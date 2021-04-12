import {
  BrowserWindow, app, session,
} from 'electron';

import { homedir } from 'os';
import { writeFileSync, existsSync, mkdirSync } from 'fs';
import { v4 as uuidv4 } from 'uuid';
import axios, { AxiosResponse } from 'axios';
import * as jwt from 'jsonwebtoken';

const configPath = '.config/fossteams';
const DEBUG = false;
const MICROSOFT_TENANT_ID = 'f8cdef31-a31e-4b4a-93e4-5f571e91255a';
const TEAMS_APP_ID = '5e3ce6c0-2b1f-4285-8d4b-75ee78787346';
const SKYPE_RESOURCE = 'https://api.spaces.skype.com';
const CHAT_SVC_AGG_RESOURCE = 'https://chatsvcagg.teams.microsoft.com';
const USER_AGENT = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) MicrosoftTeams-Preview/1.4.00.7556 Chrome/80.0.3987.163 Electron/8.5.5 Safari/537.36';

type TeamsSkype = 'teams' | 'skype' | 'chatsvcagg';

let win : BrowserWindow | null = null;
let tokenResponseCount = 0;
let currentTenant : (string | null) = null;

function getLoginURL(type: TeamsSkype, tenantId: string) : string {
  const loginUrl = new URL('https://login.microsoftonline.com');
  loginUrl.pathname = `/${tenantId}/oauth2/authorize`;

  const state = uuidv4();
  switch (type) {
    case 'teams':
      loginUrl.searchParams.append('response_type', 'id_token');
      loginUrl.searchParams.append('state', state);
      break;

    case 'skype':
      loginUrl.searchParams.append('response_type', 'token');
      loginUrl.searchParams.append('state', `${state}|${SKYPE_RESOURCE}`);
      loginUrl.searchParams.append('resource', SKYPE_RESOURCE);
      break;

    case 'chatsvcagg':
      loginUrl.searchParams.append('response_type', 'token');
      loginUrl.searchParams.append('state', `${state}|${CHAT_SVC_AGG_RESOURCE}`);
      loginUrl.searchParams.append('resource', CHAT_SVC_AGG_RESOURCE);
      break;

    default:
      break;
  }
  loginUrl.searchParams.append('client_id', TEAMS_APP_ID);
  loginUrl.searchParams.append('client-request-id', uuidv4());
  loginUrl.searchParams.append('redirect_uri', 'https://teams.microsoft.com/go');
  loginUrl.searchParams.append('x-client-SKU', 'Js');
  loginUrl.searchParams.append('x-client-Ver', '1.0.9');
  loginUrl.searchParams.append('prompt', 'none');
  loginUrl.searchParams.append('nonce', uuidv4());

  return loginUrl.toString();
}

function authorize(type: TeamsSkype, tenantId: string) {
  console.log(`Authorizing ${type} with tenantId=${tenantId}`);
  win.loadURL(getLoginURL(type, tenantId), {
    userAgent: USER_AGENT,
  });
}

function saveTeamsToken(token: string, type: TeamsSkype) {
  if (!existsSync(`${homedir}/${configPath}/`)) {
    mkdirSync(`${homedir}/${configPath}`, { recursive: true });
  }
  writeFileSync(`${homedir}/${configPath}/token-${type}.jwt`, token);
}

type Tenant = {
  tenantId: string;
  tenantName: string;
  userId: string;
  isInvitationRedeemed: boolean;
  userType: string;
  tenantType: string;
}

function getTenants(token: string) : Promise<AxiosResponse<Tenant[]>> {
  const req = axios.get('https://teams.microsoft.com/api/mt/emea/beta/users/tenants',
    {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });
  return req;
}

app.whenReady().then(() => {
  // Check process arguments
  const p = new Promise<void>((resolve, _) => {
    if (process.argv.length === 3) {
      // Check if 3rd argument is "logout"
      switch (process.argv[2]) {
        case 'logout':
          session.defaultSession.clearStorageData().then(() => {
            resolve();
          });
          break;

        case 'get-url':
          console.log(getLoginURL('teams', 'common'));
          app.exit();
          break;

        default:
          break;
      }
    }
    resolve();
  });

  p.then(() => {
    win = new BrowserWindow({ width: 800, height: 600 });
    if (DEBUG) {
      win.webContents.openDevTools();
    }
    win.webContents.on('will-navigate', (e, url) => {
      if (url.startsWith('https://teams.microsoft.com/')) {
        e.preventDefault();
      }
    });

    win.webContents.on('did-navigate', async (e, url) => {
      if (url.startsWith('https://teams.microsoft.com/go')) {
        const token = url.replace('https://teams.microsoft.com/go#', '');
        const searchParams = new URLSearchParams(token);
        let teamsToken = searchParams.get('id_token');

        if (teamsToken === null) {
          teamsToken = searchParams.get('access_token');
        }

        const decoded = jwt.decode(teamsToken);
        if (decoded === null) {
          console.warn(`Inavlid JWT provided: ${searchParams}`);
          return;
        }

        if (typeof (decoded) === 'string') {
          console.error('Invalid decoded JWT: is a string');
          return;
        }

        if (decoded.tid === MICROSOFT_TENANT_ID && decoded.aud === SKYPE_RESOURCE) {
          // Skip, we need a tenant selection
          console.log(`Tenant ID is MICROSOFT: aud=${decoded.aud}`);
          // Get Tenant list
          try {
            const tenants = await (await getTenants(teamsToken)).data;
            // Pick the first tenant and authorize Skype
            currentTenant = tenants[0].tenantId;
            authorize('skype', currentTenant);
          } catch (err) {
            console.error(`Unable to get tenants: ${err}`);
          }
          // win.webContents.stop();
          // win.webContents.loadURL('https://teams.microsoft.com/go');
          return;
        }

        tokenResponseCount += 1;

        if (tokenResponseCount > 5) {
          console.error('Redirecting too many times, stopping');
          e.preventDefault();
          win.webContents.stop();
          return;
        }

        console.log(`Audience: ${decoded.aud}`);
        console.log('Decoded', decoded);


        win.webContents.stop();

        if (decoded.aud === TEAMS_APP_ID) {
        // Teams Token
          console.log('Got a Teams token');
          saveTeamsToken(teamsToken, 'teams');
          win.destroy();
          app.quit();
        } else if (decoded.aud === SKYPE_RESOURCE) {
          console.log('Got a Skype token');
          saveTeamsToken(teamsToken, 'skype');
          authorize('chatsvcagg', currentTenant);
        } else if (decoded.aud === CHAT_SVC_AGG_RESOURCE) {
          console.log('Got a ChatSvcAgg token');
          saveTeamsToken(teamsToken, 'chatsvcagg');
          authorize('teams', currentTenant);
        } else {
          console.error(`Invalid audience ${decoded.aud} found.`);
        }
      }
    });

    authorize('skype', 'common');
  });
});
