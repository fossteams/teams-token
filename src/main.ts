import {
  BrowserWindow, app,
} from 'electron';

const uuid = require('uuid');

const USER_AGENT = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) MicrosoftTeams-Preview/1.4.00.7556 Chrome/80.0.3987.163 Electron/8.5.5 Safari/537.36';

function getLoginURL() : string {
  const loginUrl = new URL('https://login.microsoftonline.com');
  loginUrl.pathname = '/common/oauth2/authorize';
  loginUrl.searchParams.append('response_type', 'id_token');
  loginUrl.searchParams.append('client_id', '5e3ce6c0-2b1f-4285-8d4b-75ee78787346');
  loginUrl.searchParams.append('redirect_uri', 'https://teams.microsoft.com/go');
  loginUrl.searchParams.append('state', uuid.v4());
  loginUrl.searchParams.append('x-client-SKU', 'Js');
  loginUrl.searchParams.append('x-client-Ver', '1.0.9');
  loginUrl.searchParams.append('nonce', uuid.v4());

  return loginUrl.toString();
}

app.whenReady().then(() => {
  const win = new BrowserWindow({ width: 800, height: 600 });
  win.webContents.on('will-navigate', (e, url) => {
    if (url.startsWith('https://teams.microsoft.com/')) {
      e.preventDefault();
    }
  });

  win.webContents.on('did-navigate', (e, url) => {
    if (url.startsWith('https://teams.microsoft.com/go')) {
      e.preventDefault();
      const token = url.replace('https://teams.microsoft.com/go#', '');
      const searchParams = new URLSearchParams(token);
      console.log(searchParams.get('id_token'));
      win.destroy();
      app.quit();
    }
  });

  win.loadURL(getLoginURL(), {
    userAgent: USER_AGENT,
  });
});
