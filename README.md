# teams-token

A super simple Electron app that will save your Microsoft Teams tokens in `~/.config/fossteams/`.

## Instructions

```bash
yarn install
yarn start
```

## Note

Log-in with your Microsoft Teams account (your corprorate / school account) and let this app do the rest.
If you don't trust putting your credentials in a random GitHub app, audit the code for yourself: we're basically
using Electron to intercept Oauth redirections and storing the tokens on your computer.

The Microsoft Teams authentication flow is described [here](https://github.com/fossteams/teams-api/blob/master/notes/login-flow.md). We're
getting the token that is passed to `https://teams.microsoft.com/go`.