var msalConfig = {
  "auth": {
    "clientId": "25773fae-dc08-46dc-abf3-eb73030ea422",
    "authority": "https://login.microsoftonline.com/common",
    // "redirectUri": "msal25773fae-dc08-46dc-abf3-eb73030ea422://auth"
    // "redirectUri": "https://app.bernhard-webstudio.ch/category-counter/taskpane.html"
  },
  "cache": {
    "cacheLocation": "localStorage",
    "storeAuthStateInCookie": true
  },
  "graphScopes":
    ['user.read', 'Calendars.Read', 'Calendars.Read.Shared']
};
export default msalConfig;
