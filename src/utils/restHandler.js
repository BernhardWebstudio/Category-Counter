// libraries: require compilation :/
var Msal = require('msal');
import msalConfig from '../config.js';

class RESTHandler {
  constructor() {
    this.standalone = true;
  }

  setStandalone(standalone) {
    this.standalone = standalone;
  }

  async makeGetRequest(url) {
    return new Promise(resolve => {
      this.getAuthToken().then(() => {
        fetch(url, {
          headers: {
            'Authorization': 'Bearer ' + this.accessToken,
            'Content-Type': 'application/json'
          }
        }).then(function (response) {
          return response.json()
        }).then(function (results) {
          resolve(results)
        }).catch(function (error) {
          console.error(error)
          throw error;
        })
      }).catch(function (error) {
        console.error(error)
        throw error;
      });
    });
  }

  getAuthToken() {
    if (this.accessToken) {
      // we already got it
      return new Promise(resolve => { resolve() })
    }
    if (this.standalone) {
      // Get login from/for standalone web version
      if (this.myMSALObj.getAccount()) {
        // account signed in, fetch auth token
        return new Promise(resolve => {
          this.acquireToken((token) => {
            this.accessToken = token;
            resolve();
          });
        });
      } else {
        // sign in required
        return new Promise(resolve => {
          this.signIn((token) => {
            this.accessToken = token;
            resolve();
          })
        })
      }
    } else {// Get login from/for Outlook
      let self = this;
      return new Promise(resolve => {
        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
          if (result.status === "succeeded") {
            self.accessToken = result.value;
          } else {
            console.error("Got error logging in: " + result.status)
          }
          resolve();
        })
      })
    }
  }

  signIn(tokenCallback) {
    this.myMSALObj.loginPopup({
      scopes: msalConfig.graphScopes
    }).then(function (loginResponse) {
      //Login Success
      console.log(loginResponse);
      this.acquireToken(tokenCallback);
    }).catch(function (error) {
      console.error(error);
    });
  }

  acquireToken(callback) {
    let self = this;
    // Always start with acquireTokenSilent to obtain a token in the signed in user from cache
    this.myMSALObj.acquireTokenSilent({
      scopes: msalConfig.graphScopes
    }).then(function (tokenResponse) {
      callback(tokenResponse.accessToken);
    }).catch(function (error) {
      console.error(error);
      // Upon acquireTokenSilent failure (due to consent or interaction or login required ONLY)
      // Call acquireTokenPopup(popup window)
      if (self.requiresInteraction(error.errorCode)) {
        self.myMSALObj.acquireTokenPopup({
          scopes: msalConfig.graphScopes
        }).then(function (tokenResponse) {
          callback(tokenResponse.accessToken);
        }).catch(function (error) {
          console.error(error);
        });
      }
    });
  }

  requiresInteraction(errorCode) {
    if (!errorCode || !errorCode.length) {
      return false;
    }
    return errorCode === "consent_required" ||
      errorCode === "interaction_required" ||
      errorCode === "login_required";
  }
}

export default new RESTHandler();
