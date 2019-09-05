// libraries: require compilation :/
var Msal = require('msal');
// global variables, ha ha
let accessToken = "";
let eventCache = [];
let standalone = true;
let storage = new (require('../utils/storage'));
let baseUrl = 'https://outlook.office.com/api';

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    standalone = false;

    baseUrl = Office.context.mailbox.restUrl;
  } else {
    console.error("Office context but not in Outlook :P")
  }
});

export async function run() {
  /**
   * Insert your Outlook code here
   */
  await setupTaskpane();
}

async function setupTaskpane() {
  // on value change: handleForm
  inputs = ['from', 'to', 'category', 'rate'];
  inputs.forEach(inputId => {
    let input = document.getElementById(inputId);
    input.onblur = handleForm();
  });
  rateInput = document.getElementById('rate');
  rateInput.onblur(() => {
    storage.setItem('rate', rateInput.value);
  })
  let rate = storage.getItem('rate');
  if (rate) {
    rateInput.value = rate;
  }
  // set default values
  let allCategories = await findCategories();
  let options = document.getElementsByName("category").innerHTML;
  allCategories.forEach(category => {
    options += "<option value='" + category + "'>" + category + "</option>";
  })
  document.getElementsByName("category").innerHTML = options;
  setStatus("Setup finished");
}

async function handleForm() {
  // get values
  setStatus("Retrieving Form Inputs")
  let from = document.getElementById('from').value;
  let to = document.getElementById('to').value;
  let category = document.getElementById('category').value;
  let rate = document.getElementById('rate').value;
  // load events
  let events = await loadEvents();
  // calculate overall time
  setStatus("Filtering Events")
  let totalTime = 0;
  events.forEach(event => {
    if (from <= event.End && to >= event.Start) {
      // dates intersect
      if (category == "" || event.Categories.includes(category)) {
        // category applies
        workStart = Math.max(from, event.Start);
        workEnd = Math.min(to, Event.End);
        totalTime += workEnd - workStart;
      }
    }
  });
  // present results
  setStatus("Calculating results...");
  let resultsHtml = "<table>";
  resultsHtml += "<tr><td>Time worked:</td><td>" + totalTime + "</td></tr>";
  resultsHtml += "<tr><td>That makes:</td><td>" + (totalTime * rate) + "</td></tr>";
  resultsHtml += "</table>";
  document.getElementById('results').innerHTML = resultsHtml
  // finish up
  setStatus("Done.")
}

async function findCategories() {
  setStatus("Loading categories.");
  let events = await loadEvents();
  setStatus("Finding categories.");
  let categories = []
  events.forEach(event => {
    event.Categories.forEach(category => {
      categories.push(category)
    })
  })
  return categories
}

async function loadEvents() {
  setStatus("Loading Events")
  if (eventCache.length > 0) {
    return eventCache
  }
  if (!accessToken) {
    await login();
  }
  let restUrl = baseUrl + '/v2.0/me/events?$select=Start,End,Categories';
  return new Promise(resolve => {
    fetch(restUrl, {
      headers: {
        'Authorization': 'Bearer ' + accessToken,
        'Content-Type': 'application/json'
      }
    }).then(function (response) {
      return response.json()
    }).then(function (results) {
      resolve(results)
    }).catch(function (error) {
      console.error(error)
      setStatus("Got error fetching events: " + JSON.stringify(error))
    })
  })
}

async function login() {
  return new Promise(resolve => {
    if (standalone) {
      // try login via OAuth2
      var msalConfig = require('../config.js');
      var myMSALObj = new Msal.UserAgentApplication(msalConfig);
      var requestObj = {
        scopes: ["user.read"]
      }
      var loggedIn = storage.getItem('loggedIn');
      if (!loggedIn) {
        myMSALObj.loginPopup(requestObj).then(function (loginResponse) {
          //Login Success callback
          storage.setItem('loggedIn', true);
        }).catch(function (error) {
          console.log(error);
          setStatus("Got error logging in: " + error);
        });
      }
      // Acquire access token
      myMSALObj.acquireTokenSilent(requestObj).then(function (tokenResponse) {
        accessToken = tokenResponse.accessToken;
        resolve();
      }).catch(function (error) {
        console.log(error);
        // silent failed. Try interactively:
        myMSALObj.acquireTokenPopup(requestObj).then(function (tokenResponse) {
          accessToken = tokenResponse.accessToken;
          resolve();
        }).catch(function (error) {
          console.log(error);
          setStatus("Got error logging in: " + error);
        });
      });
    } else {
      // Get login from 
      Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
        if (result.status === "succeeded") {
          accessToken = result.value;
        } else {
          setStatus("Got error logging in: " + result.status)
        }
        resolve();
      })
    }
  }
  );
}

function setStatus(status) {
  let statusSpan = document.getElementById('status');
  statusSpan.innerText = status;
}

function emptyCache() {
  eventCache = []
}
