// global variables, ha ha
let accessToken = "";
let eventCache = [];

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  /**
   * Insert your Outlook code here
   */
  setupTaskpane();
}

function setupTaskpane() {
  // on value change: handleForm
  inputs = ['from', 'to', 'category', 'rate']
  inputs.forEach(inputId => {
    let input = document.getElementById(inputId)
    input.onblur = handleForm()
  })
  rateInput = document.getElementById('rate')
  rateInput.onblur(() => {
    Office.context.document.settings.set('rate', rateInput.value)
  })
  let rate = Office.context.document.settings.get('rate')
  if (rate) {
    rateInput.value = rate;
  }
  // set default values
  let allCategories = await findCategories()
  let options = document.getElementsByName("category").innerHTML
  allCategories.forEach(category => {
    options += "<option value='" + category + "'>" + category + "</option>"
  })
  document.getElementsByName("category").innerHTML = options
  setStatus("Setup finished")
}

async function handleForm() {
  // get values
  setStatus("Retrieving Form Inputs")
  let from = document.getElementById('from').value;
  let to = document.getElementById('to').value;
  let category = document.getElementById('category').value;
  let rate = document.getElementById('rate').value;
  // load events
  let events = await loadEvents()
  // calculate overall time
  setStatus("Filtering Events")
  let totalTime = 0;
  events.forEach(event => {
    if (from <= event.End && to >= event.Start) {
      // dates intersect
      if (category == "" || event.Categories.includes(category)) {
        // category applies
        workStart = Math.max(from, event.Start)
        workEnd = Math.min(to, Event.End)
        totalTime += workEnd - workStart
      }
    }
  });
  // present results
  setStatus("Calculating results...")
  let resultsHtml = "<table>"
  resultsHtml += "<tr><td>Time worked:</td><td>" + totalTime + "</td></tr>"
  resultsHtml += "<tr><td>That makes:</td><td>" + (totalTime * rate) + "</td></tr>"
  resultsHtml += "</table>"
  document.getElementById('results').innerHTML = resultsHtml
  // finish up
  setStatus("Done.")
}

async function findCategories() {
  setStatus("Loading categories.")
  let events = await loadEvents()
  setStatus("Finding categories.")
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
  let restUrl = Office.context.mailbox.restUrl + '/v2.0/me/events?$select=Start,End,Categories';
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
  return new Promise(resolve =>
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
      if (result.status === "succeeded") {
        accessToken = result.value;
      } else {
        setStatus("Got error logging in: " + result.status)
      }
      resolve();
    })
  );
}

function setStatus(status) {
  let statusSpan = document.getElementById('status');
  statusSpan.innerText = status;
}

function emptyCache() {
  eventCache = []
}
