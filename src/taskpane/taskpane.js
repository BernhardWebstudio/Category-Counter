console.log("0");
const Moment = require('moment');
console.log("1");

const MomentRange = require('moment-range');
console.log("2");

const moment = MomentRange.extendMoment(Moment);
console.log("3");

// global variables, ha ha
let eventCache = [];
const storage = (require('../utils/storage')).default;
console.log("5");
const restHandler = (require('../utils/restHandler')).default;
console.log("6");
let baseUrl = 'https://graph.microsoft.com/';
console.log("7");

// Office ready
Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    baseUrl = Office.context.mailbox.restUrl;
    restHandler.setStandalone(false);
    storage.setStandalone(false);
  } else {
    console.error('Office context available but not in Outlook :P')
  }
  run();
});
// Window (standalone) ready
if (window && !Office) {
  window.addEventListener('load', () => {
    run();
  });
}

export async function run() {
  await setupTaskpane();
}

async function setupTaskpane() {
  // set default values
  try {
    // await categories so we have them in cache afterwards and reduce number of 
    // unnecessary API calls
    let allCategories = await findCategories();
    let options = document.getElementById('category').innerHTML;
    if (!options) {
      options = '<option value=\'\'>Any</option>';
    }
    allCategories.forEach(category => {
      options +=
        '<option value=\'' + category + '\'>' + category + '</option>';
    })
    console.log('options', options);
    document.getElementById('category').innerHTML = options;
    setStatus('Setting up categories finished');
  } catch (error) {
    console.error(error);
    setStatus('Failed setting up categories: ' + error);
  }
  // on value change: handleForm
  let inputs = ['daterange', 'category', 'rate'];
  inputs.forEach(inputId => {
    let input = document.getElementById(inputId);
    input.onblur = handleForm();
  });
  let rateInput = document.getElementById('rate');
  rateInput.onblur = () => {
    storage.setItem('rate', rateInput.value);
  };
  let rate = storage.getItem('rate');
  if (rate) {
    rateInput.value = rate;
  }
  await setupDatepicker();
  let submitBtn = document.getElementById('submitBtn');
  submitBtn.addEventListener('click', function (event) {
    event.preventDefault();
    handleForm();
  })

  setStatus('Setup finished');
}

async function setupDatepicker() {
  setStatus('Setting up DatePicker');
  let genericDatePickerConfig = {
    showDropdowns: true,
    // timePicker: true,
    // timePickerIncrement: 60,
    linkedCalendars: false,
    locale: {
      "format": "YYYY/MM/DD", // DD.MM.YYYY unfortunately leads to problems
      "separator": " - ",
      "applyLabel": "Anwenden",
      "cancelLabel": "Abbrechen",
      "fromLabel": "Von",
      "toLabel": "Bis",
      "customRangeLabel": "Custom",
      "weekLabel": "W",
      "daysOfWeek": [
        "So",
        "Mo",
        "Di",
        "Mi",
        "Do",
        "Fr",
        "Sa"
      ],
      "monthNames": [
        "Januar",
        "Februar",
        "März",
        "April",
        "Mai",
        "Juni",
        "July",
        "August",
        "September",
        "Oktober",
        "November",
        "Dezember"
      ],
      "firstDay": 1
    }
  };
  loadEvents().then(events => {
    let minDate = parseOutlookDate(events[0].start);
    let maxDate = parseOutlookDate(events[0].end);
    // find first & last events
    events.forEach(event => {
      let eventStart = parseOutlookDate(event.start);
      let eventEnd = parseOutlookDate(event.end);
      if (minDate > eventStart) {
        minDate = eventStart;
      }
      if (maxDate < eventEnd) {
        maxDate = eventEnd;
      }
    });
    genericDatePickerConfig.minDate = minDate;
    genericDatePickerConfig.maxDate = maxDate;
    // initialize date picker
    $('input[name="daterange"]')
      .daterangepicker(genericDatePickerConfig, handleForm);
    setStatus('Setting up DatePicker finished');
  }).catch(error => {
    console.error(error);
    setStatus('Failed setting up daterangepicker: ' + error);
    // setup anyways
    $('input[name="daterange"]')
      .daterangepicker(genericDatePickerConfig, handleForm);
  });
}

function parseOutlookDate(date) {
  if (date.timeZone === 'UTC') {
    // return moment.utc(date.dateTime)
    return moment(date.dateTime);
  } else {
    console.error('Unsupported TimeZone: ' + date.timeZone);
    return moment(date.dateTime);
  }
}

async function handleForm() {
  // get values
  setStatus('Retrieving Form Inputs')
  var drp = $('input[name="daterange"]').data('daterangepicker');
  let from = drp.startDate;
  let to = drp.endDate;
  let targetRange = moment().range(from, to);
  let category = document.getElementById('category').value.trim();
  let rate = document.getElementById('rate').value;
  // load events
  let events = await loadEvents();
  // calculate overall time
  setStatus('Filtering Events');
  let totalTime = 0;
  events.forEach(event => {
    let eventStart = parseOutlookDate(event.start);
    let eventEnd = parseOutlookDate(event.end);
    let eventRange = moment().range(eventStart, eventEnd);
    if (targetRange.overlaps(eventRange)) {
      // dates intersect
      if (category === '' || event.categories.includes(category)) {
        // category applies
        let intersection = targetRange.intersect(eventRange);
        totalTime += intersection + 0;  // milliseconds
      } else {
        console.log(
          'Category \'' + category + '\' not included.', event.categories)
      }
    } else {
      console.log('Do not apply:')
      console.log([targetRange, eventRange]);
    }
  });
  // present results
  setStatus('Calculating results...');
  let resultsHtml = '<table>';
  resultsHtml +=
    '<tr><td>Time worked:</td><td>' + (totalTime / 3.6e+6) + ' h</td></tr>';
  resultsHtml += '<tr><td>That makes:</td><td>' + (totalTime / 3.6e+6 * rate) +
    '</td></tr>';
  resultsHtml += '</table>';
  document.getElementById('results').innerHTML = resultsHtml;
  // finish up
  setStatus('Done.');
}

function findCategories() {
  return new Promise(resolve => {
    setStatus('Loading categories.')
    loadEvents()
      .then(events => {
        setStatus('Finding categories.')
        let categories = []
        console.log(events);
        events.forEach(
          event => {
            event.categories.forEach(
              category => { categories.push(category) })
          })
        resolve(categories);
      })
      .catch(error => {
        console.error(error);
        setStatus('Got error finding categories: ' + JSON.stringify(error));
        throw error;  // throw futher up
      });
  });
}

async function loadEvents() {
  return new Promise(resolve => {
    setStatus('Loading Events')
    if (eventCache.length > 0) {
      resolve(eventCache);
      return;
    }
    let restUrl = baseUrl + '/v1.0/me/events?$select=Start,End,Categories';
    try {
      let p = restHandler.makeGetRequest(restUrl);
      console.log(p);
      p.then((results) => {
        resolve(results.value);
      }).catch(error => {
        console.error(error);
        setStatus('Got error fetching events: ' + JSON.stringify(error));
        throw error;
      })
    } catch (error) {
      console.error(error);
      setStatus('Got error fetching events: ' + JSON.stringify(error));
      throw error;  // throw futher up
    }
  });
}

function setStatus(status) {
  let statusSpan = document.getElementById('status');
  statusSpan.innerText = status;
  console.log(status);
}

function emptyCache() {
  eventCache = []
}
