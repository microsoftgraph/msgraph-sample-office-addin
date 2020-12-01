// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

'use strict';

// <AuthUiSnippet>
// Handle to authentication pop dialog
let authDialog = undefined;

// Build a base URL from the current location
function getBaseUrl() {
  return location.protocol + '//' + location.hostname +
  (location.port ? ':' + location.port : '');
}

// Process the response back from the auth dialog
function processConsent(result) {
  const message = JSON.parse(result.message);

  authDialog.close();
  if (message.status === 'success') {
    showMainUi();
  } else {
    const error = JSON.stringify(message.result, Object.getOwnPropertyNames(message.result));
    showStatus(`An error was returned from the consent dialog: ${error}`, true);
  }
}

// Use the Office Dialog API to show the interactive
// login UI
function showConsentPopup() {
  const authDialogUrl = `${getBaseUrl()}/consent.html`;

  Office.context.ui.displayDialogAsync(authDialogUrl,
    {
      height: 60,
      width: 30,
      promptBeforeOpen: false
    },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        authDialog = result.value;
        authDialog.addEventHandler(Office.EventType.DialogMessageReceived, processConsent);
      } else {
        // Display error
        const error = JSON.stringify(error, Object.getOwnPropertyNames(error));
        showStatus(`Could not open consent prompt dialog: ${error}`, true);
      }
    });
}

// Inform the user we need to get their consent
function showConsentUi() {
  $('.container').empty();
  $('<p/>', {
    class: 'ms-fontSize-24 ms-fontWeight-bold',
    text: 'Consent for Microsoft Graph access needed'
  }).appendTo('.container');
  $('<p/>', {
    class: 'ms-fontSize-16 ms-fontWeight-regular',
    text: 'In order to access your calendar, we need to get your permission to access the Microsoft Graph.'
  }).appendTo('.container');
  $('<p/>', {
    class: 'ms-fontSize-16 ms-fontWeight-regular',
    text: 'We only need to do this once, unless you revoke your permission.'
  }).appendTo('.container');
  $('<p/>', {
    class: 'ms-fontSize-16 ms-fontWeight-regular',
    text: 'Please click or tap the button below to give permission (opens a popup window).'
  }).appendTo('.container');
  $('<button/>', {
    class: 'primary-button',
    text: 'Give permission'
  }).on('click', showConsentPopup)
  .appendTo('.container');
}

// Display a status
function showStatus(message, isError) {
  $('.status').empty();
  $('<div/>', {
    class: `status-card ms-depth-4 ${isError ? 'error-msg' : 'success-msg'}`
  }).append($('<p/>', {
    class: 'ms-fontSize-24 ms-fontWeight-bold',
    text: isError ? 'An error occurred' : 'Success'
  })).append($('<p/>', {
    class: 'ms-fontSize-16 ms-fontWeight-regular',
    text: message
  })).appendTo('.status');
}

function toggleOverlay(show) {
  $('.overlay').css('display', show ? 'block' : 'none');
}
// </AuthUiSnippet>

// <MainUiSnippet>
function showMainUi() {
  $('.container').empty();

  // Use luxon to calculate the start
  // and end of the current week. Use
  // those dates to set the initial values
  // of the date pickers
  const now = luxon.DateTime.local();
  const startOfWeek = now.startOf('week');
  const endOfWeek = now.endOf('week');

  $('<h2/>', {
    class: 'ms-fontSize-24 ms-fontWeight-semibold',
    text: 'Select a date range to import'
  }).appendTo('.container');

  // Create the import form
  $('<form/>').on('submit', getCalendar)
  .append($('<label/>', {
    class: 'ms-fontSize-16 ms-fontWeight-semibold',
    text: 'Start'
  })).append($('<input/>', {
    class: 'form-input',
    type: 'date',
    value: startOfWeek.toISODate(),
    id: 'viewStart'
  })).append($('<label/>', {
    class: 'ms-fontSize-16 ms-fontWeight-semibold',
    text: 'End'
  })).append($('<input/>', {
    class: 'form-input',
    type: 'date',
    value: endOfWeek.toISODate(),
    id: 'viewEnd'
  })).append($('<input/>', {
    class: 'primary-button',
    type: 'submit',
    id: 'importButton',
    value: 'Import'
  })).appendTo('.container');

  $('<hr/>').appendTo('.container');

  $('<h2/>', {
    class: 'ms-fontSize-24 ms-fontWeight-semibold',
    text: 'Add event to calendar'
  }).appendTo('.container');

  // Create the new event form
  $('<form/>').on('submit', createEvent)
  .append($('<label/>', {
    class: 'ms-fontSize-16 ms-fontWeight-semibold',
    text: 'Subject'
  })).append($('<input/>', {
    class: 'form-input',
    type: 'text',
    required: true,
    id: 'eventSubject'
  })).append($('<label/>', {
    class: 'ms-fontSize-16 ms-fontWeight-semibold',
    text: 'Start'
  })).append($('<input/>', {
    class: 'form-input',
    type: 'datetime-local',
    required: true,
    id: 'eventStart'
  })).append($('<label/>', {
    class: 'ms-fontSize-16 ms-fontWeight-semibold',
    text: 'End'
  })).append($('<input/>', {
    class: 'form-input',
    type: 'datetime-local',
    required: true,
    id: 'eventEnd'
  })).append($('<input/>', {
    class: 'primary-button',
    type: 'submit',
    id: 'importButton',
    value: 'Create'
  })).appendTo('.container');
}
// </MainUiSnippet>

// <WriteToSheetSnippet>
const DAY_MILLISECONDS = 86400000;
const DAY_MINUTES = 1440;
const EXCEL_DATE_OFFSET = 25569;

// Excel date cells require an OLE Automation date format
// You can use the Moment-MSDate plug-in
// (https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-ranges-advanced#work-with-dates-using-the-moment-msdate-plug-in)
// Or you can do the conversion yourself
function convertDateToOAFormat(dateTime) {
  const date = new Date(dateTime);

  // Get the time zone offset for the browser's time zone
  // since all of the dates here are handled in that time zone
  const tzOffset = date.getTimezoneOffset() / DAY_MINUTES;

  // Calculate the OLE Automation date, which is
  // the number of days since midnight, December 30, 1899
  const oaDate = date.getTime() / DAY_MILLISECONDS + EXCEL_DATE_OFFSET - tzOffset;
  return oaDate;
}

async function writeEventsToSheet(events) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const eventsTable = sheet.tables.add('A1:D1', true);

    // Create the header row
    eventsTable.getHeaderRowRange().values = [[
      'Subject',
      'Organizer',
      'Start',
      'End'
    ]];

    // Create the data rows
    const data = [];
    events.forEach((event) => {
      data.push([
        event.subject,
        event.organizer.emailAddress.name,
        convertDateToOAFormat(event.start.dateTime),
        convertDateToOAFormat(event.end.dateTime)
      ]);
    });

    eventsTable.rows.add(null, data);

    const tableRange = eventsTable.getRange();
    tableRange.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];
    tableRange.format.autofitColumns();
    tableRange.format.autofitRows();

    try {
      await context.sync();
    } catch (err) {
      console.log(`Error: ${JSON.stringify(err)}`);
      showStatus(err, true);
    }
  });
}
// </WriteToSheetSnippet>

// <GetCalendarSnippet>
async function getCalendar(evt) {
  evt.preventDefault();
  toggleOverlay(true);

  const apiToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });

  const viewStart = $('#viewStart').val();
  const viewEnd = $('#viewEnd').val();

  const requestUrl =
    `${getBaseUrl()}/graph/calendarview?viewStart=${viewStart}&viewEnd=${viewEnd}`;

  const response = await fetch(requestUrl, {
    headers: {
      authorization: `Bearer ${apiToken}`
    }
  });

  if (response.ok) {
    const events = await response.json();
    writeEventsToSheet(events);
    showStatus(`Imported ${events.length} events`, false);
  } else {
    const error = await response.json();
    showStatus(`Error getting events from calendar: ${JSON.stringify(error)}`, true);
  }

  toggleOverlay(false);
}
// </GetCalendarSnippet>

// <CreateEventSnippet>
async function createEvent(evt) {
  evt.preventDefault();
  toggleOverlay(true);

  const apiToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });

  const payload = {
    eventSubject: $('#eventSubject').val(),
    eventStart: $('#eventStart').val(),
    eventEnd: $('#eventEnd').val()
  };

  const requestUrl = `${getBaseUrl()}/graph/newevent`;

  const response = await fetch(requestUrl, {
    method: 'POST',
    headers: {
      authorization: `Bearer ${apiToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(payload)
  });

  if (response.ok) {
    showStatus('Event created', false);
  } else {
    const error = await response.json();
    showStatus(`Error creating event: ${JSON.stringify(error)}`, true);
  }

  toggleOverlay(false);
}
// </CreateEventSnippet>

// <OfficeReadySnippet>
Office.onReady(info => {
  // Only run if we're inside Excel
  if (info.host === Office.HostType.Excel) {
    $(async function() {
      let apiToken = '';
      try {
        apiToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });
        console.log(`API Token: ${apiToken}`);
      } catch (error) {
        console.log(`getAccessToken error: ${error}`);
        // Fall back to interactive login
        showConsentUi();
      }

      // Call auth status API to see if we need to get consent
      const authStatusResponse = await fetch(`${getBaseUrl()}/auth/status`, {
        headers: {
          authorization: `Bearer ${apiToken}`
        }
      });

      const authStatus = await authStatusResponse.json();
      if (authStatus.status === 'consent_required') {
        showConsentUi();
      } else {
        // report error
        if (authStatus.status === 'error') {
          const error = JSON.stringify(authStatus.error,
            Object.getOwnPropertyNames(authstatus.error));
          showStatus(`Error checking auth status: ${error}`, true);
        } else {
          showMainUi();
        }
      }
    });
  }
});
// </OfficeReadySnippet>
