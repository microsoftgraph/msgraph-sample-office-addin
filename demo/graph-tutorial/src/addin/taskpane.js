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
    showError(`An error was returned from the consent dialog: ${error}`);
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
        showError(`Could not open consent prompt dialog: ${error}`);
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

// Display an error
function showError(message) {
  $('.container').empty();
  $('<div/>', {
    class: 'ms-depth-4 error-msg'
  }).append('<p/>', {
    class: 'ms-fontSize-24 ms-fontWeight-bold',
    text: 'An error occurred'
  }).append('<p/>', {
    class: 'ms-fontSize-26 ms-fontWeight-regular',
    text: message
  }).appendTo('.container');
}
// </AuthUiSnippet>

function showMainUi() {
  $('.container').empty();
  $('<p/>', {
    class: 'ms-fontSize-24 ms-fontWeight-bold',
    text: 'Authenticated!'
  }).appendTo('.container');
}

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
          showError(`Error checking auth status: ${error}`);
        } else {
          showMainUi();
        }
      }
    });
  }
});
// </OfficeReadySnippet>
