// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/* global authConfig, localStorage, Office */

// <ConsentJsSnippet>
'use strict';

// @ts-ignore
var msal = msal || {
  PublicClientApplication: () => {
    throw new Error('MSAL not loaded');
  },
};

const msalClient = new msal.PublicClientApplication({
  auth: {
    // authConfig is defined in config.js
    // @ts-ignore
    clientId: authConfig.clientId,
    navigateToLoginRequestUrl: false,
  },
  cache: {
    cacheLocation: 'localStorage',
    storeAuthStateInCookie: true,
  },
});

const msalRequest = {
  scopes: ['https://graph.microsoft.com/.default'],
};

// Function that handles the redirect back to this page
// once the user has signed in and granted consent
/**
 * @param {{ account: { homeId: string; }; accessToken: any; } | null} response
 */
function handleResponse(response) {
  localStorage.removeItem('msalCallbackExpected');
  if (response !== null) {
    localStorage.setItem('msalAccountId', response.account.homeId);
    Office.context.ui.messageParent(
      JSON.stringify({ status: 'success', result: response.accessToken }),
    );
  }
}

Office.onReady(() => {
  if (Office.context.ui.messageParent) {
    // Let MSAL process a redirect response if that's what
    // caused this page to load.
    msalClient
      .handleRedirectPromise()
      .then(handleResponse)
      .catch((/** @type {any} */ error) => {
        console.log(error);
        Office.context.ui.messageParent(
          JSON.stringify({ status: 'failure', result: error }),
        );
      });

    // If we're not expecting a callback (because this is
    // the first time the page has loaded), then start the
    // login process
    if (!localStorage.getItem('msalCallbackExpected')) {
      // Set the msalCallbackExpected property so we don't
      // make repeated token requests
      localStorage.setItem('msalCallbackExpected', 'yes');

      // If the user has signed into this machine before
      // do a token request, otherwise do a login
      if (localStorage.getItem('msalAccountId')) {
        msalClient.acquireTokenRedirect(msalRequest);
      } else {
        msalClient.loginRedirect(msalRequest);
      }
    }
  }
});
// </ConsentJsSnippet>
