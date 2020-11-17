<!-- markdownlint-disable MD002 MD041 -->

In this exercise you will enable [Office Add-in single sign-on (SSO)](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins) in the add-in, and extend the web API to support [on-behalf-of flow](https://docs.microsoft.com/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow). This is required to obtain the necessary OAuth access token to call the Microsoft Graph.

## Overview

Office Add-in SSO provides an access token, but that token is only enables the add-in to call it's own web API. It does not enable direct access to the Microsoft Graph. The process works as follows.

1. The add-in gets a token by calling [getAccessToken](https://docs.microsoft.com/javascript/api/office-runtime/officeruntime.auth?view=common-js#getaccesstoken-options-). This token's audience (the `aud` claim) is the application ID of the add-in's app registration.
1. The add-in sends this token in the `Authorization` header when it makes a call to the web API.
1. The web API validates the token, then uses the on-behalf-of flow to exchange this token for a Microsoft Graph token. This new token's audience is `https://graph.microsoft.com`.
1. The web API uses the new token to make calls to the Microsoft Graph, and returns the results back to the add-in.

## Configure the solution

1. Open **./.env** and update the `AZURE_APP_ID`, `AZURE_CLIENT_SECRET`, and `AZURE_TENANT_ID` with the application ID, client secret, and tenant ID from your app registration.

    > [!IMPORTANT]
    > If you're using source control such as git, now would be a good time to exclude the **.env** file from source control to avoid inadvertently leaking your app ID and client secret.

1. Open **./manifest/manifest.xml** and replace all instances of `YOUR_APP_ID_HERE` with the application ID from your app registration.

1. Create a new file in the **./src/addin** directory named **config.js** and add the following code, replacing `YOUR_APP_ID_HERE` with the application ID from your app registration.

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/config.example.js":::

## Implement sign-in

1. Open **./src/api/auth.ts** and add the following `import` statements at the top of the file.

    ```typescript
    import jwt, { SigningKeyCallback, JwtHeader } from 'jsonwebtoken';
    import jwksClient from 'jwks-rsa';
    import * as msal from '@azure/msal-node';
    ```

1. Add the following code after the `import` statements.

    :::code language="typescript" source="../demo/graph-tutorial/src/api/auth.ts" id="TokenExchangeSnippet":::

    This code [initializes an MSAL confidential client](https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-node/docs/initialize-confidential-client-application.md), and exports a function to get a Graph token from the token sent by the add-in.

1. Add the following code before the `export default authRouter;` line.

    :::code language="typescript" source="../demo/graph-tutorial/src/api/auth.ts" id="GetAuthStatusSnippet":::

    This code implements an API (`GET /auth/status`) that checks if the add-in token can be silently exchanged for a Graph token. The add-in will use this API to determine if it needs to present an interactive login to the user.

1. Open **./src/addin/taskpane.js** and add the following code to the file.

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="AuthUiSnippet":::

    This code adds functions to update the UI, and to use the [Office Dialog API](https://docs.microsoft.com/office/dev/add-ins/develop/dialog-api-in-office-add-ins) to initiate an interactive authentication flow.

1. Add the following function to implement a temporary main UI.

    ```javascript
    function showMainUi() {
      $('.container').empty();
      $('<p/>', {
        class: 'ms-fontSize-24 ms-fontWeight-bold',
        text: 'Authenticated!'
      }).appendTo('.container');
    }
    ```

1. Replace the existing `Office.onReady` call with the following.

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="OfficeReadySnippet":::

    Consider what this code does.

    - When the task pane first loads, it calls `getAccessToken` to get a token scoped for the add-in's web API.
    - It uses that token to call the `/auth/status` API to check if the user has given consent to the Microsoft Graph scopes yet.
        - If the user has not consented, it uses a pop-up window to get the user's consent through an interactive login.
        - If the user has consented, it loads the main UI.

### Getting user consent

Even though the add-in is using SSO, the user still has to consent to the add-in accessing their data via Microsoft Graph. Getting consent is a one-time process. Once the user has granted consent, the SSO token can be exchanged for a Graph token without any user interaction. In this section you'll implement the consent experience in the add-in using [msal-browser](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-browser).

1. Create a new file in the **./src/addin** directory named **consent.js** and add the following code.

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/consent.js" id="ConsentJsSnippet":::

    This code does login for the user, requesting the set of Microsoft Graph permissions that are configured on the app registration.

1. Create a new file in the **./src/addin** directory named **consent.html** and add the following code.

    :::code language="html" source="../demo/graph-tutorial/src/addin/consent.html" id="ConsentHtmlSnippet":::

    This code implements a basic HTML page to load the **consent.js** file. This page will be loaded in a pop-up dialog.

1. Save all of your changes and restart the server.

1. Re-upload your **manifest.xml** file using the same steps in [Side-load the add-in in Excel](02-create-app.md#side-load-the-add-in-in-excel).

1. Select the **Import Calendar** button on the **Home** tab to open the task pane.

1. Select the **Give permission** button in the task pane to launch the consent dialog in a pop-up window. Sign in and grant consent.

1. The task pane updates with an "Authenticated!" message. You can check the tokens as follows.

    - In your brower's developer tools, the API token is shown in the Console.
    - In your CLI where you are running the Node.js server, the Graph token is printed.

    You can compare these token at [https://jwt.ms](https://jwt.ms). Notice that the API token's audience (`aud`) is set to the application ID of your app registration, and the scope (`scp`) is `access_as_user`.
