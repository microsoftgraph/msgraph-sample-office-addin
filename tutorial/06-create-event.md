<!-- markdownlint-disable MD002 MD041 -->

In this section you will add the ability to create events on the user's calendar.

## Implement the API

1. Open **./src/api/graph.ts** and add the following code to implement a new event API (`POST /graph/newevent`).

    :::code language="typescript" source="../demo/graph-tutorial/src/api/graph.ts" id="CreateEventSnippet":::

1. Open **./src/addin/taskpane.js** and add the following function to call the new event API.

    :::code language="javascript" source="../demo/graph-tutorial/src/addin/taskpane.js" id="CreateEventSnippet":::

1. Save all of your changes, restart the server, and refresh the task pane in Excel (close any open task panes and re-open).

    ![A screenshot of the create event form](images/create-event-ui.png)

1. Fill in the form and choose **Create**. Verify that the event is added to the user's calendar.
