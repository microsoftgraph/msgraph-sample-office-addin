// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import Router from 'express-promise-router';
import { zonedTimeToUtc } from 'date-fns-tz';
import { findIana } from 'windows-iana';
import * as graph from '@microsoft/microsoft-graph-client';
import { Event, MailboxSettings } from 'microsoft-graph';
import 'isomorphic-fetch';
import { getTokenOnBehalfOf } from './auth';

// <GetClientSnippet>
async function getAuthenticatedClient(authHeader: string): Promise<graph.Client> {
  const accessToken = await getTokenOnBehalfOf(authHeader);

  return graph.Client.init({
    authProvider: (done) => {
      // Call the callback with the
      // access token
      done(null, accessToken || '');
    }
  });
}
// </GetClientSnippet>

// <GetTimeZonesSnippet>
interface TimeZones {
  // The string returned by Microsoft Graph
  // Could be Windows name or IANA identifier.
  graph: string;
  // The IANA identifier
  iana: string;
}

async function getTimeZones(client: graph.Client): Promise<TimeZones> {
  // Get mailbox settings to determine user's
  // time zone
  const settings: MailboxSettings = await client
    .api('/me/mailboxsettings')
    .get();

  // Time zone from Graph can be in IANA format or a
  // Windows time zone name. If Windows, convert to IANA
  const ianaTzs = findIana(settings.timeZone || '');
  const ianaTz = ianaTzs ? ianaTzs[0] : null;

  const returnValue: TimeZones = {
    graph: settings.timeZone || '',
    iana: ianaTz ?? settings.timeZone ?? ''
  };

  return returnValue;
}
// </GetTimeZonesSnippet>

const graphRouter = Router();

// <GetCalendarViewSnippet>
graphRouter.get('/calendarview',
  async function(req, res) {
    const authHeader = req.headers['authorization'];

    if (authHeader) {
      try {
        const client = await getAuthenticatedClient(authHeader);

        const viewStart = req.query['viewStart']?.toString();
        const viewEnd = req.query['viewEnd']?.toString();

        const timeZones = await getTimeZones(client);

        // Convert the start and end times into UTC from the user's time zone
        const utcViewStart = zonedTimeToUtc(viewStart || '', timeZones.iana);
        const utcViewEnd = zonedTimeToUtc(viewEnd || '', timeZones.iana);

        // GET events in the specified window of time
        const eventPage: graph.PageCollection = await client
          .api('/me/calendarview')
          // Header causes start and end times to be converted into
          // the requested time zone
          .header('Prefer', `outlook.timezone="${timeZones.graph}"`)
          // Specify the start and end of the calendar view
          .query({
            startDateTime: utcViewStart.toISOString(),
            endDateTime: utcViewEnd.toISOString()
          })
          // Only request the fields used by the app
          .select('subject,start,end,organizer')
          // Sort the results by the start time
          .orderby('start/dateTime')
          // Limit to at most 25 results in a single request
          .top(25)
          .get();

        const events: Event[] = [];

        // Set up a PageIterator to process the events in the result
        // and request subsequent "pages" if there are more than 25
        // on the server
        const callback: graph.PageIteratorCallback = (event) => {
          // Add each event into the array
          events.push(event);
          return true;
        };

        const iterator = new graph.PageIterator(client, eventPage, callback, {
          headers: {
            'Prefer': `outlook.timezone="${timeZones.graph}"`
          }
        });
        await iterator.iterate();

        // Return the array of events
        res.status(200).json(events);
      } catch (error) {
        console.log(error);
        res.status(500).json(error);
      }
    } else {
      // No auth header
      res.status(401).end();
    }
  }
);
// </GetCalendarViewSnippet>

// <CreateEventSnippet>
graphRouter.post('/newevent',
  async function(req, res) {
    const authHeader = req.headers['authorization'];

    if (authHeader) {
      try {
        const client = await getAuthenticatedClient(authHeader);

        const timeZones = await getTimeZones(client);

        // Create a new Graph Event object
        const newEvent: Event = {
          subject: req.body['eventSubject'],
          start: {
            dateTime: req.body['eventStart'],
            timeZone: timeZones.graph
          },
          end: {
            dateTime: req.body['eventEnd'],
            timeZone: timeZones.graph
          }
        };

        // POST /me/events
        await client.api('/me/events')
          .post(newEvent);

        // Send a 201 Created
        res.status(201).end();
      } catch (error) {
        console.log(error);
        res.status(500).json(error);
      }
    } else {
      // No auth header
      res.status(401).end();
    }
  }
);
// </CreateEventSnippet>

export default graphRouter;
