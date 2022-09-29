// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// <ServerSnippet>
import express, { RequestHandler } from 'express';
import https from 'https';
import fs from 'fs';
import dotenv from 'dotenv';
import path from 'path';

// Load .env file
dotenv.config();

import authRouter from './api/auth';
import graphRouter from './api/graph';

const app = express();
const PORT = 3000;

// Support JSON payloads
app.use(express.json() as RequestHandler);
app.use(express.static(path.join(__dirname, 'addin')));
app.use(express.static(path.join(__dirname, 'dist/addin')));

app.use('/auth', authRouter);
app.use('/graph', graphRouter);

const serverOptions = {
  key: fs.readFileSync(process.env.TLS_KEY_PATH || ''),
  cert: fs.readFileSync(process.env.TLS_CERT_PATH || ''),
};

https.createServer(serverOptions, app).listen(PORT, () => {
  console.log(`⚡️[server]: Server is running at https://localhost:${PORT}`);
});
// </ServerSnippet>
