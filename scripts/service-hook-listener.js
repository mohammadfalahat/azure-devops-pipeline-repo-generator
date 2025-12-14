#!/usr/bin/env node
/**
 * Lightweight Azure DevOps Service Hook listener inspired by the official samples:
 * https://learn.microsoft.com/azure/devops/extend/develop/add-service-hook
 *
 * - Runs an HTTP server (default port 3000) that accepts POST payloads from Azure DevOps Service Hooks.
 * - Logs basic details (event type, notification ID, repository/collection context) so you can validate
 *   the webhook contract locally before deploying your extension.
 * - Includes a `--self-test` mode that exercises the parser without opening a network port.
 */
const http = require('http');

const port = parseInt(process.env.PORT || '3000', 10);
const verbose = process.env.LOG_PAYLOADS !== 'false';

const summarizePayload = (payload = {}) => {
  const eventType = payload.eventType || payload?.resource?.eventType || 'unknown';
  const notificationId =
    payload.notificationId ?? payload.id ?? payload?.resourceContainers?.notificationId ?? 'n/a';
  const collection = payload?.resourceContainers?.collection?.baseUrl || 'unknown collection';
  const project = payload?.resource?.project?.name || payload?.resource?.project?.id || 'unknown project';
  const repo = payload?.resource?.repository?.name || payload?.resource?.repository?.id || 'unknown repo';
  return `eventType=${eventType} notificationId=${notificationId} project=${project} repo=${repo} collection=${collection}`;
};

const runSelfTest = () => {
  const sample = {
    id: 42,
    eventType: 'git.push',
    resourceContainers: {
      collection: { baseUrl: 'https://dev.azure.local/DefaultCollection' }
    },
    resource: {
      project: { name: 'SampleProject' },
      repository: { name: 'SampleRepo' }
    }
  };
  const summary = summarizePayload(sample);
  if (!summary.includes('git.push') || !summary.includes('SampleRepo')) {
    throw new Error(`Unexpected summary output: ${summary}`);
  }
  console.log('Self-test passed:', summary);
};

const logRequest = (req, body) => {
  let parsed;
  try {
    parsed = body ? JSON.parse(body) : {};
  } catch (error) {
    console.warn('Received non-JSON payload:', error.message);
  }
  console.log(`[${new Date().toISOString()}] ${req.method} ${req.url}`);
  if (parsed) {
    console.log(' ', summarizePayload(parsed));
  }
  if (verbose && body) {
    console.log(' raw payload:', body);
  }
};

if (process.argv.includes('--self-test')) {
  runSelfTest();
  process.exit(0);
}

const server = http.createServer((req, res) => {
  const chunks = [];
  req.on('data', (chunk) => chunks.push(chunk));
  req.on('end', () => {
    const body = Buffer.concat(chunks).toString();
    logRequest(req, body);
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ status: 'ok' }));
  });
});

server.listen(port, () => {
  console.log(`Service Hook listener ready on http://localhost:${port}`);
  console.log('Press Ctrl+C to stop.');
});

process.on('SIGINT', () => {
  console.log('Shutting down listener...');
  server.close(() => process.exit(0));
});
