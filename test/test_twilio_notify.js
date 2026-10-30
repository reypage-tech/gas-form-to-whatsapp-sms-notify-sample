#!/usr/bin/env node
// Loads src/TwilioNotify.gs into a vm context with UrlFetchApp/PropertiesService
// stubs, since there is no live Apps Script runtime in this test environment.
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const GAS_PATH = path.join(__dirname, '..', 'src', 'TwilioNotify.gs');
const SOURCE = fs.readFileSync(GAS_PATH, 'utf8');

function makeContext({ properties, fetchResponse, appendedRows, sheets }) {
  const props = properties || {};
  const calls = { fetch: [] };

  const context = {
    PropertiesService: {
      getScriptProperties() {
        return {
          getProperty(key) {
            return Object.prototype.hasOwnProperty.call(props, key) ? props[key] : null;
          },
        };
      },
    },
    UrlFetchApp: {
      fetch(url, options) {
        calls.fetch.push({ url, options });
        const code = (fetchResponse && fetchResponse.code) || 200;
        const body = (fetchResponse && fetchResponse.body) || '{}';
        return {
          getResponseCode() { return code; },
          getContentText() { return body; },
        };
      },
    },
    Utilities: {
      base64Encode(s) { return Buffer.from(s, 'utf8').toString('base64'); },
    },
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return {
          getSheetByName(name) {
            return (sheets && sheets[name]) || {
              appendRow(row) { (appendedRows || []).push(row); },
            };
          },
        };
      },
    },
    console,
  };
  context.global = context;
  vm.createContext(context);
  vm.runInContext(SOURCE, context, { filename: 'TwilioNotify.gs' });
  return { context, calls };
}

function test(name, fn) {
  try {
    fn();
    console.log(`ok - ${name}`);
  } catch (err) {
    console.error(`FAIL - ${name}`);
    console.error(err);
    process.exitCode = 1;
  }
}

test('default DRY_RUN (property unset) never calls UrlFetchApp.fetch', () => {
  const { context, calls } = makeContext({ properties: {} });
  const result = context.sendTwilioMessage_('SMS', '+15550001111', 'hello');
  assert.strictEqual(calls.fetch.length, 0, 'fetch must not be called while dry-run');
  assert.strictEqual(result.dryRun, true);
  assert.strictEqual(result.result, 'dry-run');
});

test("DRY_RUN='TRUE' (not exact string 'false') still short-circuits", () => {
  const { calls, context } = makeContext({ properties: { DRY_RUN: 'TRUE' } });
  context.sendTwilioMessage_('SMS', '+15550001111', 'hello');
  assert.strictEqual(calls.fetch.length, 0);
});

test("DRY_RUN='false' + SMS calls fetch with correct URL, auth header, and body", () => {
  const { context, calls } = makeContext({
    properties: {
      DRY_RUN: 'false',
      TWILIO_ACCOUNT_SID: 'ACxxxx',
      TWILIO_AUTH_TOKEN: 'tokenxxx',
      TWILIO_FROM: '+15559998888',
    },
  });
  const result = context.sendTwilioMessage_('SMS', '+15550001111', 'hello');
  assert.strictEqual(calls.fetch.length, 1);
  const call = calls.fetch[0];
  assert.strictEqual(call.url, 'https://api.twilio.com/2010-04-01/Accounts/ACxxxx/Messages.json');
  assert.strictEqual(call.options.method, 'post');
  assert.strictEqual(
    call.options.headers.Authorization,
    'Basic ' + Buffer.from('ACxxxx:tokenxxx', 'utf8').toString('base64')
  );
  // payload objects are created inside the vm context, so compare by
  // own-property value rather than assert.deepStrictEqual's cross-realm
  // prototype check.
  assert.deepStrictEqual(JSON.parse(JSON.stringify(call.options.payload)), {
    To: '+15550001111',
    From: '+15559998888',
    Body: 'hello',
  });
  assert.strictEqual(result.dryRun, false);
  assert.strictEqual(result.result, 'sent');
});

test("DRY_RUN='false' + WhatsApp prefixes both To and From with 'whatsapp:'", () => {
  const { context, calls } = makeContext({
    properties: {
      DRY_RUN: 'false',
      TWILIO_ACCOUNT_SID: 'ACxxxx',
      TWILIO_AUTH_TOKEN: 'tokenxxx',
      TWILIO_FROM: '+15559998888',
    },
  });
  context.sendTwilioMessage_('WhatsApp', '+15550001111', 'hello');
  const call = calls.fetch[0];
  assert.deepStrictEqual(JSON.parse(JSON.stringify(call.options.payload)), {
    To: 'whatsapp:+15550001111',
    From: 'whatsapp:+15559998888',
    Body: 'hello',
  });
});

test('non-2xx response is recorded as failed, 2xx as sent', () => {
  const failing = makeContext({
    properties: {
      DRY_RUN: 'false',
      TWILIO_ACCOUNT_SID: 'ACxxxx',
      TWILIO_AUTH_TOKEN: 'tokenxxx',
      TWILIO_FROM: '+15559998888',
    },
    fetchResponse: { code: 500, body: '{"message":"error"}' },
  });
  const failResult = failing.context.sendTwilioMessage_('SMS', '+15550001111', 'hello');
  assert.strictEqual(failResult.result, 'failed');
  assert.strictEqual(failResult.responseCode, 500);

  const ok = makeContext({
    properties: {
      DRY_RUN: 'false',
      TWILIO_ACCOUNT_SID: 'ACxxxx',
      TWILIO_AUTH_TOKEN: 'tokenxxx',
      TWILIO_FROM: '+15559998888',
    },
    fetchResponse: { code: 201, body: '{}' },
  });
  const okResult = ok.context.sendTwilioMessage_('SMS', '+15550001111', 'hello');
  assert.strictEqual(okResult.result, 'sent');
  assert.strictEqual(okResult.responseCode, 201);
});

test('sendTwilioNotificationAndLog_ appends one Send Log row matching the 10-column schema', () => {
  const appendedRows = [];
  const { context } = makeContext({ properties: {}, appendedRows });
  const sheets = { 'Send Log': { appendRow(row) { appendedRows.push(row); } } };
  const ctx2 = makeContext({ properties: {}, sheets });
  const result = ctx2.context.sendTwilioNotificationAndLog_('Agent A', 'WhatsApp', '+15550001111', 'hello');
  assert.strictEqual(appendedRows.length, 1);
  assert.strictEqual(appendedRows[0].length, 10);
  assert.strictEqual(result.result, 'dry-run');
});

if (process.exitCode) {
  process.exit(process.exitCode);
}
console.log('test_twilio_notify.js: all checks passed');
