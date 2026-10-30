/**
 * Optional Twilio Messages API sender (SMS and WhatsApp), separate from
 * Code.gs on purpose: Code.gs's own header comment says that file makes no
 * outbound HTTP calls, and this one does when explicitly enabled.
 *
 * Credentials and the DRY_RUN switch live in Script Properties
 * (Project Settings > Script Properties in the Apps Script editor), never
 * in this file or in README.md:
 *   - TWILIO_ACCOUNT_SID
 *   - TWILIO_AUTH_TOKEN
 *   - TWILIO_FROM      (the Twilio-verified sending number, e.g. +15551234567)
 *   - DRY_RUN          (defaults to dry run; set to the exact string
 *                        "false" to actually call the Twilio API)
 *
 * While DRY_RUN is unset or any value other than the exact string "false",
 * sendTwilioMessage_ never calls UrlFetchApp.fetch; it only reports what
 * would have been sent, and sendTwilioNotificationAndLog_ writes that as a
 * "dry-run" row to the Send Log sheet.
 */

var TWILIO_API_BASE_ = 'https://api.twilio.com/2010-04-01/Accounts/';

function twilioConfig_() {
  var props = PropertiesService.getScriptProperties();
  return {
    accountSid: props.getProperty('TWILIO_ACCOUNT_SID'),
    authToken: props.getProperty('TWILIO_AUTH_TOKEN'),
    from: props.getProperty('TWILIO_FROM'),
    dryRun: props.getProperty('DRY_RUN') !== 'false'
  };
}

function twilioIsWhatsApp_(channel) {
  return String(channel).toLowerCase() === 'whatsapp';
}

function twilioRecipient_(channel, phone) {
  return twilioIsWhatsApp_(channel) ? 'whatsapp:' + phone : phone;
}

function twilioSender_(channel, from) {
  if (twilioIsWhatsApp_(channel) && String(from).indexOf('whatsapp:') !== 0) {
    return 'whatsapp:' + from;
  }
  return from;
}

function twilioAuthHeader_(accountSid, authToken) {
  return 'Basic ' + Utilities.base64Encode(accountSid + ':' + authToken);
}

/**
 * Builds the Twilio request for one message and, unless DRY_RUN is the
 * exact string "false", sends it via UrlFetchApp.fetch.
 * channel is "SMS" or "WhatsApp" (case-insensitive).
 */
function sendTwilioMessage_(channel, toPhone, body) {
  var cfg = twilioConfig_();
  var url = TWILIO_API_BASE_ + cfg.accountSid + '/Messages.json';
  var to = twilioRecipient_(channel, toPhone);
  var from = twilioSender_(channel, cfg.from);

  if (cfg.dryRun) {
    return { dryRun: true, url: url, to: to, from: from, body: body, responseCode: null, result: 'dry-run' };
  }

  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    payload: { To: to, From: from, Body: body },
    headers: { Authorization: twilioAuthHeader_(cfg.accountSid, cfg.authToken) },
    muteHttpExceptions: true
  });
  var code = response.getResponseCode();
  var result = (code >= 200 && code < 300) ? 'sent' : 'failed';
  return { dryRun: false, url: url, to: to, from: from, body: body, responseCode: code, result: result };
}

function logTwilioAttempt_(row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Send Log');
  if (!sheet) {
    return;
  }
  sheet.appendRow([
    new Date(),
    row.agent || '',
    row.channel,
    row.to,
    String(row.body || '').slice(0, 80),
    row.dryRun ? ('dry-run:' + row.url) : row.url,
    row.result,
    row.dryRun ? 'dry-run' : String(row.responseCode),
    0,
    1
  ]);
}

/** Sends (or dry-runs) one message and logs the attempt to the Send Log sheet. */
function sendTwilioNotificationAndLog_(agent, channel, toPhone, body) {
  var result = sendTwilioMessage_(channel, toPhone, body);
  logTwilioAttempt_({
    agent: agent,
    channel: channel,
    to: result.to,
    body: body,
    url: result.url,
    dryRun: result.dryRun,
    result: result.result,
    responseCode: result.responseCode
  });
  return result;
}
