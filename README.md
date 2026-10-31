# Form-to-WhatsApp/SMS Notify Queue (sample)

![Dashboard](screens/01_dashboard.png)

Replaces tracking form follow-ups by hand with one queue that shows each message draft, assigned agent, send status and retry count.

For support or sales teams who want to see how Google Form responses could turn into a per-agent WhatsApp/SMS follow-up queue.

Each form response becomes a queue row with a message draft, an assigned agent, a send status, and a retry count. A send log tracks attempts, and the dashboard shows unsent and failed counts across the whole queue. An optional Twilio Messages API path can send real SMS/WhatsApp messages, but it stays in dry-run mode (no network call) until a Script Property is explicitly switched on.

![Send Queue](screens/02_send_queue.png)
![Send Log](screens/03_send_log.png)

## Tabs

| Tab | Purpose |
|-----|---------|
| Dashboard | Queue totals, unsent/failed/sent counts, backlog, and a per-agent chart |
| Send Queue | One row per fictional form response, with message draft, status, and retry count |
| Send Log | Mock send attempts with a fictional recipient, a mock endpoint name, and a mock response code |
| Lists (hidden) | Agents, channels, statuses, and inquiry types used for dropdown validation |

## How to try it

Build the workbook locally, open dist/form_to_whatsapp_sms_notify_queue.xlsx, and edit the light blue cells on Send Queue or Send Log (for example, change a Status to Queued) to see the Dashboard counts and chart update.

```bash
python3 tools/check_sheet_names.py job_config.json
python3 build_xlsx.py
python3 test/test_xlsx.py
node test/test_twilio_notify.js
python3 tools/make_screens.py
```

`build_xlsx.py` writes the workbook to `dist/` and keeps `src/Code.gs` in sync with `job_config.json` for the Google Sheets setup path.

## Twilio setup (optional, real send)

This sample includes an optional `src/TwilioNotify.gs` path that can send real SMS/WhatsApp messages through the Twilio Messages API. It defaults to a dry run.

Script Properties (Apps Script editor: Project Settings > Script Properties):

- `TWILIO_ACCOUNT_SID`
- `TWILIO_AUTH_TOKEN`
- `TWILIO_FROM` (your Twilio-verified sending number, e.g. `+15551234567`)
- `DRY_RUN` (defaults to dry run; set to the exact string `false` to send for real)

While `DRY_RUN` is unset or anything other than the exact string `false`, no HTTP request is made; the Send Log sheet gets a `dry-run` row instead. No Twilio keys are stored in this repository.

Note: built and tested with fictional data and local mocks; no WhatsApp Cloud API or Twilio call is made unless DRY_RUN is switched off; not used in production.
