/**
 * Google Sheets setup, generic across jobs.
 * Everything job-specific (sheet names, columns, formulas, sample rows,
 * dashboard KPIs, chart) comes from JOB_CONFIG_JSON below, which
 * build_xlsx.py keeps in sync with job_config.json.
 *
 * This file does not call the WhatsApp Cloud API or the Twilio SMS API.
 * The Send Queue and Send Log sheets only hold mock data (mock export):
 * fictional recipients, a Mock Endpoint column naming the call that would
 * be made, and a Mock Response Code column standing in for a real API
 * response. There is no outbound HTTP call anywhere in this file.
 *
 * TwilioNotify.gs in this same project adds an optional, separate path
 * that can send real SMS/WhatsApp via the Twilio Messages API; it stays
 * in dry-run mode (no network call) unless a Script Property explicitly
 * turns it on. See that file for details.
 */

var INPUT_FILL = '#DDEBF7';

function columnToLetter_(index) {
  var letter = '';
  var n = index;
  while (n > 0) {
    var rem = (n - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}

function sheetOrder_(config) {
  if (config.sheet_order) {
    return config.sheet_order;
  }
  var names = [];
  (config.data_sheets || []).forEach(function (s) { names.push(s.name); });
  if (config.dashboard) {
    names.push(config.dashboard.name);
  }
  if (config.lists) {
    names.push(config.lists.sheet_name);
  }
  return names;
}

function setupWorkbook() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = loadJobConfig_();
  var order = sheetOrder_(config);
  order.forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (sh) {
      ss.deleteSheet(sh);
    }
    ss.insertSheet(name);
  });

  var sheetMeta = {};
  (config.data_sheets || []).forEach(function (sheetCfg) {
    sheetMeta[sheetCfg.name] = buildDataSheet_(ss.getSheetByName(sheetCfg.name), sheetCfg);
  });

  if (config.lists) {
    buildLists_(ss.getSheetByName(config.lists.sheet_name), config.lists);
    ss.getSheetByName(config.lists.sheet_name).hideSheet();
  }

  if (config.dashboard) {
    buildDashboard_(ss.getSheetByName(config.dashboard.name), config.dashboard, sheetMeta, config.lists);
  }

  order.forEach(function (name) {
    protectSheet_(ss.getSheetByName(name), config, name);
  });

  SpreadsheetApp.getUi().alert((config.project && config.project.title) + ' is ready.');
}

function buildLists_(sheet, listsCfg) {
  var names = Object.keys(listsCfg.columns);
  sheet.getRange(1, 1, 1, names.length).setValues([names]).setFontWeight('bold');
  var maxLen = 0;
  names.forEach(function (n) { maxLen = Math.max(maxLen, listsCfg.columns[n].length); });
  var rows = [];
  for (var i = 0; i < maxLen; i++) {
    rows.push(names.map(function (n) {
      var values = listsCfg.columns[n];
      return i < values.length ? values[i] : '';
    }));
  }
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, names.length).setValues(rows);
  }
}

function resolveRowFormula_(template, colLetter, row) {
  var expr = template;
  Object.keys(colLetter).forEach(function (header) {
    expr = expr.split('{' + header + '}').join(colLetter[header] + row);
  });
  if (expr.charAt(0) !== '=') {
    expr = '=' + expr;
  }
  return expr;
}

function buildDataSheet_(sheet, cfg) {
  var headers = cfg.columns.map(function (c) { return c.header; });
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  var colLetter = {};
  headers.forEach(function (h, idx) { colLetter[h] = columnToLetter_(idx + 1); });

  var rows = cfg.rows || [];
  rows.forEach(function (row, ridx) {
    var r = ridx + 2;
    cfg.columns.forEach(function (colCfg, cidx) {
      var cell = sheet.getRange(r, cidx + 1);
      if (colCfg.kind === 'input') {
        var val = row[colCfg.header];
        cell.setValue(val === undefined ? '' : val);
        cell.setBackground(INPUT_FILL);
      } else if (colCfg.kind === 'formula') {
        cell.setFormula(resolveRowFormula_(colCfg.formula, colLetter, r));
      }
      if (colCfg.number_format) {
        cell.setNumberFormat(colCfg.number_format);
      } else if (colCfg.date) {
        cell.setNumberFormat('yyyy-mm-dd');
      }
    });
  });

  var lastDataRow = 1 + rows.length;
  var totalRow = null;
  if (cfg.total_row) {
    var tr = cfg.total_row;
    totalRow = lastDataRow + 1;
    var labelIdx = headers.indexOf(tr.label_column) + 1;
    sheet.getRange(totalRow, labelIdx).setValue(tr.label).setFontWeight('bold');
    tr.sum_columns.forEach(function (h) {
      var cidx = headers.indexOf(h) + 1;
      var letter = colLetter[h];
      sheet.getRange(totalRow, cidx)
        .setFormula('=SUM(' + letter + '2:' + letter + lastDataRow + ')')
        .setFontWeight('bold');
    });
  }

  return { headers: headers, colLetter: colLetter, lastDataRow: lastDataRow, totalRow: totalRow };
}

function buildDashboard_(sheet, dashCfg, sheetMeta, listsCfg) {
  sheet.getRange('A1').setValue(dashCfg.heading || dashCfg.name).setFontWeight('bold').setFontSize(14);

  var kpiAddr = {};
  var row = 3;
  (dashCfg.kpis || []).forEach(function (kpi) {
    sheet.getRange(row, 1).setValue(kpi.label).setFontWeight('bold');
    var valueCell = sheet.getRange(row, 2);
    if (kpi.kind === 'sum') {
      var meta = sheetMeta[kpi.sheet];
      var letter = meta.colLetter[kpi.column];
      valueCell.setFormula(
        "=SUM('" + kpi.sheet + "'!" + letter + '2:' + letter + meta.lastDataRow + ')'
      );
    } else if (kpi.kind === 'sum_if') {
      var m2 = sheetMeta[kpi.sheet];
      var sumLetter = m2.colLetter[kpi.column];
      var filtLetter = m2.colLetter[kpi.filter_column];
      valueCell.setFormula(
        "=SUMIFS('" + kpi.sheet + "'!" + sumLetter + '2:' + sumLetter + m2.lastDataRow + ',' +
        "'" + kpi.sheet + "'!" + filtLetter + '2:' + filtLetter + m2.lastDataRow + ',"' +
        kpi.filter_value + '")'
      );
    } else if (kpi.kind === 'input') {
      valueCell.setValue(kpi.value);
      valueCell.setBackground(INPUT_FILL);
    } else if (kpi.kind === 'expr') {
      var expr = kpi.formula;
      Object.keys(kpiAddr).forEach(function (kid) {
        expr = expr.split('{' + kid + '}').join(kpiAddr[kid]);
      });
      valueCell.setFormula(expr.charAt(0) === '=' ? expr : '=' + expr);
    }
    if (kpi.number_format) {
      valueCell.setNumberFormat(kpi.number_format);
    }
    kpiAddr[kpi.id] = 'B' + row;
    row += 1;
  });

  if (dashCfg.chart) {
    buildChart_(sheet, dashCfg.chart, sheetMeta, listsCfg);
  }
}

function buildChart_(sheet, chartCfg, sheetMeta, listsCfg) {
  var groupValues = listsCfg ? listsCfg.columns[chartCfg.group_list] : [];
  var srcMeta = sheetMeta[chartCfg.source_sheet];
  var groupLetter = srcMeta.colLetter[chartCfg.group_column];
  var srcLast = srcMeta.lastDataRow;

  var headerRow = 5;
  var groupCol = 4; // column D
  sheet.getRange(headerRow, groupCol).setValue(chartCfg.group_column).setFontWeight('bold');
  chartCfg.series.forEach(function (series, si) {
    sheet.getRange(headerRow, groupCol + 1 + si).setValue(series.label).setFontWeight('bold');
  });

  var dataStart = headerRow + 1;
  groupValues.forEach(function (groupValue, gi) {
    var r = dataStart + gi;
    sheet.getRange(r, groupCol).setValue(groupValue);
    chartCfg.series.forEach(function (series, si) {
      var sumLetter = srcMeta.colLetter[series.sum_column];
      var sq = "'" + chartCfg.source_sheet + "'!";
      sheet.getRange(r, groupCol + 1 + si).setFormula(
        '=SUMIFS(' + sq + sumLetter + '2:' + sumLetter + srcLast + ',' +
        sq + groupLetter + '2:' + groupLetter + srcLast + ',' +
        columnToLetter_(groupCol) + r + ')'
      );
    });
  });
  var dataEnd = dataStart + groupValues.length - 1;
  if (dataEnd < dataStart) {
    return;
  }

  var range = sheet.getRange(headerRow, groupCol, dataEnd - headerRow + 1, 1 + chartCfg.series.length);
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(range)
    .setPosition(dataEnd + 3, groupCol, 0, 0)
    .setOption('title', chartCfg.title || '')
    .build();
  sheet.insertChart(chart);
}

function protectSheet_(sheet, config, name) {
  var maxRows = Math.max(sheet.getLastRow() + 5, 40);
  var maxCols = 14;
  sheet.getRange(1, 1, maxRows, maxCols).setLocked(true);

  var dataCfg = (config.data_sheets || []).filter(function (s) { return s.name === name; })[0];
  if (dataCfg) {
    var lastDataRow = 1 + (dataCfg.rows || []).length;
    dataCfg.columns.forEach(function (colCfg, idx) {
      if (colCfg.kind === 'input') {
        var letter = columnToLetter_(idx + 1);
        sheet.getRange(letter + '2:' + letter + lastDataRow).setLocked(false);
      }
    });
  }
  if (config.dashboard && config.dashboard.name === name) {
    (config.dashboard.kpis || []).forEach(function (kpi, idx) {
      if (kpi.kind === 'input') {
        sheet.getRange('B' + (3 + idx)).setLocked(false);
      }
    });
  }
  sheet.protect();
}

function loadJobConfig_() {
  return JSON.parse(JOB_CONFIG_JSON);
}

function highlightOverruns_() {
  var config = loadJobConfig_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  (config.data_sheets || []).forEach(function (sheetCfg) {
    if (!sheetCfg.overrun_check) {
      return;
    }
    var sheet = ss.getSheetByName(sheetCfg.name);
    if (!sheet) {
      return;
    }
    var headers = sheetCfg.columns.map(function (c) { return c.header; });
    var estIdx = headers.indexOf(sheetCfg.overrun_check.estimate_column) + 1;
    var actIdx = headers.indexOf(sheetCfg.overrun_check.actual_column) + 1;
    var numRows = (sheetCfg.rows || []).length;
    if (numRows === 0) {
      return;
    }
    var estimates = sheet.getRange(2, estIdx, numRows, 1).getValues();
    var actuals = sheet.getRange(2, actIdx, numRows, 1).getValues();
    var backgrounds = [];
    for (var i = 0; i < numRows; i++) {
      var e = Number(estimates[i][0]) || 0;
      var a = Number(actuals[i][0]) || 0;
      var color = '#ffffff';
      if (e <= 0) {
        color = a > 0 ? '#f4cccc' : '#ffffff';
      } else if (a > e) {
        color = '#f4cccc';
      } else if (a > e * 0.9) {
        color = '#fce5cd';
      }
      var rowColors = [];
      for (var c = 0; c < headers.length; c++) {
        rowColors.push(color);
      }
      backgrounds.push(rowColors);
    }
    sheet.getRange(2, 1, numRows, headers.length).setBackgrounds(backgrounds);
  });
}

function recalculateSummaries() {
  SpreadsheetApp.flush();
  highlightOverruns_();
}

function onEdit(e) {
  highlightOverruns_();
}

function onOpen() {
  var config = loadJobConfig_();
  SpreadsheetApp.getUi()
    .createMenu(config.project && config.project.title ? config.project.title : 'Job Template')
    .addItem('Setup sample workbook', 'setupWorkbook')
    .addItem('Recalculate summaries', 'recalculateSummaries')
    .addToUi();
}

// JOB_CONFIG_JSON_START
var JOB_CONFIG_JSON = "{\"project\": {\"title\": \"Form-to-WhatsApp/SMS Notify Queue (sample)\", \"slug\": \"form_to_whatsapp_sms_notify_queue\", \"author\": \"demo-author\"}, \"sheet_order\": [\"Dashboard\", \"Send Queue\", \"Send Log\", \"Lists\"], \"lists\": {\"sheet_name\": \"Lists\", \"columns\": {\"Agents\": [\"Agent A\", \"Agent B\", \"Agent C\"], \"Channels\": [\"WhatsApp\", \"SMS\"], \"QueueStatuses\": [\"Queued\", \"Sent (mock)\", \"Failed (mock)\"], \"LogResults\": [\"Sent (mock)\", \"Failed (mock)\"], \"InquiryTypes\": [\"Order Status\", \"Product Question\", \"Support Request\", \"Billing Question\"]}}, \"data_sheets\": [{\"name\": \"Send Queue\", \"columns\": [{\"header\": \"Date\", \"kind\": \"input\", \"date\": true}, {\"header\": \"Customer Name\", \"kind\": \"input\"}, {\"header\": \"Phone (fictional)\", \"kind\": \"input\"}, {\"header\": \"Inquiry Type\", \"kind\": \"input\", \"list\": \"InquiryTypes\"}, {\"header\": \"Channel\", \"kind\": \"input\", \"list\": \"Channels\"}, {\"header\": \"Agent\", \"kind\": \"input\", \"list\": \"Agents\"}, {\"header\": \"Message Draft\", \"kind\": \"input\"}, {\"header\": \"Status\", \"kind\": \"input\", \"list\": \"QueueStatuses\"}, {\"header\": \"Retry Count\", \"kind\": \"input\", \"number_format\": \"#,##0\"}, {\"header\": \"Count\", \"kind\": \"input\", \"number_format\": \"#,##0\"}], \"rows\": [{\"Date\": \"2026-10-20\", \"Customer Name\": \"Customer 1\", \"Phone (fictional)\": \"+1-555-0101\", \"Inquiry Type\": \"Order Status\", \"Channel\": \"WhatsApp\", \"Agent\": \"Agent A\", \"Message Draft\": \"Hi, your order question has been received. Our team will follow up shortly.\", \"Status\": \"Sent (mock)\", \"Retry Count\": 0, \"Count\": 1}, {\"Date\": \"2026-10-20\", \"Customer Name\": \"Customer 2\", \"Phone (fictional)\": \"+1-555-0102\", \"Inquiry Type\": \"Product Question\", \"Channel\": \"SMS\", \"Agent\": \"Agent B\", \"Message Draft\": \"Hi, thanks for your product question. We will reply with details soon.\", \"Status\": \"Sent (mock)\", \"Retry Count\": 0, \"Count\": 1}, {\"Date\": \"2026-10-21\", \"Customer Name\": \"Customer 3\", \"Phone (fictional)\": \"+1-555-0103\", \"Inquiry Type\": \"Support Request\", \"Channel\": \"WhatsApp\", \"Agent\": \"Agent C\", \"Message Draft\": \"Hi, we received your support request and a specialist will reach out.\", \"Status\": \"Failed (mock)\", \"Retry Count\": 1, \"Count\": 1}, {\"Date\": \"2026-10-21\", \"Customer Name\": \"Customer 4\", \"Phone (fictional)\": \"+1-555-0104\", \"Inquiry Type\": \"Billing Question\", \"Channel\": \"SMS\", \"Agent\": \"Agent A\", \"Message Draft\": \"Hi, thanks for your billing question. We are checking your account now.\", \"Status\": \"Queued\", \"Retry Count\": 0, \"Count\": 1}, {\"Date\": \"2026-10-22\", \"Customer Name\": \"Customer 5\", \"Phone (fictional)\": \"+1-555-0105\", \"Inquiry Type\": \"Order Status\", \"Channel\": \"WhatsApp\", \"Agent\": \"Agent B\", \"Message Draft\": \"Hi, your order question has been received. Our team will follow up shortly.\", \"Status\": \"Queued\", \"Retry Count\": 0, \"Count\": 1}, {\"Date\": \"2026-10-22\", \"Customer Name\": \"Customer 6\", \"Phone (fictional)\": \"+1-555-0106\", \"Inquiry Type\": \"Product Question\", \"Channel\": \"SMS\", \"Agent\": \"Agent C\", \"Message Draft\": \"Hi, thanks for your product question. We will reply with details soon.\", \"Status\": \"Failed (mock)\", \"Retry Count\": 2, \"Count\": 1}, {\"Date\": \"2026-10-23\", \"Customer Name\": \"Customer 7\", \"Phone (fictional)\": \"+1-555-0107\", \"Inquiry Type\": \"Support Request\", \"Channel\": \"WhatsApp\", \"Agent\": \"Agent A\", \"Message Draft\": \"Hi, we received your support request and a specialist will reach out.\", \"Status\": \"Sent (mock)\", \"Retry Count\": 1, \"Count\": 1}, {\"Date\": \"2026-10-23\", \"Customer Name\": \"Customer 8\", \"Phone (fictional)\": \"+1-555-0108\", \"Inquiry Type\": \"Billing Question\", \"Channel\": \"SMS\", \"Agent\": \"Agent B\", \"Message Draft\": \"Hi, thanks for your billing question. We are checking your account now.\", \"Status\": \"Queued\", \"Retry Count\": 0, \"Count\": 1}, {\"Date\": \"2026-10-24\", \"Customer Name\": \"Customer 9\", \"Phone (fictional)\": \"+1-555-0109\", \"Inquiry Type\": \"Order Status\", \"Channel\": \"SMS\", \"Agent\": \"Agent C\", \"Message Draft\": \"Hi, your order question has been received. Our team will follow up shortly.\", \"Status\": \"Queued\", \"Retry Count\": 0, \"Count\": 1}, {\"Date\": \"2026-10-24\", \"Customer Name\": \"Customer 10\", \"Phone (fictional)\": \"+1-555-0110\", \"Inquiry Type\": \"Support Request\", \"Channel\": \"WhatsApp\", \"Agent\": \"Agent A\", \"Message Draft\": \"Hi, we received your support request and a specialist will reach out.\", \"Status\": \"Failed (mock)\", \"Retry Count\": 1, \"Count\": 1}], \"total_row\": {\"label\": \"Total\", \"label_column\": \"Message Draft\", \"sum_columns\": [\"Retry Count\", \"Count\"]}, \"column_widths\": {\"A\": 11, \"B\": 16, \"C\": 16, \"D\": 15, \"E\": 10, \"F\": 10, \"G\": 66, \"H\": 18, \"I\": 13, \"J\": 8}}, {\"name\": \"Send Log\", \"columns\": [{\"header\": \"Attempted At\", \"kind\": \"input\"}, {\"header\": \"Agent\", \"kind\": \"input\", \"list\": \"Agents\"}, {\"header\": \"Channel\", \"kind\": \"input\", \"list\": \"Channels\"}, {\"header\": \"Recipient (fictional)\", \"kind\": \"input\"}, {\"header\": \"Message Preview\", \"kind\": \"input\"}, {\"header\": \"Mock Endpoint\", \"kind\": \"input\"}, {\"header\": \"Result\", \"kind\": \"input\", \"list\": \"LogResults\"}, {\"header\": \"Mock Response Code\", \"kind\": \"input\"}, {\"header\": \"Retry Count\", \"kind\": \"input\", \"number_format\": \"#,##0\"}, {\"header\": \"Count\", \"kind\": \"input\", \"number_format\": \"#,##0\"}], \"rows\": [{\"Attempted At\": \"2026-10-20 09:15\", \"Agent\": \"Agent A\", \"Channel\": \"WhatsApp\", \"Recipient (fictional)\": \"+1-555-0101\", \"Message Preview\": \"Hi, your order question has been received...\", \"Mock Endpoint\": \"mock:whatsapp.sendMessage\", \"Result\": \"Sent (mock)\", \"Mock Response Code\": \"200 (mock)\", \"Retry Count\": 0, \"Count\": 1}, {\"Attempted At\": \"2026-10-20 09:40\", \"Agent\": \"Agent B\", \"Channel\": \"SMS\", \"Recipient (fictional)\": \"+1-555-0102\", \"Message Preview\": \"Hi, thanks for your product question...\", \"Mock Endpoint\": \"mock:sms.send\", \"Result\": \"Sent (mock)\", \"Mock Response Code\": \"200 (mock)\", \"Retry Count\": 0, \"Count\": 1}, {\"Attempted At\": \"2026-10-21 10:05\", \"Agent\": \"Agent C\", \"Channel\": \"WhatsApp\", \"Recipient (fictional)\": \"+1-555-0103\", \"Message Preview\": \"Hi, we received your support request...\", \"Mock Endpoint\": \"mock:whatsapp.sendMessage\", \"Result\": \"Failed (mock)\", \"Mock Response Code\": \"500 (mock)\", \"Retry Count\": 1, \"Count\": 1}, {\"Attempted At\": \"2026-10-22 11:20\", \"Agent\": \"Agent C\", \"Channel\": \"SMS\", \"Recipient (fictional)\": \"+1-555-0106\", \"Message Preview\": \"Hi, thanks for your product question...\", \"Mock Endpoint\": \"mock:sms.send\", \"Result\": \"Failed (mock)\", \"Mock Response Code\": \"500 (mock)\", \"Retry Count\": 1, \"Count\": 1}, {\"Attempted At\": \"2026-10-22 15:50\", \"Agent\": \"Agent C\", \"Channel\": \"SMS\", \"Recipient (fictional)\": \"+1-555-0106\", \"Message Preview\": \"Hi, thanks for your product question...\", \"Mock Endpoint\": \"mock:sms.send\", \"Result\": \"Failed (mock)\", \"Mock Response Code\": \"500 (mock)\", \"Retry Count\": 2, \"Count\": 1}, {\"Attempted At\": \"2026-10-23 08:30\", \"Agent\": \"Agent A\", \"Channel\": \"WhatsApp\", \"Recipient (fictional)\": \"+1-555-0107\", \"Message Preview\": \"Hi, we received your support request...\", \"Mock Endpoint\": \"mock:whatsapp.sendMessage\", \"Result\": \"Failed (mock)\", \"Mock Response Code\": \"500 (mock)\", \"Retry Count\": 1, \"Count\": 1}, {\"Attempted At\": \"2026-10-23 08:55\", \"Agent\": \"Agent A\", \"Channel\": \"WhatsApp\", \"Recipient (fictional)\": \"+1-555-0107\", \"Message Preview\": \"Hi, we received your support request...\", \"Mock Endpoint\": \"mock:whatsapp.sendMessage\", \"Result\": \"Sent (mock)\", \"Mock Response Code\": \"200 (mock)\", \"Retry Count\": 1, \"Count\": 1}, {\"Attempted At\": \"2026-10-24 09:10\", \"Agent\": \"Agent A\", \"Channel\": \"WhatsApp\", \"Recipient (fictional)\": \"+1-555-0110\", \"Message Preview\": \"Hi, we received your support request...\", \"Mock Endpoint\": \"mock:whatsapp.sendMessage\", \"Result\": \"Failed (mock)\", \"Mock Response Code\": \"500 (mock)\", \"Retry Count\": 1, \"Count\": 1}], \"total_row\": {\"label\": \"Total\", \"label_column\": \"Message Preview\", \"sum_columns\": [\"Retry Count\", \"Count\"]}, \"column_widths\": {\"A\": 17, \"B\": 10, \"C\": 10, \"D\": 18, \"E\": 40, \"F\": 26, \"G\": 16, \"H\": 20, \"I\": 12, \"J\": 8}}], \"dashboard\": {\"name\": \"Dashboard\", \"heading\": \"Send queue dashboard (mock data only)\", \"kpis\": [{\"id\": \"total_items\", \"kind\": \"sum\", \"label\": \"Total items in queue\", \"sheet\": \"Send Queue\", \"column\": \"Count\", \"number_format\": \"#,##0\"}, {\"id\": \"unsent_count\", \"kind\": \"sum_if\", \"label\": \"Unsent (status = Queued)\", \"sheet\": \"Send Queue\", \"column\": \"Count\", \"filter_column\": \"Status\", \"filter_value\": \"Queued\", \"number_format\": \"#,##0\"}, {\"id\": \"failed_count\", \"kind\": \"sum_if\", \"label\": \"Failed, needs retry\", \"sheet\": \"Send Queue\", \"column\": \"Count\", \"filter_column\": \"Status\", \"filter_value\": \"Failed (mock)\", \"number_format\": \"#,##0\"}, {\"id\": \"sent_count\", \"kind\": \"sum_if\", \"label\": \"Sent (mock)\", \"sheet\": \"Send Queue\", \"column\": \"Count\", \"filter_column\": \"Status\", \"filter_value\": \"Sent (mock)\", \"number_format\": \"#,##0\"}, {\"id\": \"backlog\", \"kind\": \"expr\", \"label\": \"Backlog (unsent + failed)\", \"formula\": \"{unsent_count}+{failed_count}\", \"number_format\": \"#,##0\"}, {\"id\": \"total_retries\", \"kind\": \"sum\", \"label\": \"Total retry attempts (queue)\", \"sheet\": \"Send Queue\", \"column\": \"Retry Count\", \"number_format\": \"#,##0\"}, {\"id\": \"logged_attempts\", \"kind\": \"sum\", \"label\": \"Mock send attempts logged\", \"sheet\": \"Send Log\", \"column\": \"Count\", \"number_format\": \"#,##0\"}, {\"id\": \"logged_failed\", \"kind\": \"sum_if\", \"label\": \"Mock send attempts failed\", \"sheet\": \"Send Log\", \"column\": \"Count\", \"filter_column\": \"Result\", \"filter_value\": \"Failed (mock)\", \"number_format\": \"#,##0\"}], \"chart\": {\"title\": \"Queue items and retries by agent\", \"y_axis_title\": \"Count\", \"group_list\": \"Agents\", \"source_sheet\": \"Send Queue\", \"group_column\": \"Agent\", \"series\": [{\"label\": \"Items in queue\", \"sum_column\": \"Count\"}, {\"label\": \"Retry attempts\", \"sum_column\": \"Retry Count\"}], \"anchor\": \"D15\"}, \"column_widths\": {\"A\": 32, \"B\": 14, \"D\": 14, \"E\": 14, \"F\": 14}}, \"screens\": {\"sheets\": [\"Dashboard\", \"Send Queue\", \"Send Log\"]}, \"readme\": {\"title\": \"Form-to-WhatsApp/SMS Notify Queue (sample)\", \"who_it_is_for\": \"For support or sales teams who want to see how Google Form responses could turn into a per-agent WhatsApp/SMS follow-up queue.\", \"effect\": \"Replaces tracking form follow-ups by hand with one queue that shows each message draft, assigned agent, send status and retry count.\", \"details\": \"Each form response becomes a queue row with a message draft, an assigned agent, a send status, and a retry count. A send log tracks attempts, and the dashboard shows unsent and failed counts across the whole queue. An optional Twilio Messages API path can send real SMS/WhatsApp messages, but it stays in dry-run mode (no network call) until a Script Property is explicitly switched on.\", \"tabs\": [{\"name\": \"Dashboard\", \"purpose\": \"Queue totals, unsent/failed/sent counts, backlog, and a per-agent chart\"}, {\"name\": \"Send Queue\", \"purpose\": \"One row per fictional form response, with message draft, status, and retry count\"}, {\"name\": \"Send Log\", \"purpose\": \"Mock send attempts with a fictional recipient, a mock endpoint name, and a mock response code\"}, {\"name\": \"Lists (hidden)\", \"purpose\": \"Agents, channels, statuses, and inquiry types used for dropdown validation\"}], \"how_to_try\": \"Build the workbook locally, open dist/form_to_whatsapp_sms_notify_queue.xlsx, and edit the light blue cells on Send Queue or Send Log (for example, change a Status to Queued) to see the Dashboard counts and chart update.\", \"twilio_setup\": \"This sample includes an optional `src/TwilioNotify.gs` path that can send real SMS/WhatsApp messages through the Twilio Messages API. It defaults to a dry run.\\n\\nScript Properties (Apps Script editor: Project Settings > Script Properties):\\n\\n- `TWILIO_ACCOUNT_SID`\\n- `TWILIO_AUTH_TOKEN`\\n- `TWILIO_FROM` (your Twilio-verified sending number, e.g. `+15551234567`)\\n- `DRY_RUN` (defaults to dry run; set to the exact string `false` to send for real)\\n\\nWhile `DRY_RUN` is unset or anything other than the exact string `false`, no HTTP request is made; the Send Log sheet gets a `dry-run` row instead. No Twilio keys are stored in this repository.\"}}";
// JOB_CONFIG_JSON_END
