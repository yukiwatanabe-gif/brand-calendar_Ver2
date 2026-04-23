// ============================================================
// スプレッドシートID（2026_campaigns_import）
// ============================================================
const SPREADSHEET_ID = '1rTa6wzZvgow63PBR-TpF5wsP6H_QTR36ji_SC3bRPgM';

function doGet_v3_legacy(e) {
  // ── デバッグ（?type=debug） ──────────────────────
  if ((e.parameter.type || '') === 'debug') {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = ss.getSheets().map(s => s.getName());
    const active = ss.getActiveSheet().getName();
    const target = ss.getSheetByName('Export_Master');
    return ContentService.createTextOutput(JSON.stringify({
      spreadsheetId: SPREADSHEET_ID,
      spreadsheetName: ss.getName(),
      sheets: sheets,
      activeSheet: active,
      exportMasterFound: target !== null,
      exportMasterRows: target ? target.getLastRow() : 0,
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // ── KPI データ返却（?type=kpi） ──────────────────
  if ((e.parameter.type || '') === 'kpi') {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('KPI');
    if (!sheet) {
      return ContentService.createTextOutput('[]')
        .setMimeType(ContentService.MimeType.JSON);
    }
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return ContentService.createTextOutput('[]')
        .setMimeType(ContentService.MimeType.JSON);
    }
    const h = data[0].map(v => String(v).trim().toLowerCase());
    const ci = h.indexOf('channel'), li = h.indexOf('label');
    const ti = h.indexOf('target'),  ai = h.indexOf('actual');
    const yi = h.indexOf('yoy');
    const result = [];
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      const ch = String(r[ci] || '').trim();
      const lb = String(r[li] || '').trim();
      if (!ch || !lb) continue;
      result.push({
        channel: ch, label: lb,
        target:  Number(r[ti]) || 0,
        actual:  Number(r[ai]) || 0,
        yoy:     (yi >= 0 && r[yi] !== '') ? (Number(r[yi]) || null) : null,
      });
    }
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Export_Master') || ss.getActiveSheet();
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const data = rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost_v3_legacy(e) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Manual_Input') || ss.getActiveSheet();
  const data = JSON.parse(e.postData.contents);

  if (data.action === 'delete') {
    sheet.deleteRow(data._row);
    return ContentService
      .createTextOutput(JSON.stringify({status:'ok'}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (data.action === 'update') {
    sheet.getRange(data._row, 1, 1, 7).setValues([[
      data.title, data.start, data.end||'',
      data.type, data.team, data.memo||'', data.tentative||''
    ]]);
    return ContentService
      .createTextOutput(JSON.stringify({status:'ok'}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  sheet.appendRow([
    data.title, data.start, data.end||'',
    data.type, data.team, data.memo||'', data.tentative||''
  ]);
  return ContentService
    .createTextOutput(JSON.stringify({status:'ok'}))
    .setMimeType(ContentService.MimeType.JSON);
}
// ============================================================
// Backlog 同期設定（ここを書き換えてください）
// ============================================================
const BACKLOG_SPACE_ID   = 'yogibo';   // 例: 'yogibo'
const BACKLOG_API_KEY    = PropertiesService.getScriptProperties().getProperty('BACKLOG_API_KEY') || '';
const BACKLOG_PROJECT_ID = 447191;             // 数値のプロジェクトID
const BACKLOG_ISSUE_COUNT = 100;

const DEFAULT_TYPE = 'event';
const DEFAULT_TEAM = 'promo';
const SHEET_BACKLOG = 'Backlog_Data';

// ============================================================
// メイン関数（トリガーから呼び出す）
// ============================================================
function syncBacklogData() {
  const issues = fetchBacklogIssues_();
  if (issues === null) { Logger.log('取得失敗'); return; }
  const rows = issues.map(issue => mapIssueToRow_(issue));
  writeToSheet_(rows);
  Logger.log(`完了: ${rows.length}件`);
}

function fetchBacklogIssues_() {
  const url = `https://${BACKLOG_SPACE_ID}.backlog.com/api/v2/issues`
    + `?apiKey=${BACKLOG_API_KEY}`
    + `&projectId[]=${BACKLOG_PROJECT_ID}`
    + `&count=${BACKLOG_ISSUE_COUNT}`
    + `&order=updated`
    + `&statusId[]=1&statusId[]=2&statusId[]=3&statusId[]=4`;
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) {
      Logger.log('APIエラー: ' + res.getResponseCode());
      return null;
    }
    return JSON.parse(res.getContentText());
  } catch(e) {
    Logger.log('例外: ' + e.message);
    return null;
  }
}

function getTeamFromAssignee_(assignee) {
  if (!assignee) return DEFAULT_TEAM;
  const name = assignee.name || '';
  if (name.includes('中堅')) return 'ec';
  if (name.includes('皆川')) return 'ec';
  if (name.includes('佐伯')) return 'design';
  return 'promo';
}

function mapIssueToRow_(issue) {
  const title         = issue.summary || '';
  const startRaw      = issue.startDate || issue.created || null;
  const start         = formatDate_(startRaw);
  const end           = issue.dueDate ? formatDate_(issue.dueDate) : '';
  const type          = DEFAULT_TYPE;
  const team          = getTeamFromAssignee_(issue.assignee);
  const memo          = `https://${BACKLOG_SPACE_ID}.backlog.com/view/${issue.issueKey}`;
  const deliveryTime  = (issue.customFields || []).find(f => f.name === '配信時間')?.value || '';
  return [title, start, end, type, team, memo, deliveryTime];
}

function formatDate_(isoString) {
  if (!isoString) return '';
  const d = new Date(isoString);
  if (isNaN(d.getTime())) return '';
  const yyyy = d.getUTCFullYear();
  const mm   = String(d.getUTCMonth() + 1).padStart(2, '0');
  const dd   = String(d.getUTCDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

function writeToSheet_(rows) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_BACKLOG);
  if (!sheet) throw new Error(`シート "${SHEET_BACKLOG}" が見つかりません`);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 7).setValues([['title','start','end','type','team','memo','delivery_time']]);
  if (rows.length > 0) sheet.getRange(2, 1, rows.length, 7).setValues(rows);
}

