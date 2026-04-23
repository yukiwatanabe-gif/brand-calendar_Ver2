// =============================================================================
// Backlog → Google スプレッドシート 同期スクリプト
// スプレッドシート: brand_calender_2026 / Backlog_Data シート
//
// 【セットアップ手順】
//  1. スプレッドシートのメニュー「拡張機能」→「Apps Script」を開く
//  2. このファイルの内容を貼り付けて保存
//  3. 下の CONFIG セクションを自環境に合わせて編集
//  4. 「syncBacklogToSheet」を一度手動実行してAuthを承認
//  5. 「setupTrigger」を実行して毎時自動同期を登録
// =============================================================================

// スタンドアロンGAS用：getActiveSpreadsheet()は使えないので固定ID
const MAIN_SS_ID = '1_IRj1_zfE9JjlckGyT1dXh4trbtBFODfPmJYNRLEKJI';

// ─── CONFIG（ここだけ編集） ───────────────────────────────────────────────────
const CONFIG = {
  // ▼ BacklogのスペースID（例: 'yogibo' → yogibo.backlog.com）
  SPACE_ID: 'yogibo',

  // ▼ Backlog APIキー（GASプロジェクトの「スクリプトプロパティ」に BACKLOG_API_KEY を設定）
  API_KEY: PropertiesService.getScriptProperties().getProperty('BACKLOG_API_KEY') || '',

  // ▼ 対象プロジェクトキー（複数指定可）
  PROJECT_KEYS: ['447191'],

  // ▼ 取得するステータスID（空配列=全ステータス）
  //   1:未対応 2:処理中 3:処理済み 4:完了
  STATUS_IDS: [1, 2, 3],

  // ▼ 書き込み先シート名
  SHEET_NAME: 'Backlog_Data',

  // ▼ 1回のAPIリクエストで取得する件数（最大100）
  FETCH_COUNT: 100,
};

// ─── メディア種別マッピング ────────────────────────────────────────────────────
// Backlogのカテゴリー名・課題種別名 → ダッシュボードのmedia_type
// ※ 部分一致で判定（大文字小文字無視）
const MEDIA_TYPE_MAP = [
  { keywords: ['instagram', 'ig', 'インスタ'],                  type: 'Instagram' },
  { keywords: ['twitter', ' x ', 'ツイート', 'ツイッター', 'xpost'], type: 'X' },
  { keywords: ['facebook', 'fb', 'フェイスブック'],             type: 'Facebook' },
  { keywords: ['blog', 'ブログ'],                                type: 'ブログ' },
  { keywords: ['prtimes', 'pr times', 'プレスリリース', 'press'], type: 'PRTIMES' },
  { keywords: ['コーポレート', 'corporate', 'hp', 'ウェブサイト', 'web'], type: 'コーポレートサイト' },
  { keywords: ['楽天', 'rakuten', 'yahoo', 'amazon', 'ポータル'], type: 'ポータルサイト' },
  { keywords: ['通常企画', '企画'],                              type: '通常企画' },
  { keywords: ['販促', 'プロモ'],                                type: '販促企画' },
  { keywords: ['配信依頼', '依頼'],                              type: '配信依頼' },
];

// ─── チームマッピング ──────────────────────────────────────────────────────────
// Backlogのカテゴリー名・担当者名 → ダッシュボードのteam
const TEAM_MAP = [
  { keywords: ['ec', 'イーコマース', 'eコマース', '通販'],      team: 'ec' },
  { keywords: ['vmd', '店舗', 'ビジュアル', 'visual'],          team: 'vmd' },
  { keywords: ['プロモ', 'promo', 'sns', '広報', 'pr'],          team: 'promo' },
  { keywords: ['デザイン', 'design', 'クリエイティブ'],          team: 'design' },
];

// =============================================================================
// メイン関数
// =============================================================================

/**
 * customFields の中身を確認するデバッグ用関数
 * GASエディタで手動実行して実行ログを確認する
 */
function debugCustomFields() {
  const projectIds = resolveProjectIds(CONFIG.PROJECT_KEYS);
  const query = `apiKey=${encodeURIComponent(CONFIG.API_KEY)}&count=5&offset=0&order=desc&${projectIds.map(id => `projectId[]=${id}`).join('&')}`;
  const url = `https://${CONFIG.SPACE_ID}.backlog.com/api/v2/issues?${query}`;
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const issues = JSON.parse(res.getContentText());
  issues.forEach(issue => {
    Logger.log(`--- ${issue.issueKey}: ${issue.summary}`);
    Logger.log(`customFields: ${JSON.stringify(issue.customFields)}`);
  });
}

/**
 * P列（delivery_time）の書き込み確認用デバッグ
 */
function debugDeliveryTime() {
  // シートのP列を確認
  const ss = SpreadsheetApp.openById(MAIN_SS_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log(`ヘッダー列数: ${headers.length}`);
  Logger.log(`ヘッダー: ${JSON.stringify(headers)}`);
  const pIdx = headers.indexOf('delivery_time');
  Logger.log(`delivery_time 列インデックス: ${pIdx} (${pIdx >= 0 ? String.fromCharCode(65 + pIdx) + '列' : '見つからず'})`);
  if (pIdx >= 0) {
    const vals = sheet.getRange(2, pIdx + 1, 10, 1).getValues().flat();
    Logger.log(`P2:P11の値: ${JSON.stringify(vals)}`);
  }

  // APIから1件取得してdelivery_timeの抽出を確認
  const projectIds = resolveProjectIds(CONFIG.PROJECT_KEYS);
  const query = `apiKey=${encodeURIComponent(CONFIG.API_KEY)}&count=3&offset=0&order=desc&${projectIds.map(id => `projectId[]=${id}`).join('&')}`;
  const url = `https://${CONFIG.SPACE_ID}.backlog.com/api/v2/issues?${query}`;
  const issues = JSON.parse(UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText());
  issues.slice(0, 1).forEach(issue => {
    Logger.log(`${issue.issueKey}: customFields件数=${(issue.customFields||[]).length}`);
    (issue.customFields || []).forEach(f => {
      Logger.log(`  name="${f.name}" (len=${f.name.length}) value=${JSON.stringify(f.value)}`);
    });
    const dt = (issue.customFields || []).find(f => f.name.trim() === '配信時間');
    Logger.log(`配信時間(trim一致): ${JSON.stringify(dt)}`);
  });
}

/**
 * BacklogのすべてのIssueをスプレッドシートに同期する
 */
function syncBacklogToSheet() {
  const ss = SpreadsheetApp.openById(MAIN_SS_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error(`シート "${CONFIG.SHEET_NAME}" が見つかりません`);
  }

  Logger.log('同期開始...');
  const startTime = new Date();

  // プロジェクトIDを解決
  const projectIds = resolveProjectIds();
  if (projectIds.length === 0) {
    Logger.log('対象プロジェクトが見つかりませんでした');
    return;
  }

  // 全Issueを取得
  const issues = fetchAllIssues(projectIds);
  Logger.log(`取得件数: ${issues.length}件`);

  // ヘッダー行
  const HEADERS = [
    'backlog_key',    // A: 課題キー (例: PROJ-123)
    'title',          // B: 件名
    'start',          // C: 開始日 (YYYY-MM-DD)
    'end',            // D: 期限日 (YYYY-MM-DD)
    'media_type',     // E: メディア種別 (X / Instagram / ブログ 等)
    'team',           // F: チーム (ec / promo / vmd / design / all)
    'assignee',       // G: 担当者名
    'status',         // H: ステータス
    'priority',       // I: 優先度
    'milestone',      // J: マイルストーン
    'category',       // K: カテゴリー（元データ）
    'issue_type',     // L: 課題種別（元データ）
    'description',    // M: 詳細（最大500文字）
    'source',         // N: データソース（固定: 'Backlog'）
    'synced_at',      // O: 同期日時
    'delivery_time',  // P: 配信時間（カスタムフィールド）
  ];

  // データ変換
  const rows = issues.map(issue => mapIssueToRow(issue, HEADERS));

  // シートに書き込み
  writeToSheet(sheet, HEADERS, rows);

  const elapsed = ((new Date() - startTime) / 1000).toFixed(1);
  Logger.log(`同期完了: ${rows.length}件 (${elapsed}秒)`);
}

// =============================================================================
// API通信
// =============================================================================

/**
 * プロジェクトキー → プロジェクトIDに変換
 */
function resolveProjectIds() {
  const ids = [];
  for (const key of CONFIG.PROJECT_KEYS) {
    try {
      const url = `https://${CONFIG.SPACE_ID}.backlog.com/api/v2/projects/${key}?apiKey=${CONFIG.API_KEY}`;
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (res.getResponseCode() === 200) {
        const project = JSON.parse(res.getContentText());
        ids.push(project.id);
        Logger.log(`プロジェクト解決: ${key} → ID:${project.id}`);
      } else {
        Logger.log(`プロジェクト取得失敗: ${key} (HTTP ${res.getResponseCode()})`);
      }
    } catch (e) {
      Logger.log(`プロジェクト取得エラー: ${key} - ${e.message}`);
    }
  }
  return ids;
}

/**
 * 全Issueをページネーションで取得
 */
function fetchAllIssues(projectIds) {
  const allIssues = [];
  let offset = 0;

  while (true) {
    // GASはURLSearchParams非対応のため手動でクエリ文字列を構築
    let query = `apiKey=${encodeURIComponent(CONFIG.API_KEY)}&count=${CONFIG.FETCH_COUNT}&offset=${offset}&order=desc`;
    projectIds.forEach(id => { query += `&projectId[]=${id}`; });
    CONFIG.STATUS_IDS.forEach(id => { query += `&statusId[]=${id}`; });

    const url = `https://${CONFIG.SPACE_ID}.backlog.com/api/v2/issues?${query}`;

    let res;
    try {
      res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    } catch (e) {
      Logger.log(`APIリクエストエラー (offset:${offset}): ${e.message}`);
      break;
    }

    if (res.getResponseCode() !== 200) {
      Logger.log(`APIエラー: HTTP ${res.getResponseCode()} - ${res.getContentText()}`);
      break;
    }

    const batch = JSON.parse(res.getContentText());
    if (!batch || batch.length === 0) break;

    allIssues.push(...batch);
    Logger.log(`取得: ${offset + 1}〜${offset + batch.length}件`);

    if (batch.length < CONFIG.FETCH_COUNT) break;
    offset += CONFIG.FETCH_COUNT;

    // レート制限対策
    Utilities.sleep(300);
  }

  return allIssues;
}

// =============================================================================
// データ変換
// =============================================================================

/**
 * BacklogのIssueオブジェクト → スプレッドシート行に変換
 */
function mapIssueToRow(issue, headers) {
  const categoryNames  = (issue.category || []).map(c => c.name).join(', ');
  const issueTypeName  = issue.issueType?.name || '';
  const milestoneName  = (issue.milestone || []).map(m => m.name).join(', ');
  const assigneeName   = issue.assignee?.name || '';
  const statusName     = issue.status?.name || '';
  const priorityName   = issue.priority?.name || '';
  const description    = (issue.description || '').substring(0, 500).replace(/\n/g, ' ');
  const deliveryTime   = (issue.customFields || []).find(f => f.name.trim() === '配信時間')?.value ?? '';

  // メディア種別・チームを推定
  const combinedText = [categoryNames, issueTypeName, issue.summary].join(' ').toLowerCase();
  const mediaType = detectMediaType(combinedText) || issueTypeName || 'タスク';
  const team      = detectTeam(combinedText, assigneeName) || 'all';

  const rowMap = {
    backlog_key:  issue.issueKey || '',
    title:        issue.summary || '',
    start:        formatDate(issue.startDate),
    end:          formatDate(issue.dueDate),
    media_type:   mediaType,
    team:         team,
    assignee:     assigneeName,
    status:       statusName,
    priority:     priorityName,
    milestone:    milestoneName,
    category:     categoryNames,
    issue_type:   issueTypeName,
    description:  description,
    source:         'Backlog',
    synced_at:      new Date().toISOString(),
    delivery_time:  deliveryTime,
  };

  return headers.map(h => rowMap[h] ?? '');
}

/**
 * テキストからメディア種別を推定
 */
function detectMediaType(text) {
  for (const { keywords, type } of MEDIA_TYPE_MAP) {
    if (keywords.some(kw => text.includes(kw))) return type;
  }
  return null;
}

/**
 * テキスト・担当者名からチームを推定
 */
function detectTeam(text, assigneeName) {
  const combined = (text + ' ' + assigneeName).toLowerCase();
  for (const { keywords, team } of TEAM_MAP) {
    if (keywords.some(kw => combined.includes(kw))) return team;
  }
  return null;
}

/**
 * Backlog日時 → YYYY-MM-DD に変換
 */
function formatDate(dateStr) {
  if (!dateStr) return '';
  try {
    return dateStr.substring(0, 10); // "2026-04-17T..." → "2026-04-17"
  } catch (e) {
    return '';
  }
}

// =============================================================================
// スプレッドシート書き込み
// =============================================================================

function writeToSheet(sheet, headers, rows) {
  // 既存データをクリア（ヘッダー除く）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  }

  // ヘッダー書き込み
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // delivery_time列を事前にテキスト形式に固定（9:00が時刻型に自動変換されるのを防ぐ）
  const dtCol = headers.indexOf('delivery_time') + 1;
  if (dtCol > 0) {
    sheet.getRange(1, dtCol, (rows.length || 1) + 1, 1).setNumberFormat('@');
  }

  // データ書き込み
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  // ヘッダー行のスタイル
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1e293b');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  // 列幅の自動調整
  sheet.autoResizeColumns(1, headers.length);

  // 凍結行
  sheet.setFrozenRows(1);
}

// =============================================================================
// Web App エンドポイント（ダッシュボードから直接呼び出す）
// 【デプロイ設定】
//   実行するユーザー: 自分
//   アクセスできるユーザー: 全員（ログイン不要）
// =============================================================================

/**
 * 2026_campaigns_import の全行に CPG-001 形式のIDを振る（初回のみ手動実行）
 * id列がなければ1列目に自動挿入する
 */
// google.script.run 経由でHTMLから直接呼び出される関数
function getCampaignData() {
  return _getData({ parameter: { sheet: 'campaigns' } });
}
function getBacklogData() {
  return _getData({ parameter: { sheet: 'backlog' } });
}

const SPONSOR_SS_ID = '13GiAsAm7OI2xnXaAHwfP6MYkaTm6805gUiQvVL6elco';

function getSponsorSheetData() {
  try {
    const ss = SpreadsheetApp.openById(SPONSOR_SS_ID);
    const sheet = ss.getSheets()[0];
    return _sheetToObjects(sheet, true);
  } catch(err) {
    return { error: err.message };
  }
}

function saveSponsorData(params) {
  try {
    const ss = SpreadsheetApp.openById(SPONSOR_SS_ID);
    const sheet = ss.getSheets()[0];
    const lastRow = sheet.getLastRow();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                        .map(h => String(h).trim());
    const col = h => headers.indexOf(h) + 1;

    if (params.action === 'update' && params._row) {
      const rowNum = parseInt(params._row);
      const fields = ['案件名','案件ステータス','担当者','開始日','終了日','協賛分類','進捗',
                      '確定内容','想定AVE','実績AVE','経費詳細','経費（数値）','協賛金',
                      '企画書URL','メモ','報告対象フラグ','進捗率','イベントURL'];
      fields.forEach(f => { if (col(f) > 0 && params[f] !== undefined) sheet.getRange(rowNum, col(f)).setValue(params[f]); });
      if (col('更新日時') > 0) sheet.getRange(rowNum, col('更新日時')).setValue(new Date().toISOString());
      return { status: 'ok', action: 'updated' };
    }

    // 新規追加
    const newRow = headers.map(h => {
      if (h === '更新日時') return new Date().toISOString();
      if (h === 'ID') return 'SPO-' + String(lastRow).padStart(3, '0');
      return params[h] !== undefined ? params[h] : '';
    });
    sheet.appendRow(newRow);
    return { status: 'ok', action: 'added' };
  } catch(err) {
    return { error: err.message };
  }
}

function assignCampaignIds() {
  const ss = SpreadsheetApp.openById(MAIN_SS_ID);
  const sheet = ss.getSheetByName('2026_campaigns_import');
  if (!sheet) throw new Error('Sheet "2026_campaigns_import" が見つかりません');

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim().toLowerCase());

  if (!headers.includes('id')) {
    sheet.insertColumnBefore(1);
    sheet.getRange(1, 1).setValue('id');
    Logger.log('id列を1列目に挿入しました');
  }

  const newHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim().toLowerCase());
  const idColNum = newHeaders.indexOf('id') + 1;
  const lastRow = sheet.getLastRow();
  let count = 0;

  for (let row = 2; row <= lastRow; row++) {
    const cell = sheet.getRange(row, idColNum);
    if (!cell.getValue()) {
      cell.setValue('CPG-' + String(row - 1).padStart(3, '0'));
      count++;
    }
  }
  Logger.log('ID振り出し完了: ' + count + '件');
}

/**
 * ダッシュボードHTMLからfetchで呼び出されるエンドポイント
 * ?sheet=backlog  → Backlog_Dataシート
 * ?sheet=campaigns → 2026_campaigns_importシート
 */
function doGet(e) {
  const sheetParam = e && e.parameter && e.parameter.sheet;

  // ?sheet=xxx → データAPI（JSON / JSONP）
  if (sheetParam) {
    const data = sheetParam === 'sponsor' ? getSponsorSheetData() : _getData(e);
    const json = JSON.stringify(data);
    const cb   = e.parameter.callback;
    if (cb) {
      return ContentService
        .createTextOutput(`${cb}(${json})`)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ?action=addTask → タスクをスプレッドシートに保存
  if (e && e.parameter && e.parameter.action === 'addTask') {
    const result = _saveManualTask(e.parameter);
    const json = JSON.stringify(result);
    const cb = e.parameter.callback;
    if (cb) {
      return ContentService
        .createTextOutput(`${cb}(${json})`)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ?action=addCampaign / updateCampaign → 企画をスプレッドシートに保存
  if (e && e.parameter && (e.parameter.action === 'addCampaign' || e.parameter.action === 'updateCampaign')) {
    const result = _saveManualCampaign(e.parameter);
    const json = JSON.stringify(result);
    const cb = e.parameter.callback;
    if (cb) {
      return ContentService
        .createTextOutput(`${cb}(${json})`)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ?gasAction=saveSponsor → 協賛案件をスプレッドシートに保存（GitHub Pages対応JSONP）
  if (e && e.parameter && e.parameter.gasAction === 'saveSponsor') {
    const result = saveSponsorData(e.parameter);
    const json = JSON.stringify(result);
    const cb = e.parameter.callback;
    if (cb) {
      return ContentService
        .createTextOutput(`${cb}(${json})`)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }

  // パラメータなし → ダッシュボードHTMLを返す
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('案件ハブ | Yogibo')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function _formatCellDate(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  return String(val || '').trim().replace(/\//g, '-');
}

function _saveManualCampaign(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SS_ID);
    const sheet = ss.getSheetByName('2026_campaigns_import');
    if (!sheet) return { error: 'Sheet not found: 2026_campaigns_import' };

    const lastRow = sheet.getLastRow();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                        .map(h => String(h).trim().toLowerCase());
    const col = h => headers.indexOf(h) + 1;
    const idCol = col('id');

    if (params.action === 'updateCampaign' && params.id && idCol > 0 && lastRow > 1) {
      // IDで行を特定して更新
      const idValues = sheet.getRange(2, idCol, lastRow - 1, 1).getValues().flat();
      const rowIdx = idValues.findIndex(v => String(v) === String(params.id));
      if (rowIdx >= 0) {
        const rowNum = rowIdx + 2;
        const fields = ['title', 'start', 'end', 'type', 'team', 'memo'];
        fields.forEach(f => { if (col(f) > 0) sheet.getRange(rowNum, col(f)).setValue(params[f] || ''); });
        return { status: 'ok', action: 'updated', id: params.id };
      }
    }

    // 新規追加：次のIDを自動採番
    let newId = '';
    if (idCol > 0) {
      const existingIds = lastRow > 1
        ? sheet.getRange(2, idCol, lastRow - 1, 1).getValues().flat().map(v => String(v))
        : [];
      const nums = existingIds.map(id => parseInt(id.replace('CPG-', '')) || 0).filter(n => n > 0);
      const nextNum = nums.length > 0 ? Math.max(...nums) + 1 : 1;
      newId = 'CPG-' + String(nextNum).padStart(3, '0');
    }
    const newRow = headers.map(h => h === 'id' ? newId : (params[h] || ''));
    sheet.appendRow(newRow);
    return { status: 'ok', action: 'added', id: newId };
  } catch(err) {
    return { error: err.message };
  }
}

function removeDuplicateCampaigns() {
  const ss = SpreadsheetApp.openById(MAIN_SS_ID);
  const sheet = ss.getSheetByName('2026_campaigns_import');
  if (!sheet) return;
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim().toLowerCase());
  const titleCol = headers.indexOf('title');
  const startCol = headers.indexOf('start');
  const seen = new Set();
  const rowsToDelete = [];
  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][titleCol] || '').trim() + '|' + _formatCellDate(values[i][startCol]);
    if (seen.has(key)) {
      rowsToDelete.push(i + 1);
    } else {
      seen.add(key);
    }
  }
  // 下から削除
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }
  Logger.log('削除した重複行数: ' + rowsToDelete.length);
}

function _saveManualTask(params) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SS_ID);
    let sheet = ss.getSheetByName('Tasks_Manual');
    if (!sheet) {
      sheet = ss.insertSheet('Tasks_Manual');
      sheet.getRange(1, 1, 1, 6).setValues([['title','start','end','type','team','memo']]);
    }
    sheet.appendRow([
      params.title || '',
      params.end   || '',
      params.end   || '',
      params.type  || 'deadline',
      params.team  || 'all',
      params.memo  || '',
    ]);
    return { status: 'ok' };
  } catch(err) {
    return { error: err.message };
  }
}

function _sheetToObjects(sheet, includeRowNum) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h).trim());
  return values.slice(1)
    .map((row, idx) => {
      const obj = {};
      if (includeRowNum) obj._row = idx + 2;
      headers.forEach((h, i) => {
        const v = row[i];
        if (v instanceof Date) {
          obj[h] = v.getFullYear() <= 1900
            ? Utilities.formatDate(v, 'Asia/Tokyo', 'HH:mm')
            : Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
        } else {
          obj[h] = String(v === null || v === undefined ? '' : v);
        }
      });
      return obj;
    })
    .filter(obj => headers.some(h => obj[h] !== ''));
}

function _getData(e) {
  try {
    const sheetParam = (e && e.parameter && e.parameter.sheet) || 'backlog';
    const ss = SpreadsheetApp.openById(MAIN_SS_ID);

    if (sheetParam === 'campaigns') {
      const mainSheet = ss.getSheetByName('2026_campaigns_import');
      if (!mainSheet) return { error: 'Sheet not found: 2026_campaigns_import' };
      return _sheetToObjects(mainSheet, true); // _row付きで返す
    }

    const sheetName = CONFIG.SHEET_NAME;
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { error: `Sheet not found: ${sheetName}` };
    return _sheetToObjects(sheet);

  } catch (err) {
    return { error: err.message };
  }
}

// =============================================================================
// トリガー設定
// =============================================================================

/**
 * 毎時自動同期トリガーを設定する（初回のみ手動実行）
 */
function setupTrigger() {
  // 既存トリガーを削除
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncBacklogToSheet') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 1時間おきに実行
  ScriptApp.newTrigger('syncBacklogToSheet')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('トリガー設定完了: syncBacklogToSheet を毎時実行します');
}

/**
 * トリガーを削除する
 */
function removeTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncBacklogToSheet') {
      ScriptApp.deleteTrigger(t);
    }
  });
  Logger.log('トリガーを削除しました');
}

// =============================================================================
// 接続テスト（設定確認用）
// =============================================================================

/**
 * API接続とプロジェクト取得をテストする
 * 最初にこれを実行して設定を確認してください
 */
function testConnection() {
  Logger.log('=== 接続テスト開始 ===');
  Logger.log(`スペース: ${CONFIG.SPACE_ID}.backlog.com`);

  for (const key of CONFIG.PROJECT_KEYS) {
    const url = `https://${CONFIG.SPACE_ID}.backlog.com/api/v2/projects/${key}?apiKey=${CONFIG.API_KEY}`;
    try {
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const code = res.getResponseCode();
      if (code === 200) {
        const proj = JSON.parse(res.getContentText());
        Logger.log(`✅ プロジェクト接続OK: ${proj.name} (ID:${proj.id})`);
      } else if (code === 401) {
        Logger.log(`❌ 認証エラー: APIキーを確認してください (${key})`);
      } else if (code === 404) {
        Logger.log(`❌ プロジェクト未発見: "${key}" のキーを確認してください`);
      } else {
        Logger.log(`❌ エラー: HTTP ${code} - ${res.getContentText()}`);
      }
    } catch (e) {
      Logger.log(`❌ 通信エラー: ${e.message}`);
    }
  }

  Logger.log('=== 接続テスト完了 ===');
}
