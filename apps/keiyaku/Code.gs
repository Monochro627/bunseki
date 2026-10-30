/**
 * 契約進捗アプリ（Monochro 賃貸）
 *
 * データの本体は Google スプレッドシート「契約進捗DB」。
 * 売上管理表には、お客様名のセルにリンクを付けるだけで、それ以外は書き込まない。
 */

// 設定（CONFIG）は「設定.gs」にあります。このファイルを貼り替えても設定は消えません。

const SHEET_DEFS = {
  items:   { name: '項目マスタ', head: ['ID', '契約方法', '段階', '項目', '補足', '種類', '有効'] },
  deals:   { name: '案件', head: ['案件ID', '作成日時', '担当', 'お客様名', '月タブ', '行', '物件', '管理会社', 'TEL', '契約方法',
                                  '重説予定日', '契約書到着予定日', '決済予定日', '入居日', '付帯', '状態', '引き継ぎ先',
                                  'キャンセル理由', 'キャンセル詳細', 'キャンセル日時', 'キャンセル者', '最終活動',
                                  '重説完了日', '宅建士', '鍵渡し予定日時', '完了日時', '顧客No'] },
  checks:  { name: 'チェック', head: ['案件ID', '項目ID', '状態', '日時', '記録者'] },
  notes:   { name: 'メモ', head: ['日時', '案件ID', '記録者', '内容'] },
  hiki:    { name: '引き継ぎ', head: ['案件ID', 'キー', '値', '日時', '記録者'] },
  log:     { name: '履歴', head: ['日時', '案件ID', '操作', '内容', '記録者'] },
  members: { name: 'メンバー', head: ['売上表の名前', 'メール'] },
  jusetsu: { name: '重説集計', head: ['月', '宅建士', '件数', 'お客様（重説完了日）'] },
};

const MODE_LIST = { '対面': '対面・郵送', '郵送': '対面・郵送', '電子': '電子' };
const MAIL_ONLY_PHASE = '郵送契約のみ';
const DATE_FIELDS = ['重説予定日', '重説完了日', '契約書到着予定日', '決済予定日', '入居日'];
const DATETIME_FIELDS = ['鍵渡し予定日時'];
const EDITABLE_FIELDS = ['契約方法', '物件', '管理会社', 'TEL', '付帯', '宅建士'].concat(DATE_FIELDS, DATETIME_FIELDS);
// 契約時の段階の先頭に入れる、枠の色を変えて目立たせる項目
const KEY_ITEMS = ['契約締結', '入金', '送金'];

// ===== 初期設定（最初に1回だけ実行） =====
function setup() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('DB_ID')) {
    throw new Error('すでに契約進捗DBがあります: ' + SpreadsheetApp.openById(props.getProperty('DB_ID')).getUrl());
  }
  const ss = SpreadsheetApp.create('契約進捗DB');
  Object.keys(SHEET_DEFS).forEach((key, i) => {
    const def = SHEET_DEFS[key];
    const sh = i === 0 ? ss.getSheets()[0].setName(def.name) : ss.insertSheet(def.name);
    sh.getRange(1, 1, 1, def.head.length).setValues([def.head]).setFontWeight('bold').setBackground('#eeeeee');
    sh.setFrozenRows(1);
  });
  const rows = seedItems_();
  const itemSheet = ss.getSheetByName(SHEET_DEFS.items.name);
  itemSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  itemSheet.getRange(2, 7, rows.length, 1).insertCheckboxes();
  itemSheet.getRange('F1').setNote('空=通常 / 続き=直前の項目の続き / 注意=チェックしない注意書き / 期日のみ=引き継ぎの期日欄');
  itemSheet.getRange('C1').setNote('「郵送契約のみ」の段階は、契約方法が郵送のときだけ表示されます');
  const members = ['小嶋勇輝', '諸本雅広', '佐光琴音', '澤井颯晟', '松井拓登', '寺田喜心', '山田園佳', '宮川大輝'].map(n => [n, '']);
  ss.getSheetByName(SHEET_DEFS.members.name).getRange(2, 1, members.length, 2).setValues(members);
  props.setProperty('DB_ID', ss.getId());
  ensureSchema_();
  Logger.log('契約進捗DBを作成しました: ' + ss.getUrl());
}

// トリガーを入れる（WEB_APP_URL を設定した後に1回実行）
function installTriggers() {
  if (!CONFIG.WEB_APP_URL) throw new Error('先に CONFIG.WEB_APP_URL を設定してください');
  ensureSchema_();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  // いま売上表に入っているお客様は、チェックを一度押すまでアラートを出さない
  syncSales_(true);
  repaintAll();
  ScriptApp.newTrigger('onSalesEdit').forSpreadsheet(CONFIG.SALES_ID).onEdit().create();
  ScriptApp.newTrigger('syncSales').timeBased().everyMinutes(10).create();
  ScriptApp.newTrigger('sendStallAlerts').timeBased().atHour(8).everyDays(1).inTimezone(CONFIG.TZ).create();
  ScriptApp.newTrigger('sendMonthlyJusetsu').timeBased().onMonthDay(1).atHour(9).inTimezone(CONFIG.TZ).create();
  refreshJusetsu_();
  Logger.log('トリガーを設定しました');
}

// DBの見出しを最新にし、足りないシートと「契約締結・入金・送金」の項目を追加する（何度実行しても同じ結果）
function ensureSchema_() {
  const ss = db_();
  Object.keys(SHEET_DEFS).forEach(key => {
    const def = SHEET_DEFS[key];
    let sh = ss.getSheetByName(def.name);
    if (!sh) { sh = ss.insertSheet(def.name); sh.setFrozenRows(1); }
    sh.getRange(1, 1, 1, def.head.length).setValues([def.head]).setFontWeight('bold').setBackground('#eeeeee');
  });
  const itemSheet = ss.getSheetByName(SHEET_DEFS.items.name);
  [['対面・郵送', 'TK'], ['電子', 'EK']].forEach(([list, prefix]) => {
    const items = readAll_('items');
    if (items.some(i => i['契約方法'] === list && i['種類'] === '重要')) return;
    const first = items.find(i => i['契約方法'] === list && i['段階'] === '契約時');
    if (!first) return;
    itemSheet.insertRowsBefore(first._row, KEY_ITEMS.length);
    delete ROWS_CACHE_.items;
    itemSheet.getRange(first._row, 1, KEY_ITEMS.length, 7)
      .setValues(KEY_ITEMS.map((t, i) => [prefix + (i + 1), list, '契約時', t, '', '重要', true]));
    itemSheet.getRange(first._row, 7, KEY_ITEMS.length, 1).insertCheckboxes();
  });
}

// テストデータを消す（案件・チェック・メモ・引き継ぎ・履歴。項目マスタとメンバーは残す）
function resetData() {
  ['deals', 'checks', 'notes', 'hiki', 'log'].forEach(key => {
    const sh = sheet_(key);
    if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
    delete ROWS_CACHE_[key];
  });
  Logger.log('テストデータを消しました');
}

// ===== 売上表との連携 =====
function onSalesEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (!CONFIG.TAB_PATTERN.test(sh.getName().trim())) return;
  const cols = salesCols_(sh);
  if (!cols) return;
  const c1 = e.range.getColumn(), c2 = e.range.getLastColumn();
  const hit = [cols.name, cols.kind, cols.sale].some(c => c && c >= c1 && c <= c2);
  if (!hit) return;
  if (scanTabs_(sh.getParent()).every(t => t.getSheetId() !== sh.getSheetId())) return;
  // 編集された行だけを見る。ほかの処理中なら待たずにやめる（10分ごとの確認で拾う）
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return;
  try {
    scanTab_(sh, readAll_('deals'), false, e.range.getRow(), e.range.getLastRow());
  } finally {
    lock.releaseLock();
  }
}

function syncSales() { syncSales_(false); }

function syncSales_(quiet) {
  const lock = LockService.getScriptLock();
  // 定期の確認は、ほかの処理中なら待たずに次の回へ回す
  if (!lock.tryLock(quiet ? 30000 : 1000)) return;
  try {
    const deals = readAll_('deals');
    scanTabs_(SpreadsheetApp.openById(CONFIG.SALES_ID)).forEach(sh => scanTab_(sh, deals, quiet));
  } finally {
    lock.releaseLock();
  }
}

// 監視するタブ（新しい順に TAB_SCAN_COUNT 枚。FIRST_TAB より古いタブは含めない）
function scanTabs_(ss) {
  const all = ss.getSheets().filter(s => CONFIG.TAB_PATTERN.test(s.getName().trim()));
  const first = all.findIndex(s => s.getName().trim() === CONFIG.FIRST_TAB);
  return (first >= 0 ? all.slice(0, first + 1) : all).slice(0, CONFIG.TAB_SCAN_COUNT);
}

function salesCols_(sh) {
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v).replace(/\s/g, ''));
  const name = head.indexOf('お客様名') + 1, kind = head.indexOf('案件種別') + 1, move = head.indexOf('入居日') + 1;
  const sale = head.indexOf('売買') + 1; // チェックが入っている行は売買案件なので対象外
  return name && kind ? { name, kind, move, sale } : null;
}

// fromRow〜toRow の行だけを調べる（省略すると全行）。ブロック名を決めるため A列だけは上から読む
function scanTab_(sh, deals, quiet, fromRow, toRow) {
  const cols = salesCols_(sh);
  const last = sh.getLastRow();
  if (!cols || last < 2) return;
  const r1 = Math.max(2, fromRow || 2), r2 = Math.min(last, toRow || last);
  if (r2 < r1) return;
  const colA = sh.getRange(2, 1, r2 - 1, 1).getValues();
  const n = r2 - r1 + 1, off = r1 - 2;
  const kinds = sh.getRange(r1, cols.kind, n, 1).getDisplayValues();
  const nameRange = sh.getRange(r1, cols.name, n, 1);
  const names = nameRange.getDisplayValues();
  const rich = nameRange.getRichTextValues();
  const moves = cols.move ? sh.getRange(r1, cols.move, n, 1).getValues() : null;
  const sales = cols.sale ? sh.getRange(r1, cols.sale, n, 1).getValues() : null;
  const bgs = nameRange.getBackgrounds();
  const tab = sh.getName().trim();
  let block = '';
  for (let k = 0; k < r2 - 1; k++) {
    const a = String(colA[k][0]).replace(/\s/g, '');
    if (a && a !== '合計') block = a;
    if (a === '合計' || k < off) continue;
    const i = k - off;
    const name = String(names[i][0]).trim();
    const kind = String(kinds[i][0]).trim();
    if (!name || !kind) continue;
    if (CONFIG.TARGET_STAFF.length && CONFIG.TARGET_STAFF.indexOf(block) < 0) continue;
    if (CONFIG.EXCLUDE_BLOCKS.some(x => block.indexOf(x) === 0)) continue;
    const link = rich[i][0] ? rich[i][0].getLinkUrl() : null;
    if (sales && sales[i][0] === true) {
      // 売買案件：付いていたリンクと色を外し、案件は「対象外」にする
      const own = link && dealIdFromUrl_(link);
      if (own) {
        const cell = sh.getRange(r1 + i, cols.name);
        cell.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(names[i][0]).build());
        cell.setBackground(null);
        const d = deals.find(x => x['案件ID'] === own);
        if (d && d['状態'] !== '対象外') { update_('deals', d, { '状態': '対象外' }); log_(own, '対象外', '売買にチェックが入ったため', 'システム'); }
      }
      continue;
    }
    if (link && dealIdFromUrl_(link)) continue; // すでにリンク済み
    if (link) continue; // 別のリンクが付いているセルには触らない

    // ① 名前のセルに顧客No（M3188）が書いてあれば、その顧客の案件（見込みボードの「成約」で作ったもの）につなぐ
    const noHit = name.match(CUSTOMER_NO_RE);
    let deal = noHit ? deals.find(d => d['顧客No'] === noHit[0] && d['状態'] !== 'キャンセル') : null;
    if (deal && (deal['月タブ'] !== tab || Number(deal['行']) !== r1 + i)) {
      update_('deals', deal, { '月タブ': tab, '行': r1 + i, '担当': block });
      log_(deal['案件ID'], '売上表とつなぐ', tab + ' ' + (r1 + i) + '行目 ' + name + '（顧客No）', 'システム');
    }
    // ② 同じタブ・同じ名前の案件があればつなぎ直す（リンクが消えた場合の保険）
    if (!deal) deal = deals.find(d => d['月タブ'] === tab && d['お客様名'] === name && d['状態'] !== 'キャンセル');
    if (deal && deal['状態'] === '対象外') {
      // 売買のチェックが外れたので賃貸として戻す
      update_('deals', deal, { '状態': '進行中' });
      log_(deal['案件ID'], '対象に戻す', '売買のチェックが外れたため', 'システム');
    }
    // ③ 見込みボードの「成約」で先に作られた案件があれば、それにつなぐ（同じ担当・ゆるく同じ名前）
    if (!deal) deal = linkBoardDeal_(deals, name, block, tab, r1 + i);
    const cell = sh.getRange(r1 + i, cols.name);
    if (!deal) {
      // 案件は見込みボードの「成約」でだけ作る。つながらない行は黄色にして知らせる（成約を押すか、名前の後ろに顧客No を書くとつながる）
      if (String(bgs[i][0]).toLowerCase() !== NO_DEAL_COLOR) {
        cell.setBackground(NO_DEAL_COLOR);
        cell.setNote('契約進捗とまだつながっていません。見込みボードでこのお客様の「成約」を押すか、名前の後ろに顧客No（例：田坂様 M3188）を書くとつながります。');
      }
      continue;
    }
    // 名前の後ろに顧客No を付けて、売上表でも番号が分かるようにする（顧客No だけ書いたときは、お客様名に置き換える）
    let text = String(names[i][0]).trim();
    const no = deal['顧客No'];
    if (noHit && text.replace(CUSTOMER_NO_RE, '').trim() === '') text = salesName_(deal['お客様名']);
    if (no && text.indexOf(no) < 0) text = text.replace(CUSTOMER_NO_RE, '').trim() + ' ' + no;
    const url = CONFIG.WEB_APP_URL + '?id=' + deal['案件ID'];
    cell.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(text).setLinkUrl(url).build());
    if (String(bgs[i][0]).toLowerCase() === NO_DEAL_COLOR) cell.clearNote();
    paintCell_(cell, deal['状態'] === 'キャンセル' ? '#ffffff' : deal['状態'] === '完了' ? (blockColor_(sh, r1 + i) || '#d9ead3') : PENDING_COLOR);
  }
}

// 見込みボードで作った案件（月タブが空）のうち、名前（空白・様・さん・カッコ内のフリガナを除く）が同じものを探してつなぐ
// 同じ担当（先頭2文字）を優先。送客の案件は売上表では送客元のブロックに入ることがあるので、担当が違っても同じ名前が1件だけならつなぐ
function linkBoardDeal_(deals, name, block, tab, row) {
  const nm = boardName_(name), st = String(block).replace(/[\s　]/g, '').slice(0, 2);
  const open = deals.filter(d => !d['月タブ'] && d['状態'] !== 'キャンセル' && d['状態'] !== '対象外');
  const same = open.filter(d => boardName_(d['お客様名']) === nm);
  let deal = same.find(d => String(d['担当']).replace(/[\s　]/g, '').slice(0, 2) === st) || (same.length === 1 ? same[0] : null);
  // 見込みボードの名前は「谷元さんご紹介　田坂様（タニモトサン）」のように前に紹介メモが付くことがあるので、
  // 同じ担当で、片方の名前がもう片方に含まれる案件が1件だけならつなぐ
  if (!deal && nm.length >= 2) {
    const part = open.filter(d => String(d['担当']).replace(/[\s　]/g, '').slice(0, 2) === st)
      .filter(d => { const b = boardName_(d['お客様名']); return b.length >= 2 && (b.indexOf(nm) >= 0 || nm.indexOf(b) >= 0); });
    if (part.length === 1) deal = part[0];
  }
  if (!deal) return null;
  update_('deals', deal, { '月タブ': tab, '行': row, '担当': block });
  log_(deal['案件ID'], '売上表とつなぐ', tab + ' ' + row + '行目 ' + name, 'システム');
  return deal;
}
function boardName_(s) {
  return String(s || '').replace(CUSTOMER_NO_RE, '').replace(/[（(][^）)]*[）)]/g, '').replace(/[\s　]/g, '').replace(/(様|さま|さん)$/, '');
}
// 顧客管理の顧客No（見込みボードが振る M0001 の形）。売上表の名前のセルにこれがあれば、その顧客の案件につなぐ
const CUSTOMER_NO_RE = /M\d{4,}/;
// 売上表で、契約進捗とまだつながっていない名前のセルの色（見込みボードで成約を押すとつながる）
const NO_DEAL_COLOR = '#fff2cc';
// 売上表に書くお客様名：見込みボードの名前からフリガナ（カッコ）を外し、最後に「様」を付ける
function salesName_(s) {
  const n = String(s || '').replace(/\s*[（(][^（()）]*[)）]\s*$/, '').trim();
  return /(様|さま)$/.test(n) ? n : n + '様';
}

// 案件の売上表上の位置を探す（行が動いていてもリンクで探し直す）
function locateSales_(deal) {
  const sh = SpreadsheetApp.openById(CONFIG.SALES_ID).getSheetByName(deal['月タブ']);
  if (!sh) return null;
  const cols = salesCols_(sh);
  if (!cols) return null;
  let row = Number(deal['行']);
  const own = r => {
    const rt = sh.getRange(r, cols.name).getRichTextValue();
    return rt && dealIdFromUrl_(rt.getLinkUrl() || '') === deal['案件ID'];
  };
  if (!row || !own(row)) {
    const rich = sh.getRange(2, cols.name, sh.getLastRow() - 1, 1).getRichTextValues();
    const idx = rich.findIndex(r => r[0] && dealIdFromUrl_(r[0].getLinkUrl() || '') === deal['案件ID']);
    if (idx < 0) return null;
    row = idx + 2;
    update_('deals', deal, { '行': row });
  }
  return { sh, row, cols };
}

function readSalesMoveIn_(deal) {
  try {
    const loc = locateSales_(deal);
    if (!loc || !loc.cols.move) return null;
    const v = loc.sh.getRange(loc.row, loc.cols.move).getValue();
    return v instanceof Date ? v : null;
  } catch (err) {
    return null;
  }
}

// 売上表のお客様名のセルを塗る：未完了＝ラベンダー、完了＝A列の担当ブロックの色、キャンセル＝白
const PENDING_COLOR = '#e4d7f5';
function paintSalesName_(deal) {
  try {
    const loc = locateSales_(deal);
    if (!loc) return;
    let color = PENDING_COLOR;
    if (deal['状態'] === 'キャンセル') color = '#ffffff';
    if (deal['状態'] === '完了') color = blockColor_(loc.sh, loc.row) || '#d9ead3';
    paintCell_(loc.sh.getRange(loc.row, loc.cols.name), color);
  } catch (err) {
    log_(deal['案件ID'], '色の変更に失敗', String(err), 'システム');
  }
}

// リンクは残したまま、文字を黒・下線にして背景を塗る（濃い背景でも読めるように）
function paintCell_(cell, color) {
  const rt = cell.getRichTextValue();
  if (rt && rt.getLinkUrl()) {
    cell.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(rt.getText()).setLinkUrl(rt.getLinkUrl())
      .setTextStyle(SpreadsheetApp.newTextStyle().setForegroundColor('#000000').setUnderline(true).setBold(false).build()).build());
  }
  cell.setBackground(color);
}

// その行が属するブロック（A列の結合セル）の背景色
function blockColor_(sh, row) {
  const colA = sh.getRange(1, 1, row, 1).getValues();
  for (let r = row; r >= 2; r--) {
    const a = String(colA[r - 1][0]).replace(/\s/g, '');
    if (a && a !== '合計') {
      const c = sh.getRange(r, 1).getBackground();
      return c && c !== '#ffffff' ? c : '';
    }
  }
  return '';
}

// いまある案件の色をまとめて塗り直す（installTriggers のときに実行）
function repaintAll() {
  readAll_('deals').forEach(d => { if (d['状態'] !== '対象外') paintSalesName_(d); });
}

// 状態を更新し、変わったときだけ売上表の色も変える
function setStatus_(deal, patch) {
  const before = deal['状態'];
  update_('deals', deal, patch);
  if (deal['状態'] !== before) paintSalesName_(deal);
}

function dealIdFromUrl_(url) {
  if (!url || !CONFIG.WEB_APP_URL) return '';
  if (url.indexOf(CONFIG.WEB_APP_URL) !== 0) return '';
  const m = url.match(/[?&]id=([A-Za-z0-9-]+)/);
  return m ? m[1] : '';
}

// ===== ウェブアプリ =====
function doGet(e) {
  const p = (e && e.parameter) || {};
  const t = HtmlService.createTemplateFromFile('App');
  t.dealId = /^[A-Za-z0-9-]{1,40}$/.test(p.id || '') ? p.id : '';
  t.printMode = p.print === '1';
  t.appUrl = CONFIG.WEB_APP_URL;
  return t.evaluate()
    .setTitle('契約進捗')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // Notion などのページに埋め込めるようにする
}

function apiList() {
  const items = readAll_('items');
  const checks = groupBy_(readAll_('checks'), '案件ID');
  const me = me_();
  // 一覧のカードに出す、いちばん新しいメモ（下の行ほど新しい）
  const memos = {};
  readAll_('notes').forEach(n => { memos[n['案件ID']] = { text: String(n['内容']), at: fmtDateTime_(n['日時']), by: n['記録者'] }; });
  return {
    me,
    deals: readAll_('deals').map(d => {
      const p = progress_(d, items, checks[d['案件ID']] || []);
      return { id: d['案件ID'], name: d['お客様名'], no: d['顧客No'], staff: d['担当'], tab: d['月タブ'], mode: d['契約方法'],
               status: d['状態'], done: p.done, total: p.total, phase: p.phase, idleHours: idleHours_(d),
               moveIn: fmtDate_(d['入居日']), phases: p.phases, phaseIdx: p.phaseIdx, memo: memos[d['案件ID']] || null };
    }).reverse(),
    ranking: monthlyRanking_(),
  };
}

// 今月「完了にする」を押した件数を担当ごとに（多い順）
function monthlyRanking_() {
  const month = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM');
  const count = {};
  readAll_('deals').forEach(d => {
    if (d['状態'] !== '完了' || !(d['完了日時'] instanceof Date)) return;
    if (Utilities.formatDate(d['完了日時'], CONFIG.TZ, 'yyyy-MM') !== month) return;
    count[d['担当']] = (count[d['担当']] || 0) + 1;
  });
  return { month: Number(month.slice(5)), list: Object.keys(count).map(k => ({ staff: k, count: count[k] })).sort((a, b) => b.count - a.count) };
}

// 売上表の入居日を読み直す（売上表は大きく開くのに時間がかかるので、画面を出したあとに裏で呼ぶ）
function apiRefreshMoveIn(id) {
  const deal = findDeal_(id);
  const salesMove = readSalesMoveIn_(deal);
  if (salesMove && fmtDate_(salesMove) !== fmtDate_(deal['入居日'])) update_('deals', deal, { '入居日': salesMove });
  return { moveIn: fmtDate_(deal['入居日']), fromSales: !!salesMove };
}

// 段階の「リセット」：指定した項目のチェック（✓も✕も）を全部消して未に戻す
function apiResetMany(id, itemIds) {
  return withLock_(() => {
    const deal = findDeal_(id);
    const me = me_();
    const targets = readAll_('checks').filter(c => c['案件ID'] === id && (itemIds || []).indexOf(c['項目ID']) >= 0)
      .sort((a, b) => b._row - a._row); // 下の行から消すと行番号がずれない
    targets.forEach(c => sheet_('checks').deleteRow(c._row));
    delete ROWS_CACHE_.checks;
    if (targets.length) log_(id, 'リセット', targets.map(c => c['項目ID']).join(','), me.name);
    const patch = { '最終活動': new Date() };
    if (deal['状態'] === '完了' && targets.length) { patch['状態'] = '進行中'; patch['完了日時'] = ''; }
    setStatus_(deal, patch);
    return { removed: targets.map(c => c['項目ID']), status: deal['状態'] };
  });
}

function apiGet(id) {
  const deal = findDeal_(id);
  return {
    me: me_(),
    members: readAll_('members').map(m => m['売上表の名前']).filter(String),
    items: readAll_('items').filter(i => i['有効'] !== false).map(i => ({
      id: i['ID'], list: i['契約方法'], phase: i['段階'], text: i['項目'], note: i['補足'], kind: i['種類'] })),
    deal: dealToClient_(deal),
    moveInFromSales: false,
    checks: readAll_('checks').filter(c => c['案件ID'] === id).map(c => ({
      item: c['項目ID'], v: c['状態'], at: fmtDateTime_(c['日時']), by: c['記録者'] })),
    notes: readAll_('notes').filter(n => n['案件ID'] === id).map(n => ({
      at: fmtDateTime_(n['日時']), by: n['記録者'], text: n['内容'] })).reverse(),
    hiki: readAll_('hiki').filter(h => h['案件ID'] === id).reduce((o, h) => { o[h['キー']] = h['値']; return o; }, {}),
  };
}

// next: 'done' / 'na' / ''（未に戻す）
function apiCheck(id, itemId, next) {
  return withLock_(() => {
    const deal = findDeal_(id);
    const me = me_();
    const now = new Date();
    const row = readAll_('checks').find(c => c['案件ID'] === id && c['項目ID'] === itemId);
    if (!next) {
      if (row) deleteRow_('checks', row);
    } else if (row) {
      update_('checks', row, { '状態': next, '日時': now, '記録者': me.name });
    } else {
      append_('checks', { '案件ID': id, '項目ID': itemId, '状態': next, '日時': now, '記録者': me.name });
    }
    log_(id, 'チェック', itemId + ' → ' + (next || '未'), me.name);
    const p = progress_(deal, readAll_('items'), readAll_('checks').filter(c => c['案件ID'] === id));
    const patch = { '最終活動': now };
    // 完了はボタンで決める。完了後に未チェックが出たら進行中に戻す
    if (deal['状態'] === '完了' && !(p.total && p.done === p.total)) { patch['状態'] = '進行中'; patch['完了日時'] = ''; }
    setStatus_(deal, patch);
    return { at: fmtDateTime_(now), by: me.name, status: deal['状態'] };
  });
}

// 「完了にする」ボタン：全部埋まっているときだけ完了にする
function apiComplete(id) {
  return withLock_(() => {
    const deal = findDeal_(id);
    const p = progress_(deal, readAll_('items'), readAll_('checks').filter(c => c['案件ID'] === id));
    if (!(p.total && p.done === p.total)) throw new Error('まだチェックが残っています（' + p.done + '/' + p.total + '）');
    const me = me_(), now = new Date();
    setStatus_(deal, { '状態': '完了', '完了日時': now, '最終活動': now });
    log_(id, '完了', '', me.name);
    return dealToClient_(deal);
  });
}

// 段階の「残りを全部✓」：未チェックの項目だけをまとめて完了にする（✕はそのまま）
function apiCheckMany(id, itemIds) {
  return withLock_(() => {
    const deal = findDeal_(id);
    const me = me_();
    const now = new Date();
    const existing = readAll_('checks').filter(c => c['案件ID'] === id);
    const added = [];
    (itemIds || []).forEach(itemId => {
      if (existing.some(c => c['項目ID'] === itemId)) return;
      append_('checks', { '案件ID': id, '項目ID': itemId, '状態': 'done', '日時': now, '記録者': me.name });
      added.push(itemId);
    });
    if (added.length) log_(id, 'まとめてチェック', added.join(','), me.name);
    const p = progress_(deal, readAll_('items'), readAll_('checks').filter(c => c['案件ID'] === id));
    const patch = { '最終活動': now };
    // 完了はボタンで決める。完了後に未チェックが出たら進行中に戻す
    if (deal['状態'] === '完了' && !(p.total && p.done === p.total)) { patch['状態'] = '進行中'; patch['完了日時'] = ''; }
    setStatus_(deal, patch);
    return { added, at: fmtDateTime_(now), by: me.name, status: deal['状態'] };
  });
}

function apiField(id, field, value) {
  if (EDITABLE_FIELDS.indexOf(field) < 0) throw new Error('この項目は変更できません: ' + field);
  return withLock_(() => {
    const deal = findDeal_(id);
    let v = value;
    if (DATE_FIELDS.indexOf(field) >= 0) v = parseDate_(value);
    if (DATETIME_FIELDS.indexOf(field) >= 0) v = parseDateTime_(value);
    if (field === '契約方法' && !MODE_LIST[v]) throw new Error('契約方法が正しくありません');
    const patch = {}; patch[field] = v;
    if (field === '契約方法') {
      // 契約方法が変わるとチェック項目が変わるので、完了状態を計算し直す
      const tmp = Object.assign({}, deal, patch);
      const p = progress_(tmp, readAll_('items'), readAll_('checks').filter(c => c['案件ID'] === id));
      // 完了はボタンで決める。完了後に未チェックが出たら進行中に戻す
    if (deal['状態'] === '完了' && !(p.total && p.done === p.total)) { patch['状態'] = '進行中'; patch['完了日時'] = ''; }
    }
    setStatus_(deal, patch);
    log_(id, '変更', field + ' = ' + value, me_().name);
    if (field === '重説完了日' || field === '宅建士') refreshJusetsu_();
    return true;
  });
}

function apiNote(id, text) {
  text = String(text || '').trim();
  if (!text) throw new Error('メモが空です');
  return withLock_(() => {
    const deal = findDeal_(id);
    const me = me_(), now = new Date();
    append_('notes', { '日時': now, '案件ID': id, '記録者': me.name, '内容': text });
    update_('deals', deal, { '最終活動': now });
    return { at: fmtDateTime_(now), by: me.name, text };
  });
}

function apiHiki(id, key, value) {
  return withLock_(() => {
    findDeal_(id);
    const me = me_(), now = new Date();
    const row = readAll_('hiki').find(h => h['案件ID'] === id && h['キー'] === key);
    if (row) update_('hiki', row, { '値': value, '日時': now, '記録者': me.name });
    else append_('hiki', { '案件ID': id, 'キー': key, '値': value, '日時': now, '記録者': me.name });
    if (key === 'to') update_('deals', findDeal_(id), { '引き継ぎ先': value });
    return true;
  });
}

function apiCancel(id, reason, detail) {
  return withLock_(() => {
    const deal = findDeal_(id);
    const me = me_(), now = new Date();
    setStatus_(deal, { '状態': 'キャンセル', 'キャンセル理由': reason, 'キャンセル詳細': detail || '',
                             'キャンセル日時': now, 'キャンセル者': me.name });
    log_(id, 'キャンセル', reason + ' ' + (detail || ''), me.name);
    return dealToClient_(findDeal_(id));
  });
}

function apiUncancel(id) {
  return withLock_(() => {
    const deal = findDeal_(id);
    const p = progress_(deal, readAll_('items'), readAll_('checks').filter(c => c['案件ID'] === id));
    setStatus_(deal, { '状態': '進行中', 'キャンセル理由': '',
                             'キャンセル詳細': '', 'キャンセル日時': '', 'キャンセル者': '' });
    log_(id, 'キャンセル取消', '', me_().name);
    return dealToClient_(findDeal_(id));
  });
}

// ===== 鍵渡しの予定（担当者のカレンダー） =====
// 入居日の前日 17:00〜20:00 と当日 10:00〜12:00 に「〇〇様🔓鍵渡し」を入れる。
// 担当者が小嶋さん（このアプリの持ち主）に「予定の変更」権限でカレンダーを共有していれば、担当者のカレンダーに直接入る。
// 共有していなければ、持ち主のカレンダーに予定を作って担当者を招待する（同じ会社のアカウントなら担当者のカレンダーにも出る）。
// 入れた予定は「引き継ぎ」シートのキー keyEvents に残し、入れ直すときは前の予定を消してから作る。
const KEY_SLOTS = [
  { dayOffset: -1, from: 17, to: 20 }, // 前日 17:00〜20:00
  { dayOffset: 0, from: 10, to: 12 },  // 当日 10:00〜12:00
];
// メンバーシートにメールが無いときの担当者のメール（見込みボードの STAFF と同じ）
const STAFF_EMAILS = {
  '小嶋勇輝': 'info@monochro-inc.com', '諸本雅広': 'info@monochro-inc.com', '山田園佳': 's.yamada@monochro-inc.com', '佐光琴音': 'k.sakou@monochro-inc.com',
  '澤井颯晟': 'r.sawai@monochro-inc.com', '松井拓登': 't.matui@monochro-inc.com', '寺田喜心': 'k.terada@monochro-inc.com', '宮川大輝': 'd.miyagawa@monochro-inc.com',
};
// 担当のメール。①メンバーシートのフルネーム一致 ②STAFF_EMAILS のフルネーム一致 ③「澤井」のように苗字だけのときは、先頭がそろう人が1人だけならその人
// （「松井」のように同じ苗字が2人いて決められないときは空を返し、画面に「メンバーシートに入れてください」と出す）
// 予定を入れる専用カレンダー（小嶋さん・諸本さんは info@ のふだんのカレンダーが全員に共有されているので、ゲスト招待用に入れる）
const STAFF_CALENDARS = {
  '小嶋': 'c_bb3f0991978892a66def0ffe291316babe115e6d1d8d99a2e182d46ed55232cc@group.calendar.google.com',
  '諸本': 'c_0d189637edc0749f8a43bb4295393c633b58371c184d58f5ea3d93547422811b@group.calendar.google.com',
};
function staffCalendar_(staff) { return STAFF_CALENDARS[String(staff || '').replace(/[\s　]/g, '').slice(0, 2)] || ''; }
function staffEmail_(staff) {
  const norm = s => String(s || '').replace(/[\s　]/g, '');
  const st = norm(staff);
  if (!st) return '';
  const book = {};
  Object.keys(STAFF_EMAILS).forEach(n => { book[n] = STAFF_EMAILS[n]; });
  readAll_('members').forEach(m => { const n = norm(m['売上表の名前']); if (n && m['メール']) book[n] = String(m['メール']).trim(); else if (n && !(n in book)) book[n] = ''; });
  if (book[st]) return book[st];
  const hits = Object.keys(book).filter(n => n !== st && (n.indexOf(st) === 0 || st.indexOf(n) === 0) && book[n]);
  return hits.length === 1 ? book[hits[0]] : '';
}

function apiKeyEvents(id) {
  return withLock_(() => {
    const deal = findDeal_(id);
    const move = deal['入居日'];
    if (!(move instanceof Date)) throw new Error('先に入居日を入れてください');
    const email = staffEmail_(deal['担当']);
    if (!email) throw new Error('担当「' + deal['担当'] + '」のメールが分かりません（契約進捗DBの「メンバー」シートに入れてください）');
    const row = readAll_('hiki').find(h => h['案件ID'] === id && h['キー'] === 'keyEvents');
    let old = {};
    try { old = row ? JSON.parse(row['値']) : {}; } catch (e) { old = {}; }
    (old.refs || []).forEach(deleteEvent_);
    const name = String(deal['お客様名']).trim().replace(/(様|さま)$/, '');
    const title = name + '様🔓鍵渡し';
    const desc = '担当：' + deal['担当'] + (deal['物件'] ? '\n物件：' + deal['物件'] : '') + '\n契約進捗：' + CONFIG.WEB_APP_URL + '?id=' + id;
    const y = move.getFullYear(), mo = move.getMonth(), d = move.getDate();
    const made = KEY_SLOTS.map(s => createEvent_(staffCalendar_(deal['担当']) || email, title, new Date(y, mo, d + s.dayOffset, s.from), new Date(y, mo, d + s.dayOffset, s.to), desc, deal['担当']));
    const val = JSON.stringify({ date: fmtDate_(move), email: staffCalendar_(deal['担当']) || email, refs: made.map(m => m.ref), where: made[0].where });
    const me = me_(), now = new Date();
    if (row) update_('hiki', row, { '値': val, '日時': now, '記録者': me.name });
    else append_('hiki', { '案件ID': id, 'キー': 'keyEvents', '値': val, '日時': now, '記録者': me.name });
    log_(id, '鍵渡しの予定', fmtDate_(move) + ' ' + made[0].where, me.name);
    return { value: val, where: made[0].where };
  });
}

// 担当者のカレンダーに直接作る。書けなければ持ち主のカレンダーに作って担当者を招待する
// 予定の色（Googleカレンダーは決まった11色だけ。担当カラーにいちばん近い色。見込みボードの STAFF の event と同じ）
// 佐光さんのベージュは11色に無いので色を付けず、カレンダー全体の色（ベージュにしておく）で出す
const STAFF_EVENT_COLORS = { '小嶋': 'ORANGE', '諸本': 'CYAN', '山田': 'PALE_RED', '佐光': '', '澤井': 'MAUVE', '松井': 'YELLOW', '寺田': 'GRAY', '宮川': 'RED' };
function paintEvent_(ev, staff) {
  const c = STAFF_EVENT_COLORS[String(staff || '').replace(/[\s　]/g, '').slice(0, 2)];
  if (!c || !CalendarApp.EventColor[c]) return;
  try { ev.setColor(CalendarApp.EventColor[c]); } catch (e) { /* 色が付けられなくても予定はある */ }
}
function createEvent_(email, title, start, end, desc, staff) {
  const own = Session.getEffectiveUser().getEmail();
  if (email && email.toLowerCase() !== own.toLowerCase()) {
    try {
      const cal = CalendarApp.getCalendarById(email);
      if (cal) {
        const ev = cal.createEvent(title, start, end, { description: desc });
        if (email.indexOf('@group.calendar.google.com') < 0) paintEvent_(ev, staff); // 専用カレンダーは色を付けず、カレンダー自体の色で出す
        return { ref: cal.getId() + '|' + ev.getId(), where: email.indexOf('@group.calendar.google.com') > 0 ? '担当の専用カレンダー' : email + ' のカレンダーに直接' };
      }
    } catch (e) { /* 共有されていない・閲覧のみ → 招待にする */ }
  }
  const cal = CalendarApp.getDefaultCalendar();
  const isSelf = !email || email.toLowerCase() === own.toLowerCase();
  const ev = cal.createEvent(title, start, end, isSelf ? { description: desc } : { description: desc, guests: email, sendInvites: false });
  paintEvent_(ev, staff);
  return { ref: cal.getId() + '|' + ev.getId(), where: isSelf ? own + ' のカレンダー' : email + ' を招待' };
}
function deleteEvent_(ref) {
  if (!ref) return;
  const [calId, eventId] = String(ref).split('|');
  try { const ev = CalendarApp.getCalendarById(calId).getEventById(eventId); if (ev) ev.deleteEvent(); } catch (e) { /* 手で消されていれば何もしない */ }
}

/** エディタから実行：担当者ごとに、カレンダーへ直接書き込めるかを「実行ログ」に出す（予定は作らない） */
function checkCalendars() {
  const own = Session.getEffectiveUser().getEmail();
  const names = readAll_('members').map(m => String(m['売上表の名前']).trim()).filter(String);
  [...new Set(names.concat(Object.keys(STAFF_EMAILS)))].forEach(st => {
    const email = staffCalendar_(st) || staffEmail_(st);
    let msg;
    if (!email) msg = 'メール未登録';
    else if (email.toLowerCase() === own.toLowerCase()) msg = '自分のカレンダー（直接入る）';
    else {
      let cal = null;
      try { cal = CalendarApp.getCalendarById(email); } catch (e) { cal = null; }
      msg = !cal ? '共有されていない → 招待で入る'
        : cal.isOwnedByMe() || canEdit_(cal) ? '直接書き込める' : '閲覧だけ共有 → 招待で入る（「予定の変更」にしてもらう）';
    }
    Logger.log(st + '（' + email + '）：' + msg);
  });
}
// カレンダーに書き込めるか：予定を作れるか試して、すぐ消す
function canEdit_(cal) {
  try { const ev = cal.createEvent('（確認用）', new Date(2000, 0, 1), new Date(2000, 0, 1, 0, 1)); ev.deleteEvent(); return true; } catch (e) { return false; }
}

// ===== 48時間アラート =====
function sendStallAlerts() {
  refreshJusetsu_();
  const items = readAll_('items');
  const checks = groupBy_(readAll_('checks'), '案件ID');
  const mails = readAll_('members').reduce((o, m) => { o[m['売上表の名前']] = m['メール']; return o; }, {});
  const admin = CONFIG.ADMIN_EMAIL || Session.getEffectiveUser().getEmail();
  const byTo = {};
  readAll_('deals').forEach(d => {
    if (d['状態'] !== '進行中') return;
    const h = idleHours_(d);
    if (h < CONFIG.STALL_HOURS) return;
    const p = progress_(d, items, checks[d['案件ID']] || []);
    const to = mails[d['担当']] || admin;
    (byTo[to] = byTo[to] || []).push({ d, h, p });
  });
  Object.keys(byTo).forEach(to => {
    const list = byTo[to].sort((a, b) => b.h - a.h);
    const lines = list.map(x => '・' + x.d['お客様名'] + ' 様（' + x.d['担当'] + '）' + Math.floor(x.h / 24) + '日動きなし　'
      + x.p.done + '/' + x.p.total + '　いま：' + (x.p.phase || '-') + '\n　' + CONFIG.WEB_APP_URL + '?id=' + x.d['案件ID']);
    MailApp.sendEmail({
      to,
      subject: '【契約進捗】48時間動きのない案件が' + list.length + '件あります',
      body: '次の案件は、48時間チェックもメモも入っていません。\n'
        + '進んでいればチェックを、待ちの状態なら「面談・メモ」に一言書いてください（カウントがリセットされます）。\n\n'
        + lines.join('\n\n'),
    });
  });
}

// ===== 重説の月次集計 =====
function jusetsuByMonth_() {
  const out = {};
  readAll_('deals').forEach(d => {
    if (!(d['重説完了日'] instanceof Date)) return;
    const month = Utilities.formatDate(d['重説完了日'], CONFIG.TZ, 'yyyy-MM');
    const who = String(d['宅建士'] || '').trim() || '（宅建士未入力）';
    const m = out[month] = out[month] || {};
    (m[who] = m[who] || []).push(d);
  });
  return out;
}

// 「重説集計」シートを作り直す（新しい月が上）
function refreshJusetsu_() {
  const data = jusetsuByMonth_();
  const rows = [];
  Object.keys(data).sort().reverse().forEach(month => {
    Object.keys(data[month]).sort().forEach(who => {
      const list = data[month][who].sort((a, b) => a['重説完了日'] - b['重説完了日']);
      rows.push([month, who, list.length, list.map(d => d['お客様名'] + '（' + Utilities.formatDate(d['重説完了日'], CONFIG.TZ, 'M/d') + '）').join('、')]);
    });
  });
  const sh = sheet_('jusetsu');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 4).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 4).setValues(rows);
}

// 毎月1日：前月の重説件数を宅建士ごとにメールする
function sendMonthlyJusetsu() {
  refreshJusetsu_();
  const now = new Date();
  const prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const month = Utilities.formatDate(prev, CONFIG.TZ, 'yyyy-MM');
  const data = jusetsuByMonth_()[month] || {};
  const names = Object.keys(data).sort((a, b) => data[b].length - data[a].length);
  const total = names.reduce((n, k) => n + data[k].length, 0);
  const lines = names.map(who => '■ ' + who + '：' + data[who].length + '件\n' + data[who]
    .sort((a, b) => a['重説完了日'] - b['重説完了日'])
    .map(d => '　・' + Utilities.formatDate(d['重説完了日'], CONFIG.TZ, 'M/d') + '　' + d['お客様名'] + ' 様（担当 ' + d['担当'] + '）').join('\n'));
  MailApp.sendEmail({
    to: CONFIG.ADMIN_EMAIL || Session.getEffectiveUser().getEmail(),
    subject: '【重説集計】' + (prev.getMonth() + 1) + '月の重説は' + total + '件でした',
    body: (prev.getMonth() + 1) + '月の重説（重説完了日で集計）\n\n' + (lines.join('\n\n') || '重説完了の記録はありませんでした。')
      + '\n\n契約進捗DBの「重説集計」シートで、過去の月も見られます。',
  });
}

// ===== 計算 =====
function progress_(deal, items, checks) {
  const mode = deal['契約方法'] || '対面';
  const list = MODE_LIST[mode] || '対面・郵送';
  const state = checks.reduce((o, c) => { o[c['項目ID']] = c['状態']; return o; }, {});
  let done = 0, total = 0, phase = '';
  const phases = [];
  items.forEach(i => {
    if (i['有効'] === false || i['契約方法'] !== list || i['種類'] === '注意') return;
    if (i['段階'] === MAIL_ONLY_PHASE && mode !== '郵送') return;
    if (phases.indexOf(i['段階']) < 0) phases.push(i['段階']);
    const v = state[i['ID']];
    if (v === 'na') return;
    total++;
    if (v === 'done') done++; else if (!phase) phase = i['段階'];
  });
  return { done, total, phase, phases, phaseIdx: phase ? phases.indexOf(phase) : phases.length };
}

function idleHours_(d) {
  const t = d['最終活動'];
  return t instanceof Date ? (Date.now() - t.getTime()) / 36e5 : 0;
}

// ===== スプレッドシート操作 =====
// 1回の処理の中では、DBを開くのもシートを読むのも1回だけにする（速度のため）
let DB_ = null;
const SHEET_CACHE_ = {}, ROWS_CACHE_ = {};
function db_() {
  if (DB_) return DB_;
  const id = PropertiesService.getScriptProperties().getProperty('DB_ID');
  if (!id) throw new Error('先に setup() を実行してください');
  DB_ = SpreadsheetApp.openById(id);
  return DB_;
}

function sheet_(key) {
  if (!SHEET_CACHE_[key]) SHEET_CACHE_[key] = db_().getSheetByName(SHEET_DEFS[key].name);
  return SHEET_CACHE_[key];
}

function readAll_(key) {
  if (ROWS_CACHE_[key]) return ROWS_CACHE_[key].slice();
  const sh = sheet_(key);
  const last = sh.getLastRow();
  if (last < 2) return (ROWS_CACHE_[key] = []).slice();
  const head = SHEET_DEFS[key].head;
  ROWS_CACHE_[key] = sh.getRange(2, 1, last - 1, head.length).getValues().map((r, i) => {
    const o = { _row: i + 2 };
    head.forEach((h, j) => { o[h] = r[j]; });
    return o;
  }).filter(o => head.some(h => o[h] !== ''));
  return ROWS_CACHE_[key].slice();
}

function append_(key, obj) {
  const head = SHEET_DEFS[key].head;
  const sh = sheet_(key);
  sh.appendRow(head.map(h => (obj[h] === undefined ? '' : obj[h])));
  obj._row = sh.getLastRow();
  delete ROWS_CACHE_[key];
  return obj;
}

function update_(key, obj, patch) {
  const head = SHEET_DEFS[key].head;
  Object.assign(obj, patch);
  sheet_(key).getRange(obj._row, 1, 1, head.length).setValues([head.map(h => (obj[h] === undefined ? '' : obj[h]))]);
}

function deleteRow_(key, obj) { sheet_(key).deleteRow(obj._row); delete ROWS_CACHE_[key]; }

function log_(id, op, text, by) { append_('log', { '日時': new Date(), '案件ID': id, '操作': op, '内容': text, '記録者': by }); }

function findDeal_(id) {
  const d = readAll_('deals').find(x => x['案件ID'] === id);
  if (!d) throw new Error('案件が見つかりません: ' + id);
  return d;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(28000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// ===== 小物 =====
function me_() {
  const email = Session.getActiveUser().getEmail() || '';
  const m = readAll_('members').find(x => String(x['メール']).toLowerCase() === email.toLowerCase());
  return { email, name: m ? m['売上表の名前'] : (email.split('@')[0] || '不明') };
}

function dealToClient_(d) {
  const o = {};
  SHEET_DEFS.deals.head.forEach(h => {
    const v = d[h];
    o[h] = v instanceof Date ? (DATE_FIELDS.indexOf(h) >= 0 ? fmtDate_(v) : fmtDateTime_(v)) : v;
    if (o[h] === undefined) o[h] = '';
  });
  return o;
}

function groupBy_(rows, key) { return rows.reduce((o, r) => { (o[r[key]] = o[r[key]] || []).push(r); return o; }, {}); }
function newId_() { return Utilities.formatDate(new Date(), CONFIG.TZ, 'yyMMdd') + '-' + Utilities.getUuid().slice(0, 6); }
function fmtDate_(v) { return v instanceof Date ? Utilities.formatDate(v, CONFIG.TZ, 'yyyy-MM-dd') : ''; }
function fmtDateTime_(v) { return v instanceof Date ? Utilities.formatDate(v, CONFIG.TZ, "yyyy-MM-dd'T'HH:mm") : ''; }
function parseDateTime_(s) {
  const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4]), Number(m[5])) : '';
}
function parseDate_(s) {
  const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : '';
}

// ===== 項目マスタの初期データ =====
// 種類：空=通常 / 続き=直前の項目の続き（↪表示） / 注意=チェックしない注意書き / 期日のみ=引き継ぎの期日だけの欄
function seedItems_() {
  const T = [
    ['申込時', [
      ['申込書記入', '年収・住所確認'],
      ['NTTコンセント・エアコン・供給番号の確認', '①NTTコンセントの有無 ②エアコン（V数・室外機場所・ダクト穴）③電気・水道供給番号の写メ'],
      ['駐車場サイズを必ず確認'],
      ['過去の滞納歴を聞いておく', '家賃やクレカ'],
      ['身分証明書コピー', '保険証はマスキング必須'],
      ['管理会社FAX・到着確認'],
      ['審査時、保証会社から連絡が入る旨を伝える'],
      ['概算見積書の説明', '相見積対策で出さない場合は担当判断'],
      ['物件資料を印刷しておく', '証拠保存のため'],
      ['申込後の流れシート送付', 'PDF'],
      ['紹介依頼'],
      ['紹介者へのお礼報告'],
      ['お客様への報連相（2日に1回）', '審査・契約書作成が長引いている時']]],
    ['審査承認時', [
      ['契約書到着日を管理会社へ確認'],
      ['重説日・契約日を決定する', '宅建士とも調整', '続き'],
      ['付帯サービスの提案', '電気・ガス・ネット・WS等'],
      ['必要書類連絡 or 申込後契約条件シート送付'],
      ['決済金明細書作成', '概算の時と相違ないかも確認'],
      ['決済金明細と申込後条件シートをLINEで送る'],
      ['請求書をドライブに保存', '3種類'],
      ['鍵渡し日を自分のカレンダーに仮で入れる', '入居日確定後']]],
    ['重説作成時', [
      ['契約書到着'], ['重説作成'], ['登記簿謄本取得', '平日のうちに'], ['ハザードマップ印刷', '物件の場所をマーク'], ['重説の二重チェック']]],
    ['契約書作成時', [
      ['契約書に社判・角印・宅建士印'], ['送金依頼'], ['広告料請求書 or 領収書作成', '領収日は空けておくと便利'],
      ['管理会社宛て契約書郵送時の送付状を印刷'], ['決済金明細書・請求書保存', 'まだの場合']]],
    [MAIL_ONLY_PHASE, [
      ['お客様記入箇所に付箋と鉛筆でマーク', '貼る場所に注意'], ['返送用チェックシート作成'], ['郵送する契約書類を写メで保存', '質問に答えるため'],
      ['返信用レターパックを同封・写メをアルバムに追加'], ['郵送前に郵送先と名義を確認', '宛名不在等を防ぐため']]],
    ['契約時', [
      ['紹介依頼・お客様と写真'], ['Google口コミ依頼'], ['電気・ガス手続きの告知', 'オール電化の時は電源の話'], ['水道手続きのアナウンス'],
      ['インターネット手続きのアナウンス', '事前に電話番号を伝える'], ['署名捺印後、全て写メで送ってもらう', '郵送の場合'], ['紹介者へのお礼報告']]],
    ['契約手続き後', [
      ['契約書類・物件資料を全てスキャン'], ['共有アイテムの取引台帳に格納', 'フォルダ名は「物件名・号室・顧客名・フリガナ」', '続き'],
      ['火災保険の払込受領書をもらう・印刷'], ['契約書類を管理会社へ郵送・写メをアルバムに追加'], ['鍵・契約書控えの到着予定日を確認']]],
    ['鍵渡し時', [
      ['鍵渡し時の書類を作成・交付'], ['鍵渡し書類'], ['Google口コミ依頼', '契約時にもらっていない場合'], ['紹介依頼・お客様と写真', '契約時に撮っていない場合'],
      ['紹介者へのお礼報告'], ['控え郵送前に郵送先と名義を確認'], ['契約書の控え送付・写メをアルバムに追加']]],
  ];
  const E = [
    ['申込時', [
      ['申込書記入', '年収・住所確認'],
      ['駐輪・バイク・駐車場・ペットの申込もその場で行う＆番手確認', '', '続き'],
      ['2人入居不可の物件でないか確認', '1LDK等の場合'],
      ['月額賃料の上乗せ系の説明を先にしておく', '保証料・24Hサポート・水道・町会費など'],
      ['NTTコンセント・エアコン・供給番号の確認', '①NTTコンセントの有無 ②エアコン（V数・室外機場所・ダクト穴）③電気・水道供給番号の写メ'],
      ['駐車場・バイク置き場は車種とサイズを管理会社に必ず確認'],
      ['過去の滞納歴を聞いておく', '家賃やクレカ'],
      ['身分証明書コピー', '保険証はマスキング必須'],
      ['管理会社到着確認'],
      ['審査時、保証会社から連絡が入る旨を伝える'],
      ['概算見積書の説明', '相見積対策で出さない場合は担当判断'],
      ['物件資料（元付け）を印刷しておく', '証拠保存のため'],
      ['申込後の流れシート送付', 'PDF'],
      ['紹介依頼'], ['紹介者へのお礼報告'],
      ['お客様への報連相（2日に1回）', '審査・契約書作成が長引いている時']]],
    ['審査承認時', [
      ['審査承認'], ['契約書到着日を管理会社へ確認'], ['重説日・契約日を決定する', '宅建士とも調整', '続き'],
      ['電子契約の流れをしっかり読む', '必要書類を忘れがちなので注意', '注意'],
      ['付帯サービスの提案', '電気・ガス・ネット・WS等'], ['決済金明細書作成', '概算の時と相違ないかも確認'],
      ['請求書をドライブに保存', '3種類'], ['鍵渡し日を自分のカレンダーに仮で入れる', '入居日確定後。確定できていたらベスト']]],
    ['重説作成時', [
      ['契約書到着'], ['重説作成'],
      ['★仲介としての調査を行う', '使用用途・設備と残置・特約・概算と明細の相違がないか・駐輪駐車関係の書類もあるか'],
      ['登記簿謄本取得', '平日のうちに'], ['ハザードマップ印刷', '物件の場所をマーク'], ['重説の二重チェック']]],
    ['契約書作成時', [
      ['仲介の電子認証を終わらせる'], ['送金依頼'], ['広告料請求書 or 領収書作成', '領収日は空けておくと便利'], ['決済金明細書・請求書保存', 'まだの場合']]],
    ['契約時', [
      ['紹介依頼・お客様と写真'], ['Google口コミ依頼'], ['電気・ガス手続きの告知', 'オール電化の時は電源の話'], ['水道手続きのアナウンス'],
      ['インターネット手続きのアナウンス', '事前に電話番号を伝える'], ['鍵渡し日の確定＆LINEで日時の証拠も送っておく'],
      ['署名捺印後、全て写メで送ってもらう', '郵送の場合'], ['紹介者へのお礼報告']]],
    ['契約手続き後', [
      ['契約書類・物件資料を全てスキャン'], ['共有アイテムの取引台帳に格納', 'フォルダ名は「物件名・号室・顧客名・フリガナ」', '続き'],
      ['火災保険の払込受領書をもらう・印刷'], ['契約書類を管理会社へ郵送・写メをアルバムに追加'],
      ['メールボックス・宅配ボックス・駐輪駐車区画・ステッカーの確認', '①メールボックス ②宅配ボックスの開け方 ③駐輪・駐車場区画 ④ステッカーはもらえるか。郵送の鍵渡し書類にあればOK'],
      ['鍵・契約書控えの到着予定日を確認']]],
    ['鍵渡し時', [
      ['鍵渡し時の書類を作成・交付'], ['鍵渡し書類'], ['Google口コミ依頼', '契約時にもらっていない場合'], ['紹介依頼・お客様と写真', '契約時に撮っていない場合'],
      ['紹介者へのお礼報告'], ['控え郵送前に郵送先・名義と転送届の有無を確認', '宛名不在等を防ぐため'], ['契約書の控え送付・写メをアルバムに追加']]],
  ];
  const H = [
    ['管理から契約書到着', [['作成済み不動産業務ツールの有無'], ['重説作成・ハザード・登記簿（EV台数・エアコン台数）'], ['明細作成（手数料・割引理由・振込期日）'],
      ['請求書保存'], ['送金依頼（済・待）'], ['申込情報（イタンジ・紙面・LSTEP・その他）'], ['社判']]],
    ['郵送契約の場合', [['付箋を貼る・シャーペンでマーキング'], ['返信用チェックリスト'], ['お客様へ発送準備（レターパック青、緊急時は赤）'], ['郵送書類の写メ'], ['返信用封筒（お客様→Monochro）']]],
    ['お客様から契約書到着', [['書類スキャン'], ['領収書 or 請求書の作成'], ['管理会社への発送準備（普通郵便 or レターパック）']]],
    ['その他', [['エイワン（火災保険）申込'], ['保証会社申込（全保連・セーフティー・その他）'], ['電気ガス取次（インターネット取次の有無）'], ['オール電化の有無'], ['電話の都合の良い時間帯']]],
    ['契約書控え新居へ郵送', [['', '', '期日のみ']]],
  ];
  const rows = [];
  const push = (prefix, list, groups) => {
    let n = 0;
    groups.forEach(([phase, items]) => items.forEach(it => {
      n++;
      rows.push([prefix + ('00' + n).slice(-3), list, phase, it[0] || '', it[1] || '', it[2] || '', true]);
    }));
  };
  push('T', '対面・郵送', T);
  push('E', '電子', E);
  push('H', '引き継ぎ', H);
  return rows;
}
