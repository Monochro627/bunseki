/**
 * 見込みボード（顧客管理・進捗管理シート用 GAS Webアプリ）
 *
 * 元データ：「顧客管理」シート。アプリが使う列は右端に自動追加する。
 *   見込み月      … "2026-11" / "空中戦" / "その他" / 空欄（＝未振り分け）
 *   議事録フォルダ … 顧客ごとの Drive フォルダID（議事録ログのキーも兼ねる）
 *   最終更新      … "小嶋 10/20 14:05"
 *   追客アラート日 … "2026-10-25"（前日18時と当日8時に担当者へメール）
 *   アラート内容  … 一言メモ
 *   アラート予定  … Googleカレンダーの予定（"カレンダーID|予定ID"）
 */

const SPREADSHEET_ID = '1-F5ckS4tNKnHXTrE4jUB2Tcz804lx2vdMQkJvNANvRo';
const SHEET_NAME = '顧客管理';
const NOTES_SHEET_NAME = '議事録ログ';
const NOTES_ROOT_FOLDER_NAME = '顧客議事録';

// 売上ゲージ用：「売上管理表」の5期目シート（読むだけ。書き込みはしない）
const SALES_SPREADSHEET_ID = '1QLNfTOsi6BpFPkxZbp9e9mjxbvFGy-6TZi6uTksOuxI';
const SALES_SHEET_NAME = '5期目　年間売上表';
const SALES_PERIOD_START = '2026-04-01'; // 5期目の期間（日数経過率の計算に使う）
const SALES_PERIOD_END = '2027-03-31';
// 複数の行の合計で目標を追いかける担当者（ラベルは空白を除いた形で完全一致）
const SALES_COMBINED = {
  '諸本雅広': {
    actual: ['インスタ※元売上', '飛び込み・HP※元売上', '諸本雅広※元売上(自己集客）'],
    apps: ['インスタ※元申込件数', '飛び込み・HP※元申込件数', '諸本雅広※元申込件数'],
    saleActual: ['（分析）インスタ売買元売上', '（分析）飛び込み・口コミ・HP売買'],
    saleApps: ['（分析）インスタ売買成約件数', '（分析）飛び込み・口コミ・HP売買成約件数'],
  },
};
// 平均単価は賃貸だけで出すので、売買の売上と申込件数を引く。担当者ごとの売買の行（名前は n に入る）
const SALE_ACTUAL_LABELS = (n) => [n + '※売買案件', '【分析】' + n + '④売買総売上'];
const SALE_APPS_LABELS = (n) => ['【分析】' + n + '※売買申込数', '【分析】' + n + '④売買案件'];

// 担当者と会社アカウントの対応。開いた人の担当分を最初に表示するのに使う。
// color は売上管理表（5期目　年間売上表）の目標行の背景色。ゲージと担当者ボタンの色に使う
// event はカレンダーの予定の色（Googleカレンダーは決まった11色だけ。担当カラーにいちばん近い色）。
//   ORANGE=ミカン YELLOW=バナナ PALE_RED=フラミンゴ MAUVE=ブドウ CYAN=ピーコック GRAY=グラファイト RED=トマト
//   PALE_BLUE=ラベンダー PALE_GREEN=セージ BLUE=ブルーベリー GREEN=バジル
//   空なら色を付けない（カレンダー全体の色で出る）。佐光さんのベージュは11色に無いので、カレンダー全体の色をベージュにして使う
// calendar は予定を入れる専用カレンダー（無ければ email のカレンダー）。小嶋さん・諸本さんは info@ のふだんのカレンダーが全員に
//   共有されているので、それぞれの「ゲスト招待用」カレンダーに入れる（メールは今までどおり email に送る）
const STAFF = [
  { name: '小嶋勇輝', email: 'info@monochro-inc.com', color: '#ff9900', event: 'ORANGE', calendar: 'c_bb3f0991978892a66def0ffe291316babe115e6d1d8d99a2e182d46ed55232cc@group.calendar.google.com' },
  { name: '諸本雅広', email: 'info@monochro-inc.com', color: '#6d9eeb', event: 'CYAN', calendar: 'c_0d189637edc0749f8a43bb4295393c633b58371c184d58f5ea3d93547422811b@group.calendar.google.com' },
  { name: '山田園佳', email: 's.yamada@monochro-inc.com', color: '#f4cccc', event: 'PALE_RED' },
  { name: '佐光琴音', email: 'k.sakou@monochro-inc.com', color: '#ebdfc2', event: '' },
  { name: '澤井颯晟', email: 'r.sawai@monochro-inc.com', color: '#b4a7d6', event: 'MAUVE' },
  { name: '松井拓登', email: 't.matui@monochro-inc.com', color: '#ffff00', event: 'YELLOW' },
  { name: '寺田喜心', email: 'k.terada@monochro-inc.com', color: '#000000', event: 'GRAY' },
  { name: '宮川大輝', email: 'd.miyagawa@monochro-inc.com', color: '#ff0000', event: 'RED' },
];

// 成約までの段階。ボタンを押すと日付の列に今日の日付を入れ、B列のステータスを書き換え、進捗ログに1行残す
const STAGES = [
  { key: 'call', label: '初回電話', col: 'firstCall', status: '初回電話済み' },
  { key: 'meeting', label: '面談', col: 'meeting', status: 'カウンセリング面談済み' },
  { key: 'proposal', label: '物件提案', col: 'proposal', status: '物件提案中' },
  { key: 'viewing', label: '内覧', col: 'viewing', status: '内覧済み' }, // 押すたびに内覧回数が増える
  { key: 'apply', label: '成約', col: 'applied', status: '完了' }, // アプリでは「成約」、シートのB列は「完了」（紫になる書式に合わせる）。日付はQ列の申込日
];
// 段階以外の状態（日付は残さず、B列だけ書き換える）
const STATUS_BUTTONS = ['追客', '時期待ち', '返答待ち'];
const PROGRESS_LOG_SHEET_NAME = '進捗ログ';
const MONTH_TARGET_SHEET_NAME = '今月の目標'; // 担当者ごとの月の売上目標（アプリから手入力）

// 予定を先に入れられる段階。予定は分析に数えず、「実施した」を押したときに実施として記録する
// minutes は所要時間の初期値（画面のプルダウンで30分単位に変えられる）
const PLAN_STAGES = {
  call: { col: 'planCall', title: '【初回電話】', minutes: 30 },
  meeting: { col: 'planMeeting', title: '【面談】', minutes: 60 },
  viewing: { col: 'planViewing', title: '【内覧】', minutes: 300 }, // 内覧は移動も含めて半日見ておく
};
// 取引種別（D列のプルダウンと同じ）
const DEAL_TYPES = ['賃貸', '売買（購入）', '売買（売却）'];
const BOARD_FROM = '2026-10'; // 成約・失客をボードに出すのはこの月以降（アプリを使い始めた月。それより前は出さない）
const FUNNEL_FROM = '26/10'; // 分析の対象：紹介月（F列 "yy/mm"）がこの月以降の顧客（アプリを使い始めた10月から）

// 読み込む列（ヘッダー名で探す）
const COLS = {
  name: '顧客名',
  status: '進捗ステータス',
  memo: 'メモ',
  dealType: '取引種別',
  staff: '担当者',
  introMonth: '紹介月',
  introType: '紹介種別',
  moveTime: '引越し時期',
  firstCall: '初回電話日',
  meeting: '面談日\n（カウンセリング面談）',
  viewing: '内覧日',
  viewCount: '内覧回数',
  applied: '申込日',
  property: '物件名・号室',
  moveIn: '入居日',
};
const APP_COLS = {
  forecast: '見込み月', folder: '議事録フォルダ', updated: '最終更新',
  alertDate: '追客アラート日', alertNote: 'アラート内容', alertEvent: 'アラート予定',
  proposal: '物件提案日',
  planCall: '初回電話予定', planMeeting: '面談予定', planViewing: '内覧予定', planEvents: '予定カレンダー',
  order: '並び順', // 列の中での並び順（小さいほど上）。ボードで上下にドラッグすると入る
  contract: '契約進捗ID', // 契約進捗アプリの案件ID（成約したときに作る／つなぐ）
};

const CACHE_KEY = 'board_v17';
// 保存の順番待ち（ロック）を待つ最大の時間。混んでいても失敗しないように長め（Apps Script は1回6分まで動ける）
const LOCK_WAIT = 30000;
const CACHE_TTL = 600; // 10分。アプリから書くとき・シートを直接直したとき（onSheetEdit）は捨てるので、長めでよい
const SALES_CACHE_KEY = 'sales_v9';

function doGet(e) {
  if (e && e.parameter && e.parameter.page === 'home') return homePage_(); // ?page=home はホーム画面（アプリの入口）
  if (e && e.parameter && e.parameter.page === 'docs') return docsPage_(); // ?page=docs は営業資料（ドライブのフォルダをカードで見せる）
  const isAnalysis = e && e.parameter && e.parameter.page === 'analysis';
  const template = HtmlService.createTemplateFromFile(isAnalysis ? 'Analysis' : 'Index');
  // 見込みボードと分析ボードは同じアプリの別ページ。上のタブで行き来する（?page=analysis で分析ボード、?view=sale で売買案件）
  template.appUrl = ScriptApp.getService().getUrl();
  template.initialView = (e && e.parameter && e.parameter.view) || '';
  template.contractUrl = CONTRACT_DB_ID ? CONTRACT_APP_URL : ''; // 契約進捗アプリ（つないでいなければ空）
  template.manualUrl = MANUAL_APP_URL; // マニュアルのアプリ（空ならタブを出さない）
  if (isAnalysis) {
    template.boardUrl = template.appUrl;
    template.analysisPage = ANALYSIS_ALIASES[e.parameter.p] || e.parameter.p || ANALYSIS_PAGES[0].key;
  }
  return template
    .evaluate()
    .setTitle(isAnalysis ? '成績ダッシュボード' : '見込みボード')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // Notion などのページに埋め込めるようにする（開ける人は今までどおり社内アカウントだけ）
}

/** 初回に1回だけエディタから実行：アプリ用の列と議事録ログシートを作る */
function setup() {
  const sheet = getSheet_();
  ensureAppColumns_(sheet);
  getNotesSheet_();
  getNotesRoot_();
  installAlertTriggers_();
  installEditTrigger_();
  clearCache_();
}

/** シートを直接直したとき、アプリがためている顧客の一覧をすぐ捨てる（次に開いたとき・↻ でシートの値が出る） */
function onSheetEdit() { clearCache_(); }
/**
 * 1分ごと：顧客の一覧がためられていなければ作っておく（書き込みのあと最初に開いた人が待たされないように）。
 * ためてあれば何もしないので軽い
 */
function warmBoardCache() {
  if (!readCache_()) writeCache_(buildBoard_());
}
function installEditTrigger_() {
  ScriptApp.getProjectTriggers().filter((t) => ['onSheetEdit', 'warmBoardCache'].indexOf(t.getHandlerFunction()) >= 0).forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onSheetEdit').forSpreadsheet(SPREADSHEET_ID).onEdit().create();
  ScriptApp.newTrigger('warmBoardCache').timeBased().everyMinutes(1).create();
}

// ---------- 読み込み ----------

function getBoard(fresh) {
  const cached = fresh ? null : readCache_(); // ↻ のときはためた分を使わずシートから読み直す
  const data = cached || buildBoard_();
  if (!cached) writeCache_(data);
  return Object.assign({}, data, { me: currentStaffName_() });
}

function buildBoard_() {
  const sheet = getSheet_();
  const map = ensureAppColumns_(sheet);
  const lastRow = sheet.getLastRow();
  const values = lastRow < 2 ? [] : sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getDisplayValues();
  // 申込日は表示が「09/11」で年が無いので、日付の値から年月を取る（成約した月の列に出すため）
  const appliedDates = lastRow < 2 || map.applied == null ? [] : sheet.getRange(2, map.applied + 1, lastRow - 1, 1).getValues();
  const noteCounts = countNotes_();

  const customers = [];
  values.forEach((r, i) => {
    const name = r[map.name].trim();
    const status = r[map.status].trim();
    const result = resultOf_(status);
    // B列が「完了」でないのに申込日が入っている人は、アプリを使う前の過去の案件。ボードには出さない（売買案件の一覧の「完了・失客も表示」には出る）
    const pastApplied = !!r[map.applied] && !result;
    const closed = isClosed_(status) || pastApplied;
    // 売買案件は完了・失客も一覧で振り返れるように読み込む（ボードには出さない）
    const isSale = map.dealType != null && r[map.dealType].indexOf('売買') >= 0;
    // ×でボードから外した人（見込み月が「非表示」で始まる）。ボードには出さず、「外した人」から戻せる
    const hidden = String(r[map.forecast]).indexOf(HIDDEN_MARK) === 0;
    // 成約・失客した顧客は、その月の列に記録として残す（月をさかのぼって振り返れる）
    // 成約は申込日の月（申込日が無いか、今月より先の日付なら見込み月）、失客は見込み月の列
    // 出すのは BOARD_FROM（2026年10月）から今月まで。申込日を「11/5」のように年なしで打つと翌年の日付になることがあるので、先の月は使わない
    const now = thisMonth_();
    const inRange = (m) => (m && m >= BOARD_FROM && m <= now ? m : '');
    const hasMonth = /^\d{4}-\d{2}$/.test(r[map.forecast]);
    const d = (appliedDates[i] || [])[0];
    const appliedYm = d instanceof Date ? Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM') : '';
    const forecastYm = hasMonth ? r[map.forecast] : '';
    const wonMonth = appliedYm && appliedYm <= now ? appliedYm : forecastYm;
    const boardMonth = hidden ? '' : result === 'won' ? inRange(wonMonth) : result === 'lost' ? inRange(forecastYm) : '';
    if (!name || (closed && !isSale && !boardMonth && !hidden)) return;
    // インスタ：会社集客（インスタ）か「◯◯（インスタ送客）」は main、インスタ顧客からの紹介・リピートは sub（見た目を変えて参考に出すだけ。件数には数えない）
    const insta = instaKind_(r[map.introType], r[map.staff]);
    let owners = parseOwners_(r[map.staff]);
    // 主担当も送客元も8名以外なら出さない（インスタの人だけは「インスタ」の欄に出すので読む）
    if (!owners) {
      if (!insta) return;
      owners = { staff: String(r[map.staff] || '').replace(/[（(].*$/, '').replace(/[\s　]/g, ''), referrer: '', owners: [] };
    }
    const folderId = r[map.folder];
    const c = {
      row: i + 2, forecast: r[map.forecast], folderId: folderId, updated: r[map.updated], closed: closed,
      alertDate: r[map.alertDate], alertNote: r[map.alertNote], proposal: r[map.proposal],
      planCall: r[map.planCall], planMeeting: r[map.planMeeting], planViewing: r[map.planViewing],
      order: r[map.order] === '' ? null : Number(r[map.order]),
      result: result, boardMonth: boardMonth, hidden: hidden, contractId: r[map.contract], insta: insta,
      staffRaw: String(r[map.staff] || '').trim(), // E列そのまま（担当の変更で今の値を出すため）
    };
    Object.keys(COLS).forEach((k) => { c[k] = map[k] == null ? '' : r[map[k]]; });
    Object.assign(c, owners);
    c.notes = folderId ? (noteCounts[folderId] || 0) : 0;
    customers.push(c);
  });

  return {
    customers: customers, staff: STAFF.map((s) => s.name), statusButtons: STATUS_BUTTONS,
    stages: STAGES.map((st) => ({ key: st.key, label: st.label })), loadedAt: new Date().toISOString(),
    dealTypes: DEAL_TYPES, staffOptions: staffOptions_(sheet, map), planMinutes: Object.keys(PLAN_STAGES).reduce((o, k) => { o[k] = PLAN_STAGES[k].minutes; return o; }, {}),
    staffColors: STAFF.reduce((o, s) => { o[s.name] = s.color; return o; }, {}),
  };
}

/**
 * 担当者欄を主担当と送客元に分ける。例：「澤井颯晟（諸本送客）」→ 主担当 澤井颯晟、送客元 諸本雅広
 * 主担当が8名以外でも送客元が8名なら、送客元の画面にだけ出す。どちらも8名以外なら null。
 * 送客元が8名以外（インスタ・神戸アカ・飛び込み等）なら送客元は空。
 */
function parseOwners_(raw) {
  const text = String(raw || '').trim();
  const m = text.match(/^(.*?)[\s　]*[（(](.*)[)）]\s*$/);
  const primaryName = (m ? m[1] : text).replace(/[\s　]/g, '');
  const primary = STAFF.find((s) => s.name === primaryName);
  let referrer = '';
  if (m && m[2].indexOf('送客') >= 0) {
    const from = m[2].replace('送客', '').replace(/[\s　]/g, '');
    const hit = from && STAFF.find((s) => s.name !== primaryName && (s.name === from || s.name.indexOf(from) === 0));
    if (hit) referrer = hit.name;
  }
  const owners = [];
  if (primary) owners.push(primary.name);
  if (referrer) owners.push(referrer);
  if (!owners.length) return null;
  return { staff: primaryName, referrer: referrer, owners: owners };
}

/** インスタの顧客か。'main'＝会社集客のインスタ（KPIに入る）、'sub'＝インスタ顧客による紹介・リピート（自己集客。参考）、''＝ちがう */
function instaKind_(introType, rawStaff) {
  const t = String(introType || ''), s = String(rawStaff || '');
  if ((t.indexOf('会社集客') === 0 && t.indexOf('インスタ') >= 0) || s.indexOf('インスタ送客') >= 0) return 'main';
  if (t.indexOf('インスタ顧客') >= 0) return 'sub';
  return '';
}

/** 成約なら 'won'（B列が「完了」。アプリの「成約」ボタンで入る）、失客なら 'lost'、それ以外は '' */
function resultOf_(status) {
  if (status.indexOf('失客') >= 0) return 'lost';
  if (status.indexOf('完了') === 0) return 'won';
  return '';
}

/** 成約・失客した月としてボードに残すため、見込み月を今月にする */
function thisMonth_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
}

function isClosed_(status) {
  return status.indexOf('完了') === 0 || status.indexOf('失客') >= 0 || status.indexOf('送客被り') >= 0;
}

// ---------- 書き込み ----------

// ×でボードから外すときに見込み月へ入れる印。元の値を「非表示:2026-11」のように後ろに残し、戻すときに使う
const HIDDEN_MARK = '非表示';

/** ×ボタン：ボードから外す（見込み月を「非表示:元の値」にするだけ。B列やほかの列は触らない） */
function hideCustomer(row, name) {
  const sheet = getSheet_();
  const map = ensureAppColumns_(sheet);
  const actualRow = resolveRow_(sheet, map, row, name);
  const cur = sheet.getRange(actualRow, map.forecast + 1).getDisplayValue();
  const value = cur.indexOf(HIDDEN_MARK) === 0 ? cur : HIDDEN_MARK + (cur ? ':' + cur : '');
  const res = writeFields_(actualRow, name, { forecast: value });
  res.forecast = value;
  return res;
}

/** 外した人を戻す（見込み月を外す前の値に戻す） */
function restoreCustomer(row, name) {
  const sheet = getSheet_();
  const map = ensureAppColumns_(sheet);
  const actualRow = resolveRow_(sheet, map, row, name);
  const cur = sheet.getRange(actualRow, map.forecast + 1).getDisplayValue();
  const value = cur.indexOf(HIDDEN_MARK) === 0 ? cur.slice(HIDDEN_MARK.length).replace(/^[:：]/, '') : cur;
  const res = writeFields_(actualRow, name, { forecast: value });
  res.forecast = value;
  return res;
}

/** 見込み月を変更（ドラッグ） value: "2026-11" / "空中戦" / "その他" / "" */
function moveCustomer(row, name, value) {
  if (value && !/^\d{4}-\d{2}$/.test(value) && value !== '空中戦' && value !== 'その他') {
    throw new Error('不正な見込み月です: ' + value);
  }
  return writeFields_(row, name, { forecast: value });
}

/**
 * 複数まとめて見込み月を変える（複数選択の移動）。items: [{row, name}]。1回の順番待ちでまとめて書く
 * 戻り値：[{row, updated}]（items と同じ順）
 */
function moveCustomers(items, value) {
  if (value && !/^\d{4}-\d{2}$/.test(value) && value !== '空中戦' && value !== 'その他') throw new Error('不正な見込み月です: ' + value);
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_WAIT);
  try {
    const sheet = getSheet_();
    const map = ensureAppColumns_(sheet);
    const stamp = updateStamp_();
    const out = (items || []).map((it) => {
      const actualRow = resolveRow_(sheet, map, it.row, it.name);
      sheet.getRange(actualRow, map.forecast + 1).setValue(value);
      sheet.getRange(actualRow, map.updated + 1).setValue(stamp);
      return { row: actualRow, updated: stamp };
    });
    clearCache_();
    return out;
  } finally {
    lock.releaseLock();
  }
}

/**
 * 段階ボタン。stage: STAGES の key。undo=true で取り消し（押し間違い用）
 * 内覧は押すたびに回数+1（N列の内覧日は1回目の日付のまま、O列を「2回」のように更新）
 */
function recordStage(row, name, stageKey, undo) {
  // 成約の契約進捗づくりは時間がかかるので、画面が成約を出したあとに contractAfterWin で別に呼ぶ
  return recordStage_(row, name, stageKey, undo);
}

/** 成約のあと（画面から裏で呼ぶ）：賃貸なら契約進捗の案件を作る（同じ人の案件があればつなぐ）。{ contractId } か {} */
function contractAfterWin(row, name) {
  try { const ct = ensureContract_(row, name, true); return ct ? { contractId: ct.id, row: ct.row } : {}; }
  catch (e) { return { warning: '契約進捗を作れませんでした：' + e.message }; }
}

// 行を1回だけ読んで、変わったセルだけ書く（1セルずつ読むと遅いため）
function recordStage_(row, name, stageKey, undo) {
  const stage = STAGES.find((st) => st.key === stageKey);
  if (!stage) throw new Error('不明な段階です: ' + stageKey);
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_WAIT);
  let oldAlertRef = '';
  try {
    const sheet = getSheet_();
    const map = ensureAppColumns_(sheet);
    const actualRow = resolveRow_(sheet, map, row, name);
    const r = sheet.getRange(actualRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    const val = {}; // 今の値（書いたものはここも書き換える）
    Object.keys(map).forEach((k) => { if (map[k] != null) val[k] = r[map[k]]; });
    const put = (k, v) => { sheet.getRange(actualRow, map[k] + 1).setValue(v); val[k] = v === '' ? '' : String(v instanceof Date ? 'd' : v); };
    // 予定が入っていて、その日を過ぎていれば予定の日で実施を記録する（押すのが遅れても日付がずれない）。予定は消す
    let today = new Date();
    const plan = !undo && PLAN_STAGES[stage.key] ? parsePlan_(val[PLAN_STAGES[stage.key].col]) : null;
    if (plan) {
      if (plan < today) today = plan;
      put(PLAN_STAGES[stage.key].col, '');
      const refs = readPlanRefs_(val.planEvents);
      delete refs[stage.key]; // カレンダーの予定は実施したものとして残す
      put('planEvents', Object.keys(refs).length ? JSON.stringify(refs) : '');
    }
    const todayText = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd');
    const setDate = (k, v) => put(k, k === 'proposal' ? (v ? todayText : '') : (v ? today : ''));
    let count = parseInt(val.viewCount, 10) || 0;
    let countWarning = '';

    if (stage.key === 'viewing') {
      count = Math.max(0, count + (undo ? -1 : 1));
      if (!undo && !val.viewing) setDate('viewing', true);
      if (undo && count === 0) setDate('viewing', false);
      try {
        put('viewCount', count ? count + '回' : '');
      } catch (e) {
        countWarning = 'O列「内覧回数」のプルダウンに「' + count + '回」が無いため書き込めませんでした（進捗ログには記録済み）';
      }
    } else {
      setDate(stage.col, !undo);
    }
    if (stage.key === 'apply' && !undo) {
      put('forecast', thisMonth_());
      // 成約したら追客アラートは不要（カレンダーの予定はロックを外してから消す）
      if (val.alertDate || val.alertEvent) { oldAlertRef = val.alertEvent; put('alertDate', ''); put('alertNote', ''); put('alertEvent', ''); }
    }

    // B列：記録が残っている一番先の段階に合わせる
    const reached = STAGES.filter((st) => (st.key === 'viewing' ? count > 0 || val.viewing : val[st.col]));
    const top = reached[reached.length - 1];
    const status = top ? (top.key === 'viewing' ? '内覧済み（' + Math.max(count, 1) + '回）' : top.status) : '';
    put('status', status);
    const stamp = updateStamp_();
    put('updated', stamp);

    getProgressLogSheet_().appendRow([
      Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'), name, actualRow,
      r[map.staff], stage.label,
      stage.key === 'viewing' ? count : '', undo ? '取消' : '', currentStaffName_() || currentEmail_(),
    ]);
    clearCache_();

    const r2 = sheet.getRange(actualRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    const result = resultOf_(status);
    return {
      row: actualRow, updated: stamp, status: status, warning: countWarning,
      firstCall: r2[map.firstCall], meeting: r2[map.meeting], proposal: r2[map.proposal],
      viewing: r2[map.viewing], viewCount: r2[map.viewCount], applied: r2[map.applied],
      forecast: r2[map.forecast], result: result,
      planCall: r2[map.planCall], planMeeting: r2[map.planMeeting], planViewing: r2[map.planViewing],
      alertDate: r2[map.alertDate], alertNote: r2[map.alertNote],
      boardMonth: result ? r2[map.forecast] : '',
    };
  } finally {
    lock.releaseLock();
    if (oldAlertRef) deleteAlertEvent_(oldAlertRef);
  }
}

/**
 * 段階の予定を入れる（変更も同じ）。when: "2026-10-25T14:00"、minutes: 所要時間（30分単位）
 * 担当者のカレンダーに所要時間分の予定を作り、前の予定は消す。シートには "2026-10-25 14:00〜15:30" の形で入れる
 * 予定は分析に数えない。「実施した」を押すと recordStage で実施として記録される
 */
function setPlan(row, name, stageKey, when, minutes) {
  const ps = PLAN_STAGES[stageKey];
  if (!ps) throw new Error('この段階には予定を入れられません: ' + stageKey);
  const start = parsePlan_(when);
  if (!start) throw new Error('日時を選んでください');
  const hour = Number(Utilities.formatDate(start, 'Asia/Tokyo', 'H'));
  if (hour < 8 || hour > 22) throw new Error('開始は 8:00〜22:30 の間で入れてください');
  const mins = Number(minutes) || ps.minutes;
  if (mins < 30 || mins > 480 || mins % 30) throw new Error('所要時間は30分単位で選んでください');
  const end = new Date(start.getTime() + mins * 60000);
  const sheet = getSheet_();
  const map = ensureAppColumns_(sheet);
  const actualRow = resolveRow_(sheet, map, row, name);
  const r = sheet.getRange(actualRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const refs = readPlanRefs_(r[map.planEvents]);
  deleteAlertEvent_(refs[stageKey]);
  const event = createCalendarEvent_(alertOwner_(r[map.staff]), ps.title + calendarName_(name), start, end,
    '顧客：' + name + '\n担当：' + r[map.staff]);
  refs[stageKey] = event.ref;
  const text = Utilities.formatDate(start, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') + '〜' + Utilities.formatDate(end, 'Asia/Tokyo', 'HH:mm');
  const fields = { planEvents: JSON.stringify(refs) };
  fields[ps.col] = text;
  const res = writeFields_(actualRow, name, fields);
  res[ps.col] = text;
  res.calendar = event.where;
  return res;
}

/** 予定を取り消す（カレンダーの予定も消す） */
function clearPlan(row, name, stageKey) {
  const ps = PLAN_STAGES[stageKey];
  if (!ps) throw new Error('この段階には予定を入れられません: ' + stageKey);
  const sheet = getSheet_();
  const map = ensureAppColumns_(sheet);
  const actualRow = resolveRow_(sheet, map, row, name);
  const refs = readPlanRefs_(sheet.getRange(actualRow, map.planEvents + 1).getDisplayValue());
  deleteAlertEvent_(refs[stageKey]);
  delete refs[stageKey];
  const fields = { planEvents: Object.keys(refs).length ? JSON.stringify(refs) : '' };
  fields[ps.col] = '';
  const res = writeFields_(actualRow, name, fields);
  res[ps.col] = '';
  return res;
}

/** カレンダーの件名に使う顧客名。末尾のフリガナ「（ナカタ）」だけ外す（名前が丸ごと消えないように） */
function calendarName_(name) {
  const n = String(name || '').trim();
  return n.replace(/\s*[（(][^（()）]*[)）]\s*$/, '') || n;
}

/**
 * 顧客名とフリガナを直す（A列は「名前（フリガナ）」の形）。
 * あわせて議事録フォルダの名前と、カレンダーの予定（追客・面談などの件名）も新しい名前に変える
 */
function renameCustomer(row, name, newName, kana) {
  newName = String(newName || '').trim();
  kana = String(kana || '').trim();
  if (!newName) throw new Error('名前を入力してください');
  if (/[（()）]/.test(newName) || /[（()）]/.test(kana)) throw new Error('名前とフリガナには括弧を入れずに入力してください');
  const full = kana ? newName + '（' + kana + '）' : newName;
  const sheet = getSheet_();
  const map = ensureAppColumns_(sheet);
  const actualRow = resolveRow_(sheet, map, row, name);
  const r = sheet.getRange(actualRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const res = writeFields_(actualRow, name, { name: full });

  // ここから先は失敗しても名前の変更は済んでいるので、止めずに続ける
  if (r[map.folder]) {
    try { DriveApp.getFolderById(r[map.folder]).setName(full); } catch (e) { /* フォルダが無ければ何もしない */ }
  }
  const refs = [r[map.alertEvent]].concat(Object.values(readPlanRefs_(r[map.planEvents]))).filter(Boolean);
  refs.forEach((ref) => {
    try {
      const [calId, eventId] = ref.split('|');
      const ev = CalendarApp.getCalendarById(calId).getEventById(eventId);
      if (ev) ev.setTitle(ev.getTitle().replace(/^(【[^】]*】).*$/, '$1' + calendarName_(full)));
    } catch (e) { /* 消された予定は何もしない */ }
  });
  // 契約進捗の案件のお客様名も合わせる（売上表につなぐときに名前で探すため）
  if (map.contract != null && r[map.contract] && CONTRACT_DB_ID) {
    try {
      const sh = contractDb_().getSheetByName('案件');
      const v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
      const idCol = v[0].indexOf('案件ID'), nmCol = v[0].indexOf('お客様名');
      const i = v.findIndex((x, k) => k > 0 && x[idCol] === r[map.contract]);
      if (i > 0 && nmCol >= 0) sh.getRange(i + 1, nmCol + 1).setValue(full);
      CacheService.getScriptCache().remove(CONTRACT_CACHE_KEY);
    } catch (e) { /* 契約進捗に届かなくても名前の変更は済んでいる */ }
  }
  res.name = full;
  return res;
}

/**
 * 担当者欄（E列）の選択肢。E列のプルダウン（入力規則）の範囲をそのまま読む（「澤井颯晟（小嶋送客）」など）
 * row の行のプルダウンを見る。無ければ一番下の行から上へ、プルダウンの付いている行を探す（2行目などプルダウンの無い行があるため）
 * どこにも無ければ8名の名前だけ
 */
function staffOptions_(sheet, map, row) {
  const read = (r) => {
    try {
      const rule = sheet.getRange(r, map.staff + 1).getDataValidation();
      const v = rule && rule.getCriteriaValues()[0];
      if (v && v.getValues) return v.getDisplayValues().map((x) => String(x[0]).trim()).filter(String);
      if (Array.isArray(v)) return v.map(String).filter(String);
    } catch (e) { /* 読めなければ次へ */ }
    return null;
  };
  let list = row ? read(row) : null;
  if (!list) {
    const last = sheet.getLastRow();
    const rules = sheet.getRange(Math.max(2, last - 200), map.staff + 1, Math.min(201, last - 1), 1).getDataValidations();
    for (let i = rules.length - 1; i >= 0 && !list; i--) if (rules[i][0]) list = read(Math.max(2, last - 200) + i);
  }
  return list && list.length ? list : STAFF.map((s) => s.name);
}

/**
 * 担当を変える（E列）。小嶋さんから途中で引き継いだときなど。value は E列のプルダウンの選択肢のどれか（「澤井颯晟（小嶋送客）」）
 * 進捗ログに残し、契約進捗につながっていれば案件の担当も新しい主担当にする
 */
function setStaff(row, name, value) {
  value = String(value || '').trim();
  const sheet = getSheet_();
  const map = ensureAppColumns_(sheet);
  const actualRow = resolveRow_(sheet, map, row, name);
  if (staffOptions_(sheet, map, actualRow).indexOf(value) < 0) throw new Error('担当者欄のプルダウンにない値です：' + value);
  const r = sheet.getRange(actualRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const before = r[map.staff];
  const res = writeFields_(actualRow, name, { staff: value });
  getProgressLogSheet_().appendRow([
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'), name, res.row, value, '担当変更（' + before + ' → ' + value + '）', '', '', currentStaffName_() || currentEmail_(),
  ]);
  const owners = parseOwners_(value);
  if (owners) Object.assign(res, owners);
  res.staffRaw = value;
  // すでに入っている予定（初回電話・面談・内覧の予定、追客アラート）のうち、まだ先のものを新しい担当のカレンダーに移す
  try { res.moved = moveEventsToOwner_(sheet, map, res.row, name, r, before, value); } catch (e) { res.warning = '予定を新しい担当に移せませんでした：' + e.message; }
  res.insta = instaKind_(map.introType == null ? '' : r[map.introType], value);
  // 契約進捗の案件の担当も合わせる（契約は引き継いだ人が進めるため）
  if (owners && map.contract != null && r[map.contract] && CONTRACT_DB_ID) {
    try {
      const sh = contractDb_().getSheetByName('案件');
      const v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
      const idCol = v[0].indexOf('案件ID'), stCol = v[0].indexOf('担当');
      const i = v.findIndex((x, k) => k > 0 && x[idCol] === r[map.contract]);
      if (i > 0 && stCol >= 0) sh.getRange(i + 1, stCol + 1).setValue(owners.owners[0]);
      CacheService.getScriptCache().remove(CONTRACT_CACHE_KEY);
    } catch (e) { /* 契約進捗に届かなくても担当の変更は済んでいる */ }
  }
  return res;
}

/**
 * 担当が変わったとき、前の担当のカレンダーに入っている先の予定を消して、新しい担当のカレンダーに同じ日時で作り直す。
 * 通知先（alertOwner_）が変わらないとき（送客の付け方だけ変えたときなど）は何もしない。移した件数を返す
 */
function moveEventsToOwner_(sheet, map, row, name, r, before, after) {
  const oldOwner = alertOwner_(before), owner = alertOwner_(after);
  if ((oldOwner.calendar || oldOwner.email) === (owner.calendar || owner.email)) return 0; // 入れる先のカレンダーが同じなら何もしない
  const now = new Date();
  const today = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
  let moved = 0;
  // 初回電話・面談・内覧の予定（列の値「2026-10-25 14:00〜15:30」から、開始と終わりを読む）
  const refs = readPlanRefs_(r[map.planEvents]);
  Object.keys(PLAN_STAGES).forEach((key) => {
    const ps = PLAN_STAGES[key], text = r[map[ps.col]], start = parsePlan_(text);
    if (!refs[key] || !start || start < now) return;
    const m = String(text).match(/〜(\d{2}:\d{2})/);
    const end = m ? parsePlan_(String(text).slice(0, 10) + ' ' + m[1]) : null;
    deleteAlertEvent_(refs[key]);
    refs[key] = createCalendarEvent_(owner, ps.title + calendarName_(name), start, end && end > start ? end : new Date(start.getTime() + ps.minutes * 60000),
      '顧客：' + name + '\n担当：' + after).ref;
    moved++;
  });
  const fields = {};
  if (moved) fields.planEvents = JSON.stringify(refs);
  // 追客アラート（今日以降のもの）
  const date = r[map.alertDate];
  if (r[map.alertEvent] && /^\d{4}-\d{2}-\d{2}$/.test(date) && date >= today) {
    deleteAlertEvent_(r[map.alertEvent]);
    fields.alertEvent = createAlertEvent_(owner, name, date, r[map.alertNote]).ref;
    moved++;
  }
  Object.keys(fields).forEach((k) => sheet.getRange(row, map[k] + 1).setValue(fields[k]));
  return moved;
}

/** 取引種別を変える（D列）。売買から賃貸に変わった人など */
function setDealType(row, name, dealType) {
  if (DEAL_TYPES.indexOf(dealType) < 0) throw new Error('取引種別は ' + DEAL_TYPES.join('・') + ' から選んでください');
  const res = writeFields_(row, name, { dealType: dealType });
  res.dealType = dealType;
  return res;
}

/** "2026-10-25 14:00" / "2026-10-25T14:00"（後ろに「〜15:30」があってもよい）→ 開始の Date（日本時間）。読めなければ null */
function parsePlan_(text) {
  const m = String(text || '').match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})/);
  return m ? new Date(m[1] + '-' + m[2] + '-' + m[3] + 'T' + m[4] + ':' + m[5] + ':00+09:00') : null;
}

function readPlanRefs_(text) {
  try { return text ? JSON.parse(text) : {}; } catch (e) { return {}; }
}

/** 段階以外の状態ボタン（追客・時期待ち・返答待ち） */
function setStatus(row, name, status) {
  if (STATUS_BUTTONS.indexOf(status) < 0) throw new Error('不明なステータスです: ' + status);
  const res = writeFields_(row, name, { status: status });
  getProgressLogSheet_().appendRow([
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'), name, res.row, '', status, '', '', currentStaffName_() || currentEmail_(),
  ]);
  return res;
}

function getProgressLogSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(PROGRESS_LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PROGRESS_LOG_SHEET_NAME);
    sheet.appendRow(['日時', '顧客名', '行', '担当者', '段階・状態', '内覧回数', '取消', '操作した人']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 成約までの段階の分析（紹介月が FUNNEL_FROM 以降の顧客）。担当者ごとと全体。
 * reached：その段階まで進んだ人数（後の段階に記録があれば前の段階も通ったとみなす。B列が「完了」なら成約まで進んだとみなす）
 * lostAt：失客した人が最後にいた段階（0＝初回電話の前）
 */
function getFunnel() {
  const sheet = getSheet_();
  const map = ensureAppColumns_(sheet);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
  const from = Number(FUNNEL_FROM.replace('/', ''));
  const empty = () => ({ total: 0, reached: STAGES.map(() => 0), lost: 0, lostAt: new Array(STAGES.length + 1).fill(0), open: 0 });
  const groups = { all: empty() };
  STAFF.forEach((s) => { groups[s.name] = empty(); });

  values.forEach((r) => {
    const name = r[map.name].trim();
    const status = r[map.status].trim();
    const ym = Number(String(r[map.introMonth]).replace(/[^\d]/g, ''));
    if (!name || !ym || ym < from || status.indexOf('送客被り') >= 0) return;
    const owners = parseOwners_(r[map.staff]);
    if (!owners) return;
    const has = (st) => (st.key === 'viewing' ? r[map.viewing] || (parseInt(r[map.viewCount], 10) || 0) > 0 : r[map[st.col]]);
    let top = -1;
    STAGES.forEach((st, i) => { if (has(st)) top = i; });
    if (status.indexOf('完了') === 0) top = STAGES.length - 1;
    const lost = status.indexOf('失客') >= 0;
    [groups.all].concat(owners.owners.map((o) => groups[o])).forEach((g) => {
      g.total++;
      for (let i = 0; i <= top; i++) g.reached[i]++;
      if (lost) { g.lost++; g.lostAt[top + 1]++; } else if (top < STAGES.length - 1) g.open++;
    });
  });
  return { from: FUNNEL_FROM, stages: STAGES.map((st) => st.label), groups: groups };
}

/**
 * 列の中の並び順を保存する。items: [{row, name}] を上から順に。並び順の列にまとめて 1, 2, 3… と入れる
 * （1件ずつ書くと遅いので、列をまるごと読んで書き戻す。行がずれていたら顧客名で探し直す）
 */
function saveOrder(items) {
  if (!Array.isArray(items) || !items.length) return { ok: true };
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_WAIT);
  try {
    const sheet = getSheet_();
    const map = ensureAppColumns_(sheet);
    const n = sheet.getLastRow() - 1;
    const names = sheet.getRange(2, map.name + 1, n, 1).getDisplayValues().map((r) => r[0].trim());
    const orderRange = sheet.getRange(2, map.order + 1, n, 1);
    const orders = orderRange.getValues();
    items.forEach((it, i) => {
      let idx = it.row - 2;
      if (names[idx] !== it.name) idx = names.indexOf(it.name);
      if (idx >= 0) orders[idx][0] = i + 1;
    });
    orderRange.setValues(orders);
    clearCache_();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/** ステータス・メモの編集 */
function updateCustomer(row, name, fields) {
  const allowed = {};
  if ('status' in fields) allowed.status = String(fields.status);
  if ('memo' in fields) allowed.memo = String(fields.memo);
  return writeFields_(row, name, allowed);
}

function writeFields_(row, name, fields) {
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_WAIT);
  try {
    const sheet = getSheet_();
    const map = ensureAppColumns_(sheet);
    const actualRow = resolveRow_(sheet, map, row, name);
    Object.keys(fields).forEach((k) => {
      sheet.getRange(actualRow, map[k] + 1).setValue(fields[k]);
    });
    const stamp = updateStamp_();
    sheet.getRange(actualRow, map.updated + 1).setValue(stamp);
    clearCache_();
    return { row: actualRow, updated: stamp };
  } finally {
    lock.releaseLock();
  }
}

/** 行がずれていても（並べ替え・行挿入）顧客名で正しい行を探す */
function resolveRow_(sheet, map, row, name) {
  if (row >= 2 && row <= sheet.getLastRow()) {
    if (sheet.getRange(row, map.name + 1).getDisplayValue().trim() === name) return row;
  }
  const names = sheet.getRange(2, map.name + 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  const hits = [];
  names.forEach((n, i) => { if (n[0].trim() === name) hits.push(i + 2); });
  if (hits.length === 1) return hits[0];
  throw new Error('「' + name + '」の行が特定できませんでした。画面を再読み込みしてください。');
}

/** 失客ボタン：B列を「失客」にし、理由をC列のメモに追記する */
function loseCustomer(row, name, reason) {
  clearAlert(row, name);
  const sheet = getSheet_();
  const map = ensureAppColumns_(sheet);
  const actualRow = resolveRow_(sheet, map, row, name);
  const memo = sheet.getRange(actualRow, map.memo + 1).getDisplayValue();
  const line = '失客（' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'MM/dd') + '）' + (reason ? '：' + reason : '');
  const res = writeFields_(actualRow, name, { status: '失客', memo: memo ? memo + '\n' + line : line, forecast: thisMonth_() });
  res.forecast = res.boardMonth = thisMonth_();
  return res;
}

/**
 * 過去の失客として整理するボタン：B列を「失客」にするが、見込み月は空にするので今月の失客には数えず、ボードからも消える
 * （アプリを使う前にすでに失客していた顧客の片付け用）
 */
function archiveLostCustomer(row, name) {
  clearAlert(row, name);
  const sheet = getSheet_();
  const map = ensureAppColumns_(sheet);
  const actualRow = resolveRow_(sheet, map, row, name);
  const memo = sheet.getRange(actualRow, map.memo + 1).getDisplayValue();
  const line = '過去の失客として整理（' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'MM/dd') + '）';
  return writeFields_(actualRow, name, { status: '失客', memo: memo ? memo + '\n' + line : line, forecast: '' });
}

/**
 * 失客から戻す（復活）。B列を記録が残っている一番先の段階に戻し（何も無ければ「追客」）、メモに「失客から戻す（MM/dd）」と追記する
 * 見込み月はそのまま（「今月の失客」なら今月の列、「過去の失客（整理）」なら未振り分けに戻る）。進捗ログにも残す
 */
function reviveCustomer(row, name) {
  const sheet = getSheet_();
  const map = ensureAppColumns_(sheet);
  const actualRow = resolveRow_(sheet, map, row, name);
  const r = sheet.getRange(actualRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  if (resultOf_(String(r[map.status]).trim()) !== 'lost') throw new Error('失客になっていません（今：' + r[map.status] + '）');
  const count = parseInt(r[map.viewCount], 10) || 0;
  const reached = STAGES.filter((st) => st.key !== 'apply' && (st.key === 'viewing' ? count > 0 || r[map.viewing] : r[map[st.col]]));
  const top = reached[reached.length - 1];
  const status = top ? (top.key === 'viewing' ? '内覧済み（' + Math.max(count, 1) + '回）' : top.status) : '追客';
  const line = '失客から戻す（' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'MM/dd') + '）';
  const memo = r[map.memo];
  const res = writeFields_(actualRow, name, { status: status, memo: memo ? memo + '\n' + line : line });
  getProgressLogSheet_().appendRow([
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'), name, res.row, r[map.staff], '失客から戻す', '', '', currentStaffName_() || currentEmail_(),
  ]);
  res.status = status;
  res.memo = memo ? memo + '\n' + line : line;
  return res;
}

// ---------- 追客アラート ----------

/** アラート日を設定し、担当者のカレンダーに終日予定を入れる。date: "2026-10-25" */
function setAlert(row, name, date, note) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('日付を選んでください');
  const sheet = getSheet_();
  const map = ensureAppColumns_(sheet);
  const actualRow = resolveRow_(sheet, map, row, name);
  const r = sheet.getRange(actualRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  deleteAlertEvent_(r[map.alertEvent]);
  const event = createAlertEvent_(alertOwner_(r[map.staff]), name, date, note);
  const res = writeFields_(actualRow, name, { alertDate: date, alertNote: note || '', alertEvent: event.ref });
  res.calendar = event.where;
  return res;
}

/** アラート解除（対応済み） */
function clearAlert(row, name) {
  const sheet = getSheet_();
  const map = ensureAppColumns_(sheet);
  const actualRow = resolveRow_(sheet, map, row, name);
  const ref = sheet.getRange(actualRow, map.alertEvent + 1).getDisplayValue();
  const date = sheet.getRange(actualRow, map.alertDate + 1).getDisplayValue();
  if (!ref && !date) return { row: actualRow };
  deleteAlertEvent_(ref);
  return writeFields_(actualRow, name, { alertDate: '', alertNote: '', alertEvent: '' });
}

/** 通知先：主担当が8名なら主担当、そうでなければ送客元 */
function alertOwner_(rawStaff) {
  const o = parseOwners_(rawStaff);
  const name = o ? o.owners[0] : '';
  const staff = STAFF.find((s) => s.name === name);
  return { name: name, email: (staff && staff.email) || Session.getEffectiveUser().getEmail(), calendar: (staff && staff.calendar) || '' };
}

/**
 * 担当者のカレンダーに直接予定を作る（担当者が小嶋さんに「予定の変更」権限でカレンダーを共有している場合）。
 * 共有されていなければ、小嶋さんのカレンダーに予定を作って担当者を招待する。
 */
function createAlertEvent_(owner, name, date, note) {
  const [y, m, d] = date.split('-').map(Number);
  return createCalendarEvent_(owner, '【追客】' + calendarName_(name), new Date(y, m - 1, d), null, note);
}

/** end が null なら終日の予定 */
function createCalendarEvent_(owner, title, start, end, note) {
  const desc = (note ? note + '\n\n' : '') + '見込みボード：' + ScriptApp.getService().getUrl();
  const make = (cal, opts) => (end ? cal.createEvent(title, start, end, opts) : cal.createAllDayEvent(title, start, opts));
  // 専用カレンダー（小嶋さん・諸本さんのゲスト招待用）があれば、まずそこに入れる
  if (owner.calendar) {
    try {
      const gc = CalendarApp.getCalendarById(owner.calendar);
      if (gc) {
        const ev = make(gc, { description: desc });
        paintEvent_(ev, owner.name);
        return { ref: gc.getId() + '|' + ev.getId(), where: owner.name + 'さんの専用カレンダー' };
      }
    } catch (e) { /* 書けなければ下のふつうの入れ方にする */ }
  }
  let cal = null;
  try { cal = CalendarApp.getCalendarById(owner.email); } catch (e) { cal = null; }
  if (cal && owner.email !== Session.getEffectiveUser().getEmail()) {
    try {
      const ev = make(cal, { description: desc });
      paintEvent_(ev, owner.name);
      return { ref: cal.getId() + '|' + ev.getId(), where: owner.name + 'さんのカレンダー' };
    } catch (e) { /* 閲覧のみの共有なら招待にする */ }
  }
  const own = CalendarApp.getDefaultCalendar();
  const isSelf = owner.email === Session.getEffectiveUser().getEmail();
  const ev = make(own, isSelf ? { description: desc } : { description: desc, guests: owner.email, sendInvites: false });
  paintEvent_(ev, owner.name);
  return { ref: own.getId() + '|' + ev.getId(), where: isSelf ? '小嶋さんのカレンダー' : owner.name + 'さんを招待' };
}

/** 予定に担当の色（STAFF の event）を付ける。色が決まっていない人・失敗したときはそのまま */
function paintEvent_(ev, staffName) {
  const s = STAFF.find((x) => x.name === staffName);
  if (!s || !s.event || !CalendarApp.EventColor[s.event]) return;
  try { ev.setColor(CalendarApp.EventColor[s.event]); } catch (e) { /* 色が付けられなくても予定はある */ }
}

function deleteAlertEvent_(ref) {
  if (!ref) return;
  const [calId, eventId] = ref.split('|');
  try {
    const ev = CalendarApp.getCalendarById(calId).getEventById(eventId);
    if (ev) ev.deleteEvent();
  } catch (e) { /* 手動で消されていれば何もしない */ }
}

/** 毎朝8時：今日の分と期限切れの分 */
function sendMorningAlerts() {
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  sendAlerts_((d) => d <= today, (d) => (d < today ? '期限切れ ' : '') + d.slice(5).replace('-', '/'), '今日の追客');
}

/** 毎日18時：明日の分 */
function sendEveningAlerts() {
  const tomorrow = Utilities.formatDate(new Date(Date.now() + 864e5), 'Asia/Tokyo', 'yyyy-MM-dd');
  sendAlerts_((d) => d === tomorrow, (d) => d.slice(5).replace('-', '/'), '明日の追客');
}

function sendAlerts_(match, label, subject) {
  const sheet = getSheet_();
  const map = ensureAppColumns_(sheet);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
  const byEmail = {};
  values.forEach((r) => {
    const d = r[map.alertDate];
    if (!d || !match(d) || isClosed_(r[map.status].trim())) return;
    const owner = alertOwner_(r[map.staff]);
    (byEmail[owner.email] = byEmail[owner.email] || []).push(
      '・' + r[map.name] + '（' + label(d) + '）' + (r[map.alertNote] ? '\n　' + r[map.alertNote] : '') + (r[map.status] ? '\n　ステータス：' + r[map.status] : ''));
  });
  const url = ScriptApp.getService().getUrl();
  Object.keys(byEmail).forEach((email) => {
    const lines = byEmail[email];
    MailApp.sendEmail(email, '【見込みボード】' + subject + ' ' + lines.length + '件', lines.join('\n\n') + '\n\n見込みボード：' + url);
  });
}

function installAlertTriggers_() {
  ScriptApp.getProjectTriggers()
    .filter((t) => ['sendMorningAlerts', 'sendEveningAlerts'].indexOf(t.getHandlerFunction()) >= 0)
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendMorningAlerts').timeBased().everyDays(1).atHour(8).inTimezone('Asia/Tokyo').create();
  ScriptApp.newTrigger('sendEveningAlerts').timeBased().everyDays(1).atHour(18).inTimezone('Asia/Tokyo').create();
}

// ---------- 売上 ----------

/**
 * 担当者ごとの5期目の目標と実績（月別）を返す。
 * 目標：上段の「年間売上目標」ブロックで、ラベルが担当者名で始まる行（↪売買目標の行は除く）
 *       年間目標はラベルの【】内の金額（例：【2,000万🔥】）。月別はその行の4月〜3月の値
 *       売買の目標は件数で、担当者の行のすぐ下にある「↪売買目標」行の件数（【3件】のように【】があればそちらを優先）
 * 実績：下段の「年間売上表」ブロックで、ラベルが担当者名だけの行。無ければ「担当者名 ※元売上」の行
 * 申込件数：同シートさらに下の「年間申込件数表」ブロックで、ラベルが担当者名（または「担当者名 ※元申込件数」）の行
 * 売買の売上・申込件数（平均単価を賃貸だけで出すために引く）：SALE_ACTUAL_LABELS / SALE_APPS_LABELS の行
 */
function getSales() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(SALES_CACHE_KEY);
  if (hit) return JSON.parse(hit);

  const sheet = SpreadsheetApp.openById(SALES_SPREADSHEET_ID).getSheetByName(SALES_SHEET_NAME);
  const rows = sheet.getRange(1, 2, Math.min(sheet.getLastRow(), 300), 14).getDisplayValues(); // B〜O列
  const norm = (v) => String(v || '').replace(/[\s　]/g, '');
  const yen = (v) => Number(String(v || '').replace(/[^\d-]/g, '')) || 0;
  const targetStart = rows.findIndex((r) => norm(r[0]).indexOf('年間売上目標') === 0);
  const actualStart = rows.findIndex((r) => norm(r[0]) === '年間売上表');
  const appsStart = rows.findIndex((r) => norm(r[0]) === '年間申込件数表');
  if (targetStart < 0 || actualStart < 0 || appsStart < 0) {
    throw new Error('売上管理表の「年間売上目標」「年間売上表」「年間申込件数表」の行が見つかりません');
  }
  const actualBlock = rows.slice(actualStart + 1, appsStart);
  const appsBlock = rows.slice(appsStart + 1);
  const months = (r) => r.slice(1, 13).map(yen);
  // labels のどれかに一致する最初の行を、ラベルの組ごとに月別で合計する（見つからない組は無視）
  const sumRows = (block, groups) => {
    const out = new Array(12).fill(0);
    groups.forEach((labels) => {
      const r = block.find((row) => labels.indexOf(norm(row[0])) >= 0);
      if (r) months(r).forEach((v, i) => { out[i] += v; });
    });
    return out;
  };

  const staff = STAFF.map((s) => {
    const n = s.name;
    const targetBlock = rows.slice(targetStart + 1, actualStart);
    const tIdx = targetBlock.findIndex((r) => norm(r[0]).indexOf(n) === 0);
    const target = tIdx >= 0 ? targetBlock[tIdx] : null;
    let saleTargetCount = 0;
    for (let k = tIdx + 1; tIdx >= 0 && k < targetBlock.length && norm(targetBlock[k][0]).indexOf('↪') === 0; k++) {
      const l = norm(targetBlock[k][0]);
      if (l.indexOf('売買目標') < 0) continue;
      const hit = l.match(/【(\d+)件/) || l.match(/(\d+)件/);
      if (hit) saleTargetCount = Number(hit[1]);
      break;
    }
    const actualRows = actualBlock.filter((r) => norm(r[0]).indexOf(n) === 0);
    const combined = SALES_COMBINED[n];
    const actual = combined
      ? combined.actual.map((l) => actualBlock.find((r) => norm(r[0]) === l)).filter(Boolean)
      : [actualRows.find((r) => norm(r[0]) === n) || actualRows.find((r) => norm(r[0]).indexOf('元売上') >= 0)].filter(Boolean);
    const tMonths = target ? months(target) : null;
    const aMonths = new Array(12).fill(0);
    actual.forEach((r) => months(r).forEach((v, i) => { aMonths[i] += v; }));
    const saleMonths = sumRows(actualBlock, [SALE_ACTUAL_LABELS(n)].concat((combined ? combined.saleActual : []).map((l) => [l])));
    const appMonths = sumRows(appsBlock, combined ? combined.apps.map((l) => [l]) : [[n, n + '※元申込件数']]);
    const saleAppMonths = sumRows(appsBlock, [SALE_APPS_LABELS(n)].concat((combined ? combined.saleApps : []).map((l) => [l])));
    const label = target ? String(target[0]) : '';
    const m = label.match(/【[^】]*?([\d,]+(?:\.\d+)?)万/);
    return {
      name: n,
      targetLabel: label.trim(),
      actualLabel: actual.map((r) => String(r[0]).trim()).join(' ＋ '),
      targetMonths: tMonths,
      actualMonths: aMonths,
      targetTotal: m ? Math.round(Number(m[1].replace(/,/g, '')) * 10000) : 0,
      actualTotal: aMonths.reduce((a, b) => a + b, 0),
      appMonths: appMonths,
      saleTotal: saleMonths.reduce((a, b) => a + b, 0),
      saleAppMonths: saleAppMonths,
      saleTargetCount: saleTargetCount,
    };
  });

  // 会社の売買目標（件数）：「売買総合計【30件】」の行。見込みボードの「売買」の欄のゲージに使う
  let saleTargetTotal = typeof SALE_TARGET_DEFAULT === 'number' ? SALE_TARGET_DEFAULT : 0;
  rows.forEach((r) => { const m = norm(r[0]).match(/^売買総合計.*?【(\d+)件/); if (m) saleTargetTotal = Number(m[1]); });
  // 会社全体の売買の件数・売上：担当者ごとの行を足すと、飛び込み・HPの売買が担当者の売買申込数と二重に入るので、
  // 会社の行（「★売買総申込件数」「（分析）全体売買売上」）をそのまま使う。見つからなければ null（画面は担当者の合計にする）
  const companyRow = (block, head) => { const r = block.find((row) => norm(row[0]).indexOf(head) === 0); return r ? months(r) : null; };
  const data = {
    saleTargetTotal: saleTargetTotal,
    companySaleAppMonths: companyRow(appsBlock, '★売買総申込件数'),
    companySaleMonths: companyRow(actualBlock, '（分析）全体売買売上'),
    periodStart: SALES_PERIOD_START, periodEnd: SALES_PERIOD_END, staff: staff, loadedAt: new Date().toISOString(),
    month: thisMonth_(), monthTargets: readMonthTargets_(thisMonth_()),
  };
  cache.put(SALES_CACHE_KEY, JSON.stringify(data), 1800); // 売上管理表はあまり変わらないので30分（今月の目標を入れたときは捨てる）
  return data;
}

// ---------- 今月の目標（手入力） ----------

function getMonthTargetSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(MONTH_TARGET_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MONTH_TARGET_SHEET_NAME);
    sheet.appendRow(['月', '担当者', '目標金額', '更新者', '更新日時']);
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('@');
  }
  return sheet;
}

/** その月の目標 { 担当者名: 金額 } */
function readMonthTargets_(ym) {
  const sheet = getMonthTargetSheet_();
  const last = sheet.getLastRow();
  const out = {};
  if (last < 2) return out;
  sheet.getRange(2, 1, last - 1, 3).getValues().forEach((r) => {
    if (String(r[0]) === ym && r[1]) out[r[1]] = Number(r[2]) || 0;
  });
  return out;
}

/** 担当者の今月の目標を入れる（変更も同じ行を書き換え）。amount は円 */
function setMonthTarget(staffName, amount) {
  if (!STAFF.some((s) => s.name === staffName)) throw new Error('担当者が見つかりません: ' + staffName);
  amount = Math.round(Number(amount));
  if (!(amount >= 0) || amount > 1e9) throw new Error('目標金額を正しく入れてください');
  const ym = thisMonth_();
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_WAIT);
  try {
    const sheet = getMonthTargetSheet_();
    const last = sheet.getLastRow();
    const rows = last < 2 ? [] : sheet.getRange(2, 1, last - 1, 2).getValues();
    const i = rows.findIndex((r) => String(r[0]) === ym && r[1] === staffName);
    const values = [[ym, staffName, amount, currentStaffName_() || currentEmail_(), Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')]];
    sheet.getRange(i >= 0 ? i + 2 : last + 1, 1, 1, 5).setValues(values);
    CacheService.getScriptCache().remove(SALES_CACHE_KEY);
    return { month: ym, amount: amount };
  } finally {
    lock.releaseLock();
  }
}

// ---------- 議事録 ----------

function getNotes(row, name) {
  const sheet = getSheet_();
  const map = ensureAppColumns_(sheet);
  const actualRow = resolveRow_(sheet, map, row, name);
  const folderId = sheet.getRange(actualRow, map.folder + 1).getDisplayValue();
  if (!folderId) return { folderUrl: '', notes: [] };
  const notesSheet = getNotesSheet_();
  const last = notesSheet.getLastRow();
  const rows = last < 2 ? [] : notesSheet.getRange(2, 1, last - 1, 8).getDisplayValues();
  const notes = rows
    .filter((r) => r[0] === folderId)
    .map((r) => ({ date: r[1], type: r[2], title: r[3], url: r[4], memo: r[6], by: r[7] }))
    .reverse();
  return { folderUrl: 'https://drive.google.com/drive/folders/' + folderId, notes: notes };
}

/** Google ドキュメント等のリンクを追加 */
function addNoteLink(row, name, url, title, memo) {
  if (!/^https:\/\/(docs|drive)\.google\.com\//.test(url)) throw new Error('Google ドキュメント／ドライブのリンクを入力してください');
  const folderId = ensureCustomerFolder_(row, name);
  let fileId = '';
  const m = url.match(/\/d\/([\w-]{20,})/) || url.match(/[?&]id=([\w-]{20,})/);
  if (m) {
    fileId = m[1];
    try {
      if (!title) title = DriveApp.getFileById(fileId).getName();
      DriveApp.createShortcut(fileId).moveTo(DriveApp.getFolderById(folderId));
    } catch (e) {
      // 小嶋さんに共有されていないファイルはショートカットを作れない。リンクだけ残す。
    }
  }
  appendNote_(folderId, 'リンク', title || 'Google メモ', url, fileId, memo);
  return getNotes(row, name);
}

/** PDF 等のアップロード（base64） */
function uploadNoteFile(row, name, fileName, mimeType, base64, memo) {
  const folderId = ensureCustomerFolder_(row, name);
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, fileName);
  const file = DriveApp.getFolderById(folderId).createFile(blob);
  appendNote_(folderId, mimeType === 'application/pdf' ? 'PDF' : 'ファイル', fileName, file.getUrl(), file.getId(), memo);
  return getNotes(row, name);
}

/** テキストだけの一言メモ */
function addNoteText(row, name, memo) {
  const folderId = ensureCustomerFolder_(row, name);
  appendNote_(folderId, 'メモ', '', '', '', memo);
  return getNotes(row, name);
}

function ensureCustomerFolder_(row, name) {
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_WAIT);
  try {
    const sheet = getSheet_();
    const map = ensureAppColumns_(sheet);
    const actualRow = resolveRow_(sheet, map, row, name);
    const cell = sheet.getRange(actualRow, map.folder + 1);
    const existing = cell.getDisplayValue();
    if (existing) return existing;
    const folder = getNotesRoot_().createFolder(name);
    cell.setValue(folder.getId());
    clearCache_();
    return folder.getId();
  } finally {
    lock.releaseLock();
  }
}

function appendNote_(folderId, type, title, url, fileId, memo) {
  getNotesSheet_().appendRow([
    folderId,
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'),
    type, title, url, fileId, memo || '', currentStaffName_() || currentEmail_(),
  ]);
  clearCache_();
}

function countNotes_() {
  const sheet = getNotesSheet_();
  const last = sheet.getLastRow();
  const counts = {};
  if (last < 2) return counts;
  sheet.getRange(2, 1, last - 1, 1).getValues().forEach((r) => { counts[r[0]] = (counts[r[0]] || 0) + 1; });
  return counts;
}

// ---------- シート・列の準備 ----------

function getSheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
}

function getNotesSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(NOTES_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(NOTES_SHEET_NAME);
    sheet.appendRow(['議事録フォルダ', '登録日時', '種類', 'タイトル', 'URL', 'ファイルID', 'メモ', '登録者']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getNotesRoot_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('NOTES_ROOT_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* 削除されていたら作り直す */ }
  }
  const folder = DriveApp.createFolder(NOTES_ROOT_FOLDER_NAME);
  props.setProperty('NOTES_ROOT_ID', folder.getId());
  return folder;
}

/** ヘッダー名から列番号(0始まり)の対応表を作る。アプリ用の列が無ければ右端に追加する */
function ensureAppColumns_(sheet) {
  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map((h) => h.trim());
  const map = {};
  Object.keys(COLS).forEach((k) => {
    const idx = header.indexOf(COLS[k].trim());
    map[k] = idx >= 0 ? idx : null;
  });
  if (map.name == null || map.status == null) throw new Error('「顧客名」「進捗ステータス」列が見つかりません');

  let width = lastCol;
  Object.keys(APP_COLS).forEach((k) => {
    let idx = header.indexOf(APP_COLS[k]);
    if (idx < 0) {
      if (sheet.getMaxColumns() <= width) sheet.insertColumnsAfter(sheet.getMaxColumns(), 1);
      sheet.getRange(1, width + 1).setValue(APP_COLS[k]);
      // 右端に列を足すと左隣（広告料 相殺or後請求）の入力規則がコピーされるので、外してから使う
      sheet.getRange(1, width + 1, sheet.getMaxRows(), 1).setDataValidation(null).setNumberFormat('@');
      header[width] = APP_COLS[k];
      idx = width;
      width++;
    }
    map[k] = idx;
  });
  return map;
}

// ---------- ユーティリティ ----------

function currentEmail_() {
  return (Session.getActiveUser().getEmail() || '').toLowerCase();
}

function currentStaffName_() {
  const email = currentEmail_();
  const hit = STAFF.find((s) => s.email && s.email.toLowerCase() === email);
  return hit ? hit.name : '';
}

function updateStamp_() {
  const who = currentStaffName_() || currentEmail_().split('@')[0] || '不明';
  return who.replace(/^(..).*$/, '$1') + ' ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'MM/dd HH:mm');
}

// CacheService は1キー100KBまでなので分割して保存する
function writeCache_(data) { writeCacheKey_(CACHE_KEY, data, CACHE_TTL); }
function readCache_() { return readCacheKey_(CACHE_KEY); }

function writeCacheKey_(key, data, ttl) {
  const json = JSON.stringify(data);
  const size = 90000;
  const parts = {};
  const n = Math.ceil(json.length / size);
  for (let i = 0; i < n; i++) parts[key + '_' + i] = json.slice(i * size, (i + 1) * size);
  parts[key] = String(n);
  try { CacheService.getScriptCache().putAll(parts, ttl); } catch (e) { /* キャッシュできなくても動く */ }
}

function readCacheKey_(key) {
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get(key));
  if (!n) return null;
  const keys = [];
  for (let i = 0; i < n; i++) keys.push(key + '_' + i);
  const got = cache.getAll(keys);
  if (keys.some((k) => got[k] == null)) return null;
  return JSON.parse(keys.map((k) => got[k]).join(''));
}

function clearCache_() {
  CacheService.getScriptCache().remove(CACHE_KEY);
}

// ---------- ホーム画面（アプリの入口） ----------
// ?page=home で開く。見込みボード・成績ダッシュボード・契約進捗・マニュアル・営業資料を大きいカードで並べるだけ（データは読まないので速い）
// カードを足すときは homeApps_ に1行足す。url が空のものは出さない
function homeApps_() {
  const app = ScriptApp.getService().getUrl();
  return [
    { title: '見込みボード', desc: '見込み月の管理・成約・失客', tag: 'BOARD', icon: 'board', url: app },
    { title: '成績ダッシュボード', desc: '売上・紹介・縦掘り・昇格', tag: 'DASHBOARD', icon: 'chart', url: app + '?page=analysis' },
    { title: '契約進捗', desc: '契約のチェック・鍵渡し', tag: 'CONTRACT', icon: 'key', url: CONTRACT_DB_ID ? CONTRACT_APP_URL : '' },
    { title: 'マニュアル', desc: '社内ルール・営業マニュアル', tag: 'MANUAL', icon: 'book', url: MANUAL_APP_URL },
    { title: '営業資料', desc: '提案書・物件資料・チラシなど', tag: 'DOCS', icon: 'folder', url: SALES_DOCS_FOLDER_ID ? app + '?page=docs' : '' },
  ].filter((a) => a.url);
}
const HOME_ICONS = {
  board: '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M9 4v16M15 4v16"/>',
  chart: '<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/>',
  key: '<circle cx="8" cy="15" r="4"/><path d="M11 12l9-9M17 6l3 3M14 9l2 2"/>',
  book: '<path d="M2 4h6a4 4 0 0 1 4 4v13a3 3 0 0 0-3-3H2z"/><path d="M22 4h-6a4 4 0 0 0-4 4v13a3 3 0 0 1 3-3h7z"/>',
  folder: '<path d="M3 6a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><path d="M3 10h18"/>',
};
// ホーム画面・営業資料の画面で共通の見た目（マニュアルのアプリに合わせた、ベージュの背景・黒い帯・白いカード）
const PORTAL_CSS = ':root{--bg:#f5f4f0;--card:#fff;--fg:#111;--sub:#6b6b6b;--line:#e4e2dc;--head:#0d0d0d;--headfg:#fff}' +
  '@media (prefers-color-scheme:dark){:root{--bg:#121212;--card:#1c1c1c;--fg:#f2f2f2;--sub:#a0a0a0;--line:#2e2e2e;--head:#000;--headfg:#fff}}' +
  '*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--fg);font-family:"Hiragino Sans","Noto Sans JP",system-ui,sans-serif}' +
  '.head{background:var(--head);color:var(--headfg);padding:28px 20px 34px;border-radius:0 0 24px 24px}' +
  '.head small{display:block;font-size:12px;letter-spacing:.2em;opacity:.6;font-weight:700}' +
  '.head h1{margin:6px 0 4px;font-size:26px;letter-spacing:.06em}.head p{margin:0;font-size:13px;opacity:.7}.in{max-width:1100px;margin:0 auto}' +
  '.grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;padding:16px;max-width:1100px;margin:0 auto}' +
  '@media (min-width:900px){.grid{grid-template-columns:repeat(4,minmax(0,1fr));gap:18px;padding:24px}}' +
  '.card{display:flex;flex-direction:column;gap:8px;background:var(--card);border:1px solid var(--line);border-radius:18px;padding:20px 18px 16px;color:inherit;text-decoration:none;min-height:180px;transition:transform .12s,box-shadow .12s}' +
  '.card:hover{transform:translateY(-2px);box-shadow:0 8px 24px rgba(0,0,0,.08)}.card:active{transform:scale(.98)}' +
  '.card svg{width:34px;height:34px}.card b{font-size:17px;letter-spacing:.04em;line-height:1.4}' +
  '.desc{font-size:12px;color:var(--sub);line-height:1.5;flex:1}' +
  '.tag{align-self:flex-start;font-size:11px;font-weight:700;letter-spacing:.12em;color:var(--sub);border:1px solid var(--line);border-radius:999px;padding:3px 12px}' +
  '.tip{max-width:1100px;margin:0 auto;padding:0 16px 32px;font-size:12px;color:var(--sub);line-height:1.7}' +
  '.tip summary{cursor:pointer}' +
  '@media (min-width:900px){.head{padding:36px 0 40px}.head .in{padding:0 24px}.tip{padding:0 24px 32px}}';
const escHtml_ = (v) => String(v).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
function portalOutput_(title, css, body) {
  return HtmlService.createHtmlOutput('<!doctype html><html lang="ja"><head><meta charset="utf-8"><base target="_top"><style>' +
      PORTAL_CSS + css + '</style></head><body>' + body + '</body></html>')
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function homePage_() {
  const h = Number(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'H'));
  const hello = h < 11 ? 'おはようございます' : h < 18 ? 'こんにちは' : 'おつかれさまです';
  const cards = homeApps_().map((a) =>
    '<a class="card" href="' + escHtml_(a.url) + '" target="_top">' +
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + HOME_ICONS[a.icon] + '</svg>' +
      '<b>' + escHtml_(a.title) + '</b><span class="desc">' + escHtml_(a.desc) + '</span><span class="tag">' + escHtml_(a.tag) + '</span></a>').join('');
  return portalOutput_('Monochro ホーム', '',
    '<div class="head"><div class="in"><small>MONOCHRO</small><h1>ホーム</h1><p>' + hello + '。使うアプリを選んでください</p></div></div>' +
    '<div class="grid">' + cards + '</div>' +
    '<details class="tip"><summary>スマホのホーム画面に置くには</summary>' +
    'iPhone：Safari でこのページを開き、共有ボタン →「ホーム画面に追加」<br>Android：Chrome でこのページを開き、右上の「︙」→「ホーム画面に追加」<br>パソコン：このページをブックマークバーに入れておくと便利です</details>');
}

// ---------- 営業資料（ドライブのフォルダをカードで見せる） ----------
// ?page=docs で開く。SALES_DOCS_FOLDER_ID のフォルダの中のフォルダが、そのまま上のカテゴリ（タブ）になる。
// 資料を増やすときはドライブのそのフォルダに入れるだけ（アプリは10分ごと、または↻で読み直す）。
//   ・フォルダ名・ファイル名の頭の「01_」「1.」などの番号は並び順に使い、画面では消す
//   ・ファイル名の頭に「★」を付けると、そのカテゴリの一番上に「おすすめ」として出る
//   ・ドライブでファイルを右クリック →「ファイル情報 → 詳細」の「説明」に書いた文が、カードの説明になる
//   ・ショートカットも置ける（元のファイルを開く）。カテゴリの中のフォルダは1段まで読む（カードに小さくフォルダ名）
//   ・一番上のフォルダに直接置いたファイルは「ツール・その他」
//   ・名前に「※社外秘」があれば消して🔒の印にする。名前に「触らない」があるフォルダ・ファイル、頭が「_」のものは出さない
const SALES_DOCS_FOLDER_ID = '1uR96NKqWaKIiIuVXIW2hbXCwkDKRpWQ5'; // 共有ドライブの「営業資料（アプリ用）」
const DOCS_CACHE_KEY = 'docs_v1';
function getSalesDocs(fresh) {
  if (!SALES_DOCS_FOLDER_ID) return { categories: [], error: '営業資料のフォルダがまだ決まっていません（Code.gs の SALES_DOCS_FOLDER_ID）' };
  if (!fresh) { const hit = readCacheKey_(DOCS_CACHE_KEY); if (hit) return hit; }
  const root = DriveApp.getFolderById(SALES_DOCS_FOLDER_ID);
  const SECRET = /[\s　]*※\s*社外秘[\s　]*/;
  const hidden = (n) => /触らない|^\s*_/.test(n);
  const strip = (n) => String(n).replace(SECRET, ' ').replace(/^\s*\d+\s*[._\-\s、）)]\s*/, '').trim();
  const now = Date.now();
  const info = (f, sub) => {
    let id = f.getId(), mime = f.getMimeType();
    if (mime === MimeType.SHORTCUT) { id = f.getTargetId(); mime = f.getTargetMimeType(); }
    const raw = f.getName(), pin = /^\s*[★☆]/.test(raw), updated = f.getLastUpdated(), secret = SECRET.test(raw);
    return {
      secret: secret, name: strip(raw.replace(/^\s*[★☆]\s*/, '')).replace(/\.(pdf|png|jpe?g|gif|heic|docx?|xlsx?|pptx?|mp4|mov)$/i, ''), sortKey: raw.replace(/^\s*[★☆]\s*/, ''), pin: pin, sub: sub || '',
      url: 'https://drive.google.com/open?id=' + id, mime: mime, desc: f.getDescription() || '',
      updated: Utilities.formatDate(updated, 'Asia/Tokyo', 'yyyy/MM/dd'), isNew: now - updated.getTime() < 14 * 864e5,
    };
  };
  const filesIn = (folder, sub, out, depth) => {
    const it = folder.getFiles();
    while (it.hasNext()) { const f = it.next(); if (!hidden(f.getName())) out.push(info(f, sub)); }
    if (depth > 0) { const fs = folder.getFolders(); while (fs.hasNext()) { const c = fs.next(); if (!hidden(c.getName())) filesIn(c, strip(c.getName()), out, depth - 1); } }
    return out;
  };
  const sortFiles = (a) => a.sort((x, y) => (y.pin - x.pin) || x.sortKey.localeCompare(y.sortKey, 'ja', { numeric: true }));
  const cats = [];
  const fs = root.getFolders();
  while (fs.hasNext()) {
    const c = fs.next(), n = c.getName();
    if (!hidden(n)) cats.push({ key: n.replace(/^\s*[★☆]\s*/, ''), name: strip(n), secret: SECRET.test(n), files: sortFiles(filesIn(c, '', [], 1)) });
  }
  cats.sort((a, b) => a.key.localeCompare(b.key, 'ja', { numeric: true }));
  const loose = sortFiles(filesIn(root, '', [], 0));
  if (loose.length) cats.push({ key: '\uffff', name: 'ツール・その他', secret: false, files: loose });
  const data = { categories: cats.map((c) => ({ name: c.name, secret: c.secret, files: c.files })), folderUrl: root.getUrl(), folderName: root.getName(),
    loadedAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'MM/dd HH:mm') };
  writeCacheKey_(DOCS_CACHE_KEY, data, 600);
  return data;
}
function docsPage_() {
  const home = ScriptApp.getService().getUrl() + '?page=home';
  const css =
    '.head .row{display:flex;align-items:center;gap:10px}.head .row h1{flex:1}' +
    '.hb{display:inline-flex;align-items:center;justify-content:center;min-width:36px;height:36px;padding:0 10px;border-radius:10px;border:1px solid rgba(255,255,255,.25);background:transparent;color:#fff;text-decoration:none;font-size:15px;cursor:pointer}' +
    '.search{display:block;width:100%;margin-top:16px;padding:14px 18px;border-radius:14px;border:0;font-size:16px;background:#fff;color:#111}' +
    '.chips{display:flex;gap:8px;overflow-x:auto;padding:16px 16px 0;max-width:1100px;margin:0 auto;scrollbar-width:none}' +
    '.chip{flex:none;border:1px solid var(--line);background:var(--card);color:var(--fg);border-radius:999px;padding:8px 14px;font-size:13px;cursor:pointer;white-space:nowrap}' +
    '.chip[aria-pressed="true"]{background:var(--fg);color:var(--bg);border-color:var(--fg);font-weight:700}.chip i{font-style:normal;opacity:.6;margin-left:4px}' +
    '.sec{max-width:1100px;margin:0 auto;padding:18px 16px 0;font-size:13px;font-weight:700;letter-spacing:.08em;color:var(--sub)}' +
    '.grid{grid-template-columns:repeat(auto-fill,minmax(160px,1fr));padding-top:10px}' +
    '.card{min-height:150px;padding:16px}.card b{font-size:15px;word-break:break-all}' +
    '.ft{display:inline-flex;align-items:center;gap:6px;font-size:11px;font-weight:700;letter-spacing:.06em}.ft::before{content:"";width:10px;height:10px;border-radius:3px;background:var(--c)}' +
    '.meta{display:flex;flex-wrap:wrap;gap:6px;align-items:center;font-size:11px;color:var(--sub)}' +
    '.new{background:#e5484d;color:#fff;border-radius:6px;padding:1px 6px;font-weight:700;font-size:10px}.pin{color:#c58a00}' +
    '.secret{margin-left:8px;font-size:10px;letter-spacing:.06em;border:1px solid var(--line);border-radius:6px;padding:1px 6px;font-weight:700}' +
    '.empty{max-width:1100px;margin:0 auto;padding:40px 16px;text-align:center;color:var(--sub);font-size:14px;line-height:1.8}' +
    '.empty a{color:inherit}.foot{max-width:1100px;margin:0 auto;padding:24px 16px 40px;font-size:12px;color:var(--sub)}.foot a{color:inherit}' +
    '@media (min-width:900px){.grid{grid-template-columns:repeat(auto-fill,minmax(230px,1fr));padding:10px 24px 0}.chips,.sec,.foot{padding-left:24px;padding-right:24px}}';
  const body =
    '<div class="head"><div class="in"><small>MONOCHRO</small>' +
    '<div class="row"><a class="hb" href="' + escHtml_(home) + '" target="_top" title="ホーム" aria-label="ホーム">🏠</a><h1>営業資料</h1>' +
    '<button class="hb" id="reload" title="ドライブから読み直す" aria-label="読み直す">↻</button></div>' +
    '<p id="sub">読み込み中…</p>' +
    '<input class="search" id="q" type="search" placeholder="資料を検索…（例：トーク、決済金）"></div></div>' +
    '<div class="chips" id="chips"></div><div id="list"></div><div class="foot" id="foot"></div>' +
    '<script>' + DOCS_CLIENT_JS + '</script>';
  return portalOutput_('営業資料', css, body);
}
// 営業資料の画面の動き（前回の中身を端末に残して、開いたらすぐ出してから最新を読む）
const DOCS_CLIENT_JS = String.raw`
const TYPES=[["pdf","PDF","#e5484d"],["presentation","スライド","#f59e0b"],["document","ドキュメント","#3b82f6"],["spreadsheet","シート","#16a34a"],["image","画像","#8b5cf6"],["video","動画","#ec4899"],["form","フォーム","#7c3aed"]];
const typeOf=m=>{for(const t of TYPES)if(m.indexOf(t[0])>=0)return t;return ["","ファイル","#9ca3af"]};
const esc=v=>String(v==null?"":v).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"})[c]);
let D=null,cat=0;
try{D=JSON.parse(localStorage.getItem("docs.v1")||"null");cat=Number(localStorage.getItem("docs.cat")||0)}catch(e){}
function card(f){const t=typeOf(f.mime);return '<a class="card" href="'+esc(f.url)+'" target="_blank" rel="noopener"><span class="ft" style="--c:'+t[2]+'">'+t[1]+'</span>'+
 '<b>'+(f.pin?'<span class="pin">★ </span>':'')+esc(f.name)+'</b><span class="desc">'+esc(f.desc)+'</span>'+
 '<span class="meta">'+(f.isNew?'<span class="new">NEW</span>':'')+(f.secret?'<span>🔒 社外秘</span>':'')+(f.sub?'<span>📁 '+esc(f.sub)+'</span>':'')+'<span>更新 '+esc(f.updated)+'</span></span></a>'}
function render(){
 const sub=document.getElementById("sub"),chips=document.getElementById("chips"),list=document.getElementById("list"),foot=document.getElementById("foot");
 if(!D){return}
 if(D.error){sub.textContent="準備中";list.innerHTML='<div class="empty">'+esc(D.error)+'</div>';chips.innerHTML="";return}
 const cats=D.categories,total=cats.reduce((a,c)=>a+c.files.length,0);
 if(cat>cats.length)cat=0;
 sub.textContent="ドライブの「"+D.folderName+"」フォルダの資料 "+total+"件";
 const q=document.getElementById("q").value.trim().toLowerCase();
 chips.innerHTML=['<button class="chip" data-i="0" aria-pressed="'+(cat===0)+'">すべて<i>'+total+'</i></button>'].concat(cats.map((c,i)=>'<button class="chip" data-i="'+(i+1)+'" aria-pressed="'+(cat===i+1)+'">'+(c.secret?'🔒 ':'')+esc(c.name)+'<i>'+c.files.length+'</i></button>')).join("");
 const hit=f=>!q||(f.name+" "+f.desc+" "+f.sub).toLowerCase().indexOf(q)>=0;
 const show=cat===0?cats:[cats[cat-1]];let html="",n=0;
 show.forEach(c=>{const fs=c.files.filter(hit);n+=fs.length;if(fs.length)html+='<div class="sec">'+(c.secret?'🔒 ':'')+esc(c.name)+(c.secret?'<span class="secret">社外秘</span>':'')+'</div><div class="grid">'+fs.map(card).join("")+'</div>'});
 list.innerHTML=html||'<div class="empty">'+(q?"「"+esc(q)+"」に合う資料はありません":"まだ資料がありません。<br>ドライブのフォルダに入れると、ここに出ます")+'</div>';
 foot.innerHTML='資料を足す・直すときは <a href="'+esc(D.folderUrl)+'" target="_blank" rel="noopener">ドライブのフォルダ</a> に入れるだけです（反映は最大10分。急ぐときは右上の↻）。ファイル名の頭に★でおすすめ、ドライブの「説明」がカードの説明になります。'+(D.loadedAt?'<br>読み込み '+esc(D.loadedAt):"");
}
function load(fresh){document.getElementById("reload").disabled=true;
 google.script.run.withSuccessHandler(d=>{D=d;try{localStorage.setItem("docs.v1",JSON.stringify(d))}catch(e){}render();document.getElementById("reload").disabled=false})
 .withFailureHandler(e=>{document.getElementById("sub").textContent="読み込めませんでした："+(e&&e.message||e);document.getElementById("reload").disabled=false}).getSalesDocs(!!fresh)}
document.getElementById("chips").addEventListener("click",e=>{const b=e.target.closest("[data-i]");if(!b)return;cat=Number(b.dataset.i);try{localStorage.setItem("docs.cat",cat)}catch(x){}render()});
document.getElementById("q").addEventListener("input",render);
document.getElementById("reload").addEventListener("click",()=>load(true));
render();load(false);
`;

// ---------- 契約進捗アプリとの連携 ----------
// 契約進捗DB（契約進捗アプリのデータのスプレッドシート）。契約進捗DBを開いたときのURLの /d/ と /edit の間を貼る
const CONTRACT_DB_ID = '1JoRN8JFn1WG2n2vCOJsdJ796CYK1_n0IMadmFnNJs3k';
const CONTRACT_APP_URL = 'https://script.google.com/a/macros/monochro-inc.com/s/AKfycbwJnqpysI8KL_BsuzDaL2N1ooBBFLtQhxT0e4cVRjEwMEB_Xz0EaGde_STwbfn2Pr0oXw/exec';
const CONTRACT_CACHE_KEY = 'contract_v1';
// マニュアル（新人用）のアプリ。上のタブ「📖 マニュアル」で別のタブに開く。空ならタブを出さない
const MANUAL_APP_URL = 'https://script.google.com/a/macros/monochro-inc.com/s/AKfycbzMAYjG3jPmJq2Nr7NeqAUU5ZuaomuDbDA5cBVwbckjk-35joDFOQPuW__aW-7w0Oe9/exec';
const CONTRACT_MODE_LIST = { '対面': '対面・郵送', '郵送': '対面・郵送', '電子': '電子' };

function contractDb_() {
  if (!CONTRACT_DB_ID) throw new Error('Code.gs の CONTRACT_DB_ID（契約進捗DBのID）が空です');
  return SpreadsheetApp.openById(CONTRACT_DB_ID);
}
// 契約進捗DBのシートを、1行目の見出しをキーにしたオブジェクトの配列で読む
function readContractSheet_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const head = v[0].map(String);
  return v.slice(1).filter((r) => r.some((x) => x !== '')).map((r) => head.reduce((o, h, i) => { o[h] = r[i]; return o; }, {}));
}
// 名前をゆるく比べるための形：空白・「様」「さん」・カッコの中（フリガナ）を除く
function contractName_(s) {
  return String(s || '').replace(/[（(][^）)]*[）)]/g, '').replace(/[\s　]/g, '').replace(/(様|さま|さん)$/, '');
}
// 担当は先頭2文字で比べる（売上表のブロック名が「小嶋」「小嶋勇輝」のどちらでもよいように）
const contractStaff_ = (s) => String(s || '').replace(/[\s　]/g, '').slice(0, 2);
// staffs：主担当と送客元（送客の案件は、売上表では送客元のブロックに入っていることもあるので両方で探す）
function findContract_(deals, name, staffs) {
  const nm = contractName_(name), sts = [].concat(staffs).filter(Boolean).map(contractStaff_);
  const hits = deals.filter((d) => d['案件ID'] && contractName_(d['お客様名']) === nm && sts.indexOf(contractStaff_(d['担当'])) >= 0
    && d['状態'] !== 'キャンセル' && d['状態'] !== '対象外');
  hits.sort((a, b) => (b['作成日時'] instanceof Date ? b['作成日時'] : 0) - (a['作成日時'] instanceof Date ? a['作成日時'] : 0));
  return hits[0] || null;
}
// 顧客管理の行の顧客の、契約進捗の案件をつなぐ。create=true なら見つからないとき新しく作る。売買は作らない
function ensureContract_(row, name, create) {
  const sheet = getSheet_(), map = ensureAppColumns_(sheet);
  const actualRow = resolveRow_(sheet, map, row, name);
  const r = sheet.getRange(actualRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  if (r[map.contract]) return { id: r[map.contract], row: actualRow, created: false };
  if (map.dealType != null && r[map.dealType].indexOf('売買') >= 0) return null; // 契約進捗は賃貸だけ
  if (!CONTRACT_DB_ID) return null;
  const owners = parseOwners_(r[map.staff]);
  const staff = owners ? owners.owners[0] : String(r[map.staff]).trim();
  const ss = contractDb_();
  // 契約進捗DBを読んで探す間は、保存の順番待ち（ロック）をふさがない（読むのに数秒かかるため）
  const who = owners ? owners.owners : [staff];
  let deal = findContract_(readContractSheet_(ss, '案件'), r[map.name], who), created = false;
  if (!deal && create) {
    // 新しく作るときだけ順番待ちして、その間にほかで作られていないかもう一度探す（二重に作らないように）
    const lock = LockService.getScriptLock();
    lock.waitLock(LOCK_WAIT);
    try {
      deal = findContract_(readContractSheet_(ss, '案件'), r[map.name], who);
      if (!deal) {
        const sh = ss.getSheetByName('案件');
        const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
        const now = new Date();
        deal = { '案件ID': Utilities.formatDate(now, 'Asia/Tokyo', 'yyMMdd') + '-' + Utilities.getUuid().slice(0, 6), '作成日時': now,
          '担当': staff, 'お客様名': String(r[map.name]).trim(), '月タブ': '', '行': '', '物件': map.property != null ? r[map.property] : '',
          '契約方法': '対面', '状態': '進行中', '最終活動': now };
        sh.appendRow(head.map((h) => (deal[h] === undefined ? '' : deal[h])));
        const log = ss.getSheetByName('履歴');
        if (log) log.appendRow([now, deal['案件ID'], '作成', '見込みボードの成約から（顧客管理 ' + actualRow + '行目）', currentStaffName_() || currentEmail_()]);
        created = true;
      }
    } finally {
      lock.releaseLock();
    }
  }
  if (!deal) return null;
  sheet.getRange(actualRow, map.contract + 1).setValue(deal['案件ID']); // 1セルだけなので順番待ちしない
  clearCache_();
  CacheService.getScriptCache().remove(CONTRACT_CACHE_KEY);
  return { id: deal['案件ID'], row: actualRow, created: created };
}
/** 詳細を開いたとき：契約進捗の案件を探してつなぐ（作らない）。{ id } か null */
function linkContract(row, name) { return ensureContract_(row, name, false); }
/**
 * 名前が合わずに見つからなかったときの候補：同じ担当（主担当か送客元）の契約進捗の案件のうち、まだ顧客管理のどの顧客にもつながっていないもの
 * 名前が近い順（片方がもう片方を含む → 先頭がそろう文字数）に並べる。[{ id, name, staff, tab, created, near }]
 */
function contractCandidates(row, name) {
  if (!CONTRACT_DB_ID) return [];
  const sheet = getSheet_(), map = ensureAppColumns_(sheet);
  const actualRow = resolveRow_(sheet, map, row, name);
  const r = sheet.getRange(actualRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const owners = parseOwners_(r[map.staff]);
  const sts = (owners ? owners.owners : [r[map.staff]]).map(contractStaff_);
  const used = {};
  sheet.getRange(2, map.contract + 1, Math.max(1, sheet.getLastRow() - 1), 1).getDisplayValues().forEach((x) => { if (x[0]) used[x[0]] = true; });
  const nm = contractName_(r[map.name]);
  const score = (s) => { const t = contractName_(s); if (t === nm) return 100; if (t && nm && (t.indexOf(nm) >= 0 || nm.indexOf(t) >= 0)) return 50; let k = 0; while (k < t.length && t[k] === nm[k]) k++; return k; };
  return readContractSheet_(contractDb_(), '案件')
    .filter((d) => d['案件ID'] && !used[d['案件ID']] && sts.indexOf(contractStaff_(d['担当'])) >= 0 && d['状態'] !== 'キャンセル' && d['状態'] !== '対象外')
    .map((d) => ({ id: d['案件ID'], name: String(d['お客様名']), staff: String(d['担当']), tab: String(d['月タブ'] || ''),
      created: d['作成日時'] && d['作成日時'].getTime ? Utilities.formatDate(d['作成日時'], 'Asia/Tokyo', 'M/d') : '', near: score(d['お客様名']) }))
    .sort((a, b) => b.near - a.near || (a.id < b.id ? 1 : -1))
    .slice(0, 30);
}
/** 候補から選んだ案件とつなぐ */
function linkContractTo(row, name, id) {
  if (!/^[A-Za-z0-9-]{1,40}$/.test(String(id || ''))) throw new Error('案件IDが正しくありません');
  if (!readContractSheet_(contractDb_(), '案件').some((d) => d['案件ID'] === id)) throw new Error('契約進捗に案件が見つかりません：' + id);
  const res = writeFields_(row, name, { contract: id });
  CacheService.getScriptCache().remove(CONTRACT_CACHE_KEY);
  return { id: id, row: res.row };
}
/** 「契約進捗を作る」ボタン */
function createContract(row, name) {
  const ct = ensureContract_(row, name, true);
  if (!ct) throw new Error('契約進捗を作れませんでした（売買の顧客か、CONTRACT_DB_ID が空です）');
  return ct;
}
/** カードに出す進み具合（5分キャッシュ）：{ url, deals: { 案件ID: { done, total, status, idle, phase } } } */
function getContractProgress() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(CONTRACT_CACHE_KEY);
  if (hit) return JSON.parse(hit);
  const out = { url: CONTRACT_APP_URL, deals: {} };
  if (!CONTRACT_DB_ID) return out;
  const ss = contractDb_();
  const items = readContractSheet_(ss, '項目マスタ').filter((i) => i['有効'] !== false && i['種類'] !== '注意');
  const checks = readContractSheet_(ss, 'チェック').reduce((o, c) => { (o[c['案件ID']] = o[c['案件ID']] || {})[c['項目ID']] = c['状態']; return o; }, {});
  readContractSheet_(ss, '案件').forEach((d) => {
    const id = d['案件ID'];
    if (!id) return;
    const mode = d['契約方法'] || '対面', list = CONTRACT_MODE_LIST[mode] || '対面・郵送', st = checks[id] || {};
    let done = 0, total = 0, phase = '';
    items.forEach((i) => {
      if (i['契約方法'] !== list || (i['段階'] === '郵送契約のみ' && mode !== '郵送')) return;
      const v = st[i['ID']];
      if (v === 'na') return;
      total++;
      if (v === 'done') done++; else if (!phase) phase = i['段階'];
    });
    const last = d['最終活動'];
    out.deals[id] = { done: done, total: total, status: d['状態'], phase: phase,
      idle: last && last.getTime ? Math.floor((Date.now() - last.getTime()) / 36e5) : 0 };
  });
  try { cache.put(CONTRACT_CACHE_KEY, JSON.stringify(out), 300); } catch (e) { /* 大きすぎてもそのまま返す */ }
  return out;
}
