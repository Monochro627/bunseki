/**
 * 分析ボード（見込みボードと同じプロジェクトの別ページ。URL は見込みボードの末尾に ?page=analysis）
 * ファイル名はスクリプト「AnalysisCode」、画面は HTML「Analysis」（Apps Script では同じ名前を2つ付けられない）
 *
 * Looker Studio の成績表を1ページずつ作り直していく場所。見込みボード本体（Code.gs / Index.html）とは分けて置く。
 * ページを足すときは ANALYSIS_PAGES に1行足し、Analysis.html にそのページの描画を足す。
 * 見込みボードの「分析」タブを押すと、ここ（最後に見ていたページ）に切り替わる。上のタブで見込みボードに戻る。
 *
 * 今あるページ
 *   person … 担当者シート（担当を1人選ぶと、成績・年間予算・成約率・紹介・入社年次・期ごとの推移を1枚で。順位と平均つき）
 *   score  … 担当別成績表（売上・申込・紹介・口コミ・縦掘り・自己集客・インスタ・フォロワーを担当別に。期間を選べる。累積もここで見る）
 *   monthly… 月別推移（期と担当を選んで、4月〜3月の売上・申込・紹介・縦掘り・口コミを月ごとの縦棒で縦に5つ。前の期の同じ月と並べる。担当別成績表の下にも出す）
 *   nenji  … 入社年次比較（入社月から12か月ずつ区切って、同じ年次同士で比べる）
 *   budget … 年間予算達成率（会社と各担当の今期の売上 ÷ 目標を半円ゲージで）
 *   promo  … 昇格要件（今の役職・次の昇格の要件：売上・インスタ・口コミのゲージと、資格や試験のクリアボタン。降格ライン）
 *   rate   … 成約率＆平均単価（紹介・インスタの段階ごとの割合と単価。担当で絞れる）
 *   intro  … 紹介分析（紹介の内訳・縦掘り率・アポなど。担当で絞れる）
 *   sale   … 売買成績表（売買の売上・成約数と、今期の売買目標の達成率）
 *   mgrowth… メンバー成長率（担当ごとに1期目〜今期の売上・申込数・紹介数）
 *   cgrowth… 会社成長率（会社の年商・契約数・インスタの売上と契約数を期ごとに、昨対比つき）
 *   funnel … 成約までの段階（紹介→初回電話→面談→物件提案→内覧→成約の人数と割合、失客の内訳。Code.gs の getFunnel）
 * nenji 以外は同じデータ（getScoreData）を使い、計算は画面側でする（person は getNenjiData も使う）
 */

const ANALYSIS_PAGES = [
  { key: 'person', title: '担当者シート' },
  { key: 'score', title: '担当別成績表' },
  { key: 'monthly', title: '月別推移' },
  { key: 'nenji', title: '入社年次比較' },
  { key: 'budget', title: '年間予算達成率' },
  { key: 'promo', title: '昇格要件' },
  { key: 'rate', title: '成約率＆平均単価' },
  { key: 'intro', title: '紹介分析' },
  { key: 'sale', title: '売買成績表' },
  { key: 'mgrowth', title: 'メンバー成長率' },
  { key: 'cgrowth', title: '会社成長率' },
  { key: 'funnel', title: '成約までの段階' },
];
// 売買の案件数：顧客管理シートで「取引種別」が売買の人を、紹介月・担当者（送客は主担当）ごとに数えて足す
//   saleIntro：紹介（紹介種別が「自己集客（…）」か空）、saleOther：紹介以外（会社集客のインスタ・飛び込み（HP）・SUUMO など）
//   …Won：そのうち進捗ステータスが「完了」の人（同じ紹介月の人の中で成約率を出すため）
const SALE_DEAL_KEYS = ['saleIntro', 'saleIntroWon', 'saleOther', 'saleOtherWon'];
function addSaleIntro_(keys, rows, partial) {
  const base = keys.length;
  SALE_DEAL_KEYS.forEach((k) => keys.push(k));
  rows.forEach((r) => SALE_DEAL_KEYS.forEach(() => r.push(0)));
  try {
    const sheet = getSheet_();
    const head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map((h) => String(h).trim());
    const col = (k) => head.indexOf(COLS[k]);
    const c = { name: col('name'), status: col('status'), deal: col('dealType'), staff: col('staff'), intro: col('introMonth'), type: col('introType') };
    if (c.deal < 0 || c.intro < 0 || c.staff < 0) return;
    const last = Math.max(c.name, c.status, c.deal, c.staff, c.intro, c.type) + 1;
    const values = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), last).getDisplayValues();
    const at = {};
    rows.forEach((r, i) => { at[r[0] + '|' + r[1]] = i; });
    values.forEach((r) => {
      if (!String(r[c.name]).trim() || String(r[c.deal]).indexOf('売買') < 0) return;
      if (String(r[c.status]).indexOf('送客被り') >= 0) return;
      const m = String(r[c.intro]).match(/(\d{2,4})\D+(\d{1,2})/);
      if (!m) return;
      const ym = (m[1].length === 2 ? '20' + m[1] : m[1]) + '/' + ('0' + m[2]).slice(-2);
      if (ym > partial) return;
      const owners = parseOwners_(r[c.staff]);
      if (!owners) return;
      const name = owners.owners[0], k = ym + '|' + name;
      if (at[k] == null) { at[k] = rows.length; rows.push([ym, name].concat(keys.map(() => 0))); }
      const type = c.type >= 0 ? String(r[c.type] || '').trim() : '';
      const other = !!type && type.indexOf('自己集客') !== 0; // 会社集客・会社送客など
      const won = resultOf_(String(r[c.status]).trim()) === 'won';
      const i = 2 + base + (other ? 2 : 0);
      rows[at[k]][i]++;
      if (won) rows[at[k]][i + 1]++;
    });
  } catch (e) { /* 読めなくても他の数字は出す */ }
}

// 前にあったページのキー → 今のページ（総合成績表は担当別成績表にまとめた）
const ANALYSIS_ALIASES = { total: 'score' };

// ---------- 担当別成績表・予算 ----------
// 「★紹介分析データ用（新）」の見出しの名前で列を探す（列が増えても動くように）
const SCORE_HEADERS = {
  rev: '売上', app: '総申込数', intro: '総紹介数', kuchikomi: '口コミ',
  tate: '縦掘り合計', tateRev: '縦掘り売上', tateCnt: '縦掘り成約数',
  selfRev: '自己集客元売上', instaRev: 'インスタ元売上', instaApp: 'インスタ申込数',
  // 成約率＆平均単価
  callMeet: '初回電話or面談（紹介）', viewIntro: '内覧数（紹介）', selfApp: '自己集客申込数',
  instaMeet: 'インスタ面談数', instaView: 'インスタ内覧数', rentRev: '賃貸元売上', rentApp: '賃貸のみ申込数',
  // 売買
  saleRev: '売買元売上', saleCnt: '売買成約数',
  // 紹介分析
  kizon: '既存人脈', kizonIntro: '既存人脈による紹介', shinki: '新規人脈', shinkiIntro: '新規人脈による紹介',
  custIntro: '既存顧客による紹介', custRepeat: '既存顧客のリピート', kobe: '神戸アカによる集客',
  instaIntro: 'インスタ顧客による紹介', instaRepeatIntro: 'インスタ顧客によるリピート紹介',
  kizonApo: '既存アポ', shinkiApo: '新規アポ', newIntro: '新規紹介合計', instaTate: 'インスタの縦掘り案件',
};
const TATE_TARGET = 1000; // 紹介分析の縦掘り数の目標（全員を選んだとき。Looker の「縦掘り数累積」のゲージと同じ）
// 担当を選んだときの縦掘りの目標は「縦掘り目標」シート（顧客管理スプレッドシート。アプリが作って書く。画面の入力欄から保存）
//   A=担当者、B=種類（年次／期）、C=番号（年次なら1,2…年目、期なら5期目の5）、D=目標件数、E=更新者、F=更新日時
//   年次：入社月（NENJI_STAFF の start）から12か月ずつ。今の年次の12か月の縦掘り数をその年次の目標と比べる
const TATE_GOAL_SHEET = '縦掘り目標';
const SCORE_CACHE_KEY = 'score_v10'; // getScoreData のキャッシュ（中身を変えたら番号を上げる）
const TATE_GOAL_HEAD = ['担当者', '種類', '番号', '目標件数', '更新者', '更新日時'];
// 売買の今期の目標件数は、売上管理表「5期目　年間売上表」B列の「売買総合計【30件】」から読む（読めなければこの数）
const SALE_TARGET_DEFAULT = 30;

// 会社成長率：売上管理表の各期の「◯期目　年間売上表」から読む（Looker と同じ数字）。期によって行の名前が違うので、名前で探す
// ラベルは空白を除いて比べる。上から順に探して、最初に見つかったものを使う（足すものは全部足す）
const GROWTH_ROWS = {
  // 年商：売上の表の会社全体の行
  rev: [/^Monochro全体売上計(申込ベース)?$/, /^全体売上計※?役員込み$/],
  // 総契約数：申込件数の表の会社全体の行
  cnt: [/^Monochro全体申込件数$/, /^全体売上計※役員有り$/, /^全体計※役員有り$/, /^全体計$/],
  // インスタ売上：インスタ・飛び込み（HP）・神戸アカの元売上を足す
  irev: [/^※?インスタ※?元売上$/, /^飛び込み.*※元売上$/, /^インスタ※神戸アカ$/],
  // インスタ契約数：同じく元申込件数を足す
  icnt: [/^※?インスタ※?元申込件数$/, /^飛び込み.*※元申込件数$/, /^インスタ※神戸アカ申込件数$/],
};
const GROWTH_SUM = { rev: false, cnt: false, irev: true, icnt: true }; // true は見つかった行を全部足す
// 担当者のほかに読む行。インスタ＝インスタ等の売上（年間予算では諸本さんの分に足す）、Monochro＝会社全体
const SCORE_EXTRA_ROWS = ['インスタ', 'Monochro'];

// 年間予算達成率：目標は売上管理表の「年間売上目標」の【】の金額（見込みボードの年間ゲージと同じ。Code.gs の getSales）
// 売上は紹介分析シートの今期の「売上」。諸本さんはインスタ等の行も足す（目標がインスタ・飛び込み・神戸アカ込みのため）
const BUDGET_COMPANY_TARGET = 100000000; // 会社全体の年間目標（Monochro の行の売上と比べる）
const BUDGET_ADD_ROWS = { '諸本雅広': ['インスタ'] };
const BUDGET_TITLES = { '佐光琴音': '主任', '澤井颯晟': '主任', '松井拓登': '主任', '山田園佳': '主任', '諸本雅広': 'インスタ・飛び込み・神戸アカ込み' };

// 昇格予算達成率：これまでの累積（シートにある全部の月）が基準に届いたか。metric は rev＝売上、selfRev＝自己集客元売上
// 昇格の段階と要件（下から順に）。金額は累積で比べる
//   metric：rev＝売上、selfRev＝自己集客元売上。minusSale：売買元売上を引く（主任は売買除く）
//   insta：インスタ案件の成約（インスタ申込数）の件数、reviews：Google口コミ★5の件数（「★紹介分析データ用（新）」の「口コミ」列＝AK列の累積）
//   checks：ボタンで「クリア」にする要件、demote：降格ライン（昇格月から12か月の自己集客元売上がこれ以下で1ランクダウン）
//   chiefOthers：売上以外の主任要件クリア（自動）、demoteNotes：文字で出す降格要件
//   from：ひとつ前の段階の番号（主任の次は、エージェントマネージャーの道と営業を極める道の2つに分かれる）、path：道の名前
const PROMO_STEPS = [
  { title: '主任', short: '主任', from: -1, path: '', metric: 'rev', minusSale: true, target: 6000000, insta: 3, reviews: 30,
    checks: ['人物評価', 'ロープレ試験', '原状回復試験の合格'] },
  { title: 'エージェントマネージャー', short: 'エージェントマネージャー', from: 0, path: 'エージェントマネージャーの道', metric: 'selfRev', target: 6000000,
    chiefOthers: true, // 「売上以外の主任要件クリア」：主任の要件のうち売上以外が全部そろったら自動でクリア
    checks: ['リファラル2人以上（ルームアドバイザー認定を受けた者）', 'エージェントアドバイザー研修の受講＆試験', '経営陣の最終判断', 'インボイス登録'],
    demoteNotes: ['2年間採用0人', 'リファラルが2人未満になった場合', '3ヵ月連続で売上30万未満'] }, // 降格要件（金額で見られないので文字で出す）
  { title: 'コンサルティングルームアドバイザー', short: '課長代理', from: 0, path: '営業を極める道', metric: 'selfRev', target: 10000000, reviews: 50,
    checks: ['FP3級の合格', '人物評価', 'Monochro試験', '経営陣の最終判断', 'インボイス登録・源泉徴収10%'], demote: 5500000 },
  { title: 'プレジデントルームアドバイザー', short: '課長', from: 2, path: '営業を極める道', metric: 'rev', target: 30000000, reviews: 100,
    checks: ['敷金診断士', '人物評価', 'Monochro試験', '経営陣の最終判断', 'インボイス登録・源泉徴収10%（支払調書）'], demote: 6000000 },
  { title: 'エグゼクティブルームアドバイザー', short: '次長', from: 3, path: '営業を極める道', metric: 'selfRev', target: 40000000, reviews: 150,
    checks: ['宅地建物取引士の合格＆専任登録', '人物評価', 'Monochro試験', '経営陣の最終判断', 'インボイス登録・源泉徴収10%'], demote: 6500000 },
  { title: 'トップ・オブ・ザ・エグゼクティブルームアドバイザー', short: 'トップ', from: 4, path: '営業を極める道', metric: 'selfRev', target: 50000000, reviews: 200,
    checks: ['人物評価', 'Monochro試験', '経営陣の最終判断', 'インボイス登録・源泉徴収10%'], demote: 7000000 },
];
const PROMO_SHEET = '昇格チェック';  // 顧客管理スプレッドシート。アプリが作って書く（今の役職・昇格月・要件のクリア）
const FOLLOWER_SHEET = 'フォロワー'; // A列=名前、B列=フォロワー数、C列=更新日。「全体総フォロワー」の行は合計

/**
 * 担当別成績表のデータ。月ごとの数字を返し、期間の合計は画面側でする
 * { partial: "2026/10", keys: [...], rows: [["2026/04", "松井拓登", 数字...]], staff: [{name, color}], followers: [...], followerTotal }
 */
function getScoreData() {
  const cacheKey = SCORE_CACHE_KEY;
  const cached = readCacheKey_(cacheKey);
  if (cached) return cached;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const partial = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM');
  const keys = Object.keys(SCORE_HEADERS);
  const names = STAFF.map((s) => s.name).concat(SCORE_EXTRA_ROWS);
  const rows = readAnalysisSheet_(ss, NENJI_SHEET_NEW, keys.map((k) => SCORE_HEADERS[k]))
    .filter((r) => names.indexOf(r.name) >= 0 && r.ym <= partial)
    .map((r) => [r.ym, r.name].concat(r.vals));
  addSaleIntro_(keys, rows, partial);

  const fs = ss.getSheetByName(FOLLOWER_SHEET);
  const followers = [];
  let followerTotal = 0;
  if (fs && fs.getLastRow() > 1) {
    fs.getRange(2, 1, fs.getLastRow() - 1, 3).getDisplayValues().forEach((r) => {
      const name = String(r[0]).trim(), n = Number(String(r[1]).replace(/[^\d.]/g, '')) || 0;
      if (!name) return;
      if (name.indexOf('総フォロワー') >= 0) followerTotal = n;
      else followers.push({ name: name, count: n, updated: String(r[2]).trim() });
    });
  }
  // 年間目標（売上管理表）。読めなくても他のページは動くようにする
  let targets = {}, targetError = '', saleTarget = SALE_TARGET_DEFAULT, saleCompany = null;
  try {
    const sales = getSales();
    sales.staff.forEach((s) => { targets[s.name] = s.targetTotal; });
    // 今期の会社全体の売買（売上管理表の「★売買総申込件数」「（分析）全体売買売上」の行。月は今期の最初の月から順に12か月）
    if (sales.companySaleAppMonths || sales.companySaleMonths) {
      saleCompany = { from: SALES_PERIOD_START.slice(0, 7).replace('-', '/'), cnt: sales.companySaleAppMonths, rev: sales.companySaleMonths };
    }
    const labels = SpreadsheetApp.openById(SALES_SPREADSHEET_ID).getSheetByName(SALES_SHEET_NAME).getRange('B1:B40').getDisplayValues();
    labels.forEach((r) => {
      const m = String(r[0]).replace(/[\s　]/g, '').match(/^売買総合計.*?【(\d+)件/);
      if (m) saleTarget = Number(m[1]);
    });
  } catch (e) { targetError = e.message; }
  const data = {
    partial: partial, keys: keys, rows: rows, followers: followers,
    followerTotal: followerTotal || followers.reduce((a, f) => a + f.count, 0),
    staff: STAFF.map((s) => ({ name: s.name, color: s.color })),
    budget: {
      from: SALES_PERIOD_START.slice(0, 7).replace('-', '/'), to: SALES_PERIOD_END.slice(0, 7).replace('-', '/'),
      start: SALES_PERIOD_START, end: SALES_PERIOD_END, companyTarget: BUDGET_COMPANY_TARGET,
      targets: targets, targetError: targetError, addRows: BUDGET_ADD_ROWS, titles: BUDGET_TITLES,
    },
    promoSteps: PROMO_STEPS, tateTarget: TATE_TARGET, tateGoals: readTateGoals_(), nenjiStarts: nenjiStarts_(), saleTarget: saleTarget, saleCompany: saleCompany,
    growth: readGrowth_(),
    loadedAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'),
  };
  writeCacheKey_(cacheKey, data, 600);
  return data;
}

/**
 * 会社成長率のデータ。売上管理表の「◯期目　年間売上表」を期ごとに読む
 * [{ n: 5, months: ["4月", ...], rev: [月ごと], cnt, irev, icnt, total: {rev, cnt, irev, icnt}, missing: [見つからなかった項目] }]
 * 読めなければ { error }
 */
function readGrowth_() {
  try {
    const ss = SpreadsheetApp.openById(SALES_SPREADSHEET_ID);
    const norm = (v) => String(v || '').replace(/[\s　]/g, '');
    const num = (v) => { const s = String(v || ''); return s.indexOf('#') >= 0 ? 0 : Number(s.replace(/[^\d.\-]/g, '')) || 0; };
    const sheets = ss.getSheets().map((sh) => ({ sh: sh, m: norm(sh.getName()).match(/^(\d+)期目年間売上表$/) }))
      .filter((x) => x.m).sort((a, b) => Number(a.m[1]) - Number(b.m[1]));
    return sheets.map((x) => {
      const sh = x.sh;
      const rows = sh.getRange(1, 2, Math.min(sh.getLastRow(), 400), 20).getDisplayValues(); // B列から
      const labels = rows.map((r) => norm(r[0]));
      // 表の見出しの行（「年間売上表」「年間申込件数表」）から、月の列と「合計」の列を探す
      const block = (title, next) => {
        const start = labels.findIndex((l) => l === title);
        if (start < 0) return null;
        let end = labels.findIndex((l, i) => i > start && l === next);
        if (end < 0) end = rows.length;
        const head = rows[start].map(norm), totalCol = head.indexOf('合計');
        const monthCols = [];
        for (let c = 1; c < (totalCol > 0 ? totalCol : head.length); c++) if (/^\d+月$/.test(head[c])) monthCols.push(c);
        return { start: start, end: end, totalCol: totalCol, monthCols: monthCols, months: monthCols.map((c) => head[c]) };
      };
      const sale = block('年間売上表', '年間申込件数表'), apps = block('年間申込件数表', '年間紹介件数表');
      const out = { n: Number(x.m[1]), months: sale ? sale.months : [], total: {}, missing: [] };
      Object.keys(GROWTH_ROWS).forEach((k) => {
        const b = (k === 'rev' || k === 'irev') ? sale : apps;
        const monthly = new Array(b ? b.monthCols.length : 0).fill(0);
        let total = 0, found = false;
        if (b) {
          for (const re of GROWTH_ROWS[k]) {
            for (let i = b.start + 1; i < b.end; i++) {
              if (!re.test(labels[i])) continue;
              b.monthCols.forEach((c, j) => { monthly[j] += num(rows[i][c]); });
              total += b.totalCol > 0 ? num(rows[i][b.totalCol]) : 0;
              found = true;
              if (!GROWTH_SUM[k]) break;
            }
            if (found && !GROWTH_SUM[k]) break;
          }
        }
        if (!found) out.missing.push(k);
        out[k] = monthly;
        out.total[k] = total;
      });
      return out;
    });
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * 紹介分析シートを読む。[{ ym: "2026/04", name, vals: [見出しの順の数字] }]
 * 数字はセルの値で読む（小数も落とさない）。件数の欄の「¥…」や #REF! など、数字でないものは 0
 */
function readAnalysisSheet_(ss, sheetName, headers) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('「' + sheetName + '」シートが見つかりません');
  const lastCol = sheet.getLastColumn();
  const head = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map((h) => String(h).trim());
  const idx = headers.map((h) => {
    const i = head.indexOf(h);
    if (i < 0) throw new Error('「' + sheetName + '」に「' + h + '」の列が見つかりません');
    return i;
  });
  const n = sheet.getLastRow() - 1;
  if (n < 1) return [];
  const range = sheet.getRange(2, 1, n, lastCol);
  const shown = range.getDisplayValues(), raw = range.getValues();
  const out = [];
  shown.forEach((r, i) => {
    const ym = String(r[0]).trim();
    if (!/^\d{4}\/\d{2}$/.test(ym)) return;
    out.push({ ym: ym, name: String(r[1]).trim(), vals: idx.map((c) => analysisNum_(raw[i][c])) });
  });
  return out;
}

function analysisNum_(v) {
  if (typeof v === 'number') return v;
  const t = String(v || '').trim();
  if (!t || t.indexOf('#') >= 0) return 0;
  return Number(t.replace(/[^\d.\-]/g, '')) || 0;
}

// ---------- 入社年次比較 ----------
// 元データ：顧客管理スプレッドシートの紹介分析シート（担当者×月で1行。A列=月 "2026/04"、B列=名前）
const NENJI_SHEET_NEW = '★紹介分析データ用（新）';
const NENJI_SHEET_OLD = '★紹介分析データ用（旧）4期目まで'; // 退職した人（伊藤さん）の分
// 使う列（0始まり）：L縦掘り合計・M総紹介数・P自己集客申込数・Q自己集客元売上・V総申込数・W売上
const NENJI_COLS = { tate: 11, intro: 12, selfApp: 15, selfRev: 16, app: 21, rev: 22 };
// 入社月（1か月目）。この順に並ぶ。色は STAFF（Code.gs）の色、STAFF にいない人は color を書く
const NENJI_STAFF = [
  { name: '小嶋勇輝', start: '2021/02' },
  { name: '伊藤竣哉', start: '2022/08', end: '2025/08', sheet: NENJI_SHEET_OLD, color: '#34a853' },
  { name: '諸本雅広', start: '2022/11' },
  { name: '山田園佳', start: '2023/07' },
  { name: '佐光琴音', start: '2023/10' },
  { name: '澤井颯晟', start: '2023/12' },
  { name: '松井拓登', start: '2025/04' },
  { name: '寺田喜心', start: '2025/10' },
  { name: '宮川大輝', start: '2026/07' },
];
// シートより前の分の手入力（この期間のシートの行は使わない）。書いていない指標は「データなし」になる
//   小嶋さん 1年目：手入力の値。縦掘りは「小嶋ラホーム時代」シートD列の既存顧客による紹介＋リピート
//   小嶋さん 2年目の前半（2022/02〜2022/07）：手入力の値から、シートにある 2022/08〜2023/01 の分を引いたもの
//     縦掘りは「小嶋ラホーム時代」シート147行目以降の件数（27＋13＝40件）
const NENJI_MANUAL = {
  '小嶋勇輝': [
    { from: '2021/02', to: '2022/01', tate: 33, intro: 142, app: 100, rev: 10689976 },
    { from: '2022/02', to: '2022/07', tate: 40, intro: 107, app: 69, rev: 6838375 },
  ],
};

/**
 * 入社年次比較のデータ。月ごとの数字を担当者ごとに返し、年次への区切りは画面側でする
 * { partial: "2026/10"（今月＝途中の月）, staff: [{ name, start, color, months: { "2026/04": {tate,...} }, manual: [...] }] }
 */
function getNenjiData() {
  const cacheKey = 'nenji_v1';
  const cached = readCacheKey_(cacheKey);
  if (cached) return cached;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const partial = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM');
  const sheets = {};
  const rowsOf = (sheetName) => {
    if (!sheets[sheetName]) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) throw new Error('「' + sheetName + '」シートが見つかりません');
      // 月・名前は表示どおりの文字で、数字は小数も落とさないようにセルの値で読む
      const range = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), 23);
      const shown = range.getDisplayValues(), raw = range.getValues();
      sheets[sheetName] = shown.map((r, i) => ({ ym: String(r[0]).trim(), name: String(r[1]).trim(), raw: raw[i] }));
    }
    return sheets[sheetName];
  };
  // 数字はそのまま。「0件」「¥514,158」のような文字は数字にする。空欄や #REF! は 0
  const num = (v) => {
    if (typeof v === 'number') return v;
    const t = String(v || '').trim();
    if (!t || t.indexOf('#') >= 0) return 0;
    return Number(t.replace(/[^\d.\-]/g, '')) || 0;
  };
  const colorOf = (s) => s.color || ((STAFF.find((x) => x.name === s.name) || {}).color) || '#8a94a6';

  const staff = NENJI_STAFF.map((s) => {
    const manual = NENJI_MANUAL[s.name] || [];
    const inManual = (ym) => manual.some((m) => ym >= m.from && ym <= m.to);
    const months = {};
    rowsOf(s.sheet || NENJI_SHEET_NEW).forEach((r) => {
      const ym = r.ym;
      if (r.name !== s.name || !/^\d{4}\/\d{2}$/.test(ym)) return;
      if (ym < s.start || ym > partial || (s.end && ym > s.end) || inManual(ym)) return;
      const o = {};
      Object.keys(NENJI_COLS).forEach((k) => { o[k] = num(r.raw[NENJI_COLS[k]]); });
      months[ym] = o;
    });
    return { name: s.name, start: s.start, end: s.end || '', color: colorOf(s), months: months, manual: manual };
  });
  const data = { partial: partial, staff: staff, loadedAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') };
  writeCacheKey_(cacheKey, data, 600);
  return data;
}

// ---------- 縦掘りの目標（担当ごと・年次／期） ----------
function tateGoalSheet_(create) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(TATE_GOAL_SHEET);
  if (!sh && create) { sh = ss.insertSheet(TATE_GOAL_SHEET); sh.getRange(1, 1, 1, TATE_GOAL_HEAD.length).setValues([TATE_GOAL_HEAD]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
/** { 担当: { nenji: {年次: 件数}, ki: {期: 件数} } }。読めなくても他は動くように空で返す */
function readTateGoals_() {
  const out = {};
  try {
    const sh = tateGoalSheet_(false);
    if (!sh || sh.getLastRow() < 2) return out;
    sh.getRange(2, 1, sh.getLastRow() - 1, 4).getDisplayValues().forEach((r) => {
      const name = String(r[0]).trim(), kind = String(r[1]).trim(), n = Number(r[2]), v = String(r[3]).replace(/[^\d.]/g, '');
      if (!name || !n || v === '' || (kind !== '年次' && kind !== '期')) return;
      const o = out[name] = out[name] || { nenji: {}, ki: {} };
      o[kind === '年次' ? 'nenji' : 'ki'][n] = Number(v);
    });
  } catch (e) { /* シートが読めなくても動く */ }
  return out;
}
/** 担当の入社月（年次の起算月）。{ 担当: '2025/10' } */
function nenjiStarts_() {
  const out = {};
  NENJI_STAFF.forEach((s) => { if (!s.end) out[s.name] = s.start; });
  return out;
}
/** 縦掘りの目標を保存。kind は '年次' か '期'、value が空なら目標なしに戻す */
function setTateGoal(name, kind, n, value) {
  if (!STAFF.some((s) => s.name === name)) throw new Error('担当者が見つかりません：' + name);
  if (kind !== '年次' && kind !== '期') throw new Error('種類が違います：' + kind);
  n = Number(n);
  if (!(n >= 1 && n <= 99)) throw new Error('番号が違います：' + n);
  const v = value === '' || value == null ? '' : Math.max(0, Math.round(Number(value)));
  if (v !== '' && !isFinite(v)) throw new Error('目標の件数が数字ではありません');
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_WAIT);
  try {
    const sh = tateGoalSheet_(true);
    const by = currentStaffName_() || currentEmail_() || '不明';
    const at = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
    const cnt = sh.getLastRow() - 1;
    const rows = cnt > 0 ? sh.getRange(2, 1, cnt, 3).getDisplayValues() : [];
    const i = rows.findIndex((r) => r[0] === name && r[1] === kind && Number(r[2]) === n);
    sh.getRange(i >= 0 ? i + 2 : cnt + 2, 1, 1, 6).setValues([[name, kind, n, v, by, at]]);
  } finally { lock.releaseLock(); }
  CacheService.getScriptCache().remove(SCORE_CACHE_KEY); // 担当別成績表などのデータに目標が入っているので捨てる
  return readTateGoals_();
}

// ---------- 昇格チェック（今の役職・昇格月・要件のクリア） ----------
// 「昇格チェック」シート：A=担当者、B=種類（役職／クリア）、C=段階、D=項目、E=値、F=更新者、G=更新日時。同じ担当者・種類・段階・項目の行は上書き
const PROMO_HEAD = ['担当者', '種類', '段階', '項目', '値', '更新者', '更新日時'];
function promoSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(PROMO_SHEET);
  if (!sh) { sh = ss.insertSheet(PROMO_SHEET); sh.getRange(1, 1, 1, PROMO_HEAD.length).setValues([PROMO_HEAD]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
/** 画面用：{ roles: {担当: {step: 段階の番号（-1 はまだ）, month: '2025/04', by, at}}, checks: {担当: {段階: {項目: {by, at}}}} } */
function getPromoState() {
  const sh = promoSheet_();
  const out = { roles: {}, checks: {} };
  if (sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 7).getDisplayValues().forEach((r) => {
    const [name, type, step, item, val, by, at] = r.map((x) => String(x).trim());
    if (!name) return;
    if (type === '役職') {
      const i = PROMO_STEPS.findIndex((s) => s.title === step);
      out.roles[name] = { step: i, month: val, by: by, at: at };
    } else if (type === 'クリア' && val === 'TRUE') {
      ((out.checks[name] = out.checks[name] || {})[step] = out.checks[name][step] || {})[item] = { by: by, at: at };
    }
  });
  return out;
}
function writePromo_(name, type, step, item, val) {
  if (!STAFF.some((s) => s.name === name)) throw new Error('担当者が見つかりません：' + name);
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_WAIT);
  try {
    const sh = promoSheet_();
    const by = currentStaffName_() || currentEmail_() || '不明';
    const at = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
    const n = sh.getLastRow() - 1;
    const rows = n > 0 ? sh.getRange(2, 1, n, 4).getDisplayValues() : [];
    const i = rows.findIndex((r) => r[0] === name && r[1] === type && r[2] === step && r[3] === item);
    const line = [[name, type, step, item, val, by, at]];
    const cell = sh.getRange(i >= 0 ? i + 2 : n + 2, 1, 1, 7);
    cell.setNumberFormat('@').setValues(line); // 「2025/04」が日付に変わらないよう文字で入れる
  } finally { lock.releaseLock(); }
  return getPromoState();
}
/** 今の役職と昇格した月。stepIndex が -1 なら「まだ役職なし」 */
function setPromoRole(name, stepIndex, month) {
  const st = PROMO_STEPS[stepIndex];
  return writePromo_(name, '役職', st ? st.title : '', '', /^\d{4}\/\d{2}$/.test(month || '') ? month : '');
}
/** 要件のクリア（on=true）／取り消し（on=false） */
function setPromoCheck(name, stepTitle, item, on) {
  const st = PROMO_STEPS.find((s) => s.title === stepTitle);
  if (!st || st.checks.indexOf(item) < 0) throw new Error('要件が見つかりません：' + stepTitle + ' ' + item);
  return writePromo_(name, 'クリア', stepTitle, item, on ? 'TRUE' : 'FALSE');
}
