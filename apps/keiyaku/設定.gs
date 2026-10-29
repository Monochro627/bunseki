/**
 * 契約進捗アプリの設定
 * Code.gs を新しいものに貼り替えるときも、このファイルはそのまま残してください。
 */

// ===== 設定 =====
const CONFIG = {
  // 売上管理表のID。テスト中は「売上管理表のコピー」のIDに差し替える
  SALES_ID: '1QLNfTOsi6BpFPkxZbp9e9mjxbvFGy-6TZi6uTksOuxI',
  // ウェブアプリとしてデプロイした後に表示されるURL（.../exec）を貼る
  WEB_APP_URL: 'https://script.google.com/a/macros/monochro-inc.com/s/AKfycbwJnqpysI8KL_BsuzDaL2N1ooBBFLtQhxT0e4cVRjEwMEB_Xz0EaGde_STwbfn2Pr0oXw/exec',
  // リンクを付ける売上表のブロック（A列の名前）。[] にすると全員
  TARGET_STAFF: [],
  // 対象外にするブロック（A列の名前がこの言葉で始まるもの）
  EXCLUDE_BLOCKS: ['インスタ', '飛び込み'],
  // 監視する月次タブの名前
  TAB_PATTERN: /^月次売上\d+月※\d+期目$/,
  // 上の条件に合うタブのうち、シートの並び順で先頭から何枚を見るか（月替わりに前月分も拾うため2）
  TAB_SCAN_COUNT: 2,
  // このタブより古いタブは見ない（本番開始の月）
  FIRST_TAB: '月次売上10月※5期目',
  // 何時間チェックが動かなければアラートにするか
  STALL_HOURS: 48,
  // 重説の月次集計メールと、メール未登録の担当者のアラートの送り先。複数はカンマ区切り（空ならアプリの持ち主）
  ADMIN_EMAIL: '',
  TZ: 'Asia/Tokyo',
};
