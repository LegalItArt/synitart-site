/**
 * SynItArt 농장 웹사이트 제안 페이지(/farm-web/) 방문 통계 기록기
 *
 * 받는 값: page, t(view·pkg·copy), src(유입경로), pkg, disc, maint, total, ref(이전 사이트 도메인)
 * 받지 않는 값: 이름·연락처·농장 이름·IP·쿠키. 개인을 알아볼 수 없는 숫자만 쌓는다.
 * SynBerry 설문 시트와 섞지 않도록 새 스프레드시트에 따로 만든다.
 *
 * 배포 순서
 * 1. syn@synitart.com 계정으로 새 구글 시트를 만들고 이름을 "farm-web 방문 통계"로 한다.
 * 2. 확장 프로그램 → Apps Script → 이 파일 내용을 붙여 넣고 저장한다.
 * 3. 배포 → 새 배포 → 유형 "웹 앱", 실행 사용자 "나", 액세스 "모든 사용자" → 배포.
 * 4. 나온 웹 앱 주소(…/exec)를 farm-web/index.html의 STATS_EXEC 에 붙여 넣는다.
 */
const SHEET = '기록';
const HEAD = ['시각', '페이지', '사건', '유입경로', '패키지', '할인', '관리(만원)', '제작비(VAT포함)', '이전 사이트'];
const ALLOWED_T = ['view', 'pkg', 'copy'];
const ALLOWED_PKG = ['basic', 'standard', 'premium', 'custom', ''];
const ALLOWED_DISC = ['none', 'term', 'ref', ''];

function doPost(e) {
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const clean = (v, re, max) => {
      v = String(v == null ? '' : v).slice(0, max);
      return re.test(v) ? v : '';
    };
    const t = ALLOWED_T.indexOf(d.t) >= 0 ? d.t : '';
    if (!t || d.page !== 'farm-web') return out('skip');
    const row = [
      new Date(),
      'farm-web',
      t,
      clean(d.src, /^[A-Za-z0-9_-]{0,40}$/, 40) || 'direct',
      ALLOWED_PKG.indexOf(d.pkg || '') >= 0 ? (d.pkg || '') : '',
      ALLOWED_DISC.indexOf(d.disc || '') >= 0 ? (d.disc || '') : '',
      [0, 15, 30].indexOf(Number(d.maint)) >= 0 ? Number(d.maint) : '',
      Number.isFinite(Number(d.total)) && Number(d.total) > 0 && Number(d.total) < 1e9 ? Math.round(Number(d.total)) : '',
      clean(d.ref, /^[A-Za-z0-9.-]{0,60}$/, 60)
    ];
    const lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      sheet_().appendRow(row);
    } finally {
      lock.releaseLock();
    }
    return out('ok');
  } catch (err) {
    return out('error');
  }
}

/* 브라우저로 주소를 열면 유입경로·사건별 개수를 보여 준다(개인정보 없음). */
function doGet() {
  const rows = sheet_().getDataRange().getValues().slice(1);
  const count = {};
  rows.forEach(r => {
    const k = r[3] + ' · ' + r[2];
    count[k] = (count[k] || 0) + 1;
  });
  return ContentService.createTextOutput(JSON.stringify({ rows: rows.length, bySrcAndEvent: count }, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
  }
  return sh;
}

function out(s) {
  return ContentService.createTextOutput(s);
}
