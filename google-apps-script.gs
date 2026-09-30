/**
 * 원과 부채꼴의 숲 · 입체도형 무중력 연구소 → 구글 시트 기록 받기
 *
 * [설정 방법] (한 번만 하면 돼요)
 * 1. 구글 드라이브에서 새 구글 시트를 만들어요. (예: "수학 앱 학습 기록")
 * 2. 시트 메뉴 [확장 프로그램] → [Apps Script]를 열어요.
 * 3. 기본으로 들어 있는 코드를 모두 지우고, 이 파일 내용을 전부 붙여 넣은 뒤 저장해요.
 * 4. 오른쪽 위 [배포] → [새 배포] → 톱니바퀴에서 [웹 앱]을 골라요.
 *      - 다음 사용자 인증 정보로 실행: 나
 *      - 액세스 권한이 있는 사용자: 모든 사용자
 *    [배포]를 누르고 권한을 허용해요.
 * 5. 나오는 "웹 앱 URL"(https://script.google.com/macros/s/…/exec)을 복사해서
 *    index.html 과 solids.html 안의  const SHEET_URL="";  따옴표 사이에 붙여 넣어요.
 *
 * 코드를 고친 뒤에는 [배포] → [배포 관리] → 연필 → 버전 "새 버전"으로 다시 배포해야 반영돼요.
 *
 * [만들어지는 시트]
 *  - 기록        : 로그인과 채점이 한 줄씩 쌓여요.
 *  - 학생별 현황 : 앱·학번별로 한 줄, 맞힌 문제 수와 칭찬 도장을 최신 값으로 갱신해요.
 */

const LOG_SHEET = '기록';
const SUMMARY_SHEET = '학생별 현황';
const LOG_HEADER = ['받은 시각', '기록 시각', '앱', '학번', '이름', '종류', '문제', '문제 내용', '입력한 답', '정답 여부', '맞힌 문제 수(도토리)', '칭찬 도장'];
const SUMMARY_HEADER = ['앱', '학번', '이름', '맞힌 문제 수(도토리)', '칭찬 도장', '시도 횟수', '정답 횟수', '마지막 활동'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    const rows = Array.isArray(data.rows) ? data.rows.slice(0, 200) : [];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const log = getSheet_(ss, LOG_SHEET, LOG_HEADER);
    const summary = getSheet_(ss, SUMMARY_SHEET, SUMMARY_HEADER);
    const received = new Date();

    const logValues = rows.map(r => [
      received, r.time, r.app, r.id, r.name, r.type, r.q || '', r.text || '', r.ans || '', r.ok || '',
      r.acorns === '' || r.acorns === undefined ? '' : Number(r.acorns),
      r.stamps === '' || r.stamps === undefined ? '' : Number(r.stamps)
    ].map(safe_));
    if (logValues.length) {
      log.getRange(log.getLastRow() + 1, 1, logValues.length, LOG_HEADER.length).setValues(logValues);
    }
    rows.forEach(r => updateSummary_(summary, r));

    return json_({ ok: true, saved: rows.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ ok: true, message: '수학 앱 기록 수신 중' });
}

function updateSummary_(sheet, r) {
  const last = sheet.getLastRow();
  const keys = last > 1 ? sheet.getRange(2, 1, last - 1, 2).getValues() : [];
  let rowIndex = keys.findIndex(k => String(k[0]) === String(r.app) && String(k[1]) === String(r.id));
  const graded = r.type === '채점';
  if (rowIndex === -1) {
    sheet.appendRow([r.app, r.id, r.name, r.acorns, r.stamps, graded ? 1 : 0, graded && r.ok === 'O' ? 1 : 0, r.time].map(safe_));
    return;
  }
  rowIndex += 2;
  const cur = sheet.getRange(rowIndex, 1, 1, SUMMARY_HEADER.length).getValues()[0];
  sheet.getRange(rowIndex, 1, 1, SUMMARY_HEADER.length).setValues([[
    r.app, r.id, r.name, r.acorns, r.stamps,
    Number(cur[5] || 0) + (graded ? 1 : 0),
    Number(cur[6] || 0) + (graded && r.ok === 'O' ? 1 : 0),
    r.time
  ].map(safe_)]);
}

function getSheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#e8f1dc');
    sh.setFrozenRows(1);
  }
  return sh;
}

// 학생이 입력한 글자가 수식(=, +, -, @로 시작)으로 실행되지 않게 막아요.
function safe_(v) {
  if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
  return v;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
