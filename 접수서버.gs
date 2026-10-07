/**
 * 돈암1동 이웃살핌 — 접수 서버
 * 길음종합사회복지관 마을동행팀
 *
 * 앱에서 보낸 접수 내용을 구글 시트에 한 줄씩 쌓습니다.
 * 시트는 처음 접수가 들어올 때 자동으로 만들어집니다.
 *
 * ── 설치 방법 ────────────────────────────────────────────
 * 1. script.google.com 접속 (반드시 기관 구글 계정으로)
 * 2. [새 프로젝트]
 * 3. 코드 창을 다 지우고 이 파일 전체를 붙여넣기
 * 4. 위쪽 [저장]
 *
 * 5. 오른쪽 위 [배포] → [새 배포]
 * 6. 왼쪽 톱니바퀴 → [웹 앱] 선택
 * 7. 설명       : 이웃살핌 접수
 *    실행 계정  : 나
 *    액세스 권한: 모든 사용자        ← 반드시 이걸로
 * 8. [배포] → 권한 요청이 뜨면
 *      계정 선택 → [고급] → [(안전하지 않은 페이지)로 이동] → [허용]
 * 9. 나오는 "웹 앱 URL" 을 복사
 *      https://script.google.com/macros/s/AKfycb..../exec
 *
 * 10. index.html 의 ENDPOINT 에 그 주소를 붙여넣고 Commit
 * 11. sw.js 의 VERSION 숫자를 올리고 Commit
 *
 * ── 시트 주소 확인 ───────────────────────────────────────
 * 첫 접수가 들어온 뒤, 함수 목록에서 [시트주소보기] 를 골라 실행하면
 * 실행 로그에 시트 주소가 나옵니다.
 *
 * ── 코드를 고친 뒤에는 ───────────────────────────────────
 * [배포] → [배포 관리] → 연필 → 버전 [새 버전] → [배포]
 * 새 배포를 만들면 주소가 바뀌니, 반드시 기존 배포를 수정하세요.
 * ─────────────────────────────────────────────────────── */

var 시트파일이름 = '우리동네 이웃살핌 접수시트';
var 탭이름       = '접수목록';

var 머리글 = [
  '접수일시', '동', '걱정되는 분 성함', '사시는 곳', '연락처', '연세',
  '본인 연락 동의', '추천 이유', '자세한 내용',
  '추천하신 분', '추천하신 분 연락처',
  '접수확인', '조치', '조치일', '결과', '추천인 통보', '비고'
];


/* ── 앱이 접수를 보낼 때 실행됩니다 ───────────────────── */
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);

    // 아주 단순한 장난 방지 — 내용이 비면 받지 않습니다
    if (!d.reporter && !d.name && !d.addr) {
      return 응답({ ok: false, error: '내용이 비어 있습니다' });
    }

    var sheet = 시트가져오기();
    sheet.appendRow([
      new Date(),
      d.dong        || '',
      d.name        || '',
      d.addr        || '',
      d.tel         || '',
      d.age         || '',
      d.consent     || '',
      d.reasons     || '',
      d.detail      || '',
      d.reporter    || '',
      d.reporterTel || '',
      '', '', '', '', '', ''      // 처리대장 칸은 담당자가 채웁니다
    ]);

    return 응답({ ok: true });

  } catch (err) {
    return 응답({ ok: false, error: String(err) });
  }
}


/* ── 서버가 살아있는지 확인용 ─────────────────────────── */
function doGet() {
  return 응답({ ok: true, msg: '돈암1동 이웃살핌 접수 서버가 작동 중입니다.' });
}


/* ── 시트 준비 (없으면 만들고, 있으면 그대로 씁니다) ──── */
function 시트가져오기() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  var ss = null;

  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create(시트파일이름);
    props.setProperty('SHEET_ID', ss.getId());
  }

  var sheet = ss.getSheetByName(탭이름);
  if (!sheet) {
    sheet = (ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() === 0)
      ? ss.getSheets()[0].setName(탭이름)
      : ss.insertSheet(탭이름);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(머리글);
    꾸미기(sheet);
  } else {
    동칸만들기(sheet);          // 예전 시트를 쓰던 경우 '동' 칸을 끼워 넣습니다
  }
  return sheet;
}


/* ── 꾸미기 ──────────────────────────────────────────── */
function 꾸미기(sheet) {
  sheet.getRange(1, 1, 1, 머리글.length)
       .setFontWeight('bold').setBackground('#EEF4EF');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(2, 90);    // 동
  sheet.setColumnWidth(4, 200);   // 사시는 곳
  sheet.setColumnWidth(8, 240);   // 추천 이유
  sheet.setColumnWidth(9, 320);   // 자세한 내용
}


/* ── 예전 시트에 '동' 칸 끼워 넣기 (한 번만 실행됩니다) ─ */
function 동칸만들기(sheet) {
  var head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (head.indexOf('동') !== -1) return;      // 이미 있으면 그냥 둡니다

  sheet.insertColumnBefore(2);                // B열에 빈 칸을 하나 넣고
  sheet.getRange(1, 2).setValue('동');        // 머리글을 적습니다
  꾸미기(sheet);                               // 기존 줄은 '동'이 비어 있게 됩니다
}


/* ── 시트 주소 확인 (함수 목록에서 골라 실행) ─────────── */
function 시트주소보기() {
  var sheet = 시트가져오기();
  var url = sheet.getParent().getUrl();
  Logger.log('시트 주소: ' + url);
  Logger.log('');
  Logger.log('※ 이 시트는 절대 "웹에 게시" 하지 마세요.');
  Logger.log('   [공유] 를 "제한됨" 으로 두고 담당자만 초대하세요.');
  return url;
}


/* ── 새 접수가 오면 메일 받기 (선택) ──────────────────
   아래 주소를 바꾼 뒤, 이 함수 안의 주석을 풀고
   doPost 안에서 알림보내기(d) 를 호출하세요.               */
var 알림받을주소 = 'your-name@example.org';

function 알림보내기(d) {
  MailApp.sendEmail(
    알림받을주소,
    '[이웃살핌] 새 접수 — ' + (d.addr || d.name || ''),
    '추천하신 분: ' + (d.reporter || '') + '\n' +
    '걱정되는 분: ' + (d.name || '') + '\n' +
    '사시는 곳  : ' + (d.addr || '') + '\n' +
    '추천 이유  : ' + (d.reasons || '') + '\n\n' +
    (d.detail || '') + '\n\n' +
    '시트에서 확인하세요: ' + 시트가져오기().getParent().getUrl()
  );
}


function 응답(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
