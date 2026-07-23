// 2026 생성형 AI(LLM) 1일 집중 특강(상주) 신청·관리 백엔드 — 구글시트 저장 + 신청자 이메일 발송 + 관리자 API

// ===== 설정 =====
// 관리자 비밀번호는 코드에 적지 않고 Apps Script의 Script Properties(비공개 설정값)에 저장한다.
// 키: ADMIN_PASSWORD, 값: isidor-llm-2026  → 프로젝트 설정 > 스크립트 속성에서 입력 (SETUP-backend.md 참조)
const CAPACITY = 15;                               // 정원 (참석확정 인원 기준)
const FROM_NAME = '유소장닷컴';        // 발신자 표시 이름 (실제 발신 주소는 스크립트 소유 계정 = isidor.yu@gmail.com)
const CONTACT = 'isidor.yu@gmail.com';             // 신청자 안내 메일의 문의처 + 문의사항 알림 수신
const COURSE_TITLE = '생성형 AI(LLM), 하루 만에 일에 쓰는 법 — 1일 집중 특강';
const COURSE_DATE = '2026년 8월 (세부 일정 추후 안내)';
const COURSE_PLACE = '상주 (세부 장소는 추후 안내)';
const EXPERIENCES = ['처음이다', '써 본 적 있다', '업무에 자주 쓴다']; // 신청 폼 라디오와 일치
// ================================

function getAdminPassword_() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD') || '';
}

function authOk_(pw) {
  const real = getAdminPassword_();
  return real && pw === real; // 설정값이 비어 있으면 항상 거부
}

const HEADERS = ['접수ID', '신청일시', '성명', '소속', '직위', '이메일', '연락처', 'AI사용경험', '문의', '상태', '참석확정일시'];

function getSheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SHEET_ID');
  let ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { id = null; } }
  if (!id) {
    ss = SpreadsheetApp.create('2026 LLM 특강 신청자');
    props.setProperty('SHEET_ID', ss.getId());
  }
  let sh = ss.getSheetByName('신청자');
  if (!sh) {
    sh = ss.getSheets()[0];
    sh.setName('신청자');
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  // 비공개 관리자 페이지: 공개 사이트가 아닌 이 배포 URL(?page=admin)로만 제공된다.
  if (params.page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('Admin')
      .setTitle('LLM 특강 관리자')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  const action = params.action || 'status';
  if (action === 'status') return json_(publicStatus_());
  return json_({ ok: false, error: 'unknown' });
}

// 관리자 페이지(HtmlService)에서 google.script.run으로 호출하는 API
function adminList(pw) { return handleList_({ pw: pw }); }
function adminConfirm(pw, row) { return handleConfirm_({ pw: pw, row: row }); }
function adminUnconfirm(pw, row) { return handleUnconfirm_({ pw: pw, row: row }); }
function adminDeleteMany(pw, ids) { return handleDelete_({ pw: pw, ids: ids }); }
function adminUpdate(pw, id, data) { return handleUpdate_({ pw: pw, id: id, data: data }); }

function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad json' }); }
  if (p.action === 'apply') return json_(handleApply_(p));
  if (p.action === 'list') return json_(handleList_(p));
  if (p.action === 'confirm') return json_(handleConfirm_(p));
  if (p.action === 'unconfirm') return json_(handleUnconfirm_(p));
  if (p.action === 'delete') return json_(handleDelete_(p));
  return json_({ ok: false, error: 'unknown action' });
}

// 상태별 인원 집계 — 참석확정 인원이 정원 판정 기준
function countByStatus_() {
  const rows = getSheet_().getDataRange().getValues();
  const c = { 접수: 0, 확정: 0, 대기: 0 };
  for (let i = 1; i < rows.length; i++) {
    const s = rows[i][9];
    if (s === '참석확정') c.확정++;
    else if (s === '대기') c.대기++;
    else c.접수++;
  }
  return c;
}

function publicStatus_() {
  const c = countByStatus_();
  const remaining = Math.max(0, CAPACITY - c.확정);
  return {
    ok: true, capacity: CAPACITY,
    confirmed: c.확정, applied: c.접수, waiting: c.대기,
    remaining: remaining, full: remaining <= 0,
  };
}

function handleApply_(p) {
  const name = (p['성명'] || '').toString().trim();
  const org = (p['소속'] || '').toString().trim();
  const title = (p['직위'] || '').toString().trim();
  const email = (p['이메일'] || '').toString().trim();
  const phone = formatPhone_(p['연락처'] || '');
  const exp = (p['AI사용경험'] || '').toString().trim();
  const note = (p['문의'] || '').toString().trim();
  if (!name || !org || !email || !phone) return { ok: false, error: '필수 항목이 누락되었습니다.' };
  if (EXPERIENCES.indexOf(exp) === -1) return { ok: false, error: '생성형 AI 사용 경험을 선택해 주세요.' };
  // 개인정보 수집·이용 동의는 서버에서도 재검증한다 (브라우저 검증 우회 차단)
  if ((p['개인정보동의'] || '').toString().trim() !== '동의') {
    return { ok: false, error: '개인정보 수집·이용에 동의해야 신청할 수 있습니다.' };
  }

  // 정원 검사와 행 추가를 락으로 직렬화 — 동시 제출로 상태 판정이 어긋나는 것을 방지
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { ok: false, error: '접수가 몰려 처리가 지연되고 있습니다. 잠시 후 다시 시도해 주세요.' }; }
  let waitlisted;
  let remaining;
  try {
    const c = countByStatus_();
    remaining = Math.max(0, CAPACITY - c.확정);
    waitlisted = remaining <= 0; // 참석확정 인원이 정원에 도달하면 이후 신청자는 자동 대기자
    const id = 'A' + new Date().getTime();
    getSheet_().appendRow([id, new Date(), name, org, title, email, phone, exp, note, waitlisted ? '대기' : '신청접수', '']);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  try {
    if (waitlisted) sendWaitEmail_(email, name);
    else sendApplyEmail_(email, name);
  } catch (e) { /* 메일 실패해도 접수는 유지 */ }
  // 문의사항이 있으면 운영자에게 알림 메일 전달
  if (note) {
    try {
      GmailApp.sendEmail(CONTACT, '[LLM 특강] 신청 문의사항 — ' + name,
        '신청서의 문의사항 내용을 전달드립니다.\n\n'
        + '· 성명: ' + name + '\n'
        + '· 소속: ' + org + (title ? ' / ' + title : '') + '\n'
        + '· 이메일: ' + email + '\n'
        + '· 연락처: ' + phone + '\n'
        + '· AI 사용 경험: ' + exp + '\n'
        + '· 접수 상태: ' + (waitlisted ? '대기' : '신청접수') + '\n\n'
        + '--- 문의사항 ---\n' + note + '\n',
        { name: FROM_NAME });
    } catch (e) { /* 알림 실패해도 접수는 유지 */ }
  }
  return { ok: true, waitlisted: waitlisted, remaining: waitlisted ? 0 : remaining };
}

function handleList_(p) {
  if (!authOk_(p.pw)) return { ok: false, error: 'auth' };
  const rows = getSheet_().getDataRange().getValues();
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    list.push({
      row: i + 1, id: r[0], 신청일시: fmt_(r[1]), 성명: r[2], 소속: r[3], 직위: r[4] || '',
      이메일: r[5], 연락처: formatPhone_(r[6]), AI사용경험: r[7] || '', 문의: r[8] || '',
      상태: r[9], 참석확정일시: fmt_(r[10]),
    });
  }
  const c = countByStatus_();
  const summary = {
    정원: CAPACITY, 신청접수: c.접수, 참석확정: c.확정, 대기: c.대기,
    잔여: Math.max(0, CAPACITY - c.확정),
  };
  return { ok: true, summary: summary, list: list };
}

function handleConfirm_(p) {
  if (!authOk_(p.pw)) return { ok: false, error: 'auth' };
  const sh = getSheet_();
  const row = parseInt(p.row, 10);
  if (!row || row < 2) return { ok: false, error: 'bad row' };
  const data = sh.getRange(row, 1, 1, HEADERS.length).getValues()[0];
  sh.getRange(row, 10).setValue('참석확정');
  sh.getRange(row, 11).setValue(new Date());
  SpreadsheetApp.flush(); // 즉시 커밋 — 직후 리스트 재조회가 이전 값을 읽는 문제 방지
  try { sendConfirmEmail_(data[5], data[2]); } catch (e) {}
  return { ok: true };
}

function handleUnconfirm_(p) {
  if (!authOk_(p.pw)) return { ok: false, error: 'auth' };
  const sh = getSheet_();
  const row = parseInt(p.row, 10);
  if (!row || row < 2) return { ok: false, error: 'bad row' };
  sh.getRange(row, 10).setValue('신청접수');
  sh.getRange(row, 11).setValue('');
  SpreadsheetApp.flush();
  return { ok: true };
}

function handleDelete_(p) {
  if (!authOk_(p.pw)) return { ok: false, error: 'auth' };
  const ids = p.ids || (p.id ? [p.id] : []);
  if (!ids.length) return { ok: false, error: 'bad id' };
  const sh = getSheet_();
  const rows = sh.getDataRange().getValues();
  const targets = [];
  for (let i = 1; i < rows.length; i++) {
    if (ids.indexOf(String(rows[i][0])) !== -1) targets.push(i + 1); // 접수ID(A열)로 매칭 → 행 번호 이동에 안전
  }
  if (!targets.length) return { ok: false, error: 'not found' };
  targets.sort(function (a, b) { return b - a; }); // 아래 행부터 삭제해야 행 번호가 밀리지 않음
  targets.forEach(function (r) { sh.deleteRow(r); });
  SpreadsheetApp.flush();
  return { ok: true, deleted: targets.length };
}

// 신청 정보 수정 — 접수ID로 행을 찾아 성명·소속·직위·이메일·연락처·AI사용경험·문의를 갱신
function handleUpdate_(p) {
  if (!authOk_(p.pw)) return { ok: false, error: 'auth' };
  const id = (p.id || '').toString();
  const d = p.data || {};
  const name = (d['성명'] || '').toString().trim();
  const org = (d['소속'] || '').toString().trim();
  const title = (d['직위'] || '').toString().trim();
  const email = (d['이메일'] || '').toString().trim();
  const phone = formatPhone_(d['연락처'] || '');
  const exp = (d['AI사용경험'] || '').toString().trim();
  const note = (d['문의'] || '').toString().trim();
  if (!id) return { ok: false, error: 'bad id' };
  if (!name || !email) return { ok: false, error: '성명·이메일은 비울 수 없습니다.' };
  if (exp && EXPERIENCES.indexOf(exp) === -1) return { ok: false, error: '유효하지 않은 AI 사용 경험 값입니다.' };
  const sh = getSheet_();
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === id) {
      sh.getRange(i + 1, 3, 1, 7).setValues([[name, org, title, email, phone, exp, note]]);
      SpreadsheetApp.flush();
      return { ok: true };
    }
  }
  return { ok: false, error: 'not found' };
}

function fmt_(d) {
  if (!d) return '';
  if (Object.prototype.toString.call(d) !== '[object Date]') return d.toString();
  return Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
}

// 전화번호를 010-0000-0000 형식으로 정규화한다. 숫자만 추출하고, 엑셀/숫자 변환으로
// 앞자리 0이 떨어진 휴대폰 번호(예: 1058307048)는 0을 복원해 011 형식으로 맞춘다.
function formatPhone_(raw) {
  var d = String(raw == null ? '' : raw).replace(/\D/g, '');
  if (!d) return '';
  if (d.length === 10 && d.charAt(0) === '1') d = '0' + d;                 // 1058307048 → 01058307048
  if (d.length === 11) return d.replace(/(\d{3})(\d{4})(\d{4})/, '$1-$2-$3'); // 휴대폰 010-0000-0000
  if (d.length === 10 && d.indexOf('02') === 0) return d.replace(/(\d{2})(\d{4})(\d{4})/, '$1-$2-$3'); // 서울 02-0000-0000
  if (d.length === 10) return d.replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3'); // 지역번호 000-000-0000
  if (d.length === 9 && d.indexOf('02') === 0) return d.replace(/(\d{2})(\d{3})(\d{4})/, '$1-$2-$3'); // 서울 02-000-0000
  return d; // 형식을 알 수 없으면 숫자만 반환
}

// ===== 이메일 (발신 주소 = 스크립트 소유 계정) =====
function courseInfoBlock_() {
  return '· 과정: ' + COURSE_TITLE + '\n'
    + '· 일시: ' + COURSE_DATE + '\n'
    + '· 장소: ' + COURSE_PLACE + '\n'
    + '· 참가비: 82,500원/인 (VAT 포함) — 정가 165,000원, 7·8월 론칭 DC 50%\n'
    + '· 준비물: 개인 노트북 + 충전기 (반드시 지참)\n';
}

function sendApplyEmail_(email, name) {
  GmailApp.sendEmail(email, '[LLM 특강] 신청이 접수되었습니다',
    applyBody_(name), { name: FROM_NAME });
}

function sendWaitEmail_(email, name) {
  GmailApp.sendEmail(email, '[LLM 특강] 대기자로 등록되었습니다',
    waitBody_(name), { name: FROM_NAME });
}

function sendConfirmEmail_(email, name) {
  GmailApp.sendEmail(email, '[LLM 특강] 참석이 확정되었습니다',
    confirmBody_(name), { name: FROM_NAME });
}

function prepBlock_() {
  return '[교육 전 준비사항]\n'
    + '1. 개인 노트북 + 충전기 (종일 실습 — 배터리만으로는 부족합니다)\n'
    + '2. 최신 크롬(또는 엣지) 브라우저 설치\n'
    + '3. 클로드 에이전트(claude.ai) 사전 가입 (가입 시 휴대폰 문자 인증 필요)\n'
    + '4. 클로드 에이전트 유료 구독 — 권장 플랜: 클로드 Pro (월 US$20, 연 결제 시 월 US$17 수준)\n'
    + '   무료 계정은 사용량 한도로 종일 실습이 어렵습니다.\n'
    + '   구독료·사용량 초과 등 클로드 에이전트 이용 비용은 모두 수강자 본인 부담입니다.\n'
    + '5. 실습용 업무 파일 2~3개 (한글·PDF·엑셀, 개인정보 없는 파일 / .hwp는 PDF 변환 권장)\n'
    + '6. 기관 지급 노트북은 claude.ai 접속 차단 여부를 사전 확인\n';
}

function applyBody_(name) {
  return name + ' 님, 안녕하세요.\n\n'
    + '"' + COURSE_TITLE + '" 신청이 정상적으로 접수되었습니다.\n'
    + '참석이 확정되면 확정 안내 메일을 다시 보내드립니다.\n\n'
    + courseInfoBlock_() + '\n'
    + prepBlock_() + '\n'
    + '문의: ' + CONTACT + '\n유소장.컴';
}

function waitBody_(name) {
  return name + ' 님, 안녕하세요.\n\n'
    + '"' + COURSE_TITLE + '" 신청이 접수되었으나,\n'
    + '현재 정원(' + CAPACITY + '명)이 마감되어 대기자로 등록되었습니다.\n'
    + '결원이 생기면 등록 순서대로 참석 확정 안내를 드리겠습니다.\n\n'
    + courseInfoBlock_() + '\n'
    + '문의: ' + CONTACT + '\n유소장.컴';
}

function confirmBody_(name) {
  return name + ' 님, 안녕하세요.\n\n'
    + '"' + COURSE_TITLE + '" 참석이 확정되었습니다.\n'
    + '아래 일정을 확인하시고 당일 준비물을 꼭 챙겨 오시기 바랍니다.\n\n'
    + courseInfoBlock_() + '\n'
    + prepBlock_() + '\n'
    + '※ 세부 장소는 교육일 전 별도 안내드립니다.\n\n'
    + '문의: ' + CONTACT + '\n유소장.컴';
}
