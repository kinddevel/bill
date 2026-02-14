/**
 * 사내 결제·사용 처리 시스템 v2.2
 */

const SHEET = { EXPENSE: '지출내역', USERS: '사용자' };
const ROLE = { USER: '사용자', APPROVER: '승인자', ADMIN: '관리자' };
const STATUS = {
  SUBMITTED: '제출',
  APPROVER_OK: '승인자승인',
  FINAL_OK: '승인(최종)',
  REJECT: '반려'
};
const NOTIFY_EMAIL = 'jihye@phosem.com';

const HEADERS = [
  '번호', '등록일시', '사용일', '요청자이메일', '요청자이름',
  '사용처', '금액', '결제수단', '분류', '메모',
  '상태', '승인자이메일', '최종처리일시', '영수증', '사진',
  '세금계산서', '거래명세서', '견적서', '발주서', '삭제여부', '반려사유'
];

function run_bindSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());
  _ensureSheets_();
  SpreadsheetApp.getUi().alert('연결 완료! 이제 웹앱을 실행하세요.');
}

function _ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function _ensureSheets_() {
  const ss = _ss_();
  let expense = ss.getSheetByName(SHEET.EXPENSE);
  if (!expense) expense = ss.insertSheet(SHEET.EXPENSE);

  const currentHeaders = expense.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (String(currentHeaders[0]).trim() !== HEADERS[0]) {
    expense.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }

  if (!ss.getSheetByName(SHEET.USERS)) {
    const users = ss.insertSheet(SHEET.USERS);
    users.getRange(1, 1, 1, 4).setValues([['이메일', '구글이메일', '이름', '권한']]);
  }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('사내결제')
    .addItem('1) 시스템 연결(필수)', 'run_bindSpreadsheet')
    .addToUi();
}

function 현재사용자_() {
  const googleEmail = String(Session.getActiveUser().getEmail() || '').toLowerCase().trim();
  const sh = _ss_().getSheetByName(SHEET.USERS);
  const data = sh ? sh.getDataRange().getValues() : [];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).toLowerCase().trim() === googleEmail) {
      return { 이메일: data[i][0], 구글이메일: data[i][1], 이름: data[i][2], 권한: data[i][3], 매핑: true };
    }
  }
  return { 이메일: googleEmail, 구글이메일: googleEmail, 이름: '미등록사용자', 권한: ROLE.USER, 매핑: false };
}

function doGet() {
  _ensureSheets_();
  const t = HtmlService.createTemplateFromFile('HTML');
  t.현재사용자 = 현재사용자_();
  return t.evaluate().setTitle('사내 결제 시스템').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function API_데이터로드() {
  const u = 현재사용자_();
  const sh = _ss_().getSheetByName(SHEET.EXPENSE);
  if (!sh || sh.getLastRow() < 2) {
    return { ok: true, summary: { 제출: 0, 승인: 0, 반려: 0 }, myRows: [], allRows: [], role: u.권한 };
  }

  const data = sh.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][19] === true) continue;
    rows.push({
      번호: data[i][0], 등록일시: data[i][1], 사용일: data[i][2], 요청자이메일: data[i][3], 요청자이름: data[i][4],
      사용처: data[i][5], 금액: data[i][6], 결제수단: data[i][7], 분류: data[i][8], 메모: data[i][9],
      상태: data[i][10], 승인자이메일: data[i][11], 최종처리일시: data[i][12], 반려사유: data[i][20],
      파일: { 영수증: data[i][13], 사진: data[i][14], 세금계산서: data[i][15], 거래명세서: data[i][16], 견적서: data[i][17], 발주서: data[i][18] }
    });
  }

  const myRows = rows
    .filter(r => String(r.요청자이메일).toLowerCase() === u.이메일.toLowerCase() || String(r.요청자이메일).toLowerCase() === u.구글이메일.toLowerCase())
    .sort((a, b) => new Date(b.등록일시).getTime() - new Date(a.등록일시).getTime());

  return {
    ok: true,
    role: u.권한,
    summary: {
      제출: myRows.filter(r => r.상태 === STATUS.SUBMITTED || r.상태 === STATUS.APPROVER_OK).length,
      승인: myRows.filter(r => r.상태 === STATUS.FINAL_OK).length,
      반려: myRows.filter(r => String(r.상태).includes(STATUS.REJECT)).length
    },
    myRows,
    allRows: (u.권한 !== ROLE.USER) ? rows : []
  };
}

function API_지출저장(p) {
  p = p || {};
  p.파일 = p.파일 || {};
  const u = 현재사용자_();
  const sh = _ss_().getSheetByName(SHEET.EXPENSE);
  let rowIdx = -1;
  let nextId = Number(p.번호);
  let prevFiles = { 영수증: '', 사진: '', 세금계산서: '', 거래명세서: '', 견적서: '', 발주서: '' };

  if (p.번호) {
    const ids = sh.getRange('A:A').getValues().flat();
    rowIdx = ids.indexOf(Number(p.번호)) + 1;
    if (rowIdx > 1) {
      const prev = sh.getRange(rowIdx, 1, 1, HEADERS.length).getValues()[0];
      if (String(prev[3]).toLowerCase() !== u.이메일.toLowerCase() && u.권한 === ROLE.USER) {
        return { ok: false, error: '본인 데이터만 수정할 수 있습니다.' };
      }
      prevFiles = { 영수증: prev[13], 사진: prev[14], 세금계산서: prev[15], 거래명세서: prev[16], 견적서: prev[17], 발주서: prev[18] };
    }
  } else {
    const ids = sh.getRange('A:A').getValues().flat().filter(Number);
    nextId = ids.length ? Math.max(...ids) + 1 : 1;
  }

  const finalFiles = {
    영수증: p.파일.영수증 || prevFiles.영수증,
    사진: p.파일.사진 || prevFiles.사진,
    세금계산서: p.파일.세금계산서 || prevFiles.세금계산서,
    거래명세서: p.파일.거래명세서 || prevFiles.거래명세서,
    견적서: p.파일.견적서 || prevFiles.견적서,
    발주서: p.파일.발주서 || prevFiles.발주서
  };

  const row = [
    nextId, new Date(), p.사용일, u.이메일, u.이름, p.사용처, Number(p.금액 || 0), p.결제수단, p.분류, p.메모,
    STATUS.SUBMITTED, '', '',
    finalFiles.영수증, finalFiles.사진, finalFiles.세금계산서, finalFiles.거래명세서, finalFiles.견적서, finalFiles.발주서,
    false, ''
  ];

  if (rowIdx > 1) sh.getRange(rowIdx, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);

  _메일발송_(u, row, finalFiles);
  return { ok: true };
}

function API_승인처리(p) {
  const u = 현재사용자_();
  if (u.권한 === ROLE.USER) return { ok: false, error: '권한이 없습니다.' };

  const sh = _ss_().getSheetByName(SHEET.EXPENSE);
  const ids = sh.getRange('A:A').getValues().flat();
  const rowIdx = ids.indexOf(Number(p.번호)) + 1;
  if (rowIdx <= 1) return { ok: false, error: '내역 찾기 실패' };

  const row = sh.getRange(rowIdx, 1, 1, HEADERS.length).getValues()[0];
  const current = row[10];
  let status = current;

  if (p.결정 === STATUS.REJECT) {
    status = STATUS.REJECT;
  } else if (u.권한 === ROLE.ADMIN) {
    status = STATUS.FINAL_OK;
  } else if (u.권한 === ROLE.APPROVER && current === STATUS.SUBMITTED) {
    status = STATUS.APPROVER_OK;
  }

  sh.getRange(rowIdx, 11).setValue(status);
  sh.getRange(rowIdx, 12).setValue(u.이메일 || u.구글이메일);
  if (p.사유) sh.getRange(rowIdx, 21).setValue(p.사유);
  if (status === STATUS.FINAL_OK) sh.getRange(rowIdx, 13).setValue(new Date());

  return { ok: true, status };
}

function API_파일업로드(p) {
  const u = 현재사용자_();
  const root = _getOrCreateFolder_(DriveApp.getRootFolder(), '사내결제사진');
  const userFolder = _getOrCreateFolder_(root, u.이름 || u.구글이메일 || '미분류');

  const ext = _extractExt_(p.fileName);
  const prefix = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const cleanName = _sanitizeFileName_(p.fileName.replace(/\.[^.]+$/, ''));
  const finalName = `${prefix}_${cleanName}${ext}`;

  const blob = Utilities.newBlob(Utilities.base64Decode(p.base64), p.mimeType, finalName);
  const file = userFolder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return { ok: true, url: file.getUrl(), name: finalName };
}

function _getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function _extractExt_(fileName) {
  const m = String(fileName).match(/(\.[^.]+)$/);
  return m ? m[1] : '';
}

function _sanitizeFileName_(name) {
  return String(name).replace(/[\\/:*?"<>|#%{}~&]/g, '_').trim() || 'file';
}

function _메일발송_(u, row, files) {
  const links = Object.keys(files)
    .filter(k => files[k])
    .map(k => `- ${k}: ${files[k]}`)
    .join('\n');

  const body = [
    '사내 결제 등록 알림',
    `요청자: ${u.이름} (${u.이메일 || u.구글이메일})`,
    `사용일: ${row[2]}`,
    `사용처: ${row[5]}`,
    `금액: ${Number(row[6]).toLocaleString()}원`,
    `분류: ${row[8]}`,
    '',
    '[증빙 링크]',
    links || '- 없음'
  ].join('\n');

  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: `[사내결제] ${u.이름}님 지출 등록 (${Number(row[6]).toLocaleString()}원)`,
    body
  });
}
