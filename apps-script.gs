/**
 * 중화 적정 실험실 — Google 스프레드시트「acid–base titration」 연동 (v2: 수행평가 · 초기화 · 비밀번호 변경)
 * 1) 스프레드시트에서 확장 프로그램 → Apps Script 를 열고 이 코드를 통째로 붙여 넣기 (기존 코드는 지우기)
 * 2) 배포 → 새 배포 (이미 배포했다면: 배포 관리 → 편집(연필) → 버전: 새 버전)
 *    유형: 웹 앱 / 실행: 나 / 액세스: 모든 사용자 → 배포 → 웹 앱 URL 복사
 * 3) index.html 맨 위 SHEET_URL 에 URL 붙여 넣기
 * 교사 비밀번호 처음 값은 0405 입니다. 실험실의 교사 화면 → [비밀번호 변경]에서 바꿀 수 있어요.
 */
const DEFAULT_PW = '0405';

const HEAD = ['서버시간','기록ID','학생시간','학년','반','번호','이름','구분','항목','점수','상세'];
const RHEAD = ['서버시간','기록ID','학년','반','번호','이름','단계','시료','적정액','지시약','종말점 색 변화',
  '적정액 농도(M)','시료 부피(mL)','1회(mL)','2회(mL)','3회(mL)','평균 부피(mL)','학생 계산 농도(M)','실제 농도(M)',
  '상대오차(%)','식초 함량(%)','오차 원인','고찰','심화 실습(40)','결과 보고서(20)'];
const AHEAD = ['학생키','결과 보고서 조정 점수','수정 시각'];

function pw_() { return PropertiesService.getScriptProperties().getProperty('TEACHER_PW') || DEFAULT_PW; }

function sheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(head);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#F7E4F0');
  }
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const r = JSON.parse(e.postData.contents);
    const sh = sheet_('결과', HEAD);
    if (r.id && sh.getLastRow() > 1 &&
        sh.getRange(2, 2, sh.getLastRow() - 1, 1).createTextFinder(r.id).matchEntireCell(true).findNext()) {
      return json_({ ok: true, dup: true });
    }
    const now = new Date();
    sh.appendRow([now, r.id, r.clientTime, r.grade, r.cls, r.no, r.name, r.kind, r.item, r.score, r.detail]);
    if (r.report) {
      const p = r.report;
      sheet_('보고서', RHEAD).appendRow([now, r.id, r.grade, r.cls, r.no, r.name, p.stage + '단계', p.acid, p.titrant,
        p.ind, p.change, p.ct, p.va, p.v1, p.v2, p.v3, p.vbar, p.myC, p.trueC, p.relErr, p.pct, p.errs, p.disc, p.lab40, p.rep20]);
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const p = e.parameter || {};
  if (p.key !== pw_()) return json_({ ok: false, error: '비밀번호가 맞지 않아요' });
  const action = p.action || 'data';
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (action === 'data') return json_({ ok: true, rows: rows_(), adj: adj_(), sheetUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl() });
    if (action === 'setpw') {
      const n = String(p.newpw || '');
      if (n.length < 4 || n.length > 20) return json_({ ok: false, error: '4~20자로 입력하세요' });
      PropertiesService.getScriptProperties().setProperty('TEACHER_PW', n);
      return json_({ ok: true });
    }
    if (action === 'reset') {
      const t = String(p.target || '');
      if (!t) return json_({ ok: false, error: '대상이 없어요' });
      [['결과', HEAD, 3], ['보고서', RHEAD, 2]].forEach(([name, head, gcol]) => {
        const sh = sheet_(name, head);
        const n = sh.getLastRow() - 1;
        if (n <= 0) return;
        const v = sh.getRange(2, 1, n, head.length).getValues();
        const keep = t === 'all' ? [] : v.filter(x => [x[gcol], x[gcol + 1], x[gcol + 2], x[gcol + 3]].join('|') !== t);
        sh.getRange(2, 1, n, head.length).clearContent();
        if (keep.length) sh.getRange(2, 1, keep.length, head.length).setValues(keep);
      });
      const a = sheet_('교사조정', AHEAD); const n = a.getLastRow() - 1;
      if (n > 0) {
        const v = a.getRange(2, 1, n, 3).getValues();
        const keep = t === 'all' ? [] : v.filter(x => x[0] !== t);
        a.getRange(2, 1, n, 3).clearContent();
        if (keep.length) a.getRange(2, 1, keep.length, 3).setValues(keep);
      }
      return json_({ ok: true });
    }
    if (action === 'adjust') {
      const a = sheet_('교사조정', AHEAD); const n = a.getLastRow() - 1;
      const v = n > 0 ? a.getRange(2, 1, n, 3).getValues() : [];
      const i = v.findIndex(x => x[0] === p.skey);
      if (p.score === '' || p.score == null) { if (i >= 0) a.deleteRow(i + 2); }
      else if (i >= 0) a.getRange(i + 2, 2, 1, 2).setValues([[Number(p.score), new Date()]]);
      else a.appendRow([p.skey, Number(p.score), new Date()]);
      return json_({ ok: true });
    }
    return json_({ ok: false, error: '알 수 없는 요청' });
  } finally {
    lock.releaseLock();
  }
}

function rows_() {
  const sh = sheet_('결과', HEAD); const n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  return sh.getRange(2, 1, n, HEAD.length).getValues().filter(x => x[1] !== '').map(x => ({
    time: x[0], id: x[1], clientTime: x[2], grade: String(x[3]), cls: String(x[4]), no: String(x[5]),
    name: String(x[6]), kind: x[7], item: x[8], score: x[9], detail: x[10] }));
}
function adj_() {
  const a = sheet_('교사조정', AHEAD); const n = a.getLastRow() - 1; const o = {};
  if (n > 0) a.getRange(2, 1, n, 2).getValues().forEach(x => { if (x[0] !== '') o[x[0]] = x[1]; });
  return o;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
