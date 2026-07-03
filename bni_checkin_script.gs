// BNI Pioneer 출석체크 - Google Apps Script (JSONP 지원)
// 명단 탭: A=이름, B=전화번호뒷4자리, C=구분(멤버/비지터/대리), D=대리대상, E=호스트, F=역할(의장단/도어퍼슨)
// 출석기록 탭: A=날짜, B=이름, C=구분, D=대리대상, E=시각, F=지각여부, G=추첨여부
// 당첨기록 탭: A=날짜, B=당첨자, C=방식, D=시각 (자동생성)

const RAFFLE_TIME = '06:30';
const LATE_TIME = '07:00';

function doGet(e) {
  const action = e.parameter.action;
  const callback = e.parameter.callback;
  let result;

  try {
    if (action === 'checkin') {
      result = handleCheckin(e.parameter.pin);
    } else if (action === 'visitorCheckin') {
      result = handleVisitorCheckin(e.parameter.pin, e.parameter.name, e.parameter.host);
    } else if (action === 'hostList') {
      result = getHostCandidates();
    } else if (action === 'today') {
      result = getTodayRecords();
    } else if (action === 'logWinner') {
      result = logWinner(e.parameter.name, e.parameter.mode);
    } else {
      result = { status: 'error', message: 'Unknown action' };
    }
  } catch(err) {
    result = { status: 'error', message: err.toString() };
  }

  const json = JSON.stringify(result);

  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// 날짜를 yyyy-MM-dd 문자열로 변환
function toDateStr(val) {
  try {
    const d = new Date(val);
    return Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd');
  } catch(e) {
    return String(val).slice(0, 10);
  }
}

// 시각을 HH:mm 문자열로 변환
function toTimeStr(val) {
  if (!val) return '';
  try {
    const str = String(val);
    if (/^\d{2}:\d{2}$/.test(str)) return str;
    const d = new Date(val);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, 'Asia/Seoul', 'HH:mm');
    }
    return str;
  } catch(e) {
    return String(val);
  }
}

function handleCheckin(pin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName('명단');
  const recordSheet = ss.getSheetByName('출석기록');

  if (!pin || pin.length !== 4) {
    return { status: 'error', message: '핀번호 오류' };
  }

  const normalizedPin = pin.padStart(4, '0');

  const members = memberSheet.getDataRange().getValues();
  let found = null;

  for (let i = 1; i < members.length; i++) {
    const rowPin = String(members[i][1]).trim().padStart(4, '0');
    if (rowPin === normalizedPin) {
      found = {
        name: members[i][0],
        type: members[i][2] || '멤버',
        subFor: members[i][3] || '',
        host: members[i][4] || ''
      };
      break;
    }
  }

  if (!found) {
    return { status: 'notfound', message: '등록되지 않은 번호' };
  }

  const today = getTodayStr();
  const records = recordSheet.getDataRange().getValues();

  for (let i = 1; i < records.length; i++) {
    const recDate = toDateStr(records[i][0]);
    const recName = records[i][1];
    if (recDate === today && recName === found.name) {
      return { status: 'already', message: '이미 체크인됨', name: found.name };
    }
  }

  return doCheckinWrite(recordSheet, today, found.name, found.type, found.subFor, found.host);
}

// 호스트 후보 목록 (의장단/도어퍼슨/이미 배정된 호스트 제외)
function getHostCandidates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName('명단');
  const members = memberSheet.getDataRange().getValues();

  // 이미 호스트로 배정된 이름 수집 (E열)
  const assignedHosts = new Set();
  for (let i = 1; i < members.length; i++) {
    const host = String(members[i][4] || '').trim();
    if (host) assignedHosts.add(host);
  }

  const candidates = [];
  for (let i = 1; i < members.length; i++) {
    const name = String(members[i][0] || '').trim();
    const type = String(members[i][2] || '').trim();
    const role = String(members[i][5] || '').trim();
    if (!name) continue;
    if (type !== '멤버') continue;                    // 멤버만
    if (role === '의장단' || role === '도어퍼슨') continue;  // 역할자 제외
    if (assignedHosts.has(name)) continue;            // 이미 호스팅 배정된 멤버 제외
    candidates.push(name);
  }

  return { status: 'ok', hosts: candidates };
}

// 비지터 현장 등록 + 체크인 (호스트 포함)
function handleVisitorCheckin(pin, name, host) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName('명단');
  const recordSheet = ss.getSheetByName('출석기록');

  if (!pin || pin.length !== 4) {
    return { status: 'error', message: '핀번호 오류' };
  }
  if (!name || name.trim() === '') {
    return { status: 'error', message: '이름을 입력해주세요' };
  }

  const trimmedName = name.trim();
  const trimmedHost = (host || '').trim();
  const normalizedPin = pin.padStart(4, '0');

  const today = getTodayStr();
  const records = recordSheet.getDataRange().getValues();
  for (let i = 1; i < records.length; i++) {
    const recDate = toDateStr(records[i][0]);
    if (recDate === today && records[i][1] === trimmedName) {
      return { status: 'already', message: '이미 체크인됨', name: trimmedName };
    }
  }

  // 명단 탭에 비지터로 추가 (호스트 포함)
  memberSheet.appendRow([trimmedName, normalizedPin, '비지터', '', trimmedHost, '']);

  return doCheckinWrite(recordSheet, today, trimmedName, '비지터', '', trimmedHost);
}

// 공통 체크인 기록 저장 로직
function doCheckinWrite(recordSheet, today, name, type, subFor, host) {
  const now = new Date();
  const timeStr = Utilities.formatDate(now, 'Asia/Seoul', 'HH:mm');
  const isLate = timeToMin(timeStr) >= timeToMin(LATE_TIME);
  const isRaffle = timeToMin(timeStr) < timeToMin(RAFFLE_TIME) &&
                   (type === '멤버' || type === '대리');

  recordSheet.appendRow([
    today, name, type, subFor, timeStr,
    isLate ? 'Y' : 'N',
    isRaffle ? 'Y' : 'N'
  ]);

  return {
    status: 'ok',
    name: name,
    type: type,
    subFor: subFor,
    host: host || '',
    time: timeStr,
    isLate: isLate,
    isRaffle: isRaffle
  };
}

function getTodayRecords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recordSheet = ss.getSheetByName('출석기록');
  const memberSheet = ss.getSheetByName('명단');

  const today = getTodayStr();
  const rows = recordSheet.getDataRange().getValues();

  // 명단에서 이름→호스트 매핑
  const members = memberSheet.getDataRange().getValues();
  const hostMap = {};
  for (let i = 1; i < members.length; i++) {
    const nm = String(members[i][0] || '').trim();
    const host = String(members[i][4] || '').trim();
    if (nm && host) hostMap[nm] = host;
  }

  const records = [];
  for (let i = 1; i < rows.length; i++) {
    const recDate = toDateStr(rows[i][0]);
    if (recDate === today) {
      const nm = rows[i][1];
      records.push({
        name: nm,
        type: rows[i][2],
        subFor: rows[i][3],
        host: hostMap[nm] || '',
        time: toTimeStr(rows[i][4]),
        isLate: rows[i][5] === 'Y',
        isRaffle: rows[i][6] === 'Y'
      });
    }
  }

  const memberCount = members.slice(1).filter(r => r[0] !== '').length;

  return {
    status: 'ok',
    records: records,
    memberCount: memberCount
  };
}

// 추첨 당첨자 시트 기록
function logWinner(name, mode) {
  if (!name) return { status: 'error', message: '이름 없음' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('당첨기록');
  if (!sheet) {
    sheet = ss.insertSheet('당첨기록');
    sheet.appendRow(['날짜', '당첨자', '방식', '시각']);
  }

  const today = getTodayStr();
  const timeStr = Utilities.formatDate(new Date(), 'Asia/Seoul', 'HH:mm');
  const modeNames = { spin: '슬롯머신', roulette: '룰렛', card: '카드뒤집기', name: '이름추첨' };

  sheet.appendRow([today, name, modeNames[mode] || mode, timeStr]);

  return { status: 'ok', name: name };
}

function getTodayStr() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
}

function timeToMin(timeStr) {
  const [h, m] = timeStr.split(':').map(Number);
  return h * 60 + m;
}

function testToday() {
  Logger.log(getTodayStr());
}
