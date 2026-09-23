// BNI Pioneer 출석체크 - Google Apps Script v2 (고속화)
// 명단 탭: A=이름, B=전화번호뒷4자리, C=구분, D=대리대상, E=호스트, F=역할
// 출석기록 탭: A=날짜, B=이름, C=구분, D=대리대상, E=시각, F=지각여부, G=추첨여부
// 당첨기록 탭: A=날짜, B=당첨자, C=방식, D=시각 (자동생성)

const RAFFLE_TIME = '06:20';
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
    } else if (action === 'memberList') {
      result = getMemberList();
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

function toDateStr(val) {
  try {
    const d = new Date(val);
    return Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd');
  } catch(e) {
    return String(val).slice(0, 10);
  }
}

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

// 전체 명단 내려주기 (앱이 시작할 때 캐싱용)
function getMemberList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName('명단');
  const members = memberSheet.getDataRange().getValues();

  const list = [];
  for (let i = 1; i < members.length; i++) {
    const name = String(members[i][0] || '').trim();
    if (!name) continue;
    list.push({
      name: name,
      pin: String(members[i][1] || '').trim().padStart(4, '0'),
      type: String(members[i][2] || '멤버').trim(),
      subFor: String(members[i][3] || '').trim(),
      host: String(members[i][4] || '').trim(),
      role: String(members[i][5] || '').trim()
    });
  }

  return { status: 'ok', members: list };
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

  if (isAlreadyCheckedIn(recordSheet, found.name)) {
    return { status: 'already', message: '이미 체크인됨', name: found.name };
  }

  return doCheckinWrite(recordSheet, found.name, found.type, found.subFor, found.host);
}

// 오늘 중복 체크 (뒤에서부터 검사 - 최신 데이터가 아래에 있으므로 빠름)
function isAlreadyCheckedIn(recordSheet, name) {
  const today = getTodayStr();
  const lastRow = recordSheet.getLastRow();
  if (lastRow < 2) return false;

  // 최근 200행만 검사 (충분함)
  const startRow = Math.max(2, lastRow - 199);
  const numRows = lastRow - startRow + 1;
  const data = recordSheet.getRange(startRow, 1, numRows, 2).getValues();

  for (let i = data.length - 1; i >= 0; i--) {
    const recDate = toDateStr(data[i][0]);
    if (recDate < today) break; // 오늘보다 이전 날짜 나오면 중단
    if (recDate === today && data[i][1] === name) return true;
  }
  return false;
}

function getHostCandidates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName('명단');
  const members = memberSheet.getDataRange().getValues();

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
    if (type !== '멤버') continue;
    if (role === '의장단' || role === '도어퍼슨') continue;
    if (assignedHosts.has(name)) continue;
    candidates.push(name);
  }

  return { status: 'ok', hosts: candidates };
}

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

  if (isAlreadyCheckedIn(recordSheet, trimmedName)) {
    return { status: 'already', message: '이미 체크인됨', name: trimmedName };
  }

  memberSheet.appendRow([trimmedName, normalizedPin, '비지터', '', trimmedHost, '']);

  return doCheckinWrite(recordSheet, trimmedName, '비지터', '', trimmedHost);
}

function doCheckinWrite(recordSheet, name, type, subFor, host) {
  const now = new Date();
  const today = getTodayStr();
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
  const lastRow = recordSheet.getLastRow();
  const records = [];

  if (lastRow >= 2) {
    // 최근 300행만 읽기 (하루 출석은 그 안에 다 들어옴)
    const startRow = Math.max(2, lastRow - 299);
    const numRows = lastRow - startRow + 1;
    const rows = recordSheet.getRange(startRow, 1, numRows, 7).getValues();

    // 명단에서 이름→호스트 매핑
    const members = memberSheet.getDataRange().getValues();
    const hostMap = {};
    for (let i = 1; i < members.length; i++) {
      const nm = String(members[i][0] || '').trim();
      const host = String(members[i][4] || '').trim();
      if (nm && host) hostMap[nm] = host;
    }

    for (let i = 0; i < rows.length; i++) {
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
  }

  const memberSheet2 = ss.getSheetByName('명단');
  const allRows = memberSheet2.getDataRange().getValues();
  const memberCount = allRows.slice(1).filter(r => String(r[0]).trim() !== '' && String(r[2]).trim() === '멤버').length;

  return {
    status: 'ok',
    records: records,
    memberCount: memberCount
  };
}

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
