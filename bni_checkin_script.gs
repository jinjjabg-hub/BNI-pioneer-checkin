// BNI Pioneer 출석체크 - Google Apps Script (JSONP 지원)

const RAFFLE_TIME = '06:30';
const LATE_TIME = '07:00';

function doGet(e) {
  const action = e.parameter.action;
  const callback = e.parameter.callback;
  let result;

  try {
    if (action === 'checkin') {
      result = handleCheckin(e.parameter.pin);
    } else if (action === 'today') {
      result = getTodayRecords();
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

function handleCheckin(pin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName('명단');
  const recordSheet = ss.getSheetByName('출석기록');

  if (!pin || pin.length !== 4) {
    return { status: 'error', message: '핀번호 오류' };
  }

  const members = memberSheet.getDataRange().getValues();
  let found = null;

  for (let i = 1; i < members.length; i++) {
    const rowPin = String(members[i][1]).trim();
    if (rowPin.padStart(4,'0') === pin.padStart(4,'0')) {
      found = {
        name: members[i][0],
        type: members[i][2] || '멤버',
        subFor: members[i][3] || ''
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
    const recDate = String(records[i][0]).slice(0, 10);
    const recName = records[i][1];
    if (recDate === today && recName === found.name) {
      return { status: 'already', message: '이미 체크인됨', name: found.name };
    }
  }

  const now = new Date();
  const timeStr = Utilities.formatDate(now, 'Asia/Seoul', 'HH:mm');
  const isLate = timeToMin(timeStr) >= timeToMin(LATE_TIME);
  const isRaffle = timeToMin(timeStr) < timeToMin(RAFFLE_TIME) &&
                   (found.type === '멤버' || found.type === '대리');

  recordSheet.appendRow([
    today,
    found.name,
    found.type,
    found.subFor,
    timeStr,
    isLate ? 'Y' : 'N',
    isRaffle ? 'Y' : 'N'
  ]);

  return {
    status: 'ok',
    name: found.name,
    type: found.type,
    subFor: found.subFor,
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
  const records = [];

  for (let i = 1; i < rows.length; i++) {
    const recDate = String(rows[i][0]).slice(0, 10);
    if (recDate === today) {
      records.push({
        name: rows[i][1],
        type: rows[i][2],
        subFor: rows[i][3],
        time: rows[i][4],
        isLate: rows[i][5] === 'Y',
        isRaffle: rows[i][6] === 'Y'
      });
    }
  }

  const allMembers = memberSheet.getDataRange().getValues();
  const memberCount = allMembers.slice(1).filter(r => r[0] !== '').length;

  return {
    status: 'ok',
    records: records,
    memberCount: memberCount
  };
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

function testRecords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('출석기록');
  const rows = sheet.getDataRange().getValues();
  const today = getTodayStr();
  Logger.log('오늘: ' + today);
  Logger.log('2행 날짜: ' + String(rows[1][0]));
  Logger.log('2행 날짜 slice: ' + String(rows[1][0]).slice(0,10));
  Logger.log('일치여부: ' + (String(rows[1][0]).slice(0,10) === today));
}
