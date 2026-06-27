// BNI Pioneer 출석체크 - Google Apps Script
// 시트 구조:
//   탭1 "명단": A=이름, B=전화번호뒷4자리, C=구분(멤버/비지터/대리), D=대리대상
//   탭2 "출석기록": A=날짜, B=이름, C=구분, D=대리대상, E=시각, F=지각여부, G=추첨여부

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const RAFFLE_TIME = '06:30'; // 추첨 기준 (설정 가능)
const LATE_TIME = '07:00';   // 지각 기준 (설정 가능)

function doGet(e) {
  const action = e.parameter.action;
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

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleCheckin(pin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName('명단');
  const recordSheet = ss.getSheetByName('출석기록');

  if (!pin || pin.length !== 4) {
    return { status: 'error', message: '핀번호 오류' };
  }

  // 명단에서 검색
  const members = memberSheet.getDataRange().getValues();
  let found = null;

  for (let i = 1; i < members.length; i++) {
    const rowPin = String(members[i][1]).trim();
    if (rowPin === pin) {
      found = {
        name: members[i][0],
        pin: members[i][1],
        type: members[i][2] || '멤버',
        subFor: members[i][3] || ''
      };
      break;
    }
  }

  if (!found) {
    return { status: 'notfound', message: '등록되지 않은 번호' };
  }

  // 오늘 이미 체크인 했는지 확인
  const today = getTodayStr();
  const records = recordSheet.getDataRange().getValues();

  for (let i = 1; i < records.length; i++) {
    const recDate = String(records[i][0]).slice(0, 10);
    const recName = records[i][1];
    if (recDate === today && recName === found.name) {
      return { status: 'already', message: '이미 체크인됨' };
    }
  }

  // 시간 계산
  const now = new Date();
  const timeStr = Utilities.formatDate(now, 'Asia/Seoul', 'HH:mm');
  const isLate = timeToMin(timeStr) >= timeToMin(LATE_TIME);
  const isRaffle = timeToMin(timeStr) < timeToMin(RAFFLE_TIME) &&
                   (found.type === '멤버' || found.type === '대리');

  // 기록 저장
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

  // 전체 멤버 수
  const allMembers = memberSheet.getDataRange().getValues();
  const memberCount = allMembers.slice(1).filter(r => r[2] === '멤버').length;

  return {
    status: 'ok',
    records: records,
    memberCount: memberCount
  };
}

// ===== 유틸 =====
function getTodayStr() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
}

function timeToMin(timeStr) {
  const [h, m] = timeStr.split(':').map(Number);
  return h * 60 + m;
}
