/**
 * ==================== 주간 출석 집계 시스템 ====================
 * 방안 4: 월요일 기준 주간 집계
 * - 주 = 월~일 (7일)
 * - 각 달은 "월요일이 그 달에 속한 주"만 포함
 * - 주 4회 인증 필요 (장기오프 일수만큼 차감)
 * - OFF.md 오프제도 폐지
 */

/**
 * 특정 달의 모든 주 목록 가져오기 (월요일 기준)
 * @param {number} year - 년도
 * @param {number} month - 월 (0-based: 0=1월)
 * @returns {Array} 주 목록 [{시작: Date, 끝: Date}, ...]
 */
function 월별주목록가져오기(year, month) {
  const 주목록 = [];

  // 이 달의 첫날과 마지막날
  const 첫날 = new Date(year, month, 1);
  const 마지막날 = new Date(year, month + 1, 0);

  // 이 달의 첫 월요일 찾기
  let current = new Date(첫날);

  // 첫날이 월요일이 아니면, 첫 월요일로 이동
  while (current.getDay() !== 1) {
    current.setDate(current.getDate() + 1);
    if (current > 마지막날) {
      // 이 달에 월요일이 없음 (거의 불가능하지만)
      return [];
    }
  }

  // 첫 월요일부터 시작해서 주 단위로 반복
  while (current <= 마지막날) {
    const 주시작 = new Date(current);
    const 주끝 = new Date(current);
    주끝.setDate(주끝.getDate() + 6); // 일요일

    주목록.push({
      시작: 주시작,
      끝: 주끝
    });

    // 다음 주 월요일
    current.setDate(current.getDate() + 7);
  }

  return 주목록;
}

/**
 * 특정 주의 인증 횟수 계산
 * @param {string} memberName - 조원 이름
 * @param {Date} 주시작 - 주 시작일 (월요일)
 * @param {Date} 주끝 - 주 종료일 (일요일)
 * @param {boolean} 완료된주 - 주가 완료되었는지 여부 (기본값: true)
 * @returns {Object} {인증횟수, 장기오프일수, 필요횟수, 결석, 주완료}
 */
function 주간인증계산(memberName, 주시작, 주끝, 완료된주 = true) {
  let 인증횟수 = 0;
  let 장기오프일수 = 0;

  // 주의 각 날짜 체크
  for (let d = new Date(주시작); d <= 주끝; d.setDate(d.getDate() + 1)) {
    const dateStr = Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd');

    // 장기오프 확인 (최우선)
    const longOffInfo = 장기오프확인(memberName, dateStr);
    if (longOffInfo.isLongOff) {
      장기오프일수++;
      continue;
    }

    // 출석 확인
    const 출석여부 = 출석확인(memberName, dateStr);
    if (출석여부) {
      인증횟수++;
    }
  }

  // 필요 인증 횟수 = 4 - 장기오프일수
  const 필요횟수 = Math.max(0, 4 - 장기오프일수);

  // 전체 주가 장기오프면 (7일 모두)
  if (장기오프일수 === 7) {
    return {
      인증횟수,
      장기오프일수,
      필요횟수: 0,
      결석: 0,
      전체장기오프: true,
      주완료: 완료된주
    };
  }

  // 결석 계산 - 주가 완료된 경우만
  let 결석 = 0;
  if (완료된주 && 인증횟수 < 필요횟수) {
    const 부족 = 필요횟수 - 인증횟수;

    if (부족 === 1) 결석 = 1;
    else if (부족 === 2) 결석 = 2;
    else if (부족 === 3) 결석 = 3;
    else if (부족 >= 4) 결석 = 4;
  }

  return {
    인증횟수,
    장기오프일수,
    필요횟수,
    결석,
    전체장기오프: false,
    주완료: 완료된주
  };
}

/**
 * 출석 확인 (제출기록 시트 또는 Drive 폴더)
 * @param {string} memberName - 조원 이름
 * @param {string} dateStr - 날짜 (yyyy-MM-dd)
 * @returns {boolean} 출석 여부
 */
function 출석확인(memberName, dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) return false;

  const data = sheet.getDataRange().getValues();

  // 제출기록 시트에서 확인
  for (let i = 1; i < data.length; i++) {
    const [timestamp, name, recordDate, fileCount, links, folderLink, status, weekNum, reason] = data[i];

    const recordDateStr = typeof recordDate === 'string'
      ? recordDate
      : Utilities.formatDate(new Date(recordDate), 'Asia/Seoul', 'yyyy-MM-dd');

    if (name === memberName && recordDateStr === dateStr) {
      // O (출석)만 인정, OFF는 폐지됨
      return status === 'O';
    }
  }

  return false;
}

/**
 * 월별 주간 집계 실행
 * @param {number} year - 년도
 * @param {number} month - 월 (0-based)
 * @returns {Object} {조원명: {주차별결석: [], 총결석: N}}
 */
function 월별주간집계(year, month) {
  Logger.log(`=== ${year}년 ${month + 1}월 주간 집계 시작 ===`);

  const 주목록 = 월별주목록가져오기(year, month);
  const 조원결석 = {};

  Logger.log(`총 ${주목록.length}개 주 발견`);

  // 현재 날짜 (한국 시간)
  const 오늘 = new Date();
  const 오늘자정 = new Date(오늘.getFullYear(), 오늘.getMonth(), 오늘.getDate());

  // 각 주별 집계
  for (let weekIdx = 0; weekIdx < 주목록.length; weekIdx++) {
    const 주 = 주목록[weekIdx];
    const 주차 = weekIdx + 1;

    // 주가 완료되었는지 확인 (일요일이 지났으면 완료)
    const 완료된주 = 주.끝 < 오늘자정;
    const 주상태 = 완료된주 ? '완료' : '진행중';

    Logger.log(`\n--- ${주차}주차: ${Utilities.formatDate(주.시작, 'Asia/Seoul', 'MM/dd')} ~ ${Utilities.formatDate(주.끝, 'Asia/Seoul', 'MM/dd')} (${주상태}) ---`);

    // 각 조원별 계산
    for (const memberName of Object.keys(CONFIG.MEMBERS)) {
      const 결과 = 주간인증계산(memberName, 주.시작, 주.끝, 완료된주);

      if (!조원결석[memberName]) {
        조원결석[memberName] = {
          주차별결석: [],
          주차별상세: [],
          총결석: 0
        };
      }

      조원결석[memberName].주차별결석.push(결과.결석);
      조원결석[memberName].주차별상세.push({
        주차,
        인증횟수: 결과.인증횟수,
        장기오프일수: 결과.장기오프일수,
        필요횟수: 결과.필요횟수,
        결석: 결과.결석,
        전체장기오프: 결과.전체장기오프,
        주완료: 결과.주완료
      });

      if (!결과.전체장기오프) {
        조원결석[memberName].총결석 += 결과.결석;
      }

      const 상태표시 = 결과.주완료 ? `→ 결석 ${결과.결석}회` : '(진행중)';
      Logger.log(`  ${memberName}: 인증 ${결과.인증횟수}/${결과.필요횟수}회 (장기오프 ${결과.장기오프일수}일) ${상태표시}`);
    }
  }

  // 최종 결과 출력
  Logger.log('\n=== 월별 결석 요약 ===');
  for (const [memberName, data] of Object.entries(조원결석)) {
    const 벌칙여부 = data.총결석 >= 4 ? ' 🚨 벌칙대상' : data.총결석 === 3 ? ' ⚠️ 경고' : '';
    Logger.log(`${memberName}: 총 ${data.총결석}회 결석${벌칙여부}`);
  }

  Logger.log('\n=== 주간 집계 완료 ===');

  return 조원결석;
}

/**
 * 주간 집계 결과를 시트에 저장
 * @param {number} year - 년도
 * @param {number} month - 월 (0-based)
 * @param {Object} 집계결과 - 월별주간집계() 결과
 */
function 주간집계저장(year, month, 집계결과) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('주간집계');

  // 시트가 없으면 생성
  if (!sheet) {
    sheet = ss.insertSheet('주간집계');

    // 헤더 작성
    sheet.getRange('A1:I1').setValues([[
      '년월', '조원명', '주차', '인증', '필요', '장기오프일', '결석', '상태', '비고'
    ]]);
    sheet.getRange('A1:I1').setFontWeight('bold');
    sheet.getRange('A1:I1').setBackground('#4CAF50');
    sheet.getRange('A1:I1').setFontColor('white');
  }

  const 년월 = `${year}-${String(month + 1).padStart(2, '0')}`;

  // 기존 데이터 삭제 (해당 년월)
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === 년월) {
      sheet.deleteRow(i + 1);
    }
  }

  // 새 데이터 추가
  const rows = [];

  for (const [memberName, data] of Object.entries(집계결과)) {
    for (const 주상세 of data.주차별상세) {
      const 상태 = 주상세.주완료 ? '완료' : '진행중';
      const 비고 = 주상세.전체장기오프 ? '전체장기오프' : '';

      rows.push([
        년월,
        memberName,
        주상세.주차,
        주상세.인증횟수,
        주상세.필요횟수,
        주상세.장기오프일수,
        주상세.결석,
        상태,
        비고
      ]);
    }
  }

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 9).setValues(rows);
  }

  Logger.log(`주간집계 시트에 ${rows.length}개 행 저장`);
}

/**
 * 주간 집계 결과를 JSON 파일로 저장
 * @param {number} year - 년도
 * @param {number} month - 월 (0-based)
 * @param {Object} 집계결과 - 월별주간집계() 결과
 */
function 주간집계JSON저장(year, month, 집계결과) {
  const 년월 = `${year}-${String(month + 1).padStart(2, '0')}`;

  // JSON 데이터 생성
  const jsonData = {
    년월,
    생성일시: new Date().toISOString(),
    안내: {
      주기준: '월요일 시작',
      설명: '각 주는 월요일부터 일요일까지입니다. 월요일이 속한 달의 주로 계산됩니다.',
      예시: '11월 25일(월)~12월 1일(일) → 11월 4주차'
    },
    조원별집계: {}
  };

  // 조원별 데이터 추가
  for (const [memberName, data] of Object.entries(집계결과)) {
    jsonData.조원별집계[memberName] = {
      총결석: data.총결석,
      주차별: data.주차별상세.map(주 => ({
        주차: 주.주차,
        인증: 주.인증횟수,
        필요: 주.필요횟수,
        장기오프: 주.장기오프일수,
        결석: 주.결석,
        상태: 주.주완료 ? '완료' : '진행중',
        전체장기오프: 주.전체장기오프
      }))
    };
  }

  // JSON 파일로 저장
  const jsonString = JSON.stringify(jsonData, null, 2);
  const fileName = `weekly_summary_${년월}.json`;

  try {
    const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);

    // 기존 파일 삭제
    const existingFiles = folder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }

    // 새 파일 생성
    const file = folder.createFile(fileName, jsonString, MimeType.PLAIN_TEXT);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileUrl = `https://drive.google.com/uc?export=download&id=${file.getId()}`;
    Logger.log(`주간 집계 JSON 저장 완료: ${fileName}`);
    Logger.log(`URL: ${fileUrl}`);

  } catch (e) {
    Logger.log(`주간 집계 JSON 저장 실패: ${e.message}`);
  }
}

/**
 * 매일 자동 실행 - 주간 집계 (일요일 밤에만)
 */
function 주간집계자동실행() {
  const now = new Date();
  const dayOfWeek = now.getDay(); // 0=일요일

  // 일요일 밤에만 실행
  if (dayOfWeek !== 0) {
    Logger.log('오늘은 일요일이 아니므로 주간집계를 건너뜁니다.');
    return;
  }

  const year = now.getFullYear();
  const month = now.getMonth();

  const 집계결과 = 월별주간집계(year, month);
  주간집계저장(year, month, 집계결과);
  주간집계JSON저장(year, month, 집계결과);
}

/**
 * 수동 실행용: 이번 달 주간 집계
 */
function 이번달주간집계() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth();

  const 집계결과 = 월별주간집계(year, month);
  주간집계저장(year, month, 집계결과);
  주간집계JSON저장(year, month, 집계결과);
}

/**
 * JSON 파일 ID 확인 (HTML 설정용)
 * HTML의 JSON_FILE_IDS에 입력할 파일 ID를 출력합니다.
 */
function JSON파일ID확인() {
  const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);

  // 현재 연월 계산
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  const yearMonth = year + '-' + String(month).padStart(2, '0');

  Logger.log('');
  Logger.log('='.repeat(60));
  Logger.log('📁 JSON 파일 ID 목록 (HTML 설정용)');
  Logger.log('='.repeat(60));
  Logger.log('');

  // 1. 일간 출석 파일
  const attendanceFileName = `attendance_summary_${yearMonth}.json`;
  const attendanceFiles = folder.getFilesByName(attendanceFileName);

  if (attendanceFiles.hasNext()) {
    const file = attendanceFiles.next();
    const fileId = file.getId();
    const url = `https://drive.google.com/uc?export=download&id=${fileId}`;

    Logger.log('📄 일간 출석 파일:');
    Logger.log('   파일명: ' + attendanceFileName);
    Logger.log('   파일 ID: ' + fileId);
    Logger.log('   전체 URL: ' + url);
    Logger.log('');
  } else {
    Logger.log('❌ 일간 출석 파일을 찾을 수 없습니다: ' + attendanceFileName);
    Logger.log('   → 먼저 월말집계() 함수를 실행해주세요!');
    Logger.log('');
  }

  // 2. 주간 집계 파일
  const weeklyFileName = `weekly_summary_${yearMonth}.json`;
  const weeklyFiles = folder.getFilesByName(weeklyFileName);

  if (weeklyFiles.hasNext()) {
    const file = weeklyFiles.next();
    const fileId = file.getId();
    const url = `https://drive.google.com/uc?export=download&id=${fileId}`;

    Logger.log('📊 주간 집계 파일:');
    Logger.log('   파일명: ' + weeklyFileName);
    Logger.log('   파일 ID: ' + fileId);
    Logger.log('   전체 URL: ' + url);
    Logger.log('');
  } else {
    Logger.log('❌ 주간 집계 파일을 찾을 수 없습니다: ' + weeklyFileName);
    Logger.log('   → 먼저 이번달주간집계() 함수를 실행해주세요!');
    Logger.log('');
  }

  Logger.log('-'.repeat(60));
  Logger.log('📋 HTML 설정 방법:');
  Logger.log('-'.repeat(60));
  Logger.log('');
  Logger.log('1. GitHub에서 index.html 파일 열기');
  Logger.log('2. Ctrl+F로 "JSON_FILE_IDS" 검색');
  Logger.log('3. 위의 파일 ID들을 다음과 같이 입력:');
  Logger.log('');
  Logger.log('   const JSON_FILE_IDS = {');
  Logger.log('       attendance: \'위의_일간_출석_파일_ID\',');
  Logger.log('       weekly: \'위의_주간_집계_파일_ID\'');
  Logger.log('   };');
  Logger.log('');
  Logger.log('4. 커밋 후 GitHub Pages에서 확인');
  Logger.log('');
  Logger.log('='.repeat(60));
  Logger.log('');
}

/**
 * JSON 폴더 URL 확인
 * 전체 폴더의 파일 목록을 확인할 수 있습니다.
 */
function JSON폴더URL확인() {
  const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);

  Logger.log('');
  Logger.log('📁 JSON 폴더 정보:');
  Logger.log('   폴더 ID: ' + CONFIG.JSON_FOLDER_ID);
  Logger.log('   폴더명: ' + folder.getName());
  Logger.log('   폴더 URL: ' + folder.getUrl());
  Logger.log('');
  Logger.log('📄 폴더 내 JSON 파일 목록:');
  Logger.log('');

  const files = folder.getFiles();
  let count = 0;

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();

    if (fileName.endsWith('.json')) {
      count++;
      Logger.log(`   ${count}. ${fileName}`);
      Logger.log('      파일 ID: ' + file.getId());
      Logger.log('      URL: https://drive.google.com/uc?export=download&id=' + file.getId());
      Logger.log('');
    }
  }

  if (count === 0) {
    Logger.log('   (JSON 파일이 없습니다)');
    Logger.log('');
  }

  Logger.log('='.repeat(60));
}
