/**
 * 스터디 출석 자동화 시스템 v3.0 (주간 집계 + AI 통합 버전)
 * 마감시간 제도 + 장기오프 제도 (구글 폼) 통합 버전
 * * 주요 기능:
 * - 새벽 3시 마감 제도 (미제출 시 자동 결석)
 * - 장기오프 제도 (구글 폼으로 사전 신청)
 * - 구글 폼 응답 자동 처리
 * - 주간 집계 시스템 (월요일 기준, 주 4회 인증 필요)
 * - Gemini AI 자동 요약 및 질 평가
 * * 설치 방법:
 * 1. Google Forms에서 "장기오프 신청" 폼 생성
 * 2. 폼을 기존 스프레드시트에 연결
 * 3. Google Sheets "출석표" 열기
 * 4. 확장 프로그램 > Apps Script
 * 5. 이 코드 복사-붙여넣기
 * 6. 저장 후 "초기설정" 함수 실행
 * 7. 권한 승인
 * 8. 트리거 자동 설정됨
 */

// ==================== 설정 ====================
const CONFIG = {
  // 조원 정보
  MEMBERS: {
    '센트룸': '1Wm2l0gzgo2w6EuT3VD7ToFXQksYEC-nc',
    '길': '1mdq7dI-nE5mY0wo6iYK2otSHmC8jnToz',
    'what': '1UtswVFSZtLlbQUZx35mBc6QD6q9zJPkg',
    '머리 빗는 네오': '1XQIgvcZ4uD__JxddsxKTbk0PzNB6Js2k',
    '녹동': '1-aEr_ER-o8SxcQzLCMeAqEy-cghtUl-R',
    '오늘의너굴이': '1572mLeNrDLWLnXRronM-cfNpnUt-wBAM',
    'Dann': '1mMoVApl7GN3EUYi9oPi7Nfo_2hYDb9Dw',
    '보노보노': '1_Mqn79Y1Qp79DWBxcbP-SGVUGjJA3PGw',
    'Magnus': ['1eHjsJ8bnWcK__8EXvukqixzh4wb8CncR', '1e8HUMzD0zW0BG2rkuB3kXoGtK2fw2fhG']
  },
  
  // 시트 이름
  SHEET_NAME: '제출기록',
  ATTENDANCE_SHEET: '출석표',
  LONG_OFF_SHEET: '장기오프신청',
  ADMIN_SHEET: '관리자수정',  // 🆕 추가
  MONTHLY_SUMMARY_SHEET: '월별결산',  // 🆕 월별결산 시트
  
  // JSON 파일 출력 폴더 ID
  JSON_FOLDER_ID: '1el9NDYDGfWlUEkBzI1GT_1TULLoBnSsQ',

  // 마감시간 설정
  DEADLINE_HOUR: 3,
  
  // 장기오프 설정
  LONG_OFF_STATUS: 'LONG_OFF',
  LONG_OFF_AUTO_APPROVE: true,
  
  // 구글 폼 응답 시트 열 구조
  FORM_COLUMNS: {
    TIMESTAMP: 0,
    NAME: 1,
    START_DATE: 2,
    END_DATE: 3,
    REASON: 4,
    APPROVED: 5
  },
  
  // 🆕 관리자수정 시트 열 구조
  ADMIN_COLUMNS: {
    NAME: 0,
    DATE: 1,
    STATUS: 2,
    REASON: 3,
    PROCESSED: 4,
    PROCESSED_TIME: 5
  },
  
  // 스캔 설정
  SCAN_ALL_MONTHS: false,
  MAX_FOLDERS_TO_SCAN: 100  // 마지막 항목은 콤마 없어도 OK
};

// ==================== 메인 함수 ====================

/**
 * 매 시간 실행되는 메인 함수
 * 트리거로 설정해야 함
 * * 각 조원 폴더 안의 모든 날짜 폴더를 스캔합니다
 */
function 출석체크_메인() {
  Logger.log('=== 출석 체크 시작 (전체 폴더 스캔) ===');

  관리자수정처리(); 
  
  const results = [];
  const currentMonth = new Date().getMonth(); // 0-based (0=1월, 9=10월)
  const currentYear = new Date().getFullYear();
  
  if (CONFIG.SCAN_ALL_MONTHS) {
    Logger.log(`스캔 모드: 모든 달 체크`);
  } else {
    Logger.log(`스캔 모드: 이번 달만 체크 (${currentYear}년 ${currentMonth + 1}월)`);
  }
  Logger.log('');

  
  
  // 각 조원의 폴더 체크
  for (const [memberName, folderIdOrArray] of Object.entries(CONFIG.MEMBERS)) {
    // 폴더 ID를 배열로 정규화 (단일 문자열이면 배열로 변환)
    const folderIds = Array.isArray(folderIdOrArray) ? folderIdOrArray : [folderIdOrArray];
    
    Logger.log(`📁 ${memberName} - ${folderIds.length}개 폴더 스캔 중...`);
    
    // 중복 날짜 체크를 위한 Set
    const processedDates = new Set();
    
    // 각 폴더 스캔
    for (let folderIndex = 0; folderIndex < folderIds.length; folderIndex++) {
      const folderId = folderIds[folderIndex];
      
      try {
        Logger.log(`  📂 폴더 ${folderIndex + 1}/${folderIds.length} (ID: ${folderId.substring(0, 10)}...)`);
        
        const mainFolder = DriveApp.getFolderById(folderId);
        const subfolders = mainFolder.getFolders();
        
        let folderCount = 0;
        let processedCount = 0;
        let skippedCount = 0;
        let duplicateCount = 0;
        
        // 모든 하위 폴더 스캔
        while (subfolders.hasNext() && folderCount < CONFIG.MAX_FOLDERS_TO_SCAN) {
          const folder = subfolders.next();
          const folderName = folder.getName().trim();
          folderCount++;
          
          // 폴더 이름에서 날짜 추출
          const dateInfo = 날짜추출(folderName);
          
          if (dateInfo) {
            const { dateStr, year, month } = dateInfo;
            
            // 이미 처리한 날짜면 건너뛰기 (중복 방지)
            if (processedDates.has(dateStr)) {
              Logger.log(`    ⚠ ${dateStr} - 중복 (이미 다른 폴더에서 처리됨)`);
              duplicateCount++;
              continue;
            }

            // 🆕 마감된 날짜는 스캔하지 않음 (마감 후 제출 방지)
            if (날짜마감확인(dateStr)) {
              Logger.log(`    ⏰ ${dateStr} - 마감됨 (스캔 건너뜀)`);
              skippedCount++;
              continue;
            }
            
            // 이번 달만 체크할지, 모든 달을 체크할지 결정
            const shouldProcess = CONFIG.SCAN_ALL_MONTHS || 
                                  (year === currentYear && month === currentMonth);
            
            if (shouldProcess) {

              // 🆕 이 부분 추가: 관리자수정이 있으면 건너뛰기
              if (관리자수정존재확인(memberName, dateStr)) {
              Logger.log(`    🔧 ${dateStr} - 관리자수정됨 (자동스캔 건너뜀)`);
              processedDates.add(dateStr);
              processedCount++;
              continue;
              }

              // 🆕 장기오프 체크 (최우선)
              const longOffInfo = 장기오프확인(memberName, dateStr);

              if (longOffInfo.isLongOff) {
                Logger.log(`    🏝️ ${dateStr} - 장기오프 (${longOffInfo.reason})`);
                출석기록추가(memberName, dateStr, [], CONFIG.LONG_OFF_STATUS, longOffInfo.reason);
                processedDates.add(dateStr);
                processedCount++;
                continue;
              }

              // 일반 출석 처리
              const files = 파일목록및링크생성(folder);

              if (files.length > 0) {
                Logger.log(`    ✓ ${dateStr} - 출석 (${files.length}개 파일)`);
                출석기록추가(memberName, dateStr, files, 'O');
                processedDates.add(dateStr);
                processedCount++;
              } else {
                Logger.log(`    ⚠ ${dateStr} - 폴더는 있지만 파일 없음`);
                skippedCount++;
              }
            } else {
              skippedCount++;
            }
          }
        }
        
        if (folderCount >= CONFIG.MAX_FOLDERS_TO_SCAN) {
          Logger.log(`    ⚠️ 최대 폴더 수(${CONFIG.MAX_FOLDERS_TO_SCAN})에 도달. 일부만 스캔됨`);
        }
        
        Logger.log(`    → 폴더 ${folderIndex + 1}: ${folderCount}개 검사 / 처리: ${processedCount}개 / 중복: ${duplicateCount}개 / 건너뜀: ${skippedCount}개`);
        
      } catch (error) {
        Logger.log(`    ❌ 폴더 ${folderIndex + 1} 오류: ${error.message}`);
      }
    }
    
    Logger.log(`  ✅ ${memberName} 완료: 총 ${processedDates.size}개 날짜 처리`);
    Logger.log('');
  }

  // JSON 파일 생성
  JSON파일생성();
  
  Logger.log('=== 출석 체크 완료 ===');
}

function 마감시간체크() {
  Logger.log('=== 마감시간 체크 시작 ===');
  
  const now = new Date();
  Logger.log(`현재 시각: ${now}`);
  
  // 🔧 수정: 어제만이 아니라 오늘 이전 최근 7일 체크
  const daysToCheck = 7;
  const targetDates = [];
  
  for (let i = 1; i <= daysToCheck; i++) {
    const checkDate = new Date(now);
    checkDate.setDate(checkDate.getDate() - i);
    
    const dateStr = Utilities.formatDate(checkDate, 'Asia/Seoul', 'yyyy-MM-dd');
    
    // 마감되었는지 확인
    if (날짜마감확인(dateStr)) {
      targetDates.push(dateStr);
    }
  }
  
  Logger.log(`체크 대상 날짜: ${targetDates.join(', ')}`);
  Logger.log('');
  
  // 각 날짜별로 처리
  for (const targetDateStr of targetDates) {
    Logger.log(`📅 ${targetDateStr} 처리 중...`);
    
    // 🔧 마감 전 최종 스캔
    Logger.log('📂 마감 전 최종 스캔 시작...');
    최종스캔_특정날짜(targetDateStr);
    Logger.log('');

    // 방금 추가된 기록이 시트에 완전히 반영되도록 강제 동기화
    SpreadsheetApp.flush();
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      Logger.log('❌ 제출기록 시트가 없습니다.');
      continue;
    }
    
    // 기존 기록 확인
    const data = sheet.getDataRange().getValues();
    const 제출자명단 = new Set();
    
    for (let i = 1; i < data.length; i++) {
      const [timestamp, name, dateStr, fileCount, links, folderLink, status, weekNum, reason] = data[i];
      
      const dateStrFormatted = typeof dateStr === 'string' 
        ? dateStr 
        : Utilities.formatDate(new Date(dateStr), 'Asia/Seoul', 'yyyy-MM-dd');
      
      if (dateStrFormatted === targetDateStr) {
        제출자명단.add(name);
      }
    }
    
    Logger.log(`${targetDateStr} 제출자: ${Array.from(제출자명단).join(', ')}`);
    Logger.log('');
    
    // 미제출자 찾기 및 결석 처리
    let 결석처리수 = 0;
    let 장기오프업데이트수 = 0;
    
    for (const memberName of Object.keys(CONFIG.MEMBERS)) {
      const longOffInfo = 장기오프확인(memberName, targetDateStr);
      
      if (longOffInfo.isLongOff) {
        if (제출자명단.has(memberName)) {
          Logger.log(`🔄 ${memberName} - 기존 기록을 장기오프로 업데이트`);
          장기오프업데이트수++;
        } else {
          Logger.log(`🏝️ ${memberName} - 장기오프 (${longOffInfo.reason})`);
        }
        출석기록추가(memberName, targetDateStr, [], CONFIG.LONG_OFF_STATUS, longOffInfo.reason);
        continue;
      }
      
      if (!제출자명단.has(memberName)) {
        Logger.log(`❌ ${memberName} - 미제출 (결석 처리)`);
        출석기록추가(memberName, targetDateStr, [], 'X');
        결석처리수++;
      } else {
        Logger.log(`✅ ${memberName} - 제출 완료`);
      }
    }
    
    Logger.log('');
    Logger.log(`${targetDateStr}: ${결석처리수}명 결석 처리, ${장기오프업데이트수}명 장기오프 업데이트`);
    Logger.log('');
  }
  
  Logger.log('=== 마감시간 체크 완료 ===');
  
  // JSON 재생성
  Logger.log('JSON 파일 재생성 중...');
  JSON파일생성();
}

/**
 * 🆕 장기오프 확인 함수 (구글 폼 응답 시트 버전)
 * @param {string} memberName - 조원 이름
 * @param {string} dateStr - 확인할 날짜 (yyyy-MM-dd)
 * @returns {Object} {isLongOff: boolean, reason: string}
 */
function 장기오프확인(memberName, dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.LONG_OFF_SHEET);
  
  if (!sheet) {
    return { isLongOff: false, reason: '' };
  }
  
  const data = sheet.getDataRange().getValues();
  const targetDate = new Date(dateStr);
  
  // 첫 행(헤더)은 제외하고 검색
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const timestamp = row[CONFIG.FORM_COLUMNS.TIMESTAMP];
    const name = row[CONFIG.FORM_COLUMNS.NAME];
    const startDateValue = row[CONFIG.FORM_COLUMNS.START_DATE];
    const endDateValue = row[CONFIG.FORM_COLUMNS.END_DATE];
    const reason = row[CONFIG.FORM_COLUMNS.REASON];
    const approved = row[CONFIG.FORM_COLUMNS.APPROVED];
    
    // 🔧 이름 확인 (대소문자 무시)
    if (String(name).trim().toLowerCase() !== memberName.toLowerCase()) continue;
    
    // 승인 여부 확인 (자동 승인 모드가 아닐 때만)
    if (!CONFIG.LONG_OFF_AUTO_APPROVE && approved !== 'O' && approved !== 'o') {
      continue;
    }
    
    // 날짜 범위 확인
    try {
      let startDate, endDate;
      
      if (startDateValue instanceof Date) {
        startDate = startDateValue;
      } else {
        startDate = new Date(startDateValue);
      }
      
      if (endDateValue instanceof Date) {
        endDate = endDateValue;
      } else {
        endDate = new Date(endDateValue);
      }
      
      // 시간 부분 제거 (날짜만 비교)
      startDate.setHours(0, 0, 0, 0);
      endDate.setHours(0, 0, 0, 0);
      const compareDate = new Date(targetDate);
      compareDate.setHours(0, 0, 0, 0);
      
      // targetDate가 startDate와 endDate 사이에 있는지 확인
      if (compareDate >= startDate && compareDate <= endDate) {
        return {
          isLongOff: true,
          reason: reason || '장기오프'
        };
      }
    } catch (e) {
      Logger.log(`장기오프 날짜 파싱 오류: ${name}, ${startDateValue} ~ ${endDateValue}`);
      Logger.log(`오류 상세: ${e.message}`);
    }
  }
  
  return { isLongOff: false, reason: '' };
}

/**
 * 폴더 이름에서 날짜 정보 추출
 * @param {string} folderName - 폴더 이름
 * @returns {Object|null} {dateStr, year, month, day} 또는 null
 */
function 날짜추출(folderName) {
  // 다양한 날짜 형식 패턴
  const patterns = [
    // 2025-10-15 형식
    /^(\d{4})-(\d{1,2})-(\d{1,2})$/,
    // 20251015 형식
    /^(\d{4})(\d{2})(\d{2})$/,
    // 2025.10.15 형식
    /^(\d{4})\.(\d{1,2})\.(\d{1,2})$/,
    // 2025/10/15 형식
    /^(\d{4})\/(\d{1,2})\/(\d{1,2})$/,
  ];
  
  // 년도 포함 형식 체크
  for (const pattern of patterns) {
    const match = folderName.match(pattern);
    if (match) {
      const year = parseInt(match[1]);
      const month = parseInt(match[2]) - 1; // 0-based
      const day = parseInt(match[3]);
      
      // 유효한 날짜인지 검증
      if (month >= 0 && month <= 11 && day >= 1 && day <= 31) {
        const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        return { dateStr, year, month, day };
      }
    }
  }
  
  // 월-일 형식 (년도 없음) - 현재 년도로 간주
  const shortPatterns = [
    /^(\d{1,2})-(\d{1,2})$/,
    /^(\d{1,2})\.(\d{1,2})$/,
    /^(\d{1,2})\/(\d{1,2})$/,
  ];
  
  for (const pattern of shortPatterns) {
    const match = folderName.match(pattern);
    if (match) {
      const currentYear = new Date().getFullYear();
      const month = parseInt(match[1]) - 1; // 0-based
      const day = parseInt(match[2]);
      
      // 유효한 날짜인지 검증
      if (month >= 0 && month <= 11 && day >= 1 && day <= 31) {
        const dateStr = `${currentYear}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        return { dateStr, year: currentYear, month, day };
      }
    }
  }
  
  return null;
}

/**
 * 🆕 날짜 마감 확인 함수
 * 해당 날짜의 마감 시각이 지났는지 확인
 * @param {string} dateStr - 확인할 날짜 (yyyy-MM-dd)
 * @returns {boolean} 마감되었으면 true
 * * 예시:
 * - 10월 16일 폴더 마감: 10월 17일 03:00:00
 * - 17일 02:59:59까지는 false 반환 (스캔함)
 * - 17일 03:00:00부터는 true 반환 (스캔 안함)
 */
function 날짜마감확인(dateStr) {
  const now = new Date();

  // dateStr을 Date 객체로 변환 (예: "2025-10-16")
  const targetDate = new Date(dateStr + 'T00:00:00+09:00'); // 한국시간

  // 마감시각 = 해당 날짜 다음날 새벽 3시
  // 예: 16일 폴더 → 17일 03:00:00에 마감
  const deadline = new Date(targetDate);
  deadline.setDate(deadline.getDate() + 1); // 다음날
  deadline.setHours(CONFIG.DEADLINE_HOUR, 0, 0, 0); // 새벽 3시

  // 현재 시각이 마감시각 이후면 true (마감됨)
  const isClosed = now >= deadline;

  if (isClosed) {
    Logger.log(`      [마감체크] ${dateStr} 마감됨 (마감: ${deadline.toLocaleString('ko-KR')})`);
  }

  return isClosed;
}


// ==================== 보조 함수 ====================

/**
 * 오늘 날짜 폴더 찾기
 * 다양한 날짜 형식을 지원합니다
 */
function 오늘날짜폴더찾기(parentFolder, targetDate) {
  const folders = parentFolder.getFolders();
  
  // targetDate: "2025-10-15" 형식
  const year = targetDate.substring(0, 4);
  const month = targetDate.substring(5, 7);
  const day = targetDate.substring(8, 10);
  
  // 인식 가능한 형식들
  const validFormats = [
    targetDate,                       // 2025-10-15
    targetDate.replace(/-/g, ''),     // 20251015
    `${year}.${month}.${day}`,        // 2025.10.15
    `${year}/${month}/${day}`,        // 2025/10/15
    `${month}-${day}`,                // 10-15
    `${month}.${day}`,                // 10.15
    `${month}/${day}`,                // 10/15
    `${parseInt(month)}-${parseInt(day)}`, // 10-15 (앞자리 0 제거)
    `${parseInt(month)}.${parseInt(day)}`, // 10.15 (앞자리 0 제거)
    `${parseInt(month)}/${parseInt(day)}`, // 10/15 (앞자리 0 제거)
  ];
  
  while (folders.hasNext()) {
    const folder = folders.next();
    const folderName = folder.getName().trim();
    
    // 모든 가능한 형식과 비교
    for (const format of validFormats) {
      if (folderName === format) {
        Logger.log(`    📂 발견: "${folderName}" (형식: ${format})`);
        return folder;
      }
    }
  }
  
  return null;
}


/**
 * 폴더 내 모든 파일의 링크 생성
 */
function 파일목록및링크생성(folder) {
  const files = [];
  const fileIterator = folder.getFiles();

  while (fileIterator.hasNext()) {
    const file = fileIterator.next();
    const fileName = file.getName();
    
    // 파일 공유 설정 시도 (실패해도 무시)
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      // 에러를 로그에 남기지 않음
    }
    
    files.push({
      name: file.getName(),
      url: file.getUrl(),
      mimeType: file.getMimeType(),
      size: file.getSize()
    });
  }
  
  return files;
}

/**
 * Google Sheets에 출석 기록 추가 (폴더 ID 직접 전달 버전)
 */
function 출석기록추가(memberName, date, files, status = 'O', reason = '', folderId = '') {
  const koreaTime = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.appendRow(['타임스탬프', '이름', '날짜', '파일수', '링크', '폴더링크', '출석상태', '주차', '사유']);
    sheet.getRange('A1:I1').setFontWeight('bold').setBackground('#4CAF50').setFontColor('white');
  }
  
  const linksText = files.map(f => `${f.name}: ${f.url}`).join('\n');
  
  // 폴더 링크 생성
  const folderLink = (status === 'O' && folderId) ? 
    `https://drive.google.com/drive/folders/${folderId}` : '';
  
  const weekNum = 주차계산(new Date(date));
  let displayText = '';
  if (status === 'OFF') displayText = '오프';
  else if (status === CONFIG.LONG_OFF_STATUS) displayText = '장기오프';
  else if (status === 'X') displayText = '결석';

  const data = sheet.getDataRange().getValues();
  let existingRow = -1;
  
  const targetDateStr = typeof date === 'string' 
    ? date 
    : Utilities.formatDate(new Date(date), 'Asia/Seoul', 'yyyy-MM-dd');
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === memberName) {
      const recordDateStr = typeof data[i][2] === 'string'
        ? data[i][2]
        : Utilities.formatDate(new Date(data[i][2]), 'Asia/Seoul', 'yyyy-MM-dd');
      
      if (recordDateStr === targetDateStr) {
        existingRow = i + 1;
        break;
      }
    }
  }

  if (existingRow !== -1) {
    const oldStatus = data[existingRow - 1][6];
    Logger.log(`  🔄 ${memberName} ${date} 기록 업데이트 (이전: ${oldStatus}, 새상태: ${status})`);
    const range = sheet.getRange(existingRow, 1, 1, 9);
    range.setValues([
      [koreaTime, memberName, date, files.length, linksText || displayText, folderLink, status, weekNum, reason]
    ]);
    if (status === 'X') range.setBackground('#ffcdd2');
    else if (status === CONFIG.LONG_OFF_STATUS) range.setBackground('#e1f5fe');
    else range.setBackground(null);
    
  } else {
    sheet.appendRow([
      koreaTime, memberName, date, files.length, linksText || displayText, folderLink, status, weekNum, reason
    ]);
    
    const lastRow = sheet.getLastRow();
    const range = sheet.getRange(lastRow, 1, 1, 9);
    if (status === 'X') range.setBackground('#ffcdd2');
    else if (status === CONFIG.LONG_OFF_STATUS) range.setBackground('#e1f5fe');
  }
  
  const statusText = status === CONFIG.LONG_OFF_STATUS ? '장기오프' :
                     status === 'OFF' ? '오프' :
                     status === 'X' ? '결석' : '출석';
  Logger.log(`  ✓ ${memberName} ${statusText} 기록 처리 완료`);
}

/**
 * 날짜로 주차 계산
 */
function 주차계산(date) {
  const firstDay = new Date(date.getFullYear(), date.getMonth(), 1);
  const dayOfMonth = date.getDate();
  const firstDayOfWeek = firstDay.getDay();
  
  const weekNumber = Math.ceil((dayOfMonth + firstDayOfWeek) / 7);
  return weekNumber;
}

/**
 * Google Drive URL에서 폴더 ID 추출
 */
function 폴더ID추출(fileUrl) {
  try {
    const fileId = fileUrl.match(/[-\w]{25,}/);
    if (!fileId) return '';
    
    const file = DriveApp.getFileById(fileId[0]);
    const folders = file.getParents();
    
    if (folders.hasNext()) {
      return folders.next().getId();
    }
  } catch (e) {
    Logger.log(`폴더 ID 추출 오류: ${e.message}`);
  }
  
  return '';
}


/**
 * JSON 파일 생성 및 공개 (수정 버전)
 * 🔧 현재 월 데이터만 포함하도록 수정
 */
function JSON파일생성() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recordSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!recordSheet) {
    Logger.log('제출기록 시트가 없습니다.');
    return;
  }
  
  const records = recordSheet.getDataRange().getValues();
  const jsonData = {};
  
  // 🆕 현재 연월 계산
  const today = new Date();
  const currentYearMonth = Utilities.formatDate(today, 'Asia/Seoul', 'yyyy-MM');
  
  Logger.log(`JSON 생성: ${currentYearMonth} 데이터만 포함`);
  
  // 조원별로 데이터 구조화
  for (const memberName of Object.keys(CONFIG.MEMBERS)) {
    jsonData[memberName] = {
      출석: 0,
      결석: 0,
      오프: 0,
      장기오프: 0,
      경고: false,
      벌칙: false,
      기록: {},
      주간통계: {}
    };
  }
  
  // 기록 파싱
  for (let i = 1; i < records.length; i++) {
    const [timestamp, name, dateStr, fileCount, links, folderLink, status, weekNum, reason] = records[i];
    
    if (!jsonData[name]) continue;
    
    // 🆕 날짜 문자열 정규화
    const dateFormatted = typeof dateStr === 'string' 
      ? dateStr 
      : Utilities.formatDate(new Date(dateStr), 'Asia/Seoul', 'yyyy-MM-dd');
    
    // 🆕 현재 월 데이터만 필터링
    if (!dateFormatted.startsWith(currentYearMonth)) {
      continue;  // 다른 월 데이터는 건너뛰기
    }
    
    const date = new Date(dateFormatted);
    const day = date.getDate().toString();
    
    // 날짜별 기록
    jsonData[name].기록[day] = {
      status: status,
      link: folderLink || (links ? (links.split('\n')[0].split(': ')[1] || links) : ''),
      fileCount: fileCount || 0,
      reason: reason || ''
    };
    
    // 출석/결석/오프/장기오프 카운트
    if (status === 'O') {
      jsonData[name].출석++;
    } else if (status === 'OFF') {
      jsonData[name].오프++;
    } else if (status === CONFIG.LONG_OFF_STATUS) {
      jsonData[name].장기오프++;
    } else {
      jsonData[name].결석++;
    }
    
    // 주간 통계
    const weekKey = `${weekNum}주차`;
    if (!jsonData[name].주간통계[weekKey]) {
      jsonData[name].주간통계[weekKey] = {
        출석: 0,
        결석: 0,
        오프: 0,
        장기오프: 0
      };
    }
    
    if (status === 'O') {
      jsonData[name].주간통계[weekKey].출석++;
    } else if (status === 'OFF') {
      jsonData[name].주간통계[weekKey].오프++;
    } else if (status === CONFIG.LONG_OFF_STATUS) {
      jsonData[name].주간통계[weekKey].장기오프++;
    } else {
      jsonData[name].주간통계[weekKey].결석++;
    }
  }
  
  // 경고/벌칙 판정 (오프와 장기오프는 제외)
  for (const [name, data] of Object.entries(jsonData)) {
    if (data.결석 >= 4) {
      data.벌칙 = true;
    } else if (data.결석 === 3) {
      data.경고 = true;
    }
  }
  
  // JSON 파일로 저장
  const jsonString = JSON.stringify(jsonData, null, 2);
  const fileName = `attendance_summary_${currentYearMonth}.json`;
  
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
    
    Logger.log(`JSON 파일 생성 완료: ${file.getUrl()}`);
  } catch (error) {
    Logger.log(`JSON 파일 생성 오류: ${error.message}`);
  }
}

// ==================== 초기 설정 ====================

/**
 * 최초 1회만 실행
 * 트리거 설정 및 시트 초기화
 */
function 초기설정() {
  // 🆕 스프레드시트 시간대를 한국으로 설정
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Seoul');
  Logger.log('✅ 스프레드시트 시간대: Asia/Seoul (한국 시간)');
  Logger.log('');
  
  // 기존 트리거 삭제
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // 트리거 1: 1시간마다 출석 체크
  ScriptApp.newTrigger('출석체크_메인')
    .timeBased()
    .everyHours(1)
    .create();
  
  Logger.log('트리거 1 설정 완료: 매 1시간마다 출석 체크');
  
  // 트리거 2: 매일 새벽 3시 마감시간 체크
  ScriptApp.newTrigger('마감시간체크')
    .timeBased()
    .atHour(CONFIG.DEADLINE_HOUR)
    .everyDays(1)
    .create();
  
  Logger.log(`트리거 2 설정 완료: 매일 새벽 ${CONFIG.DEADLINE_HOUR}시 마감 체크`);
  
  // 🆕 트리거 3: 1시간마다 관리자 수정 자동 처리
  ScriptApp.newTrigger('관리자수정_자동처리')
    .timeBased()
    .everyHours(1)
    .create();
  
  Logger.log('트리거 3 설정 완료: 매 1시간마다 관리자 수정 자동 처리');
  
  // 시트 초기화
  
  // 🆕 트리거 4: 매월 1일 오전 1시에 전월 결산 생성
  ScriptApp.newTrigger('월별결산생성')
    .timeBased()
    .onMonthDay(1)
    .atHour(1)
    .create();
  
  Logger.log('트리거 4 설정 완료: 매월 1일 오전 1시 전월 결산 생성');
  
  // 제출기록 시트
  let recordSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!recordSheet) {
    recordSheet = ss.insertSheet(CONFIG.SHEET_NAME);
    recordSheet.appendRow(['타임스탬프', '이름', '날짜', '파일수', '링크', '폴더링크', '출석상태', '주차', '사유']);
    recordSheet.getRange('A1:I1').setFontWeight('bold').setBackground('#4CAF50').setFontColor('white');
  }
  
  // 🆕 장기오프신청 시트 확인 (구글 폼으로 자동 생성되어야 함)
  let longOffSheet = ss.getSheetByName(CONFIG.LONG_OFF_SHEET);
  if (!longOffSheet) {
    Logger.log('⚠️ 장기오프신청 시트가 없습니다.');
    Logger.log('구글 폼에서 "스프레드시트에 연결"을 먼저 설정해주세요.');
    Logger.log('시트 이름: ' + CONFIG.LONG_OFF_SHEET);
  } else {
    Logger.log('✅ 장기오프신청 시트 발견: ' + CONFIG.LONG_OFF_SHEET);
    
    // 승인 열(G열) 확인 및 추가
    const lastCol = longOffSheet.getLastColumn();
    if (lastCol < 7) {
      Logger.log('⚠️ 승인 열(G열)이 없습니다. 수동으로 추가해주세요.');
    } else {
      Logger.log('✅ 승인 열(G열) 확인 완료');
    }
  }
  
  // 🆕 관리자수정 시트 생성
  관리자수정시트_생성();
  Logger.log('✅ 관리자수정 시트 생성 완료');
  
  // 🆕 월별결산 시트 초기화
  let summarySheet = ss.getSheetByName(CONFIG.MONTHLY_SUMMARY_SHEET);
  if (!summarySheet) {
    summarySheet = ss.insertSheet(CONFIG.MONTHLY_SUMMARY_SHEET);
    const headers = ['연월', '조원명', '출석', '오프', '장기오프', '결석', '출석률(%)', '상태', '비고'];
    summarySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    summarySheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4CAF50').setFontColor('white');
    summarySheet.setFrozenRows(1);
    summarySheet.setColumnWidth(1, 100);
    summarySheet.setColumnWidth(2, 150);
    summarySheet.setColumnWidths(3, 4, 80);
    summarySheet.setColumnWidth(7, 100);
    summarySheet.setColumnWidth(8, 120);
    summarySheet.setColumnWidth(9, 200);
    Logger.log('✅ 월별결산 시트 생성 완료');
  } else {
    Logger.log('✅ 월별결산 시트 이미 존재함');
  }
  
  Logger.log('초기 설정 완료!');
  Logger.log('');
  Logger.log('⚠️ 다음 작업 필요:');
  Logger.log('1. CONFIG.MEMBERS에 각 조원의 Google Drive 폴더 ID 입력');
  Logger.log('    💡 여러 폴더 사용 시: [\'폴더ID1\', \'폴더ID2\'] 형식으로');
  Logger.log('    예시: \'길\': [\'집_폴더ID\', \'직장_폴더ID\']');
  Logger.log('2. CONFIG.JSON_FOLDER_ID에 JSON 출력 폴더 ID 입력');
  Logger.log('3. 각 조원이 자신의 폴더를 대표원장님과 공유');
  Logger.log('4. 구글 폼을 스프레드시트에 연결 (시트명: ' + CONFIG.LONG_OFF_SHEET + ')');
  Logger.log('5. 장기오프 시트에 G열 "승인" 헤더 추가');
  Logger.log('');
  Logger.log('💡 오프 사용법:');
  Logger.log('- 오프하려는 날 폴더에 OFF.md 파일 생성');
  Logger.log('- 주당 3회까지 가능');
  Logger.log('- 3회 초과 시 모두 결석 처리');
  Logger.log('');
  Logger.log('⏰ 마감시간 제도:');
  Logger.log(`- 매일 새벽 ${CONFIG.DEADLINE_HOUR}시가 마감시간`);
  Logger.log('- 전날 인증하지 않으면 자동 결석 처리');
  Logger.log('');
  Logger.log('🆕 장기오프 제도 (구글 폼):');
  Logger.log('- 구글 폼으로 신청');
  Logger.log('- 여행, 출장 등 장기 사유에 활용');
  Logger.log('- 주간 오프 카운트에서 제외');
  Logger.log(`- 자동 승인: ${CONFIG.LONG_OFF_AUTO_APPROVE ? '활성화' : '수동 승인 필요'}`);
  Logger.log('');
  Logger.log('🆕 여러 폴더 지원:');
  Logger.log('- 집/직장 등 여러 위치에서 작업하는 조원');
  Logger.log('- 배열로 여러 폴더 ID 설정 가능');
  Logger.log('');
  Logger.log('📊 월별결산 기능:');
  Logger.log('- 매월 1일 오전 1시에 전월 결산 자동 생성');
  Logger.log('- "월별결산" 시트에서 월별 통계 확인');
  Logger.log('- 조원별 출석률, 경고/벌칙 상태 한눈에 확인');
  Logger.log('- 중복 날짜는 자동으로 하나만 처리');
}

/**
 * 수동 테스트용
 */
function 테스트실행() {
  출석체크_메인();
}

/**
 * 마감시간 체크 수동 테스트
 */
function 마감시간체크_테스트() {
  마감시간체크();
}

/**
 * 🆕 장기오프 테스트 (구글 폼 버전)
 * 특정 날짜와 조원의 장기오프 상태 확인
 */
function 장기오프테스트() {
  const 테스트조원 = '센트룸';  // ← 테스트할 조원 이름
  const 테스트날짜 = '2025-10-22';  // ← 테스트할 날짜
  
  Logger.log(`=== 장기오프 테스트: ${테스트조원}, ${테스트날짜} ===`);
  
  const result = 장기오프확인(테스트조원, 테스트날짜);
  
  if (result.isLongOff) {
    Logger.log(`✅ 장기오프 기간입니다!`);
    Logger.log(`    사유: ${result.reason}`);
  } else {
    Logger.log(`❌ 장기오프 기간이 아닙니다.`);
  }
  
  Logger.log('=== 테스트 완료 ===');
  
  // 폼 응답 시트 구조 확인
  Logger.log('');
  Logger.log('=== 폼 응답 시트 구조 확인 ===');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.LONG_OFF_SHEET);
  
  if (sheet) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log('헤더: ' + JSON.stringify(headers));
    
    if (sheet.getLastRow() > 1) {
      const firstData = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
      Logger.log('첫 번째 데이터: ' + JSON.stringify(firstData));
    }
  } else {
    Logger.log('⚠️ 시트를 찾을 수 없음: ' + CONFIG.LONG_OFF_SHEET);
  }
}

/**
 * 🆕 관리자용: 수동 출석 기록 추가/수정
 * 시스템 오류나 특수 상황에서 관리자가 직접 출석 상태를 수정
 * * @param {string} memberName - 조원 이름
 * @param {string} date - 날짜 (YYYY-MM-DD)
 * @param {string} status - 출석 상태 ('O', 'OFF', 'LONG_OFF', 'X')
 * @param {string} reason - 사유 (선택)
 * @param {boolean} overwrite - 기존 기록 덮어쓰기 (true: 덮어쓰기, false: 중복 방지)
 */
function 관리자_출석수정(memberName, date, status = 'O', reason = '관리자 수정', overwrite = true) {
  Logger.log('=== 관리자 출석 기록 수정 ===');
  Logger.log(`조원: ${memberName}`);
  Logger.log(`날짜: ${date}`);
  Logger.log(`상태: ${status}`);
  Logger.log(`사유: ${reason}`);
  Logger.log(`덮어쓰기: ${overwrite}`);
  Logger.log('');
  
  // 유효성 검사
  if (!CONFIG.MEMBERS[memberName]) {
    Logger.log(`❌ 오류: "${memberName}"은(는) 등록된 조원이 아닙니다.`);
    Logger.log('등록된 조원: ' + Object.keys(CONFIG.MEMBERS).join(', '));
    return;
  }
  
  const validStatuses = ['O', 'OFF', 'LONG_OFF', 'X'];
  if (!validStatuses.includes(status)) {
    Logger.log(`❌ 오류: "${status}"은(는) 유효하지 않은 상태입니다.`);
    Logger.log('유효한 상태: O (출석), OFF (오프), LONG_OFF (장기오프), X (결석)');
    return;
  }
  
  // 날짜 형식 검증
  const datePattern = /^\d{4}-\d{2}-\d{2}$/;
  if (!datePattern.test(date)) {
    Logger.log(`❌ 오류: "${date}"은(는) 유효하지 않은 날짜 형식입니다.`);
    Logger.log('올바른 형식: YYYY-MM-DD (예: 2025-10-15)');
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    Logger.log('❌ 오류: 제출기록 시트가 없습니다.');
    return;
  }
  
  // 기존 기록 확인
  const data = sheet.getDataRange().getValues();
  let existingRow = -1;
  
  for (let i = 1; i < data.length; i++) {
    const [, recordName, recordDate] = data[i];
    if (recordName === memberName && recordDate === date) {
      existingRow = i + 1; // 1-based index
      break;
    }
  }
  
  if (existingRow > 0) {
    if (overwrite) {
      // 기존 기록 수정
      const oldStatus = data[existingRow - 1][6];
      const oldReason = data[existingRow - 1][8] || '';
      
      Logger.log(`📝 기존 기록 발견 (행 ${existingRow})`);
      Logger.log(`    이전 상태: ${oldStatus}${oldReason ? ' (' + oldReason + ')' : ''}`);
      Logger.log(`    새 상태: ${status}${reason ? ' (' + reason + ')' : ''}`);
      
      // 상태 업데이트
      sheet.getRange(existingRow, 7).setValue(status); // G열: 출석상태
      
      // 링크 열 업데이트
      let linkText = '';
      if (status === 'OFF') {
        linkText = '오프';
      } else if (status === CONFIG.LONG_OFF_STATUS) {
        linkText = '장기오프';
      } else if (status === 'X') {
        linkText = '결석';
      } else {
        linkText = '관리자 수정 (출석 처리)';
      }
      sheet.getRange(existingRow, 5).setValue(linkText); // E열: 링크
      
      // 사유 업데이트
      sheet.getRange(existingRow, 9).setValue(reason); // I열: 사유
      
      // 배경색 설정
      if (status === 'X') {
        sheet.getRange(existingRow, 1, 1, 9).setBackground('#ffcdd2'); // 결석: 빨간색
      } else if (status === CONFIG.LONG_OFF_STATUS) {
        sheet.getRange(existingRow, 1, 1, 9).setBackground('#e1f5fe'); // 장기오프: 파란색
      } else if (status === 'OFF') {
        sheet.getRange(existingRow, 1, 1, 9).setBackground('#fff9c4'); // 오프: 노란색
      } else {
        sheet.getRange(existingRow, 1, 1, 9).setBackground('#c8e6c9'); // 출석: 초록색
      }
      
      Logger.log('✅ 기록 수정 완료!');
      
    } else {
      Logger.log(`⚠️ 기록이 이미 존재합니다 (행 ${existingRow})`);
      Logger.log('    덮어쓰기를 원하시면 overwrite=true로 설정하세요.');
      return;
    }
    
  } else {
    // 새 기록 추가
    Logger.log('📝 새 기록 추가');
    
    출석기록추가(memberName, date, [], status, reason);
    
    Logger.log('✅ 기록 추가 완료!');
  }
  
  // JSON 파일 재생성
  JSON파일생성();
  
  Logger.log('');
  Logger.log('=== 수정 완료 ===');
}

/**
 * 🆕 관리자용: 일괄 출석 수정
 * 여러 조원의 특정 날짜 출석을 한 번에 수정
 * * @param {Array} records - 수정할 기록 배열
 * 예: [
 * {name: '센트룸', date: '2025-10-15', status: 'O', reason: '시스템 오류'},
 * {name: '길', date: '2025-10-16', status: 'OFF', reason: '정전'}
 * ]
 */
function 관리자_일괄수정(records) {
  Logger.log('=== 관리자 일괄 출석 수정 ===');
  Logger.log(`총 ${records.length}개 기록 수정`);
  Logger.log('');
  
  let successCount = 0;
  let failCount = 0;
  
  records.forEach((record, index) => {
    Logger.log(`[${index + 1}/${records.length}] 처리 중...`);
    
    try {
      관리자_출석수정(
        record.name,
        record.date,
        record.status || 'O',
        record.reason || '관리자 일괄 수정',
        record.overwrite !== false // 기본값: true
      );
      successCount++;
      Logger.log('');
    } catch (e) {
      Logger.log(`❌ 오류: ${e.message}`);
      Logger.log('');
      failCount++;
    }
  });
  
  Logger.log('=== 일괄 수정 완료 ===');
  Logger.log(`성공: ${successCount}개 / 실패: ${failCount}개`);
}

/**
 * 🆕 관리자용: 출석 기록 삭제
 * 잘못 입력된 기록을 완전히 삭제
 * * @param {string} memberName - 조원 이름
 * @param {string} date - 날짜 (YYYY-MM-DD)
 */
function 관리자_기록삭제(memberName, date) {
  Logger.log('=== 관리자 출석 기록 삭제 ===');
  Logger.log(`조원: ${memberName}`);
  Logger.log(`날짜: ${date}`);
  Logger.log('');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    Logger.log('❌ 오류: 제출기록 시트가 없습니다.');
    return;
  }
  
  // 기존 기록 찾기
  const data = sheet.getDataRange().getValues();
  let targetRow = -1;
  
  for (let i = 1; i < data.length; i++) {
    const [, recordName, recordDate] = data[i];
    if (recordName === memberName && recordDate === date) {
      targetRow = i + 1; // 1-based index
      break;
    }
  }
  
  if (targetRow > 0) {
    const oldStatus = data[targetRow - 1][6];
    const oldReason = data[targetRow - 1][8] || '';
    
    Logger.log(`📝 기록 발견 (행 ${targetRow})`);
    Logger.log(`    상태: ${oldStatus}${oldReason ? ' (' + oldReason + ')' : ''}`);
    
    // 행 삭제
    sheet.deleteRow(targetRow);
    
    Logger.log('✅ 기록 삭제 완료!');
    
    // JSON 파일 재생성
    JSON파일생성();
    
  } else {
    Logger.log(`❌ 해당 기록을 찾을 수 없습니다.`);
    Logger.log(`    조원: ${memberName}`);
    Logger.log(`    날짜: ${date}`);
  }
  
  Logger.log('');
  Logger.log('=== 삭제 완료 ===');
}

/**
 * 🆕 관리자용: 특정 조원의 모든 출석 기록 조회
 * * @param {string} memberName - 조원 이름
 * @param {string} month - 월 필터 (선택, 예: '2025-10')
 */
function 관리자_기록조회(memberName, month = null) {
  Logger.log('=== 출석 기록 조회 ===');
  Logger.log(`조원: ${memberName}`);
  if (month) {
    Logger.log(`월: ${month}`);
  }
  Logger.log('');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    Logger.log('❌ 오류: 제출기록 시트가 없습니다.');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  const records = [];
  
  for (let i = 1; i < data.length; i++) {
    const [timestamp, name, dateStr, fileCount, links, folderLink, status, weekNum, reason] = data[i];
    
    if (name === memberName) {
      // 월 필터
      if (month && !dateStr.startsWith(month)) {
        continue;
      }
      
      records.push({
        row: i + 1,
        date: dateStr,
        status: status,
        fileCount: fileCount,
        reason: reason || '',
        timestamp: timestamp
      });
    }
  }
  
  if (records.length === 0) {
    Logger.log('📝 기록이 없습니다.');
    return;
  }
  
  Logger.log(`📊 총 ${records.length}개 기록 발견:`);
  Logger.log('');
  
  // 날짜순 정렬
  records.sort((a, b) => a.date.localeCompare(b.date));
  
  // 통계
  const stats = {
    O: 0,
    OFF: 0,
    LONG_OFF: 0,
    X: 0
  };
  
  records.forEach(record => {
    const statusText = 
      record.status === 'O' ? '출석' :
      record.status === 'OFF' ? '오프' :
      record.status === CONFIG.LONG_OFF_STATUS ? '장기오프' :
      '결석';
    
    Logger.log(`  [행 ${record.row}] ${record.date} - ${statusText}${record.reason ? ' (' + record.reason + ')' : ''}`);
    
    if (stats[record.status] !== undefined) {
      stats[record.status]++;
    }
  });
  
  Logger.log('');
  Logger.log('📈 통계:');
  Logger.log(`    출석: ${stats.O}일`);
  Logger.log(`    오프: ${stats.OFF}일`);
  Logger.log(`    장기오프: ${stats.LONG_OFF}일`);
  Logger.log(`    결석: ${stats.X}일`);
  Logger.log('');
  Logger.log('=== 조회 완료 ===');
  
  return records;
}
/**
 * 🆕 구글 폼 응답 시트 구조 확인
 */
function 폼응답시트구조확인() {
  Logger.log('=== 폼 응답 시트 구조 확인 ===');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.LONG_OFF_SHEET);
  
  if (!sheet) {
    Logger.log('❌ 시트를 찾을 수 없음: ' + CONFIG.LONG_OFF_SHEET);
    Logger.log('구글 폼을 스프레드시트에 연결했는지 확인하세요.');
    return;
  }
  
  Logger.log('✅ 시트 발견: ' + CONFIG.LONG_OFF_SHEET);
  Logger.log('');
  
  // 헤더 확인
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  Logger.log('📊 열 구조:');
  headers.forEach((header, index) => {
    const columnLetter = String.fromCharCode(65 + index); // A, B, C, ...
    Logger.log(`  ${columnLetter}열 (index ${index}): ${header}`);
  });
  
  Logger.log('');
  
  // 첫 번째 데이터 확인
  if (sheet.getLastRow() > 1) {
    Logger.log('📝 첫 번째 데이터:');
    const firstData = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
    firstData.forEach((value, index) => {
      const columnLetter = String.fromCharCode(65 + index);
      Logger.log(`  ${columnLetter}열: ${value} (타입: ${typeof value})`);
    });
  } else {
    Logger.log('⚠️ 데이터가 없습니다. 구글 폼에서 테스트 신청을 해보세요.');
  }
  
  Logger.log('');
  Logger.log('💡 CONFIG.FORM_COLUMNS 설정 확인:');
  Logger.log(JSON.stringify(CONFIG.FORM_COLUMNS, null, 2));
}

// ==================== 🆕 간편 관리자 수정 시스템 ====================

/**
 * 🆕 관리자수정 시트를 통한 자동 수정
 * 시트에 입력하면 자동으로 출석 기록 수정
 * * 사용법:
 * 1. "관리자수정" 시트 열기
 * 2. 조원 이름, 날짜, 상태, 사유 입력
 * 3. 저장하면 자동 처리
 * 4. 처리 결과는 "처리상태" 열에 표시
 */
function 관리자수정_자동처리() {
  Logger.log('=== 관리자 수정 자동 처리 시작 ===');
  Logger.log('');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('관리자수정');
  
  // 시트가 없으면 생성
  if (!sheet) {
    Logger.log('⚠️ "관리자수정" 시트가 없습니다. 생성합니다...');
    관리자수정시트_생성();
    Logger.log('✅ "관리자수정" 시트 생성 완료');
    Logger.log('');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    Logger.log('처리할 수정 요청이 없습니다.');
    return;
  }
  
  let processCount = 0;
  let successCount = 0;
  let failCount = 0;
  
  // 2행부터 처리 (1행은 헤더)
  for (let i = 1; i < data.length; i++) {
    const [조원, 날짜, 상태, 사유, 처리상태] = data[i];
    
    // 이미 처리된 행은 스킵
    if (처리상태 === '✅ 완료' || 처리상태 === '⏭️ 스킵') {
      continue;
    }
    
    // 빈 행은 스킵
    if (!조원 || !날짜 || !상태) {
      continue;
    }
    
    processCount++;
    const rowNum = i + 1;
    
    Logger.log(`[${processCount}] 처리 중: ${조원} / ${날짜} / ${상태}`);
    
    try {
      // 유효성 검사
      if (!CONFIG.MEMBERS[조원]) {
        throw new Error(`"${조원}"은(는) 등록된 조원이 아닙니다.`);
      }
      
      const validStatuses = ['O', 'OFF', 'LONG_OFF', 'X', '출석', '오프', '장기오프', '결석'];
      if (!validStatuses.includes(상태)) {
        throw new Error(`"${상태}"은(는) 유효하지 않은 상태입니다.`);
      }
      
      // 한글 상태를 영문 코드로 변환
      let statusCode = 상태;
      if (상태 === '출석') statusCode = 'O';
      else if (상태 === '오프') statusCode = 'OFF';
      else if (상태 === '장기오프') statusCode = 'LONG_OFF';
      else if (상태 === '결석') statusCode = 'X';
      
      // 출석 수정 실행
      관리자_출석수정(
        조원,
        날짜,
        statusCode,
        사유 || '관리자 수정',
        true
      );
      
      // 처리 완료 표시
      sheet.getRange(rowNum, 5).setValue('✅ 완료');
      sheet.getRange(rowNum, 6).setValue(
        Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')
      );
      
      // 완료된 행은 연한 초록색
      sheet.getRange(rowNum, 1, 1, 6).setBackground('#c8e6c9');
      
      Logger.log(`    ✅ 성공`);
      successCount++;
      
    } catch (e) {
      Logger.log(`    ❌ 실패: ${e.message}`);
      
      // 에러 표시
      sheet.getRange(rowNum, 5).setValue('❌ 실패: ' + e.message);
      
      // 실패한 행은 연한 빨간색
      sheet.getRange(rowNum, 1, 1, 6).setBackground('#ffcdd2');
      
      failCount++;
    }
    
    Logger.log('');
  }
  
  Logger.log('=== 자동 처리 완료 ===');
  Logger.log(`처리: ${processCount}개 / 성공: ${successCount}개 / 실패: ${failCount}개`);
}

/**
 * 🆕 관리자수정 시트 생성
 */
function 관리자수정시트_생성() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 기존 시트가 있으면 삭제
  const existingSheet = ss.getSheetByName('관리자수정');
  if (existingSheet) {
    ss.deleteSheet(existingSheet);
  }
  
  // 새 시트 생성
  const sheet = ss.insertSheet('관리자수정');
  
  // 헤더 설정
  const headers = ['조원 이름', '날짜 (YYYY-MM-DD)', '상태', '사유 (선택)', '처리상태', '처리시간'];
  sheet.appendRow(headers);
  
  // 헤더 스타일
  sheet.getRange('A1:F1')
    .setFontWeight('bold')
    .setBackground('#FF9800')
    .setFontColor('white')
    .setHorizontalAlignment('center');
  
  // 열 너비 조정
  sheet.setColumnWidth(1, 100);  // 조원 이름
  sheet.setColumnWidth(2, 150);  // 날짜
  sheet.setColumnWidth(3, 100);  // 상태
  sheet.setColumnWidth(4, 300);  // 사유
  sheet.setColumnWidth(5, 250);  // 처리상태
  sheet.setColumnWidth(6, 150);  // 처리시간
  
  // 예시 데이터 3개 추가
  const examples = [
    ['센트룸', '2025-10-15', 'O', 'Google Drive 동기화 오류', '', ''],
    ['길', '2025-10-16', '출석', '정전으로 업로드 지연', '', ''],
    ['what', '2025-10-17', 'OFF', '긴급 병원 진료', '', '']
  ];
  
  examples.forEach(example => {
    sheet.appendRow(example);
  });
  
  // 예시 데이터는 연한 노란색
  sheet.getRange(2, 1, 3, 6).setBackground('#fff9c4');
  
  // 안내문 추가
  sheet.getRange('A5').setValue('📝 사용 방법:');
  sheet.getRange('A6').setValue('1. 위 예시를 참고하여 새 행에 정보 입력');
  sheet.getRange('A7').setValue('2. 상태는 "O", "OFF", "LONG_OFF", "X" 또는 "출석", "오프", "장기오프", "결석" 입력');
  sheet.getRange('A8').setValue('3. 트리거가 1시간마다 자동 처리하거나, "관리자수정_자동처리" 함수 직접 실행');
  sheet.getRange('A9').setValue('4. 처리 완료되면 "처리상태" 열에 ✅ 표시됨');
  sheet.getRange('A10').setValue('');
  sheet.getRange('A11').setValue('⚠️ 주의: 예시 데이터는 삭제하거나 "처리상태"를 "⏭️ 스킵"으로 변경하세요');
  
  sheet.getRange('A5:A11').setFontWeight('bold').setFontColor('#666666');
  
  // 데이터 유효성 검사 (상태 열)
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['O', 'OFF', 'LONG_OFF', 'X', '출석', '오프', '장기오프', '결석'], true)
    .setAllowInvalid(false)
    .setHelpText('출석 상태를 선택하세요')
    .build();
  
  sheet.getRange('C2:C1000').setDataValidation(statusRule);
  
  // 조원 이름 자동완성 (드롭다운)
  const memberNames = Object.keys(CONFIG.MEMBERS);
  const memberRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(memberNames, true)
    .setAllowInvalid(false)
    .setHelpText('조원 이름을 선택하세요')
    .build();
  
  sheet.getRange('A2:A1000').setDataValidation(memberRule);
  
  // 시트 보호 (처리상태, 처리시간 열은 수정 불가)
  const protection = sheet.protect().setDescription('관리자수정 시트 보호');
  protection.setUnprotectedRanges([
    sheet.getRange('A2:D1000')  // 조원, 날짜, 상태, 사유만 편집 가능
  ]);
  
  Logger.log('✅ 관리자수정 시트 생성 완료');
}

/**
 * 🆕 관리자수정 시트 초기화
 * 완료된 항목 삭제 및 시트 정리
 */
function 관리자수정시트_초기화() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('관리자수정');
  
  if (!sheet) {
    Logger.log('⚠️ "관리자수정" 시트가 없습니다.');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  
  // 뒤에서부터 삭제 (행 번호가 밀리지 않도록)
  let deletedCount = 0;
  
  for (let i = data.length - 1; i >= 1; i--) {
    const [조원, 날짜, 상태, 사유, 처리상태] = data[i];
    
    // 완료된 행 또는 빈 행 삭제
    if (처리상태 === '✅ 완료' || (!조원 && !날짜 && !상태)) {
      sheet.deleteRow(i + 1);
      deletedCount++;
    }
  }
  
  Logger.log(`✅ ${deletedCount}개 행 삭제 완료`);
}

// ==================== 관리자 함수 사용 예시 ====================

/**
 * 예시 1: 단일 기록 수정
 * 센트룸의 10월 15일 출석을 출석(O)으로 변경
 */
function 예시_단일수정() {
  관리자_출석수정(
    '센트룸',              // 조원 이름
    '2025-10-15',       // 날짜
    'O',                // 상태: 'O'(출석), 'OFF'(오프), 'LONG_OFF'(장기오프), 'X'(결석)
    '시스템 오류 수정',   // 사유
    true                // 덮어쓰기: true
  );
}

/**
 * 예시 2: 결석을 출석으로 변경
 */
function 예시_결석을출석으로() {
  관리자_출석수정('길', '2025-10-16', 'O', '인증 누락, 관리자 확인 후 출석 처리');
}

/**
 * 예시 3: 출석을 오프로 변경
 */
function 예시_출석을오프로() {
  관리자_출석수정('what', '2025-10-17', 'OFF', '사후 오프 신청 승인');
}

/**
 * 예시 4: 일괄 수정
 * 여러 조원의 기록을 한 번에 수정
 */
function 예시_일괄수정() {
  const 수정목록 = [
    {
      name: '센트룸',
      date: '2025-10-15',
      status: 'O',
      reason: '정전으로 인한 업로드 지연'
    },
    {
      name: '길',
      date: '2025-10-15',
      status: 'O',
      reason: '정전으로 인한 업로드 지연'
    },
    {
      name: 'what',
      date: '2025-10-15',
      status: 'O',
      reason: '정전으로 인한 업로드 지연'
    }
  ];
  
  관리자_일괄수정(수정목록);
}

/**
 * 예시 5: 기록 삭제
 * 중복되거나 잘못된 기록 삭제
 */
function 예시_기록삭제() {
  관리자_기록삭제('센트룸', '2025-10-18');
}

/**
 * 예시 6: 특정 조원 기록 조회
 */
function 예시_기록조회() {
  // 전체 조회
  관리자_기록조회('센트룸');
  
  // 특정 월만 조회
  // 관리자_기록조회('센트룸', '2025-10');
}

/**
 * 예시 7: 장기오프로 변경
 */
function 예시_장기오프로변경() {
  관리자_출석수정(
    '녹동',
    '2025-10-20',
    'LONG_OFF',
    '해외 출장 (사후 신청)'
  );
}

/**
 * 🆕 폴더 ID 테스트 (여러 폴더 지원)
 * 특정 조원의 폴더에 접근 가능한지 확인
 */
function 폴더ID테스트() {
  const 테스트조원 = '센트룸';  // ← 테스트할 조원 이름 (CONFIG.MEMBERS에 있는 이름)
  
  Logger.log(`=== 폴더 ID 테스트: ${테스트조원} ===`);
  Logger.log('');
  
  const folderIdOrArray = CONFIG.MEMBERS[테스트조원];
  
  if (!folderIdOrArray) {
    Logger.log(`❌ CONFIG.MEMBERS에 "${테스트조원}" 조원이 없습니다.`);
    Logger.log('사용 가능한 조원: ' + Object.keys(CONFIG.MEMBERS).join(', '));
    return;
  }
  
  // 폴더 ID를 배열로 정규화
  const folderIds = Array.isArray(folderIdOrArray) ? folderIdOrArray : [folderIdOrArray];
  
  Logger.log(`📁 총 ${folderIds.length}개 폴더 테스트`);
  Logger.log('');
  
  // 각 폴더 테스트
  folderIds.forEach((folderId, index) => {
    Logger.log(`--- 폴더 ${index + 1}/${folderIds.length} ---`);
    Logger.log(`폴더 ID: ${folderId}`);
    
    try {
      const folder = DriveApp.getFolderById(folderId);
      Logger.log(`✅ 폴더 접근 성공: "${folder.getName()}"`);
      
      // 하위 폴더 샘플 확인 (최대 5개)
      const subfolders = folder.getFolders();
      let count = 0;
      const sampleFolders = [];
      
      while (subfolders.hasNext() && count < 5) {
        const subfolder = subfolders.next();
        sampleFolders.push(subfolder.getName());
        count++;
      }
      
      if (sampleFolders.length > 0) {
        Logger.log(`📂 하위 폴더 샘플 (최대 5개):`);
        sampleFolders.forEach(name => {
          Logger.log(`  - ${name}`);
        });
      } else {
        Logger.log(`⚠️ 하위 폴더가 없습니다.`);
      }
      
      // 파일 샘플 확인 (최대 3개)
      const files = folder.getFiles();
      count = 0;
      const sampleFiles = [];
      
      while (files.hasNext() && count < 3) {
        const file = files.next();
        sampleFiles.push(file.getName());
        count++;
      }
      
      if (sampleFiles.length > 0) {
        Logger.log(`📄 파일 샘플 (최대 3개):`);
        sampleFiles.forEach(name => {
          Logger.log(`  - ${name}`);
        });
      }
      
    } catch (e) {
      Logger.log(`❌ 폴더 접근 실패: ${e.message}`);
      Logger.log(`원인:`);
      Logger.log(`  1. 폴더 ID가 잘못됨`);
      Logger.log(`  2. 공유 권한이 없음`);
      Logger.log(`  3. 폴더가 삭제됨`);
      Logger.log(`해결: Google Drive에서 폴더 URL 다시 확인`);
    }
    
    Logger.log('');
  });
  
  Logger.log('=== 테스트 완료 ===');
}

// ==================== Web App 배포 ====================

/**
 * 통합 doGet 함수 - 모든 웹앱 기능 처리
 * - date 파라미터: 다이제스트 HTML 서빙
 * - month + type 파라미터: 출석/주간 JSON 반환
 * - action=getDigest: 다이제스트 JSON 반환
 */
function doGet(e) {
  try {
    const params = e.parameter;

    // 1. 다이제스트 HTML 서빙 (date 파라미터)
    if (params.date) {
      return 다이제스트HTML서빙(params.date);
    }

    // 2. 다이제스트 JSON API (action=getDigest)
    if (params.action === 'getDigest') {
      const date = params.date || Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
      const digest = 저장된다이제스트불러오기(date);
      return ContentService
        .createTextOutput(JSON.stringify(digest))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 3. 출석/주간 통계 JSON (month 파라미터)
    const month = params.month || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
    const type = params.type || 'attendance';

    Logger.log('Web App 요청 받음. 월:', month, '타입:', type);

    const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);

    // 타입에 따라 다른 파일명 사용
    let fileName;
    if (type === 'weekly') {
      fileName = `weekly_summary_${month}.json`;
    } else {
      fileName = `attendance_summary_${month}.json`;
    }

    const files = folder.getFilesByName(fileName);

    if (!files.hasNext()) {
      Logger.log('JSON 파일 없음:', fileName);

      // 주간 통계가 없을 때는 에러 반환
      if (type === 'weekly') {
        return ContentService
          .createTextOutput(JSON.stringify({
            error: true,
            message: '주간 통계 파일이 없습니다. 이번달주간집계() 함수를 실행해주세요.'
          }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      // 일반 출석 데이터가 없을 때는 빈 데이터 반환
      const emptyData = {};
      Object.keys(CONFIG.MEMBERS).forEach(name => {
        emptyData[name] = {
          출석: 0,
          결석: 0,
          오프: 0,
          장기오프: 0,
          경고: false,
          벌칙: false,
          기록: {},
          주간통계: {}
        };
      });

      return ContentService
        .createTextOutput(JSON.stringify(emptyData))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const file = files.next();
    const content = file.getBlob().getDataAsString();

    Logger.log('JSON 파일 로드 성공:', fileName);

    return ContentService
      .createTextOutput(content)
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log('Web App 오류:', error);

    return ContentService
      .createTextOutput(JSON.stringify({
        error: true,
        message: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 다이제스트 HTML 서빙 (다이제스트 웹앱 기능)
 */
function 다이제스트HTML서빙(dateStr) {
  try {
    const htmlContent = 다이제스트HTML가져오기(dateStr);

    if (!htmlContent) {
      return HtmlService.createHtmlOutput(`
        <!DOCTYPE html>
        <html lang="ko">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>다이제스트 없음</title>
          <style>
            body {
              font-family: -apple-system, BlinkMacSystemFont, "Apple SD Gothic Neo", "Malgun Gothic", sans-serif;
              display: flex;
              justify-content: center;
              align-items: center;
              height: 100vh;
              margin: 0;
              background: #f8f9fa;
            }
            .message {
              text-align: center;
              padding: 40px;
              background: white;
              border-radius: 12px;
              box-shadow: 0 2px 12px rgba(0,0,0,0.1);
            }
            h1 { color: #e74c3c; margin-bottom: 10px; }
            p { color: #7f8c8d; }
          </style>
        </head>
        <body>
          <div class="message">
            <h1>❌ 다이제스트를 찾을 수 없습니다</h1>
            <p>${dateStr} 날짜의 다이제스트가 없습니다.</p>
            <p style="font-size: 14px; margin-top: 20px;">
              URL 형식: <code>...exec?date=2025-11-21</code>
            </p>
          </div>
        </body>
        </html>
      `);
    }

    return HtmlService.createHtmlOutput(htmlContent)
      .setTitle(`📚 ${dateStr} 스터디 다이제스트`)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    Logger.log(`다이제스트 HTML 서빙 오류: ${error.message}`);

    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html lang="ko">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>오류</title>
        <style>
          body {
            font-family: -apple-system, BlinkMacSystemFont, "Apple SD Gothic Neo", "Malgun Gothic", sans-serif;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            margin: 0;
            background: #f8f9fa;
          }
          .error {
            text-align: center;
            padding: 40px;
            background: white;
            border-radius: 12px;
            box-shadow: 0 2px 12px rgba(0,0,0,0.1);
          }
          h1 { color: #e74c3c; margin-bottom: 10px; }
          p { color: #7f8c8d; }
          code {
            display: block;
            background: #f4f4f4;
            padding: 10px;
            border-radius: 4px;
            margin-top: 10px;
            font-size: 12px;
            color: #e74c3c;
          }
        </style>
      </head>
      <body>
        <div class="error">
          <h1>⚠️ 오류 발생</h1>
          <p>다이제스트를 불러오는 중 오류가 발생했습니다.</p>
          <code>${error.message}</code>
        </div>
      </body>
      </html>
    `);
  }
}

/**
 * 저장된 HTML 다이제스트 파일 가져오기
 */
function 다이제스트HTML가져오기(dateStr) {
  try {
    const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);
    const htmlFileName = `digest-${dateStr}.html`;

    const files = folder.getFilesByName(htmlFileName);

    if (!files.hasNext()) {
      Logger.log(`HTML 파일 없음: ${htmlFileName}`);
      return null;
    }

    const file = files.next();
    const htmlContent = file.getBlob().getDataAsString('UTF-8');

    return htmlContent;

  } catch (error) {
    Logger.log(`HTML 파일 읽기 실패: ${error.message}`);
    throw error;
  }
}

/**
 * 저장된 다이제스트 JSON 불러오기
 */
function 저장된다이제스트불러오기(dateStr) {
  const fileName = `digest-${dateStr}.json`;
  const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);

  try {
    const files = folder.getFilesByName(fileName);
    if (files.hasNext()) {
      const file = files.next();
      const content = file.getBlob().getDataAsString('UTF-8');
      return JSON.parse(content);
    }
    return { error: true, message: '다이제스트 파일을 찾을 수 없습니다.' };
  } catch (error) {
    Logger.log(`다이제스트 JSON 로드 오류: ${error.message}`);
    return { error: true, message: error.message };
  }
}

/**
 * 🆕 특정 날짜만 최종 스캔하는 함수
 * 마감시간체크() 전에 해당 날짜를 한 번 더 체크
 */
function 최종스캔_특정날짜(targetDateStr) {
  memberLoop:
  for (const [memberName, folderIdOrArray] of Object.entries(CONFIG.MEMBERS)) {
    const folderIds = Array.isArray(folderIdOrArray) ? folderIdOrArray : [folderIdOrArray];
    
    for (const folderId of folderIds) {
      try {
        const mainFolder = DriveApp.getFolderById(folderId);
        const subfolders = mainFolder.getFolders();
        
        while (subfolders.hasNext()) {
          const folder = subfolders.next();
          const folderName = folder.getName().trim();
          const dateInfo = 날짜추출(folderName);
          
          if (dateInfo && dateInfo.dateStr === targetDateStr) {
            // 장기오프 체크
            const longOffInfo = 장기오프확인(memberName, targetDateStr);
            
            if (longOffInfo.isLongOff) {
              Logger.log(`  🏝️ ${memberName} - 장기오프 (${longOffInfo.reason})`);
              출석기록추가(memberName, targetDateStr, [], CONFIG.LONG_OFF_STATUS, longOffInfo.reason);
              continue memberLoop;
            }
            
            // OFF.md 체크
            const isOff = OFF파일확인(folder);
            
            if (isOff) {
              Logger.log(`  🏖️ ${memberName} - 오프`);
              출석기록추가(memberName, targetDateStr, [], 'OFF');
              continue memberLoop;
            }
            
            // 일반 출석 체크
            const files = 파일목록및링크생성(folder);
            
            if (files.length > 0) {
              Logger.log(`  ✓ ${memberName} - 출석 (${files.length}개 파일)`);
              // 🔧 수정: dateStr → targetDateStr
              출석기록추가(memberName, targetDateStr, files, 'O', '', folder.getId());
              continue memberLoop;
            }
          }
        }
      } catch (error) {
        Logger.log(`  ❌ ${memberName} 스캔 오류: ${error.message}`);
      }
    }
  }
}

// ==================== 🆕 관리자수정 함수 (기존 코드 맨 끝에 추가) ====================

/**
 * 🆕 관리자수정 시트 처리
 */
function 관리자수정처리() {
  Logger.log('');
  Logger.log('=== 관리자수정 처리 시작 ===');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adminSheet = ss.getSheetByName(CONFIG.ADMIN_SHEET);
  
  if (!adminSheet) {
    Logger.log(`⚠️ "${CONFIG.ADMIN_SHEET}" 시트가 없습니다. 건너뜀.`);
    return;
  }
  
  const data = adminSheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    Logger.log('처리할 항목이 없습니다.');
    return;
  }
  
  let processedCount = 0;
  const now = new Date();
  
  // 첫 행(헤더) 제외하고 처리
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[CONFIG.ADMIN_COLUMNS.NAME];
    const dateValue = row[CONFIG.ADMIN_COLUMNS.DATE];
    const status = row[CONFIG.ADMIN_COLUMNS.STATUS];
    const reason = row[CONFIG.ADMIN_COLUMNS.REASON] || '';
    const processed = row[CONFIG.ADMIN_COLUMNS.PROCESSED];
    
    // 이미 처리된 항목은 건너뛰기
    if (processed === '완료' || processed === 'O' || processed === '✅') {
      continue;
    }
    
    // 필수 필드 검증
    if (!name || !dateValue || !status) {
      Logger.log(`  ⚠️ ${i + 1}행: 필수 정보 누락 (이름: ${name}, 날짜: ${dateValue}, 상태: ${status})`);
      continue;
    }
    
    // 날짜 포맷 변환
    let dateStr;
    try {
      if (dateValue instanceof Date) {
        dateStr = Utilities.formatDate(dateValue, 'Asia/Seoul', 'yyyy-MM-dd');
      } else {
        dateStr = String(dateValue).trim();
        if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
          Logger.log(`  ❌ ${i + 1}행: 날짜 형식 오류 (${dateStr}). YYYY-MM-DD 형식 사용 필요`);
          continue;
        }
      }
    } catch (e) {
      Logger.log(`  ❌ ${i + 1}행: 날짜 변환 실패 (${dateValue})`);
      continue;
    }
    
    // 조원 이름 검증
    if (!CONFIG.MEMBERS[name]) {
      Logger.log(`  ❌ ${i + 1}행: 알 수 없는 조원 (${name})`);
      continue;
    }
    
    // 상태 값 검증 및 정규화
    let normalizedStatus = status.toString().toUpperCase().trim();
    if (normalizedStatus === 'O' || normalizedStatus === '출석') {
      normalizedStatus = 'O';
    } else if (normalizedStatus === 'OFF' || normalizedStatus === '오프') {
      normalizedStatus = 'OFF';
    } else if (normalizedStatus === 'X' || normalizedStatus === '결석') {
      normalizedStatus = 'X';
    } else if (normalizedStatus === 'LONG_OFF' || normalizedStatus === '장기오프') {
      normalizedStatus = 'LONG_OFF';
    } else {
      Logger.log(`  ❌ ${i + 1}행: 알 수 없는 상태 (${status}). O/OFF/X/LONG_OFF 중 하나 사용`);
      continue;
    }
    
    // 출석기록 추가/업데이트
    try {
      Logger.log(`  🔧 ${name} - ${dateStr} → ${normalizedStatus}${reason ? ' (' + reason + ')' : ''}`);
      출석기록추가(name, dateStr, [], normalizedStatus, reason);
      
      // 처리완료 표시
      const rowIndex = i + 1;
      adminSheet.getRange(rowIndex, CONFIG.ADMIN_COLUMNS.PROCESSED + 1).setValue('완료');
      adminSheet.getRange(rowIndex, CONFIG.ADMIN_COLUMNS.PROCESSED_TIME + 1).setValue(now);
      
      processedCount++;
    } catch (e) {
      Logger.log(`  ❌ ${i + 1}행: 처리 중 오류 - ${e.message}`);
    }
  }
  
  Logger.log(`✅ 관리자수정 처리 완료: ${processedCount}건`);
  Logger.log('');
  
  return processedCount;
}

/**
 * 🆕 관리자수정 존재 확인
 */
function 관리자수정존재확인(memberName, dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adminSheet = ss.getSheetByName(CONFIG.ADMIN_SHEET);
  
  if (!adminSheet) {
    return false;
  }
  
  const data = adminSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[CONFIG.ADMIN_COLUMNS.NAME];
    const dateValue = row[CONFIG.ADMIN_COLUMNS.DATE];
    const processed = row[CONFIG.ADMIN_COLUMNS.PROCESSED];
    
    if (processed !== '완료' && processed !== 'O' && processed !== '✅') {
      continue;
    }
    
    if (name !== memberName) {
      continue;
    }
    
    let checkDateStr;
    try {
      if (dateValue instanceof Date) {
        checkDateStr = Utilities.formatDate(dateValue, 'Asia/Seoul', 'yyyy-MM-dd');
      } else {
        checkDateStr = String(dateValue).trim();
      }
      
      if (checkDateStr === dateStr) {
        return true;
      }
    } catch (e) {
      continue;
    }
  }
  
  return false;
}

/**
 * 🆕 수동으로 관리자수정만 처리 (테스트용)
 */
function 관리자수정만_처리() {
  관리자수정처리();
  JSON파일생성();
  Logger.log('✅ 관리자수정 처리 및 JSON 생성 완료!');
}




// ==================== 🆕 월별결산 기능 ====================

/**
 * 🆕 월별결산 생성 (매월 1일 실행)
 * 전월 데이터를 "월별결산" 시트에 저장
 */
function 월별결산생성() {
  Logger.log('');
  Logger.log('=== 월별결산 생성 시작 ===');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let summarySheet = ss.getSheetByName(CONFIG.MONTHLY_SUMMARY_SHEET);
  
  // 시트가 없으면 생성
  if (!summarySheet) {
    Logger.log('월별결산 시트 생성...');
    summarySheet = ss.insertSheet(CONFIG.MONTHLY_SUMMARY_SHEET);
    
    // 헤더 설정
    const headers = ['연월', '조원명', '출석', '오프', '장기오프', '결석', '출석률(%)', '상태', '비고'];
    summarySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    summarySheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4CAF50').setFontColor('white');
    summarySheet.setFrozenRows(1);
    
    // 열 너비 설정
    summarySheet.setColumnWidth(1, 100);  // 연월
    summarySheet.setColumnWidth(2, 150);  // 조원명
    summarySheet.setColumnWidths(3, 4, 80);  // 출석~결석
    summarySheet.setColumnWidth(7, 100);  // 출석률
    summarySheet.setColumnWidth(8, 120);  // 상태
    summarySheet.setColumnWidth(9, 200);  // 비고
  }
  
  // 전월 데이터 계산
  const now = new Date();
  const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const targetYear = lastMonth.getFullYear();
  const targetMonth = lastMonth.getMonth(); // 0-based
  const yearMonth = `${targetYear}-${String(targetMonth + 1).padStart(2, '0')}`;
  
  // 해당 월의 총 일수 계산
  const daysInMonth = new Date(targetYear, targetMonth + 1, 0).getDate();
  
  Logger.log(`대상 연월: ${yearMonth} (총 ${daysInMonth}일)`);
  
  // 이미 해당 월 결산이 있는지 확인
  const existingData = summarySheet.getDataRange().getValues();
  let alreadyExists = false;
  for (let i = 1; i < existingData.length; i++) {
    if (existingData[i][0] === yearMonth) {
      alreadyExists = true;
      Logger.log(`⚠️ ${yearMonth} 결산이 이미 존재합니다. 업데이트합니다.`);
      break;
    }
  }
  
  // 전월 출석 데이터 집계
  const attendanceSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!attendanceSheet) {
    Logger.log('❌ 제출기록 시트가 없습니다.');
    return;
  }
  
  const data = attendanceSheet.getDataRange().getValues();
  
  // 🔧 중복 제거: 각 조원별로 날짜별 최신 상태만 저장
  const memberDateRecords = {};
  
  // 조원별 Map 초기화
  for (const memberName of Object.keys(CONFIG.MEMBERS)) {
    memberDateRecords[memberName] = new Map(); // key: 날짜, value: {status, timestamp, reason}
  }
  
  // 데이터 수집 (같은 날짜는 최신 타임스탬프만 유지)
  for (let i = 1; i < data.length; i++) {
    const [timestamp, name, dateStr, fileCount, links, folderLink, status, weekNum, reason] = data[i];
    
    if (!memberDateRecords[name]) continue;
    
    const dateStrFormatted = typeof dateStr === 'string' 
      ? dateStr 
      : Utilities.formatDate(new Date(dateStr), 'Asia/Seoul', 'yyyy-MM-dd');
    
    // 해당 월 데이터만 처리
    if (dateStrFormatted && dateStrFormatted.startsWith(yearMonth)) {
      const existing = memberDateRecords[name].get(dateStrFormatted);
      
      // 기존 기록이 없거나, 더 최신 기록이면 업데이트
      if (!existing || timestamp > existing.timestamp) {
        memberDateRecords[name].set(dateStrFormatted, {
          status: status,
          timestamp: timestamp,
          reason: reason || ''  // 🆕 사유 추가
        });
      }
    }
  }
  
  // 통계 계산
  const summaryData = [];
  
  for (const [memberName, dateMap] of Object.entries(memberDateRecords)) {
    let 출석 = 0;
    let 오프 = 0;
    let 장기오프 = 0;
    const 오프초과결석목록 = [];  // 🆕 오프 초과로 결석 전환된 날짜 목록
    
    // 날짜별 상태 카운트
    for (const [date, record] of dateMap.entries()) {
      if (record.status === 'O') {
        출석++;
      } else if (record.status === 'OFF') {
        오프++;
      } else if (record.status === 'LONG_OFF') {
        장기오프++;
      } else if (record.status === 'X' && record.reason && record.reason.includes('오프') && record.reason.includes('초과')) {
        // 🆕 오프 초과 결석 감지
        오프초과결석목록.push(date.split('-')[2] + '일');  // 날짜만 추출 (예: "01일")
      }
      // 'X'나 다른 상태는 결석으로 처리됨 (아래에서 계산)
    }
    
    // 🔧 정확한 결석 계산: 전체 일수 - (출석 + 오프 + 장기오프)
    const 결석 = daysInMonth - (출석 + 오프 + 장기오프);
    
    // 출석률 계산
    const 출석률 = daysInMonth > 0 ? ((출석 / daysInMonth) * 100).toFixed(1) : 0;
    
    // 상태 판정
    let 상태 = '정상';
    if (결석 >= 4) {
      상태 = '🚨 벌칙';
    } else if (결석 === 3) {
      상태 = '⚠️ 경고';
    } else {
      상태 = '✅ 정상';
    }
    
    let 비고 = `출석 ${출석}일 + 오프 ${오프}일 + 장기오프 ${장기오프}일 + 결석 ${결석}일 = 총 ${daysInMonth}일`;
    if (오프초과결석목록.length > 0) {
      비고 += ` | 🚨 오프 초과 결석: ${오프초과결석목록.join(', ')}`;
    }
    
    Logger.log(`${memberName}: 출석 ${출석}, 오프 ${오프}, 장기오프 ${장기오프}, 결석 ${결석} → ${상태}`);
    
    summaryData.push([
      yearMonth,
      memberName,
      출석,
      오프,
      장기오프,
      결석,
      출석률,
      상태,
      비고
    ]);
  }
  
  // 기존 데이터 삭제 (같은 연월)
  if (alreadyExists) {
    for (let i = existingData.length - 1; i >= 1; i--) {
      if (existingData[i][0] === yearMonth) {
        summarySheet.deleteRow(i + 1);
      }
    }
  }
  
  // 새 데이터 추가
  const lastRow = summarySheet.getLastRow();
  summarySheet.getRange(lastRow + 1, 1, summaryData.length, summaryData[0].length).setValues(summaryData);
  
  // 조건부 서식 적용
  for (let i = 0; i < summaryData.length; i++) {
    const rowNum = lastRow + 1 + i;
    const 출석률 = parseFloat(summaryData[i][6]);
    
    if (출석률 >= 90) {
      summarySheet.getRange(rowNum, 7).setBackground('#e8f5e9'); // 초록
    } else if (출석률 >= 70) {
      summarySheet.getRange(rowNum, 7).setBackground('#fff9c4'); // 노랑
    } else {
      summarySheet.getRange(rowNum, 7).setBackground('#ffcdd2'); // 빨강
    }
    
    // 상태에 따른 색상
    const 상태 = summaryData[i][7];
    if (상태.includes('벌칙')) {
      summarySheet.getRange(rowNum, 8).setBackground('#ffcdd2');
    } else if (상태.includes('경고')) {
      summarySheet.getRange(rowNum, 8).setBackground('#ffe0b2');
    } else {
      summarySheet.getRange(rowNum, 8).setBackground('#e8f5e9');
    }
  }
  
  Logger.log(`✅ ${yearMonth} 월별결산 저장 완료: ${summaryData.length}명`);
  Logger.log('');
}

/**
 * 🆕 월별결산 수동 실행 (테스트용)
 */
function 월별결산_수동실행() {
  월별결산생성();
  Logger.log('✅ 월별결산 생성 완료!');
}

/**
 * 🆕 특정 월의 결산 생성 (수동 실행용)
 * @param {number} year - 연도 (예: 2025)
 * @param {number} month - 월 (1-12)
 */
function 특정월_결산생성(year, month) {
  Logger.log('');
  Logger.log('=== 특정 월 결산 생성 시작 ===');
  Logger.log(`대상: ${year}년 ${month}월`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let summarySheet = ss.getSheetByName(CONFIG.MONTHLY_SUMMARY_SHEET);
  
  // 시트가 없으면 생성
  if (!summarySheet) {
    Logger.log('월별결산 시트 생성...');
    summarySheet = ss.insertSheet(CONFIG.MONTHLY_SUMMARY_SHEET);
    
    // 헤더 설정
    const headers = ['연월', '조원명', '출석', '오프', '장기오프', '결석', '출석률(%)', '상태', '비고'];
    summarySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    summarySheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4CAF50').setFontColor('white');
    summarySheet.setFrozenRows(1);
    
    // 열 너비 설정
    summarySheet.setColumnWidth(1, 100);
    summarySheet.setColumnWidth(2, 150);
    summarySheet.setColumnWidths(3, 4, 80);
    summarySheet.setColumnWidth(7, 100);
    summarySheet.setColumnWidth(8, 120);
    summarySheet.setColumnWidth(9, 200);
  }
  
  const yearMonth = `${year}-${String(month).padStart(2, '0')}`;
  
  // 해당 월의 총 일수 계산
  const daysInMonth = new Date(year, month, 0).getDate();
  Logger.log(`${yearMonth} (총 ${daysInMonth}일)`);
  
  // 기존 데이터 확인
  const existingData = summarySheet.getDataRange().getValues();
  let alreadyExists = false;
  for (let i = 1; i < existingData.length; i++) {
    if (existingData[i][0] === yearMonth) {
      alreadyExists = true;
      Logger.log(`⚠️ ${yearMonth} 결산이 이미 존재합니다. 업데이트합니다.`);
      break;
    }
  }
  
  // 출석 데이터 집계
  const attendanceSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!attendanceSheet) {
    Logger.log('❌ 제출기록 시트가 없습니다.');
    return;
  }
  
  const data = attendanceSheet.getDataRange().getValues();
  
  // 🔧 중복 제거: 각 조원별로 날짜별 최신 상태만 저장
  const memberDateRecords = {};
  
  // 조원별 Map 초기화
  for (const memberName of Object.keys(CONFIG.MEMBERS)) {
    memberDateRecords[memberName] = new Map(); // key: 날짜, value: {status, timestamp, reason}
  }
  
  // 데이터 수집 (같은 날짜는 최신 타임스탬프만 유지)
  for (let i = 1; i < data.length; i++) {
    const [timestamp, name, dateStr, fileCount, links, folderLink, status, weekNum, reason] = data[i];
    
    if (!memberDateRecords[name]) continue;
    
    const dateStrFormatted = typeof dateStr === 'string' 
      ? dateStr 
      : Utilities.formatDate(new Date(dateStr), 'Asia/Seoul', 'yyyy-MM-dd');
    
    // 해당 월 데이터만 처리
    if (dateStrFormatted && dateStrFormatted.startsWith(yearMonth)) {
      const existing = memberDateRecords[name].get(dateStrFormatted);
      
      // 기존 기록이 없거나, 더 최신 기록이면 업데이트
      if (!existing || timestamp > existing.timestamp) {
        memberDateRecords[name].set(dateStrFormatted, {
          status: status,
          timestamp: timestamp,
          reason: reason || ''  // 🆕 사유 추가
        });
      }
    }
  }
  
  // 통계 계산
  const summaryData = [];
  
  for (const [memberName, dateMap] of Object.entries(memberDateRecords)) {
    let 출석 = 0;
    let 오프 = 0;
    let 장기오프 = 0;
    const 오프초과결석목록 = [];  // 🆕 오프 초과로 결석 전환된 날짜 목록
    
    // 날짜별 상태 카운트
    for (const [date, record] of dateMap.entries()) {
      if (record.status === 'O') {
        출석++;
      } else if (record.status === 'OFF') {
        오프++;
      } else if (record.status === 'LONG_OFF') {
        장기오프++;
      } else if (record.status === 'X' && record.reason && record.reason.includes('오프') && record.reason.includes('초과')) {
        // 🆕 오프 초과 결석 감지
        오프초과결석목록.push(date.split('-')[2] + '일');  // 날짜만 추출 (예: "01일")
      }
      // 'X'나 다른 상태는 결석으로 처리됨 (아래에서 계산)
    }
    
    // 🔧 정확한 결석 계산: 전체 일수 - (출석 + 오프 + 장기오프)
    const 결석 = daysInMonth - (출석 + 오프 + 장기오프);
    
    // 출석률 계산
    const 출석률 = daysInMonth > 0 ? ((출석 / daysInMonth) * 100).toFixed(1) : 0;
    
    // 상태 판정
    let 상태 = '정상';
    if (결석 >= 4) {
      상태 = '🚨 벌칙';
    } else if (결석 === 3) {
      상태 = '⚠️ 경고';
    } else {
      상태 = '✅ 정상';
    }
    
    let 비고 = `출석 ${출석}일 + 오프 ${오프}일 + 장기오프 ${장기오프}일 + 결석 ${결석}일 = 총 ${daysInMonth}일`;
    if (오프초과결석목록.length > 0) {
      비고 += ` | 🚨 오프 초과 결석: ${오프초과결석목록.join(', ')}`;
    }
    
    Logger.log(`${memberName}: 출석 ${출석}, 오프 ${오프}, 장기오프 ${장기오프}, 결석 ${결석} → ${상태}`);
    
    summaryData.push([
      yearMonth,
      memberName,
      출석,
      오프,
      장기오프,
      결석,
      출석률,
      상태,
      비고
    ]);
  }
  
  // 기존 데이터 삭제 (같은 연월)
  if (alreadyExists) {
    for (let i = existingData.length - 1; i >= 1; i--) {
      if (existingData[i][0] === yearMonth) {
        summarySheet.deleteRow(i + 1);
      }
    }
  }
  
  // 새 데이터 추가
  const lastRow = summarySheet.getLastRow();
  summarySheet.getRange(lastRow + 1, 1, summaryData.length, summaryData[0].length).setValues(summaryData);
  
  // 조건부 서식 적용
  for (let i = 0; i < summaryData.length; i++) {
    const rowNum = lastRow + 1 + i;
    const 출석률 = parseFloat(summaryData[i][6]);
    
    if (출석률 >= 90) {
      summarySheet.getRange(rowNum, 7).setBackground('#e8f5e9');
    } else if (출석률 >= 70) {
      summarySheet.getRange(rowNum, 7).setBackground('#fff9c4');
    } else {
      summarySheet.getRange(rowNum, 7).setBackground('#ffcdd2');
    }
    
    const 상태 = summaryData[i][7];
    if (상태.includes('벌칙')) {
      summarySheet.getRange(rowNum, 8).setBackground('#ffcdd2');
    } else if (상태.includes('경고')) {
      summarySheet.getRange(rowNum, 8).setBackground('#ffe0b2');
    } else {
      summarySheet.getRange(rowNum, 8).setBackground('#e8f5e9');
    }
  }
  
  Logger.log(`✅ ${yearMonth} 월별결산 저장 완료: ${summaryData.length}명`);
  Logger.log('');
}

/**
 * 🆕 10월 결산 생성 (예시)
 */
function 결산_10월생성() {
  특정월_결산생성(2025, 10);
  Logger.log('✅ 2025년 10월 결산 생성 완료!');
}

function 녹동_폴더링크_복구() {
  Logger.log('=== 녹동 폴더 링크 복구 시작 ===');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  const 녹동폴더ID = CONFIG.MEMBERS['녹동'];
  const mainFolder = DriveApp.getFolderById(녹동폴더ID);
  
  let 수정개수 = 0;
  
  for (let i = 1; i < data.length; i++) {
    const [timestamp, name, dateStr, fileCount, links, folderLink, status] = data[i];
    
    // 녹동의 출석 기록이고 폴더 링크가 비어있거나 불완전한 경우
    if (name === '녹동' && status === 'O' && (!folderLink || folderLink.length < 50)) {
      
      const dateFormatted = typeof dateStr === 'string' ? 
        dateStr : 
        Utilities.formatDate(new Date(dateStr), 'Asia/Seoul', 'yyyy-MM-dd');
      
      Logger.log(`처리 중: ${dateFormatted}`);
      
      // 날짜 폴더 찾기
      const dateFolder = 오늘날짜폴더찾기(mainFolder, dateFormatted);
      
      if (dateFolder) {
        const newFolderLink = `https://drive.google.com/drive/folders/${dateFolder.getId()}`;
        
        // 시트의 폴더 링크 열(F열, 6번째) 업데이트
        sheet.getRange(i + 1, 6).setValue(newFolderLink);
        
        Logger.log(`  ✅ 수정 완료: ${newFolderLink}`);
        수정개수++;
      } else {
        Logger.log(`  ⚠️ ${dateFormatted} 폴더를 찾을 수 없음`);
      }
    }
  }
  
  Logger.log('');
  Logger.log(`=== 복구 완료: ${수정개수}개 기록 수정됨 ===`);
  
  // JSON 재생성
  if (수정개수 > 0) {
    Logger.log('JSON 파일 재생성 중...');
    JSON파일생성();
    Logger.log('✅ JSON 파일 재생성 완료!');
  }
}

// ==================== 🎯 원클릭 장기오프 시스템 완전 설치 ====================

/**
 * 🎯 장기오프 시스템 완전 설치 (원클릭)
 * - 기존 신청 즉시 반영
 * - 폼 제출 트리거 자동 설정
 * - 전체 시스템 검증
 */
function 장기오프시스템_완전설치() {
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('🎯 장기오프 시스템 완전 설치 시작');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('');
  
  // Step 1: 기존 신청 즉시 반영
  Logger.log('📋 Step 1: 기존 장기오프 신청 처리 중...');
  Logger.log('');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const longOffSheet = ss.getSheetByName(CONFIG.LONG_OFF_SHEET);
  
  if (!longOffSheet) {
    Logger.log('❌ 장기오프신청 시트가 없습니다.');
    return;
  }
  
  const data = longOffSheet.getDataRange().getValues();
  let processedCount = 0;
  
  // 첫 행(헤더) 제외하고 처리
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[CONFIG.FORM_COLUMNS.NAME];
    const startDateValue = row[CONFIG.FORM_COLUMNS.START_DATE];
    const endDateValue = row[CONFIG.FORM_COLUMNS.END_DATE];
    const reason = row[CONFIG.FORM_COLUMNS.REASON];
    const approved = row[CONFIG.FORM_COLUMNS.APPROVED];
    
    // 필수 정보 확인
    if (!name || !startDateValue || !endDateValue) {
      Logger.log(`  ⏭️ ${i + 1}행: 정보 누락, 건너뜀`);
      continue;
    }
    
    // 승인 여부 확인 (자동 승인 모드가 아닐 때)
    if (!CONFIG.LONG_OFF_AUTO_APPROVE && approved !== 'O' && approved !== 'o') {
      Logger.log(`  ⏭️ ${i + 1}행: 미승인, 건너뜀`);
      continue;
    }
    
    // 날짜 파싱
    let startDate, endDate;
    
    try {
      if (startDateValue instanceof Date) {
        startDate = startDateValue;
      } else {
        startDate = new Date(startDateValue);
      }
      
      if (endDateValue instanceof Date) {
        endDate = endDateValue;
      } else {
        endDate = new Date(endDateValue);
      }
      
      startDate.setHours(0, 0, 0, 0);
      endDate.setHours(0, 0, 0, 0);
      
    } catch (e) {
      Logger.log(`  ❌ ${i + 1}행: 날짜 파싱 오류`);
      continue;
    }
    
    Logger.log(`  📝 처리 중: ${name} (${Utilities.formatDate(startDate, 'Asia/Seoul', 'MM/dd')} ~ ${Utilities.formatDate(endDate, 'Asia/Seoul', 'MM/dd')})`);
    
    // 해당 기간의 모든 날짜에 장기오프 기록 추가
    let daysProcessed = 0;
    const currentDate = new Date(startDate);
    
    while (currentDate <= endDate) {
      const dateStr = Utilities.formatDate(currentDate, 'Asia/Seoul', 'yyyy-MM-dd');
      출석기록추가(name, dateStr, [], CONFIG.LONG_OFF_STATUS, reason || '장기오프');
      daysProcessed++;
      currentDate.setDate(currentDate.getDate() + 1);
    }
    
    Logger.log(`     ✅ ${daysProcessed}일 처리 완료`);
    processedCount++;
  }
  
  Logger.log('');
  Logger.log(`✅ Step 1 완료: ${processedCount}건의 장기오프 신청 처리됨`);
  Logger.log('');
  
  // Step 2: JSON 재생성
  Logger.log('📄 Step 2: JSON 파일 생성 중...');
  JSON파일생성();
  Logger.log('✅ Step 2 완료: JSON 파일 생성됨');
  Logger.log('');
  
  // Step 3: 폼 제출 트리거 설정
  Logger.log('⚙️ Step 3: 폼 제출 트리거 설정 중...');
  
  // 기존 트리거 확인
  const triggers = ScriptApp.getProjectTriggers();
  let hasFormTrigger = false;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onFormSubmit_장기오프처리') {
      hasFormTrigger = true;
    }
  });
  
  if (hasFormTrigger) {
    Logger.log('✅ Step 3 완료: 폼 제출 트리거 이미 설정되어 있음');
  } else {
    Logger.log('⚠️ 폼 제출 트리거가 없습니다.');
    Logger.log('');
    Logger.log('💡 수동 설정 방법:');
    Logger.log('   1. Apps Script 왼쪽 시계 아이콘(트리거) 클릭');
    Logger.log('   2. "트리거 추가" 클릭');
    Logger.log('   3. 함수: onFormSubmit_장기오프처리');
    Logger.log('   4. 이벤트 소스: 스프레드시트에서');
    Logger.log('   5. 이벤트 유형: 양식 제출 시');
    Logger.log('   6. 저장');
  }
  
  Logger.log('');
  
  // Step 4: 검증
  Logger.log('🔍 Step 4: 시스템 검증 중...');
  Logger.log('');
  
  // 출석표 확인
  const attendanceSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (attendanceSheet) {
    const attendanceData = attendanceSheet.getDataRange().getValues();
    let longOffCount = 0;
    
    for (let i = 1; i < attendanceData.length; i++) {
      if (attendanceData[i][6] === CONFIG.LONG_OFF_STATUS) {
        longOffCount++;
      }
    }
    
    Logger.log(`✅ 출석표에서 ${longOffCount}개의 장기오프 기록 발견`);
  }
  
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('🎉 장기오프 시스템 완전 설치 완료!');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('');
  Logger.log('📊 설치 결과:');
  Logger.log(`   - 기존 신청 처리: ${processedCount}건`);
  Logger.log(`   - 폼 제출 트리거: ${hasFormTrigger ? '설정됨' : '수동 설정 필요'}`);
  Logger.log('');
  Logger.log('💡 다음 단계:');
  if (!hasFormTrigger) {
    Logger.log('   1. 위의 "수동 설정 방법"에 따라 트리거 설정');
    Logger.log('   2. 출석표 시트에서 Magnus의 10/18-19 장기오프 확인');
  } else {
    Logger.log('   1. 출석표 시트에서 Magnus의 10/18-19 장기오프 확인');
    Logger.log('   2. 이제부터 폼 제출 시 자동 반영됩니다!');
  }
  Logger.log('');
}

// ==================== 🆕 폼 제출 즉시 처리 함수 ====================

/**
 * 🆕 구글 폼 제출 시 자동 실행
 * 장기오프 신청 즉시 출석표 반영
 */
function onFormSubmit_장기오프처리(e) {
  try {
    Logger.log('=== 폼 제출 감지: 장기오프 즉시 처리 ===');
    
    const row = e.range.getRow();
    const sheet = e.range.getSheet();
    
    // 장기오프신청 시트가 맞는지 확인
    if (sheet.getName() !== CONFIG.LONG_OFF_SHEET) {
      return;
    }
    
    // 제출된 데이터 읽기
    const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const name = data[CONFIG.FORM_COLUMNS.NAME];
    const startDateValue = data[CONFIG.FORM_COLUMNS.START_DATE];
    const endDateValue = data[CONFIG.FORM_COLUMNS.END_DATE];
    const reason = data[CONFIG.FORM_COLUMNS.REASON];
    
    Logger.log(`신청자: ${name}`);
    Logger.log(`기간: ${startDateValue} ~ ${endDateValue}`);
    
    // 유효성 검사
    if (!name || !startDateValue || !endDateValue || !CONFIG.MEMBERS[name]) {
      Logger.log('❌ 유효하지 않은 신청');
      return;
    }
    
    // 날짜 파싱
    let startDate = startDateValue instanceof Date ? startDateValue : new Date(startDateValue);
    let endDate = endDateValue instanceof Date ? endDateValue : new Date(endDateValue);
    
    startDate.setHours(0, 0, 0, 0);
    endDate.setHours(0, 0, 0, 0);
    
    // 자동 승인
    if (CONFIG.LONG_OFF_AUTO_APPROVE) {
      sheet.getRange(row, CONFIG.FORM_COLUMNS.APPROVED + 1).setValue('O');
    }
    
    // 해당 기간 모든 날짜에 장기오프 기록
    let daysProcessed = 0;
    const currentDate = new Date(startDate);
    
    while (currentDate <= endDate) {
      const dateStr = Utilities.formatDate(currentDate, 'Asia/Seoul', 'yyyy-MM-dd');
      출석기록추가(name, dateStr, [], CONFIG.LONG_OFF_STATUS, reason || '장기오프');
      daysProcessed++;
      currentDate.setDate(currentDate.getDate() + 1);
    }
    
    Logger.log(`✅ ${name}의 ${daysProcessed}일 장기오프 처리 완료`);
    
    // JSON 재생성
    JSON파일생성();
    
  } catch (error) {
    Logger.log('❌ 오류: ' + error.toString());
  }
}
