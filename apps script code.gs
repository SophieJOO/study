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
    '오늘의너굴이': '1O4JHt954kCa1nPicUitQCbnyDIO2r1Zt',
    'Dann': '1mMoVApl7GN3EUYi9oPi7Nfo_2hYDb9Dw',
    '보노보노': '1_Mqn79Y1Qp79DWBxcbP-SGVUGjJA3PGw',
    'Magnus': ['1eHjsJ8bnWcK__8EXvukqixzh4wb8CncR', '1e8HUMzD0zW0BG2rkuB3kXoGtK2fw2fhG', '14UsgZTtXpuiBv2jYItigwQGiasIWof5s'],
    'RMSK전사': '1EUT7w_-7Cbu5216j2d6rkiq7DQzZ3p_s',
    '바른몸': '1i3lxqexxksy-y_Uft8e1XlUJAmTPmrTw'
  },
  
  // 시트 이름
  SHEET_NAME: '제출기록',
  ATTENDANCE_SHEET: '출석표',
  LONG_OFF_SHEET: '장기오프신청',
  ADMIN_SHEET: '관리자수정',  // 🆕 추가
  MONTHLY_SUMMARY_SHEET: '월별결산',  // 🆕 월별결산 시트
  PENALTY_SHEET: '벌칙관리',  // 🆕 벌칙 관리 시트
  DIGEST_SHEET: '다이제스트',  // 🆕 다이제스트 시트 (드라이브 대신 시트 사용)
  
  // JSON 파일 출력 폴더 ID
  JSON_FOLDER_ID: '1el9NDYDGfWlUEkBzI1GT_1TULLoBnSsQ',

  // 마감시간 설정
  DEADLINE_HOUR: 3,
  
  // 장기오프 설정
  LONG_OFF_STATUS: 'LONG_OFF',
  LONG_OFF_AUTO_APPROVE: false,
  
  // 구글 폼 응답 시트 열 구조
  FORM_COLUMNS: {
    TIMESTAMP: 0,
    NAME: 1,
    START_DATE: 2,
    END_DATE: 3,
    REASON: 4,
    APPROVED: 5
  },
  
  // 🆕 관리자수정 시트 열 구조 (파일링크는 맨 끝)
  ADMIN_COLUMNS: {
    NAME: 0,
    DATE: 1,
    STATUS: 2,
    REASON: 3,
    PROCESSED: 4,
    PROCESSED_TIME: 5,
    FILE_LINK: 6       // 🆕 파일링크 (선택)
  },
  
  // 스캔 설정
  SCAN_ALL_MONTHS: false,
  MAX_FOLDERS_TO_SCAN: 100,

  // 웹앱 URL (다이제스트 링크용)
  WEB_APP_URL: 'https://script.google.com/macros/s/AKfycbwvp_YIPlshnQm-HJISz6AvC6g0Mf5oFMLi0uMJ7DypVOqCGmJhDFmDVQ-60QN6Df3-gg/exec'
};

// ==================== 📢 Slack 알림 ====================

/**
 * Slack Webhook URL 가져오기 (Script Properties에서)
 * 보안을 위해 코드에 직접 저장하지 않고 Script Properties 사용
 */
function getSlackWebhookUrl() {
  return PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
}

/**
 * Slack Webhook URL 설정 (최초 1회 실행 필요)
 * Apps Script에서 이 함수를 실행하고 프롬프트에 Webhook URL 입력
 */
function 슬랙웹훅설정() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '🔔 Slack Webhook 설정',
    'Slack Incoming Webhook URL을 입력하세요:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const webhookUrl = response.getResponseText().trim();
    if (webhookUrl.startsWith('https://hooks.slack.com/')) {
      PropertiesService.getScriptProperties().setProperty('SLACK_WEBHOOK_URL', webhookUrl);
      ui.alert('✅ Slack Webhook URL이 저장되었습니다!');
      Logger.log('✅ Slack Webhook URL 설정 완료');
    } else {
      ui.alert('❌ 올바른 Slack Webhook URL이 아닙니다.');
    }
  }
}

/**
 * Slack으로 메시지 전송
 * @param {string} message - 메시지 텍스트
 * @param {string} linkUrl - 링크 URL (선택)
 * @param {string} linkTitle - 링크 제목 (선택)
 */
function 슬랙알림(message, linkUrl, linkTitle) {
  const webhookUrl = getSlackWebhookUrl();
  if (!webhookUrl) {
    Logger.log('⚠️ Slack Webhook URL이 설정되지 않았습니다. 슬랙웹훅설정() 함수를 먼저 실행하세요.');
    return;
  }

  try {
    const payload = {
      text: message
    };

    if (linkUrl) {
      payload.attachments = [{
        color: '#36a64f',
        title: linkTitle || '바로가기',
        title_link: linkUrl
      }];
    }

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(webhookUrl, options);

    if (response.getResponseCode() === 200) {
      Logger.log('✅ Slack 알림 전송 완료');
    } else {
      Logger.log('⚠️ Slack 알림 실패:', response.getContentText());
    }
  } catch (e) {
    Logger.log('❌ Slack 알림 오류:', e.message);
  }
}

/**
 * 일일 다이제스트 Slack 알림
 */
function 일일다이제스트_슬랙알림(dateStr, 참여인원) {
  const url = `${CONFIG.WEB_APP_URL}?date=${dateStr}`;
  const message = `📚 *${dateStr} 일일 다이제스트*가 생성되었습니다!\n참여 인원: ${참여인원}명`;
  슬랙알림(message, url, '📖 다이제스트 보기');
}

/**
 * 월간 다이제스트 Slack 알림
 */
function 월간다이제스트_슬랙알림(yearMonth) {
  const url = `${CONFIG.WEB_APP_URL}?date=MONTHLY-${yearMonth}`;
  const message = `📊 *${yearMonth} 월간 다이제스트*가 생성되었습니다!`;
  슬랙알림(message, url, '📈 월간 분석 보기');
}

// ==================== 📢 Slack PDF 전송 ====================

/**
 * Slack Bot Token 가져오기 (Script Properties에서)
 */
function getSlackBotToken() {
  return PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
}

/**
 * Slack 채널 ID 가져오기 (Script Properties에서)
 */
function getSlackChannelId() {
  return PropertiesService.getScriptProperties().getProperty('SLACK_CHANNEL_ID');
}

/**
 * Slack Bot Token 설정 (최초 1회 실행 필요)
 * Slack App에서 Bot Token (xoxb-...)을 발급받아 입력
 * 필요 권한: files:write, chat:write
 */
function 슬랙봇토큰설정() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '🤖 Slack Bot Token 설정',
    'Slack Bot Token (xoxb-...)을 입력하세요:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const token = response.getResponseText().trim();
    if (token.startsWith('xoxb-')) {
      PropertiesService.getScriptProperties().setProperty('SLACK_BOT_TOKEN', token);
      ui.alert('✅ Slack Bot Token이 저장되었습니다!');
      Logger.log('✅ Slack Bot Token 설정 완료');
    } else {
      ui.alert('❌ 올바른 Bot Token이 아닙니다. (xoxb-로 시작해야 합니다)');
    }
  }
}

/**
 * Slack 채널 ID 설정 (PDF를 보낼 채널)
 * 채널 ID는 Slack에서 채널 우클릭 → 채널 세부정보 보기 → 하단에서 확인
 */
function 슬랙채널설정() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '📢 Slack 채널 ID 설정',
    'PDF를 전송할 채널 ID (C로 시작)를 입력하세요:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const channelId = response.getResponseText().trim();
    if (channelId.startsWith('C') || channelId.startsWith('D') || channelId.startsWith('G')) {
      PropertiesService.getScriptProperties().setProperty('SLACK_CHANNEL_ID', channelId);
      ui.alert('✅ Slack 채널 ID가 저장되었습니다!');
      Logger.log('✅ Slack 채널 ID 설정 완료: ' + channelId);
    } else {
      ui.alert('❌ 올바른 채널 ID가 아닙니다. (C, D, G로 시작해야 합니다)');
    }
  }
}

/**
 * HTML 파일을 PDF로 변환
 * @param {string} fileId - Google Drive HTML 파일 ID
 * @returns {Blob} PDF Blob
 */
function HTML파일을PDF로변환(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const htmlContent = file.getBlob().getDataAsString('UTF-8');

    // HTML을 임시 Google Doc으로 변환 후 PDF로 export
    const tempDoc = DocumentApp.create('temp_digest_pdf');
    const tempDocId = tempDoc.getId();

    // HTML 내용을 간단히 텍스트로 추출 (스타일 제거)
    const textContent = htmlContent
      .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
      .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
      .replace(/<[^>]+>/g, '\n')
      .replace(/&nbsp;/g, ' ')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&amp;/g, '&')
      .replace(/\n\s*\n/g, '\n\n')
      .trim();

    tempDoc.getBody().setText(textContent);
    tempDoc.saveAndClose();

    // PDF로 변환
    const pdfBlob = DriveApp.getFileById(tempDocId).getAs('application/pdf');

    // 임시 파일 삭제
    DriveApp.getFileById(tempDocId).setTrashed(true);

    return pdfBlob;
  } catch (e) {
    Logger.log('❌ PDF 변환 오류: ' + e.message);
    return null;
  }
}

/**
 * 다이제스트 HTML을 PDF로 변환 (시트에서 파일 ID 조회)
 * @param {string} dateStr - 날짜 (yyyy-MM-dd 또는 MONTHLY-yyyy-MM)
 * @returns {Blob} PDF Blob
 */
function 다이제스트PDF생성(dateStr) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.DIGEST_SHEET);

    if (!sheet) {
      Logger.log('⚠️ 다이제스트 시트가 없습니다.');
      return null;
    }

    const data = sheet.getDataRange().getValues();
    let fileId = null;

    // 날짜로 파일 ID 찾기
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === dateStr) {
        fileId = data[i][1]; // B열: 파일 ID
        break;
      }
    }

    if (!fileId) {
      Logger.log(`❌ ${dateStr} 다이제스트를 찾을 수 없습니다.`);
      return null;
    }

    // HTML 파일 가져오기
    const file = DriveApp.getFileById(fileId);
    const htmlContent = file.getBlob().getDataAsString('UTF-8');

    // HTML을 PDF로 변환 (Apps Script 내장 기능 활용)
    const pdfBlob = Utilities.newBlob(htmlContent, 'text/html', `digest-${dateStr}.html`)
      .getAs('application/pdf')
      .setName(`다이제스트-${dateStr}.pdf`);

    Logger.log(`✅ PDF 생성 완료: 다이제스트-${dateStr}.pdf`);
    return pdfBlob;

  } catch (e) {
    Logger.log('❌ PDF 생성 오류: ' + e.message);

    // 대안: HTML 파일 자체를 첨부
    try {
      const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
      const sheet = ss.getSheetByName(CONFIG.DIGEST_SHEET);
      const data = sheet.getDataRange().getValues();

      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === dateStr) {
          const file = DriveApp.getFileById(data[i][1]);
          return file.getBlob().setName(`다이제스트-${dateStr}.html`);
        }
      }
    } catch (e2) {
      Logger.log('❌ HTML 파일 로드도 실패: ' + e2.message);
    }
    return null;
  }
}

/**
 * Slack에 PDF 파일 업로드
 * @param {Blob} pdfBlob - PDF 파일 Blob
 * @param {string} message - 함께 보낼 메시지
 * @param {string} channelId - 채널 ID (선택, 없으면 설정된 기본 채널 사용)
 */
function 슬랙PDF업로드(pdfBlob, message, channelId) {
  const token = getSlackBotToken();
  const channel = channelId || getSlackChannelId();

  if (!token) {
    Logger.log('⚠️ Slack Bot Token이 설정되지 않았습니다. 슬랙봇토큰설정() 함수를 먼저 실행하세요.');
    return false;
  }

  if (!channel) {
    Logger.log('⚠️ Slack 채널 ID가 설정되지 않았습니다. 슬랙채널설정() 함수를 먼저 실행하세요.');
    return false;
  }

  if (!pdfBlob) {
    Logger.log('⚠️ 업로드할 파일이 없습니다.');
    return false;
  }

  try {
    // Step 1: files.getUploadURLExternal로 업로드 URL 받기
    const getUrlResponse = UrlFetchApp.fetch('https://slack.com/api/files.getUploadURLExternal', {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + token
      },
      payload: {
        filename: pdfBlob.getName(),
        length: pdfBlob.getBytes().length
      },
      muteHttpExceptions: true
    });

    const urlResult = JSON.parse(getUrlResponse.getContentText());

    if (!urlResult.ok) {
      Logger.log('❌ 업로드 URL 요청 실패: ' + urlResult.error);
      return false;
    }

    // Step 2: 파일 업로드
    const uploadResponse = UrlFetchApp.fetch(urlResult.upload_url, {
      method: 'post',
      payload: pdfBlob.getBytes(),
      muteHttpExceptions: true
    });

    if (uploadResponse.getResponseCode() !== 200) {
      Logger.log('❌ 파일 업로드 실패');
      return false;
    }

    // Step 3: files.completeUploadExternal로 업로드 완료
    const completeResponse = UrlFetchApp.fetch('https://slack.com/api/files.completeUploadExternal', {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        files: [{
          id: urlResult.file_id,
          title: pdfBlob.getName()
        }],
        channel_id: channel,
        initial_comment: message
      }),
      muteHttpExceptions: true
    });

    const completeResult = JSON.parse(completeResponse.getContentText());

    if (completeResult.ok) {
      Logger.log('✅ Slack PDF 업로드 완료: ' + pdfBlob.getName());
      return true;
    } else {
      Logger.log('❌ Slack 업로드 완료 실패: ' + completeResult.error);
      return false;
    }

  } catch (e) {
    Logger.log('❌ Slack PDF 업로드 오류: ' + e.message);
    return false;
  }
}

/**
 * 일일 다이제스트 Slack PDF 전송
 */
function 일일다이제스트_슬랙PDF전송(dateStr, 참여인원) {
  const message = `📚 *${dateStr} 일일 다이제스트*\n참여 인원: ${참여인원}명`;
  const pdfBlob = 다이제스트PDF생성(dateStr);

  if (pdfBlob) {
    슬랙PDF업로드(pdfBlob, message);
  } else {
    // PDF 생성 실패 시 링크로 대체
    Logger.log('⚠️ PDF 생성 실패, 링크로 대체 전송');
    일일다이제스트_슬랙알림(dateStr, 참여인원);
  }
}

/**
 * 월간 다이제스트 Slack PDF 전송
 */
function 월간다이제스트_슬랙PDF전송(yearMonth) {
  const message = `📊 *${yearMonth} 월간 다이제스트*`;
  const pdfBlob = 다이제스트PDF생성(`MONTHLY-${yearMonth}`);

  if (pdfBlob) {
    슬랙PDF업로드(pdfBlob, message);
  } else {
    // PDF 생성 실패 시 링크로 대체
    Logger.log('⚠️ PDF 생성 실패, 링크로 대체 전송');
    월간다이제스트_슬랙알림(yearMonth);
  }
}

// ==================== 🗑️ 멤버 삭제 ====================

/**
 * 특정 멤버의 모든 데이터를 삭제하고 JSON 재생성
 * @param {string} memberName - 삭제할 멤버 이름
 */
function 멤버데이터삭제(memberName) {
  if (!memberName) {
    Logger.log('❌ 멤버 이름을 입력해주세요.');
    return;
  }

  Logger.log(`\n=== ${memberName} 멤버 데이터 삭제 시작 ===\n`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let totalDeleted = 0;

  // 1. 주간집계 시트에서 삭제
  const weeklySheet = ss.getSheetByName('주간집계');
  if (weeklySheet && weeklySheet.getLastRow() > 1) {
    const deleted = 시트에서멤버삭제(weeklySheet, memberName, 1); // B열(이름)
    totalDeleted += deleted;
    Logger.log(`✅ 주간집계: ${deleted}행 삭제`);
  }

  // 2. 월별결산 시트에서 삭제
  const summarySheet = ss.getSheetByName(CONFIG.MONTHLY_SUMMARY_SHEET);
  if (summarySheet && summarySheet.getLastRow() > 1) {
    const deleted = 시트에서멤버삭제(summarySheet, memberName, 1); // B열(조원명)
    totalDeleted += deleted;
    Logger.log(`✅ 월별결산: ${deleted}행 삭제`);
  }

  // 3. 제출기록 시트에서 삭제
  const recordSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (recordSheet && recordSheet.getLastRow() > 1) {
    const deleted = 시트에서멤버삭제(recordSheet, memberName, 1); // B열(이름)
    totalDeleted += deleted;
    Logger.log(`✅ 제출기록: ${deleted}행 삭제`);
  }

  // 4. 벌칙관리 시트에서 삭제
  const penaltySheet = ss.getSheetByName(CONFIG.PENALTY_SHEET);
  if (penaltySheet && penaltySheet.getLastRow() > 1) {
    const deleted = 시트에서멤버삭제(penaltySheet, memberName, 1); // B열(조원명)
    totalDeleted += deleted;
    Logger.log(`✅ 벌칙관리: ${deleted}행 삭제`);
  }

  // 5. JSON 파일에서 멤버 삭제 및 재생성
  Logger.log(`\n📁 JSON 파일 재생성 중...`);
  try {
    JSON파일생성();
    Logger.log(`✅ 출석 JSON 재생성 완료`);
  } catch (e) {
    Logger.log(`⚠️ 출석 JSON 재생성 실패: ${e.message}`);
  }

  try {
    이번주주간집계();
    Logger.log(`✅ 주간집계 JSON 재생성 완료`);
  } catch (e) {
    Logger.log(`⚠️ 주간집계 JSON 재생성 실패: ${e.message}`);
  }

  Logger.log(`\n=== 완료: 총 ${totalDeleted}행 삭제됨 ===`);
  Logger.log(`💡 웹페이지에서 Ctrl+Shift+R로 새로고침하세요.`);
}

/**
 * 시트에서 특정 멤버의 행 삭제
 * @param {Sheet} sheet - 대상 시트
 * @param {string} memberName - 멤버 이름
 * @param {number} nameColIndex - 이름이 있는 열 인덱스 (0-based)
 * @returns {number} 삭제된 행 수
 */
function 시트에서멤버삭제(sheet, memberName, nameColIndex) {
  const data = sheet.getDataRange().getValues();
  let deletedCount = 0;

  // 아래에서 위로 삭제 (행 번호 변경 방지)
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][nameColIndex]).trim() === memberName) {
      sheet.deleteRow(i + 1);
      deletedCount++;
    }
  }

  return deletedCount;
}

/**
 * 황인섭, 호호 멤버 데이터 일괄 삭제 (편의 함수)
 */
function 탈퇴멤버일괄삭제() {
  멤버데이터삭제('황인섭');
  멤버데이터삭제('호호');
  멤버데이터삭제('스카피');
}

/**
 * 주간집계 시트의 모든 멤버 이름 출력 (디버그용)
 */
function 주간집계멤버목록() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('주간집계');

  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('주간집계 시트가 비어있습니다.');
    return;
  }

  const data = sheet.getDataRange().getValues();
  const members = new Set();

  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][1]).trim(); // B열 (조원명)
    if (name) members.add(name);
  }

  Logger.log('=== 주간집계 시트 멤버 목록 ===');
  for (const name of members) {
    Logger.log(`  - "${name}" (길이: ${name.length})`);
  }
  Logger.log(`총 ${members.size}명`);
}

/**
 * 특정 멤버의 폴더 내용 확인 (디버그용)
 * @param {string} memberName - 멤버 이름 (기본: 보노보노)
 * @param {string} targetDate - 확인할 날짜 (기본: 어제)
 */
function 멤버폴더확인(memberName, targetDate) {
  memberName = memberName || '보노보노';

  if (!targetDate) {
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    targetDate = Utilities.formatDate(yesterday, 'Asia/Seoul', 'yyyy-MM-dd');
  }

  const folderIdOrArray = CONFIG.MEMBERS[memberName];
  if (!folderIdOrArray) {
    Logger.log(`❌ ${memberName} 멤버를 찾을 수 없습니다.`);
    return;
  }

  const folderIds = Array.isArray(folderIdOrArray) ? folderIdOrArray : [folderIdOrArray];

  Logger.log(`=== ${memberName} 폴더 확인 (${targetDate}) ===`);
  Logger.log(`폴더 ID: ${folderIds.join(', ')}`);
  Logger.log('');

  for (const folderId of folderIds) {
    try {
      const mainFolder = DriveApp.getFolderById(folderId);
      Logger.log(`📁 메인 폴더: ${mainFolder.getName()}`);

      const subfolders = mainFolder.getFolders();
      let found = false;

      while (subfolders.hasNext()) {
        const folder = subfolders.next();
        const folderName = folder.getName().trim();
        const dateInfo = 날짜추출(folderName);

        if (dateInfo && dateInfo.dateStr === targetDate) {
          found = true;
          Logger.log(`\n✅ 날짜 폴더 발견: "${folderName}"`);
          Logger.log(`   추출된 날짜: ${dateInfo.dateStr}`);
          Logger.log(`   폴더 생성시간: ${folder.getDateCreated().toLocaleString('ko-KR')}`);
          Logger.log(`   마지막 수정: ${folder.getLastUpdated().toLocaleString('ko-KR')}`);

          const files = folder.getFiles();
          let fileCount = 0;
          Logger.log(`   파일 목록:`);
          while (files.hasNext()) {
            const file = files.next();
            fileCount++;
            Logger.log(`     - ${file.getName()} (${file.getDateCreated().toLocaleString('ko-KR')})`);
          }
          Logger.log(`   총 파일: ${fileCount}개`);

          // 마감 체크
          const isClosed = 날짜마감확인(targetDate);
          Logger.log(`   마감 여부: ${isClosed ? '마감됨' : '진행중'}`);
        }
      }

      if (!found) {
        Logger.log(`\n❌ ${targetDate} 날짜의 폴더를 찾을 수 없습니다.`);
        Logger.log(`\n최근 폴더 목록 (최대 10개):`);
        const allFolders = mainFolder.getFolders();
        let count = 0;
        while (allFolders.hasNext() && count < 10) {
          const f = allFolders.next();
          Logger.log(`  - ${f.getName()}`);
          count++;
        }
      }

    } catch (e) {
      Logger.log(`❌ 폴더 접근 오류: ${e.message}`);
    }
  }

  // 제출기록 시트에서 확인
  Logger.log(`\n=== 제출기록 시트 확인 ===`);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (sheet && sheet.getLastRow() > 1) {
    const data = sheet.getDataRange().getValues();
    let foundInSheet = false;
    for (let i = 1; i < data.length; i++) {
      const name = String(data[i][1]).trim();
      const dateStr = data[i][2];
      const formattedDate = typeof dateStr === 'string' ? dateStr : Utilities.formatDate(new Date(dateStr), 'Asia/Seoul', 'yyyy-MM-dd');
      if (name === memberName && formattedDate === targetDate) {
        foundInSheet = true;
        Logger.log(`✅ 제출기록 발견:`);
        Logger.log(`   날짜: ${formattedDate}`);
        Logger.log(`   상태: ${data[i][6]}`);
        Logger.log(`   파일수: ${data[i][3]}`);
        break;
      }
    }
    if (!foundInSheet) {
      Logger.log(`❌ ${memberName}의 ${targetDate} 제출기록 없음`);
    }
  }
}

// ==================== 🚨 벌칙 관리 ====================

/**
 * 전월 벌칙 대상자 등록 + Slack 알림 (에디터에서 바로 실행 가능)
 */
function 수동_벌칙알림() {
  var now = new Date();
  var lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  var year = lastMonth.getFullYear();
  var month = lastMonth.getMonth();  // 0-based
  var ym = year + '-' + String(month + 1).padStart(2, '0');
  Logger.log('수동 벌칙 알림 실행: ' + ym);

  // 월별주간집계 함수로 직접 계산 (시트 중복/형식 문제 우회)
  var 집계결과 = 월별주간집계(year, month);
  var memberAbsences = {};
  for (var name in 집계결과) {
    memberAbsences[name] = 집계결과[name].총결석;
  }
  Logger.log('주간집계 결석: ' + JSON.stringify(memberAbsences));

  // 벌칙 대상자 등록 (4회 이상)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var penaltySheet = ss.getSheetByName('벌칙관리');
  if (!penaltySheet) { 벌칙관리시트_생성(); penaltySheet = ss.getSheetByName('벌칙관리'); }
  var existingData = penaltySheet.getLastRow() > 1 ? penaltySheet.getRange(2, 1, penaltySheet.getLastRow() - 1, 2).getDisplayValues() : [];
  var existingKeys = {};
  for (var j = 0; j < existingData.length; j++) { existingKeys[existingData[j][0] + '|' + existingData[j][1]] = true; }

  var newRows = [];
  var penaltyNames = [];
  for (var name in memberAbsences) {
    if (memberAbsences[name] >= 4) {
      penaltyNames.push(name + '(' + memberAbsences[name] + '회)');
      if (!existingKeys[ym + '|' + name]) {
        newRows.push([ym, name, memberAbsences[name], '🚨 벌칙', '미수행', '']);
      }
    }
  }

  Logger.log('벌칙 대상자 ' + penaltyNames.length + '명: ' + penaltyNames.join(', '));

  if (newRows.length > 0) {
    penaltySheet.getRange(penaltySheet.getLastRow() + 1, 1, newRows.length, 6).setValues(newRows);
    for (var k = 0; k < newRows.length; k++) {
      penaltySheet.getRange(penaltySheet.getLastRow() - newRows.length + 1 + k, 5).setBackground('#ffcdd2');
    }
    Logger.log('🚨 ' + ym + ' 벌칙 대상자 ' + newRows.length + '명 등록');
  } else {
    Logger.log('신규 등록 대상 없음 (이미 등록됨)');
  }

  벌칙대상자_슬랙알림(ym);
}

/**
 * 특정 월 벌칙 데이터 초기화 (잘못 등록된 데이터 정리용)
 */
function 벌칙데이터_초기화() {
  var ym = '2026-02';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('벌칙관리');
  if (!sheet || sheet.getLastRow() < 2) { Logger.log('벌칙관리 시트 비어있음'); return; }
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  var deleted = 0;
  for (var i = data.length - 1; i >= 0; i--) {
    if (String(data[i][0]).trim() === ym) {
      sheet.deleteRow(i + 2);
      deleted++;
    }
  }
  Logger.log(ym + ' 벌칙 데이터 ' + deleted + '행 삭제 완료');
}

/**
 * 월별결산 시트 디버그 (데이터 확인용)
 */
function 주간집계_디버그() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('주간집계');
  if (!sheet) { Logger.log('주간집계 시트 없음'); return; }
  var lastRow = sheet.getLastRow();
  Logger.log('주간집계 총 행수: ' + lastRow);
  if (lastRow < 2) { Logger.log('데이터 없음'); return; }
  var allYM = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  var ymSet = {};
  for (var i = 0; i < allYM.length; i++) {
    var v = String(allYM[i][0]).trim();
    if (!ymSet[v]) ymSet[v] = 0;
    ymSet[v]++;
  }
  Logger.log('연월별 행 수: ' + JSON.stringify(ymSet));
  // 첫 2행 샘플
  var sample = sheet.getRange(2, 1, Math.min(2, lastRow - 1), 9).getDisplayValues();
  for (var i = 0; i < sample.length; i++) {
    Logger.log('샘플 행' + (i+2) + ': [' + sample[i].join('] [') + ']');
  }
}

/**
 * 주간집계 시트의 중복 데이터 정리 및 연월 형식 통일
 * - "2026-2" → "2026-02" 형식으로 통일
 * - 같은 연월+조원명+주차 중복 행 제거
 */
function 주간집계_중복정리() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('주간집계');
  if (!sheet || sheet.getLastRow() < 2) { Logger.log('주간집계 시트 비어있음'); return; }

  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, 9).getDisplayValues();
  Logger.log('정리 전 총 ' + data.length + '행');

  // 1단계: 연월 형식 통일 + 중복 제거
  var seen = {};
  var keepRows = [];
  var duplicateCount = 0;
  var formatFixCount = 0;

  for (var i = 0; i < data.length; i++) {
    var row = data[i].slice();  // 복사
    var ym = String(row[0]).trim();
    var parts = ym.split('-');

    // 연월 형식 통일: "2026-2" → "2026-02"
    if (parts.length === 2) {
      var fixedYM = parts[0] + '-' + String(parseInt(parts[1])).padStart(2, '0');
      if (fixedYM !== ym) {
        row[0] = fixedYM;
        formatFixCount++;
      }
    }

    // 중복 체크 (연월 + 조원명 + 주차)
    var key = row[0] + '|' + String(row[1]).trim() + '|' + String(row[2]).trim();
    if (seen[key]) {
      duplicateCount++;
    } else {
      seen[key] = true;
      keepRows.push(row);
    }
  }

  Logger.log('형식 수정: ' + formatFixCount + '건, 중복 제거: ' + duplicateCount + '건');
  Logger.log('정리 후 ' + keepRows.length + '행');

  // 2단계: 시트 재작성
  if (duplicateCount > 0 || formatFixCount > 0) {
    // 기존 데이터 삭제
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 9).clearContent();
    }
    // 정리된 데이터 쓰기 (연월 열을 텍스트 서식으로 강제 지정)
    if (keepRows.length > 0) {
      sheet.getRange(2, 1, keepRows.length, 1).setNumberFormat('@');
      sheet.getRange(2, 1, keepRows.length, 9).setValues(keepRows);
    }
    // 남은 빈 행 삭제
    var newLastRow = keepRows.length + 1;
    if (sheet.getMaxRows() > newLastRow + 1) {
      sheet.deleteRows(newLastRow + 1, sheet.getMaxRows() - newLastRow);
    }
    Logger.log('✅ 주간집계 시트 정리 완료');
  } else {
    Logger.log('✅ 정리할 항목 없음');
  }
}

/**
 * 벌칙관리 시트 생성 (초기설정에서 호출)
 */
function 벌칙관리시트_생성() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.PENALTY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.PENALTY_SHEET);
    const headers = ['연월', '조원명', '결석수', '상태', '수행여부', '확인일시'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f44336').setFontColor('white');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 100);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 80);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 100);
    sheet.setColumnWidth(6, 180);
    Logger.log('✅ 벌칙관리 시트 생성 완료');
  } else {
    Logger.log('✅ 벌칙관리 시트 이미 존재함');
  }
}

/**
 * 주간집계에서 벌칙 대상자를 벌칙관리 시트에 등록
 * 벌칙 기준: 주간집계 총결석수 4회 이상
 * @param {string} yearMonth - 연월 (예: '2026-01')
 */
function 벌칙대상자_등록(yearMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberAbsences = {};

  // yearMonth 비교용: 숫자로 비교하여 "2026-02"와 "2026-2" 모두 매칭
  const ymParts = yearMonth.split('-');
  const targetYear = parseInt(ymParts[0]);
  const targetMonth = parseInt(ymParts[1]);

  // 1차: 주간집계 시트에서 해당 월 데이터 읽기
  const weeklySheet = ss.getSheetByName('주간집계');
  if (weeklySheet && weeklySheet.getLastRow() >= 2) {
    const data = weeklySheet.getRange(2, 1, weeklySheet.getLastRow() - 1, 9).getDisplayValues();

    for (const row of data) {
      const rowYM = String(row[0]).trim();
      const rowParts = rowYM.split('-');
      if (rowParts.length !== 2 || parseInt(rowParts[0]) !== targetYear || parseInt(rowParts[1]) !== targetMonth) continue;

      const name = String(row[1]).trim();
      const absences = Number(row[6]) || 0;
      const isFullLongOff = String(row[8]).trim() === '전체장기오프';

      if (!memberAbsences[name]) memberAbsences[name] = 0;
      if (!isFullLongOff) memberAbsences[name] += absences;
    }
  }

  // 2차: 주간집계에 데이터가 없으면 월별결산 시트에서 폴백 (숫자 비교로 형식 차이 처리)
  if (Object.keys(memberAbsences).length === 0) {
    Logger.log('⚠️ 주간집계에 ' + yearMonth + ' 데이터 없음 → 월별결산에서 검색');
    const summarySheet = ss.getSheetByName(CONFIG.MONTHLY_SUMMARY_SHEET);
    if (summarySheet && summarySheet.getLastRow() >= 2) {
      const sData = summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, 9).getDisplayValues();
      for (const row of sData) {
        const rowYM = String(row[0]).trim();
        const rowParts = rowYM.split('-');
        if (rowParts.length !== 2 || parseInt(rowParts[0]) !== targetYear || parseInt(rowParts[1]) !== targetMonth) continue;
        const name = String(row[1]).trim();
        const absences = Number(row[5]) || 0;  // 월별결산 6번째 열 = 결석
        memberAbsences[name] = absences;
      }
    }
  }

  Logger.log('벌칙 대상자 검색: ' + yearMonth + ' (조원 ' + Object.keys(memberAbsences).length + '명 데이터)');

  // 총결석 4회 이상 → 벌칙 대상
  const penaltyMembers = [];
  for (const [name, totalAbsences] of Object.entries(memberAbsences)) {
    if (totalAbsences >= 4) {
      penaltyMembers.push({
        name: name,
        absences: totalAbsences,
        status: '🚨 벌칙'
      });
    }
  }

  Logger.log(`벌칙 대상자 ${penaltyMembers.length}명 발견: ${penaltyMembers.map(m => `${m.name}(${m.absences}회)`).join(', ')}`);

  if (penaltyMembers.length === 0) {
    Logger.log(`📋 ${yearMonth} 벌칙 대상자 없음`);
    return [];
  }

  // 벌칙관리 시트에 등록
  let penaltySheet = ss.getSheetByName(CONFIG.PENALTY_SHEET);
  if (!penaltySheet) {
    벌칙관리시트_생성();
    penaltySheet = ss.getSheetByName(CONFIG.PENALTY_SHEET);
  }

  // 기존 데이터 확인 (중복 방지)
  const existingData = penaltySheet.getLastRow() > 1
    ? penaltySheet.getRange(2, 1, penaltySheet.getLastRow() - 1, 2).getValues()
    : [];
  const existingKeys = new Set(existingData.map(r => `${r[0]}|${r[1]}`));

  const newRows = [];
  for (const member of penaltyMembers) {
    const key = `${yearMonth}|${member.name}`;
    if (!existingKeys.has(key)) {
      newRows.push([yearMonth, member.name, member.absences, member.status, '미수행', '']);
    }
  }

  if (newRows.length > 0) {
    penaltySheet.getRange(penaltySheet.getLastRow() + 1, 1, newRows.length, 6).setValues(newRows);
    // 색상 적용
    for (let i = 0; i < newRows.length; i++) {
      const rowNum = penaltySheet.getLastRow() - newRows.length + 1 + i;
      penaltySheet.getRange(rowNum, 5).setBackground('#ffcdd2'); // 미수행 빨강
    }
    Logger.log(`🚨 ${yearMonth} 벌칙 대상자 ${newRows.length}명 등록`);
  } else {
    Logger.log(`📋 ${yearMonth} 벌칙 대상자 이미 등록됨`);
  }

  return penaltyMembers;
}

/**
 * Slack Block Kit 벌칙 메시지 블록 생성
 * @param {string} yearMonth - 연월
 * @returns {Array} Block Kit blocks 배열
 */
function 벌칙메시지블록생성(yearMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const penaltySheet = ss.getSheetByName(CONFIG.PENALTY_SHEET);

  if (!penaltySheet || penaltySheet.getLastRow() < 2) {
    return [];
  }

  const data = penaltySheet.getRange(2, 1, penaltySheet.getLastRow() - 1, 6).getDisplayValues();
  const members = data.filter(r => String(r[0]).trim() === yearMonth);

  if (members.length === 0) return [];

  const year = yearMonth.split('-')[0];
  const month = parseInt(yearMonth.split('-')[1]);

  const blocks = [
    {
      type: 'header',
      text: { type: 'plain_text', text: `🚨 ${year}년 ${month}월 벌칙 대상자`, emoji: true }
    },
    {
      type: 'section',
      text: { type: 'mrkdwn', text: '벌칙 수행 완료 후 *✅ 완료* 버튼을 눌러주세요.' }
    },
    { type: 'divider' }
  ];

  for (const row of members) {
    const name = String(row[1]);
    const absences = Number(row[2]);
    const completed = String(row[4]) === '완료';
    const confirmedAt = row[5] ? String(row[5]) : '';

    if (completed) {
      blocks.push({
        type: 'section',
        text: { type: 'mrkdwn', text: '✅ *' + name + '* - 결석 ' + absences + '회 → 완료 (' + confirmedAt + ')' }
      });
    } else {
      var completeUrl = CONFIG.WEB_APP_URL
        + '?action=penaltyComplete'
        + '&name=' + encodeURIComponent(name)
        + '&yearMonth=' + encodeURIComponent(yearMonth);
      blocks.push({
        type: 'section',
        text: { type: 'mrkdwn', text: '🚨 *' + name + '* - 결석 ' + absences + '회' },
        accessory: {
          type: 'button',
          text: { type: 'plain_text', text: '✅ 완료', emoji: true },
          style: 'primary',
          url: completeUrl
        }
      });
    }
  }

  // 요약
  const doneCount = members.filter(r => String(r[4]) === '완료').length;
  blocks.push({ type: 'divider' });
  blocks.push({
    type: 'context',
    elements: [{ type: 'mrkdwn', text: `📊 수행 현황: ${doneCount}/${members.length}명 완료` }]
  });

  return blocks;
}

/**
 * 벌칙 대상자 Slack 알림 (웹훅 + Block Kit 인터랙티브 메시지)
 * 일간 다이제스트와 동일한 웹훅 방식 사용
 * @param {string} yearMonth - 연월
 */
function 벌칙대상자_슬랙알림(yearMonth) {
  const webhookUrl = getSlackWebhookUrl();
  if (!webhookUrl) {
    Logger.log('⚠️ Slack Webhook URL이 설정되지 않았습니다. 슬랙웹훅설정() 함수를 먼저 실행하세요.');
    return;
  }

  const blocks = 벌칙메시지블록생성(yearMonth);
  if (blocks.length === 0) {
    Logger.log(`📋 ${yearMonth} 벌칙 대상자 없음 - Slack 알림 스킵`);
    return;
  }

  try {
    const response = UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        text: '🚨 ' + yearMonth + ' 벌칙 대상자 알림',
        blocks: blocks
      }),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      Logger.log('✅ ' + yearMonth + ' 벌칙 알림 전송 완료');
    } else {
      Logger.log('❌ Slack 벌칙 알림 실패: ' + response.getContentText());
    }
  } catch (e) {
    Logger.log('❌ Slack 벌칙 알림 오류: ' + e.message);
  }
}

/**
 * 벌칙 수행 완료 처리
 * @param {string} memberName - 조원명
 * @param {string} yearMonth - 연월
 * @returns {boolean} 성공 여부
 */
function 벌칙완료처리(memberName, yearMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const penaltySheet = ss.getSheetByName(CONFIG.PENALTY_SHEET);

  if (!penaltySheet || penaltySheet.getLastRow() < 2) {
    Logger.log('⚠️ 벌칙관리 시트가 비어있습니다.');
    return false;
  }

  const data = penaltySheet.getRange(2, 1, penaltySheet.getLastRow() - 1, 6).getDisplayValues();

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === yearMonth && String(data[i][1]).trim() === memberName) {
      const rowNum = i + 2;
      const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
      penaltySheet.getRange(rowNum, 5).setValue('완료');
      penaltySheet.getRange(rowNum, 6).setValue(now);
      penaltySheet.getRange(rowNum, 5).setBackground('#e8f5e9'); // 완료 초록
      Logger.log(`✅ 벌칙 완료 처리: ${memberName} (${yearMonth}) - ${now}`);
      return true;
    }
  }

  Logger.log(`⚠️ 벌칙 기록을 찾을 수 없음: ${memberName} (${yearMonth})`);
  return false;
}

/**
 * 벌칙 리마인더 (매주 월요일 실행)
 * 미수행 벌칙이 있으면 Slack으로 리마인더 전송
 */
function 벌칙리마인더() {
  const now = new Date();
  if (now.getDay() !== 1) {
    Logger.log('벌칙리마인더: 월요일이 아니므로 스킵');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const penaltySheet = ss.getSheetByName(CONFIG.PENALTY_SHEET);

  if (!penaltySheet || penaltySheet.getLastRow() < 2) {
    Logger.log('벌칙리마인더: 벌칙 데이터 없음');
    return;
  }

  const data = penaltySheet.getRange(2, 1, penaltySheet.getLastRow() - 1, 6).getDisplayValues();

  // 미수행 건이 있는 연월 목록
  const pendingMonths = new Set();
  for (const row of data) {
    if (String(row[4]).trim() === '미수행') {
      pendingMonths.add(String(row[0]));
    }
  }

  if (pendingMonths.size === 0) {
    Logger.log('벌칙리마인더: 미수행 벌칙 없음');
    return;
  }

  const webhookUrl = getSlackWebhookUrl();
  if (!webhookUrl) {
    Logger.log('⚠️ Slack Webhook URL 설정 누락으로 리마인더 스킵');
    return;
  }

  // 각 연월별로 리마인더 전송
  for (const yearMonth of pendingMonths) {
    const blocks = 벌칙메시지블록생성(yearMonth);
    if (blocks.length === 0) continue;

    // 리마인더 헤더 추가
    blocks.unshift({
      type: 'section',
      text: { type: 'mrkdwn', text: '⏰ *벌칙 수행 확인 리마인더*\n아직 벌칙 미수행자가 있습니다. 확인해주세요.' }
    });

    try {
      const response = UrlFetchApp.fetch(webhookUrl, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          text: '⏰ ' + yearMonth + ' 벌칙 미수행 리마인더',
          blocks: blocks
        }),
        muteHttpExceptions: true
      });

      if (response.getResponseCode() === 200) {
        Logger.log('✅ ' + yearMonth + ' 벌칙 리마인더 전송 완료');
      } else {
        Logger.log('❌ 리마인더 전송 실패: ' + response.getContentText());
      }
    } catch (e) {
      Logger.log('❌ 리마인더 오류: ' + e.message);
    }
  }
}

// ==================== 📋 시트 메뉴 ====================

/**
 * 시트 열 때 자동으로 커스텀 메뉴 생성
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 출석관리')
    .addItem('🗑️ 오래된 데이터 정리...', '오래된데이터정리_다이얼로그')
    .addItem('📅 특정 월 데이터 삭제...', '특정월데이터삭제_다이얼로그')
    .addSeparator()
    .addItem('🔄 이번 주 주간집계 실행', '이번주주간집계')
    .addItem('📄 이번 달 JSON 재생성', 'JSON파일생성')
    .addSeparator()
    .addItem('⚙️ 초기설정 (트리거 재설정)', '초기설정')
    .addToUi();
}

/**
 * 오래된 데이터 정리 다이얼로그
 */
function 오래된데이터정리_다이얼로그() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    '🗑️ 오래된 데이터 정리',
    '최근 몇 개월의 데이터를 유지할까요? (예: 2 입력 시 현재월 + 이전월만 유지)\n\n' +
    '⚠️ 주의: 삭제된 데이터는 복구할 수 없습니다!',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() === ui.Button.OK) {
    const monthsToKeep = parseInt(result.getResponseText());

    if (isNaN(monthsToKeep) || monthsToKeep < 1) {
      ui.alert('❌ 오류', '1 이상의 숫자를 입력해주세요.', ui.ButtonSet.OK);
      return;
    }

    const confirmResult = ui.alert(
      '⚠️ 확인',
      `최근 ${monthsToKeep}개월 데이터만 유지하고 나머지를 삭제합니다.\n정말 진행하시겠습니까?`,
      ui.ButtonSet.YES_NO
    );

    if (confirmResult === ui.Button.YES) {
      const deletedCount = 오래된데이터정리(monthsToKeep);
      ui.alert('✅ 완료', `${deletedCount}개의 오래된 기록이 삭제되었습니다.`, ui.ButtonSet.OK);
    }
  }
}

/**
 * 특정 월 데이터 삭제 다이얼로그
 */
function 특정월데이터삭제_다이얼로그() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    '📅 특정 월 데이터 삭제',
    '삭제할 연월을 입력하세요 (예: 2025-10)\n\n' +
    '⚠️ 주의: 삭제된 데이터는 복구할 수 없습니다!',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() === ui.Button.OK) {
    const yearMonth = result.getResponseText().trim();

    // 형식 검증 (YYYY-MM)
    if (!/^\d{4}-\d{2}$/.test(yearMonth)) {
      ui.alert('❌ 오류', 'YYYY-MM 형식으로 입력해주세요. (예: 2025-10)', ui.ButtonSet.OK);
      return;
    }

    const confirmResult = ui.alert(
      '⚠️ 확인',
      `${yearMonth} 데이터를 모두 삭제합니다.\n정말 진행하시겠습니까?`,
      ui.ButtonSet.YES_NO
    );

    if (confirmResult === ui.Button.YES) {
      const deletedCount = 특정월데이터삭제(yearMonth);
      ui.alert('✅ 완료', `${yearMonth}: ${deletedCount}개의 기록이 삭제되었습니다.`, ui.ButtonSet.OK);
    }
  }
}

/**
 * 오래된 출석 데이터 정리
 * @param {number} monthsToKeep - 유지할 개월 수 (현재월 포함)
 * @returns {number} 삭제된 행 수
 */
function 오래된데이터정리(monthsToKeep = 2) {
  Logger.log(`=== 오래된 데이터 정리 시작 (최근 ${monthsToKeep}개월 유지) ===`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    Logger.log('❌ 출석 시트를 찾을 수 없습니다.');
    return 0;
  }

  // 유지할 연월 목록 생성
  const now = new Date();
  const keepMonths = new Set();

  for (let i = 0; i < monthsToKeep; i++) {
    const targetDate = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const yearMonth = `${targetDate.getFullYear()}-${String(targetDate.getMonth() + 1).padStart(2, '0')}`;
    keepMonths.add(yearMonth);
  }

  Logger.log(`유지할 월: ${Array.from(keepMonths).join(', ')}`);

  // 데이터 가져오기
  const data = sheet.getDataRange().getValues();
  const rowsToDelete = [];

  // 날짜 컬럼 (3번째 = index 2)
  for (let i = 1; i < data.length; i++) {
    const dateValue = data[i][2];  // 날짜 컬럼

    if (!dateValue) continue;

    let dateStr;
    if (dateValue instanceof Date) {
      dateStr = Utilities.formatDate(dateValue, 'Asia/Seoul', 'yyyy-MM-dd');
    } else {
      dateStr = String(dateValue).trim();
    }

    const yearMonth = dateStr.substring(0, 7);  // "2025-10"

    if (!keepMonths.has(yearMonth)) {
      rowsToDelete.push(i + 1);  // 1-based row number
    }
  }

  Logger.log(`삭제 대상: ${rowsToDelete.length}개 행`);

  // 뒤에서부터 삭제 (인덱스 변경 방지)
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }

  Logger.log(`✅ ${rowsToDelete.length}개 행 삭제 완료`);

  // JSON 재생성
  if (rowsToDelete.length > 0) {
    Logger.log('JSON 파일 재생성 중...');
    JSON파일생성();
    Logger.log('✅ JSON 파일 재생성 완료');
  }

  return rowsToDelete.length;
}

/**
 * 특정 월의 출석 데이터 삭제
 * @param {string} yearMonth - 삭제할 연월 (예: "2025-10")
 * @returns {number} 삭제된 행 수
 */
function 특정월데이터삭제(yearMonth) {
  Logger.log(`=== ${yearMonth} 데이터 삭제 시작 ===`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    Logger.log('❌ 출석 시트를 찾을 수 없습니다.');
    return 0;
  }

  const data = sheet.getDataRange().getValues();
  const rowsToDelete = [];

  for (let i = 1; i < data.length; i++) {
    const dateValue = data[i][2];  // 날짜 컬럼

    if (!dateValue) continue;

    let dateStr;
    if (dateValue instanceof Date) {
      dateStr = Utilities.formatDate(dateValue, 'Asia/Seoul', 'yyyy-MM-dd');
    } else {
      dateStr = String(dateValue).trim();
    }

    if (dateStr.startsWith(yearMonth)) {
      rowsToDelete.push(i + 1);  // 1-based row number
    }
  }

  Logger.log(`${yearMonth}: ${rowsToDelete.length}개 행 삭제 예정`);

  // 뒤에서부터 삭제
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }

  Logger.log(`✅ ${rowsToDelete.length}개 행 삭제 완료`);

  return rowsToDelete.length;
}

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

            // 이번 달만 체크할지, 모든 달을 체크할지 결정 (마감체크보다 먼저 실행하여 성능 최적화)
            const shouldProcess = CONFIG.SCAN_ALL_MONTHS ||
                                  (year === currentYear && month === currentMonth);

            if (!shouldProcess) {
              skippedCount++;
              continue;
            }

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
                // 🆕 off.md 파일 체크 (과도기 지원)
                const hasOffFile = files.some(f =>
                  f.name.toLowerCase() === 'off.md' ||
                  f.name.toLowerCase() === 'off.txt'
                );

                if (hasOffFile) {
                  Logger.log(`    🏖️ ${dateStr} - 오프 (off.md 파일 발견)`);
                  출석기록추가(memberName, dateStr, files, 'OFF', 'off.md 파일');
                } else {
                  Logger.log(`    ✓ ${dateStr} - 출석 (${files.length}개 파일)`);
                  출석기록추가(memberName, dateStr, files, 'O');
                }
                processedDates.add(dateStr);
                processedCount++;
              } else {
                Logger.log(`    ⚠ ${dateStr} - 폴더는 있지만 파일 없음`);
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

/**
 * 🆕 과거 출석 기록 재검사 (off.md 파일 누락 수정용)
 * - 최근 N일간의 출석 기록을 다시 확인
 * - off.md 파일이 있는데 '출석'으로 체크된 경우 '오프'로 수정
 * @param {number} days - 확인할 일수 (기본: 7일)
 */
function 출석기록재검사(days = 7) {
  Logger.log(`=== 최근 ${days}일 출석 기록 재검사 시작 ===\n`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    Logger.log('❌ 제출기록 시트가 없습니다.');
    return;
  }

  // 최근 N일의 날짜 생성
  const targetDates = [];
  const now = new Date();

  for (let i = 0; i < days; i++) {
    const checkDate = new Date(now);
    checkDate.setDate(checkDate.getDate() - i);
    const dateStr = Utilities.formatDate(checkDate, 'Asia/Seoul', 'yyyy-MM-dd');
    targetDates.push(dateStr);
  }

  Logger.log(`📅 검사 대상 날짜: ${targetDates.join(', ')}\n`);

  let totalChecked = 0;
  let totalFixed = 0;

  // 각 조원별로 검사
  for (const [memberName, folderIdOrArray] of Object.entries(CONFIG.MEMBERS)) {
    const folderIds = Array.isArray(folderIdOrArray) ? folderIdOrArray : [folderIdOrArray];

    Logger.log(`👤 ${memberName} 검사 중...`);
    let memberFixed = 0;

    // 각 날짜 검사
    for (const dateStr of targetDates) {
      // 해당 날짜의 현재 출석 상태 확인
      const data = sheet.getDataRange().getValues();
      let currentStatus = null;

      for (let i = 1; i < data.length; i++) {
        if (data[i][1] === memberName && data[i][2] === dateStr) {
          currentStatus = data[i][6]; // 출석상태 열
          break;
        }
      }

      // 출석('O')으로 되어 있는 경우만 재검사
      if (currentStatus !== 'O') {
        continue;
      }

      totalChecked++;

      // 폴더에서 off.md 파일 찾기
      let hasOffFile = false;

      for (const folderId of folderIds) {
        try {
          const memberFolder = DriveApp.getFolderById(folderId);

          // 여러 날짜 형식 시도
          const dateFormats = [];
          const parts = dateStr.split('-');
          if (parts.length === 3) {
            const year = parts[0];
            const month = parts[1];
            const day = parts[2];

            dateFormats.push(
              `${year}-${month}-${day}`,
              `${year}${month}${day}`,
              `${year}.${month}.${day}`,
              `${year}년 ${month}월 ${day}일`,
              `${year}-${parseInt(month)}-${parseInt(day)}`
            );
          }

          // 날짜 폴더 찾기
          let dateFolder = null;
          for (const format of dateFormats) {
            const folders = memberFolder.getFoldersByName(format);
            if (folders.hasNext()) {
              dateFolder = folders.next();
              break;
            }
          }

          if (!dateFolder) {
            continue;
          }

          // 폴더 내 파일 확인
          const files = dateFolder.getFiles();
          while (files.hasNext()) {
            const file = files.next();
            const fileName = file.getName().toLowerCase();

            if (fileName === 'off.md' || fileName === 'off.txt') {
              hasOffFile = true;
              break;
            }
          }

          if (hasOffFile) {
            break; // 찾았으면 다른 폴더 검사 안함
          }

        } catch (e) {
          // 폴더 접근 오류는 무시
        }
      }

      // off.md 파일이 있으면 오프로 수정
      if (hasOffFile) {
        Logger.log(`  🔧 ${dateStr} - 출석 → 오프로 수정 (off.md 발견)`);

        // 기록 업데이트
        const files = 파일목록및링크생성_날짜폴더찾기(memberName, folderIds, dateStr);
        if (files) {
          출석기록추가(memberName, dateStr, files, 'OFF', 'off.md 파일 (재검사로 수정됨)');
          memberFixed++;
          totalFixed++;
        }
      }
    }

    if (memberFixed > 0) {
      Logger.log(`  ✅ ${memberFixed}개 기록 수정됨\n`);
    } else {
      Logger.log(`  ✓ 수정 필요 없음\n`);
    }
  }

  Logger.log(`\n=== 재검사 완료 ===`);
  Logger.log(`📊 총 ${totalChecked}개 기록 검사`);
  Logger.log(`🔧 총 ${totalFixed}개 기록 수정`);

  // JSON 파일 재생성
  if (totalFixed > 0) {
    Logger.log('\n📁 JSON 파일 재생성 중...');
    JSON파일생성();
    Logger.log('✅ JSON 파일 업데이트 완료');
  }
}

/**
 * 날짜 폴더를 찾아서 파일 목록 반환 (재검사용)
 */
function 파일목록및링크생성_날짜폴더찾기(memberName, folderIds, dateStr) {
  for (const folderId of folderIds) {
    try {
      const memberFolder = DriveApp.getFolderById(folderId);

      // 여러 날짜 형식 시도
      const dateFormats = [];
      const parts = dateStr.split('-');
      if (parts.length === 3) {
        const year = parts[0];
        const month = parts[1];
        const day = parts[2];

        dateFormats.push(
          `${year}-${month}-${day}`,
          `${year}${month}${day}`,
          `${year}.${month}.${day}`,
          `${year}년 ${month}월 ${day}일`
        );
      }

      // 날짜 폴더 찾기
      for (const format of dateFormats) {
        const folders = memberFolder.getFoldersByName(format);
        if (folders.hasNext()) {
          const dateFolder = folders.next();
          return 파일목록및링크생성(dateFolder);
        }
      }

    } catch (e) {
      // 폴더 접근 오류는 무시
    }
  }

  return null;
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

  // JSON 재생성 (현재 달 + 처리된 날짜의 달)
  Logger.log('JSON 파일 재생성 중...');
  JSON파일생성();  // 현재 달

  // 처리된 날짜 중 이전 달이 있으면 해당 달도 JSON 재생성
  const now2 = new Date();
  const currentMonth = now2.getMonth() + 1;
  const currentYear = now2.getFullYear();

  const processedMonths = new Set();
  for (const dateStr of targetDates) {
    const [year, month] = dateStr.split('-').map(Number);
    if (year !== currentYear || month !== currentMonth) {
      processedMonths.add(`${year}-${month}`);
    }
  }

  for (const yearMonth of processedMonths) {
    const [year, month] = yearMonth.split('-').map(Number);
    Logger.log(`이전 달 JSON 재생성: ${year}년 ${month}월`);
    특정월JSON생성(year, month);           // 1-based month
    // 주간집계는 매일 새벽 이번주주간집계 트리거에서 처리 (성능 최적화)
  }
}

/**
 * 🆕 OFF 파일 확인 함수 (off.md 또는 off.txt 파일 존재 여부 확인)
 * @param {Folder} folder - 확인할 폴더 객체
 * @returns {boolean} OFF 파일 존재 여부
 */
function OFF파일확인(folder) {
  try {
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName().toLowerCase();
      if (fileName === 'off.md' || fileName === 'off.txt') {
        return true;
      }
    }
    return false;
  } catch (error) {
    Logger.log(`OFF파일확인 오류: ${error.message}`);
    return false;
  }
}

/**
 * 🆕 장기오프 확인 함수 (구글 폼 응답 시트 버전)
 * @param {string} memberName - 조원 이름
 * @param {string} dateStr - 확인할 날짜 (yyyy-MM-dd)
 * @returns {Object} {isLongOff: boolean, reason: string}
 */
function 장기오프확인(memberName, dateStr, cachedLongOffData = null) {
  // 캐시된 데이터가 없으면 시트에서 로드
  let data;
  if (cachedLongOffData) {
    data = cachedLongOffData;
  } else {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.LONG_OFF_SHEET);

    if (!sheet) {
      return { isLongOff: false, reason: '' };
    }

    data = sheet.getDataRange().getValues();
  }

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

  // 2자리 연도 패턴 (26-02-02, 26.02.02, 26/02/02, 260202)
  const shortYearPatterns = [
    /^(\d{2})-(\d{1,2})-(\d{1,2})$/,
    /^(\d{2})\.(\d{1,2})\.(\d{1,2})$/,
    /^(\d{2})\/(\d{1,2})\/(\d{1,2})$/,
    /^(\d{2})(\d{2})(\d{2})$/,
  ];

  for (const pattern of shortYearPatterns) {
    const match = folderName.match(pattern);
    if (match) {
      const shortYear = parseInt(match[1]);
      const year = shortYear >= 0 && shortYear <= 50 ? 2000 + shortYear : 1900 + shortYear;
      const month = parseInt(match[2]) - 1; // 0-based
      const day = parseInt(match[3]);

      if (month >= 0 && month <= 11 && day >= 1 && day <= 31) {
        const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        return { dateStr, year, month, day };
      }
    }
  }
  
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
function 출석기록추가(memberName, date, files, status = 'O', reason = '', folderId = '', directLink = '') {
  const koreaTime = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.appendRow(['타임스탬프', '이름', '날짜', '파일수', '링크', '폴더링크', '출석상태', '주차', '사유']);
    sheet.getRange('A1:I1').setFontWeight('bold').setBackground('#4CAF50').setFontColor('white');
  }

  const linksText = files.map(f => `${f.name}: ${f.url}`).join('\n');

  // 폴더 링크 생성 (directLink가 있으면 우선 사용)
  let folderLink = '';
  if (directLink) {
    folderLink = directLink;  // 🆕 관리자가 직접 입력한 링크
  } else if (status === 'O' && folderId) {
    folderLink = `https://drive.google.com/drive/folders/${folderId}`;
  }
  
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

/**
 * 특정 월의 데이터를 폴더에서 수동 수집
 * @param {string} yearMonth - 년월 (yyyy-MM)
 */
function 월간데이터수동수집(yearMonth) {
  Logger.log(`=== ${yearMonth} 월간 데이터 수동 수집 시작 ===\n`);

  const [year, month] = yearMonth.split('-').map(Number);
  const lastDay = new Date(year, month, 0).getDate();
  const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);
  const fileName = `monthly-data-${yearMonth}.json`;

  // JSON 데이터 초기화
  const jsonData = {
    년월: yearMonth,
    수집일시: new Date().toISOString(),
    조원데이터: {}
  };

  // 각 조원별로 데이터 수집
  for (const [memberName, folderIdOrArray] of Object.entries(CONFIG.MEMBERS)) {
    const folderIds = Array.isArray(folderIdOrArray) ? folderIdOrArray : [folderIdOrArray];

    if (!folderIds[0]) {
      Logger.log(`⚠️ ${memberName}: 폴더 ID 없음, 건너뜀`);
      continue;
    }

    Logger.log(`👤 ${memberName} 데이터 수집 중...`);

    let 한달내용 = '';
    let 출석일수 = 0;
    let 파일수 = 0;

    // 각 날짜별로 폴더 확인
    for (let day = 1; day <= lastDay; day++) {
      const dateStr = `${yearMonth}-${String(day).padStart(2, '0')}`;

      for (const folderId of folderIds) {
        try {
          const memberFolder = DriveApp.getFolderById(folderId);
          const subfolders = memberFolder.getFolders();

          while (subfolders.hasNext()) {
            const subfolder = subfolders.next();
            const folderName = subfolder.getName();

            // 날짜 형식 매칭
            if (folderName.includes(dateStr) ||
                folderName.includes(dateStr.replace(/-/g, '')) ||
                folderName.includes(dateStr.replace(/-/g, '.'))) {

              const files = subfolder.getFiles();
              let dayContent = '';
              let dayFileCount = 0;

              while (files.hasNext()) {
                const file = files.next();
                const fileName = file.getName().toLowerCase();

                // OFF 파일 제외, 텍스트/마크다운 파일만
                if (fileName === 'off.md' || fileName === 'off.txt') continue;

                if (fileName.endsWith('.md') || fileName.endsWith('.txt')) {
                  try {
                    const content = file.getBlob().getDataAsString('UTF-8');
                    dayContent += `\n### ${file.getName()}\n${content}\n`;
                    dayFileCount++;
                  } catch (e) {
                    dayFileCount++;
                  }
                } else {
                  dayFileCount++;
                }
              }

              if (dayFileCount > 0) {
                한달내용 += `\n[${dateStr}]\n${dayContent}\n`;
                출석일수++;
                파일수 += dayFileCount;
              }
              break;
            }
          }
        } catch (e) {
          // 폴더 접근 오류 무시
        }
      }
    }

    if (출석일수 > 0) {
      jsonData.조원데이터[memberName] = {
        한달내용,
        출석일수,
        파일수
      };
      Logger.log(`  ✅ ${출석일수}일 출석, ${파일수}개 파일`);
    } else {
      Logger.log(`  ⚠️ 데이터 없음`);
    }
  }

  // 기존 파일 삭제
  const existingFiles = folder.getFilesByName(fileName);
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }

  // 새 파일 저장
  const newFile = folder.createFile(fileName, JSON.stringify(jsonData, null, 2), MimeType.PLAIN_TEXT);
  newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  Logger.log(`\n✅ 월간 데이터 수집 완료: ${fileName}`);
  Logger.log(`   ${Object.keys(jsonData.조원데이터).length}명 데이터 저장`);
}

/**
 * 특정 월의 JSON 파일 생성
 * @param {number} year - 연도 (예: 2025)
 * @param {number} month - 월 (1-12)
 */
function 특정월JSON생성(year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recordSheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!recordSheet) {
    Logger.log('제출기록 시트가 없습니다.');
    return;
  }

  const records = recordSheet.getDataRange().getValues();
  const jsonData = {};

  // 대상 연월 문자열 생성
  const targetYearMonth = `${year}-${String(month).padStart(2, '0')}`;

  Logger.log(`JSON 생성: ${targetYearMonth} 데이터`);

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

    // 날짜 문자열 정규화
    const dateFormatted = typeof dateStr === 'string'
      ? dateStr
      : Utilities.formatDate(new Date(dateStr), 'Asia/Seoul', 'yyyy-MM-dd');

    // 대상 월 데이터만 필터링
    if (!dateFormatted.startsWith(targetYearMonth)) {
      continue;
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

  // 경고/벌칙 계산
  for (const memberName of Object.keys(jsonData)) {
    const member = jsonData[memberName];
    if (member.결석 >= 3) member.경고 = true;
    if (member.결석 >= 4) member.벌칙 = true;
  }

  // JSON 파일 저장
  const jsonString = JSON.stringify(jsonData, null, 2);
  const fileName = `attendance_summary_${targetYearMonth}.json`;

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

  // 🆕 트리거 5: 매일 새벽 4시 초반 AI 다이제스트 자동 생성 (전날 다이제스트)
  ScriptApp.newTrigger('일일AI다이제스트생성')
    .timeBased()
    .atHour(4)
    .nearMinute(0)
    .everyDays(1)
    .create();

  Logger.log('트리거 5 설정 완료: 매일 새벽 4시 초반 전날 다이제스트 자동 생성');

  // 🆕 트리거 6: 매일 새벽 6시 이번 주 주간집계 (빠른 버전)
  ScriptApp.newTrigger('이번주주간집계')
    .timeBased()
    .atHour(6)
    .everyDays(1)
    .create();

  Logger.log('트리거 6 설정 완료: 매일 새벽 6시 이번 주 주간집계 (빠른 버전)');

  // 🆕 트리거 6-2: 매일 새벽 7시 지난주 결석 확정 (월요일에만 실제 동작)
  ScriptApp.newTrigger('지난주결석확정')
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .create();

  Logger.log('트리거 6-2 설정 완료: 매일 새벽 7시 지난주 결석 확정 (월요일에만 실제 동작)');

  // 🆕 트리거 7: 매월 1일 오전 5시 월간 AI 분석 (전월, 누적된 데이터 사용)
  ScriptApp.newTrigger('월간AI분석_자동실행')
    .timeBased()
    .onMonthDay(1)
    .atHour(5)
    .create();

  Logger.log('트리거 7 설정 완료: 매월 1일 오전 5시 전월 AI 분석 및 다이제스트 생성 (누적 데이터)');

  // 🆕 트리거 8: 매월 1일 오전 6시 월간 원본 수집 (전월, 옵시디언용)
  ScriptApp.newTrigger('월간원본수집_자동실행')
    .timeBased()
    .onMonthDay(1)
    .atHour(6)
    .create();

  Logger.log('트리거 8 설정 완료: 매월 1일 오전 6시 전월 원본 파일 수집 (옵시디언용)');

  // 🆕 트리거 9: 매일 오전 9시 벌칙 리마인더 (월요일에만 실제 동작)
  ScriptApp.newTrigger('벌칙리마인더')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();

  Logger.log('트리거 9 설정 완료: 매일 오전 9시 벌칙 리마인더 (월요일에만 실제 동작)');

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

    // 승인 열(F열) 확인 및 추가
    const lastCol = longOffSheet.getLastColumn();
    if (lastCol < 6) {
      Logger.log('⚠️ 승인 열(F열)이 없습니다. 수동으로 추가해주세요.');
      Logger.log('   현재 열: A(타임스탬프), B(이름), C(시작일), D(종료일), E(사유), F(승인) 필요');
    } else {
      Logger.log('✅ 승인 열(F열) 확인 완료');
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

  // 🆕 벌칙관리 시트 초기화
  벌칙관리시트_생성();

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
  
  // 헤더 설정 (파일링크는 맨 끝)
  const headers = ['조원 이름', '날짜 (YYYY-MM-DD)', '상태', '사유 (선택)', '처리상태', '처리시간', '파일링크 (선택)'];
  sheet.appendRow(headers);

  // 헤더 스타일
  sheet.getRange('A1:G1')
    .setFontWeight('bold')
    .setBackground('#FF9800')
    .setFontColor('white')
    .setHorizontalAlignment('center');

  // 열 너비 조정
  sheet.setColumnWidth(1, 100);  // 조원 이름
  sheet.setColumnWidth(2, 150);  // 날짜
  sheet.setColumnWidth(3, 100);  // 상태
  sheet.setColumnWidth(4, 200);  // 사유
  sheet.setColumnWidth(5, 100);  // 처리상태
  sheet.setColumnWidth(6, 150);  // 처리시간
  sheet.setColumnWidth(7, 300);  // 파일링크

  // 예시 데이터 3개 추가 (파일링크는 맨 끝)
  const examples = [
    ['센트룸', '2025-10-15', 'O', 'Drive 동기화 오류', '', '', ''],
    ['길', '2025-10-16', '출석', '업로드 지연', '', '', 'https://drive.google.com/...'],
    ['what', '2025-10-17', 'OFF', '병원 진료', '', '', '']
  ];

  examples.forEach(example => {
    sheet.appendRow(example);
  });

  // 예시 데이터는 연한 노란색
  sheet.getRange(2, 1, 3, 7).setBackground('#fff9c4');

  // 안내문 추가
  sheet.getRange('A5').setValue('📝 사용 방법:');
  sheet.getRange('A6').setValue('1. 위 예시를 참고하여 새 행에 정보 입력');
  sheet.getRange('A7').setValue('2. 상태는 "O", "OFF", "LONG_OFF", "X" 또는 "출석", "오프", "장기오프", "결석" 입력');
  sheet.getRange('A8').setValue('3. 파일링크는 선택사항 - 비워두면 ✓만, 입력하면 ✓📁 표시');
  sheet.getRange('A9').setValue('4. 트리거가 1시간마다 자동 처리하거나, "관리자수정_자동처리" 함수 직접 실행');
  sheet.getRange('A10').setValue('5. 처리 완료되면 "처리상태" 열에 완료 표시됨');
  sheet.getRange('A11').setValue('');
  sheet.getRange('A12').setValue('⚠️ 주의: 예시 데이터는 삭제하거나 "처리상태"를 "⏭️ 스킵"으로 변경하세요');

  sheet.getRange('A5:A12').setFontWeight('bold').setFontColor('#666666');
  
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

// ==================== Web App 배포 ====================

/**
 * Slack 인터랙티브 버튼 처리 (doPost)
 * - 장기오프 승인 버튼 클릭 시 처리
 */
function doPost(e) {
  Logger.log('========================================');
  Logger.log('doPost 함수 시작! (Slack Interactive)');
  Logger.log('호출 시각:', new Date());

  try {
    // Slack interactive payload 파싱
    const payload = JSON.parse(e.parameter.payload);
    Logger.log('Payload type:', payload.type);

    // Block Actions (버튼 클릭)
    if (payload.type === 'block_actions') {
      const action = payload.actions[0];
      const actionId = action.action_id;
      const responseUrl = payload.response_url;

      Logger.log('Action ID:', actionId);

      // 장기오프 승인 버튼
      if (actionId.startsWith('longoff_approve_')) {
        const data = JSON.parse(action.value);
        const name = data.name;
        const startDate = data.startDate;
        const endDate = data.endDate;
        const reason = data.reason || '장기오프';
        const formRow = data.row ? parseInt(data.row) : 0;

        Logger.log('장기오프 승인 처리:', name, startDate, '~', endDate);

        // 1) 장기오프신청 시트에서 승인 표시
        if (formRow > 0) {
          const ss = SpreadsheetApp.getActiveSpreadsheet();
          const longOffSheet = ss.getSheetByName(CONFIG.LONG_OFF_SHEET);
          if (longOffSheet) {
            longOffSheet.getRange(formRow, CONFIG.FORM_COLUMNS.APPROVED + 1).setValue('O');
          }
        }

        // 2) 해당 기간 모든 날짜에 장기오프 출석 기록
        let current = new Date(startDate);
        const end = new Date(endDate);
        current.setHours(0, 0, 0, 0);
        end.setHours(0, 0, 0, 0);
        let daysProcessed = 0;

        while (current <= end) {
          const dateStr = Utilities.formatDate(current, 'Asia/Seoul', 'yyyy-MM-dd');
          출석기록추가(name, dateStr, [], CONFIG.LONG_OFF_STATUS, reason);
          daysProcessed++;
          current.setDate(current.getDate() + 1);
        }

        Logger.log('✅', name, '장기오프', daysProcessed, '일 승인 완료');

        // 3) Slack 메시지 업데이트 (response_url 사용)
        if (responseUrl) {
          const updatedBlocks = [
            {
              type: 'header',
              text: { type: 'plain_text', text: '✅ 장기오프 승인 완료', emoji: true }
            },
            {
              type: 'section',
              fields: [
                { type: 'mrkdwn', text: '*신청자:*\n' + name },
                { type: 'mrkdwn', text: '*기간:*\n' + startDate + ' ~ ' + endDate + ' (' + daysProcessed + '일)' }
              ]
            },
            {
              type: 'context',
              elements: [
                { type: 'mrkdwn', text: '✅ 승인됨 | ' + new Date().toLocaleString('ko-KR') }
              ]
            }
          ];

          UrlFetchApp.fetch(responseUrl, {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify({
              replace_original: true,
              blocks: updatedBlocks
            }),
            muteHttpExceptions: true
          });
        }

        // 즉시 응답 (Slack 3초 타임아웃 대응)
        return ContentService.createTextOutput('');
      }

      // 장기오프 거절 버튼
      if (actionId.startsWith('longoff_reject_')) {
        const data = JSON.parse(action.value);
        const name = data.name;
        const startDate = data.startDate;
        const endDate = data.endDate;
        const responseUrl = payload.response_url;

        Logger.log('장기오프 거절:', name, startDate, '~', endDate);

        // Slack 메시지 업데이트
        if (responseUrl) {
          const updatedBlocks = [
            {
              type: 'header',
              text: { type: 'plain_text', text: '❌ 장기오프 거절됨', emoji: true }
            },
            {
              type: 'section',
              fields: [
                { type: 'mrkdwn', text: '*신청자:*\n' + name },
                { type: 'mrkdwn', text: '*기간:*\n' + startDate + ' ~ ' + endDate }
              ]
            },
            {
              type: 'context',
              elements: [
                { type: 'mrkdwn', text: '❌ 거절됨 | ' + new Date().toLocaleString('ko-KR') }
              ]
            }
          ];

          UrlFetchApp.fetch(responseUrl, {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify({
              replace_original: true,
              blocks: updatedBlocks
            }),
            muteHttpExceptions: true
          });
        }

        return ContentService.createTextOutput('');
      }
    }

    return ContentService.createTextOutput('OK');

  } catch (error) {
    Logger.log('doPost 오류:', error.message);
    return ContentService.createTextOutput('Error: ' + error.message);
  }
}

/**
 * 통합 doGet 함수 - 모든 웹앱 기능 처리
 * - date 파라미터: 다이제스트 HTML 서빙
 * - month + type 파라미터: 출석/주간 JSON 반환
 * - action=getDigest: 다이제스트 JSON 반환
 */
function doGet(e) {
  // ⚠️ 진단용 로그 - 이 로그가 보이지 않으면 doGet 자체가 실행되지 않는 것
  Logger.log('========================================');
  Logger.log('doGet 함수 시작!');
  Logger.log('호출 시각:', new Date());
  Logger.log('파라미터:', JSON.stringify(e.parameter));
  Logger.log('========================================');

  try {
    const params = e.parameter;

    // 0. 벌칙 완료 처리 (Slack 버튼 → 브라우저)
    if (params.action === 'penaltyComplete') {
      var name = params.name;
      var yearMonth = params.yearMonth;
      Logger.log('벌칙 완료 처리: ' + name + ' (' + yearMonth + ')');
      var success = 벌칙완료처리(name, yearMonth);

      // Slack 알림
      if (success) {
        var webhookUrl = getSlackWebhookUrl();
        if (webhookUrl) {
          UrlFetchApp.fetch(webhookUrl, {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify({
              text: '✅ *' + name + '* (' + yearMonth + ') 벌칙 수행 완료 처리됨'
            }),
            muteHttpExceptions: true
          });
        }
      }

      var resultMsg = success ? '✅ ' + name + '님 벌칙 수행 완료 처리되었습니다.' : '❌ 처리 실패: 해당 기록을 찾을 수 없습니다.';
      var html = '<html><head><meta charset="utf-8"><style>'
        + 'body{font-family:sans-serif;display:flex;justify-content:center;align-items:center;height:100vh;margin:0;background:#f5f5f5}'
        + '.card{background:white;padding:40px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,0.1);text-align:center;max-width:400px}'
        + '.ok{color:#4CAF50;font-size:48px}.fail{color:#f44336;font-size:48px}'
        + '</style></head><body><div class="card">'
        + '<div class="' + (success ? 'ok' : 'fail') + '">' + (success ? '✅' : '❌') + '</div>'
        + '<h2>' + resultMsg + '</h2>'
        + '<p style="color:#888">' + yearMonth + '</p>'
        + '<p style="color:#aaa;font-size:12px">이 탭을 닫아도 됩니다.</p>'
        + '</div></body></html>';
      return HtmlService.createHtmlOutput(html).setTitle('벌칙 완료');
    }

    // 0-1. 장기오프 승인 처리 (Slack 버튼 → 브라우저)
    if (params.action === 'longOffApprove') {
      var name = params.name;
      var startDate = params.startDate;
      var endDate = params.endDate;
      var reason = params.reason || '장기오프';
      var formRow = params.row ? parseInt(params.row) : 0;
      Logger.log('장기오프 승인: ' + name + ' (' + startDate + ' ~ ' + endDate + ')');

      var success = false;
      try {
        // 1) 장기오프신청 시트에서 승인 처리
        if (formRow > 0) {
          var ss = SpreadsheetApp.getActiveSpreadsheet();
          var longOffSheet = ss.getSheetByName(CONFIG.LONG_OFF_SHEET);
          if (longOffSheet) {
            longOffSheet.getRange(formRow, CONFIG.FORM_COLUMNS.APPROVED + 1).setValue('O');
          }
        }

        // 2) 해당 기간 모든 날짜에 장기오프 출석 기록
        var current = new Date(startDate);
        var end = new Date(endDate);
        current.setHours(0, 0, 0, 0);
        end.setHours(0, 0, 0, 0);
        var daysProcessed = 0;

        while (current <= end) {
          var dateStr = Utilities.formatDate(current, 'Asia/Seoul', 'yyyy-MM-dd');
          출석기록추가(name, dateStr, [], CONFIG.LONG_OFF_STATUS, reason);
          daysProcessed++;
          current.setDate(current.getDate() + 1);
        }

        // 3) JSON 재생성
        JSON파일생성();
        success = daysProcessed > 0;
        Logger.log('✅ ' + name + ' 장기오프 ' + daysProcessed + '일 승인 완료');

        // 4) Slack 승인 완료 알림
        var webhookUrl = getSlackWebhookUrl();
        if (webhookUrl) {
          UrlFetchApp.fetch(webhookUrl, {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify({
              text: '✅ *' + name + '* 장기오프 승인 완료 (' + startDate + ' ~ ' + endDate + ', ' + daysProcessed + '일)'
            }),
            muteHttpExceptions: true
          });
        }
      } catch (err) {
        Logger.log('❌ 장기오프 승인 오류: ' + err.message);
      }

      var resultMsg = success
        ? '✅ ' + name + '님 장기오프가 승인되었습니다. (' + startDate + ' ~ ' + endDate + ')'
        : '❌ 승인 처리에 실패했습니다.';
      var html = '<html><head><meta charset="utf-8"><style>'
        + 'body{font-family:sans-serif;display:flex;justify-content:center;align-items:center;height:100vh;margin:0;background:#f5f5f5}'
        + '.card{background:white;padding:40px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,0.1);text-align:center;max-width:400px}'
        + '.ok{color:#4CAF50;font-size:48px}.fail{color:#f44336;font-size:48px}'
        + '</style></head><body><div class="card">'
        + '<div class="' + (success ? 'ok' : 'fail') + '">' + (success ? '✅' : '❌') + '</div>'
        + '<h2>' + resultMsg + '</h2>'
        + '<p style="color:#888">' + startDate + ' ~ ' + endDate + '</p>'
        + '<p style="color:#aaa;font-size:12px">이 탭을 닫아도 됩니다.</p>'
        + '</div></body></html>';
      return HtmlService.createHtmlOutput(html).setTitle('장기오프 승인');
    }

    // 1. 다이제스트 HTML 서빙 (date 파라미터)
    if (params.date) {
      Logger.log('다이제스트 HTML 서빙 시작. 날짜:', params.date);
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
            message: '주간 통계 파일이 없습니다. 이번주주간집계() 함수를 실행해주세요.'
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
    Logger.log(`📖 다이제스트 읽기 시작: ${dateStr}`);

    // 1. 스프레드시트 시트에서 파일 ID 찾기
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.DIGEST_SHEET);

    if (!sheet) {
      Logger.log('⚠️ 다이제스트 시트가 없습니다.');
      return null;
    }

    // 모든 데이터 가져오기
    const data = sheet.getDataRange().getValues();

    // 날짜로 검색 (헤더 제외)
    let fileId = null;
    for (let i = 1; i < data.length; i++) {
      const cellValue = data[i][0];
      let cellDateStr;

      // Date 객체인 경우 문자열로 변환
      if (cellValue instanceof Date) {
        cellDateStr = Utilities.formatDate(cellValue, 'Asia/Seoul', 'yyyy-MM-dd');
      } else {
        cellDateStr = String(cellValue).trim();
      }

      Logger.log(`  비교: 셀="${cellDateStr}" vs 요청="${dateStr}"`);

      if (cellDateStr === dateStr) {
        fileId = data[i][1];  // 파일ID 컬럼
        Logger.log(`✅ 다이제스트 파일 ID 찾음: ${fileId}`);
        break;
      }
    }

    if (!fileId) {
      Logger.log(`❌ ${dateStr} 다이제스트 없음`);
      return null;
    }

    // 2. 드라이브에서 HTML 파일 읽기
    const file = DriveApp.getFileById(fileId);
    const htmlContent = file.getBlob().getDataAsString('UTF-8');

    Logger.log(`✅ HTML 파일 읽기 완료: ${htmlContent.length} 문자`);
    return htmlContent;

  } catch (error) {
    Logger.log(`HTML 읽기 실패: ${error.message}`);
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
 * 수정된 월의 JSON을 자동으로 재생성
 */
function 관리자수정처리() {
  Logger.log('');
  Logger.log('=== 관리자수정 처리 시작 ===');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adminSheet = ss.getSheetByName(CONFIG.ADMIN_SHEET);

  if (!adminSheet) {
    Logger.log(`⚠️ "${CONFIG.ADMIN_SHEET}" 시트가 없습니다. 건너뜀.`);
    return { processedCount: 0, affectedMonths: [] };
  }

  const data = adminSheet.getDataRange().getValues();

  if (data.length <= 1) {
    Logger.log('처리할 항목이 없습니다.');
    return { processedCount: 0, affectedMonths: [] };
  }

  let processedCount = 0;
  const now = new Date();
  const affectedMonths = new Set(); // 수정된 월 추적

  // 첫 행(헤더) 제외하고 처리
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[CONFIG.ADMIN_COLUMNS.NAME];
    const dateValue = row[CONFIG.ADMIN_COLUMNS.DATE];
    const status = row[CONFIG.ADMIN_COLUMNS.STATUS];
    const reason = row[CONFIG.ADMIN_COLUMNS.REASON] || '';
    const fileLink = row[CONFIG.ADMIN_COLUMNS.FILE_LINK] || '';  // 🆕 파일링크 (선택)
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
    let dateObj;
    try {
      if (dateValue instanceof Date) {
        dateObj = dateValue;
        dateStr = Utilities.formatDate(dateValue, 'Asia/Seoul', 'yyyy-MM-dd');
      } else {
        dateStr = String(dateValue).trim();
        if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
          Logger.log(`  ❌ ${i + 1}행: 날짜 형식 오류 (${dateStr}). YYYY-MM-DD 형식 사용 필요`);
          continue;
        }
        const parts = dateStr.split('-');
        dateObj = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
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
      // fileLink를 문자열로 변환 (빈 값이거나 다른 타입일 수 있음)
      const fileLinkStr = fileLink ? String(fileLink).trim() : '';
      const linkInfo = fileLinkStr ? ` [링크: ${fileLinkStr.substring(0, 30)}...]` : '';
      Logger.log(`  🔧 ${name} - ${dateStr} → ${normalizedStatus}${reason ? ' (' + reason + ')' : ''}${linkInfo}`);
      출석기록추가(name, dateStr, [], normalizedStatus, reason, '', fileLinkStr);

      // 수정된 월 추적 (yyyy-MM 형식)
      const yearMonth = Utilities.formatDate(dateObj, 'Asia/Seoul', 'yyyy-MM');
      affectedMonths.add(yearMonth);

      // 처리완료 표시
      const rowIndex = i + 1;
      const formattedTime = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
      adminSheet.getRange(rowIndex, CONFIG.ADMIN_COLUMNS.PROCESSED + 1).setValue('완료');
      adminSheet.getRange(rowIndex, CONFIG.ADMIN_COLUMNS.PROCESSED_TIME + 1).setValue(formattedTime);

      processedCount++;
    } catch (e) {
      Logger.log(`  ❌ ${i + 1}행: 처리 중 오류 - ${e.message}`);
    }
  }

  Logger.log(`✅ 관리자수정 처리 완료: ${processedCount}건`);

  // 수정된 월의 JSON 재생성 (최근 2개월만 처리하여 성능 최적화)
  const affectedMonthsArray = Array.from(affectedMonths);
  if (affectedMonthsArray.length > 0) {
    Logger.log('');
    Logger.log('=== 수정된 월 JSON 재생성 ===');

    // 현재 날짜 기준으로 최근 2개월만 재생성 (현재월 + 이전월)
    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth() + 1; // 1-based

    // 재생성 대상 월 필터링
    const recentMonths = affectedMonthsArray.filter(yearMonth => {
      const [year, month] = yearMonth.split('-').map(Number);
      // 현재월과 이전월만 재생성 (최대 2개월 범위)
      const monthDiff = (currentYear - year) * 12 + (currentMonth - month);
      return monthDiff >= 0 && monthDiff <= 1;
    });

    // 오래된 월은 스킵 로그
    const skippedMonths = affectedMonthsArray.filter(m => !recentMonths.includes(m));
    if (skippedMonths.length > 0) {
      Logger.log(`⏭️ 오래된 월 스킵 (성능 최적화): ${skippedMonths.join(', ')}`);
    }

    for (const yearMonth of recentMonths) {
      const [year, month] = yearMonth.split('-').map(Number);
      Logger.log(`📁 ${year}년 ${month}월 JSON 재생성 중...`);
      try {
        특정월JSON생성(year, month);
        Logger.log(`  ✅ ${year}년 ${month}월 일간 JSON 재생성 완료`);
      } catch (e) {
        Logger.log(`  ❌ ${year}년 ${month}월 일간 JSON 재생성 실패: ${e.message}`);
      }

      // 주간집계는 매일 새벽 이번주주간집계 트리거에서 처리 (성능 최적화)
    }
  }

  Logger.log('');

  return { processedCount, affectedMonths: affectedMonthsArray };
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
 * 이제 관리자수정처리()가 자동으로 해당 월 JSON을 재생성합니다.
 */
function 관리자수정만_처리() {
  const result = 관리자수정처리();
  Logger.log('');
  Logger.log('========================================');
  Logger.log(`✅ 관리자수정 처리 완료: ${result.processedCount}건`);
  if (result.affectedMonths.length > 0) {
    Logger.log(`📁 JSON 재생성된 월: ${result.affectedMonths.join(', ')}`);
  }
  Logger.log('========================================');
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

  // 🆕 주간집계에서 벌칙 판정용 주간결석 합계 미리 계산
  const weeklySheet = ss.getSheetByName('주간집계');
  const memberWeeklyAbsences = {};
  if (weeklySheet && weeklySheet.getLastRow() > 1) {
    const weeklyData = weeklySheet.getRange(2, 1, weeklySheet.getLastRow() - 1, 9).getDisplayValues();
    for (const row of weeklyData) {
      const rowYM = String(row[0]).trim();
      if (rowYM !== yearMonth) continue;
      const name = String(row[1]).trim();
      const absences = Number(row[6]) || 0;
      const isFullLongOff = String(row[8]).trim() === '전체장기오프';
      if (!memberWeeklyAbsences[name]) memberWeeklyAbsences[name] = 0;
      if (!isFullLongOff) memberWeeklyAbsences[name] += absences;
    }
    Logger.log(`주간집계 기준 결석: ${JSON.stringify(memberWeeklyAbsences)}`);
  } else {
    Logger.log('⚠️ 주간집계 데이터 없음 - 일일 결석 기준으로 상태 판정');
  }
  
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

    // 상태 판정 (주간집계 기준, 없으면 일일 기준)
    const weeklyAbs = memberWeeklyAbsences[memberName] ?? 결석;
    let 상태 = '정상';
    if (weeklyAbs >= 4) {
      상태 = '🚨 벌칙';
    } else if (weeklyAbs === 3) {
      상태 = '⚠️ 경고';
    } else {
      상태 = '✅ 정상';
    }

    let 비고 = `출석 ${출석}일 + 오프 ${오프}일 + 장기오프 ${장기오프}일 + 결석 ${결석}일 = 총 ${daysInMonth}일 (주간결석 ${weeklyAbs}회)`;
    if (오프초과결석목록.length > 0) {
      비고 += ` | 🚨 오프 초과 결석: ${오프초과결석목록.join(', ')}`;
    }

    Logger.log(`${memberName}: 출석 ${출석}, 오프 ${오프}, 장기오프 ${장기오프}, 결석 ${결석}, 주간결석 ${weeklyAbs} → ${상태}`);
    
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

  // 전월 주간집계 재생성
  try {
    const [pYear, pMonth] = yearMonth.split('-').map(Number);
    Logger.log('🔄 전월 주간집계 재생성 중... (' + yearMonth + ')');
    var 집계결과 = 월별주간집계(pYear, pMonth - 1);  // 0-based month
    주간집계저장(pYear, pMonth - 1, 집계결과);
    SpreadsheetApp.flush();  // 시트 쓰기 즉시 반영
    Logger.log('✅ 전월 주간집계 재생성 완료');
  } catch (e) {
    Logger.log('⚠️ 주간집계 재생성 오류 (결산은 정상 완료): ' + e.message);
  }

  // 벌칙 대상자 등록 및 슬랙 알림 (주간집계 재생성 실패해도 독립 실행)
  try {
    벌칙대상자_등록(yearMonth);
    벌칙대상자_슬랙알림(yearMonth);
  } catch (e) {
    Logger.log('⚠️ 벌칙 알림 오류: ' + e.message);
  }
}

/**
 * 🆕 전월 결산 수동 실행 (에디터에서 바로 실행 가능)
 * 벌칙 대상자 등록 및 Slack 알림도 함께 실행됨
 */
function 전월결산_수동실행() {
  const now = new Date();
  const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  특정월_결산생성(lastMonth.getFullYear(), lastMonth.getMonth() + 1);
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

  // 🆕 주간집계에서 벌칙 판정용 주간결석 합계 미리 계산
  const weeklySheet = ss.getSheetByName('주간집계');
  const memberWeeklyAbsences = {};
  if (weeklySheet && weeklySheet.getLastRow() > 1) {
    const weeklyData = weeklySheet.getRange(2, 1, weeklySheet.getLastRow() - 1, 9).getDisplayValues();
    for (const row of weeklyData) {
      const rowYM = String(row[0]).trim();
      if (rowYM !== yearMonth) continue;
      const name = String(row[1]).trim();
      const absences = Number(row[6]) || 0;
      const isFullLongOff = String(row[8]).trim() === '전체장기오프';
      if (!memberWeeklyAbsences[name]) memberWeeklyAbsences[name] = 0;
      if (!isFullLongOff) memberWeeklyAbsences[name] += absences;
    }
    Logger.log(`주간집계 기준 결석: ${JSON.stringify(memberWeeklyAbsences)}`);
  } else {
    Logger.log('⚠️ 주간집계 데이터 없음 - 일일 결석 기준으로 상태 판정');
  }

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

    // 상태 판정 (주간집계 기준, 없으면 일일 기준)
    const weeklyAbs = memberWeeklyAbsences[memberName] ?? 결석;
    let 상태 = '정상';
    if (weeklyAbs >= 4) {
      상태 = '🚨 벌칙';
    } else if (weeklyAbs === 3) {
      상태 = '⚠️ 경고';
    } else {
      상태 = '✅ 정상';
    }

    let 비고 = `출석 ${출석}일 + 오프 ${오프}일 + 장기오프 ${장기오프}일 + 결석 ${결석}일 = 총 ${daysInMonth}일 (주간결석 ${weeklyAbs}회)`;
    if (오프초과결석목록.length > 0) {
      비고 += ` | 🚨 오프 초과 결석: ${오프초과결석목록.join(', ')}`;
    }

    Logger.log(`${memberName}: 출석 ${출석}, 오프 ${오프}, 장기오프 ${장기오프}, 결석 ${결석}, 주간결석 ${weeklyAbs} → ${상태}`);
    
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

  // 주간집계 생성 후 벌칙 대상자 등록 및 슬랙 알림
  try {
    Logger.log('🔄 주간집계 생성 중... (' + yearMonth + ')');
    var 집계결과 = 월별주간집계(year, month - 1);  // 0-based month
    주간집계저장(year, month - 1, 집계결과);
    SpreadsheetApp.flush();  // 시트 쓰기 즉시 반영
    Logger.log('✅ 주간집계 생성 완료');
  } catch (e) {
    Logger.log('⚠️ 주간집계 생성 오류: ' + e.message);
  }

  try {
    벌칙대상자_등록(yearMonth);
    벌칙대상자_슬랙알림(yearMonth);
  } catch (e) {
    Logger.log('⚠️ 벌칙 알림 오류: ' + e.message);
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
 * 장기오프 신청 → Slack 승인 요청 전송 (승인 전까지 출석 반영 안 함)
 */
function onFormSubmit_장기오프처리(e) {
  try {
    Logger.log('=== 폼 제출 감지: 장기오프 승인 요청 ===');

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

    const startDateStr = Utilities.formatDate(startDate, 'Asia/Seoul', 'yyyy-MM-dd');
    const endDateStr = Utilities.formatDate(endDate, 'Asia/Seoul', 'yyyy-MM-dd');

    // 일수 계산
    const days = Math.round((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;

    // Slack 승인 요청 전송 (인터랙티브 버튼)
    const webhookUrl = getSlackWebhookUrl();
    if (webhookUrl) {
      // 버튼 클릭 시 전달할 데이터
      const buttonValue = JSON.stringify({
        name: name,
        startDate: startDateStr,
        endDate: endDateStr,
        reason: reason || '장기오프',
        row: row
      });

      const blocks = [
        {
          type: 'header',
          text: { type: 'plain_text', text: '🏝️ 장기오프 승인 요청', emoji: true }
        },
        {
          type: 'section',
          fields: [
            { type: 'mrkdwn', text: '*신청자:*\n' + name },
            { type: 'mrkdwn', text: '*기간:*\n' + startDateStr + ' ~ ' + endDateStr + ' (' + days + '일)' }
          ]
        },
        {
          type: 'section',
          text: { type: 'mrkdwn', text: '*사유:*\n' + (reason || '(사유 없음)') }
        },
        { type: 'divider' },
        {
          type: 'section',
          text: { type: 'mrkdwn', text: '승인하면 해당 기간이 장기오프로 출석에 반영됩니다.' }
        },
        {
          type: 'actions',
          elements: [
            {
              type: 'button',
              text: { type: 'plain_text', text: '✅ 승인', emoji: true },
              style: 'primary',
              action_id: 'longoff_approve_' + row,
              value: buttonValue
            },
            {
              type: 'button',
              text: { type: 'plain_text', text: '❌ 거절', emoji: true },
              style: 'danger',
              action_id: 'longoff_reject_' + row,
              value: buttonValue
            }
          ]
        }
      ];

      UrlFetchApp.fetch(webhookUrl, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          text: '🏝️ ' + name + ' 장기오프 승인 요청 (' + startDateStr + ' ~ ' + endDateStr + ')',
          blocks: blocks
        }),
        muteHttpExceptions: true
      });

      Logger.log(`✅ ${name} 장기오프 승인 요청 Slack 전송 완료 (인터랙티브 버튼)`);
    } else {
      Logger.log('⚠️ Slack Webhook URL 미설정 - 승인 요청 전송 불가');
    }

  } catch (error) {
    Logger.log('❌ 오류: ' + error.toString());
  }
}

// ==================== 주간 출석 집계 시스템 ====================
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

function 주간인증계산(memberName, 주시작, 주끝, 완료된주 = true, cachedData = null) {
  let 인증횟수 = 0;
  let 장기오프일수 = 0;

  // 캐시된 데이터 추출 (인덱스 또는 원본 데이터)
  const cachedAttendanceIndex = cachedData?.출석인덱스 || null;
  const cachedLongOffData = cachedData?.장기오프데이터 || null;

  // 주의 각 날짜 체크
  for (let d = new Date(주시작); d <= 주끝; d.setDate(d.getDate() + 1)) {
    const dateStr = Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd');

    // 장기오프 확인 (최우선)
    const longOffInfo = 장기오프확인(memberName, dateStr, cachedLongOffData);
    if (longOffInfo.isLongOff) {
      장기오프일수++;
      continue;
    }

    // 출석 확인 (인덱스로 O(1) 검색)
    const 출석여부 = 출석확인(memberName, dateStr, cachedAttendanceIndex);
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
 * 출석 데이터를 인덱싱하여 O(1) 검색 가능하게 만듦
 * @param {Array} data - 시트 원본 데이터
 * @returns {Map} key: "이름-날짜", value: status
 */
function 출석데이터인덱싱(data) {
  const index = new Map();

  for (let i = 1; i < data.length; i++) {
    const [timestamp, name, recordDate, fileCount, links, folderLink, status, weekNum, reason] = data[i];

    if (!name) continue;

    // 날짜 문자열 변환 (한 번만)
    const dateStr = typeof recordDate === 'string'
      ? recordDate
      : Utilities.formatDate(new Date(recordDate), 'Asia/Seoul', 'yyyy-MM-dd');

    const key = `${name}-${dateStr}`;

    // 이미 있으면 O가 우선 (출석이 한 번이라도 있으면 출석으로 인정)
    if (!index.has(key) || status === 'O') {
      index.set(key, status);
    }
  }

  return index;
}

function 출석확인(memberName, dateStr, cachedIndex = null) {
  // 인덱스가 있으면 O(1) 검색
  if (cachedIndex instanceof Map) {
    const key = `${memberName}-${dateStr}`;
    const status = cachedIndex.get(key);
    return status === 'O';
  }

  // 인덱스가 없으면 기존 방식 (시트에서 로드)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) return false;

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const [timestamp, name, recordDate, fileCount, links, folderLink, status, weekNum, reason] = data[i];

    const recordDateStr = typeof recordDate === 'string'
      ? recordDate
      : Utilities.formatDate(new Date(recordDate), 'Asia/Seoul', 'yyyy-MM-dd');

    if (name === memberName && recordDateStr === dateStr) {
      return status === 'O';
    }
  }

  return false;
}

function 월별주간집계(year, month) {
  Logger.log(`=== ${year}년 ${month + 1}월 주간 집계 시작 ===`);

  const 주목록 = 월별주목록가져오기(year, month);
  const 조원결석 = {};

  Logger.log(`총 ${주목록.length}개 주 발견`);

  // 현재 날짜 (한국 시간)
  const 오늘 = new Date();
  const 오늘자정 = new Date(오늘.getFullYear(), 오늘.getMonth(), 오늘.getDate());

  // 🚀 시트 데이터 캐싱 (한 번만 로드)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 출석원본 = ss.getSheetByName(CONFIG.SHEET_NAME)?.getDataRange().getValues() || [];
  const 장기오프원본 = ss.getSheetByName(CONFIG.LONG_OFF_SHEET)?.getDataRange().getValues() || [];

  // 🚀 데이터 인덱싱 (O(1) 검색을 위해)
  const cachedData = {
    출석인덱스: 출석데이터인덱싱(출석원본),
    장기오프데이터: 장기오프원본
  };
  Logger.log(`📊 데이터 캐싱 완료 (출석: ${출석원본.length}행, 장기오프: ${장기오프원본.length}행)`);

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
      const 결과 = 주간인증계산(memberName, 주.시작, 주.끝, 완료된주, cachedData);

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
  const targetYear = year;
  const targetMonth = month + 1;

  // 기존 데이터 삭제 (해당 년월 - 숫자 비교로 "2026-2"와 "2026-02" 모두 매칭)
  const data = sheet.getDataRange().getDisplayValues();
  for (let i = data.length - 1; i >= 1; i--) {
    const rowYM = String(data[i][0]).trim();
    const parts = rowYM.split('-');
    if (parts.length === 2 && parseInt(parts[0]) === targetYear && parseInt(parts[1]) === targetMonth) {
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
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, 1, rows.length, 9).setValues(rows);
  }

  Logger.log(`주간집계 시트에 ${rows.length}개 행 저장`);
}

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
 * 🆕 지난주 결석 확정 (매주 월요일 트리거용)
 * - 지난주를 "완료" 상태로 변경하고 결석 확정
 * - 이미 완료된 경우 스킵
 */
function 지난주결석확정() {
  const startTime = new Date();
  Logger.log('=== 지난주 결석 확정 시작 ===');

  const now = new Date();

  // 이번 주 월요일 찾기
  const 이번주월요일 = new Date(now);
  const dayOfWeek = now.getDay();
  const daysFromMonday = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
  이번주월요일.setDate(now.getDate() - daysFromMonday);
  이번주월요일.setHours(0, 0, 0, 0);

  // 지난 주 월요일/일요일
  const 지난주월요일 = new Date(이번주월요일);
  지난주월요일.setDate(지난주월요일.getDate() - 7);
  const 지난주일요일 = new Date(지난주월요일);
  지난주일요일.setDate(지난주월요일.getDate() + 6);

  const 오늘자정 = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  // 지난주가 완료되었는지 확인
  if (지난주일요일 >= 오늘자정) {
    Logger.log('⚠️ 지난주가 아직 끝나지 않았습니다.');
    return;
  }

  // 이미 완료 처리됐는지 확인
  const 지난주이미완료 = 주차완료여부확인(지난주월요일);
  if (지난주이미완료) {
    Logger.log('📌 지난주는 이미 완료 처리됨 (스킵)');
    const endTime = new Date();
    Logger.log(`=== 완료 (${((endTime - startTime) / 1000).toFixed(1)}초) ===`);
    return;
  }

  // 🚀 시트 데이터 캐싱 (한 번만 로드)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 출석원본 = ss.getSheetByName(CONFIG.SHEET_NAME)?.getDataRange().getValues() || [];
  const 장기오프원본 = ss.getSheetByName(CONFIG.LONG_OFF_SHEET)?.getDataRange().getValues() || [];

  // 🚀 데이터 인덱싱 (O(1) 검색을 위해)
  const cachedData = {
    출석인덱스: 출석데이터인덱싱(출석원본),
    장기오프데이터: 장기오프원본
  };
  Logger.log(`📊 데이터 캐싱 완료 (출석: ${출석원본.length}행, 장기오프: ${장기오프원본.length}행)`);

  // 지난주 완료 처리
  지난주완료처리(지난주월요일, 지난주일요일, cachedData);

  const endTime = new Date();
  const 소요시간 = (endTime - startTime) / 1000;
  Logger.log(`\n=== 지난주 결석 확정 완료 (${소요시간.toFixed(1)}초) ===`);
}

/**
 * 🆕 이번 주만 빠르게 집계 (매일 트리거용)
 * - 월요일 기준으로 주를 판단
 * - 월초에 월요일이 없는 날들은 이전달 마지막 주로 처리
 * - 예: 2026년 1월 1~4일(목~일) → 2025년 12월 마지막 주
 */
function 이번주주간집계() {
  const startTime = new Date();
  Logger.log('=== 이번 주 주간집계 시작 ===');

  const now = new Date();

  // 이번 주 월요일 찾기
  const 이번주월요일 = new Date(now);
  const dayOfWeek = now.getDay(); // 0=일, 1=월, ..., 6=토
  const daysFromMonday = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
  이번주월요일.setDate(now.getDate() - daysFromMonday);
  이번주월요일.setHours(0, 0, 0, 0);

  // 이번 주 일요일
  const 이번주일요일 = new Date(이번주월요일);
  이번주일요일.setDate(이번주월요일.getDate() + 6);

  const 오늘자정 = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  // 이번 주 처리
  Logger.log('📌 이번 주 집계');

  // 월요일이 속한 월이 이 주의 소속 월
  const 소속년도 = 이번주월요일.getFullYear();
  const 소속월 = 이번주월요일.getMonth(); // 0-based

  Logger.log(`오늘: ${Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd (E)')}`);
  Logger.log(`이번 주: ${Utilities.formatDate(이번주월요일, 'Asia/Seoul', 'MM/dd(E)')} ~ ${Utilities.formatDate(이번주일요일, 'Asia/Seoul', 'MM/dd(E)')}`);
  Logger.log(`소속: ${소속년도}년 ${소속월 + 1}월`);

  // 이 주가 소속월의 몇 번째 주인지 계산
  const 주목록 = 월별주목록가져오기(소속년도, 소속월);
  let 현재주차 = -1;

  for (let i = 0; i < 주목록.length; i++) {
    const 주 = 주목록[i];
    if (주.시작.getTime() === 이번주월요일.getTime()) {
      현재주차 = i + 1;
      break;
    }
  }

  if (현재주차 === -1) {
    Logger.log('⚠️ 현재 주차를 찾을 수 없습니다.');
    return;
  }

  Logger.log(`→ ${소속월 + 1}월 ${현재주차}주차`);

  // 이번 주가 완료되었는지 확인
  const 완료된주 = 이번주일요일 < 오늘자정;

  // 🚀 시트 데이터 캐싱 (한 번만 로드)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 출석원본 = ss.getSheetByName(CONFIG.SHEET_NAME)?.getDataRange().getValues() || [];
  const 장기오프원본 = ss.getSheetByName(CONFIG.LONG_OFF_SHEET)?.getDataRange().getValues() || [];

  // 🚀 데이터 인덱싱 (O(1) 검색을 위해)
  const cachedData = {
    출석인덱스: 출석데이터인덱싱(출석원본),
    장기오프데이터: 장기오프원본
  };
  Logger.log(`📊 데이터 캐싱 완료 (출석: ${출석원본.length}행, 장기오프: ${장기오프원본.length}행)`);

  // 각 조원별 이번 주 집계
  const 이번주집계 = {};

  for (const memberName of Object.keys(CONFIG.MEMBERS)) {
    const 결과 = 주간인증계산(memberName, 이번주월요일, 이번주일요일, 완료된주, cachedData);

    이번주집계[memberName] = {
      주차: 현재주차,
      인증횟수: 결과.인증횟수,
      장기오프일수: 결과.장기오프일수,
      필요횟수: 결과.필요횟수,
      결석: 결과.결석,
      전체장기오프: 결과.전체장기오프,
      주완료: 결과.주완료
    };

    const 상태표시 = 결과.주완료 ? `→ 결석 ${결과.결석}회` : '(진행중)';
    Logger.log(`  ${memberName}: 인증 ${결과.인증횟수}/${결과.필요횟수}회 ${상태표시}`);
  }

  // 기존 JSON 파일 읽어서 이번 주만 업데이트
  이번주JSON업데이트(소속년도, 소속월, 현재주차, 이번주집계, 주목록);

  const endTime = new Date();
  const 소요시간 = (endTime - startTime) / 1000;
  Logger.log(`\n=== 이번 주 주간집계 완료 (${소요시간.toFixed(1)}초) ===`);
}

/**
 * 🆕 특정 주 주간집계 재계산 (범용)
 * - 해당 주가 이미 완료되어도 강제로 재계산
 * - 시트 메뉴 또는 수동 실행용
 * @param {string} dateStr - 해당 주의 아무 날짜 (yyyy-MM-dd 형식). 없으면 프롬프트로 입력받음
 */
function 주간집계재계산(dateStr) {
  // 날짜가 없으면 프롬프트로 입력받기
  if (!dateStr) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      '📅 주간집계 재계산',
      '재계산할 주의 날짜를 입력하세요 (예: 2025-12-15)\n해당 주의 아무 날짜나 입력하면 됩니다.',
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) {
      Logger.log('취소됨');
      return;
    }

    dateStr = response.getResponseText().trim();
  }

  // 날짜 파싱
  const targetDate = new Date(dateStr);
  if (isNaN(targetDate.getTime())) {
    Logger.log(`❌ 잘못된 날짜 형식: ${dateStr}`);
    SpreadsheetApp.getUi().alert(`잘못된 날짜 형식입니다: ${dateStr}\n예: 2025-12-15`);
    return;
  }

  // 해당 주의 월요일 찾기
  const dayOfWeek = targetDate.getDay(); // 0=일, 1=월, ..., 6=토
  const daysFromMonday = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
  const 주월요일 = new Date(targetDate);
  주월요일.setDate(targetDate.getDate() - daysFromMonday);
  주월요일.setHours(0, 0, 0, 0);

  // 해당 주의 일요일
  const 주일요일 = new Date(주월요일);
  주일요일.setDate(주월요일.getDate() + 6);

  const 시작일 = Utilities.formatDate(주월요일, 'Asia/Seoul', 'MM/dd(E)');
  const 종료일 = Utilities.formatDate(주일요일, 'Asia/Seoul', 'MM/dd(E)');

  Logger.log(`=== 주간집계 재계산 시작 ===`);
  Logger.log(`📅 대상 주: ${시작일} ~ ${종료일}`);

  // 강제로 완료 처리 (재계산)
  지난주완료처리(주월요일, 주일요일);

  Logger.log(`✅ 주간집계 재계산 완료: ${시작일} ~ ${종료일}`);

  try {
    SpreadsheetApp.getUi().alert(`✅ 주간집계 재계산 완료!\n\n대상: ${시작일} ~ ${종료일}`);
  } catch (e) {
    // UI가 없는 환경 (트리거 등)에서는 무시
  }
}

/**
 * 지난주 완료 처리 (결석 확정)
 */
function 지난주완료처리(지난주월요일, 지난주일요일, cachedData = null) {
  const 소속년도 = 지난주월요일.getFullYear();
  const 소속월 = 지난주월요일.getMonth();

  Logger.log(`지난주: ${Utilities.formatDate(지난주월요일, 'Asia/Seoul', 'MM/dd(E)')} ~ ${Utilities.formatDate(지난주일요일, 'Asia/Seoul', 'MM/dd(E)')}`);
  Logger.log(`소속: ${소속년도}년 ${소속월 + 1}월`);

  const 주목록 = 월별주목록가져오기(소속년도, 소속월);
  let 지난주차 = -1;

  for (let i = 0; i < 주목록.length; i++) {
    const 주 = 주목록[i];
    if (주.시작.getTime() === 지난주월요일.getTime()) {
      지난주차 = i + 1;
      break;
    }
  }

  if (지난주차 === -1) {
    Logger.log('⚠️ 지난주 주차를 찾을 수 없습니다.');
    return;
  }

  Logger.log(`→ ${소속월 + 1}월 ${지난주차}주차 (완료 처리)`);

  // 각 조원별 지난주 집계 (완료 상태로)
  const 지난주집계 = {};

  for (const memberName of Object.keys(CONFIG.MEMBERS)) {
    const 결과 = 주간인증계산(memberName, 지난주월요일, 지난주일요일, true, cachedData);

    지난주집계[memberName] = {
      주차: 지난주차,
      인증횟수: 결과.인증횟수,
      장기오프일수: 결과.장기오프일수,
      필요횟수: 결과.필요횟수,
      결석: 결과.결석,
      전체장기오프: 결과.전체장기오프,
      주완료: true  // 강제로 완료
    };

    Logger.log(`  ${memberName}: 인증 ${결과.인증횟수}/${결과.필요횟수}회 → 결석 ${결과.결석}회`);
  }

  // JSON 업데이트
  이번주JSON업데이트(소속년도, 소속월, 지난주차, 지난주집계, 주목록);
}

/**
 * 주차가 이미 완료 처리되었는지 JSON에서 확인
 */
function 주차완료여부확인(주월요일) {
  const 소속년도 = 주월요일.getFullYear();
  const 소속월 = 주월요일.getMonth();
  const 년월 = `${소속년도}-${String(소속월 + 1).padStart(2, '0')}`;
  const fileName = `weekly_summary_${년월}.json`;

  try {
    const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);
    const files = folder.getFilesByName(fileName);

    if (!files.hasNext()) {
      return false;  // 파일이 없으면 완료 안 됨
    }

    const file = files.next();
    const content = file.getBlob().getDataAsString('UTF-8');
    const jsonData = JSON.parse(content);

    // 해당 주차 찾기
    const 주목록 = 월별주목록가져오기(소속년도, 소속월);
    let 주차 = -1;

    for (let i = 0; i < 주목록.length; i++) {
      if (주목록[i].시작.getTime() === 주월요일.getTime()) {
        주차 = i + 1;
        break;
      }
    }

    if (주차 === -1) return false;

    // 첫 번째 조원의 해당 주차 데이터 확인
    const 첫번째조원 = Object.keys(jsonData.조원별집계)[0];
    if (!첫번째조원) return false;

    const 주차별 = jsonData.조원별집계[첫번째조원].주차별;
    const 해당주차 = 주차별.find(w => w.주차 === 주차);

    if (!해당주차) return false;

    return 해당주차.상태 === '완료';

  } catch (e) {
    Logger.log(`주차완료여부 확인 오류: ${e.message}`);
    return false;
  }
}

/**
 * 이번 주 데이터만 JSON에 업데이트 (전체 재생성 없이)
 */
function 이번주JSON업데이트(year, month, 현재주차, 이번주집계, 주목록) {
  const 년월 = `${year}-${String(month + 1).padStart(2, '0')}`;
  const fileName = `weekly_summary_${년월}.json`;

  try {
    const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);
    const files = folder.getFilesByName(fileName);

    let jsonData;

    if (files.hasNext()) {
      // 기존 파일이 있으면 읽어서 업데이트
      const file = files.next();
      const content = file.getBlob().getDataAsString('UTF-8');
      jsonData = JSON.parse(content);
      Logger.log(`📂 기존 JSON 파일 로드: ${fileName}`);
    } else {
      // 없으면 새로 생성
      jsonData = {
        년월: 년월,
        생성일: new Date().toISOString(),
        총주차: 주목록.length,
        주차정보: 주목록.map((주, idx) => ({
          주차: idx + 1,
          시작: Utilities.formatDate(주.시작, 'Asia/Seoul', 'yyyy-MM-dd'),
          끝: Utilities.formatDate(주.끝, 'Asia/Seoul', 'yyyy-MM-dd')
        })),
        규칙설명: {
          주정의: '월요일~일요일',
          예시: '11월 25일(월)~12월 1일(일) → 11월 4주차'
        },
        조원별집계: {}
      };
      Logger.log(`📝 새 JSON 파일 생성: ${fileName}`);
    }

    // 각 조원별 이번 주차 데이터 업데이트
    for (const [memberName, weekData] of Object.entries(이번주집계)) {
      if (!jsonData.조원별집계[memberName]) {
        jsonData.조원별집계[memberName] = {
          총결석: 0,
          주차별: []
        };
      }

      const memberData = jsonData.조원별집계[memberName];

      // 해당 주차 데이터 찾아서 업데이트 또는 추가
      const weekIndex = memberData.주차별.findIndex(w => w.주차 === 현재주차);

      const newWeekData = {
        주차: 현재주차,
        인증: weekData.인증횟수,
        필요: weekData.필요횟수,
        장기오프: weekData.장기오프일수,
        결석: weekData.결석,
        상태: weekData.주완료 ? '완료' : '진행중',
        전체장기오프: weekData.전체장기오프
      };

      if (weekIndex >= 0) {
        memberData.주차별[weekIndex] = newWeekData;
      } else {
        memberData.주차별.push(newWeekData);
        // 주차 순서로 정렬
        memberData.주차별.sort((a, b) => a.주차 - b.주차);
      }

      // 총결석 재계산
      memberData.총결석 = memberData.주차별
        .filter(w => !w.전체장기오프)
        .reduce((sum, w) => sum + (w.결석 || 0), 0);
    }

    // 🆕 탈퇴 멤버 제거 (CONFIG.MEMBERS에 없는 멤버 삭제)
    const activeMembers = Object.keys(CONFIG.MEMBERS);
    const jsonMembers = Object.keys(jsonData.조원별집계);
    for (const memberName of jsonMembers) {
      if (!activeMembers.includes(memberName)) {
        delete jsonData.조원별집계[memberName];
        Logger.log(`🗑️ JSON에서 탈퇴 멤버 제거: ${memberName}`);
      }
    }

    // 업데이트 시간 기록
    jsonData.최종업데이트 = new Date().toISOString();

    // 파일 저장
    const jsonString = JSON.stringify(jsonData, null, 2);

    // 기존 파일 삭제 후 새로 생성
    const existingFiles = folder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }

    const newFile = folder.createFile(fileName, jsonString, MimeType.PLAIN_TEXT);
    newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    Logger.log(`✅ JSON 업데이트 완료: ${현재주차}주차`);

  } catch (e) {
    Logger.log(`❌ JSON 업데이트 실패: ${e.message}`);
  }
}

// ==================== AI 다이제스트 HTML 생성 시스템 ====================

/**
 * ==================== 스터디 컨텐츠 수집 시스템 ====================
 * 공부 내용 자동 수집 및 정리
 */

/**
 * 마크다운 내용 클린업
 * HTML 태그, Obsidian 특수 문법 등을 제거하여 깔끔하게 정리
 * @param {string} content - 원본 마크다운 내용
 * @returns {string} 정리된 내용
 */
function 마크다운클린업(content) {
  if (!content) return '';

  let cleaned = content;

  // 1. HTML 태그 제거 (스타일 속성 포함)
  cleaned = cleaned.replace(/<span[^>]*>/g, '');
  cleaned = cleaned.replace(/<\/span>/g, '');
  cleaned = cleaned.replace(/<font[^>]*>/g, '');
  cleaned = cleaned.replace(/<\/font>/g, '');
  cleaned = cleaned.replace(/<[^>]+>/g, ''); // 나머지 HTML 태그 제거

  // 2. Obsidian 이미지 링크 변환
  cleaned = cleaned.replace(/!\[\[([^\]]+)\]\]/g, '[이미지: $1]');

  // 3. 연속된 빈 줄 정리 (3개 이상의 빈 줄을 2개로)
  cleaned = cleaned.replace(/\n{4,}/g, '\n\n\n');

  // 4. 마크다운 강조 문법은 유지 (**, *, ~~, ` 등)

  return cleaned.trim();
}

/**
 * 마크다운을 HTML로 변환
 * @param {string} markdown - 마크다운 텍스트
 * @returns {string} HTML
 */
function 마크다운을HTML로(markdown) {
  if (!markdown) return '';

  let html = markdown;

  // 1. 코드 블록 먼저 처리 (변환 전에 보호)
  const codeBlocks = [];
  html = html.replace(/`([^`]+)`/g, function(match, code) {
    codeBlocks.push(code);
    return `__CODE_BLOCK_${codeBlocks.length - 1}__`;
  });

  // 2. 제목 변환 (### → h3, ## → h2, # → h1)
  html = html.replace(/^### (.+)$/gm, '<h3>$1</h3>');
  html = html.replace(/^## (.+)$/gm, '<h2>$1</h2>');
  html = html.replace(/^# (.+)$/gm, '<h1>$1</h1>');

  // 3. 굵게 **text** → <strong>text</strong>
  html = html.replace(/\*\*([^\*]+)\*\*/g, '<strong>$1</strong>');

  // 4. 기울임 *text* → <em>text</em>
  html = html.replace(/\*([^\*]+)\*/g, '<em>$1</em>');

  // 5. 취소선 ~~text~~ → <del>text</del>
  html = html.replace(/~~([^~]+)~~/g, '<del>$1</del>');

  // 6. 리스트 변환
  const lines = html.split('\n');
  let inList = false;
  let result = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // 들여쓰기된 리스트 (탭 또는 2칸 공백)
    if (/^\t[\-\*] (.+)$/.test(line) || /^  [\-\*] (.+)$/.test(line)) {
      const content = line.replace(/^\t[\-\*] /, '').replace(/^  [\-\*] /, '');
      if (!inList) {
        result.push('<ul>');
        inList = true;
      }
      result.push(`  <li style="margin-left: 20px;">${content}</li>`);
    }
    // 일반 리스트
    else if (/^[\-\*] (.+)$/.test(line)) {
      const content = line.replace(/^[\-\*] /, '');
      if (!inList) {
        result.push('<ul>');
        inList = true;
      }
      result.push(`  <li>${content}</li>`);
    }
    // 리스트가 아닌 줄
    else {
      if (inList) {
        result.push('</ul>');
        inList = false;
      }
      result.push(line);
    }
  }

  // 마지막 리스트 닫기
  if (inList) {
    result.push('</ul>');
  }

  html = result.join('\n');

  // 7. 문단 처리
  html = html.replace(/\n\n+/g, '</p><p>');
  html = html.replace(/\n/g, '<br>\n');

  // h 태그와 ul 태그 주변의 불필요한 <br> 제거
  html = html.replace(/<br>\s*<\/h([123])>/g, '</h$1>');
  html = html.replace(/<h([123])><br>/g, '<h$1>');
  html = html.replace(/<br>\s*<ul>/g, '<ul>');
  html = html.replace(/<\/ul><br>/g, '</ul>');

  // 전체를 <p>로 감싸기
  if (!html.startsWith('<')) {
    html = '<p>' + html;
  }
  if (!html.endsWith('>')) {
    html = html + '</p>';
  }

  // 빈 <p></p> 제거
  html = html.replace(/<p><\/p>/g, '');
  html = html.replace(/<p>\s*<\/p>/g, '');
  html = html.replace(/<p><br><\/p>/g, '');

  // 8. 코드 블록 복원
  codeBlocks.forEach((code, index) => {
    html = html.replace(
      `__CODE_BLOCK_${index}__`,
      `<code>${code.replace(/</g, '&lt;').replace(/>/g, '&gt;')}</code>`
    );
  });

  return html;
}


/**
 * 일일 다이제스트 생성
 * @param {string} dateStr - 날짜 (yyyy-MM-dd). 없으면 어제
 * @returns {string} 생성된 다이제스트
 */
function 일일AI다이제스트생성(dateStr) {
  // 트리거 호출 시 이벤트 객체가 전달되므로 문자열인지 확인
  if (!dateStr || typeof dateStr !== 'string') {
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    dateStr = Utilities.formatDate(yesterday, 'Asia/Seoul', 'yyyy-MM-dd');
  }

  Logger.log(`=== ${dateStr} 다이제스트 생성 시작 ===`);

  const 조원데이터 = [];

  for (const [memberName, folderIdOrArray] of Object.entries(CONFIG.MEMBERS)) {
    const folderIds = Array.isArray(folderIdOrArray) ? folderIdOrArray : [folderIdOrArray];

    Logger.log(`\n${memberName} 파일 찾는 중... (폴더 ID: ${folderIds.length}개)`);

    for (const folderId of folderIds) {
      const content = 파일내용수집(memberName, folderId, dateStr);

      if (content && content.내용) {
        Logger.log(`  ✅ 파일 발견! (${content.파일목록.length}개)`);

        조원데이터.push({
          이름: memberName,
          내용: content.내용,
          파일목록: content.파일목록
        });

        break; // 첫 번째 폴더에서 찾으면 중단
      } else {
        Logger.log(`  ❌ 이 폴더에서 찾을 수 없음`);
      }
    }
  }

  if (조원데이터.length === 0) {
    Logger.log('\n❌ 해당 날짜에 공부한 조원이 없습니다.');
    return null;
  }

  Logger.log(`\n✅ ${조원데이터.length}명의 데이터 수집 완료`);

  // 간단한 통합 요약 생성
  let 통합다이제스트 = `📚 ${dateStr} 스터디 다이제스트\n\n`;
  통합다이제스트 += `총 ${조원데이터.length}명 참여\n\n`;

  조원데이터.forEach((data, index) => {
    통합다이제스트 += `${index + 1}. ${data.이름} - ${data.파일목록.length}개 파일 제출\n`;
  });

  Logger.log('\n=== 다이제스트 완성 ===');
  Logger.log(통합다이제스트);

  다이제스트저장(통합다이제스트, 조원데이터, dateStr);

  // 🆕 월간 누적 데이터 저장 (매일 자동 누적)
  월간데이터누적(조원데이터, dateStr);

  // 🆕 Slack PDF 전송 (Bot Token 설정 시) 또는 링크 알림
  일일다이제스트_슬랙PDF전송(dateStr, 조원데이터.length);

  return 통합다이제스트;
}

/**
 * 🆕 월간 데이터 누적 저장
 * 일간 다이제스트 생성 시 호출되어 월간 데이터 JSON에 누적 저장
 * @param {Array} 조원데이터 - 일간 조원 데이터 [{이름, 내용, 파일목록}, ...]
 * @param {string} dateStr - 날짜 (yyyy-MM-dd)
 */
function 월간데이터누적(조원데이터, dateStr) {
  const yearMonth = dateStr.substring(0, 7); // 'yyyy-MM'
  const fileName = `monthly-data-${yearMonth}.json`;
  const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);

  try {
    // 기존 JSON 파일 읽기 (없으면 새로 생성)
    let jsonData = {
      년월: yearMonth,
      수집일시: new Date().toISOString(),
      조원데이터: {}
    };

    const files = folder.getFilesByName(fileName);
    if (files.hasNext()) {
      const file = files.next();
      jsonData = JSON.parse(file.getBlob().getDataAsString('UTF-8'));
      file.setTrashed(true); // 기존 파일 삭제
    }

    // 해당 날짜의 조원 데이터 추가/업데이트
    조원데이터.forEach(data => {
      const memberName = data.이름;

      // 조원 데이터가 없으면 초기화
      if (!jsonData.조원데이터[memberName]) {
        jsonData.조원데이터[memberName] = {
          한달내용: '',
          출석일수: 0,
          파일수: 0
        };
      }

      // 데이터 누적
      jsonData.조원데이터[memberName].한달내용 += `\n[${dateStr}]\n${data.내용}\n`;
      jsonData.조원데이터[memberName].출석일수++;
      jsonData.조원데이터[memberName].파일수 += data.파일목록.length;
    });

    // 갱신 시간 업데이트
    jsonData.수집일시 = new Date().toISOString();

    // JSON 파일로 저장
    const newFile = folder.createFile(fileName, JSON.stringify(jsonData, null, 2), MimeType.PLAIN_TEXT);
    newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    Logger.log(`✅ 월간 누적 데이터 저장: ${fileName} (${Object.keys(jsonData.조원데이터).length}명)`);

  } catch (error) {
    Logger.log(`⚠️ 월간 누적 저장 실패: ${error.message}`);
  }
}

/**
 * 🆕 2단계: 월간 AI 분석 실행 (시간 초과 방지)
 * 저장된 데이터를 읽어서 AI 분석 후 HTML 생성
 * @param {string} yearMonth - 년월 (yyyy-MM). 없으면 이번 달
 */
function 월간AI분석실행(yearMonth) {
  if (!yearMonth) {
    yearMonth = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM');
  }

  Logger.log(`\n=== [2단계] ${yearMonth} 월간 AI 분석 시작 ===\n`);

  // Gemini API 키 확인
  const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!GEMINI_API_KEY) {
    Logger.log('❌ Gemini API 키가 설정되지 않았습니다.');
    Logger.log('스크립트 속성에 GEMINI_API_KEY를 설정하세요.');
    return null;
  }

  // 저장된 데이터 파일 읽기
  const fileName = `monthly-data-${yearMonth}.json`;
  const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);

  const files = folder.getFilesByName(fileName);
  if (!files.hasNext()) {
    Logger.log(`❌ 수집된 데이터 파일이 없습니다: ${fileName}`);
    Logger.log('일간 다이제스트가 먼저 생성되어야 월간 데이터가 누적됩니다.');
    return null;
  }

  const file = files.next();
  const jsonData = JSON.parse(file.getBlob().getDataAsString('UTF-8'));
  const 조원데이터 = jsonData.조원데이터;

  Logger.log(`📁 데이터 파일 로드 완료: ${Object.keys(조원데이터).length}명\n`);

  const 조원분석결과 = [];

  // 각 조원별 AI 분석
  for (const [memberName, data] of Object.entries(조원데이터)) {
    Logger.log(`👤 ${memberName} AI 분석 중...`);

    const 분석결과 = AI월간분석(memberName, data.한달내용, data.출석일수, data.파일수, GEMINI_API_KEY);

    if (분석결과) {
      조원분석결과.push({
        이름: memberName,
        출석일수: data.출석일수,
        파일수: data.파일수,
        분석내용: 분석결과
      });
      Logger.log(`  ✅ 분석 완료\n`);
    } else {
      Logger.log(`  ❌ AI 분석 실패\n`);
    }

    // API 호출 제한 고려하여 잠시 대기
    Utilities.sleep(1000);
  }

  if (조원분석결과.length === 0) {
    Logger.log('\n❌ 분석 결과가 없습니다.');
    return null;
  }

  Logger.log(`\n✅ ${조원분석결과.length}명의 AI 분석 완료`);

  // 월간 다이제스트 저장
  월간다이제스트저장(조원분석결과, yearMonth);

  return 조원분석결과;
}

/**
 * 월간 AI 다이제스트 생성
 * 누적된 데이터(월간데이터누적)를 사용하여 AI 분석 실행
 * @param {string} yearMonth - 년월 (yyyy-MM). 없으면 이번 달
 */
function 월간AI다이제스트생성(yearMonth) {
  if (!yearMonth) {
    yearMonth = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM');
  }

  Logger.log(`\n${'='.repeat(60)}`);
  Logger.log(`📊 ${yearMonth} 월간 다이제스트 생성 시작`);
  Logger.log('='.repeat(60));

  // 누적된 데이터로 AI 분석 실행
  const 분석결과 = 월간AI분석실행(yearMonth);
  if (!분석결과) {
    Logger.log('\n❌ 월간 다이제스트 생성 실패');
    Logger.log('일간 다이제스트가 먼저 생성되어 데이터가 누적되어야 합니다.');
    return null;
  }

  Logger.log(`\n${'='.repeat(60)}`);
  Logger.log(`✅ 월간 다이제스트 생성 완료`);
  Logger.log('='.repeat(60));

  // 🆕 Slack PDF 전송 (Bot Token 설정 시) 또는 링크 알림
  월간다이제스트_슬랙PDF전송(yearMonth);

  return 분석결과;
}

/**
 * 🆕 트리거용: 월간 AI 분석 자동 실행 (전월)
 */
function 월간AI분석_자동실행() {
  const now = new Date();
  const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const yearMonth = Utilities.formatDate(lastMonth, 'Asia/Seoul', 'yyyy-MM');

  Logger.log(`\n🤖 [자동 트리거] 전월 AI 분석: ${yearMonth}`);

  return 월간AI분석실행(yearMonth);
}

/**
 * 🆕 트리거용: 월간 원본 수집 자동 실행 (전월)
 */
function 월간원본수집_자동실행() {
  const now = new Date();
  const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const yearMonth = Utilities.formatDate(lastMonth, 'Asia/Seoul', 'yyyy-MM');

  Logger.log(`\n🤖 [자동 트리거] 전월 원본 수집: ${yearMonth}`);

  return 월간원본수집(yearMonth);
}

/**
 * 🆕 월간 원본 파일 수집 (옵시디언용)
 * 각 조원의 원본 파일을 폴더 구조 그대로 복사
 * @param {string} yearMonth - 년월 (yyyy-MM). 없으면 이번 달
 * @returns {string} 생성된 폴더 URL
 */
function 월간원본수집(yearMonth) {
  if (!yearMonth) {
    yearMonth = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM');
  }

  Logger.log(`\n=== ${yearMonth} 월간 원본 파일 수집 시작 ===\n`);

  // 1. 컬렉션 폴더 생성 또는 찾기
  const collectionFolderName = `${yearMonth}-스터디모음`;
  const jsonFolder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);

  // 기존 폴더가 있으면 삭제
  const existingFolders = jsonFolder.getFoldersByName(collectionFolderName);
  while (existingFolders.hasNext()) {
    const folder = existingFolders.next();
    Logger.log(`⚠️ 기존 폴더 삭제: ${folder.getName()}`);
    folder.setTrashed(true);
  }

  const collectionFolder = jsonFolder.createFolder(collectionFolderName);
  Logger.log(`📁 컬렉션 폴더 생성: ${collectionFolderName}\n`);

  // 해당 월의 일수 계산
  const [year, month] = yearMonth.split('-');
  const lastDay = new Date(parseInt(year), parseInt(month), 0).getDate();

  let 총파일수 = 0;
  let 총폴더수 = 0;

  // 2. 각 조원별로 수집
  for (const [memberName, folderIdOrArray] of Object.entries(CONFIG.MEMBERS)) {
    const folderIds = Array.isArray(folderIdOrArray) ? folderIdOrArray : [folderIdOrArray];

    Logger.log(`👤 ${memberName} 파일 수집 중...`);

    // 조원 폴더 생성
    const memberFolder = collectionFolder.createFolder(memberName);
    let 조원파일수 = 0;

    // 3. 날짜별로 수집
    for (let day = 1; day <= lastDay; day++) {
      const dateStr = `${yearMonth}-${String(day).padStart(2, '0')}`;

      // 여러 날짜 형식 시도
      const dateFormats = [];
      const parts = dateStr.split('-');
      if (parts.length === 3) {
        const y = parts[0];
        const m = parts[1];
        const d = parts[2];

        dateFormats.push(
          `${y}-${m}-${d}`,      // 2025-11-24
          `${y}${m}${d}`,        // 20251124 (what 조원)
          `${y}.${m}.${d}`,      // 2025.11.24
          `${y}년 ${m}월 ${d}일`, // 2025년 11월 24일
          `${y}-${parseInt(m)}-${parseInt(d)}`  // 2025-1-5 (zero-pad 없는 형식)
        );
      }

      // 각 폴더 ID에서 날짜 폴더 찾기
      let sourceDateFolder = null;
      for (const folderId of folderIds) {
        try {
          const mainFolder = DriveApp.getFolderById(folderId);

          for (const format of dateFormats) {
            const folders = mainFolder.getFoldersByName(format);
            if (folders.hasNext()) {
              sourceDateFolder = folders.next();
              break;
            }
          }

          if (sourceDateFolder) break;
        } catch (e) {
          // 폴더 접근 불가
          continue;
        }
      }

      // 날짜 폴더가 있으면 파일 복사
      if (sourceDateFolder) {
        const destDateFolder = memberFolder.createFolder(dateStr);
        총폴더수++;

        const files = sourceDateFolder.getFiles();
        let 날짜파일수 = 0;

        while (files.hasNext()) {
          const file = files.next();
          const fileName = file.getName().toLowerCase();

          // off.md는 제외
          if (fileName === 'off.md' || fileName === 'off.txt') {
            continue;
          }

          // 파일 복사
          file.makeCopy(file.getName(), destDateFolder);
          날짜파일수++;
          조원파일수++;
          총파일수++;
        }

        if (날짜파일수 > 0) {
          Logger.log(`  ✅ ${dateStr}: ${날짜파일수}개 파일`);
        }
      }
    }

    Logger.log(`  📊 ${memberName}: 총 ${조원파일수}개 파일\n`);
  }

  Logger.log(`\n${'='.repeat(60)}`);
  Logger.log(`✅ 월간 원본 수집 완료`);
  Logger.log(`📁 총 폴더: ${총폴더수}개`);
  Logger.log(`📄 총 파일: ${총파일수}개`);
  Logger.log(`🔗 폴더 URL: ${collectionFolder.getUrl()}`);
  Logger.log('='.repeat(60));

  return collectionFolder.getUrl();
}

/**
 * AI로 조원의 한 달 학습 내용 분석
 */
function AI월간분석(memberName, 한달내용, 출석일수, 파일수, apiKey) {
  try {
    // 내용이 너무 길면 잘라내기 (Gemini API 토큰 제한 고려)
    const maxLength = 30000; // 약 3만자로 제한
    if (한달내용.length > maxLength) {
      한달내용 = 한달내용.substring(0, maxLength) + '\n\n... (내용이 너무 길어 일부만 포함됨)';
    }

    const prompt = `당신은 한의학 스터디 그룹의 학습 내용을 상세히 정리하는 전문가입니다.

아래는 "${memberName}" 조원이 한 달 동안 공부한 내용입니다.
- 출석일수: ${출석일수}일
- 제출 파일 수: ${파일수}개

학습 내용:
${한달내용}

다음 형식으로 **풍부하고 상세하게** 분석해주세요:

## 📚 주요 학습 주제
- 이 조원이 집중적으로 공부한 핵심 주제 5-7개를 구체적으로 나열해주세요
- 각 주제별로 어떤 내용을 다뤘는지 간단히 설명해주세요

## 📖 학습 내용 상세 요약
이 조원이 한 달간 공부한 **핵심 내용을 상세히 요약**해주세요:
- 주요 개념, 이론, 원리 등을 구체적으로 설명해주세요
- 중요한 한의학 용어나 개념이 있다면 포함해주세요
- 학습한 처방, 약재, 경혈, 질환 등 구체적인 내용을 언급해주세요
- 최소 10줄 이상으로 풍부하게 작성해주세요

## 🔑 핵심 키워드 & 개념
- 이번 달 학습에서 가장 중요한 키워드 10-15개를 나열해주세요
- 키워드는 한의학 용어, 처방명, 약재명, 질환명, 경혈명 등을 포함해주세요

## 📈 학습 흐름 분석
- 한 달간 학습이 어떻게 진행되었는지 흐름을 분석해주세요
- 초반/중반/후반에 어떤 주제를 다뤘는지, 연결성이 있는지 설명해주세요

## 💡 주목할 만한 인사이트
- 학습 내용 중 특별히 깊이 있게 다룬 부분이나 흥미로운 관점이 있다면 언급해주세요
- 임상적으로 유용한 내용이나 실제 적용 가능한 지식이 있다면 강조해주세요

## 🌟 학습 특징 및 강점
- 이 조원만의 학습 스타일이나 특징을 2-3가지 언급해주세요

**중요: 공부 내용 요약은 최대한 구체적이고 풍부하게 작성해주세요. 단순 나열이 아닌, 실제로 어떤 내용을 배웠는지 알 수 있도록 상세히 기술해주세요.**`;

    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;

    const payload = {
      contents: [{
        parts: [{
          text: prompt
        }]
      }]
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    if (result.candidates && result.candidates[0] && result.candidates[0].content) {
      return result.candidates[0].content.parts[0].text;
    } else {
      Logger.log(`  AI 응답 오류: ${JSON.stringify(result)}`);
      return null;
    }

  } catch (e) {
    Logger.log(`  AI 분석 오류: ${e.message}`);
    return null;
  }
}

/**
 * 월간 다이제스트 저장 (드라이브 + 시트)
 */
function 월간다이제스트저장(조원분석결과, yearMonth) {
  Logger.log(`\n📝 월간 다이제스트 저장 시작: ${yearMonth}`);

  // HTML 생성
  let htmlContent = `<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>📊 ${yearMonth} 월간 스터디 다이제스트</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Apple SD Gothic Neo", "Malgun Gothic", sans-serif;
            line-height: 1.8;
            color: #333;
            background: #f8f9fa;
            padding: 20px;
        }
        .container {
            max-width: 900px;
            margin: 0 auto;
            background: white;
            padding: 40px;
            border-radius: 12px;
            box-shadow: 0 2px 12px rgba(0,0,0,0.1);
        }
        .header {
            text-align: center;
            margin-bottom: 50px;
            padding-bottom: 30px;
            border-bottom: 4px solid #4CAF50;
        }
        .header h1 {
            font-size: 32px;
            color: #2c3e50;
            margin-bottom: 15px;
        }
        .meta {
            color: #7f8c8d;
            font-size: 16px;
        }
        .member-analysis {
            margin-bottom: 60px;
            padding: 35px;
            background: linear-gradient(to bottom, #f8f9fa, #ffffff);
            border-radius: 10px;
            border-left: 5px solid #4CAF50;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        }
        .member-analysis h2 {
            font-size: 26px;
            color: #2c3e50;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
        }
        .member-analysis h2::before {
            content: "👤";
            margin-right: 12px;
        }
        .stats {
            display: flex;
            gap: 20px;
            margin-bottom: 25px;
            padding: 15px;
            background: white;
            border-radius: 8px;
        }
        .stat-item {
            flex: 1;
            text-align: center;
            padding: 10px;
        }
        .stat-number {
            font-size: 24px;
            font-weight: bold;
            color: #4CAF50;
        }
        .stat-label {
            font-size: 13px;
            color: #7f8c8d;
            margin-top: 5px;
        }
        .analysis-content {
            color: #555;
            font-size: 15px;
        }
        .analysis-content h2 {
            font-size: 20px;
            color: #34495e;
            margin-top: 25px;
            margin-bottom: 12px;
            border-bottom: 2px solid #ecf0f1;
            padding-bottom: 8px;
        }
        .analysis-content h2::before {
            content: "";
        }
        .analysis-content ul {
            margin: 15px 0;
            padding-left: 25px;
        }
        .analysis-content li {
            margin: 8px 0;
            line-height: 1.6;
        }
        .pdf-button {
            display: inline-block;
            margin-top: 20px;
            padding: 14px 35px;
            background: #4CAF50;
            color: white;
            border: none;
            border-radius: 8px;
            font-size: 17px;
            font-weight: bold;
            cursor: pointer;
            text-decoration: none;
            box-shadow: 0 3px 10px rgba(76, 175, 80, 0.3);
            transition: all 0.3s ease;
        }
        .pdf-button:hover {
            background: #45a049;
            box-shadow: 0 5px 15px rgba(76, 175, 80, 0.4);
            transform: translateY(-2px);
        }
        @media print {
            body {
                background: white;
                padding: 0;
            }
            .container {
                box-shadow: none;
                padding: 20px;
            }
            .pdf-button {
                display: none;
            }
            .member-analysis {
                page-break-inside: avoid;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 ${yearMonth} 월간 스터디 다이제스트</h1>
            <div class="meta">
                생성일시: ${Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')} |
                총 ${조원분석결과.length}명 분석
            </div>
            <button class="pdf-button" onclick="window.print()">
                📄 PDF로 저장하기
            </button>
        </div>
`;

  조원분석결과.forEach((data) => {
    htmlContent += `
        <div class="member-analysis">
            <h2>${data.이름}</h2>

            <div class="stats">
                <div class="stat-item">
                    <div class="stat-number">${data.출석일수}일</div>
                    <div class="stat-label">출석</div>
                </div>
                <div class="stat-item">
                    <div class="stat-number">${data.파일수}개</div>
                    <div class="stat-label">파일</div>
                </div>
                <div class="stat-item">
                    <div class="stat-number">${Math.round(data.파일수 / data.출석일수 * 10) / 10}</div>
                    <div class="stat-label">평균 파일/일</div>
                </div>
            </div>

            <div class="analysis-content">
                ${마크다운을HTML로(data.분석내용)}
            </div>
        </div>
`;
  });

  htmlContent += `
    </div>
</body>
</html>`;

  // 드라이브에 저장
  const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);
  const htmlFileName = `monthly-digest-${yearMonth}.html`;

  const existingFiles = folder.getFilesByName(htmlFileName);
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }

  const htmlFile = folder.createFile(htmlFileName, htmlContent, MimeType.HTML);
  htmlFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const fileId = htmlFile.getId();

  Logger.log(`✅ 드라이브에 HTML 파일 저장: ${htmlFileName}`);
  Logger.log(`  - 파일 ID: ${fileId}`);

  // 시트에 파일 ID 저장
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.DIGEST_SHEET);

  if (sheet) {
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

    // 기존 같은 월 데이터 삭제
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i > 0; i--) {
      if (data[i][0] === `MONTHLY-${yearMonth}`) {
        sheet.deleteRow(i + 1);
        Logger.log(`기존 ${yearMonth} 월간 다이제스트 삭제됨`);
      }
    }

    // 새 데이터 추가
    sheet.insertRowBefore(2);
    sheet.getRange(2, 1, 1, 3).setValues([[
      `MONTHLY-${yearMonth}`,
      fileId,
      timestamp
    ]]);

    Logger.log(`✅ 시트에 파일 ID 저장 완료`);
  }

  Logger.log(`\n📱 웹앱 URL로 확인:`);
  Logger.log(`https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec?date=MONTHLY-${yearMonth}`);
}

/**
 * 텍스트 파일 인코딩 자동 감지 읽기
 * UTF-8 → EUC-KR 순서로 시도하여 유효한 내용을 반환
 * @param {GoogleAppsScript.Drive.File} file - Drive 파일 객체
 * @returns {string} 파일 내용 (읽기 실패 시 빈 문자열)
 */
function 텍스트파일읽기(file) {
  const blob = file.getBlob();
  const charsets = ['UTF-8', 'EUC-KR'];

  for (const charset of charsets) {
    try {
      const content = blob.getDataAsString(charset);
      if (content && content.trim().length > 0 && !content.includes('\uFFFD')) {
        if (charset !== 'UTF-8') {
          Logger.log(`  ℹ️ ${charset} 인코딩으로 읽음: ${file.getName()}`);
        }
        return content;
      }
    } catch (e) {
      // 다음 인코딩 시도
    }
  }

  Logger.log(`  ⚠️ 텍스트 파일 인코딩 감지 실패: ${file.getName()}`);
  return '';
}

/**
 * 파일 내용 수집
 * @param {string} memberName - 조원 이름
 * @param {string} folderId - 조원 폴더 ID
 * @param {string} dateStr - 날짜 (yyyy-MM-dd)
 * @returns {Object} {내용, 파일목록}
 */
function 파일내용수집(memberName, folderId, dateStr) {
  try {
    // dateStr은 이미 yyyy-MM-dd 형식 (예: 2025-11-21)
    Logger.log(`  찾는 중: ${dateStr}`);

    // 조원 폴더
    const memberFolder = DriveApp.getFolderById(folderId);
    Logger.log(`  📁 조원 폴더: ${memberFolder.getName()}`);

    // 🔍 디버깅: 이 폴더에 어떤 하위 폴더들이 있는지 확인
    const allFolders = memberFolder.getFolders();
    const folderNames = [];
    while (allFolders.hasNext() && folderNames.length < 10) {
      folderNames.push(allFolders.next().getName());
    }
    Logger.log(`  📂 하위 폴더들: ${folderNames.join(', ')}`);

    // 여러 날짜 형식 시도 (what 조원은 yyyyMMdd 형식 사용)
    const dateFormats = [];

    // dateStr이 yyyy-MM-dd 형식이라고 가정 (예: 2025-11-24)
    const parts = dateStr.split('-');
    if (parts.length === 3) {
      const year = parts[0];
      const month = parts[1];
      const day = parts[2];

      // 다양한 날짜 형식 생성
      const shortYear = year.slice(-2);     // 2026 → 26
      dateFormats.push(
        `${year}-${month}-${day}`,           // 2025-11-24
        `${year}${month}${day}`,             // 20251124 (what 조원)
        `${year}.${month}.${day}`,           // 2025.11.24
        `${year}년 ${month}월 ${day}일`,     // 2025년 11월 24일
        `${year}-${parseInt(month)}-${parseInt(day)}`,  // 2025-1-5 (zero-pad 없는 형식)
        `${shortYear}-${month}-${day}`,      // 25-11-24 (부천개원의)
        `${shortYear}-${parseInt(month)}-${parseInt(day)}`,  // 25-1-5
        `${shortYear}${month}${day}`,        // 251124
        `${shortYear}.${month}.${day}`,      // 25.11.24
      );
    }

    Logger.log(`  🔍 시도할 날짜 형식: ${dateFormats.join(', ')}`);

    // 각 형식을 순서대로 시도
    let dateFolder = null;
    for (const format of dateFormats) {
      const folders = memberFolder.getFoldersByName(format);
      if (folders.hasNext()) {
        dateFolder = folders.next();
        Logger.log(`  ✅ 폴더 발견: ${dateFolder.getName()} (형식: ${format})`);
        break;
      }
    }

    // 모든 형식을 시도했지만 찾지 못함
    if (!dateFolder) {
      Logger.log(`  ❌ 날짜 폴더 없음 (모든 형식 시도함)`);
      Logger.log(`  💡 찾은 하위 폴더: ${folderNames.length}개`);
      return null;
    }

    // 🆕 먼저 off.md 파일이 있는지 확인 (오프한 사람은 다이제스트에서 제외)
    const allFiles = dateFolder.getFiles();
    let hasOffFile = false;

    while (allFiles.hasNext()) {
      const file = allFiles.next();
      const fileName = file.getName().toLowerCase();

      if (fileName === 'off.md' || fileName === 'off.txt') {
        hasOffFile = true;
        break;
      }
    }

    if (hasOffFile) {
      Logger.log(`  🏖️ 오프 (off.md 발견) - 다이제스트에서 제외`);
      return null;
    }

    let 전체내용 = '';
    const 파일목록 = [];
    const files = dateFolder.getFiles();

    let fileCount = 0;
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      const mimeType = file.getMimeType();

      fileCount++;

      // 마크다운 파일
      if (fileName.toLowerCase().endsWith('.md')) {
        try {
          const mdContent = 텍스트파일읽기(file);

          // 마크다운 클린업 적용
          const cleanedContent = 마크다운클린업(mdContent);

          // 제목 추출
          const titleMatch = cleanedContent.match(/^#\s+(.+)$/m);
          if (titleMatch) {
            전체내용 += `[제목: ${titleMatch[1]}]\n\n`;
          }

          전체내용 += cleanedContent + '\n\n' + '='.repeat(50) + '\n\n';

          파일목록.push({
            이름: fileName,
            타입: 'Markdown'
          });

        } catch (e) {
          Logger.log(`  MD 파일 읽기 실패: ${fileName}`);
        }
      }

      // 텍스트 파일
      else if (mimeType === MimeType.PLAIN_TEXT || fileName.toLowerCase().endsWith('.txt')) {
        try {
          const txtContent = 텍스트파일읽기(file);

          if (txtContent && txtContent.trim().length > 0) {
            전체내용 += `[텍스트 파일: ${fileName}]\n\n${txtContent}\n\n` + '='.repeat(50) + '\n\n';
          } else {
            전체내용 += `[텍스트 파일: ${fileName}] (내용 추출 실패)\n\n` + '='.repeat(50) + '\n\n';
            Logger.log(`  ⚠️ 텍스트 파일 내용 비어있음: ${fileName}`);
          }

          파일목록.push({
            이름: fileName,
            타입: 'Text'
          });
        } catch (e) {
          Logger.log(`  텍스트 파일 읽기 실패: ${fileName}`);
        }
      }

      // PDF (OCR로 텍스트 추출)
      else if (mimeType === MimeType.PDF) {
        try {
          Logger.log(`  PDF OCR 시작: ${fileName}`);
          const pdfContent = PDF텍스트추출(file);

          if (pdfContent && pdfContent.trim().length > 0) {
            전체내용 += `[PDF 문서: ${fileName}]\n\n${pdfContent}\n\n` + '='.repeat(50) + '\n\n';
            Logger.log(`  PDF OCR 성공: ${fileName} (${pdfContent.length}자)`);
          } else {
            전체내용 += `[PDF 문서: ${fileName}] (텍스트 추출 실패 또는 이미지 PDF)\n\n`;
            Logger.log(`  PDF OCR 실패 또는 빈 내용: ${fileName}`);
          }

          파일목록.push({
            이름: fileName,
            타입: 'PDF'
          });
        } catch (e) {
          Logger.log(`  PDF 처리 실패: ${fileName} - ${e.message}`);
          전체내용 += `[PDF 문서: ${fileName}] (처리 실패)\n\n`;
          파일목록.push({
            이름: fileName,
            타입: 'PDF'
          });
        }
      }

      // 이미지
      else if (mimeType.startsWith('image/')) {
        try {
          // 이미지를 base64로 인코딩 (HTML 임베드용)
          const blob = file.getBlob();
          const base64Data = Utilities.base64Encode(blob.getBytes());

          전체내용 += `[이미지: ${fileName}]\n\n`;

          파일목록.push({
            이름: fileName,
            타입: 'Image',
            mimeType: mimeType,
            base64: base64Data  // HTML 렌더링용
          });

          Logger.log(`  이미지 인코딩 완료: ${fileName}`);
        } catch (e) {
          Logger.log(`  이미지 처리 실패: ${fileName} - ${e.message}`);
          // 실패해도 파일 목록에는 추가
          파일목록.push({
            이름: fileName,
            타입: 'Image'
          });
        }
      }
    }

    Logger.log(`  총 ${fileCount}개 파일, 텍스트 추출: ${파일목록.length}개`);

    if (전체내용.trim().length > 0) {
      return { 내용: 전체내용, 파일목록 };
    }

    // 텍스트가 없어도 파일이 있으면 기본 정보 반환
    if (파일목록.length > 0) {
      return {
        내용: `${memberName}이(가) ${dateStr}에 ${파일목록.length}개 파일을 제출했습니다.`,
        파일목록
      };
    }

    return null;

  } catch (e) {
    Logger.log(`  ${memberName} 파일 수집 실패: ${e.message}`);
    Logger.log(`  Stack: ${e.stack}`);
    return null;
  }
}

/**
 * 🆕 다이제스트 시트 초기화
 * - 스프레드시트에 "다이제스트" 시트 생성
 * - 드라이브 HTML 파일 ID를 시트에 저장하여 관리
 * - 배포 설정 "실행 계정: 나"로 권한 문제 해결
 */
function 다이제스트시트초기화() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 기존 시트 삭제
  const existing = ss.getSheetByName(CONFIG.DIGEST_SHEET);
  if (existing) {
    ss.deleteSheet(existing);
    Logger.log('기존 다이제스트 시트 삭제됨');
  }

  // 새 시트 생성
  const sheet = ss.insertSheet(CONFIG.DIGEST_SHEET);

  // 헤더 설정
  const headers = ['날짜', 'JSON데이터', '생성시각'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 헤더 스타일
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#4CAF50')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // 열 너비 설정
  sheet.setColumnWidth(1, 120);  // 날짜
  sheet.setColumnWidth(2, 600);  // JSON데이터 (넓게)
  sheet.setColumnWidth(3, 180);  // 생성시각

  // 헤더 고정
  sheet.setFrozenRows(1);

  Logger.log('✅ 다이제스트 시트 초기화 완료');
  Logger.log('💡 드라이브 HTML 파일 ID를 이 시트에서 관리합니다');
}

/**
 * 이미지 압축 함수
 * - 큰 이미지를 축소하여 base64 크기 감소
 * - Google Apps Script의 이미지 처리 한계로 인해 크기 제한 방식 사용
 * @param {string} base64Data - 원본 base64 이미지 데이터
 * @param {string} mimeType - 이미지 MIME 타입
 * @returns {string} 압축된 base64 데이터 (또는 원본)
 */
function 이미지압축(base64Data, mimeType) {
  try {
    // base64 크기가 1MB 이하면 그대로 반환
    const MAX_IMAGE_SIZE = 1 * 1024 * 1024; // 1MB
    if (base64Data.length <= MAX_IMAGE_SIZE) {
      return base64Data;
    }

    Logger.log(`  🗜️ 이미지 압축 시도: ${(base64Data.length / 1024 / 1024).toFixed(2)}MB`);

    // base64를 blob으로 변환
    const decodedData = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decodedData, mimeType, 'image');

    // PNG를 JPEG로 변환하여 압축 (JPEG가 일반적으로 더 작음)
    let compressedBlob;
    if (mimeType === 'image/png') {
      // PNG -> JPEG 변환
      compressedBlob = blob.getAs('image/jpeg');
    } else {
      // 이미 JPEG인 경우 썸네일로 축소 시도
      compressedBlob = blob;
    }

    // 압축된 데이터를 base64로 재인코딩
    const compressedBase64 = Utilities.base64Encode(compressedBlob.getBytes());

    // 압축 후에도 너무 크면 null 반환 (이미지 스킵)
    if (compressedBase64.length > MAX_IMAGE_SIZE * 2) {
      Logger.log(`  ⚠️ 압축 후에도 너무 큼 - 이미지 스킵: ${(compressedBase64.length / 1024 / 1024).toFixed(2)}MB`);
      return null;
    }

    const compressionRatio = ((1 - compressedBase64.length / base64Data.length) * 100).toFixed(1);
    Logger.log(`  ✅ 이미지 압축 완료: ${(compressedBase64.length / 1024 / 1024).toFixed(2)}MB (${compressionRatio}% 감소)`);

    return compressedBase64;
  } catch (e) {
    Logger.log(`  ❌ 이미지 압축 실패: ${e.message}`);
    // 압축 실패 시 원본 반환
    return base64Data;
  }
}

/**
 * 다이제스트 저장 (드라이브 + 시트 하이브리드) 🆕
 * - HTML 파일은 드라이브에 저장 (이미지 base64 포함 가능)
 * - 파일 ID는 시트에 저장 (관리 편의성)
 * - 배포 "실행 계정: 나"로 모든 사용자가 접근 가능
 * @param {string} 통합다이제스트 - 통합 다이제스트 텍스트
 * @param {Array} 조원데이터 - 조원별 상세 데이터
 * @param {string} dateStr - 날짜
 */
function 다이제스트저장(통합다이제스트, 조원데이터, dateStr, options = {}) {
  Logger.log(`\n📝 다이제스트 저장 시작: ${dateStr}`);

  // 기본 옵션 설정
  options = {
    compressImages: options.compressImages || false,
    skipImages: options.skipImages || false,
    ...options
  };

  // 최대 파일 크기 (10MB - 여유분 고려)
  const MAX_FILE_SIZE = 10 * 1024 * 1024;

  // 1. HTML 파일 생성 (시트에 저장할 내용)
  let htmlContent = `<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>📚 ${dateStr} 스터디 다이제스트</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Apple SD Gothic Neo", "Malgun Gothic", sans-serif;
            line-height: 1.7;
            color: #333;
            background: #f8f9fa;
            padding: 20px;
        }
        .container {
            max-width: 800px;
            margin: 0 auto;
            background: white;
            padding: 40px;
            border-radius: 12px;
            box-shadow: 0 2px 12px rgba(0,0,0,0.1);
        }
        .header {
            text-align: center;
            margin-bottom: 40px;
            padding-bottom: 20px;
            border-bottom: 3px solid #4CAF50;
        }
        .header h1 {
            font-size: 28px;
            color: #2c3e50;
            margin-bottom: 10px;
        }
        .meta {
            color: #7f8c8d;
            font-size: 14px;
        }
        .member-section {
            margin-bottom: 50px;
            padding: 30px;
            background: #f8f9fa;
            border-radius: 8px;
            border-left: 4px solid #4CAF50;
        }
        .member-section h2 {
            font-size: 24px;
            color: #2c3e50;
            margin-bottom: 15px;
            display: flex;
            align-items: center;
        }
        .member-section h2::before {
            content: "👤";
            margin-right: 10px;
        }
        .file-list {
            background: white;
            padding: 15px;
            border-radius: 6px;
            margin: 15px 0;
        }
        .file-list h3 {
            font-size: 16px;
            color: #34495e;
            margin-bottom: 10px;
        }
        .file-list ul {
            list-style: none;
            padding-left: 0;
        }
        .file-list li {
            padding: 8px 0;
            border-bottom: 1px solid #ecf0f1;
            color: #555;
        }
        .file-list li:last-child {
            border-bottom: none;
        }
        .file-list li::before {
            content: "📄";
            margin-right: 8px;
        }
        .content-section {
            margin-top: 20px;
        }
        .content-section h3 {
            font-size: 18px;
            color: #34495e;
            margin-bottom: 15px;
            padding-bottom: 10px;
            border-bottom: 2px solid #ecf0f1;
        }
        .content-body {
            background: white;
            padding: 20px;
            border-radius: 6px;
            line-height: 1.8;
        }
        .content-body h1 {
            font-size: 22px;
            color: #2c3e50;
            margin: 25px 0 15px 0;
            padding-bottom: 10px;
            border-bottom: 2px solid #4CAF50;
        }
        .content-body h2 {
            font-size: 20px;
            color: #34495e;
            margin: 20px 0 12px 0;
        }
        .content-body h3 {
            font-size: 18px;
            color: #555;
            margin: 15px 0 10px 0;
        }
        .content-body ul {
            margin: 15px 0;
            padding-left: 25px;
        }
        .content-body li {
            margin: 8px 0;
        }
        .content-body strong {
            color: #2c3e50;
            font-weight: 600;
        }
        .content-body em {
            color: #7f8c8d;
            font-style: italic;
        }
        .content-body code {
            background: #f4f4f4;
            padding: 2px 6px;
            border-radius: 3px;
            font-family: "Monaco", "Courier New", monospace;
            font-size: 0.9em;
            color: #e74c3c;
        }
        .content-body p {
            margin: 12px 0;
        }
        .image-gallery {
            margin: 20px 0;
        }
        .image-gallery h4 {
            font-size: 16px;
            color: #34495e;
            margin-bottom: 10px;
        }
        .image-item {
            margin: 15px 0;
            text-align: center;
        }
        .image-item img {
            max-width: 100%;
            height: auto;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            margin-bottom: 8px;
        }
        .image-caption {
            font-size: 13px;
            color: #7f8c8d;
            font-style: italic;
        }
        @media (max-width: 768px) {
            .container {
                padding: 20px;
            }
            .header h1 {
                font-size: 22px;
            }
            .member-section {
                padding: 20px;
            }
        }

        /* PDF 다운로드 버튼 스타일 */
        .pdf-button {
            display: inline-block;
            margin-top: 20px;
            padding: 12px 30px;
            background: #4CAF50;
            color: white;
            border: none;
            border-radius: 6px;
            font-size: 16px;
            font-weight: bold;
            cursor: pointer;
            text-decoration: none;
            box-shadow: 0 2px 8px rgba(76, 175, 80, 0.3);
            transition: all 0.3s ease;
        }
        .pdf-button:hover {
            background: #45a049;
            box-shadow: 0 4px 12px rgba(76, 175, 80, 0.4);
            transform: translateY(-2px);
        }
        .pdf-button:active {
            transform: translateY(0);
        }

        /* 인쇄(PDF 생성) 시 스타일 */
        @media print {
            body {
                background: white;
                padding: 0;
            }
            .container {
                box-shadow: none;
                padding: 20px;
                max-width: 100%;
            }
            .pdf-button {
                display: none; /* PDF 생성 시 버튼 숨김 */
            }
            .member-section {
                page-break-inside: avoid; /* 섹션이 페이지 중간에 나뉘지 않도록 */
                margin-bottom: 30px;
            }
            .image-gallery img {
                max-width: 100%;
                page-break-inside: avoid;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📚 ${dateStr} 스터디 다이제스트</h1>
            <div class="meta">
                생성일시: ${Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')} |
                참여: ${조원데이터.length}명
            </div>
            <button class="pdf-button" onclick="window.print()">
                📄 PDF로 저장하기
            </button>
        </div>
`;

  조원데이터.forEach((data, index) => {
    htmlContent += `
        <div class="member-section">
            <h2>${data.이름}</h2>

            <div class="file-list">
                <h3>📁 제출 파일 (${data.파일목록.length}개)</h3>
                <ul>
`;

    data.파일목록.forEach(file => {
      htmlContent += `                    <li>${file.이름} <span style="color: #95a5a6;">(${file.타입})</span></li>\n`;
    });

    htmlContent += `                </ul>
            </div>

            <div class="content-section">
                <h3>📖 학습 내용</h3>
                <div class="content-body">
`;

    // 마크다운을 HTML로 변환
    const cleanedContent = 마크다운클린업(data.내용);
    const htmlBody = 마크다운을HTML로(cleanedContent);
    htmlContent += htmlBody;

    htmlContent += `
                </div>
            </div>
`;

    // 이미지 갤러리 추가 (base64 데이터가 있는 이미지만)
    const images = data.파일목록.filter(f => f.타입 === 'Image' && f.base64);
    if (images.length > 0 && !options.skipImages) {
      // 압축 모드에서 null인 이미지(너무 큰 이미지)는 제외
      const validImages = options.compressImages
        ? images.filter(img => img.compressedBase64 !== null)
        : images;

      if (validImages.length > 0) {
        htmlContent += `
            <div class="image-gallery">
                <h4>📸 첨부 이미지 (${validImages.length}개)</h4>
`;
        validImages.forEach(img => {
          // 압축 모드인 경우 압축된 이미지 사용
          const imageBase64 = options.compressImages ? (img.compressedBase64 || img.base64) : img.base64;
          if (!imageBase64) return; // null 체크
          const dataUri = `data:${img.mimeType};base64,${imageBase64}`;
          htmlContent += `
                <div class="image-item">
                    <img src="${dataUri}" alt="${img.이름}">
                    <div class="image-caption">${img.이름}</div>
                </div>
`;
        });
        htmlContent += `
            </div>
`;
      }
    }

    htmlContent += `
        </div>
`;
  });

  htmlContent += `
    </div>
</body>
</html>`;

  Logger.log(`\n📏 HTML 길이: ${htmlContent.length} 문자`);

  // 용량 초과 시 재시도 로직
  if (htmlContent.length > MAX_FILE_SIZE) {
    if (!options.compressImages && !options.skipImages) {
      // 1차 시도: 이미지 압축
      Logger.log(`⚠️ 파일 크기 초과 (${(htmlContent.length / 1024 / 1024).toFixed(2)}MB) - 이미지 압축 후 재시도`);

      // 이미지 압축 수행
      조원데이터.forEach(data => {
        data.파일목록.forEach(file => {
          if (file.타입 === 'Image' && file.base64) {
            file.compressedBase64 = 이미지압축(file.base64, file.mimeType);
          }
        });
      });

      return 다이제스트저장(통합다이제스트, 조원데이터, dateStr, { compressImages: true });
    } else if (options.compressImages && !options.skipImages) {
      // 2차 시도: 이미지 제외
      Logger.log(`⚠️ 압축 후에도 크기 초과 (${(htmlContent.length / 1024 / 1024).toFixed(2)}MB) - 이미지 제외 후 재시도`);
      return 다이제스트저장(통합다이제스트, 조원데이터, dateStr, { skipImages: true });
    } else {
      // 그래도 초과하면 에러
      Logger.log(`❌ 이미지 제외 후에도 크기 초과 - 텍스트 내용이 너무 큽니다.`);
      throw new Error(`다이제스트 파일 크기 초과: ${(htmlContent.length / 1024 / 1024).toFixed(2)}MB`);
    }
  }

  if (options.compressImages) {
    Logger.log(`✅ 이미지 압축 적용됨`);
  }
  if (options.skipImages) {
    Logger.log(`✅ 이미지 제외됨 (텍스트만 저장)`);
  }

  // 2. 드라이브에 HTML 파일 저장 (이미지 포함, 권한 문제 해결!)
  const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);
  const htmlFileName = `digest-${dateStr}.html`;

  // 기존 파일 삭제
  const existingFiles = folder.getFilesByName(htmlFileName);
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }

  // 새 파일 생성
  const htmlFile = folder.createFile(htmlFileName, htmlContent, MimeType.HTML);
  htmlFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const fileId = htmlFile.getId();

  Logger.log(`✅ 드라이브에 HTML 파일 저장: ${htmlFileName}`);
  Logger.log(`  - 파일 ID: ${fileId}`);

  // 3. 스프레드시트 시트에 파일 정보 저장 (HTML 대신 파일 ID 저장)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.DIGEST_SHEET);

  if (!sheet) {
    Logger.log('⚠️ 다이제스트 시트가 없습니다. 다이제스트시트초기화() 함수를 먼저 실행하세요!');
    throw new Error('다이제스트 시트가 없습니다. 다이제스트시트초기화() 함수를 먼저 실행하세요.');
  }

  // 생성 시각
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

  // 기존 같은 날짜 데이터 삭제 (있으면)
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i > 0; i--) {  // 헤더 제외하고 역순으로 검색
    if (data[i][0] === dateStr) {
      sheet.deleteRow(i + 1);
      Logger.log(`기존 ${dateStr} 다이제스트 삭제됨`);
    }
  }

  // 새 데이터 추가 (날짜, 파일ID, 생성시각)
  sheet.insertRowBefore(2);
  sheet.getRange(2, 1, 1, 3).setValues([[
    dateStr,
    fileId,  // HTML 내용 대신 드라이브 파일 ID 저장!
    timestamp
  ]]);

  Logger.log(`\n✅ 다이제스트 저장 완료!`);
  Logger.log(`  - 날짜: ${dateStr}`);
  Logger.log(`  - 파일 ID: ${fileId}`);
  Logger.log(`  - HTML 길이: ${htmlContent.length} 문자`);
  Logger.log(`  - 참여자: ${조원데이터.length}명`);
  Logger.log(`  - 생성 시각: ${timestamp}`);
  Logger.log(`\n📱 웹앱 URL로 확인 가능:`);
  Logger.log(`https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec?date=${dateStr}`);
}

/**
 * 🆕 PDF 텍스트 추출 함수
 * Google Drive OCR을 사용하여 PDF에서 텍스트 추출
 * @param {File} pdfFile - Google Drive PDF 파일 객체
 * @returns {string} 추출된 텍스트 (실패 시 빈 문자열)
 */
function PDF텍스트추출(pdfFile) {
  let tempDocId = null;

  try {
    const blob = pdfFile.getBlob();
    const fileName = pdfFile.getName().replace(/\.pdf$/i, '_OCR_TEMP');

    // Google Drive API v3를 사용하여 PDF를 Google Docs로 변환 (OCR 적용)
    const boundary = '-------314159265358979323846';
    const delimiter = "\r\n--" + boundary + "\r\n";
    const closeDelimiter = "\r\n--" + boundary + "--";

    const metadata = {
      name: fileName,
      mimeType: 'application/vnd.google-apps.document'
    };

    const requestBody =
      delimiter +
      'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
      JSON.stringify(metadata) +
      delimiter +
      'Content-Type: application/pdf\r\n' +
      'Content-Transfer-Encoding: base64\r\n\r\n' +
      Utilities.base64Encode(blob.getBytes()) +
      closeDelimiter;

    const response = UrlFetchApp.fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&ocrLanguage=ko',
      {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
          'Content-Type': 'multipart/related; boundary=' + boundary
        },
        payload: requestBody,
        muteHttpExceptions: true
      }
    );

    const result = JSON.parse(response.getContentText());

    if (result.error) {
      Logger.log(`PDF 업로드 오류: ${result.error.message}`);
      return '';
    }

    tempDocId = result.id;

    // 변환된 Google Docs에서 텍스트 추출
    const doc = DocumentApp.openById(tempDocId);
    const text = doc.getBody().getText();

    // 임시 파일 삭제
    DriveApp.getFileById(tempDocId).setTrashed(true);
    tempDocId = null;

    return text.trim();

  } catch (e) {
    Logger.log(`PDF OCR 오류: ${e.message}`);

    // 임시 파일이 생성되었다면 삭제
    if (tempDocId) {
      try {
        DriveApp.getFileById(tempDocId).setTrashed(true);
      } catch (deleteError) {
        Logger.log(`임시 파일 삭제 실패: ${deleteError.message}`);
      }
    }

    return '';
  }
}

