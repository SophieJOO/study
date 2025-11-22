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
 * 일일 다이제스트 생성
 * @param {string} dateStr - 날짜 (yyyy-MM-dd). 없으면 어제
 * @returns {string} 생성된 다이제스트
 */
function 일일AI다이제스트생성(dateStr) {
  if (!dateStr) {
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

  return 통합다이제스트;
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

    // 날짜 폴더 찾기 (yyyy-MM-dd 형식)
    const dateFolders = memberFolder.getFoldersByName(dateStr);
    if (!dateFolders.hasNext()) {
      Logger.log(`  날짜 폴더 없음: ${dateStr}`);
      return null;
    }

    const dateFolder = dateFolders.next();
    Logger.log(`  ✅ 폴더 발견: ${dateFolder.getName()}`);

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
          const mdContent = file.getBlob().getDataAsString('UTF-8');

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
          const txtContent = file.getBlob().getDataAsString('UTF-8');
          전체내용 += `[텍스트 파일: ${fileName}]\n\n${txtContent}\n\n` + '='.repeat(50) + '\n\n';

          파일목록.push({
            이름: fileName,
            타입: 'Text'
          });
        } catch (e) {
          Logger.log(`  텍스트 파일 읽기 실패: ${fileName}`);
        }
      }

      // PDF (파일명만 - OCR은 별도 구현 필요)
      else if (mimeType === MimeType.PDF) {
        전체내용 += `[PDF 문서: ${fileName}]\n\n`;
        파일목록.push({
          이름: fileName,
          타입: 'PDF'
        });
      }

      // 이미지
      else if (mimeType.startsWith('image/')) {
        전체내용 += `[이미지: ${fileName}]\n\n`;
        파일목록.push({
          이름: fileName,
          타입: 'Image'
        });
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
 * 다이제스트 저장
 * @param {string} 통합다이제스트 - 통합 다이제스트 텍스트
 * @param {Array} 조원데이터 - 조원별 상세 데이터
 * @param {string} dateStr - 날짜
 */
function 다이제스트저장(통합다이제스트, 조원데이터, dateStr) {
  const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);

  // 1. 전체 원본 내용 파일 생성
  let 전체내용 = `📚 ${dateStr} 스터디 전체 내용\n`;
  전체내용 += `생성일시: ${Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')}\n`;
  전체내용 += `총 ${조원데이터.length}명 참여\n`;
  전체내용 += '='.repeat(80) + '\n\n';

  조원데이터.forEach((data, index) => {
    전체내용 += `\n${'#'.repeat(80)}\n`;
    전체내용 += `# ${index + 1}. ${data.이름}\n`;
    전체내용 += `${'#'.repeat(80)}\n\n`;

    전체내용 += `📁 제출 파일 (${data.파일목록.length}개):\n`;
    data.파일목록.forEach(file => {
      전체내용 += `  - ${file.이름} (${file.타입})\n`;
    });
    전체내용 += '\n';

    전체내용 += `📖 전체 내용:\n`;
    전체내용 += '-'.repeat(80) + '\n';
    전체내용 += data.내용 + '\n';
    전체내용 += '-'.repeat(80) + '\n\n';
  });

  const fullFileName = `full-content-${dateStr}.txt`;
  const existingFull = folder.getFilesByName(fullFileName);
  while (existingFull.hasNext()) {
    existingFull.next().setTrashed(true);
  }

  const fullFile = folder.createFile(fullFileName, 전체내용, MimeType.PLAIN_TEXT);
  fullFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // 2. 간단 요약 파일 저장
  const summaryFileName = `summary-${dateStr}.txt`;
  const existingSummary = folder.getFilesByName(summaryFileName);
  while (existingSummary.hasNext()) {
    existingSummary.next().setTrashed(true);
  }

  const summaryFile = folder.createFile(summaryFileName, 통합다이제스트, MimeType.PLAIN_TEXT);
  summaryFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // 3. JSON 데이터 저장
  const jsonData = {
    date: dateStr,
    generated: new Date().toISOString(),
    summary: 통합다이제스트,
    memberCount: 조원데이터.length,
    members: 조원데이터.map(data => ({
      name: data.이름,
      fileCount: data.파일목록.length,
      files: data.파일목록,
      fullContent: data.내용
    }))
  };

  const jsonFileName = `digest-${dateStr}.json`;
  const existingJson = folder.getFilesByName(jsonFileName);
  while (existingJson.hasNext()) {
    existingJson.next().setTrashed(true);
  }

  const jsonFile = folder.createFile(
    jsonFileName,
    JSON.stringify(jsonData, null, 2),
    MimeType.PLAIN_TEXT
  );
  jsonFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  Logger.log(`\n파일 저장 완료:`);
  Logger.log(`  - ${fullFileName} (전체 원본 내용)`);
  Logger.log(`  - ${summaryFileName} (간단 요약)`);
  Logger.log(`  - ${jsonFileName} (JSON 데이터)`);
}

/**
 * 수동 실행용: 오늘의 다이제스트 테스트
 */
function AI다이제스트테스트() {
  const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  일일AI다이제스트생성(today);
}

/**
 * 수동 실행용: 어제의 다이제스트
 */
function 어제AI다이제스트생성() {
  일일AI다이제스트생성(); // dateStr 없으면 자동으로 어제
}

/**
 * 🔍 디버깅: 실제 폴더 구조 확인
 * 한 조원의 폴더 안에 어떤 하위 폴더들이 있는지 확인
 */
function 실제폴더구조확인() {
  // 첫 번째 조원의 폴더 ID 가져오기
  const firstMember = Object.entries(CONFIG.MEMBERS)[0];
  const memberName = firstMember[0];
  const folderIdOrArray = firstMember[1];
  const folderId = Array.isArray(folderIdOrArray) ? folderIdOrArray[0] : folderIdOrArray;

  Logger.log(`=== ${memberName} 폴더 구조 확인 ===`);
  Logger.log(`폴더 ID: ${folderId}\n`);

  try {
    const memberFolder = DriveApp.getFolderById(folderId);
    Logger.log(`📁 조원 폴더: ${memberFolder.getName()}`);
    Logger.log(`\n하위 폴더 목록:`);

    const subFolders = memberFolder.getFolders();
    let count = 0;

    while (subFolders.hasNext() && count < 20) {  // 최대 20개만 출력
      const folder = subFolders.next();
      const folderName = folder.getName();

      Logger.log(`  ${count + 1}. ${folderName}`);

      // 첫 번째 하위 폴더의 내부도 확인
      if (count === 0) {
        Logger.log(`     └─ ${folderName} 안의 하위 폴더:`);
        const subSubFolders = folder.getFolders();
        let subCount = 0;
        while (subSubFolders.hasNext() && subCount < 10) {
          const subFolder = subSubFolders.next();
          Logger.log(`        ${subCount + 1}. ${subFolder.getName()}`);
          subCount++;
        }
      }

      count++;
    }

    if (count === 0) {
      Logger.log(`  ❌ 하위 폴더가 없습니다!`);
    } else {
      Logger.log(`\n총 ${count}개의 하위 폴더가 있습니다.`);
    }

  } catch (e) {
    Logger.log(`❌ 오류 발생: ${e.message}`);
    Logger.log(e.stack);
  }
}

/**
 * 🔍 다이제스트 저장 폴더 확인
 * 다이제스트 파일이 저장되는 폴더의 이름과 URL을 출력
 */
function AI저장폴더확인() {
  try {
    const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);
    const folderName = folder.getName();
    const folderUrl = folder.getUrl();

    Logger.log(`=== 다이제스트 저장 위치 ===`);
    Logger.log(`폴더명: ${folderName}`);
    Logger.log(`폴더 URL: ${folderUrl}`);
    Logger.log(`폴더 ID: ${CONFIG.JSON_FOLDER_ID}`);

    Logger.log(`\n최근 생성된 다이제스트 파일:`);
    const files = folder.getFiles();
    const digestFiles = [];

    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      if (fileName.startsWith('full-content-') || fileName.startsWith('summary-') || fileName.startsWith('digest-')) {
        digestFiles.push({
          name: fileName,
          date: file.getLastUpdated(),
          url: file.getUrl()
        });
      }
    }

    // 날짜 순으로 정렬
    digestFiles.sort((a, b) => b.date - a.date);

    // 최근 10개만 출력
    digestFiles.slice(0, 10).forEach((file, index) => {
      Logger.log(`  ${index + 1}. ${file.name}`);
      Logger.log(`     수정일: ${Utilities.formatDate(file.date, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')}`);
      Logger.log(`     URL: ${file.url}`);
    });

    if (digestFiles.length === 0) {
      Logger.log(`  ❌ 다이제스트 파일이 없습니다.`);
    }

  } catch (e) {
    Logger.log(`❌ 오류 발생: ${e.message}`);
    Logger.log(e.stack);
  }
}
