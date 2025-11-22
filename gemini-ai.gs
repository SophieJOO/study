/**
 * ==================== Gemini AI 통합 ====================
 * 공부 내용 자동 요약 및 질 평가 시스템
 */

// Gemini API 설정
const GEMINI_CONFIG = {
  // API 키는 스크립트 속성에 저장 (보안)
  // 설정 방법: 프로젝트 설정 > 스크립트 속성 > GEMINI_API_KEY 추가
  getApiKey: function() {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) {
      throw new Error('GEMINI_API_KEY가 설정되지 않았습니다. 프로젝트 설정 > 스크립트 속성에서 추가하세요.');
    }
    return apiKey;
  },

  API_URL: 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent',

  // 요청 설정
  temperature: 0.7,
  maxOutputTokens: 3000
};

/**
 * Gemini API 호출
 * @param {string} prompt - 프롬프트
 * @returns {string} AI 응답
 */
function GeminiAPI호출(prompt) {
  const apiKey = GEMINI_CONFIG.getApiKey();

  const payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }],
    generationConfig: {
      temperature: GEMINI_CONFIG.temperature,
      maxOutputTokens: GEMINI_CONFIG.maxOutputTokens,
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(
      `${GEMINI_CONFIG.API_URL}?key=${apiKey}`,
      options
    );

    const json = JSON.parse(response.getContentText());

    if (!response.getResponseCode() === 200) {
      throw new Error(`API 오류: ${response.getResponseCode()} - ${json.error?.message || '알 수 없는 오류'}`);
    }

    const result = json.candidates?.[0]?.content?.parts?.[0]?.text;

    if (!result) {
      throw new Error('AI 응답이 비어있습니다.');
    }

    return result;

  } catch (e) {
    Logger.log(`Gemini API 호출 실패: ${e.message}`);
    throw e;
  }
}

/**
 * 개별 조원의 공부 내용 요약 및 평가
 * @param {string} memberName - 조원 이름
 * @param {string} 내용 - 공부 내용 (마크다운, 텍스트 등)
 * @param {Array} 파일목록 - 파일 정보 [{이름, 타입}, ...]
 * @returns {Object} {요약, 핵심키워드, 질평가점수, 질평가코멘트}
 */
function AI개별요약및평가(memberName, 내용, 파일목록 = []) {
  const 파일정보 = 파일목록.length > 0
    ? `\n\n[제출 파일]\n${파일목록.map(f => `- ${f.이름} (${f.타입})`).join('\n')}`
    : '';

  const prompt = `
당신은 한의학 전문가이자 학습 코치입니다.

아래는 "${memberName}" 학생이 오늘 공부한 내용입니다:

${내용}${파일정보}

위 내용을 분석하여 다음 형식으로 응답해주세요:

## 요약
[2-3문장으로 핵심 내용 요약]

## 핵심 키워드
[3-5개의 핵심 키워드, 쉼표로 구분]

## 학습 질 평가
점수: [1-10점]
평가: [학습의 깊이, 이해도, 체계성을 1-2문장으로 평가]

규칙:
1. 요약은 전문용어를 정확하게 사용하되 간결하게
2. 키워드는 한의학 핵심 개념 위주로
3. 질 평가 기준:
   - 10점: 매우 심도있는 학습, 체계적 정리, 응용/분석 포함
   - 7-9점: 충실한 학습, 핵심 개념 정리
   - 4-6점: 기본적 학습, 단순 암기 위주
   - 1-3점: 매우 형식적이거나 내용 빈약
4. 평가는 건설적이고 격려하는 톤으로
`;

  try {
    const 응답 = GeminiAPI호출(prompt);

    // 응답 파싱
    const 요약매치 = 응답.match(/##\s*요약\s*\n([\s\S]*?)(?=\n##|$)/);
    const 키워드매치 = 응답.match(/##\s*핵심\s*키워드\s*\n([\s\S]*?)(?=\n##|$)/);
    const 점수매치 = 응답.match(/점수:\s*(\d+)/);
    const 평가매치 = 응답.match(/평가:\s*([\s\S]*?)(?=\n##|$)/);

    return {
      요약: 요약매치 ? 요약매치[1].trim() : 내용.substring(0, 200) + '...',
      핵심키워드: 키워드매치 ? 키워드매치[1].trim() : '',
      질평가점수: 점수매치 ? parseInt(점수매치[1]) : 5,
      질평가코멘트: 평가매치 ? 평가매치[1].trim() : '평가 없음',
      AI처리완료: true
    };

  } catch (e) {
    Logger.log(`AI 요약 실패 (${memberName}): ${e.message}`);

    // AI 실패 시 기본 요약
    return {
      요약: 내용.substring(0, 200) + '...',
      핵심키워드: '',
      질평가점수: null,
      질평가코멘트: 'AI 처리 실패',
      AI처리완료: false
    };
  }
}

/**
 * 전체 조원 통합 AI 다이제스트 생성
 * @param {Array} 조원데이터 - [{이름, 내용, 파일목록, AI평가}, ...]
 * @param {string} dateStr - 날짜 (yyyy-MM-dd)
 * @returns {string} 카톡 공유용 통합 다이제스트
 */
function AI통합다이제스트생성(조원데이터, dateStr) {
  // 조원 요약 생성
  const 조원요약들 = 조원데이터.map(data => `
**${data.이름}** (질 평가: ${data.AI평가?.질평가점수 || '?'}/10점)
${data.AI평가?.요약 || data.내용.substring(0, 100)}
핵심: ${data.AI평가?.핵심키워드 || '키워드 없음'}
${data.AI평가?.질평가코멘트 ? `💡 ${data.AI평가.질평가코멘트}` : ''}
  `.trim()).join('\n\n---\n\n');

  const prompt = `
당신은 한의학 스터디 그룹의 학습 큐레이터입니다.

오늘(${dateStr}) 조원들이 공부한 내용은 다음과 같습니다:

${조원요약들}

위 내용을 바탕으로 카카오톡에 공유할 "오늘의 스터디 다이제스트"를 작성해주세요.

형식:
📚 오늘의 스터디 다이제스트 (${dateStr})

[각 조원별로 2-3줄 요약, 이모지 활용]

━━━━━━━━━━━━━━━━
📊 오늘의 학습 통계
- 출석: N명
- 평균 학습 질: X.X/10점
- 최고 점수: [이름] (10점)

💡 오늘의 하이라이트
[가장 인상적이었던 학습 내용이나 공통 주제]

🔗 자세히 보기: [링크]

규칙:
1. 조원별 요약은 핵심만 간결하게
2. 이모지를 적절히 활용하여 가독성 향상
3. 통계는 정확하게 계산
4. 하이라이트는 공통 주제나 특별히 잘한 점 강조
5. 격려하고 동기부여하는 톤
6. 카톡 메시지로 바로 사용 가능하도록
`;

  try {
    const 다이제스트 = GeminiAPI호출(prompt);

    return 다이제스트;

  } catch (e) {
    Logger.log(`통합 다이제스트 생성 실패: ${e.message}`);

    // 실패 시 간단한 텍스트 생성
    let fallback = `📚 오늘의 스터디 다이제스트 (${dateStr})\n\n`;

    for (const data of 조원데이터) {
      fallback += `✅ ${data.이름}\n`;
      fallback += `   ${data.AI평가?.요약 || data.내용.substring(0, 80)}\n`;
      if (data.AI평가?.핵심키워드) {
        fallback += `   핵심: ${data.AI평가.핵심키워드}\n`;
      }
      fallback += `\n`;
    }

    fallback += `━━━━━━━━━━━━━━━━\n`;
    fallback += `📊 출석: ${조원데이터.length}명\n`;

    return fallback;
  }
}

/**
 * 일일 AI 다이제스트 메인 함수
 * @param {string} dateStr - 날짜 (yyyy-MM-dd), 미지정 시 어제
 */
function 일일AI다이제스트생성(dateStr) {
  if (!dateStr) {
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    dateStr = Utilities.formatDate(yesterday, 'Asia/Seoul', 'yyyy-MM-dd');
  }

  Logger.log(`=== ${dateStr} AI 다이제스트 생성 시작 ===`);

  const 조원데이터 = [];

  // 각 조원의 파일 수집 및 AI 분석
  for (const [memberName, folderIdOrArray] of Object.entries(CONFIG.MEMBERS)) {
    const folderIds = Array.isArray(folderIdOrArray) ? folderIdOrArray : [folderIdOrArray];

    for (const folderId of folderIds) {
      const content = 파일내용수집(memberName, folderId, dateStr);

      if (content && content.내용) {
        Logger.log(`\n${memberName} AI 분석 중...`);

        // AI 요약 및 평가
        const AI평가 = AI개별요약및평가(memberName, content.내용, content.파일목록);

        Logger.log(`  요약: ${AI평가.요약.substring(0, 50)}...`);
        Logger.log(`  질 평가: ${AI평가.질평가점수}/10점`);
        Logger.log(`  키워드: ${AI평가.핵심키워드}`);

        조원데이터.push({
          이름: memberName,
          내용: content.내용,
          파일목록: content.파일목록,
          AI평가
        });

        break;
      }
    }
  }

  if (조원데이터.length === 0) {
    Logger.log('어제 공부한 조원이 없습니다.');
    return null;
  }

  Logger.log(`\n=== 통합 다이제스트 생성 중... ===`);

  // 통합 다이제스트 생성
  const 통합다이제스트 = AI통합다이제스트생성(조원데이터, dateStr);

  Logger.log('\n=== AI 다이제스트 완성 ===');
  Logger.log(통합다이제스트);

  // 결과 저장
  AI다이제스트저장(통합다이제스트, 조원데이터, dateStr);

  return 통합다이제스트;
}

/**
 * 파일 내용 수집 (기존 함수와 유사하지만 더 상세)
 * @param {string} memberName - 조원 이름
 * @param {string} folderId - 폴더 ID
 * @param {string} dateStr - 날짜
 * @returns {Object} {내용, 파일목록}
 */
function 파일내용수집(memberName, folderId, dateStr) {
  try {
    const mainFolder = DriveApp.getFolderById(folderId);
    const subfolders = mainFolder.getFolders();

    // 날짜 폴더 찾기
    let targetFolder = null;
    while (subfolders.hasNext()) {
      const folder = subfolders.next();
      const folderName = folder.getName().trim();
      const dateInfo = 날짜추출(folderName);

      if (dateInfo && dateInfo.dateStr === dateStr) {
        targetFolder = folder;
        break;
      }
    }

    if (!targetFolder) return null;

    let 전체내용 = '';
    const 파일목록 = [];
    const files = targetFolder.getFiles();

    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      const mimeType = file.getMimeType();

      // 마크다운 파일
      if (fileName.toLowerCase().endsWith('.md')) {
        try {
          const mdContent = file.getBlob().getDataAsString('UTF-8');

          // 제목 추출
          const titleMatch = mdContent.match(/^#\s+(.+)$/m);
          if (titleMatch) {
            전체내용 += `[제목: ${titleMatch[1]}]\n\n`;
          }

          전체내용 += mdContent + '\n\n';

          파일목록.push({
            이름: fileName,
            타입: 'Markdown'
          });

        } catch (e) {
          Logger.log(`MD 파일 읽기 실패: ${fileName}`);
        }
      }

      // PDF (파일명만 - OCR은 별도 구현 필요)
      else if (mimeType === MimeType.PDF) {
        전체내용 += `[PDF 문서: ${fileName}]\n`;
        파일목록.push({
          이름: fileName,
          타입: 'PDF'
        });
      }

      // 이미지
      else if (mimeType.startsWith('image/')) {
        전체내용 += `[이미지: ${fileName}]\n`;
        파일목록.push({
          이름: fileName,
          타입: 'Image'
        });
      }
    }

    return 전체내용 ? { 내용: 전체내용, 파일목록 } : null;

  } catch (e) {
    Logger.log(`${memberName} 파일 수집 실패: ${e.message}`);
    return null;
  }
}

/**
 * AI 다이제스트 저장
 * @param {string} 통합다이제스트 - 통합 다이제스트 텍스트
 * @param {Array} 조원데이터 - 조원별 상세 데이터
 * @param {string} dateStr - 날짜
 */
function AI다이제스트저장(통합다이제스트, 조원데이터, dateStr) {
  const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);

  // 1. 통합 다이제스트 텍스트 파일 저장
  const txtFileName = `ai-digest-${dateStr}.txt`;
  const existingTxt = folder.getFilesByName(txtFileName);
  while (existingTxt.hasNext()) {
    existingTxt.next().setTrashed(true);
  }

  const txtFile = folder.createFile(txtFileName, 통합다이제스트, MimeType.PLAIN_TEXT);
  txtFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // 2. 상세 데이터 JSON 저장
  const jsonData = {
    date: dateStr,
    generated: new Date().toISOString(),
    summary: 통합다이제스트,
    members: 조원데이터.map(data => ({
      name: data.이름,
      summary: data.AI평가?.요약,
      keywords: data.AI평가?.핵심키워드,
      qualityScore: data.AI평가?.질평가점수,
      qualityComment: data.AI평가?.질평가코멘트,
      files: data.파일목록
    }))
  };

  const jsonFileName = `ai-digest-${dateStr}.json`;
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

  Logger.log(`AI 다이제스트 저장 완료:`);
  Logger.log(`  - ${txtFileName}`);
  Logger.log(`  - ${jsonFileName}`);
}

/**
 * 수동 실행용: 오늘의 AI 다이제스트 테스트
 */
function AI다이제스트테스트() {
  const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  일일AI다이제스트생성(today);
}

/**
 * 수동 실행용: 어제의 AI 다이제스트
 */
function 어제AI다이제스트생성() {
  일일AI다이제스트생성(); // dateStr 없으면 자동으로 어제
}
