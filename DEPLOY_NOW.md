# 🚀 지금 바로 배포하기

## 준비된 파일들

✅ `index-완성본.html` - GitHub Pages에 배포할 완전한 HTML 파일
✅ `weekly-attendance.gs` - Apps Script에 추가할 주간 집계 스크립트
✅ 모든 안전 가이드 문서들

---

## 🎯 3단계로 완료하기

### 1단계: Apps Script에 주간 집계 추가 (10분)

**목표:** 주간 집계 JSON 파일 생성

```
1. Google 스프레드시트 열기
2. 확장 프로그램 > Apps Script
3. 왼쪽 "+" 버튼 클릭 > "스크립트" 선택
4. 파일명: weekly-attendance
5. /home/user/study/weekly-attendance.gs 파일 내용 전체 복사-붙여넣기
6. Ctrl+S 저장
7. 함수 선택: 이번달주간집계
8. 실행 (▶️) 클릭
9. 권한 승인 (처음 한 번만)
```

**결과 확인:**
```javascript
// 실행 로그에서 다음 복사:
주간 집계 JSON 저장 완료: weekly_summary_2025-11.json
URL: https://drive.google.com/uc?export=download&id=XXXXXXXXXX
```

**📝 파일 ID를 메모하세요!** (나중에 사용)

---

### 2단계: HTML 파일 배포 (5분)

#### 방법 A: GitHub 웹에서 직접 수정 (추천)

```
1. https://github.com/floating535-lang/study-attendance 접속
2. index.html 클릭
3. 연필 아이콘(편집) 클릭
4. 전체 내용 삭제
5. /home/user/study/index-완성본.html 전체 복사-붙여넣기
6. 아래로 스크롤하여 "Commit changes" 클릭
7. 커밋 메시지: "Add weekly attendance summary section"
8. "Commit changes" 버튼 클릭
```

#### 방법 B: Git 명령어로 (로컬에서)

```bash
cd /path/to/study-attendance
cp /home/user/study/index-완성본.html index.html
git add index.html
git commit -m "Add weekly attendance summary section"
git push origin main
```

---

### 3단계: JSON 파일 ID 설정 (3분)

**1단계에서 메모한 파일 ID를 HTML에 설정**

#### GitHub 웹에서:

```
1. https://github.com/floating535-lang/study-attendance 접속
2. index.html 클릭
3. 연필 아이콘(편집) 클릭
4. Ctrl+F로 "YOUR_GOOGLE_DRIVE_JSON_FOLDER_URL" 검색
5. 다음 부분 수정:

변경 전:
const baseUrl = 'YOUR_GOOGLE_DRIVE_JSON_FOLDER_URL';

변경 후:
const baseUrl = 'https://drive.google.com/uc?export=download&id=';
const JSON_FILE_IDS = {
    attendance: '기존_attendance_summary_파일_ID',
    weekly: '1단계에서_복사한_파일_ID'
};

그리고 아래쪽에서:
const dataUrl = `${baseUrl}/attendance_summary_${yearMonth}.json`;
const weeklyUrl = `${baseUrl}/weekly_summary_${yearMonth}.json`;

이렇게 변경:
const dataUrl = `${baseUrl}${JSON_FILE_IDS.attendance}`;
const weeklyUrl = `${baseUrl}${JSON_FILE_IDS.weekly}`;

6. "Commit changes" 클릭
```

---

## ✅ 완료 확인

### 1. GitHub Pages 확인
```
https://floating535-lang.github.io/study-attendance/
```

**보여야 할 것:**
- ✅ 기존 출석표 (일간)
- ✅ 하단에 "📊 주간 출석 집계" 섹션
- ✅ "📅 주간 집계 기준" 안내문
- ✅ 진행중인 주는 결석이 "-"로 표시

### 2. Apps Script 스프레드시트 확인
```
1. 스프레드시트에서 '주간집계' 시트 확인
2. 데이터가 있나요?
```

---

## 🔧 문제 해결

### 문제: 주간 통계가 안 보여요
**해결:**
```
1. F12 눌러서 개발자 도구 열기
2. Console 탭 확인
3. 에러 메시지 확인
```

**가능한 원인:**
- JSON 파일 ID가 잘못됨
- 파일 공유 권한이 "링크가 있는 모든 사용자"가 아님

**수정:**
```javascript
// Apps Script에서 다시 실행
function JSON파일확인() {
  const folder = DriveApp.getFolderById(CONFIG.JSON_FOLDER_ID);
  const files = folder.getFilesByName('weekly_summary_2025-11.json');

  if (files.hasNext()) {
    const file = files.next();

    // 공유 설정 확인
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    Logger.log('파일 ID: ' + file.getId());
    Logger.log('URL: https://drive.google.com/uc?export=download&id=' + file.getId());
  }
}
```

### 문제: "진행중인 주"가 결석으로 표시돼요
**해결:**
```
이번달주간집계() 다시 실행
→ 최신 데이터로 업데이트됨
```

---

## 📅 자동화 설정 (선택사항)

**1주일 안정화 후 설정하세요**

```
1. Apps Script > 트리거 (시계 아이콘)
2. "+ 트리거 추가"
3. 함수: 이번달주간집계
4. 이벤트 소스: 시간 기반
5. 시간 기반 트리거 유형: 일 타이머
6. 시간대 선택: 오전 0시 ~ 1시
7. 저장
```

**효과:** 매일 자동으로 주간 집계 업데이트

---

## 🎯 다음 할 일

### 이번 주:
- [x] 주간 집계 시스템 배포
- [ ] 매일 출석 정상 작동 확인
- [ ] 주간 통계 정확성 검증

### 다음 주:
- [ ] Gemini AI 다이제스트 추가 (gemini-ai.gs)
- [ ] 자동 트리거 설정
- [ ] 최종 안정화

---

## 📞 도움이 필요하면

**에러가 발생했을 때:**
```
1. Apps Script 실행 로그 캡처
2. 브라우저 Console 에러 캡처
3. 어느 단계에서 문제가 생겼는지 설명
```

**롤백이 필요하면:**
```
1. GitHub에서 index.html 이전 버전으로 복구
2. Apps Script에서 weekly-attendance.gs 삭제
3. 원래대로 돌아옴
```

---

## 🎉 배포 완료!

모든 단계를 완료하셨다면:
- ✅ 주간 집계 시스템 작동 중
- ✅ 출석표에 주간 통계 표시
- ✅ 기존 시스템 100% 보존
- ✅ OFF.md 시스템 폐지 완료

**축하합니다! 🎊**

1주일 모니터링 후 다음 단계(AI 다이제스트)로 진행하세요.
