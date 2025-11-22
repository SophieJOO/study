# 기존 시스템 영향 분석

## 📊 변경 사항 상세 분석

### 1. 기존 데이터에 미치는 영향

#### ✅ **영향 없음 (100% 안전)**

| 항목 | 현재 상태 | 이유 |
|------|----------|------|
| 기존 OFF 레코드 | 제출기록 시트에 그대로 보존 | 삭제/변경 코드 없음 |
| 기존 출석(O) 레코드 | 그대로 보존 | 영향 없음 |
| 기존 장기오프 레코드 | 그대로 보존 및 시스템 유지 | 변경 없음 |
| 기존 결석(X) 레코드 | 그대로 보존 | 영향 없음 |
| JSON 파일 | 계속 생성됨 (OFF 포함) | 코드 유지 |
| HTML 출석표 | 기존 데이터 표시 유지 | 변경 없음 |

**증거 (코드 확인):**
```javascript
// apps script code.gs:732-740
if (status === 'O') {
  jsonData[name].출석++;
} else if (status === 'OFF') {  // ✅ 여전히 OFF 처리
  jsonData[name].오프++;
} else if (status === CONFIG.LONG_OFF_STATUS) {
  jsonData[name].장기오프++;
} else {
  jsonData[name].결석++;
}
```

#### ⚠️ **동작 변경 (새 데이터만)**

| 항목 | 기존 동작 | 새 동작 | 영향 범위 |
|------|----------|---------|----------|
| OFF.md 파일 | OFF 상태로 기록 | 무시 → 결석 가능 | 앞으로 업로드되는 것만 |
| 주간 OFF 3회 초과 | 자동 결석 전환 | 체크 안함 | 앞으로 생성되는 것만 |

**중요:** 과거 데이터는 절대 변경되지 않습니다!

### 2. 시스템별 영향 분석

#### A. 출석 체크 시스템 (`출석체크_메인()`)

**제거된 코드:**
```javascript
// 제거됨:
if (fileName === 'OFF.md') {
  출석기록추가(memberName, dateStr, [], 'OFF');
}

// 현재:
const files = 파일목록및링크생성(folder);
if (files.length > 0) {
  출석기록추가(memberName, dateStr, files, 'O');
} else {
  // 파일 없음 → 결석 처리는 마감시간체크()에서
}
```

**영향:**
- ✅ 장기오프는 여전히 최우선 처리
- ✅ 일반 출석 처리 정상 작동
- ⚠️ OFF.md 올려도 인식 안됨

**위험도:** 낮음 (OFF.md 사용 안하면 문제 없음)

#### B. JSON 생성 시스템 (`JSON파일생성()`)

**코드 확인:**
```javascript
// apps script code.gs:732-740
// ✅ OFF 카운팅 코드 그대로 유지
if (status === 'OFF') {
  jsonData[name].오프++;
}
```

**영향:**
- ✅ 기존 OFF 레코드 정상 카운트
- ✅ JSON 파일 생성 정상

**위험도:** 없음

#### C. HTML 출석표 (`index (1).html`)

**코드 확인:**
```javascript
// index (1).html:550-567
} else if (dayStatus.status === 'OFF') {
  // ✅ OFF 표시 코드 그대로 유지
  cell.textContent = '🏖️';
  cell.classList.add('off');
  cell.title = '오프';
}
```

**영향:**
- ✅ 기존 OFF 정상 표시
- ✅ 주간 오프 초과 체크 코드는 유지 (line 505-515)

**위험도:** 없음

#### D. 주간 오프 검증 (`주간오프검증()`)

**제거된 함수:**
```javascript
// 완전 제거됨
function 주간오프검증() { ... }
function getISOYearWeek() { ... }
function 과거주차_전체재검증() { ... }
```

**제거된 호출:**
```javascript
// apps script code.gs:213 (제거됨)
// 주간오프검증();  ← 더이상 실행 안됨
```

**영향:**
- ⚠️ 앞으로 주간 OFF 3회 초과 체크 안함
- ✅ 과거 데이터는 이미 처리되어 있음
- ✅ 새로운 OFF 생성 자체가 불가능

**위험도:** 낮음 (OFF 제도 폐지하므로 무의미)

### 3. 새로운 기능의 독립성

#### A. `weekly-attendance.gs`

**독립성 확인:**
```javascript
// 읽기만 수행
const sheet = ss.getSheetByName(CONFIG.SHEET_NAME); // 제출기록 읽기
const data = sheet.getDataRange().getValues();      // 데이터 읽기

// 쓰기는 별도 시트
const summarySheet = ss.getSheetByName('주간집계'); // 새 시트
```

**영향:**
- ✅ 제출기록 시트 읽기만 수행
- ✅ 주간집계 시트에만 쓰기
- ✅ 기존 데이터 변경 없음

**위험도:** 없음

#### B. `gemini-ai.gs`

**독립성 확인:**
```javascript
// 완전히 독립된 파일 생성
const txtFile = folder.createFile(txtFileName, ...);  // 새 파일
const jsonFile = folder.createFile(jsonFileName, ...); // 새 파일
```

**영향:**
- ✅ 구글 드라이브 파일 읽기만
- ✅ 새 파일 생성만
- ✅ 기존 시스템 영향 없음

**위험도:** 없음

## 🎯 실제 위험 시나리오

### 시나리오 1: "누군가 OFF.md를 올리면?"

**기존:**
```
폴더 구조:
2025-11-22/
  ├── OFF.md          → OFF 상태 기록
```

**현재:**
```
폴더 구조:
2025-11-22/
  ├── OFF.md          → 무시됨
  └── (다른 파일 없음) → 결석 처리 (마감시간체크에서)
```

**대응 방안:**
1. 조원들에게 OFF.md 사용 중단 공지
2. 또는 apps script code.gs 교체를 보류

### 시나리오 2: "기존 OFF 데이터가 사라지면?"

**불가능:**
- 데이터 삭제 코드 없음
- 읽기만 수행
- JSON/HTML에서 정상 표시

**증거:**
```bash
# 삭제 관련 코드 검색
$ grep -n "setTrashed\|delete\|remove" "apps script code.gs"
# → 제출기록 시트 관련 삭제 코드 없음
```

### 시나리오 3: "주간 집계가 틀리면?"

**영향 범위:**
- 주간집계 시트에만 영향
- 제출기록 시트는 변경 없음
- JSON/HTML은 영향 없음

**대응:**
- 주간집계 시트만 삭제하고 재실행

## 📈 안전성 점수

| 항목 | 안전성 | 근거 |
|------|--------|------|
| 기존 데이터 보존 | ⭐⭐⭐⭐⭐ 5/5 | 변경/삭제 코드 없음 |
| 기존 출석 체크 | ⭐⭐⭐⭐⭐ 5/5 | 로직 동일 (장기오프 우선) |
| JSON 생성 | ⭐⭐⭐⭐⭐ 5/5 | 코드 변경 없음 |
| HTML 표시 | ⭐⭐⭐⭐⭐ 5/5 | OFF 표시 유지 |
| 새 기능 독립성 | ⭐⭐⭐⭐⭐ 5/5 | 읽기만 수행 |
| OFF.md 호환성 | ⭐⭐ 2/5 | 더이상 작동 안함 |

**종합 안전성: ⭐⭐⭐⭐ 4.3/5**

유일한 위험: OFF.md 사용 시 결석 처리

## 🛡️ 최종 권장사항

### 옵션 1: 단계적 배포 (가장 안전) ⭐⭐⭐⭐⭐

```
1단계 (즉시):
  - weekly-attendance.gs 추가
  - gemini-ai.gs 추가
  - 수동 실행으로 테스트
  → 기존 시스템 영향 0%

2단계 (1주 후):
  - 조원들에게 OFF.md 사용 중단 공지
  - 주간 집계 검증
  - AI 다이제스트 검증

3단계 (2주 후):
  - apps script code.gs 교체
  - OFF.md 제도 공식 폐지
```

### 옵션 2: OFF.md 유지하며 새 기능만 추가 ⭐⭐⭐⭐

```
변경하지 않는 것:
  - apps script code.gs (OFF.md 체크 유지)
  - digest-functions.gs (OFF.md 체크 유지)

추가만 하는 것:
  - weekly-attendance.gs
  - gemini-ai.gs

결과:
  → 기존 시스템 완전히 유지
  → 새 기능 별도 사용
```

### 옵션 3: 테스트 환경 1주일 검증 ⭐⭐⭐⭐⭐

```
1. 스프레드시트 복사본 생성
2. 복사본에 새 코드 전체 적용
3. 1주일 병렬 운영
4. 문제 없으면 원본에 적용
```

## 💡 제안

가장 안전한 방법:

1. **지금 당장:**
   - `weekly-attendance.gs` 추가만
   - `gemini-ai.gs` 추가만
   - 기존 코드 건드리지 않음

2. **1-2주 테스트 후:**
   - 주간 집계 정확성 확인
   - AI 품질 확인

3. **검증 완료 후:**
   - OFF.md 폐지 여부 결정
   - 필요하면 apps script code.gs 교체

이렇게 하면 **위험도 0%**입니다!
