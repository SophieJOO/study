# 🕐 Windows Task Scheduler 설정 가이드

매일 새벽 4시에 자동으로 스터디 요약 이미지를 생성하도록 설정합니다.

---

## 🚀 빠른 설정 (PowerShell 명령어)

PowerShell을 **관리자 권한**으로 실행한 후 아래 명령어를 복사하여 실행하세요:

```powershell
$action = New-ScheduledTaskAction -Execute "C:\Users\User\study\study_summary\run_daily.bat" -WorkingDirectory "C:\Users\User\study\study_summary"
$trigger = New-ScheduledTaskTrigger -Daily -At 4:00AM
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -WakeToRun
Register-ScheduledTask -TaskName "StudySummaryAutomation" -Action $action -Trigger $trigger -Settings $settings -Description "매일 스터디 인증 요약 이미지 생성"
```

---

## 📋 수동 설정 방법 (GUI)

### 1단계: Task Scheduler 열기
```
Win + R → taskschd.msc → Enter
```

### 2단계: 새 작업 만들기
1. 왼쪽 패널에서 "작업 스케줄러 라이브러리" 클릭
2. 오른쪽 패널에서 "기본 작업 만들기..." 클릭

### 3단계: 기본 정보 입력
```
이름: StudySummaryAutomation
설명: 매일 스터디 인증 요약 이미지 생성
```

### 4단계: 트리거 설정
```
시작: "매일" 선택
시작 시간: 오전 4:00:00
매 1일마다: 선택됨
```

### 5단계: 동작 설정
```
동작: "프로그램 시작"
프로그램/스크립트: C:\Users\User\study\study_summary\run_daily.bat
시작 위치: C:\Users\User\study\study_summary
```

### 6단계: 추가 설정 (권장)
작업을 만든 후, 해당 작업을 더블클릭하여 속성을 열고:

**일반 탭:**
- ✅ "가장 높은 권한으로 실행" 체크

**조건 탭:**
- ❌ "컴퓨터가 AC 전원에 연결된 경우에만 작업 시작" 해제
- ✅ "이 작업을 실행하기 위해 컴퓨터의 절전 모드 해제" 체크

**설정 탭:**
- ✅ "예약된 시작 시간을 놓친 경우 가능한 빨리 작업 실행" 체크

---

## ✅ 설정 확인

### 작업 목록에서 확인
```
1. Task Scheduler 열기
2. "작업 스케줄러 라이브러리" 클릭
3. "StudySummaryAutomation" 작업 확인
4. 상태가 "준비"로 표시되어야 함
```

### 수동 테스트 실행
```
1. 작업 우클릭 → "실행"
2. logs 폴더에서 최신 로그 파일 확인
3. output 폴더에서 생성된 이미지 확인
```

---

## 🔧 문제 해결

### 작업이 실행되지 않는 경우
1. **컴퓨터가 켜져 있는지 확인** - 절전 모드에서도 실행하려면 설정 필요
2. **배치 파일 경로 확인** - 경로에 한글이나 공백이 없는지 확인
3. **Python 환경 확인** - python 명령이 PATH에 있는지 확인

### 로그 확인
```
C:\Users\User\study\study_summary\logs\run_YYYYMMDD_HHMMSS.log
```

---

## 🗑️ 작업 삭제

더 이상 필요 없으면:
```powershell
Unregister-ScheduledTask -TaskName "StudySummaryAutomation" -Confirm:$false
```
