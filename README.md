# GD Agent AI Dashboard

## 프로젝트 개요
- **목적**: 60개 게이트(Gate)의 제작 진행 상황을 실시간으로 시각화
- **회사**: CS Wind AM
- **배포 경로**: https://mes-r018-analysis.pages.dev/gd-agent/
- **데이터 소스**: Excel 파일 업로드 (`GD AI Agent sample 1.xlsx`)

---

## Excel 파일 구조

### 기본 정보
- **총 행 수**: 60개 (Gate G01 ~ G60)
- **총 컬럼 수**: 58개

---

## 컬럼 구조 (A~BA열, 58개)

### A~Q: 기본 정보 (17개)

| Excel | 컬럼명 | 의미 | Gate Card 사용 | 예시 |
|-------|--------|------|----------------|------|
| A | `mcn_no` | Gate 번호 (G01~G60) | ✅ 헤더 표시 | G01, G52 |
| B | `serial_no2` | Section ID | ✅ 표시 | VB087-U, VB100-T |
| C | `rev_flag` | 투입 방향 (0=Normal, 1=Reverse) | 파싱 시 고려 | 0, 1 |
| D | `wo_dtl_id` | Work Order ID | 미표시 | M42100000006355708 |
| E | `fo_desc` | 현재 공정 (FU-C02, CSO-C02 등) | ✅ 표시 | FU-C11, CSO-C08 |
| F | `sts` | Status (S/R/H) | 시간 계산 분기 | S, R, H |
| G | `working_rate` | 작업 진행률 (%) | 미정 | 30.0, 50.0 |
| H | `start_dt` | 실제 시작 시간 | ✅ 시간 계산 | 2026-02-19 09:24:00 |
| I | `end_dt` | 실제 종료 시간 | 미사용 | NaN |
| J | `plan_start_dt` | 계획 시작 시간 | ✅ R일 때 사용 | 2026-02-22 09:50:00 |
| K | `plan_end_dt` | 계획 종료 시간 | 의미 없음 | 2026-02-22 11:32:00 |
| L | `work_st` | Standard Time (분) | ✅ 종료 시간 계산 | 102.0, 64.0 |
| M | `worker_id` | 작업자 ID | 미표시 | 8639.0 |
| N | `worker_nm` | 작업자 이름 (이름, 성) | 미정 | BERUMEN, KRIS ELIAS |
| O | `skirt_qty` | Skirt 개수 | ✅ WIP 박스 개수 | 11, 5 |
| P | `proj_color` | 프로젝트 번호 (14, 7) | 미정 | 14, 7 |
| Q | `cur_time` | 데이터 업데이트 시간 | ✅ 대기 시간 계산 | 2026-02-19 09:51:44 |

#### F열: Status 코드 (sts) 상세

Gate의 전체 진행 상태를 나타냅니다.

| 코드 | 의미 | 개수 | 설명 |
|------|------|------|------|
| **S** | Started | 25개 | 작업 시작됨, 진행 중 |
| **R** | Reserved/Ready | 33개 | 예약됨, 작업 대기 |
| **H** | **Holding** | 1개 | **보류 상태** (예: G06 - Flatness Inspection) |
| **NaN** | 정보 없음 | 1개 | Status 정보 없음 |

**Holding 예시 (Gate G06):**
- Section: VB088-U
- Process: Flatness Inspection
- Status: H (Holding)
- Working Rate: 50%
- Worker: COLLINS, TOMMY
- 작업은 시작되었으나 어떤 이유로 보류된 상태

---

### R~AK: Joint Status (20개 컬럼)

- `joint_status1` (R열) ~ `joint_status20` (AK열)
- **중요**: Joint 수 = Skirt 수 + 1 (양쪽 Flange 포함)
  - Joint 1: 시작 Flange - Skirt 1
  - Joint 2~N: Skirt 간 연결
  - Joint N+1: Skirt N - 끝 Flange
- **예시**: Skirt 11개 → Joint 12개 (Joint 1~12)
- **나머지**: B (Blank) 또는 NaN

#### Joint Status 코드 정의

| 약어 | 의미 | 설명 |
|------|------|------|
| **F** | Fit-up | 핏업 (두 Skirt 결합) |
| **D** | Done | 완료 |
| **W** | Working | 작업 중 |
| **I** | Inside welding | 내부 용접 |
| **O** | Outside welding | 외부 용접 |
| **B** | Blank | 빈 값 (해당 Joint 없음) |

#### Joint Status 조합 예시

| 코드 | 의미 | 설명 |
|------|------|------|
| **FW** | Fit-up Working | Fit-up 작업 중 |
| **FD** | Fit-up Done | Fit-up 완료 |
| **FDOW** | FD + Outside Working | Fit-up 완료 + 외부 용접 중 |
| **FDOD** | FD + Outside Done | Fit-up 완료 + 외부 용접 완료 |
| **FDIW** | FD + Inside Working | Fit-up 완료 + 내부 용접 중 |
| **FDID** | FD + Inside Done | Fit-up 완료 + 내부 용접 완료 |
| **FDIWOW** | FD + Inside Working + Outside Working | **불가능** (동시 작업 불가) |
| **FDIDOD** | FD + Inside Done + Outside Done | **모든 용접 완료 (최종 상태)** ✅ |

#### 작업 순서 (Workflow)

```
FW → FD → (FDOW 또는 FDIW) → (FDOD 또는 FDID) → (FDIW 또는 FDOW) → FDIDOD
```

**3-Can 작업 방식 (2 Joints per CAN):**

**Option A - Outside 먼저:**
1. Joint 1: FD → FDOW → FDOD
2. Joint 2: FD → FDOW → FDOD
3. Joint 1: FDOD → FDIW → FDIDOD
4. Joint 2: FDOD → FDIW → FDIDOD

**Option B - Inside 먼저:**
1. Joint 1: FD → FDIW → FDID
2. Joint 2: FD → FDIW → FDID
3. Joint 1: FDID → FDOW → FDIDOD
4. Joint 2: FDID → FDOW → FDIDOD

**이유**: SAW 붐대를 넣었다 뺐다 반복하면 비효율적 → 한 쪽(Outside 또는 Inside)을 먼저 다 하고 반대쪽 진행

#### E열 (fo_desc) 공정 코드

| 코드 | 의미 | 설명 |
|------|------|------|
| **FU-C02** | Fit-up Joint 02 | Joint 2 Fit-up 작업 |
| **CSO-C02** | Cir-Seam Outside Joint 02 | Joint 2 외부 용접 |
| **CSI-C02** | Cir-Seam Inside Joint 02 | Joint 2 내부 용접 |

---

### AL: Plant (1개)
- `plant`: 공장 코드 (예: 4210)

---

### AM~A`: Skirt Status (20개 컬럼)

- `skirt_status1` (AM열) ~ `skirt_status20` (A`열)
- **사용**: Skirt 개수만큼 사용
- **나머지**: B (Blank) 또는 NaN

#### Skirt Status 코드 정의 (빈도순)

| 코드 | 개수 | 의미 | 진행률 바 색상 | Growing Line |
|------|------|------|----------------|--------------|
| **SD** | 275 | Skirt Done (Lseam 완료) | 🟨 노란색 (#F7CD42) | ✅ 준비 완료 |
| **SN\|BN** | 102 | Skirt Not ready, Bending 대기 | ⬜ 회색 (#5a6b78) | ❌ 준비 안 됨 |
| **SW\|FU** | 47 | Skirt Working, Fit-up 대기 (Lseam 완료) | 🟨 노란색 (#F7CD42) | ✅ 준비 완료 |
| **SN\|LS** | 36 | Skirt Not ready, Lseam 대기 | ⬜ 회색 (#5a6b78) | ❌ 준비 안 됨 |
| **SN\|BV** | 11 | Skirt Not ready, Beveling 대기 | ⬜ 회색 (#5a6b78) | ❌ 준비 안 됨 |
| **SN\|PB** | 8 | Skirt Not ready, Pre-blast 대기 | ⬜ 회색 (#5a6b78) | ❌ 준비 안 됨 |
| **SW\|ROK** | 6 | Skirt Working, Re-rolling OK (Lseam 후 재 Bending) | 🟨 노란색 (#F7CD42) | ✅ 준비 완료 |
| **SW\|FFU** | 4 | Skirt Working, Flange Fit-up 대기 (Lseam 완료) | 🟨 노란색 (#F7CD42) | ✅ 준비 완료 |
| **SN\|CT** | 3 | Skirt Not ready, Cutting 대기 | ⬜ 회색 (#5a6b78) | ❌ 준비 안 됨 |
| **SW\|FCS** | 3 | Skirt Working, Flange Cir-seam 대기 (Lseam 완료) | 🟨 노란색 (#F7CD42) | ✅ 준비 완료 |
| **B** | 645 | Blank (빈 값) | - | - |

#### Prefix 의미

- **S** = Skirt
- **SW** = Skirt Working (작업 가능/완료, Lseam 완료) → 🟨 **노란색**
- **SN** = Skirt Not ready (대기 중, Lseam 전) → ⬜ **회색**

#### Suffix 의미 (공정 단계)

| Suffix | 의미 | 설명 |
|--------|------|------|
| **SD** | Skirt Done | Lseam 완료, Growing Line 위에 위치 |
| **CT** | Cutting | 절단 |
| **PB** | Pre-blast | 전처리 블라스트 |
| **BV** | Beveling | 개선 (용접 준비) |
| **BN** | Bending | 벤딩 (원형 성형) |
| **LS** | Lseam | 세로 용접 |
| **FU** | Fit-up | Fit-up 대기 (WIP 또는 라인 위) |
| **ROK** | Re-rolling OK | Lseam 후 Ovality 초과 → 재 Bending 완료 |
| **FFU** | Flange Fit-up | Flange Fit-up 대기 (끝단 Skirt만) |
| **FCS** | Flange Cir-seam | Flange Cir-seam 대기 (끝단 Skirt만) |

#### BT 공정 흐름 (Pre-Growing 단계)

```
SN|CT → SN|PB → SN|BV → SN|BN → SN|LS → SD (또는 SW|FU, SW|ROK)
                                           ↓
                                    Growing Line 투입
```

**끝단 Skirt (Flange 포함):**
```
... → SW|FFU → SW|FCS → SD → Growing Line 투입
```

---

## Skirt-Joint 관계 및 매핑

### 기본 구조

```
[Flange] ─ [Joint 1] ─ [Skirt 1] ─ [Joint 2] ─ [Skirt 2] ─ [Joint 3] ─ ... ─ [Joint N] ─ [Skirt N] ─ [Joint N+1] ─ [Flange]
```

**중요:**
- **Joint 1 존재함!** (시작 Flange - Skirt 1 연결)
- **Joint 수 = Skirt 수 + 1**
- 예: Skirt 11개 → Joint 12개 (Joint 1~12)
- Joint 1: Flange - Skirt 1
- Joint 2~11: Skirt 간 연결 (10개)
- Joint 12: Skirt 11 - Flange

### Skirt-Joint 매핑 규칙

**6개 Skirt, 7개 Joint 예시:**

```
[Flange] ─ [J1] ─ [S1] ─ [J2] ─ [S2] ─ [J3] ─ [S3] ─ [J4] ─ [S4] ─ [J5] ─ [S5] ─ [J6] ─ [S6] ─ [J7] ─ [Flange]
          index0  index0  index1  index1  index2  index2  index3  index3  index4  index4  index5  index5  index6
```

**각 Skirt는 양쪽 Joint를 확인:**
- **Skirt 1** (index 0): Joint 1 (index 0), Joint 2 (index 1)
- **Skirt 2** (index 1): Joint 2 (index 1), Joint 3 (index 2)
- **Skirt 3** (index 2): Joint 3 (index 2), Joint 4 (index 3)
- ...
- **Skirt 6** (index 5): Joint 6 (index 5), Joint 7 (index 6)

**JavaScript 매핑:**
```javascript
// 각 Skirt는 양쪽 Joint 확인
const leftJoint = jointStatuses[skirtIndex];      // 왼쪽 Joint
const rightJoint = jointStatuses[skirtIndex + 1]; // 오른쪽 Joint
```

### Joint 번호 규칙

- **Joint No. = Section 기준 연속 번호 (01~N+1)**
- 예: 11개 Can Section → Joint 12개 → 01~12
- **FU-01** = Joint 1 (시작 Flange - Skirt 1)
- **FU-02** = Joint 2 (Skirt 1 - Skirt 2)
- **FU-03** = Joint 3 (Skirt 2 - Skirt 3)
- ...
- **FU-12** = Joint 12 (Skirt 11 - 끝 Flange)
- MES 입력 시 Skirt ID 매핑: **FU-xx ↔ Txx** (동일 번호)

---

## Rev_flag - Section 투입 방향 (C열)

`rev_flag`는 Section이 Growing Line에 투입되는 방향을 나타냅니다.

### Rev_flag 값

| Rev_flag | 투입 방향 | 설명 |
|----------|-----------|------|
| **0** | Normal (정방향) | 두꺼운 쪽부터 투입, Joint No와 Skirt No 같은 방향 증가 |
| **1** | **Reverse (역방향)** | **얇은 쪽부터 투입**, Joint No와 Skirt No **반대 방향** |

### Normal 투입 (Rev_flag = 0)

```
Gate door ← [Skirt 1(두꺼움)] ─ [Joint 2] ─ [Skirt 2] ─ [Joint 3] ─ ... ─ [Skirt 11(얇음)]
```

- Section은 정방향으로 Growing Line에 투입
- 두꺼운 Skirt(하단)부터 투입
- Joint No와 Skirt No가 **같은 방향**으로 증가

### Reverse 투입 (Rev_flag = 1)

```
Gate door ← [Skirt 11(얇음)] ─ [Joint 11] ─ [Skirt 10] ─ [Joint 10] ─ ... ─ [Skirt 1(두꺼움)]
```

- Section은 역방향으로 Growing Line에 투입
- **얇은 Skirt(상단)부터 투입** (Gate door에 가까운 쪽이 얇은 Skirt)
- **Joint 번호가 높은 것부터 작업** (Joint 11 → Joint 10 → ...)
- **표시 순서는 좌→우 동일** (UI에서는 좌측부터 우측으로 표시)
- **FU-C11** = 실제로는 맨 왼쪽 Joint (Skirt 11과 Skirt 10 사이)

**파싱 로직:**
```javascript
// Excel에서 읽은 Joint Status 배열 (20개)
let jointStatuses = ["FDIDOD", "FD", "FD", "B", "B", ...]; // 예시

// Rev_flag = 1이면 Joint 순서를 뒤집음
if (rev_flag === 1) {
  jointStatuses.reverse(); // 전체 배열 뒤집기
}

// 사용 예:
// Normal: Joint 1~12 (왼쪽부터 오른쪽)
// Reverse: Joint 12~1 (왼쪽부터 오른쪽, 하지만 번호는 역순)
```

---

## Gate Card 색상 규칙 (4가지 색상)

### 진행률 바 색상 로직 (최종 확정)

**각 Skirt는 4가지 색상 중 하나:**

| 색상 | 코드 | 조건 | 의미 |
|------|------|------|------|
| ⬜ **회색** | `#5a6b78` | Skirt Status가 `SN|*` | Growing Line 준비 안 됨 (Lseam 전) |
| 🟨 **노란색** | `#F7CD42` | Skirt Status가 `SD`, `SW|*` + Joint 없음/B/NaN | Lseam 완료, Fit-up 대기 |
| ⬜ **흰색** | `#FFFFFF` | Skirt Status가 `SD`, `SW|*` + Joint = `FD*` (FDIDOD 제외) | Fit-up 완료 ~ Cir-seam 완료 전 |
| 🟩 **녹색** | `#4CAF50` | Skirt Status가 `SD`, `SW|*` + Joint = `FDIDOD` | Cir-seam (Inside & Outside) 모두 완료 ✅ |

### JavaScript 색상 결정 로직

```javascript
function getSkirtColor(skirtIndex, skirtStatus, jointStatuses) {
  // 1. Skirt Status가 SN|*이면 회색 (Growing Line 준비 안 됨)
  if (skirtStatus.startsWith('SN|')) {
    return '#5a6b78'; // ⬜ 회색
  }
  
  // 2. Skirt Status가 SD, SW|*이면 Joint 확인
  if (skirtStatus === 'SD' || skirtStatus.startsWith('SW|')) {
    // 양쪽 Joint 확인
    const leftJoint = jointStatuses[skirtIndex];      // 왼쪽 Joint
    const rightJoint = jointStatuses[skirtIndex + 1]; // 오른쪽 Joint
    
    // 양쪽 중 하나라도 FDIDOD이면 녹색 (Cir-seam 완료)
    if (leftJoint === 'FDIDOD' || rightJoint === 'FDIDOD') {
      return '#4CAF50'; // 🟩 녹색
    }
    
    // 양쪽 중 하나라도 FD*이면 흰색 (Fit-up 완료 ~ Cir-seam 완료 전)
    if ((leftJoint && leftJoint.startsWith('FD')) || 
        (rightJoint && rightJoint.startsWith('FD'))) {
      return '#FFFFFF'; // ⬜ 흰색
    }
    
    // Joint가 없거나 B(Blank)이면 노란색 (FU 전)
    return '#F7CD42'; // 🟨 노란색
  }
  
  // 기본값 (예외)
  return '#5a6b78'; // ⬜ 회색
}
```

### 색상 변화 시나리오 (3 Skirts, 4 Joints 예시)

```
[Flange] ─ [Joint 1] ─ [Skirt 1] ─ [Joint 2] ─ [Skirt 2] ─ [Joint 3] ─ [Skirt 3] ─ [Joint 4] ─ [Flange]
```

#### Step 1: Skirt 1만 라인 투입 (Lseam 완료)
```
[Flange] ─ [J1: 없음] ─ 🟨 S1 (SD, J1 없음, J2 없음) ─ ... ─ ... ─ ... ─ ...
```
- Skirt 1: 양쪽 Joint 없음 → 노란색 ✅

#### Step 2: Skirt 2 투입 후 Joint 1 Fit-up (J1 = FD)
```
[Flange] ─ [J1: FD] ─ ⬜ S1 (J1=FD, J2 없음) ─ [J2: 없음] ─ 🟨 S2 (J2 없음, J3 없음) ─ ... ─ ...
```
- Skirt 1: 왼쪽 J1=FD, 오른쪽 J2 없음 → 흰색 ✅
- Skirt 2: 양쪽 Joint 없음 → 노란색 ✅

#### Step 3: Joint 2 Fit-up (J2 = FD)
```
[Flange] ─ [J1: FD] ─ ⬜ S1 (J1=FD, J2=FD) ─ [J2: FD] ─ ⬜ S2 (J2=FD, J3 없음) ─ ... ─ ...
```
- Skirt 1: 양쪽 모두 FD → 흰색 ✅
- Skirt 2: 왼쪽 J2=FD, 오른쪽 J3 없음 → 흰색 ✅

#### Step 4: Skirt 3 투입 후 Joint 3 Fit-up (J3 = FD)
```
[Flange] ─ [J1: FD] ─ ⬜ S1 ─ [J2: FD] ─ ⬜ S2 (J2=FD, J3=FD) ─ [J3: FD] ─ ⬜ S3 (J3=FD, J4 없음) ─ ...
```
- Skirt 1: 양쪽 FD → 흰색 ✅
- Skirt 2: 양쪽 FD → 흰색 ✅
- Skirt 3: 왼쪽 J3=FD, 오른쪽 J4 없음 → 흰색 ✅

#### Step 5: Joint 2 Cir-seam 완료 (J2 = FDIDOD)
```
[Flange] ─ [J1: FD] ─ ⬜ S1 (J1=FD, J2=FDIDOD) ─ [J2: FDIDOD] ─ 🟩 S2 (J2=FDIDOD, J3=FD) ─ [J3: FD] ─ ⬜ S3 ─ ...
```
- Skirt 1: 왼쪽 J1=FD, 오른쪽 J2=FDIDOD → **녹색** ✅ (한쪽이라도 FDIDOD)
- Skirt 2: 왼쪽 J2=FDIDOD → **녹색** ✅
- Skirt 3: 양쪽 FD → 흰색 ✅

#### Step 6: Joint 1 Cir-seam 완료 (J1 = FDIDOD)
```
[Flange] ─ [J1: FDIDOD] ─ 🟩 S1 (J1=FDIDOD, J2=FDIDOD) ─ [J2: FDIDOD] ─ 🟩 S2 ─ [J3: FD] ─ ⬜ S3 ─ ...
```
- Skirt 1: 양쪽 모두 FDIDOD → 녹색 ✅
- Skirt 2: 왼쪽 FDIDOD → 녹색 ✅
- Skirt 3: 양쪽 FD → 흰색 ✅

---

## Gate Card 시간 범위 계산 로직

### 표시 형식
`06:00 AM - 07:24 AM`

### 계산 로직

```javascript
if (sts === 'R') {
  // 진행 전 (Ready)
  startTime = plan_start_dt;  // J열
  endTime = plan_start_dt + work_st;  // J열 + L열 (Standard Time)
} else if (sts === 'S' || sts === 'H') {
  // 진행 중 (Started) 또는 보류 (Hold)
  startTime = start_dt;  // H열
  endTime = start_dt + work_st;  // H열 + L열 (Standard Time)
}
```

---

## 대기 시간 계산 로직

### 표시 형식
`52 MINS WAITING`

### 계산 로직

```javascript
const waitingMinutes = Math.floor((cur_time - start_dt) / 60000);
// cur_time(Q열) - start_dt(H열) = 대기 시간 (밀리초)
```

---

## 화살표 표시 (Working Position)

**Joint Status 배열**을 검사하여 현재 작업 위치를 표시합니다:

```javascript
// 로직:
const i = jointStatuses.findIndex(s => s && s.includes('W'));

if (i !== -1) {
  // Skirt i와 Skirt i+1 사이에 화살표 표시 (해당 Joint 위치)
  // 화살표는 Skirt 막대 사이를 정확하게 가리킴
} else {
  // 첫 번째 Skirt에 화살표 표시 (대기 상태)
}
```

**검증 예시 (Gate 2)**:
- `joint_status1` = "FDOW" (**W 포함!**)
- → 화살표가 가장 왼쪽(첫 번째 Joint)에 표시됨 ✅

---

## Mod 계산 규칙

Gate 번호로 Mod를 자동 계산합니다:

```javascript
const gateNumber = parseInt(mcn_no.replace('G', '')); // "G01" → 1
const mod = Math.ceil(gateNumber / 20);
```

| Gate 범위 | Mod |
|-----------|-----|
| G01 ~ G20 | Mod 1 |
| G21 ~ G40 | Mod 2 |
| G41 ~ G60 | Mod 3 |

---

## UI 구성 요소

### Gate Card 구조 (Gate 52 예시)

```
┌─────────────────────────────────────┐
│ 2026/02/20 08:19AM          G52  →  │ (헤더: 시간, Gate 번호, 화살표)
├─────────────────────────────────────┤
│        SECTION ID                   │
│        VB100-T                      │ (Section ID, 큰 글씨)
├─────────────────────────────────────┤
│  CURRENT PROCESS                    │
│  FU-C11          WAITING            │ (현재 공정, 상태)
│                  MINS WAITING       │ (대기 시간)
│  10:21 AM        ════════  11:45 AM │ (시간 범위, 진행률 바)
│                  ════════            │ (진행률 바 2줄)
├─────────────────────────────────────┤
│  PROGRESS & WIP READINESS           │
│        ↓                            │ (화살표 - 작업 위치)
│  [🟨][⬜][🟨][🟨][⬜][⬜][⬜][⬜][⬜][⬜][⬜] │ (Skirt별 진행률 바)
└─────────────────────────────────────┘
```

### 진행률 바 (Progress Bar) 규칙

**중요**: Skirt Status와 Joint Status를 **함께** 확인하여 색상 결정!

```javascript
// 각 Skirt마다 고정 너비 박스 (30px)
Skirt별 색상:
- SN|* → ⬜ 회색 (#5a6b78) - Growing Line 준비 안 됨
- SD, SW|* + Joint 없음 → 🟨 노란색 (#F7CD42) - Lseam 완료, Fit-up 대기
- SD, SW|* + Joint FD* → ⬜ 흰색 (#FFFFFF) - Fit-up 완료 ~ Cir-seam 완료 전
- SD, SW|* + Joint FDIDOD → 🟩 녹색 (#4CAF50) - Cir-seam 완료
- "B" (Blank) → 표시 안 함
```

---

## Mod 필터링

사용자가 특정 Mod만 선택하여 볼 수 있습니다:

```javascript
필터 옵션:
- All: 전체 60개 게이트 표시
- Mod 1: Gate 1~20만 표시
- Mod 2: Gate 21~40만 표시
- Mod 3: Gate 41~60만 표시
```

---

## 데이터 예시

### Gate 1 (G01)
```javascript
{
  mcn_no: "G01",
  serial_no2: "VB087-U",
  rev_flag: 0,
  fo_desc: "VT/MT Repair",
  sts: "S",
  skirt_qty: 11,
  mod: 1,
  joint_status: ["FDIDOD", "FDIDOD", "FDIDOD", "FDIDOD", "FDIDOD", "FDIDOD", "FDIDOD", "FDIDOD", "FDIDOD", "FDIDOD", "FDIDOD", "FDIDOD"],
  skirt_status: ["SD", "SD", "SD", "SD", "SD", "SD", "SD", "SD", "SD", "SD", "SD"]
}
```
- Skirt 11개 → Joint 12개 ✅
- Joint 1~12: 모두 완료 (FDIDOD)
- 모든 Skirt: 완료 (SD)
- 화살표: 'W' 없으므로 첫 번째 위치
- 진행률 바: 모든 Skirt가 양쪽 Joint FDIDOD → 전체 녹색

### Gate 2 (G02)
```javascript
{
  mcn_no: "G02",
  serial_no2: "VB094-U",
  rev_flag: 0,
  fo_desc: "CSO-C02",
  sts: "R",
  skirt_qty: 11,
  mod: 1,
  joint_status: ["FDOW", "FD", "FD", null, null, null, null, null, null, null, null, null],
  skirt_status: ["SD", "SD", "SD", "SW|FU", "SN|LS", "SN|LS", "SN|LS", "SN|BN", "SN|BN", "SN|BN", "SN|BN"]
}
```
- Skirt 11개 → Joint 12개 (Joint 1~3만 데이터 있음)
- Joint 1: FDOW (**W 포함!** → 화살표 여기 표시)
- Joint 2~3: FD (Fit-up 완료)
- Skirt 1: J1=FDOW, J2=FD → 흰색 (FD*)
- Skirt 2: J2=FD, J3=FD → 흰색
- Skirt 3: J3=FD, J4=null → 흰색 (한쪽이라도 FD)
- Skirt 4: J4=null, J5=null, SW|FU → 노란색
- Skirt 5~11: SN|* → 회색

### Gate 52 (G52) - Reverse 예시
```javascript
{
  mcn_no: "G52",
  serial_no2: "VB100-T",
  rev_flag: 1,  // Reverse!
  fo_desc: "FU-C11",
  sts: "S",
  skirt_qty: 11,
  mod: 3,
  joint_status: [null, "B", "B", "B", "B", "B", "B", "B", "B", "B", "FD"],  // Excel에서 읽은 원본
  // Rev_flag = 1이므로 reverse 필요:
  // → [null, "FD", "B", "B", "B", "B", "B", "B", "B", "B", "B"]
  skirt_status: ["SD", "SD", "SD", "SD", "B", "B", "B", "B", "B", "B", "B"]
}
```
- **FU-C11** = 실제로는 맨 왼쪽 Joint (Skirt 1과 Skirt 2 사이)
- Rev_flag = 1이므로 Joint 배열 reverse 필요
- Skirt 1: Joint 2 = FD → 흰색
- Skirt 2~4: Joint 3~5 = B → 노란색
- Skirt 5~11: B → 표시 안 함

---

## 파일 구조

```
gd-agent/
├── README.md                           # 이 파일 (📝 핵심 문서)
├── CSWIND_CONTEXT.md                   # CS Wind AM 공장 컨텍스트
├── index.html                          # 메인 HTML (개발 버전)
├── gate-card-sample.html               # Gate Card 샘플 (디자인 참고)
├── GD Agent AI Mockup_Rev23.html       # UI 목업 (참고용)
├── assets/
│   ├── gate_card_sample_1.png          # UI 샘플 이미지 1
│   ├── gate_card_sample_2.png          # UI 샘플 이미지 2
│   ├── gate_card_sample_3.png          # UI 샘플 이미지 3
│   ├── gate_card_status.png            # 상태 정의 이미지
│   ├── wip_readiness_table.png         # WIP Readiness 테이블
│   └── waiv-logo.png                   # WAIV 로고
├── css/
│   └── (스타일시트)
├── js/
│   └── (JavaScript 파일)
└── data/
    └── GD AI Agent sample 1.xlsx       # 샘플 Excel 파일
```

---

## 개발 이력

### 완료된 작업
- ✅ Excel 파일 구조 완벽 분석 (58개 컬럼)
- ✅ 상태 코드 정의 (Joint Status: F/D/W/I/O, Skirt Status: SD/SW/SN)
- ✅ Mod 계산 규칙 정의 (Gate 번호 기반)
- ✅ **진행률 바 4색 로직 정의 (회색/노란색/흰색/녹색)** ✨
- ✅ **Skirt-Joint 매핑 규칙 확정 (Joint 1 없음, Joint = Skirt + 1)** ✨
- ✅ 화살표 표시 로직 정의 (Joint Status 'W' 검사)
- ✅ Rev_flag 처리 로직 정의 (Joint 배열 reverse)
- ✅ 시간 범위 계산 로직 정의 (F열 sts 기반 분기)
- ✅ 대기 시간 계산 로직 정의 (cur_time - start_dt)
- ✅ Gate Card 디자인 샘플 작성 (gate-card-sample.html)

### 진행 중
- 🔄 Excel 파싱 JavaScript 구현 (다음 단계)
- 🔄 60개 Gate Card 동적 렌더링
- 🔄 Mod 필터링 UI 구현
- 🔄 상태별 색상 코딩 구현

### 대기 중
- ⏳ 상태(Normal/Waiting/Delayed/QC-Delayed) 계산 로직 정의 필요
- ⏳ WIP Readiness 데이터 구조 및 표시 로직
- ⏳ 실시간 데이터 업데이트 (향후)
- ⏳ AI 분석 기능 추가 (향후)

---

## 배포 정보

- **Main URL**: https://mes-r018-analysis.pages.dev/gd-agent/
- **Git Repository**: https://github.com/twokomi/MES_R018_Analysis
- **Project Path**: `/home/user/webapp/gd-agent/`

---

## 참고 사항

### Performance Report와의 관계
- GD Agent Dashboard는 MES Performance Report와 동일한 프로젝트 내에 존재
- Performance Report: `/` (루트)
- GD Agent Dashboard: `/gd-agent/`
- 독립적인 기능이지만 동일 Git 저장소 관리

### 대화 컨텍스트 보존
이 README는 대화 압축으로 인한 정보 손실을 방지하기 위해 작성되었습니다.
모든 중요한 정보는 이 파일에 기록하여 언제든지 참조할 수 있도록 합니다.

---

**Last Updated**: 2026-02-20
**Version**: 0.2.0 (Excel 구조 완벽 분석 완료, 색상 로직 확정)
