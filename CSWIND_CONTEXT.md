# CS Wind AM 공장 - 풍력타워 제조 공정 컨텍스트

## 공정 개요
- **회사**: CS Wind AM (미국 공장)
- **제품**: 풍력타워 (Wind Tower)
- **주요 KPI**: BT Green Tag / Release 된 BT Section 수량

## 타워 구조

### 계층 구조
```
Tower (타워)
  └─ Section (섹션: Top/Mid/Bottom)
      └─ Skirt(Can) (스커트: 5~11개)
          └─ Joint (조인트: Skirt 사이 용접부, Skirt보다 1개 적음)
```

### Section 특성
- **Top**: 직경 작음 / 철판 얇음 → 용접량 적음 (일반적으로 11개 Skirt)
- **Mid**: Mid1, Mid2, Mid3… 세분화 가능
- **Bottom**: 직경 큼 / 철판 두꺼움 → 용접량 많음

### ID 명명 규칙
- Section ID: `VB001-T` (VB=프로젝트명, 001=타워번호, T/M/B=섹션구분)
- Skirt ID: `VB001-T1` (섹션 내 Skirt 순번)

## AM 공장 레이아웃

### Gate / Growing Line
- **60개 Gate = 60개 Growing Line (1:1 대응)**
- Section 완성 시 Gate가 열리고 Release (BT 완료품 출하)

### Mod 병렬 구조
- **Mod 1**: Gate 1~20 (Growing Line 20개)
- **Mod 2**: Gate 21~40 (Growing Line 20개)
- **Mod 3**: Gate 41~60 (Growing Line 20개)

### 주요 설비
- **SAW Machine**: Mod당 3대 (총 9대)
- **LW Machine (L)**: L1~L11 (Long seam 용접)
- **Bending Machine (B)**: B1~B6
- **FFU / FC(FTC)**: Flange 공정 구역
  - FC는 Mod별 1대 (FC1=Mod1, FC2=Mod2, FC3=Mod3)

## BT (Black Tower) 생산 흐름

### 전체 공정 순서
```
Steel Plate 입고
  → Pre Blasting
  → Cutting (Oxy Cut / Plasma Cut)
  → Beveling → Supermarket (버퍼)
  → Bending (B) → Can 제작
  → Tack + LW (L) → Can 완성 (Long Seam 용접)
  → [선택] FFU → FC(FTC) (끝단 Skirt Flange 부착)
  → Growing Line 투입
  → Fit-up + CW (원주 용접, 3-Can 배치 방식)
  → NDT (in-line)
  → [필요시] Repair (in-line)
  → Green Tag / Release
```

### Growing Line 조립 방식 (핵심)

**3-Can 배치 규칙:**
1. 3개 Can을 배치하고 해당 배치의 Joint들에 대해 Fit-up + CW 완료
2. **SAW 용접 순서** (효율성):
   - Outside welding을 배치의 Joint들에 대해 먼저 몰아서 수행
   - Inside welding을 배치의 Joint들에 대해 이후에 몰아서 수행
   - Outside/Inside를 번갈아 수행하지 않음 (Head 출입 반복 최소화)
3. 배치 CW 완료 후, 조립된 구간을 Gate 방향으로 이동해 작업공간 확보
4. SAW 작업 위치와 조립 구간 사이 공간에 다음 Can들을 투입 (끼워 넣기)
5. 위 과정을 반복해 최종 Section 길이 완성

### Joint 번호 규칙
- Joint No. = Section 기준 연속 번호 (01~N)
- 예: 10개 Can Section → Joint 9개 → 01~09
- **FU-03** = 해당 Section의 3번째 Joint (Fit-up)

### 용어 정리
- **LW**: Long seam (세로 용접)
- **CW**: Circumferential seam (원주 용접, Cir-seam)
- **FU**: Fit-up (맞대기)
- **ID**: Inside Done (내부 용접 완료)
- **OD**: Outside Done (외부 용접 완료)
- **FFU**: Flange Fit-up
- **FC(FTC)**: Flange Circumferential (Flange 원주 용접)
- **Green Tag**: NDT 합격 후 부착되는 완료 태그
- **Release**: BT 완료품 출하

## MES 트래킹

### 트래킹 단위 전환
- **Pre-Growing (LSeam까지)**: Can(Skirt) 단위 트래킹
- **Growing Line (Fit-up 이후)**: Section/Joint 단위 트래킹

### Joint 입력 시 Skirt ID 매핑 규칙
- Joint 진행 입력 시: **FU-xx ↔ Txx** (동일 번호 매핑)
- 예: FU-03 진행 입력 시 → VB000-T3 (Skirt #3) 입력

### Rev_flag (투입 방향)
- **Rev_flag = 0 (Normal)**: 두꺼운 쪽부터 투입 (정방향)
- **Rev_flag = 1 (Reverse)**: 얇은 쪽부터 투입 (역방향)
  - Joint No와 Skirt No 방향이 반대

## 3대 공정 분류
- **BT (Black Tower)**: Steel Plate 입고 → 용접 완료 Section (도장 전)
- **WT (White Tower)**: Washing → Blasting → Metalizing → Painting
- **IM (Internal Mounting)**: 내부 전기/기계 설비 설치

**참고**: GD Agent Dashboard는 BT 공정 중심으로 설계됨
