/**
 * GD Agent AI - Excel Parser
 * 
 * 목적: GD AI Agent sample 1.xlsx 파일 파싱
 * 컬럼: 58개 (A~BA)
 * 행: 60개 (Gate G01~G60)
 */

class ExcelParser {
  constructor() {
    this.rawData = null;      // Excel에서 읽은 원본 데이터
    this.gates = [];          // 파싱된 60개 Gate 데이터
    this.isLoaded = false;    // 데이터 로드 완료 여부
  }

  /**
   * Excel 파일 로드 및 파싱
   * @param {File} file - Excel 파일 객체
   * @returns {Promise<Array>} - 파싱된 Gate 데이터 배열
   */
  async loadExcel(file) {
    console.log('📂 Excel 파일 로드 시작:', file.name);

    try {
      // 1. 파일을 ArrayBuffer로 읽기
      const arrayBuffer = await this.readFileAsArrayBuffer(file);
      console.log('✅ 파일 읽기 완료, 크기:', arrayBuffer.byteLength, 'bytes');

      // 2. SheetJS로 Excel 파싱
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      console.log('✅ Workbook 파싱 완료, 시트 수:', workbook.SheetNames.length);
      console.log('📋 시트 이름:', workbook.SheetNames);

      // 3. 첫 번째 시트 선택
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      console.log('📄 시트 선택:', sheetName);

      // 4. JSON 형식으로 변환 (header: 1 = 배열 형태)
      this.rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
      console.log('✅ 데이터 변환 완료, 총 행 수:', this.rawData.length);

      // 5. 데이터 구조 확인
      this.validateData();

      // 6. 60개 Gate 데이터 파싱
      this.parseGates();

      this.isLoaded = true;
      console.log('🎉 Excel 파싱 완료! 총 Gate 수:', this.gates.length);

      return this.gates;

    } catch (error) {
      console.error('❌ Excel 파싱 오류:', error);
      throw error;
    }
  }

  /**
   * 파일을 ArrayBuffer로 읽기
   * @param {File} file
   * @returns {Promise<ArrayBuffer>}
   */
  readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = (e) => reject(new Error('파일 읽기 실패'));
      reader.readAsArrayBuffer(file);
    });
  }

  /**
   * 데이터 유효성 검증
   */
  validateData() {
    console.log('🔍 데이터 유효성 검증 시작...');

    if (!this.rawData || this.rawData.length === 0) {
      throw new Error('데이터가 비어있습니다.');
    }

    // 헤더 행 확인 (첫 번째 행)
    const header = this.rawData[0];
    console.log('📋 헤더 (첫 10개):', header.slice(0, 10));
    console.log('📊 총 컬럼 수:', header.length);

    // 데이터 행 수 확인 (헤더 제외)
    const dataRows = this.rawData.length - 1;
    console.log('📊 데이터 행 수 (헤더 제외):', dataRows);

    if (dataRows !== 60) {
      console.warn(`⚠️ 경고: 60개 Gate가 예상되지만 ${dataRows}개 행이 있습니다.`);
    }

    // 첫 번째 데이터 행 샘플 (디버깅용)
    console.log('🔬 첫 번째 데이터 행 샘플 (첫 10개 컬럼):', this.rawData[1]?.slice(0, 10));
  }

  /**
   * 60개 Gate 데이터 파싱
   */
  parseGates() {
    console.log('⚙️ Gate 데이터 파싱 시작...');

    // 헤더 제외하고 데이터 행만 처리
    const dataRows = this.rawData.slice(1);

    this.gates = dataRows.map((row, index) => {
      try {
        return this.parseGateRow(row, index);
      } catch (error) {
        console.error(`❌ Gate ${index + 1} 파싱 오류:`, error);
        return null;
      }
    }).filter(gate => gate !== null); // null 제거

    console.log('✅ Gate 파싱 완료:', this.gates.length, '개');
  }

  /**
   * 단일 Gate 행 파싱
   * @param {Array} row - Excel 행 데이터 (배열)
   * @param {Number} index - 행 인덱스 (0부터 시작)
   * @returns {Object} - 파싱된 Gate 객체
   */
  parseGateRow(row, index) {
    // A~Q: 기본 정보 (17개)
    const mcn_no = this.getCellValue(row, 0);           // A: Gate 번호
    const serial_no2 = this.getCellValue(row, 1);       // B: Section ID
    const rev_flag = this.getCellValue(row, 2) || 0;    // C: Rev flag (0=Normal, 1=Reverse)
    const wo_dtl_id = this.getCellValue(row, 3);        // D: Work Order ID
    const fo_desc = this.getCellValue(row, 4);          // E: 현재 공정
    const sts = this.getCellValue(row, 5);              // F: Status (S/R/H)
    const working_rate = this.getCellValue(row, 6);     // G: Working Rate (%)
    const start_dt = this.getCellValue(row, 7);         // H: 시작 시간
    const end_dt = this.getCellValue(row, 8);           // I: 종료 시간
    const plan_start_dt = this.getCellValue(row, 9);    // J: 계획 시작
    const plan_end_dt = this.getCellValue(row, 10);     // K: 계획 종료
    const work_st = this.getCellValue(row, 11);         // L: Standard Time
    const worker_id = this.getCellValue(row, 12);       // M: 작업자 ID
    const worker_nm = this.getCellValue(row, 13);       // N: 작업자 이름
    const skirt_qty = this.getCellValue(row, 14) || 0;  // O: Skirt 개수
    const proj_color = this.getCellValue(row, 15);      // P: 프로젝트 색상
    const cur_time = this.getCellValue(row, 16);        // Q: 현재 시간

    // R~AK: Joint Status (20개, index 17~36)
    let jointStatuses = [];
    for (let i = 17; i < 37; i++) {
      jointStatuses.push(this.getCellValue(row, i) || 'B'); // B = Blank
    }

    // AL: Plant (index 37)
    const plant = this.getCellValue(row, 37);

    // AM~A`: Skirt Status (20개, index 38~57)
    let skirtStatuses = [];
    for (let i = 38; i < 58; i++) {
      skirtStatuses.push(this.getCellValue(row, i) || 'B'); // B = Blank
    }

    // Mod 계산 (Gate 번호 기반)
    const gateNumber = parseInt(mcn_no?.replace('G', '') || '0');
    const mod = Math.ceil(gateNumber / 20);

    // Rev_flag 처리 (Reverse일 때 Joint 배열 뒤집기)
    if (rev_flag === 1) {
      // Joint 1은 없으므로 index 0 제외하고 reverse
      const joints = jointStatuses.slice(1);
      joints.reverse();
      jointStatuses = [null, ...joints]; // index 0에 null 추가
    }

    // Gate 객체 생성
    const gate = {
      // 기본 정보
      mcn_no,
      serial_no2,
      rev_flag,
      wo_dtl_id,
      fo_desc,
      sts,
      working_rate,
      start_dt,
      end_dt,
      plan_start_dt,
      plan_end_dt,
      work_st,
      worker_id,
      worker_nm,
      skirt_qty,
      proj_color,
      cur_time,
      plant,
      
      // 계산된 값
      mod,
      gateNumber,
      
      // 배열 데이터
      jointStatuses,
      skirtStatuses,
      
      // 메타 정보
      rowIndex: index,
      isReverse: rev_flag === 1
    };

    // 디버깅: 처음 3개 Gate만 로그
    if (index < 3) {
      console.log(`🔍 Gate ${gateNumber} (${mcn_no}) 파싱 완료:`, {
        section: serial_no2,
        process: fo_desc,
        status: sts,
        mod,
        skirt_qty,
        rev_flag: rev_flag === 1 ? 'Reverse' : 'Normal',
        jointCount: jointStatuses.filter(j => j && j !== 'B').length,
        skirtCount: skirtStatuses.filter(s => s && s !== 'B').length
      });
    }

    return gate;
  }

  /**
   * 셀 값 가져오기 (null/undefined 처리)
   * @param {Array} row - 행 배열
   * @param {Number} colIndex - 컬럼 인덱스
   * @returns {*} - 셀 값
   */
  getCellValue(row, colIndex) {
    const value = row[colIndex];
    
    // null, undefined, 빈 문자열 처리
    if (value === null || value === undefined || value === '') {
      return null;
    }
    
    // 'B' (Blank) 처리
    if (value === 'B') {
      return 'B';
    }
    
    return value;
  }

  /**
   * Gate 번호로 Gate 찾기
   * @param {String} mcn_no - Gate 번호 (예: "G01")
   * @returns {Object|null} - Gate 객체
   */
  getGateByNumber(mcn_no) {
    return this.gates.find(gate => gate.mcn_no === mcn_no) || null;
  }

  /**
   * Mod로 필터링
   * @param {Number} mod - Mod 번호 (1, 2, 3)
   * @returns {Array} - 필터링된 Gate 배열
   */
  filterByMod(mod) {
    if (!mod) return this.gates; // mod가 없으면 전체 반환
    return this.gates.filter(gate => gate.mod === mod);
  }

  /**
   * 파싱된 데이터 요약
   * @returns {Object} - 요약 정보
   */
  getSummary() {
    if (!this.isLoaded) {
      return { error: '데이터가 로드되지 않았습니다.' };
    }

    const summary = {
      totalGates: this.gates.length,
      mod1: this.filterByMod(1).length,
      mod2: this.filterByMod(2).length,
      mod3: this.filterByMod(3).length,
      statusCount: {},
      reverseCount: this.gates.filter(g => g.isReverse).length
    };

    // Status 별 카운트
    this.gates.forEach(gate => {
      const status = gate.sts || 'Unknown';
      summary.statusCount[status] = (summary.statusCount[status] || 0) + 1;
    });

    return summary;
  }
}

// 전역 인스턴스 생성
const excelParser = new ExcelParser();

console.log('✅ ExcelParser 클래스 로드 완료');
