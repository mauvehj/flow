/**
 * Sample Data Generator
 * Generates in-memory mock Excel files with distinct schemas to showcase the auto-grouping capabilities.
 */

const SampleDataGenerator = (() => {
  /**
   * Helper to create a File object from an XLSX workbook
   * @param {Object} wb
   * @param {string} filename
   * @returns {File}
   */
  const workbookToFile = (wb, filename) => {
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    return new File([blob], filename, {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      lastModified: Date.now()
    });
  };

  /**
   * Generates 5 sample Excel files across 3 distinct schemas:
   * - 2 Sales Reports (Schema 1)
   * - 2 Employee Records (Schema 2)
   * - 1 Inventory Log (Schema 3)
   * @returns {File[]}
   */
  const generateSampleFiles = () => {
    const files = [];

    // --- Schema 1: Sales Reports (매출 내역) ---
    // File 1: 1월 매출내역.xlsx
    const wbSales1 = XLSX.utils.book_new();
    const salesData1 = [
      { '일자': '2026-01-05', '지점명': '강남점', '상품명': '프리미엄 원두 (1kg)', '수량': 20, '단가': 35000, '총매출액': 700000 },
      { '일자': '2026-01-12', '지점명': '강남점', '상품명': '에스프레소 머신', '수량': 2, '단가': 1200000, '총매출액': 2400000 },
      { '일자': '2026-01-20', '지점명': '서초점', '상품명': '텀블러 블랙 (500ml)', '수량': 45, '단가': 22000, '총매출액': 990000 },
      { '일자': '2026-01-28', '지점명': '서초점', '상품명': '핸드드립 세트', '수량': 15, '단가': 48000, '총매출액': 720000 }
    ];
    XLSX.utils.book_append_sheet(wbSales1, XLSX.utils.json_to_sheet(salesData1), '1월매출');
    files.push(workbookToFile(wbSales1, '1월_매출내역_강남_서초.xlsx'));

    // File 2: 2월 매출내역.xlsx (Note: same columns, test grouping!)
    const wbSales2 = XLSX.utils.book_new();
    const salesData2 = [
      { '일자': '2026-02-03', '지점명': '판교점', '상품명': '프리미엄 원두 (1kg)', '수량': 35, '단가': 35000, '총매출액': 1225000 },
      { '일자': '2026-02-14', '지점명': '판교점', '상품명': '바닐라 시럽 (1L)', '수량': 50, '단가': 14000, '총매출액': 700000 },
      { '일자': '2026-02-22', '지점명': '송도점', '상품명': '에스프레소 머신', '수량': 3, '단가': 1200000, '총매출액': 3600000 },
      { '일자': '2026-02-27', '지점명': '송도점', '상품명': '텀블러 화이트 (500ml)', '수량': 30, '단가': 22000, '총매출액': 660000 }
    ];
    XLSX.utils.book_append_sheet(wbSales2, XLSX.utils.json_to_sheet(salesData2), '2월매출');
    files.push(workbookToFile(wbSales2, '2월_매출내역_판교_송도.xlsx'));

    // --- Schema 2: Employee Records (임직원 명부) ---
    // File 3: 임직원명부_본사.xlsx
    const wbEmp1 = XLSX.utils.book_new();
    const empData1 = [
      { '사번': 'HQ-1001', '성명': '김도현', '소속부서': '경영전략팀', '직급': '팀장', '연락처': '010-1234-5678', '입사일': '2021-03-01' },
      { '사번': 'HQ-1002', '성명': '이수아', '소속부서': '인사총무팀', '직급': '선임', '연락처': '010-2345-6789', '입사일': '2022-07-15' },
      { '사번': 'HQ-1003', '성명': '박준영', '소속부서': '재무회계팀', '직급': '책임', '연락처': '010-3456-7890', '입사일': '2020-01-10' }
    ];
    XLSX.utils.book_append_sheet(wbEmp1, XLSX.utils.json_to_sheet(empData1), '본사명단');
    files.push(workbookToFile(wbEmp1, '임직원명부_본사_경영부서.xlsx'));

    // File 4: 임직원명부_연구소.xlsx (Same columns as File 3)
    const wbEmp2 = XLSX.utils.book_new();
    const empData2 = [
      { '사번': 'RD-2001', '성명': '최민호', '소속부서': 'AI연구실', '직급': '수석연구원', '연락처': '010-4567-8901', '입사일': '2019-11-01' },
      { '사번': 'RD-2002', '성명': '정유진', '소속부서': '소프트웨어팀', '직급': '선임연구원', '연락처': '010-5678-9012', '입사일': '2023-02-01' },
      { '사번': 'RD-2003', '성명': '한승우', '소속부서': '데이터엔지니어링', '직급': '연구원', '연락처': '010-6789-0123', '입사일': '2024-05-15' }
    ];
    XLSX.utils.book_append_sheet(wbEmp2, XLSX.utils.json_to_sheet(empData2), '연구소명단');
    files.push(workbookToFile(wbEmp2, '임직원명부_R&D센터.xlsx'));

    // --- Schema 3: Inventory Log (물류 재고) ---
    // File 5: 물류창고_실사재고.xlsx (Distinct schema)
    const wbInv = XLSX.utils.book_new();
    const invData = [
      { '자재코드': 'MAT-A01', '품목명': '고급 크라프트 박스', '보관위치': 'A구역-101호', '현재고': 1500, '적정재고': 1000 },
      { '자재코드': 'MAT-A02', '품목명': '친환경 완충재 (롤)', '보관위치': 'A구역-104호', '현재고': 80, '적정재고': 100 },
      { '자재코드': 'MAT-B05', '품목명': '로고 각인 컵 홀더', '보관위치': 'B구역-202호', '현재고': 4200, '적정재고': 3000 },
      { '자재코드': 'MAT-C12', '품목명': '방수 패키지 테이프', '보관위치': 'C구역-301호', '현재고': 230, '적정재고': 200 }
    ];
    XLSX.utils.book_append_sheet(wbInv, XLSX.utils.json_to_sheet(invData), '재고실사');
    files.push(workbookToFile(wbInv, '물류창고_포장자재_실사재고.xlsx'));

    return files;
  };

  return {
    generateSampleFiles
  };
})();

// Export to global scope
window.SampleDataGenerator = SampleDataGenerator;
