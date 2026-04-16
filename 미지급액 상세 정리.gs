/**
 * 미지급액 상세 정리 최종 스크립트
 * 단일 업체와 복수 업체를 구분하여 처리
 * 
 * ⚠️ 첨부파일 관리 스크립트와 함께 사용하기 위해 수정됨
 */

// 메인 시트 이름 설정
const MAIN_SHEET_NAME = '목차';

// ===== onOpen 함수 이름 변경: 첨부파일 스크립트와 충돌 방지 =====
// 이 함수는 직접 호출하지 않고, 첨부파일 스크립트의 onOpen에서 호출됩니다
function addUnpaidMenu() {
  const ui = SpreadsheetApp.getUi();
  
  // 미지급액 관리 메뉴 추가
  ui.createMenu('💰 미지급액 관리')
    .addItem('✅ 최종 보고서 생성', 'createFinalUnpaidReport')
    .addItem('📋 간단 요약 보기', 'showQuickSummary')
    .addSeparator()
    .addItem('⚙️ 메인 시트 설정', 'configureMainSheet')
    .addItem('🔍 디버그 모드 실행', 'runDebugMode')
    .addItem('❓ 도움말', 'showUnpaidHelp')
    .addToUi();
}

/**
 * 메인 함수: 미지급액 상세 보고서 생성
 */
function createFinalUnpaidReport() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 메인 시트 가져오기
  let mainSheet;
  try {
    mainSheet = spreadsheet.getSheetByName(MAIN_SHEET_NAME);
  } catch (e) {
    // 시트 이름으로 찾지 못하면 현재 활성 시트 사용
    mainSheet = spreadsheet.getActiveSheet();
    console.log(`'${MAIN_SHEET_NAME}' 시트를 찾을 수 없어 현재 활성 시트를 사용합니다.`);
  }
  
  console.log(`메인 시트: ${mainSheet.getName()}`);
  console.log('미지급액 보고서 생성 시작...');
  
  // 보고서 시트 생성 또는 초기화
  let reportSheet;
  try {
    reportSheet = spreadsheet.getSheetByName('미지급액_최종보고서');
    reportSheet.clear();
  } catch (e) {
    reportSheet = spreadsheet.insertSheet('미지급액_최종보고서');
  }
  
  // 헤더 설정
  const baseHeaders = [
    '번호',
    '공사',
    '업체',
    '상세',
    '발행(청구)금액',
    '지급액',
    '미지급액',
    '비고'
  ];

  reportSheet.getRange(1, 1, 1, baseHeaders.length).setValues([baseHeaders]);
  reportSheet.getRange(1, 1, 1, baseHeaders.length)
    .setFontWeight('bold')
    .setBackground('#4a90e2')
    .setFontColor('#ffffff')
    .setBorder(true, true, true, true, true, true);
  
  // 메인 시트 데이터 읽기 - A1:K80 범위로 지정
  const dataRange = mainSheet.getRange('A1:K90');
  const values = dataRange.getValues();
  
  const allUnpaidItems = [];
  const monthKeySet = new Set();
  let itemNumber = 1;
  let lastMainCategory = ''; // 대분류 저장
  let lastSubItem = ''; // 마지막 세부항목 저장 (병합 셀 처리용)
  let subItemStartRow = -1; // 병합 셀 시작 행 추적
  const processedSheets = new Set(); // 이미 처리한 복수 항목 시트 추적
  const skippedIndividualRows = new Set(); // 인건비 개별 행들 추적
  
  // 각 행 검사
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    
    // A열에서 대분류 번호 확인 (1), (2), (3) 등
    const numCol = String(row[0] || '').trim();
    if (numCol.match(/^\(\d+\)$/)) {
      // 대분류 발견 - B열에서 실제 카테고리명 읽기
      lastMainCategory = String(row[1] || '').trim();
      console.log(`대분류 발견: ${lastMainCategory}`);
    }
    
    // C열 값 확인 (새로운 세부항목인지 체크)
    const currentSubItem = String(row[2] || '').trim();
    const companyInD = String(row[3] || '').trim();
    
    // C열에 값이 있고 헤더가 아닌 경우 새 세부항목으로 설정
    if (currentSubItem && currentSubItem !== '' && currentSubItem !== '세부항목' && currentSubItem !== '업체') {
      // 새로운 세부항목 발견
      lastSubItem = currentSubItem;
      subItemStartRow = i;
      console.log(`새 세부항목 발견: ${lastSubItem} at row ${i+1}`);
    }
    
    // I열(미지급액) 확인 - 먼저 미지급액이 있는지 확인
    const unpaidAmountInSheet = parseFloat(row[8]) || 0; // I열 - 미지급액
    
    // 미지급액이 0보다 큰 경우만 처리
    if (unpaidAmountInSheet > 0) {
      // D열(업체명) 확인
      const company = companyInD; // D열 - 업체명
      
      // 업체명이 없거나 헤더인 경우 건너뛰기
      if (!company || company === '업체/성명' || company === '업체') {
        continue;
      }
      
      // C열이 "인건비(설비)" 같은 복수 항목인지 먼저 확인
      if (currentSubItem && currentSubItem.includes('인건비') && currentSubItem.includes('(')) {
        // 인건비(설비) 같은 복수 항목 - 개별 행은 건너뛰기
        console.log(`인건비 개별 행 건너뛰기 (행 ${i+1}): ${currentSubItem} - ${company}`);
        skippedIndividualRows.add(i); // 나중에 참조하기 위해 저장
        continue;
      }
      
      // 행 번호가 7번(index 6)이고 D열이 "인건비(설비)"인 경우 건너뛰기
      // 이 행은 나중에 설비 섹션에서 처리
      if (i === 6 && company.includes('인건비') && company.includes('(')) {
        console.log(`인건비(설비) 목차 행 건너뛰기 (행 ${i+1}) - 나중에 처리`);
        continue;
      }
      
      // 실제 처리할 항목명 결정
      let actualItemName = '';
      let constructionName = '';
      
            // D열(업체명)을 먼저 확인하여 복수 항목 패턴 체크
      if (company) {
        // "그 외" 패턴
        if (company.includes('그 외') || company.includes('그외')) {
          actualItemName = company; // 그 외(공통가설공사) 전체 사용
          constructionName = currentSubItem || lastSubItem || lastMainCategory;
          console.log(`복수 항목 발견 - 그 외 패턴 (D열): ${actualItemName}`);
        }
        // "인건비" 패턴
        else if (company.includes('인건비') && company.includes('(')) {
          // 공종 상관없이 이름에 '인건비'와 괄호가 있으면 복수 항목으로 지정
          actualItemName = company; 
          constructionName = currentSubItem || lastSubItem;
          console.log(`복수 항목 발견 - 인건비 패턴 (D열): ${actualItemName}`);
        }
        // 기타 복수 항목들 (식대, 보험, 민원 등은 D열에 직접 나타나지 않으므로 C열에서 처리)
      }
      
      // actualItemName이 아직 설정되지 않은 경우 일반 처리
      if (!actualItemName) {
        actualItemName = currentSubItem || lastSubItem;
        constructionName = actualItemName;
        
        // 빈 세부항목 처리
        if (!actualItemName || actualItemName === '') {
          // 병합 셀 범위 내에 있는지 확인
          if (i - subItemStartRow < 10) {
            actualItemName = lastSubItem;
            constructionName = actualItemName;
          } else {
            // 너무 멀리 떨어져 있으면 업체명을 항목명으로 사용
            actualItemName = company;
            constructionName = actualItemName;
          }
        }
      }
      
      // 헤더나 빈 행 건너뛰기
      if (!actualItemName || actualItemName === '세부항목') {
        continue;
      }
      
      // 구분 결정 - 대분류 사용
      const category = lastMainCategory || '기타';
      
      // J열(발행금액)과 H열(지급액) 읽기
      let invoiceAmount = parseFloat(row[9]) || 0; // J열 - 계산서 발행금액(청구금액)
      let paidAmount = parseFloat(row[7]) || 0; // H열 - 지급액
      let calculatedUnpaid = invoiceAmount - paidAmount; // 미지급액 재계산
      
      console.log(`처리중: ${category} - ${actualItemName} - ${company} (row ${i+1})`);
      console.log(`  발행: ${invoiceAmount}, 지급: ${paidAmount}, 미지급: ${calculatedUnpaid}`);
      
      // 복수 업체/인원 항목인지 확인 - actualItemName으로 판단
      if (isMultipleEntityItem(actualItemName)) {
        // 이미 처리한 복수 항목인지 확인
        if (processedSheets.has(actualItemName)) {
          console.log(`복수 항목 ${actualItemName} 이미 처리됨 - 건너뛰기`);
          continue;
        }
        
        console.log(`복수 업체 항목 발견: ${actualItemName}`);
        processedSheets.add(actualItemName); // 처리한 항목으로 표시
        
        // 상세 시트에서 개별 미지급 내역 가져오기
        const detailResult = collectEntityUnpaidData(spreadsheet, [actualItemName], constructionName);
        const detailItems = (detailResult && detailResult.detailItems) || [];

        if (detailItems.length > 0) {
          console.log(`${detailItems.length}개 세부 항목 발견`);
          // 복수 내역은 개별 항목만 추가 (합계 행 제외)
          detailItems.forEach(item => {
            const monthlyBreakdown = item.monthlyBreakdown || {};
            Object.keys(monthlyBreakdown).forEach(key => {
              if (monthlyBreakdown[key] > 0) {
                monthKeySet.add(key);
              }
            });

            allUnpaidItems.push({
              number: itemNumber++,
              construction: item.constructionName,
              company: item.subItem,
              detail: item.entity,
              invoiceAmount: item.invoiceAmount,
              paidAmount: item.paidAmount,
              unpaidAmount: item.unpaidAmount,
              note: item.note || '',
              monthly: monthlyBreakdown
            });
          });
          // 복수 내역은 목차의 합계 행을 표시하지 않음
          // continue를 사용하여 이 행 처리를 완전히 건너뜀
          continue;
        } else {
          console.log('상세 시트에서 데이터를 찾을 수 없음');
          // 상세 시트가 없는 복수 항목도 건너뜀
          continue;
        }
      } else {
        // 단일 업체 항목
        console.log(`단일 업체 항목: ${actualItemName} - ${company}`);
        
        // 단일 업체 표시를 위한 정리
        let displayConstruction = constructionName;
        let displayCompany = company;
        
        // actualItemName과 company가 같으면 constructionName을 공사명으로 사용
        if (actualItemName === company) {
          displayConstruction = lastSubItem || lastMainCategory;
        }
        
        let monthlyBreakdown = {};

        const possibleSheetNames = [
          company,
          actualItemName,
          currentSubItem,
          lastSubItem,
          constructionName,
          lastMainCategory
        ];

        const detailResult = collectEntityUnpaidData(spreadsheet, possibleSheetNames, constructionName);
        if (detailResult && detailResult.entityData) {
          const entityRecord = findEntityRecord(detailResult.entityData, company);
          if (entityRecord) {
            invoiceAmount = entityRecord.invoiceAmount || invoiceAmount;
            paidAmount = entityRecord.paidAmount || paidAmount;
            calculatedUnpaid = entityRecord.unpaidAmount || calculatedUnpaid;
            monthlyBreakdown = entityRecord.monthlyUnpaid || {};
          }
        }

        Object.keys(monthlyBreakdown).forEach(key => {
          if (monthlyBreakdown[key] > 0) {
            monthKeySet.add(key);
          }
        });

        allUnpaidItems.push({
          number: itemNumber++,
          construction: displayConstruction,
          company: displayCompany,
          detail: '',
          invoiceAmount: invoiceAmount,
          paidAmount: paidAmount,
          unpaidAmount: calculatedUnpaid,
          note: '',
          monthly: monthlyBreakdown
        });
      }
    }
  }

  // 보고서에 데이터 입력
  if (allUnpaidItems.length > 0) {
    const monthKeys = Array.from(monthKeySet).sort();
    const monthLabels = monthKeys.map(key => getMonthLabel(key));

    const headers = [
      '번호',
      '공사',
      '업체',
      '상세',
      '발행(청구)금액',
      '지급액',
      '미지급액',
      ...monthLabels,
      '비고'
    ];

    reportSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    reportSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4a90e2')
      .setFontColor('#ffffff')
      .setBorder(true, true, true, true, true, true);

    const dataRows = allUnpaidItems.map(item => {
      const monthlyValues = monthKeys.map(key => item.monthly[key] || 0);
      return [
        item.number,
        item.construction,
        item.company,
        item.detail,
        item.invoiceAmount,
        item.paidAmount,
        item.unpaidAmount,
        ...monthlyValues,
        item.note
      ];
    });

    reportSheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);

    // 합계 계산
    let totalInvoice = 0;
    let totalPaid = 0;
    let totalUnpaid = 0;
    const monthTotals = new Array(monthKeys.length).fill(0);

    allUnpaidItems.forEach((item, index) => {
      totalInvoice += item.invoiceAmount || 0;
      totalPaid += item.paidAmount || 0;
      totalUnpaid += item.unpaidAmount || 0;

      monthKeys.forEach((key, idx) => {
        const value = item.monthly[key] || 0;
        monthTotals[idx] += value;
      });
    });

    // 합계 행 추가
    const summaryRow = allUnpaidItems.length + 3;
    reportSheet.getRange(summaryRow, 1).setValue('합계');
    reportSheet.getRange(summaryRow, 5).setValue(totalInvoice);
    reportSheet.getRange(summaryRow, 6).setValue(totalPaid);
    reportSheet.getRange(summaryRow, 7).setValue(totalUnpaid);

    monthTotals.forEach((value, idx) => {
      if (value > 0) {
        reportSheet.getRange(summaryRow, 8 + idx).setValue(value);
      }
    });

    reportSheet.getRange(summaryRow, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#ffffcc')
      .setBorder(true, true, true, true, true, true);

    // 서식 적용
    formatReportSheet(reportSheet, allUnpaidItems.length, monthKeys.length);

    // 완료 메시지
    const message = `✅ 보고서 생성 완료!\n\n` +
                   `📊 총 ${allUnpaidItems.length}개 항목 정리\n` +
                   `💰 총 미지급액: ${formatNumber(Math.round(totalUnpaid))}원\n\n` +
                   `'미지급액_최종보고서' 시트를 확인하세요.`;
    
    SpreadsheetApp.getUi().alert('보고서 생성 완료', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } else {
    SpreadsheetApp.getUi().alert(
      '보고서 생성',
      '미지급액이 있는 항목이 없습니다.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 복수 업체/인원 항목인지 확인
 * 항목명을 분석하여 복수 내역이 있는지 판단
 */
function isMultipleEntityItem(subItem) {
  const item = String(subItem).trim();
  
  // 복수 내역 키워드/패턴 (우선순위 높음)
  const multiplePatterns = [
    // "그 외"가 포함된 모든 항목
    /그\s*외/i,
    
    // "인건비"가 포함된 모든 항목
    /인건비/i,
    
    // 특정 복수 내역 항목들
    /^식대$/i,           // 식대 - 여러 음식점
    /^보험$/i,           // 보험 - 각종 보험
    /^민원$/i,           // 민원 - 다양한 민원 처리
    /퇴직공제부금/i,      // 퇴직공제부금 - 여러 직원
  ];
  
  // 복수 내역인지 확인
  for (let pattern of multiplePatterns) {
    if (pattern.test(item)) {
      console.log(`복수 내역으로 판단: ${item}`);
      return true;
    }
  }
  
  // 나머지는 모두 단일 업체로 처리
  // 유진기업(주)동서울, 대진상사(설비), 대진상사(잡자재), 성광골재(주)(잡자재) 등 포함
  console.log(`단일 업체로 판단: ${item}`);
  return false;
}

/**
 * 상세 시트에서 업체별 미지급 내역 및 월별 현황을 수집
 */
function collectEntityUnpaidData(spreadsheet, possibleSheetNames, constructionName) {
  const filteredNames = (possibleSheetNames || []).filter(name => !!name);
  const entityData = {};
  const detailItems = [];

  try {
    const detailSheet = findDetailSheet(spreadsheet, filteredNames);

    if (!detailSheet) {
      console.log(`상세 시트를 찾을 수 없음: ${filteredNames.join(', ')}`);
      return { sheetName: '', entityData, detailItems };
    }

    const sheetName = detailSheet.getName();
    console.log(`상세 시트 발견: ${sheetName}`);

    const dataRange = detailSheet.getDataRange();
    const values = dataRange.getValues();

    let paymentAmountCol = -1;
    let paymentCompanyCol = -1;
    let paymentDateCol = -1;
    let invoiceAmountCol = -1;
    let invoiceCompanyCol = -1;
    let invoiceDateCol = -1;

    // 헤더 후보 스캔 (최대 15행)
    for (let i = 0; i < Math.min(15, values.length); i++) {
      for (let j = 0; j < values[i].length; j++) {
        const cell = String(values[i][j] || '').toLowerCase().replace(/\s/g, '');

        if (cell.includes('지급금액')) paymentAmountCol = j;
        if (cell.includes('지급업체') || cell === '지급처') paymentCompanyCol = j;
        if (cell.includes('지급') && (cell.includes('일') || cell.includes('날짜') || cell.includes('일자') || cell.includes('입금'))) paymentDateCol = j;

        if ((cell.includes('발행') || cell.includes('청구')) && cell.includes('금액') && !cell.includes('부가세')) invoiceAmountCol = j;
        if ((cell.includes('발행') || cell.includes('청구')) && cell.includes('업체')) invoiceCompanyCol = j;
        if ((cell.includes('발행') || cell.includes('청구')) && (cell.includes('일') || cell.includes('날짜') || cell.includes('일자'))) invoiceDateCol = j;
      }
    }

    if (invoiceAmountCol < 0) {
      invoiceAmountCol = 10; // K열(0-index 10) 추정
      console.log('발행금액 열을 기본값 K열(10)로 설정');
    }
    if (invoiceCompanyCol < 0) {
      invoiceCompanyCol = 12; // M열(0-index 12) 추정
      console.log('발행업체 열을 기본값 M열(12)로 설정');
    }
    if (invoiceDateCol < 0 && invoiceAmountCol > 0) {
      invoiceDateCol = invoiceAmountCol - 1;
      console.log(`발행일 열을 금액 열(${invoiceAmountCol}) 기준으로 추정: ${invoiceDateCol}`);
    }

    console.log(`헤더 위치 - 지급금액:${paymentAmountCol}, 지급업체:${paymentCompanyCol}, 지급일:${paymentDateCol}, 발행금액:${invoiceAmountCol}, 발행업체:${invoiceCompanyCol}, 발행일:${invoiceDateCol}`);

    // 데이터 시작 행 추정
    let dataStartRow = 1;
    for (let i = 0; i < Math.min(15, values.length); i++) {
      const row = values[i];
      if ((paymentCompanyCol >= 0 && String(row[paymentCompanyCol] || '').includes('지급업체')) ||
          (invoiceCompanyCol >= 0 && String(row[invoiceCompanyCol] || '').includes('업체')) ||
          String(row[0] || '').includes('구분')) {
        dataStartRow = i + 1;
        console.log(`데이터 시작 행: ${dataStartRow + 1}`);
        break;
      }
    }

    // === 지급 내역 수집 ===
    if (paymentAmountCol >= 0 && paymentCompanyCol >= 0) {
      console.log('지급 내역 수집 시작...');
      for (let i = dataStartRow; i < values.length; i++) {
        const company = String((values[i][paymentCompanyCol] || '')).trim();
        const amount = parseAmount(values[i][paymentAmountCol]);
        if (!company || amount <= 0) continue;
        if (company.includes('지급') || company.includes('업체') || company.includes('합계')) continue;

        const entityRecord = ensureEntityRecord(entityData, company);
        if (!entityRecord) continue;

        entityRecord.paidAmount += amount;
        console.log(`지급 내역 - ${company}: ${amount}원 (행 ${i + 1})`);
      }
    }

    // === 발행(청구) 내역 수집 ===
    console.log(`발행 내역 수집 시작... (금액: ${invoiceAmountCol}열, 업체: ${invoiceCompanyCol}열, 발행일: ${invoiceDateCol}열)`);
    for (let i = dataStartRow; i < values.length; i++) {
      const company = String((values[i][invoiceCompanyCol] || '')).trim();
      const amount = parseAmount(values[i][invoiceAmountCol]);

      if (!company || amount <= 0) continue;
      if (company.includes('업체') || company.includes('청구') || company.includes('발행') || company.includes('합계')) continue;

      const entityRecord = ensureEntityRecord(entityData, company);
      if (!entityRecord) continue;

      entityRecord.invoiceAmount += amount;

      let monthKey = '';
      if (invoiceDateCol >= 0) {
        const rawDate = values[i][invoiceDateCol];
        const parsedDate = parseDateValue(rawDate);
        if (parsedDate) monthKey = getMonthKey(parsedDate);
      }

      if (monthKey) {
        if (!entityRecord.invoiceByMonth[monthKey]) entityRecord.invoiceByMonth[monthKey] = 0;
        entityRecord.invoiceByMonth[monthKey] += amount;
        console.log(`발행 내역 - ${company}: ${amount}원 (${monthKey}) (행 ${i + 1})`);
      } else {
        console.log(`발행 내역 - ${company}: ${amount}원 (행 ${i + 1}) - 월 정보 없음`);
      }
    }

    // === 집계 ===
    console.log('\n=== 상세 데이터 집계 결과 ===');
    Object.keys(entityData).forEach(company => {
      const data = entityData[company];
      const invoiceMonths = Object.keys(data.invoiceByMonth || {}).sort();
      let remainingPaid = data.paidAmount || 0;
      const monthlyUnpaid = {};

      invoiceMonths.forEach(monthKey => {
        let invoiceValue = data.invoiceByMonth[monthKey] || 0;
        let unpaidValue = invoiceValue;
        if (remainingPaid > 0) {
          const deduction = Math.min(remainingPaid, unpaidValue);
          unpaidValue -= deduction;
          remainingPaid -= deduction;
        }
        if (unpaidValue > 0) monthlyUnpaid[monthKey] = unpaidValue;
      });

      data.unpaidAmount = Math.max(0, (data.invoiceAmount || 0) - (data.paidAmount || 0));
      data.monthlyUnpaid = monthlyUnpaid;

      console.log(`${company}: 발행 ${data.invoiceAmount}원, 지급 ${data.paidAmount}원, 미지급 ${data.unpaidAmount}원`);

      if (data.unpaidAmount > 0) {
        detailItems.push({
          constructionName: constructionName || '',
          subItem: sheetName,
          entity: company,
          invoiceAmount: data.invoiceAmount || 0,
          paidAmount: data.paidAmount || 0,
          unpaidAmount: data.unpaidAmount || 0,
          note: '',
          monthlyBreakdown: monthlyUnpaid
        });
      }
    });

    detailItems.sort((a, b) => b.unpaidAmount - a.unpaidAmount);
    console.log(`${sheetName} 시트에서 총 ${detailItems.length}개 미지급 항목 발견`);

    return { sheetName, entityData, detailItems };

  } catch (e) {
    console.log(`상세 시트 처리 중 오류: ${filteredNames.join(', ')}`, e.toString());
  }

  return { sheetName: '', entityData, detailItems };
}

function findDetailSheet(spreadsheet, possibleSheetNames) {
  const candidates = (possibleSheetNames || [])
    .map(name => ({
      original: name,
      normalized: normalizeName(name)
    }))
    .filter(item => item.normalized);

  if (candidates.length === 0) {
    return null;
  }

  const sheets = spreadsheet.getSheets();

  for (let candidate of candidates) {
    for (let sheet of sheets) {
      const sheetName = sheet.getName();
      const normalizedSheet = normalizeName(sheetName);

      if (!normalizedSheet) {
        continue;
      }

      if (sheetName === candidate.original ||
          normalizedSheet === candidate.normalized ||
          normalizedSheet.includes(candidate.normalized) ||
          candidate.normalized.includes(normalizedSheet)) {
        return sheet;
      }
    }
  }

  return null;
}

function normalizeName(text) {
  if (!text) {
    return '';
  }
  return String(text)
    .toLowerCase()
    .replace(/[\s\-_.]/g, '')
    .replace(/[()\[\]]/g, '');
}

function findEntityRecord(entityData, targetName) {
  if (!entityData || !targetName) {
    return null;
  }

  const normalizedTarget = normalizeName(targetName);
  if (!normalizedTarget) {
    return null;
  }

  for (let key in entityData) {
    if (!Object.prototype.hasOwnProperty.call(entityData, key)) {
      continue;
    }
    const normalizedKey = normalizeName(key);
    if (!normalizedKey) {
      continue;
    }

    if (normalizedKey === normalizedTarget ||
        normalizedKey.includes(normalizedTarget) ||
        normalizedTarget.includes(normalizedKey)) {
      const data = entityData[key];
      return {
        name: key,
        invoiceAmount: data.invoiceAmount || 0,
        paidAmount: data.paidAmount || 0,
        unpaidAmount: data.unpaidAmount || 0,
        monthlyUnpaid: data.monthlyUnpaid || {},
        invoiceByMonth: data.invoiceByMonth || {}
      };
    }
  }

  return null;
}

function ensureEntityRecord(entityData, company) {
  if (!company) return null;
  if (!entityData[company]) {
    entityData[company] = {
      invoiceAmount: 0,
      paidAmount: 0,
      invoiceByMonth: {},
      monthlyUnpaid: {}
    };
  }
  return entityData[company];
}

function parseDateValue(value) {
  if (!value) return null;

  if (value instanceof Date && !isNaN(value)) return value;

  if (typeof value === 'number') {
    const dateFromNumber = new Date(Math.round((value - 25569) * 86400 * 1000));
    if (!isNaN(dateFromNumber)) return dateFromNumber;
  }

  const str = String(value).trim();
  if (!str) return null;

  const cleaned = str
    .replace(/년|월/g, '-')
    .replace(/일/g, '')
    .replace(/[./\\]/g, '-')
    .replace(/\s+/g, '-')
    .replace(/--+/g, '-')
    .replace(/^-|-$/g, '');

  const parts = cleaned.split('-').filter(Boolean);
  if (parts.length === 3 || parts.length === 2) {
    let [year, month, day] = parts;
    if (parts.length === 2) day = '01';

    if (year.length === 2) year = (parseInt(year, 10) > 50 ? '19' : '20') + year;
    if (month.length === 1) month = '0' + month;
    if (day.length === 1) day = '0' + day;

    const isoString = `${year}-${month}-${day}`;
    const parsed = new Date(isoString);
    if (!isNaN(parsed)) return parsed;
  }

  const fallback = new Date(str);
  if (!isNaN(fallback)) return fallback;

  return null;
}

function getMonthKey(date) {
  if (!(date instanceof Date) || isNaN(date)) {
    return '';
  }
  return Utilities.formatDate(date, 'Asia/Seoul', 'yyyy-MM');
}

function getMonthLabel(monthKey) {
  if (!monthKey) return '';
  try {
    const parts = String(monthKey).split('-');
    if (parts.length !== 2) return String(monthKey);

    const year = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10);
    const date = new Date(year, month - 1, 1);

    if (!isNaN(date)) {
      // ‘25년 9월’ 형식 (원하면 'yyyy년 M월'로 변경 가능)
      return Utilities.formatDate(date, 'Asia/Seoul', 'yy년 M월');
      // return Utilities.formatDate(date, 'Asia/Seoul', 'yyyy년 M월');
    }
  } catch (e) {
    console.log(`월 라벨 변환 오류: ${monthKey}`, e.toString());
  }
  return String(monthKey);
}

/**
 * 금액 파싱 (문자열을 숫자로 변환)
 */
function parseAmount(value) {
  if (typeof value === 'number') return value;
  if (!value) return 0;

  // 문자열에서 숫자만 추출 (콤마, 원 등 제거)
  const numStr = String(value).replace(/[^0-9.-]/g, '');
  const num = parseFloat(numStr);

  return isNaN(num) ? 0 : num;
}

/**
 * 보고서 시트 서식 적용
 */
function formatReportSheet(sheet, dataRows, monthCount) {
  const numericColumnCount = 3 + Math.max(0, monthCount);

  sheet.getRange(2, 5, dataRows + 2, numericColumnCount).setNumberFormat('#,##0');

  sheet.setColumnWidth(1, 50);  // 번호
  sheet.setColumnWidth(2, 180); // 공사
  sheet.setColumnWidth(3, 150); // 업체
  sheet.setColumnWidth(4, 150); // 상세
  sheet.setColumnWidth(5, 130); // 발행(청구)금액
  sheet.setColumnWidth(6, 130); // 지급액
  sheet.setColumnWidth(7, 130); // 미지급액

  for (let idx = 0; idx < monthCount; idx++) {
    sheet.setColumnWidth(8 + idx, 100); // 월별 미지급액
  }

  sheet.setColumnWidth(8 + Math.max(0, monthCount), 200); // 비고

  sheet.getRange(2, 1, dataRows, 8 + monthCount).setBorder(
    true, true, true, true, true, true,
    '#cccccc',
    SpreadsheetApp.BorderStyle.SOLID
  );

  sheet.getRange(2, 7, dataRows, 1).setBackground('#ffe6e6');
  sheet.getRange(2, 5, dataRows, 2).setBackground('#e6f2ff');

  if (monthCount > 0) {
    sheet.getRange(2, 8, dataRows, monthCount).setBackground('#f3f0ff');
  }
}

/**
 * 숫자를 천단위 콤마 형식으로 변환
 */
function formatNumber(num) {
  return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

/**
 * 간단한 미지급액 요약 보기
 */
function showQuickSummary() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 메인 시트 가져오기
  let mainSheet;
  try {
    mainSheet = spreadsheet.getSheetByName(MAIN_SHEET_NAME);
  } catch (e) {
    mainSheet = spreadsheet.getActiveSheet();
  }
  
  const dataRange = mainSheet.getRange('A1:K80');
  const values = dataRange.getValues();
  
  let unpaidCount = 0;
  let totalUnpaid = 0;
  const categories = new Map();
  
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const unpaidAmount = row[8]; // I열
    
    if (typeof unpaidAmount === 'number' && unpaidAmount > 0) {
      unpaidCount++;
      totalUnpaid += unpaidAmount;
      
      const category = row[1] || '기타';
      if (!categories.has(category)) {
        categories.set(category, 0);
      }
      categories.set(category, categories.get(category) + unpaidAmount);
    }
  }
  
  let summary = '📊 미지급액 요약\n' + '─'.repeat(30) + '\n\n';
  
  categories.forEach((amount, category) => {
    summary += `▪ ${category}: ${formatNumber(Math.round(amount))}원\n`;
  });
  
  summary += '\n' + '─'.repeat(30) + '\n';
  summary += `📍 총 ${unpaidCount}개 항목\n`;
  summary += `💰 총 미지급액: ${formatNumber(Math.round(totalUnpaid))}원`;
  
  SpreadsheetApp.getUi().alert('미지급액 요약', summary, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 디버그 모드 실행 (상세 로그 출력)
 */
function runDebugMode() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '디버그 모드',
    '디버그 모드로 실행하시겠습니까?\n\n' +
    '상세한 처리 과정이 로그에 기록됩니다.\n' +
    '실행 후 보기 > 로그에서 확인하세요.',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    console.log('========== 디버그 모드 시작 ==========');
    createFinalUnpaidReport();
    console.log('========== 디버그 모드 종료 ==========');
    ui.alert('디버그 완료', '로그를 확인하려면:\n보기 > 로그를 선택하세요.', ui.ButtonSet.OK);
  }
}

/**
 * 메인 시트 설정
 */
function configureMainSheet() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  
  let sheetList = '현재 시트 목록:\n';
  sheets.forEach((sheet, index) => {
    sheetList += `${index + 1}. ${sheet.getName()}\n`;
  });
  
  const response = ui.prompt(
    '메인 시트 설정',
    `${sheetList}\n현재 설정: ${MAIN_SHEET_NAME}\n\n` +
    '메인 시트 이름을 입력하세요:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const newSheetName = response.getResponseText().trim();
    try {
      const testSheet = spreadsheet.getSheetByName(newSheetName);
      if (testSheet) {
        ui.alert('설정 완료', `메인 시트가 '${newSheetName}'로 설정되었습니다.\n\n` +
                '스크립트 편집기에서 MAIN_SHEET_NAME 변수를 수정해주세요.', ui.ButtonSet.OK);
      }
    } catch (e) {
      ui.alert('오류', `'${newSheetName}' 시트를 찾을 수 없습니다.`, ui.ButtonSet.OK);
    }
  }
}

/**
 * 도움말
 */
function showUnpaidHelp() {
  const helpText = 
    '📊 미지급액 관리 시스템\n' +
    '═══════════════════════════\n\n' +
    '✅ 최종 보고서 생성\n' +
    '• 단일 업체: 메인 시트에서 직접 추출\n' +
    '• 복수 업체: 상세 시트에서 업체별 합산\n\n' +
    '처리 방식:\n' +
    '─────────────\n' +
    '1️⃣ 단일 업체 (한일타워, 대진상사 등)\n' +
    '   → 메인 시트 정보 그대로 사용\n\n' +
    '2️⃣ 복수 업체 (그 외, 인건비, 식대, 보험 등)\n' +
    '   → 상세 시트에서 업체별 건별 합산\n' +
    '   → 발행금액 합계 - 지급금액 합계 = 미지급액\n\n' +
    '📋 간단 요약 보기\n' +
    '• 구분별 미지급액 요약 표시\n\n' +
    '🔍 디버그 모드\n' +
    '• 상세한 처리 과정을 로그에 기록\n' +
    '• 문제 해결시 유용\n\n' +
    '═══════════════════════════';
  
  SpreadsheetApp.getUi().alert('도움말', helpText, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 스크립트 초기화 함수
 * 첨부파일 관리 스크립트에서 호출됩니다
 */
function initializeUnpaidScript() {
  console.log('미지급액 관리 스크립트 초기화 완료');
}
