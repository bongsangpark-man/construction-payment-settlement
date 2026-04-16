// ============================================
// 드롭다운 D7 - 데이터 표시 + 자동 시트 이동
// 항목명 = 시트명으로 자동 매칭
// D열에서 항목 찾기
// ============================================

// 드롭다운 설정
function setupDropdown() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var dropdownCell = sheet.getRange('D7');
  
  // 기존 데이터 확인 규칙 제거
  dropdownCell.clearDataValidations();
  
  // D열 전체 데이터 가져오기
  var lastRow = sheet.getLastRow();
  var dColumnData = sheet.getRange(1, 4, lastRow, 1).getValues(); // 4번째 열 = D열
  
  var items = [];
  var itemRows = {};
  
  // D열에서 항목 찾기
  for (var i = 0; i < dColumnData.length; i++) {
    var value = dColumnData[i][0];
    
    if (value && value !== '') {
      var itemName = String(value).trim();
      
      // 헤더나 숫자 제외
      if (itemName !== '항목' && 
          itemName !== '구분' && 
          !itemName.match(/^[\d\.,]+$/)) {
        
        if (!itemRows[itemName]) {
          items.push(itemName);
          itemRows[itemName] = i + 1;
        }
      }
    }
  }
  
  if (items.length === 0) {
    SpreadsheetApp.getUi().alert('D열에서 항목을 찾을 수 없습니다.');
    return;
  }
  
  // 항목 정보 저장
  PropertiesService.getScriptProperties().setProperty('itemRows', JSON.stringify(itemRows));
  
  // 드롭다운 생성
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(items)
    .setAllowInvalid(false)
    .build();
  
  dropdownCell.setDataValidation(rule);
  dropdownCell.setValue('항목 선택');
  dropdownCell.setBackground('#e8f0fe')
              .setFontWeight('bold')
              .setFontColor('#1a73e8');
  
  // 데이터 영역 초기화
  sheet.getRange('E7:K7').clearContent();
  
  // 시트 이동 버튼 영역 설정
  sheet.getRange('D8').setValue('📄 시트 이동 →')
       .setFontWeight('bold')
       .setFontColor('#5f6368');
  
  sheet.getRange('E8').setValue('[클릭하여 이동]')
       .setFontColor('#1a73e8')
       .setFontStyle('italic')
       .setBackground('#f0f7ff')
       .setBorder(true, true, true, true, false, false);
  
  // 이동 버튼에 메모 추가
  sheet.getRange('E8').setNote('선택한 항목의 시트로 이동하려면 클릭하세요');
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ 드롭다운 생성 완료\n' + items.length + '개 항목 발견', 
    '완료', 
    3
  );
}

// 드롭다운 선택 시 실행
function onEdit(e) {
  if (!e) return;
  
  var range = e.range;
  var sheet = e.source.getActiveSheet();
  
  // D7 셀 선택 시
  if (range.getRow() === 7 && range.getColumn() === 4) {
    var selectedItem = range.getValue();
    
    if (selectedItem && selectedItem !== '항목 선택') {
      showItemData(selectedItem);
      updateSheetLink(selectedItem);
    } else {
      sheet.getRange('E7:K7').clearContent();
      sheet.getRange('E8').setValue('[클릭하여 이동]')
           .setFontColor('#1a73e8');
    }
  }
  
  // E8 셀 클릭 시 (시트 이동)
  if (range.getRow() === 8 && range.getColumn() === 5) {
    var selectedItem = sheet.getRange('D7').getValue();
    if (selectedItem && selectedItem !== '항목 선택') {
      goToSheet(selectedItem);
    }
  }
}

// 시트 링크 업데이트
function updateSheetLink(itemName) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 항목명과 같은 이름의 시트가 있는지 확인
  var targetSheet = spreadsheet.getSheetByName(itemName);
  
  if (targetSheet) {
    sheet.getRange('E8').setValue('📑 [' + itemName + ' 시트로 이동]')
         .setFontColor('#0b5394')
         .setFontWeight('bold')
         .setBackground('#cfe2f3');
  } else {
    sheet.getRange('E8').setValue('⚠️ [' + itemName + ' 시트 없음 - 클릭하여 생성]')
         .setFontColor('#cc0000')
         .setBackground('#fce5cd');
  }
}

// 선택한 항목의 데이터 표시
function showItemData(itemName) {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  try {
    var itemRows = JSON.parse(PropertiesService.getScriptProperties().getProperty('itemRows') || '{}');
    var targetRow = itemRows[itemName];
    
    if (!targetRow) {
      var lastRow = sheet.getLastRow();
      for (var i = 1; i <= lastRow; i++) {
        var cellValue = sheet.getRange(i, 4).getValue(); // D열 = 4번째 열
        if (String(cellValue).trim() === String(itemName).trim()) {
          targetRow = i;
          break;
        }
      }
    }
    
    if (!targetRow) {
      sheet.getRange('E7').setValue('데이터 없음');
      return;
    }
    
    // E열부터 K열까지 데이터 가져오기
    var rowData = sheet.getRange(targetRow, 5, 1, 7).getValues()[0];
    
    // E7부터 K7에 데이터 표시
    sheet.getRange('E7:K7').setValues([rowData]);
    sheet.getRange('E7:K7').setNumberFormat('#,##0')
                           .setBackground('#f8f9fa')
                           .setBorder(true, true, true, true, true, true);
    
  } catch (error) {
    console.error('Error:', error);
    sheet.getRange('E7').setValue('오류: ' + error.toString());
  }
}

// 시트로 이동 (E8 클릭 또는 메뉴 실행)
function goToSheet(itemName) {
  if (!itemName) {
    var sheet = SpreadsheetApp.getActiveSheet();
    itemName = sheet.getRange('D7').getValue();
  }
  
  if (!itemName || itemName === '항목 선택') {
    SpreadsheetApp.getUi().alert('알림', 'D7 셀에서 항목을 먼저 선택하세요.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = spreadsheet.getSheetByName(itemName);
  
  if (!targetSheet) {
    // 시트 생성 여부 확인
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('시트 생성', 
      '"' + itemName + '" 시트가 없습니다.\n새로 만드시겠습니까?', 
      ui.ButtonSet.YES_NO);
    
    if (response === ui.Button.YES) {
      targetSheet = createDetailSheet(itemName);
    } else {
      return;
    }
  }
  
  // 시트로 이동
  spreadsheet.setActiveSheet(targetSheet);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ ' + itemName + ' 시트로 이동했습니다.', 
    '이동 완료', 
    2
  );
}

// 상세 시트 생성
function createDetailSheet(itemName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var newSheet = spreadsheet.insertSheet(itemName);
  
  // 헤더 설정
  newSheet.getRange('A1').setValue(itemName + ' 상세 정보')
          .setFontSize(18)
          .setFontWeight('bold')
          .setFontColor('#1a73e8');
  
  newSheet.getRange('A3').setValue('생성일: ' + new Date().toLocaleDateString('ko-KR'))
          .setFontColor('#5f6368');
  
  // 메인으로 돌아가기 버튼
  newSheet.getRange('A5').setValue('🏠 메인으로 돌아가기')
          .setFontColor('#1a73e8')
          .setFontWeight('bold')
          .setBackground('#e8f0fe')
          .setBorder(true, true, true, true, false, false);
  
  newSheet.getRange('A5').setNote('클릭하여 메인 시트로 돌아가기');
  
  // 기본 테이블 헤더
  var headers = [['날짜', '구분', '내용', '금액', '담당자', '비고']];
  newSheet.getRange('A7:F7').setValues(headers)
          .setFontWeight('bold')
          .setBackground('#e8f0fe')
          .setBorder(true, true, true, true, true, true);
  
  // 컬럼 너비 조정
  newSheet.setColumnWidth(1, 100);
  newSheet.setColumnWidth(2, 100);
  newSheet.setColumnWidth(3, 250);
  newSheet.setColumnWidth(4, 120);
  newSheet.setColumnWidth(5, 100);
  newSheet.setColumnWidth(6, 200);
  
  return newSheet;
}

// 메인 시트로 돌아가기
function goToMainSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  
  // 첫 번째 시트를 메인으로 가정
  spreadsheet.setActiveSheet(sheets[0]);
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ 메인 시트로 돌아왔습니다.', 
    '이동 완료', 
    2
  );
}

// 모든 항목의 시트 확인
function checkAllSheets() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 드롭다운 항목 가져오기
  var itemRows = JSON.parse(PropertiesService.getScriptProperties().getProperty('itemRows') || '{}');
  var items = Object.keys(itemRows);
  
  if (items.length === 0) {
    SpreadsheetApp.getUi().alert('먼저 드롭다운을 생성하세요.');
    return;
  }
  
  var existingSheets = [];
  var missingSheets = [];
  
  items.forEach(function(itemName) {
    var targetSheet = spreadsheet.getSheetByName(itemName);
    if (targetSheet) {
      existingSheets.push(itemName);
    } else {
      missingSheets.push(itemName);
    }
  });
  
  var message = '📊 시트 확인 결과\n\n';
  message += '✅ 존재하는 시트 (' + existingSheets.length + '개):\n';
  message += existingSheets.join(', ') + '\n\n';
  message += '❌ 없는 시트 (' + missingSheets.length + '개):\n';
  message += missingSheets.join(', ');
  
  SpreadsheetApp.getUi().alert('시트 확인', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// 없는 시트 모두 생성
function createAllMissingSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var itemRows = JSON.parse(PropertiesService.getScriptProperties().getProperty('itemRows') || '{}');
  var items = Object.keys(itemRows);
  
  if (items.length === 0) {
    SpreadsheetApp.getUi().alert('먼저 드롭다운을 생성하세요.');
    return;
  }
  
  var createdSheets = [];
  
  items.forEach(function(itemName) {
    var targetSheet = spreadsheet.getSheetByName(itemName);
    if (!targetSheet) {
      createDetailSheet(itemName);
      createdSheets.push(itemName);
    }
  });
  
  if (createdSheets.length > 0) {
    SpreadsheetApp.getUi().alert(
      '시트 생성 완료', 
      createdSheets.length + '개 시트 생성:\n' + createdSheets.join(', '), 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert('알림', '모든 시트가 이미 존재합니다.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// D열의 모든 항목에 시트 링크 추가
function addSheetLinksToColumn() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var lastRow = sheet.getLastRow();
  
  for (var i = 1; i <= lastRow; i++) {
    var cellValue = sheet.getRange(i, 4).getValue(); // D열 = 4번째 열
    
    if (cellValue && cellValue !== '' && 
        cellValue !== '항목' && cellValue !== '구분' &&
        !String(cellValue).match(/^[\d\.,]+$/)) {
      
      var itemName = String(cellValue).trim();
      var targetSheet = spreadsheet.getSheetByName(itemName);
      
      if (targetSheet) {
        // 시트가 있으면 링크 스타일 적용
        sheet.getRange(i, 4)
             .setFontColor('#1a73e8')
             .setFontStyle('italic')
             .setBackground('#e8f0fe');
        
        // 메모 추가
        sheet.getRange(i, 4).setNote('📑 ' + itemName + ' 시트로 이동 가능');
      }
    }
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ D열 항목에 시트 링크 스타일 적용 완료', 
    '완료', 
    3
  );
}

// ===== onOpen 함수 삭제됨 =====
// 이제 Code.gs에서 통합 메뉴를 관리합니다
// ==============================

// 수동 데이터 표시
function manualShowData() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var selectedItem = sheet.getRange('D7').getValue();
  
  if (!selectedItem || selectedItem === '항목 선택') {
    SpreadsheetApp.getUi().alert('D7 셀에서 항목을 먼저 선택하세요.');
    return;
  }
  
  showItemData(selectedItem);
  updateSheetLink(selectedItem);
}

// 사용 방법
function showHelp() {
  var html = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: 'Malgun Gothic', sans-serif;
            padding: 20px;
            line-height: 1.8;
          }
          h3 {
            color: #1a73e8;
            border-bottom: 2px solid #1a73e8;
            padding-bottom: 10px;
          }
          .section {
            margin: 20px 0;
            padding: 15px;
            background: #f8f9fa;
            border-radius: 8px;
          }
          .highlight {
            background: #fef7e0;
            padding: 2px 4px;
            border-radius: 3px;
            font-weight: bold;
            color: #ea4335;
          }
          .link-style {
            color: #1a73e8;
            text-decoration: underline;
            font-style: italic;
          }
          code {
            background: #e8f0fe;
            padding: 2px 6px;
            border-radius: 3px;
            color: #1a73e8;
          }
        </style>
      </head>
      <body>
        <h3>📖 사용 방법</h3>
        
        <div class="section">
          <h4>✨ 핵심 기능</h4>
          <p>• <span class="highlight">항목명 = 시트명</span> 자동 매칭</p>
          <p>• D7 드롭다운 선택 → E7~K7 데이터 표시</p>
          <p>• <span class="link-style">E8 셀 클릭</span> → 해당 시트로 이동</p>
          <p>• 시트가 없으면 자동 생성 옵션</p>
        </div>
        
        <div class="section">
          <h4>1️⃣ 초기 설정</h4>
          <ol>
            <li>메뉴 → <code>📌 D7에 드롭다운 생성</code></li>
            <li>D열의 모든 항목이 드롭다운에 추가됨</li>
          </ol>
        </div>
        
        <div class="section">
          <h4>2️⃣ 시트 이동 방법</h4>
          <p><strong>방법 1:</strong> D7에서 항목 선택 → <span class="link-style">E8 셀 클릭</span></p>
          <p><strong>방법 2:</strong> 메뉴 → 시트 이동 → 선택 항목 시트로 이동</p>
          <p><strong>방법 3:</strong> D열의 항목 직접 클릭 (링크 스타일 적용 후)</p>
        </div>
        
        <div class="section" style="background: #e8f0fe;">
          <h4>⚙️ 시트 관리 기능</h4>
          <p><strong>모든 시트 확인:</strong> 어떤 시트가 있고 없는지 확인</p>
          <p><strong>없는 시트 모두 생성:</strong> 한 번에 모든 시트 생성</p>
          <p><strong>D열에 링크 스타일:</strong> 시트가 있는 항목을 파란색으로 표시</p>
        </div>
        
        <div class="section" style="background: #fef7e0;">
          <h4>💡 팁</h4>
          <p>• 시트명과 항목명이 정확히 일치해야 합니다</p>
          <p>• E8 셀이 <span style="color: #0b5394;">[시트로 이동]</span>로 표시되면 시트 존재</p>
          <p>• E8 셀이 <span style="color: #cc0000;">[시트 없음]</span>으로 표시되면 클릭하여 생성</p>
          <p>• 생성된 시트에는 기본 템플릿이 자동 적용됩니다</p>
        </div>
      </body>
    </html>`;
  
  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(700);
  
  SpreadsheetApp.getUi()
    .showModalDialog(htmlOutput, '사용 방법');
}

// 초기 실행 함수
function myFunction() {
  setupDropdown();
  SpreadsheetApp.getUi().alert(
    '초기 설정 완료',
    '✅ 드롭다운이 D7 셀에 생성되었습니다.\n\n' +
    '사용 방법:\n' +
    '1. D7에서 항목 선택\n' +
    '2. E7:K7에 데이터가 자동 표시됨\n' +
    '3. E8 클릭하여 해당 시트로 이동\n\n' +
    '추가 기능은 메뉴 → 📊 데이터 관리를 확인하세요.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}