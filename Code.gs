// ===== 메인 스크립트 파일 (Code.gs) : 최종 완성본 =====

// 전역 변수
const SHARED_DRIVE_FOLDER_ID = '1IxjWPKJaNEkWAZyc6DxqDZvDdoK_QppD'; // 공유드라이브 첨부파일 폴더
const ARCHIVE_SUBFOLDER_NAME = '보관함';

// 1. 스프레드시트 열 때 실행
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('📎 첨부파일 관리')
    .addItem('📤 현재 셀에 파일 첨부', 'showAttachmentDialog')
    .addItem('💾 첨부파일 다운로드 (선택범위)', 'downloadAttachment')
    .addItem('📦 첨부파일 보관', 'archiveAttachments')
    .addItem('🗑️ 첨부파일 완전 삭제', 'deleteAttachmentsPermanently')
    .addSeparator()
    .addItem('📁 첨부파일 폴더 열기', 'openAttachmentFolder')
    .addItem('🔗 저장 폴더 ID 변경', 'changeAttachmentFolderId')
    .addItem('🔄 첨부파일 위치 동기화', 'manualSync')
    .addItem('⚙️ 자동 동기화 트리거 설정', 'setupChangeTrigger')
    .addToUi();
  
  try { if (typeof addUnpaidMenu === 'function') addUnpaidMenu(); } catch (e) {}
  try { if (typeof OCR !== 'undefined' && typeof OCR.addMenu === 'function') OCR.addMenu(); } catch (e) {}
}

// 2. 첨부파일 다이얼로그 표시
function showAttachmentDialog() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('AttachmentDialog')
      .setWidth(550).setHeight(550);
    SpreadsheetApp.getUi().showModalDialog(html, '📎 서류 첨부');
  } catch (e) {
    SpreadsheetApp.getUi().alert('오류', 'AttachmentDialog.html 파일이 필요합니다.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// 3. 파일 업로드 처리 (시트명/열헤더 기반 폴더 구조)
function attachFile(base64Data, fileName, mimeType, rowOffset) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const cell = sheet.getActiveCell().offset(rowOffset || 0, 0);
    const cellA1 = cell.getA1Notation();
    
    // 저장할 폴더 결정 (시트명/열헤더 구조)
    const folder = getTargetFolder(sheet, cell);
    
    // 폴더 내 파일 수를 세어 다음 번호 부여
    const nextNumber = getNextFileNumber(folder);
    const numberedFileName = nextNumber + '. ' + fileName;
    
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, numberedFileName);
    const file = folder.createFile(blob);
    
    // [참고] 공유드라이브는 드라이브 레벨에서 권한이 관리되므로 개별 설정은 건너뜀
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      console.log('공유드라이브 파일 - 개별 권한 설정 건너뜀: ' + e.message);
    }
    
    // 아이콘 설정
    let icon = '📎';
    const lower = fileName.toLowerCase();
    if (lower.includes('계산서') || lower.includes('invoice')) icon = '🧾';
    else if (lower.includes('계약')) icon = '📄';
    else if (lower.includes('입금') || lower.includes('receipt')) icon = '💰';
    else if (lower.includes('청구')) icon = '📋';
    else if (lower.endsWith('.pdf')) icon = '📑';
    else if (lower.match(/\.(jpg|jpeg|png)$/)) icon = '🖼️';
    
    const userProperties = PropertiesService.getUserProperties();
    const key = `${sheet.getSheetId()}_${cellA1}`;
    let attachments = [];
    try { attachments = JSON.parse(userProperties.getProperty(key)) || []; } catch (e) {}
    
    attachments.push({
      fileId: file.getId(),
      fileName: fileName,
      icon: icon,
      uploadDate: new Date().toISOString(),
      fileSize: blob.getBytes().length
    });
    
    userProperties.setProperty(key, JSON.stringify(attachments));
    updateCellDisplay(cell, attachments);
    
    // 저장 경로 안내 메시지
    const sheetName = sheet.getName();
    const colHeader = getColumnHeader(sheet, cell.getColumn());
    let pathMsg;
    if (sheetName === '목차') {
      const companyName = String(sheet.getRange(cell.getRow(), 4).getValue() || '').trim();
      pathMsg = companyName ? `${companyName}/${colHeader}` : '루트 폴더';
    } else {
      pathMsg = `${sheetName}/${colHeader}`;
    }
    
    return { success: true, message: `업로드 완료 (${pathMsg})`, fileCount: attachments.length };
    
  } catch (error) {
    console.error(error);
    return { success: false, message: error.toString() };
  }
}

// 4. 셀 디스플레이 업데이트
function updateCellDisplay(cell, attachments) {
  const count = attachments.length;
  if (count === 0) {
    cell.clear();
    cell.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
    return;
  }
  
  cell.setBackground('#e3f2fd').setFontColor('#1565c0').setFontWeight('bold');
  
  if (count === 1) {
    const f = attachments[0];
    cell.setFormula(`=HYPERLINK("https://drive.google.com/file/d/${f.fileId}/view", "${f.icon} ${f.fileName}")`);
    cell.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
  } else {
    try {
      const r = SpreadsheetApp.newRichTextValue();
      let text = '', current = 0, ranges = [];
      
      attachments.forEach((f, i) => {
        const line = `${i + 1}. ${f.icon} ${f.fileName}`;
        ranges.push({ s: current, e: current + line.length, url: `https://drive.google.com/file/d/${f.fileId}/view` });
        text += line + (i < attachments.length - 1 ? '\n' : '');
        current = text.length + 1;
      });
      
      r.setText(text);
      ranges.forEach(rg => r.setLinkUrl(rg.s, rg.e, rg.url));
      cell.setRichTextValue(r.build());
      cell.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      
      const minH = 21 * count + 10;
      if (cell.getSheet().getRowHeight(cell.getRow()) < minH) cell.getSheet().setRowHeight(cell.getRow(), minH);
    } catch (e) {
      cell.setValue(attachments.map((f, i) => `${i+1}. ${f.icon} ${f.fileName}`).join('\n'));
    }
  }
  
  const note = `📎 파일 ${count}개\n` + attachments.map(f => `- ${f.fileName} (${formatFileSize(f.fileSize)})`).join('\n');
  cell.setNote(note);
}

// 5. 다운로드 (Ctrl+클릭 지원 + IFrame 자동 다운로드)
function downloadAttachment() {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // [핵심] Ctrl+클릭으로 선택된 모든 범위(Range) 리스트를 가져옴
  const rangeList = sheet.getActiveRangeList(); 
  
  if (!rangeList) {
    SpreadsheetApp.getUi().alert('선택된 셀이 없습니다.');
    return;
  }

  const ranges = rangeList.getRanges(); // 선택된 모든 범위 배열
  const ui = SpreadsheetApp.getUi();

  const allProperties = PropertiesService.getUserProperties().getProperties();
  const sheetId = sheet.getSheetId();
  let allFiles = [];

  // 선택된 모든 덩어리(Range)를 순회하며 파일 수집
  ranges.forEach(range => {
    const numRows = range.getNumRows();
    const numCols = range.getNumColumns();
    const startRow = range.getRow();
    const startCol = range.getColumn();

    for (let i = 0; i < numRows; i++) {
      for (let j = 0; j < numCols; j++) {
        const cell = sheet.getRange(startRow + i, startCol + j);
        const key = `${sheetId}_${cell.getA1Notation()}`;
        
        if (allProperties[key]) {
          try {
            const attachments = JSON.parse(allProperties[key]);
            attachments.forEach(f => {
              allFiles.push({
                name: f.fileName,
                url: `https://drive.google.com/uc?export=download&id=${f.fileId}`,
                icon: f.icon || '📄',
                cell: cell.getA1Notation()
              });
            });
          } catch (e) {}
        }
      }
    }
  });

  if (allFiles.length === 0) { 
    ui.alert('파일 없음', '선택한 범위에 첨부파일이 없습니다.', ui.ButtonSet.OK); 
    return; 
  }

  // 팝업 HTML 생성 (Iframe 방식 + 수동 버튼 포함)
  const htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: 'Segoe UI', Tahoma, sans-serif; padding: 20px; background: #f9f9f9; text-align: center; }
        .header { margin-bottom: 20px; }
        h3 { margin: 0; color: #333; font-size: 18px; }
        p { margin: 5px 0 0; font-size: 13px; color: #666; }
        .file-list { max-height: 250px; overflow-y: auto; border: 1px solid #eee; background: white; border-radius: 8px; padding: 10px; margin-bottom: 20px; text-align: left; }
        .file-item { padding: 8px; border-bottom: 1px solid #f0f0f0; display: flex; justify-content: space-between; align-items: center; }
        .file-item:last-child { border-bottom: none; }
        .file-info { display: flex; align-items: center; gap: 8px; font-size: 13px; color: #333; overflow: hidden; }
        .cell-badge { background: #eee; color: #555; padding: 2px 6px; border-radius: 4px; font-size: 11px; font-weight: bold; min-width: 30px; text-align: center; }
        .btn-group { display: flex; gap: 10px; }
        .btn { flex: 1; padding: 12px; border: none; border-radius: 6px; font-weight: bold; cursor: pointer; font-size: 14px; transition: 0.2s; }
        .btn-primary { background: #2196F3; color: white; }
        .btn-primary:hover { background: #1976D2; }
        .btn-primary:disabled { background: #ccc; cursor: wait; }
        .btn-secondary { background: #ddd; color: #333; }
        .btn-secondary:hover { background: #ccc; }
        .manual-btn { text-decoration: none; background: #4CAF50; color: white; padding: 4px 10px; border-radius: 4px; font-size: 12px; font-weight: bold; white-space: nowrap;}
        .manual-btn:hover { background: #45a049; }
        .status-bar { margin-bottom: 15px; background: #e3f2fd; color: #1565c0; padding: 10px; border-radius: 6px; font-size: 13px; display: none; }
      </style>
    </head>
    <body>
      <div class="header">
        <h3>💾 파일 다운로드</h3>
        <p>총 <b>${allFiles.length}</b>개의 파일이 선택되었습니다.</p>
      </div>
      <div id="statusBar" class="status-bar">준비 중...</div>
      <div class="file-list" id="list"></div>
      <div class="btn-group">
        <button class="btn btn-secondary" onclick="google.script.host.close()">닫기</button>
        <button class="btn btn-primary" id="downloadAllBtn" onclick="downloadAll()">📥 전체 다운로드 시작</button>
      </div>
      <script>
        const files = ${JSON.stringify(allFiles)};
        const list = document.getElementById('list');
        const status = document.getElementById('statusBar');
        const btn = document.getElementById('downloadAllBtn');

        files.forEach(file => {
          const div = document.createElement('div');
          div.className = 'file-item';
          div.innerHTML = \`
            <div class="file-info">
              <span class="cell-badge">\${file.cell}</span>
              <span title="\${file.name}">\${file.icon} \${file.name.length > 20 ? file.name.substring(0,20)+'...' : file.name}</span>
            </div>
            <a href="\${file.url}" class="manual-btn" download target="_blank">받기</a>
          \`;
          list.appendChild(div);
        });

        async function downloadAll() {
          btn.disabled = true;
          btn.textContent = '진행 중...';
          status.style.display = 'block';
          
          let count = 0;
          for (const file of files) {
            count++;
            status.textContent = count + ' / ' + files.length + ' 다운로드 시도...';
            
            // iframe을 이용한 강제 다운로드 (브라우저 차단 우회)
            const iframe = document.createElement('iframe');
            iframe.style.display = 'none';
            iframe.src = file.url;
            document.body.appendChild(iframe);
            
            // 메모리 해제
            setTimeout(() => document.body.removeChild(iframe), 60000);
            
            // 딜레이 (1.5초) - 중요: 너무 빠르면 브라우저가 막음
            await new Promise(resolve => setTimeout(resolve, 1500));
          }
          
          status.style.background = '#d4edda';
          status.style.color = '#155724';
          status.innerHTML = '✅ 완료! 반응이 없으면<br>목록 옆의 <b>[받기]</b> 버튼을 누르세요.';
          btn.textContent = '완료';
        }
      </script>
    </body>
    </html>
  `;

  const output = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(500)
    .setHeight(allFiles.length > 5 ? 500 : 350);
    
  ui.showModalDialog(output, `💾 파일 다운로드 (총 ${allFiles.length}개)`);
}

// 6. 보관 기능
function archiveAttachments() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell = sheet.getActiveCell();
  const key = `${sheet.getSheetId()}_${cell.getA1Notation()}`;
  const props = PropertiesService.getUserProperties();
  const data = props.getProperty(key);
  
  if (!data) { ui.alert('파일 없음', '보관할 파일이 없습니다.', ui.ButtonSet.OK); return; }
  const attachments = JSON.parse(data);
  
  if (ui.alert('📦 보관', `${attachments.length}개 파일을 보관함으로 이동합니다.\n링크는 계속 유지됩니다.`, ui.ButtonSet.YES_NO) === ui.Button.YES) {
    try {
      const archiveFolder = getOrInitArchiveFolder();
      const mainFolderId = getOrInitFolderId();
      const mainFolder = DriveApp.getFolderById(mainFolderId);
      
      attachments.forEach(a => {
        try {
          const file = DriveApp.getFileById(a.fileId);
          archiveFolder.addFile(file);
          mainFolder.removeFile(file);
        } catch (e) {}
      });
      
      props.deleteProperty(key);
      cell.clearContent(); cell.clearFormat(); cell.clearNote(); cell.setBackground(null);
      cell.getSheet().setRowHeight(cell.getRow(), 21);
      ui.alert('✅ 보관 완료', '파일이 보관함(하위폴더)으로 이동되었습니다.', ui.ButtonSet.OK);
    } catch (e) { ui.alert('❌ 오류', e.toString(), ui.ButtonSet.OK); }
  }
}

// 7. 완전 삭제
function deleteAttachmentsPermanently() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell = sheet.getActiveCell();
  const key = `${sheet.getSheetId()}_${cell.getA1Notation()}`;
  const props = PropertiesService.getUserProperties();
  const data = props.getProperty(key);
  
  if (!data) { ui.alert('파일 없음', ui.ButtonSet.OK); return; }
  
  if (ui.alert('🗑️ 삭제', '휴지통으로 이동하시겠습니까?', ui.ButtonSet.YES_NO) === ui.Button.YES) {
    JSON.parse(data).forEach(a => { try { DriveApp.getFileById(a.fileId).setTrashed(true); } catch(e){} });
    props.deleteProperty(key);
    cell.clearContent(); cell.clearFormat(); cell.clearNote(); cell.setBackground(null);
    cell.getSheet().setRowHeight(cell.getRow(), 21);
    ui.alert('✅ 삭제 완료', ui.ButtonSet.OK);
  }
}

// 8. 폴더 열기
function openAttachmentFolder() {
  try {
    const folderId = getOrInitFolderId();
    const url = DriveApp.getFolderById(folderId).getUrl();
    const html = `<script>window.open('${url}', '_blank');google.script.host.close();</script>`;
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html), '폴더 여는 중...');
  } catch (e) { SpreadsheetApp.getUi().alert('오류', e.toString(), SpreadsheetApp.getUi().ButtonSet.OK); }
}

// 9. 헬퍼 함수들
function deleteSingleFile(fileId, cellA1) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const key = `${sheet.getSheetId()}_${cellA1}`;
    const props = PropertiesService.getUserProperties();
    const data = props.getProperty(key);
    if (!data) return false;
    let attachments = JSON.parse(data);
    try {
      const file = DriveApp.getFileById(fileId);
      const archiveFolder = getOrInitArchiveFolder();
      const mainFolder = DriveApp.getFolderById(getOrInitFolderId());
      archiveFolder.addFile(file);
      mainFolder.removeFile(file);
    } catch (e) {}
    attachments = attachments.filter(a => a.fileId !== fileId);
    if (attachments.length === 0) {
      props.deleteProperty(key);
      const cell = sheet.getRange(cellA1);
      cell.clearContent(); cell.setBackground(null); cell.clearNote();
    } else {
      props.setProperty(key, JSON.stringify(attachments));
      updateCellDisplay(sheet.getRange(cellA1), attachments);
    }
    return true;
  } catch (e) { return false; }
}

function formatFileSize(bytes) {
  if (!bytes) return '0 B';
  const i = Math.floor(Math.log(bytes) / Math.log(1024));
  return (bytes / Math.pow(1024, i)).toFixed(1) + ' ' + ['B', 'KB', 'MB', 'GB'][i];
}

function getCurrentCellInfo() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell = sheet.getActiveCell();
  const data = PropertiesService.getUserProperties().getProperty(`${sheet.getSheetId()}_${cell.getA1Notation()}`);
  return { sheet: sheet.getName(), row: cell.getRow(), column: cell.getColumn(), a1Notation: cell.getA1Notation(), attachmentCount: data ? JSON.parse(data).length : 0 };
}

function getFileInfo(fileId) {
  try { const f = DriveApp.getFileById(fileId); return { name: f.getName(), url: f.getUrl(), mimeType: f.getMimeType() }; } catch(e) { return null; }
}

function onEdit(e) {
  if (e && ['첨부', 'attach', '@'].includes(e.value)) {
    e.range.clear(); SpreadsheetApp.getActive().setActiveRange(e.range); showAttachmentDialog();
  }
}

function getOrInitFolderId() {
  // ScriptProperties에 저장된 커스텀 폴더 ID 우선 사용
  const customId = PropertiesService.getScriptProperties().getProperty('ATTACHMENT_FOLDER_ID');
  const folderId = customId || SHARED_DRIVE_FOLDER_ID;
  
  try {
    DriveApp.getFolderById(folderId);
    return folderId;
  } catch (e) {
    throw new Error('첨부파일 폴더에 접근할 수 없습니다. 폴더 ID를 확인하세요: ' + folderId);
  }
}

/**
 * 첨부파일 저장 폴더 ID 변경
 * 메뉴에서 호출하여 새로운 구글 드라이브 폴더 ID를 설정
 */
function changeAttachmentFolderId() {
  const ui = SpreadsheetApp.getUi();
  const scriptProps = PropertiesService.getScriptProperties();
  
  // 현재 사용 중인 폴더 ID 표시
  const currentCustomId = scriptProps.getProperty('ATTACHMENT_FOLDER_ID');
  const currentId = currentCustomId || SHARED_DRIVE_FOLDER_ID;
  const isCustom = currentCustomId ? '(사용자 설정)' : '(기본값)';
  
  let currentFolderName = '';
  try {
    currentFolderName = DriveApp.getFolderById(currentId).getName();
  } catch (e) {
    currentFolderName = '(접근 불가)';
  }
  
  const response = ui.prompt(
    '🔗 저장 폴더 ID 변경',
    `현재 폴더: ${currentFolderName} ${isCustom}\n` +
    `현재 ID: ${currentId}\n\n` +
    '폴더 ID 또는 드라이브 URL을 입력하세요.\n' +
    '(예: https://drive.google.com/drive/folders/폴더ID)\n' +
    '(빈칸 입력 시 기본값으로 복원)',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const rawInput = response.getResponseText().trim();
  
  // 빈칸 입력 시 기본값 복원
  if (!rawInput) {
    scriptProps.deleteProperty('ATTACHMENT_FOLDER_ID');
    ui.alert('✅ 복원 완료', '기본 폴더로 복원되었습니다.\nID: ' + SHARED_DRIVE_FOLDER_ID, ui.ButtonSet.OK);
    return;
  }
  
  // URL에서 폴더 ID 추출 (URL이 아니면 입력값 그대로 사용)
  const newId = extractFolderIdFromInput(rawInput);
  
  // 새 폴더 ID 검증
  try {
    const folder = DriveApp.getFolderById(newId);
    const folderName = folder.getName();
    
    // 확인 다이얼로그
    const confirm = ui.alert(
      '📁 폴더 확인',
      `폴더명: ${folderName}\n추출된 ID: ${newId}\n\n이 폴더로 변경하시겠습니까?`,
      ui.ButtonSet.YES_NO
    );
    
    if (confirm === ui.Button.YES) {
      scriptProps.setProperty('ATTACHMENT_FOLDER_ID', newId);
      ui.alert('✅ 변경 완료', `저장 폴더가 변경되었습니다.\n폴더명: ${folderName}`, ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('❌ 오류', '해당 폴더에 접근할 수 없습니다.\n입력값을 다시 확인하세요.\n\n입력값: ' + rawInput + '\n추출된 ID: ' + newId, ui.ButtonSet.OK);
  }
}

/**
 * 입력값에서 구글 드라이브 폴더 ID 추출
 * - URL 형태: https://drive.google.com/drive/folders/폴더ID?usp=... → 폴더ID 추출
 * - ID만 입력: 그대로 반환
 */
function extractFolderIdFromInput(input) {
  if (!input) return '';
  
  // URL 패턴: /folders/ 뒤의 ID 추출 (? 또는 / 이전까지)
  const urlMatch = input.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (urlMatch) {
    console.log('URL에서 폴더 ID 추출: ' + urlMatch[1]);
    return urlMatch[1];
  }
  
  // URL이 아니면 입력값 그대로 반환 (ID로 간주)
  return input;
}

/**
 * 열 번호로 폴더명 결정
 * 목차 시트: L(12)=계약서, M(13)=견적서_내역서, N(14)=사업자등록증, O(15)=통장사본
 * 기타 시트: H(8)=입금증, O(15)=계산서(청구서)
 */
function getColumnHeader(sheet, colIndex) {
  const col = Number(colIndex);
  const sheetName = sheet.getName();
  console.log(`getColumnHeader 호출 - 시트: ${sheetName}, colIndex: ${col}`);
  
  // === 목차 시트 열 매핑 ===
  if (sheetName === '목차') {
    if (col === 12) return '계약서';              // L열
    if (col === 13) return '견적서_내역서';        // M열
    if (col === 14) return '사업자등록증';          // N열
    if (col === 15) return '통장사본';              // O열
  }
  
  // === 기타 시트 열 매핑 ===
  if (col === 8) return '입금증';                  // H열
  if (col === 15) return '계산서(청구서)';          // O열
  
  // 그 외 열은 1행 헤더값 시도
  const headerValue = String(sheet.getRange(1, col).getValue() || '').trim();
  if (headerValue) {
    console.log('헤더값 사용 → ' + headerValue);
    return headerValue;
  }
  
  // 헤더도 없으면 열 문자 사용
  const colLetter = String.fromCharCode(64 + col);
  console.log('폴백 사용 → ' + colLetter + '열');
  return colLetter + '열';
}

/**
 * 첨부파일 저장 대상 폴더 결정
 * - 목차 시트: 공유드라이브/D열 업체명/문서종류/ 구조로 저장
 * - 그 외 시트: 공유드라이브/시트명/열헤더명/ 구조로 저장
 */
function getTargetFolder(sheet, cell) {
  const rootFolder = DriveApp.getFolderById(getOrInitFolderId());
  const sheetName = sheet.getName();
  
  if (sheetName === '목차') {
    // D열(4번째)에서 업체명 가져오기
    const companyName = String(sheet.getRange(cell.getRow(), 4).getValue() || '').trim();
    if (!companyName) {
      console.log('목차 시트 - D열 업체명 없음, 루트 폴더에 저장');
      return rootFolder;
    }
    
    // 1단계: 업체명 폴더 (예: 만방토건)
    const companyFolder = getOrCreateSubFolder(rootFolder, companyName);
    
    // 2단계: 문서종류 폴더 (예: 계약서, 견적서_내역서)
    const docType = getColumnHeader(sheet, cell.getColumn());
    const docFolder = getOrCreateSubFolder(companyFolder, docType);
    
    console.log(`저장 경로: ${companyName}/${docType}`);
    return docFolder;
  }
  
  // 기타 시트: 시트명/열헤더 구조
  // 1단계: 시트명 폴더 (예: 만방토건)
  const sheetFolder = getOrCreateSubFolder(rootFolder, sheetName);
  
  // 2단계: 열 헤더명 폴더 (예: 입금증, 계산서(청구서))
  const colHeader = getColumnHeader(sheet, cell.getColumn());
  const headerFolder = getOrCreateSubFolder(sheetFolder, colHeader);
  
  console.log(`저장 경로: ${sheetName}/${colHeader}`);
  return headerFolder;
}

/**
 * 폴더 내 파일 수를 세어 다음 번호 반환
 */
function getNextFileNumber(folder) {
  const files = folder.getFiles();
  let count = 0;
  while (files.hasNext()) {
    files.next();
    count++;
  }
  return count + 1;
}

/**
 * 하위 폴더를 찾거나 없으면 생성
 */
function getOrCreateSubFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  console.log(`새 폴더 생성: ${folderName}`);
  return parentFolder.createFolder(folderName);
}

function getOrInitArchiveFolder() {
  const mainFolderId = getOrInitFolderId();
  const mainFolder = DriveApp.getFolderById(mainFolderId);
  const folders = mainFolder.getFoldersByName(ARCHIVE_SUBFOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return mainFolder.createFolder(ARCHIVE_SUBFOLDER_NAME);
}

// ===== 첨부파일 위치 동기화 =====

/**
 * onChange 트리거 핸들러
 * 행/열 삽입·삭제 시 자동으로 첨부파일 위치를 동기화
 */
function onSheetChange(e) {
  if (!e) return;
  // 구조 변경(행/열 삽입·삭제)만 처리
  const structuralChanges = ['INSERT_ROW', 'REMOVE_ROW', 'INSERT_COLUMN', 'REMOVE_COLUMN'];
  if (!structuralChanges.includes(e.changeType)) return;
  
  console.log('구조 변경 감지: ' + e.changeType);
  try {
    const moved = syncAttachmentPositions();
    console.log(`동기화 완료: ${moved}개 항목 이동`);
  } catch (err) {
    console.error('동기화 오류: ' + err.toString());
  }
}

/**
 * 수동 동기화 (메뉴에서 호출)
 */
function manualSync() {
  const ui = SpreadsheetApp.getUi();
  try {
    const moved = syncAttachmentPositions();
    if (moved > 0) {
      ui.alert('✅ 동기화 완료', `${moved}개 첨부파일 위치가 갱신되었습니다.`, ui.ButtonSet.OK);
    } else {
      ui.alert('✅ 동기화 완료', '모든 첨부파일 위치가 정상입니다.', ui.ButtonSet.OK);
    }
  } catch (err) {
    ui.alert('❌ 오류', err.toString(), ui.ButtonSet.OK);
  }
}

/**
 * 핵심: 첨부파일 위치 동기화
 * 셀의 하이퍼링크/리치텍스트에서 fileId를 추출하여
 * Properties 키를 현재 셀 위치로 갱신
 */
function syncAttachmentPositions() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const props = PropertiesService.getUserProperties();
  const allProps = props.getProperties();
  let movedCount = 0;
  
  sheets.forEach(sheet => {
    const sheetId = sheet.getSheetId();
    const prefix = sheetId + '_';
    
    // 이 시트의 기존 Properties에서 fileId → 옛 키 매핑
    const fileIdToOldKey = {};
    
    Object.keys(allProps).forEach(key => {
      if (!key.startsWith(prefix)) return;
      try {
        const attachments = JSON.parse(allProps[key]);
        attachments.forEach(att => {
          fileIdToOldKey[att.fileId] = key;
        });
      } catch (e) {}
    });
    
    // 이 시트에 첨부파일이 없으면 건너뛰기
    if (Object.keys(fileIdToOldKey).length === 0) return;
    
    const dataRange = sheet.getDataRange();
    const formulas = dataRange.getFormulas();
    const notes = dataRange.getNotes();
    const numRows = dataRange.getNumRows();
    const numCols = dataRange.getNumColumns();
    
    // 셀 스캔: fileId → 현재 셀 위치 매핑
    const fileIdToNewKey = {};
    
    for (let i = 0; i < numRows; i++) {
      for (let j = 0; j < numCols; j++) {
        const formula = formulas[i][j];
        const note = notes[i][j];
        
        // 첨부파일이 없는 셀은 건너뛰기
        if (!formula && (!note || !note.includes('📎'))) continue;
        
        const cellA1 = sheet.getRange(i + 1, j + 1).getA1Notation();
        const newKey = prefix + cellA1;
        
        // 1) 단일 파일: HYPERLINK 수식에서 fileId 추출
        if (formula && formula.includes('drive.google.com/file/d/')) {
          const match = formula.match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/);
          if (match && fileIdToOldKey[match[1]]) {
            fileIdToNewKey[match[1]] = newKey;
          }
        }
        
        // 2) 복수 파일: 리치텍스트에서 fileId 추출
        if (note && note.includes('📎') && !formula.includes('drive.google.com')) {
          try {
            const richText = sheet.getRange(i + 1, j + 1).getRichTextValue();
            if (richText) {
              const runs = richText.getRuns();
              for (let r = 0; r < runs.length; r++) {
                const url = runs[r].getLinkUrl();
                if (url && url.includes('drive.google.com/file/d/')) {
                  const match = url.match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/);
                  if (match && fileIdToOldKey[match[1]]) {
                    fileIdToNewKey[match[1]] = newKey;
                  }
                }
              }
            }
          } catch (e) {}
        }
      }
    }
    
    // Properties 키 갱신 (옛 키 → 새 키)
    const processedOldKeys = new Set();
    
    Object.keys(fileIdToNewKey).forEach(fileId => {
      const oldKey = fileIdToOldKey[fileId];
      const newKey = fileIdToNewKey[fileId];
      
      if (oldKey && oldKey !== newKey && !processedOldKeys.has(oldKey)) {
        processedOldKeys.add(oldKey);
        const data = props.getProperty(oldKey);
        if (data) {
          props.deleteProperty(oldKey);
          props.setProperty(newKey, data);
          movedCount++;
          console.log(`위치 이동: ${oldKey.replace(prefix, '')} → ${newKey.replace(prefix, '')}`);
        }
      }
    });
  });
  
  return movedCount;
}

/**
 * onChange 트리거 설치 (최초 1회만 실행)
 * Apps Script 편집기에서 이 함수를 한 번 실행하세요
 */
function setupChangeTrigger() {
  const ui = SpreadsheetApp.getUi();
  
  // 기존 onChange 트리거 제거
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  triggers.forEach(trigger => {
    if (trigger.getEventType() === ScriptApp.EventType.ON_CHANGE) {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });
  
  // 새 onChange 트리거 설치
  ScriptApp.newTrigger('onSheetChange')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onChange()
    .create();
  
  ui.alert(
    '✅ 트리거 설정 완료',
    '행 삽입/삭제 시 첨부파일 위치가 자동으로 동기화됩니다.\n\n' +
    (removed > 0 ? `(기존 트리거 ${removed}개 제거 후 재설치)` : ''),
    ui.ButtonSet.OK
  );
}
