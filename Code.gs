// ===== 메인 스크립트 파일 (Code.gs) : 최종 완성본 =====

// 전역 변수
const INVOICE_FOLDER_NAME = '스프레드시트_첨부서류';
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

// 3. 파일 업로드 처리 (권한 강제 설정 포함)
function attachFile(base64Data, fileName, mimeType, rowOffset) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const cell = sheet.getActiveCell().offset(rowOffset || 0, 0);
    const cellA1 = cell.getA1Notation();
    
    const folderId = getOrInitFolderId();
    const folder = DriveApp.getFolderById(folderId);
    
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
    const file = folder.createFile(blob);
    
    // [중요] 파일별 권한 강제 설정 (액세스 오류 방지)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
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
    
    return { success: true, message: '업로드 완료', fileCount: attachments.length };
    
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
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('INVOICE_FOLDER_ID');
  if (savedId) { try { DriveApp.getFolderById(savedId); return savedId; } catch (e) {} }
  const folders = DriveApp.getFoldersByName(INVOICE_FOLDER_NAME);
  let folder;
  if (folders.hasNext()) folder = folders.next();
  else {
    folder = DriveApp.createFolder(INVOICE_FOLDER_NAME);
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }
  props.setProperty('INVOICE_FOLDER_ID', folder.getId());
  return folder.getId();
}

function getOrInitArchiveFolder() {
  const mainFolderId = getOrInitFolderId();
  const mainFolder = DriveApp.getFolderById(mainFolderId);
  const folders = mainFolder.getFoldersByName(ARCHIVE_SUBFOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return mainFolder.createFolder(ARCHIVE_SUBFOLDER_NAME);
}
