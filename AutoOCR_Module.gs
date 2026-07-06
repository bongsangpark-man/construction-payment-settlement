/**
 * AutoOCR_Module.gs [v41 - 카드사용내역 멀티행 OCR 추가]
 * 
 * Vision API + 정규식/블랙리스트 → Gemini 2.5 Flash 멀티모달로 전환
 * - 이미지/PDF를 직접 Gemini에 전송 → OCR + 구조화 추출 동시 처리
 * - 블랙리스트/정규식 제거, 프롬프트 기반 의미 추출
 * - [v41] 카드(페이)사용내역 PDF → 멀티 행 추출/입력 기능 추가
 * 
 * (이전: v40 - Gemini 2.0 Flash Direct)
 */
var OCR = (function () {
  'use strict';

  // ===================== [1. 설정] =====================
  // ⚠️ 보안 경고: API 키는 가급적 '스크립트 속성'을 이용하세요.
  var API_KEY = 'AIzaSyBwlYgVYL_lszUwNT7mlpHVrF6PyWE14RE';
  
  var GEMINI_MODEL = 'gemini-2.5-flash';
  var GEMINI_ENDPOINT = 'https://generativelanguage.googleapis.com/v1beta/models/'
                        + GEMINI_MODEL + ':generateContent';

  var RAW_OCR_TO_SHEET = true;
  var RAW_OCR_SHEET_NAME = '📄_RAW_OCR';

  // 우리 회사 이름 (프롬프트에서 제외 대상으로 사용)
  var OUR_COMPANY_NAME = '가현종합건설';

  // ===================== [2. 메인 실행 함수] =====================
  function runAutoOCR(fileId, row, column) {
    var runId = (Utilities.getUuid() || '').slice(0, 8);
    _dlog('OCR', '=== Gemini OCR 시작 ===', { runId: runId, fileId: fileId });

    try {
      if (column !== 8 && column !== 15) {
        _log('ERROR', '허용되지 않은 열(허용: H=8, O=15) → 중단');
        return;
      }

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getActiveSheet();
      var sheetName = sheet.getName();
      var sheetId = sheet.getSheetId();

      var columnType = (column === 8) ? 'payment' : 'invoice';
      ss.toast('🤖 Gemini 문서 분석 중...', columnType === 'payment' ? '지급내역' : '증빙', 5);

      // --- Gemini 호출 (OCR + 추출 동시) ---
      var result = _extractWithGemini(fileId, runId);

      if (!result) {
        _log('ERROR', 'Gemini 추출 결과 없음');
        ss.toast('❌ 실패', '내용을 읽을 수 없습니다', 3);
        return;
      }

      // RAW 로그 기록
      var a1 = sheet.getRange(row, column).getA1Notation();
      _logRawOCRText_(runId, fileId, a1, JSON.stringify(result, null, 2), sheet);

      // 시트 참조 재확인 (RAW 로그 기록 시 시트가 변경될 수 있음)
      sheet = ss.getSheetByName(sheetName) || _getSheetById_(ss, sheetId) || sheet;

      var docType = result.docType || columnType;
      _dlog('INFO', '문서 유형: ' + docType, result);

      if (!result.amount && !result.company) {
        _log('WARNING', '데이터 추출 실패');
        ss.toast('⚠️ 확인 필요', '데이터를 찾지 못했습니다', 3);
        return;
      }

      // 카드이면서 H열(지급)인 경우 별도 타입 지정
      if (docType === 'card' && column === 8) {
        docType = 'card_payment_auto_create';
      }

      // Gemini 응답을 기존 _autoFillData 형식에 맞게 변환
      var extracted = {
        amount:     result.amount || null,
        date:       result.date || null,
        company:    result.company || '',
        merchant:   result.company || '',  // 카드용 호환
        account:    result.account || '',
        cardMasked: (docType === 'card' || docType === 'card_payment_auto_create') ? (result.account || '') : ''
      };

      _autoFillData(sheet, extracted, row, column, docType);

      var finalName = extracted.company || '완료';
      ss.toast('✅ 입력 완료', finalName, 3);
      _log('SUCCESS', '처리 완료 (' + finalName + ')');

    } catch (e) {
      _dlog('ERROR', '시스템 오류', { error: String(e), stack: e.stack });
      SpreadsheetApp.getActiveSpreadsheet().toast('⚠️ 오류 발생', '로그 확인', 5);
    }
  }

  // ===================== [3. Gemini API 호출] =====================
  function _extractWithGemini(fileId, runId) {
    try {
      var file = DriveApp.getFileById(fileId);
      var blob = file.getBlob();
      var base64 = Utilities.base64Encode(blob.getBytes());
      var mime = (blob.getContentType() || 'image/jpeg').toLowerCase();

      // PDF의 경우 mime 확인
      if (mime === 'application/pdf') {
        // Gemini는 PDF를 직접 지원
        mime = 'application/pdf';
      }

      var prompt = [
        '당신은 건설 공사대금 정산 시스템의 문서 분석 전문가입니다.',
        '첨부된 문서에서 다음 정보를 정확하게 추출하세요.',
        '',
        '## 문서 유형 판별',
        '- "invoice": 세금계산서 (전자세금계산서, 수정세금계산서 포함)',
        '- "card": 카드 영수증, 카드 매출전표',
        '- "payment": 입출금 명세서, 이체 확인서, 입금증',
        '',
        '## 추출 규칙',
        '1. company: 거래 상대방의 상호명(회사명)을 추출합니다.',
        '   - "' + OUR_COMPANY_NAME + '"은 우리 회사이므로 반드시 제외하고, 거래 상대방 이름만 추출하세요.',
        '   - 세금계산서: "공급자" 쪽의 상호를 추출 (우리가 공급받는 자이므로)',
        '   - 카드영수증: 가맹점명을 추출',
        '   - 입출금명세서: 송금/입금 상대방 이름을 추출',
        '   - "(주)", "주식회사" 등은 포함해도 됩니다.',
        '',
        '2. amount: 최종 합계금액/결제금액/이체금액 (숫자만, 쉼표 없이)',
        '   - 세금계산서: "합계금액" 또는 "청구금액"의 값',
        '   - 카드: "합계" 또는 "결제금액"의 값',
        '   - 입출금: "출금액" 또는 "이체금액"의 값 (잔액은 제외)',
        '',
        '3. date: 거래일자 (YYYY-MM-DD 형식)',
        '   - 세금계산서: "작성일자"',
        '   - 카드: "승인일시" 또는 "거래일시"',  
        '   - 입출금: "거래일" 또는 "이체일"',
        '',
        '4. account: 계좌번호 또는 카드번호 (있으면 추출, 없으면 빈 문자열)',
        '',
        '## 주의사항',
        '- 금액에서 부가세, 공급가액이 아닌 "합계금액"을 추출하세요.',
        '- 날짜가 여러 개면 가장 메인이 되는 거래일자를 선택하세요.',
        '- 읽을 수 없거나 해당 정보가 없으면 빈 문자열("")을 반환하세요.',
        '- 반드시 JSON만 반환하세요. 다른 텍스트는 포함하지 마세요.'
      ].join('\n');

      var responseSchema = {
        type: 'OBJECT',
        properties: {
          docType:  { type: 'STRING' },
          company:  { type: 'STRING' },
          amount:   { type: 'STRING' },
          date:     { type: 'STRING' },
          account:  { type: 'STRING' }
        },
        required: ['docType', 'company', 'amount', 'date', 'account']
      };

      var payload = {
        contents: [{
          parts: [
            { text: prompt },
            { inlineData: { mimeType: mime, data: base64 } }
          ]
        }],
        generationConfig: {
          responseMimeType: 'application/json',
          responseSchema: responseSchema,
          temperature: 0
        }
      };

      _dlog('GEMINI', 'API 호출 시작', { fileId: fileId, mime: mime, runId: runId });

      var response = UrlFetchApp.fetch(GEMINI_ENDPOINT + '?key=' + API_KEY, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      var httpCode = response.getResponseCode();
      var body = response.getContentText();

      if (httpCode !== 200) {
        _dlog('GEMINI_ERROR', 'HTTP ' + httpCode, { body: body.substring(0, 500) });
        return null;
      }

      var parsed = JSON.parse(body);

      // candidates[0].content.parts[0].text에서 JSON 추출
      if (parsed.candidates && parsed.candidates[0] &&
          parsed.candidates[0].content && parsed.candidates[0].content.parts &&
          parsed.candidates[0].content.parts[0]) {
        
        var textResult = parsed.candidates[0].content.parts[0].text;
        _dlog('GEMINI', 'Raw 응답', { text: textResult });

        var extracted = JSON.parse(textResult);

        // 금액 정리: 쉼표/공백 제거, 숫자만
        if (extracted.amount) {
          extracted.amount = String(extracted.amount).replace(/[^0-9]/g, '');
        }

        // 날짜 형식 통일
        if (extracted.date) {
          extracted.date = String(extracted.date).replace(/[./]/g, '-');
        }

        _dlog('GEMINI', '추출 완료', extracted);
        return extracted;
      }

      _dlog('GEMINI_ERROR', '응답 구조 이상', { parsed: JSON.stringify(parsed).substring(0, 500) });
      return null;

    } catch (e) {
      _dlog('GEMINI_ERROR', 'API 호출 실패', { error: String(e), stack: e.stack });
      return null;
    }
  }

  // ===================== [3-B. 카드사용내역 전용 Gemini 추출] =====================
  /**
   * 카드(페이)사용내역 PDF에서 모든 거래를 배열로 추출
   * - 멀티 페이지 PDF 지원
   * - 각 거래: {date, merchant, amount}
   */
  function _extractCardStatementWithGemini(fileId, runId) {
    try {
      var file = DriveApp.getFileById(fileId);
      var blob = file.getBlob();
      var base64 = Utilities.base64Encode(blob.getBytes());
      var mime = (blob.getContentType() || 'application/pdf').toLowerCase();

      var prompt = [
        '당신은 건설 공사대금 정산 시스템의 카드사용내역 분석 전문가입니다.',
        '첨부된 문서는 "카드(페이)사용내역"으로, 여러 건의 카드 거래가 테이블 형태로 나열되어 있습니다.',
        '',
        '## 작업',
        '문서의 모든 페이지에서 모든 거래 행을 빠짐없이 추출하세요.',
        '',
        '## 추출 규칙',
        '각 거래에서 다음 3가지를 추출합니다:',
        '',
        '1. **date**: 사용일자 (YYYY-MM-DD 형식)',
        '2. **merchant**: 사용처 (가맹점/업체명)',
        '3. **amount**: 사용금액 (숫자만, 쉼표/원 제거)',
        '',
        '## 주의사항',
        '- 모든 페이지의 모든 거래를 누락 없이 추출하세요.',
        '- 합계/소계 행은 제외하세요.',
        '- 헤더 행(사용일자, 사용처, 카드별칭 등)은 제외하세요.',
        '- "총 N건" 같은 요약 정보는 제외하세요.',
        '- 금액은 숫자만 반환하세요 (쉼표, 원, 공백 제거).',
        '- 반드시 JSON 배열만 반환하세요. 다른 텍스트는 포함하지 마세요.'
      ].join('\n');

      var responseSchema = {
        type: 'ARRAY',
        items: {
          type: 'OBJECT',
          properties: {
            date:     { type: 'STRING' },
            merchant: { type: 'STRING' },
            amount:   { type: 'STRING' }
          },
          required: ['date', 'merchant', 'amount']
        }
      };

      var payload = {
        contents: [{
          parts: [
            { text: prompt },
            { inlineData: { mimeType: mime, data: base64 } }
          ]
        }],
        generationConfig: {
          responseMimeType: 'application/json',
          responseSchema: responseSchema,
          temperature: 0
        }
      };

      _dlog('CARD_STMT', 'Gemini API 호출 시작 (카드사용내역)', { fileId: fileId, mime: mime, runId: runId });

      var response = UrlFetchApp.fetch(GEMINI_ENDPOINT + '?key=' + API_KEY, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      var httpCode = response.getResponseCode();
      var body = response.getContentText();

      if (httpCode !== 200) {
        _dlog('CARD_STMT_ERROR', 'HTTP ' + httpCode, { body: body.substring(0, 500) });
        return null;
      }

      var parsed = JSON.parse(body);

      if (parsed.candidates && parsed.candidates[0] &&
          parsed.candidates[0].content && parsed.candidates[0].content.parts &&
          parsed.candidates[0].content.parts[0]) {

        var textResult = parsed.candidates[0].content.parts[0].text;
        _dlog('CARD_STMT', 'Raw 응답 길이: ' + textResult.length);

        var transactions = JSON.parse(textResult);

        if (!Array.isArray(transactions)) {
          _dlog('CARD_STMT_ERROR', '응답이 배열이 아님', { type: typeof transactions });
          return null;
        }

        // 금액/날짜 정리
        transactions = transactions.map(function(tx) {
          if (tx.amount) tx.amount = String(tx.amount).replace(/[^0-9]/g, '');
          if (tx.date) tx.date = String(tx.date).replace(/[./]/g, '-');
          return tx;
        }).filter(function(tx) {
          // 빈 거래 제거
          return tx.amount && tx.merchant;
        });

        _dlog('CARD_STMT', '추출 완료: ' + transactions.length + '건', transactions.slice(0, 3));
        return transactions;
      }

      _dlog('CARD_STMT_ERROR', '응답 구조 이상', { parsed: JSON.stringify(parsed).substring(0, 500) });
      return null;

    } catch (e) {
      _dlog('CARD_STMT_ERROR', 'API 호출 실패', { error: String(e), stack: e.stack });
      return null;
    }
  }

  // ===================== [4. 시트 입력 (매핑 + 스마트매칭 + 자동생성)] =====================
  function _autoFillData(sheet, data, row, column, docType) {
    try {
      var extractedDate = data.date;
      var extractedAmount = _toNumberValue(data.amount);
      var extractedCompany = data.company || data.merchant || '';
      var extractedAccount = '';

      if (data.cardMasked) extractedAccount = data.cardMasked;
      else if (data.account) extractedAccount = data.account;

      if (column === 8) {
        // H열 (지급내역): B=금액, C=날짜, D=계좌, F=업체명
        if (extractedAmount) sheet.getRange(row, 2).setValue(extractedAmount);
        if (extractedDate) sheet.getRange(row, 3).setValue(extractedDate);
        if (extractedAccount.trim()) sheet.getRange(row, 4).setValue(extractedAccount.trim());

        var finalCompany = '';

        // [1순위: 금액 매칭] - K열(발행내역)에서 동일 금액 찾기
        if (docType !== 'card_payment_auto_create' && extractedAmount) {
          var lastRow = sheet.getLastRow();
          if (lastRow > 0) {
            var rangeK = sheet.getRange(1, 11, lastRow, 1).getValues();
            var rangeM = sheet.getRange(1, 13, lastRow, 1).getValues();
            for (var i = 0; i < lastRow; i++) {
              var invoiceAmt = _toNumberValue(rangeK[i][0]);
              if (invoiceAmt === extractedAmount) {
                var invoiceCompany = rangeM[i][0];
                if (invoiceCompany && String(invoiceCompany).trim() !== '') {
                  finalCompany = invoiceCompany;
                  _dlog('MATCH', '발행내역(K열) 매칭 성공: ' + finalCompany);
                  break;
                }
              }
            }
          }
        }

        // [2순위: 단일 업체 시트 감지]
        if (!finalCompany && docType !== 'card_payment_auto_create') {
          var dominant = _getDominantCompanyName(sheet);
          if (dominant) {
            finalCompany = dominant;
            _dlog('MATCH', '단일 업체 시트 감지됨: ' + finalCompany);
          }
        }

        // [3순위: 유사 업체명 보정]
        if (!finalCompany && extractedCompany && docType !== 'card_payment_auto_create') {
          var similar = _findSimilarExistingName(sheet, extractedCompany);
          if (similar) {
            finalCompany = similar;
            _dlog('MATCH', '유사 업체명 보정: ' + extractedCompany + ' -> ' + finalCompany);
          }
        }

        // [4순위: Gemini 추출 원본]
        if (!finalCompany) finalCompany = extractedCompany;
        if (finalCompany) sheet.getRange(row, 6).setValue(finalCompany);

        // 카드 + H열 → 발행내역 자동생성
        if (docType === 'card_payment_auto_create' && extractedAmount && extractedDate && finalCompany) {
          _createInvoiceRecord(sheet, extractedAmount, extractedDate, finalCompany);
        }

        _colorRow(sheet, row, 2, 8);
        _stamp(sheet, row, 8);
      }
      else if (column === 15) {
        // O열 (증빙): K=금액, L=날짜, M=업체명
        if (extractedAmount) sheet.getRange(row, 11).setValue(extractedAmount);
        if (extractedDate) sheet.getRange(row, 12).setValue(extractedDate);
        if (extractedCompany) sheet.getRange(row, 13).setValue(extractedCompany);
        _colorRow(sheet, row, 11, 15);
        _stamp(sheet, row, 15);

        // 카드 + O열 → 지급내역 자동생성 (입금증이 없으므로)
        if (docType === 'card' && extractedAmount && extractedDate && extractedCompany) {
          sheet.getRange(row, 14).setValue('법인카드');
          _createPaymentRecord(sheet, extractedAmount, extractedDate, extractedCompany);
        }
      }
    } catch (e) {
      _dlog('ERROR', '시트 입력 실패', { msg: String(e) });
    }
  }

  // ===================== [5. 스마트 매칭 헬퍼] =====================

  // F열에 한 가지 회사 이름만 있는지 확인 (단일업체 시트 감지)
  function _getDominantCompanyName(sheet) {
    try {
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) return null;
      var startCheck = Math.max(1, lastRow - 100);
      var vals = sheet.getRange(startCheck, 6, lastRow - startCheck + 1, 1).getValues();

      var counts = {};
      var uniqueNames = [];

      for (var i = 0; i < vals.length; i++) {
        var name = String(vals[i][0]).trim();
        if (name && name !== '지급업체') {
          if (!counts[name]) {
            counts[name] = 0;
            uniqueNames.push(name);
          }
          counts[name]++;
        }
      }
      if (uniqueNames.length === 1 && counts[uniqueNames[0]] >= 1) {
        return uniqueNames[0];
      }
    } catch (e) {}
    return null;
  }

  // F열에 있는 기존 회사명들 중 OCR 결과와 비슷한 게 있는지 찾기
  function _findSimilarExistingName(sheet, targetName) {
    try {
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) return null;
      var startCheck = Math.max(1, lastRow - 50);
      var vals = sheet.getRange(startCheck, 6, lastRow - startCheck + 1, 1).getValues();

      var cleanTarget = _cleanNameForCompare(targetName);
      if (cleanTarget.length < 2) return null;

      for (var i = 0; i < vals.length; i++) {
        var existing = String(vals[i][0]).trim();
        if (!existing || existing === '지급업체') continue;

        var cleanExisting = _cleanNameForCompare(existing);
        if (cleanExisting.indexOf(cleanTarget) >= 0) return existing;
        if (cleanTarget.indexOf(cleanExisting) >= 0) return existing;
      }
    } catch (e) {}
    return null;
  }

  function _cleanNameForCompare(name) {
    return name.replace(/\(주\)|주식회사|\(유\)|유한회사|\s/g, '');
  }

  // 카드 → 발행내역 자동 생성
  function _createInvoiceRecord(sheet, amount, date, company) {
    try {
      var startRow = 7;
      var maxRow = sheet.getMaxRows();
      var values = sheet.getRange(startRow, 11, maxRow - startRow + 1, 1).getValues();
      var targetRow = -1;

      for (var i = 0; i < values.length; i++) {
        if (values[i][0] === "" || values[i][0] === null || typeof values[i][0] === 'undefined') {
          targetRow = startRow + i;
          break;
        }
      }
      if (targetRow === -1) targetRow = sheet.getLastRow() + 1;

      sheet.getRange(targetRow, 11).setValue(amount);
      sheet.getRange(targetRow, 12).setValue(date);
      sheet.getRange(targetRow, 13).setValue(company);
      sheet.getRange(targetRow, 14).setValue("법인카드");

      _colorRow(sheet, targetRow, 11, 15);
      sheet.getRange(targetRow, 15).setNote("🤖 카드 자동생성: " + new Date().toLocaleTimeString());

      _dlog('AUTO_CREATE', '발행내역 자동 생성 완료 (' + company + ')', { row: targetRow });
    } catch (e) {
      _dlog('ERROR', '발행내역 자동생성 실패', e.toString());
    }
  }

  // 카드 O열 → 지급내역 자동 생성
  function _createPaymentRecord(sheet, amount, date, company) {
    try {
      var startRow = 7;
      var maxRow = sheet.getMaxRows();
      var values = sheet.getRange(startRow, 2, maxRow - startRow + 1, 1).getValues();
      var targetRow = -1;

      for (var i = 0; i < values.length; i++) {
        if (values[i][0] === "" || values[i][0] === null || typeof values[i][0] === 'undefined') {
          targetRow = startRow + i;
          break;
        }
      }
      if (targetRow === -1) targetRow = sheet.getLastRow() + 1;

      sheet.getRange(targetRow, 2).setValue(amount);
      sheet.getRange(targetRow, 3).setValue(date);
      sheet.getRange(targetRow, 6).setValue(company);

      _colorRow(sheet, targetRow, 2, 8);
      sheet.getRange(targetRow, 8).setNote("🤖 카드→지급 자동생성: " + new Date().toLocaleTimeString());

      _dlog('AUTO_CREATE', '지급내역 자동 생성 완료 (' + company + ')', { row: targetRow });
    } catch (e) {
      _dlog('ERROR', '지급내역 자동생성 실패', e.toString());
    }
  }

  // ===================== [5-B. 카드사용내역 멀티 행 입력] =====================
  /**
   * 카드사용내역 거래 배열을 시트에 멀티 행으로 입력
   * @param {Sheet} sheet - 대상 시트
   * @param {Array} transactions - [{date, merchant, amount}, ...]
   * @param {number} startRow - 시작 행
   * @param {number} column - 첨부된 열 (8=H열 지급, 15=O열 증빙)
   */
  function _autoFillCardStatement(sheet, transactions, startRow, column) {
    try {
      var count = transactions.length;
      _dlog('CARD_STMT_FILL', '멀티 행 입력 시작', { count: count, startRow: startRow, column: column });

      // 카드사용내역은 입금증이 없으므로, H열/O열 상관없이 항상 지급+발행 양쪽 입력
      // [발행내역 측] 빈 행 찾기 (K열 기준, startRow 또는 7행부터)
      var invoiceStartRow = (column === 15) ? startRow : 7;
      var invoiceRows = _findEmptyRows(sheet, invoiceStartRow, 11, count);

      // [지급내역 측] 빈 행 찾기 (B열 기준, startRow 또는 7행부터)
      var paymentStartRow = (column === 8) ? startRow : 7;
      var paymentRows = _findEmptyRows(sheet, paymentStartRow, 2, count);

      for (var i = 0; i < count; i++) {
        var tx = transactions[i];
        var amount = _toNumberValue(tx.amount);
        var date = tx.date || '';
        var merchant = tx.merchant || '';

        // 발행내역 입력: K=금액, L=날짜, M=업체명, N=법인카드
        var iRow = invoiceRows[i];
        if (amount) sheet.getRange(iRow, 11).setValue(amount);
        if (date) sheet.getRange(iRow, 12).setValue(date);
        if (merchant) sheet.getRange(iRow, 13).setValue(merchant);
        sheet.getRange(iRow, 14).setValue('법인카드');
        _colorRow(sheet, iRow, 11, 15);
        sheet.getRange(iRow, 15).setNote('🤖 카드사용내역 자동생성: ' + new Date().toLocaleTimeString());

        // 지급내역 입력: B=금액, C=날짜, F=업체명
        var pRow = paymentRows[i];
        if (amount) sheet.getRange(pRow, 2).setValue(amount);
        if (date) sheet.getRange(pRow, 3).setValue(date);
        if (merchant) sheet.getRange(pRow, 6).setValue(merchant);
        _colorRow(sheet, pRow, 2, 8);
        _stamp(sheet, pRow, 8);
      }

      _dlog('CARD_STMT_FILL', '지급+발행 양쪽 입력 완료', {
        count: count,
        paymentRows: paymentRows.length,
        invoiceRows: invoiceRows.length
      });

    } catch (e) {
      _dlog('ERROR', '카드사용내역 멀티 행 입력 실패', { msg: String(e), stack: e.stack });
    }
  }

  /**
   * 지정 열 기준으로 빈 행을 count개 찾기
   * 부족하면 시트 끝에 행 추가
   */
  function _findEmptyRows(sheet, searchStartRow, checkCol, count) {
    var rows = [];
    var maxRow = sheet.getMaxRows();
    var lastRow = sheet.getLastRow();
    var endScan = Math.max(maxRow, lastRow + count);

    // 기존 행에서 빈 셀 찾기
    if (maxRow >= searchStartRow) {
      var scanRange = sheet.getRange(searchStartRow, checkCol, maxRow - searchStartRow + 1, 1).getValues();
      for (var i = 0; i < scanRange.length && rows.length < count; i++) {
        if (scanRange[i][0] === '' || scanRange[i][0] === null || typeof scanRange[i][0] === 'undefined') {
          rows.push(searchStartRow + i);
        }
      }
    }

    // 부족하면 시트 끝에 추가
    var remaining = count - rows.length;
    if (remaining > 0) {
      var addStart = maxRow + 1;
      sheet.insertRowsAfter(maxRow, remaining);
      for (var r = 0; r < remaining; r++) {
        rows.push(addStart + r);
      }
      _dlog('CARD_STMT_FILL', remaining + '개 행 추가', { from: addStart });
    }

    return rows;
  }

  // ===================== [6. 유틸리티] =====================
  function _toNumberValue(v) { return v ? Number(String(v).replace(/[^0-9.\-]/g, '')) : null; }
  function _colorRow(s, r, c1, c2) { s.getRange(r, c1, 1, c2 - c1 + 1).setBackground('#e8f5e9'); }
  function _stamp(s, r, c) { s.getRange(r, c).setNote((s.getRange(r, c).getNote() || '') + '\n🤖 Gemini: ' + new Date().toLocaleTimeString()); }
  function _getAttachmentsForCell(sid, a1) { try { var d = PropertiesService.getUserProperties().getProperty(sid + '_' + a1); return d ? JSON.parse(d) : []; } catch (e) { return []; } }
  function _getSheetById_(ss, id) { return ss.getSheets().filter(function(s) { return s.getSheetId() === id; })[0]; }

  // ===================== [7. 로깅] =====================
  function _log(category, message) { _dlog(category, message, ''); }
  function _dlog(c, m, d) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sh = ss.getSheetByName('📋_OCR_로그');
      if (!sh) { sh = ss.insertSheet('📋_OCR_로그'); sh.appendRow(['시간', '구분', '내용', '상세']); }
      sh.appendRow([new Date().toLocaleTimeString(), c, m, JSON.stringify(d || '')]);
    } catch (e) {}
  }

  function _logRawOCRText_(r, f, a, t, s) {
    if (!RAW_OCR_TO_SHEET) return;
    var l = s.getParent().getSheetByName(RAW_OCR_SHEET_NAME);
    if (!l) {
      l = s.getParent().insertSheet(RAW_OCR_SHEET_NAME);
      l.appendRow(['시간', 'runId', 'fileId', '셀', '길이', '내용']);
    }
    l.appendRow([new Date().toLocaleString(), r, f, a, (t || '').length, t]);
  }

  // ===================== [8. 카드사용내역 OCR 메인 함수] =====================
  /**
   * 카드(페이)사용내역 PDF → 멀티 행 OCR 처리
   * @param {string} fileId - Drive 파일 ID
   * @param {number} row - 첨부 셀 행 번호
   * @param {number} column - 첨부 셀 열 번호 (8=H, 15=O)
   */
  function runCardStatementOCR(fileId, row, column) {
    var runId = (Utilities.getUuid() || '').slice(0, 8);
    _dlog('CARD_STMT', '=== 카드사용내역 OCR 시작 ===', { runId: runId, fileId: fileId });

    try {
      if (column !== 8 && column !== 15) {
        _log('ERROR', '허용되지 않은 열(허용: H=8, O=15) → 중단');
        return;
      }

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getActiveSheet();
      var sheetName = sheet.getName();
      var sheetId = sheet.getSheetId();

      var typeLabel = (column === 8) ? '지급내역' : '증빙';
      ss.toast('🤖 카드사용내역 분석 중... (멀티 행)', typeLabel, 10);

      // --- Gemini 호출 (카드사용내역 전용) ---
      var transactions = _extractCardStatementWithGemini(fileId, runId);

      if (!transactions || transactions.length === 0) {
        _log('ERROR', '카드사용내역 추출 결과 없음');
        ss.toast('❌ 실패', '거래 내역을 읽을 수 없습니다', 3);
        return;
      }

      // RAW 로그 기록
      var a1 = sheet.getRange(row, column).getA1Notation();
      _logRawOCRText_(runId, fileId, a1, JSON.stringify(transactions, null, 2), sheet);

      // 시트 참조 재확인
      sheet = ss.getSheetByName(sheetName) || _getSheetById_(ss, sheetId) || sheet;

      _dlog('CARD_STMT', '총 ' + transactions.length + '건 추출 완료, 시트 입력 시작');

      // 멀티 행 입력
      _autoFillCardStatement(sheet, transactions, row, column);

      ss.toast('✅ 카드사용내역 입력 완료!', transactions.length + '건 처리됨', 5);
      _log('SUCCESS', '카드사용내역 ' + transactions.length + '건 처리 완료');

    } catch (e) {
      _dlog('ERROR', '카드사용내역 OCR 시스템 오류', { error: String(e), stack: e.stack });
      SpreadsheetApp.getActiveSpreadsheet().toast('⚠️ 오류 발생', '로그 확인', 5);
    }
  }

  // ===================== [9. UI 메뉴 + 수동 실행] =====================
  function addMenu() {
    try {
      SpreadsheetApp.getUi()
        .createMenu('🤖 OCR 처리')
        .addItem('▶️ 선택 영역 OCR 실행 (Gemini)', 'OCR.manualRun')
        .addSeparator()
        .addItem('📋 카드사용내역 OCR (멀티 행)', 'OCR.manualCardStatementRun')
        .addToUi();
    } catch (e) {
      _log('ERROR', '메뉴 에러: ' + e);
    }
  }

  function manualRun() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var rangeList = sheet.getActiveRangeList();
    if (!rangeList) { SpreadsheetApp.getUi().alert('선택된 셀이 없습니다.'); return; }

    var ranges = rangeList.getRanges();
    var processed = 0;
    var validColumns = [8, 15];

    ranges.forEach(function(range) {
      var startRow = range.getRow();
      var endRow = startRow + range.getNumRows() - 1;
      var startCol = range.getColumn();
      var endCol = startCol + range.getNumColumns() - 1;

      for (var r = startRow; r <= endRow; r++) {
        for (var c = startCol; c <= endCol; c++) {
          if (validColumns.indexOf(c) === -1) continue;
          var cellA1 = sheet.getRange(r, c).getA1Notation();
          var attachments = _getAttachmentsForCell(sheet.getSheetId(), cellA1);
          if (attachments.length > 0 && attachments[0].fileId) {
            SpreadsheetApp.getActiveSpreadsheet().toast('🤖 Gemini 처리 중... (' + cellA1 + ')', '진행 중', -1);
            runAutoOCR(attachments[0].fileId, r, c);
            processed++;
            Utilities.sleep(500);
          }
        }
      }
    });

    if (processed === 0) SpreadsheetApp.getUi().alert('선택된 범위에 처리할 파일(H/O열)이 없습니다.');
    else SpreadsheetApp.getUi().alert('✅ 총 ' + processed + '건 Gemini OCR 처리 완료!');
  }

  /**
   * 카드사용내역 수동 실행 (메뉴에서 호출)
   * 선택한 셀에 첨부된 PDF를 카드사용내역으로 처리
   */
  function manualCardStatementRun() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var range = sheet.getActiveRange();
    if (!range) { SpreadsheetApp.getUi().alert('선택된 셀이 없습니다.'); return; }

    var row = range.getRow();
    var col = range.getColumn();
    var validColumns = [8, 15];

    if (validColumns.indexOf(col) === -1) {
      SpreadsheetApp.getUi().alert('H열(지급내역) 또는 O열(증빙)의 셀을 선택해주세요.');
      return;
    }

    var cellA1 = sheet.getRange(row, col).getA1Notation();
    var attachments = _getAttachmentsForCell(sheet.getSheetId(), cellA1);

    if (!attachments || attachments.length === 0 || !attachments[0].fileId) {
      SpreadsheetApp.getUi().alert('선택한 셀(' + cellA1 + ')에 첨부된 파일이 없습니다.\n\n먼저 카드사용내역 PDF를 첨부해주세요.');
      return;
    }

    // 확인 대화상자
    var ui = SpreadsheetApp.getUi();
    var confirm = ui.alert(
      '📋 카드사용내역 OCR',
      '셀 ' + cellA1 + '에 첨부된 파일을 카드사용내역으로 처리합니다.\n' +
      '\n모든 거래가 개별 행으로 입력됩니다.\n' +
      (col === 8 ? '\n⚡ H열이므로 지급내역 + 발행내역 양쪽에 자동 입력됩니다.' : '\n📝 O열이므로 발행내역(K~N열)에만 입력됩니다.') +
      '\n\n계속하시겠습니까?',
      ui.ButtonSet.YES_NO
    );

    if (confirm !== ui.Button.YES) return;

    SpreadsheetApp.getActiveSpreadsheet().toast('🤖 카드사용내역 분석 중...', '처리 중', -1);
    runCardStatementOCR(attachments[0].fileId, row, col);
  }

  // ===================== [공개 API] =====================
  return {
    runAutoOCR: runAutoOCR,
    runCardStatementOCR: runCardStatementOCR,
    addMenu: addMenu,
    manualRun: manualRun,
    manualCardStatementRun: manualCardStatementRun
  };
})();