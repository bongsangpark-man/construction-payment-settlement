/**
 * AutoOCR_Module.gs [v35 FINAL - Fixed Missing Function]
 * - 1. 세금계산서: v12 정밀 로직
 * - 2. 카드: v29 로직 유지
 * - 3. 입출금(UPGRADE): 날짜/금액 보정 + 헬퍼 함수 누락 수정
 * - 4. 실행: 선택 영역 드래그 일괄 처리
 */
var OCR = (function () {
  'use strict';

  // ===================== [1. 설정 및 필터] =====================
  // ⚠️ 보안 경고: API 키는 가급적 '스크립트 속성'을 이용하세요.
  var API_KEY = 'AIzaSyDvuexZvQijq5r4a1HBXHB65cvGBxaPXZk'; 
  
  var VISION_IMAGES_ENDPOINT = 'https://vision.googleapis.com/v1/images:annotate';
  var VISION_FILES_ENDPOINT = 'https://vision.googleapis.com/v1/files:annotate';
  var RAW_OCR_TO_SHEET = true; 
  var RAW_OCR_SHEET_NAME = '📄_RAW_OCR'; 

  var BUYER_BLACKLIST = ['가현종합건설']; 
  var BUYER_NAME_PROP = 'INVOICE_BUYER_NAME';

  var LABEL_BLACKLIST = [
    '전자세금계산서', '수정전자세금계산서', '세금계산서', '계산서', '영수증', '청구서', '견적서', 
    '승인번호', '작성일자', '수정사유', '이메일', '참조',
    '공급가액', '세액', '합계금액', '청구금액', '현금', '수표', '어음', '외상미수금',
    '공급자', '매출자', '공급받는자', '상호(법인명)', '상호', '법인명', '성명', '대표자',
    '등록', '등록번호', '사업자번호', '주민번호',
    '업태', '종목', '업종', '사업장', '사업장주소', '주소', '연락처', '전화번호', '팩스',
    '품목', '규격', '수량', '단가', '비고', '번호',
    '이금액을', '영수', '청구', '함', '귀하', 
    '도소매', '도매', '소매', '서비스', '제조', '부동산', '임대', '자재', '용품', '건축물', '윤용진', '순숙', '이갑희', '손숙,이강희', '손숙','박봉춘'
  ];

  var CARD_JUNK_WORDS = [
    '대표자', '가맹점', '주소', '전화', 'TEL', '사업자', '등록번호', '승인', '일시', '금액', 
    '부가세', '합계', '매입', '매출', '카드', '신용', '일시불', '할부', '전자', '서명', 
    '가맹점명', '상호', '성명', 'No', 'NO', '영수증', '매출전표', '고객용', '회원용', '문의', '안내',
    '현장', '박봉', '롯(', '수협-', '법인' 
  ];

  var CARD_CONTEXT_BLACKLIST = [
    '서울', '경기', '인천', '부산', '대전', '대구', '광주', '울산', '제주', '강원', '충북', '충남', '전북', '전남', '경북', '경남', 
    '시', '군', '구', '읍', '면', '동', '리', '로', '길', '층', '호', 
    '운영', '농협운영', '기타', '도소매'
  ];

  // ===================== [2. 메인 실행 함수] =====================
  function runAutoOCR(fileId, row, column) {
    var runId = (Utilities.getUuid() || '').slice(0, 8);
    _dlog('OCR', '=== OCR 시작 ===', { runId: runId, fileId: fileId });

    try {
      if (column !== 8 && column !== 15) {
        _log('ERROR', '허용되지 않은 열(허용: H=8, O=15) → 중단');
        return;
      }

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getActiveSheet();
      var sheetId = sheet.getSheetId();
      var sheetName = sheet.getName();
      
      var columnType = (column === 8) ? 'payment' : 'invoice';
      ss.toast('🤖 문서 분석 중...', columnType === 'payment' ? '지급내역' : '증빙', 3);

      var text = _performOCR_toText(fileId);
      if (!text) { 
        _log('ERROR', 'OCR 결과 없음');
        ss.toast('❌ 실패', '내용을 읽을 수 없습니다', 3);
        return; 
      }

      var a1 = sheet.getRange(row, column).getA1Notation();
      _logRawOCRText_(runId, fileId, a1, text, sheet);

      sheet = ss.getSheetByName(sheetName) || _getSheetById_(ss, sheetId) || sheet;

      var detected = _detectDocType(text); 
      var docType = (detected !== 'unknown') ? detected : columnType;
      _dlog('INFO', '문서 유형 감지됨: ' + docType);

      var extracted;
      if (docType === 'invoice') {
         extracted = _extractInvoice(text, runId); 
      } else if (docType === 'card') {
         extracted = _extractCard(text); 
      } else {
         extracted = _extractPayment(text); 
      }

      // [중요] 카드이면서 H열(지급)인 경우 별도 타입 지정
      if (docType === 'card' && column === 8) {
        docType = 'card_payment_auto_create'; 
      }

      _dlog('OCR', '추출 데이터', extracted || {});
      
      if (!extracted || (!extracted.amount && !extracted.company && !extracted.merchant)) { 
        _log('WARNING', '데이터 추출 실패'); 
        ss.toast('⚠️ 확인 필요', '데이터를 찾지 못했습니다', 3);
        return; 
      }

      _autoFillData(sheet, extracted, row, column, docType);
      
      var finalName = extracted.company || extracted.merchant || '완료';
      ss.toast('✅ 입력 완료', finalName, 3);
      _log('SUCCESS', '처리 완료 (' + finalName + ')');

    } catch (e) {
      _dlog('ERROR', '시스템 오류', { error: String(e), stack: e.stack });
      SpreadsheetApp.getActiveSpreadsheet().toast('⚠️ 오류 발생', '로그 확인', 5);
    }
  }

  // ===================== [3. 문서 유형 감지] =====================
  function _detectDocType(text) {
    var t = _collapseVerticalKoreanStacks(_joinSplitLabels(_normSpaces(text || '')));
    var scoreInv = 4 * /계산서/.test(t) + 3 * /공급가액/.test(t) + 2 * /공급자/.test(t) + 2 * /등록번호/.test(t);
    var scorePay = 3 * /입출금|입금증|명세서/.test(t) + 3 * /출금액|잔액/.test(t) + 2 * /의뢰인|적요/.test(t);
    var scoreCard = 4 * /카드|영수증|매출전표/.test(t) + 3 * /할부|가맹점/.test(t) + 2 * /승인번호/.test(t);
    
    if (scoreCard > scoreInv && scoreCard > scorePay) return 'card';
    if (scoreInv >= scorePay && scoreInv >= scoreCard) return 'invoice';
    return 'payment';
  }

  // ===================== [4-A. 입출금 추출] =====================
  function _extractPayment(text) {
    var t = text || '';
    var data = { amount: null, date: null, account: null, bank: null, company: null };

    // [1. 날짜 로직] 문서 내 모든 날짜를 찾아 "가장 과거(오래된)" 날짜 선택
    var dateRegex = /([0-9]{4})[\.\-\/년]\s*([0-9]{1,2})[\.\-\/월]\s*([0-9]{1,2})/g;
    var datesFound = [];
    var match;
    while ((match = dateRegex.exec(t)) !== null) {
        datesFound.push(match[1] + '-' + _pad2(match[2]) + '-' + _pad2(match[3]));
    }
    if (datesFound.length > 0) {
        datesFound.sort(); 
        data.date = datesFound[0]; 
    } else {
        var mShortDate = t.match(/([0-9]{1,2})[\.\-\/월]\s*([0-9]{1,2})[\.\-\/일]/);
        if (mShortDate) {
              var year = new Date().getFullYear();
              data.date = year + '-' + _pad2(mShortDate[1]) + '-' + _pad2(mShortDate[2]);
        }
    }

    // [2. 금액 로직] 줄바꿈이나 노이즈로 인해 금액이 밀려난 경우 대비
    var lines = t.split(/\r?\n/);
    var foundAmount = false;

    for (var i = 0; i < lines.length; i++) {
        var line = lines[i].trim();
        if (!line) continue;
        if (/잔액/.test(line)) continue; 

        if (/출금|입금|이체|송금|거래금액/.test(line)) {
            // Case A: 같은 줄에 금액이 온전하게 있는 경우
            var mVal = line.match(/([0-9,]+)\s*원?/);
            if (mVal && mVal[1] !== '0' && mVal[1] !== '00') {
                data.amount = mVal[1].replace(/[^0-9]/g, '');
                foundAmount = true;
                break;
            } 
            // Case B: 금액이 다음 줄로 밀려난 경우
            else if (lines[i+1]) {
                var nextLine = lines[i+1].trim();
                var mNextVal = nextLine.match(/^([0-9,]+)\s*원?$/); 
                if (mNextVal && mNextVal[1] !== '0' && mNextVal[1] !== '00') {
                    data.amount = mNextVal[1].replace(/[^0-9]/g, '');
                    foundAmount = true;
                    break;
                }
            }
        }
    }

    // Case C: 키워드 근처에서 못 찾았다면, 전체 텍스트에서 가장 큰 금액 추정
    if (!foundAmount) {
        var cleanText = t.replace(/잔액\s*[:]?\s*[0-9,]+/g, ''); 
        var rawAmt = cleanText.match(/([0-9,]{3,})\s*원/);
        if (rawAmt) {
             data.amount = rawAmt[1].replace(/[^0-9]/g, '');
        }
    }

    // [3. 업체명/계좌 등 나머지 추출]
    for (var i = 0; i < lines.length; i++) {
        var line = lines[i].trim();
        if (/적요|받는분|받는사람|내용|상호|가맹점/.test(line)) {
            var sameLine = line.replace(/적요|받는분|받는사람|내용|상호|가맹점|[:]/g, '').trim();
            if (_isValidCompanyString(sameLine)) { data.company = sameLine; break; }
            if (lines[i+1]) {
                var nextLine = lines[i+1].trim();
                if (_isValidCompanyString(nextLine)) { data.company = nextLine; break; }
            }
        }
    }
    if (!data.company) {
        lines.forEach(function(l) {
            if (l.indexOf('출금') > -1 && !/\d/.test(l)) { 
                 var cand = l.replace('출금', '').trim();
                 if (_isValidCompanyString(cand)) data.company = cand;
            }
        });
    }

    var mBank = t.match(/(?:은행|계좌)[\s\S]{0,10}?([가-힣]+)\s?[\/]\s?([0-9\-]+)/);
    if (mBank) {
       data.bank = mBank[1];
       data.account = mBank[2];
    } else {
       var mAccOnly = t.match(/(\d{3,}-\d{2,}-\d{3,})/);
       if(mAccOnly) data.account = mAccOnly[1];
    }
    return data;
  }

  // ⚠️ [중요] 여기가 누락되었던 부분입니다! (입출금 내역 검증 함수)
  function _isValidCompanyString(str) {
      if (!str) return false;
      if (str.endsWith('원')) return false;
      if (/^[0-9,]+$/.test(str)) return false;
      if (str.length < 2) return false;
      if (/잔액|출금액|입금액|수수료|거래수단/.test(str)) return false;
      return true;
  }

  // ===================== [4-B. 카드 추출] =====================
  function _extractCard(text) {
    var t = text || '';
    var data = { amount: null, date: null, merchant: null, cardMasked: null };
    
    var rawLines = t.split(/\r?\n/);
    var lines = rawLines.map(function(s){ return s.trim(); }).filter(function(s){ return s.length > 0; });

    var mTotalLabel = t.match(/(?:합계|결제금액|승인금액|청구금액)\s*[:]?\s*([0-9,]+)/);
    if (mTotalLabel) {
        data.amount = mTotalLabel[1].replace(/[^0-9]/g, '');
    } else {
        for (var i = 0; i < lines.length; i++) {
            var line = lines[i];
            if (/공급가액|부가세|세액|과세/.test(line)) continue; 
            var standalone = line.match(/^([0-9,]+)\s*원$/);
            if (standalone) {
                 data.amount = standalone[1].replace(/[^0-9]/g, '');
                 break; 
            }
            var mRawAmt = line.match(/([0-9,]{4,})\s*원/); 
            if (mRawAmt) {
                 var val = mRawAmt[1].replace(/[^0-9]/g, '');
                 if (!data.amount) data.amount = val;
            }
        }
    }

    var mDate = t.match(/(?:승인|거래|사용)일시?\s*[:]?\s*([0-9]{4}[-./]\d{2}[-./]\d{2})/);
    if (mDate) {
        data.date = mDate[1].replace(/[./]/g, '-');
    } else {
        var mRawDate = t.match(/(20\d{2}[-./]\d{2}[-./]\d{2})/);
        if (mRawDate) data.date = mRawDate[1].replace(/[./]/g, '-');
    }

    var mCard = t.match(/(\d{3,4}[-–]\d{3,4}[-–][*\-xX]{4,}[-–]\d{3,4})/);
    if (mCard) data.cardMasked = mCard[1].replace(/[–]/g, '-');

    if (lines.length > 2) {
        var target = lines[2]; 
        var clean = _cleanMerchantString(target);
        if (_isValidMerchant(clean)) data.merchant = clean;
    }
    if (!data.merchant && lines.length > 1) {
        var target = lines[1];
        var clean = _cleanMerchantString(target);
        if (_isValidMerchant(clean)) data.merchant = clean;
    }
    if (!data.merchant) {
        for (var i = 0; i < Math.min(lines.length, 8); i++) {
             var raw = lines[i];
             if (/(주)|주식회사/.test(raw)) {
                 var clean = _cleanMerchantString(raw);
                 if (_isValidMerchant(clean) && !/가현종합/.test(clean)) { data.merchant = clean; break; }
             }
        }
    }
    if (!data.merchant) {
        for (var i = 0; i < lines.length; i++) {
            var line = lines[i];
            if (/카드|[\*]{4}/.test(line) || /\d{4}-\d{4}/.test(line)) {
                if (lines[i+1]) {
                    var cand1 = _cleanMerchantString(lines[i+1]);
                    if (_isValidMerchant_Loose(cand1)) { data.merchant = cand1; break; }
                }
                if (lines[i+2]) {
                      var cand2 = _cleanMerchantString(lines[i+2]);
                      if (_isValidMerchant_Loose(cand2)) { data.merchant = cand2; break; }
                }
            }
        }
    }
    if (!data.merchant) {
        for (var k = 0; k < Math.min(lines.length, 6); k++) {
            var raw = lines[k];
            var clean = _cleanMerchantString(raw);
            if (_isValidMerchant(clean)) { data.merchant = clean; break; }
        }
    }
    return data;
  }

  function _cleanMerchantString(str) {
      return str.split(/\s(Tel|TEL|전화|대표|0\d{1,2}-)/)[0].trim();
  }
  function _isValidMerchant_Loose(str) {
      if (!str) return false;
      var n = str.replace(/\s/g, '');
      if (n.length < 2) return false;
      if (/^[0-9\-:.\/]+$/.test(n)) return false; 
      if (/[0-9]/.test(n) && n.endsWith('원')) return false; 
      return !CARD_JUNK_WORDS.some(function(bad) { return n.indexOf(bad) >= 0; });
  }
  function _isValidMerchant(str) {
      if (!_isValidMerchant_Loose(str)) return false;
      var n = str.replace(/\s/g, '');
      var addressPattern = /(서울|경기|인천|부산|대구|광주|대전|울산|강원|충북|충남|전북|전남|경북|경남|제주).*(시|군|구|읍|면|동|로|길|번지)/;
      if (addressPattern.test(n)) return false;
      if (/답십리|상봉동|신내동/.test(n)) return false; 
      if (CARD_CONTEXT_BLACKLIST.some(function(bad) { return n.indexOf(bad) >= 0; })) return false;
      return true;
  }

  // ===================== [4-C. 세금계산서 추출] =====================
  function _extractInvoice(text, runId) {
    var t = _collapseVerticalKoreanStacks(_joinSplitLabels(_normSpaces(text || '')));
    var data = { amount: null, date: null, company: null };
    data.amount = _findAmountNear(t.split('\n'), /합\s*계\s*금\s*액|청\s*구\s*금\s*액/i, 10);
    var m = t.match(/작성일자\s*(\d{4})[.\-/년]\s*(\d{1,2})[.\-/월]\s*(\d{1,2})/) || t.match(/(20\d{2})[.\-/년]\s*(\d{1,2})[.\-/월]\s*(\d{1,2})/);
    if (m) data.date = m[1].slice(2) + '-' + _pad2(m[2]) + '-' + _pad2(m[3]);
    var sup = _extractSupplierBlockName(t);
    data.company = sup.value || _selectCompanyWithTrace(t).company;
    if (data.company) {
       var cleaned = _stripLabels(data.company);
       if (!cleaned || _evaluateSupplierCandidate(cleaned).reasons.length > 0) data.company = '';
       else data.company = cleaned;
    }
    return data;
  }
  function _evaluateSupplierCandidate(raw) {
    var cleaned = _stripLabels(raw).replace(/\s+/g, ' ').trim();
    var reasons = [];
    var strongCompanyRegex = /(주식회사|\(주\)|건설|기공|산업|기업|엔지니어링|테크|시스템|유진|이앤제이|신영|구청|시청|상사|유통|공단|협회|조합|디지털|사무기|OA|모바일)/;
    var isStrong = strongCompanyRegex.test(cleaned);
    if (/^(건설업|공사업|제조업|서비스업|도소매업|운수업|임대업|부동산|부동산업|도소매|건재|도.소매|도매|소매)$/i.test(cleaned)) reasons.push('forbidden-strict');
    else if (/공사업|주거용|비주거용|배관|난방|세금계산서|영수증|청구서|자재|용품|안전진단|건축물|시설/.test(cleaned)) {
        if (!isStrong) reasons.push('forbidden-contains-weak');
    }
    if (reasons.length === 0) {
        if (!cleaned) reasons.push('empty');
        else {
          if (!isStrong && _isLabelLike(cleaned)) reasons.push('label-like');
          if (_isBlacklistedBuyer(cleaned)) reasons.push('blacklisted');
          if (!/[가-힣]/.test(cleaned)) reasons.push('no-hangul');
          if (cleaned.length < 2 || cleaned.length > 30) reasons.push('len-out');
          if (/(대표자|성명|사업자등록번호|공급가액|세액|합계금액)/.test(cleaned)) reasons.push('forbidden-terms');
          if (/\d/.test(cleaned) && /\s/.test(cleaned)) reasons.push('digit-phrase');
        }
    }
    return { raw: raw, cleaned: cleaned, accepted: reasons.length === 0, reasons: reasons };
  }
  function _looksLikeCompanyName(name) {
    if (!name) return false;
    if (/^(건설업|공사업|제조업|서비스업|도소매업|도소매|부동산업|부동산)$/.test(name)) return false;
    if (/(\s및\s|청구|영수|이금액을|세금계산서|계산서|영수증)/.test(name)) return false; 
    var hasCorp = /(주식회사|\(주\)|상사|기업|건설)/.test(name);
    if (!hasCorp && /안전진단|용품|자재|건축물/.test(name)) return false;
    if (_isLabelLike(name) || _isBlacklistedBuyer(name)) return false;
    if (/\d/.test(name) && /\s/.test(name)) return false;
    var keywords = /(주|㈜|회사|공사|건설|산업|상사|기공|기업|엔지니어링|ENG|테크|시스템|이앤제이|유진|신영|구청|시청|유통|공단|디지털|사무기|OA|모바일)/i;
    if (keywords.test(name)) return true;
    if (/\(.*\)/.test(name)) return true;
    var compact = _stripSpaces(name);
    if (compact.length >= 3 && !/\s/.test(name)) return true;
    return false;
  }
  function _extractSupplierBlockName(text) {
    var normalized = _collapseVerticalKoreanStacks(_joinSplitLabels(_normSpaces(text || '')));
    var matched = normalized.match(/공급.?자([\s\S]{0,3000}?)(공급.?받는.?자|$)/i);
    var lines = matched ? matched[1].split(/\r?\n/) : normalized.split(/\r?\n/);
    var fallback = '', lastLabel = '';
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim(); if(!line) continue;
      if (/공급받는자|매출자/i.test(line)) { lastLabel = ''; continue; }
      var inline = line.match(/(상호\(법인명\)|상호|법인명)\s*[:：]?\s*(.+)$/i);
      if (inline) { var res = _evaluateSupplierCandidate(inline[2]); if(res.accepted) return {value:res.cleaned}; lastLabel='company'; continue; }
      if (/(상호\(법인명\)|상호|법인명)\s*[:：]?$/.test(line)) { lastLabel = 'company'; continue; }
      if (/(성명|대표자)/.test(line)) { lastLabel = ''; continue; }
      var ev = _evaluateSupplierCandidate(line);
      if (ev.reasons.indexOf('blacklisted') >= 0) { lastLabel = ''; continue; }
      if (lastLabel === 'company') { if(ev.accepted) return {value:ev.cleaned}; if(!fallback&&ev.accepted&&_looksLikeCompanyName(ev.cleaned)) fallback=ev.cleaned; }
      else if(!fallback&&ev.accepted&&_looksLikeCompanyName(ev.cleaned)) fallback=ev.cleaned;
    }
    if (!fallback) {
        for (var i = 0; i < lines.length; i++) {
            var line = lines[i].trim();
            if (/(주)|주식회사/.test(line)) {
                var ev = _evaluateSupplierCandidate(line);
                if (ev.accepted) return { value: ev.cleaned };
            }
        }
    }
    return { value: fallback };
  }
  function _selectCompanyWithTrace(t, runId) {
    var chosen = _matchSupplierName(t);
    if (!chosen || _isBlacklistedBuyer(chosen)) { var cands = _gatherSupplierCandidates(t); if (cands.length > 0) chosen = cands[0]; }
    return { company: chosen ? _stripLabels(chosen) : '' };
  }

  // ===================== [5. 시트 입력 (Mapping + 스마트매칭 + 1+1자동생성)] =====================
  function _autoFillData(sheet, data, row, column, docType) {
    try {
      var extractedDate = data.date;
      var extractedAmount = _toNumberValue(data.amount);
      var extractedCompany = data.company || data.merchant || '';
      var extractedAccount = '';

      if (data.cardMasked) extractedAccount = data.cardMasked;
      else if (data.bank || data.account) extractedAccount = (data.bank || '') + ' ' + (data.account || '');

      if (column === 8) {
          if (extractedAmount) sheet.getRange(row, 2).setValue(extractedAmount); 
          if (extractedDate) sheet.getRange(row, 3).setValue(extractedDate); 
          if (extractedAccount.trim()) sheet.getRange(row, 4).setValue(extractedAccount.trim()); 

          var finalCompany = '';
          
          // [1순위: 금액 매칭]
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
          
          // [4순위: OCR 원본]
          if (!finalCompany) finalCompany = extractedCompany;
          if (finalCompany) sheet.getRange(row, 6).setValue(finalCompany); 

          if (docType === 'card_payment_auto_create' && extractedAmount && extractedDate && finalCompany) {
              _createInvoiceRecord(sheet, extractedAmount, extractedDate, finalCompany);
          }

          _colorRow(sheet, row, 2, 8); 
          _stamp(sheet, row, 8);
      } 
      else if (column === 15) {
          if (extractedAmount) sheet.getRange(row, 11).setValue(extractedAmount); 
          if (extractedDate) sheet.getRange(row, 12).setValue(extractedDate); 
          if (extractedCompany) sheet.getRange(row, 13).setValue(extractedCompany); 
          _colorRow(sheet, row, 11, 15); _stamp(sheet, row, 15);
      }
    } catch (e) {
       _dlog('ERROR', '시트 입력 실패', {msg: String(e)});
    }
  }

  // [Helper] F열(지급업체)에 한 가지 회사 이름만 적혀있는지 확인 (단일업체 시트 감지)
  function _getDominantCompanyName(sheet) {
      try {
          var lastRow = sheet.getLastRow();
          if (lastRow < 2) return null;
          var startCheck = Math.max(1, lastRow - 50);
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
      } catch(e) {}
      return null;
  }

  // [Helper] F열에 있는 기존 회사명들 중에서 OCR 결과와 비슷한 게 있는지 찾기
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
      } catch(e) {}
      return null;
  }

  function _cleanNameForCompare(name) {
      return name.replace(/\(주\)|주식회사|\(유\)|유한회사|\s/g, '');
  }

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
          
          _dlog('AUTO_CREATE', '발행내역 자동 생성 완료 (' + company + ')', {row: targetRow});
      } catch(e) {
          _dlog('ERROR', '발행내역 자동생성 실패', e.toString());
      }
  }

  // ===================== [6. 유틸리티] =====================
  function _getApiKey() { if(API_KEY) return API_KEY; throw new Error('API Key Missing'); }
  function _performOCR_toText(fileId) {
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    var mime = (blob.getContentType() || '').toLowerCase();
    var key = _getApiKey();
    if (mime === 'application/pdf') {
      var payloadPDF = { requests: [{ inputConfig: { mimeType: 'application/pdf', content: Utilities.base64Encode(blob.getBytes()) }, features: [{ type: 'DOCUMENT_TEXT_DETECTION', model: 'builtin/latest' }], imageContext: { languageHints: ['ko', 'en'] }, pages: [] }] };
      try {
        var resPDF = UrlFetchApp.fetch(VISION_FILES_ENDPOINT + '?key=' + key, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payloadPDF), muteHttpExceptions: true });
        var dataPDF = JSON.parse(resPDF.getContentText() || '{}');
        var full = '';
        (dataPDF.responses || []).forEach(function(r) {
          if (r.fullTextAnnotation) full += r.fullTextAnnotation.text + '\n';
          if (r.responses) { r.responses.forEach(function(sub) { if(sub.fullTextAnnotation) full += sub.fullTextAnnotation.text + '\n'; }); }
        });
        return _collapseVerticalKoreanStacks(_normSpaces(_joinSplitLabels(full)));
      } catch(e) { _dlog('NET_ERROR', String(e)); return ''; }
    } else {
        var payloadImg = { requests: [{ image: { content: Utilities.base64Encode(blob.getBytes()) }, features: [{ type: 'DOCUMENT_TEXT_DETECTION', model: 'builtin/latest' }], imageContext: { languageHints: ['ko', 'en'] } }] };
        try {
            var resImg = UrlFetchApp.fetch(VISION_IMAGES_ENDPOINT + '?key=' + key, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payloadImg), muteHttpExceptions: true });
            var dataImg = JSON.parse(resImg.getContentText() || '{}');
            var text = (dataImg.responses && dataImg.responses[0] && dataImg.responses[0].fullTextAnnotation && dataImg.responses[0].fullTextAnnotation.text) || '';
            return _collapseVerticalKoreanStacks(_normSpaces(_joinSplitLabels(text)));
        } catch(e) { _dlog('NET_ERROR', String(e)); return ''; }
    }
  }

  function _cleanEntityName(n) { return String(n).replace(/\(법인명\)/g,'').replace(/㈜/g,'(주)').replace(/주식회사/g,'(주)').replace(/\s+/g,' ').trim(); }
  function _joinSplitLabels(s) { return s.replace(/공급\s*자/g,'공급자').replace(/공급\s*받는\s*자/g,'공급받는자').replace(/상\s*호/g,'상호').replace(/합\s*계\s*금\s*액/g,'합계금액'); }
  function _collapseVerticalKoreanStacks(s) { return s.replace(/(^|\n)((?:[가-힣]\s*(?:\n|$)){2,8})/g, function(_,p,c){ return p+c.replace(/\s+/g,'')+'\n'; }); }
  function _normSpaces(s) { return (s||'').replace(/[\u00A0\u202F]/g, ' '); }
  function _pad2(v) { return ('0'+v).slice(-2); }
  function _toNumberValue(v) { return v ? Number(String(v).replace(/[^0-9.\-]/g,'')) : null; }
  function _isLabelLike(s) { var v = _stripSpaces(s); if (LABEL_BLACKLIST.some(function(l){ return v.indexOf(_stripSpaces(l))>=0; })) return true; if (/(업태|종목|업종|시|도|군|구|읍|면|동|리|로|길|호|층|원)$/.test(v)) return true; return false; }
  function _stripSpaces(s) { return (s||'').replace(/\s+/g,''); }
  function _stripLabels(s) { var t = s; LABEL_BLACKLIST.forEach(function(l){ t = t.replace(new RegExp(l,'gi'),' '); }); return _cleanEntityName(t); }
  function _matchSupplierName(t) { return _guessNameAroundBizNo(t); }
  function _gatherSupplierCandidates(t) { var c=[]; t.split('\n').forEach(function(l){ var v=_stripLabels(l); if(_looksLikeCompanyName(v)&&_evaluateSupplierCandidate(v).accepted) c.push(v); }); return c; }
  function _guessNameAroundBizNo(t) { var l=t.split('\n'); var i=l.findIndex(function(ln){ return /\d{3}[-\s]*\d{2}[-\s]*\d{5}/.test(ln); }); if(i<0)return''; for(var k=Math.max(0,i-3);k<Math.min(l.length,i+4);k++){ var v=_stripLabels(l[k]); if(_evaluateSupplierCandidate(v).accepted&&_looksLikeCompanyName(v)) return v; } return ''; }
  function _isBlacklistedBuyer(n) { var t=_normalizeEntityName(n); var l=BUYER_BLACKLIST.concat([PropertiesService.getScriptProperties().getProperty(BUYER_NAME_PROP)||'']); return l.some(function(b){ return b && t.indexOf(_normalizeEntityName(b))>=0; }); }
  function _normalizeEntityName(s) { return s.replace(/[^가-힣0-9A-Za-z]/g,''); }
  function _findAmountNear(lines, re, win) { var idx = lines.findIndex(function(l){ return re.test(l); }); if(idx<0)return null; var txt = lines.slice(idx, idx+win).join(' '); var m = txt.match(/([0-9,]+)\s*원?/); return m ? m[1] : null; }
  function _getAttachmentsForCell(sid, a1) { try{var d=PropertiesService.getUserProperties().getProperty(sid+'_'+a1); return d?JSON.parse(d):[];}catch(e){return[];} }
  function _colorRow(s,r,c1,c2) { s.getRange(r,c1,1,c2-c1+1).setBackground('#e8f5e9'); }
  function _stamp(s,r,c) { s.getRange(r,c).setNote((s.getRange(r,c).getNote()||'')+'\n🤖: '+new Date().toLocaleTimeString()); }
  function _logRawOCRText_(r,f,a,t,s) { if(!RAW_OCR_TO_SHEET)return; var l=s.getParent().getSheetByName(RAW_OCR_SHEET_NAME); if(!l){l=s.getParent().insertSheet(RAW_OCR_SHEET_NAME);l.appendRow(['시간','runId','fileId','셀','길이','내용']);} l.appendRow([new Date().toLocaleString(),r,f,a,(t||'').length,t]); }
  function _getSheetById_(ss, id) { return ss.getSheets().filter(function(s){return s.getSheetId()===id;})[0]; }
  
  function _log(category, message) { _dlog(category, message, ''); }
  function _dlog(c,m,d) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sh = ss.getSheetByName('📋_OCR_로그');
      if(!sh) { sh=ss.insertSheet('📋_OCR_로그'); sh.appendRow(['시간','구분','내용','상세']); }
      sh.appendRow([new Date().toLocaleTimeString(), c, m, JSON.stringify(d||'')]);
    } catch(e) {}
  }

  function addMenu() { try { SpreadsheetApp.getUi().createMenu('🤖 OCR 처리').addItem('▶️ 선택 영역 OCR 실행', 'OCR.manualRun').addToUi(); } catch (e) { _log('ERROR', '메뉴 에러: ' + e); } }
  function manualRun() {
    var sheet = SpreadsheetApp.getActiveSheet(); var rangeList = sheet.getActiveRangeList();
    if (!rangeList) { SpreadsheetApp.getUi().alert('선택된 셀이 없습니다.'); return; }
    var ranges = rangeList.getRanges(); var processed = 0; var validColumns = [8, 15];
    ranges.forEach(function(range) {
      var startRow = range.getRow(); var endRow = startRow + range.getNumRows() - 1;
      var startCol = range.getColumn(); var endCol = startCol + range.getNumColumns() - 1;
      for (var r = startRow; r <= endRow; r++) {
        for (var c = startCol; c <= endCol; c++) {
          if (validColumns.indexOf(c) === -1) continue;
          var cellA1 = sheet.getRange(r, c).getA1Notation();
          var attachments = _getAttachmentsForCell(sheet.getSheetId(), cellA1);
          if (attachments.length > 0 && attachments[0].fileId) {
            SpreadsheetApp.getActiveSpreadsheet().toast('OCR 처리 중... (' + cellA1 + ')', '진행 중', -1);
            runAutoOCR(attachments[0].fileId, r, c); processed++; Utilities.sleep(500);
          }
        }
      }
    });
    if (processed === 0) SpreadsheetApp.getUi().alert('선택된 범위에 처리할 파일(H/O열)이 없습니다.');
    else SpreadsheetApp.getUi().alert('✅ 총 ' + processed + '건 처리 완료!');
  }

  return { runAutoOCR: runAutoOCR, addMenu: addMenu, manualRun: manualRun };
})();