/* ============================================
   Google Apps Script - 주문 관리 시스템
   
   사용법:
   1. Google Sheets 새 파일 만들기
   2. 확장 프로그램 → Apps Script 클릭
   3. 이 코드 전체 복사 → 붙여넣기
   4. 배포 → 새 배포 → 웹 앱 → 모든 사용자 접근 허용
   5. 생성된 URL을 order-form.html의 CONFIG.scriptUrl에 붙여넣기
   ============================================ */

/* ============================================
   ⚙️ 설정
   ============================================ */
const SHEET_NAME = '주문목록';        // 주문 데이터
const SETTINGS_SHEET = '설정';        // 가게/계좌 설정
const PRODUCTS_SHEET = '상품';        // 상품 목록
const ADMIN_PASSWORD = 'admin1234';   // 관리자 페이지 비밀번호 (변경 필수!)

/* ============================================
   시트 자동 생성 (3개 시트)
   ============================================ */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupOrdersSheet(ss);
  setupSettingsSheet(ss);
  setupProductsSheet(ss);
  SpreadsheetApp.getUi().alert(
    '✅ 시트 설정 완료!\n\n' +
    '• 주문목록 - 손님 주문이 자동으로 들어옴\n' +
    '• 설정 - 가게이름·계좌번호 등 (수정 시 폼에 즉시 반영)\n' +
    '• 상품 - 판매중 체크박스로 노출 제어'
  );
}

function setupOrdersSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  const headers = [
    '주문번호', '주문일시', '상품명', '단가', '수량', '총금액',
    '주문자명', '연락처', '입금자명', '보내는분', '보내는분연락처', '입금확인',
    '우편번호', '주소', '배송메모', '송장번호', '발송완료'
  ];

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#E8F5E9');

  const lastRow = Math.max(sheet.getLastRow(), 100);

  // 입금확인 드롭다운 (L열)
  const confirmRange = sheet.getRange(2, 12, lastRow - 1, 1);
  const confirmRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['미확인', '확인완료', '환불']).build();
  confirmRange.setDataValidation(confirmRule);

  // 발송완료 드롭다운 (Q열)
  const shipRange = sheet.getRange(2, 17, lastRow - 1, 1);
  const shipRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['미발송', '발송완료', '반품']).build();
  shipRange.setDataValidation(shipRule);

  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 160);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 130);
  sheet.setColumnWidth(12, 300);

  const dataRange = sheet.getRange(2, 1, lastRow - 1, headers.length);
  const greenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$L2="확인완료"').setBackground('#E8F5E9')
    .setRanges([dataRange]).build();
  const yellowRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$L2="미확인"').setBackground('#FFF9C4')
    .setRanges([dataRange]).build();
  sheet.setConditionalFormatRules([greenRule, yellowRule]);
}

function setupSettingsSheet(ss) {
  let sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (sheet) return;  // 이미 있으면 사용자 데이터 보존

  sheet = ss.insertSheet(SETTINGS_SHEET);
  const data = [
    ['설정 항목', '값'],
    ['가게 이름', '배플루언서'],
    ['헤더 이모지', '🍓'],
    ['은행 이름', '농협'],
    ['계좌번호', '123-4567-8901-23'],
    ['예금주', '현농프레쉬'],
    ['최대 주문 수량', 10],
    ['발송지 주소', '광주광역시 OO구 OO로 123']
  ];
  sheet.getRange(1, 1, data.length, 2).setValues(data);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#E8F5E9');
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 280);
  sheet.getRange('A1').setNote(
    '값 열(B)을 변경하면 주문폼에 즉시 반영됩니다.\n행 순서 변경/삭제 금지!'
  );
}

function setupProductsSheet(ss) {
  let sheet = ss.getSheetByName(PRODUCTS_SHEET);
  if (sheet) return;

  sheet = ss.insertSheet(PRODUCTS_SHEET);
  const data = [
    ['상품명', '가격', '판매중'],
    ['설향 딸기 2kg (250g×8팩)', 35000, true],
    ['설향 딸기 2kg (500g×4팩)', 33000, true],
    ['설향 딸기 1.32kg (330g×4팩)', 25000, true]
  ];
  sheet.getRange(1, 1, data.length, 3).setValues(data);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#E8F5E9');
  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 80);

  // 판매중 체크박스 (100행까지)
  const checkRange = sheet.getRange(2, 3, 100, 1);
  checkRange.setDataValidation(
    SpreadsheetApp.newDataValidation().requireCheckbox().build()
  );

  sheet.getRange('A1').setNote(
    '상품 추가: 행 추가 + 판매중 체크\n중단: 판매중 체크 해제 (행 삭제 X — 이력 보존)'
  );
}

/* ============================================
   POST 요청 처리

   public:  newOrder        — 손님 주문폼 제출
   admin:   updateConfig    — 관리자 페이지에서 설정/상품 저장
   ============================================ */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action || 'newOrder';

    // 공개 액션 (비번 X)
    if (action === 'newOrder') {
      return handleNewOrder(data);
    }

    // 관리자 액션 (비번 필요)
    if (data.pw !== ADMIN_PASSWORD) {
      return jsonResponse({ success: false, error: 'Unauthorized' });
    }

    switch (action) {
      case 'updateConfig':
        return updateConfig(data);
      default:
        return jsonResponse({ success: false, error: 'Unknown action: ' + action });
    }
  } catch (error) {
    return jsonResponse({ success: false, error: error.message });
  }
}

/* ============================================
   새 주문 접수 (public)
   ============================================ */
function handleNewOrder(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    setupOrdersSheet(ss);
    sheet = ss.getSheetByName(SHEET_NAME);
  }

  sheet.appendRow([
    data.orderId || '',
    data.orderDate || new Date().toLocaleString('ko-KR'),
    data.product || '',
    data.unitPrice || 0,
    data.quantity || 1,
    data.totalPrice || 0,
    data.name || '',
    data.phone || '',
    data.depositor || data.name || '',
    data.senderName || data.name || '',
    data.senderPhone || data.phone || '',
    '미확인',
    data.zipcode || '',
    data.address || '',
    data.memo || '',
    '',
    '미발송'
  ]);

  return jsonResponse({ success: true, orderId: data.orderId });
}

/* ============================================
   GET 요청 처리 (공개 + 관리자 API)
   ============================================ */
function doGet(e) {
  const action = e.parameter.action || 'getConfig';

  try {
    // 공개 API (비번 불필요) — 주문폼이 호출
    if (action === 'getConfig') {
      return getConfig();
    }

    // 관리자 API (비번 필요)
    const password = e.parameter.pw || '';
    if (password !== ADMIN_PASSWORD) {
      return jsonResponse({ error: 'Unauthorized' });
    }

    switch (action) {
      case 'getAdminConfig':
        return getAdminConfig();
      case 'getOrderStats':
        return getOrderStats();
      default:
        return jsonResponse({ error: 'Unknown action' });
    }
  } catch (error) {
    return jsonResponse({ error: error.message });
  }
}

/* ============================================
   설정·상품 헬퍼 (시트에서 읽기)
   ============================================ */
function readSettings(ss) {
  const settings = {};
  const sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (sheet && sheet.getLastRow() >= 2) {
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    values.forEach(row => {
      if (row[0]) settings[String(row[0]).trim()] = row[1];
    });
  }
  return settings;
}

function readProducts(ss, onlyActive = true) {
  const products = [];
  const sheet = ss.getSheetByName(PRODUCTS_SHEET);
  if (sheet && sheet.getLastRow() >= 2) {
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    values.forEach(row => {
      if (!row[0]) return;
      if (onlyActive && row[2] !== true) return;
      products.push({
        name: String(row[0]),
        price: Number(row[1]) || 0,
        active: row[2] === true
      });
    });
  }
  return products;
}

/* ============================================
   설정·상품 정보 조회 (주문폼이 페이지 로드 시 호출, 공개)
   ============================================ */
function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = readSettings(ss);
  const products = readProducts(ss, true).map(p => ({ name: p.name, price: p.price }));

  return jsonResponse({
    shopName: settings['가게 이름'] || '주문폼',
    headerEmoji: settings['헤더 이모지'] || '🛒',
    bank: {
      name: settings['은행 이름'] || '',
      account: String(settings['계좌번호'] || ''),
      holder: settings['예금주'] || ''
    },
    maxQty: Number(settings['최대 주문 수량']) || 10,
    products: products
  });
}

/* ============================================
   관리자: 설정·상품 raw 조회 (비활성 상품 포함)
   ============================================ */
function getAdminConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return jsonResponse({
    settings: readSettings(ss),
    products: readProducts(ss, false)
  });
}

/* ============================================
   관리자: 설정·상품 일괄 수정 (admin → POST)
   data = { settings: {키:값,...}, products: [{name,price,active},...] }
   ============================================ */
function updateConfig(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 설정 업데이트 (기존 행의 B열만 갱신, 키 매칭)
  if (data.settings && typeof data.settings === 'object') {
    const sheet = ss.getSheetByName(SETTINGS_SHEET);
    if (!sheet) return jsonResponse({ success: false, error: '설정 시트 없음' });

    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      values.forEach((row, i) => {
        const key = String(row[0]).trim();
        if (Object.prototype.hasOwnProperty.call(data.settings, key)) {
          sheet.getRange(i + 2, 2).setValue(data.settings[key]);
        }
      });
    }
  }

  // 상품 전체 교체 (clear + write)
  if (Array.isArray(data.products)) {
    const sheet = ss.getSheetByName(PRODUCTS_SHEET);
    if (!sheet) return jsonResponse({ success: false, error: '상품 시트 없음' });

    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
    }

    if (data.products.length > 0) {
      const rows = data.products.map(p => [
        String(p.name || ''),
        Number(p.price) || 0,
        p.active !== false
      ]);
      sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    }
  }

  return jsonResponse({ success: true });
}

/* ============================================
   주문 통계 (관리자)
   ============================================ */
function getOrderStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  const empty = {
    summary: {
      today: { count: 0, revenue: 0, confirmed: 0 },
      week: { count: 0, revenue: 0 },
      month: { count: 0, revenue: 0 },
      all: { count: 0, revenue: 0 }
    },
    pending: { unconfirmed: 0, awaitingShipping: 0 },
    daily: [],
    topProducts: []
  };

  if (!sheet || sheet.getLastRow() < 2) return jsonResponse(empty);

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getValues();

  // 기준 시각
  const tz = 'Asia/Seoul';
  const now = new Date();
  const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const monthStr = Utilities.formatDate(now, tz, 'yyyy-MM');

  // 이번 주 시작 (월요일 00:00)
  const weekStart = new Date(now);
  const dow = (weekStart.getDay() + 6) % 7;  // 월=0, 일=6
  weekStart.setDate(weekStart.getDate() - dow);
  weekStart.setHours(0, 0, 0, 0);

  // 최근 7일 (오늘 포함)
  const dailyMap = {};
  for (let i = 6; i >= 0; i--) {
    const d = new Date(now);
    d.setDate(d.getDate() - i);
    const key = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    dailyMap[key] = { date: key, count: 0, revenue: 0 };
  }

  // 집계 변수
  let allCount = 0, allRevenue = 0;
  let monthCount = 0, monthRevenue = 0;
  let weekCount = 0, weekRevenue = 0;
  let todayCount = 0, todayRevenue = 0, todayConfirmed = 0;
  let unconfirmed = 0, awaitingShipping = 0;
  const productMap = {};  // 확인완료된 것만 집계

  data.forEach(row => {
    if (!row[0]) return;  // 주문번호 없으면 skip

    const orderDate = row[1] instanceof Date ? row[1] : new Date(row[1]);
    if (isNaN(orderDate.getTime())) return;

    const dateKey = Utilities.formatDate(orderDate, tz, 'yyyy-MM-dd');
    const monthKey = Utilities.formatDate(orderDate, tz, 'yyyy-MM');
    const total = Number(row[5]) || 0;
    const paymentStatus = row[11];
    const shipStatus = row[16];

    // 환불은 매출에서 제외
    if (paymentStatus === '환불') return;

    // 전체
    allCount++;
    allRevenue += total;

    // 이번 달
    if (monthKey === monthStr) {
      monthCount++;
      monthRevenue += total;
    }

    // 이번 주
    if (orderDate >= weekStart) {
      weekCount++;
      weekRevenue += total;
    }

    // 오늘
    if (dateKey === todayStr) {
      todayCount++;
      todayRevenue += total;
      if (paymentStatus === '확인완료') todayConfirmed++;
    }

    // 일자별 (최근 7일)
    if (dailyMap[dateKey]) {
      dailyMap[dateKey].count++;
      dailyMap[dateKey].revenue += total;
    }

    // 상태별 카운트
    if (paymentStatus !== '확인완료') unconfirmed++;
    if (paymentStatus === '확인완료' && shipStatus !== '발송완료') awaitingShipping++;

    // 상품별 매출 (확인완료만)
    if (paymentStatus === '확인완료') {
      const name = row[2] || '(상품 미상)';
      if (!productMap[name]) productMap[name] = { name: name, count: 0, revenue: 0 };
      productMap[name].count += Number(row[4]) || 0;
      productMap[name].revenue += total;
    }
  });

  const topProducts = Object.values(productMap)
    .sort((a, b) => b.revenue - a.revenue)
    .slice(0, 10);

  return jsonResponse({
    summary: {
      today: { count: todayCount, revenue: todayRevenue, confirmed: todayConfirmed },
      week: { count: weekCount, revenue: weekRevenue },
      month: { count: monthCount, revenue: monthRevenue },
      all: { count: allCount, revenue: allRevenue }
    },
    pending: { unconfirmed: unconfirmed, awaitingShipping: awaitingShipping },
    daily: Object.values(dailyMap),
    topProducts: topProducts
  });
}

/* ============================================
   유틸리티
   ============================================ */
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================================
   메뉴 추가 (시트에서 바로 실행 가능)
   ============================================ */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🍓 주문관리')
    .addItem('📋 시트 초기 설정', 'setupSheet')
    .addItem('📊 오늘 주문 요약', 'showTodaySummary')
    .addItem('📦 송장 엑셀 다운로드 (확인완료 건)', 'generateShippingSheet')
    .addToUi();
}

/* ============================================
   오늘 주문 요약
   ============================================ */
function showTodaySummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('주문 데이터가 없습니다.');
    return;
  }
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getValues();
  const today = new Date().toLocaleDateString('ko-KR');
  
  let totalOrders = 0;
  let confirmedOrders = 0;
  let pendingOrders = 0;
  let totalRevenue = 0;
  
  data.forEach(row => {
    if (!row[0]) return;
    const orderDate = new Date(row[1]).toLocaleDateString('ko-KR');
    if (orderDate === today) {
      totalOrders++;
      if (row[11] === '확인완료') {
        confirmedOrders++;
        totalRevenue += Number(row[5]);
      } else if (row[11] === '미확인') {
        pendingOrders++;
      }
    }
  });
  
  SpreadsheetApp.getUi().alert(
    `📊 오늘 주문 요약 (${today})\n\n` +
    `총 주문: ${totalOrders}건\n` +
    `입금 확인: ${confirmedOrders}건\n` +
    `입금 대기: ${pendingOrders}건\n` +
    `확인된 매출: ${totalRevenue.toLocaleString()}원`
  );
}

/* ============================================
   송장 시트 생성 — 핵심 로직 (UI 호출 없음)
   ============================================ */
function doGenerateShipping() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet || sheet.getLastRow() < 2) {
    return { success: false, error: '주문 데이터가 없습니다.' };
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 17).getValues();
  const toShip = data.filter(row =>
    row[0] && row[11] === '확인완료' && row[16] !== '발송완료'
  );

  if (toShip.length === 0) {
    return { success: false, error: '발송할 주문이 없습니다. (입금확인 완료 + 미발송 건만 대상)' };
  }

  const sheetName = '송장_' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd_HHmm');
  let shipSheet = ss.getSheetByName(sheetName);
  if (shipSheet) ss.deleteSheet(shipSheet);
  shipSheet = ss.insertSheet(sheetName);

  const shipHeaders = [
    '받는분성명', '받는분전화번호', '받는분기타연락처',
    '받는분우편번호', '받는분주소',
    '운송장번호', '상품명', '수량',
    '보내는분성명', '보내는분전화번호', '보내는분주소',
    '배송메모', '주문번호'
  ];
  shipSheet.getRange(1, 1, 1, shipHeaders.length).setValues([shipHeaders]);
  shipSheet.getRange(1, 1, 1, shipHeaders.length).setFontWeight('bold').setBackground('#BBDEFB');

  // 발송지 주소 - 설정 시트에서 읽음
  const settings = readSettings(ss);
  const SENDER_ADDRESS = settings['발송지 주소'] || '발송지 주소 미설정';

  const rows = toShip.map(row => [
    row[6], row[7], '',
    row[12], row[13],
    '', row[2], row[4],
    row[9] || row[6], row[10] || row[7], SENDER_ADDRESS,
    row[14], row[0]
  ]);
  shipSheet.getRange(2, 1, rows.length, shipHeaders.length).setValues(rows);

  shipSheet.setColumnWidth(1, 100);
  shipSheet.setColumnWidth(2, 130);
  shipSheet.setColumnWidth(4, 80);
  shipSheet.setColumnWidth(5, 300);
  shipSheet.setColumnWidth(7, 200);

  const sheetUrl = ss.getUrl() + '#gid=' + shipSheet.getSheetId();
  return {
    success: true,
    sheetName: sheetName,
    count: rows.length,
    url: sheetUrl
  };
}

/* ============================================
   송장 생성 — 메뉴 호출용 (UI 알림 + 활성 시트 전환)
   ============================================ */
function generateShippingSheet() {
  const result = doGenerateShipping();
  const ui = SpreadsheetApp.getUi();
  if (!result.success) {
    ui.alert(result.error);
    return;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName(result.sheetName));
  ui.alert(
    `✅ 송장 시트 생성 완료!\n\n` +
    `시트명: "${result.sheetName}"\n` +
    `발송 건수: ${result.count}건\n\n` +
    `엑셀(.xlsx)로 다운로드 후 택배사 시스템에 업로드하세요.`
  );
}

/* ============================================
   시트 편집 감지: 송장번호 입력 시 자동 "발송완료"

   ⚠️ 설정 필요 (한 번만):
   Apps Script → ⏰ 트리거 → + 트리거 추가 →
   함수: onSheetEdit / 이벤트: 스프레드시트 수정 시
   ============================================ */
function onSheetEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    if (sheet.getName() !== SHEET_NAME) return;

    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row < 2) return;

    // 송장번호(P열=16) 입력 시 발송완료(Q열=17) 자동 체크
    if (col === 16) {
      const tracking = (e.value || e.range.getValue() || '').toString().trim();
      if (tracking) {
        sheet.getRange(row, 17).setValue('발송완료');
      }
    }
  } catch (error) {
    console.error('onSheetEdit error:', error);
  }
}
