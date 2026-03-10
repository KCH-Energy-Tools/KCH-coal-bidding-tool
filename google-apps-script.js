/**
 * KCHenergy 석탄 입찰 평가 시스템 - Google Apps Script API
 * Web App으로 배포하여 프론트엔드에서 호출
 */

// ==================== 설정 ====================

// 관리자 이메일 목록 - 이 사람들만 공식 입찰 평가 작업 가능
const ADMIN_EMAILS = [
  'lechen@kchenergy.com'
];

// 워크시트 이름
const SHEETS = {
  ICI_BASELINE: 'ICI_기준',
  MINE_DATABASE: '광산DB',
  OFFICIAL_BIDS: '공식입찰',
  PERSONAL_DRAFTS: '개인초안',
  AUDIT_LOG: '작업로그'
};

// ==================== Web App 진입점 ====================

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    const params = e.parameter;
    const action = params.action;
    const userEmail = Session.getActiveUser().getEmail() || params.userEmail || 'anonymous';
    
    let result;
    
    switch (action) {
      // ICI 기준
      case 'getICIBaseline':
        result = getICIBaseline();
        break;
      case 'updateICIBaseline':
        result = updateICIBaseline(JSON.parse(params.data), userEmail);
        break;
        
      // 광산 데이터베이스
      case 'getMineDatabase':
        result = getMineDatabase();
        break;
      case 'addMine':
        result = addMine(JSON.parse(params.data), userEmail);
        break;
      case 'updateMine':
        result = updateMine(JSON.parse(params.data), userEmail);
        break;
        
      // 공식 입찰
      case 'getOfficialBids':
        result = getOfficialBids();
        break;
      case 'createOfficialBid':
        result = createOfficialBid(JSON.parse(params.data), userEmail);
        break;
      case 'updateOfficialBid':
        result = updateOfficialBid(JSON.parse(params.data), userEmail);
        break;
        
      // 개인 초안
      case 'getMyDrafts':
        result = getMyDrafts(userEmail);
        break;
      case 'saveDraft':
        result = saveDraft(JSON.parse(params.data), userEmail);
        break;
      case 'deleteDraft':
        result = deleteDraft(params.draftId, userEmail);
        break;
        
      // 사용자 정보
      case 'getUserInfo':
        result = getUserInfo(userEmail);
        break;
        
      default:
        result = { success: false, error: 'Unknown action: ' + action };
    }
    
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==================== 헬퍼 함수 ====================

function isAdmin(email) {
  return ADMIN_EMAILS.includes(email.toLowerCase());
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    initializeSheet(sheet, name);
  }
  return sheet;
}

function initializeSheet(sheet, name) {
  const headers = {
    [SHEETS.ICI_BASELINE]: ['날짜', 'ICI5000가격', 'NAR', '황분', '회분', '질소함량', '기준운임', '세금', '전력상수', '수정자'],
    [SHEETS.MINE_DATABASE]: ['광산ID', '광산명', '공급업체', '기준NAR', '기준황분', '기준회분', '기준질소', '비고', '등록자', '등록일'],
    [SHEETS.OFFICIAL_BIDS]: ['배치ID', '배치명', '생성일', 'Genco', 'ICI기준평가값', '최적광산', '최적평가값', '전체데이터JSON', '생성자', '상태'],
    [SHEETS.PERSONAL_DRAFTS]: ['초안ID', '사용자이메일', '초안명', '생성일', '수정일', '전체데이터JSON'],
    [SHEETS.AUDIT_LOG]: ['타임스탬프', '사용자이메일', '작업유형', '상세내용']
  };
  
  if (headers[name]) {
    sheet.getRange(1, 1, 1, headers[name].length).setValues([headers[name]]);
    sheet.getRange(1, 1, 1, headers[name].length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

function logAction(userEmail, actionType, details) {
  const sheet = getSheet(SHEETS.AUDIT_LOG);
  sheet.appendRow([new Date().toISOString(), userEmail, actionType, details]);
}

function generateId() {
  return Utilities.getUuid();
}

// ==================== 사용자 정보 ====================

function getUserInfo(email) {
  return {
    success: true,
    data: {
      email: email,
      isAdmin: isAdmin(email),
      permissions: {
        canEditICI: isAdmin(email),
        canCreateOfficialBid: isAdmin(email),
        canEditOfficialBid: isAdmin(email),
        canEditMineDB: true,
        canCreateDraft: true
      }
    }
  };
}

// ==================== ICI 기준 ====================

function getICIBaseline() {
  const sheet = getSheet(SHEETS.ICI_BASELINE);
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return { success: true, data: null };
  }
  
  const lastRow = data[data.length - 1];
  return {
    success: true,
    data: {
      date: lastRow[0],
      price: lastRow[1],
      nar: lastRow[2],
      sulfur: lastRow[3],
      ash: lastRow[4],
      nitrogen: lastRow[5],
      freight: lastRow[6],
      tax: lastRow[7],
      powerConstant: lastRow[8],
      updatedBy: lastRow[9]
    },
    history: data.slice(1).map(row => ({
      date: row[0],
      price: row[1],
      updatedBy: row[9]
    }))
  };
}

function updateICIBaseline(data, userEmail) {
  if (!isAdmin(userEmail)) {
    return { success: false, error: '권한 부족: 관리자만 ICI 기준을 수정할 수 있습니다' };
  }
  
  const sheet = getSheet(SHEETS.ICI_BASELINE);
  sheet.appendRow([
    new Date().toISOString(),
    data.price,
    data.nar,
    data.sulfur,
    data.ash,
    data.nitrogen,
    data.freight,
    data.tax,
    data.powerConstant || 0,
    userEmail
  ]);
  
  logAction(userEmail, 'UPDATE_ICI', `ICI5000: $${data.price}`);
  
  return { success: true };
}

// ==================== 광산 데이터베이스 ====================

function getMineDatabase() {
  const sheet = getSheet(SHEETS.MINE_DATABASE);
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return { success: true, data: [] };
  }
  
  const mines = data.slice(1).map(row => ({
    id: row[0],
    name: row[1],
    supplier: row[2],
    nar: row[3],
    sulfur: row[4],
    ash: row[5],
    nitrogen: row[6],
    notes: row[7],
    createdBy: row[8],
    createdAt: row[9]
  }));
  
  return { success: true, data: mines };
}

function addMine(data, userEmail) {
  const sheet = getSheet(SHEETS.MINE_DATABASE);
  const mineId = generateId();
  
  sheet.appendRow([
    mineId,
    data.name,
    data.supplier,
    data.nar,
    data.sulfur,
    data.ash,
    data.nitrogen,
    data.notes || '',
    userEmail,
    new Date().toISOString()
  ]);
  
  logAction(userEmail, 'ADD_MINE', `광산: ${data.name}`);
  
  return { success: true, data: { id: mineId } };
}

function updateMine(data, userEmail) {
  const sheet = getSheet(SHEETS.MINE_DATABASE);
  const allData = sheet.getDataRange().getValues();
  
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][0] === data.id) {
      sheet.getRange(i + 1, 2, 1, 6).setValues([[
        data.name,
        data.supplier,
        data.nar,
        data.sulfur,
        data.ash,
        data.nitrogen
      ]]);
      
      logAction(userEmail, 'UPDATE_MINE', `광산: ${data.name}`);
      return { success: true };
    }
  }
  
  return { success: false, error: '광산을 찾을 수 없습니다' };
}

// ==================== 공식 입찰 ====================

function getOfficialBids() {
  const sheet = getSheet(SHEETS.OFFICIAL_BIDS);
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return { success: true, data: [] };
  }
  
  const bids = data.slice(1).map(row => ({
    id: row[0],
    name: row[1],
    createdAt: row[2],
    genco: row[3],
    iciEvalValue: row[4],
    bestMine: row[5],
    bestEvalValue: row[6],
    fullData: row[7] ? JSON.parse(row[7]) : null,
    createdBy: row[8],
    status: row[9]
  }));
  
  return { success: true, data: bids };
}

function createOfficialBid(data, userEmail) {
  if (!isAdmin(userEmail)) {
    return { success: false, error: '권한 부족: 관리자만 공식 입찰 평가를 생성할 수 있습니다' };
  }
  
  const sheet = getSheet(SHEETS.OFFICIAL_BIDS);
  const bidId = generateId();
  
  sheet.appendRow([
    bidId,
    data.name,
    new Date().toISOString(),
    data.genco,
    data.iciEvalValue,
    data.bestMine,
    data.bestEvalValue,
    JSON.stringify(data.fullData),
    userEmail,
    data.status || 'draft'
  ]);
  
  logAction(userEmail, 'CREATE_OFFICIAL_BID', `배치: ${data.name}`);
  
  return { success: true, data: { id: bidId } };
}

function updateOfficialBid(data, userEmail) {
  if (!isAdmin(userEmail)) {
    return { success: false, error: '권한 부족: 관리자만 공식 입찰 평가를 수정할 수 있습니다' };
  }
  
  const sheet = getSheet(SHEETS.OFFICIAL_BIDS);
  const allData = sheet.getDataRange().getValues();
  
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][0] === data.id) {
      sheet.getRange(i + 1, 2, 1, 9).setValues([[
        data.name,
        allData[i][2],
        data.genco,
        data.iciEvalValue,
        data.bestMine,
        data.bestEvalValue,
        JSON.stringify(data.fullData),
        allData[i][8],
        data.status
      ]]);
      
      logAction(userEmail, 'UPDATE_OFFICIAL_BID', `배치: ${data.name}`);
      return { success: true };
    }
  }
  
  return { success: false, error: '입찰 평가 기록을 찾을 수 없습니다' };
}

// ==================== 개인 초안 ====================

function getMyDrafts(userEmail) {
  const sheet = getSheet(SHEETS.PERSONAL_DRAFTS);
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return { success: true, data: [] };
  }
  
  const isUserAdmin = isAdmin(userEmail);
  
  const drafts = data.slice(1)
    .filter(row => isUserAdmin || row[1] === userEmail)
    .map(row => ({
      id: row[0],
      userEmail: row[1],
      name: row[2],
      createdAt: row[3],
      updatedAt: row[4],
      fullData: row[5] ? JSON.parse(row[5]) : null
    }));
  
  return { success: true, data: drafts };
}

function saveDraft(data, userEmail) {
  const sheet = getSheet(SHEETS.PERSONAL_DRAFTS);
  const allData = sheet.getDataRange().getValues();
  const now = new Date().toISOString();
  
  if (data.id) {
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === data.id && allData[i][1] === userEmail) {
        sheet.getRange(i + 1, 3, 1, 4).setValues([[
          data.name,
          allData[i][3],
          now,
          JSON.stringify(data.fullData)
        ]]);
        
        logAction(userEmail, 'UPDATE_DRAFT', `초안: ${data.name}`);
        return { success: true, data: { id: data.id } };
      }
    }
  }
  
  const draftId = generateId();
  sheet.appendRow([
    draftId,
    userEmail,
    data.name,
    now,
    now,
    JSON.stringify(data.fullData)
  ]);
  
  logAction(userEmail, 'CREATE_DRAFT', `초안: ${data.name}`);
  
  return { success: true, data: { id: draftId } };
}

function deleteDraft(draftId, userEmail) {
  const sheet = getSheet(SHEETS.PERSONAL_DRAFTS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === draftId) {
      if (data[i][1] !== userEmail && !isAdmin(userEmail)) {
        return { success: false, error: '권한 부족' };
      }
      
      sheet.deleteRow(i + 1);
      logAction(userEmail, 'DELETE_DRAFT', `초안ID: ${draftId}`);
      return { success: true };
    }
  }
  
  return { success: false, error: '초안을 찾을 수 없습니다' };
}

// ==================== 초기화 함수 (수동으로 한 번 실행) ====================

function initializeAllSheets() {
  Object.values(SHEETS).forEach(name => getSheet(name));
  
  const mineSheet = getSheet(SHEETS.MINE_DATABASE);
  if (mineSheet.getLastRow() <= 1) {
    const defaultMines = [
      [generateId(), 'Adaro Indonesia', 'PT Adaro Energy', 4200, 0.10, 3.00, 0.70, '', 'system', new Date().toISOString()],
      [generateId(), 'Kideco Jaya Agung', 'Kideco', 4200, 0.08, 4.50, 0.80, '', 'system', new Date().toISOString()],
      [generateId(), 'Kaltim Prima Coal', 'KPC', 5000, 0.50, 8.00, 1.00, '', 'system', new Date().toISOString()],
      [generateId(), 'Arutmin Indonesia', 'Arutmin', 4700, 0.35, 6.50, 0.90, '', 'system', new Date().toISOString()],
      [generateId(), 'Berau Coal', 'Berau Coal Energy', 4200, 0.20, 5.00, 0.75, '', 'system', new Date().toISOString()],
      [generateId(), 'Indominco Mandiri', 'ITM', 4800, 0.40, 7.00, 0.95, '', 'system', new Date().toISOString()],
      [generateId(), 'Trubaindo Coal', 'Trubaindo', 5200, 0.55, 8.50, 1.05, '', 'system', new Date().toISOString()],
      [generateId(), 'Jembayan Muarabara', 'Jembayan', 4000, 0.15, 4.00, 0.65, '', 'system', new Date().toISOString()],
      [generateId(), 'Bharinto Ekatama', 'Bharinto', 4400, 0.25, 5.50, 0.80, '', 'system', new Date().toISOString()],
      [generateId(), 'Multi Harapan Utama', 'MHU', 4300, 0.18, 4.80, 0.78, '', 'system', new Date().toISOString()]
    ];
    
    defaultMines.forEach(mine => mineSheet.appendRow(mine));
  }
  
  Logger.log('모든 워크시트가 초기화되었습니다!');
}
