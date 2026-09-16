/**
 * 匹克球預約與媒合系統 - 後端核心 RESTful API (Code.gs v2.2)
 * 負責處理全三端 (學員、教練、委員會) 之 GET 讀取與 POST 寫入交易
 * 整合 LockService 防併發衝堂、自動核發單號與財務核算
 */

// ==========================================
// 1. GET 路由分流 (資料讀取)
// ==========================================
function doGet(e) {
  try {
    const params = e.parameter || {};
    const action = params.action;
    const uid = params.uid;

    let responseData = {};

    switch (action) {
      // 學員端初始化讀取
      case 'getStudentData':
        responseData = handleGetStudentData(uid);
        break;

      // 教練端初始化讀取
      case 'getCoachData':
        responseData = handleGetCoachData(uid);
        break;

      // 委員會端初始化讀取
      case 'getAdminData':
        responseData = handleGetAdminData(uid);
        break;

      // 系統健康檢查
      case 'ping':
      default:
        responseData = { status: 'success', message: 'Pickleball System API is operational.' };
        break;
    }

    return createJsonResponse(responseData);
  } catch (error) {
    return createJsonResponse({ status: 'error', message: error.toString() });
  }
}

// ==========================================
// 2. POST 路由分流 (資料寫入與狀態流轉)
// ==========================================
function doPost(e) {
  // 啟動全域 ScriptLock，防止並行搶單與衝突
  const lock = LockService.getScriptLock();
  try {
    // 最長等待 15 秒以取得鎖
    lock.waitLock(15000);

    let payload = {};
    if (e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else {
      payload = e.parameter || {};
    }

    const action = payload.action;
    let result = {};

    switch (action) {
      // --- 學員端動作 ---
      case 'submitStudentRequest':
        result = handleSubmitStudentRequest(payload);
        break;
      case 'cancelStudentBooking':
        result = handleCancelStudentBooking(payload);
        break;

      // --- 教練端動作 ---
      case 'applyCoach':
        result = handleApplyCoach(payload);
        break;
      case 'submitCoachSlot':
        result = handleSubmitCoachSlot(payload);
        break;
      case 'expressCoachInterest':
        result = handleExpressCoachInterest(payload);
        break;
      case 'coachCompleteLesson':
        result = handleCoachCompleteLesson(payload);
        break;

      // --- 委員會端動作 ---
      case 'applyAdmin':
        result = handleApplyAdmin(payload);
        break;
      case 'approveCoach':
        result = handleApproveCoach(payload);
        break;
      case 'approveAdmin':
        result = handleApproveAdmin(payload);
        break;
      case 'confirmMatch':
        result = handleConfirmMatch(payload);
        break;
      case 'settleLessonAccounting':
        result = handleSettleLessonAccounting(payload);
        break;
      case 'updatePayrollStatus':
        result = handleUpdatePayrollStatus(payload);
        break;
      case 'adminCancelBooking':
        result = handleAdminCancelBooking(payload);
        break;
      case 'adminProxyRequest':
        result = handleSubmitStudentRequest(payload, true);
        break;
      case 'approveCoach':
        result = handleApproveCoach(payload);
        break;
      case 'rejectCoach': // ⭐ 新增這兩行
        result = handleRejectCoach(payload);
        break;
      default:
        result = { status: 'error', message: 'Unknown action requested: ' + action };
        break;
      
    }

    return createJsonResponse(result);

  } catch (error) {
    return createJsonResponse({ status: 'error', message: 'Transaction Locked or Failed: ' + error.toString() });
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 3. GET 處理函式群
// ==========================================

// 學員端資料彙集
function handleGetStudentData(studentUid) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfigMap();

  // 1. 讀取可預約的教練空檔 (Slots_SLOT where status == 'Open')
  const slotSheet = ss.getSheetByName('Slots_SLOT');
  const slotRows = getRowsData(slotSheet);
  const openSlots = slotRows.filter(r => r.status === 'Open');

  // 2. 讀取個人預約紀錄 (Requests_REQ + Matches_MAT)
  const reqSheet = ss.getSheetByName('Requests_REQ');
  const reqRows = getRowsData(reqSheet);
  const myRequests = reqRows.filter(r => r.studentUid === studentUid);

  const matSheet = ss.getSheetByName('Matches_MAT');
  const matRows = getRowsData(matSheet);
  const myMatches = matRows.filter(r => r.studentUid === studentUid);

  // 3. 計算該學員進行中的有效單數 (Pending 或 Confirmed)
  let activeCount = 0;
  myRequests.forEach(r => {
    if (r.status === 'Pending' || r.status === 'Interested') activeCount++;
  });
  myMatches.forEach(m => {
    if (m.status === 'Confirmed') activeCount++;
  });

  // ⭐ 讀取已核准之教練名單供學員指定（加在這裡）
  const coachSheet = ss.getSheetByName('Coaches');
  const coachRows = getRowsData(coachSheet);
  const approvedCoaches = coachRows
    .filter(c => c.status === 'Approved')
    .map(c => c.realName ? (c.realName.includes('教練') ? c.realName : c.realName + ' 教練') : (c.displayName || '教練'));

  return {
    status: 'success',
    studentUid: studentUid,
    activeCount: activeCount,
    maxLimit: parseInt(config.MAX_ACTIVE_BOOKINGS || '2'),
    openSlots: openSlots,
    myRequests: myRequests,
    myMatches: myMatches,
    approvedCoaches: approvedCoaches, // ⭐ 回傳名冊給前端
    config: config
  };
}

// 教練端資料彙集
function handleGetCoachData(coachUid) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 檢查白名單資格
  const coachSheet = ss.getSheetByName('Coaches');
  const coaches = getRowsData(coachSheet);
  const coachInfo = coaches.find(c => c.coachUid === coachUid);
  const isApproved = coachInfo && coachInfo.status === 'Approved';

  // 2. 需求池 (Requests_REQ where status in ['Pending', 'Interested'])
  const reqSheet = ss.getSheetByName('Requests_REQ');
  const reqRows = getRowsData(reqSheet);
  const poolRequests = reqRows.filter(r => r.status === 'Pending' || r.status === 'Interested');

  // 3. 我的授課課表 (Matches_MAT)
  const matSheet = ss.getSheetByName('Matches_MAT');
  const matRows = getRowsData(matSheet);
  const coachName = coachInfo ? coachInfo.realName : '';
  const myLessons = matRows.filter(m => {
    return m.coachNames && (m.coachNames.includes(coachName) || m.coachNames.includes(coachUid));
  });

  // 4. 我的釋出空檔 (Slots_SLOT)
  const slotSheet = ss.getSheetByName('Slots_SLOT');
  const slotRows = getRowsData(slotSheet);
  const mySlots = slotRows.filter(s => s.coachUid === coachUid);

  return {
    status: 'success',
    isApproved: !!isApproved,
    coachInfo: coachInfo || null,
    poolRequests: poolRequests,
    myLessons: myLessons,
    mySlots: mySlots
  };
}

// 委員會端資料彙集
function handleGetAdminData(adminUid) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfigMap();

  // 1. 權限檢核 (Admins 表或創世管理員)
  const adminSheet = ss.getSheetByName('Admins');
  const admins = getRowsData(adminSheet);
  const adminInfo = admins.find(a => a.adminUid === adminUid);
  const isSuperAdmin = adminUid === config.INITIAL_SUPER_ADMIN;
  const isApproved = isSuperAdmin || (adminInfo && adminInfo.status === 'Approved');

  if (!isApproved) {
    return {
      status: 'unauthorized',
      message: '尚未取得委員會管理員權限',
      adminUid: adminUid
    };
  }

  // 2. 撈取撮合工作台與記帳所需資料
  const reqSheet = ss.getSheetByName('Requests_REQ');
  const slotSheet = ss.getSheetByName('Slots_SLOT');
  const matSheet = ss.getSheetByName('Matches_MAT');
  const accSheet = ss.getSheetByName('Accounting_ACC');
  const coachSheet = ss.getSheetByName('Coaches');

  return {
    status: 'success',
    role: isSuperAdmin ? 'SuperAdmin' : (adminInfo ? adminInfo.role : 'Admin'),
    requests: getRowsData(reqSheet),
    slots: getRowsData(slotSheet),
    matches: getRowsData(matSheet),
    accounting: getRowsData(accSheet),
    coaches: getRowsData(coachSheet),
    admins: admins,
    config: config
  };
}

// ==========================================
// 4. POST 交易處理函式群
// ==========================================

// 學員提單 (含長輩代客提單)
function handleSubmitStudentRequest(p, isProxy) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfigMap();
  const studentUid = p.studentUid;

  // 1. 檢驗未結課上限 (< 2 筆)
  if (!isProxy) {
    const studentData = handleGetStudentData(studentUid);
    if (studentData.activeCount >= parseInt(config.MAX_ACTIVE_BOOKINGS || '2')) {
      return { status: 'error', message: '您目前已有 2 筆進行中的預約，已達上限！' };
    }
  }

  // 2. 自動生成 REQ- 流水單號
  const reqId = generateSequenceId('REQ');
  const createdAt = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');

  const reqSheet = ss.getSheetByName('Requests_REQ');
  reqSheet.appendRow([
    reqId,
    createdAt,
    studentUid,
    p.studentName || 'LINE學員',
    p.courseType || '初階5堂包套',
    p.preferredDate,
    p.timeSlot,
    p.level || '2.0',
    p.studentCount || 4,
    p.designatedCoach || '不指定(由委員會推薦)',
    p.venuePreference || '宜蘭國民運動中心',
    p.notes || '',
    JSON.stringify([]), // interestedCoaches 初始為空
    'Pending',
    isProxy ? (p.proxyAdminUid || 'ADMIN') : ''
  ]);

  return { status: 'success', reqId: reqId };
}

// 學員線上自主取消預約 (含 24 小時門檻防呆)
function handleCancelStudentBooking(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfigMap();
  const limitHours = parseFloat(config.CANCELLATION_LIMIT_HOURS || '24');

  const lessonDateStr = p.lessonDate; // YYYY-MM-DD
  const lessonDate = new Date(lessonDateStr + 'T08:00:00');
  const now = new Date();
  const diffHours = (lessonDate.getTime() - now.getTime()) / (1000 * 60 * 60);

  if (diffHours < limitHours) {
    return { status: 'error', message: '距開課已不足 24 小時，請直接在 LINE 群組聯絡委員會幹部協助處理。' };
  }

  // 執行取消狀態變更
  if (p.idType === 'REQ') {
    updateRowStatus(ss.getSheetByName('Requests_REQ'), 'reqId', p.id, 'Cancelled_User');
  } else if (p.idType === 'MAT') {
    updateRowStatus(ss.getSheetByName('Matches_MAT'), 'matId', p.id, 'Cancelled_User');
  }

  return { status: 'success', message: '預約已自主取消' };
}

// 教練發布空檔
function handleSubmitCoachSlot(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const slotId = generateSequenceId('SLOT');
  const createdAt = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');

  const slotSheet = ss.getSheetByName('Slots_SLOT');
  slotSheet.appendRow([
    slotId,
    createdAt,
    p.coachUid,
    p.coachName,
    p.slotDate,
    p.timeSlot,
    p.venueOption || '宜蘭國民運動中心',
    'Open'
  ]);

  return { status: 'success', slotId: slotId };
}

// 教練意向認領
function handleExpressCoachInterest(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Requests_REQ');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === p.reqId) {
      let interested = [];
      try {
        interested = JSON.parse(data[i][12] || '[]');
      } catch (e) {
        interested = [];
      }

      if (!interested.includes(p.coachName)) {
        interested.push(p.coachName);
      }

      sheet.getRange(i + 1, 13).setValue(JSON.stringify(interested));
      sheet.getRange(i + 1, 14).setValue('Interested');
      return { status: 'success', message: '已成功登記承接意向！' };
    }
  }

  return { status: 'error', message: '查無該筆需求單號' };
}

// 委員會確認撮合 (二次防衝堂檢核)
function handleConfirmMatch(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const matSheet = ss.getSheetByName('Matches_MAT');
  const matRows = getRowsData(matSheet);

  // 1. 執行後端防衝堂比對
  const assignedCoaches = p.coachNames.split(',').map(s => s.trim());
  const conflict = matRows.find(m => {
    if (m.status !== 'Confirmed') return false;
    if (m.lessonDate === p.lessonDate && m.lessonTime === p.lessonTime) {
      return assignedCoaches.some(c => m.coachNames.includes(c));
    }
    return false;
  });

  if (conflict) {
    return {
      status: 'error',
      message: `衝堂警告！教練名冊在 ${p.lessonDate} ${p.lessonTime} 已有正式排程 (${conflict.matId})，禁止重複排課！`
    };
  }

  // 2. 核發 MAT- 媒合單號
  const matId = generateSequenceId('MAT');
  matSheet.appendRow([
    matId,
    p.reqId || '',
    p.lessonDate,
    p.lessonTime,
    p.venue,
    p.studentUid,
    p.studentName,
    p.coachNames,
    p.courseType,
    p.expectedFee || 0,
    'Confirmed',
    p.managedBy
  ]);

  // 3. 同步將來源需求標記為 Confirmed
  if (p.reqId) {
    updateRowStatus(ss.getSheetByName('Requests_REQ'), 'reqId', p.reqId, 'Confirmed');
  }

  // 4. 若有關聯空檔單，標記為 Matched
  if (p.slotId) {
    updateRowStatus(ss.getSheetByName('Slots_SLOT'), 'slotId', p.slotId, 'Matched');
  }

  return { status: 'success', matId: matId };
}

// 教練完課簽到回報
function handleCoachCompleteLesson(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const matSheet = ss.getSheetByName('Matches_MAT');
  const data = matSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === p.matId) {
      matSheet.getRange(i + 1, 10).setValue(p.actualRevenue); // 更新實收金額
      matSheet.getRange(i + 1, 11).setValue('Completed');     // 標記完課
      if (p.coachNames) {
        matSheet.getRange(i + 1, 8).setValue(p.coachNames);   // 更新出勤教練清單
      }
      return { status: 'success', message: '已完成完課簽到，等待委員會均分核銷！' };
    }
  }

  return { status: 'error', message: '找不到對應的媒合排程單號' };
}

// 委員會課堂收支均分核銷 (生成 ACC- 傳票)
function handleSettleLessonAccounting(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const actualRevenue = parseFloat(p.actualRevenue || '0');
  const venueCost = parseFloat(p.venueCost || '0');
  const coachCount = parseInt(p.coachCount || '1');

  // 計算可分配酬勞與均分款
  const netCoachPool = actualRevenue - venueCost;
  const perCoachPay = Math.round(netCoachPool / coachCount);

  const accId = generateSequenceId('ACC');
  const accSheet = ss.getSheetByName('Accounting_ACC');

  accSheet.appendRow([
    accId,
    p.matId,
    p.lessonDate,
    actualRevenue,
    venueCost,
    netCoachPool,
    coachCount,
    perCoachPay,
    p.coachList,
    'Unsettled', // 薪酬發放初始狀態為待結算發放
    p.settledBy
  ]);

  // 同步 Matches_MAT 為 Settled
  updateRowStatus(ss.getSheetByName('Matches_MAT'), 'matId', p.matId, 'Settled');

  return { status: 'success', accId: accId, perCoachPay: perCoachPay };
}

// 委員會教練薪資發放結清標記（安全修正版）
function handleUpdatePayrollStatus(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accSheet = ss.getSheetByName('Accounting_ACC');
  const data = accSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    // 比對記帳單號 accId (第 1 欄，索引 0)
    if (String(data[i][0]) === String(p.accId)) {
      // 第 10 欄為 settlementStatus（核銷狀態）
      accSheet.getRange(i + 1, 10).setValue('Settled');
      return { status: 'success', message: '該堂薪酬已標記發放結清！' };
    }
  }
  return { status: 'error', message: '查無該筆記帳單號：' + p.accId };
}


// 教練申請
function handleApplyCoach(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Coaches');
  sheet.appendRow([
    p.coachUid,
    p.realName || '',
    p.displayName || '教練申請人',
    p.duprId || '',
    p.duprRating || '',
    p.licenseNote || '',
    p.phone || '',
    'Coach_Pending',
    ''
  ]);
  return { status: 'success', message: '教練資格申請已送出，等待委員會審核！' };
}

// 核准教練 (精確更新狀態為 Coach_Pending 的該列)
function handleApproveCoach(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Coaches');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    // 同時比對 UID 與待審核狀態
    if (String(data[i][0]) === String(p.coachUid) && data[i][7] === 'Coach_Pending') {
      sheet.getRange(i + 1, 8).setValue('Approved');
      sheet.getRange(i + 1, 9).setValue(p.approvedBy || 'ADMIN');
      return { status: 'success', message: '教練已成功核准加入白名單！' };
    }
  }
  return { status: 'error', message: '查無待審核記錄' };
}

// 拒絕教練申請 (直接將誤按/不符資格的那一列刪除)
function handleRejectCoach(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Coaches');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    // 找到該筆待審核紀錄直接刪除整列
    if (String(data[i][0]) === String(p.coachUid) && data[i][7] === 'Coach_Pending') {
      sheet.deleteRow(i + 1);
      return { status: 'success', message: '已拒絕並刪除該筆教練申請！' };
    }
  }
  return { status: 'error', message: '查無該筆待審核申請' };
}

// 委員會幹部申請
function handleApplyAdmin(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Admins');
  sheet.appendRow([
    p.adminUid,
    p.realName || '',
    p.displayName || '幹部申請人',
    'Admin',
    'Pending',
    ''
  ]);
  return { status: 'success', message: '幹部授權申請已送出！' };
}

// 核准幹部
function handleApproveAdmin(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Admins');
  updateRowStatus(sheet, 'adminUid', p.adminUid, 'Approved', 5);
  sheet.getRange(getRowIndex(sheet, 'adminUid', p.adminUid), 6).setValue(p.approvedBy);
  return { status: 'success', message: '新幹部權限已開通！' };
}

// 委員會代客強制取消
function handleAdminCancelBooking(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (p.idType === 'REQ') {
    updateRowStatus(ss.getSheetByName('Requests_REQ'), 'reqId', p.id, 'Cancelled_Admin');
  } else {
    updateRowStatus(ss.getSheetByName('Matches_MAT'), 'matId', p.id, 'Cancelled_Admin');
  }
  return { status: 'success', message: '已執行委員會代客強制取消' };
}

// ==========================================
// 5. 底層通用輔助工具 (Helpers)
// ==========================================

// 生成標準單號 (REQ / SLOT / MAT / ACC - YYYYMMDD - XXX)
function generateSequenceId(prefix) {
  const dateStr = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd');
  const randomSuffix = Math.floor(100 + Math.random() * 900); // 隨機三位流水碼
  return `${prefix}-${dateStr}-${randomSuffix}`;
}

// 讀取工作表並轉換為物件陣列 (Header 映射)
function getRowsData(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  // 清理欄位標題 (例如 "需求單號 (reqId)" 抽取為 "reqId")
  const headers = data[0].map(h => {
    const match = h.match(/\(([^)]+)\)/);
    return match ? match[1] : h.trim();
  });

  const result = [];
  for (let i = 1; i < data.length; i++) {
    const rowObj = {};
    for (let j = 0; j < headers.length; j++) {
      let val = data[i][j];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, 'Asia/Taipei', 'yyyy-MM-dd');
      }
      rowObj[headers[j]] = val;
    }
    result.push(rowObj);
  }
  return result;
}
// 更新指定單號之狀態欄位
function updateRowStatus(sheet, idKey, targetId, newStatus, fallbackColIndex) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => {
    const match = h.match(/\(([^)]+)\)/);
    return match ? match[1] : h.trim();
  });

  const idCol = headers.indexOf(idKey);
  let statusCol = headers.indexOf('status');
  
  if (statusCol === -1) {
    statusCol = headers.indexOf('settlementStatus');
  }
  if (statusCol === -1 && fallbackColIndex !== undefined) {
    statusCol = fallbackColIndex - 1;
  }
  
  // 欄位防呆：若找不到欄位，強制設為合法欄位，避免報出起始欄過小
  const finalCol = (statusCol >= 0 ? statusCol : (fallbackColIndex || 10) - 1) + 1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol >= 0 ? idCol : 0]) === String(targetId)) {
      sheet.getRange(i + 1, finalCol).setValue(newStatus);
      break;
    }
  }
}


// 取得目標 ID 所在的 Row Index
function getRowIndex(sheet, idKey, targetId) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => {
    const match = h.match(/\(([^)]+)\)/);
    return match ? match[1] : h.trim();
  });
  const idCol = headers.indexOf(idKey);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(targetId)) {
      return i + 1;
    }
  }
  return -1;
}

// 讀取 Config 工作表設定值
function getConfigMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');
  const map = {};
  if (!sheet) return map;
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0]) {
      map[values[i][0]] = values[i][1];
    }
  }
  return map;
}

// 產生標準 JSON 輸出 (處理跨域 CORS 轉向)
function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
