/**
 * 匹克球預約與媒合系統 - 後端核心 RESTful API (Code.gs v2.5 - 初階個別點名、補課、三個月保留期)
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
      case 'getSubstituteLessons':
        result = handleGetSubstituteLessons(payload);
        break;
      case 'assignSubstitute':
        result = handleAssignSubstitute(payload);
        break;
      case 'adminProxyRequest':
        result = handleSubmitStudentRequest(payload, true);
        break;
      case 'getLearningData': result = handleLearningData(payload); break;
      case 'saveBasicAttendance': result = handleSaveBasicAttendance(payload); break;
      case 'saveBasicRoster': result = handleSaveBasicRoster(payload); break;
      case 'requestBasicMakeup': result = handleRequestBasicMakeup(payload); break;
      case 'reviewBasicMakeup': result = handleReviewBasicMakeup(payload); break;
      case 'deleteAttendanceRecord': result = deleteAttendanceRecord(payload); break;
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
  const myMatches = matRows.filter(r => r.studentUid === studentUid || parseRoster(r.rosterJson).some(m=>m.ownerUid===studentUid));

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
    return !isSubstituteLesson(m) && m.coachNames && ((coachName && m.coachNames.includes(coachName)) || (coachUid && String(m.coachUid || '').split(/[,、]/).includes(coachUid)));
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
    coaches: getRowsData(coachSheet).concat(SUBSTITUTE_COACHES),
    admins: admins,
    config: config
  };
}

// ==========================================
// 4. POST 交易處理函式群
// ==========================================

// 學員提單 (含長輩代客提單)
function handleSubmitStudentRequestOriginal(p, isProxy) {
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


// 代課選項僅供委員會指派，不建立可登入的假 LINE 帳號。
const SUBSTITUTE_COACHES = [
  { coachUid: 'SUB_COACH_1', realName: '代課教練一號', displayName: '代課教練一號', status: 'Approved', isSubstitute: true },
  { coachUid: 'SUB_COACH_2', realName: '代課教練二號', displayName: '代課教練二號', status: 'Approved', isSubstitute: true }
];
const SUBSTITUTE_DELEGATES = ['Ud471b1a7217b649b3599b54700d352c9', 'U24533fdb919248f20b70d39f89f237c1'];

function substituteOption(value) {
  return SUBSTITUTE_COACHES.find(c => c.coachUid === value || c.realName === value);
}
function isSubstituteLesson(m) {
  return !!(substituteOption(m.coachUid) || substituteOption(m.coachNames));
}
function verifiedLineActor(p) {
  if (!p.accessToken || typeof p.accessToken !== 'string') throw new Error('請由 LINE 登入後再操作。');
  const response = UrlFetchApp.fetch('https://api.line.me/v2/profile', {
    headers: { Authorization: 'Bearer ' + p.accessToken }, muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) throw new Error('LINE 登入已失效，請關閉頁面並重新登入。');
  const profile = JSON.parse(response.getContentText());
  if (!profile.userId) throw new Error('無法驗證 LINE 身分。');
  return profile;
}
function requireSubstituteDelegate(p) {
  const actor = verifiedLineActor(p);
  if (!SUBSTITUTE_DELEGATES.includes(actor.userId)) throw new Error('僅杜和益或黃嘉文可協助代課簽到。');
  return actor;
}
function requireVerifiedAdmin(p) {
  const actor = verifiedLineActor(p);
  const config = getConfigMap();
  const admins = getRowsData(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Admins'));
  if (actor.userId !== config.INITIAL_SUPER_ADMIN && !admins.some(a => a.adminUid === actor.userId && a.status === 'Approved')) {
    throw new Error('需要委員會管理員權限。');
  }
  return actor;
}
function headerKeys(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => {
    const s = String(h).trim();
    const match = s.match(/\(([^)]+)\)/);
    return match ? match[1] : s;
  });
}
function ensureFields(sheet, fields) {
  const headers = headerKeys(sheet);
  const missing = fields.filter(f => !headers.includes(f));
  if (missing.length) {
    const needed = headers.length + missing.length;
    if (needed > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), needed - sheet.getMaxColumns());
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
  }
  return headers.concat(missing);
}
function safeCell(value) {
  return typeof value === 'string' && /^[=+@-]/.test(value) ? "'" + value : value;
}
function appendFields(sheet, values) {
  const headers = ensureFields(sheet, Object.keys(values));
  sheet.appendRow(headers.map(h => safeCell(values[h] === undefined ? '' : values[h])));
}
function updateFields(sheet, idKey, id, values) {
  const row = getRowIndex(sheet, idKey, id);
  if (row < 2) throw new Error('找不到單號：' + id);
  const headers = ensureFields(sheet, Object.keys(values));
  Object.keys(values).forEach(key => sheet.getRange(row, headers.indexOf(key) + 1).setValue(safeCell(values[key])));
}
function requiredSubstituteName(value) {
  const name = String(value || '').trim();
  if (!name || name.length > 80 || /[,、\r\n]/.test(name) || substituteOption(name)) {
    throw new Error('請填寫一位實際代課教練姓名（80 字內，不使用逗號或頓號）。');
  }
  return name;
}
function lessonNotes(value) {
  const notes = String(value || '').trim();
  if (notes.length > 1000) throw new Error('備註請控制在 1000 字內。');
  return notes;
}
function moneyValue(value, label) {
  if (value === '' || value === null || value === undefined || !Number.isFinite(Number(value)) || Number(value) < 0) {
    throw new Error(label + '須填寫零或正數。');
  }
  return Number(value);
}
function lessonWindow(value) {
  const match = String(value || '').match(/^(\d{1,2}):(\d{2})\s*[-~～–—]\s*(\d{1,2}):(\d{2})$/);
  if (!match) throw new Error('課程時段格式不正確。');
  const nums = match.slice(1).map(Number);
  const start = nums[0] * 60 + nums[1], end = nums[2] * 60 + nums[3];
  if (nums[0] > 23 || nums[2] > 23 || nums[1] > 59 || nums[3] > 59 || start >= end) throw new Error('課程起訖時間不正確。');
  return [start, end];
}
function checkAssignmentConflict(rows, proposed, excludedId) {
  const window = lessonWindow(proposed.lessonTime);
  const normalize = s => String(s || '').replace(/教練/g, '').replace(/\s/g, '');
  const proposedNames = String(proposed.substituteCoachName || proposed.coachNames || '').split(/[,、]/).map(normalize).filter(Boolean);
  const conflict = rows.find(m => {
    if (m.matId === excludedId || m.status !== 'Confirmed' || m.lessonDate !== proposed.lessonDate) return false;
    const uids = String(m.coachUid || '').split(/[,、]/);
    const names = String(m.substituteCoachName || m.coachNames || '').split(/[,、]/).map(normalize);
    if (!(proposed.coachUid && uids.includes(proposed.coachUid)) && !proposedNames.some(n => names.includes(n)) && m.coachNames !== proposed.coachNames) return false;
    const existing = lessonWindow(m.lessonTime);
    return window[0] < existing[1] && existing[0] < window[1];
  });
  if (conflict) throw new Error('教練或代課名額與 ' + conflict.matId + ' 時段重疊，請改派其他教練或調整時間。');
}
function handleGetSubstituteLessons(p) {
  const actor = requireSubstituteDelegate(p);
  const lessons = getRowsData(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Matches_MAT'))
    .filter(m => isSubstituteLesson(m) && m.status === 'Confirmed');
  return { status: 'success', lessons: lessons, delegateUid: actor.userId };
}
function handleAssignSubstitute(p) {
  const actor = requireVerifiedAdmin(p);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Matches_MAT');
  const rows = getRowsData(sheet);
  const match = rows.find(m => m.matId === p.matId);
  if (!match || match.status !== 'Confirmed') throw new Error('僅能改派尚未完課的正式課程。');
  const option = substituteOption(p.coachUid);
  if (!option) throw new Error('請選擇代課教練一號或二號。');
  const changes = {
    coachUid: option.coachUid, coachNames: option.realName,
    substituteCoachName: requiredSubstituteName(p.substituteCoachName), notes: lessonNotes(p.notes),
    substituteAssignedBy: actor.userId, substituteAssignedAt: new Date().toISOString()
  };
  checkAssignmentConflict(rows, Object.assign({}, match, changes), match.matId);
  if (!match.originalCoachNames) changes.originalCoachNames = match.coachNames;
  if (!match.originalCoachUid) changes.originalCoachUid = match.coachUid || '';
  updateFields(sheet, 'matId', match.matId, changes);
  return { status: 'success', message: '已保存代課安排。' };
}

function handleSubmitStudentRequest(p, isProxy) {
  if (String(p.courseType || '').includes('初階')) rosterFromRequest(p, 'draft');
  const option = substituteOption(p.designatedCoachUid) || substituteOption(p.designatedCoach);
  let name = '';
  if (option) {
    if (!isProxy) throw new Error('代課教練由委員會安排。');
    const actor = requireVerifiedAdmin(p);
    p.proxyAdminUid = actor.userId;
    p.designatedCoach = option.realName;
    p.designatedCoachUid = option.coachUid;
    name = requiredSubstituteName(p.substituteCoachName);
  }
  p.notes = safeCell(lessonNotes(p.notes));
  const result = handleSubmitStudentRequestOriginal(p, isProxy);
  if (result.status === 'success') {
    updateFields(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Requests_REQ'), 'reqId', result.reqId, {
      designatedCoachUid: p.designatedCoachUid || '', substituteCoachName: name,
      seriesId: p.seriesId || '', participantsJson: JSON.stringify(rosterFromRequest(p,result.reqId))
    });
  }
  return result;
}

function handleConfirmMatch(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Matches_MAT');
  const rows = getRowsData(sheet);
  const request = getRowsData(ss.getSheetByName('Requests_REQ')).find(r => r.reqId === p.reqId);
  if (p.reqId && (!request || !['Pending', 'Interested'].includes(request.status) || rows.some(m => m.reqId === p.reqId && !String(m.status).startsWith('Cancelled')))) {
    throw new Error('需求單已處理或不存在，請重新整理。');
  }
  const linkedIds = Array.from(new Set([p.reqId].concat(Array.isArray(p.linkedReqIds)?p.linkedReqIds:[]).filter(Boolean)));
  const reqRows = getRowsData(ss.getSheetByName('Requests_REQ'));
  const linked = linkedIds.map(id=>reqRows.find(r=>r.reqId===id));
  if (linked.some(r=>!r || !['Pending','Interested'].includes(r.status))) throw new Error('合班需求已處理或不存在。');
  if (linked.length > 1 && linked.some(r=>!basicLesson(r) || r.preferredDate!==p.lessonDate || r.timeSlot!==p.lessonTime)) throw new Error('僅可合併同一天、同時段的初階需求。');
  const linkedCount = linked.length ? linked.reduce((n,r)=>n+Number(r.studentCount||0),0) : Number(p.studentCount);
  if (linked.length > 1 && (linkedCount<4 || linkedCount>6)) throw new Error('合班後須為4～6人。');
  if (basicLesson(p) && (linkedCount<4 || linkedCount>6)) throw new Error('初階正式班須為4～6人，請先合班。');
  const roster = linked.flatMap(r=>{const actual=parseRoster(r.participantsJson);return actual.length?actual:Array.from({length:Number(r.studentCount||0)},(_,i)=>({id:r.reqId+':legacy:'+i,name:i===0?r.studentName:'待補姓名 '+(i+1),ownerUid:r.studentUid,groupName:'舊單待確認'}));});
  const rosterNeedsReview = linked.some(r=>!parseRoster(r.participantsJson).length);
  if (basicLesson(p) && roster.length!==linkedCount) throw new Error('個別名冊人數與合班人數不符，請補齊姓名。');
  const option = substituteOption(p.coachUid) || substituteOption(p.coachNames);
  const actor = requireVerifiedAdmin(p);
  const record = {
    reqId: p.reqId || '', lessonDate: p.lessonDate, lessonTime: p.lessonTime, venue: p.venue,
    studentUid: p.studentUid, studentName: p.studentName,
    coachUid: option ? option.coachUid : (p.coachUid || ''), coachNames: option ? option.realName : p.coachNames,
    courseType: p.courseType, expectedFee: expectedLessonFee(Object.assign({}, p, request || {}, {studentCount:linkedCount})), studentCount: linkedCount || '', linkedReqIds: JSON.stringify(linkedIds), rosterJson: JSON.stringify(roster), rosterNeedsReview:rosterNeedsReview, classId: linked.length ? linked.map(r=>r.seriesId||r.reqId).sort().join('+') : '', status: 'Confirmed',
    managedBy: actor ? actor.userId : p.managedBy,
    substituteCoachName: option ? requiredSubstituteName(p.substituteCoachName || (request && request.substituteCoachName)) : '',
    notes: lessonNotes(p.notes === undefined ? (request && request.notes) : p.notes)
  };
  checkAssignmentConflict(rows, record);
  record.matId = generateSequenceId('MAT');
  appendFields(sheet, record);
  linkedIds.forEach(id=>updateRowStatus(ss.getSheetByName('Requests_REQ'),'reqId',id,'Confirmed'));
  if (p.slotId) updateRowStatus(ss.getSheetByName('Slots_SLOT'), 'slotId', p.slotId, 'Matched');
  return { status: 'success', matId: record.matId };
}

function handleCoachCompleteLesson(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Matches_MAT');
  const match = getRowsData(sheet).find(m => m.matId === p.matId);
  if (!match || match.status !== 'Confirmed') throw new Error('課程不存在或已處理，請重新整理。');
  // 代課路徑無論從哪個頁面呼叫，都需驗證指定兩位代理人的 LINE 身分。
  const actor = isSubstituteLesson(match) ? requireSubstituteDelegate(p) : verifiedLineActor(p);
  if (!isSubstituteLesson(match)) {
    const config = getConfigMap();
    const admins = getRowsData(ss.getSheetByName('Admins'));
    const coach = getRowsData(ss.getSheetByName('Coaches')).find(c => c.coachUid === actor.userId && c.status === 'Approved');
    const assigned = String(match.coachUid || '').split(/[,、]/).includes(actor.userId) ||
      (!match.coachUid && coach && String(match.coachNames || '').split(/[,、]/).some(n => n.replace(/\s*教練\s*/g, '').trim() === String(coach.realName || '').replace(/\s*教練\s*/g, '').trim()));
    if (!assigned && actor.userId !== config.INITIAL_SUPER_ADMIN && !admins.some(a => a.adminUid === actor.userId && a.status === 'Approved')) throw new Error('無此課程簽到權限。');
  }
  if (isSubstituteLesson(match)) requiredSubstituteName(match.substituteCoachName);
  if (basicLesson(match)) {
    const roster=confirmedRoster(match);
    const marked=getRowsData(learningSheet('Basic_Attendance')).filter(x=>x.matId===match.matId);
    if (!roster.length || roster.some(member=>!marked.some(x=>x.memberId===member.id))) throw new Error('請先完成全班逐人點名。');
  }
  const revenue = moneyValue(p.reportedRevenue === undefined ? p.actualRevenue : p.reportedRevenue, '回報實收');
  updateFields(sheet, 'matId', match.matId, {
    reportedRevenue: revenue, attendance: lessonNotes(p.attendance || '全員到課'), completionNotes: lessonNotes(p.notes),
    completedBy: actor.userId, completedByName: actor.displayName || '', completedAt: new Date().toISOString(), status: 'Completed'
  });
  return { status: 'success', message: '已完成簽到，等待委員會核對實收與核銷。' };
}

function handleSettleLessonAccounting(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const matSheet = ss.getSheetByName('Matches_MAT');
  const match = getRowsData(matSheet).find(m => m.matId === p.matId);
  const accSheet = ss.getSheetByName('Accounting_ACC');
  if (!match || match.status !== 'Completed' || getRowsData(accSheet).some(a => a.matId === p.matId)) throw new Error('課程尚未完課或已核銷，請重新整理。');
  const actor = isSubstituteLesson(match) ? requireVerifiedAdmin(p) : null;
  const actualRevenue = moneyValue(p.actualRevenue, '實收學費');
  const venueCost = moneyValue(p.venueCost === undefined || p.venueCost === null || p.venueCost === '' ? 400 : p.venueCost, '場地費');
  const substituteName = isSubstituteLesson(match) ? requiredSubstituteName(match.substituteCoachName) : '';
  const coachList = substituteName || match.coachNames;
  const coachCount = String(coachList || '').split(/[,、]/).filter(s => s.trim()).length;
  if (!coachCount || Number(p.coachCount) !== coachCount) throw new Error('出勤人數與課程教練名單不一致，請先核對名單。');
  const netCoachPool = Math.max(0, actualRevenue - venueCost);
  const perCoachPay = Math.round(netCoachPool / coachCount);
  const accId = generateSequenceId('ACC');
  appendFields(accSheet, {
    accId: accId, matId: match.matId, lessonDate: match.lessonDate, actualRevenue: actualRevenue,
    venueCost: venueCost, netCoachPool: netCoachPool, coachCount: coachCount, perCoachPay: perCoachPay,
    coachList: coachList, settlementStatus: 'Unsettled', settledBy: actor ? actor.userId : p.settledBy,
    substituteSlot: substituteName ? match.coachNames : '', substituteCoachName: substituteName,
    notes: match.notes || '', reportedRevenue: match.reportedRevenue === undefined ? '' : match.reportedRevenue,
    attendance: match.attendance || '', completionNotes: match.completionNotes || '', completedBy: match.completedBy || ''
  });
  updateFields(matSheet, 'matId', match.matId, { actualRevenue: actualRevenue, status: 'Settled' });
  return { status: 'success', accId: accId, perCoachPay: perCoachPay };
}

// 初階是每人五堂套課1500元；MAT 預估費用代表全班的單堂金額。
function expectedLessonFee(request) {
  const count = Number(request.studentCount);
  if (String(request.courseType || '').includes('初階')) {
    if (!Number.isInteger(count) || count < 4 || count > 6) throw new Error('初階團班請填寫 4～6 人。');
    return (1500 / 5) * count;
  }
  // 進階維持原本整堂計價。
  if (Number.isInteger(count) && count >= 1 && count <= 4) return count <= 2 ? 1400 : 1500;
  return moneyValue(request.expectedFee, '預估學費');
}

// 個別名冊、點名及補課。新工作表在首次使用時建立，不更動既有工作表欄位。
const LEARNING_HEADERS = {
  Basic_Attendance: ['attendanceId','matId','lessonDate','memberId','memberName','ownerUid','groupName','status','notes','recordedBy','recordedAt'],
  Basic_Makeup: ['makeupId','sourceAttendanceId','sourceMatId','memberId','memberName','ownerUid','groupName','status','targetMatId','requestedBy','requestedAt','approvedBy','approvedAt','notes']
};
function learningSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1,1,1,LEARNING_HEADERS[name].length).setValues([LEARNING_HEADERS[name]]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
function basicLesson(m) { return String(m.courseType || '').includes('初階'); }
function parseRoster(value) {
  if (Array.isArray(value)) return value;
  try { const result = JSON.parse(value || '[]'); return Array.isArray(result) ? result : []; }
  catch (_) { return []; }
}
function cleanRosterMember(m, fallbackId, owner) {
  const name = String(m.name || '').trim();
  const groupName = String(m.groupName || '').trim();
  if (!name || name.length > 40 || groupName.length > 40) throw new Error('每位學員姓名與組別請填寫 40 字以內。');
  return { id: String(m.id || fallbackId).slice(0,100), name: name,
    ownerUid: String(owner || m.ownerUid || '').slice(0,100), groupName: groupName };
}
function rosterFromRequest(p, reqId) {
  if (!String(p.courseType || '').includes('初階')) return [];
  const source = parseRoster(p.participants);
  const count = Number(p.studentCount);
  if (!Number.isInteger(count) || count < 1 || count > 6 || source.length !== count) {
    throw new Error('初階報名請逐一填寫每位學員姓名，人數須為 1～6 人。');
  }
  const prefix = String(p.seriesId || reqId).replace(/[^a-zA-Z0-9_-]/g,'').slice(0,60);
  return source.map((m,i) => cleanRosterMember(m, prefix + ':' + (i+1), p.studentUid));
}
function confirmedRoster(match) {
  const base = parseRoster(match.rosterJson);
  const extra = getRowsData(learningSheet('Basic_Makeup'))
    .filter(x => x.targetMatId === match.matId && x.status === 'Scheduled')
    .map(x => ({ id: x.memberId, name: x.memberName, ownerUid: x.ownerUid, groupName: x.groupName, makeupId: x.makeupId }));
  return base.concat(extra);
}
function handleLearningData(p) {
  const actor = verifiedLineActor(p);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfigMap();
  const admins = getRowsData(ss.getSheetByName('Admins'));
  const isAdmin = actor.userId === config.INITIAL_SUPER_ADMIN || admins.some(a => a.adminUid === actor.userId && a.status === 'Approved');
  const coach = getRowsData(ss.getSheetByName('Coaches')).find(c => c.coachUid === actor.userId && c.status === 'Approved');
  const matches = getRowsData(ss.getSheetByName('Matches_MAT'));
  const visible = matches.filter(m => basicLesson(m) && (isAdmin || (coach && (String(m.coachUid || '').split(/[,、]/).includes(actor.userId) || (!m.coachUid && String(m.coachNames || '').includes(coach.realName)))) || (isSubstituteLesson(m) && SUBSTITUTE_DELEGATES.includes(actor.userId)) || m.studentUid === actor.userId || parseRoster(m.rosterJson).some(x => x.ownerUid === actor.userId)));
  const ids = new Set(visible.map(m => m.matId));
  const attendance = getRowsData(learningSheet('Basic_Attendance')).filter(x => ids.has(x.matId) && (isAdmin || coach || SUBSTITUTE_DELEGATES.includes(actor.userId) || x.ownerUid === actor.userId));
  const makeup = getRowsData(learningSheet('Basic_Makeup')).filter(x => isAdmin || ids.has(x.sourceMatId) && (coach || SUBSTITUTE_DELEGATES.includes(actor.userId) || x.ownerUid === actor.userId));
  return {status:'success', matches:visible.map(m => {
    const privileged = isAdmin || coach || SUBSTITUTE_DELEGATES.includes(actor.userId);
    const view = privileged ? Object.assign({},m) : {matId:m.matId,lessonDate:m.lessonDate,lessonTime:m.lessonTime,courseType:m.courseType,status:m.status};
    view.roster=confirmedRoster(m).filter(x=>privileged || x.ownerUid===actor.userId);
    return view;
  }), attendance:attendance, makeup:makeup, isAdmin:!!isAdmin};
}
function requireLessonTeacher(p,match) {
  if (isSubstituteLesson(match)) return requireSubstituteDelegate(p);
  const actor = verifiedLineActor(p);
  const coach = getRowsData(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Coaches')).find(c => c.coachUid === actor.userId && c.status === 'Approved');
  const assigned = String(match.coachUid || '').split(/[,、]/).includes(actor.userId) || (coach && !match.coachUid && String(match.coachNames || '').includes(coach.realName));
  if (!assigned) return requireVerifiedAdmin(p);
  return actor;
}
function handleSaveBasicAttendance(p) {
  const match = getRowsData(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Matches_MAT')).find(m => m.matId === p.matId);
  if (!match || !basicLesson(match) || !['Confirmed','Completed','Settled'].includes(match.status)) throw new Error('找不到可點名的初階課程。');
  const actor = requireLessonTeacher(p,match);
  if (match.rosterNeedsReview === true || String(match.rosterNeedsReview).toLowerCase()==='true') throw new Error('舊課程名冊尚待委員會確認姓名。');
  const roster = confirmedRoster(match);
  if (!roster.length) throw new Error('請先由委員會建立個別學員名冊。');
  const entries = Array.isArray(p.entries) ? p.entries : [];
  if (entries.length !== roster.length || new Set(entries.map(e => e.memberId)).size !== roster.length) throw new Error('請逐一點名全部學員。');
  const sheet = learningSheet('Basic_Attendance'), old = getRowsData(sheet);
  const makeups = getRowsData(learningSheet('Basic_Makeup'));
  roster.forEach(member => {
    const e = entries.find(x => x.memberId === member.id);
    if (!e || !['Present','Excused','Absent'].includes(e.status)) throw new Error('點名狀態不完整。');
    const id = match.matId + ':' + member.id;
    const existing = old.find(x => x.attendanceId === id);
    if (existing && existing.status !== e.status && makeups.some(x => x.sourceAttendanceId === id && !['Rejected','Completed'].includes(x.status))) throw new Error('已有待處理補課，請先處理後再修改缺課狀態。');
  });
  roster.forEach(member => {
    const e = entries.find(x => x.memberId === member.id), id = match.matId + ':' + member.id;
    const fields = {matId:match.matId,lessonDate:match.lessonDate,memberId:member.id,memberName:member.name,ownerUid:member.ownerUid || '',groupName:member.groupName || '',status:e.status,notes:lessonNotes(e.notes),recordedBy:actor.userId,recordedAt:new Date().toISOString()};
    if (old.some(x => x.attendanceId === id)) updateFields(sheet,'attendanceId',id,fields);
    else appendFields(sheet,Object.assign({attendanceId:id},fields));
    if (member.makeupId && e.status === 'Present') updateFields(learningSheet('Basic_Makeup'),'makeupId',member.makeupId,{status:'Completed'});
  });
  return {status:'success',message:'個別點名已儲存。'};
}
function handleSaveBasicRoster(p) {
  const actor = requireVerifiedAdmin(p);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Matches_MAT');
  const match = getRowsData(sheet).find(m => m.matId === p.matId);
  if (!match || !basicLesson(match) || match.status !== 'Confirmed') throw new Error('僅能編輯尚未完課的初階名冊。');
  if (getRowsData(learningSheet('Basic_Attendance')).some(x => x.matId === match.matId)) throw new Error('已有點名紀錄，請先處理點名再調整名冊。');
  const roster = parseRoster(p.roster);
  if (roster.length < 4 || roster.length > 6) throw new Error('正式初階班名冊須為 4～6 人。');
  const cleaned = roster.map((m,i) => cleanRosterMember(m,match.matId + ':' + (i+1),m.ownerUid || match.studentUid));
  if (new Set(cleaned.map(x=>x.id)).size !== cleaned.length) throw new Error('名冊識別重複。');
  updateFields(sheet,'matId',match.matId,{rosterJson:JSON.stringify(cleaned),studentCount:cleaned.length,expectedFee:cleaned.length*300,rosterNeedsReview:false,rosterUpdatedBy:actor.userId});
  return {status:'success',message:'名冊已更新。'};
}
function handleRequestBasicMakeup(p) {
  const actor = verifiedLineActor(p);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = getRowsData(learningSheet('Basic_Attendance')).find(x => x.attendanceId === p.attendanceId);
  if (!source || !['Absent','Excused'].includes(source.status)) throw new Error('須有個別缺課或請假紀錄才能申請補課。');
  const admins = getRowsData(ss.getSheetByName('Admins'));
  const admin = actor.userId === getConfigMap().INITIAL_SUPER_ADMIN || admins.some(a=>a.adminUid===actor.userId && a.status==='Approved');
  if (!admin && source.ownerUid !== actor.userId) throw new Error('只能為自己的學員提出補課。');
  const sheet = learningSheet('Basic_Makeup');
  if (getRowsData(sheet).some(x => x.sourceAttendanceId === source.attendanceId && x.status !== 'Rejected')) throw new Error('此缺課已有補課申請。');
  const makeupId = generateSequenceId('MAKEUP');
  appendFields(sheet,{makeupId:makeupId,sourceAttendanceId:source.attendanceId,sourceMatId:source.matId,memberId:source.memberId,memberName:source.memberName,ownerUid:source.ownerUid,groupName:source.groupName,status:'Pending',targetMatId:'',requestedBy:actor.userId,requestedAt:new Date().toISOString(),approvedBy:'',approvedAt:'',notes:lessonNotes(p.notes)});
  return {status:'success',makeupId:makeupId};
}
function handleReviewBasicMakeup(p) {
  const actor = requireVerifiedAdmin(p);
  const sheet = learningSheet('Basic_Makeup');
  const item = getRowsData(sheet).find(x => x.makeupId === p.makeupId);
  if (!item || !['Pending','Approved','Scheduled'].includes(item.status)) throw new Error('找不到待處理補課申請。');
  if (!['Approved','Rejected','Scheduled'].includes(p.status)) throw new Error('補課狀態不正確。');
  let target = '';
  if (p.status === 'Scheduled') {
    target = getRowsData(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Matches_MAT')).find(m=>m.matId===p.targetMatId);
    if (!target || target.status !== 'Confirmed' || !basicLesson(target)) throw new Error('請選擇尚未上課的初階課程。');
    if (target.matId === item.sourceMatId || target.lessonDate <= String(getRowsData(learningSheet('Basic_Attendance')).find(x=>x.attendanceId===item.sourceAttendanceId)?.lessonDate || '')) throw new Error('補課須安排在原缺課日之後。');
    if (confirmedRoster(target).length >= 6) throw new Error('補課班已達六人上限。');
    if (confirmedRoster(target).some(x=>x.id===item.memberId)) throw new Error('學員已在該堂名冊中。');
  }
  updateFields(sheet,'makeupId',item.makeupId,{status:p.status,targetMatId:target ? target.matId : '',approvedBy:actor.userId,approvedAt:new Date().toISOString(),notes:lessonNotes(p.notes === undefined ? item.notes : p.notes)});
  return {status:'success',message:'補課狀態已更新。'};
}
function deleteAttendanceRecord(p) {
  requireVerifiedAdmin(p);
  const sheet = learningSheet('Basic_Attendance');
  const row = getRowIndex(sheet,'attendanceId',p.attendanceId);
  if (row < 2) throw new Error('找不到點名紀錄。');
  if (getRowsData(learningSheet('Basic_Makeup')).some(x=>x.sourceAttendanceId===p.attendanceId && !['Rejected','Completed'].includes(x.status))) throw new Error('仍有待處理補課，不能刪除原點名。');
  sheet.deleteRow(row);
  return {status:'success',message:'點名紀錄已刪除。'};
}
function purgeOldBasicAttendance() {
  const sheet=learningSheet('Basic_Attendance');
  const rows=getRowsData(sheet);
  const makeups=getRowsData(learningSheet('Basic_Makeup'));
  const today=Utilities.formatDate(new Date(),'Asia/Taipei','yyyy-MM-dd');
  const cutoff=new Date(today+'T00:00:00+08:00');
  cutoff.setUTCMonth(cutoff.getUTCMonth()-3);
  const before=Utilities.formatDate(cutoff,'Asia/Taipei','yyyy-MM-dd');
  let removed=0;
  for(let i=rows.length-1;i>=0;i--){
    const x=rows[i];
    if (String(x.lessonDate||'') < before && !makeups.some(m=>m.sourceAttendanceId===x.attendanceId && !['Rejected','Completed'].includes(m.status))){sheet.deleteRow(i+2);removed++;}
  }
  return removed;
}
// 部署後於 Apps Script 執行一次，安裝每日自動清理；重跑不會重複安裝。
function setupBasicAttendanceRetention() {
  if (!ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()==='purgeOldBasicAttendance'))
    ScriptApp.newTrigger('purgeOldBasicAttendance').timeBased().everyDays(1).atHour(3).inTimezone('Asia/Taipei').create();
  return '已設定每日清理三個月前的點名紀錄';
}
