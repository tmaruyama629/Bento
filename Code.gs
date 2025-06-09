/**
 * ファイル名: Code.gs
 *
 * 変更履歴:
 * 2025/05/08 T.Maruyama  新規作成
 */

function doGet() {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate().setTitle('お弁当注文フォーム');
}

const CONFIG = {
  MASTER_ID: '1s00XO8VNkN4NMi1OSbRhquqU-H8gj-Lg_bQbVLKwfP0',
  ORDER_ID: '1Mzj9Oxz3NWVmvYebdw3Bne9HrJ-v_-0nJo9TZ9xM2pI',
  ORDER_SHEET: 'D_Order'
};

// 汎用設定取得
function getConfigValue(key) {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.MASTER_ID);
  const configSheet = spreadsheet.getSheetByName('M_Config');
  const configValues = configSheet.getDataRange().getValues(); // すべて取得

  for (let i = 0; i < configValues.length; i++) {
    if (configValues[i][0] === key) {
      return configValues[i][1];
    }
  }
  return null; // 見つからなかった場合
}

// パスワード検証（汎用設定取得を使用）
function verifyPassword(inputPw) {
  const adminPw = getConfigValue('AdminPassword');
  const userPw = getConfigValue('UserPassword');

  if (inputPw === adminPw) {
    return {
      isValid: true,
      isAdmin: true
    };
  } else if (inputPw === userPw) {
    return {
      isValid: true,
      isAdmin: false
    };
  } else {
    return {
      isValid: false,
      isAdmin: false
    };
  }
}

// 共通のスプレッドシート取得関数
function getSpreadsheet(sheetId) {
  try {
    return SpreadsheetApp.openById(sheetId);
  } catch (error) {
    throw new Error(`スプレッドシートの取得に失敗しました: ${sheetId}`);
  }
}

// データ取得共通化
function getDataFromSheet(sheetId, sheetName) {
  const sheet = getSpreadsheet(sheetId).getSheetByName(sheetName);
  const values = sheet.getDataRange().getValues();
  return values.slice(1); // ヘッダー除外
}

// 社員一覧取得
function getEmployeeList() {
  const values = getDataFromSheet(CONFIG.MASTER_ID, 'M_Employee');
  return values.map(row => ({
    EmployeeCD: row[0],
    EmployeeName: row[1]
  }));
}

// 工場一覧取得
function getFactoryList() {
  const values = getDataFromSheet(CONFIG.MASTER_ID, 'M_Factory');
  return values.map(row => ({
    FactoryCD: row[0],
    FactoryName: row[1]
  }));
}

// 社員のデフォルト工場を取得
function getEmployeeData(empCD) {
  const values = getDataFromSheet(CONFIG.MASTER_ID, 'M_Employee');
  const employee = values.find(row => row[0] === empCD);
  return { defaultFactory: employee ? employee[2] : '' };
}

// 休日取得
function getHolidayMap() {
  const values = getDataFromSheet(CONFIG.MASTER_ID, 'M_Holiday');
  const holidayMap = {};
  values.forEach(row => {
    const date = formatDate(row[0]); // yyyy-MM-dd に変換
    holidayMap[date] = true;
  });
  return holidayMap;
}

// コメント取得
function getComment() {
  const values = getDataFromSheet(CONFIG.MASTER_ID, 'M_Comment');
  return values.map(row => ({
    CommentCD: row[0],
    CommentText: row[1],
    HyperLink: row[2]
  }));
}


// 指定週のメニューを取得
function getMenuForWeek(empCD, startDate, endDate, isAdmin = false) {
  const now = new Date();
  const currentDateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  const weekDays = getDaysInRange(startDate, endDate);
  const spreadsheet = getSpreadsheet(CONFIG.MASTER_ID);

  // M_Config から締切時間を取得
  const deadlineStr = getConfigValue('DeadlineTime') || '09:00';
  const [deadlineHour, deadlineMinute] = deadlineStr.split(':').map(Number);

  const menuSheet = spreadsheet.getSheetByName('M_Menu');
  const venderSheet = spreadsheet.getSheetByName('M_Vender');
  const orderSheet = getSpreadsheet(CONFIG.ORDER_ID).getSheetByName(CONFIG.ORDER_SHEET);
  const holidayMap = getHolidayMap();

  if (!menuSheet || !venderSheet || !orderSheet) {
    Logger.log("必要なシートが見つかりません。");
    return [];
  }

  const menuValues = menuSheet.getDataRange().getValues();
  const venderValues = venderSheet.getDataRange().getValues();
  const orderValues = orderSheet.getDataRange().getValues();

  const venderMap = createVenderMap(venderValues);
  const menuMap = createMenuMap(menuValues, venderMap);

  const weekData = weekDays.map(date => {
    const isHoliday = !!holidayMap[date];
    const dateObj = new Date(date);
    const isSameDay = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy/MM/dd') === currentDateStr;

    // 締切時間の Date を生成して比較
    const deadline = new Date(dateObj);
    deadline.setHours(deadlineHour, deadlineMinute, 0, 0);
    const isPastDeadline = !isAdmin && isSameDay && now > deadline;
    const isClosed = isHoliday || (isAdmin ? false : isPastDeadline);

    // 注文情報取得
    const order = findOrderForDate(orderValues, empCD, date);

    // 注文済み情報セット
    let orderedMenuCD = '';
    let orderedMenuName = '';
    let orderedVenderCD = '';
    let orderedVenderName = '';
    let orderedFactoryCD = '';
    let ordered = false;

    if (order) {
      const menuCD = order[4]?.toString().trim();
      const venderCD = order[5]?.toString().trim();
      const factoryCD = order[3]?.toString().trim();
      const key = `${menuCD}_${venderCD}`;
      if (menuMap[key]) {
        orderedMenuCD = menuMap[key].MenuCD;
        orderedMenuName = menuMap[key].MenuName;
        orderedVenderCD = venderCD;
        orderedVenderName = venderMap[venderCD] || '';
        orderedFactoryCD = factoryCD;
        ordered = true;
      }
    }

    // メニュー詳細リスト生成
    const menuDetails = getMenuDetailsForDate(date, menuValues, venderMap, isAdmin);

    return {
      Date: date,
      IsHoliday: isHoliday,
      IsClosed: isClosed,
      Ordered: ordered,
      MenuCD: orderedMenuCD,
      MenuName: orderedMenuName,
      VenderCD: orderedVenderCD,
      VenderName: orderedVenderName,
      FactoryCD: order ? order[3] : '',
      Menus: menuDetails
    };
  });

  return weekData;
}

// ベンダーマップ生成
function createVenderMap(venderValues) {
  const venderMap = {};
  venderValues.slice(1).forEach(([venderCD, venderName]) => {
    venderMap[venderCD] = venderName;
  });
  return venderMap;
}

// メニューマップ生成
function createMenuMap(menuValues, venderMap) {
  const menuMap = {};
  menuValues.slice(1).forEach(([menuCD, venderCD, menuName]) => {
    const key = `${menuCD}_${venderCD}`;
    menuMap[key] = {
      MenuCD: menuCD,
      MenuName: menuName,
      VenderCD: venderCD,
      DisplayName: `${menuCD}_${menuName}[${venderMap[venderCD] || ''}]`
    };
  });
  return menuMap;
}

// 注文情報取得
function findOrderForDate(orderValues, empCD, date) {
  return orderValues.find(row =>
    row[2] === empCD.padStart(6, '0') && formatDate(row[1]) === date && row[6] === 1
  );
}

// メニュー詳細リスト生成
function getMenuDetailsForDate(date, menuValues, venderMap, isAdmin) {
  const menus = filterMenusForDate(date, menuValues);
  return menus
    .filter(m => isAdmin || m[10] === 1)
    .map(m => {
      const venderCD = m[1];
      const venderName = venderMap[venderCD] || '';
      return {
        MenuCD: m[0],
        VenderCD: venderCD,
        MenuName: m[2],
        VenderName: venderName,
        DisplayName: `${m[0]}_${m[2]}[${venderName}]`
      };
    });
}

// 日付範囲取得
function getDaysInRange(startDate, endDate) {
  const days = [];
  let currentDate = new Date(startDate);
  const end = new Date(endDate);
  while (currentDate <= end) {
    days.push(formatDate(currentDate));
    currentDate.setDate(currentDate.getDate() + 1);
  }
  return days;
}

// メニュー抽出
function filterMenusForDate(date, menuValues) {
  const jsDate = new Date(date);
  const dow = jsDate.getDay(); // 0:Sun ～ 6:Sat
  const colIndex = 3 + dow; // M_Menuシートの曜日列は3列目から
  return menuValues.slice(1).filter(row => row[colIndex] === 1);
}

// 日付フォーマット
function formatDate(value) {
  if (typeof value === 'number') value = value.toString();
  if (/^\d{8}$/.test(value)) {
    return `${value.substring(0,4)}-${value.substring(4,6)}-${value.substring(6,8)}`;
  }

  const date = new Date(value);
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

// 注文保存処理
function saveOrderData(employeeCD, orders, isAdmin) {
  const holidayMap = getHolidayMap();
  const dbSheet = SpreadsheetApp.openById(CONFIG.ORDER_ID).getSheetByName(CONFIG.ORDER_SHEET);
  const now = new Date();
  const insertedOrders = [];
  const data = dbSheet.getDataRange().getValues();
  const header = data[0];

  // マスタデータ取得
  const master = getMasterData();

  orders.forEach(order => {
    const nowDateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    if (!isAdmin && order.OrderDate < nowDateStr) return;

    const orderDate = parseOrderDate(order.OrderDate);
    const orderDateStr = Utilities.formatDate(orderDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');

    // 締切・休日判定
    if (isOrderClosed(orderDate, now, nowDateStr, orderDateStr, holidayMap, isAdmin, master.deadlineHour, master.deadlineMinute)) return;

    // 既存注文検索
    const existingRow = findExistingOrderRow(data, header, employeeCD, orderDateStr);

    // 工場のみ変更判定
    if (!order.MenuCD || !order.VenderCD) {
      if (!isFactoryOnlyChange(order, existingRow)) return;
    }

    // 既存注文があればActiveFlgを0に
    if (existingRow > 0) {
      dbSheet.getRange(existingRow + 1, master.activeFlgIdx + 1).setValue(0);
      if ((!order.MenuCD || !order.VenderCD)) {
        const oldRow = data[existingRow];
        order.MenuCD = oldRow[master.menuCDIdx] || '';
        order.VenderCD = oldRow[master.venderCDIdx] || '';
      }
    }

    // 注文行作成
    const newRow = createOrderRow(order, orderDate, employeeCD, now, master);

    const lastRow = dbSheet.getLastRow() + 1;
    dbSheet.getRange(lastRow, 1, 1, 8).setValues([newRow]);

    insertedOrders.push({
      OrderDate: orderDateStr,
      MenuCD: order.MenuCD,
      VenderCD: order.VenderCD,
      MenuName: master.menuMap.get(`${order.MenuCD}|${order.VenderCD}`) || '不明なメニュー',
      FactoryCD: order.FactoryCD ?? '',
      VenderName: master.venderMap.get(order.VenderCD) || '不明なベンダー'
    });
  });

  Logger.log('insertedOrders: ' + JSON.stringify(insertedOrders));
  return insertedOrders;
}

// マスタデータ取得
function getMasterData() {
  const menuMasterSheet = SpreadsheetApp.openById(CONFIG.MASTER_ID).getSheetByName('M_Menu');
  const menuData = menuMasterSheet.getDataRange().getValues();
  const menuHeader = menuData[0];
  const menuCDIdxInMaster = menuHeader.indexOf('MenuCD');
  const venderCDIdxInMaster = menuHeader.indexOf('VenderCD');
  const menuNameIdx = menuHeader.indexOf('MenuName');

  const venderSheet = SpreadsheetApp.openById(CONFIG.MASTER_ID).getSheetByName('M_Vender');
  const venderData = venderSheet.getDataRange().getValues();
  const venderHeader = venderData[0];
  const venderCDIdxInVender = venderHeader.indexOf('VenderCD');
  const venderNameIdx = venderHeader.indexOf('VenderName');

  const configSheet = SpreadsheetApp.openById(CONFIG.MASTER_ID).getSheetByName('M_Config');
  const configData = configSheet.getDataRange().getValues();
  const configMap = {};
  for (let i = 1; i < configData.length; i++) {
    const [key, value] = configData[i];
    configMap[key] = value;
  }
  const deadlineStr = configMap['DeadlineTime'] || '09:00';
  const [deadlineHour, deadlineMinute] = deadlineStr.split(':').map(Number);

  const menuMap = new Map();
  for (let i = 1; i < menuData.length; i++) {
    const row = menuData[i];
    const key = `${row[menuCDIdxInMaster]}|${row[venderCDIdxInMaster]}`;
    menuMap.set(key, row[menuNameIdx]);
  }

  const venderMap = new Map();
  for (let i = 1; i < venderData.length; i++) {
    const row = venderData[i];
    venderMap.set(row[venderCDIdxInVender], row[venderNameIdx]);
  }

  // dbSheet header index
  const dbSheet = SpreadsheetApp.openById(CONFIG.ORDER_ID).getSheetByName(CONFIG.ORDER_SHEET);
  const dbHeader = dbSheet.getDataRange().getValues()[0];
  return {
    menuMap,
    venderMap,
    deadlineHour,
    deadlineMinute,
    menuCDIdx: dbHeader.indexOf('MenuCD'),
    venderCDIdx: dbHeader.indexOf('VenderCD'),
    activeFlgIdx: dbHeader.indexOf('ActiveFlg')
  };
}

// 締切・休日判定
function isOrderClosed(orderDate, now, nowDateStr, orderDateStr, holidayMap, isAdmin, deadlineHour, deadlineMinute) {
  const isSameDay = nowDateStr === orderDateStr;
  const isPastDate = new Date(orderDateStr) < new Date(nowDateStr);
  const deadline = new Date(orderDate);
  deadline.setHours(deadlineHour, deadlineMinute, 0, 0);
  const isPastDeadline = isSameDay && now > deadline;
  const isHoliday = !!holidayMap[orderDateStr];
  return !isAdmin && (isHoliday || isPastDeadline || isPastDate);
}

// 既存注文検索
function findExistingOrderRow(data, header, employeeCD, orderDateStr) {
  const orderDateIdx = header.indexOf('OrderDate');
  const employeeCDIdx = header.indexOf('EmployeeCD');
  const activeFlgIdx = header.indexOf('ActiveFlg');
  return data.findIndex((row, idx) => {
    if (idx === 0) return false;
    return row[employeeCDIdx] == employeeCD &&
      Utilities.formatDate(new Date(row[orderDateIdx]), Session.getScriptTimeZone(), 'yyyy/MM/dd') === orderDateStr &&
      row[activeFlgIdx] == 1;
  });
}

// 工場のみ変更判定
function isFactoryOnlyChange(order, existingRow) {
  return (!order.MenuCD || !order.VenderCD) && order.FactoryCD && existingRow >= 0;
}

// 注文行作成
function createOrderRow(order, orderDate, employeeCD, now, master) {
  const orderNo = generateOrderNo();
  return [
    orderNo.toString(),
    orderDate,
    employeeCD.padStart(6, '0'),
    (order.FactoryCD ?? '').toString(),
    (order.MenuCD ?? '').toString(),
    (order.VenderCD ?? '').toString(),
    1,
    now
  ];
}

// OrderDateをDate型に変換
function parseOrderDate(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return value;
  }

  if (typeof value === 'string') {
    if (/^\d{8}$/.test(value)) {
      const y = Number(value.slice(0, 4));
      const m = Number(value.slice(4, 6));
      const d = Number(value.slice(6, 8));
      return new Date(y, m - 1, d);
    } else if (value.includes('-')) {
      const parts = value.split('-');
      if (parts.length === 3) {
        const [y, m, d] = parts.map(Number);
        return new Date(y, m - 1, d);
      }
    }
  }

  throw new Error(`無効な日付形式: ${value}`);
}

// オーダ番号発番
function generateOrderNo() {
  const sheet = SpreadsheetApp.openById(CONFIG.ORDER_ID).getSheetByName(CONFIG.ORDER_SHEET);
  const lastRow = sheet.getLastRow();

  let nextNumber = 1;
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const numbers = data
      .filter(orderNo => orderNo)
      .map(orderNo => Number(orderNo))
      .filter(n => !isNaN(n));
    if (numbers.length > 0) {
      nextNumber = Math.max(...numbers) + 1;
    }
  }

  return String(nextNumber).padStart(5, '0');
}
