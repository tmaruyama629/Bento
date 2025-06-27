/**
 * @fileoverview お弁当注文システムのサーバーサイドロジックです。
 * Webアプリのバックエンド処理、スプレッドシートとのデータ連携などを担当します。
 * @author T.Maruyama
 * @since 2025-05-08
 * @version 1.2.0
 */

/**
 * =================================================================================
 * 変更履歴
 * =================================================================================
 *2025-06-27 T.Maruyama v1.2.0
 *- [機能追加] メニュー選択時に当日は選択不可とするメニューを設定する機能を追加 
 *
 * 2025-06-27 T.Maruyama v1.1.9
 * - [機能追加] 工場マスタから有効な工場のみを取得するフィルタを追加
 *
 * 2025-06-27 T.Maruyama v1.1.8
 * - [機能追加] 休日マスタから有効な休日のみを取得するフィルタを追加
 *
 * 2025-06-27 T.Maruyama v1.1.7
 * - [機能追加] メニューマスタから有効なメニューのみを取得するフィルタを追加
 * 
 * 2025-06-27 T.Maruyama v1.1.6
 * - [バグ修正] 管理者ログイン時でも休日はメニュー/工場プルダウン非活性・注文登録不可となるようサーバー・フロント両方の判定式を修正
 *
 * 2025-06-25 T.Maruyama v1.1.5
 * - [機能追加] 社員マスタから有効な社員のみを取得するフィルタを追加
 *
 * 2025-06-25 T.Maruyama v1.1.1
 * - [改善] エラーハンドリング強化・APIレスポンス形式統一
 * 
 * 2025-06-23 T.Maruyama v1.1.0
 * - [機能変更] ログイン認証を社員CDとパスワードの組み合わせで行うように変更
 * - [機能追加] ログイン画面の入力値制限（社員CD:半角数字6桁, PW:半角英数記号）
 * - [機能追加] パスワードの表示/非表示切り替え機能を追加
 * 
 * 2025-06-19 T.Maruyama v1.0.2
 * - [修正] LockServiceを導入し、注文保存処理の同時実行を防止 (saveOrderData, generateOrderNo)
 * 
 * 2025-06-10 T.Maruyama v1.0.1
 * - [リファクタリング] getMenuForWeek, saveOrderData の処理を改善
 * - [追加] 開発用のスプレッドシートIDをCONFIGに追加
 *
 * 2025-05-08 T.Maruyama v1.0.0
 * - 新規作成
 * =================================================================================
 */

function doGet() {
  try {
    const template = HtmlService.createTemplateFromFile('index');
    // return template.evaluate().setTitle('お弁当注文フォーム');
    return template.evaluate().setTitle('お弁当注文フォーム【開発系】');    
  } catch (error) {
    Logger.log('doGet error: ' + error.message + '\n' + error.stack);
    throw new Error('システムエラーが発生しました。');
  }
}

const CONFIG = {
  // ope
  // MASTER_ID: '1s00XO8VNkN4NMi1OSbRhquqU-H8gj-Lg_bQbVLKwfP0',
  //ORDER_ID: '1Mzj9Oxz3NWVmvYebdw3Bne9HrJ-v_-0nJo9TZ9xM2pI',

  // dev
  MASTER_ID: '10LOh1EtyaavGZWWPRctqFjk0WTRZZwP12llpU2efz-E',
  ORDER_ID: '1S5tGMNR6RRJfPEDp_eiIafAGtwhQlZZheMPaazrFjI4',

  ORDER_SHEET: 'D_Order'
};

// 汎用設定取得
function getConfigValue(key) {
  try {
    if (!key) throw new Error('keyが未指定です');
    const spreadsheet = SpreadsheetApp.openById(CONFIG.MASTER_ID);
    const configSheet = spreadsheet.getSheetByName('M_Config');
    if (!configSheet) throw new Error('M_Configシートが見つかりません');
    const configValues = configSheet.getDataRange().getValues();
    for (let i = 0; i < configValues.length; i++) {
      if (configValues[i][0] === key) {
        return configValues[i][1];
      }
    }
    return null;
  } catch (error) {
    Logger.log('getConfigValue error: ' + error.message + '\n' + error.stack);
    throw error;
  }
}

// 共通のスプレッドシート取得関数
function getSpreadsheet(sheetId) {
  try {
    if (!sheetId) throw new Error('sheetIdが未指定です');
    return SpreadsheetApp.openById(sheetId);
  } catch (error) {
    Logger.log('getSpreadsheet error: ' + error.message + '\n' + error.stack);
    throw new Error(`スプレッドシートの取得に失敗しました: ${sheetId}`);
  }
}

// データ取得共通化
function getDataFromSheet(sheetId, sheetName) {
  try {
    if (!sheetId || !sheetName) throw new Error('sheetIdまたはsheetNameが未指定です');
    const sheet = getSpreadsheet(sheetId).getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`シート ${sheetName} が見つかりません。`);
    }
    const values = sheet.getDataRange().getValues();
    return values.slice(1);
  } catch (error) {
    Logger.log('getDataFromSheet error: ' + error.message + '\n' + error.stack);
    throw error;
  }
}

// 社員一覧取得
function getEmployeeList() {
  try {
    const values = getDataFromSheet(CONFIG.MASTER_ID, 'M_Employee');
    // ActiveFlg=1のみ返す
    return { status: 'success', data: values.filter(row => row[4] === 1).map(row => ({
      EmployeeCD: row[0],
      EmployeeName: row[1]
    })) };
  } catch (error) {
    Logger.log('getEmployeeList error: ' + error.message + '\n' + error.stack);
    return { status: 'error', message: '社員一覧の取得に失敗しました。' };
  }
}

// 工場一覧取得
function getFactoryList() {
  try {
    const values = getDataFromSheet(CONFIG.MASTER_ID, 'M_Factory');
    // ActiveFlg=1のみ返す
    return { status: 'success', data: values.filter(row => row[2] === 1).map(row => ({
      FactoryCD: row[0],
      FactoryName: row[1]
    })) };
  } catch (error) {
    Logger.log('getFactoryList error: ' + error.message + '\n' + error.stack);
    return { status: 'error', message: '工場一覧の取得に失敗しました。' };
  }
}

// 社員のデフォルト工場を取得
function getEmployeeData(empCD) {
  try {
    if (!empCD) throw new Error('社員CDが未指定です');
    const values = getDataFromSheet(CONFIG.MASTER_ID, 'M_Employee');
    // ActiveFlg=1のみ有効
    const employee = values.find(row => row[0] === empCD && row[4] === 1);
    return { status: 'success', data: { defaultFactory: employee ? employee[2] : '' } };
  } catch (error) {
    Logger.log('getEmployeeData error: ' + error.message + '\n' + error.stack);
    return { status: 'error', message: '社員データの取得に失敗しました。' };
  }
}

// 休日取得
function getHolidayMap() {
  try {
    // ヘッダー＋データ全体を取得
    const sheet = getSpreadsheet(CONFIG.MASTER_ID).getSheetByName('M_Holiday');
    if (!sheet) throw new Error('M_Holidayシートが見つかりません');
    const values = sheet.getDataRange().getValues();
    const header = values[0];
    const dateIdx = header.indexOf('HolidayDate');
    const activeFlgIdx = header.indexOf('ActiveFlg');
    const holidayMap = {};
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      // ActiveFlg=1のみ有効な休日として扱う
      if (row[activeFlgIdx] !== 1) continue;
      const date = row[dateIdx];
      if (date) {
        const formatted = formatDate(date);
        holidayMap[formatted] = true;
      }
    }
    return holidayMap;
  } catch (error) {
    Logger.log('getHolidayMap error: ' + error.message + '\n' + error.stack);
    throw error;
  }
}

// コメント取得
function getComment() {
  try {
    const values = getDataFromSheet(CONFIG.MASTER_ID, 'M_Comment');
    return { status: 'success', data: values.map(row => ({
      CommentCD: row[0],
      CommentText: row[1],
      HyperLink: row[2]
    })) };
  } catch (error) {
    Logger.log('getComment error: ' + error.message + '\n' + error.stack);
    return { status: 'error', message: 'コメントの取得に失敗しました。' };
  }
}

// メニュー制限マップ取得
function getMenuRestrictionMap() {
  try {
    const sheet = getSpreadsheet(CONFIG.MASTER_ID).getSheetByName('M_MenuRestriction');
    if (!sheet) return {};
    const values = sheet.getDataRange().getValues();
    const header = values[0];
    const dateIdx = header.indexOf('MenuRestrictionDate');
    const menuCDIdx = header.indexOf('MenuCD');
    const venderCDIdx = header.indexOf('VenderCD');
    const activeFlgIdx = header.indexOf('ActiveFlg');
    const map = {};
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (row[activeFlgIdx] !== 1) continue;
      const date = formatDate(row[dateIdx]);
      const menuCD = row[menuCDIdx];
      const venderCD = row[venderCDIdx];
      if (!date || !menuCD || !venderCD) continue;
      if (!map[date]) map[date] = {};
      map[date][`${menuCD}_${venderCD}`] = true;
    }
    return map;
  } catch (error) {
    Logger.log('getMenuRestrictionMap error: ' + error.message + '\n' + error.stack);
    return {};
  }
}

// 指定週のメニューを取得
function getMenuForWeek(empCD, startDate, endDate, isAdmin) {
  try {
    if (!empCD || !startDate || !endDate) throw new Error('パラメータが不足しています');
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
    const restrictionMap = getMenuRestrictionMap(); // 追加

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
      const isClosed = isHoliday || isPastDeadline;

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

      // 選択不可メニューリスト
      const restrictedMenus = restrictionMap[date] ? Object.keys(restrictionMap[date]) : [];

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
        Menus: menuDetails,
        RestrictedMenus: restrictedMenus // 追加
      };
    });

    return { status: 'success', data: weekData };
  } catch (error) {
    Logger.log('getMenuForWeek error: ' + error.message + '\n' + error.stack);
    return { status: 'error', message: 'メニュー情報の取得に失敗しました。' };
  }
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

// メニュー抽出
function filterMenusForDate(date, menuValues) {
  const jsDate = new Date(date);
  const dow = jsDate.getDay(); // 0:Sun ～ 6:Sat
  const colIndex = 3 + dow; // M_Menuシートの曜日列は3列目から
  // ActiveFlg=1のみ返す
  return menuValues.slice(1).filter(row => row[colIndex] === 1 && row[11] === 1);
}

// メニュー詳細リスト生成
function getMenuDetailsForDate(date, menuValues, venderMap, isAdmin) {
  const menus = filterMenusForDate(date, menuValues);
  return menus
    .filter(m => isAdmin || m[10] === 1) // さらにDispFlg=1のみ（従来通り）
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

// 日付フォーマット
function formatDate(value) {
  if (typeof value === 'number') value = value.toString();
  if (/^\d{8}$/.test(value)) {
    return `${value.substring(0,4)}-${value.substring(4,6)}-${value.substring(6,8)}`;
  }

  const date = new Date(value);
  if (isNaN(date.getTime())) { // Invalid Date かどうかをチェック
    Logger.log(`formatDate に無効な日付値が渡されました: ${value}`);
    // 動作を決定: エラーをスローするか、null を返すか、デフォルトのエラー文字列を返す
    throw new Error(`無効な日付形式です: ${value}`);
    // return null;
  }
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

// ログイン認証
function verifyLogin(empCD, inputPw) {
  try {
    if (!empCD || !inputPw) throw new Error('社員CDまたはパスワードが未入力です');
    if (empCD === '999999') {
      const adminPw = getConfigValue('AdminPassword');
      if (inputPw === adminPw) {
        return { status: 'success', data: { isValid: true, isAdmin: true, employeeCD: empCD } };
      }
    } else {
      const sheet = getSpreadsheet(CONFIG.MASTER_ID).getSheetByName('M_Employee');
      if (!sheet) throw new Error('M_Employee シートが見つかりません。');
      const values = sheet.getDataRange().getValues();
      // ActiveFlg=1のみ有効
      const employee = values.find(row => row[0] === empCD && row[3] === inputPw && row[4] === 1);
      if (employee) {
        return { status: 'success', data: { isValid: true, isAdmin: false, employeeCD: empCD, employeeName: employee[1] } };
      }
    }
    return { status: 'success', data: { isValid: false, isAdmin: false } };
  } catch (error) {
    Logger.log('verifyLogin error: ' + error.message + '\n' + error.stack);
    return { status: 'error', message: '認証処理でエラーが発生しました。' };
  }
}

// 注文保存処理
function saveOrderData(employeeCD, orders, isAdmin) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30);
  try {
    if (!employeeCD || !Array.isArray(orders)) throw new Error('パラメータが不正です');
    const holidayMap = getHolidayMap();
    const dbSheet = SpreadsheetApp.openById(CONFIG.ORDER_ID).getSheetByName(CONFIG.ORDER_SHEET);
    const now = new Date();
    const insertedOrders = [];
    const data = dbSheet.getDataRange().getValues();
    const header = data[0];

    // マスタデータ取得
    const master = getMasterData();

    orders.forEach(order => {
      try { // 追加: 各注文処理毎のエラーハンドリング
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
      } catch (error) {
        Logger.log(`注文処理でエラーが発生しました (Employee: ${employeeCD}, OrderDate: ${order.OrderDate}): ${error.message}`);
      }
    });

    Logger.log('insertedOrders: ' + JSON.stringify(insertedOrders));
    return { status: 'success', data: insertedOrders };
  } catch (error) {
    Logger.log('saveOrderData error: ' + error.message + '\n' + error.stack);
    return { status: 'error', message: '注文保存処理でエラーが発生しました。' };
  } finally {
    lock.releaseLock();
  }
}

// マスタデータ取得
// 修正: 各マスタシートが取得できなかった場合にエラーをスローするチェックを追加
function getMasterData() {
  const menuMasterSheet = SpreadsheetApp.openById(CONFIG.MASTER_ID).getSheetByName('M_Menu');
  if (!menuMasterSheet) {
    throw new Error('M_Menu シートが見つかりません。');
  }
  const menuData = menuMasterSheet.getDataRange().getValues();
  const menuHeader = menuData[0];
  const menuCDIdxInMaster = menuHeader.indexOf('MenuCD');
  const venderCDIdxInMaster = menuHeader.indexOf('VenderCD');
  const menuNameIdx = menuHeader.indexOf('MenuName');

  const venderSheet = SpreadsheetApp.openById(CONFIG.MASTER_ID).getSheetByName('M_Vender');
  if (!venderSheet) {
    throw new Error('M_Vender シートが見つかりません。');
  }
  const venderData = venderSheet.getDataRange().getValues();
  const venderHeader = venderData[0];
  const venderCDIdxInVender = venderHeader.indexOf('VenderCD');
  const venderNameIdx = venderHeader.indexOf('VenderName');

  const configSheet = SpreadsheetApp.openById(CONFIG.MASTER_ID).getSheetByName('M_Config');
  if (!configSheet) {
    throw new Error('M_Config シートが見つかりません。');
  }
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
  if (!dbSheet) {
    throw new Error(`注文シート ${CONFIG.ORDER_SHEET} が見つかりません。`);
  }
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
  const isHoliday = !!holidayMap[orderDateStr];
  const isPastDate = orderDateStr < nowDateStr;
  const deadline = new Date(orderDate);
  deadline.setHours(deadlineHour, deadlineMinute, 0, 0);
  const isPastDeadline = now > deadline && orderDateStr === nowDateStr;
  // 管理者でも休日は常に登録不可
  return isHoliday || (!isAdmin && (isPastDeadline || isPastDate));
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
  const lock = LockService.getScriptLock();
  lock.waitLock(30); // 30秒
  try {
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
  } finally {
    lock.releaseLock();
  }
}
