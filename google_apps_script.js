/**
 * 便當訂購與收款管理系統 - Google Apps Script 雲端後端 (Code.gs)
 * 
 * 【部署步驟】
 * 1. 前往 Google 雲端硬碟 -> 新增「Google 試算表」，命名為「便當訂購系統資料庫」
 * 2. 點選上方選單「擴充功能」 -> 「Apps Script」
 * 3. 將本檔案所有程式碼複製並貼上（覆蓋原本的 Code.gs）
 * 4. 點選右上角「部署」 -> 「新增部署作業」
 * 5. 左側齒輪「選取類型」選擇「網頁應用程式 (Web App)」
 * 6. 「執行身分」選擇：「我 (您的帳號)」
 * 7. 「誰可以存取」選擇：「所有人 (Anyone)」
 * 8. 點選「部署」，授權後複製產生的「網頁應用程式網址 (https://script.google.com/macros/s/.../exec)」
 * 9. 回到便當系統管理後台 -> 貼上該網址即可全公司連線同步！
 */

function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let configSheet = ss.getSheetByName('Config');
  if (!configSheet) {
    configSheet = ss.insertSheet('Config');
    configSheet.appendRow(['Key', 'Value']);
    configSheet.appendRow(['session', JSON.stringify({
      date: Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd'),
      activeStoreIds: ['store_1'],
      deadline: '10:00',
      isClosed: false,
      announcement: ''
    })]);
    configSheet.appendRow(['customStores', '[]']);
  }

  let ordersSheet = ss.getSheetByName('Orders');
  if (!ordersSheet) {
    ordersSheet = ss.insertSheet('Orders');
    ordersSheet.appendRow(['id', 'date', 'userName', 'storeId', 'storeName', 'itemsJson', 'totalPrice', 'note', 'paid', 'createdAt']);
  }
  return { configSheet, ordersSheet };
}

function getConfigMap(configSheet) {
  const data = configSheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    map[data[i][0]] = data[i][1];
  }
  return map;
}

function setConfigValue(configSheet, key, valueStr) {
  const data = configSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      configSheet.getRange(i + 1, 2).setValue(valueStr);
      return;
    }
  }
  configSheet.appendRow([key, valueStr]);
}

function getOrdersByDate(ordersSheet, targetDate) {
  const data = ordersSheet.getDataRange().getValues();
  const orders = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const rowDate = String(row[1]);
    if (!targetDate || rowDate === targetDate) {
      let items = [];
      try { items = JSON.parse(row[5] || '[]'); } catch (e) {}
      orders.push({
        id: String(row[0]),
        date: rowDate,
        userName: String(row[2]),
        storeId: String(row[3]),
        storeName: String(row[4]),
        items: items,
        totalPrice: Number(row[6]) || 0,
        note: String(row[7] || ''),
        paid: row[8] === true || row[8] === 'TRUE' || row[8] === 'true',
        createdAt: String(row[9] || '')
      });
    }
  }
  return orders;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const { configSheet, ordersSheet } = initSheets();
    const configMap = getConfigMap(configSheet);
    let session = {};
    let customStores = [];
    try { session = JSON.parse(configMap['session'] || '{}'); } catch (err) {}
    try { customStores = JSON.parse(configMap['customStores'] || '[]'); } catch (err) {}
    
    const dateParam = (e && e.parameter && e.parameter.date) ? e.parameter.date : session.date;
    const orders = getOrdersByDate(ordersSheet, dateParam);

    return jsonResponse({
      ok: true,
      session: session,
      customStores: customStores,
      orders: orders
    });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const { configSheet, ordersSheet } = initSheets();
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action;

    if (action === 'saveSession') {
      setConfigValue(configSheet, 'session', JSON.stringify(body.session));
      return jsonResponse({ ok: true });
    }

    if (action === 'saveCustomStores') {
      setConfigValue(configSheet, 'customStores', JSON.stringify(body.customStores || []));
      return jsonResponse({ ok: true });
    }

    if (action === 'submitOrder') {
      const o = body.order;
      ordersSheet.appendRow([
        o.id,
        o.date,
        o.userName,
        o.storeId,
        o.storeName,
        JSON.stringify(o.items || []),
        o.totalPrice,
        o.note || '',
        o.paid ? true : false,
        o.createdAt || new Date().toISOString()
      ]);
      return jsonResponse({ ok: true });
    }

    if (action === 'togglePaid') {
      const orderId = body.orderId;
      const paid = body.paid ? true : false;
      const data = ordersSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(orderId)) {
          ordersSheet.getRange(i + 1, 9).setValue(paid);
          break;
        }
      }
      return jsonResponse({ ok: true });
    }

    if (action === 'deleteOrder') {
      const orderId = body.orderId;
      const data = ordersSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(orderId)) {
          ordersSheet.deleteRow(i + 1);
          break;
        }
      }
      return jsonResponse({ ok: true });
    }

    if (action === 'clearDateOrders') {
      const targetDate = body.date;
      const data = ordersSheet.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][1]) === String(targetDate)) {
          ordersSheet.deleteRow(i + 1);
        }
      }
      return jsonResponse({ ok: true });
    }

    return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  } finally {
    lock.releaseLock();
  }
}
