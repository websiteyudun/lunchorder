/**
 * 便當訂購與收款管理系統 - 核心邏輯與 Excel 智慧解析引擎 (app.js)
 */

(function () {
  // ===== 預設狀態與常數 =====
  const STORAGE_KEYS = {
    STORES: 'bento_stores_v1',
    SESSION: 'bento_session_v1',
    ORDERS: 'bento_orders_v1',
    CLOUD_CONFIG: 'bento_cloud_config_v1',
    LAST_USER: 'bento_last_username_v1'
  };

  function getTodayString() {
    const d = new Date();
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  }

  // ===== 狀態管理 =====
  const state = {
    mode: 'order', // 'order' | 'admin'
    adminTab: 'dashboard', // 'dashboard' | 'session' | 'menus' | 'cloud'
    stores: [],
    session: {
      date: getTodayString(),
      activeStoreIds: ['store_1'],
      deadline: '10:00',
      isClosed: false,
      announcement: ''
    },
    orders: [],
    selectedStoreId: 'store_1',
    selectedCategory: 'ALL',
    menuSearch: '',
    cart: [], // [{ storeId, storeName, name, price, qty, note }]
    userName: localStorage.getItem(STORAGE_KEYS.LAST_USER) || '',
    orderNote: '',
    cloudConfig: {
      type: 'local', // 'local' | 'gas' | 'firebase'
      gasUrl: '',
      firebaseUrl: ''
    },
    isSyncing: false,
    editingStoreId: null
  };

  // ===== 1. Excel 智慧解析引擎 (支援 SheetJS 2D 陣列) =====
  function cleanText(s) {
    if (s === null || s === undefined) return '';
    let t = String(s).trim().replace(/　/g, ' ');
    // 合併中文字之間的空白，例如 "排 骨 飯" -> "排骨飯"
    let prev;
    do {
      prev = t;
      t = t.replace(/([\u4e00-\u9fa5])\s+([\u4e00-\u9fa5])/g, '$1$2');
    } while (t !== prev);
    return t.replace(/\s+/g, ' ').trim();
  }

  function parsePrice(s) {
    if (s === null || s === undefined) return null;
    const t = String(s).trim().replace(/　/g, '');
    if (/^\d+$/.test(t)) {
      const v = parseInt(t, 10);
      if (v >= 5 && v <= 1000) return { price: v, prefix: '' };
    }
    const m = t.match(/^([大小])\s*(\d+)$/);
    if (m) {
      const v = parseInt(m[2], 10);
      if (v >= 5 && v <= 1000) return { price: v, prefix: `(${m[1]})` };
    }
    return null;
  }

  function isHeaderOrIgnored(s) {
    const t = cleanText(s);
    if (!t) return true;
    if (/^(品項|品名|便當|單價|價格|價錢|金額|姓名|簽名)$/.test(t)) return true;
    if (/外送|電話|截止|之前訂|前訂購|點前|休息|公休|更新|漲$/.test(t)) return true;
    if (/^09\d{2}|^\(03\)|^03-|^\d{7,8}$/.test(t)) return true;
    if (/蘆竹|南崁|中正路|南山路|南竹路|光明路|仁愛路|大新一街|南上路|蘆興南路/.test(t)) return true;
    return false;
  }

  function isCategoryHeader(s) {
    const t = cleanText(s);
    return /^(單點|飯類|麵類|湯麵|炒麵|炒飯|燴飯|湯類|小菜|小菜類|乾麵類.*|湯麵類.*|捲壽司|豆皮|精緻小菜|焗麵|焗飯|酥皮包|濃湯|其他|日式拉麵|韓式石鍋拌飯.*|日式醬燒石鍋拌飯.*|韓式料理.*|起司豆腐煲.*|單點\(.*\)|品項\(單點\))$/.test(t);
  }

  function parseWorkbookToStores(workbook) {
    const parsedStores = [];
    workbook.SheetNames.forEach((sheetName) => {
      if (sheetName === '好記涼麵 (2)') return;
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      if (!rows || rows.length === 0) return;

      const infoLines = [];
      const items = [];
      const notes = [];

      for (let rIdx = 0; rIdx < Math.min(3, rows.length); rIdx++) {
        const row = rows[rIdx] || [];
        row.forEach((cell) => {
          const cText = cleanText(cell);
          if (!cText) return;
          if (/^(品項|品名|便當|單價|價格|價錢|金額|姓名|簽名|飯)$/.test(cText)) return;
          if (rIdx === 0 && infoLines.length === 0 && !/^\d/.test(cText)) {
            infoLines.push(cText);
          } else if (/電話|外送|03-|09\d{2}|\d{3}-\d{4}|\d{7}|前訂|截止|休|路|街|號|元/.test(cText)) {
            if (!infoLines.includes(cText)) infoLines.push(cText);
          }
        });
      }

      const colState = {};

      for (let rIdx = 0; rIdx < rows.length; rIdx++) {
        const row = rows[rIdx] || [];
        const nonEmptyCells = [];
        for (let c = 0; c < row.length; c++) {
          const val = cleanText(row[c]);
          if (val) nonEmptyCells.push({ col: c, val });
        }
        if (nonEmptyCells.length === 0) continue;

        if (/^[\*＊]/.test(nonEmptyCells[0].val)) {
          notes.push(nonEmptyCells[0].val);
          continue;
        }

        const rowJoined = nonEmptyCells.map((x) => x.val).join(' ');
        if (/姓名|簽名/.test(rowJoined) && /品項|品名|便當|單價|價格|價錢|金額|炒飯/.test(rowJoined)) {
          if (nonEmptyCells[0].val === '炒飯') {
            colState[0] = { category: '炒飯', lastMainItem: '', lastPrice: 0 };
          }
          continue;
        }
        if (rIdx === 0 && (/電話|外送|\d{3}-\d{4}|\d{7}/.test(rowJoined) || nonEmptyCells.length <= 3)) {
          continue;
        }
        if (rIdx === 1 && /外送電話/.test(rowJoined)) {
          continue;
        }

        let c = 0;
        while (c < row.length) {
          const c0 = cleanText(row[c]);
          if (!c0) {
            c++;
            continue;
          }

          if (!colState[c]) {
            colState[c] = { category: '餐點', lastMainItem: '', lastPrice: 0 };
          }
          const st = colState[c];

          if (isHeaderOrIgnored(c0)) {
            c++;
            continue;
          }

          const c1 = c + 1 < row.length ? cleanText(row[c + 1]) : '';
          const c2 = c + 2 < row.length ? cleanText(row[c + 2]) : '';
          const p1 = parsePrice(c1);
          const p2 = parsePrice(c2);

          if (isCategoryHeader(c0) && !p1 && !p2) {
            st.category = c0;
            st.lastMainItem = '';
            c++;
            continue;
          }

          // 優先檢查子選項 (前一欄為空且有上一筆主品項)
          if (c > 0 && !cleanText(row[c - 1])) {
            const prevCol = c - 1;
            if (colState[prevCol] && colState[prevCol].lastMainItem) {
              const mainItem = colState[prevCol].lastMainItem;
              const cat = colState[prevCol].category;
              const p0 = parsePrice(c0);
              if (p0 && p0.prefix) {
                items.push({ category: cat, name: `${mainItem}${p0.prefix}`, price: p0.price });
                c++;
                continue;
              }
              if (p1) {
                colState[prevCol].lastPrice = p1.price;
                items.push({ category: cat, name: `${mainItem} - ${c0}`, price: p1.price });
                c += 2;
                continue;
              } else if (colState[prevCol].lastPrice > 0 && c0.length <= 6 && /麵|粉|條|飯|五穀|醬/.test(c0)) {
                items.push({ category: cat, name: `${mainItem} - ${c0}`, price: colState[prevCol].lastPrice });
                c++;
                continue;
              }
            }
          }

          // Case 1: c0 是品名, c1 是價格
          if (p1) {
            let itemName = c0 + p1.prefix;
            let cat = st.category;
            if (['焗麵', '焗飯'].includes(cat) && /湯|酥皮包|咖啡|紅茶|沙拉|麵包/.test(c0)) {
              cat = '湯品與點心';
              st.category = cat;
            }
            if (['炒飯', '燴飯', '湯麵', '炒麵'].includes(cat)) {
              if (!itemName.includes(cat) && !/飯|麵/.test(itemName)) {
                itemName = `${itemName}${cat}`;
              }
            } else if (['焗麵', '焗飯'].includes(cat)) {
              if (!itemName.includes('焗')) {
                itemName = `${itemName} (${cat})`;
              }
            }
            st.lastMainItem = c0;
            st.lastPrice = p1.price;
            items.push({ category: cat, name: itemName, price: p1.price });
            c += 2;
            continue;
          }

          // Case 2: c0 是品名, c1 是子項或空, c2 是價格
          if (p2) {
            const cat = st.category;
            if (c1 && !isHeaderOrIgnored(c1)) {
              st.lastMainItem = c0;
              st.lastPrice = p2.price;
              const itemName = `${c0} - ${c1}${p2.prefix}`;
              items.push({ category: cat, name: itemName, price: p2.price });
            } else {
              let itemName = c0 + p2.prefix;
              if (['焗麵', '焗飯'].includes(cat) && !itemName.includes('焗')) {
                itemName = `${itemName} (${cat})`;
              }
              st.lastMainItem = c0;
              st.lastPrice = p2.price;
              items.push({ category: cat, name: itemName, price: p2.price });
            }
            c += 3;
            continue;
          }

          // Case 5: 梁社漢那種飯跟麵同價但麵沒填價格
          if (/麵$/.test(c0) && /飯$/.test(st.lastMainItem) && st.lastPrice > 0) {
            const base1 = c0.replace(/麵$/, '');
            const base2 = st.lastMainItem.replace(/飯$/, '');
            if (base1 === base2) {
              items.push({ category: st.category, name: c0, price: st.lastPrice });
              c++;
              continue;
            }
          }

          c++;
        }
      }

      if (items.length === 0) return;

      let cleanName = sheetName
        .replace(/\(.*?\)/g, '')
        .replace(/\d+\.\d+(\.\d+)?/g, '')
        .replace(/已更新|新/g, '')
        .trim();
      if (sheetName === '陳老師') cleanName = '陳媽媽(陳老師)';

      parsedStores.push({
        id: 'store_' + (parsedStores.length + 1),
        sheetName,
        name: cleanName || sheetName,
        info: infoLines.join(' | '),
        notes: notes.join(' / '),
        itemCount: items.length,
        items
      });
    });

    return parsedStores;
  }

  // ===== 2. 初始化與網址參數解析 =====
  function initApp() {
    // 1. 讀取雲端設定
    try {
      const savedCloud = JSON.parse(localStorage.getItem(STORAGE_KEYS.CLOUD_CONFIG) || 'null');
      if (savedCloud) state.cloudConfig = { ...state.cloudConfig, ...savedCloud };
    } catch (e) {}

    // 2. 讀取店家資料 (優先使用 localStorage 更新過的，否則使用內建 DEFAULT_STORES)
    try {
      const savedStores = JSON.parse(localStorage.getItem(STORAGE_KEYS.STORES) || 'null');
      if (savedStores && Array.isArray(savedStores) && savedStores.length > 0) {
        state.stores = savedStores;
      } else if (window.DEFAULT_STORES) {
        state.stores = JSON.parse(JSON.stringify(window.DEFAULT_STORES));
      }
    } catch (e) {
      if (window.DEFAULT_STORES) state.stores = JSON.parse(JSON.stringify(window.DEFAULT_STORES));
    }

    // 3. 讀取今日開團設定
    try {
      const savedSession = JSON.parse(localStorage.getItem(STORAGE_KEYS.SESSION) || 'null');
      if (savedSession && savedSession.activeStoreIds) {
        state.session = { ...state.session, ...savedSession };
      }
    } catch (e) {}

    // 4. 讀取本地訂單
    try {
      const savedOrders = JSON.parse(localStorage.getItem(STORAGE_KEYS.ORDERS) || '[]');
      if (Array.isArray(savedOrders)) state.orders = savedOrders;
    } catch (e) {}

    // 5. 解析網址參數 (讓訂購者點開連結就能自動套用雲端網址、日期、指定店家與截止時間！)
    const params = new URLSearchParams(window.location.search);
    if (params.get('gas')) {
      state.cloudConfig.type = 'gas';
      state.cloudConfig.gasUrl = params.get('gas');
      localStorage.setItem(STORAGE_KEYS.CLOUD_CONFIG, JSON.stringify(state.cloudConfig));
    }
    if (params.get('fb')) {
      state.cloudConfig.type = 'firebase';
      state.cloudConfig.firebaseUrl = params.get('fb');
      localStorage.setItem(STORAGE_KEYS.CLOUD_CONFIG, JSON.stringify(state.cloudConfig));
    }
    if (params.get('date')) {
      state.session.date = params.get('date');
    }
    if (params.get('stores')) {
      const storeIds = params.get('stores').split(',').filter(Boolean);
      if (storeIds.length > 0) {
        state.session.activeStoreIds = storeIds;
      }
    }
    if (params.get('deadline')) {
      state.session.deadline = params.get('deadline');
    }
    if (params.get('note')) {
      state.session.announcement = params.get('note');
    }
    if (params.get('admin') === '1') {
      state.mode = 'admin';
    }

    ensureValidSelectedStore();
    renderAll();

    // 若有設定雲端同步，立即從雲端拉取最新開團狀態與訂單
    if (state.cloudConfig.type !== 'local') {
      syncFromCloud();
      // 每 20 秒自動背景更新訂單
      setInterval(() => {
        if (!state.isSyncing) syncFromCloud(true);
      }, 20000);
    }

    // 跨分頁即時同步
    window.addEventListener('storage', (e) => {
      if (e.key === STORAGE_KEYS.ORDERS || e.key === STORAGE_KEYS.SESSION || e.key === STORAGE_KEYS.STORES) {
        reloadFromLocalStorage();
        renderAll();
      }
    });
  }

  function reloadFromLocalStorage() {
    try {
      const s = JSON.parse(localStorage.getItem(STORAGE_KEYS.SESSION) || 'null');
      if (s) state.session = s;
      const o = JSON.parse(localStorage.getItem(STORAGE_KEYS.ORDERS) || '[]');
      if (o) state.orders = o;
      const st = JSON.parse(localStorage.getItem(STORAGE_KEYS.STORES) || 'null');
      if (st && st.length) state.stores = st;
    } catch (e) {}
  }

  function ensureValidSelectedStore() {
    const activeIds = state.session.activeStoreIds || [];
    if (activeIds.length > 0 && !activeIds.includes(state.selectedStoreId)) {
      state.selectedStoreId = activeIds[0];
    } else if (!state.selectedStoreId && state.stores.length > 0) {
      state.selectedStoreId = state.stores[0].id;
    }
  }

  // ===== 3. 雲端與本地資料同步層 =====
  function saveLocalState() {
    localStorage.setItem(STORAGE_KEYS.SESSION, JSON.stringify(state.session));
    localStorage.setItem(STORAGE_KEYS.ORDERS, JSON.stringify(state.orders));
    localStorage.setItem(STORAGE_KEYS.STORES, JSON.stringify(state.stores));
    localStorage.setItem(STORAGE_KEYS.CLOUD_CONFIG, JSON.stringify(state.cloudConfig));
  }

  async function syncFromCloud(silent = false) {
    if (state.cloudConfig.type === 'local') return;
    if (!silent) setSyncBadge('同步中...', 'syncing');
    state.isSyncing = true;

    try {
      if (state.cloudConfig.type === 'gas' && state.cloudConfig.gasUrl) {
        const url = new URL(state.cloudConfig.gasUrl);
        url.searchParams.set('date', state.session.date);
        url.searchParams.set('_t', Date.now());
        const res = await fetch(url.toString(), { method: 'GET' });
        const data = await res.json();
        if (data && data.ok) {
          if (data.session && data.session.activeStoreIds) {
            state.session = { ...state.session, ...data.session };
          }
          if (Array.isArray(data.customStores) && data.customStores.length > 0) {
            state.stores = data.customStores;
          }
          if (Array.isArray(data.orders)) {
            // 合併或取代當日訂單
            const otherOrders = state.orders.filter((o) => o.date !== state.session.date);
            state.orders = [...otherOrders, ...data.orders];
          }
          ensureValidSelectedStore();
          saveLocalState();
          renderAll();
          setSyncBadge('雲端已連線 (Google 試算表)', 'online');
        }
      } else if (state.cloudConfig.type === 'firebase' && state.cloudConfig.firebaseUrl) {
        const baseUrl = state.cloudConfig.firebaseUrl.replace(/\/$/, '');
        const res = await fetch(`${baseUrl}/bento.json?_t=${Date.now()}`);
        const data = await res.json();
        if (data) {
          if (data.session && data.session.activeStoreIds) {
            state.session = { ...state.session, ...data.session };
          }
          if (Array.isArray(data.customStores) && data.customStores.length > 0) {
            state.stores = data.customStores;
          }
          if (data.orders) {
            const cloudOrders = Object.values(data.orders);
            state.orders = cloudOrders;
          } else {
            state.orders = [];
          }
          ensureValidSelectedStore();
          saveLocalState();
          renderAll();
          setSyncBadge('雲端已連線 (Firebase)', 'online');
        }
      }
    } catch (err) {
      console.warn('Cloud sync error:', err);
      if (!silent) setSyncBadge('雲端連線失敗 (使用本機暫存)', 'error');
    } finally {
      state.isSyncing = false;
    }
  }

  async function pushToCloud(action, payload) {
    saveLocalState();
    if (state.cloudConfig.type === 'local') return true;

    setSyncBadge('儲存至雲端中...', 'syncing');
    try {
      if (state.cloudConfig.type === 'gas' && state.cloudConfig.gasUrl) {
        await fetch(state.cloudConfig.gasUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify({ action, ...payload })
        });
        setSyncBadge('雲端已同步 (Google 試算表)', 'online');
        return true;
      } else if (state.cloudConfig.type === 'firebase' && state.cloudConfig.firebaseUrl) {
        const baseUrl = state.cloudConfig.firebaseUrl.replace(/\/$/, '');
        if (action === 'saveSession') {
          await fetch(`${baseUrl}/bento/session.json`, {
            method: 'PUT',
            body: JSON.stringify(payload.session)
          });
        } else if (action === 'saveCustomStores') {
          await fetch(`${baseUrl}/bento/customStores.json`, {
            method: 'PUT',
            body: JSON.stringify(payload.customStores)
          });
        } else if (action === 'submitOrder') {
          const o = payload.order;
          await fetch(`${baseUrl}/bento/orders/${o.id}.json`, {
            method: 'PUT',
            body: JSON.stringify(o)
          });
        } else if (action === 'togglePaid') {
          await fetch(`${baseUrl}/bento/orders/${payload.orderId}/paid.json`, {
            method: 'PUT',
            body: JSON.stringify(payload.paid)
          });
        } else if (action === 'deleteOrder') {
          await fetch(`${baseUrl}/bento/orders/${payload.orderId}.json`, {
            method: 'DELETE'
          });
        } else if (action === 'clearDateOrders') {
          // 重新覆寫保留非當日訂單
          const remaining = {};
          state.orders.forEach((o) => {
            remaining[o.id] = o;
          });
          await fetch(`${baseUrl}/bento/orders.json`, {
            method: 'PUT',
            body: JSON.stringify(remaining)
          });
        }
        setSyncBadge('雲端已同步 (Firebase)', 'online');
        return true;
      }
    } catch (err) {
      console.error('Push to cloud error:', err);
      setSyncBadge('雲端寫入失敗 (已存本機)', 'error');
      return false;
    }
  }

  function setSyncBadge(text, status) {
    const el = document.getElementById('syncStatusBadge');
    if (!el) return;
    el.textContent = text;
    el.className = `sync-badge sync-${status}`;
  }

  // ===== 4. 檢查是否已截止收單 =====
  function isOrderingClosed() {
    if (state.session.isClosed) return { closed: true, reason: '管理員已手動結單關閉點餐' };
    if (state.session.deadline && state.session.date === getTodayString()) {
      const now = new Date();
      const [hh, mm] = state.session.deadline.split(':').map(Number);
      if (!isNaN(hh) && !isNaN(mm)) {
        const deadlineTime = new Date();
        deadlineTime.setHours(hh, mm, 0, 0);
        if (now > deadlineTime) {
          return { closed: true, reason: `已超過今日收單截止時間 (${state.session.deadline})` };
        }
      }
    }
    return { closed: false, reason: '' };
  }

  // ===== 5. 畫面渲染邏輯 =====
  function renderAll() {
    renderHeader();
    if (state.mode === 'order') {
      document.getElementById('orderView').classList.remove('hidden');
      document.getElementById('adminView').classList.add('hidden');
      renderOrderView();
    } else {
      document.getElementById('orderView').classList.add('hidden');
      document.getElementById('adminView').classList.remove('hidden');
      renderAdminView();
    }
  }

  function renderHeader() {
    const btnOrder = document.getElementById('navOrderBtn');
    const btnAdmin = document.getElementById('navAdminBtn');
    if (state.mode === 'order') {
      btnOrder.classList.add('active');
      btnAdmin.classList.remove('active');
    } else {
      btnOrder.classList.remove('active');
      btnAdmin.classList.add('active');
    }

    const todayOrders = getTodayOrders();
    const badge = document.getElementById('todayOrderCountBadge');
    if (badge) badge.textContent = todayOrders.length;

    if (state.cloudConfig.type === 'local') {
      setSyncBadge('單機/網址分享模式 (可於後台綁定雲端)', 'local');
    }
  }

  function getTodayOrders() {
    return state.orders.filter((o) => o.date === state.session.date);
  }

  // ----- 訂購前台渲染 -----
  function renderOrderView() {
    const activeIds = state.session.activeStoreIds || [];
    const activeStores = state.stores.filter((s) => activeIds.includes(s.id));
    const closeStatus = isOrderingClosed();

    // 頂部開團資訊列
    document.getElementById('displaySessionDate').textContent = state.session.date;
    document.getElementById('displayDeadline').textContent = state.session.deadline
      ? `今日 ${state.session.deadline} 截止`
      : '未設限';

    const statusBanner = document.getElementById('orderStatusBanner');
    if (closeStatus.closed) {
      statusBanner.className = 'status-banner closed';
      statusBanner.innerHTML = `<i class="fa-solid fa-lock"></i> <strong>目前暫停收單：</strong>${escapeHtml(closeStatus.reason)}（如需加訂請聯繫管理人）`;
    } else {
      statusBanner.className = 'status-banner open';
      const announce = state.session.announcement
        ? ` ｜ 📢 公告：<strong>${escapeHtml(state.session.announcement)}</strong>`
        : '';
      statusBanner.innerHTML = `<i class="fa-solid fa-circle-check"></i> <strong>開團訂購中！</strong>請選擇今日便當並輸入姓名送出${announce}`;
    }

    // 今日開放店家選擇器
    const storeTabsEl = document.getElementById('activeStoreTabs');
    if (activeStores.length === 0) {
      storeTabsEl.innerHTML = `<div class="empty-hint">今日尚未開放任何便當店，請由管理後台勾選今日要訂購的店家。</div>`;
      document.getElementById('menuContainer').innerHTML = '';
      return;
    }

    if (!activeStores.some((s) => s.id === state.selectedStoreId)) {
      state.selectedStoreId = activeStores[0].id;
    }

    storeTabsEl.innerHTML = activeStores
      .map((s) => {
        const isSelected = s.id === state.selectedStoreId;
        return `
          <button class="store-pill ${isSelected ? 'active' : ''}" data-store-id="${s.id}">
            <span class="store-pill-name">🍱 ${escapeHtml(s.name)}</span>
            <span class="store-pill-count">${s.items.length} 品項</span>
          </button>
        `;
      })
      .join('');

    storeTabsEl.querySelectorAll('.store-pill').forEach((btn) => {
      btn.addEventListener('click', () => {
        state.selectedStoreId = btn.getAttribute('data-store-id');
        state.selectedCategory = 'ALL';
        renderOrderView();
      });
    });

    const currentStore = state.stores.find((s) => s.id === state.selectedStoreId);
    if (!currentStore) return;

    // 店家資訊卡
    document.getElementById('currentStoreTitle').textContent = currentStore.name;
    document.getElementById('currentStoreInfo').textContent = currentStore.info || '暫無電話或外送備註';
    const noteEl = document.getElementById('currentStoreNotes');
    if (currentStore.notes) {
      noteEl.textContent = `💡 店家加料/附餐說明：${currentStore.notes}`;
      noteEl.classList.remove('hidden');
    } else {
      noteEl.classList.add('hidden');
    }

    // 分類標籤
    const categories = ['ALL', ...new Set(currentStore.items.map((i) => i.category || '餐點'))];
    const catTabsEl = document.getElementById('categoryTabs');
    if (categories.length <= 2) {
      catTabsEl.innerHTML = '';
    } else {
      catTabsEl.innerHTML = categories
        .map((c) => {
          const label = c === 'ALL' ? '全部餐點' : c;
          const active = state.selectedCategory === c ? 'active' : '';
          return `<button class="cat-chip ${active}" data-cat="${escapeHtml(c)}">${escapeHtml(label)}</button>`;
        })
        .join('');
      catTabsEl.querySelectorAll('.cat-chip').forEach((btn) => {
        btn.addEventListener('click', () => {
          state.selectedCategory = btn.getAttribute('data-cat');
          renderOrderView();
        });
      });
    }

    // 過濾品項
    const keyword = state.menuSearch.trim().toLowerCase();
    const filteredItems = currentStore.items.filter((item) => {
      if (state.selectedCategory !== 'ALL' && (item.category || '餐點') !== state.selectedCategory) {
        return false;
      }
      if (keyword && !item.name.toLowerCase().includes(keyword) && !(item.category || '').toLowerCase().includes(keyword)) {
        return false;
      }
      return true;
    });

    // 依分類分組渲染菜單
    const grouped = {};
    filteredItems.forEach((item) => {
      const cat = item.category || '餐點';
      if (!grouped[cat]) grouped[cat] = [];
      grouped[cat].push(item);
    });

    const menuContainer = document.getElementById('menuContainer');
    if (filteredItems.length === 0) {
      menuContainer.innerHTML = `<div class="empty-state">找不到符合「${escapeHtml(state.menuSearch)}」的餐點品項</div>`;
    } else {
      menuContainer.innerHTML = Object.entries(grouped)
        .map(([cat, list]) => {
          const cardsHtml = list
            .map((item) => {
              const inCart = state.cart.find(
                (c) => c.storeId === currentStore.id && c.name === item.name && c.price === item.price
              );
              const qty = inCart ? inCart.qty : 0;
              return `
                <div class="menu-item-card ${qty > 0 ? 'in-cart' : ''}">
                  <div class="menu-item-main">
                    <div class="menu-item-name">${escapeHtml(item.name)}</div>
                    <div class="menu-item-price">$${item.price}</div>
                  </div>
                  <div class="menu-item-actions">
                    ${
                      qty > 0
                        ? `
                      <button class="qty-btn minus" data-name="${escapeHtml(item.name)}" data-price="${item.price}">−</button>
                      <span class="qty-num">${qty}</span>
                      <button class="qty-btn plus" data-name="${escapeHtml(item.name)}" data-price="${item.price}" ${closeStatus.closed ? 'disabled' : ''}>＋</button>
                    `
                        : `
                      <button class="add-item-btn" data-name="${escapeHtml(item.name)}" data-price="${item.price}" ${closeStatus.closed ? 'disabled' : ''}>
                        <i class="fa-solid fa-plus"></i> 點餐
                      </button>
                    `
                    }
                  </div>
                </div>
              `;
            })
            .join('');

          return `
            <div class="menu-category-block">
              <h3 class="menu-category-title"><span>${escapeHtml(cat)}</span> <small>${list.length} 項</small></h3>
              <div class="menu-grid">${cardsHtml}</div>
            </div>
          `;
        })
        .join('');

      menuContainer.querySelectorAll('.add-item-btn, .qty-btn.plus').forEach((btn) => {
        btn.addEventListener('click', () => {
          if (closeStatus.closed) return;
          const name = btn.getAttribute('data-name');
          const price = Number(btn.getAttribute('data-price'));
          updateCartItem(currentStore.id, currentStore.name, name, price, 1);
        });
      });

      menuContainer.querySelectorAll('.qty-btn.minus').forEach((btn) => {
        btn.addEventListener('click', () => {
          const name = btn.getAttribute('data-name');
          const price = Number(btn.getAttribute('data-price'));
          updateCartItem(currentStore.id, currentStore.name, name, price, -1);
        });
      });
    }

    renderCartSidebar(closeStatus);
    renderPublicTodayOrders();
  }

  function updateCartItem(storeId, storeName, itemName, price, delta) {
    const idx = state.cart.findIndex(
      (c) => c.storeId === storeId && c.name === itemName && c.price === price
    );
    if (idx > -1) {
      state.cart[idx].qty += delta;
      if (state.cart[idx].qty <= 0) {
        state.cart.splice(idx, 1);
      }
    } else if (delta > 0) {
      state.cart.push({ storeId, storeName, name: itemName, price, qty: delta });
    }
    renderOrderView();
  }

  function renderCartSidebar(closeStatus) {
    const cartListEl = document.getElementById('cartItemsList');
    const totalEl = document.getElementById('cartTotalPrice');
    const submitBtn = document.getElementById('submitOrderBtn');
    const nameInput = document.getElementById('orderUserName');

    if (nameInput && nameInput.value !== state.userName) {
      nameInput.value = state.userName;
    }

    if (state.cart.length === 0) {
      cartListEl.innerHTML = `<div class="cart-empty">尚未選擇便當，請從左側菜單點擊「＋點餐」</div>`;
      totalEl.textContent = '$0';
      submitBtn.disabled = true;
      return;
    }

    let total = 0;
    cartListEl.innerHTML = state.cart
      .map((item, idx) => {
        const sub = item.price * item.qty;
        total += sub;
        return `
          <div class="cart-row">
            <div class="cart-row-info">
              <span class="cart-store-tag">${escapeHtml(item.storeName)}</span>
              <div class="cart-item-title">${escapeHtml(item.name)}</div>
              <div class="cart-item-sub">$${item.price} × ${item.qty} = <strong>$${sub}</strong></div>
            </div>
            <div class="cart-row-ctrl">
              <button class="qty-btn-sm" data-idx="${idx}" data-delta="-1">−</button>
              <span>${item.qty}</span>
              <button class="qty-btn-sm" data-idx="${idx}" data-delta="1">＋</button>
            </div>
          </div>
        `;
      })
      .join('');

    totalEl.textContent = `$${total}`;
    submitBtn.disabled = closeStatus.closed;

    cartListEl.querySelectorAll('.qty-btn-sm').forEach((btn) => {
      btn.addEventListener('click', () => {
        const idx = Number(btn.getAttribute('data-idx'));
        const delta = Number(btn.getAttribute('data-delta'));
        const item = state.cart[idx];
        if (item) updateCartItem(item.storeId, item.storeName, item.name, item.price, delta);
      });
    });
  }

  function renderPublicTodayOrders() {
    const container = document.getElementById('publicTodayOrdersList');
    const todayOrders = getTodayOrders();
    if (todayOrders.length === 0) {
      container.innerHTML = `<div class="empty-state-sm">今日尚無人送出訂單，搶先成為第一位吧！</div>`;
      return;
    }

    container.innerHTML = todayOrders
      .slice()
      .reverse()
      .map((o) => {
        const itemsText = (o.items || []).map((i) => `${escapeHtml(i.name)} ×${i.qty}`).join('、');
        const noteHtml = o.note ? `<span class="order-note-pill">備註：${escapeHtml(o.note)}</span>` : '';
        const paidBadge = o.paid
          ? `<span class="paid-pill yes"><i class="fa-solid fa-check"></i> 已繳費</span>`
          : `<span class="paid-pill no">未繳費 ($${o.totalPrice})</span>`;
        return `
          <div class="public-order-row">
            <div class="public-order-left">
              <span class="public-user-avatar">${escapeHtml(o.userName.slice(0, 1))}</span>
              <div>
                <div class="public-user-name">${escapeHtml(o.userName)} <small class="store-label-sm">${escapeHtml(o.storeName)}</small></div>
                <div class="public-order-items">${itemsText} ${noteHtml}</div>
              </div>
            </div>
            <div class="public-order-right">
              ${paidBadge}
            </div>
          </div>
        `;
      })
      .join('');
  }

  // ----- 送出訂單 -----
  async function handleOrderSubmit() {
    const closeStatus = isOrderingClosed();
    if (closeStatus.closed) {
      showToast(closeStatus.reason, 'error');
      return;
    }

    const nameInput = document.getElementById('orderUserName');
    const noteInput = document.getElementById('orderNoteInput');
    const userName = (nameInput.value || '').trim();
    const note = (noteInput.value || '').trim();

    if (!userName) {
      showToast('請輸入您的姓名！', 'error');
      nameInput.focus();
      return;
    }
    if (state.cart.length === 0) {
      showToast('請至少選擇一項餐點！', 'error');
      return;
    }

    state.userName = userName;
    localStorage.setItem(STORAGE_KEYS.LAST_USER, userName);

    // 依店家拆單或合併送出 (若同時點多家，依店家產生對應訂單方便統計電話叫餐)
    const byStore = {};
    state.cart.forEach((c) => {
      if (!byStore[c.storeId]) {
        byStore[c.storeId] = { storeId: c.storeId, storeName: c.storeName, items: [], totalPrice: 0 };
      }
      byStore[c.storeId].items.push({ name: c.name, price: c.price, qty: c.qty });
      byStore[c.storeId].totalPrice += c.price * c.qty;
    });

    const submitBtn = document.getElementById('submitOrderBtn');
    submitBtn.disabled = true;
    submitBtn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> 送出訂單中...`;

    for (const group of Object.values(byStore)) {
      const newOrder = {
        id: 'ord_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
        date: state.session.date,
        userName,
        storeId: group.storeId,
        storeName: group.storeName,
        items: group.items,
        totalPrice: group.totalPrice,
        note,
        paid: false,
        createdAt: new Date().toLocaleTimeString('zh-TW', { hour: '2-digit', minute: '2-digit' })
      };
      state.orders.push(newOrder);
      await pushToCloud('submitOrder', { order: newOrder });
    }

    state.cart = [];
    noteInput.value = '';
    submitBtn.disabled = false;
    submitBtn.innerHTML = `<i class="fa-solid fa-paper-plane"></i> 確認送出訂購`;

    renderAll();
    showToast(`🎉 訂購成功！${userName} 的訂單已送出`, 'success');
  }

  // ===== 6. 管理後台渲染邏輯 =====
  function renderAdminView() {
    // 切換後台子分頁
    document.querySelectorAll('.admin-tab-btn').forEach((btn) => {
      const tab = btn.getAttribute('data-tab');
      btn.classList.toggle('active', tab === state.adminTab);
    });
    document.querySelectorAll('.admin-panel').forEach((panel) => {
      panel.classList.toggle('hidden', panel.id !== `adminPanel_${state.adminTab}`);
    });

    if (state.adminTab === 'dashboard') {
      renderAdminDashboard();
    } else if (state.adminTab === 'session') {
      renderAdminSessionSetup();
    } else if (state.adminTab === 'menus') {
      renderAdminMenuManager();
    } else if (state.adminTab === 'cloud') {
      renderAdminCloudSetup();
    }
  }

  // ----- 後台 Tab 1: 今日訂單統計 (打電話訂購清單) & 繳費管理 -----
  function renderAdminDashboard() {
    const todayOrders = getTodayOrders();
    let totalAmount = 0;
    let paidAmount = 0;
    let totalItemsCount = 0;
    let paidOrdersCount = 0;

    todayOrders.forEach((o) => {
      totalAmount += o.totalPrice;
      if (o.paid) {
        paidAmount += o.totalPrice;
        paidOrdersCount++;
      }
      (o.items || []).forEach((i) => {
        totalItemsCount += i.qty;
      });
    });

    const unpaidAmount = totalAmount - paidAmount;

    document.getElementById('statTotalOrders').textContent = `${todayOrders.length} 筆 (${totalItemsCount} 份)`;
    document.getElementById('statTotalAmount').textContent = `$${totalAmount}`;
    document.getElementById('statPaidAmount').textContent = `$${paidAmount} (${paidOrdersCount}/${todayOrders.length} 人已繳)`;
    document.getElementById('statUnpaidAmount').textContent = `$${unpaidAmount}`;

    // 快速收單切換按鈕
    const toggleCloseBtn = document.getElementById('quickToggleCloseBtn');
    if (state.session.isClosed) {
      toggleCloseBtn.className = 'btn btn-success';
      toggleCloseBtn.innerHTML = `<i class="fa-solid fa-door-open"></i> 重新開放點餐`;
    } else {
      toggleCloseBtn.className = 'btn btn-danger';
      toggleCloseBtn.innerHTML = `<i class="fa-solid fa-gavel"></i> 立即結單 / 停止收單`;
    }

    // 1. 依店家與品項彙整「電話訂購總表」
    const phoneContainer = document.getElementById('phoneCallSummaryContainer');
    if (todayOrders.length === 0) {
      phoneContainer.innerHTML = `<div class="empty-state">今日 (${state.session.date}) 尚無任何訂購紀錄</div>`;
    } else {
      const storeGroups = {};
      todayOrders.forEach((o) => {
        const key = o.storeId || o.storeName;
        if (!storeGroups[key]) {
          const stObj = state.stores.find((s) => s.id === o.storeId || s.name === o.storeName);
          storeGroups[key] = {
            storeName: o.storeName,
            storeInfo: stObj ? stObj.info : '',
            totalStorePrice: 0,
            totalStoreQty: 0,
            itemMap: {} // itemName -> { price, qty, buyers: [] }
          };
        }
        const g = storeGroups[key];
        g.totalStorePrice += o.totalPrice;
        (o.items || []).forEach((it) => {
          g.totalStoreQty += it.qty;
          if (!g.itemMap[it.name]) {
            g.itemMap[it.name] = { name: it.name, price: it.price, qty: 0, buyers: [] };
          }
          g.itemMap[it.name].qty += it.qty;
          const buyerLabel = o.note
            ? `${o.userName}${it.qty > 1 ? '×' + it.qty : ''}(${o.note})`
            : `${o.userName}${it.qty > 1 ? '×' + it.qty : ''}`;
          g.itemMap[it.name].buyers.push(buyerLabel);
        });
      });

      phoneContainer.innerHTML = Object.values(storeGroups)
        .map((g) => {
          const itemRows = Object.values(g.itemMap)
            .sort((a, b) => b.qty - a.qty)
            .map(
              (item) => `
              <tr>
                <td class="call-item-name"><strong>${escapeHtml(item.name)}</strong></td>
                <td class="call-item-qty"><span class="qty-badge-lg">${item.qty} 份</span></td>
                <td>$${item.price}</td>
                <td><strong>$${item.price * item.qty}</strong></td>
                <td class="call-item-buyers">${escapeHtml(item.buyers.join('、'))}</td>
              </tr>
            `
            )
            .join('');

          return `
            <div class="call-store-card">
              <div class="call-store-header">
                <div>
                  <h3 class="call-store-title">📞 ${escapeHtml(g.storeName)}</h3>
                  <div class="call-store-phone">${escapeHtml(g.storeInfo || '請參考店家電話資訊')}</div>
                </div>
                <div class="call-store-summary">
                  <span>共 <strong>${g.totalStoreQty}</strong> 份餐點</span>
                  <span class="call-store-total">應付總計：<strong>$${g.totalStorePrice}</strong></span>
                  <button class="btn btn-sm btn-outline copy-call-btn" data-store="${escapeHtml(g.storeName)}">
                    <i class="fa-regular fa-copy"></i> 複製叫餐清單
                  </button>
                </div>
              </div>
              <table class="admin-table call-table">
                <thead>
                  <tr>
                    <th>餐點品項</th>
                    <th>統計數量 (唸給店家)</th>
                    <th>單價</th>
                    <th>小計</th>
                    <th>訂購人明細</th>
                  </tr>
                </thead>
                <tbody>${itemRows}</tbody>
              </table>
            </div>
          `;
        })
        .join('');

      phoneContainer.querySelectorAll('.copy-call-btn').forEach((btn) => {
        btn.addEventListener('click', () => {
          const sName = btn.getAttribute('data-store');
          const grp = Object.values(storeGroups).find((x) => x.storeName === sName);
          if (!grp) return;
          const lines = [
            `【${grp.storeName}】訂購清單 (${state.session.date})`,
            ...Object.values(grp.itemMap).map((i) => `・${i.name} x ${i.qty} 份 ($${i.price * i.qty})`),
            `----------------`,
            `總數量：${grp.totalStoreQty} 份｜總金額：$${grp.totalStorePrice}`
          ];
          navigator.clipboard.writeText(lines.join('\n'));
          showToast('已複製電話叫餐清單！', 'success');
        });
      });
    }

    // 2. 訂購人收款管理表
    const paymentTableBody = document.getElementById('paymentManagementBody');
    if (todayOrders.length === 0) {
      paymentTableBody.innerHTML = `<tr><td colspan="7" class="empty-cell">尚無訂單資料</td></tr>`;
    } else {
      paymentTableBody.innerHTML = todayOrders
        .map((o) => {
          const itemsStr = (o.items || []).map((i) => `${escapeHtml(i.name)} ×${i.qty}`).join('、');
          return `
            <tr class="${o.paid ? 'row-paid' : 'row-unpaid'}">
              <td><strong>${escapeHtml(o.userName)}</strong></td>
              <td>${escapeHtml(o.storeName)}</td>
              <td>${itemsStr}</td>
              <td>${escapeHtml(o.note || '-')}</td>
              <td class="price-col"><strong>$${o.totalPrice}</strong></td>
              <td>
                <button class="paid-toggle-btn ${o.paid ? 'is-paid' : 'is-unpaid'}" data-id="${o.id}">
                  ${
                    o.paid
                      ? `<i class="fa-solid fa-circle-check"></i> 已繳錢`
                      : `<i class="fa-regular fa-circle"></i> 未繳錢 (點擊標記)`
                  }
                </button>
              </td>
              <td>
                <button class="btn-icon delete-order-btn" data-id="${o.id}" title="刪除此筆訂單">
                  <i class="fa-regular fa-trash-can"></i>
                </button>
              </td>
            </tr>
          `;
        })
        .join('');

      paymentTableBody.querySelectorAll('.paid-toggle-btn').forEach((btn) => {
        btn.addEventListener('click', async () => {
          const id = btn.getAttribute('data-id');
          const ord = state.orders.find((x) => x.id === id);
          if (!ord) return;
          ord.paid = !ord.paid;
          renderAll();
          await pushToCloud('togglePaid', { orderId: id, paid: ord.paid });
        });
      });

      paymentTableBody.querySelectorAll('.delete-order-btn').forEach((btn) => {
        btn.addEventListener('click', async () => {
          const id = btn.getAttribute('data-id');
          const ord = state.orders.find((x) => x.id === id);
          if (!ord) return;
          if (!confirm(`確定要刪除「${ord.userName}」的訂單嗎？`)) return;
          state.orders = state.orders.filter((x) => x.id !== id);
          renderAll();
          await pushToCloud('deleteOrder', { orderId: id });
          showToast('已刪除該筆訂單', 'info');
        });
      });
    }
  }

  // ----- 後台 Tab 2: 今日開團設定與分享網址 -----
  function renderAdminSessionSetup() {
    document.getElementById('sessionDateInput').value = state.session.date;
    document.getElementById('sessionDeadlineInput').value = state.session.deadline || '';
    document.getElementById('sessionAnnouncementInput').value = state.session.announcement || '';

    const checklistEl = document.getElementById('sessionStoresChecklist');
    const activeIds = state.session.activeStoreIds || [];

    checklistEl.innerHTML = state.stores
      .map((s) => {
        const checked = activeIds.includes(s.id);
        return `
          <label class="store-check-card ${checked ? 'checked' : ''}">
            <input type="checkbox" value="${s.id}" ${checked ? 'checked' : ''} />
            <div class="store-check-body">
              <div class="store-check-name">${escapeHtml(s.name)}</div>
              <div class="store-check-meta">${s.items.length} 品項 ｜ ${escapeHtml(s.info || '')}</div>
            </div>
          </label>
        `;
      })
      .join('');

    checklistEl.querySelectorAll('input[type="checkbox"]').forEach((chk) => {
      chk.addEventListener('change', () => {
        chk.closest('.store-check-card').classList.toggle('checked', chk.checked);
      });
    });

    updateShareUrlBox();
  }

  function updateShareUrlBox() {
    const url = new URL(window.location.href.split('?')[0]);
    url.searchParams.set('date', state.session.date);
    if (state.session.activeStoreIds && state.session.activeStoreIds.length > 0) {
      url.searchParams.set('stores', state.session.activeStoreIds.join(','));
    }
    if (state.session.deadline) {
      url.searchParams.set('deadline', state.session.deadline);
    }
    if (state.session.announcement) {
      url.searchParams.set('note', state.session.announcement);
    }
    if (state.cloudConfig.type === 'gas' && state.cloudConfig.gasUrl) {
      url.searchParams.set('gas', state.cloudConfig.gasUrl);
    } else if (state.cloudConfig.type === 'firebase' && state.cloudConfig.firebaseUrl) {
      url.searchParams.set('fb', state.cloudConfig.firebaseUrl);
    }
    const shareInput = document.getElementById('shareOrderUrlInput');
    if (shareInput) shareInput.value = url.toString();
  }

  // ----- 後台 Tab 3: 匯入便當 Excel 與菜單編輯管理 -----
  function renderAdminMenuManager() {
    document.getElementById('totalStoresCountLabel').textContent = state.stores.length;
    const listEl = document.getElementById('adminStoreEditorList');

    if (!state.editingStoreId && state.stores.length > 0) {
      state.editingStoreId = state.stores[0].id;
    }

    listEl.innerHTML = state.stores
      .map((s) => {
        const active = s.id === state.editingStoreId ? 'active' : '';
        return `
          <div class="editor-store-item ${active}" data-id="${s.id}">
            <span>${escapeHtml(s.name)}</span>
            <small>${s.items.length} 項</small>
          </div>
        `;
      })
      .join('');

    listEl.querySelectorAll('.editor-store-item').forEach((el) => {
      el.addEventListener('click', () => {
        state.editingStoreId = el.getAttribute('data-id');
        renderAdminMenuManager();
      });
    });

    const editStore = state.stores.find((s) => s.id === state.editingStoreId);
    const detailEl = document.getElementById('adminStoreEditorDetail');
    if (!editStore) {
      detailEl.innerHTML = `<div class="empty-state">請選擇左側店家檢視或編輯菜單</div>`;
      return;
    }

    const rowsHtml = editStore.items
      .map(
        (it, idx) => `
        <tr>
          <td><input type="text" class="input-sm edit-item-cat" data-idx="${idx}" value="${escapeHtml(it.category || '餐點')}" /></td>
          <td><input type="text" class="input-sm edit-item-name" data-idx="${idx}" value="${escapeHtml(it.name)}" /></td>
          <td><input type="number" class="input-sm edit-item-price" data-idx="${idx}" value="${it.price}" /></td>
          <td>
            <button class="btn-icon remove-menu-item-btn" data-idx="${idx}" title="刪除此品項">
              <i class="fa-regular fa-trash-can"></i>
            </button>
          </td>
        </tr>
      `
      )
      .join('');

    detailEl.innerHTML = `
      <div class="editor-detail-header">
        <div class="form-row">
          <div class="form-group">
            <label>店家名稱</label>
            <input type="text" id="editStoreNameInput" class="form-control" value="${escapeHtml(editStore.name)}" />
          </div>
          <div class="form-group flex-2">
            <label>電話與外送條件資訊</label>
            <input type="text" id="editStoreInfoInput" class="form-control" value="${escapeHtml(editStore.info || '')}" />
          </div>
        </div>
        <div class="editor-actions-bar">
          <button class="btn btn-sm btn-outline" id="addMenuItemRowBtn"><i class="fa-solid fa-plus"></i> 新增品項</button>
          <button class="btn btn-sm btn-primary" id="saveStoreMenuChangesBtn"><i class="fa-solid fa-floppy-disk"></i> 儲存店家菜單修改</button>
          <button class="btn btn-sm btn-danger" id="deleteWholeStoreBtn"><i class="fa-regular fa-trash-can"></i> 刪除此店家</button>
        </div>
      </div>
      <div class="table-scroll">
        <table class="admin-table">
          <thead>
            <tr>
              <th style="width: 140px;">分類</th>
              <th>品項名稱</th>
              <th style="width: 110px;">價格 ($)</th>
              <th style="width: 60px;">刪除</th>
            </tr>
          </thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      </div>
    `;

    document.getElementById('addMenuItemRowBtn').addEventListener('click', () => {
      editStore.items.unshift({ category: '餐點', name: '新品項', price: 100 });
      renderAdminMenuManager();
    });

    detailEl.querySelectorAll('.remove-menu-item-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        const idx = Number(btn.getAttribute('data-idx'));
        editStore.items.splice(idx, 1);
        renderAdminMenuManager();
      });
    });

    document.getElementById('saveStoreMenuChangesBtn').addEventListener('click', async () => {
      editStore.name = document.getElementById('editStoreNameInput').value.trim() || editStore.name;
      editStore.info = document.getElementById('editStoreInfoInput').value.trim();
      const cats = detailEl.querySelectorAll('.edit-item-cat');
      const names = detailEl.querySelectorAll('.edit-item-name');
      const prices = detailEl.querySelectorAll('.edit-item-price');
      const updatedItems = [];
      for (let i = 0; i < names.length; i++) {
        const nm = names[i].value.trim();
        const pr = Number(prices[i].value) || 0;
        const ct = cats[i].value.trim() || '餐點';
        if (nm && pr > 0) updatedItems.push({ category: ct, name: nm, price: pr });
      }
      editStore.items = updatedItems;
      editStore.itemCount = updatedItems.length;
      await pushToCloud('saveCustomStores', { customStores: state.stores });
      renderAll();
      showToast(`已儲存「${editStore.name}」菜單！`, 'success');
    });

    document.getElementById('deleteWholeStoreBtn').addEventListener('click', async () => {
      if (!confirm(`確定要刪除「${editStore.name}」整家店的菜單嗎？`)) return;
      state.stores = state.stores.filter((s) => s.id !== editStore.id);
      state.editingStoreId = state.stores[0] ? state.stores[0].id : null;
      await pushToCloud('saveCustomStores', { customStores: state.stores });
      renderAll();
      showToast('已刪除該店家', 'info');
    });
  }

  // ----- 後台 Tab 4: 雲端同步與多人連線設定 -----
  function renderAdminCloudSetup() {
    document.getElementById('cloudTypeSelect').value = state.cloudConfig.type;
    document.getElementById('gasUrlInput').value = state.cloudConfig.gasUrl || '';
    document.getElementById('firebaseUrlInput').value = state.cloudConfig.firebaseUrl || '';

    document.getElementById('gasConfigBlock').classList.toggle('hidden', state.cloudConfig.type !== 'gas');
    document.getElementById('firebaseConfigBlock').classList.toggle('hidden', state.cloudConfig.type !== 'firebase');
  }

  // ===== 7. 事件綁定 =====
  function bindEvents() {
    // 頂部模式切換
    document.getElementById('navOrderBtn').addEventListener('click', () => {
      state.mode = 'order';
      renderAll();
    });
    document.getElementById('navAdminBtn').addEventListener('click', () => {
      state.mode = 'admin';
      renderAll();
    });

    // 菜單搜尋
    const searchInput = document.getElementById('menuSearchInput');
    searchInput.addEventListener('input', () => {
      state.menuSearch = searchInput.value;
      renderOrderView();
    });

    // 訂購人姓名同步
    document.getElementById('orderUserName').addEventListener('input', (e) => {
      state.userName = e.target.value;
    });

    // 送出訂單
    document.getElementById('submitOrderBtn').addEventListener('click', handleOrderSubmit);

    // 後台子分頁切換
    document.querySelectorAll('.admin-tab-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        state.adminTab = btn.getAttribute('data-tab');
        renderAdminView();
      });
    });

    // 後台：手動結單/開放點餐
    document.getElementById('quickToggleCloseBtn').addEventListener('click', async () => {
      state.session.isClosed = !state.session.isClosed;
      renderAll();
      await pushToCloud('saveSession', { session: state.session });
      showToast(state.session.isClosed ? '已手動結單關閉點餐！' : '已重新開放點餐！', 'info');
    });

    // 後台：立即重新整理雲端訂單
    document.getElementById('refreshOrdersBtn').addEventListener('click', () => {
      syncFromCloud(false);
      showToast('已更新最新訂單狀態', 'info');
    });

    // 後台：匯出今日訂單 Excel
    document.getElementById('exportTodayExcelBtn').addEventListener('click', exportTodayOrdersToExcel);

    // 後台：清空今日訂單
    document.getElementById('clearTodayOrdersBtn').addEventListener('click', async () => {
      if (!confirm(`確定要清空 ${state.session.date} 的所有訂單紀錄嗎？此操作無法復原。`)) return;
      state.orders = state.orders.filter((o) => o.date !== state.session.date);
      renderAll();
      await pushToCloud('clearDateOrders', { date: state.session.date });
      showToast('已清空當日訂單紀錄', 'info');
    });

    // 後台：儲存今日開團設定
    document.getElementById('saveSessionBtn').addEventListener('click', async () => {
      const checkedIds = Array.from(
        document.querySelectorAll('#sessionStoresChecklist input[type="checkbox"]:checked')
      ).map((el) => el.value);

      if (checkedIds.length === 0) {
        showToast('請至少勾選一家今日要開放訂購的便當店！', 'error');
        return;
      }

      state.session.date = document.getElementById('sessionDateInput').value || getTodayString();
      state.session.deadline = document.getElementById('sessionDeadlineInput').value;
      state.session.announcement = document.getElementById('sessionAnnouncementInput').value.trim();
      state.session.activeStoreIds = checkedIds;
      state.session.isClosed = false;

      ensureValidSelectedStore();
      await pushToCloud('saveSession', { session: state.session });
      renderAll();
      showToast('✅ 已發布今日開團菜單！可複製下方網址傳給同事點餐', 'success');
    });

    // 全選 / 取消全選店家
    document.getElementById('selectAllStoresBtn').addEventListener('click', () => {
      document.querySelectorAll('#sessionStoresChecklist input[type="checkbox"]').forEach((c) => {
        c.checked = true;
        c.closest('.store-check-card').classList.add('checked');
      });
    });
    document.getElementById('clearAllStoresBtn').addEventListener('click', () => {
      document.querySelectorAll('#sessionStoresChecklist input[type="checkbox"]').forEach((c) => {
        c.checked = false;
        c.closest('.store-check-card').classList.remove('checked');
      });
    });

    // 複製點餐網址
    document.getElementById('copyShareUrlBtn').addEventListener('click', () => {
      updateShareUrlBox();
      const input = document.getElementById('shareOrderUrlInput');
      input.select();
      navigator.clipboard.writeText(input.value);
      showToast('📋 已複製訂購網址！快貼到 LINE 或 Teams 群組給同事吧', 'success');
    });

    // 後台：上傳匯入 Excel 檔案
    const excelInput = document.getElementById('excelFileInput');
    const dropZone = document.getElementById('excelDropZone');

    excelInput.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (file) handleExcelFileUpload(file);
    });

    dropZone.addEventListener('dragover', (e) => {
      e.preventDefault();
      dropZone.classList.add('dragover');
    });
    dropZone.addEventListener('dragleave', () => {
      dropZone.classList.remove('dragover');
    });
    dropZone.addEventListener('drop', (e) => {
      e.preventDefault();
      dropZone.classList.remove('dragover');
      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
        handleExcelFileUpload(e.dataTransfer.files[0]);
      }
    });

    // 後台：還原預設 35 家便當店資料
    document.getElementById('resetDefaultStoresBtn').addEventListener('click', async () => {
      if (!confirm('確定要還原為系統內建的 35 家便當店菜單嗎？')) return;
      if (window.DEFAULT_STORES) {
        state.stores = JSON.parse(JSON.stringify(window.DEFAULT_STORES));
        state.editingStoreId = state.stores[0].id;
        await pushToCloud('saveCustomStores', { customStores: state.stores });
        renderAll();
        showToast('已還原預設 35 家便當店菜單！', 'success');
      }
    });

    // 後台：新增空白店家
    document.getElementById('createNewStoreBtn').addEventListener('click', () => {
      const name = prompt('請輸入新便當店名稱：');
      if (!name || !name.trim()) return;
      const newStore = {
        id: 'store_custom_' + Date.now(),
        sheetName: name.trim(),
        name: name.trim(),
        info: '外送電話：請填寫',
        notes: '',
        itemCount: 1,
        items: [{ category: '便當', name: '招牌便當', price: 100 }]
      };
      state.stores.unshift(newStore);
      state.editingStoreId = newStore.id;
      saveLocalState();
      renderAdminMenuManager();
      showToast(`已新增「${newStore.name}」，請編輯品項後點擊儲存`, 'success');
    });

    // 後台：雲端設定切換與儲存
    document.getElementById('cloudTypeSelect').addEventListener('change', (e) => {
      state.cloudConfig.type = e.target.value;
      renderAdminCloudSetup();
    });

    document.getElementById('saveCloudConfigBtn').addEventListener('click', async () => {
      state.cloudConfig.type = document.getElementById('cloudTypeSelect').value;
      state.cloudConfig.gasUrl = document.getElementById('gasUrlInput').value.trim();
      state.cloudConfig.firebaseUrl = document.getElementById('firebaseUrlInput').value.trim();
      saveLocalState();
      updateShareUrlBox();
      if (state.cloudConfig.type !== 'local') {
        await pushToCloud('saveSession', { session: state.session });
        await syncFromCloud(false);
      } else {
        renderHeader();
      }
      showToast('已儲存雲端連線設定！', 'success');
    });
  }

  function handleExcelFileUpload(file) {
    if (typeof XLSX === 'undefined') {
      showToast('Excel 解析元件尚未載入，請確認網路連線', 'error');
      return;
    }
    const reader = new FileReader();
    reader.onload = async function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const parsed = parseWorkbookToStores(workbook);
        if (parsed.length === 0) {
          showToast('無法從此 Excel 解析出任何菜單品項，請確認格式', 'error');
          return;
        }
        state.stores = parsed;
        state.editingStoreId = parsed[0].id;
        // 若原本勾選的店家 ID 仍存在則保留，否則預設勾選第 1 家
        const validActive = (state.session.activeStoreIds || []).filter((id) =>
          parsed.some((s) => s.id === id)
        );
        state.session.activeStoreIds = validActive.length > 0 ? validActive : [parsed[0].id];

        await pushToCloud('saveCustomStores', { customStores: state.stores });
        renderAll();
        const totalItems = parsed.reduce((acc, s) => acc + s.items.length, 0);
        showToast(`🎉 匯入成功！共解析出 ${parsed.length} 家便當店、${totalItems} 道餐點`, 'success');
      } catch (err) {
        console.error(err);
        showToast('解析 Excel 檔案時發生錯誤：' + err.message, 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function exportTodayOrdersToExcel() {
    const todayOrders = getTodayOrders();
    if (todayOrders.length === 0) {
      showToast('今日尚無訂單可匯出', 'error');
      return;
    }
    if (typeof XLSX === 'undefined') {
      showToast('Excel 匯出元件未載入', 'error');
      return;
    }

    const rows = [
      ['訂購日期', '訂購人姓名', '店家名稱', '餐點明細', '備註', '應付金額', '繳費狀態', '送出時間']
    ];
    todayOrders.forEach((o) => {
      const itemsText = (o.items || []).map((i) => `${i.name} x${i.qty}`).join('、');
      rows.push([
        o.date,
        o.userName,
        o.storeName,
        itemsText,
        o.note || '',
        o.totalPrice,
        o.paid ? '已繳費' : '未繳費',
        o.createdAt || ''
      ]);
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, `訂單明細_${state.session.date}`);
    XLSX.writeFile(wb, `便當訂購與收款清單_${state.session.date}.xlsx`);
    showToast('已匯出 Excel 訂購與收款清單！', 'success');
  }

  // ===== 工具函式 =====
  function escapeHtml(str) {
    if (str === null || str === undefined) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function showToast(msg, type = 'info') {
    const container = document.getElementById('toastContainer');
    if (!container) return;
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.textContent = msg;
    container.appendChild(toast);
    setTimeout(() => {
      toast.classList.add('fade-out');
      setTimeout(() => toast.remove(), 300);
    }, 3200);
  }

  // 啟動系統
  document.addEventListener('DOMContentLoaded', () => {
    bindEvents();
    initApp();
  });
})();
