// ============================================
// Amazon Inventory Assistant - Web Frontend
// Core logic rewritten in JavaScript for Cloudflare
// ============================================

const API_BASE = ''; // Same origin for Pages + Functions
const APP_VERSION = '1.0.0';

let authToken = localStorage.getItem('token');
let currentUser = null;
let workbook = null;
let worksheet = null;
let headers = [];
let colIdxMap = {};
let greenRows = new Set();
let seckillItems = [];
let tableData = [];
let originalFileBuffer = null; // Store original file ArrayBuffer for ExcelJS format-preserving export

// ========== Utility Functions ==========
function parseNum(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  const cleaned = String(v).replace(/,/g, '').replace(/\s/g, '');
  const n = parseFloat(cleaned);
  return isNaN(n) ? 0 : n;
}

function getNextFriday() {
  const d = new Date();
  const day = d.getDay();
  const diff = (5 - day + 7) % 7 || 7;
  d.setDate(d.getDate() + diff);
  return `${d.getMonth() + 1}月${d.getDate()}日`;
}

function estimateShipDate(remaining, dailySales) {
  if (remaining <= 0) return [getNextFriday(), ''];
  if (dailySales <= 0) return [`${Math.round(remaining)}-充公给别人`, '【充公建议】销量为0且库存不足，预计超过3个月才可发出，建议充公给别的团队卖'];

  const daysNeeded = remaining / dailySales;
  if (daysNeeded > 90) return [`${Math.round(remaining)}-充公给别人`, '【充公建议】预计超过3个月才可发出，建议充公给别的团队卖'];
  if (daysNeeded > 30) {
    return ['下个月看', `【预计发货】剩余${Math.round(remaining)}个，按日均${dailySales.toFixed(1)}个约需${Math.round(daysNeeded)}天，预计下个月可发`];
  }
  const est = new Date();
  est.setDate(est.getDate() + Math.round(daysNeeded));
  return [`${est.getMonth() + 1}月${est.getDate()}日`, `【预计发货】剩余${Math.round(remaining)}个，约${Math.round(daysNeeded)}天后可发(${est.getMonth() + 1}月${est.getDate()}日)`];
}

function detectProductType(sku, code) {
  const s = String(sku || '').toLowerCase();
  const c = String(code || '').toLowerCase();
  const combined = s + ' ' + c;
  if (combined.includes('asus') || combined.includes('square') || combined.includes('方形')) return 'square';
  if (combined.includes('gan-50') || combined.includes('gan50') || combined.includes('氮化镓50') || combined.includes('50pcs')) return 'gan50';
  if (combined.includes('gan') || combined.includes('氮化镓') || combined.includes('xiaomi') || combined.includes('小米') || combined.includes('huawei') || combined.includes('华为') || combined.includes('小壳子')) return 'gan100';
  if (combined.includes('25装') || combined.includes('25pcs') || combined.includes('25 pc') || combined.includes('25个/箱')) return 'hp25';
  if (combined.includes('100w') || combined.includes('120w') || combined.includes('140w') || combined.includes('240w') || combined.includes('大功率')) return 'hp40';
  return 'long';
}

const PRODUCT_TYPES = {
  square: { name: '方形壳/华硕壳', qty: 60 },
  long: { name: '长条普通壳', qty: 50 },
  hp40: { name: '大功率(40装)', qty: 40 },
  hp25: { name: '大功率(25装)', qty: 25 },
  gan100: { name: '氮化镓/小壳子(100装)', qty: 100 },
  gan50: { name: '氮化镓/小壳子(50装)', qty: 50 },
};

function isGreenCell(cell) {
  if (!cell || !cell.s || !cell.s.fgColor) return false;
  const rgb = cell.s.fgColor.rgb;
  if (!rgb || rgb === '00000000' || rgb === 'FFFFFFFF') return false;
  if (rgb.length >= 6) {
    const r = parseInt(rgb.slice(-6, -4), 16);
    const g = parseInt(rgb.slice(-4, -2), 16);
    const b = parseInt(rgb.slice(-2), 16);
    return g > r + 25 && g > b + 25 && g > 60;
  }
  return false;
}

// ========== Column Mapping ==========
function autoMapColumns(hdrs) {
  const aliases = {
    asin: ['asin'],
    sku: ['sku', 'fnsku'],
    code: ['编码'],
    sales30: ['近30天销量', '30天销量', '月销量'],
    sales15: ['15日销量', '15天销量', '十五日销量'],
    sales7: ['7日销量', '7天销量', '七日销量'],
    sales3: ['3日销量', '3天销量', '三日销量'],
    monthlySales: ['月销量'],
    available: ['可售库存', '库存', '可用库存', 'fba库存', '现有库存'],
    unshipped: ['在途', '在途库存'],
    unshippedOrders: ['未出货订单', '未出货'],
    purchasing: ['采购中', '采购中库存', '下单中', '已采购'],
    weekShip: ['本周要出库库存', '本周要出', '本周出库', '本周要出库存'],
    seaSup: ['海运补单', '海运补'],
    seaOrder: ['海运下单', '下单', '海运'],
    date: ['改出货日期', '出货日期', '日期'],
    remark: ['备注', '说明'],
    handling: ['清货建议', '处理建议', '充公建议', '其他建议', '行动建议'],
    confiscate: ['充公数量', '充公', '转其他团队'],
    aiAdvice: ['AI建议', '发货建议', '渠道建议', '物流建议'],
  };

  const colMap = {};
  // Special: X月销量 pattern
  for (const h of hdrs) {
    const hStr = h !== null && h !== undefined ? String(h) : '';
    if (/\d+月销量/.test(hStr)) {
      colMap.monthlySales = hStr;
      break;
    }
  }

  for (const [key, alts] of Object.entries(aliases)) {
    if (key === 'monthlySales' && colMap.monthlySales) continue;
    for (const a of alts) {
      const aLower = a.toLowerCase();
      // Exact match first
      for (const h of hdrs) {
        const hStr = h !== null && h !== undefined ? String(h) : '';
        if (hStr.toLowerCase() === aLower) { colMap[key] = hStr; break; }
      }
      if (colMap[key]) break;
      // Substring match
      for (const h of hdrs) {
        const hStr = h !== null && h !== undefined ? String(h) : '';
        if (hStr.toLowerCase().includes(aLower)) { colMap[key] = hStr; break; }
      }
      if (colMap[key]) break;
    }
  }

  // FBA shipment columns
  const fbaCols = [];
  for (const h of hdrs) {
    const hStr = h !== null && h !== undefined ? String(h) : '';
    if (hStr.toUpperCase().startsWith('FBA') && hStr.length > 3) fbaCols.push(hStr);
  }
  if (fbaCols.length) colMap.fba_shipments = fbaCols;

  return colMap;
}

// ========== Core Analysis ==========
function analyzeRow(row, colMap, isGreen = false) {
  const sales30 = parseNum(row[colMap.sales30]);
  const sales15 = parseNum(row[colMap.sales15]);
  const sales7 = parseNum(row[colMap.sales7]);
  const sales3 = parseNum(row[colMap.sales3]);
  const monthlySales = parseNum(row[colMap.monthlySales]);
  const available = parseNum(row[colMap.available]);
  let unshipped = parseNum(row[colMap.unshipped]);
  for (const col of colMap.fba_shipments || []) {
    unshipped += parseNum(row[col]);
  }
  const purchasing = parseNum(row[colMap.purchasing]);
  const unshippedOrders = parseNum(row[colMap.unshippedOrders]);
  const sku = row[colMap.sku] || '';
  const code = row[colMap.code] || '';
  const asin = row[colMap.asin] || '';

  const daily30 = sales30 > 0 ? sales30 / 30 : 0;
  const daily15 = sales15 > 0 ? sales15 / 15 : daily30;
  const daily7 = sales7 > 0 ? sales7 / 7 : daily30;
  const daily3 = sales3 > 0 ? sales3 / 3 : daily7;
  const dailyWeighted = daily3 * 0.30 + daily7 * 0.20 + daily15 * 0.25 + daily30 * 0.25;

  const totalInv = available + unshipped + purchasing;
  const isDeep = sales30 > 600;
  const targetStock = isDeep ? Math.round(sales30 * 4) : Math.round(sales30 * 2);
  const suggestRestock = Math.max(0, targetStock - totalInv);

  const excessThreshold = sales30 * 5;
  const availableExcess = (sales30 > 0 && available > excessThreshold) ? available - excessThreshold : 0;
  const isExcessAvailable = availableExcess > 0;
  const isExcessRestock = suggestRestock > excessThreshold && sales30 > 0;
  let restockExcessQty = 0;
  let actualRestock = suggestRestock;
  if (isExcessRestock) {
    restockExcessQty = Math.round(suggestRestock - excessThreshold);
    actualRestock = excessThreshold;
  }
  const isExcess = isExcessAvailable || isExcessRestock;

  const ptype = detectProductType(sku, code);
  const qtyPerBox = PRODUCT_TYPES[ptype]?.qty || 50;

  const isHighVolume = monthlySales > 150;
  let isStagnant = false;
  let stagnantMsg = '';
  if (sales30 > 0 && totalInv > sales30 * 6) {
    isStagnant = true;
    stagnantMsg = '【滞销提醒】总库存>30天销量6倍，建议加快动销';
  }

  // Stagnant interception
  if (isStagnant && unshippedOrders > 0) {
    const confiscateQty = Math.round(unshippedOrders);
    const handlingRemarks = [];
    if (isGreen) handlingRemarks.push('【绿标/翔标】该产品可能有绿标/翔标，请复制ASIN到详情页确认后再操作');
    handlingRemarks.push(`【充公建议】总库存滞销，未出货${confiscateQty}个建议全部充公给别的团队卖`);
    return {
      weekShip: '', seaOrder: '', date: `充公${confiscateQty}个`, remark: '',
      handling: handlingRemarks.join(' | '), confiscate: confiscateQty,
      aiAdvice: stagnantMsg,
      _meta: { is_excess: isExcess, available_excess: Math.round(availableExcess), restock_excess_qty: restockExcessQty,
        is_deep: isDeep, is_high_volume: isHighVolume, is_stagnant: true, suggest_restock: Math.round(suggestRestock),
        actual_restock: Math.round(actualRestock), ptype, qty_per_box: qtyPerBox, daily_sales: daily30.toFixed(1),
        daily_weighted: dailyWeighted.toFixed(2), target_stock: targetStock, total_inv: totalInv,
        unshipped_orders: Math.round(unshippedOrders), week_ship: 0, monthly_sales: Math.round(monthlySales), asin, sku }
    };
  }

  // Core logic: centered on unshipped orders
  let weekShip = 0, shipDate = '', seaOrder = '', shipEstimateMsg = '';
  if (unshippedOrders > 0) {
    if (totalInv >= unshippedOrders) {
      weekShip = Math.round(unshippedOrders);
      shipDate = getNextFriday();
      seaOrder = weekShip;
    } else {
      weekShip = Math.round(totalInv);
      const remaining = unshippedOrders - totalInv;
      [shipDate, shipEstimateMsg] = estimateShipDate(remaining, daily30);
      if (weekShip > 0) seaOrder = weekShip;
    }
  }

  // AI advice
  const aiAdviceParts = [];
  if (unshippedOrders > 0) {
    if (dailyWeighted > 0) {
      const invDays = totalInv / dailyWeighted;
      if (invDays < 20) aiAdviceParts.push('【发货方式】建议发美森');
      else if (invDays < 30) aiAdviceParts.push('【发货方式】建议发快船');
      else if (invDays < 40) aiAdviceParts.push('【发货方式】建议发一半慢船一半美森');
      else aiAdviceParts.push('【发货方式】建议发慢船');
    } else if (totalInv > 0) {
      aiAdviceParts.push('【发货方式】销量数据不足，建议优先发美森');
    }
  }
  if (suggestRestock > 0) {
    const boxes = qtyPerBox > 0 ? Math.ceil(suggestRestock / qtyPerBox) : 0;
    aiAdviceParts.push(`【下单提醒】建议补货${suggestRestock}个（约${boxes}箱）`);
  }
  if (stagnantMsg) aiAdviceParts.push(stagnantMsg);
  const aiAdvice = aiAdviceParts.join(' | ');

  const handlingRemarks = [];
  if (isGreen) handlingRemarks.push('【绿标/翔标】该产品可能有绿标/翔标，请复制ASIN到详情页确认后再操作');
  if (shipEstimateMsg) handlingRemarks.push(shipEstimateMsg);
  const handling = handlingRemarks.join(' | ');
  const confiscate = (unshippedOrders > 0) ? Math.max(0, unshippedOrders - totalInv) : 0;

  return {
    weekShip: weekShip > 0 ? weekShip : '',
    seaOrder: seaOrder !== '' ? seaOrder : '',
    date: shipDate, remark: '', handling, confiscate: confiscate > 0 ? confiscate : '',
    aiAdvice, asin, sku,
    _meta: { is_excess: isExcess, available_excess: Math.round(availableExcess), restock_excess_qty: restockExcessQty,
      is_deep: isDeep, is_high_volume: isHighVolume, is_stagnant: isStagnant, suggest_restock: Math.round(suggestRestock),
      actual_restock: Math.round(actualRestock), ptype, qty_per_box: qtyPerBox, daily_sales: daily30.toFixed(1),
      daily_weighted: dailyWeighted.toFixed(2), target_stock: targetStock, total_inv: totalInv,
      unshipped_orders: Math.round(unshippedOrders), week_ship: Math.round(weekShip), monthly_sales: Math.round(monthlySales), asin, sku }
  };
}

// ========== Auth ==========
async function apiFetch(path, opts = {}) {
  const res = await fetch(API_BASE + path, {
    ...opts,
    headers: { 'Content-Type': 'application/json', ...(authToken ? { Authorization: 'Bearer ' + authToken } : {}), ...opts.headers },
  });
  if (res.status === 401) { logout(); return null; }
  return res;
}

async function doLogin() {
  const username = document.getElementById('login-username').value.trim();
  const password = document.getElementById('login-password').value.trim();
  if (!username || !password) { showLoginError('请输入用户名和密码'); return; }

  // Check version first
  try {
    const verRes = await fetch(API_BASE + '/api/check-version', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ version: APP_VERSION })
    });
    const verData = await verRes.json();
    if (!verData.allowed) { showLoginError('当前版本已被禁用，请联系管理员'); return; }
  } catch(e) { /* If API down, allow login */ }

  const res = await fetch(API_BASE + '/api/login', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ username, password })
  });
  const data = await res.json();
  if (!res.ok) { showLoginError(data.error || '登录失败'); return; }

  authToken = data.token;
  currentUser = { username: data.username, role: data.role };
  localStorage.setItem('token', authToken);
  document.getElementById('login-screen').style.display = 'none';
  document.getElementById('app').classList.add('active');
  document.getElementById('user-display').textContent = `${data.username} · ${data.role === 'admin' ? '管理员' : data.role === 'manager' ? '组长' : '运营'}`;
  document.getElementById('admin-btn').style.display = data.role === 'admin' ? 'inline-block' : 'none';
}

function showLoginError(msg) {
  const el = document.getElementById('login-error');
  el.textContent = msg;
  el.style.display = 'block';
}

function doLogout() {
  logout();
}

function logout() {
  authToken = null;
  currentUser = null;
  localStorage.removeItem('token');
  document.getElementById('login-screen').style.display = 'flex';
  document.getElementById('app').classList.remove('active');
  clearData();
}

// Auto login
if (authToken) {
  // We can't verify token without API, so just show login for now
  // In production, you'd call a /api/me endpoint
}

// ========== Excel Upload ==========
function handleFileSelect(event) {
  const file = event.target.files[0];
  if (!file) return;
  loadExcel(file);
}

function loadExcel(file) {
  const reader = new FileReader();
  reader.onload = function(e) {
    // Save original file buffer for ExcelJS format-preserving export
    originalFileBuffer = e.target.result.slice(0);
    const data = new Uint8Array(e.target.result);
    workbook = XLSX.read(data, { type: 'array', cellStyles: true });
    worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    if (jsonData.length < 2) { alert('表格数据为空'); return; }
    headers = jsonData[0];
    tableData = jsonData.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });

    // Detect green rows
    greenRows.clear();
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    for (let r = 1; r <= range.e.r; r++) {
      const cellRef = XLSX.utils.encode_cell({ r, c: 0 });
      if (isGreenCell(worksheet[cellRef])) greenRows.add(r - 1);
    }

    // Ensure calc columns exist
    const colMap = autoMapColumns(headers);
    let modified = false;
    if (!colMap.aiAdvice) { headers.push('AI建议'); modified = true; }
    if (!colMap.handling) { headers.push('处理建议'); modified = true; }
    if (modified) {
      tableData.forEach(row => {
        headers.forEach(h => { if (!(h in row)) row[h] = ''; });
      });
    }

    colIdxMap = {};
    for (const [key, header] of Object.entries(colMap)) {
      const idx = headers.indexOf(header);
      if (idx >= 0) colIdxMap[key] = idx;
    }

    renderTable();
    document.getElementById('toolbar').style.display = 'flex';
    document.getElementById('status-bar').style.display = 'flex';
    document.getElementById('table-container').style.display = 'block';
    document.getElementById('status-text').textContent = `已加载 ${file.name}`;
    document.getElementById('row-count').textContent = `${tableData.length} 行数据`;

    // Save history
    apiFetch('/api/history', { method: 'POST', body: JSON.stringify({ filename: file.name, row_count: tableData.length }) });
  };
  reader.readAsArrayBuffer(file);
}

// Drag & drop
document.addEventListener('DOMContentLoaded', () => {
  const uploadArea = document.getElementById('upload-area');
  uploadArea.addEventListener('dragover', (e) => { e.preventDefault(); uploadArea.classList.add('dragover'); });
  uploadArea.addEventListener('dragleave', () => uploadArea.classList.remove('dragover'));
  uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length && (files[0].name.endsWith('.xlsx') || files[0].name.endsWith('.xls'))) {
      loadExcel(files[0]);
    }
  });
});

// ========== Table Rendering ==========
function renderTable() {
  const thead = document.querySelector('#data-table thead');
  const tbody = document.querySelector('#data-table tbody');
  thead.innerHTML = '<tr>' + headers.map(h => `<th>${h}</th>`).join('') + '</tr>';

  tbody.innerHTML = tableData.map((row, rIdx) => {
    return '<tr>' + headers.map((h, cIdx) => {
      const val = row[h];
      let display = val !== undefined && val !== null ? String(val) : '';
      if (typeof val === 'number' && Number.isFinite(val)) display = Number.isInteger(val) ? val : val.toFixed(2);
      const isCalc = ['weekShip', 'seaOrder', 'date', 'handling', 'confiscate', 'aiAdvice'].some(k => colIdxMap[k] === cIdx);
      return `<td class="${isCalc ? 'calc-cell' : ''}" contenteditable="true">${display}</td>`;
    }).join('') + '</tr>';
  }).join('');
}

// ========== Auto Calculate ==========
function autoCalculate() {
  if (!tableData.length) return;
  const colMapByHeader = autoMapColumns(headers);
  const fullColMap = { ...colMapByHeader };

  seckillItems = [];
  for (let rIdx = 0; rIdx < tableData.length; rIdx++) {
    const row = tableData[rIdx];
    const rowData = {};
    for (const [key, colIdx] of Object.entries(colIdxMap)) {
      rowData[headers[colIdx]] = row[headers[colIdx]];
    }
    for (const fbaCol of fullColMap.fba_shipments || []) {
      if (headers.includes(fbaCol)) rowData[fbaCol] = row[fbaCol];
    }

    const isGreen = greenRows.has(rIdx);
    const result = analyzeRow(rowData, colMapByHeader, isGreen);

    if (result._meta.monthly_sales > 150) {
      seckillItems.push({ asin: result._meta.asin, sku: result._meta.sku, monthly_sales: result._meta.monthly_sales });
    }

    const calcMapping = { weekShip: result.weekShip, seaOrder: result.seaOrder, date: result.date,
      handling: result.handling, confiscate: result.confiscate, aiAdvice: result.aiAdvice };
    for (const [key, val] of Object.entries(calcMapping)) {
      if (key in colIdxMap) row[headers[colIdxMap[key]]] = val;
    }
  }

  renderTable();

  if (seckillItems.length) {
    document.getElementById('seckill-panel').style.display = 'block';
    document.getElementById('seckill-btn').style.display = 'inline-block';
    document.getElementById('seckill-items').innerHTML = seckillItems.map(item => `
      <div class="seckill-badge">
        <strong>${item.sku || 'N/A'}</strong>
        <span style="color:#FF9500;">${item.asin || 'N/A'}</span>
        <span style="color:#8E8E93; font-size:12px;">月销 ${item.monthly_sales}</span>
      </div>
    `).join('');
  }

  document.getElementById('status-text').textContent = '计算完成';
  alert(`自动计算完成！共处理 ${tableData.length} 行数据。`);
}

// ========== Export ==========
async function exportResult() {
  if (!tableData.length) { showToast('没有数据可导出'); return; }
  // Sync edits from DOM
  const tbody = document.querySelector('#data-table tbody');
  tbody.querySelectorAll('tr').forEach((tr, rIdx) => {
    tr.querySelectorAll('td').forEach((td, cIdx) => {
      tableData[rIdx][headers[cIdx]] = td.textContent;
    });
  });

  // Try ExcelJS first (preserves original formatting)
  if (originalFileBuffer && typeof ExcelJS !== 'undefined') {
    try {
      await exportWithExcelJS();
      return;
    } catch(e) {
      console.error('ExcelJS export failed:', e);
    }
  }
  // Fallback to SheetJS
  exportWithSheetJS();
}

async function exportWithExcelJS() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(originalFileBuffer);
  const ws = wb.worksheets[0];
  if (!ws) throw new Error('No worksheet found');

  // Read header row to find column positions
  const headerRow = ws.getRow(1);
  const headerMap = {}; // header text -> column number (1-based)
  const colCount = ws.columnCount;
  for (let c = 1; c <= colCount; c++) {
    const cell = headerRow.getCell(c);
    const val = cell.value;
    if (val !== null && val !== undefined) {
      // Handle rich text objects
      const text = (typeof val === 'object' && val.richText) 
        ? val.richText.map(r => r.text).join('') 
        : String(val).trim();
      if (text) headerMap[text] = c;
    }
  }

  // Define calculated columns and their aliases for matching
  const calcColAliases = {
    weekShip:  ['本周要出库库存', '本周要出', '本周出库', '本周要出库存'],
    seaOrder:  ['海运下单', '下单', '海运'],
    date:      ['改出货日期', '出货日期', '日期'],
    handling:  ['清货建议', '处理建议', '充公建议', '其他建议', '行动建议'],
    confiscate:['充公数量', '充公', '转其他团队'],
    aiAdvice:  ['AI建议', '发货建议', '渠道建议', '物流建议'],
  };

  // Map each calc key to a column number (1-based), finding or creating as needed
  const calcColNums = {};
  let nextNewCol = colCount + 1;

  for (const [key, aliases] of Object.entries(calcColAliases)) {
    let found = false;
    // First try exact match from colIdxMap headers
    if (key in colIdxMap) {
      const hdrName = String(headers[colIdxMap[key]]).trim();
      if (headerMap[hdrName]) {
        calcColNums[key] = headerMap[hdrName];
        found = true;
      }
    }
    // Then try aliases
    if (!found) {
      for (const alias of aliases) {
        if (headerMap[alias]) {
          calcColNums[key] = headerMap[alias];
          found = true;
          break;
        }
      }
    }
    // If still not found, also check by substring in existing headers
    if (!found) {
      for (const [hdr, colNum] of Object.entries(headerMap)) {
        for (const alias of aliases) {
          if (hdr.includes(alias) || alias.includes(hdr)) {
            calcColNums[key] = colNum;
            found = true;
            break;
          }
        }
        if (found) break;
      }
    }
    // Create new column if not found
    if (!found) {
      const newColNum = nextNewCol++;
      const displayName = (key in colIdxMap) ? String(headers[colIdxMap[key]]) : aliases[0];
      headerRow.getCell(newColNum).value = displayName;
      // Style the new header to match existing headers
      const refHeaderCell = headerRow.getCell(1);
      const headerCell = headerRow.getCell(newColNum);
      if (refHeaderCell.style) {
        headerCell.style = JSON.parse(JSON.stringify(refHeaderCell.style));
      }
      headerCell.font = { bold: true, ...(refHeaderCell.font || {}) };
      calcColNums[key] = newColNum;
    }
  }

  // Update only the calculated cells (leave all original cells untouched)
  for (let rIdx = 0; rIdx < tableData.length; rIdx++) {
    const excelRow = rIdx + 2; // Excel row (1-indexed, row 1 is header)
    const row = tableData[rIdx];
    for (const [key, colNum] of Object.entries(calcColNums)) {
      if (!(key in colIdxMap)) continue;
      const cell = ws.getCell(excelRow, colNum);
      let val = row[headers[colIdxMap[key]]];
      // Convert numeric strings to numbers
      if (typeof val === 'string') {
        const numVal = Number(val);
        if (!isNaN(numVal) && val.trim() !== '') val = numVal;
      }
      cell.value = (val !== '' && val !== null && val !== undefined) ? val : '';
    }
  }

  // Generate and download
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `备货分析_${new Date().toISOString().slice(0,10)}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  showToast('导出成功（保留原始格式）');
}

function exportWithSheetJS() {
  const ws = XLSX.utils.aoa_to_sheet([headers, ...tableData.map(row => headers.map(h => row[h]))]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  XLSX.writeFile(wb, `备货分析_${new Date().toISOString().slice(0,10)}.xlsx`);
  showToast('导出成功（SheetJS兼容模式）');
}

function showToast(msg) {
  // Simple toast notification
  let toast = document.getElementById('toast-notification');
  if (!toast) {
    toast = document.createElement('div');
    toast.id = 'toast-notification';
    toast.style.cssText = 'position:fixed;bottom:80px;left:50%;transform:translateX(-50%);background:rgba(0,0,0,0.8);color:white;padding:12px 24px;border-radius:12px;font-size:14px;z-index:9999;transition:opacity 0.3s;pointer-events:none;';
    document.body.appendChild(toast);
  }
  toast.textContent = msg;
  toast.style.opacity = '1';
  clearTimeout(toast._timer);
  toast._timer = setTimeout(() => { toast.style.opacity = '0'; }, 2500);
}

function clearData() {
  workbook = null; worksheet = null; headers = []; colIdxMap = {}; greenRows.clear();
  seckillItems = []; tableData = []; originalFileBuffer = null;
  document.getElementById('toolbar').style.display = 'none';
  document.getElementById('status-bar').style.display = 'none';
  document.getElementById('table-container').style.display = 'none';
  document.getElementById('seckill-panel').style.display = 'none';
  document.getElementById('seckill-btn').style.display = 'none';
  document.querySelector('#data-table thead').innerHTML = '';
  document.querySelector('#data-table tbody').innerHTML = '';
}

// ========== Seckill Dialog ==========
function openSeckill() {
  if (!seckillItems.length) return;
  document.getElementById('seckill-modal').classList.add('active');
  const tbody = document.querySelector('#seckill-table tbody');
  const defaultDate = getDefaultDateRange();
  tbody.innerHTML = seckillItems.map((item, i) => `
    <tr>
      <td contenteditable="true">${defaultDate}</td>
      <td><select><option>LD</option><option>BD</option><option>黑五LD</option><option>黑五BD</option><option>网一LD</option><option>网一BD</option><option>秋季</option></select></td>
      <td><select><option>美国</option><option>英国</option><option>德国</option><option>日本</option><option>加拿大</option></select></td>
      <td contenteditable="true">${item.asin}</td>
      <td contenteditable="true">${item.asin}</td>
      <td contenteditable="true">${item.sku}</td>
      <td contenteditable="true"></td>
      <td contenteditable="true"></td>
    </tr>
  `).join('');

  const ziniaoTbody = document.querySelector('#ziniao-table tbody');
  ziniaoTbody.innerHTML = Array(5).fill(0).map(() => `
    <tr>
      <td><select><option>紫鸟</option><option>VPS</option><option>向日葵</option></select></td>
      <td contenteditable="true"></td><td contenteditable="true"></td>
      <td contenteditable="true"></td><td contenteditable="true"></td>
    </tr>
  `).join('');
}

function getDefaultDateRange() {
  const d = new Date();
  const daysToMonday = (7 - d.getDay() + 1) % 7 || 7;
  const nextMon = new Date(d); nextMon.setDate(d.getDate() + daysToMonday);
  const nextFri = new Date(nextMon); nextFri.setDate(nextMon.getDate() + 4);
  return `${nextMon.getFullYear()}/${nextMon.getMonth()+1}/${nextMon.getDate()}-${nextFri.getFullYear()}/${nextFri.getMonth()+1}/${nextFri.getDate()}`;
}

async function saveSeckill() {
  const items = [];
  document.querySelectorAll('#seckill-table tbody tr').forEach(tr => {
    const tds = tr.querySelectorAll('td');
    items.push({ date: tds[0].textContent, type: tds[1].querySelector('select')?.value, site: tds[2].querySelector('select')?.value,
      asin: tds[3].textContent, child_asin: tds[4].textContent, sku: tds[5].textContent, qty: tds[6].textContent, shop: tds[7].textContent });
  });
  const ziniao = [];
  document.querySelectorAll('#ziniao-table tbody tr').forEach(tr => {
    const tds = tr.querySelectorAll('td');
    const row = { type: tds[0].querySelector('select')?.value, company: tds[1].textContent, username: tds[2].textContent, password: tds[3].textContent, shop: tds[4].textContent };
    if (Object.values(row).some(v => v)) ziniao.push(row);
  });

  await apiFetch('/api/seckill', { method: 'POST', body: JSON.stringify({ items, ziniao_info: ziniao }) });
  alert('秒杀提报表已保存');
  closeModal('seckill-modal');
}

function exportSeckill() {
  alert('导出功能开发中，请使用保存到系统后在管理后台导出。');
}

// ========== Admin Panel ==========
function openAdmin() {
  document.getElementById('admin-modal').classList.add('active');
  loadUsers();
}

function closeModal(id) {
  document.getElementById(id).classList.remove('active');
}

function switchTab(tab) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
  event.target.classList.add('active');
  document.getElementById('tab-' + tab).classList.add('active');
  if (tab === 'users') loadUsers();
  if (tab === 'logs') loadLogs();
  if (tab === 'history-admin') loadAdminHistory();
  if (tab === 'seckill-admin') loadAdminSeckill();
  if (tab === 'versions') loadVersions();
}

async function loadUsers() {
  const res = await apiFetch('/api/users');
  if (!res) return;
  const users = await res.json();
  const tbody = document.querySelector('#users-table tbody');
  const roleMap = { admin: '管理员', manager: '组长', user: '运营' };
  tbody.innerHTML = users.map(u => `
    <tr>
      <td>${u.id}</td><td>${u.username}</td><td><span class="badge badge-${u.role}">${roleMap[u.role] || u.role}</span></td>
      <td>${u.department || ''}</td><td>${u.is_active ? '启用' : '禁用'}</td>
      <td>
        <button class="btn btn-secondary" style="padding:4px 10px; font-size:12px;" onclick="toggleUser('${u.username}', ${!u.is_active})">${u.is_active ? '禁用' : '启用'}</button>
        <button class="btn btn-danger" style="padding:4px 10px; font-size:12px;" onclick="deleteUser('${u.username}')">删除</button>
      </td>
    </tr>
  `).join('');
}

async function loadLogs() {
  const res = await apiFetch('/api/logs');
  if (!res) return;
  const logs = await res.json();
  document.querySelector('#logs-table tbody').innerHTML = logs.map(l => `
    <tr><td>${l.id}</td><td>${l.username}</td><td>${l.login_time?.slice(0,19)}</td><td>${l.success ? '成功' : '失败'}</td></tr>
  `).join('');
}

async function loadAdminHistory() {
  const res = await apiFetch('/api/history');
  if (!res) return;
  const history = await res.json();
  document.querySelector('#history-admin-table tbody').innerHTML = history.map(h => `
    <tr><td>${h.id}</td><td>${h.username}</td><td>${h.filename}</td><td>${h.uploaded_at?.slice(0,19)}</td><td>${h.row_count}</td></tr>
  `).join('');
}

async function loadAdminSeckill() {
  const res = await apiFetch('/api/seckill');
  if (!res) return;
  const reports = await res.json();
  document.querySelector('#seckill-admin-table tbody').innerHTML = reports.map(r => `
    <tr><td>${r.id}</td><td>${r.username}</td><td>${r.created_at?.slice(0,19)}</td><td>${r.items ? JSON.parse(r.items).length : 0}</td></tr>
  `).join('');
}

async function loadVersions() {
  const res = await apiFetch('/api/versions');
  if (!res) return;
  const versions = await res.json();
  document.querySelector('#versions-table tbody').innerHTML = versions.map(v => `
    <tr>
      <td>${v.id}</td><td>${v.version}</td><td>${v.description || ''}</td><td>${v.is_active ? '启用' : '禁用'}</td>
      <td>
        <button class="btn btn-secondary" style="padding:4px 10px; font-size:12px;" onclick="toggleVersion(${v.id}, ${!v.is_active})">${v.is_active ? '禁用' : '启用'}</button>
        <button class="btn btn-danger" style="padding:4px 10px; font-size:12px;" onclick="deleteVersion(${v.id})">删除</button>
      </td>
    </tr>
  `).join('');
}

function showAddUser() { document.getElementById('add-user-modal').classList.add('active'); }
function showAddVersion() { document.getElementById('add-version-modal').classList.add('active'); }

async function doAddUser() {
  const username = document.getElementById('new-username').value.trim();
  const password = document.getElementById('new-password').value.trim();
  const role = document.getElementById('new-role').value;
  const department = document.getElementById('new-department').value.trim();
  if (!username || !password) { alert('用户名和密码不能为空'); return; }
  const res = await apiFetch('/api/register', { method: 'POST', body: JSON.stringify({ username, password, role, department }) });
  if (!res) return;
  const data = await res.json();
  if (!res.ok) { alert(data.error); return; }
  closeModal('add-user-modal');
  loadUsers();
}

async function doAddVersion() {
  const version = document.getElementById('new-version').value.trim();
  const description = document.getElementById('new-version-desc').value.trim();
  if (!version) { alert('请输入版本号'); return; }
  const res = await apiFetch('/api/versions', { method: 'POST', body: JSON.stringify({ version, description }) });
  if (!res) return;
  closeModal('add-version-modal');
  loadVersions();
}

async function toggleUser(username, active) {
  await apiFetch('/api/users/' + encodeURIComponent(username), { method: 'PUT', body: JSON.stringify({ active }) });
  loadUsers();
}

async function deleteUser(username) {
  if (!confirm('确定删除用户 ' + username + ' 吗？')) return;
  await apiFetch('/api/users/' + encodeURIComponent(username), { method: 'DELETE' });
  loadUsers();
}

async function toggleVersion(id, active) {
  await apiFetch('/api/versions', { method: 'PUT', body: JSON.stringify({ id, active }) });
  loadVersions();
}

async function deleteVersion(id) {
  if (!confirm('确定删除该版本吗？')) return;
  await apiFetch('/api/versions', { method: 'DELETE', body: JSON.stringify({ id }) });
  loadVersions();
}

// ========== History Dialog ==========
async function openHistory() {
  document.getElementById('history-modal').classList.add('active');
  const res = await apiFetch('/api/history');
  if (!res) return;
  const history = await res.json();
  document.querySelector('#history-table tbody').innerHTML = history.map(h => `
    <tr><td>${h.id}</td><td>${h.filename}</td><td>${h.uploaded_at?.slice(0,19)}</td><td>${h.row_count}</td></tr>
  `).join('');
}
