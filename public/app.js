// ============================================
// Amazon Inventory Assistant - Web Frontend
// Complete feature set with all fixes
// ============================================

const API_BASE = ''; // Same origin for Pages + Functions
const APP_VERSION = '1.0.1';

let authToken = localStorage.getItem('token');
let currentUser = null;
let workbook = null;
let worksheet = null;
let headers = [];
let colIdxMap = {};
let greenRows = new Set();
let seckillItems = [];
let tableData = [];
let batchFiles = [];
let manualColumnMappings = {}; // User's manual column mapping overrides
let originalFileName = ''; // Store original uploaded filename
let originalFileBuffer = null; // Store original file ArrayBuffer for ZIP-based export
let restockConfig = { weight_3d: 20, weight_7d: 30, weight_15d: 30, weight_30d: 20, stock_months: 4, deduct_unshipped: true, deduct_week_outbound: true };
let currentUserPermissions = null; // Current user's permissions object
let currentPermUser = null; // Username being edited in permission modal
let usersPermissionsMap = {}; // Cache of all users' permissions { username: permissionsObj }

// ========== Utility Functions ==========
function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = String(str || '');
  return div.innerHTML;
}

function showToast(msg, duration = 3000) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.classList.add('show');
  setTimeout(() => el.classList.remove('show'), duration);
}

function showProgress(pct) {
  const c = document.getElementById('progress-container');
  const b = document.getElementById('progress-bar');
  c.style.display = pct < 100 ? 'block' : 'none';
  b.style.width = pct + '%';
}

function parseNum(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  const cleaned = String(v).replace(/,/g, '').replace(/\s/g, '').replace(/[，]/g, '');
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
  if (dailySales <= 0) return [`${Math.round(remaining)}-新品待观察`, '【新品提醒】无历史销量数据，建议人工评估后少量备货测试市场，暂不充公'];

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

// Auxiliary calculation columns — matching original version output exactly
const AUX_COLUMNS = [
  { key: 'daily_3', name: '3天日均', format: v => (v !== undefined && v !== '' && !isNaN(v)) ? Number(Number(v).toFixed(2)).toString() : '' },
  { key: 'daily_7', name: '7天日均', format: v => (v !== undefined && v !== '' && !isNaN(v)) ? Number(Number(v).toFixed(2)).toString() : '' },
  { key: 'daily_15', name: '15天日均', format: v => (v !== undefined && v !== '' && !isNaN(v)) ? Number(Number(v).toFixed(2)).toString() : '' },
  { key: 'daily_30', name: '30天日均', format: v => (v !== undefined && v !== '' && !isNaN(v)) ? Number(Number(v).toFixed(2)).toString() : '' },
  { key: 'daily_weighted', name: '加权日销量', format: v => (v !== undefined && v !== '' && !isNaN(v)) ? Number(Number(v).toFixed(2)).toString() : '' },
  { key: 'monthly_weighted', name: '加权月销量', format: v => (v !== undefined && v !== '' && !isNaN(v)) ? Number(Number(v).toFixed(2)).toString() : '' },
  { key: 'stock_months', name: '备货月数', format: v => (v !== undefined && !isNaN(v)) ? Math.round(v) : '' },
  { key: 'target_stock', name: '目标库存', format: v => (v !== undefined && !isNaN(v)) ? Math.round(v) : '' },
  { key: 'fba_available', name: 'FBA现有库存', format: v => (v !== undefined && !isNaN(v)) ? Math.round(v) : '' },
  { key: 'transit_total', name: '在途库存合计', format: v => (v !== undefined && !isNaN(v)) ? Math.round(v) : '' },
  { key: 'purchasing_total', name: '采购中库存合计', format: v => (v !== undefined && !isNaN(v)) ? Math.round(v) : '' },
  { key: 'total_inv', name: '总库存', format: v => (v !== undefined && !isNaN(v)) ? Math.round(v) : '' },
  { key: 'recommended_order', name: '推荐下单库存', format: v => (v !== undefined && !isNaN(v)) ? Math.round(v) : '' },
  { key: 'inventory_months', name: '库存可售月数', format: v => (v !== undefined && v !== '' && !isNaN(v)) ? Number(Number(v).toFixed(2)).toString() : '' },
];

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

// ========== Column Mapping (with manual overrides) ==========
function autoMapColumns(hdrs) {
  const aliases = {
    asin: ['asin'],
    sku: ['sku', 'fnsku'],
    code: ['编码'],
    sales30: ['近30天销量', '30天销量', '月销量'],
    sales15: ['15日销量', '15天销量', '十五日销量'],
    sales7: ['7日销量', '7天销量', '七日销量'],
    sales3: ['3日销量', '3天销量', '三日销量'],
    yesterdaySales: ['昨日销量', '昨天销量', '1日销量', '1天销量', '近1天销量', '近1日销量'],
    monthlySales: ['月销量'],
    available: ['可售库存', '库存', '可用库存', 'fba库存', '现有库存'],
    unshippedOrders: ['未出货订单', '未出货'],
    weekOutbound: ['本周要出库库存', '本周要出', '本周出库', '本周要出库存'],
    purchasing: ['采购中', '采购中库存', '下单中', '已采购'],
    weekShip: ['本周要出库库存', '本周要出', '本周出库', '本周要出库存'],
    seaSup: ['海运补单', '海运补'],
    seaOrder: ['海运', '海运下单', '下单'],
    date: ['改出货日期', '出货日期', '日期'],
    remark: ['备注', '说明'],
    handling: ['清货建议', '处理建议', '充公建议', '其他建议', '行动建议'],
    confiscate: ['充公数量', '充公', '转其他团队'],
    aiAdvice: ['AI建议', '发货建议', '渠道建议', '物流建议'],
  };

  const colMap = {};
  const overrides = manualColumnMappings || {};

  // 1) Apply manual overrides for single-value columns first
  const singleKeys = ['available', 'unshippedOrders', 'weekOutbound', 'yesterdaySales',
    'sales3', 'sales7', 'sales15', 'sales30', 'asin', 'sku', 'code',
    'monthlySales', 'weekShip', 'seaOrder', 'date', 'remark', 'handling', 'confiscate', 'aiAdvice'];
  for (const key of singleKeys) {
    if (overrides[key] && overrides[key].length > 0) {
      for (const h of hdrs) {
        if (String(h) === overrides[key][0]) { colMap[key] = h; break; }
      }
    }
  }

  // 2) Apply manual overrides for multi-value columns
  if (Array.isArray(overrides.all_in_transit) && overrides.all_in_transit.length > 0) {
    const matched = hdrs.filter(h => overrides.all_in_transit.includes(String(h)));
    if (matched.length) colMap.all_in_transit = matched.map(String);
  }
  if (Array.isArray(overrides.all_purchasing) && overrides.all_purchasing.length > 0) {
    const matched = hdrs.filter(h => overrides.all_purchasing.includes(String(h)));
    if (matched.length) colMap.all_purchasing = matched.map(String);
  }

  // 3) Build ignore set
  const ignoreSet = new Set(overrides.ignore_columns || []);

  // 4) X月销量 regex
  for (const h of hdrs) {
    const hStr = h !== null && h !== undefined ? String(h) : '';
    if (/\d+月销量/.test(hStr) && !colMap.monthlySales) {
      colMap.monthlySales = hStr;
      break;
    }
  }

  // 5) Auto-detect remaining columns
  for (const [key, alts] of Object.entries(aliases)) {
    if (colMap[key]) continue;
    for (const a of alts) {
      const aLower = a.toLowerCase();
      for (const h of hdrs) {
        const hStr = h !== null && h !== undefined ? String(h) : '';
        if (hStr.toLowerCase() === aLower) { colMap[key] = hStr; break; }
      }
      if (colMap[key]) break;
      for (const h of hdrs) {
        const hStr = h !== null && h !== undefined ? String(h) : '';
        if (hStr.toLowerCase().includes(aLower)) {
          if (key === 'available') {
            const hLower = hStr.toLowerCase();
            const excludeKw = ['在途', '采购', '下单', '货件', 'shipment', '充公', '未出货', '出库', '推荐', '目标', '总计', '合计'];
            if (excludeKw.some(kw => hLower.includes(kw))) continue;
          }
          colMap[key] = hStr; break;
        }
      }
      if (colMap[key]) break;
    }
  }

  // 6) FBA shipment columns (multi)
  const fbaCols = [];
  for (const h of hdrs) {
    const hStr = h !== null && h !== undefined ? String(h) : '';
    if (ignoreSet.has(hStr)) continue;
    if (hStr.toUpperCase().includes('FBA') && hStr.length > 3) {
      // Exclude auxiliary/display columns like "FBA现有库存"
      const hLower = hStr.toLowerCase();
      if (['现有', '库存', '合计', '总计'].some(kw => hLower.includes(kw))) continue;
      fbaCols.push(hStr);
    }
  }
  if (fbaCols.length) colMap.fba_shipments = fbaCols;

  // 7) In-transit columns (multi) - supplement manual with auto-detect
  const transitCols = colMap.all_in_transit ? [...colMap.all_in_transit] : [];
  const inTransitKw = ['在途', '海运在途', '在途中', 'shipment', 'transit', '在海上'];
  for (const h of hdrs) {
    const hStr = h !== null && h !== undefined ? String(h) : '';
    if (ignoreSet.has(hStr) || transitCols.includes(hStr)) continue;
    const hLower = hStr.toLowerCase();
    if (inTransitKw.some(kw => hLower.includes(kw))) transitCols.push(hStr);
  }
  if (transitCols.length) colMap.all_in_transit = transitCols;

  // 8) Purchasing columns (multi) - supplement manual with auto-detect
  const purchCols = colMap.all_purchasing ? [...colMap.all_purchasing] : [];
  const purchKw = ['采购中', '采购', '下单中', '已采购', '已下单', 'purchasing'];
  for (const h of hdrs) {
    const hStr = h !== null && h !== undefined ? String(h) : '';
    if (ignoreSet.has(hStr) || purchCols.includes(hStr)) continue;
    const hLower = hStr.toLowerCase();
    if (purchKw.some(kw => hLower.includes(kw))) purchCols.push(hStr);
  }
  if (purchCols.length) colMap.all_purchasing = purchCols;

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
  // Sum in-transit inventory (avoid double-counting pre-computed totals with individual FBA columns)
  let unshipped = 0;
  const transitCols = colMap.all_in_transit || [];
  const fbaCols = colMap.fba_shipments || [];
  const transitTotalCol = transitCols.find(h => /合计|总计/.test(h));
  if (transitTotalCol && transitCols.length === 1) {
    // Pre-computed total column like "在途库存合计" — use directly
    unshipped = parseNum(row[transitTotalCol]);
  } else {
    // Sum individual in-transit and FBA shipment columns
    for (const col of transitCols) {
      unshipped += parseNum(row[col]);
    }
    for (const col of fbaCols) {
      unshipped += parseNum(row[col]);
    }
  }
  // Sum purchasing columns (deduplicate single key vs array)
  let purchasing = 0;
  const purchCols = colMap.all_purchasing || [];
  for (const col of purchCols) {
    purchasing += parseNum(row[col]);
  }
  if (colMap.purchasing && !purchCols.includes(colMap.purchasing)) {
    purchasing += parseNum(row[colMap.purchasing]);
  }
  const unshippedOrders = parseNum(row[colMap.unshippedOrders]);
  const weekOutbound = parseNum(row[colMap.weekOutbound]);
  const sku = row[colMap.sku] || '';
  const code = row[colMap.code] || '';
  const asin = row[colMap.asin] || '';

  const yesterdaySales = parseNum(row[colMap.yesterdaySales]);

  const daily30 = sales30 > 0 ? sales30 / 30 : 0;
  const daily7 = sales7 > 0 ? sales7 / 7 : daily30;
  const daily15 = sales15 > 0 ? sales15 / 15 : (sales7 > 0 ? daily7 : daily30);
  const daily3 = sales3 > 0 ? sales3 / 3 : (yesterdaySales > 0 ? yesterdaySales : daily7);
  const dailyWeighted = daily3 * (restockConfig.weight_3d / 100) + daily7 * (restockConfig.weight_7d / 100) + daily15 * (restockConfig.weight_15d / 100) + daily30 * (restockConfig.weight_30d / 100);

  // Apply deduct options from restock config
  let effectiveAvailable = available;
  if (restockConfig.deduct_unshipped && unshippedOrders > 0) {
    effectiveAvailable -= unshippedOrders;
  }
  if (restockConfig.deduct_week_outbound && weekOutbound > 0) {
    effectiveAvailable -= weekOutbound;
  }
  const totalInv = effectiveAvailable + unshipped + purchasing;
  const rawTotalInv = available + unshipped + purchasing; // For auxiliary column display (no deductions)
  const monthlyWeighted = dailyWeighted * 30; // Weighted monthly sales

  // New product detection: no historical sales data at all
  const isNewProduct = (monthlySales === 0 && dailyWeighted === 0);

  // Use stocking rules to determine multiplier (priority-sorted, first match wins)
  let stockMultiplier = restockConfig.stock_months || 4;
  const sortedRules = [...stockingRules].filter(r => r.is_active !== 0 && r.is_active !== false).sort((a, b) => (a.priority || 0) - (b.priority || 0));
  for (const rule of sortedRules) {
    const minSales = rule.min_monthly_sales || 0;
    const maxSales = rule.max_monthly_sales || 0; // 0 means no upper limit
    if (monthlySales >= minSales && (maxSales === 0 || monthlySales <= maxSales)) {
      stockMultiplier = rule.multiplier || rule.stock_multiplier || stockMultiplier;
      break;
    }
  }
  const isDeep = sales30 > 600;
  const targetStock = Math.round(monthlyWeighted * stockMultiplier);
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

  // New product early return: skip confiscation, show observation notice
  if (isNewProduct) {
    const handlingRemarks = [];
    if (isGreen) handlingRemarks.push('【绿标/翔标】该产品可能有绿标/翔标，请复制ASIN到详情页确认后再操作');
    handlingRemarks.push('🆕 新品待观察，建议人工评估首次备货量');
    // AI silence rule: only output advice if there's some data (transit/purchasing/unshipped)
    const hasAnyData = unshipped > 0 || purchasing > 0 || unshippedOrders > 0;
    const aiAdvice = hasAnyData ? '【新品】无历史销售数据，建议人工评估首次备货量' : '';
    return {
      weekShip: '', seaOrder: '', date: '新品待观察', remark: '',
      handling: handlingRemarks.join(' | '), confiscate: '',
      aiAdvice, asin, sku,
      _meta: { is_excess: false, available_excess: 0, restock_excess_qty: 0,
        is_deep: false, is_high_volume: false, is_stagnant: false,
        is_new_product: true, suggest_restock: 0,
        actual_restock: 0, ptype, qty_per_box: qtyPerBox, daily_sales: '0.0',
        daily_weighted: '0.00', target_stock: 0, total_inv: rawTotalInv,
        unshipped_orders: Math.round(unshippedOrders), week_ship: 0, monthly_sales: 0, asin, sku,
        daily_3: 0, daily_7: 0, daily_15: 0, daily_30: 0,
        monthly_weighted: 0, stock_months: stockMultiplier,
        fba_available: available, transit_total: unshipped, purchasing_total: purchasing,
        recommended_order: 0, inventory_months: 0 }
    };
  }

  let isStagnant = false;
  let stagnantMsg = '';
  if (sales30 > 0 && totalInv > sales30 * 6) {
    isStagnant = true;
    stagnantMsg = '【滞销提醒】总库存>30天销量6倍，建议加快动销';
  }

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
        daily_weighted: dailyWeighted.toFixed(2), target_stock: targetStock, total_inv: rawTotalInv,
        unshipped_orders: Math.round(unshippedOrders), week_ship: 0, monthly_sales: Math.round(monthlySales), asin, sku,
        daily_3: daily3, daily_7: daily7, daily_15: daily15, daily_30: daily30,
        monthly_weighted: monthlyWeighted, stock_months: stockMultiplier,
        fba_available: available, transit_total: unshipped, purchasing_total: purchasing,
        recommended_order: Math.round(suggestRestock), inventory_months: monthlyWeighted > 0 ? rawTotalInv / monthlyWeighted : 0 }
    };
  }

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
  // Seckill suggestion
  if (monthlySales > 150) {
    aiAdviceParts.push('【秒杀建议】月销量>150，建议每月报3次秒杀+1次周秒杀，考虑多备7天库存');
  }
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
      daily_weighted: dailyWeighted.toFixed(2), target_stock: targetStock, total_inv: rawTotalInv,
      unshipped_orders: Math.round(unshippedOrders), week_ship: Math.round(weekShip), monthly_sales: Math.round(monthlySales), asin, sku,
      daily_3: daily3, daily_7: daily7, daily_15: daily15, daily_30: daily30,
      monthly_weighted: monthlyWeighted, stock_months: stockMultiplier,
      fba_available: available, transit_total: unshipped, purchasing_total: purchasing,
      recommended_order: Math.round(suggestRestock), inventory_months: monthlyWeighted > 0 ? rawTotalInv / monthlyWeighted : 0 }
  };
}

// ========== Auth ==========
async function apiFetch(path, opts = {}) {
  const mergedHeaders = { 'Content-Type': 'application/json', ...(authToken ? { Authorization: 'Bearer ' + authToken } : {}), ...opts.headers };
  try {
    const res = await fetch(API_BASE + path, { ...opts, headers: mergedHeaders });
    if (res.status === 401) { logout(); return null; }
    return res;
  } catch (e) {
    console.error('API fetch error:', e);
    return null;
  }
}

async function doLogin() {
  const username = document.getElementById('login-username').value.trim();
  const password = document.getElementById('login-password').value.trim();
  if (!username || !password) { showLoginError('请输入用户名和密码'); return; }

  try {
    const verRes = await fetch(API_BASE + '/api/check-version', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ version: APP_VERSION })
    });
    const verData = await verRes.json();
    if (!verData.allowed) { showLoginError('当前版本已被禁用，请联系管理员'); return; }
  } catch(e) { /* Allow login if API down */ }

  try {
    const res = await fetch(API_BASE + '/api/login', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ username, password })
    });
    const data = await res.json();
    if (!res.ok) { showLoginError(data.error || '登录失败'); return; }

    authToken = data.token;
    currentUser = { username: data.username, role: data.role };
    currentUserPermissions = data.permissions || null;
    localStorage.setItem('token', authToken);
    // Load manual column mappings and restock config after login
    loadColumnMappingsFromServer();
    loadRestockConfig();
    loadStockingRules();
    showApp();
  } catch(e) {
    showLoginError('网络连接失败，请检查网络');
  }
}

function showApp() {
  if (!currentUser) {
    try {
      const parts = authToken.split('.');
      const payload = JSON.parse(atob(parts[1]));
      currentUser = { username: payload.username, role: payload.role };
    } catch(e) { logout(); return; }
  }
  document.getElementById('login-screen').style.display = 'none';
  document.getElementById('app').classList.add('active');
  const roleLabel = currentUser.role === 'admin' ? '管理员' : currentUser.role === 'manager' ? '组长' : '运营';
  document.getElementById('user-display').textContent = `${currentUser.username} · ${roleLabel}`;
  document.getElementById('admin-btn').style.display = currentUser.role === 'admin' ? 'inline-block' : 'none';
  // Show weighted algorithm button based on permissions (admin always has it)
  const weightBtn = document.getElementById('weight-btn');
  if (weightBtn) {
    if (currentUser.role === 'admin') {
      weightBtn.style.display = 'inline-block';
    } else {
      weightBtn.style.display = hasPermission('weight_config') ? 'inline-block' : 'none';
    }
  }
  // Apply permission-based UI restrictions (non-admin users)
  applyPermissionUI();
}

function applyPermissionUI() {
  if (!currentUser || currentUser.role === 'admin') return; // Admin sees everything
  const perms = currentUserPermissions;
  if (!perms) return; // No permissions set = full access
  // Map permission keys to toolbar button IDs
  const permBtnMap = {
    'upload': ['toolbar-upload-btn'],
    'column_settings': ['toolbar-colsettings-btn'],
    'batch_analysis': ['batch-btn', 'batch-btn-2'],
    'auto_calculate': ['toolbar-autocalc-btn'],
    'export': ['toolbar-export-btn'],
    'seckill': ['seckill-btn'],
    'history': ['history-btn'],
    'clear_data': ['toolbar-clear-btn'],
    'rules': ['toolbar-rules-btn'],
    'weight_config': ['weight-btn']
  };
  for (const [perm, btnIds] of Object.entries(permBtnMap)) {
    if (perms[perm] === false) {
      for (const id of btnIds) {
        const el = document.getElementById(id);
        if (el) el.style.display = 'none';
      }
    }
  }
  // Hide upload area if upload permission is denied
  if (perms.upload === false) {
    const uploadCard = document.getElementById('upload-card');
    if (uploadCard) uploadCard.style.display = 'none';
  }
}

function hasPermission(perm) {
  if (!currentUser) return false;
  if (currentUser.role === 'admin') return true;
  if (!currentUserPermissions) return true; // No restrictions set
  return currentUserPermissions[perm] !== false;
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

// Auto-login on page load
async function tryAutoLogin() {
  if (!authToken) return;
  try {
    const parts = authToken.split('.');
    if (parts.length !== 3) throw new Error('Invalid token');
    const payload = JSON.parse(atob(parts[1]));
    currentUser = { username: payload.username, role: payload.role };
    // Load column mappings and restock config
    loadColumnMappingsFromServer();
    loadRestockConfig();
    loadStockingRules();
    // Verify token with a simple API call
    const res = await apiFetch('/api/history');
    if (res && res.ok) {
      showApp();
    } else {
      logout();
    }
  } catch(e) {
    logout();
  }
}
document.addEventListener('DOMContentLoaded', tryAutoLogin);

// ========== Column Settings (Manual Mapping) ==========
async function loadColumnMappingsFromServer() {
  try {
    const res = await apiFetch('/api/column-mappings');
    if (res && res.ok) {
      manualColumnMappings = await res.json();
    }
  } catch(e) { /* ignore */ }
}

async function openColumnSettings() {
  document.getElementById('col-settings-modal').classList.add('active');
  // Populate select options from current headers (if table loaded) or from saved mappings
  const allHeaders = headers.length > 0 ? [...headers] : [];
  // If no table loaded, try to get headers from last session
  const comboIds = ['cs-available', 'cs-sales3', 'cs-sales7', 'cs-sales15', 'cs-sales30',
    'cs-yesterday', 'cs-unshipped', 'cs-week-outbound'];
  for (const id of comboIds) {
    const sel = document.getElementById(id);
    const currentVal = sel.value;
    sel.innerHTML = '<option value="">（自动识别）</option>';
    for (const h of allHeaders) {
      if (h) {
        const opt = document.createElement('option');
        opt.value = h; opt.textContent = h;
        sel.appendChild(opt);
      }
    }
    sel.value = currentVal || '';
  }

  // Populate multi-select chips
  const multiIds = ['cs-transit', 'cs-purchasing', 'cs-ignore'];
  for (const id of multiIds) {
    const container = document.getElementById(id);
    container.innerHTML = '';
    for (const h of allHeaders) {
      if (h) {
        const chip = document.createElement('span');
        chip.className = 'chip';
        chip.textContent = h;
        chip.dataset.value = h;
        chip.onclick = () => chip.classList.toggle('selected');
        container.appendChild(chip);
      }
    }
  }

  // Load saved mappings and apply to UI
  const m = manualColumnMappings || {};
  const comboMap = {
    'cs-available': 'available', 'cs-sales3': 'sales3', 'cs-sales7': 'sales7',
    'cs-sales15': 'sales15', 'cs-sales30': 'sales30', 'cs-yesterday': 'yesterdaySales',
    'cs-unshipped': 'unshippedOrders', 'cs-week-outbound': 'weekOutbound'
  };
  for (const [selId, mapKey] of Object.entries(comboMap)) {
    if (m[mapKey] && m[mapKey].length > 0) {
      document.getElementById(selId).value = m[mapKey][0] || '';
    }
  }
  // Multi-selects
  const multiMap = { 'cs-transit': 'all_in_transit', 'cs-purchasing': 'all_purchasing', 'cs-ignore': 'ignore_columns' };
  for (const [containerId, mapKey] of Object.entries(multiMap)) {
    const selected = m[mapKey] || [];
    const container = document.getElementById(containerId);
    for (const chip of container.querySelectorAll('.chip')) {
      if (selected.includes(chip.dataset.value)) chip.classList.add('selected');
    }
  }
}

async function saveColumnSettings() {
  const comboMap = {
    'cs-available': 'available', 'cs-sales3': 'sales3', 'cs-sales7': 'sales7',
    'cs-sales15': 'sales15', 'cs-sales30': 'sales30', 'cs-yesterday': 'yesterdaySales',
    'cs-unshipped': 'unshippedOrders', 'cs-week-outbound': 'weekOutbound'
  };
  const mappings = {};
  for (const [selId, mapKey] of Object.entries(comboMap)) {
    const val = document.getElementById(selId).value;
    mappings[mapKey] = val ? [val] : [];
  }
  // Multi-selects
  const multiMap = { 'cs-transit': 'all_in_transit', 'cs-purchasing': 'all_purchasing', 'cs-ignore': 'ignore_columns' };
  for (const [containerId, mapKey] of Object.entries(multiMap)) {
    const selected = [];
    for (const chip of document.getElementById(containerId).querySelectorAll('.chip.selected')) {
      selected.push(chip.dataset.value);
    }
    mappings[mapKey] = selected;
  }

  const res = await apiFetch('/api/column-mappings', { method: 'POST', body: JSON.stringify(mappings) });
  if (res && res.ok) {
    manualColumnMappings = mappings;
    showToast('列映射设置已保存，下次上传表格时将优先使用手动指定的列');
    closeModal('col-settings-modal');
  } else {
    showToast('保存失败');
  }
}

async function clearColumnSettings() {
  const allMappings = {
    available: [], sales3: [], sales7: [], sales15: [], sales30: [],
    yesterdaySales: [], unshippedOrders: [], weekOutbound: [],
    all_in_transit: [], all_purchasing: [], ignore_columns: []
  };
  const res = await apiFetch('/api/column-mappings', { method: 'POST', body: JSON.stringify(allMappings) });
  if (res && res.ok) {
    manualColumnMappings = {};
    // Reset UI
    for (const sel of document.querySelectorAll('#col-settings-form select')) sel.value = '';
    for (const chip of document.querySelectorAll('#col-settings-form .chip')) chip.classList.remove('selected');
    showToast('已清除手动设置，恢复自动识别模式');
  }
}

// ========== Excel Upload ==========
function handleFileSelect(event) {
  const file = event.target.files[0];
  if (!file) return;
  event.target.value = ''; // Reset so same file can be re-selected
  loadExcel(file);
}

function loadExcel(file) {
  originalFileName = file.name; // Store for export naming
  showProgress(10);
  const reader = new FileReader();
  reader.onload = async function(e) {
    showProgress(30);
    originalFileBuffer = e.target.result.slice(0); // Save copy for ZIP-based export
    const data = new Uint8Array(e.target.result);
    workbook = XLSX.read(data, { type: 'array', cellStyles: true });
    worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    if (jsonData.length < 2) { alert('表格数据为空'); showProgress(100); return; }
    headers = jsonData[0].map(h => String(h || '').trim());
    tableData = jsonData.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? row[i] : ''; });
      return obj;
    });

    showProgress(50);

    // Detect green rows
    greenRows.clear();
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    for (let r = 1; r <= range.e.r; r++) {
      const cellRef = XLSX.utils.encode_cell({ r, c: 0 });
      if (isGreenCell(worksheet[cellRef])) greenRows.add(r - 1);
    }

    const colMap = autoMapColumns(headers);
    let modified = false;
    if (!colMap.aiAdvice) { headers.push('AI建议'); colMap.aiAdvice = 'AI建议'; modified = true; }
    if (!colMap.handling) { headers.push('处理建议'); colMap.handling = '处理建议'; modified = true; }
    // Ensure auxiliary calculation columns exist
    for (const aux of AUX_COLUMNS) {
      if (headers.indexOf(aux.name) < 0) { headers.push(aux.name); modified = true; }
    }
    if (modified) {
      tableData.forEach(row => {
        headers.forEach(h => { if (!(h in row)) row[h] = ''; });
      });
    }

    colIdxMap = {};
    for (const [key, header] of Object.entries(colMap)) {
      if (key === 'fba_shipments' || key === 'all_in_transit' || key === 'all_purchasing') {
        if (Array.isArray(header)) {
          colIdxMap[key] = header.map(h => headers.indexOf(h));
        }
      } else {
        const idx = headers.indexOf(header);
        if (idx >= 0) colIdxMap[key] = idx;
      }
    }

    showProgress(80);
    renderTable();
    document.getElementById('toolbar').style.display = 'flex';
    document.getElementById('status-bar').style.display = 'flex';
    document.getElementById('table-container').style.display = 'block';
    document.getElementById('status-text').textContent = `已加载 ${file.name}`;
    document.getElementById('row-count').textContent = `${tableData.length} 行数据`;

    // Show batch button when table is loaded (with permission check)
    const batchBtn = document.getElementById('batch-btn');
    if (batchBtn) batchBtn.style.display = hasPermission('batch_analysis') ? 'inline-block' : 'none';
    const batchBtn2 = document.getElementById('batch-btn-2');
    if (batchBtn2) batchBtn2.style.display = 'none';

    // Save upload history AND file data to server for admin online viewing
    try {
      const histRes = await apiFetch('/api/history', { method: 'POST', body: JSON.stringify({ filename: file.name, row_count: tableData.length }) });
      if (histRes && histRes.ok) {
        const histData = await histRes.json();
        if (histData.id) {
          // Save parsed table data for online viewing
          await apiFetch('/api/upload-data', {
            method: 'POST',
            body: JSON.stringify({
              history_id: histData.id,
              headers: headers,
              data: tableData,
              green_rows: Array.from(greenRows)
            })
          });
        }
      }
    } catch(err) { /* non-critical */ }

    showProgress(100);
  };
  reader.readAsArrayBuffer(file);
}

// Drag & drop
document.addEventListener('DOMContentLoaded', () => {
  const uploadArea = document.getElementById('upload-area');
  if (uploadArea) {
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
  }
});

// ========== Table Rendering ==========
function renderTable() {
  const thead = document.querySelector('#data-table thead');
  const tbody = document.querySelector('#data-table tbody');

  const colMap = autoMapColumns(headers);
  const skuLikeKeys = new Set(['sku', 'asin', 'code']);
  const numLikeKeys = new Set(['sales30', 'sales15', 'sales7', 'sales3', 'monthlySales', 'available', 'unshipped', 'purchasing', 'unshippedOrders', 'yesterdaySales', 'weekOutbound']);
  const calcKeys = new Set(['weekShip', 'seaOrder', 'date', 'handling', 'confiscate', 'aiAdvice']);

  const colTypes = headers.map(h => {
    for (const [key, hdr] of Object.entries(colMap)) {
      if (hdr === h) {
        if (skuLikeKeys.has(key)) return 'sku';
        if (numLikeKeys.has(key)) return 'num';
        if (calcKeys.has(key)) return 'calc';
      }
    }
    const vals = tableData.slice(0, 10).map(r => r[h]);
    const numCount = vals.filter(v => !isNaN(parseFloat(v)) && v !== '').length;
    if (numCount > vals.length * 0.5) return 'num';
    return 'text';
  });

  thead.innerHTML = '<tr>' + headers.map(h => `<th>${escapeHtml(h) || '&nbsp;'}</th>`).join('') + '</tr>';

  tbody.innerHTML = tableData.map((row, rIdx) => {
    const isGreen = greenRows.has(rIdx);
    const rowClass = isGreen ? 'green-row' : '';
    return `<tr class="${rowClass}">` + headers.map((h, cIdx) => {
      const val = row[h];
      let display = val !== undefined && val !== null ? String(val) : '';
      if (typeof val === 'number' && Number.isFinite(val)) display = Number.isInteger(val) ? val : val.toFixed(2);

      const cType = colTypes[cIdx];
      let cls = '';
      if (cType === 'calc') cls = 'calc-cell';
      else if (cType === 'sku') cls = 'sku-cell';
      else if (cType === 'num') cls = 'num-cell';

      return `<td class="${cls}" contenteditable="true">${display}</td>`;
    }).join('') + '</tr>';
  }).join('');
}

// ========== Auto Calculate ==========
function autoCalculate() {
  if (!tableData.length) { showToast('请先上传表格数据'); return; }
  const colMapByHeader = autoMapColumns(headers);

  seckillItems = [];
  for (let rIdx = 0; rIdx < tableData.length; rIdx++) {
    const row = tableData[rIdx];
    const rowData = {};
    for (const [key, colIdx] of Object.entries(colIdxMap)) {
      if (key === 'fba_shipments' || key === 'all_in_transit' || key === 'all_purchasing') {
        if (Array.isArray(colIdx)) {
          for (const idx of colIdx) {
            if (idx >= 0 && idx < headers.length) rowData[headers[idx]] = row[headers[idx]];
          }
        }
      } else if (colIdx >= 0 && colIdx < headers.length) {
        rowData[headers[colIdx]] = row[headers[colIdx]];
      }
    }
    for (const fbaCol of colMapByHeader.fba_shipments || []) {
      if (headers.includes(fbaCol)) rowData[fbaCol] = row[fbaCol];
    }
    for (const tCol of colMapByHeader.all_in_transit || []) {
      if (headers.includes(tCol)) rowData[tCol] = row[tCol];
    }
    for (const pCol of colMapByHeader.all_purchasing || []) {
      if (headers.includes(pCol)) rowData[pCol] = row[pCol];
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
    // Write auxiliary calculation columns
    for (const aux of AUX_COLUMNS) {
      const colIdx = headers.indexOf(aux.name);
      if (colIdx >= 0) {
        row[headers[colIdx]] = aux.format(result._meta[aux.key]);
      }
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
        <span style="color:#8E8E93; font-size:11px;">月销 ${item.monthly_sales}</span>
      </div>
    `).join('');
  }

  document.getElementById('status-text').textContent = '计算完成';
  showToast(`自动计算完成！共处理 ${tableData.length} 行数据，${seckillItems.length} 个秒杀推荐`);
}

// ========== Batch Analysis ==========
function batchUpload() {
  document.getElementById('batch-file-input').click();
}

async function handleBatchSelect(event) {
  const files = Array.from(event.target.files);
  if (!files.length) return;
  event.target.value = '';

  showToast(`开始批量分析 ${files.length} 个文件...`);
  let successCount = 0;
  let failCount = 0;

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    showProgress(Math.round((i / files.length) * 100));
    try {
      const result = await processBatchFile(file);
      if (result) {
        downloadBatchResult(result, file.name);
        successCount++;
      } else {
        failCount++;
      }
    } catch(e) {
      console.error('Batch file error:', file.name, e);
      failCount++;
    }
  }

  showProgress(100);
  showToast(`批量分析完成！成功 ${successCount} 个，失败 ${failCount} 个`);
}

function processBatchFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = function(e) {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array', cellStyles: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        if (jsonData.length < 2) { resolve(null); return; }

        const hdrs = jsonData[0].map(h => String(h || '').trim());
        const rows = jsonData.slice(1).map(row => {
          const obj = {};
          hdrs.forEach((h, i) => { obj[h] = row[i] !== undefined ? row[i] : ''; });
          return obj;
        });

        const colMap = autoMapColumns(hdrs);
        let modified = false;
        if (!colMap.aiAdvice) { hdrs.push('AI建议'); modified = true; }
        if (!colMap.handling) { hdrs.push('处理建议'); modified = true; }
        for (const aux of AUX_COLUMNS) {
          if (hdrs.indexOf(aux.name) < 0) { hdrs.push(aux.name); modified = true; }
        }
        if (modified) {
          rows.forEach(row => hdrs.forEach(h => { if (!(h in row)) row[h] = ''; }));
        }

        const greens = new Set();
        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
        for (let r = 1; r <= range.e.r; r++) {
          const cellRef = XLSX.utils.encode_cell({ r, c: 0 });
          if (isGreenCell(ws[cellRef])) greens.add(r - 1);
        }

        for (let rIdx = 0; rIdx < rows.length; rIdx++) {
          const result = analyzeRow(rows[rIdx], colMap, greens.has(rIdx));
          const calcMapping = { weekShip: result.weekShip, seaOrder: result.seaOrder, date: result.date,
            handling: result.handling, confiscate: result.confiscate, aiAdvice: result.aiAdvice };
          for (const [key, val] of Object.entries(calcMapping)) {
            if (key in colMap) rows[rIdx][colMap[key]] = val;
          }
          // Write auxiliary columns
          for (const aux of AUX_COLUMNS) {
            if (hdrs.indexOf(aux.name) >= 0) {
              rows[rIdx][aux.name] = aux.format(result._meta[aux.key]);
            }
          }
        }

        resolve({ headers: hdrs, rows });
      } catch(err) {
        resolve(null);
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

function downloadBatchResult(result, origName) {
  const ws = XLSX.utils.aoa_to_sheet([result.headers, ...result.rows.map(row => result.headers.map(h => row[h]))]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  const newName = origName.replace(/\.xlsx?$/i, '') + '_已填充.xlsx';
  XLSX.writeFile(wb, newName);
}

// ========== Export ==========
function exportResult() {
  if (!worksheet || !workbook || !originalFileBuffer) { showToast('没有数据可导出'); return; }
  if (typeof JSZip === 'undefined') { showToast('缺少JSZip库，无法导出'); return; }

  // ZIP-based export: preserve original format/styles/colors/formulas
  // Only update specific AI result columns (本周要出库存, 改出货日期)
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

  async function doExport() {
    try {
      showToast('正在生成导出文件...');
      const zip = await JSZip.loadAsync(originalFileBuffer);

      // Find sheet XML path
      const sheetPath = zip.file('xl/worksheets/sheet1.xml') ? 'xl/worksheets/sheet1.xml' :
        Object.keys(zip.files).find(f => f.match(/xl\/worksheets\/sheet\d+\.xml/)) || 'xl/worksheets/sheet1.xml';
      const sheetXml = await zip.file(sheetPath).async('string');

      const parser = new DOMParser();
      const doc = parser.parseFromString(sheetXml, 'text/xml');

      // Load shared strings
      let ssStrings = [];
      const ssFile = zip.file('xl/sharedStrings.xml');
      if (ssFile) {
        const ssXml = await ssFile.async('string');
        const ssDoc = parser.parseFromString(ssXml, 'text/xml');
        const sis = ssDoc.getElementsByTagName('si');
        for (let i = 0; i < sis.length; i++) {
          let text = '';
          const ts = sis[i].getElementsByTagName('t');
          for (let j = 0; j < ts.length; j++) text += (ts[j].textContent || '');
          ssStrings.push(text);
        }
      }

      // Read header row to build column name → column index mapping
      const allRows = doc.getElementsByTagName('row');
      if (!allRows.length) { showToast('工作表为空'); return; }
      const headerRow = allRows[0];
      const hCells = headerRow.getElementsByTagName('c');
      const colNameMap = {}; // headerName → 0-based col index
      let actualMaxCol = -1; // Track actual rightmost column (including empty ones)
      const emptyCols = []; // Columns with no header text (can be reused)
      
      for (let i = 0; i < hCells.length; i++) {
        const cell = hCells[i];
        const ref = cell.getAttribute('r') || '';
        const colLetters = ref.replace(/\d+/g, '');
        let ci = 0;
        for (let k = 0; k < colLetters.length; k++) ci = ci * 26 + colLetters.charCodeAt(k) - 64;
        ci--;
        if (ci > actualMaxCol) actualMaxCol = ci;
        
        let val = '';
        if (cell.getAttribute('t') === 's') {
          const vEl = cell.getElementsByTagName('v')[0];
          if (vEl) val = ssStrings[parseInt(vEl.textContent)] || '';
        } else {
          const vEl = cell.getElementsByTagName('v')[0];
          if (vEl) val = vEl.textContent || '';
        }
        val = String(val).trim();
        if (val) {
          colNameMap[val] = ci;
        } else {
          // Empty header cell - can be reused
          emptyCols.push(ci);
        }
      }

      // Determine target columns to update (only 本周要出库存 and 改出货日期)
      const targetHeaders = ['本周要出库存', '改出货日期'];
      const targets = {};
      for (const h of targetHeaders) {
        if (colNameMap[h] !== undefined) {
          targets[colNameMap[h]] = h;
        }
      }

      // Add columns to the right side in order:
      // 1) 处理建议, AI建议 (right after 备注)
      // 2) AUX_COLUMNS (3天日均 etc. after 处理建议/AI建议)
      // First, try to reuse empty columns, then create new ones
      const auxColMap = {}; // aux column name -> column index
      let nextNewCol = actualMaxCol + 1;
      let emptyIdx = 0;

      // Helper: assign a column index for a new export column
      function assignCol(name) {
        if (colNameMap[name] !== undefined) {
          auxColMap[name] = colNameMap[name];
        } else if (emptyIdx < emptyCols.length) {
          auxColMap[name] = emptyCols[emptyIdx++];
        } else {
          auxColMap[name] = nextNewCol++;
        }
      }

      // 1) 处理建议 and AI建议 come first (between 备注 and calculation columns)
      const extraHeaders = ['处理建议', 'AI建议'];
      for (const eh of extraHeaders) {
        if (auxColMap[eh] === undefined) assignCol(eh);
      }

      // 2) AUX_COLUMNS (3天日均, 7天日均, ...) come after
      for (let i = 0; i < AUX_COLUMNS.length; i++) {
        const auxName = AUX_COLUMNS[i].name;
        if (auxColMap[auxName] === undefined) assignCol(auxName);
      }

      if (Object.keys(targets).length === 0 && Object.keys(auxColMap).length === 0) {
        showToast('未找到目标列，请确认表格包含"本周要出库存"和"改出货日期"列');
        return;
      }

      // Helper: convert 0-based column index to Excel column letters (A, B, ..., Z, AA, AB, ...)
      function colIndexToLetters(ci) {
        let s = '';
        ci++;
        while (ci > 0) {
          ci--;
          s = String.fromCharCode(65 + (ci % 26)) + s;
          ci = Math.floor(ci / 26);
        }
        return s;
      }

      // Add/update auxiliary column headers in headerRow
      // Build a map of existing cells by column index for quick lookup
      const headerCellMap = {};
      for (let i = 0; i < hCells.length; i++) {
        const cell = hCells[i];
        const ref = cell.getAttribute('r') || '';
        const colLetters = ref.replace(/\d+/g, '');
        let ci = 0;
        for (let k = 0; k < colLetters.length; k++) ci = ci * 26 + colLetters.charCodeAt(k) - 64;
        ci--;
        headerCellMap[ci] = cell;
      }
      
      for (const [auxName, ci] of Object.entries(auxColMap)) {
        const cellRef = colIndexToLetters(ci) + '1';
        let cell = headerCellMap[ci];
        
        if (cell) {
          // Update existing empty cell - change to inlineStr with header name
          cell.setAttribute('t', 'inlineStr');
          cell.removeAttribute('s'); // Remove old style reference
          // Remove existing <v> if present
          const vEl = cell.getElementsByTagName('v')[0];
          if (vEl) vEl.parentNode.removeChild(vEl);
          // Create or update <is> element
          let isEl = cell.getElementsByTagName('is')[0];
          if (!isEl) {
            isEl = doc.createElementNS(NS, 'is');
            cell.appendChild(isEl);
          }
          let tEl = isEl.getElementsByTagName('t')[0];
          if (!tEl) {
            tEl = doc.createElementNS(NS, 't');
            isEl.appendChild(tEl);
          }
          tEl.textContent = auxName;
        } else {
          // Create new cell
          const newCell = doc.createElementNS(NS, 'c');
          newCell.setAttribute('r', cellRef);
          newCell.setAttribute('t', 'inlineStr');
          const isEl = doc.createElementNS(NS, 'is');
          const tEl = doc.createElementNS(NS, 't');
          tEl.textContent = auxName;
          isEl.appendChild(tEl);
          newCell.appendChild(isEl);
          headerRow.appendChild(newCell);
        }
      }

      // Process data rows (skip header row 0)
      for (let ri = 1; ri < allRows.length; ri++) {
        const dataRow = allRows[ri];
        const rIdx = ri - 1;
        if (rIdx >= tableData.length) break;
        const cells = Array.from(dataRow.getElementsByTagName('c'));

        for (const cell of cells) {
          const ref = cell.getAttribute('r') || '';
          const colLetters = ref.replace(/\d+/g, '');
          let ci = 0;
          for (let k = 0; k < colLetters.length; k++) ci = ci * 26 + colLetters.charCodeAt(k) - 64;
          ci--;
          if (!(ci in targets)) continue;

          const hdr = targets[ci];
          const val = tableData[rIdx][hdr];
          if (val === undefined || val === null || val === '') continue;

          const numVal = parseNum(val);
          const isNum = !isNaN(numVal) && String(numVal) === String(val).trim();
          const vEl = cell.getElementsByTagName('v')[0];

          if (isNum) {
            cell.removeAttribute('t');
            if (vEl) { vEl.textContent = String(numVal); }
            else {
              const v = doc.createElementNS(NS, 'v');
              v.textContent = String(numVal);
              cell.appendChild(v);
            }
          } else {
            // Text: use inlineStr to avoid modifying shared string table
            const sv = String(val);
            cell.setAttribute('t', 'inlineStr');
            // Remove existing <v> if present
            if (vEl) vEl.parentNode.removeChild(vEl);
            // Create or update <is> element
            let isEl = cell.getElementsByTagName('is')[0];
            if (!isEl) {
              isEl = doc.createElementNS(NS, 'is');
              cell.appendChild(isEl);
            }
            let tEl = isEl.getElementsByTagName('t')[0];
            if (!tEl) {
              tEl = doc.createElementNS(NS, 't');
              isEl.appendChild(tEl);
            }
            tEl.textContent = sv;
          }
        }

        // Add auxiliary column data cells for this row
        // Build a map of existing cells in this row by column index
        const rowDataCellMap = {};
        for (const cell of cells) {
          const ref = cell.getAttribute('r') || '';
          const colLetters = ref.replace(/\d+/g, '');
          let ci = 0;
          for (let k = 0; k < colLetters.length; k++) ci = ci * 26 + colLetters.charCodeAt(k) - 64;
          ci--;
          rowDataCellMap[ci] = cell;
        }

        for (const [auxName, ci] of Object.entries(auxColMap)) {
          const val = tableData[rIdx][auxName];
          if (val === undefined || val === null || val === '') continue;

          const rowNum = ri + 1; // Excel row numbers are 1-based
          const cellRef = colIndexToLetters(ci) + rowNum;
          const numVal = parseNum(val);
          const isNum = !isNaN(numVal) && String(numVal) === String(val).trim();

          let cell = rowDataCellMap[ci];
          if (cell) {
            // Update existing cell (reused empty column)
            // Remove old type and value
            cell.removeAttribute('t');
            const oldV = cell.getElementsByTagName('v')[0];
            if (oldV) oldV.parentNode.removeChild(oldV);
            const oldIS = cell.getElementsByTagName('is')[0];
            if (oldIS) oldIS.parentNode.removeChild(oldIS);

            if (isNum) {
              const vEl = doc.createElementNS(NS, 'v');
              vEl.textContent = String(numVal);
              cell.appendChild(vEl);
            } else {
              cell.setAttribute('t', 'inlineStr');
              const isEl = doc.createElementNS(NS, 'is');
              const tEl = doc.createElementNS(NS, 't');
              tEl.textContent = String(val);
              isEl.appendChild(tEl);
              cell.appendChild(isEl);
            }
          } else {
            // Create new cell (brand new column)
            const newCell = doc.createElementNS(NS, 'c');
            newCell.setAttribute('r', cellRef);

            if (isNum) {
              const vEl = doc.createElementNS(NS, 'v');
              vEl.textContent = String(numVal);
              newCell.appendChild(vEl);
            } else {
              newCell.setAttribute('t', 'inlineStr');
              const isEl = doc.createElementNS(NS, 'is');
              const tEl = doc.createElementNS(NS, 't');
              tEl.textContent = String(val);
              isEl.appendChild(tEl);
              newCell.appendChild(isEl);
            }
            dataRow.appendChild(newCell);
          }
        }
      }

      // Update dimension element to include new auxiliary columns
      const maxColFinal = nextNewCol - 1;
      if (Object.keys(auxColMap).length > 0) {
        const dimensionEl = doc.querySelector('dimension');
        if (dimensionEl) {
          const oldRef = dimensionEl.getAttribute('ref');
          const match = oldRef.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
          if (match) {
            const newEndCol = colIndexToLetters(maxColFinal);
            const newRef = `${match[1]}${match[2]}:${newEndCol}${match[4]}`;
            dimensionEl.setAttribute('ref', newRef);
          }
        }
      }

      // Serialize sheet XML
      const serializer = new XMLSerializer();
      let newSheetXml = serializer.serializeToString(doc);
      zip.file(sheetPath, newSheetXml);

      // Generate and download
      const blob = await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      let exportName = originalFileName ? originalFileName.replace(/\.xlsx?$/i, '') + '_已填充.xlsx' : '备货分析_已填充.xlsx';
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = exportName;
      document.body.appendChild(a); a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      showToast('导出成功（已保留原始格式、颜色、公式）');
    } catch(e) {
      console.error('ZIP export error:', e);
      showToast('导出失败: ' + e.message);
    }
  }

  doExport();
}

function clearData() {
  workbook = null; worksheet = null; headers = []; colIdxMap = {}; greenRows.clear();
  seckillItems = []; tableData = []; originalFileName = ''; originalFileBuffer = null;
  document.getElementById('toolbar').style.display = 'none';
  document.getElementById('status-bar').style.display = 'none';
  document.getElementById('table-container').style.display = 'none';
  document.getElementById('seckill-panel').style.display = 'none';
  document.getElementById('seckill-btn').style.display = 'none';
  document.getElementById('progress-container').style.display = 'none';
  const batchBtn = document.getElementById('batch-btn');
  if (batchBtn) batchBtn.style.display = 'none';
  const batchBtn2 = document.getElementById('batch-btn-2');
  if (batchBtn2) batchBtn2.style.display = 'none';
  const thead = document.querySelector('#data-table thead');
  const tbody = document.querySelector('#data-table tbody');
  if (thead) thead.innerHTML = '';
  if (tbody) tbody.innerHTML = '';
  // Re-apply permission UI restrictions
  applyPermissionUI();
}

// ========== Seckill Dialog ==========
function openSeckill() {
  if (!seckillItems.length) { showToast('暂无秒杀推荐SKU'); return; }
  document.getElementById('seckill-modal').classList.add('active');
  const tbody = document.querySelector('#seckill-table tbody');
  const defaultDate = getDefaultDateRange();
  tbody.innerHTML = seckillItems.map((item, i) => `
    <tr>
      <td contenteditable="true">${defaultDate}</td>
      <td><select><option>LD</option><option>BD</option><option>黑五LD</option><option>黑五BD</option><option>网一LD</option><option>网一BD</option><option>秋季</option></select></td>
      <td><select><option>美国</option><option>英国</option><option>德国</option><option>日本</option><option>加拿大</option><option>法国</option><option>意大利</option><option>西班牙</option><option>墨西哥</option><option>澳大利亚</option></select></td>
      <td contenteditable="true">${item.asin || ''}</td>
      <td contenteditable="true">${item.asin || ''}</td>
      <td contenteditable="true">${item.sku || ''}</td>
      <td contenteditable="true"></td>
      <td contenteditable="true"></td>
      <td contenteditable="true">${item.monthly_sales}</td>
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
    if (tds.length >= 9) {
      items.push({
        date: tds[0].textContent, type: tds[1].querySelector('select')?.value || tds[1].textContent,
        site: tds[2].querySelector('select')?.value || tds[2].textContent,
        asin: tds[3].textContent, child_asin: tds[4].textContent,
        sku: tds[5].textContent, qty: tds[6].textContent,
        shop: tds[7].textContent, monthly_sales: tds[8].textContent
      });
    }
  });
  const ziniao = [];
  document.querySelectorAll('#ziniao-table tbody tr').forEach(tr => {
    const tds = tr.querySelectorAll('td');
    if (tds.length >= 5) {
      const row = {
        type: tds[0].querySelector('select')?.value || tds[0].textContent,
        company: tds[1].textContent, username: tds[2].textContent,
        password: tds[3].textContent, shop: tds[4].textContent
      };
      if (Object.values(row).some(v => v && v.trim())) ziniao.push(row);
    }
  });

  const res = await apiFetch('/api/seckill', { method: 'POST', body: JSON.stringify({ items, ziniao_info: ziniao }) });
  if (res && res.ok) {
    showToast('秒杀提报表已保存');
    closeModal('seckill-modal');
  } else {
    showToast('保存失败');
  }
}

function exportSeckill() {
  const items = [];
  document.querySelectorAll('#seckill-table tbody tr').forEach(tr => {
    const tds = tr.querySelectorAll('td');
    if (tds.length >= 8) {
      items.push({
        date: tds[0].textContent, type: tds[1].querySelector('select')?.value || tds[1].textContent,
        site: tds[2].querySelector('select')?.value || tds[2].textContent,
        parent_asin: tds[3].textContent, child_asin: tds[4].textContent,
        sku: tds[5].textContent, qty: tds[6].textContent, shop: tds[7].textContent,
        monthly_sales: tds.length >= 9 ? tds[8].textContent : ''
      });
    }
  });

  const ziniao = [];
  document.querySelectorAll('#ziniao-table tbody tr').forEach(tr => {
    const tds = tr.querySelectorAll('td');
    if (tds.length >= 5) {
      ziniao.push({
        type: tds[0].querySelector('select')?.value || tds[0].textContent,
        company: tds[1].textContent, username: tds[2].textContent,
        password: tds[3].textContent, shop: tds[4].textContent
      });
    }
  });

  const wb = XLSX.utils.book_new();
  const headerRow = ['活动时间段', '秒杀类型', '站点', '父体ASIN', '子体ASIN', '子体SKU', '提报数量', '店铺名', '月销量'];
  const dataRows = items.map(i => [i.date, i.type, i.site, i.parent_asin, i.child_asin, i.sku, i.qty, i.shop, i.monthly_sales || '']);
  const ws = XLSX.utils.aoa_to_sheet([
    ['代报LD、BD收集表'],
    [],
    headerRow,
    ...dataRows,
    [],
    ['温馨提示：请在活动开始前7天提交，确保库存充足'],
    ['注意事项：每个ASIN仅限提报一次，费用另计']
  ]);

  if (ziniao.length) {
    const vpsStart = XLSX.utils.decode_range(ws['!ref']).e.r + 2;
    const vpsHeader = ['类型', 'IP/公司/识别码', '用户名/验证码', '密码', '店铺名'];
    XLSX.utils.sheet_add_aoa(ws, [vpsHeader, ...ziniao.map(z => [z.type, z.company, z.username, z.password, z.shop])], { origin: `A${vpsStart}` });
  }

  ws['!cols'] = [
    { wch: 22 }, { wch: 10 }, { wch: 8 }, { wch: 14 }, { wch: 14 }, { wch: 18 }, { wch: 10 }, { wch: 16 }, { wch: 10 }
  ];

  XLSX.utils.book_append_sheet(wb, ws, '秒杀提报');
  XLSX.writeFile(wb, `秒杀提报_${new Date().toISOString().slice(0,10)}.xlsx`);
  showToast('秒杀Excel已导出！');
}

// ========== Rules Dialog (Complete rules from desktop version) ==========
function openRules() {
  document.getElementById('rules-modal').classList.add('active');
  document.getElementById('rules-content').innerHTML = RULES_HTML;
}

const RULES_HTML = `
<h3>一、补货下单规则</h3>
<h4>1.1 下单时间</h4>
<ul>
  <li><strong>固定下单时间</strong>：每周二、周五。</li>
  <li><strong>截止时间</strong>：跟单当天未明确说明截止时间的，默认下午5:00前截止。</li>
  <li><strong>急单处理</strong>：海运订单超过下午5:00一律不接，如有急单需提前安排助理下单给跟单。</li>
</ul>

<h4>1.2 下单流程</h4>
<ul>
  <li>每周二查看助理提供的ASIN总结，重点关注【本周要出库存】列。</li>
  <li>判断各ASIN是否需要出货或补货，并及时告知助理。</li>
  <li>助理根据确认结果下单，避免反复确认导致时间延误。</li>
  <li>美国海运出货日期固定写"每周五"。</li>
</ul>

<h4>1.3 下单数量与日期填写规范</h4>
<ul>
  <li>一个数量对应一个出货日期，不可只写一个总数量和一个日期。</li>
  <li>助理将直接筛选对应数量和日期的数据下单，降低下错单的概率。</li>
  <li>出货日期必须准确，工厂将根据该日期排单生产。</li>
</ul>

<h3>二、出货清单管理</h3>
<h4>2.1 ASIN总结与库存判断</h4>
<ul>
  <li>助理已在ASIN总结中新增【本周要出库存】列，运营人员只需判断该ASIN是否需要出货或补货，并及时反馈。若未及时沟通，可能导致助理来不及做货件，影响出货进度。</li>
</ul>

<h4>2.2 下周出货清单填写与回传</h4>
<ul>
  <li>由于销量不稳定及库容限制，部分ASIN前期下单过多，无法一次性出完。需每周判断以下事项：</li>
  <li>· 哪些库存下周可以出掉</li>
  <li>· 哪些需改期出货</li>
  <li>· 改期到什么时候出</li>
  <li>建议每周二看完ASIN总结后，一次性将上述信息告知助理，避免反复确认。</li>
</ul>

<h4>2.3 下单数量标注规范</h4>
<ul>
  <li>确定本周出货的：出货数量标从未出货订单里面根据销量写入本周要出货的库存里面。</li>
  <li>部分出货的：例如本周有500个只需出250个，则将250个标，剩余数量在其他列备注，并注明原出货日期改到何时。</li>
</ul>

<h4>2.4 新增备货规范</h4>
<ul>
  <li>如果要新增下单的海运订单，多少数量/什么时候下，分批的话写在海运下单里面的数量/数量/数量/，日历列写每一票日期/日期/日期。</li>
</ul>

<h3>三、货件创建规则</h3>
<h4>3.1 货件时间</h4>
<ul>
  <li>美国海运做货件时间为每周三、周四，以周四居多。</li>
</ul>

<h4>3.2 分箱规则</h4>
<ul>
  <li>美国站货件会分仓，为避免产生入库配置费，每个SKU需发5箱以上（含5箱）。下单时应根据销量的稳定性和上升趋势，尽量保证补货5箱以上。</li>
</ul>

<h4>3.3 货件拆分原则</h4>
<ul>
  <li>若存在低于5箱的SKU（如测品1-2箱），需单独做货件。</li>
</ul>

<h3>四、库存调拨规则</h3>
<h4>4.1 海外仓优先原则</h4>
<ul>
  <li>下单海运前，可先登录海外仓系统查询产品是否有库存。若急需补货，可优先从海外仓发货。</li>
</ul>

<h4>4.2 跨部门调货</h4>
<ul>
  <li>若产品库存低于200个，发走前需确认该库存是否属于其他部门。若需调其他部门库存，必须提前征得对方同意，避免库存发走后导致对方断货。</li>
</ul>

<h3>五、产品包装规格</h3>
<p>规格汇总表【接口备注】列会注明壳子类型。不同产品单个重量不同，装箱个数也不同，请按以下标准执行：</p>
<ul>
  <li>方形壳/华硕壳 = <strong>60个/箱</strong></li>
  <li>小壳子（氮化镓/小米/华为）= <strong>100个/箱</strong></li>
  <li>长条普通壳 = <strong>50个/箱</strong></li>
  <li>大功率（100W以上）= <strong>40个/箱</strong> 或 <strong>25个/箱</strong></li>
</ul>
<p>注：新品如暂不清楚具体装箱数量，可先写下单1箱或2箱。</p>

<h3>六、库存监控与预警</h3>
<h4>6.1 备货标准</h4>
<ul>
  <li>默认备货量（FBA库存+在途库存+采购中库存）应为30天销量的4倍（即约4个月库存），管理员可通过自定义规则调整不同销量区间的备货倍数。</li>
</ul>

<h4>6.2 销量激增判断</h4>
<ul>
  <li>若最近7天销量 × 4 ＞ 30天销量 × 2，则定性为销量激增，需提示是否需要加大备货。</li>
</ul>

<h4>6.3 销量停滞判断</h4>
<ul>
  <li>若总库存（FBA库存+在途+采购中库存）＞ 30天销量 × 6，则定性为销量停滞库存，需给出提醒加快动销。</li>
</ul>

<h4>6.4 秒杀建议</h4>
<ul>
  <li>当月销量＞150的库存，建议每月报3次秒杀+1次周秒杀，并给出是否需要多备7天左右库存的建议。</li>
</ul>

<h3>七、其他注意事项</h3>
<h4>7.1 绿标/翔标识别</h4>
<ul>
  <li>表格中标注绿色的行，代表该产品可能有做绿标。可复制ASIN到详情页查看是否有翔标标识，确认后再做后续操作。</li>
</ul>

<h4>7.2 表格字段说明</h4>
<ul>
  <li>从左至右依次为：产品信息、销量、FBA库存、在途库存、采购中库存、下单和发货区域。</li>
</ul>

<h4>7.3 发货方式建议</h4>
<ul>
  <li>库存天数 < 20天：建议发<strong>美森</strong>（快船，约12-15天到仓）</li>
  <li>库存天数 20-30天：建议发<strong>快船</strong></li>
  <li>库存天数 30-40天：建议发<strong>一半慢船一半美森</strong></li>
  <li>库存天数 ≥ 40天：建议发<strong>慢船</strong>（经济，约35-45天到仓）</li>
</ul>
`;

// ========== Admin Panel ==========
function openAdmin() {
  document.getElementById('admin-modal').classList.add('active');
  loadUsers();
}

function closeModal(id) {
  document.getElementById(id).classList.remove('active');
}

function switchTab(tab, el) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
  if (el) el.classList.add('active');
  document.getElementById('tab-' + tab).classList.add('active');
  if (tab === 'users') loadUsers();
  if (tab === 'logs') loadLogs();
  if (tab === 'history-admin') loadAdminHistory();
  if (tab === 'seckill-admin') loadAdminSeckill();
  if (tab === 'permissions') loadPermissionsTab();
  if (tab === 'versions') loadVersions();
  if (tab === 'llm-config') loadLLMConfig();
}

async function loadUsers() {
  const res = await apiFetch('/api/users');
  if (!res || !res.ok) return;
  const users = await res.json();
  const tbody = document.querySelector('#users-table tbody');
  const roleMap = { admin: '管理员', manager: '组长', user: '运营' };
  // Cache permissions for each user to avoid HTML attribute escaping issues
  usersPermissionsMap = {};
  for (const u of users) {
    usersPermissionsMap[u.username] = u.permissions || null;
  }
  tbody.innerHTML = users.map(u => {
    const hasCustomPerms = u.permissions && Object.keys(u.permissions).length > 0;
    const safeName = u.username.replace(/'/g, "\\'");
    return `
    <tr>
      <td>${u.id}</td><td>${u.username}</td>
      <td><span class="badge badge-${u.role}">${roleMap[u.role] || u.role}</span></td>
      <td>${u.department || '-'}</td>
      <td>${u.is_active ? '<span style="color:#34C759;">启用</span>' : '<span style="color:#FF3B30;">禁用</span>'}</td>
      <td>
        <button class="btn btn-purple" style="padding:3px 8px; font-size:11px;" onclick="openPermEditor('${safeName}')">${hasCustomPerms ? '✏ 权限' : '🔒 权限'}</button>
        <button class="btn btn-outline" style="padding:3px 8px; font-size:11px;" onclick="showResetPw('${safeName}')">重置密码</button>
        <button class="btn btn-secondary" style="padding:3px 8px; font-size:11px;" onclick="toggleUser('${safeName}', ${u.is_active ? 0 : 1})">${u.is_active ? '禁用' : '启用'}</button>
        <button class="btn btn-danger" style="padding:3px 8px; font-size:11px;" onclick="deleteUser('${safeName}')">删除</button>
      </td>
    </tr>`;
  }).join('');
}

async function loadLogs() {
  const res = await apiFetch('/api/logs');
  if (!res) return;
  const logs = await res.json();
  document.querySelector('#logs-table tbody').innerHTML = logs.map(l => `
    <tr>
      <td>${l.id}</td><td>${l.username}</td>
      <td>${l.login_time ? l.login_time.replace('T',' ').slice(0,19) : ''}</td>
      <td>${l.success ? '<span style="color:#34C759;">成功</span>' : '<span style="color:#FF3B30;">失败</span>'}</td>
    </tr>
  `).join('');
}

async function loadAdminHistory() {
  const res = await apiFetch('/api/history');
  if (!res) return;
  const history = await res.json();
  // Load notes counts for all history items
  let notesMap = {};
  try {
    const notesRes = await apiFetch('/api/notes/counts');
    if (notesRes && notesRes.ok) notesMap = await notesRes.json();
  } catch(e) {}
  document.querySelector('#history-admin-table tbody').innerHTML = history.map(h => {
    const noteCount = notesMap[h.id] || 0;
    return `<tr>
      <td>${h.id}</td><td>${h.username}</td><td>${h.filename}</td>
      <td>${h.uploaded_at ? h.uploaded_at.replace('T',' ').slice(0,19) : ''}</td>
      <td>${h.row_count}</td>
      <td><button class="btn btn-info" style="padding:3px 10px; font-size:11px;" onclick="viewUploadData(${h.id}, '${(h.filename||'').replace(/'/g,"\\'")}')">在线查看</button></td>
      <td><button class="btn btn-outline" style="padding:3px 10px; font-size:11px;" onclick="viewUploadData(${h.id}, '${(h.filename||'').replace(/'/g,"\\'")}')">📝 ${noteCount > 0 ? noteCount + '条' : '添加'}</button></td>
    </tr>`;
  }).join('');
}

async function loadAdminSeckill() {
  const res = await apiFetch('/api/seckill');
  if (!res) return;
  const reports = await res.json();
  document.querySelector('#seckill-admin-table tbody').innerHTML = reports.map(r => {
    let itemCount = 0;
    try { itemCount = JSON.parse(r.items).length; } catch(e) {}
    const created = r.created_at ? r.created_at.replace('T',' ').slice(0,16) : '';
    return `<tr>
      <td>${r.id}</td><td>${r.username}</td><td>${created}</td><td>${itemCount}</td>
      <td><button class="btn btn-danger" style="padding:3px 8px; font-size:11px;" onclick="deleteSeckillReport(${r.id})">删除</button></td>
    </tr>`;
  }).join('');
}

async function loadVersions() {
  const res = await apiFetch('/api/versions');
  if (!res) return;
  const versions = await res.json();
  document.querySelector('#versions-table tbody').innerHTML = versions.map(v => `
    <tr>
      <td>${v.id}</td><td>${v.version}</td><td>${v.description || '-'}</td>
      <td>${v.is_active ? '<span style="color:#34C759;">启用</span>' : '<span style="color:#8E8E93;">禁用</span>'}</td>
      <td>
        <button class="btn btn-secondary" style="padding:3px 8px; font-size:11px;" onclick="toggleVersion(${v.id}, ${v.is_active ? 0 : 1})">${v.is_active ? '禁用' : '启用'}</button>
        <button class="btn btn-danger" style="padding:3px 8px; font-size:11px;" onclick="deleteVersion(${v.id})">删除</button>
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
  showToast('用户创建成功');
  loadUsers();
}

async function doAddVersion() {
  const version = document.getElementById('new-version').value.trim();
  const description = document.getElementById('new-version-desc').value.trim();
  if (!version) { alert('请输入版本号'); return; }
  const res = await apiFetch('/api/versions', { method: 'POST', body: JSON.stringify({ version, description }) });
  if (!res) return;
  closeModal('add-version-modal');
  showToast('版本添加成功');
  loadVersions();
}

async function toggleUser(username, active) {
  await apiFetch('/api/users/' + encodeURIComponent(username), { method: 'PUT', body: JSON.stringify({ active: !!active }) });
  loadUsers();
}

async function deleteUser(username) {
  if (!confirm('确定删除用户 ' + username + ' 吗？此操作不可恢复。')) return;
  await apiFetch('/api/users/' + encodeURIComponent(username), { method: 'DELETE' });
  showToast('用户已删除');
  loadUsers();
}

async function toggleVersion(id, active) {
  await apiFetch('/api/versions', { method: 'PUT', body: JSON.stringify({ id, active: !!active }) });
  loadVersions();
}

async function deleteVersion(id) {
  if (!confirm('确定删除该版本吗？')) return;
  await apiFetch('/api/versions', { method: 'DELETE', body: JSON.stringify({ id }) });
  loadVersions();
}

async function deleteSeckillReport(id) {
  if (!confirm('确定删除该秒杀报告吗？')) return;
  await apiFetch('/api/seckill/' + id, { method: 'DELETE' });
  showToast('报告已删除');
  loadAdminSeckill();
}

// Reset Password
function showResetPw(username) {
  document.getElementById('reset-pw-username').value = username;
  document.getElementById('reset-new-pw').value = '';
  document.getElementById('reset-pw-modal').classList.add('active');
}

async function doResetPassword() {
  const username = document.getElementById('reset-pw-username').value;
  const newPw = document.getElementById('reset-new-pw').value.trim();
  if (!newPw) { alert('请输入新密码'); return; }
  const res = await apiFetch('/api/reset-password', { method: 'POST', body: JSON.stringify({ username, new_password: newPw }) });
  if (res && res.ok) {
    closeModal('reset-pw-modal');
    showToast('密码已重置');
  } else {
    showToast('重置失败');
  }
}

// LLM Config
async function loadLLMConfig() {
  const res = await apiFetch('/api/llm-config');
  if (!res) return;
  const data = await res.json();
  document.getElementById('llm-api-key').value = data.api_key || '';
  document.getElementById('llm-base-url').value = data.base_url || '';
  document.getElementById('llm-model-name').value = data.model_name || '';
}

async function saveLLMConfig() {
  const api_key = document.getElementById('llm-api-key').value.trim();
  const base_url = document.getElementById('llm-base-url').value.trim();
  const model_name = document.getElementById('llm-model-name').value.trim();
  const res = await apiFetch('/api/llm-config', { method: 'POST', body: JSON.stringify({ api_key, base_url, model_name }) });
  if (res && res.ok) {
    showToast('LLM配置已保存');
  } else {
    showToast('保存失败');
  }
}

// ========== History Dialog ==========
async function openHistory() {
  document.getElementById('history-modal').classList.add('active');
  const res = await apiFetch('/api/history');
  if (!res) return;
  const history = await res.json();
  document.querySelector('#history-table tbody').innerHTML = history.map(h => `
    <tr>
      <td>${h.id}</td><td>${h.filename}</td>
      <td>${h.uploaded_at ? h.uploaded_at.replace('T',' ').slice(0,19) : ''}</td>
      <td>${h.row_count}</td>
    </tr>
  `).join('');
}

// ========== Admin Online Table Viewer ==========
async function viewUploadData(historyId, filename) {
  document.getElementById('view-data-modal').classList.add('active');
  document.getElementById('view-data-title').textContent = `📊 ${filename || '上传数据'}`;
  document.getElementById('view-data-loading').style.display = 'block';
  document.getElementById('view-data-table-container').style.display = 'none';
  document.getElementById('view-data-info').textContent = '';

  // Show export button for admin
  const exportBtn = document.getElementById('view-data-export-btn');
  if (exportBtn) exportBtn.style.display = (currentUser && currentUser.role === 'admin') ? 'inline-block' : 'none';

  // Show notes section for admin
  const notesSection = document.getElementById('view-data-notes-section');
  if (notesSection) notesSection.style.display = (currentUser && currentUser.role === 'admin') ? 'block' : 'none';

  viewDataHistoryId = historyId;
  viewDataHeaders = null;
  viewDataRows = null;

  const res = await apiFetch('/api/upload-data/' + historyId);
  document.getElementById('view-data-loading').style.display = 'none';

  if (!res || !res.ok) {
    document.getElementById('view-data-info').textContent = '未找到该文件的在线数据（仅支持新上传的文件）';
    return;
  }

  const data = await res.json();
  const vHeaders = data.headers || [];
  const vData = data.data || [];
  const vGreenRows = new Set(data.green_rows || []);

  viewDataHeaders = vHeaders;
  viewDataRows = vData;

  document.getElementById('view-data-table-container').style.display = 'block';
  document.getElementById('view-data-info').textContent = `共 ${vData.length} 行数据`;

  const thead = document.querySelector('#view-data-table thead');
  const tbody = document.querySelector('#view-data-table tbody');
  thead.innerHTML = '<tr>' + vHeaders.map(h => `<th>${h}</th>`).join('') + '</tr>';
  tbody.innerHTML = vData.map((row, rIdx) => {
    const isGreen = vGreenRows.has(rIdx) || vGreenRows.has(String(rIdx));
    const rowClass = isGreen ? 'green-row' : '';
    return `<tr class="${rowClass}">` + vHeaders.map(h => {
      const val = row[h];
      let display = val !== undefined && val !== null ? String(val) : '';
      return `<td>${display}</td>`;
    }).join('') + '</tr>';
  }).join('');

  // Load admin notes for this upload
  if (currentUser && currentUser.role === 'admin') {
    loadAdminNotes(historyId);
  }
}

// ========== Theme Toggle (Dark/Light Mode) ==========
function toggleTheme() {
  const html = document.documentElement;
  const current = html.getAttribute('data-theme') || 'light';
  const next = current === 'light' ? 'dark' : 'light';
  html.setAttribute('data-theme', next);
  document.getElementById('theme-icon').textContent = next === 'dark' ? '🌙' : '☀️';
  try { localStorage.setItem('theme', next); } catch(e) {}
}

function initTheme() {
  try {
    const saved = localStorage.getItem('theme');
    if (saved === 'dark') {
      document.documentElement.setAttribute('data-theme', 'dark');
      const icon = document.getElementById('theme-icon');
      if (icon) icon.textContent = '🌙';
    }
  } catch(e) {}
}
document.addEventListener('DOMContentLoaded', initTheme);

// ========== Permissions Tab (Admin) ==========
const PERM_LABELS = {
  upload: '📁 上传表格', column_settings: '⚙ 列设置', batch_analysis: '📋 批量分析',
  auto_calculate: '🧮 自动计算', export: '📊 导出表格', seckill: '⚡ 提报秒杀',
  history: '📋 历史记录', clear_data: '🗑 清除记录', rules: '📖 规则说明',
  weight_config: '🧮 加权算法'
};

async function loadPermissionsTab() {
  const res = await apiFetch('/api/users');
  if (!res || !res.ok) return;
  const users = await res.json();
  const tbody = document.querySelector('#permissions-table tbody');
  const roleMap = { admin: '管理员', manager: '组长', user: '运营' };
  const nonAdminUsers = users.filter(u => u.role !== 'admin');
  if (!nonAdminUsers.length) {
    tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;color:#8E8E93;padding:16px;">暂无子账号，请先在"用户管理"中添加用户</td></tr>';
    return;
  }
  tbody.innerHTML = nonAdminUsers.map(u => {
    const perms = u.permissions || {};
    const enabledModules = Object.entries(PERM_LABELS)
      .filter(([key]) => perms[key] !== false)
      .map(([, label]) => label)
      .join('、') || '无';
    const disabledCount = Object.values(perms).filter(v => v === false).length;
    return `<tr>
      <td><strong>${u.username}</strong></td>
      <td><span class="badge badge-${u.role}">${roleMap[u.role] || u.role}</span></td>
      <td style="font-size:11px;color:#8E8E93;max-width:300px;overflow:hidden;text-overflow:ellipsis;">${disabledCount > 0 ? enabledModules + ' (已限制' + disabledCount + '项)' : '全部模块'}</td>
      <td>
        <button class="btn btn-purple" style="padding:3px 10px;font-size:11px;" onclick="openPermEditor('${u.username.replace(/'/g, "\\'")}')">设置权限</button>
      </td>
    </tr>`;
  }).join('');
  // Update the cache for the perm editor
  usersPermissionsMap = {};
  for (const u of users) { usersPermissionsMap[u.username] = u.permissions || null; }
}

// ========== Close modals on overlay click ==========
document.addEventListener('click', (e) => {
  if (e.target.classList.contains('modal-overlay') && e.target.classList.contains('active')) {
    e.target.classList.remove('active');
  }
});

// ========== Keyboard shortcuts ==========
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    document.querySelectorAll('.modal-overlay.active').forEach(m => m.classList.remove('active'));
  }
});

// ========== Weighted Algorithm Config ==========
async function loadRestockConfig() {
  try {
    const res = await apiFetch('/api/restock-config');
    if (res && res.ok) {
      const data = await res.json();
      if (data && data.weight_3d !== undefined) {
        restockConfig = {
          weight_3d: data.weight_3d ?? 20,
          weight_7d: data.weight_7d ?? 30,
          weight_15d: data.weight_15d ?? 30,
          weight_30d: data.weight_30d ?? 20,
          stock_months: data.stock_months ?? 4,
          deduct_unshipped: Number(data.deduct_unshipped) !== 0,
          deduct_week_outbound: Number(data.deduct_week_outbound) !== 0
        };
      }
    }
  } catch(e) { /* use defaults */ }
}

async function openWeightConfig() {
  if (!currentUser) { showToast('请先登录'); return; }
  if (currentUser.role !== 'admin' && !hasPermission('weight_config')) { showToast('无权访问此功能'); return; }
  // Apply current config to UI
  document.getElementById('weight-3d').value = restockConfig.weight_3d;
  document.getElementById('weight-7d').value = restockConfig.weight_7d;
  document.getElementById('weight-15d').value = restockConfig.weight_15d;
  document.getElementById('weight-30d').value = restockConfig.weight_30d;
  document.getElementById('stock-months').value = restockConfig.stock_months;
  document.getElementById('deduct-unshipped').checked = restockConfig.deduct_unshipped;
  document.getElementById('deduct-week-outbound').checked = restockConfig.deduct_week_outbound;
  updateWeightHint();
  // Load stocking rules
  await loadStockingRules();
  document.getElementById('weight-config-modal').classList.add('active');
}

function updateWeightHint() {
  const w3 = parseInt(document.getElementById('weight-3d').value) || 0;
  const w7 = parseInt(document.getElementById('weight-7d').value) || 0;
  const w15 = parseInt(document.getElementById('weight-15d').value) || 0;
  const w30 = parseInt(document.getElementById('weight-30d').value) || 0;
  document.getElementById('weight-3d-val').textContent = w3 + '%';
  document.getElementById('weight-7d-val').textContent = w7 + '%';
  document.getElementById('weight-15d-val').textContent = w15 + '%';
  document.getElementById('weight-30d-val').textContent = w30 + '%';
  const total = w3 + w7 + w15 + w30;
  const hint = document.getElementById('weight-hint');
  if (total === 100) {
    hint.textContent = `当前总计：${total}% ✓`;
    hint.className = 'weight-hint ok';
  } else {
    hint.textContent = `当前总计：${total}%（需等于100%）`;
    hint.className = 'weight-hint error';
  }
}

async function saveWeightConfig() {
  const w3 = parseInt(document.getElementById('weight-3d').value) || 0;
  const w7 = parseInt(document.getElementById('weight-7d').value) || 0;
  const w15 = parseInt(document.getElementById('weight-15d').value) || 0;
  const w30 = parseInt(document.getElementById('weight-30d').value) || 0;
  if (w3 + w7 + w15 + w30 !== 100) {
    showToast('四项权重之和必须等于100%，当前为' + (w3+w7+w15+w30) + '%');
    return false;
  }
  const config = {
    weight_3d: w3, weight_7d: w7, weight_15d: w15, weight_30d: w30,
    stock_months: parseInt(document.getElementById('stock-months').value) || 4,
    deduct_unshipped: document.getElementById('deduct-unshipped').checked ? 1 : 0,
    deduct_week_outbound: document.getElementById('deduct-week-outbound').checked ? 1 : 0
  };
  const res = await apiFetch('/api/restock-config', { method: 'POST', body: JSON.stringify(config) });
  if (res && res.ok) {
    restockConfig = {
      weight_3d: w3, weight_7d: w7, weight_15d: w15, weight_30d: w30,
      stock_months: config.stock_months,
      deduct_unshipped: !!config.deduct_unshipped,
      deduct_week_outbound: !!config.deduct_week_outbound
    };
    showToast('加权算法配置已保存，下次自动计算时生效');
    closeModal('weight-config-modal');
    return true;
  } else {
    showToast('保存失败，请重试');
    return false;
  }
}

function resetWeightDefaults() {
  document.getElementById('weight-3d').value = 20;
  document.getElementById('weight-7d').value = 30;
  document.getElementById('weight-15d').value = 30;
  document.getElementById('weight-30d').value = 20;
  document.getElementById('stock-months').value = 4;
  document.getElementById('deduct-unshipped').checked = true;
  document.getElementById('deduct-week-outbound').checked = true;
  updateWeightHint();
  showToast('已恢复默认值，请点击保存以生效');
}

// ========== Stocking Rules ==========
let stockingRules = [];

async function loadStockingRules() {
  const res = await apiFetch('/api/stocking-rules');
  if (res && res.ok) {
    stockingRules = await res.json();
  } else {
    stockingRules = [];
  }
  renderStockingRules();
}

function renderStockingRules() {
  const tbody = document.querySelector('#stocking-rules-table tbody');
  if (!stockingRules.length) {
    tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;color:#8E8E93;padding:12px;">暂无规则，点击"添加规则"创建</td></tr>';
    return;
  }
  stockingRules.sort((a, b) => (a.priority || 0) - (b.priority || 0));
  tbody.innerHTML = stockingRules.map((r, i) => {
    const isActive = r.is_active !== 0 && r.is_active !== false;
    const maxDisplay = (r.max_monthly_sales === 0 || r.max_monthly_sales === 99999) ? '无限制' : (r.max_monthly_sales || 99999);
    return `
    <tr style="${isActive ? '' : 'opacity:0.5;'}">
      <td><input type="text" value="${(r.name || '').replace(/"/g, '&quot;')}" placeholder="规则名称" style="width:90px;padding:3px 5px;border-radius:6px;border:1px solid #E5E5EA;font-size:12px;" onchange="stockingRules[${i}].name=this.value"></td>
      <td><input type="number" value="${r.min_monthly_sales || 0}" min="0" style="width:70px;padding:3px;border-radius:6px;border:1px solid #E5E5EA;text-align:center;" onchange="stockingRules[${i}].min_monthly_sales=parseInt(this.value)"></td>
      <td><input type="number" value="${r.max_monthly_sales || 0}" min="0" style="width:70px;padding:3px;border-radius:6px;border:1px solid #E5E5EA;text-align:center;" onchange="stockingRules[${i}].max_monthly_sales=parseInt(this.value)"></td>
      <td><input type="number" value="${r.multiplier || r.stock_multiplier || 2}" min="0.5" max="20" step="0.5" style="width:55px;padding:3px;border-radius:6px;border:1px solid #E5E5EA;text-align:center;" onchange="stockingRules[${i}].multiplier=parseFloat(this.value); stockingRules[${i}].stock_multiplier=parseFloat(this.value)"></td>
      <td><input type="number" value="${r.priority || i + 1}" min="1" max="99" style="width:45px;padding:3px;border-radius:6px;border:1px solid #E5E5EA;text-align:center;" onchange="stockingRules[${i}].priority=parseInt(this.value)"></td>
      <td><span style="color:${isActive ? '#34C759' : '#FF3B30'}; font-size:12px; font-weight:600;">${isActive ? '启用' : '禁用'}</span></td>
      <td>
        <button class="btn btn-secondary" style="padding:2px 7px;font-size:11px;" onclick="toggleStockingRule(${i})">${isActive ? '禁用' : '启用'}</button>
        <button class="btn btn-danger" style="padding:2px 7px;font-size:11px;" onclick="removeStockingRule(${i})">删除</button>
      </td>
    </tr>`;
  }).join('');
}

function addStockingRule() {
  stockingRules.push({ name: '新规则', priority: stockingRules.length + 1, min_monthly_sales: 0, max_monthly_sales: 0, multiplier: 2, stock_multiplier: 2, is_active: 1 });
  renderStockingRules();
}

function toggleStockingRule(idx) {
  const r = stockingRules[idx];
  const newState = (r.is_active === 0 || r.is_active === false) ? 1 : 0;
  stockingRules[idx].is_active = newState;
  // If rule has an id, update on server immediately
  if (r.id) {
    apiFetch('/api/stocking-rules', { method: 'PUT', body: JSON.stringify({ ...r, is_active: newState }) });
  }
  renderStockingRules();
}

function removeStockingRule(idx) {
  // If the rule has an id, delete from server
  const rule = stockingRules[idx];
  if (rule && rule.id) {
    apiFetch('/api/stocking-rules', { method: 'DELETE', body: JSON.stringify({ id: rule.id }) });
  }
  stockingRules.splice(idx, 1);
  renderStockingRules();
}

// Override saveWeightConfig to also save stocking rules
const _origSaveWeightConfig = saveWeightConfig;
saveWeightConfig = async function() {
  // Save weight config first (validates weights sum to 100)
  const success = await _origSaveWeightConfig();
  if (!success) return; // Don't save stocking rules if weight config failed
  // Then save stocking rules
  try {
    await apiFetch('/api/stocking-rules', { method: 'POST', body: JSON.stringify({ rules: stockingRules }) });
  } catch(e) { /* non-critical */ }
};

// ========== Permission Editor ==========
function openPermEditor(username) {
  currentPermUser = username;
  const currentPerms = usersPermissionsMap[username] || null;
  document.getElementById('perm-username').textContent = username;
  // Reset all checkboxes
  const checkboxes = document.querySelectorAll('#perm-grid input[type="checkbox"]');
  checkboxes.forEach(cb => {
    if (currentPerms && currentPerms[cb.dataset.perm] !== undefined) {
      cb.checked = !!currentPerms[cb.dataset.perm];
    } else {
      cb.checked = true; // Default: all enabled
    }
  });
  document.getElementById('perm-editor-modal').classList.add('active');
}

function toggleAllPerms(state) {
  document.querySelectorAll('#perm-grid input[type="checkbox"]').forEach(cb => cb.checked = state);
}

async function saveUserPermissions() {
  if (!currentPermUser) return;
  const perms = {};
  document.querySelectorAll('#perm-grid input[type="checkbox"]').forEach(cb => {
    perms[cb.dataset.perm] = cb.checked;
  });
  const res = await apiFetch('/api/user-permissions/' + encodeURIComponent(currentPermUser), {
    method: 'PUT', body: JSON.stringify({ permissions: perms })
  });
  if (res && res.ok) {
    showToast(`${currentPermUser} 的权限已更新，用户下次登录时生效`);
    closeModal('perm-editor-modal');
    loadUsers(); // Refresh user list
  } else {
    showToast('权限保存失败');
  }
}

// ========== Export View Data (Admin) ==========
let viewDataHistoryId = null;
let viewDataHeaders = null;
let viewDataRows = null;

function exportViewData() {
  if (!viewDataHeaders || !viewDataRows) { showToast('没有数据可导出'); return; }
  const ws = XLSX.utils.aoa_to_sheet([viewDataHeaders, ...viewDataRows.map(row => viewDataHeaders.map(h => row[h] ?? ''))]);
  ws['!cols'] = viewDataHeaders.map(() => ({ wch: 16 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  const title = document.getElementById('view-data-title').textContent || '导出数据';
  XLSX.writeFile(wb, `${title.replace(/[^a-zA-Z0-9\u4e00-\u9fa5]/g,'_')}_${new Date().toISOString().slice(0,10)}.xlsx`);
  showToast('导出成功');
}

// ========== Admin Notes ==========
let currentNoteHistoryId = null;

async function loadAdminNotes(historyId) {
  currentNoteHistoryId = historyId;
  const res = await apiFetch('/api/notes/' + historyId);
  const list = document.getElementById('admin-notes-list');
  if (!res || !res.ok) {
    list.innerHTML = '<p style="color:#8E8E93;font-size:12px;text-align:center;padding:8px;">暂无备注</p>';
    return;
  }
  const notes = await res.json();
  if (!notes.length) {
    list.innerHTML = '<p style="color:#8E8E93;font-size:12px;text-align:center;padding:8px;">暂无备注，在下方添加第一条备注</p>';
    return;
  }
  list.innerHTML = notes.map(n => `
    <div style="padding:8px 10px; margin-bottom:6px; background:rgba(0,122,255,0.05); border-radius:8px; border-left:3px solid ${n.note_type === 'warning' ? '#FF9500' : n.note_type === 'error' ? '#FF3B30' : '#007AFF'};">
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:3px;">
        <strong style="font-size:12px; color:#007AFF;">${escapeHtml(n.author || '管理员')}</strong>
        <span style="font-size:10px; color:#8E8E93;">${escapeHtml(n.created_at ? n.created_at.replace('T',' ').slice(0,16) : '')}</span>
      </div>
      <p style="font-size:12px; color:#333; margin:0;">${escapeHtml(n.content)}</p>
      ${n.target_user ? '<span style="font-size:10px; color:#8E8E93;">→ 提醒: ' + escapeHtml(n.target_user) + '</span>' : ''}
    </div>
  `).join('');
}

async function addAdminNote() {
  const input = document.getElementById('admin-note-input');
  const content = input.value.trim();
  if (!content) { showToast('请输入备注内容'); return; }
  if (!currentNoteHistoryId) { showToast('请先选择一个上传记录'); return; }

  const res = await apiFetch('/api/notes', {
    method: 'POST',
    body: JSON.stringify({
      history_id: currentNoteHistoryId,
      content: content,
      note_type: 'info'
    })
  });
  if (res && res.ok) {
    input.value = '';
    showToast('备注已添加');
    loadAdminNotes(currentNoteHistoryId);
  } else {
    showToast('添加备注失败');
  }
}

// ========== Chat Sidebar ==========
let chatSidebarOpen = false;
let chatPollInterval = null;
let chatPastedImage = null; // Base64 data URL of pasted image
let chatAllMessages = []; // Cache for search filtering
let chatUserList = []; // Cached user list for @mentions

function toggleChatSidebar() {
  const sidebar = document.getElementById('chat-sidebar');
  chatSidebarOpen = !chatSidebarOpen;
  sidebar.classList.toggle('open', chatSidebarOpen);
  if (chatSidebarOpen) {
    loadChatUsers();
    loadChatMessages();
    if (!chatPollInterval) {
      chatPollInterval = setInterval(loadChatMessages, 10000); // Poll every 10s
    }
  } else {
    if (chatPollInterval) { clearInterval(chatPollInterval); chatPollInterval = null; }
    removeChatImage();
    document.getElementById('chat-search-input').value = '';
    const dropdown = document.getElementById('chat-mention-dropdown');
    if (dropdown) dropdown.style.display = 'none';
  }
}

async function loadChatUsers() {
  const res = await apiFetch('/api/chat-users');
  if (!res || !res.ok) return;
  const users = await res.json();
  chatUserList = users;
  const select = document.getElementById('chat-recipient-select');
  const currentVal = select.value;
  select.innerHTML = '<option value="__all__">全体消息</option>';
  for (const u of users) {
    if (u.username === currentUser?.username) continue;
    const opt = document.createElement('option');
    opt.value = u.username;
    opt.textContent = u.username + (u.role === 'admin' ? ' (管理员)' : '');
    select.appendChild(opt);
  }
  select.value = currentVal || '__all__';
}

async function loadChatMessages() {
  if (!chatSidebarOpen) return;
  const recipient = document.getElementById('chat-recipient-select').value;
  const url = recipient === '__all__' ? '/api/messages' : '/api/messages?with=' + encodeURIComponent(recipient);
  const res = await apiFetch(url);
  const container = document.getElementById('chat-messages');
  if (!res || !res.ok) {
    container.innerHTML = '<div style="text-align:center;color:#8E8E93;padding:20px;font-size:12px;">暂无消息</div>';
    chatAllMessages = [];
    return;
  }
  chatAllMessages = await res.json();
  renderChatMessages(chatAllMessages);
}

function renderChatMessages(messages) {
  const container = document.getElementById('chat-messages');
  if (!messages.length) {
    container.innerHTML = '<div style="text-align:center;color:#8E8E93;padding:20px;font-size:12px;">暂无消息，发送第一条吧</div>';
    return;
  }
  // Apply search filter
  const searchTerm = (document.getElementById('chat-search-input')?.value || '').trim().toLowerCase();
  const filtered = searchTerm
    ? messages.filter(m => (m.content || '').toLowerCase().includes(searchTerm) || (m.sender || '').toLowerCase().includes(searchTerm))
    : messages;

  if (!filtered.length) {
    container.innerHTML = '<div style="text-align:center;color:#8E8E93;padding:20px;font-size:12px;">没有匹配的消息</div>';
    return;
  }

  container.innerHTML = filtered.map(m => {
    const isMine = m.sender === currentUser?.username;
    // Highlight @mentions in content
    let displayContent = escapeHtml(m.content || '');
    displayContent = displayContent.replace(/@(\S+)/g, '<span class="mention-tag">@$1</span>');
    // Show image if present
    let imageHtml = '';
    if (m.image_data) {
      imageHtml = `<br><img src="${m.image_data}" onclick="window.open(this.src,'_blank')" alt="图片">`;
    }
    return `<div class="chat-bubble ${isMine ? 'mine' : 'theirs'}">
      <div class="chat-bubble-meta">${isMine ? '' : '<strong>' + escapeHtml(m.sender || '') + '</strong>'} <span>${escapeHtml(m.created_at ? m.created_at.replace('T',' ').slice(0,16) : '')}</span></div>
      <div class="chat-bubble-text">${displayContent}${imageHtml}</div>
    </div>`;
  }).join('');
  container.scrollTop = container.scrollHeight;
}

function filterChatMessages() {
  renderChatMessages(chatAllMessages);
}

// Image paste support (Ctrl+V)
function setupImagePaste() {
  const chatInput = document.getElementById('chat-input');
  if (!chatInput) return;
  chatInput.addEventListener('paste', function(e) {
    const items = e.clipboardData?.items;
    if (!items) return;
    for (const item of items) {
      if (item.type.startsWith('image/')) {
        e.preventDefault();
        const file = item.getAsFile();
        if (!file) return;
        const reader = new FileReader();
        reader.onload = function(ev) {
          chatPastedImage = ev.target.result;
          const preview = document.getElementById('chat-image-preview');
          const img = document.getElementById('chat-preview-img');
          img.src = chatPastedImage;
          preview.style.display = 'block';
        };
        reader.readAsDataURL(file);
        return;
      }
    }
  });
}

function triggerImagePaste() {
  showToast('请使用 Ctrl+V 粘贴截图，或从聊天工具栏点击 @ 提及成员');
}

function removeChatImage() {
  chatPastedImage = null;
  const preview = document.getElementById('chat-image-preview');
  if (preview) preview.style.display = 'none';
}

// @mention support
function insertMentionTrigger() {
  const input = document.getElementById('chat-input');
  input.value += '@';
  input.focus();
  showMentionDropdown('');
}

function showMentionDropdown(filter) {
  const dropdown = document.getElementById('chat-mention-dropdown');
  if (!chatUserList.length) { dropdown.style.display = 'none'; return; }
  const filtered = chatUserList.filter(u =>
    u.username !== currentUser?.username &&
    u.username.toLowerCase().includes(filter.toLowerCase())
  );
  if (!filtered.length) { dropdown.style.display = 'none'; return; }
  dropdown.innerHTML = filtered.map(u =>
    `<div class="chat-mention-item" onmousedown="selectMention('${escapeHtml(u.username)}')">${escapeHtml(u.username)} <span style="color:#8E8E93;font-size:11px;">${u.role === 'admin' ? '管理员' : '运营'}</span></div>`
  ).join('');
  dropdown.style.display = 'block';
}

function selectMention(username) {
  const input = document.getElementById('chat-input');
  const val = input.value;
  const atIdx = val.lastIndexOf('@');
  if (atIdx >= 0) {
    input.value = val.slice(0, atIdx) + '@' + username + ' ';
  } else {
    input.value += '@' + username + ' ';
  }
  document.getElementById('chat-mention-dropdown').style.display = 'none';
  input.focus();
}

function handleChatInput(e) {
  const val = e.target.value;
  const atIdx = val.lastIndexOf('@');
  if (atIdx >= 0 && (atIdx === 0 || val[atIdx - 1] === ' ')) {
    const filter = val.slice(atIdx + 1);
    if (!filter.includes(' ') && filter.length < 20) {
      showMentionDropdown(filter);
      return;
    }
  }
  document.getElementById('chat-mention-dropdown').style.display = 'none';
}

function handleChatKeydown(e) {
  if (e.key === 'Enter') {
    e.preventDefault();
    sendChatMessage();
  }
  if (e.key === 'Escape') {
    document.getElementById('chat-mention-dropdown').style.display = 'none';
  }
}

async function sendChatMessage() {
  const input = document.getElementById('chat-input');
  const content = input.value.trim();
  if (!content && !chatPastedImage) return;
  const recipient = document.getElementById('chat-recipient-select').value;
  const body = { content: content || '[图片]' };
  if (recipient !== '__all__') body.recipient = recipient;
  if (chatPastedImage) body.image_data = chatPastedImage;

  const res = await apiFetch('/api/messages', {
    method: 'POST',
    body: JSON.stringify(body)
  });
  if (res && res.ok) {
    input.value = '';
    removeChatImage();
    loadChatMessages();
  } else {
    showToast('发送失败');
  }
}

// Initialize image paste on DOM ready
document.addEventListener('DOMContentLoaded', setupImagePaste);
