// ════════════════════════════════════════════════════════════════════════════
//  VIP Auto Showroom — Google Apps Script Backend
//  Deploy as: Web App → Execute as: Me → Who has access: Anyone
// ════════════════════════════════════════════════════════════════════════════

// ── Sheet names (rename to match your workbook) ──────────────────────────
const SHEETS = {
  SALES:     'المبيعات',     // Every individual sale transaction
  INVENTORY: 'المخزون',     // Current vehicle stock
  BRANCHES:  'الفروع',      // Branch master data (optional, for display names)
};

// ── Column indices (0-based) inside the SALES sheet ──────────────────────
// IMPORTANT: The order below must match your actual sheet columns.
// See the "Sheet Structure" section in README for full details.
const COL = {
  // ── SALES sheet ──────────────────────────────────────────────────────
  DATE:       0,   // A  تاريخ_البيع        (YYYY-MM-DD or DD/MM/YYYY)
  BRANCH:     1,   // B  اسم_الفرع
  BRAND:      2,   // C  الماركة
  MODEL:      3,   // D  الموديل
  YEAR:       4,   // E  السنة
  PRICE:      5,   // F  سعر_البيع           (full number, e.g. 95000)
  COST:       6,   // G  التكلفة             (full number)
  PROFIT:     7,   // H  الأرباح             (Price - Cost, can be formula)
  PAYMENT:    8,   // I  طريقة_الدفع        (نقد / تمويل / بطاقة)
  CLIENT_VIP: 9,   // J  عميل_VIP           (TRUE / FALSE  or  نعم / لا)
  FINANCE_REQ:10,  // K  طلب_تمويل          (TRUE / FALSE)
  DAYS_STOCK: 11,  // L  أيام_في_المخزون    (number)  ← filled in Inventory sheet

  // ── INVENTORY sheet (used for صحة المخزون and معدل الدوران) ──────────
  INV_ENTRY_DATE: 0,  // A  تاريخ_الدخول
  INV_BRAND:      1,  // B  الماركة
  INV_BRANCH:     2,  // C  اسم_الفرع
  INV_STATUS:     3,  // D  الحالة  (متاح / محجوز / مباع)
  INV_DAYS:       4,  // E  أيام_في_المخزون  (=TODAY()-A2)
  INV_PAYMENT:    5,  // F  طريقة_الدفع      (optional — leave blank if unknown)
};

// ── Inventory health thresholds (days) ───────────────────────────────────
const THRESHOLDS = { NEEDS: 60, STALE: 90 };

// ═════════════════════════════════════════════════════════════════════════
//  ENTRY POINT
// ═════════════════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    const params  = e && e.parameter ? e.parameter : {};
    const period  = params.period  || 'month';
    const branch  = params.branch  || 'all';
    const brand   = params.brand   || 'all';
    const payment = params.payment || 'all';

    const payload = buildPayload(period, branch, brand, payment);

    return ContentService
      .createTextOutput(JSON.stringify(payload))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    const errPayload = { status: 'error', message: err.message };
    return ContentService
      .createTextOutput(JSON.stringify(errPayload))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ═════════════════════════════════════════════════════════════════════════
//  BUILD FULL PAYLOAD
// ═════════════════════════════════════════════════════════════════════════
function buildPayload(period, branch, brand, payment) {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const salesRows   = getRows(ss, SHEETS.SALES);
  const invRows     = getRows(ss, SHEETS.INVENTORY);

  // Date windows
  const { start: curStart, end: curEnd } = getDateRange(period);
  const { start: prevStart, end: prevEnd } = getPrevDateRange(period);

  // Filter helpers
  const matchBranch  = r => branch  === 'all' || normalize(r[COL.BRANCH])  === normalize(branch);
  const matchBrand   = r => brand   === 'all' || normalize(r[COL.BRAND])   === normalize(brand);
  const matchPayment = r => payment === 'all' || normalize(r[COL.PAYMENT]) === normalize(payment);
  const matchCommon  = r => matchBranch(r) && matchBrand(r) && matchPayment(r);

  // Apply current-period filter
  const curRows  = salesRows.filter(r => {
    const d = parseDate(r[COL.DATE]);
    return d && d >= curStart && d <= curEnd && matchCommon(r);
  });

  // Apply previous-period filter (for % change)
  const prevRows = salesRows.filter(r => {
    const d = parseDate(r[COL.DATE]);
    return d && d >= prevStart && d <= prevEnd && matchCommon(r);
  });

  // Inventory rows with matching branch/brand filters only
  const filteredInv = invRows.filter(r => {
    const brnMatch = branch === 'all' || normalize(r[COL.INV_BRANCH]) === normalize(branch);
    const brdMatch = brand  === 'all' || normalize(r[COL.INV_BRAND])  === normalize(brand);
    return brnMatch && brdMatch;
  });

  // ── KPI: current ─────────────────────────────────────────────────────
  const revenue  = sum(curRows, COL.PRICE);
  const profit   = sum(curRows, COL.PROFIT);
  const carsSold = curRows.length;

  // Inventory turnover = (total sold this period) / (avg inventory count)
  const avgInv       = filteredInv.length || 1;
  const inventoryRate = parseFloat((carsSold / avgInv * (365 / daysDiff(curStart, curEnd))).toFixed(2));

  // ── KPI: previous (for % change) ─────────────────────────────────────
  const prevRevenue  = sum(prevRows, COL.PRICE);
  const prevProfit   = sum(prevRows, COL.PROFIT);
  const prevCarsSold = prevRows.length;

  const changes = {
    revenue:   pctChange(prevRevenue,  revenue),
    profit:    pctChange(prevProfit,   profit),
    sold:      pctChange(prevCarsSold, carsSold),
    inventory: pctChange(prevCarsSold, carsSold), // reuse as proxy
  };

  // ── Monthly trend (always last 6 months, ignores period filter) ───────
  const monthly = buildMonthlyTrend(salesRows, branch, brand, payment);

  // ── Inventory health ──────────────────────────────────────────────────
  const inventory = buildInventoryHealth(filteredInv);

  // ── Branches ─────────────────────────────────────────────────────────
  const branches = buildBranches(salesRows, curStart, curEnd, brand, payment);

  // ── Top brands ────────────────────────────────────────────────────────
  const brands = buildBrands(curRows);

  // ── Today summary ─────────────────────────────────────────────────────
  const today = buildToday(salesRows, invRows);

  return {
    status: 'ok',
    revenue,
    profit,
    carsSold,
    inventoryRate,
    changes,
    monthly,
    inventory,
    branches,
    brands,
    today,
  };
}

// ═════════════════════════════════════════════════════════════════════════
//  MONTHLY TREND  — last 6 calendar months
// ═════════════════════════════════════════════════════════════════════════
function buildMonthlyTrend(allRows, branch, brand, payment) {
  const months = [];
  const now    = new Date();

  for (let i = 5; i >= 0; i--) {
    const d     = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const start = new Date(d.getFullYear(), d.getMonth(), 1);
    const end   = new Date(d.getFullYear(), d.getMonth() + 1, 0, 23, 59, 59);

    const rows = allRows.filter(r => {
      const rd = parseDate(r[COL.DATE]);
      if (!rd || rd < start || rd > end) return false;
      if (branch  !== 'all' && normalize(r[COL.BRANCH])  !== normalize(branch))  return false;
      if (brand   !== 'all' && normalize(r[COL.BRAND])   !== normalize(brand))   return false;
      if (payment !== 'all' && normalize(r[COL.PAYMENT]) !== normalize(payment)) return false;
      return true;
    });

    const AR_MONTHS = ['يناير','فبراير','مارس','أبريل','مايو','يونيو',
                       'يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر'];
    months.push({
      month:   AR_MONTHS[d.getMonth()],
      revenue: sum(rows, COL.PRICE),
      profit:  sum(rows, COL.PROFIT),
    });
  }

  return months;
}

// ═════════════════════════════════════════════════════════════════════════
//  INVENTORY HEALTH  (from Inventory sheet)
// ═════════════════════════════════════════════════════════════════════════
function buildInventoryHealth(invRows) {
  const active = invRows.filter(r => normalize(r[COL.INV_STATUS]) !== 'مباع');
  let healthy = 0, needs = 0, stale = 0;

  active.forEach(r => {
    const days = parseFloat(r[COL.INV_DAYS]) || 0;
    if      (days >= THRESHOLDS.STALE) stale++;
    else if (days >= THRESHOLDS.NEEDS) needs++;
    else                               healthy++;
  });

  const total = healthy + needs + stale || 1;
  return {
    healthy: Math.round(healthy / total * 100),
    needs:   Math.round(needs   / total * 100),
    stale:   100 - Math.round(healthy / total * 100) - Math.round(needs / total * 100),
  };
}

// ═════════════════════════════════════════════════════════════════════════
//  BRANCH PERFORMANCE
// ═════════════════════════════════════════════════════════════════════════
function buildBranches(allRows, curStart, curEnd, brand, payment) {
  const curRows  = allRows.filter(r => {
    const d = parseDate(r[COL.DATE]);
    return d && d >= curStart && d <= curEnd &&
           (brand   === 'all' || normalize(r[COL.BRAND])   === normalize(brand)) &&
           (payment === 'all' || normalize(r[COL.PAYMENT]) === normalize(payment));
  });

  const { start: pS, end: pE } = getPrevByRange(curStart, curEnd);
  const prevRows = allRows.filter(r => {
    const d = parseDate(r[COL.DATE]);
    return d && d >= pS && d <= pE &&
           (brand   === 'all' || normalize(r[COL.BRAND])   === normalize(brand)) &&
           (payment === 'all' || normalize(r[COL.PAYMENT]) === normalize(payment));
  });

  const branchMap  = groupBy(curRows,  COL.BRANCH, COL.PRICE);
  const prevBrnMap = groupBy(prevRows, COL.BRANCH, COL.PRICE);

  return Object.entries(branchMap)
    .sort((a, b) => b[1] - a[1])
    .map(([name, rev]) => {
      const prevRev = prevBrnMap[name] || 0;
      return {
        name,
        revenue: rev,
        growth:  prevRev > 0 ? parseFloat(pctChange(prevRev, rev).toFixed(1)) : 0,
      };
    });
}

// ═════════════════════════════════════════════════════════════════════════
//  TOP BRANDS
// ═════════════════════════════════════════════════════════════════════════
function buildBrands(rows) {
  const revMap  = groupBy(rows, COL.BRAND, COL.PRICE);
  const profMap = groupBy(rows, COL.BRAND, COL.PROFIT);

  return Object.entries(revMap)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 6)
    .map(([name, rev]) => ({
      name,
      revenue: rev,
      profit:  profMap[name] || 0,
    }));
}

// ═════════════════════════════════════════════════════════════════════════
//  DAILY SUMMARY  — today only
// ═════════════════════════════════════════════════════════════════════════
function buildToday(salesRows, invRows) {
  const todayStart = new Date(); todayStart.setHours(0, 0, 0, 0);
  const todayEnd   = new Date(); todayEnd.setHours(23, 59, 59, 999);

  const todaySales = salesRows.filter(r => {
    const d = parseDate(r[COL.DATE]);
    return d && d >= todayStart && d <= todayEnd;
  });

  const todayInv = invRows.filter(r => {
    const d = parseDate(r[COL.INV_ENTRY_DATE]);
    return d && d >= todayStart && d <= todayEnd;
  });

  const sales    = todaySales.length;
  const inventory = todayInv.length;
  const finance  = todaySales.filter(r => isTruthy(r[COL.FINANCE_REQ])).length;
  const vip      = todaySales.filter(r => isTruthy(r[COL.CLIENT_VIP])).length;

  return { sales, inventory, finance, vip };
}

// ═════════════════════════════════════════════════════════════════════════
//  DATE RANGE HELPERS
// ═════════════════════════════════════════════════════════════════════════
function getDateRange(period) {
  const now = new Date();
  let start, end;

  if (period === 'month') {
    start = new Date(now.getFullYear(), now.getMonth(), 1);
    end   = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59);

  } else if (period === 'last_month') {
    start = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    end   = new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59);

  } else if (period === 'quarter') {
    const qStart = Math.floor(now.getMonth() / 3) * 3;
    start = new Date(now.getFullYear(), qStart, 1);
    end   = new Date(now.getFullYear(), qStart + 3, 0, 23, 59, 59);

  } else { // year
    start = new Date(now.getFullYear(), 0, 1);
    end   = new Date(now.getFullYear(), 11, 31, 23, 59, 59);
  }

  return { start, end };
}

function getPrevDateRange(period) {
  const now = new Date();
  let start, end;

  if (period === 'month') {
    start = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    end   = new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59);

  } else if (period === 'last_month') {
    start = new Date(now.getFullYear(), now.getMonth() - 2, 1);
    end   = new Date(now.getFullYear(), now.getMonth() - 1, 0, 23, 59, 59);

  } else if (period === 'quarter') {
    const qStart = Math.floor(now.getMonth() / 3) * 3;
    start = new Date(now.getFullYear(), qStart - 3, 1);
    end   = new Date(now.getFullYear(), qStart, 0, 23, 59, 59);

  } else { // year
    start = new Date(now.getFullYear() - 1, 0, 1);
    end   = new Date(now.getFullYear() - 1, 11, 31, 23, 59, 59);
  }

  return { start, end };
}

function getPrevByRange(curStart, curEnd) {
  const diff  = curEnd - curStart;
  const pE    = new Date(curStart.getTime() - 1);
  const pS    = new Date(pE.getTime() - diff);
  return { start: pS, end: pE };
}

// ═════════════════════════════════════════════════════════════════════════
//  LOW-LEVEL HELPERS
// ═════════════════════════════════════════════════════════════════════════

/** Read all data rows (skip header) from a named sheet */
function getRows(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  return data.slice(1); // drop header row
}

/** Sum a numeric column across rows */
function sum(rows, col) {
  return rows.reduce((acc, r) => acc + (parseFloat(String(r[col]).replace(/,/g, '')) || 0), 0);
}

/** Group rows by a key column and sum a value column */
function groupBy(rows, keyCol, valCol) {
  return rows.reduce((map, r) => {
    const k = String(r[keyCol] || '').trim();
    if (!k) return map;
    map[k] = (map[k] || 0) + (parseFloat(String(r[valCol]).replace(/,/g, '')) || 0);
    return map;
  }, {});
}

/** Percentage change from prev to cur */
function pctChange(prev, cur) {
  if (!prev) return 0;
  return parseFloat(((cur - prev) / prev * 100).toFixed(1));
}

/** Days between two dates */
function daysDiff(a, b) {
  return Math.max(1, Math.round((b - a) / 86400000));
}

/** Normalize Arabic strings for loose comparison */
function normalize(v) {
  return String(v || '').trim().replace(/\s+/g, ' ');
}

/** Treat sheet boolean-ish values as true/false */
function isTruthy(v) {
  const s = String(v).toLowerCase().trim();
  return s === 'true' || s === 'نعم' || s === '1' || s === 'yes';
}

/**
 * Parse a date value from a sheet cell.
 * Handles: JS Date objects (returned by GAS), ISO strings, DD/MM/YYYY strings.
 */
function parseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v) ? null : v;
  const s = String(v).trim();
  // DD/MM/YYYY
  const dmy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (dmy) return new Date(+dmy[3], +dmy[2] - 1, +dmy[1]);
  const d = new Date(s);
  return isNaN(d) ? null : d;
}
