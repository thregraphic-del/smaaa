/**
 * ═══════════════════════════════════════════════════════════════════
 *  VIP Car Showroom – Executive Dashboard Backend
 *  Google Apps Script Web App
 *
 *  Deploy as:  Execute as → Me  |  Access → Anyone
 *  After deploy, paste the /exec URL into index.html → API_URL
 * ═══════════════════════════════════════════════════════════════════
 *
 *  REQUIRED GOOGLE SHEETS STRUCTURE
 *  ─────────────────────────────────
 *  Sheet 1 → "Sales"
 *  ┌────────────┬────────┬─────────┬─────────────┬────────────┬────────┬────────────┬─────────────┐
 *  │    Date    │ Branch │  Brand  │ PaymentType │ SalePrice  │ Profit │ CarStatus  │ ClientType  │
 *  ├────────────┼────────┼─────────┼─────────────┼────────────┼────────┼────────────┼─────────────┤
 *  │ 2024-01-05 │ jeddah │ toyota  │    cash     │  185000    │ 22000  │   sold     │    vip      │
 *  │ 2024-01-07 │ riyadh │ lexus   │   finance   │  320000    │ 38000  │   sold     │   normal    │
 *  └────────────┴────────┴─────────┴─────────────┴────────────┴────────┴────────────┴─────────────┘
 *
 *  Sheet 2 → "Inventory"
 *  ┌────────────┬────────┬─────────┬────────────────┬────────────────┐
 *  │    Date    │ Branch │  Brand  │  InventoryDays │ InventoryStatus│
 *  ├────────────┼────────┼─────────┼────────────────┼────────────────┤
 *  │ 2024-01-10 │  abha  │ hyundai │       45       │    healthy     │
 *  │ 2024-01-12 │dammam  │ nissan  │      120       │     stale      │
 *  └────────────┴────────┴─────────┴────────────────┴────────────────┘
 *
 *  Allowed values
 *  ──────────────
 *  Branch:          jeddah | riyadh | abha | dammam
 *  Brand:           toyota | lexus | hyundai | nissan | bmw  (+ any others)
 *  PaymentType:     cash | finance
 *  ClientType:      vip | normal
 *  InventoryStatus: healthy | needs_followup | stale
 *
 * ═══════════════════════════════════════════════════════════════════
 */

// ── Column indices for "Sales" sheet (0-based) ──────────────────────
var SALES_COL = {
  DATE:        0,
  BRANCH:      1,
  BRAND:       2,
  PAYMENT:     3,
  SALE_PRICE:  4,
  PROFIT:      5,
  CAR_STATUS:  6,
  CLIENT_TYPE: 7
};

// ── Column indices for "Inventory" sheet (0-based) ──────────────────
var INV_COL = {
  DATE:   0,
  BRANCH: 1,
  BRAND:  2,
  DAYS:   3,
  STATUS: 4
};

// ── Brand display names ──────────────────────────────────────────────
var BRAND_NAMES = {
  toyota:  'تويوتا',
  lexus:   'لكزس',
  hyundai: 'هيونداي',
  nissan:  'نيسان',
  bmw:     'BMW'
};

// ── Branch display names ─────────────────────────────────────────────
var BRANCH_NAMES = {
  jeddah: 'جدة',
  riyadh: 'الرياض',
  abha:   'أبها',
  dammam: 'الدمام'
};

// ── Month labels (Arabic) ────────────────────────────────────────────
var MONTH_AR = ['يناير','فبراير','مارس','أبريل','مايو','يونيو',
                'يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر'];

// ════════════════════════════════════════════════════════════════════
//  Entry point – doGet
// ════════════════════════════════════════════════════════════════════
function doGet(e) {
  var params  = e && e.parameter ? e.parameter : {};
  var period  = (params.period  || 'this_month').toLowerCase();
  var branch  = (params.branch  || 'all').toLowerCase();
  var brand   = (params.brand   || 'all').toLowerCase();
  var payment = (params.payment || 'all').toLowerCase();

  try {
    var ss        = SpreadsheetApp.getActiveSpreadsheet();
    var salesRows = getRows(ss, 'Sales');
    var invRows   = getRows(ss, 'Inventory');

    // Date windows
    var window    = getDateWindow(period);
    var curRange  = window.current;
    var prevRange = window.previous;

    // Filter sales for current & previous periods
    var curSales  = filterSales(salesRows, curRange,  branch, brand, payment);
    var prevSales = filterSales(salesRows, prevRange, branch, brand, payment);

    // Filter inventory for current period
    var curInv    = filterInventory(invRows, curRange, branch, brand);

    var payload = {
      kpi:            buildKpi(curSales, prevSales, curInv),
      monthly:        buildMonthly(salesRows, curRange, branch, brand, payment),
      inventoryHealth: buildInventoryHealth(curInv),
      branches:       buildBranches(salesRows, curRange, prevRange, branch, brand, payment),
      brands:         buildBrands(curSales),
      dailySummary:   buildDailySummary(salesRows, invRows, branch, brand)
    };

    return buildResponse(payload);

  } catch (err) {
    return buildResponse({ error: err.message });
  }
}

// ════════════════════════════════════════════════════════════════════
//  Sheet reader
// ════════════════════════════════════════════════════════════════════
function getRows(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  return data.slice(1); // skip header row
}

// ════════════════════════════════════════════════════════════════════
//  Date window helpers
// ════════════════════════════════════════════════════════════════════
function getDateWindow(period) {
  var now   = new Date();
  var year  = now.getFullYear();
  var month = now.getMonth(); // 0-based

  var cur, prev;

  switch (period) {
    case 'last_month':
      cur  = { start: new Date(year, month - 1, 1), end: new Date(year, month, 0) };
      prev = { start: new Date(year, month - 2, 1), end: new Date(year, month - 1, 0) };
      break;

    case 'this_quarter':
      var qStart = Math.floor(month / 3) * 3;
      cur  = { start: new Date(year, qStart,     1), end: new Date(year, qStart + 3, 0) };
      prev = { start: new Date(year, qStart - 3, 1), end: new Date(year, qStart,     0) };
      break;

    case 'this_year':
      cur  = { start: new Date(year,     0, 1), end: new Date(year,     11, 31) };
      prev = { start: new Date(year - 1, 0, 1), end: new Date(year - 1, 11, 31) };
      break;

    case 'this_month':
    default:
      cur  = { start: new Date(year, month,     1), end: new Date(year, month + 1, 0) };
      prev = { start: new Date(year, month - 1, 1), end: new Date(year, month,     0) };
      break;
  }

  // Normalise to midnight
  [cur.start, cur.end, prev.start, prev.end].forEach(function(d) {
    d.setHours(0, 0, 0, 0);
  });
  cur.end.setHours(23, 59, 59, 999);
  prev.end.setHours(23, 59, 59, 999);

  return { current: cur, previous: prev };
}

function inRange(dateVal, range) {
  var d = parseDate(dateVal);
  return d && d >= range.start && d <= range.end;
}

function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  var d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

// ════════════════════════════════════════════════════════════════════
//  Row filters
// ════════════════════════════════════════════════════════════════════
function filterSales(rows, range, branch, brand, payment) {
  return rows.filter(function(r) {
    if (!inRange(r[SALES_COL.DATE], range)) return false;
    if (branch  !== 'all' && String(r[SALES_COL.BRANCH]).toLowerCase()  !== branch)  return false;
    if (brand   !== 'all' && String(r[SALES_COL.BRAND]).toLowerCase()   !== brand)   return false;
    if (payment !== 'all' && String(r[SALES_COL.PAYMENT]).toLowerCase() !== payment) return false;
    return true;
  });
}

function filterInventory(rows, range, branch, brand) {
  return rows.filter(function(r) {
    if (!inRange(r[INV_COL.DATE], range)) return false;
    if (branch !== 'all' && String(r[INV_COL.BRANCH]).toLowerCase() !== branch) return false;
    if (brand  !== 'all' && String(r[INV_COL.BRAND]).toLowerCase()  !== brand)  return false;
    return true;
  });
}

// ════════════════════════════════════════════════════════════════════
//  KPI builder
// ════════════════════════════════════════════════════════════════════
function buildKpi(curSales, prevSales, curInv) {
  var curRevenue  = sum(curSales,  SALES_COL.SALE_PRICE);
  var prevRevenue = sum(prevSales, SALES_COL.SALE_PRICE);
  var curProfit   = sum(curSales,  SALES_COL.PROFIT);
  var prevProfit  = sum(prevSales, SALES_COL.PROFIT);
  var curSold     = curSales.length;
  var prevSold    = prevSales.length;

  // Inventory turnover = cars sold ÷ avg inventory items (simple proxy)
  var avgInv      = curInv.length || 1;
  var curTurnover = parseFloat((curSold / avgInv).toFixed(2));
  var prevAvgInv  = Math.max(prevSales.length, 1);
  var prevTurnover= parseFloat((prevSold / prevAvgInv).toFixed(2));

  return {
    revenue:          { value: curRevenue,  change: pctChange(curRevenue,  prevRevenue)  },
    profit:           { value: curProfit,   change: pctChange(curProfit,   prevProfit)   },
    carsSold:         { value: curSold,     change: pctChange(curSold,     prevSold)     },
    inventoryTurnover:{ value: curTurnover, change: pctChange(curTurnover, prevTurnover) }
  };
}

// ════════════════════════════════════════════════════════════════════
//  Monthly trend builder
// ════════════════════════════════════════════════════════════════════
function buildMonthly(salesRows, curRange, branch, brand, payment) {
  // Build one entry per month inside curRange
  var startMonth = curRange.start.getMonth();
  var startYear  = curRange.start.getFullYear();
  var endMonth   = curRange.end.getMonth();
  var endYear    = curRange.end.getFullYear();

  var months = [];
  var y = startYear, m = startMonth;
  while (y < endYear || (y === endYear && m <= endMonth)) {
    months.push({ year: y, month: m });
    m++;
    if (m > 11) { m = 0; y++; }
  }

  return months.map(function(ym) {
    var mStart = new Date(ym.year, ym.month, 1);
    var mEnd   = new Date(ym.year, ym.month + 1, 0, 23, 59, 59, 999);
    var mRows  = filterSales(salesRows, { start: mStart, end: mEnd }, branch, brand, payment);
    return {
      month:   MONTH_AR[ym.month],
      revenue: sum(mRows, SALES_COL.SALE_PRICE),
      profit:  sum(mRows, SALES_COL.PROFIT)
    };
  });
}

// ════════════════════════════════════════════════════════════════════
//  Inventory health builder
// ════════════════════════════════════════════════════════════════════
function buildInventoryHealth(invRows) {
  var total = invRows.length || 1;
  var counts = { healthy: 0, needsFollowup: 0, stale: 0 };

  invRows.forEach(function(r) {
    var status = String(r[INV_COL.STATUS]).toLowerCase().trim();
    if (status === 'healthy')           counts.healthy++;
    else if (status === 'needs_followup') counts.needsFollowup++;
    else if (status === 'stale')          counts.stale++;
    else {
      // Classify by InventoryDays if status is missing
      var days = parseFloat(r[INV_COL.DAYS]) || 0;
      if (days <= 60)       counts.healthy++;
      else if (days <= 120) counts.needsFollowup++;
      else                  counts.stale++;
    }
  });

  return {
    healthy:       Math.round((counts.healthy       / total) * 100),
    needsFollowup: Math.round((counts.needsFollowup / total) * 100),
    stale:         Math.round((counts.stale         / total) * 100)
  };
}

// ════════════════════════════════════════════════════════════════════
//  Branch performance builder
// ════════════════════════════════════════════════════════════════════
function buildBranches(salesRows, curRange, prevRange, branchFilter, brand, payment) {
  var branches = Object.keys(BRANCH_NAMES);

  return branches.map(function(b) {
    // If a specific branch is selected, only show that one
    if (branchFilter !== 'all' && branchFilter !== b) return null;

    var cur  = filterSales(salesRows, curRange,  b, brand, payment);
    var prev = filterSales(salesRows, prevRange, b, brand, payment);
    var curRev  = sum(cur,  SALES_COL.SALE_PRICE);
    var prevRev = sum(prev, SALES_COL.SALE_PRICE);

    return {
      name:    BRANCH_NAMES[b],
      revenue: curRev,
      growth:  pctChange(curRev, prevRev)
    };
  }).filter(Boolean);
}

// ════════════════════════════════════════════════════════════════════
//  Top brands builder
// ════════════════════════════════════════════════════════════════════
function buildBrands(curSales) {
  var agg = {};
  curSales.forEach(function(r) {
    var b = String(r[SALES_COL.BRAND]).toLowerCase().trim();
    if (!agg[b]) agg[b] = { sold: 0, revenue: 0 };
    agg[b].sold++;
    agg[b].revenue += parseFloat(r[SALES_COL.SALE_PRICE]) || 0;
  });

  return Object.keys(agg)
    .sort(function(a, b) { return agg[b].sold - agg[a].sold; })
    .slice(0, 7)
    .map(function(b) {
      return {
        name:    BRAND_NAMES[b] || b,
        sold:    agg[b].sold,
        revenue: agg[b].revenue
      };
    });
}

// ════════════════════════════════════════════════════════════════════
//  Daily summary builder
// ════════════════════════════════════════════════════════════════════
function buildDailySummary(salesRows, invRows, branch, brand) {
  var today     = new Date();
  today.setHours(0, 0, 0, 0);
  var todayEnd  = new Date(today);
  todayEnd.setHours(23, 59, 59, 999);
  var todayRange = { start: today, end: todayEnd };

  var todaySales = filterSales(salesRows, todayRange, branch, brand, 'all');
  var todayInv   = filterInventory(invRows, todayRange, branch, brand);

  var financeReqs = todaySales.filter(function(r) {
    return String(r[SALES_COL.PAYMENT]).toLowerCase() === 'finance';
  }).length;

  var vipClients = todaySales.filter(function(r) {
    return String(r[SALES_COL.CLIENT_TYPE]).toLowerCase() === 'vip';
  }).length;

  return {
    sales:            todaySales.length,
    inventoryEntries: todayInv.length,
    financeRequests:  financeReqs,
    vipClients:       vipClients
  };
}

// ════════════════════════════════════════════════════════════════════
//  Utility functions
// ════════════════════════════════════════════════════════════════════
function sum(rows, colIdx) {
  return rows.reduce(function(acc, r) {
    return acc + (parseFloat(r[colIdx]) || 0);
  }, 0);
}

function pctChange(current, previous) {
  if (!previous || previous === 0) return current > 0 ? 100 : 0;
  return parseFloat(((current - previous) / Math.abs(previous) * 100).toFixed(1));
}

function buildResponse(payload) {
  var json = JSON.stringify(payload);
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════════
//  Quick test function – run manually in the GAS editor to verify
// ════════════════════════════════════════════════════════════════════
function testDoGet() {
  var fakeEvent = {
    parameter: {
      period:  'this_month',
      branch:  'all',
      brand:   'all',
      payment: 'all'
    }
  };
  var result = doGet(fakeEvent);
  Logger.log(result.getContent());
}
