// ════════════════════════════════════════════════════════════════════════════
//  VIP Auto Showroom — Google Apps Script Backend
//  يتناسب مع هيكل ملف معرض_3.xlsx الفعلي
//  Deploy: Web App → Execute as: Me → Who has access: Anyone
// ════════════════════════════════════════════════════════════════════════════

// ── أسماء الأوراق كما هي في الملف ───────────────────────────────────────
const SHEETS = {
  SALES:     'المبيعات',
  INVENTORY: 'المخزون',
  CARS:      'السيارات',
  BRANCHES:  'الفروع',
  CUSTOMERS: 'العملاء',
  SERVICES:  'الصيانة والخدمات',
};

// ── أعمدة ورقة المبيعات (0-based) ──────────────────────────────────────
// معرف البيع | معرف العميل | معرف خط البيع | معرف السيارة | معرف الموظف |
// معرف الفرع | تاريخ البيع | وقت البيع | السعر الأساسي | نسبة الخصم |
// قيمة الخصم الفعلية | السعر بعد الخصم | تكلفة السيارة لنا | الربح الإجمالي |
// نسبة الربح | طريقة الدفع | اسم البنك | ...
const S = {
  ID:           0,
  CUSTOMER_ID:  1,
  LEAD_ID:      2,
  CAR_ID:       3,
  EMP_ID:       4,
  BRANCH_ID:    5,
  DATE:         6,
  TIME:         7,
  BASE_PRICE:   8,
  DISCOUNT_PCT: 9,
  DISCOUNT_VAL: 10,
  FINAL_PRICE:  11,  // السعر بعد الخصم  ← الإيرادات
  COST:         12,  // تكلفة السيارة لنا
  PROFIT:       13,  // الربح الإجمالي
  PROFIT_PCT:   14,
  PAYMENT:      15,  // طريقة الدفع: كاش / تمويل بنكي / بطاقة
};

// ── أعمدة ورقة المخزون (0-based) ────────────────────────────────────────
// معرف المخزون | معرف السيارة | تاريخ الطلب | تاريخ الشراء | تاريخ الوصول |
// تاريخ البيع | تاريخ التسليم | حالة السيارة | نوع التوريد | تكلفة الشراء
const I = {
  ID:           0,
  CAR_ID:       1,
  ORDER_DATE:   2,
  BUY_DATE:     3,
  ARRIVE_DATE:  4,
  SALE_DATE:    5,
  DELIVER_DATE: 6,
  STATUS:       7,  // مباعة / محجوزة / متاحة
  SUPPLY_TYPE:  8,
  COST:         9,
};

// ── أعمدة ورقة السيارات (0-based) ──────────────────────────────────────
// معرف السيارة | الماركة | الموديل | سنة الموديل | الفئة | السعر الأساسي |
// تكلفة الشراء | حالة السيارة | ...
const C = {
  ID:      0,
  BRAND:   1,  // الماركة
  MODEL:   2,
  YEAR:    3,
  CAT:     4,
  PRICE:   5,
  COST:    6,
  STATUS:  7,
};

// ── أعمدة ورقة الفروع (0-based) ─────────────────────────────────────────
// معرف الفرع | اسم الفرع | المدينة | الحي | ...
const B = {
  ID:   0,
  NAME: 1,
  CITY: 2,
};

// ── أعمدة ورقة العملاء (0-based) ────────────────────────────────────────
// معرف العميل | الاسم | العمر | الجنس | المدينة | مستوى الدخل المتوقع | ...
const CUST = {
  ID:     0,
  NAME:   1,
  INCOME: 5,  // عالي / متوسط / منخفض
};

// ── عتبات صحة المخزون (بالأيام) ─────────────────────────────────────────
const THRESHOLDS = { NEEDS: 60, STALE: 90 };

// ═════════════════════════════════════════════════════════════════════════
//  ENTRY POINT
// ═════════════════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    const p       = (e && e.parameter) ? e.parameter : {};
    const period  = p.period  || 'month';
    const branch  = p.branch  || 'all';
    const brand   = p.brand   || 'all';
    const payment = p.payment || 'all';

    const payload = buildPayload(period, branch, brand, payment);
    return buildResponse(JSON.stringify(payload));

  } catch (err) {
    return buildResponse(JSON.stringify({ status: 'error', message: err.message, stack: err.stack }));
  }
}

// إضافة CORS headers صريحة للسماح بالطلبات من أي مصدر
function buildResponse(jsonStr) {
  return ContentService
    .createTextOutput(jsonStr)
    .setMimeType(ContentService.MimeType.JSON);
}

// اختبار مباشر من محرر Apps Script
function testRun() {
  const result = buildPayload('month', 'all', 'all', 'all');
  Logger.log(JSON.stringify(result, null, 2));
}

// ═════════════════════════════════════════════════════════════════════════
//  BUILD FULL PAYLOAD
// ═════════════════════════════════════════════════════════════════════════
function buildPayload(period, branchFilter, brandFilter, paymentFilter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── تحميل جميع الأوراق ───────────────────────────────────────────────
  const salesRows  = getRows(ss, SHEETS.SALES);
  const invRows    = getRows(ss, SHEETS.INVENTORY);
  const carRows    = getRows(ss, SHEETS.CARS);
  const branchRows = getRows(ss, SHEETS.BRANCHES);
  const custRows   = getRows(ss, SHEETS.CUSTOMERS);

  // ── جداول بحث سريع (lookup maps) ─────────────────────────────────────
  // معرف الفرع → { name, city }
  const branchMap = {};
  branchRows.forEach(r => {
    branchMap[str(r[B.ID])] = { name: str(r[B.NAME]), city: str(r[B.CITY]) };
  });

  // معرف السيارة → الماركة
  const carBrandMap = {};
  carRows.forEach(r => { carBrandMap[str(r[C.ID])] = str(r[C.BRAND]); });

  // معرف العميل → مستوى الدخل
  const custIncomeMap = {};
  custRows.forEach(r => { custIncomeMap[str(r[CUST.ID])] = str(r[CUST.INCOME]); });

  // ── إثراء كل صف مبيعات بالمدينة والماركة وطريقة الدفع الموحدة ──────
  const enriched = salesRows.map(r => {
    const bi      = branchMap[str(r[S.BRANCH_ID])] || {};
    const brand   = carBrandMap[str(r[S.CAR_ID])]  || '';
    const payRaw  = str(r[S.PAYMENT]);
    const pay     = normalizePayment(payRaw);
    const date    = parseDate(r[S.DATE]);
    const income  = custIncomeMap[str(r[S.CUSTOMER_ID])] || '';
    return { row: r, city: bi.city || '', branchName: bi.name || '', brand, pay, date, income };
  });

  // ── نوافذ التاريخ ─────────────────────────────────────────────────────
  const { start: curS, end: curE }   = getDateRange(period);
  const { start: prevS, end: prevE } = getPrevDateRange(period);

  // ── دالة مطابقة الفلاتر ───────────────────────────────────────────────
  const matchFilters = (en, start, end) => {
    if (!en.date || en.date < start || en.date > end) return false;
    if (branchFilter  !== 'all' && en.city  !== branchFilter)  return false;
    if (brandFilter   !== 'all' && en.brand !== brandFilter)   return false;
    if (paymentFilter !== 'all' && en.pay   !== paymentFilter) return false;
    return true;
  };

  const curRows  = enriched.filter(en => matchFilters(en, curS,  curE));
  const prevRows = enriched.filter(en => matchFilters(en, prevS, prevE));

  // ══ 1. KPI ═══════════════════════════════════════════════════════════
  const revenue  = sumCol(curRows,  r => num(r.row[S.FINAL_PRICE]));
  const profit   = sumCol(curRows,  r => num(r.row[S.PROFIT]));
  const carsSold = curRows.length;

  const prevRevenue  = sumCol(prevRows, r => num(r.row[S.FINAL_PRICE]));
  const prevProfit   = sumCol(prevRows, r => num(r.row[S.PROFIT]));
  const prevCarsSold = prevRows.length;

  // معدل دوران المخزون: عدد المبيعات ÷ متوسط المخزون × (365 ÷ أيام الفترة)
  const avgInv       = invRows.filter(r => str(r[I.STATUS]) !== 'مباعة').length || 1;
  const periodDays   = Math.max(1, Math.round((curE - curS) / 86400000));
  const inventoryRate = parseFloat((carsSold / avgInv * (365 / periodDays)).toFixed(2));

  const prevInvRate   = parseFloat((prevCarsSold / avgInv * (365 / periodDays)).toFixed(2));

  const changes = {
    revenue:   pct(prevRevenue,  revenue),
    profit:    pct(prevProfit,   profit),
    sold:      pct(prevCarsSold, carsSold),
    inventory: pct(prevInvRate,  inventoryRate),
  };

  // ══ 2. الرسم البياني الشهري (آخر 6 أشهر) ═══════════════════════════
  const monthly = buildMonthly(enriched, branchFilter, brandFilter, paymentFilter);

  // ══ 3. صحة المخزون ══════════════════════════════════════════════════
  const inventory = buildInventoryHealth(invRows, branchMap, branchFilter);

  // ══ 4. أداء الفروع ══════════════════════════════════════════════════
  const branches = buildBranches(enriched, curS, curE, brandFilter, paymentFilter, branchMap);

  // ══ 5. أفضل الماركات ════════════════════════════════════════════════
  const brands = buildBrands(curRows);

  // ══ 6. ملخص اليوم ═══════════════════════════════════════════════════
  const today = buildToday(enriched, invRows, custIncomeMap);

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
//  الرسم البياني الشهري
// ═════════════════════════════════════════════════════════════════════════
function buildMonthly(enriched, branchFilter, brandFilter, paymentFilter) {
  const now = new Date();
  const AR  = ['يناير','فبراير','مارس','أبريل','مايو','يونيو',
               'يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر'];
  const months = [];

  for (let i = 5; i >= 0; i--) {
    const s = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const e = new Date(now.getFullYear(), now.getMonth() - i + 1, 0, 23, 59, 59);

    const rows = enriched.filter(en => {
      if (!en.date || en.date < s || en.date > e) return false;
      if (branchFilter  !== 'all' && en.city  !== branchFilter)  return false;
      if (brandFilter   !== 'all' && en.brand !== brandFilter)   return false;
      if (paymentFilter !== 'all' && en.pay   !== paymentFilter) return false;
      return true;
    });

    months.push({
      month:   AR[s.getMonth()],
      revenue: Math.round(sumCol(rows, r => num(r.row[S.FINAL_PRICE]))),
      profit:  Math.round(sumCol(rows, r => num(r.row[S.PROFIT]))),
    });
  }
  return months;
}

// ═════════════════════════════════════════════════════════════════════════
//  صحة المخزون
// ═════════════════════════════════════════════════════════════════════════
function buildInventoryHealth(invRows, branchMap, branchFilter) {
  const today    = new Date();
  let healthy = 0, needs = 0, stale = 0;

  invRows.forEach(r => {
    const status = str(r[I.STATUS]);
    if (status === 'مباعة') return;

    // تصفية بالمدينة عبر معرف السيارة غير متاح هنا — نتجاوز فلتر الفرع للمخزون
    const arriveDate = parseDate(r[I.ARRIVE_DATE]);
    if (!arriveDate) return;

    const days = Math.round((today - arriveDate) / 86400000);
    if      (days >= THRESHOLDS.STALE) stale++;
    else if (days >= THRESHOLDS.NEEDS) needs++;
    else                               healthy++;
  });

  const total = (healthy + needs + stale) || 1;
  const hPct  = Math.round(healthy / total * 100);
  const nPct  = Math.round(needs   / total * 100);
  const sPct  = 100 - hPct - nPct;

  return { healthy: hPct, needs: nPct, stale: sPct };
}

// ═════════════════════════════════════════════════════════════════════════
//  أداء الفروع — مجمّع حسب المدينة
// ═════════════════════════════════════════════════════════════════════════
function buildBranches(enriched, curS, curE, brandFilter, paymentFilter, branchMap) {
  const { start: pS, end: pE } = getPrevByRange(curS, curE);

  const filter = (en, s, e) =>
    en.date && en.date >= s && en.date <= e &&
    (brandFilter   === 'all' || en.brand === brandFilter) &&
    (paymentFilter === 'all' || en.pay   === paymentFilter);

  const curRows  = enriched.filter(en => filter(en, curS,  curE));
  const prevRows = enriched.filter(en => filter(en, pS, pE));

  const curMap  = groupByKey(curRows,  en => en.city, en => num(en.row[S.FINAL_PRICE]));
  const prevMap = groupByKey(prevRows, en => en.city, en => num(en.row[S.FINAL_PRICE]));

  return Object.entries(curMap)
    .sort((a, b) => b[1] - a[1])
    .map(([city, rev]) => ({
      name:    city,
      revenue: Math.round(rev),
      growth:  parseFloat(pct(prevMap[city] || 0, rev).toFixed(1)),
    }));
}

// ═════════════════════════════════════════════════════════════════════════
//  أفضل الماركات
// ═════════════════════════════════════════════════════════════════════════
function buildBrands(curRows) {
  const revMap  = groupByKey(curRows, en => en.brand, en => num(en.row[S.FINAL_PRICE]));
  const profMap = groupByKey(curRows, en => en.brand, en => num(en.row[S.PROFIT]));

  return Object.entries(revMap)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 6)
    .map(([name, rev]) => ({
      name,
      revenue: Math.round(rev),
      profit:  Math.round(profMap[name] || 0),
    }));
}

// ═════════════════════════════════════════════════════════════════════════
//  ملخص اليوم
// ═════════════════════════════════════════════════════════════════════════
function buildToday(enriched, invRows, custIncomeMap) {
  const ts = new Date(); ts.setHours(0, 0, 0, 0);
  const te = new Date(); te.setHours(23, 59, 59, 999);

  const todaySales = enriched.filter(en => en.date && en.date >= ts && en.date <= te);

  // إدخالات المخزون اليوم: تاريخ الوصول = اليوم
  const todayInv = invRows.filter(r => {
    const d = parseDate(r[I.ARRIVE_DATE]);
    return d && d >= ts && d <= te;
  });

  const sales    = todaySales.length;
  const inventory = todayInv.length;
  const finance  = todaySales.filter(en => en.pay === 'finance').length;

  // عملاء VIP = مستوى الدخل المتوقع "عالي"
  const vip = todaySales.filter(en => {
    const income = custIncomeMap[str(en.row[S.CUSTOMER_ID])] || '';
    return income === 'عالي';
  }).length;

  return { sales, inventory, finance, vip };
}

// ═════════════════════════════════════════════════════════════════════════
//  DATE HELPERS
// ═════════════════════════════════════════════════════════════════════════
function getDateRange(period) {
  const now = new Date();
  let start, end;
  if (period === 'month') {
    start = new Date(now.getFullYear(), now.getMonth(), 1);
    end   = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59);
  } else if (period === 'last_month') {
    start = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    end   = new Date(now.getFullYear(), now.getMonth(),     0, 23, 59, 59);
  } else if (period === 'quarter') {
    const q = Math.floor(now.getMonth() / 3) * 3;
    start = new Date(now.getFullYear(), q,     1);
    end   = new Date(now.getFullYear(), q + 3, 0, 23, 59, 59);
  } else {
    start = new Date(now.getFullYear(), 0,  1);
    end   = new Date(now.getFullYear(), 11, 31, 23, 59, 59);
  }
  return { start, end };
}

function getPrevDateRange(period) {
  const now = new Date();
  let start, end;
  if (period === 'month') {
    start = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    end   = new Date(now.getFullYear(), now.getMonth(),     0, 23, 59, 59);
  } else if (period === 'last_month') {
    start = new Date(now.getFullYear(), now.getMonth() - 2, 1);
    end   = new Date(now.getFullYear(), now.getMonth() - 1, 0, 23, 59, 59);
  } else if (period === 'quarter') {
    const q = Math.floor(now.getMonth() / 3) * 3;
    start = new Date(now.getFullYear(), q - 3, 1);
    end   = new Date(now.getFullYear(), q,     0, 23, 59, 59);
  } else {
    start = new Date(now.getFullYear() - 1, 0,  1);
    end   = new Date(now.getFullYear() - 1, 11, 31, 23, 59, 59);
  }
  return { start, end };
}

function getPrevByRange(curS, curE) {
  const diff = curE - curS;
  const pE   = new Date(curS.getTime() - 1);
  const pS   = new Date(pE.getTime()   - diff);
  return { start: pS, end: pE };
}

// ═════════════════════════════════════════════════════════════════════════
//  UTILITY HELPERS
// ═════════════════════════════════════════════════════════════════════════

function getRows(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1);
}

function num(v) {
  return parseFloat(String(v || 0).replace(/,/g, '')) || 0;
}

function str(v) {
  return String(v == null ? '' : v).trim();
}

function sumCol(rows, fn) {
  return rows.reduce((acc, r) => acc + fn(r), 0);
}

function groupByKey(rows, keyFn, valFn) {
  return rows.reduce((map, r) => {
    const k = keyFn(r);
    if (!k) return map;
    map[k] = (map[k] || 0) + valFn(r);
    return map;
  }, {});
}

function pct(prev, cur) {
  if (!prev) return 0;
  return parseFloat(((cur - prev) / prev * 100).toFixed(1));
}

/**
 * توحيد طريقة الدفع مع قيم فلتر الـ HTML
 * HTML يرسل: cash / finance / card
 * الشيت يحتوي: كاش / تمويل بنكي / بطاقة
 */
function normalizePayment(v) {
  const s = str(v).toLowerCase();
  if (s === 'كاش'        || s === 'cash')          return 'cash';
  if (s.includes('تمويل') || s === 'finance')       return 'finance';
  if (s.includes('بطاقة') || s === 'card')          return 'card';
  return s;
}

function parseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v) ? null : v;
  const s   = String(v).trim();
  const dmy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (dmy) return new Date(+dmy[3], +dmy[2] - 1, +dmy[1]);
  const d = new Date(s);
  return isNaN(d) ? null : d;
}
