// ═══════════════════════════════════════════════════════
// 彩虹CFO Apps Script v3.28
// 更新日期：2026/06/19
// ───────────────────────────────────────────────────────
// 新增（vs v3.20）：
//   ★ getFixedAssets：whitelist 支援玉山日幣/富邦美金/旅遊基金/Suica/現金日幣
//   ★ updateLiquidCash：preferredKey 優先扣特定 row
//   ★ addTravelRecord：旅遊頁付款帳戶自動扣現金（依帳戶優先扣對應 row）
//   ★ findVtMarketValue：從新 table 格式抓 VT 市值
// ═══════════════════════════════════════════════════════

const SS_ID              = '1PcD6z0CWAMghLgjXgY69W176DQf0LHPV-pbWnTDyjYI';
const TRAVEL_SHEET       = '旅遊記帳';
const CLOSING_SHEET      = '月度結算';
const DEBT_SHEET         = '負債管理';
const ASSET_CONFIG_SHEET = '資產設定';
const INSTALLMENT_SHEET  = '分期負債';

const RETIRE_GOAL_GROSS = 12380000;
const RETIRE_GOAL_NET   = 7269000;

const FALLBACK_CASH           = 600000;
const FALLBACK_HUSBAND_RETIRE = 1607275;
const FALLBACK_WIFE_RETIRE    = 1052243;

const BUDGET_DEF = {
  '餐飲':22000, '生活雜支':21500, '交通':2500, '醫療保健':6000,
  '娛樂休閒':5000, '學習成長':5000, '水電費':3000, '房貸':15515,
  '保險費':9560, '健身房':2376, '手機費':1098, '投資':15000
  // 旅遊基金改為行程制，不計入月預算
};
const BUDGET_TOTAL = Object.values(BUDGET_DEF).reduce((a,b) => a+b, 0);

const GITHUB_REPO   = 'rainbowsupermans2/rainbow-bookkeep';
const SNAPSHOT_PATH = 'data/snapshot.json';

function doGet(e) {
  try {
    const type     = e.parameter.type     || '';
    const callback = e.parameter.callback || 'cb';
    const data     = e.parameter.data     || '';
    const project  = e.parameter.project  || '';
    const months   = parseInt(e.parameter.months) || 12;
    const year     = parseInt(e.parameter.year)   || 0;
    const month    = parseInt(e.parameter.month)  || 0;

    let result;
    if      (type === 'expense')       result = getExpense();
    else if (type === 'asset')         result = getAsset();
    else if (type === 'history')       result = getHistory();
    else if (type === 'add')           result = addRecord(data);
    else if (type === 'addTravel')     result = addTravelRecord(data);
    else if (type === 'getTravel')     result = getTravelRecords(project);
    else if (type === 'monthlyReport') result = monthlyAnalysis(months);
    else if (type === 'closing')       result = monthEndClosing(year, month);
    else if (type === 'getClosing')    result = getClosing(parseInt(e.parameter.months) || 6);
    else if (type === 'debts')         result = { success: true, data: getDebts() };
    else if (type === 'fixedAssets')   result = { success: true, data: getFixedAssets() };
    else if (type === 'installments')  result = { success: true, data: getInstallments() };
    else if (type === 'mortgagePost')  result = monthlyMortgageAutoPost(e.parameter.force === '1');
    else if (type === 'installPost')   result = monthlyInstallmentAutoPost(e.parameter.force === '1');
    else if (type === 'syncInstDebt')  result = syncInstallmentDebtBalances();
    else result = { success: false, error: 'unknown type: ' + type };

    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  } catch(err) {
    const cb = (e && e.parameter && e.parameter.callback) || 'cb';
    return ContentService
      .createTextOutput(cb + '({"success":false,"error":' + JSON.stringify(err.message) + '})')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
}

// ════════════════════════════════════════════
//  ★ v3.21：讀取「資產設定」分頁（whitelist 多 cash row）
// ════════════════════════════════════════════
function getFixedAssets() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(ASSET_CONFIG_SHEET);
    const result = {
      cash:          FALLBACK_CASH,
      fixedDeposit:  FALLBACK_CASH,
      liquidCash:    0,
      husbandRetire: FALLBACK_HUSBAND_RETIRE,
      wifeRetire:    FALLBACK_WIFE_RETIRE,
    };
    if (!sheet) return result;
    const rows = sheet.getDataRange().getValues();
    if (!rows.length) return result;

    let keyCol = 0, valCol = 1;
    if (String(rows[0][0]||'').trim() === '日期') { keyCol = 1; valCol = 2; }

    const CASH_KEYS = ['現金', '活存合計', '玉山日幣', '富邦美金', '旅遊基金',
                       '現金日幣', '彩虹Suica', '先生Suica', 'Suica（日幣）',
                       '現金菲幣'];
    let liquidCashSum = 0;

    for (let i = 1; i < rows.length; i++) {
      const key = String(rows[i][keyCol]||'').trim();
      const val = parseFloat(rows[i][valCol])||0;
      if (val < 0) continue;
      if (key === '定存')        result.fixedDeposit = val;
      else if (key === '先生勞退') result.husbandRetire = val;
      else if (key === '彩虹勞退') result.wifeRetire = val;
      else if (CASH_KEYS.indexOf(key) >= 0) liquidCashSum += val;
    }
    result.liquidCash = liquidCashSum;
    result.cash = result.fixedDeposit + result.liquidCash;
    return result;
  } catch(e) {
    Logger.log('getFixedAssets error: ' + e.message);
    return { cash: FALLBACK_CASH, fixedDeposit: FALLBACK_CASH, liquidCash: 0,
             husbandRetire: FALLBACK_HUSBAND_RETIRE, wifeRetire: FALLBACK_WIFE_RETIRE };
  }
}

// ════════════════════════════════════════════
//  ★ v3.21：更新「現金」餘額（preferredKey 優先扣對應 row）
// ════════════════════════════════════════════
function updateLiquidCash(delta, preferredKey) {
  try {
    if (!delta) return false;
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(ASSET_CONFIG_SHEET);
    if (!sheet) return false;
    const rows = sheet.getDataRange().getValues();
    if (!rows.length) return false;

    let keyCol = 0, valCol = 1, dateCol = -1;
    if (String(rows[0][0]||'').trim() === '日期') { keyCol = 1; valCol = 2; dateCol = 0; }

    const SEARCH_ORDER = preferredKey
      ? [preferredKey, '活存合計', '現金', '玉山日幣', '富邦美金']
      : ['活存合計', '現金', '玉山日幣', '富邦美金'];

    for (const key of SEARCH_ORDER) {
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][keyCol]||'').trim() === key) {
          const cur = parseFloat(rows[i][valCol])||0;
          const next = cur + delta;
          sheet.getRange(i+1, valCol+1).setValue(next);
          if (dateCol >= 0) {
            sheet.getRange(i+1, dateCol+1).setValue(Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd'));
          }
          Logger.log('updateLiquidCash: ' + key + ' ' + cur + ' → ' + next + '（Δ' + delta + '）');
          return true;
        }
      }
    }
    Logger.log('updateLiquidCash: 找不到任何現金列');
    return false;
  } catch(e) { Logger.log('updateLiquidCash error: ' + e.message); return false; }
}

function getHistory() {
  try {
    const ss    = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName('資產記錄');
    if (!sheet) return { success: false, error: '找不到資產記錄工作表' };
    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return { success: true, data: [] };
    const data = rows.slice(1).filter(r => r[0]).map(r => r.map((cell, i) => {
      if (i === 0 && cell instanceof Date) return Utilities.formatDate(cell, 'Asia/Taipei', 'yyyy-MM-dd');
      return cell;
    }));
    return { success: true, data: data };
  } catch(e) { return { success: false, error: e.message }; }
}

function getAsset() {
  try {
    const result = getHistory();
    if (!result.success || !result.data.length) return { success: false, error: '無資產數據' };
    const last = result.data[result.data.length - 1];
    return { success: true, data: {
      date:String(last[0]||''), stock:parseFloat(last[1])||0, vt:parseFloat(last[2])||0,
      fund:parseFloat(last[3])||0, cash:parseFloat(last[4])||0,
      husbandRetire:parseFloat(last[5])||0, wifeRetire:parseFloat(last[6])||0,
      total:parseFloat(last[7])||0, netWorth:parseFloat(last[8])||0,
      progress:parseFloat(last[9])||0, distToTarget:parseFloat(last[10])||0,
    }};
  } catch(e) { return { success: false, error: e.message }; }
}

function getExpense() {
  try {
    const ss    = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName('收支記帳');
    if (!sheet) return { success: false, error: '找不到收支記帳工作表' };
    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return { success: true, data: [rows[0]] };
    const data = rows.map(r => r.map((cell, i) => {
      if (i === 0 && cell instanceof Date) return Utilities.formatDate(cell, 'Asia/Taipei', 'yyyy-MM-dd');
      return cell;
    }));
    return { success: true, data: data };
  } catch(e) { return { success: false, error: e.message }; }
}

function getDebts() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(DEBT_SHEET);
    if (!sheet) return { rows:[], byType:{}, totalDebt:0, mortgage:0, creditCard:0, error:'找不到負債管理工作表' };
    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return { rows:[], byType:{}, totalDebt:0, mortgage:0, creditCard:0 };
    const data = rows.slice(1).filter(r => r[1]).map(r => ({
      date:       r[0] instanceof Date ? Utilities.formatDate(r[0], 'Asia/Taipei', 'yyyy-MM-dd') : String(r[0]||''),
      type:       String(r[1]||''),
      balance:    parseFloat(r[2])||0,
      monthly:    parseFloat(r[3])||0,
      rate:       parseFloat(r[4])||0,
      payoffDate: r[5] instanceof Date ? Utilities.formatDate(r[5], 'Asia/Taipei', 'yyyy-MM') : String(r[5]||''),
      note:       String(r[6]||''),
    }));
    const byType = {};
    let totalDebt = 0, mortgage = 0, creditCard = 0;
    data.forEach(d => {
      byType[d.type] = d;
      totalDebt += d.balance;
      if (d.type === '房貸') mortgage = d.balance;
      else creditCard += d.balance;
    });
    return { rows: data, byType, totalDebt, mortgage, creditCard };
  } catch(e) {
    Logger.log('getDebts error: ' + e.message);
    return { rows:[], byType:{}, totalDebt:0, mortgage:0, creditCard:0, error:e.message };
  }
}

function updateDebtBalance(debtType, delta) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(DEBT_SHEET);
    if (!sheet) return false;
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]||'') === debtType) {
        const cur = parseFloat(rows[i][2])||0;
        const next = cur + delta;
        const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');
        sheet.getRange(i+1, 1).setValue(today);
        sheet.getRange(i+1, 3).setValue(next);
        Logger.log('updateDebtBalance: ' + debtType + ' ' + cur + ' → ' + next);
        return true;
      }
    }
    return false;
  } catch(e) { Logger.log('updateDebtBalance error: ' + e.message); return false; }
}

// ════════════════════════════════════════════
// 同步分期負債餘額 → 負債管理（連動收支記帳）
// 根據「分期負債」的 passedPeriods 計算剩餘未還總額，
// 並更新「負債管理」對應列的債務餘額欄位。
// ════════════════════════════════════════════
function syncInstallmentDebtBalances() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const debtSheet = ss.getSheetByName(DEBT_SHEET);
    if (!debtSheet) return { success: false, error: '找不到負債管理工作表' };

    const inst = getInstallments();
    const debtRows = debtSheet.getDataRange().getValues();
    const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');
    const updated = [], notFound = [];

    inst.rows.forEach(function(i) {
      // 計算正確的剩餘未還總額
      const correctBalance = i.remainingAmount;
      let matchRow = -1;

      for (let r = 1; r < debtRows.length; r++) {
        const type = String(debtRows[r][1] || '');
        if (!type || type === '房貸') continue;
        // ★ v3.25 CRITICAL: 只匹配分期追蹤列，絕不匹配信用卡帳單列
        if (!/分期|人壽|保險/.test(type)) continue;
        const typeClean = type.split('（')[0].trim();

        // 同時含「分期/人壽/保險」關鍵字 + 相同末尾數字
        const typeNum  = (typeClean.match(/\d+$|\d+(?=期|\s)/) || typeClean.match(/\d+/) || [''])[0];
        const instStr  = i.item + ' ' + i.card;
        const instNum  = (instStr.match(/\d+/) || [''])[0];
        if (typeNum && instNum && typeNum === instNum) {
          const banks = ['富邦','中信','玉山','南山','新光','台灣人壽','國泰'];
          const bankMatch = banks.some(function(k) {
            return (type.includes(k) || typeClean.includes(k)) &&
                   (i.item.includes(k) || i.card.includes(k));
          });
          if (bankMatch || (/分期/.test(type) && /分期/.test(instStr))) {
            matchRow = r; break;
          }
        }
      }

      if (matchRow < 0) { notFound.push(i.item); return; }

      const currentBalance = parseFloat(debtRows[matchRow][2]) || 0;
      if (currentBalance !== correctBalance) {
        debtSheet.getRange(matchRow + 1, 1).setValue(today);
        debtSheet.getRange(matchRow + 1, 3).setValue(correctBalance);
        updated.push(String(debtRows[matchRow][1]) + '：' + currentBalance + ' → ' + correctBalance);
        Logger.log('syncInstDebt: ' + debtRows[matchRow][1] + ' ' + currentBalance + ' → ' + correctBalance);
      }
    });

    return { success: true, updated: updated, notFound: notFound };
  } catch(e) {
    Logger.log('syncInstallmentDebtBalances error: ' + e.message);
    return { success: false, error: e.message };
  }
}

function getInstallments() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(INSTALLMENT_SHEET);
    if (!sheet) return { rows: [], totalRemaining: 0 };
    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return { rows: [], totalRemaining: 0 };

    const items = [];
    rows.slice(1).forEach((r, idx) => {
      if (!r[1]) return;
      let startYm = '';
      if (r[0] instanceof Date) {
        startYm = Utilities.formatDate(r[0], 'Asia/Taipei', 'yyyy/MM');
      } else if (r[0]) {
        startYm = String(r[0]).trim();
      }
      const item          = String(r[1]||'');
      const category      = String(r[2]||'');
      const card          = String(r[3]||'');
      const monthly       = parseFloat(r[4]) || 0;
      const totalPeriods  = parseInt(r[5]) || 0;
      const passedPeriods = parseInt(r[6]) || 0;
      const note          = String(r[7]||'');
      if (monthly <= 0 || totalPeriods <= 0) return;

      const currentPeriod    = passedPeriods + 1;
      const remainingPeriods = totalPeriods - passedPeriods;
      const isActive         = remainingPeriods > 0;

      items.push({
        rowIndex: idx + 2,
        startYm, item, category, card, monthly, totalPeriods, passedPeriods, note,
        currentPeriod, remainingPeriods,
        remainingAmount: remainingPeriods * monthly,
        isActive,
      });
    });

    return { rows: items, totalRemaining: items.reduce((s,i) => s + i.remainingAmount, 0) };
  } catch(e) {
    Logger.log('getInstallments error: ' + e.message);
    return { rows: [], totalRemaining: 0, error: e.message };
  }
}

function addRecord(dataStr) {
  try {
    const d     = JSON.parse(decodeURIComponent(dataStr));
    const ss    = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName('收支記帳');
    const row   = d.row || [d.date, d.type, d.category, d.item || d.desc, d.amount, d.account, d.project];
    sheet.appendRow(row);

    const amount  = parseFloat(d.amount) || 0;
    const account = String(d.account  || '');
    const cat     = String(d.category || '');
    const recType = String(d.type     || '');
    const updates = [];

    if (amount > 0) {
      const debts = getDebts();

      if (recType === '支出') {
        if (cat === '房貸' && debts.byType['房貸']) {
          if (updateDebtBalance('房貸', -amount)) updates.push('房貸 -' + amount);
        } else if (cat === '還卡費' && debts.byType[account]) {
          if (updateDebtBalance(account, -amount)) updates.push(account + ' -' + amount);
        } else if (debts.byType[account] && account !== '房貸') {
          if (updateDebtBalance(account, amount)) updates.push(account + ' +' + amount);
        }
      }

      const cashAccounts = ['現金', '富邦銀行', '富邦美金', '第一銀行', '板信', '郵局', 'LINE Bank', 'LINE Pay'];
      const isCashAcct = cashAccounts.indexOf(account) >= 0;
      let liquidDelta = 0;

      if (recType === '支出') {
        if (isCashAcct) liquidDelta = -amount;
        else if (cat === '還卡費') liquidDelta = -amount;
      } else if (recType === '收入') {
        const nonCashAccounts = ['玉山日幣', 'Suica（日幣）'];
        if (nonCashAccounts.indexOf(account) < 0) liquidDelta = amount;
      }

      if (liquidDelta !== 0) {
        if (updateLiquidCash(liquidDelta, account)) {
          updates.push('現金 ' + (liquidDelta > 0 ? '+' : '') + liquidDelta);
        }
      }
    }

    return { success: true, updates: updates };
  } catch(e) { return { success: false, error: e.message }; }
}

// ════════════════════════════════════════════
//  ★ v3.21：旅遊記帳（自動扣現金 + 信用卡負債）
// ════════════════════════════════════════════
function addTravelRecord(dataStr) {
  try {
    const d     = JSON.parse(decodeURIComponent(dataStr));
    const ss    = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(TRAVEL_SHEET) || ss.insertSheet(TRAVEL_SHEET);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['日期','項目','類別','原始金額','幣別','匯率','台幣換算','付款方式','專案代碼','備註']);
    }
    const amountTwd = parseFloat(d.amountTwd) || 0;
    sheet.appendRow([
      d.date||'', d.item||'', d.category||'', d.amount||0, d.currency||'JPY',
      d.rate||0, amountTwd, d.payment||'', d.project||'', d.note||'',
    ]);

    const updates = [];
    const payment = String(d.payment || '');
    if (payment && amountTwd > 0) {
      const debts = getDebts();
      if (debts.byType[payment] && payment !== '房貸') {
        if (updateDebtBalance(payment, amountTwd)) updates.push(payment + ' +' + amountTwd);
      } else {
        const TRIP_CASH = ['現金日幣', '玉山日幣', '彩虹Suica', '先生Suica',
                           'LINE Pay', '台幣預付', 'Kimi地陪現金',
                           '現金菲幣', '富邦銀行', 'LINE Bank'];
        if (TRIP_CASH.indexOf(payment) >= 0) {
          if (updateLiquidCash(-amountTwd, payment)) {
            updates.push('現金(' + payment + ') -' + amountTwd);
          }
        }
      }
    }
    return { success: true, updates };
  } catch(e) { return { success: false, error: e.message }; }
}

function getTravelRecords(project) {
  try {
    const ss    = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(TRAVEL_SHEET);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, data: [] };
    const rows = sheet.getDataRange().getValues();
    const data = rows.slice(1).filter(r => {
      if (!r[0]) return false;
      if (project && r[8] && r[8] !== project) return false;
      return true;
    }).map(r => r.map((cell, i) => {
      if (i === 0 && cell instanceof Date) return Utilities.formatDate(cell, 'Asia/Taipei', 'yyyy-MM-dd');
      return cell;
    }));
    return { success: true, data: data };
  } catch(e) { return { success: false, error: e.message }; }
}

function getFundNAV(fundCode) {
  try {
    let apiCode = fundCode;
    if (fundCode.startsWith('B') && fundCode.length === 6) {
      apiCode = fundCode.substring(0, 3) + ',' + fundCode.substring(3);
    }
    const url = 'https://fund.api.cnyes.com/fund/api/v2/funds/' + apiCode + '/nav?format=json';
    const res  = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(res.getContentText());
    return json.items ? json.items.nav : null;
  } catch(e) { Logger.log('getFundNAV error: ' + e.message); return null; }
}

function refreshFundNAVs() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName('1150421鉅亨基金');
    if (!sheet) return { success: false, error: '找不到基金工作表' };
    const rows = sheet.getDataRange().getValues();
    let updated = 0, failed = 0;
    const log = [];
    for (let i = 1; i < rows.length; i++) {
      const code = String(rows[i][0] || '').trim();
      if (!code || code.indexOf('台幣市值') >= 0 || code.indexOf('1150421') >= 0) continue;
      const nav = getFundNAV(code);
      if (nav !== null && nav > 0) {
        sheet.getRange(i + 1, 6).setValue(nav);
        log.push(code + '=' + nav); updated++;
      } else { log.push(code + '=FAIL'); failed++; }
    }
    SpreadsheetApp.flush();
    return { success: true, updated, failed, detail: log };
  } catch(e) { return { success: false, error: e.message }; }
}

function getEffectiveDate(dateValue, type, category) {
  const date = dateValue instanceof Date ? new Date(dateValue) : new Date(dateValue);
  if (isNaN(date.getTime())) return null;
  if (type !== '收入' || category !== '薪資') return date;
  const tomorrow = new Date(date);
  tomorrow.setDate(date.getDate() + 1);
  if (tomorrow.getMonth() === date.getMonth()) return date;
  const next = new Date(date);
  next.setDate(1);
  next.setMonth(next.getMonth() + 1);
  return next;
}

function findMarketValues(sheet) {
  const rows = sheet.getDataRange().getValues();
  const values = [];
  for (let i = 0; i < rows.length; i++) {
    const label = String(rows[i][0] || '');
    if (label.indexOf('台幣市值') >= 0) {
      for (let j = 2; j < rows[i].length; j++) {
        const v = parseFloat(rows[i][j]);
        if (!isNaN(v) && v > 100000) { values.push(v); break; }
      }
    }
  }
  return values;
}

function findVtMarketValue(sheet) {
  const rows = sheet.getDataRange().getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][1] || '').trim().toUpperCase() === 'VT') {
      for (let j = 2; j < rows[i].length; j++) {
        const v = parseFloat(rows[i][j]);
        if (!isNaN(v) && v > 100000) return v;
      }
    }
  }
  return 0;
}

function dailyAssetUpdate(force) {
  try {
    const now = new Date();
    const weekday = now.getDay();
    if (!force && (weekday === 0 || weekday === 6)) {
      return { success: true, skipped: true, reason: 'weekend' };
    }
    const ss         = SpreadsheetApp.openById(SS_ID);
    const stockSheet = ss.getSheetByName('1150421台股');
    const fundSheet  = ss.getSheetByName('1150421鉅亨基金');
    const assetSheet = ss.getSheetByName('資產記錄');
    if (!stockSheet || !fundSheet || !assetSheet) return { success: false, error: '找不到必要工作表' };

    // ★ v3.28: 同一天已有記錄時「覆寫」該列，而非跳過
    // 這樣手動執行 dailyUpdateAndPush 可隨時更新到最新股價
    const today = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd');
    let todayRowIndex = -1;  // 1-based row index in sheet, -1 = 尚無今日記錄
    const existingRows = assetSheet.getDataRange().getValues();
    for (let i = 1; i < existingRows.length; i++) {
      const d = existingRows[i][0] instanceof Date
        ? Utilities.formatDate(existingRows[i][0], 'Asia/Taipei', 'yyyy/MM/dd')
        : String(existingRows[i][0]);
      if (d === today) { todayRowIndex = i + 1; break; }  // +1 for 1-based
    }
    }

    const stockSheetValues = findMarketValues(stockSheet);
    const fundSheetValues  = findMarketValues(fundSheet);
    // ★ v3.26: F15 手動覆蓋值需 > 1,000,000 才採用，避免小數值污染
    const stockManual = parseFloat(stockSheet.getRange('F15').getValue()) || 0;
    const stockValue  = stockManual > 1000000 ? stockManual : (stockSheetValues[0] || 0);
    const vtNew = findVtMarketValue(stockSheet);
    const vtValue = vtNew > 0 ? vtNew : (stockSheetValues[1] || 0);
    const fundValue = fundSheetValues[0] || 0;
    // ★ v3.26: 最低合理門檻，任一市值 < 100,000 視為異常資料
    const MIN_VALID = 100000;
    if (stockValue < MIN_VALID || fundValue < MIN_VALID || vtValue < MIN_VALID) {
      const missing = [];
      if (stockValue < MIN_VALID) missing.push('台股市值(' + stockValue + ')');
      if (fundValue  < MIN_VALID) missing.push('基金市值(' + fundValue + ')');
      if (vtValue    < MIN_VALID) missing.push('VT市值(' + vtValue + ')');
      return { success: true, skipped: true, reason: 'invalid_data', missing };
    }

    const fixed = getFixedAssets();
    const debts = getDebts();
    const installments = getInstallments();
    // ★ v3.27: 與 renderAllAsset() 使用相同邏輯，避免雙重計算
    // 負債管理裡的分期追蹤列（富邦人壽分期等）已被 installments.totalRemaining 覆蓋，
    // 故此處 ccRows 只取信用卡帳單列（排除 分期|人壽|保険費 關鍵字）
    const mortgageBalance = (debts.byType && debts.byType['房貸'] && debts.byType['房貸'].balance) || 0;
    const ccRows = debts.rows.filter(function(d) {
      return d.type !== '房貸' && !/分期|人壽|保險費/.test(d.type);
    });
    const ccTotal = ccRows.reduce(function(s, d) { return s + d.balance; }, 0);
    const totalDebt = mortgageBalance + ccTotal + installments.totalRemaining;

    const total    = stockValue + vtValue + fundValue + fixed.cash + fixed.husbandRetire + fixed.wifeRetire;
    const netWorth = total - totalDebt;
    const progress = Math.round(netWorth / RETIRE_GOAL_GROSS * 10000) / 100;
    const gap      = RETIRE_GOAL_GROSS - netWorth;
    const newRow = [today, stockValue, vtValue, fundValue, fixed.cash, fixed.husbandRetire, fixed.wifeRetire, total, netWorth, progress, gap];
    if (todayRowIndex > 0) {
      assetSheet.getRange(todayRowIndex, 1, 1, newRow.length).setValues([newRow]);  // 覆寫今日列
    } else {
      assetSheet.appendRow(newRow);  // 新增
    }
    return { success: true, date: today, stockValue, vtValue, fundValue, total, netWorth, totalDebt, progress, gap, forced: !!force, updated: todayRowIndex > 0 };
  } catch(e) { return { success: false, error: e.message }; }
}

function pushSnapshotToGitHub() {
  try {
    const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
    if (!token) return { success: false, error: 'GITHUB_TOKEN 未設定' };
    const assetResult   = getAsset();
    const historyResult = getHistory();
    const expenseResult = getExpense();
    const monthlyResult = monthlyAnalysis(1);
    const closingResult = getClosing(6);
    const debtsData     = getDebts();
    const fixedData     = getFixedAssets();
    const installData   = getInstallments();
    const historyData = (historyResult.success && historyResult.data) ? historyResult.data.slice(-90) : [];
    const expenseData = (expenseResult.success && expenseResult.data) ? expenseResult.data.slice(-120) : [];
    const snapshot = {
      generated: Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm'),
      version:   '3.21',
      asset:     assetResult.success ? assetResult.data : null,
      history:   historyData,
      expense:   expenseData,
      monthly:   monthlyResult.success ? monthlyResult : null,
      closing:   closingResult.success ? closingResult.data : [],
      debts:     debtsData,
      fixedAssets: fixedData,
      installments: installData,
      constants: {
        retireGoalGross: RETIRE_GOAL_GROSS,
        retireGoalNet:   RETIRE_GOAL_NET,
        mortgage:        debtsData.mortgage,
        totalDebt:       debtsData.totalDebt,
        creditCard:      debtsData.creditCard,
        installmentRemaining: installData.totalRemaining,
        husbandRetire:   fixedData.husbandRetire,
        wifeRetire:      fixedData.wifeRetire,
        cash:            fixedData.cash,
        fixedDeposit:    fixedData.fixedDeposit,
        liquidCash:      fixedData.liquidCash,
      }
    };
    const contentB64 = Utilities.base64Encode(JSON.stringify(snapshot), Utilities.Charset.UTF_8);
    const apiBase = 'https://api.github.com/repos/' + GITHUB_REPO + '/contents/' + SNAPSHOT_PATH;
    const headers = {
      'Authorization': 'token ' + token,
      'Accept':        'application/vnd.github.v3+json',
      'User-Agent':    'RainbowCFO-AppsScript'
    };
    let existingSha = '';
    const getRes = UrlFetchApp.fetch(apiBase, { method: 'get', headers, muteHttpExceptions: true });
    if (getRes.getResponseCode() === 200) existingSha = JSON.parse(getRes.getContentText()).sha || '';
    const payload = {
      message: 'auto: daily snapshot ' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm'),
      content: contentB64
    };
    if (existingSha) payload.sha = existingSha;
    const putRes = UrlFetchApp.fetch(apiBase, {
      method:  'put',
      headers: Object.assign({}, headers, { 'Content-Type': 'application/json' }),
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const code = putRes.getResponseCode();
    if (code === 200 || code === 201) return { success: true, httpCode: code, generated: snapshot.generated };
    return { success: false, httpCode: code, error: putRes.getContentText().substring(0, 300) };
  } catch(e) { return { success: false, error: e.message }; }
}

function dailyUpdateAndPush() {
  const updateResult = dailyAssetUpdate();
  if (updateResult.success && !updateResult.skipped) {
    Utilities.sleep(3000);
    const pushResult = pushSnapshotToGitHub();
    return { update: updateResult, push: pushResult };
  } else {
    return { update: updateResult, pushSkipped: true };
  }
}

function monthEndClosing(year, month) {
  try {
    const now = new Date();
    if (!year)  year  = now.getFullYear();
    if (!month) month = now.getMonth() + 1;
    const ss = SpreadsheetApp.openById(SS_ID);
    const expSheet = ss.getSheetByName('收支記帳');
    if (!expSheet) return { success: false, error: '找不到收支記帳' };
    let closingSheet = ss.getSheetByName(CLOSING_SHEET);
    if (!closingSheet) {
      closingSheet = ss.insertSheet(CLOSING_SHEET);
      closingSheet.appendRow(['結算月份','總收入','總支出','結餘','預算總額','使用率%','超支類別','節餘類別','關鍵建議','結算時間','本月還卡費']);
    }
    const startDate = new Date(year, month - 1, 1);
    const endDate   = new Date(year, month, 0, 23, 59, 59);
    const monthLabel = year + '/' + String(month).padStart(2, '0');
    const expRows = expSheet.getDataRange().getValues().slice(1);
    const monthExp = expRows.filter(r => {
      if (!r[0]) return false;
      const d = getEffectiveDate(r[0], r[1], r[2]);
      if (!d) return false;
      return d >= startDate && d <= endDate;
    });
    let totalIncome = 0, totalExpense = 0, cardPayment = 0;
    const catMap = {};
    monthExp.forEach(r => {
      const type = r[1], cat = r[2], amt = parseFloat(r[4]) || 0;
      if (type === '收入') { totalIncome += amt; return; }
      if (cat === '還卡費') { cardPayment += amt; return; }
      catMap[cat] = (catMap[cat] || 0) + amt;
      totalExpense += amt;
    });
    const balance = totalIncome - totalExpense;
    const utilization = Math.round(totalExpense / BUDGET_TOTAL * 100);
    const overspent = [];
    const saved = [];
    Object.keys(BUDGET_DEF).forEach(cat => {
      const actual = catMap[cat] || 0;
      const budget = BUDGET_DEF[cat];
      if (budget <= 0) return;
      const pct = actual / budget * 100;
      if (pct > 100) overspent.push(cat + '：實 NT$' + Math.round(actual).toLocaleString() + '（超 ' + Math.round(pct - 100) + '%）');
      if (pct < 50 && budget > 3000) saved.push(cat + '：實 NT$' + Math.round(actual).toLocaleString() + '（僅用 ' + Math.round(pct) + '%）');
    });
    const suggestions = [];
    if (overspent.length === 0) suggestions.push('✅ 所有類別預算都在控制中');
    else if (overspent.length <= 2) suggestions.push('⚠️ 少數類別超支，下月可微調');
    else suggestions.push('🚨 多項超支（' + overspent.length + ' 類）');
    if (balance > 0) suggestions.push('💰 月結餘 NT$' + balance.toLocaleString() + '，建議轉投資');
    if (balance < 0) suggestions.push('🔴 月赤字 NT$' + Math.abs(balance).toLocaleString());
    if (utilization < 70) suggestions.push('📉 預算使用率僅 ' + utilization + '%');
    if (utilization > 110) suggestions.push('📈 預算使用率 ' + utilization + '%，整體緊縮');
    if (cardPayment > 0) suggestions.push('💸 本月實際還卡 NT$' + cardPayment.toLocaleString());
    const closingDate = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm');
    const existingRows = closingSheet.getDataRange().getValues();
    let existingIdx = -1;
    for (let i = 1; i < existingRows.length; i++) {
      if (existingRows[i][0] === monthLabel) { existingIdx = i; break; }
    }
    const newRow = [monthLabel, Math.round(totalIncome), Math.round(totalExpense), Math.round(balance),
      BUDGET_TOTAL, utilization, overspent.join('\n')||'無', saved.join('\n')||'無', suggestions.join('\n'), closingDate, Math.round(cardPayment)];
    if (existingIdx > 0) closingSheet.getRange(existingIdx + 1, 1, 1, newRow.length).setValues([newRow]);
    else closingSheet.appendRow(newRow);
    return { success: true, month: monthLabel, totalIncome: Math.round(totalIncome), totalExpense: Math.round(totalExpense),
      balance: Math.round(balance), utilization, cardPayment: Math.round(cardPayment), overspent, saved, suggestions, closingDate };
  } catch(e) { return { success: false, error: e.message }; }
}

function monthEndClosingAuto() {
  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);
  if (tomorrow.getMonth() === today.getMonth()) {
    return { success: true, skipped: true, reason: 'not_last_day' };
  }
  const result = monthEndClosing();
  Utilities.sleep(2000);
  pushSnapshotToGitHub();
  return result;
}

function getClosing(months) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(CLOSING_SHEET);
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, data: [] };
    const rows = sheet.getDataRange().getValues();
    const dataRows = rows.slice(1).filter(r => r[0]);
    const limited = dataRows.slice(-(months || 6));
    const data = limited.map(r => ({
      month: String(r[0]||''), totalIncome: parseFloat(r[1])||0, totalExpense: parseFloat(r[2])||0,
      balance: parseFloat(r[3])||0, budgetTotal: parseFloat(r[4])||0, utilization: parseFloat(r[5])||0,
      overspent: String(r[6]||''), saved: String(r[7]||''), suggestions: String(r[8]||''), closingDate: String(r[9]||''),
      cardPayment: parseFloat(r[10])||0,
    }));
    return { success: true, data: data };
  } catch(e) { return { success: false, error: e.message }; }
}

function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'dailyAssetUpdate' || fn === 'dailyUpdateAndPush' ||
        fn === 'monthEndClosingAuto' || fn === 'monthlyInstallmentAutoPost' ||
        fn === 'monthlyMortgageAutoPost') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('dailyUpdateAndPush').timeBased().everyDays(1).atHour(13).create();
  ScriptApp.newTrigger('monthEndClosingAuto').timeBased().everyDays(1).atHour(23).create();
  ScriptApp.newTrigger('monthlyInstallmentAutoPost').timeBased().onMonthDay(1).atHour(2).create();
  ScriptApp.newTrigger('monthlyMortgageAutoPost').timeBased().onMonthDay(19).atHour(8).create();
}

function monthlyInstallmentAutoPost(force) {
  try {
    const today = new Date();
    if (!force && today.getDate() !== 1) {
      return { success: true, skipped: true, reason: 'not_first_day', date: today.getDate() };
    }
    const inst = getInstallments();
    if (!inst.rows || !inst.rows.length) {
      return { success: true, skipped: true, reason: 'no_installments' };
    }

    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName('收支記帳');
    const debtSheet = ss.getSheetByName(INSTALLMENT_SHEET);
    if (!sheet || !debtSheet) return { success: false, error: '找不到必要工作表' };

    const dateStr = Utilities.formatDate(today, 'Asia/Taipei', 'yyyy-MM-dd');
    const ym = Utilities.formatDate(today, 'Asia/Taipei', 'yyyy-MM');

    const existing = sheet.getDataRange().getValues();
    const alreadyPosted = new Set();
    existing.slice(1).forEach(r => {
      if (!r[0]) return;
      const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
      if (isNaN(d.getTime())) return;
      if (Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM') !== ym) return;
      if (String(r[5]||'').trim() !== '分期預付') return;
      const desc = String(r[3]||'');
      inst.rows.forEach(it => {
        if (desc.indexOf(it.item) >= 0) alreadyPosted.add(it.item);
      });
    });

    const posted = [], skipped = [];
    inst.rows.forEach(it => {
      if (!it.isActive) { skipped.push(it.item + '（已繳完）'); return; }
      if (alreadyPosted.has(it.item)) { skipped.push(it.item + '（本月已記）'); return; }
      const desc = it.item + '（第' + it.currentPeriod + '/' + it.totalPeriods + '期）';
      sheet.appendRow([dateStr, '支出', it.category, desc, it.monthly, '分期預付', '一般']);
      debtSheet.getRange(it.rowIndex, 7).setValue(it.passedPeriods + 1);
      // ★ v3.25: Removed updateDebtBalance(it.card, it.monthly) — installments tracked via syncInstallmentDebtBalances()
      posted.push(desc + ' ' + it.monthly);
    });

    Logger.log('monthlyInstallmentAutoPost：post ' + posted.length + ' / skip ' + skipped.length);
    // 自動同步 負債管理 的分期餘額（連動）
    if (posted.length > 0) {
      try { syncInstallmentDebtBalances(); } catch(se) { Logger.log('sync error: ' + se.message); }
    }
    return { success: true, date: dateStr, posted, skipped, forced: !!force };
  } catch(e) {
    Logger.log('monthlyInstallmentAutoPost error: ' + e.message);
    return { success: false, error: e.message };
  }
}

function postInstallmentsNow() {
  return monthlyInstallmentAutoPost(true);
}

function monthlyMortgageAutoPost(force) {
  try {
    const today = new Date();
    if (!force && today.getDate() !== 19) {
      return { success: true, skipped: true, reason: 'not_payday', date: today.getDate() };
    }
    const debts = getDebts();
    const mortgage = debts.byType['房貸'];
    if (!mortgage || mortgage.balance <= 0) {
      return { success: true, skipped: true, reason: 'no_mortgage' };
    }

    const ACCOUNT = '富邦銀行';
    const balance = mortgage.balance;
    let rate = parseFloat(mortgage.rate) || 2.24;
    if (rate > 1) rate = rate / 100;
    const monthly = parseFloat(mortgage.monthly) || 15515;

    const interestAmt = Math.round(balance * rate / 12);
    const principalAmt = monthly - interestAmt;
    if (principalAmt <= 0) {
      return { success: false, error: '本金計算異常：' + principalAmt };
    }

    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName('收支記帳');
    if (!sheet) return { success: false, error: '找不到收支記帳' };
    const ym = Utilities.formatDate(today, 'Asia/Taipei', 'yyyy-MM');
    const existing = sheet.getDataRange().getValues();
    const alreadyPosted = existing.slice(1).some(r => {
      if (!r[0]) return false;
      const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
      if (isNaN(d.getTime()) || Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM') !== ym) return false;
      return String(r[2]||'').trim() === '房貸' && String(r[3]||'').includes('房貸本金');
    });
    if (alreadyPosted) return { success: true, skipped: true, reason: 'already_posted_this_month' };

    const dateStr = Utilities.formatDate(today, 'Asia/Taipei', 'yyyy-MM-dd');

    sheet.appendRow([dateStr, '支出', '房貸', '房貸本金', principalAmt, ACCOUNT, '一般']);
    updateDebtBalance('房貸', -principalAmt);
    updateLiquidCash(-principalAmt, ACCOUNT);

    sheet.appendRow([dateStr, '支出', '利息費', '房貸利息', interestAmt, ACCOUNT, '一般']);
    updateLiquidCash(-interestAmt, ACCOUNT);

    Logger.log('monthlyMortgageAutoPost：本金 ' + principalAmt + ' + 利息 ' + interestAmt + ' = ' + monthly);
    return { success: true, date: dateStr, principal: principalAmt, interest: interestAmt, total: monthly, balanceAfter: balance - principalAmt };
  } catch(e) {
    Logger.log('monthlyMortgageAutoPost error: ' + e.message);
    return { success: false, error: e.message };
  }
}

function postMortgageNow() {
  const r = monthlyMortgageAutoPost(true);
  Logger.log(JSON.stringify(r));
  return r;
}

function testForceUpdateNow() {
  const result = dailyAssetUpdate(true);
  Logger.log(JSON.stringify(result));
  return result;
}

function monthlyAnalysis(months) {
  try {
    const n = months || 12;
    const ss = SpreadsheetApp.openById(SS_ID);
    const expSheet   = ss.getSheetByName('收支記帳');
    const assetSheet = ss.getSheetByName('資產記錄');
    const expRows = expSheet.getDataRange().getValues().slice(1);
    const cutoff  = new Date();
    cutoff.setMonth(cutoff.getMonth() - n);
    const recentExp = expRows.filter(r => {
      if (!r[0]) return false;
      const d = getEffectiveDate(r[0], r[1], r[2]);
      return d && d >= cutoff;
    });
    const catMap = {};
    let totalIncome = 0, totalExpense = 0, cardPayment = 0;
    recentExp.forEach(r => {
      const type = r[1], cat = r[2], amt = parseFloat(r[4]) || 0;
      if (type === '收入') { totalIncome += amt; return; }
      if (cat === '還卡費') { cardPayment += amt; return; }
      if (!catMap[cat]) catMap[cat] = 0;
      catMap[cat] += amt;
      totalExpense += amt;
    });
    const avgMonthlyExpense = totalExpense / n;
    const avgMonthlyIncome  = totalIncome  / n;
    const avgMonthlyCardPay = cardPayment / n;
    const catAnalysis = Object.keys(BUDGET_DEF).map(cat => {
      const actual = (catMap[cat] || 0) / n;
      const budget = BUDGET_DEF[cat];
      return { cat, actual: Math.round(actual), budget,
        diff: Math.round(actual - budget),
        pct: budget > 0 ? Math.round(actual / budget * 100) : 0 };
    });
    const assetRows = assetSheet.getDataRange().getValues().slice(1).filter(r => r[0]);
    const latestRow = assetRows[assetRows.length - 1];
    const cutoffAsset = new Date();
    cutoffAsset.setMonth(cutoffAsset.getMonth() - n);
    const recentAssets = assetRows.filter(r => {
      const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
      return d >= cutoffAsset;
    });
    const oldestRow = recentAssets[0] || latestRow;
    const netWorthNow  = parseFloat(latestRow[8]) || 0;
    const netWorthPrev = parseFloat(oldestRow[8]) || 0;
    const netWorthChg  = netWorthNow - netWorthPrev;
    const progressNow  = parseFloat(latestRow[9]) || 0;
    const MONTHLY_PENSION = 49362, MONTHLY_BUDGET_RET = 63000;
    const POST_SHORTFALL = MONTHLY_BUDGET_RET - MONTHLY_PENSION;
    const monthlyReturn = netWorthNow * 0.03 / 12;
    const sustainable = monthlyReturn >= POST_SHORTFALL;
    const suggestions = [];
    catAnalysis.forEach(c => {
      if (c.pct > 120) suggestions.push('⚠️ ' + c.cat + '：月均 NT$' + c.actual.toLocaleString() + '，超預算 ' + (c.pct - 100) + '%');
      if (c.pct < 50 && c.budget > 3000) suggestions.push('✅ ' + c.cat + '：節省優異（僅用 ' + c.pct + '%）');
    });
    if (netWorthChg > 0) suggestions.push('📈 近' + n + '個月資產增加 NT$' + Math.round(netWorthChg).toLocaleString());
    if (avgMonthlyExpense > BUDGET_TOTAL * 1.1) suggestions.push('🔴 月均支出超預算10%');
    if (avgMonthlyCardPay > 0) suggestions.push('💸 月均實際還卡 NT$' + Math.round(avgMonthlyCardPay).toLocaleString() + '（含旅遊/分期）');
    if (sustainable) suggestions.push('🏆 3%年報酬下投資帳戶永不歸零！月收益 NT$' + Math.round(monthlyReturn).toLocaleString() + ' > 缺口 NT$' + POST_SHORTFALL.toLocaleString());
    return {
      success: true, period: n,
      generated: Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd'),
      income:  { monthlyAvg: Math.round(avgMonthlyIncome) },
      expense: { monthlyAvg: Math.round(avgMonthlyExpense),
        cardPayment: Math.round(avgMonthlyCardPay),
        budgetTotal: BUDGET_TOTAL,
        utilization: Math.round(avgMonthlyExpense / BUDGET_TOTAL * 100), byCategory: catAnalysis },
      asset: { netWorth: Math.round(netWorthNow), change: Math.round(netWorthChg), progress: progressNow,
        retireGoalGross: RETIRE_GOAL_GROSS, retireGoalNet: RETIRE_GOAL_NET,
        gap: Math.round(RETIRE_GOAL_GROSS - netWorthNow) },
      debts: getDebts(),
      fixedAssets: getFixedAssets(),
      installments: getInstallments(),
      retirement: { monthlyPension: MONTHLY_PENSION, monthlyBudget: MONTHLY_BUDGET_RET,
        shortfall: POST_SHORTFALL, monthlyReturn: Math.round(monthlyReturn), sustainable },
      suggestions
    };
  } catch(e) { return { success: false, error: e.message }; }
}
