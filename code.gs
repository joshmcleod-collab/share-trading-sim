// ==========================================
// SETUP & GLOBALS
// ==========================================
const SHEET_ID = '1SISgF_XNE0YJx9AfzbaLT9lFxa7PXhipxcIKIDG_its';

// Initialize globally to prevent opening the spreadsheet 30+ times per second
const SS = SpreadsheetApp.openById(SHEET_ID);

// ==========================================
// WEB APP ROUTING
// ==========================================
function doGet(e) {
  if (e.parameter.view === 'teacher') {
    return HtmlService.createHtmlOutputFromFile('Teacher')
        .setTitle('Teacher Dashboard - Economy Tracker')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
    return HtmlService.createHtmlOutputFromFile('Student')
        .setTitle('Student Portal - Economy Tracker')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

// ==========================================
// HELPER FUNCTIONS
// ==========================================
function getStudentRow(studentId) {
  const sheet = SS.getSheetByName('Student Economy Tracker') || SS.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const targetId = String(studentId).trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === targetId) { 
      return { rowIndex: i, rowData: data[i], sheet: sheet };
    }
  }
  return null;
}

function parseCleanNumber(val) {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  const cleaned = String(val).replace(/[^0-9.-]+/g, '');
  const num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

// ==========================================
// TRANSACTION LOGGING FUNCTIONS
// ==========================================
function logTransaction(studentId, action, detailsObj) {
  let logSheet = SS.getSheetByName('TransactionLogs'); 
  if (!logSheet) {
    logSheet = SS.insertSheet('TransactionLogs');
    logSheet.appendRow(['Date', 'Student ID', 'Action', 'Details']);
  }
  const timestamp = new Date();
  const detailsString = JSON.stringify(detailsObj);
  logSheet.appendRow([timestamp, String(studentId).trim(), action, detailsString]);
}

function getTransactionLogs(studentId = null) {
  const logSheet = SS.getSheetByName('TransactionLogs'); 
  if (!logSheet) return [];
  
  const data = logSheet.getDataRange().getValues();
  data.shift(); // Remove header
  let filteredLogs = data;
  
  if (studentId) {
    const targetId = String(studentId).trim().toLowerCase();
    filteredLogs = data.filter(log => String(log[1]).trim().toLowerCase() === targetId);
  }
  
  return filteredLogs.map(log => {
    let dateStr = log[0] instanceof Date ? log[0].toLocaleString() : log[0];
    return [dateStr, log[1], log[2], log[3]];
  });
}

// ==========================================
// PASSWORD HASHING & AUTHENTICATION
// ==========================================
function hashPassword_(plainPassword, existingSalt) {
  const salt = existingSalt || Utilities.getUuid().replace(/-/g, '');
  const rawHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    salt + plainPassword,
    Utilities.Charset.UTF_8
  );
  const hexHash = rawHash
    .map(b => ('0' + ((b < 0 ? b + 256 : b).toString(16))).slice(-2))
    .join('');
  return `${salt}:${hexHash}`;
}

function verifyPasswordForRow_(rowData, inputPassword) {
  const plainPass = String(rowData[8] || '').trim();      // Column I (index 8)
  const encryptedPass = String(rowData[9] || '').trim();  // Column J (index 9)

  // 1. Check encrypted password in Column J first
  if (encryptedPass && encryptedPass.includes(':')) {
    const [salt] = encryptedPass.split(':');
    return hashPassword_(inputPassword, salt) === encryptedPass;
  }

  // 2. Fallback to plaintext in Column I
  return plainPass !== '' && inputPassword === plainPass;
}

function authenticateStudent(studentId, password) {
  const cleanId = String(studentId).trim();
  const cleanPass = String(password).trim();
  const student = getStudentRow(cleanId);
  
  if (!student) return { success: false, error: "Student ID not found." };
  
  if (verifyPasswordForRow_(student.rowData, cleanPass)) {
    const canonicalId = String(student.rowData[0]).trim();
    return { success: true, data: getStudentDataById(canonicalId) };
  }
  
  logTransaction(cleanId, 'FAILED_LOGIN', { attempted: new Date().toISOString() });
  return { success: false, error: "Invalid password." };
}

function changeStudentPassword(studentId, oldPassword, newPassword) {
  const cleanId = String(studentId).trim();
  const cleanOld = String(oldPassword).trim();
  const cleanNew = String(newPassword).trim();

  if (!cleanNew || cleanNew.length < 6) {
    return { success: false, error: 'New password must be at least 6 characters.' };
  }

  const student = getStudentRow(cleanId);
  if (!student) {
    return { success: false, error: 'Student ID not found.' };
  }

  if (!verifyPasswordForRow_(student.rowData, cleanOld)) {
    return { success: false, error: 'Current password is incorrect.' };
  }

  const canonicalId = String(student.rowData[0]).trim();
  const encrypted = hashPassword_(cleanNew);
  const rowNum = student.rowIndex + 1;

  // Write encrypted password to Column J (Col 10) and clear plaintext in Column I (Col 9)
  student.sheet.getRange(rowNum, 10).setValue(encrypted);
  student.sheet.getRange(rowNum, 9).clearContent();

  CacheService.getScriptCache().remove('STUDENT_TRACKER');
  logTransaction(canonicalId, 'PASSWORD_CHANGED', { migratedToEncrypted: true });

  return { success: true, message: 'Password updated and encrypted! Please sign in.' };
}

// ==========================================
// DATA RETRIEVAL (BATCH OPTIMIZED)
// ==========================================
function getStudentData() {
  const studentSheet = SS.getSheetByName('Student Economy Tracker') || SS.getSheets()[0];
  const studentData = studentSheet.getDataRange().getValues();
  
  const stockSheet = SS.getSheetByName('Stocks');
  const stocksData = stockSheet ? stockSheet.getDataRange().getValues() : [];
  
  let holdingsSheet = SS.getSheetByName('Holdings');
  if (!holdingsSheet) holdingsSheet = SS.insertSheet('Holdings');
  const holdingsData = holdingsSheet.getDataRange().getValues();
  
  if (stocksData.length > 0) stocksData.shift();
  const liveStocks = {};
  stocksData.forEach(row => {
    const ticker = String(row[0]).trim();
    const locked = parseCleanNumber(row[5]);
    const live = parseCleanNumber(row[4]);
    const price = locked > 0 ? locked : live;
    if (!liveStocks[ticker] || price > 0) {
      liveStocks[ticker] = price;
    }
  });
  
  if (holdingsData.length > 0) holdingsData.shift();
  const studentPortfolios = {};
  holdingsData.forEach(row => {
    const studId = String(row[0]).trim().toLowerCase();
    const ticker = String(row[1]).trim();
    const qty = parseCleanNumber(row[2]);
    
    if (!studentPortfolios[studId]) studentPortfolios[studId] = 0;
    if (liveStocks[ticker]) {
      studentPortfolios[studId] += (qty * liveStocks[ticker]);
    }
  });
  
  studentData.shift(); 
  return studentData.map(row => {
    const studId = String(row[0]).trim();
    const name = String(row[1]).trim();
    const cash = parseCleanNumber(row[3]);
    const termDeposit = parseCleanNumber(row[4]);
    const savings = parseCleanNumber(row[5]);
    const bonds = parseCleanNumber(row[6]);
    const stockValue = studentPortfolios[studId.toLowerCase()] || 0;
    
    return {
      studentId: studId,
      name: name,
      className: row[2],
      cash: cash,
      termDeposit: termDeposit,
      savings: savings,
      bonds: bonds,
      total: cash + termDeposit + savings + bonds + stockValue
    };
  });
}

function getStudentDataById(studentId) {
  const cleanId = String(studentId).trim();
  const student = getStudentRow(cleanId);
  if (!student) return null;
  const row = student.rowData;
  const canonicalId = String(row[0]).trim();
  
  const cash = parseCleanNumber(row[3]);
  const termDeposit = parseCleanNumber(row[4]);
  const savings = parseCleanNumber(row[5]);
  const bonds = parseCleanNumber(row[6]);
  const fixedIncomeTotal = termDeposit + savings + bonds;

  return {
    studentId: canonicalId,
    name: String(row[1]).trim(),
    className: row[2],
    cash: cash,
    termDeposit: termDeposit,
    savings: savings,
    bonds: bonds,
    total: calculateLiveNetWorth(canonicalId, cash + fixedIncomeTotal)
  };
}

function getLeaderboard() {
  const data = getStudentData();
  return data.sort((a, b) => b.total - a.total).map((student, rank) => ({
    ...student,
    rank: rank + 1
  }));
}

// ==========================================
// TRADING ENGINE & HOLDINGS
// ==========================================
function getAvailableStocks() {
  const cache = CacheService.getScriptCache();
  const cachedStocks = cache.get('MARKET_STOCKS');
  if (cachedStocks) return JSON.parse(cachedStocks);

  const stockSheet = SS.getSheetByName('Stocks');
  if (!stockSheet) return [];

  const lastRow = stockSheet.getLastRow();
  if (lastRow < 2) return [];

  const data = stockSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const richTextB = stockSheet.getRange(2, 2, lastRow - 1, 1).getRichTextValues();

  // Columns: A=0 (ticker), B=1 (name + link), C=2 (category), D=3 (marketType), E=4 (livePrice), F=5 (lockedClose)
  const mappedStocks = data.map((row, i) => {
    const richVal = richTextB[i][0];
    let url = richVal ? richVal.getLinkUrl() : '';
    if (!url && richVal && richVal.getRuns().length > 0) {
      url = richVal.getRuns()[0].getLinkUrl() || '';
    }

    const livePrice = parseCleanNumber(row[4]);
    const lockedClose = parseCleanNumber(row[5]);

    return {
      ticker: String(row[0]).trim(),
      name: String(row[1]).trim(),
      url: url,
      category: String(row[2]).trim(),
      marketType: String(row[3]).trim(),
      livePrice: livePrice,
      lockedClose: lockedClose,
      price: lockedClose > 0 ? lockedClose : livePrice
    };
  });

  cache.put('MARKET_STOCKS', JSON.stringify(mappedStocks), 300);
  return mappedStocks;
}

function calculateLiveNetWorth(studentId, cash) {
  const holdings = getStudentHoldings(studentId);
  const stocks = getAvailableStocks();
  let totalStockValue = 0;

  holdings.forEach(h => {
    const liveStock = stocks.find(s => s.ticker === h.ticker && s.price > 0) 
                   || stocks.find(s => s.ticker === h.ticker);
    if (liveStock && !isNaN(liveStock.price)) {
      totalStockValue += (h.quantity * liveStock.price);
    }
  });

  return parseCleanNumber(cash) + totalStockValue;
}

function getStudentHoldings(studentId) {
  let holdingsSheet = SS.getSheetByName('Holdings');
  if (!holdingsSheet) {
    holdingsSheet = SS.insertSheet('Holdings');
    holdingsSheet.appendRow(['Student ID', 'Stock ID', '#', 'Value']);
    return [];
  }
  
  const data = holdingsSheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const targetId = String(studentId).trim().toLowerCase();
  data.shift(); // Remove header
  return data
    .filter(row => String(row[0]).trim().toLowerCase() === targetId) 
    .map(row => ({
      ticker: String(row[1]).trim(),
      quantity: parseCleanNumber(row[2]),
      avgPrice: parseCleanNumber(row[3])
    }));
}

function processTrade(studentId, action, ticker, tradeValue) {
  const lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(15000); 
    
    const cleanId = String(studentId).trim();
    tradeValue = Number(tradeValue);
    if (tradeValue <= 0) return { success: false, error: "Amount must be greater than zero." };

    const stocks = getAvailableStocks();
    const stock = stocks.find(s => s.ticker === ticker && s.price > 0) || stocks.find(s => s.ticker === ticker);
    if (!stock || stock.price <= 0) return { success: false, error: "Asset price unavailable." };

    const student = getStudentRow(cleanId);
    if (!student) return { success: false, error: "Student ID not found." };

    const canonicalId = String(student.rowData[0]).trim();
    const currentCash = parseCleanNumber(student.rowData[3]);
    const holdingsSheet = SS.getSheetByName('Holdings');
    let holdingsData = holdingsSheet.getDataRange().getValues();
    
    let holdingRowIndex = -1;
    let currentQty = 0;
    for (let i = 1; i < holdingsData.length; i++) {
      if (String(holdingsData[i][0]).trim().toLowerCase() === canonicalId.toLowerCase() && String(holdingsData[i][1]).trim() === ticker) {
        holdingRowIndex = i + 1;
        currentQty = parseCleanNumber(holdingsData[i][2]);
        break;
      }
    }

    if (action === 'BUY') {
      if (currentCash < tradeValue) return { success: false, error: "Insufficient funds." };
      
      const quantityBought = tradeValue / stock.price; 
      
      // Deduct Cash
      student.sheet.getRange(student.rowIndex + 1, 4).setValue(currentCash - tradeValue);
      
      // Update Holdings
      if (holdingRowIndex > -1) {
        holdingsSheet.getRange(holdingRowIndex, 3).setValue(currentQty + quantityBought);
      } else {
        holdingsSheet.appendRow([canonicalId, ticker, quantityBought, stock.price]);
      }
      
      logTransaction(canonicalId, 'BUY', { ticker: ticker, amountInvested: tradeValue, sharesBought: quantityBought, price: stock.price });
      return { success: true, message: `Successfully bought $${tradeValue.toFixed(2)} of ${ticker}.` };

    } else if (action === 'SELL') {
      if (currentQty <= 0) return { success: false, error: "You don't own any of this asset." };

      let quantityToSell = tradeValue / stock.price;
      let actualRevenue = tradeValue;

      if (quantityToSell >= currentQty - 0.000001) {
        quantityToSell = currentQty;
        actualRevenue = currentQty * stock.price;
      }
      
      // Add Cash
      student.sheet.getRange(student.rowIndex + 1, 4).setValue(currentCash + actualRevenue);
      
      // Update Holdings
      if (quantityToSell === currentQty) {
         holdingsSheet.getRange(holdingRowIndex, 3).setValue(0);
      } else {
         holdingsSheet.getRange(holdingRowIndex, 3).setValue(currentQty - quantityToSell);
      }
      
      logTransaction(canonicalId, 'SELL', { ticker: ticker, amountSold: actualRevenue, sharesSold: quantityToSell, price: stock.price });
      return { success: true, message: `Successfully sold $${actualRevenue.toFixed(2)} of ${ticker}.` };
    }
    
  } catch (e) {
    return { success: false, error: "The market is busy processing other trades. Please try again in 5 seconds." };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// FIXED INCOME & SAVINGS ALLOCATION
// ==========================================
function processAllocation(studentId, action, accountType, amount) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    const cleanId = String(studentId).trim();
    amount = Number(amount);
    if (isNaN(amount) || amount <= 0) {
      return { success: false, error: "Amount must be greater than zero." };
    }

    // Map account type to sheet column number (E=5, F=6, G=7)
    const columnMap = {
      'TERM_DEPOSIT': { col: 5, index: 4, label: 'Term Deposit' },
      'SAVINGS':      { col: 6, index: 5, label: 'Savings' },
      'BONDS':        { col: 7, index: 6, label: 'Bonds' }
    };

    const target = columnMap[accountType];
    if (!target) return { success: false, error: "Invalid account type." };

    const student = getStudentRow(cleanId);
    if (!student) return { success: false, error: "Student ID not found." };

    const canonicalId = String(student.rowData[0]).trim();
    const currentCash = parseCleanNumber(student.rowData[3]);
    const currentAccountBal = parseCleanNumber(student.rowData[target.index]);
    const rowNum = student.rowIndex + 1;

    if (action === 'DEPOSIT') {
      if (currentCash < amount) {
        return { success: false, error: "Insufficient Liquid Cash." };
      }
      student.sheet.getRange(rowNum, 4).setValue(currentCash - amount);
      student.sheet.getRange(rowNum, target.col).setValue(currentAccountBal + amount);
      logTransaction(canonicalId, 'ALLOCATE_IN', { account: target.label, amount: amount });
      return { success: true, message: `Allocated $${amount.toFixed(2)} to ${target.label}.` };

    } else if (action === 'WITHDRAW') {
      if (currentAccountBal < amount) {
        return { success: false, error: `Insufficient funds in ${target.label}.` };
      }
      student.sheet.getRange(rowNum, 4).setValue(currentCash + amount);
      student.sheet.getRange(rowNum, target.col).setValue(currentAccountBal - amount);
      logTransaction(canonicalId, 'ALLOCATE_OUT', { account: target.label, amount: amount });
      return { success: true, message: `Withdrew $${amount.toFixed(2)} from ${target.label} to Liquid Cash.` };
    }

    return { success: false, error: "Invalid action." };
  } catch (e) {
    return { success: false, error: "Server busy. Please try again in a few seconds." };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// MARKET & SHARE PRICE LOCKING (HOURLY TRIGGER)
// ==========================================
function lockSharePrices() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return; // Prevent overlapping runs

  try {
    const cache = CacheService.getScriptCache();
    const stockSheet = SS.getSheetByName('Stocks');

    if (stockSheet) {
      const lastRow = stockSheet.getLastRow();
      if (lastRow >= 2) {
        const numRows = lastRow - 1;

        // 1. Get the live prices from Column E (col 5, starting at row 2)
        const livePrices = stockSheet.getRange(2, 5, numRows, 1).getValues();

        // 2. Paste them as static values into Column F (col 6, Locked Close)
        stockSheet.getRange(2, 6, numRows, 1).setValues(livePrices);

        // 3. Update Column G (col 7, Last Updated) with the current date/time
        const now = new Date();
        const timestamps = Array(numRows).fill([now]);
        stockSheet.getRange(2, 7, numRows, 1).setValues(timestamps);
      }
    }

    // 4. Clear stale stock and student caches so new prices appear immediately
    cache.removeAll(['MARKET_STOCKS', 'STUDENT_TRACKER']);

  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// DAILY ACCOUNT INTEREST (DAILY TRIGGER)
// ==========================================
/**
 * Fetches interest rates from the 'variables' tab (cached for 1 hour)
 */
function getInterestRates() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('VARIABLES_RATES');
  if (cached) {
    return JSON.parse(cached);
  }

  const varSheet = SS.getSheetByName('variables');
  if (!varSheet) return { termDeposit: 0, savings: 0, bonds: 0 };

  const [termRate, savingsRate, bondsRate] = varSheet.getRange('A2:C2').getValues()[0];
  const rates = {
    termDeposit: Number(termRate) || 0,
    savings: Number(savingsRate) || 0,
    bonds: Number(bondsRate) || 0
  };

  cache.put('VARIABLES_RATES', JSON.stringify(rates), 3600);
  return rates;
}

/**
 * Applies daily interest to Term Deposit, Savings, and Bonds on 'Student Economy Tracker'
 */
function applyAccountInterest() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return; // Prevent overlapping runs or collisions with trades

  try {
    const cache = CacheService.getScriptCache();
    // Clear stale rate/student caches before reading so any updated rates apply immediately
    cache.removeAll(['VARIABLES_RATES', 'STUDENT_TRACKER']);

    const trackerSheet = SS.getSheetByName('Student Economy Tracker');
    if (!trackerSheet) return;

    const lastStudentRow = trackerSheet.getLastRow();
    if (lastStudentRow < 2) return;

    const rateObj = getInterestRates();
    const rates = [rateObj.termDeposit, rateObj.savings, rateObj.bonds];
    const numStudents = lastStudentRow - 1;

    // Columns E, F, G (5, 6, 7) on 'Student Economy Tracker'
    const START_COL = 5;
    const range = trackerSheet.getRange(2, START_COL, numStudents, 3);
    const balances = range.getValues();

    const updatedBalances = balances.map(row => {
      return row.map((balance, i) => {
        const current = Number(balance) || 0;
        if (current <= 0) return current;
        return Math.round(current * (1 + rates[i]) * 100) / 100;
      });
    });

    range.setValues(updatedBalances);

    // Clear student data cache again so students immediately see their new balances
    cache.remove('STUDENT_TRACKER');

  } finally {
    lock.releaseLock();
  }
}

function recordDailySnapshot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName("Student Economy Tracker");
  const holdingsSheet = ss.getSheetByName("Holdings");
  const stocksSheet = ss.getSheetByName("Stocks");
  const historySheet = ss.getSheetByName("History");

  const cleanNum = (val) => {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    return Number(String(val).replace(/[^0-9.-]+/g, "")) || 0;
  };

  const today = new Date();
  const dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), "MMM dd");

  // 1. Build a lookup map using Ticker (Col A / [0]) and Locked Close (Col F / [5])
  const stocksData = stocksSheet.getDataRange().getValues();
  const lockedPrices = {};
  for (let i = 1; i < stocksData.length; i++) {
    const ticker = String(stocksData[i][0]).trim(); // Col A: Ticker
    const lockedClose = cleanNum(stocksData[i][5]);  // Col F: Locked Close
    if (ticker) lockedPrices[ticker] = lockedClose;
  }

  // 2. Multiply Holdings Shares (Col C / [2]) by Locked Close from Stocks
  const holdingsData = holdingsSheet.getDataRange().getValues();
  const studentStockTotals = {};
  for (let i = 1; i < holdingsData.length; i++) {
    const studId = String(holdingsData[i][0]).trim(); // Col A: Student ID
    const ticker = String(holdingsData[i][1]).trim(); // Col B: Ticker
    const shares = cleanNum(holdingsData[i][2]);        // Col C: # of shares

    if (studId && ticker) {
      const price = lockedPrices[ticker] || 0;
      studentStockTotals[studId] = (studentStockTotals[studId] || 0) + (shares * price);
    }
  }

  // 3. Add Student Economy Tracker D:G + Locked Close Stock Totals
  const trackerData = trackerSheet.getDataRange().getValues();
  for (let i = 1; i < trackerData.length; i++) {
    const studId = String(trackerData[i][0]).trim(); // Col A: Student ID
    const name = String(trackerData[i][1]).trim();   // Col B: Student Name
    if (!studId || !name) continue;

    const cash = cleanNum(trackerData[i][3]);        // Col D
    const termDeposit = cleanNum(trackerData[i][4]); // Col E
    const savings = cleanNum(trackerData[i][5]);     // Col F
    const bonds = cleanNum(trackerData[i][6]);       // Col G

    const stockValue = studentStockTotals[studId] || 0;
    const totalNetWorth = Math.round((cash + termDeposit + savings + bonds + stockValue) * 100) / 100;

    historySheet.appendRow([dateString, name, totalNetWorth]);
  }
}

function getPortfolioHistory(traderName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName("History");
  const data = historySheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  
  const entries = [];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === String(traderName).trim().toLowerCase()) {
      const rawDate = data[i][0];
      let label = "";
      let sortTime = i;

      if (rawDate instanceof Date) {
        label = Utilities.formatDate(rawDate, tz, "MMM dd");
        sortTime = rawDate.getTime();
      } else {
        label = String(rawDate).trim();
        const parsed = Date.parse(label + " " + new Date().getFullYear());
        if (!isNaN(parsed)) sortTime = parsed;
      }

      entries.push({
        label: label,
        value: Number(data[i][2]) || 0,
        sortTime: sortTime
      });
    }
  }

  // Sort chronologically so Nov 16 -> Nov 17 -> Nov 18 -> Nov 19 renders left-to-right!
  entries.sort((a, b) => a.sortTime - b.sortTime);

  return {
    labels: entries.map(e => e.label),
    values: entries.map(e => e.value)
  };
}
