/**
 * FeedHope - Google Sheets Backend (Google Apps Script)
 * Copy and paste this code into Extensions -> Apps Script in your Google Sheet,
 * then Deploy as Web App (Execute as: Me, Access: Anyone).
 */

function doGet(e) {
  try {
    let ss = null;
    try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch(err) {}
    
    const params = (e && e.parameter) ? e.parameter : {};

    // 1. DEDUPLICATION CHECK (Prevents Google Apps Script 302 redirect double execution)
    if (params.reqId) {
      try {
        const cache = CacheService.getScriptCache();
        if (cache.get(params.reqId)) {
          return returnCurrentData(ss);
        }
        cache.put(params.reqId, '1', 60);
      } catch(cacheErr) {}
    }

    // 2. ACTION HANDLING
    if (params.action === 'deleteDonation' || params.action === 'delete') {
      deleteDonationFromParams(ss, params);
    } else if (params.action === 'deleteExpense') {
      deleteExpenseFromParams(ss, params);
    } else if (params.action === 'addExpense') {
      addExpenseFromParams(ss, params);
    } else if (params.action === 'clearAll') {
      clearAllSheetsData(ss);
    } else if (params.action === 'addDonation' || (params.amount && !params.action)) {
      addDonationFromParams(ss, params);
    } else if (params.name && !params.action) {
      addDonationFromParams(ss, params);
    }

    return returnCurrentData(ss);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    let ss = null;
    try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch(err) {}
    let data = {};

    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    // 1. DEDUPLICATION CHECK
    if (data.reqId) {
      try {
        const cache = CacheService.getScriptCache();
        if (cache.get(data.reqId)) {
          return returnCurrentData(ss);
        }
        cache.put(data.reqId, '1', 60);
      } catch(cacheErr) {}
    }

    // 2. ACTION HANDLING
    if (data.action === 'deleteDonation' || data.action === 'delete') {
      deleteDonationFromParams(ss, data);
    } else if (data.action === 'deleteExpense') {
      deleteExpenseFromParams(ss, data);
    } else if (data.action === 'addExpense') {
      addExpenseFromParams(ss, data);
    } else if (data.action === 'clearAll') {
      clearAllSheetsData(ss);
    } else if (data.amount || data.name) {
      addDonationFromParams(ss, data);
    }

    return returnCurrentData(ss);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function addDonationFromParams(ss, params) {
  let ssObj = ss;
  if (!ssObj) {
    try { ssObj = SpreadsheetApp.getActiveSpreadsheet(); } catch(e) {}
  }
  if (!ssObj) return;

  const donationsSheet = ssObj.getSheetByName('Donations') || ssObj.insertSheet('Donations');
  const name = params.name || 'Anonymous';
  const amount = Number(params.amount) || 0;
  const msg = params.msg || '❤️ Supporter';
  const time = params.time || new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  const avatar = params.avatar || (name !== 'Anonymous' ? name.charAt(0).toUpperCase() : '?');

  donationsSheet.appendRow([name, amount, msg, time, avatar]);
}

function deleteDonationFromParams(ss, params) {
  let ssObj = ss;
  if (!ssObj) {
    try { ssObj = SpreadsheetApp.getActiveSpreadsheet(); } catch(e) {}
  }
  if (!ssObj) return;

  const sheet = ssObj.getSheetByName('Donations');
  if (!sheet) return;

  const index = parseInt(params.index, 10);
  const targetName = params.name ? String(params.name).trim().toLowerCase() : '';

  const rows = sheet.getDataRange().getValues();
  const startRow = (rows.length > 0 && isNaN(Number(rows[0][1]))) ? 1 : 0;
  const dataRowsCount = rows.length - startRow;

  if (!isNaN(index) && index >= 0 && index < dataRowsCount) {
    const sheetRowToDelete = rows.length - index;
    sheet.deleteRow(sheetRowToDelete);
    return;
  }

  if (targetName) {
    for (let i = rows.length - 1; i >= startRow; i--) {
      const rowName = String(rows[i][0] || '').trim().toLowerCase();
      if (rowName === targetName) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
  }
}

function addExpenseFromParams(ss, params) {
  let ssObj = ss;
  if (!ssObj) {
    try { ssObj = SpreadsheetApp.getActiveSpreadsheet(); } catch(e) {}
  }
  if (!ssObj) return;

  const sheet = ssObj.getSheetByName('Expenses') || ssObj.insertSheet('Expenses');
  const purpose = params.purpose || params.title || 'Expense';
  const amount = Number(params.amount) || 0;
  const receipt = params.receipt || '#';
  const date = params.date || new Date().toISOString().split('T')[0];
  const status = params.status || 'Verified';

  sheet.appendRow([purpose, amount, receipt, date, status]);
}

function deleteExpenseFromParams(ss, params) {
  let ssObj = ss;
  if (!ssObj) {
    try { ssObj = SpreadsheetApp.getActiveSpreadsheet(); } catch(e) {}
  }
  if (!ssObj) return;

  const sheet = ssObj.getSheetByName('Expenses');
  if (!sheet) return;

  const index = parseInt(params.index, 10);
  const rows = sheet.getDataRange().getValues();
  const startRow = (rows.length > 0 && isNaN(Number(rows[0][1]))) ? 1 : 0;
  const dataRowsCount = rows.length - startRow;

  if (!isNaN(index) && index >= 0 && index < dataRowsCount) {
    const sheetRowToDelete = rows.length - index;
    sheet.deleteRow(sheetRowToDelete);
  }
}

function clearAllSheetsData(ss) {
  let ssObj = ss;
  if (!ssObj) {
    try { ssObj = SpreadsheetApp.getActiveSpreadsheet(); } catch(e) {}
  }
  if (!ssObj) return;

  const donationsSheet = ssObj.getSheetByName('Donations');
  if (donationsSheet) donationsSheet.clearContents();

  const expensesSheet = ssObj.getSheetByName('Expenses');
  if (expensesSheet) expensesSheet.clearContents();
}

function returnCurrentData(ss) {
  let ssObj = ss;
  if (!ssObj) {
    try { ssObj = SpreadsheetApp.getActiveSpreadsheet(); } catch(e) {}
  }

  let goal = 500000, meals = 0;
  
  if (ssObj) {
    const configSheet = ssObj.getSheetByName('Config');
    if (configSheet) {
      const data = configSheet.getDataRange().getValues();
      data.forEach(row => {
        const key = String(row[0]).toLowerCase().trim();
        const val = Number(row[1]) || 0;
        if (key.includes('goal')) goal = val;
        if (key.includes('meal')) meals = val;
      });
    }
  }

  const donations = [];
  if (ssObj) {
    const donationsSheet = ssObj.getSheetByName('Donations');
    if (donationsSheet) {
      const rows = donationsSheet.getDataRange().getValues();
      const startRow = (rows.length > 0 && isNaN(Number(rows[0][1]))) ? 1 : 0;
      for (let i = startRow; i < rows.length; i++) {
        const row = rows[i];
        if (row[0] || row[1]) {
          donations.push({
            name: row[0] || 'Anonymous',
            amount: Number(row[1]) || 0,
            msg: row[2] || 'Thank you for supporting!',
            time: row[3] || 'Just now',
            avatar: row[4] || (row[0] ? String(row[0]).charAt(0).toUpperCase() : '?')
          });
        }
      }
    }
  }

  donations.reverse();

  const expenses = [];
  if (ssObj) {
    const expensesSheet = ssObj.getSheetByName('Expenses');
    if (expensesSheet) {
      const rows = expensesSheet.getDataRange().getValues();
      const startRow = (rows.length > 0 && isNaN(Number(rows[0][1]))) ? 1 : 0;
      for (let i = startRow; i < rows.length; i++) {
        const row = rows[i];
        if (row[0] || row[1]) {
          expenses.push({
            purpose: row[0] || 'Expense',
            amount: Number(row[1]) || 0,
            receipt: row[2] || '#',
            date: row[3] || '2025-02-15',
            status: row[4] || 'Verified'
          });
        }
      }
    }
  }
  expenses.reverse();

  const totalRaised = donations.reduce((sum, d) => sum + Number(d.amount), 0);
  const donors = donations.length;

  const response = {
    status: 'success',
    totalRaised: totalRaised,
    goal: goal,
    donors: donors,
    mealsServed: meals > 0 ? meals : Math.floor(totalRaised / 40),
    donations: donations,
    expenses: expenses
  };

  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}
