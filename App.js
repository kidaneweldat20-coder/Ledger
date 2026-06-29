// Configuration
const SHEET_NAME = "Ledger"; 
const USERS_SHEET = "Users";
const AUDIT_SHEET = "Audit_Log"; // <-- 1. Add this configuration

// Helper: Response method
function respond(responseObject) {
  return ContentService.createTextOutput(JSON.stringify(responseObject))
    .setMimeType(ContentService.MimeType.JSON);
}

// Helper: Log Action (Add this anywhere before doPost)
function logAction(user, actionType, details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let auditSheet = ss.getSheetByName(AUDIT_SHEET);
  
  // If the sheet doesn't exist yet, create it and add headers automatically
  if (!auditSheet) {
    auditSheet = ss.insertSheet(AUDIT_SHEET);
    auditSheet.appendRow(["Timestamp", "User", "Action", "Details"]);
    auditSheet.getRange("A1:D1").setFontWeight("bold"); // Make headers bold
  }
  
  const timestamp = new Date(); // Logs the exact server time
  auditSheet.appendRow([timestamp, user, actionType, details]);
}

// GET: Load Ledger Data
function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) return respond([]);

  const data = sheet.getDataRange().getValues();
  const result = [];

  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    if (!row[0]) continue; 
    
    let formattedDate = row[1];
    if (row[1] instanceof Date) {
      formattedDate = Utilities.formatDate(row[1], Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    
    result.push({
      id: row[0],
      date: formattedDate,
      group: row[2],
      category: row[3],
      amount: parseFloat(row[4]) || 0,
      type: row[5]
    });
  }
  return respond(result);
}

// POST: All logic (Login, Security, Password Recovery, and Ledger)
function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledgerSheet = ss.getSheetByName(SHEET_NAME);
  const userSheet = ss.getSheetByName(USERS_SHEET);
  
  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return respond({ status: "error", message: "Invalid JSON" });
  }

  const action = payload.action;

  // --- 1. AUTHENTICATION & SECURITY ---
  if (action === 'LOGIN') {
    const users = userSheet.getDataRange().getValues();
    for (let i = 1; i < users.length; i++) {
      if (users[i][0] == payload.username && users[i][1] == payload.password) {
        return respond({ status: "success" });
      }
    }
    return respond({ status: "error", message: "Invalid credentials" });
  }

  if (action === 'CHANGE_PWD') {
    const users = userSheet.getDataRange().getValues();
    for (let i = 1; i < users.length; i++) {
      if (users[i][0] == payload.username && users[i][1] == payload.oldPassword) {
        userSheet.getRange(i + 1, 2).setValue(payload.newPassword);
        return respond({ status: "success", message: "Password updated successfully!" });
      }
    }
    return respond({ status: "error", message: "Old password incorrect" });
  }

  // --- RECOVERY LOGIC ---
  if (action === 'RECOVER_PASSWORD') {
    // CHANGE THIS KEY to something only you know!
    const MASTER_RECOVERY_KEY = "MY_SUPER_SECRET_KEY_2026"; 
    
    if (payload.recoveryCode === MASTER_RECOVERY_KEY) {
      // Resets Admin password (Row 2, Column 2)
      userSheet.getRange(2, 2).setValue("123456");
      return respond({ status: "success", message: "Password reset to '123456'. Please log in and change it immediately." });
    } else {
      return respond({ status: "error", message: "Invalid Recovery Key." });
    }
  }
  // --- 2. LEDGER OPERATIONS ---
  if (action === "DELETE" || action === "CLEAR_ALL") {
    if (!payload.secretToken || !payload.username) {
      return respond({ status: "error", message: "Unauthorized: Missing credentials" });
    }

    // Actually verify the password against the database
    const users = userSheet.getDataRange().getValues();
    let isValidUser = false;
    for (let i = 1; i < users.length; i++) {
      if (users[i][0] == payload.username && users[i][1] == payload.secretToken) {
        isValidUser = true;
        break;
      }
    }

    if (!isValidUser) {
      return respond({ status: "error", message: "Unauthorized: Incorrect password" });
    }
  }


  if (action === 'INSERT') {
    ledgerSheet.appendRow([payload.id, payload.date, payload.group, payload.category, payload.amount, payload.type]);
    return respond({ status: "success" });
  }

  if (action === 'UPDATE') {
    const values = ledgerSheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (values[i][0].toString() === payload.id.toString()) {
        ledgerSheet.getRange(i + 1, 2, 1, 5).setValues([[payload.date, payload.group, payload.category, payload.amount, payload.type]]);
        return respond({ status: "success" });
      }
    }
  }

    if (action === 'DELETE') {
    const values = ledgerSheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (values[i][0].toString() === payload.id.toString()) {
        // Grab the row's data BEFORE deleting it so we can log what was removed
        const date = values[i][1] instanceof Date ? Utilities.formatDate(values[i][1], Session.getScriptTimeZone(), "yyyy-MM-dd") : values[i][1];
        const group = values[i][2];
        const category = values[i][3];
        const amount = values[i][4];
        const type = values[i][5];
        
        const deletedDetails = `Date: ${date} | Group: ${group} | Cat: ${category} | Amount: £${amount} | Type: ${type}`;
        
        // Delete the row
        ledgerSheet.deleteRow(i + 1);
        
        // Save to Audit Log
        logAction(payload.username, "DELETE_RECORD", deletedDetails);
        
        return respond({ status: "success" });
      }
    }
  }

  if (action === 'CLEAR_ALL') {
    if (ledgerSheet.getLastRow() > 1) {
      // Delete all data rows
      ledgerSheet.deleteRows(2, ledgerSheet.getLastRow() - 1);
      
      // Save to Audit Log
      logAction(payload.username, "CLEAR_ALL_DATA", "User wiped the entire ledger.");
    }
    return respond({ status: "success" });
  }
  }
