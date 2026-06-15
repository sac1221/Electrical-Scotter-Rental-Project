function doGet(e) {
  if (e && e.parameter && e.parameter.admin === 'true') {
    const template = HtmlService.createTemplateFromFile('admin');
    try {
      template.initialDashboard = JSON.stringify(getActiveRentals());
      template.initialBookings = JSON.stringify(getAllBookings());
      template.initialInventory = JSON.stringify(getAllScooters());
    } catch(err) {
      template.initialDashboard = 'null';
      template.initialBookings = 'null';
      template.initialInventory = 'null';
    }
    return template.evaluate()
      .setTitle('EV Rentals Pro - Admin Panel')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else {
    return HtmlService.createHtmlOutputFromFile('guest')
      .setTitle('EV Rentals - Guest Booking')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

function getSettings() {
  const props = PropertiesService.getScriptProperties();
  return {
    spreadsheetId: props.getProperty('SPREADSHEET_ID') || '',
    reportEmail: props.getProperty('REPORT_EMAIL') || ''
  };
}

function saveSettings(settings) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SPREADSHEET_ID', settings.spreadsheetId);
  props.setProperty('REPORT_EMAIL', settings.reportEmail);
  return { success: true };
}

function getSheetApp() {
  const cache = CacheService.getScriptCache();
  let sid = cache.get('SPREADSHEET_ID');
  if (!sid) {
    const props = PropertiesService.getScriptProperties();
    sid = props.getProperty('SPREADSHEET_ID');
    if (sid) cache.put('SPREADSHEET_ID', sid, 21600); // cache for 6 hours
  }
  if (sid) {
    try { return SpreadsheetApp.openById(sid); } catch(e) { return null; }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function updateSheet() {
  const props = PropertiesService.getScriptProperties();
  let existingId = props.getProperty('SPREADSHEET_ID');
  let ss;
  
  try {
    if (existingId) {
      ss = SpreadsheetApp.openById(existingId);
    } else {
      ss = SpreadsheetApp.create('EV Rentals Database');
      existingId = ss.getId();
      props.setProperty('SPREADSHEET_ID', existingId);
    }
  } catch(e) {
    ss = SpreadsheetApp.create('EV Rentals Database');
    existingId = ss.getId();
    props.setProperty('SPREADSHEET_ID', existingId);
  }

  try {
    let vehiclesSheet = ss.getSheetByName('Vehicles');
    if (!vehiclesSheet) {
      vehiclesSheet = ss.insertSheet('Vehicles');
      vehiclesSheet.appendRow(['Scooter ID', 'Status']);
      vehiclesSheet.appendRow(['S-001', 'Available']);
      vehiclesSheet.appendRow(['S-002', 'Available']);
      vehiclesSheet.appendRow(['S-003', 'Available']);
      vehiclesSheet.appendRow(['S-004', 'Available']);
    }
    
    let bookingsSheet = ss.getSheetByName('Bookings');
    if (!bookingsSheet) {
      bookingsSheet = ss.insertSheet('Bookings');
      bookingsSheet.appendRow([
        'Booking ID', 'Customer Name', 'Phone', 'Alt Phone', 'Area', 
        'Booking Date', 'Booking Time', 'Expected Return Date', 'Days Rented', 
        'Scooter ID', 'Rent Amount', 'Security Deposit', 'Signature', 
        'Return Status', 'Late Fee', 'Deposit Refunded', 'Timestamp', 'Return Timestamp'
      ]);
    }
    
    let settingsSheet = ss.getSheetByName('Settings');
    if (!settingsSheet) {
      settingsSheet = ss.insertSheet('Settings');
      settingsSheet.appendRow(['Cottages']);
      const defaultCottages = ['Nadi Cottage', 'Shivapadam 1', 'Shivapadam 2', 'Shivapadam 3', 'Shivapadam 4', 'Nalanda Cottage', 'Brahmaputra', 'Adiyogi Alayam'];
      defaultCottages.forEach(c => settingsSheet.appendRow([c]));
    }
    
    let sheet1 = ss.getSheetByName('Sheet1');
    if (sheet1) ss.deleteSheet(sheet1);
    
    setupTriggers();
    return { success: true, message: "Database sheet updated/initialized successfully!", newId: ss.getId() };
  } catch(e) {
    return { success: false, message: "Error updating spreadsheet: " + e.toString() };
  }
}

function getAppConfig() {
  const ss = getSheetApp();
  if(!ss) return { cottages: [] };
  
  const settingsSheet = ss.getSheetByName('Settings');
  if(!settingsSheet) return { cottages: [] };
  
  const data = settingsSheet.getDataRange().getValues();
  const cottages = [];
  
  // Assuming column A is Cottages, starting from row 2
  for(let i=1; i<data.length; i++) {
    if(data[i][0]) cottages.push(data[i][0]);
  }
  
  return { cottages: cottages };
}

function getDashboardData() {
  const ss = getSheetApp();
  if(!ss) return { error: "No Spreadsheet connected." };
  
  const vehiclesSheet = ss.getSheetByName('Vehicles');
  const bookingsSheet = ss.getSheetByName('Bookings');
  
  let available = 0, rented = 0;
  if(vehiclesSheet) {
    const vData = vehiclesSheet.getDataRange().getValues();
    for(let i=1; i<vData.length; i++) {
      if(vData[i][1] === 'Available') available++;
      if(vData[i][1] === 'Rented') rented++;
    }
  }
  
  let todayProfit = 0; 
  let todayCashInHand = 0; 
  let depositsHeld = 0; 
  let returnsTodayCount = 0;
  let upcomingReturns = [];
  
  const todayIsoDate = new Date().toISOString().substring(0, 10);
  
  if(bookingsSheet) {
    const bData = bookingsSheet.getDataRange().getValues();
    for(let i=1; i<bData.length; i++) {
      const row = bData[i];
      const status = row[13];
      
      const isReturnToday = (status === 'Returned' && row[17] && row[17] instanceof Date && row[17].toISOString().substring(0,10) === todayIsoDate);
      const isExpectedToday = (status === 'Pending' && row[7] && row[7] instanceof Date && row[7].toISOString().substring(0,10) === todayIsoDate);

      // Bookings made today
      if(row[16] && row[16] instanceof Date && row[16].toISOString().substring(0, 10) === todayIsoDate) {
        let rentAmt = Number(row[10] || 0);
        let depAmt = Number(row[11] || 0);
        todayProfit += rentAmt;
        todayCashInHand += (rentAmt + depAmt);
      }
      
      // Returns processed today
      if(isReturnToday) {
        let lateAmt = Number(row[14] || 0);
        let refAmt = Number(row[15] || 0);
        todayProfit += lateAmt;
        todayCashInHand += lateAmt; 
        todayCashInHand -= refAmt;  
      }
      
      // Active Deposits (Allocated bikes)
      if(status === 'Pending' && row[9] !== 'Pending Allocation') {
        depositsHeld += Number(row[11] || 0);
      }
      
      // Upcoming Returns today
      if(isExpectedToday && row[9] !== 'Pending Allocation') {
        returnsTodayCount++;
        upcomingReturns.push({
          bookingId: row[0],
          name: row[1],
          phone: row[2],
          scooterId: row[9]
        });
      }
    }
  }
  
  return {
    available: available,
    rented: rented,
    todayProfit: todayProfit,
    todayCashInHand: todayCashInHand,
    depositsHeld: depositsHeld,
    returnsToday: returnsTodayCount,
    upcomingReturns: upcomingReturns
  };
}

function getAvailableScooters() {
  const ss = getSheetApp();
  const sheet = ss.getSheetByName('Vehicles');
  const data = sheet.getDataRange().getValues();
  let scooters = [];
  for(let i=1; i<data.length; i++) {
    if(data[i][1] === 'Available') scooters.push(data[i][0]);
  }
  return scooters;
}

function invalidateCache() {
  const cache = CacheService.getScriptCache();
  cache.removeAll(['EV_ACTIVE_RENTALS', 'EV_ALL_BOOKINGS', 'EV_ALL_SCOOTERS', 'EV_DASHBOARD']);
}

function getAllScooters() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('EV_ALL_SCOOTERS');
  if (cached) return JSON.parse(cached);

  const ss = getSheetApp();
  const sheet = ss.getSheetByName('Vehicles');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  let scooters = [];
  for(let i=1; i<data.length; i++) {
    scooters.push({ id: data[i][0], status: data[i][1] });
  }
  cache.put('EV_ALL_SCOOTERS', JSON.stringify(scooters), 21600);
  return scooters;
}

function getActiveRentals() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('EV_ACTIVE_RENTALS');
  if (cached) return JSON.parse(cached);

  const ss = getSheetApp();
  const sheet = ss.getSheetByName('Bookings');
  if (!sheet) return { rentals: [], pendingAllocations: [] };
  const data = sheet.getDataRange().getValues();
  let rentals = [];
  let pendingAllocations = [];
  
  for(let i=1; i<data.length; i++) {
    if(data[i][13] === 'Pending') {
      const obj = {
        rowIdx: i + 1,
        bookingId: data[i][0],
        name: data[i][1],
        phone: data[i][2],
        altPhone: data[i][3],
        area: data[i][4],
        bookDate: data[i][5],
        bookTime: data[i][6],
        expReturn: data[i][7],
        daysRented: data[i][8],
        scooterId: data[i][9],
        rentAmount: data[i][10],
        securityDeposit: data[i][11]
      };
      
      if (data[i][9] === 'Pending Allocation') {
        pendingAllocations.push(obj);
      } else {
        rentals.push(obj);
      }
    }
  }
  const result = { rentals, pendingAllocations };
  cache.put('EV_ACTIVE_RENTALS', JSON.stringify(result), 21600);
  return result;
}

function getAllBookings() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('EV_ALL_BOOKINGS');
  if (cached) return JSON.parse(cached);

  const ss = getSheetApp();
  const sheet = ss.getSheetByName('Bookings');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  let bookings = [];
  
  for(let i=1; i<data.length; i++) {
    bookings.push({
      rowIdx: i + 1,
      bookingId: data[i][0],
      name: data[i][1],
      phone: data[i][2],
      altPhone: data[i][3],
      area: data[i][4],
      bookDate: data[i][5],
      bookTime: data[i][6],
      expReturn: data[i][7],
      daysRented: data[i][8],
      scooterId: data[i][9],
      rentAmount: data[i][10],
      securityDeposit: data[i][11],
      status: data[i][13]
    });
  }
  cache.put('EV_ALL_BOOKINGS', JSON.stringify(bookings), 21600);
  return bookings;
}

function createBooking(data) {
  if (!data || !data.name || !data.phone) return { success: false, error: "Invalid data" };
  
  // Security fix: Enforce server-side rent calculation
  const days = parseInt(data.daysRented);
  if(isNaN(days) || days < 1) return { success: false, error: "Invalid days" };
  const enforceRentAmount = days * 300;
  const enforceSecurityDeposit = 2000;
  
  const ss = getSheetApp();
  const bookingsSheet = ss.getSheetByName('Bookings');
  const bookingId = 'EV' + new Date().getTime().toString().slice(-6);
  const bookTimestamp = new Date().toISOString();
  const status = 'Pending';
  
  const rowData = [
    bookingId,
    data.name.replace(/</g, "&lt;"), // Basic XSS sanitization
    data.phone,
    data.altPhone,
    data.area,
    data.bookDate,
    data.bookTime,
    data.expReturn,
    days,
    data.scooterId || 'Pending Allocation',
    enforceRentAmount,
    enforceSecurityDeposit,
    data.signature,
    status,
    0, 0, 
    bookTimestamp,
    ''
  ];
  
  bookingsSheet.appendRow(rowData);
  
  // If a specific scooter was selected (legacy support), mark it Rented. Otherwise, ignore.
  if (data.scooterId !== 'Pending Allocation') {
    updateScooterStatus(data.scooterId, 'Rented');
  }
  invalidateCache();
  return { success: true, bookingId: bookingId };
}

function processReturn(data) {
  const ss = getSheetApp();
  const bookingsSheet = ss.getSheetByName('Bookings');
  const bData = bookingsSheet.getDataRange().getValues();
  let rowIdx = -1;
  for(let i=1; i<bData.length; i++) {
    if(bData[i][0] === data.bookingId) {
      rowIdx = i + 1;
      break;
    }
  }
  if (rowIdx === -1) return { success: false, error: "Booking not found" };
  
  const returnTimestamp = new Date().toISOString();
  
  // Status, Late Fee, Deposit Refunded, Timestamp, Return Timestamp
  bookingsSheet.getRange(rowIdx, 14, 1, 5).setValues([['Returned', data.lateFee, data.depositRefunded, bData[rowIdx-1][16], returnTimestamp]]);
  
  let nextStatus = data.condition === 'Damaged' ? 'Maintenance' : 'Available';
  updateScooterStatus(data.scooterId, nextStatus);
  invalidateCache();
  
  return { success: true };
}

function allocateBike(data) {
  if (!data || !data.bookingId || !data.scooterId) return { success: false, error: "Invalid data" };

  const ss = getSheetApp();
  const bookingsSheet = ss.getSheetByName('Bookings');
  
  // Find booking
  const bData = bookingsSheet.getDataRange().getValues();
  let rowIdx = -1;
  for(let i=1; i<bData.length; i++) {
    if(bData[i][0] === data.bookingId) {
      rowIdx = i + 1;
      break;
    }
  }
  if (rowIdx === -1) return { success: false, error: "Booking not found" };
  
  // Update Scooter ID in bookings
  bookingsSheet.getRange(rowIdx, 10).setValue(data.scooterId);
  
  // Mark scooter as Rented using shared utility
  const statusUpdate = updateScooterStatus(data.scooterId, 'Rented');
  if (!statusUpdate.success) {
    return { success: false, error: "Scooter not found in Vehicles sheet" };
  }
  invalidateCache();
  
  return { success: true };
}


function updateScooterStatus(scooterId, newStatus) {
  const ss = getSheetApp();
  const sheet = ss.getSheetByName('Vehicles');
  const data = sheet.getDataRange().getValues();
  
  for(let i=1; i<data.length; i++) {
    if(data[i][0] === scooterId) {
      sheet.getRange(i+1, 2).setValue(newStatus);
      invalidateCache();
      return { success: true };
    }
  }
  return { success: false };
}

function setupTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for(let i=0; i<triggers.length; i++) {
    if(triggers[i].getHandlerFunction() === 'sendDailyReport') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('sendDailyReport')
    .timeBased()
    .everyDays(1)
    .atHour(21) // App Script runs this between 21:00 and 22:00 (approx 9:30 PM)
    .create();
}

function sendDailyReport() {
  const props = PropertiesService.getScriptProperties();
  const email = props.getProperty('REPORT_EMAIL');
  if(!email) return "No email configured.";

  const data = getDashboardData();
  if(data.error) return data.error;

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: auto; border: 1px solid #ccc; padding: 20px; border-radius: 10px;">
        <h2 style="color: #0EA5E9; text-align: center;">EV Rentals - Daily Report</h2>
        <p style="text-align: center; color: #555;">Date: <strong>${new Date().toDateString()}</strong></p>
        <hr>
        <table style="width: 100%; border-collapse: collapse; margin-top: 20px;">
            <tr><td style="padding: 10px; border-bottom: 1px solid #eee;"><strong>Today's Profit (Earnings)</strong><br><small>Rent + Late Fees</small></td><td style="padding: 10px; border-bottom: 1px solid #eee; text-align: right; color: #10B981; font-weight: bold; font-size: 1.2em;">₹${data.todayProfit}</td></tr>
            <tr><td style="padding: 10px; border-bottom: 1px solid #eee;"><strong>Net Cash In Hand</strong><br><small>Rent + Deposits In - Refunds</small></td><td style="padding: 10px; border-bottom: 1px solid #eee; text-align: right; font-weight: bold; font-size: 1.2em;">₹${data.todayCashInHand}</td></tr>
            <tr><td style="padding: 10px; border-bottom: 1px solid #eee;"><strong>Total Deposits Held</strong><br><small>For all active rentals</small></td><td style="padding: 10px; border-bottom: 1px solid #eee; text-align: right; font-weight: bold; font-size: 1.2em;">₹${data.depositsHeld}</td></tr>
            <tr><td style="padding: 10px; border-bottom: 1px solid #eee;"><strong>Available EVs</strong></td><td style="padding: 10px; border-bottom: 1px solid #eee; text-align: right; font-weight: bold;">${data.available}</td></tr>
            <tr><td style="padding: 10px;"><strong>Rented EVs</strong></td><td style="padding: 10px; text-align: right; font-weight: bold;">${data.rented}</td></tr>
        </table>
        <br>
        <p style="font-size: 12px; color: #888; text-align: center;">Generated automatically by EV Rentals System</p>
    </div>
  `;
  
  const blob = Utilities.newBlob(htmlBody, 'text/html', 'report.html');
  const pdf = blob.getAs('application/pdf');
  pdf.setName(`EV_Rental_Report_${new Date().toISOString().split('T')[0]}.pdf`);

  MailApp.sendEmail({
    to: email,
    subject: `EV Rental Daily Summary - ${new Date().toDateString()}`,
    htmlBody: htmlBody,
    attachments: [pdf]
  });
  
  return "Report sent to " + email;
}

