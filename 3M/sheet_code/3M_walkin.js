// ==================================================
// 3M CAR CARE CRM - Walk-in & Call Sheet Handler
// ==================================================

const SHEET_NAME = "WalkinCallLeads";

function getLocalISOString() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const hours = String(now.getHours()).padStart(2, '0');
  const minutes = String(now.getMinutes()).padStart(2, '0');
  const seconds = String(now.getSeconds()).padStart(2, '0');
  
  const offsetMinutes = now.getTimezoneOffset();
  const offsetHours = String(Math.floor(Math.abs(offsetMinutes) / 60)).padStart(2, '0');
  const offsetMins = String(Math.abs(offsetMinutes) % 60).padStart(2, '0');
  const offsetSign = offsetMinutes > 0 ? '-' : '+';
  
  return `${year}-${month}-${day}T${hours}:${minutes}:${seconds}${offsetSign}${offsetHours}:${offsetMins}`;
}

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // 3M Column Headers
    const headers = [
      "Date", "platform", "service", "Vehicle", "Full name", "Contact", "Source",
      "Service Required", "Nearest Location", "Vehicle Age (Month)", "Estimated Amount",
      "Customer Budget", "Interest Level", "From Hyderabad?", "Master Data Correct?",
      "Additional Notes", "Call 1 date", "Call 1 Status", "Call 2 date", "Call 2 Status",
      "Call 3 date", "Call 3 Status", "Confirmation Call Date", "Confirmation call",
      "Status", "Next Action"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim());

  function getColIdx(name) {
    // Direct match
    let idx = headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
    if (idx !== -1) return idx;
    // Partial match (for flexibility)
    idx = headers.findIndex(h => h.toLowerCase().includes(name.toLowerCase()));
    return idx;
  }

  try {
    const lead = JSON.parse(e.postData.contents);
    
    // Get phone from multiple possible field names
    let incomingPhone = '';
    if (lead.Contact) incomingPhone = String(lead.Contact).replace(/\D/g, '');
    else if (lead.phone) incomingPhone = String(lead.phone).replace(/\D/g, '');
    else if (lead.contact) incomingPhone = String(lead.contact).replace(/\D/g, '');
    else if (lead.contactNo) incomingPhone = String(lead.contactNo).replace(/\D/g, '');
    else if (lead['Contact No.']) incomingPhone = String(lead['Contact No.']).replace(/\D/g, '');
    
    if (!incomingPhone) {
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "error", 
        message: "No phone number provided" 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Find Contact column (try multiple variations)
    let contactIdx = getColIdx("Contact");
    if (contactIdx === -1) contactIdx = getColIdx("Contact No");
    if (contactIdx === -1) contactIdx = getColIdx("Phone");
    
    if (contactIdx === -1) {
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "error", 
        message: "Contact column not found" 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      const rowPhone = String(data[i][contactIdx]).replace(/\D/g, '');
      if (rowPhone === incomingPhone && rowPhone !== "") {
        targetRow = i + 1;
        break;
      }
    }

    const isNew = targetRow === -1;
    if (isNew) {
      targetRow = sheet.getLastRow() + 1;
    }

    function setValue(colName, value) {
      if (value === undefined || value === null) return;
      // Allow empty strings for clearing fields
      const idx = getColIdx(colName);
      if (idx !== -1 && targetRow !== -1) {
        sheet.getRange(targetRow, idx + 1).setValue(value);
      }
    }

    const sd = lead.structuredData || {};

    // ----- BASIC FIELDS (only if new) -----
    if (isNew) {
      setValue("Date", getLocalISOString());
      setValue("platform", lead.platform || (lead.leadType === 'walkin' ? 'Walk-in' : 'Call'));
      setValue("service", lead.service || '');
      setValue("Vehicle", lead.Vehicle || lead.vehicle || '');
      setValue("Full name", lead.name || lead['Full name'] || lead.fullName || '');
      setValue("Contact", lead.Contact || lead.phone || lead.contact || lead.contactNo || lead['Contact No.'] || '');
      setValue("Source", lead.Source || lead.source || (lead.leadType === 'walkin' ? 'Walk-in Customer' : 'Phone Call'));
    }

    // ----- STRUCTURED DATA FIELDS (always update) -----
    setValue("Service Required", sd.services ? (Array.isArray(sd.services) ? sd.services.join(', ') : sd.services) : '');
    setValue("Nearest Location", sd.nearestLocation || lead.nearestLocation || '');
    setValue("Vehicle Age (Month)", sd.vehicleAgeMonths || lead.vehicleAgeMonths || '');
    setValue("Estimated Amount", sd.estimatedAmount || lead.estimatedAmount || '');
    setValue("Customer Budget", sd.budget || lead.budget || '');
    setValue("Interest Level", sd.interestLevel || lead.interestLevel || '');
    setValue("From Hyderabad?", sd.fromHyderabad || lead.fromHyderabad || lead.location || '');
    setValue("Master Data Correct?", sd.masterDataCorrect || lead.masterDataCorrect || '');
    setValue("Additional Notes", sd.additionalNotes || lead.notes || lead.additionalNotes || '');

    // ----- CALL TRACKING FIELDS (always update) -----
    setValue("Call 1 date", lead.call1Date || '');
    setValue("Call 1 Status", lead.call1Status || 'Pending');
    setValue("Call 2 date", lead.call2Date || '');
    setValue("Call 2 Status", lead.call2Status || 'Pending');
    setValue("Call 3 date", lead.call3Date || '');
    setValue("Call 3 Status", lead.call3Status || 'Pending');
    setValue("Confirmation Call Date", lead.confirmDate || '');
    
    // ----- CONFIRMATION & STATUS -----
    let confirmationStatus = "Pending";
    if (lead.converted) confirmationStatus = "Completed";
    else if (lead.confirmDate) confirmationStatus = "Scheduled";
    setValue("Confirmation call", confirmationStatus);
    
    let rowStatus = "Active";
    if (lead.lost) rowStatus = "Lost";
    if (lead.converted) rowStatus = "Converted";
    setValue("Status", rowStatus);
    setValue("Next Action", lead.nextAction || getNextActionFromLead(lead));

    // ----- RESPONSE -----
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "success", 
      row: targetRow,
      isNew: isNew,
      message: isNew ? "Lead created successfully" : "Lead updated successfully"
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error('Error in doPost:', error);
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "error", 
      message: error.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function getNextActionFromLead(lead) {
  if (lead.lost) return 'Lost';
  if (lead.converted) return 'Converted';
  if (lead.call1Status === 'Pending') return 'Call 1';
  if (lead.call2Status === 'Pending') return 'Call 2';
  if (lead.call3Status === 'Pending') return 'Call 3';
  if (!lead.converted && !lead.confirmDate) return 'Confirmation';
  return 'Completed';
}

// ==================================================
// DOGET - Fetch all leads from the sheet
// ==================================================

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "success", 
      data: [],
      totalCount: 0,
      message: "Sheet not found"
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "success", 
      data: [],
      totalCount: 0,
      message: "No data found"
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  const headers = data[0].map(h => h.toString().trim());
  
  // Parse query parameters
  const params = e?.parameter || {};
  const startDate = params.startDate;
  const endDate = params.endDate;
  const offset = parseInt(params.offset) || 0;
  const limit = parseInt(params.limit) || 50;
  const search = params.search ? params.search.toLowerCase() : '';
  const platform = params.platform;
  
  let leads = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const lead = {};
    headers.forEach((header, idx) => {
      let value = row[idx];
      if (value instanceof Date) {
        value = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      lead[header] = value;
    });
    
    // Filter by platform
    if (platform && lead.platform !== platform) continue;
    
    // Filter by date range
    if (startDate && endDate && lead.Date) {
      const leadDate = new Date(lead.Date);
      const start = new Date(startDate);
      const end = new Date(endDate);
      end.setHours(23, 59, 59);
      if (leadDate < start || leadDate > end) continue;
    }
    
    // Filter by search term
    if (search) {
      const nameMatch = (lead['Full name'] || '').toLowerCase().includes(search);
      const phoneMatch = (lead['Contact'] || '').toLowerCase().includes(search);
      const vehicleMatch = (lead['Vehicle'] || '').toLowerCase().includes(search);
      const serviceMatch = (lead['service'] || '').toLowerCase().includes(search);
      if (!nameMatch && !phoneMatch && !vehicleMatch && !serviceMatch) continue;
    }
    
    leads.push(lead);
  }
  
  // Sort by Date descending (newest first)
  leads.sort((a, b) => {
    const dateA = a.Date || '';
    const dateB = b.Date || '';
    if (dateA > dateB) return -1;
    if (dateA < dateB) return 1;
    return 0;
  });
  
  const totalCount = leads.length;
  const paginatedLeads = leads.slice(offset, offset + limit);
  
  return ContentService.createTextOutput(JSON.stringify({ 
    status: "success", 
    data: paginatedLeads,
    totalCount: totalCount,
    offset: offset,
    limit: limit,
    hasMore: offset + limit < totalCount
  })).setMimeType(ContentService.MimeType.JSON);
}

// ==================================================
// OPTIONAL: Helper to get column headers
// ==================================================

function getHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return [];
  const headers = sheet.getDataRange().getValues()[0];
  return headers.map(h => h.toString().trim());
}

// ==================================================
// OPTIONAL: Test function to verify column mapping
// ==================================================

function testColumnMapping() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log('Sheet not found. Creating...');
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = [
      "Date", "platform", "service", "Vehicle", "Full name", "Contact", "Source",
      "Service Required", "Nearest Location", "Vehicle Age (Month)", "Estimated Amount",
      "Customer Budget", "Interest Level", "From Hyderabad?", "Master Data Correct?",
      "Additional Notes", "Call 1 date", "Call 1 Status", "Call 2 date", "Call 2 Status",
      "Call 3 date", "Call 3 Status", "Confirmation Call Date", "Confirmation call",
      "Status", "Next Action"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  
  const headers = sheet.getDataRange().getValues()[0].map(h => h.toString().trim());
  
  const testColumns = [
    'Date', 'platform', 'service', 'Vehicle', 'Full name', 'Contact', 'Source',
    'Service Required', 'Nearest Location', 'Vehicle Age (Month)', 'Estimated Amount',
    'Customer Budget', 'Interest Level', 'From Hyderabad?', 'Master Data Correct?',
    'Additional Notes', 'Call 1 date', 'Call 1 Status', 'Call 2 date', 'Call 2 Status',
    'Call 3 date', 'Call 3 Status', 'Confirmation Call Date', 'Confirmation call',
    'Status', 'Next Action'
  ];
  
  const results = {};
  testColumns.forEach(col => {
    const idx = headers.findIndex(h => h.toLowerCase() === col.toLowerCase());
    results[col] = idx !== -1 ? `Found at column ${idx + 1}` : 'NOT FOUND';
  });
  
  Logger.log(results);
  return results;
}

// ==================================================
// OPTIONAL: Function to update a specific lead by phone
// ==================================================

function updateLeadByPhone(phone, updates) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { status: 'error', message: 'Sheet not found' };
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim());
  
  const contactIdx = headers.findIndex(h => h.toLowerCase() === 'contact');
  if (contactIdx === -1) return { status: 'error', message: 'Contact column not found' };
  
  const cleanPhone = String(phone).replace(/\D/g, '');
  let targetRow = -1;
  
  for (let i = 1; i < data.length; i++) {
    const rowPhone = String(data[i][contactIdx]).replace(/\D/g, '');
    if (rowPhone === cleanPhone && rowPhone !== '') {
      targetRow = i + 1;
      break;
    }
  }
  
  if (targetRow === -1) return { status: 'error', message: 'Lead not found' };
  
  function setValue(colName, value) {
    if (value === undefined || value === null) return;
    const idx = headers.findIndex(h => h.toLowerCase() === colName.toLowerCase());
    if (idx !== -1) {
      sheet.getRange(targetRow, idx + 1).setValue(value);
    }
  }
  
  // Apply updates
  Object.keys(updates).forEach(key => {
    setValue(key, updates[key]);
  });
  
  return { status: 'success', message: 'Lead updated successfully', row: targetRow };
}