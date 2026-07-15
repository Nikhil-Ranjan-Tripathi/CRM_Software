// ==================================================
// 3M CAR CARE CRM - Social Media Sheet Handler
// WITH PAGINATION SUPPORT
// ==================================================

const SHEET_NAME = "Sheet1";
const DEFAULT_LIMIT = 100;

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = [
      "Date", "ad_id", "platform", "service", "Vehicle", "Full name", "Contact", "Source",
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
    let idx = headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
    if (idx !== -1) return idx;
    const cleanName = name.replace(/[^a-zA-Z0-9 ]/g, '').trim().toLowerCase();
    idx = headers.findIndex(h => h.toLowerCase().replace(/[^a-zA-Z0-9 ]/g, '').trim() === cleanName);
    return idx;
  }

  try {
    let lead;
    if (typeof e.postData.contents === 'string') {
      lead = JSON.parse(e.postData.contents);
    } else {
      lead = e.postData.contents;
    }
    
    let incomingPhone = lead.Contact || lead.phone || lead.contact || '';
    incomingPhone = String(incomingPhone).replace(/\D/g, '');
    
    if (!incomingPhone) {
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "error", 
        message: "No phone number provided" 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const contactIdx = getColIdx("Contact");
    if (contactIdx === -1) {
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "error", 
        message: "Contact column not found" 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      const rowPhone = String(data[i][contactIdx] || '').replace(/\D/g, '');
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
      const strValue = String(value).trim();
      if (strValue === '') return;
      const idx = getColIdx(colName);
      if (idx !== -1 && targetRow !== -1) {
        try {
          sheet.getRange(targetRow, idx + 1).setValue(value);
        } catch (e) {
          console.error('Error setting', colName, ':', e);
        }
      }
    }

    const sd = lead.structuredData || {};

    if (isNew) {
      setValue("Date", lead.Date || lead.date || new Date().toISOString().split('T')[0]);
      setValue("ad_id", lead.ad_id || '');
      setValue("platform", lead.platform || '');
      setValue("service", lead.service || '');
      setValue("Vehicle", lead.Vehicle || lead.vehicle || '');
      setValue("Full name", lead.name || lead['Full name'] || lead['Full Name'] || '');
      setValue("Contact", lead.Contact || lead.phone || lead.contact || '');
      setValue("Source", lead.Source || lead.source || '');
    }

    // Structured data
    const serviceRequired = sd.services ? (Array.isArray(sd.services) ? sd.services.join(', ') : sd.services) : (lead['Service Required'] || '');
    setValue("Service Required", serviceRequired);
    setValue("Nearest Location", sd.nearestLocation || lead['Nearest Location'] || '');
    setValue("Vehicle Age (Month)", sd.vehicleAgeMonths || lead['Vehicle Age (Month)'] || '');
    setValue("Estimated Amount", sd.estimatedAmount || lead['Estimated Amount'] || '');
    setValue("Customer Budget", sd.budget || lead['Customer Budget'] || '');
    setValue("Interest Level", sd.interestLevel || lead['Interest Level'] || '');
    setValue("From Hyderabad?", sd.fromHyderabad || lead['From Hyderabad?'] || '');
    setValue("Master Data Correct?", sd.masterDataCorrect || lead['Master Data Correct?'] || '');
    setValue("Additional Notes", sd.additionalNotes || lead.notes || lead['Additional Notes'] || '');

    // Call tracking
    setValue("Call 1 date", lead.call1Date || lead['Call 1 date'] || '');
    setValue("Call 1 Status", lead.call1Status || lead['Call 1 Status'] || 'Pending');
    setValue("Call 2 date", lead.call2Date || lead['Call 2 date'] || '');
    setValue("Call 2 Status", lead.call2Status || lead['Call 2 Status'] || 'Pending');
    setValue("Call 3 date", lead.call3Date || lead['Call 3 date'] || '');
    setValue("Call 3 Status", lead.call3Status || lead['Call 3 Status'] || 'Pending');
    setValue("Confirmation Call Date", lead.confirmDate || lead['Confirmation Call Date'] || '');
    
    let confirmationStatus = "Pending";
    if (lead.converted || lead.Status === 'Converted') confirmationStatus = "Completed";
    else if (lead.confirmDate || lead['Confirmation Call Date']) confirmationStatus = "Scheduled";
    setValue("Confirmation call", confirmationStatus);
    
    let rowStatus = "Active";
    if (lead.lost || lead.Status === 'Lost') rowStatus = "Lost";
    else if (lead.converted || lead.Status === 'Converted') rowStatus = "Converted";
    setValue("Status", rowStatus);
    
    let nextAction = lead.nextAction || '';
    if (!nextAction) {
      if (rowStatus === 'Lost') nextAction = 'Lost';
      else if (rowStatus === 'Converted') nextAction = 'Converted';
      else if (lead.call1Status === 'Pending' || lead['Call 1 Status'] === 'Pending') nextAction = 'Call 1';
      else if (lead.call2Status === 'Pending' || lead['Call 2 Status'] === 'Pending') nextAction = 'Call 2';
      else if (lead.call3Status === 'Pending' || lead['Call 3 Status'] === 'Pending') nextAction = 'Call 3';
      else if (confirmationStatus !== 'Completed') nextAction = 'Confirmation';
      else nextAction = 'Completed';
    }
    setValue("Next Action", nextAction);

    return ContentService.createTextOutput(JSON.stringify({ 
      status: "success", 
      row: targetRow,
      isNew: isNew,
      message: isNew ? "Lead created" : "Lead updated"
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error('Error:', error);
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "error", 
      message: error.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "success", 
      data: [],
      totalCount: 0,
      offset: 0,
      limit: 0,
      hasMore: false
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "success", 
      data: [],
      totalCount: 0,
      offset: 0,
      limit: 0,
      hasMore: false
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  const headers = data[0].map(h => h.toString().trim());
  const params = e.parameter || {};
  
  // Pagination parameters
  const offset = parseInt(params.offset) || 0;
  const limit = parseInt(params.limit) || 100;
  const search = params.search ? params.search.toLowerCase() : '';
  const startDate = params.startDate;
  const endDate = params.endDate;
  
  // First, collect ALL leads with proper sorting by date (newest first)
  let allLeads = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const lead = {};
    headers.forEach((header, idx) => {
      lead[header] = row[idx] !== undefined ? row[idx] : '';
    });
    
    // Skip empty rows
    if (!lead['Full name'] && !lead['Full Name'] && !lead['Contact']) continue;
    
    // Apply date filter if provided
    if (startDate && endDate && lead['Date']) {
      try {
        const leadDate = new Date(lead['Date']);
        const start = new Date(startDate);
        const end = new Date(endDate);
        end.setHours(23, 59, 59);
        if (leadDate < start || leadDate > end) continue;
      } catch (e) {}
    }
    
    // Apply search filter if provided
    if (search) {
      const nameMatch = (lead['Full name'] || lead['Full Name'] || '').toLowerCase().includes(search);
      const phoneMatch = (lead['Contact'] || '').toLowerCase().includes(search);
      if (!nameMatch && !phoneMatch) continue;
    }
    
    allLeads.push(lead);
  }
  
  // Sort by date (newest first)
  allLeads.sort((a, b) => {
    const dateA = a['Date'] || a['created_time'] || '';
    const dateB = b['Date'] || b['created_time'] || '';
    if (dateA > dateB) return -1;
    if (dateA < dateB) return 1;
    return 0;
  });
  
  const totalCount = allLeads.length;
  
  // Apply pagination
  const paginatedLeads = allLeads.slice(offset, offset + limit);
  const hasMore = (offset + limit) < totalCount;
  
  return ContentService.createTextOutput(JSON.stringify({ 
    status: "success", 
    data: paginatedLeads,
    totalCount: totalCount,
    offset: offset,
    limit: limit,
    hasMore: hasMore,
    nextOffset: hasMore ? offset + limit : -1
  })).setMimeType(ContentService.MimeType.JSON);
}