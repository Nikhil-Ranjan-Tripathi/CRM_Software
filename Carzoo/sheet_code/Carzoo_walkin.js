// ==================================================
// 3M CAR CARE CRM - Walk-in & Call Sheet Handler
// ==================================================

const SHEET_NAME = "Walkin_Call";

// ----- Helper: UTC ISO string with milliseconds (e.g., 2026-06-25T18:33:39.940Z) -----
function getUTCISOString() {
  return new Date().toISOString();
}

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = [
      "created_time", "Platform", "Full name", "Contact No.", "Vehicle", "Service",
      "Service Required", "Vehicle Age (Month)", "Estimated Amount", "Customer Budget",
      "Interest Level", "Location", "Master Data Correct?", "Additional Notes",
      "Call 1 Date", "Call 1 Status", "Call 2 Date", "Call 2 Status",
      "Call 3 Date", "Call 3 Status", "Confirmation Call Date", "Confirmation call",
      "Status", "Next Action"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim());

  function getColIdx(name) {
    // exact match, then partial match
    let idx = headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
    if (idx !== -1) return idx;
    idx = headers.findIndex(h => h.toLowerCase().includes(name.toLowerCase()));
    return idx;
  }

  try {
    const lead = JSON.parse(e.postData.contents);
    
    // Extract phone numbers
    const oldPhone = String(lead.oldPhone || lead.originalPhone || '').replace(/\D/g, '');
    const newPhone = String(lead['Contact No.'] || lead.Contact || lead.phone || lead.contact || lead.contactNo || '').replace(/\D/g, '');
    
    Logger.log('oldPhone: ' + oldPhone + ', newPhone: ' + newPhone);
    Logger.log('lead.Vehicle: ' + lead.Vehicle + ', lead.vehicle: ' + lead.vehicle);
    
    if (!newPhone && !oldPhone) {
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "error", 
        message: "No phone number provided" 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Find Contact No. column
    let contactIdx = getColIdx("Contact No.");
    if (contactIdx === -1) contactIdx = getColIdx("Contact");
    if (contactIdx === -1) {
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "error", 
        message: "Contact No. column not found. Headers: " + headers.join(', ')
      })).setMimeType(ContentService.MimeType.JSON);
    }

    let targetRow = -1;
    const lookupPhone = oldPhone || newPhone;
    for (let i = 1; i < data.length; i++) {
      const rowPhone = String(data[i][contactIdx]).replace(/\D/g, '');
      if (rowPhone === lookupPhone && rowPhone !== "") {
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
      if (value === '') return; // skip empty strings
      const idx = getColIdx(colName);
      if (idx !== -1 && targetRow !== -1) {
        sheet.getRange(targetRow, idx + 1).setValue(value);
      }
    }

    const sd = lead.structuredData || {};

    // ----- Core fields -----
    const vehicleValue = lead.Vehicle || lead.vehicle || '';
    const fullNameValue = lead.name || lead['Full name'] || '';
    const serviceValue = lead.service || '';
    const platformValue = lead.platform || (lead.leadType === 'walkin' ? 'Walk-in' : 'Call');

    if (isNew) {
      // --- New record: set created_time with UTC ISO string ---
      setValue("created_time", getUTCISOString());
      // Also set Date if exists (for compatibility)
      if (getColIdx("Date") !== -1) setValue("Date", getUTCISOString());
      
      setValue("Platform", platformValue);
      setValue("Full name", fullNameValue);
      setValue("Contact No.", newPhone);
      setValue("Vehicle", vehicleValue);
      setValue("Service", serviceValue);
      // Also set "Contact" if exists (backward compatibility)
      if (getColIdx("Contact") !== -1) setValue("Contact", newPhone);
    } else {
      // --- Update existing ---
      if (fullNameValue) setValue("Full name", fullNameValue);
      if (vehicleValue) setValue("Vehicle", vehicleValue);
      if (serviceValue) setValue("Service", serviceValue);
      if (lead.ad_id) setValue("ad_id", lead.ad_id);
      if (lead.Source || lead.source) setValue("Source", lead.Source || lead.source || '');
      // If phone changed, update Contact No.
      if (oldPhone && newPhone && oldPhone !== newPhone) {
        setValue("Contact No.", newPhone);
        if (getColIdx("Contact") !== -1) setValue("Contact", newPhone);
      }
    }

    // ----- Structured data -----
    if (sd.services) setValue("Service Required", Array.isArray(sd.services) ? sd.services.join(', ') : sd.services);
    if (sd.nearestLocation) setValue("Nearest Location", sd.nearestLocation);
    if (sd.vehicleAgeMonths) setValue("Vehicle Age (Month)", sd.vehicleAgeMonths);
    if (sd.estimatedAmount) setValue("Estimated Amount", sd.estimatedAmount);
    if (sd.budget) setValue("Customer Budget", sd.budget);
    if (sd.interestLevel) setValue("Interest Level", sd.interestLevel);
    if (sd.fromHyderabad) setValue("Location", sd.fromHyderabad); // using Location column
    if (sd.masterDataCorrect) setValue("Master Data Correct?", sd.masterDataCorrect);
    if (sd.additionalNotes) setValue("Additional Notes", sd.additionalNotes);
    if (lead.notes) setValue("Additional Notes", lead.notes);

    // ----- Call tracking (column names match your sheet) -----
    if (lead.call1Date) setValue("Call 1 Date", lead.call1Date);
    if (lead.call1Status) setValue("Call 1 Status", lead.call1Status);
    if (lead.call2Date) setValue("Call 2 Date", lead.call2Date);
    if (lead.call2Status) setValue("Call 2 Status", lead.call2Status);
    if (lead.call3Date) setValue("Call 3 Date", lead.call3Date);
    if (lead.call3Status) setValue("Call 3 Status", lead.call3Status);
    if (lead.confirmDate) setValue("Confirmation Call Date", lead.confirmDate);
    
    let confirmationStatus = "Pending";
    if (lead.converted) confirmationStatus = "Completed";
    else if (lead.confirmDate) confirmationStatus = "Scheduled";
    setValue("Confirmation call", confirmationStatus);
    
    let rowStatus = "Active";
    if (lead.lost) rowStatus = "Lost";
    if (lead.converted) rowStatus = "Converted";
    setValue("Status", rowStatus);
    if (lead.nextAction) setValue("Next Action", lead.nextAction);

    Logger.log('Update successful. Row: ' + targetRow + ', isNew: ' + isNew);

    return ContentService.createTextOutput(JSON.stringify({ 
      status: "success", 
      row: targetRow,
      isNew: isNew,
      message: isNew ? "Lead created successfully" : "Lead updated successfully"
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log('Error: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "error", 
      message: error.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ----- Helper: compute Next Action -----
function getNextActionFromLead(lead) {
  if (lead.lost) return 'Lost';
  if (lead.converted) return 'Converted';
  if (lead.call1Status === 'Pending') return 'Call 1';
  if (lead.call2Status === 'Pending') return 'Call 2';
  if (lead.call3Status === 'Pending') return 'Call 3';
  if (!lead.converted && !lead.confirmDate) return 'Confirmation';
  return 'Completed';
}

// ----- doGet: fetch leads (unchanged, but column names updated) -----
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
  
  const params = e.parameter || {};
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
      lead[header] = row[idx];
    });
    
    if (platform && lead.Platform !== platform) continue;
    
    if (startDate && endDate && lead.created_time) {
      const leadDate = new Date(lead.created_time);
      const start = new Date(startDate);
      const end = new Date(endDate);
      end.setHours(23, 59, 59);
      if (leadDate < start || leadDate > end) continue;
    }
    
    if (search) {
      const nameMatch = (lead['Full name'] || '').toLowerCase().includes(search);
      const phoneMatch = (lead['Contact No.'] || '').toLowerCase().includes(search);
      const salesMatch = (lead['Sales Person'] || '').toLowerCase().includes(search);
      if (!nameMatch && !phoneMatch && !salesMatch) continue;
    }
    
    leads.push(lead);
  }
  
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