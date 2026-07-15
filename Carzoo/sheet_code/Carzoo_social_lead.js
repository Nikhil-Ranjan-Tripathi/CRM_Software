// ==================================================
// CARZOO CRM - Social Media Leads Sheet Handler
// ==================================================

function doPost(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();
    
    const headers = data[0].map(h => h.toString().trim());

    function getColIdx(name) {
      let idx = headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
      if (idx !== -1) return idx;
      idx = headers.findIndex(h => h.toLowerCase().includes(name.toLowerCase()));
      return idx;
    }

    const lead = JSON.parse(e.postData.contents);
    const incomingPhone = String(lead.phone || '').replace(/\D/g, '');
    
    if (!incomingPhone) {
      output.setContent(JSON.stringify({ 
        status: "error", 
        message: "No phone number provided" 
      }));
      return output;
    }

    let contactIdx = getColIdx("Contact");
    if (contactIdx === -1) contactIdx = getColIdx("Phone");
    if (contactIdx === -1) contactIdx = getColIdx("Contact No");
    
    if (contactIdx === -1) {
      output.setContent(JSON.stringify({ 
        status: "error", 
        message: "Contact column not found" 
      }));
      return output;
    }

    let targetRowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      const rowPhone = String(data[i][contactIdx]).replace(/\D/g, '');
      if (rowPhone === incomingPhone && rowPhone !== "") {
        targetRowIndex = i + 1;
        break;
      }
    }

    const isNewRecord = (targetRowIndex === -1);
    if (isNewRecord) {
      targetRowIndex = sheet.getLastRow() + 1;
    }

    function updateCell(columnName, value) {
      if (value === undefined || value === null) return;
      if (value === '') return;
      const idx = getColIdx(columnName);
      if (idx !== -1 && targetRowIndex !== -1) {
        sheet.getRange(targetRowIndex, idx + 1).setValue(value);
      }
    }

    const sd = lead.structuredData || {};

    // Set basic info for new records
    if (isNewRecord) {
      updateCell("created_time", new Date().toISOString());
      updateCell("Platform", lead.platform);
      updateCell("Full name", lead.name);
      updateCell("Contact No.", lead.phone);
      updateCell("Vehicle", lead.vehicle);
      updateCell("Service", lead.service);
      updateCell("created_time", lead.date);
    }

    // Always update call progress fields
    updateCell("Call 1 Date", lead.call1Date);
    updateCell("Call 1 Status", lead.call1Status);
    updateCell("Call 2 Date", lead.call2Date);
    updateCell("Call 2 Status", lead.call2Status);
    updateCell("Call 3 Date", lead.call3Date);
    updateCell("Call 3 Status", lead.call3Status);
    updateCell("Confirmation Call Date", lead.confirmDate);
    
    let confirmationStatus = "Pending";
    if (lead.converted) confirmationStatus = "Completed";
    else if (lead.confirmDate) confirmationStatus = "Scheduled";
    updateCell("Confirmation call", confirmationStatus);

    // Update structured data
    updateCell("Service Required", sd.services ? sd.services.join(', ') : '');
    updateCell("Vehicle Age (Month)", sd.vehicleAgeMonths);
    updateCell("Estimated Amount", sd.estimatedAmount);
    updateCell("Customer Budget", sd.budget);
    updateCell("Interest Level", sd.interestLevel);
    updateCell("Location", sd.fromHyderabad);
    updateCell("Master Data Correct?", sd.masterDataCorrect);
    updateCell("Additional Notes", sd.additionalNotes || lead.notes);
    
    let rowStatus = "Active";
    if (lead.lost) rowStatus = "Lost";
    if (lead.converted) rowStatus = "Converted";
    updateCell("Status", rowStatus);
    updateCell("Next Action", lead.nextAction);

    output.setContent(JSON.stringify({ 
      status: "success", 
      row: targetRowIndex,
      isNew: isNewRecord,
      message: isNewRecord ? "Lead created successfully" : "Lead updated successfully"
    }));
    return output;

  } catch (error) {
    console.error('Error in doPost:', error);
    output.setContent(JSON.stringify({ 
      status: "error", 
      message: error.toString() 
    }));
    return output;
  }
}

function doGet(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length < 2) {
      output.setContent(JSON.stringify({ 
        status: "success", 
        data: [],
        message: "No data found"
      }));
      return output;
    }
    
    const headers = data[0].map(h => h.toString().trim());
    const leads = [];
    
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
      leads.push(lead);
    }
    
    output.setContent(JSON.stringify({ 
      status: "success", 
      data: leads,
      count: leads.length
    }));
    return output;
    
  } catch (error) {
    output.setContent(JSON.stringify({ 
      status: "error", 
      message: error.toString() 
    }));
    return output;
  }
}