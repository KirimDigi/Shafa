/**
 * Google Apps Script for Shafa's Aqiqah Digital Invitation
 * Spreadsheet URL: https://docs.google.com/spreadsheets/d/1f8EymowFz6bGgJSL6yt7lfGlfFECMKCIPMkY1KT9GG8/edit?gid=0#gid=0
 * Sheet Name: Sheet1
 * 
 * Instructions:
 * 1. Open your Google Sheet.
 * 2. Click Extensions > Apps Script.
 * 3. Delete any default code and paste this code.
 * 4. Replace the SPREADSHEET_ID variable below with your actual sheet ID (already preset below).
 * 5. Click Save (disk icon).
 * 6. Click Deploy > Manage deployments.
 * 7. Choose "New version" or deploy as a new Web App (set execution as "Me" and access to "Anyone").
 * 8. Copy the Web App URL and replace the scriptUrl variable in index.html.
 */

var SPREADSHEET_ID = "1f8EymowFz6bGgJSL6yt7lfGlfFECMKCIPMkY1KT9GG8";
var SHEET_NAME = "Sheet1";

function doGet(e) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    
    // Auto-create headers in Row 1 if missing or blank
    var firstRow = sheet.getRange(1, 1, 1, 5).getValues()[0];
    var isHeaderMissing = !firstRow[0] || firstRow[0].toString().trim() === "";
    if (isHeaderMissing) {
      sheet.getRange(1, 1, 1, 5).setValues([["Timestamp", "Nama Tamu", "Ucapan", "Konfirmasi Kehadiran", "Jumlah Tamu"]]);
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    // Check if headers are still somehow empty, map fallback keys
    if (!headers[0] || headers[0].toString().trim() === "") {
      headers = ["timestamp", "nama tamu", "ucapan", "konfirmasi kehadiran", "jumlah tamu"];
    }
    
    var jsonArray = [];
    for (var i = 1; i < data.length; i++) {
      // Skip empty rows
      if (data[i].join("").trim() === "") continue;
      
      var row = {};
      for (var j = 0; j < headers.length; j++) {
        var headerName = headers[j] ? headers[j].toString().toLowerCase().trim() : "column_" + j;
        row[headerName] = data[i][j];
      }
      jsonArray.push(row);
    }
    
    return ContentService.createTextOutput(JSON.stringify(jsonArray))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    
    // Ensure headers exist in Row 1 before appending any data
    var firstRow = sheet.getRange(1, 1, 1, 5).getValues()[0];
    var isHeaderMissing = !firstRow[0] || firstRow[0].toString().trim() === "";
    if (isHeaderMissing) {
      sheet.getRange(1, 1, 1, 5).setValues([["Timestamp", "Nama Tamu", "Ucapan", "Konfirmasi Kehadiran", "Jumlah Tamu"]]);
    }
    
    var timestamp = new Date();
    var name = "";
    var ucapan = "";
    var kehadiran = "";
    var jumlah = 1;
    
    if (e.postData && e.postData.contents) {
      try {
        var json = JSON.parse(e.postData.contents);
        name = json.name || json.nama || json["nama tamu"] || "";
        ucapan = json.comment || json.ucapan || "";
        kehadiran = json.attendance || json.kehadiran || json["konfirmasi kehadiran"] || "";
        jumlah = json.guest_count || json.jumlah || json["jumlah tamu"] || 1;
      } catch (err) {
        name = e.parameter.name || e.parameter.nama || e.parameter["nama tamu"] || "";
        ucapan = e.parameter.comment || e.parameter.ucapan || "";
        kehadiran = e.parameter.attendance || e.parameter.kehadiran || e.parameter["konfirmasi kehadiran"] || "";
        jumlah = e.parameter.guest_count || e.parameter.jumlah || e.parameter["jumlah tamu"] || 1;
      }
    } else {
      name = e.parameter.name || e.parameter.nama || e.parameter["nama tamu"] || "";
      ucapan = e.parameter.comment || e.parameter.ucapan || "";
      kehadiran = e.parameter.attendance || e.parameter.kehadiran || e.parameter["konfirmasi kehadiran"] || "";
      jumlah = e.parameter.guest_count || e.parameter.jumlah || e.parameter["jumlah tamu"] || 1;
    }
    
    // Append the row matching columns:
    // 1. timestamp, 2. nama tamu, 3. ucapan, 4. konfirmasi kehadiran, 5. jumlah tamu
    sheet.appendRow([timestamp, name, ucapan, kehadiran, jumlah]);
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Data saved successfully" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
