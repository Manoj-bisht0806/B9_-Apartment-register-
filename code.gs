function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("B9 Chauhan Apartment Collection Register")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSheetsInfo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var result = {
    spreadsheetId: ss.getId(),
    sheets: []
  };
  
  for (var i = 0; i < sheets.length; i++) {
    result.sheets.push({
      name: sheets[i].getName(),
      gid: sheets[i].getSheetId()
    });
  }
  return result;
}

// -------------------------------------------------------------
// OnEdit Trigger: Syncs Subtotals & Summaries on Manual Edits
// -------------------------------------------------------------
function onEdit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  
  for (var k = 0; k < sheets.length; k++) {
    if (sheets[k].getName() !== "Overall Summary") {
      updateIndividualFloorSheet(sheets[k]);
    }
  }
  updateOverallCollectionSummary(ss);
}

// -------------------------------------------------------------
// Add Collection Record Function
// -------------------------------------------------------------
function addCollectionRecord(pass, targetSheet, flatNo, residentName, date, amount, paymentMode, transactionRef, fileData) {
  var SECRET_PASSWORD = "admin"; // Apna Admin Password yahan dalein
  
  if (pass !== SECRET_PASSWORD) {
    return { success: false, message: "गलत पासवर्ड! आप एंट्री नहीं कर सकते।" };
  }
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(targetSheet);
    if (!sheet) return { success: false, message: "Sheet नहीं मिली!" };
    
    // 1. Payment Receipt/Screenshot Upload to Drive
    var proofUrl = "Not Required";
    if (fileData) {
      var folderName = "Apartment_Collection_Receipts";
      var folders = DriveApp.getFoldersByName(folderName);
      var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
      
      var blob = Utilities.newBlob(Utilities.base64Decode(fileData.bytes), fileData.mimeType, fileData.fileName);
      var uploadedFile = folder.createFile(blob);
      uploadedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      proofUrl = uploadedFile.getUrl();
    }
    
    // 2. Find Insert Position & Last S.No.
    var data = sheet.getDataRange().getDisplayValues();
    var insertRowIndex = -1;
    var lastSNo = 0;
    
    for (var i = 0; i < data.length; i++) {
      var rowStr = data[i].join(" ");
      if (rowStr.indexOf("Floor Total") !== -1 || rowStr.indexOf("Month-wise subtotal") !== -1) {
        insertRowIndex = i + 1;
        break;
      }
      var val = parseInt(data[i][0]);
      if (!isNaN(val) && val > 0) {
        lastSNo = val;
      }
    }
    
    var d = new Date(date);
    var monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    var monthStr = monthNames[d.getMonth()] + " " + d.getFullYear();
    var newSNo = lastSNo + 1;
    var parsedAmount = parseFloat(amount) || 0;
    
    var formattedDate = ("0" + d.getDate()).slice(-2) + "/" + ("0" + (d.getMonth() + 1)).slice(-2) + "/" + d.getFullYear();
    var formattedAmount = "₹" + parsedAmount.toLocaleString('en-IN');
    
    var newRowData = [newSNo, monthStr, formattedDate, flatNo, residentName, formattedAmount, paymentMode, transactionRef, proofUrl];

    var targetRow = insertRowIndex;
    if (insertRowIndex !== -1) {
      sheet.insertRowBefore(insertRowIndex);
      targetRow = insertRowIndex;
    } else {
      sheet.appendRow(newRowData);
      targetRow = sheet.getLastRow();
    }
    
    sheet.getRange(targetRow, 1, 1, newRowData.length).setValues([newRowData]);

    // 3. AUTO-APPLY STANDARD ROW FORMATTING
    applyStandardCollectionRowFormatting(sheet, targetRow, newRowData.length);
    
    // 4. Update Subtotals & Summary
    updateIndividualFloorSheet(sheet);
    updateOverallCollectionSummary(ss);
    
    SpreadsheetApp.flush();
    return { success: true, message: "कलेक्शन रिकॉर्ड सफलतापूर्वक जुड़ गया!" };
  } catch (e) {
    return { success: false, message: "एरर: " + e.toString() };
  }
}

// -------------------------------------------------------------
// Universal Standard Formatting Helper
// -------------------------------------------------------------
function applyStandardCollectionRowFormatting(sheet, rowNum, colCount) {
  var rowRange = sheet.getRange(rowNum, 1, 1, colCount);
  
  rowRange.setFontWeight("normal");
  rowRange.setBackground("#ffffff");
  rowRange.setFontColor("#000000");
  rowRange.setFontFamily("Arial");
  rowRange.setFontSize(10);
  rowRange.setBorder(true, true, true, true, true, true, "#d3d3d3", SpreadsheetApp.BorderStyle.SOLID);
  rowRange.setWrap(true);

  sheet.getRange(rowNum, 1).setHorizontalAlignment("center"); // S.No
  sheet.getRange(rowNum, 2).setHorizontalAlignment("left");   // Month
  sheet.getRange(rowNum, 3).setHorizontalAlignment("center"); // Date
  sheet.getRange(rowNum, 4).setHorizontalAlignment("center"); // Flat No
  sheet.getRange(rowNum, 5).setHorizontalAlignment("left");   // Resident Name
  sheet.getRange(rowNum, 6).setHorizontalAlignment("right");  // Amount
  sheet.getRange(rowNum, 7).setHorizontalAlignment("center"); // Payment Mode
  sheet.getRange(rowNum, 8).setHorizontalAlignment("left");   // Txn Ref
  if (colCount >= 9) {
    sheet.getRange(rowNum, 9).setHorizontalAlignment("center"); // Receipt Link
  }
}

// -------------------------------------------------------------
// Floor Subtotals Calculation Function
// -------------------------------------------------------------
function updateIndividualFloorSheet(sheet) {
  var data = sheet.getDataRange().getDisplayValues();
  var overallTotal = 0;
  var monthMap = {};
  var floorTotalRow = -1;

  for (var i = 0; i < data.length; i++) {
    var sNo = parseInt(data[i][0]);
    var monthName = data[i][1] ? data[i][1].trim() : "";
    var amtStr = String(data[i][5]).replace(/[^0-9.]/g, '');
    var amtVal = parseFloat(amtStr);

    if (!isNaN(sNo) && sNo > 0 && monthName !== "") {
      if (!isNaN(amtVal)) {
        overallTotal += amtVal;
        if (!monthMap[monthName]) {
          monthMap[monthName] = { total: 0, count: 0 };
        }
        monthMap[monthName].total += amtVal;
        monthMap[monthName].count += 1;
      }
    }

    var fullRowStr = data[i].join(" ");
    if (fullRowStr.indexOf("Floor Total") !== -1) {
      floorTotalRow = i + 1;
    }
  }

  if (floorTotalRow !== -1) {
    sheet.getRange(floorTotalRow, 6).setValue("₹" + overallTotal.toLocaleString('en-IN')).setHorizontalAlignment("right");
  }
}

// -------------------------------------------------------------
// Overall Collection Summary Updating Function
// -------------------------------------------------------------
function updateOverallCollectionSummary(ss) {
  var summarySheet = ss.getSheetByName("Overall Summary");
  if (!summarySheet) return;

  var sheets = ss.getSheets();
  var globalMonthMap = {};
  var globalFloorMap = {};

  for (var k = 0; k < sheets.length; k++) {
    var s = sheets[k];
    var sName = s.getName();
    if (sName !== "Overall Summary") {
      var data = s.getDataRange().getDisplayValues();
      var fTotal = 0;
      var fCount = 0;

      for (var i = 0; i < data.length; i++) {
        var sNo = parseInt(data[i][0]);
        var mName = data[i][1] ? data[i][1].trim() : "";
        var amtStr = String(data[i][5]).replace(/[^0-9.]/g, '');
        var amtVal = parseFloat(amtStr);

        if (!isNaN(sNo) && sNo > 0) {
          if (!isNaN(amtVal)) {
            fTotal += amtVal;
            fCount += 1;

            if (mName !== "") {
              if (!globalMonthMap[mName]) {
                globalMonthMap[mName] = { total: 0, count: 0 };
              }
              globalMonthMap[mName].total += amtVal;
              globalMonthMap[mName].count += 1;
            }
          }
        }
      }
      globalFloorMap[sName] = { total: fTotal, count: fCount };
    }
  }

  var sData = summarySheet.getDataRange().getDisplayValues();
  var monthGrandTotalRow = -1;

  for (var i = 0; i < sData.length; i++) {
    var colA = sData[i][0] ? sData[i][0].trim() : "";

    if (globalMonthMap[colA]) {
      summarySheet.getRange(i + 1, 2).setValue("₹" + globalMonthMap[colA].total.toLocaleString('en-IN')).setHorizontalAlignment("right");
      summarySheet.getRange(i + 1, 3).setValue(globalMonthMap[colA].count).setHorizontalAlignment("center");
    }

    if (globalFloorMap[colA]) {
      summarySheet.getRange(i + 1, 2).setValue("₹" + globalFloorMap[colA].total.toLocaleString('en-IN')).setHorizontalAlignment("right");
      summarySheet.getRange(i + 1, 3).setValue(globalFloorMap[colA].count).setHorizontalAlignment("center");
    }

    if (colA === "Grand Total") {
      monthGrandTotalRow = i + 1;
    }
  }
}
