
var newBoLNumber= -1;
var newInvoiceNumber = -1;
var removeOldPDFFilesWithSameName = true;

//var productVolumes = new Map(["1/2 Barrel",15.5],["1/6 Barrel", 5.167], ["750 mL", .1981]);
//var keyValue = {
/**
 Called when spreasheet is opened. Create menus.
*/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  //SpreadsheetApp.getUi().alert("Creating Menu ");
  // Or DocumentApp or FormApp.
  ui.createMenu('Invoices')
      .addItem('New Invoice', 'newInvoice')
      //.addItem('Generate Bill of Lading', 'generateBillofLading')
      .addItem('Finalize Invoice', 'finalizeInvoice')
      //.addSeparator()
      //.addItem('Generate BoL', 'generateBillofLading')
      //.addSubMenu(ui.createMenu('Sub-menu')
          //.addItem('Second item', 'menuItem2'))
      .addToUi();
}

function menuItem1() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('You clicked the first menu item!');
}

function menuItem2() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('You clicked the second menu item!');
}


/**

*/
function triggerOnEdit(e) {
  //SpreadsheetApp.getUi().alert("trigger called ");
  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();
  if(sheetName == invoiceSheetName) {
    //setInvoiceValues(e);  Converted to use vlookup()
  } else if (sheetName == bolSheetName) {
    //setBoLValues(e);
  }
}

/**
  Set price and volume when product is selected in Invoice sheet 
    called from trigger
*/
function setInvoiceValues(e) {
  //SpreadsheetApp.getUi().alert("in showMessageOnUpdate()");
  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();
  var value = range.getValue();
  //SpreadsheetApp.getUi().alert("row: " +row +" column:" +col +" value: " +value);
  if(col == invoiceProductColumn && row >= invoiceProductStartRow) {
    //SpreadsheetApp.getUi().alert("getting product info");
    var wholesalePrice = "";
    var productVolume = "";
    var productSize = "";
    
    if(value != "") {
      var productsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(productsSheetName);
      var productRow = findProductNameRow(value, productsSheet);
  
      console.log("range updated row: " +row +" column:" +col +" value: " +value +" productRow = " +productRow);//   range.getA1Notation());
      wholesalePrice = getCellValue(productsSheet, productRow, productWholesalePriceColumn);
      productVolume = getCellValue(productsSheet, productRow, productVolumeColumn);
      productSize = getCellValue(productsSheet, productRow, productSizeColumn);
      //SpreadsheetApp.getUi().alert("wholesale price = " +wholesalePrice +" volume = " +productVolume);
    }
    
    // Set price
    var priceCell = range.offset(0, (invoicePriceColumn - col), 1, 1);
    priceCell.setValue(wholesalePrice);
    
    //Set volume
    var volCell = range.offset(0, -2, 1, 1);
    volCell.setValue(productVolume);
    
    //Set Size
    var sizeCell = range.offset(0, -3, 1, 1);
    sizeCell.setValue(productSize);
  }
} //setInvoiceValues

/**
  Set weight and volume when product is selected in BoL sheet 
    called from trigger
*/
function setBoLValues(e) {
  //SpreadsheetApp.getUi().alert("in setBoLValues()");
  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();
  var value = range.getValue();
  //SpreadsheetApp.getUi().alert("row: " +row +" column:" +col +" value: " +value);
  if(col == invoiceProductColumn && row >= invoiceProductStartRow) {
    //SpreadsheetApp.getUi().alert("getting product info");
    var productWeight = "";
    var productVolume = "";
    
    if(value != "") {
      var productsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(productsSheetName);
      var productRow = findProductNameRow(value, productsSheet);
  
      //SpreadsheetApp.getUi().alert("range updated row: " +row +" column:" +col +" value: " +value +" productRow = " +productRow);//   range.getA1Notation());
      productWeight = getCellValue(productsSheet, productRow, productWeightColumn);
      productVolume = getCellValue(productsSheet, productRow, productVolumeColumn);
      console.log("setBoLValues(): weight = " +productWeight +" volume = " +productVolume);
    }
    
    // Set price
    var weightCell = range.offset(0, (bolWeightColumn - col), 1, 1);
    weightCell.setValue(productWeight);
    
    //Set volume
    var volCell = range.offset(0, -(col - bolVolumeColumn), 1, 1);
    volCell.setValue(productVolume);
  }
} //setBoLValues


/**
 findProductNameRow
  find row number of product name in sheet
  
  @return row number
*/
function findProductNameRow(productName, sheet) {
  var productNameColumn = 1;
  
  console.log("finding row for product: " +productName);
  for (var i = 1; i < sheet.getMaxRows(); i++) {
    var vals = sheet.getSheetValues(i, productNameColumn, 1, 1);
    var val = vals[0][0];
    console.log("value: " +val);
    console.log("in findProductNameRow() - checking cell:" +i +":" +productNameColumn);
    if (val == productName) {
      return i;
    }
  }
  return -1;
} //findProductNameRow

/**
 initialize new invoice
*/
function newInvoice() {
  var start = new Date();
  //SpreadsheetApp.getUi().alert("in newInvoice()");
  var invoiceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(invoiceSheetName);
  var summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(summarySheetName);
  var bolSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(bolSheetName);
  
  // Clear spreadsheet
  for(row = invoiceProductStartRow; row <= invoiceProductEndRow; row++) {
    range = invoiceSheet.getRange(row, invoiceProductColumn);
    range.setValue("");
    /* Not needed after converting to use vlookup()
    range = invoiceSheet.getRange(row, invoicePriceColumn);
    range.setValue("");
    range = invoiceSheet.getRange(row, invoiceVolumeColumn);
    range.setValue("");
    range = invoiceSheet.getRange(row, invoiceSizeColumn);
    range.setValue("");
    */
    range = invoiceSheet.getRange(row, invoiceQuantityColumn);
    range.setValue("");
  }
  
  // zero keg returns
  range = invoiceSheet.getRange(kegReturnRow, halfReturnColumn);
  range.setValue(0);
  range = invoiceSheet.getRange(kegReturnRow, sixtelReturnColumn);
  range.setValue(0);
  
  // Set date
  var d = new Date();
  var today = (d.getMonth()+1) +"/" +d.getDate() +"/" +d.getFullYear();
  console.log("date: " +today);
  range = invoiceSheet.getRange(invoiceDateRow, invoiceDateColumn);
  range.setValue(today);
  range = bolSheet.getRange(bolDateRow, bolDateColumn);
  range.setValue(today);
  
  // Set new invoice number
  var lastRow = getLastRowOfColumn(summarySheet, invoiceSummaryInvoiceNumberColumn);
  var lastInvoiceNumber = getCellValue(summarySheet, lastRow, invoiceSummaryInvoiceNumberColumn);
  var lastBoLNumber = getCellValue(summarySheet, lastRow, invoiceSummaryBoLNumberColumn);
  console.log("lastInvoiceNumber: " +lastInvoiceNumber, +" lastBoLNumber: " +lastBoLNumber);
  range = invoiceSheet.getRange(invoiceNumberRow, invoiceNumberColumn);
  range.setValue(lastInvoiceNumber+1);
  range = bolSheet.getRange(bolNumberRow, bolNumberColumn);
  range.setValue(lastBoLNumber + 1);
  newInvoiceNumber = lastInvoiceNumber+1;
  newBoLNumber = lastBoLNumber+1;
  
  var end = new Date()
  console.log(" newInvoice() time: " +end.getTime() - start.getTime());
} //newInvoice

/*
   Update totals on summary sheet and generate PDF copies of invoice and bill of lading

*/
function finalizeInvoice() {
  console.log("in finalizeInvoice()");
  var start = new Date();
  
  // generating BoL
  //generateBillofLading();
  
  var invoiceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(invoiceSheetName);
  var bolSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(bolSheetName);
  var summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(summarySheetName);
  var productsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(productsSheetName);  
  
  var invoiceNumber = getCellValue(invoiceSheet, invoiceNumberRow, invoiceNumberColumn);
  
  var lastRow = getLastRowOfColumn(summarySheet, invoiceSummaryInvoiceNumberColumn);
  var lastInvoiceNumber = getCellValue(summarySheet, lastRow, invoiceSummaryInvoiceNumberColumn);
  if(lastInvoiceNumber == invoiceNumber) {
    console.log("Replacing Summary row");
    summarySheet.deleteRow(lastRow);
  } else {
    var newRow = lastRow+1;
  }
  
  console.log("setting summary date");
  var dateString = getCellValue(invoiceSheet, invoiceDateRow, invoiceDateColumn);
  console.log("date:" +dateString);

  console.log("setting summary invoice number");
  var invoiceNumber = getCellValue(invoiceSheet, invoiceNumberRow, invoiceNumberColumn);
  console.log("invoiceNumber:" +invoiceNumber);
  
  var bolNumber = getCellValue(bolSheet, bolNumberRow, bolNumberColumn);

  console.log("setting summary beer sale");
  var beerSale = getCellValue(invoiceSheet, invoiceBeerTotalSaleRow, invoiceTotalSaleColumn);
  
  console.log("setting summary total sale");
  var totalSale = getCellValue(invoiceSheet, invoiceTotalSaleRow, invoiceTotalSaleColumn);

  console.log("setting summary sale");    
  var totalVolume = getCellValue(invoiceSheet, invoiceTotalVolumeRow, invoiceTotalVolumeColumn);
  
  summarySheet.appendRow([dateString, invoiceNumber, bolNumber, beerSale, totalSale, totalVolume]);
  console.log("Total Sale: " +totalSale);
  
  calculateVolumeTotalsByProduct(invoiceSheet, summarySheet, productsSheet);
  
  // Update totals in summary sheet
  
  // get total sales w/o deposits and returns
  
  // get product volume totals
  
  
  // generate PDF copy of Invoice in Invoices directory
  
  console.log("lastInvoiceNumber: " +lastInvoiceNumber);
  range = invoiceSheet.getRange(invoiceNumberRow, invoiceNumberColumn);
  var invoiceNumber = range.getValue();
  var invoiceFileName = "Inv" +invoiceNumber;
  var invoicesDirectory = "Invoices";
  //getDirectory(invoicesDirectory);
  generatePDF(invoiceSheet, invoicesDirectory, invoiceFileName);
  
  // generate PDF copy of BoL in BoL directory
  range = bolSheet.getRange(bolNumberRow, bolNumberColumn);
  var bolNumber = range.getValue();
  var bolFileName = "BOL " +bolNumber;
  var bolDirectory = "BOL's";
  //getDirectory(invoicesDirectory);
  generatePDF(bolSheet, bolDirectory, bolFileName);
  
  var end = new Date();
  console.log(" finalizeInvoice() time: " +(end.getTime() - start.getTime()) +"ms");
} //finalizeInvoice

/**
*/
function getDirectory(dirName) {
  console.log(" in getDirectory for name: " +dirName);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var file = DriveApp.getFileById(ss.getId());
  var folders = file.getParents();
  var currentDir = null;
  var newDir = null;
  console.log("getting parent folder");
  while (folders.hasNext()){
    currentDir = folders.next();
    console.log('folder name = '+currentDir.getName());
  }
  if(currentDir != null) {
    console.log("looking for child dir: " +dirName);
    var childDirs = currentDir.getFoldersByName(dirName);
    
    while(childDirs.hasNext()) {
      newDir = childDirs.next();
    }
    if(newDir == null) {
      newDir = currentDir.createFolder(dirName);
    }
  }
  return newDir;
} //getDirectory

/**
  Save a sheet as a PDF file
*/
function generatePDF(sheet, directory, filename) {
  var pdfStart = new Date().getTime();
  var pdfFilename = filename +".pdf";
  console.log("in generatePDF()");
  var newSpreadsheet = SpreadsheetApp.create(filename);
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //sheet = originalSpreadsheet.getActiveSheet();
  sheet.copyTo(newSpreadsheet);
  var initialSheet = newSpreadsheet.getSheetByName("Sheet1");
  newSpreadsheet.deleteSheet(initialSheet);
  
  // Repace cell values with text (to avoid broken references).
  var destSheet = newSpreadsheet.getSheets()[0];
  
  var sourceRange = sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns());
  var sourcevalues = sourceRange.getValues();
  var destRange = destSheet.getRange(1, 1, destSheet.getMaxRows(), destSheet.getMaxColumns());
  destRange.setValues(sourcevalues);
  
  //Save the desired sheet as pdf
  var pdf = DriveApp.getFileById(newSpreadsheet.getId()).getAs('application/pdf');
  var dir = getDirectory(directory);
  if(dir != null) {
    if(removeOldPDFFilesWithSameName) {
      console.log("calling deleteNamedFiles(): dir: " +directory +" filename: " +pdfFilename);
      deleteNamedFiles(dir, pdfFilename);
    } 
    var saveCopy = dir.createFile(pdf);
  } else {
    SpreadsheetApp.getUi().alert("Error: Can't create file");
  }
  
  //Delete temporary spreadsheet
  DriveApp.getFilesByName(filename).next().setTrashed(true);

  console.log("generatePDF() time: " +(new Date().getTime() - pdfStart) +"ms");
} //generatePDF

/**
  Calculate product volume totals
*/
function calculateVolumeTotalsByProduct(invoiceSheet, summarySheet, productsSheet) {
  var volStart = new Date().getTime();
  var product = "";
  var quantity = 0;
  var productCell;
  var quantityCell;
  var productTypeTotals = Object.create(FauxMap);
  var entry = null;
  
  for(row = invoiceProductStartRow; row <= invoiceProductEndRow; row++) {
    productCell = invoiceSheet.getRange(row, invoiceSizeColumn);
    if(!productCell.isBlank()) {
      product = productCell.getValue();
      
      quantityCell = invoiceSheet.getRange(row, invoiceQuantityColumn);
      quantity = quantityCell.getValue();
      if(!quantityCell.isBlank() && typeof quantity == "number") {
        if(productTypeTotals.containsKey(product)) {
          entry = productTypeTotals.get(product);
          entry += quantity;          
        } else {
          entry = quantity;
        }
        //console.log("Adding to FauxMap - key: " +product +" new quantity: " +entry +" row: " +row);
        productTypeTotals.put(product, entry);
      }
    }
  } 
  
  var keys = productTypeTotals.keys();
  var product;
  var columnIndex;
  var range;
  var unitVolume;
  
  var row = summarySheet.getLastRow();
  console.log("keys length: " +keys.length);
  
  //add product type quantity and total volume to row in summary sheet
  for(i = 0; i< keys.length; i++) {   
    console.log("product quantity loop, index: " +i);
    //get product quantity from invoice sheet
    product = keys[i];
    quantity = productTypeTotals.get(keys[i]);
    console.log("product quantities[" +i +"]: " +product +" : " + quantity);
    columnIndex = getColumnNumberByName(summarySheet, product);
    if(columnIndex > 0) {
      range = summarySheet.getRange(row, columnIndex);
      range.setValue(quantity); //productTypeTotals.get(product));
      console.log("just set quantity: " +quantity +" i: " +i);
    
      //get unit volume from product sheet
      unitVolume = getVolumeForProduct(productsSheet, product);
      range = summarySheet.getRange(row, columnIndex+1);
      range.setValue(unitVolume * quantity);
    }
    
  }
  
  console.log("toString test: " +productTypeTotals.toString());
  console.log("calculateVolumeTotalsByProduct() time: " +(new Date().getTime() - volStart) +"ms");
} //calculateVolumeTotalsByProduct

/**
  Find 1st row of product in the product sheet and get associated volume
  
  @return volume of product
*/
function getVolumeForProduct(productsSheet, product) {
  //var row = getRowNumberByName(productsSheet, productSizeColumn, product);
  var row = getRowNumberByName(productsSheet, productSizeColumn, product);
  console.log("getVolumeForProduct(): row: " +row +" for product: " +product);
  var volume = productsSheet.getRange(row, productVolumeColumn).getValue();
  
  return volume;
} //getVolumeForProduct

/**
  Find 1st row of product in the product sheet and get associated volume
  
  @return volume of product
*/
function getWeightForProduct(productsSheet, product) {
  //var row = getRowNumberByName(productsSheet, productSizeColumn, product);
  var weight = "";
  var row = getRowNumberByName(productsSheet, productNameColumn, product);
  console.log("getVolumeForProduct(): row: " +row +" for product: " +product);
  
  if(row > -1) {
    weight = productsSheet.getRange(row, productWeightColumn).getValue();
  }
  
  return weight;
} //getWeightForProduct

/**
*/
function getProductWeightFromSheet(product) {
  var productsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(productsSheetName);

  return getWeightForProduct(productsSheet, product);
}

/**
 get info and copy to Bill of Lading tab
 
 @deprecated
*/
function generateBillofLading () {
  console.log("in generateBillofLading()");
  var bolStart = new Date().getTime();
  var invoiceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(invoiceSheetName);
  var productSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(productsSheetName);
  var bolSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(bolSheetName);
  var summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(summarySheetName);
  var product;
  var size;
  var productRow;
  var weight;
  var volume;
  var quantity;
  var date = invoiceSheet.getRange(invoiceDateRow, invoiceDateColumn).getValue();
  
  // Set date from invoice
  bolSheet.getRange(bolDateRow, bolDateColumn).setValue(date);
  
  // Get new BoL number
  
  // Clear products
  console.log("clearing sheet");
  for(row = bolProductStartRow; row <= bolProductEndRow; row++) {
    range = bolSheet.getRange(row, bolProductColumn);
    range.setValue("");
    range = bolSheet.getRange(row, bolWeightColumn);
    range.setValue("");
    range = bolSheet.getRange(row, bolVolumeColumn);
    range.setValue("");
    range = bolSheet.getRange(row, bolQuantityColumn);
    range.setValue("");
    bolSheet.getRange(row, bolSizeColumn).setValue("");
  }
  
  // Copy rows from invoice
  console.log("copying rows from invoice");
  for(row = bolProductStartRow; row <= bolProductEndRow; row++) {
    var product = getCellValue(invoiceSheet, row, invoiceProductColumn);
    
    if(product != "") {
      volume = getCellValue(invoiceSheet, row, invoiceVolumeColumn);
      quantity = getCellValue(invoiceSheet, row, invoiceQuantityColumn);
      size = getCellValue(invoiceSheet, row, invoiceSizeColumn);
    
      productRow = findProductNameRow(product, productSheet);
      weight = getCellValue(productSheet, productRow, productWeightColumn);
      console.log("bol product: " +product +" weight: " +weight +" row: " +row);
    
      bolSheet.getRange(row, bolProductColumn).setValue(product);
      bolSheet.getRange(row, bolSizeColumn).setValue(size);
      bolSheet.getRange(row, bolQuantityColumn).setValue(quantity);
      bolSheet.getRange(row, bolVolumeColumn).setValue(volume);
      bolSheet.getRange(row, bolWeightColumn).setValue(weight);
    }
  }

  //get BoL number
  console.log("geting new BoL number");
  var lastRow = getLastRowOfColumn(summarySheet, invoiceSummaryBoLNumberColumn);
  console.log("Bol number in row: " +lastRow +" col: " +invoiceSummaryBoLNumberColumn);
  //var lastBoLNumber = getCellValue(summarySheet, lastRow, invoiceSummaryBolNumberColumn);
  //var newBoLNumber = lastBoLNumber + 1;
  //console.log(lastBoLNumber +" : " +newBoLNumber);
  //var bolNumberRange = bolSheet.getRange(bolNumberRow, bolNumberColumn);
  //bolNumberRange.setValue(newBoLNumber);
  console.log("generateBillofLading() time: " +(new Date().getTime() - bolStart) +"ms");
} //generateBillofLading

/**
   return last row number number with data in given column
   
   @return last row number
*/
function getLastRowOfColumn(sheet, column) {
  console.log("in getLastRowOfColumn()");
  if(sheet == null) { 
    SpreadsheetApp.getUi().alert("summary sheet not found");
    return -1;
  }
  var dataRange = sheet.getDataRange();
  //SpreadsheetApp.getUi().alert("found summary data: " +dataRange.getNumRows());
  var data = dataRange.getValues();
  var numRows = data.length;
  //SpreadsheetApp.getUi().alert("got Range: " +numRows);

 // loop from bottom to top for last row in given row "RowNumber"
  
  while( data[numRows - 1][column-1] == "" && numRows > 0 ) {
    //SpreadsheetApp.getUi().alert("row: " +numRows-1 +" col: " +column +" value: " +data[numRows - 1][column-1]);
    numRows--;
  }
  
  return numRows;
} //getLastRowOfColumn



