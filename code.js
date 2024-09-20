function copyData(row) {
  const thisSS = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = thisSS.getSheetByName("REKAP ORDER");
  var sourceData = activeSheet.getRange(row,1,1,16).getDisplayValues()[0];
  var template = thisSS.getSheetByName("Invoice Template");
  template.getRange("i2").setValue(sourceData[1]);
  template.getRange("i3").setNumberFormat('@STRING@').setValue(sourceData[0]);
  template.getRange("i4").setValue(sourceData[5]);
  template.getRange("i5").setValue(sourceData[6]);
  template.getRange("d8").setValue("Plat Cutting " + sourceData[7]);
  var raw = sourceData[8];
  var items = [];
  raw.split("\n").forEach(function(line){
    if (line.match(/^\d*\..*×.*=.*/gm) != null){
      items.push(line);
    }
  });
  for (let i = 0; i < 10; i++) {
    var itemIndex = 9 + i;
    if (items[i] != null) {
//      var no = line.split(".")[0].match(/\d+/)[0];
      var no = i+1;
      var pieces = items[i].split("=")[1].match(/\d+/)[0];
      var panjangxlebar = items[i].split("=")[0].split(".").slice(1).join(".");
      template.getRange("a"+itemIndex+":d"+itemIndex).setValues([[no,pieces,"Pcs",panjangxlebar]]);
    } else {
      template.getRange("a"+itemIndex+":d"+itemIndex).clearContent();
    }
  }
  
  if (sourceData[10] != "") {
    var qty = parseFloat(sourceData[10]);
    template.getRange("A20:D20").setValues([[1,qty,"Lbr","Plat "+sourceData[7]]]);
  } else {
    template.getRange("A20:D20").clearContent();
  }
  template.getRange("J9").setValue(sourceData[11]);
  template.getRange("J20").setValue(sourceData[12]);
  template.getRange("J23").setValue(sourceData[13]);
  template.getRange("J24").setValue(sourceData[14]);
  template.getRange("J25").setValue(sourceData[15]);
  removeEmptyColumns(template);
  removeEmptyRows(template);
}

//Remove All Empty Columns in the Entire Workbook
function removeEmptyColumns(sheet) {
var maxColumns = sheet.getMaxColumns(); 
var lastColumn = sheet.getLastColumn();
if (maxColumns-lastColumn != 0){
      sheet.deleteColumns(lastColumn+1, maxColumns-lastColumn);
      }
  }

//Remove All Empty Rows in the Entire Workbook
function removeEmptyRows(sheet) {
var maxRows = sheet.getMaxRows(); 
var lastRow = sheet.getLastRow();
if (maxRows-lastRow != 0){
      sheet.deleteRows(lastRow+1, maxRows-lastRow);
      }
  }

var PRINT_OPTIONS = {
//  'size': 7,               // paper size. 0=letter, 1=tabloid, 2=Legal, 3=statement, 4=executive, 5=folio, 6=A3, 7=A4, 8=A5, 9=B4, 10=B
  'size' : "8.26x5.51",
  'fzr': false,            // repeat row headers
  'portrait': false,        // false=landscape
//  'fitw': true,            // fit window or actual size
  'scale': 4,
  'gridlines': false,      // show gridlines
  'printtitle': false,
  'sheetnames': false,
  'pagenum': 'UNDEFINED',  // CENTER = show page numbers / UNDEFINED = do not show
  'attachment': false,
  'top_margin':0.2,
  'bottom_margin':0.07,
  'left_margin':0.2,
  'right_margin':0.2,
}

var PDF_OPTS = objectToQueryString(PRINT_OPTIONS);

function printSelectedRange() {
  SpreadsheetApp.flush();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Invoice Template");
  var range = sheet.getRange("a1:j26");

  var gid = sheet.getSheetId();
  var printRange = objectToQueryString({
    'c1': range.getColumn() - 1,
    'r1': range.getRow() - 1,
    'c2': range.getColumn() + range.getWidth() - 1,
    'r2': range.getRow() + range.getHeight() - 1
  });
  var url = ss.getUrl().replace(/edit$/, '') + 'export?format=pdf' + PDF_OPTS + printRange + "&gid=" + gid;

  var htmlTemplate = HtmlService.createTemplateFromFile('js');
  htmlTemplate.url = url;
  SpreadsheetApp.getUi().showModalDialog(htmlTemplate.evaluate().setHeight(10).setWidth(100), 'Printing...');
}

function objectToQueryString(obj) {
  return Object.keys(obj).map(function(key) {
    return Utilities.formatString('&%s=%s', key, obj[key]);
  }).join('');
}

function printasPDF() {
  var aSheet = SpreadsheetApp.getActiveSheet();
  var aCell = aSheet.getActiveCell();
  var aColumn = aCell.getColumn();
  var aRow = aCell.getRow();

  if (aSheet.getName() == "REKAP ORDER" && aColumn == "17" && aCell.getValue() == 'PRINT') {
    copyData(aRow);
    printSelectedRange();
    aCell.clearContent();
  } else {
    //do nothing
  }
  return
}
