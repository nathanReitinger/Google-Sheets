// Script allows automatic sorting of rows per date and 
// whenever cell values change--i.e., character match to
// "Resolved"--then the script moves that row to a second sheet

var sortCol=0; // 0 for first column and 1 for second column and so on...
var asc=false; // set variable asc to false for descending sort

// whenever cells are edited, look for values of "Resolved" and move those to new sheet
function onEdit(event) {
    // the range to srt
    sortRange("Sheet1","B4:Y126");
  
    // get the right spreadsheet 
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var s = ss.getActiveSheet();
    var r = ss.getActiveRange();
    if(s.getName() == "sheet1" && r.getColumn() == 3 && r.getValue() == "Resolved") {
        // get values to transfer over to new sheet
        var row = r.getRow();
        var numColumns = s.getLastColumn();
        var targetSheet = ss.getSheetByName("sheet2");
        var target = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
        s.getRange(row, 1, 1, numColumns).moveTo(target);
        // remove the row that was just moved
        s.deleteRow(row);
    }
};

// for entire spreadsheet, sort by date
function sortRange(sheetName,rangeName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sortSheet = ss.getSheetByName(sheetName);
  var range = sortSheet.getRange(rangeName);
  var activeSheet = ss.getActiveSheet();
  var activeRange = activeSheet.getActiveRange();
  var sortedValues;
  if( sortSheet.getName() == activeSheet.getName() &&
      activeRange.getLastRow() >= range.getRow() && 
      activeRange.getRow() <= range.getLastRow() &&
      activeRange.getLastColumn() >= range.getColumn() && 
      activeRange.getColumn() <= range.getLastColumn() )
      { 
        sortedValues=range.getValues().sort(mySortFunction);
        range.setValues(sortedValues);
      }
};

// see http://igoogledrive.blogspot.com/2013/08/auto-sort-rows-after-modifying.html
var mySortFunction = function(a,b) {
  try{x=a[sortCol].toLowerCase();
      y=b[sortCol].toLowerCase();}
  catch(e){x=a[sortCol];y=b[sortCol];}
  return (x>y)?(asc?1:-1):(x<y)?(asc?-1:1):0

}