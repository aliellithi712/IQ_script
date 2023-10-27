function onEdit(e) {
  if (!e) return; // Check if event object is defined

var spreadsheet = SpreadsheetApp.getActive();
var sheet = SpreadsheetApp.getActiveSheet();

var cell = e.range; // Current Cell

var cellRange = sheet.getActiveCell();
var cellContent = sheet.getActiveCell().getValue();
var selectedColumn = cellRange.getColumn();
var selectedRow = cellRange.getRow();



/*  Format of the names */
function setFontAndSize(range, fontFamily, fontSize) {
  range.setFontFamily(fontFamily).setFontSize(fontSize);
}
var rangeB = sheet.getRange("B:B");
setFontAndSize(rangeB, "Calibri", 12);
var rangeL = sheet.getRange("L:L");
setFontAndSize(rangeL, "Calibri", 12);

///////////////////////////////////////////////////
/*  Write the people who flee */
var cellP1 = sheet.getRange("L1").getValue();
Logger.log("Cell P2 Value: " + cellP1);
if (cellP1 == "Yes"){
for(var j=2 ; j< 150; j++){
if (sheet.getRange("C"+ j).getValue() != '' && sheet.getRange("D"+j).getValue() == '') {
var name = sheet.getRange("B" + j).getValue();
var name2 = sheet.getRange("E" + j).getValue();
var isfound = false;
for (var i = 2; i < 150 ; i++){
  if(isfound){
    break;
  }
  var celltest = sheet.getRange("L"+ i );
  var celltest2 = sheet.getRange("M"+ i );
  if (celltest.getValue() == ''){
    celltest.setValue(name) 
    if (name2 == '') { celltest2.setValue(0) }
    else {celltest2.setValue(name2)}
    isfound = true;
  }
}
}
}
}
///////////////////////////////////////////////////

/*  Caught The name if flee */ 
function setter( j ,variable, cellE_t, name) { 
var temp = j + 1;
var name2 = sheet.getRange("M" + temp).getValue();
if (name2 == ''){
  var res = 35;
  variable.setValue(  ` عليه ${res} - ${name}` ); 
  cellE_t.setValue(res);
}
else {
  var intValue = parseInt(name2, 10); 
  var res = 35 + intValue;
  variable.setValue(  ` عليه ${res} - ${name}` ); 
  cellE_t.setValue(res);
}
}

///////////////////////////////////////////////////
/*  Clear if name came back */

if  (selectedColumn == 2 && e.oldValue !== cell.getValue()&& cell.getValue()!= ''){

var substring = "عليه";
var rangeP = sheet.getRange("L:L");
var valuesP = rangeP.getValues();
var rangeO = sheet.getRange("O:O");
var valuesO = rangeO.getValues();
var rangeN = sheet.getRange("N:N");
var valuesN = rangeN.getValues();
var rangeP2 = sheet.getRange("P:P");
var valuesP2 = rangeP.getValues();
var cellB = sheet.getRange("B" + selectedRow).getValue();
var cellB_t = sheet.getRange("B" + selectedRow)
var cellE_t = sheet.getRange("E" + selectedRow)






var limit = sheet.getRange("U" + 8).getValue();
for(var i=1; i<=limit ; i++){
if ((cellB == valuesO[i][0] && cellB != '') || (cellB == valuesN[i][0] && cellB != '')){
  var cell_1 = sheet.getRange("C" + selectedRow);
  var cell_2 = sheet.getRange("F" + selectedRow);
  var cell_3 = sheet.getRange("G" + selectedRow);
  var cell_O = sheet.getRange("O" + (i+1));
  var cell_N = sheet.getRange("N" + (i+1));
  var check = sheet.getRange("N" + (i+1)).getValue();
  var restOfDays = sheet.getRange("P" + (i+1));
  var startDate = new Date (sheet.getRange("R" + (i+1)).getValue());
  var endDate = new Date (sheet.getRange("S" + 2).getValue());
  var lifetime = parseInt((endDate - startDate) / 1000 / 60 / 60 / 24);
  restOfDays.setValue(lifetime);
  if (check === '--' && lifetime>30){
    cell_O.setValue(cell_O.getValue() + ' -- Expired');
    break;}
  else if (check !== '--' && lifetime>7){
    cell_N.setValue(cell_N.getValue() + ' -- Expired');
    break;}
  cell_1.setValue('###');
  cell_3.setValue('0');
  cell_2.setValue('0');
}}

for (var i = 1; i < selectedRow ; i++) {

if ((cellB == valuesP[i][0] && cellB != '') || ((valuesP[i][0].indexOf(cellB) !== -1)  && cellB != '' ) ) {
  if (valuesP[i][0].indexOf(substring) !== -1){
    var name = valuesP[i][0].split(" - ")[1];
  }
  else{
    var name = valuesP[i][0];
  }
  setter(i , cellB_t, cellE_t, name);
  var temp = i+1
  var cell = sheet.getRange("L" + temp);
  cell.clearContent();
  var cell = sheet.getRange("M" + temp); 
  cell.clearContent();
  // var cell = sheet.getRange("N" + temp);
  // cell.setValue('here');
  break;
}
}
}
///////////////////////////////////////////////////
/*  Subscriptions */

if (selectedColumn == 14 || selectedColumn == 15 && e.oldValue !== cell.getValue()&& cell.getValue()!= '') {
var restOfDays = sheet.getRange("P" + selectedRow);
var other = sheet.getRange("R" + selectedRow);
var startDate = new Date();
other.setValue(startDate);
var endDate = new Date (sheet.getRange("S" + 2).getValue());
var lifetime = parseInt((endDate - startDate) / 1000 / 60 / 60 / 24);
restOfDays.setValue(lifetime);

if (selectedColumn == 14) {
  var fees = sheet.getRange("Q" + selectedRow);
  fees.setValue(100);
  var other = sheet.getRange("O" + selectedRow);
  other.setValue('--');
}
else if (selectedColumn == 15){
  var fees = sheet.getRange("Q" + selectedRow);
  fees.setValue(200);
  var other = sheet.getRange("N" + selectedRow);
  other.setValue('--');
}
}

if (selectedColumn == 21 && selectedRow == 2){
if (cellContent == 'Yes'){
  var limit = sheet.getRange("U" + 8).getValue();
  limit = limit + 1;
  for(var i=2 ; i <= limit ; i++ ){
    var restOfDays = sheet.getRange("P" + i);
    var startDate = new Date (sheet.getRange("R" + i).getValue());
    var endDate = new Date (sheet.getRange("S" + 2).getValue());
    var lifetime = parseInt((endDate - startDate) / 1000 / 60 / 60 / 24);
    restOfDays.setValue(lifetime);

  }
}
}


///////////////////////////////////////////////////
/*  Set the time and prohibt delete */ 
if (selectedColumn >= 3 && selectedColumn <= 4) {
  if (cellContent == ";" || cellContent == "ك" ) {
      var time = new Date();
      var hours = time.getHours();
      var minutes = time.getMinutes();
      var seconds = time.getSeconds();
      var milliseconds = time.getMilliseconds();
      var formattedTime = hours.toString().padStart(2, '0') + ':' +
                  minutes.toString().padStart(2, '0') + ':' +
                  seconds.toString().padStart(2, '0')
                  milliseconds.toString().padStart(3, '0');
    cellRange.setValue(formattedTime);
  }
   else {
    var password = Browser.inputBox("Ilegal Input:", Browser.Buttons.OK_CANCEL);
    if (password == "myPassword") {
      Browser.msgBox("Correct Password");
    } else {
      Browser.msgBox(`Incorrect password. Entry not allowed: `);
      sheet.getActiveCell().setValue("CHANGE");
      }
    }
  }
 


///////////////////////////////////////////////////
/*  Calculate the time difference */

if (selectedColumn == 4 && e.oldValue !== cell.getValue() && cell.getValue()!= ''){
  var result;

  var col_G = sheet.getRange("G" + selectedRow);
  var col_C = sheet.getRange("C" + selectedRow).getValue();
  var col_D = sheet.getRange("D" + selectedRow).getValue();
  var time1 = Utilities.formatDate(col_C, Session.getScriptTimeZone(), "HH:mm:ss");
  var time2 = Utilities.formatDate(col_D, Session.getScriptTimeZone(), "HH:mm:ss");
  var time1Parts = time1.split(":");
  var time2Parts = time2.split(":");
  var hours1 = parseInt(time1Parts[0], 10);
  var minutes1 = parseInt(time1Parts[1], 10);
  var hours2 = parseInt(time2Parts[0], 10);
  var minutes2 = parseInt(time2Parts[1], 10);
  var totalMinutes1 = (hours1 * 60) + minutes1;
  var totalMinutes2 = (hours2 * 60) + minutes2;
  var minutesDifference = totalMinutes2 - totalMinutes1;
  var hoursDiff = Math.floor(minutesDifference / 60);
  var minutesDiff = minutesDifference % 60;

  if(hoursDiff < 0){
    hoursDiff += 24;
    minutesDiff = minutesDiff * -1;
  }
  var res = Number(hoursDiff) + (Number(minutesDiff) / 100);
  
  var cell = sheet.getRange("F" + selectedRow);
  cell.setValue(res);

  if (cell.getValue() < 0.2) {
    col_G.setValue(10);
  } else if (cell.getValue() >= 3.20) {
    col_G.setValue(35);
  } else if (cell.getValue() - Math.floor(cell.getValue()) > 0.20) {
    col_G.setValue(10 * (Math.floor(cell.getValue()) + 1));
  } else {
    col_G.setValue(10 * Math.floor(cell.getValue()));
  }  
}
}
