function onEdit(e) {
  if (!e) return; // Check if event object is defined

var sheet = SpreadsheetApp.getActiveSheet();
var cell = e.range; // Current Cell
var cellRange = sheet.getActiveCell();
var cellContent = sheet.getActiveCell().getValue();
var selectedColumn = cellRange.getColumn();
var selectedRow = cellRange.getRow();



/*  Change the empty and do not change the changed */

  if(selectedColumn != 1  && selectedColumn != 5){
    if( selectedColumn == 12 || selectedColumn == 13 || selectedColumn == 6 || selectedColumn == 7 ){
      e.range.setValue(e.oldValue);
      e.source.toast("You cannot modify non-empty cells.");
    }
    else if((selectedColumn == 3 || selectedColumn == 4) && ( (cellContent == ';' || cellContent == 'ك') && (e.oldValue == '-') ) || ((cellContent == ';;' || cellContent == 'كك') && (e.oldValue == ';') || (e.oldValue == 'ك') )){
      e.range.setValue('-');
      e.source.toast("Time will be inserted");    
    }
    else if((selectedColumn == 3 || selectedColumn == 4) && ( (cellContent == ';' || cellContent == 'ك') && (e.oldValue != '-') )){
      var decimalTime = e.oldValue; // Replace with your decimal time value
      var totalSeconds = decimalTime * 86400; // 86400 seconds in a day
      var hours = Math.floor(totalSeconds / 3600); var minutes = Math.floor((totalSeconds % 3600) / 60); var seconds = Math.round(totalSeconds % 60);
      var timeString = hours.toString().padStart(2, '0') + ':' + minutes.toString().padStart(2, '0') + ':' + seconds.toString().padStart(2, '0');
      e.range.setValue(timeString);
      return;

    }
    else if (selectedColumn == 2 && (e.range.getValue() === '' ||e.range.getValue() === "" || e.range.getValue() === null) && e.oldValue === 'none') {
      e.source.toast("You can't delete this cell");    
      e.range.setValue('none');
      return;
    }

    else if ((selectedColumn == 14 || selectedColumn == 15) && e.oldValue ==  '-'  ){
      e.source.toast("Subscription will be inserted"); 
    }
    else{
      var newValue = e.value;
      var oldValue = e.oldValue;
      if (oldValue !== 'none' && newValue !== oldValue ) {
      e.range.setValue(e.oldValue);
      e.source.toast("You cannot modify non-empty cells.");
      if (selectedColumn == 14 || selectedColumn == 15|| selectedColumn == 3|| selectedColumn == 4){return;}}}}            



///////////////////////////////////////////////////

/*  Restart Button */

if (selectedColumn == 21 && selectedRow == 14 && cellContent =='Yes'){


  var limit1 = sheet.getRange("U" + 12).getValue();
      var limit11 = sheet.getRange("U" + 10).getValue();
      for(var j=2 ; j<= limit1; j++){
      if (sheet.getRange("C"+ j).getValue() != '-' && sheet.getRange("D"+j).getValue() == '-') {
      var name = sheet.getRange("B" + j).getValue();
      if (name != 'none'){
      var name2 = sheet.getRange("E" + j).getValue();
      var isfound = false;
      for (var i = 2; i <= (limit11 + 20) ; i++){
      if(isfound){ break;}
      var celltest = sheet.getRange("L"+ i );
      var celltest2 = sheet.getRange("M"+ i );
      var celltest3 = sheet.getRange("C"+ j );
      var celltest4 = sheet.getRange("E"+ j );
      if (celltest.getValue() == '-' && ( (celltest3.getValue()!= '###' ) || ( celltest3.getValue() == '###' && celltest4.getValue() != '' ) ) ){
        celltest.setValue(name) 
        if (name2 == '') { celltest2.setValue(0) }
        else {celltest2.setValue(name2)}
        isfound = true;
      }}}}}

    
      var sum = 0;
      for (var j=2 ; j<= limit1; j++){
          var cellValue = sheet.getRange("H" + j).getValue();
          var intValue = parseInt(cellValue, 10); // Parse as integer with base 10
          sum += intValue;
      }
      sum+= sheet.getRange("T" + 7).getValue(); sum+= sheet.getRange("T" + 9).getValue();
      sheet.getRange("U" + 15).setValue(sum);


      
  var limit1 = sheet.getRange("U" + 12).getValue();
  for(var i=2 ; i<= limit1; i++){
    sheet.getRange("B" + i).setValue('none');
    sheet.getRange("C" + i).setValue('-');
    sheet.getRange("D" + i).setValue('-');
    sheet.getRange("E" + i).setValue('');
    sheet.getRange("F" + i).setValue('');
    sheet.getRange("G" + i).setValue('');
    }


}


if (selectedColumn == 21 && selectedRow == 14 && cellContent =='Hide'){
  sheet.getRange("U" + 15).setValue('');
}


///////////////////////////////////////////////////



/*  Format of the names */

function setFontAndSize(range, fontFamily, fontSize) { range.setFontFamily(fontFamily).setFontSize(fontSize); }
var rangeB = sheet.getRange("B:B");
setFontAndSize(rangeB, "Calibri", 12);
var rangeL = sheet.getRange("L:L");
setFontAndSize(rangeL, "Calibri", 12);
var rangeN = sheet.getRange("N:N");
setFontAndSize(rangeN, "Calibri", 12);
var rangeO = sheet.getRange("O:O");
setFontAndSize(rangeO, "Calibri", 12);


///////////////////////////////////////////////////

/*  Caught The flee name if came back */ 
function setter( j ,variable, cellE_t, name) { 
var temp = j + 1;
var name2 = sheet.getRange("M" + temp).getValue();

if (name2 == 0){
    var res = 35;
    variable.setValue(  ` عليه ${res} - ${name}` ); 
    cellE_t.setValue(res);
  }
else {
  var intValue = parseInt(name2, 10); 
  var res = 35 + intValue;
  variable.setValue(  `  عليه ${res} - ${name}` ); 
  cellE_t.setValue(res);
  }}

///////////////////////////////////////////////////

/*  Clear if name came back */

if  (selectedColumn == 2 && cell.getValue()!= '' && cell.getValue()!= 'none'){
var substringg = "عليه";
var rangeP = sheet.getRange("L:L");
var valuesP = rangeP.getValues();
var rangeO = sheet.getRange("O:O");
var valuesO = rangeO.getValues();
var rangeN = sheet.getRange("N:N");
var valuesN = rangeN.getValues();
var cellB = sheet.getRange("B" + selectedRow).getValue();
var cellB_t = sheet.getRange("B" + selectedRow)
var cellE_t = sheet.getRange("E" + selectedRow)


var limit2 = sheet.getRange("U" + 10).getValue();
for (var i = 1; i <= limit2 ; i++) {
if ( cellB == valuesP[i][0] && cellB != 'none' ) {
  if (valuesP[i][0].indexOf(substringg) !== -1){
    var name = valuesP[i][0].split(" - ")[1];
  }
  else{
    var name = valuesP[i][0];
  }
  // var flag = sheet.getRange("C" + (i+1)).getValue() === '###' ? true : false;
  setter(i , cellB_t, cellE_t, name);
  var temp = i+1
  var cell = sheet.getRange("L" + temp);
  cell.setValue('-');
  var cell = sheet.getRange("M" + temp); 
  cell.setValue('-');
  // var cell = sheet.getRange("N" + temp);
  // cell.setValue('here');
  break;
}
}
}
///////////////////////////////////////////////////

/*  Subscriptions */

if ((selectedColumn == 14 || selectedColumn == 15) && cell.getValue() !== '-') {
var restOfDays = sheet.getRange("P" + selectedRow);
var other = sheet.getRange("R" + selectedRow);
if (selectedColumn == 14 && cell.getValue() !== '' && selectedRow != 1) {
  restOfDays.setValue(0);
  other.setValue(sheet.getRange("S" + 2).getValue());
  var fees = sheet.getRange("Q" + selectedRow);
  fees.setValue(100);
  var other = sheet.getRange("O" + selectedRow);
  other.setValue('//');
  sheet.getRange("T" + 7).setValue( sheet.getRange("T" + 7).getValue() + 100  );
  
}
else if (selectedColumn == 15 && cell.getValue() !== '' && selectedRow != 1){
  restOfDays.setValue(0);
  other.setValue(sheet.getRange("S" + 2).getValue());
  var fees = sheet.getRange("Q" + selectedRow);
  fees.setValue(800);
  var other = sheet.getRange("N" + selectedRow);
  other.setValue('//');
  sheet.getRange("T" + 9).setValue( sheet.getRange("T" + 9).getValue() + 800  );
}
}

if (selectedColumn == 21 && selectedRow == 2){
if (cellContent == 'Yes'){
    
  var limit = sheet.getRange("U" + 8).getValue();
   for(var i=2 ; i <= limit ; i++ ){

    
    var cell_O = sheet.getRange("O" + (i));
    var cell_N = sheet.getRange("N" + (i));

    if (cell_O.getValue() != '-' || cell_N.getValue() != '-' ){
    var check = sheet.getRange("N" + (i)).getValue();
    var restOfDays = sheet.getRange("P" + i);
    var startDate = new Date (sheet.getRange("R" + i).getValue());
    var endDate = new Date (sheet.getRange("S" + 2).getValue());
    var lifetime = Math.ceil (parseInt((endDate - startDate) / 1000 / 60 / 60 / 24));
    
    restOfDays.setValue(lifetime);
    sheet.getRange("T" + 7).setValue( 0 );
    sheet.getRange("T" + 9).setValue( 0 );

    if (check == '//' && lifetime>30){
      // cell_O.setValue(cell_O.getValue() + ' -- Expired');
      cell_O.setValue('-');
      cell_N.setValue('-');
      restOfDays.setValue('');
      sheet.getRange("Q" + i).setValue('')
      sheet.getRange("R" + i).setValue('')
      return; }
    else if (check !== '//' && lifetime>7){
      cell_O.setValue('-');
      cell_N.setValue('-');
      restOfDays.setValue('');
      sheet.getRange("Q" + i).setValue('')
      sheet.getRange("R" + i).setValue('')
      // cell_N.setValue(cell_N.getValue() + ' -- Expired');
      return;
      }
    }
       
      }}}


///////////////////////////////////////////////////
/*  Set the time and prohibt delete */ 


if (selectedColumn >= 3 && selectedColumn <= 4) {
  if (cellContent == ";" || cellContent == "ك" ) {

    var rangeO = sheet.getRange("O:O");
    var valuesO = rangeO.getValues();
    var rangeN = sheet.getRange("N:N");
    var valuesN = rangeN.getValues();
    var limit = sheet.getRange("U" + 8).getValue();
    var cell_b = sheet.getRange("B" + selectedRow).getValue();
    
    if (selectedColumn == 3){
    for(var i=1; i<=limit ; i++){
      if (( cell_b == valuesO[i][0] && cell_b != 'none') || (cell_b == valuesN[i][0] && cell_b != 'none')) {
        sheet.getRange("C" + selectedRow).setValue('###');
        sheet.getRange("D" + selectedRow).setValue('###');
        sheet.getRange("F" + selectedRow).setValue(0);
        sheet.getRange("G" + selectedRow).setValue(0);
        return;
      }
    }
    }

    var time = new Date();
    var formattedTime = Utilities.formatDate(time, Session.getScriptTimeZone(), "HH:mm:ss");
    cellRange.setValue(formattedTime);
  }

  /* Calculate the Cost */
  if (selectedColumn == 4 ){
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
    if (minutesDiff == 0) 
    var res = Number(hoursDiff) + (Number( minutesDiff) / 100);
    else
    var res = Number(hoursDiff) + (Number( 60 - minutesDiff) / 100);
  }
  else{
    var res = Number(hoursDiff) + (Number( minutesDiff) / 100);
  }
  
  
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
  }}

  } }