Finisherpixinfo
function finisherPixInfo() {


var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();
var ui = SpreadsheetApp.getUi();


// checks to see if there is a url that needs to be input. But, will close the program if there isn't any input. 
var question = ui.prompt(
      'FinisherPix Info Update',
      'Please enter FinisherPix url snippet',
      ui.ButtonSet.OK_CANCEL);
var fpi = question.getResponseText();      
Logger.log("url snippet entered = " + fpi);

if (fpi == '') {
      ui.alert('No FinisherPix url defined. Please try again. Script execution canceled.');
      throw new Error('No FinisherPix url defined. Please try again. Script execution canceled.');
    }


//sets up an array with all of the sheet names, and prepares the outermost for loop; which cycles the 'splicer' through each sheet. 
var sheetnames = ["M18-24", "M25-29", "M30-34", "M35-39", "M40-44", "M45-49", "M50-54", "M55-59", "M60-64", "M65-69",
  "F18-24", "F25-29", "F30-34", "F35-39", "F40-44", "F45-49", "F50-54", "F55-59", "F60-64", "F65-69",];
var s = sheetnames.length
Logger.log(s);

for (i = 0 ; i < sheetnames.length; i++){
Logger.log("current sheet " + sheetnames[i]);
var sheet = ss.getSheetByName(sheetnames[i]);
Logger.log(sheet);

if (sheet == null) {
}

else {



//start of the splicer. sets up the array that will be used to push info to the sheets. initializes the sheet that data will be called from, 
//and sorts it in preparation for the splice
var finisherpix = [];
var range = sheet.getRange(1,1, ss.getLastRow(), ss.getLastColumn());
sortsheet = range.sort(1);




//For loop needs to accomplish two things, collect the bib number, and then push the full finisherpix tring to the array. 
//Once the loop is finished, the script will push the array to each corresponding cell. 
  var lastRow = sheet.getLastRow();
Logger.log("sheet is " + lastRow + " rows deep");
  for (r=1; r <= lastRow; r++) {
     var r1 = r + 1

//set the range below to reflect the 'bib' range. 
var thcell = "O" + r1;
Logger.log(thcell);
var bibs = sheet.getRange(thcell).getValues().toString();
//Logger.log("r = " + r);
//Logger.log("r1 = " + r1);
//Logger.log("bibs = " + bibs);
    
finisherpix.push(['=JOIN("", {"http://www.finisherpix.com/photos/my-photos/currency/USD/pctrl/Photos/paction/search/pevent/' + fpi + '/pbib/"},' + bibs + ', {".html"})']);


}
//pushes the concatenated string to the destination shee. 
var destrange = sheet.getRange(2,11,sheet.getLastRow(),1);
destrange.setFormulas(finisherpix);
}
//closes out app with a simple note that the script is complete, and returns the sheet variable to match the global scope... I think :)
}
ui.alert('script complete');
var sheet = ss.getActiveSheet();
}