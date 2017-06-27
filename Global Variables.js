//Global Variables
var ss = SpreadsheetApp.getActiveSpreadsheet();

//Create macros menu on open
function onOpen(e) {
      var menuEntries = [];

                      menuEntries.push({ name: "Email-Finding Data", functionName: "getEmailFindingData" });   
                      menuEntries.push(null); // line separator
                      menuEntries.push({name: "Update Finisher Pix", functionName: "finisherPixInfo"});
                      
 ss.addMenu("Macros", menuEntries);
}