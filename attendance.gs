
function onOpen(){
  menu()

}

function menu(){
 var menu =  SpreadsheetApp.getUi().createMenu("Managed Attendance Data");
 menu.addItem("Save Attendance Data","sendData").addToUi()
}

function sendData() {
  var ss = SpreadsheetApp.getActive();

  //sheets
  var dashBoard = ss.getSheetByName("Dashboard");
  var tab1 =ss.getSheetByName("Tier 1 (69%<)")
  var tab2 =ss.getSheetByName("Tier 2(70-89%)")
  var tab3 =ss.getSheetByName("Tier 3(90%>)")

  tab3.activate();

  //last rows
  var tab1Lr =ss.getSheetByName("Tier 1 (69%<)").getLastRow();
  var tab2Lr =ss.getSheetByName("Tier 2(70-89%)").getLastRow();
  var tab3Lr =ss.getSheetByName("Tier 3(90%>)").getLastRow();
  
  //get sheet data
  var tier1 = ss.getSheetByName("Tier 1 (69%<)").getRange(1,1,tab1Lr,5).getValues();
  var tier2 = ss.getSheetByName("Tier 2(70-89%)").getRange(1,1,tab2Lr,5).getValues();
  var tier3 = ss.getSheetByName("Tier 3(90%>)").getRange(1,1,tab3Lr,5).getValues();

  //make new worksheets 
  var newTier1 = ss.insertSheet().setName("Tier 1 "+tab1.getRange("G3").getValue().toString());
  var newTier2 = ss.insertSheet().setName("Tier 2 "+tab2.getRange("G3").getValue().toString());
  var newTier3 = ss.insertSheet().setName("Tier 3 "+tab3.getRange("G3").getValue().toString());

  //send data to new tab
  newTier1.getRange(1,1,tab1Lr,5).setValues(tier1);
  newTier2.getRange(1,1,tab2Lr,5).setValues(tier2);
  newTier3.getRange(1,1,tab3Lr,5).setValues(tier3);

   //clear data 
  tab1.getRange(2,1,tab1Lr,).clearContent();
  tab2.getRange(2,1,tab2Lr,5).clearContent();
  tab3.getRange(2,1,tab3Lr,5).clearContent(); 
  dashBoard.getRange(2,1,dashBoard.getLastRow(),3).clearContent();

  //reset formulas
  tab1.getRange(2,1,1,1).setFormula("=filter(Dashboard!A2:E,Dashboard!E2:E<=69)")
  tab2.getRange(2,1,1,1).setFormula("=filter(Dashboard!A2:E,Dashboard!E2:E>=70,Dashboard!E2:E<90)")
  tab3.getRange(2,1,1,1).setFormula("=filter(Dashboard!A2:E,Dashboard!E2:E>90)")

//hide new sheets
  newTier1.hideSheet();
  newTier2.hideSheet();
  newTier3.hideSheet();
}


