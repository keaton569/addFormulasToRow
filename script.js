/** @OnlyCurrentDoc */

function macro_update_close_filter() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  var criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['TRUE'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(1, criteria);
};




function onEdit() {
  var Sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Responses");
  var row = Sheet1.getLastRow();
  var formulas = ['=IF($A'+row+'>0,VLOOKUP($B'+row+',\'EF DATA\'!$B$1:$G$5000,2,false),"")', 
                  '=IF($A'+row+'>0,VLOOKUP($B'+row+',\'EF DATA\'!$B$1:$G$5000,5,false),"")', 
                  '=IF($A'+row+'>0,VLOOKUP($Z'+row+',\'EF AM #s\'!$A:$B,2,false),"")',
                 '=IF($A'+row+'>0,if(VLOOKUP($B'+row+',\'EF DATA\'!$B$1:$G$5000,3,false)>0,VLOOKUP($B'+row+',\'EF DATA\'!$B$1:$G$5000,3,false),"missing"),"")',
                 '=IF($A'+row+'>0,if(VLOOKUP($B'+row+',\'EF DATA\'!$B$1:$G$5000,4,false)>0,VLOOKUP($B'+row+',\'EF DATA\'!$B$1:$G$5000,4,false),"missing"),"")'];
  
  Sheet1.getRange(Sheet1.getLastRow(), 25).setFormula(formulas[0]);
  Sheet1.getRange(Sheet1.getLastRow(), 26).setFormula(formulas[1]);
  Sheet1.getRange(Sheet1.getLastRow(), 27).setFormula(formulas[2]);
  Sheet1.getRange(Sheet1.getLastRow(), 28).setFormula(formulas[3]);
  Sheet1.getRange(Sheet1.getLastRow(), 29).setFormula(formulas[4]);
}
