function idPopUp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi()
  let attempt = ""
  var studentIDs = ['52K13J','68C62S','45G31R','78G05U','19R04Q','51P22S','58D85H',
  '20C28P','40W01W','36Z47A','11K99U','13L02H','47U55R','52C92X','12Y15Z','78X79S','33V38G','93F74N','78Y37M','54P63C','23F61T','72G39P','69P49D','17Y94X','47C02F']
  while (!(studentIDs.includes(attempt))){
    attempt = ui.prompt("Enter Student ID").getResponseText()
  }
  insertUserSheet()
  var finishButton = newSheet.insertImage("https://upload.wikimedia.org/wikipedia/commons/thumb/c/c0/OOjs_UI_icon_cancel.svg/20px-OOjs_UI_icon_cancel.svg.png", 1, 3) 
  finishButton.assignScript("finishSession").setHeight(20).setWidth(20)
  let studentID = attempt.toString()
  newSheet.getRange(2,1).setValue(studentID)
  newSheet.setName(studentID)
}
function insertUserSheet(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.insertSheet("New Sheet");
      newSheet = ss.getSheetByName("New Sheet")
      newSheet.getRange(2,2).setFormula("=IFNA(VLOOKUP(A2,IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1vE5K8VghNGc4zUb-PTDjudg1maEpeeZihRpkhA2uFEc/edit?usp=sharing\",\"Student_IDs!A1:B26\"), 2, FALSE))")
      newSheet.getRange(2,3).setFormula("=IFNA(index(FILTER(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1vE5K8VghNGc4zUb-PTDjudg1maEpeeZihRpkhA2uFEc/edit?usp=sharing\",\"Sample_Data!C2:C\"),IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1vE5K8VghNGc4zUb-PTDjudg1maEpeeZihRpkhA2uFEc/edit?usp=sharing\",\"Sample_Data!B2:B\")=B2),countif(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1vE5K8VghNGc4zUb-PTDjudg1maEpeeZihRpkhA2uFEc/edit?usp=sharing\",\"Sample_Data!B2:B\"),B2)), \"\")")
      newSheet.getRange(2,4).setFormula("=IF(B2 <> \"\",TO_DATE(DMAX(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1vE5K8VghNGc4zUb-PTDjudg1maEpeeZihRpkhA2uFEc/edit?usp=sharing\",\"Sample_Data!A1:G358\"), \"Payment Date\", $B$1:$B$2)), \"\")")
      newSheet.getRange(2,5).setFormula("=IF(B2 <> \"\",IFNA(QUERY(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1vE5K8VghNGc4zUb-PTDjudg1maEpeeZihRpkhA2uFEc/edit?usp=sharing\",\"Sample_Data!A1:G358\"), \"SELECT Col5 WHERE Col2 Contains '\"&B2&\"' ORDER BY Col7 DESC LIMIT 1\", 0), \"\"))")
      newSheet.getRange(2,6).setFormula("=IFNA(index(FILTER(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1vE5K8VghNGc4zUb-PTDjudg1maEpeeZihRpkhA2uFEc/edit?usp=sharing\",\"Sample_Data!D2:D\"),IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1vE5K8VghNGc4zUb-PTDjudg1maEpeeZihRpkhA2uFEc/edit?usp=sharing\",\"Sample_Data!B2:B\")=B2),countif(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1vE5K8VghNGc4zUb-PTDjudg1maEpeeZihRpkhA2uFEc/edit?usp=sharing\",\"Sample_Data!B2:B\"),B2)),\"\")")
      newSheet.getRange(2,7).setFormula("=IF(B2<> 0, F2-C2, \"\")")
      newSheet.getRange(2,8).setFormula("=IF(B2 <> \"\",TO_DATE((DMAX(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1vE5K8VghNGc4zUb-PTDjudg1maEpeeZihRpkhA2uFEc/edit?usp=sharing\",\"Sample_Data!A1:G358\"), \"Date\", $B$1:$B$2))+((G2)+1)*7), \"\")")

      newSheet.getRange(6, 1).setFormula("=IF(B2 <> \"\", QUERY(IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1vE5K8VghNGc4zUb-PTDjudg1maEpeeZihRpkhA2uFEc/edit?usp=sharing\",\"Sample_Data!A1:G358\"), \"Select Col2, Col1, Col3, Col4, Col5, Col6, Col7 where LOWER(Col2) = LOWER(\'\"&B2&\"\')\"))")
      newSheet.getRange(3,2).setValue("Click ⨂ on left to finish session")
 formatSheet()
}

function formatSheet(){
  let titles = [["Student ID", "Student","Amount of Lessons", "Last Payment", "Payment Amount", "No. of Lessons", "Lessons Left", "Next Payment Date"]]
  newSheet.getRange("A1:H1").setValues(titles)
  newSheet.getRange("A1:H1").setBackground("#1c4587").setFontColor("#ffffff")
  newSheet.getRange("A6:H6").setBackground("#1c4587").setFontColor("#ffffff")
}

function finishSession(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi()
  var response = ui.alert("Do you want to close the session?",ui.ButtonSet.YES_NO );
  if(response == ui.Button.YES) {
      ss.deleteActiveSheet();
  }
}
