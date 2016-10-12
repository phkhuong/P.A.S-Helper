/*
P.A.S Helper v1.01
What's new:
- Improved phonenumber function's chance of find company email.
*/

function phonenumber(company,city) {
  if(company === "" || city === "")
    return "";
  else{
  var company_array = company.split(" ");
  var city_array = city.split(" ");
  var url ="https://www.google.com/search?q=";
  for(i=0, len = company_array.length; i< len; i++)
    url = url + company_array[i] + "+";
  for(i=0, len = city_array.length; i< len; i++){
    if(i<city_array.length-1)
      url = url + city_array[i] + "+";
    else
      url = url + city_array[i] +"+phone+number";}
  var response = UrlFetchApp.fetch(url);
  Utilities.sleep(1000);
  var string = response.getContentText();
  if(string.search('<span class="_m3b">') !== -1){
    var a = string.indexOf('<span class="_m3b">');
    var cut = string.substring(a, string.length);
    var output = cut.substring(19,cut.indexOf("</span>"));
  }
  else if(string.search('</span> &middot; ') !== -1){
    var a = string.indexOf('</span> &middot; ');
    var cut = string.substring(a, string.length);
    var output = cut.substring(16,cut.indexOf("</div>"));
  }
  else
    return "";
  return output;
  //Logger.log(output);
  }
}



function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Report Helper', functionName: 'report'},
    {name: 'Duplicate Detect', functionName: 'duplicate_finder'}
  ];
  spreadsheet.addMenu('PAS Helper', menuItems);
}

function report() {
  var activespreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var activesheet = SpreadsheetApp.getActiveSheet();
  var spreadsheetname = activespreadsheet.getName();
  var data = activesheet.getDataRange().getValues();
  var lrow = activesheet.getLastRow();
  var total = lrow - 1;
  var duplicate = 0;
  var dupnm = 0;
  var nomail = 0;
  var dsource = 0;
  var workcity = 0;
  var workstate = 0;
  var phonenumber = 0;
  var facebookurl = 0;
  var twitterurl = 0;
  var github = 0;
  var facebookemail = 0;
  for(i=0, len = activesheet.getLastColumn(); i<len; i++){
    if(data[0][i] === "Tag \(dsource\)" || data[0][i] === "Tag dsource" || data[0][i] === "Tag Dsource" || data[0][i] === "Tag \(Dsource\)" || data[0][i] === "Tag \(Female\)" || data[0][i] === "Tag Female" )
      for(j=1; j<lrow; j++){
        if(data[j][i] !== "")
          dsource++;
      }
    else if(data[0][i] === "Facebook Profile")
      for(j=1; j<lrow; j++){
        if(data[j][i] !== "")
          facebookurl++;
      }
    else if(data[0][i] === "Twitter URL")
      for(j=1; j<lrow; j++){
        if(data[j][i] !== "")
          twitterurl++;
      }
    else if(data[0][i] === "Github Profile")
      for(j=1; j<lrow; j++){
        if(data[j][i] !== "")
          github++;
      }
    else if(data[0][i] === "Work City/Town")
      for(j=1; j<lrow; j++){
        if(data[j][i] !== "")
          workcity++;
      }
    else if(data[0][i] === "Work State")
      for(j=1; j<lrow; j++){
        if(data[j][i] !== "")
          workstate++;
      }
    else if(data[0][i] === "Facebook Email")
      for(j=1; j<lrow; j++){
        if(data[j][i] !== "")
          facebookemail++;
      }
    else if(data[0][i] === "Home Email"){
      for(j=1; j<lrow; j++){
        if(data[j][i] === "" && data[j][i+1] === ""){
          nomail++
          for(l=0; l<len; l++){
            if( l===0 || data[0][l] === "NOTE" || data[0][l] === "Note" || data[0][l] === "note"){
              if(data[j][l].indexOf("Dup") !== -1 || data[j][l].indexOf("dup") !== -1 || data[j][l].indexOf("Duplicate") !== -1 || data[j][l].indexOf("duplicate") !== -1 ){
                dupnm++;
              }
            }
          }
        }
      }
    }
    else if(data[0][i] === "Mobile Phone")
      for(j=1; j<lrow; j++){
        if(data[j][i] !== "")
          phonenumber++;
      }
    else if(data[0][i] === "Home Phone")
      for(j=1; j<lrow; j++){
        if(data[j][i] !== "")
          phonenumber++;
      }
    else if(data[0][i] === "Work Phone")
      for(j=1; j<lrow; j++){
        if(data[j][i] !== "")
          phonenumber++;
      }
    else if( i===0 || data[0][i] === "NOTE" || data[0][i] === "Note" || data[0][i] === "note")
      for(j=1; j<lrow; j++){
        if(data[j][i].indexOf("Dup") !== -1 || data[j][i].indexOf("dup") !== -1 || data[j][i].indexOf("Duplicate") !== -1 || data[j][i].indexOf("duplicate") !== -1 )
          duplicate++;
      }
  }

  var anchor = activesheet.getLastColumn() +2;
  activesheet.getRange(4, anchor).setValue(spreadsheetname).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, anchor+1).setValue("").setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, anchor+2).setValue("").setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, anchor+3).setValue("").setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, anchor+4).setValue("").setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, anchor+5).setValue(total).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, anchor+6).setValue(duplicate).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, anchor+7).setValue(total-nomail-(duplicate-dupnm)).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, anchor+8).setValue(nomail - dupnm).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, anchor+9).setValue(phonenumber).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, anchor+10).setValue(dsource).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, anchor+11).setValue(workcity).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, anchor+12).setValue(workstate).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, anchor+13).setValue(facebookurl).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, anchor+14).setValue(twitterurl).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, anchor+15).setValue(github).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, anchor+16).setValue(facebookemail).setBorder(true, true, true, true, true, true);
  
  activesheet.getRange(2,anchor,2).merge().setValue("List name").setBackground("#c9daf8").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
  activesheet.getRange(2,anchor+1,2).merge().setValue("Sent by? \(Sourcer name\)").setBackground("#c9daf8").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
  activesheet.getRange(2,anchor+2,2).merge().setValue("Sent to? N or T or WE name").setBackground("#c9daf8").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
  activesheet.getRange(2,anchor+3,2).merge().setValue("QC? x").setBackground("#c9daf8").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
  activesheet.getRange(2,anchor+4,2).merge().setValue("Format check").setBackground("#c9daf8").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
  activesheet.getRange(2,anchor+5,2).merge().setValue("Total").setBackground("#c9daf8").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
  activesheet.getRange(2,anchor+6,2).merge().setValue("Duplicate").setBackground("#c9daf8").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
  activesheet.getRange(2,anchor+7,1,2).merge().setValue("Email").setBackground("#c9daf8").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
  activesheet.getRange(3,anchor+7).setValue("Yes").setBackground("#c9daf8").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
  activesheet.getRange(3,anchor+8).setValue("No").setBackground("#c9daf8").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
  activesheet.getRange(2,anchor+9,2).merge().setValue("Phone").setBackground("#c9daf8").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
  activesheet.getRange(2,anchor+10,2).merge().setValue("Dsource").setBackground("#c9daf8").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
  activesheet.getRange(2,anchor+11,2).merge().setValue("Work City").setBackground("#c9daf8").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
  activesheet.getRange(2,anchor+12,2).merge().setValue("Work State").setBackground("#c9daf8").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
  activesheet.getRange(2,anchor+13,2).merge().setValue("FB URL").setBackground("#c9daf8").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
  activesheet.getRange(2,anchor+14,2).merge().setValue("Twitter URL").setBackground("#c9daf8").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
  activesheet.getRange(2,anchor+15,2).merge().setValue("Github URL").setBackground("#c9daf8").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
  activesheet.getRange(2,anchor+16,2).merge().setValue("FB Email").setBackground("#c9daf8").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
  
  activesheet.getRange(2,anchor,3,1).copyTo(activesheet.getRange(6, anchor));
  activesheet.getRange(2,anchor+5,3,13).copyTo(activesheet.getRange(6, anchor+1));

}

function duplicate_finder(){
  var activespreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var activesheet = SpreadsheetApp.getActiveSheet();
  var lastrow = activesheet.getLastRow();
  var lastcolumn = activesheet.getLastColumn();
    // Prompt the user for a row number.
  var selectedcolumn = Browser.inputBox('Duplicate Detect','Please enter the column number to use:',
      Browser.Buttons.OK_CANCEL);
  if (selectedcolumn == 'cancel') {
    return;
  }
  var columnNumber = Number(selectedcolumn);
  if (isNaN(columnNumber) || columnNumber < 1 ||
      columnNumber > lastcolumn) {
    Browser.msgBox('Error',
        Utilities.formatString('Column "%s" is not valid.', selectedcolumn),
        Browser.Buttons.OK);
    return;
  }
  var data = activesheet.getRange(2, columnNumber, lastrow).getValues();
  var formula = activesheet.getRange(2, columnNumber, lastrow).getFormulas();
  Logger.log(formula[2]);
  for(i=0;i<lastrow-1;i++){
    for(j=i+1;j<lastrow;j++){
      if(String(data[i]) == String(data[j])){
        if(String(formula[i]) == String(formula[j])){
          activesheet.getRange(i+2, columnNumber).setBackground("Orange");
          activesheet.getRange(j+2, columnNumber).setBackground("Orange");
        }
        else{
          activesheet.getRange(i+2, columnNumber).setBackground("Yellow");
          activesheet.getRange(j+2, columnNumber).setBackground("Yellow");
        }       
      }
    }
  }
}
