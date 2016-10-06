function phonenumber(company,city) {
  //var sheet = SpreadsheetApp.getActiveSheet();
  //var data = sheet.getDataRange().getValues();
  //var company = data[0][0];
  //var city = data[0][1];
  var company_array = company.split(" ");
  var city_array = city.split(" ");
  var url ="https://www.google.com.vn/search?q=";
  for(i=0, len = company_array.length; i< len; i++)
    url = url + company_array[i] + "+";
  for(i=0, len = city_array.length; i< len; i++){
    if(i<city_array.length-1)
      url = url + city_array[i] + "+";
    else
      url = url + city_array[i] +"+phone+number";}
 var response = UrlFetchApp.fetch(url);
 var string = response.getContentText();
  if(string.search('<span class="_m3b">') === -1)
    return "";
 var a = string.indexOf('<span class="_m3b">');
 var cut = string.substring(a, string.length);
  var output = cut.substring(19,cut.indexOf("</span>"));
  return output;
  //sheet.appendRow([final]);
  //Logger.log(output);
}



function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Report Helper', functionName: 'report'}
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
          for(l=0; l<len; l++){
            if( l===0 || data[0][l] === "NOTE" || data[0][l] === "Note" || data[0][l] === "note"){
              if(data[j][l].indexOf("Dup") === -1 || data[j][l].indexOf("dup") === -1 || data[j][l].indexOf("Duplicate") === -1 || data[j][l].indexOf("duplicate") === -1 )
                nomail++;
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

  var archor = activesheet.getLastColumn() +2;
  activesheet.getRange(4, archor).setValue(spreadsheetname).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, archor+1).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, archor+2).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, archor+3).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, archor+4).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, archor+5).setValue(total).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, archor+6).setValue(duplicate).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, archor+7).setValue(total-nomail).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, archor+8).setValue(nomail).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, archor+9).setValue(phonenumber).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, archor+10).setValue(dsource).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, archor+11).setValue(workcity).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, archor+12).setValue(workstate).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, archor+13).setValue(facebookurl).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, archor+14).setValue(twitterurl).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, archor+15).setValue(github).setBorder(true, true, true, true, true, true);
  activesheet.getRange(4, archor+16).setValue(facebookemail).setBorder(true, true, true, true, true, true);
  
  activesheet.getRange(2,archor,2).merge().setValue("List name").setBackground("#c9daf8").setHorizontalAlignment("center");
  activesheet.getRange(2,archor+1,2).merge().setValue("Sent by? \(Sourcer name\)").setBackground("#c9daf8").setHorizontalAlignment("center");
  activesheet.getRange(2,archor+2,2).merge().setValue("Sent to? N or T or WE name").setBackground("#c9daf8").setHorizontalAlignment("center");
  activesheet.getRange(2,archor+3,2).merge().setValue("QC? x").setBackground("#c9daf8").setHorizontalAlignment("center");
  activesheet.getRange(2,archor+4,2).merge().setValue("Format check").setBackground("#c9daf8").setHorizontalAlignment("center");
  activesheet.getRange(2,archor+5,2).merge().setValue("Total").setBackground("#c9daf8").setHorizontalAlignment("center");
  activesheet.getRange(2,archor+6,2).merge().setValue("Duplicate").setBackground("#c9daf8").setHorizontalAlignment("center");
  activesheet.getRange(2,archor+7,1,2).merge().setValue("Email").setBackground("#c9daf8").setHorizontalAlignment("center");
  activesheet.getRange(3,archor+7).setValue("Yes").setBackground("#c9daf8").setHorizontalAlignment("center");
  activesheet.getRange(3,archor+8).setValue("No").setBackground("#c9daf8").setHorizontalAlignment("center");
  activesheet.getRange(2,archor+9,2).merge().setValue("Phone").setBackground("#c9daf8").setHorizontalAlignment("center");
  activesheet.getRange(2,archor+10,2).merge().setValue("Dsource").setBackground("#c9daf8").setHorizontalAlignment("center");
  activesheet.getRange(2,archor+11,2).merge().setValue("Work City").setBackground("#c9daf8").setHorizontalAlignment("center");
  activesheet.getRange(2,archor+12,2).merge().setValue("Work State").setBackground("#c9daf8").setHorizontalAlignment("center");
  activesheet.getRange(2,archor+13,2).merge().setValue("FB URL").setBackground("#c9daf8").setHorizontalAlignment("center");
  activesheet.getRange(2,archor+14,2).merge().setValue("Twitter URL").setBackground("#c9daf8").setHorizontalAlignment("center");
  activesheet.getRange(2,archor+15,2).merge().setValue("Github URL").setBackground("#c9daf8").setHorizontalAlignment("center");
  activesheet.getRange(2,archor+16,2).merge().setValue("FB Email").setBackground("#c9daf8").setHorizontalAlignment("center");
  
  activesheet.getRange(2,archor,3,1).copyTo(activesheet.getRange(6, archor));
  activesheet.getRange(2,archor+5,3,13).copyTo(activesheet.getRange(6, archor+1));

}
