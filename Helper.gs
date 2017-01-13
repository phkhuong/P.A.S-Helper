/*



CÁC BẠN KHI DÙNG SCRIPT GĂP LỖI GÌ THÌ REPORT CHO MÌNH ĐỂ GIÚP MÌNH HOÀN THIỆN SCRIPT NHÉ !!!

UPDATED
- Create working sheet sẽ tự động kết hợp First Name, Last Name và LinkedIn Profile để tạo cột Full Name/Link nếu Sheet gốc không có cột đấy (tốc độ thực thi của  Create working sheet sẽ chậm hơn khá nhiều nếu rơi vào trường hợp naỳ).


CÁC BAN LƯU Ý:
+ TRUƠC KHI DÙNG REPORT HELPER CÁC BẠN NHỚ TURN OFF HẾT FILTER ĐỂ KHÔNG BỊ LỖI, VÀ NHỚ XEM LẠI TIÊU ĐỀ CÁC CỘT XEM CÓ ĐÚNG KHÔNG.
+ TRÁNH KÉO HÀM phonenumber() QUÁ DÒNG CUỐI CÙNG CỦA LIST VÌ CÓ THỂ LÀM REPORT HELPER TÍNH SAI TOTAL VÀ NO MAIL.
+ CÁCH TÌM WORK PHONE KHI BỊ #ERROR: Xóa ô bị Error, sau đó nhấn Ctrl + Z để undo lại, hàm phonenumber sẽ tự động load lại, làm lại cho đên khi load được Work Phone.
+ Sau khi tim được hết Work Phone, COPY NGUYEN CỘT WORK PHONE SAU ĐÓ BẤM CHUỘT PHẢI CHỌN PASTE SPECIAL/PASTE VALUES ĐỂ XÓA CÔNG THỨC HÀM (ĐỂ TRÁNH TRƯỜNG HỢP HÔM SAU MỞ LẠI LIST THÌ CÔNG THỨC SẼ LOAD LẠI NÊN CÓ THỂ BỊ ERROR)
+ COPY cả phần hướng dẫn này cũng không sao nên các bạn cứ nhấn Ctrl+A để copy nhé



P.A.S Helper v1.02d
Contact: khuong.pham@vsource.io

------------------------------------------------
What's new:
- Create working sheet sẽ tự động kết hợp First Name, Last Name và LinkedIn Profile để tạo cột Full Name/Link nếu Sheet gốc không có cột đấy (tốc độ thực thi của script sẽ chậm hơn khá nhiều nếu rơi vào trường hợp naỳ).
- Hàm phonenumber() không còn báo #ERROR khi Work city hoặc Company có ký tự đặc biệt (~!@#$%&*(:".,[]{}_-=+<>/\|';:) 
- Sửa lỗi  hàm phonenumber() làm google sheet count sai số lượng khi chọn cột work phone.
- Updated "Total" counting formula
- Create Working Sheet đã có thể dùng cho các list bị trống vài dòng đầu tiên (trống <10 dòng)

________________________________________________
Hướng dẫn sử dụng:
- Copy tất cả dòng script tại đây ( Ctrl + A -> Ctrl + C )
- Vào list, chọn Tools/Script Editor trên menu bar
- Xóa tất cả nội dung có sẵn trong Script Editor, paste script đã copy vào, nhấn Ctrl+S, chọn OK để save script lại
- Refresh list, sau đó bạn sẽ thấy xuất hiện mục PAS Helper trên menu bar 
 + Chọn PAS Helper/Create working sheet (chọn Allow nếu xuất hiện bảng hỏi) sẽ tạo sheet mới và tự động copy các cột cần thiết để làm việc
 + Chọn PAS Helper/Duplication Detector (chọn Allow nếu xuất hiện bảng hỏi), điền số thứ tự cột muốn tìm duplicate (cột Full Name), các ô có tên trùng sẽ tô màu vàng, ô có linkedin trùng sẽ tô màu cam
 + Chọn PAS Helper/Report Helper sẽ tính tất cả các dữ liệu cần thiết để report và xuất ra phía ngoài cùng bên phải của sheet
- Nhập =phonenumber(Ô chứa tên company, Ô chứa work city) để tìm work phone. 
- Sau khi tim được Work Phone, COPY NGUYEN CỘT WORK PHONE SAU ĐÓ BẤM CHUỘT PHẢI CHỌN PASTE SPECIAL/PASTE VALUES ĐỂ XÓA CÔNG THỨC HÀM (ĐỂ TRÁNH TRƯỜNG HỢP HÔM SAU MỞ LẠI LIST THÌ CÔNG THỨC SẼ LOAD LẠI CO THỂ BỊ ERROR)
_________________________________________________
Thắc mắc thường gặp:
+ Cột Full Name/Link bị nhảy hàng hoặc bị lỗi ?
-> Do khi copy từ file gốc, Create working sheet sẽ copy luôn hàm tạo liên kết nên nếu hàm đó không thể copy thì sẽ gây lỗi ( thường truờng hợp này bạn không thể copy cột Full Nmae/ Link được )
+ Lỗi #NAME khi dùng hàm phonenumber ?
-> Lỗi này xuất hiện khi bạn gõ sai cú pháp của hàm
+ Làm gì khi dùng phonenumber chỉ ra #Error ?
->  Xóa ô bị Error, sau đó nhấn Ctrl + Z để undo lại, hàm phonenumber sẽ tự động load lại, làm lại cho đên khi load được Work Phone .Nếu Error quá nhiều, tạo một sheet mới hoặc dùng 1 sheet cũ mà hàm phonenumber hoạt động tốt. Copy 2 cột company và work city vào và dùng =phonenumber để tìm. Hoặc đợi sang ngày hôm sau dùng lại.
+ Cách hàm phonenumber tìm work phone ?
-> Dùng company và work city làm keyword để tìm work phone trên google ([Company] [Work city] phone number)
+ Cách tính Duplicate trong Report Helper ?
-> Tất cả các dòng có dữ liệu xuất hiện chữ "dup" hoặc "duplicate" (cả chữ hoa và chữ thường) trong cột "Note"
+C ách tính no mail  ?
-> Tất cả các dòng không có Home Phone, Work phone và Additional Phone.
+ Cách tính Phone number ?
-> Tất cả các dòng không có cả Mobile, Home và Work phone
*/

function phonenumber(company,city) {
  if(company !== "" && city !== ""){ 
    var company_array = company.replace(/\W+/g, " ").split(" ");
    var city_array = city.replace(/\W+/g, " ").split(" ");
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
      return output;
    }
    else if(string.search('<span data-dtype="d3ph">') !== -1){
      var a = string.indexOf('<span data-dtype="d3ph">');
      var cut = string.substring(a, string.length);
      var output = cut.substring(29,cut.indexOf("</span>"));
      return output;
    }
    else if(string.search('</span> &middot; ') !== -1){
      var a = string.indexOf('</span> &middot; ');
      var cut = string.substring(a, string.length);
      var output = cut.substring(16,cut.indexOf("</div>"));
      return output;
    }
    /*else
    return "";
    return output;
    Logger.log(output);*/
  }
}

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Create Working Sheet', functionName: 'create_working_sheet'},
    {name: 'Duplicate Detector', functionName: 'duplicate_finder'},
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
  var total = 0;
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
    if(data[0][i].toLowerCase() === "tag (dsource)" || data[0][i].toLowerCase() === "tag dsource" || data[0][i].toLowerCase() === "tag \(female\)" || data[0][i].toLowerCase() === "tag female" || data[0][i].indexOf("Dsource") !== -1 || data[0][i].indexOf("dsource") !== -1 || data[0][i].indexOf("DSOURCE") !== -1 || data[0][i].indexOf("Tag") !== -1)
      for(j=1; j<lrow; j++){
        if(data[j][i] !== "")
          dsource++;
      }
    else if(data[0][i].toLowerCase().indexOf("name") !== -1)
      for(j=1; j<lrow; j++){
        if(data[j][i] !== "")
          total++;
    }
    else if(data[0][i].toLowerCase() === "facebook profile")
      for(j=1; j<lrow; j++){
        if(data[j][i] !== "")
          facebookurl++;
      }
    else if(data[0][i].toLowerCase() === "twitter url")
      for(j=1; j<lrow; j++){
        if(data[j][i] !== "")
          twitterurl++;
      }
    else if(data[0][i].toLowerCase() === "github profile")
      for(j=1; j<lrow; j++){
        if(data[j][i] !== "")
          github++;
      }
    else if(data[0][i].toLowerCase() === "work city/town" || data[0][i].toLowerCase() === "work city")
      for(j=1; j<lrow; j++){
        if(data[j][i] !== "")
          workcity++;
      }
    else if(data[0][i].toLowerCase() === "work city/town.")
      workcity = 0;
    else if(data[0][i].toLowerCase() === "work state")
      for(j=1; j<lrow; j++){
        if(data[j][i] !== "")
          workstate++;
      }
    else if(data[0][i].toLowerCase() === "facebook email")
      for(j=1; j<lrow; j++){
        if(data[j][i] !== "")
          facebookemail++;
      }
    else if(data[0][i].toLowerCase() === "home email"){
      if(total>lrow-1){
        for(j=1; j<lrow; j++){
          if(data[j][i] === "" && data[j][i+1] === "" && data[j][i+2] === ""){
            nomail++
              for(l=0; l<len; l++){
                if(l===0 || data[0][l].toLowerCase() === "note"){
                  if(data[j][l].toLowerCase().indexOf("dup") !== -1  || data[j][l].toLowerCase().indexOf("duplicate") !== -1){
                    dupnm++;
                  }
                }
              }
          }
        }
      }
      else{
      for(j=1; j<total+1; j++){
        if(data[j][i] === "" && data[j][i+1] === "" && data[j][i+2] === ""){
          nomail++
          for(l=0; l<len; l++){
            if(data[0][l].toLowerCase() === "note"){
              if(data[j][l].toLowerCase().indexOf("dup") !== -1  || data[j][l].toLowerCase().indexOf("duplicate") !== -1){
                dupnm++;
              }
              }
            }
          }
        }
      }
    }
    /*else if(data[0][i].toLowerCase() === "mobile phone")
      for(j=1; j<lrow; j++){
        if(data[j][i] !== "")
          phonenumber++;
      }
    else if(data[0][i].toLowerCase() === "home phone")
      for(j=1; j<lrow; j++){
        if(data[j][i] !== "")
          phonenumber++;
      }*/
    else if(data[0][i].toLowerCase() === "work phone")
      for(j=1; j<lrow; j++){
        if(data[j][i] !== "" || data[j][i-1] !== "" || data [j][i-2] !== "")
          phonenumber++;
      }
    else if( i===0 || data[0][i].toLowerCase() === "note")
      for(j=1; j<lrow; j++){
        if(data[j][i].toLowerCase().indexOf("dup") !== -1  || data[j][i].toLowerCase().indexOf("duplicate") !== -1)
          duplicate++;
      }
  }
  
  if(total > lrow -1)
    total = lrow - 1;

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

function create_working_sheet(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var distributed_sheet = SpreadsheetApp.getActiveSheet()
  ss.insertSheet();
  var working_sheet = SpreadsheetApp.getActiveSheet()
  //ss.setActiveSheet(ss.getSheets()[0]);
  //var distributed_sheet = ss.getSheets()[0];
  var clen = distributed_sheet.getLastColumn();
  var rlen = distributed_sheet.getLastRow();
  //Logger.log(ws_rlen);
  var frow = 0;
  if(distributed_sheet.getRange(1, 8).getValue() !== "")
    frow = 1;
  else{
    for(i=2; i<10;i++){
      if(distributed_sheet.getRange(i, 8).getValue() !== ""){
        frow = i;
        break
      }
    }
  }
  var first_row = distributed_sheet.getRange(frow, 1, 1, clen).getValues();
  var values = [
    ['Home Email', 'Work Email', 'Additional Email', 'Mobile Phone', 'Home Phone', 'Work Phone','Facebook Profile', 'Twitter URL', 'Github Profile']
  ];
  //Logger.log(rlen);
  //Logger.log(frow);
  var colnum = 0;
  var fnchecker = false;
  var lpchecker = false;
  for(i=0; i<clen; i++){
    if(first_row[0][i].toLowerCase() === "full name/link"){
      distributed_sheet.getRange(frow, i+1, rlen-frow+1).copyTo(working_sheet.getRange("C1"));
      fnchecker = true; 
    }
    else if(first_row[0][i].toLowerCase() === "company")
      distributed_sheet.getRange(frow, i+1, rlen-frow+1).copyTo(working_sheet.getRange("D1"), {contentsOnly:true});
    else if(first_row[0][i].toLowerCase() === "position title")
      distributed_sheet.getRange(frow, i+1, rlen-frow+1).copyTo(working_sheet.getRange("E1"), {contentsOnly:true});
    else if(first_row[0][i].toLowerCase() === "work city/town"){
      distributed_sheet.getRange(frow, i+1, rlen-frow+1).copyTo(working_sheet.getRange("F1"), {contentsOnly:true});
      var work_city = distributed_sheet.getRange(frow+1, i+1).getValues();
      //Logger.log(work_city);
      if(work_city[0][0] === ""){
        working_sheet.getRange('G1').setValue('Work State').setFontFamily("Calibri").setFontWeight("bold").setFontSize(14).setBackground("#c9daf8");
        colnum++
      }
    }
    else if(first_row[0][i].toLowerCase() === "facebook email")
      working_sheet.getRange(1, colnum+16,1,1).setValue('Facebook Email').setFontFamily("Calibri").setFontWeight("bold").setFontSize(14).setBackground("#c9daf8");
    
    else if(first_row[0][i].toLowerCase() === "first name"){
      var fncl_id = 65 + i;  
    }
    else if(first_row[0][i].toLowerCase() === "last name"){
      var lncl_id = 65 + i; 
    }
    else if(first_row[0][i].toLowerCase() === "linkedin profile"){
      var lkcl_id_number = 65 + i;
      lpchecker = true;
      if(lkcl_id_number>90){
        var lkcl_id_0 = 65;
        var lkcl_id_1 = 64+ (65+i)%90;
      }
      else
        var lkcl_id = 65 + i; 
    }
    
    if(fnchecker === false && i===clen-1 && lpchecker === true){
      for(var j = 2; j<=rlen;j++){
        if(lkcl_id_number>90)
          var formula = '=HYPERLINK('+distributed_sheet.getName()+'!'+String.fromCharCode(lkcl_id_0,lkcl_id_1)+j+',CONCATENATE('+distributed_sheet.getName()+'!'+String.fromCharCode(fncl_id)+j+'," ",'+distributed_sheet.getName()+'!'+String.fromCharCode(lncl_id)+j+'))';
        else
          var formula = '=HYPERLINK('+distributed_sheet.getName()+'!'+String.fromCharCode(lkcl_id)+j+',CONCATENATE('+distributed_sheet.getName()+'!'+String.fromCharCode(fncl_id)+j+'," ",'+distributed_sheet.getName()+'!'+String.fromCharCode(lncl_id)+j+'))';
        working_sheet.getRange("C"+j).setFormula(formula);
      }
    }
    else if(fnchecker === false && i===clen-1 && lpchecker === false){
      for(var j = 2; j<=rlen;j++){
        var formula = '=HYPERLINK(,CONCATENATE('+distributed_sheet.getName()+'!'+String.fromCharCode(fncl_id)+j+'," ",'+distributed_sheet.getName()+'!'+String.fromCharCode(lncl_id)+j+'))';
        working_sheet.getRange("C"+j).setFormula(formula);
        
    }
    Browser.msgBox('Không tìm thấy cột Linkedin Profile ở sheet gốc! Nhấn OK để tiếp tục.');
  }
  }
  var ws_rlen = working_sheet.getLastRow();
  
  working_sheet.getRange('A1').setValue('Note').setFontFamily("Calibri").setFontWeight("bold").setFontSize(14).setBackground("#c9daf8");
  working_sheet.getRange('B1').setValue('Tag (dsource)').setFontFamily("Calibri").setFontWeight("bold").setFontSize(14).setBackground("#c9daf8");
  working_sheet.getRange('C1').setValue('Full Name/Link').setFontFamily("Calibri").setFontWeight("bold").setFontSize(14).setBackground("#c9daf8");
  working_sheet.getRange('D1').setFontFamily("Calibri").setFontWeight("bold").setFontSize(14).setBackground("#c9daf8");
  working_sheet.getRange('E1').setFontFamily("Calibri").setFontWeight("bold").setFontSize(14).setBackground("#c9daf8");
  if((working_sheet.getRange('F2').getValue() !== '' && working_sheet.getRange('F3').getValue() !== '') && (working_sheet.getRange('F2').getValue().toLowerCase() !== 'not found' && working_sheet.getRange('F3').getValue().toLowerCase() !== 'not found')){
    working_sheet.getRange('F1').setValue('Work City/Town.').setFontFamily("Calibri").setFontWeight("bold").setFontSize(14).setBackground("#c9daf8");
  }
  working_sheet.getRange('F1').setFontFamily("Calibri").setFontWeight("bold").setFontSize(14).setBackground("#c9daf8");
  working_sheet.getRange(1, colnum+7,1,9).setValues(values).setFontFamily("Calibri").setFontWeight("bold").setFontSize(14).setBackground("#c9daf8");
  working_sheet.getRange(2, colnum+13,ws_rlen-1,3).setBackground("#ead1dc");
  working_sheet.getRange(2, 3,ws_rlen-1,1).setBackground("white");
  working_sheet.getRange(2, 6,ws_rlen-1,1).setBackground("white");
  working_sheet.setFrozenRows(1);
  working_sheet.setFrozenColumns(5);

}
