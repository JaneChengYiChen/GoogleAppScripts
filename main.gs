// @ts-nocheck
function myFunction() {
  var SpreadSheet = SpreadsheetApp.getActive();
  //原始回應
  var ResponseSheetName = "表單回應";
  var ResponseSheet = SpreadSheet.getSheetByName(ResponseSheetName);
  var ResponselastRow = ResponseSheet.getLastRow();
  
  //原始回應的欄位賦值
  var Time = ResponseSheet.getSheetValues(ResponselastRow,1,1,1);
  var Branch = ResponseSheet.getSheetValues(ResponselastRow,2,1,1);
  var Name = ResponseSheet.getSheetValues(ResponselastRow,3,1,1);
  var LifeNumber = ResponseSheet.getSheetValues(ResponselastRow,4,1,1);
  var PropertyNumber = ResponseSheet.getSheetValues(ResponselastRow,5,1,1);
  var Mobile = ResponseSheet.getSheetValues(ResponselastRow,6,1,1);
  var ID = ResponseSheet.getSheetValues(ResponselastRow,7,1,1);
  var Birthday = ResponseSheet.getSheetValues(ResponselastRow,8,1,1);
  var Email = ResponseSheet.getSheetValues(ResponselastRow,9,1,1);
  var Remark = ResponseSheet.getSheetValues(ResponselastRow,22,1,1);
  
  // 現在應該要被匯入到哪張資料表(YYYYMM)
  var SalesmanSheetName = getSalesmanSheetName();
  
  // 資料表是否存在，如果不存在，那就新增一個吧！
  var SalesmanSheet = SpreadSheet.getSheetByName(SalesmanSheetName);
  if (SalesmanSheet == null) { 
    createNewSalesmanSheet(SpreadSheet, SalesmanSheetName);
    var SalesmanSheet = SpreadSheet.getSheetByName(SalesmanSheetName);
  }
  
  //此筆的代號
  var uuid = Utilities.getUuid();
  
  //將需要調整的保險公司加入欄位
  var supAllianzArray = [10, 'Allianz','英文大寫身分證10碼'];
  var supTaiwanArray = [11, 'Taiwan','英文大寫身分證10碼'];
  var supTaiwanMobileArray = [12, 'TaiwanMobile','英文大寫身分證10碼'];
  var supAiaArray = [13, 'Aia','登錄證字號10碼'];
  var supTransGlobeArray = [14, 'TransGlobe','WK4-英文大寫身分證10碼'];
  var supSklArray = [15, 'Skl','登錄證字號10碼'];
  var supChubbArray = [16, 'Chubb','LS+英文身分證10碼'];
  var supYauntaArray = [17, 'Yaunta','英文大寫身分證10碼'];
  var supHontaiArray = [18, 'Hontai','英文大寫身分證10碼'];
  var supChubbPropertyArray = [19, 'ChubbProperty','英文大寫身分證10碼'];
  var supHotainsArray = [20, 'Hotains','英文大寫身分證10碼'];
  var supFubonArray = [21, 'Fubon','英文大寫身分證10碼'];
  var supRemark = [22, 'Remark',Remark];
  
  var supArrayList = [supAllianzArray, supTaiwanArray, supTaiwanMobileArray, supAiaArray,
                     supTransGlobeArray, supSklArray, supChubbArray, supYauntaArray,
                     supHontaiArray, supChubbPropertyArray, supHotainsArray, supFubonArray, supRemark];
  
  var supList = [];
  var number = 0;
  
  for (var key in supArrayList) {
    var value = supArrayList[key];
    var SalesmanAnswer =  ResponseSheet.getSheetValues(ResponselastRow,value[0],1,1); //業務員的答案
    
    if(SalesmanAnswer =="是" || SalesmanAnswer != ''){
      var SupName = value[1]; //保險公司名稱
      var SupAccount = value[2]; //保險公司帳號
      var SupSheet = SpreadSheet.getSheetByName(SupName); 
      var SuplastRow = SupSheet.getLastRow();
      var InsertSupRowPosition = SuplastRow+1;

      if (SalesmanAnswer =="是"){
        var InsertSupData = [uuid, Name, LifeNumber, PropertyNumber, Mobile,ID, Birthday, null ,SalesmanSheetName,  Branch,Email];
      var InsertSupRange = SupName+"!A"+ InsertSupRowPosition +":K"+ InsertSupRowPosition;
      SupSheet.getRange(InsertSupRange).setValues([InsertSupData]);
      }

      if (SalesmanAnswer != '' && SalesmanAnswer !="是"){
        var InsertSupData = [uuid, Name, LifeNumber, PropertyNumber, Mobile,ID, Remark, null ,SalesmanSheetName,  Branch,Email];
      var InsertSupRange = SupName+"!A"+ InsertSupRowPosition +":K"+ InsertSupRowPosition;
      SupSheet.getRange(InsertSupRange).setValues([InsertSupData]);
      }
      
      
      if (number ==0 ){
        supList.push([uuid, Time, Branch, Name, LifeNumber, PropertyNumber, Mobile,ID, Birthday, Email, SupName, SupAccount]);
      } else {
        supList.push([null, null, null, null, null, null, null, null, null, null, SupName, SupAccount]);
      }
      
      var number = number + 1;  
    }
  }
  
  
  //將原始回應的欄位加入業務員統計的Sheet
  var SalesmanInsertTopRow = SalesmanSheet.getLastRow() +1;
  var SalesmanInsertBottomRow = SalesmanInsertTopRow + supList.length -1;
  var InsertSalesmanRange = SalesmanSheetName+"!A"+ SalesmanInsertTopRow +":L"+ SalesmanInsertBottomRow;
  SalesmanSheet.getRange(InsertSalesmanRange).setValues(supList);
  
  //已完成按鈕
  SalesmanSheet.getRange("O"+SalesmanInsertTopRow).insertCheckboxes();
  //是否要寄送
  SalesmanSheet.getRange("P"+SalesmanInsertTopRow).insertCheckboxes();
  //總共有多少個欄位
  SalesmanSheet.getRange("Q"+SalesmanInsertTopRow).setValue(supList.length);
}

function dataSync(e) {
 
  var SpreadSheet = e.source;
  var Sheets = SpreadSheet.getActiveSheet();
  var SheetName = Sheets.getName();
  var ResponseSheetName = "表單回應";
  var SupSheet = SpreadSheet.getSheetByName(SheetName);
  var activeRng = Sheets.getActiveRange();
  var activeRow = activeRng.getRow();
  
  //取得應該要匯入到哪一個月份的資料
  if( SupSheet == null || ExclusionSheets(SheetName) < 0) {
    return;
  }
  
  var SalesmanSheetName = SupSheet.getRange('I'+ activeRow).getValue();
  //此兩表單寫的方式不一樣 先return
  if (SheetName === ResponseSheetName || SalesmanSheetName == null || SalesmanSheetName == ''){
    return;
  } 
 
  var SalesmanSheet = SpreadSheet.getSheetByName(SalesmanSheetName);
  var uuidNumber = SupSheet.getRange('A'+ activeRow).getValue();
  var column = onSearch(uuidNumber, SalesmanSheetName, 1); 
  var row = onSearch(SheetName, SalesmanSheetName, 11, column) + column - 2;
  SupSheet.getRange("H"+ activeRow).copyTo(SalesmanSheet.getRange(row,13,1,1), {contentsOnly:true});
}


function uuid() {
  return Utilities.getUuid();
}

function getSalesmanSheetName() {
  var date = new Date(); 
  var month = (date.getMonth()+1).toString();
  var year = (date.getYear()+1900).toString();
  var SalesmanSheetName = year + month;
  
  return SalesmanSheetName;
}

function createNewSalesmanSheet(SpreadSheet, SalesmanSheetName) {
  var SalesmanSheet = SpreadSheet.insertSheet();
  SalesmanSheet.setName(SalesmanSheetName);
  var InsertHeader = ['編號', '時間戳記','分行','姓名','壽險證號','產險證號','手機號碼','身分證字號'
                        ,'生日','信箱','申請的保險公司','帳號(預設)','密碼','備註','已完成','寄送'];
  var InsertHeaderRange = SalesmanSheetName +"!A1:P1";
  SalesmanSheet.getRange(InsertHeaderRange).setValues([InsertHeader]);
  SalesmanSheet.getRange("O1:P1").setBackground("#46bdc6"); 
  SalesmanSheet.getRange("L1:N1").setBackground("#e6b8af");
  SalesmanSheet.getRange("A1:K1").setBackground("#434343");
  SalesmanSheet.getRange("A1:P1").setFontColor('white').setHorizontalAlignment("center");
  SalesmanSheet.setFrozenRows(1); //第一列凍結
  SalesmanSheet.getRange('A1:P1').protect().setDescription('protected range'); //設定保護
}


//ref https://stackoverflow.com/questions/18482143/search-spreadsheet-by-column-return-rows
function onSearch(keyword ,sheetname, whichColumn, startRow = 2)
{
    var searchString = keyword;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetname); 
    var column = whichColumn; //是在第幾欄  
    var columnValues = sheet.getRange(startRow, column, sheet.getLastRow()).getValues(); //第一欄是標題
    var searchResult = columnValues.findIndex(searchString); //Row Index - 2

    if(searchResult != -1) { //searchResult + 2 = 實際row
        return searchResult + 2;
    }
  
    return false;
}

Array.prototype.findIndex = function(search){
  if(search == "") return false;
  
  for (var i = 0; i < this.length; i++)
    if (this[i] == search) return i;

  return -1;
}

//按鈕寄信
function sendmail(e) {
  //看看還有多少寄信扣打
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  Logger.log("Remaining email quota: " + emailQuotaRemaining); 
  
  var SpreadSheet = e.source;
  var Sheets = SpreadSheet.getActiveSheet();
  var SheetName = Sheets.getName();
  
  if (ExclusionSheets(SheetName) > 0) {
    return;
  };
  
  var SalesmanSheet = SpreadSheet.getSheetByName(SheetName);
  
  var activeRng = Sheets.getActiveRange();
  var activeRow = activeRng.getRow();
  var activeColumn = activeRng.getColumn();
  var activeRngValue = activeRng.getValue();
  var rowStarPosition = activeRow;
  var rowNumber = SalesmanSheet.getRange('Q'+ activeRow).getValue();
  var salesmanEmail = SalesmanSheet.getRange('J'+ activeRow).getValue();
  var data = SalesmanSheet.getRange(rowStarPosition, 11, rowNumber, 4).getValues();
  if (activeColumn == 16 && activeRngValue != ''){  //如果點選寄送的按鈕
    //輸入寄送的日期
    SalesmanSheet.getRange(activeRow, 18).setValue(new Date());
    
    //準備信件內容
    var message = "夥伴您好，以下為您的申請結果：\n\n";
    for (var i = 0; i < data.length; ++i) {
      var row = data[i];
      var supplier = row[0]; //supplier
      var account = row[1]; //account
      var password = row[2]; //password
      var remark = row[3]; //remark
      message = message + supplier + ":\n";
      message = message + '  帳號：' + account + "\n";
      message = message + '  密碼：' + password + "\n";
      message = message + '  備註：' + remark + "\n";    
    }
    
    var subject = '帳號申請結果';
    message = message + "\n\n若有相關疑問，請寄至 steve.lee@gmail.com.tw \n行政團隊敬上";
    
    MailApp.sendEmail({
      to: salesmanEmail,
      subject: subject,
      cc: 'steve.lee@gmail.com.tw',
      name: '策劃室',
      body: message
    });
  } 
  
  return;
}

function ExclusionSheets(SheetName){
  var exclusion = ['表單回應', 'Allianz', 'TaiwanMobile', 'Aia',
                   'TransGlobe', 'Skl', 'Chubb', 'Yaunta','Hontai',
                   'ChubbProperty', 'Hotains', 'Fubon', 'Taiwan', 'Remark'];
  
  for (var i = 0; i < exclusion.length; i++)
    if (exclusion[i] == SheetName) return i;
  
  return -1; 
  
}