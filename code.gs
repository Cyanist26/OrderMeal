try{
  /* 加載當天日期 */
  var today = new Date();
  /* 設置截單時間 */
  var viewDeadline = new Date();
  var submitDeadline = new Date();
  if( today.getDay() != 6 ){
    /* 非週六的查看截止時間 */
    viewDeadline.setHours(10,15,0);
    /* 非週六的下單截止時間 */
    submitDeadline.setHours(9,55,0);
  }  
  else{
    /* 週六的查看截止時間 */
    viewDeadline.setHours(10,30,0);
    /* 週六的下單截止時間 */
    submitDeadline.setHours(10,20,0);
  }
    
  /* 判斷是否已過查看截止時間 */
  var isDateChanged = false;
  if ( today >= viewDeadline ) 
  {
    today.setDate(today.getDate() + 1 );
    isDateChanged = true;
  }
  /* 設置加載天數限制 */
  var year = today.getFullYear();
  var month = today.getMonth() + 1;
  var date = today.getDate();
  var days = new Date(year, month, 0).getDate();
  /* 格式化日期 */
  var formatToday = Utilities.formatDate(today, "GMT+8", "yyyy-MM-dd");
  
  /* 加載需要使用的文件或表格 */
  var hostFile = SpreadsheetApp.openById("1ClRplJYKcLJuYS64c2FJj1AmuUQxLpJzHNgV1QaTS6g"); 
  var orderSheet = hostFile.getSheetByName("本月訂餐");
  var orderData = orderSheet.getDataRange().getDisplayValues();  
  
  /* 加載日期所在列 */
  var todayIndexInOrdCol = FtofsStandardLibrary.getIndexByContent(true, formatToday, orderData)[0][1];
  
  /* 記錄日誌所需的資源 */
  var userInfo = getUserInfo();
  var logSheet = SpreadsheetApp.openById("1DcHqkayhEoET1D1L0QzWpCHjV5gOeQcsyflj9pohMxg").getSheetByName("訂餐表日誌");
}
catch(e){
  SpreadsheetApp.getUi().alert(e.toString());
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui
    .createMenu('FTOFS') 
    .addItem('訂餐請戳我','showOrderDialog')
    .addToUi();
    
//  ui.alert("週六截單時間提早到10:20！！");
  showOrderDialog();
  myLogger("INFO", "打開訂餐表");
}

function showOrderDialog(){
  try{
    var html = HtmlService.createHtmlOutputFromFile('order')
      .setWidth(1200)
      .setHeight(550);
    SpreadsheetApp.getUi().showModalDialog(html, '  ');
  }
  catch(e){
    SpreadsheetApp.getUi().alert(e.toString());
  }
}

function getUserInfo(){
  try{ 
    var userEmail = Session.getEffectiveUser().getEmail();
    var cache = CacheService.getScriptCache();
    var userEmailCached = cache.get(userEmail);
    if ( false/*userEmailCached != null*/ )
    {
      return JSON.parse(userEmailCached);
    }
    else
    {
      var data = SpreadsheetApp.openById("1YMMT1I4LQvypqus_b-AYSyEH-2nKr0L2ah14Y5oz5mw").getSheetByName("花名冊").getDataRange().getDisplayValues();  
      var emailIndex = FtofsStandardLibrary.getIndexByContent(true, userEmail, data); 
      
      for( var i = 0; i < emailIndex.length; i++ )
      {
        var emailIndexRow = emailIndex[i][0];
        var leftday = data[emailIndexRow - 1][6];
        var leftdayDate = new Date(leftday.replace(/-/g, "/"));
        leftdayDate.setMonth(leftdayDate.getMonth() + 1 );
        leftdayDate.setDate(leftdayDate.getDate() + 1 );
        
        
        if( leftday == "" || today - leftdayDate < 0 )
        {
          var info = {
            division : data[emailIndexRow - 1][0],
            department : data[emailIndexRow - 1][1],
            name : data[emailIndexRow - 1][3]
          };
          cache.put(userEmail, JSON.stringify(info), 21600);
          return info;
        }
      }
      throw "賬號信息錯誤。"
    }
  }
  catch(e){
    throw e;
  }
}

function getFormatedDateAndMenu(name){
  /* 調試模式，屏蔽其他用戶 */
//  if( Session.getEffectiveUser().getEmail() != "cwq@ftofs.info" )
//    throw "休息片刻，馬上回來~"
  /* 訂餐表中名字所在行 */
  var nameIndexInOrdRow = FtofsStandardLibrary.getIndexByContent(true, name, orderData)[0][0];
  /* 加載參數表 */
  var menuData = hostFile.getSheetByName("參數表").getRange(1, 1, 35, 17).getDisplayValues();
  /* 參數表中日期所在行 */
  var todayIndexRow = FtofsStandardLibrary.getIndexByContent(true, formatToday, menuData)[0][0];
  
  var formatedDateAndMenu = [];
  
  for( var day = 1; day < 7; day++ )
  {
    /* 超出本月範圍 */
    if( ( date + day - 1 ) > days )
      formatedDateAndMenu.push(["下月訂餐","敬請期待~","","","","","","","","","","","","","","","",""]); 
    else
    {
      /* 寫入菜單 */
      formatedDateAndMenu.push(menuData[todayIndexRow + day - 2]);
      /* 寫入已選菜式 */
      formatedDateAndMenu[day - 1].push(orderData[nameIndexInOrdRow - 1][todayIndexInOrdCol + day - 2]);
    }
  }
  
  return formatedDateAndMenu;
}

function submitOrder(orderResult){
  /* 訂餐表中名字所在行 */
  var nameIndexInOrdRow = FtofsStandardLibrary.getIndexByContent(true, orderResult[0], orderData)[0][0];
  /* 提交前加載參數表確認 */
  var confirmMenuData = hostFile.getSheetByName("參數表").getRange(1, 1, 35, 17).getDisplayValues();
  /* 參數表中日期所在行 */
  var confirmTodayIndexRow = FtofsStandardLibrary.getIndexByContent(true, formatToday, confirmMenuData)[0][0];
  /* 用於截單的日期對象 */
  var confirmDate = new Date();
  
  try{
    /* 調試模式，屏蔽其他用戶 */
//    if( Session.getEffectiveUser().getEmail() != "ly@ftofs.info" )
//       throw "調試中，請稍後~"
    var error = "";
    /* 修改訂餐表day1 */
    /* 已定闡釋序號 */
    var menuNum = eval(orderResult[1].slice(0,1));
    /* 已定菜式名稱 */
    var menu = orderResult[1].slice(1);
    /* 提交當天菜式的截止時間 */
    if ( !isDateChanged && ( confirmDate >= submitDeadline ) ) 
      throw "吉時已過，無法修改當天訂餐！"
    /* 選擇不訂 */
    else if( menu == "none" ) 
      orderSheet.getRange(nameIndexInOrdRow, todayIndexInOrdCol).clearContent();
    /* 提交時菜式已滿 */
    else if( (eval(confirmMenuData[confirmTodayIndexRow - 1][menuNum + 6]) <= eval(confirmMenuData[confirmTodayIndexRow - 1][menuNum + 11]))
          && (orderData[nameIndexInOrdRow - 1][todayIndexInOrdCol - 1] != menu) )
      error += "你點慢了，第 1";
    /* 新增或修改菜式 */
    else if( orderData[nameIndexInOrdRow - 1][todayIndexInOrdCol - 1] != menu ) 
      orderSheet.getRange(nameIndexInOrdRow, todayIndexInOrdCol).setValue(menu);
    
    /* 修改訂餐表day2-day6 */
    for( var day = 2; day < 7; day++)
    {
      /* 已定菜式序號 */
      var menuNum = new Number(orderResult[day].slice(0,1));
      /* 已定菜式名稱 */
      var menu = orderResult[day].slice(1);
      /* 超出本月範圍 */
      if( (( date + day - 1 ) > days) )
        continue;
      /* 選擇不訂 */
      else if( menu == "none" ) {
        orderSheet.getRange(nameIndexInOrdRow, todayIndexInOrdCol + day - 1).clearContent();
        continue;
      }
      /* 提交時改菜式已滿 */
      else if( (eval(confirmMenuData[confirmTodayIndexRow + day - 2][(menuNum + 6)]) <= eval(confirmMenuData[confirmTodayIndexRow + day - 2][menuNum + 11]))
          && (orderData[nameIndexInOrdRow - 1][todayIndexInOrdCol + day - 2] != menu) ) {
        if( error == "" )
          error += "你點慢了，第 " + day;
        else
          error += "， " + day;
      }
      /* 新增或修改菜式 */
      else if( orderData[nameIndexInOrdRow - 1][todayIndexInOrdCol + day - 2] != menu )
        orderSheet.getRange(nameIndexInOrdRow, todayIndexInOrdCol + day - 1).setValue(menu);
    }
      
    myLogger("INFO", orderResult.slice(1, 6).toString());
    
    if( error != "" ) {
      error += " 天點的菜已經被搶光了！其他餐已訂好！";
      myLogger("ERROR", "提交出錯：" + error);
      throw error
    }
  }
  catch(e) {
    throw e;
  }
}

function myLogger(type, content) {
  var date = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");
  if(userInfo != null)
    logSheet.appendRow([userInfo.name, date, type, content]);
  else
    logSheet.appendRow(["", date, type, content]);
}