try{
  /* 加載當天日期 */
  var today = new Date();
  /* 10點之後不再加載當天菜式 */
  var isDateChanged = false;
  if ( today.getHours() >= 10 && today.getMinutes() >= 0 ) {today.setDate(today.getDate() + 1 );isDateChanged = true;}
  /* 格式化日期 */
  var formatToday = Utilities.formatDate(today, "GMT+8", "yyyy-MM-dd");
  
  /* 加載需要使用的文件或表格 */
  var hostFile = SpreadsheetApp.openById("1ClRplJYKcLJuYS64c2FJj1AmuUQxLpJzHNgV1QaTS6g"); 
  var orderSheet = hostFile.getSheetByName("本月訂餐");
  var orderData = orderSheet.getDataRange().getDisplayValues();  
  
  /* 加載日期所在列 */
  var todayIndexInOrdCol = FtofsStandardLibrary.getIndexByContent(true, formatToday, orderData)[0][1];
}
catch(e){
  SpreadsheetApp.getUi().alert(e.toString())
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var test = today.getHours();
  ui
  .createMenu('FTOFS') 
  .addItem('訂餐請戳我','showOrderDialog')
  .addToUi();
  
  showOrderDialog();
}

function showOrderDialog(){
  try{
    var html = HtmlService.createHtmlOutputFromFile('order')
      .setWidth(1000)
      .setHeight(550);
    SpreadsheetApp.getUi().showModalDialog(html, '訂餐');
  }
  catch(e){
    ui.alert(e.toString());
  }
}

function getUserInfo(){
  try{ 
    var userEmail = Session.getEffectiveUser().getEmail();
    var data = SpreadsheetApp.openById("1YMMT1I4LQvypqus_b-AYSyEH-2nKr0L2ah14Y5oz5mw").getSheetByName("花名冊").getDataRange().getDisplayValues();  
    
    var emailIndex = FtofsStandardLibrary.getIndexByContent(true, userEmail, data); 
    
    for( var i = 0; i < emailIndex.length; i++ )
    {
      var emailIndexRow = emailIndex[i][0];
      var leftday = data[emailIndexRow - 1][6];

      if(  leftday == "" )
      {     
        info = {
          division : data[emailIndexRow - 1][0],
          department : data[emailIndexRow - 1][1],
          name : data[emailIndexRow - 1][3]
        };       
        return info;
      }
    }
  }
  catch(e){
    throw e;
  }
}

function getFormatedDateAndMenu(name){
  /* 訂餐表中名字所在行 */
  var nameIndexInOrdRow = FtofsStandardLibrary.getIndexByContent(true, name, orderData)[0][0];
  /* 加載參數表 */
  var menuData = hostFile.getSheetByName("參數表").getRange(1, 1, 35, 17).getDisplayValues();
  /* 參數表中日期所在行 */
  var todayIndexRow = FtofsStandardLibrary.getIndexByContent(true, formatToday, menuData)[0][0];
  
  var formatedDateAndMenu = [];
  /* 寫入菜單 */
  formatedDateAndMenu.push(menuData[todayIndexRow - 1]);
  formatedDateAndMenu.push(menuData[todayIndexRow]    );
  formatedDateAndMenu.push(menuData[todayIndexRow + 1]);
  formatedDateAndMenu.push(menuData[todayIndexRow + 2]);
  formatedDateAndMenu.push(menuData[todayIndexRow + 3]);
  formatedDateAndMenu.push(menuData[todayIndexRow + 4]);
  /* 寫入已選菜式 */
  formatedDateAndMenu[0].push(orderData[nameIndexInOrdRow - 1][todayIndexInOrdCol - 1]);
  formatedDateAndMenu[1].push(orderData[nameIndexInOrdRow - 1][todayIndexInOrdCol]    );
  formatedDateAndMenu[2].push(orderData[nameIndexInOrdRow - 1][todayIndexInOrdCol + 1]);
  formatedDateAndMenu[3].push(orderData[nameIndexInOrdRow - 1][todayIndexInOrdCol + 2]);
  formatedDateAndMenu[4].push(orderData[nameIndexInOrdRow - 1][todayIndexInOrdCol + 3]);
  formatedDateAndMenu[5].push(orderData[nameIndexInOrdRow - 1][todayIndexInOrdCol + 4]);  
  
  return formatedDateAndMenu;
}

function submitOrder(orderResult){
  /* 訂餐表中名字所在行 */
  var nameIndexInOrdRow = FtofsStandardLibrary.getIndexByContent(true, orderResult[0], orderData)[0][0];
  /* 提交前加載參數表確認 */
  var confirmMenuData = hostFile.getSheetByName("參數表").getRange(1, 1, 35, 17).getDisplayValues();
  /* 參數表中日期所在行 */
  var confirmTodayIndexRow = FtofsStandardLibrary.getIndexByContent(true, formatToday, confirmMenuData)[0][0];
  
  try{
    /* 修改訂餐表day1 */
    /* 已定闡釋序號 */
    var menuNum = eval(orderResult[1].slice(0,1));
    /* 已定菜式名稱 */
    var menu = orderResult[1].slice(1);
    /* 提交時已過截止時間 */
    if ( !isDateChanged && today.getHours() >= 10 ) 
      throw "吉時已過，無法修改當天訂餐！"
    /* 選擇不訂 */
    else if( menu == "none" ) 
      orderSheet.getRange(nameIndexInOrdRow, todayIndexInOrdCol).clearContent();
    /* 提交時菜式已滿 */
    else if( eval(confirmMenuData[confirmTodayIndexRow - 1][menuNum + 6]) <= eval(confirmMenuData[confirmTodayIndexRow - 1][menuNum + 11]))
      throw "你點慢了，第 1 天點的菜已經被搶光了！"
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
      /* 選擇不訂 */
      if( menu == "none" ) {
        orderSheet.getRange(nameIndexInOrdRow, todayIndexInOrdCol + day - 1).clearContent();
        continue;
      }
      /* 提交時改菜式已滿 */
      else if( new Number(confirmMenuData[confirmTodayIndexRow + day - 2][(menuNum + 6)]) <= new Number(confirmMenuData[confirmTodayIndexRow + day - 2][menuNum + 11]))
        throw "你點慢了，第 " + day + " 天點的菜已經被搶光了！前 " + (day - 1) + " 天的菜已訂好！"
      /* 新增或修改菜式 */
      else if( orderData[nameIndexInOrdRow - 1][todayIndexInOrdCol + day - 2] != menu )
        orderSheet.getRange(nameIndexInOrdRow, todayIndexInOrdCol + day - 1).setValue(menu);
    }  
  }
  catch(e) {throw e;}
}