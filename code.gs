try{
  //加載當天日期
  var today = new Date();
  //10點之後不再加載當天菜式
  var isDateChanged = false;
  if ( today.getHours() >= 10 && today.getMinutes() >= 0 ) {today.setDate(today.getDate() + 1 );isDateChanged = true;}
  var formatToday = Utilities.formatDate(today, "GMT+8", "yyyy-MM-dd");
  
  //加載需要使用的文件或表格
  var hostFile = SpreadsheetApp.openById("1ClRplJYKcLJuYS64c2FJj1AmuUQxLpJzHNgV1QaTS6g"); 
  var orderSheet = hostFile.getSheetByName("本月訂餐");
  var orderData = orderSheet.getDataRange().getDisplayValues();  
  
  //加載日期和姓名地址 
  var todayIndexInOrd = FtofsStandardLibrary.getIndexByContent(true, formatToday, orderData);
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
  
}

function showOrderDialog(){
  try{
    var html = HtmlService.createHtmlOutputFromFile('order')
      .setWidth(800)
      .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, '訂餐');
  }
  catch(e){
    ui.alert(e.toString());
  }
}

function getUserInfo(){
  try{ 
    var userEmail = Session.getEffectiveUser().getEmail();
    var data = SpreadsheetApp.openById("1YMMT1I4LQvypqus_b-AYSyEH-2nKr0L2ah14Y5oz5mw").getSheetByName("花名冊").getDataRange().getValues();  
    var odate = new Date();
    var emailIndex = FtofsStandardLibrary.getIndexByContent(true, userEmail, data);
    for( var i = 0; i < emailIndex.length; i++ )
    {
      if(  odate.getTime() - data[(emailIndex[i][0])-1][6] > 0 )
      {     
        info = {
          division : data[(emailIndex[i][0])-1][0],
          department : data[(emailIndex[i][0])-1][1],
          name : data[(emailIndex[i][0])-1][3]
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
  var nameIndexInOrd = FtofsStandardLibrary.getIndexByContent(true, name, orderData);
  var menuData = hostFile.getSheetByName("參數表").getRange(1, 1, 35, 17).getDisplayValues();
  var todayIndex = FtofsStandardLibrary.getIndexByContent(true, formatToday, menuData);
  
  var formatedDateAndMenu = [];//last index = 16 已選菜式
  formatedDateAndMenu.push(menuData[(todayIndex[0][0] - 1)]);
  formatedDateAndMenu.push(menuData[todayIndex[0][0]]);
  formatedDateAndMenu.push(menuData[(todayIndex[0][0] + 1)]);
  formatedDateAndMenu.push(menuData[(todayIndex[0][0] + 2)]);
  formatedDateAndMenu.push(menuData[(todayIndex[0][0] + 3)]);
  formatedDateAndMenu.push(menuData[(todayIndex[0][0] + 4)]);
  formatedDateAndMenu[0].push(orderData[nameIndexInOrd[0][0] - 1][todayIndexInOrd[0][1] - 1]);
  formatedDateAndMenu[1].push(orderData[nameIndexInOrd[0][0] - 1][todayIndexInOrd[0][1]]);
  formatedDateAndMenu[2].push(orderData[nameIndexInOrd[0][0] - 1][todayIndexInOrd[0][1] + 1]);
  formatedDateAndMenu[3].push(orderData[nameIndexInOrd[0][0] - 1][todayIndexInOrd[0][1] + 2]);
  formatedDateAndMenu[4].push(orderData[nameIndexInOrd[0][0] - 1][todayIndexInOrd[0][1] + 3]);
  formatedDateAndMenu[5].push(orderData[nameIndexInOrd[0][0] - 1][todayIndexInOrd[0][1] + 4]);  
  
  return formatedDateAndMenu;
}

function submitOrder(orderResult){
  var nameIndexInOrd = FtofsStandardLibrary.getIndexByContent(true, orderResult[0], orderData);
  var confirmMenuData = hostFile.getSheetByName("參數表").getRange(1, 1, 35, 17).getDisplayValues();
  var todayIndex = FtofsStandardLibrary.getIndexByContent(true, formatToday, confirmMenuData);
  //Logger.log(orderResult);
  
  try{
    //修改訂餐表day2-day6
    for( var day = 2; day < 7; day++)
    {
      var menuNum = new Number(orderResult[day].slice(0,1));
      if( orderResult[day] == "0none" ) {
        orderSheet.getRange(nameIndexInOrd[0][0], todayIndexInOrd[0][1] + day - 1).clearContent();
        continue;
      }
      else if( orderData[nameIndexInOrd[0][0] - 1][todayIndexInOrd[0][1] + day - 2] == orderResult[day].slice(1) ) continue;
      else if( new Number(confirmMenuData[(todayIndex[0][0] + day - 2)][(menuNum + 6)]) <= new Number(confirmMenuData[(todayIndex[0][0] + day - 2)][(menuNum + 11)]))
        throw "你點慢了，第 " + day + " 天點的菜已經被搶光了！前 " + (day - 1) + " 天的菜已訂好！"
      orderSheet.getRange(nameIndexInOrd[0][0], todayIndexInOrd[0][1] + day - 1).setValue(orderResult[day].slice(1));
    }
    
    //修改訂餐表day1
    var menuNum = eval(orderResult[1].slice(0,1));
    if ( !isDateChanged && today.getHours() >= 10 ) 
      throw "吉時已過，無法修改當天訂餐！"
    else if( orderResult[1] == "0none" ) 
      orderSheet.getRange(nameIndexInOrd[0][0], todayIndexInOrd[0][1]).clearContent();
    else if( eval(confirmMenuData[(todayIndex[0][0] - 1)][(menuNum + 6)]) <= eval(confirmMenuData[(todayIndex[0][0] - 1)][(menuNum + 11)]))
      throw "你點慢了，第 1 天點的菜已經被搶光了！"
    else if( orderData[nameIndexInOrd[0][0] - 1][todayIndexInOrd[0][1] - 1] != orderResult[1].slice(1) ) 
      orderSheet.getRange(nameIndexInOrd[0][0], todayIndexInOrd[0][1]).setValue(orderResult[1].slice(1));
  }
  catch(e) {throw e;}
}