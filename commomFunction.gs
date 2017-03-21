/* 生成郵箱-姓名鍵值對字典 */
function getDict(){
  var listSheet = SpreadsheetApp.openById("1YMMT1I4LQvypqus_b-AYSyEH-2nKr0L2ah14Y5oz5mw").getSheetByName("花名冊");
  /* 獲取郵箱 */
  var mailValue = listSheet.getRange("H2:H").getValues();
  /* 獲取姓名 */
  var nameValue = listSheet.getRange("D2:D").getValues();
  /* 獲取離職時間 */
  var leftDate = listSheet.getRange("G2:G").getValues();
  
  for( var i = 0; i < mailValue.length; i++ )
  {
    /* 剔除郵箱信息丟失的行 */
    if( mailValue[i] == "" )
    {
      mailValue.splice(i, 1);
      nameValue.splice(i, 1);
      leftDate.splice(i, 1);
      i--;
    }
    /* 剔除已離職的行 */
    else if( leftDate[i] != "" )
    {
      mailValue.splice(i, 1);
      nameValue.splice(i, 1);
      leftDate.splice(i, 1);
      i--;
    }
  }
  
  var listDict = new FtofsStandardLibrary.dictionary();
  for( var i = 0; i < mailValue.length; i++ )
  {
    listDict.push(mailValue[i], nameValue[i]);
  }
  
  return listDict;
}

/* 獲取當前用戶郵箱和姓名 */
function getInfo(){
  try{
    var userEmail = Session.getEffectiveUser().getEmail();
    var cache = CacheService.getScriptCache();
    var dictCached = cache.get("mail-name-cache");
    /* 字典緩存存在或未過期 */
    if (dictCached != null) {
      var Dict = JSON.parse(dictCached);
      var info ={   
        email : userEmail,
        name : Dict[userEmail].toString()
      };
    }
    /* 字典緩存不存在或已過期 */
    else{
      var Dict = getDict();
      cache.put("mail-name-cache", Dict.toString(), 21600);
      Logger.log("put cache succeed");
      var info ={   
        email : userEmail,
        name : Dict.getValue(userEmail).toString()
      };
    }
    
    return info;
  }
  catch(e){
    throw e;
  }
}

/* 編輯時自動添加編輯者姓名和編輯時間 */
function onEdit(event){
  var actSht = event.source.getActiveSheet();
  if( actSht.getName() == "評論區" ){
    var actRng = event.source.getActiveRange();
    var index = actRng.getRowIndex();
    var user = getInfo().name;
    var date = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");
    var userCol = 2;
    var dateCol = 3;
    var userCell = actSht.getRange(index,userCol);
    var dateCell = actSht.getRange(index,dateCol);  
    userCell.setValue(user);
    dateCell.setValue(date);
  }
  else if( actSht.getName() == "這裡有好東西" ){
    var actRng = event.source.getActiveRange();
    var index = actRng.getRowIndex();
    var user = getInfo().name;
    var date = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");
    var userCol = 3;
    var dateCol = 4;
    var userCell = actSht.getRange(index,userCol);
    var dateCell = actSht.getRange(index,dateCol);  
    userCell.setValue(user);
    dateCell.setValue(date);
  }
}