function sendTofinish(){
  
  //有効なGooglesプレッドシートを開く
  var ss = SpreadsheetApp.openById('1gMn6N77DZW4j9h7eRhI7jr3pIK4EzG0IW_mGlf3YzOQ');
  var sheet = ss.getSheetByName('フォーム')
  var cam_list = ss.getSheetByName('カメラの個数');
  
  //新規予約された行番号を取得
  var num_row = sheet.getLastRow();
  //「カメラの個数」シートの一番下の行番号を取得
  var num_row2 = cam_list.getLastRow();
  
  //今あるカメラのリスト
  var list = cam_list.getRange(1,1,num_row2,1).getValues();
  
  //返却日3日前と前日と当日にメールを送る
  //返却していない人たちの終了日 == 今日　ならその人にメールを送る
  
  //今日
  var today = new Date();
  today = new Date(today.getFullYear(),today.getMonth(),today.getDate());
  Logger.log(today);
  //３日前
  var three_days_ago;
  //１日前
  var one_days_ago;
  
  //メール送信用
  var day;
  var name;
  var mail;
  var camera;
  var tr = "\n--------------\n返却の手順は、まずslackで部長に「返却をしたい」ことを伝え、具体的に会う場所や日にちを決めます。\nその後、実際に会い、カメラを返して終了となります。\nその他、わからないことがあった際にも連絡をお願いします。\n必ず、返し忘れがないかどうかも確認をお願いします。\n--------------\n";
function mail_keywords(day){
  var thing = "カメラの返却日まであと"+ day + "日となりました。\n\n返却を忘れないようにしてください";
  return thing;
  }
  //返却済か、未返却
  var rental;
  
  
  //２行目からデータが入力されている
  var j = 2;
  while(j <= num_row)
  {
    rental = sheet.getRange(j, 9).getValue();
    //返却したかどうか
    Logger.log(rental);
    
    if(rental == "未返却")
    {
    
      //返却日
      var last_day = new Date(sheet.getRange(j,8).getValue());
      var last_day_h = Moment.moment(last_day);//比較できるようにmomentオブジェクト化
      Logger.log(last_day);
      //返却３日前
      three_days_ago = new Date(last_day.getFullYear(),last_day.getMonth(),last_day.getDate()-3);
      var three_days_ago_h = Moment.moment(three_days_ago); //比較できるようにmomentオブジェクト化
      Logger.log(three_days_ago);
      //返却1日前
      one_days_ago = new Date(last_day.getFullYear(),last_day.getMonth(),last_day.getDate()-1);
      var one_days_ago_h = Moment.moment(one_days_ago);//比較できるようにmomentオブジェクト化
      Logger.log(one_days_ago);
      
      // ”day”を指定することでtodayとthree_days_ago_hの年月日を比較している
      if(Moment.moment(today).isSame(three_days_ago_h,'day')) //今日が返却期限の3日前
      {
        day = 3;
        name = sheet.getRange(j, 3).getValue();
        mail = sheet.getRange(j,2).getValue();
        var thing = mail_keywords(day);
        MailApp.sendEmail(mail,"カメラの返却期限のお知らせ",tr + thing); //メールを送信
        Logger.log("day = 3 成功");
      }
      if(Moment.moment(today).isSame(one_days_ago_h,'day')) //今日が返却期限の1日前
      {
        day = 1;
        name = sheet.getRange(j, 3).getValue();
        mail = sheet.getRange(j,2).getValue();
        var thing = mail_keywords(day);
        MailApp.sendEmail(mail,"カメラの返却期限のお知らせ",tr + thing); //メールを送信
        Logger.log("day = 1 成功");
        
      }
      if(Moment.moment(today).isSame(last_day_h,"day")) //今日が返却期限当日
      {
        day = "当";
        name = sheet.getRange(j, 3).getValue();
        mail = sheet.getRange(j,2).getValue();
        var thing = mail_keywords(day);
        MailApp.sendEmail(mail,"カメラの返却期限のお知らせ",tr + thing); //メールを送信
        Logger.log("day = 0 成功");
        
      }
      
      if(Moment.moment(today).isAfter(last_day_h,'day')) //todayが、last_dayよりも後の場合　＝返却期限を超えている
      {
        thing = "返却期限を過ぎています。\n至急、返却をお願いします。"
        name = sheet.getRange(j, 3).getValue();
        mail = sheet.getRange(j,2).getValue();
        MailApp.sendEmail(mail,"カメラの返却期限のお知らせ",tr + thing); //メールを送信
        Logger.log("day = -1 成功");
      }
    }
    j = j + 1;
  }
}