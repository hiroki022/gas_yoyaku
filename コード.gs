function doGet(e) {
var htmlutput =HtmlService.createTemplateFromFile("result").evaluate();
return htmlutput;

//下記はtumblerを使うことにより、必要なくなった。各ページをhtmlから送られてきたパラメータによってページを切り替える部分だった。
 /* var page=e.parameter["p"];
  if(page == "result")
  {  
    htmloutput = HtmlService.createTemplateFromFile("result").evaluate();
    return htmloutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  else if(page == "index")
  {
    htmloutput = HtmlService.createTemplateFromFile("index").evaluate();
    return htmloutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  else if(page == "top")
  {
    htmloutput = HtmlService.createTemplateFromFile("top").evaluate();
    return htmloutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  else if(page == "test")
  {
    htmloutput = HtmlService.createTemplateFromFile("test").evaluate();
    return htmloutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  else
  {
    var htmloutput = HtmlService.createTemplateFromFile("top").evaluate();
    return htmloutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }*/
}

function doPost(e){

  var ss=spreadsheet = SpreadsheetApp.openById('1gMn6N77DZW4j9h7eRhI7jr3pIK4EzG0IW_mGlf3YzOQ');
  var sheet = ss.getSheetByName('フォーム'); 
  
  var day = new Date();
  day = new Date(day.getFullYear(),day.getMonth(),day.getDate());
  var email=e.parameter.email;
  var nname=e.parameter.nname;
  var grad=e.parameter.grad;
  var course=e.parameter.course;
  var camera=e.parameter.camera;
  var bihin=e.parameter.bihin;
  
  //メールアドレスから学籍番号を抜き出す（メールアドレスから、数字以外を消して、「s」を足して格納）
  var reg = /\D/g;
  var number= "s"+email.replace(reg,"");
  
  sheet.appendRow([day,email,nname,grad,course,camera,"","","",number,bihin]);
  
  sendToCalendar();
  
  var htmloutput = HtmlService.createTemplateFromFile("result").evaluate();
  return htmloutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


/*function comp() {
  //有効なGooglesプレッドシートを開く
  var ss = SpreadsheetApp.openById('1gMn6N77DZW4j9h7eRhI7jr3pIK4EzG0IW_mGlf3YzOQ');
  var sheet = ss.getSheetByName('フォーム');
  var cam_list = ss.getSheetByName('カメラの個数');
  
  //新規予約された行番号を取得
  var num_row = sheet.getLastRow();
  //「カメラの個数」シートの一番下の行番号を取得
  var num_row2 = cam_list.getLastRow();
  //学籍番号を配列に格納
  var list_number = sheet.getRange().getValues();
  
  var thing = "返却が完了しました";
  var name;
  var mail;
  
  var rental;
  var j = 2;
  rental = sheet.getRange(1,num_row,1,1).getValue();
  Logger.log("rental= "+ rental);
  while(j <= num_row)
  {
    //rental = sheet.getRange(j, 9).getValue();
    //if(rental == "未返却")
    if(rental[j] == "未返却")
    {
      if(list_number[j] == number)
      {
        sheet.getRange(j,9).setValue("返却済み");
        name = sheet.getRange(j, 3).getValue();
        mail = sheet.getRange(j,2).getValue();
        MailApp.sendEmail(mail,name + "さんのカメラ返却完了のお知らせ",thing); //メールを送信
        name = sheet.getRange(j, 3).getValue();
        mail = cam_list.getRange(2,7).getValue();
        MailApp.sendEmail(mail,name + "さんのカメラ返却完了のお知らせ",thing); //メールを送信
      }
    }
    j = j + 1;
  }
}
*/

function sendToCalendar(e){
  try{
    //有効なGooglesプレッドシートを開く
    var ss = SpreadsheetApp.openById('1gMn6N77DZW4j9h7eRhI7jr3pIK4EzG0IW_mGlf3YzOQ');
    var sheet = ss.getSheetByName('フォーム')
    var cam_list = ss.getSheetByName('カメラの個数');
    
    //新規予約された行番号を取得
    var num_row = sheet.getLastRow();
    //「カメラの個数」シートの一番下の行番号を取得
    var num_row2 = cam_list.getLastRow();
    
    //新規予約の行の3列目
    var nname = sheet.getRange(num_row, 3).getValue();
    
    //メールアドレスの取得
    //2行目
    var nmail = sheet.getRange(num_row,2).getValue();
    //借りたいカメラの名前　F列
    var rental_camera = sheet.getRange(num_row,6).getValue();
//借りたい備品の名前　K列
    var rental_bihin = sheet.getRange(num_row,11).getValue();
    
    //今あるカメラのリスト
    var list = cam_list.getRange(1,1,num_row2,1).getValues();
    //今ある備品のリスト
    var bihin_list = cam_list.getRange(1,6,num_row2,1).getValues();
    Logger.log("備品" + bihin_list);
    
    //その人が借りたいカメラ
    //リストの中に借りたいカメラはどの行か調べる
    var i = 0;
    Logger.log(rental_camera);
    while( i <= rental_camera)
    {
      //iは借りたいカメラの行を知るため、しかし、スプレッドシート上では1列目に名称が入っているため、list[i-1]で一致する行にしないと1行ずれる
      if(list[i - 1] == rental_camera)
      {
        Logger.log('一致しました');
        break;
      }
      else
      {
        i = i +1;
      }
    }
    
    //備品でも同じことをする　関数化するのがめんどくさかった
    var k = 0;
    Logger.log(rental_bihin);
    while( k <= rental_bihin)
    {
      //iは借りたいカメラの行を知るため、しかし、スプレッドシート上では1列目に名称が入っているため、list[i-1]で一致する行にしないと1行ずれる
      if(bihin_list[k - 1] == rental_bihin)
      {
        Logger.log('一致しました');
        break;
      }
      else
      {
        k = k +1;
      }
    }
    
    var ndate = new Date(sheet.getRange(num_row, 1).getDisplayValue());//1行目から値読み取り
    var ndates = new Date(ndate.getFullYear(),ndate.getMonth(),ndate.getDate()); //開始日を抜き出す
    var ndatee = new Date(ndate.getFullYear(),ndate.getMonth() + 1,ndate.getDate() + 7); //終了日を抜き出す
    
    sheet.getRange(num_row,7).setValue(ndates); //開始日をスプレッドシートに入力
    sheet.getRange(num_row,8).setValue(ndatee); //終了日をスプレッドシートに入力
    sheet.getRange(num_row,9).setValue("未返却"); //未返却 を入れる
    
    //整数型に変換
    list.map(function (element) { return Number(element); });
    bihin_list.map(function (element) { return Number(element); });
    
    var cam_num = cam_list.getRange(i,3).getValue();
    var rent = cam_list.getRange(i,4).getValue();
    
    var bihin_num = cam_list.getRange(k,8).getValue();
    var bihin_rent = cam_list.getRange(k,9).getValue();
    
    Logger.log("cam_num=" + cam_num);
    Logger.log("rent=" + rent);
     Logger.log("bihin_num=" + cam_num);
    Logger.log("bihin_rent=" + rent);
    
    //備品の残り個数をセット
     var j = bihin_num - bihin_rent;
    //jを決めてから、rentに＋１をする
    cam_list.getRange(k,9).setValue(bihin_rent + 1);
    var bihin = cam_list.getRange(k,7).getDisplayValue();
    var no_bihin = "";
    if(j > bihin_rent){
    no_bihin = "<br/> " + bihin + "　は借りられているので、現在貸し出せません。"
    }
    
    j=0;
    //カメラの残り個数をセット
    var j = cam_num - rent;
    //jを決めてから、rentに＋１をする
    cam_list.getRange(i,4).setValue(rent + 1);
    
    Logger.log('j = ' + j);
    
    if(j > 0)
    { //在庫がまだある
      j = j + 1;
      
      //予約を記載するカレンダーを取得
      var cals = CalendarApp.getCalendarById("ug50e7guqcnmpog1eest35dmk0@group.calendar.google.com");
      
      cals.getEvents(ndates, ndatee); //その日に予定が入っていないか
      
      var camera = cam_list.getRange(i,2).getDisplayValue();
      var thing = nname+"様　"+camera+"　のご予約"
      
      //予約情報をカレンダーに追加
      
      var r = cals.createEvent(thing, ndates, ndatee);
      var thing = nname+"様　\n\n 予約を承りました。\n\n ありがとうございました \n借りた備品\n" + bihin
      
      MailApp.sendEmail(nmail,camera + "の予約",thing); //メールを送信
    }
    else{
      sheet.deleteRows(num_row);
      cam_list.getRange(i,4).setValue(rent - 1);
      var thing = nname + "様　\n\n カメラに先約がありましたので、\n 申し訳ございませんが、ご予約いただけませんでした。\n\n ご予定を変更して再度お申込みください";
      MailApp.sendEmail(nmail,"ご予約できませんでした",thing);
    }
    }
  
  catch(exp){
      //実行に失敗した時に通知
      Logger.log(exp.message, exp.message);
    }
}