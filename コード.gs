const color=[['#76a5af','#a2c4c9'],['#93c47d','#b6d7a8'],['#ffd966','#ffe599'],['#f6b26b','#f9cb9c'],['#e06666','#ea9999'],['#77d9a8','#bfe4ff'],['#c8c8cb','#84919e']];
const arr_day = new Array('日', '月', '火', '水', '木', '金', '土');

/**　初期設定用関数　**/
function initialization() {
  const tags=['活動日','開始時刻','終了時刻','活動場所','希望者1','優先度','参加','希望者2','優先度','参加','希望者3','優先度','参加','希望者4','優先度','参加','希望者5','優先度','参加','状態','結果'];
  const tags2=['学籍番号','氏名','学年','学科','参加回数','参加回数調整','','活動場所','','活動時間'];

  //スプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('予定一覧');
  var sheet_data = spreadsheet.getSheetByName('データ');

  sheet_data.getRange(/*始行(1,2,3)*/1,/*始列(A,B,C)*/1,/*行範囲(1~)*/1,/*列範囲(1~)*/6).setBackground(color[1][0]);
  sheet_data.getRange(/*始行(1,2,3)*/1,/*始列(A,B,C)*/8,/*行範囲(1~)*/1,/*列範囲(1~)*/1).setBackground(color[1][0]);
  sheet_data.getRange(/*始行(1,2,3)*/1,/*始列(A,B,C)*/10,/*行範囲(1~)*/1,/*列範囲(1~)*/1).setBackground(color[1][0]);

  //日付の入力規定・表示形式を設定
  var cell = sheet.getRange(2,1,sheet.getMaxRows()-1,1);
  var rule = SpreadsheetApp.newDataValidation().requireDate().build();
  cell.setDataValidation(rule);
  cell.setNumberFormat('MM/DD')

  //開始・終了時刻の入力規定・表示形式を設定
  var range = sheet_data.getRange(2,10,sheet.getMaxRows()-1,1);
  range.setNumberFormat('@');
  cell = sheet.getRange(2,2,sheet.getMaxRows()-1,2);
  rule = SpreadsheetApp.newDataValidation().requireValueInRange(range).build();
  cell.setDataValidation(rule);
  cell.setNumberFormat('@');

  //参加者の入力規定を設定
  range = sheet_data.getRange(2,2,sheet.getMaxRows()-1,1);
  for(var i=0;i<5;i++){
    cell = sheet.getRange(2,5+3*i,sheet.getMaxRows()-1,1);
    rule = SpreadsheetApp.newDataValidation().requireValueInRange(range).build();
    cell.setDataValidation(rule);
  }

  //活動場所の入力規定を設定
  range = sheet_data.getRange(2,8,sheet.getMaxRows()-1,1);
  cell = sheet.getRange(2,4,sheet.getMaxRows()-1,1);
  rule = SpreadsheetApp.newDataValidation().requireValueInRange(range).build();
  cell.setDataValidation(rule);

  values = ['参加','不参加'];
  for(var i=0;i<5;i++){
    cell = sheet.getRange(2,7+3*i,sheet.getMaxRows()-1,1);
    rule = SpreadsheetApp.newDataValidation().requireValueInList(values).build();
    cell.setDataValidation(rule);
  }

  values = ['募集中','仮決定','連絡済み','終了','中止'];
  cell = sheet.getRange(2,20,sheet.getMaxRows()-1,1);
  rule = SpreadsheetApp.newDataValidation().requireValueInList(values).build();
  cell.setDataValidation(rule);

  
  sheet.getRange(/*始行(1,2,3)*/1,/*始列(A,B,C)*/1,/*行範囲(1~)*/1,/*列範囲(1~)*/4).setBackground(color[1][0]);
  for(var i=0;i<5;i++){
    sheet.getRange(/*始行(1,2,3)*/1,/*始列(A,B,C)*/5+i*3,/*行範囲(1~)*/1,/*列範囲(1~)*/1).setBackground(color[i][0]);
    sheet.getRange(/*始行(1,2,3)*/1,/*始列(A,B,C)*/6+i*3,/*行範囲(1~)*/1,/*列範囲(1~)*/2).setBackground(color[i][1]);
  }
  for(var i=0;i<21;i++){
    sheet.getRange(/*始行(1,2,3)*/1,/*始列(A,B,C)*/1+i,/*行範囲(1~)*/1,/*列範囲(1~)*/1).setValue(tags[i]);
    sheet.getRange(/*始行(1,2,3)*/1,/*始列(A,B,C)*/1+i,/*行範囲(1~)*/1,/*列範囲(1~)*/1).setHorizontalAlignment('center');
  }

  for(var i=0;i<10;i++){
    sheet_data.getRange(/*始行(1,2,3)*/1,/*始列(A,B,C)*/1+i,/*行範囲(1~)*/1,/*列範囲(1~)*/1).setValue(tags2[i]);
    sheet_data.getRange(/*始行(1,2,3)*/1,/*始列(A,B,C)*/1+i,/*行範囲(1~)*/1,/*列範囲(1~)*/1).setHorizontalAlignment('center');
  }

  //縦線設定
  var border=[4,7,10,13,16,19];
  for(var i=0;i<6;i++){
    sheet.getRange(1,border[i],sheet.getMaxRows(),1).setBorder(false, false, false, true, false, false,'black',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }

  //横線設定
  sheet.getRange(1,1,1,21).setBorder(false, null, true, null, null, null,'black',SpreadsheetApp.BorderStyle.SOLID_THICK);

}


/**　メイン処理　**/
function process() {
  //スプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('予定一覧');
  var dataLen = sheet.getDataRange().getValues().length;
  var myRange = sheet.getRange(/*始行(1,2,3)*/2,/*始列(A,B,C)*/1,/*行範囲(1~)*/dataLen,/*列範囲(1~)*/21);
  var sheet_s = myRange.getValues();

  var sheet_s_color = myRange.getValues();
  var sheet_s_font_color = myRange.getValues();

  var sheet_data = spreadsheet.getSheetByName('データ');
  var dataLen_data = sheet_data.getDataRange().getValues().length;
  var myRange_data = sheet_data.getRange(/*始行(1,2,3)*/2,/*始列(A,B,C)*/1,/*行範囲(1~)*/dataLen_data,/*列範囲(1~)*/6);
  var sheet_data_s = myRange_data.getValues();

  //色を入れる変数として初期化
  for(var i=0;i<dataLen;i++){
    for(var j=0;j<21;j++){
      sheet_s_color[i][j]=null;
    }
  }
  for(var i=0;i<dataLen;i++){
    for(var j=0;j<21;j++){
      sheet_s_font_color[i][j]='black';
    }
  }


  //セルの色を設定
  for(var i=0;i<dataLen;i++){
    if(sheet_s[i][0]!=''){
      if(sheet_s[i][19]=='募集中'){
        for(var j=0;j<21;j++){
          sheet_s_color[i][j]=color[5][0];
        }
      }else if(sheet_s[i][19]=='仮決定'||sheet_s[i][19]=='連絡済み'){
        for(var j=0;j<21;j++){
          sheet_s_color[i][j]=color[5][1];
        }
      }else if(sheet_s[i][19]=='終了'){
        for(var j=0;j<21;j++){
          sheet_s_color[i][j]=color[6][0];
        }
      }else if(sheet_s[i][19]=='中止'){
        for(var j=0;j<21;j++){
          sheet_s_color[i][j]=color[6][1];
        }
        for(var j=0;j<5;j++){
          sheet_s[i][6+j*3]='不参加';
        }
      }
    }
  }

  //参加回数書き込み処理
  for(var i=0;i<dataLen_data-1;i++){
    var count=0;
    var name = sheet_data_s[i][1];
    if(name!=''){
      for(var j=0;j<dataLen;j++){
        for(var k=0;k<5;k++){
          if(name==sheet_s[j][4+k*3]&&sheet_s[j][6+k*3]=='参加'){
            sheet_s_font_color[j][4+k*3]='white';
            count++;
          }
        }
      }
    }
    if(sheet_data_s[i][5]!=''){
      sheet_data_s[i][4]=count+parseInt(sheet_data_s[i][5]);
    }else{
      sheet_data_s[i][4]=count+0;
    }
  }

  //優先度書き込み処理
  for(var i=0;i<dataLen;i++){
    var nums=[-1,-1,-1,-1,-1];
    for(var j=0;j<5;j++){
      var name = sheet_s[i][4+j*3];
      if(name != ''){
        for(var k=0;k<dataLen_data;k++){
          if(sheet_data_s[k][1]==name){
            nums[j]= sheet_data_s[k][4]+j*0.5;
          }
        }
      }
    }
    for(var l=0;l<5;l++){
      var rank=1;
      for(var n=-l;n<(-l+5);n++){
        if(l!=(l+n)&&nums[l]>nums[l+n]){
          if(nums[l+n]!=-1){
            rank++;
          }
        }
      }
      if(nums[l]!=-1){
        sheet_s[i][5+l*3]=rank+'位';
      }
    }
  }

  //文章書き込み処理
  for(var i=0;i<dataLen;i++){
    if(sheet_s[i][0]!=''){
      var dates=new Date(sheet_s[i][0]);
      var sentece=(dates.getMonth()+1)+'/'+dates.getDate()+'('+arr_day[dates.getDay()]+')は';
      var count=0;
      for(var j=0;j<5;j++){
        var name = sheet_s[i][4+j*3];
        if(name!='' && sheet_s[i][6+j*3]=='参加'){
          sentece+=name+'、';
          count++;
        }
      }
      if(count==0){
          sentece+='参加者が集まりませんでした'
        }
      sheet_s[i][20] = sentece;
    }
  }

  //シートに反映
  myRange.setBackgrounds(sheet_s_color);
  myRange.setFontColors(sheet_s_font_color);
  myRange_data.setValues(sheet_data_s);
  myRange.setValues(sheet_s);
}

/**　メール送信処理　**/
function sendMail(){
  //スプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('予定一覧');
  var dataLen = sheet.getDataRange().getValues().length;
  var myRange = sheet.getRange(/*始行(1,2,3)*/2,/*始列(A,B,C)*/1,/*行範囲(1~)*/dataLen,/*列範囲(1~)*/21);
  var sheet_s = myRange.getValues();

  var nowdate = new Date();

  for(var i = 0; i < dataLen; i++) {
    var status = sheet_s[i][19];
    var day = new Date(sheet_s[i][0]);
    var names = sheet_s[i][20];
    var basyo = sheet_s[i][3];
    var t_stert = sheet_s[i][1];
    var t_end = sheet_s[i][2];

    if(day.getMonth()==nowdate.getMonth() && day.getDate()==nowdate.getDate()){
      //メールの件名
      const subject = '【自動配信】本日活動あり';
      //メールの本文
      const body = '本日'+names+'が'+basyo+'で活動予定です。'+'時間は'+t_stert+'~'+t_end+'です。';
      
      //メールを送信する
      GmailApp.sendEmail('', subject, body);
    }  
  }
};
