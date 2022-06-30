const color=[['#76a5af','#a2c4c9'],['#93c47d','#b6d7a8'],['#ffd966','#ffe599'],['#f6b26b','#f9cb9c'],['#e06666','#ea9999'],['#77d9a8','#bfe4ff'],['#c8c8cb','#84919e']];
const arr_day = new Array('日', '月', '火', '水', '木', '金', '土');

/**　初期設定用関数　**/
function initialization() {
  const tags=['活動日','開始時刻','終了時刻','活動場所','希望者1','優先度','参加','希望者2','優先度','参加','希望者3','優先度','参加','希望者4','優先度','参加','希望者5','優先度','参加','状態','結果'];
  const tags2=['学籍番号','氏名','学年','学科','参加回数','参加回数調整','','活動場所','','活動時間'];

  //スプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_ScheduleList = spreadsheet.getSheetByName('予定一覧');
  var sheet_Data = spreadsheet.getSheetByName('データ');

  sheet_Data.getRange(/*始行(1,2,3)*/1,/*始列(A,B,C)*/10,/*行範囲(1~)*/1,/*列範囲(1~)*/1).setBackground(color[1][0]);

  //日付の入力規定・表示形式を設定
  var cell = sheet_ScheduleList.getRange(2,1,sheet_ScheduleList.getMaxRows()-1,1);
  var rule = SpreadsheetApp.newDataValidation().requireDate().build();
  cell.setDataValidation(rule);
  cell.setNumberFormat('MM/DD')

  //開始・終了時刻の入力規定・表示形式を設定
  var range = sheet_Data.getRange(2,10,sheet_ScheduleList.getMaxRows()-1,1);
  range.setNumberFormat('@');
  cell = sheet_ScheduleList.getRange(2,2,sheet_ScheduleList.getMaxRows()-1,2);
  rule = SpreadsheetApp.newDataValidation().requireValueInRange(range).build();
  cell.setDataValidation(rule);
  cell.setNumberFormat('@');

  //参加者の入力規定を設定
  range = sheet_Data.getRange(2,2,sheet_ScheduleList.getMaxRows()-1,1);
  for(var i=0;i<5;i++){
    cell = sheet_ScheduleList.getRange(2,5+3*i,sheet_ScheduleList.getMaxRows()-1,1);
    rule = SpreadsheetApp.newDataValidation().requireValueInRange(range).build();
    cell.setDataValidation(rule);
  }

  //活動場所の入力規定を設定
  range = sheet_Data.getRange(2,8,sheet_ScheduleList.getMaxRows()-1,1);
  cell = sheet_ScheduleList.getRange(2,4,sheet_ScheduleList.getMaxRows()-1,1);
  rule = SpreadsheetApp.newDataValidation().requireValueInRange(range).build();
  cell.setDataValidation(rule);

  values = ['参加','不参加'];
  for(var i=0;i<5;i++){
    cell = sheet_ScheduleList.getRange(2,7+3*i,sheet_ScheduleList.getMaxRows()-1,1);
    rule = SpreadsheetApp.newDataValidation().requireValueInList(values).build();
    cell.setDataValidation(rule);
  }

  values = ['募集中','仮決定','連絡済み','終了','中止'];
  cell = sheet_ScheduleList.getRange(2,20,sheet_ScheduleList.getMaxRows()-1,1);
  rule = SpreadsheetApp.newDataValidation().requireValueInList(values).build();
  cell.setDataValidation(rule);

  
  sheet_ScheduleList.getRange(1,1,1,4).setBackground(color[1][0]);
  for(var i=0;i<5;i++){
    sheet_ScheduleList.getRange(1,5+i*3,1,1).setBackground(color[i][0]);
    sheet_ScheduleList.getRange(1,6+i*3,1,2).setBackground(color[i][1]);
  }

  //枠名記入
  for(var i=0;i<21;i++){
    sheet_ScheduleList.getRange(1,1+i,1,1).setValue(tags[i]);
    sheet_ScheduleList.getRange(1,1+i,1,1).setHorizontalAlignment('center');
  }
  for(var i=0;i<10;i++){
    sheet_Data.getRange(1,1+i,1,1).setValue(tags2[i]);
    sheet_Data.getRange(1,1+i,1,1).setHorizontalAlignment('center');
  }

  //縦線設定
  var border=[4,7,10,13,16,19];
  for(var i=0;i<6;i++){
    sheet_ScheduleList.getRange(1,border[i],sheet_ScheduleList.getMaxRows(),1).setBorder(false, false, false, true, false, false,'black',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }

  //横線設定
  sheet_ScheduleList.getRange(1,1,1,21).setBorder(false, null, true, null, null, null,'black',SpreadsheetApp.BorderStyle.SOLID_THICK);

}


/**　メイン処理　**/
function process() {
  //スプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_ScheduleList = spreadsheet.getSheetByName('予定一覧');
  var length_ScheduleList = sheet_ScheduleList.getDataRange().getValues().length;
  var range_ScheduleList = sheet_ScheduleList.getRange(2,1,length_ScheduleList,21);
  var value_ScheduleList = range_ScheduleList.getValues();

  var color_BG = range_ScheduleList.getValues();
  var color_font = range_ScheduleList.getValues();

  var sheet_Data = spreadsheet.getSheetByName('データ');
  var length_Data = sheet_Data.getDataRange().getValues().length;
  var range_Data = sheet_Data.getRange(2,1,length_Data,6);
  var value_Data = range_Data.getValues();

  //色を入れる変数として初期化
  for(var i=0;i<length_ScheduleList;i++){
    for(var j=0;j<21;j++){
      color_BG[i][j]=null;
    }
  }
  for(var i=0;i<length_ScheduleList;i++){
    for(var j=0;j<21;j++){
      color_font[i][j]='black';
    }
  }


  //セルの色を設定
  for(var i=0;i<length_ScheduleList;i++){
    if(value_ScheduleList[i][0]!=''){
      if(value_ScheduleList[i][19]=='募集中'){
        for(var j=0;j<21;j++){
          color_BG[i][j]=color[5][0];
        }
      }else if(value_ScheduleList[i][19]=='仮決定'||value_ScheduleList[i][19]=='連絡済み'){
        for(var j=0;j<21;j++){
          color_BG[i][j]=color[5][1];
        }
      }else if(value_ScheduleList[i][19]=='終了'){
        for(var j=0;j<21;j++){
          color_BG[i][j]=color[6][0];
        }
      }else if(value_ScheduleList[i][19]=='中止'){
        for(var j=0;j<21;j++){
          color_BG[i][j]=color[6][1];
        }
        for(var j=0;j<5;j++){
          value_ScheduleList[i][6+j*3]='不参加';
        }
      }
    }
  }

  //参加回数書き込み処理
  for(var i=0;i<length_Data-1;i++){
    var count=0;
    var name = value_Data[i][1];
    if(name!=''){
      for(var j=0;j<length_ScheduleList;j++){
        for(var k=0;k<5;k++){
          if(name==value_ScheduleList[j][4+k*3]&&value_ScheduleList[j][6+k*3]=='参加'){
            color_font[j][4+k*3]='white';
            count++;
          }
        }
      }
    }
    if(value_Data[i][5]!=''){
      value_Data[i][4]=count+parseInt(value_Data[i][5]);
    }else{
      value_Data[i][4]=count+0;
    }
  }

  //優先度書き込み処理
  for(var i=0;i<length_ScheduleList;i++){
    var nums=[-1,-1,-1,-1,-1];
    for(var j=0;j<5;j++){
      var name = value_ScheduleList[i][4+j*3];
      if(name != ''){
        for(var k=0;k<length_Data;k++){
          if(value_Data[k][1]==name){
            nums[j]= value_Data[k][4]+j*0.5;
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
        value_ScheduleList[i][5+l*3]=rank+'位';
      }
    }
  }

  //文章書き込み処理
  for(var i=0;i<length_ScheduleList;i++){
    if(value_ScheduleList[i][0]!=''){
      var dates=new Date(value_ScheduleList[i][0]);
      var sentece=(dates.getMonth()+1)+'/'+dates.getDate()+'('+arr_day[dates.getDay()]+')は';
      var count=0;
      for(var j=0;j<5;j++){
        var name = value_ScheduleList[i][4+j*3];
        if(name!='' && value_ScheduleList[i][6+j*3]=='参加'){
          sentece+=name+'、';
          count++;
        }
      }
      if(count==0){
          sentece+='参加者が集まりませんでした'
        }
      value_ScheduleList[i][20] = sentece;
    }
  }

  //シートに反映
  range_ScheduleList.setBackgrounds(color_BG);
  range_ScheduleList.setFontColors(color_font);
  range_Data.setValues(value_Data);
  range_ScheduleList.setValues(value_ScheduleList);
}

function today() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_ScheduleList = spreadsheet.getSheetByName('予定一覧');
  var length_ScheduleList = sheet_ScheduleList.getDataRange().getValues().length;
  var range_ScheduleList = sheet_ScheduleList.getRange(2,1,length_ScheduleList,21);
  var value_ScheduleList = range_ScheduleList.getValues();

  var sheet_Today = spreadsheet.getSheetByName('本日の活動');
  var length_Today = 20; //sheet_Today.getDataRange().getValues().length;
  var range_Today = sheet_Today.getRange(1,1,length_Today,4);
  var value_Today = range_Today.getValues();

  var nowdate = new Date();

  for(var i = 0; i < length_ScheduleList; i++) {
    var day = new Date(value_ScheduleList[i][0]);
    var names = value_ScheduleList[i][20];
    var basyo = value_ScheduleList[i][3];
    var t_stert = value_ScheduleList[i][1];
    var t_end = value_ScheduleList[i][2];
    var dayteSentense=day.getFullYear()+'/'+(day.getMonth()+1)+'/'+day.getDate()+'('+arr_day[day.getDay()]+')'

    if(day.getMonth()==nowdate.getMonth() && day.getDate()==nowdate.getDate()){
      value_Today[0]=['本日の日付','活動場所','開始時間','終了時間'];
      value_Today[1]=[dayteSentense,basyo,t_stert,t_end];
      var count=0;
      for(var j=0;j<5;j++){
        var name = value_ScheduleList[i][4+j*3];
        if(name!='' && value_ScheduleList[i][6+j*3]=='参加'){
          value_Today[3+count*3]=['活動予定者'+(count+1),'確認','',''];
          value_Today[4+count*3]=[name,'','',''];
          count++;
        }
      }
      
    } 
  }
  range_Today.setValues(value_Today);
}

/**　メール送信処理　**/
function sendMail(){
  //スプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_ScheduleList = spreadsheet.getSheetByName('予定一覧');
  var length_ScheduleList = sheet_ScheduleList.getDataRange().getValues().length;
  var range_ScheduleList = sheet_ScheduleList.getRange(2,1,length_ScheduleList,21);
  var value_ScheduleList = range_ScheduleList.getValues();

  var nowdate = new Date();

  for(var i = 0; i < length_ScheduleList; i++) {
    var status = value_ScheduleList[i][19];
    var day = new Date(value_ScheduleList[i][0]);
    var names = value_ScheduleList[i][20];
    var basyo = value_ScheduleList[i][3];
    var t_stert = value_ScheduleList[i][1];
    var t_end = value_ScheduleList[i][2];

    if(day.getMonth()==nowdate.getMonth() && day.getDate()==nowdate.getDate()){
      //メールの件名
      const subject = '【自動配信】本日活動あり';
      //メールの本文
      const body = '本日'+names+'が'+basyo+'で活動予定です。'+'時間は'+t_stert+'~'+t_end+'です。';
      
      //メールを送信する
      GmailApp.sendEmail('メールアドレス', subject, body);
    }  
  }
};
