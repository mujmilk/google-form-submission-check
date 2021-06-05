// ---------------------------
// 提出状況チェックスクリプト
// ---------------------------
// このスクリプトが使える状況: フォルダの中に 課題名 - 名前.拡張子 というファイルがいっぱいある
// (Google Formでファイルを提出すると，ファイル名 - Googleアカウント名.拡張子というファイル名に自動変換されることを利用している)
//
// 使い方:
//   1: Google スプレッドシートを作成
//   2: 作成したスプレッドシートを開いて，ツール＞スクリプトエディタを起動
//   3: このコードをコピペして実行
//
//   【要確認】
//   実行後，表示＞ログ または 実行ログ で 提出されたファイル名のうち課題一覧になかったファイル名，提出した人の中で名前一覧に名前が無かった人の一覧を表示する
//
function checkSubmit() {
  
  // 毎回変更する箇所ここから ------------------------------------
  
  // 課題があるフォルダのID
  // URLの最後(https://drive.google.com/drive/u/3/folders/xxxのxxx)
  var folderId = '';
  
  // 必須課題名
  // 最初が正しい課題名, 後ろに表記ゆれを任意の個数指定できる
  // 表記ゆれファイル名で提出されたファイルは，スプレッドシートの出力に！表記ゆれマークが付く
  var assignmentList = [["work11", "work1.1"],
                        ["work12", "work1.2"],
                        ["work13", "work1.3"],
                        ["work14", "work1.4"]];
  
  // 課題の提出締切日時     (YYYY, MM, dd, hh, mm, ss)
  var deadline = new Date(2021,  12, 31, 23, 59, 59);
  
  // 毎回変更する箇所ここまで ------------------------------------
  
  
  // 提出ファイルの拡張子
  var fileExtension = ".cpp";
  
  // 履修者の人数
  var studentsNum = 100;
  
  // 名簿取得
  // 名簿の形式は，一行目が['学籍番号','名前'](ヘッダ),二行目以降が[学籍番号, 名前]であることを想定
  var namesheetId = '名簿のスプレッドシートのID';
  var nameList = SpreadsheetApp.openById(namesheetId).getSheetByName('list').getRange(2, 2, studentsNum).getValues();
  
  var folder = DriveApp.getFolderById(folderId),
      files = folder.getFiles(),
      list = {},
      outputList = [],
      sheet, range;
  
  for(let i=0; i<nameList.length; ++i) {
    list[nameList[i]] = {};
    list[nameList[i]]['sum'] = 0;
    for(let j=0; j<assignmentList.length; ++j) {
      list[nameList[i]][assignmentList[j][0]] = "";
    }
  }
  
  while(files.hasNext()) {
    var buff = files.next();

    // 締切よりも後に更新されたファイルをチェックしない
    var lastUpdatedDate = buff.getLastUpdated();
    if (lastUpdatedDate > deadline) {
      Logger.log('締切を過ぎています: '+buff.getUrl());
      continue;
    }

    // buff.getName (ファイル名 - 名前.pen) から名前を切り出す
    var name = buff.getName().split(" - ")[1].split(fileExtension)[0];
    var submittedFileName = buff.getName().split(" - ")[0];
    
    // list[名前][ファイル名] = 1
    if (name in list) {
      var notInListFlag = 1;
      for (let i=0; i<assignmentList.length; ++i) {
        for (let j=0; j<assignmentList[i].length; ++j) {
          if (submittedFileName == assignmentList[i][j]) {
            notInListFlag = 0;
            // 複数提出されている場合，最初のファイルへリンクされる
            if (list[name][assignmentList[i][0]] == "") {
              list[name]['sum'] += 1;
              
              if (j != 0) {
                var str = '=HYPERLINK("' + buff.getUrl() + '","〇！表記ゆれ")';
                list[name][assignmentList[i][0]] = str;
              } else {
                var str = '=HYPERLINK("' + buff.getUrl() + '","〇")';
                list[name][assignmentList[i][0]] = str;
              }

              break;
            }
          }
        }
      }
      if (notInListFlag == 1) {
        // 指定されたファイル名と異なる名前で提出されたファイルがあった場合,ログに表示する
        Logger.log("不明なファイル名:"+name+","+submittedFileName);
      }
    }else{
      // リストにない名前で提出されたファイルがあった場合,ログに表示する
      Logger.log("名簿に名前が存在しません:"+name);
    }
  };
  
  // スプレッドシートのヘッダ作成
  var header = ["名前","判定","必須課題提出数"];
  for(let i=0; i<assignmentList.length; ++i) {
    header.push(assignmentList[i][0]);
  }
  outputList.push(header);
  
  for(let i=0; i<nameList.length; ++i) {
    var row = [];
    
    row.push(nameList[i]);
    
    if(list[nameList[i]]['sum'] == assignmentList.length){
      row.push("〇");
    }else if(list[nameList[i]]['sum'] == 0){ // 提出数0
      row.push("");
    }else{
      row.push("△");
    }
    
    row.push(list[nameList[i]]['sum']);  // row[2]
    
    for(let j=0; j<assignmentList.length; ++j) {
      row.push(list[nameList[i]][assignmentList[j][0]]);
    }
    
    outputList.push(row);
  }

  sheet = SpreadsheetApp.getActiveSheet();
  range = sheet.getRange(1, 1, outputList.length, outputList[0].length);
  range.setValues(outputList); // 書き出し
}