// ---------------------------
// 提出状況チェックスクリプト：ファイル名自由版
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
  
  // 課題があるフォルダのID 任意の個数指定できる
  // URLの最後(https://drive.google.com/drive/u/3/folders/xxxのxxx)
  var folderId = ['提出ファイルが格納されているフォルダのID_1',　'提出ファイルが格納されているフォルダのID_2'];
  
  // 最低限提出する必要があるファイルの数
  // 提出数がこれ以上であれば〇，これ以下0より上であれば△，0であれば×がスプレッドシートの名前の横に記録される
  var fileNum = 5;
  
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
  
  var   list = {},
        outputList = [],
        sheet, range;
  
  for(let i=0; i<nameList.length; ++i) {
    list[nameList[i]] = {};
    list[nameList[i]]['sum'] = 0;
  }

  for (let folderNum = 0; folderNum < folderId.length; folderNum++){
    var folder = DriveApp.getFolderById(folderId[folderNum]),
        files = folder.getFiles();
    
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
        // 複数提出されている場合、最初に提出されたファイルへリンクされる 修正の必要があるかも
        var str = '=HYPERLINK("' + buff.getUrl() + '","'+ submittedFileName +'")';
        list[name][submittedFileName] = str;
        list[name]['sum']++;
      }else{
        // リストにない名前で提出されたファイルがあった場合,ログに表示する
        Logger.log("名簿に名前が存在しません:"+name);
      }
    };
  }

  // スプレッドシートのヘッダ作成
  var header = ["名前","判定","課題提出数"];
  while(header.length<50){
      header.push("");
  }
  outputList.push(header);
  
  for(let i=0; i<nameList.length; ++i) {
    var row = [];
    
    row.push(nameList[i]);
    
    if(list[nameList[i]]['sum'] >= fileNum){
      row.push("〇");
    }else if(list[nameList[i]]['sum'] == 0){ // 提出数0
      row.push("");
    }else{
      row.push("△");
    }
    
    for(var key in list[nameList[i]]) {
      row.push(list[nameList[i]][key]);
    }

    while(row.length < 50){
      row.push("");
    }
    
    outputList.push(row);
  }

  sheet = SpreadsheetApp.getActiveSheet();
  range = sheet.getRange(1, 1, outputList.length, 50);
  range.setValues(outputList); // 書き出し
}