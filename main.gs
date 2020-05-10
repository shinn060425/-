let id = "1HuaAdg-xdo6n_JJXCYgyzD8bAHjl09ngtJbVwF8SXNw";

// htmlからhtmlへの遷移用。"https://surleconomiejp.blogspot.com/2017/02/google-apps-script-htmlhtml.html"を参考にした。
function doGet(e) {
  let htmlOutput = "";
  if (!e.parameter.page) {
    htmlOutput = HtmlService.createTemplateFromFile('index').evaluate();
  } else {
    htmlOutput = HtmlService.createTemplateFromFile(e.parameter['page']).evaluate();
  }
  //スマホで表示する際に大きさを調整するため、metaタグを追加。"https://tonari-it.com/gas-web-add-meta-viewport/"を参考にした。
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return htmlOutput;
}

// htmlからhtmlへの遷移用。"https://surleconomiejp.blogspot.com/2017/02/google-apps-script-htmlhtml.html"を参考にした。
function getScriptUrl() {
  let url = ScriptApp.getService().getUrl();
  return url;
}

// cssやjavascriptの外部ファイルを読み込むための関数。"https://qiita.com/taromorimotohf/items/5e52cb9062600e8ccac3"を参考にした。
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

//スプレッドシートのデータを格納。"https://qiita.com/non-programmer/items/e822dbeb7b3ab6e403b3"を参考にした。
function loading(sheetname) {
  let ss = SpreadsheetApp.openById(id);
  let sh = ss.getSheetByName(sheetname);
  //最終行・列を調べる
  let lastrow = sh.getLastRow();
  let lastcolumn = sh.getLastColumn();
  //2行1列目最終行3列目までを二次元配列に入れています。
  let list = sh.getRange(1, 1, lastrow, lastcolumn).getValues();
  //時間要素をplane textに直す。objectの型のチェックには"http://blog.ko-atrandom.com/?p=310"を参考にした。
  for (let i=0; i<list.length; i++) {
    for (let j=0; j<list[i].length; j++) {
      if (typeof(list[i][j]) == "object") {
        let date = list[i][j];
        let dateString = Utilities.formatDate(date,"JST","yyyy/MM/dd");
        list[i].splice(j, 1, dateString);
      }
    }
  }
  return list;
}

// 屠殺の時にspreadsheetからデータを削除するための関数。
function delete_from_sheet_gs(sheetname, value, experiment, num, k) {
  let nowtime = new Date();
  let time = new Date(nowtime.getFullYear(), nowtime.getMonth(), nowtime.getDate());
  let ss = SpreadsheetApp.openById(id);
  let sh = ss.getSheetByName(sheetname);
  let shdel = ss.getSheetByName("処分マウス");
  let shdelrow = shdel.getLastRow() + 1 + k; //複数行入力の際、処理時間がほぼ同値のためlastrowがかぶってしまうことの一時的な解決。
  let row = num + 2;
  let lastcolumn = sh.getLastColumn();
  let dellist = sh.getRange(row, 1, 1, lastcolumn).getValues();
  let date = dellist[0][1];
  shdel.getRange(shdelrow, 1, 1, 6).setValues([[value, date, , sheetname, experiment, time]]);
  shdel.getRange(shdelrow, 3).setFormulaR1C1("=ROUNDDOWN((R" + shdelrow + "C6-R" + shdelrow + "C2)/7)");
  //4列目以降の記述を備考として追記する。
  for (let i=1; i <= lastcolumn - 3; i++) {
    shdel.getRange(shdelrow, i + 6).setValue(dellist[0][i + 2]);
  }
  //valueがマウスの数よりも少なければ値変更のみ、同値であれば列ごと消去
  if (value == dellist[0][0]) {
    sh.deleteRows(row);
  } else if (value < dellist[0][0]) {
    sh.getRange(row,1).setValue(dellist[0][0] - value);
  }
}

// 生まれたときにspreadsheetにhtmlから直接書き込むための関数。"https://www.pre-practice.net/2017/10/web.html"を参考にした。
function input_to_sheet_gs(value, born_date, mo_prefix) {
  
  //日付の足し算は"https://qiita.com/pppolon/items/b58f05b7534fe4b8ec72"を参考にした。
  let nowtime = new Date();
  let time = new Date(nowtime.getFullYear(), nowtime.getMonth(), nowtime.getDate() - born_date);
  
  let ss = SpreadsheetApp.openById(id);
  let sh = ss.getSheetByName("仔分け前");
  let shmo = ss.getSheetByName("交配メス");
  let shfa = ss.getSheetByName("交配オス");
   
  let lastRow = sh.getLastRow() + 1;
  sh.getRange(lastRow, 1, 1, 5).setValues([[value[0], value[1], time, , mo_prefix]]);
  sh.getRange(lastRow, 4).setFormulaR1C1("=ROUNDDOWN((TODAY()-R" + lastRow + "C3)/7)");
  
  //母マウスのprefixから「交配メス」シート内での行数を調べる
  let row = 0;
  let lastrow_mo = shmo.getLastRow();
  let prefix_list_mo = shmo.getRange(2, 4, lastrow_mo - 1, 1).getValues();
  for (let j=0; j < lastrow_mo - 1; j++) {
    if (prefix_list_mo[j][0] == mo_prefix) {
      row = j + 2;
    }
  }
  
  //母マウスの状態、交配成績を変更。
  shmo.getRange(row, 6).setValue("仔育て");
  let mating_num_mo = shmo.getRange(row, 7).getValue();
  shmo.getRange(row, 7).setValue(mating_num_mo + 1);
  
  //父親のprefixを取得し、「交配オス」シートの交配成績を変更。
  let fa_prefix = shmo.getRange(row, 9).getValue();
  let lastrow_fa = shfa.getLastRow();
  let prefix_list_fa = shfa.getRange(2, 4, lastrow_fa - 1, 1).getValues();
  for (let j=0; j < lastrow_fa - 1; j++) {
    if (prefix_list_fa[j][0] == fa_prefix) {
      let mating_num_fa = shfa.getRange(j + 2, 7).getValue();
      shfa.getRange(j + 2, 7).setValue(mating_num_fa + 1);
    }
  }
}

// 仔分け時にスプレッドシートを変更するための関数。
function discrimination_on_sheet_gs(input, num, mo_prefix) {
  let nowtime = new Date();
  let time = new Date(nowtime.getFullYear(), nowtime.getMonth(), nowtime.getDate());
  let ss = SpreadsheetApp.openById(id);
  let sh = ss.getSheetByName("仔分け前");
  let shdel = ss.getSheetByName("処分マウス");
  let shdelrow = shdel.getLastRow() + 1;
  let row = num + 1
  let lastcolumn = sh.getLastColumn();
  let dellist = sh.getRange(row, 1, 1, lastcolumn).getValues();
  let date = dellist[0][2];
  //オス、メスシート、処分マウスシートに追記。元データは列ごと消去。
  let sex = ["メス", "オス"]
  for (let k=0; k < sex.length; k++) {
    //オスorメスシートに書き込み
    if (input[k] > 0) {
      let shsex = ss.getSheetByName(sex[k]);
      let shsexrow = shsex.getLastRow() + 1;
      shsex.getRange(shsexrow, 1, 1, 2).setValues([[input[k], date]]);
      shsex.getRange(shsexrow, 3).setFormulaR1C1("=ROUNDDOWN((TODAY()-R" + shsexrow + "C2)/7)");
    }
    //処分マウスシートに書き込み
    if (input[k + 2] > 0) {
      shdel.getRange(shdelrow, 1, 1, 6).setValues([[input[k + 2], date, , "仔分け前(" + sex[k] + ")", "間引き", time]]);
      shdel.getRange(shdelrow, 3).setFormulaR1C1("=ROUNDDOWN((TODAY()-R" + shdelrow + "C2)/7)");
      for (let i=1; i <= lastcolumn - 4; i++) {
        shdel.getRange(shdelrow, i + 6).setValue(dellist[0][i + 2]);
      }
      shdelrow = shdelrow + 1;
    }
  }
  
  //prefixから母親の状態を変更。
  let shmo = ss.getSheetByName("交配メス");
  let lastrow_mo = shmo.getLastRow();
  let prefix_list = shmo.getRange(2, 4, lastrow_mo - 1, 1).getValues();
  for (let j=0; j < lastrow_mo - 1; j++) {
    if (prefix_list[j][0] == mo_prefix) {
      shmo.getRange(j + 2, 6).setValue("休憩");
    }
  }
  
  //「仔分け」シートから行を削除
  sh.deleteRows(row);
}

function mating_on_sheet_gs(which_mating, mo_prefix, fa_prefix) {
  let ss = SpreadsheetApp.openById(id);
  let shmo = ss.getSheetByName("交配メス");
  let shfa = ss.getSheetByName("交配オス");
  
  //which_matingがstart or finishでセルに入れる値を変える
  let which_cond_mo = "";
  let which_cond_fa = "";
    if (which_mating == "start") {
      which_cond_mo = "交配中";
      which_cond_fa = "交配中";
    } else if (which_mating == "finish") {
      which_cond_mo = "妊娠";
      which_cond_fa = "休憩";
    }
  
  //prefixからそれぞれのシートの「現在の状況」、「全交配数」を書き変える。同時に、「交配メス」シートに交配相手のprefixを追記する。
  let lastrow_mo = shmo.getLastRow();
  let prefix_list_mo = shmo.getRange(2, 4, lastrow_mo - 1, 1).getValues();
  let num_mo = 0;
  for (let j=0; j < lastrow_mo - 1; j++) {
    if (prefix_list_mo[j][0] == mo_prefix) {
      shmo.getRange(j + 2, 6).setValue(which_cond_mo);
      shmo.getRange(j + 2, 8).setValue(fa_prefix);
      num_mo = shmo.getRange(j + 2, 1).getValue();
      let num_all_mo = shmo.getRange(j + 2, 8).getValue();
      shmo.getRange(j + 2, 8).setValue(num_all_mo + num_mo);
    }
  }
  let lastrow_fa = shfa.getLastRow();
  let prefix_list_fa = shfa.getRange(2, 4, lastrow_fa - 1, 1).getValues();
  for (let j=0; j < lastrow_fa - 1; j++) {
    if (prefix_list_fa[j][0] == fa_prefix) {
      shfa.getRange(j + 2, 6).setValue(which_cond_fa);
      let num_all_fa = shfa.getRange(j + 2, 8).getValue();
      shfa.getRange(j + 2, 8).setValue(num_all_fa + num_mo);
    }
  }
}