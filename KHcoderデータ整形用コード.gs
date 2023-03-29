const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName('シート1');
const ui = SpreadsheetApp.getUi();

function copy_and_paste() {
  //分野と年代の決定
  const sobj = ui.prompt('分野名を入力').getResponseText();
  const start = ui.prompt('始まりの年代を入力').getResponseText();
  const end = ui.prompt('終わりの年代を入力').getResponseText();
  let url = `https://cir.nii.ac.jp/dissertations?q=${sobj}&from=${start}&until=${end}&languageType=en&count=200&sortorder=0`;
  
  //スクレイピング
  let res = UrlFetchApp.fetch(url).getContentText();
  let getDatas = Parser.data(res).from('"sendEvent(\'タイトル\')">').to('</a>').iterate();
  
  //スクレイピングしたものをシートに追加
  for(data of getDatas){
    datas = [sobj, start, data.toLowerCase()];
    sheet.appendRow(datas);
  }
}

function sheet_to_excel() {
  
  const sheetID = "";
  const folderID = "";
  const xlsxName = "CiNii博論タイトル分析";

  const xlsxFile = excelExport(sheetID, xlsxName);
  
  gdriveUpload(folderID, xlsxFile);

}

function excelExport(sheetID, fileName){
  let today = Utilities.formatDate(new Date(),"JST","yyyyMMdd");
  let xlsxName = `${today}_${fileName}.xlsx`;
  let options = {
    method: 'get',
    headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    muteHttpExceptions: true
  };
  let Url = `https://docs.google.com/feeds/download/spreadsheets/Export?key=${sheetID}&exportFormat=xlsx`;
  let xlsxFile = UrlFetchApp.fetch(Url,options).getBlob().setName(xlsxName);
  return xlsxFile
}

function gdriveUpload(folderID, uploadFile){
  let folder = DriveApp.getFolderById(folderID);
  let drive_file = folder.createFile(uploadFile);
  return drive_file;
}