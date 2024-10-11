/**
 * Copyright (c) 彬彬 2024
 * All rights reserved.
 * 
 * Created by: 彬彬
 * Contact: C110156220@nkust.edu.tw / robin92062574@gmail.com
 */

// 此腳本要由nkustic2019@gmail.com運行!

function startCountingStudent() {
  // 輸入學年度
  let year = "110"
  // 找到目標資料夾
  let folderId = findOrCreateTargetFolder(year)
  // 輸出所有資料夾以及裡面Google文件的ID
  let data = listAllSubfolderNamesAndDocId(folderId[1])

  // 開始數
  let studentInfo = checkingStudentGoogleDoc(data)

  // 建立統計完的試算表
  createSpreadsheetWithStudentDataInFolder_OutputCounting(
    year,
    folderId[1],
    studentInfo
  )
  
  
}

/**
 * 根據資料夾 ID 列出該資料夾下的所有子資料夾名稱，並輸出每個資料夾中的 Google 文件的 ID
 * 
 * @param folderId 資料夾 ID
 * 
 * @return {Object} 包含所有子資料夾名稱及其 Google 文件的 ID 和分享連結
 */
function listAllSubfolderNamesAndDocId(folderId) {
  try {
    // 根據資料夾 ID 取得主資料夾
    var folder = DriveApp.getFolderById(folderId);
    
    // 取得主資料夾下的所有子資料夾
    var subfolders = folder.getFolders();
    
    // 用來儲存子資料夾名稱、Google 文件 ID 和分享連結的陣列
    var folderNames = [], docIds = [], folderShareLinks = [];
    
    // 遍歷所有子資料夾
    while (subfolders.hasNext()) {
      var subfolder = subfolders.next();
      folderNames.push(subfolder.getName());
      folderShareLinks.push(subfolder.getUrl());

      // 取得該子資料夾下的所有 Google 文件
      var files = subfolder.getFilesByType(MimeType.GOOGLE_DOCS);
      
      // 檢查該子資料夾是否有 Google 文件
      if (files.hasNext()) {
        var file = files.next(); // 取得第一個 Google 文件
        docIds.push(file.getId());
        Logger.log("子資料夾: " + subfolder.getName() + " 中的 Google 文件 ID: " + file.getId());
      } else {
        Logger.log("子資料夾: " + subfolder.getName() + " 中沒有 Google 文件。");
        docIds.push(null); // 若無文件，則記錄 null
      }
    }


    // 回傳子資料夾名稱、Google 文件 ID 和分享連結
    return {
      folderNames: folderNames,
      docIds: docIds,
      shareLinks: folderShareLinks
    };
    
  } catch (error) {
    throw new Error("列出子資料夾名稱或取得 Google 文件 ID 時發生錯誤: " + error.message);
  }
}


function checkTableValues(googleDocId) {
  // 開啟 Google 文件
  var doc = DocumentApp.openById(googleDocId);
  var body = doc.getBody();
  var tables = body.getTables();
  let times = 0

  // 確認文件是否有表格
  if (tables.length > 0) {
    var table = tables[0]; // 假設序號表格是文件中的第一個表格

    for (var i = 6; i < table.getNumRows(); i++) { // 從第一列開始迭代 (跳過表頭)
      var row = table.getRow(i);
      var certItem = row.getCell(1).getText(); // 認證項目
      var description = row.getCell(2).getText(); // 簡要說明內容
      var date = row.getCell(3).getText(); // 取得日期


      if (certItem && description && date) {
        times = times + 1 
      }
    }
  } else {
    Logger.log("文件中沒有找到表格。");
  }
  Logger.log(times)
  return times
}

/**
 * 創建或找尋是否有新資料夾，尋找是否有新學年度的非正式課程資料夾
 * 
 * @param year 學年度
 * 
 * return [
 * 非正式課程存放地點資料夾ID,
 * 該目標資料夾ID
 * ]
 */
function findOrCreateTargetFolder(year) {

  try {
    // 資料夾名稱
    let parentFolderName = "智商系日四技【非正式課程學習護照】";
    let subFolderName = "日四技_"+year+"學年度入學";

    var parentFolder = DriveApp.getFoldersByName(parentFolderName)
    if (!parentFolder.hasNext()) {
      throw new Error('未找到名為'+parentFolderName+'的資料夾');
    }
    // 去下一個
    parentFolder = parentFolder.next()

    var subFolder = parentFolder.getFoldersByName(subFolderName);
    if (!subFolder.hasNext()) {
      // 若無，則建立資料夾
      subFolder = parentFolder.createFolder(subFolderName);
    }
    else{
      subFolder = subFolder.next()
    }
  
    return [
      parentFolder.getId(),
      subFolder.getId()
    ];

  } catch (error) {
    throw new Error('建立資料夾以及查詢特定資料夾時出現錯誤: ' + error.message);
  }
}

/**
 * 
 */
function checkingStudentGoogleDoc(folderAndDoc){
  let folderName = folderAndDoc.folderNames
  let docsId = folderAndDoc.docIds
  let studentValueResult = []
  let studentEmail = []
  let studentId = []

  for (let i = 0 ; i < folderName.length ; i++ ){
  
    // 確認表格填的次數
    let studentValue = checkTableValues(docsId[i])
    studentValueResult.push(studentValue)

    // 正規表達式
    let tmp_student = folderName[i].match(/C\d{9}/i)

    studentId.push(tmp_student[0])
    studentEmail.push(tmp_student[0]+"@nkust.edu.tw")

  }

  return {
    id :studentId,
    email:studentEmail,
    times:studentValueResult,
    shareLink:folderAndDoc.shareLinks
  }

}

/**
 * 
 */
function createSpreadsheetWithStudentDataInFolder_OutputCounting(
  year,
  targetfolderId ,
  studentInfo
  ) {

  // 4 個列表的數據
  var studentNumberList = studentInfo.id; // 學生編號列表
  var emailList = studentInfo.email; // 學生郵件列表
  var links = studentInfo.shareLink
  var times = studentInfo.times
  

  // 創建新的試算表
  var spreadsheet = SpreadsheetApp.create(year +"【非正式課程】初步統計名單");
  var sheet = spreadsheet.getActiveSheet();
  
  // 設置表頭
  sheet.getRange(1, 1).setValue("學號");

  // 本資料無姓名
  // sheet.getRange(1, 2).setValue("姓名");

  sheet.getRange(1, 3).setValue("郵件地址");
  sheet.getRange(1, 4).setValue("已填寫次數");
  sheet.getRange(1, 5).setValue("資料夾分享連結")
  

  // 將數據寫入試算表
  for (let i = 0; i < studentNumberList.length; i++) {
    sheet.getRange(i + 2, 1).setValue(studentNumberList[i]); // 學號
    sheet.getRange(i + 2, 3).setValue(emailList[i]); // 郵件地址
    sheet.getRange(i + 2, 4).setValue(times[i]); // 已填寫次數
    sheet.getRange(i + 2, 5).setValue(links[i]); // 資料夾分享連結
  }

  // 使用 DriveApp 移動試算表到特定資料夾
  var file = DriveApp.getFileById(spreadsheet.getId()); // 獲取試算表文件
  var folder = DriveApp.getFolderById(targetfolderId); // 獲取指定資料夾
  file.moveTo(folder); // 將試算表移動到指定資料夾

  Logger.log('試算表已創建並移動到指定資料夾，鏈接為：' + spreadsheet.getUrl());
}
