

/**
 * Copyright (c) 彬彬 2024
 * All rights reserved.
 * 
 * Created by: 彬彬
 * Contact: C110156220@nkust.edu.tw / robin92062574@gmail.com
 */

// 此腳本要由nkustic2019@gmail.com運行!

/**
 * 主執行
 */
function runFindFolderPermission(){
  // 這裡輸入第幾學年度
  let year = "110"

  // 找尋
  let folderId = findOrCreateTargetFolder(year)
  Logger.log("已定位到目標資料夾:"+"智商系日四技【非正式課程學習護照】"+"、"+"日四技_"+year+"學年度入學")

  // 列出該年度所有資料夾名稱
  let allFolderName = listAllSubfolderNames(folderId[1])
  Logger.log("已取得所有學生資料(請確保資料夾名稱正確)")

  // 處理並查詢，且輸出為試算表存放於該年度
  let studentInfo = deal_folderCheckAndCreatGoogleSheet_1(allFolderName)
  Logger.log("已確認將學生擁有資料夾權限")

  createSpreadsheetWithStudentDataInFolder_OutputUrls(
    year,folderId[1],studentInfo
  )
  Logger.log("已輸出學生名單以及其雲端分享連結，請去日四技_"+year+"學年度入學資料夾查看")
  
}

/**
 * 讀取指定資料夾的分享連結，並檢查是否有特定使用者的權限
 * 
 * @param folderId 要檢查的資料夾 ID
 * @param studentEmail 要檢查的使用者 Email
 * 
 * return {Object} 含有資料夾的分享連結與權限檢查結果
 */
function getFolderShareLinkAndCheckPermissions(folderId, studentEmail) {
  try {
    // 取得資料夾
    var folder = DriveApp.getFolderById(folderId);
    
    // 取得資料夾的分享連結
    var shareLink = folder.getUrl();
    
    Logger.log("資料夾分享連結: " + shareLink);

    // 檢查該學生是否擁有資料夾的編輯權限
    var hasPermission = false;
    var editors = folder.getEditors();
    
    for (var i = 0; i < editors.length; i++) {
      if (editors[i].getEmail() === studentEmail) {
        hasPermission = true;
        break;
      }
    }

    // 如果該學生沒有編輯權限，則賦予編輯權限
    if (!hasPermission) {
      folder.addEditor(studentEmail);
      Logger.log("已為 " + studentEmail + " 賦予編輯權限");
    }

    // 回傳分享連結與權限檢查結果
    return {
      shareLink: shareLink,
      hasPermission: hasPermission
    };
    
  } catch (error) {
    throw new Error("讀取資料夾分享連結或檢查權限時發生錯誤: " + error.message);
  }
}

/**
 * 根據資料夾 ID 列出該資料夾下的所有子資料夾名稱
 * 
 * @param folderId 資料夾 ID
 * 
 * return {Array} 包含所有子資料夾名稱的陣列
 */
function listAllSubfolderNames(folderId) {
  try {
    // 根據資料夾 ID 取得資料夾
    var folder = DriveApp.getFolderById(folderId);
    
    // 取得資料夾下的所有子資料夾
    var subfolders = folder.getFolders();
    
    // 用來儲存子資料夾名稱的陣列
    var folderNames = []; var allfolderId = []
    
    // 遍歷所有子資料夾，並將名稱加入陣列
    while (subfolders.hasNext()) {
      var subfolder = subfolders.next();
      folderNames.push(subfolder.getName());
      allfolderId.push(subfolder.getId())
    }

    // 輸出所有子資料夾名稱到日志
    Logger.log("子資料夾名稱: " + folderNames.join(", "));

    // 回傳子資料夾名稱陣列
    return [folderNames,allfolderId];
    
  } catch (error) {
    throw new Error("列出子資料夾名稱時發生錯誤: " + error.message);
  }
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
function deal_folderCheckAndCreatGoogleSheet_1(allFolderNames){

  let result = {
    id:[],
    email:[],
    shareLink:[]
  }
  
  for (let i = 0 ; i < allFolderNames[0].length ; i++){

    // 正規表達式，取C開頭的十碼(含C)
    const match = allFolderNames[0][i].match(/C\d{9}/i);

    if (match){
      // 學生學號的電子郵件
      const students_Email = match[0].charAt(0).toLowerCase() + match[0].slice(1)+"@nkust.edu.tw"

      let find_result = getFolderShareLinkAndCheckPermissions(
        // 取得ID
        allFolderNames[1][i],
        students_Email
      )

      // output
      result.email.push(students_Email)
      result.id.push(match[0])
      result.shareLink.push(find_result.shareLink)

    }else{
      Logger.log("此人可能不是大寫C導致出現錯誤，此人資料夾名稱:" + allFolderNames[i])
    }

  }
  Logger.log(result)

  return result
}

/**
 * 
 */
function createSpreadsheetWithStudentDataInFolder_OutputUrls(
  year,
  targetfolderId ,
  studentInfo
  ) {

  // 4 個列表的數據
  var studentNumberList = studentInfo.id; // 學生編號列表
  var emailList = studentInfo.email; // 學生郵件列表
  var shareLinkList = studentInfo.shareLink
  // 創建新的試算表
  var spreadsheet = SpreadsheetApp.create(year +"【非正式課程】「重新」寄信名單");
  var sheet = spreadsheet.getActiveSheet();
  
  // 設置表頭
  sheet.getRange(1, 1).setValue("學號");

  // 本資料無姓名
  // sheet.getRange(1, 2).setValue("姓名");

  sheet.getRange(1, 3).setValue("郵件地址");
  sheet.getRange(1, 4).setValue("分享連結");
  

  // 將數據寫入試算表
  for (let i = 0; i < studentNumberList.length; i++) {
    sheet.getRange(i + 2, 1).setValue(studentNumberList[i]); // 學號
    sheet.getRange(i + 2, 3).setValue(emailList[i]); // 郵件地址
    sheet.getRange(i + 2, 4).setValue(shareLinkList[i]); // 分享連結
  }

  // 使用 DriveApp 移動試算表到特定資料夾
  var file = DriveApp.getFileById(spreadsheet.getId()); // 獲取試算表文件
  var folder = DriveApp.getFolderById(targetfolderId); // 獲取指定資料夾
  file.moveTo(folder); // 將試算表移動到指定資料夾

  Logger.log('試算表已創建並移動到指定資料夾，鏈接為：' + spreadsheet.getUrl());
}
