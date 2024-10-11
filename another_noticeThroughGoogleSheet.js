/**
 * Copyright (c) 彬彬 2024
 * All rights reserved.
 * 
 * Created by: 彬彬
 * Contact: C110156220@nkust.edu.tw / robin92062574@gmail.com
 */

// 由於寄信都是用公務信箱寄信。請使用vhoffice01@nkust.edu.tw啟動本腳本!

// 使用情境：需共用資料夾雲端(已經設定好權限)並且統整為一個試算表(Excel)

// 若需要內文需變動在請到sendEmailTemplate進行變更

/**
 * 發送郵件通知，根據指定的檔案名稱取得試算表，並根據資料寄送郵件
 * @param {string} fileName - 試算表檔案名稱
 * @param {string} year - 年度資訊，用於郵件主旨顯示
 */
function sendEmailNotification(fileName, year) {
  // 取得雲端硬碟中名稱為 fileName 的試算表檔案
  var files = DriveApp.getFilesByName(fileName);
  
  // 如果找到檔案，則繼續處理
  if (files.hasNext()) {
    var file = files.next();
    var spreadsheet = SpreadsheetApp.open(file);  // 開啟試算表
    var sheet = spreadsheet.getActiveSheet();     // 取得目前的工作表

    // 取得所有資料列
    var dataRange = sheet.getDataRange();
    var data = dataRange.getValues();  // 以陣列形式取得所有資料
    
    // 從第2行開始迭代資料（忽略標題行）
    for (var i = 1; i < data.length; i++) {
      var studentId = data[i][0];  // 學號 (A欄)
      var email = data[i][1];      // 電子郵件 (B欄)
      var cloudLink = data[i][2];  // 雲端連結 (C欄)
      
      // 呼叫 sendEmailTemplate，將學號、郵件、雲端連結及年度資訊傳入
      sendEmailTemplate(email, cloudLink, year);
    }
    
    // 記錄成功發送的訊息
    Logger.log("所有通知信件已發送！");
  } else {
    // 如果找不到檔案，記錄錯誤訊息
    Logger.log("找不到名稱為 '" + fileName + "' 的試算表檔案。");
  }
}

/**
 * 寄送郵件模板，動態生成郵件主旨及內容
 * @param {string} email - 收件人的電子郵件
 * @param {string} cloudLink - 雲端連結
 * @param {string} year - 年度資訊，用於郵件主旨顯示
 */
function sendEmailTemplate(email, cloudLink, year) {
  // 設定郵件主旨，將年度資訊動態插入
  var subject = "【智商系辦通知】" + year + "級實務專題成果報告及Demo影片雲端連結已開通";
  
  // 設定郵件內容，根據提供的雲端連結和模板格式生成郵件內文
  var body = "Dear 專題小組長：\n\n" +
             "　如題，雲端連結如下。\n" +
             "　" + cloudLink + "\n\n" +
             "　再請您於期限內繳交資料，謝謝！\n" +
             "　期限及相關規定請參照line群組近期傳送的檔案，\n\n" +
             "　如有問題再請您回信告知，如雲端權限未開啟的問題。\n\n" +
             "　　　　　　　By 系辦工讀生";

  // 使用 MailApp.sendEmail 來寄送郵件
  MailApp.sendEmail(email, subject, body);
}

/**
 * 執行主函數，設定檔案名稱及年度後，觸發郵件發送過程
 */
function runEmailSent() {
  // 試算表檔案名稱
  let fileName = "寄信統計表"; 
  // 年度資訊
  let year = "110";

  // 呼叫發送通知函數，傳入檔案名稱和年度資訊
  sendEmailNotification(fileName, year);
}
