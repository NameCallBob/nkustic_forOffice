/**
 * Copyright (c) 彬彬 2024 - 2026
 * 功能：不論是「簽到表模式」或「全體模式」，均執行深層內容核查並產出試算表報告。
 */

function startCheckRecord_all() {
 
  // ================= 1. 使用者設定區 =================
  
  let targetYear = "114"; // 目標學年度
  let targetTerm = "1";   // 目標學期
  
  let table_content = [
    "本系系周會",
    // targetYear + "-" + targetTerm + " 智慧商務系系周會", 
    targetYear + "-" + targetTerm + " 交通安全暨反詐騙講座", 
    "114.10.01",
    ""
  ];

  // 模式開關：true = 僅查核簽到表名單 / false = 查核資料夾內「所有」學生
  let isCheckListOnly = true; 
  let sheetFileName = "本次系周會簽到表"; 
  
  // 入學年度資料夾 (例如 111學年度入學)
  let checkFolderYear = "114"; 
  
  // =================================================

  // 找到目標資料夾
  let folders = findOrCreateTargetFolder(checkFolderYear);
  let parentFolderId = folders[0];
  let targetFolderId = folders[1];

  // 取得學生文件清單
  let allStudentData = listAllSubfolderNamesAndDocId(targetFolderId);

  // 獲取試算表物件 (不論模式為何，我們都需要一個地方輸出報告)
  let parentFolder = DriveApp.getFolderById(parentFolderId);
  let files = parentFolder.getFilesByName(sheetFileName);
  if (!files.hasNext()) throw new Error("找不到檔案：「" + sheetFileName + "」，請確保檔案位於主資料夾內。");
  
  let ss = SpreadsheetApp.open(files.next());
  let reportSheetName = targetYear + "-" + targetTerm + " 總檢核報告";
  let reportSheet = ss.getSheetByName(reportSheetName) || ss.insertSheet(reportSheetName);
  reportSheet.clear();
  
  let headers = [["學號/資料夾", "姓名(若有)", "檢核結果", "更新時間", "文件原始項目", "文件原始內容", "文件原始日期"]];
  reportSheet.getRange(1, 1, 1, 7).setValues(headers).setBackground("#cfe2f3").setFontWeight("bold");

  let checkList = [];

  // --- 邏輯分支：決定「待檢查名單」 ---
  if (isCheckListOnly) {
    Logger.log("--- 模式：簽到表名單查核 ---");
    let sourceSheet = ss.getSheets()[0]; 
    let sourceData = sourceSheet.getDataRange().getValues();
    for (let i = 0; i < sourceData.length; i++) {
      let sId = sourceData[i][0] ? sourceData[i][0].toString().trim() : "";
      let sName = sourceData[i][1] ? sourceData[i][1].toString().trim() : "";
      if (sId) checkList.push({ id: sId, name: sName });
    }
  } else {
    Logger.log("--- 模式：全體資料夾查核 (共 " + allStudentData.folderNames.length + " 個目錄) ---");
    // 將資料夾內的所有學生加入待查名單
    for (let i = 0; i < allStudentData.folderNames.length; i++) {
      checkList.push({ 
        id: allStudentData.folderNames[i], // 全體模式直接用資料夾名稱去對
        name: "" 
      });
    }
  }

  // --- 執行深層查核 ---
  let reportRows = [];
  checkList.forEach(student => {
    let result = performUltraFuzzyCheck(allStudentData, student.id, student.name, table_content, targetYear, targetTerm);
    reportRows.push([
      student.id, student.name, result.status, result.time,
      result.rawCat, result.rawDesc, result.rawDate
    ]);
  });

  // --- 輸出結果至試算表 ---
  if (reportRows.length > 0) {
    reportSheet.getRange(2, 1, reportRows.length, 7).setValues(reportRows);
    reportSheet.autoResizeColumns(1, 7);
    
    // 簡單的美化：將「缺漏」標記為紅色字體
    let range = reportSheet.getRange(2, 3, reportRows.length, 1);
    let rules = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("❓")
      .setFontColor("#FF0000")
      .setRanges([range])
      .build();
    let sheetRules = reportSheet.getConditionalFormatRules();
    sheetRules.push(rules);
    reportSheet.setConditionalFormatRules(sheetRules);
  }

  Logger.log("🎉 全體檢核完成！請查看分頁：「" + reportSheetName + "」");
}

/**
 * 核心：超強模糊比對邏輯 (保持不變，但確保讀取每個文件)
 */
function performUltraFuzzyCheck(allData, sId, sName, targetArr, tYear, tTerm) {
  let fNames = allData.folderNames;
  let dIds = allData.docIds;
  let idx = -1;
  let now = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm");

  // 1. 搜尋學生資料夾
  for (let i = 0; i < fNames.length; i++) {
    if (fNames[i].includes(sId) || (sName && fNames[i].includes(sName))) {
      idx = i; break;
    }
  }

  if (idx === -1) return { status: "❌ 找不到資料夾", time: "", rawCat: "", rawDesc: "", rawDate: "" };
  if (!dIds[idx]) return { status: "⚠️ 找不到護照文件", time: "", rawCat: "", rawDesc: "", rawDate: "" };

  try {
    let doc = DocumentApp.openById(dIds[idx]);
    let tables = doc.getBody().getTables();
    if (tables.length === 0) return { status: "🛑 無表格內容", time: now, rawCat: "", rawDesc: "", rawDate: "" };
    
    let table = tables[0];
    let bestMatch = null;
    let targetClean = targetArr[1].replace(/\s+/g, "").replace(/週/g, "周");

    for (let r = 5; r < table.getNumRows(); r++) {
      let row = table.getRow(r);
      if (row.getNumCells() < 3) continue;

      let rCat = row.getCell(1).getText().trim();
      let rDesc = row.getCell(2).getText().trim();
      let rDate = row.getCell(3).getText().trim();

      let cDesc = rDesc.replace(/\s+/g, "").replace(/週/g, "周").replace(/（/g, "(").replace(/）/g, ")");
      
      let yearMatch = cDesc.match(/\d{3}/);
      if (yearMatch && yearMatch[0] !== tYear) continue;

      if (cDesc.includes(targetClean) || targetClean.includes(cDesc)) {
        if (cDesc.length > 5) return { status: "✅ 已填寫", time: now, rawCat: rCat, rawDesc: rDesc, rawDate: rDate };
      }

      let hasYear = cDesc.includes(tYear);
      let hasTerm = new RegExp(tTerm + "|-" + tTerm + "|\\(" + tTerm + "\\)").test(cDesc);
      let hasKey = /系周會|智商|智慧商務/.test(cDesc);

      if (hasYear && hasKey && hasTerm) {
        bestMatch = { status: "⚠️ 模糊匹配(待核對)", time: now, rawCat: rCat, rawDesc: rDesc, rawDate: rDate };
      }
    }
    return bestMatch || { status: "❓ 缺漏", time: now, rawCat: "", rawDesc: "", rawDate: "" };
  } catch (e) {
    return { status: "🛑 讀取錯誤", time: "", rawCat: "", rawDesc: "", rawDate: "" };
  }
}

// ================= 輔助函式 (保持不變) =================

function listAllSubfolderNamesAndDocId(folderId) {
  var folder = DriveApp.getFolderById(folderId);
  var subfolders = folder.getFolders();
  var folderNames = [], docIds = [];
  while (subfolders.hasNext()) {
    var subfolder = subfolders.next();
    folderNames.push(subfolder.getName());
    var files = subfolder.getFilesByType(MimeType.GOOGLE_DOCS);
    docIds.push(files.hasNext() ? files.next().getId() : null);
  }
  return { folderNames: folderNames, docIds: docIds };
}

function findOrCreateTargetFolder(year) {
  let parentName = "智商系日四技【非正式課程學習護照】";
  let subName = "日四技_"+year+"學年度入學";
  var parentFolder = DriveApp.getFoldersByName(parentName).next();
  var subFolder = parentFolder.getFoldersByName(subName);
  var targetId = subFolder.hasNext() ? subFolder.next().getId() : parentFolder.createFolder(subName).getId();
  return [parentFolder.getId(), targetId];
}
