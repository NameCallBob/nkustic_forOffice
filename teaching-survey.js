/**
 * Copyright (c) 彬彬 2024
 * All rights reserved.
 * 
 * Created by: 彬彬
 * Contact: C110156220@nkust.edu.tw / robin92062574@gmail.com
 */


//以下其設計邏輯，參考郭粟閣/林祜緻之原先著作，將其原代碼進行優化及註解

// 授課意見調查表之路徑設計為
// / -> /系上表單及文件 -> /授課意見調查表 -> /新增的新學年授課意見調查表

// 若要修改其路徑，請詢問彬彬或相關維護人員

// 1.使用前請先確保你有去學校取得所有智慧商務系課程
// 並儲存於/系上表單及文件/授課意見調查表/課程統整(Google試算表)
// 之後先進行資料的初步處理，將不相關課程進行刪除。

// 筆記:何謂不相關課程:
// 實用英文(一、二、三、四)、中文閱讀與表達、資料分析套裝軟體、體育(一、二)。

// 2.修改 init的變數 year是哪個學年哪個學期(113(1)或是115(2))
// 以及依照情況分成幾個step
// (1,2一定可以跑，但你改更大之後出錯請通知彬彬或相關維護人員)

// 3.點選上方的執行並且選擇init

// 註:如果你不小心動到了程式碼還不會跑，以下是原代碼的github
// https://github.com/NameCallBob/nkusticsa_forOffice


/**
 * 建立授課意見調查表之資料表
 * return {Array} 
 * 
 * 值如下
 * [
 * 系上表單及文件的資料夾ID,
 * 授課意見調查表的資料夾ID,
 * 學年授課意見調查表的資料夾ID
 * ]
 */
function createParentFolder(parentFolderName="系上表單及文件",schoolyear){
  function createSchoolYearFolder(schoolyear) {
    var folder = DriveApp.getFoldersByName("系上表單及文件").next() ; folder = folder.getFoldersByName("授課意見調查表").next() ; 
    let folderId = folder.getId()
    var foldername = schoolyear + "授課意見調查表"; 
    var existingFolder=folder.getFoldersByName(foldername)

    if (!existingFolder.hasNext()){//如果沒有同名表單，則創建新表單
        folder.createFolder(foldername);  
        Logger.log("創建學年度資料夾：" + foldername)
      }
      else{
        Logger.log("資料夾已存在，跳過創建"+  foldername)
      }
      return [folderId,folder.getFoldersByName(foldername).next().getId()]
  }
  // 檢查父資料夾是否存在，如果不存在就建立
    var parentFolder = DriveApp.getFoldersByName(parentFolderName);
    if (!parentFolder.hasNext()) {
      parentFolder = DriveApp.createFolder(parentFolderName);
      Logger.log("已建立父資料夾：" + parentFolder.getName());
    } else {
      parentFolder = parentFolder.next();
      Logger.log("父資料夾存在：" + parentFolder.getName());
    }
    var childfolder = createSchoolYearFolder(schoolyear)
    
    return [parentFolder.getId(),childfolder[0],childfolder[1]]
}

/**
 * 得取所有課程名稱(需人工審核該課程是否需建立授課意見調查，例如:體育不用)
 * 檔案位置:/系上表單及文件/授課意見調查表/課程統整(試算表)
 * 
 * @param {String} ParentFolderId - 授課意見調查表的資料夾ID
 * 
 * return Array 
 */
function getClassName(ParentFolderId){

  // 根據資料夾 ID 和檔名取得試算表
  function getSpreadsheet(folderId, fileName) {
    var folder = DriveApp.getFoldersByName("授課意見調查表").next();
    var files = folder.getFilesByName(fileName);
    if (files.hasNext()) {
      var file = files.next();
      var spreadsheet = SpreadsheetApp.openById(file.getId());
      return spreadsheet;
    } else {
      return null;
    }
  }
  //偵測資料
  try{
      var folder = DriveApp.getFolderById(ParentFolderId);
  }catch{
    Logger.log("找不到/系上表單及資料/授課意見調查表/，建立資料夾")
    ParentFolderId = createParentFolder()
    var folder = DriveApp.getFolderById(ParentFolderId);
    // 移動新試算表至指定資料夾
    var file = DriveApp.getFileById(newSpreadsheet.getId());
    folder.createFile(file);
  }
  
  //讀取試算表
  var spreadsheet = getSpreadsheet(folder, "課程統整");
  if (spreadsheet) {
    // 如果試算表存在，讀取欄位 B 的數據
    try{
      var sheet = spreadsheet.getSheetByName("工作表1");
    }catch{
      throw Error("資料表名稱非「工作表1」，請手動更改")
    }

    var lastRow = sheet.getLastRow();
    if (lastRow > 1) { // 確保至少有一行資料（假設第一行是標題）
      classId = sheet.getRange("A1:A"+lastRow).getValues();
      teachClass_1 = sheet.getRange("E1:E" + lastRow).getValues();
      teachClass_2 = sheet.getRange("F1:F"+ lastRow).getValues();
      className = sheet.getRange("H1:H"+ lastRow).getValues();
      classLoc = sheet.getRange("N1:N"+ lastRow).getValues();
      teacherName = sheet.getRange("M1:M"+ lastRow).getValues();
      return [teachClass_1,teachClass_2,className,classLoc,teacherName,classId]
    }
    else{
      throw Error("該資料表未有資料")
    }
    return null

  } else {
    //建立新試算表後移動至特定位置
    var newSpreadsheet = SpreadsheetApp.create("課程統整");
    var file = DriveApp.getFileById(newSpreadsheet.getId());
    file.moveTo(folder)
    Logger.log("試算表已建立，請去校務系統統整本學期的課表");
    Logger.log("試算表位置：/系上表單及文件／授課意見調查表/課程統整（試算表）")
    Logger.log("資料新增後請移至授課意見調查表（資料夾）")
    Logger.log("若發現該資料在/系上表單及文件，請協助移到授課意見調查表")
    throw  Error("終止運行")

  }

 
}

/**
 * 確認並建立表單的放置地點，若已存在則只放置學年資料夾
 * 
 * @param {String} yearfolderName  -  學年資料夾
 * return {String}  -  資料夾ID
 */
function createFolders(yearfolderName) {
  let 
  var parentFolderName = "系上表單及文件";
  var childFolderName = yearfolderName;

  // 檢查父資料夾是否存在
  var parentFolder = DriveApp.getFoldersByName(parentFolderName);
  if (!parentFolder.hasNext()) {
    parentFolder = DriveApp.createFolder(parentFolderName);
    Logger.log("已建立父資料夾：" + parentFolder.getName());
  } else {
    parentFolder = parentFolder.next();

  }
  if (parentFolder.getFoldersByName("授課意見調查表".hasNext())){
  // 在父資料夾中檢查子資料夾是否存在，如果不存在就建立
      var childFolder = parentFolder.getFoldersByName(childFolderName);
      var childFolderId;
      if (!childFolder.hasNext()) {
        childFolder = parentFolder.createFolder(childFolderName);
        Logger.log("已建立子資料夾：" + childFolder.getName());
        childFolderId = childFolder.getId();
      } else {
        childFolder = childFolder.next();
        Logger.log("子資料夾存在：" + parentFolder.getName());
        childFolderId = childFolder.getId();
      }
      Logger.log("子資料夾的ID：" + childFolderId);
      return childFolderId
    }
  else{
    throw Error("路徑有問題，未找到已存在的授課意見調查表，請確認系上表單及文件是否存在")
  }
  
}


/**
 * 建立問卷
 * @param {String} data  -  課程名稱
 * @param {String} FolderId - 放置資料夾ID
 * 變數 data 為 ["課程名稱","授課班級","授課老師","課程編號","授課地點"]
 */
function createNewForm(data,FolderId){
    // 指定Google雲端硬碟位置的文件夾ID
    var folderId = FolderId;


    /**
     * (重複使用)
     *  在Google表單中創建五分量表之選項
     * @param {} form  - Google表單之物件
     * @param {String} title - 題目名稱
     */
    function createScaleQuestion(form, title) {
      var item = form.addMultipleChoiceItem();
      item.setTitle(title).setChoiceValues(["非常同意", "同意", "普通", "不同意", "非常不同意"]);
      item.setRequired(true)
    }


    // 在指定的文件夾中建立一個新的Google表單
    var folder = DriveApp.getFolderById(folderId);
    // 自定義表單簡介
    let descrip = "同學您好!\n為增進良好的教學體驗，請填寫教學意見調查表!\
                  提供您對於本堂課的想法，請如實填寫且避免攻擊性詞語，\
                  \n感謝您的填寫，以下是該堂課的基本資訊。\
                  \n授課地點："+data[4]+"\
                  \n課程編號："+data[3]+"\
                  \n課程名稱："+data[0]+"\
                  \n授課班級："+data[1]+"\
                  \n授課老師："+data[2]+"\
                  \n\n如有任何有關問卷上的問題，請洽智商系系辦!"

    var form = FormApp.create(data[3]+"_"+data[1]+"_"+data[0]).setDescription(descrip);

    //問題設置
    createScaleQuestion(form, "本課程的授課方式能幫助我有效的學習課程內容");
    createScaleQuestion(form, "本課程的授課方式能激勵同學學習");
    form.addParagraphTextItem().setTitle("您對本課程的授課方式的意見");
    createScaleQuestion(form, "您對本課程的學習成效感到");
    form.addParagraphTextItem().setTitle("承上題,原因為何?");
    form.addParagraphTextItem().setTitle("您對本課程評量的方式");

    //表單設定
    form.setAcceptingResponses(true); //開放問答
    form.setCollectEmail(false);  //不收集電子郵件
    // form.setRequireLogin(false)
    // 因kuasmis此系辦帳號非機構帳號，上方不須打
    form.setCustomClosedFormMessage("感謝您的填寫!")
    let formId = form.getId()
    var formFile = DriveApp.getFileById(formId);
    formFile.moveTo(folder)
    let formUrl = form.getPublishedUrl();

    return formUrl
  
}

/** 
 * 建立所有課程的表單資料
 * @params {Array} data  - 如下
 * data = [開課班級,合開課程班級,課程名稱,開課地點,開課導師]
 * return {Array}  - 如 [className,classTeacher,classWho,classLoc,classId]
 */
function createAllClass(data){

  /**
   * 對於有相同名稱的課程進行判斷是否生成表單
   * @param {Array} 
   */
  function dealSameNameLesson(){
    // 先針對將有合併班級的課程
  }
  // 資料處理

  // 在這裡填寫不要生成的課程
  const no = ["服務學習","體育","實務專題","實習","中文閱讀與表達","服務設計與企劃執行"
  ]
  //變數;className:課程名稱、classTeacher:授課老師、classLoc:上課地點、classWho:授課班級
  let className = new Array() ; let classTeacher = new Array() ; let classLoc = new Array() ; let classWho = new Array() ; let classId = new Array();
  /**確認字串是否含某些字 */
  function checkIfIncludesAnyElement(word, array) {
    return array.some(element => word.includes(element));
  }
  for (let i = 0 ; i < data[0].length ; i++){
    //判斷課程是否需編寫授課意見調查
    if (checkIfIncludesAnyElement(data[2][i][0],no)) {
      Logger.log("課程:"+data[2][i][0]+"跳過")
      continue
    }
    else{
      if (data[1][i].length != 0 && data[1][i][0] != ""){
        //若有合開班級，要將其班級合併
        res_class = data[0][i][0]+"、"+data[1][i][0]
      }
      else{
        res_class = data[0][i][0]
      }
    className.push(data[2][i][0]) ; classTeacher.push(data[4][i][0]) ; classWho.push(res_class) ; classLoc.push(data[3][i][0]);classId.push(data[5][i][0])
    Logger.log ("第"+i+"堂課，\
      編號"+data[5][i][0]+"，\
      名稱:"+data[2][i][0]+"，\
      授課班級:"+res_class+"，\
      開課教授:"+data[4][i][0]+"，\
      位置:"+data[3][i][0]+"。")
    }

  }
  return [className,classTeacher,classWho,classLoc,classId]
}

/**
 * 生成各課程表單後最後放入Word 
 * 
 * 所有課數：data[0].length
 * 
 * @param {Integer} step -  生成表單要分成幾步驟
 * @param {String} folderId -  資料夾ID
 * @param {Array} data -  表單的資訊
 * 
 * return null 
 */
function createForm(step,folderId,data){
  let urls = [] ; let num , end
  // 由於有函數執行上限，設為兩個step
  if (step == 1){
    num = 0 ; end = data[0].length/2
  }
  else if (step == 2){
    num = data[0].length/2+1 ; end = data[0].length-1
  }
  // 執行表單生成並儲存表單連結
  for ( let i = num ; i <= end ; i++){
    tmp_data = [data[0][i],data[2][i],data[1][i],data[4][i],data[5][i]]
    let url = createNewForm(tmp_data,folderId)
    urls.push(url)
  }
  Logger.log("表單建置完畢!開始QrCode生成");

  let doc = DocumentApp.create('授課意見QrCode_'+step);
  var body = doc.getBody(); let urlnum = 0

  for (let i = num ; i <= end ; i++){
    var course = data[4][i]+"_"+data[0][i]+"_"+data[1][i];
    if (i > 1) { 
      body.appendPageBreak();
    }
    // Word內容
    body.appendParagraph(course).setHeading(DocumentApp.ParagraphHeading.HEADING1);   
    body.appendParagraph(data[2][i]).setHeading(DocumentApp.ParagraphHeading.HEADING2) 
    try{
    //第一個生成QrCode的服務
    var imageUrl = "https://chart.googleapis.com/chart?chs=500x500&cht=qr&chl=" + encodeURIComponent(urls[urlnum]);
    var imageBlob = UrlFetchApp.fetch(imageUrl).getBlob();
    }catch{
      //第二個外部套件的生成服務  
      var imageUrl = "https://api.qrserver.com/v1/create-qr-code/?size=500x500&data=" + encodeURIComponent(urls[urlnum]);
      var imageBlob = UrlFetchApp.fetch(imageUrl).getBlob();
    }
    body.appendImage(imageBlob);
    urlnum = urlnum + 1
  }
  DriveApp.getFileById(doc.getId()).moveTo(DriveApp.getFolderById(folderId))
  
}


/**主執行
 * 邏輯流程：建立新學年度資料夾 -> 讀取課程資料 -> 建立表單 -> 建立Word 
 * 基於GoogleBot只有6分鐘的處理時間，所以分成兩步驟進行生成
 */
function init(){
  try{
  var year = "113(1)" ; step = 2
  Logger.log("開始")
  folderId = createParentFolder("系上表單及文件",year)
  Logger.log("讀取課程資料")
  data = getClassName(folderId[1])
  Logger.log("讀取成功，開始建置表單")
  data = createAllClass(data)
  Logger.log("開始建置表單")
  createForm(step,folderId[2],data)
  Logger.log("結束，授課意見調查整體生成完畢")

  }catch(error){
    throw Error("shit!出問題摟，錯誤訊息是"+error)
  }
  
}
