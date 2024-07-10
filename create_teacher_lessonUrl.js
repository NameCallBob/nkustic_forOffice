/**
 * Copyright (c) 彬彬 2024
 * All rights reserved.
 * 
 * Created by: 彬彬
 * Contact: C110156220@nkust.edu.tw / robin92062574@gmail.com
 */


// 在使用此程式碼之前請記得要先執行表單製作且確認init裡面的schoolyear是否是你想要輸出的學年度。
// 此程式碼會因為校務系統的授課清單錯誤，而有錯誤，絕對不是程式碼的錯!
// 若出現以上問題要進行手動調整，(什麼問題?授課老師打錯、名稱錯誤、教室空白)

// 若出現程式碼邏輯問題，請聯絡彬彬或其他維護人員。

/**
 * 鎖定資料夾位置
 * 
 * @param {String} parentFolderName - 父資料夾(指雲端的第一層)
 * @param {String} schoolyear - 學年和學期(e.g. 113(1) )
 * 
 * 依照以上的範例，通常會鎖定113(1)授課意見調查表
 * return {String} folderId - 該年度授課意見調查表的ID 
 */
function getFolderloc(parentFolderName="系上表單及文件",schoolyear){
  // 註:若是這裡出錯，請記得檢查變數(parentFolder,parentNextFolder,oldFolder,targetFolder)是否正確
  // 會混淆的變數如下。
  // parentFolder -> 父資料夾
  // parentNextFolder -> 系上表單及文件 或 父資料夾(若parentFolderName為null)
  // oldFolder -> 授課意見調查表
  // targetFolder -> 使用者要找的資料夾(e.g.112(2)授課意見調查表)

  // 先判斷是否有父資料夾
  if (parentFolderName == null){
    Logger.log("疑似無父資料夾(原:系上表單及文件)")  
  }
  else{
  // 尋找資料夾
  let parentFolder = DriveApp.getFoldersByName(parentFolderName);
  if (!parentFolder.hasNext()){
    throw Error(`找不到${parentFolderName}，有改路徑需聯絡相關維護人員`)
     }
  }
  let parentNextFolder
  try{
    parentNextFolder = parentFolder.next()
  }
  catch(err){
    // 若使用者未輸入父資料夾，將直接尋找授課意見調查表的資料夾
    parentNextFolder = DriveApp
  }



    // 尋找下一個資料夾
    let oldFolder = parentNextFolder.getFoldersByName("授課意見調查表")
    if (!oldFolder.hasNext()){
      throw Error("找不到授課意見調查表，路徑有改要連絡相關維護人員喔")
    }
    let targetFolder = oldFolder.next()



    let targetFolderName = schoolyear + "授課意見調查表"
    // 尋找下一個資料夾
    if(!targetFolder.getFoldersByName(targetFolderName).hasNext()){
        throw Error(`\
        找不到該年度授課意見調查表，\
        您要找的資料夾為${targetFolderName}\
        ，路徑有改要連絡相關維護人員喔`
      )
    }
    else{
      targetFolder = targetFolder.getFoldersByName(targetFolderName).next()
      return targetFolder.getId()
    }
}


/**
 * 透過資料夾ID取得該資料的所有檔案
 * 透過迴圈得取表單ID
 * 
 * @params {String} FolderId - 資料夾ID
 * 
 * return ['表單ID','表單ID','表單ID','表單ID'.....] 
 */
function getAllForm(FolderId) {
    var folder = DriveApp.getFolderById(FolderId);
    var files = folder.getFiles();

    // 儲存表單ID
    let result = []

    Logger.log("----開始尋找!----")
    while (files.hasNext()) {
    var file = files.next();
    var mimeType = file.getMimeType();
    
    // 檢查是否為 Google 表單
    if (mimeType === MimeType.GOOGLE_FORMS) {
      Logger.log('Google 表單名稱：' + file.getName());
      Logger.log('Google 表單 ID：' + file.getId());
      // 如果只需要一個表單 ID，可以在這裡直接回傳 file.getId();
      result.push(file.getId()) ;
    }
  }
  Logger.log("----尋找結束-----")

  // 如果資料夾內沒有 Google 表單
  if (result.length === 0){
      throw Error(
        "在這個資料夾中找不到任何 Google 表單，您有先表單製作嗎?"
      )
  }
  return result;
}


/**
 * 透過formID拿取該表單的基本資訊以及簡介裡的資訊
 * 
 * @param {String} id - Google表單ID
 * 
 */
function deal_formInfo(id){

  // ------------------func1-----------------------
  /**
   *  由於表單簡介格式固定
   *  利用正規化表達式
   *  取得老師名稱、及上課班級
   *  註:
   * 
   *  @param {String} Content - 授課意見調查表_表單簡介
   * 
   *  return {Array} 長這樣['授課班級','授課老師','授課教室','課程名稱]
   *
   */
  function getClassInfo(Content){

    // 使用正規表達式來擷取授課班級和授課老師
    var regexClassLoc = /授課教室：(.+?)\n/;
    var regexForWho = /授課班級：(.+?)\n/;
    var regexTeacher = /授課老師：(.+?)\n/;
    var regexLessonName = /課程名稱：(.+?)\n/;

    var forwhoMatch = formDescription.match(regexForWho);
    var teacherMatch = formDescription.match(regexTeacher);
    var classLocMatch = formDescription.match(regexClassLoc);
    var lessonNameMatch = formDescription.match(regexLessonName);

    // 將其儲存為object並檢查是否存在
    var courseInfo = {
      who: forwhoMatch ? forwhoMatch[1] : "未知",
      teacher: teacherMatch ? teacherMatch[1] : "未知",
      classLoc: classLocMatch ? classLocMatch[1] : "未知",
      lessonName: lessonNameMatch ? lessonNameMatch[1] : "未知"
    };

    return [

      courseInfo.who,
      courseInfo.teacher,
      courseInfo.classLoc,
      courseInfo.lessonName

    ];
  }

  // -------------------------------------------
  /**
   * 測試:輸出結果
   * @param {String} formTitle - 表單標頭
   * @param {String} formDescription - 表單簡介
   * @param {String} formUrl - 表單連結
   * @param {String} formDesData - 表單簡介拿取的資料 ref -> getClassInfo
   */
  function output_logTest(formTitle,formDescription,formUrl,formDesData){
    Logger.log('表單標題: ' + formTitle);
    Logger.log('表單簡介: ' + formDescription);
    Logger.log('表單連結: ' + formUrl);
    Logger.log('courseInfo: ' + formDesData) ;
  }

  // -------------------------------------------

  // run 
  var form = FormApp.openById(id); 
  // 拿取表單資訊
  var formTitle = form.getTitle();
  var formDescription = form.getDescription();
  var formUrl = form.getPublishedUrl();

  let content = getClassInfo(formDescription);

  // 測試!!輸出使用，若需要再解除註解
  // output_logTest(formTitle,formDescription,formUrl,content)
  
  // 將原結果再添加表單連結、表單標頭(若有需要使用)進去
  content.push(formUrl)
  content.push(formTitle)

  return content
}

/**
 * 處理取得每個表單的實際資訊後，彙整每位老師所上的課以及資訊
 * 並且使用output_AS_Word()將結果用Word儲存
 * 註:裡面含邏輯func -> check、
 * 
 * @param {Array} all_formId 該資料夾所有的表單ID
 * 
 * return null
 */
function arrange_all_data(all_formId,targetFolder){

  // ------------------func1-----------------------
  /**
   * 得取的資料利用迴圈，進行整理
   * @param {Array} all_formId - 該資料夾所有的表單ID(Array)
   * return 結果
   */
  function check(all_formId){
    // 結果
    let result = {

    // 授課地點 (暫停使用)
    classLoc : [],

    // 授課老師
    teacher : [],

    // 授課名稱
    lessonName : [],

    // 授課班級
    forwho : [],

    // 表單連結
    url : []

    }
    for (let j = 0 ; j < all_formId.length ; j++){
      // 利用迴圈確保每個表單可被搜尋進去
      tmp_res = deal_formInfo(all_formId[j])

      // 將結果用result進行儲存
      result.forwho.push(tmp_res[0])
      result.teacher.push(tmp_res[1])
      result.lessonName.push(tmp_res[3])
      result.classLoc.push(tmp_res[2])
      result.url.push(tmp_res[4])

      }
    return result;
  }
  // -------------------------------------------

  // ------------------func2-----------------------
  /**
   * 將已經讀取的表單資訊進行整理為
   * 一位老師有多少的課
   * 
   * @param {object} data - 表單取得內容
   * 
   * return {object}  - 老師授課資料
   */
  function formDataToInfo(data){
    /**
     * 重複使用:找出value所對應的多個index
     * 
     * @param {Array} need_to_find - 需要進行找尋的陣列
     * @param {String} value - 要找的值
     * 
     * return {Array} indexes - 找到的索引
     */
    function getIndexes(need_to_find,value){
      let indexes = []
      
      idx = need_to_find.indexOf(value)
      while (idx !== -1) {
        indexes.push(idx);
        idx = need_to_find.indexOf(value, idx + 1);
      }

      return indexes
    }


    // 取得老師
    let teacherList = [... new Set(data.teacher)]

    // 結果
    // 預覽結果為["老師姓名",[ ["班級",['','','']] , ["班級",['','','']] ]]
    var result = []

    // 將所有老師透過迴圈整理
    for (let i = 0 ; i < teacherList.length ; i++){
      // 此邏輯為由於在先前資料整理(function check)有確保相同的index是可以得取該表單的內容
      // 所以先透過index拿取後再進行整理

      let indexesFind = getIndexes(data.teacher,teacherList[i])

      // 暫存找到的結果
      let tmp_TeacherResult = [teacherList[i],[]]
      let tmp_ClassLoc = []
      let tmp_LessonName = []
      let tmp_ForWho = []
      let tmp_Url = []

      // 利用找到的index取得對應的表單內容
      for (let j = 0 ; j < indexesFind.length ; j ++){        
        // NOTE:若這裡出錯，要記得看一下屬性的名稱是否正確
          tmp_ClassLoc.push(data.classLoc[indexesFind[j]])
          tmp_LessonName.push(data.lessonName[indexesFind[j]])
          tmp_ForWho.push(data.forwho[indexesFind[j]])
          tmp_Url.push(data.url[indexesFind[j]])
      }

      // 找出教室位置unique值
      let classLocList = [... new Set(tmp_ClassLoc)]

      // 開始透過教室進行輸出
      for (let k = 0 ; k < classLocList.length ; k ++){
        // 暫存結果的預設位置，假設情況為有老師在MA218有兩堂課要上，預覽其中一筆的結果如下
        // [ ['MA218',[['授課班級','課程名稱','課程URL'],['授課班級','課程名稱','課程URL']]].....]

        let indexesFind = getIndexes(tmp_ClassLoc,classLocList[k])
        let tmp_result = [classLocList[k],[]]

        // 找課程資訊
        for (let kk = 0 ; kk < indexesFind.length ; kk++){
          let tmp_ClassOutput = []
          tmp_ClassOutput.push(tmp_ForWho[indexesFind[kk]])
          tmp_ClassOutput.push(tmp_LessonName[indexesFind[kk]])
          tmp_ClassOutput.push(tmp_Url[indexesFind[kk]])
          tmp_result[1].push(tmp_ClassOutput)
        }
        tmp_TeacherResult[1].push(tmp_result)
      }      
      result.push(tmp_TeacherResult)
      Logger.log(tmp_TeacherResult)
    } 
    return result; 
  }

  // -------------------------------------------
  
  // 儲存從表單讀取的所有資訊
  result_ob = check(all_formId)
  result_info = formDataToInfo(result_ob)
  return result_info;
}

/**
 * 先檢查是否存在給老師們表單連結的資料夾，反之建立。
 * 最後將其結果儲存為Google文件
 * 內容將以授課地點為段落將其上的課程名稱及表單連結放進去。
 * 
 * @param {Array} result_info - 所有老師處理後的結果
 * @param {String} formFolderId - 授課意見調查表的資料夾ID
 * 
 * return null  
 */
function output_AS_Word(result_info,formFolderId){
  /**
   * 在目標文件夾中建立文件
   * 
   * @param {Array} content - 該位老師的授課資訊
   * @param {String} targetFolderId - 目標資料夾的ID
   * 
   * return null;
   */
  function createDocument(content,targetFolderId) {
    let doc = DocumentApp.create(content[0] + '老師_授課意見調查表_表單連結');
    var body = doc.getBody();
    
    // 插入標題(大標題註記是誰)
    body.appendParagraph(content[0] +'老師_表單連結')
        .setHeading(DocumentApp.ParagraphHeading.HEADING1);

    for (let i = 0 ; i < content[1].length ; i++){

      // 插入標題(教室)
      body.appendParagraph(content[1][i][0])
        .setHeading(DocumentApp.ParagraphHeading.HEADING2)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .setUnderline(true);

      // 插入表格
      let table = body.appendTable();

      // 資料
      let data = content[1][i][1];

      // 表格標題
      let tableHeaderRow = table.appendTableRow();
      tableHeaderRow.appendTableCell("授課班級")
      tableHeaderRow.appendTableCell("課程名稱")
      tableHeaderRow.appendTableCell("表單連結")

      // 插入資料到表格
      for (let i = 0; i < data.length; i++) {
        let row = table.appendTableRow();
        for (let j = 0 ; j < data[i].length ; j++){
          row.appendTableCell(data[i][j]);
        }
      }
    }
    // 確保內容已經儲存過後進行位置轉移
    doc.saveAndClose()

    // 取得文件連結
    let url = doc.getUrl();
    let docId = doc.getId();

    // 將文件存在特定位置
    let targetFolder = DriveApp.getFolderById(targetFolderId);
    let newDocFile = DriveApp.getFileById(docId);

    targetFolder.addFile(newDocFile);
    DriveApp.getRootFolder().removeFile(newDocFile);

    Logger.log(content[0]+"老師")
    Logger.log('文件已建立：' + url);
  }

  // 先確保有沒有資料夾存放這些文件
  let parentFolder = DriveApp.getFolderById(formFolderId);
  if (!parentFolder.getFoldersByName("老師們的表單連結").hasNext()){
    parentFolder.createFolder("老師們的表單連結")
    Logger.log("偵測到無資料夾，已建立「老師們的表單連結」")
    }
  var targetFolderId = parentFolder.getFoldersByName("老師們的表單連結").next().getId()
  for (let i = 0 ; i < result_info.length ; i++){
    createDocument(result_info[i],targetFolderId)
  }
  
}

/**
 * 主執行
 * 請選擇學年以及學期
 */
function init(){

  // 填寫學年和學期
  let schoolyear = "112(2)"

  // 先取得該學年或學期的授課意見調查表的位置
  let targetFolderId = getFolderloc(
    "系上表單及文件",
    schoolyear
  );

  // 取得該資料夾內所有formId
  let all_formid = getAllForm(targetFolderId);

  // 開始整理資訊
  let result_info = arrange_all_data(all_formid);

  // 輸出為Word並放在授課意見調查表該資料夾下的一個資料夾
  output_AS_Word(result_info,targetFolderId);

  Logger.log("執行結束，可去資料夾看是否有所有老師的表單，審核後可寄出")
}
