/**
 * Copyright (c) 彬彬 2024
 * All rights reserved.
 * 
 * Created by: 彬彬
 * Contact: C110156220@nkust.edu.tw / robin92062574@gmail.com
 */


// 在使用此程式碼之前請記得要先執行表單製作且確認init裡面的schoolyear是否是你想要輸出的學年度
// 若出現程式碼問題，請聯絡彬彬或其他維護人員。

/**
 * 鎖定資料夾位置
 * 
 * @param {String} parentFolderName 父資料夾(指雲端的第一層)
 * @param {String} schoolyear 學年和學期(e.g. 113(1) )
 * 
 * 依照以上的範例，通常會鎖定113(1)授課意見調查表
 * return folderId
 */
function getFolderloc(parentFolderName="系上表單及文件",schoolyear){
  let parentFolder = DriveApp.getFoldersByName(parentFolderName);
  if (!parentFolder.hasNext()){
    throw Error("找不到系上表單及文件，有改路徑需聯絡相關維護人員，")
  }
  else{
    let newFolder = parentFolder.next()
    newFolder = newFolder.getFilesByName("授課意見調查表").next()
  
    if(!newFolder.getFilesByName(schoolyear + "授課意見調查表").hasNext()){
        throw Error("找不到授課意見調查表，路徑有改要連絡相關維護人員喔")
      }
    else{
      newFolder = newFolder.getFilesByName(schoolyear + "授課意見調查表").next()
      return newFolder.getId()
    }
  }
}


/**
 * 透過資料夾ID取得該資料的所有檔案
 * 透過迴圈得取表單ID
 * 
 * @params {String} FolderId 資料夾ID
 * 
 * return ['表單ID','表單ID','表單ID','表單ID'.....] 
 */
function getAllForm(FolderId) {
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();

    // 儲存表單ID
    let result = []
  
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
 * @param {String} id Google表單ID
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
   *  @param {String} Content 授課意見調查表_表單簡介
   * 
   *  return {Array} 長這樣['授課班級','授課老師','授課教室','課程名稱]
   *
   */
  function getClassInfo(Content){

     // 使用正規表達式來擷取授課班級和授課老師
    
    var regexClassLoc = /授課教室：(.+?)\n/;
    var regexClass = /授課班級：(.+?)\n/;
    var regexTeacher = /授課老師：(.+?)\n/;
    var regexLessonName = /課程名稱：(.+?)\n/;

    var classMatch = formDescription.match(regexClass);
    var teacherMatch = formDescription.match(regexTeacher);
    var classLocMatch = formDescription.match(regexClassLoc);
    var lessonNameMatch = formDescription.match(regexLessonName);

    var courseInfo = {
      class: classMatch ? classMatch[1] : "未知",
      teacher: teacherMatch ? teacherMatch[1] : "未知",
      classLoc: classLocMatch ? classLocMatch[1] : "未知",
      lessonName: lessonNameMatch ? lessonNameMatch[1] : "未知"
    };

    return [

      courseInfo.class,
      courseInfo.teacher,
      courseInfo.classLoc,
      courseInfo.lessonName

    ];
  }

  // -------------------------------------------
  /**
   * 測試:輸出結果
   * @param {String} formTitle 表單標頭
   * @param {String} formDescription 表單簡介
   * @param {String} formUrl 表單連結
   * @param {String} formDesData 表單簡介拿取的資料 ref -> getClassInfo
   */
  function output_logTest(formTitle,formDescription,formUrl,formDesData){
    Logger.log('表單標題: ' + formTitle);
    Logger.log('表單簡介: ' + formDescription);
    Logger.log('courseInfo' + formDesData) ;
    Logger.log('表單連結: ' + formUrl);
  }

  // -------------------------------------------

  // run 
  var form = FormApp.openById(id); 
  // 拿取表單資訊
  var formTitle = form.getTitle();
  var formDescription = form.getDescription();
  var formUrl = form.getPublishedUrl();

  let content = getClassInfo(formDescription);

  output_logTest(formTitle,formDescription,formUrl,content)
  
  // 將原結果再添加表單連結進去
  content.push(formUrl)

  return content
}

/**
 * 處理取得每個表單的實際資訊
 * @param {Array} all_formId 該資料夾所有的表單ID
 */
function arrange_all_data(all_formId){

  /**
   * 得取的資料進行整理
   * @param {Array} all_formId 該資料夾所有的表單ID(Array)
   * return 結果
   */
  function check(all_formId){
    // 結果
    let result = {

    // 授課地點
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
      result.teacher.push(tmp_res[1])
      result.lessonName.push(tmp_res[3])
      result.forwho.push(tmp_res[2])

    }
    return result;
    
  }
  // 儲存老師、課程名稱、開課班級、表單連結
  result_ob = check(all_formId)

  // 整理後老師的名子
  teacher = []


}

/**
 * 將其結果儲存為Google文件
 * @param {object} 所有老師授課的資訊
 */
function output_AS_Word(){
  
}

/**
 * 主執行
 * 請選擇學年以及學期
 */
function init(){

  // 填寫學年和學期
  let schoolyear = 112(2)

  // 先取得該學年或學期的授課意見調查表的位置
  let targetFolderId = getFolderloc(schoolyear);

  // 取得該資料夾內所有formId
  let all_formid = getAllForm(targetFolderId);

  // 開始整理資訊
  arrange_all_data(all_formid)
  
}
