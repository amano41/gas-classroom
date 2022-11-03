const MY_EMAIL_ROW = 1;
const MATERIALS_FOLDER_ROW = 2;
const COURSES_LIST_ROW = 6;


function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu("授業管理");
  menu.addItem("授業回の追加", "createNewLesson");
  menu.addItem("クラス一覧の更新", "listAllCourses");
  menu.addSeparator();
  menu.addItem("未提出ファイルの掃除", "removeUnsubmittedFiles");
  menu.addToUi();
}


/**
 * クラス一覧を取得
 */
function listAllCourses() {

  const optionalArgs = {
    courseStates: ["ACTIVE"]
  }

  const response = Classroom.Courses.list(optionalArgs);
  const courses = response.courses;
  if (!courses || courses.length === 0) {
    Browser.msgBox("クラスがありません。");
    return;
  }

  const sheet = SpreadsheetApp.getActiveSheet();
  const myEmail = sheet.getRange(MY_EMAIL_ROW, 2).getValue();
  if (myEmail === "") {
    Browser.msgBox("Email が入力されていません。");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow >= COURSES_LIST_ROW) {
    const lastCol = sheet.getLastColumn();
    const range = sheet.getRange(COURSES_LIST_ROW, 1, lastRow - COURSES_LIST_ROW + 1, lastCol);
    range.clearContent();
  }

  let row = COURSES_LIST_ROW;
  for (const course of courses) {
    const owner = Classroom.UserProfiles.get(course.ownerId);
    if (owner.emailAddress === myEmail) {
      console.log(row, course.name, course.section, course.id, course.teacherFolder.id, course.teacherGroupEmail);
      sheet.getRange(row, 1).setValue(course.name + " (" + course.section + ") ");
      sheet.getRange(row, 3).setValue(course.id);
      sheet.getRange(row, 4).setValue(course.teacherFolder.id);
      sheet.getRange(row, 5).setValue(course.teacherGroupEmail);
      row = row + 1;
    }
  }
}


/**
 * 新しい授業回を作成
 */
function createNewLesson() {

  const sheet = SpreadsheetApp.getActiveSheet();

  // クラス一覧が取得されているか確認
  const lastRow = sheet.getLastRow();
  if (lastRow < COURSES_LIST_ROW) {
    Browser.msgBox("クラス情報がありません。\\nメニューから［クラス一覧の取得］を実行してください。");
    return;
  }

  // 授業回数
  const lessonNumber = Browser.inputBox("授業の回数を 00 形式で入力してください。");
  if (lessonNumber === "cancel") {
    console.log("Canceled.");
    return;
  }

  // タイトル
  const lessonTitle = Browser.inputBox("授業のタイトルを入力してください。");
  if (lessonTitle === "cancel") {
    console.log("Canceled.");
    return;
  }

  // 実施日
  const lessonDate = Browser.inputBox("授業の実施日を YYYY/MM/DD 形式で入力してください。");
  if (lessonDate === "cancel") {
    console.log("Canceled.");
    return
  }

  // 出席確認フォームのコピー
  copyAttendanceForm(lessonNumber);

  // クラス一覧を順番に処理
  for (let row = COURSES_LIST_ROW; row < lastRow + 1; row++) {

    const courseName = sheet.getRange(row, 1).getValue();
    console.log("Course: %s", courseName);

    const courseId = sheet.getRange(row, 3).getValue();
    const startTime = sheet.getRange(row, 2).getDisplayValue();

    // 予約投稿時間は授業開始 10 分前
    const scheduledDate = new Date(lessonDate + " " + startTime);
    scheduledDate.setMinutes(scheduledDate.getMinutes() - 10);

    // 課題の提出期限は授業日の一週間後
    const assignmentDue = new Date(lessonDate + " " + startTime);
    assignmentDue.setDate(assignmentDue.getDate() + 7);

    // 出席確認フォームの回答期限は授業開始 30 分後
    const attendanceDue = new Date(lessonDate + " " + startTime);
    attendanceDue.setMinutes(attendanceDue.getMinutes() + 30);

    // トピック作成
    const topicName = lessonNumber + " - " + lessonTitle
    topicId = createTopic(courseId, topicName);

    // 資料の作成
    createCourseWorkMaterial(courseId, topicId, "Slido", "", scheduledDate);
    createCourseWorkMaterial(courseId, topicId, "参考資料", "", scheduledDate);
    createCourseWorkMaterial(courseId, topicId, "授業資料", "", scheduledDate);
    createCourseWorkMaterial(courseId, topicId, "課題" + lessonNumber + "：チェックデータ", "", scheduledDate);
    createCourseWorkMaterial(courseId, topicId, "課題" + lessonNumber + "：解答例", "", scheduledDate);

    // 課題の作成
    createCourseWork(courseId, topicId, "課題" + lessonNumber, "", assignmentDue, scheduledDate);

    // 出席確認の作成
    createCourseWork(courseId, topicId, "出席確認" + lessonNumber, "", attendanceDue, scheduledDate);
  }

}


/**
 * 出席確認フォームのコピー
 */
function copyAttendanceForm(lessonNumber) {

  const sheet = SpreadsheetApp.getActiveSheet();
  const materialsFolderId = sheet.getRange(MATERIALS_FOLDER_ROW, 2).getValue();

  // Template フォルダ
  const materialsFolder = DriveApp.getFolderById(materialsFolderId);
  let folders = materialsFolder.getFoldersByName("Template")
  if (!folders.hasNext()) {
    console.log("Not Found: Template フォルダ")
    return;
  }

  // 出席確認フォームのテンプレート
  const templateFolder = folders.next();
  const files = templateFolder.getFilesByName("出席確認");
  if (!files.hasNext()) {
    console.log("Not Found: 出席確認フォーム")
    return;
  }
  const formTemplate = files.next();

  // コピー先となる授業回フォルダ
  folders = materialsFolder.getFoldersByName(lessonNumber);
  if (folders.hasNext()) {
    lessonFolder = folders.next();
  }
  else {
    lessonFolder = materialsFolder.createFolder(lessonNumber);
    console.log("Folder Created: %s", lessonNumber);
  }

  // コピー
  const fileName = "出席確認" + lessonNumber
  const copiedFile = formTemplate.makeCopy(fileName, lessonFolder);
  const form = FormApp.openById(copiedFile.getId());
  form.setTitle(fileName);

  // 回答先をスプレッドシートにする
  const destSheet = SpreadsheetApp.create(fileName + "（回答）");
  SpreadsheetApp.flush();
  const sheetFile = DriveApp.getFileById(destSheet.getId());
  sheetFile.moveTo(lessonFolder);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, destSheet.getId());
  destSheet.deleteSheet(destSheet.getSheetByName("シート1"));

  console.log("Attendance forms created: 出席確認%s", lessonNumber);
}


/**
 * トピックの作成
 */
function createTopic(courseId, topicName) {
  const resource = {
    name: topicName
  }
  const response = Classroom.Courses.Topics.create(resource, courseId);
  console.log("Topic created: %s", topicName);
  return response.topicId;
}


/**
 * 資料の作成
 */
function createCourseWorkMaterial(courseId, topicId, title, description, scheduledDate, fileId = null) {

  const scheduledTime = Utilities.formatDate(scheduledDate, "UTC", "yyyy-MM-dd'T'HH:mm:ss'Z'")

  const resource = {
    "title": title,
    "description": description,
    "state": "DRAFT",
    "scheduledTime": scheduledTime,
    "topicId": topicId
  }

  if (fileId) {
    resource["materials"] = [{
      "driveFile": {
        "driveFile": { "id": fileId }
      }
    }]
  }

  const response = Classroom.Courses.CourseWorkMaterials.create(resource, courseId);
  console.log("Material created: %s", title);
}


/**
 * 課題の作成
 */
function createCourseWork(courseId, topicId, title, description, dueDate, scheduledDate, fileId = null) {

  const year = dueDate.getUTCFullYear();
  const month = dueDate.getUTCMonth() + 1; // Month は 0-11
  const day = dueDate.getUTCDate();
  const hours = dueDate.getUTCHours();
  const minutes = dueDate.getUTCMinutes();
  const scheduledTime = Utilities.formatDate(scheduledDate, "UTC", "yyyy-MM-dd'T'HH:mm:ss'Z'")

  const resource = {
    "title": title,
    "description": description,
    "workType": "ASSIGNMENT",
    "state": "DRAFT",
    "dueDate": {
      "year": year,
      "month": month,
      "day": day,
    },
    "dueTime": {
      "hours": hours,
      "minutes": minutes
    },
    "scheduledTime": scheduledTime,
    "topicId": topicId
  }

  if (fileId) {
    resource["materials"] = [{
      "driveFile": {
        "driveFile": { "id": fileId }
      }
    }]
  }

  const response = Classroom.Courses.CourseWork.create(resource, courseId);
  console.log("Assignment created: %s", title);
}


/**
 * 学生が提出を取り消したファイルを一括削除する
 */
function removeUnsubmittedFiles() {

  const confirm = Browser.msgBox(
    "未提出ファイルの掃除", "実行してもよろしいですか？", Browser.Buttons.YES_NO
  );

  if (confirm !== "yes") {
    console.log("Canceled.");
    return;
  }

  const sheet = SpreadsheetApp.getActiveSheet();

  // 教師アカウントで実行しているか確認
  const userEmail = Session.getActiveUser().getEmail();
  const myEmail = sheet.getRange(MY_EMAIL_ROW, 2).getValue();
  if (userEmail !== myEmail) {
    Browser.msgBox("教師アカウントでログインしていません。\\n\\n" + userEmail);
    return;
  }

  // クラス一覧が取得されているか確認
  const lastRow = sheet.getLastRow();
  if (lastRow < COURSES_LIST_ROW) {
    Browser.msgBox("クラス情報がありません。\\nメニューから［クラス一覧の取得］を実行してください。");
    return;
  }

  // クラス一覧を順番に処理
  for (let row = COURSES_LIST_ROW; row < lastRow + 1; row++) {

    const folderId = sheet.getRange(row, 4).getValue();
    const folder = DriveApp.getFolderById(folderId);

    // 対象フォルダのフルパスを作成
    let folderPath = folder;
    let parents = folder.getParents();
    while (parents.hasNext()) {
      const parent = parents.next();
      folderPath = parent.getName() + "/" + folderPath;
      parents = parent.getParents();
    }
    folderPath = "/" + folderPath;

    // Classroom のフォルダが取得できているか確認
    if (!folderPath.match("/Classroom/")) {
      Browser.msgBox("Invalid Folder: " + folderPath);
      return;
    }

    // 編集権限を解除する教師アカウントの Email アドレス
    let teacherEmails = [];
    teacherEmails.push(myEmail);
    teacherEmails.push(sheet.getRange(row, 5).getValue());

    removeEditors(folder, folderPath, teacherEmails);
  }

}


/**
 * 指定フォルダ内の自分以外がオーナーのファイルから編集権限を削除する
 */
function removeEditors(targetFolder, targetFolderPath, targetEditors) {

  console.info(targetFolderPath);

  // 自分以外がオーナーで自分に編集権限があるファイルを処理
  let files = targetFolder.searchFiles('not ("me" in owners)');
  while (files.hasNext()) {

    let file = files.next();
    console.info("\t" + file.getName());

    let editors = file.getEditors();
    for (let i in editors) {
      if (targetEditors.includes(editors[i].getEmail())) {
        file.removeEditor(editors[i]);
      }
    }
  }

  // 子フォルダを再帰的に処理
  let folders = targetFolder.getFolders();
  while (folders.hasNext()) {
    let folder = folders.next();
    let folderPath = targetFolderPath + "/" + folder.getName();
    removeEditors(folder, folderPath, targetEditors);
  }
}
