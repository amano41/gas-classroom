const MY_EMAIL_ROW = 1;
const MATERIALS_FOLDER_ROW = 2;
const COURSES_LIST_ROW = 6;


function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu("授業管理");
  menu.addItem("授業回の追加", "createNewLesson");
  menu.addSeparator();
  menu.addItem("クラス一覧の更新", "listActiveCourses");
  menu.addSeparator();
  menu.addItem("未提出ファイルの掃除", "removeUnsubmittedFiles");
  menu.addItem("提出物ファイルのリネーム", "renameSubmittedFiles");
  menu.addToUi();
}


function debug_listCourses() {
  const courses = listCourses();
  for (const course of courses) {
    console.log("course: %s %s (%s)", course.name, course.section, course.id);
  }
}


function debug_listTopics() {
  const courseId = "xxxxxxxxxxxx";
  const topics = listTopics(courseId);
  for (const topic of topics) {
    console.log("topic: %s (%s)", topic.name, topic.topicId);
  }
}


function debug_listCourseWorks() {
  const courseId = "xxxxxxxxxxxx";
  const courseWorks = listCourseWorks(courseId);
  for (const courseWork of courseWorks) {
    console.log("coursework: %s (%s)", courseWork.title, courseWork.id);
  }
}


function debug_listSubmissions() {
  const courseId = "xxxxxxxxxxxx";
  const courseWorkId = "xxxxxxxxxxxx";
  const submissions = listSubmissions(courseId, courseWorkId);
  for (const submission of submissions) {
    if (submission.courseWorkType !== "ASSIGNMENT") {
      continue;
    }
    if (submission.state !== "TURNED_IN") {
      continue;
    }
    console.log("submission: %s [%s]", submission.id, submission.alternateLink);
  }
}


/**
 * クラスの一覧を取得
 */
function listCourses(courseStates = ["ACTIVE"], pageToken = null) {
  const optionalArgs = {
    courseStates: courseStates,
    pageToken: pageToken
  }
  const response = Classroom.Courses.list(optionalArgs);
  if (response.nextPageToken) {
    return response.courses.concat(listCourses(response.nextPageToken));
  }
  else {
    return response.courses;
  }
}


/**
 * トピックの一覧を取得
 */
function listTopics(courseId, pageToken = null) {
  const optionalArgs = {
    pageToken: pageToken
  }
  const response = Classroom.Courses.Topics.list(courseId, optionalArgs);
  if (response.nextPageToken) {
    return response.topic.concat(listTopics(courseId, response.nextPageToken));
  }
  else {
    return response.topic;
  }
}



/**
 * 課題の一覧を取得
 */
function listCourseWorks(courseId, pageToken = null) {
  const optionalArgs = {
    pageToken: pageToken
  }
  const response = Classroom.Courses.CourseWork.list(courseId, optionalArgs);
  if (response.nextPageToken) {
    return response.courseWork.concat(listCourseWorks(courseId, response.nextPageToken));
  }
  else {
    return response.courseWork;
  }
}


/**
 * 提出物の一覧を取得
 */
function listSubmissions(courseId, courseWorkId, pageToken = null) {
  const optionalArgs = {
    pageToken: pageToken
  }
  const response = Classroom.Courses.CourseWork.StudentSubmissions.list(courseId, courseWorkId, optionalArgs);
  if (response.nextPageToken) {
    return response.studentSubmissions.concat(listSubmissions(courseId, courseWorkId, response.nextPageToken));
  }
  else {
    return response.studentSubmissions;
  }
}


/**
 * 開講中のクラス一覧を取得
 */
function listActiveCourses() {

  const courses = listCourses();
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
  const lessonNumber = Browser.inputBox("授業回を 00 形式で入力してください。");
  if (lessonNumber === "cancel") {
    Browser.msgBox("実行を中止します。");
    return;
  }

  // タイトル
  const lessonTitle = Browser.inputBox("授業のタイトルを入力してください。");
  if (lessonTitle === "cancel") {
    Browser.msgBox("実行を中止します。");
    return;
  }

  // 実施日
  const lessonDate = Browser.inputBox("授業の実施日を YYYY/MM/DD 形式で入力してください。");
  if (lessonDate === "cancel") {
    Browser.msgBox("実行を中止します。");
    return
  }

  // Slido Event Code
  const eventCode = Browser.inputBox("Slido のイベントコードを入力してください。");
  if (eventCode === "cancel") {
    Browser.msgBox("実行を中止します。");
    return
  }

  // Slido Event URL
  const eventURL = Browser.inputBox("Slido のイベント URL を入力してください。");
  if (eventURL === "cancel") {
    Browser.msgBox("実行を中止します。");
    return
  }

  // 実行の確認
  const confirm = Browser.msgBox(
    "実行の確認",
    "以下の設定で授業を作成します。よろしいですか？\\n\\n" +
    "授業名：　" + lessonNumber + " - " + lessonTitle + "\\n" +
    "実施日：　" + lessonDate + "\\n\\n" +
    "作成数：　" + (lastRow - COURSES_LIST_ROW + 1) + " クラス",
    Browser.Buttons.YES_NO);
  if (confirm !== "yes") {
    Browser.msgBox("キャンセルしました。");
    return;
  }

  // 授業回の教材フォルダ
  const lessonFolder = getFolderByPath(lessonNumber, getMaterialsFolder());

  // 出席確認フォームのコピー
  copyAttendanceForm(lessonNumber);

  // クラス一覧を順番に処理
  for (let row = COURSES_LIST_ROW; row < lastRow + 1; row++) {

    const courseName = sheet.getRange(row, 1).getValue();
    const courseId = sheet.getRange(row, 3).getValue();
    const startTime = sheet.getRange(row, 2).getDisplayValue();

    // 進捗状況の報告
    Browser.msgBox("クラス「" + courseName + "」（" + courseId + "）に授業を作成します。");

    // 予約投稿時間は授業開始 10 分前
    const scheduledDate = new Date(lessonDate + " " + startTime);
    scheduledDate.setMinutes(scheduledDate.getMinutes() - 10);

    // 課題の提出期限は授業日の一週間後
    const assignmentDue = new Date(lessonDate + " " + startTime);
    assignmentDue.setDate(assignmentDue.getDate() + 7);

    // 出席確認フォームの回答期限は授業開始 30 分後
    const attendanceDue = new Date(lessonDate + " " + startTime);
    attendanceDue.setMinutes(attendanceDue.getMinutes() + 30);

    // トピックの作成
    const topicName = lessonNumber + " - " + lessonTitle
    topicId = createTopic(courseId, topicName);

    // 資料の作成
    createSlido(courseId, topicId, scheduledDate, lessonFolder, eventCode, eventURL);
    createMaterial(courseId, topicId, "参考資料", "", scheduledDate, lessonFolder, "参考資料");
    createMaterial(courseId, topicId, "授業資料", "", scheduledDate, lessonFolder, "授業資料");
    createMaterial(courseId, topicId, "課題" + lessonNumber + "：チェックデータ", "", scheduledDate, lessonFolder, "課題/チェックデータ");
    createMaterial(courseId, topicId, "課題" + lessonNumber + "：解答例", "", scheduledDate, lessonFolder, "課題/解答例");

    // 課題の作成
    createAssignment(courseId, topicId, assignmentDue, scheduledDate, lessonFolder);

    // 出席確認の作成
    createAttendance(courseId, topicId, attendanceDue, scheduledDate, lessonFolder);
  }

  Browser.msgBox("授業「" + lessonNumber + " - " + lessonTitle + "」の作成が完了しました！");
}


/**
 * 指定したフォルダの内容から資料を作成する
 */
function createMaterial(courseId, topicId, title, description, scheduledDate, lessonFolder, sourceFolderPath = "資料") {

  const sourceFolder = getFolderByPath(sourceFolderPath, lessonFolder);
  if (!sourceFolder) {
    Browser.msgBox("フォルダ '" + lessonFolder.getName() + "/" + sourceFolderPath + "' が見つかりません。");
    return;
  }

  let attachments = [];
  const files = sourceFolder.getFiles();
  while (files.hasNext()) {
    attachments.push(files.next());  // まずはファイル自体を格納
  }

  // 大文字・小文字の違いを無視してファイル名の順に並び替える
  attachments = attachments.sort((a, b) => {
    return (a.getName().toLowerCase() > b.getName().toLowerCase()) ? 1 : -1;
  });

  // ファイル ID の配列に変換する
  attachments = attachments.map(v => v.getId());

  createCourseWorkMaterial(courseId, topicId, title, description, scheduledDate, attachments);
}


/**
 * 指定したフォルダの内容から課題を作成する
 */
function createAssignment(courseId, topicId, dueDate, scheduledDate, lessonFolder, sourceFolderPath = "課題") {

  const lessonNumber = lessonFolder.getName();

  // 課題フォルダ
  const sourceFolder = getFolderByPath(sourceFolderPath, lessonFolder);
  if (!sourceFolder) {
    Browser.msgBox("フォルダ '" + lessonNumber + "/" + sourceFolderPath + "' が見つかりません。");
    return;
  }

  // 指示用ファイル
  const fileName = "課題" + lessonNumber + ".pdf";
  const fileId = getFileIdByPath(fileName, sourceFolder);
  if (!fileId) {
    Browser.msgBox("ファイル '" + fileName + "' が見つかりません。");
    return;
  }

  // 解答用ファイルがなければ指示用ファイルだけを添付して終了
  const supplementFolder = getFolderByPath("解答用ファイル", sourceFolder);
  if (!supplementFolder) {
    console.log("No supplement files.");
    createCourseWork(courseId, topicId, "課題" + lessonNumber, "", dueDate, scheduledDate, [fileId]);
    return;
  }

  // 解答用ファイルを添付
  let attachments = [];
  const files = supplementFolder.getFiles();
  while (files.hasNext()) {
    attachments.push(files.next());  // まずはファイル自体を格納
  }

  // 大文字・小文字の違いを無視してファイル名の順に並び替える
  attachments = attachments.sort((a, b) => {
    return (a.getName().toLowerCase() > b.getName().toLowerCase()) ? 1 : -1;
  });

  // ファイル ID の配列に変換する
  attachments = attachments.map(v => v.getId());

  // 指示用ファイルは常に先頭
  attachments.unshift(fileId);

  createCourseWork(courseId, topicId, "課題" + lessonNumber, "", dueDate, scheduledDate, attachments);
}


/**
 * 出席確認フォームの課題を作成する
 */
function createAttendance(courseId, topicId, dueDate, scheduledDate, lessonFolder) {

  const lessonNumber = lessonFolder.getName();

  const attendanceFile = getFileByPath("出席確認" + lessonNumber, lessonFolder);
  if (!attendanceFile) {
    Browser.msgBox("フォーム '出席確認" + lessonNumber + "' が見つかりません。");
    return;
  }

  createCourseWork(courseId, topicId, "出席確認" + lessonNumber, "", dueDate, scheduledDate, [attendanceFile.getUrl()]);
}


/**
 * Slido のリンク資料を作成する
 */
function createSlido(courseId, topicId, scheduledDate, lessonFolder, eventCode, eventURL) {

  const description = "#" + eventCode;

  const attachments = [eventURL];
  const lessonNumber = lessonFolder.getName();
  const qrcode = getFileIdByPath("slido" + lessonNumber + ".png", lessonFolder);
  if (qrcode === null) {
    Browser.msgBox("Slido の QR コードが見つかりません。");
  }
  else {
    attachments.push(qrcode);
  }

  createCourseWorkMaterial(courseId, topicId, "Slido", description, scheduledDate, attachments);
}


/**
 * パスからフォルダを取得
 */
function getFolderByPath(folderPath, baseFolder = null) {

  const parts = folderPath.split("/");
  while (parts[0] === "" || parts[0] === "マイドライブ") {
    parts.shift();
  }

  let folder = baseFolder || DriveApp.getRootFolder();

  while (parts.length > 0) {

    const target = parts.shift();
    if (target === "") {
      continue
    }

    const folders = folder.getFoldersByName(target);
    if (folders.hasNext()) {
      folder = folders.next();
    }
    else {
      console.log("Folder '%s' not found in '%s'", target, folder);
      return null;
    }
  }

  return folder;
}


/**
 * パスからフォルダの ID を取得
 */
function getFolderIdByPath(folderPath, baseFolder = null) {
  const folder = getFolderByPath(folderPath, baseFolder);
  if (folder) {
    return folder.getId();
  }
  else {
    return null;
  }
}


/**
 * パスからファイルを取得
 */
function getFileByPath(filePath, baseFolder = null) {

  const parts = filePath.split("/");
  while (parts[0] === "" || parts[0] === "マイドライブ") {
    parts.shift();
  }

  let folder = baseFolder || DriveApp.getRootFolder();

  while (parts.length > 0) {

    const target = parts.shift();
    if (target === "") {
      continue
    }

    // パス要素が残っていれば途中のフォルダ
    if (parts.length > 0) {
      const folders = folder.getFoldersByName(target);
      if (folders.hasNext()) {
        folder = folders.next();
      }
      else {
        console.log("Folder '%s' not found in '%s'", target, folder);
        return null;
      }
    }
    // 最後のパス要素であれば目的のファイル
    else {
      const files = folder.getFilesByName(target);
      if (files.hasNext()) {
        return files.next();
      }
      else {
        console.log("File '%s' not found in '%s'", target, folder);
        return null;
      }
    }
  }

  return null;
}


/**
 * パスからファイルの ID を取得
 */
function getFileIdByPath(filePath, baseFolder = null) {
  const file = getFileByPath(filePath, baseFolder);
  if (file) {
    return file.getId();
  }
  else {
    return null;
  }
}


/**
 * 教材フォルダを取得
 */
function getMaterialsFolder() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const materialsFolderId = sheet.getRange(MATERIALS_FOLDER_ROW, 2).getValue();
  return DriveApp.getFolderById(materialsFolderId);
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
function createCourseWorkMaterial(courseId, topicId, title, description, scheduledDate, attachments = []) {

  const scheduledTime = Utilities.formatDate(scheduledDate, "UTC", "yyyy-MM-dd'T'HH:mm:ss'Z'")

  const resource = {
    "title": title,
    "description": description,
    "state": "DRAFT",
    "scheduledTime": scheduledTime,
    "materials": [],
    "topicId": topicId
  }

  for (const attachment of attachments) {
    if (attachment.match(/^https?:/)) {
      resource.materials.push({
        "link": {
          "url": attachment
        }
      });
    }
    else {
      resource.materials.push({
        "driveFile": {
          "driveFile": { "id": attachment }
        }
      });
    }
  }

  const response = Classroom.Courses.CourseWorkMaterials.create(resource, courseId);
  console.log("Material created: %s", title);
  return response;
}


/**
 * 課題の作成
 */
function createCourseWork(courseId, topicId, title, description, dueDate, scheduledDate, attachments = []) {

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
    "materials": [],
    "topicId": topicId
  }

  for (const attachment of attachments) {
    if (attachment.match(/^https?:/)) {
      resource.materials.push({
        "link": {
          "url": attachment
        }
      });
    }
    else {
      resource.materials.push({
        "driveFile": {
          "driveFile": { "id": attachment }
        }
      });
    }
  }

  const response = Classroom.Courses.CourseWork.create(resource, courseId);
  console.log("Assignment created: %s", title);
  return response;
}


/**
 * 出席確認フォームのコピー
 */
function copyAttendanceForm(lessonNumber) {

  const materialsFolder = getMaterialsFolder();

  // 出席確認フォームのテンプレート
  const templateFile = getFileByPath("Template/出席確認", materialsFolder);
  if (!templateFile) {
    console.log("Template file not found: 出席確認");
    return;
  }

  // コピー先となる授業回フォルダ
  let lessonFolder = getFolderByPath(lessonNumber, materialsFolder);
  if (!lessonFolder) {
    lessonFolder = materialsFolder.createFolder(lessonNumber);
    console.log("Lesson folder created: %s", lessonNumber);
  }

  // コピー
  const destName = "出席確認" + lessonNumber
  const destFile = templateFile.makeCopy(destName, lessonFolder);
  const form = FormApp.openById(destFile.getId());
  form.setTitle(destName);

  // 回答先をスプレッドシートにする
  const sheet = SpreadsheetApp.create(destName + "（回答）");
  SpreadsheetApp.flush();
  const sheetFile = DriveApp.getFileById(sheet.getId());
  sheetFile.moveTo(lessonFolder);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, sheet.getId());
  sheet.deleteSheet(sheet.getSheetByName("シート1"));

  console.log("Attendance forms created: 出席確認%s", lessonNumber);
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
      Browser.msgBox("Invalid folder: " + folderPath);
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


/**
 * 提出物ファイルをリネームする
 */
function renameSubmittedFiles() {

  const confirm = Browser.msgBox(
    "提出物ファイルのリネーム", "実行してもよろしいですか？", Browser.Buttons.YES_NO
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

    const courseId = sheet.getRange(row, 3).getValue();
    const courseWorks = listCourseWorks(courseId);

    for (courseWork of courseWorks) {

      // 種類が「課題」でなければ飛ばす
      if (courseWork.workType !== "ASSIGNMENT") {
        continue;
      }

      // 状態が「公開」でなければ飛ばす
      if (courseWork.state !== "PUBLISHED") {
        continue;
      }

      // 提出期限前であれば飛ばす
      today = new Date();
      dueDate = new Date(courseWork.dueDate.year, courseWork.dueDate.month - 1, courseWork.dueDate.day);
      if (today.getTime() <= dueDate.getTime()) {
        continue;
      }

      // 提出物ファイルの先頭に学籍番号をつける
      ensureStudentIdPrefix(courseId, courseWork.id);
    }
  }
}


/**
 * 提出物のファイル名の先頭に学籍番号がつくようにリネームする
 */
function ensureStudentIdPrefix(courseId, courseWorkId) {

  const submissions = listSubmissions(courseId, courseWorkId);

  for (const submission of submissions) {

    // 状態が「提出済み」でなければ飛ばす
    if (submission.state !== "TURNED_IN") {
      continue;
    }

    // 種類が「課題」でなければ飛ばす
    if (submission.courseWorkType !== "ASSIGNMENT") {
      continue;
    }

    // assignmentSubmission は courseWorkType === "ASSIGNMENT" の場合のみ有効
    const attachments = submission.assignmentSubmission.attachments;

    // ファイルが添付されていなければ飛ばす
    if (!attachments || attachments.length === 0) {
      continue;
    }

    // ファイル名をひとつずつ確認
    for (const attachment of attachments) {
      if ("driveFile" in attachment) {

        const fileName = attachment.driveFile.title;
        let newFileName = null;

        // 命名規則に従っていれば飛ばす
        if (fileName.match(/^ne\d{6}_/)) {
          continue;
        }

        const userId = submission.userId;
        const userName = Classroom.Courses.Students.get(courseId, userId).profile.name.fullName;

        const match = fileName.match(/^(ne|NE)(\d{2})-?(\d{4})[A-Za-z]?(_| )(.+)$/);
        if (match) {
          newFileName = "ne" + match[2] + match[3] + "_" + match[5];
        }
        else {
          const studentId = userName.replace(/^(ne|NE)(\d{2})-?(\d{4}).+$/, "ne$2$3");
          newFileName = studentId + "_" + fileName;
        }

        const fileId = attachment.driveFile.id;
        const file = DriveApp.getFileById(fileId);

        // 念のためファイル名が一致するか確認してからリネーム
        if (file.getName() === fileName) {
          //file.setName(newFileName);
          console.log(userName + " : " + fileName + " ==> " + newFileName);
        }
        else {
          console.error("File name does not match.");
          console.error(userName + " : " + fileName + " vs. " + file.getName());
        }
      }
    }
  }
}
