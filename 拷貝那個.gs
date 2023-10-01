/**
 * Project: 拷貝那個！ CopyThat!
 * Author: Chang, Chia-Cheng; 張家誠
 * Date: 2023-8-26 v1.0
 * Copyright (c) 2023 Chang, Chia-Cheng 張家誠
 * Fix Bug:
 ** 2023-10-2 修正Apps Script類型檔案複製至根目錄的問題
 */

/**
 * A built-in function on GAS.
 */
function onOpen() {
  SpreadsheetApp.getActive().addMenu('進階功能', [{name: '拷貝', functionName: 'makeCopy'}]); //add menu on sheet
}

sumFolders = 0; //variable for count folders copied
sumFiles = 0; //variable for count files copied

/**
 * Click the menu to execute this function. It will check whether the folder to be copied and the folder to be stored are correct and share permissions. If it is correct, check the files and folders from the first layer and copy those files and folders. If there is a folder at the next level, call the goNext function.
 * 
 * Regardless of whether it is correct or incorrect, the program will pop up a message window to display the execution status.
 * 
 * Finally, it will count how many files and folders have been copied in total, and the number of seconds used for this copying process.
 */
function makeCopy() {
  var startTime = parseInt( new Date().getTime() / 1000, 0); //get start time of execution
  var ss = SpreadsheetApp.getActive();
  st = ss.getSheetByName('main'); //get sheet
  st.getRange(3,2).setValue(''); //clear message last execute
  try{
    var copyFolderId = st.getRange(1,2).getValue().split("/folders/")[1].split("?")[0]; //get copied folder's id
  }catch{
    st.getRange(3,2).setValue('程式中斷：複製資料夾的連結有錯誤，未提供複製資料夾的連結。'); //record message for wrong folder's id to copy
    var errMsg = HtmlService.createHtmlOutput('程式偵測錯誤：<br />複製資料夾的連結有錯誤，未提供複製資料夾的連結。')
    .setTitle('程式中斷')
    .setWidth(300)
    .setHeight(200);
    ss.show(errMsg); //show message for wrong folder's id to copy
    return;
  }
  try{
    var copyFolder = DriveApp.getFolderById(copyFolderId); //get copied folder
  }catch{
    st.getRange(3,2).setValue('程式中斷：複製資料夾的連結有錯誤；可能的原因是「複製資料夾連結格式錯誤」或者「複製資料夾未提供『檢視者』權限」'); //record message for wrong folder's id to copy
    var errMsg = HtmlService.createHtmlOutput('程式偵測錯誤：<br />複製資料夾的連結有錯誤；可能的原因是：<br /><br />「複製資料夾連結格式錯誤」<br />「複製資料夾未提供『檢視者』權限」')
    .setTitle('程式中斷')
    .setWidth(300)
    .setHeight(200);
    ss.show(errMsg); //show message for wrong folder's id to copy
    return;
  }
  try{
    var saveFolderId = st.getRange(2,2).getValue().split("/folders/")[1].split("?")[0]; //get saved folder's id
  }catch{
    st.getRange(3,2).setValue('程式中斷：存放資料夾的連結有錯誤，未提供存放資料夾的連結。'); //record message for wrong folder's id to save
    var errMsg = HtmlService.createHtmlOutput('程式偵測錯誤：<br />存放資料夾的連結有錯誤，未提供存放資料夾的連結。')
    .setTitle('程式中斷')
    .setWidth(300)
    .setHeight(200);
    ss.show(errMsg); //show message for wrong folder's id to save
    return;
  }
  try{
    var saveFolder = DriveApp.getFolderById(saveFolderId); //get saved folder
  }catch{
    st.getRange(3,2).setValue('程式中斷：存放資料夾的連結有錯誤，存放資料夾連結格式錯誤。'); //record message for wrong folder's id to save
    var errMsg = HtmlService.createHtmlOutput('程式偵測錯誤：<br />存放資料夾的連結有錯誤，存放資料夾連結格式錯誤。')
    .setTitle('程式中斷:')
    .setWidth(300)
    .setHeight(200);
    ss.show(errMsg); //show message for wrong folder's id to save; If you do not wish to display an advertisement page, please delete this line or change it to your own page. However, I would very much like you to keep this message. Thank you very much.
    return;
  }
  ss.show(copyright); //show message for program is running
  var files = copyFolder.getFiles(); //get and copy files in first level
  while(files.hasNext()){
    var file = files.next();
    var newFile = file.makeCopy(file.getName(), saveFolder);
    newFile.moveTo(saveFolder);
    sumFiles++;
  }
  //go to next level folder to find files and folders
  var folders = copyFolder.getFolders();
  while(folders.hasNext()){
    var folder = folders.next();
    goNext(folder, saveFolder); //call function to find and copy
  }
  var endTime = parseInt( new Date().getTime() / 1000, 0); //get end time of execution
  var useTime = endTime - startTime; //get seconds of program executes
  st.getRange(3,2).setValue(`完成拷貝！總共複製 ${sumFolders} 個資料夾, ${sumFiles} 個檔案，程式耗時 ${useTime} 秒。`); //record message for result of program execute
  ss.show(end); //show message for result of program execute; If you do not wish to display an advertisement page, please delete this line or change it to your own page. However, I would very much like you to keep this message. Thank you very much.
}

/**
 * This function can be used to check the files and folders of the next layer of folders. When encountering a lower layer of folders, it will call and execute itself again until each layer of folders has been scanned.
 */
function goNext(subFolder, saveFolder){
  var saveSubFolder = saveFolder.createFolder(subFolder); //create sub folder
  sumFolders++;
  var files = subFolder.getFiles();
  //copied all files in sub folder
  while(files.hasNext()){
    var file = files.next();
    var newFile = file.makeCopy(file.getName(), saveSubFolder);
    newFile.moveTo(saveSubFolder);
    sumFiles++
  }
  //if find sub folder, and call function goNext again
  var folders = subFolder.getFolders();
  while(folders.hasNext()){
    var folder = folders.next();
    goNext(folder, saveSubFolder);
  }
}

/**
 * Show my copyright on pop-up message window. If you do not wish to display an advertisement page, please delete this part or change it to your own page. However, I would very much like you to keep this message. Thank you very much.
 */
copyright = HtmlService.createHtmlOutput('<iframe scrolling="no" style="position:fixed; top:0; left:0; bottom:0; right:0; width:100%; height:100%; border:none; margin:0; padding:0; overflow:hidden; z-index:999999;" src="https://php-pie.net/GAS/CopyThat-zh-tw.php"></iframe>')
.setTitle('程式正在執行，敬請稍待')
.setHeight(600);

/**
 * Show result on pop-up message window. If you do not wish to display an advertisement page, please delete this part or change it to your own page. However, I would very much like you to keep this message. Thank you very much.
 */
end = HtmlService.createHtmlOutput('<iframe scrolling="no" style="position:fixed; top:0; left:0; bottom:0; right:0; width:100%; height:100%; border:none; margin:0; padding:0; overflow:hidden; z-index:999999;" src="https://php-pie.net/GAS/CopyThat-done-zh-tw.php"></iframe>')
.setTitle('拷貝那個！')
.setHeight(600);

