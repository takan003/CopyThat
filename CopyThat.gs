/**
 * Project: CopyThat!
 * Author: Chang, Chia-Cheng; 張家誠
 * Date: 2023-8-26 v1.0
 * Copyright (c) 2023 Chang, Chia-Cheng 張家誠
 * Fix Bugs:
 ** 2023-10-2 Fixed the problem of copying Apps Script type files to the root directory. 
 */

/**
 * A built-in functions on GAS.
 */
function onOpen() {
  SpreadsheetApp.getActive().addMenu('Advanced', [{name: 'Make a copy', functionName: 'makeCopy'}]); //add menu on sheet
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
    st.getRange(3,2).setValue('Interrupt: Url of folder to copy is worong. Url of folder to copy not provided.'); //record message for wrong folder's id to copy
    var errMsg = HtmlService.createHtmlOutput('Program found error:<br />Url of folder to copy is worong. Url of folder to copy not provided.')
    .setTitle('Program Interrupt:')
    .setWidth(300)
    .setHeight(200);
    ss.show(errMsg); //show message for wrong folder's id to copy
    return;
  }
  try{
    var copyFolder = DriveApp.getFolderById(copyFolderId); //get copied folder
  }catch{
    st.getRange(3,2).setValue('Interrupt: Url of folder to copy is worong. The possible reasons are as follows: Wrong ID, provider turn off \'Viewer\' permission.'); //record message for wrong folder's id to copy
    var errMsg = HtmlService.createHtmlOutput('Program found error:<br />Url of folder to copy is worong. The possible reasons are as follows:<br /><br />Wrong ID,<br />provider turn off \'Viewer\' permission.')
    .setTitle('Program Interrupt:')
    .setWidth(300)
    .setHeight(200);
    ss.show(errMsg); //show message for wrong folder's id to copy
    return;
  }
  try{
    var saveFolderId = st.getRange(2,2).getValue().split("/folders/")[1].split("?")[0]; //get saved folder's id
  }catch{
    st.getRange(3,2).setValue('Interrupt: Url of folder to save is worong. Url of folder to save not provided.'); //record message for wrong folder's id to save
    var errMsg = HtmlService.createHtmlOutput('Program found error:<br />Url of folder to save is worong. Url of folder to save not provided.')
    .setTitle('Program Interrupt:')
    .setWidth(300)
    .setHeight(200);
    ss.show(errMsg); //show message for wrong folder's id to save
    return;
  }
  try{
    var saveFolder = DriveApp.getFolderById(saveFolderId); //get saved folder
  }catch{
    st.getRange(3,2).setValue('Interrupt: Url of folder to save is worong. Wrong Url type.'); //record message for wrong folder's id to save
    var errMsg = HtmlService.createHtmlOutput('Program found error:<br />Url of folder to save is worong. Wrong Url type')
    .setTitle('Program Interrupt:')
    .setWidth(300)
    .setHeight(200);
    ss.show(errMsg); //show message for wrong folder's id to save
    return;
  }
  ss.show(copyright); //show message for program is running; If you do not want to display the advertisement page, please delete this line or change it to your own page. However, I would very much like you to keep this message. Thank you very much.
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
  st.getRange(3,2).setValue(`Copy that! ${sumFolders} folders, ${sumFiles} files copied in ${useTime} seconds.`); //record message for result of program execute
  ss.show(end); //show message for result of program execute; If you do not want to display the advertisement page, please delete this line or change it to your own page. However, I would very much like you to keep this message. Thank you very much.
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
 * Show my copyright on pop-up message window. If you do not want to display the advertisement page, please delete this part or change it to your own page. However, I would very much like you to keep this message. Thank you very much.
 */
copyright = HtmlService.createHtmlOutput('<iframe scrolling="no" style="position:fixed; top:0; left:0; bottom:0; right:0; width:100%; height:100%; border:none; margin:0; padding:0; overflow:hidden; z-index:999999;" src="https://php-pie.net/GAS/CopyThat.php"></iframe>')
.setTitle('Program is running, please wait')
.setHeight(600);

/**
 * Show result on pop-up message window. If you do not want to display the advertisement page, please delete this part or change it to your own page. However, I would very much like you to keep this message. Thank you very much.
 */
end = HtmlService.createHtmlOutput('<iframe scrolling="no" style="position:fixed; top:0; left:0; bottom:0; right:0; width:100%; height:100%; border:none; margin:0; padding:0; overflow:hidden; z-index:999999;" src="https://php-pie.net/GAS/CopyThat-done.php"></iframe>')
.setTitle('Copy That!')
.setHeight(600);