![Copy That!](https://www.php-pie.net/images/CopyThat.png "Copy That!")
# Download 下載使用
👉 [English version](https://docs.google.com/spreadsheets/d/1oWTF7TRUZlMzjav9LxExCMWfXQuLtIKh_PFhP3j8syE/copy) 👉 [中文版本](https://docs.google.com/spreadsheets/d/1Zn68t9-4FbS2dqQwphm-0xGUUFbKlIuUQ_jLFeHzcWM/copy)
# Features 特徵
1. No installations, no bloatware, no updates: this works in any modern browser, including Google Chrome, Firefox, Edge and Safari.  
不需要安裝，不需要外掛，不用升級，適用於最新版本的各種瀏覽器。
2. Through the program, you can copy the Google drive data shared by others (including files and folders of different layers) to your own Google drive intact.  
透過程式，你可以將他人共用的雲端硬碟資料（包括檔案與不同層的資料夾）原封不動地複製到自己的雲端硬碟中。
3. Files in Google format in the Google drive, such as gdoc, gsheet, gslides, gscript, etc., will not be forced to convert to Microsoft Office format or cannot be downloaded (copied).  
雲端硬碟中Google格式的檔案，比如gdoc、gsheet、gslides、gscript等，文件格式將不會被迫轉換為微軟office的格式或者無法下載(複製)。
4. The copied spreadsheet can be used in anyone's Google drive, and you can also modify the program by yourself to add other functions.  
建立副本的試算表可以在任何Google帳號的雲端硬碟上使用，您也可以自行修改程式，增加其他的功能。
# Foreword 前言
This spreadsheet application script is mainly used to solve the multi-level folder and file copy function that Google Drive does not provide. Through the program, it can be performed between two different Google accounts and different folders, or the same account copying of files and folders between different folders.  
The program structure and process are unexpectedly simple, and the time spent writing the program is shorter than I searched for various solutions on the Internet. This fact surprised me, a rookie programmer.  
這支試算表應用腳本主要是用來解決Google雲端硬碟不提供的多層次資料夾與檔案複製的功能，透過程式能在兩個不同的Google帳號，或同一個帳號不同資料夾之間進行檔案與檔案夾的複製。  
程式架構與流程出乎意料的簡單，撰寫程式所花費的時間，比我在網路上搜尋各種解決方案還要短，這件事讓我這個菜鳥程式設計人員感到驚訝。
# Restricted 限制
There is a big problem with this program that needs special attention, otherwise it will be a disaster for users. Do not use a third party to perform the copy for the other two, otherwise a complete copy of files and folders will also appear in the root directory of your own Google drive.  
> Example:<br />User A copies 10 folders and 500 files shared by user B to the cloud drive of user C through this program, and user C will successfully receive 10 folders and 500 files , but user A will also get 10 folders and 500 files in the root directory of Google drive.  

這支程式有一個很大的問題要特別注意，否則對使用者來說將是一場災難。不要以第三方替其他兩人執行複製(不要當媒人)，否則在您自己的雲端硬碟根目錄中，也會得到一份完整的檔案與資料。
>舉例：<br />使用者A透過這個程式，將使用者B分享的10個資料夾與500個檔案複製到使用者C的雲端硬碟，使用者C會成功地收到10個資料夾與500個檔案，但是使用者A的雲端硬碟根目錄中同樣會得到10個資料夾與500個檔案。
# How to use 使用說明
1. Please copy the "CopyThat"(English version) spreadsheet to your Google drive by "Make a copy".  
請以建立副本的方式，將「拷貝那個」(中文版)試算表複製到自己的雲端硬碟中。
<img src="https://www.php-pie.net/images/gas/copythat/copythat-001.gif" alt="Please copy the 'CopyThat' spreadsheet to your Google drive by 'Make a copy'." />
2. After opening the spreadsheet, wait for the "Advanced" function on menu to be displayed, if it is not displayed, please refresh the page once.  
打開試算表之後，等待「進階功能」功能顯示，若沒有顯示請重新整理網頁一次。
<img src="https://www.php-pie.net/images/gas/copythat/copythat-002.gif" alt="After opening the spreadsheet, wait for the 'Advanced' function on menu to be displayed, if it is not displayed, please refresh the page once." />
3. Paste the "**Url** of folder to copy" into the B1 cell, and paste the "**Url** of folder to save" into the B2 cell.  
將「複製資料夾的連結」貼進B1的儲存格，將「存放資料夾的連結」貼近B2的儲存格。
<img src="https://www.php-pie.net/images/gas/copythat/copythat-003.gif" alt="Paste the 'Url of folder to copy' into the B1 cell, and paste the 'Url of folder to save' into the B2 cell." />
3. On first run, the program needs your authorization to gain access to Drive.  
第一次執行時，程式需要您的授權，才能取用雲端硬碟的存取權限。
<img src="https://www.php-pie.net/images/gas/copythat/copythat-004.gif" alt="On first run, the program needs your authorization to gain access to Drive." />
<img src="https://www.php-pie.net/images/gas/copythat/copythat-004-1.gif" alt="On first run, the program needs your authorization to gain access to Drive." />
5. Click "Make a Copy" in the advanced function menu, and wait for the program to complete.  
點選進階功能選單中的「拷貝」，靜候程式執行完成即可。
<img src="https://www.php-pie.net/images/gas/copythat/copythat-005.gif" alt="Click 'Make a Copy' in the advanced function menu, and wait for the program to complete." />

# Troubleshoot 錯誤與解決
When the program was designed, various situations were preset, such as forgetting to enter the url of the folder, the format of the link was wrong, or the permissions was not shared for the folder to copy, etc. When encountering the above errors, the program will display the error message in a dialogue window or written in "Execution Status". Please follow the message description to troubleshoot the problem.  
當初在程式設計的時候，已經有預設各種狀況，比如忘記輸入資料夾的連結、連結的格式錯誤，或者複製的資料夾未開啟共用權限等，遇到以上的錯誤時，程式會將錯誤訊息以對話視窗或者寫在「程式執行情形」方式顯示，請您依照訊息說明，排除問題即可。
# Visit & Sponsor 參觀與贊助
Welcome to my website [PHP-Pie](https://php-pie.net "PHP-Pie"), there may be some tools, programs, or even inspiration that can help you. We are more than happy to accept your [sponsor](https://p.ecpay.com.tw/36FF207 "sponsor")ship if you wish.:heart:  
歡迎參觀我的網站 [PHP-Pie](https://php-pie.net "PHP-Pie")，裡頭也許會有一些能幫助您的工具、程式，甚至是提供靈感。如果您願意的話，也非常樂意接受您的[打賞](https://p.ecpay.com.tw/36FF207 "打賞")。:heart:  
Copyright (c) 2023 Chang, Chia-Cheng 張家誠
# About Name 關於名稱
![Program's names](https://www.php-pie.net/images/gas/copythat/programName.png "Program's names")  
When I first picked the name of the program, because I wanted to move files from user A’s Google drive to user B’s Google drive, I thought of the homophonic Japanese word Shogun (MOVE Sogun). However, in fact, it should be "copy" rather than "move", so I thought of the homonym of the great basketball superstar Kobe Bryant and named it Copy Bryant.  
Later, I felt that using Kobe’s name might cause copyright infringement. After thinking hard for a while, I thought of the word porter the homonym for Harry Potter’s Potter. So I added the circular glasses intention to the logo design. However, I decided to use the more reasonable "copy", so I finally decided on the name "CopyThat!".  
"Copy that!" is a military term in British and American English. It originated from the meaning of copying and receiving a telegram in the past. Now it is generally used to mean "received/clear/understand". The purpose of this program is to copy completely and help you copy the files and folders you specify.  
However, thinking up the name takes far more time than writing the program!  
當初取程式的名稱的時候，因為想向要將檔案從使用者A的雲端硬碟MOVE到使用者B的雲端硬碟，所以聯想到了諧音日文的幕府將軍(MOVE Sogun)。但是，事實上應該是「複製」而不是「移動」，因此又想到了諧音偉大的籃球巨星柯比布萊恩，取名Copy Bryant。  
後來覺得使用柯比的名字可能會有侵權的疑慮，苦思了一段時間之後，想到了搬運工這個詞，英文是porter，諧音為哈利波特的波特，所以就把圓形的眼鏡意向加入logo的設計中。不過，還是決定採用較合理的「複製」，於是最後才終於決定取名為「拷貝那個！」。  
「拷貝那個！」是英美語中的軍事術語，源自於過去接收電報時表示完整抄收的意思，現在一般用做「收到/清楚了/明白」的意思。這支程式的目的，就是完全抄收，幫您拷貝那個您指定的檔案與資料夾。  
然而，想名稱所花費的時間，遠遠超過寫程式的時間！
