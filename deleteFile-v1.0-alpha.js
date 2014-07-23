fso = new ActiveXObject("Scripting.FileSystemObject");

var today = new Date();         //get current date
//參數設定
var timeUnit = "isDay";         //time unit configuration，value available: isSecond,isMinute,isHour,isDay,isMonth,isYear
var delTime = 1;                // use it with timeUnit to determine the files before which time to delete.
var startFolder = "c:\\storage";        //the target folder
var isBak = false;              //choose backup or not
var bakFolder = "C:\\storageBak" + "-" + today.toLocaleDateString();    //backup folder name, use date to name it.

//parameter init, do not change it!
var fileCounter = 0;            //file counter
var folderCounter = 0;          //folder counter
var errorCounter = 0;           //error counter
var result = "";                //operation reslut
var output = new ResultWriter();    //result writer

//operation start!
//caculate which time you choose to delete files before it
var delDate = getDelDate();

//if choose backup and the backup folder isn't exist, create it, else create a new folder with current time.
if(isBak){              
    if(fso.FolderExists(bakFolder) == false){
        fso.CreateFolder(bakFolder);
    }else{
        bakFolder += today.getHours() + "H" + today.getMinutes() + "M" + today.getSeconds() + "S";
        fso.CreateFolder(bakFolder);
    }
}

output.line("Delete files before " + delDate.toLocaleString() + ".");
output.line("Program starts at " + today.toLocaleString() + ".");
output.line("The target folder is " + startFolder + ".");

if(isBak){
    output.line("The Backup folder is " + bakFolder + ".");
}

output.line("The detail information is as follows: ");

DeleteOldFiles(startFolder,delDate);         //delete files operation
DeleteEmptyFolders(startFolder);            //delete empty folders

//Task finish window, delete the annotion characters to enable it.
//WScript.Echo("Task Finshed！\n, deleted " + fileCounter + " file(s), " + folderCounter + "folder(s), "
//    + "you can check the detail information in the log file.");
output.line("");
output.line("Delete " + fileCounter + " file(s), " + folderCounter + " folder(s).");
output.line((fileCounter - errorCounter) + " file(s) deleted successful, " + errorCounter + "file(s) failed.");  

function DeleteOldFiles(folderName,date){
    var folder,selFile,fileCollection;
    try{
        folder = fso.GetFolder(folderName);
    }catch(e1){
        output.line("Error: " + e1.description + "\r\n");
        output.line("The target folder has an error, please check it.\r\n");
        return;
    }
    fileCollection = folder.Files;
    var e = new Enumerator(fileCollection); 
    for(;!e.atEnd();e.moveNext()){
        var selFile = e.item();
        
        if(selFile.DateCreated <= date){
            //operation log
            fileCounter++;
            output.line(fileCounter + ":" + selFile.Name + " in " + selFile.ParentFolder + "\r\n");
            output.line("create time: " + selFile.DateCreated+ "\r\n");
            output.line("last accessed: " +selFile.DateLastAccessed + "\r\n")
            output.line("last modified: " + selFile.DateLastModified + "\r\n")

            if(isBak == true){
                //create new path for backup
                var flPath = selFile.Path.substring(startFolder.length,selFile.Path.length - selFile.Name.length);
                var newPath = bakFolder + flPath;
                if(!fso.FolderExists(newPath)){
                    fso.CreateFolder(newPath);
                }
                fso.CopyFile(selFile.Path,newPath,true);
            }
            //delete raw files
            try{
                fso.deleteFile(selFile.path,true);
            }catch(e2){
                output.line("Error: " + e2.description + "\r\n");
                output.line("The file may be in use, continue for next file.\r\n");
                errorCounter++;
            }finally{
                output.line(result);
                continue;
            }
        }
    }
    var enumSubFolder = new Enumerator(folder.SubFolders);
    //opearation in sub folders
    for(;!enumSubFolder.atEnd();enumSubFolder.moveNext()){
        DeleteOldFiles(enumSubFolder.item().Path,date);
    }
}
function DeleteEmptyFolders(folderName){
    var folder = fso.GetFolder(folderName);
    if(folder.Files.Count == 0 && folder.SubFolders.Count == 0){
        output.line(folder.Name + " in " + folder.ParentFolder + "\r\n");
        output.line("create time: " + folder.DateCreated + "\r\n");
        output.line("last accessed: " + folder.DateLastAccessed + "\r\n");
        output.line("last modified: " + folder.DateLastModified + "\r\n")
        fso.DeleteFolder(folder.Path);
        folderCounter++;
    }else if(folder.SubFolders.Count != 0){
        var enumSubFolder = new Enumerator(folder.SubFolders);
        for(;!enumSubFolder.atEnd();enumSubFolder.moveNext()){
            DeleteEmptyFolders(enumSubFolder.item());
        }
    }
}
//calculate delete date
function getDelDate(){
    
    var OlderThanDate = new Date();
    var time = today.getTime();     //get current time
    var MinMilli = 1000 * 60;       //the millseconds in one minute
    var HrMilli = MinMilli * 60;    //the millseconds in one hour
    var DyMilli = HrMilli * 24;     //the millseconds in one day

    switch(timeUnit){
        case "isSecond" :
            OlderThanDate.setTime(time - (delTime * 1000));break;
        case "isMinute" :
            OlderThanDate.setTime(time - (delTime * MinMilli));break;
        case "isHour"   :
            OlderThanDate.setTime(time - (delTime * HrMilli));break;
        case "isDay"    :
            OlderThanDate.setTime(time - (delTime * DyMilli));break;
        case "isMonth" :
            OlderThanDate.setMonth(today.getMonth() - delTime);break;
        case "isYear"   :
            OlderThanDate.setYear(today.getYear() - delTime);break;
    }
    return OlderThanDate;
}
//result writer
function ResultWriter(){
    var savepath = WScript.ScriptFullName.substr(0,(WScript.ScriptFullName.length-WScript.ScriptName.length));
    var rflPath = savepath+"deleteFiles-" + today.toLocaleDateString();
    if(fso.FileExists(rflPath + ".log")){
        ResultFile = fso.CreateTextFile(rflPath + today.getHours() + "時" + today.getMinutes() + "分" + today.getSeconds() + "秒.log",false,true);
    }else{
        ResultFile = fso.CreateTextFile(rflPath + ".log",false,true);
    }
    this.file=ResultFile;
    this.line=ResultWriter_Line;
}
function ResultWriter_Line(strings){
    lineFeed = "\r\n";
    this.file.Write(strings + lineFeed);
}
