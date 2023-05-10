var FSO = WScript.CreateObject("Scripting.FileSystemObject");
var log = FSO.OpenTextFile("log.txt", 8, true);
var iterator = new Enumerator(WScript.Arguments);
for (; !iterator.atEnd(); iterator.moveNext()) {
    var file = iterator.item();
    if(file != ""){
        try{
            renameFile(file);
        } catch(error){
            WScript.Echo("Error= " + error.message);
        }
    } else break;
}

function renameFile(initialPath){
    var fileName = FSO.GetBaseName(initialPath);
    var fileExt = FSO.GetExtensionName(initialPath);
    fileName = "FILE_" + fileName + "_OK." + fileExt;
    var newPath = FSO.BuildPath("C:\\Users\\MyPC\\Documents\\Scripts", fileName);
    var report = "";
    try{
        FSO.CopyFile(initialPath, newPath);
        report = initialPath + " ==> " + newPath + " at " + getNow();
    } catch(error){
        WScript.Echo("CopyFile Error= " + error.message);
        report = "\tfailed to copy: " + initialPath + " at " + getNow();
    }
    log.WriteLine(report);
    return fileName;
}

function getNow(){
    var d = new Date();
    var segments = d.toString().split(" ");
    //get the time and date in format "hh:mm:ss MMM DD,YYYY"
    var now = segments[3] + " " +  segments[1] + " " + segments[2] + "," + segments[5]; 
    return now;
}