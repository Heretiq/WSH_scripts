var FSO = WScript.CreateObject("Scripting.FileSystemObject");
var log = FSO.OpenTextFile("log.txt", 8, true);
var rootFolder = "C:\\Users\\MyPC\\Documents\\Scripts";

//получение элементов текущей даты
var NOW = new Date();
var YEAR = NOW.getYear();
var MONTH = addZero(NOW.getMonth()+1);
var DAY = addZero(NOW.getDate());
//создание иерархии вложенных подкаталогов по текущей дате
var yearFolder = createSubfolder(rootFolder, YEAR);
var monthFolder = createSubfolder(yearFolder, YEAR + "-" + MONTH);
var dayFolder = createSubfolder(monthFolder, YEAR + "-" + MONTH + "-" + DAY);

//перебор файлов, переданных в скрипт
var iterator = new Enumerator(WScript.Arguments);
for (; !iterator.atEnd(); iterator.moveNext()) {
    var file = iterator.item();
    if(file != ""){  
        try{
            moveFile(file, dayFolder);
        } catch(error){
            WScript.Echo("Move error = " + error.message);
        }
    } else break;
}
//открытие вновь созданной конечной папки
var sh = WScript.CreateObject("WScript.Shell");
sh.Run(format("explorer.exe \"${dayFolder}\"", [dayFolder]));

//создает вложенный подкаталог
function createSubfolder(parent, child){
    var childFolder = FSO.BuildPath(parent, child);
    try{
        return FSO.CreateFolder(childFolder);
    } catch(error){
        return childFolder;
    }
}

//добавляет 1 к номеру месяца (т.к. WSH нумерует месяцы от 0 до 11)
function addZero(str){
    if(str.toString().length === 1){
        str = "0" + str;
    }
    return str;
}

//подготовка файла к перемещению и трансфер
function moveFile(filePath, newFolder){
    var report = "";
    var fileName = FSO.GetFileName(filePath); 
    newPath = renameFile(fileName, newFolder);
    try{
        FSO.MoveFile(filePath, newPath);
        report = fileName + " ==> " + newFolder + " at " + getNow();
    } catch(error){
        report = "\tfailed to move " + fileName + " to " + newFolder;
    }
    log.WriteLine(report);
    return report;
 }

 //проверяет, требуется ли переименование файла в связи с обнаружением файла с таким же именем в целевой папке
 function renameFile(fileName, newFolder){
    var counter = 1;
    var newPath = FSO.BuildPath(newFolder, fileName);
    while(FSO.FileExists(newPath)){
        var baseName = FSO.GetBaseName(fileName);
        //если собственно имя файла заканчивется числом в скобках, то вместо этого числа подставляется counter
        if(baseName.match(/\(\d\)$/)){
            baseName = baseName.replace(/\(\d\)$/, "(" + counter + ").");
        } else{
            //иначе к собственно имени файла добавляется число в скобках
            baseName = baseName + "(" + counter + ").";
        } 
        var altFileName = baseName + FSO.GetExtensionName(fileName);
        newPath = FSO.BuildPath(newFolder, altFileName);
        counter++;
    }
    return newPath;
 }
 
 //получает текущую дату
 function getNow(){
    var d = new Date();
    var segments = d.toString().split(" ");
    //get the time and date in format "hh:mm:ss MMM DD,YYYY"
    var now = segments[3] + " " +  segments[1] + " " + segments[2] + "," + segments[5]; 
    return now;
}

//сокращенная запись
function format(s, array){var items = s.split(/\$\{.*?\}/);var result="";for(var i = 0; i < array.length; i++){result += items[i]+array[i];}if(items.length > array.length){result += items[items.length - 1];}return result;}