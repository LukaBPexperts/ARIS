var nLocale   = Context.getSelectedLanguage();
var scriptAdminComponent = Context.getComponent("ScriptAdmin");
var exportLocation = "C:/ARISScriptExports";
var df = new java.text.SimpleDateFormat("yyyy_MM_dd_hh_mm");

/* CREATE OUTPUT SHEETS */
var oOutput = Context.createExcelWorkbook(Context.getSelectedFile());
var oExportedFilesSheet = oOutput.createSheet("Exported Files");
var headerRow = oExportedFilesSheet.createRow(0);
headerRow.createCell(0).setCellValue("Folder");
headerRow.createCell(1).setCellValue("File");
var oIssuesSheet = oOutput.createSheet("Issues");
headerRow = oIssuesSheet.createRow(0);
headerRow.createCell(0).setCellValue("Report Folder");
headerRow.createCell(1).setCellValue("Report Name");
headerRow.createCell(2).setCellValue("Issue description");
var intCurrentIssueRow = 1;
var intCurrentExportedFileRow = 1;

//var testFileName = "Find task \\ role connection to req in matrix model which does not have occ in the same validity model";
//var bFileOK = fileNameOk(testFileName);

var aScriptComponentInfo = scriptAdminComponent.getScriptComponents();

// get ComponentIDs
var ComponentID_CommonFiles = aScriptComponentInfo[0].getComponentID();
var ComponentID_Reports = aScriptComponentInfo[1].getComponentID();

// Set the Root folders for Common Files and Reports
var ExportFolder_CommonFiles ="Common files";
var ExportFolder_Reports ="Reports";

var outputLocation = exportLocation + "/" + df.format( new Date() );

var writeErrorCount = 0;
var errors = [];

main();

oOutput.write();

///////////////////////////////////////////////////////////////////////////////////////////////////////
// Main function
///////////////////////////////////////////////////////////////////////////////////////////////////////
function main(){
    
    // prepare target folder
    createFolder(outputLocation);
    emptyFolder(outputLocation);
    
    //Export Common Files
    var listCommonFilesScriptInfos = scriptAdminComponent.getScriptInfos(ComponentID_CommonFiles, null, nLocale);
    exportScripts(ComponentID_CommonFiles, null, 
                    ExportFolder_CommonFiles, null, listCommonFilesScriptInfos);
    
    // get all Report folders / Categories
    var reportSubFolders = scriptAdminComponent.getCategories(ComponentID_Reports, nLocale);
    for (var j=0; j<reportSubFolders.length; j++)
    {
        //Filter only Reports in Folders containing some string
        //if ( reportSubFolders[j].getName().indexOf("Equinor") != -1 || reportSubFolders[j].getName().equals("APG_Scripts") ) { 
            
            var listScriptInfos = scriptAdminComponent.getScriptInfos(ComponentID_Reports, reportSubFolders[j].getCategoryID(), nLocale);
            exportScripts(ComponentID_Reports, reportSubFolders[j].getCategoryID(),
                            ExportFolder_Reports, reportSubFolders[j].getName(), listScriptInfos);
        //}
    }
           
}
function logExportedFile(strFolder, strName){
    if(strFolder == null)
        strFolder = "Common files";
    
    var dataRow = oExportedFilesSheet.createRow(intCurrentExportedFileRow);
    dataRow.createCell(0).setCellValue(strFolder);
    dataRow.createCell(1).setCellValue(strName);
    intCurrentExportedFileRow++;
}
function logIssue(strFolder, strName, strIssue){
    if(strFolder == null)
        strFolder = "Common files";
    
    var dataRow = oIssuesSheet.createRow(intCurrentIssueRow);
    dataRow.createCell(0).setCellValue(strFolder);
    dataRow.createCell(1).setCellValue(strName);
    dataRow.createCell(2).setCellValue(strIssue);
    intCurrentIssueRow++;
}

///////////////////////////////////////////////////////////////////////////////////////////////////////
// function createFolder
///////////////////////////////////////////////////////////////////////////////////////////////////////
function createFolder(location) {
    var dir =new Packages.java.io.File(location);
    if (dir!=null) {
        if (! dir.exists()) {
             Packages.org.apache.commons.io.FileUtils.forceMkdir(dir);
        }
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////////
// function emptyFolder
///////////////////////////////////////////////////////////////////////////////////////////////////////
function emptyFolder(location) {
    var dir =new Packages.java.io.File(location);
    if (dir!=null) {
        if (dir.exists()) {
             Packages.org.apache.commons.io.FileUtils.cleanDirectory(dir);
        }
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////////
// function validateStringFilenameUsingRegex
// Checks if given filename is ok, returns true if illegal charachter is used
///////////////////////////////////////////////////////////////////////////////////////////////////////
function fileNameOk(filename) {
    var REGEX_PATTERN = "^[A-za-z0-9._-\)\(]{1,255}$";
    
    if (filename.indexOf("\\") != -1 || filename.indexOf("/") != -1) {
        return false;
    }
    return true;
}

///////////////////////////////////////////////////////////////////////////////////////////////////////
// function exportScripts
///////////////////////////////////////////////////////////////////////////////////////////////////////
function exportScripts(componentID, categoryID, exportRootFolder, subFolderName, aScriptInfos) {
    
    //Create subfolder
    if(subFolderName != null && !subFolderName.equals(""))
        createFolder(outputLocation + "/" + exportRootFolder + "/"  + subFolderName);
    
     for(var k=0; k<aScriptInfos.length; k++)
     {
        var scriptID = aScriptInfos[k].getID();
        var fileName = aScriptInfos[k].getName();
        var fileNameARX = aScriptInfos[k].getName();
        var isJSFile = true;
        
        
        
        // export *.js files and scripts
        if (fileName.indexOf(".js")!=-1 || ! aScriptInfos[k].isSimpleFile() ) {
            
            //Validate file name
            if(!fileNameOk(fileName)){
                logIssue(subFolderName, fileName, "Report name contains illegal characters!");
                continue;
            }
            
            if (fileName.indexOf(".js")==-1) {
                isJSFile = false;
                fileName = fileName + ".js";
                fileNameARX = fileNameARX + ".arx";
            }
            // script
            var zipEntry = scriptAdminComponent.exportFile( componentID, categoryID, scriptID );
            var arxByteArray = null;
            if(componentID == ComponentID_Reports)
                arxByteArray = scriptAdminComponent.exportScript(componentID, scriptID, new Array());
           
            var file =null;
            var fileARX = null;
            if(subFolderName != null){
                file =new Packages.java.io.File(outputLocation + "/" + exportRootFolder + "/" + subFolderName + "/" + fileName);
                fileARX =new Packages.java.io.File(outputLocation + "/" + exportRootFolder + "/" + subFolderName + "/" + fileNameARX);
            }
            else{
                file =new Packages.java.io.File(outputLocation + "/" + exportRootFolder + "/" + fileName);
                fileARX =new Packages.java.io.File(outputLocation + "/" + exportRootFolder + "/" + fileNameARX);
            }
            
            try {
                Packages.org.apache.commons.io.FileUtils.writeByteArrayToFile(file, zipEntry.getData());
                logExportedFile(subFolderName, fileName);
                if(arxByteArray != null && !isJSFile){
                    Packages.org.apache.commons.io.FileUtils.writeByteArrayToFile(fileARX, arxByteArray);
                    logExportedFile(subFolderName, fileNameARX);
                }
            } catch ( err ) {
                logIssue(subFolderName, fileName, err.toString());
            }
            
        }
     }

}    


