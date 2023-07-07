var workingDBName = "z_Telia Test W 2"
var publishDBName = "z_Telia Test 1"
var admUser = "system"
var dirPath = "C:\\PublishDBCleanup\\" //location of the log file
//*********************************************************************
//*********************************************************************

var nLocale = Context.getSelectedLanguage()
//log file global variables
var g_exc = Context.createExcelWorkbook(Context.getSelectedFile());
var g_table_Groups = new Array();
var g_headers_Groups = new Array("Group name", "Path", "Status");
var g_table_Models = new Array();
var g_headers_Models = new Array("Model name", "Model type", "GUID", "Path", "Status");

var selectedDatabases = ArisData.getSelectedDatabases()
var selectedDatabase = selectedDatabases[0]

main();

//write file
g_exc.write();

//create file on disk
createDir(dirPath)
var dateStr = getTodayDateAsString()
var fileName = "PublishingDBcleanup_" + dateStr + ".xlsx"
var contentByteArray = Context.getFile(Context.getSelectedFile(), Constants.LOCATION_OUTPUT)
writeUsingFileWriter(contentByteArray, dirPath, fileName)


function main(){
    
    //check if correct user is trying to start the script
    if(!checkUser()){
        Dialogs.MsgBox("Only admin user can start this script!", 0, "Aris report")
        return false;
    }
    
    //check if the started DB is TC_PublishingDB
    if (selectedDatabase.Name(nLocale) != publishDBName) {
        Dialogs.MsgBox("Report started on the wrong database!", 0, "Aris report")
        return false;
    }
    
        var btnPressed = Dialogs.MsgBox("You are about to start cleanup script!\nPress NO if you want to run simulation, without making any changes", Constants.MSGBOX_ICON_QUESTION | Constants.MSGBOX_BTN_YESNO + 512, "Message from Report")
        //var btnPressed = Constants.MSGBOX_RESULT_YES
        
        if (btnPressed == Constants.MSGBOX_RESULT_YES) {
            var simStr = ""
        }
        else{
            var simStr = " - Simulation"
        }
        
        //*********************************************************
        //*********************GROUP CHECK*************************
        //*********************************************************
        var allGroups = selectedDatabase.Find(Constants.SEARCH_GROUP)
        allGroups.reverse() //start from last group
        
        for (var i = 0; i < allGroups.length; i++) {
            if(allGroups[i].IsValid()){
                Context.writeStatus("1. Check Groups" + (i + 1) + "/" + allGroups.length)
        
                var groupName = allGroups[i].Name(nLocale)
                var groupPath = allGroups[i].Path(nLocale)
        
                var childGroups = allGroups[i].Childs()
                var modelList = allGroups[i].ModelList()
                var objDefList = allGroups[i].ObjDefList()
        
                if (childGroups.length == 0 && modelList.length == 0 && objDefList.length == 0) {
                    //delete group 
                    var parentGroup = allGroups[i].Parent()
                    if (btnPressed == Constants.MSGBOX_RESULT_YES) {                
                        var delGroup = parentGroup.Delete(allGroups[i])
                        //check if the parent group is empty, if yes delete it
                        if (delGroup && parentGroup.Childs().length == 0 && parentGroup.ModelList().length == 0 && parentGroup.ObjDefList().length == 0) {
                            // delete parent group but first remove it from list
                            var p_groupName = parentGroup.Name(nLocale)
                            var p_groupPath = parentGroup.Path(nLocale)
                            
                            var grandParentGroup = parentGroup.Parent();
                            var delGroupParent = grandParentGroup.Delete(parentGroup);
                            
                            if(delGroupParent){
                                g_table_Groups.push([p_groupName, p_groupPath, "Deleted" + simStr]);
                            }
                            else{
                                g_table_Groups.push([p_groupName, p_groupPath, "Error! Unable to delete group!"]);
                            }
                        }
                        
                    }
                    else {
                        var delGroup = true;                       
                    }
                    
                    if (delGroup) {
                        //write log group deleted
                        g_table_Groups.push([groupName, groupPath, "Deleted" + simStr]);
                    }
                    else {
                        g_table_Groups.push([groupName, groupPath, "Error! Unable to delete group!"]);
                    }
                
                }
                else {
                    g_table_Groups.push([groupName, groupPath, "Skipped" + simStr]);
                }
                allGroups[i] = null; //clear cache
            }
        }
        allGroups = null; //clear cache
    
        //PRINT LOG FILE
        printInformation_Groups("Deleted Groups");
        g_table_Groups = null; //clear cache
        g_headers_Groups = null; //clear cache
    
        //*********************************************************
        //********************MODEL CHECK**************************
        //*********************************************************
        var allModels = selectedDatabase.Find(Constants.SEARCH_MODEL)
        //var workingDB = ArisData.openDatabase(workingDBName, admUser, admPass, "dd838074-ac29-11d4-85b8-00005a4053ff", nLocale, false)
        var workingDB = ArisData.openDatabase(workingDBName, "dd838074-ac29-11d4-85b8-00005a4053ff", nLocale, true)
        for (var i = 0; i < allModels.length; i++) {
            Context.writeStatus("2. Check Models" + (i + 1) + "/" + allModels.length)
    
            var delModel = false;
    
            var modelName = allModels[i].Name(nLocale)
            var modelType = allModels[i].Type()
    
            var modelGroup = allModels[i].Group()
            var modelPath = modelGroup.Path(nLocale)
    
            var modelGUID = allModels[i].GUID()
            var modelWorkingDB = workingDB.FindGUID(modelGUID, Constants.CID_MODEL)
    
            if (!modelWorkingDB.IsValid()) {
                //model not found remove it
                var modelGroup = allModels[i].Group()
                if (btnPressed == Constants.MSGBOX_RESULT_YES) {
                    var delModel = modelGroup.Delete(allModels[i]);
                }
                else {
                    if (allModels[i].canWrite(true)) {
                        var delModel = true;
                    }
                }
                if (delModel) {
                    //write log model deleted
                    g_table_Models.push([modelName, modelType, modelGUID, modelPath, "Deleted" + simStr]);
                }
                else {
                    g_table_Models.push([modelName, modelType, modelGUID, modelPath, "Error! Unable to delete model!"]);
                }
            }
            else {
                g_table_Models.push([modelName, modelType, modelGUID, modelPath, "Skipped" + simStr]);
            }
            allModels[i] = null; //clear cache
        }
        allModels = null; //clear cache
    
        //PRINT LOG FILE
        printInformation_Models("Deleted Models");
        g_table_Models = null; //clear cache
        g_headers_Models = null; //clear cache
}

function printInformation_Groups(p_sheetName) {
    var headerFont = g_exc.createFont();
    headerFont.setFontName("headerFont");
    headerFont.setBoldweight(Constants.XL_FONT_WEIGHT_BOLD);
    var headerStyle = g_exc.createCellStyle(headerFont, Constants.BORDER_NONE, Constants.BORDER_NONE, Constants.BORDER_NONE, Constants.BORDER_NONE, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, Constants.C_TRANSPARENT, Constants.C_GREY_80_PERCENT, Constants.FILLTYPE_TRANSPARENT);
    //createCellStyle ( XlsFont p_Font, short p_TopBorder, short p_RightBorder, short p_BottomBorder, short p_LeftBorder, int p_TopBorderColor, int p_RightBorderColor, int p_BottomBorderColor, int p_LeftBorderColor, short p_HorizontalAlignment, short p_VerticalAlignment, int p_BackgroundColor, int p_ForegroundColor, short p_FillPattern )
    var sObjects = g_exc.createSheet(p_sheetName);

    var headRow = sObjects.createRow(0);
    for (var i = 0; i < g_headers_Groups.length; i++) {
        headRow.createCell(i).setCellValue(g_headers_Groups[i]);
    }

    for (var i = 0; i < g_table_Groups.length; i++) {
        Context.writeStatus("1. Print Groups" + (i + 1) + "/" + g_table_Groups.length)
        var excRow = sObjects.createRow(i + 1);
        for (var j = 0; j < g_table_Groups[i].length; j++) {
            writeExcelCell(excRow, j, j, g_table_Groups);
        }
    }

    //format work sheet
    sObjects.setDefaultColumnWidth(40);
    //sObjects.setColumnWidth(GUIDColumn,50*256); //changing columnwidth for selected columns
    //sObjects.setColumnWidth(2,25*256);
    sObjects.setZoom(3, 4);
    sObjects.createFreezePane(1, 1);

    //bold first row
    var headerCells = sObjects.getRowAt(0).getCells();
    for (var i = 0; i < headerCells.length; i++) {
        headerCells[i].setCellStyle(headerStyle);
    }

    //Wrap text and set vertical alignment
    var contentCell = sObjects.getCell(1, 0);
    if (contentCell != null) {
        contentCell.getCellStyle().setWrapText(true);
        contentCell.getCellStyle().setVerticalAlignment(Constants.VERTICAL_TOP);
        contentCell.getCellStyle().setLocked(false);
    }
    sObjects.getCell(0, 0).getCellStyle().setWrapText(true);
    sObjects.getCell(0, 0).getCellStyle().setLocked(true);

    //g_exc.setWriteProtect(true,"pwd","user");
    //sObjects.setProtect(true);
}

function printInformation_Models(p_sheetName) {
    var headerFont = g_exc.createFont();
    headerFont.setFontName("headerFont");
    headerFont.setBoldweight(Constants.XL_FONT_WEIGHT_BOLD);
    var headerStyle = g_exc.createCellStyle(headerFont, Constants.BORDER_NONE, Constants.BORDER_NONE, Constants.BORDER_NONE, Constants.BORDER_NONE, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, Constants.C_TRANSPARENT, Constants.C_GREY_80_PERCENT, Constants.FILLTYPE_TRANSPARENT);
    //createCellStyle ( XlsFont p_Font, short p_TopBorder, short p_RightBorder, short p_BottomBorder, short p_LeftBorder, int p_TopBorderColor, int p_RightBorderColor, int p_BottomBorderColor, int p_LeftBorderColor, short p_HorizontalAlignment, short p_VerticalAlignment, int p_BackgroundColor, int p_ForegroundColor, short p_FillPattern )
    var sObjects = g_exc.createSheet(p_sheetName);

    var headRow = sObjects.createRow(0);
    for (var i = 0; i < g_headers_Models.length; i++) {
        headRow.createCell(i).setCellValue(g_headers_Models[i]);
    }

    for (var i = 0; i < g_table_Models.length; i++) {
        Context.writeStatus("2. Print Models" + (i + 1) + "/" + g_table_Models.length)
        var excRow = sObjects.createRow(i + 1);
        for (var j = 0; j < g_table_Models[i].length; j++) {
            writeExcelCell(excRow, j, j, g_table_Models);
        }
    }

    //format work sheet
    sObjects.setDefaultColumnWidth(40);
    //sObjects.setColumnWidth(GUIDColumn,50*256); //changing columnwidth for selected columns
    //sObjects.setColumnWidth(2,25*256);
    sObjects.setZoom(3, 4);
    sObjects.createFreezePane(1, 1);

    //bold first row
    var headerCells = sObjects.getRowAt(0).getCells();
    for (var i = 0; i < headerCells.length; i++) {
        headerCells[i].setCellStyle(headerStyle);
    }

    //Wrap text and set vertical alignment
    var contentCell = sObjects.getCell(1, 0);
    if (contentCell != null) {
        contentCell.getCellStyle().setWrapText(true);
        contentCell.getCellStyle().setVerticalAlignment(Constants.VERTICAL_TOP);
        contentCell.getCellStyle().setLocked(false);
    }
    sObjects.getCell(0, 0).getCellStyle().setWrapText(true);
    sObjects.getCell(0, 0).getCellStyle().setLocked(true);

    //g_exc.setWriteProtect(true,"pwd","user");
    //sObjects.setProtect(true);

}

function writeExcelCell(p_row, p_tableIndex, p_columnIndex, g_table) {
    var rowIndex = p_row.getRowIndex();
    var excCell = p_row.createCell(p_columnIndex);
    var cellData = "";
    if (isArray(g_table[rowIndex - 1][p_tableIndex])) {
        for (i1 = 0; i1 < g_table[rowIndex - 1][p_tableIndex].length; i1++) {
            cellData = cellData + g_table[rowIndex - 1][p_tableIndex][i1] + "; " + "\n";
        }
        var nr = cellData.length;
        cellData = cellData.slice(0, nr - 3);

    }
    else {
        cellData = g_table[rowIndex - 1][p_tableIndex];
    }
    excCell.setCellValue(cellData);
}

function isArray(testObject) {
    return testObject && !(testObject.propertyIsEnumerable('length')) && typeof testObject === 'object' && typeof testObject.length === 'number';
}

function createDir(path) {
    var directory = new java.io.File(path);
    if (!directory.exists()) {
        directory.mkdir();
    }
    //var files = directory.listFiles();

    //for (var i = 0; i < files.length; i++) {
    //files[i].delete()
    //}
}

function writeUsingFileWriter(byteArray, pathName, fileName) {
    var file = new java.io.File(pathName + fileName);

    // Check if file with same name already exists
    if (file.exists()) {
        var i = 1;
        var newFileName = fileName;

        // Rename file by appending a number until a unique name is found
        while (file.exists()) {
            var dotIndex = fileName.lastIndexOf(".");
            newFileName = fileName.slice(0, dotIndex) + " (" + i + ")" + fileName.slice(dotIndex);
            file = new java.io.File(pathName + newFileName);
            i++;
        }
    }

    var fr = new java.io.FileOutputStream(file);

    fr.write(byteArray);
    fr.close();
}

function getTodayDateAsString() {
    var today = new Date();
    var year = today.getFullYear().toString().slice(-2);
    var month = (today.getMonth() + 1).toString();
    var day = today.getDate().toString();

    // Add leading zero to month if necessary
    if (month.length === 1) {
        month = "0" + month;
    }

    // Add leading zero to day if necessary
    if (day.length === 1) {
        day = "0" + day;
    }

    return year + month + day;
}

function checkUser(){
    var currentUser = ArisData.getActiveUser()
    if(currentUser.Name(nLocale) != admUser){
        return false;
    }
    return true;
}