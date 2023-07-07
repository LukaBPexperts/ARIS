//*******************************************************************************************************************************************
//******************************************************CUSTOMIZE****************************************************************************
//*******************************************************************************************************************************************
var adminUsername = "arisservice" //admin username (user must have access to the whole database)
var adminPassword = "arisservice" // admin user password
var entireMethodFilterGUID = "dd838074-ac29-11d4-85b8-00005a4053ff" //entire method filter GUID

//OBJECTS WHICH ARE NOT GOING TO BE MOVED
var excludedObjects = [Constants.OT_APPL_SYS_TYPE, Constants.OT_APPL_SYS_CLS, Constants.OT_ENT_TYPE, Constants.OT_TECH_TRM, Constants.OT_CLS, Constants.OT_ORG_UNIT, Constants.OT_LOC]
//*******************************************************************************************************************************************
//*******************************************************************************************************************************************

//**********************************************************************************************************************
//**************************************DON'T CHANGE ANYTHING BELOW THIS LINE********************************************
//**********************************************************************************************************************

var nLocale = Context.getSelectedLanguage()

//log file global variables
var g_exc = Context.createExcelWorkbook(Context.getSelectedFile());
var g_table = new Array();
var g_headers = new Array("Name", "Model/Object", "Type", "Old Location", "New Location", "Moved", "Log");

var selectedModels = ArisData.getSelectedModels()

var dbName = ArisData.getActiveDatabase().Name(Context.getSelectedLanguage())
var selection = Dialogs.BrowseArisItems("Select group", "Select the target group", "(servername is obsolete)", dbName, Constants.CID_GROUP)
var aOIDs = selection.split(";")

var dbAdm = ArisData.openDatabase(dbName, adminUsername, adminPassword, entireMethodFilterGUID, Context.getSelectedLanguage(), false);

var targetGroup = dbAdm.FindOID(aOIDs[0])

for (var i = 0; i < selectedModels.length; i++) {
    var model = dbAdm.FindGUID(selectedModels[i].GUID())
    var oldGroupModel = model.Group().Path(nLocale)

    var objDefs = model.ObjDefList()
    var filteredObjDefs = filterObjDefs(objDefs, excludedObjects)


    //move model
    var moveModel = false
    try {
        var moveModel = model.ChangeGroup(targetGroup)
    } catch (e) {}
    if (moveModel != false) {
        //LOG FOR MODEL MOVED   
        g_table.push([model.Name(nLocale), "Model", model.Type(), oldGroupModel, model.Group().Path(nLocale), moveModel, "Model successfully moved"]);
    } else {
        //LOG FOR MODEL NOT MOVED   
        g_table.push([model.Name(nLocale), "Model", model.Type(), oldGroupModel, "", moveModel, "Error, model is not moved"]);
    }

    //move objects
    for (var j = 0; j < filteredObjDefs.length; j++) {
        oldGroupObject = filteredObjDefs[j].Group().Path(nLocale)
        var moveObject = false
        try {
            var moveObject = filteredObjDefs[j].ChangeGroup(targetGroup)
        } catch (e) {}

        if (moveObject != false) {
            //LOG FOR OBJECT MOVED 
            g_table.push([filteredObjDefs[j].Name(nLocale), "Object", filteredObjDefs[j].Type(), oldGroupObject, filteredObjDefs[j].Group().Path(nLocale), moveObject, "Object successfully moved"]);
        } else {
            //LOG FOR OBJECT NOT MOVED   
            g_table.push([filteredObjDefs[j].Name(nLocale), "Object", filteredObjDefs[j].Type(), oldGroupObject, "", moveObject, "Error, object is not moved"]);
        }
    }
}

//PRINT LOG FILE
printInformation("Log");
g_exc.write();

function printInformation(p_sheetName) {

    var headerFont = g_exc.createFont();
    headerFont.setFontName("headerFont");
    headerFont.setBoldweight(Constants.XL_FONT_WEIGHT_BOLD);
    var headerStyle = g_exc.createCellStyle(headerFont, Constants.BORDER_NONE, Constants.BORDER_NONE, Constants.BORDER_NONE, Constants.BORDER_NONE, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, Constants.C_TRANSPARENT, Constants.C_GREY_80_PERCENT, Constants.FILLTYPE_TRANSPARENT);
    //createCellStyle ( XlsFont p_Font, short p_TopBorder, short p_RightBorder, short p_BottomBorder, short p_LeftBorder, int p_TopBorderColor, int p_RightBorderColor, int p_BottomBorderColor, int p_LeftBorderColor, short p_HorizontalAlignment, short p_VerticalAlignment, int p_BackgroundColor, int p_ForegroundColor, short p_FillPattern )
    var sObjects = g_exc.createSheet(p_sheetName);

    var headRow = sObjects.createRow(0);
    for (var i = 0; i < g_headers.length; i++) {
        headRow.createCell(i).setCellValue(g_headers[i]);
    }

    for (var i = 0; i < g_table.length; i++) {
        var excRow = sObjects.createRow(i + 1);
        for (var j = 0; j < g_table[i].length; j++) {
            writeExcelCell(excRow, j, j);

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
/*--------------------------------------------
Help functions
------------------------------------------*/
function writeExcelCell(p_row, p_tableIndex, p_columnIndex) {
    var rowIndex = p_row.getRowIndex();
    var excCell = p_row.createCell(p_columnIndex);
    var cellData = "";
    if (isArray(g_table[rowIndex - 1][p_tableIndex])) {
        for (i1 = 0; i1 < g_table[rowIndex - 1][p_tableIndex].length; i1++) {
            cellData = cellData + g_table[rowIndex - 1][p_tableIndex][i1] + "; " + "\n";
        }
        var nr = cellData.length;
        cellData = cellData.slice(0, nr - 3);

    } else {
        cellData = g_table[rowIndex - 1][p_tableIndex];
    }
    excCell.setCellValue(cellData);
}


function isArray(testObject) {
    return testObject && !(testObject.propertyIsEnumerable('length')) && typeof testObject === 'object' && typeof testObject.length === 'number';
}


function filterObjDefs(objDefs, excludedObjects) {
    var filteredObjDefs = new Array()
    for (var i = 0; i < objDefs.length; i++) {
        var flag = false
        for (var j = 0; j < excludedObjects.length; j++) {
            if (objDefs[i].TypeNum() == excludedObjects[j]) {
                var flag = true
            }
        }
        if (flag != true) {
            filteredObjDefs.push(objDefs[i])
        }
    }
    return filteredObjDefs
}