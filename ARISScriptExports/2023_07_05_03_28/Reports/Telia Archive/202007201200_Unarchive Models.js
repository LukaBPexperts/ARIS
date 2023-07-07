var prodDbName = "Archive_Test"
var admUser = "system"
var admPass = "manager"

var nLocale = Context.getSelectedLanguage()
//var oOutput = Context.createOutputObject()

//log file global variables
var g_exc = Context.createExcelWorkbook(Context.getSelectedFile());
var g_table = new Array();
var g_headers = new Array("Model", "Model Type", "Location", "Person Responsible", "Model Archived", "Log");

var selectedDatabase = ArisData.getActiveDatabase()

selectedDatabase.setAutoTouch(false);

//FIND MODELS MARKED FOR UNARCHIVE
var searchItemSpecial = selectedDatabase.createSearchItem(Constants.AT_STATE_1, nLocale, "UNARCHIVE", Constants.SEARCH_CMP_EQUAL, false, false) //TODO
var selectedModels = selectedDatabase.Find(Constants.SEARCH_MODEL, null, searchItemSpecial)
//SPECIAL FIND MULTI MIX
//var modelsForArchive = design_DB.Find(Constants.SEARCH_MODEL, [Constants.MT_EEPC_COLUMN], searchItemSpecial1.and(searchItemSpecial2.and(searchItemSpecial3)))

//CHECK MODELS, ASSIGMENTS, AND ATTRIBUTE and add them to log file as array
//oTable.push([oCurrModelName, oCurrModelPath, oSourceOccName, oTargetOccName, "Model locked", false]);
//var modelsForArchive = checkModels(selectedModels)
var modelsForUnarchive = selectedModels

//*****************************************************
var btnPressed = Dialogs.MsgBox(modelsForUnarchive.length + " model will be unarchived. Do you want to continue?", Constants.MSGBOX_ICON_QUESTION | Constants.MSGBOX_BTN_YESNO + 512, "Message from Report")
if (btnPressed == Constants.MSGBOX_RESULT_YES) {

    //MERGE MODELS TO PROD DB
    var prod_DB = ArisData.openDatabase(prodDbName, admUser, admPass, "dd838074-ac29-11d4-85b8-00005a4053ff", nLocale, false)
    var merge = Context.getComponent("Merge")
    var mergeOptions = merge.createMergeOptions(selectedDatabase)
    mergeOptions.setAssignmentDepth(0);
    mergeOptions.setConsiderBorderAssignment(true);
    mergeOptions.setConsiderBorderCxnDefs(true);
    mergeOptions.setConsiderSourceGroupPath(true);
    mergeOptions.setCxnDefSourceWins(true);
    mergeOptions.setFontSourceWins(true);
    mergeOptions.setInheritAccessRightsForNewGroups(true);
    mergeOptions.setMixAttributes(false);
    mergeOptions.setModelSourceWins(true);
    mergeOptions.setObjDefSourceWins(true);
    mergeOptions.setSynchronizeAssignments(false);

    for (var i = 0; i < modelsForUnarchive.length; i++) {

        var mergeResult = merge.startMerge([modelsForUnarchive[i]], prod_DB, mergeOptions, true)

        if (mergeResult.isSuccessful()) {
            var groupTemp = modelsForUnarchive[i].Group()
            var deleteModel = groupTemp.Delete(modelsForUnarchive[i])
            g_table.push([modelsForUnarchive[i].Name(nLocale), modelsForUnarchive[i].Type(), modelsForUnarchive[i].Group().Path(nLocale), modelsForUnarchive[i].Attribute(Constants.AT_PERS_RESP, nLocale).getValue(), mergeResult.isSuccessful(), "Model Unarchived"]);
        } else {
            g_table.push([modelsForUnarchive[i].Name(nLocale), modelsForUnarchive[i].Type(), modelsForUnarchive[i].Group().Path(nLocale), modelsForUnarchive[i].Attribute(Constants.AT_PERS_RESP, nLocale).getValue(), mergeResult.isSuccessful(), "Merge Error"]);
        }
    }


    prod_DB.clearCaches()
}

selectedDatabase.setAutoTouch(true);

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