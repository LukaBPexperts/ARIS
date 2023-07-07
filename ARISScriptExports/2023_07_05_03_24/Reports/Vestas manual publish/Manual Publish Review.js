var nLocale = Context.getSelectedLanguage()
var selectedGroups = ArisData.getSelectedGroups()

//log file global variables
var g_exc = Context.createExcelWorkbook(Context.getSelectedFile());
var g_table = new Array();
var g_headers = new Array("Model Name", "Model Type", "Location", "Log");

var publishModels = new Array()
for (var i = 0; i < selectedGroups.length; i++) {

    var publishModelsTemp = selectedGroups[i].ModelList(true, null)

    var publishModels = publishModels.concat(publishModelsTemp)

}

publishModels = ArisData.Unique(publishModels)

var btnPressed = Dialogs.MsgBox(publishModels.length + " models will be published . Do you want to continue?", Constants.MSGBOX_ICON_QUESTION | Constants.MSGBOX_BTN_YESNO + 512, "Message from Report")
if (btnPressed == Constants.MSGBOX_RESULT_YES) {

    //*****************************************************
    //MERGE MODELS TO Publishied DB
    //var review_DB = ArisData.openDatabase("Review", "system", "tutqL4LWQADAs33W", "dd838074-ac29-11d4-85b8-00005a4053ff", nLocale, false)
    //var design_DB = ArisData.openDatabase("Design", "system", "tutqL4LWQADAs33W", "dd838074-ac29-11d4-85b8-00005a4053ff", nLocale, false)
    
    var review_DB = ArisData.openDatabase("Review", "system", "manager", "dd838074-ac29-11d4-85b8-00005a4053ff", nLocale, false)
    var design_DB = ArisData.openDatabase("Design", "system", "manager", "dd838074-ac29-11d4-85b8-00005a4053ff", nLocale, false)

    //var published_DB = ArisData.openDatabase("Published", "system", "manager", "dd838074-ac29-11d4-85b8-00005a4053ff", nLocale, false) //FOR TEST
    //var design_DB = ArisData.openDatabase("Design", "system", "manager", "dd838074-ac29-11d4-85b8-00005a4053ff", nLocale, false) //FOR TEST

    if (publishModels.length > 0) {
        var publishModelsGUID = getModelsGUID(publishModels)
        var mergeModels = design_DB.FindGUID(publishModelsGUID)

        var merge = Context.getComponent("Merge")
        var mergeOptions = merge.createMergeOptions(design_DB)
        mergeOptions.setAssignmentDepth(0) //***this MUST be zero (0)***
        mergeOptions.setMixAttributes(false) //do not merge attributes
        mergeOptions.setConsiderBorderAssignment(true)
        mergeOptions.setConsiderBorderCxnDefs(true)
        mergeOptions.setMixUserRelations(false) // Specifies whether relationships between users and user groups are to be merged.  
        mergeOptions.setConsiderSourceGroupPath(true) // If this option is set, then the merge algorithm will use the source group path in the case of a conflict.
        mergeOptions.setInheritAccessRightsForNewGroups(true) //true if access rights should be inherited when creating new groups
        //mergeOptions.setSynchronizeAssignments(true) //delete assignment on target side

        for (var i = 0; i < mergeModels.length; i++) {
            var mergeResult = merge.startMerge([mergeModels[i]], review_DB, mergeOptions, true)
            if (mergeResult.isSuccessful()) {
                g_table.push([mergeModels[i].Name(nLocale), mergeModels[i].Type(), mergeModels[i].Group().Path(nLocale), "Model published"]);

                var modelReadyForTestValue_Released = ArisData.ActiveFilter().UserDefinedAttributeValueTypeNum("3cfde370-e44c-11e9-004f-eabd1a8bb6a2", "3cfde372-e44c-11e9-004f-eabd1a8bb6a2")
                var modelReadyForTestWrite = mergeModels[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("3cfde370-e44c-11e9-004f-eabd1a8bb6a2"), nLocale).setValue(modelReadyForTestValue_Released)
                
                var modelTemp = review_DB.FindGUID(mergeModels[i].GUID())
                var modelReadyForTestWrite = modelTemp.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("3cfde370-e44c-11e9-004f-eabd1a8bb6a2"), nLocale).setValue(modelReadyForTestValue_Released)
            } else {
                g_table.push([mergeModels[i].Name(nLocale), mergeModels[i].Type(), mergeModels[i].Group().Path(nLocale), "Model is not published"]);
            }
        }

        review_DB.clearCaches()
        design_DB.clearCaches()
    }
    //*****************************************************

    design_DB.close()
    review_DB.close()

    //REORGANIZE DB
    reorganizeDb("Published")

    //PRINT LOG FILE
    printInformation("Log");
    g_exc.write();
}

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

function getModelsGUID(models) {
    var guidsArray = new Array()
    for (var i = 0; i < models.length; i++) {
        var guidStr = models[i].GUID()
        guidsArray.push(guidStr)
    }
    return guidsArray
}

function reorganizeDb(dbName) {
    var process = Context.execute("\"d:\\ARIS10\\tools\\ArisAdm\\arisadm10.exe\" -l reorg.log -s localhost -u system tutqL4LWQADAs33W reorg " + dbName);
    var inReader = new java.io.BufferedReader(new java.io.InputStreamReader(process.getInputStream()));
    // read all chars from the stream and concatenate them
    var sResult = "";
    var text;
    while ((text = inReader.readLine()) != null) {
        sResult += text;
    }
    process.waitFor(); //ends when process terminates
}