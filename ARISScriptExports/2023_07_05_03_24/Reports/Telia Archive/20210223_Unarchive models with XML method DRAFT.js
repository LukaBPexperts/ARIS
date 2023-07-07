//****************CUSTOMIZE**********************************
var prodDbName = "_TeliaCompany_WORKING"
var admUser = "arisservice"
var admPass = "arisservice"
//***********************************************************

var nLocale = Context.getSelectedLanguage()

//log file global variables
var g_exc = Context.createExcelWorkbook(Context.getSelectedFile());
var g_table = new Array();
var g_headers = new Array("Model", "Model Type", "Location", "Person Responsible", "Model Unarchived", "Log");

//var selectedDatabase = ArisData.getActiveDatabase()
var selectedDatabases = ArisData.getSelectedDatabases()
var selectedModels = ArisData.getSelectedModels()

if (selectedDatabases.length != 0) {
    //find models marked as ready for archive
    var selectedDatabase = selectedDatabases[0]
    var searchItemSpecial = selectedDatabase.createSearchItem(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("24557581-cc28-11ea-4a36-00505682171a"), nLocale, "To be Unarchived (for Admins)", Constants.SEARCH_CMP_EQUAL, false, false)
    var selectedModels = selectedDatabase.Find(Constants.SEARCH_MODEL, null, searchItemSpecial)
    var modelsForUnarchive = selectedModels
} else {
    var selectedDatabase = ArisData.getActiveDatabase()
    //var modelsForUnarchive = checkModels(selectedModels)
    var modelsForUnarchive = selectedModels
}

//*****************************************************
var btnPressed = Dialogs.MsgBox(modelsForUnarchive.length + " model will be unarchived. Press NO if you want to simulate unarchival, without making any changes", Constants.MSGBOX_ICON_QUESTION | Constants.MSGBOX_BTN_YESNO + 512, "Message from Report")
if (btnPressed == Constants.MSGBOX_RESULT_YES) {
    selectedDatabase.setAutoTouch(false);

    //var prod_DB = ArisData.openDatabase(prodDbName, admUser, admPass, "dd838074-ac29-11d4-85b8-00005a4053ff", nLocale, false)
    var prod_DB = ArisData.openDatabase(prodDbName)

    //XML EXPORT/IMPORT   
    var xmlExportComponent = Context.getComponent("XMLExport")
    var xmlImportComponent = Context.getComponent("XMLImport")
    var selection = Dialogs.BrowseArisItems("Select group", "Select the target group", "(servername is obsolete)", prodDbName, Constants.CID_GROUP)
    var aOIDs = selection.split(";")
    var targetGroup = prod_DB.FindOID(aOIDs[0])


    for (var i = 0; i < modelsForUnarchive.length; i++) {
        //change archive status to Active
        var attValue_Restored = ArisData.ActiveFilter().UserDefinedAttributeValueTypeNum("24557581-cc28-11ea-4a36-00505682171a", "2468d677-cc28-11ea-4a36-00505682171a")
        var archivingStatusAtt = modelsForUnarchive[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("24557581-cc28-11ea-4a36-00505682171a"), nLocale).setValue(attValue_Restored)

        // var modelStatus = modelsForUnarchive[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("d2ac7740-af65-11df-599b-005056aa0ab6"), nLocale).Delete()

        //XML EXPORT
        var xmlExportOptions = xmlExportComponent.createDefaultXMLExportOptions(modelsForUnarchive[i], [nLocale]) //add array of models, only one model for test    
        xmlExportOptions.setExportCompleteGroupPaths(false)
        xmlExportOptions.setExportFormatedAttributeDefinitions(true)
        xmlExportOptions.setExportUsersAndUsergroups(false)
        xmlExportOptions.setHierarchyDepth(1) //assignments depth
        xmlExportOptions.setRelationDepth(1) //connections depth
        xmlExportOptions.setIsSimpleExport(false)
        xmlExportOptions.setZipExport(false)

        var xmlExport = xmlExportComponent.startExport(xmlExportOptions)
        var xmlFile = xmlExport.getFileResult()

        //XML IMPORT 
        var modelOptions = xmlImportComponent.createDefaultModelOptions()
        var xmlImportOptionsAml = xmlImportComponent.createDefaultAMLOptions(xmlFile, targetGroup, nLocale, modelOptions)

        var xmlImportOptionsMerge = xmlImportComponent.createDefaultMergeOptions()
        xmlImportOptionsMerge.setConsiderBorderAssignments(true);
        xmlImportOptionsMerge.setConsiderBorderCxnDefs(true);
        xmlImportOptionsMerge.setConsiderGroupPath(false);
        xmlImportOptionsMerge.setCxnDefSourceWins(true);
        xmlImportOptionsMerge.setFontsSourceWins(true);
        xmlImportOptionsMerge.setModelSourceWins(true);
        xmlImportOptionsMerge.setObjDefSourceWins(true);

        var xmlImportResult = xmlImportComponent.startImport(xmlImportOptionsAml, xmlImportOptionsMerge)

        if (xmlImportResult.isSuccessful()) {
            var groupTemp = modelsForUnarchive[i].Group()
            g_table.push([modelsForUnarchive[i].Name(nLocale), modelsForUnarchive[i].Type(), modelsForUnarchive[i].Group().Path(nLocale), modelsForUnarchive[i].Attribute(Constants.AT_PERS_RESP, nLocale).getValue(), xmlImportResult.isSuccessful(), "Model Unarchived"]);
            var deleteModel = groupTemp.Delete(modelsForUnarchive[i])
        } else {
            g_table.push([modelsForUnarchive[i].Name(nLocale), modelsForUnarchive[i].Type(), modelsForUnarchive[i].Group().Path(nLocale), modelsForUnarchive[i].Attribute(Constants.AT_PERS_RESP, nLocale).getValue(), xmlImportResult.isSuccessful(), "Merge Error"]);
        }
    }

    prod_DB.clearCaches()
    selectedDatabase.setAutoTouch(true);

    //PRINT LOG FILE
    printInformation("Log");
    g_exc.write();
    
} else {
    for (var i = 0; i < modelsForUnarchive.length; i++) {
        g_table.push([modelsForUnarchive[i].Name(nLocale), modelsForUnarchive[i].Type(), modelsForUnarchive[i].Group().Path(nLocale), modelsForUnarchive[i].Attribute(Constants.AT_PERS_RESP, nLocale).getValue(), "TRUE", "Model Ready for Unarchival"]);
    }

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

function checkModels(models) {
    var modelsForArchive = new Array();

    for (var i = 0; i < models.length; i++) {
        var errorFlag = true
        var errorMessage = ""

        var modelArchiveStatus = models[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("24557581-cc28-11ea-4a36-00505682171a"), nLocale).getValue()

        if (modelArchiveStatus != "To be Unarchived (for Admins)") {
            errorFlag = false
            errorMessage = errorMessage + "archive status is not set to To be Unarchived (for Admins)"
        }

        if (errorFlag != false) {
            modelsForArchive.push(models[i])
            //g_table.push([models[i].Name(nLocale), models[i].Type(), models[i].Group().Path(nLocale), models[i].Attribute (Constants.AT_PERS_RESP, nLocale).getValue(), errorFlag, errorMessage]);
        } else {
            g_table.push([models[i].Name(nLocale), models[i].Type(), models[i].Group().Path(nLocale), models[i].Attribute(Constants.AT_PERS_RESP, nLocale).getValue(), errorFlag, errorMessage]);
        }
    }
    return modelsForArchive
}