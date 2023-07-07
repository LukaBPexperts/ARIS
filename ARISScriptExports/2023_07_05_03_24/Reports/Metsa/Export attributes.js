// static variables
var nLocale = Context.getSelectedLanguage();
gExcelObject = Context.createExcelWorkbook(Context.getSelectedFile());
var giStandardColumnWidth = 8000;
var giRowCounterModels = 0;
var giRowCounterObjects = 0;

// set style
var crFontHeader = gExcelObject.createFont();
crFontHeader.setBoldweight(Constants.XL_FONT_WEIGHT_NORMAL);
crFontHeader.setBold(true)
var crLiBorderColor = Constants.C_GRAY;
crHeaderRowStyle = gExcelObject.createCellStyle(crFontHeader, 1, 1, 1, 1, crLiBorderColor, crLiBorderColor, crLiBorderColor, crLiBorderColor, Constants.ALIGN_LEFT, Constants.VERTICAL_TOP, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);
crHeaderRowStyle.setWrapText(true);

var crFont = gExcelObject.createFont();
crFont.setBoldweight(Constants.XL_FONT_WEIGHT_NORMAL);
var crLiBorderColor = Constants.C_GRAY;
crDataRowStyle = gExcelObject.createCellStyle(crFont, 1, 1, 1, 1, crLiBorderColor, crLiBorderColor, crLiBorderColor, crLiBorderColor, Constants.ALIGN_LEFT, Constants.VERTICAL_TOP, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);
crDataRowStyle.setWrapText(true);

var crFontGUID = gExcelObject.createFont();
crFontGUID.setBoldweight(Constants.XL_FONT_WEIGHT_BOLD);
crFontGUID.setBold(true)
crDataRowStyleGUID = gExcelObject.createCellStyle(crFontGUID, 1, 1, 1, 1, crLiBorderColor, crLiBorderColor, crLiBorderColor, crLiBorderColor, Constants.ALIGN_LEFT, Constants.VERTICAL_TOP, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.SOLID_FOREGROUND);
crDataRowStyleGUID.setWrapText(true);

printHeadersModels(gExcelObject);
printHeadersObjects(gExcelObject);

main();

function main() {
    var selectedGroups = ArisData.getSelectedGroups()
    var selectedModels = ArisData.getSelectedModels()

    if (selectedGroups.length > 0) {
        for (var i = 0; i < selectedGroups.length; i++) {
            var selectedGroup = selectedGroups[i]
            var models = selectedGroup.ModelList()
            collectAndPrintData(models, gExcelObject);
        }
    }
    else if (selectedModels.length > 0) {
        for (var i = 0; i < selectedModels.length; i++) {
            collectAndPrintData(selectedModels, gExcelObject);
        }
    }
    // write report
    gExcelObject.write();
}

function printHeadersModels(pExcelWorkBook) {
    var sSheetName = "Models";
    var lMainSheet = pExcelWorkBook.createSheet(sSheetName);
    var lHeaderRow = lMainSheet.createRow(giRowCounterModels);

    // set column width    
    lMainSheet.setColumnWidth(0, giStandardColumnWidth);
    lMainSheet.setColumnWidth(1, giStandardColumnWidth);
    lMainSheet.setColumnWidth(2, giStandardColumnWidth);
    lMainSheet.setColumnWidth(3, giStandardColumnWidth);
    lMainSheet.setColumnWidth(4, giStandardColumnWidth);
    lMainSheet.setColumnWidth(5, giStandardColumnWidth);
    lMainSheet.setColumnWidth(6, giStandardColumnWidth);
    lMainSheet.setColumnWidth(7, giStandardColumnWidth);
    lMainSheet.setColumnWidth(8, giStandardColumnWidth);
    lMainSheet.setColumnWidth(9, giStandardColumnWidth);
    lMainSheet.setColumnWidth(10, giStandardColumnWidth);
    lMainSheet.setColumnWidth(11, 9000);


    /* Freeze top row alone */
    lMainSheet.createFreezePane(0, 1);

    var lCellL1 = lHeaderRow.createCell(0);
    lCellL1.setCellValue("Name");
    lCellL1.setCellStyle(crHeaderRowStyle);

    var lCellL2 = lHeaderRow.createCell(1);
    lCellL2.setCellValue("Target");
    lCellL2.setCellStyle(crHeaderRowStyle);

    var lCellL3 = lHeaderRow.createCell(2);
    lCellL3.setCellValue("Description");
    lCellL3.setCellStyle(crHeaderRowStyle);

    var lCellL4 = lHeaderRow.createCell(3);
    lCellL4.setCellValue("Short Description");
    lCellL4.setCellStyle(crHeaderRowStyle);

    var lCell5 = lHeaderRow.createCell(4);
    lCell5.setCellValue("Process Frequency");
    lCell5.setCellStyle(crHeaderRowStyle);

    var lCell6 = lHeaderRow.createCell(5);
    lCell6.setCellValue("Planning Horizon");
    lCell6.setCellStyle(crHeaderRowStyle);

    var lCell7 = lHeaderRow.createCell(6);
    lCell7.setCellValue("Process Level");
    lCell7.setCellStyle(crHeaderRowStyle);

    var lCell8 = lHeaderRow.createCell(7);
    lCell8.setCellValue("Process Owner");
    lCell8.setCellStyle(crHeaderRowStyle);

    var lCell9 = lHeaderRow.createCell(8);
    lCell9.setCellValue("Process Manager");
    lCell9.setCellStyle(crHeaderRowStyle);

    var lCell10 = lHeaderRow.createCell(9);
    lCell10.setCellValue("Title 1");
    lCell10.setCellStyle(crHeaderRowStyle);

    var lCell11 = lHeaderRow.createCell(10);
    lCell11.setCellValue("Link 1");
    lCell11.setCellStyle(crHeaderRowStyle);

    var lCell12 = lHeaderRow.createCell(11);
    lCell12.setCellValue("MODEL GUID");
    lCell12.setCellStyle(crDataRowStyleGUID);

    giRowCounterModels++
}

function printRowModels(psName, psTarget, psDescription, psShortDescription, psProcessFrequency, psPlanningHorizon, psProcessLevel, psProcessOwner, psProcessManager, psTitle1, psLink1, psMODELGUID, pExcelWorkBook) {
    var sheets = pExcelWorkBook.getSheets()
    for (var i = 0; i < sheets.length; i++) {
        if (pExcelWorkBook.getSheetName(i) == "Models") {
            var lMainSheet = sheets[i]
        }
    }

    // second row data
    var lModelRow = lMainSheet.createRow(giRowCounterModels);

    var lCell = lModelRow.createCell(0);
    lCell.setCellValue(psName);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(1);
    lCell.setCellValue(psTarget);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(2);
    lCell.setCellValue(psDescription);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(3);
    lCell.setCellValue(psShortDescription);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(4);
    lCell.setCellValue(psProcessFrequency);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(5);
    lCell.setCellValue(psPlanningHorizon);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(6);
    lCell.setCellValue(psProcessLevel);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(7);
    lCell.setCellValue(psProcessOwner);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(8);
    lCell.setCellValue(psProcessManager);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(9);
    lCell.setCellValue(psTitle1);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(10);
    lCell.setCellValue(psLink1);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(11);
    lCell.setCellValue(psMODELGUID);
    lCell.setCellStyle(crDataRowStyle);

    // up counter
    giRowCounterModels++;
}

function printHeadersObjects(pExcelWorkBook) {
    var sSheetName = "Objects";
    var lMainSheet = pExcelWorkBook.createSheet(sSheetName);
    var lHeaderRow = lMainSheet.createRow(giRowCounterObjects);

    // set column width    
    lMainSheet.setColumnWidth(0, giStandardColumnWidth);
    lMainSheet.setColumnWidth(1, giStandardColumnWidth);
    lMainSheet.setColumnWidth(2, giStandardColumnWidth);
    lMainSheet.setColumnWidth(3, giStandardColumnWidth);
    lMainSheet.setColumnWidth(4, giStandardColumnWidth);
    lMainSheet.setColumnWidth(5, giStandardColumnWidth);
    lMainSheet.setColumnWidth(6, giStandardColumnWidth);
    lMainSheet.setColumnWidth(7, 9000);

    /* Freeze top row alone */
    lMainSheet.createFreezePane(0, 1);

    var lCellL1 = lHeaderRow.createCell(0);
    lCellL1.setCellValue("Model Name");
    lCellL1.setCellStyle(crHeaderRowStyle);

    var lCellL2 = lHeaderRow.createCell(1);
    lCellL2.setCellValue("Object Type");
    lCellL2.setCellStyle(crHeaderRowStyle);

    var lCellL3 = lHeaderRow.createCell(2);
    lCellL3.setCellValue("Name");
    lCellL3.setCellStyle(crHeaderRowStyle);

    var lCellL4 = lHeaderRow.createCell(3);
    lCellL4.setCellValue("Description");
    lCellL4.setCellStyle(crHeaderRowStyle);

    var lCell5 = lHeaderRow.createCell(4);
    lCell5.setCellValue("Process Level (object)");
    lCell5.setCellStyle(crHeaderRowStyle);

    var lCell6 = lHeaderRow.createCell(5);
    lCell6.setCellValue("Title 1");
    lCell6.setCellStyle(crHeaderRowStyle);

    var lCell7 = lHeaderRow.createCell(6);
    lCell7.setCellValue("Link 1");
    lCell7.setCellStyle(crHeaderRowStyle);

    var lCell8 = lHeaderRow.createCell(7);
    lCell8.setCellValue("OBJECT GUID");
    lCell8.setCellStyle(crDataRowStyleGUID);

    giRowCounterObjects++
}

function printRowObjects(psModelName, psObjectType, psName, psDescription, psProcessLevel, psTitle1, psLink1, psOBJECTGUID, pExcelWorkBook) {
    var sheets = pExcelWorkBook.getSheets()
    for (var i = 0; i < sheets.length; i++) {
        if (pExcelWorkBook.getSheetName(i) == "Objects") {
            var lMainSheet = sheets[i]
        }
    }

    var lModelRow = lMainSheet.createRow(giRowCounterObjects);

    var lCell = lModelRow.createCell(0);
    lCell.setCellValue(psModelName);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(1);
    lCell.setCellValue(psObjectType);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(2);
    lCell.setCellValue(psName);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(3);
    lCell.setCellValue(psDescription);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(4);
    lCell.setCellValue(psProcessLevel);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(5);
    lCell.setCellValue(psTitle1);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(6);
    lCell.setCellValue(psLink1);
    lCell.setCellStyle(crDataRowStyle);

    var lCell = lModelRow.createCell(7);
    lCell.setCellValue(psOBJECTGUID);
    lCell.setCellStyle(crDataRowStyle);

    // up counter
    giRowCounterObjects++;
}

function collectAndPrintData(pModels, pExcelWorkBook) {
    for (var i = 0; i < pModels.length; i++) {
        var model = pModels[i];
        
        var mName = model.Name(nLocale);
        var mTarget = model.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("6b6954e1-36bf-11ea-5df4-00505699050f"), nLocale).getValue();
        var mDescription = model.Attribute(Constants.AT_DESC, nLocale).getValue();
        var mShortDescription = model.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("c21714b1-734d-11ea-5df4-00505699050f"), nLocale).getValue();
        var mProcessFrequency = model.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("20a55bb1-36c0-11ea-5df4-00505699050f"), nLocale).getValue();
        var mPlanningHorizon = model.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("759a2db1-41c6-11ea-5df4-00505699050f"), nLocale).getValue();
        var mProcessLevel = model.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("a908e230-d8dc-11de-7aef-000c295dfb26"), nLocale).getValue();
        var mProcessOwner = model.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("e70ccb60-d01c-11dd-7315-000c295dfb26"), nLocale).getValue();
        var mProcessManager = model.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("0cf91db0-a399-11de-7aef-000c295dfb26"), nLocale).getValue();
        var mTitle1 = model.Attribute(Constants.AT_TITL1, nLocale).getValue();
        var mLink1 = model.Attribute(Constants.AT_EXT_1, nLocale).getValue();
        var mMODELGUID = model.GUID()

        printRowModels(mName, mTarget, mDescription, mShortDescription, mProcessFrequency, mPlanningHorizon, mProcessLevel, mProcessOwner, mProcessManager, mTitle1, mLink1, mMODELGUID, pExcelWorkBook)

        //get all func and evt objects, print data
        var objDefs = model.ObjDefListByTypes([Constants.OT_FUNC, Constants.OT_EVT]);
        for (var j = 0; j < objDefs.length; j++) {
            var objDef = objDefs[j];

            var oObjectType = objDef.Type()
            var oName = objDef.Name(nLocale)
            var oDescription = objDef.Attribute(Constants.AT_DESC, nLocale).getValue();
            var oProcessLevel = objDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("54398371-9079-11eb-5df4-00505699050f"), nLocale).getValue();
            var oTitle1 = objDef.Attribute(Constants.AT_TITL1, nLocale).getValue();
            var oLink1 = objDef.Attribute(Constants.AT_EXT_1, nLocale).getValue();
            var oOBJECTGUID = objDef.GUID()

            printRowObjects(mName, oObjectType, oName, oDescription, oProcessLevel, oTitle1, oLink1, oOBJECTGUID, pExcelWorkBook)

        }
    }
}