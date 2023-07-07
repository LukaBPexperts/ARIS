var nLocale = Context.getSelectedLanguage()
var selectedDatabases = ArisData.getSelectedDatabases()
var selectedDB = selectedDatabases[0]

//*****************************************
var giStandardColumnWidth = 15000;
var giRowCounter = 0;
var lDataRowStyle = null;

gExcelObject = Context.createExcelWorkbook(Context.getSelectedFile());

// set style header
var lHeaderFont = gExcelObject.createFont();
lHeaderFont.setBoldweight(Constants.XL_FONT_WEIGHT_BOLD);
var headLiBorderColor = Constants.C_GRAY;
var headerStyle = gExcelObject.createCellStyle(lHeaderFont, 1, 1, 1, 1, headLiBorderColor, headLiBorderColor, headLiBorderColor, headLiBorderColor, Constants.ALIGN_CENTER, Constants.VERTICAL_CENTER, Constants.C_TRANSPARENT, Constants.C_LIGHT_BLUE, Constants.SOLID_FOREGROUND);
// set style rows
var crFont = gExcelObject.createFont();
crFont.setBoldweight(Constants.XL_FONT_WEIGHT_NORMAL);
var crLiBorderColor = Constants.C_GRAY;
var crDataRowStyle = gExcelObject.createCellStyle(crFont, 1, 1, 1, 1, crLiBorderColor, crLiBorderColor, crLiBorderColor, crLiBorderColor, Constants.ALIGN_LEFT, Constants.VERTICAL_TOP, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);
crDataRowStyle.setWrapText(true);

printRowHeadersLog(gExcelObject);
//*****************************************

var excelFile = getExcelFile() //get excell file
var sheets = excelFile.getSheets() //get all sheets

//IMPORT MODEL DATA
var sheetModels = getSheetByName(sheets, "Models")
var mColumnNumGUID = getColumnNum(sheetModels, "MODEL GUID")

var rowCount = 0
if (sheetModels != null) {
    var rowCount = sheetModels.getRows().length
}

for (var i = 1; i < rowCount; i++) {

    var modelGUID = sheetModels.getCell(i, mColumnNumGUID).getCellValue().trim()
    var model = selectedDB.FindGUID(modelGUID, Constants.CID_MODEL)

    if (model.IsValid()) {
        //map API numbers
        var mNameApi = Constants.AT_NAME
        var mTargetApi = ArisData.ActiveFilter().UserDefinedAttributeTypeNum("6b6954e1-36bf-11ea-5df4-00505699050f")
        var mDescApi = Constants.AT_DESC
        var mShortDescApi = ArisData.ActiveFilter().UserDefinedAttributeTypeNum("c21714b1-734d-11ea-5df4-00505699050f")
        var mProcessFrequencyApi = ArisData.ActiveFilter().UserDefinedAttributeTypeNum("20a55bb1-36c0-11ea-5df4-00505699050f")
        var mPlanningHorizonApi = ArisData.ActiveFilter().UserDefinedAttributeTypeNum("759a2db1-41c6-11ea-5df4-00505699050f")
        var mProcessLevelApi = ArisData.ActiveFilter().UserDefinedAttributeTypeNum("a908e230-d8dc-11de-7aef-000c295dfb26")
        var mProcessOwnerApi = ArisData.ActiveFilter().UserDefinedAttributeTypeNum("e70ccb60-d01c-11dd-7315-000c295dfb26")
        var mProcessManagerApi = ArisData.ActiveFilter().UserDefinedAttributeTypeNum("0cf91db0-a399-11de-7aef-000c295dfb26")
        var mTitle1Api = Constants.AT_TITL1
        var mLink1Api = Constants.AT_EXT_1

        //get column num from excel
        var mNameNum = getColumnNum(sheetModels, "Name")
        var mTargetNum = getColumnNum(sheetModels, "Target")
        var mDescNum = getColumnNum(sheetModels, "Description")
        var mShortDescNum = getColumnNum(sheetModels, "Short Description")
        var mProcessFrequencyNum = getColumnNum(sheetModels, "Process Frequency")
        var mPlanningHorizonNum = getColumnNum(sheetModels, "Planning Horizon")
        var mProcessLevelNum = getColumnNum(sheetModels, "Process Level")
        var mProcessOwnerNum = getColumnNum(sheetModels, "Process Owner")
        var mProcessManagerNum = getColumnNum(sheetModels, "Process Manager")
        var mTitle1Num = getColumnNum(sheetModels, "Title 1")
        var mLink1Num = getColumnNum(sheetModels, "Link 1")

        //get values from excel
        var mNameValue = getDataFromExcel(sheetModels, i, mNameNum)
        var mTargetValue = getDataFromExcel(sheetModels, i, mTargetNum)
        var mDescValue = getDataFromExcel(sheetModels, i, mDescNum)
        var mShortDescValue = getDataFromExcel(sheetModels, i, mShortDescNum)
        var mProcessFrequencyValue = getDataFromExcel(sheetModels, i, mProcessFrequencyNum)
        var mPlanningHorizonValue = getDataFromExcel(sheetModels, i, mPlanningHorizonNum)
        var mProcessLevelValue = getDataFromExcel(sheetModels, i, mProcessLevelNum)
        var mProcessOwnerValue = getDataFromExcel(sheetModels, i, mProcessOwnerNum)
        var mProcessManagerValue = getDataFromExcel(sheetModels, i, mProcessManagerNum)
        var mTitle1Value = getDataFromExcel(sheetModels, i, mTitle1Num)
        var mLink1Value = getDataFromExcel(sheetModels, i, mLink1Num)

        //write values, and update log
        updateAttValue(model, mNameApi, mNameValue)
        updateAttValue(model, mTargetApi, mTargetValue)
        updateAttValue(model, mDescApi, mDescValue)
        updateAttValue(model, mShortDescApi, mShortDescValue)
        updateAttValue(model, mProcessFrequencyApi, mProcessFrequencyValue)
        updateAttValue(model, mPlanningHorizonApi, mPlanningHorizonValue)
        updateAttValue(model, mProcessLevelApi, mProcessLevelValue)
        updateAttValue(model, mProcessOwnerApi, mProcessOwnerValue)
        updateAttValue(model, mProcessManagerApi, mProcessManagerValue)
        updateAttValue(model, mTitle1Api, mTitle1Value)
        updateAttValue(model, mLink1Api, mLink1Value)

    }
    else {
        //write log, model is not found in the DB!
    }
}


//IMPORT OBJECT DATA
var sheetObjects = getSheetByName(sheets, "Objects")
var oColumnNumGUID = getColumnNum(sheetObjects, "OBJECT GUID")

var rowCount = 0
if (sheetObjects != null) {
    var rowCount = sheetObjects.getRows().length
}

for (var i = 1; i < rowCount; i++) {

    var objectGUID = sheetObjects.getCell(i, oColumnNumGUID).getCellValue().trim()
    var object = selectedDB.FindGUID(objectGUID, Constants.CID_OBJDEF)

    if (object.IsValid()) {
        //map API numbers   
        var oNameApi = Constants.AT_NAME
        var oDescApi = Constants.AT_DESC
        var oProcessLevelApi = ArisData.ActiveFilter().UserDefinedAttributeTypeNum("54398371-9079-11eb-5df4-00505699050f") //Process level (object)	     		
        var oTitle1Api = Constants.AT_TITL1
        var oLink1Api = Constants.AT_EXT_1

        //get column num from excel
        var oNameNum = getColumnNum(sheetObjects, "Name")
        var oDescNum = getColumnNum(sheetObjects, "Description")
        var oProcessLevelNum = getColumnNum(sheetObjects, "Process Level (object)")
        var oTitle1Num = getColumnNum(sheetObjects, "Title 1")
        var oLink1Num = getColumnNum(sheetObjects, "Link 1")

        //get values from excel
        var oNameValue = getDataFromExcel(sheetObjects, i, oNameNum)
        var oDescValue = getDataFromExcel(sheetObjects, i, oDescNum)
        var oProcessLevelValue = getDataFromExcel(sheetObjects, i, oProcessLevelNum)
        var oTitle1Value = getDataFromExcel(sheetObjects, i, oTitle1Num)
        var oLink1Value = getDataFromExcel(sheetObjects, i, oLink1Num)

        //write values, and update log
        updateAttValue(object, oNameApi, oNameValue)
        updateAttValue(object, oDescApi, oDescValue)
        updateAttValue(object, oProcessLevelApi, oProcessLevelValue)
        updateAttValue(object, oTitle1Api, oTitle1Value)
        updateAttValue(object, oLink1Api, oLink1Value)

    }
    else {
        //write log, object is not found in the DB!
    }
}

//write file
gExcelObject.write();

function getDataFromExcel(sheet, rowNum, columnNum) {
    try {
        var value = sheet.getCell(rowNum, columnNum).getCellValue()
    }
    catch (e) {
        //could not read, write log
        return ""
    }
    return value;
}

function getExcelFile() {
    var sdefname = ""
    sdefext = __toString("*.xls!!Excel|*.xls; *.xlsx||")
    var sdefdir = ""
    var stitle = "Chose excel file"

    var files = Dialogs.BrowseForFiles(sdefname, sdefext, sdefdir, stitle, 0)

    var excelFile = Context.getExcelReader(files[0].getData())

    return excelFile;
}

function getSheetByName(sheets, sheetName) {
    var returnSheet = null
    for (var i = 0; i < sheets.length; i++) {
        if (sheets[i].getName() == sheetName) {
            returnSheet = sheets[i]
        }
    }
    return returnSheet
}

function getColumnNum(sheet, columnName) {
    var temp = null;
    for (var i = 0; i < 50; i++) {
        if (sheet.getCell(0, i).getCellValue().trim().toLowerCase() == columnName.trim().toLowerCase()) {
            var temp = i;
            break;
        }
    }
    return temp;
}

function updateAttValue(objDef, attApi, attValue) {
    if (objDef.Attribute(attApi, nLocale).getValue() != attValue) {
        var oldValue = objDef.Attribute(attApi, nLocale).getValue()
        var updateAtt = objDef.Attribute(attApi, nLocale).setValue(attValue)
        if (updateAtt) {
            //printRowSheet2(objDef.Name(nLocale), objDef.Attribute(apmNumberAttNum, nLocale).getValue(), objDef.Attribute(attApi, nLocale).Type() + " - Attribute Updated", oldValue, attValue, gExcelObject)
            printRowLog(objDef.Type(), objDef.Name(nLocale), objDef.Attribute(attApi, nLocale).Type(), oldValue, attValue, "Attribute updated", gExcelObject)
        }
        else {
            printRowLog(objDef.Type(), objDef.Name(nLocale), objDef.Attribute(attApi, nLocale).Type(), oldValue, attValue, "ERROR: Attribute not updated", gExcelObject)
        }
    }
}

function printRowHeadersLog(pExcelWorkBook) {
    //var sDate = java.text.SimpleDateFormat("dd-MM-yyyy HH.mm").format(new java.util.Date());
    var sSheetName = "LOG"
    var lMainSheet = pExcelWorkBook.createSheet(sSheetName);
    var lHeaderRow = lMainSheet.createRow(giRowCounter);

    lMainSheet.setDefaultColumnWidth(30);

    /* Freeze top row alone */
    lMainSheet.createFreezePane(0, 1);

    // set style
    //var lHeaderFont = pExcelWorkBook.createFont();
    //lHeaderFont.setBoldweight(Constants.XL_FONT_WEIGHT_BOLD);
    //var headLiBorderColor = Constants.C_GRAY;

    //headerStyle = pExcelWorkBook.createCellStyle(lHeaderFont, 1, 1, 1, 1, headLiBorderColor, headLiBorderColor, headLiBorderColor, headLiBorderColor, Constants.ALIGN_CENTER, Constants.VERTICAL_CENTER, Constants.C_TRANSPARENT, Constants.C_LIGHT_BLUE, Constants.SOLID_FOREGROUND);

    var lCellL1 = lHeaderRow.createCell(0);
    lCellL1.setCellValue("Type");
    lCellL1.setCellStyle(headerStyle);

    var lCellL2 = lHeaderRow.createCell(1);
    lCellL2.setCellValue("Name");
    lCellL2.setCellStyle(headerStyle);

    var lCellL3 = lHeaderRow.createCell(2);
    lCellL3.setCellValue("Attribute");
    lCellL3.setCellStyle(headerStyle);

    var lCellL4 = lHeaderRow.createCell(3);
    lCellL4.setCellValue("Old Value");
    lCellL4.setCellStyle(headerStyle);

    var lCellL5 = lHeaderRow.createCell(4);
    lCellL5.setCellValue("New Value");
    lCellL5.setCellStyle(headerStyle);

    var lCellL6 = lHeaderRow.createCell(5);
    lCellL6.setCellValue("Status");
    lCellL6.setCellStyle(headerStyle);

    giRowCounter++
}

function printRowLog(l_type, l_name, l_attribute, l_oldValue, l_newValue, l_status, pExcelWorkBook) {
    try {

        //var crFont = pExcelWorkBook.createFont();
        //crFont.setBoldweight(Constants.XL_FONT_WEIGHT_NORMAL);
        //var crLiBorderColor = Constants.C_GRAY;

        //crDataRowStyle = pExcelWorkBook.createCellStyle(crFont, 1, 1, 1, 1, crLiBorderColor, crLiBorderColor, crLiBorderColor, crLiBorderColor, Constants.ALIGN_LEFT, Constants.VERTICAL_TOP, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.SOLID_FOREGROUND);

        //crDataRowStyle.setWrapText(true);

        var lMainSheet = pExcelWorkBook.getSheetAt(0);

        // second row data
        var lModelRow = lMainSheet.createRow(giRowCounter);

        var lCell = lModelRow.createCell(0);
        lCell.setCellValue(l_type);
        lCell.setCellStyle(crDataRowStyle);

        var lCell = lModelRow.createCell(1);
        lCell.setCellValue(l_name);
        lCell.setCellStyle(crDataRowStyle);

        var lCell = lModelRow.createCell(2);
        lCell.setCellValue(l_attribute);
        lCell.setCellStyle(crDataRowStyle);

        // Input Name
        var lCell = lModelRow.createCell(3);
        lCell.setCellValue(l_oldValue);
        lCell.setCellStyle(crDataRowStyle);

        var lCell = lModelRow.createCell(4);
        lCell.setCellValue(l_newValue);
        lCell.setCellStyle(crDataRowStyle);

        var lCell = lModelRow.createCell(5);
        lCell.setCellValue(l_status);
        lCell.setCellStyle(crDataRowStyle);

        // up counter
        giRowCounter++;
    }
    catch (ex) {
        Dialogs.MsgBox("Error : " + ex);
    }
}