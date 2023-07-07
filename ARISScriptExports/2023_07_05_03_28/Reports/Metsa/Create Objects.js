var nLocale = Context.getSelectedLanguage()

var selectedModels = ArisData.getSelectedModels()
var selectedModel = selectedModels[0]
var modelGroup = selectedModel.Group()

var excelFile = getExcelFile() //get excell file

//MAP ATTRIBUTES
var oDescApi = Constants.AT_DESC
var oTitle1Api = Constants.AT_TITLE_1
//var oLink1Api = Constants.AT_LINK

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

//var sheets = excelFile.getSheets() //get all sheets
//var sheet_0 = getSheetByName(sheets, "Models")
var sheet_0 = excelFile.getSheetAt(0)
//var mColumnNumGUID = getColumnNum(sheet_0, "MODEL GUID")

var rowCount = 0
if (sheet_0 != null) {
    var rowCount = sheet_0.getRows().length
}

var c_Y = getMaxY(selectedModel) + 300
var c_X = 100
for (var i = 1; i < rowCount; i++) {

    //get column num from excel
    var oNameNum = getColumnNum(sheet_0, "Name")
    var oTypeNum = getColumnNum(sheet_0, "Object type")
    var oDescNum = getColumnNum(sheet_0, "Description")
    var oTitle1Num = getColumnNum(sheet_0, "Title 1")
    var oLink1Num = getColumnNum(sheet_0, "Link 1")
    var oProcessLevelNum = getColumnNum(sheet_0, "Process level (object)")
    var oSynonymsNum = getColumnNum(sheet_0, "Synonyms")

    //get values from excel
    var oNameValue = getDataFromExcel(sheet_0, i, oNameNum)
    var oTypeValue = getDataFromExcel(sheet_0, i, oTypeNum)
    var oDescValue = getDataFromExcel(sheet_0, i, oDescNum)
    var oTitle1Value = getDataFromExcel(sheet_0, i, oTitle1Num)
    var oLink1Value = getDataFromExcel(sheet_0, i, oLink1Num)
    var oProcessLevelValue = getDataFromExcel(sheet_0, i, oProcessLevelNum)
    var oSynonymsValue = getDataFromExcel(sheet_0, i, oSynonymsNum)

    //create object
    var objDefType = getObjDefType(oTypeValue)
    if (objDefType != null) {
        var symbolNum = getSymbol(oTypeValue)

        //create objDef
        var sourceObjDef = modelGroup.CreateObjDef(objDefType, oNameValue, Context.getSelectedLanguage())
        if (sourceObjDef.IsValid()) {
            printRowLog(sourceObjDef.Type(), sourceObjDef.Name(nLocale), sourceObjDef.Group().Path(nLocale), "Object definition created", gExcelObject)
            //create objOcc on model
            var sourceObjOcc = selectedModel.createObjOcc(symbolNum, sourceObjDef, c_X, c_Y, true, false)
            if (c_X < 600) {
                c_X = c_X + 500
            }
            else {
                c_X = 100
                c_Y = c_Y + 300
            }

            if (sourceObjOcc.IsValid()) {
                printRowLog(sourceObjDef.Type(), sourceObjDef.Name(nLocale), sourceObjDef.Group().Path(nLocale), "Object occurence created", gExcelObject)

                //write values, and update log
                //updateAttValue(sourceObjDef, oNameApi, oNameValue)
                updateAttValue(sourceObjDef, oDescApi, oDescValue)
                //updateAttValue(sourceObjDef, oProcessLevelApi, oProcessLevelValue)

            }
            else {
                printRowLog(sourceObjDef.Type(), sourceObjDef.Name(nLocale), sourceObjDef.Group().Path(nLocale), "ERROR: Object occurence is not created", gExcelObject)
            }
        }
        else {
            //objdef error
            printRowLog(oTypeValue, oNameValue, "", "ERROR: Object definition is not created", gExcelObject)
        }
    }
}

//write file
gExcelObject.write();

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
    for (var i = 0; i < 50; i++) { //excel file with max 50 columns
        if (sheet.getCell(0, i).getCellValue().trim().toLowerCase() == columnName.trim().toLowerCase()) {
            var temp = i;
            break;
        }
    }
    return temp;
}

function getDataFromExcel(sheet, rowNum, columnNum) {
    try {
        var value = sheet.getCell(rowNum, columnNum).getStringCellValue()
    }
    catch (e) {
        //could not read, write log
        return ""
    }
    return value;
}


function updateAttValue(objDef, attApi, attValue) {
    if (objDef.Attribute(attApi, nLocale).getValue() != attValue) {
        var oldValue = objDef.Attribute(attApi, nLocale).getValue()
        var updateAtt = objDef.Attribute(attApi, nLocale).setValue(attValue)
        if (updateAtt) {
            //printRowSheet2(objDef.Name(nLocale), objDef.Attribute(apmNumberAttNum, nLocale).getValue(), objDef.Attribute(attApi, nLocale).Type() + " - Attribute Updated", oldValue, attValue, gExcelObject)
            printRowLog(objDef.Type(), objDef.Name(nLocale), objDef.Group().Path(nLocale), "Attribute: " + objDef.Attribute(attApi, nLocale).Type() + " is updated", gExcelObject)
        }
        else {
            printRowLog(objDef.Type(), objDef.Name(nLocale), objDef.Group().Path(nLocale), "ERROR: Attribute: " + objDef.Attribute(attApi, nLocale).Type() + " is not updated", gExcelObject)
        }
    }
}

function getMaxY(model) {
    var tempY = 0;
    var objOccs = model.ObjOccList()

    for (var i = 0; i < objOccs.length; i++) {
        if (objOccs[i].Y() > tempY) {
            tempY = objOccs[i].Y()
        }
    }
    return tempY
}

function getObjDefType(oTypeValue) {
    if (oTypeValue == "Function") {
        return Constants.OT_FUNC
    }

    if (oTypeValue == "Event") {
        return Constants.OT_EVT
    }

    if (oTypeValue == "Role") {
        return Constants.OT_PERS_TYPE
    }

    if (oTypeValue == "Application system type") {
        return Constants.OT_APPL_SYS_TYPE
    }

    if (oTypeValue == "Information carrier") {
        return Constants.OT_INFO_CARR
    }

    if (oTypeValue == "Data") {
        return Constants.OT_ENT_TYPE
    }

    if (oTypeValue == "Business rule") {
        return Constants.OT_BUSINESS_RULE
    }

    return null;
}

function getSymbol(oTypeValue) {
    if (oTypeValue == "Function") {
        return Constants.ST_FUNC
    }

    if (oTypeValue == "Event") {
        return Constants.ST_EV
    }

    if (oTypeValue == "Role") {
        return Constants.ST_EMPL_TYPE
    }

    if (oTypeValue == "Application system type") {
        return Constants.ST_APPL_SYS_TYPE
    }

    if (oTypeValue == "Information carrier") {
        return Constants.ST_INFO_CARR
    }

    if (oTypeValue == "Data") {
        return Constants.ST_ENT_TYPE
    }

    if (oTypeValue == "Business rule") {
        return Constants.ST_BUSINESS_RULE
    }

    return null;
}

function printRowHeadersLog(pExcelWorkBook) {
    //var sDate = java.text.SimpleDateFormat("dd-MM-yyyy HH.mm").format(new java.util.Date());
    var sSheetName = "LOG"
    var lMainSheet = pExcelWorkBook.createSheet(sSheetName);
    var lHeaderRow = lMainSheet.createRow(giRowCounter);

    lMainSheet.setDefaultColumnWidth(30);

    /* Freeze top row alone */
    lMainSheet.createFreezePane(0, 1);

    var lCellL1 = lHeaderRow.createCell(0);
    lCellL1.setCellValue("Type");
    lCellL1.setCellStyle(headerStyle);

    var lCellL2 = lHeaderRow.createCell(1);
    lCellL2.setCellValue("Name");
    lCellL2.setCellStyle(headerStyle);

    var lCellL3 = lHeaderRow.createCell(2);
    lCellL3.setCellValue("Location");
    lCellL3.setCellStyle(headerStyle);

    var lCellL4 = lHeaderRow.createCell(3);
    lCellL4.setCellValue("Status");
    lCellL4.setCellStyle(headerStyle);

    giRowCounter++
}

function printRowLog(l_type, l_name, l_location, l_status, pExcelWorkBook) {
    try {

        var lMainSheet = pExcelWorkBook.getSheetAt(0);

        var lModelRow = lMainSheet.createRow(giRowCounter);

        var lCell = lModelRow.createCell(0);
        lCell.setCellValue(l_type);
        lCell.setCellStyle(crDataRowStyle);

        var lCell = lModelRow.createCell(1);
        lCell.setCellValue(l_name);
        lCell.setCellStyle(crDataRowStyle);

        var lCell = lModelRow.createCell(2);
        lCell.setCellValue(l_location);
        lCell.setCellStyle(crDataRowStyle);

        var lCell = lModelRow.createCell(3);
        lCell.setCellValue(l_status);
        lCell.setCellStyle(crDataRowStyle);

        // up counter
        giRowCounter++;
    }
    catch (ex) {
        Dialogs.MsgBox("Error : " + ex);
    }
}