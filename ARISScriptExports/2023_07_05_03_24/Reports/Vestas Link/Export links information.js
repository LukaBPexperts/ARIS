var nLocale = Context.getSelectedLanguage()
//var oOutput = Context.createOutputObject()
var excelObject = Context.createExcelWorkbook(Context.getSelectedFile());

var selectedDatabases = ArisData.getSelectedDatabases()

var selectedDatabase = selectedDatabases[0]

//setup excell file
//*************************************************************
var objectsName = "Objects"
var objectsSheet = excelObject.createSheet(objectsName);
//mainSheet.setColumnWidth(0, 12000);
objectsSheet.setDefaultColumnWidth(30)

var modelsName = "Models"
var modelsSheet = excelObject.createSheet(modelsName);
modelsSheet.setDefaultColumnWidth(30)

createHeaderRow(objectsSheet)
createHeaderRow(modelsSheet)

//*************************************************************

//OBJECTS
//var objDefs = selectedDatabase.Find(Constants.SEARCH_OBJDEF,[Constants.OT_FUNC, Constants.OT_INFO_CARR])
var objDefs = selectedDatabase.Find(Constants.SEARCH_OBJDEF, -1)
var rowCounter = 1;
for (var i = 0; i < objDefs.length; i++) {

    var objName = objDefs[i].Name(nLocale)

    //for debug
    //if(objName == "File at Team Site")
    //var a = 2

    var objType = objDefs[i].Type()
    var objGUID = objDefs[i].GUID()

    var link_1Att = objDefs[i].Attribute(Constants.AT_EXT_1, nLocale)
    var link_2Att = objDefs[i].Attribute(Constants.AT_EXT_2, nLocale)
    var link_3Att = objDefs[i].Attribute(Constants.AT_EXT_3, nLocale)
    var link_4Att = objDefs[i].Attribute(Constants.AT_LINK, nLocale)
    var link_5Att = objDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("fcba0a10-f523-11e6-004f-eabd1a8bb6a2"), nLocale)
    var link_6Att = objDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("6cb95000-19d2-11e7-004f-eabd1a8bb6a2"), nLocale)

    var link_1 = link_1Att.getValue()
    var link_2 = link_2Att.getValue()
    var link_3 = link_3Att.getValue()
    var link_4 = link_4Att.getValue()
    var link_5 = link_5Att.getValue()			
    var link_6 = link_6Att.getValue() 

    if (link_1Att.IsMaintained() || link_2Att.IsMaintained() || link_3Att.IsMaintained() || link_4Att.IsMaintained() || link_5Att.IsMaintained() || link_6Att.IsMaintained()) {
        var newRow = objectsSheet.createRow(rowCounter);
        rowCounter++

        var newCell = newRow.createCell(0);
        newCell.setCellValue(objGUID);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(1);
        newCell.setCellValue(objType);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(2);
        newCell.setCellValue(objName);
        //newCell.setCellStyle(cellStyle);

        try {
            var newCell = newRow.createCell(3);
            newCell.setCellValue(link_1);
            //newCell.setCellStyle(cellStyle);
        }
        catch (err) {
            newCell.setCellValue("Error, link not available");
        }

        try {
            var newCell = newRow.createCell(4);
            newCell.setCellValue(link_2);
            //newCell.setCellStyle(cellStyle);
        }
        catch (err) {
            newCell.setCellValue("Error, link not available");
        }

        try {
            var newCell = newRow.createCell(5);
            newCell.setCellValue(link_3);
            //newCell.setCellStyle(cellStyle);
        }
        catch (err) {
            newCell.setCellValue("Error, link not available");
        }

        try {
            var newCell = newRow.createCell(6);
            newCell.setCellValue(link_4);
            //newCell.setCellStyle(cellStyle);
        }
        catch (err) {
            newCell.setCellValue("Error, link not available");
        }

        try {
            var newCell = newRow.createCell(7);
            newCell.setCellValue(link_5);
            //newCell.setCellStyle(cellStyle);
        }
        catch (err) {
            newCell.setCellValue("Error, link not available");
        }

        try {
            var newCell = newRow.createCell(8);
            newCell.setCellValue(link_6);
            //newCell.setCellStyle(cellStyle);
        }
        catch (err) {
            newCell.setCellValue("Error, link not available");
        }
    }
}

//MODELS
var models = selectedDatabase.Find(Constants.SEARCH_MODEL, -1)
var rowCounter = 1;
for (var i = 0; i < models.length; i++) {

    var objName = models[i].Name(nLocale)

    var objType = models[i].Type()
    var objGUID = models[i].GUID()

    var link_1Att = models[i].Attribute(Constants.AT_EXT_1, nLocale)
    var link_2Att = models[i].Attribute(Constants.AT_EXT_2, nLocale)
    var link_3Att = models[i].Attribute(Constants.AT_EXT_3, nLocale)
    var link_4Att = models[i].Attribute(Constants.AT_LINK, nLocale)
    var link_5Att = models[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("fcba0a10-f523-11e6-004f-eabd1a8bb6a2"), nLocale)
    var link_6Att = models[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("6cb95000-19d2-11e7-004f-eabd1a8bb6a2"), nLocale)

    var link_1 = link_1Att.getValue()
    var link_2 = link_2Att.getValue()
    var link_3 = link_3Att.getValue()
    var link_4 = link_4Att.getValue()
    var link_5 = link_5Att.getValue()
    var link_6 = link_6Att.getValue()

    if (link_1Att.IsMaintained() || link_2Att.IsMaintained() || link_3Att.IsMaintained() || link_4Att.IsMaintained() || link_5Att.IsMaintained() || link_6Att.IsMaintained()) {
        var newRow = modelsSheet.createRow(rowCounter);
        rowCounter++

        var newCell = newRow.createCell(0);
        newCell.setCellValue(objGUID);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(1);
        newCell.setCellValue(objType);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(2);
        newCell.setCellValue(objName);
        //newCell.setCellStyle(cellStyle);

        try {
            var newCell = newRow.createCell(3);
            newCell.setCellValue(link_1);
            //newCell.setCellStyle(cellStyle);
        }
        catch (err) {
            newCell.setCellValue("Error, link not available");
        }

        try {
            var newCell = newRow.createCell(4);
            newCell.setCellValue(link_2);
            //newCell.setCellStyle(cellStyle);
        }
        catch (err) {
            newCell.setCellValue("Error, link not available");
        }

        try {
            var newCell = newRow.createCell(5);
            newCell.setCellValue(link_3);
            //newCell.setCellStyle(cellStyle);
        }
        catch (err) {
            newCell.setCellValue("Error, link not available");
        }

        try {
            var newCell = newRow.createCell(6);
            newCell.setCellValue(link_4);
            //newCell.setCellStyle(cellStyle);
        }
        catch (err) {
            newCell.setCellValue("Error, link not available");
        }

        try {
            var newCell = newRow.createCell(7);
            newCell.setCellValue(link_5);
            //newCell.setCellStyle(cellStyle);
        }
        catch (err) {
            newCell.setCellValue("Error, link not available");
        }

        try {
            var newCell = newRow.createCell(8);
            newCell.setCellValue(link_6);
            //newCell.setCellStyle(cellStyle);
        }
        catch (err) {
            newCell.setCellValue("Error, link not available");
        }
    }
}

excelObject.write()

function createHeaderRow(sheet) {
    var headerRow = sheet.createRow(0);
    //rowCounter++;
    var font = excelObject.createFont();
    cellStyle1 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(169, 208, 142), Constants.SOLID_FOREGROUND);
    cellStyle2 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(155, 194, 230), Constants.SOLID_FOREGROUND);
    cellStyle2.setLocked(true)
    cellStyle3 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(237, 28, 36), Constants.SOLID_FOREGROUND);
    cellStyle3.setLocked(true)
    cellStyle4 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(34, 177, 76), Constants.SOLID_FOREGROUND);

    var newCell = headerRow.createCell(0);
    newCell.setCellValue("GUID");
    newCell.setCellStyle(cellStyle3);

    var newCell = headerRow.createCell(1);
    newCell.setCellValue("Type");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(2);
    newCell.setCellValue("Name");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(3);
    newCell.setCellValue("Link 1");
    newCell.setCellStyle(cellStyle4);

    var newCell = headerRow.createCell(4);
    newCell.setCellValue("Link 2");
    newCell.setCellStyle(cellStyle4);

    var newCell = headerRow.createCell(5);
    newCell.setCellValue("Link 3");
    newCell.setCellStyle(cellStyle4);

    var newCell = headerRow.createCell(6);
    newCell.setCellValue("Link 4");
    newCell.setCellStyle(cellStyle4);

    var newCell = headerRow.createCell(7);
    newCell.setCellValue("Link 5");
    newCell.setCellStyle(cellStyle4);

    var newCell = headerRow.createCell(8);
    newCell.setCellValue("Link 6");
    newCell.setCellStyle(cellStyle4);

}

function RGB(r, g, b) {
    return (new java.awt.Color(r / 255.0, g / 255.0, b / 255.0, 1)).getRGB() & 0xFFFFFF
}