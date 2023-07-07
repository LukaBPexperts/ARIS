var nLocale = Context.getSelectedLanguage()
var oOutput = Context.createOutputObject()
var excelObject = Context.createExcelWorkbook(Context.getSelectedFile());

//setup excell file
//*************************************************************
var sheetName = "Sheet1"
var mainSheet = excelObject.createSheet(sheetName);
//mainSheet.setColumnWidth(0, 12000);
//mainSheet.setDefaultColumnWidth(22)
mainSheet.setColumnWidth(0, 15000)
mainSheet.setColumnWidth(1, 20000)
var rowCounter = 0;

createHeaderRow(mainSheet)
//*************************************************************

var selectedModels = ArisData.getSelectedModels()

for (var i = 0; i < selectedModels.length; i++) {
    var activityObjects = selectedModels[i].ObjDefListByTypes([Constants.OT_FUNC])

    for (var j = 0; j < activityObjects.length; j++) {
        var activityUrl = getPortalLink(activityObjects[j])
        var activityName = activityObjects[j].Name(nLocale)

        var newRow = mainSheet.createRow(rowCounter);
        rowCounter++

        var newCell = newRow.createCell(0);
        newCell.setCellValue(activityName);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(1);
        newCell.setCellValue(activityUrl);
        //newCell.setCellStyle(cellStyle);
    }
}

excelObject.write()

function createHeaderRow(headerRow) {
    var headerRow = mainSheet.createRow(rowCounter);
    rowCounter++;
    var font = excelObject.createFont();
    cellStyle1 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(169, 208, 142), Constants.SOLID_FOREGROUND);
    cellStyle2 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(155, 194, 230), Constants.SOLID_FOREGROUND);
    //cellStyle2.setLocked(true)
    cellStyle3 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(255, 217, 102), Constants.SOLID_FOREGROUND);

    var newCell = headerRow.createCell(0);
    newCell.setCellValue("Activity Name");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(1);
    newCell.setCellValue("VPP Link");
    newCell.setCellStyle(cellStyle2);

}

function getPortalLink(selectedObject) {
    //http://<servername>/#<tenant>/item/c.<factSheet>.<DBName>.<GUID>.<Version>

    var oDatabase = ArisData.getActiveDatabase();
    var sTenantName = ArisData.getTenantName();
    var serverName = oDatabase.ServerName();
    var sDBName = oDatabase.Name(nLocale);
    sDBName = sDBName.replace(" ", "%20");
    var mGUID = selectedObject.GUID()

    var linkURL = "https://" + serverName + "/#" + sTenantName + "/item/c." + sDBName + "." + mGUID + ".-1"
    return linkURL
}

function RGB(r, g, b) {
    return (new java.awt.Color(r / 255.0, g / 255.0, b / 255.0, 1)).getRGB() & 0xFFFFFF
}