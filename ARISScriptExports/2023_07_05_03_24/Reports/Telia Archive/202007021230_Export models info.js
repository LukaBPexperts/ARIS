var nLocale = Context.getSelectedLanguage()
//var oOutput = Context.createOutputObject()
var excelObject = Context.createExcelWorkbook(Context.getSelectedFile());

var selectedGroups = ArisData.getSelectedGroups()

//setup excell file
//*************************************************************
var sheetName = "Sheet1"
var mainSheet = excelObject.createSheet(sheetName);
//mainSheet.setColumnWidth(0, 12000);
mainSheet.setDefaultColumnWidth(22)
var rowCounter = 0;

createHeaderRow(mainSheet)
//*************************************************************

for (var i = 0; i < selectedGroups.length; i++) {
    //var models = selectedGroups[i].ModelList(true, [Constants.MT_ABS_COMPOSITE_SCHEDULE, Constants.MT_ACS_DGM])
  
    var objDefs = selectedGroups[i].ObjDefList(true, null)

    for (var j = 0; j < objDefs.length; j++) {

        var objName = objDefs[j].Name(nLocale)
        var objType = objDefs[j].Type()
        var objGUID = objDefs[j].GUID()
        var objDesc = objDefs[j].Attribute(Constants.AT_DESC, nLocale).getValue()
        
        //var leveragePotentialAsIsCompliance = capabilityObjDefs[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("d9bc0080-89b7-11ec-1a81-0022487eead8"), nLocale).getValue()
        
        /* var personResponsible = models[j].Attribute(Constants.AT_PERS_RESP, nLocale).getValue()
        var modelPath = models[j].Group().Path(nLocale)
        var modelLink = getPortalLink(models[j])
        var lastUser = models[j].getLastModifier()
        var lastChangeDate = models[j].Attribute(Constants.AT_LAST_CHNG_2, nLocale).getValue() */
        //var archivingAssessmentDate = models[j].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("8fda52b0-6188-11de-4732-00237d347a8a"), nLocale).getValue()
        //var archivingStatus = models[j].Attribute(Constants.AT, nLocale).getValue()

        
        if(objDesc == ""){
        
        var newRow = mainSheet.createRow(rowCounter);
        rowCounter++

        var newCell = newRow.createCell(0);
        newCell.setCellValue(objName);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(1);
        newCell.setCellValue(objType);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(2);
        newCell.setCellValue(objGUID);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(3);
        newCell.setCellValue(objDesc);
        //newCell.setCellStyle(cellStyle);
        
        }

       /*  var newCell = newRow.createCell(4);
        newCell.setCellValue(modelPath);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(5);
        newCell.setCellValue(modelLink);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(6);
        newCell.setCellValue(lastUser);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(7);
        newCell.setCellValue(lastChangeDate);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(8);
        newCell.setCellValue("archivingAssessmentDate");
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(9);
        newCell.setCellValue("archivingStatus");
        //newCell.setCellStyle(cellStyle); */

    }

}

excelObject.write()


function createHeaderRow(headerRow) {
    var headerRow = mainSheet.createRow(rowCounter);
    rowCounter++;
    var font = excelObject.createFont();
    cellStyle1 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(169, 208, 142), Constants.SOLID_FOREGROUND);
    cellStyle2 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(155, 194, 230), Constants.SOLID_FOREGROUND);
    cellStyle2.setLocked(true)
    cellStyle3 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(255, 217, 102), Constants.SOLID_FOREGROUND);

    var newCell = headerRow.createCell(0);
    newCell.setCellValue("Name");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(1);
    newCell.setCellValue("Type");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(2);
    newCell.setCellValue("GUID");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(3);
    newCell.setCellValue("Description");
    newCell.setCellStyle(cellStyle2);

  /*   var newCell = headerRow.createCell(4);
    newCell.setCellValue("Model path");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(5);
    newCell.setCellValue("Portal link to model");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(6);
    newCell.setCellValue("Last user");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(7);
    newCell.setCellValue("Last change date");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(8);
    newCell.setCellValue("Archiving assessment date");
    newCell.setCellStyle(cellStyle3);

    var newCell = headerRow.createCell(9);
    newCell.setCellValue("Archiving status");
    newCell.setCellStyle(cellStyle3); */



}

function getPortalLink(model) {
    //http://<servername>/#<tenant>/item/c.<factSheet>.<DBName>.<GUID>.<Version>

    var oDatabase = ArisData.getActiveDatabase();
    var sTenantName = ArisData.getTenantName();
    var serverName = oDatabase.ServerName();
    var sDBName = oDatabase.Name(nLocale);
    sDBName = sDBName.replace(" ", "%20");
    var mGUID = model.GUID()

    var linkURL = "http://" + serverName + "/#" + sTenantName + "/item/c." + sDBName + "." + mGUID + ".-1"
    return linkURL
}

function RGB(r, g, b) {
    return (new java.awt.Color(r / 255.0, g / 255.0, b / 255.0, 1)).getRGB() & 0xFFFFFF
}