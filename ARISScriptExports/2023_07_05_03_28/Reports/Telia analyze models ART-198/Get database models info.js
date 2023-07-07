var nLocale = Context.getSelectedLanguage()
var oOutput = Context.createOutputObject()
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
    var models = selectedGroups[i].ModelList(true, null)

    var groupName = selectedGroups[i].Name(nLocale)
    var numberTotal = models.length
    var numberApproved = getNumberApproved(models)
    var numberPersonResponEmpty = getNumberPersonResponEmpty(models)
    var numberLast1Year = getNumberLastChange1Year(models, 1)
    var numberLast2Year = getNumberLastChange1YearRetention(models, 2, 5)
    var numberLast5Year = getNumberLastChange1YearRetention(models, 5, 10)
    var numberLast10Year = getNumberLastChange1YearRetention(models, 10, null)

    var newRow = mainSheet.createRow(rowCounter);
    rowCounter++

    var newCell = newRow.createCell(0);
    newCell.setCellValue(groupName);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(1);
    newCell.setCellValue(numberTotal);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(2);
    newCell.setCellValue(numberApproved);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(3);
    newCell.setCellValue(numberPersonResponEmpty);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(4);
    newCell.setCellValue(numberLast1Year);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(5);
    newCell.setCellValue(numberLast2Year);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(6);
    newCell.setCellValue(numberLast5Year);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(7);
    newCell.setCellValue(numberLast10Year);
    //newCell.setCellStyle(cellStyle);

}

excelObject.write()

function getNumberLastChange1Year(models, yearBr) {
    var br = 0
    for (var i = 0; i < models.length; i++) {

        var lastMod = models[i].Attribute(Constants.AT_LAST_CHNG_2, nLocale).getValue()
        var lastModDate = new Date(lastMod)

        var date = new Date();
        date.setFullYear(date.getFullYear() - yearBr);

        if (lastModDate.getTime() > date.getTime()) {
            br++
        }
    }
    return br
}

function getNumberLastChange1YearRetention(models, yearBr, yearL) {
    var br = 0
    for (var i = 0; i < models.length; i++) {

        var lastMod = models[i].Attribute(Constants.AT_LAST_CHNG_2, nLocale).getValue()
        var lastModDate = new Date(lastMod)

        var date = new Date();
        date.setFullYear(date.getFullYear() - yearBr);

        if (yearL != null) {

            var dateL = new Date();
            dateL.setFullYear(dateL.getFullYear() - yearL);

            if (lastModDate.getTime() < date.getTime() && lastModDate.getTime() > dateL.getTime()) {
                br++
            }
        } else if (lastModDate.getTime() < date.getTime()) {
            br++
        }
    }
    return br
}

function getNumberPersonResponEmpty(models) {
    var br = 0;
    for (var i = 0; i < models.length; i++) {
        var personResponsibleFlag = models[i].Attribute(Constants.AT_PERS_RESP, nLocale).IsMaintained()

        if (personResponsibleFlag == false) {
            br++
        }
    }
    return br
}

function getNumberApproved(models) {
    var br = 0;
    for (var i = 0; i < models.length; i++) {
        var approvedValue = models[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("d2ac7740-af65-11df-599b-005056aa0ab6"), nLocale).getValue()

        if (approvedValue == "Approved") {
            br++
        }
    }
    return br

}

function createHeaderRow(headerRow) {
    var headerRow = mainSheet.createRow(rowCounter);
    rowCounter++;
    var font = excelObject.createFont();
    cellStyle1 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(169, 208, 142), Constants.SOLID_FOREGROUND);
    cellStyle2 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(155, 194, 230), Constants.SOLID_FOREGROUND);
    cellStyle3 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(255, 217, 102), Constants.SOLID_FOREGROUND);

    var newCell = headerRow.createCell(0);
    newCell.setCellValue("Top Level Group Name");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(1);
    newCell.setCellValue("Total Number of Models");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(2);
    newCell.setCellValue("Approved models");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(3);
    newCell.setCellValue("Person Responsible empty");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(4);
    newCell.setCellValue("Last change < 1 Year");
    newCell.setCellStyle(cellStyle3);

    var newCell = headerRow.createCell(5);
    newCell.setCellValue("Last change > 2 and < 5 Years");
    newCell.setCellStyle(cellStyle3);

    var newCell = headerRow.createCell(6);
    newCell.setCellValue("Last change > 5 and < 10 Years");
    newCell.setCellStyle(cellStyle3);

    var newCell = headerRow.createCell(7);
    newCell.setCellValue("Last change > 10 Years");
    newCell.setCellStyle(cellStyle3);

}

function RGB(r, g, b) {
    return (new java.awt.Color(r / 255.0, g / 255.0, b / 255.0, 1)).getRGB() & 0xFFFFFF
}