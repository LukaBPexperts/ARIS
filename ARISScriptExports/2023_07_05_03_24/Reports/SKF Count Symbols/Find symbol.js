var nLocale = Context.getSelectedLanguage()
var oOutput = Context.createOutputObject()
var excelObject = Context.createExcelWorkbook(Context.getSelectedFile());

//setup excell file
//*************************************************************
var sheetName = "Sheet1"
var mainSheet = excelObject.createSheet(sheetName);
//mainSheet.setColumnWidth(0, 12000);
mainSheet.setDefaultColumnWidth(22)
var rowCounter = 0;

createHeaderRow(mainSheet)
//*************************************************************

var sValue = Dialogs.InputBox("Enter symbol number", "Symbol number", null)

if (sValue != "") {

    var selectedDatabases = ArisData.getSelectedDatabases()
    var selectedGroups = ArisData.getSelectedGroups()

    if (selectedDatabases.length != 0) {
        var selectedDatabase = selectedDatabases[0]
        var models = selectedDatabase.Find(Constants.SEARCH_MODEL, null)
    } else if (selectedGroups.length != 0) {
        var models = new Array()
        for (var i = 0; i < selectedGroups.length; i++) {
            var tempModels = selectedGroups[i].ModelList(true, null)
            models = models.concat(tempModels)
        }
    }

    var modelTypes = new Array()
    for (var i = 0; i < models.length; i++) {

        var occs = models[i].ObjOccList()

        for (var j = 0; j < occs.length; j++) {

            var occType = occs[j].SymbolNum()
            var occName = occs[j].SymbolName()

            var modelGroup = models[i].Group()
            var modelPath = modelGroup.Path(nLocale)

            if (occType == sValue) {
                //print out
                //*********************
                var newRow = mainSheet.createRow(rowCounter);
                rowCounter++

                var newCell = newRow.createCell(0);
                newCell.setCellValue(occName + " (" + occType + ")");
                //newCell.setCellStyle(cellStyle);

                var newCell = newRow.createCell(1);
                newCell.setCellValue(models[i].Name(nLocale));
                //newCell.setCellStyle(cellStyle);

                var newCell = newRow.createCell(2);
                newCell.setCellValue(models[i].Type());
                //newCell.setCellStyle(cellStyle);

                var newCell = newRow.createCell(3);
                newCell.setCellValue(modelPath);
                //newCell.setCellStyle(cellStyle);
                //*********************          
            }
        }
    }

    excelObject.write()

}

function createHeaderRow(headerRow) {
    var headerRow = mainSheet.createRow(rowCounter);
    rowCounter++;
    var font = excelObject.createFont();
    cellStyle1 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(169, 208, 142), Constants.SOLID_FOREGROUND);
    cellStyle2 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(155, 194, 230), Constants.SOLID_FOREGROUND);
    cellStyle2.setLocked(true)
    cellStyle3 = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(255, 217, 102), Constants.SOLID_FOREGROUND);

    var newCell = headerRow.createCell(0);
    newCell.setCellValue("Symbol Name (Number)");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(1);
    newCell.setCellValue("Occurrence in model (Name)");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(2);
    newCell.setCellValue("Model type");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(3);
    newCell.setCellValue("Model path");
    newCell.setCellStyle(cellStyle2);

}

function RGB(r, g, b) {
    return (new java.awt.Color(r / 255.0, g / 255.0, b / 255.0, 1)).getRGB() & 0xFFFFFF
}