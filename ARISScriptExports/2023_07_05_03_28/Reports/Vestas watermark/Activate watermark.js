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

var selectedGroups = ArisData.getSelectedGroups()

var models = new Array()
for (var i = 0; i < selectedGroups.length; i++) {
    var tempModels = selectedGroups[i].ModelList(true, null)
    models = models.concat(tempModels)
}

for (var i = 0; i < models.length; i++) {
    Context.writeStatus("Working on model: " + (i + 1) + '/' + (models.length + 1))
    
    var modelWrite = models[i].canWrite(true) 
    
    var flag = false
    if(modelWrite == true){
    var flag = models[i].changeFlag(Constants.MODEL_FILLBYLOGO, true)
    }
    
    var modelName = models[i].Name(nLocale)
    var modelTypeName = models[i].Type() //used in output
    var modelGroup = models[i].Group()
    var modelPath = modelGroup.Path(nLocale)

    if (flag == true && modelWrite == true) {
        //*********************
        var newRow = mainSheet.createRow(rowCounter);
        rowCounter++

        var newCell = newRow.createCell(0);
        newCell.setCellValue(modelName);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(1);
        newCell.setCellValue(modelTypeName);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(2);
        newCell.setCellValue(modelPath);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(3);
        newCell.setCellValue("Watermark Activated");
        //newCell.setCellStyle(cellStyle);
        //*********************
    } else {
        //*********************
        var newRow = mainSheet.createRow(rowCounter);
        rowCounter++

        var newCell = newRow.createCell(0);
        newCell.setCellValue(modelName);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(1);
        newCell.setCellValue(modelTypeName);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(2);
        newCell.setCellValue(modelPath);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(3);
        newCell.setCellValue("Error, watermark is not activated");
        //newCell.setCellStyle(cellStyle);
        //*********************


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
    newCell.setCellValue("Model name");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(1);
    newCell.setCellValue("Model Type");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(2);
    newCell.setCellValue("Model Location");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(3);
    newCell.setCellValue("Status");
    newCell.setCellStyle(cellStyle2);
}

function RGB(r, g, b) {
    return (new java.awt.Color(r / 255.0, g / 255.0, b / 255.0, 1)).getRGB() & 0xFFFFFF
}