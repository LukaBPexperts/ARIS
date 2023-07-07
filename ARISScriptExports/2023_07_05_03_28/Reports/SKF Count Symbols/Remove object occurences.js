var nLocale = Context.getSelectedLanguage()
var oOutput = Context.createOutputObject()
var excelObject = Context.createExcelWorkbook(Context.getSelectedFile());
var br = 0;

var modelsSymbolsExcel = new Array()

var excelFile = getExcelFile() //get excell file
var sheets = excelFile.getSheets() //get all sheets

//get number of rows in excel table
var rowCount = 0
while (sheets[0].getCell(rowCount, 0) != null) {
    rowCount++
}

//get marked models
for (var j = 1; j < rowCount; j++) {
    Context.writeStatus("1. Reading excel file row " + (j + 1) + "/" + rowCount)
    try {
        var valueTemp = sheets[0].getCell(j, 9).getCellValue()
        if (valueTemp.toUpperCase() == "NO") {
            var modelsSymbolsTemp = new Array()

            modelsSymbolsTemp.push(sheets[0].getCell(j, 0).getCellValue())
            modelsSymbolsTemp.push(sheets[0].getCell(j, 1).getCellValue())
            modelsSymbolsTemp.push(sheets[0].getCell(j, 2).getCellValue())

            modelsSymbolsExcel.push(modelsSymbolsTemp)
        }
    } catch (e) {}
}

//setup excell file
//*************************************************************
var sheetName = "Sheet1"
var mainSheet = excelObject.createSheet(sheetName);
//mainSheet.setColumnWidth(0, 12000);
mainSheet.setDefaultColumnWidth(22)
var rowCounter = 0;

createHeaderRow(mainSheet)
//*************************************************************

var selectedDatabases = ArisData.getSelectedDatabases()
var selectedDatabase = selectedDatabases[0]

var models = selectedDatabase.Find(Constants.SEARCH_MODEL, null)

for (var i = 0; i < models.length; i++) {
    Context.writeStatus("2. Working on model " + (i + 1) + "/" + models.length)

    var modelTypeNum = models[i].TypeNum()
    var modelTypeName = models[i].Type() //used in output
    var modelGroup = models[i].Group()
    var modelPath = modelGroup.Path(nLocale)
    var modelName = models[i].Name(nLocale)


    for (var j = 0; j < modelsSymbolsExcel.length; j++) {
        try {
            if (modelsSymbolsExcel[j][0] == modelTypeName) {

                var occs = models[i].ObjOccList()

                for (var k = 0; k < occs.length; k++) {
                    var objTypeNum = occs[k].ObjDef().TypeNum()
                    var objTypeName = occs[k].ObjDef().Type()
                    var occTypeNum = occs[k].SymbolNum()
                    var occTypeName = occs[k].SymbolName()

                    if (modelsSymbolsExcel[j][1] == (objTypeName + " (" + objTypeNum + ")") && modelsSymbolsExcel[j][2] == (occTypeName + " (" + occTypeNum + ")")) {

                        //delete occ and create log file
                        var remFlag = occs[k].Remove()
                        if (remFlag == false) {
                            //print log
                            //*********************
                            var newRow = mainSheet.createRow(rowCounter);
                            rowCounter++

                            var newCell = newRow.createCell(0);
                            newCell.setCellValue(modelTypeName);
                            //newCell.setCellStyle(cellStyle);

                            var newCell = newRow.createCell(1);
                            newCell.setCellValue(objTypeName + " (" + objTypeNum + ")");
                            //newCell.setCellStyle(cellStyle);

                            var newCell = newRow.createCell(2);
                            newCell.setCellValue(occTypeName + " (" + occTypeNum + ")");
                            //newCell.setCellStyle(cellStyle);

                            var newCell = newRow.createCell(3);
                            newCell.setCellValue("Symbol is not removed from the database");
                            //newCell.setCellStyle(cellStyle);

                            var newCell = newRow.createCell(4);
                            newCell.setCellValue(modelPath + "/" + modelName);
                            //newCell.setCellStyle(cellStyle);           
                            //*********************                           
                        } else {
                            br = br + 1 //count how many objects are removed
                        }
                    }
                }
            }
        } catch (e) {}
    }
}

var newRow = mainSheet.createRow(rowCounter + 2);
rowCounter++

 var newCell = newRow.createCell(0);
 newCell.setCellValue("Symbols removed: " + br);
 //newCell.setCellStyle(cellStyle); 

excelObject.write()

function getExcelFile() {
    var sdefname = ""
    sdefext = __toString("*.xls!!Excel|*.xls; *.xlsx||")
    var sdefdir = ""
    var stitle = "Chose excel file"

    var files = Dialogs.BrowseForFiles(sdefname, sdefext, sdefdir, stitle, 0)

    var excelFile = Context.getExcelReader(files[0].getData())

    return excelFile;
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
    newCell.setCellValue("Model type");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(1);
    newCell.setCellValue("Object type");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(2);
    newCell.setCellValue("Symbol type");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(3);
    newCell.setCellValue("Status");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(4);
    newCell.setCellValue("Model path");
    newCell.setCellStyle(cellStyle2);

}

function RGB(r, g, b) {
    return (new java.awt.Color(r / 255.0, g / 255.0, b / 255.0, 1)).getRGB() & 0xFFFFFF
}