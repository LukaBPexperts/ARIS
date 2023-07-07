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

var selectedDatabases = ArisData.getSelectedDatabases()

var selectedDatabase = selectedDatabases[0]

var models = selectedDatabase.Find(Constants.SEARCH_MODEL, null)

var modelTypes = new Array()
for (var i = 0; i < models.length; i++) {

    modelTypes.push(models[i].TypeNum())

}

modelTypes = ArisData.Unique(modelTypes)

for (var i = 0; i < modelTypes.length; i++) {
    Context.writeStatus("Analyzing model type: " + (i + 1) + '/' + (modelTypes.length + 1))

    var models = selectedDatabase.Find(Constants.SEARCH_MODEL, modelTypes[i])
    var modelTypeName = models[0].Type() //used in output

    var objectOccs = new Array()
    for (var j = 0; j < models.length; j++) {
        var occs = models[j].ObjOccList()
        objectOccs = objectOccs.concat(occs)
    }
    objectOccs = ArisData.Unique(objectOccs)

    var objTypes = getObjectsType(objectOccs)

    for (var j = 0; j < objTypes.length; j++) {

        var symbols = new Array()
        var symbolsUnique = new Array()
        for (var k = 0; k < objectOccs.length; k++) {
            if (objectOccs[k].ObjDef().TypeNum() == objTypes[j][0]) {
                //if(objectOccs[k].ObjDef().Type() == objTypes[j]){
                //get symbol and count it or save in array
                //symbols.push(objectOccs[k].SymbolNum())
                //symbols.push(objectOccs[k].SymbolName())

                var symbolsTemp = new Array()
                symbols.push(objectOccs[k].SymbolNum())

                symbolsTemp.push(objectOccs[k].SymbolNum())
                symbolsTemp.push(objectOccs[k].SymbolName())

                symbolsUnique.push(symbolsTemp)

            }
        }

        //symbolsUnique = ArisData.Unique(symbols)
        symbolsUnique = uniqueMultiArray(symbolsUnique) //return unique pairs
        for (var k = 0; k < symbolsUnique.length; k++) {
            var br = 0;
            for (var l = 0; l < symbols.length; l++) {
                if (symbolsUnique[k][0] == symbols[l]) {
                    br = br + 1
                }
            }

            //*********************
            var newRow = mainSheet.createRow(rowCounter);
            rowCounter++

            var newCell = newRow.createCell(0);
            newCell.setCellValue(modelTypeName);
            //newCell.setCellStyle(cellStyle);

            var newCell = newRow.createCell(1);
            newCell.setCellValue(objTypes[j][1] + " (" + objTypes[j][0] + ")");
            //newCell.setCellStyle(cellStyle);

            var newCell = newRow.createCell(2);
            newCell.setCellValue(symbolsUnique[k][1] + " (" + symbolsUnique[k][0] + ")");
            //newCell.setCellStyle(cellStyle);

            var newCell = newRow.createCell(3);
            newCell.setCellValue(br);
            //newCell.setCellStyle(cellStyle);
            //*********************
        }
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
    newCell.setCellValue("Model type");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(1);
    newCell.setCellValue("Object type");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(2);
    newCell.setCellValue("Symbol type");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(3);
    newCell.setCellValue("Number of occurrences");
    newCell.setCellStyle(cellStyle2);

}

function getObjectsType(objectOccs) {
    var returnTypePlusName = new Array()

    for (var i = 0; i < objectOccs.length; i++) {
        var objTypesNameTemp = new Array()

        objTypesNameTemp.push(objectOccs[i].ObjDef().TypeNum())
        objTypesNameTemp.push(objectOccs[i].ObjDef().Type())
        //objTypes.push(objectOccs[i].ObjDef().Type())

        returnTypePlusName.push(objTypesNameTemp)
    }
    returnTypePlusName = uniqueMultiArray(returnTypePlusName) //return unique pairs

    return returnTypePlusName

}

function uniqueMultiArray(array) {
    for (var i = 0; i < array.length; i++) {
        for (var j = i + 1; j < array.length;) {
            if (array[i][0] == array[j][0] && array[i][1] == array[j][1])
                array.splice(j, 1);
            else
                j++;
        }
    }
    return array
}

function RGB(r, g, b) {
    return (new java.awt.Color(r / 255.0, g / 255.0, b / 255.0, 1)).getRGB() & 0xFFFFFF
}