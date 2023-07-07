var modelTypeDelete = Constants.MT_BPMN_COLLABORATION_DIAGRAM
var modelTypeCreate = Constants.MT_VAL_ADD_CHN_DGM

var excelObject = Context.createExcelWorkbook(Context.getSelectedFile());
var nLocale = Context.getSelectedLanguage()
var selectedDatabases = ArisData.getSelectedDatabases()
var selectedDatabase = selectedDatabases[0]

var bpmnModelsAll = selectedDatabase.Find(Constants.SEARCH_MODEL, [modelTypeDelete])

//get models for delete
var bpmnModels = new Array()
for (var i = 0; i < bpmnModelsAll.length; i++) {

    var occList1 = bpmnModelsAll[i].ObjOccListBySymbol([ArisData.ActiveFilter().UserDefinedSymbolTypeNum("b12afd21-975f-11e5-7464-0050560116c9")]) //TU STAVI API NAME OD SIMBOLA PRIMJER JE: "ST_BPMN_SUBPROCESS"
    var occList2 = bpmnModelsAll[i].ObjOccListBySymbol(Constants.ST_BPMN_SUBPROCESS)
    var occList3 = bpmnModelsAll[i].ObjOccListBySymbol(Constants.ST_BPMN_TASK)

    if (occList1.length > 0 && occList2.length == 0 && occList3.length == 0) {
        bpmnModels.push(bpmnModelsAll[i])
    }
}

bpmnModels = ArisData.Unique(bpmnModels)

//generate new and delete old models
var modelGeneration = Context.getComponent("ModelGeneration")
var generatorOptions = modelGeneration.createModelGenerationOptions()

var newModels = new Array()
for (var i = 0; i < bpmnModels.length; i++) {
    var modelName = bpmnModels[i].Name(nLocale)
    var modelGroup = bpmnModels[i].Group()

    var modelTypedArray = ArisData.createTypedArray(Constants.CID_MODEL);
    modelTypedArray.push(bpmnModels[i])

    var generatedModel = modelGeneration.generateModel(modelTypedArray, modelName, modelTypeCreate, modelGroup, generatorOptions)
    newModels.push(generatedModel)

    var deleteModel = modelGroup.Delete(bpmnModels[i])
}

//setup excell file
//*************************************************************
var sheetName = "Sheet1"
var mainSheet = excelObject.createSheet(sheetName);
//mainSheet.setColumnWidth(0, 12000);
mainSheet.setDefaultColumnWidth(20)
var rowCounter = 0;

createHeaderRow(mainSheet)
//*************************************************************

for (var i = 0; i < newModels.length; i++) {

    var modelName = newModels[i].Name(nLocale)
    var modelPath = newModels[i].Group().Path(nLocale)
    var modelGuid = newModels[i].GUID()

    var newRow = mainSheet.createRow(rowCounter);
    rowCounter++

    var newCell = newRow.createCell(0);
    newCell.setCellValue(modelName);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(1);
    newCell.setCellValue(modelPath);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(2);
    newCell.setCellValue(modelGuid);
    //newCell.setCellStyle(cellStyle);   
}

excelObject.write()

function createHeaderRow(headerRow) {
    var headerRow = mainSheet.createRow(rowCounter);
    rowCounter++;
    var font = excelObject.createFont();
    cellStyleObject = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(102, 136, 102), Constants.SOLID_FOREGROUND);
    cellStyleCxn = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, RGB(146, 208, 80), Constants.SOLID_FOREGROUND);
    cellStyleMark = excelObject.createCellStyle(font, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.C_GREY_80_PERCENT, Constants.ALIGN_LEFT, Constants.ALIGN_LEFT, Constants.C_TRANSPARENT, Constants.C_RED, Constants.SOLID_FOREGROUND);

    var newCell = headerRow.createCell(0);
    newCell.setCellValue("Model name");
    newCell.setCellStyle(cellStyleObject);

    var newCell = headerRow.createCell(1);
    newCell.setCellValue("Path");
    newCell.setCellStyle(cellStyleObject);

    var newCell = headerRow.createCell(2);
    newCell.setCellValue("GUID");
    newCell.setCellStyle(cellStyleObject);
}

function RGB(r, g, b) {
    return (new java.awt.Color(r / 255.0, g / 255.0, b / 255.0, 1)).getRGB() & 0xFFFFFF
}