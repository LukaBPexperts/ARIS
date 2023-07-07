var modelTypeDelete = Constants.MT_EEPC
var modelTypeCreate = Constants.MT_VAL_ADD_CHN_DGM

//var modelTypeCreate = ArisData.ActiveFilter().UserDefinedModelTypeNum("af9f2f40-78aa-11ea-2ba6-0a002700001a")



var excelObject = Context.createExcelWorkbook(Context.getSelectedFile());
var nLocale = Context.getSelectedLanguage()
var selectedDatabases = ArisData.getSelectedDatabases()
var selectedDatabase = selectedDatabases[0]

var bpmnModelsAll = selectedDatabase.Find(Constants.SEARCH_MODEL, [modelTypeDelete])

//get models for delete
var bpmnModels = new Array()
/* for (var i = 0; i < bpmnModelsAll.length; i++) {

    var occList1 = bpmnModelsAll[i].ObjOccListBySymbol([ArisData.ActiveFilter().UserDefinedSymbolTypeNum("b12afd21-975f-11e5-7464-0050560116c9")]) //TU STAVI API NAME OD SIMBOLA PRIMJER JE: "ST_BPMN_SUBPROCESS"
    var occList2 = bpmnModelsAll[i].ObjOccListBySymbol(Constants.ST_BPMN_SUBPROCESS)
    var occList3 = bpmnModelsAll[i].ObjOccListBySymbol(Constants.ST_BPMN_TASK)

    if (occList1.length > 0 && occList2.length == 0 && occList3.length == 0) {
        bpmnModels.push(bpmnModelsAll[i])
    }
}
 */
bpmnModels = bpmnModelsAll

//generate new and delete old models
var modelGeneration = Context.getComponent("ModelGeneration")
var generatorOptions = modelGeneration.createModelGenerationOptions()
	
generatorOptions.onlyCxnsWithOccs(true)

var newModels = new Array()
for (var i = 0; i < bpmnModels.length; i++) {
    var modelName = bpmnModels[i].Name(nLocale)
    var modelGroup = bpmnModels[i].Group()
    
    var objOccs = bpmnModels[i].ObjOccList()
    var cxnOccs = bpmnModels[i].CxnOccList()
    
    //create model
    var newModel = modelGroup.CreateModel(modelTypeCreate, modelName, nLocale)
    
    for (var j=0; j<objOccs.length; j++){
        var occSymbol = objOccs[j].getSymbol()
        var occDef = objOccs[j].ObjDef()
        var occX = objOccs[j].X()
        var occY = objOccs[j].Y()
        var newObjOcc = newModel.createObjOcc(occSymbol, occDef, occX, occY, true)
    }
    
    for (var j=0; j<cxnOccs.length; j++){
        var sourceObjOcc =  getObjectOccSource(cxnOccs[j], newModel)  
        var targetObjOcc = getObjectOccTarget(cxnOccs[j], newModel)  
        var cxnType =  Constants.CT_IS_PREDEC_OF_1
        var pointList = cxnOccs[j].getPoints()
        
        var newCxn =newModel.CreateCxnOcc(sourceObjOcc, targetObjOcc, cxnType, pointList)
    }
    
    
    
    
    //doLayout on geberatedModel
    //layoutModel = ArisData.getModelOptions().getLayoutParameters(modelTypeCreate)
    //layoutModel.setSpacingX(100)
    //layoutModel.setSpacingY(100)
    
    //var layoutFlag = generatedModel.doLayout()
    
    //newModels.push(generatedModel)

    //var deleteModel = modelGroup.Delete(bpmnModels[i])
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

function getObjectOccSource(cxnOcc, newModel){   
    var sourceObjOcc = cxnOcc.getSource()
    var sourceObjOccX = sourceObjOcc.X()
    var sourceObjOccY = sourceObjOcc.Y()
    var sourceObjDef = sourceObjOcc.ObjDef()
    
    var occListModel = sourceObjDef.OccListInModel(newModel)
    
    for(var i=0; i<occListModel.length; i++){
        var tempX = occListModel[i].X()
        var tempY = occListModel[i].Y()
        
        if((tempX == sourceObjOccX) && (tempY == sourceObjOccY)){
         
            return occListModel[i]
        }
    }
}

function getObjectOccTarget(cxnOcc, newModel){
    var targetObjOcc = cxnOcc.getTarget()
    var targetObjOccX = targetObjOcc.X()
    var targetObjOccY = targetObjOcc.Y()
    var targetObjDef = targetObjOcc.ObjDef()
    
    var occListModel = targetObjDef.OccListInModel(newModel)
    
    for(var i=0; i<occListModel.length; i++){
        var tempX = occListModel[i].X()
        var tempY = occListModel[i].Y()
        
        if((tempX == targetObjOccX) && (tempY == targetObjOccY)){
         
            return occListModel[i]
        }
    }
}