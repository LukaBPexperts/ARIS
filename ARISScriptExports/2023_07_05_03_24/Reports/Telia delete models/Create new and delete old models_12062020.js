var modelTypeDelete = Constants.MT_BPMN_COLLABORATION_DIAGRAM
var modelTypeCreate = ArisData.ActiveFilter().UserDefinedModelTypeNum("9f0ac401-0e0f-11e6-7464-0050560116c9")
var cxnTypeCreate = Constants.CT_IS_PREDEC_OF_1

var excelObject = Context.createExcelWorkbook(Context.getSelectedFile());
var nLocale = Context.getSelectedLanguage()
var selectedDatabases = ArisData.getSelectedDatabases()
var selectedDatabase = selectedDatabases[0]

var bpmnModelsAll = selectedDatabase.Find(Constants.SEARCH_MODEL, [modelTypeDelete])

//get models for delete
var bpmnModels = new Array()
for (var i = 0; i < bpmnModelsAll.length; i++) {

    var occList = bpmnModelsAll[i].ObjOccListBySymbol([ArisData.ActiveFilter().UserDefinedSymbolTypeNum("b12afd21-975f-11e5-7464-0050560116c9")]) //TU STAVI API NAME OD SIMBOLA PRIMJER JE: "ST_BPMN_SUBPROCESS"
    var occList2 = bpmnModelsAll[i].ObjOccListBySymbol(Constants.ST_BPMN_SUBPROCESS)
    var occList3 = bpmnModelsAll[i].ObjOccListBySymbol(Constants.ST_BPMN_TASK)
    var occList4 = bpmnModelsAll[i].ObjOccListBySymbol(Constants.ST_BPMN_DATA_STORE)
    var occList5 = bpmnModelsAll[i].ObjOccListBySymbol(Constants.ST_BPMN_CALL_ACTIVITY)
    var occList6 = bpmnModelsAll[i].ObjOccListBySymbol(Constants.ST_BPMN_CALL_ACTIVITY_COLLAPSED)
    var occList7 = bpmnModelsAll[i].ObjOccListBySymbol(Constants.ST_BPMN_DATA_OBJECT)
    var occList8 = bpmnModelsAll[i].ObjOccListBySymbol(Constants.ST_INFO_OBJ)
    var occList9 = bpmnModelsAll[i].ObjOccListBySymbol(Constants.ST_BPMN_GROUPING)
    var occList10 = bpmnModelsAll[i].ObjOccListBySymbol(Constants.ST_BPMN_GROUPING_1)
    var occList11 = bpmnModelsAll[i].ObjOccListBySymbol(Constants.ST_BPMN_MESSAGE_2)
    var occList12 = bpmnModelsAll[i].ObjOccListBySymbol(Constants.ST_BPMN_MESSAGE)


    if (occList.length > 0 && occList2.length == 0 && occList3.length == 0 && occList4.length == 0 && occList5.length == 0 && occList6.length == 0 && occList7.length == 0 && occList8.length == 0 && occList9.length == 0 && occList10.length == 0 && occList11.length == 0 && occList12.length == 0) {
        //puniNiz
        bpmnModels.push(bpmnModelsAll[i])
    }
}


bpmnModels = ArisData.Unique(bpmnModels)

var newModels = new Array()
for (var i = 0; i < bpmnModels.length; i++) {
    var modelName = bpmnModels[i].Name(nLocale)
    var modelGroup = bpmnModels[i].Group()

    var objOccs = bpmnModels[i].ObjOccList()
    var cxnOccs = bpmnModels[i].CxnOccList()

    //create model
    var newModel = modelGroup.CreateModel(modelTypeCreate, modelName, nLocale)

    for (var j = 0; j < objOccs.length; j++) {
        var occSymbol = objOccs[j].getSymbol()
        var occDef = objOccs[j].ObjDef()
        var occX = objOccs[j].X()
        var occY = objOccs[j].Y()
        var newObjOcc = newModel.createObjOcc(occSymbol, occDef, occX, occY, true)
    }

    for (var j = 0; j < cxnOccs.length; j++) {
        var sourceObjOcc = getObjectOccSource(cxnOccs[j], newModel)
        var targetObjOcc = getObjectOccTarget(cxnOccs[j], newModel)
        var cxnType = cxnTypeCreate
        var pointList = cxnOccs[j].getPoints()

        var newCxn = newModel.CreateCxnOcc(sourceObjOcc, targetObjOcc, cxnType, pointList)
    }

    newModels.push(newModel)

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

function getObjectOccSource(cxnOcc, newModel) {
    var sourceObjOcc = cxnOcc.getSource()
    var sourceObjOccX = sourceObjOcc.X()
    var sourceObjOccY = sourceObjOcc.Y()
    var sourceObjDef = sourceObjOcc.ObjDef()

    var occListModel = sourceObjDef.OccListInModel(newModel)

    for (var i = 0; i < occListModel.length; i++) {
        var tempX = occListModel[i].X()
        var tempY = occListModel[i].Y()

        if ((tempX == sourceObjOccX) && (tempY == sourceObjOccY)) {

            return occListModel[i]
        }
    }
}

function getObjectOccTarget(cxnOcc, newModel) {
    var targetObjOcc = cxnOcc.getTarget()
    var targetObjOccX = targetObjOcc.X()
    var targetObjOccY = targetObjOcc.Y()
    var targetObjDef = targetObjOcc.ObjDef()

    var occListModel = targetObjDef.OccListInModel(newModel)

    for (var i = 0; i < occListModel.length; i++) {
        var tempX = occListModel[i].X()
        var tempY = occListModel[i].Y()

        if ((tempX == targetObjOccX) && (tempY == targetObjOccY)) {

            return occListModel[i]
        }
    }
}