var excelObject = Context.createExcelWorkbook(Context.getSelectedFile());
var nLocale = Context.getSelectedLanguage()

//setup excell file
//*************************************************************
var sheetName = "Sheet1"
var mainSheet = excelObject.createSheet(sheetName);
//mainSheet.setColumnWidth(0, 12000);
mainSheet.setDefaultColumnWidth(22)
var rowCounter = 0;
createHeaderRow(mainSheet)
//*************************************************************

var variantParent = "000"
var modelGuid = "000"
var modelObjectsStart = ArisData.getSelectedObjOccs()
if (modelObjectsStart.length == 0) {

    var selectedModels = ArisData.getSelectedModels()
    var selectedModel = selectedModels[0]

    var modelName = selectedModel.Name(nLocale)
    var modelGuid = selectedModel.GUID()
    var hierarchyLevel = "L0" //first model always L0
    var description = selectedModel.Attribute(Constants.AT_DESC, nLocale).getValue()

    var newRow = mainSheet.createRow(rowCounter);
    rowCounter++

    var newCell = newRow.createCell(0);
    newCell.setCellValue(modelName);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(1);
    newCell.setCellValue(modelGuid);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(2);
    newCell.setCellValue(hierarchyLevel);
    //newCell.setCellStyle(cellStyle);

    var newCell = newRow.createCell(3);
    newCell.setCellValue(description);
    //newCell.setCellStyle(cellStyle);

    var objOccListTemp = selectedModel.ObjOccListFilter(Constants.OT_FUNC)

    var modelObjectsStart = new Array()
    for (var i = 0; i < objOccListTemp.length; i++) {
        var objDefTemp = objOccListTemp[i].ObjDef()
        //check Hierarchy Level
        if (objDefTemp.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("49e3ac80-b16f-11e2-424c-005056aa0ab6"), nLocale).IsMaintained()) {
            modelObjectsStart.push(objOccListTemp[i])
        }
    }
}

for (var ix = 0; ix < modelObjectsStart.length; ix++) {
    getDataHier(modelObjectsStart[ix], modelGuid, variantParent, 0)
}

excelObject.write()

function getDataHier(objOcc, parent, variantParent, variantDepth) {
    var objDef = objOcc.ObjDef()

    var id = objDef.GUID()
    var o_name = objDef.Name(nLocale)
    var description = objDef.Attribute(Constants.AT_DESC, nLocale).getValue()
    var hierarchyLevelAtt = objDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("49e3ac80-b16f-11e2-424c-005056aa0ab6"), nLocale)
    if (hierarchyLevelAtt.IsMaintained()) {
        var hierarchyLevel = objDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("49e3ac80-b16f-11e2-424c-005056aa0ab6"), nLocale).getValue()
    }
    else {
        var hierarchyLevel = ""
    }

    var occModel = objOcc.Model()
    var modelTypeNum = occModel.TypeNum()

    if (modelTypeNum == Constants.MT_FUNC_TREE || modelTypeNum == ArisData.ActiveFilter().UserDefinedModelTypeNum("2972efc1-c981-11e4-1c0c-0050560116c6")) {

        var newRow = mainSheet.createRow(rowCounter);
        rowCounter++

        var newCell = newRow.createCell(0);
        newCell.setCellValue(o_name);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(1);
        newCell.setCellValue(id);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(2);
        newCell.setCellValue(hierarchyLevel);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(3);
        newCell.setCellValue(description);
        //newCell.setCellStyle(cellStyle);
    }

    //if variant
    if (modelTypeNum == Constants.MT_PRCS_SLCT_MTX) {
        variantDepth = variantDepth + 1

        var newRow = mainSheet.createRow(rowCounter);
        rowCounter++

        var newCell = newRow.createCell(0);
        newCell.setCellValue(o_name);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(1);
        newCell.setCellValue(id);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(2);
        newCell.setCellValue(hierarchyLevel);
        //newCell.setCellStyle(cellStyle);

        var newCell = newRow.createCell(3);
        newCell.setCellValue(description);
        //newCell.setCellStyle(cellStyle);
    }

    //get recursion objects
    var assignedModels = objDef.AssignedModels([Constants.MT_FUNC_TREE, Constants.MT_PRCS_SLCT_MTX])
    var funcObjOccs = new Array()
    for (var i = 0; i < assignedModels.length; i++) {
        var tempObjOccs = assignedModels[i].ObjOccListFilter(Constants.OT_FUNC)
        for (var j = 0; j < tempObjOccs.length; j++) {
            var tempObjDef = tempObjOccs[j].ObjDef()
            if (tempObjDef.GUID() == objDef.GUID()) {
                var connectedOccs = tempObjOccs[j].getConnectedObjOccs(0)
                for (k in connectedOccs) {
                    //return true if assigment is in loop
                    if(!checkLoop(connectedOccs[k], occModel)){
                        funcObjOccs.push(connectedOccs[k])
                    }
                }
            }
        }
    }

    //recursion
    for (var i = 0; i < funcObjOccs.length; i++) {
        getDataHier(funcObjOccs[i], id, parent, variantDepth)
    }
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
    newCell.setCellValue("Process name");
    newCell.setCellStyle(cellStyle1);

    var newCell = headerRow.createCell(1);
    newCell.setCellValue("GUID");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(2);
    newCell.setCellValue("Hierarchy level");
    newCell.setCellStyle(cellStyle2);

    var newCell = headerRow.createCell(3);
    newCell.setCellValue("Description");
    newCell.setCellStyle(cellStyle2);

    //var newCell = headerRow.createCell(8);
    //newCell.setCellValue("Archiving assessment date");
    //newCell.setCellStyle(cellStyle3);
}

function RGB(r, g, b) {
    return (new java.awt.Color(r / 255.0, g / 255.0, b / 255.0, 1)).getRGB() & 0xFFFFFF
}

function checkLoop(targetObjOcc, occModel){
//check if the assigment is in loop, if yes return true
var flag = false
var assignedModelsTemp = targetObjOcc.ObjDef().AssignedModels([Constants.MT_FUNC_TREE, Constants.MT_PRCS_SLCT_MTX])
for (var i=0; i<assignedModelsTemp.length; i++){
    if(assignedModelsTemp[i].GUID() == occModel.GUID()){
        var flag = true
    }
}  
return flag
}