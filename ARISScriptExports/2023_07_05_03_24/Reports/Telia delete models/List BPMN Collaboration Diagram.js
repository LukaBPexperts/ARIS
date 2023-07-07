var excelObject = Context.createExcelWorkbook(Context.getSelectedFile());
var nLocale = Context.getSelectedLanguage()

var selectedDatabase = ArisData.getSelectedDatabases()


var bpmnModelsAll = selectedDatabase[0].Find(Constants.SEARCH_MODEL, [Constants.MT_BPMN_COLLABORATION_DIAGRAM])
var bpmnModels = new Array()
//get all models which have process module occ in it
for (var i = 0; i < bpmnModelsAll.length; i++) {
	
var occList = bpmnModelsAll[i].ObjOccListBySymbol([ArisData.ActiveFilter().UserDefinedSymbolTypeNum("b12afd21-975f-11e5-7464-0050560116c9")]) //TU STAVI API NAME OD SIMBOLA PRIMJER JE: "ST_BPMN_SUBPROCESS"
var occList2 = bpmnModelsAll[i].ObjOccListBySymbol(Constants.ST_BPMN_SUBPROCESS)
var occList3 = bpmnModelsAll[i].ObjOccListBySymbol(Constants.ST_BPMN_TASK)

if(occList.length > 0  && occList2.length==0  && occList3.length==0){
    //puniNiz
    bpmnModels.push(bpmnModelsAll[i])
}
}

//izbaci duple ako ih ima
bpmnModels = ArisData.Unique(bpmnModels)

//setup excell file
    //*************************************************************
    var sheetName = "Sheet1"
    var mainSheet = excelObject.createSheet(sheetName);
    //mainSheet.setColumnWidth(0, 12000);
    mainSheet.setDefaultColumnWidth(20)
    var rowCounter = 0;

    createHeaderRow(mainSheet)
    //*************************************************************

for (var i = 0; i < bpmnModels.length; i++) {

    
            
            var modelName = bpmnModels[i].Name(nLocale)
            var modelPath = bpmnModels[i].Group().Path(nLocale)
            var modelGuid = bpmnModels[i].GUID()
            var modelCreator = bpmnModels[i].Attribute(Constants.AT_CREATOR,nLocale).getValue()
            
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
            
            var newCell = newRow.createCell(3);
            newCell.setCellValue(modelCreator);
            //newCell.setCellStyle(cellStyle);  
}

excelObject.write()

function getCxnArray(objDefA, objDefB) {
    var cxnArray = new Array()

    var cxnsListA = objDefA.CxnListFilter(Constants.EDGES_OUT, Constants.CT_SENDS_3)

    if (objDefA.IsEqual(objDefB)) {
        for (var k = 0; k < cxnsListA.length; k++) {
            var sourceObj = cxnsListA[k].SourceObjDef()
            var targetObj = cxnsListA[k].TargetObjDef()
            if (sourceObj.IsEqual(targetObj)) {
                cxnArray.push(cxnsListA[k])
                cxnsListA[k] = null;
            }
        }
    } else {
        var cxnsListB = objDefB.CxnListFilter(Constants.EDGES_IN, Constants.CT_SENDS_3)
        var lookup = new Array;
        for (var j in cxnsListB) {
            lookup[cxnsListB[j].GUID()] = cxnsListB[j].GUID();
        }

        for (var i in cxnsListA) {
            if (typeof lookup[cxnsListA[i].GUID()] != 'undefined') {
                cxnArray.push(cxnsListA[i])
                //break;
            }
        }
    }
    cxnsListA = null
    cxnsListB = null

    return cxnArray;
}

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
    
    var newCell = headerRow.createCell(3);
    newCell.setCellValue("Creator");
    newCell.setCellStyle(cellStyleObject);
    }

function RGB(r, g, b) {
    return (new java.awt.Color(r / 255.0, g / 255.0, b / 255.0, 1)).getRGB() & 0xFFFFFF
}