var excelObject = Context.createExcelWorkbook(Context.getSelectedFile());
var nLocale = Context.getSelectedLanguage()
var selectedObjDefs = ArisData.getSelectedObjDefs()

for (var i = 0; i < selectedObjDefs.length; i++) {

    var sourceObjName = selectedObjDefs[i].Name(nLocale)
    var sourceObjType = selectedObjDefs[i].Type()
    var sourceObjGUID = selectedObjDefs[i].GUID()

    //setup excell file
    //*************************************************************
    var sheetName = sourceObjName
    var mainSheet = excelObject.createSheet(sheetName);
    //mainSheet.setColumnWidth(0, 12000);
    mainSheet.setDefaultColumnWidth(20)
    var rowCounter = 0;

    createHeaderRow(mainSheet)
    //*************************************************************

    var connectedObjDefs = selectedObjDefs[i].getConnectedObjs([Constants.OT_APPL_SYS_TYPE, Constants.OT_APPL_SYS_CLS], Constants.EDGES_INOUT, [Constants.CT_SENDS_3])

    for (var j = 0; j < connectedObjDefs.length; j++) {

        var targetObjName = connectedObjDefs[j].Name(nLocale)
        var targetObjType = connectedObjDefs[j].Type()
        var targetObjGUID = connectedObjDefs[j].GUID()

        cxnsArray = getCxnArray(selectedObjDefs[i], connectedObjDefs[j])
        
        connectedObjDefs[j] = null;
        
        for (var k = 0; k < cxnsArray.length; k++) {
            var cxnType = cxnsArray[k].ActiveType()
            var cxnGUID = cxnsArray[k].GUID()
            var cxnAttRole = cxnsArray[k].Attribute(Constants.AT_CXN_ROLE, nLocale).getValue()
            var cxnAttDescription = cxnsArray[k].Attribute(Constants.AT_DESC, nLocale).getValue()
            var cxnAttRemarks = cxnsArray[k].Attribute(Constants.AT_REM, nLocale).getValue()
            var cxnOccs = cxnsArray[k].OccList()
            
            cxnsArray[k] = null;

            var newRow = mainSheet.createRow(rowCounter);
            rowCounter++

            var newCell = newRow.createCell(0);
            newCell.setCellValue(sourceObjName);
            //newCell.setCellStyle(cellStyle);

            var newCell = newRow.createCell(1);
            newCell.setCellValue(sourceObjType);
            //newCell.setCellStyle(cellStyle);

            var newCell = newRow.createCell(2);
            newCell.setCellValue(sourceObjGUID);
            //newCell.setCellStyle(cellStyle);

            var newCell = newRow.createCell(3);
            newCell.setCellValue(cxnType);
            //newCell.setCellStyle(cellStyle);

            var newCell = newRow.createCell(4);
            newCell.setCellValue(cxnGUID);
            //newCell.setCellStyle(cellStyle);

            var newCell = newRow.createCell(5);
            newCell.setCellValue(cxnAttRole);
            //newCell.setCellStyle(cellStyle);

            var newCell = newRow.createCell(6);
            newCell.setCellValue(cxnAttDescription);
            //newCell.setCellStyle(cellStyle);

            var newCell = newRow.createCell(7);
            newCell.setCellValue(cxnAttRemarks);
            //newCell.setCellStyle(cellStyle);

            var newCell = newRow.createCell(8);
            newCell.setCellValue(cxnOccs.length);
            //newCell.setCellStyle(cellStyle);

            var newCell = newRow.createCell(9);
            newCell.setCellValue(targetObjType);
            //newCell.setCellStyle(cellStyle);

            var newCell = newRow.createCell(10);
            newCell.setCellValue(targetObjName);
            //newCell.setCellStyle(cellStyle);

            var newCell = newRow.createCell(11);
            newCell.setCellValue(targetObjGUID);
            //newCell.setCellStyle(cellStyle);

            var newCell = newRow.createCell(12);
            newCell.setCellValue("NO");
            //newCell.setCellStyle(cellStyle);
            
        }
    }
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
    newCell.setCellValue("Source object name");
    newCell.setCellStyle(cellStyleObject);

    var newCell = headerRow.createCell(1);
    newCell.setCellValue("Source object type");
    newCell.setCellStyle(cellStyleObject);

    var newCell = headerRow.createCell(2);
    newCell.setCellValue("Source object GUID");
    newCell.setCellStyle(cellStyleObject);

    var newCell = headerRow.createCell(3);
    newCell.setCellValue("Connection type");
    newCell.setCellStyle(cellStyleCxn);

    var newCell = headerRow.createCell(4);
    newCell.setCellValue("Relationship GUID");
    newCell.setCellStyle(cellStyleCxn);

    var newCell = headerRow.createCell(5);
    newCell.setCellValue("Attribute: Connection Role");
    newCell.setCellStyle(cellStyleCxn);

    var newCell = headerRow.createCell(6);
    newCell.setCellValue("Attribute: Description");
    newCell.setCellStyle(cellStyleCxn);

    var newCell = headerRow.createCell(7);
    newCell.setCellValue("Attribute: Remarks");
    newCell.setCellStyle(cellStyleCxn);

    var newCell = headerRow.createCell(8);
    newCell.setCellValue("# of occurence for this connection");
    newCell.setCellStyle(cellStyleCxn);

    var newCell = headerRow.createCell(9);
    newCell.setCellValue("Target object type");
    newCell.setCellStyle(cellStyleObject);

    var newCell = headerRow.createCell(10);
    newCell.setCellValue("Target object name");
    newCell.setCellStyle(cellStyleObject);

    var newCell = headerRow.createCell(11);
    newCell.setCellValue("Target object GUID");
    newCell.setCellStyle(cellStyleObject);

    var newCell = headerRow.createCell(12);
    newCell.setCellValue("To be deleted");
    newCell.setCellStyle(cellStyleMark);
}

function RGB(r, g, b) {
    return (new java.awt.Color(r / 255.0, g / 255.0, b / 255.0, 1)).getRGB() & 0xFFFFFF
}