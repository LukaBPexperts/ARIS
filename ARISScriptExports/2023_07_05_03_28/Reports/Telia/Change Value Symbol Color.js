var nLocale = Context.getSelectedLanguage()
var selectedDatabases = ArisData.getSelectedDatabases()
var selectedDatabase = selectedDatabases[0]
selectedDatabase.setAutoTouch(false);

//log file global variables
var g_exc = Context.createExcelWorkbook(Context.getSelectedFile());
var g_table_Log = new Array();
var g_headers_Log = new Array("Model name", "Object Occurence Name", "Status");

var allProcessChainModels = selectedDatabase.Find(Constants.SEARCH_MODEL, ArisData.ActiveFilter().UserDefinedModelTypeNum("8397e581-6239-11eb-306f-0050568286bc")) //Process chain

for (var i = 0; i < allProcessChainModels.length; i++) {
    var processChain = allProcessChainModels[i]

    if (processChain.canWrite(true)) {
        var valueObjOccs = processChain.ObjOccListBySymbol([ArisData.ActiveFilter().UserDefinedSymbolTypeNum("22b1a871-3492-11eb-306f-0050568286bc")])
        for (var j = 0; j < valueObjOccs.length; j++) {
            var objOcc = valueObjOccs[j]
            var setColorFlag = objOcc.setFillColor(RGB(92, 82, 78)) //color hex code #5C524E
            if (setColorFlag) {
                g_table_Log.push([processChain.Name(nLocale), objOcc.ObjDef().Name(nLocale), "Color changed"]);
            }
            else {
                g_table_Log.push([processChain.Name(nLocale), objOcc.ObjDef().Name(nLocale), "ERROR - Color is not changed"]);
            }
        }
    }
    else {
        g_table_Log.push([processChain.Name(nLocale), "", "ERROR - Model is locked!"]);
    }
}

selectedDatabase.setAutoTouch(true);

//PRINT LOG FILE
printInformation_Log("LOG");

//write file
g_exc.write();

function RGB(r, g, b) {
    return (new java.awt.Color(r / 255.0, g / 255.0, b / 255.0, 1)).getRGB() & 0xFFFFFF
}

function writeExcelCell(p_row, p_tableIndex, p_columnIndex, g_table) {
    var rowIndex = p_row.getRowIndex();
    var excCell = p_row.createCell(p_columnIndex);
    var cellData = "";
    if (isArray(g_table[rowIndex - 1][p_tableIndex])) {
        for (i1 = 0; i1 < g_table[rowIndex - 1][p_tableIndex].length; i1++) {
            cellData = cellData + g_table[rowIndex - 1][p_tableIndex][i1] + "; " + "\n";
        }
        var nr = cellData.length;
        cellData = cellData.slice(0, nr - 3);

    }
    else {
        cellData = g_table[rowIndex - 1][p_tableIndex];
    }
    excCell.setCellValue(cellData);
}

function isArray(testObject) {
    return testObject && !(testObject.propertyIsEnumerable('length')) && typeof testObject === 'object' && typeof testObject.length === 'number';
}

function printInformation_Log(p_sheetName) {
    var headerFont = g_exc.createFont();
    headerFont.setFontName("headerFont");
    headerFont.setBoldweight(Constants.XL_FONT_WEIGHT_BOLD);
    var headerStyle = g_exc.createCellStyle(headerFont, Constants.BORDER_NONE, Constants.BORDER_NONE, Constants.BORDER_NONE, Constants.BORDER_NONE, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, Constants.C_TRANSPARENT, Constants.C_GREY_80_PERCENT, Constants.FILLTYPE_TRANSPARENT);
    //createCellStyle ( XlsFont p_Font, short p_TopBorder, short p_RightBorder, short p_BottomBorder, short p_LeftBorder, int p_TopBorderColor, int p_RightBorderColor, int p_BottomBorderColor, int p_LeftBorderColor, short p_HorizontalAlignment, short p_VerticalAlignment, int p_BackgroundColor, int p_ForegroundColor, short p_FillPattern )
    var sObjects = g_exc.createSheet(p_sheetName);

    var headRow = sObjects.createRow(0);
    for (var i = 0; i < g_headers_Log.length; i++) {
        headRow.createCell(i).setCellValue(g_headers_Log[i]);
    }

    for (var i = 0; i < g_table_Log.length; i++) {
        Context.writeStatus("1. Print Groups" + (i + 1) + "/" + g_table_Log.length)
        var excRow = sObjects.createRow(i + 1);
        for (var j = 0; j < g_table_Log[i].length; j++) {
            writeExcelCell(excRow, j, j, g_table_Log);
        }
    }

    //format work sheet
    sObjects.setDefaultColumnWidth(40);
    //sObjects.setColumnWidth(GUIDColumn,50*256); //changing columnwidth for selected columns
    //sObjects.setColumnWidth(2,25*256);
    sObjects.setZoom(3, 4);
    sObjects.createFreezePane(1, 1);

    //bold first row
    var headerCells = sObjects.getRowAt(0).getCells();
    for (var i = 0; i < headerCells.length; i++) {
        headerCells[i].setCellStyle(headerStyle);
    }

    //Wrap text and set vertical alignment
    var contentCell = sObjects.getCell(1, 0);
    if (contentCell != null) {
        contentCell.getCellStyle().setWrapText(true);
        contentCell.getCellStyle().setVerticalAlignment(Constants.VERTICAL_TOP);
        contentCell.getCellStyle().setLocked(false);
    }
    sObjects.getCell(0, 0).getCellStyle().setWrapText(true);
    sObjects.getCell(0, 0).getCellStyle().setLocked(true);

    //g_exc.setWriteProtect(true,"pwd","user");
    //sObjects.setProtect(true);
}