//------------- Connection type to be replaced --------------
//FIRST ROUND
var g_originalCxnTypeNr = Constants.CT_DECD_ON; //Is responsible for Role-Function
//var g_originalCxnTypeNr = Constants.CT_DECID_ON; //Is responsible for Role-Function

//SECOND ROUND
//var g_originalCxnTypeNr = Constants.CT_IS_TECH_RESP_3; //Is responsible for Role-Function
//var g_originalCxnTypeNr = Constants.CT_IS_TECH_RESP_1; //Is responsible for Role-Function

//---------------Replacing connection type ------------------
//FIRST ROUND
var g_newCxnTypeNr = Constants.CT_CONTR_TO_2 //Function reads Business Term
//var g_newCxnTypeNr = Constants.CT_CONTR_TO_1 //Function reads Business Term

//SECOND ROUND
//var g_newCxnTypeNr = Constants.CT_DECD_ON //Function reads Business Term
//var g_newCxnTypeNr = Constants.CT_DECID_ON //Function reads Business Term


var g_invertConnection = false;//true if old connection type is in the opposite direction than the new

//var g_newCxnTypeNr = Constants.CT_HAS_CONSLT_ROLE_IN_2 //Has consulting role in Group-Function
//var g_newCxnTypeNr = Constants.CT_EXEC_1 //Carries out Group-Function


//Global variables
var g_nLoc = Context.getSelectedLanguage();
var g_exc = Context.createExcelWorkbook(Context.getSelectedFile());
var g_table = new Array();
var g_headers = new Array("Model", "Location", "Source objects", "Target objects", "New cxn created", "Old cxn deleted");
var g_itemType = Constants.CID_OBJDEF;



function main(){
    
    var currentDB = ArisData.getSelectedDatabases()[0]   
    var oSelModels = currentDB.Find(Constants.SEARCH_MODEL, [Constants.MT_FUNC_ALLOC_DGM, Constants.MT_EEPC_COLUMN])

    
    //var oSelModels = ArisData.getSelectedModels();
    //var currentDB = oSelModels[0].Database();
    currentDB.setAutoTouch(false);
    var oCurrFilter = ArisData.ActiveFilter();
    var cxnTypeName = oCurrFilter.ActiveCxnTypeName(g_originalCxnTypeNr);
    
    //
    for (var i=0; i<oSelModels.length; i++){
        Context.writeStatus("Model " + (i + 1) + "/" + (oSelModels.length))
        var oCurrModel = oSelModels[i];
        var oCurrModelName = oCurrModel.Name(g_nLoc);
        var oCurrModelPath = oCurrModel.Group().Path(g_nLoc);
        var oOrigCxnOccs = oCurrModel.CxnOccListFilter(g_originalCxnTypeNr);
        
        for (var j=0; j<oOrigCxnOccs.length; j++){
            //Collect information of the original object
            var oOrigCxnOcc = oOrigCxnOccs[j];
            var oPointList = oOrigCxnOcc.getPoints();
            var oOrigCxn = oOrigCxnOcc.getDefinition();
            var oSourceOcc = oOrigCxnOcc.getSource();
            var oSourceOccName = oSourceOcc.ObjDef().Name(g_nLoc);
            var oTargetOcc = oOrigCxnOcc.getTarget();
            var oTargetOccName = oTargetOcc.ObjDef().Name(g_nLoc);
            
            //Create a new connection with the same information
            var oNewCxnOcc;
            try{
                if (!g_invertConnection){
                    oNewCxnOcc = oCurrModel.CreateCxnOcc(true, oSourceOcc, oTargetOcc, g_newCxnTypeNr, oPointList);
                } else {
                    oNewCxnOcc = oCurrModel.CreateCxnOcc(true, oTargetOcc, oSourceOcc, g_newCxnTypeNr, oPointList);
                }
            } catch(e){
                //Model is locked
                g_table.push([oCurrModelName, oCurrModelPath, oSourceOccName, oTargetOccName, "Model locked", false]);
                continue;
            }
            var bCreateSuccess = false;
            if (oNewCxnOcc != null){
             bCreateSuccess = oNewCxnOcc.IsValid();
            }
            if (!bCreateSuccess){
                g_table.push([oCurrModelName, oCurrModelPath, oSourceOccName, oTargetOccName, false, false]);
                continue;
            }
            //oNewCxnOcc.applyTemplate();
            //copy attributes from original to new cxn
            //oNewCxn = oNewCxnOcc.getDefinition();
            //Dialogs.MsgBox(oSourceOccName + " "+j);
            //Delete original cxnOcc (reorganization will take care of the def)
            
            var bDeleteSuccess = oOrigCxnOcc.Remove();
            
            g_table.push([oCurrModelName, oCurrModelPath, oSourceOccName, oTargetOccName, bCreateSuccess, bDeleteSuccess]);
            
        }
    }
    
    currentDB.setAutoTouch(true);
    printInformation(cxnTypeName);
    g_exc.write();
    
}

main();

function printInformation(p_sheetName){
    
    var headerFont = g_exc.createFont();
    headerFont.setFontName("headerFont");
    headerFont.setBoldweight(Constants.XL_FONT_WEIGHT_BOLD);
    var headerStyle = g_exc.createCellStyle (headerFont, Constants.BORDER_NONE, Constants.BORDER_NONE, Constants.BORDER_NONE, Constants.BORDER_NONE, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, Constants.C_TRANSPARENT, Constants.C_GREY_80_PERCENT, Constants.FILLTYPE_TRANSPARENT);
    //createCellStyle ( XlsFont p_Font, short p_TopBorder, short p_RightBorder, short p_BottomBorder, short p_LeftBorder, int p_TopBorderColor, int p_RightBorderColor, int p_BottomBorderColor, int p_LeftBorderColor, short p_HorizontalAlignment, short p_VerticalAlignment, int p_BackgroundColor, int p_ForegroundColor, short p_FillPattern )
    var sObjects = g_exc.createSheet(p_sheetName);
    
    var headRow = sObjects.createRow(0);
    for (var i=0; i<g_headers.length; i++){
        headRow.createCell(i).setCellValue(g_headers[i]);
    }
    
    
    for (var i=0;i<g_table.length;i++){
        var excRow = sObjects.createRow(i+1);
        for (var j=0;j<g_table[i].length;j++){
            writeExcelCell(excRow, j, j);
            
        }
    }
    
    //format work sheet
    sObjects.setDefaultColumnWidth(40);
    //sObjects.setColumnWidth(GUIDColumn,50*256); //changing columnwidth for selected columns
    //sObjects.setColumnWidth(2,25*256);
    sObjects.setZoom(3,4);
    sObjects.createFreezePane(1,1);
    
    //bold first row
    var headerCells = sObjects.getRowAt(0).getCells();
    for (var i=0;i<headerCells.length;i++){
        headerCells[i].setCellStyle(headerStyle);
    }
    
    
    //Wrap text and set vertical alignment
    var contentCell = sObjects.getCell(1,0);
    if (contentCell != null){
        contentCell.getCellStyle().setWrapText(true);
        contentCell.getCellStyle().setVerticalAlignment(Constants.VERTICAL_TOP);
        contentCell.getCellStyle().setLocked(false);
    }
    sObjects.getCell(0,0).getCellStyle().setWrapText(true);
    sObjects.getCell(0,0).getCellStyle().setLocked(true);
    
    
    
    //g_exc.setWriteProtect(true,"pwd","user");
    //sObjects.setProtect(true);
    
    
}
/*--------------------------------------------
Help functions
------------------------------------------*/
function writeExcelCell(p_row, p_tableIndex, p_columnIndex){
    var rowIndex = p_row.getRowIndex();
    var excCell = p_row.createCell(p_columnIndex);
    var cellData = "";
    if (isArray(g_table[rowIndex-1][p_tableIndex])) {
        for (i1=0;i1<g_table[rowIndex-1][p_tableIndex].length;i1++){
            cellData = cellData + g_table[rowIndex-1][p_tableIndex][i1] + "; " + "\n";
        }
        var nr = cellData.length;
        cellData = cellData.slice(0, nr-3);
        
    } else {
        cellData = g_table[rowIndex-1][p_tableIndex];
    }
    excCell.setCellValue(cellData);
}


function isArray(testObject) {   
    return testObject && !(testObject.propertyIsEnumerable('length')) && typeof testObject === 'object' && typeof testObject.length === 'number';
}

