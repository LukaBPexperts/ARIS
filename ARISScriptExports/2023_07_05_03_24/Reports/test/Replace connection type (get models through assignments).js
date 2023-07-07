
//------------- Connection type to be replaced --------------
//var g_originalCxnTypeNr = Constants.CT_CONTR_TO_2; //Contributes to Role-Function
//var g_originalCxnTypeNr = Constants.CT_DECD_ON; //Is responsible for Role-Function
var g_originalCxnTypeNr = Constants.CT_IS_INP_FOR//Business Term Is input for Function

//var g_originalCxnTypeNr = Constants.CT_CONTR_TO_1; //Contributes to Group-Function
//var g_originalCxnTypeNr = Constants.CT_DECID_ON; //Is responsible for Group function

//---------------Replacing connection type ------------------
//var g_newCxnTypeNr = Constants.CT_HAS_CONSLT_ROLE_IN_2 //Has consulting role in Role-Function
//var g_newCxnTypeNr = Constants.CT_EXEC_2 //Carries out Role-Function
var g_newCxnTypeNr = Constants.CT_READ_1 //Function reads Business Term

//var g_newCxnTypeNr = Constants.CT_HAS_CONSLT_ROLE_IN_2 //Has consulting role in Group-Function
//var g_newCxnTypeNr = Constants.CT_EXEC_1 //Carries out Group-Function


var g_invertConnection = true;//true if old connection type is in the opposite direction than the new
var g_modifiedModelType = Constants.MT_FUNC_ALLOC_DGM; //Model type number in which connection types are to be changed
var g_assignmentLevelLimit = 6; //Number of assignment levels to be checked, 0=selected model only 


//Global variables
var g_modelList = new Array();
var g_currentModelPath = new Array();
var g_nLoc = Context.getSelectedLanguage();
var g_targetModels = new Array();
var g_exc = Context.createExcelWorkbook(Context.getSelectedFile());
var g_table = new Array();
var g_table2 = new Array();
var g_headers = new Array("Model name", "Model Type","Level", "Assignment path", "Group path (location)");
var g_headers2 = new Array("Model", "Location", "Source objects", "Target objects", "New cxn created", "Old cxn deleted");
var g_itemType = Constants.CID_OBJDEF;



function main(){
    var oSelModels = ArisData.getSelectedModels();
    var currentDB = oSelModels[0].Database();
    currentDB.setAutoTouch(false);
    var oCurrFilter = ArisData.ActiveFilter();
    var cxnTypeName = oCurrFilter.ActiveCxnTypeName(g_originalCxnTypeNr);
    for (var i=0; i<oSelModels.length; i++){
        
        getNextLevel(oSelModels[i]); //Recursive function to traverse assignment hierarchy. Models of the defined type are stored in g_targetModels
        //To aggregate information from lower levels, make the getNextLevel function to return that information and store it here
    }
    //Recursive function reverses table. Reverse back before printing out.
    g_table.reverse();
    
    printInformation("Found models",g_headers,g_table);
    for (var i=0; i<g_targetModels.length; i++){
        var oCurrModel = g_targetModels[i];
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
                    oPointList.reverse();
                    oNewCxnOcc = oCurrModel.CreateCxnOcc(true, oTargetOcc, oSourceOcc, g_newCxnTypeNr, oPointList);
                }
            } catch(e){
                //Model is locked
                g_table2.push([oCurrModelName, oCurrModelPath, oSourceOccName, oTargetOccName, "Model locked", false]);
                continue;
            }
            var bCreateSuccess = false;
            if (oNewCxnOcc != null){
             bCreateSuccess = oNewCxnOcc.IsValid();
            }
            if (!bCreateSuccess){
                g_table2.push([oCurrModelName, oCurrModelPath, oSourceOccName, oTargetOccName, false, false]);
                continue;
            }
            oNewCxnOcc.applyTemplate();
            //copy attributes from original to new cxn
            //oNewCxn = oNewCxnOcc.getDefinition();
            //Dialogs.MsgBox(oSourceOccName + " "+j);
            //Delete original cxnOcc (reorganization will take care of the def)
            
            var bDeleteSuccess = oOrigCxnOcc.Remove();
            
            g_table2.push([oCurrModelName, oCurrModelPath, oSourceOccName, oTargetOccName, bCreateSuccess, bDeleteSuccess]);
            
        }
    }
    
    currentDB.setAutoTouch(true);
    printInformation(cxnTypeName,g_headers2,g_table2);
    g_exc.write();
    
}

main();

function getNextLevel(p_model){
    if (g_currentModelPath.length == g_assignmentLevelLimit+1){ //Level limit is achieved. Ignore this model
        return;
    }
    if (existsInArray(p_model, g_modelList)){
        return;
    }
    g_modelList.push(p_model);
    /* -------------------------
    Actions for current model
    ------------------------- */
    var tableRow = new Array(); //Array with the following information
    //Get properties
    var aModelName = p_model.Name(g_nLoc);
    var aModelType = p_model.Type();
    var aGroupPath = p_model.Group().Path(g_nLoc);
   
    //Add to g_targetModels if criteria is fulfilled
    if (p_model.TypeNum() == g_modifiedModelType){
        g_targetModels.push(p_model);
    }
    //Put model to end of current breadcrumb path
    
    g_currentModelPath.push(aModelName);
    
    /*  ----------------------
    Get next level models (not if the model is a FAD)
    ---------------------- */
    //Get included objects
    if (p_model.TypeNum() != Constants.MT_FUNC_ALLOC_DGM){
        var oIncludedObjects = p_model.ObjDefListFilter(); //Alternatively e.g. ObjDefListByTypes(int[] objTypeNum), For EPC models need to filter out Process interfaces
        
        //Get assigned models
        var oAssignedModels = new Array();
        for (var i=0; i<oIncludedObjects.length; i++){
            var oCurrAssignedModels = oIncludedObjects[i].AssignedModels([Constants.MT_FUNC_TREE, Constants.MT_PRCS_SLCT_MTX, Constants.MT_FUNC_ALLOC_DGM]); //Possible to filter by model type or array of model types
            for (var j=0; j<oCurrAssignedModels.length; j++){
                if (!oCurrAssignedModels[j].equals(p_model)){//Do not include assignment to self. E.g. in VACD or a model that is already included
                    oAssignedModels.push(oCurrAssignedModels[j]);
                }
            }
        }
        //Sort models per wanted parameters
        oAssignedModels = ArisData.sort(oAssignedModels, Constants.AT_NAME, g_nLoc);
        for (var i=0; i<oAssignedModels.length; i++){
            getNextLevel(oAssignedModels[i]); //Recursive call
        }
    }
    
    
    
    /* -------------------------
    Actions for current model (in reversed order)
    ------------------------- */
    g_table.push([aModelName, aModelType, g_currentModelPath.length, g_currentModelPath.join(), aGroupPath]);
    //This model has been handled, remove from end of current breadcrumb path
    g_currentModelPath.splice(-1);
}

function printInformation(p_sheetName,p_headers,p_table){
    
    var headerFont = g_exc.createFont();
    headerFont.setFontName("headerFont");
    headerFont.setBoldweight(Constants.XL_FONT_WEIGHT_BOLD);
    var headerStyle = g_exc.createCellStyle (headerFont, Constants.BORDER_NONE, Constants.BORDER_NONE, Constants.BORDER_NONE, Constants.BORDER_NONE, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, Constants.C_TRANSPARENT, Constants.C_GREY_80_PERCENT, Constants.FILLTYPE_TRANSPARENT);
    //createCellStyle ( XlsFont p_Font, short p_TopBorder, short p_RightBorder, short p_BottomBorder, short p_LeftBorder, int p_TopBorderColor, int p_RightBorderColor, int p_BottomBorderColor, int p_LeftBorderColor, short p_HorizontalAlignment, short p_VerticalAlignment, int p_BackgroundColor, int p_ForegroundColor, short p_FillPattern )
    var sObjects = g_exc.createSheet(p_sheetName);
    
    var headRow = sObjects.createRow(0);
    for (var i=0; i<p_headers.length; i++){
        headRow.createCell(i).setCellValue(p_headers[i]);
    }
    
    
    for (var i=0;i<p_table.length;i++){
        var excRow = sObjects.createRow(i+1);
        for (var j=0;j<p_table[i].length;j++){
            writeExcelCell(p_table,excRow, j, j);
            
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
        //contentCell.getCellStyle().setWrapText(true);
        contentCell.getCellStyle().setVerticalAlignment(Constants.VERTICAL_TOP);
        contentCell.getCellStyle().setLocked(false);
    }
    //sObjects.getCell(0,0).getCellStyle().setWrapText(true);
    sObjects.getCell(0,0).getCellStyle().setLocked(true);
    
    
    
    //g_exc.setWriteProtect(true,"pwd","user");
    //sObjects.setProtect(true);
    
    
}
/*--------------------------------------------
Help functions
------------------------------------------*/
function writeExcelCell(p_table,p_row, p_tableIndex, p_columnIndex){
    var rowIndex = p_row.getRowIndex();
    var excCell = p_row.createCell(p_columnIndex);
    var cellData = "";
    if (isArray(p_table[rowIndex-1][p_tableIndex])) {
        for (i1=0;i1<p_table[rowIndex-1][p_tableIndex].length;i1++){
            cellData = cellData + p_table[rowIndex-1][p_tableIndex][i1] + "; " + "\n";
        }
        var nr = cellData.length;
        cellData = cellData.slice(0, nr-3);
        
    } else {
        cellData = p_table[rowIndex-1][p_tableIndex];
    }
    excCell.setCellValue(cellData);
}


function isArray(testObject) {   
    return testObject && !(testObject.propertyIsEnumerable('length')) && typeof testObject === 'object' && typeof testObject.length === 'number';
}

function existsInArray(p_obj, p_objArray){
    var exist = false;
    for (i=0; i<p_objArray.length; i++){
        if (p_objArray[i].equals(p_obj)){
            exist = true;
        }
    }
    return exist;
}