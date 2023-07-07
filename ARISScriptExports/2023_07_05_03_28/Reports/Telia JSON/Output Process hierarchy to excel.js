var activeFilter = ArisData.ActiveFilter();

var g_nLoc = Context.getSelectedLanguage();
var g_env   = Context.getEnvironment();                         // Determines the script runtime environment {BP,STD}
var g_exc = Context.createExcelWorkbook(Context.getSelectedFile());
var g_resultTable = new Array();
var g_entireMethodGUID = "dd838074-ac29-11d4-85b8-00005a4053ff";
var g_publish = false;


var g_versioning;
var g_modelList = new Array();
var g_table = new Array();
var g_headers = new Array("Level (calculated)", "L0 Process location", "L1 Process location", "L2 Process location", "L3 Process location", "L4 Process step", "Name","Level", "GUID", "Group Path","Type", "Last user", "Last changed", "Process Owner", "Description", "Link to Development");
var g_publishedOrNot = "";
var g_currentModelPath = [];
var g_versionComments = "";
var g_assignmentLevelLimit = 0; //Number of assignment levels considered, 0=only selected models are evaluated

var aLastUser = Constants.AT_LUSER;
var aDescription = Constants.AT_DESC;
var aLastChanged = Constants.AT_LAST_CHNG_2;
var aCreator = Constants.AT_CREATOR;
var aTimeOfCreation = Constants.AT_CREAT_TIME_STMP;
var aModelStatus = ArisData.ActiveFilter().UserDefinedAttributeTypeNum("f02e3820-948c-11e3-01db-b51dcc951f78");
var aIdentifier = Constants.AT_ID;
var aPersonResponsible = Constants.AT_PERS_RESP;
var aProcessOwner = ArisData.ActiveFilter().UserDefinedAttributeTypeNum("e70ccb60-d01c-11dd-7315-000c295dfb26");
var aProcessLevel = ArisData.ActiveFilter().UserDefinedAttributeTypeNum("a908e230-d8dc-11de-7aef-000c295dfb26");

var g_attributes = [aLastUser, aLastChanged, aProcessOwner, aDescription];


function main(){
    var oStartingModels = ArisData.getSelectedModels();
    if (oStartingModels.length == 0){
        var oSelectedGroup = ArisData.getSelectedGroups()[0];
        oStartingModels = oSelectedGroup.ModelList(true);
    }
    var contextDB = oStartingModels[0].Database();
    var currentDateTime = new Date().toISOString();
    
    /* ------------------------------------------------
    Dialog for selecting number of assignment levels
    ------------------------------------------------ */
    var result = Dialogs.showDialog(new numberSelectionDialog(), Constants.DIALOG_TYPE_ACTION, "Select number of Assignment levels");
    if (result == -1){
        Context.setScriptError(Constants.ERR_CANCEL);
        return;
    }else{
        g_assignmentLevelLimit = parseInt(result);
    }
    //Dialogs.MsgBox(g_publish);  
    for (var i=0; i<oStartingModels.length; i++){
        determineStartingModelPath(oStartingModels[i]);
        getNextLevel(oStartingModels[i]);
        //To aggregate information from lower levels, make the getNextLevel function to return that information and store it here
    }
    
    //Recursive function reverses table. Reverse back before printing out.
    //g_table.reverse();
    
    printInformation(oStartingModels[0].Name(g_nLoc));
}
main();
function determineStartingModelPath(p_model){
    var procMapGuids = ["fa3b3d81-2b82-11e9-0473-00505699050f","a8187441-061b-11ea-5df4-00505699050f"];
    var allGuids = [];
    var oCurrentModel = p_model;
    for (var i=0; i<3; i++){ //Don't try further than 3 levels up
        if (existsInList(procMapGuids,oCurrentModel.GUID())){
            return;
        }
        allGuids.push(oCurrentModel.GUID());
        g_currentModelPath.unshift(oCurrentModel.Name(g_nLoc));
        var oParentObjects = oCurrentModel.SuperiorObjDefs();
        for (var j=0; j<oParentObjects.length; j++){
            var oParentObject = oParentObjects[j];
            var oParentOccs = oParentObject.OccList();
            for (var k=0; k<oParentOccs.length; k++){
                var oParentOcc = oParentOccs[k];
                if (oParentOcc.getSymbol() == Constants.ST_VAL_ADD_CHN_SML_1){
                    var oParentModel = oParentOcc.Model();
                    //Create a tobe variant of the parent model, If parent already exists, it will end there
                    if (oParentModel.TypeNum() == Constants.MT_VAL_ADD_CHN_DGM && !existsInList(allGuids, oParentModel.GUID())){//To exclude e.g. matrix models
                        oCurrentModel = oParentModel;
                    }
                }
            }
        }
        
    }
}

function existsInList(p_list, p_value){
    for (var i=0; i<p_list.length; i++){
        if (p_list[i] == p_value){
            return true;
        }
    }
    return false;
}
        
function numberSelectionDialog()
{
    var oResult = 0;
    // all member functions except for getPages can access the property "dialog" of the dialog class. Type is "UserDialog" (see help).
    // examples:
    // - Get the page with the specified index (e.g. 0): this.dialog.getPage(0)
    // - Define the list of visible wizard pages (e.g. 0,1): this.dialog.setActiveWizardPages([0,1])
    
    // returns DialogTemplate[]
    // non-optional
    this.getPages = function()
    {
        var oDialogTemplate = Dialogs.createNewDialogTemplate(600, 200, "Select Number");
        oDialogTemplate.Text(10, 10, 300, 40, "Select number of assignment levels to be included.\n0 = selected model(s) only");
        oDialogTemplate.ComboBox(10, 40, 100, 40, [0,1,2,3,4,5,6,7,8], "ASSIGNMENTS");
        return [oDialogTemplate];
    }
    
    // returns true if the page is in a valid state. In this case "Ok", "Finish", or "Next" is enabled.
    // called each time a dialog value is changed by the user (button pressed, list selection, text field value, table entry, radio button,...)
    // pageNumber: the current page number, 0-based
    this.isInValidState = function(pageNumber)
    {
        return true;
    }
    
    // called after "Ok"/"Finish" has been pressed and the current state data has been applied
    // can be used to update your data
    // pageNumber: the current page number
    // bOK: true=Ok/finish, false=cancel pressed
    // optional
    this.onClose = function(pageNumber, bOk)
    {
        if (!bOk){
            this.oResult = -1;
        }
    }
    
    // the result of this function is returned as result of Dialogs.showDialog(). Can be any object.
    // optional
    this.getResult = function()
    {
        if (this.oResult == -1){
            return this.oResult;
        } else {
            return this.dialog.getPage(0).getDialogElement("ASSIGNMENTS").getValue();
        }
    }
    
}

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
    //Check that model can be published
    var aModelName = p_model.Name(g_nLoc);
    var aModelType = p_model.Type();
    var aGroupPath = p_model.Group().Path(g_nLoc);
    var sProcessLevel = p_model.Attribute(aProcessLevel, g_nLoc).getValue();
    var oSuperiorObjects = p_model.SuperiorObjDefs();
    /* var aSuperiorIdentifier = "";
    for (var i=0; i<oSuperiorObjects.length; i++){
        var aIdentifier = oSuperiorObjects[i].Attribute(Constants.AT_ID, g_nLoc).getValue();
        if (aIdentifier != ""){
            aSuperiorIdentifier = aIdentifier;
            break;
        }
    } */
    
    
    
    
    /*  ----------------------
    Get next level models
    ---------------------- */
    //Get included objects
    
    //var oIncludedObjects = p_model.ObjDefListFilter(); //Alternatively e.g. ObjDefListByTypes(int[] objTypeNum), For EPC models need to filter out Process interfaces
    var oIncludedObjects = new Array();
    var oIncludedOccs = p_model.ObjOccListBySymbol([Constants.ST_SOLAR_VAC, Constants.ST_VAL_ADD_CHN_SML_1, Constants.ST_VAL_ADD_CHN_SML_2, Constants.ST_FUNC, Constants.ST_SOLAR_FUNC, Constants.ST_CONTR]);//Constants.ST_SOLAR_SL_VAC_OCC excluded
    for (var i=0; i<oIncludedOccs.length; i++){
        oIncludedObjects.push(oIncludedOccs[i].ObjDef());
    } 
    //Get assigned models
    var oAssignedModels = new Array();
    for (var i=0; i<oIncludedObjects.length; i++){
        var oCurrAssignedModels = oIncludedObjects[i].AssignedModels([Constants.MT_VAL_ADD_CHN_DGM, Constants.MT_EEPC, Constants.MT_EEPC_ROW]); //Possible to filter by model type or array of model types
        for (var j=0; j<oCurrAssignedModels.length; j++){
            if (!oCurrAssignedModels[j].equals(p_model)){//Do not include assignment to self. E.g. in VACD or a model that is already included
                oAssignedModels.push(oCurrAssignedModels[j]);
            }
        }
    }
    var oModelData = [];
    oModelData.push(g_currentModelPath.length-1);
    var tempModelPath = g_currentModelPath.concat([aModelName]);
    for (var i=0; i<5; i++){
        var levelModel= tempModelPath[i+1]; //exclude Process Map - model
        if (levelModel == null){
            levelModel = "";
        }
        oModelData.push(levelModel);
    }
    oModelData.push(aModelName);
    oModelData.push(sProcessLevel);
    oModelData.push(p_model.GUID());
    
    oModelData.push(p_model.Group().Path(g_nLoc));
    oModelData.push(aModelType);
    for (var i=0; i<g_attributes.length; i++){
        oModelData.push(p_model.Attribute(g_attributes[i], g_nLoc).getValue());
    }
    oModelData.push(getConnectLink(p_model));
    
    g_table.push(oModelData);
    
    //Put model to end of current breadcrumb path
    g_currentModelPath.push(aModelName);
    
    //if model is EPC, include each process step
    if (p_model.TypeNum() == Constants.MT_EEPC){
        for (var i=0; i<oIncludedObjects.length; i++){
            var oFunc = oIncludedObjects[i];
            var oObjectData = [];
            oObjectData.push(g_currentModelPath.length-1);
            
            //oObjectData.push(g_currentModelPath.concat([oFunc.Name(g_nLoc)]));
            var tempModelPath = g_currentModelPath.concat([oFunc.Name(g_nLoc)]);
            for (var j=0; j<5; j++){
                var levelModel= tempModelPath[j+1]; //exclude Process Map - model
                if (levelModel == null){
                    levelModel = "";
                }
                oObjectData.push(levelModel);
            }
            oObjectData.push(oFunc.Name(g_nLoc));
            oObjectData.push("Process step (L4)");
            oObjectData.push(oFunc.GUID());
            oObjectData.push(oFunc.Group().Path(g_nLoc));
            oObjectData.push(oFunc.Type());
            for (var j=0; j<g_attributes.length; j++){
                oObjectData.push(oFunc.Attribute(g_attributes[j], g_nLoc).getValue());
            }
            g_table.push(oObjectData);
        }
    }
    
    //Sort models per wanted parameters
    oAssignedModels = ArisData.sort(oAssignedModels, Constants.AT_NAME, g_nLoc);
    for (var i=0; i<oAssignedModels.length; i++){
        getNextLevel(oAssignedModels[i]); //Recursive call
    }
    /* -------------------------
    Actions for current model (in reversed order)
    ------------------------- */
    
    
    //This model has been handled, remove from end of current breadcrumb path
    g_currentModelPath.splice(-1);
}

function getConnectLink(p_model){
    var aModelGUID = p_model.GUID();
    /* var oModelInPublishing = g_publishDB.FindGUID(aModelGUID, Constants.CID_MODEL);
    if (!oModelInPublishing.IsValid()){
        return "";
    } */
    
    var oCompressedGUID = compressGuid(aModelGUID);
    var aLink = "http://aris.metsagroup.com/#default/item/c.temp.UNITED.[COMPRESSEDGUID]";
    //aLink = aLink.replace("[TYPE]", sPortalType);
    aLink = aLink.replace("[COMPRESSEDGUID]", oCompressedGUID);
    return aLink;
}

function compressGuid(guid) {
    var code = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_";
    var compressedGuid = "";
    var bucket = 0;
    var bits = 0;
    for (var i = 0; i < guid.length(); i++) {
 
        var rValue = guid.charAt(i)
        if (rValue >= 48 && rValue <=57)
            fValue = rValue - 48;
        else if (rValue >= 97 && rValue <= 102)
            fValue = 10 + rValue - 97;
        else if (rValue >= 65 && rValue <= 90)
            fValue = 10 + rValue - 65;
        else             
            continue;
 
       bucket = (bucket << 4) + fValue;
        bits+=4;
        if (bits >= 6) {
            bits-=6;
            sValue = bucket >> bits;
            bucket = bucket ^ (sValue << bits);
            compressedGuid += code.charAt(sValue);
        }
    }
    if (bits > 0) {
        compressedGuid += code.charAt(bucket << (6 - bits));
    }
    return compressedGuid;
}


/* -------------------------
    Print to Excel
    ------------------------- */
function printInformation(p_sheetName){
    
    var headerFont = g_exc.createFont();
    headerFont.setFontName("headerFont");
    headerFont.setBoldweight(Constants.XL_FONT_WEIGHT_BOLD);
    var headerStyle = g_exc.createCellStyle (headerFont, Constants.BORDER_NONE, Constants.BORDER_NONE, Constants.BORDER_NONE, Constants.BORDER_NONE, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, Constants.C_TRANSPARENT, Constants.C_GREY_80_PERCENT, Constants.FILLTYPE_TRANSPARENT);
    var normalFont = g_exc.createFont();
    var normalStyle = g_exc.createCellStyle (normalFont, Constants.BORDER_NONE, Constants.BORDER_NONE, Constants.BORDER_NONE, Constants.BORDER_NONE, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.FILLTYPE_TRANSPARENT);
    
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
    sObjects.setDefaultColumnWidth(30);
    sObjects.setColumnWidth(0,0*256);
    sObjects.setColumnWidth(1,20*256);
    sObjects.setColumnWidth(6,50*256);
    sObjects.setColumnWidth(7,30*256);
    sObjects.setColumnWidth(8,0*256);
    sObjects.setColumnWidth(9,0*256);
    sObjects.setColumnWidth(10,20*256);
    sObjects.setColumnWidth(11,20*256);
    sObjects.setColumnWidth(13,20*256);
    sObjects.setColumnWidth(14,70*256);
    
    sObjects.setZoom(3,4);
    sObjects.createFreezePane(6,1);
    
    //bold Header row
    var headerCells = sObjects.getRowAt(0).getCells();
    for (var i=0;i<headerCells.length;i++){
        headerCells[i].setCellStyle(headerStyle);
    }
    
    
    //Wrap text and set vertical alignment
    var contentCell = sObjects.getCell(2,0);
    if (contentCell != null){
        contentCell.getCellStyle().setWrapText(true);
        contentCell.getCellStyle().setVerticalAlignment(Constants.VERTICAL_CENTER);
        contentCell.getCellStyle().setLocked(false);
    }
    sObjects.getCell(1,0).getCellStyle().setWrapText(true);
    //sObjects.getCell(0,0).getCellStyle().setLocked(true);
    
    //mergeSameCells(sObjects, 0);
    //mergeSameCells(sObjects, 1);
    //mergeSameCells(sObjects, 2);
            
    //g_exc.setWriteProtect(true,"pwd","user");
    //sObjects.setProtect(true);
    g_exc.write();
    
}
/*--------------------------------------------
Help functions
------------------------------------------*/
function existsInArray(p_obj, p_objArray){
    var exist = false;
    for (i=0; i<p_objArray.length; i++){
        if (p_objArray[i].equals(p_obj)){
            exist = true;
        }
    }
    return exist;
}

function mergeSameCells(p_sheet, p_column, p_style){
    //Merge cells with same value
    var lastRowNum = p_sheet.getLastFilledRowNum();
    var lastValue = "";
    var mergeStartRow = 0;
    var mergeEndRow = 0;
    for (var i=0; i<lastRowNum+1; i++){
        var oCell = p_sheet.getCell(i, p_column);
        if (oCell == null){//there is no value
            if (mergeEndRow > mergeStartRow){
                p_sheet.addRegion(p_column, p_column, mergeStartRow, mergeEndRow);
            }
            mergeStartRow = i;
            mergeEndRow = i;
            continue;
        }    
        var cellValue = oCell.getCellValue();
        //Check if there is a change in the previous column
        var changeAtLeftCell = false; 
        if (p_column > 0 && i > 0){//don not check first column or first row
            var oLeftCell = p_sheet.getCell(i, p_column-1);
            var oPrevLeftCell = p_sheet.getCell(i-1, p_column-1);
            if (oLeftCell != null && oPrevLeftCell != null){
                var oLeftValue = oLeftCell.getCellValue();
                var oPrevLeftValue = oPrevLeftCell.getCellValue();
                if (!oLeftValue.equals(oPrevLeftValue)){
                    changeAtLeftCell = true;
                }
            } else {//either cell is null, merge
                changeAtLeftCell = true;
            }
        }
        if (cellValue.equals(lastValue) && !changeAtLeftCell){//there is no change at this row or the previous
            mergeEndRow = i;
            if (i==lastRowNum && mergeEndRow > mergeStartRow){ //End of table, merge if area longer than 1
                p_sheet.addRegion(p_column, p_column, mergeStartRow, lastRowNum)
            }
        } else {
            
            if (mergeEndRow > mergeStartRow){//If only one cell in area, do not merge
                p_sheet.addRegion(p_column, p_column, mergeStartRow, mergeEndRow);
            }
            mergeStartRow = i;
            mergeEndRow = i;
        }
        lastValue = cellValue;
    } 
}

function writeExcelCell(p_row, p_tableIndex, p_columnIndex){
    var rowIndex = p_row.getRowIndex();
    var excCell = p_row.createCell(p_columnIndex);
    var cellData = "";
    if (isArray(g_table[rowIndex-1][p_tableIndex])) {
        for (i1=0;i1<g_table[rowIndex-1][p_tableIndex].length;i1++){
            cellData = cellData + g_table[rowIndex-1][p_tableIndex][i1] + "; ";
        }
        var nr = cellData.length;
        cellData = cellData.slice(0, nr-2);
        
    } else {
        cellData = g_table[rowIndex-1][p_tableIndex];
    }
    if (cellData.length > 4){
        if (cellData.slice(0,4) == "http"){
            excCell.setHyperlink(cellData, cellData);
        } else {
            excCell.setCellValue(cellData);
        }
    }
    excCell.setCellValue(cellData);
    
    
}


function isArray(testObject) {   
    return testObject && !(testObject.propertyIsEnumerable('length')) && typeof testObject === 'object' && typeof testObject.length === 'number';
}



function getAttrValue(p_oItem, p_nAttrType, p_bRemoveLF) {
    if (p_oItem != null) {
        var oAttr = p_oItem.Attribute(p_nAttrType, g_nLoc);
        var oAttrBaseType = ArisData.ActiveFilter().AttrBaseType(p_nAttrType);
        if(oAttrBaseType == 9 && oAttr != null && oAttr.IsMaintained()) {        //Date
            var d = oAttr.MeasureValue();
            var oDate = d.getDate();
            var oMonth = d.getMonth();
            oMonth++;
            var oYear = d.getYear();
            oYear = oYear + 1900;   //getFullYear() is not working
            return "" + oDate + "." + oMonth + "." + oYear;
        } else if(oAttr != null && oAttr.IsMaintained()) {
            return "" + oAttr.GetValue(p_bRemoveLF);
        }
    }
    return "";
}

