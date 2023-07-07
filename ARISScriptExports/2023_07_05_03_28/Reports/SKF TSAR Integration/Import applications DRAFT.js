/* -------ARIS - Capability mappings----------------------------------------
  A table for mapping excel header to ARIS attribute
  To add your own mapping, add a row with the following format:
g_mapTable.push([Name in excel header, ARIS Attribute type number or GUID]);         //Attribute type name in ARIS

  The order of mappings is not meaningful, i.e does not have to match position in Excel file
----------------------------------------------------------------------- */
var g_mapTable = new Array();
g_mapTable.push(["System ID", "AT_ID"]);    //PPM ID
g_mapTable.push(["System Name", Constants.AT_NAME]);                           //Name
g_mapTable.push(["Description", Constants.AT_DESC]);                    //Description
//g_mapTable.push(["Global / Local", "d8a12551-2688-11ec-4841-0a5f6f9452ec"]);                    //Description
//g_mapTable.push(["Title 4", Constants.AT_TITL4]);         
//g_mapTable.push(["Link 4", Constants.AT_LINK]);         

/* ---------------Special columns-----------------------------------------------
Columns that need more functionality than just writing excel cell value to corresponding ARIS attribute.
----------------------------------------------------------------------------- */
var g_applicationID = "System ID";               //system ID column header in import file, used for matching system between ARIS and PPM

//Only update existing or also create new
var g_updateOnly = false; //If set to true it will only update existing applications. When set to false will create new objects

/* ----------General settings-------------------------------------
 Definitions for settings needed by the script.
 E.g. which attribute is used for indicating import status
--------------------------------------------------------------- */                                                 
var g_applicationGroupPath = "SKF/test";                       //Group path for objects
var g_applicationModelGroupPath = g_applicationGroupPath;      //Group path for system library model
var g_applicationObjType = Constants.OT_APPL_SYS_CLS;                                   //system object type
var g_applicationObjSymbolType = "12e88be1-e5cb-11e5-6fc6-005056a95aef";                             //IT system symbol type
//Application	Application system class	262158	12e88be1-e5cb-11e5-6fc6-005056a95aef	Derived	

var g_applicationModelName = "Applications map";                                    //Name of system library model
var g_applicationModelType = "2c1d6381-d494-11e5-6fc6-005056a95aef" //Constants.MT_APPL_SYS_TYPE_DGM;              //system landscape model


var g_sourceObjType = Constants.OT_ORG_UNIT;                                        //Object type for 
var g_cxnModelGroupPath = g_applicationGroupPath;
var g_cxnModelType = Constants.MT_ACS_DGM; 
var g_cxnModelName = "Systems and organizations relation";
var g_applicationOrgMatrixGroupPath = g_applicationGroupPath;                                 //Group path for Systems-Organizations Matrix
var g_applicationOrgMatrixName = "Systems and organizations relation";                   //Name of Systems-Orgs Matrix
var g_applicationOrgCxnType = Constants.CT_CAN_BE_USER;                                  //Systems-Orgs relationship type (Org - can be user - System)

/* ------Settings needed for running script by scheduler-----------------

------------------------------------------------------------- */
var isRunScheduled = false;
var g_clientLogFileName = "importLog.txt";
//Outfile
var g_nLoc = Context.getSelectedLanguage();
var g_oOutfile = Context.createOutputObject(Context.getSelectedFormat(), Context.getSelectedFile());
g_oOutfile.Init(g_nLoc);
g_oOutfile.DefineF("Normal", "Times New Roman", 12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0, 21, 0, 0, 0, 1);

var g_nEng = 1033;                                                                  //Language used for all attributes - English (United states)
var g_currentDB = ArisData.getActiveDatabase();                                     //Current ARIS database (on which the report is run)
var g_currentFilter = ArisData.ActiveFilter();                                      //Current filter

var s_importGroup = pathToGroup(g_applicationGroupPath);
var s_libraryModelGroup = pathToGroup(g_applicationModelGroupPath);
var s_libraryModelType = toModelTypeNr(g_applicationModelType);

var g_applicationObjSymbolType = toSymbolTypeNr(g_applicationObjSymbolType)

//var s_cxnModelGroup = pathToGroup(g_cxnModelGroupPath);
//var s_cxnModelType = toModelTypeNr(g_cxnModelType);
var s_sourceObjects = new Array();
var s_targetObjects = new Array();
var nrOfUpdated = 0;
var nrOfCreated = 0;

/* -------------
Functions for writing to Excel
------------- */
var g_exc = Context.createExcelWorkbook(Context.getSelectedFile());
var g_table = [];
var g_headers = ["Update status"]; 
var g_columnWidths = [];
//g_columnWidths.push([0, 13]);

var g_colorGrey = getColorByRGB( 54, 53, 52);
var g_colorLighterGrey = getColorByRGB( 94, 93, 92);
var g_colorLightestGrey = getColorByRGB( 244, 243, 242);
var g_colorGreen = getColorByRGB( 143, 212, 9);

function main(){
        
    //Check that groups in configuration exist in database
    if(!checkGroupExistence([s_importGroup], g_oOutfile)){
        logOutput("All defined groups are not existing", g_oOutfile);
        printLogOutput(g_oOutfile);
        return; //All groups are not ok
    } 
        
    /* Read excel file and create table for object creation*/
    
    var selFile = Dialogs.getFilePath("","*.xlsx!!Worksheet Files|*.xlsx;*.xls||","","Select import file",0)[0].getData();
    
    try{
        var workbookReader = Context.getExcelReader(selFile);
    } catch (e){
        logOutput("No valid file has been selected for data import.", g_oOutfile);
        var workbookReader = null;
    }
    if (workbookReader != null) {
        var sheets = workbookReader.getSheets();
        //Check mapTable
        /* if(!checkMapTable(sheets[0], g_mapTable, g_oOutfile)){
            printLogOutput(g_oOutfile);
            return; //map Table not ok
        } */
        
        //Find or create libraryModel. Interrupt if libraryModel is locked
        //var libraryModel = findOrCreateLibraryModel(s_libraryModelGroup, s_libraryModelType, g_applicationModelName, g_oOutfile);
        //if (!libraryModel.IsValid()){
        //    printLogOutput(g_oOutfile);
        //    return; //Librarymodel locked or not created for some reason
        //}
                
        readImportExcel(sheets[0]);
        //updateLibraryModel(libraryModel, s_importGroup, g_applicationObjType);
        
    } else {
        //Dialogs.MsgBox("No file has been selected for data import.", 0, "Aris Report");
        Context.setScriptError(Constants.ERR_CANCEL);
    }
    printInformation(g_table, g_headers, "Imported Applications", g_columnWidths);
g_exc.write();
    
}

main();
function getOriginalCellValue(p_cell){
    var isDate = p_cell.isCellADate();
    if (isDate){
        var oCellValue = p_cell.getDateCellValue();
    } else if (p_cell.getNumericCellValue() > 0) {
        var oCellValue = p_cell.getNumericCellValue();
    } else {
        var oCellValue = p_cell.getCellValue().trim();
    }
    return oCellValue;
}
function readImportExcel(sheet){
    var IDColumn = 0;
    var parentIdColumn = 100;
    var itemTypeColumn = 100;
    var IDArisTypeNr = mapHeaderToTypeNr(g_applicationID, g_mapTable);
    var attrTypeList = new Array();
    var attrDateList = new Array();
    
    
    
    //read header row to get attributes and IDColumn       
    var row = 0;
    var nrOfColumns = sheet.getRowAt(0).getLastCellNum();
    for (var column = 0; column < nrOfColumns; column++){
        var columnHeader = getOriginalCellValue(sheet.getCell(row, column));
        if (columnHeader == g_applicationID){ //Mark ID column
            IDColumn = column;
        
        }
        for (var i = 0; i < g_mapTable.length; i++){
            if(columnHeader == g_mapTable[i][0]){
                g_headers.push(g_mapTable[i][0]); //also writes header
            }
        }
        
        var headerTypeNr = mapHeaderToTypeNr(columnHeader, g_mapTable);//also writes output header
        
        //get columns with date attributes
        if (headerTypeNr == 0){ //attribute not mapped
            attrDateList.push(0);
        } else {
            if (g_currentFilter.AttrBaseType(headerTypeNr) == Constants.ABT_DATE){
                attrDateList.push(1);
            } else {
                attrDateList.push(0);
            }
        }
        attrTypeList.push(headerTypeNr);
    }
    
    //read through Excel file
    for (row = 1; row < sheet.getLastFilledRowNum()+1; row++){
        var rowData = [];
        var oObject;
        var oCellType = sheet.getCell(row, IDColumn).getCellType().getType();
        if (oCellType == Constants.XL_CELL_TYPE_NUMERIC){
            var objID = sheet.getCell(row, IDColumn).getNumericCellValue();
        }
        var objID = sheet.getCell(row, IDColumn).getCellValue();
        var numericValue = parseInt(objID);
        if (typeof numericValue === "number"){
            objID = numericValue.toString();
        }
        var objTypeNumber = g_applicationObjType;
        var oObjectSymbolType = g_applicationObjSymbolType;
        var newItem = false;

        var searchItem = g_currentDB.createSearchItem(IDArisTypeNr, g_nEng, objID, Constants.SEARCH_CMP_EQUAL, false, false);
        var objCandidates = s_importGroup.ObjDefList(true, objTypeNumber, searchItem);
        if (objCandidates.length > 0){
            oObject = objCandidates[0];
            rowData.push("No change");
        }else{
            if (g_updateOnly){
                continue; //Does not create new objects
            }
            oObject = s_importGroup.CreateObjDef(objTypeNumber, "N/A", g_nEng);
            oObject.setDefaultSymbolNum(oObjectSymbolType, false);
            rowData.push("New object");
            newItem = true;
        }
    
        
        //for each column update attribute value
        for (var column = 0; column < attrTypeList.length ; column++){
            if (attrTypeList[column] != 0){ //only import attributes defined in mapping table
                var oCurrentAttribute = oObject.Attribute(attrTypeList[column], g_nEng);
                var oCurrentValue = oCurrentAttribute.getValue();
                var selCell = sheet.getCell(row, column);
                var cellValue = "";
                if (selCell == null && oCurrentValue != "") { //empty cell
                    oCurrentAttribute.Delete();
                    if (!newItem){
                    rowData[0] = "Updated";
                    }
                } else if(attrDateList[column] == 1){ //ARIS attribute base type is "Date"
                    cellValue = selCell.getDateCellValue();
                    
                    try{
                        oCurrentAttribute.setValue(cellValue);
                        if(!newItem){
                        rowData[0] = "Updated";
                        }
                        
                    }catch (err){
                        //Cell has date formatting but no value
                    }
                } else {
                    cellValue = getOriginalCellValue(selCell);
                    if (cellValue != oCurrentValue){ //Column corresponds to normal attribute and has changed
                        oCurrentAttribute.setValue(cellValue);
                        if(!newItem){
                        rowData[0] = "Updated";
                        }
                    }
                    
                }
                rowData.push(cellValue);
            }
        }
        g_table.push(rowData);
        s_targetObjects.push(oObject); 
        
        
               
    }
    
}


function checkMapTable(p_sheet, p_mapTable, p_oOutfile){
    var mapTableOk = true;
    var missingHeaders = new Array();
    var nonExistingAttributes = new Array();
    var nrOfColumns = p_sheet.getRowAt(0).getLastCellNum();
    for (var i=0; i<p_mapTable.length; i++){
        var mapTableHeaderFound = false;
        var mapTableHeader = p_mapTable[i][0];
        for (var column=0; column<nrOfColumns; column++){
            if (mapTableHeader == p_sheet.getCell(0, column).getCellValue()){
                mapTableHeaderFound = true;
                //check that attribute exists in method
                if(toTypeNr(p_mapTable[i][1]) == -1){
                    nonExistingAttributes.push(p_mapTable[i][1]);
                    nonExistingAttributes.push("(" + mapTableHeader + ")");
                }
            }
        }
        if(!mapTableHeaderFound){
            missingHeaders.push(mapTableHeader);
        }
    }
    if (missingHeaders.length > 0){
        logOutput("Following headers are missing in import file: " + missingHeaders.toString(), p_oOutfile);
        mapTableOk = false;
    }
    if (nonExistingAttributes.length > 0){
        logOutput("Following GUIDS are not found in ARIS method: " + nonExistingAttributes.toString(), p_oOutfile);
        mapTableOk = false;
    }
    return mapTableOk;
}

function toTypeNr(p_GUIDorTypeNr){
    var typeNr = -1; 
    if(typeof p_GUIDorTypeNr === "string"){
        try{
            typeNr = g_currentFilter.UserDefinedAttributeTypeNum(p_GUIDorTypeNr);
        } catch (e){
        }
    }else{
        typeNr = p_GUIDorTypeNr;
    }
    return typeNr;
}

function toSymbolTypeNr(p_GUIDorTypeNr){
    var typeNr = -1; 
    if(typeof p_GUIDorTypeNr === "string"){
        typeNr = g_currentFilter.UserDefinedSymbolTypeNum(p_GUIDorTypeNr); 
    }else{
        typeNr = p_GUIDorTypeNr;
    }
    return typeNr;
}

function toModelTypeNr(p_GUIDorTypeNr){
    var typeNr = -1; 
    if(typeof p_GUIDorTypeNr === "string"){
        typeNr = g_currentFilter.UserDefinedModelTypeNum(p_GUIDorTypeNr); 
    }else{
        typeNr = p_GUIDorTypeNr;
    }
    return typeNr;
}

function checkGroupExistence(p_groupArray, p_oOutfile){
    var groupsExist = true;
    var nonExistingGroups = new Array();
    for (var i=0; i<p_groupArray.length; i++){
        if(typeof p_groupArray[i] === "string"){ //String instead of a group object
            nonExistingGroups.push(p_groupArray[i]);
        }
    }
    if (nonExistingGroups.length > 0){
        logOutput("Following group paths are not found in the database: " + nonExistingGroups.toString(), p_oOutfile);
        groupsExist = false;
    }
    return groupsExist;
}

function pathToGroup(p_groupPath){
    var oGroup = g_currentDB.Group(p_groupPath, g_nEng);
    if(!oGroup.IsValid()){
       oGroup = p_groupPath;
    }
    return oGroup;
}

function findOrCreateLibraryModel(p_modelGroup, p_modelType, p_modelName, p_oOutfile){
    var libraryModel;
    var searchItem = g_currentDB.createSearchItem(Constants.AT_NAME, g_nEng, p_modelName, Constants.SEARCH_CMP_EQUAL, false, false);
    var candidateModels = p_modelGroup.ModelList(false, [p_modelType], searchItem);
    if (candidateModels.length > 0){
        libraryModel = candidateModels[0];
        var lockingComponent = Context.getComponent("Locking");
        var oLockOwner = lockingComponent.getLockOwner(libraryModel);
        if (oLockOwner != ""){
           logOutput("Library model is locked by " + oLockOwner + ". Import is cancelled.", p_oOutfile);
           libraryModel = null;
        }
    }else{
        libraryModel = p_modelGroup.CreateModel(p_modelType, p_modelName, g_nEng);
    }
    return libraryModel;
}

function updateLibraryModel(p_libraryModel, p_objGroup, p_objType){
    var xInit = 50;
    var xStep = 500;
    var xMax = 4000;
    var yInit = 50;
    var yStep = 250;
    var xPos = xInit;
    var yPos = yInit;
    
    //Find last occurrence position in libraryModel
    var allOccs = p_libraryModel.ObjOccList();
    var testOccs = new Array();
    for (var i=0; i<allOccs.length; i++){
        var testOcc = allOccs[i];
        var testY = testOcc.Y();
        if (testY > yPos){
            yPos = testY;
            testOccs = [testOcc];
        } else if (testY == yPos){
            testOccs.push(testOcc);
        }
    }
    for (var i=0; i<testOccs.length; i++){
        var testX = testOccs[i].X()
        if (testX > xPos){
            xPos = testX;
        }
    }
    //from import group, find all objects not in library model and add them there
    var allObjects = p_objGroup.ObjDefListFilter(p_objType);
    allObjects = ArisData.sort(allObjects, Constants.AT_NAME, g_nEng);
    for (var i=0; i<allObjects.length; i++){
        var oObjectDef = allObjects[i];
        if(oObjectDef.OccListInModel(p_libraryModel).length == 0){
            var defSymbolType = oObjectDef.getDefaultSymbolNum();
            xPos = xPos + xStep;
            if(xPos > xMax){
                xPos = xInit;
                yPos = yPos + yStep;
            }
            var newOcc = p_libraryModel.createObjOcc(defSymbolType, oObjectDef, xInit, yInit, true, true);
            newOcc.SetPosition(xPos, yPos);
        }
    }
}

function mapHeaderToTypeNr(p_header, p_mapTable){
    var attrTypeNr = 0;
    for (var i = 0; i < p_mapTable.length; i++){
        if(p_header == p_mapTable[i][0]){
            attrTypeNr = toTypeNr(p_mapTable[i][1]);
        }
    }
    return attrTypeNr; //Return value 0 = header string not found in mapTable 
}

function initMatrix(p_matrixGroup, p_matrixName, p_oOutfile){
    var oMatrix;
    var oMatrixCandidates = p_matrixGroup.ModelListFilter(p_matrixName, g_nEng, Constants.MT_MATRIX_MOD, Constants.AT_NAME);
    if (oMatrixCandidates.length > 0){
        oMatrix = oMatrixCandidates[0].getMatrixModel();
        var lockingComponent = Context.getComponent("Locking");
        var oLockOwner = lockingComponent.getLockOwner(oMatrix);
        if (oLockOwner != ""){
           logOutput("Matrix model is locked by " + oLockOwner + ". Import is cancelled.", p_oOutfile);
           return null;
        }
        //Empty matrix column and row headers
        var rowHeaders = oMatrix.getHeader(true).getCells();
        for (var i=0; i<rowHeaders.length; i++){
            try{
                oMatrix.deleteHeaderCell(rowHeaders[i]);
            }catch (e){
                 logOutput("Matrix model is locked. Import is cancelled.", p_oOutfile);
                 return null;
            }
        }
        var columnHeaders = oMatrix.getHeader(false).getCells();
        for (var i=0; i<columnHeaders.length; i++){
            oMatrix.deleteHeaderCell(columnHeaders[i]);
        }
    } else {
        oMatrix = p_matrixGroup.CreateModel(Constants.MT_MATRIX_MOD, p_matrixName, g_nEng);
    }
    return oMatrix;
}

function updateCxnLibraryModel(p_cxnLibraryModel, p_sourceObjects, p_targetObjects){
    var fromNoticeTypeNr = toTypeNr(g_fromNoticeAttribute);
    try {
        var sourceObjType = p_sourceObjects[0].TypeNum();
        var targetObjType = p_targetObjects[0].TypeNum();
    } catch (e) {
        logOutput("Either no source or target objects found", g_oOutfile);
        return;
    }
    for (var i=0; i<p_targetObjects.length; i++){
        var oTargetObjDef = p_targetObjects[i];
        var oTargetOccs = oTargetObjDef.OccListInModel(p_cxnLibraryModel);
        var oTargetOcc;
        if (oTargetOccs.length == 0){
            var defSymbolType = oTargetObjDef.getDefaultSymbolNum();
            oTargetOcc = p_cxnLibraryModel.createObjOcc(defSymbolType, oTargetObjDef, 100, 100, true, false);
        } else {
            oTargetOcc = oTargetOccs[0];
        }
        if (!oTargetOcc.IsValid()){
           logOutput("Could not create occurrence of "+ oTargetObjDef.Name(g_nEng), g_oOutfile);
           continue;
        }
        var inCxnDefs = oTargetObjDef.CxnListFilter(Constants.EDGES_IN, g_applicationOrgCxnType);
        for (var j=0; j<inCxnDefs.length; j++){
            var currCxnDef = inCxnDefs[j];
            var isImportedCxn = currCxnDef.Attribute(fromNoticeTypeNr, g_nEng).GetValue(false);
            var cxnOccsInModel = currCxnDef.OccList([p_cxnLibraryModel]);
            if (isImportedCxn && cxnOccsInModel.length == 0){ //create new connection occ
                var oSourceObjDef = currCxnDef.SourceObjDef();
                var oSourceOccs = oSourceObjDef.OccListInModel(p_cxnLibraryModel);
                var oSourceOcc;
                if (oSourceOccs.length == 0){
                    var defSourceSymbolType = oSourceObjDef.getDefaultSymbolNum();
                    oSourceOcc = p_cxnLibraryModel.createObjOcc(defSourceSymbolType, oSourceObjDef, 200, 200, true, false);
                } else {
                    oSourceOcc = oSourceOccs[0];
                }
                if (!oSourceOcc.IsValid()){
                    logOutput("Could not create occurrence of "+ oSourceObjDef.Name(g_nEng), g_oOutfile);
                    continue;
                }
                                
                g_ntranspoint[0] = new PublicInterface.Point(oSourceOcc.X(),oSourceOcc.Y());
                g_ntranspoint[1] = new PublicInterface.Point(oTargetOcc.X(),oTargetOcc.Y());
                p_cxnLibraryModel.CreateCxnOcc(oSourceOcc, oTargetOcc, currCxnDef, g_ntranspoint, false);
            }
        }
    }
    
    //Sort objects
    var xInit = 50;
    var xStep = 500;
    var xMax = 4000;
    var yInit = 50;
    var yStep = 250;
    var xPos = xInit;
    var yPos = yInit;
    var allSourceOccs = p_cxnLibraryModel.ObjOccListFilter(sourceObjType);
    var allTargetOccs = p_cxnLibraryModel.ObjOccListFilter(targetObjType);
    //place sourceOccs
    for (var i=0; i<allSourceOccs.length; i++){
        yPos = yPos + yStep;
        allSourceOccs[i].SetPosition(xPos, yPos);
    }
    //place targetOccs
    xInit = 700;
    xPos = xInit;
    yPos = yInit;
    for (var i=0; i<allTargetOccs.length; i++){
        xPos = xPos + xStep;
        if(xPos > xMax){
            xPos = xInit;
            yPos = yPos + yStep;
        }
        allTargetOccs[i].SetPosition(xPos, yPos);
    }
    
}      

function updateMatrix(p_matrix, p_columnHeaderObjects, p_rowHeaderObjects){
    var oMatrix = p_matrix.getMatrixModel();
    var rowObjects = p_rowHeaderObjects;
    rowObjects = ArisData.sort(rowObjects, Constants.AT_NAME, g_nEng);
    var oRowSymbols = new Array;
    for (var i=0; i<rowObjects.length; i++){
        var oSymbol = rowObjects[i].getDefaultSymbolNum();
        var currHeaderCell;
        try {
            currHeaderCell = oMatrix.createHeaderCell(null, rowObjects[i], -1, -1, true);
            var size = currHeaderCell.getSize();
            var oAttributes = rowObjects[i].AttrList(g_nEng);
        } catch (e) {
            logOutput("Problem putting header "+ rowObjects[i].Name(g_nEng) +" into matrix, trying again", g_oOutfile);
            
        }
       
        if (oRowSymbols.length == 0 || !existsInList(oSymbol, oRowSymbols)){
            oRowSymbols.push(oSymbol);
        }
    }
    var columnObjects = p_columnHeaderObjects;
    columnObjects = ArisData.sort(columnObjects, Constants.AT_NAME, g_nEng);
    var oColumnSymbols = new Array;
    for (var i=0; i<columnObjects.length; i++){
        var oSymbol = columnObjects[i].getDefaultSymbolNum();
        try {
            oMatrix.createHeaderCell(null, columnObjects[i], -1, -1, false);
        } catch (e) {
            logOutput("Problem putting header "+ columnObjects[i].Name(g_nEng) +" into matrix", g_oOutfile);
        }
        if (oColumnSymbols.length == 0 || !existsInList(oSymbol, oColumnSymbols)){
            oColumnSymbols.push(oSymbol);
        }
    }
    //go through matrixheaders and correct headers that went wrong
    var hCells = oMatrix.getHeader(true).getCells();
    for (var i=0; i<hCells.length; i++){
        var hCell = hCells[i];
        try{ //store wrong symbol and try changing it to correct one
            var tempSymbolNum = hCell.getSymbolNum();
            if (oRowSymbols.length == 0 || !existsInList(tempSymbolNum, oRowSymbols)){
                oRowSymbols.push(tempSymbolNum);
            }
            hCell.setSymbolNum(oRowSymbols[0]);
        } catch (e){
        }
    }
    
    //go through matrixheaders and correct headers that went wrong
    var hCells = oMatrix.getHeader(false).getCells();
    for (var i=0; i<hCells.length; i++){
        var hCell = hCells[i];
        try{ //store wrong symbol and try changing it to correct one
            var tempSymbolNum = hCell.getSymbolNum();
            if (oColumnSymbols.length == 0 || !existsInList(tempSymbolNum, oColumnSymbols)){
                oColumnSymbols.push(tempSymbolNum);
            }
            hCell.setSymbolNum(oColumnSymbols[0]);
        } catch (e){
        }
    }
    
    
    
    oMatrix.setVisibleObjectSymbolTypes(oColumnSymbols, false); //Make rowheaders visible
    oMatrix.setVisibleObjectSymbolTypes(oRowSymbols, true); //Make rowheaders visible
    if (oRowSymbols.length > 0 || oColumnSymbols.length > 0){
        var newCxnData = oMatrix.createNewMatrixConnectionDataObject(oColumnSymbols[0], oRowSymbols[0], g_applicationOrgCxnType, null, true, false);
        oMatrix.setCxnData([newCxnData]);
    }
    
}


function existsInList(p_value, p_array){
    var exists = false;
    for (var k=0; k<p_array.length; k++){
        if(p_array[k] == p_value){
            exists = true;
            break;
        }
    }
    return exists;
}

function translateNumberToCategory(p_number, p_catMapTable){
    var category;
    for (var i=0; i<p_catMapTable.length; i++){
        var catUpperLimit = p_catMapTable[i][0];
        if(p_number <= catUpperLimit){
            category = p_catMapTable[i][1];
            break;
        }
    }
    return category;
}
/* -------------------------
    Print to Excel
    ------------------------- */
function printInformation(p_table, p_headers, p_sheetName, p_columnWidths) {

    var headerFont = g_exc.getFont("Arial", 11, Constants.C_BLACK, true, false, false, false, Constants.XL_FONT_SS_NONE);
    var headerStyle = g_exc.createCellStyle(headerFont, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, Constants.C_TRANSPARENT, Constants.C_GREY_80_PERCENT, Constants.FILLTYPE_TRANSPARENT);
    headerStyle.setWrapText(true);
    var normalFont = g_exc.getFont("Calibri", 11, Constants.C_BLACK, false, false, false, false, Constants.XL_FONT_SS_NONE);
    var normalStyle = g_exc.createCellStyle(normalFont, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.BORDER_THIN, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_LEFT, Constants.ALIGN_CENTER, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.FILLTYPE_TRANSPARENT);
    normalStyle.setWrapText(true);
    //createCellStyle ( XlsFont p_Font, short p_TopBorder, short p_RightBorder, short p_BottomBorder, short p_LeftBorder, int p_TopBorderColor, int p_RightBorderColor, int p_BottomBorderColor, int p_LeftBorderColor, short p_HorizontalAlignment, short p_VerticalAlignment, int p_BackgroundColor, int p_ForegroundColor, short p_FillPattern )
    var sObjects = g_exc.createSheet(p_sheetName);

    var headRow = sObjects.createRow(0);
    for (var i = 0; i < p_headers.length; i++) {
        headRow.createCell(i).setCellValue(p_headers[i]);
    }


    for (var i = 0; i < p_table.length; i++) {
        var excRow = sObjects.createRow(i + 1);
        for (var j = 0; j < p_table[i].length; j++) {
            writeExcelCell(p_table, excRow, j, j);
            excRow.getCellAt(j).setCellStyle(normalStyle);
        }
    }

    //format work sheet
    sObjects.setDefaultColumnWidth(20);
    for (var i = 0; i < p_columnWidths.length; i++) {
        sObjects.setColumnWidth(p_columnWidths[i][0], p_columnWidths[i][1] * 256);
    }

    sObjects.setZoom(3, 4);
    sObjects.createFreezePane(0,1);//Header row frozen

    //bold Header row
    var headerCells = sObjects.getRowAt(0).getCells();
    for (var i = 0; i < headerCells.length; i++) {
        headerCells[i].setCellStyle(headerStyle);
    }

    //mergeSameCells(sObjects, 0);
    //mergeSameCells(sObjects, 1);
    //mergeSameCells(sObjects, 2);

    //g_exc.setWriteProtect(true,"pwd","user");
    //sObjects.setProtect(true);


}
/*--------------------------------------------
Help functions
------------------------------------------*/
function mergeSameCells(p_sheet, p_column, p_referenceColumn) {
    //Merge cells with same value
    if (p_referenceColumn == null){
        p_referenceColumn = 0;
    }
    var lastRowNum = p_sheet.getLastFilledRowNum();
    var lastValue = "";
    var mergeStartRow = 0;
    var mergeEndRow = 0;
    for (var i = 0; i < lastRowNum + 1; i++) {
        var oCell = p_sheet.getCell(i, p_column);
        if (oCell == null) {//there is no value
            if (mergeEndRow > mergeStartRow) {
                p_sheet.addRegion(p_column, p_column, mergeStartRow, mergeEndRow);
            }
            mergeStartRow = i;
            mergeEndRow = i;
            continue;
        }
        var cellValue = oCell.getCellValue();
        
        //Check if there is a change in the previous column
        var changeAtLeftCell = false;
        if (p_column > 0 && i > 0) {//don not check first column or first row
            var oLeftCell = p_sheet.getCell(i, p_referenceColumn);
            var oPrevLeftCell = p_sheet.getCell(i - 1, p_referenceColumn);
            if (oLeftCell != null && oPrevLeftCell != null) {
                var oLeftValue = oLeftCell.getCellValue();
                var oPrevLeftValue = oPrevLeftCell.getCellValue();
                if (!oLeftValue.equals(oPrevLeftValue)) {
                    changeAtLeftCell = true;
                }
            } else {//either cell is null, merge
                changeAtLeftCell = true;
            }
        }
        if (cellValue.equals(lastValue) && !changeAtLeftCell) {//there is no change at this row or the previous
            mergeEndRow = i;
            if (i == lastRowNum && mergeEndRow > mergeStartRow) { //End of table, merge if area longer than 1
                p_sheet.addRegion(p_column, p_column, mergeStartRow, lastRowNum)
            }
        } else {

            if (mergeEndRow > mergeStartRow) {//If only one cell in area, do not merge
                p_sheet.addRegion(p_column, p_column, mergeStartRow, mergeEndRow);
            }
            mergeStartRow = i;
            mergeEndRow = i;
        }
        lastValue = cellValue;
    }
}

function changeColumnStyle(p_sheet, p_column, p_style){
    //Merge cells with same value
    var lastRowNum = p_sheet.getLastFilledRowNum();
    var lastValue = "";
    var mergeStartRow = 0;
    var mergeEndRow = 0;
    for (var i=1; i<lastRowNum+1; i++){
        var oCell = p_sheet.getCell(i, p_column);
        if (oCell != null){//there is a value
            oCell.setCellStyle(p_style);
            continue;
        }
    }
}
function writeExcelCell(p_table, p_row, p_tableIndex, p_columnIndex) {
    var rowIndex = p_row.getRowIndex();
    var excCell = p_row.createCell(p_columnIndex);
    var cellData = "";
    if (isArray(p_table[rowIndex - 1][p_tableIndex])) {
        for (i1 = 0; i1 < p_table[rowIndex - 1][p_tableIndex].length; i1++) {
            cellData = cellData + p_table[rowIndex - 1][p_tableIndex][i1] + "; ";
        }
        var nr = cellData.length;
        cellData = cellData.slice(0, nr - 2);

    } else {
        cellData = p_table[rowIndex - 1][p_tableIndex];
    }
    if (cellData.length > 4) {
        if (cellData.slice(0, 4) == "http") {
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
        if (oAttrBaseType == 9 && oAttr != null && oAttr.IsMaintained()) {        //Date
            var d = oAttr.MeasureValue();
            var oDate = d.getDate();
            var oMonth = d.getMonth();
            oMonth++;
            var oYear = d.getYear();
            oYear = oYear + 1900;   //getFullYear() is not working
            return "" + oDate + "." + oMonth + "." + oYear;
        } else if (oAttr != null && oAttr.IsMaintained()) {
            return "" + oAttr.GetValue(p_bRemoveLF);
        }
    }
    return "";
}
