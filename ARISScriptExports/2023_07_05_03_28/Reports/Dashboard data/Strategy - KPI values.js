/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

const outputParameters = new OutputParameters("");
const xOutput = createXmlOutput(outputParameters);

var nLocale = Context.getSelectedLanguage();
var aDatabases = ArisData.getActiveDatabase();
    
main();

/*************************************************************************************************/

//------functions------
function main()
{
    outHeader();
    readDatabase();
    uploadXmlOutputToADS(xOutput, outputParameters);
}

/**
**  Function to read models from the database
**  @return void
**/
function readDatabase()
{
    var aModels = aDatabases.Find(Constants.SEARCH_MODEL, [Constants.MT_KPI_ALLOC_DGM]);
    for (var i in aModels) 
    {
        getArisAttributes(aModels[i]);
    }  
}

/**
**  Function to search for KPI object in the model
**  @return void
**/
function getArisAttributes(oModel)
{
    var allOccs = oModel.ObjOccList();

    for(var i in allOccs) 
    {
        var curOcc = allOccs[i];
        var curObj = curOcc.ObjDef();
        
        if(curObj.TypeNum() == Constants.OT_KPI)
        {
            getKpiValue(curObj, oModel);
        } 
    }
}

/**
**  Function for reading the KPI attributes and output to an XML in ADS
**  @return void
**/
function getKpiValue(oKPI, oModel)
{
    var actVal = getKpiValue(Constants.AT_ACT_VAL);
    var plVal = getKpiValue(Constants.AT_PL_VAL);
    var preVal = getKpiValue(Constants.AT_PREVIOUS_VALUE);
    var tol = getKpiValue(Constants.AT_VAL_TOL);
    
    xOutput.addRow([oKPI.Name(nLocale), oKPI.GUID(),actVal, plVal, preVal, tol, oModel.GUID()]);
    
    function getKpiValue(nAttrType)
    {
         var oAttr = oKPI.Attribute(nAttrType, nLocale);
         if (!oAttr.IsMaintained()) return "";
         return oAttr = oAttr.getValue();
    }
}

/**
**  Function to set the columns of the xml file
**  @return void
**/
function outHeader()
{
     xOutput.setColumns([
                       ["KPI name","text"],
                       ["KPI GUID","text"],
                       ["KPI actual value", "number"],
                       ["KPI plan value", "number"],
                       ["KPI previous value", "number"],
                       ["KPI tol value", "number"],
                       ["Model GUID", "text"]
                       ]);   
}