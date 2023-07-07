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
    var aModels = aDatabases.Find(Constants.SEARCH_MODEL, [Constants.MT_CAUSE_EFFECT_DGM]);
    for (var i in aModels) 
    {
        getArisAttributes(aModels[i]);
    }  
}

/**
**  Function to check all connection from the model
**  @return void
**/
function getArisAttributes (oModel)
{ 
    var allOccs = oModel.ObjOccList();

    for(var i in allOccs) 
    {
        var curOcc = allOccs[i];
        var curObj = curOcc.ObjDef(); 
        var allConnections = curOcc.CxnOccList();

        for(var x in allConnections) 
        {
            curConOcc = allConnections[x];
            curConDef = curConOcc.getDefinition();
            curTarget = curConOcc.getTarget();
            curSource = curConOcc.getSource();
            
            if (curConDef.TypeNum() == Constants.CT_CONTAINS_1)
            {
                if(curObj.TypeNum() == Constants.OT_PERSPECT && curTarget.ObjDef().TypeNum() == Constants.OT_OBJECTIVE) 
                {
                    var perspective = curObj;
                    checkAss(curTarget, oModel, perspective);               
                }
            }
        }    
    }
}

/**
**  Function to check assignments
**  @return void
**/
function checkAss(curTarget, oModel, perspective)
{
    var perResp = "";
    var initativ = "";
    var allAssignments = curTarget.ObjDef().AssignedModels();
    if (allAssignments != "")
    {
        for(var i in allAssignments) 
        {
            var curAssignment = allAssignments[i];
            var allCurOccs = curAssignment.ObjOccList();
            for(var i in allCurOccs) 
            {
                var assCurOcc = allCurOccs[i];
                var assCurObj = assCurOcc.ObjDef();
                  
                if(assCurObj.TypeNum() == Constants.OT_ORG_UNIT)
                {
                    var allAssConnections = assCurOcc.CxnOccList();
        
                    for(var x in allAssConnections) 
                    {
                        curAssConOcc = allAssConnections[x];
                        curAssConDef = curAssConOcc.getDefinition();
                        curAssTarget = curAssConOcc.getTarget();
                        curSource = curAssConOcc.getSource();
                        
                        if(assCurObj.TypeNum() == Constants.OT_ORG_UNIT && curAssTarget.ObjDef().TypeNum() == Constants.OT_FUNC_INST)
                        {
                            perResp = assCurObj;
                            initativ = curAssTarget;
                            buildRow(curTarget, oModel, perspective, perResp, initativ);
                        }
                    }
                }
            }
            if(perResp == "" && initativ == "" )
            {
                buildRow(curTarget, oModel, perspective, perResp, initativ);
            }
        }
    }
    else
    {
       buildRow(curTarget, oModel, perspective, perResp, initativ); 
    }
}

function buildRow (curTarget, oModel, perspective, perResp, initativ)
{
    if(perResp != "" && initativ != "" )
    {
        xOutput.addRow([oModel.Name(nLocale),oModel.GUID(),perspective.Name(nLocale), perspective.GUID(),curTarget.ObjDef().Name(nLocale),curTarget.ObjDef().GUID(),initativ.ObjDef().Name(nLocale),perResp.Name(nLocale)]);
    }
    else
    {
         xOutput.addRow([oModel.Name(nLocale),oModel.GUID(),perspective.Name(nLocale), perspective.GUID(),curTarget.ObjDef().Name(nLocale),curTarget.ObjDef().GUID(),initativ,perResp]);
    }
}

/**
**  Function to set the columns of the xml file
**  @return void
**/
function outHeader()
{
     xOutput.setColumns([
                       ["Model name","text"],
                       ["Model GUID","text"],
                       ["Perspective name","text"],
                       ["Perspective GUID","text"],
                       ["Objective name","text"],
                       ["Objective GUID","text"],
                       ["Initiative name", "text"],
                       ["Person Respons","text"]
                       ]);   
}