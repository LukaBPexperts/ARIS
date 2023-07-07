/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

//------constants------
const outputParameters = new OutputParameters("");
const xOutput = createXmlOutput(outputParameters);
const oModelType = ["247", "246", "272", "273", "13", "134", "173", "50", "140", "154", "221"];

//------global------
var nLocale = Context.getSelectedLanguage();
var oSelGroups = ArisData.getSelectedGroups();
var oModelCounter2 = 0;
var oModelsCount = 0;    
main();

/*************************************************************************************************/

//------functions------
function main()
{
    outHeader();
    var oModelCounter = 0;
    var modelCount = oChildsCounter(oSelGroups,oModelCounter);
    oChilds(oSelGroups, modelCount);
    uploadXmlOutputToADS(xOutput, outputParameters);
}

/**
**  Function to check if subfolder exist
**  @return void
**/
function oChilds(groups, modelCount)
{
    for(var i in groups)
    {
        readDatabase(groups[i], modelCount);
        var childGroups = groups[i].Childs();
        if(childGroups.length > 0)
        {
            oChilds(childGroups, modelCount);
        }
    }
}

/**
**  Function to read models from the database and output to an XML in ADS
**  @return void
**/
function readDatabase(group, modelCount)
{
    var oModels = group.ModelList();
    for (var j in oModels) 
    {
        //var oObjDef = oModels.ObjDef();
        for (var i in oModelType)
        {
            if (oModels[j].TypeNum() == oModelType[i])
            {
                var oIOTobjects = oModels[j].ObjDefListByTypes([Constants.OT_IOT_OBJECT])
                if (oIOTobjects.length > 0) 
                {          
                    xOutput.addRow([oModels[j].Name(nLocale),oModels[j].GUID(),oIOTobjects.length,modelCount]);
                }
                else
                {
                    xOutput.addRow([oModels[j].Name(nLocale),oModels[j].GUID(),oIOTobjects.length,modelCount]);
                }
            }
        }
    }
}

/**
**  Function for specifying the number of models
**  @return oModelCounter2
**/
function oChildsCounter(groups,oModelCounter)
{
    for(var i in groups)
    {
        oModelCounter = oModelCounter + readDatabase2(groups[i]);
        var childGroups = groups[i].Childs();
        if(childGroups.length > 0)
        {
            oChildsCounter(childGroups,oModelCounter);
        }
        else
        {
            oModelCounter2 = oModelCounter;
        }
    }
   return oModelCounter2;
}

/**
**  Function for specifying the number of models with IoT object
**  @return oModelsCount
**/
function readDatabase2(group)
{
    var oModels = group.ModelList();
    for (var j in oModels) 
    {
        for (var i in oModelType)
        {
            if (oModels[j].TypeNum() == oModelType[i])
            {
                oModelsCount++;
            }
        }
    }
    return oModelsCount;
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
                       ["IoT Objecte","number"],
                       ["Model Count", "number"]
                       ]);   
}