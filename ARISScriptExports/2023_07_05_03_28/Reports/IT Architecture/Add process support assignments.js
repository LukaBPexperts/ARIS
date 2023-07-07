/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

////////////////////////////////////////////////////////
//
// Report zum Hinzufügen eines neuen AST
//
////////////////////////////////////////////////////////

var g_oPsmConfig = new psmConfig(false);
var g_bIgnoreCheckBox_ToBe = g_oPsmConfig.ignoreCheckBox_ToBe();
    
var aAstData = new Array();

var theAttributeString = Context.getProperty("theAttributeString");
if (theAttributeString != null) {
    var theAttributeArray = theAttributeString.split("##");

    aAstData[0] = new AST_DATA(null);
    aAstData[0].aCxn2UnitData[0] = new CXN2UNIT_DATA(null, theAttributeArray);
}
   
var theModel = ArisData.getActiveDatabase().FindGUID(Context.getProperty("theModelGUID"));
var theUnitDef = ArisData.getActiveDatabase().FindGUID(Context.getProperty("theUnitGuid"));

createEditDialog(-1/*selectedIdx*/, false/*p_bMultiple*/, aAstData, [theUnitDef], g_bIgnoreCheckBox_ToBe);
