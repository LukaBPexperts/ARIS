/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

var nLoc = Context.getSelectedLanguage();
var oDB = ArisData.getActiveDatabase();

const outputParameters = new OutputParameters("");
const output = createXmlOutput(outputParameters);

outHeader();
main();
uploadXmlOutputToADS(output, outputParameters);

/*************************************************************************************************/

function main() {
    var oModelList = oDB.Find(Constants.SEARCH_MODEL, [Constants.MT_STRATEGY_DIAGRAM, Constants.MT_OBJ_DGM]);    
    for (var i in oModelList) {
        var oModel = oModelList[i];
        var oObjList = oModel.ObjDefListByTypes([Constants.OT_OBJECTIVE]);
        for (var j in oObjList) {
            var oObj = oObjList[j];
            outRow(oModel, oObj);
        }
    }    
}

function outHeader() {
   output.setColumns([["Model GUID", "text"],
                      ["Objective GUID", "text"]]);
}

function outRow(oModel, oObj) {
    output.addRow([oModel.GUID(), // Model GUID
                   oObj.GUID() // Objective GUID
                   ]);
}
