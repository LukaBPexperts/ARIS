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
    var oObjList = oDB.Find(Constants.SEARCH_OBJDEF, [Constants.OT_OBJECTIVE]);    
    for (var i in oObjList) {
        var oObj = oObjList[i];
        var oCxnList = oObj.CxnListFilter(Constants.EDGES_OUT, Constants.CT_MEASURED_BY);
        for (var j in oCxnList) {
            var oCxn = oCxnList[j];
            outRow(oObj, oCxn.TargetObjDef());
        }
    }    
}

function outHeader() {
   output.setColumns([["Objective name", "text"],
                      ["Objective GUID", "text"],
                      ["KPI name", "text"],
                      ["KPI GUID", "text"]]);
}

function outRow(oObj, oKpi, sWeighting) {
    output.addRow([oObj.Name(nLoc), // Objective name
                   oObj.GUID(), // Objective GUID
                   oKpi.Name(nLoc), // KPI name
                   oKpi.GUID() // KPI GUID
                   ]);
}
