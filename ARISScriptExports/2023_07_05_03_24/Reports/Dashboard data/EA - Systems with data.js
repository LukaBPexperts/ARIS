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
    var oSystems = oDB.Find(Constants.SEARCH_OBJDEF, [Constants.OT_APPL_SYS_TYPE]);    
    for (var i in oSystems) {
        outArisAttributes(oSystems[i]);
    }    
}

function outHeader() {
   output.setColumns([["System", "text"],
                      ["System-Guid", "text"],
                      ["Data center costs", "number"],
                      ["Criticality", "text"],
                      ["Criticality (numeric)", "number"],
                      ["Interface count", "number"]]);
}

function outArisAttributes(oSystem) {
    output.addRow([oSystem.Name(nLoc),
                   oSystem.GUID(),
                   getMeasureValue(Constants.AT_COST_COMP_CNT), // Data center costs
                   getValue(Constants.AT_CRITICALITY), // Criticality
                   getNumericCriticality(), // Criticality (numeric)
                   getInterfaceCount(oSystem)] // Interface count
                   ); 

    function getValue(nAttrType) {
        var oAttr = oSystem.Attribute(nAttrType, nLoc);
        if (oAttr.IsMaintained()) {
            return oAttr.getValue(); 
        }
        return "";
    }

    function getMeasureValue(nAttrType) {
        var oAttr = oSystem.Attribute(nAttrType, nLoc);
        if (oAttr.IsMaintained()) {
            return oAttr.MeasureValue(); 
        }
        return "";
    }

    function getNumericCriticality() {
        var oAttr = oSystem.Attribute(Constants.AT_CRITICALITY, nLoc);
        if (oAttr.IsMaintained()) {
            var numCrit = 0;
            switch (oAttr.MeasureUnitTypeNum()) {
                case Constants.AVT_VERY_LOW_1:
                numCrit = 1;
                break;
                case Constants.AVT_LOW_3:
                numCrit = 2;
                break;
                case Constants.AVT_AVG_1:
                numCrit = 3;
                break;
                case Constants.AVT_HIGH_1:
                numCrit = 4;
                break;
                case Constants.AVT_VERY_HIGH_1:
                numCrit = 5;
                break;
            }
            return numCrit;
        }
        return "";
    }
    
    function getInterfaceCount(sourceSystem) {
        var interfaceCxnDefList = sourceSystem.CxnListFilter(Constants.EDGES_INOUT, Constants.CT_SENDS_3);
        if(interfaceCxnDefList != null)
            return interfaceCxnDefList.length;
    
        return 0;
    }

}
