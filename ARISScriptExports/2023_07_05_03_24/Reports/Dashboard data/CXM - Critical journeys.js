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

const outputParameters = new OutputParameters("");
const output = createXmlOutput(outputParameters);

outHeader();

var oJourneyLandscapeModels = ArisData.getActiveDatabase().Find(Constants.SEARCH_MODEL, Constants.MT_CUSTOMER_JOURNEY_LANDSCAPE);
for (var i in oJourneyLandscapeModels) {
    var oJourneyLandscapeModel = oJourneyLandscapeModels[i];

    var oJourneyObjDefs = oJourneyLandscapeModel.ObjDefListByTypes([Constants.OT_CUSTOMER_JOURNEY]);
    for (var j in oJourneyObjDefs) {
        var oJourneyObjDef = oJourneyObjDefs[j];
        
        var oJourneyMapModels = oJourneyObjDef.AssignedModels(Constants.MT_CUSTOMER_JOURNEY_MAP);
        if (oJourneyMapModels.length > 0) {
            for (var k in oJourneyMapModels) {
                var oJourneyMapModel = oJourneyMapModels[k];
                
                var oTouchpointObjDefs = oJourneyMapModel.ObjDefListByTypes([Constants.OT_CUSTOMER_TOUCHPOINT]);
                for (var t in oTouchpointObjDefs) {
                    var oTouchpointObjDef = oTouchpointObjDefs[t];
                    
                    var oInitiativeObjDefs = getConnectedSourceObjects(oJourneyMapModel, oTouchpointObjDef, Constants.OT_FUNC_INST, Constants.CT_SUPPORTS_1);
                    var oRiskObjDefs =       getConnectedSourceObjects(oJourneyMapModel, oTouchpointObjDef, Constants.OT_RISK, Constants.CT_IS_ASSOCIATED_WITH);
                    
                    for (var m in oInitiativeObjDefs) {
                        var oInitiativeObjDef = oInitiativeObjDefs[m];
    
                        for (var n in oRiskObjDefs) {
                            var oRiskObjDef = oRiskObjDefs[n];
    
                            output.addRow([oJourneyLandscapeModel.Name(nLoc),
                                           oJourneyLandscapeModel.GUID(),
                                           oJourneyObjDef.Name(nLoc),
                                           oJourneyObjDef.GUID(),
                                           oJourneyMapModel.Name(nLoc),
                                           oJourneyMapModel.GUID(),
                                           oTouchpointObjDef.Name(nLoc),
                                           oTouchpointObjDef.GUID(),
                                           oInitiativeObjDef.Name(nLoc),
                                           oRiskObjDef.Name(nLoc),
                                           getAttrValue(oRiskObjDef, Constants.AT_AAM_PROBABILITY),
                                           getAttrValue(oRiskObjDef, Constants.AT_AAM_IMPACT)]);
                        }
                    }
                }
            }
        } else {
            output.addRow([oJourneyLandscapeModel.Name(nLoc),
                           oJourneyLandscapeModel.GUID(),
                           oJourneyObjDef.Name(nLoc),
                           oJourneyObjDef.GUID(),
                           "",
                           "",
                           "",
                           "",
                           "",
                           "",
                           "",
                           ""]);
        }
    }
}
uploadXmlOutputToADS(output, outputParameters);

/**************************************************************************************************************/

function getConnectedSourceObjects(oJourneyMapModel, oTrgObjDef, nSrcObjType, nCxnType) {
    var oSourceObjDefs = new Array();
    
    var oTrgObjOccs = oTrgObjDef.OccListInModel(oJourneyMapModel);
    for (var i in oTrgObjOccs) {
        var oInCxns = oTrgObjOccs[i].Cxns(Constants.EDGES_IN);
        for (var j in oInCxns) {
            var oInCxn = oInCxns[j];
            if (oInCxn.CxnDef().TypeNum() != nCxnType) continue;

            var oSrcObjOcc = oInCxn.SourceObjOcc();
            var oSrcObjDef = oSrcObjOcc.ObjDef();
            if (oSrcObjDef.TypeNum() != nSrcObjType) continue;
            
            oSourceObjDefs.push(oSrcObjDef);
        }
    }
    return ArisData.Unique(oTouchpointObjDefs);
}

function getAttrValue(oItem, nAttrType) {
    var oAttr = oItem.Attribute(nAttrType, nLoc);
    if (!oAttr.IsMaintained()) return "";
    
    return oAttr.getValue();
}

function outHeader() {
    output.setColumns([["Customer-journey-landscape","text"],
                       ["Customer-journey-landscape-Guid", "text"],
                       ["Customer-journey", "text"],
                       ["Customer-journey-Guid", "text"],
                       ["Customer-journey-map", "text"],
                       ["Customer-journey-map-Guid", "text"],
                       ["Customer touchpoint", "text"],
                       ["Customer touchpoint-Guid", "text"],
                       ["Initiative", "text"],
                       ["Risk", "text"],
                       ["Probability", "text"],
                       ["Impact", "text"]]);
}

