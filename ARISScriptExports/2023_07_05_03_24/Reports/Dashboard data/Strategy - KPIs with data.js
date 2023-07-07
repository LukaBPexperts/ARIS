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

const GOAL_ACC_POOR = "Poor";
const GOAL_ACC_AVERAGE = "Average";
const GOAL_ACC_GOOD = "Good";

const outputParameters = new OutputParameters("");
const output = createXmlOutput(outputParameters);

outHeader();
main();
uploadXmlOutputToADS(output, outputParameters);

/*************************************************************************************************/

function main() {
    var oKPIs = oDB.Find(Constants.SEARCH_OBJDEF, [Constants.OT_KPI]);
    for (var i in oKPIs) {
        outArisAttributes(oKPIs[i]);
    }
}

function outHeader() {
   output.setColumns([["KPI", "text"],
                      ["KPI GUID", "text"],
                      ["Goal accomplishment (quantitative)", "number"],
                      ["Goal accomplishment (qualitative)", "text"]]);
}

function outArisAttributes(oKPI) {

    var actVal = getMeasureValue(Constants.AT_ACT_VAL);
    var plVal = getMeasureValue(Constants.AT_PL_VAL);
    var minVal = getMeasureValue(Constants.AT_PL_VAL_MIN);
    var maxVal = getMeasureValue(Constants.AT_PL_VAL_MAX);
    var tol = getMeasureValue(Constants.AT_VAL_TOL);
    var isInverse = getInverse();

    output.addRow([oKPI.Name(nLoc),
                   oKPI.GUID(),
                   calculateGoalAccQuant(actVal, plVal, minVal, maxVal, isInverse),
                   calculateGoalAccQual(actVal, plVal, tol, isInverse)
                   ]);

    function getMeasureValue(nAttrType) {
        var oAttr = oKPI.Attribute(nAttrType, nLoc);
        if (oAttr.IsMaintained()) {
            return oAttr.MeasureValue();
        }
        return "";
    }

    function zeroIfEmpty(val) {
        return val === "" ? 0 : val;
    }

    function getInverse() {
        var oAttr = oKPI.Attribute(Constants.AT_KEY_INDIC_RATING, nLoc);
        return oAttr.IsMaintained() && oAttr.MeasureUnitTypeNum() == Constants.AVT_INVERS_PROP ? true : false;
    }

    // Calculate the quantitative goal accomplishment according to BSC methodology.
    // See "How to calculate the degree of goal accomplishment" in the ARIS help.
    // This function is also used in the report "EA - Objectives with data".
    function calculateGoalAccQuant(actVal, plVal, minVal, maxVal, isInverse) {
        var result = "";
        // Do not calculate anything if we don't have at least the actual and the plan value
        if (actVal === "" || plVal === "") {
            return result;
        }
        // In the inverse case we need a valid maximum value
        if (isInverse && maxVal === "") {
            return result;
        }
        if (isInverse) {
            maxVal = zeroIfEmpty(maxVal);
            result = (maxVal - actVal) / (maxVal - plVal);
        } else {
            minVal = zeroIfEmpty(minVal);
            result = (actVal - minVal) / (plVal - minVal);
        }
        return result;
    }

    // Determine the qualitative goal accomplishment according to BSC methodology.
    // See "How to calculate the degree of goal accomplishment" in the ARIS help.
    function calculateGoalAccQual(actVal, plVal, tol, isInverse) {
        var result = "";
        // Do not calculate anything if we don't have at least the actual and the plan value
        if (actVal === "" || plVal === "") {
            return result;
        }
        tol = zeroIfEmpty(tol);
        if (isInverse) {
            var tolVal = plVal + tol;
            if (actVal <= plVal) {
                result = GOAL_ACC_GOOD;
            } else if (plVal < actVal && actVal <= tolVal) {
                result = GOAL_ACC_AVERAGE;
            } else {
                result = GOAL_ACC_POOR;
            }
        } else {
            var tolVal = plVal - tol;
            if (actVal < tolVal) {
                result = GOAL_ACC_POOR;
            } else if (tolVal <= actVal && actVal < plVal) {
                result = GOAL_ACC_AVERAGE;
            } else {
                result = GOAL_ACC_GOOD;
            }
        }
        return result;
    }
}
