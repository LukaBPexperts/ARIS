/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

Context.setProperty("excel-formulas-allowed", false); //default value is provided by server setting (tenant-specific): "abs.report.excel-formulas-allowed" 

Context.setProperty("Prop_attributes", "true");
Context.setProperty("Prop_enhancedAttrs", "true");
Context.setProperty("Prop_cxns", "true");    
Context.setProperty("Prop_enhancedCxns", "true");
Context.setProperty("Prop_occurenceLevel", evaluateOccLevel());

outputObjectInformation();  // cmp. report "Object information library" (standard/3fc05d50-a9f4-11df-03f2-c09e38d23e7e)

/***********************************************************/

function evaluateOccLevel() {
    if (ArisData.getSelectedObjOccs().length > 0 || ArisData.getSelectedModels().length > 0) return "true";
    return "false";
}
