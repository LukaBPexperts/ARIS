/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

// Get either the selected models or the models in the selected groups
var aModels = [];
var aSelectedGroups = ArisData.getSelectedGroups();
if (aSelectedGroups.length == 0) {
    aModels = ArisData.getSelectedModels();
} else {
    for (var i in aSelectedGroups) {
        aModels = aModels.concat(aSelectedGroups[i].ModelList(true, null));
    }
}

// Clear the attributes at all models found
var nLoc = Context.getSelectedLanguage();
var nCount = 0;
for (var i in aModels) {
    var model = aModels[i];
    var bChanged = false;
    var oAttrPurpose = model.Attribute(Constants.AT_ARCHIMATE_PURPOSE, nLoc);
    if (oAttrPurpose.IsMaintained()) {
        oAttrPurpose.Delete();
        bChanged = true;
    }
    var oAttrContent = model.Attribute(Constants.AT_ARCHIMATE_CONTENT, nLoc);
    if (oAttrContent.IsMaintained()) {
        oAttrContent.Delete();
        bChanged = true;
    }
    if (bChanged) nCount++;
}
Dialogs.MsgBox(formatstring1(getString("TEXT_1"), nCount.toString()), Constants.MSGBOX_BTN_OK, getString("TEXT_2"));
