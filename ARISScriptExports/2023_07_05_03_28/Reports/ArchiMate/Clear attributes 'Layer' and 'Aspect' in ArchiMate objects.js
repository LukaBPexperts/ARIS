/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

// Get either the selected objects or the objects in the selected groups
var aObjects = [];
var aSelectedGroups = ArisData.getSelectedGroups();
if (aSelectedGroups.length == 0) {
    aObjects = ArisData.getSelectedObjDefs();
} else {
    for (var i in aSelectedGroups) {
        aObjects = aObjects.concat(aSelectedGroups[i].ObjDefList(true, null));
    }
}

// Clear the attributes at all objects found
var nLoc = Context.getSelectedLanguage();
var nCount = 0;
for (var i in aObjects) {
    var obj = aObjects[i];
    var bChanged = false;
    var oAttrAspect = obj.Attribute(Constants.AT_ARCHIMATE_ASPECT, nLoc);
    if (oAttrAspect.IsMaintained()) {
        oAttrAspect.Delete();
        bChanged = true;
    }
    var oAttrLayer = obj.Attribute(Constants.AT_ARCHIMATE_LAYER, nLoc);
    if (oAttrLayer.IsMaintained()) {
        oAttrLayer.Delete();
        bChanged = true;
    }
    if (bChanged) nCount++;
}
Dialogs.MsgBox(formatstring1(getString("TEXT_1"), nCount.toString()), Constants.MSGBOX_BTN_OK, getString("TEXT_2"));
