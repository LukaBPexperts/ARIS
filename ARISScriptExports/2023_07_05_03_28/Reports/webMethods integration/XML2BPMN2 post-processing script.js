/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

/**
 * BPMN 2.0 Import transformation post-processing script
 * @author y12501
 */

/**
 * Determines if the script should run in debug mode.
 */
const IS_DEBUGGING_ENABLED = false;

/**
 * Determines if the part layout should be performed.
 */
const IS_PART_LAYOUT_ENABLED = false;

var DEBUG_MESSAGE_ID = 1;

var database = ArisData.getActiveDatabase();
if (IS_PART_LAYOUT_ENABLED) {
    main();
}

function main() {
    var guids = getSelectedModelGuids();
    if (guids.length <= 0) {
        print("Couldn't find any model GUID!");
        return;
    }

    var layoutedOccurrencies = [];
    for (var i = 0; i < guids.length; i++) {
        var guid = guids[i];
        var model = database.FindGUID(guid);
        var occurrencies = doPartLayout(model);
        layoutedOccurrencies = layoutedOccurrencies.concat(occurrencies);
    }

    printLayoutedOccs(layoutedOccurrencies);
}

function getSelectedModelGuids() {
    var selectedModels = ArisData.getSelectedModels();
    var guids = [];

    for (var i = 0; i < selectedModels.length; i++) {
        var guid = selectedModels[i].GUID();
        guids.push(guid);
    }

    return guids;
}

function doPartLayout(model) {
    var selection = [];
    var objects = model.ObjOccList();

    for (var i = 0; i < objects.length; i++) {
        var object = objects[i];
        if (object != null && (object.SymbolNum() == Constants.ST_BPMN_POOL_1 || object.SymbolNum() == Constants.ST_BPMN_LANE_1)) {
            selection.push(object);
        }
    }

    model.doLayout(selection);
    return selection;
}

function printLayoutedOccs(message) {
    if (IS_DEBUGGING_ENABLED) {
        Context.setProperty("DEBUG_MESSAGE_" + DEBUG_MESSAGE_ID++, message);
        if (typeof(Dialogs) === "object") {
            Dialogs.MsgBox(message);
        }
    }
}

function printLayoutedOccs(occurrencies) {
    if (!IS_DEBUGGING_ENABLED) {
        return;
    }

    var dbLocale = database.getDbLanguage().LocaleId();
    var message = "";
    for (var i = 0; i < occurrencies.length; i++) {
        var occ = occurrencies[i];
        var def = occ.ObjDef();
        var guid = def.GUID();
        var name = def.Name(dbLocale);
        var symbolName = occ.SymbolName();
        message += guid + " - [" + symbolName + "] - " + name + "\n";
    }
    Context.setProperty("DEBUG_LAYOUT_PERFORMED_ON", message);
    if (typeof(Dialogs) === "object") {
        Dialogs.MsgBox(message);
    }
}
