/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

main();

function main() {
    try {
        if(checkActiveFilter()) {
            Context.setProperty("SUCCESS", true);
        } else {
            Context.setProperty("SUCCESS", false);
        }
    } catch (e) {
        var message = e.message;
        
        if(message == null) {
            message = "Unexpected error";
        }
        
        Context.setProperty("SUCCESS", false);
        Context.setProperty("ERROR_MESSAGE", message);
    }
}

function checkActiveFilter() {
    var bIsFilterOk = checkFilter(ArisData.ActiveFilter());
    return bIsFilterOk;
}
