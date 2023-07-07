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
 * The default item finder returning items by the given guid.
 */
function DefaultItemFinder() {
    
    this.findItem = function(propertyName, mappedItemTypeName, guid, propertyMap, db) {
        
        if(db != null && guid != null)
            return db.FindGUID(guid);   
        
        return null;
    }
}