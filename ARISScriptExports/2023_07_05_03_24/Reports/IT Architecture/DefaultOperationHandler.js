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
 * Handles all operations for which no handler is defined by marking them as failed.
 *
 * @author MIPO
 */
function DefaultOperationHandler() {
    this.handle = function(db, processor) {
        var processor;
        while ((operation = processor.next()) != null) {
            processor.setStatus(2, "No handler found for operation."); 
        }
    }
}

