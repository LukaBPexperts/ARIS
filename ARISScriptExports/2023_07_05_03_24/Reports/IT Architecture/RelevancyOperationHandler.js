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
 * This handler use the main-functionality from the default mapping operation handler.
 * Only an additional history extension will be added
 *
 * @author rico
 */
function RelevancyOperationHandler() {
    
    var handler = new MappingOperationHandler();
    
    this.handle = function(database, processor) { 
        handler.handle(database, processor);
    }
    
    this.changeHistoryValue = function(value) {
        return value += " |relevancy";
    }
    
    this.getModelsToLayout = function(){
        return handler.getModelsToLayout();
    }
    
    //register
    handler.registerPostProcessDescFunction(this.changeHistoryValue);
}
