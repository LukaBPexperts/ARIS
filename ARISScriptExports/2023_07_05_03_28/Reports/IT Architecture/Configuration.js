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
 * Import configuration
 * 
 * Each configuration entry contains the following settings:
 *
 *    database   - name of ARIS database to import data into
 *    url        - url of Business Publisher with enabled IT Inventory
 *    export     - name of export on BP server to import changes from
 *    secret     - shared secret between BP server and import report
 *                 (used to establish trust between the two machines;
 *                 must be identical to secret given in IT Inventory
 *                 configuration)
 *    history    - API name of the attribute to use for the change history
 */
 
var configuration = new Array();
configuration.push(
    {
        "database": "United Motor Group",
        "url": "http://localhost:19990/businesspublisher/layouts/extensions/ext02/inventoryimport.jsp",
        "export": "UMG",
        "secret": "",
        "history": "AT_CHANGE_HISTORY",
        "layouter": "DefaultLayouter",
        "itemFinder": "DefaultItemFinder",
        "postImportProcessor": ""
    }
);

var handlerMappings = new Array();
handlerMappings.push(
    {
        "candidateOperations" : "CandidateOperationHandler",
        "requestOperations" : "RequestOperationHandler",
        "mappingOperations" : "MappingOperationHandler",
        "interfaceOperations" : "InterfaceOperationHandler",
        "relevancyOperations" : "RelevancyOperationHandler"
    }
);
