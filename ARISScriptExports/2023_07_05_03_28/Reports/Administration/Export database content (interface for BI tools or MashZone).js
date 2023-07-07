/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

// Export settings
var g_bExportCompleteGroupPaths = true;     // set true, if the export should consider the complete group paths
var g_bExportFormatedAttrDefs   = false;    // set true, if the export should export formated attribute definition
var g_bExportUsersAndUsergroups = false;    // set true, if the export should contain users and user groups too
var g_nHierarchyDepth           = 0;        // sets the hierarchy depth for the export
var g_bIsSimpleExport           = false;    // set true, if the export must be done in simple mode
var g_nRelationDepth            = 0;        // sets the relation depth for the export
var g_bZipExport                = false;    // set true, if the export result should be zipped
var g_bAddDTD                   = true;     // set true, if the DTD file should be transfered to the client
  
  
var xmlExport = Context.getComponent("XMLExport");

var aSelection = getSelection();
var aLocaleIDs = [Context.getSelectedLanguage()]
var exportOptions = getExportOptions(xmlExport, aSelection, aLocaleIDs);

var result = xmlExport.startExport(exportOptions);
if (result.isSuccessful()) {
    Context.addOutputFile(Context.getSelectedFile(), result.getFileResult());
    if (g_bAddDTD) {
        var aDTD = Context.getFile("ARIS-Export.dtd", Constants.LOCATION_PROCESSGENERATOR);
        Context.addOutputFile("ARIS-Export.dtd", aDTD);
    } 
}

/*************************************************************************************************/ 
 
function getExportOptions(xmlExport, aSelection, aLocaleIDs) {
    var exportOptions = xmlExport.createDefaultXMLExportOptions(aSelection, aLocaleIDs);
    
    exportOptions.setExportCompleteGroupPaths(g_bExportCompleteGroupPaths);
    exportOptions.setExportFormatedAttributeDefinitions(g_bExportFormatedAttrDefs);
    exportOptions.setExportUsersAndUsergroups(g_bExportUsersAndUsergroups);
    exportOptions.setHierarchyDepth(g_nHierarchyDepth);
    exportOptions.setIsSimpleExport(g_bIsSimpleExport);
    exportOptions.setRelationDepth(g_nRelationDepth);
    exportOptions.setZipExport(g_bZipExport);
    
    return exportOptions;
}
 
function getSelection() {
    selection = ArisData.getSelectedDatabases();
    if (selection.length > 0) selection = [selection[0].RootGroup()];       // Database selected -> return main group

    if (selection.length == 0) selection = ArisData.getSelectedGroups();
    if (selection.length == 0) selection = ArisData.getSelectedModels();
    if (selection.length == 0) selection = ArisData.getSelectedObjDefs();
    return selection;
} 