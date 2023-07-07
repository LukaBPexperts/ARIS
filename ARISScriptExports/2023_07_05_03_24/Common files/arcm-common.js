/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

var g_approach = ""; 
 
var g_bTestManagementExport = false;
var g_bRiskManagementExport = false;
var g_bSurveyManagementExport = false;
var g_bAuditManagementExport = false;
var g_bPolicyManagementExport = false;
var g_bGDPRExport = false;

//-------------------------------------------------------------------------------------------
//---------------------------------- APPROACH HANDLING --------------------------------------

/*---------------------------------------------------------------------------------------
    Determine the approach - "rba" or "cba".
    
    If inside the report context the parameter "ARCM_APPROACH" is specified as "rba"/"RBA"
    or "cba"/"CBA" then the approach is set accordingly.
    
    If the report context contains the parameter "AUTO_EXECUTION" which indicates that the 
    report was not directly executed by a user (but by another macro or report instead) and 
    the parameter "ARCM_APPROACH" then approach rba is used as fallback.
    
    If the report context does neither contain the parameters "ARCM_APPROACH" nor "AUTO_EXECUTION"
    then a dialog is displayed where the user must pick the desired approach. As pre-selection
	the last approach is used that was chosen at the previous display of this dialog.
---------------------------------------------------------------------------------------*/
function determineApproach() {
    determineApproach(false);
}

function determineApproach(showStandAloneHint) {
	
	if (g_approach != "") {
        return;
    }
	
    var paramValueApproach = Context.getProperty("ARCM_APPROACH");
    
    //---- approach is determined by passed report parameter (auto export report uses this parameter)
    if ( paramValueApproach == "rba" || paramValueApproach == "RBA" ) {
        g_approach = "RBA";
        return;
    }
    if ( paramValueApproach == "cba" || paramValueApproach == "CBA" ) {
        g_approach = "CBA";
        return;
    }
    
    //---- report is called by auto export report but somehow the approach parameter is missing
    var paramValueExecution = Context.getProperty("AUTO_EXECUTION");
    if ( paramValueExecution != null ) {
        //report is executed automated but approach was not set -> use "rba" as fallback
        g_approach = "RBA";
        return;
    }

    //---- report is executed manually
    //ask for approach -> dialog finish function will set the global variable
    return Dialogs.showDialog(new approachDialog(showStandAloneHint), Constants.DIALOG_TYPE_WIZARD, getString("COMMON_18"));
}

//---- Dialog ----
function approachDialog(showStandAloneHint) {
    // isOk = false means dialog finished with cancel
    // isOk = true means dialog finish with ok
    var isOk = false;
    
	// all member functions except for getPages can access the property "dialog" of the dialog class. Type is "UserDialog" (see help).
	// examples:
	// - Get the page with the specified index (e.g. 0): this.dialog.getPage(0)
	// - Define the list of visible wizard pages (e.g. 0,1): this.dialog.setActiveWizardPages([0,1])

	// see script help for object "DialogObject"
	
	// returns DialogTemplate[]
	// non-optional
	this.getPages = function() {
        var iDialogTemplate1 = Dialogs.createNewDialogTemplate(450, 100, "Select approach");
		
		iDialogTemplate1.Text(15,15,275,15,getString("COMMON_19"));
    
		iDialogTemplate1.OptionGroup("APPROACH_GROUP");
        iDialogTemplate1.OptionButton(25,35,275,15,"RBA","APPROACH");
        iDialogTemplate1.OptionButton(25,55,275,15,"CBA","APPROACH");
        if (showStandAloneHint == true) {
            iDialogTemplate1.Text(25,75,400,45,getString("TEXT_6"),"HINT");
        } 
		iDialogTemplate1.OKButton();
        
        return [iDialogTemplate1];
	}
    
    // initialize dialog pages (are already created and pre-initialized with static data from XML or template)
    // parameter: Array of DialogPage
    // see Help: DialogPage
    // user can set control values
    // optional
    this.init = function(aPages) {
        var previousDecision = null
        //ask for the decision the user took during last manual execution
        try {
            previousDecision = Context.getProfileString("ARCM_Export", "approach", "RBA");
        }
        catch (e) {
            //won't work for semantic macro - in this case just continue without this info ...
        }
        if (previousDecision != null) {
            g_approach = previousDecision;
        }
        
        if (g_approach == "RBA") {this.dialog.getPage(0).getDialogElement("APPROACH_GROUP").setSelection(0);}
		if (g_approach == "CBA") {this.dialog.getPage(0).getDialogElement("APPROACH_GROUP").setSelection(1);}
		
		//fallback: RBA as default
        if (g_approach == null || g_approach == "") {this.dialog.getPage(0).getDialogElement("APPROACH_GROUP").setSelection(0);}
     }
    
    // called after "Ok"/"Finish" has been pressed and the current state data has been applied
    // can be used to update your data
    // pageNumber: the current page number
    // bOK: true=Ok/finish, false=cancel pressed
    // optional
    this.onClose = function(pageNumber, bOk) {
        isOk = bOk;
        if (bOk) {
            var selectedApproachIndex = this.dialog.getPage(0).getDialogElement("APPROACH_GROUP").getSelectedIndex();
            if (selectedApproachIndex == 0) {
                g_approach = "RBA";
            }
            if (selectedApproachIndex == 1) {
                g_approach = "CBA";
            }
            //store approach decision for further executions
            Context.writeProfileString("ARCM_Export", "approach", g_approach);
        }    
    }  
    this.getResult = function() {
        return isOk;
    }
        
}



//---------------------------------------------------------------------------------------
//---------------------------------- MODE HANDLING --------------------------------------

/*---------------------------------------------------------------------------------------
    Determine the modules to export data for; default is:
    - do Test management
    - do Risk management
    - do Survey management
    - do Audit management
    - do Policy management
    - do GDPR
    Returns true if at least one mode is active, otherwise false.
---------------------------------------------------------------------------------------*/
function determineModules() {

    var licenseList = Context.getComponent("UMC").getAllLicenses();
    for (var i=0; i<licenseList.size(); i++) {
        var license = licenseList.get(i);
        var prod = license.getProduct();
        if (prod != null) {
            if (prod.getCode() == "YRCTM") {g_bTestManagementExport = true;}
            if (prod.getCode() == "YRCRM") {g_bRiskManagementExport = true;}
            if (prod.getCode() == "YRCSM") {g_bSurveyManagementExport = true;}
            if (prod.getCode() == "YRCAM") {g_bAuditManagementExport = true;}
            if (prod.getCode() == "YRCPM") {g_bPolicyManagementExport = true;}
        }
    }

    g_bGDPRExport = true;
}



//---------------------------------------------------------------------------------------------------
//---------------------------------- TOPOLOGICAL SORT HANDLING --------------------------------------

/*---------------------------------------------------------------------------------------
    Performs a topological sort on a given Array of elements.
    Returns the follwoing result object:
		- property "objDefSet": 	new sorted LinkedHashSet with all elements from
                                    objDefSet which are not part of any cycle
		- property "cycleArrays":	array of arrays which in turn contain the elements
									of the different detected cycles
---------------------------------------------------------------------------------------*/
function sortTopological(objDefSet, hm_parentObjDef2ChildObjDefsSet, hm_childObjDef2ParentObjDefsSet) {
    
    if (objDefSet == null) {
        return {
            objDefSet: objDefSet,
            cycleArrays: []
        }
    }
    
    var objDefArray = convertHashSetToJSArray(objDefSet);
    var unsortedObjDefArray = objDefArray.concat(new Array());
    if (objDefArray.length == 0 || hm_parentObjDef2ChildObjDefsSet == null || hm_childObjDef2ParentObjDefsSet == null) {
        return {
            objDefSet: new java.util.LinkedHashSet(),
            cycleArrays: []
        }
    }  
    
    //create successor and predecessor count maps
    var objDefSuccessorCounts = new java.util.HashMap(); //ObjDef -> count
    var objDefPredecessorCounts = new java.util.HashMap(); //ObjDef -> count
    for (var i=0; i<objDefArray.length; i++) {
        //successor count
		var count = 0; 
        var childrenSet = hm_parentObjDef2ChildObjDefsSet.get(objDefArray[i]);
        if (childrenSet != null) {
            count = childrenSet.size();
        }
        objDefSuccessorCounts.put(objDefArray[i], count);
		//predecessor count
		count = 0; 
        var childrenSet = hm_childObjDef2ParentObjDefsSet.get(objDefArray[i]);
        if (childrenSet != null) {
            count = childrenSet.size();
        }
        objDefPredecessorCounts.put(objDefArray[i], count);
    }
    
    //sort all cycle-free elements from bottom-up direction (i.e. no child elements)...
    var sortedObjDefArray = new Array();
    var bottomUpSortedObjDefArray = new Array();
    sortTopologicalBottomUp(unsortedObjDefArray, bottomUpSortedObjDefArray, objDefSuccessorCounts, hm_childObjDef2ParentObjDefsSet);
    sortedObjDefArray = sortedObjDefArray.concat(bottomUpSortedObjDefArray);
    //...if there are still unsorted elements left then sort all cycle-free elements from top-down direction (i.e. no parent elements)
    //this step is only done to narrow down the cycles more exactly
    if (unsortedObjDefArray.length > 0) {
        var topDownSortedObjDefArray = new Array();
        sortTopologicalTopDown(unsortedObjDefArray, topDownSortedObjDefArray, objDefPredecessorCounts, hm_parentObjDef2ChildObjDefsSet);
        sortedObjDefArray = sortedObjDefArray.concat(topDownSortedObjDefArray);
    }
    
    //if there remain unsorted elements then sorting was not successful - determine the cycles between them
    var cycles = new Array(); //array of cycle arrays
    if (unsortedObjDefArray.length > 0) {
        cycles = determineCycles(unsortedObjDefArray, hm_parentObjDef2ChildObjDefsSet);
    }
    
    var sortedObjDefLinkedHashSet = new java.util.LinkedHashSet();
    for (var m=0; m<sortedObjDefArray.length; m++) {
        sortedObjDefLinkedHashSet.add(sortedObjDefArray[m]);
    }
    
    return {
        objDefSet: sortedObjDefLinkedHashSet,
        cycleArrays: cycles
    }
}


/*---------------------------------------------------------------------------------------
    Fills bottomUpSortedObjDefArray with topologically sorted elements - order is
    children first.
    Leaves all elements that could not be sorted topologically bottom up in
    unsortedObjDefArray.
---------------------------------------------------------------------------------------*/
function sortTopologicalBottomUp(unsortedObjDefArray, bottomUpSortedObjDefArray, objDefSuccessorCounts, hm_childObjDef2ParentObjDefsSet) {
    var furtherIteration = true;
    while (furtherIteration) {
        //separate the elements by the criterion "successors" / "no successors"
        var objDefsWithoutSuccessors = new Array();
        var objDefsWithSuccessors = new Array();
        for (var i=0; i<unsortedObjDefArray.length; i++) {
            var count = objDefSuccessorCounts.get(unsortedObjDefArray[i]);
            if (count == 0) {
                bottomUpSortedObjDefArray.push(unsortedObjDefArray[i]);
                objDefsWithoutSuccessors.push(unsortedObjDefArray[i]); //mark the ObjDef as relevant for count updates of the parent ObjDefs
     
            } else {
                objDefsWithSuccessors.push(unsortedObjDefArray[i]);
            }
        }
        
        furtherIteration = objDefsWithoutSuccessors.length > 0 && objDefsWithSuccessors.length != unsortedObjDefArray.length;
        //in the next iteration only the elements with (unprocessed) successors will be checked
        unsortedObjDefArray.length = 0; //clear the array
        for (var h=0; h<objDefsWithSuccessors.length; h++) {
            unsortedObjDefArray.push(objDefsWithSuccessors[h]);
        }
        
        //update the successor count map
        for (var j=0; j<objDefsWithoutSuccessors.length; j++) {
            var parentObjDefsSet = hm_childObjDef2ParentObjDefsSet.get(objDefsWithoutSuccessors[j]);
            if (parentObjDefsSet != null) {
                var parentObjDefsArray = convertHashSetToJSArray(parentObjDefsSet);
                for (var k=0; k<parentObjDefsArray.length; k++) { 
                    var currentChildCount = objDefSuccessorCounts.get(parentObjDefsArray[k]);
                    objDefSuccessorCounts.put(parentObjDefsArray[k], --currentChildCount);
                }
            }
        }
    }
}


/*---------------------------------------------------------------------------------------
    Fills topDownSortedObjDefArray with topologically sorted elements - order is
    children(!) first.
    Leaves all elements that could not be  sorted topologically top down in
    unsortedObjDefArray.
---------------------------------------------------------------------------------------*/
function sortTopologicalTopDown(unsortedObjDefArray, topDownSortedObjDefArray, objDefPredecessorCounts, hm_parentObjDef2ChildObjDefsSet) {
    var furtherIteration = true;
    while (furtherIteration) {
        //separate the elements by the criterion "predecessors" / "no predecessors"
        var objDefsWithoutPredecessors = new Array();
        var objDefsWithPredecessors = new Array();
        for (var i=0; i<unsortedObjDefArray.length; i++) {
            var count = objDefPredecessorCounts.get(unsortedObjDefArray[i]);
            if (count == 0) {
                topDownSortedObjDefArray.push(unsortedObjDefArray[i]);
                objDefsWithoutPredecessors.push(unsortedObjDefArray[i]); //mark the ObjDef as relevant for count updates of the parent ObjDefs
     
            } else {
                objDefsWithPredecessors.push(unsortedObjDefArray[i]);
            }
        }
        
        furtherIteration = objDefsWithoutPredecessors.length > 0 && objDefsWithPredecessors.length != unsortedObjDefArray.length;
        //in the next iteration only the elements with (unprocessed) predecessors will be checked
        unsortedObjDefArray.length = 0; //clear the array
        for (var h=0; h<objDefsWithPredecessors.length; h++) {
            unsortedObjDefArray.push(objDefsWithPredecessors[h]);
        }
        
        //update the predecessor count map
        for (var j=0; j<objDefsWithoutPredecessors.length; j++) {
            var childObjDefsSet = hm_parentObjDef2ChildObjDefsSet.get(objDefsWithoutPredecessors[j]);
            if (childObjDefsSet != null) {
                var childObjDefsArray = convertHashSetToJSArray(childObjDefsSet);
                for (var k=0; k<childObjDefsArray.length; k++) { 
                    var currentParentCount = objDefPredecessorCounts.get(childObjDefsArray[k]);
                    objDefPredecessorCounts.put(childObjDefsArray[k], --currentParentCount);
                }
            }
        }
    }
    
    topDownSortedObjDefArray.reverse();
}


function determineCycles(unsortedObjDefsToCheckArray, hm_parentObjDef2ChildObjDefsSet) {
    
    var uncheckedObjDefs = unsortedObjDefsToCheckArray.concat(new Array());
    var cycles = new Array();
    
    while (uncheckedObjDefs.length > 0) {
        var startObjDef = uncheckedObjDefs[0];
        //find out to which cycle the start ObjDef belongs - cycleArray is filled by recursion
        var cycleArray = new Array();
        determineCycleForObjDefRecursive(startObjDef, unsortedObjDefsToCheckArray, hm_parentObjDef2ChildObjDefsSet, cycleArray);
        //add the ObjDef array representing the recognized cycle to the result
        cycles.push(cycleArray);
        //remove all cycle member from the array of objects to check
        for (var i=0; i<cycleArray.length; i++) {
            var index = uncheckedObjDefs.indexOf(cycleArray[i]);
            uncheckedObjDefs.splice(index, 1);
        }
    }
    
    return cycles;
}


function determineCycleForObjDefRecursive(ojDef, objDefsToCheckArray, hm_parentObjDef2ChildObjDefsSet, cycleArray) {
    var index = cycleArray.indexOf(ojDef);
    if (index >= 0) {
        return;
    }
    cycleArray.push(ojDef);
    var childObjDefSet = hm_parentObjDef2ChildObjDefsSet.get(ojDef);
    var childObjDefArray = convertHashSetToJSArray(childObjDefSet);
    for (var i=0; i<childObjDefArray.length; i++) {
        var childIndex = objDefsToCheckArray.indexOf(childObjDefArray[i]);
        //do recursion only for those children whose cycle membership is to check
        if (childIndex >= 0) {
            determineCycleForObjDefRecursive(childObjDefArray[i], objDefsToCheckArray, hm_parentObjDef2ChildObjDefsSet, cycleArray);
        }
    }
}