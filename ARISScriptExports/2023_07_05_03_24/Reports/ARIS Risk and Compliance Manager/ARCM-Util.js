/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

//---------------------------------------------------------------------------------------
//------------------------------------- COMMON ------------------------------------------

/*---------------------------------------------------------------------------------------
    Converts a Java HashSet to a JavaScript Array.
 ---------------------------------------------------------------------------------------*/   
function convertHashSetToJSArray(p_hashSet) {
    var jsArray = new Array();
    if (p_hashSet == null) {return jsArray;} 
    for (var it = p_hashSet.iterator();it.hasNext();) {
        jsArray.push(it.next());
    }
    return jsArray;
}

/*---------------------------------------------------------------------------------------
    Converts a JavaScript Array to a Java HashSet.
 ---------------------------------------------------------------------------------------*/   
function convertJSArrayToHashSet(p_jsArray) {
    var oHashSet = new java.util.HashSet();
    if (p_jsArray == null) {return oHashSet;} 
    for (var i=0; i<p_jsArray.length; i++) {
        oHashSet.add(p_jsArray[i]);
    }
    return oHashSet;
}

//---------------------------------------------------------------------------------------
//------------------------------------- SORTING -----------------------------------------


/*---------------------------------------------------------------------------------------
    Sort a JS-Array of child ObjDefs according to the XY position of their occurrences in 
    a given model type.
    
    For each ObjDefs the occurrence in the *first* found model with the given type is 
    used. After this the XY positions are used for sorting.
    ObjDefs that have either no occurences at all or no occurences in a model of the 
    given model type are ignored for sorting - they are appended at the end of the returned
    sorted list.
    
    Returns a sorted JS-Array with the ObjDefs given as input.
---------------------------------------------------------------------------------------*/
function sortObjDefsByXYPosInModel( aObjDefs, iModelType ) {
    return sortObjDefsByPositionInModel( aObjDefs, iModelType, sortObjOccsByXYPos ); 
}

/*---------------------------------------------------------------------------------------
    Does the same as sortObjDefsByXYPosInModel but with Y position before X position.
---------------------------------------------------------------------------------------*/
function sortObjDefsByYXPosInModel( aObjDefs, iModelType ) {
    return sortObjDefsByPositionInModel( aObjDefs, iModelType, sortObjOccsByYXPos ); 
}

/*---------------------------------------------------------------------------------------
    Sort a JS-Array of child ObjDefs according to the XY position of their occurrences in 
    a given model type.
    
    For each ObjDefs the occurrence in the *first* found model with the given type is 
    used. After this the XY positions are used for sorting.
    ObjDefs that have either no occurences at all or no occurences in a model of the 
    given model type are ignored for sorting - they are appended at the end of the returned
    sorted list.
    
    Returns a sorted JS-Array with the ObjDefs given as input.
---------------------------------------------------------------------------------------*/
function sortObjDefsByPositionInModel( aObjDefs, iModelTypeNum, sSortFunctionName ) {

    //determine valid ObjOccs and ObjDefs without valid occurences
    var nonOccurrenceObjDefs = new Array();
    var sortableObjOccs = new Array();
    
    //if a model type is given
    if (iModelTypeNum != null && iModelTypeNum != -1) {
        
        for (var i=0; i<aObjDefs.length; i++) {       
            var bOccurrenceFound = false;
            var aObjDefOccurrences = aObjDefs[i].OccList();
            for (var j=0; j<aObjDefOccurrences.length; j++) {
                //stop at the first found ObjOcc - add it to the ObjOcc array
                if (aObjDefOccurrences[j].Model().TypeNum() == iModelTypeNum) {
                    sortableObjOccs.push( aObjDefOccurrences[j] );
                    bOccurrenceFound = true;
                    break;    
                } 
            }
            //if nothing found - add the ObjDef to the 'not sortable' list
            if (!bOccurrenceFound) {
                nonOccurrenceObjDefs.push( aObjDefs[i] );
            }
        }
    }
    else {
        sortableObjOccs = findChildObjOccsToSort(aObjDefs);
    }
    
    
    //sort the occurences according to the sort function
    sortableObjOccs.sort( eval(sSortFunctionName) );
    
    //remove duplicate occurrences
    //sortableObjOccs = removeDuplicateObjOccs(sortableObjOccs);
    
    //create the result array
    var resultArray = new Array();
    for (var k=0; k<sortableObjOccs.length; k++) {
        resultArray.push( sortableObjOccs[k].ObjDef() );
    }
    for (var l=0; l<nonOccurrenceObjDefs.length; l++) {
        resultArray.push( nonOccurrenceObjDefs[l] );
    }
    
    return resultArray;
}


/*---------------------------------------------------------------------------------------
    Sorts the given list of child ObjDefs if the following conditions are fulfilled:
    - There is a model where the parent ObjDef as well all child ObjDefs have at least
      one ObjOcc
    - All child ObjOccs are connected to the parent ObjOcc
    
    If so then the child ObjOccs are sorted by their XY position. An Array of their 
    sorted ObjDefs is returned.
    If not then the unsorted ObjDef list is returned.
---------------------------------------------------------------------------------------*/
function sortObjDefsByXYOfLinkedOccs( oParentObjDef, aChildObjDefs ) {
    return sortObjDefsByXYOfLinkedOccs( oParentObjDef, aChildObjDefs, sortObjOccsByXYPos );
}

/*---------------------------------------------------------------------------------------
    Does the same as sortObjDefsByXYOfLinkedOccs but with Y position before X position.
---------------------------------------------------------------------------------------*/
function sortObjDefsByYXOfLinkedOccs( oParentObjDef, aChildObjDefs ) {
    return sortObjDefsByYXOfLinkedOccs( oParentObjDef, aChildObjDefs, sortObjOccsByYXPos );
}

/*---------------------------------------------------------------------------------------
    Sorts the given list of child ObjDefs if the following conditions are fulfilled:
    - There is a model where the parent ObjDef as well all child ObjDefs have at least
      one ObjOcc
    - All child ObjOccs are connected to the parent ObjOcc
    
    If so then the child ObjOccs are sorted by the given sort function. An Array of their 
    sorted ObjDefs is returned.
    If not then the unsorted ObjDef list is returned.
---------------------------------------------------------------------------------------*/
function sortObjDefsByPositionOfLinkedOccs( oParentObjDef, aChildObjDefs, sSortFunctionName ) {
    
    if (aChildObjDefs == null) {return null;}
    if (aChildObjDefs.length <= 1) {return aChildObjDefs;}
    
    var aObjOccsToSort = findParentLinkedChildObjOccsToSort( oParentObjDef, aChildObjDefs );
    
    //if no model could be found where the conditions for the ObjOccs were fulfilled then return unsorted list
    if (aObjOccsToSort.length == 0) {
        return aChildObjDefs;    
    }
    
    //otherwise: sort the child ObjDefs according to the sort function (remove the not needed
    //parent ObjOcc at the beginning)
    aObjOccsToSort.shift();
    aObjOccsToSort.sort( eval(sSortFunctionName) );
    
    //remove duplicate occurrences
    //aObjOccsToSort = removeDuplicateObjOccs(aObjOccsToSort);
    
    //create the result array
    var resultArray = new Array();
    for (var o=0; o<aObjOccsToSort.length; o++) {
        resultArray.push( aObjOccsToSort[o].ObjDef() );    
    }
    return resultArray;
}


/*---------------------------------------------------------------------------------------
 Sorts ObjOccs by their XY position. ObjOccs with smaller X get a smaller index.
 If elements have the same X then the Y is compared. Again ObjOccs with smaller
 Y get a smaller index.
---------------------------------------------------------------------------------------*/
function sortObjOccsByXYPos( firstObjOcc, secondObjOcc ) {
    
    if (firstObjOcc.X() - secondObjOcc.X() != 0) {
        return firstObjOcc.X() - secondObjOcc.X();
    } else {
        return firstObjOcc.Y() - secondObjOcc.Y();
    }    
}

/*---------------------------------------------------------------------------------------
 Sorts ObjOccs by their YX position. ObjOccs with smaller Y get a smaller index.
 If elements have the same Y then the X is compared. Again ObjOccs with smaller
 X get a smaller index.
---------------------------------------------------------------------------------------*/
function sortObjOccsByYXPos( firstObjOcc, secondObjOcc ) {
    
    if (firstObjOcc.Y() - secondObjOcc.Y() != 0) {
        return firstObjOcc.Y() - secondObjOcc.Y();
    } else {
        return firstObjOcc.X() - secondObjOcc.X();
    }    
}


/*---------------------------------------------------------------------------------------
    Sorts the given list of child ObjDefs if the following conditions are fulfilled:
    - There is a model where the parent ObjDef as well all child ObjDefs have at least
      one ObjOcc
    - All child ObjOccs are connected to the parent ObjOcc
    
    If so then the child ObjOccs are sorted by their distance to the parent ObjOcc. An 
    Array of their sorted ObjDefs is returned.
    If not then the unsorted ObjDef list is returned.
---------------------------------------------------------------------------------------*/
function sortObjDefsByDistanceOfLinkedOccs( oParentObjDef, aChildObjDefs ) {
    
    if (aChildObjDefs == null) {return null;}
    if (aChildObjDefs.length <= 1) {return aChildObjDefs;}
    
    var aObjOccsToSort = findParentLinkedChildObjOccsToSort( oParentObjDef, aChildObjDefs );
    
    //if no model could be found where the conditions for the ObjOccs were fulfilled then return unsorted list
    if (aObjOccsToSort.length == 0) {
        return aChildObjDefs;    
    }
    
    //calculate the distances of the child ObjOccs (from center to center)
    var oParentObjOcc = aObjOccsToSort.shift();
    var aDistanceArray = new Array();
    for (var i=0; i<aObjOccsToSort.length; i++) {
        
        var iParentXCenter = oParentObjOcc.X() + (oParentObjOcc.Width()/2);
        var iParentYCenter = oParentObjOcc.Y() - (oParentObjOcc.Height()/2);
        var iChildXCenter = aObjOccsToSort[i].X() + (aObjOccsToSort[i].Width()/2);
        var iChildYCenter = aObjOccsToSort[i].Y() - (aObjOccsToSort[i].Height()/2);
        
        var xDist = Math.abs( iParentXCenter - iChildXCenter );
        var yDist = Math.abs( iParentYCenter - iChildYCenter );
        var fDistance = Math.sqrt( Math.pow(xDist,2) + Math.pow(yDist,2) );
        
        var oObjOccDistanceContainer = new ObjOccDistanceContainer();
        oObjOccDistanceContainer.Occ = aObjOccsToSort[i];
        oObjOccDistanceContainer.Distance = fDistance;
        aDistanceArray.push( oObjOccDistanceContainer );
    }
    
    //sort the distance container array
    aDistanceArray.sort(sortDistanceContainers);
    
    //remove the duplicate containers
    //aDistanceArray = removeDuplicateDistanceContainers(aDistanceArray);
    
    //create the result array
    var resultArray = new Array();
    for (var j=0; j<aDistanceArray.length; j++) {
        resultArray.push( aDistanceArray[j].Occ.ObjDef() );    
    }
    return resultArray;
}

/*---------------------------------------------------------------------------------------
 Sorts ObjOccs by their XY position. ObjOccs with smaller X get a smaller index.
 If elements have the same X then the Y is compared. Again ObjOccs with smaller
 Y get a smaller index.
---------------------------------------------------------------------------------------*/
function sortByGUID( firstObject, secondObject ) {
    if (firstObject == null || secondObject == null) {return 0;}
    var firstGUID = firstObject.GUID();
    var secondGUID = secondObject.GUID();
    return firstGUID.localeCompare(secondGUID);        
}

/*---------------------------------------------------------------------------------------
    Data container for sorting ObjOccs by distance
---------------------------------------------------------------------------------------*/
function ObjOccDistanceContainer( oChildObjOcc, fDistance ) {    
    this.Occ = oChildObjOcc;
    this.Distance = fDistance;   
}

/*---------------------------------------------------------------------------------------
 Sorts ObjOccs by their XY position. ObjOccs with smaller X get a smaller index.
 If elements have the same X then the Y is compared. Again ObjOccs with smaller
 Y get a smaller index.
---------------------------------------------------------------------------------------*/
function sortDistanceContainers( firstContainer, secondContainer ) {
    return firstContainer.Distance - secondContainer.Distance;        
}

/*---------------------------------------------------------------------------------------
 Determines a set of child ObjOccs to be sorted by the following steps:
 - There must be a model where the parent ObjDef as well as all child ObjDefs have
   occurrences; if no such model exists then an empty Array is returned
 - Inside this model the child ObjOccs must be connected to the parent ObjOcc; if this
   condition is not fulfilled by any qualified model from the first step then an empty 
   Array is returned
 
 If both conditions are met then a list of ObjOccs is returned where the parent ObjOcc
 is the first element, followed by all child ObjOccs to be sorted.
 If no parent or child ObjDefs are passed to the function then an empty Array is 
 returned.
---------------------------------------------------------------------------------------*/
function findParentLinkedChildObjOccsToSort( oParentObjDef, aChildObjDefs ) {

    if (oParentObjDef == null || aChildObjDefs == null || aChildObjDefs.length <= 1) {return new Array();}
    
    var aObjOccsToSort = new Array();
    
    //find a model where parent and all child objdefs have linked occurrences
    var aParentObjOccs = oParentObjDef.OccList();
    for (var i=0; i<aParentObjOccs.length; i++) {

        var oParentObjDefModel = aParentObjOccs[i].Model();
        
        //all children must have ObjOccs in this model
        var aFoundObjOccs = new Array();
        for (j=0; j<aChildObjDefs.length; j++) {
            var aChildObjOccs = aChildObjDefs[j].OccList();
            for (k=0; k<aChildObjOccs.length; k++) {
                if (aChildObjOccs[k].Model().equals( oParentObjDefModel )) {
                    //there can be multiple occurrences of the same ObjDef in this model -> keep them all and find out
                    //the real connected one in the next step
                    aFoundObjOccs.push(aChildObjOccs[k]);
                }
            }
        }
        
        //the children ObjOccs must be linked to the parent ObjOcc
        var aConnectedObjOccs = new Array();
        var aParentObjOccCxns = aParentObjOccs[i].CxnOccList();
        for (var m=0; m<aFoundObjOccs.length; m++) {
            for (var n=0; n<aParentObjOccCxns.length; n++) {
                if (aParentObjOccCxns[n].SourceObjOcc().equals( aFoundObjOccs[m] )
                    || aParentObjOccCxns[n].TargetObjOcc().equals( aFoundObjOccs[m] )
                    ) {
                    aConnectedObjOccs.push( aFoundObjOccs[m] );
                    break;
                }
            }
        }
        //if the Array lengths don't match then at least one found child ObjOcc is not connected to the parent ObjOcc -> try the next model
        if (aConnectedObjOccs.length < aChildObjDefs.length) {
             var x=0;
            continue;
        }
    
        //if everything is fulfilled then work with these child ObjOccs
        aObjOccsToSort.push(aParentObjOccs[i]);
        for (var p=0; p<aConnectedObjOccs.length; p++) {
            aObjOccsToSort.push(aConnectedObjOccs[p]);    
        }
        break;
    }
    
    return aObjOccsToSort;    
}

/*---------------------------------------------------------------------------------------
 Determines a set of child ObjOccs to be sorted by the following steps:
 - There must be a model where the all child ObjDefs have an occurrence; 
   if no such model exists then an empty Array is returned
 
 If this condition is met then a list of the found ObjOccs is returned.
 If no parent or child ObjDefs are passed to the function then an empty Array is 
 returned.
---------------------------------------------------------------------------------------*/
function findChildObjOccsToSort( aChildObjDefs ) {

    if (aChildObjDefs == null || aChildObjDefs.length <= 1) {return new Array();}
    
    var aObjOccsToSort = new Array();
    var oFirstChildObjDef = aChildObjDefs[0];
    
    //find a model where all child objdefs have occurrences
    var aFirstChildObjOccs = oFirstChildObjDef.OccList();
    for (var i=0; i<aFirstChildObjOccs.length; i++) {

        var oFirstChildObjDefModel = aFirstChildObjOccs[i].Model();
        
        //all other children must have ObjOccs in this model
        var aFoundObjOccs = new Array();
        aFoundObjOccs.push(aFirstChildObjOccs[i]);
        for (j=1; j<aChildObjDefs.length; j++) {
            var aChildObjOccs = aChildObjDefs[j].OccList();
            for (k=0; k<aChildObjOccs.length; k++) {
                if (aChildObjOccs[k].Model().equals( oParentObjDefModel )) {
                    //there can be multiple occurrences of the same ObjDef in this model -> keep them all and find out
                    //the real connected one in the next step
                    aFoundObjOccs.push(aChildObjOccs[k]);
                }
            }
        }
        
        
        //if the Array lengths don't match then at least one found child ObjOcc is not connected to the parent ObjOcc -> try the next model
        if (aFoundObjOccs.length < aChildObjDefs.length) {
            continue;
        }
    
        //if everything is fulfilled then work with these child ObjOccs
        for (var p=0; p<aFoundObjOccs.length; p++) {
            aObjOccsToSort.push(aFoundObjOccs[p]);    
        }
        break;
    }
    
    return aObjOccsToSort;    
}

/*---------------------------------------------------------------------------------------
 Removes all ObjOcc "duplicates" i.e. if several ObjOccs have the same ObjDef then only
 the first ObjOcc is retained, all others are removed.
 This function takes a JS-Array of ObjOccs and returns a new JS-Array of ObjOccs.
---------------------------------------------------------------------------------------*/
function removeDuplicateObjOccs( aObjOccs ) {
    
    var singleSortableObjOccs = new Array();
    var oObjDefSet = new Packages.java.util.HashSet();
    
    for (var m=0; m<aObjOccs.length; m++) {
        if (oObjDefSet.add(aObjOccs[m].ObjDef())) {
            singleSortableObjOccs.push( aObjOccs[m] );
        }
    }
    
    return singleSortableObjOccs;
}

/*---------------------------------------------------------------------------------------
 Removes all DistanceContainer "duplicates" i.e. if several Containers have the same 
 ObjDef then only the first Container is retained, all others are removed.
 This function takes a JS-Array of ObjOccDistanceContainers and returns a new JS-Array 
 of ObjOccDistanceContainer.
---------------------------------------------------------------------------------------*/
function removeDuplicateDistanceContainers( aDistanceArray ) {
    
    var singleSortableDistanceContainers = new Array();
    var oObjDefSet = new Packages.java.util.HashSet();
    
    for (var m=0; m<aDistanceArray.length; m++) {
        if (oObjDefSet.add(aDistanceArray[m].Occ.ObjDef())) {
            singleSortableDistanceContainers.push( aObjOccs[m] );
        }
    }
    
    return singleSortableDistanceContainers;    
}


//---------------------------------------------------------------------------------------
//---------------------------- CONTEXT PROPERTY HANDLING --------------------------------

var g_sHashSetDelimiter = ",";
var g_sHashMapAssignDelimiter = "=";
var g_sHashMapEntryDelimiter = "#";

/*---------------------------------------------------------------------------------------
 Convert a HashSet into a property String.
 HashSet format: contains only ObjDefs
---------------------------------------------------------------------------------------*/
function convertHashSetToPropertyString(p_aObjDefHashSet) {
    return convertJSArrayToPropertyString( convertHashSetToJSArray(p_aObjDefHashSet), true );
}

/*---------------------------------------------------------------------------------------
 Convert a HashSet into a property String.
 HashSet format: contains only Strings
---------------------------------------------------------------------------------------*/
function convertStringHashSetToPropertyString(p_aStringHashSet) {
    return convertJSArrayToPropertyString( convertHashSetToJSArray(p_aStringHashSet), false );
}

/*---------------------------------------------------------------------------------------
 Convert a HashSet into a property String.
 HashSet format: contains only ObjDefs
---------------------------------------------------------------------------------------*/
function convertJSArrayToPropertyString(p_aJSArray, p_bStoreObjDefs) {
    if (p_aJSArray == null) {return null;}
    var oStringBuffer = new java.lang.StringBuffer( p_aJSArray.length * (36 + g_sHashSetDelimiter.length) );
    var bFirst = true;
    for (var i=0; i<p_aJSArray.length; i++) {
        if (!bFirst) {oStringBuffer.append(g_sHashSetDelimiter);}
        if (p_bStoreObjDefs != null && p_bStoreObjDefs) {
            oStringBuffer.append(p_aJSArray[i].GUID());
        } else {
            oStringBuffer.append(p_aJSArray[i]);
        }
        bFirst = false;
    }
    return oStringBuffer.toString();
}

/*---------------------------------------------------------------------------------------
 Convert a HashMap into a property String.
 HashMap format: ObjDef | JSArray <ObjDef>
---------------------------------------------------------------------------------------*/
function convertHashMapToPropertyString(p_aObjDefHashMap) {
    
    var iKeysStringLength = p_aObjDefHashMap.keySet().size() * 6 * (36 + g_sHashSetDelimiter.length);    
    var oStringBuffer = new java.lang.StringBuffer( iKeysStringLength );
    var bFirst = true;
    for (var oIterator = p_aObjDefHashMap.keySet().iterator(); oIterator.hasNext(); ) {
        var oCurrentKeyObjDef = oIterator.next();
        var aCurrentValueObjDefs = p_aObjDefHashMap.get(oCurrentKeyObjDef);
        
        if (aCurrentValueObjDefs == null) {continue;}
        if (!bFirst) {oStringBuffer.append(g_sHashMapEntryDelimiter);}
        
        oStringBuffer.append( oCurrentKeyObjDef.GUID() )
                     .append( g_sHashMapAssignDelimiter )
                     .append( convertHashSetToPropertyString( convertJSArrayToHashSet(aCurrentValueObjDefs) ) );
        bFirst = false;
    }
    return oStringBuffer.toString();
}

/*---------------------------------------------------------------------------------------
 Stores the content of a HashSet as a property String in the report's Context.
 HashSet format: contains only ObjDefs
---------------------------------------------------------------------------------------*/
function storeHashSetAsProperty(p_sPropertyKey, p_aObjDefHashSet) {
    Context.setProperty( p_sPropertyKey, convertHashSetToPropertyString(p_aObjDefHashSet) );
}

/*---------------------------------------------------------------------------------------
 Stores the content of a HashSet as a property String in the report's Context.
 HashSet format: contains only Strings
---------------------------------------------------------------------------------------*/
function storeStringHashSetAsProperty(p_sPropertyKey, p_aStringHashSet) {
    Context.setProperty( p_sPropertyKey, convertStringHashSetToPropertyString(p_aStringHashSet) );
}

/*---------------------------------------------------------------------------------------
 Stores the content of a HashMap as a property String in the report's Context.
 HashMap format: ObjDef | JSArray <ObjDef>
---------------------------------------------------------------------------------------*/
function storeHashMapAsProperty(p_sPropertyKey, p_aObjDefHashMap) {
    Context.setProperty( p_sPropertyKey, convertHashMapToPropertyString(p_aObjDefHashMap) );
}


/*---------------------------------------------------------------------------------------
 Convert a HashSet into a property String.
 HashSet format: contains only ObjDefs
---------------------------------------------------------------------------------------*/
function convertPropertyStringToHashSet(p_sPropertyString, p_oDatabase, p_bRestoreObjDefs) {
    
    var oHashSet = new java.util.HashSet();
    if (p_sPropertyString == null) {return oHashSet;}
    
    var sJSPropertyString = p_sPropertyString + ""; //Rhino String conversion issue for method "split"
    if (sJSPropertyString == "") {return oHashSet;}
    
    var splitRegex = java.util.regex.Pattern.compile(g_sHashSetDelimiter);
    var aGUIDs = splitRegex.split(sJSPropertyString);
    for ( var i=0; i<aGUIDs.length; i++) {
        //restore the ObjDefs from the property GUIDs
        if (p_bRestoreObjDefs != null && p_bRestoreObjDefs) {
            oHashSet.add( p_oDatabase.FindGUID(aGUIDs[i], Constants.CID_OBJDEF) );
        }
        //store the property String directly
        else {
            oHashSet.add( aGUIDs );  
        }
    }
    return oHashSet;
}

/*---------------------------------------------------------------------------------------
 Convert a HashMap into a property String.
 HashMap format: ObjDef | JSArray <ObjDef>
---------------------------------------------------------------------------------------*/
function convertPropertyStringToHashMap(p_sPropertyString, p_oDatabase) {
    
    var oHashMap = new java.util.HashMap();
    if (p_sPropertyString == null) {return oHashMap;}
    
    var sJSPropertyString = p_sPropertyString + ""; //Rhino String conversion issue for method "split"
    if (sJSPropertyString == "") {return oHashMap;}
    
    var splitRegexEntry = java.util.regex.Pattern.compile(g_sHashMapEntryDelimiter);
    var splitRegexAssign = java.util.regex.Pattern.compile(g_sHashMapAssignDelimiter);
    
    var aEntries = splitRegexEntry.split(sJSPropertyString);
    for (var i=0; i<aEntries.length; i++) {
        var aAssign = splitRegexAssign.split(aEntries[i] + ""); //Rhino String conversion issue
        
        var oKeyObjDef = p_oDatabase.FindGUID(aAssign[0], Constants.CID_OBJDEF);
        var oValueObjDefsHashSet = convertPropertyStringToHashSet(aAssign[1], p_oDatabase, true);
        oHashMap.put( oKeyObjDef, convertHashSetToJSArray(oValueObjDefsHashSet) );
    }
    return oHashMap;
}

/*---------------------------------------------------------------------------------------
 Restores a HashSet from property String of the report's Context.
 HashSet format: contains only ObjDefs
---------------------------------------------------------------------------------------*/
function restoreHashSetFromProperty(p_sPropertyKey, p_oDatabase) {
    if (p_sPropertyKey == null || Context.getProperty(p_sPropertyKey) == null) {return new java.util.HashSet();}
    return convertPropertyStringToHashSet( Context.getProperty( p_sPropertyKey ), p_oDatabase, true );
}

/*---------------------------------------------------------------------------------------
 Restores a HashSet from property String of the report's Context.
 HashSet format: contains only Strings
---------------------------------------------------------------------------------------*/
function restoreStringHashSetFromProperty(p_sPropertyKey, p_oDatabase) {
    if (p_sPropertyKey == null || Context.getProperty(p_sPropertyKey) == null) {return new java.util.HashSet();}
    return convertPropertyStringToHashSet( Context.getProperty( p_sPropertyKey ), p_oDatabase, false );
}

/*---------------------------------------------------------------------------------------
 Restores a HashMap from property String of the report's Context.
 HashMap format: ObjDef | JSArray <ObjDef>
---------------------------------------------------------------------------------------*/
function restoreHashMapFromProperty(p_sPropertyKey, p_oDatabase) {
    if (p_sPropertyKey == null || Context.getProperty(p_sPropertyKey) == null) {return new java.util.HashMap();}
    return convertPropertyStringToHashMap( Context.getProperty( p_sPropertyKey ), p_oDatabase );
}