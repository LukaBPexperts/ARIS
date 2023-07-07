/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

var gLanguage = Context.getSelectedLanguage();

/*----------------------------------------------------------------------------------------------
 Write a given error string (pOutput) into a new output file (pOutputFilename)
 -----------------------------------------------------------------------------------------------*/ 
function writeErrorreport(pOutput, pHeader, pOutputFilename){
    
    if (pOutput.equals("")){
        // no error found
        return false;
    }
    else{
        pOutput =  (pHeader==null?"":pHeader) + pOutput;
        Context.setSelectedFile(pOutputFilename);
        var outFile = Context.createOutputObject(Constants.OUTTEXT, pOutputFilename);
        outFile.Init(gLanguage);
        outFile.OutputLn(pOutput, getString("COMMON_1"),  12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0);
        outFile.WriteReport();
        return true;
    } 
    
}

/*----------------------------------------------------------------------------------------------
 Write a given error string (pOutput) into a new output file (pOutputFilename)
 -----------------------------------------------------------------------------------------------*/ 
function writeErrorreport(pOutput, pOutputFilename){
    
    if (pOutput.equals("")){
        // no error found
        return false;
    }
    else{
        Context.setSelectedFile(pOutputFilename);
        var outFile = Context.createOutputObject(Constants.OUTTEXT, pOutputFilename);
        outFile.Init(gLanguage);
        outFile.OutputLn(pOutput, getString("COMMON_1"),  12, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0);
        outFile.WriteReport();
        return true;
    } 
    
}


/*---------------------------------------------------------------------------------------
    Checks the attribute type number:
    - for empty values -1 is returned
    - numerical values are not changed but only returned
    - if they belong to standard attributes the value of "Constants." + <sTypeNum>
      is returned
    - if they belong to custom attributes with GUID the currently valid TypeNumber
      is returned.
    - if it does not belong to any category above -1 is returned
	
	TODO Accessed by QuestionnaireTemplate semantic report - duplicate in export report
	-> shift to Classification or Utils
---------------------------------------------------------------------------------------*/
function getTypeNum(p_typeNum) {
    
    var sTypeNum = new java.lang.String(p_typeNum);
    
    //if no TypeNum was set
    if (sTypeNum.length() == 0) {
        return -1;    
    }
    
    //if TypeNum is already a number just return it
    if (!isNaN(sTypeNum)) {
        return p_typeNum;
    }
 
    //if TypeNum belongs to a standard attribute return the corresponding Constant
    var iTypeNum;
    try {
        iTypeNum = eval("Constants." + sTypeNum);
        return iTypeNum;
    }
    catch (e) {
        //ok, it does not
    }
 
    //check if TypeNum is the GUID of a user-defined attribute  
    try {
        iTypeNum = ArisData.ActiveFilter().UserDefinedAttributeTypeNum( sTypeNum );
        return iTypeNum;
    }
    catch (e) {
        //ok, it is not
    }
    
    //unknown TypeNum
    return -1; 
}