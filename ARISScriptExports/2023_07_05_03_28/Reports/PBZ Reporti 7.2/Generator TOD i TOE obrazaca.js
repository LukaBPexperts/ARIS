var nLocale = Context.getSelectedLanguage();
var gActiveDatabase = ArisData.getActiveDatabase();
var g_objVACD;
var g_objEPC;
var g_listCoveredRisks = new Array();
var g_listCoveredRiskFunctions = new Array();
var g_listControls = new Array();
var g_objISPTree;
var g_LLCOList = null;
var g_HLCOList = null;
var g_DomainList = null;

var g_isStandardProcess = true;
var g_generateTOD = true;
var g_generateTOE = true;
var g_tipCobitProcesa = "Sve";


var database = ArisData.getActiveDatabase();
var filter = database.ActiveFilter();
var objL262attribute = filter.UserDefinedAttributeTypeNum("eb39c830-88cd-11df-453f-005056b228c7");
var objISPProces4attribute = filter.UserDefinedAttributeTypeNum("e03dc090-88d1-11df-453f-005056b228c7");
var objDatumVerifikacijeattribute = filter.UserDefinedAttributeTypeNum("950e02b0-88d1-11df-453f-005056b228c7");
var objCheckEvidenceattribute = filter.UserDefinedAttributeTypeNum("317ba912-7df8-11da-0c60-cf8e338f9a0b");
var objPristupTestiranjuAttribute = filter.UserDefinedAttributeTypeNum("c05f4e80-88d4-11df-453f-005056b228c7");
var objTipCOBITCiljaAttribute = filter.UserDefinedAttributeTypeNum("7a9da400-88d4-11df-453f-005056b228c7");

main()

function isL262Relevant(objObject)
{
    
    if(objObject.Attribute(objL262attribute, nLocale).getValue() == "True")
        return true;
    else
        return false;
}

function main()
{
    var fileTemplateContent;
    var objXlsWorkBookTOD;
    var objXlsWorkBookTOE;
    
    showOutputOptionsDialog();
    
    var selClientFolders = Dialogs.getClientFolder("Odaberite gdje želite spremiti generirane obrasce:", "", false);
    Context.setSelectedPath( selClientFolders[0] ) ;
    g_objVACD = ArisData.getSelectedModels()[0];

	//Load ISP Process Tree
    fileTemplateContent = Context.getFile("ISP stablo procesa.xls", Constants.LOCATION_SCRIPT)
    var objISPTreeWorkBook = Context.getExcelReader(fileTemplateContent);
    g_objISPTree = objISPTreeWorkBook.getSheetAt(0);
       
	   
	if(g_isStandardProcess)   
	{
		getL262ControlListForStandardProces();
		
		if(g_generateTOD)
		{
			Context.setSelectedFile(g_objVACD.Name(nLocale) + "_TOD.xls");
			fileTemplateContent = Context.getFile("TOD - predložak.xls", Constants.LOCATION_SCRIPT)
			objXlsWorkBookTOD = Context.createExcelWorkbook(Context.getSelectedFile(), fileTemplateContent);
    
			for(var i=0; i<g_listControls.length; i++)
			{
				g_objEPC = getOccurenceInEPC(g_listControls[i]).Model();
				g_listCoveredRisks = getCoveredRisks(g_listControls[i]);
				g_listCoveredRiskFunctions = new Array();
				getCoveredRisksFunctions();
        
				objXlsWorkBookTOD.cloneSheet(0);
				objXlsWorkBookTOD.setSheetName(objXlsWorkBookTOD.getSheetCount()-1, g_listControls[i].Attribute(Constants.AT_AAM_CTRL_ID,nLocale).getValue())
				
				writeStandardControlTODSheet(g_listControls[i], objXlsWorkBookTOD.getSheetAt(objXlsWorkBookTOD.getSheetCount()-1));
				
			}
            
            objXlsWorkBookTOD.removeSheetAt(0);
            objXlsWorkBookTOD.write();
		}
        if(g_generateTOE)
		{
			Context.setSelectedFile(g_objVACD.Name(nLocale) + "_TOE.xls");
			fileTemplateContent = Context.getFile("TOE - predložak.xls", Constants.LOCATION_SCRIPT)
			objXlsWorkBookTOE = Context.createExcelWorkbook(Context.getSelectedFile(), fileTemplateContent);
    
			for(var i=0; i<g_listControls.length; i++)
			{
				g_objEPC = getOccurenceInEPC(g_listControls[i]).Model();
				g_listCoveredRisks = getCoveredRisks(g_listControls[i]);
				g_listCoveredRiskFunctions = new Array();
				getCoveredRisksFunctions();
        
				objXlsWorkBookTOE.cloneSheet(0);
				objXlsWorkBookTOE.setSheetName(objXlsWorkBookTOE.getSheetCount()-1, g_listControls[i].Attribute(Constants.AT_AAM_CTRL_ID,nLocale).getValue())
				
				writeStandardControlTOESheet(g_listControls[i], objXlsWorkBookTOE.getSheetAt(objXlsWorkBookTOE.getSheetCount()-1));
				
			}
            
            objXlsWorkBookTOE.removeSheetAt(0);
            objXlsWorkBookTOE.write();
		}
        
        
	}
    else
    {
        
        getL262ControlListForITProces();
        
        if(g_generateTOD)
		{
			Context.setSelectedFile(g_objVACD.Name(nLocale) + "_TOD.xls");
			fileTemplateContent = Context.getFile("ToD_IT - predložak.xls", Constants.LOCATION_SCRIPT)
			objXlsWorkBookTOD = Context.createExcelWorkbook(Context.getSelectedFile(), fileTemplateContent);
    
			for(var i=0; i<g_listControls.length; i++)
			{
				g_objEPC = getOccurenceInEPC(g_listControls[i]).Model();
				g_listCoveredRisks = getCoveredRisks(g_listControls[i]);
				g_listCoveredRiskFunctions = new Array();
				getCoveredRisksFunctions();
        
				objXlsWorkBookTOD.cloneSheet(0);
				objXlsWorkBookTOD.setSheetName(objXlsWorkBookTOD.getSheetCount()-1, g_listControls[i].Attribute(Constants.AT_AAM_CTRL_ID,nLocale).getValue())
				
				writeITControlTODSheet(g_listControls[i], objXlsWorkBookTOD.getSheetAt(objXlsWorkBookTOD.getSheetCount()-1));
				
			}
            
            objXlsWorkBookTOD.removeSheetAt(0);
            objXlsWorkBookTOD.write();
		}
        if(g_generateTOE)
		{
			Context.setSelectedFile(g_objVACD.Name(nLocale) + "_TOE.xls");
			fileTemplateContent = Context.getFile("TOE_IT - predložak.xls", Constants.LOCATION_SCRIPT)
			objXlsWorkBookTOE = Context.createExcelWorkbook(Context.getSelectedFile(), fileTemplateContent);
    
			for(var i=0; i<g_listControls.length; i++)
			{
				g_objEPC = getOccurenceInEPC(g_listControls[i]).Model();
				g_listCoveredRisks = getCoveredRisks(g_listControls[i]);
				g_listCoveredRiskFunctions = new Array();
				getCoveredRisksFunctions();
        
				objXlsWorkBookTOE.cloneSheet(0);
				objXlsWorkBookTOE.setSheetName(objXlsWorkBookTOE.getSheetCount()-1, g_listControls[i].Attribute(Constants.AT_AAM_CTRL_ID,nLocale).getValue())
				
				writeITControlTOESheet(g_listControls[i], objXlsWorkBookTOE.getSheetAt(objXlsWorkBookTOE.getSheetCount()-1));
				
			}
            
            objXlsWorkBookTOE.removeSheetAt(0);
            objXlsWorkBookTOE.write();
		}
        
    }    
    
}

function getL262ControlListForITProces()
{
    if(g_tipCobitProcesa == "Sve")
      getL262ControlListForStandardProces();
    else
    {
        var objConnectedEPCList = new Array();
	
        var objVACDFunctionList = g_objVACD.ObjDefListFilter(22);
        objVACDFunctionList =  ArisData.sort(objVACDFunctionList, Constants.AT_USER_ATTR5, nLocale);
        
        for(var i=0; i<objVACDFunctionList.length; i++)
        {
            objVACDFunction = objVACDFunctionList[i];
            objAssignedEPC = objVACDFunction.AssignedModels(13)[0];
            if(objAssignedEPC != null && isL262Relevant(objAssignedEPC))
                objConnectedEPCList[objConnectedEPCList.length] = objAssignedEPC;
        }
        
        for(var i=0; i<objConnectedEPCList.length; i++)
        {
            objVACDFunctionList = objConnectedEPCList[i].ObjOccListFilter(-1, 131407);
            objVACDFunctionList =  ArisData.sort(objVACDFunctionList, Constants.AT_USER_ATTR5, nLocale);
            
            if(objVACDFunctionList != null)
            {
                for(var j=0; j<objVACDFunctionList.length; j++)
                {   
                    if(objVACDFunctionList[j].ObjDef() != null && isL262Relevant(objVACDFunctionList[j].ObjDef()))
                    {
                        g_listCoveredRisks = getCoveredRisks(objVACDFunctionList[j].ObjDef());
                        
                        if(isTipCOBITCiljaEqual())
                            g_listControls[g_listControls.length] = objVACDFunctionList[j].ObjDef();
                    }
                }
            }
        }
    }
  
}

function isTipCOBITCiljaEqual()
{
    var oRiskCat = null;
    
    for(var i=0; i<g_listCoveredRisks.length; i++)
    {
        oRiskCat = getRiskCategory(g_listCoveredRisks[i]);
        
        if(oRiskCat == null)
            return false
        else
        {
            if( vbStrComp(oRiskCat.Attribute(objPristupTestiranjuAttribute,nLocale).getValue(), g_tipCobitProcesa )  == 0)
                return true;
        }
    }
    
    return false;
}

function getL262ControlListForStandardProces()
{
	var objConnectedEPCList = new Array();
	
    var objVACDFunctionList = g_objVACD.ObjDefListFilter(22);
    objVACDFunctionList =  ArisData.sort(objVACDFunctionList, Constants.AT_USER_ATTR5, nLocale);
    
    for(var i=0; i<objVACDFunctionList.length; i++)
    {
        objVACDFunction = objVACDFunctionList[i];
        objAssignedEPC = objVACDFunction.AssignedModels(13)[0];
        if(objAssignedEPC != null && isL262Relevant(objAssignedEPC))
            objConnectedEPCList[objConnectedEPCList.length] = objAssignedEPC;
    }
    
    for(var i=0; i<objConnectedEPCList.length; i++)
    {
        objVACDFunctionList = objConnectedEPCList[i].ObjOccListFilter(-1, 131407);
        objVACDFunctionList =  ArisData.sort(objVACDFunctionList, Constants.AT_USER_ATTR5, nLocale);
        
        if(objVACDFunctionList != null)
        {
            for(var j=0; j<objVACDFunctionList.length; j++)
            {
                if(objVACDFunctionList[j].ObjDef() != null && isL262Relevant(objVACDFunctionList[j].ObjDef()))
                    g_listControls[g_listControls.length] = objVACDFunctionList[j].ObjDef();
            }
        }
    }
}


function writeStandardControlTODSheet(objControl, objSheet)
{
    var strLevel4ISPProcess = objControl.Attribute(objISPProces4attribute, nLocale).getValue();
    var strLevel5ForComparison = strLevel4ISPProcess.replace(" ", "");
    
    var objProcessLevels = getISPProcessLevels(strLevel5ForComparison);
    
	objSheet.cell(9, 2).setCellValue(objProcessLevels[0]);
	objSheet.cell(10, 2).setCellValue(objProcessLevels[1]);
    objSheet.cell(11, 2).setCellValue(objProcessLevels[2]);
    objSheet.cell(12, 2).setCellValue(strLevel4ISPProcess);
    objSheet.cell(13, 2).setCellValue(g_objVACD.Name(nLocale));
    
    objSheet.cell(18, 1).setCellValue(g_objEPC.Name(nLocale));
    objSheet.cell(18, 2).setCellValue(getDateFormatForExcel(g_objVACD.Attribute(objDatumVerifikacijeattribute, nLocale).getValue()));
    objSheet.cell(18, 3).setCellValue(getFunctList());
    objSheet.cell(18, 5).setCellValue(getRiskList());
    objSheet.cell(18, 6).setCellValue(clean(getControlDescription(objControl)));
    objSheet.cell(18, 7).setCellValue(getControlExecutors(objControl));
    
    objSheet.cell(32, 6).setCellValue(objControl.Attribute(Constants.AT_AAM_CTRL_FREQUENCY, nLocale).getValue());
    objSheet.cell(34, 6).setCellValue(objControl.Attribute(objCheckEvidenceattribute, nLocale).getValue());
    
    objSheet.cell(44, 6).setCellValue( writeX(objControl.Attribute(Constants.AT_AAM_ASSERTIONS_EXIST_OCCURRENCE, nLocale).getValue()) );
    objSheet.cell(46, 6).setCellValue( writeX(objControl.Attribute(Constants.AT_AAM_ASSERTIONS_NA, nLocale).getValue()) );
    objSheet.cell(48, 6).setCellValue( writeX(objControl.Attribute(Constants.AT_AAM_ASSERTIONS_VALUATION_ALLOCATION, nLocale).getValue()) );
    objSheet.cell(50, 6).setCellValue( writeX(objControl.Attribute(Constants.AT_AAM_ASSERTIONS_COMPLETENESS, nLocale).getValue()) );
    objSheet.cell(52, 6).setCellValue( writeX(objControl.Attribute(Constants.AT_AAM_ASSERTIONS_RIGHTS_OBLIGATIONS, nLocale).getValue()) );
    objSheet.cell(54, 6).setCellValue( writeX(objControl.Attribute(Constants.AT_AAM_ASSERTIONS_PRESENTATION_DISCLOSURE, nLocale).getValue()) );
}

function writeStandardControlTOESheet(objControl, objSheet)
{
    var strLevel4ISPProcess = objControl.Attribute(objISPProces4attribute, nLocale).getValue();
    var strLevel5ForComparison = strLevel4ISPProcess.replace(" ", "");
    
    var objProcessLevels = getISPProcessLevels(strLevel5ForComparison);
    
	objSheet.cell(9, 2).setCellValue(objProcessLevels[0]);
	objSheet.cell(10, 2).setCellValue(objProcessLevels[1]);
    objSheet.cell(11, 2).setCellValue(objProcessLevels[2]);
    objSheet.cell(12, 2).setCellValue(strLevel4ISPProcess);
    objSheet.cell(13, 2).setCellValue(g_objVACD.Name(nLocale));
    
    objSheet.cell(18, 1).setCellValue(g_objEPC.Name(nLocale));
    objSheet.cell(18, 2).setCellValue(getDateFormatForExcel(g_objVACD.Attribute(objDatumVerifikacijeattribute, nLocale).getValue()));
    objSheet.cell(18, 3).setCellValue(getFunctList());
    objSheet.cell(18, 6).setCellValue(getRiskList());
    objSheet.cell(18, 7).setCellValue(clean(getControlDescription(objControl)));
    objSheet.cell(18, 8).setCellValue(getControlExecutors(objControl));
    
    objSheet.cell(23, 6).setCellValue(objControl.Attribute(Constants.AT_AAM_CTRL_ID, nLocale).getValue());
    objSheet.cell(29, 4).setCellValue(objControl.Attribute(Constants.AT_AAM_CTRL_FREQUENCY, nLocale).getValue());
    
    objSheet.cell(53, 6).setCellValue( writeX(objControl.Attribute(Constants.AT_AAM_ASSERTIONS_EXIST_OCCURRENCE, nLocale).getValue()) );
    objSheet.cell(55, 6).setCellValue( writeX(objControl.Attribute(Constants.AT_AAM_ASSERTIONS_NA, nLocale).getValue()) );
    objSheet.cell(57, 6).setCellValue( writeX(objControl.Attribute(Constants.AT_AAM_ASSERTIONS_VALUATION_ALLOCATION, nLocale).getValue()) );
    objSheet.cell(59, 6).setCellValue( writeX(objControl.Attribute(Constants.AT_AAM_ASSERTIONS_COMPLETENESS, nLocale).getValue()) );
    objSheet.cell(61, 6).setCellValue( writeX(objControl.Attribute(Constants.AT_AAM_ASSERTIONS_RIGHTS_OBLIGATIONS, nLocale).getValue()) );
    objSheet.cell(63, 6).setCellValue( writeX(objControl.Attribute(Constants.AT_AAM_ASSERTIONS_PRESENTATION_DISCLOSURE, nLocale).getValue()) );
}

function writeITControlTODSheet(objControl, objSheet)
{
    objSheet.cell(13, 3).setCellValue(getDomainList());
    objSheet.cell(14, 3).setCellValue(getHLCOList());
    objSheet.cell(15, 3).setCellValue(getLLCOList());
    
    objSheet.cell(20, 1).setCellValue(getRiskList());
    objSheet.cell(20, 2).setCellValue(clean(getControlDescription(objControl)));
    objSheet.cell(20, 4).setCellValue(getDateFormatForExcel(g_objVACD.Attribute(objDatumVerifikacijeattribute, nLocale).getValue()));
    objSheet.cell(20, 7).setCellValue(objControl.Attribute(Constants.AT_AAM_CTRL_ID, nLocale).getValue());
    objSheet.cell(20, 8).setCellValue(getControlExecutors(objControl));
    
    objSheet.cell(34, 5).setCellValue(objControl.Attribute(Constants.AT_AAM_CTRL_FREQUENCY, nLocale).getValue());
    objSheet.cell(36, 5).setCellValue(objControl.Attribute(objCheckEvidenceattribute, nLocale).getValue());
}

function writeITControlTOESheet(objControl, objSheet)
{
    objSheet.cell(13, 3).setCellValue(getDomainList());
    objSheet.cell(14, 3).setCellValue(getHLCOList());
    objSheet.cell(15, 3).setCellValue(getLLCOList());
    
    objSheet.cell(20, 1).setCellValue(getRiskList());
    objSheet.cell(20, 2).setCellValue(clean(getControlDescription(objControl)));
    objSheet.cell(20, 6).setCellValue(getDateFormatForExcel(g_objVACD.Attribute(objDatumVerifikacijeattribute, nLocale).getValue()));
    objSheet.cell(20, 7).setCellValue(objControl.Attribute(Constants.AT_AAM_CTRL_ID, nLocale).getValue());
    objSheet.cell(20, 8).setCellValue(getControlExecutors(objControl));
    
    objSheet.cell(25, 11).setCellValue(objControl.Attribute(Constants.AT_AAM_CTRL_ID, nLocale).getValue());
    
    objSheet.cell(32, 3).setCellValue(objControl.Attribute(Constants.AT_AAM_CTRL_FREQUENCY, nLocale).getValue());
}




function getISPProcessLevels(strLevel4Pocess)
{ 
    var listRows = g_objISPTree.getRows();
    var listProceses = new Array();
	listProceses[0] = "";
	listProceses[1] = "";
	listProceses[2] = "";
    var level4ExcelProces;
    
    for (var i = 0 ; i < listRows.length ; i++ )
    {
        level4ExcelProces = listRows[i].getCellAt(3).getCellValue().toString().replace(" ", "");
        
        if(  vbStrComp(level4ExcelProces,strLevel4Pocess) == 0)
        {
            listProceses[0] = listRows[i].getCellAt(0).getCellValue();
			listProceses[1] = listRows[i].getCellAt(1).getCellValue();
			listProceses[2] = listRows[i].getCellAt(2).getCellValue();
        }
    }
    
    return listProceses;
}
function writeX(strAttributeVal)
{
    if(strAttributeVal == "True")
        return "X";
    else
        return "";
}


function getCoveredRisks(objControlDef)
{
    var objTypes = new Array();
	objTypes[0] = Constants.OT_RISK;
    var conTypes = new Array();
	conTypes[0] = Constants.CT_IS_REDU_BY;
	
	return objControlDef.getConnectedObjs(objTypes, Constants.EDGES_INOUTASSIGN, conTypes);
	
}

function getRiskCategory(objControlDef)
{
    var objTypes = new Array();
	objTypes[0] = Constants.OT_RISK_CATEGORY;
    var conTypes = new Array();
	conTypes[0] = Constants.CT_BELONG_CAT;
    conTypes[1] = Constants.CT_BELONGS;
    conTypes[2] = Constants.CT_BELONGS_TO_1;
    conTypes[3] = Constants.CT_BELONGS_TO_2;
    conTypes[4] = Constants.CT_BELONGS_TO_3;
    conTypes[5] = Constants.CT_BELONGS_TO_4;
    conTypes[6] = Constants.CT_BELONGS_TO_5;
    conTypes[7] = Constants.CT_BELONGS_TO_6;
    conTypes[8] = Constants.CT_BELONGS_TO_7;
    conTypes[9] = Constants.CT_BELONGS_TO_8;
    conTypes[10] = Constants.CT_BELONGS_TO_9;
    conTypes[11] = Constants.CT_BELONGS_TO_10;
	
    conTypes[12] = Constants.CT_INCLUDES;
    conTypes[13] = Constants.CT_SUBS_1;
    conTypes[14] = Constants.CT_SUBS_2;
    conTypes[15] = Constants.CT_SUBS_3;
    conTypes[16] = Constants.CT_SUBS_4;
    conTypes[17] = Constants.CT_SUBS_5;
    
	return objControlDef.getConnectedObjs(objTypes, Constants.EDGES_IN, conTypes)[0];
	
}

function getCoveredRisksFunctions()
{
    var objRiskFuncList = new Array();
    var objTypes = new Array();
	objTypes[0] = Constants.OT_FUNC;
    var conTypes = new Array();
	conTypes[0] = Constants.CT_OCCUR;
	
    for (var i = 0 ; i < g_listCoveredRisks.length ; i++ )
    {
	     objRiskFuncList = g_listCoveredRisks[i].getConnectedObjs(objTypes, Constants.EDGES_INOUTASSIGN, conTypes);
         
         if(objRiskFuncList != null)
         {
             for (var j = 0 ; j < objRiskFuncList.length ; j++ )
            {
                g_listCoveredRiskFunctions[g_listCoveredRiskFunctions.length] = objRiskFuncList[j];
            }
         }
    }
	
}

function getControlExecutors(p_Object) 
{
	var outText = "";
    var objTypes = new Array();
    objTypes[0] = Constants.OT_ORG_UNIT;
    objTypes[1] = Constants.OT_POS;
    objTypes[2] = Constants.OT_PERS;
    objTypes[3] = Constants.OT_PERS_TYPE;
    objTypes[4] = Constants.OT_ORG_UNIT_TYPE;
    objTypes[5] = Constants.OT_GRP;
    var conTypes = new Array();
    conTypes[0] = Constants.CT_CONTR_TO_1;
    conTypes[1] = Constants.CT_CONTR_TO_2;
    conTypes[2] = Constants.CT_EXEC_1;
    conTypes[3] = Constants.CT_EXEC_2;
    //conTypes[4] = Constants.CT_EXEC_3; //Custom 20220201
    conTypes[4] = Constants.CT_EXEC_4;
    conTypes[5] = Constants.CT_EXEC_5;
	
	var objConnectedOrg = p_Object.getConnectedObjs(objTypes, Constants.EDGES_INOUTASSIGN, conTypes);
	
    for (var i = 0 ; i < objConnectedOrg.length ; i++ )
    {
        outText = outText + objConnectedOrg[i].Name(nLocale) + "\n\n";
    }  
    
    return outText.substring(0, outText.length-2);
}

function getRiskList()
{
    var strRiskNameList = "";
    
    if(g_listCoveredRisks != null)
    {
        for (var i = 0 ; i < g_listCoveredRisks.length ; i++ )
        {
            strRiskNameList = strRiskNameList + g_listCoveredRisks[i].Name(nLocale) + " (" + g_listCoveredRisks[i].Attribute(Constants.AT_AAM_RISK_ID, nLocale).getValue()  + ")" + "\n\n";
        }
    }
    
    return strRiskNameList.substring(0, strRiskNameList.length-2);
}
function getFunctList()
{
    var strRiskNameList = "";
    
    if(g_listCoveredRiskFunctions != null)
    {
        for (var i = 0 ; i < g_listCoveredRiskFunctions.length ; i++ )
        {
            strRiskNameList = strRiskNameList + g_listCoveredRiskFunctions[i].Name(nLocale) + "\n\n";
        }
    }
    
    return strRiskNameList.substring(0, strRiskNameList.length-2);
}

function getControlDescription(objControl)
{
    var strDesc = "";
    
    strDesc = objControl.Name(nLocale) + " (" + objControl.Attribute(Constants.AT_AAM_CTRL_ID, nLocale).getValue() + ")\n\n";
    strDesc = strDesc + objControl.Attribute(Constants.AT_DESC, nLocale).getValue();
    
    return strDesc;
}


function getLLCOList()
{
    var strRiskNameList = "";
    var oRiskCatName = "";
    
    if(g_listCoveredRisks != null)
    {
        g_LLCOList = new Array();
        
        for (var i = 0 ; i < g_listCoveredRisks.length ; i++ )
        {
            oRiskCatName = getLLCOCategoryName(g_listCoveredRisks[i]);
            if(oRiskCatName != "" && !isInList(oRiskCatName,g_LLCOList))
            {
                strRiskNameList = strRiskNameList + oRiskCatName + ", ";
                g_LLCOList[g_LLCOList.length] = oRiskCatName;
            }
        }
    }
    
    return strRiskNameList.substring(0, strRiskNameList.length-2);
}

function getLLCOCategoryName(oRisk)
{
    var oRiskCat = getRiskCategory(oRisk);
    
    if(oRiskCat != null)
    {
        if(vbStrComp( oRiskCat.Attribute(objTipCOBITCiljaAttribute, nLocale).getValue(), "LLCO") == 0)
            return oRiskCat.Name(nLocale) + " - " + oRiskCat.Attribute(Constants.AT_SHORT_DESC, nLocale).GetValue(true);
    }
    
    return "";
}
function getHLCOList()
{
    var strRiskNameList = "";
    var oRiskCatName = "";
    
    if(g_listCoveredRisks != null)
    {
        g_HLCOList = new Array();
        
        for (var i = 0 ; i < g_listCoveredRisks.length ; i++ )
        {
            oRiskCatName = getHLCOCategoryName(g_listCoveredRisks[i]);
            if(oRiskCatName != "" && !isInList(oRiskCatName,g_HLCOList))
            {
                strRiskNameList = strRiskNameList + oRiskCatName + ", ";
                g_HLCOList[g_HLCOList.length] = oRiskCatName;
            }
        }
    }
    
    return strRiskNameList.substring(0, strRiskNameList.length-2);
}

function getHLCOCategoryName(oRisk)
{
    var oRiskCat = getRiskCategory(oRisk);
    
    if(oRiskCat != null)
    {
        oRiskCat = getRiskCategory(oRiskCat);
        
        if(oRiskCat != null)
        {
            if(vbStrComp( oRiskCat.Attribute(objTipCOBITCiljaAttribute, nLocale).getValue(), "HLCO") == 0)
               return oRiskCat.Name(nLocale) + " - " + oRiskCat.Attribute(Constants.AT_SHORT_DESC, nLocale).GetValue(true);
    
        }
        else
            return "";
    }
    
    return "";
}

function getDomainList()
{
    var strRiskNameList = "";
    var oRiskCatName = "";
    
    if(g_listCoveredRisks != null)
    {
        g_DomainList = new Array();
        
        for (var i = 0 ; i < g_listCoveredRisks.length ; i++ )
        {
            oRiskCatName = getDomainCategoryName(g_listCoveredRisks[i]);
            if(oRiskCatName != "" && !isInList(oRiskCatName,g_DomainList))
            {
                strRiskNameList = strRiskNameList + oRiskCatName + ", ";
                g_DomainList[g_DomainList.length] = oRiskCatName;
            }
        }
    }
    
    return strRiskNameList.substring(0, strRiskNameList.length-2);
}

function getDomainCategoryName(oRisk)
{
    var oRiskCat = getRiskCategory(oRisk);
    
    if(oRiskCat != null)
    {
        oRiskCat = getRiskCategory(oRiskCat);
        
        if(oRiskCat != null)
        {
            oRiskCat = getRiskCategory(oRiskCat);
            
            if(oRiskCat != null)
            {
                //if(vbStrComp( oRiskCat.Attribute(objTipCOBITCiljaAttribute, nLocale).getValue(), "DOMAIN") == 0)
                    return oRiskCat.Name(nLocale) + " - " + oRiskCat.Attribute(Constants.AT_SHORT_DESC, nLocale).GetValue(true);
    
        
            }
            else
                return "";
        }
        else
            return "";
    }
    
    return "";
}





function isInList(strRiskCatName, list)
{
     for (var i = 0 ; i < list.length ; i++ )
     {
         if(vbStrComp(list[i], strRiskCatName) == 0)
             return true;
     }
     
     return false;
}


function getOccurenceInBCD(objDestDef)
{
        var occ = null;
        var epccount = 0;
        var oOccList = objDestDef.OccList();
        
        for (var i = 0 ; i < oOccList.length ; i++ )
        {
            if(oOccList[i].Model().TypeNum() == Constants.MT_BUSY_CONTR_DGM)
            {
                occ = oOccList[i];
                epccount ++;
            }
        }
       
        return occ;
}
function getOccurenceInEPC(objDestDef)
{
        var occ = null;
        var epccount = 0;
        var oOccList = objDestDef.OccList();
        
        for (var i = 0 ; i < oOccList.length ; i++ )
        {
            if(oOccList[i].Model().TypeNum() == Constants.MT_EEPC)
            {
                occ = oOccList[i];
                epccount ++;
            }
        }
       
        return occ;
}


function vbStrComp()
{
  var lhs = "";
  var rhs = "";
  var compMode = 0;

  if(arguments.length==2)
  {
    lhs = ""+arguments[0];
    rhs = ""+arguments[1];
  }
  else if(arguments.length==3)
  {
    lhs      = ""+arguments[0];
    rhs      = ""+arguments[1];
    compMode = 0+arguments[2];
  }

  if(lhs==null||rhs==null||compMode==null)
    return null;

  lhs = "" + lhs;
  rhs = "" + rhs;

  if(typeof(lhs) == "object")
	  lhs = lhs.toString();

  if(typeof(rhs) == "object")
	  rhs = rhs.toString();

  if( (typeof(lhs) != "string") || (typeof(rhs) != "string")|| (typeof(compMode)!="number"))
    return 0;

  if(compMode==0)
  {
    var i, c_lhs, c_rhs;
    for(i=0;i<lhs.length&&i<rhs.length;i++) {
      c_lhs = lhs.charCodeAt(i);
      c_rhs = rhs.charCodeAt(i);
      if(lhs!=rhs) {
        if(lhs < rhs) return -1
        return 1;
      }
    }
    if(i<lhs.length) {
      return -1;
    } else if(i<rhs.length) {
      return 1;
    }
    return 0;
  }
  else if(compMode==1)
  {
    var tmp_lhs = new java.lang.String(lhs.toLowerCase());
    var res = tmp_lhs.compareTo(new java.lang.String(rhs.toLowerCase()));
    return (res < 0) ? -1 : ((res > 0) ? 1 : 0);
  }
  else throw "argument error: vbStrComp expects compMode 0 or 1";
}



/** 
  *  Dialog function for output options dialog
  * @param dlgItem   dialog item name
  * @param act action
  * @param suppVal support value
  * @return bool value, for act!=1, true means continue dialog and false ends dialog, vice versa for act=1 
  */
function dlgFuncOutputOptions(dlgItem, act, suppVal)
{
    var bResult = true;
    
}


function showOutputOptionsDialog()
{
    var dicOutputFormat = "dlgTipKontrola";
  // Output format
  var userdialog = Dialogs.createNewDialogTemplate(0, 0, 385, 225, "Opcije izvještaja");

  userdialog.GroupBox(7, 0, 426, 55, "Odaberite tip procesa/kontrola");
  userdialog.OptionGroup(dicOutputFormat);
  userdialog.OptionButton(20, 15, 380, 15, "Administrativni");
  userdialog.OptionButton(20, 30, 380, 15, "COBIT");


  userdialog.GroupBox(7, 60, 426, 70, "Koje obrasce želite kreirati:");
  userdialog.CheckBox(20, 75, 220, 15, "TOD obrazac", "TOD");
  userdialog.CheckBox(20, 90, 220, 15, "TOE obrazac", "TOE");
  
  var lisValues = new Array();
  lisValues[0] = "Sve";
  lisValues[1] = "A";
  lisValues[2] = "S";
  
  userdialog.GroupBox(7, 130, 426, 50, "COBIT Pristup testiranju:");
  userdialog.DropListBox(20, 150, 200, 15, lisValues, "dlgTipCOBITCilja")
 
  userdialog.OKButton(20, 180, 75, 15);
  userdialog.CancelButton(160, 180, 75, 15);  
    
  dlgFuncOutput = Dialogs.createUserDialog(userdialog); 
  Dialogs.show(dlgFuncOutput)
  
  g_isStandardProcess = (dlgFuncOutput.getDlgValue(dicOutputFormat) == 0);
  g_generateTOD = (dlgFuncOutput.getDlgValue("TOD") == 1);
  g_generateTOE = (dlgFuncOutput.getDlgValue("TOE") == 1);
  
  var dlgTipCOBIT = dlgFuncOutput.getDlgValue("dlgTipCOBITCilja");
  
  if(dlgTipCOBIT == 0)
      g_tipCobitProcesa = "Sve";
  
  if(dlgTipCOBIT == 1)
      g_tipCobitProcesa = "A";
  
  if(dlgTipCOBIT == 2)
      g_tipCobitProcesa = "S";
  
}

function getDateFormatForExcel(strDateFromARIS)
{
    if(strDateFromARIS != "")
    {
    var mainParts = newsplitfunction(strDateFromARIS, ",");
    
    var strDateMonthPart = mainParts[0];
    var strYearPart = vbLTrim(mainParts[1]);
    
    var dateMonthParts = newsplitfunction(strDateMonthPart, " ");
    var strDate = dateMonthParts[1];
    
    var strMonth = "";
    if(dateMonthParts[0] == "Jan")
        strMonth = "01";
    if(dateMonthParts[0] == "Feb")
        strMonth = "02";
    if(dateMonthParts[0] == "Mar")
        strMonth = "03";
    if(dateMonthParts[0] == "Apr")
        strMonth = "04";
    if(dateMonthParts[0] == "May")
        strMonth = "05";
    if(dateMonthParts[0] == "Jun")
        strMonth = "06";
    if(dateMonthParts[0] == "Jul")
        strMonth = "07";
    if(dateMonthParts[0] == "Aug")
        strMonth = "08";
    if(dateMonthParts[0] == "Sep")
        strMonth = "09";
    if(dateMonthParts[0] == "Oct")
        strMonth = "10";
    if(dateMonthParts[0] == "Nov")
        strMonth = "11";
    if(dateMonthParts[0] == "Dec")
        strMonth = "12";
    
    return strDate + "." + strMonth + "." + strYearPart;
    }
    else
        return "";
    
}

function newsplitfunction(sstring, sseparator)
{
  // in VB implicitly declared local variables
  var stemp = undefined;
  var sprivatestring = undefined;
  var istart = undefined;
  //////

  var arefarray = __createArray(""); 
  arefarray = new Array(); // from VB: ReDim(0);
  sprivatestring = sstring;

  while (sprivatestring != "") {
    if (vbInStr(sprivatestring, sseparator) > 0) {
      stemp = vbLeft(sprivatestring, (vbInStr(sprivatestring, sseparator) - 1));
      istart = vbInStr(sprivatestring, sseparator);
      if (istart != vbLen(sprivatestring)) {
        sprivatestring = vbMid(sprivatestring, (istart + 1), (vbLen(sprivatestring) - istart));
      }
      else {
        sprivatestring = "";
      }

    }
    else {
      stemp = sprivatestring;
      sprivatestring = "";
    }


    if (stemp != "") {
      if (arefarray[__getArrayIndex(arefarray,1,0)] != "") {
        arefarray = __redimArray(arefarray, new __ArrayInfo(0, (__getUBound(arefarray, 1) + 1)), "", true); 
      }

      arefarray[__getArrayIndex(arefarray,1,__getUBound(arefarray, 1))] = stemp;
    }

  }


  return arefarray;
}
function clean(str)
{
    return str.replace("\r", "");
}
