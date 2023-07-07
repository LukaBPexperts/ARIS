var nLocale = Context.getSelectedLanguage();
var gActiveDatabase = ArisData.getActiveDatabase();
var g_objVACD;
var g_objEPC;
var g_objRootRiskCat;
var g_objCurrentControl;
var g_objCurrentRisk;
var g_objCurrentRiskFunction;
var g_listCoveredRisks = new Array();
var g_listCoveredRiskFunctions = new Array();
var g_listControls = new Array();
var objXlsWorkBookRCM;


var database = ArisData.getActiveDatabase();
var filter = database.ActiveFilter();
var objL262attribute = filter.UserDefinedAttributeTypeNum("eb39c830-88cd-11df-453f-005056b228c7");
var objISPProces4attribute = filter.UserDefinedAttributeTypeNum("e03dc090-88d1-11df-453f-005056b228c7");
var objDatumVerifikacijeattribute = filter.UserDefinedAttributeTypeNum("950e02b0-88d1-11df-453f-005056b228c7");
var objCheckEvidenceattribute = filter.UserDefinedAttributeTypeNum("317ba912-7df8-11da-0c60-cf8e338f9a0b");
var objPristupTestiranjuAttribute = filter.UserDefinedAttributeTypeNum("c05f4e80-88d4-11df-453f-005056b228c7");
var objTipCOBITCiljaAttribute = filter.UserDefinedAttributeTypeNum("7a9da400-88d4-11df-453f-005056b228c7");

var objImpactAttribute = filter.UserDefinedAttributeTypeNum("0159fb10-88d5-11df-453f-005056b228c7");
var objProbabilityAttribute = filter.UserDefinedAttributeTypeNum("29802600-88d5-11df-453f-005056b228c7");


main()


function main()
{
    var fileTemplateContent;
    
    var selClientFolders = Dialogs.getClientFolder("Odaberite gdje želite spremiti generirani obrazac:", "", false);
    Context.setSelectedPath( selClientFolders[0] ) ;
    g_objRootRiskCat = ArisData.getSelectedObjDefs()[0];
    
    Context.setSelectedFile(g_objRootRiskCat.Name(nLocale) + "_RCM.xls");
    fileTemplateContent = Context.getFile("RCM - predložak.xls", Constants.LOCATION_SCRIPT)
	objXlsWorkBookRCM = Context.createExcelWorkbook(Context.getSelectedFile(), fileTemplateContent);
	   
	getRiskCategoryRisks(g_objRootRiskCat);
    
    objXlsWorkBookRCM.getSheetAt(0).cell(9, 1).setCellValue(g_objRootRiskCat.Name(nLocale));
    
    var listControls = null;
    var counter = 0;
    var objControlOcc = null;
    
    g_listCoveredRisks =  ArisData.sort(g_listCoveredRisks, Constants.AT_AAM_RISK_ID, nLocale);
    for(var i=0; i<g_listCoveredRisks.length; i++)
    {
        g_objCurrentRisk = g_listCoveredRisks[i];
        listControls = getRiskControls(g_listCoveredRisks[i]);
        listControls =  ArisData.sort(listControls, Constants.AT_AAM_CTRL_ID, nLocale);
        if(listControls != null && listControls.length > 0)
        {
            for(var j=0; j<listControls.length; j++)
            {
                g_objCurrentControl = listControls[j];
                objControlOcc = getOccurenceInEPC(g_objCurrentControl);
                if(objControlOcc != null)
                {
                    g_objEPC = objControlOcc.Model();
                    if(g_objEPC != null)
                        g_objVACD = getVACD(g_objEPC);
                }
                
                writeRCMRow(objXlsWorkBookRCM.getSheetAt(0), counter);
                counter ++;
            }
        }
    }
    
    
    
    objXlsWorkBookRCM.getSheetAt(0).cell(20 + counter, 0).setCellValue("Type");
    objXlsWorkBookRCM.getSheetAt(0).cell(20 + counter, 1).setCellValue("Reference");
    objXlsWorkBookRCM.getSheetAt(0).cell(20 + counter, 2).setCellValue("Activity Owner");
    objXlsWorkBookRCM.getSheetAt(0).addRegion(2, 3, 20 + counter,20 + counter);
    objXlsWorkBookRCM.getSheetAt(0).cell(20 + counter, 4).setCellValue("Anomaly description");
    objXlsWorkBookRCM.getSheetAt(0).addRegion(4, 21, 20 + counter,20 + counter);
    objXlsWorkBookRCM.getSheetAt(0).cell(20 + counter, 22).setCellValue("Rem. Plan re.");
    
    objXlsWorkBookRCM.getSheetAt(0).cell(22 + counter, 0).setCellValue("");
    objXlsWorkBookRCM.getSheetAt(0).cell(22 + counter, 1).setCellValue("");
    objXlsWorkBookRCM.getSheetAt(0).cell(22 + counter, 2).setCellValue("");
    objXlsWorkBookRCM.getSheetAt(0).addRegion(2, 3, 22 + counter,22 + counter);
    objXlsWorkBookRCM.getSheetAt(0).cell(22 + counter, 4).setCellValue("");
    objXlsWorkBookRCM.getSheetAt(0).addRegion(4, 21, 22 + counter,22 + counter);
    objXlsWorkBookRCM.getSheetAt(0).cell(22 + counter, 22).setCellValue("");
    
    var headerStyle = objXlsWorkBookRCM.getSheetAt(0).getRowAt(16).getCellAt(0).getCellStyle();
    var normalStyle = objXlsWorkBookRCM.getSheetAt(0).getRowAt(18).getCellAt(12).getCellStyle();
    
    for(var i=0; i<23; i++)
    {
    objXlsWorkBookRCM.getSheetAt(0).cell(20 + counter, i).setCellStyle(headerStyle);
    }
    for(var i=0; i<23; i++)
    {
    objXlsWorkBookRCM.getSheetAt(0).cell(22 + counter, i).setCellStyle(normalStyle);
    }
    
    objXlsWorkBookRCM.write();
       
    
}


function writeRCMRow(objSheet, index)
{   
    
    
    if(g_objEPC != null)
    {
        objSheet.cell(18 + index, 0).setCellValue(g_objEPC.Name(nLocale));
        
    }
    else
        objSheet.cell(18 + index, 0).setCellValue("");
    objSheet.cell(18 + index, 0).setCellStyle(CellStyle(0));
    
    if(g_objVACD!=null)
        objSheet.cell(18 + index, 1).setCellValue(getDateFormatForExcel(g_objVACD.Attribute(objDatumVerifikacijeattribute, nLocale).getValue()));
    else
        objSheet.cell(18 + index, 1).setCellValue("");
    objSheet.cell(18 + index, 1).setCellStyle(CellStyle(1));
    
    objSheet.cell(18 + index, 2).setCellValue(getRiskFunctionName());
    objSheet.cell(18 + index, 2).setCellStyle(CellStyle(2));
    
    if(g_objCurrentRiskFunction == null)
        objSheet.cell(18 + index, 3).setCellValue("");
    else
        objSheet.cell(18 + index, 3).setCellValue(getControlExecutors(g_objCurrentRiskFunction));
    objSheet.cell(18 + index, 3).setCellStyle(CellStyle(3));
    
    objSheet.cell(18 + index, 4).setCellValue(g_objCurrentRisk.Attribute(Constants.AT_AAM_RISK_ID, nLocale).getValue());
    objSheet.cell(18 + index, 4).setCellStyle(CellStyle(4));
    
    objSheet.cell(18 + index, 5).setCellValue(clean(g_objCurrentRisk.Attribute(Constants.AT_NAME, nLocale).GetValue(true)) + "\n\n" + clean(g_objCurrentRisk.Attribute(Constants.AT_DESC, nLocale).GetValue(true)) );
    objSheet.cell(18 + index, 5).setCellStyle(CellStyle(5));
    
    objSheet.cell(18 + index, 6).setCellValue( writeX(g_objCurrentControl.Attribute(Constants.AT_AAM_ASSERTIONS_EXIST_OCCURRENCE, nLocale).getValue()) );
    objSheet.cell(18 + index, 6).setCellStyle(CellStyle(6));
    
    objSheet.cell(18 + index, 7).setCellValue( writeX(g_objCurrentControl.Attribute(Constants.AT_AAM_ASSERTIONS_NA, nLocale).getValue()) );
    objSheet.cell(18 + index, 7).setCellStyle(CellStyle(7));
    
    objSheet.cell(18 + index, 8).setCellValue( writeX(g_objCurrentControl.Attribute(Constants.AT_AAM_ASSERTIONS_VALUATION_ALLOCATION, nLocale).getValue()) );
    objSheet.cell(18 + index, 8).setCellStyle(CellStyle(8));
     
    objSheet.cell(18 + index, 9).setCellValue( writeX(g_objCurrentControl.Attribute(Constants.AT_AAM_ASSERTIONS_COMPLETENESS, nLocale).getValue()) );
    objSheet.cell(18 + index, 9).setCellStyle(CellStyle(9));
    
    objSheet.cell(18 + index, 10).setCellValue( writeX(g_objCurrentControl.Attribute(Constants.AT_AAM_ASSERTIONS_RIGHTS_OBLIGATIONS, nLocale).getValue()) );
    objSheet.cell(18 + index, 10).setCellStyle(CellStyle(10));
    
    objSheet.cell(18 + index, 11).setCellValue( writeX(g_objCurrentControl.Attribute(Constants.AT_AAM_ASSERTIONS_PRESENTATION_DISCLOSURE, nLocale).getValue()) );
    objSheet.cell(18 + index, 11).setCellStyle(CellStyle(11));
    
    objSheet.cell(18 + index, 12).setCellValue("");
    objSheet.cell(18 + index, 12).setCellStyle(CellStyle(12));
    
    objSheet.cell(18 + index, 13).setCellValue(g_objCurrentRisk.Attribute(objImpactAttribute, nLocale).getValue());
    objSheet.cell(18 + index, 13).setCellStyle(CellStyle(13));
    
    objSheet.cell(18 + index, 14).setCellValue(g_objCurrentRisk.Attribute(objProbabilityAttribute, nLocale).getValue());
    objSheet.cell(18 + index, 14).setCellStyle(CellStyle(14));
    
    objSheet.cell(18 + index, 15).setCellValue(getGRE());
    objSheet.cell(18 + index, 15).setCellStyle(CellStyle(15));
    
    objSheet.cell(18 + index, 16).setCellValue("");
    objSheet.cell(18 + index, 16).setCellStyle(CellStyle(16));
    
    objSheet.cell(18 + index, 17).setCellValue(g_objCurrentControl.Attribute(Constants.AT_AAM_CTRL_ID, nLocale).getValue());
    objSheet.cell(18 + index, 17).setCellStyle(CellStyle(17));
    
    objSheet.cell(18 + index, 18).setCellValue(clean(g_objCurrentControl.Attribute(Constants.AT_NAME, nLocale).GetValue(true)) + "\n\n" + clean(g_objCurrentControl.Attribute(Constants.AT_DESC, nLocale).GetValue(true)) );
    objSheet.cell(18 + index, 18).setCellStyle(CellStyle(18));
    
    objSheet.cell(18 + index, 19).setCellValue(getControlExecutors(g_objCurrentControl));
    objSheet.cell(18 + index, 19).setCellStyle(CellStyle(19));
    
    objSheet.cell(18 + index, 20).setCellValue(getKeyNotKey());
    objSheet.cell(18 + index, 20).setCellStyle(CellStyle(20));
    
    objSheet.cell(18 + index, 21).setCellValue("Adequate");
    objSheet.cell(18 + index, 21).setCellStyle(CellStyle(21));
    
    objSheet.cell(18 + index, 22).setCellValue("Achieved");
    objSheet.cell(18 + index, 22).setCellStyle(CellStyle(22));
    
    objSheet.cell(18 + index, 23).setCellValue("Adequate Design");
    objSheet.cell(18 + index, 23).setCellStyle(CellStyle(23));
    
    objSheet.cell(18 + index, 24).setCellValue("");
    objSheet.cell(18 + index, 24).setCellStyle(CellStyle(24));
    
}

function getKeyNotKey()
{
    var val = g_objCurrentControl.Attribute(Constants.AT_AAM_KEY_CTRL, nLocale).getValue();
    
    if(vbStrComp(val, "True") == 0)
        return "K";
    
    if(vbStrComp(val, "False") == 0)
        return "NK";
    
    return "";
}

function clean(str)
{
    return str.replace("\r", "");
}
function getVACD(objEPC)
{
  var objSupObjDefs = objEPC.getSuperiorObjDefs();
  var objVACDOcc = null;
  
  for(var i=0; i<objSupObjDefs.length;i++)
  {
      if(objSupObjDefs[i].TypeNum() == Constants.OT_FUNC)
      {
          objVACDOcc = getOccurenceInVACD(objSupObjDefs[i]);
          
          if(objVACDOcc != null)
              return objVACDOcc.Model();
      }
  }
  
  return null;
}

function CellStyle(column)
{
    return objXlsWorkBookRCM.getSheetAt(0).getRowAt(18).getCellAt(column).getCellStyle();
}

function getGRE()
{
    var impact = g_objCurrentRisk.Attribute(objImpactAttribute, nLocale).getValue();
    var probability = g_objCurrentRisk.Attribute(objProbabilityAttribute, nLocale).getValue();
    
    if(vbStrComp(impact, "Significant") == 0)
    {
        if(vbStrComp(probability, "Low") == 0)
            return "M";
        if(vbStrComp(probability, "Medium") == 0)
            return "H";
        if(vbStrComp(probability, "High") == 0)
            return "H";
    }
    
    if(vbStrComp(impact, "Average") == 0)
    {
        if(vbStrComp(probability, "Low") == 0)
            return "L";
        if(vbStrComp(probability, "Medium") == 0)
            return "M";
        if(vbStrComp(probability, "High") == 0)
            return "H";
    }
    
    if(vbStrComp(impact, "Insignificant") == 0)
    {
        if(vbStrComp(probability, "Low") == 0)
            return "L";
        if(vbStrComp(probability, "Medium") == 0)
            return "L";
        if(vbStrComp(probability, "High") == 0)
            return "M";
    }
    
    return "";
}

function getRiskControls(objRiskDef)
{
    var objTypes = new Array();
	objTypes[0] = Constants.OT_FUNC;
	var conTypes = new Array();
	conTypes[0] = Constants.CT_IS_REDU_BY;
	
	var objConnectedObjects = objRiskDef.getConnectedObjs(objTypes, Constants.EDGES_OUT, conTypes);
    
    return objConnectedObjects;
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
function getRiskCategoryRisks(objRiskCatDef)
{
    var objTypes = new Array();
	objTypes[0] = Constants.OT_RISK;
    objTypes[1] = Constants.OT_RISK_CATEGORY;
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
    //conTypes[10] = Constants.CT_BELONGS_TO_9; //ANTE
    conTypes[10] = Constants.CT_BELONGS_TO_10;
	
    conTypes[11] = Constants.CT_INCLUDES;
    conTypes[12] = Constants.CT_SUBS_1;
    conTypes[13] = Constants.CT_SUBS_2;
    conTypes[14] = Constants.CT_SUBS_3;
    conTypes[15] = Constants.CT_SUBS_4;
    conTypes[16] = Constants.CT_SUBS_5;
	
	var objConnectedObjectsOcc = objRiskCatDef.getConnectedObjs(objTypes, Constants.EDGES_OUT, conTypes);
    var objConnectedOcc;
    
    for(var i=0; i<objConnectedObjectsOcc.length;i++)
    {
        objConnectedOcc = objConnectedObjectsOcc[i];
        
        if(objConnectedOcc.TypeNum() == Constants.OT_RISK_CATEGORY)
            getRiskCategoryRisks(objConnectedOcc)
        else
            g_listCoveredRisks[g_listCoveredRisks.length] = objConnectedOcc;
    }
	
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

function getRiskFunctionName()
{
    var objRiskFuncList = new Array();
    var objTypes = new Array();
	objTypes[0] = Constants.OT_FUNC;
    var conTypes = new Array();
	conTypes[0] = Constants.CT_OCCUR;
	
    objRiskFuncList = g_objCurrentRisk.getConnectedObjs(objTypes, Constants.EDGES_INOUTASSIGN, conTypes);
    
    if(objRiskFuncList.length == 0)
    {
        g_objCurrentRiskFunction = null;
        return "";
    }
    else
    {
        g_objCurrentRiskFunction = objRiskFuncList[0];
        return objRiskFuncList[0].Attribute(Constants.AT_NAME, nLocale).GetValue(true);
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
    //conTypes[4] = Constants.CT_EXEC_3; //ANTE
    conTypes[4] = Constants.CT_EXEC_4;
    conTypes[5] = Constants.CT_EXEC_5;
	
	var objConnectedOrg = p_Object.getConnectedObjs(objTypes, Constants.EDGES_INOUTASSIGN, conTypes);
	
    for (var i = 0 ; i < objConnectedOrg.length ; i++ )
    {
        outText = outText + objConnectedOrg[i].Name(nLocale) + ",";
    }  
    
    return outText.substring(0, outText.length-1);
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

function getOccurenceInVACD(objDestDef)
{
        var occ = null;
        var epccount = 0;
        var oOccList = objDestDef.OccList();
        
        for (var i = 0 ; i < oOccList.length ; i++ )
        {
            if(oOccList[i].Model().TypeNum() == Constants.MT_VAL_ADD_CHN_DGM)
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

