var oOutput = Context.createOutputObject()
oOutput.DefineF("ID_STYLE_RD_HEADING_4", getString("ID_DEFAULT_FONT"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_ITALIC | Constants.FMT_BOLD | Constants.FMT_LEFT| Constants.FMT_VTOP| Constants.FMT_TOCENTRY3 , 0, 0, 0, 0, 0, 1)
oOutput.DefineF("ID_STYLE_RD_HEADING_2", getString("ID_DEFAULT_FONT"), 14, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_ITALIC | Constants.FMT_BOLD | Constants.FMT_LEFT| Constants.FMT_VTOP| Constants.FMT_TOCENTRY1 , 0, 0, 2, 2, 0, 1)
oOutput.DefineF("ID_STYLE_RD_DEFAULT", getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_LEFT| Constants.FMT_VTOP, 0, 0, 0, 0, 0, 1)
oOutput.DefineF("ID_STYLE_RD_INFO", getString("ID_DEFAULT_FONT"), 14, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_BOLD | Constants.FMT_CENTER| Constants.FMT_VTOP, 0, 0, 1.76, 8.82, 0, 1)
oOutput.DefineF("ID_STYLE_RD_HEADING_3", getString("ID_DEFAULT_FONT"), 12, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_ITALIC | Constants.FMT_BOLD | Constants.FMT_LEFT| Constants.FMT_VTOP| Constants.FMT_TOCENTRY2 , 0, 0, 1, 1, 0, 1)
oOutput.DefineF("ID_STYLE_RD_HEADING_1", getString("ID_DEFAULT_FONT"), 18, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_BOLD | Constants.FMT_LEFT| Constants.FMT_VTOP| Constants.FMT_TOCENTRY0 , 0, 0, 4, 4, 0, 1)
oOutput.DefineF("ID_STYLE_RD_TITLE", getString("ID_DEFAULT_FONT"), 21, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_BOLD | Constants.FMT_CENTER| Constants.FMT_VTOP, 0, 0, 1.76, 8.82, 0, 1)
oOutput.DefineF("ID_STYLE_RD_TABLE_CONTENT", getString("ID_DEFAULT_FONT"), 8, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_LEFT| Constants.FMT_VTOP, 0, 0, 0, 0, 0, 1)
oOutput.DefineF("ID_STYLE_RD_TABLE_HEAD", getString("ID_DEFAULT_FONT"), 8, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_BOLD | Constants.FMT_CENTER| Constants.FMT_VTOP, 0, 0, 0, 0, 0, 1)
oOutput.DefineF("ID_STYLE_RD_HEADER_FOOTER", getString("ID_DEFAULT_FONT"), 9, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_LEFT| Constants.FMT_VTOP, 0, 0, 0, 0, 0, 1)
setupOutputObject( oOutput )
oOutput.SetTitle(Context.getScriptInfo(Constants.SCRIPT_NAME))



var nLocale = Context.getSelectedLanguage();
var g_objCurrentRiskGroup;
createSection1(oOutput, ArisData.getSelectedObjOccs());
oOutput.WriteReport()

/** Apply default page format settings to output object
 * @param {Output} outputObj The output object
 */
function setupOutputObject(outputObj)
{
	outputObj.SetPageWidth(210)
	outputObj.SetPageHeight(297)
	outputObj.SetLeftMargin(20)
	outputObj.SetRightMargin(20)
	outputObj.SetTopMargin(30)
	outputObj.SetBottomMargin(30)
	outputObj.SetDistHeader(10)
	outputObj.SetDistFooter(10)
	outputObj.SetAutoTOCNumbering(true)
	globalHeader(outputObj)

	globalFooter(outputObj)
}
function globalHeader(outputObj) {
	outputObj.BeginHeader()
	outputObj.EndHeader()
}
function globalFooter(outputObj) {
	outputObj.BeginFooter()
	outputObj.EndFooter()
}

/**
 * @param {Output} p_output The output object
 * @param {ObjDef[]} p_aObjDef 
*/
function createSection1(p_output, selectedRiskGroups)
{
	// do not create new section if all data is empty
	if(selectedRiskGroups.length==0)
		return;

	p_output.BeginSection(false, Constants.SECTION_DEFAULT);
	setupOutputObject(p_output); //use defaults
    
	// repetition of queried data:
	for(var i=0; i<selectedRiskGroups.length; i++) 
    {
        g_objCurrentRiskGroup = selectedRiskGroups[i];
        var objRiskOccList = new Array();
        getRiskOccurances(selectedRiskGroups[i], objRiskOccList);
        sortRiskOccurancesByY(objRiskOccList);
        
		dtable(p_output, objRiskOccList);
	}
		
	p_output.EndSection();
}

function getRiskOccurances(objRiskCategoryOcc, objRiskOccList)
{
    var objConnectedOcc;
    
    var SymbolTypes = new Array();
    SymbolTypes[0] = Constants.ST_RISK;
    SymbolTypes[1] = Constants.ST_RISK_1;
    SymbolTypes[2] = Constants.ST_RISK_CATEGORY;
    SymbolTypes[3] = Constants.ST_RISK_PIC;
	
	var objConnectedObjectsOcc = objRiskCategoryOcc.getConnectedObjOccs(SymbolTypes, Constants.EDGES_OUT);
    
    for(var i=0; i<objConnectedObjectsOcc.length;i++)
    {
        objConnectedOcc = objConnectedObjectsOcc[i];
        
        if(objConnectedOcc.SymbolNum() == Constants.ST_RISK_CATEGORY)
            getRiskOccurances(objConnectedOcc, objRiskOccList)
        else
            objRiskOccList[objRiskOccList.length] = objConnectedOcc;
    }
}

function sortRiskOccurancesByY(objRiskOccList)
{
    var temp;
    
   for(var i=0; i<objRiskOccList.length-1;i++)
   {
       for(var j=i+1; j<objRiskOccList.length;j++)
       {
           if(objRiskOccList[j].Y() < objRiskOccList[i].Y())
           {
               temp = objRiskOccList[i];
               objRiskOccList[i] = objRiskOccList[j];
               objRiskOccList[j] = temp;
           }
       }
   }
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

function getOccuresAtActivity(objRiskDef)
{
    var objTypes = new Array();
	objTypes[0] = Constants.OT_FUNC;
	var conTypes = new Array();
	conTypes[0] = Constants.CT_OCCUR;
    conTypes[1] = Constants.CT_OCCURS_AT;
	
	var objConnectedObjects = objRiskDef.getConnectedObjs(objTypes, Constants.EDGES_INOUTASSIGN, conTypes);
    
    return objConnectedObjects[0];
}

function writeOrgUnitsPositions(p_output, p_Object) 
{
    var outText = "";
    var objTypes = new Array();
    objTypes[0] = Constants.OT_ORG_UNIT;
    objTypes[1] = Constants.OT_POS;
    objTypes[2] = Constants.OT_PERS;
    objTypes[3] = Constants.OT_PERS_TYPE;
    objTypes[4] = Constants.OT_ORG_UNIT_TYPE;
    var conTypes = new Array();
    conTypes[0] = Constants.CT_CONTR_TO_1;
    conTypes[1] = Constants.CT_CONTR_TO_2;
    conTypes[2] = Constants.CT_EXEC_1;
    conTypes[3] = Constants.CT_EXEC_2;
    //conTypes[4] = Constants.CT_EXEC_3;  //Custom 20220201
    conTypes[4] = Constants.CT_EXEC_4;
    conTypes[5] = Constants.CT_EXEC_5;
                
    var objConnectedOrg = p_Object.getConnectedObjs(objTypes, Constants.EDGES_INOUTASSIGN, conTypes);
    p_output.BeginParagraphF("ID_STYLE_RD_TABLE_CONTENT");
    for(var index=0; index<objConnectedOrg.length; index++) 
    {    
         outText = objConnectedOrg[index].Attribute(Constants.AT_NAME, nLocale).GetValue(true);
         p_output.OutputLn(outText, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
         if(index != (objConnectedOrg.length-1))
             p_output.OutputLn("", getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
                    
    }
    p_output.EndParagraph();
}

function getProcessPhaseName(objCurrControlDef)
{
    if(objCurrControlDef == null)
        return "";
    
    var objControlOccEPCList =  getModelListOfType(objCurrControlDef, Constants.MT_EEPC);
    
    if(objControlOccEPCList.length == 0)
        return "";
    else
        return objControlOccEPCList[0].Name(nLocale);
    
    /*
    var objRiskOccEPCList =  getModelListOfType(objRiskDef, Constants.MT_EEPC);
    
    for(var index=0; index<objRiskOccEPCList.length; index++) 
    {  
        if(objFunctionDef.OccListInModel(objRiskOccEPCList[index]).length > 0)
            return objRiskOccEPCList[index].Name(nLocale);
    }
    */
 
}

function getModelListOfType(objDef, enmModelType)
{
    var objModelList = new Array();
    var objOccList = objDef.OccList();
    
    for(var index=0; index<objOccList.length; index++) 
    {   
        if(objOccList[index].Model().TypeNum() == enmModelType)
            objModelList[objModelList.length] = objOccList[index].Model();
    }
    
    return objModelList;
}

function writeTableHeader(p_output)
{
    if(Context.getSelectedFormat() != Constants.OUTEXCEL)
    {
        p_output.BeginTable(100, Constants.C_BLACK, Constants.C_WHITE,  Constants.FMT_LEFT, 0)
        p_output.TableRow()
	    p_output.TableCell("RISK & CONTROL MATRIX", 100, getString("ID_DEFAULT_FONT"), 14, Constants.C_BLACK, Constants.C_YELLOW, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
        p_output.TableRow()
	    p_output.TableCell("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_WHITE, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
        p_output.TableRow()
	    p_output.TableCell("CPE Name", 10, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
        p_output.TableCell("Activity", 30, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
        p_output.TableCell("Potential Risk", 20, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
        p_output.TableCell("Control", 40, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
    }
    else
    {
        p_output.BeginTable(100, Constants.C_BLACK, Constants.C_WHITE,  Constants.FMT_LEFT, 0)
        p_output.TableRow()
	    p_output.TableCell("RISK & CONTROL MATRIX", 100, getString("ID_DEFAULT_FONT"), 14, Constants.C_BLACK, Constants.C_YELLOW, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
        p_output.TableRow()
	    p_output.TableCell("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_WHITE, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
        p_output.TableRow()
	    p_output.TableCell("CPE Name", 10, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
        p_output.TableCell("Activity", 15, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
        p_output.TableCell("", 15, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
        p_output.TableCell("Potential Risk", 5, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
        p_output.TableCell("", 15, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
        p_output.TableCell("Control", 5, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
        p_output.TableCell("", 15, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
        p_output.TableCell("", 15, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
        p_output.TableCell("", 5, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
    }
    
    p_output.TableRow()
	p_output.TableCell("", 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
    p_output.TableCell("Description", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
	p_output.TableCell("Activity Owner", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
	p_output.TableCell("Risk ID", 5, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
    p_output.TableCell("Description", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
	p_output.TableCell("Control ID", 5, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
    p_output.TableCell("Description", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
    p_output.TableCell("Control Owner", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
    p_output.TableCell("Key control", 5, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
	
}
function dtable(p_output, objRiskOccList)
{
    var textRisk;
    var textControl;
	// do not create new table if all data is empty
	if(objRiskOccList.length==0)
		return;
    
	//Table Header
    writeTableHeader(p_output);
    
    // repetition of queried data:
	for(var i=0; i<objRiskOccList.length; i++)
    {
        var objCurrRiskDef = objRiskOccList[i].ObjDef();
        var objControlDefList = getRiskControls(objCurrRiskDef);
        var objRiskActivity;
        var txtRiskActivityName;
        var blnIsKeyControl;
        var strProcessPhaseName;
        
        for(var j=0; j<objControlDefList.length; j++)
        {     
            objRiskActivity = getOccuresAtActivity(objCurrRiskDef);
            strProcessPhaseName = getProcessPhaseName(objControlDefList[j]); 
            
		    p_output.TableRow();
            
            p_output.TableCell(strProcessPhaseName, 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0)
            
            /*
            if(i==0)
                p_output.TableCell(g_objCurrentRiskGroup.ObjDef().Attribute(Constants.AT_NAME,nLocale).GetValue(true), 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0)
            else
                p_output.TableCell("", 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0)
            */
            
            if(objRiskActivity != null)
            {
                txtRiskActivityName = objRiskActivity.Attribute(Constants.AT_NAME,nLocale).GetValue(true);
                p_output.TableCell(txtRiskActivityName , 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0);
                p_output.TableCell("" , 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0);
                writeOrgUnitsPositions(p_output, objRiskActivity);
            }
            else
            {
                txtRiskActivityName = "";
                p_output.TableCell(txtRiskActivityName , 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0);
                p_output.TableCell(txtRiskActivityName , 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_BOLD, 0);
            }
            
            
		    p_output.TableCell(objCurrRiskDef.Attribute(Constants.AT_AAM_RISK_ID,nLocale).GetValue(true), 5, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_UNDERLINE, 0);
            
            p_output.TableCell("", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT, 0)
            p_output.BeginParagraphF("ID_STYLE_RD_TABLE_CONTENT")
			p_output.OutputLn(objCurrRiskDef.Attribute(Constants.AT_NAME,nLocale).GetValue(true), getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_VTOP | Constants.FMT_LEFT | Constants.FMT_BOLD, 0)
            p_output.OutputLn("", getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
            p_output.OutputLn(objCurrRiskDef.Attribute(Constants.AT_DESC,nLocale).GetValue(true), getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
		    p_output.EndParagraph();
           
		    p_output.TableCell(objControlDefList[j].Attribute(Constants.AT_AAM_CTRL_ID,nLocale).GetValue(true), 5, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER | Constants.FMT_UNDERLINE, 0);
            
		    p_output.TableCell("", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
            p_output.BeginParagraphF("ID_STYLE_RD_TABLE_CONTENT")
			p_output.OutputLn(objControlDefList[j].Attribute(Constants.AT_NAME,nLocale).GetValue(true), getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_VTOP | Constants.FMT_LEFT | Constants.FMT_BOLD, 0)
            p_output.OutputLn("", getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
            p_output.OutputLn(objControlDefList[j].Attribute(Constants.AT_DESC,nLocale).GetValue(true), getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
		    p_output.EndParagraph()
            
            p_output.TableCell("" , 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_BOLD, 0);
            writeOrgUnitsPositions(p_output, objControlDefList[j]);
            
            blnIsKeyControl = objControlDefList[j].Attribute(Constants.AT_AAM_KEY_CTRL,nLocale).GetValue(true);
            if(blnIsKeyControl == "True")
                p_output.TableCell("K" , 5, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0);
            else
            {
                if(blnIsKeyControl == "False")
                    p_output.TableCell("NK" , 5, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0);
                else
                    p_output.TableCell("" , 5, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0);
            }
        }
	}

	p_output.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
}


function RGB(r, g, b) {
	return (new java.awt.Color(r/255.0,g/255.0,b/255.0,1)).getRGB() & 0xFFFFFF
}
