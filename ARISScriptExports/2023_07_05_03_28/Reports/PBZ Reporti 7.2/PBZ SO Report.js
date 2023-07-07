var g_atributetypenumstoignore = new Array();
var g_modelatributetypenumstoignore = new Array();
var g_VACatributetypenumstoignore = new Array();
var g_POSatributetypenumstoignore = new Array();
var g_APPatributetypenumstoignore = new Array();
var g_atributetypenums = new Array();
var g_modelatributetypenums = new Array(); 
var g_VACatributetypenums = new Array();
var g_POSatributetypenums = new Array(); 
var g_APPatributetypenums = new Array();  

__usertype_tsettings = function() {
  
  this.bConfigureModAtrToIgnore = false;
  this.bConfigureObjAtrToIgnore = false;
  this.bConfigureVACAtrToIgnore = false;
  this.bConfigureAPPAtrToIgnore = false;
  this.bConfigurePOSAtrToIgnore = false;
  
  this.__createNew = function() { return new __usertype_tsettings; }
}
var settings = new __holder(new __usertype_tsettings());

var nLocale = Context.getSelectedLanguage();
var gActiveDatabase = ArisData.getActiveDatabase();
var g_selectedGroups = ArisData.getSelectedGroups();
var g_selectedGroup;
var g_rowCount = 3;
var g_currentVACD;
var g_currentPhaseFunct;
var g_currentEPC;
var g_currentLevel1 = " ";
var g_currentLevel2 = " ";
var g_currentLevel3 = " ";
var g_currentLevel4 = " ";
var g_currentLevel5 = " ";
var g_currentLevel6 = " ";
var g_currentLevel7 = " ";
var g_currentLevel8 = " ";
var g_currentLevel9 = " ";
var g_currentLevel10 = " ";

var g_orgLevel1 = "";
var g_orgLevel2 = "";
var g_orgLevel3 = "";
var g_orgLevel4 = "";
var g_orgLevel5 = "";
var g_orgLevel6 = "";

if(displayui(settings.value) != 0)
{
Context.setSelectedFile("SO Total Report.xls");
fileTemplateContent = Context.getFile("Template.xls", Constants.LOCATION_SCRIPT)
var g_objXlsWorkBook = Context.createExcelWorkbook(Context.getSelectedFile(), fileTemplateContent);
var g_objSheet = g_objXlsWorkBook.getSheets()[0];

writeAdditionalHeaderColumns();
main();
}

function main()
{
    for (var gr = 0;gr<g_selectedGroups.length;gr++)
    {
    g_selectedGroup = g_selectedGroups[gr];
    
    var listGroupVACDs = g_selectedGroup.ModelList(true, getTypeNumList(Constants.MT_VAL_ADD_CHN_DGM));
    var listEPC;
    //var listEPC = g_selectedGroup.ModelList(true, getTypeNumList(Constants.MT_EEPC));
    var listEPCFunctions = new Array();
    
    for (var v = 0;v<listGroupVACDs.length;v++)
    {
    listEPC = getAssignedEPCs(listGroupVACDs[v]);    
        
    for (var i = 0;i<listEPC.length;i++)
    {
        //Set current context
        g_currentEPC = listEPC[i];
        g_currentVACD = listGroupVACDs[v];
        getCurrentLevelNames(g_currentVACD);
        //Set current context
        
        listEPCFunctions = listEPC[i].ObjOccListBySymbol(getTypeNumList3(Constants.ST_FUNC, Constants.ST_SYS_FUNC_ACT, 131407));
        
        for (var j = 0;j<listEPCFunctions.length;j++)
        {
            var listApplications = getApplications(listEPCFunctions[j].ObjDef());
            var listPositions = getPositions(listEPCFunctions[j].ObjDef());
            
            if(listApplications.length <= 1 && listPositions.length <= 1)
            {
                var oApp = null;
                var oPos = null;
                if(listApplications.length > 0)
                    oApp = listApplications[0];
                if(listPositions.length > 0)
                    oPos = listPositions[0];
                
                writeFunction(listEPCFunctions[j].ObjDef(), oPos, oApp);
            }
            else
            {
                if(listApplications.length == 0)
                {
                    for (var p = 0;p<listPositions.length;p++)
                    {
                        writeFunction(listEPCFunctions[j].ObjDef(), listPositions[p], null);
                    }
                }
                
                if(listPositions.length == 0)
                {
                    for (var a = 0;a<listApplications.length;a++)
                    {
                        writeFunction(listEPCFunctions[j].ObjDef(), null, listApplications[a]);
                    }
                }
                
                if(listApplications.length >= 1 && listPositions.length >= 1)
                {
                    for (var p = 0;p<listPositions.length;p++)
                    {
                        for (var a = 0;a<listApplications.length;a++)
                        {
                            writeFunction(listEPCFunctions[j].ObjDef(), listPositions[p], listApplications[a]);
                        }
                    }
                }
            }
        }
    }
    }//vacds
    }//groups
    
    g_objXlsWorkBook.write();
}
function getVACDFunction(oEPC)
{
    var listSuperiorFuncts = oEPC.getSuperiorObjDefs();
    var occInVACD = null;
    
    for (var i = 0;i<listSuperiorFuncts.length;i++)
    {
        if(listSuperiorFuncts[i].OccListInModel(g_currentVACD).length > 0)
            return listSuperiorFuncts[i];
        
    }
    
    return listSuperiorFuncts[0];
}

function writeFunction(objFunction, oPosition, oApplication)
{   
    if(oPosition!= null && oPosition.TypeNum() == Constants.OT_GRP)
    {
        var listGroupPositions = getGroupPositions(oPosition);
        
        if(listGroupPositions.length > 0)
        {
            for (var i = 0;i<listGroupPositions.length;i++)
            {
                writeFunctionForGroup(objFunction, oPosition, listGroupPositions[i], oApplication);
            }
        }
        
        //Group org. units
        var listGroupOrgUnits = getGroupOrgUnits(oPosition);
        if(listGroupOrgUnits.length > 0)
        {
            for (var i = 0;i<listGroupOrgUnits.length;i++)
            {
                writeFunctionForOrgUnit(objFunction, oPosition, listGroupOrgUnits[i], oApplication);
            }
        }
        
        if(listGroupOrgUnits.length == 0 && listGroupPositions.length == 0)
            writeFunctionForGroup(objFunction, oPosition, null, oApplication);
        
         return 1;
    }
    
    
    g_objSheet.cell(g_rowCount, 0).setCellValue(g_currentLevel1);
    g_objSheet.cell(g_rowCount, 1).setCellValue(g_currentLevel2);
    g_objSheet.cell(g_rowCount, 2).setCellValue(g_currentLevel3);
    g_objSheet.cell(g_rowCount, 3).setCellValue(g_currentLevel4);
    g_objSheet.cell(g_rowCount, 4).setCellValue(g_currentLevel5);
    g_objSheet.cell(g_rowCount, 5).setCellValue(g_currentLevel6);
    g_objSheet.cell(g_rowCount, 6).setCellValue(g_currentLevel7);
    g_objSheet.cell(g_rowCount, 7).setCellValue(g_currentLevel8);
    g_objSheet.cell(g_rowCount, 8).setCellValue(g_currentLevel9);
    g_objSheet.cell(g_rowCount, 9).setCellValue(g_currentLevel10);
    
    //VACD
    if(g_currentVACD != null)
    {
     g_objSheet.cell(g_rowCount, 10).setCellValue(g_currentVACD.Name(nLocale));
     g_objSheet.cell(g_rowCount, 11).setCellValue(g_currentVACD.GUID().toString());
     g_objSheet.cell(g_rowCount, 12).setCellValue(g_currentVACD.Attribute(133845, nLocale).GetValue(false));
     g_objSheet.cell(g_rowCount, 13).setCellValue(g_currentVACD.Attribute(264917, nLocale).GetValue(false));
     g_objSheet.cell(g_rowCount, 14).setCellValue(g_currentVACD.Attribute(68314, nLocale).GetValue(false));
     g_objSheet.cell(g_rowCount, 15).setCellValue(g_currentVACD.Attribute(527061, nLocale).GetValue(false));
     g_objSheet.cell(g_rowCount, 16).setCellValue(g_currentVACD.Attribute(264919, nLocale).GetValue(false));
     g_objSheet.cell(g_rowCount, 17).setCellValue(writeLinks(getVACDLinks(g_currentVACD)));
    }
    //VACD
    
    //EPC
    g_objSheet.cell(g_rowCount, 18).setCellValue(g_currentEPC.Name(nLocale));
    g_objSheet.cell(g_rowCount, 19).setCellValue(g_currentEPC.GUID().toString());
    g_objSheet.cell(g_rowCount, 20).setCellValue(getVACDFunction(g_currentEPC).Attribute(989, nLocale).GetValue(false));
    
    //ORG Dio
    if(oPosition != null)
    {
        var oOrgDio = getOrgDio(oPosition);
        getOrgLevels(oOrgDio);
        //write levels
        g_objSheet.cell(g_rowCount, 21).setCellValue(g_orgLevel1);
        g_objSheet.cell(g_rowCount, 22).setCellValue(g_orgLevel2);
        g_objSheet.cell(g_rowCount, 23).setCellValue(g_orgLevel3);
        g_objSheet.cell(g_rowCount, 24).setCellValue(g_orgLevel4);
        g_objSheet.cell(g_rowCount, 25).setCellValue(g_orgLevel5);
        g_objSheet.cell(g_rowCount, 26).setCellValue(g_orgLevel6);
        g_orgLevel1 = "";
        g_orgLevel2 = "";
        g_orgLevel3 = "";
        g_orgLevel4 = "";
        g_orgLevel5 = "";
        g_orgLevel6 = "";
        
        if(oOrgDio != null)
        {
            g_objSheet.cell(g_rowCount, 27).setCellValue(oOrgDio.Name(nLocale));
            g_objSheet.cell(g_rowCount, 28).setCellValue(oOrgDio.Attribute(Constants.AT_PERS_RESP, nLocale).GetValue(false));
            g_objSheet.cell(g_rowCount, 29).setCellValue(oOrgDio.Attribute(Constants.AT_NUM_OF_EMPL, nLocale).GetValue(false));
        }
        
       
    }
    
    //Function
    g_objSheet.cell(g_rowCount, 30).setCellValue(objFunction.Name(nLocale));
    g_objSheet.cell(g_rowCount, 31).setCellValue(objFunction.GUID().toString());
    g_objSheet.cell(g_rowCount, 32).setCellValue(objFunction.Attribute(989, nLocale).GetValue(false));
    
    //Radno mjesto
    if(oPosition != null)
        g_objSheet.cell(g_rowCount, 33).setCellValue(oPosition.Name(nLocale));
    
    //Aplikacija
    if(oApplication != null)
    {
        g_objSheet.cell(g_rowCount, 34).setCellValue(oApplication.Name(nLocale));
    }
    
    //InputOutput docs
    g_objSheet.cell(g_rowCount, 35).setCellValue(getInputDocString(objFunction));
    g_objSheet.cell(g_rowCount, 36).setCellValue(getOutputDocString(objFunction));
    
    writeAdditionalValues(objFunction,oPosition, oApplication, g_rowCount);
    
    g_rowCount += 1;
}
function writeFunctionForGroup(objFunction, objGroup, oPosition, oApplication)
{
    g_objSheet.cell(g_rowCount, 0).setCellValue(g_currentLevel1);
    g_objSheet.cell(g_rowCount, 1).setCellValue(g_currentLevel2);
    g_objSheet.cell(g_rowCount, 2).setCellValue(g_currentLevel3);
    g_objSheet.cell(g_rowCount, 3).setCellValue(g_currentLevel4);
    g_objSheet.cell(g_rowCount, 4).setCellValue(g_currentLevel5);
    g_objSheet.cell(g_rowCount, 5).setCellValue(g_currentLevel6);
    g_objSheet.cell(g_rowCount, 6).setCellValue(g_currentLevel7);
    g_objSheet.cell(g_rowCount, 7).setCellValue(g_currentLevel8);
    g_objSheet.cell(g_rowCount, 8).setCellValue(g_currentLevel9);
    g_objSheet.cell(g_rowCount, 9).setCellValue(g_currentLevel10);
    
    //VACD
    if(g_currentVACD != null)
    {
     g_objSheet.cell(g_rowCount, 10).setCellValue(g_currentVACD.Name(nLocale));
     g_objSheet.cell(g_rowCount, 11).setCellValue(g_currentVACD.GUID().toString());
     g_objSheet.cell(g_rowCount, 12).setCellValue(g_currentVACD.Attribute(133845, nLocale).GetValue(false));
     g_objSheet.cell(g_rowCount, 13).setCellValue(g_currentVACD.Attribute(264917, nLocale).GetValue(false));
     g_objSheet.cell(g_rowCount, 14).setCellValue(g_currentVACD.Attribute(68314, nLocale).GetValue(false));
     g_objSheet.cell(g_rowCount, 15).setCellValue(g_currentVACD.Attribute(527061, nLocale).GetValue(false));
     g_objSheet.cell(g_rowCount, 16).setCellValue(g_currentVACD.Attribute(264919, nLocale).GetValue(false));
     g_objSheet.cell(g_rowCount, 17).setCellValue(writeLinks(getVACDLinks(g_currentVACD)));
    }
    //VACD
    
    //EPC
    g_objSheet.cell(g_rowCount, 18).setCellValue(g_currentEPC.Name(nLocale));
    g_objSheet.cell(g_rowCount, 19).setCellValue(g_currentEPC.GUID().toString());
    g_objSheet.cell(g_rowCount, 20).setCellValue(getVACDFunction(g_currentEPC).Attribute(989, nLocale).GetValue(false));
    
    //ORG Dio
    if(oPosition != null)
    {
        var oOrgDio = getOrgDio(oPosition);
        getOrgLevels(oOrgDio);
        //write levels
        g_objSheet.cell(g_rowCount, 21).setCellValue(g_orgLevel1);
        g_objSheet.cell(g_rowCount, 22).setCellValue(g_orgLevel2);
        g_objSheet.cell(g_rowCount, 23).setCellValue(g_orgLevel3);
        g_objSheet.cell(g_rowCount, 24).setCellValue(g_orgLevel4);
        g_objSheet.cell(g_rowCount, 25).setCellValue(g_orgLevel5);
        g_objSheet.cell(g_rowCount, 26).setCellValue(g_orgLevel6);
        g_orgLevel1 = "";
        g_orgLevel2 = "";
        g_orgLevel3 = "";
        g_orgLevel4 = "";
        g_orgLevel5 = "";
        g_orgLevel6 = "";
        
        if(oOrgDio != null)
        {
            g_objSheet.cell(g_rowCount, 27).setCellValue(oOrgDio.Name(nLocale));
            g_objSheet.cell(g_rowCount, 28).setCellValue(oOrgDio.Attribute(Constants.AT_PERS_RESP, nLocale).GetValue(false));
            g_objSheet.cell(g_rowCount, 29).setCellValue(oOrgDio.Attribute(Constants.AT_NUM_OF_EMPL, nLocale).GetValue(false));
        }
        
        
    }
    
    //Function
    g_objSheet.cell(g_rowCount, 30).setCellValue(objFunction.Name(nLocale));
    g_objSheet.cell(g_rowCount, 31).setCellValue(objFunction.GUID().toString());
    g_objSheet.cell(g_rowCount, 32).setCellValue(objFunction.Attribute(989, nLocale).GetValue(false));
    
    //Radno mjesto
    if(objGroup != null)
        g_objSheet.cell(g_rowCount, 33).setCellValue(objGroup.Name(nLocale));
    
    //Aplikacija
    if(oApplication != null)
    {
        g_objSheet.cell(g_rowCount, 34).setCellValue(oApplication.Name(nLocale));
    }
    
    //InputOutput docs
    g_objSheet.cell(g_rowCount, 35).setCellValue(getInputDocString(objFunction));
    g_objSheet.cell(g_rowCount, 36).setCellValue(getOutputDocString(objFunction));
    
    writeAdditionalValues(objFunction,oPosition, oApplication, g_rowCount);
    
    g_rowCount += 1;
}
function writeFunctionForOrgUnit(objFunction, objGroup, oOrgUnit, oApplication)
{
    g_objSheet.cell(g_rowCount, 0).setCellValue(g_currentLevel1);
    g_objSheet.cell(g_rowCount, 1).setCellValue(g_currentLevel2);
    g_objSheet.cell(g_rowCount, 2).setCellValue(g_currentLevel3);
    g_objSheet.cell(g_rowCount, 3).setCellValue(g_currentLevel4);
    g_objSheet.cell(g_rowCount, 4).setCellValue(g_currentLevel5);
    g_objSheet.cell(g_rowCount, 5).setCellValue(g_currentLevel6);
    g_objSheet.cell(g_rowCount, 6).setCellValue(g_currentLevel7);
    g_objSheet.cell(g_rowCount, 7).setCellValue(g_currentLevel8);
    g_objSheet.cell(g_rowCount, 8).setCellValue(g_currentLevel9);
    g_objSheet.cell(g_rowCount, 9).setCellValue(g_currentLevel10);
    
    //VACD
    if(g_currentVACD != null)
    {
     g_objSheet.cell(g_rowCount, 10).setCellValue(g_currentVACD.Name(nLocale));
     g_objSheet.cell(g_rowCount, 11).setCellValue(g_currentVACD.GUID().toString());
     g_objSheet.cell(g_rowCount, 12).setCellValue(g_currentVACD.Attribute(133845, nLocale).GetValue(false));
     g_objSheet.cell(g_rowCount, 13).setCellValue(g_currentVACD.Attribute(264917, nLocale).GetValue(false));
     g_objSheet.cell(g_rowCount, 14).setCellValue(g_currentVACD.Attribute(68314, nLocale).GetValue(false));
     g_objSheet.cell(g_rowCount, 15).setCellValue(g_currentVACD.Attribute(527061, nLocale).GetValue(false));
     g_objSheet.cell(g_rowCount, 16).setCellValue(g_currentVACD.Attribute(264919, nLocale).GetValue(false));
     g_objSheet.cell(g_rowCount, 17).setCellValue(writeLinks(getVACDLinks(g_currentVACD)));
    }
    //VACD
    
    //EPC
    g_objSheet.cell(g_rowCount, 18).setCellValue(g_currentEPC.Name(nLocale));
    g_objSheet.cell(g_rowCount, 19).setCellValue(g_currentEPC.GUID().toString());
    g_objSheet.cell(g_rowCount, 20).setCellValue(getVACDFunction(g_currentEPC).Attribute(989, nLocale).GetValue(false));
    
    //ORG Dio
   
        var oOrgDio = oOrgUnit;
        getOrgLevels(oOrgDio);
        //write levels
        g_objSheet.cell(g_rowCount, 21).setCellValue(g_orgLevel1);
        g_objSheet.cell(g_rowCount, 22).setCellValue(g_orgLevel2);
        g_objSheet.cell(g_rowCount, 23).setCellValue(g_orgLevel3);
        g_objSheet.cell(g_rowCount, 24).setCellValue(g_orgLevel4);
        g_objSheet.cell(g_rowCount, 25).setCellValue(g_orgLevel5);
        g_objSheet.cell(g_rowCount, 26).setCellValue(g_orgLevel6);
        g_orgLevel1 = "";
        g_orgLevel2 = "";
        g_orgLevel3 = "";
        g_orgLevel4 = "";
        g_orgLevel5 = "";
        g_orgLevel6 = "";
        
        if(oOrgDio != null)
        {
            g_objSheet.cell(g_rowCount, 27).setCellValue(oOrgDio.Name(nLocale));
            g_objSheet.cell(g_rowCount, 28).setCellValue(oOrgDio.Attribute(Constants.AT_PERS_RESP, nLocale).GetValue(false));
            g_objSheet.cell(g_rowCount, 29).setCellValue(oOrgDio.Attribute(Constants.AT_NUM_OF_EMPL, nLocale).GetValue(false));
        }
        
        
    
    //Function
    g_objSheet.cell(g_rowCount, 30).setCellValue(objFunction.Name(nLocale));
    g_objSheet.cell(g_rowCount, 31).setCellValue(objFunction.GUID().toString());
    g_objSheet.cell(g_rowCount, 32).setCellValue(objFunction.Attribute(989, nLocale).GetValue(false));
    
    //Radno mjesto
    if(objGroup != null)
        g_objSheet.cell(g_rowCount, 33).setCellValue(objGroup.Name(nLocale));
    
    //Aplikacija
    if(oApplication != null)
    {
        g_objSheet.cell(g_rowCount, 34).setCellValue(oApplication.Name(nLocale));
    }
    
    //InputOutput docs
    g_objSheet.cell(g_rowCount, 35).setCellValue(getInputDocString(objFunction));
    g_objSheet.cell(g_rowCount, 36).setCellValue(getOutputDocString(objFunction));
    
    writeAdditionalValues(objFunction,null, oApplication, g_rowCount);
    
    g_rowCount += 1;
}


function writeAdditionalHeaderColumns()
{   
    var startingPoint = 37;
    
    for (var i = 0;i<g_atributetypenums.length;i++)
    {
         g_objSheet.cell(2, i + startingPoint).setCellValue("Atribut funkcije(" + ArisData.ActiveFilter().AttrTypeName(g_atributetypenums[i]) + ")");
         g_objSheet.cell(2, i+ startingPoint).setCellStyle(g_objSheet.cell(2,26).getCellStyle());
    }
    
    startingPoint = startingPoint + g_atributetypenums.length;
    for (var i = 0;i<g_modelatributetypenums.length;i++)
    {
         g_objSheet.cell(2, i + startingPoint).setCellValue("EPC atribut(" + ArisData.ActiveFilter().AttrTypeName(g_modelatributetypenums[i]) + ")");
         g_objSheet.cell(2, i + startingPoint).setCellStyle(g_objSheet.cell(2,26).getCellStyle());
    }
    
    startingPoint = startingPoint + g_modelatributetypenums.length;
    for (var i = 0;i<g_VACatributetypenums.length;i++)
    {
         g_objSheet.cell(2, i + startingPoint).setCellValue("VAC atribut(" + ArisData.ActiveFilter().AttrTypeName(g_VACatributetypenums[i]) + ")");
         g_objSheet.cell(2, i + startingPoint).setCellStyle(g_objSheet.cell(2,26).getCellStyle());
    }
    
    startingPoint = startingPoint + g_VACatributetypenums.length;
    for (var i = 0;i<g_POSatributetypenums.length;i++)
    {
         g_objSheet.cell(2, i + startingPoint).setCellValue("Atribut radnog mjesta(" + ArisData.ActiveFilter().AttrTypeName(g_POSatributetypenums[i]) + ")");
         g_objSheet.cell(2, i + startingPoint).setCellStyle(g_objSheet.cell(2,26).getCellStyle());
    }
    
    startingPoint = startingPoint + g_POSatributetypenums.length;
    for (var i = 0;i<g_APPatributetypenums.length;i++)
    {
         g_objSheet.cell(2, i + startingPoint).setCellValue("Atribut aplikacije(" + ArisData.ActiveFilter().AttrTypeName(g_APPatributetypenums[i]) + ")");
         g_objSheet.cell(2, i + startingPoint).setCellStyle(g_objSheet.cell(2,26).getCellStyle());
    }
}
function writeAdditionalValues(objFunction,objPosition, objApplication, currentRow)
{   
    var startingPoint = 37;
    for (var i = 0;i<g_atributetypenums.length;i++)
    {
         g_objSheet.cell(currentRow, i+startingPoint).setCellValue(objFunction.Attribute(g_atributetypenums[i], nLocale).GetValue(false));
         g_objSheet.cell(currentRow, i+startingPoint).setCellStyle(g_objSheet.cell(currentRow,26).getCellStyle());
    }
    
    startingPoint = startingPoint + g_atributetypenums.length;
    for (var i = 0;i<g_modelatributetypenums.length;i++)
    {
         g_objSheet.cell(currentRow, i + startingPoint).setCellValue(g_currentEPC.Attribute(g_modelatributetypenums[i], nLocale).GetValue(false));
         g_objSheet.cell(currentRow, i+startingPoint).setCellStyle(g_objSheet.cell(currentRow,26).getCellStyle());
    }
    
    if(g_currentVACD != null)
    {
    startingPoint = startingPoint + g_modelatributetypenums.length;
    for (var i = 0;i<g_VACatributetypenums.length;i++)
    {
         g_objSheet.cell(currentRow, i + startingPoint).setCellValue(g_currentVACD.Attribute(g_VACatributetypenums[i], nLocale).GetValue(false));
         g_objSheet.cell(currentRow, i+startingPoint).setCellStyle(g_objSheet.cell(currentRow,26).getCellStyle());
    }
    }
    
    if(objPosition != null)
    {
    startingPoint = startingPoint + g_VACatributetypenums.length;
    for (var i = 0;i<g_POSatributetypenums.length;i++)
    {
         g_objSheet.cell(currentRow, i + startingPoint).setCellValue(objPosition.Attribute(g_POSatributetypenums[i], nLocale).GetValue(false));
         g_objSheet.cell(currentRow, i+startingPoint).setCellStyle(g_objSheet.cell(currentRow,26).getCellStyle());
    }
    }
    
    if(objApplication != null)
    {
    startingPoint = startingPoint + g_POSatributetypenums.length;
    for (var i = 0;i<g_APPatributetypenums.length;i++)
    {
         g_objSheet.cell(currentRow, i + startingPoint).setCellValue(objApplication.Attribute(g_APPatributetypenums[i], nLocale).GetValue(false));
         g_objSheet.cell(currentRow, i+startingPoint).setCellStyle(g_objSheet.cell(currentRow,26).getCellStyle());
    }
    }
}

function getOrgDio(p_Object)
{
    var objTypes = new Array();
	objTypes[0] = Constants.OT_ORG_UNIT;
	var conTypes = new Array();
	conTypes[0] = Constants.CT_IS_CRT_BY;
	conTypes[1] = Constants.CT_IS_COMPOSED_OF;
	conTypes[2] = Constants.CT_IS_COMPONENT_OF;
	conTypes[3] = Constants.CT_SUPPORTS;
    
    var list = p_Object.getConnectedObjs(objTypes, Constants.EDGES_INOUTASSIGN,conTypes);
	
    if(list.length == 0)
        return null;
    else
        return list[0];
	
}

function getSuperiorOrg(p_Object)
{
    var objTypes = new Array();
	objTypes[0] = Constants.OT_ORG_UNIT;
	var conTypes = new Array();
	//conTypes[0] = Constants.CT_IS_SUPERIOR;
    conTypes[0] = Constants.CT_IS_COMPOSED_OF;
	conTypes[1] = Constants.CT_IS_COMPONENT_OF;
	//conTypes[3] = Constants.CT_IS_SUPERIOR_1;
    conTypes[2] = Constants.CT_IS_CRT_BY;
    
    var list = p_Object.getConnectedObjs(objTypes, Constants.EDGES_IN,conTypes);
	
    if(list.length == 0)
        return null;
    else
        return list[0];
	
}
function getOrgLevels(objOrg)
{
    var sOrgHL = "";
    var sOrgName = "";
    if(objOrg == null)
    {
        return;
    }
    else
        sOrgName = vbToLowerCase(objOrg.Name(nLocale));
        sOrgHL = objOrg.Attribute(1390,nLocale).GetValue(false);
    
    if( sOrgHL == 6 )
        g_orgLevel6 = objOrg.Name(nLocale);
    
    if( sOrgHL == 5)
        g_orgLevel5 = objOrg.Name(nLocale);
    
    if( sOrgHL == 4)
        g_orgLevel4 = objOrg.Name(nLocale);    
    
         if( sOrgHL == 3)
        g_orgLevel3 = objOrg.Name(nLocale);
    
    if( sOrgHL == 2)
        g_orgLevel2 = objOrg.Name(nLocale);
    
        if( sOrgHL == 1)
        g_orgLevel1 = objOrg.Name(nLocale);
         
    
    getOrgLevels(getSuperiorOrg(objOrg));
}

function getLink(oVACD)
{
    var listOLEObjects = oVACD.ComObjOccs();
    
    for (var j = 0;j<listOLEObjects.length;j++)
    {
        if(listOLEObjects[j].IsLink())
            return listOLEObjects[j].LinkedFile();
    }
    
    return "";
}
function getApplications(p_Object)
{
	var objTypes = new Array();
	objTypes[0] = Constants.OT_APPL_SYS_TYPE;
	var conTypes = new Array();
	conTypes[0] = Constants.CT_SUPP_1;
	conTypes[1] = Constants.CT_SUPP_2;
	conTypes[2] = Constants.CT_SUPP_3;
	conTypes[3] = Constants.CT_SUPPORTS;
	conTypes[4] = Constants.CT_CAN_SUPP_1;
	conTypes[5] = Constants.CT_CAN_SUPP_2;
	
	return p_Object.getConnectedObjs(objTypes, Constants.EDGES_INOUTASSIGN,conTypes);
}

function getPositions(p_Object) 
{
	var objTypes = new Array();
	objTypes[0] = Constants.OT_POS;
    objTypes[1] = Constants.OT_GRP;

	var conTypes = new Array();
	conTypes[0] = Constants.CT_CONTR_TO_1;
	conTypes[1] = Constants.CT_CONTR_TO_2;
	conTypes[2] = Constants.CT_EXEC_1;
	conTypes[3] = Constants.CT_EXEC_2;
	//conTypes[4] = Constants.CT_EXEC_3; //Custom 20220201
	conTypes[4] = Constants.CT_EXEC_4;
	conTypes[5] = Constants.CT_EXEC_5;
	
	return p_Object.getConnectedObjs(objTypes, Constants.EDGES_INOUTASSIGN, conTypes);
    
}
function getGroupPositions(p_Object) 
{
	var objTypes = new Array();
	objTypes[0] = Constants.OT_POS;
    
	var conTypes = new Array();
	conTypes[0] = Constants.CT_IS_CRT_BY;
	conTypes[1] = Constants.CT_CONTR_TO_2;
	conTypes[2] = Constants.CT_EXEC_1;
	conTypes[3] = Constants.CT_EXEC_2;
	//conTypes[4] = Constants.CT_EXEC_3; //Custom 20220201
	conTypes[4] = Constants.CT_EXEC_4;
	conTypes[5] = Constants.CT_EXEC_5;
	
	return p_Object.getConnectedObjs(objTypes, Constants.EDGES_INOUTASSIGN, conTypes);
    
}
function getGroupOrgUnits(p_Object) 
{
	var objTypes = new Array();
	objTypes[0] = Constants.OT_ORG_UNIT;
    
	var conTypes = new Array();
	conTypes[0] = Constants.CT_IS_CRT_BY;
	conTypes[1] = Constants.CT_CONTR_TO_2;
	conTypes[2] = Constants.CT_EXEC_1;
	conTypes[3] = Constants.CT_EXEC_2;
	//conTypes[4] = Constants.CT_EXEC_3; //Custom 20220201
	conTypes[4] = Constants.CT_EXEC_4;
	conTypes[5] = Constants.CT_EXEC_5;
    conTypes[6] = Constants.CT_IS_ASSIG_1;
    //conTypes[8] = Constants.CT_IS_ASSIG_2;//Custom 20220201
    conTypes[7] = Constants.CT_IS_ASSIG_3;
	
	return p_Object.getConnectedObjs(objTypes, Constants.EDGES_INOUTASSIGN, conTypes);
    
}


function getInputDocString(p_Object) 
{
	var outText = "";
	var objTypes = new Array();
	objTypes[0] = Constants.OT_TECH_TRM;
	objTypes[1] = Constants.OT_INFO_CARR;
	objTypes[2] = Constants.OT_LST;
	objTypes[3] = Constants.OT_LST_DSGN;
	var conTypes = new Array();
	conTypes[0] = Constants.CT_IS_INP_FOR;
	
	var objConnectedInputDocsList = p_Object.getConnectedObjs(objTypes, Constants.EDGES_IN, conTypes);
    
	if(objConnectedInputDocsList.length > 0)
    {
       
	    for(var index=0; index<objConnectedInputDocsList.length; index++) 
        {
            outText = outText + objConnectedInputDocsList[index].Attribute(Constants.AT_NAME, nLocale).GetValue(true) + "\r\n" + "\r\n";
        }
    }
    
    return outText;
}

function getOutputDocString(p_Object) 
{
	var outText = "";
	var objTypes = new Array();
	objTypes[0] = Constants.OT_TECH_TRM;
	objTypes[1] = Constants.OT_INFO_CARR;
	objTypes[2] = Constants.OT_LST;
	objTypes[3] = Constants.OT_LST_DSGN;
	var conTypes = new Array();
	conTypes[0] = Constants.CT_HAS_OUT;
	
	var objConnectedOutputDocsList = p_Object.getConnectedObjs(objTypes, Constants.EDGES_OUT, conTypes);
	
    if(objConnectedOutputDocsList.length > 0)
    {
        
	    for(var index=0; index<objConnectedOutputDocsList.length; index++) 
        {
            outText = outText + objConnectedOutputDocsList[index].Attribute(Constants.AT_NAME, nLocale).GetValue(true) + "\r\n" + "\r\n";           
        }
    }
    
    return outText;
}

function getAssignedEPClist(listGroupVACDs)
{
    var listLinkedEPCList = new Array();
    var listAssignedEPCs = new Array();
    
    for (var i = 0;i<listGroupVACDs.length;i++)
    {
        var listFuncOccs = listGroupVACDs[i].ObjOccListBySymbol(getTypeNumList2(Constants.ST_VAL_ADD_CHN_SML_1, Constants.ST_VAL_ADD_CHN_SML_2));
        
        for (var j = 0;j<listFuncOccs.length;j++)
        {
            listAssignedEPCs = listFuncOccs[j].ObjDef().AssignedModels(Constants.MT_EEPC);
            for (var k = 0;k<listAssignedEPCs.length;k++)
            {
                addNewObjectToList(listLinkedEPCList, listAssignedEPCs[k]);
            }
        }
    }
    
    return listLinkedEPCList;
}
function getAssignedEPCs(listGroupVACD)
{
    var listLinkedEPCList = new Array();
    var listAssignedEPCs = new Array();
    
    
    var listFuncOccs = listGroupVACD.ObjOccListBySymbol(getTypeNumList2(Constants.ST_VAL_ADD_CHN_SML_1, Constants.ST_VAL_ADD_CHN_SML_2));
    
    for (var j = 0;j<listFuncOccs.length;j++)
    {
        listAssignedEPCs = listFuncOccs[j].ObjDef().AssignedModels(Constants.MT_EEPC);
        for (var k = 0;k<listAssignedEPCs.length;k++)
        {
            addNewObjectToList(listLinkedEPCList, listAssignedEPCs[k]);
        }
    }
    
    
    return listLinkedEPCList;
}
function getVACDLinks(listGroupVACD)
{
    var listLinkList = new Array();
    
    var listFuncOccs = listGroupVACD.ObjOccListBySymbol(getTypeNumList2(Constants.ST_VAL_ADD_CHN_SML_1, Constants.ST_VAL_ADD_CHN_SML_2));
    
    for (var j = 0;j<listFuncOccs.length;j++)
    {
        
        if(listFuncOccs[j].ObjDef().Attribute(Constants.AT_EXT_1, nLocale).IsMaintained())
            addNewObjectToList(listLinkList, listFuncOccs[j].ObjDef().Attribute(Constants.AT_EXT_1, nLocale).GetValue(false));
        if(listFuncOccs[j].ObjDef().Attribute(Constants.AT_EXT_2, nLocale).IsMaintained())
            addNewObjectToList(listLinkList, listFuncOccs[j].ObjDef().Attribute(Constants.AT_EXT_2, nLocale).GetValue(false));
        if(listFuncOccs[j].ObjDef().Attribute(Constants.AT_EXT_3, nLocale).IsMaintained())
            addNewObjectToList(listLinkList, listFuncOccs[j].ObjDef().Attribute(Constants.AT_EXT_3, nLocale).GetValue(false));
        
    }
    
    
    return listLinkList;
}
function writeLinks(listLinks)
{
    var sLinkList ="";
    
    for (var j = 0;j<listLinks.length;j++)
    {
        sLinkList = sLinkList + listLinks[j] + "\r\n\r\n";
    }
    
    return sLinkList;
}

function getTypeNumList(iTypeNum)
{
    var listTypeNums = new Array();
    listTypeNums[0] = iTypeNum;
    
    return listTypeNums;
}
function getTypeNumList2(iTypeNum, iTypeNum2)
{
    var listTypeNums = new Array();
    listTypeNums[0] = iTypeNum;
    listTypeNums[1] = iTypeNum2;
    
    return listTypeNums;
}
function getTypeNumList3(iTypeNum, iTypeNum2, iTypeNum3)
{
    var listTypeNums = new Array();
    listTypeNums[0] = iTypeNum;
    listTypeNums[1] = iTypeNum2;
    listTypeNums[2] = iTypeNum3;
    
    return listTypeNums;
}

function objectExistsInList(list, objDef)
{
    for (var i = 0 ; i < list.length ; i++ )
    {
        if( vbStrComp(objDef.GUID(), list[i].GUID()) == 0)
        {
            return true;
        }
    }
    
    return false;
}
function addNewObjectToList(list, objDef)
{     
        if(!objectExistsInList(objDef) )
            list[list.length] = objDef;
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

function getCurrentLevelNames(objVACD)
{
    g_currentLevel1 = " ";
    g_currentLevel2 = " ";
    g_currentLevel3 = " ";
    g_currentLevel4 = " ";
    g_currentLevel5 = " ";
    g_currentLevel6 = " ";
    g_currentLevel7 = " ";
    g_currentLevel8 = " ";
    g_currentLevel9 = " ";
    g_currentLevel10 = " ";

    var listNames = objVACD.Group().Path(nLocale).replace("\\", "#").split("#");
    
    if(listNames.length == 1)
        g_currentLevel1 = listNames[0];
    
    if(listNames.length == 2)
    {
        g_currentLevel1 = listNames[0];
        g_currentLevel2 = listNames[1];
    }
    
    if(listNames.length == 3)
    {
        g_currentLevel1 = listNames[0];
        g_currentLevel2 = listNames[1];
        g_currentLevel3 = listNames[2];
    }
    
    if(listNames.length == 4)
    {
        g_currentLevel1 = listNames[0];
        g_currentLevel2 = listNames[1];
        g_currentLevel3 = listNames[2];
        g_currentLevel4 = listNames[3];
    }
    
    if(listNames.length == 5)
    {
        g_currentLevel1 = listNames[0];
        g_currentLevel2 = listNames[1];
        g_currentLevel3 = listNames[2];
        g_currentLevel4 = listNames[3];
        g_currentLevel5 = listNames[4];
    }
     if(listNames.length == 6)
    {
        g_currentLevel1 = listNames[0];
        g_currentLevel2 = listNames[1];
        g_currentLevel3 = listNames[2];
        g_currentLevel4 = listNames[3];
        g_currentLevel5 = listNames[4];
        g_currentLevel6 = listNames[5];
    }
     if(listNames.length == 7)
    {
        g_currentLevel1 = listNames[0];
        g_currentLevel2 = listNames[1];
        g_currentLevel3 = listNames[2];
        g_currentLevel4 = listNames[3];
        g_currentLevel5 = listNames[4];
        g_currentLevel6 = listNames[5];
        g_currentLevel7 = listNames[6];
    }
    if(listNames.length == 8)
    {
        g_currentLevel1 = listNames[0];
        g_currentLevel2 = listNames[1];
        g_currentLevel3 = listNames[2];
        g_currentLevel4 = listNames[3];
        g_currentLevel5 = listNames[4];
        g_currentLevel6 = listNames[5];
        g_currentLevel7 = listNames[6];
        g_currentLevel8 = listNames[7];
    }
    if(listNames.length == 9)
    {
        g_currentLevel1 = listNames[0];
        g_currentLevel2 = listNames[1];
        g_currentLevel3 = listNames[2];
        g_currentLevel4 = listNames[3];
        g_currentLevel5 = listNames[4];
        g_currentLevel6 = listNames[5];
        g_currentLevel7 = listNames[6];
        g_currentLevel8 = listNames[7];
        g_currentLevel9 = listNames[8];
    }
    if(listNames.length >= 10)
    {
        g_currentLevel1 = listNames[0];
        g_currentLevel2 = listNames[1];
        g_currentLevel3 = listNames[2];
        g_currentLevel4 = listNames[3];
        g_currentLevel5 = listNames[4];
        g_currentLevel6 = listNames[5];
        g_currentLevel7 = listNames[6];
        g_currentLevel8 = listNames[7];
        g_currentLevel9 = listNames[8];
        g_currentLevel10 = listNames[9];
    }
}

//Dialog functions

function displayui(settings)
{
  var __functionResult = false;
  //g_atributetypenumstoignore = getinitialatributetypenums();
  //g_modelatributetypenumstoignore = getinitialmodatributetypenums();
  
  var ndlgresult2 = 0; 
  ndlgresult2 = __toInteger(showVipDialog2(settings));
  
  if(ndlgresult2 == 0)
      return 0;
  else
  {
      if(settings.bConfigureObjAtrToIgnore)
      {
          var ndlgresult3 = 0; 
          var h_atrnums = new __holder(g_atributetypenumstoignore);
          ndlgresult3 = __toInteger(showVipDialog3(h_atrnums));
          if(ndlgresult3 == 0)
            return 0;
          
          //g_atributetypenumstoignore = h_atrnums.value;
      }
      
      if(settings.bConfigureModAtrToIgnore)
      {
          var ndlgresult4 = 0; 
          var h_atrnums = new __holder(g_modelatributetypenumstoignore);
          ndlgresult4 = __toInteger(showFilterModelAtributesDialog(h_atrnums));
          if(ndlgresult4 == 0)
            return 0;
          //g_modelatributetypenumstoignore = h_atrnums.value;
      }
      
      if(settings.bConfigureVACAtrToIgnore)
      {
          var ndlgresult5 = 0; 
          var h_atrnums = new __holder(g_VACatributetypenumstoignore);
          ndlgresult5 = __toInteger(showFilterVACModelAtributesDialog(h_atrnums));
          if(ndlgresult5 == 0)
            return 0;
          //g_modelatributetypenumstoignore = h_atrnums.value;
      }
      
      if(settings.bConfigurePOSAtrToIgnore)
      {
          var ndlgresult4 = 0; 
          var h_atrnums = new __holder(g_modelatributetypenumstoignore);
          ndlgresult4 = __toInteger(showPOSAttributesDialog(h_atrnums));
          //g_modelatributetypenumstoignore = h_atrnums.value;
          if(ndlgresult4 == 0)
            return 0;
      }
      
      if(settings.bConfigureAPPAtrToIgnore)
      {
          var ndlgresult4 = 0; 
          var h_atrnums = new __holder(g_modelatributetypenumstoignore);
          ndlgresult4 = __toInteger(showAPPAttributesDialog(h_atrnums));
          //g_modelatributetypenumstoignore = h_atrnums.value;
          if(ndlgresult4 == 0)
            return 0;
      }
      
      return true;
  }
  
  __functionResult = true;
  return __functionResult;
}

function showVipDialog2(settings)
{
  var __functionResult = 0;
  var chkModAtrIgnore = 0; 
  var chkObjAtrIgnore = 0; 
  var chkVACAtrIgnore = 0;
  var chkPOSAtrIgnore = 0;
  var chkAPPAtrIgnore = 0;
  
  chkModAtrIgnore = __toInteger(vbIIf(settings.bConfigureModAtrToIgnore, 1, 0));
  chkObjAtrIgnore = __toInteger(vbIIf(settings.bConfigureObjAtrToIgnore, 1, 0));
  chkVACAtrIgnore = __toInteger(vbIIf(settings.bConfigureVACAtrToIgnore, 1, 0));
  chkAPPAtrIgnore = __toInteger(vbIIf(settings.bConfigureAPPAtrToIgnore, 1, 0));
  chkPOSAtrIgnore = __toInteger(vbIIf(settings.bConfigurePOSAtrToIgnore, 1, 0));
 
  
  var userdialog = Dialogs.createNewDialogTemplate(300, 150, "Odaberite dodatne atribute:", "vipDialog2_Handler");
  
  userdialog.CheckBox(15, 20, 300, 15, "Želim dodatne atribute modela (EPC)", "chkModAtrIgnore");
  userdialog.CheckBox(15, 50, 300, 15, "Želim dodatne atribute funkcije/aktivnosti", "chkObjAtrIgnore");
  userdialog.CheckBox(15, 80, 300, 15, "Želim dodatne atribute modela (VAC)", "chkVACAtrIgnore");
  userdialog.CheckBox(15, 110, 300, 15, "Želim dodatne atribute za objekt Aplikacija", "chkAPPAtrIgnore");
  userdialog.CheckBox(15, 140, 300, 15, "Želim dodatne atribute za objekt Radno mjesto", "chkPOSAtrIgnore");
  
  userdialog.OKButton(15, 80, 50, 20);
  userdialog.CancelButton(80, 80, 50, 20);

  var dlg = Dialogs.createUserDialog(userdialog); 
  
  //set initial values
  dlg.setDlgValue("chkModAtrIgnore", chkModAtrIgnore);
  dlg.setDlgValue("chkObjAtrIgnore", chkObjAtrIgnore);
  dlg.setDlgValue("chkVACAtrIgnore", chkVACAtrIgnore);
  dlg.setDlgValue("chkAPPAtrIgnore", chkAPPAtrIgnore);
  dlg.setDlgValue("chkPOSAtrIgnore", chkPOSAtrIgnore);

  __functionResult = Dialogs.show( __currentDialog = dlg);

  if (__functionResult == - 1)
  {
    settings.bConfigureModAtrToIgnore = __toBool((dlg.getDlgValue("chkModAtrIgnore") == 1));
    settings.bConfigureObjAtrToIgnore = __toBool((dlg.getDlgValue("chkObjAtrIgnore") == 1));
    settings.bConfigureVACAtrToIgnore = __toBool((dlg.getDlgValue("chkVACAtrIgnore") == 1));
    settings.bConfigurePOSAtrToIgnore = __toBool((dlg.getDlgValue("chkPOSAtrIgnore") == 1));
    settings.bConfigureAPPAtrToIgnore = __toBool((dlg.getDlgValue("chkAPPAtrIgnore") == 1));
  }

  return __functionResult;
}
function vipDialog2_Handler(dlgitem, action, suppvalue)
{
  var __functionResult = false;

  switch(action) {
    case 2:
      if(dlgitem == "btnAtributesToIgnore")
      {
          var ndlgresult3 = 0; 
          var h_atrnums = new __holder(g_atributetypenumstoignore);
          ndlgresult3 = __toInteger(showVipDialog3(h_atrnums));
          g_atributetypenumstoignore = h_atrnums.value;
          __functionResult = false;
      }
    break;
  }

  return __functionResult;
}

var g_dlg3_atributetypenums;
var g_dlg3_atributetypenumstoignore;
function showVipDialog3(atributetypenumstoignore)
{
  var __functionResult = 0;

  var satributetypenames = __createArray(""); 
  var satributeignorenames = __createArray(""); 

  g_dlg3_atributetypenums = g_atributetypenums;
  g_dlg3_atributetypenumstoignore = getinitialatributetypenums();

  sobjecttypenames = getatributetypenames(g_dlg3_atributetypenums);
  signorenames = getatributetypenames(g_dlg3_atributetypenumstoignore);

  var userdialog = Dialogs.createNewDialogTemplate(0, 0, 470, 252,"Dodatni atributi funkcije/aktivnosti", "vipDialog3_Handler");
  // %GRID:10,7,1,1
  userdialog.ListBox(10, 77, 180, 140, sobjecttypenames, "lstObjectTypes");
  userdialog.ListBox(270, 77, 180, 140, signorenames, "lstIgnore");
  userdialog.Text(10, 7, 440, 49, "Odaberite dodatne atributi iz liste desno i dodajte ih u lijevu listu.", "Text1");
  userdialog.Text(10, 50, 190, 28, "Odabrani atributi", "Text2");
  userdialog.Text(270, 50, 150, 28, "Mogući dodatni atributi", "lblIgnore");
  userdialog.PushButton(200, 84, 60, 21, ">>", "butAddAll");
  userdialog.PushButton(200, 112, 60, 21, ">", "butAdd");
  userdialog.PushButton(200, 161, 60, 21, "<", "butRemove");
  userdialog.PushButton(200, 189, 60, 21, "<<", "butRemoveAll");
  userdialog.OKButton(100, 224, 90, 21);
  userdialog.CancelButton(270, 224, 90, 21);


  var dlg = Dialogs.createUserDialog(userdialog); 

  __functionResult = Dialogs.show( __currentDialog = dlg);

  if (__functionResult == - 1)
  {
    g_atributetypenumstoignore = g_dlg3_atributetypenums;
  }

  return __functionResult;
}
function showAPPAttributesDialog(atributetypenumstoignore)
{
  var __functionResult = 0;

  var satributetypenames = __createArray(""); 
  var satributeignorenames = __createArray(""); 

  g_dlg3_atributetypenums = g_APPatributetypenums;
  g_dlg3_atributetypenumstoignore = getinitialatributetypenums();

  sobjecttypenames = getatributetypenames(g_dlg3_atributetypenums);
  signorenames = getatributetypenames(g_dlg3_atributetypenumstoignore);

  var userdialog = Dialogs.createNewDialogTemplate(0, 0, 470, 252,"Dodatni atributi aplikacije", "vipDialog3_Handler");
  // %GRID:10,7,1,1
  userdialog.ListBox(10, 77, 180, 140, sobjecttypenames, "lstObjectTypes");
  userdialog.ListBox(270, 77, 180, 140, signorenames, "lstIgnore");
  userdialog.Text(10, 7, 440, 49, "Odaberite dodatne atributi iz liste desno i dodajte ih u lijevu listu.", "Text1");
  userdialog.Text(10, 50, 190, 28, "Odabrani atributi", "Text2");
  userdialog.Text(270, 50, 150, 28, "Mogući dodatni atributi", "lblIgnore");
  userdialog.PushButton(200, 84, 60, 21, ">>", "butAddAll");
  userdialog.PushButton(200, 112, 60, 21, ">", "butAdd");
  userdialog.PushButton(200, 161, 60, 21, "<", "butRemove");
  userdialog.PushButton(200, 189, 60, 21, "<<", "butRemoveAll");
  userdialog.OKButton(100, 224, 90, 21);
  userdialog.CancelButton(270, 224, 90, 21);


  var dlg = Dialogs.createUserDialog(userdialog); 

  __functionResult = Dialogs.show( __currentDialog = dlg);

  if (__functionResult == - 1)
  {
    g_APPatributetypenumstoignore = g_dlg3_atributetypenums;
  }

  return __functionResult;
}
function showPOSAttributesDialog(atributetypenumstoignore)
{
  var __functionResult = 0;

  var satributetypenames = __createArray(""); 
  var satributeignorenames = __createArray(""); 

  g_dlg3_atributetypenums = g_POSatributetypenums;
  g_dlg3_atributetypenumstoignore = getinitialatributetypenums();

  sobjecttypenames = getatributetypenames(g_dlg3_atributetypenums);
  signorenames = getatributetypenames(g_dlg3_atributetypenumstoignore);

  var userdialog = Dialogs.createNewDialogTemplate(0, 0, 470, 252,"Dodatni atributi radnog mjesta", "vipDialog3_Handler");
  // %GRID:10,7,1,1
  userdialog.ListBox(10, 77, 180, 140, sobjecttypenames, "lstObjectTypes");
  userdialog.ListBox(270, 77, 180, 140, signorenames, "lstIgnore");
  userdialog.Text(10, 7, 440, 49, "Odaberite dodatne atributi iz liste desno i dodajte ih u lijevu listu.", "Text1");
  userdialog.Text(10, 50, 190, 28, "Odabrani atributi", "Text2");
  userdialog.Text(270, 50, 150, 28, "Mogući dodatni atributi", "lblIgnore");
  userdialog.PushButton(200, 84, 60, 21, ">>", "butAddAll");
  userdialog.PushButton(200, 112, 60, 21, ">", "butAdd");
  userdialog.PushButton(200, 161, 60, 21, "<", "butRemove");
  userdialog.PushButton(200, 189, 60, 21, "<<", "butRemoveAll");
  userdialog.OKButton(100, 224, 90, 21);
  userdialog.CancelButton(270, 224, 90, 21);


  var dlg = Dialogs.createUserDialog(userdialog); 

  __functionResult = Dialogs.show( __currentDialog = dlg);

  if (__functionResult == - 1)
  {
    g_POSatributetypenumstoignore = g_dlg3_atributetypenums;
  }

  return __functionResult;
}

function vipDialog3_Handler(dlgitem, action, suppvalue)
{
  var __functionResult  = false;
  var sobjecttypenames  = __createArray(""); 
  var signorenames      = __createArray(""); 
  var idxobjecttype     = 0; 

  switch(action) {
    case 1:
      __currentDialog.setDlgEnable("butAddAll", true);
      __currentDialog.setDlgEnable("butAdd", true);
      __currentDialog.setDlgEnable("butRemove", true);
      __currentDialog.setDlgEnable("butRemoveAll", true);
    break;
    
    case 2:
      __functionResult = true;
      
      switch(dlgitem) {
        case "butAddAll":
          g_dlg3_atributetypenumstoignore = getallatributetypenums(Constants.CID_OBJDEF);
          g_dlg3_atributetypenums = new Array();
          signorenames = getatributetypenames(g_dlg3_atributetypenumstoignore);
          __currentDialog.setDlgListBoxArray("lstObjectTypes", sobjecttypenames);
          __currentDialog.setDlgListBoxArray("lstIgnore", signorenames);

          __currentDialog.setDlgEnable("butAddAll", false);
          __currentDialog.setDlgEnable("butAdd", false);
          __currentDialog.setDlgEnable("butRemoveAll", true);
        break;
        
        case "butAdd":
          idxobjecttype = __toLong( __currentDialog.getDlgValue("lstObjectTypes"));
          if (idxobjecttype != - 1) {
            addtolonglist(g_dlg3_atributetypenumstoignore, g_dlg3_atributetypenums[idxobjecttype]);
            g_dlg3_atributetypenums = removefromlonglist(g_dlg3_atributetypenums, idxobjecttype);
            sobjecttypenames = getatributetypenames(g_dlg3_atributetypenums);
            signorenames = getatributetypenames(g_dlg3_atributetypenumstoignore);

            __currentDialog.setDlgListBoxArray("lstObjectTypes", sobjecttypenames);
            __currentDialog.setDlgListBoxArray("lstIgnore", signorenames);
            __currentDialog.setDlgEnable("butAdd", false);
            __currentDialog.setDlgEnable("butRemoveAll", true);            
          }
        break;
        
        case "butRemove":
          idxobjecttype = __toLong( __currentDialog.getDlgValue("lstIgnore"));
          if (idxobjecttype != - 1) {
            addtolonglist(g_dlg3_atributetypenums, g_dlg3_atributetypenumstoignore[idxobjecttype]);
            g_dlg3_atributetypenumstoignore = removefromlonglist(g_dlg3_atributetypenumstoignore, idxobjecttype);
            sobjecttypenames = getatributetypenames(g_dlg3_atributetypenums);
            signorenames = getatributetypenames(g_dlg3_atributetypenumstoignore);

            __currentDialog.setDlgListBoxArray("lstObjectTypes", sobjecttypenames);
            __currentDialog.setDlgListBoxArray("lstIgnore", signorenames);
            __currentDialog.setDlgEnable("butRemove", false);
            __currentDialog.setDlgEnable("butAddAll", true);            
          }
        break;
        
        case "butRemoveAll":
          g_dlg3_atributetypenums = getallatributetypenums(Constants.CID_OBJDEF);
          g_dlg3_atributetypenumstoignore = new Array();
          sobjecttypenames = getatributetypenames(g_dlg3_atributetypenums);

          __currentDialog.setDlgListBoxArray("lstObjectTypes", sobjecttypenames);
          __currentDialog.setDlgListBoxArray("lstIgnore", signorenames);

          __currentDialog.setDlgEnable("butRemoveAll", false);
          __currentDialog.setDlgEnable("butRemove", false);
          __currentDialog.setDlgEnable("butAddAll", true);          
        break;
        
        case "lstObjectTypes":
          idxobjecttype = __toLong( __currentDialog.getDlgValue("lstObjectTypes"));
          if (idxobjecttype != - 1) {
            __currentDialog.setDlgEnable("butAdd", true);
          }
        break;
        
        case "lstIgnore":
          idxobjecttype = __toLong( __currentDialog.getDlgValue("lstIgnore"));
          if (idxobjecttype != - 1) {
            __currentDialog.setDlgEnable("butRemove", true);
          }
        break;
        
        case "OK":
        case "Cancel":
          __functionResult = false;
        break;
      }
    break;
  }

  return __functionResult;
}

var g_dlg_mod_atributetypenums;
var g_dlg_mod_atributetypenumstoignore;
function showFilterModelAtributesDialog(modatributetypenumstoignore)
{
  var __functionResult = 0;

  var satributetypenames = __createArray(""); 
  var satributeignorenames = __createArray(""); 

  g_dlg_mod_atributetypenums = g_modelatributetypenums;
  g_dlg_mod_atributetypenumstoignore = getinitialmodatributetypenums();

  sobjecttypenames = getatributetypenames(g_dlg_mod_atributetypenums);
  signorenames = getatributetypenames(g_dlg_mod_atributetypenumstoignore);

  var userdialog = Dialogs.createNewDialogTemplate(0, 0, 470, 252,"Odaberite dodatne atribute za EPC model", "filterModelAtributesDialog_Handler");
  // %GRID:10,7,1,1
  userdialog.ListBox(10, 77, 180, 140, sobjecttypenames, "lstObjectTypes");
  userdialog.ListBox(270, 77, 180, 140, signorenames, "lstIgnore");
  userdialog.Text(10, 7, 440, 49, "Odaberite dodatne atributi iz liste desno i dodajte ih u lijevu listu.", "Text1");
  userdialog.Text(10, 50, 190, 28, "Odabrani atributi", "Text2");
  userdialog.Text(270, 50, 150, 28, "Mogući dodatni atributi", "lblIgnore");
  userdialog.PushButton(200, 84, 60, 21, ">>", "butAddAll");
  userdialog.PushButton(200, 112, 60, 21, ">", "butAdd");
  userdialog.PushButton(200, 161, 60, 21, "<", "butRemove");
  userdialog.PushButton(200, 189, 60, 21, "<<", "butRemoveAll");
  userdialog.OKButton(100, 224, 90, 21);
  userdialog.CancelButton(270, 224, 90, 21);


  var dlg = Dialogs.createUserDialog(userdialog); 

  __functionResult = Dialogs.show( __currentDialog = dlg);

  if (__functionResult == - 1)
  {
    g_modelatributetypenumstoignore = g_dlg_mod_atributetypenums;
  }

  return __functionResult;
}
function showFilterVACModelAtributesDialog(modatributetypenumstoignore)
{
  var __functionResult = 0;

  var satributetypenames = __createArray(""); 
  var satributeignorenames = __createArray(""); 

  g_dlg_mod_atributetypenums = g_VACatributetypenums;
  g_dlg_mod_atributetypenumstoignore = getinitialmodatributetypenums();

  sobjecttypenames = getatributetypenames(g_dlg_mod_atributetypenums);
  signorenames = getatributetypenames(g_dlg_mod_atributetypenumstoignore);

  var userdialog = Dialogs.createNewDialogTemplate(0, 0, 470, 252,"Odaberite dodatne atribute za VAC model", "filterModelAtributesDialog_Handler");
  // %GRID:10,7,1,1
  userdialog.ListBox(10, 77, 180, 140, sobjecttypenames, "lstObjectTypes");
  userdialog.ListBox(270, 77, 180, 140, signorenames, "lstIgnore");
  userdialog.Text(10, 7, 440, 49, "Odaberite dodatne atributi iz liste desno i dodajte ih u lijevu listu.", "Text1");
  userdialog.Text(10, 50, 190, 28, "Odabrani atributi", "Text2");
  userdialog.Text(270, 50, 150, 28, "Mogući dodatni atributi", "lblIgnore");
  userdialog.PushButton(200, 84, 60, 21, ">>", "butAddAll");
  userdialog.PushButton(200, 112, 60, 21, ">", "butAdd");
  userdialog.PushButton(200, 161, 60, 21, "<", "butRemove");
  userdialog.PushButton(200, 189, 60, 21, "<<", "butRemoveAll");
  userdialog.OKButton(100, 224, 90, 21);
  userdialog.CancelButton(270, 224, 90, 21);


  var dlg = Dialogs.createUserDialog(userdialog); 

  __functionResult = Dialogs.show( __currentDialog = dlg);

  if (__functionResult == - 1)
  {
    g_VACatributetypenumstoignore = g_dlg_mod_atributetypenums;
  }

  return __functionResult;
}

function filterModelAtributesDialog_Handler(dlgitem, action, suppvalue)
{
  var __functionResult  = false;
  var sobjecttypenames  = __createArray(""); 
  var signorenames      = __createArray(""); 
  var idxobjecttype     = 0; 

  switch(action) {
    case 1:
      __currentDialog.setDlgEnable("butAddAll", true);
      __currentDialog.setDlgEnable("butAdd", true);
      __currentDialog.setDlgEnable("butRemove", true);
      __currentDialog.setDlgEnable("butRemoveAll", true);
    break;
    
    case 2:
      __functionResult = true;
      
      switch(dlgitem) {
        case "butAddAll":
          g_dlg_mod_atributetypenumstoignore = getallatributetypenums(Constants.CID_MODEL);
          g_dlg_mod_atributetypenums = new Array();
          signorenames = getatributetypenames(g_dlg_mod_atributetypenumstoignore);
          __currentDialog.setDlgListBoxArray("lstObjectTypes", sobjecttypenames);
          __currentDialog.setDlgListBoxArray("lstIgnore", signorenames);

          __currentDialog.setDlgEnable("butAddAll", false);
          __currentDialog.setDlgEnable("butAdd", false);
          __currentDialog.setDlgEnable("butRemoveAll", true);
        break;
        
        case "butAdd":
          idxobjecttype = __toLong( __currentDialog.getDlgValue("lstObjectTypes"));
          if (idxobjecttype != - 1) {
            addtolonglist(g_dlg_mod_atributetypenumstoignore, g_dlg_mod_atributetypenums[idxobjecttype]);
            g_dlg_mod_atributetypenums = removefromlonglist(g_dlg_mod_atributetypenums, idxobjecttype);
            sobjecttypenames = getatributetypenames(g_dlg_mod_atributetypenums);
            signorenames = getatributetypenames(g_dlg_mod_atributetypenumstoignore);

            __currentDialog.setDlgListBoxArray("lstObjectTypes", sobjecttypenames);
            __currentDialog.setDlgListBoxArray("lstIgnore", signorenames);
            __currentDialog.setDlgEnable("butAdd", false);
            __currentDialog.setDlgEnable("butRemoveAll", true);            
          }
        break;
        
        case "butRemove":
          idxobjecttype = __toLong( __currentDialog.getDlgValue("lstIgnore"));
          if (idxobjecttype != - 1) {
            addtolonglist(g_dlg_mod_atributetypenums, g_dlg_mod_atributetypenumstoignore[idxobjecttype]);
            g_dlg_mod_atributetypenumstoignore = removefromlonglist(g_dlg_mod_atributetypenumstoignore, idxobjecttype);
            sobjecttypenames = getatributetypenames(g_dlg_mod_atributetypenums);
            signorenames = getatributetypenames(g_dlg_mod_atributetypenumstoignore);

            __currentDialog.setDlgListBoxArray("lstObjectTypes", sobjecttypenames);
            __currentDialog.setDlgListBoxArray("lstIgnore", signorenames);
            __currentDialog.setDlgEnable("butRemove", false);
            __currentDialog.setDlgEnable("butAddAll", true);            
          }
        break;
        
        case "butRemoveAll":
          g_dlg_mod_atributetypenums = getallatributetypenums(Constants.CID_MODEL);
          g_dlg_mod_atributetypenumstoignore = new Array();
          sobjecttypenames = getatributetypenames(g_dlg_mod_atributetypenums);

          __currentDialog.setDlgListBoxArray("lstObjectTypes", sobjecttypenames);
          __currentDialog.setDlgListBoxArray("lstIgnore", signorenames);

          __currentDialog.setDlgEnable("butRemoveAll", false);
          __currentDialog.setDlgEnable("butRemove", false);
          __currentDialog.setDlgEnable("butAddAll", true);          
        break;
        
        case "lstObjectTypes":
          idxobjecttype = __toLong( __currentDialog.getDlgValue("lstObjectTypes"));
          if (idxobjecttype != - 1) {
            __currentDialog.setDlgEnable("butAdd", true);
          }
        break;
        
        case "lstIgnore":
          idxobjecttype = __toLong( __currentDialog.getDlgValue("lstIgnore"));
          if (idxobjecttype != - 1) {
            __currentDialog.setDlgEnable("butRemove", true);
          }
        break;
        
        case "OK":
        case "Cancel":
          __functionResult = false;
        break;
      }
    break;
  }

  return __functionResult;
}

function getatributetypenames(atributetypenums)
{
  var __functionResult = __createArray("");
  var result = __createArray(""); 
  var ofilter = ArisData.getActiveDatabase().ActiveFilter();

  if (atributetypenums.length > 0)
  {
    var i = 0; 
    result = new Array();

    for ( i = 0 ; i < (atributetypenums.length - 1)+1 ; i++ )
    {
      result[i] = ofilter.AttrTypeName(atributetypenums[i]);
    }
  }

  __functionResult = result;
  return __functionResult;
}
function getallatributetypenums(nEntityType)
{
  var result = new Array();
  var atributetypenums = __createArray(0); 
  var i = 0; 

  var ofilter = null; 
  ofilter = ArisData.getActiveDatabase().ActiveFilter();
  atributetypenums = ofilter.AttrTypes(nEntityType);

  var mNamesTypes = new java.util.TreeMap();
  for ( i = 0 ; i < atributetypenums.length ; i++ )
  {
    mNamesTypes.put("" + ofilter.AttrTypeName(atributetypenums[i]) + " " + atributetypenums[i], atributetypenums[i]);
  }
  
  result = new Array();

  var iter = mNamesTypes.keySet().iterator();
  while (iter.hasNext())
  {
    result.push(mNamesTypes.get(iter.next()));    
  }
/* 
  for ( i = 0 ; i < __getUBound(objecttypenums, 1)+1 ; i++ ){
    result[i] = objecttypenums[i];
  }
 */
  
  return result;
}
function getinitialatributetypenums()
{
  var result = new Array();
  var atributetypenums = __createArray(0); 
  var i = 0; 

  var ofilter = null; 
  ofilter = ArisData.getActiveDatabase().ActiveFilter();
  atributetypenums = ofilter.AttrTypes(Constants.CID_OBJDEF);

  var mNamesTypes = new java.util.TreeMap();
  for ( i = 0 ; i < atributetypenums.length ; i++ )
  {
      if(atributetypenums[i] != 1 &&
          atributetypenums[i] != 989)
    mNamesTypes.put("" + ofilter.AttrTypeName(atributetypenums[i]) + " " + atributetypenums[i], atributetypenums[i]);
  }
  
  result = new Array();

  var iter = mNamesTypes.keySet().iterator();
  while (iter.hasNext())
  {
    result.push(mNamesTypes.get(iter.next()));    
  }
/* 
  for ( i = 0 ; i < __getUBound(objecttypenums, 1)+1 ; i++ ){
    result[i] = objecttypenums[i];
  }
 */
  
  return result;
}

function getinitialmodatributetypenums()
{
  var result = new Array();
  var atributetypenums = __createArray(0); 
  var i = 0; 

  var ofilter = null; 
  ofilter = ArisData.getActiveDatabase().ActiveFilter();
  atributetypenums = ofilter.AttrTypes(Constants.CID_MODEL);

  var mNamesTypes = new java.util.TreeMap();
  for ( i = 0 ; i < atributetypenums.length ; i++ )
  {
      if(atributetypenums[i] != 1 &&
      atributetypenums[i] != 989)
    mNamesTypes.put("" + ofilter.AttrTypeName(atributetypenums[i]) + " " + atributetypenums[i], atributetypenums[i]);
  }
  
  result = new Array();

  var iter = mNamesTypes.keySet().iterator();
  while (iter.hasNext())
  {
    result.push(mNamesTypes.get(iter.next()));    
  }
/* 
  for ( i = 0 ; i < __getUBound(objecttypenums, 1)+1 ; i++ ){
    result[i] = objecttypenums[i];
  }
 */
  
  return result;
}

function addtolonglist(list, value)
{
  list[list.length] = value;
}

function addtostringlist(list, value)
{
  list[list.length] = value;
}

function removefromlonglist(list, index)
{
  var i = 0; 

  if (list.length <= 1) {
    list.length = 0;
  }
  else {
    for ( i = index ; i < list.length - 1 ; i++ ){
      list[i] = list[i + 1];
    }
    list.length = list.length-1;
  }
  __functionResult = list;
  return __functionResult;
}

//Dialog functions