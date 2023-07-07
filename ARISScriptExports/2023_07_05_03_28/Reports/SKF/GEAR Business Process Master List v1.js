var gsSettingsFile = "GEAR_Configuration.xml"


var msCaption = Context.getScriptInfo(Constants.SCRIPT_TITLE);
var msCaption2 = "GEAR Business Process Master List"
var mCurrentDatabase = ArisData.getActiveDatabase();
var mActiveDB = mCurrentDatabase
var mActiveMethod = mCurrentDatabase.ActiveFilter();

var mOutputFormat = Context.getSelectedFormat();
var gnLoc = Context.getSelectedLanguage();
var mOutput = null;

var mlBackCol_BusinessProcessArea = fcLng("&h" + GEAR_GetAttrValueInConfig("ReportRegistry/SymbolTypes/" + "BusinessProcessArea", "BckCol"))

var mlModelType_BusProcessMap_L3_L4 = GEAR_GetTypeInConfig("ModelTypes", "BusProcessMap_L3_L4")

var mlAttrType_Level                  = GEAR_GetTypeInConfig("AttributeTypes", "Level")
var mlAttrType_ReleaseSKF             = GEAR_GetTypeInConfig("AttributeTypes", "ReleaseSKF") 
var mlAttrType_OwnerProject           = GEAR_GetTypeInConfig("AttributeTypes", "OwnerProject")
var mlAttrType_OwnerBusiness          = GEAR_GetTypeInConfig("AttributeTypes", "OwnerBusiness")
var mlAttrType_GovernanceStatusObject = GEAR_GetTypeInConfig("AttributeTypes", "GovernanceStatusObject")  // APG stage
var mlAttrType_EAMClassification      = GEAR_GetTypeInConfig("AttributeTypes", "EAMClassification")



var msReportNameInConfig ="GEAR_BusinessProcessMasterList"
var msStartModelGUID = GEAR_GetValueInConfig("ScriptSettings/" + msReportNameInConfig  + "/StartModel")
var msGUIDFolderSemaphor = GEAR_GetValueInConfig("ScriptSettings/" + msReportNameInConfig  + "/FolderSemaphor")
var mFolderSemaphor = null
var msSemaphor = " -ACTIV- "
var msOutputFiles=[]
var msCurrentExportDate = fsJavaNow("MMM dd, yyyy hh:mm:ss a")   //fsJavaNow("yyyy-MM-dd-hh-mm-ss")
var gb_AutomaticScriptRun = false
var msFolderNameDebaredObjects= GEAR_GetAttrValueInConfig("ReportRegistry/DebaredFolder", "StartsWith")   //"Z "
var mFolderNameDebaredObjects = null


function main(){
    
    gs_Env = Context.getEnvironment();    
    Script_getAutomaticRun()
    if (gs_Env!="STD" || gb_AutomaticScriptRun){
        mbDialogAllowed=false
    }
    //gb_AutomaticScriptRun=false
    Script_GetSpecificFolders()
   
    if (gb_AutomaticScriptRun){
        Script_TestPreviousRun()
        Script_CreateSemaphore()
    }

    // rename output file
    var lsFileName = msCaption2 + "." + Files_ExtractExtention(Context.getSelectedFile())
    Context.setSelectedFile(lsFileName)
    msOutputFiles.push(lsFileName)
    // get outputfolder
    var lOutput= Context.createOutputObject()
    msTempFolder = lOutput.getCurrentPath()
    lOutput.WriteReport()  // close file

        
    Script_InitOutput( mlBackCol_BusinessProcessArea,  mlBackCol_BusinessProcessArea);    
    Script_Overview_Objecten();
    
    mOutput.WriteReport(); 
    
    if (gb_AutomaticScriptRun){
        Script_CopyFilesToOutboundFolder()
        Script_RemoveSemaphore()
    }
    
    
    return true
    
}


function Script_Overview_Objecten(){    
    
    var lsCaption = msCaption2
    mOutput.OutputLnF(lsCaption,"Header2"); 
    mOutput.BeginTable(100, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);              
    mOutput.TableRow();  
    mOutput.TableCellF("Process Area (Level 1)"                             ,30,"B1LEFT");            
    mOutput.TableCellF("Process Area (Level 2)"                             ,30,"B1LEFT");                
    mOutput.TableCellF("Process Area (Level 3)"                             ,30,"B1LEFT")
    mOutput.TableCellF("Release"                                            ,30,"B1LEFT")
    mOutput.TableCellF("Owner (Project)"                                    ,30,"B1LEFT")
    mOutput.TableCellF("Owner (Business)"                                   ,30,"B1LEFT")
    mOutput.TableCellF("APG Stage (Object)"                                 ,30,"B1LEFT")
    mOutput.TableCellF("EAM Classification"                                 ,30,"B1LEFT")
    
    
    var lSelModel = mCurrentDatabase.FindGUID(msStartModelGUID, Constants.CID_MODEL)
    if (!lSelModel.IsValid()) return false
    
    var lObjDefsLevel1 = ArisData.sort(lSelModel.ObjDefListFilter("L1", gnLoc, Constants.OT_FUNC, mlAttrType_Level), Constants.AT_NAME, gnLoc)
    for each (var lObjDefLevel1 in lObjDefsLevel1){
        if (Script_IsDaredObject(lObjDefLevel1)) continue  
            
        var lCxnDefs = lObjDefLevel1.CxnList(Constants.EDGES_OUT, [lSelModel])
        
        var lObjDefsLevel2 = []
        for each (var lCxnDef in lCxnDefs){
            if (lCxnDef.TypeNum() == Constants.CT_IS_PRCS_ORNT_SUPER){
                var lTargetObjDef = lCxnDef.TargetObjDef()
                
                if (Script_IsDaredObject(lTargetObjDef)) continue                      
                if (Aris_GetAttrValue(lTargetObjDef, mlAttrType_Level)!="L2") continue
                    
                lObjDefsLevel2.push(lTargetObjDef)                
            }
        }
        lObjDefsLevel2 = ArisData.sort(lObjDefsLevel2, Constants.AT_NAME, gnLoc)
        if (lObjDefsLevel2.length == 0){
                mOutput.TableRow();  
                mOutput.TableCellF(lObjDefLevel1.Name(-1)            ,30,"LEFT");            
                mOutput.TableCellF("n/a"                             ,30,"LEFT");                
                mOutput.TableCellF("n/a"                             ,30,"LEFT")
                mOutput.TableCellF("n/a"                             ,30,"LEFT")
                mOutput.TableCellF("n/a"                             ,30,"LEFT")
                mOutput.TableCellF("n/a"                             ,30,"LEFT")
                mOutput.TableCellF("n/a"                             ,30,"LEFT")
                mOutput.TableCellF("n/a"                             ,30,"LEFT")
        }else{
            for each (var lObjDefLevel2 in lObjDefsLevel2){
    
                var lAssignedModels = lObjDefLevel2.AssignedModels(mlModelType_BusProcessMap_L3_L4)
                var lObjDefsLevel3 = []
                for each (var lAssignedModel in lAssignedModels){
    
                    var lObjDefsAdHoc = ArisData.sort(lAssignedModel.ObjDefListFilter("L3", gnLoc, Constants.OT_FUNC, mlAttrType_Level), Constants.AT_NAME, gnLoc)
                    for each (var lObjDefAdHoc in lObjDefsAdHoc){
                        
                        if (Script_IsDaredObject(lObjDefAdHoc)) continue                         
                            
                        lObjDefsLevel3.push(lObjDefAdHoc)
                    }
                }   
                
                if (lObjDefsLevel3.length == 0){
                    mOutput.TableRow();  
                    mOutput.TableCellF(lObjDefLevel1.Name(-1)            ,30,"LEFT");            
                    mOutput.TableCellF(lObjDefLevel2.Name(-1)            ,30,"LEFT");                
                    mOutput.TableCellF("n/a"                             ,30,"LEFT")
                    mOutput.TableCellF("n/a"                             ,30,"LEFT")
                    mOutput.TableCellF("n/a"                             ,30,"LEFT")
                    mOutput.TableCellF("n/a"                             ,30,"LEFT")
                    mOutput.TableCellF("n/a"                             ,30,"LEFT")
                    mOutput.TableCellF("n/a"                             ,30,"LEFT")
                }else{
                    lObjDefsLevel3 = ArisData.Unique(lObjDefsLevel3)
                    lObjDefsLevel3 = ArisData.sort(lObjDefsLevel3, Constants.AT_NAME, gnLoc)
                    for each (var lObjDefLevel3 in lObjDefsLevel3){
                        mOutput.TableRow();  
                        mOutput.TableCellF(lObjDefLevel1.Name(-1)                                             ,30,"LEFT");            
                        mOutput.TableCellF(lObjDefLevel2.Name(-1)                                             ,30,"LEFT");                
                        mOutput.TableCellF(lObjDefLevel3.Name(-1)                                             ,30,"LEFT")
                        mOutput.TableCellF(Aris_GetAttrValue(lObjDefLevel3, mlAttrType_ReleaseSKF)            ,30,"LEFT")
                        mOutput.TableCellF(Aris_GetAttrValue(lObjDefLevel3, mlAttrType_OwnerProject)          ,30,"LEFT")
                        mOutput.TableCellF(Aris_GetAttrValue(lObjDefLevel3, mlAttrType_OwnerBusiness)         ,30,"LEFT")
                        mOutput.TableCellF(Aris_GetAttrValue(lObjDefLevel3, mlAttrType_GovernanceStatusObject),30,"LEFT")
                        mOutput.TableCellF(Aris_GetAttrValue(lObjDefLevel3, mlAttrType_EAMClassification)     ,30,"LEFT")
                    }
                }
            }
        }
    }
    
    if (mOutputFormat==Constants.OUTEXCEL){
        var lsEndTable = lsCaption;
    }else{
        var lsEndTable = "";
    }
    mOutput.EndTable(lsEndTable, 100, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);    
}
function Script_GetFolderSemaphor(){
    mFolderSemaphor = null
    var lbFolderFound = false
    if (fcIsGUID(msGUIDFolderSemaphor)){
        var lFolder = mActiveDB.FindGUID(msGUIDFolderSemaphor)
        if (lFolder){
            if (lFolder.IsValid()){
                mFolderSemaphor = lFolder                
            }
        }
    }
}
function Script_GetDebareFolder(){
    if (msFolderNameDebaredObjects.length>0){        
        var lRootFolders = mCurrentDatabase.RootGroup().Childs()
        for each ( var lFolder in lRootFolders) {
            if (lFolder.Name(-1).substr(0, msFolderNameDebaredObjects.length).toLowerCase() == msFolderNameDebaredObjects.toLowerCase()){
                mFolderNameDebaredObjects = lFolder
                break
            }
        }
    }
}
function Script_IsDaredObject(pObject){
    // objecten in a folder that is child of rootfolder that start with "Z "
    var lBool = false
    
    if (mFolderNameDebaredObjects==null) return false
    if (fIsMissing(pObject)) return false
        
    var lGroup = null

    switch(pObject.KindNum()){
    case Constants.CID_GROUP:
        lGroup = pObject
        break
    case Constants.CID_OBJDEF:        
    case Constants.CID_MODEL:            
        lGroup = pObject.Group()
        break
    }
    if (lGroup){
        var lsDaredPath = (mFolderNameDebaredObjects.Path(-1)) + ""
        if (lGroup.Path(-1).substr(0,lsDaredPath.length) == lsDaredPath){
            lBool=true
        }
    }
    return lBool
}

function Script_GetSpecificFolders(){
    Script_GetDebareFolder()    
    Script_GetFolderSemaphor()    
    
    if (mFolderSemaphor == null){    
        var lsEmailContent = "Warning: 'Semaphor folder' is missing."
                       + "\n\r" + "Nevertheless the script will continue."
                       + "\n\r" + "Check the GUIDs of the relevant folders and the settings in the configuration file."
                       + "\n\r" + "Export date: " + msCurrentExportDate
        Script_Email(lsEmailContent)            
    }
}

function Script_CopyFilesToOutboundFolder(){
    var lsPath2Element = "ScriptSettings/" + msReportNameInConfig  + "/Outbound"
    var lsOutboundFolder = GEAR_GetValueInConfig(lsPath2Element)
    var lbOK = Script_CheckFolder(lsOutboundFolder)

    if (lbOK){
        /*
        // remove old files in folder
        var lsFileNames = Files_GetFileNamesInDir(lsOutboundFolder)
        for each (var lsFileName in lsFileNames){
           Files_DeleteFile (lsOutboundFolder + "/" + lsFileName)
        }
        */

        
        // copy new created files to folder
        for each (lsOutputFile in msOutputFiles){
            if (!Files_CopyFile(msTempFolder + "/" + lsOutputFile, lsOutboundFolder + "/" + lsOutputFile)){
                var lsEmailContent = "Copy file '" + lsOutputFile + "' to outbound folder ('" + lsOutboundFolder + "') failed."
                                   + "\n\r" + "Check access rights of folder '" + lsOutboundFolder + "' or contact ARIS software manager."                                   
                                   + "\n\r" + "Export date: " + msCurrentExportDate
                                  
                Script_Email(lsEmailContent)   
                lbOK=false
                break
            }
        }
    }else{
        var lsEmailContent = "Outbound folder not specified in Configuration file"
                       + "\n\r" + "or doesn't exist on ARIS server"
                       + "\n\r" + "Check configuration file (see tag "
                       + "\n\r" + "'ScriptSettings/" + msReportNameInConfig  + "/Outbound' )."
                       + "\n\r" + "and check folder on ARIS server"
                       + "\n\r" + "Export date: " + msCurrentExportDate
        Script_Email(lsEmailContent)            
    }    
}
function Script_CheckFolder(psFolder){
    var lbOK= false
    var lsFolder = fValueIfMissing(fTrim(psFolder), "")
    if (lsFolder>""){
        lsFolder = Files_GetCorrectFileName(lsFolder)
        if (Files_Exists(lsFolder)){
            lbOK=true
        }else{
            lbOK = Files_CreatePath(lsFolder)
        }
    }
    return lbOK
}
function Script_getAutomaticRun(){  
    try{if (Dialogs){gb_AutomaticScriptRun=false}}catch(ex){gb_AutomaticScriptRun=true}    
    if (Context.getProperty("AutomaticRun")) gb_AutomaticScriptRun=true;        
//gb_AutomaticScriptRun=true    
}
function Script_Email(psContent){
    var lsPath2Element = "ScriptSettings/" + msReportNameInConfig  + "/Email"
    var lsFrom = GEAR_GetValueInConfig(lsPath2Element + "/SendFROM")
    var lsTOs  = GEAR_GetValueInConfig(lsPath2Element + "/SendTO").split(";")    
    var lsCCs  = GEAR_GetValueInConfig(lsPath2Element + "/SendCC").split(";")        
    var lsSubj = GEAR_GetValueInConfig(lsPath2Element + "/Subject")
    return EMAIL_SendMessage(lsFrom, lsTOs, lsCCs, lsSubj, psContent)
}
function Script_TestPreviousRun(){
    if (mFolderSemaphor){
        var lsFolderName = mFolderSemaphor.Name(-1)
        var lIdx = lsFolderName.lastIndexOf(msSemaphor)
        if (lIdx>-1){           
            var lsEmailContent = "Warning: Last run of the script is not finished properly"
                      + "\n\r" + "Nevertheless the script will continue."
                      + "\n\r" + "Export date: " + msCurrentExportDate 
            Script_Email(lsEmailContent)               
            var lsFolderName = lsFolderName.substr(0,lIdx)
            mFolderSemaphor.Attribute(Constants.AT_NAME,-1).setValue(lsFolderName)
        }
    }           
}
function Script_CreateSemaphore(){
    if (mFolderSemaphor){
        var lsFolderName = mFolderSemaphor.Name(-1)
        var lIdx = lsFolderName.lastIndexOf(msSemaphor)
        if (lIdx>-1){
            var lsFolderName = lsFolderName.substr(0,lIdx)
        }
        lsFolderName = lsFolderName + msSemaphor + msCurrentExportDate
        mFolderSemaphor.Attribute(Constants.AT_NAME,-1).setValue(lsFolderName)
    }
}
function Script_RemoveSemaphore(){
     if (mFolderSemaphor){
        var lsFolderName = mFolderSemaphor.Name(-1)
        var lIdx = lsFolderName.lastIndexOf(msSemaphor)
        if (lIdx>-1){           
            var lsFolderName = lsFolderName.substr(0,lIdx)
            mFolderSemaphor.Attribute(Constants.AT_NAME,-1).setValue(lsFolderName)
        }
    }
}

function Script_GetTextColorIndex(plBackCol){
    var lsHex = fcHex(plBackCol,6)
    var lTotal=0
    for (var i =0;i<lsHex.length;i+=2){
        if (parseInt(("0x" + lsHex.substr(i,2))*1)<0xa0)lTotal++
    }

    if (lTotal>=1){ 
        return new org.apache.poi.hssf.util.HSSFColor.WHITE().getIndex() //9 //Constants.C_WHITE  // index white
    }else{
        return new org.apache.poi.hssf.util.HSSFColor.BLACK().getIndex() //8 //Constants.C_BLACK
    }
}

function Script_InitOutput(plH1BackCol, plH2BackCol){
    var lLIndent=3;
    var lRIndent=3;

    
    mOutput = new PrivateOutputObject(null, plH1BackCol, plH2BackCol)

    mOutput.DefineF("DATE", "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT , lLIndent, lRIndent, 0, 0, 0, 1);
    mOutput.DefineF("LEFT", "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT , lLIndent, lRIndent, 0, 0, 0, 1);
    mOutput.DefineF("CENTER", "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_CENTER, lLIndent, lRIndent, 0, 0, 0, 1);
    mOutput.DefineF("RIGHT", "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_RIGHT, lLIndent, lRIndent, 0, 0, 0, 1);
    mOutput.DefineF("NUMERIC0", "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_RIGHT, lLIndent, lRIndent, 0, 0, 0, 1);    
    mOutput.DefineF("NUMERIC2", "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_RIGHT, lLIndent, lRIndent, 0, 0, 0, 1);        
  
    mOutput.DefineF("B1LEFT", "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_BOLD  , lLIndent, lRIndent, 0, 0, 0, 1);
    mOutput.DefineF("B1CENTER", "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_CENTER | Constants.FMT_BOLD , lLIndent, lRIndent, 0, 0, 0, 1);
    mOutput.DefineF("B1RIGHT", "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_RIGHT | Constants.FMT_BOLD , lLIndent, lRIndent, 0, 0, 0, 1);
    mOutput.DefineF("B1NUMERIC", "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_RIGHT | Constants.FMT_BOLD , lLIndent, lRIndent, 0, 0, 0, 1);
    mOutput.DefineF("B1DATE", "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_CENTER | Constants.FMT_BOLD , lLIndent, lRIndent, 0, 0, 0, 1);    
    
    mOutput.DefineF("B2LEFT", "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_BOLD  , lLIndent, lRIndent, 0, 0, 0, 1);
    mOutput.DefineF("B2CENTER", "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_CENTER | Constants.FMT_BOLD , lLIndent, lRIndent, 0, 0, 0, 1);
    mOutput.DefineF("B2RIGHT", "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_RIGHT | Constants.FMT_BOLD , lLIndent, lRIndent, 0, 0, 0, 1);
    mOutput.DefineF("B2NUMERIC", "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_RIGHT | Constants.FMT_BOLD , lLIndent, lRIndent, 0, 0, 0, 1);
    mOutput.DefineF("B2DATE", "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_CENTER | Constants.FMT_BOLD , lLIndent, lRIndent, 0, 0, 0, 1);    

    
    mOutput.DefineF("LEFTWRAP", "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_BOLD  , lLIndent, lRIndent, 0, 0, 0, 1);
    mOutput.DefineF("CENTERWRAP", "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_CENTER | Constants.FMT_BOLD , lLIndent, lRIndent, 0, 0, 0, 1);
    mOutput.DefineF("RIGHTWRAP", "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_RIGHT | Constants.FMT_BOLD , lLIndent, lRIndent, 0, 0, 0, 1);


}

///////////////////////////////////////
// output object


function PrivateOutputObject(psOptionalTemplate, plH1BackCol, plH2BackCol){
    var lWorkBook=null;
    var lSheet = null;
    var lSheetIndex = -1;
    var lRow = null;
    var lRowTeller = 0;
    var lColumnTeller = 0;
    var lKSStyles = new KeySet();
    var lStyles = new Array();
    
        
    this.OutputLnF = function (psText1,pstext2){
        return
    }    
    this.BeginTable = function(p_iWidthPercent, p_BorderColor, p_BkColor, p_Format, p_iIndent ){
        lSheet = lWorkBook.createSheet("xxxxyyyyzzzz");
        lSheetIndex = lWorkBook.getSheetCount()-1;
        lRowTeller = 0;

    }
    this.EndTable = function (psSheetName, plWidth, psFontName, plFontsize, plFontColor, plBackColor, plShading, plFormat, plIndent){
        if ((psSheetName+"")!=""){
            lWorkBook.setSheetName(lSheetIndex, psSheetName);
        }
    }
    
    this.TableRow =function (){        
        lRow = lSheet.createRow(lRowTeller);
        lRowTeller++;
        lColumnTeller =0;
    }  
    this.TableCellF = function(psText, plWidth, psStyleName){
        var lCell = lRow.createCell(lColumnTeller);
        
        var lIndex = lKSStyles.GetIndex(psStyleName);
        if (lIndex>-1){
            lCell.setCellStyle(lStyles[lIndex]);
        }else{
            var a=1/0;
        }
        lSheet.setColumnWidth(lColumnTeller, plWidth*250);
        
  
        if (psText=="" || fIsMissing(psText)){
            lCell.setCellType(Constants.XL_CELL_TYPE_BLANK); 
            lCell.setCellValue("");
        }else{
             switch(psStyleName){
                case "BRIGHT"  : lCell.setCellType(Constants.XL_CELL_TYPE_STRING);break;
                case "BLEFT"   : lCell.setCellType(Constants.XL_CELL_TYPE_STRING);break;
                case "BDATE"   : lCell.setCellType(Constants.XL_CELL_TYPE_STRING);break;
                case "BCENTER" : lCell.setCellType(Constants.XL_CELL_TYPE_STRING);break;
                case "BNUMERIC": lCell.setCellType(Constants.XL_CELL_TYPE_STRING);break;
                case "RIGHT"   : lCell.setCellType(Constants.XL_CELL_TYPE_STRING);break;
                case "LEFT"    : lCell.setCellType(Constants.XL_CELL_TYPE_STRING);break;
                case "DATE"    : lCell.setCellType(Constants.XL_CELL_TYPE_STRING);break;
                case "CENTER"  : lCell.setCellType(Constants.XL_CELL_TYPE_STRING);break;
                case "RIGHTWRAP":lCell.setCellType(Constants.XL_CELL_TYPE_STRING);break;
                case "LEFTWRAP": lCell.setCellType(Constants.XL_CELL_TYPE_STRING);break;
                case "CENTERWRAP":
                lCell.setCellType(Constants.XL_CELL_TYPE_STRING);break;

                case "NUMERIC0" : lCell.setCellType(Constants.XL_CELL_TYPE_NUMERIC);break;
                case "NUMERIC2" : lCell.setCellType(Constants.XL_CELL_TYPE_NUMERIC);break;
            }
            switch(psStyleName){
                case "DATE":                    
                    var lsDate = psText.substr(6,4) + "/" + psText.substr(3,2) +"/" + psText.substr(0,2) + " 0:0:0" 
                    lCell.setCellValue(fcJavaDate(lsDate));
                    break;
                default:
                    var lsText = psText; 
                    if (isNaN(psText)){
                        lCell.setCellValue(psText)
                    }else{
                        lCell.setDoubleCellValue(psText * 1)
                    }        
            }
        }        
        
        lColumnTeller++;
    }
    
    this.setWrapText= function(psStyleName,pBool){
       var lIndex = lKSStyles.GetIndex(psStyleName);
        if (lIndex>-1){
            lStyles[lIndex].setWrapText(pBool)
        }else{
            var a=1/0;
        }
    }
    
    this.DefineF = function (psStyleName
                           , psFontName, plFontSize
                           , plFontColor, plBGColor
                           , plAlignment 
                           , plLIndent, pRIndent, plDistTop, plDistBottom
                           , p_iIndFirstLine, p_iLineSpacing ) {
                               
        var lIndex = lKSStyles.AddItem(psStyleName);      

        switch(psStyleName){
            case "B1RIGHT"  : lStyles[lIndex] = lCellStyleHR1; break;
            case "B1LEFT"   : lStyles[lIndex] = lCellStyleHL1; break;
            case "B1DATE"   : lStyles[lIndex] = lCellStyleHD1; break;
            case "B1CENTER" : lStyles[lIndex] = lCellStyleHC1; break;
            case "B1NUMERIC": lStyles[lIndex] = lCellStyleHN1; break;            
            
            case "B2RIGHT"  : lStyles[lIndex] = lCellStyleHR2; break;
            case "B2LEFT"   : lStyles[lIndex] = lCellStyleHL2; break;
            case "B2DATE"   : lStyles[lIndex] = lCellStyleHD2; break;
            case "B2CENTER" : lStyles[lIndex] = lCellStyleHC2; break;
            case "B2NUMERIC": lStyles[lIndex] = lCellStyleHN2; break;            
            
            case "RIGHT"   : lStyles[lIndex] = lCellStyleR ; break;
            case "LEFT"    : lStyles[lIndex] = lCellStyleL ; break;
            case "DATE"    : lStyles[lIndex] = lCellStyleD ; break;
            case "CENTER"  : lStyles[lIndex] = lCellStyleC ; break;
            case "RIGHTWRAP":lStyles[lIndex] = lCellStyleRW ; break;
            case "LEFTWRAP": lStyles[lIndex] = lCellStyleLW ; break;
            case "CENTERWRAP":lStyles[lIndex]= lCellStyleCW ; break;

            case "NUMERIC" : lStyles[lIndex] = lCellStyleN0; break; 
            case "NUMERIC0" : 
               lStyles[lIndex] = lCellStyleN0 ; break;             
            case "NUMERIC2" : lStyles[lIndex] = lCellStyleN2 ; break;                         
            default:
            var a=1/0;
        }
    }
    
    this.WriteReport = function(){
        lWorkBook.write();
    }

    
    if (!fIsMissing(psOptionalTemplate)){
        lWorkBook = Context.createExcelWorkbook(Context.getSelectedFile(),psOptionalTemplate);
    }else{
        lWorkBook = Context.createExcelWorkbook(Context.getSelectedFile());        
    }

    var lFontH1 = lWorkBook.createFont();
    var lFontH2 = lWorkBook.createFont();    
    var lFontN =lWorkBook.createFont();
    var lFontN2 =lWorkBook.createFont();    

    lFontN.setFontName("ARIAL");
    lFontN.setFontHeightInPoints(10);
    lFontN2.setFontName("ARIAL");
    lFontN2.setFontHeightInPoints(10);
    //lFontNormal.setBoldweight (800);
    lFontH1.setFontName("ARIAL");
    lFontH1.setBoldweight (800);
    lFontH1.setFontHeightInPoints(12);
    //lFontH1.setColor(Script_GetTextColorIndex(plH1BackCol))
    lFontH2.setFontName("ARIAL");
    lFontH2.setBoldweight (800);
    lFontH2.setFontHeightInPoints(12);
    //lFontH2.setColor(Script_GetTextColorIndex(plH2BackCol))

    
    //Output variables
    var lBackColorHeader      = Constants.C_GREY_80_PERCENT;
    var lBackColorObject      = Constants.C_WHITE;
    var lBackColorComplex     = Constants.C_LIGHT_YELLOW;
    var lBackColorAttribute   = Constants.C_WHITE;

    
    var lCellStyleHL2= lWorkBook.createCellStyle (lFontH2, 1, 1, 1, 1, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_LEFT ,Constants.VERTICAL_CENTER , plH2BackCol , plH2BackCol , 1 );
    var lCellStyleHR2= lWorkBook.createCellStyle (lFontH2, 1, 1, 1, 1, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_RIGHT ,Constants.VERTICAL_CENTER , plH2BackCol , plH2BackCol , 1 );
    var lCellStyleHC2= lWorkBook.createCellStyle (lFontH2, 1, 1, 1, 1, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_CENTER ,Constants.VERTICAL_CENTER , plH2BackCol , plH2BackCol , 1 );    
    var lCellStyleHD2= lWorkBook.createCellStyle (lFontH2, 1, 1, 1, 1, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_CENTER ,Constants.VERTICAL_CENTER , plH2BackCol , plH2BackCol , 1 );        
    var lCellStyleHN2= lWorkBook.createCellStyle (lFontH2, 1, 1, 1, 1, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_RIGHT ,Constants.VERTICAL_CENTER , plH2BackCol , plH2BackCol , 1 );

    var lCellStyleHL1= lWorkBook.createCellStyle (lFontH1, 1, 1, 1, 1, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_LEFT ,Constants.VERTICAL_CENTER , plH1BackCol , plH1BackCol , 1 );
    var lCellStyleHR1= lWorkBook.createCellStyle (lFontH1, 1, 1, 1, 1, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_RIGHT ,Constants.VERTICAL_CENTER , plH1BackCol , plH1BackCol , 1 );
    var lCellStyleHC1= lWorkBook.createCellStyle (lFontH1, 1, 1, 1, 1, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_CENTER ,Constants.VERTICAL_CENTER , plH1BackCol , plH1BackCol , 1 );    
    var lCellStyleHD1= lWorkBook.createCellStyle (lFontH1, 1, 1, 1, 1, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_CENTER ,Constants.VERTICAL_CENTER , plH1BackCol , plH1BackCol , 1 );        
    var lCellStyleHN1= lWorkBook.createCellStyle (lFontH1, 1, 1, 1, 1, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_RIGHT ,Constants.VERTICAL_CENTER , plH1BackCol , plH1BackCol , 1 );
    
    
    var lCellStyleL = lWorkBook.createCellStyle (lFontN, 1, 1, 1, 1, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_LEFT ,Constants.VERTICAL_CENTER ,lBackColorAttribute , lBackColorAttribute ,1 );
    var lCellStyleR = lWorkBook.createCellStyle (lFontN, 1, 1, 1, 1, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_RIGHT ,Constants.VERTICAL_CENTER ,lBackColorAttribute , lBackColorAttribute ,1 );
    var lCellStyleC = lWorkBook.createCellStyle (lFontN, 1, 1, 1, 1, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_CENTER ,Constants.VERTICAL_CENTER ,lBackColorAttribute , lBackColorAttribute ,1 );
    
    var lCellStyleLW = lWorkBook.createCellStyle (lFontN2, 1, 1, 1, 1, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_LEFT ,Constants.VERTICAL_CENTER ,lBackColorAttribute , lBackColorAttribute ,1 );
    var lCellStyleRW = lWorkBook.createCellStyle (lFontN2, 1, 1, 1, 1, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_RIGHT ,Constants.VERTICAL_CENTER ,lBackColorAttribute , lBackColorAttribute ,1 );
    var lCellStyleCW = lWorkBook.createCellStyle (lFontN2, 1, 1, 1, 1, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_CENTER ,Constants.VERTICAL_CENTER ,lBackColorAttribute , lBackColorAttribute ,1 );
    
    
    var lCellStyleD = lWorkBook.createCellStyle (lFontN, 1, 1, 1, 1, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_CENTER ,Constants.VERTICAL_CENTER ,lBackColorAttribute , lBackColorAttribute ,1 );
    var lCellStyleN0 = lWorkBook.createCellStyle (lFontN, 1, 1, 1, 1, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_RIGHT ,Constants.VERTICAL_CENTER ,lBackColorAttribute , lBackColorAttribute ,1 );    
    var lCellStyleN2 = lWorkBook.createCellStyle (lFontN2, 1, 1, 1, 1, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.C_BLACK, Constants.ALIGN_RIGHT ,Constants.VERTICAL_CENTER ,lBackColorAttribute , lBackColorAttribute ,1 );        
    
    lCellStyleHL1.setWrapText(true);
    lCellStyleHR1.setWrapText(true);
    lCellStyleHC1.setWrapText(true);
    lCellStyleHD1.setWrapText(true);
    lCellStyleHN1.setWrapText(true);
    
    lCellStyleHL2.setWrapText(true);
    lCellStyleHR2.setWrapText(true);
    lCellStyleHC2.setWrapText(true);
    lCellStyleHD2.setWrapText(true);
    lCellStyleHN2.setWrapText(true);
    
    lCellStyleLW.setWrapText(true);
    lCellStyleRW.setWrapText(true);
    lCellStyleCW.setWrapText(true);    
    
    lCellStyleL.setWrapText(false);
    lCellStyleR.setWrapText(false);
    lCellStyleC.setWrapText(false);
    lCellStyleD.setWrapText(false);
    lCellStyleN0.setWrapText(false);
    lCellStyleN2.setWrapText(false);    
    
    lCellStyleN0.setDataFormat(Constants.XL_CELL_DATAFORMAT_INT);
    lCellStyleN2.setDataFormat(Constants.XL_CELL_DATAFORMAT_2DIG);    
    lCellStyleD.setDataFormat(Constants.XL_CELL_DATAFORMAT_M_D_YY);
    
}



var lbOK = main()
if (!lbOK)Context.setScriptError(Constants.ERR_CANCEL)
if (gb_AutomaticScriptRun)Context.setScriptError(Constants.ERR_NOFILECREATED)    
