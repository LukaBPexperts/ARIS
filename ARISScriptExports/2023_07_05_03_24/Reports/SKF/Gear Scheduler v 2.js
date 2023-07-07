var gsSettingsFile = "GEAR_Configuration.xml"
var msCaption = "Scheduler"
var mcProfileFile = "GEAR Scheduler Settings.xml"
var mcReportComponent = 1
var mScriptAdminComponent = Context.getComponent("ScriptAdmin") 
var gnLoc = Context.getSelectedLanguage()
var mSelDBs = ArisData.getSelectedDatabases()
var mActiveDB = ArisData.getActiveDatabase()

function main(){
    glTimerIntervalInLog = 60000     // see routine Log_Text for more info
    gbBreakScriptRun = false         // see routine Log_Text for more info 

    // is user allowed to see scripts?
    if (fIsMissing(mScriptAdminComponent)) return false
    
    var lScriptCategory = Script_GetScriptCategory()
        
    Log_Text("Start GEAR Scheduler")    
    var lScriptInfos = mScriptAdminComponent.getScriptInfos (mcReportComponent, lScriptCategory.getCategoryID(), gnLoc)
    for each (var lScriptInfo in lScriptInfos){        
        if (!lScriptInfo.isSimpleFile() 
        && lScriptInfo.getName()!=Context.getScriptInfo(Constants.SCRIPT_NAME)){
            
            // real script and not the current script
            Log_Text("Start Script " + lScriptInfo.getName())
            if (gbBreakScriptRun) break 
                
            var lReportComponent =  Context.getComponent("Report")  
            var lReportExecData = lReportComponent.createExecInfo(lScriptInfo.getID()
                                                                , mSelDBs
                                                                , Context.getSelectedLanguage()
                                                                , Context.getSelectedFormat()
                                                                , Context.getSelectedFile());
            lReportExecData.setProperty("AutomaticRun", "true")                                                            
            var lResult = lReportComponent.execute(lReportExecData)

        }
    }
    Log_Text("End GEAR Scheduler")    
}

function Script_GetScriptCategory(){ //psProfileName,psKeyPrefix){
     
    var lCategory = null
    
    var lsTmp = "" + fsNow("yyyyMMdd hhmmss") + " " + parseInt(Math.random() * 1000); 

    var lsKey="___";
    Profile_WriteKeyValue(lsKey,lsTmp);
    
    var lsKeyInclPrefix = lsKey;
    //if (psKeyPrefix!=null && psKeyPrefix!=undefined){
    //    lsKeyInclPrefix = psKeyPrefix + "_" + lsKeyInclPrefix;
    //}
    var lsCategories = new Array()    
    
    var lScriptAdminComponent = Context.getComponent("ScriptAdmin")
    lsCategories = lScriptAdminComponent.getCategories(mcReportComponent, Context.getSelectedLanguage())
   
    var lbFound = false;
    for (var i in lsCategories){
        var lCatID = lsCategories[i].getCategoryID();
        var lScriptInfos= lScriptAdminComponent.getScriptInfos (mcReportComponent, lCatID, Context.getSelectedLanguage())             
        for (var j in lScriptInfos){
            if (new java.lang.String(lScriptInfos[j].getName()).equalsIgnoreCase(mcProfileFile)){
                // bingo!, nu nog controleren of lsTmpFile klopt
                var lScriptID = lScriptInfos[j].getID();
                var lZipEntry = lScriptAdminComponent.exportFile(1, lCatID, lScriptID) 
                var lByteArray= lZipEntry.getData(); 
                var lsValue=Profile_GetKeyValueByByteArray(lByteArray, lsKey, "");                    
                if (lsValue==lsTmp){
                    // dubbel bingo, we hebben de juiste folder te pakken
                    lCategory = lsCategories[i]
                    lbFound=true;
                    break
                }
            }
        }
        if (lbFound)break;
        
    } 
    return lCategory      
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
function Script_CreateSemaphore(psReportName){
    if (mFolderSemaphor){
        var lsFolderName = mFolderSemaphor.Name(-1)
        var lIdx = lsFolderName.lastIndexOf(msSemaphor)
        if (lIdx>-1){
            var lsFolderName = lsFolderName.substr(0,lIdx)
        }
        lsFolderName = lsFolderName + msSemaphor + msCurrentExportDate + " " + psReportName
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


///////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////+ Profile
////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//dinow-frbar -> allow users like 000_System 
function Profile_GetXmlCompatibleName(psValue){
    return "a"+psValue;
}

function Profile_WriteKeyValue (psKey, psValue){    
    var lsKey = Profile_GetCorrectedKey(msCaption + "_" + psKey);
    var lsUserName = "" + Profile_GetXmlCompatibleName(ArisData.ActiveUser().Name(-1));
    lsUserName = lsUserName.length==0 ? "-" : lsUserName;
    lsUserName=lsUserName.replace(/[-]/g,"_");
    Context.writePrivateProfileString(lsUserName,lsKey, psValue, mcProfileFile);
}
function Profile_GetKeyValue (psKey, psDefault){    
    var lsValue2Return = "";
    
    var lsKey = Profile_GetCorrectedKey(msCaption + "_" + psKey);
    var lsUserName = "" +Profile_GetXmlCompatibleName(Aris_GetName(ArisData.ActiveUser()));
    lsUserName = lsUserName.length==0 ? "-" : lsUserName;
    lsUserName=lsUserName.replace(/[-]/g,"_");
    lsValue2Return = Context.getPrivateProfileString(lsUserName,lsKey, psDefault, mcProfileFile);
    
    return lsValue2Return;
}
function Profile_GetKeyValueByXML_file (psFileName, psKey, psDefault){ 
    var lsValue2Return = "";
    var lFileContent = Files_ReadByteArrayByFile(psFileName, true);
    var lsKey = Profile_GetCorrectedKey(msCaption + "_" + psKey);

    
    if (lFileContent.length>0){    
        try{            
            var lXMLParser = Context.getXMLParser(lFileContent);        
        }catch(ex){var lex=ex;
            var aa=1;
            if (lex.name=="InternalError"){
               try{      
                    var lFileContent = Files_ReadByteArrayByFile(psFileName, false);
                    var lXMLParser = Context.getXMLParser(lFileContent);                    
                }catch(ex){var lex=ex;
                    var aa=1;
                }
            }       
        }    
        try{            
            
            if (lXMLParser!=null){
                var lsUserName = "a"+Aris_GetName(ArisData.ActiveUser());
                lsUserName=lsUserName.replace(/[-]/g,"_");
                var lsUserNameAsElementName = Profile_GetCorrectedKey(lsUserName);
                var lSection = lXMLParser.getRootElement().getChild(lsUserNameAsElementName);
                if (lSection!=null){
                   lsValue2Return = XML_Element_GetText(lSection, lsKey)
                }
            }
        }catch(ex){var lex=ex;
            var aa=1;
        }    
              
    }
    if (lsValue2Return == "" ){
        if (psDefault!=null && psDefault!=undefined){
            lsValue2Return = psDefault;
        }
    }
    return lsValue2Return;
}
function Profile_GetKeyValueByByteArray(pByteArray, psKey, psDefault){ 
    var lsValue2Return = "";
    
    var lsKey = Profile_GetCorrectedKey(msCaption + "_" + psKey);

    
    if (pByteArray.length>0){    
        try{            
            var lXMLParser = Context.getXMLParser(pByteArray);        
        }catch(ex){var lex=ex;
        }    
        try{                        
            if (lXMLParser!=null){
                var lsUserName = "a"+ArisData.ActiveUser().Name(-1);
                lsUserName=lsUserName.replace(/[-]/g,"_");
                var lsUserNameAsElementName = Profile_GetCorrectedKey(lsUserName);
                var lSection = lXMLParser.getRootElement().getChild(lsUserNameAsElementName);
                if (lSection!=null){
                   lsValue2Return = XML_Element_GetText(lSection, lsKey)
                }
            }
        }catch(ex){var lex=ex;
            var aa=1;
        }    
              
    }
    if (lsValue2Return == "" ){
        if (psDefault!=null && psDefault!=undefined){
            lsValue2Return = psDefault;
        }
    }
    return lsValue2Return;
}


function Profile_GetCorrectedKey(psKey){
    var lsKey2Return = fValueIfMissing(psKey + "","");
    lsKey2Return = lsKey2Return.replace(/[ .()]/g,"_");

    return lsKey2Return;
}


main()