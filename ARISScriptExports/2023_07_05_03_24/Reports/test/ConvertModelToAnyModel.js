var gaSelModels=new Array();
var giModelTypeNum=-1;
var gSupObjDefGuidMap=new java.util.HashMap();
var gaiLocs=new Array();
var gsModelTypeNumDerivedGuid="";
var gsModelTypeAPIName="";
var gOutputLog=null;
var gsCurrentModelName="";
var pLocaleID = Context.getSelectedLanguage();
var lGroup="";

function XMLConversionResult(){
    this.bSuccessFul=false;
    this.aFileBytes=new Array();
}
function main(){    
    gaSelModels=ArisData.getSelectedModels();
    gaiLocs.push(Context.getSelectedLanguage());
    copyDTDFile("ARIS-Export.dtd");
    if(ArisData.getActiveDatabase().ActiveFilter().IsFullMethod()){
        if(checkSelModels()){
            if(selectModelType()){
                gOutputLog=Context.createOutputObject (Context.getSelectedFormat(), Context.getSelectedFile());
                writeToLog("masse test" + gsModelTypeNumDerivedGuid,"I");
                storeSupObjDefRelations();
                var lIter=gSupObjDefGuidMap.keySet().iterator();
                while(lIter.hasNext()){
                    var lsModelGUID=lIter.next();
                    var lModel=ArisData.getActiveDatabase().FindGUID(lsModelGUID); 
                    gsCurrentModelName=lModel.Name(gaiLocs[0]);
                    var lXMLExportResultWrapper=exportModel(lModel);
                    lGroup=lModel.Group();
                    if(lXMLExportResultWrapper.isSuccessful()){                    
                        Context.setFile("Export2.xml",Constants.LOCATION_SCRIPT ,lXMLExportResultWrapper.getFileResult());
                        var lXMLConversionResult=getChangedXML(lXMLExportResultWrapper.getFileResult());
                        if(lXMLConversionResult.bSuccessFul){
                            Context.setFile("Import2.xml",Constants.LOCATION_SCRIPT,lXMLConversionResult.aFileBytes);                        
                            if(deleteModel(lModel)){
                                var lXMLImportResultWrapper=importModel(lXMLConversionResult.aFileBytes,gaiLocs[0]);
                                if(lXMLImportResultWrapper.isSuccessful()){
                                    if(lXMLImportResultWrapper.getMessages().length>0){
                                         writeToLog("Import messages","W",lXMLImportResultWrapper.getMessages());
                                    }
                                    if(restoreSupObjDefRelations(lsModelGUID)){
                                        writeToLog("Succesful processed","I");  
                                    }
                                }else{                               
                                    writeToLog("Import failed","E",lXMLImportResultWrapper.getMessages());
                                }
                            }else{
                                writeToLog("Removal of model failed","E");
                            }
                        }
                    }else{
                        writeToLog("Export failed","E",lXMLExportResultWrapper.getMessages());
                    }
                }
                gOutputLog.WriteReport();
            }
        }else{
            Dialogs.MsgBox("All selected models should be of the same type!");
        }
    }else{
        Dialogs.MsgBox("Script should be evaluated with entire method!");
    }
}

function copyDTDFile(sDTDName) {
    var inputStream = null;
    var outputStream = null;
    try {
        
        var targetDTD = new java.io.File(sDTDName);

        var sSourcePath = new java.io.File(Context.getTempPath()).getParent();
        sSourcePath = new java.io.File(sSourcePath, "templates");
        sSourcePath = new java.io.File(sSourcePath, "scriptservice");
        sSourcePath = new java.io.File(sSourcePath, "processgen");
       
        var sourceDTD = new java.io.File(sSourcePath, sDTDName);

        if (StrComp(targetDTD.getAbsolutePath(), sourceDTD.getAbsolutePath()) != 0) {
            inputStream = new java.io.FileInputStream(sourceDTD);
            outputStream = new java.io.FileOutputStream(targetDTD);
            
            var buffer = java.lang.reflect.Array.newInstance(java.lang.Byte.TYPE, 1024);
            var i = inputStream.read(buffer);
            while (i != -1) {
                outputStream.write(buffer, 0, i);
                i = inputStream.read(buffer);
            }
            inputStream.close();
            outputStream.close();
            return targetDTD;        

        } else {
            return null;
        }
    } catch(e) {
        if(inputStream!=null) inputStream.close(); 
        if(outputStream!=null) outputStream.close();         
        return null;
    }
}
function selectModelType(){
    var laiModelTypes=ArisData.getActiveDatabase().ActiveFilter().ModelTypes(Constants.ARISVIEW_ALL);
    var liBaseModelTypeNum=giModelTypeNum & 0xffff;
    var laiModelTypesSel=new Array();
    var laiModelTypesSelNames=new Array();
    for(var i=0;i<laiModelTypes.length;i++){
// All model types? Yes:
//       if(laiModelTypes[i]!=giModelTypeNum && (laiModelTypes[i] & 0xffff)==liBaseModelTypeNum){
            laiModelTypesSel.push(laiModelTypes[i]);
            laiModelTypesSelNames.push(ArisData.getActiveDatabase().ActiveFilter().ModelTypeName(laiModelTypes[i]));
//        }
    }
    var lDlgTemplate=Dialogs.createNewDialogTemplate(100,100,400,250,"Select model type");
    lDlgTemplate.Text(10,10,150,20,"Convert to model type")
    lDlgTemplate.DropListBox(180,10,200,20,laiModelTypesSelNames,"dlModelSel");
    lDlgTemplate.OKButton();
    lDlgTemplate.CancelButton();
    var lDlg=Dialogs.createUserDialog(lDlgTemplate);
    var liResult=Dialogs.show(lDlg);
    if(liResult==-1){
        var liSelIdx=lDlg.getDlgValue("dlModelSel");
        gsModelTypeAPIName=""+ArisData.getActiveDatabase().ActiveFilter().getAPIName(Constants.CID_MODEL, laiModelTypesSel[liSelIdx]); //Added by Matias 14.1.2011: Conversion to other standard model type, APIname needed
        if(laiModelTypesSel[liSelIdx]==liBaseModelTypeNum){
            gsModelTypeNumDerivedGuid="";
        }else{
            gsModelTypeNumDerivedGuid=""+ArisData.getActiveDatabase().ActiveFilter().UserDefinedModelTypeGUID(laiModelTypesSel[liSelIdx]);
        }
        return true;
    }else{
        return false;
    }
    
}
function writeToLog(psMsg,psType,pasMsgs){
    var lsMsg=psType + " - Model: "+ gsCurrentModelName + ", Msg: " + psMsg + "\r\n";
    if(pasMsgs!=undefined){
        if(pasMsgs.length>0){
            lsMsg+="\r\n";
            for(var i=0;i<pasMsgs.length;i++){
                lsMsg=lsMsg+ pasMsgs[i] +"\r\n";
            }
            lsMsg+="\r\n";
        }
    }
    gOutputLog.OutputTxt(lsMsg); 
}
function getChangedXML(paFileBytes){
    var lXMLConversionResult=new XMLConversionResult();
    var lXMLStream = new java.io.ByteArrayInputStream(paFileBytes);
    var lBuilder = new Packages.org.jdom.input.SAXBuilder();
    var lDocument = lBuilder.build(lXMLStream);
    var lElementRoot = lDocument.getRootElement();
    
    changeModelInGroup(lElementRoot);
    
    var lOutputStream= new java.io.ByteArrayOutputStream();    
    var lXMLOutput = new Packages.org.jdom.output.XMLOutputter();
           
    lXMLOutput.output(lDocument,lOutputStream);
    lXMLConversionResult.bSuccessFul=true;
    lXMLConversionResult.aFileBytes=lOutputStream.toByteArray();
    return lXMLConversionResult;
}

function changeModelInGroup(pGroupElem){
    try{
         var lModelElems=pGroupElem.getChildren("Model");
         for(var i=0;i<lModelElems.size();i++){
            var lModelElem=lModelElems.get(i);
            var lTypeGuidElem=lModelElem.getChild("TypeGUID");
            lModelElem.setAttribute("Model.Type", gsModelTypeAPIName); //Added by Matias 14.1.2011
            if(lTypeGuidElem!=null){              //if source type is derived
                if(gsModelTypeNumDerivedGuid.length==0){ //if target type is standard
                    lModelElem.removeChild("TypeGUID")
                }else{
                    lTypeGuidElem.setText(gsModelTypeNumDerivedGuid); //if target type is derived
                }
            }else{                                 //if source type is standard
                if(gsModelTypeNumDerivedGuid.length>0){  //if target type is derived
                    lTypeGuidElem=new Packages.org.jdom.Element("TypeGUID");
                    lTypeGuidElem.setText(gsModelTypeNumDerivedGuid);
                    var liIndex=lModelElem.indexOf(lModelElem.getChild("GUID"))+1;                     
                    lModelElem.addContent(liIndex,lTypeGuidElem);                    
                }
            }        
         }
         var lGroupElems=pGroupElem.getChildren("Group");
         for(var i=0;i<lGroupElems.size();i++){
            var lGroupElem=lGroupElems.get(i);
            changeModelInGroup(lGroupElem);
         }
    }catch(ex){
        writeToLog("Xml conversion failed. error: " + ex,"E") 
    }
}
function checkSelModels(){    
    var lbResult=true;
    for(var i=0;i<gaSelModels.length;i++){
        if(giModelTypeNum==-1 || giModelTypeNum==gaSelModels[i].TypeNum()){
            giModelTypeNum=gaSelModels[i].TypeNum();
        }else{
            lbResult=false;
            break
        }
    }
    return lbResult;
}
function storeSupObjDefRelations(){
    for(var i=0;i<gaSelModels.length;i++){
        var lModel=gaSelModels[i];
        var laSupObjDefs=lModel.getSuperiorObjDefs();
        var laSupObjDefsGuids=new Array();
        for(var j=0;j<laSupObjDefs.length;j++){
            laSupObjDefsGuids.push(laSupObjDefs[j].GUID());
        }
        gSupObjDefGuidMap.put(lModel.GUID(),laSupObjDefsGuids);
    }
}
function exportModel(pModel){
    var lXMLExport=Context.getComponent("XMLExport");
    var lXMLExportOptionsWrapper=lXMLExport.createDefaultXMLExportOptions(pModel,gaiLocs);
    lXMLExportOptionsWrapper.setZipExport(false);
    lXMLExportOptionsWrapper.setIsSimpleExport(false);
    lXMLExportOptionsWrapper.setExportCompleteGroupPaths(false);
    lXMLExportOptionsWrapper.setHierarchyDepth(0);
    lXMLExportOptionsWrapper.setRelationDepth(0);
    var lXMLExportResultWrapper=lXMLExport.startExport(lXMLExportOptionsWrapper);
    return lXMLExportResultWrapper;
}
function importModel(pFileBytes,pLocaleID){
    var lXMLImport=Context.getComponent("XMLImport");
    var lDefaultModelOptions=lXMLImport.createDefaultModelOptions();
    var lAMLOptions=lXMLImport.createDefaultAMLOptions(pFileBytes, lGroup, pLocaleID, lDefaultModelOptions);
//    var lAMLOptions=lXMLImport.createDefaultAMLOptions(pFileBytes, ArisData.getActiveDatabase().RootGroup(), pLocaleID, lDefaultModelOptions);
    var lMergeOptions=lXMLImport.createDefaultMergeOptions();
    lMergeOptions.setConsiderGroupPath(false);
    return lXMLImport.startImport(lAMLOptions, lMergeOptions);   
}
function deleteModel(pModel){
    return pModel.Group().Delete(pModel);    
}
function restoreSupObjDefRelations(psModelGUID){
    try{
        var lModel=ArisData.getActiveDatabase().FindGUID(psModelGUID);
        var laSupObjDefGuids=gSupObjDefGuidMap.get(psModelGUID);
        for(var i=0;i<laSupObjDefGuids.length;i++){
            var lSupObjDef=ArisData.getActiveDatabase().FindGUID(laSupObjDefGuids[i]);
            lSupObjDef.CreateAssignment(lModel);
        }
        return true
    }catch(ex){
        writeToLog("Restore of assignments failed. Error: " + ex,"E");
        return false
    }        
}

main();