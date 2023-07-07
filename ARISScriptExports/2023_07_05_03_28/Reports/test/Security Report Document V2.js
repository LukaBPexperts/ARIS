//Context.setProperty("model-as-emf", true)
var oOutput = Context.createOutputObject()
var xlTableActive = false
oOutput.DefineF(getString("ID_STYLE_RD_DEFAULT"), getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VTOP, 0, 0, 0, 0, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_HEADING_1"), getString("ID_DEFAULT_FONT"), 18, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_TOCENTRY0, 0, 0, 4, 4, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_HEADING_2"), getString("ID_DEFAULT_FONT"), 14, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_ITALIC | Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_TOCENTRY1, 0, 0, 2, 2, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_TABLE_CONTENT"), getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VTOP, 0, 0, 0, 0, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_HEADING_3"), getString("ID_DEFAULT_FONT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_ITALIC | Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_TOCENTRY2, 0, 0, 1, 1, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_HEADING_4"), getString("ID_DEFAULT_FONT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_ITALIC | Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_TOCENTRY3, 0, 0, 0, 0, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_INFO"), getString("ID_DEFAULT_FONT"), 14, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VTOP, 0, 0, 1.76, 8.82, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_TABLE_HEAD"), getString("ID_DEFAULT_FONT"), 8, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VTOP, 0, 0, 0, 0, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_TITLE"), getString("ID_DEFAULT_FONT"), 21, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER | Constants.FMT_VTOP, 0, 0, 1.76, 8.82, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_HEADER_FOOTER"), getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_VTOP, 0, 0, 0, 0, 0, 1)
setupOutputObject(oOutput)
oOutput.SetTitle(Context.getScriptInfo(Constants.SCRIPT_NAME))

var startObject = ArisData.getSelectedObjDefs() //dohvati objekt na kojem se pokrece report

var selectedModel = ArisData.getSelectedModels() //dohvati model na kojem se pokrece report

var nLocale = Context.getSelectedLanguage() //dohvati jezik



main(selectedModel);

oOutput.WriteReport() //ispisi report

/** Apply default page format settings to output object
 * @param {Output} outputObj The output object
 */

function setupOutputObject(outputObj) {
    outputObj.SetAutoTOCNumbering(true)
    globalHeader(outputObj)

    globalFooter(outputObj)
}

function globalHeader(outputObj) {
    outputObj.BeginHeader()
    if (Context.getSelectedFormat() != Constants.OUTEXCEL)
        outputObj.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
    outputObj.TableRow()
    outputObj.ResetFrameStyle()
    outputObj.SetFrameStyle(Constants.FRAME_TOP, 0, 0)
    outputObj.SetFrameStyle(Constants.FRAME_LEFT, 0, 0)
    outputObj.SetFrameStyle(Constants.FRAME_BOTTOM, 51, Constants.BRDR_NORMAL)
    outputObj.SetFrameStyle(Constants.FRAME_RIGHT, 0, 0)
    outputObj.TableCell("", 36.72, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
    outputObj.BeginParagraph(Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
    var image = Context.createPicture("logo_report.png")
    outputObj.OutGraphic(image, -1, 55, 20)

    outputObj.EndParagraph()
    outputObj.TableCell("", 63.28, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
    outputObj.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    outputObj.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
    outputObj.EndParagraph()
    outputObj.BeginParagraph(Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
    outputObj.OutputLn("Security Report", getString("ID_REPORTDEF_11"), 24, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0.71)
    outputObj.EndParagraph()
    if (Context.getSelectedFormat() != Constants.OUTEXCEL)
        outputObj.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
    else
        outputObj.TableRow()
    outputObj.EndHeader()
}

function globalFooter(outputObj) {
    outputObj.BeginFooter()
    if (Context.getSelectedFormat() != Constants.OUTEXCEL)
        outputObj.BeginTable(100, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_NOBORDER, 0)
    outputObj.TableRow()
    outputObj.ResetFrameStyle()
    outputObj.SetFrameStyle(Constants.FRAME_TOP, 0, 0)
    outputObj.SetFrameStyle(Constants.FRAME_LEFT, 0, 0)
    outputObj.SetFrameStyle(Constants.FRAME_BOTTOM, 0, 0)
    outputObj.SetFrameStyle(Constants.FRAME_RIGHT, 0, 0)
    outputObj.TableCell("", 1.66, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
    outputObj.TableCell("", 98.34, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_CENTER, 0)
    outputObj.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    outputObj.Output(getString("ID_REPORTDEF_2"), getString("ID_REPORTDEF_11"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
    outputObj.OutputField(Constants.FIELD_PAGE, getString("ID_REPORTDEF_11"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER)
    outputObj.Output(getString("ID_REPORTDEF_3"), getString("ID_REPORTDEF_11"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
    outputObj.OutputField(Constants.FIELD_NUMPAGES, getString("ID_REPORTDEF_11"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER)
    outputObj.OutputLn("", getString("ID_REPORTDEF_11"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
    outputObj.EndParagraph()
    if (Context.getSelectedFormat() != Constants.OUTEXCEL)
        outputObj.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
    else
        outputObj.TableRow()
    outputObj.EndFooter()
}

/**
 * @param {Output} p_output The output object
 * @param {ObjDef[]} p_aObjDef 
 */

//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
//                                                                              FUNCTIONS
//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------


function main(selectedModel) {
    for (var i = 0; i < selectedModel.length; i++) {
        
      var allFunctionsFromEpc = selectedModel[i].ObjOccListFilter(Constants.OT_FUNC)   //sve funkcije sa modela
        
      var decisionFunctions = getFunctionsConnectedToXorRule(allFunctionsFromEpc)  //vraca decision funkcije iz niza svih funkcija (one koje iza sebe imaju XOR rule)
      
      sortByPosition(decisionFunctions)
      
      for(var j=0;j<decisionFunctions.length;j++){
           
           oOutput.BeginSection(false, Constants.SECTION_DEFAULT); //zapocni novu stranicu

            globalHeader(oOutput)
            globalFooter(oOutput)
           
           //1. DECISION NAME NASLOV
            oOutput.BeginParagraphF(getString("ID_STYLE_RD_HEADING_1"))
            oOutput.Output("Decision: " + decisionFunctions[j].ObjDef().Name(nLocale), getString("ID_STYLE_RD_DEFAULT"), 18, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0) 
            oOutput.EndParagraph()
            
            //Ispod naslova tekst
            oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))        
            oOutput.Output("Decision Name: ", getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
            oOutput.Output(decisionFunctions[j].ObjDef().Name(nLocale), getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)//ovako ispisi i ostale ispod
            oOutput.EndParagraph()
            
            //do tu !!!!!!!!
            oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))        
            oOutput.Output("Decision Type: ", getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
            oOutput.Output(getObjectAtrValue(decisionFunctions[j], AT_DECISION_TYPE), getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)//ovako ispisi i ostale ispod
            oOutput.EndParagraph()
            
            oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))        
            oOutput.Output("Decision Subtype: ", getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
            oOutput.Output(getObjectAtrValue(decisionFunctions[j], AT_DECISION_SUB_TYPE), getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)//ovako ispisi i ostale ispod
            oOutput.EndParagraph()
            
            oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))        
            oOutput.Output("Decision Criticality: ", getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
            oOutput.Output(getObjectAtrValue(decisionFunctions[j], AT_DECISION_CRITICALITY), getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)//ovako ispisi i ostale ispod
            oOutput.EndParagraph()
           
            emptyrow(1)
          
            
            // DECISION MAKER MATRIX (RASCI)
            oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
            oOutput.Output("Decision Maker (RASCI)", getString("ID_STYLE_RD_DEFAULT"), 14, RGB(0, 0, 0), RGB(198, 217, 241), Constants.FMT_BOLD | Constants.FMT_LEFT, 0) 
            oOutput.EndParagraph()
            
            emptyrow(1)
            
            matrixRASCI(decisionFunctions[j])
            
             
             
             
            emptyrow(1)
            
             //POLICY
            oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
            oOutput.Output("Policy ", getString("ID_STYLE_RD_DEFAULT"), 14, RGB(0, 0, 0), RGB(198, 217, 241), Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
            oOutput.EndParagraph()
                    
                    emptyrow(1)
                    
                    var policyObjects = decisionFunctions[j].getConnectedObjOccs(new Array(Constants.ST_TRANSFORM_OPERATOR_XOR,Constants.ST_OPR_XOR_1,Constants.ST_XOR,Constants.ST_OPR_XOR_INST), Constants.EDGES_OUT)
                    
                    for(var p=0;p<policyObjects.length;p++){
                    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
                    oOutput.OutputLnF("• "+policyObjects[p].ObjDef().Name(nLocale), getString("ID_STYLE_RD_DEFAULT"))
                    oOutput.EndParagraph()
                    }
            
            emptyrow(1)
            
            
            // RESULT OF DECISION
            oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
            oOutput.Output("Results of decision ", getString("ID_STYLE_RD_DEFAULT"), 14, RGB(0, 0, 0), RGB(198, 217, 241), Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
            oOutput.EndParagraph()
            
            var decisionEvents = getEventsFromDecisionFunctions(decisionFunctions[j]) //ulazi 1 funkcija vraca niz evenata
            
            emptyrow(1)
                    
             for(var p=0;p<decisionEvents.length;p++){
             oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
             oOutput.OutputLnF("• "+decisionEvents[p].ObjDef().Name(nLocale), getString("ID_STYLE_RD_DEFAULT"))
             oOutput.EndParagraph()
             }
            
           //CIA Cyber Security SECTION
             
             emptyrow(1)
             
            oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
            oOutput.Output("CIA Cyber Security: ", getString("ID_STYLE_RD_DEFAULT"), 14, RGB(0, 0, 0), RGB(198, 217, 241), Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
            oOutput.EndParagraph()
              
              emptyrow(1)
              
            getCSdata(decisionFunctions[j])   //cita podatke iz CSI
      
     
        oOutput.EndSection(); //zavrsi stranicu
      
    
       
    
}
    //var connectioNameFromPositionToFunctio = connectionName(positionsFromFunction[0])
    //getCSdata(decisionFunctions[0])   //cita podatke iz CSI      
}
}

function matrixRASCI(decisionFunctions){

     var positionsFromFunction = getPositionsFromFunction(decisionFunctions) //ulaz funkcija vraca niz pozicija i rola spojeni na funkciju
    
    if(positionsFromFunction != ""){
    
    oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
            oOutput.TableRow()
            oOutput.ResetFrameStyle()
            oOutput.TableCell("", 100 / 3, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(224,224,224), 0, Constants.FMT_VBOTTOM | Constants.FMT_CENTER, 0)
            oOutput.BeginParagraph(Constants.FMT_VBOTTOM | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
            oOutput.OutputLn("Decision / Position", getString("ID_DEFAULT_FONT"), 9.5, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VBOTTOM | Constants.FMT_CENTER | Constants.FMT_BOLD, 0.71)
            oOutput.EndParagraph()
            
            
           
            
            for(var k=0; k<positionsFromFunction.length;k++){
            oOutput.TableCell("", 100 / 3, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(249,182,0), 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
            oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
            oOutput.OutputLn(positionsFromFunction[k].ObjDef().Name(nLocale), getString("ID_DEFAULT_FONT"), 9.5, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_BOLD, 0.71)
            oOutput.EndParagraph()
            }
            
            
            oOutput.TableRow()
            oOutput.ResetFrameStyle()
            oOutput.TableCell("", 100 / 3, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(41,138,37), 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
            oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
            oOutput.OutputLn(decisionFunctions.ObjDef().Name(nLocale), getString("ID_DEFAULT_FONT"), 9.5, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER | Constants.FMT_BOLD, 0.71)
            oOutput.EndParagraph()
            
            //ispisi imena veze
            for(var k1=0; k1<positionsFromFunction.length;k1++){
            
            oOutput.TableCell("", 100 / 3, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
            oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
            oOutput.OutputLn(connectionName(positionsFromFunction[k1]), getString("ID_DEFAULT_FONT"), 9.5, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER | Constants.FMT_BOLD, 0.71)
            oOutput.EndParagraph()
            
            
            
            }
            
          oOutput.EndTable("", 100, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
    }
}


function getCSdata(decisionFunction){
    
    var selectedObjectOcc = decisionFunction.getConnectedObjOccs(new Array(Constants.ST_ATTR_1,Constants.ST_ATTR_3,Constants.ST_ATTR_LINK,Constants.ST_XML_ATTRTYPE), Constants.EDGES_INOUT) //dohvati objekte sa funkcije
    
    var selectedObject = new Array()
    for(var k2=0; k2<selectedObjectOcc.length;k2++){
             //if(selectedObjectOcc[k2].getFillColor() == -1){ //uzmi samo plave objekte (postoje zeleni i crveni jos)
             selectedObject.push(selectedObjectOcc[k2].ObjDef())
             }
    
for (var i = 0; i < selectedObject.length; i++) {
     
     var objectName = selectedObject[i].Name(nLocale) //ime objekta
     var classObjects = selectedObject[i].getConnectedObjs([Constants.OT_CLS], Constants.EDGES_INOUTASSIGN, new Array(Constants.CT_HAS_MEMB,Constants.CT_HAS_MEMBER))// belongs to class
     
     var classObjectsName = ""
     
     for(var k=0;k<classObjects.length;k++){
        classObjectsName = classObjectsName + classObjects[k].Name(nLocale) + ", "
     }
   
     
     var allModelsFromObject = collectModelsFromObject(selectedObject[i].OccList()) 
     
     var allModelsFromObjectName = ""
     
     for(var k=0;k<allModelsFromObject.length;k++){
        allModelsFromObjectName = allModelsFromObjectName + allModelsFromObject[k].Name(nLocale) + ", "
     }
     
     var epcModels = new Array(); //occurrences in process 

    for (var j = 0; j < allModelsFromObject.length; j++) { //u  varijablu orgChartModels spremi sve modele tipa organizational chart
        if (allModelsFromObject[j].TypeNum() == Constants.MT_EEPC) {
            epcModels.push(allModelsFromObject[j])
        }
    }
    
    var epcModelsName = ""
     
     for(var k=0;k<epcModels.length;k++){
        epcModelsName = epcModelsName + epcModels[k].Name(nLocale) + ", "
     }
     
     var allFunctionsFromObject = selectedObject[i].getConnectedObjs([Constants.OT_FUNC], Constants.EDGES_INOUTASSIGN, new Array(Constants.CT_IS_INP_FOR, Constants.CT_IS_INPUT_FOR, Constants.CT_HAS_OUT)) //input output of functions
     
     var allFunctionsFromObjectName = ""
     
     for(var k=0;k<allFunctionsFromObject.length;k++){
        allFunctionsFromObjectName = allFunctionsFromObjectName + allFunctionsFromObject[k].Name(nLocale) + ", "
     }
     
     var screenModels = new Array(); //all screen designs models

    for (var j = 0; j < allModelsFromObject.length; j++) { //u  varijablu orgChartModels spremi sve modele tipa organizational chart
        if (allModelsFromObject[j].TypeNum() == Constants.MT_SCREEN_DES || allModelsFromObject[j].TypeNum() == Constants.MT_SCREEN_DESIGN_SD ) {
            screenModels.push(allModelsFromObject[j])
        }
    }
     
     var screenModelsName = ""
     
     for(var k=0;k<screenModels.length;k++){
        screenModelsName = screenModelsName + screenModels[k].Name(nLocale) + ", "
     }
     
     
     
    var relatedScreensFromObject = selectedObject[i].getConnectedObjs(null, Constants.EDGES_INOUTASSIGN, [Constants.CT_IS_REPRESEN]) // related screens
     
      var relatedScreensFromObjectName = ""
     
     for(var k=0;k<relatedScreensFromObject.length;k++){
        relatedScreensFromObjectName = relatedScreensFromObjectName + relatedScreensFromObject[k].Name(nLocale) + ", "
     }
     
     //ISPIS PODATAKA NA STRANICU
            //ispisi ime objekta tj funkcije
            oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))        
            oOutput.Output((i+1)+"." + " Name of data asset: ", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
            oOutput.Output(objectName, getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)//ovako ispisi i ostale ispod
            oOutput.EndParagraph()
            emptyrow(1)
            
            //umetni tablicu sa CIA
            oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
            oOutput.TableRow()
            oOutput.ResetFrameStyle()
            oOutput.TableCell("", 100 / 3, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
            oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
            oOutput.OutputLn("Confidentiality", getString("ID_DEFAULT_FONT"), 9.5, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71)
            oOutput.EndParagraph()
            oOutput.TableCell("", 100 / 3, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
            oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
            oOutput.OutputLn("Integrity", getString("ID_DEFAULT_FONT"), 9.5, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71)
            oOutput.EndParagraph()
            oOutput.TableCell("", 100 / 3, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
            oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
            oOutput.OutputLn("Availability", getString("ID_DEFAULT_FONT"), 9.5, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71)
            oOutput.EndParagraph()
            
            oOutput.TableRow()
            oOutput.ResetFrameStyle()
            oOutput.TableCell("", 100 / 3, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
            oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
            oOutput.OutputLn(getObjectDefAtrValue(selectedObject[i], AT_CONFIDENTIALITY), getString("ID_DEFAULT_FONT"), 9.5, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71)
            oOutput.EndParagraph()
            oOutput.TableCell("", 100 / 3, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
            oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
            oOutput.OutputLn(getObjectDefAtrValue(selectedObject[i], AT_INTEGRITY), getString("ID_DEFAULT_FONT"), 9.5, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71)
            oOutput.EndParagraph()
            oOutput.TableCell("", 100 / 3, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
            oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
            oOutput.OutputLn(getObjectDefAtrValue(selectedObject[i], AT_AVAILABILITY_CUSTOM), getString("ID_DEFAULT_FONT"), 9.5, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71)
            oOutput.EndParagraph()
            
            oOutput.EndTable("", 100, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
            
            emptyrow(1)
            
            //ispisi Belongs to class
            oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))        
            oOutput.Output("Belongs to Data Class: ", getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
            oOutput.Output(classObjectsName, getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)//ovako ispisi i ostale ispod
            oOutput.EndParagraph()
            emptyrow(1)
            //ispisi Used in Process
            oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))        
            oOutput.Output("Data Asset Used in Following Processes: ", getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
            oOutput.Output(epcModelsName, getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)//ovako ispisi i ostale ispod
            oOutput.EndParagraph()
            emptyrow(1)
            //Input/Output of Function
            oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))        
            oOutput.Output("Data Asset Used in Following Functions: ", getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
            oOutput.Output(allFunctionsFromObjectName, getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)//ovako ispisi i ostale ispod
            oOutput.EndParagraph()
            emptyrow(1)
            //Role
            oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))        
            oOutput.Output("Data Asset Used by Following Roles: ", getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
            oOutput.Output(" ", getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)//ovako ispisi i ostale ispod
            oOutput.EndParagraph()
            emptyrow(1)
            
            //Used in Screen
            oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))        
            oOutput.Output("Data Asset Used in Following Screens: ", getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
            oOutput.Output(screenModelsName, getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)//ovako ispisi i ostale ispod
            oOutput.EndParagraph()
            emptyrow(1)
            //Related Field in Screen
            oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))        
            oOutput.Output("Related Field in Screen: ", getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
            oOutput.Output(relatedScreensFromObjectName, getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)//ovako ispisi i ostale ispod
            oOutput.EndParagraph()
            emptyrow(1)
            /*
            //Confidentiality
            oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))        
            oOutput.Output("Confidentiality: ", getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
            oOutput.Output(getObjectDefAtrValue(selectedObject[i], AT_CONFIDENTIALITY), getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)//ovako ispisi i ostale ispod
            oOutput.EndParagraph()
            emptyrow(1)
            //Integrity
            oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))        
            oOutput.Output("Integrity: ", getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
            oOutput.Output(getObjectDefAtrValue(selectedObject[i], AT_INTEGRITY), getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)//ovako ispisi i ostale ispod
            oOutput.EndParagraph()
            emptyrow(1)
            //AVAILABILITY
            oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))        
            oOutput.Output("Availability: ", getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
            oOutput.Output(getObjectDefAtrValue(selectedObject[i], AT_AVAILABILITY_CUSTOM), getString("ID_STYLE_RD_DEFAULT"), 12, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)//ovako ispisi i ostale ispod
            oOutput.EndParagraph()
            */
            emptyrow(2)
     
       //ispis ne u tablicu nego normalno jedno pokraj drugoga ili kako vec
       /*
            oOutput.TableRow()
            oOutput.ResetFrameStyle()
            oOutput.TableCell("", 100 / 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0)
            oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
            oOutput.OutputLn(objectName, getString("ID_DEFAULT_FONT"), 9.5, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71)
            oOutput.EndParagraph()
            oOutput.ResetFrameStyle()
            oOutput.TableCell("", 100 / 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_LEFT, 0)
            oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
            oOutput.OutputLn(classObjectsName, getString("ID_DEFAULT_FONT"), 9.5, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT , 0.71)
            oOutput.EndParagraph()
            oOutput.TableCell("", 100 / 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_LEFT, 0)
            oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
            oOutput.OutputLn(epcModelsName, getString("ID_DEFAULT_FONT"), 9.5, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT , 0.71)
            oOutput.EndParagraph()
            oOutput.TableCell("", 100 / 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_LEFT, 0)
            oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
            oOutput.OutputLn(allFunctionsFromObjectName, getString("ID_DEFAULT_FONT"), 9.5, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT , 0.71)
            oOutput.EndParagraph()
            oOutput.TableCell("", 100 / 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_LEFT, 0)
            oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
            oOutput.OutputLn(" ", getString("ID_DEFAULT_FONT"), 9.5, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT , 0.71)
            oOutput.EndParagraph()  
            oOutput.TableCell("", 100 / 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_LEFT, 0)
            oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
            oOutput.OutputLn(screenModelsName, getString("ID_DEFAULT_FONT"), 9.5, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT , 0.71)
            oOutput.EndParagraph()  
            oOutput.TableCell("", 100 / 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_LEFT, 0)
            oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
            oOutput.OutputLn(relatedScreensFromObjectName, getString("ID_DEFAULT_FONT"), 9.5, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT , 0.71)
            oOutput.EndParagraph()  
            oOutput.TableCell("", 100 / 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_LEFT, 0)
            oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
            oOutput.OutputLn(getObjectDefAtrValue(selectedObject[i], AT_CONFIDENTIALITY), getString("ID_DEFAULT_FONT"), 9.5, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT , 0.71)
            oOutput.EndParagraph()  
            oOutput.TableCell("", 100 / 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_LEFT, 0)
            oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
            oOutput.OutputLn(getObjectDefAtrValue(selectedObject[i], AT_INTEGRITY), getString("ID_DEFAULT_FONT"), 9.5, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT , 0.71)
            oOutput.EndParagraph()  
            oOutput.TableCell("", 100 / 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VCENTER | Constants.FMT_LEFT, 0)
            oOutput.BeginParagraph(Constants.FMT_VCENTER | Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
            oOutput.OutputLn(getObjectDefAtrValue(selectedObject[i], AT_AVAILABILITY_CUSTOM), getString("ID_DEFAULT_FONT"), 9.5, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT , 0.71)
            oOutput.EndParagraph()  
            //oOutput.EndTable("", 100, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0);
            */
            

}
    

   
    
}

function connectionName(position){ //vraca tip veze izmedju pozicije i funkcije

                    /*
                    switch (position.CxnOccList()[0].CxnDef().TypeNum()) {
                        case 65:
                            var ime_veze = "carries out";
                            break;
                        case 218:
                            var ime_veze = "carries out";
                            break;
                        case 232:
                            var ime_veze = "decides on";
                            break;
                        case 323:
                            var ime_veze = "decides on";
                            break;
                        case 233:
                            var ime_veze = "contributes to";
                            break;
                        case 324:
                            var ime_veze = "contributes to";
                            break;
                        case 266:
                            var ime_veze = "must be informed about";
                            break;
                        case 326:
                            var ime_veze = "must be informed about";
                            break;
                    }
                        */
   //var ime_veze = position.CxnOccList()[0].CxnDef().AttrList(nLocale, null)[0].GetValue(false) //ovo isto radi, ali uzima niz i prvi u nizu
   var ime_veze = position.CxnOccList()[0].CxnDef().Attribute(389,nLocale,true).getValue() //potrebno testirati sto ce se dogoditi ako veza ima vise atributa
   return ime_veze

}


function getPositionsFromFunction(decisionFunction){ //vraca niz pozicija i rola spojeni na funkciju
 
   var positions = decisionFunction.getConnectedObjOccs(new Array(Constants.ST_POS,Constants.ST_POS_1,Constants.ST_POS_2,Constants.ST_EMPL_TYPE,Constants.ST_PERS_TYPE_PIC,Constants.ST_PERSON_TYPE,Constants.ST_SOLAR_ROLE), Constants.EDGES_INOUT)
    
      
    
   return positions
}


function getEventsFromDecisionFunctions(decisionFunction){ //vraca niz evenata spojenih na decision funkciju (1 function --> array of events)
      
      if(decisionFunction.getConnectedObjOccs(new Array(Constants.ST_TRANSFORM_OPERATOR_XOR,Constants.ST_OPR_XOR_1,Constants.ST_XOR,Constants.ST_OPR_XOR_INST), Constants.EDGES_OUT) != ""){
        
      var xorRule = decisionFunction.getConnectedObjOccs(new Array(Constants.ST_TRANSFORM_OPERATOR_XOR,Constants.ST_OPR_XOR_1,Constants.ST_XOR,Constants.ST_OPR_XOR_INST), Constants.EDGES_OUT)
        
      }      

        for(var k=0;k<xorRule.length;k++){
         
        var decisionEvents = xorRule[k].getConnectedObjOccs(new Array(Constants.ST_EV,Constants.ST_EVENT_PIC,Constants.ST_EV_INST), Constants.EDGES_OUT)   
        }
        
  return decisionEvents
  
}

function getFunctionsConnectedToXorRule(allFunctionsFromEpc){
 
        var xorRule = new Array();
        var functionXorRule = new Array();
        
        for(var j=0;j<allFunctionsFromEpc.length;j++){
        
        if(allFunctionsFromEpc[j].getConnectedObjOccs(new Array(Constants.ST_TRANSFORM_OPERATOR_XOR,Constants.ST_OPR_XOR_1,Constants.ST_XOR,Constants.ST_OPR_XOR_INST), Constants.EDGES_OUT) != ""){
        
        xorRule.push(allFunctionsFromEpc[j].getConnectedObjOccs(new Array(Constants.ST_TRANSFORM_OPERATOR_XOR,Constants.ST_OPR_XOR_1,Constants.ST_XOR,Constants.ST_OPR_XOR_INST), Constants.EDGES_OUT))
        
        functionXorRule.push(allFunctionsFromEpc[j])
        
        }
        }
 
   return functionXorRule
    
}

function collectModelsFromObject(p_array) {
    var result = new Array();
    for (var i = 0; i < p_array.length; i++) {
        result.push(p_array[i].Model())
    }
    return ArisData.Unique(result)
}

function getObjectDefAtrValue(objectDefList, atribut) {
    //getObjectDefAtrValue(objectOccList[i], "AT_IME_I_PREZIME_NA_POZICIJI")  
    if (objectDefList.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum(atribut), nLocale).GetValue(false) != undefined) {
        var atrValue = objectDefList.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum(atribut), nLocale).GetValue(false)
        return atrValue
    } else {
        var atrValue = ""
        return atrValue
    }
}

function getObjectAtrValue(objectOccList, atribut) {
    //getObjectAtrValue(objectOccList[i], "AT_IME_I_PREZIME_NA_POZICIJI")  
    if (objectOccList.ObjDef().Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum(atribut), nLocale).GetValue(false) != undefined) {
        var atrValue = objectOccList.ObjDef().Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum(atribut), nLocale).GetValue(false)
        return atrValue
    } else {
        var atrValue = ""
        return atrValue
    }
}

function sortByPosition(oobjectocclist) {
    for (var i = 0; i < oobjectocclist.length - 1; i++)
        for (var j = i + 1; j < oobjectocclist.length; j++) {
            if (oobjectocclist[i].Y() > oobjectocclist[j].Y()) {
                var temp = oobjectocclist[i]
                oobjectocclist[i] = oobjectocclist[j]
                oobjectocclist[j] = temp
            } else if (oobjectocclist[i].Y() == oobjectocclist[j].Y() && oobjectocclist[i].X() > oobjectocclist[j].X()) {
                var temp = oobjectocclist[i]
                oobjectocclist[i] = oobjectocclist[j]
                oobjectocclist[j] = temp
            }
        }
    return oobjectocclist;
}

function emptyrow(number) {
    for (var i = 0; i < number; i++)
        oOutput.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
}

function RGB(r, g, b) {
    return (new java.awt.Color(r / 255.0, g / 255.0, b / 255.0, 1)).getRGB() & 0xFFFFFF
}