// GEAR_GetAttrValueTypeInConfig let op in aris_cs_nl is iets aangepast!!! klopt het nu nog wel???

var gsCRLF=String.fromCharCode(13) + String.fromCharCode(10);
var gsLFCR=String.fromCharCode(10) + String.fromCharCode(13);

// indexen van boolean attribute values    
var gc_IX_UNDEFINED  = 0;    
var gc_IX_TRUE  = 1;
var gc_IX_FALSE = 2;
eTrilian = {UNKNOWN:0, FALSE:1, TRUE:2}

var gHASHMAP = java.util.HashMap
var lhm = new gHASHMAP()

var glTimerIntervalInLog = 0
var gsDefaultLogFile = ""
var gbBreakScriptRun = false

function fNewLineChar(){
    return "\n"
}
if (fIsMissing(mCurrentDatabase)){
    var mCurrentDatabase = ArisData.getActiveDatabase()
}

function XLS_CloneCellStyle(pWB, pCellStyle){
    var lCellStyle2Return = null;
    try{
    lCellStyle2Return = pWB.createCellStyle(pWB.getFontAt(pCellStyle.getFontIndex())
                                          , pCellStyle.getBorderTop()
                                          , pCellStyle.getBorderRight()
                                          , pCellStyle.getBorderBottom()
                                          , pCellStyle.getBorderLeft()
                                          , pCellStyle.getTopBorderColor()
                                          , pCellStyle.getRightBorderColor()
                                          , pCellStyle.getBottomBorderColor()
                                          , pCellStyle.getLeftBorderColor()
                                          , pCellStyle.getAlignment()
                                          , pCellStyle.getVerticalAlignment()
                                          , pCellStyle.getFillBackgroundColor()
                                          , pCellStyle.getFillForegroundColor()
                                          , pCellStyle.getFillPattern())
    }catch(ex){
        var lex=ex
        var aa=0
    }
    return lCellStyle2Return;                                          
}
/*
function XLS_CopyCellStyle(pXSSFCell, pCellStyle){
    try{
        var lWB = pXXSFCell.getSheet().getWorkbook()
        var lStyle = pXXSFCell.getCellStyle()
        var lFontIndex = pCellstyle.getFontIndex()
        var lFont = lWB.getFontAt(lFontIndex)
        
        lStyle.setFont(pWB.getFontAt(pCellStyle.getFontIndex())
        lStyle.set, pCellStyle.getBorderTop()
        lStyle.set, pCellStyle.getBorderRight()
        lStyle.set, pCellStyle.getBorderBottom()
        lStyle.set, pCellStyle.getBorderLeft()
        lStyle.set, pCellStyle.getTopBorderColor()
        lStyle.set, pCellStyle.getRightBorderColor()
        lStyle.set, pCellStyle.getBottomBorderColor()
        lStyle.set, pCellStyle.getLeftBorderColor()
        lStyle.set, pCellStyle.getAlignment()
        lStyle.set, pCellStyle.getVerticalAlignment()
        lStyle.set, pCellStyle.getFillBackgroundColor()
        lStyle.set, pCellStyle.getFillForegroundColor()
        lStyle.set, pCellStyle.getFillPattern())
    }catch(ex){}
    return lCellStyle2Return;                                          
}
*/


function XLS_CopyRow(pBaseRow, pRow){
    var lColNr = 0
    while (lColNr<pBaseRow.getLastCellNum()){
        var lCell = pBaseRow.createCell(lColNr)
        XLS_CopyBaseCell(lCell, pRow)
        lColNr++
    }
}
function XLS_CopyBaseCell(pBaseCell, pRow){
    var lNewCell = null
    try{
        var lNewCell=pRow.createCell(pBaseCell.getCellIndex());
    }catch(ex){
        var lex=ex;
        var aa=1;
        return null
    }
    if (lNewCell.getClass().indexOf("poi.xssf")>-1){
        lNewCell.setCellType(pBaseCell.getCellType().getType());
        var lStyle = XLS_CloneCellStyle(lNewCell.getSheet().getWorkbook(), pBaseCell.getCellStyle())
        lNewCell.setCellStyle(lStyle())
    }else if (lNewCell.getClass() == pBaseCell.getClass()){
        lNewCell.setCellType(pBaseCell.getCellType())       
    }     
    return lNewCell
}
function XLS_DeleteColumn(pSheet, plColIndexToDelete ){
    
    var lMaxColumn = 0;
    var lbArisClass = (pSheet.getClass()== com.idsscheer.report.output.excelaccess.logic.AXlsSheet) ? true: false
    
    if (lbArisClass){
        var lLastRowNum = pSheet.getLastFilledRowNum()
    }else{
        var lLastRowNum = pSheet.getLastRowNum()
    }
    for (var lRowIdx= 0; lRowIdx < lLastRowNum + 1; lRowIdx++ ){
        if (lbArisClass){
            var lRow = pSheet.getRowAt(lRowIdx)
        }else{
            var lRow = pSheet.getRow(lRowIdx)
        }

        // if no row exists here; then nothing to do; next!
        if (lRow == null ) continue;

        // if the row doesn't have this many columns then we are good; next!
        if (lbArisClass){
            var lLastColumn = lRow.getLastCellNum();
        }else{
            var lLastColumn = lRow.getLastCellNum() -1 // Gets the index of the last cell contained in this row PLUS ONE.
        }
        if (lLastColumn > lMaxColumn )
            lMaxColumn = lLastColumn;

        if ( lLastColumn < plColIndexToDelete )continue;

        for (var lColIdx = plColIndexToDelete + 1; lColIdx < lLastColumn + 1; lColIdx++ ){
            if (lbArisClass){
                var lOldCell = lRow.getCellAt(lColIdx - 1);
            }else{
                var lOldCell = lRow.getCell(lColIdx - 1);
            }
            if (lOldCell != null ){
                if (lbArisClass){
                    lRow.removeCell(lColIdx - 1)
                }else{
                    lRow.removeCell(lOldCell );
                }
            }
            if (lbArisClass){
                var lNextCell = lRow.getCellAt(lColIdx);
            }else{
                var lNextCell = lRow.getCell(lColIdx);
            }
            if (lNextCell != null ){
                if (lbArisClass){
                    var lNewCell = lRow.createCell(lColIdx -1)
                    lNewCell.setCellType(lNextCell.getCellType());
                }else{
                    var lNewCell = lRow.createCell(lColIdx -1, lNextCell.getCellType() );
                }
                    
                XLS_CloneCell(lNewCell, lNextCell);
                
                if (lbArisClass){
                    lNextCell.setCellType(Constants.XL_CELL_TYPE_BLANK)
                    lRow.removeCell(lColIdx)
                }else{
                     if (lNewCell.getClass().indexOf("poi.xssf")>-1){
                         var POICELL = org.apache.poi.hssf.usermodel.HSSFCell
                    }else{
                         var POICELL = org.apache.poi.xssf.usermodel.XSSFCell
                    }  
                    lNextCell.setCellType(POICELL.CELL_TYPE_BLANK);
                    lRow.removeCell(lNextCell );
                    
                }
            }
            if (lColIdx == lLastColumn -1){
                var aa=0
            }
        }
    }


    // Adjust the column widths
    for (var lColIdx = plColIndexToDelete; lColIdx<lMaxColumn; lColIdx++ ){
        pSheet.setColumnWidth(lColIdx, pSheet.getColumnWidth(lColIdx + 1) );
    }
    
}


    /*
     * Takes an existing Cell and merges all the styles and forumla
     * into the new one
     */
function XLS_CloneCell(pNewCell, pOldCell){
    var lbArisClass = (pNewCell.getClass()== com.idsscheer.report.output.excelaccess.logic.AXlsCell) ? true: false
    if (!lbArisClass){    
        if (lNewCell.getClass().indexOf("poi.xssf")>-1){
            var POICELL = org.apache.poi.hssf.usermodel.HSSFCell
        }else{
            var POICELL = org.apache.poi.xssf.usermodel.XSSFCell
        }        
    }
    if (lbArisClass){
        if (pOldCell.getCellComment()){
            pNewCell.setCellComment( pOldCell.getCellComment().getText(), pOldCell.getCellComment().getAuthor() );
        }else{
            pNewCell.removeCellComment()
        }
    }else{
        pNewCell.setCellComment( pOldCell.getCellComment() );
    }
    pNewCell.setCellStyle( pOldCell.getCellStyle() );

    if (lbArisClass){ 
        switch ( pNewCell.getCellType() ){
        case Constants.XL_CELL_TYPE_BOOLEAN:
            pNewCell.setCellValue( pOldCell.getBooleanCellValue() );
            break;
            
        case Constants.XL_CELL_TYPE_NUMERIC:
            pNewCell.setCellValue( pOldCell.getNumericCellValue() );
            break;
        
        case Constants.XL_CELL_TYPE_STRING:
            pNewCell.setCellValue( pOldCell.getStringCellValue() );
            break;
        
        case Constants.XL_CELL_TYPE_ERROR:
            pNewCell.setCellValue( pOldCell.getErrorCellValue() );
            break;
        
        case Constants.XL_CELL_TYPE_FORMULA:
            pNewCell.setCellFormula( pOldCell.getCellFormula() );
            break;
        }        
    }else{
        switch ( pNewCell.getCellType() ){
            case POICELL.CELL_TYPE_BOOLEAN:{
                pNewCell.setCellValue( pOldCell.getBooleanCellValue() );
                break;
            }
            case POICELL.CELL_TYPE_NUMERIC:{
                pNewCell.setCellValue( pOldCell.getNumericCellValue() );
                break;
            }
            case POICELL.CELL_TYPE_STRING:{
                pNewCell.setCellValue( pOldCell.getStringCellValue() );
                break;
            }
            case POICELL.CELL_TYPE_ERROR:{
                pNewCell.setCellValue( pOldCell.getErrorCellValue() );
                break;
            }
            case POICELL.CELL_TYPE_FORMULA:{
                pNewCell.setCellFormula( pOldCell.getCellFormula() );
                break;
            }
        }
    }
}
function XLS_SetAutoFilter(psTempPath, plRowIdx){

    if (Files_ExtractExtention(Context.getSelectedFile())=="xlsx"){

        var lXLS_workBook = new org.apache.poi.xssf.usermodel.XSSFWorkbook(new java.io.FileInputStream(psTempPath + "/" + Context.getSelectedFile()))

    }else{

        var lXLS_workBook = new org.apache.poi.hssf.usermodel.HSSFWorkbook(new java.io.FileInputStream(psTempPath + "/" + Context.getSelectedFile()))        

    }

    var lWS = lXLS_workBook.getSheetAt(0)

    var lFirstRowNum = fValueIfMissing(plRowIdx, lWS.getFirstRowNum())
    var lLastRowNum = lWS.getLastRowNum()
    var lFirstRow = lWS.getRow(lFirstRowNum)        
    var lFirstColNum = lFirstRow.getFirstCellNum()
    var lLastColNum = lFirstRow.getLastCellNum() - 1       // getLastCellNum : Gets the index of the last cell contained in this row PLUS ONE.

    lWS.setAutoFilter(new org.apache.poi.ss.util.CellRangeAddress(lFirstRowNum, lLastRowNum, lFirstColNum, lLastColNum))

    lXLS_workBook.write(new java.io.FileOutputStream(psTempPath + "/" + Context.getSelectedFile()))

}


function XML_GetDocumentByFileName(psFileName){
    var lDocument2Return=null;
    
    var lContent = Context.getFile(psFileName, Constants.LOCATION_SCRIPT)
    if (lContent.length>0){
        try{
            var lXMLStream = new Packages.java.io.ByteArrayInputStream(lContent);
            var lBuilder = new Packages.org.jdom.input.SAXBuilder();
            lDocument2Return = lBuilder.build(lXMLStream);
        }catch(ex){
            var lex=ex;
            //AAS_ErrorDialog("Error in Attributes value(s)\nXML syntax invalid! Error message: " + ex, "ATSC_Settings");
        }
    }
    return lDocument2Return;
}

function XML_GetRootElement(pDocument){
    return pDocument.getRootElement()  
}
function XML_GetChild(pParent, psName){
    var lElement = null
    var lsName = fValueIfMissing(psName,"")
    if (lsName>""){
        try{
            lElement = pParent.getChild(new Packages.java.lang.String(lsName))
        }catch(ex){}
    }
    return lElement
}
function XML_GetChilds(pParent){
    var lChildElements = []
    var lElement = null
    try{
        lChildElements = pParent.getChildren().toArray()
        
    }catch(ex){}

    return lChildElements
}


function XML_Element_GetAttrValue(pElement, psKeySet, psDefault, pbKeyCaseSensitive){
    var ls2Return = "";

    lsDefault = fValueIfMissing(psDefault, "") + ""
    lbKeyCaseSensitive = fValueIfMissing(pbKeyCaseSensitive, true)
    if (lbKeyCaseSensitive!=false)lbKeyCaseSensitive=true
        
    var lsKeys= new Array()
    if (fIsArray(psKeySet)){
        var lsKeys = psKeySet
    }else{
        lsKeys[0]=psKeySet
    }
    
    var lbFound = false;
    
    if (!lbKeyCaseSensitive){
        for (var i in lsKeys){
            var lsKey = (lsKeys[i]+"").toUpperCase()
            try{
                var lAttributes = pElement.getAttributes()
                for (var j=0; j<lAttributes.size();j++){
                    if (lAttributes.get(j).name.toUpperCase()==lsKey){
                        ls2Return = lAttributes.get(j).value + ""
                        lbFound = true
                        break
                    }
                }
            }
            catch(ex){
                var lex=ex;
                var a=1;
                ls2Return = lsDefault;
            }
            if (lbFound)break
        }
    }else{
        for (var i in lsKeys){        
            try{
                ls2Return = pElement.getAttribute(lsKeys[i]).getValue()+""
                lbFound=true
            }
            catch(ex){
                var lex=ex;
                var a=1;
            }
            if (lbFound)break
        }
    }
    if (!lbFound)ls2Return = lsDefault;
    
    return (ls2Return + "");
}

function GEAR_GetAttrValueInConfig(psPath2Element, psAttrName){
    var lsValue =""
    var lXMLDoc = XML_GetDocumentByFileName(gsSettingsFile)
    if (lXMLDoc){
        var lsPath2Element = (psPath2Element + "").replace(/\\/g,"/")+"";              
        var lsNames = lsPath2Element.split("/")
        var lElement = XML_GetRootElement(lXMLDoc)
        
        if (lsNames[0]+""==lElement.getName()){
            lsNames =lsNames.slice(1)
        }
            
        for each (var lsName in lsNames){
            lElement = XML_GetChild(lElement, lsName)
            if (lElement==null){
                break
            }
        }    
        if (lElement){
            lsValue = XML_Element_GetAttrValue(lElement, psAttrName)
        }
    }
    
    return lsValue
}

function GEAR_GetChildNamesInConfig(psSection){
    var lsChildNames=[]
    var lXMLDoc = XML_GetDocumentByFileName(gsSettingsFile)
    if (lXMLDoc){
        var lRootElement = XML_GetRootElement(lXMLDoc)
        var lsSection = (psSection + "").replace(/\\/g,"/")+"";              
        var lSection =  XML_GetChild(lRootElement, lsSection)
        var lChilds = XML_GetChilds(lSection)
        for each (var lChild in lChilds){
            lsChildNames.push(lChild.getName()+"")
        }
    }
    return lsChildNames
    
}

function GEAR_GetValueInConfig(psPath2Element, psEntry){
    if (fIsMissing(psEntry)){
        // psSection =  psPath2Element
        var lsPath2Element = psPath2Element
        var lXMLDoc = XML_GetDocumentByFileName(gsSettingsFile)
        if (lXMLDoc){
            var lsPath2Element = (lsPath2Element + "").replace(/\\/g,"/")+"";              
            var lsNames = lsPath2Element.split("/")
            var lElement = XML_GetRootElement(lXMLDoc)
            
            if (lsNames[0]+""==lElement.getName()){
                lsNames =lsNames.slice(1)
            }
                
            for each (var lsName in lsNames){
                lElement = XML_GetChild(lElement, lsName)
                if (lElement==null){
                    break
                }
            }    
            if (lElement){
                return fTrim(XML_Element_GetText(lElement, "", ""))
            }
        }
         
    }else{
        var lsSection = psPath2Element
        return fTrim(Context.getPrivateProfileString(lsSection, psEntry,"",gsSettingsFile))
    }
    return ""
}




         
function GEAR_GetTypeInConfig(psSection, psEntry){
    
    var lType = -1    

   
    var lsValue = Context.getPrivateProfileString(psSection, psEntry,"",gsSettingsFile)
    lsValue = fTrim(lsValue)
    if (lsValue=="")lsValue=psEntry
    switch(psSection.toUpperCase()){
    case "MODELTYPES":
        lType = Aris_GetType("M", lsValue)
        break
    case "OBJECTTYPES":        
        lType = Aris_GetType("", lsValue)    
        break
    case "ATTRIBUTETYPES":    
        lType = Aris_GetType("A", lsValue)
        break
    case "SYMBOLTYPES":    
        lType = Aris_GetType("S", lsValue)
        break        
    case "RELATIONTYPES":
        lType = Aris_GetType("", lsValue)    
        break    
    }
    return lType
}
function Aris_IsValid(pObject){
    var lb2Return = false;
    try{
        if (fIsMissing(pObject)){
              // niet goed
        }else if (pObject.IsValid()){
            lb2Return = true;
        }
    }catch(ex){}
        
    return lb2Return;
}
function Aris_Object4Attr(pObject){
    var lObject2Return = pObject;   ///default
    
    if (Aris_IsValid(pObject)){
        var lKindNum = Aris_GetKindNum(pObject);        
        switch (lKindNum){
        case Constants.CID_OBJOCC:
           lObject2Return = pObject.ObjDef();  break;
        case Constants.CID_CXNOCC:            
            lObject2Return = pObject.CxnDef();  break;
        }
    }
    return lObject2Return;
}
function Aris_GetName(pObject, plLoc){    
    var lsName2Return ="";
    if (fIsMissing(pObject))return "";
    
    try{    
        var lObject4Attr = Aris_Object4Attr(pObject);
        var lLoc = fValueIfMissing(plLoc,gnLoc);    
    
        if (fcLeft(pObject.toString() + "",("ConfigurationFilter").length)=="ConfigurationFilter"){
            lsName2Return = pObject.getName();
        }else{
            if (Aris_IsValid(lObject4Attr)){
                if (lObject4Attr.KindNum() == Constants.CID_USER){
                    lsName2Return = Aris_GetAttrValue(lObject4Attr, Constants.AT_NAME_LGINDEP, lLoc, true); 
                }else if (lObject4Attr.KindNum() == Constants.CID_USERGROUP){
                    lsName2Return = Aris_GetAttrValue(lObject4Attr, Constants.AT_NAME_LGINDEP, lLoc, true); 
                }else{
                    lsName2Return = Aris_GetAttrValue(lObject4Attr, Constants.AT_NAME, lLoc, true); 
                }
            }
        }
    }catch(ex){var lex = ex;
    var aa=1;
    }    
    return lsName2Return;
}

function Aris_GetKindNum(pObject){
    var lKIndNum2Return= 0;
    if (!fIsMissing(pObject)){
        try {
            lKIndNum2Return = pObject.KindNum()
        }catch(ex){
            var lex=ex;
        }
    }
    return lKIndNum2Return;
}
function Aris_SetAttrValue(pObject, plAttrNr, psAttrValue, plLoc){
    var lBool2Return = false
    
    if (fIsMissing(plAttrNr))return;
    if (plAttrNr==-1)return;
    
    var lsAttrValue = fValueIfMissing(psAttrValue,"");
    var lLoc = fValueIfMissing(plLoc,-1);
    
    var lObject4Attr = Aris_Object4Attr(pObject)
    if (Aris_IsValid(lObject4Attr)){            
        
        var lAttrBaseType = Aris_AttributeBaseType(plAttrNr);
        if (lLoc==-1){
            // zet alle talen
            var lDBLanguages = lObject4Attr.Database().LanguageList();
            for (var i=0; i<lDBLanguages.length; i++){                            
                var lLocaleID = lDBLanguages[i].LocaleId();
                var lAttr = lObject4Attr.Attribute(plAttrNr, lLocaleID);    
                lBool2Return=Aris_Attr_SetValue(lAttr, lAttrBaseType, lsAttrValue);
            }            
        }else{
            var lAttr = lObject4Attr.Attribute(plAttrNr, lLoc);    
            lBool2Return=Aris_Attr_SetValue(lAttr, lAttrBaseType, lsAttrValue);
        }
    }
    return lBool2Return
}

function Aris_Attr_SetValue(pAttribute,  plAttrBaseType, psValue){
    var lBool2Return = false
    
    var lsValue     = "";
    var lBoolean    = false;
  
    switch (plAttrBaseType){
    case Constants.ABT_BOOL:
        if (psValue==""){
            lBool2Return=pAttribute.Delete();
        }else{
            var lBoolean = fcBool(psValue)
            switch (lBoolean){
            case true:
                lBool2Return=pAttribute.setValue("1");   break;
            case false:
                lBool2Return=pAttribute.setValue("0");   break;
            }
        }
        break;
    case Constants.ABT_TIMESTAMP:
        lsValue = "" + psValue.toString() + "";
        if (lsValue == ""){
            lBool2Return=pAttribute.Delete();
        }else{
            lBool2Return=pAttribute.setValue(lsValue);
        }
        break;
    case Constants.ABT_DATE:
        lsValue = "" + psValue.toString() + "";
        if (lsValue == ""){
            lBool2Return=pAttribute.Delete();
        }else{
            lBool2Return=pAttribute.setValue(new Date(lsValue));
        }
        break;        
    default:    
        lsValue = "" + psValue.toString() + "";
        if (lsValue == ""){
            lBool2Return=pAttribute.Delete();
        }else{
            lBool2Return=pAttribute.setValue(lsValue);
        }
        break;
    }
    return lBool2Return
}
function Aris_IsPseudoAttribute(plAttrNum){
    var lbValue2Return=false;
    if(plAttrNum>=5000 && plAttrNum<=32000){
        lbValue2Return=true;
    }
    return lbValue2Return;
}
    

function Aris_GetAttrBoolValue(pObject, plAttrNr, plLoc, pbDefault){
    var lbResult=false;
    
    var lObject4Attr = Aris_Object4Attr(pObject)
    
    var lLoc = fValueIfMissing(plLoc,gnLoc);
    
    if (Aris_IsValid(lObject4Attr)){            
        
        if (!Aris_IsPseudoAttribute(plAttrNr)){
            var lAttr = lObject4Attr.Attribute(plAttrNr, lLoc)
    
            
            if(lAttr.IsMaintained())
            {
                if ((lAttr.MeasureUnitTypeNum() == Constants.AVT_ONE )
                ||  (lAttr.MeasureUnitTypeNum() == Constants.AVT_TRUE_2)
                ||  (lAttr.MeasureUnitTypeNum() == 2)){
                    lbResult=true;
                }else{
                    lbResult = false;
                }
            }else{            
                if (fcBool(pbDefault)){
                    lbResult=true;
                }else{
                    lbResult=false;
                }
            }            
        }else{
            // AAS-pseudoattribute            
            lbResult= fcBool(Aris_GetAttrValue(pObject, plAttrNr, plLoc, true, true));
        }
    }
    return lbResult;       
}


function Aris_GetAttrValue(pObject, plAttrNum){
    var lsValue = ""
    var lAttr = pObject.Attribute(plAttrNum, gnLoc)
    if (lAttr.IsMaintained()){
        lsValue = lAttr.getValue()
    }
    return lsValue
}
function Aris_AttributeBaseType(plAttrNum){
    var lType2Return = -1;
    try{
        if (plAttrNum!=45){
            // 45 geeft een keiharde fout
            if (fIsMissing(mCurrentDatabase)){
                lType2Return = ArisData.getActiveDatabase().ActiveFilter().AttrBaseType(plAttrNum);
            }else{            
                var lType2Return = mCurrentDatabase.ActiveFilter().AttrBaseType(plAttrNum);
            }
        }
    }catch(ex){}
    return lType2Return;
}
function Aris_GetAttrValueTypes(plAttrNr, pOptionalMethodFilter){
    var lsValueTypes2Return = new Array();
     
    if (plAttrNr<=0)return lsValueTypes2Return;
    
    var lMethodFilter = fIsMissing(pOptionalMethodFilter) ? ArisData.getActiveDatabase().ActiveFilter() : pOptionalMethodFilter
    
    var lAttrBaseType = Aris_AttributeBaseType(plAttrNr);
    switch(lAttrBaseType){
    case Constants.ABT_BOOL:
        lsValueTypes2Return[gc_IX_UNDEFINED]="";
         var lAttrValueTypeNums = lMethodFilter.AttrValueTypeNums(plAttrNr)
        if (lAttrValueTypeNums.length==2){
            lsValueTypes2Return[gc_IX_FALSE] = lMethodFilter.AttrValueType(plAttrNr, lAttrValueTypeNums[0])
            lsValueTypes2Return[gc_IX_TRUE] = lMethodFilter.AttrValueType(plAttrNr, lAttrValueTypeNums[1])
        }else{
            switch(plAttrNr){
            case Constants.AT_USER_ATTR3:    
            case Constants.AT_USER_ATTR4:            
                lsValueTypes2Return[gc_IX_FALSE] = lMethodFilter.AttrValueType(plAttrNr,Constants.AVT_NULL);
                lsValueTypes2Return[gc.IX_TRUE] = lMethodFilter.AttrValueType(plAttrNr,Constants.AVT_ONE);
                break;
            default:
                lsValueTypes2Return[gc_IX_FALSE] = lMethodFilter.AttrValueType(plAttrNr,Constants.AVT_FALSE_2);
                lsValueTypes2Return[gc_IX_TRUE] = lMethodFilter.AttrValueType(plAttrNr,Constants.AVT_TRUE_2);
            }
        }
        break;
    case Constants.ABT_VALUE:
    case Constants.ABT_COMBINED:
    case Constants.ABT_LONGTEXT:
        var lTypeNums = lMethodFilter.AttrValueTypeNums(plAttrNr);
        for (var i=0;i<lTypeNums.length;i++){
            var lsValue = "" + fTrim(lMethodFilter.AttrValueType(plAttrNr,lTypeNums[i]));
            if (lsValue.length>0)lsValueTypes2Return.push(lsValue);
        }
        break;
    }
    return lsValueTypes2Return;    
}



function Aris_GetType(psCode,psGuidOrType){
    var lTypeNr2Return = -1;
    var lsGuidOrType=fTrim(psGuidOrType).toLowerCase()
    if (fcIsGUID(lsGuidOrType)){
        var lsCode = (psCode+"").toUpperCase();
        if (lsCode=="M")lTypeNr2Return= Aris_UserDefinedModelTypeNum(lsGuidOrType);
        if (lsCode=="S")lTypeNr2Return= Aris_UserDefinedSymbolTypeNum(lsGuidOrType);
        if (lsCode=="A")lTypeNr2Return= Aris_UserDefinedAttributeTypeNum(lsGuidOrType);
        if (lsCode=="AVT")lTypeNr2Return= Aris_UserDefinedAttributeTypeNum(lsGuidOrType);
    }else if(isNaN(psGuidOrType)){
        lTypeNr2Return=Constants[psGuidOrType.toUpperCase()]                
    }else{
        lTypeNr2Return=psGuidOrType * 1;
    }
    if (isNaN(lTypeNr2Return)){
        lTypeNr2Return=-1
    }
    return lTypeNr2Return;
}

function Aris_UserDefinedModelTypeNum(psGUID){
    var lTypeNr2Return = 0;
    try{lTypeNr2Return = mCurrentDatabase.ActiveFilter().UserDefinedModelTypeNum (psGUID);
    }catch(ex){var lex=ex;
        lTypeNr2Return = ArisData.getActiveDatabase().ActiveFilter().UserDefinedModelTypeNum (psGUID);
        var aa=1;
    }
    return lTypeNr2Return;
}
function Aris_UserDefinedAttributeValueTypeNum(pAttrNrOrGUID, psGUID){
    var lTypeNum = -1
    try{
        if (!fcIsGUID(pAttrNrOrGUID)){
            var lsAttrGUID = mCurrentDatabase.ActiveFilter().UserDefinedAttributeTypeGUID(plAttrNr)
        }else{
            var lsAttrGUID = pAttrNrOrGUID
        }
    }catch(ex){
        return lTypeNum
    }
    try{
        lTypeNum =mCurrentDatabase.ActiveFilter().UserDefinedAttributeValueTypeNum(lsAttrGUID, psGUID)
    }catch(ex){}
    return lTypeNum
}
 
function Aris_UserDefinedSymbolTypeNum(psGUID){
    var lTypeNr2Return = 0;
    try{lTypeNr2Return = mCurrentDatabase.ActiveFilter().UserDefinedSymbolTypeNum (psGUID);
    }catch(ex){var lex=ex;
        var aa=1;
    }
    return lTypeNr2Return;
}
function Aris_UserDefinedAttributeTypeNum(psGUID){
    var lTypeNr2Return = 0;
    try{lTypeNr2Return = mCurrentDatabase.ActiveFilter().UserDefinedAttributeTypeNum (psGUID);
    }catch(ex){var lex=ex;
        var aa=1;
    }
    return lTypeNr2Return;
}
function Aris_GetAttrValueType(pMethodFilter, plAttrNr, plLoc, psValueTypeNum){
    var lsValue2Return = "";
    
    if (plAttrNr<=0)return lsValue2Return;
    
    var lMethodfilter = fIsMissing(pMethodFilter) ? ArisData.getActiveDatabase().AtiveFilter() : pMethodFilter;    
    var lLoc = fIsMissing(plLoc) ? -1 : plLoc;        // wordt niet gebruikt!!
    
    var lAttrBaseType = ArisData.getActiveDatabase().ActiveFilter().AttrBaseType(plAttrNr);

    if (lAttrBaseType==Constants.ABT_VALUE 
    || lAttrBaseType==Constants.ABT_COMBINED   
    || lAttrBaseType==Constants.ABT_LONGTEXT    ){ 
        lsValue2Return = pMethodFilter.AttrValueType(plAttrNr,psValueTypeNum);        
    }
    return lsValue2Return;    
}
function Aris_GetAttrBoolValue(pObject, plAttrNr, plLoc, pbDefault){
    var lbResult=false;
    
    var lObject4Attr = Aris_Object4Attr(pObject)
    
    var lLoc = fValueIfMissing(plLoc,gnLoc);
    
    if (Aris_IsValid(lObject4Attr)){            
        
        if (!Aris_IsPseudoAttribute(plAttrNr)){
            var lAttr = lObject4Attr.Attribute(plAttrNr, lLoc)
    
            
            if(lAttr.IsMaintained())
            {
                if ((lAttr.MeasureUnitTypeNum() == Constants.AVT_ONE )
                ||  (lAttr.MeasureUnitTypeNum() == Constants.AVT_TRUE_2)
                ||  (lAttr.MeasureUnitTypeNum() == 2)){
                    lbResult=true;
                }else{
                    lbResult = false;
                }
            }else{            
                if (fcBool(pbDefault)){
                    lbResult=true;
                }else{
                    lbResult=false;
                }
            }            
        }else{
            // AAS-pseudoattribute            
            lbResult= fcBool(Aris_GetAttrValue(pObject, plAttrNr, plLoc, true, true));
        }
    }
    return lbResult;       
}


function GEAR_GetAttrValueTypeInConfig(psEntryAttr, psEntryAttrValueType) {  
    var lsGUIDOfAttr = fTrim(Context.getPrivateProfileString("AttributeTypes", psEntryAttr,"",gsSettingsFile))
    var lsGUIDOfAttrValue = fTrim(Context.getPrivateProfileString("AttrValueTypeNums", psEntryAttrValueType,"",gsSettingsFile)    )
    var lAttrValueTypeNum = fcLng(Aris_GetAttrValueTypeNum(lsGUIDOfAttr, lsGUIDOfAttrValue))  // is met 1 verhoogd per 23-8-2016  klopt het nu nog wel????
    try {
        if (fcIsGUID(lsGUIDOfAttr)){
            var lAttrNum = mCurrentDatabase.ActiveFilter().UserDefinedAttributeTypeNum(lsGUIDOfAttr)  // letop: functie in cs_NL_Akr is aangepast
        }else{
            var lAttrNum = fcLng(lsGUIDOfAttr)            
        }
        return mCurrentDatabase.ActiveFilter().AttrValueType(lAttrNum, lAttrValueTypeNum)
    }catch(ex){
        return ""
    }
} 



function Aris_GetAttrValueTypeNum(psGUIDOrTypeNumOfAttr, psGUIDOfAttrValue){
    var lAttrValueTypeNum2Return = -1
    try {

        if (fcIsGUID(psGUIDOrTypeNumOfAttr)){
            var lsGUIDOfAttr = psGUIDOrTypeNumOfAttr
            var lAttrNum = mCurrentDatabase.ActiveFilter().UserDefinedAttributeTypeNum(lsGUIDOfAttr)
        }else{
            var lAttrNum = fcLng(psGUIDOrTypeNumOfAttr)            
            var lsGUIDOfAttr = mCurrentDatabase.ActiveFilter().UserDefinedAttributeTypeGUID(lAttrNum)
        }
        
        lAttrValueTypeNum2Return = mCurrentDatabase.ActiveFilter().UserDefinedAttributeValueTypeNum(lsGUIDOfAttr, psGUIDOfAttrValue)
        
    }catch(ex){var lex=ex
        var aa=1       
    }
    return lAttrValueTypeNum2Return
}  

function Aris_ModelIsOpen(pModel){ 
    var lBool = false
    ArisData.getActiveDatabase().setAutoTouch(false);
    try{
        pModel.setBgColor(pModel.getBgColor());
        ArisData.getActiveDatabase().setAutoTouch(true);    
        var adhoc =pModel.ObjOccList()
    }
    catch(e){
        lBool = true
    }
    return lBool
}
function tCxnOccAndObjOcc(){
    this.ObjOcc = null;
    this.CxnOcc = null;
    this.Direction = 0;          //krijgt waarde Constants.EDGES_IN of Constants.EDGES_OUT;
    this.bOnAllocDiag = false;   // = true als connectie op FAD, of andere AllocDiagram
}  

function Aris_GetCxnsAndObjs(pObjOcc, pbInclOn_Alloc_Diagram, plDirection, plCxnTypes){
    var ltCxnOccAndObjOccs2Return = new Array();
    
    var lbInclOn_Alloc_Diagram  = fValueIfMissing(pbInclOn_Alloc_Diagram, false)    
	var lDirection				= fValueIfMissing(plDirection, Constants.EDGES_INOUT)
    var lCxnTypes				= fValueIfMissing(plCxnTypes, [])
    
    var lActObjOccs = new Array();          // array waarin alle objocc's die onderocht moeten worden
                                            // worden gestopt. Naast pObjOcc ook ObjOcc's op FAD's
    
    // verzamel alle ActObjOcc's
    lActObjOccs.push(pObjOcc);
    
    if (lbInclOn_Alloc_Diagram) {
        
        var lAssModels=pObjOcc.ObjDef().AssignedModels()
        for (var i in lAssModels){
            var lAssModel = lAssModels[i]
            if (Aris_ModelIsAllocDgm(lAssModel)){
                var lFADObjOccs = pObjOcc.ObjDef().OccListInModel(lAssModel)
                for (var j in lFADObjOccs){
                    lActObjOccs.push(lFADObjOccs[j]);
                }        
            }
        }
    }
        
    // per verzamelde ActObjOcc de relaties verzamelen
    for (var i=0; i<lActObjOccs.length;i++){        
    
        for (var lkeer =1; lkeer<=2; lkeer ++){
            
            if (lkeer==1)var lEdges = Constants.EDGES_IN;
            if (lkeer==2)var lEdges = Constants.EDGES_OUT;
            if ((lEdges & lDirection) == lEdges){
                var lCxnOccs= lActObjOccs[i].Cxns(lEdges);
                
                for (var j in lCxnOccs){
                    var lCxnOcc = lCxnOccs[j];
                    if (lCxnTypes.length>0){                        
                        var lbOK=false
                        if (lCxnTypes.indexOf(lCxnOcc.CxnDef().TypeNum())>-1){
                            var lbOK=true
                        }
                    }else{
                        var lbOK= true
                    }
                    
                    if (lbOK){
                    
                        var ltCxnOccAndObjOcc = new tCxnOccAndObjOcc();
                        ltCxnOccAndObjOcc.Direction = lEdges;            
                        ltCxnOccAndObjOcc.CxnOcc = lCxnOcc;
                        if (lkeer==1)ltCxnOccAndObjOcc.ObjOcc = lCxnOcc.SourceObjOcc();
                        if (lkeer==2)ltCxnOccAndObjOcc.ObjOcc = lCxnOcc.TargetObjOcc();        
                        if (i>0)ltCxnOccAndObjOcc.bOnAllocDiag = true;
                        // voeg toe aan array
                        ltCxnOccAndObjOccs2Return.push (ltCxnOccAndObjOcc);                  
                    }
                }
            }
        }
    }
    return ltCxnOccAndObjOccs2Return;
}




function fIndexOf(pArisObjArray,pArisObject){
    var lResult = -1
    try{
        for (var i in pArisObjArray){
            if (pArisObjArray[i].IsEqual(pArisObject)){
                lResult = i
                break
            }
        }
    }catch(ex){}
    return lResult
}

function fcLng(val){
    
  var l2Return = 0;

  if(typeof(val)=="object")
    val = "" + val;
    
  if(typeof(val)=="string")
  {
    if(val.substr(0,2)=="&H" || val.substr(0,2)=="&h")
      l2Return = parseInt(val.substr(2),16);
    else if(val.substr(0,2)=="&O" || val.substr(0,2)=="&o")
      l2Return = parseInt(val.substr(2),8);
    else if(val.toLowerCase+"" == "true")
      l2Return = -1;
    else if(val.toLowerCase+"" == "false")
      l2Return = 0;
    else{
        var lsVal = ""
        for (i=0 ;i<val.length;i++){
            if (fIsNumeric(val.substr(i,1))
            || val.substr(i,1)=="."){
                lsVal+=val.substr(i,1);
            }else{
                break
            }
        }    
        l2Return = parseInt(lsVal * 1);
    }
  }else{
     l2Return = parseInt(val * 1);
  }
  if (isNaN(l2Return)){
     l2Return =0;
  }
  return l2Return;
}
function fcRight(psString, plLen){
    var lsString2Return = psString;
    
    if(psString==null)
        return "";

    if (typeof(psString) == "object"){
        lsString2Return = psString.toString() 
    }
    lsString2Return =lsString2Return+"";

    if (typeof(lsString2Return) != "string") return "";

    if (plLen <= 0) return "";

    if (plLen > lsString2Return.length) return lsString2Return;

    return lsString2Return.substring(lsString2Return.length-plLen);
}
function fIsNumeric(psText){
    var lbIsNumber2Return = true;
    var lsValidChars = "0123456789.";
    var lsChar;
    var lsText = fTrim(psText);
    
    if (lsText.length==0){
        var lbIsNumber2Return = false;
    }else{
        for (var i = 0; i < lsText.length && lbIsNumber2Return == true; i++) {
            lsChar = lsText.charAt(i); 
            if (lsValidChars.indexOf(lsChar) == -1){
                if (lsChar!="-" || i>0){
                    lbIsNumber2Return=false
                }
            }
        }  
    }
    return lbIsNumber2Return;   
}

function fTrim(psValue){
//  return ("" + psValue).trim()

  var lsValue2Return= "" + psValue;
  lsValue2Return=lsValue2Return.replace(/^ */,"");  
  lsValue2Return=lsValue2Return.replace(/ *$/,"");  
  return lsValue2Return;
}

function fcLeft(psString, plLen){
    var lsString2Return = psString;
    
    if(psString==null)
        return "";

    if (typeof(psString) == "object"){
        lsString2Return = psString.toString()
    }
    lsString2Return =lsString2Return+"";

    if (typeof(lsString2Return) != "string") return "";

    if (plLen <= 0) return "";

    if (plLen > lsString2Return.length) return lsString2Return;

    return lsString2Return.substr(0,plLen);
}

function fRGB(plred,plgreen,plblue){
    var lred=parseInt(fcLng(plred));
    var lgreen=parseInt(fcLng(plgreen));
    var lblue=parseInt(fcLng(plblue));    
    if (lred<0)lred=0
    if (lred>255) lred=255
    if (lgreen<0)lgreen=0
    if (lgreen>255) lgreen=255
    if (lblue<0)lblue=0
    if (lblue>255) lblue=255
    return fcLng("&H" + fcHex(lred,2) + fcHex(lgreen,2) + fcHex(lblue,2))
}

function fJavaByteArrayByLen(plLen){   
    return new java.lang.reflect.Array.newInstance(java.lang.Byte.TYPE, plLen);
}

function fRed(plRGB){
    return (plRGB  &  0xff0000) >> 0x10
}
function fGreen(plRGB){
    return (plRGB  &  0xff00) >> 0x8
}
function fBlue(plRGB){
    return (plRGB  &  0xff) 
}
function fcHex(plNum, plLen){
    var lLen = fValueIfMissing(plLen,0)
    var lsHex="";
    if(fIsNumeric(plNum)){
        lsHex = plNum.toString(16).toUpperCase() + "";   
        if (lLen>0)for(var i=lsHex.length; i<lLen; i++)lsHex= "0" + lsHex;             
    }
    return lsHex;
}
function fcBool(pVar){
	var lBool2Return = false;	
     if (typeof(pVar)=='boolean')return pVar
    
    var lVar = pVar;    
    if (fIsMissing(pVar)){
    }else if (pVar==""){    
    }else if (isNaN(pVar)){	
		lVar = "" + pVar;
		lVar = lVar.toLowerCase();
		if (lVar=="true" || lVar=="waar" || lVar=="y" || lVar == "yes" || lVar=="vrai" ){		
			lBool2Return = true;
		}
	}else{
		if (parseInt(lVar)!=0){
			lBool2Return = true;
		}
	}		  
	return lBool2Return;
}

// Checks, whether it's a GUID
function fcIsGUID(p_sGuid) {
    var sGuid = new String(p_sGuid);
    if (sGuid.length != 36) return false;
    if (sGuid.charAt(8)  != '-') return false;
    if (sGuid.charAt(13) != '-') return false;
    if (sGuid.charAt(18) != '-') return false;
    if (sGuid.charAt(23) != '-') return false;    
    return true;
}
function fcTimer(){
    return fGetTimer();
}
function fGetTimer(){
    var lTimer2Return = 0;
    var d = new Date();
    lTimer2Return = d.valueOf();
    return lTimer2Return;
}
function fIsArray(pArray){
    var lbOK2Return = false;
    if (pArray instanceof Array)lbOK2Return=true
    return lbOK2Return    
}
function fIsInArray(pArray,pElement, pbIgnoreCase){
    var lbIgnoreCase = fValueIfMissing(pbIgnoreCase, false)
    if (lbIgnoreCase){
        var lbIsInArray2Return = (("," + pArray.toString().toLowerCase() + ",").indexOf("," + pElement.toLowerCase() + ",")>-1);
    }else{
        var lbIsInArray2Return = (("," + pArray.toString() + ",").indexOf("," + pElement + ",")>-1);
    }
    return lbIsInArray2Return;    
}
function fUniqueArray(pArray){
    this.onlyUnique = function(value, index, self) { 
        return self.indexOf(value) === index;
    }
    return pArray.filter(this.onlyUnique)
}
function fNumOrdA(a, b){ return (a-b); }
function fNumOrdD(a, b){ return (b-a); } 
function fSortNumericArray(pArray){
    return pArray.sort(fNumOrdA)
}

function fIsMissing(pOptionalPar){
    var lbIsMissing2Return = false;
    if (pOptionalPar==undefined || pOptionalPar ==null){
        lbIsMissing2Return = true;        
    }
    return lbIsMissing2Return;
}
function fValueIfMissing(pOptional,pDefault){
    return (fIsMissing(pOptional) ? pDefault: pOptional)
}
function fsJavaNow(psFormat){
    var lsNow2Return = ""
    var ldDate=new Packages.java.util.Date();
    var lsFormat = fValueIfMissing(psFormat,"yyyy/MM/dd HH:mm:ss");    
    if (lsFormat=="")lsFormat = "yyyy/MM/dd HH:mm:ss";
    var lDateFormat = Packages.java.text.SimpleDateFormat(lsFormat,java.util.Locale.US);
    lsNow2Return=lDateFormat.format(ldDate);
    return lsNow2Return;
}
function fsNow(psOptionalFormat){
    return fsJavaNow(psOptionalFormat);
}
function fcNow(){
    var lDate2Return = new Date();
    var lsNow = fsJavaNow("yyyy-MM-dd-hh-mm-ss");
    lDate2Return = fcJavascriptDate(lsNow.split("-"));
    return lDate2Return;
}
function fsJavascriptNow(){
    return fcNow();
}
function fcJavascriptDate(yy,MM,dd,hh,mm,ss,nn){            // alternatief fcJavascriptDate([]);
    var lDate2Return = new Date();
    if (yy != undefined){
        if (fcTypeOf(yy)=="array"){
            var lyy = fcLng(fValueIfMissing(yy[0],0));
            var lMM = fcLng(fValueIfMissing(yy[1],0));
            var ldd = fcLng(fValueIfMissing(yy[2],0));
            var lhh = fcLng(fValueIfMissing(yy[3],0));
            var lmm = fcLng(fValueIfMissing(yy[4],0));
            var lss = fcLng(fValueIfMissing(yy[5],0));
            var lnn = fcLng(fValueIfMissing(yy[6],0));
        }else{
            var lyy = fcLng(fValueIfMissing(yy,0))
            var lMM = fcLng(fValueIfMissing(MM,0))
            var ldd = fcLng(fValueIfMissing(dd,0))
            var lhh = fcLng(fValueIfMissing(hh,0))
            var lmm = fcLng(fValueIfMissing(mm,0))
            var lss = fcLng(fValueIfMissing(ss,0))
            var lnn = fcLng(fValueIfMissing(nn,0))
        }
        lDate2Return.setFullYear(lyy,(lMM-1),ldd);
        lDate2Return.setHours(lhh,lmm,lss,lnn);
    }
    return lDate2Return;
}
function fcTypeOf(v){
    var lsType2Return = typeof(v);
    switch(lsType2Return){
    case "number": return lsType2Return;
    case "string": return lsType2Return;
    case "boolean": return lsType2Return;
    case "function": return lsType2Return;
    case "string": return lsType2Return;
    default:
        var lfConstructor = v.constructor;
        if( lfConstructor != null ){
        	switch( lfConstructor ){
        	case Array:
	        	lsType2Return = "array"; return lsType2Return;
        	case Date:
                lsType2Return = "date"; return lsType2Return;
            }
            var lsConstructor = lfConstructor.toString() + "";
            var lInstr_1 = lsConstructor.indexOf("function ");
            if (lInstr_1>=0){
                var lInstr_2 = lsConstructor.indexOf("(", lInstr_1);
                if (lInstr_2>=0){
                    lsType2Return = fTrim(lsConstructor.substring(lInstr_1+ ("function ").length,lInstr_2));
                }
            }
        }else{
            try{
                if (v.getClass().getSuperclass()=="class java.util.Date"){
                    lsType2Return = "date"; return lsType2Return;
                }
            }catch(ex){}
        }
	}
    return lsType2Return;
}
function fcDatesCompare(pDate1, pDate2){
    try{
        var ldt1 = pDate1.getTime()
        var ldt2 = pDate2.getTime()    
        return (ldt1===ldt2) ? 0 : (ldt1>ldt2) ? 1 : -1
        
    }catch(ex){
        var lex=ex
        var aa=0
        return NaN
    }
    
}


function fGetArisObjectArrayIndex(pArray, pObject){
    var lIndex = -1
    try{
        for (var i in pArray){
            var lItem = pArray[i]
            if (lItem.IsEqual(pObject)){
                lIndex = i
                break
            }
        }
    }catch(ex){var lex=ex
        var aa=0
    }
    return lIndex
}
function fcBeep(){
    Packages.java.awt.Toolkit.getDefaultToolkit().beep(); 
    
    

    var lClip = new javax.sound.sampled.AudioSystem.getClip();
    // getAudioInputStream() also accepts a File or InputStream
    var lFile =  new java.io.File("C:/Windows/Media/Windows Background.wav")
    var lAudioInputStream = new javax.sound.sampled.AudioSystem.getAudioInputStream(lFile)
    lClip.open(lAudioInputStream);
    lClip.loop(1);

}
function Files_ReadJavaByteArrayByFile(psFileLocation){
    var lJavaByteArray2Return = fJavaByteArrayByLen(0);
    
    if (Files_Exists(psFileLocation)){
        try{
           var lInputStream = new java.io.FileInputStream(psFileLocation);  
           var lLen = lInputStream.available();
            // Transfer bytes from in to out
            var lJavaByteArray2Return = fJavaByteArrayByLen(lLen)

            lLen = lInputStream.read(lJavaByteArray2Return);
            lInputStream.close();
            lbSuccess2Return = true;
        }catch(ex){
            var lex=ex;
            var aa=1;
        }
    }
    return lJavaByteArray2Return
}
function Files_ReadFile(psFileLocation){
    var lsText2Return="";
    
    var lsFileLocation = Files_GetCorrectFileName(psFileLocation)
    
    try{// create BufferedWriter to write text to file
        //var lWriter=new Packages.java.io.BufferedWriter(new Packages.java.io.FileWriter(pFileLocation));
        
        var lFileInputStream=new Packages.java.io.FileInputStream(lsFileLocation);
        var lInputStreamReader = new Packages.java.io.InputStreamReader(lFileInputStream, "UTF8");
        var lReader=new Packages.java.io.BufferedReader(lInputStreamReader);

        var lLine="";        
        try {
            while((lLine = lReader.readLine()) !=null){
                lsText2Return+=lLine;
            }
        }catch(ex){
        }
        
        try{lReader.close()}catch(ex){} 
        try{lInputStreamReader.close()}catch(ex){}        
        try{lFileInputStream.close()}catch(ex){}      
        
    }catch(object){        
    }
    return lsText2Return;
}

function Files_CreateZip(psSrcFileNames, psTargetFileName){      

    var lbOK=true
    
    try {
        // Create the ZIP file
        
        var lOutputStream = new Packages.java.io.ByteArrayOutputStream();
        var lZipOutputStream= new Packages.java.util.zip.ZipOutputStream(lOutputStream);
        for (var i in psSrcFileNames){
            var lsSrcFileName = psSrcFileNames[i]
            if (Files_Exists(lsSrcFileName)){               
                var lSrcBytes = Files_ReadJavaByteArrayByFile(lsSrcFileName)
                var lZipEntry = new Packages.java.util.zip.ZipEntry(Files_ExtractFileName(lsSrcFileName))
                lZipEntry.setSize(lSrcBytes.length);
                lZipOutputStream.putNextEntry(lZipEntry);
                lZipOutputStream.write(lSrcBytes, 0, lSrcBytes.length)
            }
        }
        
        lZipOutputStream.closeEntry();
        lZipOutputStream.close();
        
        var lOutputStream4File = Packages.java.io.FileOutputStream(psTargetFileName);
        
        var lBuffer = lOutputStream.toByteArray()    
        lOutputStream4File.write(lBuffer, 0, lBuffer.length);
        
        lOutputStream.close();
        lOutputStream4File.close()
        
    } catch (ex) {var lex=ex;
        lbOK = false
        var aa=1;
    }
    return lbOK
}


// prototype array.indexOf( element,optional from)   
if (!Array.prototype.indexOf){  
    Array.prototype.indexOf = function(pElement/*,pfrom*/)  {
        var lIndexOf2Return = -1;    
        var lLength = this.length;    
        var lFrom = Number(arguments[1]) || 0;
        lFrom = (lFrom < 0)? Math.ceil(lFrom): Math.floor(lFrom);
        if (lFrom < 0){
            lFrom += len;
        }
        for (; lFrom < lLength; lFrom++){
            if (lFrom in this && this[lFrom] === pElement){
                lIndexOf2Return = lForm
                break;
            }
        }
        return lIndexOf2Return;
    }
}

if (!String.prototype.trim){  
    String.prototype.trim = function(){
       var ls=this.replace(/^ */,"");  
       ls=ls.replace(/ *$/,"");
       return ls
    }
}

var mlLastTimer= 0
var msFile2Listen2=""
function Log_Text(psText, pbIsError){
    var lTimer= new Date().valueOf();
    if (pbIsError){
    }else{
        if (lTimer - mlLastTimer< glTimerIntervalInLog) return
    }
    mlLastTimer = lTimer
    var lsLogFile= fValueIfMissing(gsDefaultLogFile,"")
    if (lsLogFile==""){
        gsDefaultLogFile = GEAR_GetValueInConfig("LogFile")        
        lsLogFile= fValueIfMissing(gsDefaultLogFile,"")        
        msFile2Listen2 = GEAR_GetValueInConfig("ListenToFile")        
    }
    
    if (lsLogFile=="")return
    try{
        Files_AppendToFile(lsLogFile,  fsNow() + " " + psText + gsLFCR);
    }catch(ex){
        var lex=ex;
        var aa=0        
    }
    if(msFile2Listen2>""){
        if (Files_ReadFile(msFile2Listen2).toUpperCase()=="BREAK"){
            Files_WriteToFile(msFile2Listen2, "Done")
            gbBreakScriptRun = true
        }
    }
    
}
function Utils_SplitString(psString, psSplitChar, pbOptionalNoEmptyStrings){
    var lsStrings2Return = new Array();
    

    var lbNoEmptyStrings = fValueIfMissing(fcBool(pbOptionalNoEmptyStrings), false)

    var lsSplitChar  = fValueIfMissing(psSplitChar + "", ",")
    if (lsSplitChar == ""){    
        lsStrings2Return.push(psString);
    }else{
        var lsAdhocs = psString.split(psSplitChar);
        for (var i in lsAdhocs){
            var lsAdhoc = fTrim(lsAdhocs[i]);
            if (lbNoEmptyStrings){            
                if (lsAdhoc!=""){
                    lsStrings2Return.push(lsAdhoc)
                }
            }else{
                lsStrings2Return.push(lsAdhoc)
            }
        }
    }return lsStrings2Return
}
function Files_WriteToFile(pFileLocation, psText){
    try{// create BufferedWriter to write text to file
        //var lWriter=new Packages.java.io.BufferedWriter(new Packages.java.io.FileWriter(pFileLocation));
        var lsFileName = Files_CheckedFilename(Files_ExtractFileName(pFileLocation));
        var lsPath2File= Files_ExtractPath(pFileLocation);
        if (Files_Exists(lsPath2File)){ 
            var lFileLocation = lsPath2File + "/" + lsFileName;
        
            var lWriter=new Packages.java.io.BufferedWriter(new Packages.java.io.OutputStreamWriter(new Packages.java.io.FileOutputStream(lFileLocation), "UTF8"));
            lWriter.write(psText);
            lWriter.flush();
            lWriter.close();            
            return true;
        }
    }catch(object){
        return false;
    }
}

function Files_AppendToFile(pFileLocation, psText){
    try{// create BufferedWriter to write text to file
        //var lWriter=new Packages.java.io.BufferedWriter(new Packages.java.io.FileWriter(pFileLocation));
        var lsFileName = Files_CheckedFilename(Files_ExtractFileName(pFileLocation));
        var lsPath2File= Files_ExtractPath(pFileLocation);

        if (Files_Exists(lsPath2File)){ 
            var lFileLocation = lsPath2File + "/" + lsFileName;
            var lWriter=new Packages.java.io.RandomAccessFile(lFileLocation, "rw");
            lWriter.seek(lWriter.length());
            lWriter.writeBytes(psText);
            lWriter.close();            
            return true;
        }
    }catch(ex){
        var lex=ex
        var aa=0
        return false;
    }
}
function Files_DeleteFile(psFileLocation){
    var lbSuccess2Return =false;
    var lsFileLocation = Files_GetCorrectFileName(psFileLocation);
    if (Files_Exists(lsFileLocation)){
        try{
            var lbSuccess2Return = (new Packages.java.io.File(lsFileLocation))["delete"]();
        }catch(ex){
            var lex=ex;
            var aa=1;
            
        }
    }
    return lbSuccess2Return;
}
function Files_CopyFile(pSrcFileName,pDstFileName){
    // Copies src file to dst file.
    // If the dst file does not exist, it is created  
    var lbSuccess2Return = true;
    
    if (Files_Exists(pSrcFileName)){
        var lSrcFile = new java.io.File(pSrcFileName);
        if (!lSrcFile.canRead()){
            lbSuccess2Return=false
        }else{
            var lDestFile = new java.io.File(pDstFileName);
            if (Files_Exists(pDstFileName)){
                if (!lDestFile.canWrite()){
                    lbSuccess2Return=false
                }
            }
        }
        if (lbSuccess2Return){                       
            
            lInputStream = Packages.java.io.FileInputStream(pSrcFileName);
            try{
               lOutputStream = Packages.java.io.FileOutputStream(pDstFileName);
        
                // Transfer bytes from in to out
                var lBuffer = fJavaByteArrayByLen(1024)
                var lLen =0;
                while ((lLen = lInputStream.read(lBuffer)) > 0) {
                    lOutputStream.write(lBuffer, 0, lLen);
                }
                lOutputStream.close();
            }catch(ex){
               lbSuccess2Return=false
           }                    
            lInputStream.close();
            
        }
    }
    return lbSuccess2Return
}
function Files_CreateFolder (psFolderName){
    var lbSuccess2Return = false;
    var lsFolderName = Files_GetCorrectFileName(psFolderName);
     // Create a directory; all ancestor directories must exist
    lbSuccess2Return = (new Packages.java.io.File(lsFolderName)).mkdir();
    return lbSuccess2Return;   
}
function Files_CreatePath (psPathFolderName){
    var lbSuccess2Return = false;
    var lsPathFolderName = Files_GetCorrectFileName(psPathFolderName);
    lbSuccess2Return = (new Packages.java.io.File(lsPathFolderName)).mkdirs();
    return lbSuccess2Return;   
}
function Files_CheckedFilename(psFileName, psReplace, psPrefix){
	var lsFileName2Return = "" + psFileName;
    
    var lsPrefix = "" + fTrim(psPrefix);
    if (!lsPrefix.equals(""))lsPrefix=Files_CheckedFilename(lsPrefix, psReplace,"");
    if (lsPrefix.equals(""))lsPrefix = "Report";
    var lsForbidden = "*/:?<>|" + String.fromCharCode(34);
    for (var i = 0; i<lsForbidden.length;i++){
      lsFileName2Return = lsFileName2Return.replace(lsForbidden.substr(i,1),psReplace);
    }
    lsFileName2Return=fTrim(lsFileName2Return);
    if (lsFileName2Return.equals("")){
        lsFileName2Return = lsPrefix + " " + fsNow("yyyy/MM/dd HH:mm:ss");
    }
    var lFirstChar=lsFileName2Return.substr(0,1);
    var lsNumbers="0123456789";
    if (lsNumbers.indexOf(lFirstChar)>-1){
        // filename begint met een nummer, magniet
        lsFileName2Return = lsPrefix + " " + lsFileName2Return;
    }
    return lsFileName2Return;
}

function Files_ExtractExtention (psFileName){
    var lsExtention2Return = "";
    var lsFileName = psFileName + "";
    var lInstr = lsFileName.lastIndexOf(".");
    if (lInstr>=0)lsExtention2Return = lsFileName.substr(lInstr+1);
    return lsExtention2Return;   
}
function Files_ExtractFileName (psFileName){
    var lsFileName2Return = "";
    var lsFileName = psFileName + "";
    lsFileName = Files_GetCorrectFileName(lsFileName);
    var lInstr = lsFileName.lastIndexOf("/");
    if (lInstr>=0) lsFileName=lsFileName.substr(lInstr + 1)
    lsFileName2Return = lsFileName;
    return lsFileName2Return;
}
function Files_ExtractName(psFileName){
    var lsName2Return = "";
    var lsFileName = psFileName + "";
    var lInstr = lsFileName.lastIndexOf(".");
    if (lInstr>=0)lsFileName = lsFileName.substr(0,lInstr);           
    lsName2Return = Files_ExtractFileName(lsFileName);
    return lsName2Return;
}

function Files_ExtractPath(psFileName){
    var lsPath2Return = "";
    var lsFileName = psFileName + "";
    lsFileName = Files_GetCorrectFileName(lsFileName);
    var lInstr = lsFileName.lastIndexOf("/");
    if (lInstr>=0) lsPath2Return=lsFileName.substr(0,lInstr)

    return lsPath2Return;
}
function Files_Exists(psFileLocation){
    var lBool2Return = false;
    var lsFileLocation = Files_GetCorrectFileName(psFileLocation);

    var lsFileNameComponents= lsFileLocation.split("/");
    
    if (fcIsGUID(lsFileNameComponents[0])){
        // is sprake van in Aris 9
         var lScriptAdminComponent = Context.getComponent("ScriptAdmin")
         if (fIsMissing(lScriptAdminComponent)){
             // user - not allowed to see scripts
         }else{
             var lsCatID = lsFileNameComponents[0];
             var lScriptInfos= lScriptAdminComponent.getScriptInfos (1, lsCatID, Context.getSelectedLanguage())             
             for (var i in lScriptInfos){
                 if (lScriptInfos[i].getName()+"" == lsFileNameComponents[1]){
                     lBool2Return=true;
                     break;
                 }
             }
        }
    }else{
        lBool2Return = new Packages.java.io.File(lsFileLocation).exists();
    }
    return lBool2Return;
    
}
function Files_SortFileByLastModDate(psFileLocations){  // sort from young to old
    var ltmSeconds = new java.util.TreeMap()
    for (var i in psFileLocations){        
        var lsFileLocation = psFileLocations[i]
        var lLong = new java.io.File(lsFileLocation).lastModified()
        ltmSeconds.put(new java.lang.Long(lLong),i)
    }
    var lSort = ltmSeconds.values().toArray()
    lSort.reverse()
    var lsFileLocations = []
    for (var i in psFileLocations){       
        lsFileLocations.push(psFileLocations[lSort[i]])
    }
    return lsFileLocations
}
function Files_GetFileNamesInDir(psDirectoryName, psOptionalPrefix, psOptionalExtention){
    //psOptionalExtention kan uit meer dan een extention bestaan, in dat geval gebruik maken van scheidingsteken "|"

    var lsFileNames2Return = new Array();
    
    
    var lsOptionalPrefix = fValueIfMissing(psOptionalPrefix, "") + ""
    lsOptionalPrefix=lsOptionalPrefix.replace("*","");
    lsOptionalPrefix=lsOptionalPrefix.toUpperCase();
   
    var lsOptionalExtention = fValueIfMissing(psOptionalExtention, "") + ""
    lsOptionalExtention=lsOptionalExtention.replace("*","");
    lsOptionalExtention=lsOptionalExtention.replace(".","");
    lsOptionalExtention=lsOptionalExtention.toUpperCase();
    
    var lsOptionalExtentions = Utils_SplitString(lsOptionalExtention, "|", false);
    
    var lsFileNames = new Array();
    
    var lsDirectoryName = Files_GetCorrectFileName(psDirectoryName);
    var lDir = new Packages.java.io.File(lsDirectoryName);
    if (lDir.isDirectory()){;
        var lFiles=lDir.listFiles();
        for (var i=0 in lFiles){
            if(!(lFiles[i].isDirectory())){
                lsFileNames.push(lFiles[i].getName());
            }
        }
    }
    for (var i in lsFileNames){
        var lbOk=true;
        var lsFileName = lsFileNames[i]
        var lInstr=lsFileName.lastIndexOf(".");
        if (lInstr==0)lInstr=lsFileName.length;
        var lsExtention = lsFileName.substr(lInstr + 1);
        
        
        if ((lsOptionalPrefix+"")!=""){
            if ((fcLeft(lsFileName, lsOptionalPrefix.length)+"").toUpperCase() != lsOptionalPrefix){
                lbOk=false;  
            }
        }
        if (lsOptionalExtentions.length>0){
            var lbOk4Extention = false;
            for (var j=0;j<lsOptionalExtentions.length;j++){
                var lsOptionalExtention = lsOptionalExtentions[j];
                if ((lsOptionalExtention+"")==""){
                    lbOk4Extention=true;
                }else if ((lsExtention+"").toUpperCase() == lsOptionalExtention){
                    lbOk4Extention=true;
                }
                if (lbOk4Extention==true)break;
            }
        }else{
            var lbOk4Extention = true;
        }                   
        if (lbOk4Extention==false){
            lbOk = false;
        }   
        if (lbOk){
            lsFileNames2Return.push(lsFileName);
        }
    }
    lsFileNames2Return = lsFileNames2Return.sort();
    return lsFileNames2Return;
}
function Files_RemoveForbiddenCharsInFileName(psFileName){
    var lsFileName2Return = psFileName;
    lsFileName2Return = (lsFileName2Return + "").replace(/\x5c/g,"_")+"";           //  verander alle \ in _
    lsFileName2Return = (lsFileName2Return + "").replace(/\x2f/g,"_")+"";          // verander alle / in _
    lsFileName2Return = (lsFileName2Return + "").replace(/[:*?<>|]/g , "_")+"";    // verander alle :*?<>| in _
    lsFileName2Return = (lsFileName2Return + "").replace(/\x22/g,"_")+"";          // verander alle " in _    
    return lsFileName2Return;
}

function Files_GetCorrectFileName(psFileName){
    var lsFileName2Return = psFileName;
    lsFileName2Return = (lsFileName2Return + "").replace(/\\/g,"/")+"";              
    return lsFileName2Return;
}


//-----------------------------------------------------------------------------------
//IDS-Scheer NL
//version 1.0: amkr
//-----------------------------------------------------------------------------------

function KeySet(psID)
{
  this.TreeMap = new Packages.java.util.TreeMap();
  this.KeyIndex = new Array();  
  
  this.toString = function (psKeyText){
      var lsKeyText = "";  
      if (psKeyText==null) return "";
      if (psKeyText.size==0) return "";      
      if (this.CaseSensitive){
          return psKeyText.toString();
      }else{
          return psKeyText.toString().toLowerCase(); 
      }
  } 
  
  this.ID = fValueIfMissing(psID, "");    // te gebruiken om een bepaalde keyset te traceren
                                          // (!, let op, dit gaat dus niet over een individueel item)   
  this.mIsNumeric=false;
  this.isNumeric=function(pbOptionalIsNumeric){
      if (fIsMissing(pbOptionalIsNumeric)){
          return this.mIsNumeric;
      }else{
          this.mIsNumeric = pbOptionalIsNumeric;
      }
  }

  this.AddItem = function(psKeyText){
    //returns index of key (old one, when psKeyText already exists, new one when psKeyText is added.
    if (this.mIsNumeric){
        var lsKeyText = this.toString(100000000000 + fcLng(psKeyText));
    }else{
    var lsKeyText = this.toString(psKeyText);
    }
    if (lsKeyText == "") return -1;
    if (!this.TreeMap.containsKey(lsKeyText)){  
        var lIndex = this.TreeMap.size();
        this.TreeMap.put(lsKeyText,lIndex);
        this.KeyIndex[lIndex]=psKeyText;
    }else{
        lIndex = this.TreeMap.get(lsKeyText);  
    } 
    return parseInt(lIndex);
  }
  this.iterator = function(){
    return this.TreeMap.keySet().iterator();
  }

  this.Count = function(){
       return this.TreeMap.size();
  }
  this.GetKeyText = function(plIndex){
    if ((plIndex>=0) && (plIndex<this.KeyIndex.length)){
        if (this.mIsNumeric){
            return fcLng(this.KeyIndex[plIndex]) - 100000000000;
        }else{
        return this.KeyIndex[plIndex];
        }
    }else{
        return "";
    }  
  }
  
  this.CaseSensitive = false;

  this.GetIndex = function(psKeyText){
    var lsKeyText = this.toString(psKeyText);
    if (lsKeyText == "") return -1;

    if (this.TreeMap.containsKey(lsKeyText)){ 
        return parseInt(this.TreeMap.get(lsKeyText));
    }else{
        return -1;   
    }
  }

  this.Remove = function(psKeyText){
    var lsKeyText = this.toString(psKeyText);
    if (lsKeyText== "") return false;
  
    var lIndex=this.GetIndex(psKeyText);  
    if (lIndex>=0){  
        this.KeyIndex[lIndex]=-1;
        this.TreeMap.remove(lsKeyText);  
        return true;
      }else{
        return false;
      }
    }
  this.Sort =function(){
    var lreturn = new Array();  
    if (this.Count()>0){
        lreturn = this.TreeMap.values().toArray();
    }  
    return lreturn;
  }
}


/**************************************************************************************/
/* XML Helper functions
/**************************************************************************************/

function AAS_GetXmlDocument(psXMLContent){
    var lDocument=null;   

    if(psXMLContent!=""){
        try{
            var lsXML = new Packages.java.lang.String(psXMLContent);
            var lXMLStream = new Packages.java.io.ByteArrayInputStream(lsXML.getBytes());
            var lBuilder = new Packages.org.jdom.input.SAXBuilder();
             var lDocument = lBuilder.build(lXMLStream);
//             var lDocument = lBuilder.build(lsXML);
        }catch(ex){
            var lex=ex;            
            try{
                var lsXML = new Packages.java.lang.String(fDecode4Utf8(psXMLContent));
                var lBytes = lsXML.getBytes("US-ASCII");
                var lDocument = Context.getXMLParser(lBytes);
//                var lsXML = new Packages.java.lang.String(lBytes);
//              var lDocument = lBuilder.build(lsXML);
            }catch(ex){
                var lex=ex;            
                AAS_ErrorDialog ("Error while opening XML document."
                           + "\nError is: " + ex, "routine: getXmlDocument");
            }

            
            

        }
    }
    return lDocument;
}
function getXmlDocument(psXMLContent){
    var lDocument=null;
    if(psXMLContent!=""){
        try{
            var lsXML = new Packages.java.lang.String(psXMLContent);
            var lXMLStream = new Packages.java.io.ByteArrayInputStream(lsXML.getBytes());
            var lBuilder = new Packages.org.jdom.input.SAXBuilder();
            var lDocument = lBuilder.build(lXMLStream);
        }catch(ex){
            var lex=ex;            
            ATSC_ErrorDialog ("Error while opening XML document."
                           + "\nError is: " + ex, "routine: getXmlDocument");

        }
    }
    return lDocument;
}

function XML_GetStringByDocument(pDocument){
    return XML_getXMLOutputSting(pDocument);
}
function getXMLOutputSting(pDocument){
    return XML_getXMLOutputSting(pDocument);
}
function XML_getXMLOutputSting(pDocument){
     var lsString2Return = "";
    if(pDocument==null){
        return "";   
    }else{
        var lOutputter=new Packages.org.jdom.output.XMLOutputter(Packages.org.jdom.output.Format.getPrettyFormat());
         lsString2Return = lOutputter.outputString(pDocument);           
    }
    return lsString2Return;
}

function XML_GetDocumentByString(psContent){
    var lDocument2Return=null;
    
    if(psContent!=""){
        try{
            var lsContent = new Packages.java.lang.String(psContent);
            var lXMLStream = new Packages.java.io.ByteArrayInputStream(lsContent.getBytes());
            var lBuilder = new Packages.org.jdom.input.SAXBuilder();
            lDocument2Return = lBuilder.build(lXMLStream);
        }catch(ex){
            var lex=ex;
            AAS_ErrorDialog("Error in Attributes value(s)\nXML syntax invalid! Error message: " + ex, "ATSC_Settings");
        }
    }
    return lDocument2Return;
}

function XML_GetContentByDocument(pDocument){
    if(pDocument==null){
        return "";   
    }else{
        var lOutputter=new Packages.org.jdom.output.XMLOutputter(Packages.org.jdom.output.Format.getPrettyFormat());
        return lOutputter.outputString(pDocument);           
    }
}

function XML_GetRootElement(pDocument){
    return pDocument.getRootElement()  
}


function XML_Element_GetName(pElement){
    var lsName =""
    if (pElement)lsName = pElement.getName()
    return lsName
}

function XML_Element_GetAllChildren(pElement, psKey){
    var lChilds2Return = new Array()
    
    var lChilds = pElement.getChildren()
    for (var i=0;i<lChilds.size();i++){
        var lChild = lChilds.get(i)
        if (lChild.getName()+"" == psKey + ""){
            lChilds2Return.push(lChild);
        }
        lChilds2Return=lChilds2Return.concat(XML_Element_GetAllChildren(lChild, psKey))
    }
    return lChilds2Return    
}

function XML_Element_GetAttrValue(pElement, psKeySet, psDefault, pbKeyCaseSensitive){
    var ls2Return = "";

    lsDefault = fValueIfMissing(psDefault, "") + ""
    lbKeyCaseSensitive = fValueIfMissing(pbKeyCaseSensitive, true)
    if (lbKeyCaseSensitive!=false)lbKeyCaseSensitive=true
        
    var lsKeys= new Array()
    if (fIsArray(psKeySet)){
        var lsKeys = psKeySet
    }else{
        lsKeys[0]=psKeySet
    }
    
    var lbFound = false;
    
    if (!lbKeyCaseSensitive){
        for (var i in lsKeys){
            var lsKey = (lsKeys[i]+"").toUpperCase()
            try{
                var lAttributes = pElement.getAttributes()
                for (var j=0; j<lAttributes.size();j++){
                    if (lAttributes.get(j).name.toUpperCase()==lsKey){
                        ls2Return = lAttributes.get(j).value + ""
                        lbFound = true
                        break
                    }
                }
            }
            catch(ex){
                var lex=ex;
                var a=1;
                ls2Return = lsDefault;
            }
            if (lbFound)break
        }
    }else{
        for (var i in lsKeys){        
            try{
                ls2Return = pElement.getAttribute(lsKeys[i]).getValue()+""
                lbFound=true
            }
            catch(ex){
                var lex=ex;
                var a=1;
            }
            if (lbFound)break
        }
    }
    if (!lbFound)ls2Return = lsDefault;
    
    return (ls2Return + "");
}
function XML_Element_GetText(pElement, psOptionalKey, psOptionalDefaultValue){
    var ls2Return = "";
    var lsKey = fValueIfMissing(psOptionalKey,"");
    var lsDefault = fValueIfMissing(psOptionalDefaultValue,"")
    
    try{
        if (lsKey==""){
           ls2Return = pElement.getTextNormalize()+""
        }else{
           ls2Return = pElement.getChild(lsKey).getTextNormalize()+""
        }
    }
    catch(ex){
        var lex=ex;
        var a=1;
    }
    ls2Return= fTrim(ls2Return+"");
    if (ls2Return=="")ls2Return=lsDefault
    return ls2Return ;
}

function XML_EvaluateXpath(pXMLFile, psXpathExpr){
    var lsKeyValue2Return=""
    
    var lXpath = Packages.javax.xml.xpath.XPathFactory.newInstance().newXPath();
    var lBuilder = Packages.javax.xml.parsers.DocumentBuilderFactory.newInstance().newDocumentBuilder();
    
    var lDocument = lBuilder.parse(pXMLFile);    
    var lsKeyValue2Return = lXpath.evaluate(psXpathExpr, lDocument, Packages.javax.xml.xpath.XPathConstants.STRING);
    
    return lsKeyValue2Return;
}

function fMyXMLTagEncrypt(psTag){
    var lsNewTag2Return = psTag + "";
    var lsTag = psTag+"";
    for (var i=lsTag.length-1;i>=0;i--){
        var lAsc = lsTag.charCodeAt(i);
        if (lAsc==32){
            lsNewTag2Return = lsNewTag2Return.substr(0,i) +  "_"  + lsNewTag2Return.substr(i+1)
        }else         
        if ((lAsc>=48 && lAsc<58)
        || (lAsc>=65 && lAsc<=90)
        || (lAsc>=97 && lAsc<=122)                    
        || (lAsc==95)){
            //oke
        }else{
            lsNewTag2Return = lsNewTag2Return.substr(0,i)+ ".h" + fcHex(lAsc,2) + lsNewTag2Return.substr(i+1)
        }
    }
    return lsNewTag2Return;    
}

function EMAIL_DoSendMessage(psHost, plPort, psFrom, psTOs, psCCs, psSubject, psContent){
    
 try{

        var lProperties = java.util.Properties()
        lProperties.put("mail.smtp.host", psHost);         // for yahoo use smtp.mail.yahoo.com and for gmail use smtp.gmail.com
        lProperties.put("mail.smtp.port", plPort);                // The SMTP server port to connect to, if the connect() method doesn't 
                                                                // explicitly specify one. Defaults to 25.  
        var lMailSession = javax.mail.Session.getInstance(lProperties)   // The Session class represents a mail session and is not subclassed
        lMimeMessage = new javax.mail.internet.MimeMessage( lMailSession );

        lMimeMessage.setFrom( new javax.mail.internet.InternetAddress( psFrom));
        for each (var lsTO in psTOs){
            lMimeMessage.addRecipient(javax.mail.Message.RecipientType.TO, new javax.mail.internet.InternetAddress(lsTO))
        }
        for each (var lsCC in psCCs){
            lMimeMessage.addRecipient(javax.mail.Message.RecipientType.CC,  new javax.mail.internet.InternetAddress(lsCC))
        }
        lMimeMessage.setSentDate( new java.util.Date());
        lMimeMessage.setSubject( psSubject);
        lMimeMessage.setText( psContent);
        
        javax.mail.Transport.send( lMimeMessage );    
        var lBool = true
    }catch(ex){
        var lex=ex
        var aa=0
        var lBool = false
    }
    return lBool
}
function EMAIL_SendMessage(psFrom, psTOs, psCCs, psSubj, psContent){
    var lsHost = GEAR_GetValueInConfig("SMTPMailServer", "Host")
    var lPort  = GEAR_GetValueInConfig("SMTPMailServer", "Port")
   
    var lsContent = Context.getScriptInfo(Constants.SCRIPT_TITLE) 
                  + "\n\r" + "========================================================"
                  + "\n\r" + psContent 
                  + "\n\r" + "========================================================"
                  + "\n\r" + "Name server: " + ArisData.getActiveDatabase().ServerName()
                  + "\n\r" + "Database name: " + ArisData.getActiveDatabase().Name(-1)
                  + "\n\r" + "Filter name: " + ArisData.getActiveDatabase().ActiveFilter().Name(-1)
                  + "\n\r" + "User name: " + ArisData.getActiveDatabase().ActiveUser().Name(-1)
                  + "\n\r" + "========================================================"                  
                  + "\n\r" + "This is an automatic send message."
    return EMAIL_DoSendMessage(lsHost, lPort, psFrom, psTOs, psCCs, psSubj, lsContent)
}
