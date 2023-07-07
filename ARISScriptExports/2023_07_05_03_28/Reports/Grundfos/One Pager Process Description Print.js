var gPdfSettings = null;
var gDocument = null;
var gLoc = Context.getSelectedLanguage();
var gWriter = null;
var gOutStream = null
var gDateFormatter = new java.text.SimpleDateFormat("MMM dd, yyyy", new java.util.Locale("en"));
var gObjectMap= new java.util.HashMap();
var gPageSize=Context.getProperty("pageSize");
var gExludeAttrNr=ArisData.ActiveFilter().UserDefinedAttributeTypeNum("6fcf1e11-d899-11e7-27f1-00155dff4f3e");
var gFuncAssignedModelTypeMap= new java.util.HashMap();

var gBGColor=new function(){
    this.cornerCell=new java.awt.Color(14803422);
    this.processHeaderCell=new java.awt.Color(1132923);
    this.rowHeaderCell=new java.awt.Color(14803422);
    this.activityHeaderCell=new java.awt.Color(9869980);
}
function PDFSettings(){
    this.pageSize=com.lowagie.text.PageSize.A3.rotate();    
    this.leftMargin=30;
    this.topMargin=50;
    this.rightMargin=30;
    this.bottomMargin=75;   
}

function main(){
    var selModels=ArisData.getSelectedModels();
    
    createDocument();
    for(var n  in selModels){ 
        var selModel=selModels[n];
        
        var funcObjDefs=sortFunctions(selModel, Constants.SORT_X, Constants.SORT_Y);
        if(funcObjDefs.length >0){            
            for each(var funcObjDef in funcObjDefs){
               var assignedModels=funcObjDef.AssignedModels([Constants.MT_EEPC_TAB]);
               if(assignedModels.length>0){
                   lbTableDisplay=true;
                   gFuncAssignedModelTypeMap.put(funcObjDef.GUID(),Constants.MT_EEPC_TAB)
               }else{
                   gFuncAssignedModelTypeMap.put(funcObjDef.GUID(),Constants.MT_EEPC)
               }
            }
            if(n>0) gDocument.newPage();
            var nrOfProcesses=3;
            if(gPageSize=="A0"){
                nrOfProcesses=6;
            }
            if(funcObjDefs.length <= nrOfProcesses){
                createPage(selModel,funcObjDefs.length, funcObjDefs, 0);
            }else{                
                /*if(Math.ceil(funcObjDefs.length/4)<Math.ceil(funcObjDefs.length/3)){
                    nrOfProcesses=4;
                }*/
                i=0;
                do{                    
                    if(i+nrOfProcesses>funcObjDefs.length){
                        nrOfProcesses=funcObjDefs.length-i;
                    }
                    if(i>0) gDocument.newPage();
                    createPage(selModel,nrOfProcesses, funcObjDefs, i);
                    i+=nrOfProcesses;
                }while(i<funcObjDefs.length && nrOfProcesses>0)
            }
            createObjectOverviewPage(selModel);
        }
                
    }
    closeDocument();
    gObjectMap.clear();
}
function createDocument(){
    try{
       
        gPdfSettings = new PDFSettings();
        if(gPageSize=="A0"){
            gPdfSettings.pageSize=com.lowagie.text.PageSize.A0.rotate();
            gPdfSettings.bottomMargin=2.25* gPdfSettings.bottomMargin;
        }
        gOutStream = new java.io.ByteArrayOutputStream();
        //gOutStream = new java.io.FileOutputStream(new java.io.File(lFileInternalOutputPath,Context.getSelectedFile()))
        
        gDocument = new com.lowagie.text.Document(gPdfSettings.pageSize, gPdfSettings.leftMargin, gPdfSettings.rightMargin, gPdfSettings.topMargin, gPdfSettings.bottomMargin);						    
        gWriter=com.lowagie.text.pdf.PdfWriter.getInstance(gDocument, gOutStream);
        gWriter.setLinearPageMode();
        
        gDocument.addAuthor(ArisData.getActiveUser().Name(gLoc)); 
        gDocument.addSubject(Context.getScriptInfo(Constants.SCRIPT_NAME));
        
        gDocument.open();
        
        gPageEventHelper = new JavaAdapter(com.lowagie.text.pdf.PdfPageEventHelper, {
            onEndPage: function (writer, document) {                
                gCurrentPageNumber = writer.getCurrentPageNumber();
            }
        });
        
        gWriter.setPageEvent(gPageEventHelper);
    }catch(ex){
        
        Dialogs.MsgBox("Unexpected error: "+ex);
        
    }
}

function closeDocument(){
    gDocument.close();
       
    var lUnusedOutputObject = Context.createOutputObject();
    var lFileInternalOutputPath = lUnusedOutputObject.getCurrentPath();//returns a java.io.File object
        
    lUnusedOutputObject.OutputLn("unused", "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
    lUnusedOutputObject.WriteReport();
    
    var reader = new com.lowagie.text.pdf.PdfReader(gOutStream.toByteArray());
    var n = reader.getNumberOfPages();
    
    var stamp = new com.lowagie.text.pdf.PdfStamper(reader, new java.io.FileOutputStream(new java.io.File(lFileInternalOutputPath, Context.getSelectedFile())));
    stamp.close();
}


function createPage(pSelModel,piNrOfProcesses,pFuncObjDefs, pStartIndex){
    var lContentPadding=5;       
    var firstOuterColumnWidth=30;
    var firstOuterRowHeight=30;
    var outerTableWidth=gPdfSettings.pageSize.getWidth()-gPdfSettings.leftMargin-gPdfSettings.rightMargin;
    var outerTableHeight=gPdfSettings.pageSize.getHeight()-gPdfSettings.topMargin-gPdfSettings.bottomMargin;
    var intArray = java.lang.reflect.Array.newInstance(java.lang.Integer.TYPE, piNrOfProcesses+1);
    
    var actRespHeight=(outerTableHeight-firstOuterRowHeight)*3/8;
    var actDescHeight=(outerTableHeight-firstOuterRowHeight)*5/8;
    
    var lProcessWidth=(outerTableWidth-firstOuterColumnWidth)/piNrOfProcesses;
    intArray[0]=1000*firstOuterColumnWidth/outerTableWidth;
    for(var i=1;i<=piNrOfProcesses;i++){
        intArray[i]=1000*((outerTableWidth-firstOuterColumnWidth)/piNrOfProcesses)/outerTableWidth;
    }
    
    var table = new com.lowagie.text.pdf.PdfPTable(piNrOfProcesses+1);
    table.setWidths(intArray);
    table.setTotalWidth(outerTableWidth);
    table.setWidthPercentage(100);
    
    var tableHeaderFont = getTextFont(true,16);
    tableHeaderFont.setColor(new java.awt.Color(16777215));
    var rowHeaderFont = getTextFont(true,12);
    rowHeaderFont.setColor(new java.awt.Color(0));
    
    var tcell = new com.lowagie.text.pdf.PdfPCell(new com.lowagie.text.Phrase("", tableHeaderFont));
    tcell.setFixedHeight(firstOuterRowHeight);    
    tcell.setBackgroundColor(gBGColor.cornerCell);
    tcell.setBorderColor(gBGColor.processHeaderCell);
    table.addCell(tcell);
    
    
    for(var i=0;i<piNrOfProcesses;i++){
        
        var tcell = new com.lowagie.text.pdf.PdfPCell(new com.lowagie.text.Phrase(pFuncObjDefs[i+pStartIndex].Name(gLoc), tableHeaderFont));
        tcell.setHorizontalAlignment(com.lowagie.text.Element.ALIGN_CENTER);
        tcell.setVerticalAlignment(com.lowagie.text.Element.ALIGN_MIDDLE);
        tcell.setBackgroundColor(gBGColor.processHeaderCell);
        tcell.setBorderColor(gBGColor.processHeaderCell);
        table.addCell(tcell);        
    }
    
    //print next row, activity responsibilities
    var tcell = new com.lowagie.text.pdf.PdfPCell(new com.lowagie.text.Phrase("Activity Responsibilities", rowHeaderFont));
    tcell.setFixedHeight(actRespHeight);
    tcell.setRotation(90);
    tcell.setHorizontalAlignment(com.lowagie.text.Element.ALIGN_CENTER);
    tcell.setVerticalAlignment(com.lowagie.text.Element.ALIGN_MIDDLE);
    tcell.setBackgroundColor(gBGColor.rowHeaderCell);
    tcell.setBorderColor(gBGColor.processHeaderCell);
    table.addCell(tcell);
    
    var processActivityTables=new Array();
    for(var i=0;i<piNrOfProcesses;i++){        
        processActivityTables.push(createActivityTables(pFuncObjDefs[pStartIndex+i],[actRespHeight-2*lContentPadding,actDescHeight  -2*lContentPadding],lProcessWidth-20));
    }
    for(var i=0;i<piNrOfProcesses;i++){        
        //var activityTables=createActivityTables(pFuncObjDefs[pStartIndex+i],(outerTableHeight-firstOuterRowHeight)/3,lProcessWidth);
        var tcell=new com.lowagie.text.pdf.PdfPCell(processActivityTables[i][0]);
        //tcell.setFixedHeight((outerTableHeight-firstOuterRowHeight)/3);
        tcell.setPadding(lContentPadding);
        tcell.setBorderColor(gBGColor.processHeaderCell);
        table.addCell(tcell);
    }
    var tcell = new com.lowagie.text.pdf.PdfPCell(new com.lowagie.text.Phrase("Main activities", rowHeaderFont));
    tcell.setFixedHeight(actDescHeight) 
    tcell.setRotation(90);
    tcell.setHorizontalAlignment(com.lowagie.text.Element.ALIGN_CENTER);
    tcell.setVerticalAlignment(com.lowagie.text.Element.ALIGN_MIDDLE);
    tcell.setBackgroundColor(gBGColor.rowHeaderCell);
    tcell.setBorderColor(gBGColor.processHeaderCell);
    table.addCell(tcell);
    
    for(var i=0;i<piNrOfProcesses;i++){
        var tcell=new com.lowagie.text.pdf.PdfPCell(processActivityTables[i][1]);
        //tcell.setFixedHeight(actDescHeight);
        tcell.setPadding(lContentPadding);
        table.addCell(tcell);       
    }
    
    var canvas = gWriter.getDirectContent();
    table.writeSelectedRows(0, 6, gPdfSettings.leftMargin, gPdfSettings.pageSize.getHeight()- gPdfSettings.topMargin, canvas);
            
    printFooterAndHeader(pSelModel);    
}

function printFooterAndHeader(pSelModel){
    var businessOwnerAttrnr= ArisData.ActiveFilter().UserDefinedAttributeTypeNum("0c933f70-2ac7-11e0-5fa8-0050569d7bc2");
    var processDriverAttrnr= ArisData.ActiveFilter().UserDefinedAttributeTypeNum("94758b20-6f22-11e4-02fe-0050569d1299");
    var lastChangeDate = gDateFormatter.format(pSelModel.Attribute(Constants.AT_LAST_CHNG_2,gLoc).MeasureValue());
    var businessOwner = pSelModel.Attribute(businessOwnerAttrnr,gLoc).getValue();
    var processDriver = pSelModel.Attribute(processDriverAttrnr,gLoc).getValue();
        
    //create header
    var table = new com.lowagie.text.pdf.PdfPTable(1);    
    table.setTotalWidth(gPdfSettings.pageSize.getWidth()-gPdfSettings.leftMargin-gPdfSettings.rightMargin);
   
    var tableHeaderFont = getTextFont(true,18);
    
    var tcell = new com.lowagie.text.pdf.PdfPCell(new com.lowagie.text.Phrase(pSelModel.Name(gLoc), tableHeaderFont));    
    tcell.setHorizontalAlignment(com.lowagie.text.Element.ALIGN_CENTER);
    tcell.setVerticalAlignment(com.lowagie.text.Element.ALIGN_MIDDLE);
    tcell.setBorderWidth(0);
    table.addCell(tcell);
    
    var canvas = gWriter.getDirectContent();
    table.writeSelectedRows(0, 6, gPdfSettings.leftMargin, gPdfSettings.pageSize.getHeight() - gPdfSettings.topMargin+30, canvas);
    
    //footer    
    var table = new com.lowagie.text.pdf.PdfPTable(1);
    /*var intArray = java.lang.reflect.Array.newInstance(java.lang.Integer.TYPE, 2);
    intArray[0]=0;
    intArray[1]=100;
    table.setWidths(intArray);*/
    table.setTotalWidth(gPdfSettings.pageSize.getWidth()-gPdfSettings.leftMargin-gPdfSettings.rightMargin);
    //table.setWidthPercentage(100);
    
    var tableFooterFont = getTextFont(false,8);
    tableFooterFont.setColor(new java.awt.Color(16777215));
    var footertext= "Last change: " + lastChangeDate + "    Business Owner: " + businessOwner + "   Process driver: " + processDriver
    var tFTcell = new com.lowagie.text.pdf.PdfPCell(new com.lowagie.text.Phrase(footertext, tableFooterFont));    
    tFTcell.setVerticalAlignment(com.lowagie.text.Element.ALIGN_MIDDLE);
    tFTcell.setHorizontalAlignment(com.lowagie.text.Element.ALIGN_CENTER);
    tFTcell.setBorderWidth(0);
    table.addCell(tFTcell);
    
    var footerImageBytes = Context.getFile("Grundfos Footer.png", Constants.LOCATION_SCRIPT);
    var footerImage = com.lowagie.text.Image.getInstance(footerImageBytes);
    footerImage.scaleToFit(gPdfSettings.pageSize.getWidth()-gPdfSettings.leftMargin-gPdfSettings.rightMargin, gPdfSettings.bottomMargin-20)
    
    footerImage.setAbsolutePosition(gPdfSettings.leftMargin, 10);
    gDocument.add(footerImage);
    
    //var tcell = new com.lowagie.text.pdf.PdfPCell(footerImage);
    //tcell.setBorderWidth(0);
    //tcell.setHorizontalAlignment(com.lowagie.text.Element.ALIGN_RIGHT);
    //tcell.addElement(tFTParagraph);
    //table.addCell(tcell);
    var canvas = gWriter.getDirectContent();
    table.writeSelectedRows(0, 6, gPdfSettings.leftMargin, gPdfSettings.bottomMargin/2, canvas);
}
function createObjectOverviewPage(pSelModel){
    gDocument.newPage();
    
    var pageTitleFont = getTextFont(true,14);
    var tableHeaderFont = getTextFont(true,10);
    var cellFont = getTextFont(false,10);
    
    var pageTitle = new com.lowagie.text.Paragraph("Complete overview of documents in the process",pageTitleFont);
    pageTitle.setAlignment(com.lowagie.text.Element.ALIGN_LEFT);
    pageTitle.setSpacingAfter(20);    
    gDocument.add(pageTitle);
    
    /*var subTitle = new com.lowagie.text.Paragraph("Guides, best practices, examples, etc.",subTitleFont);
    subTitle.setAlignment(com.lowagie.text.Element.ALIGN_LEFT);
    subTitle.setSpacingBefore(20);
    gDocument.add(subTitle);
    */
        
    var outerTableWidth=gPdfSettings.pageSize.getWidth()-gPdfSettings.leftMargin-gPdfSettings.rightMargin;
    var outerTableHeight=gPdfSettings.pageSize.getHeight()-gPdfSettings.topMargin-gPdfSettings.bottomMargin;
    var intArray = java.lang.reflect.Array.newInstance(java.lang.Integer.TYPE, 3);
    intArray[0]=333;
    intArray[1]=333;
    intArray[2]=333;
    
    var table = new com.lowagie.text.pdf.PdfPTable(3);
    table.setWidths(intArray);
    table.setTotalWidth(outerTableWidth);
    table.setWidthPercentage(100);
    
    var tcell = new com.lowagie.text.pdf.PdfPCell(new com.lowagie.text.Phrase("Guides, best practices, examples, etc.", tableHeaderFont));
    tcell.setHorizontalAlignment(com.lowagie.text.Element.ALIGN_LEFT);
    tcell.setVerticalAlignment(com.lowagie.text.Element.ALIGN_MIDDLE);
    tcell.setBackgroundColor(gBGColor.activityHeaderCell);
    tcell.setBorderColor(gBGColor.activityHeaderCell);
    table.addCell(tcell);
    
    var tcell = new com.lowagie.text.pdf.PdfPCell(new com.lowagie.text.Phrase("Documents, templates, instruction, procedure, etc.", tableHeaderFont));
    tcell.setHorizontalAlignment(com.lowagie.text.Element.ALIGN_LEFT);
    tcell.setVerticalAlignment(com.lowagie.text.Element.ALIGN_MIDDLE);
    tcell.setBackgroundColor(gBGColor.activityHeaderCell);
    tcell.setBorderColor(gBGColor.activityHeaderCell);
    table.addCell(tcell);
    
    var tcell = new com.lowagie.text.pdf.PdfPCell(new com.lowagie.text.Phrase("Roles", tableHeaderFont));
    tcell.setHorizontalAlignment(com.lowagie.text.Element.ALIGN_LEFT);
    tcell.setVerticalAlignment(com.lowagie.text.Element.ALIGN_MIDDLE);
    tcell.setBackgroundColor(gBGColor.activityHeaderCell);
    tcell.setBorderColor(gBGColor.activityHeaderCell);
    table.addCell(tcell);
    
    table.addCell(getObjectListCell([Constants.OT_DOC_KNWLDG],cellFont));
    table.addCell(getObjectListCell([Constants.OT_INFO_CARR],cellFont));
    table.addCell(getObjectListCell([Constants.OT_PERS_TYPE,Constants.OT_POS,Constants.OT_GROUP, Constants.OT_ORG_UNIT],cellFont));
    
    gDocument.add(table);    
    
    printFooterAndHeader(pSelModel);
}
function createActivityTables(pObjDef, pMaxHeights,pWidth){        
    var intArray = java.lang.reflect.Array.newInstance(java.lang.Integer.TYPE, 4);
    
    var actRespTable = new com.lowagie.text.pdf.PdfPTable(1);
    actRespTable.setWidthPercentage(100);
    actRespTable.setTotalWidth(pWidth);
    
    var actDescTable = new com.lowagie.text.pdf.PdfPTable(1);
    actDescTable.setWidthPercentage(100);
    actDescTable.setTotalWidth(pWidth);
    
    var cellFont = getTextFont(false,8);
    var cellHeaderFont = getTextFont(true,8);
    
    intArray[0]=30;
    intArray[1]=30;
    intArray[2]=20;
    intArray[3]=20;
    
    var respTable = new com.lowagie.text.pdf.PdfPTable(4);
    respTable.setWidths(intArray);
    respTable.setTotalWidth(pWidth);
    
    var descTable = new com.lowagie.text.pdf.PdfPTable(1);    
    descTable.setTotalWidth(pWidth);
    
    respTable.addCell(getActivityCellHeader("Activity", cellHeaderFont));
    respTable.addCell(getActivityCellHeader("Output", cellHeaderFont));
    respTable.addCell(getActivityCellHeader("Responsible", cellHeaderFont));
    respTable.addCell(getActivityCellHeader("Accountable", cellHeaderFont));
    
    var resulTables=[actRespTable,actDescTable];
    var assignedModels=pObjDef.AssignedModels(gFuncAssignedModelTypeMap.get(pObjDef.GUID()));
    if(assignedModels.length>0){
        var funcObjDefs=sortFunctions(assignedModels[0], Constants.SORT_Y, Constants.SORT_X);
        
        var rowSpanMapResp=new java.util.HashMap();
        var prevObjGuidsR="";
        var lastRowR=-1;
        var rowSpanMapAccountable=new java.util.HashMap();
        var prevObjGuidsA="";
        var lastRowA=-1;
        var rowRespObjDefMap=new java.util.HashMap();
        var rowAccountableObjDefMap=new java.util.HashMap();
        for (var i in funcObjDefs){
            var funcObjDef=funcObjDefs[i];
            /*if(funcObjDef.GUID().equals("0108afb5-81f3-11e4-14ea-8d92e160db95")){
                var t=1;
            }*/
                
            var respObjDefs=getFADConnectedObjDefs(funcObjDef,[Constants.CT_EXEC_1,Constants.CT_EXEC_2],Constants.EDGES_IN);           
            var objGuids=getObjGuids(respObjDefs);
            if(objGuids!=prevObjGuidsR){
                rowSpanMapResp.put(i,1);
                lastRowR=i;
                prevObjGuidsR=objGuids;
                rowRespObjDefMap.put(i,respObjDefs);
            }else{
                rowSpanMapResp.put(lastRowR,parseInt(rowSpanMapResp.get(lastRowR))+1);
            }
            var accountObjDefs=getFADConnectedObjDefs(funcObjDef,[Constants.CT_DECID_ON,Constants.CT_DECD_ON],Constants.EDGES_IN);           
            var objGuids=getObjGuids(accountObjDefs);
            if(objGuids!=prevObjGuidsA){
                rowSpanMapAccountable.put(i,1);
                lastRowA=i;
                prevObjGuidsA=objGuids;
                rowAccountableObjDefMap.put(i,accountObjDefs);
            }else{
                rowSpanMapAccountable.put(lastRowA,parseInt(rowSpanMapAccountable.get(lastRowA))+1);
            }
            
            var consultObjDefs=getFADConnectedObjDefs(funcObjDef,[Constants.CT_HAS_CONSLT_ROLE_IN_1,Constants.CT_CONTR_TO_1,Constants.CT_HAS_CONSLT_ROLE_IN_2,Constants.CT_CONTR_TO_2],Constants.EDGES_IN);           
            var informedObjDefs=getFADConnectedObjDefs(funcObjDef,[Constants.CT_MUST_BE_INFO_ABT_1,Constants.CT_MUST_BE_INFO_ABT_2],Constants.EDGES_IN);           
        }
               
        for(var i in funcObjDefs){
            var funcObjDef=funcObjDefs[i]
            var outPutObjDefs=getFADConnectedObjDefs(funcObjDef,[Constants.CT_HAS_OUT],Constants.EDGES_OUT);                       
            var guidesObjDefs=getFADConnectedObjDefs(funcObjDef,[Constants.CT_IS_NEEDED_BY],Constants.EDGES_IN);
            var documentObjDefs=getFADConnectedObjDefs(funcObjDef,[Constants.CT_PROV_INP_FOR,Constants.CT_CRT_OUT_TO],Constants.EDGES_IN);
            
            respTable.addCell(getActivityCell(funcObjDef.Name(gLoc), cellFont));            
            respTable.addCell(getActivityRelationCell(outPutObjDefs, cellFont, null));
            var rowSpan=rowSpanMapResp.get(i);
            if(rowSpan!=null){
                var respObjDefs=rowRespObjDefMap.get(i);
                var tcell=getActivityRelationCell(respObjDefs, cellFont, Constants.AT_EXT_1);
                tcell.setRowspan(parseInt(rowSpan));
                if(parseInt(rowSpan)>0) tcell.setVerticalAlignment(com.lowagie.text.Element.ALIGN_MIDDLE);
                respTable.addCell(tcell);
            }
            var rowSpan=rowSpanMapAccountable.get(i);
            if(rowSpan!=null){
                var accountObjDefs=rowAccountableObjDefMap.get(i);
                var tcell=getActivityRelationCell(accountObjDefs, cellFont, Constants.AT_EXT_1);
                tcell.setRowspan(parseInt(rowSpan));
                if(parseInt(rowSpan)>0) tcell.setVerticalAlignment(com.lowagie.text.Element.ALIGN_MIDDLE);
                respTable.addCell(tcell);
            }                                   
            
            var tcell = new com.lowagie.text.pdf.PdfPCell(new com.lowagie.text.Phrase(funcObjDef.Name(gLoc), cellHeaderFont)); 
            tcell.setBorderWidth(0);
            descTable.addCell(tcell);

            var tcell = new com.lowagie.text.pdf.PdfPCell(new com.lowagie.text.Phrase(funcObjDef.Attribute(Constants.AT_DESC,gLoc).getValue(), cellFont));        
            tcell.setBorderWidth(0);
            descTable.addCell(tcell);
        }
        
        var tables=[respTable,descTable];
        for (var i in resulTables){
            var table=resulTables[i];
            var tableHeight=tables[i].calculateHeights(true);
            var tcell=new com.lowagie.text.pdf.PdfPCell(tables[i]);
            tcell.setBorderWidth(0);
            if(tableHeight>pMaxHeights[i]){
                tcell.setFixedHeight(pMaxHeights[i]);
            }
            table.addCell(tcell);
            var tcell = new com.lowagie.text.pdf.PdfPCell(new com.lowagie.text.Phrase("", cellFont));
            tcell.setBorderWidth(0);
            table.addCell(tcell);        
        }
        
    }
    return resulTables;
}

function getObjectListCell(pObjectTypeNums,pCellFont){
    
    var sortedMap=new java.util.TreeMap();
    for each(var objectTypeNum in pObjectTypeNums){
        var objectSet=gObjectMap.get(objectTypeNum);
        if(objectSet!=null){
            var iter=objectSet.iterator();
            while(iter.hasNext()){
                var objDef=iter.next();
                sortedMap.put(objDef.Name(gLoc) + "_" + objDef.GUID(), objDef);
            }
        }
    }
    var objDefs=sortedMap.values().toArray();
    return getActivityRelationCell(objDefs,pCellFont,Constants.AT_EXT_1)
    
}
function getActivityCell(pVal, pCellFont){
    var tcell = new com.lowagie.text.pdf.PdfPCell(new com.lowagie.text.Phrase(pVal, pCellFont));    
    tcell.setHorizontalAlignment(com.lowagie.text.Element.ALIGN_LEFT);
    tcell.setVerticalAlignment(com.lowagie.text.Element.ALIGN_TOP);
    tcell.setBorderColor(gBGColor.activityHeaderCell);
    return tcell;
}

function getActivityCellHeader(pVal, pCellFont){
    var tcell = new com.lowagie.text.pdf.PdfPCell(new com.lowagie.text.Phrase(pVal, pCellFont));    
    tcell.setHorizontalAlignment(com.lowagie.text.Element.ALIGN_LEFT);
    tcell.setVerticalAlignment(com.lowagie.text.Element.ALIGN_TOP);
    tcell.setBackgroundColor(gBGColor.activityHeaderCell);
    tcell.setBorderColor(gBGColor.activityHeaderCell);
    return tcell;
}


function getActivityRelationCell(pRelObjDefs, pCellFont,pLinkAttrNr){
    
    var tcell = null; 
    var linkFont=new com.lowagie.text.Font(pCellFont)
    //linkFont.setStyle("underline");
    linkFont.setColor(new java.awt.Color(26804));
    for each(var relObjDef in pRelObjDefs){        
        var linkVal=null;
        if(pLinkAttrNr!=null){
            var linkAttr=relObjDef.Attribute(pLinkAttrNr,gLoc);
            if(linkAttr.IsMaintained()){
                linkVal=linkAttr.getValue();
            }
        }
        var lPhrase=null;
        if(linkVal!=null){
            var linkChunk = new com.lowagie.text.Chunk(relObjDef.Name(gLoc), linkFont);
            linkChunk.setAction(new com.lowagie.text.pdf.PdfAction(linkVal));
            if(tcell==null){
                tcell=new com.lowagie.text.pdf.PdfPCell();
                lPhrase=tcell.setPhrase(new com.lowagie.text.Phrase(linkChunk));
            }else{
                lPhrase=tcell.getPhrase();
                lPhrase.add(new com.lowagie.text.Phrase("\r\n", pCellFont));
                lPhrase.add(linkChunk);
            }
            
        }else{            
            if(tcell==null){
                tcell=new com.lowagie.text.pdf.PdfPCell();
                lPhrase=tcell.setPhrase(new com.lowagie.text.Phrase(relObjDef.Name(gLoc), pCellFont));
            }else{
                lPhrase=tcell.getPhrase();
                lPhrase.add(new com.lowagie.text.Phrase("\r\n", pCellFont));
                lPhrase.add(new com.lowagie.text.Phrase(relObjDef.Name(gLoc), pCellFont));
            }
        }                        
    }
    if(tcell==null) tcell=new com.lowagie.text.pdf.PdfPCell();
    tcell.setHorizontalAlignment(com.lowagie.text.Element.ALIGN_LEFT);
    tcell.setVerticalAlignment(com.lowagie.text.Element.ALIGN_TOP);
    tcell.setPaddingBottom(5); 
    tcell.setBorderColor(gBGColor.activityHeaderCell);
    return tcell;
}
function sortFunctions(pModel,pSortCrit1, pSortCrit2){
    var funcObjDefs= new Array();
    var funcObjOccs=pModel.ObjOccListFilter(Constants.OT_FUNC);
    funcObjOccs=ArisData.sort(funcObjOccs, pSortCrit1, pSortCrit2, gLoc);
    for each(var funcObjOcc in funcObjOccs){
        if(funcObjOcc.getSymbol ( )!=Constants.ST_PRCS_IF){
            if(!getAttrBoolVal(funcObjOcc.ObjDef(),gExludeAttrNr,gLoc,false)){
                funcObjDefs.push(funcObjOcc.ObjDef());
            }
        }
    }
    return funcObjDefs;
}

function getFADConnectedObjDefs(pFuncObjDef,pCxnTypeNums, pDirection){
    var cxnTypeNumSet= new java.util.HashSet();    
    var resultObjDefs=new Array();    
    var FADModels=pFuncObjDef.AssignedModels(Constants.MT_FUNC_ALLOC_DGM);
    
    for each(var cxnTypeNum in pCxnTypeNums){
        cxnTypeNumSet.add(cxnTypeNum);
    }
    if(FADModels.length>0){
        var funcObjOccs=pFuncObjDef.OccListInModel(FADModels[0]);
        if(funcObjOccs.length>0){
            for each(var funcObjOcc in funcObjOccs){
                for each(var cxnOcc in funcObjOcc.Cxns(pDirection)){
                    if(cxnTypeNumSet.contains(cxnOcc.CxnDef().TypeNum())){
                        if(pDirection==Constants.EDGES_IN){
                            resultObjDefs.push(cxnOcc.CxnDef().SourceObjDef());
                        }else{
                            resultObjDefs.push(cxnOcc.CxnDef().TargetObjDef()); 
                        }
                    }
                }
            }
        }
    }
    for each(var resultObjDef in resultObjDefs){
        var objectSet=gObjectMap.get(resultObjDef.TypeNum());
        if(objectSet== null) objectSet= new java.util.HashSet();
        objectSet.add(resultObjDef);
        gObjectMap.put(resultObjDef.TypeNum(),objectSet)
    }
    
    resultObjDefs=ArisData.Unique(resultObjDefs);
    return ArisData.sort(resultObjDefs,Constants.AT_NAME,gLoc);
}

function getObjGuids(pObjDefs){
    var resultText="";
    for each(objDef in pObjDefs){
        resultText+=objDef.GUID();
    }
    return resultText;
}
function getTextFont(isBold, pSize){
    var font = new com.lowagie.text.Font(com.lowagie.text.Font.HELVETICA);
    font.setColor(java.awt.Color.BLACK);
    
    if(gPageSize=="A0"){
        pSize=pSize+2;
    }
    font.setSize(pSize);
    
    if(isBold){
        font.setStyle(com.lowagie.text.Font.BOLD);
    }
    
    return font;
}
function getAttrBoolVal(pArisItem, piAttrNr, piLoc, pdDefault){
    var lAttr=pArisItem.Attribute(piAttrNr,piLoc);
    if(lAttr.IsMaintained()){
        var laAttrValues = ArisData.getActiveDatabase().ActiveFilter().AttrValueTypeNums(piAttrNr);
        if (laAttrValues.length == 2) {
            var liRefValue;                      // refValue is the geater value
            if (laAttrValues[0] > laAttrValues[1]) {
                liRefValue = laAttrValues[0];
            } else {
                liRefValue = laAttrValues[1];
            }
        }
        if(lAttr.MeasureUnitTypeNum() == liRefValue){                
            return true;
        }else{
            return false;
        }
    }else{
        return pdDefault;
    }        
}

main();