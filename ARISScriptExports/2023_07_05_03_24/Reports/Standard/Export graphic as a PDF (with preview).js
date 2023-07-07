/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

// --------------------------------------------------------------------
// ----- MAIN part
// --------------------------------------------------------------------
    var SHOW_DIALOG = true  // Show options dialog - Default=true

    var FORCE_VECTOR_IMAGES = true //this would force this report to use vector image even if the image is larger than 55,87cm (will be scaled down to 55,87cm max h/w)
    var USE_VECTOR_FOR_1_IMAGE = true //true would use vector image for 1-model-selections of size<=558.7 x 558.7mm

    // output settings
    var PAGE_WIDTH    = 210
    var PAGE_HEIGHT   = 297
    var MARGIN_LEFT   = 10
    var MARGIN_RIGHT  = 10    
    var MARGIN_TOP    = 15
    var MARGIN_BOTTOM = 10
    var DIST_TOP      = 5
    var DIST_BOTTOM   = 5
    var CUT_OBJECTS   = false
    var FULL_SIZE     = true
    var ZOOM          = -1
    var USE_BGCOLOR   = false       // BLUE-13448 Use background color

    var oOutput = Context.createOutputObject();
    
    if (showOptionsDialog()) {

        var aModels = ArisData.getSelectedModels()
        var nLocale = Context.getSelectedLanguage();

        if(FULL_SIZE && aModels.length==1)//only 1 model. check if it is small enough to use new output object and emf output
        {
            createOneModelDocument( aModels[0])
        }
        else
        {
            if(FORCE_VECTOR_IMAGES)
                Context.setProperty("use-new-output", true); 
            else
                Context.setProperty("use-new-output", false); 
            
            oOutput = Context.createOutputObject();     // Necessary here because depends on property settings above
            setupOutputDefault( oOutput, PAGE_WIDTH, PAGE_HEIGHT );
            if (FULL_SIZE) {
                createDocument(oOutput, aModels);
            } else {
                createDocumentAdaptedToPage(oOutput, aModels);
            }
        }
        oOutput.WriteReport();
    } else {
        Context.setScriptError(Constants.ERR_CANCEL);  
    }

    
// --------------------------------------------------------------------
// ----- FUNCTIONS
// --------------------------------------------------------------------

// Apply default page format settings to output object
function setupOutputDefault( outputObj, p_pageWidth, p_pageHeight ){
    outputObj.DefineF(getString("ID_STYLE_RD_HEADER_FOOTER"), getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_LEFT| Constants.FMT_VTOP, 0, 0, 0, 0, 0, 1);
	outputObj.SetPageWidth(p_pageWidth);
	outputObj.SetPageHeight(p_pageHeight);
	outputObj.SetLeftMargin(MARGIN_LEFT);
	outputObj.SetRightMargin(MARGIN_RIGHT);
	outputObj.SetTopMargin(MARGIN_TOP);
	outputObj.SetBottomMargin(MARGIN_BOTTOM);
	outputObj.SetDistHeader(DIST_TOP);
	outputObj.SetDistFooter(DIST_BOTTOM);
	outputObj.SetAutoTOCNumbering(true);
    outputObj.SetTitle(Context.getScriptInfo(Constants.SCRIPT_NAME));
}

function setHeader( outputObj, pTxt ){
	outputObj.BeginHeader();
    outputObj.OutputLnF( pTxt, getString("ID_STYLE_RD_HEADER_FOOTER") );
	outputObj.EndHeader();
}

//java.awt.Dimension in mm
function getImageSize(graphic, p_zoomFactor)
{
    var nGWidth     = p_zoomFactor * graphic.getWidth(Constants.SIZE_LOMETRIC);
    var nGHeight    = p_zoomFactor * graphic.getHeight(Constants.SIZE_LOMETRIC);
    
    // define a minimum size (for very small models) of 5cm x 5cm
    nGWidth  = Math.max(500, nGWidth)
    nGHeight = Math.max(500, nGHeight)

    if(FORCE_VECTOR_IMAGES)
    {
        var maxwidth  = Math.min(5587-(MARGIN_LEFT*10+MARGIN_RIGHT*10), nGWidth)
        var maxheight = Math.min(5587-(MARGIN_TOP*10+MARGIN_BOTTOM*10+50), nGHeight)
        
        if(maxwidth<nGWidth || maxheight<nGHeight)
        {
            var xfactor = maxwidth / nGWidth
            var yfactor = maxheight / nGHeight
            var factor  = Math.min(xfactor, yfactor)
            
            nGWidth  = nGWidth  * factor
            nGHeight = nGHeight * factor
        }
    }
    else
    {
        // AGA-13846 Otherwise to high memory usage up to 6GB
        // define a maximum size (for very small models) of 3m x 3m 
        nGWidth  = Math.min(30000, nGWidth)
        nGHeight = Math.min(30000, nGHeight)
    }
    
    var result = new java.awt.Dimension()
    result.setSize(nGWidth, nGHeight)
    return result;
}

function getPageSize(imageDimension)
{
    var result = new java.awt.Dimension()
    result.setSize(imageDimension.width+MARGIN_LEFT*10+MARGIN_RIGHT*10, imageDimension.height+MARGIN_TOP*10+MARGIN_BOTTOM*10+50)
    return result
}

function getZoomFactor(p_model) {   // BLUE-14499
    var nZoom = ZOOM;
    if (nZoom == -1)   nZoom = 100;
    if (nZoom == null) nZoom = p_model.getPrintScale();
    if (nZoom > 100) nZoom = 100;
    return nZoom/100;
}

function createOneModelDocument(p_model)
{
    var aGraphic = p_model.Graphic( false, false, nLocale, !USE_BGCOLOR );	// BLUE-13448  
    var dimSizeImage = getImageSize(aGraphic, getZoomFactor(p_model))
    var dimSizePage  = getPageSize(dimSizeImage)

    if(dimSizePage.getWidth()<=5587 && dimSizePage.getHeight()<=5587 && USE_VECTOR_FOR_1_IMAGE) //maximum dimensions for MS WORD = maximum dimensions for new output object
    {
        Context.setProperty("use-new-output", true); 
        Context.setProperty("model-as-emf", true); 
    }
    else
    {
        Context.setProperty("use-new-output", false); 
    }
    
    oOutput = Context.createOutputObject();     // Necessary here because depends on property settings above
    setupOutputDefault( oOutput, dimSizePage.getWidth()/10.0, dimSizePage.getHeight()/10.0 );
    setHeader(oOutput, getHeader(p_model));
    oOutput.OutGraphic( aGraphic, -1, dimSizeImage.width/10.0, dimSizeImage.height/10.0 );    
}

function createDocument(p_output, p_aModel)
{
	// do not create new section if all data is empty
	if(p_aModel.length==0)    return;
    
    if(FORCE_VECTOR_IMAGES)
        Context.setProperty("model-as-emf", true); 
        
	// repetition of queried data:
	for(var i=0; i<p_aModel.length; i++) 
    {
        var aGraphic     = p_aModel[i].Graphic( false, false, nLocale, !USE_BGCOLOR );	// BLUE-13448
        var dimSizeImage = getImageSize(aGraphic, getZoomFactor(p_aModel[i]))
        var dimSizePage  = getPageSize(dimSizeImage)

        p_output.BeginSection(dimSizePage.height/10.0, dimSizePage.width/10.0,
                              DIST_TOP, DIST_BOTTOM, MARGIN_LEFT, MARGIN_RIGHT, MARGIN_TOP, MARGIN_BOTTOM, 
                              false, Constants.SECTION_DEFAULT);

        setHeader(p_output, getHeader(p_aModel[i]));
        p_output.OutGraphic( aGraphic, -1, dimSizeImage.width/10.0, dimSizeImage.height/10.0 );

    	p_output.EndSection();
	}
}

function createDocumentAdaptedToPage(p_output, p_aModel)
{
	// do not create new section if all data is empty
	if(p_aModel.length==0)    return;
    
    if(FORCE_VECTOR_IMAGES)
        Context.setProperty("model-as-emf", true); 
        
	// repetition of queried data:
	for(var i=0; i<p_aModel.length; i++) 
    {
        var aGraphic     = p_aModel[i].Graphic( CUT_OBJECTS, false, nLocale, !USE_BGCOLOR );	// BLUE-13448
        var dimSizeImage = getImageSizeAdaptedToPage()
        var dimSizePage  = getPageSizeAdaptedToPage()
        
        p_output.BeginSection(dimSizePage.height/10.0, dimSizePage.width/10.0,
                              DIST_TOP, DIST_BOTTOM, MARGIN_LEFT, MARGIN_RIGHT, MARGIN_TOP, MARGIN_BOTTOM, 
                              false, Constants.SECTION_DEFAULT);

        setHeader(p_output, getHeader(p_aModel[i]));
        if (ZOOM != null) {
            p_output.OutGraphic( aGraphic, ZOOM, dimSizeImage.width/10.0, dimSizeImage.height/10.0 );    
        } else {
            p_output.OutGraphic( aGraphic, p_aModel[i].getPrintScale(), dimSizeImage.width/10.0, dimSizeImage.height/10.0);    
        }

    	p_output.EndSection();
	}
    
    function getImageSizeAdaptedToPage() {
        var result = new java.awt.Dimension()
        var nWidth = PAGE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT
        var nHeight = PAGE_HEIGHT - MARGIN_TOP - MARGIN_BOTTOM
        result.setSize(nWidth*10, nHeight*10 - 20)
        return result    
    }    
    
    function getPageSizeAdaptedToPage() {
        var result = new java.awt.Dimension()
        result.setSize(PAGE_WIDTH*10, PAGE_HEIGHT*10)
        return result    
    }
}

function getHeader(oModel) {
    var format;    
    try {
        format = new java.text.SimpleDateFormat(getString("DATE_FORMAT"));        
    } catch(ex) {
        format = new java.text.SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
    }
    format.setTimeZone( Context.getClientTimeZone() )
	
    var sHeadline = getString("HEADLINE");  // BLUE-11460
    if (sHeadline == "") sHeadline = "@1 (@2; last change by @3 at @4)";
    return formatHeader(sHeadline, 
                    oModel.Name(nLocale, true),
                    oModel.Type(),
                    oModel.Attribute(Constants.AT_LUSER, nLocale).getValue(),
                    formatDate(oModel, format));                                // BLUE-12303

    function formatHeader(stext, svalue1, svalue2, svalue3, svalue4) {
        var snewtext = (new String(stext)).replace("@1", svalue1); // ensure we have a js-String
        snewtext = snewtext.replace("@2", svalue2);
        snewtext = snewtext.replace("@3", svalue3);
        snewtext = snewtext.replace("@4", svalue4);
        return snewtext;
    }
    
    function formatDate(oModel, format) {
        var attr = oModel.Attribute(Constants.AT_LAST_CHNG_2, nLocale);
        if (attr.IsValid() && attr.IsMaintained()) {
            return format.format(attr.MeasureValue());
        }
        return "";
    }    
}

function tPaperFormatType(sName, nWidth, nHeight) {
    this.sName = sName;
    this.nWidth = nWidth;
    this.nHeight = nHeight;
    return this;
} 

function showOptionsDialog() {
    if ( SHOW_DIALOG && !startedByTA() ) {
        return (Dialogs.showDialog(new optionsDialog(), Constants.DIALOG_TYPE_ACTION, getString("DLG_TITLE")) == true);
    }
    return true;
    
    function startedByTA() {
        return Context.getProperty("no-dialogs") != null;       // This property is set by TA
    }
}    

function optionsDialog() {  
    var bClosedByOk = false;
    
    var aPaperFormats = getPaperFormats();
    var aPaperFormatNames = getPaperFormatNames(aPaperFormats);
    
    var sSection = "SCRIPT_f7015530_60ab_11e2_1592_5cf9dd7305ef"; 
   
    this.getPages = function() {
        var iDialogTemplate = Dialogs.createNewDialogTemplate(585, 295, getString("DLG_TITLE"));
        
        iDialogTemplate.GroupBox(5, 10, 280, 80, getString("DLG_CUT"));
        iDialogTemplate.OptionGroup("OPTION_CUT");
        iDialogTemplate.OptionButton(15, 20, 250, 15, getString("DLG_USEMODELSIZE"), "OPT_NOCUT");
        iDialogTemplate.OptionButton(15, 36, 250, 15, getString("DLG_CUTOBJECTS"));
        iDialogTemplate.OptionButton(15, 52, 250, 15, getString("DLG_OVERLAYOBJECTS"));
        
        iDialogTemplate.GroupBox(300, 10, 280, 80, getString("DLG_SCALING"));
        iDialogTemplate.OptionGroup("OPTION_SCALE");
        iDialogTemplate.OptionButton(310, 20, 80, 15, getString("DLG_100PERCENT"));
        iDialogTemplate.OptionButton(310, 36, 135, 15, getString("DLG_USERDEFINED"));
        iDialogTemplate.OptionButton(310, 52, 250, 15, getString("DLG_PRINTSCALE"));
        iDialogTemplate.OptionButton(310, 68, 250, 15, getString("DLG_SCALETOFIT"), "OPT_FIT");
        iDialogTemplate.TextBox(445, 33, 60, 20, "TEXT_SCALE");
        iDialogTemplate.Text(510, 36, 80, 15, "% (10-100)");      
        
        iDialogTemplate.GroupBox(5, 100, 280, 200, getString("DLG_PAPERFORMAT"));
        iDialogTemplate.DropListBox(20, 115, 250, 20, aPaperFormatNames, "LIST_FORMAT");
        iDialogTemplate.Text(20, 145, 100, 15, getString("DLG_WIDTH"));
        iDialogTemplate.TextBox(85, 142, 50, 20, "TEXT_WIDTH");
        iDialogTemplate.Text(145, 145, 120, 15, getString("DLG_MM_MIN_50"));
        iDialogTemplate.Text(20, 175, 100, 15, getString("DLG_HEIGHT"));
        iDialogTemplate.TextBox(85, 172, 50, 20, "TEXT_HEIGHT");    
        iDialogTemplate.Text(145, 175, 120, 15, getString("DLG_MM_MIN_50"));
        iDialogTemplate.GroupBox(20, 205, 250, 50, getString("DLG_ORIENTATION"));
        iDialogTemplate.OptionGroup("OPTION_ORIENTATION");
        iDialogTemplate.OptionButton(30, 215, 150, 15, getString("DLG_PORTRAIT"), "OPT_PORTRAIT");
        iDialogTemplate.OptionButton(30, 231, 150, 15, getString("DLG_LANDSCAPE"), "OPT_LANDSCAPE");
        
        iDialogTemplate.GroupBox(300, 100, 280, 135, getString("DLG_MARGINS"));
        iDialogTemplate.Text(315, 115, 100, 15, getString("DLG_TOP"));
        iDialogTemplate.TextBox(380, 112, 50, 20, "TEXT_TOP");
        iDialogTemplate.Text(440, 115, 25, 15, "mm");
        iDialogTemplate.Text(315, 145, 100, 15, getString("DLG_BOTTOM"));
        iDialogTemplate.TextBox(380, 142, 50, 20, "TEXT_BOTTOM");
        iDialogTemplate.Text(440, 145, 25, 15, "mm");
        iDialogTemplate.Text(315, 175, 100, 15, getString("DLG_LEFT"));
        iDialogTemplate.TextBox(380, 172, 50, 20, "TEXT_LEFT");
        iDialogTemplate.Text(440, 175, 25, 15, "mm");
        iDialogTemplate.Text(315, 205, 100, 15, getString("DLG_RIGHT"));
        iDialogTemplate.TextBox(380, 202, 50, 20, "TEXT_RIGHT");
        iDialogTemplate.Text(440, 205, 25, 15, "mm");

        iDialogTemplate.GroupBox(300, 245, 280, 40, getString("DLG_BGCOLOR"));						// BLUE-13448
        iDialogTemplate.CheckBox(315, 255, 250, 15, getString("DLG_USEBGCOLOR"), "CHECK_BGCOLOR");
        
        iDialogTemplate.GroupBox(600, 10, 280, 303, getString("DLG_MODEL_PREVIEW"));
        iDialogTemplate.Picture(610, 20, 258, 285, "MODEL_PREVIEW");
        iDialogTemplate.Picture(610, 20, 258, 285, "PAGES_PREVIEW");
        iDialogTemplate.Text(613, 20, 200, 100, getString("DLG_NO_PREVIEW_AVAILABLE"), "NO_PREVIEW_AVAILABLE");

        return [iDialogTemplate];
    }
    
    // BLUE-14786
    this.getThisDialogElement = function(sField) {
        try{
            return this.dialog.getPage(0).getDialogElement(sField)            
        } catch(e) {
            return this.dialog.getDialogElement(sField)
        }
    }
    
    this.init = function(aPages) {
        var page = aPages[0];
                
        readSettings(page, sSection);
        this.turnPaperSize();

        var bEnableControls = (page.getDialogElement("OPTION_CUT").getValue() != 0);
        this.enableControls(bEnableControls);

        var bEnableScale = (page.getDialogElement("OPTION_SCALE").getValue() ==  1);
        page.getDialogElement("TEXT_SCALE").setEnabled(bEnableScale);
        
        this.updatePageSettingsModel();
        this.updatePrintPreview(page);
    }    

    this.LIST_FORMAT_selChanged = function(newSelection) {
        this.getThisDialogElement("TEXT_WIDTH").setText(aPaperFormats[newSelection].nWidth);
        this.getThisDialogElement("TEXT_HEIGHT").setText(aPaperFormats[newSelection].nHeight);
        PAGE_WIDTH  = parseInt(aPaperFormats[newSelection].nWidth, 10);
        PAGE_HEIGHT = parseInt(aPaperFormats[newSelection].nHeight, 10);
        this.updatePrintPreview();
    }
    
    this.TEXT_WIDTH_changed = function(newValue){
        PAGE_WIDTH = this.getNumericElementValue("TEXT_WIDTH", 50);
        if(PAGE_WIDTH < 50) PAGE_WIDTH = 50;
        this.updatePrintPreview();
    }

    this.TEXT_HEIGHT_changed = function(newValue){
        PAGE_HEIGHT = this.getNumericElementValue("TEXT_HEIGHT", 50);
        this.updatePrintPreview();
    }

    this.OPTION_SCALE_selChanged = function(newSelection) {
        var bEnableScale = (newSelection == 1);
        ZOOM             = this.getZoom();
        this.getThisDialogElement("TEXT_SCALE").setEnabled(bEnableScale);
        this.updatePrintPreview();
    }     
    
    
    this.OPTION_ORIENTATION_selChanged = function(newSelection) {
        this.turnPaperSize();
        PAGE_WIDTH  = this.getNumericElementValue("TEXT_WIDTH", 50);
        PAGE_HEIGHT = this.getNumericElementValue("TEXT_HEIGHT", 50);
        this.updatePrintPreview();
    }     

    this.OPTION_CUT_selChanged = function(newSelection) {
        var bEnableControls = (newSelection != 0);
        this.enableControls(bEnableControls);
        this.updatePrintPreview();
    }
    this.TEXT_SCALE_changed = function(newSelection){
        ZOOM = this.getZoom();
        this.updatePrintPreview();
    }
    this.TEXT_LEFT_changed = function(){
        MARGIN_LEFT   = this.getNumericElementValue("TEXT_LEFT", 0, PAGE_WIDTH/3);
        this.updatePrintPreview();
    }
    this.TEXT_RIGHT_changed = function(){
        MARGIN_RIGHT  = this.getNumericElementValue("TEXT_RIGHT", 0, PAGE_WIDTH/3);
        this.updatePrintPreview();
    }
    this.TEXT_TOP_changed = function(){
        MARGIN_TOP = this.getNumericElementValue("TEXT_TOP", 0, PAGE_HEIGHT/3);
        this.updatePrintPreview();
    }
    this.TEXT_BOTTOM_changed = function(){
        MARGIN_BOTTOM = this.getNumericElementValue("TEXT_BOTTOM", 0, PAGE_HEIGHT/3);
        this.updatePrintPreview();
    }
    
    this.getNumericElementValue = function(elementName, min, max, minFallbackValue, maxFallbackValue){
        return this.getNumericValue(this.getThisDialogElement(elementName).getText(), min, max, minFallbackValue, maxFallbackValue);
    }
    
    this.getNumericValue = function(value, min, max, minFallbackValue, maxFallbackValue){
        var number = parseInt(value, 10);
        var min    = parseInt(min, 10);
        var max    = parseInt(max, 10);
        var minFallbackValue = parseInt(minFallbackValue, 10);
        var maxFallbackValue = parseInt(maxFallbackValue, 10);
        
        if(isNaN(number) || !isNaN(min) && number < min){
            number = !isNaN(minFallbackValue) ? minFallbackValue : min;
        } else if(!isNaN(max) && number > max){
            number = !isNaN(maxFallbackValue) ? maxFallbackValue : max;
        }
        return number;
    }
    
    this.isInValidState = function(pageNumber) {
        if (this.getThisDialogElement("OPTION_SCALE").getValue() == 1 && !this.isValidScale())  return false;
        if (this.getThisDialogElement("OPTION_CUT").getValue() != 0   && !this.isValidMargin()) return false;
        if (this.getThisDialogElement("OPTION_CUT").getValue() != 0   && !this.isValidFormat()) return false;
        return true;
    }
    
    this.onClose = function(pageNumber, bOk) {
        bClosedByOk = bOk;
    }    

    this.getResult = function() {
        if (bClosedByOk) {
            this.writeSettings();
            this.updatePageSettingsModel();
            return true;
        }
        return false;
    }
    
    this.getZoom = function() {
        var nScaleOption = this.getThisDialogElement("OPTION_SCALE").getValue();
        switch(nScaleOption) {
            case 0: return 100;     // 100%
            case 1:
                return this.getNumericElementValue("TEXT_SCALE", 10, 100); // User-defined
            case 2: return null;    // Model print scale
            case 3: return -1;      // Scale to fit page size
        }
    }
    
    this.enableControls = function(bEnable) {
        this.getThisDialogElement("LIST_FORMAT").setEnabled(bEnable);
        this.getThisDialogElement("TEXT_WIDTH").setEnabled(bEnable);
        this.getThisDialogElement("TEXT_HEIGHT").setEnabled(bEnable);
        this.getThisDialogElement("OPT_PORTRAIT").setEnabled(bEnable);
        this.getThisDialogElement("OPT_LANDSCAPE").setEnabled(bEnable);
        this.getThisDialogElement("TEXT_TOP").setEnabled(bEnable);
        this.getThisDialogElement("TEXT_BOTTOM").setEnabled(bEnable);
        this.getThisDialogElement("TEXT_LEFT").setEnabled(bEnable);
        this.getThisDialogElement("TEXT_RIGHT").setEnabled(bEnable);

        this.getThisDialogElement("OPT_FIT").setEnabled(bEnable);
        if (!bEnable && this.getThisDialogElement("OPTION_SCALE").getValue() == 3) {
            this.getThisDialogElement("OPTION_SCALE").setSelection(0);
        } 
    }

    this.turnPaperSize = function() {
        var nOrientation = this.getThisDialogElement("OPTION_ORIENTATION").getValue();
        var nWidth = this.getThisDialogElement("TEXT_WIDTH").getText();
        var nHeight = this.getThisDialogElement("TEXT_HEIGHT").getText();

        if (nOrientation == 0 && nWidth > nHeight || nOrientation == 1 && nWidth < nHeight) {
            this.getThisDialogElement("TEXT_WIDTH").setText(nHeight);
            this.getThisDialogElement("TEXT_HEIGHT").setText(nWidth);
        }
    }

    this.isValidScale = function() {
        return this.isValidNumberField("TEXT_SCALE", 10, 100);
    }
    
    this.isValidMargin = function() {
        if(!this.isValidNumberField("TEXT_TOP",    0, PAGE_HEIGHT/3)) return false;
        if(!this.isValidNumberField("TEXT_LEFT",   0, PAGE_WIDTH/3 )) return false;
        if(!this.isValidNumberField("TEXT_RIGHT",  0, PAGE_WIDTH/3 )) return false;
        if(!this.isValidNumberField("TEXT_BOTTOM", 0, PAGE_HEIGHT/3)) return false;
        return true;
    }
    
    this.isValidFormat = function() {
        if(!this.isValidNumberField("TEXT_WIDTH",  50, 10000)) return false;
        if(!this.isValidNumberField("TEXT_HEIGHT", 50, 10000)) return false;
        return true;
    }
    
    this.isValidNumberField = function(fieldId, minVal, maxVal){
        var fieldVal = this.getThisDialogElement(fieldId).getText();
        if (isNoNumber(fieldVal)) return false;
        if(isNaN(minVal) || minVal != null && fieldVal < minVal) return false;
        if(isNaN(maxVal) || maxVal != null && fieldVal > maxVal) return false;
        return true;
    }
    
    this.writeSettings = function() {
        // Read dialog settings from config            
        this.writeSettings_Value(sSection, "OPTION_CUT");
        this.writeSettings_Value(sSection, "OPTION_SCALE");
        this.writeSettings_Text(sSection, "TEXT_SCALE");
        this.writeSettings_Value(sSection, "LIST_FORMAT");
        this.writeSettings_Text(sSection, "TEXT_WIDTH");
        this.writeSettings_Text(sSection, "TEXT_HEIGHT");
        this.writeSettings_Value(sSection, "OPTION_ORIENTATION");
        this.writeSettings_Text(sSection, "TEXT_TOP");
        this.writeSettings_Text(sSection, "TEXT_BOTTOM");
        this.writeSettings_Text(sSection, "TEXT_LEFT");
        this.writeSettings_Text(sSection, "TEXT_RIGHT");    
        this.writeSettings_Bool(sSection, "CHECK_BGCOLOR");
    }

    this.writeSettings_Value = function(sSection, sField) {
        var nValue = this.getThisDialogElement(sField).getValue();
        if (nValue < 0) nValue = 0;
        Context.writeProfileInt(sSection, sField, nValue);
    }
    this.writeSettings_Bool = function(sSection, sField) {
        var nValue = this.getThisDialogElement(sField).isChecked() ? 1 : 0;
        Context.writeProfileInt(sSection, sField, nValue);
    }
    this.writeSettings_Text = function(sSection, sField) {
        Context.writeProfileString(sSection, sField, this.getThisDialogElement(sField).getText());
    }
    
    this.updatePageSettingsModel = function(){
        FULL_SIZE   = this.getThisDialogElement("OPTION_CUT").getValue() == 0;      // Use model size (no cut)
        ZOOM        = this.getZoom();
        USE_BGCOLOR = this.getThisDialogElement("CHECK_BGCOLOR").isChecked();       // Use background color
        if (this.getThisDialogElement("OPTION_CUT").getValue() != 0) {
            PAGE_WIDTH    = parseInt(this.getThisDialogElement("TEXT_WIDTH").getText());
            PAGE_HEIGHT   = parseInt(this.getThisDialogElement("TEXT_HEIGHT").getText());
            MARGIN_LEFT   = parseInt(this.getThisDialogElement("TEXT_LEFT").getText());
            MARGIN_RIGHT  = parseInt(this.getThisDialogElement("TEXT_RIGHT").getText());
            MARGIN_TOP    = parseInt(this.getThisDialogElement("TEXT_TOP").getText());
            MARGIN_BOTTOM = parseInt(this.getThisDialogElement("TEXT_BOTTOM").getText());
            CUT_OBJECTS   = this.getThisDialogElement("OPTION_CUT").getValue() == 1;    // Cut objects at margin
        }
    }
    
    this.updatePrintPreview = function(page){
        this.getThisDialogElement("NO_PREVIEW_AVAILABLE").setVisible(false);
        if(!this.isInValidState()){
            this.drawPreviewFallback();
            return;
        }
        try {
            var model = ArisData.getSelectedModels()[0];
            var modelAttributes = this.drawModelAndGetAttributes(model);
            if(modelAttributes){
                this.drawPagesPreview(model, modelAttributes);
            } else {
                this.drawPreviewFallback();
            }
        } catch (e) {
            this.drawPreviewFallback();
        }
    }
    
    this.drawPreviewFallback = function(){
        this.getThisDialogElement("NO_PREVIEW_AVAILABLE").setVisible(true);
        this.getThisDialogElement("MODEL_PREVIEW").setPicture(null, "jpg");
        this.getThisDialogElement("PAGES_PREVIEW").setPicture(null, "png");
    }
    
    this.drawPagesPreview = function(model, modelAttributes){
        var image = Context.createPicture();
        image.SelectPen(Constants.PS_DOT, 1, 0);
        if (this.getThisDialogElement("OPTION_CUT").getValue() != 0) {
            var pageSize = this.calculatePreviewPageSize(model, modelAttributes);
            this.drawPages(image, pageSize, modelAttributes);
        } else {
            image.FrameRect(0, 0, modelAttributes.modelWidth, modelAttributes.modelHeight, 0);
        }
        var pagesPreview = this.getThisDialogElement("PAGES_PREVIEW");
        pagesPreview.setPicture(image.getAsOneImage(600, 1000, "png")[0], "png");
    }
    
    this.calculatePreviewPageSize = function(model, modelAttributes){
        var zoomFactor = getZoomFactor(model);
        var pageSize = [PAGE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT, PAGE_HEIGHT - MARGIN_TOP - MARGIN_BOTTOM];
        var coef     = 1.06*modelAttributes.graphicsZoom/10;
        var yPageCountCorrection = (modelAttributes.modelHeight / pageSize[1] / 850);
        pageSize = [pageSize[0]*coef, pageSize[1]*(coef - yPageCountCorrection)];
        if(this.getThisDialogElement("OPTION_SCALE").getValue() == 3){
            zoomFactor = this.getFitToElementZoomFactor(modelAttributes.modelWidth, modelAttributes.modelHeight, pageSize[0], pageSize[1]);
        }
        zoomFactor = (zoomFactor != 0 ? zoomFactor : 1);
        return [pageSize[0]/zoomFactor, pageSize[1]/zoomFactor];
    }
    
    this.drawPages = function(image, pageSize, modelAttributes){
        var currentY = 0;
        while(currentY < modelAttributes.modelHeight){
            var currentX = 0;
            while(currentX < modelAttributes.modelWidth){
                var color = getColorByRGB(142, 142, 142);
                image.FrameRect(currentX, currentY, currentX+pageSize[0]+1, currentY+pageSize[1]+1+(currentY == 0 ? 1 : 0), color);
                currentX += pageSize[0]-1;
            }
            currentY += pageSize[1]-1;
        }
    }
    
    this.drawModelAndGetAttributes = function(model){
        var node       = model.StartNodeList()[0];
        var aGraphic   = model.Graphic( false, false, Context.getSelectedLanguage(), !USE_BGCOLOR );
        var pic = this.getThisDialogElement("MODEL_PREVIEW");
        var modelAttributes = this.fitModelGraphicsToDialogElement(aGraphic, pic);
        if(modelAttributes){
            pic.setPicture(aGraphic.getAsOneImage(modelAttributes.graphicsZoom, "jpg")[0], "png");
        }
        return modelAttributes;
    }
    
    this.fitModelGraphicsToDialogElement = function(aGraphic, picElement){
        var modelWidth = aGraphic.getWidth(Constants.SIZE_LOMETRIC);
        var modelHeight = aGraphic.getHeight(Constants.SIZE_LOMETRIC);
        if(modelWidth > 20000 || modelHeight > 20000) return null;
        var elementWidth = 258*2.5;
        var elementHeight = 380*2.5; //380 is the height to which is image height recalculated with getAdjustedControlHeight in dialogbuilder
        var graphicsZoomFactor = this.getFitToElementZoomFactor(modelWidth, modelHeight, elementWidth, elementHeight);
        modelWidth  *= graphicsZoomFactor*1.06;
        modelHeight *= graphicsZoomFactor*1.06;
        if(graphicsZoomFactor < 0.01) return null;
        if(graphicsZoomFactor > 2) graphicsZoomFactor = 2;
        return {
            graphicsZoom: graphicsZoomFactor * 100,
            modelHeight: modelHeight,
            modelWidth: modelWidth
        };
    }
    
    this.getFitToElementZoomFactor = function(modelWidth, modelHeight, elementWidth, elementHeight){
        var zoomFactor;
        if((modelWidth/elementWidth) < (modelHeight/elementHeight)){
            zoomFactor = elementHeight/modelHeight;
        } else {
            zoomFactor = elementWidth/modelWidth;
        }
        return Math.floor(zoomFactor*100) / 100;
    }
}

function isNoNumber(sNumericText) {
    // BLUE-17443
    if (sNumericText.trim() == "") return true;
    if (sNumericText.trim().startsWith(".")) return true; // AKC-17578 - parseInt(".210") returns NaN, but isNaN(".210") returns false
    if (isNaN(sNumericText)) return true;
    return false;
}

function readSettings(page, sSection) {
    // Read dialog settings from config            
    ReadSettings_Value(page, sSection, "OPTION_CUT", 0);
    ReadSettings_Value(page, sSection, "OPTION_SCALE", 0);
    ReadSettings_Text(page, sSection, "TEXT_SCALE", 100);
    ReadSettings_Value(page, sSection, "LIST_FORMAT", 0);
    ReadSettings_Text(page, sSection, "TEXT_WIDTH", oOutput.GetPageWidth());
    ReadSettings_Text(page, sSection, "TEXT_HEIGHT", oOutput.GetPageHeight());
    ReadSettings_Value(page, sSection, "OPTION_ORIENTATION", 0);
    ReadSettings_Text(page, sSection, "TEXT_TOP", oOutput.GetTopMargin());
    ReadSettings_Text(page, sSection, "TEXT_BOTTOM", oOutput.GetBottomMargin());
    ReadSettings_Text(page, sSection, "TEXT_LEFT", oOutput.GetLeftMargin());
    ReadSettings_Text(page, sSection, "TEXT_RIGHT", oOutput.GetRightMargin());    
    ReadSettings_Bool(page, sSection, "CHECK_BGCOLOR", 0);
    
    function ReadSettings_Value(page, sSection, sField, nDefault) {
        page.getDialogElement(sField).setSelection(Context.getProfileInt(sSection, sField, nDefault));
    }
    function ReadSettings_Bool(page, sSection, sField, nDefault) {
        page.getDialogElement(sField).setChecked(Context.getProfileInt(sSection, sField, nDefault)==1);
    }
    function ReadSettings_Text(page, sSection, sField, sDefault) {
        var sText = Context.getProfileString(sSection, sField, sDefault);
        if (isNoNumber(sText)) {
            sText = sDefault;           // BLUE-17443
        }        
        page.getDialogElement(sField).setText(sText);        
    }
}
    
function getPaperFormats() {
    var aPaperFormats = new Array();
    aPaperFormats[0] = new tPaperFormatType(getString("DLG_USERDEFINED"), 210, 297);  // userdefined (BLUE-12786)
    aPaperFormats[1] = new tPaperFormatType(getString("DLG_A3"), 297, 420);           // A3
    aPaperFormats[2] = new tPaperFormatType(getString("DLG_A4"), 210, 297);           // A4
    aPaperFormats[3] = new tPaperFormatType(getString("DLG_A5"), 148, 210);           // A5
    aPaperFormats[4] = new tPaperFormatType(getString("DLG_LETTER"), 216, 280);       // Letter
    aPaperFormats[5] = new tPaperFormatType(getString("DLG_LEGAL"), 216, 356);        // Legal
    return aPaperFormats;
}

function getPaperFormatNames(aPaperFormats) {
    var aPaperFormatNames = new Array();
    for (var i in aPaperFormats) {
        aPaperFormatNames.push(aPaperFormats[i].sName);
    }
    return aPaperFormatNames;
}


// --------------------------------------------------------------------
// ----- END
// --------------------------------------------------------------------
