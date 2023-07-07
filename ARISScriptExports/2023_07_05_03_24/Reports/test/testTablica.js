//Context.setProperty("model-as-emf", true)
var oOutput = Context.createOutputObject()
var xlTableActive=false
oOutput.DefineF(getString("ID_STYLE_RD_DEFAULT"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT| Constants.FMT_VTOP, 0, 0, 0, 0, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_HEADING_1"), getString("ID_DEFAULT_FONT"), 18, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD | Constants.FMT_LEFT| Constants.FMT_VTOP| Constants.FMT_TOCENTRY0 , 0, 0, 4, 4, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_HEADING_2"), getString("ID_DEFAULT_FONT"), 14, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_ITALIC | Constants.FMT_BOLD | Constants.FMT_LEFT| Constants.FMT_VTOP| Constants.FMT_TOCENTRY1 , 0, 0, 2, 2, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_TABLE_CONTENT"), getString("ID_DEFAULT_FONT"), 8, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT| Constants.FMT_VTOP, 0, 0, 0, 0, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_HEADING_3"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_ITALIC | Constants.FMT_BOLD | Constants.FMT_LEFT| Constants.FMT_VTOP| Constants.FMT_TOCENTRY2 , 0, 0, 1, 1, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_HEADING_4"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_ITALIC | Constants.FMT_BOLD | Constants.FMT_LEFT| Constants.FMT_VTOP| Constants.FMT_TOCENTRY3 , 0, 0, 0, 0, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_INFO"), getString("ID_DEFAULT_FONT"), 14, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD | Constants.FMT_CENTER| Constants.FMT_VTOP, 0, 0, 1.76, 8.82, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_TABLE_HEAD"), getString("ID_DEFAULT_FONT"), 8, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD | Constants.FMT_CENTER| Constants.FMT_VTOP, 0, 0, 0, 0, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_TITLE"), getString("ID_DEFAULT_FONT"), 21, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD | Constants.FMT_CENTER| Constants.FMT_VTOP, 0, 0, 1.76, 8.82, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_HEADER_FOOTER"), getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT| Constants.FMT_VTOP, 0, 0, 0, 0, 0, 1)
setupOutputObject( oOutput )
oOutput.SetTitle(Context.getScriptInfo(Constants.SCRIPT_NAME))

var nLocale = Context.getSelectedLanguage()
createSection1(oOutput, ArisData.getSelectedGroups())
createSection2(oOutput, ArisData.getSelectedModels())
oOutput.WriteReport()

/** Check for excel format output
 */
function isExcel()
{
	if (Context.getSelectedFormat()==Constants.OutputXLS) return true;
	if (Context.getSelectedFormat()==Constants.OutputXLSX) return true;
	return false;
}
/** Apply default page format settings to output object
 * @param {Output} outputObj The output object
 */
function setupOutputObject(outputObj)
{
	outputObj.SetAutoTOCNumbering(true)
	globalHeader(outputObj)

	globalFooter(outputObj)
}
function globalHeader(outputObj) {
	outputObj.BeginHeader()
		if( !isExcel() )
			outputObj.BeginTable(100, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT,  Constants.FMT_LEFT | Constants.FMT_NOBORDER, 0)
			outputObj.TableRow()
				outputObj.ResetFrameStyle()
				outputObj.SetFrameStyle( Constants.FRAME_TOP, 0, 0)
				outputObj.SetFrameStyle( Constants.FRAME_LEFT, 0, 0)
				outputObj.SetFrameStyle( Constants.FRAME_BOTTOM, 0, 0)
				outputObj.SetFrameStyle( Constants.FRAME_RIGHT, 0, 0)
				outputObj.TableCell("", 50, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
			outputObj.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
			outputObj.OutputLn(Context.getScriptInfo(Constants.SCRIPT_NAME), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
			outputObj.EndParagraph()
				outputObj.TableCell("", 50, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_RIGHT, 0)
			outputObj.BeginParagraph( Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
			outputObj.OutputField(Constants.FIELD_FILENAME, getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT)
			outputObj.OutputLn("", getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT, 0.71)
			outputObj.EndParagraph()
		if( !isExcel() )
			outputObj.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
		else
			outputObj.TableRow()
	outputObj.EndHeader()
}
function globalFooter(outputObj) {
	outputObj.BeginFooter()
		if( !isExcel() )
			outputObj.BeginTable(100, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT,  Constants.FMT_LEFT | Constants.FMT_NOBORDER, 0)
			outputObj.TableRow()
				outputObj.ResetFrameStyle()
				outputObj.SetFrameStyle( Constants.FRAME_TOP, 0, 0)
				outputObj.SetFrameStyle( Constants.FRAME_LEFT, 0, 0)
				outputObj.SetFrameStyle( Constants.FRAME_BOTTOM, 0, 0)
				outputObj.SetFrameStyle( Constants.FRAME_RIGHT, 0, 0)
				outputObj.TableCell("", 50, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
			outputObj.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
			outputObj.OutputField(Constants.FIELD_DATE, getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT)
			outputObj.OutputLn("", getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
			outputObj.EndParagraph()
				outputObj.TableCell("", 50, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_RIGHT, 0)
			outputObj.BeginParagraph( Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
			outputObj.Output(getString("ID_REPORTDEF_2"), getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT, 0.71)
			outputObj.OutputField(Constants.FIELD_PAGE, getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT)
			outputObj.Output(getString("ID_REPORTDEF_3"), getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT, 0.71)
			outputObj.OutputField(Constants.FIELD_NUMPAGES, getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT)
			outputObj.OutputLn("", getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT, 0.71)
			outputObj.EndParagraph()
		if( !isExcel() )
			outputObj.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
		else
			outputObj.TableRow()
	outputObj.EndFooter()
}

/**
 * @param {Output} p_output The output object
 * @param {Group[]} p_aGroup 
*/
function createSection1(p_output, p_aGroup)
{
	// do not create new section if all data is empty
	if(p_aGroup.length==0)
		return

	//produce one table sheet per section
	if(xlTableActive && isExcel()) {
		p_output.EndTable("", 100, "", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
		xlTableActive=false
	}
	p_output.BeginSection(false, Constants.SECTION_DEFAULT)
	setupOutputObject( p_output ) //use defaults
	if( isExcel() ) {
		p_output.BeginTable(100, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
		xlTableActive=true
	}
	// repetition of queried data:
	for(var i=0; i<p_aGroup.length; i++) {
	}
	ArisData.getActiveDatabase().clearCaches()

	if( isExcel() ) {
		p_output.EndTable(getString("ID_REPORTDEF_4"), 100, "", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
		xlTableActive=false
	}
	p_output.EndSection()


}

/**
 * @param {Output} p_output The output object
 * @param {Model[]} p_aModel 
*/
function createSection2(p_output, p_aModel)
{
	// do not create new section if all data is empty
	if(p_aModel.length==0)
		return

	//produce one table sheet per section
	if(xlTableActive && isExcel()) {
		p_output.EndTable("", 100, "", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
		xlTableActive=false
	}
	p_output.BeginSection(false, Constants.SECTION_DEFAULT)
	setupOutputObject( p_output ) //use defaults
	if( isExcel() ) {
		p_output.BeginTable(100, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
		xlTableActive=true
	}
	// repetition of queried data:
	for(var i=0; i<p_aModel.length; i++) {
	}
	ArisData.getActiveDatabase().clearCaches()

	if( isExcel() ) {
		p_output.EndTable(getString("ID_REPORTDEF_5"), 100, "", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
		xlTableActive=false
	}
	p_output.EndSection()


}


function RGB(r, g, b) {
	return (new java.awt.Color(r/255.0,g/255.0,b/255.0,1)).getRGB() & 0xFFFFFF
}
