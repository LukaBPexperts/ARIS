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
createSection1(oOutput, ArisData.getSelectedObjDefs())
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
function createSection1(p_output, p_aObjDef)
{
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
	dtable(p_output, p_aObjDef)
	// createSection1 local functions:
	/**
	 * @param {Output} p_output The output object
	 * @param {ObjDef[]} p_aObjDef1 
	*/
	function dtable(p_output, p_aObjDef1)
	{
		// do not create new table if all data is empty
		if(p_aObjDef1.length==0)
			return
	
		if( !isExcel() )
			p_output.BeginTable(100, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
		p_output.TableRow()
			p_output.ResetFrameStyle()
			p_output.TableCell("", 50, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(204,204,204), 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
		p_output.OutputLn(getString("ID_REPORTDEF_7"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
			p_output.TableCell("", 50, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(204,204,204), 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
		p_output.OutputLn(getString("ID_REPORTDEF_6"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
		// repetition of queried data:
		for(var i=0; i<p_aObjDef1.length; i++) {
			p_output.TableRow()
				p_output.ResetFrameStyle()
				p_output.TableCell("", 50, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
			p_output.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
			writeData(p_output, p_aObjDef1[i].Name(nLocale))
			p_output.EndParagraph()
				p_output.TableCell("", 50, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
			p_output.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
			writeData1(p_output, p_aObjDef1[i].GUID())
			p_output.EndParagraph()
		}
		ArisData.getActiveDatabase().clearCaches()

		p_output.TableRow()
			p_output.ResetFrameStyle()
			p_output.TableCell("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
		if( !isExcel() )
			p_output.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
		else
			p_output.TableRow()

		// dtable local functions:
		/**
		 * @param {Output} p_output The output object
		 * @param {String } p_String
		*/
		function writeData(p_output, p_String) {
			p_output.OutputLn(p_String, getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
		}
		/**
		 * @param {Output} p_output The output object
		 * @param {String } p_String
		*/
		function writeData1(p_output, p_String) {
			p_output.OutputLn(p_String, getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
		}
	}
	
	if( isExcel() ) {
		p_output.EndTable(getString("ID_REPORTDEF_10"), 100, "", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
		xlTableActive=false
	}
	p_output.EndSection()


}


function RGB(r, g, b) {
	return (new java.awt.Color(r/255.0,g/255.0,b/255.0,1)).getRGB() & 0xFFFFFF
}
