var oOutput = Context.createOutputObject()
var xlTableActive=false
oOutput.DefineF(getString("ID_STYLE_RD_INFO"), getString("ID_DEFAULT_FONT"), 14, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD | Constants.FMT_CENTER| Constants.FMT_VTOP, 0, 0, 1.76, 8.82, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_TITLE"), getString("ID_DEFAULT_FONT"), 21, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD | Constants.FMT_CENTER| Constants.FMT_VTOP, 0, 0, 1.76, 8.82, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_DEFAULT"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT| Constants.FMT_VTOP, 0, 0, 0, 0, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_HEADING_4"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_ITALIC | Constants.FMT_BOLD | Constants.FMT_LEFT| Constants.FMT_VTOP| Constants.FMT_TOCENTRY3 , 0, 0, 0, 0, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_TABLE_CONTENT"), getString("ID_DEFAULT_FONT"), 8, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT| Constants.FMT_VTOP, 0, 0, 0, 0, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_HEADING_3"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_ITALIC | Constants.FMT_BOLD | Constants.FMT_LEFT| Constants.FMT_VTOP| Constants.FMT_TOCENTRY2 , 0, 0, 1, 1, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_HEADER_FOOTER"), getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT| Constants.FMT_VTOP, 0, 0, 0, 0, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_HEADING_2"), getString("ID_DEFAULT_FONT"), 14, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_ITALIC | Constants.FMT_BOLD | Constants.FMT_LEFT| Constants.FMT_VTOP| Constants.FMT_TOCENTRY1 , 0, 0, 2, 2, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_HEADING_1"), getString("ID_DEFAULT_FONT"), 18, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD | Constants.FMT_LEFT| Constants.FMT_VTOP| Constants.FMT_TOCENTRY0 , 0, 0, 4, 4, 0, 1)
oOutput.DefineF(getString("ID_STYLE_RD_TABLE_HEAD"), getString("ID_DEFAULT_FONT"), 8, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD | Constants.FMT_CENTER| Constants.FMT_VTOP, 0, 0, 0, 0, 0, 1)
setupOutputObject( oOutput )
oOutput.SetTitle(Context.getScriptInfo(Constants.SCRIPT_NAME))

var nLocale = Context.getSelectedLanguage()
function main(){
    createSection1(oOutput, ArisData.getSelectedDatabases())
    createSection2(oOutput, ArisData.getSelectedModels())
    oOutput.WriteReport()
}

/** Apply default page format settings to output object
 * @param {Output} outputObj The output object
 */
function setupOutputObject(outputObj)
{
	outputObj.SetPageWidth(210.10)
	outputObj.SetPageHeight(297.20)
	outputObj.SetLeftMargin(20)
	outputObj.SetRightMargin(20)
	outputObj.SetTopMargin(30)
	outputObj.SetBottomMargin(30)
	outputObj.SetDistHeader(10)
	outputObj.SetDistFooter(10)
	outputObj.SetAutoTOCNumbering(true)
	globalHeader(outputObj)

	globalFooter(outputObj)
}
function globalHeader(outputObj) {
	outputObj.BeginHeader()
		if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
			outputObj.BeginTable(100, RGB(255,255,255), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
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
		if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
			outputObj.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
		else
			outputObj.TableRow()
	outputObj.EndHeader()
}
function globalFooter(outputObj) {
	outputObj.BeginFooter()
		if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
			outputObj.BeginTable(100, RGB(255,255,255), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
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
			outputObj.Output(getString("ID_REPORTDEF_1"), getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT, 0.71)
			outputObj.OutputField(Constants.FIELD_PAGE, getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT)
			outputObj.Output(getString("ID_REPORTDEF_2"), getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT, 0.71)
			outputObj.OutputField(Constants.FIELD_NUMPAGES, getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT)
			outputObj.OutputLn("", getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT, 0.71)
			outputObj.EndParagraph()
		if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
			outputObj.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
		else
			outputObj.TableRow()
	outputObj.EndFooter()
}

/**
 * @param {Output} p_output The output object
 * @param {Database[]} p_aDatabase 
*/
function createSection1(p_output, p_aDatabase)
{
	//produce one table sheet per section
	if(xlTableActive && Context.getSelectedFormat()==Constants.OUTEXCEL) {
		p_output.EndTable("", 100, "", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
		xlTableActive=false
	}
	p_output.BeginSection(false, Constants.SECTION_INDEX)
	sectionPageSetup( p_output )

	if(Context.getSelectedFormat()==Constants.OUTEXCEL) {
		p_output.BeginTable(100, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
		xlTableActive=true
	}
	p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
		//to format the TOC output use output.SetTOCFormat(iLevel, sFont, iFontSize, nFontColor, nBackgroundColor, nFormat, nLeftIndentation, nRightIndentation, nDistTop, nDistBottom)
	p_output.OutputField( Constants.FIELD_TOC, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_LEFT)
	p_output.EndParagraph()
	if(Context.getSelectedFormat()==Constants.OUTEXCEL) {
		p_output.EndTable(getString("ID_REPORTDEF_7"), 100, "", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
		xlTableActive=false
	}
	p_output.EndSection()

	/** Apply default page format settings to output object
	 * @param {Output} outputObj The output object
	 */
	function sectionPageSetup(outputObj)
	{
		createHeader(outputObj)

		createFooter(outputObj)
		// local functions:
		function createHeader(outputObj) {
			outputObj.BeginHeader()
				if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
					outputObj.BeginTable(100, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
					outputObj.TableRow()
						outputObj.ResetFrameStyle()
						outputObj.TableCell("", 66.39, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
					outputObj.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
					writeData(outputObj, collect1(collect(ArisData.getSelectedModels())))
					outputObj.EndParagraph()
						outputObj.TableCell("", 33.61, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_CENTER, 0)
					outputObj.BeginParagraph( Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
										var image = Context.createPicture("Vestas logo.jpg")
					outputObj.OutGraphic(image, -1, 44, 10 )

					outputObj.EndParagraph()
				if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
					outputObj.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
				else
					outputObj.TableRow()
			outputObj.EndHeader()
				/**
				 * @param {Output} p_output The output object
				 * @param {Object[] } p_aObject
				*/
				function writeData(p_output, p_aObject) {
					for(var i=0; i<p_aObject.length; i++) {
						p_output.OutputLn(p_aObject[i], getString("ID_DEFAULT_FONT"), 13, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_LEFT, 0)
					}
				}
				function collect1(p_array){
					var result = new Array();
					for(var i=0; i<p_array.length; i++){
						result.push(p_array[i].GetValue(false))
					}
					return ArisData.Unique(result)
				}
				function collect(p_array){
					var result = new Array();
					for(var i=0; i<p_array.length; i++){
						result.push(p_array[i].Attribute(Constants.AT_NAME, nLocale))
					}
					return ArisData.Unique(result)
				}
		}
		function createFooter(outputObj) {
			outputObj.BeginFooter()
				if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
					outputObj.BeginTable(100, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
					outputObj.TableRow()
						outputObj.ResetFrameStyle()
						outputObj.TableCell("", 23.13, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
					outputObj.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
					outputObj.OutputLn(getString("ID_REPORTDEF_3"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
					outputObj.EndParagraph()
					outputObj.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
					outputObj.OutputField(Constants.FIELD_DATE, getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT)
					outputObj.Output(getString("ID_REPORTDEF_4"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
					outputObj.OutputField(Constants.FIELD_TIME, getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT)
					outputObj.OutputLn("", getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
					outputObj.EndParagraph()
						outputObj.TableCell("", 56.4, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
					outputObj.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
					outputObj.OutputLn(getString("ID_REPORTDEF_5"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
					outputObj.EndParagraph()
						outputObj.TableCell("", 20.83, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_RIGHT, 0)
					outputObj.BeginParagraph( Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
					outputObj.OutputLn(getString("ID_REPORTDEF_6"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT, 0.71)
					outputObj.EndParagraph()
				if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
					outputObj.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
				else
					outputObj.TableRow()
			outputObj.EndFooter()
		}
	}

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
	if(xlTableActive && Context.getSelectedFormat()==Constants.OUTEXCEL) {
		p_output.EndTable("", 100, "", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
		xlTableActive=false
	}
	p_output.BeginSection(false, Constants.SECTION_DEFAULT)
	setupOutputObject( p_output ) //use defaults
	if(Context.getSelectedFormat()==Constants.OUTEXCEL) {
		p_output.BeginTable(100, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
		xlTableActive=true
	}
	// repetition of queried data:
	for(var i=0; i<p_aModel.length; i++) {
		iteration_section(p_output, p_aModel[i].Attribute(Constants.AT_NAME, nLocale))
		iteration_section1(p_output, p_aModel[i].Attribute(Constants.AT_NAME, nLocale))
        var lObjOccs = Sort_SortModel(p_aModel[i], eSortType.Flow, eSortOrientation.YX, eSortNonStruct.Exclude, nLocale)
		iteration_section2(p_output, lObjOccs)
	}

	// createSection2 local functions:
		/**
		 * @param {Output} p_output The output object
		 * @param {Attr} p_Attr 
		*/
		function iteration_section(p_output, p_Attr)
		{
			//produce one table sheet per section
			if(xlTableActive && Context.getSelectedFormat()==Constants.OUTEXCEL) {
				p_output.EndTable(getString("ID_REPORTDEF_35"), 100, "", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
				xlTableActive=false
			}
			p_output.BeginSection(false, Constants.SECTION_DEFAULT)
			sectionPageSetup( p_output )

			if(Context.getSelectedFormat()==Constants.OUTEXCEL) {
				p_output.BeginTable(100, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
				xlTableActive=true
			}
			p_output.BeginParagraph( Constants.FMT_LEFT, 0, 0, 0, 0, 0)
			p_output.OutputLn(getString("ID_REPORTDEF_8"), getString("ID_DEFAULT_FONT"), 14, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_LEFT, 0)
			p_output.EndParagraph()
			p_output.BeginParagraph( Constants.FMT_LEFT, 0, 0, 0, 0, 0)
			writeData2(p_output, p_aModel[i].Graphic(false, false, nLocale))
			p_output.EndParagraph()
			// iteration_section local functions:
		/**
		 * @param {Output} p_output The output object
		 * @param {IModelPicture } p_IModelPicture
		*/
		function writeData2(p_output, p_IModelPicture) {
			p_output.OutGraphic(p_IModelPicture, -1, p_output.getCurrentMaxWidth(), p_output.GetPageHeight() - p_output.GetTopMargin() - p_output.GetBottomMargin() -10)
		}
			if(Context.getSelectedFormat()==Constants.OUTEXCEL) {
				p_output.EndTable("Model graphic", 100, "", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
				xlTableActive=false
			}
			p_output.EndSection()
		
					/** Apply default page format settings to output object
			 * @param {Output} outputObj The output object
			 */
			function sectionPageSetup(outputObj)
			{
				createHeader(outputObj)

				createFooter(outputObj)
				// local functions:
				function createHeader(outputObj) {
					outputObj.BeginHeader()
						if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
							outputObj.BeginTable(100, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
							outputObj.TableRow()
								outputObj.ResetFrameStyle()
								outputObj.TableCell("", 66.60, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_CENTER, 0)
							outputObj.BeginParagraph( Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
							writeData1(outputObj, p_aModel[i].Name(nLocale))
							outputObj.EndParagraph()
								outputObj.TableCell("", 33.40, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_RIGHT, 0)
							outputObj.BeginParagraph( Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
														var image = Context.createPicture("vestas.png")
							outputObj.OutGraphic(image, -1, 51, 11 )

							outputObj.EndParagraph()
						if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
							outputObj.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
						else
							outputObj.TableRow()
					outputObj.EndHeader()
						/**
						 * @param {Output} p_output The output object
						 * @param {String } p_String
						*/
						function writeData1(p_output, p_String) {
							p_output.OutputLn(p_String, getString("ID_DEFAULT_FONT"), 13, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_CENTER, 0.71)
						}
				}
				function createFooter(outputObj) {
					outputObj.BeginFooter()
						if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
							outputObj.BeginTable(100, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
							outputObj.TableRow()
								outputObj.ResetFrameStyle()
								outputObj.TableCell("", 21, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
							outputObj.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
							outputObj.OutputLn(getString("ID_REPORTDEF_3"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
							outputObj.EndParagraph()
							outputObj.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
							outputObj.OutputField(Constants.FIELD_DATE, getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT)
							outputObj.Output(getString("ID_REPORTDEF_4"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
							outputObj.OutputField(Constants.FIELD_TIME, getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT)
							outputObj.OutputLn("", getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
							outputObj.EndParagraph()
								outputObj.TableCell("", 58.63, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
							outputObj.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
							outputObj.OutputLn(getString("ID_REPORTDEF_5"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
							outputObj.EndParagraph()
								outputObj.TableCell("", 20.37, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_RIGHT, 0)
							outputObj.BeginParagraph( Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
							outputObj.Output(getString("ID_REPORTDEF_1"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT, 0.71)
							outputObj.OutputField(Constants.FIELD_PAGE, getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT)
							outputObj.Output(getString("ID_REPORTDEF_2"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT, 0.71)
							outputObj.OutputField(Constants.FIELD_NUMPAGES, getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT)
							outputObj.OutputLn("", getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT, 0.71)
							outputObj.EndParagraph()
						if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
							outputObj.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
						else
							outputObj.TableRow()
					outputObj.EndFooter()
				}
			}

		}
		
		/**
		 * @param {Output} p_output The output object
		 * @param {Attr} p_Attr 
		*/
		function iteration_section1(p_output, p_Attr)
		{
			//produce one table sheet per section
			if(xlTableActive && Context.getSelectedFormat()==Constants.OUTEXCEL) {
				p_output.EndTable(getString("ID_REPORTDEF_35"), 100, "", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
				xlTableActive=false
			}
			p_output.BeginSection(false, Constants.SECTION_DEFAULT)
			sectionPageSetup( p_output )

			if(Context.getSelectedFormat()==Constants.OUTEXCEL) {
				p_output.BeginTable(100, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
				xlTableActive=true
			}
			p_output.BeginParagraph( Constants.FMT_LEFT, 0, 0, 0, 0, 0)
			p_output.OutputLn(getString("ID_REPORTDEF_9"), getString("ID_DEFAULT_FONT"), 14, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_LEFT, 0)
			p_output.EndParagraph()
			if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
				p_output.BeginTable(100, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
				p_output.TableRow()
					p_output.ResetFrameStyle()
					p_output.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
				p_output.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
				p_output.Output(getString("ID_REPORTDEF_10"), getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
				writeData4(p_output, p_aModel[i].Attribute(Constants.AT_AUTH, nLocale))
				p_output.EndParagraph()
					p_output.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
					p_output.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
				p_output.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
				p_output.Output(getString("ID_REPORTDEF_11"), getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
				writeData5(p_output, p_aModel[i].Attribute(Constants.AT_VALID_FROM, nLocale))
				p_output.EndParagraph()
					p_output.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
				p_output.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
				p_output.Output(getString("ID_REPORTDEF_12"), getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
				writeData6(p_output, p_aModel[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("db3b0fc0-5fd4-11de-4732-00237d347a8a"), nLocale))
				p_output.EndParagraph()
				p_output.TableRow()
					p_output.ResetFrameStyle()
					p_output.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
				p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
				p_output.Output(getString("ID_REPORTDEF_13"), getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
				writeData7(p_output, p_aModel[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("618d4120-ea3f-11de-4732-00237d347a8a"), nLocale))
				p_output.EndParagraph()
					p_output.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
				p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
				p_output.Output(getString("ID_REPORTDEF_14"), getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
				writeData8(p_output, p_aModel[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("e1a55420-e570-11de-4732-00237d347a8a"), nLocale))
				p_output.EndParagraph()
					p_output.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
				//p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
				//p_output.Output(getString("ID_REPORTDEF_15"), getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
				//writeData9(p_output, p_aModel[i].Attribute(Constants.AT_VALID_UNTIL, nLocale))
				//p_output.EndParagraph()
					p_output.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
				p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
				p_output.Output(getString("ID_REPORTDEF_16"), getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
				writeData10(p_output, p_aModel[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("8fda52b0-6188-11de-4732-00237d347a8a"), nLocale))
				p_output.EndParagraph()
			if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
				p_output.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
			else
				p_output.TableRow()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			p_output.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			p_output.OutputLn(getString("ID_REPORTDEF_17"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_LEFT, 0)
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData11(p_output, p_aModel[i].Attribute(Constants.AT_NAME, nLocale))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			p_output.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			p_output.OutputLn(getString("ID_REPORTDEF_18"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_LEFT, 0)
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData12(p_output, p_aModel[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("c4648a80-a28f-11de-4732-00237d347a8a"), nLocale))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			p_output.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			p_output.OutputLn(getString("ID_REPORTDEF_19"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_LEFT, 0)
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData13(p_output, p_aModel[i].Attribute(Constants.AT_DESC, nLocale))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			p_output.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			p_output.OutputLn(getString("ID_REPORTDEF_20"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_LEFT, 0)
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData14(p_output, p_aModel[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("cb2e9410-ecd9-11df-21f4-00237d347a8a"), nLocale))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			p_output.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			p_output.OutputLn(getString("ID_REPORTDEF_21"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_LEFT, 0)
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData15(p_output, p_aModel[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("e8a5da20-a28f-11de-4732-00237d347a8a"), nLocale))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			p_output.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			p_output.OutputLn(getString("ID_REPORTDEF_22"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_LEFT, 0)
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData16(p_output, p_aModel[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("4a526e00-a290-11de-4732-00237d347a8a"), nLocale))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			p_output.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			p_output.OutputLn(getString("ID_REPORTDEF_23"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_LEFT, 0)
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData17(p_output, p_aModel[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("59d18ff0-a290-11de-4732-00237d347a8a"), nLocale))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			p_output.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			p_output.OutputLn(getString("ID_REPORTDEF_24"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_LEFT, 0)
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData18(p_output, p_aModel[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("ad4098b0-3021-11df-4732-00237d347a8a"), nLocale))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			p_output.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			p_output.OutputLn(getString("ID_REPORTDEF_25"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_LEFT, 0)
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData19(p_output, p_aModel[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("71fc3d00-a290-11de-4732-00237d347a8a"), nLocale))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			p_output.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			p_output.OutputLn(getString("ID_REPORTDEF_26"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_LEFT, 0)
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData20(p_output, p_aModel[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("d2a555a0-c334-11de-4732-00237d347a8a"), nLocale))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			p_output.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
			p_output.EndParagraph()
			p_output.BeginParagraph( Constants.FMT_LEFT, 0, 0, 0, 0, 0)
			p_output.OutputLn(getString("ID_REPORTDEF_27"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_LEFT, 0)
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData21(p_output, p_aModel[i].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("16b5b110-a290-11de-4732-00237d347a8a"), nLocale))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			p_output.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			p_output.OutputLn(getString("ID_REPORTDEF_28"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_LEFT, 0)
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData22(p_output, p_aModel[i].Attribute(Constants.AT_TITL1, nLocale).GetValue(false))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData23(p_output, p_aModel[i].Attribute(Constants.AT_EXT_1, nLocale).GetValue(false))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData24(p_output, p_aModel[i].Attribute(Constants.AT_TITL2, nLocale).GetValue(false))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData25(p_output, p_aModel[i].Attribute(Constants.AT_EXT_2, nLocale).GetValue(false))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData26(p_output, p_aModel[i].Attribute(Constants.AT_TITL3, nLocale).GetValue(false))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData27(p_output, p_aModel[i].Attribute(Constants.AT_EXT_3, nLocale).GetValue(false))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData28(p_output, p_aModel[i].Attribute(Constants.AT_TITL4, nLocale).GetValue(false))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData29(p_output, p_aModel[i].Attribute(Constants.AT_LINK, nLocale).GetValue(false))
			p_output.EndParagraph()
			// iteration_section1 local functions:
			/**
			 * @param {Output} p_output The output object
			 * @param {Attr } p_Attr
			*/
			function writeData4(p_output, p_Attr) {
				p_output.OutputLn(p_Attr.getValue(), getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
			}
			/**
			 * @param {Output} p_output The output object
			 * @param {Attr } p_Attr
			*/
			function writeData5(p_output, p_Attr) {
				p_output.OutputLn(p_Attr.getValue(), getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
			}
			/**
			 * @param {Output} p_output The output object
			 * @param {Attr } p_Attr
			*/
			function writeData6(p_output, p_Attr) {
				p_output.OutputLn(p_Attr.getValue(), getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
			}
			/**
			 * @param {Output} p_output The output object
			 * @param {Attr } p_Attr
			*/
			function writeData7(p_output, p_Attr) {
				p_output.OutputLn(p_Attr.getValue(), getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
			}
			/**
			 * @param {Output} p_output The output object
			 * @param {Attr } p_Attr
			*/
			function writeData8(p_output, p_Attr) {
				p_output.OutputLn(p_Attr.getValue(), getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
			}
			/**
			 * @param {Output} p_output The output object
			 * @param {Attr } p_Attr
			*/
			function writeData9(p_output, p_Attr) {
				p_output.OutputLn(p_Attr.getValue(), getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
			}
			/**
			 * @param {Output} p_output The output object
			 * @param {Attr } p_Attr
			*/
			function writeData10(p_output, p_Attr) {
				p_output.OutputLn(p_Attr.getValue(), getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
			}
		/**
		 * @param {Output} p_output The output object
		 * @param {Attr } p_Attr
		*/
		function writeData11(p_output, p_Attr) {
			p_output.OutputLn(p_Attr.getValue(), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
		}
		/**
		 * @param {Output} p_output The output object
		 * @param {Attr } p_Attr
		*/
		function writeData12(p_output, p_Attr) {
			p_output.OutputLn(p_Attr.getValue(), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
		}
		/**
		 * @param {Output} p_output The output object
		 * @param {Attr } p_Attr
		*/
		function writeData13(p_output, p_Attr) {
			p_output.OutputLn(p_Attr.getValue(), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
		}
		/**
		 * @param {Output} p_output The output object
		 * @param {Attr } p_Attr
		*/
		function writeData14(p_output, p_Attr) {
			p_output.OutputLn(p_Attr.getValue(), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
		}
		/**
		 * @param {Output} p_output The output object
		 * @param {Attr } p_Attr
		*/
		function writeData15(p_output, p_Attr) {
			p_output.OutputLn(p_Attr.getValue(), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
		}
		/**
		 * @param {Output} p_output The output object
		 * @param {Attr } p_Attr
		*/
		function writeData16(p_output, p_Attr) {
			p_output.OutputLn(p_Attr.getValue(), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
		}
		/**
		 * @param {Output} p_output The output object
		 * @param {Attr } p_Attr
		*/
		function writeData17(p_output, p_Attr) {
			p_output.OutputLn(p_Attr.getValue(), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
		}
		/**
		 * @param {Output} p_output The output object
		 * @param {Attr } p_Attr
		*/
		function writeData18(p_output, p_Attr) {
			p_output.OutputLn(p_Attr.getValue(), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
		}
		/**
		 * @param {Output} p_output The output object
		 * @param {Attr } p_Attr
		*/
		function writeData19(p_output, p_Attr) {
			p_output.OutputLn(p_Attr.getValue(), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
		}
		/**
		 * @param {Output} p_output The output object
		 * @param {Attr } p_Attr
		*/
		function writeData20(p_output, p_Attr) {
			p_output.OutputLn(p_Attr.getValue(), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
		}
		/**
		 * @param {Output} p_output The output object
		 * @param {Attr } p_Attr
		*/
		function writeData21(p_output, p_Attr) {
			p_output.OutputLn(p_Attr.getValue(), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
		}
		/**
		 * @param {Output} p_output The output object
		 * @param {Object } p_Object
		*/
		function writeData22(p_output, p_Object) {
			p_output.OutputLn(p_Object, getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
		}
		/**
		 * @param {Output} p_output The output object
		 * @param {Object } p_Object
		*/
		function writeData23(p_output, p_Object) {
			p_output.OutputLn(p_Object, getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
		}
		/**
		 * @param {Output} p_output The output object
		 * @param {Object } p_Object
		*/
		function writeData24(p_output, p_Object) {
			p_output.OutputLn(p_Object, getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
		}
		/**
		 * @param {Output} p_output The output object
		 * @param {Object } p_Object
		*/
		function writeData25(p_output, p_Object) {
			p_output.OutputLn(p_Object, getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
		}
		/**
		 * @param {Output} p_output The output object
		 * @param {Object } p_Object
		*/
		function writeData26(p_output, p_Object) {
			p_output.OutputLn(p_Object, getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
		}
		/**
		 * @param {Output} p_output The output object
		 * @param {Object } p_Object
		*/
		function writeData27(p_output, p_Object) {
			p_output.OutputLn(p_Object, getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
		}
		/**
		 * @param {Output} p_output The output object
		 * @param {Object } p_Object
		*/
		function writeData28(p_output, p_Object) {
			p_output.OutputLn(p_Object, getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
		}
		/**
		 * @param {Output} p_output The output object
		 * @param {Object } p_Object
		*/
		function writeData29(p_output, p_Object) {
			p_output.OutputLn(p_Object, getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
		}
			if(Context.getSelectedFormat()==Constants.OUTEXCEL) {
				p_output.EndTable("Model attributes", 100, "", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
				xlTableActive=false
			}
			p_output.EndSection()
		
					/** Apply default page format settings to output object
			 * @param {Output} outputObj The output object
			 */
			function sectionPageSetup(outputObj)
			{
				createHeader(outputObj)

				createFooter(outputObj)
				// local functions:
				function createHeader(outputObj) {
					outputObj.BeginHeader()
						if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
							outputObj.BeginTable(100, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
							outputObj.TableRow()
								outputObj.ResetFrameStyle()
								outputObj.TableCell("", 67.1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_CENTER, 0)
							outputObj.BeginParagraph( Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
							writeData3(outputObj, p_aModel[i].Name(nLocale))
							outputObj.EndParagraph()
								outputObj.TableCell("", 32.99, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_RIGHT, 0)
							outputObj.BeginParagraph( Constants.FMT_RIGHT, 1, 1, 0, 0, 0)
														var image = Context.createPicture("vestas.png")
							outputObj.OutGraphic(image, -1, 51, 11 )

							outputObj.EndParagraph()
						if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
							outputObj.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
						else
							outputObj.TableRow()
					outputObj.EndHeader()
						/**
						 * @param {Output} p_output The output object
						 * @param {String } p_String
						*/
						function writeData3(p_output, p_String) {
							p_output.OutputLn(p_String, getString("ID_DEFAULT_FONT"), 13, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_CENTER, 0.71)
						}
				}
				function createFooter(outputObj) {
					outputObj.BeginFooter()
						if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
							outputObj.BeginTable(100, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
							outputObj.TableRow()
								outputObj.ResetFrameStyle()
								outputObj.TableCell("", 21, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
							outputObj.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
							outputObj.OutputLn(getString("ID_REPORTDEF_3"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
							outputObj.EndParagraph()
							outputObj.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
							outputObj.OutputField(Constants.FIELD_DATE, getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT)
							outputObj.Output(getString("ID_REPORTDEF_4"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
							outputObj.OutputField(Constants.FIELD_TIME, getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT)
							outputObj.OutputLn("", getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
							outputObj.EndParagraph()
								outputObj.TableCell("", 58.21, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
							outputObj.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
							outputObj.OutputLn(getString("ID_REPORTDEF_5"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
							outputObj.EndParagraph()
								outputObj.TableCell("", 20.79, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_RIGHT, 0)
							outputObj.BeginParagraph( Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
							outputObj.Output(getString("ID_REPORTDEF_1"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT, 0.71)
							outputObj.OutputField(Constants.FIELD_PAGE, getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT)
							outputObj.Output(getString("ID_REPORTDEF_2"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT, 0.71)
							outputObj.OutputField(Constants.FIELD_NUMPAGES, getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT)
							outputObj.OutputLn("", getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT, 0.71)
							outputObj.EndParagraph()
						if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
							outputObj.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
						else
							outputObj.TableRow()
					outputObj.EndFooter()
				}
			}

		}
		
		/**
		 * @param {Output} p_output The output object
		 * @param {ObjOcc[]} p_aObjOcc 
		*/
		function iteration_section2(p_output, pObjOccs)
		{
			var p_aObjOcc = new Array()
            for (var ii in pObjOccs){
                if (pObjOccs[ii].ObjDef().TypeNum()==Constants.OT_FUNC){
                    p_aObjOcc.push(pObjOccs[ii])
                }
            }
            // do not create new section if all data is empty
			if(p_aObjOcc.length==0)
				return
		
			//produce one table sheet per section
			if(xlTableActive && Context.getSelectedFormat()==Constants.OUTEXCEL) {
				p_output.EndTable(getString("ID_REPORTDEF_35"), 100, "", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
				xlTableActive=false
			}
			p_output.BeginSection(false, Constants.SECTION_DEFAULT)
			sectionPageSetup( p_output )

			if(Context.getSelectedFormat()==Constants.OUTEXCEL) {
				p_output.BeginTable(100, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
				xlTableActive=true
			}
			p_output.BeginParagraph( Constants.FMT_LEFT, 0, 0, 0, 0, 0)
			p_output.OutputLn(getString("ID_REPORTDEF_29"), getString("ID_DEFAULT_FONT"), 14, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_LEFT, 0)
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			p_output.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
			p_output.EndParagraph()
			// repetition of queried data:
			for(var i1=0; i1<p_aObjOcc.length; i1++) {
				p_output.BeginParagraph( Constants.FMT_LEFT, 0, 0, 0, 0, 0)
				writeData31a(p_output, p_aObjOcc[i1].getObjDefName(nLocale))
				p_output.EndParagraph()
				p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
				writeData32(p_output, p_aObjOcc[i1].ObjDef().Attribute(Constants.AT_DESC, nLocale))
				p_output.EndParagraph()
				dtable(p_output, p_aObjOcc[i1].ObjDef().CxnListFilter(Constants.EDGES_IN, new Array(Constants.CT_EXEC_1,Constants.CT_EXEC_2,Constants.CT_IS_RESP_1,Constants.CT_IS_RESP_2,Constants.CT_IS_RESP_FOR_1,Constants.CT_IS_RESP_FOR_2)))
				dtable1(p_output, collect1(collect(p_aObjOcc[i1].ObjDef().CxnListFilter(Constants.EDGES_OUT, [Constants.CT_HAS_OUT]))))
				//dtable2(p_output, collect3(collect2(ArisData.sort(p_aObjOcc[i1].ObjDef().CxnListFilter(Constants.EDGES_IN, new Array(Constants.CT_BPEL_IS_INPUT,Constants.CT_IS_INP_FOR)),Constants.AT_NAME, nLocale))))
                dtable2(p_output, collect3(collect2(ArisData.sort(p_aObjOcc[i1].ObjDef().CxnListFilter(Constants.EDGES_IN, Constants.CT_IS_INP_FOR),Constants.AT_NAME, nLocale))))
				dtable3(p_output, collect5(collect4(p_aObjOcc[i1].ObjDef().CxnListFilter(Constants.EDGES_IN, new Array(Constants.CT_IS_SUPPORTED_BY,Constants.CT_SUPP_1,Constants.CT_SUPP_2,Constants.CT_SUPP_3)))))
				dtable4(p_output, collect7(collect6(p_aObjOcc[i1].ObjDef().CxnListFilter(Constants.EDGES_OUT, new Array(Constants.CT_CREATES,Constants.CT_CRT_1,Constants.CT_CRT_2,Constants.CT_CRT_3,Constants.CT_CRT_4,Constants.CT_CRT_5,Constants.CT_CRT_6)))))
				dtable5(p_output, p_aObjOcc[i1].ObjDef().AttrList(nLocale, new Array(Constants.AT_EXT_1,Constants.AT_EXT_2,Constants.AT_TITL2,Constants.AT_EXT_3,Constants.AT_TITL1,Constants.AT_LINK,Constants.AT_TITL3,Constants.AT_TITL4)))
				p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
				p_output.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
				p_output.EndParagraph()
				p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
				p_output.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
				p_output.EndParagraph()
			}

			// iteration_section2 local functions:
			/**
			 * @param {Output} p_output The output object
			 * @param {String } p_String
			*/
            function writeData31a(p_output, p_String) {
				p_output.OutputLn(p_String, getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_LEFT, 0)
			}
			function writeData31(p_output, p_String) {
				p_output.OutputLn(p_String, getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_LEFT, 0)
			}
			/**
			 * @param {Output} p_output The output object
			 * @param {Attr } p_Attr
			*/
			function writeData32(p_output, p_Attr) {
				p_output.OutputLn(p_Attr.getValue(), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
			}
				/**
				 * @param {Output} p_output The output object
				 * @param {CxnDef[]} p_aCxnDef 
				*/
				function dtable(p_output, p_aCxnDef)
				{
					// do not create new table if all data is empty
					if(p_aCxnDef.length==0)
						return
				
					if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
						p_output.BeginTable(100, RGB(0,0,0), RGB(255,255,255),  Constants.FMT_LEFT, 0)
					// repetition of queried data:
					for(var i2=0; i2<p_aCxnDef.length; i2++) {
						p_output.TableRow()
							p_output.ResetFrameStyle()
							p_output.SetFrameStyle( Constants.FRAME_TOP, 0, 0)
							p_output.SetFrameStyle( Constants.FRAME_LEFT, 0, 0)
							p_output.SetFrameStyle( Constants.FRAME_BOTTOM, 0, 0)
							p_output.SetFrameStyle( Constants.FRAME_RIGHT, 0, 0)
							p_output.TableCell("", 50.62, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255,255,255), 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
						p_output.BeginParagraph( Constants.FMT_LEFT, 12.70, 0, 0, 0, 0)
						p_output.OutputLn(getString("ID_REPORTDEF_30"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 12.70)
						p_output.EndParagraph()
							p_output.TableCell("", 49.38, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255,255,255), 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
						p_output.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
						writeData33(p_output, p_aCxnDef[i2].SourceObjDef().Name(nLocale))
						p_output.EndParagraph()
					}

					if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
						p_output.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
					else
						p_output.TableRow()

					// dtable local functions:
					/**
					 * @param {Output} p_output The output object
					 * @param {String } p_String
					*/
					function writeData33(p_output, p_String) {
						p_output.OutputLn(p_String, getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
					}
				}
				
				/**
				 * @param {Output} p_output The output object
				 * @param {Attr[]} p_aAttr 
				*/
				function dtable1(p_output, p_aAttr)
				{
					// do not create new table if all data is empty
					if(p_aAttr.length==0)
						return
				
					if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
						p_output.BeginTable(100, RGB(255,255,255), RGB(255,255,255),  Constants.FMT_LEFT, 0)
					// repetition of queried data:
					for(var i2=0; i2<p_aAttr.length; i2++) {
						p_output.TableRow()
							p_output.ResetFrameStyle()
							p_output.SetFrameStyle( Constants.FRAME_TOP, 0, 0)
							p_output.SetFrameStyle( Constants.FRAME_LEFT, 0, 0)
							p_output.SetFrameStyle( Constants.FRAME_BOTTOM, 0, 0)
							p_output.SetFrameStyle( Constants.FRAME_RIGHT, 0, 0)
							p_output.TableCell("", 50.62, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255,255,255), 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
						p_output.BeginParagraph( Constants.FMT_LEFT, 13.41, 0.71, 0, 0, 0)
						p_output.OutputLn(getString("ID_REPORTDEF_31"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 13.41)
						p_output.EndParagraph()
							p_output.TableCell("", 49.38, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255,255,255), 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
						p_output.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
						writeData34(p_output, p_aAttr[i2].GetValue(false))
						p_output.EndParagraph()
					}

					if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
						p_output.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
					else
						p_output.TableRow()

					// dtable1 local functions:
					/**
					 * @param {Output} p_output The output object
					 * @param {Object } p_Object
					*/
					function writeData34(p_output, p_Object) {
						p_output.OutputLn(p_Object, getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
					}
				}
				
				function collect1(p_array){
					var result = new Array();
					for(var i=0; i<p_array.length; i++){
						result.push(p_array[i].Attribute(Constants.AT_NAME, nLocale))
					}
					return ArisData.Unique(result)
				}
				function collect(p_array){
					var result = new Array();
					for(var i=0; i<p_array.length; i++){
						result.push(p_array[i].TargetObjDef())
					}
					return ArisData.Unique(result)
				}
				/**
				 * @param {Output} p_output The output object
				 * @param {Attr[]} p_aAttr 
				*/
				function dtable2(p_output, p_aAttr)
				{
					// do not create new table if all data is empty
					if(p_aAttr.length==0)
						return
				
					if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
						p_output.BeginTable(100, RGB(255,255,255), RGB(255,255,255),  Constants.FMT_LEFT, 0)
					// repetition of queried data:
					for(var i2=0; i2<p_aAttr.length; i2++) {
						p_output.TableRow()
							p_output.ResetFrameStyle()
							p_output.SetFrameStyle( Constants.FRAME_TOP, 0, 0)
							p_output.SetFrameStyle( Constants.FRAME_LEFT, 0, 0)
							p_output.SetFrameStyle( Constants.FRAME_BOTTOM, 0, 0)
							p_output.SetFrameStyle( Constants.FRAME_RIGHT, 0, 0)
							p_output.TableCell("", 50.62, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255,255,255), 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
						p_output.BeginParagraph( Constants.FMT_LEFT, 13.41, 0.71, 0, 0, 0)
						p_output.OutputLn(getString("ID_REPORTDEF_32"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 13.41)
						p_output.EndParagraph()
							p_output.TableCell("", 49.38, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255,255,255), 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
						p_output.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
						writeData35(p_output, p_aAttr[i2].GetValue(false))
						p_output.EndParagraph()
					}

					if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
						p_output.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
					else
						p_output.TableRow()

					// dtable2 local functions:
					/**
					 * @param {Output} p_output The output object
					 * @param {Object } p_Object
					*/
					function writeData35(p_output, p_Object) {
						p_output.OutputLn(p_Object, getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
					}
				}
				
				function collect2(p_array){
					var result = new Array();
					for(var i=0; i<p_array.length; i++){
						result.push(p_array[i].SourceObjDef())
					}
					return ArisData.Unique(result)
				}
				function collect3(p_array){
					var result = new Array();
					for(var i=0; i<p_array.length; i++){
						result.push(p_array[i].Attribute(Constants.AT_NAME, nLocale))
					}
					return ArisData.Unique(result)
				}
				/**
				 * @param {Output} p_output The output object
				 * @param {Attr[]} p_aAttr 
				*/
				function dtable3(p_output, p_aAttr)
				{
					// do not create new table if all data is empty
					if(p_aAttr.length==0)
						return
				
					if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
						p_output.BeginTable(100, RGB(255,255,255), RGB(255,255,255),  Constants.FMT_LEFT, 0)
					// repetition of queried data:
					for(var i2=0; i2<p_aAttr.length; i2++) {
						p_output.TableRow()
							p_output.ResetFrameStyle()
							p_output.SetFrameStyle( Constants.FRAME_TOP, 0, 0)
							p_output.SetFrameStyle( Constants.FRAME_LEFT, 0, 0)
							p_output.SetFrameStyle( Constants.FRAME_BOTTOM, 0, 0)
							p_output.SetFrameStyle( Constants.FRAME_RIGHT, 0, 0)
							p_output.TableCell("", 50.62, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255,255,255), 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
						p_output.BeginParagraph( Constants.FMT_LEFT, 13.41, 0.71, 0, 0, 0)
						p_output.OutputLn(getString("ID_REPORTDEF_33"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 13.41)
						p_output.EndParagraph()
							p_output.TableCell("", 49.38, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255,255,255), 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
						p_output.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
						writeData36(p_output, p_aAttr[i2].GetValue(false))
						p_output.EndParagraph()
					}

					if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
						p_output.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
					else
						p_output.TableRow()

					// dtable3 local functions:
					/**
					 * @param {Output} p_output The output object
					 * @param {Object } p_Object
					*/
					function writeData36(p_output, p_Object) {
						p_output.OutputLn(p_Object, getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
					}
				}
				
				function collect4(p_array){
					var result = new Array();
					for(var i=0; i<p_array.length; i++){
						result.push(p_array[i].SourceObjDef())
					}
					return ArisData.Unique(result)
				}
				function collect5(p_array){
					var result = new Array();
					for(var i=0; i<p_array.length; i++){
						result.push(p_array[i].Attribute(Constants.AT_NAME, nLocale))
					}
					return ArisData.Unique(result)
				}
				/**
				 * @param {Output} p_output The output object
				 * @param {Attr[]} p_aAttr 
				*/
				function dtable4(p_output, p_aAttr)
				{
					// do not create new table if all data is empty
					if(p_aAttr.length==0)
						return
				
					if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
						p_output.BeginTable(100, RGB(255,255,255), RGB(255,255,255),  Constants.FMT_LEFT, 0)
					// repetition of queried data:
					for(var i2=0; i2<p_aAttr.length; i2++) {
						p_output.TableRow()
							p_output.ResetFrameStyle()
							p_output.SetFrameStyle( Constants.FRAME_TOP, 0, 0)
							p_output.SetFrameStyle( Constants.FRAME_LEFT, 0, 0)
							p_output.SetFrameStyle( Constants.FRAME_BOTTOM, 0, 0)
							p_output.SetFrameStyle( Constants.FRAME_RIGHT, 0, 0)
							p_output.TableCell("", 50.62, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255,255,255), 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
						p_output.BeginParagraph( Constants.FMT_LEFT, 13.41, 0.71, 0, 0, 0)
						p_output.OutputLn(getString("ID_REPORTDEF_34"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 13.41)
						p_output.EndParagraph()
							p_output.TableCell("", 49.38, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255,255,255), 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
						p_output.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
						writeData37(p_output, p_aAttr[i2].GetValue(false))
						p_output.EndParagraph()
					}

					if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
						p_output.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
					else
						p_output.TableRow()

					// dtable4 local functions:
					/**
					 * @param {Output} p_output The output object
					 * @param {Object } p_Object
					*/
					function writeData37(p_output, p_Object) {
						p_output.OutputLn(p_Object, getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
					}
				}
				
				function collect6(p_array){
					var result = new Array();
					for(var i=0; i<p_array.length; i++){
						result.push(p_array[i].TargetObjDef())
					}
					return ArisData.Unique(result)
				}
				function collect7(p_array){
					var result = new Array();
					for(var i=0; i<p_array.length; i++){
						result.push(p_array[i].Attribute(Constants.AT_NAME, nLocale))
					}
					return ArisData.Unique(result)
				}
				/**
				 * @param {Output} p_output The output object
				 * @param {Attr[]} p_aAttr 
				*/
				function dtable5(p_output, p_aAttr)
				{
					// do not create new table if all data is empty
					if(p_aAttr.length==0)
						return
				
					if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
						p_output.BeginTable(100, RGB(255,255,255), RGB(255,255,255),  Constants.FMT_LEFT, 0)
					p_output.TableRow()
						p_output.ResetFrameStyle()
						p_output.SetFrameStyle( Constants.FRAME_TOP, 0, 0)
						p_output.SetFrameStyle( Constants.FRAME_LEFT, 0, 0)
						p_output.SetFrameStyle( Constants.FRAME_BOTTOM, 0, 0)
						p_output.SetFrameStyle( Constants.FRAME_RIGHT, 0, 0)
						p_output.TableCell("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255,255,255), 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
					p_output.BeginParagraph( Constants.FMT_LEFT, 0, 0, 0, 0, 0)
					p_output.OutputLn(getString("ID_REPORTDEF_28"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_UNDERLINE  |  Constants.FMT_LEFT, 0)
					p_output.EndParagraph()
					// repetition of queried data:
					for(var i2=0; i2<p_aAttr.length; i2++) {
						p_output.TableRow()
							p_output.ResetFrameStyle()
							p_output.SetFrameStyle( Constants.FRAME_TOP, 0, 0)
							p_output.SetFrameStyle( Constants.FRAME_LEFT, 0, 0)
							p_output.SetFrameStyle( Constants.FRAME_BOTTOM, 0, 0)
							p_output.SetFrameStyle( Constants.FRAME_RIGHT, 0, 0)
							p_output.TableCell("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255,255,255), 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
						p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
						writeData38(p_output, p_aAttr[i2].GetValue(false))
						p_output.EndParagraph()
					}

					if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
						p_output.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
					else
						p_output.TableRow()

					// dtable5 local functions:
					/**
					 * @param {Output} p_output The output object
					 * @param {Object } p_Object
					*/
					function writeData38(p_output, p_Object) {
						p_output.OutputLn(p_Object, getString("ID_DEFAULT_FONT"), 12, RGB(0,0,255), Constants.C_TRANSPARENT,  Constants.FMT_UNDERLINE  |  Constants.FMT_LEFT, 0)
					}
				}
				
			if(Context.getSelectedFormat()==Constants.OUTEXCEL) {
				p_output.EndTable("Object attributes", 100, "", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
				xlTableActive=false
			}
			p_output.EndSection()
		
					/** Apply default page format settings to output object
			 * @param {Output} outputObj The output object
			 */
			function sectionPageSetup(outputObj)
			{
				createHeader(outputObj)

				createFooter(outputObj)
				// local functions:
				function createHeader(outputObj) {
					outputObj.BeginHeader()
						if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
							outputObj.BeginTable(100, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
							outputObj.TableRow()
								outputObj.ResetFrameStyle()
								outputObj.TableCell("", 67.1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_CENTER, 0)
							outputObj.BeginParagraph( Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
							writeData30(outputObj, p_aModel[i].Name(nLocale))
							outputObj.EndParagraph()
								outputObj.TableCell("", 32.99, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_RIGHT, 0)
							outputObj.BeginParagraph( Constants.FMT_RIGHT, 1, 1, 0, 0, 0)
														var image = Context.createPicture("vestas.png")
							outputObj.OutGraphic(image, -1, 51, 11 )

							outputObj.EndParagraph()
						if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
							outputObj.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
						else
							outputObj.TableRow()
					outputObj.EndHeader()
						/**
						 * @param {Output} p_output The output object
						 * @param {String } p_String
						*/
						function writeData30(p_output, p_String) {
							p_output.OutputLn(p_String, getString("ID_DEFAULT_FONT"), 13, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_CENTER, 0.71)
						}
				}
				function createFooter(outputObj) {
					outputObj.BeginFooter()
						if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
							outputObj.BeginTable(100, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
							outputObj.TableRow()
								outputObj.ResetFrameStyle()
								outputObj.TableCell("", 22.66, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
							outputObj.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
							outputObj.OutputLn(getString("ID_REPORTDEF_3"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
							outputObj.EndParagraph()
							outputObj.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
							outputObj.OutputField(Constants.FIELD_DATE, getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT)
							outputObj.Output(getString("ID_REPORTDEF_4"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
							outputObj.OutputField(Constants.FIELD_TIME, getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT)
							outputObj.OutputLn("", getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
							outputObj.EndParagraph()
								outputObj.TableCell("", 58, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
							outputObj.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
							outputObj.OutputLn(getString("ID_REPORTDEF_5"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
							outputObj.EndParagraph()
								outputObj.TableCell("", 19.33, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_RIGHT, 0)
							outputObj.BeginParagraph( Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
							outputObj.Output(getString("ID_REPORTDEF_1"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT, 0.71)
							outputObj.OutputField(Constants.FIELD_PAGE, getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT)
							outputObj.Output(getString("ID_REPORTDEF_2"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT, 0.71)
							outputObj.OutputField(Constants.FIELD_NUMPAGES, getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT)
							outputObj.OutputLn("", getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT, 0.71)
							outputObj.EndParagraph()
						if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
							outputObj.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
						else
							outputObj.TableRow()
					outputObj.EndFooter()
				}
			}

		}
		
	if(Context.getSelectedFormat()==Constants.OUTEXCEL) {
		p_output.EndTable(getString("ID_REPORTDEF_35"), 100, "", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
		xlTableActive=false
	}
	p_output.EndSection()


}


function RGB(r, g, b) {
	return (new java.awt.Color(r/255.0,g/255.0,b/255.0,1)).getRGB() & 0xFFFFFF
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


//==========================================================================================================================
//  Sort
//==========================================================================================================================

eSortType = {  
        // values > 0: sort by attribute
    	Default              :  0,    
        Position             :  0,
        Flow                 : -1               
        //PositionPerLane      ; -2,                // nog niet voorzien
        //Custom               : -3,

}
eSortOrientation = {
        // values > 0: sort by attribute        
    	Default              :  0,
        YX                   :  0,
        XY                   : -1 

}
eSortNonStruct = {            
	    Default              :  0, 
        nvt                  :  0,       
        InFlow               :  0,
        Add                  : -1,               
        Exclude              : -2        
}

function Sort_SortModel(pModel, peSortType, peSortOrientation, peSortNonStruct, plLocale){
    // returns: sorted lObjOccs 
    // optional parameters;

    if (arguments.length<2){peSortType=eSortType.Default}
    if (arguments.length<3){peSortOrientation=eSortOrientation.Default}
    if (arguments.length<4){peSortNonStruct=eSortNonStruct.Default}
    if (arguments.length<5){plLocale = -1}

    //returns sorted ObjOccs[]

    var lObjOccs = [];
    var lLocale  = -1;

    if (plLocale!=0){lLocale=plLocale}
    
    if (peSortType==eSortType.Position){
        lObjOccs=Sort_SortObjOccsByPosition(pModel.ObjOccList(),peSortOrientation, plLocale);
    }else if(peSortType==eSortType.Flow){
        lObjOccs = Sort_GetFlowSort(pModel, peSortType, peSortOrientation, peSortNonStruct, plLocale );         
    }else if (peSortType>0){    
        // sortering o.b.v. attribuut ( indien <>1, aanvullen met kriterium "sort op naam")
        lObjOccs = ArisData.sort(pObjOccs, peSortType, constants.AT_NAME, plLocale);      
    }else{        
        lObjOccs = pModel.ObjOccList();
	lObjOccs=Sort_SortObjOccsByPosition(lObjOccs,peSortOrientation, plLocale);                
    }    
    return lObjOccs;       
}
function Sort_SortObjOccsByPosition(pObjOccs, peSortOrientation, plLocale){
    var lSortedObjOccs2Return = new Array();
    // optional parameters:
    if (arguments.length<2){peSortOrientation=eSortOrientation.Default}
    if (arguments.length<3){plLocale = -1}    
    //returns: sorted ObjOccs

    var lSort = Sort_GetSortIdxsByOrientation(pObjOccs, peSortOrientation, plLocale);

    for (var i=0;i<lSort.length;i++){
        lSortedObjOccs2Return.push(pObjOccs[lSort[i]]);
    }
    return lSortedObjOccs2Return;
}

function Sort_GetFlowSort(pModel, peSortType, peSortOrientation, peSortNonStruct, plLocale){ 
    // function returns: gesorteerde array met alle ObjOccs van pModel
    
    // optional parameters;
    if (arguments.length<2){peSortType=eSortType.Default}
    if (arguments.length<3){peSortOrientation=eSortOrientation.Default}
    if (arguments.length<4){peSortNonStruct=eSortNonStruct.Default}
    if (arguments.length<5){plLocale = -1}


      
    var i               = 0;
    var lStartObjOccs   = [];
    var lObjOccs        = [];
    var lSort           = [];
    var lSortIdx        = 0;  
    var lKSObjOccs      = new KeySet();
    var ltObjOccs       = [];
    this.tObjOcc         = function(pObjOcc, plIndex){
        this.ObjOcc    = pObjOcc;
        this.Prepared  = false;
        this.Handled   = false;        
        this.Index     = plIndex;
        this.SortIdx   = -1;  //init
    };
 
    // verzamel alles ObjOccs van model in lKSObjOccs en ltObjOccs[]
    lObjOccs = pModel.ObjOccList();
    for (i=0; i<lObjOccs.length; i++){
        lIndex = lKSObjOccs.AddItem(lObjOccs[i].ObjectID());
        ltObjOccs[lIndex] = new this.tObjOcc(lObjOccs[i],i);
    }
    // verzamel startobjects, en ga daarmee aan de slag, d.w.z. sorteer de flow-elementen
    lStartObjOccs = Aris_GetStartObjOccs(pModel, peSortType, peSortOrientation, plLocale);
    
    if (peSortNonStruct==eSortNonStruct.InFlow){
        // zoek eerst uit welke flowelemeneten er zijn....
        Sort_FlowSort_SortOneLevel (lStartObjOccs, ltObjOccs, lKSObjOccs, lSortIdx
                                        , peSortType, peSortOrientation, peSortNonStruct, plLocale, true); 
    }    
    lSortIdx = Sort_FlowSort_SortOneLevel (lStartObjOccs, ltObjOccs, lKSObjOccs, lSortIdx
                                                , peSortType, peSortOrientation, peSortNonStruct, plLocale); 

    // verzamel de mogelijk overgebleven, en dus nog niet gesorteerde ObjOccs, sorteer ze
    // en werk de struct bij
    if (peSortNonStruct!=eSortNonStruct.Exclude){
        lObjOccs = []; //reset
        for (i=0; i<ltObjOccs.length;i++){
            if (!ltObjOccs[i].Handled){
                lObjOccs[lObjOccs.length]=ltObjOccs[i].ObjOcc;
            }
        }    
    	lObjOccs=Sort_SortObjOccsByPosition(lObjOccs,peSortOrientation, plLocale);                
        for (i=0;i<lObjOccs.length;i++){
            var lIndex = lKSObjOccs.GetIndex(lObjOccs[i].ObjectID());          
            if (!ltObjOccs[lIndex].Handled){
                ltObjOccs[lIndex].Handled=true;
                ltObjOccs[lIndex].SortIdx = lSortIdx++;  
            }
        }
    }
         
    //haal de sortering op in array lSort
    lSort = [];   //reset
    lSortIdx = 0; //reset
    for (i=0;i<ltObjOccs.length;i++){
        lSortIdx = ltObjOccs[i].SortIdx;
        if (lSortIdx>-1){
            lSort[lSortIdx] = i;
        }
    }
    //verzamel gesorteerde ObjOccs aan de hand van lSort
    lObjOccs=[]; //reset
    for (i=0;i<lSort.length;i++){
        lObjOccs[i]=ltObjOccs[lSort[i]].ObjOcc;
    }

    return lObjOccs;
}

function Sort_FlowSort_SortOneLevel (pObjOccs, ptObjOccs, pKSObjOccs, plSortIdx
                                          , peSortType, peSortOrientation, peSortNonStruct, plLocale, pbPrepare){    
    // returns: eerstvolgende Sortidx
    // wijzigt: ptObjOccs en pKSObjOccs
                                              
    // optional parameters;
    if (arguments.length<5){peSortType=eSortType.Default}
    if (arguments.length<6){peSortOrientation=eSortOrientation.Default}
    if (arguments.length<7){peSortNonStruct=eSortNonStruct.Default}
    if (arguments.length<8){pblocale=-1};    
    if (arguments.length<9){pbPrepare=false};
    

    var lIndex   =0;
    var lSortIdx = plSortIdx;
    var lSort    = [];
    var lCxnOccs = [];
    var i        = 0;
      
    
    lSort = Sort_GetSortIdxsByOrientation(pObjOccs, peSortOrientation, plLocale);

    for (i=0;i<lSort.length;i++){
        var lIndex = pKSObjOccs.GetIndex(pObjOccs[lSort[i]].ObjectID());          

        if ((pbPrepare)&&(!ptObjOccs[lIndex].Prepared)
        ||  (!pbPrepare)&&(!ptObjOccs[lIndex].Handled)){
            if (pbPrepare){
                ptObjOccs[lIndex].Prepared = true;
            }else{
                ptObjOccs[lIndex].Handled = true;
                ptObjOccs[lIndex].SortIdx = lSortIdx++;  
            }    
            lObjOcc = ptObjOccs[lIndex].ObjOcc;
            
            var lCxnOccs=[]; //reset
            // indien flowsort, inclusief objects in flow, maar slechts die niet (later) ook als flowelement voorkomen             
            if (!pbPrepare && (peSortNonStruct==eSortNonStruct.InFlow)){
                // voeg de gerelateerde niet-flowelemeneten toe
                var lCxnOccsAdhoc = lObjOcc.OutEdges(Constants.EDGES_NONSTRUCTURE);
                for (j=0;j<lCxnOccsAdhoc.length;j++){
                    var lOthObjOcc = Aris_GetOtherObjOcc(lCxnOccsAdhoc[j], lObjOcc);
                    lIndex = pKSObjOccs.GetIndex(lOthObjOcc.ObjectID());          
                    if (!ptObjOccs[lIndex].Prepared){
                        lCxnOccs[lCxnOccs.length]=lCxnOccsAdhoc[j]
                    }
                }
            }  
            // voeg er de flow elementen aan toe
            lCxnOccs = lCxnOccs.concat.call(lObjOcc.OutEdges(Constants.EDGES_STRUCTURE));
            
            if (lCxnOccs.length > 0){
                lObjOccs=[]; //reset
                for (j=0;j<lCxnOccs.length;j++){
                    lObjOccs[j]=Aris_GetOtherObjOcc(lCxnOccs[j], lObjOcc);                  
                }
                lSortIdx = Sort_FlowSort_SortOneLevel(lObjOccs
                                                          , ptObjOccs
                                                          , pKSObjOccs
                                                          , lSortIdx
                                                          , peSortType 
                                                          , peSortOrientation
                                                          , peSortNonStruct
                                                          , plLocale
                                                          , pbPrepare);
            }
        }      
    }    
    return lSortIdx;
}

function Aris_GetStartObjOccs(pModel, peSortType, peSortOrientation, plLocale){
    if (arguments.length<3){peSortOrientation=eSortOrientation.Default};
    if (arguments.length<4){pbLocale=-1};    
    var lObjOccs = [];
    var lModelType = pModel.TypeNum();
    if ((lModelType==Constants.MT_EEPC)
    ||  (lModelType==Constants.MT_EEPC_COLUMN)  
    ||  (lModelType==Constants.MT_EEPC_ROW)  
    ||  (lModelType==Constants.MT_EEPC_INST)        
    ||  (lModelType==Constants.MT_EEPC_MAT)  
    ||  (lModelType==Constants.MT_EEPC_TAB)){
        lObjOccs = Aris_GetStartObjOccsOfEPC(pModel);
    }else{//default
        lObjOccs = Aris_GetStartObjOccsGeneral(pModel);
    }        
    lObjOccs=Sort_SortObjOccsByPosition(lObjOccs,peSortOrientation, plLocale);
    return lObjOccs;
}
function Aris_GetStartObjOccsOfEPC(pModel){
      var lStartObjOccs2Return = [];
      var lStartObjOccs = [];   
          
      pModel.BuildGraph(true);
      lStartObjOccs = pModel.StartNodeList();
      for (var i=0; i<lStartObjOccs.length;i++){
         var lObjOcc= lStartObjOccs[i];
         if (lObjOcc.InDegree(Constants.EDGES_STRUCTURE) == 0) {
             lStartObjOccs2Return[lStartObjOccs2Return.length]=lObjOcc;             
         }
      }
      return lStartObjOccs2Return;
}
function Aris_GetStartObjOccsGeneral(pModel){
      var lStartObjOccs2Return = [];
      var lStartObjOccs = [];
          
      pModel.BuildGraph(true);
      lStartObjOccs = pModel.StartNodeList();
      for (var i=0; i<lStartObjOccs.length;i++){
          var lObjOcc= lStartObjOccs[i];    
          if (lObjOcc.InDegree(Constants.EDGES_ALL) == 0) {
              lStartObjOccs2Return[lStartObjOccs2Return.length]=lObjOcc;             
          }
      }
      return lStartObjOccs2Return;
}

function Aris_GetStartObjOccsAllActivObjects(pModel){  
    var lStartObjOccs2Return = [];      
    var lActivObjectIter=gActivObjectTypeSet.iterator();
    while(lActivObjectIter.hasNext()){
        lStartObjOccs2Return=lStartObjOccs2Return.concat(pModel.ObjOccListFilter(parseInt(lActivObjectIter.next())));
    }      
    return lStartObjOccs2Return;  
}
function Aris_GetStartObjOccsAllocationDgm(pModel){        
      var lObjectType=-1;    
      if (Aris_ModelIsAllocDgm(pModel)){
          switch(pModel.TypeNum()){
              case Constants.MT_FUNC_ALLOC_DGM:
                lObjectType=Constants.OT_FUNC;
                break;
          }
      }
      return pModel.ObjOccListFilter(lObjectType);      
      
}

function Sort_GetSortIdxsByOrientation(pObjOccs, peSortOrientation, plLocale){
    var lSort2Return = new Array();
    var lKeySet   = new KeySet();
    var lObjOcc   = null;
    var lsAttrVal = "";
    var i         = 0;
    if (peSortOrientation>0){         
        // sortering op attribuutwaarde
        for (i=0;i<pObjOccs.length;i++){
            lObjOcc = pObjOccs[i];
            lsAttrVal= CS_NL.ArisUtils.getAttrValue(lObjOcc.ObjDef(), peSortOrientation, plLocale, true);
            lKeySet.AddItem(lsAttrVal + String.fromCharCode(1) + lObjOcc.ObjDef().Name(plLocale));
        }
        lSort2Return = lKeySet.Sort();
    }else if (peSortOrientation==eSortOrientation.XY==true){         
        lSort2Return = Aris_GetSortIdxsByXY(pObjOccs);
    }else{//default
        lSort2Return = Aris_GetSortIdxsByYX(pObjOccs);            
    }
    return lSort2Return;
}

    
function Aris_GetSortedObjOccssByYX(pObjOccs){
    var lSortedObjOccs2Return = new Array();
    var lSort = Aris_GetSortIdxsByYX(pObjOccs);
    for (var i=0;i<lSort.length;i++){
        lSortedObjOccs2Return.push(pObjOccs[lSort[i]]);
    }        
    return lSortedObjOccs2Return;
}

function Aris_GetSortIdxsByXY(pObjOccs){
    var lSort2Return = new Array();
    // returns lSort[] met de gesorteerde indexes van pObjOccs
    // routine houdt rekening met lichte afwijkingen in de posities 
    //(daarom is de routine zo razend ingewikkeld)

    if (pObjOccs.length==0){
        return [];
    }else if (pObjOccs.length==1){
        return [0];
    }else{
          
        var i             = 0;
        var j             = 0;
        var lObjOccs      = [];
        var lKeySet       = new KeySet();        
        var lKeySet2      = new KeySet();
        var lSortIdx      = 0;
        var lIndex        = 0;
        var lLanesIndex   = 0;
        var lObjIndex     = 0;

        var tLane = function(){
            this.Left    = 0;
            this.Right   = 0;
            this.Ql      = 0;
            this.Qr      = 0;            
            this.ObjIdxs = [];
        }         
        var lLanes = [];
        
        var tObj= function(){
            this.Lane = -1;
            this.Zx = 0;
            this.Right = 0;
        }            
        var ltObjs = [];        

        // sorteer de objectbreedtes
        for (var i=0; i<pObjOccs.length; i++){
            lKeySet2.AddItem(fcHex(pObjOccs[i].Width(),8)+ i);   // sort breedte
            lKeySet.AddItem(pObjOccs[i].ObjectID());       // verzamel OjectIDs
            ltObjs[i]= new tObj();
            ltObjs[i].Lane=-1;   
            ltObjs[i].Zx = parseInt(pObjOccs[i].X() + (pObjOccs[i].Width()/2));    
            ltObjs[i].Right = parseInt(pObjOccs[i].X() + pObjOccs[i].Width());                        
        }
        var lSort = lKeySet2.Sort();      //lSort bevat de gesorteerde objectbreedtes
        
        // probeer de ObjOccs in zo'n min mogelijk lanes te stoppen, begin met de smalste
        // een Objocc zit in een lane, als zijn x-midellpunt erin valt 
        // of als hij links dan wel rechtes over meer dan de helft van de lane valt
        var lLanesIndex   = 0;
        var lObjIndex     = 0;
        var lSortIndex    = 0;
        lObjIndex = lSort[lSortIndex];
        do{
            lLanes[lLanesIndex]= new tLane();
            lLanes[lLanesIndex].ObjIndex = lObjIndex;            // deze index is de owner van de lane          
            lLanes[lLanesIndex].Left     = pObjOccs[lObjIndex].X();
            lLanes[lLanesIndex].Ql       = pObjOccs[lObjIndex].X() + pObjOccs[lObjIndex].Width()/4;
            lLanes[lLanesIndex].Qr       = ltObjs[lObjIndex].Right - pObjOccs[lObjIndex].Width()/4;;
            lLanes[lLanesIndex].Right    = ltObjs[lObjIndex].Right;
            lLanes[lLanesIndex].ObjIdxs  = [];
            lObjIndex = -1; //reset
            for (var i=0; i<pObjOccs.length; i++){    
                j= lSort[i];
                if (ltObjs[j].Lane==-1){
                    if ((ltObjs[j].Zx >= lLanes[lLanesIndex].Left)
                    &&  (ltObjs[j].Zx <= lLanes[lLanesIndex].Right)){
                        ltObjs[j].Lane = -2; 
                    }else 
                    if ((pObjOccs[j].X() < lLanes[lLanesIndex].Ql)
                    &&  (ltObjs[j].Right > lLanes[lLanesIndex].Qr)){
                        ltObjs[j].Lane = -2; 
                    }else{
                        //volgende, nieuwe lane
                        if (lObjIndex ==-1){
                            lObjIndex = j;
                        }
                    }
                }
            }     
            if (lObjIndex!=-1){
                lLanesIndex++;
            }
        }while (lObjIndex!=-1)  
        
        //voor alle pPObjOccs is er nu tenminste 1 lane        
        // sorteer de lanes op de x-positie
        lKeySet2 = new KeySet(); 
        for (i=0; i<lLanes.length;i++){
          lKeySet2.AddItem(fcHex(lLanes[i].Left,8));
        }
        var lLaneSort = lKeySet2.Sort();
        
        // deel de ObjOccs nu definitief in in een lane        
        for(i=0; i<pObjOccs.length; i++){    
            for(j=0;j<lLaneSort.length;j++){                
                var k= lLaneSort[j];
                if ((ltObjs[i].Zx >= lLanes[k].Left)
                &&  (ltObjs[i].Zx <= lLanes[k].Right)){
                    ltObjs[i].Lane = k;  //lLanesIndex;
                    lLanes[k].ObjIdxs[lLanes[k].ObjIdxs.length] = i;
                    break;
                }else 
                if ((pObjOccs[i].X() < lLanes[k].Ql)
                &&  (ltObjs[i].Right > lLanes[k].Qr)){
                    ltObjs[i].Lane = k; //lLanesIndex;
                    lLanes[k].ObjIdxs[lLanes[k].ObjIdxs.length] = i;                        
                    break;
                }
            }
        }
        
        // sorteer per lane de ObjOccs, per YX(!), bewaar de sortering in lSort2Return
        for(i=0;i<lLaneSort.length;i++){
            lObjOccs=[];
            for (j=0;j<lLanes[lLaneSort[i]].ObjIdxs.length;j++){
                lObjOccs[lObjOccs.length]=pObjOccs[lLanes[lLaneSort[i]].ObjIdxs[j]];
            }
            lObjOccs = ArisData.sort(lObjOccs, Constants.SORT_Y, Constants.SORT_X, -1);      
            for (j=0;j<lObjOccs.length;j++){
                lIndex = lKeySet.GetIndex(lObjOccs[j].ObjectID());
                lSort2Return[lSortIdx++] = lIndex;
            }
        }                
        return lSort2Return;
    }
}

function Aris_GetSortedObjOccssByXY(pObjOccs){
    var lSortedObjOccs2Return = new Array();
    var lSort = Aris_GetSortIdxsByXY(pObjOccs);
    for (var i=0;i<lSort.length;i++){
        lSortedObjOccs2Return.push(pObjOccs[lSort[i]]);
    }        
    return lSortedObjOccs2Return;
}

function Aris_GetSortIdxsByYX(pObjOccs){
    var lSort2Return= new Array();
    
    // returns lSort[] met de gesorteerde indexes van pObjOccs
    // routine houdt rekening met lichte afwijkingen in de posities 
    //(daarom is de routine zo razend ingewikkeld)
    
    if (pObjOccs.length==0){
        return [];
    }else if (pObjOccs.length==1){
        return [0];
    }else{
          
        var i             = 0;
        var j             = 0;
        var lObjOccs      = [];
        var lKeySet       = new KeySet();        
        var lKeySet2      = new KeySet();
        var lSortIdx      = 0;
        var lIndex        = 0;        

        var tLane = function(){
            this.Top     = 0;
            this.Bottom  = 0;
            this.Qt      = 0;
            this.Qb      = 0;
            this.ObjIdxs = [];
        }         
        var lLanes = [];
        
        var tObj= function(){
            this.Lane = -1;
            this.Zy = 0;
            this.Bottom = 0;
        }            
        var ltObjs = [];      
        
        // sorteer de objectbreedtes
        for (var i=0; i<pObjOccs.length; i++){
            lKeySet2.AddItem(fcHex(pObjOccs[i].Height(),8)+ fcHex(pObjOccs[i].Y(),8)+ fcHex(pObjOccs[i].X(),8)+ i);   // sort de objoccs op hoogte
            lKeySet.AddItem(pObjOccs[i].ObjectID());              // verzamel ook de OjectIDs
            ltObjs[i]= new tObj();
            ltObjs[i].Lane=-1;   
            ltObjs[i].Zy = parseInt(pObjOccs[i].Y() + (pObjOccs[i].Height()/2));    
            ltObjs[i].Bottom = parseInt(pObjOccs[i].Y() + pObjOccs[i].Height());                        
        }
        var lSort = lKeySet2.Sort();      //lSort bevat de gesorteerde objecthoogtes
        
        var lLanesIndex   = 0;
        var lObjIndex     = 0;
        var lSortIndex    = 0 
        // probeer de ObjOccs in zo'n min mogelijk lanes te stoppen, begin met de smalste
        // een Objocc zit in een lane, als zijn y-midellpunt erin valt 
        // of als hij boven dan wel onder over meer dan de helft van de lane valt

        lObjIndex = lSort[lSortIndex];
        do{           
            lLanes[lLanesIndex]= new tLane();
            lLanes[lLanesIndex].ObjIndex = lObjIndex;            // deze index is de owner van de lane          
            lLanes[lLanesIndex].Top      = pObjOccs[lObjIndex].Y();
            lLanes[lLanesIndex].Bottom   = ltObjs[lObjIndex].Bottom;
            lLanes[lLanesIndex].Qt       = pObjOccs[lObjIndex].Y() + pObjOccs[lObjIndex].Height()/4;
            lLanes[lLanesIndex].Qb       = ltObjs[lObjIndex].Bottom - pObjOccs[lObjIndex].Height()/4;;


            lLanes[lLanesIndex].ObjIdxs=[];
            lObjIndex = -1; //reset
            for (i=0; i<pObjOccs.length; i++){    
                j= lSort[i];
                if (ltObjs[j].Lane==-1){
                    if ((ltObjs[j].Zy >= lLanes[lLanesIndex].Top)
                    &&  (ltObjs[j].Zy <= lLanes[lLanesIndex].Bottom)){
                        ltObjs[j].Lane = -2;  //lLanesIndex;
                        //lLanes[lLanesIndex].ObjIdxs[lLanes[lLanesIndex].ObjIdxs.length] = j;
                    }else 
                    if ((pObjOccs[j].Y() < lLanes[lLanesIndex].Qt)
                    &&  (ltObjs[j].Bottom > lLanes[lLanesIndex].Qb)){
                        ltObjs[j].Lane = -2; //lLanesIndex;
                       // lLanes[lLanesIndex].ObjIdxs[lLanes[lLanesIndex].ObjIdxs.length] = j;                        
                    }else{
                        //volgende, nieuwe lane
                        if (lObjIndex ==-1){
                            lObjIndex = j;
                        }
                    }
                }
            }     
            if (lObjIndex!=-1){
                lLanesIndex++;
            }
        }while (lObjIndex!=-1)  
        
        //voor alle pPObjOccs is er nu tenminste 1 lane
        // sorteer de lanes op de y-positie
        lKeySet2 = new KeySet(); 
        for (i=0; i<lLanes.length;i++){
          lKeySet2.AddItem(fcHex(lLanes[i].Top,8));
        }
        var lLaneSort = lKeySet2.Sort();
        // deel de ObjOccs nu definitief in in een lane        
        for(i=0; i<pObjOccs.length; i++){    
            for(j=0;j<lLaneSort.length;j++){                
                var k= lLaneSort[j];
                if ((ltObjs[i].Zy >= lLanes[k].Top)
                &&  (ltObjs[i].Zy <= lLanes[k].Bottom)){
                    ltObjs[i].Lane = k;  //lLanesIndex;
                    lLanes[k].ObjIdxs[lLanes[k].ObjIdxs.length] = i;
                    break;
                }else 
                if ((pObjOccs[i].Y() < lLanes[k].Qt)
                &&  (ltObjs[i].Bottom > lLanes[k].Qb)){
                    ltObjs[i].Lane = k; //lLanesIndex;
                    lLanes[k].ObjIdxs[lLanes[k].ObjIdxs.length] = i;                        
                    break;
                }
            }
        }
        
        // sorteer per lane de ObjOccs, per XY(!), bewaar de sortering in lSort2Return
        for(i=0;i<lLaneSort.length;i++){
            lObjOccs=[];
            for (j=0;j<lLanes[lLaneSort[i]].ObjIdxs.length;j++){
                lObjOccs[lObjOccs.length]=pObjOccs[lLanes[lLaneSort[i]].ObjIdxs[j]];
            }
            lObjOccs = ArisData.sort(lObjOccs, Constants.SORT_X, Constants.SORT_Y, -1);      
            for (j=0;j<lObjOccs.length;j++){
                lIndex = lKeySet.GetIndex(lObjOccs[j].ObjectID());
                lSort2Return[lSortIdx++] = lIndex;
            }
        }                
        return lSort2Return;
    }
}

function Sort_GetModelSortSettings(pModel){
    // returns array: SortType, SortOrienation, SortNonStruct
    
    var lReturn=[];
    
    // routine zal moeten beginnen aan het model te vragen wat zijn settings zijn (komt nog).
    // als die ontbreken worden de defaultsettings van de database uitgevraagd (komt nog)
    // als ook die ontbreken dan vragen we aan bestand 'ModelSortSettings.txt' in de 
    // reportfolder wat de defaultsettings zijn
    var lSettings=Sort_GetModelSortSettingsByFile(pModel);
    if (lSettings.length<1){
        lReturn[0]=eSortType.Default;
    }else{
        lReturn[0]=lSettings[0];
    }
    if (lSettings.length<2){
        lReturn[1]=eSortOrientation.Default;
    }else{
        lReturn[1]=lSettings[1];
    }
    if (lSettings.length<3){
        lReturn[2]=eSortNonStruct.Default;
    }else{
        lReturn[2]=lSettings[2];
    }
    return lReturn;
}    
    
function Sort_GetModelSortSettingsByFile(pModel){        
    // returns array: SortType, SortOrienation, SortNonStruct
    var lReturn = [];
    var lbOk = false;
    
    var leSortType = 0;
    var leSortOrientation = 0;
    var leSortNonStruct = 0;
    
    this.getSortType=function(pPars){
        var leSortType = eSortType.Default;
        if (pPars.length>=2){
            switch (fTrim(lPars[1].toLowerCase())){
            case "flow"     : leSortType = eSortType.Flow; break;
            case "position" : leSortType = eSortType.Position; break; 
            }
        }
        return leSortType;
    }
    this.getSortOrientation=function(pPars){
        var leSortOrientation = eSortOrientation.Default;
        if (lPars.length>=3){
            if (parseInt(lPars[2])>0){
                leSortOrientation = parseInt(lPars[2]);
            }else{
                switch (fTrim(lPars[2].toLowerCase())){
                case "yx"       : leSortOrientation = eSortOrientation.YX; break;
                case "xy"       : leSortOrientation = eSortOrientation.XY; break; 
                case "at_name"  : leSortOrientation = Constants.AT_NAME; break; 
                }
            }                     
        }
        return leSortOrientation;
    }
    this.getNonStruct=function(pPars){
        var leNonStruct = eNonStruct.Default;
        if (lPars.length>=3){
            switch (fTrim(lPars[3].toLowerCase())){
            case "inflow"   : leSortNonStruct = eSortNonStruct.InFlow; break;
            case "add"      : leSortNonStruct = eSortNonStruct.Add; break; 
            case "exclude"  : leSortNonStruct = eSortNonStruct.Exclude; break;                     
            }                     
        }
        return leSortNonStruct;
    }
    
    var lByteArray=Context.getFile("ModelSortSettings.cfg", Constants.LOCATION_SCRIPT);
    
    if (lByteArray.length=0){return lReturn};
        
    var lFileContent="";
    for(var i=0;i<lByteArray.length;i++){
        lFileContent+=String.fromCharCode(lByteArray[i]);
    }
    var lLines =lFileContent.split("\n");
    var lDefLine = -1;
    for (var i=0;i<lLines.length;i++){
        var lPars=lLines[i].split(",");
        try{
            if (eval("Constants." + lPars[0].toUpperCase())==pModel.TypeNum()){
                lbOk = true;               
                leSortType        = this.getSortType(lPars);
                leSortOrientation = this.getSortOrientation(lPars);
                leSortNonStruct   = this.getSortNonStruct(lPars);
            }
        }catch(err){
            if (fTrim(lPars[0])=="0"){                
                lDefLine = i;
            }
        }
    }
    if ((!lbOk)&&(lDefLine>-1)){
        // modeltype zelf is niet gedefinieerd, wel een default bij waarde 0
        var lPars=lLines[lDefLine].split(",");
        leSortType        = this.getSortType(lPars);
        leSortOrientation = this.getSortOrientation(lPars);
        leSortNonStruct   = this.getSortNonStruct(lPars);
    }else if (lbOk){
        lReturn[0]=leSortType;
        lReturn[1]=leSortOrientation;
        lReturn[2]=leSortNonStruct;
    }
    return lReturn;   
}

function Aris_GetDefaultFontStyle(){
    var lFontStyle2Return = null;

	var lFontStyleList= ArisData.getActiveDatabase().FontStyleList();
    for (var i=0;i<lFontStyleList.length;i++){
        var lFontStyle = lFontStyleList[i];
        
		if (lFontStyle.IsDefaultFontStyle()){
            lFontStyle2Return = lFontStyle;
            break;
        }
    }
    return lFontStyle2Return
}
function fIsMissing(pOptionalPar){
    var lbIsMissing2Return = false;
    if (pOptionalPar==undefined || pOptionalPar ==null){
        lbIsMissing2Return = true;        
    }
    return lbIsMissing2Return;
}
function fValueIfMissing(pOptional,pDefault){
    return (fIsMissing(pOptional)?pDefault:pOptional)
}
function fcHex(plNum, plLen){
    var lLen = fValueIfMissing(plLen,0)
    var lsHex="";
    if(fIsNumeric(plNum)){
        lsHex = plNum.toString(16).toUpperCase();   
        if (lLen>0)for(var i=lsHex.length; i<lLen; i++)lsHex= "0" + lsHex;             
    }
    return lsHex;
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
  var lsValue2Return= "" + psValue;
  lsValue2Return=lsValue2Return.replace(/^ */,"");  
  lsValue2Return=lsValue2Return.replace(/ *$/,"");  
  return lsValue2Return;
}
function Aris_GetOtherObjOcc (pCxnOcc, pObjOcc){
    var lObjOcc2Return = null;
    
    var lSource = pCxnOcc.SourceObjOcc(); 
    var lTarget = pCxnOcc.TargetObjOcc();
    
    if (pObjOcc.IsEqual(lSource)){
        lObjOcc2Return = lTarget;
    }else if (pObjOcc.IsEqual(lTarget)){
        lObjOcc2Return = lSource;
    }else{
        lObjOcc2Return = null;
    }
    return lObjOcc2Return;
}

main();