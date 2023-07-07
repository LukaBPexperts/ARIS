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
createSection1(oOutput, ArisData.getSelectedDatabases())
createSection2(oOutput, ArisData.getSelectedModels())
oOutput.WriteReport()

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
		iteration_section2(p_output, ArisData.sort(p_aModel[i].ObjOccListFilter(Constants.OT_FUNC),Constants.SORT_OBJDEF_NAME, nLocale))
		p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
		p_output.OutputLn(getString("ID_REPORTDEF_33"), getString("ID_DEFAULT_FONT"), 14, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_LEFT, 0)
		p_output.EndParagraph()
		dtable(p_output, ArisData.sort(p_aModel[i].ObjDefListByTypes([Constants.OT_INFO_CARR]),Constants.AT_NAME, nLocale))
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
				p_output.EndTable(getString("ID_REPORTDEF_37"), 100, "", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
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
				p_output.EndTable(getString("ID_REPORTDEF_37"), 100, "", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
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
			p_output.BeginParagraph( Constants.FMT_LEFT, 0, 0, 0, 0, 0)
			p_output.OutputLn(getString("ID_REPORTDEF_26"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_LEFT, 0)
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData20(p_output, p_aModel[i].Attribute(Constants.AT_TITL1, nLocale).GetValue(false))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData21(p_output, p_aModel[i].Attribute(Constants.AT_EXT_1, nLocale).GetValue(false))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData22(p_output, p_aModel[i].Attribute(Constants.AT_TITL2, nLocale).GetValue(false))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData23(p_output, p_aModel[i].Attribute(Constants.AT_EXT_2, nLocale).GetValue(false))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData24(p_output, p_aModel[i].Attribute(Constants.AT_TITL3, nLocale).GetValue(false))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData25(p_output, p_aModel[i].Attribute(Constants.AT_EXT_3, nLocale).GetValue(false))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData26(p_output, p_aModel[i].Attribute(Constants.AT_TITL4, nLocale).GetValue(false))
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			writeData27(p_output, p_aModel[i].Attribute(Constants.AT_LINK, nLocale).GetValue(false))
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
		 * @param {Object } p_Object
		*/
		function writeData20(p_output, p_Object) {
			p_output.OutputLn(p_Object, getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
		}
		/**
		 * @param {Output} p_output The output object
		 * @param {Object } p_Object
		*/
		function writeData21(p_output, p_Object) {
			p_output.OutputLn(p_Object, getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0)
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
		function iteration_section2(p_output, p_aObjOcc)
		{
			// do not create new section if all data is empty
			if(p_aObjOcc.length==0)
				return
		
			//produce one table sheet per section
			if(xlTableActive && Context.getSelectedFormat()==Constants.OUTEXCEL) {
				p_output.EndTable(getString("ID_REPORTDEF_37"), 100, "", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
				xlTableActive=false
			}
			p_output.BeginSection(false, Constants.SECTION_DEFAULT)
			sectionPageSetup( p_output )

			if(Context.getSelectedFormat()==Constants.OUTEXCEL) {
				p_output.BeginTable(100, Constants.C_BLACK, Constants.C_TRANSPARENT,  Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
				xlTableActive=true
			}
			p_output.BeginParagraph( Constants.FMT_LEFT, 0, 0, 0, 0, 0)
			p_output.OutputLn(getString("ID_REPORTDEF_27"), getString("ID_DEFAULT_FONT"), 14, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_LEFT, 0)
			p_output.EndParagraph()
			p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
			p_output.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
			p_output.EndParagraph()
			// repetition of queried data:
			for(var i1=0; i1<p_aObjOcc.length; i1++) {
				p_output.BeginParagraph( Constants.FMT_LEFT, 0, 0, 0, 0, 0)
				writeData29(p_output, p_aObjOcc[i1].getObjDefName(nLocale))
				p_output.EndParagraph()
				p_output.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
				writeData30(p_output, p_aObjOcc[i1].ObjDef().Attribute(Constants.AT_DESC, nLocale))
				p_output.EndParagraph()
                
                var responsibleCxnDefList = filterCxnByTypeNums(p_aObjOcc[i1], new Array(Constants.CT_EXEC_1,Constants.CT_EXEC_2,Constants.CT_IS_RESP_1,Constants.CT_IS_RESP_2,Constants.CT_IS_RESP_FOR_1,Constants.CT_IS_RESP_FOR_2));                
				//dtable(p_output, p_aObjOcc[i1].ObjDef().CxnListFilter(Constants.EDGES_IN, new Array(Constants.CT_EXEC_1,Constants.CT_EXEC_2,Constants.CT_IS_RESP_1,Constants.CT_IS_RESP_2,Constants.CT_IS_RESP_FOR_1,Constants.CT_IS_RESP_FOR_2)))
				dtable(p_output, responsibleCxnDefList);
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
            
            function filterCxnByTypeNums(pObjOcc, allowedCxnTypeNums){
                var allowedCxnDefList = new Array();
                var cxnOccList = pObjOcc.CxnOccList();
                for(var cxnOccCount=0; cxnOccCount<cxnOccList.length; cxnOccCount++){
                    var cxnOcc = cxnOccList[cxnOccCount];
                    for(var cxnTypeNumCount=0; cxnTypeNumCount<allowedCxnTypeNums.length; cxnTypeNumCount++){
                        if(allowedCxnTypeNums[cxnTypeNumCount]==cxnOcc.CxnDef().TypeNum()){
                            allowedCxnDefList.push(cxnOcc.CxnDef());
                        }
                    }
                }
                
                return allowedCxnDefList;
            }

			// iteration_section2 local functions:
			/**
			 * @param {Output} p_output The output object
			 * @param {String } p_String
			*/
			function writeData29(p_output, p_String) {
				p_output.OutputLn(p_String, getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_LEFT, 0)
			}
			/**
			 * @param {Output} p_output The output object
			 * @param {Attr } p_Attr
			*/
			function writeData30(p_output, p_Attr) {
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
						p_output.OutputLn(getString("ID_REPORTDEF_28"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 12.70)
						p_output.EndParagraph()
							p_output.TableCell("", 49.38, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255,255,255), 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
						p_output.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
						writeData31(p_output, p_aCxnDef[i2].SourceObjDef().Name(nLocale))
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
					function writeData31(p_output, p_String) {
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
						p_output.OutputLn(getString("ID_REPORTDEF_29"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 13.41)
						p_output.EndParagraph()
							p_output.TableCell("", 49.38, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255,255,255), 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
						p_output.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
						writeData32(p_output, p_aAttr[i2].GetValue(false))
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
					function writeData32(p_output, p_Object) {
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
						p_output.OutputLn(getString("ID_REPORTDEF_30"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 13.41)
						p_output.EndParagraph()
							p_output.TableCell("", 49.38, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255,255,255), 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
						p_output.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
						writeData33(p_output, p_aAttr[i2].GetValue(false))
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
					function writeData33(p_output, p_Object) {
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

					// dtable3 local functions:
					/**
					 * @param {Output} p_output The output object
					 * @param {Object } p_Object
					*/
					function writeData34(p_output, p_Object) {
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

					// dtable4 local functions:
					/**
					 * @param {Output} p_output The output object
					 * @param {Object } p_Object
					*/
					function writeData35(p_output, p_Object) {
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
					p_output.OutputLn(getString("ID_REPORTDEF_26"), getString("ID_DEFAULT_FONT"), 12, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_UNDERLINE  |  Constants.FMT_LEFT, 0)
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
						writeData36(p_output, p_aAttr[i2].GetValue(false))
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
					function writeData36(p_output, p_Object) {
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
							writeData28(outputObj, p_aModel[i].Name(nLocale))
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
						function writeData28(p_output, p_String) {
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
		
		/**
		 * @param {Output} p_output The output object
		 * @param {ObjDef[]} p_aObjDef 
		*/
		function dtable(p_output, p_aObjDef)
		{
			// do not create new table if all data is empty
			if(p_aObjDef.length==0)
				return
		
			if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
				p_output.BeginTable(100, RGB(0,0,0), RGB(255,255,255),  Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
			p_output.TableRow()
				p_output.ResetFrameStyle()
				p_output.TableCell("", 31.19, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255,255,255), 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
			p_output.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
			p_output.OutputLn(getString("ID_REPORTDEF_34"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_LEFT, 0.71)
			p_output.EndParagraph()
				p_output.TableCell("", 35.34, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255,255,255), 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
			p_output.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
			p_output.OutputLn(getString("ID_REPORTDEF_35"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_LEFT, 0.71)
			p_output.EndParagraph()
				p_output.TableCell("", 33.47, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255,255,255), 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
			p_output.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
			p_output.OutputLn(getString("ID_REPORTDEF_36"), getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_BOLD  |  Constants.FMT_LEFT, 0.71)
			p_output.EndParagraph()
			// repetition of queried data:
			for(var i1=0; i1<p_aObjDef.length; i1++) {
				p_output.TableRow()
					p_output.ResetFrameStyle()
					p_output.TableCell("", 31.19, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255,255,255), 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
				p_output.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
				writeData31(p_output, p_aObjDef[i1].Attribute(Constants.AT_NAME, nLocale).GetValue(false))
				p_output.EndParagraph()
					p_output.TableCell("", 35.34, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255,255,255), 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
				p_output.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
				writeData32(p_output, p_aObjDef[i1].Attribute(Constants.AT_EXT_1, nLocale).GetValue(false))
				p_output.EndParagraph()
					p_output.TableCell("", 33.47, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(255,255,255), 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
				p_output.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
				writeData33(p_output, p_aObjDef[i1].Attribute(Constants.AT_EXT_2, nLocale).GetValue(false))
				p_output.EndParagraph()
			}

			if(Context.getSelectedFormat()!=Constants.OUTEXCEL)
				p_output.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
			else
				p_output.TableRow()

			// dtable local functions:
			/**
			 * @param {Output} p_output The output object
			 * @param {Object } p_Object
			*/
			function writeData31(p_output, p_Object) {
				p_output.OutputLn(p_Object, getString("ID_DEFAULT_FONT"), 11, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
			}
			/**
			 * @param {Output} p_output The output object
			 * @param {Object } p_Object
			*/
			function writeData32(p_output, p_Object) {
				p_output.OutputLn(p_Object, getString("ID_DEFAULT_FONT"), 11, RGB(0,0,255), Constants.C_TRANSPARENT,  Constants.FMT_UNDERLINE  |  Constants.FMT_LEFT, 0.71)
			}
			/**
			 * @param {Output} p_output The output object
			 * @param {Object } p_Object
			*/
			function writeData33(p_output, p_Object) {
				p_output.OutputLn(p_Object, getString("ID_DEFAULT_FONT"), 11, RGB(0,0,255), Constants.C_TRANSPARENT,  Constants.FMT_UNDERLINE  |  Constants.FMT_LEFT, 0.71)
			}
		}
		
	if(Context.getSelectedFormat()==Constants.OUTEXCEL) {
		p_output.EndTable(getString("ID_REPORTDEF_37"), 100, "", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
		xlTableActive=false
	}
	p_output.EndSection()


}


function RGB(r, g, b) {
	return (new java.awt.Color(r/255.0,g/255.0,b/255.0,1)).getRGB() & 0xFFFFFF
}
