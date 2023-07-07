var nLocale = Context.getSelectedLanguage()
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

var startModels = ArisData.getSelectedModels()

main(startModels);

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
    outputObj.BeginTable(100, [40,30,30], RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0 )
        //outputObj.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
    outputObj.TableRow()
    outputObj.ResetFrameStyle()
    outputObj.SetFrameStyle(Constants.FRAME_TOP, 20, Constants.BRDR_NORMAL)
    outputObj.SetFrameStyle(Constants.FRAME_LEFT, 20, Constants.BRDR_NORMAL)
    outputObj.SetFrameStyle(Constants.FRAME_BOTTOM, 20, Constants.BRDR_NORMAL)
    outputObj.SetFrameStyle(Constants.FRAME_RIGHT,20, Constants.BRDR_NORMAL)
    
    outputObj.TableCell("", 3, 1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0 )   
    //outputObj.TableCell("", 40, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0)
    outputObj.BeginParagraph(Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
    var image = Context.createPicture("apisLogo.png")
    outputObj.OutGraphic(image, -1, 55, 20)
    outputObj.EndParagraph()
    
    outputObj.TableCell("", 1, 1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0 )
    //outputObj.TableCell("", 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
    outputObj.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    outputObj.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
    outputObj.EndParagraph()
    
    outputObj.TableCell("", 1, 1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0 )
    //outputObj.TableCell("", 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
    outputObj.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    outputObj.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
    outputObj.EndParagraph()
    
    outputObj.TableRow()
    outputObj.TableCell("", 1, 1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0 )
    //outputObj.TableCell("", 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
    outputObj.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    outputObj.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
    outputObj.EndParagraph()
    
    outputObj.TableCell("", 1, 1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0 )
    //outputObj.TableCell("", 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
    outputObj.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    outputObj.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
    outputObj.EndParagraph()
    
    

    if (Context.getSelectedFormat() != Constants.OUTEXCEL)
        outputObj.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
    else
        outputObj.TableRow()
    outputObj.EndHeader()
}

/* function globalFooter(outputObj) {
    outputObj.BeginFooter()
    if (Context.getSelectedFormat() != Constants.OUTEXCEL)
        outputObj.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
    outputObj.TableRow()
    outputObj.ResetFrameStyle()
    outputObj.SetFrameStyle(Constants.FRAME_TOP, 20, Constants.BRDR_NORMAL)
    outputObj.SetFrameStyle(Constants.FRAME_LEFT, 20, Constants.BRDR_NORMAL)
    outputObj.SetFrameStyle(Constants.FRAME_BOTTOM, 20, Constants.BRDR_NORMAL)
    outputObj.SetFrameStyle(Constants.FRAME_RIGHT,20, Constants.BRDR_NORMAL)
    outputObj.TableCell(Constants.SCRIPT_NAME, 70, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
    outputObj.TableCell("Str." + Constants.FIELD_PAGE + " od " + Constants.FIELD_NUMPAGES, 30, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_CENTER, 0)
    
    if (Context.getSelectedFormat() != Constants.OUTEXCEL)
        outputObj.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
    else
        outputObj.TableRow()
    outputObj.EndFooter()
} */

function globalFooter(outputObj) {
	outputObj.BeginFooter()
		//if( !isExcel() )
			outputObj.BeginTable(100, Constants.C_TRANSPARENT, Constants.C_TRANSPARENT,  Constants.FMT_LEFT | Constants.FMT_NOBORDER, 0)
			outputObj.TableRow()
				outputObj.ResetFrameStyle()
				outputObj.SetFrameStyle( Constants.FRAME_TOP, 0, 0)
				outputObj.SetFrameStyle( Constants.FRAME_LEFT, 0, 0)
				outputObj.SetFrameStyle( Constants.FRAME_BOTTOM, 0, 0)
				outputObj.SetFrameStyle( Constants.FRAME_RIGHT, 0, 0)
				outputObj.TableCell("", 50, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_LEFT, 0)
			outputObj.BeginParagraph( Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
			outputObj.OutputField(Constants.FIELD_DATE, "Arial", 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT)
			outputObj.OutputLn("", "Arial", 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_LEFT, 0.71)
			outputObj.EndParagraph()
				outputObj.TableCell("", 50, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0,  Constants.FMT_VTOP |  Constants.FMT_RIGHT, 0)
			outputObj.BeginParagraph( Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
			outputObj.Output("Page", "Arial", 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT, 0.71)
			outputObj.OutputField(Constants.FIELD_PAGE, "Arial", 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT)
			outputObj.Output("of", "Arial", 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT, 0.71)
			outputObj.OutputField(Constants.FIELD_NUMPAGES, getString("ID_DEFAULT_FONT"), 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT)
			outputObj.OutputLn("", "Arial", 10, RGB(0,0,0), Constants.C_TRANSPARENT,  Constants.FMT_RIGHT, 0.71)
			outputObj.EndParagraph()
		//if( !isExcel() )
			outputObj.EndTable("", 100, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
		//else
			//outputObj.TableRow()
	outputObj.EndFooter()
}


function main(startModels) {

    
//create first page
firstPage()

//create second page
oOutput.BeginSection(false, Constants.SECTION_DEFAULT)

    oOutput.SetAutoTOCNumbering(true)
    globalHeader(oOutput)
    globalFooter(oOutput)
    
    emptyrow(2)
    
    //Izrada/Izmjene:
    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("Izrada/Izmjene:", getString("ID_STYLE_RD_DEFAULT"), 16, RGB(0, 0, 0),  Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()
    
    emptyrow(1)
    
	oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
	oOutput.TableRow()
	oOutput.ResetFrameStyle()
	oOutput.TableCell("", 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Verzija", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Opis", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Ime i Prezime", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
    oOutput.TableCell("", 30, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Organizacijska jedinica", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
    oOutput.TableCell("", 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Datum", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
    
    for(var i=0; i<5; i++){      
        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)	
        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)	
        oOutput.TableCell("", 30, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
        oOutput.TableCell("", 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    }    
    oOutput.EndTable("", 100, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)
	
    
    emptyrow(2)
    
    //Pregledao/la
    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("Pregledao/la:", getString("ID_STYLE_RD_DEFAULT"), 16, RGB(0, 0, 0),  Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()
    
    emptyrow(1)
    
	oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
	oOutput.TableRow()
	oOutput.ResetFrameStyle()
	oOutput.TableCell("", 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Verzija", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Ime i Prezime", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
    oOutput.TableCell("", 40, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Organizacijska jedinica", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Datum pregleda", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
    
      
    oOutput.TableRow()
    oOutput.ResetFrameStyle()
    oOutput.TableCell("", 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)	
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    oOutput.TableCell("", 40, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)	
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)

  
    oOutput.EndTable("", 100, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)
    
    
    emptyrow(2)
    
    //Odobrio/la
    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("Odobrio/la:", getString("ID_STYLE_RD_DEFAULT"), 16, RGB(0, 0, 0),  Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()
    
    emptyrow(1)
    
	oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
	oOutput.TableRow()
	oOutput.ResetFrameStyle()
	oOutput.TableCell("", 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Verzija", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Ime i Prezime", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
    oOutput.TableCell("", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Potpis", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Organizacijska jedinica", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Datum pregleda", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
    
      
    oOutput.TableRow()
    oOutput.ResetFrameStyle()
    oOutput.TableCell("", 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)	
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    oOutput.TableCell("", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)	
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)

  
    oOutput.EndTable("", 100, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)
    
    
emptyrow(2)
    
    //Lista za distribuciju:
    oOutput.addLocalBookmark("TOC");
    oOutput.BeginParagraphF(getString("ID_STYLE_RD_HEADING_1"));
    oOutput.OutputLnF("Control de cambios", getString("ID_STYLE_RD_DEFAULT"));
    oOutput.EndParagraph();
    
    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("Lista za distribuciju:", getString("ID_STYLE_RD_DEFAULT"), 16, RGB(0, 0, 0),  Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()
    
    emptyrow(1)
    
	oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
	oOutput.TableRow()
	oOutput.ResetFrameStyle()
	oOutput.TableCell("", 30, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Ime i prezime", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()   
    oOutput.TableCell("", 70, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Organizacijska jedinica", getString("ID_DEFAULT_FONT"), 9, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER, 0.71)
	oOutput.EndParagraph()
    
    
    for(var i=0; i<10; i++){      
        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 30, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)	
        oOutput.TableCell("", 75, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
        
    }    
    oOutput.EndTable("", 100, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)
    
    
oOutput.EndSection()

//TOC
tableOfContent()

//UVOD
oOutput.BeginSection(false, Constants.SECTION_DEFAULT)
    globalHeader(oOutput)
    globalFooter(oOutput)
    
    emptyrow(1)
    
    //UVOD
    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("1. Uvod", getString("ID_STYLE_RD_DEFAULT"), 16, RGB(0, 0, 0),  Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()
    
    emptyrow(10)
    
    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("    1.1. Pravni okvir / Referentni dokumenti", getString("ID_STYLE_RD_DEFAULT"), 16, RGB(0, 0, 0),  Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()
    
    oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
	oOutput.TableRow()
	oOutput.ResetFrameStyle()
	oOutput.TableCell("", 80, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_LEFT, 0)
	oOutput.BeginParagraph(Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Dokument", getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT| Constants.FMT_BOLD, 0.71)
	oOutput.EndParagraph()   
    oOutput.TableCell("", 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
	oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
	oOutput.OutputLn("Datum", getString("ID_DEFAULT_FONT"),10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER | Constants.FMT_BOLD, 0.71)
	oOutput.EndParagraph()
    
    
    for(var i=0; i<11; i++){      
        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 80, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)	
        oOutput.TableCell("", 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
        
    }    
    oOutput.EndTable("", 100, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)
    
    
oOutput.EndSection()

//***************************************
//create rest pages with tables
    
for(var i=0; i<startModels.length; i++){

//******************************************************************************************
//RJECNIK SKRACENICA
//******************************************************************************************
oOutput.BeginSection(false, Constants.SECTION_DEFAULT)
    globalHeader(oOutput)
    globalFooter(oOutput)

//emptyrow(1)
   
oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
oOutput.Output("2. Rječnik skraćenica", getString("ID_STYLE_RD_DEFAULT"), 16, RGB(0, 0, 0),  Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
oOutput.EndParagraph()

emptyrow(1)

oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
oOutput.TableRow()
oOutput.ResetFrameStyle()
oOutput.TableCell("", 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_LEFT, 0)
oOutput.BeginParagraph(Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
oOutput.OutputLn("Skraćenica", getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT| Constants.FMT_BOLD, 0.71)
oOutput.EndParagraph()   
oOutput.TableCell("", 80, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
oOutput.OutputLn("Objašnjenje", getString("ID_DEFAULT_FONT"),10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER | Constants.FMT_BOLD, 0.71)
oOutput.EndParagraph()

var skracenice = startModels[0].ObjDefListFilter("Skracenica", nLocale, Constants.OT_TECH_TRM, ArisData.ActiveFilter().UserDefinedAttributeTypeNum("9fec8670-6240-11ec-4a03-10604b88abdc"))



for(var i=0; i<skracenice.length; i++){      
    var skraceniceDescription = skracenice[i].Attribute(Constants.AT_DESC, nLocale).getValue()
    oOutput.TableRow()
    oOutput.ResetFrameStyle()
    oOutput.TableCell(skracenice[i].Name(nLocale), 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)	
    oOutput.TableCell(skraceniceDescription, 80, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    
}    
oOutput.EndTable("", 100, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)

    
oOutput.EndSection()   

//******************************************************************************************
//RJECNIK POJMOVA
//******************************************************************************************
oOutput.BeginSection(false, Constants.SECTION_DEFAULT)

    globalHeader(oOutput)
    globalFooter(oOutput)

//emptyrow(1)
   
oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
oOutput.Output("3. Rječnik pojmova", getString("ID_STYLE_RD_DEFAULT"), 16, RGB(0, 0, 0),  Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
oOutput.EndParagraph()

emptyrow(1)

oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
oOutput.Output("    3.1. Popis pojmova", getString("ID_STYLE_RD_DEFAULT"), 16, RGB(0, 0, 0),  Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
oOutput.EndParagraph()

emptyrow(1)

oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
oOutput.TableRow()
oOutput.ResetFrameStyle()
oOutput.TableCell("", 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_LEFT, 0)
oOutput.BeginParagraph(Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
oOutput.OutputLn("Pojam", getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT| Constants.FMT_BOLD, 0.71)
oOutput.EndParagraph()   
oOutput.TableCell("", 80, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
oOutput.OutputLn("Objašnjenje", getString("ID_DEFAULT_FONT"),10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_CENTER | Constants.FMT_BOLD, 0.71)
oOutput.EndParagraph()

var pojmovi = startModels[0].ObjDefListFilter("Pojam", nLocale, Constants.OT_TECH_TRM, ArisData.ActiveFilter().UserDefinedAttributeTypeNum("9fec8670-6240-11ec-4a03-10604b88abdc"))

for(var i=0; i<pojmovi.length; i++){      
    //var pojmoviDescription = pojmovi[i].Attribute(Constants.AT_DESC, nLocale).getValue()
    oOutput.TableRow()
    oOutput.ResetFrameStyle()
    oOutput.TableCell(pojmovi[i].Name(nLocale), 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)	
    oOutput.TableCell("", 80, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    
}    
oOutput.EndTable("", 100, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)

emptyrow(1)

oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
oOutput.Output("    3.1. Objašnjenje pojmova", getString("ID_STYLE_RD_DEFAULT"), 16, RGB(0, 0, 0),  Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
oOutput.EndParagraph()

emptyrow(1)

for(var i=0; i<pojmovi.length; i++){      
    var pojmoviDescription = pojmovi[i].Attribute(Constants.AT_DESC, nLocale).getValue()
        
    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output(pojmoviDescription, getString("ID_STYLE_RD_DEFAULT"), 11, RGB(0, 0, 0),  Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()
    
    emptyrow(1)
} 
    
oOutput.EndSection()  

//********************************************************************************
//NAPOMENE I OTVORENA PITANJA
//********************************************************************************

oOutput.BeginSection(false, Constants.SECTION_DEFAULT)
    globalHeader(oOutput)
    globalFooter(oOutput)


//emptyrow(1)
   
oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
oOutput.Output("4. Napomene", getString("ID_STYLE_RD_DEFAULT"), 16, RGB(0, 0, 0),  Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
oOutput.EndParagraph()

emptyrow(10)

oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
oOutput.Output("5. Otvorena pitanja", getString("ID_STYLE_RD_DEFAULT"), 16, RGB(0, 0, 0),  Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
oOutput.EndParagraph()


oOutput.EndSection()  

//********************************************************************************
//INDEKS POJMOVA
//********************************************************************************

oOutput.BeginSection(false, Constants.SECTION_DEFAULT)

    globalHeader(oOutput)
    globalFooter(oOutput)

//emptyrow(1)
   
oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
oOutput.Output("6. Indeks pojmova", getString("ID_STYLE_RD_DEFAULT"), 16, RGB(0, 0, 0),  Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
oOutput.EndParagraph()


oOutput.EndSection()
}

}

function RGB(r, g, b) {
    return (new java.awt.Color(r / 255.0, g / 255.0, b / 255.0, 1)).getRGB() & 0xFFFFFF
}

function emptyrow(number) {
    for (var i = 0; i < number; i++)
        oOutput.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
}

function firstPage() {

	oOutput.BeginSection(false, Constants.SECTION_COVER)

	//oOutput.SetAutoTOCNumbering(true)
	globalHeader(oOutput)

	emptyrow(15)

	oOutput.BeginParagraph(Constants.FMT_CENTER, 0, 0, 0, 2, 0)
	oOutput.OutputLn("F109 eFiskalizacija", getString("ID_DEFAULT_FONT"), 21, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0)
	oOutput.OutputLn("eFiskalizacija bezgotovinskih računa", getString("ID_DEFAULT_FONT"), 21, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0)
    oOutput.OutputLn("Rječnik", getString("ID_DEFAULT_FONT"), 21, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0)
	oOutput.EndParagraph()

  	oOutput.EndSection()

}

function tableOfContent(){

oOutput.BeginSection(false, Constants.SECTION_INDEX);
//Header
oOutput.BeginHeader();
oOutput.BeginTable(100, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_CENTER, 0);
oOutput.TableRow();
oOutput.TableCell("Content table header", 100, "Arial", 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
oOutput.EndTable("Test", 100, "Arial", 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT, 0);
oOutput.EndHeader();
//Footer
oOutput.BeginFooter();
oOutput.BeginTable(100, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_CENTER, 0);
oOutput.TableRow();
oOutput.TableCell("Content table footer", 100, "Arial", 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
oOutput.EndTable("Test", 100, "Arial", 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT, 0);
oOutput.EndFooter();
//Sets the autonumbering on
oOutput.SetAutoTOCNumbering(false);
//Defines the style for the four levels in the TOC
oOutput.SetTOCFormat(0, "Arial", 20, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_LEFT | Constants.FMT_BOLD | Constants.FMT_ITALIC, 0, 0, 0, 0);
oOutput.SetTOCFormat(1, "Arial", 16, Constants.C_GREEN, Constants.C_WHITE, Constants.FMT_LEFT | Constants.FMT_BOLD, 5, 5, 2, 2);
oOutput.SetTOCFormat(2, "Arial", 12, Constants.C_RED, Constants.C_WHITE, Constants.FMT_LEFT | Constants.FMT_ITALIC, 10 , 5, 2, 2);
oOutput.SetTOCFormat(3, "Arial", 10, Constants.C_BLUE, Constants.C_WHITE, Constants.FMT_LEFT, 15, 5, 2, 2);
//Output text
oOutput.OutputLn("Table_of_Contents", "Arial", 30, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_LEFT, 0);
//Adds the table of content
oOutput.OutputField(Constants.FIELD_TOC, "Arial", 10, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_LEFT);
oOutput.EndSection();




}