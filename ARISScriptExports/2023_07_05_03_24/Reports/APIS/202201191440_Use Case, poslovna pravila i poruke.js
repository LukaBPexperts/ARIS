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

var startObjDefs = ArisData.getSelectedObjDefs()

main(startObjDefs);

oOutput.WriteReport()

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
        //outputObj.BeginTable(100, [40, 30, 30], RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
        outputObj.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
    outputObj.TableRow()
    outputObj.ResetFrameStyle()
    outputObj.SetFrameStyle(Constants.FRAME_TOP, 20, Constants.BRDR_NORMAL)
    outputObj.SetFrameStyle(Constants.FRAME_LEFT, 20, Constants.BRDR_NORMAL)
    outputObj.SetFrameStyle(Constants.FRAME_BOTTOM, 20, Constants.BRDR_NORMAL)
    outputObj.SetFrameStyle(Constants.FRAME_RIGHT, 20, Constants.BRDR_NORMAL)

    //outputObj.TableCell("", 3, 1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0)
    outputObj.TableCell("", 40, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0)
    outputObj.BeginParagraph(Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
    var image = Context.createPicture("apisLogo.png")
    outputObj.OutGraphic(image, -1, 55, 20)
    outputObj.EndParagraph()

    //outputObj.TableCell("", 1, 1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0)
    outputObj.TableCell("", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
    outputObj.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    outputObj.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
    outputObj.EndParagraph()

    //outputObj.TableCell("", 1, 1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0)
    outputObj.TableCell("", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
    outputObj.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    outputObj.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
    outputObj.EndParagraph()

    //outputObj.TableCell("", 1, 1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0)
    outputObj.TableCell("", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
    outputObj.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    outputObj.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
    outputObj.EndParagraph()

    //outputObj.TableCell("", 1, 1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0)
    outputObj.TableCell("", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
    outputObj.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    outputObj.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
    outputObj.EndParagraph()

    for (var i = 0; i < 1; i++) {
        outputObj.TableRow()
        //outputObj.TableCell("", 1, 1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0)
        outputObj.TableCell("", 40, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
        outputObj.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
        outputObj.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
        outputObj.EndParagraph()

        //outputObj.TableCell("", 1, 1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0)
        outputObj.TableCell("", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
        outputObj.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
        outputObj.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
        outputObj.EndParagraph()

        //outputObj.TableCell("", 1, 1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0)
        outputObj.TableCell("", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
        outputObj.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
        outputObj.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
        outputObj.EndParagraph()

        //outputObj.TableCell("", 1, 1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0)
        outputObj.TableCell("", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
        outputObj.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
        outputObj.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
        outputObj.EndParagraph()

        //outputObj.TableCell("", 1, 1, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_CENTER | Constants.FMT_LEFT, 0)
        outputObj.TableCell("", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
        outputObj.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
        outputObj.OutputLnF("", getString("ID_STYLE_RD_DEFAULT"))
        outputObj.EndParagraph()
    }


    if (Context.getSelectedFormat() != Constants.OUTEXCEL)
        outputObj.EndTable("", 100, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
    else
        outputObj.TableRow()
    outputObj.EndHeader()
}

function globalFooter(outputObj) {
    outputObj.BeginFooter()
    //if( !isExcel() )
    outputObj.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
    outputObj.TableRow()
    outputObj.ResetFrameStyle()
    outputObj.SetFrameStyle(Constants.FRAME_TOP, 20, Constants.BRDR_NORMAL)
    outputObj.SetFrameStyle(Constants.FRAME_LEFT, 20, Constants.BRDR_NORMAL)
    outputObj.SetFrameStyle(Constants.FRAME_BOTTOM, 20, Constants.BRDR_NORMAL)
    outputObj.SetFrameStyle(Constants.FRAME_RIGHT, 20, Constants.BRDR_NORMAL)

    outputObj.TableCell("", 50, "Segoe UI", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_LEFT, 0)
    outputObj.BeginParagraph(Constants.FMT_LEFT, 0.71, 0.71, 0, 0, 0)
    outputObj.OutputField(Constants.FIELD_FILENAME, "Segoe UI", 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT)
    outputObj.OutputLn("", "Segoe UI", 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0.71)
    outputObj.EndParagraph()

    outputObj.TableCell("", 50, "Segoe UI", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_VTOP | Constants.FMT_RIGHT, 0)
    outputObj.BeginParagraph(Constants.FMT_RIGHT, 0.71, 0.71, 0, 0, 0)
    outputObj.Output("Str. ", "Segoe UI", 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
    outputObj.OutputField(Constants.FIELD_PAGE, "Segoe UI", 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT)
    outputObj.Output(" od ", "Segoe UI", 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
    outputObj.OutputField(Constants.FIELD_NUMPAGES, getString("ID_DEFAULT_FONT"), 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT)
    outputObj.OutputLn("", "Segoe UI", 10, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_RIGHT, 0.71)
    outputObj.EndParagraph()
    //if( !isExcel() )
    outputObj.EndTable("", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
    //else
    //outputObj.TableRow()
    outputObj.EndFooter()
}

function main(startObjDefs) {

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
    oOutput.Output("Izradio/la:", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    emptyrow(1)

    oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
    oOutput.TableRow()
    oOutput.ResetFrameStyle()
    oOutput.TableCell("", 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Verzija", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Ime i Prezime", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 40, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Organizacijska jedinica", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Datum izrade", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()

    for (var i = 0; i < 4; i++) {
        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 10, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
        oOutput.TableCell("", 40, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    }
    oOutput.EndTable("", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)

    emptyrow(2)

    //Pregledao/la
    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("Pregledao/la:", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    emptyrow(1)

    oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
    oOutput.TableRow()
    oOutput.ResetFrameStyle()
    oOutput.TableCell("", 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Verzija", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Ime i Prezime", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 40, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Organizacijska jedinica", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Datum pregleda", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()

    oOutput.TableRow()
    oOutput.ResetFrameStyle()
    oOutput.TableCell("", 10, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    oOutput.TableCell("", 40, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)

    oOutput.EndTable("", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)

    emptyrow(2)

    //Odobrio/la
    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("Odobrio/la:", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    emptyrow(1)

    oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
    oOutput.TableRow()
    oOutput.ResetFrameStyle()
    oOutput.TableCell("", 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Verzija", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Ime i Prezime", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 15, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Potpis", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Funkcija/Organizacijska jedinica", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Datum odobrenja", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()

    oOutput.TableRow()
    oOutput.ResetFrameStyle()
    oOutput.TableCell("", 10, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    oOutput.TableCell("", 15, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)

    oOutput.EndTable("", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)

    emptyrow(2)

    //Povijest izmjena
    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("Povijest izmjena:", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    emptyrow(1)

    oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
    oOutput.TableRow()
    oOutput.ResetFrameStyle()
    oOutput.TableCell("", 10, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Verzija", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Ime i Prezime", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Organizacijska jedinica", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 40, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Opis", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()

    for (var i = 0; i < 2; i++) {
        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 10, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
        oOutput.TableCell("", 25, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
        oOutput.TableCell("", 40, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    }

    oOutput.EndTable("", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)

    emptyrow(2)

    //Lista za distribuciju:
    //oOutput.addLocalBookmark("TOC");
    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("Lista za distribuciju:", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    emptyrow(1)

    oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
    oOutput.TableRow()
    oOutput.ResetFrameStyle()
    oOutput.TableCell("", 30, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Ime i prezime", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 70, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Organizacijska jedinica", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()

    for (var i = 0; i < 1; i++) {
        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell("", 30, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
        oOutput.TableCell("", 75, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    }
    oOutput.EndTable("", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)

    oOutput.EndSection()

    //TOC
    tableOfContent()

    //UVOD
    oOutput.BeginSection(false, Constants.SECTION_DEFAULT)
    globalHeader(oOutput)
    globalFooter(oOutput)

    emptyrow(1)

    //UVOD
    oOutput.BeginParagraphF(getString("HEADING_1"))
    oOutput.Output("1. Rječnik skraćenica i pojmova", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("[Napomena – Upisati tekst]", getString("ID_STYLE_RD_DEFAULT"), 11, RGB(237, 28, 36), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_ITALIC, 0)
    oOutput.EndParagraph()

    emptyrow(2)

    //2. Opis
    oOutput.BeginParagraphF(getString("HEADING_1"))
    oOutput.Output("2. Opis", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()
    //ispis atribut sa objekta  Opis UC		2428133	9c45d280-479a-11ec-4a03-10604b88abdc	User-defined	\APIS; 	
    var opisUC = startObjDefs[0].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("9c45d280-479a-11ec-4a03-10604b88abdc"), nLocale).getValue()
    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output(opisUC, getString("ID_STYLE_RD_DEFAULT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    oOutput.EndSection()

    //******************************************************************************************
    //3. SUDIONICI (AKTERI)
    //******************************************************************************************
    oOutput.BeginSection(false, Constants.SECTION_DEFAULT)
    globalHeader(oOutput)
    globalFooter(oOutput)

    //emptyrow(1)

    oOutput.BeginParagraphF(getString("HEADING_1"))
    oOutput.Output("3. Sudionici (akteri)", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    //ispis slike modela sa assignanong FADa MT_FUNC_ALLOC_DGM
    var assignedFads = startObjDefs[0].AssignedModels(Constants.MT_FUNC_ALLOC_DGM)

    //get first assigned FAD model
    var modelGraphic = assignedFads[0].Graphic(false, false, nLocale)

    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutGraphic(modelGraphic, -1, 200, 200)
    oOutput.EndParagraph()

    //ISPIS AKTERE, ROLE SPOJENE NA OBJEKT DEF //DORADITI
    emptyrow(2)
    var primarniAkteriRole = startObjDefs[0].getConnectedObjs([Constants.OT_PERS_TYPE], Constants.EDGES_INOUT, [Constants.CT_EXEC_2])
    var ostaliAkteriRole = startObjDefs[0].getConnectedObjs([Constants.OT_PERS_TYPE], Constants.EDGES_INOUT, [Constants.CT_CONTR_TO_2])

    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("Primarni akteri:", getString("ID_STYLE_RD_DEFAULT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    for (var i = 0; i < primarniAkteriRole.length; i++) {
        var longDescriptionAtt = primarniAkteriRole[i].Attribute(Constants.AT_DESC, nLocale).getValue()
        oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
        oOutput.Output(primarniAkteriRole[i].Name(nLocale) + " - ", getString("ID_STYLE_RD_DEFAULT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
        oOutput.Output(longDescriptionAtt, getString("ID_STYLE_RD_DEFAULT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
        oOutput.EndParagraph()
    }

    emptyrow(1)
    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("Ostali akteri:", getString("ID_STYLE_RD_DEFAULT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    for (var i = 0; i < ostaliAkteriRole.length; i++) {
        var longDescriptionAtt = ostaliAkteriRole[i].Attribute(Constants.AT_DESC, nLocale).getValue()
        oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
        oOutput.Output(ostaliAkteriRole[i].Name(nLocale) + " - ", getString("ID_STYLE_RD_DEFAULT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
        oOutput.Output(longDescriptionAtt, getString("ID_STYLE_RD_DEFAULT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0)
        oOutput.EndParagraph()
    }

    oOutput.EndSection()
    //******************************************************************************************
    //4. Scenariji
    //******************************************************************************************
    oOutput.BeginSection(false, Constants.SECTION_DEFAULT)
    globalHeader(oOutput)
    globalFooter(oOutput)

    oOutput.BeginParagraphF(getString("HEADING_1"))
    oOutput.Output("4. Scenariji", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    //ispisi atribut Scenarij		2100453	118b5651-479b-11ec-4a03-10604b88abdc	User-defined	\APIS; 	
    var scenarij = startObjDefs[0].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("118b5651-479b-11ec-4a03-10604b88abdc"), nLocale).getHTMLFormattedValue()
    oOutput.OutputFormattedText(scenarij.getHTML())

    emptyrow(1)

    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("Iznimke:", getString("ID_STYLE_RD_DEFAULT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    var iznimke = startObjDefs[0].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("facce371-479a-11ec-4a03-10604b88abdc"), nLocale).getHTMLFormattedValue()
    oOutput.OutputFormattedText(iznimke.getHTML())

    emptyrow(1)
    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("Alternativni scenariji:", getString("ID_STYLE_RD_DEFAULT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    var alternativniScenariji = startObjDefs[0].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("3478d5c1-479b-11ec-4a03-10604b88abdc"), nLocale).getHTMLFormattedValue()
    oOutput.OutputFormattedText(alternativniScenariji.getHTML())

    oOutput.EndSection()

    //******************************************************************************************
    //5. Poslovna pravila
    //******************************************************************************************
    oOutput.BeginSection(false, Constants.SECTION_DEFAULT)
    globalHeader(oOutput)
    globalFooter(oOutput)

    oOutput.BeginParagraphF(getString("HEADING_1"))
    oOutput.Output("5. Poslovna pravila", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    //TABLE
    emptyrow(1)

    oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
    oOutput.TableRow()
    oOutput.ResetFrameStyle()
    oOutput.TableCell("", 20, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Naziv", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()
    oOutput.TableCell("", 80, getString("ID_DEFAULT_FONT"), 10, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER, 0)
    oOutput.BeginParagraph(Constants.FMT_CENTER, 0.71, 0.71, 0, 0, 0)
    oOutput.OutputLn("Opis", getString("ID_DEFAULT_FONT"), 11, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0.71)
    oOutput.EndParagraph()

    var businessRules = startObjDefs[0].getConnectedObjs([Constants.OT_BUSINESS_RULE], Constants.EDGES_INOUT)
    for (var i = 0; i < businessRules.length; i++) {
        var longDescriptionAtt = businessRules[i].Attribute(Constants.AT_DESC, nLocale).getValue()

        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell(businessRules[i].Name(nLocale), 20, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VCENTER, 0)
        oOutput.TableCell(longDescriptionAtt, 80, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
    }
    oOutput.EndTable("", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)

    oOutput.EndSection()

    //******************************************************************************************
    //6. Podatkovni skupovi
    //******************************************************************************************
    oOutput.BeginSection(false, Constants.SECTION_DEFAULT)
    globalHeader(oOutput)
    globalFooter(oOutput)

    oOutput.BeginParagraphF(getString("HEADING_1"))
    oOutput.Output("6. Podatkovni skupovi", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    emptyrow(1)

    var entityTypeObjects = startObjDefs[0].getConnectedObjs([Constants.OT_ENT_TYPE], Constants.EDGES_INOUT)
    for (var i = 0; i < entityTypeObjects.length; i++) {

        oOutput.BeginParagraphF(getString("HEADING_2"))
        oOutput.Output("6." + (i + 1) + " " + entityTypeObjects[i].Name(nLocale), getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
        oOutput.EndParagraph()

        var ermAttributeObjects = entityTypeObjects[i].getConnectedObjs([Constants.OT_ERM_ATTR], Constants.EDGES_INOUT)

        if (ermAttributeObjects.length != 0) {
            oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
            oOutput.TableRow()
            oOutput.ResetFrameStyle()
            oOutput.TableCell("R.br", 7, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_BOLD | Constants.FMT_CENTER, 0)
            oOutput.TableCell("Podatak", 16, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_BOLD | Constants.FMT_CENTER, 0)
            oOutput.TableCell("Tip", 14, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_BOLD | Constants.FMT_CENTER, 0)
            oOutput.TableCell("Obvezan", 12, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_BOLD | Constants.FMT_CENTER, 0)
            oOutput.TableCell("Pojavnost [opcionalno]", 17, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_BOLD | Constants.FMT_CENTER, 0)
            oOutput.TableCell("Opis / Napomena", 17, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_BOLD | Constants.FMT_CENTER, 0)
            oOutput.TableCell("Izvor podatka [opcionalno]", 17, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_BOLD | Constants.FMT_CENTER, 0)

            for (var j = 0; j < ermAttributeObjects.length; j++) {
                var tip = ermAttributeObjects[j].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("b4aaaaa0-617e-11ec-4a03-10604b88abdc"), nLocale).getValue()
                var obvezan = ermAttributeObjects[j].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("ea4cebf0-5d96-11ec-4a03-10604b88abdc"), nLocale).getValue()
                var pojavnost = ermAttributeObjects[j].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("6f3d2870-6701-11ec-4a03-10604b88abdc"), nLocale).getValue()
                var opisNapomena = ermAttributeObjects[j].Attribute(Constants.AT_DESC, nLocale).getValue()
                var izvorPodatka = ermAttributeObjects[j].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("5dd5b761-6183-11ec-4a03-10604b88abdc"), nLocale).getValue()

                oOutput.TableRow()
                oOutput.ResetFrameStyle()
                oOutput.TableCell((j + 1), 7, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
                oOutput.TableCell(ermAttributeObjects[j].Name(nLocale), 16, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
                oOutput.TableCell(tip, 14, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
                oOutput.TableCell(obvezan, 12, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
                oOutput.TableCell(pojavnost, 17, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
                oOutput.TableCell(opisNapomena, 17, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
                oOutput.TableCell(izvorPodatka, 17, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_CENTER, 0)
            }
            oOutput.EndTable("", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)
        }
        else {
            emptyrow(1)
        }
    }

    oOutput.EndSection()

    var technicalTerms = startObjDefs[0].getConnectedObjs([Constants.OT_TECH_TRM], Constants.EDGES_INOUT)
    //******************************************************************************************
    //7. Poruke u aplikaciji
    //******************************************************************************************
    oOutput.BeginSection(false, Constants.SECTION_DEFAULT)
    globalHeader(oOutput)
    globalFooter(oOutput)

    oOutput.BeginParagraphF(getString("HEADING_1"))
    oOutput.Output("7. Poruke u aplikaciji", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    emptyrow(1)

    oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
    oOutput.TableRow()
    oOutput.ResetFrameStyle()
    oOutput.TableCell("Naziv", 15, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
    oOutput.TableCell("Opis", 85, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)


    for (var j = 0; j < technicalTerms.length; j++) {
        var attributeTemp = technicalTerms[j].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("9fec8670-6240-11ec-4a03-10604b88abdc"), nLocale).getValue()

        if (attributeTemp == "Poruka sustava") {
            var longDescriptionAtt = technicalTerms[j].Attribute(Constants.AT_DESC, nLocale).getValue()
            oOutput.TableRow()
            oOutput.ResetFrameStyle()
            oOutput.TableCell(technicalTerms[j].Name(nLocale), 15, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
            oOutput.TableCell(longDescriptionAtt, 85, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
        }
    }
    oOutput.EndTable("", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)

    oOutput.EndSection()

    //******************************************************************************************
    //8. Poslovni zahtijevi
    //******************************************************************************************
    oOutput.BeginSection(false, Constants.SECTION_DEFAULT)
    globalHeader(oOutput)
    globalFooter(oOutput)

    oOutput.BeginParagraphF(getString("HEADING_1"))
    oOutput.Output("8. Poslovni zahtijevi", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    emptyrow(1)

    oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
    oOutput.TableRow()
    oOutput.ResetFrameStyle()
    oOutput.TableCell("Naziv", 15, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
    oOutput.TableCell("Opis", 85, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)

    var requirements = startObjDefs[0].getConnectedObjs([Constants.OT_REQUIREMENT], Constants.EDGES_INOUT)
    for (var j = 0; j < requirements.length; j++) {
        var longDescriptionAtt = requirements[j].Attribute(Constants.AT_DESC, nLocale).getValue()
        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell(requirements[j].Name(nLocale), 15, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
        oOutput.TableCell(longDescriptionAtt, 85, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)

    }
    oOutput.EndTable("", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)

    oOutput.EndSection()

    //******************************************************************************************
    //9. Korisnička sučelja
    //******************************************************************************************
    oOutput.BeginSection(false, Constants.SECTION_DEFAULT)
    globalHeader(oOutput)
    globalFooter(oOutput)

    oOutput.BeginParagraphF(getString("HEADING_1"))
    oOutput.Output("9. Korisnička sučelja", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    emptyrow(1)

    oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
    oOutput.TableRow()
    oOutput.ResetFrameStyle()
    oOutput.TableCell("Naziv", 15, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
    oOutput.TableCell("Opis", 85, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)

    var screens = startObjDefs[0].getConnectedObjs([Constants.OT_SCRN], Constants.EDGES_INOUT)
    for (var j = 0; j < screens.length; j++) {
        var longDescriptionAtt = screens[j].Attribute(Constants.AT_DESC, nLocale).getValue()
        oOutput.TableRow()
        oOutput.ResetFrameStyle()
        oOutput.TableCell(screens[j].Name(nLocale), 15, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
        oOutput.TableCell(longDescriptionAtt, 85, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)

    }
    oOutput.EndTable("", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)

    oOutput.EndSection()

    //******************************************************************************************
    //10. Pojmovi
    //******************************************************************************************
    oOutput.BeginSection(false, Constants.SECTION_DEFAULT)
    globalHeader(oOutput)
    globalFooter(oOutput)

    oOutput.BeginParagraphF(getString("HEADING_1"))
    oOutput.Output("10. Pojmovi", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    emptyrow(1)

    oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
    oOutput.TableRow()
    oOutput.ResetFrameStyle()
    oOutput.TableCell("Naziv", 15, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
    oOutput.TableCell("Opis", 85, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)


    for (var j = 0; j < technicalTerms.length; j++) {
        var attributeTemp = technicalTerms[j].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("9fec8670-6240-11ec-4a03-10604b88abdc"), nLocale).getValue()

        if (attributeTemp == "Pojam") {
            var longDescriptionAtt = technicalTerms[j].Attribute(Constants.AT_DESC, nLocale).getValue()
            oOutput.TableRow()
            oOutput.ResetFrameStyle()
            oOutput.TableCell(technicalTerms[j].Name(nLocale), 15, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
            oOutput.TableCell(longDescriptionAtt, 85, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
        }
    }
    oOutput.EndTable("", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)

    oOutput.EndSection()

    //******************************************************************************************
    //11. Skraćenice
    //******************************************************************************************
    oOutput.BeginSection(false, Constants.SECTION_DEFAULT)
    globalHeader(oOutput)
    globalFooter(oOutput)

    oOutput.BeginParagraphF(getString("HEADING_1"))
    oOutput.Output("11. Skraćenice", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    emptyrow(1)

    oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
    oOutput.TableRow()
    oOutput.ResetFrameStyle()
    oOutput.TableCell("Naziv", 15, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
    oOutput.TableCell("Opis", 85, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)


    for (var j = 0; j < technicalTerms.length; j++) {
        var attributeTemp = technicalTerms[j].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("9fec8670-6240-11ec-4a03-10604b88abdc"), nLocale).getValue()

        if (attributeTemp == "Skracenica") {
            var longDescriptionAtt = technicalTerms[j].Attribute(Constants.AT_DESC, nLocale).getValue()
            oOutput.TableRow()
            oOutput.ResetFrameStyle()
            oOutput.TableCell(technicalTerms[j].Name(nLocale), 15, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
            oOutput.TableCell(longDescriptionAtt, 85, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
        }
    }
    oOutput.EndTable("", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)

    oOutput.EndSection()

    //******************************************************************************************
    //12. Dokumenti
    //******************************************************************************************
    oOutput.BeginSection(false, Constants.SECTION_DEFAULT)
    globalHeader(oOutput)
    globalFooter(oOutput)

    oOutput.BeginParagraphF(getString("HEADING_1"))
    oOutput.Output("12. Dokumenti", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    emptyrow(1)

    oOutput.BeginTable(100, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_REPEAT_HEADER, 0)
    oOutput.TableRow()
    oOutput.ResetFrameStyle()
    oOutput.TableCell("Naziv", 15, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)
    oOutput.TableCell("Opis", 85, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, RGB(217, 217, 217), 0, Constants.FMT_CENTER | Constants.FMT_BOLD, 0)

    for (var j = 0; j < technicalTerms.length; j++) {
        var attributeTemp = technicalTerms[j].Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum("9fec8670-6240-11ec-4a03-10604b88abdc"), nLocale).getValue()

        if (attributeTemp == "Dokument") {
            var longDescriptionAtt = technicalTerms[j].Attribute(Constants.AT_DESC, nLocale).getValue()
            oOutput.TableRow()
            oOutput.ResetFrameStyle()
            oOutput.TableCell(technicalTerms[j].Name(nLocale), 15, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
            oOutput.TableCell(longDescriptionAtt, 85, getString("ID_DEFAULT_FONT"), 11, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0)
        }
    }
    oOutput.EndTable("", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP, 0)

    oOutput.EndSection()

    //******************************************************************************************
    //13. Otvorena pitanja
    //******************************************************************************************
    oOutput.BeginSection(false, Constants.SECTION_DEFAULT)
    globalHeader(oOutput)
    globalFooter(oOutput)

    oOutput.BeginParagraphF(getString("HEADING_1"))
    oOutput.Output("13. Otvorena pitanja", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("[Napomena – Upisati tekst]", getString("ID_STYLE_RD_DEFAULT"), 11, RGB(237, 28, 36), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_ITALIC, 0)
    oOutput.EndParagraph()

    emptyrow(1)

    oOutput.EndSection()

    //******************************************************************************************
    //14. Prilozi
    //******************************************************************************************
    oOutput.BeginSection(false, Constants.SECTION_DEFAULT)
    globalHeader(oOutput)
    globalFooter(oOutput)

    oOutput.BeginParagraphF(getString("HEADING_1"))
    oOutput.Output("14. Prilozi", getString("ID_STYLE_RD_DEFAULT"), 13, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_LEFT, 0)
    oOutput.EndParagraph()

    oOutput.BeginParagraphF(getString("ID_STYLE_RD_DEFAULT"))
    oOutput.Output("[Napomena – Upisati tekst]", getString("ID_STYLE_RD_DEFAULT"), 11, RGB(237, 28, 36), Constants.C_TRANSPARENT, Constants.FMT_LEFT | Constants.FMT_ITALIC, 0)
    oOutput.EndParagraph()

    emptyrow(1)

    oOutput.EndSection()
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
    oOutput.OutputLn("eFiskalizacija - MVP1", getString("ID_DEFAULT_FONT"), 21, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0)
    oOutput.OutputLn("UC01 Evidentiraj račun  (specifikacija)", getString("ID_DEFAULT_FONT"), 21, RGB(0, 0, 0), Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0)
    oOutput.EndParagraph()

    oOutput.EndSection()

}

function tableOfContent() {

    oOutput.BeginSection(false, Constants.SECTION_INDEX);
    //Header
    // oOutput.BeginHeader();
    // oOutput.BeginTable(100, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_CENTER, 0);
    // oOutput.TableRow();
    // oOutput.TableCell("Content table header", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
    // oOutput.EndTable("Test", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT, 0);
    // oOutput.EndHeader();
    //Footer
    // oOutput.BeginFooter();
    // oOutput.BeginTable(100, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_CENTER, 0);
    // oOutput.TableRow();
    // oOutput.TableCell("Content table footer", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_CENTER | Constants.FMT_VCENTER, 0);
    // oOutput.EndTable("Test", 100, "Segoe UI", 10, Constants.C_BLACK, Constants.C_WHITE, 0, Constants.FMT_LEFT, 0);
    // oOutput.EndFooter();
    //Sets the autonumbering on
    oOutput.SetAutoTOCNumbering(false);
    //Defines the style for the four levels in the TOC
    oOutput.SetTOCFormat(0, "Segoe UI", 12, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_LEFT | Constants.FMT_BOLD, 0, 0, 0, 0);
    oOutput.SetTOCFormat(1, "Segoe UI", 11, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_LEFT | Constants.FMT_BOLD, 5, 5, 2, 2);
    oOutput.SetTOCFormat(2, "Segoe UI", 11, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_LEFT | Constants.FMT_ITALIC, 10, 5, 2, 2);
    oOutput.SetTOCFormat(3, "Segoe UI", 11, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_LEFT, 15, 5, 2, 2);
    //Output text
    emptyrow(1)
    oOutput.OutputLn("SADRŽAJ:", "Segoe UI", 12, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_LEFT | Constants.FMT_BOLD, 0);
    //Adds the table of content
    emptyrow(2)
    oOutput.OutputField(Constants.FIELD_TOC, "Segoe UI", 10, Constants.C_BLACK, Constants.C_WHITE, Constants.FMT_LEFT);
    oOutput.EndSection();
}