// selected language (localeId)
var g_nloc = 0; 
// declare the output Object
var g_ooutfile = null; 

function main()
{
    
    // create and initialize the output Object and variable for selected language
    g_ooutfile = Context.createOutputObject(Context.getSelectedFormat(), Context.getSelectedFile());
    g_nloc = __toLong(Context.getSelectedLanguage());
    g_ooutfile.Init(g_nloc);
    
    // definitions of used stylesheets
    g_ooutfile.DefineF("REPORT1", "Arial", 24, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_BOLD | Constants.FMT_CENTER, 0, 21, 0, 0, 0, 1);
    g_ooutfile.DefineF("REPORT2", "Times New Roman", 14, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0, 21, 0, 0, 0, 1);
    
    
    var selectedGroups = ArisData.getSelectedGroups()
    
    var objDefs = new Array()
    for(var i=0 ; i < selectedGroups.length; i++){
        var tempObjDefs = selectedGroups[i].ObjDefList(true, null)
        objDefs = objDefs.concat(tempObjDefs)
    }
    
    var oselectedobjdefs1 = objDefs; 
    var oselectedobjdefs1obj = null; 
    var i1 = 0; 

    g_ooutfile.BeginTable(100, Constants.C_BLACK, Constants.C_TRANSPARENT, Constants.FMT_LEFT, 0);
    g_ooutfile.TableRow();
    g_ooutfile.TableCell("Name", 50, "Arial", 10, Constants.C_BLACK, vbRgb(192, 192, 192), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    g_ooutfile.TableCell("Type", 50, "Arial", 10, Constants.C_BLACK, vbRgb(192, 192, 192), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    g_ooutfile.TableCell("Path", 50, "Arial", 10, Constants.C_BLACK, vbRgb(192, 192, 192), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    
    g_ooutfile.TableCell("Creation date", 50, "Arial", 10, Constants.C_BLACK, vbRgb(192, 192, 192), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    g_ooutfile.TableCell("Last change date", 50, "Arial", 10, Constants.C_BLACK, vbRgb(192, 192, 192), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    g_ooutfile.TableCell("Description", 50, "Arial", 10, Constants.C_BLACK, vbRgb(192, 192, 192), 0, Constants.FMT_BOLD | Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
    
    for( i1 = 0; i1 <= (oselectedobjdefs1.length - 1); i1++ ){
        oselectedobjdefs1obj = oselectedobjdefs1[i1];
        // iterate through oOccList1
        var oocclist1 = null; 
        var oocclist1obj = null; 
        var i2 = 0; 
        oocclist1 = oselectedobjdefs1obj.OccList();
        if (oocclist1.length == 0) {
            g_ooutfile.TableRow();
            g_ooutfile.TableCell(oselectedobjdefs1obj.Name(g_nloc), 50, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
            g_ooutfile.TableCell(oselectedobjdefs1obj.Type(), 50, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
            g_ooutfile.TableCell(oselectedobjdefs1obj.Group().Path(g_nloc), 50, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
            
            g_ooutfile.TableCell(oselectedobjdefs1obj.getTimeOfCreation(), 50, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
            g_ooutfile.TableCell(oselectedobjdefs1obj.getTimeOfLastModification(), 50, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
            g_ooutfile.TableCell(oselectedobjdefs1obj.Attribute(Constants.AT_DESC, g_nloc).getValue(), 50, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT | Constants.FMT_VTOP | Constants.FMT_EXCELMODIFY, 0);
        }
        
        
        oocclist1 = null;
        
        oselectedobjdefs1obj = null;
    }
    
    g_ooutfile.EndTable("", 100, "Arial", 10, Constants.C_BLACK, Constants.C_TRANSPARENT, 0, Constants.FMT_LEFT, 0);
    oselectedobjdefs1 = null;
    
    g_ooutfile.WriteReport(Context.getSelectedPath(), Context.getSelectedFile());
}



main();
