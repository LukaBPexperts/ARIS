function main(){
    var g_ooutfile = Context.createOutputObject();
    var db = ArisData.getActiveDatabase();
    var locale = Context.getSelectedLanguage();
    var objOccs = ArisData.getSelectedObjOccs();
    var filter = ArisData.ActiveFilter();
    
    for(var i=0; i<objOccs.length; i++){
        var objOcc = objOccs[i];
        g_ooutfile.OutputTxt(objOcc.ObjDef().Name(locale)+" - Symbol: "+objOcc.SymbolNum()+" ("+objOcc.SymbolName()+")\n");
    }
    
    g_ooutfile.WriteReport()
}

main();