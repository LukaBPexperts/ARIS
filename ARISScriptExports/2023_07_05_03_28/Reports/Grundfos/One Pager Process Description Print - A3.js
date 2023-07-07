function main(){
    var lReportComponent = Context.getComponent("Report");
        
    var lReportExecData=lReportComponent.createExecInfo ("30084210-9c4a-11e7-3732-28f10e35e86c"
                                                       , ArisData.getSelectedModels() 
                                                       , Context.getSelectedLanguage()
                                                       ,Context.getSelectedFormat(),Context.getSelectedFile());
                                                       
    lReportExecData.setProperty("pageSize","A3");
   
    var lReportResult=lReportComponent.execute(lReportExecData);
              
    var lResultFileData = lReportResult.getResultFileData()
    var maxLength=0;
    var resultIndex=-1;
    for (var i in lResultFileData){
        var lReportData=lReportResult.getResultFileData()[i].getData();
        if (lReportData.length>maxLength){
            resultIndex=i
            maxLength=lReportData.length;
        }
    }
    if(resultIndex>-1){
        var lReportData=lReportResult.getResultFileData()[resultIndex].getData();            
        Context.addOutputFile(Context.getSelectedFile(), lReportData);         
    }
}
main();