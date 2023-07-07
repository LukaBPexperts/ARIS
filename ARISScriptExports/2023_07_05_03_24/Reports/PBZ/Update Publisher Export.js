//****************CUSTOMIZE**********************
var serverAddress = "Server Address"  //unesi adresu publisher servera
var exportName = "Export name"        // unesi ime Publish exporta
//**********************************************


var oOutput = Context.createOutputObject()


if(serverAddress != null && exportName != null){

var businessPublisherComp = Context.getComponent("BusinessPublisher")
businessPublisherComp.activateExport(serverAddress, exportName, false)

var aResult = businessPublisherComp.updateExport(ArisData.getActiveDatabase(), serverAddress, exportName)

businessPublisherComp.activateExport(serverAddress, exportName, true)

oOutput.OutputLn("Duration=" + aResult.getDuration(), "Tahoma", 18,0,-1,1,0)
oOutput.OutputLn("Success=" + aResult.isSuccessfull(), "Tahoma", 18,0,-1,1,0)

oOutput.OutputLn("ServerAddress=" + serverAddress, "Tahoma", 18,0,-1,1,0)
oOutput.OutputLn("ExportName=" + exportName, "Tahoma", 18,0,-1,1,0)


}

oOutput.WriteReport()