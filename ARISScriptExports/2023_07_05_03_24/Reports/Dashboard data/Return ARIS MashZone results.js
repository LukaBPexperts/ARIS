/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

Context.setProperty("excel-formulas-allowed", false); //default value is provided by server setting (tenant-specific): "abs.report.excel-formulas-allowed"

var schedulerComponent = Context.getComponent("ReportScheduler")

var sScheduleGUID = Context.getProperty("task-guid")
var sUserName     = Context.getProperty("user-name")
var sUserPwdEnc   = Context.getProperty("user-password")
var reportResult = schedulerComponent.getDecryptedScheduleResult(sScheduleGUID, sUserName, sUserPwdEnc, true)

if(reportResult.length>0)
{
    var aFileData = schedulerComponent.getZipEntries(reportResult)
    for(var i=0; i<aFileData.length; i++)
    {
        var sFileName = "" + aFileData[i].getName();
        if( sFileName=="report.result" )
            sFileName = Context.getSelectedFile()
        
        Context.addOutputFile(sFileName, aFileData[i].getData())
    }
}
