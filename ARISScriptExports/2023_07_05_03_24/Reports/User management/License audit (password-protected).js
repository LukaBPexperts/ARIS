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
 
var g_Output    = Context.createOutputObject();
var g_Locale    = 1033;
var UMC         = Context.getComponent("UMC");

var licenseCodes = new java.util.TreeSet(); 

main();

function main() {
    initializeOutput();
    startTableSheet();
    writeHeaderSection();    
    fillLicenseCodes();
    writeTableHeaderCell(getString("COLUMN_DATE"), 20);
    var licenseIter = licenseCodes.iterator();
    while(licenseIter.hasNext()) {
        writeTableHeaderCell(licenseIter.next(), 15);
    }
    writeLicenseMetrics();
    g_Output.WriteReport();
}

function writeHeaderSection() {
    newTableRow();
    writeHeader(getString("REPORT_TITLE"), 14);
    writeHeader(formatstring1(getString("GENERATED_TIME"), currentTime()), 5);
    newTableRow();
    writeDescription(getString("REPORT_DESCRIPTION1"), 10);
    newTableRow();
    writeDescription(getString("REPORT_DESCRIPTION2"), 10);
    newTableRow();
    writeTableCell("");
    writeLastActivationTime();
    writeAggregationInterval();
    newTableRow();
    writeTableCell("");
    newTableRow();  
}

function writeLastActivationTime() {
    newTableRow();
    writeSubTitleLevelThree(monitoringStatusMessage(), 10);
}

function writeAggregationInterval() {
    newTableRow();
    var interval = UMC.getLicenseMonitoringInterval();
    writeSubTitleLevelThree(formatstring1(getString("INTERVAL_DESCRIPTION"), interval), 10);
}

function fillLicenseCodes() {
    var licenses = UMC.getAllLicenses();
    for(var i = 0 ; i < licenses.size() ; i++) {
        if (licenses.get(i).getType() != "TEN") {
            licenseCodes.add(licenses.get(i).getProduct().getCode());
        }
    }
}

function writeLicenseMetrics() {
    var dateProductUsageMap = UMC.getAllLicenseUsageMapByDay();
    var dates = dateProductUsageMap.keySet();
    var datesIter = dates.iterator();
    
    var isEven = false;
    while(datesIter.hasNext()) {
        newTableRow();
        var date = datesIter.next();
        writeTableCell(date, 20, isEven);
        
        var productUsageMap = dateProductUsageMap.get(date);
        var licenseIter = licenseCodes.iterator();
        while(licenseIter.hasNext()) {
            var license = licenseIter.next();           
            writeTableCell(productUsageMap.get(license), 15, isEven);           
        }
        isEven = !isEven;
    }
}

function monitoringStatusMessage() {
    var activationDate = UMC.getLicenseMonitoringActivationTime();   
    var monActive = UMC.isLicenseMonitoringActive();
    if (monActive) {
        /**if (activationDate != null) {
            return formatstring1(getString("ACTIVATE_TIME_CURRENTLY_ACTIVATE"), formatTime(activationDate));
        } else {*/
            return getString("MONITORING_ACTIVE");
        /**}*/
    } else {
        /**if (activationDate != null) {
              return formatstring1(getString("ACTIVATE_TIME_CURRENTLY_INACTIVE"), formatTime(activationDate)); 
        } else {*/
            return getString("MONITORING_INACTIVATED");
        /**}*/
    }
}

function newTableRow() {
    if(!addTableRow()) {
        endTableSheet(getIndexedSheetName(g_CurrentSheetStartName, g_CurrentRowName));
        startTableSheet();
        g_CurrentSheetStartName = g_CurrentRowName;
        addTableRow();
    }
}