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

var g_CurrentRowName = "";
var g_CurrentSheetStartName = "";

main();

function main() {
    initializeOutput();
    startTableSheet();
    
    newTableRow();
    writeHeader(getString("REPORT_TITLE"), 7);
    newTableRow();
    writeDescription(getString("REPORT_DESCRIPTION"), 7);
    newTableRow();
    writeTableCell("");
    newTableRow();
    
    writeTableHeaderCell(getString("COLUMN_USERGROUP_NAME"), 40);
    writeTableHeaderCell(getString("COLUMN_LICENSE_PRIVILEGES"), 40);
	writeTableHeaderCell(getString("COLUMN_LICENSE_ASSIGNED_SEATS"), 30);
    writeTableHeaderCell(getString("COLUMN_FUNCTION_PRIVILEGES"), 40);
    writeTableHeaderCell(getString("COLUMN_LAST_NAME"), 30);
    writeTableHeaderCell(getString("COLUMN_FIRST_NAME"), 30);
    writeTableHeaderCell(getString("COLUMN_EMAIL_ADDRESS"), 50);
    writeTableHeaderCell(getString("COLUMN_USER_NAME"), 30);
    
    var selectedUsergroupNames = getSelectionList();
    
    var usergroups = UMC.getAllUsergroups();
    
    var isEvenCell = true;
    
    for(var i = 0 ; i < usergroups.size() ; i++) {
        isEvenCell = !isEvenCell;
        var usergroup = usergroups.get(i);
        var name = usergroup.getName();
        
        if(!selectedUsergroupNames.isEmpty() && !selectedUsergroupNames.contains(name)) {
            continue;
        }
        
        g_CurrentRowName = name;
		if(i == 0) {
			g_CurrentSheetStartName = name;
		}
        
        newTableRow();
        
        var licensePrivileges = UMC.getPrivilegesForUsergroup(usergroup, false, true);
        var functionPrivileges = UMC.getPrivilegesForUsergroup(usergroup, true, false);
        var assignedUsers = UMC.getAssignedUsersForUsergroup(usergroup);
        
        writeTableCell(name, 40, isEvenCell);
        var rows = getMax(licensePrivileges.size(), functionPrivileges.size(), assignedUsers.size());
        for(var j = 0 ; j < rows ; j++) {
            if(j != 0) {
                newTableRow();
                writeTableCell("", 40, isEvenCell);
            }
            if(licensePrivileges.size() > j) {
                writeTableCell(licensePrivileges.get(j).getDisplayName(), 40, isEvenCell);  
				writeTableCell(licensePrivileges.get(j).getAssignedSeats(), 30, isEvenCell);  				
            } else {				
                writeTableCell("", 40, isEvenCell); 
				writeTableCell("", 30, isEvenCell);				
            }
            if(functionPrivileges.size() > j) {
                writeTableCell(functionPrivileges.get(j).getDisplayName(), 40, isEvenCell);    
            } else {
                writeTableCell("", 40, isEvenCell);    
            }
            if(assignedUsers.size() > j) {
                var user = assignedUsers.get(j);
                writeTableCell(user.getLastName(), 30, isEvenCell);
                writeTableCell(user.getFirstName(), 30, isEvenCell);  
                writeTableEmailCell(user.getEmail(), 50, isEvenCell);
                writeTableCell(user.getName(), 30, isEvenCell);  
            } else {
                writeTableCell("", 30, isEvenCell);   
                writeTableCell("", 30, isEvenCell); 
                writeTableCell("", 40, isEvenCell); 
                writeTableCell("", 30, isEvenCell); 
            }
        }   
    }
    endTableSheet(getIndexedSheetName(g_CurrentSheetStartName, g_CurrentRowName));
    g_Output.WriteReport();
}

function newTableRow() {
    if(!addTableRow()) {
        var sheetName = getIndexedSheetName(g_CurrentSheetStartName, g_CurrentRowName);
        endTableSheet(sheetName);
        startTableSheet();
        g_CurrentSheetStartName = g_CurrentRowName;
        addTableRow();
    }
}

