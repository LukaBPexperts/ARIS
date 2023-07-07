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
    writeHeader(getString("REPORT_TITLE"), 9);
    newTableRow();
    writeDescription(getString("REPORT_DESCRIPTION"), 9);
    newTableRow();
    writeTableCell("");
    newTableRow();
    
    writeTableHeaderCell(getString("COLUMN_USERNAME"), 30);
    writeTableHeaderCell(getString("COLUMN_EMAIL_ADDRESS"), 50);
    writeTableHeaderCell(getString("COLUMN_LAST_NAME"), 30);
    writeTableHeaderCell(getString("COLUMN_FIRST_NAME"), 30);
    writeTableHeaderCell(getString("COLUMN_USERGROUPS"), 40);
    writeTableHeaderCell(getString("COLUMN_LICENSE_PRIVILEGES"), 40);
    writeTableHeaderCell(getString("COLUMN_LICENSE_PRIVILEGES_USER"), 40);
    writeTableHeaderCell(getString("COLUMN_FUNCTION_PRIVILEGES"), 40);
    writeTableHeaderCell(getString("COLUMN_FUNCTION_PRIVILEGES_USER"), 40);
    
    var selectedUserNames = getSelectionList();
    
    var users = UMC.getAllUsers();
    
    var isEvenRow = true;
    
    for(var i = 0 ; i < users.size() ; i++) {
        
        isEvenRow = !isEvenRow;
        var user = users.get(i);
        var username = user.getName();
        
        if(!selectedUserNames.isEmpty() && !selectedUserNames.contains(username)) {
            continue;
        }
		
		g_CurrentRowName = username;
		if(i == 0) {
			g_CurrentSheetStartName = username;
		}
           
        var email = user.getEmail();
        var lastname = user.getLastName();
        var firstname = user.getFirstName();
        var usergroups = UMC.getAssignedUsergroupNamesForUser(user);
        
        var allLicensePrivileges = new java.util.ArrayList();
        var userLicensePrivileges = new java.util.ArrayList();
        var allFunctionPrivileges = new java.util.ArrayList();
        var userFunctionPrivileges = new java.util.ArrayList();
        
        var allPrivileges = UMC.getPrivilegesForUser(user, true, true, true);
        if(allPrivileges.size() > 0) {
            for(var j = 0 ; j < allPrivileges.size() ; j++) {
                var next = allPrivileges.get(j);
                if(next.isLicensePrivilege()) {
                    allLicensePrivileges.add(next);
                } else {
                    allFunctionPrivileges.add(next);
                }
            }
            var allDirectPrivileges = UMC.getPrivilegesForUser(user, true, true, false);
            for(var j = 0 ; j < allDirectPrivileges.size() ; j++) {
                var next = allDirectPrivileges.get(j);
                if(next.isLicensePrivilege()) {
                    userLicensePrivileges.add(next);
                } else {
                    userFunctionPrivileges.add(next);
                }
            }
        }
        
        newTableRow();
        writeTableCell(username, 30, isEvenRow);
        writeTableEmailCell(email, 50, isEvenRow);
        writeTableCell(lastname, 30, isEvenRow);
        writeTableCell(firstname, 30, isEvenRow);
        var rows = getMax(usergroups.size(), allLicensePrivileges.size(), userLicensePrivileges.size(), allFunctionPrivileges.size(), userFunctionPrivileges.size());
        
        if(rows == 0) {
            writeTableCell("", 40, isEvenRow); 
            writeTableCell("", 40, isEvenRow); 
            writeTableCell("", 40, isEvenRow); 
            writeTableCell("", 40, isEvenRow); 
            writeTableCell("", 40, isEvenRow); 
        } else {
            for(var j = 0 ; j < rows ; j++) {
                if(j != 0) {
                    newTableRow();
                    writeTableCell("", 30, isEvenRow);
                    writeTableCell("", 50, isEvenRow);
                    writeTableCell("", 30, isEvenRow);
                    writeTableCell("", 30, isEvenRow);
                }
                if(usergroups.size() > j) {
                    writeTableCell(usergroups.get(j), 40, isEvenRow);    
                } else {
                    writeTableCell("", 40, isEvenRow);    
                }
                if(allLicensePrivileges.size() > j) {
                    writeTableCell(allLicensePrivileges.get(j).getDisplayName(), 40, isEvenRow);    
                } else {
                    writeTableCell("", 40, isEvenRow);    
                }
                if(userLicensePrivileges.size() > j) {
                    writeTableCell(userLicensePrivileges.get(j).getDisplayName(), 40, isEvenRow);    
                } else {
                    writeTableCell("", 40, isEvenRow);    
                }
                if(allFunctionPrivileges.size() > j) {
                    writeTableCell(allFunctionPrivileges.get(j).getDisplayName(), 40, isEvenRow);    
                } else {
                    writeTableCell("", 40, isEvenRow);    
                }
                if(userFunctionPrivileges.size() > j) {
                    writeTableCell(userFunctionPrivileges.get(j).getDisplayName(), 40, isEvenRow);    
                } else {
                    writeTableCell("", 40, isEvenRow);    
                }
            }   
        }
    }
    endTableSheet(getIndexedSheetName(g_CurrentSheetStartName, g_CurrentRowName));
    g_Output.WriteReport();
}

function newTableRow() {
    if(!addTableRow()) {
        endTableSheet(getIndexedSheetName(g_CurrentSheetStartName, g_CurrentRowName));
        startTableSheet();
        g_CurrentSheetStartName = g_CurrentRowName;
        addTableRow();
    }
}