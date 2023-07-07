/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */


//Context.writePrivateProfileInt("mySection", "myInt1", 10, "myFile.xml");
//Context.writePrivateProfileString("mySection", "myString1", "Test", "myFile.xml");

var numberOfUnits = Context.getPrivateProfileString("QuadrantConfiguration", "numberOfUnits", 0, "quadrants configuration.xml");
var attributeName = Context.getPrivateProfileString("QuadrantConfiguration", "configuration", "Default", "quadrants configuration.xml");
var useRGB = Context.getPrivateProfileString("QuadrantConfiguration", "useRGB", "Default", "quadrants configuration.xml");

var unitCount = parseInt(numberOfUnits);

var mapString = "";

for (var i=1; i<=unitCount; i++) {
    mapString = mapString + Context.getPrivateProfileString("QuadrantConfiguration", "attributeVal" + i, "Default", "quadrants configuration.xml") + ":";
    mapString = mapString + Context.getPrivateProfileString("QuadrantConfiguration", "AssignedColor" + i, "Default", "quadrants configuration.xml") + "#";
}

Context.setProperty("mapString", mapString);
Context.setProperty("useRGB", useRGB);
Context.setProperty("attributeName", attributeName);

//Dialogs.MsgBox(attributeName + " - " + numberOfUnits);
