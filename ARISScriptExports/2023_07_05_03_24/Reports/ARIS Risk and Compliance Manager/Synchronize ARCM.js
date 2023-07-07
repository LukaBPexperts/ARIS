/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

// This script call the synchronizeARCM from "Synchronie ARCM - Library" with flag started from ABA
 
try {    
    synchronizeARCM(false);        
}    
catch(ex) {
    errorHandling(ex);
}
