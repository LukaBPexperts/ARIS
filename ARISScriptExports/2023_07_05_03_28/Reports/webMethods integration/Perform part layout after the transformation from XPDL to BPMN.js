/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

main();

function main()
{   
    var model = ArisData.getSelectedModels()[0];
    
    var webMethodsIntegrationComponent = Context.getComponent("webMethodsIntegration");
         
    doCxnLayout(model);
        
    var dependingModels = webMethodsIntegrationComponent.getDependingModels(model);
        
    if(dependingModels != null)
    {
        for(var i=0; i < dependingModels.length;i++)
        {
            doCxnLayout(dependingModels[i]);
            doPoolLaneLayout(dependingModels[i]);
        }
    }
    
    doPoolLaneLayout(model);
}            
            
function doPoolLaneLayout(mainModel)
{
    var poolArray = [];

	var objects = mainModel.ObjOccList();
    
    for(var q = 0 ; q < objects.length ; q++)
	{
		if(objects[q] != null && (objects[q].SymbolNum() == Constants.ST_BPMN_POOL_1))
		{
			poolArray.push(objects[q]);
		}
	}
    for (var r = 0 ; r < poolArray.length ; r++){
        var poolLaneArray = new Array();
        poolLaneArray.push(poolArray[r]);
        var oPool = poolArray[r].ObjDef();
        //collect Lanes and Pools
        for(var q = 0 ; q < objects.length ; q++)
        {
            if(objects[q] != null && (objects[q].SymbolNum() == Constants.ST_BPMN_LANE_1 && checkBelongsToPool(oPool,objects[q].ObjDef())))
            {
                poolLaneArray.push(objects[q]);
            }
        }
    
        mainModel.doLayout(poolLaneArray);
        /*
        //correct Pools and Lanes
        for(var j = 0 ; j < poolLaneArray.length ; j++)
        {
            poolLaneArray[j].set3DEffect(false);
    
            if(poolLaneArray[j].SymbolNum() ==  1531)
            {
                poolLaneArray[j].SetPosition(poolLaneArray[j].X() + (19.5 * (274/72)), poolLaneArray[j].Y());
                poolLaneArray[j].SetSize(poolLaneArray[j].Width() - (19.5 * (274/72)), poolLaneArray[j].Height());
            }
        }
        */
    }    

	
}

function checkBelongsToPool(p_oPool,p_oLane){
    var oPools = p_oLane.getConnectedObjs([Constants.OT_BPMN_POOL], Constants.EDGES_OUT, [Constants.CT_BELONGS_TO_1]);
    for (var i = 0; i < oPools.length; i++) {
        if (oPools[i] != null && oPools[i].equals(p_oPool)) return true;
    }    
    return false;
}    
function doCxnLayout(model)
{
	var cxn = model.CxnOccList();

	if(cxn != null)
	{
		model.doLayout(cxn);
	}
}            