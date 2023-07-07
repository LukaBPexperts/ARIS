/**
 * Copyright (C) 2019 Software AG, Darmstadt, Germany and/or Software AG USA Inc., Reston, VA, USA,
 * and/or its subsidiaries and/or its affiliates and/or their licensors.
 *
 * Use, reproduction, transfer, publication or disclosure is prohibited
 * except as specifically provided for in your License Agreement with Software AG.
 *
 * Version: 10.0.11.0.1400208
 */

/**
 * Default layouter for models which were changed during the import.
 *
 * @author JURM
 */
function DefaultLayouter() {
    
    this.doLayout = function(modelsToLayout) {

        var modelIt = modelsToLayout.iterator();
        while (modelIt.hasNext()) {
            var model = modelIt.next();
            
            // bugfix anubis id: 326204
            try {
               // get the layout parameters for the model type
               var modelOptions = ArisData.getModelOptions();
               var layoutParameters = modelOptions.getLayoutParameters(model.TypeNum());
               if(layoutParameters != undefined && layoutParameters != null) {
                    //set layout strategy to hierarchy layout
                    layoutParameters.setStrategy(2);
               }
            } catch(e) {
                // ignore
            }
            // bugfix anubis id: 326204
            
            // apply the template of the model to all newly added items
            var template = model.getTemplate();
            model.setTemplate(template);
            // start the layouting algorithm
            model.doLayout();   
        }
    }
  
}

