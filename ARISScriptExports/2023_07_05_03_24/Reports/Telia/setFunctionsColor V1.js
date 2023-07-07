var nLocale = Context.getSelectedLanguage()
var models = ArisData.getSelectedModels()

main()

function main() {
	for (var i = 0; i < models.length; i++) {

		var funcionsOcc = models[i].ObjOccListFilter(Constants.OT_FUNC)

		for (var j = 0; j < funcionsOcc.length; j++) {
			var funcionDef = funcionsOcc[j].ObjDef()
            var attTypeOfFunctionValue = funcionDef.Attribute(ArisData.ActiveFilter().UserDefinedAttributeTypeNum(AT_TYPE_OF_FUNCTION), nLocale).GetValue(false)
            
            funcionsOcc[j].setFillColor(RGB(41,138,37))
            
			//promijeni atribut u kojem gleda ovaj podatak !!!
			if (attTypeOfFunctionValue == "Automatic") {
				funcionsOcc[j].setFillColor(RGB(0, 153, 204))
           }
           
           if (attTypeOfFunctionValue == "Not available") {
				funcionsOcc[j].setFillColor(RGB(153, 153, 153))
           }
           
           if (attTypeOfFunctionValue == "Manual" || attTypeOfFunctionValue == "Semi manual" ) {
				funcionsOcc[j].setFillColor(RGB(41,138,37))
           }
		}
	}
}

Dialogs.MsgBox("Done!")

function RGB(r, g, b) {
	return (new java.awt.Color(r / 255.0, g / 255.0, b / 255.0, 1)).getRGB() & 0xFFFFFF
}