var serverAdmin = Context.getComponent("ServerAdmin")

var selectedDatabases = ArisData.getSelectedDatabases()
	
var reorg = serverAdmin.reorganiseArisDatabase(selectedDatabases[0], null)
