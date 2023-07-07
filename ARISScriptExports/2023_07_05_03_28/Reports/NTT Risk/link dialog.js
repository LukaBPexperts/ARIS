var g_requestType = Dialogs.showDialog(new requestTypeDialog(), Constants.DIALOG_TYPE_ACTION, "Create Link");

var a=2

///////////////////////////////////////////////////////////////////////////////////////////////////////
// First Dialog to get the the Request Type
///////////////////////////////////////////////////////////////////////////////////////////////////////
function requestTypeDialog() {

    var oResult = null;

    this.getPages = function() {

        var dlgTemplateDocument = Dialogs.createNewDialogTemplate(440, 220, "Please copy link?");
        dlgTemplateDocument.GroupBox(15, 20, 440, 220, "Share model");
        dlgTemplateDocument.OptionGroup("OPT_REQUEST_TYPE");
        
        dlgTemplateDocument.Text(165, 40, 200, 20, "Link for Designers");
        dlgTemplateDocument.TextBox(30, 55, 400, 20, "Link_Designers", 0);
        
        dlgTemplateDocument.Text(160, 100, 200, 20, "Link for Collaborators");
        dlgTemplateDocument.TextBox(30, 115, 400, 20, "Link_Collaborators", 0);
        
        dlgTemplateDocument.Text(170, 160, 200, 20, "Link for Viewers");
        dlgTemplateDocument.TextBox(30, 175, 400, 20, "Link_Viewers", 0);

        //Check which pages should be showed based on the selected Function Type
        var listPagesToShow = [];
        listPagesToShow.push(dlgTemplateDocument);

        return listPagesToShow;

    }

    this.onClose = function(pageNumber, bOk) {
        if (bOk)
            oResult = "abcrec"
        
         var selection =java.awt.datatransfer.StringSelection("abc")
        var clipboard = java.awt.Toolkit.getDefaultToolkit().getSystemClipboard()
        
        
        
        clipboard.setContents(selection, selection);
    }


    this.getResult = function() {

        return oResult;
    }
    
    
    this.onActivatePage = function(pageNumber) {

        
        
        
        this.dialog.getPage(pageNumber).getDialogElement("Link_Designers").setText("http://desktop-dv8s6lr/#default/item/c.process.United%20Motor%20Group.BGz502wSEeUSGQBQVp1Emw.-1/~AfBbImNvbXBvbmVudF9jb21tZW50U3RyZWFtMTc2Il0~");
        
        

        //this.dialog.getPage(pageNumber).getDialogElement("Link_Designers").setEnabled(false)
        
    }

 
 
 
 
}