var test = 0;
 function myDialog()
 {
  // all member functions except for getPages can access the property "dialog" of the dialog class. Type is "UserDialog" (see help).
  // examples:
  // - Get the page with the specified index (e.g. 0): this.dialog.getPage(0)
  // - Define the list of visible wizard pages (e.g. 0,1): this.dialog.setActiveWizardPages([0,1])

  // returns DialogTemplate[]
  // non-optional
  this.getPages = function()
  {
      var iDialogTemplate1 = Dialogs.createNewDialogTemplate(600, 200, "First page");
      iDialogTemplate1.PushButton(35, 5, 80, 15, "Press", "BUTTON_1");
      iDialogTemplate1.CheckBox(120, 5, 100, 15, "CheckMe", "CHECKBOX_1");
      iDialogTemplate1.ComboBox(10, 60, 100, 40, ["Value 1", "Value 3", "Value 2", "Value 4"], "COMBO_1");
      iDialogTemplate1.DropListBox(355, 60, 100, 40, ["sorted 1", "sorted 3", "sorted 2", "sorted 4"], "COMBO_4", 3); //sorted + editable
      iDialogTemplate1.Text(10, 170, 100, 16, "stat. Text");
      iDialogTemplate1.TextBox(120, 170, 100, 16, "TXT_EDIT_1");
      var iDialogTemplate2 = Dialogs.createNewDialogTemplate(400, 200, "Second page");
      iDialogTemplate2.Text(10, 10, 100, 16, "stat. Text");
      iDialogTemplate2.TextBox(10, 40, 100, 16, "TXT_EDIT_2");
      iDialogTemplate2.TextBox(10, 80, 150, 16, "TXT_EDIT_3");
      
      var iDialogTemplate3 = Dialogs.createNewDialogTemplate(400, 200, "Treca page");
      iDialogTemplate3.TextBox(10, 80, 150, 16, "TXT_EDIT_4");
      
      return [iDialogTemplate1, iDialogTemplate2, iDialogTemplate3];
  }

 // initialize dialog pages (are already created and pre-initialized with static data from XML or template)
 // parameter: Array of DialogPage
 // see Help: DialogPage
 // user can set control values
 // optional
 this.init = function(aPages)
 {
    //use this function also to store the page data locally (for example to access it in "onClose")
    aPages[0].setFocusedElement("TXT_EDIT_1");
    aPages[0].getDialogElement("TXT_EDIT_1").setText("Start value 1");

    aPages[1].getDialogElement("TXT_EDIT_2").setText("Start value 2");
    aPages[1].getDialogElement("TXT_EDIT_3").setText("Start value text 3");
    aPages[1].setFocusedElement("TXT_EDIT_2");
    
    aPages[2].getDialogElement("TXT_EDIT_4").setText("Ante");
 }

 // returns true if the page is in a valid state. In this case "Ok", "Finish", or "Next" is enabled.
 // called each time a dialog value is changed by the user (button pressed, list selection, text field value, table entry, radio button,...)
 // pageNumber: the current page number, 0-based
 this.isInValidState = function(pageNumber)
 {
    return true;
 }

 // called when the page is displayed
 // pageNumber: the current page number, 0-based
 // optional
 this.onActivatePage = function(pageNumber)
 {
 }

 // returns true if the "Finish" or "Ok" button should be visible on this page.
 // pageNumber: the current page number, 0-based
 // optional. if not present: always true
 this.canFinish = function(pageNumber)
 {
    return true;
 }

 // returns true if the user can switch to another page.
 // pageNumber: the current page number, 0-based
 // optional. if not present: always true
 this.canChangePage = function(pageNumber)
 {
    return true;
 }

 // returns true if the user can switch to next page.
 // called when the "Next" button is pressed and thus not suitable for activation/deactivation of this button
 // can prevent the display of the next page
 // pageNumber: the current page number, 0-based
 // optional. if not present: always true
 this.canGotoNextPage = function(pageNumber)
 {
    return true;
 }

 // returns true if the user can switch to previous page.
 // called when the "Back" button is pressed and thus not suitable for activation/deactivation of this button
 // can prevent the display of the previous page
 // pageNumber: the current page number, 0-based
 // optional. if not present: always true
 this.canGotoPreviousPage = function(pageNumber)
 {
    return true;
 }

 // called after "Ok"/"Finish" has been pressed and the current state data has been applied
 // can be used to update your data
 // pageNumber: the current page number
 // bOK: true=Ok/finish, false=cancel pressed
 // optional
 this.onClose = function(pageNumber, bOk)
 {
 }

 // the result of this function is returned as result of Dialogs.showDialog(). Can be any object.
 // optional
 this.getResult = function()
 {
    return this.dialog.getPage(1).getDialogElement("TXT_EDIT_2").getText()
 }

 // other methods (all optional):
 // - [ControlID]_pressed(),
 // - [ControlID]_focusChanged(boolean lost=false, gained=true)
 // - [ControlID]_changed() for TextBox and DropListBox
 // - [ControlID]_selChanged(int newSelection)
 // - [ControlID]_cellEdited(row, column) for editable tables, row and column are 0-based
 this.BUTTON_1_pressed = function()
 {
    this.dialog.getPage(0).getDialogElement("TXT_EDIT_1").setVisible(false)
    this.dialog.getPage(1).getDialogElement("TXT_EDIT_2").setEnabled(false)
    test = 1;
 }
}

 var result = Dialogs.showDialog(new myDialog(), Constants.DIALOG_TYPE_WIZARD, "Multi page dialog");
 Dialogs.MsgBox(result)