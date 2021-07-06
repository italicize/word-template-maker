# Make an all-purpose UserForm

1. Open Word and open a new blank document.
1. Click the **View** menu, click **Macros**, and click **Edit**.
1. In the project pane, click **Normal**.
1. Click the **Insert** menu and click **UserForm**.
1. In the toolbox, click the **Label** control. 
1. Click and drag to outline a rectangle at the top of the UserForm.
1. Click the **TextBox** control.
1. Click and drag to outline a rectangle in the middle of the UserForm.
1. Click the **CommandButton** control.
1. Click and drag to outline a rectangle in the lower left of the UserForm.
1. Click the **CommandButton** control again.
1. Click and drag to outline a rectangle in the lower right of the User Form.
1. Click the header of the UserForm.
1. In the properties pane, type a name for the UserForm, like **frmGeneralUserForm**.
1. Click the label.
1. For its Name, type **Label**.
1. Click the text box.
1. For itsName, type **txtInput**.
1. For EnterKeyBehavior, select **True**.
1. For MultiLine, select **True**.
1. Click the lower-left command button.
1. For its Name, type **cmdCancel**.
1. For Cancel, select **True**.
1. For its Caption, type **Cancel**.
1. Click the lower-right command button.
1. For Name, type **cmdContinue**.
1. For its Caption, type **Continue**.
1. Right-click the lower-left button and select **View Code**.
1. Type `Me.Tag = 0: Me. Hide` in the blank line below Private Sub cmdCancel_Click().
1. From the drop-down menu above the code, select **cmdContinue**.
1. Type `Me.Tag = 1: Me. Hide` in the blank line below Private Sub cmdContinue_Click().
1. Close the toolbox.
1. Click the **File** menu and click **Save Normal**.



