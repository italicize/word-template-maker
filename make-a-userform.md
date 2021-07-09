# Make a LongInputBox function

Sometimes you need a LongInputBox function that accepts 32,000 characters, including line breaks, because the [InputBox function](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/inputbox-function) accepts only 254 characters.

1. Make a new UserForm:
   1. Open Word and open a new blank document.
   1. Click the **View** menu, click **Macros**, and click **Edit**.
   1. In the project pane, click **Normal**. \
      (Or click the template where you save macros, if not Normal.dotm.)
   1. Click the **Insert** menu and click **UserForm**.
1. Add controls to the UserForm:
   1. In the toolbox, click the **Label** control. \
      Click and drag to outline a rectangle at the top of the UserForm.
   1. Click the **TextBox** control. \
      Click and drag to outline a rectangle in the middle of the UserForm.
   1. Click the **CommandButton** control. \
      Click and drag to outline a rectangle in the lower left of the UserForm.
   1. Click the **CommandButton** control again. \
      Click and drag to outline a rectangle in the lower right of the User Form.
   1. Close the toolbox.
1. Name the UserForm and controls:
   1. In the properties pane, for Name type **frmLongInputBox**.
   1. Click the label. \
      For its Name, type **lblPrompt**.
   1. Click the text box. \
      For its Name, type **txtInput**. \
      For EnterKeyBehavior, select **True**. \
      For MultiLine, select **True**.
   1. Click the lower-left command button. \
      For its Name, type **cmdCancel**. \
      For Cancel, select **True**. \
      For its Caption, type **Cancel**.
   1. Click the lower-right command button. \
      For Name, type **cmdContinue**. \
      For its Caption, type **Continue**.
1. Add code to the UserForm: \
   In the project pane, right-click **frmLongInputBox** and select **View Code**. \
   Type or paste this code:

```vb
Private Sub cmdCancel_Click()
    Me.Tag = 0
    Me.Hide
End Sub

Private Sub cmdContinue_Click()
    Me.Tag = 1
    Me.Hide
End Sub
```

1. Save the function: \
   In the project pane, double-click a module. \
   Click in the code pane and type or paste this code:

```vb
Public Function LongInputBox(ByVal strPrompt As String, _
    ByVal strTitle As String) As String
    On Error GoTo The_End:
    Dim objForm As frmLongInputBox
    Set objForm = New frmLongInputBox
    With objForm
        'Changes the prompt and title on the form.
        .lblPrompt.Caption = strPrompt
        .Caption = strTitle
        .Show
        LongInputBox = .txtInput.Value
        'If Cancel is clicked, returns no string.
        If .Tag = "0" Then LongInputBox = ""
    End With
    Unload objForm
    Set objForm = Nothing
The_End:
End Function
```

Click the **File** menu and click **Save Normal**.

# Save shortcuts to a macro

1. For add a button, right-click the quick access toolbar and select **Customize Quick Access Toolbar**.
   1. For "Choose commands from" select **Macros**.
   1. Click a macro and click **Add**.
   1. Click **Modify**, click an image, and click **OK**.
1. To add a keyboard shortcut, right-click the menu ribbon and select **Customize the Ribbon**.
   1. Next to "Keyboard shortcuts" click **Customize**.
   1. For "Categories" select **Macros**.
   1. Click a macro.
   1. Click the box for "Press new shortcut key."
   1. Type a shortcut, such as **Alt+.**.
   1. Click **Assign**, click **Close**, and click **OK**.

