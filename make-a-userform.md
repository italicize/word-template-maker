# Make a LongInputBox function

Sometimes you need a LongInputBox function that accepts 32,000 characters, including line break, because the [InputBox function](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/inputbox-function) accepts only 254 characters.

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
1. Add the code: \
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
    Dim strInput As String, strTag As String
    Dim objForm As frmLongInputBox
    Set objForm = New frmLongInputBox
    With objForm
        .lblPrompt.Caption = strPrompt
        .Caption = strTitle
        .Show
        strInput = .txtInput.Value
        strTag = .Tag
    End With
    Unload objForm
    Set objForm = Nothing
    If strTag = "0" Or strInput = "" Then
        LongInputBox = ""
    Else
        LongInputBox = strInput
    End If
End Function
```

Click the **File** menu and click **Save Normal**.



