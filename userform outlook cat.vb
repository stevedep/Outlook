
Private Sub addToMail_Click()
 Dim obApp As Object
    Dim NewEmail As MailItem
    Set obApp = Outlook.Application
    'Set NewEmail = obApp.ActiveInspector.CurrentItem
    'Set NewEmail = obApp.Selection.Item(1)
    Set myOlExp = Application.ActiveExplorer
    myOlExp.Activate
 Set myOlsel = myOlExp.Selection
    'MsgBox myOlsel.Item(1).Subject
    c = myOlsel.Item(1).Categories
    'MsgBox c
    'If you want to set a specific category to the new email manually
    'You can use the following line instead to show the Category dialog
    'NewEmail.ShowCategoriesDialog
  For intCurrentRow = 0 To lstCategories.ListCount - 1
    If lstCategories.Selected(intCurrentRow) Then
        stritems = stritems & lstCategories.Column(0, _
        intCurrentRow) & ";"
    End If
 Next intCurrentRow
    
'Dim individualItem As Object
'MsgBox myOlsel.Count
'For Each individualItem In myOlsel
   myOlsel.Item(1).Categories = c & "; " & stritems
  ' individualItem.Categories = c & "; " & stritems
  'MsgBox individualItem.Subject
'Next
  
   
    'myOlsel.Item(1).Save
    Set obApp = Nothing
    Set NewEmail = Nothing
    UserForm1.Hide
    UserForm1.Show
    TextBox1.SetFocus
    TextBox1.Text = ""
   ' cmdSave.SetFocus
End Sub


Private Sub cmdDone_Click()
 Dim myNameSpace As Outlook.NameSpace
 Dim myInbox As Outlook.Folder
 Dim myDestFolder As Outlook.Folder
 Dim myItems As Outlook.Items
 Dim myItem As Object
 
 Set myNameSpace = Application.GetNamespace("MAPI")
 Set myInbox = myNameSpace.GetDefaultFolder(olFolderInbox)
Dim obApp As Object
    Dim NewEmail As MailItem

    Set obApp = Outlook.Application
    Set myOlExp = Application.ActiveExplorer
    myOlExp.Activate
    Set myOlsel = myOlExp.Selection
    
    myOlsel.Item(1).Move myInbox.Folders("Done")
    
'    myOlsel.Item(1).Save
    Set myOlExp = Nothing
    Set myOlsel = Nothing
    Set obApp = Nothing
    Set NewEmail = Nothing
    UserForm1.Hide
    UserForm1.Show
    UserForm1.TextBox1.SetFocus
    UserForm1.TextBox1.Text = ""
    
End Sub

Private Sub KeyHandler_KeyDown(KeyCode As Integer, _
     Shift As Integer)
    Dim intShiftDown As Integer, intAltDown As Integer
    Dim intCtrlDown As Integer
 
If KeyCode = vbKeyReturn Then MsgBox ("enter")
' Use bit masks to determine which key was pressed.
    intShiftDown = (Shift And acShiftMask) > 0
    intAltDown = (Shift And acAltMask) > 0
    intCtrlDown = (Shift And acCtrlMask) > 0
    intEnter = acCtrlMask > 0
    ' Display message telling user which key was pressed.
    If intShiftDown Then MsgBox "You pressed the Shift key."
     If intEnter Then MsgBox "You pressed the enter key."
    If intAltDown Then MsgBox "You pressed the Alt key."
    If intCtrlDown Then MsgBox "You pressed the Ctrl key."
End Sub

Private Sub cmdRemoveToDo_Click()
    Dim obApp As Object
    Dim NewEmail As MailItem
    Set obApp = Outlook.Application
    Set myOlExp = Application.ActiveExplorer
    myOlExp.Activate
    Set myOlsel = myOlExp.Selection
        c = myOlsel.Item(1).Categories
        'MsgBox c
        c = Replace(c, ", ToDo", "")
        c = Replace(c, "ToDo,", "")
        'MsgBox c
        myOlsel.Item(1).Categories = c
        myOlsel.Item(1).Save
     
    Set obApp = Nothing
    Set NewEmail = Nothing
    UserForm1.Hide
    UserForm1.Show
    UserForm1.TextBox1.SetFocus
    
    TextBox1.SetFocus
   
End Sub

Private Sub cmdSave_Click()
 Dim obApp As Object
    Dim NewEmail As MailItem

    Set obApp = Outlook.Application
    Set myOlExp = Application.ActiveExplorer
    myOlExp.Activate
    Set myOlsel = myOlExp.Selection
    
    myOlsel.Item(1).Save
    Set obApp = Nothing
    Set NewEmail = Nothing
    Set myOlsel = Nothing
    UserForm1.Hide
    UserForm1.Show
    UserForm1.TextBox1.SetFocus
    
End Sub


Private Sub cmdMeeting_Click()
 Dim myNameSpace As Outlook.NameSpace
 Dim myInbox As Outlook.Folder
 Dim myDestFolder As Outlook.Folder
 Dim myItems As Outlook.Items
 Dim myItem As Object
 
 Set myNameSpace = Application.GetNamespace("MAPI")
 Set myInbox = myNameSpace.GetDefaultFolder(olFolderInbox)
 'MsgBox olFolderInbox
Dim obApp As Object
    Dim NewEmail As MailItem

    Set obApp = Outlook.Application
    Set myOlExp = Application.ActiveExplorer
    myOlExp.Activate
    Set myOlsel = myOlExp.Selection
    
    myOlsel.Item(1).Move myInbox.Folders("Meetings")
    
'   myOlsel.Item(1).Save
    Set myOlExp = Nothing
    Set myOlsel = Nothing
    Set obApp = Nothing
    Set NewEmail = Nothing
    Set myNameSpace = Nothing
    Set myInbox = Nothing

    UserForm1.Hide
    UserForm1.Show
    UserForm1.TextBox1.SetFocus
End Sub

Private Sub cmdMultiple_Click()
 Dim obApp As Object
    Dim NewEmail As MailItem
    Set obApp = Outlook.Application
    Set myOlExp = Application.ActiveExplorer
    myOlExp.Activate
 Set myOlsel = myOlExp.Selection
    c = myOlsel.Item(1).Categories

  For intCurrentRow = 0 To lstCategories.ListCount - 1
    If lstCategories.Selected(intCurrentRow) Then
        stritems = stritems & lstCategories.Column(0, _
        intCurrentRow) & ";"
    End If
 Next intCurrentRow
    
Dim individualItem As Object

For Each individualItem In myOlsel
   'myOlsel.Item(1).Categories = c & "; " & stritems
   individualItem.Categories = c & "; " & stritems
    individualItem.Save
Next
    
   
    'myOlsel.Item(1).Save
    Set obApp = Nothing
    Set NewEmail = Nothing
    Set obApp = Nothing
    Set myOlExp = Nothing
    UserForm1.Hide
    UserForm1.Show
    UserForm1.TextBox1.SetFocus
    
End Sub

Private Sub addCategoryAndSave_Click()
 Dim obApp As Object
    Dim NewEmail As MailItem

    Set obApp = Outlook.Application
    Set myOlExp = Application.ActiveExplorer
    myOlExp.Activate
 Set myOlsel = myOlExp.Selection
    c = myOlsel.Item(1).Categories
  For intCurrentRow = 0 To lstCategories.ListCount - 1
    If lstCategories.Selected(intCurrentRow) Then
        stritems = stritems & lstCategories.Column(0, _
        intCurrentRow) & ";"
    End If
 Next intCurrentRow
   
   myOlsel.Item(1).Categories = c & "; " & stritems
        myOlsel.Item(1).Save
    
    Set obApp = Nothing
    Set NewEmail = Nothing
    UserForm1.Hide
    UserForm1.Show
    UserForm1.TextBox1.SetFocus
    TextBox1.Text = ""

End Sub

Private Sub CommandButton1_Click()
 Dim obApp As Object
 
    Set obApp = Outlook.Application
 
 MsgBox obApp.ActiveWindow
End Sub


Private Sub lstCategories_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'MsgBox KeyCode
    'bij enter toevoegen en bewaren cat
    If KeyCode = vbKeyReturn Then addCategoryAndSave_Click
    'bij ctrl toevoegen
    If KeyCode = 17 Or KeyCode = 16 Then addToMail_Click
    If KeyCode = 32 Then cmdSave_Click
    If KeyCode > 64 Then
    TextBox1.SetFocus
    TextBox1.Text = LCase(ChrW(KeyCode))
    End If
    
End Sub

Private Sub TextBox1_Change()
    Set oOutlook = GetObject(, "Outlook.Application")
    Set ns = oOutlook.GetNamespace("MAPI")
    lstCategories.Clear
    For Each objCategory In ns.Categories
         If LCase(objCategory.Name) Like "*" & LCase(TextBox1.Text) & "*" Then
             UserForm1.lstCategories.AddItem objCategory.Name
         End If
     Next
    Set oOutlook = Nothing
    Set ns = Nothing
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Set myOlExp = Application.ActiveExplorer
 
'MsgBox KeyCode
 
 If KeyCode = 13 Then lstCategories.ListIndex = 0
 If KeyCode = 37 Then cmdDone_Click
 If KeyCode = 39 Then
    'MsgBox "meeting"
    cmdMeeting_Click
 End If
 If KeyCode = 18 Then cmdRemoveToDo_Click
 If KeyCode = 38 Then
    myOlExp.Activate
    SendKeys "{UP}"
    UserForm1.Hide
    UserForm1.Show
    UserForm1.TextBox1.SetFocus
 End If

If KeyCode = 40 Then
    myOlExp.Activate
    SendKeys "{DOWN}"
    UserForm1.Hide
    UserForm1.Show
    UserForm1.TextBox1.SetFocus
 End If

Set myOlExp = Nothing
End Sub

