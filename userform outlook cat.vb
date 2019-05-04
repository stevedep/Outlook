Private Sub CommandButton1_Click()
 Dim obApp As Object
    Dim NewEmail As MailItem

    Set obApp = Outlook.Application
    'Set NewEmail = obApp.ActiveInspector.CurrentItem
    'Set NewEmail = obApp.Selection.Item(1)
    Set myOlExp = Application.ActiveExplorer
 Set myOlsel = myOlExp.Selection
    'MsgBox myOlsel.Item(1).Subject
    c = myOlsel.Item(1).Categories
    'If you want to set a specific category to the new email manually
    'You can use the following line instead to show the Category dialog
    'NewEmail.ShowCategoriesDialog
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
End Sub

Private Sub CommandButton2_Click()
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
    Set myOlsel = myOlExp.Selection
    myOlsel.Item(1).Move myInbox.Folders("Done")
    
'    myOlsel.Item(1).Save
    Set myOlExp = Nothing
    Set myOlsel = Nothing
    Set obApp = Nothing
    Set NewEmail = Nothing
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
