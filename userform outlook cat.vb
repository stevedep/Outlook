Private Sub CommandButton1_Click()
 Dim obApp As Object
    Dim NewEmail As MailItem

    Set obApp = Outlook.Application
    Set NewEmail = obApp.ActiveInspector.CurrentItem
    c = NewEmail.Categories
    'If you want to set a specific category to the new email manually
    'You can use the following line instead to show the Category dialog
    'NewEmail.ShowCategoriesDialog
  For intCurrentRow = 0 To lstCategories.ListCount - 1
    If lstCategories.Selected(intCurrentRow) Then
        stritems = stritems & lstCategories.Column(0, _
        intCurrentRow) & ";"
    End If
 Next intCurrentRow
    
    NewEmail.Categories = c & "; " & stritems

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
