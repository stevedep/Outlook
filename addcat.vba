Sub add_cat()


  Dim objApp As Outlook.Application
    Set oOutlook = GetObject(, "Outlook.Application")
    Set ns = oOutlook.GetNamespace("MAPI")

    
    For Each objCategory In ns.Categories
        UserForm1.lstCategories.AddItem objCategory.Name
    Next
    UserForm1.Show
        
    Set oOutlook = Nothing
    Set ns = Nothing
End Sub
