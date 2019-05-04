Sub openmail()
    'MsgBox ActiveCell.Value
    Set oOutlook = GetObject(, "Outlook.Application")
    Set NS = oOutlook.GetNamespace("MAPI")
    NS.Logon
    Set msg = NS.GetItemFromID(ActiveCell.Value)
    'MsgBox msg.Subject
    msg.display
End Sub
