Sub openmail()
    
    Set oOutlook = GetObject(, "Outlook.Application")
    Set NS = oOutlook.GetNamespace("MAPI")
    NS.Logon
    Set msg = NS.GetItemFromID(ActiveCell.Value)
    msg.display
    Set objInsp = msg.GetInspector
    Set objDoc = objInsp.WordEditor
    Application.Wait (Now + TimeValue("0:00:02"))
    Set myRange = objDoc.Range(Start:=ActiveCell.Offset(0, 5), End:=ActiveCell.Offset(0, 6))
    myRange.Select
    
    
    Set oOutlook = Nothing
    Set NS = Nothing
    Set msg = Nothing
    Set objInsp = Nothing
    Set objDoc = Nothing
    Set objsel = Nothing
   
End Sub
