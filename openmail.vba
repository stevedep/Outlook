Sub openmail()

    'MsgBox ActiveCell.Value
    Set oOutlook = GetObject(, "Outlook.Application")
    Set NS = oOutlook.GetNamespace("MAPI")
    NS.Logon
    Set msg = NS.GetItemFromID(ActiveCell.Value)
    'MsgBox msg.Subject
    msg.display
    Set objInsp = msg.GetInspector
    Set objDoc = objInsp.WordEditor
  '  Set objSel = objDoc.Selection
    Set myRange = objDoc.Range(Start:=ActiveCell.Offset(0, 5), End:=ActiveCell.Offset(0, 6))
    'objDoc.Range.Start = ActiveCell.Offset(0, 5)
    'objDoc.Range.End = ActiveCell.Offset(0, 6)
    myRange.Select
    Set oOutlook = Nothing
    Set NS = Nothing
    Set msg = Nothing
    Set objInsp = Nothing
    Set objDoc = Nothing
    Set objsel = Nothing
   
End Sub
