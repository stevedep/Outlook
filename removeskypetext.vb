Sub removeskypetext()

    Dim objApp As Outlook.Application
    Dim objItem As Object
    
    Set objApp = CreateObject("Outlook.Application")
    Set objItem = objApp.ActiveInspector.CurrentItem
    
    Find = "Trouble"
    rt = objItem.Body
    
    
    If 1 = 1 Then
            st = InStr(rt, Find)
            Find = "original message"
            eind = InStr(rt, Find)
       
            If st > 0 Then
                repl = Mid(rt, st, eind - st + 20)
                newv = Replace(rt, repl, "")
                newv = Replace(newv, "!OC([1033])!]", "")
                stl = InStr(newv, "<")
                eindl = InStr(newv, ">")
                Link = Mid(newv, stl + 1, eindl - stl - 1)
                strLink = Link
                strLinkText = "Join Skype Meeting here"
                Set objInsp = objItem.GetInspector
                Set objDoc = objInsp.WordEditor
                Set objsel = objDoc.Windows(1).Selection
                objsel.Text = "Hi," & vbCrLf & vbCrLf & "Hope this timeslot works." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Kind regards," & vbCrLf & vbCrLf & "Steve"
                Set myRange = objDoc.Range(Start:=objsel.End - 24, End:=objsel.End - 23)
                objsel.Hyperlinks.Add myRange, strLink, "", "", strLinkText, ""
           End If
    End If
    
    Set objApp = Nothing
    Set objItem = Nothing
End Sub
