Sub removeskypetext()

    Dim objApp As Outlook.Application
    Dim objItem As Object
    
    Set objApp = CreateObject("Outlook.Application")
    Set objItem = objApp.ActiveInspector.CurrentItem
    Set objAttendees = objItem.Recipients
    
    Find = "Trouble"
    rt = objItem.Body
    
    For x = 1 To objAttendees.Count
    If objAttendees(x).Type = olRequired Then
            naam = objAttendees(x).Name
            co = InStr(naam, ",")
            naam2 = Mid(naam, co + 2, Len(naam))
            sp = InStr(naam2, " ")
            
            If sp > 0 Then
                naam3 = Mid(naam2, 1, sp - 1)
            Else
                naam3 = naam2
            End If
               
               
           If naam3 <> "Steve" Then
                strNames = strNames & naam3 & ", "
            End If
    End If
    Next
    
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
                objsel.Text = "Hi " & strNames & vbCrLf & vbCrLf & "Hope this timeslot works." & vbCrLf & _
                "----------------------------------------------------------------------------" & vbCrLf & vbCrLf & vbCrLf & _
                "----------------------------------------------------------------------------" & _
                vbCrLf & vbCrLf & "Kind regards," & vbCrLf & vbCrLf & "Steve"
                Set myRange = objDoc.Range(Start:=objsel.End - 101, End:=objsel.End - 100)
                objsel.Hyperlinks.Add myRange, strLink, "", "", strLinkText, ""
           End If
    End If
    
    Set objApp = Nothing
    Set objItem = Nothing
End Sub
