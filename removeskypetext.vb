
Sub removeskypetext()

Dim objApp As Outlook.Application
    Dim objItem As Object
    
    Set objApp = CreateObject("Outlook.Application")
    Set objItem = objApp.ActiveInspector.CurrentItem
    
    Find = "Trouble"
    rt = objItem.Body
    MsgBox rt
    't = Replace(rt, vbCrLf, "")
    'objItem.Body = "Hi, " & vbCrLf & vbCrLf & "Hope this time slot works." & vbCrLf
    
    
    If 1 = 1 Then
            st = InStr(rt, Find)
            ' MsgBox st
            Find = "original message"
            eind = InStr(rt, Find)
            'MsgBox eind
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
                Set objSel = objDoc.Windows(1).Selection
                objSel.Text = "Hi," & vbCrLf & vbCrLf
                'strLink = "http://www.outlookcode.com"
                
                objDoc.Hyperlinks.Add objSel.Range, strLink, "", "", strLinkText, ""
                
           End If
    
           ' objItem.Body = newv
    End If

