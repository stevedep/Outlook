Sub removeskypetext()

Dim objApp As Outlook.Application
    Dim objItem As Object
    
    Set objApp = CreateObject("Outlook.Application")
    Set objItem = objApp.ActiveInspector.CurrentItem
    
    Find = "Trouble"
    rt = objItem.Body
    MsgBox rt
    st = InStr(rt, Find)
    ' MsgBox st
    Find = "original message"
    eind = InStr(rt, Find)
    'MsgBox eind
    If st > 0 Then
        repl = Mid(rt, st, eind - st + 20)
        newv = Replace(rt, repl, "")
        newv = Replace(newv, "!OC([1033])!]", "")
        newv = Replace(newv, "... vbCrLf", "")
        
        End If
    
    objItem.Body = newv
    
    Set objApp = Nothing
    Set objItem = Nothing
End Sub

