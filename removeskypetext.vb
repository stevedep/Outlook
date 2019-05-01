
Sub removeskypetext()

Dim objApp As Outlook.Application
    Dim objItem As Object
    
    Set objApp = CreateObject("Outlook.Application")
    Set objItem = objApp.ActiveInspector.CurrentItem
    
    Find = "Trouble"
    rt = objItem.Body
    st = InStr(rt, Find)
    ' MsgBox st
    Find = "original message"
    eind = InStr(rt, Find)
    'MsgBox eind
    repl = Mid(rt, st, eind - st + 20)
    newv = Replace(rt, repl, "")
    
    objItem.Body = newv
    
    Set objApp = Nothing
    Set objItem = Nothing
End Sub
