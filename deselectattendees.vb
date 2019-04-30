
Sub deselect()

Dim objApp As Outlook.Application

    Dim objItem As Object
    Dim objAttendees As Outlook.Recipients     

    Set objApp = CreateObject("Outlook.Application")
    Set objItem = objApp.ActiveInspector.CurrentItem
    Set objAttendees = objItem.Recipients  

    For x = 1 To objAttendees.Count
    On Error Resume Next
            If objAttendees(x).Type = 2 Or objAttendees(x).Type = 3 Then
                objAttendees(x).Sendable = False
               '  MsgBox (objAttendees(x).Type)
            End If
    Next  

    Set objApp = Nothing
    Set objItem = Nothing
    Set objAttendees = Nothing

End Sub

