Sub CopySpecificAttendees()
    
   
'Sub GetAttendeeList()
 
Dim objApp As Outlook.Application
Dim objItem As Object
Dim objAttendees As Outlook.Recipients
Dim objAttendeeReq As String
Dim objAttendeeOpt As String
Dim objOrganizer As String
Dim dtStart As Date
Dim dtEnd As Date
Dim strSubject As String
Dim strLocation As String
Dim strNotes As String
Dim strMeetStatus As String
Dim strCopyData As String
Dim strCount  As String
Dim strDistribution As String
Dim strPresent As String


'On Error Resume Next
 
Set objApp = CreateObject("Outlook.Application")
Set objItem = Outlook.Application.ActiveExplorer.Selection.Item(1)
Set objAttendees = Outlook.Application.ActiveExplorer.Selection.Item(1).Recipients

'MsgBox objItem.Subject

For x = 1 To objAttendees.Count
    If objAttendees(x).Type = olRequired Then
        strPresent = strPresent & objAttendees(x).Name & vbNewLine
    Else
    
        strDistribution = strDistribution & objAttendees(x).Name & vbNewLine
    End If
Next

Dim objPPT As Object, _
    PPTPrez As PowerPoint.Presentation, _
    pSlide As PowerPoint.Slide

Set objPPT = CreateObject("PowerPoint.Application")
objPPT.Visible = True

Set PPTPrez = objPPT.Presentations.Open("C:\Users\310267217\Documents\IB\Template meeting log.pptx")
Set pSlide1 = PPTPrez.Slides(1)
Set pSlide2 = PPTPrez.Slides(2)

Set oTitle = pSlide1.Shapes("Title")
oTitle.TextFrame2.TextRange.Characters.Text = objItem.Subject

Set oDate = pSlide1.Shapes("Date")
oDate.TextFrame2.TextRange.Characters.Text = Date


Set oWhenWhere = pSlide2.Shapes("Present")
oWhenWhere.TextFrame2.TextRange.Characters.Text = strPresent

Set oDistribution = pSlide2.Shapes("Distribution")
oDistribution.TextFrame2.TextRange.Characters.Text = strDistribution

Set oLocation = pSlide2.Shapes("WhenWhere")
oLocation.TextFrame2.TextRange.Characters.Text = Date & " " & objItem.Location

Set oObjective = pSlide2.Shapes("Objective")
oObjective.TextFrame2.TextRange.Characters.Text = objItem.Body

Set objApp = Nothing
Set objItem = Nothing
Set objAttendees = Nothing

'End Sub
 '   Dim olSel As Selection
 '   Dim olItem As AppointmentItem
  '  Dim olAttendees As Recipients
  '  Dim obj As Object
  '  Dim strAddrs As String
  '   Dim DataObj As MSForms.DataObject
 
   '  Set DataObj = New MSForms.DataObject
  '  Set olSel = Outlook.Application.ActiveExplorer.Selection
  '  Set olItem = olSel.Item(1)
 
  '  For Each obj In olAttendees
        'To copy the attendees who have accepted the meeting request
'  MsgBox obj.Name
 ' MsgBox obj.Type
 ' MsgBox olAttendees.Item(0).
 '       If obj.MeetingResponseStatus = olResponseAccepted Then
        'To copy who declined - "If olAttendee.MeetingResponseStatus = olResponseDeclined Then"
        'To copy who haven't respond - "If olAttendee.MeetingResponseStatus = olResponseNone" Then
 '          strAddrs = strAddrs & ";" & obj.Address
          '  DataObj.SetText strAddrs
          '  DataObj.PutInClipboard
 '       End If
 '   Next
End Sub
