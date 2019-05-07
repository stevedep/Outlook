Sub outlookweava()
    
    toelichting = InputBox("Toelichting")

    Dim objApp As Outlook.Application
    Dim objItem As Object

    Set objApp = CreateObject("Outlook.Application")
    Set objItem = objApp.ActiveInspector.CurrentItem
    
    'MsgBox objItem.EntryID
    
    ' Excel Application, workbook, and sheet object

    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Object
    ' Filename
    Dim fileDoesExist As Boolean
    Dim FileName As String

     Set objInsp = objItem.GetInspector
     Set objDoc = objInsp.WordEditor
     Set objsel = objDoc.Windows(1).Selection

    ' Create Excel Application
    on error resume next
    Set xlApp = GetObject(, "Excel.Application")
        If Err <> 0 Then
            Set xlApp = CreateObject("Excel.Application")
        End If
        
    xlApp.Visible = True
    
    FileName = "outlook.xlsm"
    
    Set xlBook = xlApp.Workbooks(FileName)
    
    If xlBook Is Nothing Then
          fileDoesExist = Dir("C:\Users\310267217\OneDrive - Philips\Outlook Weava\" & FileName) > ""
        ' Check for existing file
        If fileDoesExist Then
            ' Open Excel file
            Set xlBook = xlApp.Workbooks.Open("C:\Users\310267217\OneDrive - Philips\Outlook Weava\" & FileName)
            Set xlSheet = xlBook.Sheets(1)
        Else
            ' Add Excel file
            Set xlBook = xlApp.Workbooks.Add
            With xlBook
                .Title = "All Sales"
                .Subject = "Sales"
                .SaveAs FileName:="C:\Users\310267217\OneDrive - Philips\Outlook Weava\" & FileName
            End With
        End If
    End If
    
  
    Set xlSheet = xlBook.Sheets(1)
    
    ' Do stuff with Excel workbook
    With xlApp
        With xlBook
            ' Add Excel VBA code to update workbook here
            Dim tbl As ListObject
            Set tbl = xlSheet.ListObjects("Tabel1")
            'Set tbl = Range("Tabel1").ListObject
            Set newrow = tbl.ListRows.Add(AlwaysInsert:=True)
            newrow.Range(1, 1).Value = objItem.EntryID
            newrow.Range(1, 2).Value = objItem.Subject
            newrow.Range(1, 3).Value = objItem.Sender
            newrow.Range(1, 4).Value = Format(objItem.SentOn, "MMM d, yyyy")
            newrow.Range(1, 5).Value = objsel
            newrow.Range(1, 6).Value = objsel.Start
            newrow.Range(1, 7).Value = objsel.End
            newrow.Range(1, 8).Value = toelichting
            
            '.Close SaveChanges:=True
        End With
    End With
    
    'xlApp.Quit
    
    
    Set objInsp = Nothing
    Set objDoc = Nothing
    Set objsel = Nothing
    Set xlApp = Nothing
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set newrow = Nothing
    Set tbl = Nothing
    Set objApp = Nothing
    Set objItem = Nothing
End Sub
