
Sub readResults()

    '     This is intended to be used in another workbook named "WheelData.xlsm"
    '     Use this to sum all of the results in ArchivedResults.csv
    '     -> useful for creating a stacked percentage chart.
  
On Error GoTo finishUp

With Application
    .DisplayAlerts = False
    .ScreenUpdating = False
End With
    
    Dim wheeldata As Workbook
    Dim datasheet As Worksheet
    Dim archivePath, archivefilepath, archiveStr As String
    Dim values, totalrows As Integer
    Dim theheader As Boolean
    
    Set wheeldata = Workbooks("wheelData")
    Set datasheet = wheeldata.Worksheets("Sheet1")
    
    ' Change this to the folder containing ArchivedResults.csv
    archivePath = "C:\path\to\folder\"
    
    archivefilepath = archivePath & "ArchivedResults.csv"
      
    With datasheet

        Columns("A:L").Delete
        
        DoEvents
        
        values = 0
        
        For Each cell In .Range("$A$1:$L$1")
            values = values + 1
            cell.Value = values
            .Range(cell.Address).Offset(1, 0) = 0
            DoEvents
        Next cell
        
        DoEvents

        Open archivefilepath For Input As #1
    
        theheader = True
        
        Do Until EOF(1)
        
            totalrows = .Range("A1").End(xlDown).Row
                
            Debug.Print "totalrows: " & totalrows
                
            Line Input #1, archiveStr
            Debug.Print archiveStr
            
            If theheader <> True Then
            
                For Each cell In .Range("$A$1:$L$1")
                
                    If cell.Value = archiveStr Then
                        .Range(cell.Address).Offset(totalrows, 0) = _
                            .Range(cell.Address).Offset(totalrows - 1, 0) + 1
                    Else
                        .Range(cell.Address).Offset(totalrows, 0) = _
                            .Range(cell.Address).Offset(totalrows - 1, 0)
                    End If
                    
                Next cell
            Else
                theheader = False
            End If
    
        Loop
        
        Close #1

        DoEvents
        
        noerrors = True
        
        .Rows("2:2").Delete Shift:=xlUp
        
        DoEvents
        
        .ListObjects.Add(xlSrcRange, Range("$A$1:$L$101"), , xlYes).Name = _
            "Table3"
        
    End With

finishUp:

Application.ScreenUpdating = True

If noerrors <> True Then MsgBox "Something is wrong! Did you remember to change archivePath?"

End Sub
