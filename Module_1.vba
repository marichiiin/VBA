Sub PivotUpdate()
'
' PivotUpdate Macro
' Macro recorded 12/15/2009 by a0213227
'

' Update all pivot table data.
    'Sheets("KEP_WarePivot").PivotTables("PivotTable1").PivotCache.Refresh
    'Sheets("WagoPivot").PivotTables("PivotTable1").PivotCache.Refresh
    'Sheets("IntellutionAAPivot").PivotTables("PivotTable1").PivotCache.Refresh
    'Sheets("IntellutionDAPivot").PivotTables("PivotTable1").PivotCache.Refresh
    'Sheets("AnalogPivot").PivotTables("PivotTable1").PivotCache.Refresh
    'Sheets("DiscretePivot").PivotTables("PivotTable1").PivotCache.Refresh
    'Sheets("CntlLogixPivot").PivotTables("PivotTable1").PivotCache.Refresh
    
Sheets("IO List").Select
Dim c As Object
Dim matchFoundIndex As Long

For Each c In Sheets("IO List").Range("C:C")
    If c.Value <> "" Then
        matchFoundIndex = WorksheetFunction.Match(Cells(c.Row, 3), Range("C:C"), 0)
        If matchFoundIndex <> c.Row Then
            MsgBox "Duplicate Tagname: " & Cells(matchFoundIndex, 3).Value
            Exit Sub
            Exit For
        End If
   End If
Next

Sheets("WWPivot").PivotTables("PivotTable1").PivotCache.Refresh

End Sub


Sub ExportCSV()

        
    Application.ScreenUpdating = False
    
    FileName = Range("D2").Text
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(FileName, True)
        
   'Range("A4").Select
   'Set rng = Range(Selection, Selection.End(xlDown))
   'Count = Application.WorksheetFunction.CountA(rng)
    iCount = CInt(Range("D3").Value)
           
    Range("A4:Q4").Select
    Set Rng = Range(Selection, Selection.End(xlDown))
    
    iCount = iCount + 4
    s = ""
    
    For r = 4 To iCount
        For c = 1 To 46
            If c = 46 Then
                s = s & """" & Cells(r, c) & """"
            Else
                
                If c = 30 And r > 4 And Cells(r, c) = "" Then
                    s = s & Chr(34) & Chr(34) & ","
                Else
                    If c = 30 And r > 4 Then
                        s = s & Chr(34) & Chr(34) & Cells(r, c) & Chr(34) & Chr(34) & ","
                    Else
                        s = s & """" & Cells(r, c) & ""","
            End If
            End If
            End If
        Next c
    
    a.writeline s
    s = ""
    
    Next r
    
    Range("A5").Select
    Application.ScreenUpdating = True
End Sub

Sub ExportCntlLogix()
       
  Application.ScreenUpdating = False
    
    FileName = Range("D2").Text
    sRange = Range("C2").Text
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(FileName, True)
        
   'Range("A4").Select
   'Set rng = Range(Selection, Selection.End(xlDown))
   'Count = Application.WorksheetFunction.CountA(rng)
    iCount = CInt(Range("D3").Value)
    iCCount = CInt(Range("e3").Value)
           
    Range(sRange).Select
    Set Rng = Range(Selection, Selection.End(xlDown))
    
    iCount = iCount + 10
    s = ""
    
    For r = 4 To iCount
        For c = 1 To iCCount
            If c = iCCount Then
'                s = s & """" & Cells(r, c) & """"
                s = s & Cells(r, c)
            Else
                
                If c = 30 And r > 4 And Cells(r, c) = "" Then
'                    s = s & Chr(34) & Chr(34) & ","
                    s = s & ","
                Else
                    If c = iCCount And r > 4 Then
'                        s = s & Chr(34) & Chr(34) & Cells(r, c) & Chr(34) & Chr(34) & ","
                        s = s & Cells(r, c) & ","
                    Else
'                        s = s & """" & Cells(r, c) & ""","
                        s = s & Cells(r, c) & ","
            End If
            End If
            End If
        Next c
    
    a.writeline s
    s = ""
    
    Next r
    
    Range("A5").Select
    Application.ScreenUpdating = True
End Sub


Sub Export_AWX_A()
  Application.ScreenUpdating = False
    
    FileName = Range("D2").Text
    sRange = Range("C2").Text
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(FileName, True)
        
   'Range("A4").Select
   'Set rng = Range(Selection, Selection.End(xlDown))
   'Count = Application.WorksheetFunction.CountA(rng)
    iCount = CInt(Range("D3").Value)
    iCCount = CInt(Range("L3").Value)
           
    Range(sRange).Select
    Set Rng = Range(Selection, Selection.End(xlDown))
    
    iCount = iCount + 4
    s = ""
    
    For r = 4 To iCount
        For c = 1 To iCCount
            If c = iCCount Then
'                s = s & """" & Cells(r, c) & """"
                s = s & Cells(r, c)
            Else
                
                If c = 30 And r > 4 And Cells(r, c) = "" Then
'                    s = s & Chr(34) & Chr(34) & ","
                    s = s & ","
                Else
                    If c = iCCount And r > 4 Then
'                        s = s & Chr(34) & Chr(34) & Cells(r, c) & Chr(34) & Chr(34) & ","
                        s = s & Cells(r, c) & ","
                    Else
'                        s = s & """" & Cells(r, c) & ""","
                        s = s & Cells(r, c) & ","
            End If
            End If
            End If
        Next c
    
    a.writeline s
    s = ""
    
    Next r
    
    Range("A5").Select
    Application.ScreenUpdating = True
End Sub
