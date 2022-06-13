# 20220605_2158
 20220605_2158


 ## final


### on calc sheet
```vba
Private Sub Worksheet_Change(ByVal Target As Range)

If Intersect(Target, Range("A21")) Is Nothing Then
    'pass or ActiveWorkbook.Connections("Query - filter_prov").Refresh
Else
	With ActiveWorkbook
	    For lCnt = 1 To .Connections.Count
	      'Excludes PowerPivot and other connections
	      If .Connections(lCnt).Type = xlConnectionTypeOLEDB Then
	        .Connections(lCnt).Refresh
	      End If
	    Next lCnt
	End With
End If

End Sub


```

### workbook
```vba
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'run before save

' disable bg refresh
With ActiveWorkbook
    For lCnt = 1 To .Connections.Count
      'Excludes PowerPivot and other connections
      If .Connections(lCnt).Type = xlConnectionTypeOLEDB Then
        .Connections(lCnt).OLEDBConnection.BackgroundQuery = False
      End If
    Next lCnt
End With

'ActiveWorkbook.Connections("Query - test table").Refresh

With ActiveWorkbook
    For lCnt = 1 To .Connections.Count
      'Excludes PowerPivot and other connections
      If .Connections(lCnt).Type = xlConnectionTypeOLEDB Then
        .Connections(lCnt).Refresh
      End If
    Next lCnt
End With

Application.EnableEvents = False
ThisWorkbook.Save
Application.EnableEvents = True

MsgBox "Data refreshed and saved"

' re-enable bg refresh
With ActiveWorkbook
    For lCnt = 1 To .Connections.Count
      'Excludes PowerPivot and other connections
      If .Connections(lCnt).Type = xlConnectionTypeOLEDB Then
        .Connections(lCnt).OLEDBConnection.BackgroundQuery = True
      End If
    Next lCnt
End With



End Sub
 
```
