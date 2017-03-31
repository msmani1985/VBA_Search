Private Sub Add_Web1_Click()
If Selection.Information(wdWithInTable) = True Then
'Selection.Tables(1).Rows(6).Select
Selection.InsertRowsBelow 1
Selection.InsertRowsBelow 1
Selection.InsertRowsBelow 1
'Selection.Tables(1).Rows(7).Select
'Selection.Tables(1).Rows(7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True

'MsgBox Selection.Cells(1).RowIndex

rowID = Selection.Cells(1).RowIndex
'Selection.Rows(1).Cells.Split NumColumns:=2
'Selection.Tables(1).Cell(1, 1).Select
'Selection.Tables(1).Cell(7, 1).Select


Selection.Tables(1).Rows(rowID).Select
Selection.Tables(1).Rows(rowID).Cells.Split NumColumns:=2, MergeBeforeSplit:=True

Selection.Tables(1).Cell(rowID, 1).Select
Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
'tt.SetPlaceholderText , , "Web Link"
tt.Range.Style = "Remediation_Link_Type"
tt.SetPlaceholderText Text:="Web Link"
''tt.Range.Style = "Normal"
tt.Range.Font.Size = "10"
tt.Range.Font.ColorIndex = wdBlack
tt.Range.Font.Bold = False
tt.Range.Rows.Height = 30
tt.Range.Font.Name = "Verdana"
tt.LockContentControl = True
tt.LockContents = True
Set tt = Nothing


Selection.Tables(1).Cell(rowID, 2).Select
Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
tt.Range.Style = "Web_Remediation_Text"
tt.Title = "Remediation Text"
'tt.SetPlaceholderText Text:="Enter remediation text"
tt.Range.Font.Size = "10"
tt.Range.Font.ColorIndex = wdBlack
tt.Range.Font.Bold = False
'tt.Range.Rows.Height = 30
tt.Range.Font.Name = "Verdana"
tt.LockContentControl = True
tt.LockContents = False
Set tt = Nothing

'Selection.Tables(1).Rows(rowID).Select

'Selection.InsertRowsBelow 1
'Selection.Tables(1).Rows(7).Select
'Selection.Tables(1).Rows(7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True

'MsgBox Selection.Cells(1).RowIndex

'rowID = Selection.Cells(1).RowIndex
'Selection.Rows(1).Cells.Split NumColumns:=2
'Selection.Tables(1).Cell(1, 1).Select
'Selection.Tables(1).Cell(7, 1).Select


Selection.Tables(1).Rows(rowID - 1).Select
Selection.Tables(1).Rows(rowID - 1).Cells.Split NumColumns:=2, MergeBeforeSplit:=True

Selection.Tables(1).Cell(rowID - 1, 2).Select
Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
tt.Range.Style = "Web_Remediation_Link"
tt.Title = "Remediation Link"
'tt.SetPlaceholderText Text:="Enter remediation link"
tt.Range.Font.Size = "10"
tt.Range.Font.ColorIndex = wdBlack
tt.Range.Font.Bold = False
'tt.Range.Rows.Height = 30
tt.Range.Font.Name = "Verdana"
tt.LockContentControl = True
tt.LockContents = False
Set tt = Nothing


'Selection.InsertRowsBelow 1
'Selection.Tables(1).Rows(7).Select
'Selection.Tables(1).Rows(7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True

'MsgBox Selection.Cells(1).RowIndex

'rowID = Selection.Cells(1).RowIndex
'Selection.Rows(1).Cells.Split NumColumns:=2
'Selection.Tables(1).Cell(1, 1).Select
'Selection.Tables(1).Cell(7, 1).Select


Selection.Tables(1).Rows(rowID - 2).Select
Selection.Tables(1).Rows(rowID - 2).Cells.Split NumColumns:=2, MergeBeforeSplit:=True


Selection.Tables(1).Cell(rowID - 2, 2).Select
Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
tt.Range.Style = "Web_Remediation_Tooltip"
tt.Title = "Remediation Tooltip"
'tt.SetPlaceholderText Text:="Enter remediation tooltip"
tt.Range.Font.Size = "10"
tt.Range.Font.ColorIndex = wdBlack
tt.Range.Font.Bold = False
'tt.Range.Rows.Height = 30
tt.Range.Font.Name = "Verdana"
tt.LockContentControl = True
tt.LockContents = False
Set tt = Nothing
'Selection.InsertRowsBelow 1
'Selection.Rows.Add BeforeRow:=Selection.Rows(1)
'Selection.InsertRowsBelow (1)
End If
End Sub