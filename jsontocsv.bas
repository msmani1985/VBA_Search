Attribute VB_Name = "Module1"
Private Sub JsontoCSV_cnversion()
Dim FSO As New FileSystemObject
Dim JsonTS As TextStream
Dim JsonText As String
Dim Parsed As Dictionary

' Read .json file
Set JsonTS = FSO.OpenTextFile("C:\development\DDS932_Subbu\VBA\JSON\VBA-JSON-master\VBA-JSON-master\specs\example.json", ForReading)
JsonText = JsonTS.ReadAll
JsonTS.Close

' Parse json to Dictionary
' "values" is parsed as Collection
' each item in "values" is parsed as Dictionary
Set Parsed = JsonConverter.ParseJson(JsonText)

' Prepare and write values to sheet
Dim Values As Variant
ReDim Values(Parsed("metadata").Count, 3)

Dim Value As Dictionary
Dim i As Long

i = 0
'For Each Value In Parsed("metadata")
  Values(i, 0) = Parsed("metadata")("id")
  Values(i, 1) = Parsed("metadata")("name")
  Values(i, 2) = Parsed("metadata")("description")
  For Each Value In Parsed("metadata")("children")
    MsgBox Value("id")
  Next
  i = i + 1
'Next Value

'Sheets("example").Range(Cells(1, 1), Cells(Parsed("values").Count, 3)) = Values
End Sub
