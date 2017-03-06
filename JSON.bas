Attribute VB_Name = "JSON"
Private Sub JsontoCSV_cnversion()
Dim FSO As New FileSystemObject
Dim JsonTS As TextStream
Dim JsonText As String
Dim Parsed As Dictionary
Dim CSVcont As String
' Read .json file
Set JsonTS = FSO.OpenTextFile("C:\development\DDS932_Subbu\VBA\JSON\VBA-JSON-master\VBA-JSON-master\specs\example.json", ForReading)
JsonText = JsonTS.ReadAll
JsonTS.Close
Set Parsed = JSONParser.ParseJson(JsonText)
CSVcont = """id"",""name"",""description"",""searchkey"""
CSVcont = CSVcont & Chr(13) & """" & Parsed("metadata")("id") & """,""" & Parsed("metadata")("name") & """,""" & Parsed("metadata")("description") & """,""" & """"
If Parsed("metadata")("children").Count > 1 Then
    For Each Value In Parsed("metadata")("children")
      CSVcont = CSVcont & Chr(13) & "" & Value("id") & """,""" & Value("name") & """,""" & Value("description") & """,""" & Parsed("metadata")("name") & "/" & Value("name") & """"
      If Value("children").Count >= 1 Then
         For Each value1 In Value("children")
            CSVcont = CSVcont & Chr(13) & "" & value1("id") & """,""" & value1("name") & """,""" & value1("description") & """,""" & Parsed("metadata")("name") & "/" & Value("name") & "/" & value1("name") & """"
              
                    If value1("children").Count >= 1 Then
                        For Each value2 In value1("children")
                             'CSVcont = CSVcont & Chr(13) & "" & value1("id") & """,""" & value1("name") & """,""" & value1("description") & """,""" & Parsed("metadata")("name") & "/" & Value("name") & "/" & value1("name") & """"
                             On Error Resume Next
                                If value2("children").Count >= 1 Then
                                   If Err.Number = 424 Then
                                   On Error GoTo 0
                                     CSVcont = CSVcont & Chr(13) & "" & value2("id") & """,""" & value2("name") & """,""" & value2("description") & """,""" & Parsed("metadata")("name") & "/" & Value("name") & "/" & value1("name") & "/" & value2("name") & """"
                                   Else
                                    'On Error GoTo 0
                                    For Each value3 In value2("children")
                                        CSVcont = CSVcont & Chr(13) & "" & value3("id") & """,""" & value3("name") & """,""" & value3("description") & """,""" & Parsed("metadata")("name") & "/" & Value("name") & "/" & value1("name") & "/" & value2("name") & """"
                                        If value3("children").Count >= 1 Then
                                            For Each value4 In value3("children")
                                                CSVcont = CSVcont & Chr(13) & "" & value4("id") & """,""" & value4("name") & """,""" & value4("description") & """,""" & Parsed("metadata")("name") & "/" & Value("name") & "/" & value1("name") & "/" & value2("name") & "/" & value2("name") & """"
                                            Next
                                        End If
                                
                                    Next
                                    End If
                                End If
                        
                            Next
                      End If
                Next
            End If
         Next
      End If
 
On Error GoTo 0
MsgBox CSVcont
Open ActiveDocument.Path & "\Test.csv" For Output As #1
Print #1, CSVcont
Close #1


End Sub
Public Sub search(csvfolder, csvfile, searchterm)

Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adCmdText = &H1

Dim strPathtoTextFile, objConnection, objRecordSet, objNetwork
Dim wshshell, Username

Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

'Set objNetwork = CreateObject("WScript.Network")
'Username = objNetwork.Username

strPathtoTextFile = csvfolder & "\"

objConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
          "Data Source=" & strPathtoTextFile & ";" & _
          "Extended Properties=""text;HDR=YES;FMT=Delimited"""

objRecordSet.Open "SELECT * FROM " & csvfile & " where [name] like '" & searchterm & "'", _
          objConnection, adOpenStatic, adLockOptimistic, adCmdText
Do Until objRecordSet.EOF
     UserForm1.ComboBox1.AddItem objRecordSet.Fields.Item("searchkey")
     objRecordSet.MoveNext
Loop
End Sub
