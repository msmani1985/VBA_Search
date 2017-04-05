VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private objXmlDom
Private objXmlNode
 
  ' GetObjectParam() ?????????????PJBlog
  Private Function GetObjectParam()
   On Error Resume Next
   Dim Temp
   GetObjectParam = "Microsoft.XMLDOM"
   Err = 0
   Dim TmpObj
   Set TmpObj = CreateObject(GetObjectParam)
   Temp = Err.Number
   If Temp = 1 Or Temp = -2147221005 Then
       GetObjectParam = "Msxml2.DOMDocument.5.0"
    End If
    Err.Clear
    Set TmpObj = Nothing
   Err = 0
  End Function
 
  Private Sub Class_Initialize()
      Set objXmlDom = CreateObject(GetObjectParam())
  End Sub
 
  Private Sub Class_Terminate()
      Set objXmlDom = Nothing
  End Sub
 
  Public Function encode(AnsiCode)
    encode = ""
    Set objXmlNode = objXmlDom.createElement("tmp")
    objXmlNode.DataType = "bin.base64"
    objXmlNode.nodeTypedvalue = AnsiCode
    encode = objXmlNode.Text
    Set objXmlNode = Nothing
  End Function
 
  Public Function decode(base64Code)
    On Error Resume Next
    decode = ""
    Set objXmlNode = objXmlDom.createElement("tmp")
    objXmlNode.DataType = "bin.base64"
    objXmlNode.Text = base64Code
    decode = objXmlNode.nodeTypedvalue
    Set objXmlNode = Nothing
    If Err Then
        Err.Clear
    End If
  End Function
 
  ' ?????????
  Public Function encodeText(ByVal str)
    On Error Resume Next
    Dim ado, r: r = ""
    If str <> "" Then
    Set ado = Server.CreateObject("ADODB.Stream")
        With ado
            .Charset = "gb2312"
            .Type = 2
            If .State = 0 Then .Open
            .WriteText str
            .Position = 0
            .Type = 1
            r = encode(.Read(-1))
            .Close
        End With
    Set ado = Nothing
    End If
    If Err Then Err.Clear: r = ""
    encodeText = r
  End Function
 
  ' ?????????
  Public Function decodeText(ByVal str)
    On Error Resume Next
    Dim ado, r: r = ""
    If str <> "" Then
    Set ado = CreateObject("ADODB.Stream")
        With ado
            .Charset = "gb2312"
            .Type = 1
            If .State = 0 Then .Open
            .Write (decode(str))
            .Position = 0
            .Type = 2
            r = .ReadText(-1)
            .Close
        End With
    Set ado = Nothing
    End If
    If Err Then Err.Clear: r = ""
    decodeText = r
  End Function

