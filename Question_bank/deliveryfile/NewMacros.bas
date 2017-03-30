Attribute VB_Name = "NewMacros"
Dim FSO As New FileSystemObject
Dim msword As New Word.Application
Dim flo As Folder
Dim flo1 As Folder
Dim fil As File
Dim varerrorno As Integer
Dim vartxt
Dim vartxt1
Dim strFile As String
Dim varCloseTag As Boolean
Dim txt1 As TextStream
Dim txt2 As TextStream
Dim txt3 As TextStream
Dim UpdateQuestionNo As Integer
Dim varcsms As Boolean
Dim varopencsms As Boolean
Dim varopencsms1 As Boolean
Public Sub CreateXML()
Dim vartxt
Dim para As Paragraph
Dim txt As TextStream
Dim varcopy As Boolean
Dim varfront As Boolean
Dim bodycont As Boolean
Dim varmeta As Boolean
Dim varcover As Boolean
Dim assetcount
Dim doccount
Dim imgcount
Dim manifestcount
Dim filcorupt As Boolean
'Set flo = fso.GetFolder(Dir1.Path)
varcopy = False
varfront = False
bodycont = False
varmeta = False
varcover = False
filcorupt = False




'varpathexe = Replace(varsplit1(0), Chr(34), "")
varpathinp = ActiveDocument.Path
varpathout = ActiveDocument.Path


Set flo = FSO.GetFolder(varpathinp)
        'If fso.FolderExists(varpathout) = False Then fso.CreateFolder (varpathout)
        
        'For Each flo1 In flo.SubFolders
        
        'If fso.FolderExists(varpathout & "\" & flo1.Name) = False Then fso.CreateFolder (varpathout & "\" & flo1.Name)
        
        If FSO.FileExists(varpathout & "\XML_Error_Log.txt") = True Then FSO.DeleteFile (varpathout & "\XML_Error_Log.txt"), True
        Open varpathout & "\XML_Error_Log.txt" For Output As #35
        varerrorno = 1

        'Print #35, "-------------------------------------------------------------------------------------------------------------"
        'Print #35, flo1.Name
        'Print #35, "-------------------------------------------------------------------------------------------------------------"
        
            'For Each fil In flo.Files
            
                'If fso.GetExtensionName(fil.Path) = "docm" Then
                    'Set aa1 = ActiveDocument
                    'Visible = True
                    Call Super_Find
                    Call Find_Italic1
                    Call Find_Bold
                    Call find_sub
                    Call Find_small
                    Call REFSPLCHAR1
                    Call ReplaceBullets
                    Call Replacenumberlist
                    Call Find_character
                    Call crossref
                    Call Find_Symbolcharacter
                    'If fso.FileExists(ActiveDocument.Path & "\Question_bank.zip") = True Then fso.DeleteFile (ActiveDocument.Path & "\Question_bank.zip")
                    'If fso.FileExists(ActiveDocument.Path & "\allQuestions.xml") = True Then fso.DeleteFile (ActiveDocument.Path & "\allQuestions.xml")
                    'If fso.FolderExists(ActiveDocument.Path & "\media") = True Then fso.DeleteFolder (ActiveDocument.Path & "\media")
                    Call XMLprocess(varpathout)
                    Call CreateManifest
                    Call zip
                    'Call Zipper
                    'strFile = varpathout & "\allQuestions.xml"
                    'Call ValidateFile(strFile)
                    'Call SaveImageAsPicture
                    'If fso.FolderExists(ActiveDocument.Path & "\media") = True Then fso.DeleteFolder (ActiveDocument.Path & "\media")
                    'If fso.FileExists(ActiveDocument.Path & "\allQuestions.xml") = True Then fso.DeleteFile (ActiveDocument.Path & "\allQuestions.xml")
                    'If fso.FileExists(ActiveDocument.Path & "\XML_Error_Log.txt") = True Then fso.DeleteFile (ActiveDocument.Path & "\XML_Error_Log.txt")
                    'ActiveDocument.Close (wdDoNotSaveChanges)
                    'If fso.FileExists(ActiveDocument.Path & "\Question_bank.docm") Then fso.DeleteFile (ActiveDocument.Path & "\Question_bank.docm")
                    ActiveDocument.Close (wdDoNotSaveChanges)
                    msword.Quit
                'End If
            'Next
       ' Next
        'Quit
        'MsgBox "completed"

End Sub

Public Sub REFSPLCHAR1()
    
    Dim txt As TextStream
    Selection.HomeKey wdStory
    Selection.Find.ClearFormatting
    Selection.Find.MatchWildcards = False
    Selection.Find.Text = "*"
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Text = "&#x002A;"
        
        If Selection.Find.Execute = True Then
        Selection.Find.Execute Replace:=wdReplaceAll
        End If
        Set txt = FSO.OpenTextFile(ActiveDocument.Path & "\ascii.txt")
        'MsgBox (App.Path)
        Do While Not txt.AtEndOfStream
        vartxt = txt.ReadLine
        vartxt1 = Split(vartxt, vbTab)
        Selection.HomeKey wdStory
        Selection.Find.ClearFormatting
        Selection.Find.MatchWildcards = False
        Selection.Find.MatchCase = True
        Selection.Find.Text = "^u" & Replace(vartxt1(0), "D+", "")
        Selection.Find.Replacement.ClearFormatting
            If Selection.Find.Execute = True Then
            Selection.Find.Replacement.Text = "&#x" & Replace(vartxt1(2), "U+", "") & ";"
            'If Selection.Find.Execute = True Then
            'MsgBox vartxt1(0)
            Selection.Find.Execute Replace:=wdReplaceAll
            End If
        Loop
        
    txt.Close
End Sub

Public Sub Find_Italic1()
    Selection.HomeKey wdStory
    With Selection.Find
        .ClearFormatting
        .Text = "^p"
        .Font.Italic = 1
        .MatchWildcards = False
        .Replacement.Text = "^p"
        .Replacement.Font.Italic = 0
        .Execute Replace:=wdReplaceAll
    End With
    Selection.HomeKey wdStory
    With Selection.Find
        .ClearFormatting
        .Text = ""
        .Font.Italic = 1
        .MatchWildcards = False
        .Replacement.Text = "<i>" & "^&" & "</i>"
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Public Sub Super_Find()
    Selection.HomeKey wdStory
    With Selection.Find
        .ClearFormatting
        .Text = "^p"
        .Font.Superscript = 1
        .MatchWildcards = False
        .Replacement.Text = "^p"
        .Replacement.Font.Superscript = 0
        .Execute Replace:=wdReplaceAll
    End With
    
    Selection.HomeKey unit:=wdStory
    Selection.Find.ClearFormatting
    Do
        With Selection.Find
            .ClearFormatting
            .Text = ""
            .Font.Superscript = 1
            .MatchWildcards = False
        End With
        Selection.Find.Execute
        If Selection.Find.Found = True Then
            If Selection.Style = "Hyperlink" Or Selection.Fields.Count > 0 Then
                'MsgBox "OK"
                suptxt = Selection.Characters.Parent
                If Trim(suptxt) <> "" Then
                     Selection.InsertBefore "<sup>"
                     Selection.InsertAfter "</sup>"
                     Selection.MoveRight wdCharacter, 1
                End If
            Else
                suptxt = Selection.Characters.Parent
                If Trim(suptxt) <> "" Then
                    Selection.Font.Superscript = 0
                    Selection.Delete
                    
                    Selection.TypeText ("<sup>" & suptxt & "</sup>")
                    Selection.MoveRight wdCharacter, 1
                End If
            End If
        End If
    Loop Until Selection.Find.Found = False
End Sub

Public Sub Find_Italic()
    Selection.HomeKey wdStory
    With Selection.Find
        .ClearFormatting
        .Text = "^p"
        .Font.Italic = 1
        .MatchWildcards = False
        .Replacement.Text = "^p"
        .Replacement.Font.Italic = 0
        .Execute Replace:=wdReplaceAll
    End With
    Selection.HomeKey wdStory
     With Selection.Find
       .ClearFormatting
        .Text = ""
        .Font.Italic = 1
        .MatchWildcards = False
        .Replacement.Text = "<i>" & "^&" & "</i>"
    End With
        Do
              With Selection.Find
            .ClearFormatting
            .Text = ""
            .Font.Italic = 1
            .MatchWildcards = False
            Do While .Execute = True
                If InStr("<i>" & Selection.Text, ".<sup>") > 0 Then
                
                Selection.Text = "<i>" & Replace(Replace(Selection.Text, "<sup>", "</i><sup><i>"), "</sup>", "</i></sup>")
                ElseIf Selection.Text = " " Then
               ' MsgBox "text"
                Else
                Selection.Text = "<i>" & Selection.Text & "</i>"
            '.Execute Replace:=wdReplaceAll
            End If
             Selection.Font.Italic = False
            Selection.MoveRight wdCharacter, 1
            Loop
            End With
        Loop Until Selection.Find.Found = False
End Sub

Public Sub Find_character()

On Error Resume Next
    Selection.HomeKey wdStory
    With Selection.Find
        .ClearFormatting
        .Text = "^p"
       ' .Font.Bold = 1
       
       .Style = "Glossary_Term"
        
        .MatchWildcards = False
        .Replacement.Text = "^p"
        .Replacement.Style = "Default Paragraph Font"
        '.Replacement.Font.Bold = 0
        .Execute Replace:=wdReplaceAll
    End With
    Selection.HomeKey wdStory
    With Selection.Find
        .ClearFormatting
        .Text = ""
        '.Font.Bold = 1
        .Style = "Glossary_Term"
        .MatchWildcards = False
        .Replacement.Text = "<Glossary_Term>" & "^&" & "</Glossary_Term>"
        .Execute Replace:=wdReplaceAll
    End With

End Sub

Public Sub Find_small()
     Selection.HomeKey wdStory
    With Selection.Find
        .ClearFormatting
        .Text = "^p"
        .Font.SmallCaps = 1
        .MatchWildcards = False
        .Replacement.Text = "^p"
        .Replacement.Font.SmallCaps = 0
        .Execute Replace:=wdReplaceAll
    End With
    Selection.HomeKey unit:=wdStory
    Selection.Find.ClearFormatting
    Do
    With Selection.Find
        .ClearFormatting
        .Text = ""
        .Font.SmallCaps = 1
        .MatchWildcards = False
        Do While .Execute
        '.Replacement.Text = "<sc>" & "^&" & "</sc>"
        Selection.Text = "<small>" & UCase(Selection.Text) & "</small>"
        Selection.Font.SmallCaps = False
        Selection.MoveRight wdCharacter, 1
        '.Execute Replace:=wdReplaceAll
        Loop
    End With
    Loop Until Selection.Find.Found = False
End Sub

Public Sub find_sub()
    Selection.HomeKey wdStory
    With Selection.Find
        .ClearFormatting
        .Text = "^p"
        .Font.Subscript = 1
        .MatchWildcards = False
        .Replacement.Text = "^p"
        .Replacement.Font.Subscript = 0
        .Execute Replace:=wdReplaceAll
    End With
    Selection.HomeKey unit:=wdStory
    Selection.Find.ClearFormatting
    Do
        With Selection.Find
            .ClearFormatting
            .Text = ""
            .Font.Subscript = 1
            .MatchWildcards = False
        End With
        Selection.Find.Execute
        If Selection.Find.Found = True Then
            If Selection.Style = "Hyperlink" Or Selection.Fields.Count > 0 Then
                'MsgBox "OK"
                suptxt1 = Selection.Characters.Parent
                If Trim(suptxt1) <> "" Then
                     Selection.InsertBefore "<sub>"
                     Selection.InsertAfter "</sub>"
                     Selection.MoveRight wdCharacter, 1
                End If
            Else
                suptxt1 = Selection.Characters.Parent
                If Trim(suptxt1) <> "" Then
                    Selection.Font.Subscript = 0
                    Selection.Delete
                    
                    Selection.TypeText ("<sub>" & suptxt1 & "</sub>")
                    Selection.MoveRight wdCharacter, 1
                End If
            End If
        End If
    Loop Until Selection.Find.Found = False
End Sub

 Public Sub Find_Bold()
    Selection.HomeKey wdStory
    With Selection.Find
        .ClearFormatting
        .Text = "^p"
        .Font.Bold = 1
        .MatchWildcards = False
        .Replacement.Text = "^p"
        .Replacement.Font.Bold = 0
        .Execute Replace:=wdReplaceAll
    End With
    Selection.HomeKey wdStory
    With Selection.Find
        .ClearFormatting
        .Text = ""
        .Font.Bold = 1
        .MatchWildcards = False
        .Replacement.Text = "<b>" & "^&" & "</b>"
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Sub FindBullet()
    Dim rngTarget As Word.Range
    Dim oPara As Word.Paragraph
    Selection.HomeKey wdStory
    Set rngTarget = Selection.Range
    With rngTarget
        Call .Collapse(wdCollapseEnd)
        .End = ActiveDocument.Range.End

        For Each oPara In .Paragraphs
            If oPara.Range.ListFormat.ListType = _
               WdListType.wdListBullet Then
                oPara.Range.Select
                Selection.InsertBefore "<ulli>"
                Selection.MoveRight wdCharacter, 1
                Selection.InsertAfter "</ulli>"
              ' Selection.TypeText ("<ulli>" & oPara.Range.Text)
                 Selection.MoveRight wdCharacter, 1
                  '  suptxt1 = oPara.Range.Text
'                    If Trim(suptxt1) <> "" Then
'                     Selection.InsertBefore "<sub>"
'                     Selection.InsertAfter "</sub>"
'                     Selection.MoveRight wdCharacter, 1
'                     End If
               ' End If
                'Exit For
            End If
        Next
    End With
End Sub

Sub FindNumber()
    Dim rngTarget As Word.Range
    Dim oPara As Word.Paragraph
    Selection.HomeKey wdStory
    Set rngTarget = Selection.Range
    With rngTarget
        Call .Collapse(wdCollapseEnd)
        .End = ActiveDocument.Range.End

        For Each oPara In .Paragraphs
            If oPara.Range.ListFormat.ListType = _
               WdListType.wdListSimpleNumbering Then
                oPara.Range.Select
                 Selection.TypeText ("<olli>" & oPara.Range.Text)
                 Selection.MoveRight wdCharacter, 1
                'Exit For
            End If
        Next
    End With
End Sub

Sub bullformat()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    'Selection.Find.Replacement.Style = ActiveDocument.Styles(wdStyleListBullet)
        With Selection.Find
        .Text = "·" & vbTab
        .Replacement.Text = "<lib>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        End With
    Selection.Find.Execute Replace:=wdReplaceAl
End Sub

Sub ConvertBulletChar()
    Const BulletStyle As String = "List Bullet"
    Dim para As Paragraph
    Dim BulletChar As String
    Dim s As String
   
    BulletChar = Chr$(149)
   
    With ActiveDocument
        For Each para In .Paragraphs
            With para
                s = Mid$(.Range.Text, 1, 1)
                If InStr(1, BulletChar, s) Then
                    .Range.Characters.Item(1).Delete    ' Removes 1st char of Paragraph chr(149)
                    .Range.ListFormat.ApplyListTemplate ListTemplate:=ListGalleries( _
        wdBulletGallery).ListTemplates(1), ContinuePreviousList:=False, ApplyTo:= _
        wdListApplyToWholeList, DefaultListBehavior:=wdWord9ListBehavior  ' Apply bullet list
                                     
                End If
            End With
        Next
    End With
End Sub

Sub Find_Symbolcharacter()
On Error Resume Next
    Selection.HomeKey wdStory
    With Selection.Find
        .ClearFormatting
        .Text = "^p"
       ' .Font.Bold = 1
       
       .Style = "Symbol1"
        
        .MatchWildcards = False
        .Replacement.Text = "^p"
        .Replacement.Style = "Default Paragraph Font"
        '.Replacement.Font.Bold = 0
        .Execute Replace:=wdReplaceAll
    End With
    Selection.HomeKey wdStory
    With Selection.Find
        .ClearFormatting
        .Text = ""
        '.Font.Bold = 1
        .Style = "Symbol1"
        .MatchWildcards = False
        .Replacement.Text = "<font face=""Symbol"" size=""4"">" & "^&" & "</font>"
        .Execute Replace:=wdReplaceAll
    End With

End Sub
 
Sub ReplaceBullets()
'    Dim oPara As Paragraph
'    Dim r As Range
'    For Each oPara In ActiveDocument.Paragraphs()
'        Set r = oPara.Range
'        If r.ListFormat.ListType = wdListBullet Then
'            r.ListFormat.RemoveNumbers _
'            NumberType:=wdNumberParagraph
'            r.InsertBefore Text:="<lib>"
'        End If
'        Set r = Nothing
'    Next
'Sub ReplaceBullets()
   Dim oPara As Paragraph
    Dim r As Range
    For Each oPara In ActiveDocument.Paragraphs()
        Set r = oPara.Range
        r.Select
'        If InStr(r.Text, "History of possible or probable PSP and histological evidence of typical") Then
'            MsgBox "ok"
'        End If
        Selection.HomeKey
        If r.ListFormat.ListType = wdListBullet Then
            r.ListFormat.RemoveNumbers _
            NumberType:=wdNumberParagraph
            If oPara.Range.Previous(wdParagraph, 1).ListFormat.ListType = wdListBullet Then
                
                If oPara.Range.Next(wdParagraph, 1).ListFormat.ListType = wdListBullet Then
                    r.InsertBefore Text:="<li>"
                    
                    'r.InsertAfter Text:="</lib>"
                   
                    Else
                    'r.InsertBefore Text:="</ul><li>"
                    r.InsertBefore Text:="<li>"
                    Selection.EndOf unit:=wdParagraph, Extend:=wdExtend
                    Selection.EndKey
                    Selection.Text = "</li></ul>"
                    Selection.EndKey unit:=wdLine
                End If
            Else
            If InStr(oPara.Range.Previous(wdParagraph, 1).Text, "<li>") Then
              '  r.InsertBefore Text:="<lib>"
                    If oPara.Range.Next(wdParagraph, 1).ListFormat.ListType = wdListBullet Then
                    r.InsertBefore Text:="<li>"
                    Selection.EndOf unit:=wdParagraph, Extend:=wdExtend
                    Selection.EndKey
                    Selection.Text = "</li>"
                    Selection.EndKey unit:=wdLine
                    Else
'                     r.InsertBefore Text:="<li>"
'                    Selection.EndKey Unit:=wdLine
'
'                    Selection.Text = "</li></ul"
                    r.Select
                    r.InsertBefore Text:="<li>"
                    Selection.EndOf unit:=wdParagraph, Extend:=wdExtend
                    Selection.EndKey
                    Selection.Text = "</li></ul>"
                    Selection.EndKey unit:=wdLine
                    
                  '  r.InsertBefore Text:="</ul><li>"
                End If
                Else
                
                r.InsertBefore Text:="<ul><li>"
                Selection.EndOf unit:=wdParagraph, Extend:=wdExtend
                    Selection.EndKey
                    If oPara.Range.Next(wdParagraph, 1).ListFormat.ListType = wdListBullet Then
                    Selection.Text = "</li>"
                    Else
                     Selection.Text = "</li></ul>"
                     End If
                Selection.EndKey unit:=wdLine
                
                'oPara.Range.Text = "</li>"
                'Selection.Text = "</li>"
                
               ' Selection.EndKey Unit:=wdLine
            End If
            'r.InsertBefore Text:="<lib>"
            
            End If
        End If
        Set r = Nothing
    Next
'End Sub
End Sub

Sub Replacenumberlist()
'    Dim oPara As Paragraph
'    Dim r As Range
'    For Each oPara In ActiveDocument.Paragraphs()
'        Set r = oPara.Range
'        If r.ListFormat.ListType = wdListSimpleNumbering Then
'            r.ListFormat.RemoveNumbers _
'            NumberType:=wdNumberParagraph
'            r.InsertBefore Text:="<liol>"
'            'r.InsertAfter Text:="</liol>"
'        End If
'        Set r = Nothing
'    Next
Dim r As Range
    For Each oPara In ActiveDocument.Paragraphs()
        Set r = oPara.Range
        r.Select
'        If InStr(r.Text, "History of possible or probable PSP and histological evidence of typical") Then
'            MsgBox "ok"
'        End If
        Selection.HomeKey
        If r.ListFormat.ListType = wdListSimpleNumbering Then
            r.ListFormat.RemoveNumbers _
            NumberType:=wdNumberParagraph
            If oPara.Range.Previous(wdParagraph, 1).ListFormat.ListType = wdListSimpleNumbering Then
                
                If oPara.Range.Next(wdParagraph, 1).ListFormat.ListType = wdListSimpleNumbering Then
                    r.InsertBefore Text:="<li>"
                    
                    'r.InsertAfter Text:="</lib>"
                   
                    Else
                    'r.InsertBefore Text:="</ul><li>"
                    r.InsertBefore Text:="<li>"
                    Selection.EndOf unit:=wdParagraph, Extend:=wdExtend
                    Selection.EndKey
                    Selection.Text = "</li></ol>"
                    Selection.EndKey unit:=wdLine
                End If
            Else
            If InStr(oPara.Range.Previous(wdParagraph, 1).Text, "<li>") Then
              '  r.InsertBefore Text:="<lib>"
                    If oPara.Range.Next(wdParagraph, 1).ListFormat.ListType = wdListSimpleNumbering Then
                    r.InsertBefore Text:="<li>"
                    Selection.EndOf unit:=wdParagraph, Extend:=wdExtend
                    Selection.EndKey
                    Selection.Text = "</li>"
                    Selection.EndKey unit:=wdLine
                    Else
'                     r.InsertBefore Text:="<li>"
'                    Selection.EndKey Unit:=wdLine
'
'                    Selection.Text = "</li></ul"
                    r.Select
                    r.InsertBefore Text:="<li>"
                    Selection.EndOf unit:=wdParagraph, Extend:=wdExtend
                    Selection.EndKey
                    Selection.Text = "</li></ol>"
                    Selection.EndKey unit:=wdLine
                    
                  '  r.InsertBefore Text:="</ul><li>"
                End If
                Else
                
                r.InsertBefore Text:="<ol><li>"
                Selection.EndOf unit:=wdParagraph, Extend:=wdExtend
                    Selection.EndKey
                    If oPara.Range.Next(wdParagraph, 1).ListFormat.ListType = wdListSimpleNumbering Then
                    Selection.Text = "</li>"
                    Else
                     Selection.Text = "</li></ol>"
                     startno = startno + 1
                     End If
                Selection.EndKey unit:=wdLine
                
                'oPara.Range.Text = "</li>"
                'Selection.Text = "</li>"
                
               ' Selection.EndKey Unit:=wdLine
            End If
            'r.InsertBefore Text:="<lib>"
            
            End If
        End If
        Set r = Nothing
    Next
End Sub

Sub crossref()
 Set colHyperlinks = ActiveDocument.Hyperlinks
    For Each objHyperlink In colHyperlinks
    'Debug.Print objHyperlink.SubAddress
    varcross = Trim(objHyperlink.SubAddress)
    varfind1 = objHyperlink.Range.Text
    If varcross <> "" Then
    Selection.Find.ClearFormatting
                 With Selection.Find
                 '.Style = "Pubmed_link"
                     .Text = varfind1
                     .Replacement.Text = ""
                     .Forward = True
                     .Wrap = wdFindContinue
                     .Format = False
                     .MatchCase = True
                     .MatchWholeWord = False
                     .MatchWildcards = False
                     .MatchSoundsLike = False
                     .MatchAllWordForms = False
                End With
                Selection.Find.Execute
                '<span id="typeText" class="greenText" data-toggle="modal" data-target="#myModal" data-info=""></span>
                If InStr(varcross, "R") Then
                Selection.TypeText ("<span id=""typeText"" class=""greenText"" data-toggle=""modal"" data-target=""#myModal"" data-info=" & Chr(34) & varcross & Chr(34) & ">" & varfind1 & "</span>")
                End If
                'ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:= _
                "http://www.ncbi.nlm.nih.gov/pubmed/" & varpubid & "?dopt=Abstract", SubAddress:="", ScreenTip:="", _
                TextToDisplay:=varpubid
    'Debug.Print varcross
    'Selection.TypeText ("<link>" & objHyperlink.SubAddress & "</link>")
    '                 Selection.MoveRight wdCharacter, 1
    Else
    Selection.Find.ClearFormatting
                 With Selection.Find
                 '.Style = "Pubmed_link"
                     .Text = varfind1
                     .Replacement.Text = ""
                     .Forward = True
                     .Wrap = wdFindContinue
                     .Format = False
                     .MatchCase = True
                     .MatchWholeWord = False
                     .MatchWildcards = False
                     .MatchSoundsLike = False
                     .MatchAllWordForms = False
                End With
                Selection.Find.Execute
            Selection.TypeText ("<a href=" & Chr(34) & objHyperlink.Address & Chr(34) & " class=""greenText"" target=""_blank"" >" & _
            objHyperlink.TextToDisplay & "</a>")
    'Debug.Print objHyperlink.Address
    End If
    Next
End Sub
Private Function metaxml(t As Table, varkno) As String
varrow = 1

metaxml = "<question_meta_tag qmtmode=""C"">" & Chr(13)


                                        '    Print #99, "            <meta_tag ucx=""C"" metaTagId="""">"
                                        '    Print #99, "                <meta_tag_type>" & varcontent1 & "</meta_tag_type>"
                                        'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_VALUE" Then
                                        '    Print #99, "                <meta_tag_value>" & varcontent1 & "</meta_tag_value>"
                                        '    Print #99, "            </meta_tag>"
Do While varrow <= t.Rows.Count
If Replace(t.Cell(varrow, 1).Range.Text, "", "") = "Free Text" Then

metaxml = metaxml & "            <meta_tag ucx=""C"" metaTagId="""">" & Chr(13)
metaxml = metaxml & "                  <meta_tag_type>" & Replace(t.Cell(varrow, 1).Range.Text, "", "") & "</meta_tag_type>" & Chr(13)
metaxml = metaxml & "                  <meta_tag_value>" & Replace(t.Cell(varrow, 2).Range.Text, "", "") & "</meta_tag_value>" & Chr(13)
metaxml = metaxml & "              </meta_tag>" & Chr(13)
ElseIf Replace(t.Cell(varrow, 1).Range.Text, "", "") = "Look Up" Then
If t.Cell(varrow + 1, 1).Range.InlineShapes.Count > 0 Then

varuid = "___" & Split(t.Cell(varrow + 1, 1).Range.InlineShapes(1).OLEFormat.Object.Name, "___")(1)
varuid = "{" & Replace(Replace(varuid, "__", "-"), "-_", "") & "}"
For Each XMLNode In ActiveDocument.CustomXMLParts.SelectByID(varuid).SelectNodes("//metanode[@selected=""True""]")
metaxml = metaxml & " <meta_tag ucx=""C"" metaTagId=""" & XMLNode.ChildNodes(3).Text & """  metadataId=""" & XMLNode.ParentNode.Attributes(1).Text & """>" & vbCrLf
               metaxml = metaxml & " <meta_tag_type>" & XMLNode.ChildNodes(4).Text & "</meta_tag_type>" & vbCrLf
               metaxml = metaxml & "<meta_tag_value>" & XMLNode.ChildNodes(2).Text & "</meta_tag_value>" & vbCrLf
                metaxml = metaxml & "<meta_tag_path>" & XMLNode.ChildNodes(1).Text & "</meta_tag_path>" & vbCrLf
metaxml = metaxml & "</meta_tag>" & Chr(13)
    
Next
End If
ElseIf Replace(t.Cell(varrow, 1).Range.Text, "", "") = "Hierarchy" Then
If t.Cell(varrow + 1, 1).Range.InlineShapes.Count > 0 Then
'MsgBox t.Cell(varrow + 1, 1).Range.InlineShapes(1).OLEFormat.Object.Name
varuid = "___" & Split(t.Cell(varrow + 1, 1).Range.InlineShapes(1).OLEFormat.Object.Name, "___")(1)
varuid = "{" & Replace(Replace(varuid, "__", "-"), "-_", "") & "}"
For Each XMLNode In ActiveDocument.CustomXMLParts.SelectByID(varuid).SelectNodes("//metanode[@selected=""True""]")
metaxml = metaxml & " <meta_tag ucx=""C"" metaTagId=""" & XMLNode.ChildNodes(3).Text & """  metadataId=""" & XMLNode.ParentNode.Attributes(1).Text & """>" & vbCrLf
               metaxml = metaxml & " <meta_tag_type>" & XMLNode.ChildNodes(4).Text & "</meta_tag_type>" & vbCrLf
               metaxml = metaxml & "<meta_tag_value>" & XMLNode.ChildNodes(2).Text & "</meta_tag_value>" & vbCrLf
                metaxml = metaxml & "<meta_tag_path>" & XMLNode.ChildNodes(1).Text & "</meta_tag_path>" & vbCrLf
metaxml = metaxml & "</meta_tag>" & Chr(13)
    
Next
End If
End If
varrow = varrow + 1
Loop
metaxml = metaxml & "</question_meta_tag>" & Chr(13)
'MsgBox "ok"
End Function
Private Function remxml(t As Table, varkno, varspecialtype) As String
vartype1 = ""
varrow = 1
If varspecialtype = True Then
vartag = "<cs_sub_"
Else
vartag = "<"
End If
remxml = vartag & "question_remediation_link qrlmode=""C"">"
Do While varrow <= t.Rows.Count
 'For varcol = 1 To t.Columns.Count
    If Replace(t.Cell(varrow, 1).Range.Text, "", "") = "Web Link" Then
        vartype1 = Replace(t.Cell(varrow, 1).Range.Text, "", "")
        'MsgBox t.Cell(varrow, 2)
    ElseIf Replace(t.Cell(varrow, 1).Range.Text, "", "") = "EBook" Then
       ' MsgBox t.Cell(varrow, 2)
        vartype1 = Replace(t.Cell(varrow, 1).Range.Text, "", "")
    ElseIf Replace(t.Cell(varrow, 1).Range.Text, "", "") = "Text" Then
        'MsgBox t.Cell(varrow, 2)
        vartype1 = Replace(t.Cell(varrow, 1).Range.Text, "", "")
    End If
    If vartype1 = "Web Link" Then
        'vartype1 = t.Cell(varrow, 1)
        'MsgBox t.Cell(varrow, 2)
        remxml = remxml & vartag & "remediation_type ucx=""C"" redLinkId=""1"" remediation_link_type=""Web Link"">" & Chr(13)
        remxml = remxml & vartag & "remediation_type_text>" & Replace(t.Cell(varrow, 2).Range.Text, "", "") & Replace(vartag, "<", "</") & "remediation_type_text>" & Chr(13)
        remxml = remxml & vartag & "remediation_type_link>" & Replace(t.Cell(varrow + 1, 2).Range.Text, "", "") & Replace(vartag, "<", "</") & "remediation_type_link>" & Chr(13)
        remxml = remxml & vartag & "remediation_type_tooltip>" & Replace(t.Cell(varrow + 2, 1).Range.Text, "", "") & Replace(vartag, "<", "</") & "remediation_type_tooltip>" & Chr(13)
        remxml = remxml & Replace(vartag, "<", "</") & "remediation_type>" & Chr(13)
        varrow = varrow + 2
        vartype1 = ""
    ElseIf vartype1 = "EBook" Then
        remxml = remxml & vartag & "remediation_type ucx=""C"" redLinkId=""1"" remediation_link_type=""EBook"">" & Chr(13)
        remxml = remxml & vartag & "remediation_type_text>" & Replace(t.Cell(varrow, 2).Range.Text, "", "") & Replace(vartag, "<", "</") & "remediation_type_text>" & Chr(13)
        remxml = remxml & Replace(vartag, "<", "</") & "remediation_type>" & Chr(13)
        vartype1 = ""
        
    ElseIf vartype1 = "Text" Then
        remxml = remxml & vartag & "remediation_type ucx=""C"" redLinkId=""1"" remediation_link_type=""Text"">" & Chr(13)
        remxml = remxml & vartag & "remediation_type_text>" & Replace(t.Cell(varrow, 2).Range.Text, "", "") & Replace(vartag, "<", "</") & "remediation_type_text>"
     '   remxml = remxml & "<remediation_type_link>" & Replace(t.Cell(varrow + 1, 2).Range.Text, "", "") & "</remediation_type_link>"
     '   remxml = remxml & "<remediation_type_tooltip>" & Replace(t.Cell(varrow + 2, 1).Range.Text, "", "") & "</remediation_type_tooltip>"
        remxml = remxml & Replace(vartag, "<", "</") & "remediation_type>" & Chr(13)
      '  varrow = varrow + 2
        vartype1 = ""
        'vartype1 = t.Cell(varrow, 1)
    End If
      varrow = varrow + 1
 'Next
Loop
remxml = remxml & Replace(vartag, "<", "</") & "question_remediation_link>" & Chr(13)
If varspecialtype = True Then remxml = remxml & "</cs_sub_question>"
End Function
Sub XMLprocess(varpathout)
Dim tblOne As Table
Dim para As Paragraph
QuestionNo = 1
varcsms = False
varopencsms = False
varopencsms1 = False
Open varpathout & "\allQuestions.xml" For Output As #99
Print #99, "<?xml version=""1.0"" encoding=""UTF-8""?>"
Print #99, "<wk_question_root mode=""C"">"


For Each para In ActiveDocument.Paragraphs
    
    If para.Range.Style = "Question_Type" Then
        If InStr(para.Range.Text, "Multiple Choice") Then
            questiontype = "MC"
            questiontypename = "Multiple Choice"
            tagtype = "radioButton"
        ElseIf InStr(para.Range.Text, "Choice Multiple") Then
            questiontype = "CM"
            questiontypename = "Choice Multiple"
            tagtype = "checkBox"
        ElseIf InStr(para.Range.Text, "True or Fasle") Then
            questiontype = "TF"
            questiontypename = "True / False"
            tagtype = "radioButton"
        ElseIf InStr(para.Range.Text, "Video Questions") Then
            questiontype = "VQ"
            questiontypename = "Video Questions"
            tagtype = "radioButton"
        ElseIf InStr(para.Range.Text, "Image Integration") Then
            questiontype = "II"
            questiontypename = "Image Integration"
            tagtype = "radioButton"
        ElseIf InStr(para.Range.Text, "Graphic Option") Then
            questiontype = "GO"
            questiontypename = "Graphic Option"
            tagtype = "radioButton"
        ElseIf InStr(para.Range.Text, "Clinical Symptoms") Then
            questiontype = "CS"
            questiontypename = "Clinical Symptoms"
            tagtype = "radioButton"
        ElseIf InStr(para.Range.Text, "Medical Case") Then
            questiontype = "MEDC"
            questiontypename = "Medical Case"
            tagtype = "radioButton"
        End If
    End If
    
    If para.Range.Style = "Question_Type" Then
        
        If para.Range.Tables.Count >= 1 Then
            
            If para.Range.Previous(wdParagraph, 1) Is Nothing Then
            Else
                
                
'                If Para.Range.Style = "Question_Text" Then
'                Open App.Path & "\output\" & flo1.Name & "\" & varfilename & varTable & ".xml" For Output As #99
                varcorrectcm = ""
                K = 1
                'varCloseTag = False
                'Print #99, "    <wk_question identificationId=""0"" qtype=""" & questiontype & """ qmode=""C"">"
                'Print #99, "        <question_type ucx=""C"" >" & questiontypename & "</question_type>"
                
                
                vartabl = para.Range.Tables.Count
                    For varrow = 1 To para.Range.Tables(vartabl).Rows.Count
                    ' MsgBox Para.Range.Tables(vartabl).Rows(varrow).Cells.Count
                         For varcol = 1 To para.Range.Tables(vartabl).Rows(varrow).Cells.Count
                            para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Select
                        
                                varcontent1 = Replace(Replace(Replace(Replace(Replace(Replace(para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Text, Chr(13), ""), vbTab, ""), Chr(10), ""), Chr(11), ""), Chr(7), ""), "", "")
                                  On Error Resume Next
                                   Debug.Print para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style
                                  If Err.Number = 91 Then
                                 On Error GoTo 0
                                    Selection.MoveLeft unit:=1
                                    If (Selection.Tables(1).id = "Remidation") Then
                                        varxml = remxml(Selection.Tables(1), varkno, varspecialtype)
                                          Print #99, varxml
                                    Else
                                    If InStr(Selection.Tables(1).Range.Text, "Remediation Type") > 0 Then
                                    varxml = remxml(Selection.Tables(1), varkno, varspecialtype)
                                    Print #99, varxml
                                    Else
                                    'MsgBox "ok"
                                    If (Selection.Tables(1).id = "Metadata") Then
                                    varxml = metaxml(Selection.Tables(1), varkno)
                                     Print #99, varxml
                                       Else
                                     If (para.Range.Tables(vartabl).Rows(varrow - 1).Cells(varcol).Range.Text = "Metadata Attributes") Then
                                     varxml = metaxml(Selection.Tables(1), varkno)
                                     Print #99, varxml
                                     End If
                                    End If
                                    End If
                                    End If
                                    
                               
                                'MsgBox Err.Number
                                Else
                                 On Error GoTo 0
                                If para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Delete_box" Then
                                    varCheckboxvalue = ActiveDocument.FormFields("Delete_" & QuestionNo).CheckBox.Value
                                    If varCheckboxvalue = True Then GoTo E222:
                                End If
                                If questiontype = "MC" And varspecialtype = False Then
                                 If varcsms = True Then
                                    Print #99, "</question_cs_sub_questions>"
                                    Print #99, "</wk_question>"
                                    varcsms = False
                               End If
                                        varspecialtype = False
                                        If para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Type" Then
                                            varIdentificationId = Replace(Replace(para.Range.Tables(vartabl).Rows(varrow + 1).Cells(varcol + 1).Range.Text, "Identification Id: ", ""), "", "")
                                            Print #99, "    <wk_question identificationId=""" & varIdentificationId & """ qtype=""" & questiontype & """ qmode=""C"">"
                                            Print #99, "        <question_type ucx=""C"" >" & questiontypename & "</question_type>"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Title" Then
                                            Print #99, "        <question_title ucx=""C"">" & varcontent1 & "</question_title>"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Text" Then
                                            Print #99, "        <question_text ucx=""C"">" & varcontent1 & "</question_text>"
                                            Print #99, "        <question_multiple_choices qmcmode=""C"" tagtype=""" & tagtype & """ >"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Answer_Text" Then
                                            Print #99, "            <question_choice ucx=""C"" refId="""">"
                                            Print #99, "                <question_answer_text>" & varcontent1 & "</question_answer_text>"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Rationale_Text" Then
                                            Print #99, "                <question_rationale>" & varcontent1 & "</question_rationale>"
                                            Print #99, "            </question_choice>"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Correct_Answer" Then
                                            Print #99, "        </question_multiple_choices>"
                                            Print #99, "        <correct_answer>" & varcontent1 & "</correct_answer>"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Score" Then
                                            Print #99, "        <question_score ucx=""C"">" & varcontent1 & "</question_score>"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Difficulty" Then
                                            Print #99, "        <question_difficulty ucx=""C"" >" & varcontent1 & "</question_difficulty>"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Correct_Answer_Rationale" Then
                                            Print #99, "        <question_correct_rationale ucx=""C"">" & varcontent1 & "</question_correct_rationale>"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Incorrect_Answer_Rationale" Then
                                            Print #99, "        <question_incorrect_rationale ucx=""C"">" & varcontent1 & "</question_incorrect_rationale>"
                                        '    Print #99, "        <question_remediation_link qrlmode=""C"">"
                                        'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Text" Then
                                        '    Print #99, "            <remediation_type ucx=""C"" redLinkId=""1"" remediation_link_type=""Web Link"">"
                                        '    Print #99, "                <remediation_type_text>" & varcontent1 & "</remediation_type_text>"
                                        'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Link" Then
                                                'If InStr(varcontent1, "http") Then
                                                'Else
                                                'varcontent1 = "http://" & varcontent1
                                                'End If
                                         '   Print #99, "                <remediation_type_link>" & varcontent1 & "</remediation_type_link>"
                                        'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Tooltip" Then
                                        '    Print #99, "                <remediation_type_tooltip>" & varcontent1 & "</remediation_type_tooltip>"
                                        '    Print #99, "            </remediation_type>"
                                         '   Print #99, "             <remediation_type ucx=""C"" redLinkId=""2"" remediation_link_type=""Ebook""></remediation_type>"
                                        'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Text_Remediation_Text" Then
                                         '   Print #99, "             <remediation_type ucx=""C"" redLinkId=""3"" remediation_link_type=""Text"">"
                                          '  Print #99, "                <remediation_type_text>" & varcontent1 & "</remediation_type_text>"
                                           ' Print #99, "            </remediation_type>"
                                        'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Meta_Data_Attributes_head" Then
                                        '    Print #99, "        </question_remediation_link>"
                                        '    Print #99, "        <question_meta_tag qmtmode=""C"">"
                                        'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_TAG" Then
                                        '    Print #99, "            <meta_tag ucx=""C"" metaTagId="""">"
                                        '    Print #99, "                <meta_tag_type>" & varcontent1 & "</meta_tag_type>"
                                        'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_VALUE" Then
                                        '    Print #99, "                <meta_tag_value>" & varcontent1 & "</meta_tag_value>"
                                        '    Print #99, "            </meta_tag>"
                                        End If
                                 ElseIf questiontype = "MC" And varspecialtype = True Then
                                 
                                        If para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Type" Then
                                            If varopencsms = False Then
                                               ' varopencsms = True
                                              Print #99, "<question_cs_sub_questions>"
                                            End If
                                              
                                            varIdentificationId = Replace(Replace(para.Range.Tables(vartabl).Rows(varrow + 1).Cells(varcol + 1).Range.Text, "Identification Id: ", ""), "", "")
                                            Print #99, "    <cs_sub_question identificationId=""" & varIdentificationId & """ qtype=""" & questiontype & """ qmode=""C"">"
                                            'Print #99, "        <cs_sub_question identificationId=""0"" qtype=""" & questiontype & """ qmode=""C"">"
                                            Print #99, "            <cs_sub_question_type ucx=""C"" >" & questiontypename & "</cs_sub_question_type>"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Title" Then
                                            Print #99, "            <cs_sub_question_title ucx=""C"">" & varcontent1 & "</cs_sub_question_title>"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Text" Then
                                            Print #99, "            <cs_sub_question_text ucx=""C"">" & varcontent1 & "</cs_sub_question_text>"
                                            Print #99, "            <cs_sub_question_multiple_choices qmcmode=""C"" tagtype=""" & tagtype & """ >"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Answer_Text" Then
                                            Print #99, "                <cs_sub_question_choice ucx=""C"" refId="""">"
                                            Print #99, "                    <cs_sub_question_answer_text>" & varcontent1 & "</cs_sub_question_answer_text>"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Rationale_Text" Then
                                            Print #99, "                    <cs_sub_question_rationale>" & varcontent1 & "</cs_sub_question_rationale>"
                                            Print #99, "                </cs_sub_question_choice>"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Correct_Answer" Then
                                            Print #99, "            </cs_sub_question_multiple_choices>"
                                            Print #99, "            <cs_sub_correct_answer>" & varcontent1 & "</cs_sub_correct_answer>"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Score" Then
                                            Print #99, "            <cs_sub_question_score ucx=""C"">" & varcontent1 & "</cs_sub_question_score>"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Difficulty" Then
                                            Print #99, "            <cs_sub_question_difficulty ucx=""C"" >" & varcontent1 & "</cs_sub_question_difficulty>"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Correct_Answer_Rationale" Then
                                            Print #99, "            <cs_sub_question_correct_rationale ucx=""C"">" & varcontent1 & "</cs_sub_question_correct_rationale>"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Incorrect_Answer_Rationale" Then
                                            Print #99, "            <cs_sub_question_incorrect_rationale ucx=""C"">" & varcontent1 & "</cs_sub_question_incorrect_rationale>"
                                        '    Print #99, "            <cs_sub_question_remediation_link qrlmode=""C"">"
                                        'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Text" Then
                                        '    Print #99, "                <cs_sub_remediation_type ucx=""C"" redLinkId=""1"" remediation_link_type=""Web Link"">"
                                        ''    Print #99, "                    <cs_sub_remediation_type_text>" & varcontent1 & "</cs_sub_remediation_type_text>"
                                        'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Link" Then
                                        '    Print #99, "                    <cs_sub_remediation_type_link>" & varcontent1 & "</cs_sub_remediation_type_link>"
                                        'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Tooltip" Then
                                        '    Print #99, "                    <cs_sub_remediation_type_tooltip>" & varcontent1 & "</cs_sub_remediation_type_tooltip>"
                                        '    Print #99, "                </cs_sub_remediation_type>"
                                        '    Print #99, "                 <cs_sub_remediation_type ucx=""C"" redLinkId=""2"" remediation_link_type=""Ebook""></cs_sub_remediation_type>"
                                        'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Text_Remediation_Text" Then
                                        '    Print #99, "                 <cs_sub_remediation_type ucx=""C"" redLinkId=""3"" remediation_link_type=""Text"">"
                                        ''    Print #99, "                    <cs_sub_remediation_type_text>" & varcontent1 & "</cs_sub_remediation_type_text>"
                                         '   Print #99, "                </cs_sub_remediation_type>"
                                         '   Print #99, "            </cs_sub_question_remediation_link>"
                                           'Print #99, "            </cs_sub_question>"
                                        End If
                                ElseIf questiontype = "CM" And varspecialtype = False Then
                                 If varcsms = True Then
                                    Print #99, "</question_cs_sub_questions>"
                                    Print #99, "</wk_question>"
                                    varcsms = False
                               End If
                                    varspecialtype = False
                                    If para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Type" Then
                                            varIdentificationId = Replace(Replace(para.Range.Tables(vartabl).Rows(varrow + 1).Cells(varcol + 1).Range.Text, "Identification Id: ", ""), "", "")
                                            Print #99, "    <wk_question identificationId=""" & varIdentificationId & """ qtype=""" & questiontype & """ qmode=""C"">"
                                            'Print #99, "    <wk_question identificationId=""0"" qtype=""" & questiontype & """ qmode=""C"">"
                                            Print #99, "        <question_type ucx=""C"" >" & questiontypename & "</question_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Title" Then
                                        Print #99, "        <question_title ucx=""C"">" & varcontent1 & "</question_title>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Text" Then
                                        Print #99, "        <question_text ucx=""C"">" & varcontent1 & "</question_text>"
                                        Print #99, "        <question_choices_multiple qmcmode=""C"" tagtype=""" & tagtype & """ >"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Answer_Text" Then
                                        Print #99, "            <question_choice ucx=""C"" refId="""">"
                                        Print #99, "                <question_answer_text>" & varcontent1 & "</question_answer_text>"
                                        'Print #99, "            </question_choice>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Rationale_Text" Then
                                        Print #99, "                <question_rationale>" & varcontent1 & "</question_rationale>"
                                        Print #99, "            </question_choice>"
                                    
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Correct_Answer" Then
                                        'ActiveDocument.FormFields("Check1").CheckBox.Value
                                        varcorrectcm = "<span1>"
                                        For K = 1 To para.Range.Tables(vartabl).Rows.Count - 16
                                            varCheckboxvalue = ActiveDocument.FormFields("Check_" & QuestionNo & "_" & K).CheckBox.Value
                                            If varCheckboxvalue = True Then
                                                varcorrectcm = Replace(varcorrectcm, "<span1>", K & ",<span1>")
                                            End If
                                        Next
                                        Print #99, "        </question_choices_multiple>"
                                        Print #99, "        <correct_answer>" & Replace(Replace(varcorrectcm, ",<span1>", ""), "<span1>", "") & "</correct_answer>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Score" Then
                                        Print #99, "        <question_score ucx=""C"">" & varcontent1 & "</question_score>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Difficulty" Then
                                        Print #99, "        <question_difficulty ucx=""C"" >" & varcontent1 & "</question_difficulty>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Correct_Answer_Rationale" Then
                                        Print #99, "        <question_correct_rationale ucx=""C"">" & varcontent1 & "</question_correct_rationale>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Incorrect_Answer_Rationale" Then
                                        Print #99, "        <question_incorrect_rationale ucx=""C"">" & varcontent1 & "</question_incorrect_rationale>"
                                      '  Print #99, "        <question_remediation_link qrlmode=""C"">"
                                    'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Text" Then
                                    '    Print #99, "            <remediation_type ucx=""C"" redLinkId=""1"" remediation_link_type=""Web Link"">"
                                    '    Print #99, "                <remediation_type_text>" & varcontent1 & "</remediation_type_text>"
                                    'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Link" Then
                                    '    Print #99, "                <remediation_type_link>" & varcontent1 & "</remediation_type_link>"
                                    'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Tooltip" Then
                                    ''    Print #99, "                <remediation_type_tooltip>" & varcontent1 & "</remediation_type_tooltip>"
                                    '    Print #99, "            </remediation_type>"
                                    '    Print #99, "             <remediation_type ucx=""C"" redLinkId=""2"" remediation_link_type=""Ebook""></remediation_type>"
                                    'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Text_Remediation_Text" Then
                                     '   Print #99, "             <remediation_type ucx=""C"" redLinkId=""3"" remediation_link_type=""Text"">"
                                     '   Print #99, "                <remediation_type_text>" & varcontent1 & "</remediation_type_text>"
                                     '   Print #99, "            </remediation_type>"
                                    'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Meta_Data_Attributes_head" Then
                                     '   Print #99, "        </question_remediation_link>"
                                     '   Print #99, "        <question_meta_tag qmtmode=""C"">"
                                    'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_TAG" Then
                                     '   Print #99, "            <meta_tag ucx=""C"" metaTagId="""">"
                                     '   Print #99, "                <meta_tag_type>" & varcontent1 & "</meta_tag_type>"
                                    'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_VALUE" Then
                                     '   Print #99, "                <meta_tag_value>" & varcontent1 & "</meta_tag_value>"
                                     '   Print #99, "            </meta_tag>"
                                    End If
                                ElseIf questiontype = "CM" And varspecialtype = True Then
                                    If para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Type" Then
                                            If varopencsms = False Then
                                                'varopencsms = True
                                                Print #99, "<question_cs_sub_questions>"
                                            End If
                                            varIdentificationId = Replace(Replace(para.Range.Tables(vartabl).Rows(varrow + 1).Cells(varcol + 1).Range.Text, "Identification Id: ", ""), "", "")
                                            Print #99, "    <cs_sub_question identificationId=""" & varIdentificationId & """ qtype=""" & questiontype & """ qmode=""C"">"
                                            'Print #99, "        <cs_sub_question identificationId=""0"" qtype=""" & questiontype & """ qmode=""C"">"
                                            Print #99, "            <cs_sub_question_type ucx=""C"" >" & questiontypename & "</cs_sub_question_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Title" Then
                                        Print #99, "            <cs_sub_question_title ucx=""C"">" & varcontent1 & "</cs_sub_question_title>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Text" Then
                                        Print #99, "            <cs_sub_question_text ucx=""C"">" & varcontent1 & "</cs_sub_question_text>"
                                        Print #99, "            <cs_sub_question_choices_multiple qmcmode=""C"" tagtype=""" & tagtype & """ >"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Answer_Text" Then
                                        Print #99, "                <cs_sub_question_choice ucx=""C"" refId="""">"
                                        Print #99, "                    <cs_sub_question_answer_text>" & varcontent1 & "</cs_sub_question_answer_text>"
                                        'Print #99, "                </question_choice>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Rationale_Text" Then
                                        Print #99, "                    <cs_sub_question_rationale>" & varcontent1 & "</cs_sub_question_rationale>"
                                        Print #99, "                </cs_sub_question_choice>"
                                    
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Correct_Answer" Then
                                        'ActiveDocument.FormFields("Check1").CheckBox.Value
                                        varcorrectcm = "<span1>"
                                        For K = 1 To para.Range.Tables(vartabl).Rows.Count - 16
                                            varCheckboxvalue = ActiveDocument.FormFields("Check_" & QuestionNo & "_" & K).CheckBox.Value
                                            If varCheckboxvalue = True Then
                                                varcorrectcm = Replace(varcorrectcm, "<span1>", K & ",<span1>")
                                            End If
                                        Next
                                        Print #99, "            </cs_sub_question_choices_multiple>"
                                        Print #99, "            <cs_sub_correct_answer>" & Replace(Replace(varcorrectcm, ",<span1>", ""), "<span1>", "") & "</cs_sub_correct_answer>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Score" Then
                                        Print #99, "            <cs_sub_question_score ucx=""C"">" & varcontent1 & "</cs_sub_question_score>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Difficulty" Then
                                        Print #99, "            <cs_sub_question_difficulty ucx=""C"" >" & varcontent1 & "</cs_sub_question_difficulty>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Correct_Answer_Rationale" Then
                                        Print #99, "            <cs_sub_question_correct_rationale ucx=""C"">" & varcontent1 & "</cs_sub_question_correct_rationale>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Incorrect_Answer_Rationale" Then
                                        Print #99, "            <cs_sub_question_incorrect_rationale ucx=""C"">" & varcontent1 & "</cs_sub_question_incorrect_rationale>"
                                        'Print #99, "            <cs_sub_question_remediation_link qrlmode=""C"">"
                                    'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Text" Then
                                     '   Print #99, "                <cs_sub_remediation_type ucx=""C"" redLinkId=""1"" remediation_link_type=""Web Link"">"
                                     '   Print #99, "                    <cs_sub_remediation_type_text>" & varcontent1 & "</cs_sub_remediation_type_text>"
                                    'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Link" Then
                                     '   Print #99, "                    <cs_sub_remediation_type_link>" & varcontent1 & "</cs_sub_remediation_type_link>"
                                    'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Tooltip" Then
                                     '   Print #99, "                    <cs_sub_remediation_type_tooltip>" & varcontent1 & "</cs_sub_remediation_type_tooltip>"
                                     '   Print #99, "                </cs_sub_remediation_type>"
                                     '   Print #99, "                 <cs_sub_remediation_type ucx=""C"" redLinkId=""2"" remediation_link_type=""Ebook""></cs_sub_remediation_type>"
                                   ' ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Text_Remediation_Text" Then
                                     '   Print #99, "                 <cs_sub_remediation_type ucx=""C"" redLinkId=""3"" remediation_link_type=""Text"">"
                                     '   Print #99, "                    <cs_sub_remediation_type_text>" & varcontent1 & "</cs_sub_remediation_type_text>"
                                     '   Print #99, "                </cs_sub_remediation_type>"
                                     '   Print #99, "            </cs_sub_question_remediation_link>"
                                        'Print #99, "            </cs_sub_question>"
                                    End If
                                ElseIf (questiontype = "TF") Then
                                 If varcsms = True Then
                                    Print #99, "</question_cs_sub_questions>"
                                    Print #99, "</wk_question>"
                                    varcsms = False
                               End If
                                    varspecialtype = False
                                    If para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Type" Then
                                            varIdentificationId = Replace(Replace(para.Range.Tables(vartabl).Rows(varrow + 1).Cells(varcol + 1).Range.Text, "Identification Id: ", ""), "", "")
                                            Print #99, "    <wk_question identificationId=""" & varIdentificationId & """ qtype=""" & questiontype & """ qmode=""C"">"
                                            'Print #99, "    <wk_question identificationId=""0"" qtype=""" & questiontype & """ qmode=""C"">"
                                            Print #99, "        <question_type ucx=""C"" >" & questiontypename & "</question_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Title" Then
                                        Print #99, "        <question_title ucx=""C"">" & varcontent1 & "</question_title>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Text" Then
                                        Print #99, "        <question_text ucx=""C"">" & varcontent1 & "</question_text>"
                                        Print #99, "        <question_true_false qmcmode=""C"" tagtype=""" & tagtype & """ >"
                                        
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Answer_Text" Then
                                        Print #99, "            <question_choice ucx=""C"" refId="""">"
                                        Print #99, "                <question_answer_text>" & varcontent1 & "</question_answer_text>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Rationale_Text" Then
                                        Print #99, "                <question_rationale>" & varcontent1 & "</question_rationale>"
                                        Print #99, "            </question_choice>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Correct_Answer" Then
                                       Print #99, "        </question_true_false>"
                                            If varcontent1 = True Then
                                                Print #99, "        <correct_answer>1</correct_answer>"
                                            Else
                                                Print #99, "        <correct_answer>2</correct_answer>"
                                            End If
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Score" Then
                                        Print #99, "        <question_score ucx=""C"">" & varcontent1 & "</question_score>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Difficulty" Then
                                        Print #99, "        <question_difficulty ucx=""C"" >" & varcontent1 & "</question_difficulty>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Correct_Answer_Rationale" Then
                                        Print #99, "        <question_correct_rationale ucx=""C"">" & varcontent1 & "</question_correct_rationale>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Incorrect_Answer_Rationale" Then
                                        Print #99, "        <question_incorrect_rationale ucx=""C"">" & varcontent1 & "</question_incorrect_rationale>"
                                       ' Print #99, "        <question_remediation_link qrlmode=""C"">"
                                    'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Text" Then
                                    ''    Print #99, "            <remediation_type ucx=""C"" redLinkId=""1"" remediation_link_type=""Web Link"">"
                                    '    Print #99, "                <remediation_type_text>" & varcontent1 & "</remediation_type_text>"
                                    'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Link" Then
                                    '    Print #99, "                <remediation_type_link>" & varcontent1 & "</remediation_type_link>"
                                    'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Tooltip" Then
                                    '    Print #99, "                <remediation_type_tooltip>" & varcontent1 & "</remediation_type_tooltip>"
                                    '    Print #99, "            </remediation_type>"
                                    '    Print #99, "             <remediation_type ucx=""C"" redLinkId=""2"" remediation_link_type=""Ebook""></remediation_type>"
                                    'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Text_Remediation_Text" Then
                                    '    Print #99, "             <remediation_type ucx=""C"" redLinkId=""3"" remediation_link_type=""Text"">"
                                    '    Print #99, "                <remediation_type_text>" & varcontent1 & "</remediation_type_text>"
                                    '    Print #99, "            </remediation_type>"
                                    'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Meta_Data_Attributes_head" Then
                                    '    Print #99, "        </question_remediation_link>"
                                    '    Print #99, "        <question_meta_tag qmtmode=""C"">"
                                    ''ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_TAG" Then
                                    '    Print #99, "            <meta_tag ucx=""C"" metaTagId="""">"
                                    '    Print #99, "                <meta_tag_type>" & varcontent1 & "</meta_tag_type>"
                                    'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_VALUE" Then
                                   '     Print #99, "                <meta_tag_value>" & varcontent1 & "</meta_tag_value>"
                                   '     Print #99, "            </meta_tag>"
                                    End If
                                ElseIf (questiontype = "VQ") Then
                                 If varcsms = True Then
                                    Print #99, "</question_cs_sub_questions>"
                                    Print #99, "</wk_question>"
                                    varcsms = False
                               End If
                                    'On Error Resume Next
                                    varspecialtype = False
                                    If para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Type" Then
                                            varIdentificationId = Replace(Replace(para.Range.Tables(vartabl).Rows(varrow + 1).Cells(varcol + 1).Range.Text, "Identification Id: ", ""), "", "")
                                            Print #99, "    <wk_question identificationId=""" & varIdentificationId & """ qtype=""" & questiontype & """ qmode=""C"">"
                                            'Print #99, "    <wk_question identificationId=""0"" qtype=""" & questiontype & """ qmode=""C"">"
                                            Print #99, "        <question_type ucx=""C"" >" & questiontypename & "</question_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Title" Then
                                        Print #99, "        <question_title ucx=""C"">" & varcontent1 & "</question_title>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Text" Then
                                        Print #99, "        <question_text ucx=""C"">" & varcontent1 & "</question_text>"
                                        Print #99, "        <question_video_questions qmcmode=""C"" tagtype=""" & tagtype & """ >"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Answer_Text" Then
                                        Print #99, "            <question_choice ucx=""C"" refId="""">"
                                        Print #99, "                <question_answer_text>" & varcontent1 & "</question_answer_text>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Rationale_Text" Then
                                        Print #99, "                <question_rationale>" & varcontent1 & "</question_rationale>"
                                        Print #99, "            </question_choice>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Correct_Answer" Then
                                        Print #99, "        </question_video_questions>"
                                        Print #99, "        <correct_answer>" & varcontent1 & "</correct_answer>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Additional_file_path" Then
                                        
                                        If UCase(FSO.GetExtensionName(varcontent1)) = "JPGE" Or UCase(FSO.GetExtensionName(varcontent1)) = "JPG" Then
                                            varMediaType = "image"
                                        ElseIf UCase(FSO.GetExtensionName(varcontent1)) = "MP4" Then
                                            varMediaType = "video"
                                        ElseIf UCase(FSO.GetExtensionName(varcontent1)) = "MP3" Then
                                            varMediaType = "audio"
                                        End If
                                        
                                        'Print #99, "            <question_choice ucx=""C"" refId="""">"
                                        Print #99, "        <question_additional_fields uck=""C"" referencevalue="""" mediaType=""" & varMediaType & """>"
                                        Print #99, "            <question_additional_file_path>media\" & varcontent1 & "</question_additional_file_path>"
                                        Print #99, "        </question_additional_fields>"
                                        If FSO.FolderExists(ActiveDocument.Path & "\media") = False Then
                                            Print #35, vbCrLf
                                            varmsgtext = varerrorno & ". " & varcontent1 & " - This file is not available in the media folder"
                                            Print #35, varmsgtext
                                            varerrorno = varerrorno + 1
                                        End If
                                        
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Score" Then
                                        Print #99, "        <question_score ucx=""C"">" & varcontent1 & "</question_score>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Difficulty" Then
                                        Print #99, "        <question_difficulty ucx=""C"" >" & varcontent1 & "</question_difficulty>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Correct_Answer_Rationale" Then
                                        Print #99, "        <question_correct_rationale ucx=""C"">" & varcontent1 & "</question_correct_rationale>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Incorrect_Answer_Rationale" Then
                                        Print #99, "        <question_incorrect_rationale ucx=""C"">" & varcontent1 & "</question_incorrect_rationale>"
                                        'Print #99, "        <question_remediation_link qrlmode=""C"">"
                                   'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Text" Then
                                   '     Print #99, "            <remediation_type ucx=""C"" redLinkId=""1"" remediation_link_type=""Web Link"">"
                                   '     Print #99, "                <remediation_type_text>" & varcontent1 & "</remediation_type_text>"
                                   ' ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Link" Then
                                   '     Print #99, "                <remediation_type_link>" & varcontent1 & "</remediation_type_link>"
                                   ' ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Tooltip" Then
                                   '     Print #99, "                <remediation_type_tooltip>" & varcontent1 & "</remediation_type_tooltip>"
                                   '     Print #99, "            </remediation_type>"
                                   '     Print #99, "             <remediation_type ucx=""C"" redLinkId=""2"" remediation_link_type=""Ebook""></remediation_type>"
                                   ' ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Text_Remediation_Text" Then
                                   '     Print #99, "             <remediation_type ucx=""C"" redLinkId=""3"" remediation_link_type=""Text"">"
                                   '     Print #99, "                <remediation_type_text>" & varcontent1 & "</remediation_type_text>"
                                   '     Print #99, "            </remediation_type>"
                                   ' ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Meta_Data_Attributes_head" Then
                                    '    Print #99, "        </question_remediation_link>"
                                    '    Print #99, "        <question_meta_tag qmtmode=""C"">"
                                   ' ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_TAG" Then
                                   '     Print #99, "            <meta_tag ucx=""C"" metaTagId="""">"
                                   '     Print #99, "                <meta_tag_type>" & varcontent1 & "</meta_tag_type>"
                                   ' ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_VALUE" Then
                                   '     Print #99, "                <meta_tag_value>" & varcontent1 & "</meta_tag_value>"
                                   '     Print #99, "            </meta_tag>"
                                    End If
                                
                                ElseIf (questiontype = "GO") Then
                                 If varcsms = True Then
                                    Print #99, "</question_cs_sub_questions>"
                                    Print #99, "</wk_question>"
                                    varcsms = False
                               End If
                                    'On Error Resume Next
                                    varspecialtype = False
                                    If para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Type" Then
                                            varIdentificationId = Replace(Replace(para.Range.Tables(vartabl).Rows(varrow + 1).Cells(varcol + 1).Range.Text, "Identification Id: ", ""), "", "")
                                            Print #99, "    <wk_question identificationId=""" & varIdentificationId & """ qtype=""" & questiontype & """ qmode=""C"">"
                                            'Print #99, "    <wk_question identificationId=""0"" qtype=""" & questiontype & """ qmode=""C"">"
                                            Print #99, "        <question_type ucx=""C"" >" & questiontypename & "</question_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Title" Then
                                        Print #99, "        <question_title ucx=""C"">" & varcontent1 & "</question_title>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Text" Then
                                        Print #99, "        <question_text ucx=""C"">" & varcontent1 & "</question_text>"
                                        Print #99, "        <question_graphic_option qmcmode=""C"" tagtype=""" & tagtype & """ >"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Additional_file_path" Then
                                        
                                        If UCase(FSO.GetExtensionName(varcontent1)) = "JPGE" Or UCase(FSO.GetExtensionName(varcontent1)) = "JPG" Then
                                            varMediaType = "image"
                                        ElseIf UCase(FSO.GetExtensionName(varcontent1)) = "MP4" Then
                                            varMediaType = "video"
                                        ElseIf UCase(FSO.GetExtensionName(varcontent1)) = "MP3" Then
                                            varMediaType = "audio"
                                        End If
                                        
                                        Print #99, "            <question_choice ucx=""C"" refId="""">"
                                        Print #99, "        <question_additional_fields uck=""C"" referencevalue="""" mediaType=""" & varMediaType & """>"
                                        Print #99, "            <question_additional_file_path>media\" & varcontent1 & "</question_additional_file_path>"
                                        Print #99, "        </question_additional_fields>"
                                        
                                        If FSO.FolderExists(ActiveDocument.Path & "\media") = False Then
                                            Print #35, vbCrLf
                                            varmsgtext = varerrorno & ". " & varcontent1 & " - This file is not available in the media folder"
                                            Print #35, varmsgtext
                                            varerrorno = varerrorno + 1
                                        End If
                                        
                                        'Print #99, "            <question_choice ucx=""C"" refId="""">"
                                        'Print #99, "                <question_answer_text>" & varcontent1 & "</question_answer_text>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Rationale_Text" Then
                                        Print #99, "                <question_rationale>" & varcontent1 & "</question_rationale>"
                                        Print #99, "            </question_choice>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Correct_Answer" Then
                                        Print #99, "        </question_graphic_option>"
                                        Print #99, "        <correct_answer>" & varcontent1 & "</correct_answer>"
                                    'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Additional_file_path" Then
                                        'Print #99, "        <question_additional_fields uck=""C"" referencevalue="""">"
                                        'Print #99, "            <question_additional_file_path>media\" & varcontent1 & "</question_additional_file_path>"
                                        'Print #99, "        </question_additional_fields>"
                                        
                                        'If fso.FolderExists(ActiveDocument.Path & "\media") = False Then
                                            'Print #35, vbCrLf
                                            'varmsgtext = varerrorno & ". " & varcontent1 & " - This file is not available in the media folder"
                                            'Print #35, varmsgtext
                                            'varerrorno = varerrorno + 1
                                        'End If
                                        
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Score" Then
                                        Print #99, "        <question_score ucx=""C"">" & varcontent1 & "</question_score>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Difficulty" Then
                                        Print #99, "        <question_difficulty ucx=""C"" >" & varcontent1 & "</question_difficulty>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Correct_Answer_Rationale" Then
                                        Print #99, "        <question_correct_rationale ucx=""C"">" & varcontent1 & "</question_correct_rationale>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Incorrect_Answer_Rationale" Then
                                        Print #99, "        <question_incorrect_rationale ucx=""C"">" & varcontent1 & "</question_incorrect_rationale>"
                                   '     Print #99, "        <question_remediation_link qrlmode=""C"">"
                                   ' ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Text" Then
                                   '     Print #99, "            <remediation_type ucx=""C"" redLinkId=""1"" remediation_link_type=""Web Link"">"
                                   '     Print #99, "                <remediation_type_text>" & varcontent1 & "</remediation_type_text>"
                                   ' ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Link" Then
                                   '     Print #99, "                <remediation_type_link>" & varcontent1 & "</remediation_type_link>"
                                   ' ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Tooltip" Then
                                   '     Print #99, "                <remediation_type_tooltip>" & varcontent1 & "</remediation_type_tooltip>"
                                   '     Print #99, "            </remediation_type>"
                                   '     Print #99, "             <remediation_type ucx=""C"" redLinkId=""2"" remediation_link_type=""Ebook""></remediation_type>"
                                   ' ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Text_Remediation_Text" Then
                                   '     Print #99, "             <remediation_type ucx=""C"" redLinkId=""3"" remediation_link_type=""Text"">"
                                   '     Print #99, "                <remediation_type_text>" & varcontent1 & "</remediation_type_text>"
                                   '     Print #99, "            </remediation_type>"
                                   ' ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Meta_Data_Attributes_head" Then
                                   '     Print #99, "        </question_remediation_link>"
                                   '     Print #99, "        <question_meta_tag qmtmode=""C"">"
                                   ' ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_TAG" Then
                                   '     Print #99, "            <meta_tag ucx=""C"" metaTagId="""">"
                                   '     Print #99, "                <meta_tag_type>" & varcontent1 & "</meta_tag_type>"
                                   ' ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_VALUE" Then
                                   '     Print #99, "                <meta_tag_value>" & varcontent1 & "</meta_tag_value>"
                                   '     Print #99, "            </meta_tag>"
                                    End If
                                
                                ElseIf (questiontype = "II") Then
                                If varcsms = True Then
                                    Print #99, "</question_cs_sub_questions>"
                                    Print #99, "</wk_question>"
                                    varcsms = False
                               End If
                                    'On Error Resume Next
                                    varspecialtype = False
                                    If para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Type" Then
                                            varIdentificationId = Replace(Replace(para.Range.Tables(vartabl).Rows(varrow + 1).Cells(varcol + 1).Range.Text, "Identification Id: ", ""), "", "")
                                            Print #99, "    <wk_question identificationId=""" & varIdentificationId & """ qtype=""" & questiontype & """ qmode=""C"">"
                                            'Print #99, "    <wk_question identificationId=""0"" qtype=""" & questiontype & """ qmode=""C"">"
                                            Print #99, "        <question_type ucx=""C"" >" & questiontypename & "</question_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Title" Then
                                        Print #99, "        <question_title ucx=""C"">" & varcontent1 & "</question_title>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Text" Then
                                        Print #99, "        <question_text ucx=""C"">" & varcontent1 & "</question_text>"
                                        Print #99, "        <question_image_integration qmcmode=""C"" tagtype=""" & tagtype & """ >"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Answer_Text" Then
                                        Print #99, "            <question_choice ucx=""C"" refId="""">"
                                        Print #99, "                <question_answer_text>" & varcontent1 & "</question_answer_text>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Rationale_Text" Then
                                        Print #99, "                <question_rationale>" & varcontent1 & "</question_rationale>"
                                        Print #99, "            </question_choice>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Correct_Answer" Then
                                        Print #99, "        </question_image_integration>"
                                        Print #99, "        <correct_answer>" & varcontent1 & "</correct_answer>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Additional_file_path" Then
                                     
                                        If UCase(FSO.GetExtensionName(varcontent1)) = "JPGE" Or UCase(FSO.GetExtensionName(varcontent1)) = "JPG" Then
                                            varMediaType = "image"
                                        ElseIf UCase(FSO.GetExtensionName(varcontent1)) = "MP4" Then
                                            varMediaType = "video"
                                        ElseIf UCase(FSO.GetExtensionName(varcontent1)) = "MP3" Then
                                            varMediaType = "audio"
                                        End If
                                        
                                        'Print #99, "            <question_choice ucx=""C"" refId="""">"
                                        Print #99, "        <question_additional_fields uck=""C"" referencevalue="""" mediaType=""" & varMediaType & """>"
                                        Print #99, "            <question_additional_file_path>media\" & varcontent1 & "</question_additional_file_path>"
                                        Print #99, "        </question_additional_fields>"
                                        
                                        If FSO.FolderExists(ActiveDocument.Path & "\media") = False Then
                                            Print #35, vbCrLf
                                            varmsgtext = varerrorno & ". " & varcontent1 & " - This file is not available in the media folder"
                                            Print #35, varmsgtext
                                            varerrorno = varerrorno + 1
                                        End If
                                        
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Score" Then
                                        Print #99, "        <question_score ucx=""C"">" & varcontent1 & "</question_score>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Difficulty" Then
                                        Print #99, "        <question_difficulty ucx=""C"" >" & varcontent1 & "</question_difficulty>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Correct_Answer_Rationale" Then
                                        Print #99, "        <question_correct_rationale ucx=""C"">" & varcontent1 & "</question_correct_rationale>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Incorrect_Answer_Rationale" Then
                                        Print #99, "        <question_incorrect_rationale ucx=""C"">" & varcontent1 & "</question_incorrect_rationale>"
                                   '     Print #99, "        <question_remediation_link qrlmode=""C"">"
                                   ' ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Text" Then
                                   '     Print #99, "            <remediation_type ucx=""C"" redLinkId=""1"" remediation_link_type=""Web Link"">"
                                   '     Print #99, "                <remediation_type_text>" & varcontent1 & "</remediation_type_text>"
                                 '   ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Link" Then
                                   '     Print #99, "                <remediation_type_link>" & varcontent1 & "</remediation_type_link>"
                                  '  ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Tooltip" Then
                                  '      Print #99, "                <remediation_type_tooltip>" & varcontent1 & "</remediation_type_tooltip>"
                                  '      Print #99, "            </remediation_type>"
                                  '      Print #99, "             <remediation_type ucx=""C"" redLinkId=""2"" remediation_link_type=""Ebook""></remediation_type>"
                                  '  ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Text_Remediation_Text" Then
                                  '      Print #99, "             <remediation_type ucx=""C"" redLinkId=""3"" remediation_link_type=""Text"">"
                                  '      Print #99, "                <remediation_type_text>" & varcontent1 & "</remediation_type_text>"
                                 '       Print #99, "            </remediation_type>"
                                  '  ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Meta_Data_Attributes_head" Then
                                  '      Print #99, "        </question_remediation_link>"
                                  ''      Print #99, "        <question_meta_tag qmtmode=""C"">"
                                   ' ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_TAG" Then
                                  ''      Print #99, "            <meta_tag ucx=""C"" metaTagId="""">"
                                   '     Print #99, "                <meta_tag_type>" & varcontent1 & "</meta_tag_type>"
                                   ' ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_VALUE" Then
                                   '     Print #99, "                <meta_tag_value>" & varcontent1 & "</meta_tag_value>"
                                   '     Print #99, "            </meta_tag>"
                                    End If
                                    
                                ElseIf (questiontype = "CS") Then
                             
                                varcsms = True
                                    'varCloseTag = True
                                    varspecialtype = True
                                    If para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Type" Then
                                            varIdentificationId = Replace(Replace(para.Range.Tables(vartabl).Rows(varrow + 1).Cells(varcol + 1).Range.Text, "Identification Id: ", ""), "", "")
                                            Print #99, "    <wk_question identificationId=""" & varIdentificationId & """ qtype=""" & questiontype & """ qmode=""C"">"
                                            'Print #99, "    <wk_question identificationId=""0"" qtype=""" & questiontype & """ qmode=""C"">"
                                            Print #99, "        <question_type ucx=""C"" >" & questiontypename & "</question_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Title" Then
                                        Print #99, "        <question_title ucx=""C"">" & varcontent1 & "</question_title>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Text" Then
                                        Print #99, "        <question_text ucx=""C"">" & varcontent1 & "</question_text>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Additional_file_path" Then
                                       If UCase(FSO.GetExtensionName(varcontent1)) = "JPGE" Or UCase(FSO.GetExtensionName(varcontent1)) = "JPG" Then
                                            varMediaType = "image"
                                        ElseIf UCase(FSO.GetExtensionName(varcontent1)) = "MP4" Then
                                            varMediaType = "video"
                                        ElseIf UCase(FSO.GetExtensionName(varcontent1)) = "MP3" Then
                                            varMediaType = "audio"
                                        End If
                                        
                                        'Print #99, "            <question_choice ucx=""C"" refId="""">"
                                        Print #99, "        <question_additional_fields uck=""C"" referencevalue="""" mediaType=""" & varMediaType & """>"
                                        Print #99, "            <question_additional_file_path>media\" & varcontent1 & "</question_additional_file_path>"
                                        Print #99, "        </question_additional_fields>"
                                    
                                        'Print #99, "        <question_additional_fields ucx=""C"">" & varcontent1 & "</question_additional_fields>"
                                        If FSO.FolderExists(ActiveDocument.Path & "\media") = False Then
                                            Print #35, vbCrLf
                                            varmsgtext = varerrorno & ". " & varcontent1 & " - This file is not available in the media folder"
                                            Print #35, varmsgtext
                                            varerrorno = varerrorno + 1
                                        End If
                                    'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Meta_Data_Attributes_head" Then
                                    '    Print #99, "        <question_meta_tag qmtmode=""C"">"
                                    'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_TAG" Then
                                    '    Print #99, "            <meta_tag ucx=""C"" metaTagId="""">"
                                    '    Print #99, "                <meta_tag_type>" & varcontent1 & "</meta_tag_type>"
                                    'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_VALUE" Then
                                    '    Print #99, "                <meta_tag_value>" & varcontent1 & "</meta_tag_value>"
                                    '    Print #99, "            </meta_tag>"
                                        'Print #99, "        </question_meta_tag>"
                                        'Print #99, "        <question_cs_sub_questions>"
                                    End If
                               ElseIf (questiontype = "MEDC") Then
                              
                               varcsms = True
                                    'On Error Resume Next
                                    'varCloseTag = True
                                    varspecialtype = True
                                    If para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Type" Then
                                            'Print #99, "        </question_cs_sub_questions>"
                                            'Print #99, "    </wk_question>"
                                            varIdentificationId = Replace(Replace(para.Range.Tables(vartabl).Rows(varrow + 1).Cells(varcol + 1).Range.Text, "Identification Id: ", ""), "", "")
                                            Print #99, "    <wk_question identificationId=""" & varIdentificationId & """ qtype=""" & questiontype & """ qmode=""C"">"
                                            'Print #99, "    <wk_question identificationId=""0"" qtype=""" & questiontype & """ qmode=""C"">"
                                            Print #99, "        <question_type ucx=""C"" >" & questiontypename & "</question_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Title" Then
                                        Print #99, "        <question_title ucx=""C"">" & varcontent1 & "</question_title>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Text" Then
                                        Print #99, "        <question_text ucx=""C"">" & varcontent1 & "</question_text>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Additional_file_path" Then
                                        If UCase(FSO.GetExtensionName(varcontent1)) = "JPGE" Or UCase(FSO.GetExtensionName(varcontent1)) = "JPG" Then
                                            varMediaType = "image"
                                        ElseIf UCase(FSO.GetExtensionName(varcontent1)) = "MP4" Then
                                            varMediaType = "video"
                                        ElseIf UCase(FSO.GetExtensionName(varcontent1)) = "MP3" Then
                                            varMediaType = "audio"
                                        End If
                                        
                                        'Print #99, "            <question_choice ucx=""C"" refId="""">"
                                        Print #99, "        <question_additional_fields uck=""C"" referencevalue="""" mediaType=""" & varMediaType & """>"
                                        Print #99, "            <question_additional_file_path>media\" & varcontent1 & "</question_additional_file_path>"
                                        Print #99, "        </question_additional_fields>"
                                        If FSO.FolderExists(ActiveDocument.Path & "\media") = False Then
                                            Print #35, vbCrLf
                                            varmsgtext = varerrorno & ". " & varcontent1 & " - This file is not available in the media folder"
                                            Print #35, varmsgtext
                                            varerrorno = varerrorno + 1
                                        End If
                                    'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Meta_Data_Attributes_head" Then
                                    '    Print #99, "        <question_meta_tag qmtmode=""C"">"
                                    'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_TAG" Then
                                    '    Print #99, "            <meta_tag ucx=""C"" metaTagId="""">"
                                    '    Print #99, "                <meta_tag_type>" & varcontent1 & "</meta_tag_type>"
                                    'ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_VALUE" Then
                                    '    Print #99, "                <meta_tag_value>" & varcontent1 & "</meta_tag_value>"
                                    '    Print #99, "            </meta_tag>"
                                        
                                        'Print #99, "        </question_meta_tag>"
                                        'Print #99, "        <question_cs_sub_questions>"
                                    End If
                                    
                                End If
                                End If
                         Next
                     Next
                'Print #99, "        <question_" & LCase(questiontype) & "_version>1</question_" & LCase(questiontype) & "_version>"
                'Print #99, "        <question_" & LCase(questiontype) & "_status>Create</question_" & LCase(questiontype) & "_status>"
                If questiontype = "MEDC" Or questiontype = "CS" Then
                'Print #99, "        </question_meta_tag>"
                'Print #99, "        <question_cs_sub_questions>"
                End If
                If varspecialtype = False Then
                 '   Print #99, "        </question_meta_tag>"
                    Print #99, "    </wk_question>"
               
                End If
                QuestionNo = QuestionNo + 1
E222:
                'End If
            End If
        End If
    End If
Next
    Print #99, "</wk_question_root>"
Close #99

                    Set txt2 = FSO.OpenTextFile(varpathout & "\allQuestions.xml")
                    vartxt4 = txt2.ReadAll
                    txt2.Close
                    vartxt5 = vartxt4
                    '</question_meta_tag>''<cs_sub_question '
                    'replace(vartxt5,"</cs_sub_question>" & vbnewline & "<question_cs_sub_questions>" & vbnewline & "    <cs_sub_question ","</cs_sub_question>" & vbnewline  & vbnewline & "    <cs_sub_question ")
                    vartxt5 = Replace(vartxt5, "</cs_sub_question>" & vbNewLine & "<question_cs_sub_questions>" & vbNewLine & "    <cs_sub_question ", "</cs_sub_question>" & vbNewLine & vbNewLine & "    <cs_sub_question ")
                    vartxt5 = Replace(vartxt5, "</question_meta_tag>" & vbNewLine & "    <cs_sub_question ", "</question_meta_tag>" & vbNewLine & "<question_cs_sub_questions>" & vbNewLine & "    <cs_sub_question ")
                    vartxt5 = Replace(vartxt5, "</cs_sub_question>" & vbNewLine & "    <wk_question", "</cs_sub_question>" & vbNewLine & "        </question_cs_sub_questions>" & vbNewLine & "        </wk_question>" & vbNewLine & "    <wk_question")
                    vartxt5 = Replace(vartxt5, "</cs_sub_question>" & vbNewLine & "</wk_question_root>", "</cs_sub_question>" & vbNewLine & "        </question_cs_sub_questions>" & vbNewLine & "        </wk_question>" & vbNewLine & "</wk_question_root>")
                    Open varpathout & "\allQuestions.xml" For Output As #55
                    Print #55, vartxt5
                    Close #55
' </question_cs_sub_questions>
        '</wk_question>


Call ValidateFile(ActiveDocument.Path & "\allQuestions.xml")
End Sub
Sub ValidateFile(strFile)
    'Create an XML DOMDocument object.
    'MsgBox ("here")
    Dim X As New DOMDocument
    'Load and validate the specified file into the DOM.
    X.async = False
    X.validateOnParse = True
    X.resolveExternals = True
    X.Load strFile
    'Return validation results in message to the user.
    If X.parseError.ErrorCode <> 0 Then
        ValidateFile1 = "Validation failed on " & _
                       strFile & vbCrLf & _
                       "=====================" & vbCrLf & _
                       "Reason: " & X.parseError.reason & _
                       vbCrLf & "Source: " & _
                       X.parseError.srcText & _
                       vbCrLf & "Line: " & _
                       X.parseError.Line & vbCrLf
        Print #35, vbCrLf
        varmsgtext = varerrorno & ". " & ValidateFile1
        Print #35, varmsgtext
        varerrorno = varerrorno + 1
    Else
        ValidateFile1 = "Validation succeeded for " & _
                       strFile & vbCrLf

        Print #35, vbCrLf
        varmsgtext = varerrorno & ". " & ValidateFile1
        Print #35, varmsgtext
        varerrorno = varerrorno + 1
    End If
    Close #35
End Sub
Sub SaveImageAsPicture()
    totalshapes = ActiveDocument.Shapes.Count
    If ActiveDocument.Shapes.Count > 0 Then
    'For Each Shape In ActiveDocument.Shapes
    i = 1
        Do
        Set Shape = ActiveDocument.InlineShapes(i)
                    sigObj = Shape.Name
                    Debug.Print Shape.Name
                    If InStr(Shape.Name, "Text Box") Then
                    'Debug.Print Shape.Name
                    ActiveDocument.InlineShapes(sigObj).Select
                   ' ActiveDocument.Shapes(sigObj).TextFrame.TextRange.Text = "<extra>" & ActiveDocument.Shapes(sigObj).TextFrame.TextRange.Text & "</extra>"
                    'ActiveDocument.Shapes(sigObj).Delete
                  '  Selection.Text = "ok"
        '                If Replace(Trim(ActiveDocument.InlineShapes(sigObj).TextRange.Text), Chr(13), "") <> "" Then
        '                    ActiveDocument.InlineShapes(sigObj).TextFrame.TextRange.Text = "<image_copy>" & ActiveDocument.InlineShapes(sigObj).TextFrame.TextRange.Text
        '                End If
                     Set oFrame = Shape.ConvertToFrame
                     'oFrame.Range.ParagraphFormat.Borders.OutsideLineStyle = wdLineStyleNone
                     oFrame.Range.ParagraphFormat.Borders.Enable = False
                     'oFrame.Delete
                     totalshapes = ActiveDocument.InlineShapes.Count
                     If i > 1 Then
                     i = i - 1
                     End If
                    Else
                     i = i + 1
                    End If
                    If i >= totalshapes Then Exit Do
        
        Loop
    End If
End Sub




Sub copytableMC(MyValue, QuestionNo, varTotalMetadata, varMetaDataName)
Attribute copytableMC.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
Dim table1 As Table
Dim tt As ContentControl
Dim TABLE2 As Table
Dim addbutton
varNumRows = MyValue + 6
'QuestionNo = QuestionNo + 1
Selection.HomeKey unit:=wdStory
Selection.EndKey unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 10, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.id = "MC"
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        
        table1.Rows(1).Select
         table1.Rows(1).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        
        table1.Cell(1, 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Delete_Question"
        tt.Title = "Delete Question"
        
        tt.SetPlaceholderText Text:="Delete Question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdRed
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        table1.Cell(1, 4).Select
        
        Selection.Collapse Direction:=wdCollapseStart
        Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
        With ffield
                .Name = "Delete_" & QuestionNo
                .Range.Style = "Delete_box"
        End With
        
        table1.Rows(2).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Question Text - Multiple Choice"
        tt.Range.Style = "Question_Type"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(3, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Question_no"
        tt.Title = "Question No"
        tt.SetPlaceholderText Text:="Question No: " & QuestionNo
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(3, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Identification_Id"
        tt.Title = "Identification Id"
        tt.SetPlaceholderText Text:="Identification Id: 0"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Rows(4).Select
        table1.Rows(4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(4, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(4, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Title"
        tt.Range.Style = "Question_Title"
        tt.SetPlaceholderText Text:="Enter question title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'Selection.PasteSpecial DataType:=wdPasteText
        'tt.Range.Collapse Direction:=wdCollapseStart
        'tt.Range.PasteSpecial DataType:=wdPasteText
        'tt.Selection.Copy
        'tt.Range.PasteAndFormat wdFormatPlainText
        'tt.Range.PasteSpecial DataType:=wdPasteText
                    
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        'Question Text *
        
        table1.Rows(5).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
      
        
         table1.Rows(6).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Text"
        tt.Range.Style = "Question_Text"
        tt.SetPlaceholderText Text:="Enter question text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 70
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        'Answer option heading
        table1.Rows(7).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Answer_head"
        tt.SetPlaceholderText Text:="Answer Text - Multiple Choice"
        
        tt.Range.Font.Size = "13"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
    
         
        
      For i = 8 To MyValue + 7
        
       ' table1.Rows(5).SetHeight rowHeight:=ActiveDocument.PageSetup.PageHeight - 300, HeightRule:=wdRowHeightExactly
        
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 6
        tt.Range.Style = "Answer_Text"
        tt.SetPlaceholderText Text:="Enter answer text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
         table1.Cell(i, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Rationale_Text"
        tt.Title = "Rationale Text " & i - 6
        tt.SetPlaceholderText Text:="Enter rationale for this answer"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
    Next
        
         'table1.Rows(4).Cells.Split NumRows:=5, NumColumns:=2, MergeBeforeSplit:=True
         'table1.Cell(4, 2).Merge table1.Cell(8, 2)
               
      
               
        table1.Rows((MyValue + 8)).Select
        table1.Rows((MyValue + 8)).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 8, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_answer_head"
        tt.SetPlaceholderText Text:="Correct Answer:"
        
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(MyValue + 8, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Title = "Correct Answer"
        tt.LockContentControl = True
        tt.Range.Style = "Correct_Answer"
        tt.SetPlaceholderText , , "Select your correct answer"
        For i = 1 To MyValue
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        'Set tt = Nothing
        
        table1.Rows(MyValue + 9).Select
        table1.Rows(MyValue + 9).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 9, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Score_head"
        tt.SetPlaceholderText Text:="Score"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(MyValue + 9, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Difficulty_head"
        tt.SetPlaceholderText Text:="Difficulty"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(MyValue + 10).Select
        table1.Rows(MyValue + 10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 10, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose score in the range 1 to 10"
        
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

        table1.Cell(MyValue + 10, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose difficulty in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(MyValue + 11).Select
        table1.Rows(MyValue + 11).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 11, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale_head"
        tt.SetPlaceholderText Text:="Correct Rationale"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(MyValue + 11, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Incorrect_Answer_Rationale_head"
         tt.SetPlaceholderText Text:="Incorrect Rationale"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(MyValue + 12).Select
        table1.Rows(MyValue + 12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 12, 1).Select
        
       
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale"
        tt.SetPlaceholderText Text:="Enter correct rationale for the question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(MyValue + 12, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.SetPlaceholderText Text:="Enter incorrect rationale for the question"
        tt.Range.Style = "Incorrect_Answer_Rationale"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
      meta_rem_head table1, table1.id, QuestionNo
  
        

End Sub
Sub meta_rem_head(table1 As Table, qtype, qno, Optional ByVal remidatiaon As Boolean = True, Optional ByVal metadata As Boolean = True)
  Dim VARNAME As String
Dim tt As ContentControl

If remidatiaon Then
'MsgBox "table1"
        ''''''''''''''''''''' Rem ''''''''''''''''''''''''''''
        table1.Rows(table1.Rows.Count - 3).Shading.BackgroundPatternColor = RGB(255, 140, 0)
        table1.Rows(table1.Rows.Count - 3).Select
        table1.Rows(table1.Rows.Count - 3).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Cell(table1.Rows.Count - 3, 1).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Links_head"
        tt.SetPlaceholderText Text:="Remediation"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 3, 2).Select
        Set RWL = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
        With RWL.OLEFormat.Object
        .Name = Replace(.Name, "CommandButton", "Rem_WL_cmd_qno" & qno & "_rdtWL_qt" & qtype)
        .Caption = "Weblink"
        .Font.Size = "13"
       
        .Font.Bold = True
        .Font.Name = "Verdana"
        .BackStyle = 0
        End With
      
        VARNAME = RWL.OLEFormat.Object.Name
        HeadingButtoncode VARNAME, table1, True, qno, qtype
        Set RWL = Nothing
        



        table1.Cell(table1.Rows.Count - 3, 3).Select
        Set REB = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
        'REB.OLEFormat.Object.Name = Replace(shp.OLEFormat.Object.Name, "CommandButton", "Rem_EB_cmd_qno" & qno & "_rdtEB_qt" & qtype)
        'REB.OLEFormat.Object.Caption = "EBoook"
         With REB.OLEFormat.Object
        .Name = Replace(.Name, "CommandButton", "Rem_EB_cmd_qno" & qno & "_rdtEB_qt" & qtype)
        .Caption = "EBoook"
        .Font.Size = "13"
        .Enabled = False
        .Font.Bold = True
        .Font.Name = "Verdana"
        '.BackStyle = 0
        End With
        
        VARNAME = REB.OLEFormat.Object.Name
        HeadingButtoncode VARNAME, table1, False, qtype, qno
        Set REB = Nothing
        
        table1.Cell(table1.Rows.Count - 3, 4).Select
        Set RText = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
           With RText.OLEFormat.Object
        .Name = Replace(.Name, "CommandButton", "Rem_Text_cmd_qno" & qno & "_rdtText_qt" & qtype)
        .Caption = "Text"
        .Font.Size = "13"
     
        .Font.Bold = True
        .Font.Name = "Verdana"
        .BackStyle = 0
        End With
        'RText.OLEFormat.Object.Name = Replace(shp.OLEFormat.Object.Name, "CommandButton", "Rem_Text_cmd_qno" & qno & "_rdtText_qt" & qtype)
        'RText.OLEFormat.Object.Caption = "Text"
         VARNAME = RText.OLEFormat.Object.Name
        HeadingButtoncode VARNAME, table1, False, qno, qtype
        Set RText = Nothing
End If
 If metadata Then
        '''''''''Metadata Attributes'''''''''
        
        table1.Rows(table1.Rows.Count - 1).Shading.BackgroundPatternColor = RGB(169, 169, 169)
        table1.Rows(table1.Rows.Count - 1).Select
        table1.Rows(table1.Rows.Count - 1).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Cell(table1.Rows.Count - 1, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Metadata Attributes"
        tt.Range.Style = "Meta_Data_Attributes_head"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 1, 2).Select
        Set RWL = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
        With RWL.OLEFormat.Object
        .Name = Replace(.Name, "CommandButton", "MD_FT_cmd_qno" & qno & "_MDtFT_qt" & qtype)
        .Caption = "Free Text"
        .Font.Size = "13"
        .AutoSize = True
        .Font.Bold = True
        .Font.Name = "Verdana"
        .BackStyle = 0
        End With
         VARNAME = RWL.OLEFormat.Object.Name
        HeadingButtoncode VARNAME, table1, True, qno, qtype
        Set RWL = Nothing
        
        table1.Cell(table1.Rows.Count - 1, 3).Select
        Set REB = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
        'REB.OLEFormat.Object.Name = Replace(shp.OLEFormat.Object.Name, "CommandButton", "Rem_EB_cmd_qno" & qno & "_rdtEB_qt" & qtype)
        'REB.OLEFormat.Object.Caption = "EBoook"
         With REB.OLEFormat.Object
        .Name = Replace(.Name, "CommandButton", "MD_LU_cmd_qno" & qno & "_MDtLU_qt" & qtype)
        .Caption = "Lookup"
        .AutoSize = True
        .Font.Size = "13"
        
        .Font.Bold = True
        .Font.Name = "Verdana"
        .BackStyle = 0
        End With
         VARNAME = REB.OLEFormat.Object.Name
        HeadingButtoncode VARNAME, table1, False, qno, qtype
        
        Set REB = Nothing
        
        table1.Cell(table1.Rows.Count - 1, 4).Select
        Set RText = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
           With RText.OLEFormat.Object
        .Name = Replace(.Name, "CommandButton", "MD_HI_cmd_qno" & qno & "_MDtHI_qt" & qtype)
        .Caption = "Hierarchy"
        .Font.Size = "13"
        .AutoSize = True
        .Font.Bold = True
        .Font.Name = "Verdana"
        .BackStyle = 0
        End With
        'RText.OLEFormat.Object.Name = Replace(shp.OLEFormat.Object.Name, "CommandButton", "Rem_Text_cmd_qno" & qno & "_rdtText_qt" & qtype)
        'RText.OLEFormat.Object.Caption = "Text"
         VARNAME = RText.OLEFormat.Object.Name
        HeadingButtoncode VARNAME, table1, False, qno, qtype
        Set RText = Nothing
 End If
End Sub
Sub copytablecs(MyValue, QuestionNo)
Dim table1 As Table
Dim tt As ContentControl
Dim TABLE2 As Table
Dim addbutton
'varNumRows = MyValue + 6
'QuestionNo = QuestionNo + 1
Selection.HomeKey unit:=wdStory
Selection.EndKey unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=8, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.id = "MC"
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        
        
            table1.Rows(1).Select
         table1.Rows(1).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        
        table1.Cell(1, 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Delete_Question"
        tt.Title = "Delete Question"
        
        tt.SetPlaceholderText Text:="Delete Question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdRed
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        table1.Cell(1, 4).Select
        
        Selection.Collapse Direction:=wdCollapseStart
        Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
        With ffield
                .Name = "Delete_" & QuestionNo
                .Range.Style = "Delete_box"
        End With
        
        table1.Rows(2).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Question Text - Clinical Symptoms"
        tt.Range.Style = "Question_Type"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(3, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Question_no"
        tt.Title = "Question No"
        tt.SetPlaceholderText Text:="Question No: " & QuestionNo
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(3, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Identification_Id"
        tt.Title = "Identification Id"
        tt.SetPlaceholderText Text:="Identification Id: 0"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Rows(4).Select
        table1.Rows(4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(4, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(4, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Title"
        tt.Range.Style = "Question_Title"
        tt.SetPlaceholderText Text:="Enter question title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
                
        table1.Rows(5).Select
        
        'table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Clinical Presenting Symptoms"
        tt.Range.Style = "Question_Text"
        tt.SetPlaceholderText Text:="Enter question text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 70
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
      
        
        table1.Rows(6).Select
        table1.Rows(6).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        table1.Cell(6, 1).TopPadding = 10
        table1.Cell(6, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Additional_Fields_head"
        tt.SetPlaceholderText Text:="Additional Fields"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(6, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Additional_Fields_head"
        tt.SetPlaceholderText Text:=""
        tt.LockContentControl = True
        
        table1.Cell(6, 3).Select
        'Selection.InlineShapes.AddOLEControl ClassType:="Forms.CommandButton.1"
        Set shp = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
        
        'shp.Range.Style = "Chosse_file"
        With shp.OLEFormat.Object
            .Object.Caption = "Select File"
            '.Object.Style = "Chosse_file"
        End With
        varbtn = Replace(shp.OLEFormat.Object.Name, "CommandButton", "")
        Set shp = Nothing
        
        
        table1.Cell(6, 4).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "File Name" & varbtn
        tt.Range.Style = "Additional_file_path"
        tt.Range.Font.Size = "10"
        tt.SetPlaceholderText , , ""
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
       table1.id = "CS"
       meta_rem_head table1, table1.id, QuestionNo, False
        
        

End Sub
Sub copytableCSMC(MyValue, QuestionNo)
Dim table1 As Table
Dim tt As ContentControl
Dim TABLE2 As Table
Dim addbutton
varNumRows = MyValue + 6
'QuestionNo = QuestionNo + 1
Selection.HomeKey unit:=wdStory
Selection.EndKey unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 10, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.id = "MC"
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        
            table1.Rows(1).Select
         table1.Rows(1).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        
        table1.Cell(1, 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Delete_Question"
        tt.Title = "Delete Question"
        
        tt.SetPlaceholderText Text:="Delete Question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdRed
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        table1.Cell(1, 4).Select
        
        Selection.Collapse Direction:=wdCollapseStart
        Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
        With ffield
                .Name = "Delete_" & QuestionNo
                .Range.Style = "Delete_box"
        End With
        
        
        table1.Rows(2).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Question Text - Multiple Choice"
        tt.Range.Style = "Question_Type"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(3, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Question_no"
        tt.Title = "Question No"
        tt.SetPlaceholderText Text:="Question No: " & QuestionNo
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(3, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Identification_Id"
        tt.Title = "Identification Id"
        tt.SetPlaceholderText Text:="Identification Id: 0"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Rows(4).Select
        table1.Rows(4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(4, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(4, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Title"
        tt.Range.Style = "Question_Title"
        tt.SetPlaceholderText Text:="Enter question title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
                
        table1.Rows(5).Select
          Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
         table1.Rows(6).Select
        'table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Text"
        tt.Range.Style = "Question_Text"
        tt.SetPlaceholderText Text:="Enter question text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 70
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        'Answer option heading
        table1.Rows(7).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Answer_head"
        tt.SetPlaceholderText Text:="Answer Text - Multiple Choice"
        
        tt.Range.Font.Size = "13"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
    
         
        
      For i = 8 To MyValue + 7
        
       ' table1.Rows(5).SetHeight rowHeight:=ActiveDocument.PageSetup.PageHeight - 300, HeightRule:=wdRowHeightExactly
        
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 6
        tt.Range.Style = "Answer_Text"
        tt.SetPlaceholderText Text:="Enter answer text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
         table1.Cell(i, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Rationale_Text"
        tt.Title = "Rationale Text " & i - 6
        tt.SetPlaceholderText Text:="Enter rationale for this answer"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
    Next
        
         'table1.Rows(4).Cells.Split NumRows:=5, NumColumns:=2, MergeBeforeSplit:=True
         'table1.Cell(4, 2).Merge table1.Cell(8, 2)
               
      
               'MsgBox "ok"
        table1.Rows(MyValue + 8).Select
        table1.Rows(MyValue + 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 8, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_answer_head"
        tt.SetPlaceholderText Text:="Correct Answer:"
        
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(MyValue + 8, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Title = "Correct Answer"
        tt.LockContentControl = True
        tt.Range.Style = "Correct_Answer"
        tt.SetPlaceholderText , , "Select your correct answer"
        For i = 1 To MyValue
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        'Set tt = Nothing
        
        table1.Rows(MyValue + 9).Select
        table1.Rows(MyValue + 9).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 9, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Score_head"
        tt.SetPlaceholderText Text:="Score"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(MyValue + 9, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Difficulty_head"
        tt.SetPlaceholderText Text:="Difficulty"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(MyValue + 10).Select
        table1.Rows(MyValue + 10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 10, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose score in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

        table1.Cell(MyValue + 10, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose difficulty in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(MyValue + 11).Select
        table1.Rows(MyValue + 11).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 11, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale_head"
        tt.SetPlaceholderText Text:="Correct Rationale"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(MyValue + 11, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Incorrect_Answer_Rationale_head"
         tt.SetPlaceholderText Text:="Incorrect Rationale"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(MyValue + 12).Select
        table1.Rows(MyValue + 12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 12, 1).Select
        
       
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale"
        tt.SetPlaceholderText Text:="Enter correct rationale for the question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(MyValue + 12, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.SetPlaceholderText Text:="Enter incorrect rationale for the question"
        tt.Range.Style = "Incorrect_Answer_Rationale"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
    'table1.Rows.Add
    'table1.Rows.Add
    'table1.Rows.Add
         table1.id = "CSMC"
        meta_rem_head table1, table1.id, QuestionNo, True, False
        table1.Rows(table1.Rows.Count).Delete
        table1.Rows(table1.Rows.Count).Delete
        
        
        

End Sub

Sub copytableCSCM(MyValue, QuestionNo)
Dim table1 As Table
Dim tt As ContentControl
Dim TABLE2 As Table
Dim addbutton
varNumRows = MyValue + 6
'QuestionNo = QuestionNo + 1
Selection.HomeKey unit:=wdStory
Selection.EndKey unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 10, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        
        
            table1.Rows(1).Select
         table1.Rows(1).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        
        table1.Cell(1, 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Delete_Question"
        tt.Title = "Delete Question"
        
        tt.SetPlaceholderText Text:="Delete Question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdRed
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        table1.Cell(1, 4).Select
        
        Selection.Collapse Direction:=wdCollapseStart
        Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
        With ffield
                .Name = "Delete_" & QuestionNo
                .Range.Style = "Delete_box"
        End With
        
        table1.Rows(2).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Question Text - Choice Multiple"
        tt.Range.Style = "Question_Type"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
         table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(3, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Question_no"
        tt.Title = "Question No"
        tt.SetPlaceholderText Text:="Question No: " & QuestionNo
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(3, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Identification_Id"
        tt.Title = "Identification Id"
        tt.SetPlaceholderText Text:="Identification Id: 0"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        
        table1.Rows(4).Select
        table1.Rows(4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(4, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(4, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Title"
        tt.Range.Style = "Question_Title"
        tt.SetPlaceholderText Text:="Enter question title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
        table1.Rows(5).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
         table1.Rows(6).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Text"
        tt.Range.Style = "Question_Text"
        tt.SetPlaceholderText Text:="Enter question text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        tt.Range.Rows.Height = 70
        Set tt = Nothing
        
        'Answer option heading
        table1.Rows(7).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Answer_head"
        tt.SetPlaceholderText Text:="Answer Text - Choice Multiple"
        
        tt.Range.Font.Size = "13"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
      
        
       ' table1.Rows(5).SetHeight rowHeight:=ActiveDocument.PageSetup.PageHeight - 300, HeightRule:=wdRowHeightExactly
        For i = 8 To MyValue + 7
            
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 6
        tt.Range.Style = "Answer_Text"
        tt.SetPlaceholderText Text:="Enter answer text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
         table1.Cell(i, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Rationale_Text"
        tt.Title = "Rationale Text " & i - 6
        tt.SetPlaceholderText Text:="Enter rationale for this answer"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
         
        Next
               
        table1.Rows(MyValue + 8).Select
        table1.Rows(MyValue + 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 8, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Correct Answers:"
        tt.Range.Style = "Correct_Answer"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        For i = 2 To (MyValue * 2) + 1
        
        If i = 2 Then
            j = 1
            K = 1
            table1.Cell(MyValue + 8, i).Split NumColumns:=(MyValue * 2) + 1
            table1.Cell(MyValue + 8, 2).Select
            Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
            tt.SetPlaceholderText Text:=" "
            tt.Range.Style = "Check_box_no"
            tt.Range.Font.Size = "10"
            tt.Range.Font.Name = "Verdana"
            tt.LockContentControl = True
            tt.LockContents = False
            
        End If
        
        If i Mod 2 > 0 Then
            table1.Cell(MyValue + 8, i).Select
            
            Selection.Collapse Direction:=wdCollapseEnd
            Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
            With ffield
                .Name = "Check_" & QuestionNo & "_" & K
                .Range.Style = "Check_box"
                '.CheckBox.Value = False
            End With
            K = K + 1
        Else
            table1.Cell(MyValue + 8, i).Select
            Selection.Collapse Direction:=wdCollapseEnd
            Set tt = ActiveDocument.ContentControls.Add(wdContentControlText)
                
            tt.SetPlaceholderText Text:=j
            tt.Range.Style = "Check_box_no"
            Set tt = Nothing
            j = j + 1
        End If
        
        Next
        
        
        'Set tt = Nothing
        table1.Rows(MyValue + 9).Select
        table1.Rows(MyValue + 9).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 9, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Score_head"
        tt.SetPlaceholderText Text:="Score"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(MyValue + 9, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Difficulty_head"
        tt.SetPlaceholderText Text:="Difficulty"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(MyValue + 10).Select
        table1.Rows(MyValue + 10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 10, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose score in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

    table1.Cell(MyValue + 10, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose difficulty in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(MyValue + 11).Select
        table1.Rows(MyValue + 11).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 11, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale_head"
        tt.SetPlaceholderText Text:="Correct Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(MyValue + 11, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Incorrect_Answer_Rationale_head"
         tt.SetPlaceholderText Text:="Incorrect Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(MyValue + 12).Select
        table1.Rows(MyValue + 12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 12, 1).Select
        
       
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale"
        tt.SetPlaceholderText Text:="Enter correct rationale for the question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(MyValue + 12, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Incorrect_Answer_Rationale"
        tt.SetPlaceholderText Text:="Enter incorrect rationale for the question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
          
       

    
    table1.id = "CSCM"
    meta_rem_head table1, table1.id, QuestionNo, True, False

       ' table1.Rows(table1.Rows.Count).Delete
       ' table1.Rows(table1.Rows.Count).Delete
End Sub

Sub copytableMCase(MyValue, QuestionNo)
Dim table1 As Table
Dim tt As ContentControl
Dim TABLE2 As Table
Dim addbutton
'varNumRows = MyValue + 6
'QuestionNo = QuestionNo + 1
Selection.HomeKey unit:=wdStory
Selection.EndKey unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=9, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        
            table1.Rows(1).Select
         table1.Rows(1).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        
        table1.Cell(1, 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Delete_Question"
        tt.Title = "Delete Question"
        
        tt.SetPlaceholderText Text:="Delete Question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdRed
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        table1.Cell(1, 4).Select
        
        Selection.Collapse Direction:=wdCollapseStart
        Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
        With ffield
                .Name = "Delete_" & QuestionNo
                .Range.Style = "Delete_box"
        End With
        
        table1.Rows(2).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Question Text - Medical Case"
        tt.Range.Style = "Question_Type"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(3, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Question_no"
        tt.Title = "Question No"
        tt.SetPlaceholderText Text:="Question No: " & QuestionNo
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(3, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Identification_Id"
        tt.Title = "Identification Id"
        tt.SetPlaceholderText Text:="Identification Id: 0"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Rows(4).Select
        table1.Rows(4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(4, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(4, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Title"
        tt.Range.Style = "Question_Title"
        tt.SetPlaceholderText Text:="Enter question title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
                
        table1.Rows(5).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
         table1.Rows(6).Select
        'table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Clinical Presenting Symptoms"
        tt.Range.Style = "Question_Text"
        tt.SetPlaceholderText Text:="Enter question text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 70
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
      
        
        table1.Rows(7).Select
        table1.Rows(7).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        table1.Cell(7, 1).TopPadding = 10
        table1.Cell(7, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Additional_Fields_head"
        tt.SetPlaceholderText Text:="Additional Fields"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(7, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Additional_Fields_head"
        tt.SetPlaceholderText Text:=""
        tt.LockContentControl = True
        
        table1.Cell(7, 3).Select
        'Selection.InlineShapes.AddOLEControl ClassType:="Forms.CommandButton.1"
        Set shp = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
        
        'shp.Range.Style = "Chosse_file"
        With shp.OLEFormat.Object
            .Object.Caption = "Select File"
            '.Object.Style = "Chosse_file"
        End With
        varbtn = Replace(shp.OLEFormat.Object.Name, "CommandButton", "")
        Set shp = Nothing
        
        
        table1.Cell(7, 4).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "File Name" & varbtn
        tt.Range.Style = "Additional_file_path"
        tt.Range.Font.Size = "10"
        tt.SetPlaceholderText , , ""
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
       table1.id = "MS"
       meta_rem_head table1, table1.id, QuestionNo, False
        

End Sub
Sub copytableMCaseMC(MyValue, QuestionNo)
Dim table1 As Table
Dim tt As ContentControl
Dim TABLE2 As Table
Dim addbutton
varNumRows = MyValue + 6
'QuestionNo = QuestionNo + 1
Selection.HomeKey unit:=wdStory
Selection.EndKey unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak
'MsgBox "ok"
   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 10, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        
        table1.Rows(1).Select
         table1.Rows(1).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        
        table1.Cell(1, 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Delete_Question"
        tt.Title = "Delete Question"
        
        tt.SetPlaceholderText Text:="Delete Question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdRed
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        table1.Cell(1, 4).Select
        
        Selection.Collapse Direction:=wdCollapseStart
        Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
        With ffield
                .Name = "Delete_" & QuestionNo
                .Range.Style = "Delete_box"
        End With
        
        
        table1.Rows(2).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Question Text - Multiple Choice"
        tt.Range.Style = "Question_Type"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(3, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Question_no"
        tt.Title = "Question No"
        tt.SetPlaceholderText Text:="Question No: " & QuestionNo
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(3, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Identification_Id"
        tt.Title = "Identification Id"
        tt.SetPlaceholderText Text:="Identification Id: 0"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Rows(4).Select
        table1.Rows(4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(4, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(4, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Title"
        tt.Range.Style = "Question_Title"
        tt.SetPlaceholderText Text:="Enter question title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
                
        table1.Rows(5).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
         table1.Rows(6).Select
        'table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Text"
        tt.Range.Style = "Question_Text"
        tt.SetPlaceholderText Text:="Enter question text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 70
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        'Answer option heading
        table1.Rows(7).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Answer_head"
        tt.SetPlaceholderText Text:="Answer Text - Multiple Choice"
        
        tt.Range.Font.Size = "13"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
    
         
        
      For i = 8 To MyValue + 7
        
       ' table1.Rows(5).SetHeight rowHeight:=ActiveDocument.PageSetup.PageHeight - 300, HeightRule:=wdRowHeightExactly
        
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 6
        tt.Range.Style = "Answer_Text"
        tt.SetPlaceholderText Text:="Enter answer text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
         table1.Cell(i, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Rationale_Text"
        tt.Title = "Rationale Text " & i - 6
        tt.SetPlaceholderText Text:="Enter rationale for this answer"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
    Next
        
         'table1.Rows(4).Cells.Split NumRows:=5, NumColumns:=2, MergeBeforeSplit:=True
         'table1.Cell(4, 2).Merge table1.Cell(8, 2)
               
      
               
        table1.Rows(MyValue + 8).Select
        table1.Rows(MyValue + 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 8, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_answer_head"
        tt.SetPlaceholderText Text:="Correct Answer:"
        
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(MyValue + 8, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Title = "Correct Answer"
        tt.LockContentControl = True
        tt.Range.Style = "Correct_Answer"
        tt.SetPlaceholderText , , "Select your correct answer"
        For i = 1 To MyValue
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        'Set tt = Nothing
        
        table1.Rows(MyValue + 9).Select
        table1.Rows(MyValue + 9).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 9, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Score_head"
        tt.SetPlaceholderText Text:="Score"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(MyValue + 9, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Difficulty_head"
        tt.SetPlaceholderText Text:="Difficulty"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(MyValue + 10).Select
        table1.Rows(MyValue + 10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 10, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose score in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

        table1.Cell(MyValue + 10, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose difficulty in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(MyValue + 11).Select
        table1.Rows(MyValue + 11).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 11, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale_head"
        tt.SetPlaceholderText Text:="Correct Rationale"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(MyValue + 11, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Incorrect_Answer_Rationale_head"
         tt.SetPlaceholderText Text:="Incorrect Rationale"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
    
        
        table1.Rows(MyValue + 12).Select
        table1.Rows(MyValue + 12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 12, 1).Select
        
       
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale"
        tt.SetPlaceholderText Text:="Enter correct rationale for the question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(MyValue + 12, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.SetPlaceholderText Text:="Enter incorrect rationale for the question"
        tt.Range.Style = "Incorrect_Answer_Rationale"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing

  table1.id = "MCaseMC"
         meta_rem_head table1, table1.id, QuestionNo, True, False
        
        
        

End Sub

Sub copytableMCaseCM(MyValue, QuestionNo)
Dim table1 As Table
Dim tt As ContentControl
Dim TABLE2 As Table
Dim addbutton
varNumRows = MyValue + 6
'QuestionNo = QuestionNo + 1
Selection.HomeKey unit:=wdStory
Selection.EndKey unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 10, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        
            table1.Rows(1).Select
         table1.Rows(1).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        
        table1.Cell(1, 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Delete_Question"
        tt.Title = "Delete Question"
        
        tt.SetPlaceholderText Text:="Delete Question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdRed
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        table1.Cell(1, 4).Select
        
        Selection.Collapse Direction:=wdCollapseStart
        Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
        With ffield
                .Name = "Delete_" & QuestionNo
                .Range.Style = "Delete_box"
        End With
        
        
        table1.Rows(2).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Question Text - Choice Multiple"
        tt.Range.Style = "Question_Type"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
         table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(3, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Question_no"
        tt.Title = "Question No"
        tt.SetPlaceholderText Text:="Question No: " & QuestionNo
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(3, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Identification_Id"
        tt.Title = "Identification Id"
        tt.SetPlaceholderText Text:="Identification Id: 0"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Rows(4).Select
        table1.Rows(4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(4, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(4, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Title"
        tt.Range.Style = "Question_Title"
        tt.SetPlaceholderText Text:="Enter question title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
        table1.Rows(5).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
         table1.Rows(6).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Text"
        tt.Range.Style = "Question_Text"
        tt.SetPlaceholderText Text:="Enter question text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        tt.Range.Rows.Height = 70
        Set tt = Nothing
        
        'Answer option heading
        table1.Rows(7).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Answer_head"
        tt.SetPlaceholderText Text:="Answer Text - Choice Multiple"
        
        tt.Range.Font.Size = "13"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
      
        
       ' table1.Rows(5).SetHeight rowHeight:=ActiveDocument.PageSetup.PageHeight - 300, HeightRule:=wdRowHeightExactly
        For i = 8 To MyValue + 7
            
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 6
        tt.Range.Style = "Answer_Text"
        tt.SetPlaceholderText Text:="Enter answer text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
         table1.Cell(i, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Rationale_Text"
        tt.Title = "Rationale Text " & i - 6
        tt.SetPlaceholderText Text:="Enter rationale for this answer"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
         
        Next
               
        table1.Rows(MyValue + 8).Select
        table1.Rows(MyValue + 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 8, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Correct Answers:"
        tt.Range.Style = "Correct_Answer"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        For i = 2 To (MyValue * 2) + 1
        
        If i = 2 Then
            j = 1
            K = 1
            table1.Cell(MyValue + 8, i).Split NumColumns:=(MyValue * 2) + 1
            table1.Cell(MyValue + 8, 2).Select
            Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
            tt.SetPlaceholderText Text:=" "
            tt.Range.Style = "Check_box_no"
            tt.Range.Font.Size = "10"
            tt.Range.Font.Name = "Verdana"
            tt.LockContentControl = True
            tt.LockContents = False
            
        End If
        
        If i Mod 2 > 0 Then
            table1.Cell(MyValue + 8, i).Select
            
            Selection.Collapse Direction:=wdCollapseEnd
            Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
            With ffield
                .Name = "Check_" & QuestionNo & "_" & K
                .Range.Style = "Check_box"
                '.CheckBox.Value = False
            End With
            K = K + 1
        Else
            table1.Cell(MyValue + 8, i).Select
            Selection.Collapse Direction:=wdCollapseEnd
            Set tt = ActiveDocument.ContentControls.Add(wdContentControlText)
                
            tt.SetPlaceholderText Text:=j
            tt.Range.Style = "Check_box_no"
            Set tt = Nothing
            j = j + 1
        End If
        
        Next
        
        
        'Set tt = Nothing
        table1.Rows(MyValue + 9).Select
        table1.Rows(MyValue + 9).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 9, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Score_head"
        tt.SetPlaceholderText Text:="Score"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(MyValue + 9, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Difficulty_head"
        tt.SetPlaceholderText Text:="Difficulty"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(MyValue + 10).Select
        table1.Rows(MyValue + 10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 10, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose score in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

    table1.Cell(MyValue + 10, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose difficulty in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(MyValue + 11).Select
        table1.Rows(MyValue + 11).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 11, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale_head"
        tt.SetPlaceholderText Text:="Correct Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(MyValue + 11, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Incorrect_Answer_Rationale_head"
         tt.SetPlaceholderText Text:="Incorrect Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(MyValue + 12).Select
        table1.Rows(MyValue + 12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 12, 1).Select
        
       
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale"
        tt.SetPlaceholderText Text:="Enter correct rationale for the question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(MyValue + 12, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Incorrect_Answer_Rationale"
        tt.SetPlaceholderText Text:="Enter incorrect rationale for the question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
    
  'CaseCM
  table1.id = "CaseCM"
       meta_rem_head table1, table1.id, QuestionNo, True, False
End Sub
Sub copytablecm(MyValue, QuestionNo)
Dim table1 As Table
Dim tt As ContentControl
Dim TABLE2 As Table
Dim addbutton
varNumRows = MyValue + 6
'QuestionNo = QuestionNo + 1
Selection.HomeKey unit:=wdStory
Selection.EndKey unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 10, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        
        table1.Rows(1).Select
         table1.Rows(1).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        
        table1.Cell(1, 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Delete_Question"
        tt.Title = "Delete Question"
        
        tt.SetPlaceholderText Text:="Delete Question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdRed
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        table1.Cell(1, 4).Select
        
        Selection.Collapse Direction:=wdCollapseStart
        Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
        With ffield
                .Name = "Delete_" & QuestionNo
                .Range.Style = "Delete_box"
        End With
        
        table1.Rows(2).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(2).Select
       
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Question Text - Choice Multiple"
        tt.Range.Style = "Question_Type"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
         table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(3, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Question_no"
        tt.Title = "Question No"
        tt.SetPlaceholderText Text:="Question No: " & QuestionNo
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(3, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Identification_Id"
        tt.Title = "Identification Id"
        tt.SetPlaceholderText Text:="Identification Id: 0"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Rows(4).Select
        table1.Rows(4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(4, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(4, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Title"
        tt.Range.Style = "Question_Title"
        tt.SetPlaceholderText Text:="Enter question title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
        table1.Rows(5).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Rows(6).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Text"
        tt.Range.Style = "Question_Text"
        tt.SetPlaceholderText Text:="Enter question text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        tt.Range.Rows.Height = 70
        Set tt = Nothing
        
        'Answer option heading
        table1.Rows(7).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Answer_head"
        tt.SetPlaceholderText Text:="Answer Text - Choice Multiple"
        
        tt.Range.Font.Size = "13"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
      
        
       ' table1.Rows(5).SetHeight rowHeight:=ActiveDocument.PageSetup.PageHeight - 300, HeightRule:=wdRowHeightExactly
        For i = 8 To MyValue + 7
            
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 6
        tt.Range.Style = "Answer_Text"
        tt.SetPlaceholderText Text:="Enter answer text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
         table1.Cell(i, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Rationale_Text"
        tt.Title = "Rationale Text " & i - 6
        tt.SetPlaceholderText Text:="Enter rationale for this answer"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
         
        Next
               
        table1.Rows(MyValue + 8).Select
        table1.Rows(MyValue + 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 8, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Correct Answers:"
        tt.Range.Style = "Correct_Answer"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        For i = 2 To (MyValue * 2) + 1
        
        If i = 2 Then
            j = 1
            K = 1
            table1.Cell(MyValue + 8, i).Split NumColumns:=(MyValue * 2) + 1
            table1.Cell(MyValue + 8, 2).Select
            Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
            tt.SetPlaceholderText Text:=" "
            tt.Range.Style = "Check_box_no"
            tt.Range.Font.Size = "10"
            tt.Range.Font.Name = "Verdana"
            tt.LockContentControl = True
            tt.LockContents = False
            
        End If
        
        If i Mod 2 > 0 Then
            table1.Cell(MyValue + 8, i).Select
            
            Selection.Collapse Direction:=wdCollapseEnd
            Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
            With ffield
                .Name = "Check_" & QuestionNo & "_" & K
                .Range.Style = "Check_box"
                '.CheckBox.Value = False
            End With
            K = K + 1
        Else
            table1.Cell(MyValue + 8, i).Select
            Selection.Collapse Direction:=wdCollapseEnd
            Set tt = ActiveDocument.ContentControls.Add(wdContentControlText)
                
            tt.SetPlaceholderText Text:=j
            tt.Range.Style = "Check_box_no"
            Set tt = Nothing
            j = j + 1
        End If
        
        Next
        
        
        'Set tt = Nothing
        table1.Rows(MyValue + 9).Select
        table1.Rows(MyValue + 9).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 9, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Score_head"
        tt.SetPlaceholderText Text:="Score"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(MyValue + 9, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Difficulty_head"
        tt.SetPlaceholderText Text:="Difficulty"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(MyValue + 10).Select
        table1.Rows(MyValue + 10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 10, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose score in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

    table1.Cell(MyValue + 10, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose difficulty in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(MyValue + 11).Select
        table1.Rows(MyValue + 11).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 11, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale_head"
        tt.SetPlaceholderText Text:="Correct Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(MyValue + 11, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Incorrect_Answer_Rationale_head"
         tt.SetPlaceholderText Text:="Incorrect Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(MyValue + 12).Select
        table1.Rows(MyValue + 12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 12, 1).Select
        
       
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale"
        tt.SetPlaceholderText Text:="Enter correct rationale for the question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(MyValue + 12, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Incorrect_Answer_Rationale"
        tt.SetPlaceholderText Text:="Enter incorrect rationale for the question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        table1.id = "CM"
        meta_rem_head table1, table1.id, QuestionNo
End Sub

Sub copytabletf(QuestionNo)
Dim table1 As Table
Dim tt As ContentControl
Dim TABLE2 As Table
Dim addbutton
'QuestionNo = QuestionNo + 1
Selection.HomeKey unit:=wdStory
Selection.EndKey unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=8 + 10, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        
        table1.Rows(1).Select
         table1.Rows(1).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        
        table1.Cell(1, 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Delete_Question"
        tt.Title = "Delete Question"
        
        tt.SetPlaceholderText Text:="Delete Question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdRed
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        table1.Cell(1, 4).Select
        
        Selection.Collapse Direction:=wdCollapseStart
        Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
        With ffield
                .Name = "Delete_" & QuestionNo
                .Range.Style = "Delete_box"
        End With
        
        table1.Rows(2).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Question Text - True or Fasle"
        tt.Range.Style = "Question_Type"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
         table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(3, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Question_no"
        tt.Title = "Question No"
        tt.SetPlaceholderText Text:="Question No: " & QuestionNo
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(3, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Identification_Id"
        tt.Title = "Identification Id"
        tt.SetPlaceholderText Text:="Identification Id: 0"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Rows(4).Select
        table1.Rows(4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(4, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(4, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Title"
        tt.Range.Style = "Question_Title"
        tt.SetPlaceholderText Text:="Enter question title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
        table1.Rows(5).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Rows(6).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Text"
        tt.Range.Style = "Question_Text"
        tt.SetPlaceholderText Text:="Enter question text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        tt.Range.Rows.Height = 70
        Set tt = Nothing
        
        table1.Rows(7).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Answer_head"
        tt.SetPlaceholderText Text:="Answer Text - True or Fasle"
       
        tt.Range.Font.Size = "13"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        
        table1.Rows(8).Select
        table1.Rows(8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(8, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & 1
        tt.Range.Style = "Answer_Text"
        tt.SetPlaceholderText Text:="True"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
         table1.Cell(8, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Rationale_Text"
        tt.Title = "Rationale Text " & 1
        tt.SetPlaceholderText Text:="Enter rationale for this answer"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
        table1.Rows(9).Select
        table1.Rows(9).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(9, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & 2
        tt.Range.Style = "Answer_Text"
        tt.SetPlaceholderText Text:="False"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
         table1.Cell(9, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Rationale_Text"
        tt.Title = "Rationale Text " & 2
        tt.SetPlaceholderText Text:="Enter rationale for this answer"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
        'Answer option heading
        
        
        table1.Rows(10).Select
        table1.Rows(10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 8, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_answer_head"
        tt.SetPlaceholderText Text:="Correct Answer:"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        table1.Cell(10, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Title = "Correct Answer"
        tt.LockContentControl = True
        tt.Range.Style = "Correct_Answer"
        tt.SetPlaceholderText , , "Select your correct answer: True or False"
        tt.DropdownListEntries.Add ("True")
        tt.DropdownListEntries.Add ("False")
        
        'Set tt = Nothing
        
        table1.Rows(11).Select
        table1.Rows(11).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(11, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Score_head"
        tt.SetPlaceholderText Text:="Score"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(11, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Difficulty_head"
        tt.SetPlaceholderText Text:="Difficulty"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(12).Select
        table1.Rows(12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(12, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose score in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

    table1.Cell(12, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose difficulty in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(13).Select
        table1.Rows(13).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 5, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale_head"
        tt.SetPlaceholderText Text:="Correct Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(13, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Incorrect_Answer_Rationale_head"
         tt.SetPlaceholderText Text:="Incorrect Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(14).Select
        table1.Rows(14).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(14, 1).Select
        
       
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale"
        tt.SetPlaceholderText Text:="Enter correct rationale for the question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(14, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Incorrect_Answer_Rationale"
        tt.SetPlaceholderText Text:="Enter incorrect rationale for the question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        table1.id = "TF"
        meta_rem_head table1, table1.id, QuestionNo
  
       
End Sub

Sub copytableVQ(MyValue, QuestionNo)
Dim table1 As Table
Dim tt As ContentControl
Dim TABLE2 As Table
Dim addbutton
varNumRows = MyValue + 7
'QuestionNo = QuestionNo + 1
Selection.HomeKey unit:=wdStory
Selection.EndKey unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 10, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.id = "VQ"
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        
            table1.Rows(1).Select
         table1.Rows(1).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        
        table1.Cell(1, 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Delete_Question"
        tt.Title = "Delete Question"
        
        tt.SetPlaceholderText Text:="Delete Question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdRed
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        table1.Cell(1, 4).Select
        
        Selection.Collapse Direction:=wdCollapseStart
        Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
        With ffield
                .Name = "Delete_" & QuestionNo
                .Range.Style = "Delete_box"
        End With
        
        
        
        table1.Rows(2).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Question Text - Video Questions"
        tt.Range.Style = "Question_Type"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(3, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Question_no"
        tt.Title = "Question No"
        tt.SetPlaceholderText Text:="Question No: " & QuestionNo
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(3, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Identification_Id"
        tt.Title = "Identification Id"
        tt.SetPlaceholderText Text:="Identification Id: 0"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Rows(4).Select
        table1.Rows(4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(4, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(4, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Title"
        tt.Range.Style = "Question_Title"
        tt.SetPlaceholderText Text:="Enter question title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
                
        table1.Rows(5).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        'table1.Rows(2).Select
        
        table1.Rows(6).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Text"
        tt.Range.Style = "Question_Text"
        tt.SetPlaceholderText Text:="Enter question text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 70
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        'Answer option heading
        table1.Rows(7).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Answer_head"
        tt.SetPlaceholderText Text:="Answer Text - Video Questions"
        
        tt.Range.Font.Size = "13"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        'table1.Cell(4, 2).Select
    
                    
       
        
      For i = 8 To MyValue + 7
        
       ' table1.Rows(5).SetHeight rowHeight:=ActiveDocument.PageSetup.PageHeight - 300, HeightRule:=wdRowHeightExactly
        
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 6
        tt.Range.Style = "Answer_Text"
        tt.SetPlaceholderText Text:="Enter answer text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
         table1.Cell(i, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Rationale_Text"
        tt.Title = "Rationale Text " & i - 6
        tt.SetPlaceholderText Text:="Enter rationale for this answer"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
    Next
        
         'table1.Rows(4).Cells.Split NumRows:=5, NumColumns:=2, MergeBeforeSplit:=True
         'table1.Cell(4, 2).Merge table1.Cell(8, 2)
             
        table1.Rows(MyValue + 8).Select
        table1.Rows(MyValue + 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 8, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_answer_head"
        tt.SetPlaceholderText Text:="Correct Answer:"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(MyValue + 8, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Title = "Correct Answer"
        tt.LockContentControl = True
        tt.Range.Style = "Correct_Answer"
        tt.SetPlaceholderText , , "Select your correct answer"
        For i = 1 To MyValue
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
               
        table1.Rows(MyValue + 9).Select
        table1.Rows(MyValue + 9).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 9, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Additional_Fields_head"
        tt.SetPlaceholderText Text:="Additional Fields"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(MyValue + 9, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Additional_Fields_head"
        tt.SetPlaceholderText Text:=""
        tt.LockContentControl = True
        
        table1.Cell(MyValue + 9, 3).Select
        'Selection.InlineShapes.AddOLEControl ClassType:="Forms.CommandButton.1"
        Set shp = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
        
        'shp.Range.Style = "Chosse_file"
        With shp.OLEFormat.Object
            .Object.Caption = "Select File"
            '.Object.Style = "Chosse_file"
        End With
        varbtn = Replace(shp.OLEFormat.Object.Name, "CommandButton", "")
        Set shp = Nothing
        
        
        table1.Cell(MyValue + 9, 4).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "File Name" & varbtn
        tt.Range.Style = "Additional_file_path"
        tt.Range.Font.Size = "10"
        tt.SetPlaceholderText , , ""
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        
        'Set tt = Nothing
        
        table1.Rows(MyValue + 10).Select
        table1.Rows(MyValue + 10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 10, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Score_head"
        tt.SetPlaceholderText Text:="Score"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(MyValue + 10, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Difficulty_head"
        tt.SetPlaceholderText Text:="Difficulty"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
       table1.Rows(MyValue + 11).Select
        table1.Rows(MyValue + 11).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 11, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose score in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

    table1.Cell(MyValue + 11, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose difficulty in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(MyValue + 12).Select
        table1.Rows(MyValue + 12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 12, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale_head"
        tt.SetPlaceholderText Text:="Correct Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(MyValue + 12, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Incorrect_Answer_Rationale_head"
        tt.SetPlaceholderText Text:="Incorrect Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(MyValue + 13).Select
        table1.Rows(MyValue + 13).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 13, 1).Select
        
       
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale"
        tt.SetPlaceholderText Text:="Enter correct rationale for the question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(MyValue + 13, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Incorrect_Answer_Rationale"
        tt.SetPlaceholderText Text:="Enter incorrect rationale for the question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
  
       table1.id = "VQ"
        meta_rem_head table1, table1.id, QuestionNo
        
       
        
        'UserForm.UserForm_activate
        '
        'frmLaunch.Show
        'table1.Rows(table1.Rows.Count).Select
        
        'Call GetMyPicture
        
               
       ' Set tt = Nothing
        

End Sub
Sub copytableII(MyValue, QuestionNo)
Dim table1 As Table
Dim tt As ContentControl
Dim TABLE2 As Table
Dim addbutton
varNumRows = MyValue + 7
'QuestionNo = QuestionNo + 1
Selection.HomeKey unit:=wdStory
Selection.EndKey unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 10, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.id = "MC"
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        
        
            table1.Rows(1).Select
         table1.Rows(1).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        
        table1.Cell(1, 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Delete_Question"
        tt.Title = "Delete Question"
        
        tt.SetPlaceholderText Text:="Delete Question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdRed
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        table1.Cell(1, 4).Select
        
        Selection.Collapse Direction:=wdCollapseStart
        Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
        With ffield
                .Name = "Delete_" & QuestionNo
                .Range.Style = "Delete_box"
        End With
        
        
        table1.Rows(2).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Question Text - Image Integration"
        tt.Range.Style = "Question_Type"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(3, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Question_no"
        tt.Title = "Question No"
        tt.SetPlaceholderText Text:="Question No: " & QuestionNo
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(3, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Identification_Id"
        tt.Title = "Identification Id"
        tt.SetPlaceholderText Text:="Identification Id: 0"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Rows(4).Select
        table1.Rows(4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(4, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(4, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Title"
        tt.Range.Style = "Question_Title"
        tt.SetPlaceholderText Text:="Enter question title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
                
        table1.Rows(5).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Rows(6).Select
        'table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Text"
        tt.Range.Style = "Question_Text"
        tt.SetPlaceholderText Text:="Enter question text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 70
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        'Answer option heading
        table1.Rows(7).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Answer_head"
        tt.SetPlaceholderText Text:="Answer Text - Image Integration"
        
        tt.Range.Font.Size = "13"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        'table1.Cell(4, 2).Select
    
                    
       
        
      For i = 8 To MyValue + 7
        
       ' table1.Rows(5).SetHeight rowHeight:=ActiveDocument.PageSetup.PageHeight - 300, HeightRule:=wdRowHeightExactly
        
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 6
        tt.Range.Style = "Answer_Text"
        tt.SetPlaceholderText Text:="Enter answer text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
         table1.Cell(i, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Rationale_Text"
        tt.Title = "Rationale Text " & i - 6
        tt.SetPlaceholderText Text:="Enter rationale for this answer"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
    Next
        
         'table1.Rows(4).Cells.Split NumRows:=5, NumColumns:=2, MergeBeforeSplit:=True
         'table1.Cell(4, 2).Merge table1.Cell(8, 2)
               
        table1.Rows(MyValue + 8).Select
        table1.Rows(MyValue + 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 8, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_answer_head"
        tt.SetPlaceholderText Text:="Correct Answer:"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        table1.Cell(MyValue + 8, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Title = "Correct Answer"
        tt.LockContentControl = True
        tt.Range.Style = "Correct_Answer"
        tt.SetPlaceholderText , , "Select your correct answer"
        For i = 1 To MyValue
        tt.DropdownListEntries.Add (i)
        Next
               
        table1.Rows(MyValue + 9).Select
        table1.Rows(MyValue + 9).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 9, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Additional_Fields_head"
        tt.SetPlaceholderText Text:="Additional Fields"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(MyValue + 9, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Additional_Fields_head"
        tt.SetPlaceholderText Text:=""
        tt.LockContentControl = True
        
        table1.Cell(MyValue + 9, 3).Select
        'Selection.InlineShapes.AddOLEControl ClassType:="Forms.CommandButton.1"
        Set shp = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
        
        'shp.Range.Style = "Chosse_file"
        With shp.OLEFormat.Object
            .Object.Caption = "Select File"
            '.Object.Name = "Chosse_file"
        End With
        varbtn = Replace(shp.OLEFormat.Object.Name, "CommandButton", "")
        Set shp = Nothing
        
        table1.Cell(MyValue + 9, 4).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "File Name" & varbtn
        tt.Range.Style = "Additional_file_path"
        tt.Range.Font.Size = "10"
        tt.SetPlaceholderText , , ""
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        'Set tt = Nothing
        
         table1.Rows(MyValue + 10).Select
        table1.Rows(MyValue + 10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 10, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Score_head"
        tt.SetPlaceholderText Text:="Score"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(MyValue + 10, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Difficulty_head"
        tt.SetPlaceholderText Text:="Difficulty"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(MyValue + 11).Select
        table1.Rows(MyValue + 11).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 11, 1).Select
        
        
         table1.Rows(MyValue + 11).Select
        table1.Rows(MyValue + 11).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 11, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose score in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

        table1.Cell(MyValue + 11, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose difficulty in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(MyValue + 12).Select
        table1.Rows(MyValue + 12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 12, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale_head"
        tt.SetPlaceholderText Text:="Correct Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(MyValue + 12, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Incorrect_Answer_Rationale_head"
         tt.SetPlaceholderText Text:="Incorrect Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(MyValue + 13).Select
        table1.Rows(MyValue + 13).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 13, 1).Select
        
       
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale"
        tt.SetPlaceholderText Text:="Enter correct rationale for the question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(MyValue + 13, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Incorrect_Answer_Rationale"
        tt.SetPlaceholderText Text:="Enter incorrect rationale for the question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
  
        table1.id = "II"
        meta_rem_head table1, table1.id, QuestionNo
        

End Sub
Sub copytableGO(MyValue, QuestionNo)
Dim table1 As Table
Dim tt As ContentControl
Dim TABLE2 As Table
Dim addbutton
varNumRows = MyValue + 7
'QuestionNo = QuestionNo + 1
Selection.HomeKey unit:=wdStory
Selection.EndKey unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 9, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.id = "GO"
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        
        
            table1.Rows(1).Select
         table1.Rows(1).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        
        table1.Cell(1, 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Delete_Question"
        tt.Title = "Delete Question"
        
        tt.SetPlaceholderText Text:="Delete Question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdRed
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        table1.Cell(1, 4).Select
        
        Selection.Collapse Direction:=wdCollapseStart
        Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
        With ffield
                .Name = "Delete_" & QuestionNo
                .Range.Style = "Delete_box"
        End With
        
        
        table1.Rows(2).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Question Text - Graphic Option"
        tt.Range.Style = "Question_Type"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(3, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Question_no"
        tt.Title = "Question No"
        tt.SetPlaceholderText Text:="Question No: " & QuestionNo
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(3, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Identification_Id"
        tt.Title = "Identification Id"
        tt.SetPlaceholderText Text:="Identification Id: 0"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Rows(4).Select
        table1.Rows(4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(4, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(4, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Title"
        tt.Range.Style = "Question_Title"
        tt.SetPlaceholderText Text:="Enter question title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
                
        table1.Rows(5).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        'table1.Rows(2).Select
         table1.Rows(6).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Text"
        tt.Range.Style = "Question_Text"
        tt.SetPlaceholderText Text:="Enter question text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 70
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        'Answer option heading
        table1.Rows(7).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Answer_head"
        tt.SetPlaceholderText Text:="Answer Text - Graphic Option"
        
        tt.Range.Font.Size = "13"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        'table1.Cell(4, 2).Select
    
                    
       
        
      For i = 8 To MyValue + 7
        
       ' table1.Rows(5).SetHeight rowHeight:=ActiveDocument.PageSetup.PageHeight - 300, HeightRule:=wdRowHeightExactly
        
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=3, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        'Selection.InlineShapes.AddOLEControl ClassType:="Forms.CommandButton.1"
        Set shp = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
        
        'shp.Range.Style = "Chosse_file"
        With shp.OLEFormat.Object
            .Object.Caption = "Select File"
            '.Object.Style = "Chosse_file"
        End With
        varbtn = Replace(shp.OLEFormat.Object.Name, "CommandButton", "")
        Set shp = Nothing
        'table1.Cell(i, 1).Width = 20
        
        
        table1.Cell(i, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "File Name" & varbtn
        tt.Range.Style = "Additional_file_path"
        tt.Range.Font.Size = "10"
        tt.SetPlaceholderText , , ""
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(i, 3).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Rationale_Text"
        tt.Title = "Rationale Text " & i - 6
        tt.SetPlaceholderText Text:="Enter rationale for this answer"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
    Next
        
         'table1.Rows(4).Cells.Split NumRows:=5, NumColumns:=2, MergeBeforeSplit:=True
         'table1.Cell(4, 2).Merge table1.Cell(8, 2)
             
        table1.Rows(MyValue + 8).Select
        table1.Rows(MyValue + 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 8, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_answer_head"
        tt.SetPlaceholderText Text:="Correct Answer:"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(MyValue + 8, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Title = "Correct Answer"
        tt.LockContentControl = True
        tt.Range.Style = "Correct_Answer"
        tt.SetPlaceholderText , , "Select your correct answer"
        For i = 1 To MyValue
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
               
        
        
        
        
        'Set tt = Nothing
        
        table1.Rows(MyValue + 9).Select
        table1.Rows(MyValue + 9).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 9, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Score_head"
        tt.SetPlaceholderText Text:="Score"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(MyValue + 9, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Difficulty_head"
        tt.SetPlaceholderText Text:="Difficulty"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(MyValue + 10).Select
        table1.Rows(MyValue + 10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 10, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose score in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

    table1.Cell(MyValue + 10, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose difficulty in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(MyValue + 11).Select
        table1.Rows(MyValue + 11).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 11, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale_head"
        tt.SetPlaceholderText Text:="Correct Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(MyValue + 11, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Incorrect_Answer_Rationale_head"
        tt.SetPlaceholderText Text:="Incorrect Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(MyValue + 12).Select
        table1.Rows(MyValue + 12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(MyValue + 12, 1).Select
        
       
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale"
        tt.SetPlaceholderText Text:="Enter correct rationale for the question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(MyValue + 12, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Incorrect_Answer_Rationale"
        tt.SetPlaceholderText Text:="Enter incorrect rationale for the question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
       table1.id = "GO"
        meta_rem_head table1, table1.id, QuestionNo
           
       
        
       
        

End Sub
Public Function CreateManifest()
If FSO.FileExists(ActiveDocument.Path & "\manifest.xml") = True Then FSO.DeleteFile (ActiveDocument.Path & "\manifest.xml")
Open ActiveDocument.Path & "\manifest.xml" For Output As #55
Print #55, "<?xml version=""1.0"" encoding=""utf-8"" ?>"
Print #55, "<publisher-upload-manifest>"
    Set flo = FSO.GetFolder(ActiveDocument.Path)
    For Each fil In flo.Files
        If fil.Name = "unzip.exe" Or fil.Name = "zip.exe" Or fil.Name = "test.bat" Or fil.Name = "ascii.txt" Or fil.Name = "manifest.xml" Or InStr(fil.Name, "~$") Then
        Else
            Print #55, "<asset filename=""" & fil.Name & """ filepath = """ & fil.Path & """ type=""" & fil.Type & """ />"
        End If
    Next
    On Error Resume Next
    Set flo1 = FSO.GetFolder(ActiveDocument.Path & "\media\")
    For Each fil In flo1.Files
        Print #55, "<asset filename=""" & fil.Name & """ filepath = """ & fil.Path & """ type=""" & fil.Type & """ />"
    Next
Print #55, "</publisher-upload-manifest>"
Close #55
End Function

Public Sub zip()
        Dim ws As New WshShell
        varDocPath = ActiveDocument.Path
        'If fso.FolderExists(varDocPath & "\Question_bank") = False Then fso.CreateFolder (varDocPath & "\Question_bank")
        'ActiveDocument.Close (wdDoNotSaveChanges)
        'If fso.FileExists(varDocPath & "\Question_bank.docm") = True Then fso.CopyFile varDocPath & "\Question_bank.docm", varDocPath & "\Question_bank"
        'If fso.FileExists(varDocPath & "\manifest.xml") = True Then fso.CopyFile varDocPath & "\manifest.xml", varDocPath & "\Question_bank"
        'If fso.FileExists(varDocPath & "\XML_Error_Log.txt") = True Then fso.CopyFile varDocPath & "\XML_Error_Log.txt", varDocPath & "\Question_bank"
        'If fso.FileExists(varDocPath & "\allQuestions.xml") = True Then fso.CopyFile varDocPath & "\allQuestions.xml", varDocPath & "\Question_bank"
        'If fso.FolderExists(varDocPath & "\media") = True Then fso.CopyFile varDocPath & "\media", varDocPath & "\Question_bank"
        Open varDocPath & "\test.bat" For Output As #111
        Set tempflo = FSO.GetFolder(varDocPath)
        Print #111, tempflo.Drive
        Print #111, "cd " & Chr(34) & tempflo.Path & Chr(34)
        'Print #111, "Copy" & " " & Chr(34) & App.Path & "\zip.exe" & Chr(34) & "," & Chr(34) & flo2.Path & Chr(34)
        'zip.exe
        'Print #111, "c:"
        Print #111, tempflo.Drive
        Print #111, "cd " & Chr(34) & tempflo.Path & Chr(34)
        Print #111, Chr(34) & varDocPath & "\zip.exe" & Chr(34) & " -Xr9D" & " " & Chr(34) & varDocPath & "\Question_bank.zip" & Chr(34) & " " & "*.*" & " " & "-x" & " " & Chr(34) & "zip.exe" & Chr(34) & " " & "-x" & " " & Chr(34) & "test.bat" & Chr(34) & " " & "-x" & " " & Chr(34) & "ascii.txt" & Chr(34) & " " & "-x" & " " & Chr(34) & "unzip.exe" & Chr(34)
      '  Print #111, "del zip.exe"
        Close #111
        ws.Run Chr(34) & varDocPath & "\test.bat" & Chr(34), 0, True
        MsgBox "XML created and zipped in the same path."
        
        'Set fso = CreateObject("Scripting.FileSystemObject")
        'fso.DeleteFile fso.BuildPath(dst, fName), True
        
End Sub
Public Function Zipper()
'Zips A File
'ZipName must be FULL Path\Filename.zip - name Zip File to Create OR ADD To
'FileToZip must be Full Path\Filename.xls  -  Name of file you want to zip
If FSO.FileExists(ActiveDocument.Path & "\Question_bank.zip") Then FSO.DeleteFile (ActiveDocument.Path & "\Question_bank.zip")
ZipName = ActiveDocument.Path & "\Question_bank.zip"
FileToZip = ActiveDocument.Path & "\Question_bank.docm"
FileToZip1 = ActiveDocument.Path & "\media"
FileToZip2 = ActiveDocument.Path & "\allQuestions.xml"
FileToZip3 = ActiveDocument.Path & "\XML_Error_Log.txt"

'''''Create manifest

FileToZip4 = ActiveDocument.Path & "\manifest.xml"

'Dim fso As Object
Dim oApp As Object
If Dir(ZipName) = "" Then
    Open ZipName For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End If
Set oApp = CreateObject("Shell.Application")
oApp.Namespace(ZipName).CopyHere (FileToZip)
oApp.Namespace(ZipName).CopyHere (FileToZip1)
oApp.Namespace(ZipName).CopyHere (FileToZip2)
oApp.Namespace(ZipName).CopyHere (FileToZip3)
oApp.Namespace(ZipName).CopyHere (FileToZip4)

On Error Resume Next
Set FSO = CreateObject("scripting.filesystemobject")
FSO.DeleteFolder Environ("Temp") & "\Temporary Directory*", True
Set oApp = Nothing
Set FSO = Nothing
MsgBox "XML created and Zipped in the same path."
End Function
Sub copytableupdatemc(vartitle, MyValue, varIdentificationId, varqtext, varanswer, varrationale, varcorrecttxt, varscoretxt, vardifficultytxt, varcorrectrationaletxt, varincorrectrationaletxt, varWeblinkRemediationTxt, varWeblinkRemediationLink, varWeblinkRemediationTooltip, varTextRemediationTxt, vartagtype, vartagvalue)
Dim table1 As Table
Dim tt As ContentControl
Dim TABLE2 As Table
Dim addbutton
varNumRows = MyValue + 6
UpdateQuestionNo = UpdateQuestionNo + 1
Selection.HomeKey unit:=wdStory
Selection.EndKey unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 17, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.id = "MC"
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(1).Select
         table1.Rows(1).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        
        table1.Cell(1, 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Delete_Question"
        tt.Title = "Delete Question"
        
        tt.SetPlaceholderText Text:="Delete Question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdRed
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        table1.Cell(1, 4).Select
        
        Selection.Collapse Direction:=wdCollapseStart
        Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
        With ffield
                .Name = "Delete_" & QuestionNo
                .Range.Style = "Delete_box"
        End With
        
        table1.Rows(2).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Question Text - Multiple Choice"
        tt.Range.Style = "Question_Type"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
       
        table1.Cell(3, 1).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Question_no"
        tt.Title = "Question No"
        tt.SetPlaceholderText Text:="Question No: " & UpdateQuestionNo
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(3, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Identification_Id"
        tt.Title = "Identification Id"
        tt.SetPlaceholderText Text:="Identification Id: " & varIdentificationId
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        
        table1.Rows(4).Select
        table1.Rows(4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(4, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(4, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Title"
        tt.Range.Style = "Question_Title"
        tt.SetPlaceholderText Text:=vartitle
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
                
        table1.Rows(5).Select
        
        'table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Text"
        tt.Range.Style = "Question_Text"
        'tt.ShowingPlaceholderText = varqtext
        tt.SetPlaceholderText Text:=varqtext
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 70
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        'Answer option heading
        table1.Rows(6).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Answer_head"
        tt.SetPlaceholderText Text:="Answer Text - Multiple Choice"
        
        tt.Range.Font.Size = "13"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
    
         
        
      For i = 7 To MyValue + 6
        
       ' table1.Rows(5).SetHeight rowHeight:=ActiveDocument.PageSetup.PageHeight - 300, HeightRule:=wdRowHeightExactly
        
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 6
        tt.Range.Style = "Answer_Text"
        tt.SetPlaceholderText Text:=varanswer(i - 7)
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
         table1.Cell(i, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Rationale_Text"
        tt.Title = "Rationale Text " & i - 6
        tt.SetPlaceholderText Text:=varrationale(i - 7)
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
    Next
        
         'table1.Rows(4).Cells.Split NumRows:=5, NumColumns:=2, MergeBeforeSplit:=True
         'table1.Cell(4, 2).Merge table1.Cell(8, 2)
               
      
               
        table1.Rows(table1.Rows.Count - 8).Select
        table1.Rows(table1.Rows.Count - 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 8, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_answer_head"
        tt.SetPlaceholderText Text:="Correct Answer:"
        
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 8, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Title = "Correct Answer"
        tt.LockContentControl = True
        tt.Range.Style = "Correct_Answer"
        tt.SetPlaceholderText , , varcorrecttxt
        For i = 1 To MyValue
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        'Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 7).Select
        table1.Rows(table1.Rows.Count - 7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 7, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Score_head"
        tt.SetPlaceholderText Text:="Score"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 7, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Difficulty_head"
        tt.SetPlaceholderText Text:="Difficulty"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 6).Select
        table1.Rows(table1.Rows.Count - 6).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 6, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , varscoretxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

        table1.Cell(table1.Rows.Count - 6, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , vardifficultytxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 5).Select
        table1.Rows(table1.Rows.Count - 5).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 5, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale_head"
        tt.SetPlaceholderText Text:="Correct Rationale"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 5, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Incorrect_Answer_Rationale_head"
         tt.SetPlaceholderText Text:="Incorrect Rationale"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 4).Select
        table1.Rows(table1.Rows.Count - 4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 4, 1).Select
        
       
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale"
        tt.SetPlaceholderText Text:=varcorrectrationaletxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 4, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.SetPlaceholderText Text:=varincorrectrationaletxt
        tt.Range.Style = "Incorrect_Answer_Rationale"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
  
        table1.Rows(table1.Rows.Count - 3).Shading.BackgroundPatternColor = RGB(255, 140, 0)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Links_head"
        tt.SetPlaceholderText Text:="Remediation"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 2).Select
        table1.Rows(table1.Rows.Count - 2).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 2, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Link_Type_head"
        tt.SetPlaceholderText Text:="Remediation Type"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 2, 2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Remediation_head"
        tt.SetPlaceholderText Text:="Remediation Detail"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 1).Select
        table1.Rows(table1.Rows.Count - 1).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Cell(table1.Rows.Count - 1, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
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
        
        table1.Cell(table1.Rows.Count - 1, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Text"
        tt.Title = "Remediation Text"
        tt.SetPlaceholderText Text:=varWeblinkRemediationTxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 8).Select
        table1.Rows(table1.Rows.Count - 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
               
        table1.Cell(table1.Rows.Count - 8, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Link"
        tt.Title = "Remediation Link"
        tt.SetPlaceholderText Text:=varWeblinkRemediationLink
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 7).Select
        table1.Rows(table1.Rows.Count - 7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
                
        table1.Cell(table1.Rows.Count - 7, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Tooltip"
        tt.Title = "Remediation Tooltip"
        tt.SetPlaceholderText Text:=varWeblinkRemediationTooltip
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        
      table1.Rows(table1.Rows.Count - 6).Select
        'table1.Rows(table1.Rows.Count - 6).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Rows(table1.Rows.Count - 6).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Link_Type"
        tt.SetPlaceholderText Text:="Ebook"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 5).Select
        table1.Rows(table1.Rows.Count - 5).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        table1.Cell(table1.Rows.Count - 5, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Link_Type"
        tt.SetPlaceholderText Text:="Text"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count - 5, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Text_Remediation_Text"
        tt.Title = "Remediation Text"
        tt.SetPlaceholderText Text:=varTextRemediationTxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        '''''''''Metadata Attributes'''''''''
        
        table1.Rows(table1.Rows.Count - 4).Shading.BackgroundPatternColor = RGB(169, 169, 169)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 4).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Metadata Attributes"
        tt.Range.Style = "Meta_Data_Attributes_head"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 3).Select
        table1.Rows(table1.Rows.Count - 3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 3, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="METADATA TAG"
        tt.Range.Style = "METADATA_TAG_head"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 3, 2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.SetPlaceholderText Text:="METADATA VALUE"
        tt.Range.Style = "METADATA_VALUE_head"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        
        table1.Rows(table1.Rows.Count - 2).Select
        table1.Rows(table1.Rows.Count - 2).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 2, 1).Select
       
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.LockContentControl = True
        tt.Range.Style = "METADATA_TAG"
        tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagtype(0)
        tt.DropdownListEntries.Add "Status"
        tt.DropdownListEntries.Add "Keyword"
        tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count - 2, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagvalue(0)
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 1).Select
        table1.Rows(table1.Rows.Count - 1).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 1, 1).Select
       
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.LockContentControl = True
        tt.Range.Style = "METADATA_TAG"
        tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagtype(1)
        tt.DropdownListEntries.Add "Status"
        tt.DropdownListEntries.Add "Keyword"
        tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count - 1, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
         tt.SetPlaceholderText , , vartagvalue(1)
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        '''''Create Table Row''''
        table1.Rows(table1.Rows.Count).Select
        table1.Rows(table1.Rows.Count).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count, 1).Select
       
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.LockContentControl = True
        tt.Range.Style = "METADATA_TAG"
        tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagtype(2)
        tt.DropdownListEntries.Add "Status"
        tt.DropdownListEntries.Add "Keyword"
        tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
         tt.SetPlaceholderText , , vartagvalue(2)
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing

End Sub

Sub copytableupdatecm(vartitle, MyValue, varIdentificationId, varqtext, varanswer, varrationale, varcorrecttxt, varscoretxt, vardifficultytxt, varcorrectrationaletxt, varincorrectrationaletxt, varWeblinkRemediationTxt, varWeblinkRemediationLink, varWeblinkRemediationTooltip, varTextRemediationTxt, vartagtype, vartagvalue)
Dim table1 As Table
Dim tt As ContentControl
Dim TABLE2 As Table
Dim addbutton
varNumRows = MyValue + 6
UpdateQuestionNo = UpdateQuestionNo + 1
Selection.HomeKey unit:=wdStory
Selection.EndKey unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 17, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        
        table1.Rows(1).Select
         table1.Rows(1).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        
        table1.Cell(1, 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Delete_Question"
        tt.Title = "Delete Question"
        
        tt.SetPlaceholderText Text:="Delete Question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdRed
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        table1.Cell(1, 4).Select
        
        Selection.Collapse Direction:=wdCollapseStart
        Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
        With ffield
                .Name = "Delete_" & QuestionNo
                .Range.Style = "Delete_box"
        End With
        
        
        
        table1.Rows(2).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Question Text - Choice Multiple"
        tt.Range.Style = "Question_Type"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(3, 1).Select
        
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Question_no"
        tt.Title = "Question No"
        
        tt.SetPlaceholderText Text:="Question No: " & UpdateQuestionNo
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(3, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Identification_Id"
        tt.Title = "Identification Id"
        tt.SetPlaceholderText Text:="Identification Id: " & varIdentificationId
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Rows(4).Select
        table1.Rows(4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(4, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(4, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Title"
        tt.Range.Style = "Question_Title"
        tt.SetPlaceholderText Text:=vartitle
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
        table1.Rows(5).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Text"
        tt.Range.Style = "Question_Text"
        tt.SetPlaceholderText Text:=varqtext
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        tt.Range.Rows.Height = 70
        Set tt = Nothing
        
        'Answer option heading
        table1.Rows(6).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Answer_head"
        tt.SetPlaceholderText Text:="Answer Text - Choice Multiple"
        
        tt.Range.Font.Size = "13"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
      
        
       ' table1.Rows(5).SetHeight rowHeight:=ActiveDocument.PageSetup.PageHeight - 300, HeightRule:=wdRowHeightExactly
        For i = 7 To MyValue + 6
            
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 6
        tt.Range.Style = "Answer_Text"
        tt.SetPlaceholderText Text:=varanswer(i - 7)
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
         table1.Cell(i, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Rationale_Text"
        tt.Title = "Rationale Text " & i - 6
        tt.SetPlaceholderText Text:=varrationale(i - 7)
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
         
        Next
               
        table1.Rows(table1.Rows.Count - 8).Select
        table1.Rows(table1.Rows.Count - 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 8, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Correct Answers:"
        tt.Range.Style = "Correct_Answer"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        Dim strCorrect As String
        Dim strArray() As String
        Dim intCount As Integer
   
        strCorrect = varcorrecttxt
        strArray = Split(strCorrect, ",")
        
        For i = 2 To (MyValue * 2) + 1
        
        If i = 2 Then
            j = 1
            K = 1
            table1.Cell(table1.Rows.Count - 8, i).Split NumColumns:=(MyValue * 2) + 1
            table1.Cell(table1.Rows.Count - 8, 2).Select
            Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
            tt.SetPlaceholderText Text:=" "
            tt.Range.Style = "Check_box_no"
            tt.Range.Font.Size = "10"
            tt.Range.Font.Name = "Verdana"
            tt.LockContentControl = True
            tt.LockContents = False
            
        End If
        
        If i Mod 2 > 0 Then
            table1.Cell(table1.Rows.Count - 8, i).Select
            
            Selection.Collapse Direction:=wdCollapseEnd
            Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
            With ffield
                .Name = "Check_" & UpdateQuestionNo & "_" & K
                .Range.Style = "Check_box"
                '.CheckBox.Value = False
                For intCount = LBound(strArray) To UBound(strArray)
                    If strArray(intCount) = K Then
                        .CheckBox.Value = True
                    End If
                Next
            End With
            K = K + 1
        Else
            table1.Cell(table1.Rows.Count - 8, i).Select
            Selection.Collapse Direction:=wdCollapseEnd
            Set tt = ActiveDocument.ContentControls.Add(wdContentControlText)
                
            tt.SetPlaceholderText Text:=j
            tt.Range.Style = "Check_box_no"
            Set tt = Nothing
            j = j + 1
        End If
        
        Next
        
        
        'Set tt = Nothing
        table1.Rows(table1.Rows.Count - 7).Select
        table1.Rows(table1.Rows.Count - 7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 7, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Score_head"
        tt.SetPlaceholderText Text:="Score"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 7, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Difficulty_head"
        tt.SetPlaceholderText Text:="Difficulty"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 6).Select
        table1.Rows(table1.Rows.Count - 6).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 6, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , varscoretxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

    table1.Cell(table1.Rows.Count - 6, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , vardifficultytxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 5).Select
        table1.Rows(table1.Rows.Count - 5).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 5, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale_head"
        tt.SetPlaceholderText Text:="Correct Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 5, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Incorrect_Answer_Rationale_head"
         tt.SetPlaceholderText Text:="Incorrect Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 4).Select
        table1.Rows(table1.Rows.Count - 4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 4, 1).Select
        
       
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale"
        tt.SetPlaceholderText Text:=varcorrectrationaletxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 4, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Incorrect_Answer_Rationale"
        tt.SetPlaceholderText Text:=varincorrectrationaletxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
  
        table1.Rows(table1.Rows.Count - 3).Shading.BackgroundPatternColor = RGB(255, 140, 0)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Links_head"
        tt.SetPlaceholderText Text:="Remediation"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 2).Select
        table1.Rows(table1.Rows.Count - 2).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 2, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Link_Type_head"
        tt.SetPlaceholderText Text:="Remediation Type"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 2, 2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Remediation_head"
        tt.SetPlaceholderText Text:="Remediation Detail"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        
        
        table1.Rows(table1.Rows.Count - 1).Select
        table1.Rows(table1.Rows.Count - 1).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Cell(table1.Rows.Count - 1, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
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
        
        table1.Cell(table1.Rows.Count - 1, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Text"
        tt.Title = "Remediation Text"
        tt.SetPlaceholderText Text:=varWeblinkRemediationTxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 8).Select
        table1.Rows(table1.Rows.Count - 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
               
        table1.Cell(table1.Rows.Count - 8, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Link"
        tt.Title = "Remediation Link"
        tt.SetPlaceholderText Text:=varWeblinkRemediationLink
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 7).Select
        table1.Rows(table1.Rows.Count - 7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
                
        table1.Cell(table1.Rows.Count - 7, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Tooltip"
        tt.Title = "Remediation Tooltip"
        tt.SetPlaceholderText Text:=varWeblinkRemediationTooltip
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        
       table1.Rows(table1.Rows.Count - 6).Select
        'table1.Rows(table1.Rows.Count - 6).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Rows(table1.Rows.Count - 6).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Link_Type"
        tt.SetPlaceholderText Text:="Ebook"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 5).Select
        table1.Rows(table1.Rows.Count - 5).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        table1.Cell(table1.Rows.Count - 5, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Link_Type"
        tt.SetPlaceholderText Text:="Text"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count - 5, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Text_Remediation_Text"
        tt.Title = "Remediation Text"
        tt.SetPlaceholderText Text:=varTextRemediationTxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        '''''''''Metadata Attributes'''''''''
        
        table1.Rows(table1.Rows.Count - 4).Shading.BackgroundPatternColor = RGB(169, 169, 169)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 4).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Metadata Attributes"
        tt.Range.Style = "Meta_Data_Attributes_head"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 3).Select
        table1.Rows(table1.Rows.Count - 3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 3, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="METADATA TAG"
        tt.Range.Style = "METADATA_TAG_head"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 3, 2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.SetPlaceholderText Text:="METADATA VALUE"
        tt.Range.Style = "METADATA_VALUE_head"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 2).Select
        table1.Rows(table1.Rows.Count - 2).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 2, 1).Select
       
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.LockContentControl = True
        tt.Range.Style = "METADATA_TAG"
        tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagtype(0)
        tt.DropdownListEntries.Add "Status"
        tt.DropdownListEntries.Add "Keyword"
        tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count - 2, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagvalue(0)
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 1).Select
        table1.Rows(table1.Rows.Count - 1).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 1, 1).Select
       
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.LockContentControl = True
        tt.Range.Style = "METADATA_TAG"
        tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagtype(1)
        tt.DropdownListEntries.Add "Status"
        tt.DropdownListEntries.Add "Keyword"
        tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count - 1, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagvalue(1)
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        '''''Create Table Row''''
        table1.Rows(table1.Rows.Count).Select
        table1.Rows(table1.Rows.Count).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count, 1).Select
       
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.LockContentControl = True
        tt.Range.Style = "METADATA_TAG"
        tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagtype(2)
        tt.DropdownListEntries.Add "Status"
        tt.DropdownListEntries.Add "Keyword"
        tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagvalue(2)
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing

End Sub

Sub copytableupdatetf(vartitle, MyValue, varIdentificationId, varqtext, varanswer, varrationale, varcorrecttxt, varscoretxt, vardifficultytxt, varcorrectrationaletxt, varincorrectrationaletxt, varWeblinkRemediationTxt, varWeblinkRemediationLink, varWeblinkRemediationTooltip, varTextRemediationTxt, vartagtype, vartagvalue)
Dim table1 As Table
Dim tt As ContentControl
Dim TABLE2 As Table
Dim addbutton
UpdateQuestionNo = UpdateQuestionNo + 1
Selection.HomeKey unit:=wdStory
Selection.EndKey unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=8 + 17, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        
        table1.Rows(1).Select
         table1.Rows(1).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        table1.Cell(1, 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Delete_Question"
        tt.Title = "Delete Question"
        
        tt.SetPlaceholderText Text:="Delete Question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdRed
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        table1.Cell(1, 4).Select
        
        Selection.Collapse Direction:=wdCollapseStart
        Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
        With ffield
                .Name = "Delete_" & QuestionNo
                .Range.Style = "Delete_box"
        End With
        
        table1.Rows(2).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Question Text - True or Fasle"
        tt.Range.Style = "Question_Type"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(3, 1).Select
               
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Question_no"
        tt.Title = "Question No"
        
        tt.SetPlaceholderText Text:="Question No: " & UpdateQuestionNo
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(3, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Identification_Id"
        tt.Title = "Identification Id"
        tt.SetPlaceholderText Text:="Identification Id: " & varIdentificationId
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Rows(4).Select
        table1.Rows(4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(4, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(4, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Title"
        tt.Range.Style = "Question_Title"
        tt.SetPlaceholderText Text:=vartitle
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
        table1.Rows(5).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Text"
        tt.Range.Style = "Question_Text"
        tt.SetPlaceholderText Text:=varqtext
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        tt.Range.Rows.Height = 70
        Set tt = Nothing
        
        table1.Rows(6).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Answer_head"
        tt.SetPlaceholderText Text:="Answer Text - True or Fasle"
       
        tt.Range.Font.Size = "13"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        
        table1.Rows(7).Select
        table1.Rows(7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(7, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & 1
        tt.Range.Style = "Answer_Text"
        tt.SetPlaceholderText Text:="True"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
         table1.Cell(7, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Rationale_Text"
        tt.Title = "Rationale Text " & 1
        tt.SetPlaceholderText Text:=varrationale(0)
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
        table1.Rows(8).Select
        table1.Rows(8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(8, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & 2
        tt.Range.Style = "Answer_Text"
        tt.SetPlaceholderText Text:="False"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
         table1.Cell(8, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Rationale_Text"
        tt.Title = "Rationale Text " & 2
        tt.SetPlaceholderText Text:=varrationale(1)
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
        'Answer option heading
        
        
        table1.Rows(table1.Rows.Count - 8).Select
        table1.Rows(table1.Rows.Count - 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 8, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_answer_head"
        tt.SetPlaceholderText Text:="Correct Answer:"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        table1.Cell(table1.Rows.Count - 8, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Title = "Correct Answer"
        tt.LockContentControl = True
        tt.Range.Style = "Correct_Answer"
        tt.SetPlaceholderText , , varcorrecttxt
        tt.DropdownListEntries.Add ("True")
        tt.DropdownListEntries.Add ("False")
        
        'Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 7).Select
        table1.Rows(table1.Rows.Count - 7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 7, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Score_head"
        tt.SetPlaceholderText Text:="Score"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 7, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Difficulty_head"
        tt.SetPlaceholderText Text:="Difficulty"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 6).Select
        table1.Rows(table1.Rows.Count - 6).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 6, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , varscoretxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

    table1.Cell(table1.Rows.Count - 6, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , vardifficultytxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 5).Select
        table1.Rows(table1.Rows.Count - 5).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 5, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale_head"
        tt.SetPlaceholderText Text:="Correct Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 5, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Incorrect_Answer_Rationale_head"
         tt.SetPlaceholderText Text:="Incorrect Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 4).Select
        table1.Rows(table1.Rows.Count - 4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 4, 1).Select
        
       
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale"
        tt.SetPlaceholderText Text:=varcorrectrationaletxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 4, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Incorrect_Answer_Rationale"
        tt.SetPlaceholderText Text:=varincorrectrationaletxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
  
        table1.Rows(table1.Rows.Count - 3).Shading.BackgroundPatternColor = RGB(255, 140, 0)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Links_head"
        tt.SetPlaceholderText Text:="Remediation"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 2).Select
        table1.Rows(table1.Rows.Count - 2).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 2, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Link_Type_head"
        tt.SetPlaceholderText Text:="Remediation Type"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 2, 2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Remediation_head"
        tt.SetPlaceholderText Text:="Remediation Detail"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        
        table1.Rows(table1.Rows.Count - 1).Select
        table1.Rows(table1.Rows.Count - 1).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Cell(table1.Rows.Count - 1, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
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
        
        table1.Cell(table1.Rows.Count - 1, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Text"
        tt.Title = "Remediation Text"
        tt.SetPlaceholderText Text:=varWeblinkRemediationTxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 8).Select
        table1.Rows(table1.Rows.Count - 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
               
        table1.Cell(table1.Rows.Count - 8, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Link"
        tt.Title = "Remediation Link"
        tt.SetPlaceholderText Text:=varWeblinkRemediationLink
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 7).Select
        table1.Rows(table1.Rows.Count - 7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
                
        table1.Cell(table1.Rows.Count - 7, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Tooltip"
        tt.Title = "Remediation Tooltip"
        tt.SetPlaceholderText Text:=varWeblinkRemediationTooltip
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        
       table1.Rows(table1.Rows.Count - 6).Select
        'table1.Rows(table1.Rows.Count - 6).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Rows(table1.Rows.Count - 6).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Link_Type"
        tt.SetPlaceholderText Text:="Ebook"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 5).Select
        table1.Rows(table1.Rows.Count - 5).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        table1.Cell(table1.Rows.Count - 5, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Link_Type"
        tt.SetPlaceholderText Text:="Text"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count - 5, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Text_Remediation_Text"
        tt.Title = "Remediation Text"
        tt.SetPlaceholderText Text:=varTextRemediationTxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        '''''''''Metadata Attributes'''''''''
        
        table1.Rows(table1.Rows.Count - 4).Shading.BackgroundPatternColor = RGB(169, 169, 169)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 4).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Metadata Attributes"
        tt.Range.Style = "Meta_Data_Attributes_head"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 3).Select
        table1.Rows(table1.Rows.Count - 3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 3, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="METADATA TAG"
        tt.Range.Style = "METADATA_TAG_head"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 3, 2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.SetPlaceholderText Text:="METADATA VALUE"
        tt.Range.Style = "METADATA_VALUE_head"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        
        table1.Rows(table1.Rows.Count - 2).Select
        table1.Rows(table1.Rows.Count - 2).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 2, 1).Select
       
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.LockContentControl = True
        tt.Range.Style = "METADATA_TAG"
        tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagtype(0)
        tt.DropdownListEntries.Add "Status"
        tt.DropdownListEntries.Add "Keyword"
        tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count - 2, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagvalue(0)
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 1).Select
        table1.Rows(table1.Rows.Count - 1).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 1, 1).Select
       
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.LockContentControl = True
        tt.Range.Style = "METADATA_TAG"
        tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagtype(1)
        tt.DropdownListEntries.Add "Status"
        tt.DropdownListEntries.Add "Keyword"
        tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count - 1, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagvalue(1)
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        '''''Create Table Row''''
        table1.Rows(table1.Rows.Count).Select
        table1.Rows(table1.Rows.Count).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count, 1).Select
       
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.LockContentControl = True
        tt.Range.Style = "METADATA_TAG"
        tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagtype(2)
        tt.DropdownListEntries.Add "Status"
        tt.DropdownListEntries.Add "Keyword"
        tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagvalue(2)
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
End Sub

Sub copytableupdatevq(vartitle, MyValue, varIdentificationId, varqtext, varanswer, varrationale, varcorrecttxt, varscoretxt, vardifficultytxt, varcorrectrationaletxt, varincorrectrationaletxt, varWeblinkRemediationTxt, varWeblinkRemediationLink, varWeblinkRemediationTooltip, varTextRemediationTxt, vartagtype, vartagvalue, varadditionalfile)
Dim table1 As Table
Dim tt As ContentControl
Dim TABLE2 As Table
Dim addbutton
varNumRows = MyValue + 7
UpdateQuestionNo = UpdateQuestionNo + 1
Selection.HomeKey unit:=wdStory
Selection.EndKey unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 17, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.id = "VQ"
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        
        table1.Rows(1).Select
         table1.Rows(1).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
       
        table1.Cell(1, 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Delete_Question"
        tt.Title = "Delete Question"
       
        tt.SetPlaceholderText Text:="Delete Question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdRed
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
       
        table1.Cell(1, 4).Select
       
        Selection.Collapse Direction:=wdCollapseStart
        Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
        With ffield
                .Name = "Delete_" & QuestionNo
                .Range.Style = "Delete_box"
        End With
        
        table1.Rows(2).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Question Text - Video Questions"
        tt.Range.Style = "Question_Type"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
       
        table1.Cell(3, 1).Select
        
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Question_no"
        tt.Title = "Question No"
        
        tt.SetPlaceholderText Text:="Question No: " & UpdateQuestionNo
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(3, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Identification_Id"
        tt.Title = "Identification Id"
        tt.SetPlaceholderText Text:="Identification Id: " & varIdentificationId
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
       
        Set tt = Nothing
        
        
        table1.Rows(4).Select
        table1.Rows(4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(4, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(4, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Title"
        tt.Range.Style = "Question_Title"
        tt.SetPlaceholderText Text:=vartitle
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
                
        table1.Rows(5).Select
        
        'table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Text"
        tt.Range.Style = "Question_Text"
        tt.SetPlaceholderText Text:=varqtext
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 70
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        'Answer option heading
        table1.Rows(6).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Answer_head"
        tt.SetPlaceholderText Text:="Answer Text - Video Questions"
        
        tt.Range.Font.Size = "13"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        'table1.Cell(4, 2).Select
    
                    
       
        
      For i = 7 To MyValue + 6
        
       ' table1.Rows(5).SetHeight rowHeight:=ActiveDocument.PageSetup.PageHeight - 300, HeightRule:=wdRowHeightExactly
        
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 6
        tt.Range.Style = "Answer_Text"
        tt.SetPlaceholderText Text:=varanswer(i - 7)
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
         table1.Cell(i, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Rationale_Text"
        tt.Title = "Rationale Text " & i - 6
        tt.SetPlaceholderText Text:=varrationale(i - 7)
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
    Next
        
         'table1.Rows(4).Cells.Split NumRows:=5, NumColumns:=2, MergeBeforeSplit:=True
         'table1.Cell(4, 2).Merge table1.Cell(8, 2)
             
        table1.Rows(table1.Rows.Count - 17).Select
        table1.Rows(table1.Rows.Count - 17).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 17, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_answer_head"
        tt.SetPlaceholderText Text:="Correct Answer:"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 17, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Title = "Correct Answer"
        tt.LockContentControl = True
        tt.Range.Style = "Correct_Answer"
        tt.SetPlaceholderText , , varcorrecttxt
        For i = 1 To MyValue
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
               
        table1.Rows(table1.Rows.Count - 8).Select
        table1.Rows(table1.Rows.Count - 8).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 8, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Additional_Fields_head"
        tt.SetPlaceholderText Text:="Additional Fields"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 8, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Additional_Fields_head"
        tt.SetPlaceholderText Text:=""
        tt.LockContentControl = True
        
        table1.Cell(table1.Rows.Count - 8, 3).Select
        'Selection.InlineShapes.AddOLEControl ClassType:="Forms.CommandButton.1"
        Set shp = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
        
        'shp.Range.Style = "Chosse_file"
        With shp.OLEFormat.Object
            .Object.Caption = "Select File"
            '.Object.Style = "Chosse_file"
        End With
        varbtn = Replace(shp.OLEFormat.Object.Name, "CommandButton", "")
        Set shp = Nothing
        
        
        table1.Cell(table1.Rows.Count - 8, 4).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "File Name" & varbtn
        tt.Range.Style = "Additional_file_path"
        tt.Range.Font.Size = "10"
        tt.SetPlaceholderText , , varadditionalfile
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        
        'Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 7).Select
        table1.Rows(table1.Rows.Count - 7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 7, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Score_head"
        tt.SetPlaceholderText Text:="Score"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 7, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Difficulty_head"
        tt.SetPlaceholderText Text:="Difficulty"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
       table1.Rows(table1.Rows.Count - 6).Select
        table1.Rows(table1.Rows.Count - 6).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 6, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , varscoretxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

    table1.Cell(table1.Rows.Count - 6, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , vardifficultytxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 5).Select
        table1.Rows(table1.Rows.Count - 5).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 5, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale_head"
        tt.SetPlaceholderText Text:="Correct Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 5, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Incorrect_Answer_Rationale_head"
        tt.SetPlaceholderText Text:="Incorrect Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 4).Select
        table1.Rows(table1.Rows.Count - 4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 4, 1).Select
        
       
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale"
        tt.SetPlaceholderText Text:=varcorrectrationaletxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 4, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Incorrect_Answer_Rationale"
        tt.SetPlaceholderText Text:=varincorrectrationaletxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
  
       table1.Rows(table1.Rows.Count - 3).Shading.BackgroundPatternColor = RGB(255, 140, 0)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Links_head"
        tt.SetPlaceholderText Text:="Remediation"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 2).Select
        table1.Rows(table1.Rows.Count - 2).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 2, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Link_Type_head"
        tt.SetPlaceholderText Text:="Remediation Type"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 2, 2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Remediation_head"
        tt.SetPlaceholderText Text:="Remediation Detail"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 1).Select
        table1.Rows(table1.Rows.Count - 1).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Cell(table1.Rows.Count - 1, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
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
        
        table1.Cell(table1.Rows.Count - 1, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Text"
        tt.Title = "Remediation Text"
        tt.SetPlaceholderText Text:=varWeblinkRemediationTxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 8).Select
        table1.Rows(table1.Rows.Count - 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
               
        table1.Cell(table1.Rows.Count - 8, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Link"
        tt.Title = "Remediation Link"
        tt.SetPlaceholderText Text:=varWeblinkRemediationLink
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 7).Select
        table1.Rows(table1.Rows.Count - 7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
                
        table1.Cell(table1.Rows.Count - 7, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Tooltip"
        tt.Title = "Remediation Tooltip"
        tt.SetPlaceholderText Text:=varWeblinkRemediationTooltip
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        
        table1.Rows(table1.Rows.Count - 6).Select
        'table1.Rows(table1.Rows.Count - 6).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Rows(table1.Rows.Count - 6).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Link_Type"
        tt.SetPlaceholderText Text:="Ebook"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 5).Select
        table1.Rows(table1.Rows.Count - 5).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        table1.Cell(table1.Rows.Count - 5, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Link_Type"
        tt.SetPlaceholderText Text:="Text"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count - 5, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Text_Remediation_Text"
        tt.Title = "Remediation Text"
        tt.SetPlaceholderText Text:=varTextRemediationTxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        '''''''''Metadata Attributes'''''''''
        
        table1.Rows(table1.Rows.Count - 4).Shading.BackgroundPatternColor = RGB(169, 169, 169)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 4).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Metadata Attributes"
        tt.Range.Style = "Meta_Data_Attributes_head"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 3).Select
        table1.Rows(table1.Rows.Count - 3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 3, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="METADATA TAG"
        tt.Range.Style = "METADATA_TAG_head"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 3, 2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.SetPlaceholderText Text:="METADATA VALUE"
        tt.Range.Style = "METADATA_VALUE_head"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        
        table1.Rows(table1.Rows.Count - 2).Select
        table1.Rows(table1.Rows.Count - 2).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 2, 1).Select
       
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.LockContentControl = True
        tt.Range.Style = "METADATA_TAG"
        tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagtype(0)
        tt.DropdownListEntries.Add "Status"
        tt.DropdownListEntries.Add "Keyword"
        tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count - 2, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagvalue(0)
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 1).Select
        table1.Rows(table1.Rows.Count - 1).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 1, 1).Select
       
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.LockContentControl = True
        tt.Range.Style = "METADATA_TAG"
        tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagtype(1)
        tt.DropdownListEntries.Add "Status"
        tt.DropdownListEntries.Add "Keyword"
        tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count - 1, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
         tt.SetPlaceholderText , , vartagvalue(1)
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        '''''Create Table Row''''
        table1.Rows(table1.Rows.Count).Select
        table1.Rows(table1.Rows.Count).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count, 1).Select
       
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.LockContentControl = True
        tt.Range.Style = "METADATA_TAG"
        tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagtype(2)
        tt.DropdownListEntries.Add "Status"
        tt.DropdownListEntries.Add "Keyword"
        tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
         tt.SetPlaceholderText , , vartagvalue(2)
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        

End Sub
Sub copytableupdateII(vartitle, MyValue, varIdentificationId, varqtext, varanswer, varrationale, varcorrecttxt, varscoretxt, vardifficultytxt, varcorrectrationaletxt, varincorrectrationaletxt, varWeblinkRemediationTxt, varWeblinkRemediationLink, varWeblinkRemediationTooltip, varTextRemediationTxt, vartagtype, vartagvalue, varadditionalfile)
Dim table1 As Table
Dim tt As ContentControl
Dim TABLE2 As Table
Dim addbutton
varNumRows = MyValue + 7
UpdateQuestionNo = UpdateQuestionNo + 1
Selection.HomeKey unit:=wdStory
Selection.EndKey unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 17, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.id = "MC"
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        
        table1.Rows(1).Select
         table1.Rows(1).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
       
       
        table1.Cell(1, 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Delete_Question"
        tt.Title = "Delete Question"
       
        tt.SetPlaceholderText Text:="Delete Question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdRed
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
       
        table1.Cell(1, 4).Select
       
        Selection.Collapse Direction:=wdCollapseStart
        Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
        With ffield
                .Name = "Delete_" & QuestionNo
                .Range.Style = "Delete_box"
        End With
        
        
        table1.Rows(2).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Question Text - Image Integration"
        tt.Range.Style = "Question_Type"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
         table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
       
        table1.Cell(3, 1).Select
        
        
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Question_no"
        tt.Title = "Question No"
        
        tt.SetPlaceholderText Text:="Question No: " & UpdateQuestionNo
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
         table1.Cell(3, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Identification_Id"
        tt.Title = "Identification Id"
        tt.SetPlaceholderText Text:="Identification Id: " & varIdentificationId
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
       
        Set tt = Nothing
        
        table1.Rows(4).Select
        table1.Rows(4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(4, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(4, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Title"
        tt.Range.Style = "Question_Title"
        tt.SetPlaceholderText Text:=vartitle
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
                
        table1.Rows(5).Select
        
        'table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Text"
        tt.Range.Style = "Question_Text"
        tt.SetPlaceholderText Text:=varqtext
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 70
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        'Answer option heading
        table1.Rows(6).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Answer_head"
        tt.SetPlaceholderText Text:="Answer Text - Image Integration"
        
        tt.Range.Font.Size = "13"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        'table1.Cell(4, 2).Select
    
                    
       
        
      For i = 7 To MyValue + 6
        
       ' table1.Rows(5).SetHeight rowHeight:=ActiveDocument.PageSetup.PageHeight - 300, HeightRule:=wdRowHeightExactly
        
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 6
        tt.Range.Style = "Answer_Text"
        tt.SetPlaceholderText Text:=varanswer(i - 7)
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
         table1.Cell(i, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Rationale_Text"
        tt.Title = "Rationale Text " & i - 6
        tt.SetPlaceholderText Text:=varrationale(i - 7)
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
    Next
        
         'table1.Rows(4).Cells.Split NumRows:=5, NumColumns:=2, MergeBeforeSplit:=True
         'table1.Cell(4, 2).Merge table1.Cell(8, 2)
               
        table1.Rows(table1.Rows.Count - 17).Select
        table1.Rows(table1.Rows.Count - 17).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 17, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_answer_head"
        tt.SetPlaceholderText Text:="Correct Answer:"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        table1.Cell(table1.Rows.Count - 17, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Title = "Correct Answer"
        tt.LockContentControl = True
        tt.Range.Style = "Correct_Answer"
        tt.SetPlaceholderText , , varcorrecttxt
        For i = 1 To MyValue
        tt.DropdownListEntries.Add (i)
        Next
               
        table1.Rows(table1.Rows.Count - 8).Select
        table1.Rows(table1.Rows.Count - 8).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 8, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Additional_Fields_head"
        tt.SetPlaceholderText Text:="Additional Fields"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 8, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Additional_Fields_head"
        tt.SetPlaceholderText Text:=""
        tt.LockContentControl = True
        
        table1.Cell(table1.Rows.Count - 8, 3).Select
        'Selection.InlineShapes.AddOLEControl ClassType:="Forms.CommandButton.1"
        Set shp = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
        
        'shp.Range.Style = "Chosse_file"
        With shp.OLEFormat.Object
            .Object.Caption = "Select File"
            '.Object.Name = "Chosse_file"
        End With
        varbtn = Replace(shp.OLEFormat.Object.Name, "CommandButton", "")
        Set shp = Nothing
        
        table1.Cell(table1.Rows.Count - 8, 4).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "File Name" & varbtn
        tt.Range.Style = "Additional_file_path"
        tt.Range.Font.Size = "10"
        tt.SetPlaceholderText , , varadditionalfile
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        'Set tt = Nothing
        
         table1.Rows(table1.Rows.Count - 7).Select
        table1.Rows(table1.Rows.Count - 7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 7, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Score_head"
        tt.SetPlaceholderText Text:="Score"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 7, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Difficulty_head"
        tt.SetPlaceholderText Text:="Difficulty"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 6).Select
        table1.Rows(table1.Rows.Count - 6).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 6, 1).Select
        
        
         table1.Rows(table1.Rows.Count - 6).Select
        table1.Rows(table1.Rows.Count - 6).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 6, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , varscoretxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

        table1.Cell(table1.Rows.Count - 6, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , vardifficultytxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 5).Select
        table1.Rows(table1.Rows.Count - 5).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 5, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale_head"
        tt.SetPlaceholderText Text:="Correct Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 5, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Incorrect_Answer_Rationale_head"
         tt.SetPlaceholderText Text:="Incorrect Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 4).Select
        table1.Rows(table1.Rows.Count - 4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 4, 1).Select
        
       
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale"
        tt.SetPlaceholderText Text:=varcorrectrationaletxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 4, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Incorrect_Answer_Rationale"
        tt.SetPlaceholderText Text:=varincorrectrationaletxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
  
        table1.Rows(table1.Rows.Count - 3).Shading.BackgroundPatternColor = RGB(255, 140, 0)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Links_head"
        tt.SetPlaceholderText Text:="Remediation"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 2).Select
        table1.Rows(table1.Rows.Count - 2).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 2, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Link_Type_head"
        tt.SetPlaceholderText Text:="Remediation Type"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 2, 2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Remediation_head"
        tt.SetPlaceholderText Text:="Remediation Detail"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 1).Select
        table1.Rows(table1.Rows.Count - 1).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Cell(table1.Rows.Count - 1, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
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
        
        table1.Cell(table1.Rows.Count - 1, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Text"
        tt.Title = "Remediation Text"
        tt.SetPlaceholderText Text:=varWeblinkRemediationTxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 8).Select
        table1.Rows(table1.Rows.Count - 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
               
        table1.Cell(table1.Rows.Count - 8, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Link"
        tt.Title = "Remediation Link"
        tt.SetPlaceholderText Text:=varWeblinkRemediationLink
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 7).Select
        table1.Rows(table1.Rows.Count - 7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
                
        table1.Cell(table1.Rows.Count - 7, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Tooltip"
        tt.Title = "Remediation Tooltip"
        tt.SetPlaceholderText Text:=varWeblinkRemediationTooltip
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        
       table1.Rows(table1.Rows.Count - 6).Select
        'table1.Rows(table1.Rows.Count - 6).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Rows(table1.Rows.Count - 6).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Link_Type"
        tt.SetPlaceholderText Text:="Ebook"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 5).Select
        table1.Rows(table1.Rows.Count - 5).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        table1.Cell(table1.Rows.Count - 5, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Link_Type"
        tt.SetPlaceholderText Text:="Text"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count - 5, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Text_Remediation_Text"
        tt.Title = "Remediation Text"
        tt.SetPlaceholderText Text:=varTextRemediationTxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        '''''''''Metadata Attributes'''''''''
        
        table1.Rows(table1.Rows.Count - 4).Shading.BackgroundPatternColor = RGB(169, 169, 169)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 4).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Metadata Attributes"
        tt.Range.Style = "Meta_Data_Attributes_head"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 3).Select
        table1.Rows(table1.Rows.Count - 3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 3, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="METADATA TAG"
        tt.Range.Style = "METADATA_TAG_head"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 3, 2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.SetPlaceholderText Text:="METADATA VALUE"
        tt.Range.Style = "METADATA_VALUE_head"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        
        table1.Rows(table1.Rows.Count - 2).Select
        table1.Rows(table1.Rows.Count - 2).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 2, 1).Select
       
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.LockContentControl = True
        tt.Range.Style = "METADATA_TAG"
        tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagtype(0)
        tt.DropdownListEntries.Add "Status"
        tt.DropdownListEntries.Add "Keyword"
        tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count - 2, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagvalue(0)
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 1).Select
        table1.Rows(table1.Rows.Count - 1).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 1, 1).Select
       
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.LockContentControl = True
        tt.Range.Style = "METADATA_TAG"
        tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagtype(1)
        tt.DropdownListEntries.Add "Status"
        tt.DropdownListEntries.Add "Keyword"
        tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count - 1, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
         tt.SetPlaceholderText , , vartagvalue(1)
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        '''''Create Table Row''''
        table1.Rows(table1.Rows.Count).Select
        table1.Rows(table1.Rows.Count).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count, 1).Select
       
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.LockContentControl = True
        tt.Range.Style = "METADATA_TAG"
        tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagtype(2)
        tt.DropdownListEntries.Add "Status"
        tt.DropdownListEntries.Add "Keyword"
        tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
         tt.SetPlaceholderText , , vartagvalue(2)
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        

End Sub

Sub copytableupdatego(vartitle, MyValue, varIdentificationId, varqtext, varanswer, varrationale, varcorrecttxt, varscoretxt, vardifficultytxt, varcorrectrationaletxt, varincorrectrationaletxt, varWeblinkRemediationTxt, varWeblinkRemediationLink, varWeblinkRemediationTooltip, varTextRemediationTxt, vartagtype, vartagvalue, varadditionalpath)
Dim table1 As Table
Dim tt As ContentControl
Dim TABLE2 As Table
Dim addbutton
varNumRows = MyValue + 7
UpdateQuestionNo = UpdateQuestionNo + 1
Selection.HomeKey unit:=wdStory
Selection.EndKey unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 16, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.id = "VQ"
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        
        
        table1.Rows(1).Select
         table1.Rows(1).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
       
       
        table1.Cell(1, 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Delete_Question"
        tt.Title = "Delete Question"
       
        tt.SetPlaceholderText Text:="Delete Question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdRed
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
       
        table1.Cell(1, 4).Select
       
        Selection.Collapse Direction:=wdCollapseStart
        Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
        With ffield
                .Name = "Delete_" & QuestionNo
                .Range.Style = "Delete_box"
        End With
        
        
        table1.Rows(2).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        
        
        table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Question Text - Graphic Option"
        tt.Range.Style = "Question_Type"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
       
        table1.Cell(3, 1).Select
        
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Question_no"
        tt.Title = "Question No"
        
        tt.SetPlaceholderText Text:="Question No: " & UpdateQuestionNo
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        
        table1.Cell(3, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Identification_Id"
        tt.Title = "Identification Id"
        tt.SetPlaceholderText Text:="Identification Id: " & varIdentificationId
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
       
        Set tt = Nothing
        
        table1.Rows(4).Select
        table1.Rows(4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(4, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(4, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Title"
        tt.Range.Style = "Question_Title"
        tt.SetPlaceholderText Text:=vartitle
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
                
        table1.Rows(5).Select
        
        'table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Text"
        tt.Range.Style = "Question_Text"
        tt.SetPlaceholderText Text:=varqtext
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 70
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        'Answer option heading
        table1.Rows(6).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Answer_head"
        tt.SetPlaceholderText Text:="Answer Text - Graphic Option"
        
        tt.Range.Font.Size = "13"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        'table1.Cell(4, 2).Select
    
                    
       
        
      For i = 7 To MyValue + 6
        
       ' table1.Rows(5).SetHeight rowHeight:=ActiveDocument.PageSetup.PageHeight - 300, HeightRule:=wdRowHeightExactly
        
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=3, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        'Selection.InlineShapes.AddOLEControl ClassType:="Forms.CommandButton.1"
        Set shp = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
        
        'shp.Range.Style = "Chosse_file"
        With shp.OLEFormat.Object
            .Object.Caption = "Select File"
            '.Object.Style = "Chosse_file"
        End With
        varbtn = Replace(shp.OLEFormat.Object.Name, "CommandButton", "")
        Set shp = Nothing
        'table1.Cell(i, 1).Width = 20
        
        
        table1.Cell(i, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "File Name" & varbtn
        tt.Range.Style = "Additional_file_path"
        tt.Range.Font.Size = "10"
        tt.SetPlaceholderText , , varadditionalpath(i - 7)
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(i, 3).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Rationale_Text"
        tt.Title = "Rationale Text " & i - 6
        tt.SetPlaceholderText Text:="Enter rationale for this answer"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
    Next
        
         'table1.Rows(4).Cells.Split NumRows:=5, NumColumns:=2, MergeBeforeSplit:=True
         'table1.Cell(4, 2).Merge table1.Cell(8, 2)
             
        table1.Rows(table1.Rows.Count - 8).Select
        table1.Rows(table1.Rows.Count - 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 8, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_answer_head"
        tt.SetPlaceholderText Text:="Correct Answer:"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 8, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Title = "Correct Answer"
        tt.LockContentControl = True
        tt.Range.Style = "Correct_Answer"
        tt.SetPlaceholderText , , varcorrecttxt
        For i = 1 To MyValue
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
               
        table1.Rows(table1.Rows.Count - 7).Select
        table1.Rows(table1.Rows.Count - 7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 7, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Score_head"
        tt.SetPlaceholderText Text:="Score"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 7, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Difficulty_head"
        tt.SetPlaceholderText Text:="Difficulty"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
       table1.Rows(table1.Rows.Count - 6).Select
        table1.Rows(table1.Rows.Count - 6).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 6, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , varscoretxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

    table1.Cell(table1.Rows.Count - 6, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , vardifficultytxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 5).Select
        table1.Rows(table1.Rows.Count - 5).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 5, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale_head"
        tt.SetPlaceholderText Text:="Correct Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 5, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Incorrect_Answer_Rationale_head"
        tt.SetPlaceholderText Text:="Incorrect Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 4).Select
        table1.Rows(table1.Rows.Count - 4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 4, 1).Select
        
       
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale"
        tt.SetPlaceholderText Text:=varcorrectrationaletxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 4, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Incorrect_Answer_Rationale"
        tt.SetPlaceholderText Text:=varincorrectrationaletxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
  
       table1.Rows(table1.Rows.Count - 3).Shading.BackgroundPatternColor = RGB(255, 140, 0)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Links_head"
        tt.SetPlaceholderText Text:="Remediation"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 2).Select
        table1.Rows(table1.Rows.Count - 2).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 2, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Link_Type_head"
        tt.SetPlaceholderText Text:="Remediation Type"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 2, 2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Remediation_head"
        tt.SetPlaceholderText Text:="Remediation Detail"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 1).Select
        table1.Rows(table1.Rows.Count - 1).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Cell(table1.Rows.Count - 1, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
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
        
        table1.Cell(table1.Rows.Count - 1, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Text"
        tt.Title = "Remediation Text"
        tt.SetPlaceholderText Text:=varWeblinkRemediationTxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 8).Select
        table1.Rows(table1.Rows.Count - 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
               
        table1.Cell(table1.Rows.Count - 8, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Link"
        tt.Title = "Remediation Link"
        tt.SetPlaceholderText Text:=varWeblinkRemediationLink
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 7).Select
        table1.Rows(table1.Rows.Count - 7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
                
        table1.Cell(table1.Rows.Count - 7, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Tooltip"
        tt.Title = "Remediation Tooltip"
        tt.SetPlaceholderText Text:=varWeblinkRemediationTooltip
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        
       table1.Rows(table1.Rows.Count - 6).Select
        'table1.Rows(table1.Rows.Count - 6).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Rows(table1.Rows.Count - 6).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Link_Type"
        tt.SetPlaceholderText Text:="Ebook"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 5).Select
        table1.Rows(table1.Rows.Count - 5).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        table1.Cell(table1.Rows.Count - 5, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Link_Type"
        tt.SetPlaceholderText Text:="Text"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count - 5, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Text_Remediation_Text"
        tt.Title = "Remediation Text"
        tt.SetPlaceholderText Text:=varTextRemediationTxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        '''''''''Metadata Attributes'''''''''
        
        table1.Rows(table1.Rows.Count - 4).Shading.BackgroundPatternColor = RGB(169, 169, 169)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 4).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Metadata Attributes"
        tt.Range.Style = "Meta_Data_Attributes_head"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 3).Select
        table1.Rows(table1.Rows.Count - 3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 3, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="METADATA TAG"
        tt.Range.Style = "METADATA_TAG_head"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 3, 2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.SetPlaceholderText Text:="METADATA VALUE"
        tt.Range.Style = "METADATA_VALUE_head"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        
        table1.Rows(table1.Rows.Count - 2).Select
        table1.Rows(table1.Rows.Count - 2).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 2, 1).Select
       
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.LockContentControl = True
        tt.Range.Style = "METADATA_TAG"
        tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagtype(0)
        tt.DropdownListEntries.Add "Status"
        tt.DropdownListEntries.Add "Keyword"
        tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count - 2, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagvalue(0)
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 1).Select
        table1.Rows(table1.Rows.Count - 1).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 1, 1).Select
       
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.LockContentControl = True
        tt.Range.Style = "METADATA_TAG"
        tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagtype(1)
        tt.DropdownListEntries.Add "Status"
        tt.DropdownListEntries.Add "Keyword"
        tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count - 1, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
         tt.SetPlaceholderText , , vartagvalue(1)
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        '''''Create Table Row''''
        table1.Rows(table1.Rows.Count).Select
        table1.Rows(table1.Rows.Count).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count, 1).Select
       
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.LockContentControl = True
        tt.Range.Style = "METADATA_TAG"
        tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagtype(2)
        tt.DropdownListEntries.Add "Status"
        tt.DropdownListEntries.Add "Keyword"
        tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
         tt.SetPlaceholderText , , vartagvalue(2)
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        

End Sub

Sub copytableupdatecs(vartitle, varIdentificationId, varqtext, vartagtype, vartagvalue, varadditionalfile)
Dim table1 As Table
Dim tt As ContentControl
Dim TABLE2 As Table
Dim addbutton
'varNumRows = MyValue + 7
UpdateQuestionNo = UpdateQuestionNo + 1
Selection.HomeKey unit:=wdStory
Selection.EndKey unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=11, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.id = "VQ"
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        
        
        table1.Rows(1).Select
         table1.Rows(1).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
       
       
        table1.Cell(1, 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Delete_Question"
        tt.Title = "Delete Question"
       
        tt.SetPlaceholderText Text:="Delete Question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdRed
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
       
        table1.Cell(1, 4).Select
       
        Selection.Collapse Direction:=wdCollapseStart
        Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
        With ffield
                .Name = "Delete_" & QuestionNo
                .Range.Style = "Delete_box"
        End With
        
        
        table1.Rows(2).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        
        
        table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Question Text - Clinical Symptoms"
        tt.Range.Style = "Question_Type"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
       
        table1.Cell(3, 1).Select
        
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Question_no"
        tt.Title = "Question No"
        
        tt.SetPlaceholderText Text:="Question No: " & UpdateQuestionNo
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        
        table1.Cell(3, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Identification_Id"
        tt.Title = "Identification Id"
        tt.SetPlaceholderText Text:="Identification Id: " & varIdentificationId
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
       
        Set tt = Nothing
        
        table1.Rows(4).Select
        table1.Rows(4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(4, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(4, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Title"
        tt.Range.Style = "Question_Title"
        tt.SetPlaceholderText Text:=vartitle
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
                
        table1.Rows(5).Select
        
        'table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Text"
        tt.Range.Style = "Question_Text"
        tt.SetPlaceholderText Text:=varqtext
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 70
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        'Answer option heading
        
        
        table1.Rows(6).Select
        table1.Rows(6).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
                
        table1.Cell(6, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Additional_Fields_head"
        tt.SetPlaceholderText Text:="Additional Fields"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(6, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Additional_Fields_head"
        tt.SetPlaceholderText Text:=""
        tt.LockContentControl = True
        
        table1.Cell(6, 3).Select
        'Selection.InlineShapes.AddOLEControl ClassType:="Forms.CommandButton.1"
        Set shp = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
        
        'shp.Range.Style = "Chosse_file"
        With shp.OLEFormat.Object
            .Object.Caption = "Select File"
            '.Object.Style = "Chosse_file"
        End With
        varbtn = Replace(shp.OLEFormat.Object.Name, "CommandButton", "")
        Set shp = Nothing
        'table1.Cell(i, 1).Width = 20
        
        
        table1.Cell(6, 4).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "File Name" & varbtn
        tt.Range.Style = "Additional_file_path"
        tt.Range.Font.Size = "10"
        tt.SetPlaceholderText , , varadditionalfile
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        '''''''''Metadata Attributes'''''''''
        
        table1.Rows(table1.Rows.Count - 4).Shading.BackgroundPatternColor = RGB(169, 169, 169)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 4).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Metadata Attributes"
        tt.Range.Style = "Meta_Data_Attributes_head"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 3).Select
        table1.Rows(table1.Rows.Count - 3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 3, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="METADATA TAG"
        tt.Range.Style = "METADATA_TAG_head"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 3, 2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.SetPlaceholderText Text:="METADATA VALUE"
        tt.Range.Style = "METADATA_VALUE_head"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        
        table1.Rows(table1.Rows.Count - 2).Select
        table1.Rows(table1.Rows.Count - 2).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 2, 1).Select
       
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.LockContentControl = True
        tt.Range.Style = "METADATA_TAG"
        tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagtype(0)
        tt.DropdownListEntries.Add "Status"
        tt.DropdownListEntries.Add "Keyword"
        tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count - 2, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagvalue(0)
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 1).Select
        table1.Rows(table1.Rows.Count - 1).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 1, 1).Select
       
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.LockContentControl = True
        tt.Range.Style = "METADATA_TAG"
        tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagtype(1)
        tt.DropdownListEntries.Add "Status"
        tt.DropdownListEntries.Add "Keyword"
        tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count - 1, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
         tt.SetPlaceholderText , , vartagvalue(1)
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        '''''Create Table Row''''
        table1.Rows(table1.Rows.Count).Select
        table1.Rows(table1.Rows.Count).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count, 1).Select
       
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.LockContentControl = True
        tt.Range.Style = "METADATA_TAG"
        tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagtype(2)
        tt.DropdownListEntries.Add "Status"
        tt.DropdownListEntries.Add "Keyword"
        tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
         tt.SetPlaceholderText , , vartagvalue(2)
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        

End Sub

Sub copytableupdateMEDC(vartitle, varIdentificationId, varqtext, vartagtype, vartagvalue, varadditionalfile)
Dim table1 As Table
Dim tt As ContentControl
Dim TABLE2 As Table
Dim addbutton
'varNumRows = MyValue + 7
UpdateQuestionNo = UpdateQuestionNo + 1
Selection.HomeKey unit:=wdStory
Selection.EndKey unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=11, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.id = "VQ"
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        
        
        table1.Rows(1).Select
         table1.Rows(1).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
       
       
        table1.Cell(1, 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Delete_Question"
        tt.Title = "Delete Question"
       
        tt.SetPlaceholderText Text:="Delete Question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdRed
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
       
        table1.Cell(1, 4).Select
       
        Selection.Collapse Direction:=wdCollapseStart
        Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
        With ffield
                .Name = "Delete_" & QuestionNo
                .Range.Style = "Delete_box"
        End With
        
        
        table1.Rows(2).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        
        
        table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Question Text - Medical Case"
        tt.Range.Style = "Question_Type"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
       
        table1.Cell(3, 1).Select
        
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Question_no"
        tt.Title = "Question No"
        
        tt.SetPlaceholderText Text:="Question No: " & UpdateQuestionNo
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        
        table1.Cell(3, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Identification_Id"
        tt.Title = "Identification Id"
        tt.SetPlaceholderText Text:="Identification Id: " & varIdentificationId
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
       
        Set tt = Nothing
        
        table1.Rows(4).Select
        table1.Rows(4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(4, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(4, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Title"
        tt.Range.Style = "Question_Title"
        tt.SetPlaceholderText Text:=vartitle
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
                
        table1.Rows(5).Select
        
        'table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Text"
        tt.Range.Style = "Question_Text"
        tt.SetPlaceholderText Text:=varqtext
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 70
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        'Answer option heading
        
        
        table1.Rows(6).Select
        table1.Rows(6).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
                
        table1.Cell(6, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Additional_Fields_head"
        tt.SetPlaceholderText Text:="Additional Fields"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(6, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Additional_Fields_head"
        tt.SetPlaceholderText Text:=""
        tt.LockContentControl = True
        
        table1.Cell(6, 3).Select
        Set shp = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
        With shp.OLEFormat.Object
            .Object.Caption = "Select File"
            
        End With
        varbtn = Replace(shp.OLEFormat.Object.Name, "CommandButton", "")
        Set shp = Nothing
        
        
        table1.Cell(6, 4).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "File Name" & varbtn
        tt.Range.Style = "Additional_file_path"
        tt.Range.Font.Size = "10"
        tt.SetPlaceholderText , , varadditionalfile
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        '''''''''Metadata Attributes'''''''''
        
        table1.Rows(table1.Rows.Count - 4).Shading.BackgroundPatternColor = RGB(169, 169, 169)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 4).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Metadata Attributes"
        tt.Range.Style = "Meta_Data_Attributes_head"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 3).Select
        table1.Rows(table1.Rows.Count - 3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 3, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="METADATA TAG"
        tt.Range.Style = "METADATA_TAG_head"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 3, 2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.SetPlaceholderText Text:="METADATA VALUE"
        tt.Range.Style = "METADATA_VALUE_head"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        
        table1.Rows(table1.Rows.Count - 2).Select
        table1.Rows(table1.Rows.Count - 2).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 2, 1).Select
       
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.LockContentControl = True
        tt.Range.Style = "METADATA_TAG"
        tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagtype(0)
        tt.DropdownListEntries.Add "Status"
        tt.DropdownListEntries.Add "Keyword"
        tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count - 2, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagvalue(0)
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 1).Select
        table1.Rows(table1.Rows.Count - 1).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 1, 1).Select
       
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.LockContentControl = True
        tt.Range.Style = "METADATA_TAG"
        tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagtype(1)
        tt.DropdownListEntries.Add "Status"
        tt.DropdownListEntries.Add "Keyword"
        tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count - 1, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
         tt.SetPlaceholderText , , vartagvalue(1)
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        '''''Create Table Row''''
        table1.Rows(table1.Rows.Count).Select
        table1.Rows(table1.Rows.Count).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count, 1).Select
       
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.LockContentControl = True
        tt.Range.Style = "METADATA_TAG"
        tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , vartagtype(2)
        tt.DropdownListEntries.Add "Status"
        tt.DropdownListEntries.Add "Keyword"
        tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
         tt.SetPlaceholderText , , vartagvalue(2)
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        

End Sub

Sub copytableupdatemcCSMEDC(vartitle, MyValue, varIdentificationId, varqtext, varanswer, varrationale, varcorrecttxt, varscoretxt, vardifficultytxt, varcorrectrationaletxt, varincorrectrationaletxt, varWeblinkRemediationTxt, varWeblinkRemediationLink, varWeblinkRemediationTooltip, varTextRemediationTxt)
Dim table1 As Table
Dim tt As ContentControl
Dim TABLE2 As Table
Dim addbutton
varNumRows = MyValue + 6
UpdateQuestionNo = UpdateQuestionNo + 1
Selection.HomeKey unit:=wdStory
Selection.EndKey unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 12, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.id = "MC"
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(1).Select
         table1.Rows(1).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        
        table1.Cell(1, 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Delete_Question"
        tt.Title = "Delete Question"
        
        tt.SetPlaceholderText Text:="Delete Question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdRed
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        table1.Cell(1, 4).Select
        
        Selection.Collapse Direction:=wdCollapseStart
        Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
        With ffield
                .Name = "Delete_" & QuestionNo
                .Range.Style = "Delete_box"
        End With
        
        table1.Rows(2).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Question Text - Multiple Choice"
        tt.Range.Style = "Question_Type"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
       
        table1.Cell(3, 1).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Question_no"
        tt.Title = "Question No"
        tt.SetPlaceholderText Text:="Question No: " & UpdateQuestionNo
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(3, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Identification_Id"
        tt.Title = "Identification Id"
        tt.SetPlaceholderText Text:="Identification Id: " & varIdentificationId
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        
        table1.Rows(4).Select
        table1.Rows(4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(4, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(4, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Title"
        tt.Range.Style = "Question_Title"
        tt.SetPlaceholderText Text:=vartitle
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
                
        table1.Rows(5).Select
        
        'table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Text"
        tt.Range.Style = "Question_Text"
        'tt.ShowingPlaceholderText = varqtext
        tt.SetPlaceholderText Text:=varqtext
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 70
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        'Answer option heading
        table1.Rows(6).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Answer_head"
        tt.SetPlaceholderText Text:="Answer Text - Multiple Choice"
        
        tt.Range.Font.Size = "13"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
    
         
        
      For i = 7 To MyValue + 6
        
       ' table1.Rows(5).SetHeight rowHeight:=ActiveDocument.PageSetup.PageHeight - 300, HeightRule:=wdRowHeightExactly
        
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 6
        tt.Range.Style = "Answer_Text"
        tt.SetPlaceholderText Text:=varanswer(i - 7)
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
         table1.Cell(i, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Rationale_Text"
        tt.Title = "Rationale Text " & i - 6
        tt.SetPlaceholderText Text:=varrationale(i - 7)
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
    Next
        
         'table1.Rows(4).Cells.Split NumRows:=5, NumColumns:=2, MergeBeforeSplit:=True
         'table1.Cell(4, 2).Merge table1.Cell(8, 2)
               
      
               
        table1.Rows(table1.Rows.Count - 3).Select
        table1.Rows(table1.Rows.Count - 3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 3, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_answer_head"
        tt.SetPlaceholderText Text:="Correct Answer:"
        
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 3, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Title = "Correct Answer"
        tt.LockContentControl = True
        tt.Range.Style = "Correct_Answer"
        tt.SetPlaceholderText , , varcorrecttxt
        For i = 1 To MyValue
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        'Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 2).Select
        table1.Rows(table1.Rows.Count - 2).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 2, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Score_head"
        tt.SetPlaceholderText Text:="Score"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 2, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Difficulty_head"
        tt.SetPlaceholderText Text:="Difficulty"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 1).Select
        table1.Rows(table1.Rows.Count - 1).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 1, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , varscoretxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

        table1.Cell(table1.Rows.Count - 1, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , vardifficultytxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 8).Select
        table1.Rows(table1.Rows.Count - 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 8, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale_head"
        tt.SetPlaceholderText Text:="Correct Rationale"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 8, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Incorrect_Answer_Rationale_head"
         tt.SetPlaceholderText Text:="Incorrect Rationale"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 7).Select
        table1.Rows(table1.Rows.Count - 7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 7, 1).Select
        
       
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale"
        tt.SetPlaceholderText Text:=varcorrectrationaletxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 7, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.SetPlaceholderText Text:=varincorrectrationaletxt
        tt.Range.Style = "Incorrect_Answer_Rationale"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
  
        table1.Rows(table1.Rows.Count - 6).Shading.BackgroundPatternColor = RGB(255, 140, 0)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 6).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Links_head"
        tt.SetPlaceholderText Text:="Remediation"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 5).Select
        table1.Rows(table1.Rows.Count - 5).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 5, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Link_Type_head"
        tt.SetPlaceholderText Text:="Remediation Type"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 5, 2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Remediation_head"
        tt.SetPlaceholderText Text:="Remediation Detail"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 4).Select
        table1.Rows(table1.Rows.Count - 4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Cell(table1.Rows.Count - 4, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
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
        
        table1.Cell(table1.Rows.Count - 4, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Text"
        tt.Title = "Remediation Text"
        tt.SetPlaceholderText Text:=varWeblinkRemediationTxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 3).Select
        table1.Rows(table1.Rows.Count - 3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
               
        table1.Cell(table1.Rows.Count - 3, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Link"
        tt.Title = "Remediation Link"
        tt.SetPlaceholderText Text:=varWeblinkRemediationLink
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 2).Select
        table1.Rows(table1.Rows.Count - 2).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
                
        table1.Cell(table1.Rows.Count - 2, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Tooltip"
        tt.Title = "Remediation Tooltip"
        tt.SetPlaceholderText Text:=varWeblinkRemediationTooltip
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        
      table1.Rows(table1.Rows.Count - 1).Select
        'table1.Rows(table1.Rows.Count - 6).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Rows(table1.Rows.Count - 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Link_Type"
        tt.SetPlaceholderText Text:="Ebook"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count).Select
        table1.Rows(table1.Rows.Count).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        table1.Cell(table1.Rows.Count, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Link_Type"
        tt.SetPlaceholderText Text:="Text"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Text_Remediation_Text"
        tt.Title = "Remediation Text"
        tt.SetPlaceholderText Text:=varTextRemediationTxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        

End Sub

Sub copytableupdatecmCSMEDC(vartitle, MyValue, varIdentificationId, varqtext, varanswer, varrationale, varcorrecttxt, varscoretxt, vardifficultytxt, varcorrectrationaletxt, varincorrectrationaletxt, varWeblinkRemediationTxt, varWeblinkRemediationLink, varWeblinkRemediationTooltip, varTextRemediationTxt)
Dim table1 As Table
Dim tt As ContentControl
Dim TABLE2 As Table
Dim addbutton
varNumRows = MyValue + 6
UpdateQuestionNo = UpdateQuestionNo + 1
Selection.HomeKey unit:=wdStory
Selection.EndKey unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 12, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        
        table1.Rows(1).Select
         table1.Rows(1).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        
        table1.Cell(1, 3).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Delete_Question"
        tt.Title = "Delete Question"
        
        tt.SetPlaceholderText Text:="Delete Question"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdRed
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        table1.Cell(1, 4).Select
        
        Selection.Collapse Direction:=wdCollapseStart
        Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
        With ffield
                .Name = "Delete_" & QuestionNo
                .Range.Style = "Delete_box"
        End With
        
        
        
        table1.Rows(2).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Question Text - Choice Multiple"
        tt.Range.Style = "Question_Type"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(3, 1).Select
        
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Question_no"
        tt.Title = "Question No"
        
        tt.SetPlaceholderText Text:="Question No: " & UpdateQuestionNo
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(3, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Identification_Id"
        tt.Title = "Identification Id"
        tt.SetPlaceholderText Text:="Identification Id: " & varIdentificationId
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Rows(4).Select
        table1.Rows(4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(4, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.Range.Style = "Question_head"
        tt.SetPlaceholderText Text:="Question Title"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = True
        
        Set tt = Nothing
        
        table1.Cell(4, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Title"
        tt.Range.Style = "Question_Title"
        tt.SetPlaceholderText Text:=vartitle
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        'tt.Range.Rows.Height = 10
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
        table1.Rows(5).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Question Text"
        tt.Range.Style = "Question_Text"
        tt.SetPlaceholderText Text:=varqtext
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        tt.Range.Rows.Height = 70
        Set tt = Nothing
        
        'Answer option heading
        table1.Rows(6).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Answer_head"
        tt.SetPlaceholderText Text:="Answer Text - Choice Multiple"
        
        tt.Range.Font.Size = "13"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
      
        
       ' table1.Rows(5).SetHeight rowHeight:=ActiveDocument.PageSetup.PageHeight - 300, HeightRule:=wdRowHeightExactly
        For i = 7 To MyValue + 6
            
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 6
        tt.Range.Style = "Answer_Text"
        tt.SetPlaceholderText Text:=varanswer(i - 7)
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
        
         table1.Cell(i, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Rationale_Text"
        tt.Title = "Rationale Text " & i - 6
        tt.SetPlaceholderText Text:=varrationale(i - 7)
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
         
        Next
               
        table1.Rows(table1.Rows.Count - 3).Select
        table1.Rows(table1.Rows.Count - 3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 3, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Correct Answers:"
        tt.Range.Style = "Correct_Answer"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        Dim strCorrect As String
        Dim strArray() As String
        Dim intCount As Integer
   
        strCorrect = varcorrecttxt
        strArray = Split(strCorrect, ",")
        
        For i = 2 To (MyValue * 2) + 1
        
        If i = 2 Then
            j = 1
            K = 1
            table1.Cell(table1.Rows.Count - 3, i).Split NumColumns:=(MyValue * 2) + 1
            table1.Cell(table1.Rows.Count - 3, 2).Select
            Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
            tt.SetPlaceholderText Text:=" "
            tt.Range.Style = "Check_box_no"
            tt.Range.Font.Size = "10"
            tt.Range.Font.Name = "Verdana"
            tt.LockContentControl = True
            tt.LockContents = False
            
        End If
        
        If i Mod 2 > 0 Then
            table1.Cell(table1.Rows.Count - 3, i).Select
            
            Selection.Collapse Direction:=wdCollapseEnd
            Set ffield = ActiveDocument.FormFields _
                        .Add(Range:=Selection.Range, Type:=wdFieldFormCheckBox)
            With ffield
                .Name = "Check_" & UpdateQuestionNo & "_" & K
                .Range.Style = "Check_box"
                '.CheckBox.Value = False
                For intCount = LBound(strArray) To UBound(strArray)
                    If strArray(intCount) = K Then
                        .CheckBox.Value = True
                    End If
                Next
            End With
            K = K + 1
        Else
            table1.Cell(table1.Rows.Count - 3, i).Select
            Selection.Collapse Direction:=wdCollapseEnd
            Set tt = ActiveDocument.ContentControls.Add(wdContentControlText)
                
            tt.SetPlaceholderText Text:=j
            tt.Range.Style = "Check_box_no"
            Set tt = Nothing
            j = j + 1
        End If
        
        Next
        
        
        'Set tt = Nothing
        table1.Rows(table1.Rows.Count - 2).Select
        table1.Rows(table1.Rows.Count - 2).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 2, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Score_head"
        tt.SetPlaceholderText Text:="Score"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 2, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Difficulty_head"
        tt.SetPlaceholderText Text:="Difficulty"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 1).Select
        table1.Rows(table1.Rows.Count - 1).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 1, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , varscoretxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

    table1.Cell(table1.Rows.Count - 1, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , vardifficultytxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 8).Select
        table1.Rows(table1.Rows.Count - 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 8, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale_head"
        tt.SetPlaceholderText Text:="Correct Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 8, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Incorrect_Answer_Rationale_head"
         tt.SetPlaceholderText Text:="Incorrect Rationale"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 7).Select
        table1.Rows(table1.Rows.Count - 7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 7, 1).Select
        
       
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Correct_Answer_Rationale"
        tt.SetPlaceholderText Text:=varcorrectrationaletxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 7, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Incorrect_Answer_Rationale"
        tt.SetPlaceholderText Text:=varincorrectrationaletxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
  
        table1.Rows(table1.Rows.Count - 6).Shading.BackgroundPatternColor = RGB(255, 140, 0)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 6).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Links_head"
        tt.SetPlaceholderText Text:="Remediation"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 5).Select
        table1.Rows(table1.Rows.Count - 5).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 5, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Link_Type_head"
        tt.SetPlaceholderText Text:="Remediation Type"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        table1.Cell(table1.Rows.Count - 5, 2).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
         tt.Range.Style = "Remediation_head"
        tt.SetPlaceholderText Text:="Remediation Detail"
        'tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        
        
        table1.Rows(table1.Rows.Count - 4).Select
        table1.Rows(table1.Rows.Count - 4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Cell(table1.Rows.Count - 4, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
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
        
        table1.Cell(table1.Rows.Count - 4, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Text"
        tt.Title = "Remediation Text"
        tt.SetPlaceholderText Text:=varWeblinkRemediationTxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 3).Select
        table1.Rows(table1.Rows.Count - 3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
               
        table1.Cell(table1.Rows.Count - 3, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Link"
        tt.Title = "Remediation Link"
        tt.SetPlaceholderText Text:=varWeblinkRemediationLink
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 2).Select
        table1.Rows(table1.Rows.Count - 2).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
                
        table1.Cell(table1.Rows.Count - 2, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Tooltip"
        tt.Title = "Remediation Tooltip"
        tt.SetPlaceholderText Text:=varWeblinkRemediationTooltip
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        
       table1.Rows(table1.Rows.Count - 1).Select
        'table1.Rows(table1.Rows.Count - 6).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Rows(table1.Rows.Count - 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Link_Type"
        tt.SetPlaceholderText Text:="Ebook"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count).Select
        table1.Rows(table1.Rows.Count).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        table1.Cell(table1.Rows.Count, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Remediation_Link_Type"
        tt.SetPlaceholderText Text:="Text"
        ''tt.Range.Style = "Normal"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Text_Remediation_Text"
        tt.Title = "Remediation Text"
        tt.SetPlaceholderText Text:=varTextRemediationTxt
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
       

End Sub

Public Function getobject(varobjectname As String) As OLEFormat
'"Qno_" & QuestionNo & "_txt_Metadata_Search_keyword"
For Each shp In ActiveDocument.InlineShapes
    If shp.OLEFormat.Object.Name = varobjectname Then
     Set getobject = shp.OLEFormat
     Exit For
    End If
Next
End Function
Sub remfun(objname As String, table1 As Table, id As Integer, qno, qtype)
'MsgBox qtype
Dim table21 As Table
Dim tt As ContentControl
Dim scode As String
scode = ""
'table21.Shading.BackgroundPatternColor = wdColorBlack
Set table21 = table1
Select Case (Split(objname, "_")(0))
Case "Rem"
            
     For vari = 0 To UBound(Split(objname, "_"))
        If Split(objname, "_")(vari) = "rdtWL" Then
             table21.Rows(table21.Rows.Count).Select
            ' If InStr(Selection.Range.Text, "Enter") > 0 Or InStr(Selection.Range.Text, "Meta") > 0 Then table21.Rows.Add
             table21.Rows(table21.Rows.Count).Select
             table21.Rows(table21.Rows.Count).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
             table21.Cell(table21.Rows.Count, 1).Select
             Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
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
        
             table21.Cell(table21.Rows.Count, 2).Select
             Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
             tt.Range.Style = "Web_Remediation_Text"
             tt.Title = "Remediation Text"
             tt.SetPlaceholderText Text:="Enter remediation text"
             tt.Range.Font.Size = "10"
             tt.Range.Font.ColorIndex = wdBlack
             tt.Range.Font.Bold = False
             'tt.Range.Rows.Height = 30
             tt.Range.Font.Name = "Verdana"
             tt.LockContentControl = True
             tt.LockContents = False
             Set tt = Nothing
            
             table21.Rows.Add
             table21.Rows(table21.Rows.Count).Select
             table21.Rows(table21.Rows.Count).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
               
        table21.Cell(table21.Rows.Count, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Link"
        tt.Title = "Remediation Link"
        tt.SetPlaceholderText Text:="Enter remediation link"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
         table21.Rows.Add
        table21.Rows(table21.Rows.Count).Select
        table21.Rows(table21.Rows.Count).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
                
        table21.Cell(table21.Rows.Count, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Web_Remediation_Tooltip"
        tt.Title = "Remediation Tooltip"
        tt.SetPlaceholderText Text:="Enter remediation tooltip"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
         table21.Rows.Add
           'remediation_type_text
           'remediation_type_link
           'remediation_type_tooltip
            Exit For
        ElseIf Split(objname, "_")(vari) = "rdtEB" Then
         table21.Rows(table21.Rows.Count).Select
        ' If InStr(Selection.Range.Text, "Enter") > 0 Or InStr(Selection.Range.Text, "Meta") > 0 Then table21.Rows.Add
             table21.Rows(table21.Rows.Count).Select
             table21.Rows(table21.Rows.Count).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
             table21.Cell(table21.Rows.Count, 1).Select
             Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
             tt.Range.Style = "Remediation_Link_Type"
             tt.SetPlaceholderText Text:="EBook"
             ''tt.Range.Style = "Normal"
             tt.Range.Font.Size = "10"
             tt.Range.Font.ColorIndex = wdBlack
             tt.Range.Font.Bold = False
             tt.Range.Rows.Height = 30
             tt.Range.Font.Name = "Verdana"
             tt.LockContentControl = True
             tt.LockContents = True
             Set tt = Nothing
        
             table21.Cell(table21.Rows.Count, 2).Select
             Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
             tt.Range.Style = "Web_Remediation_Text"
             tt.Title = "Remediation Text"
             tt.SetPlaceholderText Text:="Enter EBook text"
             tt.Range.Font.Size = "10"
             tt.Range.Font.ColorIndex = wdBlack
             tt.Range.Font.Bold = False
             'tt.Range.Rows.Height = 30
             tt.Range.Font.Name = "Verdana"
             tt.LockContentControl = True
             tt.LockContents = False
             Set tt = Nothing
            
             table21.Rows.Add
             table21.Rows(table21.Rows.Count).Select
             table21.Rows(table21.Rows.Count).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
               
'        table21.Cell(table21.Rows.Count, 2).Select
'        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
'        tt.Range.Style = "Web_Remediation_Link"
'        tt.Title = "Remediation Link"
'        tt.SetPlaceholderText Text:="Enter EBook link"
'        tt.Range.Font.Size = "10"
'        tt.Range.Font.ColorIndex = wdBlack
'        tt.Range.Font.Bold = False
'        'tt.Range.Rows.Height = 30
 '       tt.Range.Font.Name = "Verdana"
 '       tt.LockContentControl = True
 '       tt.LockContents = False
 '       Set tt = Nothing
 '        table21.Rows.Add
 '       table21.Rows(table21.Rows.Count).Select
 '       table21.Rows(table21.Rows.Count).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
                
  '      table21.Cell(table21.Rows.Count, 2).Select
  '      Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
  '      tt.Range.Style = "Web_Remediation_Tooltip"
  '      tt.Title = "Remediation Tooltip"
  '      tt.SetPlaceholderText Text:="Enter EBook tooltip"
   '     tt.Range.Font.Size = "10"
   '     tt.Range.Font.ColorIndex = wdBlack
   '     tt.Range.Font.Bold = False
   '     'tt.Range.Rows.Height = 30
   '     tt.Range.Font.Name = "Verdana"
   '     tt.LockContentControl = True
   '     tt.LockContents = False
   '     Set tt = Nothing
         table21.Rows.Add
           'remediation_type_text
           'remediation_type_link
           'remediation_type_tooltip
            Exit For
        ElseIf Split(objname, "_")(vari) = "rdtText" Then
             table21.Rows(table21.Rows.Count).Select
         '    If InStr(Selection.Range.Text, "Enter") > 0 Or InStr(Selection.Range.Text, "Meta") > 0 Then table21.Rows.Add
             table21.Rows(table21.Rows.Count).Select
             table21.Rows(table21.Rows.Count).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
             table21.Cell(table21.Rows.Count, 1).Select
             Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
             tt.Range.Style = "Remediation_Link_Type"
             tt.SetPlaceholderText Text:="Text"
             ''tt.Range.Style = "Normal"
             tt.Range.Font.Size = "10"
             tt.Range.Font.ColorIndex = wdBlack
             tt.Range.Font.Bold = False
             tt.Range.Rows.Height = 30
             tt.Range.Font.Name = "Verdana"
             tt.LockContentControl = True
             tt.LockContents = True
             Set tt = Nothing
        
             table21.Cell(table21.Rows.Count, 2).Select
             Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
             tt.Range.Style = "Web_Remediation_Text"
             tt.Title = "Remediation Text"
             tt.SetPlaceholderText Text:="Enter remediation text"
             tt.Range.Font.Size = "10"
             tt.Range.Font.ColorIndex = wdBlack
             tt.Range.Font.Bold = False
             'tt.Range.Rows.Height = 30
             tt.Range.Font.Name = "Verdana"
             tt.LockContentControl = True
             tt.LockContents = False
             Set tt = Nothing
            
            
               
                
       
         table21.Rows.Add
            'remediation_type_text
            Exit For
        End If
     Next
     
     
    'MsgBox "ok"
Case "MD"
    For vari = 0 To UBound(Split(objname, "_"))
        If Split(objname, "_")(vari) = "MDtFT" Then
             table21.Rows(table21.Rows.Count).Select
        '     If InStr(Selection.Range.Text, "Enter") > 0 Or InStr(Selection.Range.Text, "Meta") > 0 Then table21.Rows.Add
             table21.Rows(table21.Rows.Count).Select
             table21.Rows(table21.Rows.Count).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
             table21.Cell(table21.Rows.Count, 1).Select
             Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
             tt.Range.Style = "Remediation_Link_Type"
             tt.SetPlaceholderText Text:="Free Text"
             ''tt.Range.Style = "Normal"
             tt.Range.Font.Size = "10"
             tt.Range.Font.ColorIndex = wdBlack
             tt.Range.Font.Bold = False
             tt.Range.Rows.Height = 30
             tt.Range.Font.Name = "Verdana"
             tt.LockContentControl = True
             tt.LockContents = True
             Set tt = Nothing
             table21.Cell(table21.Rows.Count, 2).Select
             Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
             tt.Range.Style = "Web_Remediation_Text"
             tt.Title = "Free Text"
             tt.SetPlaceholderText Text:="Enter Free Text"
             tt.Range.Font.Size = "10"
             tt.Range.Font.ColorIndex = wdBlack
             tt.Range.Font.Bold = False
             'tt.Range.Rows.Height = 30
             tt.Range.Font.Name = "Verdana"
             tt.LockContentControl = True
             tt.LockContents = False
             Set tt = Nothing
             table21.Rows.Add
        ElseIf Split(objname, "_")(vari) = "MDtLU" Then
            table21.Rows(table21.Rows.Count).Select
            ' If InStr(Selection.Range.Text, "Enter") > 0 Or InStr(Selection.Range.Text, "Meta") > 0 Then table21.Rows.Add
             table21.Rows(table21.Rows.Count).Select
             table21.Rows(table21.Rows.Count).Cells.Split NumColumns:=3, MergeBeforeSplit:=True
               
             table21.Cell(table21.Rows.Count, 1).Select
             Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
             tt.Range.Style = "Remediation_Link_Type"
             tt.SetPlaceholderText Text:="Look Up"
             ''tt.Range.Style = "Normal"
             tt.Range.Font.Size = "10"
             tt.Range.Font.ColorIndex = wdBlack
             tt.Range.Font.Bold = False
             tt.Range.Rows.Height = 30
             tt.Range.Font.Name = "Verdana"
             tt.LockContentControl = True
             tt.LockContents = True
             Set tt = Nothing
             table21.Cell(table21.Rows.Count, 2).Select
                Set shp = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.TextBox.1")
            '  shp.OLEFormat.Object.Object.Width = 200
            shp.OLEFormat.Object.Name = "MD_LU_txt_Search_qno" & qno & "_MDtLU_qt" & qtype & "_" & Selection.Cells(1).RowIndex
             table21.Cell(table21.Rows.Count, 3).Select
             Set shp = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
             'shp.Range.Style = "Chosse_file"
              With shp.OLEFormat.Object
                   .Object.Caption = "Search"
                   '.Object.Name = "Chosse_file"
              End With
              shp.OLEFormat.Object.Name = "MD_LU_cmd_Search_qno" & qno & "_MDtLU_qt" & qtype & "_" & Selection.Cells(1).RowIndex
              varbtn = shp.OLEFormat.Object.Name
               scode = scode & "Private Sub " & varbtn & "_Click()" & vbCrLf
            scode = scode & "If Len(" & Replace(varbtn, "_cmd", "_txt") & ".text) >= 3 Then" & vbCrLf
            scode = scode & " call LU(""" & varbtn & """,selection.tables(1)," & table1.Rows.Count & "," & Replace(varbtn, "_cmd", "_txt") & ".text)" & vbCrLf
            'scode = scode & varbtn & ".enabled=false" & vbCrLf
            scode = scode & "Else" & vbCrLf
            scode = scode & "MsgBox ""Search Key word minimum three character.."", vbCritical, ""WK Quizzing Platform""" & vbCrLf
            scode = scode & "End If" & vbCrLf
            
            scode = scode & "End Sub"
          '  ActiveDocument.VBProject.VBComponents("ThisDocument").CodeModule.AddFromString scode
            table21.Rows.Add
            table21.Rows.Add
            Exit For
        ElseIf Split(objname, "_")(vari) = "MDtHI" Then
              table21.Rows(table21.Rows.Count).Select
            '  If InStr(Selection.Range.Text, "Enter") > 0 Or InStr(Selection.Range.Text, "Meta") > 0 Then table21.Rows.Add
             table21.Rows(table21.Rows.Count).Select
             table21.Rows(table21.Rows.Count).Cells.Split NumColumns:=3, MergeBeforeSplit:=True
               
             table21.Cell(table21.Rows.Count, 1).Select
             Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
             tt.Range.Style = "Remediation_Link_Type"
             tt.SetPlaceholderText Text:="Hierarchy"
             ''tt.Range.Style = "Normal"
             tt.Range.Font.Size = "10"
             tt.Range.Font.ColorIndex = wdBlack
             tt.Range.Font.Bold = False
             tt.Range.Rows.Height = 30
             tt.Range.Font.Name = "Verdana"
             tt.LockContentControl = True
             tt.LockContents = True
             Set tt = Nothing
             table21.Cell(table21.Rows.Count, 2).Select
                Set shp = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.TextBox.1")
              'shp.OLEFormat.Object.Object.Width = 200
              
            shp.OLEFormat.Object.Name = "MD_HI_txt_Search_qno" & qno & "_MDtHI_qt" & qtype & "_" & Selection.Cells(1).RowIndex
'            MsgBox shp.OLEFormat.Object.Name
            
             table21.Cell(table21.Rows.Count, 3).Select
             Set shp = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
             'shp.Range.Style = "Chosse_file"
              With shp.OLEFormat.Object
                   .Object.Caption = "Search"
                   '.Object.Name = "Chosse_file"
              End With
        shp.OLEFormat.Object.Name = "MD_HI_cmd_Search_qno" & qno & "_MDtHI_qt" & qtype & "_" & Selection.Cells(1).RowIndex
        varbtn = shp.OLEFormat.Object.Name
            scode = scode & "Private Sub " & varbtn & "_Click()" & vbCrLf
            scode = scode & "If Len(" & Replace(varbtn, "_cmd", "_txt") & ".text) >= 3 Then" & vbCrLf
            scode = scode & " call HI(""" & varbtn & """,selection.tables(1)," & table1.Rows.Count & "," & Replace(varbtn, "_cmd", "_txt") & ".text)" & vbCrLf
            scode = scode & "Else" & vbCrLf
            scode = scode & "MsgBox ""Enter minimum 3 characters to search meta data."", vbCritical, ""WK Quizzing Platform""" & vbCrLf
            scode = scode & "End If" & vbCrLf
            
            scode = scode & "End Sub"
            'MsgBox scode
           
             table21.Rows.Add
            table21.Rows.Add
            Exit For
        End If
     Next
End Select
If ActiveDocument.VBProject.VBComponents("ThisDocument").CodeModule.ProcOfLine(ActiveDocument.VBProject.VBComponents("ThisDocument").CodeModule.CountOfDeclarationLines + 1, vbext_pk_Proc) = "Private Sub " & varbtn & "_Click()" Then
'MsgBox "ok"
Else
 ActiveDocument.VBProject.VBComponents("ThisDocument").CodeModule.AddFromString scode
End If
'table21.Rows.Add
End Sub
Sub LU(objname As String, table1 As Table, id As Integer, searchtext As String)
addinglist objname, table1, id, searchtext, "LOOKUP", "MC"

End Sub
Sub HI(objname As String, table1 As Table, id As Integer, searchtext As String)
'MsgBox objname
addinglist objname, table1, id, searchtext, "Hierarchy", "MC"
End Sub
Sub addinglist(objname As String, table1 As Table, id As Integer, searchtext As String, metadatatype, qtype)
'Sub HI()
Dim metadata As Collection
Dim JSONTEXT As String
If table1.Rows(id + 1).Range.InlineShapes.Count = 0 Then
table1.Rows(id + 1).Cells.Merge
table1.Cell(id + 1, 1).Select
Set metadata = config(ActiveDocument.Path & "\config.ini")
JSONTEXT = webservices(metadata.Item(2), metadata.Item(1), metadata.Item(3), metadata.Item(4), searchtext, metadatatype)
varuid = customxmlcreation(JSONTEXT, 1, qtype)
varuid = Replace(Replace(varuid, "{", "___"), "}", "____")
varuid = Replace(varuid, "-", "__")
varreplaceid = Replace(Replace(Replace(varuid, "___", "{"), "{_", "}"), "__", "-")

'ActiveDocument.CustomXMLParts.SelectByID(VARREPLACEID).XML
        Set shp = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.ListBox.1")
         'MsgBox shp.OLEFormat.Object.Name
        shp.OLEFormat.Object.Name = Replace(objname, "_cmd", "_lst") & varuid
        'shp.OLEFormat.Object.Height = table1.Cell(table1.Rows.Count - 2, 1).Height
        shp.OLEFormat.Object.Width = table1.Cell(id + 1, 1).Width - 20
        'shp.OLEFormat.Object.ColumnCount = 3
        shp.OLEFormat.Object.MultiSelect = 1
        shp.OLEFormat.Object.ListStyle = fmListStyleOption
        'shp.OLEFormat.Object.ColumnHeads = True
Else
Set shp = table1.Rows(id + 1).Range.InlineShapes(1)
varuid1 = "___" & (Split(shp.OLEFormat.Object.Name, "___")(1)) & "____"
varreplaceid = Replace(Replace(Replace(varuid1, "___", "{"), "{_", "}"), "__", "-")


On Error Resume Next
ActiveDocument.CustomXMLParts.SelectByID(varreplaceid).Delete
If Err.Number = 91 Then
Err.Clear
On Error GoTo 0
End If
shp.OLEFormat.Object.Clear

Set metadata = config(ActiveDocument.Path & "\config.ini")
JSONTEXT = webservices(metadata.Item(2), metadata.Item(1), metadata.Item(3), metadata.Item(4), searchtext, metadatatype)
varuid = customxmlcreation(JSONTEXT, 1, qtype)
varuid = Replace(Replace(varuid, "{", "___"), "}", "____")
varuid = Replace(varuid, "-", "__")
varreplaceid = Replace(Replace(Replace(varuid, "___", "{"), "{_", "}"), "__", "-")
End If
    shp.OLEFormat.Object.Name = Replace(objname, "_cmd", "_lst") & varuid
For Each aa In ActiveDocument.CustomXMLParts.SelectByID(varreplaceid).SelectNodes("//metadata[@nodepath]")
        shp.OLEFormat.Object.AddItem aa.Text
Next
scode = scode & "Private Sub " & shp.OLEFormat.Object.Name & "_Change()" & vbCrLf
 'If varlistchange = False Then
    scode = scode & "If varlistchange = False Then" & vbCrLf
    scode = scode & "For varcount = 0 To " & shp.OLEFormat.Object.Name & ".ListCount - 1" & vbCrLf
    scode = scode & "ActiveDocument.CustomXMLParts.SelectByID(""" & varreplaceid & """).SelectSingleNode(""//metadata[@nodepath='"" & " & shp.OLEFormat.Object.Name & ".List(varcount) & ""']"").ParentNode.Attributes(1).NodeValue = " & shp.OLEFormat.Object.Name & ".Selected(varcount)" & vbCrLf
    scode = scode & "Next" & vbCrLf
    scode = scode & "end if" & vbCrLf
    scode = scode & "End Sub"
            'scode = scode & "Private Sub " & shp.OLEFormat.Object.Name & "_Change()" & vbCrLf
            
            
           ' scode = scode & "End Sub"
      '  MsgBox shp.OLEFormat.Object.Name
Set shp = Nothing
'MsgBox scode

ActiveDocument.VBProject.VBComponents("ThisDocument").CodeModule.AddFromString scode


End Sub
Private Function config(configfile) As Collection

Set txt = FSO.OpenTextFile(configfile)
vartext = txt.ReadAll
txt.Close
Dim base64 As New base64
varconfig = base64.decodeText(vartext)
clientcode = Trim(Split(Split(varconfig, Chr(13))(0), "=")(1))
secretKey = Trim(Split(Split(varconfig, Chr(13))(1), "=")(1))
metadata = Trim(Split(Split(varconfig, Chr(13))(2), "=")(1))
sitedata = Trim(Split(Split(varconfig, Chr(13))(3), "=")(1))
Set config = New Collection
config.Add secretKey, "secrekey"
config.Add clientcode, "clientcode"
config.Add metadata, "metadata"
config.Add sitedata, "sitedata"
'MsgBox webservices(clientcode, secretKey, metadata, sitedata)

End Function

Function webservices(clientcode, secretKey, metadata, sitedata, search, metadatatype) As String
Dim JSONTEXT As String
Dim objHttp As Object
Set objHttp = CreateObject("Msxml2.ServerXMLHTTP.6.0")
varData = "{""clientCode"" : """ & clientcode & """ , ""secretKey"" : """ & secretKey & """ , ""email"" : ""abdul.rahman@impelsys.com"", ""firstName"" : ""Abdul"", ""lastName"" : ""Rahman"", ""clientUserId"" : ""101""}"
Call objHttp.Open("POST", sitedata & "/api/authenticate", False)
Call objHttp.setRequestHeader("Content-Type", "application/json")
Call objHttp.setRequestHeader("Accept", "application/json")
Call objHttp.send(varData)
''Response.AddHeader "Content-Type", "application/json;charset=UTF-8"
''Response.Charset = "UTF-8"
varKey = Replace(Replace(objHttp.responseText, "{""token"":""", ""), """}", "")
'Call objHttp.Open("GET", sitedata & "/api/products/110000/metadata/" & metadata & metadatatype, False)
'/api/taxonomy-search?taxonomyName={search_string}&metadataType=HIERARCHY&metadataId={metadataId}
Call objHttp.Open("GET", sitedata & "/api/metadata-search?taxonomyName=" & search & "&metadataType=" & metadatatype & "&metadataId=" & metadata, False)
'Call objHttp.Open("GET", "http://qa-quizzingplatform.impelsys.com/api/taxonomy-search?taxonomyName=hie&metadataType=HIERARCHY&metadataId=SC_03", False)
'Call objHttp.Open("GET", "http://qa-quizzingplatform.impelsys.com/api/taxonomy-search", False)
'http://qa-quizzingplatform.impelsys.com/api/taxonomy-search
Call objHttp.setRequestHeader("Authorization", varKey)
Call objHttp.send("")
JSONTEXT = objHttp.responseText
'JsonText = "{""totalMetadata"":5,""metadataType"":""HIERARCHY"",""metadata"":[{""id"":1,""name"":""HLLibrary"",""nodePath"":""HLLibrary""},{""id"":2,""name"":""HealthProfessional"",""nodePath"":""HLLibrary\HealthProfessional""},{""id"":3,""name"":""medicine"",""nodePath"":""HLLibrary\HealthProfessional\medicine""},{""id"":5,""name"":""surgery"",""nodePath"":""HLLibrary\HealthProfessional\medicine\surgery""},{""id"":2,""name"":""Medicaleducation"",""nodePath"":""HLLibrary\Medicaleducation""},{""id"":3,""name"":""ENT"",""nodePath"":""HLLibrary\Medicaleducation\ENT""},{""id"":4,""name"":""surgery"",""nodePath"":""HLLibrary\Medicaleducation\ENT\surgery""},{""id"":6,""name"":""Nursing"",""nodePath"":""HLLibrary\Nursing""},{""id"":8,""name"":""Diabetic"",""nodePath"":""HLLibrary\Nursing\Diabetic""},{""id"":9,""name"":""surgery"",""nodePath"":""HLLibrary\Nursing\Diabetic\surgery""}]}"

JSONTEXT = "{""totalMetadata"":4,""metadata"":[{""metadataId"":""465"",""metadataType"":""LOOKUP"",""status"":""Active"",""taxonomyId"":""149"",""taxonomyName"":""Medical Education"",""taxonomyPath"":""Medical Education""},{""metadataId"":""465"",""metadataType"":""LOOKUP"",""status"":""Active"",""taxonomyId"":""148"",""taxonomyName"":""Emergency Medicine"",""taxonomyPath"":""Emergency Medicine""},{""metadataId"":""92"",""metadataType"":""LOOKUP"",""status"":""Active"",""taxonomyId"":""127"",""taxonomyName"":""medlab lookup2"",""taxonomyPath"":""medlab lookup2""},{""metadataId"":""92"",""metadataType"":""LOOKUP"",""status"":""Active"",""taxonomyId"":""128"",""taxonomyName"":""medlab lookup3"",""taxonomyPath"":""medlab lookup3""}]}"
webservices = JSONTEXT
End Function
Function customxmlcreation(JSONTEXT As String, qno, qtype) As String
Dim JSON As Object
Set JSON = JSONParser.ParseJson(Replace(JSONTEXT, "\", "\\"))
vari = 0
Dim varxml As String
varxml = "<metadataroot qno=""" & qno & """ qtype=""" & qtype & """>"
For Each Value In JSON("metadata")

varxml = varxml & "<metanode1 metadataid=""" & Value("metadataId") & """>"
varxml = varxml & "<metanode selected=""false"">"
varxml = varxml & "<metadata nodepath=""" & Value("taxonomyPath") & """>" & Value("taxonomyPath") & "</metadata>"
varxml = varxml & "<metadata name=""" & Value("taxonomyName") & """>" & Value("taxonomyName") & "</metadata>"
varxml = varxml & "<metadata id=""" & Value("taxonomyId") & """>" & Value("taxonomyId") & "</metadata>"
varxml = varxml & "<metadata metadataType=""" & Value("metadataType") & """>" & Value("metadataType") & "</metadata>"
varxml = varxml & "</metanode>"
varxml = varxml & "</metanode1>"

   ' Medata_form.ListBox1.AddItem
   ' Medata_form.ListBox1.List(vari, 0) = Value("taxonomyPath")
  '  Medata_form.ListBox1.List(vari, 1) = Value("taxonomyName")
  '  Medata_form.ListBox1.List(vari, 2) = Value("taxonomyId")
   ' vari = vari + 1
Next
varxml = varxml & "</metadataroot>"

customxmlcreation = ActiveDocument.CustomXMLParts.Add(varxml).id

End Function
Sub HeadingButtoncode(objname As String, table1 As Table, varflag As Boolean, qno, qtype)

Dim tt As ContentControl

Select Case (Split(objname, "_")(0))
Case "Rem"
            If varflag = True Then
            table1.Cell(table1.Rows.Count - 2, 1).Select
            Set table21 = Selection.Tables.Add(Selection.Range, 1, 1)
            table21.id = "Remidation"
            table21.Range.id = "Remidation"
            table21.Rows(table21.Rows.Count).Select
            table21.Rows(table21.Rows.Count).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
            table21.Cell(table21.Rows.Count, 1).Select
            Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
            tt.Range.Style = "Remediation_Link_Type_head"
            tt.SetPlaceholderText Text:="Remediation Type"
            ''tt.Range.Style = "Normal"
            tt.Range.Font.Size = "10"
            tt.Range.Font.ColorIndex = wdBlack
            tt.Range.Font.Bold = True
            tt.Range.Rows.Height = 30
            tt.Range.Font.Name = "Verdana"
            tt.LockContentControl = True
            tt.LockContents = True
            Set tt = Nothing
            
            table21.Cell(table21.Rows.Count, 2).Select
             Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
             tt.Range.Style = "Remediation_head"
            tt.SetPlaceholderText Text:="Remediation Detail"
            tt.Range.Style = "Normal"
            tt.Range.Font.Size = "10"
            tt.Range.Font.ColorIndex = wdBlack
            tt.Range.Font.Bold = True
            'tt.Range.Rows.Height = 30
            tt.Range.Font.Name = "Verdana"
            tt.LockContentControl = True
            tt.LockContents = True
            Set tt = Nothing
            table21.Rows.Add
            End If
            scode = scode & "Private Sub " & objname & "_Click()" & vbCrLf
            scode = scode & " call remfun(""" & objname & """,Selection.Tables(1).Cell(" & table1.Rows.Count - 2 & ", 1).Tables(1)," & table1.Rows.Count - 2 & "," & qno & ",""" & qtype & """)" & vbCrLf
             scode = scode & "End Sub"
Case "MD"
             table1.Cell(table1.Rows.Count, 1).Select
             Set table21 = Selection.Tables.Add(Selection.Range, 1, 1)
                table21.id = "Metadata"
                table21.Range.id = "Metadata"
            scode = scode & "Private Sub " & objname & "_Click()" & vbCrLf
            scode = scode & " call remfun(""" & objname & """,selection.tables(1).Cell(" & table1.Rows.Count & ", 1).tables(1)," & table1.Rows.Count & "," & qno & ",""" & qtype & """)" & vbCrLf
             scode = scode & "End Sub"
'             MsgBox scode
End Select

ActiveDocument.VBProject.VBComponents("ThisDocument").CodeModule.AddFromString scode
End Sub
