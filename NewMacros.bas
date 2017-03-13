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
    
    Selection.HomeKey Unit:=wdStory
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
    Selection.HomeKey Unit:=wdStory
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
    Selection.HomeKey Unit:=wdStory
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
                    Selection.EndOf Unit:=wdParagraph, Extend:=wdExtend
                    Selection.EndKey
                    Selection.Text = "</li></ul>"
                    Selection.EndKey Unit:=wdLine
                End If
            Else
            If InStr(oPara.Range.Previous(wdParagraph, 1).Text, "<li>") Then
              '  r.InsertBefore Text:="<lib>"
                    If oPara.Range.Next(wdParagraph, 1).ListFormat.ListType = wdListBullet Then
                    r.InsertBefore Text:="<li>"
                    Selection.EndOf Unit:=wdParagraph, Extend:=wdExtend
                    Selection.EndKey
                    Selection.Text = "</li>"
                    Selection.EndKey Unit:=wdLine
                    Else
'                     r.InsertBefore Text:="<li>"
'                    Selection.EndKey Unit:=wdLine
'
'                    Selection.Text = "</li></ul"
                    r.Select
                    r.InsertBefore Text:="<li>"
                    Selection.EndOf Unit:=wdParagraph, Extend:=wdExtend
                    Selection.EndKey
                    Selection.Text = "</li></ul>"
                    Selection.EndKey Unit:=wdLine
                    
                  '  r.InsertBefore Text:="</ul><li>"
                End If
                Else
                
                r.InsertBefore Text:="<ul><li>"
                Selection.EndOf Unit:=wdParagraph, Extend:=wdExtend
                    Selection.EndKey
                    If oPara.Range.Next(wdParagraph, 1).ListFormat.ListType = wdListBullet Then
                    Selection.Text = "</li>"
                    Else
                     Selection.Text = "</li></ul>"
                     End If
                Selection.EndKey Unit:=wdLine
                
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
                    Selection.EndOf Unit:=wdParagraph, Extend:=wdExtend
                    Selection.EndKey
                    Selection.Text = "</li></ol>"
                    Selection.EndKey Unit:=wdLine
                End If
            Else
            If InStr(oPara.Range.Previous(wdParagraph, 1).Text, "<li>") Then
              '  r.InsertBefore Text:="<lib>"
                    If oPara.Range.Next(wdParagraph, 1).ListFormat.ListType = wdListSimpleNumbering Then
                    r.InsertBefore Text:="<li>"
                    Selection.EndOf Unit:=wdParagraph, Extend:=wdExtend
                    Selection.EndKey
                    Selection.Text = "</li>"
                    Selection.EndKey Unit:=wdLine
                    Else
'                     r.InsertBefore Text:="<li>"
'                    Selection.EndKey Unit:=wdLine
'
'                    Selection.Text = "</li></ul"
                    r.Select
                    r.InsertBefore Text:="<li>"
                    Selection.EndOf Unit:=wdParagraph, Extend:=wdExtend
                    Selection.EndKey
                    Selection.Text = "</li></ol>"
                    Selection.EndKey Unit:=wdLine
                    
                  '  r.InsertBefore Text:="</ul><li>"
                End If
                Else
                
                r.InsertBefore Text:="<ol><li>"
                Selection.EndOf Unit:=wdParagraph, Extend:=wdExtend
                    Selection.EndKey
                    If oPara.Range.Next(wdParagraph, 1).ListFormat.ListType = wdListSimpleNumbering Then
                    Selection.Text = "</li>"
                    Else
                     Selection.Text = "</li></ol>"
                     startno = startno + 1
                     End If
                Selection.EndKey Unit:=wdLine
                
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

Sub XMLprocess(varpathout)
Dim tblOne As Table
Dim para As Paragraph
QuestionNo = 1

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
            questiontypename = "True False"
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
                                 
                                
                                If questiontype = "MC" And varSpecialType = False Then
                                        varSpecialType = False
                                        If para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Type" Then
                                            Print #99, "    <wk_question identificationId=""0"" qtype=""" & questiontype & """ qmode=""C"">"
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
                                            Print #99, "        <question_remediation_link qrlmode=""C"">"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Text" Then
                                            Print #99, "            <remediation_type ucx=""C"" redLinkId="""" remediation_link_type=""Web-link"">"
                                            Print #99, "                <remediation_type_text>" & varcontent1 & "</remediation_type_text>"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Link" Then
                                            Print #99, "                <remediation_type_link>" & varcontent1 & "</remediation_type_link>"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Tooltip" Then
                                            Print #99, "                <remediation_type_tooltip>" & varcontent1 & "</remediation_type_tooltip>"
                                            Print #99, "            </remediation_type>"
                                            Print #99, "             <remediation_type ucx=""C"" redLinkId="""" remediation_link_type=""Ebook""></remediation_type>"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Text_Remediation_Text" Then
                                            Print #99, "             <remediation_type ucx=""C"" redLinkId="""" remediation_link_type=""Text"">"
                                            Print #99, "                <remediation_type_text>" & varcontent1 & "</remediation_type_text>"
                                            Print #99, "            </remediation_type>"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Meta_Data_Attributes_head" Then
                                            Print #99, "        </question_remediation_link>"
                                            Print #99, "        <question_meta_tag qmtmode=""C"">"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_TAG" Then
                                            Print #99, "            <meta_tag ucx=""C"" metaTagId="""">"
                                            Print #99, "                <meta_tag_type>" & varcontent1 & "</meta_tag_type>"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_VALUE" Then
                                            Print #99, "                <meta_tag_value>" & varcontent1 & "</meta_tag_value>"
                                            Print #99, "            </meta_tag>"
                                        End If
                                 ElseIf questiontype = "MC" And varSpecialType = True Then
                                        If para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Type" Then
                                            Print #99, "        <cs_sub_question identificationId=""0"" qtype=""" & questiontype & """ qmode=""C"">"
                                            'Print #99, "            <question_type ucx=""C"" >" & questiontypename & "</question_type>"
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
                                            Print #99, "            <cs_sub_question_remediation_link qrlmode=""C"">"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Text" Then
                                            Print #99, "                <cs_sub_remediation_type ucx=""C"" redLinkId="""" remediation_link_type=""Web-link"">"
                                            Print #99, "                    <cs_sub_remediation_type_text>" & varcontent1 & "</cs_sub_remediation_type_text>"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Link" Then
                                            Print #99, "                    <cs_sub_remediation_type_link>" & varcontent1 & "</cs_sub_remediation_type_link>"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Tooltip" Then
                                            Print #99, "                    <cs_sub_remediation_type_tooltip>" & varcontent1 & "</cs_sub_remediation_type_tooltip>"
                                            Print #99, "                </cs_sub_remediation_type>"
                                            Print #99, "                 <cs_sub_remediation_type ucx=""C"" redLinkId="""" remediation_link_type=""Ebook""></cs_sub_remediation_type>"
                                        ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Text_Remediation_Text" Then
                                            Print #99, "                 <cs_sub_remediation_type ucx=""C"" redLinkId="""" remediation_link_type=""Text"">"
                                            Print #99, "                    <cs_sub_remediation_type_text>" & varcontent1 & "</cs_sub_remediation_type_text>"
                                            Print #99, "                </cs_sub_remediation_type>"
                                            Print #99, "            </cs_sub_question_remediation_link>"
                                            Print #99, "            </cs_sub_question>"
                                        End If
                                ElseIf questiontype = "CM" And varSpecialType = False Then
                                    varSpecialType = False
                                    If para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Type" Then
                                            Print #99, "    <wk_question identificationId=""0"" qtype=""" & questiontype & """ qmode=""C"">"
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
                                        For K = 1 To para.Range.Tables(vartabl).Rows.Count - 22
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
                                        Print #99, "        <question_remediation_link qrlmode=""C"">"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Text" Then
                                        Print #99, "            <remediation_type ucx=""C"" redLinkId="""" remediation_link_type=""Web-link"">"
                                        Print #99, "                <remediation_type_text>" & varcontent1 & "</remediation_type_text>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Link" Then
                                        Print #99, "                <remediation_type_link>" & varcontent1 & "</remediation_type_link>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Tooltip" Then
                                        Print #99, "                <remediation_type_tooltip>" & varcontent1 & "</remediation_type_tooltip>"
                                        Print #99, "            </remediation_type>"
                                        Print #99, "             <remediation_type ucx=""C"" redLinkId="""" remediation_link_type=""Ebook""></remediation_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Text_Remediation_Text" Then
                                        Print #99, "             <remediation_type ucx=""C"" redLinkId="""" remediation_link_type=""Text"">"
                                        Print #99, "                <remediation_type_text>" & varcontent1 & "</remediation_type_text>"
                                        Print #99, "            </remediation_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Meta_Data_Attributes_head" Then
                                        Print #99, "        </question_remediation_link>"
                                        Print #99, "        <question_meta_tag qmtmode=""C"">"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_TAG" Then
                                        Print #99, "            <meta_tag ucx=""C"" metaTagId="""">"
                                        Print #99, "                <meta_tag_type>" & varcontent1 & "</meta_tag_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_VALUE" Then
                                        Print #99, "                <meta_tag_value>" & varcontent1 & "</meta_tag_value>"
                                        Print #99, "            </meta_tag>"
                                    End If
                                ElseIf questiontype = "CM" And varSpecialType = True Then
                                    If para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Type" Then
                                            Print #99, "        <cs_sub_question identificationId=""0"" qtype=""" & questiontype & """ qmode=""C"">"
                                            'Print #99, "            <question_type ucx=""C"" >" & questiontypename & "</question_type>"
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
                                        For K = 1 To para.Range.Tables(vartabl).Rows.Count - 22
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
                                        Print #99, "            <cs_sub_question_remediation_link qrlmode=""C"">"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Text" Then
                                        Print #99, "                <cs_sub_remediation_type ucx=""C"" redLinkId="""" remediation_link_type=""Web-link"">"
                                        Print #99, "                    <cs_sub_remediation_type_text>" & varcontent1 & "</cs_sub_remediation_type_text>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Link" Then
                                        Print #99, "                    <cs_sub_remediation_type_link>" & varcontent1 & "</cs_sub_remediation_type_link>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Tooltip" Then
                                        Print #99, "                    <cs_sub_remediation_type_tooltip>" & varcontent1 & "</cs_sub_remediation_type_tooltip>"
                                        Print #99, "                </cs_sub_remediation_type>"
                                        Print #99, "                 <cs_sub_remediation_type ucx=""C"" redLinkId="""" remediation_link_type=""Ebook""></cs_sub_remediation_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Text_Remediation_Text" Then
                                        Print #99, "                 <cs_sub_remediation_type ucx=""C"" redLinkId="""" remediation_link_type=""Text"">"
                                        Print #99, "                    <cs_sub_remediation_type_text>" & varcontent1 & "</cs_sub_remediation_type_text>"
                                        Print #99, "                </cs_sub_remediation_type>"
                                        Print #99, "            </cs_sub_question_remediation_link>"
                                        Print #99, "            </cs_sub_question>"
                                    End If
                                ElseIf (questiontype = "TF") Then
                                    varSpecialType = False
                                    If para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Type" Then
                                            Print #99, "    <wk_question identificationId=""0"" qtype=""" & questiontype & """ qmode=""C"">"
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
                                        Print #99, "        <question_remediation_link qrlmode=""C"">"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Text" Then
                                        Print #99, "            <remediation_type ucx=""C"" redLinkId="""" remediation_link_type=""Web-link"">"
                                        Print #99, "                <remediation_type_text>" & varcontent1 & "</remediation_type_text>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Link" Then
                                        Print #99, "                <remediation_type_link>" & varcontent1 & "</remediation_type_link>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Tooltip" Then
                                        Print #99, "                <remediation_type_tooltip>" & varcontent1 & "</remediation_type_tooltip>"
                                        Print #99, "            </remediation_type>"
                                        Print #99, "             <remediation_type ucx=""C"" redLinkId="""" remediation_link_type=""Ebook""></remediation_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Text_Remediation_Text" Then
                                        Print #99, "             <remediation_type ucx=""C"" redLinkId="""" remediation_link_type=""Text"">"
                                        Print #99, "                <remediation_type_text>" & varcontent1 & "</remediation_type_text>"
                                        Print #99, "            </remediation_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Meta_Data_Attributes_head" Then
                                        Print #99, "        </question_remediation_link>"
                                        Print #99, "        <question_meta_tag qmtmode=""C"">"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_TAG" Then
                                        Print #99, "            <meta_tag ucx=""C"" metaTagId="""">"
                                        Print #99, "                <meta_tag_type>" & varcontent1 & "</meta_tag_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_VALUE" Then
                                        Print #99, "                <meta_tag_value>" & varcontent1 & "</meta_tag_value>"
                                        Print #99, "            </meta_tag>"
                                    End If
                                ElseIf (questiontype = "VQ") Then
                                    'On Error Resume Next
                                    varSpecialType = False
                                    If para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Type" Then
                                            Print #99, "    <wk_question identificationId=""0"" qtype=""" & questiontype & """ qmode=""C"">"
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
                                        Print #99, "        <question_additional_fields uck=""C"" referencevalue="""">"
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
                                        Print #99, "        <question_remediation_link qrlmode=""C"">"
                                   ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Text" Then
                                        Print #99, "            <remediation_type ucx=""C"" redLinkId="""" remediation_link_type=""Web-link"">"
                                        Print #99, "                <remediation_type_text>" & varcontent1 & "</remediation_type_text>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Link" Then
                                        Print #99, "                <remediation_type_link>" & varcontent1 & "</remediation_type_link>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Tooltip" Then
                                        Print #99, "                <remediation_type_tooltip>" & varcontent1 & "</remediation_type_tooltip>"
                                        Print #99, "            </remediation_type>"
                                        Print #99, "             <remediation_type ucx=""C"" redLinkId="""" remediation_link_type=""Ebook""></remediation_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Text_Remediation_Text" Then
                                        Print #99, "             <remediation_type ucx=""C"" redLinkId="""" remediation_link_type=""Text"">"
                                        Print #99, "                <remediation_type_text>" & varcontent1 & "</remediation_type_text>"
                                        Print #99, "            </remediation_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Meta_Data_Attributes_head" Then
                                        Print #99, "        </question_remediation_link>"
                                        Print #99, "        <question_meta_tag qmtmode=""C"">"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_TAG" Then
                                        Print #99, "            <meta_tag ucx=""C"" metaTagId="""">"
                                        Print #99, "                <meta_tag_type>" & varcontent1 & "</meta_tag_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_VALUE" Then
                                        Print #99, "                <meta_tag_value>" & varcontent1 & "</meta_tag_value>"
                                        Print #99, "            </meta_tag>"
                                    End If
                                
                                ElseIf (questiontype = "GO") Then
                                    'On Error Resume Next
                                    varSpecialType = False
                                    If para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Type" Then
                                            Print #99, "    <wk_question identificationId=""0"" qtype=""" & questiontype & """ qmode=""C"">"
                                            Print #99, "        <question_type ucx=""C"" >" & questiontypename & "</question_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Title" Then
                                        Print #99, "        <question_title ucx=""C"">" & varcontent1 & "</question_title>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Text" Then
                                        Print #99, "        <question_text ucx=""C"">" & varcontent1 & "</question_text>"
                                        Print #99, "        <question_graphic_option qmcmode=""C"" tagtype=""" & tagtype & """ >"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Additional_file_path" Then
                                        Print #99, "            <question_choice ucx=""C"" refId="""">"
                                        Print #99, "        <question_additional_fields uck=""C"" referencevalue="""">"
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
                                        Print #99, "        <question_remediation_link qrlmode=""C"">"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Text" Then
                                        Print #99, "            <remediation_type ucx=""C"" redLinkId="""" remediation_link_type=""Web-link"">"
                                        Print #99, "                <remediation_type_text>" & varcontent1 & "</remediation_type_text>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Link" Then
                                        Print #99, "                <remediation_type_link>" & varcontent1 & "</remediation_type_link>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Tooltip" Then
                                        Print #99, "                <remediation_type_tooltip>" & varcontent1 & "</remediation_type_tooltip>"
                                        Print #99, "            </remediation_type>"
                                        Print #99, "             <remediation_type ucx=""C"" redLinkId="""" remediation_link_type=""Ebook""></remediation_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Text_Remediation_Text" Then
                                        Print #99, "             <remediation_type ucx=""C"" redLinkId="""" remediation_link_type=""Text"">"
                                        Print #99, "                <remediation_type_text>" & varcontent1 & "</remediation_type_text>"
                                        Print #99, "            </remediation_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Meta_Data_Attributes_head" Then
                                        Print #99, "        </question_remediation_link>"
                                        Print #99, "        <question_meta_tag qmtmode=""C"">"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_TAG" Then
                                       
                                        
                                        Print #99, "            <meta_tag ucx=""C"" metaTagId="""">"
                                        Print #99, "                <meta_tag_type>" & varcontent1 & "</meta_tag_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_VALUE" Then
                                        Print #99, "                <meta_tag_value>" & varcontent1 & "</meta_tag_value>"
                                        Print #99, "            </meta_tag>"
                                    End If
                                
                                ElseIf (questiontype = "II") Then
                                    'On Error Resume Next
                                    varSpecialType = False
                                    If para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Type" Then
                                            Print #99, "    <wk_question identificationId=""0"" qtype=""" & questiontype & """ qmode=""C"">"
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
                                     
                                        
                                        'para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.InlineShapes(1).LinkFormat.SourceFullName
                                        Print #99, "        <question_additional_fields uck=""C"" referencevalue="""">"
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
                                        Print #99, "        <question_remediation_link qrlmode=""C"">"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Text" Then
                                        Print #99, "            <remediation_type ucx=""C"" redLinkId="""" remediation_link_type=""Web-link"">"
                                        Print #99, "                <remediation_type_text>" & varcontent1 & "</remediation_type_text>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Link" Then
                                        Print #99, "                <remediation_type_link>" & varcontent1 & "</remediation_type_link>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Web_Remediation_Tooltip" Then
                                        Print #99, "                <remediation_type_tooltip>" & varcontent1 & "</remediation_type_tooltip>"
                                        Print #99, "            </remediation_type>"
                                        Print #99, "             <remediation_type ucx=""C"" redLinkId="""" remediation_link_type=""Ebook""></remediation_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Text_Remediation_Text" Then
                                        Print #99, "             <remediation_type ucx=""C"" redLinkId="""" remediation_link_type=""Text"">"
                                        Print #99, "                <remediation_type_text>" & varcontent1 & "</remediation_type_text>"
                                        Print #99, "            </remediation_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Meta_Data_Attributes_head" Then
                                        Print #99, "        </question_remediation_link>"
                                        Print #99, "        <question_meta_tag qmtmode=""C"">"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_TAG" Then
                                        Print #99, "            <meta_tag ucx=""C"" metaTagId="""">"
                                        Print #99, "                <meta_tag_type>" & varcontent1 & "</meta_tag_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_VALUE" Then
                                        Print #99, "                <meta_tag_value>" & varcontent1 & "</meta_tag_value>"
                                        Print #99, "            </meta_tag>"
                                    End If
                                    
                                ElseIf (questiontype = "CS") Then
                                    'varCloseTag = True
                                    varSpecialType = True
                                    If para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Type" Then
                                            Print #99, "    <wk_question identificationId=""0"" qtype=""" & questiontype & """ qmode=""C"">"
                                            Print #99, "        <question_type ucx=""C"" >" & questiontypename & "</question_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Title" Then
                                        Print #99, "        <question_title ucx=""C"">" & varcontent1 & "</question_title>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Text" Then
                                        Print #99, "        <question_text ucx=""C"">" & varcontent1 & "</question_text>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Additional_file_path" Then
                                        Print #99, "        <question_additional_fields ucx=""C"">" & varcontent1 & "</question_additional_fields>"
                                        If FSO.FolderExists(ActiveDocument.Path & "\media") = False Then
                                            Print #35, vbCrLf
                                            varmsgtext = varerrorno & ". " & varcontent1 & " - This file is not available in the media folder"
                                            Print #35, varmsgtext
                                            varerrorno = varerrorno + 1
                                        End If
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Meta_Data_Attributes_head" Then
                                        Print #99, "        <question_meta_tag qmtmode=""C"">"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_TAG" Then
                                        Print #99, "            <meta_tag ucx=""C"" metaTagId="""">"
                                        Print #99, "                <meta_tag_type>" & varcontent1 & "</meta_tag_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_VALUE" Then
                                        Print #99, "                <meta_tag_value>" & varcontent1 & "</meta_tag_value>"
                                        Print #99, "            </meta_tag>"
                                        'Print #99, "        </question_meta_tag>"
                                        'Print #99, "        <question_cs_sub_questions>"
                                    End If
                               ElseIf (questiontype = "MEDC") Then
                                    'On Error Resume Next
                                    'varCloseTag = True
                                    varSpecialType = True
                                    If para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Type" Then
                                            'Print #99, "        </question_cs_sub_questions>"
                                            'Print #99, "    </wk_question>"
                                            Print #99, "    <wk_question identificationId=""0"" qtype=""" & questiontype & """ qmode=""C"">"
                                            Print #99, "        <question_type ucx=""C"" >" & questiontypename & "</question_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Title" Then
                                        Print #99, "        <question_title ucx=""C"">" & varcontent1 & "</question_title>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Question_Text" Then
                                        Print #99, "        <question_text ucx=""C"">" & varcontent1 & "</question_text>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Additional_file_path" Then
                                        Print #99, "        <question_additional_fields ucx=""C"">" & varcontent1 & "</question_additional_fields>"
                                        If FSO.FolderExists(ActiveDocument.Path & "\media") = False Then
                                            Print #35, vbCrLf
                                            varmsgtext = varerrorno & ". " & varcontent1 & " - This file is not available in the media folder"
                                            Print #35, varmsgtext
                                            varerrorno = varerrorno + 1
                                        End If
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "Meta_Data_Attributes_head" Then
                                        Print #99, "        <question_meta_tag qmtmode=""C"">"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_TAG" Then
                                        Print #99, "            <meta_tag ucx=""C"" metaTagId="""">"
                                        Print #99, "                <meta_tag_type>" & varcontent1 & "</meta_tag_type>"
                                    ElseIf para.Range.Tables(vartabl).Rows(varrow).Cells(varcol).Range.Style = "METADATA_VALUE" Then
                                        Print #99, "                <meta_tag_value>" & varcontent1 & "</meta_tag_value>"
                                        Print #99, "            </meta_tag>"
                                        
                                        'Print #99, "        </question_meta_tag>"
                                        'Print #99, "        <question_cs_sub_questions>"
                                    End If
                                    
                                End If
                         Next
                     Next
                'Print #99, "        <question_" & LCase(questiontype) & "_version>1</question_" & LCase(questiontype) & "_version>"
                'Print #99, "        <question_" & LCase(questiontype) & "_status>Create</question_" & LCase(questiontype) & "_status>"
                If questiontype = "MEDC" Or questiontype = "CS" Then
                Print #99, "        </question_meta_tag>"
                Print #99, "        <question_cs_sub_questions>"
                End If
                If varSpecialType = False Then
                    Print #99, "        </question_meta_tag>"
                    Print #99, "    </wk_question>"
                End If
                'If varCloseTag = True And questiontype = "CM" Then
                    'Print #99, "        </question_cs_sub_questions>"
                    'Print #99, "        </wk_question>"
               ' End If
                QuestionNo = QuestionNo + 1
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
                    
                    vartxt5 = Replace(vartxt5, "</cs_sub_question>" & vbNewLine & "    <wk_question", "</cs_sub_question>" & vbNewLine & "        </question_cs_sub_questions>" & vbNewLine & "        </wk_question>" & vbNewLine & "    <wk_question")
                    vartxt5 = Replace(vartxt5, "</cs_sub_question>" & vbNewLine & "</wk_question_root>", "</cs_sub_question>" & vbNewLine & "        </question_cs_sub_questions>" & vbNewLine & "        </wk_question>" & vbNewLine & "</wk_question_root>")
                    Open varpathout & "\allQuestions.xml" For Output As #55
                    Print #55, vartxt5
                    Close #55



Call ValidateFile(ActiveDocument.Path & "\allQuestions.xml")
End Sub
Sub ValidateFile(strFile)
    'Create an XML DOMDocument object.
    'MsgBox ("here")
    Dim x As New DOMDocument
    'Load and validate the specified file into the DOM.
    x.async = False
    x.validateOnParse = True
    x.resolveExternals = True
    x.Load strFile
    'Return validation results in message to the user.
    If x.parseError.ErrorCode <> 0 Then
        ValidateFile1 = "Validation failed on " & _
                       strFile & vbCrLf & _
                       "=====================" & vbCrLf & _
                       "Reason: " & x.parseError.reason & _
                       vbCrLf & "Source: " & _
                       x.parseError.srcText & _
                       vbCrLf & "Line: " & _
                       x.parseError.Line & vbCrLf
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
Dim table2 As Table
Dim addbutton
varNumRows = MyValue + 6
'QuestionNo = QuestionNo + 1
Selection.HomeKey Unit:=wdStory
Selection.EndKey Unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 16, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.ID = "MC"
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        table1.Rows(1).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(1).Select
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
        
        
        
        table1.Rows(2).Select
        
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
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(3, 1).Select
        
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
        
        table1.Cell(3, 2).Select
        
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
                
        table1.Rows(4).Select
        
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
        table1.Rows(5).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Answer_head"
        tt.SetPlaceholderText Text:="Answer Text - Multiple Choice"
        
        tt.Range.Font.Size = "13"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
    
         
        
      For i = 6 To MyValue + 5
        
       ' table1.Rows(5).SetHeight rowHeight:=ActiveDocument.PageSetup.PageHeight - 300, HeightRule:=wdRowHeightExactly
        
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 5
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
        tt.Title = "Rationale Text " & i - 5
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
               
      
               
        table1.Rows(table1.Rows.Count - 16).Select
        table1.Rows(table1.Rows.Count - 16).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 16, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 16, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 15).Select
        table1.Rows(table1.Rows.Count - 15).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 15, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 15, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 14).Select
        table1.Rows(table1.Rows.Count - 14).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 14, 1).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose score in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

        table1.Cell(table1.Rows.Count - 14, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose difficulty in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 13).Select
        table1.Rows(table1.Rows.Count - 13).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 13, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 13, 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 12).Select
        table1.Rows(table1.Rows.Count - 12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 12, 1).Select
        
       
        
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
        
        table1.Cell(table1.Rows.Count - 12, 2).Select
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
        
        
  
        table1.Rows(table1.Rows.Count - 11).Shading.BackgroundPatternColor = RGB(255, 140, 0)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 11).Select
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
        
        
        table1.Rows(table1.Rows.Count - 10).Select
        table1.Rows(table1.Rows.Count - 10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 10, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 10, 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 9).Select
        table1.Rows(table1.Rows.Count - 9).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Cell(table1.Rows.Count - 9, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 9, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 8).Select
        table1.Rows(table1.Rows.Count - 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
               
        table1.Cell(table1.Rows.Count - 8, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 7).Select
        table1.Rows(table1.Rows.Count - 7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
                
        table1.Cell(table1.Rows.Count - 7, 2).Select
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
        tt.SetPlaceholderText Text:="Enter remediation text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        '''''''''Meta data Attributes'''''''''
        
        table1.Rows(table1.Rows.Count - 4).Shading.BackgroundPatternColor = RGB(169, 169, 169)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 4).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Meta Data Attributes"
        tt.Range.Style = "Meta_Data_Attributes_head"
        tt.Range.Font.Size = "15"
        tt.Range.Font.ColorIndex = wdWhite
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 3).Select
        table1.Rows(table1.Rows.Count - 3).Cells.Split NumColumns:=3, MergeBeforeSplit:=True
        table1.Cell(table1.Rows.Count - 3, 1).Select
        
        Set shp = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.ComboBox.1")
        shp.OLEFormat.Object.Name = Replace(shp.OLEFormat.Object.Name, "ComboBox", "Qno_" & QuestionNo & "_cbo_Metadata_Search_keyword")
       ' shp.OLEFormat.Object.Name = Replace(shp.OLEFormat.Object.Name, "Label", "lbl_Metadata_Search_keyword")
        Set shp = Nothing
        
        
        table1.Cell(table1.Rows.Count - 3, 2).Select
        
        Set shp = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.TextBox.1")
        shp.OLEFormat.Object.Name = Replace(shp.OLEFormat.Object.Name, "TextBox", "Qno_" & QuestionNo & "_txt_Metadata_Search_keyword")
        shp.OLEFormat.Object.Width = 150
         table1.Cell(table1.Rows.Count - 3, 2).Select
        'shp.OLEFormat.Object.Name = Replace(shp.OLEFormat.Object.Name, "TextBox", "txt_Metadata_Search_keyword")
        'MsgBox shp.OLEFormat.Object.Object.ListCount
        Set shp = Nothing
        
        table1.Cell(table1.Rows.Count - 3, 3).Select
        
        Set shp = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
        shp.OLEFormat.Object.Name = Replace(shp.OLEFormat.Object.Name, "CommandButton", "Qno_" & QuestionNo & "_cmd_Metadata_Search_keyword")
        With shp.OLEFormat.Object
            .Object.Caption = "Search"
            '.Object.Width = 100
            '.Object.BackColor = RGB(212, 215, 219)
            
            '.Object.Style = "Chosse_file"
        End With
        MsgBox shp.OLEFormat.Object.Name
        vba_code shp.OLEFormat.Object.Name
        Set shp = Nothing
        
        'table1.Rows(table1.Rows.Count - 3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        'table1.Cell(table1.Rows.Count - 3, 1).Select
        'Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        'tt.SetPlaceholderText Text:="METADATA TAG"
        'tt.Range.Style = "METADATA_TAG_head"
        'tt.Range.Font.Size = "10"
        'tt.Range.Font.ColorIndex = wdBlack
        'tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        'tt.Range.Font.Name = "Verdana"
        'tt.LockContentControl = True
        'tt.LockContents = True
        'Set tt = Nothing
        
        'table1.Cell(table1.Rows.Count - 3, 2).Select
        ' Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        'tt.SetPlaceholderText Text:="METADATA VALUE"
        'tt.Range.Style = "METADATA_VALUE_head"
        'tt.Range.Font.Size = "10"
        'tt.Range.Font.ColorIndex = wdBlack
        'tt.Range.Font.Bold = True
        'tt.Range.Rows.Height = 30
        'tt.Range.Font.Name = "Verdana"
        'tt.LockContentControl = True
        'tt.LockContents = True
        'Set tt = Nothing
        
        
        
        table1.Rows(table1.Rows.Count - 2).Select
         Set shp = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.ListBox.1")
         MsgBox shp.OLEFormat.Object.Name
        shp.OLEFormat.Object.Name = Replace(shp.OLEFormat.Object.Name, "ListBox", "Qno_" & QuestionNo & "_lst_Metadata_Search_keyword")
        'shp.OLEFormat.Object.Height = table1.Cell(table1.Rows.Count - 2, 1).Height
        shp.OLEFormat.Object.Width = table1.Cell(table1.Rows.Count - 2, 1).Width - 20
        shp.OLEFormat.Object.ColumnCount = 3
        shp.OLEFormat.Object.MultiSelect = fmMultiSelectSingle
        shp.OLEFormat.Object.ListStyle = fmListStyleOption
        shp.OLEFormat.Object.ColumnHeads = True

        MsgBox shp.OLEFormat.Object.Name
        Set shp = Nothing
        
        
        'table1.Rows(table1.Rows.Count - 2).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
       ' table1.Cell(table1.Rows.Count - 2, 1).Select
       
        'Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        'tt.LockContentControl = True
        'tt.Range.Style = "METADATA_TAG"
        'tt.Range.Rows.Height = 30
        'tt.SetPlaceholderText , , "Select your metadata tag"
            
         '   For i = 0 To varTotalMetadata - 1
         '       tt.DropdownListEntries.Add (varMetaDataName(i))
         '   Next
        
        'tt.DropdownListEntries.Add "Status"
        'tt.DropdownListEntries.Add "Keyword"
        'tt.DropdownListEntries.Add "Topic"
        'Set tt = Nothing
        
        
        'table1.Cell(table1.Rows.Count - 2, 2).Select
        'Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        'tt.Range.Style = "METADATA_VALUE"
        'tt.Range.Font.Size = "10"
        'tt.Range.Font.ColorIndex = wdBlack
        'tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        'tt.Range.Font.Name = "Verdana"
        'tt.LockContentControl = True
        'tt.LockContents = False
        'Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 1).Select
        table1.Rows(table1.Rows.Count - 1).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
       
        table1.Cell(table1.Rows.Count - 1, 1).Select
       
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.LockContentControl = True
        tt.Range.Style = "METADATA_TAG"
        tt.Range.Rows.Height = 30
        tt.SetPlaceholderText , , "Select your metadata tag"
            For i = 0 To varTotalMetadata - 1
                tt.DropdownListEntries.Add (varMetaDataName(i))
            Next
        'tt.DropdownListEntries.Add "Status"
        'tt.DropdownListEntries.Add "Keyword"
        'tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count - 1, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
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
        tt.SetPlaceholderText , , "Select your metadata tag"
            For i = 0 To varTotalMetadata - 1
                tt.DropdownListEntries.Add (varMetaDataName(i))
            Next
        'tt.DropdownListEntries.Add "Status"
        'tt.DropdownListEntries.Add "Keyword"
        'tt.DropdownListEntries.Add "Topic"
        Set tt = Nothing
        
        
        table1.Cell(table1.Rows.Count, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "METADATA_VALUE"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        

End Sub

Sub copytablecs(MyValue, QuestionNo)
Dim table1 As Table
Dim tt As ContentControl
Dim table2 As Table
Dim addbutton
'varNumRows = MyValue + 6
'QuestionNo = QuestionNo + 1
Selection.HomeKey Unit:=wdStory
Selection.EndKey Unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=10, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.ID = "MC"
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        table1.Rows(1).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(1).Select
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
        
        
        
        table1.Rows(2).Select
        
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
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(3, 1).Select
        
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
        
        table1.Cell(3, 2).Select
        
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
                
        table1.Rows(4).Select
        
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
      
        
        table1.Rows(5).Select
        table1.Rows(5).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        table1.Cell(5, 1).Select
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
        
        table1.Cell(5, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Additional_Fields_head"
        tt.SetPlaceholderText Text:=""
        tt.LockContentControl = True
        
        table1.Cell(5, 3).Select
        'Selection.InlineShapes.AddOLEControl ClassType:="Forms.CommandButton.1"
        Set shp = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
        
        'shp.Range.Style = "Chosse_file"
        With shp.OLEFormat.Object
            .Object.Caption = "Select File"
            '.Object.Style = "Chosse_file"
        End With
        varbtn = Replace(shp.OLEFormat.Object.Name, "CommandButton", "")
        Set shp = Nothing
        
        
        table1.Cell(5, 4).Select
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
        
        '''''''''Meta data Attributes'''''''''
        
        table1.Rows(table1.Rows.Count - 4).Shading.BackgroundPatternColor = RGB(169, 169, 169)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 4).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Meta Data Attributes"
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
        tt.SetPlaceholderText , , "Select your metadata tag"
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
        tt.SetPlaceholderText , , "Select your metadata tag"
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
        tt.SetPlaceholderText , , "Select your metadata tag"
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
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        

End Sub
Sub copytableCSMC(MyValue, QuestionNo)
Dim table1 As Table
Dim tt As ContentControl
Dim table2 As Table
Dim addbutton
varNumRows = MyValue + 6
'QuestionNo = QuestionNo + 1
Selection.HomeKey Unit:=wdStory
Selection.EndKey Unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 12, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.ID = "MC"
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        table1.Rows(1).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(1).Select
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
        
        
        
        table1.Rows(2).Select
        
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
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(3, 1).Select
        
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
        
        table1.Cell(3, 2).Select
        
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
                
        table1.Rows(4).Select
        
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
        table1.Rows(5).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Answer_head"
        tt.SetPlaceholderText Text:="Answer Text - Multiple Choice"
        
        tt.Range.Font.Size = "13"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
    
         
        
      For i = 6 To MyValue + 5
        
       ' table1.Rows(5).SetHeight rowHeight:=ActiveDocument.PageSetup.PageHeight - 300, HeightRule:=wdRowHeightExactly
        
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 5
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
        tt.Title = "Rationale Text " & i - 5
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
               
      
               
        table1.Rows(table1.Rows.Count - 12).Select
        table1.Rows(table1.Rows.Count - 12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 12, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 12, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 11).Select
        table1.Rows(table1.Rows.Count - 11).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 11, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 11, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 10).Select
        table1.Rows(table1.Rows.Count - 10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 10, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose score in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

        table1.Cell(table1.Rows.Count - 10, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose difficulty in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 9).Select
        table1.Rows(table1.Rows.Count - 9).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 9, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 9, 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 8).Select
        table1.Rows(table1.Rows.Count - 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 8, 1).Select
        
       
        
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
        
        table1.Cell(table1.Rows.Count - 8, 2).Select
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
        
        
  
        table1.Rows(table1.Rows.Count - 7).Shading.BackgroundPatternColor = RGB(255, 140, 0)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 7).Select
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
        
        
        table1.Rows(table1.Rows.Count - 6).Select
        table1.Rows(table1.Rows.Count - 6).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 6, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 6, 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 5).Select
        table1.Rows(table1.Rows.Count - 5).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Cell(table1.Rows.Count - 5, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 5, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 4).Select
        table1.Rows(table1.Rows.Count - 4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
               
        table1.Cell(table1.Rows.Count - 4, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 3).Select
        table1.Rows(table1.Rows.Count - 3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
                
        table1.Cell(table1.Rows.Count - 3, 2).Select
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
        
        
        
        table1.Rows(table1.Rows.Count - 2).Select
        'table1.Rows(table1.Rows.Count - 6).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Rows(table1.Rows.Count - 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 1).Select
        table1.Rows(table1.Rows.Count - 1).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        table1.Cell(table1.Rows.Count - 1, 1).Select
        
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
        
        
        table1.Cell(table1.Rows.Count - 1, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Text_Remediation_Text"
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
        
        
        '''''''''Meta data Attributes'''''''''
        
        
        
        

End Sub

Sub copytableCSCM(MyValue, QuestionNo)
Dim table1 As Table
Dim tt As ContentControl
Dim table2 As Table
Dim addbutton
varNumRows = MyValue + 6
'QuestionNo = QuestionNo + 1
Selection.HomeKey Unit:=wdStory
Selection.EndKey Unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 12, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        table1.Rows(1).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(1).Select
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
        
         table1.Rows(2).Select
        
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
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(3, 1).Select
        
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
        
        table1.Cell(3, 2).Select
        
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
        
        table1.Rows(4).Select
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
        table1.Rows(5).Select
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
        For i = 6 To MyValue + 5
            
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 5
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
        tt.Title = "Rationale Text " & i - 5
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
               
        table1.Rows(table1.Rows.Count - 12).Select
        table1.Rows(table1.Rows.Count - 12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 12, 1).Select
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
            table1.Cell(table1.Rows.Count - 12, i).Split NumColumns:=(MyValue * 2) + 1
            table1.Cell(table1.Rows.Count - 12, 2).Select
            Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
            tt.SetPlaceholderText Text:=" "
            tt.Range.Style = "Check_box_no"
            tt.Range.Font.Size = "10"
            tt.Range.Font.Name = "Verdana"
            tt.LockContentControl = True
            tt.LockContents = False
            
        End If
        
        If i Mod 2 > 0 Then
            table1.Cell(table1.Rows.Count - 12, i).Select
            
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
            table1.Cell(table1.Rows.Count - 12, i).Select
            Selection.Collapse Direction:=wdCollapseEnd
            Set tt = ActiveDocument.ContentControls.Add(wdContentControlText)
                
            tt.SetPlaceholderText Text:=j
            tt.Range.Style = "Check_box_no"
            Set tt = Nothing
            j = j + 1
        End If
        
        Next
        
        
        'Set tt = Nothing
        table1.Rows(table1.Rows.Count - 11).Select
        table1.Rows(table1.Rows.Count - 11).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 11, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 11, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 10).Select
        table1.Rows(table1.Rows.Count - 10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 10, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose score in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

    table1.Cell(table1.Rows.Count - 10, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose difficulty in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 9).Select
        table1.Rows(table1.Rows.Count - 9).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 9, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 9, 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 8).Select
        table1.Rows(table1.Rows.Count - 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 8, 1).Select
        
       
        
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
        
        table1.Cell(table1.Rows.Count - 8, 2).Select
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
        
        
  
        table1.Rows(table1.Rows.Count - 7).Shading.BackgroundPatternColor = RGB(255, 140, 0)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 7).Select
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
        
        
        table1.Rows(table1.Rows.Count - 6).Select
        table1.Rows(table1.Rows.Count - 6).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 6, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 6, 2).Select
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
        
        
        
        
        table1.Rows(table1.Rows.Count - 5).Select
        table1.Rows(table1.Rows.Count - 5).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Cell(table1.Rows.Count - 5, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 5, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 4).Select
        table1.Rows(table1.Rows.Count - 4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
               
        table1.Cell(table1.Rows.Count - 4, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 3).Select
        table1.Rows(table1.Rows.Count - 3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
                
        table1.Cell(table1.Rows.Count - 3, 2).Select
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
        
        
        
        table1.Rows(table1.Rows.Count - 2).Select
        'table1.Rows(table1.Rows.Count - 6).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Rows(table1.Rows.Count - 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 1).Select
        table1.Rows(table1.Rows.Count - 1).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        table1.Cell(table1.Rows.Count - 1, 1).Select
        
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
        
        
        table1.Cell(table1.Rows.Count - 1, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Text_Remediation_Text"
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
        
        '''''''''Meta data Attributes'''''''''
        
       
End Sub

Sub copytableMCase(MyValue, QuestionNo)
Dim table1 As Table
Dim tt As ContentControl
Dim table2 As Table
Dim addbutton
'varNumRows = MyValue + 6
'QuestionNo = QuestionNo + 1
Selection.HomeKey Unit:=wdStory
Selection.EndKey Unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=10, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.ID = "MC"
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        table1.Rows(1).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(1).Select
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
        
        
        
        table1.Rows(2).Select
        
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
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(3, 1).Select
        
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
        
        table1.Cell(3, 2).Select
        
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
                
        table1.Rows(4).Select
        
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
      
        
        table1.Rows(5).Select
        table1.Rows(5).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        table1.Cell(5, 1).Select
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
        
        table1.Cell(5, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Additional_Fields_head"
        tt.SetPlaceholderText Text:=""
        tt.LockContentControl = True
        
        table1.Cell(5, 3).Select
        'Selection.InlineShapes.AddOLEControl ClassType:="Forms.CommandButton.1"
        Set shp = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
        
        'shp.Range.Style = "Chosse_file"
        With shp.OLEFormat.Object
            .Object.Caption = "Select File"
            '.Object.Style = "Chosse_file"
        End With
        varbtn = Replace(shp.OLEFormat.Object.Name, "CommandButton", "")
        Set shp = Nothing
        
        
        table1.Cell(5, 4).Select
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
        
        '''''''''Meta data Attributes'''''''''
        
        table1.Rows(table1.Rows.Count - 4).Shading.BackgroundPatternColor = RGB(169, 169, 169)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 4).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Meta Data Attributes"
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
        tt.SetPlaceholderText , , "Select your metadata tag"
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
        tt.SetPlaceholderText , , "Select your metadata tag"
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
        tt.SetPlaceholderText , , "Select your metadata tag"
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
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        

End Sub
Sub copytableMCaseMC(MyValue, QuestionNo)
Dim table1 As Table
Dim tt As ContentControl
Dim table2 As Table
Dim addbutton
varNumRows = MyValue + 6
'QuestionNo = QuestionNo + 1
Selection.HomeKey Unit:=wdStory
Selection.EndKey Unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 12, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.ID = "MC"
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        table1.Rows(1).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(1).Select
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
        
        
        
        table1.Rows(2).Select
        
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
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(3, 1).Select
        
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
        
        table1.Cell(3, 2).Select
        
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
                
        table1.Rows(4).Select
        
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
        table1.Rows(5).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Answer_head"
        tt.SetPlaceholderText Text:="Answer Text - Multiple Choice"
        
        tt.Range.Font.Size = "13"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
    
         
        
      For i = 6 To MyValue + 5
        
       ' table1.Rows(5).SetHeight rowHeight:=ActiveDocument.PageSetup.PageHeight - 300, HeightRule:=wdRowHeightExactly
        
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 5
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
        tt.Title = "Rationale Text " & i - 5
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
               
      
               
        table1.Rows(table1.Rows.Count - 12).Select
        table1.Rows(table1.Rows.Count - 12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 12, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 12, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 11).Select
        table1.Rows(table1.Rows.Count - 11).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 11, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 11, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 10).Select
        table1.Rows(table1.Rows.Count - 10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 10, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose score in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

        table1.Cell(table1.Rows.Count - 10, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose difficulty in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 9).Select
        table1.Rows(table1.Rows.Count - 9).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 9, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 9, 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 8).Select
        table1.Rows(table1.Rows.Count - 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 8, 1).Select
        
       
        
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
        
        table1.Cell(table1.Rows.Count - 8, 2).Select
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
        
        
  
        table1.Rows(table1.Rows.Count - 7).Shading.BackgroundPatternColor = RGB(255, 140, 0)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 7).Select
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
        
        
        table1.Rows(table1.Rows.Count - 6).Select
        table1.Rows(table1.Rows.Count - 6).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 6, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 6, 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 5).Select
        table1.Rows(table1.Rows.Count - 5).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Cell(table1.Rows.Count - 5, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 5, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 4).Select
        table1.Rows(table1.Rows.Count - 4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
               
        table1.Cell(table1.Rows.Count - 4, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 3).Select
        table1.Rows(table1.Rows.Count - 3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
                
        table1.Cell(table1.Rows.Count - 3, 2).Select
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
        
        
        
          table1.Rows(table1.Rows.Count - 2).Select
        'table1.Rows(table1.Rows.Count - 6).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Rows(table1.Rows.Count - 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 1).Select
        table1.Rows(table1.Rows.Count - 1).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        table1.Cell(table1.Rows.Count - 1, 1).Select
        
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
        
        
        table1.Cell(table1.Rows.Count - 1, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Text_Remediation_Text"
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
        
        
        '''''''''Meta data Attributes'''''''''
        
        
        
        

End Sub

Sub copytableMCaseCM(MyValue, QuestionNo)
Dim table1 As Table
Dim tt As ContentControl
Dim table2 As Table
Dim addbutton
varNumRows = MyValue + 6
'QuestionNo = QuestionNo + 1
Selection.HomeKey Unit:=wdStory
Selection.EndKey Unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 12, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        table1.Rows(1).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(1).Select
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
        
         table1.Rows(2).Select
        
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
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(3, 1).Select
        
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
        
        table1.Cell(3, 2).Select
        
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
        
        table1.Rows(4).Select
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
        table1.Rows(5).Select
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
        For i = 6 To MyValue + 5
            
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 5
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
        tt.Title = "Rationale Text " & i - 5
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
               
        table1.Rows(table1.Rows.Count - 12).Select
        table1.Rows(table1.Rows.Count - 12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 12, 1).Select
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
            table1.Cell(table1.Rows.Count - 12, i).Split NumColumns:=(MyValue * 2) + 1
            table1.Cell(table1.Rows.Count - 12, 2).Select
            Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
            tt.SetPlaceholderText Text:=" "
            tt.Range.Style = "Check_box_no"
            tt.Range.Font.Size = "10"
            tt.Range.Font.Name = "Verdana"
            tt.LockContentControl = True
            tt.LockContents = False
            
        End If
        
        If i Mod 2 > 0 Then
            table1.Cell(table1.Rows.Count - 12, i).Select
            
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
            table1.Cell(table1.Rows.Count - 12, i).Select
            Selection.Collapse Direction:=wdCollapseEnd
            Set tt = ActiveDocument.ContentControls.Add(wdContentControlText)
                
            tt.SetPlaceholderText Text:=j
            tt.Range.Style = "Check_box_no"
            Set tt = Nothing
            j = j + 1
        End If
        
        Next
        
        
        'Set tt = Nothing
        table1.Rows(table1.Rows.Count - 11).Select
        table1.Rows(table1.Rows.Count - 11).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 11, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 11, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 10).Select
        table1.Rows(table1.Rows.Count - 10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 10, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose score in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

    table1.Cell(table1.Rows.Count - 10, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose difficulty in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 9).Select
        table1.Rows(table1.Rows.Count - 9).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 9, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 9, 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 8).Select
        table1.Rows(table1.Rows.Count - 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 8, 1).Select
        
       
        
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
        
        table1.Cell(table1.Rows.Count - 8, 2).Select
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
        
        
  
        table1.Rows(table1.Rows.Count - 7).Shading.BackgroundPatternColor = RGB(255, 140, 0)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 7).Select
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
        
        
        table1.Rows(table1.Rows.Count - 6).Select
        table1.Rows(table1.Rows.Count - 6).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 6, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 6, 2).Select
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
        
        
        
        
        table1.Rows(table1.Rows.Count - 5).Select
        table1.Rows(table1.Rows.Count - 5).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Cell(table1.Rows.Count - 5, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 5, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 4).Select
        table1.Rows(table1.Rows.Count - 4).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
               
        table1.Cell(table1.Rows.Count - 4, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 3).Select
        table1.Rows(table1.Rows.Count - 3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
                
        table1.Cell(table1.Rows.Count - 3, 2).Select
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
        
        
        
        table1.Rows(table1.Rows.Count - 2).Select
        'table1.Rows(table1.Rows.Count - 6).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Rows(table1.Rows.Count - 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 1).Select
        table1.Rows(table1.Rows.Count - 1).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        table1.Cell(table1.Rows.Count - 1, 1).Select
        
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
        
        
        table1.Cell(table1.Rows.Count - 1, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Text_Remediation_Text"
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
        
        '''''''''Meta data Attributes'''''''''
        
       
End Sub
Sub copytablecm(MyValue, QuestionNo)
Dim table1 As Table
Dim tt As ContentControl
Dim table2 As Table
Dim addbutton
varNumRows = MyValue + 6
'QuestionNo = QuestionNo + 1
Selection.HomeKey Unit:=wdStory
Selection.EndKey Unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 16, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        table1.Rows(1).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(1).Select
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
        
         table1.Rows(2).Select
        
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
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(3, 1).Select
        
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
        
        table1.Cell(3, 2).Select
        
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
        
        table1.Rows(4).Select
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
        table1.Rows(5).Select
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
        For i = 6 To MyValue + 5
            
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 5
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
        tt.Title = "Rationale Text " & i - 5
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
               
        table1.Rows(table1.Rows.Count - 16).Select
        table1.Rows(table1.Rows.Count - 16).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 16, 1).Select
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
            table1.Cell(table1.Rows.Count - 16, i).Split NumColumns:=(MyValue * 2) + 1
            table1.Cell(table1.Rows.Count - 16, 2).Select
            Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
            tt.SetPlaceholderText Text:=" "
            tt.Range.Style = "Check_box_no"
            tt.Range.Font.Size = "10"
            tt.Range.Font.Name = "Verdana"
            tt.LockContentControl = True
            tt.LockContents = False
            
        End If
        
        If i Mod 2 > 0 Then
            table1.Cell(table1.Rows.Count - 16, i).Select
            
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
            table1.Cell(table1.Rows.Count - 16, i).Select
            Selection.Collapse Direction:=wdCollapseEnd
            Set tt = ActiveDocument.ContentControls.Add(wdContentControlText)
                
            tt.SetPlaceholderText Text:=j
            tt.Range.Style = "Check_box_no"
            Set tt = Nothing
            j = j + 1
        End If
        
        Next
        
        
        'Set tt = Nothing
        table1.Rows(table1.Rows.Count - 15).Select
        table1.Rows(table1.Rows.Count - 15).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 15, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 15, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 14).Select
        table1.Rows(table1.Rows.Count - 14).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 14, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose score in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

    table1.Cell(table1.Rows.Count - 14, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose difficulty in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 13).Select
        table1.Rows(table1.Rows.Count - 13).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 13, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 13, 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 12).Select
        table1.Rows(table1.Rows.Count - 12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 12, 1).Select
        
       
        
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
        
        table1.Cell(table1.Rows.Count - 12, 2).Select
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
        
        
  
        table1.Rows(table1.Rows.Count - 11).Shading.BackgroundPatternColor = RGB(255, 140, 0)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 11).Select
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
        
        
        table1.Rows(table1.Rows.Count - 10).Select
        table1.Rows(table1.Rows.Count - 10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 10, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 10, 2).Select
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
        
        
        
        
        table1.Rows(table1.Rows.Count - 9).Select
        table1.Rows(table1.Rows.Count - 9).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Cell(table1.Rows.Count - 9, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 9, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 8).Select
        table1.Rows(table1.Rows.Count - 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
               
        table1.Cell(table1.Rows.Count - 8, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 7).Select
        table1.Rows(table1.Rows.Count - 7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
                
        table1.Cell(table1.Rows.Count - 7, 2).Select
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
        tt.SetPlaceholderText Text:="Enter remediation text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        '''''''''Meta data Attributes'''''''''
        
        table1.Rows(table1.Rows.Count - 4).Shading.BackgroundPatternColor = RGB(169, 169, 169)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 4).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Meta Data Attributes"
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
        tt.SetPlaceholderText , , "Select your metadata tag"
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
        tt.SetPlaceholderText , , "Select your metadata tag"
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
        tt.SetPlaceholderText , , "Select your metadata tag"
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
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
       
End Sub

Sub copytabletf(QuestionNo)
Dim table1 As Table
Dim tt As ContentControl
Dim table2 As Table
Dim addbutton
'QuestionNo = QuestionNo + 1
Selection.HomeKey Unit:=wdStory
Selection.EndKey Unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=8 + 16, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        table1.Rows(1).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(1).Select
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
        
         table1.Rows(2).Select
        
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
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(3, 1).Select
        
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
        
        table1.Cell(3, 2).Select
        
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
        
        table1.Rows(4).Select
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
        
        table1.Rows(5).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Answer_head"
        tt.SetPlaceholderText Text:="Answer Text - True or Fasle"
       
        tt.Range.Font.Size = "13"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        
        table1.Rows(6).Select
        table1.Rows(6).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(6, 1).Select
        
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
        
         table1.Cell(6, 2).Select
        
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
        
        table1.Rows(7).Select
        table1.Rows(7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(7, 1).Select
        
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
        
         table1.Cell(7, 2).Select
        
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
        
        
        table1.Rows(table1.Rows.Count - 16).Select
        table1.Rows(table1.Rows.Count - 16).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 16, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 16, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Title = "Correct Answer"
        tt.LockContentControl = True
        tt.Range.Style = "Correct_Answer"
        tt.SetPlaceholderText , , "Select your correct answer: True or False"
        tt.DropdownListEntries.Add ("True")
        tt.DropdownListEntries.Add ("False")
        
        'Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 15).Select
        table1.Rows(table1.Rows.Count - 15).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 15, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 15, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 14).Select
        table1.Rows(table1.Rows.Count - 14).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 14, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose score in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

    table1.Cell(table1.Rows.Count - 14, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose difficulty in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 13).Select
        table1.Rows(table1.Rows.Count - 13).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 13, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 13, 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 12).Select
        table1.Rows(table1.Rows.Count - 12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 12, 1).Select
        
       
        
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
        
        table1.Cell(table1.Rows.Count - 12, 2).Select
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
        
        
  
        table1.Rows(table1.Rows.Count - 11).Shading.BackgroundPatternColor = RGB(255, 140, 0)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 11).Select
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
        
        
        table1.Rows(table1.Rows.Count - 10).Select
        table1.Rows(table1.Rows.Count - 10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 10, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 10, 2).Select
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
        
        
        
        table1.Rows(table1.Rows.Count - 9).Select
        table1.Rows(table1.Rows.Count - 9).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Cell(table1.Rows.Count - 9, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 9, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 8).Select
        table1.Rows(table1.Rows.Count - 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
               
        table1.Cell(table1.Rows.Count - 8, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 7).Select
        table1.Rows(table1.Rows.Count - 7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
                
        table1.Cell(table1.Rows.Count - 7, 2).Select
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
        tt.SetPlaceholderText Text:="Enter remediation text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        '''''''''Meta data Attributes'''''''''
        
        table1.Rows(table1.Rows.Count - 4).Shading.BackgroundPatternColor = RGB(169, 169, 169)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 4).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Meta Data Attributes"
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
        tt.SetPlaceholderText , , "Select your metadata tag"
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
        tt.SetPlaceholderText , , "Select your metadata tag"
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
        tt.SetPlaceholderText , , "Select your metadata tag"
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
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
End Sub

Sub copytableVQ(MyValue, QuestionNo)
Dim table1 As Table
Dim tt As ContentControl
Dim table2 As Table
Dim addbutton
varNumRows = MyValue + 7
'QuestionNo = QuestionNo + 1
Selection.HomeKey Unit:=wdStory
Selection.EndKey Unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 16, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.ID = "VQ"
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        table1.Rows(1).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(1).Select
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
        
        table1.Rows(2).Select
        
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
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(3, 1).Select
        
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
        
        table1.Cell(3, 2).Select
        
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
                
        table1.Rows(4).Select
        
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
        table1.Rows(5).Select
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
    
                    
       
        
      For i = 6 To MyValue + 5
        
       ' table1.Rows(5).SetHeight rowHeight:=ActiveDocument.PageSetup.PageHeight - 300, HeightRule:=wdRowHeightExactly
        
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 5
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
        tt.Title = "Rationale Text " & i - 5
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
        tt.SetPlaceholderText , , "Select your correct answer"
        For i = 1 To MyValue
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
               
        table1.Rows(table1.Rows.Count - 16).Select
        table1.Rows(table1.Rows.Count - 16).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 16, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 16, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Additional_Fields_head"
        tt.SetPlaceholderText Text:=""
        tt.LockContentControl = True
        
        table1.Cell(table1.Rows.Count - 16, 3).Select
        'Selection.InlineShapes.AddOLEControl ClassType:="Forms.CommandButton.1"
        Set shp = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
        
        'shp.Range.Style = "Chosse_file"
        With shp.OLEFormat.Object
            .Object.Caption = "Select File"
            '.Object.Style = "Chosse_file"
        End With
        varbtn = Replace(shp.OLEFormat.Object.Name, "CommandButton", "")
        Set shp = Nothing
        
        
        table1.Cell(table1.Rows.Count - 16, 4).Select
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
        
        table1.Rows(table1.Rows.Count - 15).Select
        table1.Rows(table1.Rows.Count - 15).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 15, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 15, 2).Select
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
        
       table1.Rows(table1.Rows.Count - 14).Select
        table1.Rows(table1.Rows.Count - 14).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 14, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose score in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

    table1.Cell(table1.Rows.Count - 14, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose difficulty in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 13).Select
        table1.Rows(table1.Rows.Count - 13).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 13, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 13, 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 12).Select
        table1.Rows(table1.Rows.Count - 12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 12, 1).Select
        
       
        
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
        
        table1.Cell(table1.Rows.Count - 12, 2).Select
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
        
        
  
        table1.Rows(table1.Rows.Count - 11).Shading.BackgroundPatternColor = RGB(255, 140, 0)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 11).Select
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
        
        
        table1.Rows(table1.Rows.Count - 10).Select
        table1.Rows(table1.Rows.Count - 10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 10, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 10, 2).Select
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
        
        
        
        
       table1.Rows(table1.Rows.Count - 9).Select
        table1.Rows(table1.Rows.Count - 9).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Cell(table1.Rows.Count - 9, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 9, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 8).Select
        table1.Rows(table1.Rows.Count - 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
               
        table1.Cell(table1.Rows.Count - 8, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 7).Select
        table1.Rows(table1.Rows.Count - 7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
                
        table1.Cell(table1.Rows.Count - 7, 2).Select
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
        tt.SetPlaceholderText Text:="Enter remediation text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        '''''''''Meta data Attributes'''''''''
        
        table1.Rows(table1.Rows.Count - 4).Shading.BackgroundPatternColor = RGB(169, 169, 169)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 4).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Meta Data Attributes"
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
        tt.SetPlaceholderText , , "Select your metadata tag"
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
        tt.SetPlaceholderText , , "Select your metadata tag"
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
        tt.SetPlaceholderText , , "Select your metadata tag"
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
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
       
        
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
Dim table2 As Table
Dim addbutton
varNumRows = MyValue + 7
'QuestionNo = QuestionNo + 1
Selection.HomeKey Unit:=wdStory
Selection.EndKey Unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 16, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.ID = "MC"
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        table1.Rows(1).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(1).Select
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
        
        table1.Rows(2).Select
        
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
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(3, 1).Select
        
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
        
        table1.Cell(3, 2).Select
        
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
                
        table1.Rows(4).Select
        
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
        table1.Rows(5).Select
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
    
                    
       
        
      For i = 6 To MyValue + 5
        
       ' table1.Rows(5).SetHeight rowHeight:=ActiveDocument.PageSetup.PageHeight - 300, HeightRule:=wdRowHeightExactly
        
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 5
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
        tt.Title = "Rationale Text " & i - 5
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
        tt.SetPlaceholderText , , "Select your correct answer"
        For i = 1 To MyValue
        tt.DropdownListEntries.Add (i)
        Next
               
        table1.Rows(table1.Rows.Count - 16).Select
        table1.Rows(table1.Rows.Count - 16).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 16, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 16, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Additional_Fields_head"
        tt.SetPlaceholderText Text:=""
        tt.LockContentControl = True
        
        table1.Cell(table1.Rows.Count - 16, 3).Select
        'Selection.InlineShapes.AddOLEControl ClassType:="Forms.CommandButton.1"
        Set shp = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
        
        'shp.Range.Style = "Chosse_file"
        With shp.OLEFormat.Object
            .Object.Caption = "Select File"
            '.Object.Name = "Chosse_file"
        End With
        varbtn = Replace(shp.OLEFormat.Object.Name, "CommandButton", "")
        Set shp = Nothing
        
        table1.Cell(table1.Rows.Count - 16, 4).Select
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
        
         table1.Rows(table1.Rows.Count - 15).Select
        table1.Rows(table1.Rows.Count - 15).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 15, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 15, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 14).Select
        table1.Rows(table1.Rows.Count - 14).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 14, 1).Select
        
        
         table1.Rows(table1.Rows.Count - 14).Select
        table1.Rows(table1.Rows.Count - 14).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 14, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose score in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

        table1.Cell(table1.Rows.Count - 14, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose difficulty in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 13).Select
        table1.Rows(table1.Rows.Count - 13).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 13, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 13, 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 12).Select
        table1.Rows(table1.Rows.Count - 12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 12, 1).Select
        
       
        
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
        
        table1.Cell(table1.Rows.Count - 12, 2).Select
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
        
        
  
        table1.Rows(table1.Rows.Count - 11).Shading.BackgroundPatternColor = RGB(255, 140, 0)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 11).Select
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
        
        
        table1.Rows(table1.Rows.Count - 10).Select
        table1.Rows(table1.Rows.Count - 10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 10, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 10, 2).Select
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
        
        
        
        
        table1.Rows(table1.Rows.Count - 9).Select
        table1.Rows(table1.Rows.Count - 9).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Cell(table1.Rows.Count - 9, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 9, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 8).Select
        table1.Rows(table1.Rows.Count - 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
               
        table1.Cell(table1.Rows.Count - 8, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 7).Select
        table1.Rows(table1.Rows.Count - 7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
                
        table1.Cell(table1.Rows.Count - 7, 2).Select
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
        tt.SetPlaceholderText Text:="Enter remediation text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        '''''''''Meta data Attributes'''''''''
        
        table1.Rows(table1.Rows.Count - 4).Shading.BackgroundPatternColor = RGB(169, 169, 169)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 4).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Meta Data Attributes"
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
        tt.SetPlaceholderText , , "Select your metadata tag"
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
        tt.SetPlaceholderText , , "Select your metadata tag"
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
        tt.SetPlaceholderText , , "Select your metadata tag"
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
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        'table1.Rows(table1.Rows.Count).Select
        
        'Call GetMyPicture
        
               
       ' Set tt = Nothing
        

End Sub
Sub copytableGO(MyValue, QuestionNo)
Dim table1 As Table
Dim tt As ContentControl
Dim table2 As Table
Dim addbutton
varNumRows = MyValue + 7
'QuestionNo = QuestionNo + 1
Selection.HomeKey Unit:=wdStory
Selection.EndKey Unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 15, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.ID = "VQ"
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        table1.Rows(1).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(1).Select
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
        
        table1.Rows(2).Select
        
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
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(3, 1).Select
        
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
        
        table1.Cell(3, 2).Select
        
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
                
        table1.Rows(4).Select
        
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
        table1.Rows(5).Select
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
    
                    
       
        
      For i = 6 To MyValue + 5
        
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
        tt.Title = "Rationale Text " & i - 5
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
             
        table1.Rows(table1.Rows.Count - 16).Select
        table1.Rows(table1.Rows.Count - 16).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 16, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 16, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 15).Select
        table1.Rows(table1.Rows.Count - 15).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 15, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 15, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 14).Select
        table1.Rows(table1.Rows.Count - 14).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 14, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose score in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

    table1.Cell(table1.Rows.Count - 14, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , "Choose difficulty in the range 1 to 10"
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 13).Select
        table1.Rows(table1.Rows.Count - 13).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 13, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 13, 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 12).Select
        table1.Rows(table1.Rows.Count - 12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 12, 1).Select
        
       
        
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
        
        table1.Cell(table1.Rows.Count - 12, 2).Select
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
        
        
  
        table1.Rows(table1.Rows.Count - 11).Shading.BackgroundPatternColor = RGB(255, 140, 0)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 11).Select
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
        
        
        table1.Rows(table1.Rows.Count - 10).Select
        table1.Rows(table1.Rows.Count - 10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 10, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 10, 2).Select
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
        
        
        
        
        table1.Rows(table1.Rows.Count - 9).Select
        table1.Rows(table1.Rows.Count - 9).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Cell(table1.Rows.Count - 9, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 9, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 8).Select
        table1.Rows(table1.Rows.Count - 8).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
               
        table1.Cell(table1.Rows.Count - 8, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 7).Select
        table1.Rows(table1.Rows.Count - 7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
                
        table1.Cell(table1.Rows.Count - 7, 2).Select
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
        tt.SetPlaceholderText Text:="Enter remediation text"
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        'tt.Range.Rows.Height = 30
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        '''''''''Meta data Attributes'''''''''
        
        table1.Rows(table1.Rows.Count - 4).Shading.BackgroundPatternColor = RGB(169, 169, 169)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 4).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Meta Data Attributes"
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
        tt.SetPlaceholderText , , "Select your metadata tag"
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
        tt.SetPlaceholderText , , "Select your metadata tag"
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
        tt.SetPlaceholderText , , "Select your metadata tag"
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
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        Set tt = Nothing
        
        
        'table1.Rows(table1.Rows.Count).Select
        
        'Call GetMyPicture
        
               
       ' Set tt = Nothing
        

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
        Dim ws As New wshshell
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
Sub copytableupdatemc(vartitle, MyValue, varqtext, varanswer, varrationale, varcorrecttxt, varscoretxt, vardifficultytxt, varcorrectrationaletxt, varincorrectrationaletxt, varWeblinkRemediationTxt, varWeblinkRemediationLink, varWeblinkRemediationTooltip, varTextRemediationTxt, vartagtype, vartagvalue)
Dim table1 As Table
Dim tt As ContentControl
Dim table2 As Table
Dim addbutton
varNumRows = MyValue + 6
UpdateQuestionNo = UpdateQuestionNo + 1
Selection.HomeKey Unit:=wdStory
Selection.EndKey Unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 16, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.ID = "MC"
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        table1.Rows(1).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(1).Select
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
        
        
        
        table1.Rows(2).Select
        
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
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(3, 1).Select
        
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
        
        table1.Cell(3, 2).Select
        
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
                
        table1.Rows(4).Select
        
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
        table1.Rows(5).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Answer_head"
        tt.SetPlaceholderText Text:="Answer Text - Multiple Choice"
        
        tt.Range.Font.Size = "13"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
    
         
        
      For i = 6 To MyValue + 5
        
       ' table1.Rows(5).SetHeight rowHeight:=ActiveDocument.PageSetup.PageHeight - 300, HeightRule:=wdRowHeightExactly
        
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 5
        tt.Range.Style = "Answer_Text"
        tt.SetPlaceholderText Text:=varanswer(i - 6)
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
        tt.Title = "Rationale Text " & i - 5
        tt.SetPlaceholderText Text:=varrationale(i - 6)
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
               
      
               
        table1.Rows(table1.Rows.Count - 16).Select
        table1.Rows(table1.Rows.Count - 16).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 16, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 16, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 15).Select
        table1.Rows(table1.Rows.Count - 15).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 15, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 15, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 14).Select
        table1.Rows(table1.Rows.Count - 14).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 14, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , varscoretxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

        table1.Cell(table1.Rows.Count - 14, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , vardifficultytxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 13).Select
        table1.Rows(table1.Rows.Count - 13).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 13, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 13, 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 12).Select
        table1.Rows(table1.Rows.Count - 12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 12, 1).Select
        
       
        
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
        
        table1.Cell(table1.Rows.Count - 12, 2).Select
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
        
        
  
        table1.Rows(table1.Rows.Count - 11).Shading.BackgroundPatternColor = RGB(255, 140, 0)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 11).Select
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
        
        
        table1.Rows(table1.Rows.Count - 10).Select
        table1.Rows(table1.Rows.Count - 10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 10, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 10, 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 9).Select
        table1.Rows(table1.Rows.Count - 9).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Cell(table1.Rows.Count - 9, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 9, 2).Select
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
        
        
        '''''''''Meta data Attributes'''''''''
        
        table1.Rows(table1.Rows.Count - 4).Shading.BackgroundPatternColor = RGB(169, 169, 169)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 4).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Meta Data Attributes"
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

Sub copytableupdatecm(vartitle, MyValue, varqtext, varanswer, varrationale, varcorrecttxt, varscoretxt, vardifficultytxt, varcorrectrationaletxt, varincorrectrationaletxt, varWeblinkRemediationTxt, varWeblinkRemediationLink, varWeblinkRemediationTooltip, varTextRemediationTxt, vartagtype, vartagvalue)
Dim table1 As Table
Dim tt As ContentControl
Dim table2 As Table
Dim addbutton
varNumRows = MyValue + 6
UpdateQuestionNo = UpdateQuestionNo + 1
Selection.HomeKey Unit:=wdStory
Selection.EndKey Unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 16, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        table1.Rows(1).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(1).Select
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
        
         table1.Rows(2).Select
        
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
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(3, 1).Select
        
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
        
        table1.Cell(3, 2).Select
        
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
        
        table1.Rows(4).Select
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
        table1.Rows(5).Select
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
        For i = 6 To MyValue + 5
            
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 5
        tt.Range.Style = "Answer_Text"
        tt.SetPlaceholderText Text:=varanswer(i - 6)
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
        tt.Title = "Rationale Text " & i - 5
        tt.SetPlaceholderText Text:=varrationale(i - 6)
        tt.Range.Font.Size = "10"
        tt.Range.Font.ColorIndex = wdBlack
        tt.Range.Font.Bold = False
        tt.Range.Rows.Height = 50
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = False
        
        Set tt = Nothing
         
        Next
               
        table1.Rows(table1.Rows.Count - 16).Select
        table1.Rows(table1.Rows.Count - 16).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 16, 1).Select
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
            table1.Cell(table1.Rows.Count - 16, i).Split NumColumns:=(MyValue * 2) + 1
            table1.Cell(table1.Rows.Count - 16, 2).Select
            Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
            tt.SetPlaceholderText Text:=" "
            tt.Range.Style = "Check_box_no"
            tt.Range.Font.Size = "10"
            tt.Range.Font.Name = "Verdana"
            tt.LockContentControl = True
            tt.LockContents = False
            
        End If
        
        If i Mod 2 > 0 Then
            table1.Cell(table1.Rows.Count - 16, i).Select
            
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
            table1.Cell(table1.Rows.Count - 16, i).Select
            Selection.Collapse Direction:=wdCollapseEnd
            Set tt = ActiveDocument.ContentControls.Add(wdContentControlText)
                
            tt.SetPlaceholderText Text:=j
            tt.Range.Style = "Check_box_no"
            Set tt = Nothing
            j = j + 1
        End If
        
        Next
        
        
        'Set tt = Nothing
        table1.Rows(table1.Rows.Count - 15).Select
        table1.Rows(table1.Rows.Count - 15).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 15, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 15, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 14).Select
        table1.Rows(table1.Rows.Count - 14).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 14, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , varscoretxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

    table1.Cell(table1.Rows.Count - 14, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , vardifficultytxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 13).Select
        table1.Rows(table1.Rows.Count - 13).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 13, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 13, 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 12).Select
        table1.Rows(table1.Rows.Count - 12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 12, 1).Select
        
       
        
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
        
        table1.Cell(table1.Rows.Count - 12, 2).Select
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
        
        
  
        table1.Rows(table1.Rows.Count - 11).Shading.BackgroundPatternColor = RGB(255, 140, 0)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 11).Select
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
        
        
        table1.Rows(table1.Rows.Count - 10).Select
        table1.Rows(table1.Rows.Count - 10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 10, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 10, 2).Select
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
        
        
        
        
        table1.Rows(table1.Rows.Count - 9).Select
        table1.Rows(table1.Rows.Count - 9).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Cell(table1.Rows.Count - 9, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 9, 2).Select
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
        
        
        '''''''''Meta data Attributes'''''''''
        
        table1.Rows(table1.Rows.Count - 4).Shading.BackgroundPatternColor = RGB(169, 169, 169)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 4).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Meta Data Attributes"
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

Sub copytableupdatetf(vartitle, MyValue, varqtext, varanswer, varrationale, varcorrecttxt, varscoretxt, vardifficultytxt, varcorrectrationaletxt, varincorrectrationaletxt, varWeblinkRemediationTxt, varWeblinkRemediationLink, varWeblinkRemediationTooltip, varTextRemediationTxt, vartagtype, vartagvalue)
Dim table1 As Table
Dim tt As ContentControl
Dim table2 As Table
Dim addbutton
UpdateQuestionNo = UpdateQuestionNo + 1
Selection.HomeKey Unit:=wdStory
Selection.EndKey Unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=8 + 16, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        table1.Rows(1).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(1).Select
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
        
         table1.Rows(2).Select
        
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
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(3, 1).Select
        
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
        
        table1.Cell(3, 2).Select
        
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
        
        table1.Rows(4).Select
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
        
        table1.Rows(5).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Answer_head"
        tt.SetPlaceholderText Text:="Answer Text - True or Fasle"
       
        tt.Range.Font.Size = "13"
        tt.Range.Font.ColorIndex = wdBlue
        tt.Range.Font.Bold = True
        tt.Range.Font.Name = "Verdana"
        tt.LockContentControl = True
        tt.LockContents = True
        
        
        table1.Rows(6).Select
        table1.Rows(6).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(6, 1).Select
        
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
        
         table1.Cell(6, 2).Select
        
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
        
        table1.Rows(7).Select
        table1.Rows(7).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(7, 1).Select
        
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
        
         table1.Cell(7, 2).Select
        
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
        
        
        table1.Rows(table1.Rows.Count - 16).Select
        table1.Rows(table1.Rows.Count - 16).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 16, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 16, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Title = "Correct Answer"
        tt.LockContentControl = True
        tt.Range.Style = "Correct_Answer"
        tt.SetPlaceholderText , , varcorrecttxt
        tt.DropdownListEntries.Add ("True")
        tt.DropdownListEntries.Add ("False")
        
        'Set tt = Nothing
        
        table1.Rows(table1.Rows.Count - 15).Select
        table1.Rows(table1.Rows.Count - 15).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 15, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 15, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 14).Select
        table1.Rows(table1.Rows.Count - 14).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 14, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , varscoretxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

    table1.Cell(table1.Rows.Count - 14, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , vardifficultytxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 13).Select
        table1.Rows(table1.Rows.Count - 13).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 13, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 13, 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 12).Select
        table1.Rows(table1.Rows.Count - 12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 12, 1).Select
        
       
        
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
        
        table1.Cell(table1.Rows.Count - 12, 2).Select
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
        
        
  
        table1.Rows(table1.Rows.Count - 11).Shading.BackgroundPatternColor = RGB(255, 140, 0)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 11).Select
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
        
        
        table1.Rows(table1.Rows.Count - 10).Select
        table1.Rows(table1.Rows.Count - 10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 10, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 10, 2).Select
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
        
        
        
        table1.Rows(table1.Rows.Count - 9).Select
        table1.Rows(table1.Rows.Count - 9).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Cell(table1.Rows.Count - 9, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 9, 2).Select
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
        
        
        '''''''''Meta data Attributes'''''''''
        
        table1.Rows(table1.Rows.Count - 4).Shading.BackgroundPatternColor = RGB(169, 169, 169)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 4).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Meta Data Attributes"
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

Sub copytableupdatevq(vartitle, MyValue, varqtext, varanswer, varrationale, varcorrecttxt, varscoretxt, vardifficultytxt, varcorrectrationaletxt, varincorrectrationaletxt, varWeblinkRemediationTxt, varWeblinkRemediationLink, varWeblinkRemediationTooltip, varTextRemediationTxt, vartagtype, vartagvalue, varadditionalfile)
Dim table1 As Table
Dim tt As ContentControl
Dim table2 As Table
Dim addbutton
varNumRows = MyValue + 7
UpdateQuestionNo = UpdateQuestionNo + 1
Selection.HomeKey Unit:=wdStory
Selection.EndKey Unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 16, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.ID = "VQ"
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        table1.Rows(1).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(1).Select
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
        
        table1.Rows(2).Select
        
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
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(3, 1).Select
        
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
        
        table1.Cell(3, 2).Select
        
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
                
        table1.Rows(4).Select
        
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
        table1.Rows(5).Select
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
    
                    
       
        
      For i = 6 To MyValue + 5
        
       ' table1.Rows(5).SetHeight rowHeight:=ActiveDocument.PageSetup.PageHeight - 300, HeightRule:=wdRowHeightExactly
        
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 5
        tt.Range.Style = "Answer_Text"
        tt.SetPlaceholderText Text:=varanswer(i - 6)
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
        tt.Title = "Rationale Text " & i - 5
        tt.SetPlaceholderText Text:=varrationale(i - 6)
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
        
        
               
        table1.Rows(table1.Rows.Count - 16).Select
        table1.Rows(table1.Rows.Count - 16).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 16, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 16, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Additional_Fields_head"
        tt.SetPlaceholderText Text:=""
        tt.LockContentControl = True
        
        table1.Cell(table1.Rows.Count - 16, 3).Select
        'Selection.InlineShapes.AddOLEControl ClassType:="Forms.CommandButton.1"
        Set shp = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
        
        'shp.Range.Style = "Chosse_file"
        With shp.OLEFormat.Object
            .Object.Caption = "Select File"
            '.Object.Style = "Chosse_file"
        End With
        varbtn = Replace(shp.OLEFormat.Object.Name, "CommandButton", "")
        Set shp = Nothing
        
        
        table1.Cell(table1.Rows.Count - 16, 4).Select
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
        
        table1.Rows(table1.Rows.Count - 15).Select
        table1.Rows(table1.Rows.Count - 15).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 15, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 15, 2).Select
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
        
       table1.Rows(table1.Rows.Count - 14).Select
        table1.Rows(table1.Rows.Count - 14).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 14, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , varscoretxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

    table1.Cell(table1.Rows.Count - 14, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , vardifficultytxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 13).Select
        table1.Rows(table1.Rows.Count - 13).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 13, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 13, 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 12).Select
        table1.Rows(table1.Rows.Count - 12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 12, 1).Select
        
       
        
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
        
        table1.Cell(table1.Rows.Count - 12, 2).Select
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
        
        
  
       table1.Rows(table1.Rows.Count - 11).Shading.BackgroundPatternColor = RGB(255, 140, 0)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 11).Select
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
        
        
        table1.Rows(table1.Rows.Count - 10).Select
        table1.Rows(table1.Rows.Count - 10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 10, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 10, 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 9).Select
        table1.Rows(table1.Rows.Count - 9).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Cell(table1.Rows.Count - 9, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 9, 2).Select
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
        
        
        '''''''''Meta data Attributes'''''''''
        
        table1.Rows(table1.Rows.Count - 4).Shading.BackgroundPatternColor = RGB(169, 169, 169)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 4).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Meta Data Attributes"
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
Sub copytableupdateII(vartitle, MyValue, varqtext, varanswer, varrationale, varcorrecttxt, varscoretxt, vardifficultytxt, varcorrectrationaletxt, varincorrectrationaletxt, varWeblinkRemediationTxt, varWeblinkRemediationLink, varWeblinkRemediationTooltip, varTextRemediationTxt, vartagtype, vartagvalue, varadditionalfile)
Dim table1 As Table
Dim tt As ContentControl
Dim table2 As Table
Dim addbutton
varNumRows = MyValue + 7
UpdateQuestionNo = UpdateQuestionNo + 1
Selection.HomeKey Unit:=wdStory
Selection.EndKey Unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 16, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.ID = "MC"
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        table1.Rows(1).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(1).Select
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
        
        table1.Rows(2).Select
        
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
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(3, 1).Select
        
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
        
        table1.Cell(3, 2).Select
        
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
                
        table1.Rows(4).Select
        
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
        table1.Rows(5).Select
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
    
                    
       
        
      For i = 6 To MyValue + 5
        
       ' table1.Rows(5).SetHeight rowHeight:=ActiveDocument.PageSetup.PageHeight - 300, HeightRule:=wdRowHeightExactly
        
        table1.Rows(i).Select
        table1.Rows(i).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(i, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Title = "Answer Text " & i - 5
        tt.Range.Style = "Answer_Text"
        tt.SetPlaceholderText Text:=varanswer(i - 6)
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
        tt.Title = "Rationale Text " & i - 5
        tt.SetPlaceholderText Text:=varrationale(i - 6)
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
               
        table1.Rows(table1.Rows.Count - 16).Select
        table1.Rows(table1.Rows.Count - 16).Cells.Split NumColumns:=4, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 16, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 16, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        tt.Range.Style = "Additional_Fields_head"
        tt.SetPlaceholderText Text:=""
        tt.LockContentControl = True
        
        table1.Cell(table1.Rows.Count - 16, 3).Select
        'Selection.InlineShapes.AddOLEControl ClassType:="Forms.CommandButton.1"
        Set shp = Selection.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
        
        'shp.Range.Style = "Chosse_file"
        With shp.OLEFormat.Object
            .Object.Caption = "Select File"
            '.Object.Name = "Chosse_file"
        End With
        varbtn = Replace(shp.OLEFormat.Object.Name, "CommandButton", "")
        Set shp = Nothing
        
        table1.Cell(table1.Rows.Count - 16, 4).Select
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
        
         table1.Rows(table1.Rows.Count - 15).Select
        table1.Rows(table1.Rows.Count - 15).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 15, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 15, 2).Select
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
        
        table1.Rows(table1.Rows.Count - 14).Select
        table1.Rows(table1.Rows.Count - 14).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 14, 1).Select
        
        
         table1.Rows(table1.Rows.Count - 14).Select
        table1.Rows(table1.Rows.Count - 14).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 14, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , varscoretxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

        table1.Cell(table1.Rows.Count - 14, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , vardifficultytxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 13).Select
        table1.Rows(table1.Rows.Count - 13).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 13, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 13, 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 12).Select
        table1.Rows(table1.Rows.Count - 12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 12, 1).Select
        
       
        
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
        
        table1.Cell(table1.Rows.Count - 12, 2).Select
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
        
        
  
        table1.Rows(table1.Rows.Count - 11).Shading.BackgroundPatternColor = RGB(255, 140, 0)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 11).Select
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
        
        
        table1.Rows(table1.Rows.Count - 10).Select
        table1.Rows(table1.Rows.Count - 10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 10, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 10, 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 9).Select
        table1.Rows(table1.Rows.Count - 9).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Cell(table1.Rows.Count - 9, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 9, 2).Select
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
        
        
        '''''''''Meta data Attributes'''''''''
        
        table1.Rows(table1.Rows.Count - 4).Shading.BackgroundPatternColor = RGB(169, 169, 169)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 4).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Meta Data Attributes"
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

Sub copytableupdatego(vartitle, MyValue, varqtext, varanswer, varrationale, varcorrecttxt, varscoretxt, vardifficultytxt, varcorrectrationaletxt, varincorrectrationaletxt, varWeblinkRemediationTxt, varWeblinkRemediationLink, varWeblinkRemediationTooltip, varTextRemediationTxt, vartagtype, vartagvalue, varadditionalpath)
Dim table1 As Table
Dim tt As ContentControl
Dim table2 As Table
Dim addbutton
varNumRows = MyValue + 7
UpdateQuestionNo = UpdateQuestionNo + 1
Selection.HomeKey Unit:=wdStory
Selection.EndKey Unit:=wdStory
Selection.InsertBreak WdBreakType.wdPageBreak

   Set table1 = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=varNumRows + 15, NumColumns:= _
        1, AutoFitBehavior:= _
        wdAutoFitWindow)
        
        table1.Select
        table1.ID = "VQ"
        table1.Shading.BackgroundPatternColor = RGB(203, 203, 203)
        table1.Rows(1).Shading.BackgroundPatternColor = RGB(81, 169, 77)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(1).Select
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
        
        table1.Rows(2).Select
        
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
        
        table1.Rows(3).Select
        table1.Rows(3).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
         table1.Cell(3, 1).Select
        
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
        
        table1.Cell(3, 2).Select
        
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
                
        table1.Rows(4).Select
        
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
        table1.Rows(5).Select
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
    
                    
       
        
      For i = 6 To MyValue + 5
        
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
        tt.SetPlaceholderText , , varadditionalpath(i - 6)
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
        tt.Title = "Rationale Text " & i - 5
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
             
        table1.Rows(table1.Rows.Count - 16).Select
        table1.Rows(table1.Rows.Count - 16).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 16, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 16, 2).Select
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Title = "Correct Answer"
        tt.LockContentControl = True
        tt.Range.Style = "Correct_Answer"
        tt.SetPlaceholderText , , varcorrecttxt
        For i = 1 To MyValue
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
               
        table1.Rows(table1.Rows.Count - 15).Select
        table1.Rows(table1.Rows.Count - 15).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 15, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 15, 2).Select
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
        
       table1.Rows(table1.Rows.Count - 14).Select
        table1.Rows(table1.Rows.Count - 14).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 14, 1).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Score"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , varscoretxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing

    table1.Cell(table1.Rows.Count - 14, 2).Select
        
        Set tt = ActiveDocument.ContentControls.Add(wdContentControlDropdownList)
        tt.Range.Style = "Difficulty"
        tt.LockContentControl = True
        tt.SetPlaceholderText , , vardifficultytxt
        For i = 1 To 10
        tt.DropdownListEntries.Add (i)
        Next
        Set tt = Nothing
        
        
        table1.Rows(table1.Rows.Count - 13).Select
        table1.Rows(table1.Rows.Count - 13).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 13, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 13, 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 12).Select
        table1.Rows(table1.Rows.Count - 12).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 12, 1).Select
        
       
        
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
        
        table1.Cell(table1.Rows.Count - 12, 2).Select
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
        
        
  
       table1.Rows(table1.Rows.Count - 11).Shading.BackgroundPatternColor = RGB(255, 140, 0)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 11).Select
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
        
        
        table1.Rows(table1.Rows.Count - 10).Select
        table1.Rows(table1.Rows.Count - 10).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
        
        table1.Cell(table1.Rows.Count - 10, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 10, 2).Select
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
        
        
        table1.Rows(table1.Rows.Count - 9).Select
        table1.Rows(table1.Rows.Count - 9).Cells.Split NumColumns:=2, MergeBeforeSplit:=True
               
        table1.Cell(table1.Rows.Count - 9, 1).Select
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
        
        table1.Cell(table1.Rows.Count - 9, 2).Select
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
        
        
        '''''''''Meta data Attributes'''''''''
        
        table1.Rows(table1.Rows.Count - 4).Shading.BackgroundPatternColor = RGB(169, 169, 169)
        'table1.Rows(3).Shading.ForegroundPatternColor = wdColorBlue
        'Question heading
        table1.Rows(table1.Rows.Count - 4).Select
         Set tt = ActiveDocument.ContentControls.Add(wdContentControlRichText)
        
        tt.SetPlaceholderText Text:="Meta Data Attributes"
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

Public Function getobject(varobjectname As String) As OLEFormat
'"Qno_" & QuestionNo & "_txt_Metadata_Search_keyword"
For Each shp In ActiveDocument.InlineShapes
    If shp.OLEFormat.Object.Name = varobjectname Then
     Set getobject = shp.OLEFormat
     Exit For
    End If
Next
End Function



Private Sub vba_code(objname As String)
Dim sCode As String
vartxtname = Replace(objname, "_cmd_", "_txt_")
varlstname = Replace(objname, "_cmd_", "_lst_")

sCode = sCode & "Private Sub " & objname & "_Click()" & vbCrLf
sCode = sCode & "If Len(" & vartxtname & ".text" & ") >= 3 Then" & vbCrLf

sCode = sCode & "Dim objHttp As Object" & vbCrLf

sCode = sCode & "Set objHttp = CreateObject(""Msxml2.ServerXMLHTTP.6.0"")" & vbCrLf
sCode = sCode & "varData = ""{""""clientCode"""" : """"impelsys"""" , """"secretKey"""" : """"impelsys#20#17"""" , """"email"""" : """"abdul.rahman@impelsys.com"""", """"firstName"""" : """"Abdul"""", """"lastName"""" : """"Rahman"""", """"clientUserId"""" : """"101""""}""" & vbCrLf

sCode = sCode & "Call objHttp.Open(""POST"", ""http://qa-quizzingplatform.impelsys.com/api/authenticate"", False)" & vbCrLf
sCode = sCode & "Call objHttp.setRequestHeader(""Content-Type"", ""application/json"")" & vbCrLf
sCode = sCode & "Call objHttp.setRequestHeader(""Accept"", ""application/json"")" & vbCrLf
sCode = sCode & "Call objHttp.send(varData)"
sCode = sCode & "''Response.AddHeader ""Content-Type"", ""application/json;charset=UTF-8""" & vbCrLf
sCode = sCode & "''Response.Charset = ""UTF-8"""
sCode = sCode & "varKey = Replace(Replace(objHttp.responseText, ""{""""token"""":"""""", """"), """"""}"", """")" & vbCrLf

sCode = sCode & "Call objHttp.Open(""GET"", ""http://qa-quizzingplatform.impelsys.com/api/products/110000/metadata/SC_01"", False)" & vbCrLf
sCode = sCode & "Call objHttp.setRequestHeader(""Authorization"", varKey)" & vbCrLf
sCode = sCode & "Call objHttp.send("""")" & vbCrLf
sCode = sCode & "'JsonText = objHttp.responseText" & vbCrLf
sCode = sCode & "JsonText = ""{""""totalMetadata"""":5,""""metadataType"""":""""HIERARCHY"""",""""metadata"""":[{""""id"""":1,""""name"""":""""HLLibrary"""",""""nodePath"""":""""HLLibrary""""},{""""id"""":2,""""name"""":""""HealthProfessional"""",""""nodePath"""":""""HLLibrary\HealthProfessional""""},{""""id"""":3,""""name"""":""""medicine"""",""""nodePath"""":""""HLLibrary\HealthProfessional\medicine""""},{""""id"""":5,""""name"""":""""surgery"""",""""nodePath"""":""""HLLibrary\HealthProfessional\medicine\surgery""""},{""""id"""":2,""""name"""":""""Medicaleducation"""",""""nodePath"""":""""HLLibrary\Medicaleducation""""},{""""id"""":3,""""name"""":""""ENT"""",""""nodePath"""":""""HLLibrary\Medicaleducation\ENT""""},{""""id"""":4,""""name"""":""""surgery"""",""""nodePath"""":""""HLLibrary\Medicaleducation\ENT\surgery""""},{""""id"""":6,""""name"""":""""Nursing"""",""""nodePath"""":""""HLLibrary\Nursing""""},{""""id"""":8,""""name"""":""""Diabetic"""",""""nodePath"""":""""HLLibrary\Nursing\Diabetic""""}" & _
",{""""id"""":9,""""name"""":""""surgery"""",""""nodePath"""":""""HLLibrary\Nursing\Diabetic\surgery""""}]}""" & vbCrLf
sCode = sCode & "Dim JSON As Object" & vbCrLf
sCode = sCode & "Set JSON = JSONParser.ParseJson(Replace(JsonText, ""\"", ""\\""))" & vbCrLf
sCode = sCode & "vari = 0" & vbCrLf
sCode = sCode & varlstname & ".Clear" & vbCrLf
sCode = sCode & varlstname & ".ColumnWidths = ""220;100;60""" & vbCrLf
sCode = sCode & "For Each Value In JSON(""metadata"")" & vbCrLf
    sCode = sCode & varlstname & ".AddItem" & vbCrLf
    sCode = sCode & varlstname & ".List(vari, 0) = Value(""nodePath"")" & vbCrLf
    sCode = sCode & varlstname & ".List(vari, 1) = Value(""name"")" & vbCrLf
    sCode = sCode & varlstname & ".List(vari, 2) = Value(""id"")" & vbCrLf
    sCode = sCode & "vari = vari + 1" & vbCrLf
sCode = sCode & "Next" & vbCrLf

sCode = sCode & "Else" & vbCrLf
sCode = sCode & "MsgBox ""Search Key word minimum three character.."", vbCritical, ""WK Quizzing Platform""" & vbCrLf
sCode = sCode & "End If" & vbCrLf
 
 sCode = sCode & "End Sub"
ActiveDocument.VBProject.VBComponents("ThisDocument").CodeModule.AddFromString sCode
End Sub
