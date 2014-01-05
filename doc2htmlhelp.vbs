' vim:set sw=4 ts=4 expandtab:
'The MIT License
'
'Copyright (c) 2008 s7taka@gmail.com
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in
'all copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
'THE SOFTWARE.

Option Explicit

Const wdPropertyTitle = 1
Const wdDialogFileOpen = 80
Const wdFormatHTML = 8
Const wdFormatFilteredHTML = 10

Class Parameters
    Public WordDocFileName
    Public DestDir
    Public CHMTitle
    Public DivisionLevel
    Public MarginLeft

    Public Function Clone()
        Set Clone = New Parameters
        Clone.WordDocFileName = WordDocFileName
        Clone.DestDir = DestDir
        Clone.CHMTitle = CHMTitle
        Clone.DivisionLevel = DivisionLevel
        Clone.MarginLeft = MarginLeft
    End Function
End Class

Class TocItem
    Public Title
    Public Child
    Public Link

    Public Sub Class_Initialize()
        Title = ""
        Link = ""
        Set Child = Nothing
    End Sub
End Class

Dim g_regexp: Set g_regexp = CreateObject("VBScript.RegExp")
Private Function RETest(str, Pattern)
    g_regexp.Pattern = Pattern
    RETest = g_regexp.Test(str)
End Function

Private Function REMatches(str, Pattern)
    g_regexp.Pattern = Pattern
    Set REMatches = g_regexp.Execute(str)
End Function

Private Sub WriteConsole(str)
    If (LCase(Right(WScript.FullName, 11)) = "cscript.exe") Then
        WScript.Echo str
    End If
End Sub

Private Function EscapeHHCTitle(str)
    Dim tmp, tmp2

    tmp = Replace(str, "&nbsp;", " ")

    Do
        tmp2 = tmp
        tmp = Replace(tmp2, "  ", " ")
        If tmp2 = tmp Then Exit Do
    Loop

    tmp2 = ""
    Dim i
    For i = 1 To Len(tmp)
        Dim ch: ch = Mid(tmp, i, 1)  
        If ch <> "&" then
            tmp2 = tmp2 & ch
        Else
            If Mid(tmp, i, 2) = "&#" Then
                Dim PosSC: PosSC = InStr(i, tmp, ";")
                If PosSC > 0 Then
                    ch = ChrW(Mid(tmp, i + 2, PosSC - (i + 2)))
                    i = PosSC
                End If
            End If
            tmp2 = tmp2 & ch
        End If
    Next
    tmp = tmp2

    tmp2 = ""
    For i = 1 To Len(tmp)
        tmp2 = tmp2 & Chr(Asc(Mid(tmp, i, 1)))
    Next

    EscapeHHCTitle = Trim(tmp2)
End Function

Private Sub SafeReadAllFile(FileName, Text)
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")

    Dim fi: Set fi = fso.OpenTextFile(FileName, 1, False, False)
    Text = fi.ReadAll()
    fi.Close

    If InStr(Text, Chr(0)) > 0 Then
        Dim TempFileName
        TempFileName = WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%\" & fso.GetFileName(FileName) & ".tmp")

        Set fi = fso.OpenTextFile(FileName, 1, False, False)
        Dim fo: Set fo = fso.CreateTextFile(TempFileName, True, False)
        Do While Not fi.AtEndOfStream
            fo.WriteLine Replace(fi.ReadLine(), Chr(0), "")
        Loop
        fo.Close
        fi.Close

        Set fi = fso.OpenTextFile(TempFileName, 1, False, False)
        Text = fi.ReadAll()
        fi.Close

        fso.DeleteFile TempFileName
    End If
End Sub

Private Function OpenWordFileAndResolveUnspecifiedParameters(Params)
    Set OpenWordFileAndResolveUnspecifiedParameters = Nothing
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")

    On Error Resume Next
    Dim wd: Set wd = CreateObject("Word.Application")
    If wd Is Nothing Then
        MsgBox "Microsoft Wordがインストールされていません。", 16, "doc2htmlhelp"
        Exit Function
    End If
    On Error GoTo 0

    Dim doc
    If Params.WordDocFileName = "" Then
        wd.Visible = True
        wd.Dialogs(wdDialogFileOpen).Show
        If wd.Documents.Count = 0 Then
            wd.Quit
            Set wd = Nothing
            MsgBox "Wordドキュメントファイルを指定してください。", 16, "doc2htmlhelp"
            Exit Function
        End If
        Set doc = wd.Documents(1)
        Params.WordDocFileName = doc.Path & "\" & doc.Name
        wd.Visible = False
    Else
        Params.WordDocFileName = fso.GetAbsolutePathName(Params.WordDocFileName)
        If Not fso.FileExists(Params.WordDocFileName) Then
            wd.Quit
            Set wd = Nothing
            MsgBox "Wordドキュメント(" & Params.WordDocFileName & ")が見つかりません。", 16, "doc2htmlhelp"
            Exit Function
        End If
        Set doc = wd.Documents.Open(Params.WordDocFileName, , True)
    End If

    If Params.CHMTitle = "" Then
        Params.CHMTitle = doc.BuiltInDocumentProperties(wdPropertyTitle)
        If Params.CHMTitle = "" Then
            Params.CHMTitle = fso.GetBaseName(doc.name)
        End If
    End If
    
    If Params.DestDir = "" Then
        Params.DestDir = fso.GetParentFolderName(Params.WordDocFileName) & "\" & fso.GetBaseName(doc.name)
    End If
    If Not fso.FolderExists(Params.DestDir) Then
        On Error Resume Next
        fso.CreateFolder Params.DestDir
        If Err.Number <> 0 Then
            doc.Close
            Set doc = Nothing
            wd.Quit
            Set wd = Nothing
            MsgBox "フォルダ(" & Params.DestDir & ")の作成に失敗しました。" & vbCrLf & Err.Description, 16, "doc2htmlhelp"
            Exit Function
        End If
        On Error GoTo 0
    End If

    If Params.MarginLeft = -99999 Then
        Params.MarginLeft = 0
        On Error Resume Next
        Dim style: Set Style = doc.Styles("見出し 1")
        If Err.Number = 0 Then
            If Style.ParagraphFormat.CharacterUnitLeftIndent < 0 Then
                Params.MarginLeft = -Style.ParagraphFormat.CharacterUnitLeftIndent * Style.Font.Size / 0.75
            Else
                If Style.ParagraphFormat.LeftIndent < 0 Then
                    Params.MarginLeft = -Style.ParagraphFormat.LeftIndent / 0.75
                End If
            End If
        End If
        On Error GoTo 0
    End If
    Set wd = Nothing
    Set OpenWordFileAndResolveUnspecifiedParameters = doc
End Function

Private Sub SaveAsHTMLFileAndClose(doc, HTMLFileName)
    Dim WordVersion: WordVersion = doc.Application.Version
    If WordVersion <= 9 Then
        doc.SaveAs HTMLFileName, wdFormatHTML
    Else
        doc.SaveAs HTMLFileName, wdFormatFilteredHTML
    End If
    doc.Application.Quit
    If WordVersion <= 9 Then
        MSOfficeFilter HTMLFileName
    End If
End Sub

Private Sub MSOfficeFilter(HTMLFileName)
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")

    Dim HTMLText
    SafeReadAllFile HTMLFileName, HTMLText

    Dim fo: Set fo = fso.CreateTextFile(HTMLFileName, True, False)

    Dim InConditionalComments: InConditionalComments = 0
    Dim PosLt, PosGt, PosText
    Dim Text, TagText

    PosLt = InStr(HTMLText, "<")
    Do While PosLt > 0
        PosGt = InStr(PosLt, HTMLText, ">")
        If PosGt = 0 Then
            TagText = Mid(HTMLText, PosLt)
        Else
            TagText = Mid(HTMLText, PosLt, PosGt + 1 - PosLt)
        End if
        If InStr(TagText, "<html") = 1 Then
            fo.Write "<html>"
        ElseIf InStr(TagText, "<link") = 1 Then
            '
        ElseIf InStr(TagText, "<meta") = 1 Then
            If InStr(TagText, "name=ProgId") = 0 Then
                fo.Write TagText
            End If
        ElseIf InStr(TagText, "<!--[if") = 1 Then
            InConditionalComments = InConditionalComments + 1
        ElseIf InStr(TagText, "<![endif]-->") = 1 Then
            InConditionalComments = InConditionalComments - 1
        ElseIf InStr(TagText, "<![if !") = 1 Or InStr(TagText, "<![endif]>") = 1 Then
            '
        ElseIf InStr(TagText, "<!--") = 1 Then
            If InStr(TagText, "/* Font Definitions */") Then
                Dim PosAt, PosRb, PosMso, PosSc: PosAt = 1: PosMso = 1
                Do While True
                    PosAt = InStr(PosAt, TagText, "@")
                    If PosAt = 0 Then
                        Exit Do
                    End If
                    PosRb = InStr(PosAt, TagText, "}")
                    If PosRb = 0 Then
                        Exit Do
                    End If
                    TagText = Left(TagText, PosAt - 1) & Mid(TagText, PosRb + 1)
                Loop
                Do While True
                    PosMso = InStr(PosMso, TagText, "mso-")
                    If PosMso = 0 Then
                        Exit Do
                    End If
                    PosSc = InStr(PosMso, TagText, ";")
                    If PosSc = 0 Then
                        Exit Do
                    End If
                    TagText = Left(TagText, PosMso - 1) & Mid(TagText, PosSc + 1)
                Loop
            End If
            fo.Write TagText
        ElseIf InStr(TagText, "<o:p>") Or InStr(TagText, "</o:p>") = 1 Then
            '
        Else
            If InConditionalComments = 0 Then
                If InStr(TagText, "style='") > 0 Then
                    Dim Pos1, Pos2
                    Do While True
                        Pos1 = InStr(TagText, "mso-")
                        If Pos1 = 0 Then
                            Exit Do
                        End If
                        Pos2 = InStr(Pos1, TagText, ";")
                        If Pos2 = 0 Then
                            Pos2 = InStr(Pos1, TagText, "'")
                            If Pos2 = 0 Then
                                Exit Do
                            End If
                            Pos2 = Pos2 - 1
                        End If
                        TagText = Left(TagText, Pos1 - 1) & Mid(TagText, Pos2 + 1)
                    Loop
                    TagText = Replace(TagText, " style=''", "")
                    TagText = Replace(TagText, "style=''", "")
                End If
                fo.Write TagText
            End If
        End If
        If PosGt = 0 Or PosGt = Len(HTMLText) Then
            Exit Do
        End If

        PosText = PosGt + 1

        PosLt = InStr(PosText, HTMLText, "<")
        If PosLt = 0 Then
            Text = Mid(HTMLText, PosText)
        Else
            Text = Mid(HTMLText, PosText, PosLt - PosText)
        End If
        If InConditionalComments = 0 Then
            fo.Write Text
        End If
        If PosLt = 0 Or PosLt = Len(HTMLText) Then
            Exit Do
        End If

    Loop
    fo.Close
End Sub

Private Sub SplitHTML(HTMLFileName, Params, MaxTocLevel, dicHTMLFiles, dicTocLink, dicTocTree)
    Dim HTMLText
    SafeReadAllFile HTMLFileName, HTMLText

    Dim TopHTMLFileName
    Dim HTMLHeadBlock: HTMLHeadBlock = ""
    Dim HTMLStyleDef: HTMLStyleDef = ""
    Dim InHeadTag: InHeadTag = False
    Dim InStyleTag: InStyleTag = False
    Dim InStyleTagWithId: InStyleTagWithId = False
    Dim InTitleTag: InTitleTag = False

    Dim PosLt, PosGt, PosText
    Dim Text, TagText

    ' HTML HEAD部解析
    PosLt = InStr(HTMLText, "<")
    Do While PosLt > 0
        PosGt = InStr(PosLt, HTMLText, ">")
        If PosGt = 0 Then
            TagText = Mid(HTMLText, PosLt)
        Else
            TagText = Mid(HTMLText, PosLt, PosGt + 1 - PosLt)
        End if

        If TagText = "<head>" Then
            InHeadTag = True
        ElseIf TagText = "</head>" Then
            HTMLHeadBlock = HTMLHeadBlock & "<link rel=""stylesheet"" href=""doc.css"" type=""text/css"">" & vbCrLf 
            InHeadTag = False
            Exit Do
        ElseIf InStr(TagText, "<style") > 0 Then
            InStyleTag = True
            If InStr(TagText, "id=") > 0 Then
                HTMLHeadBlock = HTMLHeadBlock & TagText
                InStyleTagWithId = True
            End If
        ElseIf TagText = "</style>" Then
            InStyleTag = False
            If InStyleTagWithId Then
                HTMLHeadBlock = HTMLHeadBlock & TagText
                InStyleTagWithId = False
            End If
        ElseIf TagText = "<title>" Then
            InTitleTag = True
        ElseIf TagText = "</title>" Then
            InTitleTag = False
        ElseIf RETest(TagText, "^<!--.*") > 0 Then
            If InStyleTag And Not InStyleTagWithId Then
                HTMLStyleDef = Mid(TagText, 3, Len(TagText) - 6)
            End If
        Else
            If InStyleTag And Not InStyleTagWithId Then
                HTMLStyleDef = HTMLStyleDef & TagText
            ElseIf InHeadTag Then
                HTMLHeadBlock = HTMLHeadBlock & TagText
            End If
        End If

        PosText = PosGt + 1

        PosLt = InStr(PosText, HTMLText, "<")
        If PosLt = 0 Then
            Text = Mid(HTMLText, PosText)
        Else
            Text = Mid(HTMLText, PosText, PosLt - PosText)
        End If
        If InTitleTag Then
            '
        ElseIf InStyleTag And Not InStyleTagWithId Then
            HTMLStyleDef = HTMLStyleDef & Text
        ElseIf InHeadTag Then
            HTMLHeadBlock = HTMLHeadBlock & Text
        End If
        If PosLt = 0 Or PosLt = Len(HTMLText) Then
            Exit Do
        End If
    Loop
    
    ' スタイルシート書き込み
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim fo: Set fo = fso.CreateTextFile(Params.DestDir & "\doc.css", True, False)
    fo.Write HTMLStyleDef
    fo.Close

    ' HTML BODY部解析＆HTML分割
    Dim CurOutputFileName
    Dim CurTitle
    Dim CurHTMLTitle
    Dim ParsingHeading
    Dim HTMLBodyTag
    Dim HTMLHeadingTag 
    Dim DocLevel: DocLevel = 0
    Dim DocLevelCount(9)
    Dim dicSavedHTMLParentTag: Set dicSavedHTMLParentTag = CreateObject("Scripting.Dictionary")
    Dim dicHTMLParentTag: Set dicHTMLParentTag = CreateObject("Scripting.Dictionary")
    Dim dicHTMLParentTagPos: Set dicHTMLParentTagPos = CreateObject("Scripting.Dictionary")
    Dim dicCurLinkId: Set dicCurLinkId = CreateObject("Scripting.Dictionary")
    Dim tiTop 
    Dim PosMarkStart, PosMarkEnd
    Dim PosHeadingTag 
    Dim i

    MaxTocLevel = 0

    For i = 0 To UBound(DocLevelCount)
        DocLevelCount(i) = 0
    Next

    dicTocTree.RemoveAll
    TopHTMLFileName = Params.DestDir & "\doc"
    For i = 0 to Params.DivisionLevel
        TopHTMLFileName = TopHTMLFileName & "_0"
    Next
    TopHTMLFileName = TopHTMLFileName & ".htm"
    Set tiTop = New TocItem
    tiTop.Title = EscapeHHCTitle(Params.CHMTitle)
    tiTop.Link = fso.GetFileName(TopHTMLFileName)
    Set tiTop.Child = CreateObject("Scripting.Dictionary")
    dicTocTree.Add "top", tiTop
    
    CurHTMLTitle = Params.CHMTitle
    CurOutputFileName = TopHTMLFileName
    PosLt = InStr(HTMLText, "<body")
    Do While PosLt > 0
        PosGt = InStr(PosLt, HTMLText, ">")
        If PosGt = 0 Then
            TagText = Mid(HTMLText, PosLt)
        Else
            TagText = Mid(HTMLText, PosLt, PosGt + 1 - PosLt)
        End if

        If RETest(TagText, "^<body.*") Then
            HTMLBodyTag = TagText
            PosMarkStart = PosGt + 1
        ElseIf TagText = "</body>" Then
            dicHTMLFiles.Add dicHTMLFiles.Count, CurOutputFileName
            Set fo = fso.CreateTextFile(CurOutputFileName, True, False)
            fo.WriteLine "<html>"
            fo.WriteLine "<head>"
            fo.Write HTMLHeadBlock
            fo.WriteLine "<title>" & EscapeHHCTitle(CurHTMLTitle) & "</title>"
            fo.WriteLine "</head>"
            fo.WriteLine HTMLBodyTag
            If Params.MarginLeft > 0 Then
                fo.WriteLine "<div style='margin-left:" & Params.MarginLeft & "px'>"
            End If
            For i = 0 To dicSavedHTMLParentTag.Count - 1
                fo.WriteLine dicSavedHTMLParentTag.Item(i)
            Next
            PosMarkEnd = PosLt
            fo.Write Mid(HTMLText, PosMarkStart, PosMarkEnd - PosMarkStart)
            If Params.MarginLeft > 0 Then
                fo.WriteLine "</div>"
            End If
            fo.WriteLine "</body>"
            fo.WriteLine "</html>"
            fo.Close

            Exit Do
        ElseIf RETest(TagText, "^<(div|table|tr|td).*") Then
            dicHTMLParentTag.Add dicHTMLParentTag.Count, TagText
            dicHTMLParentTagPos.Add dicHTMLParentTagPos.Count, CStr(PosLt)
        ElseIf RETest(TagText, "</(div|table|tr|td)>") Then
            dicHTMLParentTag.Remove dicHTMLParentTag.Count - 1
            dicHTMLParentTagPos.Remove dicHTMLParentTagPos.Count - 1
        ElseIf RETest(TagText, "^<p\s+.*class=MsoToc\d.*") Then
            Dim TocLevel: TocLevel = CInt(Mid(TagText, InStr(TagText, "MsoToc") + 6, 1))
            If MaxTocLevel < TocLevel Then
                MaxTocLevel = TocLevel
            End If
        ElseIf RETest(TagText, "^<h\d.*") Or RETest(TagText, "^<p\s+.*class=(MsoHeading\d|Heading\d|Appendix|aa).*") Then
            HTMLHeadingTag = TagText
            PosHeadingTag = PosLt
            CurTitle = ""
            dicCurLinkId.RemoveAll
            ParsingHeading = True
        ElseIf RETest(TagText, "^</h\d.*") Or TagText = "</p>" Then
            If ParsingHeading Then
                If dicCurLinkId.Count > 0 Then
                    If Mid(HTMLHeadingTag, 2, 1) = "h" Then
                        DocLevel = CInt(Mid(HTMLHeadingTag, 3, 1)) - 1
                    Else
                        Dim PosClass: PosClass = InStr(HTMLHeadingTag, "class=")
                        DocLevel = 0
                        If RETest(Mid(HTMLHeadingTag, PosClass + 6), "Appendix.*") Then
                            DocLevel = 2 - 1
                        End If
                        For i = PosClass + 6 To Len(HTMLHeadingTag)
                            Dim ch: ch =  Mid(HTMLHeadingTag, i, 1)
                            If IsNumeric(ch) Then
                                DocLevel = CInt(ch) - 1
                                Exit For
                            End If
                        Next
                    End If
                    DocLevelCount(DocLevel) = DocLevelCount(DocLevel) + 1
                    For i = DocLevel + 1 To Params.DivisionLevel
                        DocLevelCount(i) = 0
                    Next
                    If DocLevel <= Params.DivisionLevel Then
                        dicHTMLFiles.Add dicHTMLFiles.Count, CurOutputFileName
                        Set fo = fso.CreateTextFile(CurOutputFileName, True, False)
                        fo.WriteLine "<html>"
                        fo.WriteLine "<head>"
                        fo.Write HTMLHeadBlock
                        fo.WriteLine "<title>" & EscapeHHCTitle(CurHTMLTitle) & "</title>"
                        fo.WriteLine "</head>"
                        fo.WriteLine HTMLBodyTag
                        If Params.MarginLeft > 0 Then
                            fo.WriteLine "<div style='margin-left:" & Params.MarginLeft & "px'>"
                        End If
                        For i = 0 To dicSavedHTMLParentTag.Count - 1
                            fo.WriteLine dicSavedHTMLParentTag.Item(i)
                        Next
                        PosMarkEnd = PosHeadingTag
                        For i = dicHTMLParentTag.Count - 1 To 0 Step -1
                            If InStr(dicHTMLParentTag.Item(i), "class=Section") > 0 Then
                                Exit For
                            End If
                            PosMarkEnd = dicHTMLParentTagPos.Item(i)
                        Next
                        fo.Write Mid(HTMLText, PosMarkStart, PosMarkEnd - PosMarkStart)
                        For i = dicSavedHTMLParentTag.Count - 1 To 0 Step -1
                            fo.WriteLine "</" & REMatches(dicSavedHTMLParentTag.Item(i), "^<(\w+).*")(0).SubMatches(0) & ">"
                        Next
                        dicSavedHTMLParentTag.RemoveAll
                        For i = 0 To dicHTMLParentTag.Count - 1
                            If InStr(dicHTMLParentTag.Item(i), "class=Section") = 0 Then
                                Exit For
                            End If
                            dicSavedHTMLParentTag.Add i, dicHTMLParentTag.Item(i)
                        Next
                        If Params.MarginLeft > 0 Then
                            fo.WriteLine "</div>"
                        End If
                        fo.WriteLine "</body>"
                        fo.WriteLine "</html>"
                        fo.Close

                        PosMarkStart = PosMarkEnd

                        CurOutputFileName = Params.DestDir & "\doc"
                        For i = 0 To Params.DivisionLevel
                            CurOutputFileName = CurOutputFileName & "_" & DocLevelCount(i)
                        Next
                        CurOutputFileName = CurOutputFileName & ".htm"
                    End If

                    For i = 0 To dicCurLinkId.Count - 1
                        dicTocLink.Add dicCurLinkId.Item(i), fso.GetFileName(CurOutputFileName & "#" & dicCurLinkId.Item(i))
                    Next

                    Dim node: Set node = dicTocTree
                    For i = 0 To DocLevel
                        If Not node.Exists("n" & CStr(DocLevelCount(i))) Then
                            Dim ti: Set ti = New TocItem
                            If i = DocLevel Then
                                ti.Title = EscapeHHCTitle(CurTitle)
                                ti.Link = fso.GetFileName(CurOutputFileName & "#" & dicCurLinkId.Item(0))
                                WriteConsole ti.Title
                            End If
                            Set ti.Child = CreateObject("Scripting.Dictionary")
                            node.Add "n" & CStr(DocLevelCount(i)), ti
                        Else
                            Set ti = node("n" & CStr(DocLevelCount(i)))
                        End If
                        Dim tmp: Set tmp = node.Item("n" & CStr(DocLevelCount(i))).Child
                        Set node = tmp
                    Next
                    If DocLevel <= Params.DivisionLevel Then
                        CurHTMLTitle = CurTitle
                    End If
                End If
                ParsingHeading = False
            End If
        ElseIf RETest(TagText, "^<a\s+.*name=""_Toc\d+") Then
            If ParsingHeading Then
                Dim Id: Id = REMatches(TagText, "name=""(_Toc\d+)")(0).SubMatches(0)
                dicCurLinkId.Add dicCurLinkId.Count, Id
            End If
        End If
        
        If PosGt = 0 Or PosGt = Len(HTMLText) Then
            Exit Do
        End If

        PosText = PosGt + 1

        PosLt = InStr(PosText, HTMLText, "<")
        If PosLt = 0 Then
            Text = Mid(HTMLText, PosText)
        Else
            Text = Mid(HTMLText, PosText, PosLt - PosText)
        End If
        If ParsingHeading Then
            CurTitle = CurTitle & Replace(Text, vbCrLf, "")
        End If
        If PosLt = 0 Or PosLt = Len(HTMLText) Then
            Exit Do
        End If
    Loop
End Sub

Private Sub ReplaceTocLink(dicHTMLFiles, dicTocLink)
    Dim FileName
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    For Each FileName In dicHTMLFiles.Items
        Dim fi: Set fi = fso.OpenTextFile(FileName, 1, False, False)
        Dim fo: Set fo = fso.CreateTextFile(FileName & ".tmp", True, False)
        Do While Not fi.AtEndOfStream
            Dim htmlLine: htmlLine = fi.ReadLine()
            If RETest(htmlLine, "href=""#_Toc\d+") Then
                Dim Id, m
                For Each m In REMatches(htmlLine, "href=""#(_Toc\d+)")
                    Id = m.SubMatches(0) 
                    htmlLine = Replace(htmlLine, "#" & Id, dicTocLink.Item(Id)) 
                Next
            End If
            fo.WriteLine htmlLine
        Loop
        fo.Close
        fi.Close

        fso.DeleteFile FileName
        fso.MoveFile FileName & ".tmp", FileName
    Next
End Sub

Private Sub CreateHHP(HHPFileName, TopHTMLFileName, Title)
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim fo: Set fo = fso.CreateTextFile(HHPFileName, True, False)
    Dim BaseFileName: BaseFileName = fso.GetParentFoldername(HHPFileName) & "\" & fso.GetBaseName(HHPFileName)
    fo.WriteLine _
        "[Options]" & vbCrLf & _
        "Binary Index = No" & vbCrLf & _
        "Compatibility = 1.1 Or later" & vbCrLf & _
        "Compiled file = " & fso.GetFileName(BaseFileName & ".chm") & vbCrLf & _
        "Contents file = " & fso.GetFileName(BaseFileName & ".hhc") & vbCrLf & _
        "Display compile progress=Yes" & vbCrLf & _
        "Full-text search=Yes" & vbCrLf & _
        "Language=0x411 日本語" & vbCrLf & _
        "Title = " & Title & vbCrLf & _
        "" & vbCrLf & _
        "" & vbCrLf & _
        "[Files]" & vbCrLf & _
        fso.GetFileName(TopHTMLFileName) & vbCrLf & _
        "" & vbCrLf & _
        "[INFOTYPES]"
    fo.Close
End Sub

Private Sub InternalCreateHHC(fo, node, MaxTocLevel, ByVal Level)
    Dim itm
    For Each itm In node.Items
        If Not IsEmpty(itm) Then
            If itm.Title <> "" And (itm.Link <> "" Or itm.Child.Count > 0) Then
                fo.WriteLine "<LI><OBJECT type=""text/sitemap"">"
                fo.WriteLine "<param name=""Name"" value=""" & itm.Title & """>"
                fo.WriteLine "<param name=""Local"" value=""" & itm.Link & """>"
                If itm.Child.Count > 0 And MaxTocLevel > Level Then
	                fo.WriteLine "<param name=""ImageNumber"" value=""1"">"
                End If
                fo.WriteLine "</OBJECT>"
                If itm.Child.Count > 0 And MaxTocLevel > Level Then
                    fo.WriteLine "<UL>"
                    InternalCreateHHC fo, itm.Child, MaxTocLevel, Level + 1
                    fo.WriteLine "</UL>"
                End If
            ElseIf itm.Title = "" And itm.Child.Count > 0 Then
                InternalCreateHHC fo, itm.Child, MaxTocLevel, Level + 1
            End If
        End If
    Next
End Sub

Private Sub CreateHHC(FileName, dicTocTree, MaxTocLevel)
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim fo: Set fo = fso.CreateTextFile(FileName, True, False)
    
    fo.WriteLine _
        "<!DOCTYPE HTML PUBLIC ""-//IETF//DTD HTML//EN"">" & vbCrLf & _
        "<HTML>" & vbCrLf & _
        "<HEAD>" & vbCrLf & _
        "</HEAD><BODY>" & vbCrLf & _
        "<OBJECT type=""text/site properties"">" & vbCrLf & _
        "<param name=""ImageType"" value=""Folder"">" & vbCrLf & _
        "</OBJECT>"
    
    fo.WriteLine "<UL>"
    InternalCreateHHC fo, dicTocTree, MaxTocLevel, 1
    fo.WriteLine "</UL>"
    
    fo.WriteLine _
        "</BODY></HTML>"
    
    fo.Close
End Sub

Private Sub CompileHHP(FileName)
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim HHCFileName: HHCFileName = fso.GetParentFolderName(WScript.ScriptFullName) & "\hhc.exe"

    If Not fso.FileExists(HHCFileName) Then
        HHCFileName = "C:\Program Files (x86)\HTML Help Workshop\hhc.exe"
        If Not fso.FileExists(HHCFileName) Then
            Dim ProgramFiles: ProgramFiles = CreateObject("Shell.Application").Namespace(&H26).Self.Path 
            HHCFileName = ProgramFiles & "\HTML Help Workshop\hhc.exe"
            Do While Not fso.FileExists(HHCFileName)
                HHCFileName = InputBox("HTML Helpコンパイラ(" & HHCFileName & ")が見つかりません。hhc.exeが存在するパスを指定してください", "doc2htmlhelp", HHCFileName)
                If HHCFileName = "" Then Exit Sub
            Loop
        End If
    End If
    
    Dim oExec: Set oExec = CreateObject("WScript.Shell").Exec("""" & HHCFileName & """ """ & FileName & """")
    Do Until oExec.StdOut.AtEndOfStream
        WriteConsole oExec.StdOut.ReadLine
    Loop
End Sub

Private Sub ConvertWordDocToHTMLHelp(Params)
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")

    WriteConsole "*** Wordドキュメントを開いています"
    Dim doc
    Set doc = OpenWordFileAndResolveUnspecifiedParameters(Params)
    If doc Is Nothing Then
        Exit Sub
    End If

    WriteConsole "*** WordドキュメントをHTML形式で保存しています"
    Dim TempHTMLFileName: TempHTMLFileName = Params.DestDir & "\doc.htm"
    SaveAsHTMLFileAndClose doc, TempHTMLFileName

    Set doc = Nothing
    
    WriteConsole "*** HTMLファイルを分割しています"
    Dim MaxTocLevel
    Dim dicHTMLFiles: Set dicHTMLFiles = CreateObject("Scripting.Dictionary")
    Dim dicTocLink: Set dicTocLink = CreateObject("Scripting.Dictionary")
    Dim dicTocTree: Set dicTocTree = CreateObject("Scripting.Dictionary")
    SplitHTML TempHTMLFileName, Params, MaxTocLevel, dicHTMLFiles, dicTocLink, dicTocTree
    
    WriteConsole "*** 目次のリンクを書き換えています"
    ReplaceTocLink dicHTMLFiles, dicTocLink
       
    WriteConsole "*** HTML Help Project ファイルを生成しています"
    CreateHHP Params.DestDir & "\" & fso.GetBaseName(Params.WordDocFileName) & ".hhp", dicHTMLFiles(0), Params.CHMTitle

    WriteConsole "*** HTML Help コンテンツファイルを生成しています"
    CreateHHC Params.DestDir & "\" & fso.GetBaseName(Params.WordDocFileName) & ".hhc", dicTocTree, MaxTocLevel

    WriteConsole "*** HTML Help プロジェクトをコンパイルしています"
    CompileHHP Params.DestDir & "\" & fso.GetBaseName(Params.WordDocFileName) & ".hhp"
End Sub

Sub Main()
    If WScript.Arguments.Named.Exists("?") Or Wscript.Arguments.Named.Exists("Help") Then
        MsgBox "使用方法: doc2htmlhelp.vbs [Wordドキュメントファイルパス] [/?] [/Title:HTML Help タイトル] [/DestDir:HTML Help生成先フォルダ] [/DivDocLevel:HTML分割するドキュメントレベル] [/MarginLeft:左余白幅]", 0, "doc2htmlhelp"
        Exit Sub
    End If

    Dim Start: Start = Timer

    Dim Params: Set Params = New Parameters
    Params.DestDir = WScript.Arguments.Named.Item("DestDir")
    Params.CHMTitle = WScript.Arguments.Named.Item("Title")
    Params.DivisionLevel = WScript.Arguments.Named.Item("DivDocLevel")
    If Params.DivisionLevel = "" Or Not IsNumeric(Params.DivisionLevel) Then
        Params.DivisionLevel = 1 - 1
    Else
        Params.DivisionLevel = CLng(Params.DivisionLevel) - 1
    End If
    Params.MarginLeft = WScript.Arguments.Named.Item("MarginLeft")
    If Params.MarginLeft = "" Or Not IsNumeric(Params.MarginLeft) Then
        Params.MarginLeft = -99999
    Else
        Params.MarginLeft = CLng(Params.MarginLeft)
    End If

    If WScript.Arguments.Unnamed.Count > 0 Then
        Dim FileName
        For Each FileName In WScript.Arguments.Unnamed
            Params.WordDocFileName = FileName
            Call ConvertWordDocToHTMLHelp(Params.Clone)
        Next
    Else
        Params.WordDocFileName = ""
        Call ConvertWordDocToHTMLHelp(Params.Clone)
    End If

    WScript.Echo "doc2htmlhelp: 完了 処理時間=" & Timer - Start & "[sec]"
End Sub

Call Main()

