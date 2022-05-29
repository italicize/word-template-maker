Option Explicit
'---5---10---15---20---25---30---35---40---45---50---55---60---65---70---75---80

'Notes: For these macros, a style name cannot end with the words List or Style.
'   A style can be named Arch, for example, but cannot be named Arch Style.

Const strDefaultStyleGallery As String = "Normal, No Spacing, Heading 1, " _
    & "Heading 2, Heading 3, Heading 4, Heading 5, Heading 6, Heading 7, " _
    & "Heading 8, Heading 9, Title, Subtitle, Subtle Emphasis, Emphasis, " _
    & "Intense Emphasis, Strong, Quote, Intense Quote, Subtle Reference, " _
    & "Intense Reference, Book Title, List Paragraph, Caption, TOC Heading"
    'Those built-in styles appear in the default style gallery in Word 2016.

Sub TestOfCommandParser()
    Call vbaAskForStyleDescriptions

'    Dim strDesc As String, strDescLow As String
'    strDesc = _
'        LongInputBox("Describe the style changes or other instructions.", _
'        "Input, please")
'    If strDesc = "" Then Exit Sub
'
'    strDesc = RemoveExtraSpaces(strDesc)
'    strDesc = CutLeftOrRight(strDesc, ".", wdRight)
'    strDesc = CutLeftOrRight(strDesc, "call ", wdLeft)
'    strDesc = CutLeftOrRight(strDesc, "run macro ", wdLeft)
'    strDesc = CutLeftOrRight(strDesc, "run ", wdLeft)
'
'    If "GenericTestMacro" = strDesc Then
'        GenericTestMacro
'    ElseIf "GenericTestMacro" = strDesc Then
'        vbaApplyStyleDescriptions strDesc 'Why doesn't this work?
'    End If
End Sub

'Function RemoveExtraSpaces(ByVal strString As String) As String
'    Dim lngA As Long
'    strString = Trim(strString)
'    For lngA = 1 To 3
'        strString = _
'            Replace(Replace(Replace(Replace(Replace(strString, _
'            "      ", " "), "     ", " "), "    ", " "), "   ", " "), "  ", " ")
'    Next lngA
'    RemoveExtraSpaces = strString
'End Function
'
'Function ReplaceWhiteSpaceWithSpaces(ByVal strString As String) As String
'    strString = Replace(strString, vbTab, " ")
''Add other kinds of white space, like nonbreaking spaces.
'    ReplaceWhiteSpaceWithSpaces = strString
'End Function
'
'Function TrimWS(ByVal str As String) As String
''Source: https://stackoverflow.com/questions/25184019/trim-all-types-of-whitespace-including-tabs
'    str = Trim(str)
'    Do Until Not Left(str, 1) = Chr(9)
'        str = Trim(Mid(str, 2, Len(str) - 1))
'        str = Trim(str)
'    Loop
'    Do Until Not Right(str, 1) = Chr(9)
'        str = Trim(Left(str, Len(str) - 1))
'        str = Trim(str)
'    Loop
'    TrimWS = str
'End Function
'
'Function CutLeftOrRight(ByVal strString As String, ByVal strExtra As String, _
'    ByVal lngPlace As Long) As String
'    Dim strLowercaseString As String, strLowercaseExtra As String
'    strLowercaseString = LCase(strString)
'    strLowercaseExtra = LCase(strExtra)
'    If lngPlace = wdLeft Then
'        If Left(strLowercaseString, Len(strExtra)) = strLowercaseExtra Then
'            CutLeftOrRight = Right(strString, Len(strString) - Len(strExtra))
'        Else
'            CutLeftOrRight = strString
'        End If
'    ElseIf lngPlace = wdRight Then
'        If Right(strLowercaseString, Len(strExtra)) = strLowercaseExtra Then
'            CutLeftOrRight = Left(strString, Len(strString) - Len(strExtra))
'        Else
'            CutLeftOrRight = strString
'        End If
'    Else
'        CutLeftOrRight = strString
'    End If
'End Function

Public Sub vbaAskForStyleDescriptions()
    Dim strDesc As String
    'Asks for style descriptions.
    strDesc = _
        LongInputBox("Describe the style changes or other instructions.", _
        "Settings to change")
    If strDesc = "" Then Exit Sub
    vbaApplyStyleDescriptions strDesc
End Sub

Public Sub vbaApplyStyleDescriptions(ByVal strDesc As String) '...as of 05/28/22
'Maybe add shading. See TrialOfShading.
'Maybe add Remove from Style Gallery.
'Maybe add Assign Value. Maybe reassign values to styles removed from gallery.
'Maybe add right 1" tab. Now it's only 1" right tab.
'Make it work with tab-delimited specifications.
'Investigate why 6 pt after didn't work as a default.
'Investigate why a List Number heading didn't have an indent of 0.5" for number.
'Explain that "10 pt font" needs to say "10 pt size" or "10 pt font size"
    Dim arrParas() As String ', strDesc As String
    Dim strPara As String, lngPara As Long, lngListPara As Long
    Dim strLabel As String, strLabelLow As String
    Dim arrSpecs() As String, strSpec As String
    Dim strSpecLow As String, lngSpec As Long, dblSpec As Double
    Dim arrStyles() As String, lngStyles As Long, strStyle As String
    Dim arrDefaultStyleGallery() As String, lngType As Long
    Dim arrList() As Variant, strList As String, lngList As Long
    Dim lngLevel As Long, lngLevels As Long
    Dim objListTemplate As ListTemplate
    Dim objActiveDocument As Document: Set objActiveDocument = ActiveDocument
    
    'Saves every line in an array.
    arrParas = vbaSaveParagraphsInAnArray(strDesc)
    'Reads each line of the style descriptions.
    For lngPara = LBound(arrParas) To UBound(arrParas)
        strPara = arrParas(lngPara)
        'Saves the specifications on each line (between commas) in an array.
        arrSpecs = Split(strPara, ", ")
        'Saves the first specification on each line, such as "Body Text style."
        strLabel = arrSpecs(0)
        strLabelLow = LCase(strLabel)
'Margins'
'-------'Sets the margins and page size.
        If strLabelLow = "margins" Or strLabelLow = "margin" _
            Or strLabelLow = "page" Or strLabelLow = "page size" _
            Or strLabelLow = "paper size" Then
            For lngSpec = 1 To UBound(arrSpecs)
                strSpec = arrSpecs(lngSpec)
                strSpecLow = LCase(strSpec)
                dblSpec = Val(strSpec)
                With objActiveDocument.PageSetup
                    If InStr(strSpecLow, "left") <> 0 Then
                        .LeftMargin = InchesToPoints(dblSpec)
                    ElseIf InStr(strSpecLow, "right") <> 0 Then
                        .RightMargin = InchesToPoints(dblSpec)
                    ElseIf InStr(strSpecLow, "top") <> 0 Then
                        .TopMargin = InchesToPoints(dblSpec)
                    ElseIf InStr(strSpecLow, "bottom") <> 0 Then
                        .BottomMargin = InchesToPoints(dblSpec)
                    ElseIf InStr(strSpecLow, "header") <> 0 Then
                        .HeaderDistance = InchesToPoints(dblSpec)
                    ElseIf InStr(strSpecLow, "footer") <> 0 Then
                        .FooterDistance = InchesToPoints(dblSpec)
                    ElseIf strSpecLow = "mirror margins" Then
                        .MirrorMargins = True
                    ElseIf strSpecLow = "no mirror margins" Then
                        .MirrorMargins = False
                    ElseIf InStr(strSpecLow, "portrait") _
                        Or InStr(strSpecLow, "vertical") Then
                        .Orientation = wdOrientPortrait
                    ElseIf InStr(strSpecLow, "landscape") _
                        Or InStr(strSpecLow, "horizontal") Then
                        .Orientation = wdOrientLandscape
                    ElseIf InStr(strSpecLow, "width") _
                        Or InStr(strSpecLow, "wide") Then
                        .PageWidth = InchesToPoints(dblSpec)
                    ElseIf InStr(strSpecLow, "height") _
                        Or InStr(strSpecLow, "high") _
                        Or InStr(strSpecLow, "tall") Then
                        .PageHeight = InchesToPoints(dblSpec)
                    ElseIf InStr(strSpecLow, "letter") Then
                        .PageWidth = InchesToPoints(8.5)
                        .PageHeight = InchesToPoints(11)
                    ElseIf InStr(strSpecLow, "tabloid") _
                        Or InStr(strSpecLow, "ledger") Then
                        .PageWidth = InchesToPoints(11)
                        .PageHeight = InchesToPoints(17)
                    ElseIf InStr(strSpecLow, "legal") Then
                        .PageWidth = InchesToPoints(8.5)
                        .PageHeight = InchesToPoints(14)
                    ElseIf InStr(strSpecLow, "executive") Then
                        .PageWidth = InchesToPoints(7.25)
                        .PageHeight = InchesToPoints(10.5)
                    ElseIf InStr(strSpecLow, "a3") Then
                        .PageWidth = InchesToPoints(11.69)
                        .PageHeight = InchesToPoints(16.54)
                    ElseIf InStr(strSpecLow, "a4") Then
                        .PageWidth = InchesToPoints(8.27)
                        .PageHeight = InchesToPoints(11.69)
                    ElseIf InStr(strSpecLow, "a5") Then
                        .PageWidth = InchesToPoints(5.83)
                        .PageHeight = InchesToPoints(8.27)
                    ElseIf InStr(strSpecLow, "a6") Then
                        .PageWidth = InchesToPoints(4.13)
                        .PageHeight = InchesToPoints(5.83)
                    ElseIf InStr(strSpecLow, "screen") Then
                        .PageWidth = InchesToPoints(6.5)
                        .PageHeight = InchesToPoints(5.18)
                    ElseIf InStr(strSpecLow, "ansi c") Then
                        .PageWidth = InchesToPoints(17)
                        .PageHeight = InchesToPoints(22)
                    ElseIf InStr(strSpecLow, "arch a") Then
                        .PageWidth = InchesToPoints(9)
                        .PageHeight = InchesToPoints(12)
                    ElseIf InStr(strSpecLow, "arch b") Then
                        .PageWidth = InchesToPoints(12)
                        .PageHeight = InchesToPoints(18)
                    ElseIf InStr(strSpecLow, "iso b5") Then
                        .PageWidth = InchesToPoints(6.93)
                        .PageHeight = InchesToPoints(9.85)
                    ElseIf InStr(strSpecLow, "iso b4") Then
                        .PageWidth = InchesToPoints(9.85)
                        .PageHeight = InchesToPoints(13.9)
                    ElseIf InStr(strSpecLow, "c5") Then
                        .PageWidth = InchesToPoints(6.37)
                        .PageHeight = InchesToPoints(9.01)
                    ElseIf InStr(strSpecLow, "jis b4") Then
                        .PageWidth = InchesToPoints(10.12)
                        .PageHeight = InchesToPoints(14.33)
                    ElseIf InStr(strSpecLow, "jis b3") Then
                        .PageWidth = InchesToPoints(14.33)
                        .PageHeight = InchesToPoints(20.28)
                    ElseIf InStr(strSpecLow, "slide") Then
                        .PageWidth = InchesToPoints(7.5)
                        .PageHeight = InchesToPoints(10)
                    ElseIf InStr(strSpecLow, "pocket") _
                        Or InStr(strSpecLow, "mass market") _
                        Or InStr(strSpecLow, "mass-market") Then
                        .PageWidth = InchesToPoints(4.25)
                        .PageHeight = InchesToPoints(6.87)
                    ElseIf InStr(strSpecLow, "digest") Then
                        .PageWidth = InchesToPoints(5.5)
                        .PageHeight = InchesToPoints(8.5)
                    ElseIf InStr(strSpecLow, "trade") Then
                        .PageWidth = InchesToPoints(6)
                        .PageHeight = InchesToPoints(9)
                    ElseIf InStr(strSpecLow, "different first page") _
                        Or InStr(strSpecLow, "different 1st page") Then
                        .DifferentFirstPageHeaderFooter = True
                    ElseIf InStr(strSpecLow, "different odd & even page") _
                        Or InStr(strSpecLow, "different odd and even page") _
                        Or InStr(strSpecLow, "different even & odd page") _
                        Or InStr(strSpecLow, "different even and odd page") _
                        Or InStr(strSpecLow, "different left and right page") _
                        Or InStr(strSpecLow, "different right and left page") _
                        Then
                        .OddAndEvenPagesHeaderFooter = True
                    End If
                End With
            Next lngSpec
'Styles '
'-------'Saves the style names in an array.
        ElseIf Right(strLabelLow, 6) = " style" _
            And Right(strLabelLow, 11) <> " list style" _
            And Right(strSpecLow, 11) <> " base style" _
            And Right(strSpecLow, 14) <> " default style" Then
            strStyle = Left(strLabel, Len(strLabel) - 6)
            lngStyles = lngStyles + 1
            If lngStyles = 1 Then
                ReDim arrStyles(1 To 1)
            Else
                ReDim Preserve arrStyles(1 To lngStyles)
            End If
            arrStyles(lngStyles) = strStyle
            
            'Adds a style if it doesn't exist.
            If Not vbaStyleExists(strStyle, objActiveDocument) Then
                If InStr(strPara, ", character style") <> 0 _
                    Or InStr(strPara, ", new character style") <> 0 Then
                    dblSpec = wdStyleTypeCharacter
                Else
                    dblSpec = wdStyleTypeParagraph
                End If
                objActiveDocument.Styles.Add strStyle, dblSpec
            End If
        End If
    Next lngPara
    
    'Reads each line of the styles descriptions again.
    For lngPara = LBound(arrParas) To UBound(arrParas)
        strPara = arrParas(lngPara)
        'Saves the specifications on each line (between commas) in an array.
        arrSpecs = Split(strPara, ", ")
        'Saves the first specification on each line, such as "Body Text style."
        strLabel = arrSpecs(0)
        strLabelLow = LCase(strLabel)
        
        'If any line begins "Style defaults," then...
        If strLabelLow = "style defaults" Or strLabelLow = "style default" _
            Or strLabelLow = "defaults for all defined styles" _
            Or strLabelLow = "defaults for defined styles" _
            Or strLabelLow = "default for all defined styles" _
            Or strLabelLow = "default for defined styles" Then
            'Applies the default specifications to all defined paragraph styles.
            For lngSpec = LBound(arrStyles) To UBound(arrStyles)
                strStyle = arrStyles(lngSpec)
                lngType = objActiveDocument.Styles(strStyle).Type
                If lngType = wdStyleTypeParagraph Then
                    'Sends style name and specs to the vbaDefineOneStyle macro.
                    vbaDefineOneStyle strStyle, arrSpecs
                End If
            Next lngSpec
        End If
    Next lngPara
    
    'Reads each line of the style descriptions again.
    For lngPara = LBound(arrParas) To UBound(arrParas)
        strPara = arrParas(lngPara)
        'Saves the specifications on each line (between commas) in an array.
        arrSpecs = Split(strPara, ", ")
        'Saves the first specification on each line, such as "Body Text style."
        strLabel = arrSpecs(0)
        strLabelLow = LCase(strLabel)
        
        'If the line begins with a style name, then...
        If Right(strLabelLow, 6) = " style" _
            And Right(strLabelLow, 11) <> " list style" _
            And Right(strSpecLow, 11) <> " base style" Then
            'Applies the specifications on the line to the style.
            strStyle = Left(strLabel, Len(strLabel) - 6)
            vbaDefineOneStyle strStyle, arrSpecs
'Gallery'
'-------'Customizes the style gallery on the Home menu.
        ElseIf strLabelLow = "styles gallery" _
            Or strLabelLow = "style gallery" Then
            'Removes the defaults from the style gallery.
            arrDefaultStyleGallery = Split(strDefaultStyleGallery, ", ")
            For lngSpec = LBound(arrDefaultStyleGallery) _
                To UBound(arrDefaultStyleGallery)
                strStyle = arrDefaultStyleGallery(lngSpec)
                objActiveDocument.Styles(strStyle).QuickStyle = False
            Next lngSpec
            'Adds styles to the style gallery.
            For lngSpec = 1 To UBound(arrSpecs)
                strStyle = arrSpecs(lngSpec)
                With objActiveDocument.Styles(strStyle)
                    .QuickStyle = True ' True means include in the gallery.
                    .UnhideWhenUsed = False ' False means never hidden.
                    .Visibility = False ' False (sic) means always visible.
                    .Priority = lngSpec
                End With
            Next lngSpec
'Lists  '
'-------'Or if the line begins with a list template name, then...
        ElseIf Right(strLabelLow, 5) = " list" _
            Or Right(strLabelLow, 11) = " list style" _
            Or Right(strLabelLow, 12) = " list styles" Then
            
            'Saves the list name.
            If Right(strLabelLow, 4) = "list" _
                And InStr(strLabelLow, " multi") <> 0 Then
                strList = Left(strLabel, InStr(strLabelLow, " multi") - 1)
            Else
                strList = Left(strLabel, InStr(strLabelLow, " list") - 1)
            End If
            'Counts the styles in the list.
            lngLevels = UBound(arrSpecs)
            If lngLevels > 9 Then lngLevels = 9
            
            'Saves the style names in an array (_, 1).
            Erase arrList
            ReDim arrList(1 To lngLevels, 1 To 27)
            For lngLevel = 1 To lngLevels
                arrList(lngLevel, 1) = arrSpecs(lngLevel)
            Next lngLevel
            
'            Key to specifications in the array.
'             2. NumberFormat
'             3. TrailingCharacter
'             4. NumberStyle
'             5. NumberPosition
'             6. Alignment
'             7. TextPosition
'             8. TabPosition
'             9. ResetOnHigher
'            10. StartAt
'            12. Font.Bold
'            13. Font.Italic
'            14. Font.StrikeThrough
'            15. Font.Subscript
'            16. Font.Superscript
'            17. Font.Shadow
'            18. Font.Outline
'            19. Font.Emboss
'            20. Font.Engrave
'            21. Font.AllCaps
'            22. Font.Hidden
'            23. Font.Underline
'            24. Font.Color
'            25. Font.Size
'            26. Font.Animation
'            27. Font.DoubleStrikeThrough
'            11. Font.Name
'             1. LinkedStyle
            
            'Reads each line in the document, looking for specs for the list.
            For lngListPara = LBound(arrParas) To UBound(arrParas)
                strPara = arrParas(lngListPara)
                'Saves the specifications on each line (between commas).
                arrSpecs = Split(strPara, ", ")
                'Saves the first specification on each line.
                strLabel = arrSpecs(0)
                strLabelLow = LCase(strLabel)
                
                'If a line has defaults for the list...
                If Left(strLabel, Len(strList)) = strList _
                    And (Right(strLabelLow, 9) = " defaults" _
                    Or Right(strLabelLow, 8) = " default") Then
                    For lngLevel = 1 To lngLevels
                        '...saves the specs in the array...
                        vbaDefineList arrList, lngLevel, arrSpecs
                        '...and applies any style specs.
                        vbaDefineOneStyle arrList(lngLevel, 1), arrSpecs
                    Next lngLevel
                
                'If a line has specs for a style...
                ElseIf Right(strLabelLow, 5) = "style" Then
                    strStyle = Left(strLabel, InStr(strLabelLow, " style") - 1)
                    '...and if the style is in the list...
                    For lngLevel = 1 To lngLevels
                        If arrList(lngLevel, 1) = strStyle Then
                            '...saves the specs in the array.
                            vbaDefineList arrList, lngLevel, arrSpecs
                        End If
                    Next lngLevel
                End If
            Next lngListPara
            
            'Reapplies any style specs, in case they override the list defaults.
            For lngListPara = LBound(arrParas) To UBound(arrParas)
                strPara = arrParas(lngListPara)
                'Saves the specifications on each line (between commas).
                arrSpecs = Split(strPara, ", ")
                'Saves the first specification on each line.
                strLabel = arrSpecs(0)
                strLabelLow = LCase(strLabel)
                
                'If a line has specs for a style...
                If Right(strLabelLow, 5) = "style" Then
                    strStyle = Left(strLabel, InStr(strLabelLow, " style") - 1)
                    '...and if the style is in the list...
                    For lngLevel = 1 To lngLevels
                        If arrList(lngLevel, 1) = strStyle Then
                            '...applies any style specs again.
                            vbaDefineOneStyle strStyle, arrSpecs
                        End If
                    Next lngLevel
                End If
            Next lngListPara
            
            'Adds a list template if it doesn't exist.
            If vbaStyleExists(strList, objActiveDocument) Then
                Set objListTemplate = objActiveDocument.ListTemplates(strList)
            Else
                Set objListTemplate = _
                    objActiveDocument.ListTemplates.Add(True, CStr(strList))
            End If
            'Applies the list template specifications.
            For lngLevel = 1 To lngLevels
                With objListTemplate.ListLevels(lngLevel)
                    .NumberFormat = ""
                    .NumberFormat = arrList(lngLevel, 2)
                    With .Font
                        If arrList(lngLevel, 11) <> "" Then
                            .Name = arrList(lngLevel, 11)
                        End If
                        If arrList(lngLevel, 12) <> "" Then
                            .Bold = arrList(lngLevel, 12)
                        End If
                        If arrList(lngLevel, 13) <> "" Then
                            .Italic = arrList(lngLevel, 13)
                        End If
                        If arrList(lngLevel, 24) <> "" Then
                            .Color = arrList(lngLevel, 24)
                        End If
                    End With
                    If arrList(lngLevel, 3) <> "" Then
                        .TrailingCharacter = arrList(lngLevel, 3)
                    End If
                    If arrList(lngLevel, 4) <> "" Then
                        .NumberStyle = arrList(lngLevel, 4)
                    End If
                    If arrList(lngLevel, 5) <> "" Then
                        .NumberPosition = arrList(lngLevel, 5)
                    End If
                    If arrList(lngLevel, 6) <> "" Then
                        .Alignment = arrList(lngLevel, 6)
                    End If
                    If arrList(lngLevel, 7) <> "" Then
                        .TextPosition = arrList(lngLevel, 7)
                    End If
'                    .TabPosition = wdUndefined      'Default
'                    .ResetOnHigher = (lngLevel - 1) 'Default
                    If arrList(lngLevel, 10) <> "" Then
                        .StartAt = arrList(lngLevel, 10)
                    End If
                    If arrList(lngLevel, 1) <> "" Then
                        .LinkedStyle = arrList(lngLevel, 1)
                    End If
                    'The linked style name must be set last.
                End With
            Next lngLevel
'Samples'
'-------'Inserts a sample of each defined style.
        ElseIf Left(strLabelLow, 13) = "insert sample" Then
            objActiveDocument.Characters.Last.Select
            With Selection
                .Collapse wdCollapseEnd
                .TypeParagraph
                .ClearFormatting
                vbaInsertSampleText arrStyles
                .EndKey Unit:=wdStory
            End With
        End If
    Next lngPara
    Set objListTemplate = Nothing
    Set objActiveDocument = Nothing
End Sub

Function vbaSaveParagraphsInAnArray(ByVal varDesc As Variant) _
    As String()
    Dim arrParas() As String, lngPara As Long, strPara As String
    'Saves paragraphs in an array.
    arrParas = Split(varDesc, vbCrLf)
    'Cleans up the text in the paragraphs.
    For lngPara = LBound(arrParas) To UBound(arrParas)
        strPara = CStr(arrParas(lngPara))
        'Removes spaces and a period at the end of a paragraph.
        strPara = Trim(strPara)
        If Right(strPara, 1) = "." Then
            strPara = Left(strPara, Len(strPara) - 1)
        End If
        'Replaces manual line breaks with a space.
        strPara = Replace(strPara, Chr(11), " ")
        'Removes extra spaces.
        Do While InStr(strPara, "  ") <> 0
            strPara = Replace(strPara, "  ", " ")
        Loop
        'Saves some text instead of an empty line.
        If strPara = "" Then strPara = "[empty line]"
        arrParas(lngPara) = strPara
    Next lngPara
    vbaSaveParagraphsInAnArray = arrParas
End Function

Private Function vbaStyleExists(ByVal strStyle As String, _
    ByVal objDocument As Document) As Boolean
    Dim objStyle As Style, objListTemplate As ListTemplate
    On Error Resume Next
    'Checks whether a style exists with the name.
    Set objStyle = objDocument.Styles(strStyle)
    vbaStyleExists = Not objStyle Is Nothing
    'If not...
    If Not vbaStyleExists Then
        'then checks whether a list exists with the name.
        Set objListTemplate = objDocument.ListTemplates(strStyle)
        vbaStyleExists = Not objListTemplate Is Nothing
    End If
    Set objStyle = Nothing: Set objListTemplate = Nothing
End Function

Private Sub vbaDefineOneStyle(ByVal strStyle As String, arrSpecs() As String)
    
    Dim lngType As Long, lngSpec As Long, strSpec As String, dblSpec As Double
    Dim strSpecLow As String, dblSpec2 As Double
'    Dim objStyle As Object, objFont As Object, objFormat As Object
    Dim objStyle As Style, objFont As Font, objFormat As ParagraphFormat
    Dim objActiveDocument As Document: Set objActiveDocument = ActiveDocument
    
    Set objStyle = objActiveDocument.Styles(strStyle)
    Set objFont = objStyle.Font
    lngType = objStyle.Type
    If lngType = wdStyleTypeParagraph Then
        Set objFormat = objStyle.ParagraphFormat
    End If
    
    'Looks at each specification in the array.
    For lngSpec = 1 To UBound(arrSpecs)
        strSpec = arrSpecs(lngSpec)
        strSpecLow = LCase(strSpec)
        dblSpec = Val(strSpec)
    
'-------'Applies specifications to both character and paragraph styles.
        If lngType = wdStyleTypeCharacter Or lngType = wdStyleTypeParagraph Then
            '(skipped) objStyle.AutomaticallyUpdate = False
            If Left(strSpec, 1) = "[" _
                Or Right(strSpec, 1) = "]" Then
                'Ignores specifications in square brackets.
            ElseIf Left(strSpecLow, 9) = "based on " _
                Or Right(strSpecLow, 11) = " base style" Then '-- based on style
                If Left(strSpecLow, 9) = "based on " Then
                    strSpec = Right(strSpec, Len(strSpec) - 9)
                    If Right(strSpecLow, 6) = " style" _
                        And Right(strSpecLow, 9) <> " no style" Then
                        strSpec = Left(strSpec, Len(strSpec) - 6)
                    End If
                ElseIf Right(strSpecLow, 11) = " base style" Then
                    strSpec = Left(strSpec, Len(strSpec) - 11)
                End If
                strSpecLow = LCase(strSpec)
                If strSpecLow = "no style" _
                    Or strSpecLow = "underlying properties" Then
                    objStyle.BaseStyle = ""
                Else
                    objStyle.BaseStyle = strSpec
                End If
            ElseIf Left(strSpecLow, 12) = "followed by " _
                Or Right(strSpecLow, 21) = " next paragraph style" _
                Or Right(strSpecLow, 11) = " next style" Then '- following style
                If Left(strSpecLow, 12) = "followed by " Then
                    strSpec = Right(strSpec, Len(strSpec) - 12)
                ElseIf Right(strSpecLow, 21) = " next paragraph style" Then
                    strSpec = Left(strSpec, Len(strSpec) - 21)
                ElseIf Right(strSpecLow, 11) = " next style" Then
                    strSpec = Left(strSpec, Len(strSpec) - 11)
                End If
                strSpecLow = LCase(strSpec)
                If Right(strSpecLow, 6) = " style" Then
                    strSpec = Left(strSpec, Len(strSpec) - 6)
                End If
                strSpecLow = LCase(strSpec)
                If strSpecLow = "the same" _
                    Or strSpecLow = "same" Then
                    objStyle.NextParagraphStyle = strStyle
                Else
                    objStyle.NextParagraphStyle = strSpec
                End If
            ElseIf Left(strSpecLow, 13) = "space between" _
                Or Left(strSpecLow, 9) = "add space" _
                Then '-------------------------------------------- space between
                objStyle.NoSpaceBetweenParagraphsOfSameStyle = False
            ElseIf Left(strSpecLow, 16) = "no space between" _
                Or Left(strSpecLow, 23) = "don't add space between" _
                Or Left(strSpecLow, 23) = "don’t add space between" _
                Or Left(strSpecLow, 24) = "do not add space between" Then
                objStyle.NoSpaceBetweenParagraphsOfSameStyle = True
            
            ElseIf Right(strSpecLow, 5) = " font" _
                And Right(strSpecLow, 12) <> " bullet font" _
                And Right(strSpecLow, 12) <> " number font" _
                And Right(strSpecLow, 12) <> " letter font" _
                And Right(strSpecLow, 13) <> " numeral font" _
                And Right(strSpecLow, 13) <> " bullets font" _
                And Right(strSpecLow, 13) <> " numbers font" _
                And Right(strSpecLow, 13) <> " letters font" Then '--- font name
                If Right(strSpecLow, 13) = " bullets font" _
                    Or Right(strSpecLow, 13) = " numbers font" _
                    Or Right(strSpecLow, 13) = " letters font" _
                    Or Right(strSpecLow, 13) = " numeral font" Then
                    strSpec = Left(strSpec, Len(strSpec) - 13)
                ElseIf Right(strSpecLow, 12) = " bullet font" _
                    Or Right(strSpecLow, 12) = " number font" _
                    Or Right(strSpecLow, 12) = " letter font" Then
                    strSpec = Left(strSpec, Len(strSpec) - 12)
                ElseIf Right(strSpecLow, 5) = " font" Then
                    strSpec = Left(strSpec, Len(strSpec) - 5)
                End If
                strSpecLow = LCase(strSpec)
                If strSpecLow = "body" Then
                    strSpec = "+Body"
                ElseIf strSpecLow = "headings" _
                    Or strSpecLow = "heading" Then
                    strSpec = "+Headings"
                End If
                If strSpecLow <> "default" Then
                    objFont.Name = strSpec
                End If
            ElseIf Right(strSpecLow, 4) = "size" Then '-------------------- size
                objFont.Size = Val(strSpec)
            ElseIf strSpecLow = "bold" Then '------------------------------ bold
                objFont.Bold = True
            ElseIf strSpecLow = "not bold" _
                Or strSpecLow = "no bold" Then
                objFont.Bold = False
            ElseIf strSpecLow = "italic" Then '-------------------------- italic
                objFont.Italic = True
            ElseIf strSpecLow = "not italic" _
                Or strSpecLow = "no italic" Then
                objFont.Italic = False
            ElseIf strSpecLow = "bold and italic" _
                Or strSpecLow = "italic and bold" _
                Or strSpecLow = "bold italic" _
                Or strSpecLow = "italic bold" Then
                objFont.Bold = True
                objFont.Italic = True
            ElseIf strSpecLow = "no bold or italic" _
                Or strSpecLow = "no bold and italic" _
                Or strSpecLow = "no italic or bold" _
                Or strSpecLow = "no italic and bold" _
                Or strSpecLow = "not bold or italic" _
                Or strSpecLow = "not bold and italic" _
                Or strSpecLow = "not italic or bold" _
                Or strSpecLow = "not italic and bold" Then
                objFont.Bold = False
                objFont.Italic = False
            ElseIf InStr(strSpecLow, "underline") <> 0 Then '--------- underline
                If strSpecLow = "no underline" _
                    Or strSpecLow = "not underlined" _
                    Or strSpecLow = "underline none" Then
                    objFont.Underline = wdUnderlineNone
                ElseIf strSpecLow = "underline dash" _
                    Or strSpecLow = "dash underline" Then
                    objFont.Underline = wdUnderlineDash
                ElseIf strSpecLow = "underline dash heavy" _
                    Or strSpecLow = "heavy dash underline" Then
                    objFont.Underline = wdUnderlineDashHeavy
                ElseIf strSpecLow = "underline dash long" _
                    Or strSpecLow = "long dash underline" Then
                    objFont.Underline = wdUnderlineDashLong
                ElseIf strSpecLow = "underline dash long heavy" _
                    Or strSpecLow = "heavy long dash underline" _
                    Or strSpecLow = "long heavy dash underline" Then
                    objFont.Underline = wdUnderlineDashLongHeavy
                ElseIf strSpecLow = "underline dot dash" _
                    Or strSpecLow = "dot dash underline" _
                    Or strSpecLow = "dot-dash underline" Then
                    objFont.Underline = wdUnderlineDotDash
                ElseIf strSpecLow = "underline dot dash heavy" _
                    Or strSpecLow = "heavy dot dash underline" _
                    Or strSpecLow = "heavy dot-dash underline" Then
                    objFont.Underline = wdUnderlineDotDashHeavy
                ElseIf strSpecLow = "underline dot dot dash" _
                    Or strSpecLow = "dot dot dash underline" _
                    Or strSpecLow = "dot-dot-dash underline" Then
                    objFont.Underline = wdUnderlineDotDotDash
                ElseIf strSpecLow = "underline dot dot dash heavy" _
                    Or strSpecLow = "heavy dot dot dash underline" _
                    Or strSpecLow = "heavy dot-dot-dash underline" Then
                    objFont.Underline = wdUnderlineDotDotDashHeavy
                ElseIf strSpecLow = "underline dotted" _
                    Or strSpecLow = "dotted underline" _
                    Or strSpecLow = "dot underline" Then
                    objFont.Underline = wdUnderlineDotted
                ElseIf strSpecLow = "underline dotted heavy" _
                    Or strSpecLow = "heavy dotted underline" _
                    Or strSpecLow = "dotted heavy underline" _
                    Or strSpecLow = "heavy dot underline" Then
                    objFont.Underline = wdUnderlineDottedHeavy
                ElseIf strSpecLow = "underline double" _
                    Or strSpecLow = "double underline" _
                    Or strSpecLow = "double underlined" Then
                    objFont.Underline = wdUnderlineDouble
                ElseIf strSpecLow = "underline single" _
                    Or strSpecLow = "single underline" _
                    Or strSpecLow = "single underlined" Then
                    objFont.Underline = wdUnderlineSingle
                ElseIf strSpecLow = "underline thick" _
                    Or strSpecLow = "thick underline" _
                    Or strSpecLow = "thick underlined" Then
                    objFont.Underline = wdUnderlineThick
                ElseIf strSpecLow = "underline wavy" _
                    Or strSpecLow = "wavy underline" Then
                    objFont.Underline = wdUnderlineWavy
                ElseIf strSpecLow = "underline wavy double" _
                    Or strSpecLow = "double wavy underline" _
                    Or strSpecLow = "wavy double underline" Then
                    objFont.Underline = wdUnderlineWavyDouble
                ElseIf strSpecLow = "underline wavy heavy" _
                    Or strSpecLow = "heavy wavy underline" _
                    Or strSpecLow = "wavy heavy underline" Then
                    objFont.Underline = wdUnderlineWavyHeavy
                ElseIf strSpecLow = "underline words" _
                    Or strSpecLow = "underline words only" _
                    Or strSpecLow = "word underline" _
                    Or strSpecLow = "words underlined" Then
                    objFont.Underline = wdUnderlineWords
                End If
            '(skipped) objFont.SmallCaps = False '------------------- small caps
            ElseIf strSpecLow = "uppercase" _
                Or strSpecLow = "uppercase letters" _
                Or strSpecLow = "all uppercase" _
                Or strSpecLow = "all uppercase letters" _
                Or strSpecLow = "caps" _
                Or strSpecLow = "all caps" _
                Or strSpecLow = "capitals" _
                Or strSpecLow = "all capitals" _
                Or strSpecLow = "capital letters" _
                Or strSpecLow = "all capital letters" _
                Or strSpecLow = "capitalize" _
                Or strSpecLow = "capitalized" _
                Or strSpecLow = "capitalized letters" _
                Or strSpecLow = "all capitalized" _
                Or strSpecLow = "all capital letters" _
                Then '------------------------------------------------- all caps
                objFont.AllCaps = True
            ElseIf strSpecLow = "not uppercase" _
                Or strSpecLow = "not uppercase letters" _
                Or strSpecLow = "not all uppercase" _
                Or strSpecLow = "not all uppercase letters" _
                Or strSpecLow = "not caps" _
                Or strSpecLow = "not all caps" _
                Or strSpecLow = "not capitals" _
                Or strSpecLow = "not all capitals" _
                Or strSpecLow = "not capital letters" _
                Or strSpecLow = "not all capital letters" _
                Or strSpecLow = "not capitalize" _
                Or strSpecLow = "not capitalized" _
                Or strSpecLow = "not capitalized letters" _
                Or strSpecLow = "not all capitalized" _
                Or strSpecLow = "not all capital letters" _
                Or strSpecLow = "no all caps" Then
                objFont.AllCaps = False
            ElseIf Right(strSpecLow, 5) = "color" _
                And Right(strSpecLow, 12) <> "bullet color" _
                And Right(strSpecLow, 12) <> "number color" _
                And Right(strSpecLow, 12) <> "letter color" _
                And Right(strSpecLow, 13) <> "numeral color" Then '------- color
                strSpec = Split(strSpec, " ")(0)
                strSpecLow = LCase(strSpec)
                dblSpec = -1
                If Left(strSpec, 1) = "#" Then
                    strSpec = Right(strSpec, Len(strSpec) - 1)
                    strSpec = Right(strSpec, 2) & Mid(strSpec, 3, 2) _
                        & Left(strSpec, 2)
                    dblSpec = Val("&H" & strSpec)
                ElseIf strSpecLow = "automatic" Or strSpecLow = "auto" _
                    Or strSpecLow = "no" Then
                    dblSpec = wdColorAutomatic
                ElseIf strSpecLow = "black" Then
                    dblSpec = wdColorBlack
                ElseIf strSpecLow = "white" Then
                    dblSpec = wdColorWhite
                ElseIf strSpecLow = "blue" Then
                    dblSpec = wdColorBlue
                End If
                If dblSpec <> -1 Then
                    objFont.Color = dblSpec
                End If
            ElseIf strSpecLow = "normal letterspacing" _
                Or strSpecLow = "normal letter spacing" _
                Or strSpecLow = "normal letter-spacing" _
                Or strSpecLow = "normal character spacing" _
                Or strSpecLow = "normal character-spacing" _
                Or strSpecLow = "no letterspacing" _
                Or strSpecLow = "no letter spacing" _
                Or strSpecLow = "no letter-spacing" _
                Or strSpecLow = "no character spacing" _
                Or strSpecLow = "no character-spacing" _
                Or strSpecLow = "not letterspaced" _
                Or strSpecLow = "not letter spaced" _
                Or strSpecLow = "not letter-spaced" _
                Or strSpecLow = "not character spaced" _
                Or strSpecLow = "not character-spaced" Then '----- letterspacing
                objFont.Spacing = 0
            ElseIf Right(strSpecLow, 13) = "letterspacing" _
                Or Right(strSpecLow, 14) = "letter spacing" _
                Or Right(strSpecLow, 14) = "letter-spacing" _
                Or Right(strSpecLow, 17) = "character spacing" _
                Or Right(strSpecLow, 17) = "character-spacing" Then
                objFont.Spacing = dblSpec
            ElseIf strSpecLow = "kerning" Then '------------------------ kerning
                objFont.Kerning = 8
            ElseIf strSpecLow = "no kerning" Or strSpecLow = "not kerned" Then
                objFont.Kerning = 0
            End If
        End If
        
'-------'Applies specifications to paragraph styles.
        If lngType = wdStyleTypeParagraph Then
            If Right(strSpecLow, 11) = "left indent" Then '--------- indents
                objFormat.LeftIndent = InchesToPoints(dblSpec)
            ElseIf Right(strSpecLow, 12) = "right indent" Then
                objFormat.RightIndent = InchesToPoints(dblSpec)
            '(postponed to the end) objFormat.SpaceBefore = dblSpec
            '(skipped) objFormat.SpaceBeforeAuto = False
            '(postponed to the end) objFormat.SpaceAfter = dblSpec
            '(skipped) objFormat.SpaceAfterAuto = False
            ElseIf Right(strSpecLow, 12) = "line spacing" Then 'line spacing
                If Split(strSpecLow, " ")(1) = "pt" _
                    Or Split(strSpecLow, " ")(1) = "pt." _
                    Or Split(strSpecLow, " ")(1) = "point" _
                    Or Split(strSpecLow, " ")(1) = "points" Then
                    objFormat.LineSpacingRule = wdLineSpaceExactly
                    objFormat.LineSpacing = dblSpec
                ElseIf Split(strSpecLow, " ")(0) = "exact" _
                    Or Split(strSpecLow, " ")(0) = "exactly" Then
                    dblSpec = Val(Split(strSpec, " ")(1))
                    objFormat.LineSpacingRule = wdLineSpaceExactly
                    objFormat.LineSpacing = dblSpec
                ElseIf Split(strSpecLow, " ")(1) = "least" Then
                    dblSpec = Val(Split(strSpec, " ")(2))
                    objFormat.LineSpacingRule = wdLineSpaceAtLeast
                    objFormat.LineSpacing = dblSpec
                ElseIf Split(strSpecLow, " ")(0) = "single" Then
                    objFormat.LineSpacingRule = wdLineSpaceSingle
                Else
                    objFormat.LineSpacingRule = wdLineSpaceMultiple
                    objFormat.LineSpacing = LinesToPoints(dblSpec)
                End If
            ElseIf strSpecLow = "left aligned" Or strSpecLow = "left align" _
                Or strSpecLow = "aligned left" Or strSpecLow = "align left" _
                Or strSpecLow = "right aligned" Or strSpecLow = "right align" _
                Or strSpecLow = "aligned right" Or strSpecLow = "align right" _
                Or strSpecLow = "centered" Or strSpecLow = "center" _
                Or strSpecLow = "center align" Or strSpecLow = "align center" _
                Or strSpecLow = "center aligned" _
                Or strSpecLow = "aligned center" _
                Or InStr(strSpecLow, "justify") <> 0 _
                Or InStr(strSpecLow, "justified") <> 0 _
                Then '-------------------------------------------- alignment
                dblSpec = wdAlignParagraphLeft
                If strSpecLow = "right aligned" _
                    Or strSpecLow = "right align" _
                    Or strSpecLow = "aligned right" _
                    Or strSpecLow = "align right" Then
                    dblSpec = wdAlignParagraphRight
                ElseIf strSpecLow = "centered" Or strSpecLow = "center" _
                    Or strSpecLow = "center aligned" _
                    Or strSpecLow = "aligned center" _
                    Or strSpecLow = "center align" _
                    Or strSpecLow = "align center" Then
                    dblSpec = wdAlignParagraphCenter
                ElseIf InStr(strSpecLow, "justify") _
                    Or InStr(strSpecLow, "justified") Then
                    dblSpec = wdAlignParagraphJustify
                End If
                objFormat.Alignment = dblSpec
            ElseIf strSpecLow = "widow/orphan control" _
                Or strSpecLow = "orphan/widow control" _
                Or strSpecLow = "widow and orphan control" _
                Or strSpecLow = "orphan and widow control" _
                Or strSpecLow = "widow control" _
                Or strSpecLow = "orphan control" Then '------- widow control
                objFormat.WidowControl = True
            ElseIf strSpecLow = "no widow/orphan control" _
                Or strSpecLow = "no orphan/widow control" _
                Or strSpecLow = "no widow and orphan control" _
                Or strSpecLow = "no orphan and widow control" _
                Or strSpecLow = "no widow or orphan control" _
                Or strSpecLow = "no orphan or widow control" _
                Or strSpecLow = "no widow control" _
                Or strSpecLow = "no orphan control" Then
                objFormat.WidowControl = False
            ElseIf Left(strSpecLow, 14) = "keep with next" _
                Or Left(strSpecLow, 24) = "keep paragraph with next" _
                Or Left(strSpecLow, 30) = "keep the paragraph with the ne" _
                Or Left(strSpecLow, 22) = "no page break after" _
                Or Left(strSpecLow, 22) = "no page break below" _
                Or Left(strSpecLow, 28) = "don't allow page break after" _
                Or Left(strSpecLow, 30) = "don't allow a page break after" _
                Or Left(strSpecLow, 28) = "don't allow page break below" _
                Or Left(strSpecLow, 30) = "don't allow a page break below" _
                Or Left(strSpecLow, 28) = "don’t allow page break after" _
                Or Left(strSpecLow, 30) = "don’t allow a page break after" _
                Or Left(strSpecLow, 28) = "don’t allow page break below" _
                Or Left(strSpecLow, 30) = "don’t allow a page break below" _
                Or Left(strSpecLow, 29) = "do not allow page break after" _
                Or Left(strSpecLow, 30) = "do not allow a page break afte" _
                Or Left(strSpecLow, 29) = "do not allow page break below" _
                Or Left(strSpecLow, 30) = "do not allow a page break belo" _
                Then '--------------------------------------- keep with next
                objFormat.KeepWithNext = True
            ElseIf Left(strSpecLow, 17) = "no keep with next" _
                Or Left(strSpecLow, 27) = "no keep paragraph with next" _
                Or Left(strSpecLow, 30) = "no keep the paragraph with nex" _
                Or Left(strSpecLow, 20) = "don't keep with next" _
                Or Left(strSpecLow, 30) = "don't keep paragraph with next" _
                Or Left(strSpecLow, 30) = "don't keep the paragraph with " _
                Or Left(strSpecLow, 20) = "don’t keep with next" _
                Or Left(strSpecLow, 30) = "don’t keep paragraph with next" _
                Or Left(strSpecLow, 30) = "don’t keep the paragraph with " _
                Or Left(strSpecLow, 21) = "do not keep with next" _
                Or Left(strSpecLow, 30) = "do not keep paragraph with nex" _
                Or Left(strSpecLow, 30) = "do not keep the paragraph with" _
                Or Left(strSpecLow, 22) = "allow page break after" _
                Or Left(strSpecLow, 24) = "allow a page break after" _
                Or Left(strSpecLow, 22) = "allow page break below" _
                Or Left(strSpecLow, 24) = "allow a page break below" _
                Then
                objFormat.KeepWithNext = False
            ElseIf Left(strSpecLow, 13) = "keep together" _
                Or Left(strSpecLow, 19) = "keep lines together" _
                Or Left(strSpecLow, 29) = "keep paragraph lines together" _
                Or Left(strSpecLow, 30) = "keep the paragraph lines toget" _
                Or Left(strSpecLow, 30) = "keep the paragraph lines on th" _
                Or Left(strSpecLow, 21) = "keep on the same page" _
                Or Left(strSpecLow, 27) = "keep lines on the same page" _
                Or Left(strSpecLow, 30) = "keep paragraph lines on the sa" _
                Then '---------------------------------- keep lines together
                objFormat.KeepTogether = True
            ElseIf strSpecLow = "no keep together" _
                Or strSpecLow = "no keep lines together" _
                Or strSpecLow = "no keep paragraph lines together" _
                Or strSpecLow = "don't keep together" _
                Or strSpecLow = "don't keep lines together" _
                Or strSpecLow = "don't keep paragraph lines together" _
                Or strSpecLow = "don’t keep together" _
                Or strSpecLow = "don’t keep lines together" _
                Or strSpecLow = "don’t keep paragraph lines together" _
                Or strSpecLow = "do not keep together" _
                Or strSpecLow = "do not keep lines together" _
                Or strSpecLow = "do not keep paragraph lines together" _
                Or Left(strSpecLow, 19) = "allow page break in" _
                Or Left(strSpecLow, 21) = "allow a page break in" _
                Or Left(strSpecLow, 23) = "allow page break within" _
                Or Left(strSpecLow, 25) = "allow a page break within" _
                Then
                objFormat.KeepTogether = False
            ElseIf Left(strSpecLow, 17) = "page break before" _
                Or Left(strSpecLow, 25) = "require page break before" _
                Or Left(strSpecLow, 27) = "require a page break before" _
                Or Left(strSpecLow, 16) = "page break above" _
                Or Left(strSpecLow, 24) = "require page break above" _
                Or Left(strSpecLow, 26) = "require a page break above" _
                Then '------------------------------------ page break before
                objFormat.PageBreakBefore = True
            ElseIf Left(strSpecLow, 20) = "no page break before" _
                Or Left(strSpecLow, 28) = "no require page break before" _
                Or Left(strSpecLow, 30) = "no require a page break before" _
                Or Left(strSpecLow, 30) = "don't require page break befor" _
                Or Left(strSpecLow, 30) = "don't require a page break bef" _
                Or Left(strSpecLow, 30) = "don’t require page break befor" _
                Or Left(strSpecLow, 30) = "don’t require a page break bef" _
                Or Left(strSpecLow, 30) = "do not require page break befo" _
                Or Left(strSpecLow, 30) = "do not require a page break be" _
                Or Left(strSpecLow, 19) = "no page break above" _
                Or Left(strSpecLow, 27) = "no require page break above" _
                Or Left(strSpecLow, 29) = "no require a page break above" _
                Or Left(strSpecLow, 30) = "don't require page break above" _
                Or Left(strSpecLow, 30) = "don't require a page break abo" _
                Or Left(strSpecLow, 30) = "don’t require page break above" _
                Or Left(strSpecLow, 30) = "don’t require a page break abo" _
                Or Left(strSpecLow, 30) = "do not require page break abov" _
                Or Left(strSpecLow, 30) = "do not require a page break ab" _
                Then
                objFormat.PageBreakBefore = False
            '(skipped) objFormat.NoLineNumber = False
            '(skipped) objFormat.Hyphenation = True
            '(skipped) objFormat.FirstLineIndent = InchesToPoints(0)
            ElseIf Left(strSpecLow, 14) = "outline level " Then '----- level
                dblSpec = Val(Right(strSpec, 1))
                If dblSpec >= 1 And dblSpec <= 9 Then
                    objFormat.OutlineLevel = dblSpec
                Else
                    objFormat.OutlineLevel = wdOutlineLevelBodyText
                End If
            ElseIf strSpecLow = "no outline level" Then
                objFormat.OutlineLevel = wdOutlineLevelBodyText
            '(skipped) objFormat.CharacterUnitLeftIndent = 0
            '(skipped) objFormat.CharacterUnitRightIndent = 0
            '(skipped) objFormat.CharacterUnitFirstLineIndent = 0
            '(skipped) objFormat.LineUnitBefore = 0
            '(skipped) objFormat.LineUnitAfter = 0
            '(skipped) objFormat.MirrorIndents = False
            '(skipped) objFormat.TextboxTightWrap = wdTightNone
            '(skipped) objFormat.CollapsedByDefault = False
            ElseIf strSpecLow = "no border" _
                Or strSpecLow = "no borders" Then '----------------- borders
                With objFormat
                    .Borders(wdBorderTop).LineStyle = wdLineStyleNone
                    .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
                    .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
                    .Borders(wdBorderRight).LineStyle = wdLineStyleNone
                End With
            ElseIf Right(strSpecLow, 6) = "border" _
                Or Right(strSpecLow, 7) = "borders" Then
                If InStr(strSpecLow, "top") <> 0 Then
                    dblSpec2 = wdBorderTop
                ElseIf InStr(strSpecLow, "bottom") <> 0 Then
                    dblSpec2 = wdBorderBottom
                ElseIf InStr(strSpecLow, "left") <> 0 Then
                    dblSpec2 = wdBorderLeft
                ElseIf InStr(strSpecLow, "right") <> 0 Then
                    dblSpec2 = wdBorderRight
                End If
                If dblSpec2 <> 0 Then
                    With objFormat.Borders(dblSpec2)
                        .LineStyle = wdLineStyleSingle
                        Select Case dblSpec
                            Case 0
                                .LineWidth = wdUndefined
                            Case Is <= 0.25
                                .LineWidth = wdLineWidth025pt
                            Case Is <= 0.5
                                .LineWidth = wdLineWidth050pt
                            Case Is <= 0.75
                                .LineWidth = wdLineWidth075pt
                            Case Is <= 1
                                .LineWidth = wdLineWidth100pt
                            Case Is <= 1.5
                                .LineWidth = wdLineWidth150pt
                            Case Is <= 2.25
                                .LineWidth = wdLineWidth225pt
                            Case Is <= 3
                                .LineWidth = wdLineWidth300pt
                            Case Is <= 4.5
                                .LineWidth = wdLineWidth450pt
                            Case Is > 4.5
                                .LineWidth = wdLineWidth600pt
                        End Select
                    End With
                End If
            ElseIf strSpecLow = "no tabs" _
                Or strSpecLow = "no tab" _
                Or strSpecLow = "clear tabs" _
                Or strSpecLow = "clear tab" _
                Or strSpecLow = "clear all tabs" _
                Or strSpecLow = "remove tabs" _
                Or strSpecLow = "remove tab" _
                Or strSpecLow = "remove all tabs" Then '--------------- tabs
                objFormat.TabStops.ClearAll
            ElseIf strSpecLow = "center tab" _
                Or strSpecLow = "centered tab" Then
                With objActiveDocument.PageSetup
                    dblSpec = .PageWidth - .LeftMargin - .RightMargin
                End With
                objFormat.TabStops.Add Position:=(dblSpec / 2), _
                    Alignment:=wdAlignTabCenter, _
                    Leader:=wdTabLeaderSpaces
            ElseIf strSpecLow = "right tab" Then
                With objActiveDocument.PageSetup
                    dblSpec = .PageWidth - .LeftMargin - .RightMargin _
                        - objFormat.RightIndent
                End With
                objFormat.TabStops.Add Position:=dblSpec, _
                    Alignment:=wdAlignTabRight, _
                    Leader:=wdTabLeaderSpaces
            ElseIf Right(strSpecLow, 3) = "tab" Then
                dblSpec2 = wdAlignTabLeft
                If InStr(strSpecLow, "center") <> 0 Then
                    dblSpec2 = wdAlignTabCenter
                ElseIf InStr(strSpecLow, "right") <> 0 Then
                    dblSpec2 = wdAlignTabRight
                ElseIf InStr(strSpecLow, "decimal") <> 0 Then
                    dblSpec2 = wdAlignTabDecimal
                End If
                objFormat.TabStops.Add Position:=InchesToPoints(dblSpec), _
                    Alignment:=dblSpec2, _
                    Leader:=wdTabLeaderSpaces
            End If
        End If
    Next lngSpec
    
'-------'Applies some specifications last, which seems to be necessary.
    
    'Again, looks at each specification in the array.
    For lngSpec = 1 To UBound(arrSpecs)
        strSpec = arrSpecs(lngSpec)
        strSpecLow = LCase(strSpec)
        dblSpec = Val(strSpec)
        
        'Applies the space before and after paragraphs.
        If lngType = wdStyleTypeParagraph Then
            If ((Right(strSpecLow, 5) = "after" _
                Or Right(strSpecLow, 5) = "below") _
                And InStr(strSpecLow, "page break") = 0) _
                Or Right(strSpecLow, 15) = "after paragraph" _
                Or Right(strSpecLow, 15) = "below paragraph" _
                Or Right(strSpecLow, 16) = "after paragraphs" _
                Or Right(strSpecLow, 16) = "below paragraphs" _
                Or Right(strSpecLow, 19) = "after the paragraph" _
                Or Right(strSpecLow, 19) = "below the paragraph" _
                Or Right(strSpecLow, 20) = "after the paragraphs" _
                Or Right(strSpecLow, 20) = "below the paragraphs" _
                Or Right(strSpecLow, 20) = "after each paragraph" _
                Or Right(strSpecLow, 20) = "below each paragraph" _
                Then '------------------------------------------ space after
                objFormat.SpaceAfter = dblSpec
            '(skipped) objFormat.SpaceAfterAuto = False
            ElseIf ((Right(strSpecLow, 6) = "before" _
                Or Right(strSpecLow, 5) = "above") _
                And InStr(strSpecLow, "page break") = 0) _
                Or Right(strSpecLow, 16) = "before paragraph" _
                Or Right(strSpecLow, 15) = "above paragraph" _
                Or Right(strSpecLow, 17) = "before paragraphs" _
                Or Right(strSpecLow, 16) = "above paragraphs" _
                Or Right(strSpecLow, 20) = "before the paragraph" _
                Or Right(strSpecLow, 19) = "above the paragraph" _
                Or Right(strSpecLow, 21) = "before the paragraphs" _
                Or Right(strSpecLow, 20) = "above the paragraphs" _
                Or Right(strSpecLow, 21) = "before each paragraph" _
                Or Right(strSpecLow, 20) = "above each paragraph" _
                Then '----------------------------------------- space before
                objFormat.SpaceBefore = dblSpec
            '(skipped) objFormat.SpaceBeforeAuto = False
            End If
        End If
    Next lngSpec
    
    Set objStyle = Nothing: Set objFont = Nothing: Set objFormat = Nothing
    Set objActiveDocument = Nothing
End Sub

Private Sub vbaDefineList(ByRef arrList() As Variant, ByVal lngLevel As Long, _
    arrSpecs() As String)
    
    Dim lngSpec As Long, strSpec As String, strSpecLow As String
    Dim dblSpec As Double
    
    'Looks at each specification on the line.
    For lngSpec = LBound(arrSpecs) To UBound(arrSpecs)
        strSpec = arrSpecs(lngSpec)
        strSpecLow = LCase(strSpec)
        dblSpec = Val(strSpec)
        
'        Key to specifications in the array.
'         2. NumberFormat
'         3. TrailingCharacter
'         4. NumberStyle
'         5. NumberPosition
'         6. Alignment
'         7. TextPosition
'         8. TabPosition
'         9. ResetOnHigher
'        10. StartAt
'        12. Font.Bold
'        13. Font.Italic
'        14. Font.StrikeThrough
'        15. Font.Subscript
'        16. Font.Superscript
'        17. Font.Shadow
'        18. Font.Outline
'        19. Font.Emboss
'        20. Font.Engrave
'        21. Font.AllCaps
'        22. Font.Hidden
'        23. Font.Underline
'        24. Font.Color
'        25. Font.Size
'        26. Font.Animation
'        27. Font.DoubleStrikeThrough
'        11. Font.Name
'         1. LinkedStyle
        
        'Saves whether no bullet or number is specified.
        If Left(strSpecLow, 9) = "no number" _
            Or Left(strSpecLow, 9) = "no bullet" _
            Or Left(strSpecLow, 9) = "no letter" _
            Or Left(strSpecLow, 10) = "no numeral" _
            Or Left(strSpecLow, 9) = """"" number" _
            Or Left(strSpecLow, 9) = """"" bullet" _
            Or Left(strSpecLow, 9) = """"" letter" _
            Or Left(strSpecLow, 10) = """"" numeral" Then
            arrList(lngLevel, 2) = ""
            arrList(lngLevel, 4) = wdListNumberStyleNone
        
        'Saves whether a tab or space or nothing follows (spec 3).
        ElseIf Right(strSpecLow, 12) = "after bullet" _
            Or Right(strSpecLow, 14) = "follows bullet" _
            Or Right(strSpecLow, 16) = "following bullet" _
            Or Right(strSpecLow, 13) = "after bullets" _
            Or Right(strSpecLow, 15) = "follows bullets" _
            Or Right(strSpecLow, 17) = "following bullets" _
            Or Right(strSpecLow, 12) = "after number" _
            Or Right(strSpecLow, 14) = "follows number" _
            Or Right(strSpecLow, 16) = "following number" _
            Or Right(strSpecLow, 13) = "after numbers" _
            Or Right(strSpecLow, 15) = "follows numbers" _
            Or Right(strSpecLow, 17) = "following numbers" _
            Or Right(strSpecLow, 12) = "after letter" _
            Or Right(strSpecLow, 14) = "follows letter" _
            Or Right(strSpecLow, 16) = "following letter" _
            Or Right(strSpecLow, 13) = "after letters" _
            Or Right(strSpecLow, 15) = "follows letters" _
            Or Right(strSpecLow, 17) = "following letters" _
            Or Right(strSpecLow, 13) = "after numeral" _
            Or Right(strSpecLow, 15) = "follows numeral" _
            Or Right(strSpecLow, 17) = "following numeral" _
            Or Right(strSpecLow, 14) = "after numerals" _
            Or Right(strSpecLow, 16) = "follows numerals" _
            Or Right(strSpecLow, 18) = "following numerals" Then
            If Split(strSpecLow, " ")(0) = "one" _
                Or Split(strSpecLow, " ")(0) = "a" _
                Or Split(strSpecLow, " ")(0) = "only" Then
                strSpecLow = Split(strSpecLow, " ")(1)
            Else
                strSpecLow = Split(strSpecLow, " ")(0)
            End If
            dblSpec = wdTrailingSpace
            If strSpecLow = "tab" _
                Or strSpecLow = "tabs" Then
                dblSpec = wdTrailingTab
            ElseIf strSpecLow = "nothing" _
                Or strSpecLow = "no" Then
                dblSpec = wdTrailingNone
            End If
            arrList(lngLevel, 3) = dblSpec
        
        'Saves the font for the bullet or number (spec 11).
        ElseIf Right(strSpecLow, 12) = " bullet font" _
            Or Right(strSpecLow, 12) = " number font" _
            Or Right(strSpecLow, 12) = " letter font" _
            Or Right(strSpecLow, 13) = " numeral font" Then
            strSpec = Trim(Left(strSpec, Len(strSpec) - 12))
            strSpecLow = LCase(strSpec)
            If strSpecLow = "body" Then
                strSpec = "+Body"
            ElseIf strSpecLow = "headings" _
                Or strSpecLow = "heading" Then
                strSpec = "+Headings"
            ElseIf strSpecLow = "default" Then
                strSpec = ""
            End If
            arrList(lngLevel, 11) = strSpec
        
        'Saves the number bold spec (spec 12) and number italic spec (spec 13).
        ElseIf strSpecLow = "bold bullet" _
            Or strSpecLow = "bold bullets" _
            Or strSpecLow = "bold number" _
            Or strSpecLow = "bold numbers" _
            Or strSpecLow = "bold letter" _
            Or strSpecLow = "bold letters" _
            Or strSpecLow = "bold numeral" _
            Or strSpecLow = "bold numerals" Then
            arrList(lngLevel, 12) = True
        ElseIf strSpecLow = "italic bullet" _
            Or strSpecLow = "italic bullets" _
            Or strSpecLow = "italic number" _
            Or strSpecLow = "italic numbers" _
            Or strSpecLow = "italic letter" _
            Or strSpecLow = "italic letters" _
            Or strSpecLow = "italic numeral" _
            Or strSpecLow = "italic numerals" Then
            arrList(lngLevel, 13) = True
        ElseIf strSpecLow = "bold italic number" _
            Or strSpecLow = "bold italic numbers" _
            Or strSpecLow = "italic bold number" _
            Or strSpecLow = "italic bold numbers" _
            Or strSpecLow = "bold and italic number" _
            Or strSpecLow = "bold and italic numbers" _
            Or strSpecLow = "italic and bold number" _
            Or strSpecLow = "italic and bold numbers" _
            Or strSpecLow = "bold italic letter" _
            Or strSpecLow = "bold italic letters" _
            Or strSpecLow = "italic bold letter" _
            Or strSpecLow = "italic bold letters" _
            Or strSpecLow = "bold and italic letter" _
            Or strSpecLow = "bold and italic letters" _
            Or strSpecLow = "italic and bold letter" _
            Or strSpecLow = "italic and bold letters" _
            Or strSpecLow = "bold italic numeral" _
            Or strSpecLow = "bold italic numerals" _
            Or strSpecLow = "italic bold numeral" _
            Or strSpecLow = "italic bold numerals" _
            Or strSpecLow = "bold and italic numeral" _
            Or strSpecLow = "bold and italic numerals" _
            Or strSpecLow = "italic and bold numeral" _
            Or strSpecLow = "italic and bold numerals" Then
            arrList(lngLevel, 12) = True
            arrList(lngLevel, 13) = True
        
        'Saves the bullet or number color (spec 14).
        ElseIf Right(strSpecLow, 12) = "bullet color" _
            Or Right(strSpecLow, 12) = "number color" _
            Or Right(strSpecLow, 12) = "letter color" _
            Or Right(strSpecLow, 13) = "numeral color" Then
            strSpec = Split(strSpec, " ")(0)
            strSpecLow = LCase(strSpec)
            If Left(strSpec, 1) = "#" Then
                strSpec = Right(strSpec, Len(strSpec) - 1)
                strSpec _
                    = Right(strSpec, 2) _
                    & Mid(strSpec, 3, 2) _
                    & Left(strSpec, 2)
                dblSpec = Val("&H" & strSpec)
                arrList(lngLevel, 24) = dblSpec
            ElseIf strSpecLow = "black" Then
                dblSpec = wdColorBlack
                arrList(lngLevel, 24) = dblSpec
            End If
        
        'Saves the bullet or number indent (spec 5) and text indent (spec 7).
        ElseIf Right(strSpecLow, 13) = "bullet indent" _
            Or Right(strSpecLow, 13) = "number indent" _
            Or Right(strSpecLow, 13) = "letter indent" _
            Or Right(strSpecLow, 14) = "numeral indent" Then
            arrList(lngLevel, 5) = InchesToPoints(dblSpec)
        ElseIf Right(strSpecLow, 11) = "text indent" Then
            arrList(lngLevel, 7) = InchesToPoints(dblSpec)
        
        'If bullets, saves bullets (spec 2) and style (spec 4).
        ElseIf (Right(strSpecLow, 6) = "bullet" _
            Or Right(strSpecLow, 7) = "bullets") _
            And Left(strSpecLow, 8) <> "based on" _
            And Left(strSpecLow, 11) <> "followed by" Then
            arrList(lngLevel, 2) = Left(strSpec, 1)
            arrList(lngLevel, 4) = wdListNumberStyleBullet
        
        'Saves the starting number (spec 10).
        ElseIf Right(strSpecLow, 8) = "start at" _
            Or Left(strSpecLow, 11) = "starting at" _
            Or Right(strSpecLow, 10) = "start with" _
            Or Left(strSpecLow, 13) = "starting with" _
            Or Right(strSpecLow, 8) = "begin at" _
            Or Left(strSpecLow, 12) = "beginning at" _
            Or Right(strSpecLow, 10) = "begin with" _
            Or Left(strSpecLow, 14) = "beginning with" Then
            arrList(lngLevel, 10) = Split(strSpec, " ")(2)
        
        'If numbers, saves the number specs.
        ElseIf (Right(strSpecLow, 7) = " number" _
            Or Right(strSpecLow, 7) = " letter" _
            Or Right(strSpecLow, 8) = " numeral" _
            Or Right(strSpecLow, 8) = " numbers" _
            Or Right(strSpecLow, 8) = " letters" _
            Or Right(strSpecLow, 9) = " numerals") _
            And Left(strSpecLow, 8) <> "based on" _
            And Left(strSpecLow, 11) <> "followed by" Then
            'Saves the number format (spec 2).
            strSpec = strSpecLow
            If Right(strSpec, 7) = " number" _
                Or Right(strSpec, 7) = " letter" Then
                strSpec = Left(strSpec, Len(strSpec) - 7)
            ElseIf Right(strSpec, 8) = " numeral" _
                Or Right(strSpec, 8) = " numbers" _
                Or Right(strSpec, 8) = " letters" Then
                strSpec = Left(strSpec, Len(strSpec) - 8)
            Else
                strSpec = Left(strSpec, Len(strSpec) - 9)
            End If
            If Right(strSpec, 6) = " roman" Then
                strSpec = Left(strSpec, Len(strSpec) - 6)
            End If
            If Right(strSpec, 7) = " arabic" Then
                strSpec = Left(strSpec, Len(strSpec) - 7)
            End If
            If Right(strSpec, 8) = " capital" Then
                strSpec = Left(strSpec, Len(strSpec) - 8)
            End If
            If Right(strSpec, 10) = " uppercase" _
                Or Right(strSpec, 10) = " lowercase" Then
                strSpec = Left(strSpec, Len(strSpec) - 10)
            End If
            If Right(strSpec, 11) = " upper-case" _
                Or Right(strSpec, 11) = " upper case" _
                Or Right(strSpec, 11) = " lower-case" _
                Or Right(strSpec, 11) = " lower case" Then
                strSpec = Left(strSpec, Len(strSpec) - 11)
            End If
            If Right(strSpec, 6) = " legal" Then
                strSpec = Left(strSpec, Len(strSpec) - 6)
            End If
            'Was strSpec = Split(strSpec, " ")(0)
                'Removes quotation marks.
                If Left(strSpec, 1) = Chr(34) Then
                    strSpec = Right(strSpec, Len(strSpec) - 1)
                ElseIf Left(strSpec, 1) = Chr(147) Then
                    strSpec = Right(strSpec, Len(strSpec) - 1)
                End If
                If Right(strSpec, 1) = Chr(34) Then
                    strSpec = Left(strSpec, Len(strSpec) - 1)
                ElseIf Right(strSpec, 1) = Chr(148) Then
                    strSpec = Left(strSpec, Len(strSpec) - 1)
                End If
            arrList(lngLevel, 2) = strSpec
            'Saves the number style (spec 4).
            dblSpec = wdListNumberStyleArabic
            If InStr(strSpecLow, "roman") <> 0 _
                Or InStr(strSpecLow, "numeral") <> 0 Then
                dblSpec = wdListNumberStyleLowercaseRoman
            End If
            If InStr(strSpecLow, "uppercase roman") <> 0 _
                Or InStr(strSpecLow, "uppercase numeral") <> 0 _
                Or InStr(strSpecLow, "upper-case roman") <> 0 _
                Or InStr(strSpecLow, "upper-case numeral") <> 0 _
                Or InStr(strSpecLow, "upper case roman") <> 0 _
                Or InStr(strSpecLow, "upper case numeral") <> 0 _
                Or InStr(strSpecLow, "capital roman") <> 0 _
                Or InStr(strSpecLow, "capital numeral") <> 0 Then
                dblSpec = wdListNumberStyleUppercaseRoman
            End If
            If InStr(strSpecLow, "letter") <> 0 Then
                dblSpec = wdListNumberStyleLowercaseLetter
            End If
            If InStr(strSpecLow, "uppercase letter") <> 0 _
                Or InStr(strSpecLow, "upper-case letter") <> 0 _
                Or InStr(strSpecLow, "upper case letter") <> 0 _
                Or InStr(strSpecLow, "capital letter") <> 0 Then
                dblSpec = wdListNumberStyleUppercaseLetter
            End If
            If InStr(strSpecLow, "legal") <> 0 Then
                dblSpec = wdListNumberStyleLegal
            End If
            arrList(lngLevel, 4) = dblSpec
            
        End If
    Next lngSpec
End Sub

Private Sub vbaInsertSampleText(arrStyles() As String)
    Dim lngStyle As Long
    For lngStyle = LBound(arrStyles) To UBound(arrStyles)
        With Selection
            .InsertAfter arrStyles(lngStyle) & " sample" & vbCrLf
            .Style = arrStyles(lngStyle)
            .Collapse wdCollapseEnd
        End With
     Next lngStyle
End Sub

'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
