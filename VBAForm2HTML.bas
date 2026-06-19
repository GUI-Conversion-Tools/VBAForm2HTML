Attribute VB_Name = "VBAForm2HTML"

' VBAForm2HTML v0.9.1
' https://github.com/GUI-Conversion-Tools/VBAForm2HTML
' Copyright (c) 2026 ZeeZeX
' This software is released under the MIT License.
' https://opensource.org/licenses/MIT

Option Explicit


#If VBA7 Then
    ' 64bit Office / VBA7 or later
    Private Declare PtrSafe Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
#Else
    ' 32bit Office
    Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
#End If


Public Sub TestRunConversion2Html()
    Call ConvertForm2HTML(UserForm1)
End Sub

Public Sub TestRunConversion2Html_2()
    Call ConvertForm2HTML(Array(UserForm1, UserForm2))
End Sub

Public Sub ConvertForm2HTML(ByVal frms As Variant, Optional ByVal usePrefix As Boolean = False, Optional htmlLang As String = "")
    
    ' frms: Variant
    '   Accepts a single UserForm object or an Array of UserForm objects to be converted.
    ' usePrefix: Boolean
    '   If set to True, the form name will be added to each element name.
    '   This is automatically set to True if frms is an array.
    
    Dim code As String
    Dim filePath As String
    Dim saveDir As String
    code = GenerateHTMLCode(frms, usePrefix, htmlLang)
    If code <> "" Then
        If ThisWorkbook.Path = "" Then
            saveDir = "C:"
        Else
            saveDir = ThisWorkbook.Path
        End If
        filePath = saveDir & "\output.html"
        Call SaveUTF8Text_NoBOM(filePath, code)
        MsgBox "Saved: " & filePath
    Else
        MsgBox "Conversion failed."
    End If
    
End Sub


Public Function GenerateHTMLCode(ByVal frms As Variant, Optional ByVal usePrefix As Boolean = False, Optional htmlLang As String = "") As String
    Dim root As Variant
    Dim indent As String
    Dim prefix As String
    Dim formName As String
    Dim controlVarName As String
    Dim parentVarName As String
    Dim unavailableNames() As Variant
    Dim ctrl As MSForms.Control
    Dim ctrls As Collection
    Dim item As Variant
    Dim r As String
    Const q As String = """"

    Dim pixelWidth As Long
    Dim pixelHeight As Long
    Dim pixelTop As Long
    Dim pixelLeft As Long
    Dim i As Long
    Dim cursorType As String
    Dim caption As String
    Dim colorSetting As String
    Dim cssSelectorProperties As Collection
    Dim tabPageProperties As Collection
    Dim activeTabProperties As Collection
    Dim inactiveTabProperties As Collection
    Dim toggleButtonInputProperties As Collection
    Dim toggleButtonSettingProperties As Collection
    Dim toggleButtonCheckedProperties As Collection
    Dim buttonTextLabelProperties As Collection
    Dim ctrlValue As String
    Dim spaceCnt As Long
    Dim temp As Variant
    Dim picSize As String
    Dim picPosition As String
    Dim elementStyles As New Collection
    Dim jsScripts As New Collection
    Dim htmlBodies As New Collection
    Dim elementStyle As String
    Dim jsScript As String
    Dim htmlBody As String
    Dim tabHeight As Long
    Dim tabPixelHeight As Long
    Dim positionProperties As Collection
    Dim colorProperties As Collection
    Dim fontProperties As Collection
    Dim otherProperties As Collection
    Dim htmlTitle As String
    Dim isHtmlTitleSet As Boolean: isHtmlTitleSet = False
    
    r = ""
    
    If IsArray(frms) Then
        usePrefix = True
    Else
        frms = VBA.Array(frms)
    End If
    
    Set cssSelectorProperties = New Collection

    With cssSelectorProperties
        .Add "background: #ffffff"
        .Add "display: flex"
        .Add "flex-direction: column"
        .Add "align-items: flex-start"
        .Add "min-width: max-content"
        .Add "gap: 0px"
    End With

    elementStyle = GenerateCssSelector("body", cssSelectorProperties, "") & vbLf
    elementStyles.Add elementStyle
    
    Set cssSelectorProperties = New Collection
    
    ' Always show up/down button of <input type="number">
    Set temp = New Collection
    With temp
        .Add "input[type=""number""]::-webkit-inner-spin-button,"
        .Add "input[type=""number""]::-webkit-outer-spin-button {"
        .Add "  opacity: 1;", "  display: block;"
        .Add "}"
    End With
    elementStyles.Add JoinCollection(temp, vbLf) & vbLf
    
    For Each root In frms
    
        Set cssSelectorProperties = New Collection
        elementStyle = ""
        jsScript = ""
        htmlBody = ""
        
        unavailableNames = VBA.Array("class", "script", "body", "id", "style")
        
        For i = LBound(unavailableNames) To UBound(unavailableNames)
            unavailableNames(i) = LCase(unavailableNames(i))
        Next
        
        If ContainsValue(unavailableNames, LCase(root.Name)) Then
            MsgBox GenerateUnavailableNameMessage(root)
            r = ""
            GenerateHTMLCode = r
            Exit Function
        End If
        
        If usePrefix Then
            prefix = root.Name & "-"
        Else
            prefix = ""
        End If
        
        pixelWidth = UserFormSizeToPixel(root.InsideWidth)
        pixelHeight = UserFormSizeToPixel(root.InsideHeight)
        
        formName = root.Name
        
        
        caption = root.caption
        caption = Convert2HTMLFormatText(caption)
        If Not isHtmlTitleSet Then
            htmlTitle = caption
            isHtmlTitleSet = True
        End If
        
        Set cssSelectorProperties = New Collection
        
        With cssSelectorProperties
            .Add "position: relative"
            .Add "margin-left: auto"
            .Add "margin-right: auto"
            .Add "box-sizing: border-box"
            .Add "overflow: hidden"
            .Add "width: " & pixelWidth & "px"
            .Add "height: " & pixelHeight & "px"
            .Add "background: " & LCase(FormColorToHex(root.BackColor))
            .Add "cursor: " & GetControlCursorType(root)
        End With
        
        Set temp = GetBorderSetting(root)
        Call ExtendCollection(cssSelectorProperties, temp)
        
        elementStyle = elementStyle & GenerateCssSelector(formName, cssSelectorProperties) & vbLf & vbLf
        
        Set cssSelectorProperties = New Collection
        

        elementStyle = elementStyle & vbLf

        Set ctrls = GetAllChildCtrlsDfs(root)
        For Each ctrl In ctrls
            controlVarName = GenerateCtrlVarName(ctrl, prefix)
            parentVarName = GenerateCtrlVarName(ctrl.Parent, prefix)
            
            Set cssSelectorProperties = New Collection
            Set tabPageProperties = New Collection
            Set activeTabProperties = New Collection
            Set inactiveTabProperties = New Collection
            Set toggleButtonInputProperties = New Collection
            Set toggleButtonSettingProperties = New Collection
            Set toggleButtonCheckedProperties = New Collection
            Set buttonTextLabelProperties = New Collection
            Set positionProperties = New Collection
            Set colorProperties = New Collection
            Set fontProperties = New Collection
            Set otherProperties = New Collection
            
            If ContainsValue(unavailableNames, LCase(ctrl.Name)) Then
                MsgBox GenerateUnavailableNameMessage(ctrl)
                r = ""
                GenerateHTMLCode = r
                Exit Function
            End If
            
            If IsSupportedCtrlType(ctrl) Then
                
                
                If TypeName(ctrl) <> "Page" Then
                    pixelLeft = UserFormSizeToPixel(ctrl.Left)
                    pixelTop = UserFormSizeToPixel(ctrl.Top)
                    pixelWidth = UserFormSizeToPixel(ctrl.Width)
                    pixelHeight = UserFormSizeToPixel(ctrl.Height)
                
                
                    With positionProperties
                        .Add "position: absolute"
                        .Add "box-sizing: border-box"
                        If TypeName(ctrl) = "Frame" Then
                            .Add "overflow: hidden"
                        End If
                        .Add "left: " & pixelLeft & "px"
                        .Add "top: " & pixelTop & "px"
                        .Add "width: " & pixelWidth & "px"
                        .Add "height: " & pixelHeight & "px"
                    End With
                End If
                
                
                If ContainsValue(Array("Label", "CommandButton", "Frame", "TextBox", "SpinButton", "ListBox", "CheckBox", "OptionButton", "ToggleButton", "ComboBox"), TypeName(ctrl)) Then
                    ' Set ForeColor
                    colorProperties.Add "color: " & LCase(FormColorToHex(ctrl.ForeColor))
                End If
                
                If ContainsValue(Array("Label", "CommandButton", "Frame", "TextBox", "SpinButton", "ListBox", "CheckBox", "OptionButton", "ToggleButton", "Image", "ComboBox"), TypeName(ctrl)) Then
                    ' Set BackColor
                    colorSetting = LCase(FormColorToHex(ctrl.BackColor))
                    If ContainsValue(Array("ComboBox", "Label", "TextBox", "CommandButton", "CheckBox", "OptionButton", "ToggleButton", "Image"), TypeName(ctrl)) Then
                        If ctrl.BackStyle = fmBackStyleTransparent Then
                            colorSetting = "transparent"
                        End If
                    End If

                    colorProperties.Add "background-color: " & colorSetting
                    
                End If
                
                If TypeName(ctrl) = "ToggleButton" Then
                    toggleButtonInputProperties.Add "display: none"
                    With buttonTextLabelProperties
                        .Add "border-top: 2px solid #ffffff"
                        .Add "border-left: 2px solid #ffffff"
                        .Add "border-right: 2px solid #7a7a7a"
                        .Add "border-bottom: 2px solid #7a7a7a"
                        .Add "box-sizing: border-box"
                        .Add "user-select: none"
                        .Add "-ms-user-select: none"
                    End With
                                        
                    With toggleButtonCheckedProperties
                        .Add "border-top: 2px solid #7a7a7a"
                        .Add "border-left: 2px solid #7a7a7a"
                        .Add "border-right: 2px solid #ffffff"
                        .Add "border-bottom: 2px solid #ffffff"
                    End With
                                        
                End If
                
                If TypeName(ctrl) = "TextBox" Then
                    If ctrl.MultiLine Then
                        With otherProperties
                            Select Case ctrl.ScrollBars
                                Case fmScrollBarsNone
                                    .Add "overflow-x: hidden"
                                    .Add "overflow-y: hidden"
                                Case fmScrollBarsHorizontal
                                    .Add "overflow-x: hidden"
                                    .Add "overflow-y: auto"
                                Case fmScrollBarsVertical
                                    .Add "overflow-x: auto"
                                    .Add "overflow-y: hidden"
                                Case fmScrollBarsBoth
                                    .Add "overflow-x: auto"
                                    .Add "overflow-y: auto"
                            End Select
                        End With
                    End If
                End If
                
                If TypeName(ctrl) = "MultiPage" Then
                    activeTabProperties.Add "padding: 2px 10px"
                    activeTabProperties.Add "background-color: " & LCase(FormColorToHex(&H8000000F))
                    inactiveTabProperties.Add "padding: 2px 8px"
                    inactiveTabProperties.Add "background-color:  " & LCase(AddRGB(FormColorToHex(&H8000000F), -20, -20, -20))
                    With tabPageProperties
                        .Add "box-sizing: border-box"
                        .Add "box-shadow: -1px -1px 0 #ffffff, 1px  1px 0 #666666, -2px -2px 0 #eeeeee, 2px  2px 0 #444444"
                        .Add "cursor: default"
                        If ctrl.Style = fmTabStyleNone Then
                            .Add "display: none;"
                        Else
                            .Add "display: inline-block;"
                        End If
                        .Add "vertical-align: top;"
                        .Add "color: " & LCase(FormColorToHex(ctrl.ForeColor))
                    End With
                    
                    
                End If
                
                
                If ContainsValue(Array("Label", "CommandButton", "Frame", "TextBox", "ListBox", "CheckBox", "OptionButton", "ComboBox", "ToggleButton", "MultiPage"), TypeName(ctrl)) Then
                    With fontProperties
                        .Add "font-family: " & ctrl.Font.Name
                        .Add "font-size: " & Round(ctrl.Font.Size * 1.33, 2) & "px"
                        
                        If ctrl.Font.Bold Then .Add "font-weight: bold"
                        If ctrl.Font.Italic Then .Add "font-style: italic"
                        If ctrl.Font.Underline And ctrl.Font.Strikethrough Then
                             .Add "text-decoration: underline line-through"
                        ElseIf ctrl.Font.Underline Then
                             .Add "text-decoration: underline"
                        ElseIf ctrl.Font.Strikethrough Then
                             .Add "text-decoration: line-through"
                        End If
                    End With
                    
                End If
                
                If TypeName(ctrl) = "Page" Then
                    
                    tabHeight = GetTextSizeFromCtrlFontSetting(ctrl.Parent, "TEST")(1)
                    tabPixelHeight = UserFormSizeToPixel(tabHeight)
                    pixelWidth = UserFormSizeToPixel(ctrl.Parent.Width)
                    pixelHeight = UserFormSizeToPixel(ctrl.Parent.Height)
                    
                    If ctrl.Parent.Style <> fmTabStyleNone Then
                        pixelHeight = pixelHeight - tabPixelHeight - 10
                    End If
                    
                    With positionProperties
                        .Add "position: relative"
                        .Add "box-sizing: border-box"
                        .Add "overflow: hidden"
                        .Add "width: " & pixelWidth & "px"
                        .Add "height: " & pixelHeight & "px"
                        
                    End With
                    
                    With otherProperties
                        .Add "padding: 2px 8px"
                        .Add "background-color: " & LCase(FormColorToHex(&H8000000F))
                        .Add "box-shadow: -1px -1px 0 #ffffff, 1px  1px 0 #666666, -2px -2px 0 #eeeeee, 2px  2px 0 #444444"
                        .Add "display: none"
                    End With
                End If
                
                If ContainsValue(Array("Frame", "TextBox", "ComboBox", "Label", "ListBox", "Image"), TypeName(ctrl)) Then
                    Set temp = GetBorderSetting(ctrl)
                    Call ExtendCollection(positionProperties, temp)
                End If

                If ContainsValue(Array("Label", "TextBox", "ComboBox", "CheckBox", "OptionButton", "ToggleButton", "ListBox"), TypeName(ctrl)) Then
                    Set temp = GetTextAlignSetting(ctrl)
                    Call ExtendCollection(positionProperties, temp)
                End If
                
                If TypeName(ctrl) = "ScrollBar" Then
                    If IsVerticalScrollBar(ctrl) Then
                        otherProperties.Add "writing-mode: vertical-lr"
                    End If
                End If
                
                ' Set mouse cursor
                If TypeName(ctrl) <> "MultiPage" And TypeName(ctrl) <> "Page" Then
                    cursorType = GetControlCursorType(ctrl)
                    otherProperties.Add "cursor: " & cursorType
                End If
                
                
                If TypeName(ctrl) = "Image" Then
                    Select Case ctrl.PictureSizeMode
                        Case fmPictureSizeModeClip
                            picSize = "auto"
                        Case fmPictureSizeModeStretch
                            picSize = "cover"
                        Case fmPictureSizeModeZoom
                            picSize = "contain"
                        Case Else
                            picSize = "auto"
                    End Select
                    
                    
                    Select Case ctrl.PictureAlignment
                        Case fmPictureAlignmentTopLeft
                            picPosition = "left top"
                        Case fmPictureAlignmentTopRight
                            picPosition = "right top"
                        Case fmPictureAlignmentCenter
                            picPosition = "center center"
                        Case fmPictureAlignmentBottomLeft
                            picPosition = "left bottom"
                        Case fmPictureAlignmentBottomRight
                            picPosition = "right bottom"
                        Case Else
                            picPosition = "left top"
                    End Select
                    
                    With otherProperties
                        .Add "background-image: url(" & q & "" & q & ")"
                        .Add "background-size: " & picSize
                        .Add "background-position: " & picPosition
                        .Add "background-repeat: no-repeat"
                    End With
                    
                End If
                
                If TypeName(ctrl) = "CheckBox" Or TypeName(ctrl) = "OptionButton" Then
                    Call ExtendCollection(buttonTextLabelProperties, positionProperties)
                    Call ExtendCollection(buttonTextLabelProperties, colorProperties)
                    Call ExtendCollection(buttonTextLabelProperties, fontProperties)
                    Call ExtendCollection(buttonTextLabelProperties, otherProperties)
                ElseIf TypeName(ctrl) = "ToggleButton" Then
                    Call ExtendCollection(buttonTextLabelProperties, positionProperties)
                    Call ExtendCollection(toggleButtonSettingProperties, colorProperties)
                    Call ExtendCollection(toggleButtonSettingProperties, fontProperties)
                    Call ExtendCollection(toggleButtonSettingProperties, otherProperties)
                ElseIf TypeName(ctrl) = "MultiPage" Then
                    Call ExtendCollection(tabPageProperties, fontProperties)
                    Call ExtendCollection(cssSelectorProperties, positionProperties)
                    Call ExtendCollection(cssSelectorProperties, otherProperties)
                Else
                    Call ExtendCollection(cssSelectorProperties, positionProperties)
                    Call ExtendCollection(cssSelectorProperties, colorProperties)
                    Call ExtendCollection(cssSelectorProperties, fontProperties)
                    Call ExtendCollection(cssSelectorProperties, otherProperties)
                End If
                
                elementStyle = elementStyle & GenerateCssSelector(controlVarName, cssSelectorProperties) & vbLf & vbLf
                
                If TypeName(ctrl) = "CheckBox" Or TypeName(ctrl) = "OptionButton" Then
                    elementStyle = elementStyle & GenerateCssSelector(controlVarName & "-label", buttonTextLabelProperties, ".") & vbLf & vbLf
                End If
                
                
                If TypeName(ctrl) = "ToggleButton" Then
                    elementStyle = elementStyle & GenerateCssSelector(controlVarName & "-label", buttonTextLabelProperties, ".") & vbLf & vbLf
                    elementStyle = elementStyle & GenerateCssSelector(controlVarName & "-input", toggleButtonInputProperties, ".") & vbLf & vbLf
                    elementStyle = elementStyle & GenerateCssSelector(controlVarName & "-setting", toggleButtonSettingProperties, ".") & vbLf & vbLf
                    elementStyle = elementStyle & GenerateCssSelector(controlVarName & "-input:checked + ." & controlVarName & "-label", toggleButtonCheckedProperties, ".") & vbLf & vbLf
                End If
                
                If TypeName(ctrl) = "MultiPage" Then
                    elementStyle = elementStyle & GenerateCssSelector(controlVarName & "-TabPages", tabPageProperties, ".") & vbLf & vbLf
                    elementStyle = elementStyle & GenerateCssSelector(controlVarName & "-ActiveTab", activeTabProperties, ".") & vbLf & vbLf
                    elementStyle = elementStyle & GenerateCssSelector(controlVarName & "-InactiveTab", inactiveTabProperties, ".") & vbLf & vbLf
                End If
                
                elementStyle = elementStyle & vbLf
                
            Else
                MsgBox GenerateUnsupportedControlMessage(ctrl)
                r = ""
                GenerateHTMLCode = r
                Exit Function
            End If
        Next ctrl
        
        
        htmlBody = htmlBody & "<div id=" & q & formName & q & ">"
        htmlBody = htmlBody & vbLf
        
        htmlBody = htmlBody & SetCssAllCtrlsDfs(root, prefix)

        htmlBody = htmlBody & vbLf
        htmlBody = htmlBody & "</div>" & vbLf
        
        For Each ctrl In ctrls
            If TypeName(ctrl) = "MultiPage" Then
                jsScript = jsScript & GenerateTabPageSwitchingFunc(ctrl, prefix)
                jsScript = jsScript & vbLf
            End If
        Next
        
        elementStyles.Add elementStyle
        htmlBodies.Add htmlBody
        jsScripts.Add jsScript
        
    Next root
    
    r = r & "<!DOCTYPE html>" & vbLf
    r = r & "<html lang=" & q & htmlLang & q & ">" & vbLf
    r = r & "<head>" & vbLf
    r = r & IndentSpaces(2) & "<meta charset=""UTF-8"">" & vbLf
    r = r & IndentSpaces(2) & "<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"" />" & vbLf
    r = r & IndentSpaces(2) & "<title>" & htmlTitle & "</title>" & vbLf
    r = r & vbLf
    r = r & IndentSpaces(2) & "<style>" & vbLf
    temp = JoinCollection(elementStyles, vbLf)
    temp = AdjustIndent(temp, 4)
    r = r & temp & vbLf
    r = r & IndentSpaces(2) & "</style>" & vbLf
    r = r & "</head>" & vbLf
    
    r = r & "<body>" & vbLf
    temp = JoinCollection(htmlBodies, vbLf) & vbLf
    temp = AdjustIndent(temp, 2)
    r = r & temp & vbLf
    
    r = r & IndentSpaces(2) & "<script>" & vbLf
    temp = JoinCollection(jsScripts, vbLf)
    temp = AdjustIndent(temp, 4)
    r = r & temp & vbLf
    r = r & IndentSpaces(2) & "</script>" & vbLf
    
    r = r & "</body>" & vbLf
    r = r & "</html>"
    
    GenerateHTMLCode = r
End Function

Private Function IsSupportedCtrlType(ByVal ctrl As Object) As Boolean
    ' Check if the control can be converted to CSS element
    Dim result As Boolean
    Select Case TypeName(ctrl)
        Case "Label", "CommandButton", "Frame", "TextBox", "SpinButton", "ListBox", "CheckBox", _
            "OptionButton", "Image", "ScrollBar", "ComboBox", "MultiPage", "Page", "ToggleButton"
            result = True
        Case Else
            result = False
    End Select
    IsSupportedCtrlType = result
End Function

Private Function SetCssElement(ByVal ctrl As Variant, ByVal prefix As String) As String
    Dim controlVarName As String
    Dim parentVarName As String
    Dim ctrlValue As String
    Dim spaceCnt As Long
    Dim r As String
    Const q As String = """"
    Dim ctrlDepth As Long
    Dim temp As Variant
    Dim listBoxSize As Long
    
    r = ""
    
    controlVarName = GenerateCtrlVarName(ctrl, prefix)
    parentVarName = GenerateCtrlVarName(ctrl.Parent, prefix)
    ctrlValue = ""

    If ContainsValue(Array("Label", "CommandButton", "CheckBox", "OptionButton", "ToggleButton"), TypeName(ctrl)) Then
        ctrlValue = Convert2HTMLFormatText(ctrl.caption)
    End If

    If ContainsValue(Array("TextBox", "ComboBox"), TypeName(ctrl)) Then
        ctrlValue = Convert2HTMLFormatText(ctrl.value, useBrTag:=False)
    End If
    
    ctrlDepth = GetFormControlDepth(ctrl)
    spaceCnt = ctrlDepth * 2

    Select Case TypeName(ctrl)
        Case "Label"
            r = r & IndentSpaces(spaceCnt) & "<div id=" & q & controlVarName & q & ">" & ctrlValue & "</div>"
        Case "TextBox"
            If ctrl.MultiLine Then
                r = r & IndentSpaces(spaceCnt) & "<textarea id=" & q & controlVarName & q
                
                If ctrl.WordWrap Then
                    r = r & " wrap=""soft"""
                Else
                    r = r & " wrap=""off"""
                End If
                
                If ctrl.Locked Then
                    r = r & " disabled"
                End If
                
                r = r & ">" & ctrlValue & "</textarea>"
            Else
                r = r & IndentSpaces(spaceCnt) & "<input id=" & q & controlVarName & q
                If ctrl.PasswordChar <> "" Then
                    r = r & " type=""password"" "
                Else
                    r = r & " type=""text"" "
                End If
                r = r & "value=" & q & ctrlValue & q
                If ctrl.Locked Then
                    r = r & " disabled"
                End If
                r = r & ">"
            End If
        Case "SpinButton"
            r = r & IndentSpaces(spaceCnt) & "<input id=" & q & controlVarName & q & " type=""number"">"
        Case "ComboBox"

            If ctrl.Style = fmStyleDropDownList Then
                r = r & IndentSpaces(spaceCnt) & "<select id=" & q & controlVarName & q & ">" & vbLf & GenerateCssSelectOptions(ctrl, spaceCnt) & IndentSpaces(spaceCnt) & "</select>"
            Else
                r = r & IndentSpaces(spaceCnt) & "<input id=" & q & controlVarName & q & " list=" & q & controlVarName & "-items-value" & q & " value=" & q & ctrlValue & q & ">" & vbLf
                r = r & GenerateCssDataList(ctrl, spaceCnt, prefix)
            End If

        Case "ListBox"
            r = r & IndentSpaces(spaceCnt) & "<select "
            If ctrl.MultiSelect = fmMultiSelectMulti Or ctrl.MultiSelect = fmMultiSelectExtended Then
                r = r & "multiple "
            End If
            ' If the size is 1 or less, <select> will appear as a combo box, so specify a value of at least 2.
            If ctrl.ListCount > 1 Then
                listBoxSize = ctrl.ListCount
            Else
                listBoxSize = 2
            End If
            r = r & "id=" & q & controlVarName & q & " size=" & listBoxSize & ">" & vbLf & GenerateCssSelectOptions(ctrl, spaceCnt) & IndentSpaces(spaceCnt) & "</select>"
            
        Case "CheckBox"
            r = r & IndentSpaces(spaceCnt) & "<label for=" & q & controlVarName & q & " class=" & q & controlVarName & "-label" & q & ">"
            temp = "<input " & "id=" & q & controlVarName & q & " type=""checkbox"""
            If ctrl.value Then
                temp = temp & " checked>"
            Else
                temp = temp & ">"
            End If
            If ctrl.Alignment = fmAlignmentLeft Then
                r = r & ctrlValue & temp & "</label>"
            Else
                r = r & temp & ctrlValue & "</label>"
            End If
            
        Case "OptionButton"
            r = r & IndentSpaces(spaceCnt) & "<label for=" & q & controlVarName & q & " class=" & q & controlVarName & "-label" & q & ">"
            temp = "<input " & "id=" & q & controlVarName & q & " type=""radio"" " & "name=" & q & parentVarName & "-radio" & q & " value=" & q & controlVarName & "-value" & q
            If ctrl.value Then
                temp = temp & " checked>"
            Else
                temp = temp & ">"
            End If
            If ctrl.Alignment = fmAlignmentLeft Then
                r = r & ctrlValue & temp & "</label>"
            Else
                r = r & temp & ctrlValue & "</label>"
            End If
        Case "ToggleButton"
            
            temp = IndentSpaces(spaceCnt) & "<input " & "id=" & q & controlVarName & q & " type=""checkbox"""
            temp = temp & " class=" & q & controlVarName & "-input " & controlVarName & "-setting" & q

            If ctrl.value Then
                temp = temp & " checked>"
            Else
                temp = temp & ">"
            End If
            
            temp = temp & "<label for=" & q & controlVarName & q & " class=" & q & controlVarName & "-label " & controlVarName & "-setting" & q & ">"
            temp = temp & ctrlValue & "</label>"
            r = r & temp
            
        Case "Image"
            r = r & IndentSpaces(spaceCnt) & "<div id=" & q & controlVarName & q & "></div>"
        Case "ScrollBar"
            r = r & IndentSpaces(spaceCnt) & "<input id=" & q & controlVarName & q & " type=""range"" min=" & q & ctrl.Min & q & " max=" & q & ctrl.Max & q & " step=""1"" value=" & q & ctrl.value & q & ">"
        Case "CommandButton"
            r = r & IndentSpaces(spaceCnt) & "<button id=" & q & controlVarName & q & ">" & ctrlValue & "</button>"
    End Select


    r = r & vbLf

    SetCssElement = r
End Function

Private Function SetCssAllCtrlsDfs(ByVal parentCtrl As Object, ByVal prefix As String) As String
    Dim children As Variant
    Dim code As String
    code = ""
    Dim results As Collection
    Set children = GetDirectChildCtrls(parentCtrl)
    Set results = SetCssAllCtrlsDfsRecursive(children, code, prefix)
    SetCssAllCtrlsDfs = code
End Function

Private Function SetCssAllCtrlsDfsRecursive(ByVal ctrls As Variant, ByRef code As String, ByVal prefix As String) As Collection
    Dim shouldProcess As Boolean: shouldProcess = False
    Dim results As New Collection
    Dim ctrl As Variant
    Dim ctrlDepth As Long
    Dim children As Variant
    Dim controlVarName As String
    Dim i As Long
    Const q As String = """"
    If ctrls.Count > 0 Then shouldProcess = True
    If shouldProcess Then
        For Each ctrl In ctrls
            results.Add ctrl
            controlVarName = GenerateCtrlVarName(ctrl, prefix)
            ctrlDepth = GetFormControlDepth(ctrl)
            code = code & SetCssElement(ctrl, prefix)
            Set children = GetDirectChildCtrls(ctrl)
            
            If TypeName(ctrl) = "Frame" Then
                code = code & IndentSpaces(ctrlDepth * 2) & "<fieldset id=" & q & controlVarName & q & ">"
                If ctrl.caption <> "" Then
                    code = code & "<legend>" & Convert2HTMLFormatText(ctrl.caption) & "</legend>"
                End If
                code = code & vbLf
            ElseIf TypeName(ctrl) = "MultiPage" Then
                code = code & IndentSpaces(ctrlDepth * 2) & "<div id=" & q & controlVarName & q & ">" & vbLf
                If ctrl.TabOrientation <> fmTabOrientationBottom Then
                    code = code & GenerateCssMultiPageTabsCode(ctrl, controlVarName, prefix, ctrlDepth) & vbLf
                End If
            ElseIf TypeName(ctrl) = "Page" Then
                code = code & IndentSpaces(ctrlDepth * 2) & "<div id=" & q & controlVarName & q
                If ctrl Is ctrl.Parent.Pages(0) Then
                    ' Setting for first tab page
                    code = code & " style=""display: block"""
                End If
                
                code = code & ">" & vbLf
            End If

            Call ExtendCollection(results, SetCssAllCtrlsDfsRecursive(children, code, prefix))
            
            If TypeName(ctrl) = "Frame" Then
                code = code & IndentSpaces(ctrlDepth * 2) & "</fieldset>" & vbLf
            ElseIf TypeName(ctrl) = "MultiPage" Then
                
                If ctrl.TabOrientation = fmTabOrientationBottom Then
                    code = code & GenerateCssMultiPageTabsCode(ctrl, controlVarName, prefix, ctrlDepth) & vbLf
                End If
                
                code = code & IndentSpaces(ctrlDepth * 2) & "</div>" & vbLf
            ElseIf TypeName(ctrl) = "Page" Then
                code = code & IndentSpaces(ctrlDepth * 2) & "</div>" & vbLf
            End If
        Next
    End If
    Set SetCssAllCtrlsDfsRecursive = results
End Function

Private Function GenerateCssMultiPageTabsCode(ByVal ctrl As Object, ByVal controlVarName As String, ByVal prefix As String, ByVal ctrlDepth As Long) As String
    ' Return code string such as:
    '  <div style="white-space: nowrap; font-size: 0;">
    '    <div id="MultiPage3_Page1-Tab" class="MultiPage3-TabPages MultiPage3-ActiveTab" onclick="MultiPage3_SwitchPage(0)">Page1</div>
    '    <div id="MultiPage3_Page2-Tab" class="MultiPage3-TabPages MultiPage3-InactiveTab" onclick="MultiPage3_SwitchPage(1)">Page2</div>
    '  </div>
    
    Dim code As String
    Dim multiPageFuncName As String
    Dim multiPageTabClassName As String
    Dim multiPageTabState As String
    Dim pageName As String
    Dim i As Long
    Dim ctrl2 As Variant
    Const q As String = """"
    
    code = IndentSpaces(ctrlDepth * 2 + 2) & "<div style=""white-space: nowrap; font-size: 0;"">" & vbLf
    multiPageFuncName = GenerateMultiPageFuncName(ctrl, prefix)
    multiPageTabClassName = controlVarName & "-TabPages"
    i = 0
    For Each ctrl2 In ctrl.Pages
        pageName = GenerateCtrlVarName(ctrl2, prefix)
        ' First Tab is active tab
        If ctrl2 Is ctrl.Pages(0) Then
            multiPageTabState = controlVarName & "-ActiveTab"
        Else
            multiPageTabState = controlVarName & "-InactiveTab"
        End If
        
        code = code & IndentSpaces(ctrlDepth * 2 + 4) & "<div id=" & q & pageName & "-Tab" & q & " class=" & q & multiPageTabClassName & " " & multiPageTabState & q & " onclick=" & q & multiPageFuncName & "(" & i & ")" & q & ">" & Convert2HTMLFormatText(ctrl2.caption) & "</div>" & vbLf
        i = i + 1
    Next
    code = code & IndentSpaces(ctrlDepth * 2 + 2) & "</div>"
    GenerateCssMultiPageTabsCode = code
End Function


Private Function GetAllChildCtrlsDfs(ByVal parentCtrl As Object) As Collection
    Dim children As Variant
    Dim results As Collection
    Set children = GetDirectChildCtrls(parentCtrl)
    Set results = GetAllChildCtrlsDfsRecursive(children)
    Set GetAllChildCtrlsDfs = results
End Function

Private Function GetAllChildCtrlsDfsRecursive(ByVal ctrls As Variant) As Collection
    Dim shouldProcess As Boolean: shouldProcess = False
    Dim results As New Collection
    Dim ctrl As Variant
    Dim children As Variant
    If ctrls.Count > 0 Then shouldProcess = True
    If shouldProcess Then
        For Each ctrl In ctrls
            results.Add ctrl
            Set children = GetDirectChildCtrls(ctrl)
            Call ExtendCollection(results, GetAllChildCtrlsDfsRecursive(children))
        Next
    End If
    Set GetAllChildCtrlsDfsRecursive = results
End Function

Private Function GetDirectChildCtrls(ByVal parentCtrl As Object) As Collection
    Dim results As Collection
    Set results = New Collection
    Dim ctrl As Object
    Dim root As Object
    
    Set root = GetUserFormObjectFromCtrl(parentCtrl)
    
    If TypeName(parentCtrl) = "MultiPage" Then
        For Each ctrl In parentCtrl.Pages
            results.Add ctrl
        Next
    Else
        For Each ctrl In root.Controls
            If parentCtrl Is ctrl.Parent Then
                results.Add ctrl
            End If
        Next ctrl
    End If
    Set GetDirectChildCtrls = results
End Function

Private Function GetUserFormObjectFromCtrl(ByVal ctrl As Object) As Object
    ' Get the ancestor (UserForm) of the control.
    
    Dim root As Object
    
    If ctrl Is Nothing Then
        Err.Raise 13
    End If
    
    Set root = ctrl
    ' Loop to get root(UserForm) object
    On Error GoTo Finally:
    Do While True
        Set root = root.Parent
    Loop
    On Error GoTo 0

Finally:
    Set GetUserFormObjectFromCtrl = root
End Function

Private Function GenerateCtrlVarName(ByVal ctrl As Object, ByVal prefix As String) As String
    ' Generates a valid, unique identifier for a control in the target language.
    Dim controlVarName As String

    If TypeName(ctrl) = "Page" Then
    ' VBA allows duplicate names for Page objects if they belong to different MultiPage controls.
    ' To ensure unique variable names in the target language (which typically uses a flat
    ' namespace), namespace the Page by prepending its parent MultiPage's name.
    ' Example: "Page1" inside "MultiPage1" becomes "MultiPage1_Page1"
        controlVarName = prefix & ctrl.Parent.Name & "_" & ctrl.Name
    Else
        controlVarName = prefix & ctrl.Name
    End If

    GenerateCtrlVarName = controlVarName
End Function


Private Function IndentSpaces(ByVal n As Long) As String
    Dim i As Long
    Dim result As String
    result = ""
    For i = 1 To n
        result = result + " "
    Next
    IndentSpaces = result
End Function


Private Function GenerateCssSelector(ByVal ctrlName As String, ByVal properties As Variant, Optional ByVal selectorSymbol As String = "#") As String
    ' Generate CSS Selector from the control name and the list of properties (Array/Collection)
    
    ' Example:
    ' ctrlName:="Label1", properties:=Array("Left: 299px", "Top: 129px", "width: 97px", "height: 16px", "color: #000000")
    ' ->
    ' #Label1 {
    '   Left: 299px;
    '   Top: 129px;
    '   width: 97px;
    '   height: 16px;
    ' color: #000000;
    ' }
    
    Dim item As Variant
    Dim result As String
    result = selectorSymbol & ctrlName & " {"
    For Each item In properties
        result = result & vbLf & IndentSpaces(2) & item & ";"
    Next
    result = result & vbLf & "}"
    GenerateCssSelector = result
End Function

Private Function GenerateMultiPageFuncName(ByVal multiPageCtrl As Object, ByVal prefix As String) As String
    ' Generate JavaScript Function Name such as "MultiPage1_SwitchPage" / "UserForm1_MultiPage1_SwitchPage"
    Dim root As Object
    Dim result As String
    Set root = GetUserFormObjectFromCtrl(multiPageCtrl)
    result = prefix & multiPageCtrl.Name & "_SwitchPage"
    result = VBA.Replace(result, "-", "_")
    GenerateMultiPageFuncName = result
End Function

Private Function GenerateTabPageSwitchingFunc(ByVal multiPageCtrl As Object, ByVal prefix As String) As String
    ' Define JavaScript Function
    Const q As String = """"
    
    Dim ctrl As Variant
    Dim collPages As New Collection
    Dim collTabs As New Collection
    Dim collFuncStr As New Collection
    Dim jsArrPages As String
    Dim jsArrTabs As String
    Dim jsFuncName As String
    Dim jsFuncStr As String
    Dim controlVarName As String
    Dim activeTabColor As String
    Dim inactiveTabColor As String
    
    activeTabColor = FormColorToHex(&H8000000F)
    inactiveTabColor = AddRGB(activeTabColor, -20, -20, -20)
    activeTabColor = LCase(activeTabColor)
    inactiveTabColor = LCase(inactiveTabColor)
    
    For Each ctrl In multiPageCtrl.Pages
        controlVarName = GenerateCtrlVarName(ctrl, prefix)
        collPages.Add q & controlVarName & q
        collTabs.Add q & controlVarName & "-Tab" & q
    Next
    jsArrPages = "[" & JoinCollection(collPages, ", ") & "]"
    jsArrTabs = "[" & JoinCollection(collTabs, ", ") & "]"
    jsFuncName = GenerateMultiPageFuncName(multiPageCtrl, prefix)
    With collFuncStr
        .Add "function " & jsFuncName & "(index) {"
        .Add "  var pages = " & jsArrPages & ";"
        .Add "  var tabs = " & jsArrTabs & ";"
        .Add "  for (var i = 0; i < pages.length; i++) {"
        .Add "    var page = document.getElementById(pages[i]);"
        .Add "    var tab = document.getElementById(tabs[i]);"
        .Add "      if (i === index) {"
        .Add "        page.style.display = ""block"";"
        .Add "        tab.style.padding = ""2px 10px"";"
        .Add "        tab.style.backgroundColor = " & q & activeTabColor & q & ";"
        .Add "      } else {"
        .Add "        page.style.display = ""none"";"
        .Add "        tab.style.padding = ""2px 8px"";"
        .Add "        tab.style.backgroundColor = " & q & inactiveTabColor & q & ";"
        .Add "      }"
        .Add "  }"
        .Add "}"
    End With
    jsFuncStr = JoinCollection(collFuncStr, vbLf)
    GenerateTabPageSwitchingFunc = jsFuncStr
End Function


Private Function GetBorderSetting(ByVal ctrl As Object) As Collection
    Dim hexBorderColor As String
    Dim result As New Collection
    hexBorderColor = FormColorToHex(ctrl.BorderColor)
    hexBorderColor = LCase(hexBorderColor)
    
    Select Case ctrl.BorderStyle
        Case fmBorderStyleSingle
            ' SpecialEffect is 0 if BorderStyle is 1
            result.Add "border: 1px solid " & hexBorderColor
        Case fmBorderStyleNone
            Select Case ctrl.SpecialEffect
                Case fmSpecialEffectFlat
                    result.Add "border: 0px solid " & hexBorderColor
                Case fmSpecialEffectRaised
                    result.Add "box-shadow: -1px -1px 0 #ffffff, 1px  1px 0 #666666, -2px -2px 0 #eeeeee, 2px  2px 0 #444444"
                Case fmSpecialEffectSunken
                    result.Add "box-shadow: 1px  1px 0 #ffffff, -1px -1px 0 #666666, 2px  2px 0 #eeeeee, -2px -2px 0 #444444"
                Case fmSpecialEffectEtched
                    result.Add "border: 1px solid #666666"
                    result.Add "outline: 1px solid #ffffff"
                    result.Add "outline-offset: -2px"
                Case fmSpecialEffectBump
                    result.Add "border: 1px solid #999999"
                    result.Add "box-shadow: -1px -1px 0 #ffffff, 1px  1px 0 #666666, -2px -2px 0 #eeeeee, 2px  2px 0 #444444"
            End Select
    End Select

    Set GetBorderSetting = result
End Function

Private Function GetTextAlignSetting(ByVal ctrl As Object) As Collection
    Dim result As New Collection
   
    If ContainsValue(Array("CheckBox", "OptionButton", "ToggleButton"), TypeName(ctrl)) Then
        Select Case ctrl.TextAlign
            Case fmTextAlignLeft
                result.Add "justify-content: flex-start"
            Case fmTextAlignCenter
                result.Add "justify-content: center"
            Case fmTextAlignRight
                result.Add "justify-content: flex-end"
            Case Else
                result.Add "justify-content: center"
        End Select
    Else
        Select Case ctrl.TextAlign
            Case fmTextAlignLeft
                result.Add "text-align: left"
            Case fmTextAlignCenter
                result.Add "text-align: center"
            Case fmTextAlignRight
                result.Add "text-align: right"
            Case Else
                result.Add "text-align: center"
        End Select
    End If
    If ContainsValue(Array("CheckBox", "OptionButton", "ToggleButton"), TypeName(ctrl)) Then
        result.Add "align-items: center"
        result.Add "display: flex"
    End If
    
    Set GetTextAlignSetting = result
End Function

Private Function IsVerticalScrollBar(ByVal ctrl As Object) As Boolean
    ' Vertical -> True, Horizontal -> False
    Dim result As Boolean
    Select Case ctrl.orientation
        Case fmOrientationAuto
            If ctrl.Width > ctrl.Height Then
                result = False
            Else
                result = True
            End If
            
        Case fmOrientationVertical
            result = True
        Case fmOrientationHorizontal
            result = False
        Case Else
            result = True
    End Select
    IsVerticalScrollBar = result
End Function


Private Function GetControlCursorType(ByVal ctrl As Object) As String
    Dim cursorType As String
    
    Select Case ctrl.MousePointer
        Case fmMousePointerDefault
            cursorType = "auto"
            
        Case fmMousePointerArrow
            cursorType = "default"
            
        Case fmMousePointerCross
            cursorType = "crosshair"
            
        Case fmMousePointerIBeam
            cursorType = "text"
            
        Case fmMousePointerSizeNESW
            cursorType = "nesw-resize"
            
        Case fmMousePointerSizeNS
            cursorType = "ns-resize"
            
        Case fmMousePointerSizeNWSE
            cursorType = "nwse-resize"
            
        Case fmMousePointerSizeWE
            cursorType = "ew-resize"
            
        Case fmMousePointerUpArrow
            cursorType = "n-resize" ' closest match
            
        Case fmMousePointerHourGlass
            cursorType = "wait"
            
        Case fmMousePointerNoDrop
            cursorType = "not-allowed"
            
        Case fmMousePointerAppStarting
            cursorType = "progress" ' better CSS equivalent
            
        Case fmMousePointerHelp
            cursorType = "help"
            
        Case fmMousePointerSizeAll
            cursorType = "move"
            
        Case Else
            cursorType = "auto"
    End Select
    
    GetControlCursorType = cursorType
End Function


Private Function GenerateCssDataList(ByVal ctrl As Object, ByVal indentOffset As Long, ByVal prefix As String) As String
    ' Generates HTML <option> elements from a ListBox or ComboBox control
    ' Each item in the control is converted into:
    
    ' <datalist id="fruits">
    '   <option value="1">
    '   <option value="2">
    '   <option value="3">
    ' </datalist>
    
    Const q As String = """"
    Dim item As Variant
    Dim i As Long: i = 0
    Dim r As String
    Dim controlVarName As String
    Dim itemValue As String
    controlVarName = GenerateCtrlVarName(ctrl, prefix)
    r = IndentSpaces(0 + indentOffset) & "<datalist id=" & q & controlVarName & "-items-value" & q & ">" & vbLf
    If ctrl.ListCount > 0 Then
        For Each item In ctrl.List
            i = i + 1
            itemValue = Convert2HTMLFormatText(item)
            r = r & IndentSpaces(2 + indentOffset) & "<option value=" & q & itemValue & q & ">" & vbLf
            If i = ctrl.ListCount Then Exit For
        Next item
    Else
        ' If no items exist, output an empty <option> element
        r = r & IndentSpaces(2 + indentOffset) & "<option value=" & q & " " & q & ">" & vbLf
    End If
    
    r = r & IndentSpaces(0 + indentOffset) & "</datalist>"
    
    
    GenerateCssDataList = r
End Function

Private Function GenerateCssSelectOptions(ByVal ctrl As Object, ByVal indentOffset As Long) As String
    ' Generates HTML <option> elements from a ListBox or ComboBox control
    ' Each item in the control is converted into:
    '
    '   <option value="1">1</option>
    '   <option value="2">2</option>
    '   <option value="3">3</option>
    '
    
    Const q As String = """"
    Dim item As Variant
    Dim i As Long: i = 0
    Dim r As String
    Dim itemValue As String
    r = ""
    If ctrl.ListCount > 0 Then
        For Each item In ctrl.List
            i = i + 1
            itemValue = Convert2HTMLFormatText(item)
            r = r & IndentSpaces(2 + indentOffset) & "<option value=" & q & itemValue & q & ">" & itemValue & "</option>" & vbLf
            If i = ctrl.ListCount Then Exit For
        Next item
    Else
        ' If no items exist, output an empty <option> element
        r = r & IndentSpaces(2 + indentOffset) & "<option value=" & q & q & ">" & item & "</option>" & vbLf
    End If
    
    GenerateCssSelectOptions = r
End Function

Private Function GetTextSizeFromCtrlFontSetting(ByVal ctrl As Object, ByVal targetText As String) As Variant()
    '------------------------------------------------------------------------------
    ' Returns the rendered text size (Width, Height) for a given text string
    ' using the same font settings as the specified control.
    ' Size is measured in points, not pixels.
    '
    ' Parameters:
    '   ctrl        - The reference control whose font settings will be used.
    '   targetText  - The text to measure. If empty, "i" is used to ensure a measurable size is returned.
    '                 (The letter  gi h is one of the ASCII characters with the narrowest rendering width.)
    '
    ' Returns:
    '   Variant() Array
    '       (0) = Text width
    '       (1) = Text height
    '
    ' Notes:
    '   - A temporary hidden Label control is dynamically created on the parent
    '     UserForm to calculate the actual rendered text dimensions.
    '   - AutoSize is enabled so the Label automatically resizes to fit the text.
    '   - The temporary control is removed immediately after measurement.
    '
    ' Compatibility Note:
    '   In Excel 2013 and earlier, it was confirmed that enabling .AutoSize does not
    '   correctly update the .Width and .Height properties, which remain 0.
    '   To avoid returning invalid measurements, this function falls back to
    '   an estimated size calculation based on the font size and text length.
    '   This fallback is less accurate than actual rendered text measurement.
    '
    '------------------------------------------------------------------------------
    Dim rootForm As Object
    Dim tempLabel As Object
    Dim tempName As String
    Dim textWidthSize As Double
    Dim textHeightSize As Double
    ' Prevent zero-size measurement for empty strings.
    If targetText = "" Then targetText = "i"
    ' Generate a unique temporary control name.
    tempName = "TempLabel_" & VBA.Replace(GenerateUUIDv4(), "-", "_")
    ' Get the parent UserForm from the specified control.
    Set rootForm = GetUserFormObjectFromCtrl(ctrl)
    ' Create a temporary Label control for text measurement.
    Set tempLabel = rootForm.Controls.Add("Forms.Label.1", tempName, True)
    ' Initialize control properties.
    tempLabel.Height = 0
    tempLabel.Width = 0
    tempLabel.caption = ""
    tempLabel.AutoSize = True
    tempLabel.WordWrap = False
    ' Optional debug background color.
    tempLabel.BackColor = &H80C0FF
    ' Copy font settings from the source control.
    tempLabel.Font.Name = ctrl.Font.Name
    tempLabel.Font.Size = ctrl.Font.Size
    tempLabel.Font.Bold = ctrl.Font.Bold
    tempLabel.Font.Italic = ctrl.Font.Italic
    tempLabel.Font.Underline = ctrl.Font.Underline
    tempLabel.Font.Strikethrough = ctrl.Font.Strikethrough
    ' Apply target text so AutoSize calculates the rendered dimensions.
    tempLabel.caption = targetText
    ' Read calculated size.
    textWidthSize = tempLabel.Width
    textHeightSize = tempLabel.Height
    
    ' In Excel 2013 and earlier, it was confirmed that the result of .AutoSize
    ' is not reflected in .Width/.Height and remains 0.
    ' As a fallback, the font size is used instead,
    ' although the measurement accuracy is reduced.
    If textWidthSize = 0 Then textWidthSize = ctrl.Font.Size * Len(targetText)
    If textHeightSize = 0 Then textHeightSize = ctrl.Font.Size
    
    ' The .Controls.Remove method does not accept a String argument; the argument must be of type Variant (String).
    ' example: tempLabel.Name or CVar(tempName)
    Call rootForm.Controls.Remove(tempLabel.Name)
    ' Release object reference.
    Set tempLabel = Nothing
    ' Return width and height as an array.
    GetTextSizeFromCtrlFontSetting = VBA.Array(textWidthSize, textHeightSize)
End Function

Private Function Convert2HTMLFormatText(ByVal text As String, Optional useBrTag As Boolean = True) As String
    ' Escape special characters in the string
    ' "&" should be replaced first
    text = VBA.Replace(text, "&", "&amp;")
    text = VBA.Replace(text, "<", "&lt;")
    text = VBA.Replace(text, ">", "&gt;")
    text = VBA.Replace(text, """", "&quot;")
    text = VBA.Replace(text, "'", "&#39;") ' Use numeric character reference for compatibility ("&apos;" is not officially supported in HTML 4.)
    text = VBA.Replace(text, " ", "&nbsp;")
    ' Convert VBA line breaks to HTML format
    ' vbCrLf should be replaced first
    text = VBA.Replace(text, vbCrLf, vbLf)
    text = VBA.Replace(text, vbCr, vbLf)
    If useBrTag Then
        text = VBA.Replace(text, vbLf, "<br>") ' For div
    Else
        text = VBA.Replace(text, vbLf, "&#13;") ' For textarea
    End If
    Convert2HTMLFormatText = text
End Function

Private Function AddRGB(ByVal hexColor As String, _
                       ByVal addR As Long, _
                       ByVal addG As Long, _
                       ByVal addB As Long) As String
    ' Example:
    ' AddRGB("#F0F0F0", -20, -20, -20) -> "#DCDCDC"
    ' AddRGB("#000000", -20, -20, -20) -> "#000000"
    ' AddRGB("#000000", 13, 14, 15) -> "#0D0E0F"
    ' AddRGB("#FEFE00", 15, 15, 15) -> "#FFFF0F"
    Dim r As Long, g As Long, b As Long

    ' Remove "#" if included
    hexColor = Replace(hexColor, "#", "")

    ' Validate length
    If Len(hexColor) <> 6 Then
        AddRGB = "#000000"
        Exit Function
    End If

    ' Convert HEX -> RGB
    r = CLng("&H" & Mid(hexColor, 1, 2))
    g = CLng("&H" & Mid(hexColor, 3, 2))
    b = CLng("&H" & Mid(hexColor, 5, 2))

    ' Add values
    r = r + addR
    g = g + addG
    b = b + addB

    ' Clamp between 0 and 255
    If r < 0 Then r = 0
    If r > 255 Then r = 255

    If g < 0 Then g = 0
    If g > 255 Then g = 255

    If b < 0 Then b = 0
    If b > 255 Then b = 255

    ' Convert back to HEX
    AddRGB = "#" & _
             Right("0" & Hex(r), 2) & _
             Right("0" & Hex(g), 2) & _
             Right("0" & Hex(b), 2)

End Function

Private Function FormColorToHex(ByVal clr As Long) As String
    ' Example:
    ' 16777215 -> "#FFFFFF"
    ' 0 -> "#000000"
    ' &H000000FF& (255) -> "#FF0000"
    ' &H00B4769E& (11826846) -> "#9E76B4"
    ' &H8000000F& (-2147483633) -> "#F0F0F0"(Windows XP[Luna Theme]/10/11), "#D4D0C8"(Windows 2000/XP[Classic Theme])
    Dim r As Long, g As Long, b As Long
    ' Convert a system color to its decimal color code when the parameter is a system color
    If 0 > clr Or clr >= 2147483648# Then
        clr = GetSysColor(clr And &HFF)
    End If
    ' Retrieve each component of the RGB color.
    r = clr And &HFF            ' Extract low-order 8 bits
    g = (clr \ &H100) And &HFF  ' Extract bits 8-15
    b = (clr \ &H10000) And &HFF ' Extract bits 16-23
    
    ' Convert the decimal RGB values to a #RRGGBB hex string and return it
    FormColorToHex = "#" & _
                     Right("0" & Hex(r), 2) & _
                     Right("0" & Hex(g), 2) & _
                     Right("0" & Hex(b), 2)
End Function


Private Function ContainsValue(ByVal itemList As Variant, ByVal value As Variant) As Boolean
    ' Check if a specific value exists in Array/Collection/Dictionary
    ' itemList - Array/Collection/Dictionary to search
    ' value - value to check
    ' Performs strict type comparison for non-numeric values
    ' Nested arrays are not supported. Objects are compared by reference
    ' Dependency: IsStrictlyEqual(helper function)
    Dim item As Variant
    Dim temp As Variant
    If LCase(TypeName(itemList)) = "dictionary" Then
        itemList = itemList.items
    End If
    If IsArray(itemList) Then
        On Error GoTo Finally
        ' Uninitialized Array -> False
        temp = LBound(itemList)
        On Error GoTo 0
    End If
    For Each item In itemList
    
        If IsStrictlyEqual(item, value) Then
            ContainsValue = True
            Exit Function
        End If
    Next
Finally:
    ContainsValue = False
    
End Function

Private Function IsStrictlyEqual(ByVal value1 As Variant, ByVal value2 As Variant) As Boolean
    ' Performs a strict equality comparison including data types.
    ' Numeric types (Integer, Long, Double, etc.) are treated as compatible.
    ' Boolean and Date types are NOT treated as numeric.
    Dim t1 As VbVarType, t2 As VbVarType
    t1 = VarType(value1)
    t2 = VarType(value2)
    
    ' Returns True if objects point to the same reference.
    ' Objects are evaluated first to prevent false matches (e.g., Empty vs empty Cells).
    ' (Also applies to variables holding both objects and other data types)
    If IsObject(value1) Or IsObject(value2) Then
        If IsObject(value1) And IsObject(value2) Then
            IsStrictlyEqual = (value1 Is value2)
        End If
        Exit Function
    End If
    
    ' Null / Empty
    If IsNull(value1) Or IsNull(value2) Then
        IsStrictlyEqual = (IsNull(value1) And IsNull(value2))
        Exit Function
    ElseIf IsEmpty(value1) Or IsEmpty(value2) Then
        IsStrictlyEqual = (IsEmpty(value1) And IsEmpty(value2))
        Exit Function
    End If
    
    
    ' Arrays are not supported (Extend if necessary).
    If IsArray(value1) Or IsArray(value2) Then
        IsStrictlyEqual = False
        Exit Function
    End If
    
    ' Error values
    If t1 = vbError Or t2 = vbError Then
        IsStrictlyEqual = (t1 = t2 And value1 = value2)
        Exit Function
    End If
    
    ' String, Date, Boolean
    If (t1 = vbString Or t2 = vbString) Or (t1 = vbDate Or t2 = vbDate) Or (t1 = vbBoolean Or t2 = vbBoolean) Then
        IsStrictlyEqual = (t1 = t2 And value1 = value2)
        Exit Function
    End If
    
    ' Other data types (e.g., Numeric)
    On Error Resume Next
    IsStrictlyEqual = (value1 = value2)
    Exit Function
    On Error GoTo 0
    IsStrictlyEqual = False
End Function


Private Function UserFormSizeToPixel(ByVal ufSize As Double) As Long
    ' Function to convert the size of a UserForm or control to pixels
    ' Excel VBA UserForm dimensions are internally handled as
    ' DPI-independent logical points based on a fixed 96 DPI system.
    ' Therefore, point-to-pixel conversion can be calculated as:
    '     pixel = point * (96 / 72)
    ' and works consistently regardless of the monitor DPI setting.
    Dim pixelSize As Long
    pixelSize = Round(ufSize * (96 / 72))
    UserFormSizeToPixel = pixelSize
End Function

Private Function GenerateUUIDv4() As String
    Dim i As Long
    Dim b(15) As Byte
    Dim s As String
    Dim hexStr As String
    
    ' Initialize random number generator
    Randomize
    
    ' Generate 16 bytes of random values
    For i = 0 To 15
        b(i) = Int(Rnd() * 256)
    Next i
    
    ' Set version (4) (set bits 7-4 to 0100)
    b(6) = (b(6) And &HF) Or &H40
    
    ' Set variant (10xx)
    b(8) = (b(8) And &H3F) Or &H80
    
    ' Convert the 16 bytes to a string (with hyphen format)
    hexStr = ""
    For i = 0 To 15
        hexStr = hexStr & Right$("0" & Hex(b(i)), 2)
        Select Case i
            Case 3, 5, 7, 9
                hexStr = hexStr & "-"
        End Select
    Next i
    
    GenerateUUIDv4 = LCase$(hexStr)
End Function

Private Sub SaveUTF8Text_NoBOM(ByVal filePath As String, ByVal textData As String)
    ' Save the specified string as UTF-8 without BOM
    Dim stream As Object
    Dim bytes() As Byte
    
    ' Normalize line endings
    textData = VBA.Replace(textData, vbCrLf, vbLf)
    textData = VBA.Replace(textData, vbCr, vbLf)
    textData = VBA.Replace(textData, vbLf, vbNewLine)
    
    ' Convert to UTF-8 and remove BOM
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' Text mode
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText textData
    stream.position = 0
    stream.Type = 1 ' Switch to binary mode
    bytes = stream.Read
    stream.Close
    Set stream = Nothing
    
    ' Remove BOM if present
    If UBound(bytes) >= 2 Then
        If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then
            bytes = MidB(bytes, 4) ' Remove BOM (EF BB BF)
        End If
    End If
    
    ' Save file in binary mode
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write bytes
    stream.SaveToFile filePath, 2
    stream.Close
    Set stream = Nothing
End Sub


Private Function GenerateUnsupportedControlMessage(ByVal ctrl As Object) As String
    Const q As String = """"
    GenerateUnsupportedControlMessage = "Control type " & q & TypeName(ctrl) & q & " is not supported."
End Function

Private Function GenerateUnavailableNameMessage(ByVal ctrl As Object) As String
    Const q As String = """"
    GenerateUnavailableNameMessage = "Object Name " & q & ctrl.Name & q & " is not available." & vbLf & "Please use a different name instead."
End Function

Private Function GetFormControlDepth(ByVal ctrl As Object) As Long
    ' Get the hierarchy depth of the control
    Dim depth As Long
    Dim temp As Variant
    depth = 0
    Set temp = ctrl
    Do While True
        If depth Mod 10 = 0 Then DoEvents
        On Error GoTo Finally
        Set temp = temp.Parent
        depth = depth + 1
        On Error GoTo 0
    Loop
Finally:
    
    If Err.Number <> 438 Then
        Err.Raise Number:=Err.Number
    End If
    
    GetFormControlDepth = depth
    
End Function

Private Function SortFormControlsByDepth(ByVal frmControls As Variant) As Collection
    ' Sort the list of UserForm controls in ascending order of hierarchy depth
    Dim tempColl As Collection
    Set tempColl = New Collection
    Dim sortedColl As Collection
    Set sortedColl = New Collection
    Dim ctrl As Variant
    Dim tempArray() As Variant
    Dim depth As Long
    Dim item As Variant
    For Each ctrl In frmControls
        depth = GetFormControlDepth(ctrl)
        tempColl.Add VBA.Array(depth, ctrl)
    Next ctrl
    If tempColl.Count > 0 Then
        tempArray = Collection2Array(tempColl)
        Call InsertionSortJaggedArray(tempArray, reverse:=False)
        For Each item In tempArray
            sortedColl.Add item(1)
        Next item
    End If
    Set SortFormControlsByDepth = sortedColl
End Function


Private Function Collection2Array(ByVal coll As Collection, Optional ByVal isStartIdx1 As Boolean = False) As Variant()
    ' Convert a Collection to an array
    ' If isStartIdx1 is True, create an array starting from index 1 (to match Collection numbering)
    Dim arr() As Variant
    Dim item As Variant
    Dim idx As Long
    If coll.Count > 0 Then
        If isStartIdx1 Then
            ReDim arr(1 To coll.Count)
        Else
            ReDim arr(0 To coll.Count - 1)
        End If
        idx = LBound(arr)
        For Each item In coll
            ' Use "Set" when assigning objects.
            If IsObject(item) Then
                Set arr(idx) = item
            Else
                arr(idx) = item
            End If
            idx = idx + 1
        Next
    Else
        arr = VBA.Array()
    End If
    Collection2Array = arr
End Function


Private Sub InsertionSortJaggedArray(ByRef arr As Variant, _
    Optional ByVal reverse As Boolean = False, _
    Optional ByVal strSort As Boolean = False, _
    Optional ByVal ignoreCase As Boolean = True)
    
    ' Sorts a jagged array using the Insertion Sort algorithm based on the first element of each nested array.
    '   e.g., [[1, "A"], [3, "B"], [2, "C"]] -> [[1, "A"], [2, "C"], [3, "B"]]
    '   Does not affect the relative order of items with the same numeric value
    '   e.g., [[3, "C"], [3, "A"], [1, "A"], [3, "B"]] -> [[1, "A"], [3, "C"], [3, "A"], [3, "B"]]
    ' reverse: Set to True for descending order.
    '   e.g., [[1, "A"], [3, "B"], [2, "C"]] -> [[3, "B"], [2, "C"], [1, "A"]]
    ' strSort: Set to True for string-based comparison, False for numeric comparison.
    ' ignoreCase: Valid only when strSort is True. Set to True to perform case-insensitive comparison.
    ' Dependency: DynamicCompare
    If Not IsArray(arr) Then Err.Raise Number:=13
    Dim minIndex As Long
    Dim maxIndex As Long
    Dim idxToRef1 As Long
    Dim idxToRef2 As Long
    Dim op As String
    
    If reverse Then
        op = "<"
    Else
        op = ">"
    End If
    
    minIndex = LBound(arr)
    maxIndex = UBound(arr)
    Dim i As Long, j As Long
    Dim swap As Variant
    For i = minIndex + 1 To maxIndex
        swap = arr(i)
        For j = i - 1 To minIndex Step -1
            idxToRef1 = LBound(arr(j))
            idxToRef2 = LBound(swap)
            If DynamicCompare(arr(j)(idxToRef1), swap(idxToRef2), op, strSort, ignoreCase) Then
                arr(j + 1) = arr(j)
            Else
                Exit For
            End If
        Next
        arr(j + 1) = swap
    Next
End Sub


Private Function DynamicCompare(ByVal a As Variant, ByVal b As Variant, ByVal op As String, _
    Optional ByVal shouldStrComp As Boolean = False, Optional ByVal ignoreCase As Boolean = True) As Boolean
    ' Performs dynamic comparison using a string representation of an operator.
    ' a, b: Values to compare.
    ' op: Comparison operator as a string (">", ">=", "<", "<=", "=", "<>").
    ' shouldStrComp: Set to True for string comparison mode, False for numeric/default comparison.
    ' ignoreCase: Valid only when shouldStrComp is True. Set to True to ignore case sensitivity.
    Dim result As Boolean
    Dim compareMode As VbCompareMethod
    
    If shouldStrComp Then
        If ignoreCase Then
            compareMode = vbTextCompare
        Else
            compareMode = vbBinaryCompare
        End If
        
        Select Case op
            Case ">"
                result = StrComp(a, b, compareMode) > 0
            Case ">="
                result = StrComp(a, b, compareMode) >= 0
            Case "<"
                result = StrComp(a, b, compareMode) < 0
            Case "<="
                result = StrComp(a, b, compareMode) <= 0
            Case "="
                result = StrComp(a, b, compareMode) = 0
            Case "<>"
                result = StrComp(a, b, compareMode) <> 0
            Case Else
                Err.Raise vbObjectError, , "Unknown operator: " & op
        End Select
    Else
        Select Case op
            Case ">"
                result = (a > b)
            Case ">="
                result = (a >= b)
            Case "<"
                result = (a < b)
            Case "<="
                result = (a <= b)
            Case "="
                result = (a = b)
            Case "<>"
                result = (a <> b)
            Case Else
                Err.Raise vbObjectError, , "Unknown operator: " & op
        End Select
    End If
    DynamicCompare = result
End Function

Private Function CollContainsKey(ByVal coll As Collection, ByVal strKey As String) As Boolean
    ' Check if a specific key exists in the Collection
    CollContainsKey = False
    If coll Is Nothing Then Exit Function
    If coll.Count = 0 Then Exit Function
     
    On Error GoTo Exception
    Call coll.item(strKey)
    On Error GoTo 0
    CollContainsKey = True
    
    Exit Function
Exception:
    CollContainsKey = False
    Exit Function
End Function


Private Sub ExtendCollection(ByRef originalColl As Variant, ByVal additionalColl As Variant)
    ' Merges two Collections or ArrayLists (modifies the original object in-place).
    ' Keys from originalColl are preserved, but keys from additionalColl will be lost.
    ' Compatible with any object that has an .Add method (excluding Dictionaries).
    
    ' Parameters:
    '   originalColl   : The target collection to be extended.
    '   additionalColl : The collection containing items to add.
    Dim item As Variant
    For Each item In additionalColl
        originalColl.Add item
    Next item
End Sub

Private Function JoinCollection(ByVal coll As Collection, Optional ByVal delimiter As String = "") As String
    Dim arr() As Variant
    Dim result As String
    arr = Collection2Array(coll)
    result = Join(arr, delimiter)
    JoinCollection = result
End Function


Private Function AdjustIndent(ByVal text As String, ByVal indentSize As Long) As String
    '------------------------------------------------------------
    ' Function: AdjustIndent
    '
    ' Description:
    '   Adjusts the indentation of each line in a given text.
    '   The text may contain multiple lines separated by line breaks.
    '
    ' Specification:
    '   - Normalizes all line breaks to vbLf before processing.
    '   - If indentSize > 0:
    '       Adds the specified number of spaces to the beginning of each line.
    '       HOWEVER, if a line is empty (i.e., contains no characters),
    '       no spaces are added to that line.
    '   - If indentSize < 0:
    '       Removes the specified number of spaces from the beginning of each line.
    '       If a line has fewer leading spaces than the absolute value of indentSize,
    '       all leading spaces are removed.
    '   - If indentSize = 0:
    '       Returns the text unchanged (except for normalized line breaks).
    '
    ' Parameters:
    '   text (String)       : Input text (may include line breaks).
    '   indentSize (Long)   : Number of spaces to add (positive) or remove (negative).
    '
    ' Returns:
    '   String              : Text with adjusted indentation.
    '------------------------------------------------------------
    Dim lines() As String
    Dim i As Long
    Dim spaces As String
    Dim removeCount As Long
    Dim leadingSpaces As Long
    
    ' Normalize line breaks
    text = VBA.Replace(text, vbCrLf, vbLf)
    text = VBA.Replace(text, vbCr, vbLf)
    
    ' Split into lines
    lines = Split(text, vbLf)
    
    If indentSize > 0 Then
        ' Add spaces (skip empty lines)
        spaces = String(indentSize, " ")
        For i = LBound(lines) To UBound(lines)
            If Len(lines(i)) > 0 Then
                lines(i) = spaces & lines(i)
            End If
        Next i
        
    ElseIf indentSize < 0 Then
        ' Remove spaces
        removeCount = -indentSize
        
        For i = LBound(lines) To UBound(lines)
            leadingSpaces = 0
            
            ' Count leading spaces
            Do While leadingSpaces < Len(lines(i)) _
                And Mid(lines(i), leadingSpaces + 1, 1) = " "
                leadingSpaces = leadingSpaces + 1
            Loop
            
            ' Remove spaces safely
            If leadingSpaces >= removeCount Then
                lines(i) = Mid(lines(i), removeCount + 1)
            Else
                lines(i) = Mid(lines(i), leadingSpaces + 1)
            End If
        Next i
    End If
    
    ' Join lines back
    AdjustIndent = Join(lines, vbLf)
End Function

