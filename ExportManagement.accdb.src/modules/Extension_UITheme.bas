Attribute VB_Name = "Extension_UITheme"
Option Compare Database
Option Explicit
'@Folder "Extensions.UI"

'====================================================================
' MODULE    : Extension_UITheme
' PURPOSE   : Form State Management, Control Capture & Styling Engine
' NOTE      : Base functions explicitly list all editable properties
'             (e.g. .Caption = .Caption) as a template manifest so that
'             theme presets know exactly which properties are available to override.
'====================================================================


'--------------------------------------------------------------------
' FUNCTION  : HexToRBGLong
' PURPOSE   : Converts Hex string (#FFFFFF or #FFF) to VBA Long color.
'--------------------------------------------------------------------
Public Function HexToRBGLong(ByVal hexColor As String, Optional ByVal fallbackColor As Long = vbWhite) As Long
    On Error GoTo ErrHandler

    ' 1. CLEAN INPUT (Remove #, spaces, etc.)
    hexColor = Replace(Replace(Trim(hexColor), "#", ""), " ", "")

    ' 2. SUPPORT SHORT-HEX (e.g., #F00 -> #FF0000)
    If Len(hexColor) = 3 Then
        hexColor = Mid(hexColor, 1, 1) & Mid(hexColor, 1, 1) & _
                   Mid(hexColor, 2, 1) & Mid(hexColor, 2, 1) & _
                   Mid(hexColor, 3, 1) & Mid(hexColor, 3, 1)
    End If

    ' 3. VALIDATE LENGTH
    If Len(hexColor) <> 6 Then GoTo ErrHandler

    ' 4. CONVERT TO LONG
    HexToRBGLong = RGB(CLng("&H" & Mid(hexColor, 1, 2)), _
                       CLng("&H" & Mid(hexColor, 3, 2)), _
                       CLng("&H" & Mid(hexColor, 5, 2)))
    Exit Function

ErrHandler:
    HexToRBGLong = fallbackColor
End Function


'--------------------------------------------------------------------
' PROCEDURE : gui_styleForm
' PURPOSE   : Standardizes Form window chrome and Detail section.
'--------------------------------------------------------------------
Public Sub gui_styleForm(ByVal formName As String)
    If Not CurrentProject.AllForms(formName).IsLoaded Then
        DoCmd.OpenForm formName, acDesign, , , , acHidden
    End If

    With Forms(formName)
        ' I. Window Chrome & Behaviors
        .Caption = .Caption
        .NavigationCaption = .Caption
        .DefaultView = DefaultView_SingleForm
        .AllowFormView = True: .AllowDatasheetView = False: .AllowLayoutView = False
        .AutoCenter = True: .AutoResize = False: .FitToScreen = False
        .BorderStyle = WindowBorderStyle_Sizable
        .RecordSelectors = False: .NavigationButtons = False: .DividingLines = False
        .ScrollBars = ScrollBars_Neither: .ControlBox = True: .CloseButton = True: .MinMaxButtons = 3: .Moveable = True
        .DataEntry = False: .AllowAdditions = True: .AllowDeletions = True: .AllowEdits = True: .AllowFilters = True
        .PopUp = True: .Modal = False

        ' II. Detail Section
        With .Section(acDetail)
            .BackColor = vbWhite: .AlternateBackColor = vbWhite: .SpecialEffect = SpecialEffect_Flat
        End With
    End With
End Sub

'--------------------------------------------------------------------
' PROCEDURE : gui_styleButton
' PURPOSE   : Standardizes command button flat styling and base color.
'--------------------------------------------------------------------
Public Sub gui_styleButton(ByVal formName As String, ByVal ControlName As String, _
                  Optional ByVal hexColor As String = COLOR_NAVY_DEEP)
    If Not CurrentProject.AllForms(formName).IsLoaded Then
        DoCmd.OpenForm formName, acDesign, , , , acHidden
    End If

    With Forms(formName).Controls(ControlName)
        ' I. Identity & Typography (7 properties)
        .Caption = .Caption
        .FontName = "Segoe UI": .fontSize = 10: .Alignment = TextAlign_Center
        .FontWeight = FontWeight_SemiBold: .FontUnderline = False: .FontItalic = False: .ForeColor = vbWhite

        ' II. Visual Shell (Back & Borders) (9 properties)
        .UseTheme = True
        .BackColor = HexToRBGLong(hexColor): .BorderWidth = 0: .BorderColor = HexToRBGLong(hexColor)
        .HoverColor = HexToRBGLong(hexColor): .PressedColor = HexToRBGLong(hexColor)
        .HoverForeColor = vbWhite: .PressedForeColor = vbWhite
        .BackStyle = BackStyle_Normal: .Transparent = False: .BorderStyle = BorderStyle_Transparent

        ' III. Positioning & Anchors (2 properties)
        .HorizontalAnchor = acHorizontalAnchorLeft: .VerticalAnchor = acVerticalAnchorTop

        ' IV. Effects Engine (6 properties)
        .Bevel = 0: .Glow = 0: .Shadow = 0: .SoftEdges = 0: .Shape = 0: .Gradient = 0
    End With
End Sub

'--------------------------------------------------------------------
' PROCEDURE : gui_styleToggleButton
' PURPOSE   : Standardizes toggle button flat styling.
'--------------------------------------------------------------------
Public Sub gui_styleToggleButton(ByVal formName As String, ByVal ControlName As String, _
                        Optional ByVal hexColor As String = "#F0F0F0")
    If Not CurrentProject.AllForms(formName).IsLoaded Then
        DoCmd.OpenForm formName, acDesign, , , , acHidden
    End If

    With Forms(formName).Controls(ControlName)
        ' I. Identity & Typography (7 properties)
        .Caption = .Caption: .FontName = "Segoe UI": .fontSize = 9
        .FontWeight = FontWeight_SemiBold: .FontUnderline = False: .FontItalic = False
        .ForeColor = HexToRBGLong("#333333")

        ' II. Visual Shell (Back & Borders) (9 properties)
        .BackColor = HexToRBGLong(hexColor): .BorderWidth = 1: .BorderColor = HexToRBGLong("#CCCCCC")
        .HoverColor = HexToRBGLong(hexColor): .PressedColor = HexToRBGLong(hexColor)
        .HoverForeColor = HexToRBGLong("#333333"): .PressedForeColor = vbWhite
        .UseTheme = True: .BorderStyle = BorderStyle_Solid

        ' IV. Effects Engine (6 properties)
        .Bevel = 0: .Glow = 0: .Shadow = 0: .SoftEdges = 0: .Shape = 0: .Gradient = 0
    End With
End Sub

'--------------------------------------------------------------------
' PROCEDURE : gui_styleLabel
' PURPOSE   : Standardizes label typography and transparent background.
'--------------------------------------------------------------------
Public Sub gui_styleLabel(ByVal formName As String, ByVal ControlName As String)
    If Not CurrentProject.AllForms(formName).IsLoaded Then
        DoCmd.OpenForm formName, acDesign, , , , acHidden
    End If

    With Forms(formName).Controls(ControlName)
        ' I. Identity & Typography (8 properties)
        .Caption = .Caption
        .FontName = "Segoe UI": .fontSize = 10: .TextAlign = TextAlign_Left
        .FontWeight = FontWeight_Normal: .FontUnderline = False: .FontItalic = False: .ForeColor = HexToRBGLong("#333333")

        ' II. Visual Shell (6 properties)
        .BackColor = vbWhite: .BorderWidth = 0: .BorderColor = vbWhite
        .SpecialEffect = SpecialEffect_Flat: .BackStyle = BackStyle_Transparent: .BorderStyle = BorderStyle_Transparent

        ' III. Absolute Layout (Margins & Line Spacing) (5 properties)
        .TopMargin = 30: .BottomMargin = 0: .LeftMargin = 72: .RightMargin = 0: .LineSpacing = 0
    End With
End Sub

'--------------------------------------------------------------------
' PROCEDURE : gui_styleRectangle
' PURPOSE   : Standardizes rectangle background and border properties.
'--------------------------------------------------------------------
Public Sub gui_styleRectangle(ByVal formName As String, ByVal ControlName As String, _
                     Optional ByVal hexBackColor As String = "#F8F9FA")
    If Not CurrentProject.AllForms(formName).IsLoaded Then
        DoCmd.OpenForm formName, acDesign, , , , acHidden
    End If

    With Forms(formName).Controls(ControlName)
        ' I. Visual Shell (6 properties)
        .BackColor = HexToRBGLong(hexBackColor)
        .BorderWidth = 0
        .BorderColor = HexToRBGLong(hexBackColor)
        .BackStyle = BackStyle_Normal
        .BorderStyle = BorderStyle_Transparent
        .SpecialEffect = SpecialEffect_Flat
    End With
End Sub

'--------------------------------------------------------------------
' PROCEDURE : gui_styleLine
' PURPOSE   : Standardizes divider line properties.
'--------------------------------------------------------------------
Public Sub gui_styleLine(ByVal formName As String, ByVal ControlName As String, _
                Optional ByVal hexBorderColor As String = "#E5E7EB")
    If Not CurrentProject.AllForms(formName).IsLoaded Then
        DoCmd.OpenForm formName, acDesign, , , , acHidden
    End If

    With Forms(formName).Controls(ControlName)
        ' I. Visual Shell (4 properties)
        .BorderWidth = 1: .BorderColor = HexToRBGLong(hexBorderColor)
        .SpecialEffect = SpecialEffect_Flat: .BorderStyle = BorderStyle_Solid
    End With
End Sub

'--------------------------------------------------------------------
' PROCEDURE : gui_styleComboBox
' PURPOSE   : Standardizes combobox layout, font, and border properties.
'--------------------------------------------------------------------
Public Sub gui_styleComboBox(ByVal formName As String, ByVal ControlName As String)
    If Not CurrentProject.AllForms(formName).IsLoaded Then
        DoCmd.OpenForm formName, acDesign, , , , acHidden
    End If

    With Forms(formName).Controls(ControlName)
        ' I. Visual Shell (6 properties)
        .Format = ""
        .BackColor = vbWhite: .BorderWidth = 1: .BorderColor = HexToRBGLong("#CCCCCC")
        .SpecialEffect = SpecialEffect_Flat
        .BackStyle = BackStyle_Normal: .BorderStyle = BorderStyle_Solid

        ' II. Typography (8 properties)
        .FontName = "Segoe UI": .TextAlign = TextAlign_Left: .fontSize = 10
        .FontWeight = FontWeight_Normal: .FontUnderline = False: .FontItalic = False
        .ForeColor = HexToRBGLong("#333333"): .DisplayAsHyperlink = 0

        ' III. Absolute Layout (Margins) (4 properties)
        .TopMargin = 30: .BottomMargin = 0: .LeftMargin = 72: .RightMargin = 0
    End With
End Sub

'--------------------------------------------------------------------
' PROCEDURE : gui_styleTextBox
' PURPOSE   : Standardizes textbox layout, font, and border properties.
'--------------------------------------------------------------------
Public Sub gui_styleTextBox(ByVal formName As String, ByVal ControlName As String)
    If Not CurrentProject.AllForms(formName).IsLoaded Then
        DoCmd.OpenForm formName, acDesign, , , , acHidden
    End If

    With Forms(formName).Controls(ControlName)
        ' I. Identity & Typography (7 properties)
        .Format = ""
        .FontName = "Segoe UI": .fontSize = 10: .TextAlign = TextAlign_Left
        .FontWeight = FontWeight_Normal: .FontUnderline = False: .FontItalic = False
        .ForeColor = HexToRBGLong("#333333")

        ' II. Visual Shell (6 properties)
        .BackColor = vbWhite: .BorderWidth = 1: .BorderColor = HexToRBGLong("#CCCCCC")
        .SpecialEffect = SpecialEffect_Flat: .BackStyle = BackStyle_Normal: .BorderStyle = BorderStyle_Solid

        ' III. Absolute Layout (Margins & Line Spacing) (5 properties)
        .TopMargin = 30: .BottomMargin = 0: .LeftMargin = 72: .RightMargin = 0: .LineSpacing = 0

        ' IV. Interactive & Scroll Logic (2 properties)
        .ScrollBars = ScrollBars_Neither: .DisplayAsHyperlink = 0
    End With
End Sub


'--------------------------------------------------------------------
' PROCEDURE : gui_styleSubform
' PURPOSE   : Standardizes subform border and anchor properties.
'--------------------------------------------------------------------
Public Sub gui_styleSubform(ByVal formName As String, ByVal ControlName As String)
    If Not CurrentProject.AllForms(formName).IsLoaded Then
        DoCmd.OpenForm formName, acDesign, , , , acHidden
    End If

    With Forms(formName).Controls(ControlName)
        ' I. Visual Shell (4 properties)
        .BorderWidth = 1: .BorderColor = HexToRBGLong("#E0E0E0")
        .SpecialEffect = SpecialEffect_Flat: .BorderStyle = BorderStyle_Solid
    End With
End Sub

'--------------------------------------------------------------------
' PROCEDURE : gui_styleCheckBox
' PURPOSE   : Standardizes checkbox border and anchor properties.
'--------------------------------------------------------------------
Public Sub gui_styleCheckBox(ByVal formName As String, ByVal ControlName As String)
    If Not CurrentProject.AllForms(formName).IsLoaded Then
        DoCmd.OpenForm formName, acDesign, , , , acHidden
    End If

    With Forms(formName).Controls(ControlName)
        ' I. Visual Shell (4 properties)
        .BorderStyle = BorderStyle_Transparent
        .BorderWidth = 0
        .BorderColor = HexToRBGLong("#CCCCCC")
        .SpecialEffect = SpecialEffect_Flat
    End With
End Sub

'--------------------------------------------------------------------
' PROCEDURE : gui_styleOptionGroup
' PURPOSE   : Standardizes option group frame properties.
'--------------------------------------------------------------------
Public Sub gui_styleOptionGroup(ByVal formName As String, ByVal ControlName As String)
    If Not CurrentProject.AllForms(formName).IsLoaded Then
        DoCmd.OpenForm formName, acDesign, , , , acHidden
    End If

    With Forms(formName).Controls(ControlName)
        ' I. Visual Shell (6 properties)
        .BackStyle = BackStyle_Transparent
        .BackColor = vbWhite
        .BorderStyle = BorderStyle_Solid
        .BorderWidth = 1
        .BorderColor = HexToRBGLong("#CCCCCC")
        .SpecialEffect = SpecialEffect_Flat
    End With
End Sub

'--------------------------------------------------------------------
' PROCEDURE : gui_styleOptionButton
' PURPOSE   : Standardizes option button properties.
'--------------------------------------------------------------------
Public Sub gui_styleOptionButton(ByVal formName As String, ByVal ControlName As String)
    If Not CurrentProject.AllForms(formName).IsLoaded Then
        DoCmd.OpenForm formName, acDesign, , , , acHidden
    End If

    With Forms(formName).Controls(ControlName)
        ' I. Visual Shell (4 properties)
        .BorderStyle = BorderStyle_Solid
        .BorderWidth = 1
        .BorderColor = HexToRBGLong("#CCCCCC")
        .SpecialEffect = SpecialEffect_Flat
    End With
End Sub

'--------------------------------------------------------------------
' PROCEDURE : gui_styleImage
' PURPOSE   : Standardizes image control border and anchor properties.
'--------------------------------------------------------------------
Public Sub gui_styleImage(ByVal formName As String, ByVal ControlName As String)
    If Not CurrentProject.AllForms(formName).IsLoaded Then
        DoCmd.OpenForm formName, acDesign, , , , acHidden
    End If

    With Forms(formName).Controls(ControlName)
        ' I. Visual Shell (6 properties)
        .BackColor = vbWhite: .BorderWidth = 1: .BorderColor = HexToRBGLong("#CCCCCC")
        .SpecialEffect = SpecialEffect_Flat
        .BackStyle = BackStyle_Normal: .BorderStyle = BorderStyle_Solid
    End With
End Sub

'--------------------------------------------------------------------
' PROCEDURE : gui_styleNavigationControl
' PURPOSE   : Standardizes navigation tab control and child tab buttons.
'--------------------------------------------------------------------
Public Sub gui_styleNavigationControl(ByVal formName As String, ByVal ControlName As String, _
                                      ByVal hexBase As String, _
                                      ByVal hexHover As String, _
                                      ByVal hexActive As String, _
                                      ByVal hexFore As String, _
                                      ByVal hexHoverFore As String, _
                                      ByVal hexActiveFore As String)
    If Not CurrentProject.AllForms(formName).IsLoaded Then
        DoCmd.OpenForm formName, acDesign, , , , acHidden
    End If

    Dim ctl As Access.Control
    Set ctl = Forms(formName).Controls(ControlName)

    ' 1. Container styling
    With ctl
        .BackStyle = BackStyle_Transparent: .BackColor = 0: .BorderStyle = BorderStyle_Transparent: .BorderWidth = 0: .BorderColor = 0
    End With

    ' 2. Child tab buttons styling
    Dim obj As Object
    For Each obj In ctl.Tabs
        With obj
            On Error Resume Next
            ' State Colors (7)
            .BackColor = HexToRBGLong(hexBase): .BorderWidth = 0: .BorderColor = HexToRBGLong(hexBase)
            .HoverColor = HexToRBGLong(hexHover): .PressedColor = HexToRBGLong(hexActive)
            .HoverForeColor = HexToRBGLong(hexHoverFore): .PressedForeColor = HexToRBGLong(hexActiveFore)

            ' Typography (7)
            .FontName = "Segoe UI": .TextAlign = TextAlign_Center: .fontSize = 9: .FontWeight = FontWeight_Bold
            .FontUnderline = False: .FontItalic = False: .ForeColor = HexToRBGLong(hexFore)

            ' Layout & Positioning (3)
            .BorderStyle = BorderStyle_Transparent: .BackStyle = BackStyle_Normal: .Transparent = False

            ' Effects Engine (6)
            .Bevel = 0: .Glow = 0: .Shadow = 0: .SoftEdges = 0: .Shape = 0: .Gradient = 0
            On Error GoTo 0
        End With
    Next obj
End Sub
