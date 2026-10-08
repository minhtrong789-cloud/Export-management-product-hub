Attribute VB_Name = "Extension_UITPrebuiltThemes"
Option Compare Database
Option Explicit
'@Folder "Extensions.UI"

'====================================================================
' MODULE    : Extension_UITPrebuiltThemes
' PURPOSE   : Design Tokens (Colors) & Semantic Theme Presets (Level 2)
' NOTE      : All presets call Level 1 atomic functions in Extension_UITheme
'             and override only semantic styling properties.
'====================================================================

'--------------------------------------------------------------------
' SECTION 1: DESIGN TOKENS (COLORS)
'--------------------------------------------------------------------
Public Const COLOR_NAVY_DEEP As String = "#2F3C7E"   ' The Core Identity (Active/Headers)
Public Const COLOR_NAVY_LIGHT As String = "#485696"  ' The Hover State
Public Const COLOR_NAVY_DARK As String = "#1C2555"   ' The Pressed State
Public Const COLOR_SHELL_PINK As String = "#FBEBEB"  ' Dashboard Sidebar & Hub BG

Public Const COLOR_ACTION_PINK As String = "#F96167" ' The CTA / Power Button (Coral)
Public Const COLOR_ACTION_HOVER As String = "#FA7A7F"
Public Const COLOR_ACTION_DARK As String = "#E04B51"

Public Const COLOR_PROSPERITY_GREEN As String = "#1B5E20" ' Success / Export
Public Const COLOR_WARNING_RED As String = "#E53935"      ' Danger / Clear
Public Const COLOR_SECONDARY_SLATE As String = "#455A64"   ' Administration / Import

Public Const COLOR_LOGISTICS_TEAL As String = "#00897B"    ' Shipping / Labels
Public Const COLOR_LOGISTICS_HOVER As String = "#26A69A"
Public Const COLOR_LOGISTICS_DARK As String = "#00695C"


'--------------------------------------------------------------------
' SECTION 2: BUTTON PRESETS
'--------------------------------------------------------------------

' Purpose: Preset for primary action buttons (Save, New Record).
Public Sub gui_styleButton_Power(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleButton(formName, ControlName, COLOR_ACTION_PINK)
    With Forms(formName).Controls(ControlName)
        .HoverColor = HexToRBGLong(COLOR_ACTION_HOVER)
        .PressedColor = HexToRBGLong(COLOR_ACTION_DARK)
        .fontSize = 9: .FontWeight = FontWeight_Bold
    End With
End Sub

' Purpose: Preset for header ribbon utility/navigation buttons.
Public Sub gui_styleButton_Header(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleButton(formName, ControlName, COLOR_NAVY_DEEP)
    With Forms(formName).Controls(ControlName)
        .HoverColor = HexToRBGLong(COLOR_NAVY_LIGHT)
        .PressedColor = HexToRBGLong(COLOR_NAVY_DARK)
        .fontSize = 9: .FontWeight = FontWeight_Bold
    End With
End Sub

' Purpose: Preset for success action buttons (Export, Confirm).
Public Sub gui_styleButton_Success(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleButton(formName, ControlName, COLOR_PROSPERITY_GREEN)
    With Forms(formName).Controls(ControlName)
        .HoverColor = HexToRBGLong("#388E3C")
        .PressedColor = HexToRBGLong("#0A210B")
    End With
End Sub

' Purpose: Preset for danger action buttons (Delete, Clear).
Public Sub gui_styleButton_Danger(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleButton(formName, ControlName, COLOR_WARNING_RED)
    With Forms(formName).Controls(ControlName)
        .HoverColor = HexToRBGLong("#EF5350")
        .PressedColor = HexToRBGLong("#C62828")
    End With
End Sub

' Purpose: Preset for secondary administrative buttons (Import, Settings).
Public Sub gui_styleButton_Secondary(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleButton(formName, ControlName, COLOR_SECONDARY_SLATE)
    With Forms(formName).Controls(ControlName)
        .HoverColor = HexToRBGLong("#607D8B")
        .PressedColor = HexToRBGLong("#263238")
    End With
End Sub

' Purpose: Preset for logistics and shipping action buttons (Print Labels).
Public Sub gui_styleButton_Logistics(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleButton(formName, ControlName, COLOR_LOGISTICS_TEAL)
    With Forms(formName).Controls(ControlName)
        .HoverColor = HexToRBGLong(COLOR_LOGISTICS_HOVER)
        .PressedColor = HexToRBGLong(COLOR_LOGISTICS_DARK)
    End With
End Sub

' Purpose: Preset for outlined ghost command buttons.
Public Sub gui_styleButton_Ghost(ByVal formName As String, ByVal ControlName As String, _
                        Optional ByVal hexColor As String = COLOR_NAVY_DEEP, _
                        Optional ByVal hexHoverBase As String = "#F8F9FA")
    Call gui_styleButton(formName, ControlName, hexColor)
    With Forms(formName).Controls(ControlName)
        .BackColor = vbWhite
        .BorderWidth = 1
        .BorderColor = HexToRBGLong(hexColor)
        .BorderStyle = BorderStyle_Solid
        .ForeColor = HexToRBGLong(hexColor)
        .HoverColor = HexToRBGLong(hexHoverBase)
        .PressedColor = HexToRBGLong("#E0E0E0")
        .HoverForeColor = HexToRBGLong(hexColor)
        .PressedForeColor = HexToRBGLong(hexColor)
        .fontSize = 9
    End With
End Sub

' Purpose: Preset for transparent backdoor utility buttons.
Public Sub gui_styleButton_Stealth(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleButton(formName, ControlName, "#000000")
    With Forms(formName).Controls(ControlName)
        .Caption = "": .ForeColor = HexToRBGLong("#AF3D5E")
        .BackStyle = BackStyle_Transparent: .Transparent = False: .BorderStyle = BorderStyle_Transparent
        .HorizontalAnchor = acHorizontalAnchorLeft: .VerticalAnchor = acVerticalAnchorBottom
        .HoverColor = HexToRBGLong("#F8D8D8"): .PressedColor = HexToRBGLong("#7F7F7F")
        .HoverForeColor = HexToRBGLong("#AF3D5E"): .PressedForeColor = HexToRBGLong("#AF3D5E")
    End With
End Sub


'--------------------------------------------------------------------
' SECTION 3: TOGGLE BUTTON PRESETS
'--------------------------------------------------------------------

' Purpose: Preset for sentinel toggle buttons (Trash/Archive switches).
Public Sub gui_styleToggleButton_Sentinel(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleToggleButton(formName, ControlName, "#2F3C7E")
    With Forms(formName).Controls(ControlName)
        .BackColor = HexToRBGLong("#2F3C7E")
        .HoverColor = HexToRBGLong("#3B4A9A")
        .PressedColor = HexToRBGLong("#1A224D")
        .ForeColor = HexToRBGLong("#FF4D4F")
        .PressedForeColor = vbWhite
        .HoverForeColor = vbWhite
        .BorderStyle = 0
    End With
End Sub


'--------------------------------------------------------------------
' SECTION 4: LABEL PRESETS
'--------------------------------------------------------------------

' Purpose: Preset for header title labels.
Public Sub gui_styleLabel_Title(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleLabel(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .fontSize = 16: .FontWeight = FontWeight_Bold: .ForeColor = vbWhite
    End With
End Sub

' Purpose: Preset for metadata and record count labels.
Public Sub gui_styleLabel_Metadata(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleLabel(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .fontSize = 9: .ForeColor = HexToRBGLong("#888888")
    End With
End Sub

' Purpose: Preset for numeric and total value labels.
Public Sub gui_styleLabel_Value(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleLabel(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .FontWeight = FontWeight_Bold
        .ForeColor = HexToRBGLong("#333333")
        .TextAlign = TextAlign_Right
    End With
End Sub

' Purpose: Preset for watermark and placeholder labels.
Public Sub gui_styleLabel_Watermark(ByVal formName As String, ByVal ControlName As String, _
                           Optional ByVal fontSize As Integer = 12)
    Call gui_styleLabel(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .fontSize = fontSize: .FontItalic = True: .FontWeight = FontWeight_Bold
        .ForeColor = HexToRBGLong("#CCCCCC")
        .TextAlign = TextAlign_Center
    End With
End Sub

' Purpose: Preset for grid column header labels.
Public Sub gui_styleLabel_GridHeader(ByVal formName As String, ByVal ControlName As String, _
                            Optional ByVal captionText As String = "")
    Call gui_styleLabel(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        If captionText <> "" Then .Caption = captionText
        .fontSize = 8: .FontWeight = FontWeight_Bold
        .ForeColor = HexToRBGLong("#6B7280")
    End With
End Sub

' Purpose: Preset for section title labels on card containers.
Public Sub gui_styleLabel_SectionTitle(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleLabel(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .FontWeight = FontWeight_Bold
        .ForeColor = HexToRBGLong(COLOR_NAVY_DEEP)
        .fontSize = 11
    End With
End Sub


'--------------------------------------------------------------------
' SECTION 5: RECTANGLE PRESETS
'--------------------------------------------------------------------

' Purpose: Preset for top header brand bar rectangle.
Public Sub gui_styleRectangle_BrandBar(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleRectangle(formName, ControlName, COLOR_NAVY_DEEP)
    Forms(formName).Controls(ControlName).HorizontalAnchor = acHorizontalAnchorBoth
End Sub

' Purpose: Preset for sidebar navigation background rectangle.
Public Sub gui_styleRectangle_Sidebar(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleRectangle(formName, ControlName, COLOR_SHELL_PINK)
    Forms(formName).Controls(ControlName).VerticalAnchor = acVerticalAnchorBoth
End Sub

' Purpose: Preset for local hub navigation blade background rectangle.
Public Sub gui_styleRectangle_Blade(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleRectangle(formName, ControlName, COLOR_SHELL_PINK)
    Forms(formName).Controls(ControlName).VerticalAnchor = acVerticalAnchorBoth
End Sub

' Purpose: Preset for card container background rectangle.
Public Sub gui_styleRectangle_Card(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleRectangle(formName, ControlName, "#F8F9FA")
    With Forms(formName).Controls(ControlName)
        .BorderWidth = 1
        .BorderStyle = BorderStyle_Solid
        .BorderColor = HexToRBGLong("#E0E0E0")
    End With
End Sub

' Purpose: Preset for sidebar card active indicator bar.
Public Sub gui_styleRectangle_CardIndicator(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleRectangle(formName, ControlName, COLOR_NAVY_DEEP)
    With Forms(formName).Controls(ControlName)
        .BorderWidth = 0: .BorderStyle = BorderStyle_Transparent: .Visible = False
    End With
End Sub


'--------------------------------------------------------------------
' SECTION 6: LINE PRESETS
'--------------------------------------------------------------------

' Purpose: Preset for section divider lines.
Public Sub gui_styleLine_Divider(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleLine(formName, ControlName, "#E5E7EB")
    Forms(formName).Controls(ControlName).HorizontalAnchor = acHorizontalAnchorBoth
End Sub


'--------------------------------------------------------------------
' SECTION 7: COMBOBOX PRESETS
'--------------------------------------------------------------------

' Purpose: Preset for standard dropdown comboboxes.
Public Sub gui_styleComboBox_Standard(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleComboBox(formName, ControlName)
End Sub

' Purpose: Preset for logistics measurement and dimension comboboxes.
Public Sub gui_styleComboBox_Logistics(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleComboBox(formName, ControlName)
    Forms(formName).Controls(ControlName).BackColor = HexToRBGLong("#F8F9FF")
End Sub

' Purpose: Preset for transparent lookup comboboxes.
Public Sub gui_styleComboBox_Lookup(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleComboBox(formName, ControlName)
    Forms(formName).Controls(ControlName).BackStyle = BackStyle_Transparent
End Sub


'--------------------------------------------------------------------
' SECTION 8: TEXTBOX PRESETS
'--------------------------------------------------------------------

' Purpose: Preset for standard data entry textboxes.
Public Sub gui_styleTextBox_Base(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleTextBox(formName, ControlName)
End Sub

' Purpose: Preset for search input textboxes with magnifying format.
Public Sub gui_styleTextBox_Magnifying(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleTextBox(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .fontSize = 11
        .ForeColor = HexToRBGLong(COLOR_NAVY_DEEP)
        .BorderColor = HexToRBGLong("#E0E0E0")
    End With
End Sub

' Purpose: Preset for primary title textbox fields.
Public Sub gui_styleTextBox_HubTitle(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleTextBox(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .fontSize = 12
        .FontWeight = FontWeight_Bold
        .ForeColor = HexToRBGLong(COLOR_NAVY_DEEP)
        .BackStyle = BackStyle_Transparent
        .BorderStyle = BorderStyle_Transparent
    End With
End Sub

' Purpose: Preset for right-aligned numeric/currency textboxes.
Public Sub gui_styleTextBox_Value(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleTextBox(formName, ControlName)
    Forms(formName).Controls(ControlName).TextAlign = TextAlign_Right
End Sub

' Purpose: Preset for fixed-width textboxes (barcodes, serial numbers).
Public Sub gui_styleTextBox_Machine(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleTextBox(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .FontName = "Courier New"
        .FontWeight = FontWeight_Bold
    End With
End Sub

' Purpose: Preset for logistics dimension and measurement textboxes.
Public Sub gui_styleTextBox_Logistics(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleTextBox(formName, ControlName)
    Forms(formName).Controls(ControlName).BackColor = HexToRBGLong("#F8F9FF")
End Sub

' Purpose: Preset for sidebar card title textboxes.
Public Sub gui_styleTextBox_CardTitle(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleTextBox(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .BackColor = vbWhite
        .BackStyle = BackStyle_Transparent: .BorderStyle = BorderStyle_Transparent
        .ForeColor = HexToRBGLong(COLOR_NAVY_DEEP)
        .fontSize = 10: .FontWeight = FontWeight_Bold: .FontName = "Segoe UI Semibold"
    End With
End Sub

' Purpose: Preset for sidebar card subtitle textboxes.
Public Sub gui_styleTextBox_CardSubtitle(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleTextBox(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .BackColor = vbWhite
        .BackStyle = BackStyle_Transparent: .BorderStyle = BorderStyle_Transparent
        .ForeColor = HexToRBGLong(COLOR_NAVY_LIGHT)
        .fontSize = 9: .FontName = "Segoe UI"
    End With
End Sub

' Purpose: Preset for multiline document editor textboxes with scrollbar.
Public Sub gui_styleTextBox_Document(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleTextBox(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .ScrollBars = 2
        .TextAlign = TextAlign_Left
        .BorderColor = HexToRBGLong("#E0E0E0")
        .BackColor = HexToRBGLong("#F8F9FA")
    End With
End Sub


'--------------------------------------------------------------------
' SECTION 9: SUBFORM PRESETS
'--------------------------------------------------------------------

' Purpose: Preset for main data subform grid container.
Public Sub gui_styleSubform_Theater(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleSubform(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .BorderColor = HexToRBGLong("#C7CCEA")
        .HorizontalAnchor = acHorizontalAnchorBoth
        .VerticalAnchor = acVerticalAnchorBoth
    End With
End Sub

' Purpose: Preset for borderless transparent subform container.
Public Sub gui_styleSubform_Ghost(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleSubform(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .BorderWidth = 0: .BorderStyle = BorderStyle_Transparent
        .HorizontalAnchor = acHorizontalAnchorBoth: .VerticalAnchor = acVerticalAnchorBoth
    End With
End Sub


'--------------------------------------------------------------------
' SECTION 10: IMAGE PRESETS
'--------------------------------------------------------------------

' Purpose: Preset for product image preview container.
Public Sub gui_styleImage_Showcase(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleImage(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .BorderColor = HexToRBGLong("#C7CCEA")
        .HorizontalAnchor = acHorizontalAnchorRight
        .VerticalAnchor = acVerticalAnchorBoth
    End With
End Sub

' Purpose: Preset for transparent brand logo images.
Public Sub gui_styleImage_Logo(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleImage(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .BackStyle = BackStyle_Transparent: .BorderStyle = BorderStyle_Transparent: .BorderWidth = 0
    End With
End Sub


'--------------------------------------------------------------------
' SECTION 11: OPTION GROUP & CONTROL PRESETS
'--------------------------------------------------------------------

' Purpose: Preset for clean silver-bordered option group frame.
Public Sub gui_styleOptionGroup_Clean(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleOptionGroup(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .BorderColor = HexToRBGLong("#BDC3C7")
        .BackStyle = BackStyle_Transparent
    End With
End Sub

' Purpose: Preset for clean silver-bordered option button.
Public Sub gui_styleOptionButton_Clean(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleOptionButton(formName, ControlName)
    Forms(formName).Controls(ControlName).BorderColor = HexToRBGLong("#BDC3C7")
End Sub

' Purpose: Preset for standard navy checkbox.
Public Sub gui_styleCheckBox_Standard(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleCheckBox(formName, ControlName)
    Forms(formName).Controls(ControlName).BorderColor = HexToRBGLong(COLOR_NAVY_DEEP)
End Sub


'--------------------------------------------------------------------
' SECTION 12: NAVIGATION CONTROL PRESETS
'--------------------------------------------------------------------

' Purpose: Preset for dashboard navigation hub tabs.
Public Sub gui_styleNavHub_Dashboard(ByVal formName As String, ByVal ControlName As String)
    Call gui_styleNavigationControl(formName, ControlName, _
                                   COLOR_SHELL_PINK, _
                                   "#F8D8D8", _
                                   COLOR_NAVY_DEEP, _
                                   "#404040", _
                                   COLOR_NAVY_DEEP, _
                                   "#FFFFFF")
End Sub
