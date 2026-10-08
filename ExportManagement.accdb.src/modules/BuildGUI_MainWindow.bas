Attribute VB_Name = "BuildGUI_MainWindow"
Option Compare Database
Option Explicit
'@Folder "App.Dashboard"

'====================================================================
' MODULE    : BuildGUI_MainWindow
' PURPOSE   : Form Wireframe and Styling Generator for Main Window Shells.
' DEPENDENCIES: Extension_UIBuilder, Extension_UITheme, Extension_UITPrebuiltThemes
'====================================================================

'--------------------------------------------------------------------
' PROCEDURE : GenerateAll_MainWindow
' PURPOSE   : Batch master trigger to build & style all Main Window forms.
'--------------------------------------------------------------------
Public Sub GenerateAll_MainWindow()
    Call Generate_frmDashboard
    Call Generate_dlgLanguagePicker
End Sub

'====================================================================
' FORM 1: frmDashboard (Master Application Shell)
' SPEC  : Navigation sidebar and main viewport host (12.5" x 6.0")
'====================================================================

Public Sub Generate_frmDashboard()
    ' PURPOSE: One-stop generator to build wireframe and apply style.
    Call BuildForm_frmDashboard
    Call StyleForm_frmDashboard
End Sub

'--------------------------------------------------------------------
' 1.1 WIREFRAME BUILDER: frmDashboard
'--------------------------------------------------------------------
Public Sub BuildForm_frmDashboard()
    Const FORM_NAME As String = "frmDashboard"

    ' 1. SIZE FORM CANVAS (12.5" x 6.0")
    Call gui_buildForm(FORM_NAME, gui_buildFormSizeDef(Width:=12.5, Height:=6#))

    ' 2. SIDEBAR WIREFRAME (L=0 to 1.5")
    Call gui_buildControl(FORM_NAME, "recNavi", acRectangle, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=1.5, Height:=6#, _
                                HorizontalAnchor:=acHorizontalAnchorLeft, VerticalAnchor:=acVerticalAnchorBoth)) ' Right=1.50, Bottom=6.00

    Call gui_buildControl(FORM_NAME, "imaSoftWareLogo", acImage, _
        gui_buildControlSizeDef(Top:=0.12, Left:=0.25, Width:=1#, Height:=1#, _
                                HorizontalAnchor:=acHorizontalAnchorLeft, VerticalAnchor:=acVerticalAnchorTop)) ' Right=1.25, Bottom=1.12

    ' Sidebar Navigation Hub (119 = acNavigationControl)
    Call gui_buildControl(FORM_NAME, "navctrlDashboard", 119, _
        gui_buildControlSizeDef(Top:=1.4, Left:=0.05, Width:=1.4, Height:=3.3, _
                                HorizontalAnchor:=acHorizontalAnchorLeft, VerticalAnchor:=acVerticalAnchorBoth)) ' Right=1.45, Bottom=4.70

    Call gui_buildControl(FORM_NAME, "cmdCallAccessWindow", acCommandButton, _
        gui_buildControlSizeDef(Top:=5.6, Left:=0.05, Width:=0.32, Height:=0.35, _
                                HorizontalAnchor:=acHorizontalAnchorLeft, VerticalAnchor:=acVerticalAnchorBottom)) ' Right=0.37, Bottom=5.95

    ' 3. MAIN VIEWPORT HOST (L=1.52", W=10.98")
    Call gui_buildControl(FORM_NAME, "sfDashboard", acSubform, _
        gui_buildControlSizeDef(Top:=0, Left:=1.52, Width:=10.98, Height:=6#, _
                                HorizontalAnchor:=acHorizontalAnchorBoth, VerticalAnchor:=acVerticalAnchorBoth)) ' Right=12.50, Bottom=6.00

    ' [FINALIZE]: Save & Release
    DoCmd.Close acForm, FORM_NAME, acSaveYes
End Sub

'--------------------------------------------------------------------
' 1.2 AESTHETICS & STYLING: frmDashboard
'--------------------------------------------------------------------
Public Sub StyleForm_frmDashboard()
    Const FORM_NAME As String = "frmDashboard"

    ' 1. OPEN IN DESIGN VIEW
    If Not CurrentProject.AllForms(FORM_NAME).IsLoaded Then
        DoCmd.OpenForm FORM_NAME, acDesign, , , , acHidden
    End If

    ' 2. LEVEL 1 & 2 BASE STYLING (Theme Presets)
    Call gui_styleForm(FORM_NAME)
    Call gui_styleRectangle_Sidebar(FORM_NAME, "recNavi")
    Call gui_styleImage_Logo(FORM_NAME, "imaSoftWareLogo")
    Call gui_styleNavHub_Dashboard(FORM_NAME, "navctrlDashboard")
    Call gui_styleButton_Stealth(FORM_NAME, "cmdCallAccessWindow")
    Call gui_styleSubform_Ghost(FORM_NAME, "sfDashboard")

    ' 3. LEVEL 3 CONTEXTUAL OVERRIDES (Form-Specific Properties)
    With Forms(FORM_NAME)
        ' Form Chrome & Modality
        .Caption = "Export Manager Software"
        .NavigationCaption = "Export Manager"
        .Modal = False
        .MinMaxButtons = 3

        ' Control Contextual Overrides
        .Controls("cmdCallAccessWindow").ControlTipText = "Toggle MS Access Workspace"
    End With

    ' 4. SAVE & CLOSE
    DoCmd.Close acForm, FORM_NAME, acSaveYes
End Sub

'====================================================================
' FORM 2: dlgLanguagePicker (Modal Dialog)
' SPEC  : Dedicated compact language picker dialog (3.5" x 2.0")
'====================================================================

Public Sub Generate_dlgLanguagePicker()
    ' PURPOSE: One-stop generator to build wireframe and apply style.
    Call BuildForm_dlgLanguagePicker
    Call StyleForm_dlgLanguagePicker
End Sub

'--------------------------------------------------------------------
' 2.1 WIREFRAME BUILDER: dlgLanguagePicker
'--------------------------------------------------------------------
Public Sub BuildForm_dlgLanguagePicker()
    Const FORM_NAME As String = "dlgLanguagePicker"

    ' 1. SIZE FORM CANVAS (3.5" x 2.0")
    Call gui_buildForm(FORM_NAME, gui_buildFormSizeDef(Width:=3.5, Height:=2#))

    ' 2. CONTROLS GEOMETRY
    Call gui_buildControl(FORM_NAME, "lblTitle", acLabel, _
        gui_buildControlSizeDef(Top:=0.1, Left:=0.2, Width:=3.1, Height:=0.3)) ' Right=3.30, Bottom=0.40

    ' Option Group Container
    Call gui_buildControl(FORM_NAME, "fraLang", acOptionGroup, _
        gui_buildControlSizeDef(Top:=0.5, Left:=0.2, Width:=3.1, Height:=0.9)) ' Right=3.30, Bottom=1.40

    ' Option Buttons (Parented to fraLang)
    Call gui_buildControl(FORM_NAME, "optVN", acOptionButton, _
        gui_buildControlSizeDef(Top:=0.85, Left:=0.35, Width:=0.2, Height:=0.2), _
        parentName:="fraLang") ' Right=0.55, Bottom=1.05

    Call gui_buildControl(FORM_NAME, "optEN", acOptionButton, _
        gui_buildControlSizeDef(Top:=0.85, Left:=1.4, Width:=0.2, Height:=0.2), _
        parentName:="fraLang") ' Right=1.60, Bottom=1.05

    Call gui_buildControl(FORM_NAME, "optBI", acOptionButton, _
        gui_buildControlSizeDef(Top:=0.85, Left:=2.45, Width:=0.2, Height:=0.2), _
        parentName:="fraLang") ' Right=2.65, Bottom=1.05

    ' Option Labels
    Call gui_buildControl(FORM_NAME, "lblVN", acLabel, _
        gui_buildControlSizeDef(Top:=0.85, Left:=0.6, Width:=0.5, Height:=0.2)) ' Right=1.10, Bottom=1.05

    Call gui_buildControl(FORM_NAME, "lblEN", acLabel, _
        gui_buildControlSizeDef(Top:=0.85, Left:=1.65, Width:=0.5, Height:=0.2)) ' Right=2.15, Bottom=1.05

    Call gui_buildControl(FORM_NAME, "lblBI", acLabel, _
        gui_buildControlSizeDef(Top:=0.85, Left:=2.7, Width:=0.5, Height:=0.2)) ' Right=3.20, Bottom=1.05

    ' Dialog Action Buttons
    Call gui_buildControl(FORM_NAME, "btnOK", acCommandButton, _
        gui_buildControlSizeDef(Top:=1.5, Left:=0.7, Width:=1#, Height:=0.35)) ' Right=1.70, Bottom=1.85

    Call gui_buildControl(FORM_NAME, "btnCancel", acCommandButton, _
        gui_buildControlSizeDef(Top:=1.5, Left:=1.8, Width:=1#, Height:=0.35)) ' Right=2.80, Bottom=1.85

    ' [FINALIZE]: Save & Release
    DoCmd.Close acForm, FORM_NAME, acSaveYes
End Sub

'--------------------------------------------------------------------
' 2.2 AESTHETICS & STYLING: dlgLanguagePicker
'--------------------------------------------------------------------
Public Sub StyleForm_dlgLanguagePicker()
    Const FORM_NAME As String = "dlgLanguagePicker"

    ' 1. OPEN IN DESIGN VIEW
    If Not CurrentProject.AllForms(FORM_NAME).IsLoaded Then
        DoCmd.OpenForm FORM_NAME, acDesign, , , , acHidden
    End If
    Dim frm As Access.Form: Set frm = Forms(FORM_NAME)

    ' 2. LEVEL 1 & 2 BASE STYLING (Theme Presets)
    Call gui_styleForm(FORM_NAME)
    Call gui_styleLabel_Title(FORM_NAME, "lblTitle")
    Call gui_styleOptionGroup_Clean(FORM_NAME, "fraLang")
    Call gui_styleOptionButton_Clean(FORM_NAME, "optVN")
    Call gui_styleLabel(FORM_NAME, "lblVN")
    Call gui_styleOptionButton_Clean(FORM_NAME, "optEN")
    Call gui_styleLabel(FORM_NAME, "lblEN")
    Call gui_styleOptionButton_Clean(FORM_NAME, "optBI")
    Call gui_styleLabel(FORM_NAME, "lblBI")
    Call gui_styleButton_Success(FORM_NAME, "btnOK")
    Call gui_styleButton_Secondary(FORM_NAME, "btnCancel")

    ' 3. LEVEL 3 CONTEXTUAL OVERRIDES (Dialog Dimensions, Modality & Captions)
    With frm
        ' Modal Shell Configuration
        .Caption = "Export Configuration"
        .Modal = True
        .MinMaxButtons = 0
        .Section(acDetail).BackColor = RGB(245, 247, 250)
        .InsideWidth = 3.5 * TWIPS_PER_INCH
        .InsideHeight = 2.2 * TWIPS_PER_INCH

        ' Title & Typography
        With .Controls("lblTitle")
            .Caption = "SELECT EXPORT LANGUAGE"
            .ForeColor = RGB(44, 62, 80)
            .TextAlign = 2 ' Center
            .fontSize = 11
        End With

        ' Option Button Labels
        .Controls("lblVN").fontSize = 9: .Controls("lblVN").Caption = "Ti" & ChrW$(7871) & "ng Vi" & ChrW$(7879) & "t"
        .Controls("lblEN").fontSize = 9: .Controls("lblEN").Caption = "English"
        .Controls("lblBI").fontSize = 9: .Controls("lblBI").Caption = "Bilingual"

        ' Action Button Captions
        .Controls("btnOK").Caption = "CONFIRM"
        .Controls("btnCancel").Caption = "CANCEL"
    End With

    ' 4. SAVE & CLOSE
    DoCmd.Close acForm, FORM_NAME, acSaveYes
End Sub
