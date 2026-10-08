Attribute VB_Name = "Extension_UIBuilder"
Option Compare Database
Option Explicit
'@Folder "Extensions.UI"

'====================================================================
' MODULE    : Extension_UIBuilder
' PURPOSE   : Form Wireframe Layout & Control Injection Engine
'====================================================================

Public Type UIControlSizeDef
    Left             As Single
    Top              As Single
    Right            As Single
    Bottom           As Single
    Width            As Single
    Height           As Single
    HorizontalAnchor As AcHorizontalAnchor
    VerticalAnchor   As AcVerticalAnchor
End Type

Public Type UIFormSizeDef
    Width        As Single
    Height       As Single
    HeaderHeight As Single
    FooterHeight As Single
End Type

'--------------------------------------------------------------------
' PROCEDURE : gui_buildForm
' PURPOSE   : Builds form canvas and section dimensions (inches to twips).
'--------------------------------------------------------------------
Public Sub gui_buildForm(ByVal formName As String, ByRef formSize As UIFormSizeDef)
    On Error GoTo ErrHandler

    ' 1. VALIDATION & CALCULATION BLOCK (FAIL-FAST)
    If formSize.Width <= 0! Then
        Err.Raise vbObjectError + 2001, "gui_buildForm", "Form width must be greater than 0."
    End If
    If formSize.Height <= 0! Then
        Err.Raise vbObjectError + 2002, "gui_buildForm", "Form height must be greater than 0."
    End If
    If (formSize.HeaderHeight + formSize.FooterHeight) >= formSize.Height Then
        Err.Raise vbObjectError + 2003, "gui_buildForm", _
                  "Header and Footer height (" & (formSize.HeaderHeight + formSize.FooterHeight) & ") must be less than total Form height (" & formSize.Height & ")."
    End If

    Dim detailH As Single
    detailH = formSize.Height - (formSize.HeaderHeight + formSize.FooterHeight)

    ' 2. CREATE FORM IF NOT EXIST & OPEN IN DESIGN VIEW
    If Not formExist(formName) Then
        Dim newFrm As Access.Form:  Set newFrm = Application.CreateForm()
        Dim tempName As String:     tempName = newFrm.name
        DoCmd.Save acForm, tempName
        DoCmd.Close acForm, tempName
        DoCmd.Rename formName, acForm, tempName
    End If

    If Not CurrentProject.AllForms(formName).IsLoaded Then
        DoCmd.OpenForm formName, acDesign, , , , acHidden
    End If
    Dim frm As Access.Form: Set frm = Forms(formName)

    ' ENABLE HEADER FOOTER IF ANY
    If (formSize.HeaderHeight > 0! Or formSize.FooterHeight > 0!) Then
        Call enableSection(frm)
    End If

    ' 3. ASSIGN FORM CANVAS & SECTION SIZES (INCHES TO TWIPS)
    frm.Width = CLng(formSize.Width * TWIPS_PER_INCH)

    If formSize.HeaderHeight > 0! Then
        frm.Section(acHeader).Height = CLng(formSize.HeaderHeight * TWIPS_PER_INCH)
    End If

    If detailH > 0! Then
        frm.Section(acDetail).Height = CLng(detailH * TWIPS_PER_INCH)
    End If

    If formSize.FooterHeight > 0! Then
        frm.Section(acFooter).Height = CLng(formSize.FooterHeight * TWIPS_PER_INCH)
    End If

    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "gui_buildForm", Err.Description
End Sub

'--------------------------------------------------------------------
' PROCEDURE : gui_buildControl
' PURPOSE   : Injects or positions a control on the form canvas.
'--------------------------------------------------------------------
Public Sub gui_buildControl(ByVal formName As String, ByVal ControlName As String, _
                           ByVal controlType As AcControlType, ByRef controlSize As UIControlSizeDef, _
                           Optional ByVal sectionID As AcSection = acDetail, _
                           Optional ByVal parentName As String = "")
    On Error GoTo ErrHandler

    ' 1. VALIDATION BLOCK (FAIL-FAST)
    If controlSize.Width <= 0! Then
        Err.Raise vbObjectError + 3001, "gui_buildControl", "Control width must be greater than 0."
    End If
    If controlSize.Height <= 0! Then
        Err.Raise vbObjectError + 3002, "gui_buildControl", "Control height must be greater than 0."
    End If

    If Not formExist(formName) Then
        Err.Raise vbObjectError + 3003, "gui_buildControl", "Form '" & formName & "' does not exist."
    End If

    ' Ensure Form is open in Design View
    If Not CurrentProject.AllForms(formName).IsLoaded Then
        DoCmd.OpenForm formName, acDesign, , , , acHidden
    End If
    Dim frm As Access.Form: Set frm = Forms(formName)

    ' Validate Section existence
    If Not sectionExist(frm, sectionID) Then
        Err.Raise vbObjectError + 3004, "gui_buildControl", _
                  "Section " & sectionID & " is not enabled on Form '" & formName & "'. Please configure it in gui_buildForm first."
    End If

    ' 2. CREATE CONTROL IF MISSING (Native Access Engine)
    If Not controlExist(formName, ControlName) Then
        Dim newCtrl As Access.Control
        Set newCtrl = Application.CreateControl(formName, controlType, sectionID, parentName, "", 0, 0, 0, 0)
        newCtrl.name = ControlName
        If controlType = acLabel Then
            newCtrl.Caption = "Label"
        End If
    End If

    ' 3. ASSIGN GEOMETRY & ANCHORS (INCHES TO TWIPS)
    With frm.Controls(ControlName)
        .Left = CLng(controlSize.Left * TWIPS_PER_INCH)
        .Top = CLng(controlSize.Top * TWIPS_PER_INCH)
        .Width = CLng(controlSize.Width * TWIPS_PER_INCH)
        .Height = CLng(controlSize.Height * TWIPS_PER_INCH)
        .HorizontalAnchor = controlSize.HorizontalAnchor
        .VerticalAnchor = controlSize.VerticalAnchor
    End With

    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "gui_buildControl", Err.Description
End Sub

'--------------------------------------------------------------------
' FUNCTION  : gui_buildFormSizeDef
' PURPOSE   : Validates and constructs UIFormSizeDef structure.
'--------------------------------------------------------------------
Public Function gui_buildFormSizeDef(ByVal Width As Single, ByVal Height As Single, _
                                     Optional ByVal HeaderHeight As Single = 0!, _
                                     Optional ByVal FooterHeight As Single = 0!) As UIFormSizeDef
    If Width <= 0! Then
        Err.Raise vbObjectError + 2001, "gui_buildFormSizeDef", "Form width must be greater than 0."
    End If
    If Height <= 0! Then
        Err.Raise vbObjectError + 2002, "gui_buildFormSizeDef", "Form height must be greater than 0."
    End If
    If (HeaderHeight + FooterHeight) >= Height Then
        Err.Raise vbObjectError + 2003, "gui_buildFormSizeDef", _
                  "Header and Footer height (" & (HeaderHeight + FooterHeight) & ") must be less than total Form height (" & Height & ")."
    End If

    Dim fSize As UIFormSizeDef
    fSize.Width = Width
    fSize.Height = Height
    fSize.HeaderHeight = HeaderHeight
    fSize.FooterHeight = FooterHeight
    gui_buildFormSizeDef = fSize
End Function

'--------------------------------------------------------------------
' FUNCTION  : gui_buildControlSizeDef
' PURPOSE   : Calculates and constructs UIControlSizeDef structure.
'--------------------------------------------------------------------
Public Function gui_buildControlSizeDef(Optional ByVal Left As Single = 0!, Optional ByVal Top As Single = 0!, _
                                        Optional ByVal Right As Single = 0!, Optional ByVal Bottom As Single = 0!, _
                                        Optional ByVal Width As Single = 0!, Optional ByVal Height As Single = 0!, _
                                        Optional ByVal HorizontalAnchor As AcHorizontalAnchor = acHorizontalAnchorLeft, _
                                        Optional ByVal VerticalAnchor As AcVerticalAnchor = acVerticalAnchorTop) _
                                        As UIControlSizeDef
    Dim hasWidthHeight As Boolean
    Dim hasRightBottom As Boolean
    Dim calcWidth As Single, calcHeight As Single
    Dim calcRight As Single, calcBottom As Single
    Dim sizeDef As UIControlSizeDef

    hasWidthHeight = (Width > 0! And Height > 0!)
    hasRightBottom = (Right > 0! And Bottom > 0!)

    ' 1. VALIDATION BLOCK
    If Not (hasWidthHeight Or hasRightBottom) Then
        Err.Raise vbObjectError + 1001, "gui_buildControlSizeDef", _
                  "Invalid size definition: You must provide either (Width > 0 and Height > 0) OR (Right > Left and Bottom > Top)."
    ElseIf hasRightBottom Then
        If Right <= Left Then
            Err.Raise vbObjectError + 1002, "gui_buildControlSizeDef", "Right (" & Right & ") must be greater than Left (" & Left & ")."
        End If
        If Bottom <= Top Then
            Err.Raise vbObjectError + 1003, "gui_buildControlSizeDef", "Bottom (" & Bottom & ") must be greater than Top (" & Top & ")."
        End If
    End If

    ' 2. CALCULATION BLOCK
    If hasWidthHeight Then
        calcWidth = Width
        calcHeight = Height
        calcRight = Left + Width
        calcBottom = Top + Height
    Else
        calcRight = Right
        calcBottom = Bottom
        calcWidth = Right - Left
        calcHeight = Bottom - Top
    End If

    ' 3. ASSIGNMENT BLOCK (RECT order: Left, Top, Right, Bottom + Anchors)
    sizeDef.Left = Left
    sizeDef.Top = Top
    sizeDef.Right = calcRight
    sizeDef.Bottom = calcBottom
    sizeDef.Width = calcWidth
    sizeDef.Height = calcHeight
    sizeDef.HorizontalAnchor = HorizontalAnchor
    sizeDef.VerticalAnchor = VerticalAnchor

    gui_buildControlSizeDef = sizeDef
End Function

'--------------------------------------------------------------------
' PRIVATE HELPERS
'--------------------------------------------------------------------
' Purpose: Checks if form exists in CurrentProject.
Private Function formExist(ByVal formName As String) As Boolean
    On Error Resume Next
    Dim formObj As Object
    Set formObj = CurrentProject.AllForms(formName)
    formExist = (Not formObj Is Nothing)
    On Error GoTo 0
End Function

' Purpose: Checks if control exists on open form.
Private Function controlExist(ByVal formName As String, ByVal ControlName As String) As Boolean
    On Error Resume Next
    Dim ctrlObj As Object
    If formExist(formName) Then
        If CurrentProject.AllForms(formName).IsLoaded Then
            Set ctrlObj = Forms(formName).Controls(ControlName)
            controlExist = (Not ctrlObj Is Nothing)
        End If
    End If
    On Error GoTo 0
End Function

' Purpose: Checks if specific section is enabled on form.
Private Function sectionExist(ByVal frm As Access.Form, ByVal sectionID As AcSection) As Boolean
    On Error Resume Next
    Dim testHeight As Long
    testHeight = frm.Section(sectionID).Height
    sectionExist = (Err.Number = 0)
    On Error GoTo 0
End Function

' Purpose: Enables form header and footer sections.
Private Sub enableSection(ByVal frm As Access.Form)
    If Not sectionExist(frm, acHeader) Then
        frm.SetFocus
        DoCmd.RunCommand acCmdFormHdrFtr
    End If
End Sub

'====================================================================
' SECTION 3: RUNTIME ANCHOR & GEOMETRY MANAGEMENT
'====================================================================

'--------------------------------------------------------------------
' PROCEDURE : gui_anchorCapture
' PURPOSE   : Captures all control anchor properties into a Dictionary cache.
'--------------------------------------------------------------------
Public Function gui_anchorCapture(ByVal frm As Access.Form) As Object
    Dim cache As Object
    Set cache = CreateObject("Scripting.Dictionary")

    Dim ctrl As Access.Control
    For Each ctrl In frm.Controls
        Select Case ctrl.controlType
            Case acLabel, acTextBox, acComboBox, acListBox, _
                 acCommandButton, acToggleButton, acCheckBox, acOptionButton, acOptionGroup, _
                 acSubform, acImage, acRectangle, acLine, _
                 acTabCtl, acAttachment, acBoundObjectFrame, acObjectFrame
                cache(ctrl.name & "_H") = ctrl.HorizontalAnchor
                cache(ctrl.name & "_V") = ctrl.VerticalAnchor
        End Select
    Next ctrl

    Set gui_anchorCapture = cache
End Function

'--------------------------------------------------------------------
' PROCEDURE : gui_anchorFreezeTopLeft
' PURPOSE   : Temporarily freezes control anchors to Top-Left (0)
'             to allow safe VBA layout resizing without Access layout conflict.
'--------------------------------------------------------------------
Public Sub gui_anchorFreezeTopLeft(ByVal frm As Access.Form)
    Dim ctrl As Access.Control
    For Each ctrl In frm.Controls
        Select Case ctrl.controlType
            Case acSubform
                ' Skip subform container controls from runtime anchor manipulation

            Case acLabel, acTextBox, acComboBox, acListBox, _
                 acCommandButton, acToggleButton, acCheckBox, acOptionButton, acOptionGroup, _
                 acImage, acRectangle, acLine, _
                 acTabCtl, acAttachment, acBoundObjectFrame, acObjectFrame
                ctrl.HorizontalAnchor = 0
                ctrl.VerticalAnchor = 0
        End Select
    Next ctrl
End Sub

'--------------------------------------------------------------------
' PROCEDURE : gui_anchorRestore
' PURPOSE   : Restores control anchors from a previously captured Dictionary cache.
'--------------------------------------------------------------------
Public Sub gui_anchorRestore(ByVal frm As Access.Form, ByVal cache As Object)
    If cache Is Nothing Then Exit Sub

    Dim ctrl As Access.Control
    For Each ctrl In frm.Controls
        Select Case ctrl.controlType
            Case acSubform
                ' Skip subform container controls from runtime anchor restoration

            Case Else
                If cache.exists(ctrl.name & "_H") Then
                    ctrl.HorizontalAnchor = cache(ctrl.name & "_H")
                    ctrl.VerticalAnchor = cache(ctrl.name & "_V")
                End If
        End Select
    Next ctrl
End Sub
