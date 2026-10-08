Attribute VB_Name = "BuildGUI_Customers"
Option Compare Database
Option Explicit
'@Folder "App.Customers"

'====================================================================
' MODULE    : BuildGUI_Customers
' PURPOSE   : Form Wireframe and Styling Generator for Customers Domain.
' DEPENDENCIES: Extension_UIBuilder, Extension_UITheme, Extension_UITPrebuiltThemes
'====================================================================

'--------------------------------------------------------------------
' PROCEDURE : GenerateAll_Customers
' PURPOSE   : Batch master trigger to build & style all Customer forms.
'--------------------------------------------------------------------
Public Sub GenerateAll_Customers()
    On Error GoTo ErrHandler
    Application.Echo False

    Call Generate_frmCustomers
    Call Generate_frmCustomerGridView
    Call Generate_frmCustomerList

    Application.Echo True
    Debug.Print "[SUCCESS] All Customer forms generated and styled."
    Exit Sub

ErrHandler:
    Application.Echo True
    Debug.Print "[ERROR] BuildGUI_Customers.GenerateAll_Customers: #" & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "GenerateAll_Customers", Err.Description
End Sub

'====================================================================
' FORM 1: frmCustomers (Master Shell Router)
' SPEC  : Header Ribbon + Full-Width Workspace Viewport (10.98" x 6.00")
'====================================================================

Public Sub Generate_frmCustomers()
    ' PURPOSE: One-stop generator to build wireframe and apply style.
    On Error GoTo ErrHandler
    Application.Echo False

    Call BuildForm_frmCustomers
    Call StyleForm_frmCustomers

    Application.Echo True
    Exit Sub

ErrHandler:
    Application.Echo True
    Debug.Print "[ERROR] BuildGUI_Customers.Generate_frmCustomers: #" & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "Generate_frmCustomers", Err.Description
End Sub

'--------------------------------------------------------------------
' 1.1 WIREFRAME BUILDER: frmCustomers
'--------------------------------------------------------------------
Public Sub BuildForm_frmCustomers()
    On Error GoTo ErrHandler
    Const FORM_NAME As String = "frmCustomers"

    ' 1. SIZE FORM CANVAS (10.9792" x 6.0")
    Call gui_buildForm(FORM_NAME, gui_buildFormSizeDef(Width:=10.9792, Height:=6#))

    ' 2. HEADER RIBBON: Global Branding & System Controls (T=0 to 0.5")
    Call gui_buildControl(FORM_NAME, "boxHeaderNavy", acRectangle, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=10.9792, Height:=0.5, _
                                HorizontalAnchor:=acHorizontalAnchorBoth, VerticalAnchor:=acVerticalAnchorTop)) ' Right=10.98, Bottom=0.50

    Call gui_buildControl(FORM_NAME, "lblTitle", acLabel, _
        gui_buildControlSizeDef(Top:=0.1, Left:=0.2, Width:=2#, Height:=0.3, _
                                HorizontalAnchor:=acHorizontalAnchorLeft, VerticalAnchor:=acVerticalAnchorTop)) ' Right=2.20, Bottom=0.40

    Call gui_buildControl(FORM_NAME, "cmdSelectAll", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.1, Left:=2.3, Width:=0.35, Height:=0.3, _
                                HorizontalAnchor:=acHorizontalAnchorLeft, VerticalAnchor:=acVerticalAnchorTop)) ' Right=2.65, Bottom=0.40

    Call gui_buildControl(FORM_NAME, "txtSearch", acTextBox, _
        gui_buildControlSizeDef(Top:=0.1, Left:=2.7, Width:=2.1, Height:=0.3, _
                                HorizontalAnchor:=acHorizontalAnchorLeft, VerticalAnchor:=acVerticalAnchorTop)) ' Right=4.80, Bottom=0.40

    Call gui_buildControl(FORM_NAME, "btnClearFilter", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.1, Left:=4.9, Width:=0.4, Height:=0.3, _
                                HorizontalAnchor:=acHorizontalAnchorLeft, VerticalAnchor:=acVerticalAnchorTop)) ' Right=5.30, Bottom=0.40

    Call gui_buildControl(FORM_NAME, "cmdRefresh", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.1, Left:=6.4, Width:=0.4, Height:=0.3, _
                                HorizontalAnchor:=acHorizontalAnchorRight, VerticalAnchor:=acVerticalAnchorTop)) ' Right=6.80, Bottom=0.40

    Call gui_buildControl(FORM_NAME, "cmdSave", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.1, Left:=6.9, Width:=1#, Height:=0.3, _
                                HorizontalAnchor:=acHorizontalAnchorRight, VerticalAnchor:=acVerticalAnchorTop)) ' Right=7.90, Bottom=0.40

    Call gui_buildControl(FORM_NAME, "togTrashbin", acToggleButton, _
        gui_buildControlSizeDef(Top:=0.1, Left:=10.1, Width:=0.4, Height:=0.3, _
                                HorizontalAnchor:=acHorizontalAnchorRight, VerticalAnchor:=acVerticalAnchorTop)) ' Right=10.50, Bottom=0.40

    Call gui_buildControl(FORM_NAME, "cmdMaximize", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.1, Left:=10.56, Width:=0.36, Height:=0.3, _
                                HorizontalAnchor:=acHorizontalAnchorRight, VerticalAnchor:=acVerticalAnchorTop)) ' Right=10.92, Bottom=0.40

    ' 3. MAIN WORKSPACE VIEWPORT (T=0.50", W=10.98", H=5.50" Standard Canvas)
    Call gui_buildControl(FORM_NAME, "sfWorkspace", acSubform, _
        gui_buildControlSizeDef(Top:=0.5, Left:=0#, Width:=10.9792, Height:=5.5, _
                                HorizontalAnchor:=acHorizontalAnchorBoth, VerticalAnchor:=acVerticalAnchorBoth)) ' Right=10.98, Bottom=6.00

    ' [FINALIZE]: Save & Release
    DoCmd.Close acForm, FORM_NAME, acSaveYes
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] BuildGUI_Customers.BuildForm_frmCustomers: #" & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "BuildForm_frmCustomers", Err.Description
End Sub

'--------------------------------------------------------------------
' 1.2 AESTHETICS & STYLING: frmCustomers
'--------------------------------------------------------------------
Public Sub StyleForm_frmCustomers()
    On Error GoTo ErrHandler
    Const FORM_NAME As String = "frmCustomers"

    ' 1. OPEN IN DESIGN VIEW
    If Not CurrentProject.AllForms(FORM_NAME).IsLoaded Then
        DoCmd.OpenForm FORM_NAME, acDesign, , , , acHidden
    End If
    Dim frm As Access.Form: Set frm = Forms(FORM_NAME)

    ' 2. LEVEL 1 & 2 BASE STYLING (Theme Presets)
    Call gui_styleForm(FORM_NAME)
    Call gui_styleRectangle_BrandBar(FORM_NAME, "boxHeaderNavy")
    Call gui_styleLabel_Title(FORM_NAME, "lblTitle")
    Call gui_styleButton_Header(FORM_NAME, "cmdSelectAll")
    Call gui_styleTextBox_Magnifying(FORM_NAME, "txtSearch")
    Call gui_styleButton_Header(FORM_NAME, "btnClearFilter")
    Call gui_styleButton_Header(FORM_NAME, "cmdRefresh")
    Call gui_styleButton_Power(FORM_NAME, "cmdSave")
    Call gui_styleToggleButton_Sentinel(FORM_NAME, "togTrashbin")
    Call gui_styleButton_Header(FORM_NAME, "cmdMaximize")
    Call gui_styleSubform_Ghost(FORM_NAME, "sfWorkspace")

    ' 3. LEVEL 3 CONTEXTUAL OVERRIDES (Captions, Tooltips, Window Chrome)
    With frm
        ' Form Window Chrome
        .Caption = "Customers Master Hub"
        .NavigationCaption = "Customers"
        .Modal = False
        .MinMaxButtons = 3

        ' Header Captions & Placeholders
        .Controls("lblTitle").Caption = "CUSTOMER HUB"
        .Controls("txtSearch").Format = "@;"" Search customers..."""
        .Controls("cmdSelectAll").Caption = "ALL"
        .Controls("btnClearFilter").Caption = "CLER"
        .Controls("cmdRefresh").Caption = "REFR"
        .Controls("cmdSave").Caption = "[ SAVE ]"
        .Controls("togTrashbin").Caption = "TRSH"
        .Controls("cmdMaximize").Caption = "MAXM"

        ' Contextual Tooltips
        .Controls("txtSearch").ControlTipText = "Search by Customer Name, Contact or Phone (Press Enter)"
        .Controls("cmdSelectAll").ControlTipText = "Select or Deselect all customers"
        .Controls("btnClearFilter").ControlTipText = "Clear current search keyword"
        .Controls("cmdRefresh").ControlTipText = "Refresh dataset from database"
        .Controls("cmdSave").ControlTipText = "Commit and save pending changes"
        .Controls("togTrashbin").ControlTipText = "Switch between Active and Trashbin archive"
        .Controls("cmdMaximize").ControlTipText = "Maximize / Restore workspace view"
    End With

    ' 4. SAVE & CLOSE
    DoCmd.Close acForm, FORM_NAME, acSaveYes
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] BuildGUI_Customers.StyleForm_frmCustomers: #" & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "StyleForm_frmCustomers", Err.Description
End Sub

'====================================================================
' FORM 2: frmCustomerGridView (Grid Workspace Page)
' SPEC  : Filter Bar + Grid Viewport + Action Sidebar + Footer (10.98" x 5.50")
'====================================================================

Public Sub Generate_frmCustomerGridView()
    ' PURPOSE: One-stop generator to build wireframe and apply style.
    On Error GoTo ErrHandler
    Application.Echo False

    Call BuildForm_frmCustomerGridView
    Call StyleForm_frmCustomerGridView

    Application.Echo True
    Exit Sub

ErrHandler:
    Application.Echo True
    Debug.Print "[ERROR] BuildGUI_Customers.Generate_frmCustomerGridView: #" & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "Generate_frmCustomerGridView", Err.Description
End Sub

'--------------------------------------------------------------------
' 2.1 WIREFRAME BUILDER: frmCustomerGridView
'--------------------------------------------------------------------
Public Sub BuildForm_frmCustomerGridView()
    On Error GoTo ErrHandler
    Const FORM_NAME As String = "frmCustomerGridView"

    ' 1. SIZE FORM CANVAS (10.9792" x 5.50" - Standard Viewport Canvas)
    Call gui_buildForm(FORM_NAME, gui_buildFormSizeDef(Width:=10.9792, Height:=5.5))

    ' 2. FILTER BAR (T=0.10" - 0.40")
    Call gui_buildControl(FORM_NAME, "cboFilterCountry", acComboBox, _
        gui_buildControlSizeDef(Top:=0.1, Left:=0.2, Width:=2.2, Height:=0.3, _
                                HorizontalAnchor:=acHorizontalAnchorLeft, VerticalAnchor:=acVerticalAnchorTop)) ' Right=2.40, Bottom=0.40

    Call gui_buildControl(FORM_NAME, "cboFilterActive", acComboBox, _
        gui_buildControlSizeDef(Top:=0.1, Left:=2.5, Width:=1.5, Height:=0.3, _
                                HorizontalAnchor:=acHorizontalAnchorLeft, VerticalAnchor:=acVerticalAnchorTop)) ' Right=4.00, Bottom=0.40

    Call gui_buildControl(FORM_NAME, "linMainDivider", acLine, _
        gui_buildControlSizeDef(Top:=0.48, Left:=0.2, Width:=9.3, Height:=0.01, _
                                HorizontalAnchor:=acHorizontalAnchorBoth, VerticalAnchor:=acVerticalAnchorTop)) ' Right=9.50, Bottom=0.49

    ' 3. VIEWPORT AREA: Main Customer Data Grid Container (T=0.55", W=9.3", H=4.35")
    Call gui_buildControl(FORM_NAME, "sfCustomerGrid", acSubform, _
        gui_buildControlSizeDef(Top:=0.55, Left:=0.2, Width:=9.3, Height:=4.35, _
                                HorizontalAnchor:=acHorizontalAnchorBoth, VerticalAnchor:=acVerticalAnchorBoth)) ' Right=9.50, Bottom=4.90

    ' 4. SIDEBAR ACTIONS: Vertical I/O Stack (L=9.60", W=1.18")
    Call gui_buildControl(FORM_NAME, "cmdNewCustomer", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.1, Left:=9.6, Width:=1.18, Height:=0.4, _
                                HorizontalAnchor:=acHorizontalAnchorRight, VerticalAnchor:=acVerticalAnchorTop)) ' Right=10.78, Bottom=0.50

    Call gui_buildControl(FORM_NAME, "cmdExport", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.6, Left:=9.6, Width:=1.18, Height:=0.3, _
                                HorizontalAnchor:=acHorizontalAnchorRight, VerticalAnchor:=acVerticalAnchorTop)) ' Right=10.78, Bottom=0.90

    Call gui_buildControl(FORM_NAME, "cmdImport", acCommandButton, _
        gui_buildControlSizeDef(Top:=1#, Left:=9.6, Width:=1.18, Height:=0.3, _
                                HorizontalAnchor:=acHorizontalAnchorRight, VerticalAnchor:=acVerticalAnchorTop)) ' Right=10.78, Bottom=1.30

    ' 5. FOOTER STRIP: Summary & Metrics (T=5.00")
    Call gui_buildControl(FORM_NAME, "linFooterSeparator", acLine, _
        gui_buildControlSizeDef(Top:=5#, Left:=0.2, Width:=9.3, Height:=0.01, _
                                HorizontalAnchor:=acHorizontalAnchorBoth, VerticalAnchor:=acVerticalAnchorBottom)) ' Right=9.50, Bottom=5.01

    Call gui_buildControl(FORM_NAME, "lblRecordCount", acLabel, _
        gui_buildControlSizeDef(Top:=5.1, Left:=0.2, Width:=3.5, Height:=0.28, _
                                HorizontalAnchor:=acHorizontalAnchorLeft, VerticalAnchor:=acVerticalAnchorBottom)) ' Right=3.70, Bottom=5.38

    ' [FINALIZE]: Save & Release
    DoCmd.Close acForm, FORM_NAME, acSaveYes
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] BuildGUI_Customers.BuildForm_frmCustomerGridView: #" & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "BuildForm_frmCustomerGridView", Err.Description
End Sub

'--------------------------------------------------------------------
' 2.2 AESTHETICS & STYLING: frmCustomerGridView
'--------------------------------------------------------------------
Public Sub StyleForm_frmCustomerGridView()
    On Error GoTo ErrHandler
    Const FORM_NAME As String = "frmCustomerGridView"

    ' 1. OPEN IN DESIGN VIEW
    If Not CurrentProject.AllForms(FORM_NAME).IsLoaded Then
        DoCmd.OpenForm FORM_NAME, acDesign, , , , acHidden
    End If
    Dim frm As Access.Form: Set frm = Forms(FORM_NAME)

    ' 2. BASE STYLING (Theme Presets)
    Call gui_styleForm(FORM_NAME)
    Call gui_styleComboBox_Standard(FORM_NAME, "cboFilterCountry")
    Call gui_styleComboBox_Standard(FORM_NAME, "cboFilterActive")
    Call gui_styleLine_Divider(FORM_NAME, "linMainDivider")
    Call gui_styleSubform_Theater(FORM_NAME, "sfCustomerGrid")
    Call gui_styleButton_Power(FORM_NAME, "cmdNewCustomer")
    Call gui_styleButton_Success(FORM_NAME, "cmdExport")
    Call gui_styleButton_Secondary(FORM_NAME, "cmdImport")
    Call gui_styleLine_Divider(FORM_NAME, "linFooterSeparator")
    Call gui_styleLabel_Metadata(FORM_NAME, "lblRecordCount")

    ' 3. CONTEXTUAL OVERRIDES
    With frm
        .Controls("cboFilterCountry").Format = "@;""[ All Countries ]"""
        .Controls("cboFilterActive").Format = "@;""[ All Status ]"""
        .Controls("linMainDivider").BorderColor = HexToRBGLong("#C7CCEA")
        .Controls("linFooterSeparator").BorderColor = HexToRBGLong("#C7CCEA")

        .Controls("cmdNewCustomer").Caption = "+ NEW CUSTOMER"
        .Controls("cmdExport").Caption = "EXPORT EXCEL"
        .Controls("cmdImport").Caption = "IMPORT"
        .Controls("lblRecordCount").Caption = "Showing: 0 / 0 customers"

        .Controls("cmdNewCustomer").ControlTipText = "Add a new customer entry"
        .Controls("cmdExport").ControlTipText = "Export customer dataset to Excel"
        .Controls("cmdImport").ControlTipText = "Import customer records from external file"
    End With

    ' 4. SAVE & CLOSE
    DoCmd.Close acForm, FORM_NAME, acSaveYes
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] BuildGUI_Customers.StyleForm_frmCustomerGridView: #" & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "StyleForm_frmCustomerGridView", Err.Description
End Sub

'====================================================================
' FORM 3: frmCustomerList (Datasheet Grid Viewport)
' SPEC  : Core Embedded Customer Dataset Table
'====================================================================

Public Sub Generate_frmCustomerList()
    ' PURPOSE: One-stop generator to build wireframe and apply style.
    On Error GoTo ErrHandler
    Application.Echo False

    Call BuildForm_frmCustomerList
    Call StyleForm_frmCustomerList

    Application.Echo True
    Exit Sub

ErrHandler:
    Application.Echo True
    Debug.Print "[ERROR] BuildGUI_Customers.Generate_frmCustomerList: #" & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "Generate_frmCustomerList", Err.Description
End Sub

'--------------------------------------------------------------------
' 3.1 WIREFRAME BUILDER: frmCustomerList
'--------------------------------------------------------------------
Public Sub BuildForm_frmCustomerList()
    On Error GoTo ErrHandler
    Const FORM_NAME As String = "frmCustomerList"

    ' 1. SIZE FORM CANVAS (10.9792" x 6.0")
    Call gui_buildForm(FORM_NAME, gui_buildFormSizeDef(Width:=10.9792, Height:=6#))

    ' 2. GRID CONTROLS (T=0, H=0.25")
    Call gui_buildControl(FORM_NAME, "chkContextAction", acCheckBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=0.4, Height:=0.25)) ' Right=0.40, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtRowNum", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=0.4, Height:=0.25)) ' Right=0.40, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtCompanyName", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=2.5, Height:=0.25)) ' Right=2.50, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtCountry", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=1.2, Height:=0.25)) ' Right=1.20, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtContinent", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=1.2, Height:=0.25)) ' Right=1.20, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtContactPerson", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=1.8, Height:=0.25)) ' Right=1.80, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtContactPersonPhone", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=1.3, Height:=0.25)) ' Right=1.30, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtContactPersonEmail", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=1.8, Height:=0.25)) ' Right=1.80, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtCompanyPhone", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=1.3, Height:=0.25)) ' Right=1.30, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtCompanyEmail", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=1.8, Height:=0.25)) ' Right=1.80, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtCompanyWebsite", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=1.8, Height:=0.25)) ' Right=1.80, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtCompanyAddress1", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=2.5, Height:=0.25)) ' Right=2.50, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtCompanyAddress2", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=2.5, Height:=0.25)) ' Right=2.50, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtTaxID", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=1.2, Height:=0.25)) ' Right=1.20, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtNotes", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=2#, Height:=0.25)) ' Right=2.00, Bottom=0.25

    ' [FINALIZE]: Save & Release
    DoCmd.Close acForm, FORM_NAME, acSaveYes
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] BuildGUI_Customers.BuildForm_frmCustomerList: #" & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "BuildForm_frmCustomerList", Err.Description
End Sub

'--------------------------------------------------------------------
' 3.2 AESTHETICS & STYLING: frmCustomerList
'--------------------------------------------------------------------
Public Sub StyleForm_frmCustomerList()
    On Error GoTo ErrHandler
    Const FORM_NAME As String = "frmCustomerList"

    ' 1. OPEN IN DESIGN VIEW
    If Not CurrentProject.AllForms(FORM_NAME).IsLoaded Then
        DoCmd.OpenForm FORM_NAME, acDesign, , , , acHidden
    End If
    Dim frm As Access.Form: Set frm = Forms(FORM_NAME)

    ' 2. FORM SHELL (Datasheet Style)
    Call gui_styleForm(FORM_NAME)
    With frm
        .DefaultView = DefaultView_Datasheet
        .AllowDatasheetView = True
        .Section(acDetail).BackColor = vbWhite
    End With

    ' 3. ATOMIC CONTROLS
    Call gui_styleCheckBox_Standard(FORM_NAME, "chkContextAction")
    Call gui_styleTextBox_Value(FORM_NAME, "txtRowNum")
    Call gui_styleTextBox_HubTitle(FORM_NAME, "txtCompanyName")
    Call gui_styleTextBox_Machine(FORM_NAME, "txtCountry")
    Call gui_styleTextBox_Machine(FORM_NAME, "txtContinent")
    Call gui_styleTextBox_Document(FORM_NAME, "txtContactPerson")
    Call gui_styleTextBox_Machine(FORM_NAME, "txtContactPersonPhone")
    Call gui_styleTextBox_Document(FORM_NAME, "txtContactPersonEmail")
    Call gui_styleTextBox_Machine(FORM_NAME, "txtCompanyPhone")
    Call gui_styleTextBox_Document(FORM_NAME, "txtCompanyEmail")
    Call gui_styleTextBox_Document(FORM_NAME, "txtCompanyWebsite")
    Call gui_styleTextBox_Document(FORM_NAME, "txtCompanyAddress1")
    Call gui_styleTextBox_Document(FORM_NAME, "txtCompanyAddress2")
    Call gui_styleTextBox_Machine(FORM_NAME, "txtTaxID")
    Call gui_styleTextBox_Document(FORM_NAME, "txtNotes")

    ' 4. CONTEXTUAL OVERRIDES (Datasheet Captions & Formats)
    With frm
        .Controls("chkContextAction").Properties("DatasheetCaption") = "Selected"
        .Controls("txtRowNum").Properties("DatasheetCaption") = "No."
        .Controls("txtCompanyName").Properties("DatasheetCaption") = "Company Name"
        .Controls("txtCountry").Properties("DatasheetCaption") = "Country"
        .Controls("txtContinent").Properties("DatasheetCaption") = "Continent"
        .Controls("txtContactPerson").Properties("DatasheetCaption") = "Contact Person"
        .Controls("txtContactPersonPhone").Properties("DatasheetCaption") = "Contact Phone"
        .Controls("txtContactPersonEmail").Properties("DatasheetCaption") = "Contact Email"
        .Controls("txtCompanyPhone").Properties("DatasheetCaption") = "Office Phone"
        .Controls("txtCompanyEmail").Properties("DatasheetCaption") = "Office Email"
        .Controls("txtCompanyWebsite").Properties("DatasheetCaption") = "Website"
        .Controls("txtCompanyAddress1").Properties("DatasheetCaption") = "Address Line 1"
        .Controls("txtCompanyAddress2").Properties("DatasheetCaption") = "Address Line 2"
        .Controls("txtTaxID").Properties("DatasheetCaption") = "Tax ID"
        .Controls("txtNotes").Properties("DatasheetCaption") = "Notes"

        ' Formats
        .Controls("txtTaxID").Format = "@"
        .Controls("txtContactPersonPhone").Format = "@"
        .Controls("txtCompanyPhone").Format = "@"
    End With

    ' 5. SAVE & CLOSE
    DoCmd.Close acForm, FORM_NAME, acSaveYes
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] BuildGUI_Customers.StyleForm_frmCustomerList: #" & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "StyleForm_frmCustomerList", Err.Description
End Sub
