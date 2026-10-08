Attribute VB_Name = "BuildGUI_Products"
Option Compare Database
Option Explicit
'@Folder "App.Products"

'====================================================================
' MODULE    : BuildGUI_Products
' PURPOSE   : Form Wireframe and Styling Generator for Products Domain.
' DEPENDENCIES: Extension_UIBuilder, Extension_UITheme, Extension_UITPrebuiltThemes
'====================================================================

'--------------------------------------------------------------------
' PROCEDURE : GenerateAll_Products
' PURPOSE   : Batch master trigger to build & style all Product forms.
'--------------------------------------------------------------------
Public Sub GenerateAll_Products()
    On Error GoTo ErrHandler
    Application.Echo False

    Call Generate_frmProducts
    Call Generate_frmProductGridView
    Call Generate_frmProductDetailView
    Call Generate_frmProductList

    Application.Echo True
    Debug.Print "[SUCCESS] All Product forms generated and styled."
    Exit Sub

ErrHandler:
    Application.Echo True
    Err.Raise Err.Number, "GenerateAll_Products", Err.Description
End Sub

'====================================================================
' FORM 1: frmProducts (Master Shell Router)
' SPEC  : Header Ribbon + Single Full-Width Workspace Viewport (10.98" x 6.00")
'====================================================================

Public Sub Generate_frmProducts()
    ' PURPOSE: One-stop generator to build wireframe and apply style.
    On Error GoTo ErrHandler
    Application.Echo False

    Call BuildForm_frmProducts
    Call StyleForm_frmProducts

    Application.Echo True
    Exit Sub

ErrHandler:
    Application.Echo True
    Err.Raise Err.Number, "Generate_frmProducts", Err.Description
End Sub

'--------------------------------------------------------------------
' 1.1 WIREFRAME BUILDER: frmProducts
'--------------------------------------------------------------------
Public Sub BuildForm_frmProducts()
    On Error GoTo ErrHandler
    Const FORM_NAME As String = "frmProducts"

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

    Call gui_buildControl(FORM_NAME, "cmdViewImage", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.1, Left:=8.6, Width:=0.4, Height:=0.3, _
                                HorizontalAnchor:=acHorizontalAnchorRight, VerticalAnchor:=acVerticalAnchorTop)) ' Right=9.00, Bottom=0.40

    Call gui_buildControl(FORM_NAME, "togTrashbin", acToggleButton, _
        gui_buildControlSizeDef(Top:=0.1, Left:=9.05, Width:=0.4, Height:=0.3, _
                                HorizontalAnchor:=acHorizontalAnchorRight, VerticalAnchor:=acVerticalAnchorTop)) ' Right=9.45, Bottom=0.40

    Call gui_buildControl(FORM_NAME, "cmdLanguage", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.1, Left:=9.5, Width:=0.36, Height:=0.3, _
                                HorizontalAnchor:=acHorizontalAnchorRight, VerticalAnchor:=acVerticalAnchorTop)) ' Right=9.86, Bottom=0.40

    Call gui_buildControl(FORM_NAME, "cmdMaximize", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.1, Left:=10.15, Width:=0.36, Height:=0.3, _
                                HorizontalAnchor:=acHorizontalAnchorRight, VerticalAnchor:=acVerticalAnchorTop)) ' Right=10.51, Bottom=0.40

    Call gui_buildControl(FORM_NAME, "cmdToggleView", acCommandButton, _
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
    Err.Raise Err.Number, "BuildForm_frmProducts", Err.Description
End Sub

'--------------------------------------------------------------------
' 1.2 AESTHETICS & STYLING: frmProducts
'--------------------------------------------------------------------
Public Sub StyleForm_frmProducts()
    On Error GoTo ErrHandler
    Const FORM_NAME As String = "frmProducts"

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
    Call gui_styleButton_Header(FORM_NAME, "cmdViewImage")
    Call gui_styleToggleButton_Sentinel(FORM_NAME, "togTrashbin")
    Call gui_styleButton_Header(FORM_NAME, "cmdLanguage")
    Call gui_styleButton_Header(FORM_NAME, "cmdMaximize")
    Call gui_styleButton_Header(FORM_NAME, "cmdToggleView")
    Call gui_styleSubform_Ghost(FORM_NAME, "sfWorkspace")

    ' 3. LEVEL 3 CONTEXTUAL OVERRIDES (Captions, Tooltips, Window Chrome)
    With frm
        ' Form Window Chrome
        .Caption = "Products Master Hub"
        .NavigationCaption = "Products"
        .Modal = False
        .MinMaxButtons = 3

        ' Header Captions & Placeholders
        .Controls("lblTitle").Caption = "PRODUCT HUB"
        .Controls("txtSearch").Format = "@;"" Search products or barcode..."""
        .Controls("cmdSelectAll").Caption = "ALL"
        .Controls("btnClearFilter").Caption = "CLER"
        .Controls("cmdRefresh").Caption = "REFR"
        .Controls("cmdSave").Caption = "[ SAVE ]"
        .Controls("cmdViewImage").Caption = "IMGE"
        .Controls("togTrashbin").Caption = "TRSH"
        .Controls("cmdLanguage").Caption = "LANG"
        .Controls("cmdMaximize").Caption = "MAXM"
        .Controls("cmdToggleView").Caption = "VIEW"

        ' Contextual Tooltips
        .Controls("txtSearch").ControlTipText = "Search by Product Name or Barcode (Press Enter)"
        .Controls("cmdSelectAll").ControlTipText = "Select or Deselect all items"
        .Controls("btnClearFilter").ControlTipText = "Clear current search keyword"
        .Controls("cmdRefresh").ControlTipText = "Refresh dataset from database"
        .Controls("cmdSave").ControlTipText = "Commit and save pending changes"
        .Controls("cmdViewImage").ControlTipText = "Toggle product image inspection panel"
        .Controls("togTrashbin").ControlTipText = "Switch between Active and Trashbin archive"
        .Controls("cmdLanguage").ControlTipText = "Switch language between English and Vietnamese"
        .Controls("cmdMaximize").ControlTipText = "Maximize / Restore workspace view"
        .Controls("cmdToggleView").ControlTipText = "Switch between Grid view and Detail view"
    End With

    ' 4. SAVE & CLOSE
    DoCmd.Close acForm, FORM_NAME, acSaveYes
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "StyleForm_frmProducts", Err.Description
End Sub

'====================================================================
' FORM 2: frmProductGridView (Grid Workspace Page)
' SPEC  : Filter Bar + Grid Viewport + Action Sidebar + Media + Footer (10.98" x 5.50")
'====================================================================

Public Sub Generate_frmProductGridView()
    ' PURPOSE: One-stop generator to build wireframe and apply style.
    On Error GoTo ErrHandler
    Application.Echo False

    Call BuildForm_frmProductGridView
    Call StyleForm_frmProductGridView

    Application.Echo True
    Exit Sub

ErrHandler:
    Application.Echo True
    Err.Raise Err.Number, "Generate_frmProductGridView", Err.Description
End Sub

'--------------------------------------------------------------------
' 2.1 WIREFRAME BUILDER: frmProductGridView
'--------------------------------------------------------------------
Public Sub BuildForm_frmProductGridView()
    On Error GoTo ErrHandler
    Const FORM_NAME As String = "frmProductGridView"

    ' 1. SIZE FORM CANVAS (10.9792" x 5.50" - Standard Viewport Canvas)
    Call gui_buildForm(FORM_NAME, gui_buildFormSizeDef(Width:=10.9792, Height:=5.5))

    ' 2. FILTER BAR (T=0.10" - 0.40")
    Call gui_buildControl(FORM_NAME, "cboFilterCategory", acComboBox, _
        gui_buildControlSizeDef(Top:=0.1, Left:=0.2, Width:=1.8, Height:=0.3, _
                                HorizontalAnchor:=acHorizontalAnchorLeft, VerticalAnchor:=acVerticalAnchorTop)) ' Right=2.00, Bottom=0.40

    Call gui_buildControl(FORM_NAME, "cboFilterLine", acComboBox, _
        gui_buildControlSizeDef(Top:=0.1, Left:=2.1, Width:=1.8, Height:=0.3, _
                                HorizontalAnchor:=acHorizontalAnchorLeft, VerticalAnchor:=acVerticalAnchorTop)) ' Right=3.90, Bottom=0.40

    Call gui_buildControl(FORM_NAME, "linMainDivider", acLine, _
        gui_buildControlSizeDef(Top:=0.48, Left:=0.2, Width:=9.3, Height:=0.01, _
                                HorizontalAnchor:=acHorizontalAnchorBoth, VerticalAnchor:=acVerticalAnchorTop)) ' Right=9.50, Bottom=0.49

    ' 3. CENTER DATA VIEWPORT (Houses frmProductList) (T=0.55", W=9.30", H=4.35")
    Call gui_buildControl(FORM_NAME, "sfProductGrid", acSubform, _
        gui_buildControlSizeDef(Top:=0.55, Left:=0.2, Width:=9.3, Height:=4.35, _
                                HorizontalAnchor:=acHorizontalAnchorBoth, VerticalAnchor:=acVerticalAnchorBoth)) ' Right=9.50, Bottom=4.90

    ' 4. SIDEBAR ACTIONS: Vertical I/O Stack (L=9.60", W=1.18")
    Call gui_buildControl(FORM_NAME, "cmdNewProduct", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.1, Left:=9.6, Width:=1.18, Height:=0.4, _
                                HorizontalAnchor:=acHorizontalAnchorRight, VerticalAnchor:=acVerticalAnchorTop)) ' Right=10.78, Bottom=0.50

    Call gui_buildControl(FORM_NAME, "cmdExport", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.6, Left:=9.6, Width:=1.18, Height:=0.3, _
                                HorizontalAnchor:=acHorizontalAnchorRight, VerticalAnchor:=acVerticalAnchorTop)) ' Right=10.78, Bottom=0.90

    ' 5. HIDDEN MEDIA OVERLAY (Split-View Image Inspection)
    Call gui_buildControl(FORM_NAME, "imgProduct", acImage, _
        gui_buildControlSizeDef(Top:=0.55, Left:=5.3, Width:=4.2, Height:=3.5, _
                                HorizontalAnchor:=acHorizontalAnchorRight, VerticalAnchor:=acVerticalAnchorBoth)) ' Right=9.50, Bottom=4.05

    Call gui_buildControl(FORM_NAME, "LabelNoImage", acLabel, _
        gui_buildControlSizeDef(Top:=1.55, Left:=5.3, Width:=4.2, Height:=0.67, _
                                HorizontalAnchor:=acHorizontalAnchorRight, VerticalAnchor:=acVerticalAnchorBoth)) ' Right=9.50, Bottom=2.22

    Call gui_buildControl(FORM_NAME, "cmdAddImage", acCommandButton, _
        gui_buildControlSizeDef(Top:=4.15, Left:=5.3, Width:=1.95, Height:=0.3, _
                                HorizontalAnchor:=acHorizontalAnchorRight, VerticalAnchor:=acVerticalAnchorBottom)) ' Right=7.25, Bottom=4.45

    Call gui_buildControl(FORM_NAME, "cmdClearImage", acCommandButton, _
        gui_buildControlSizeDef(Top:=4.15, Left:=7.45, Width:=2.05, Height:=0.3, _
                                HorizontalAnchor:=acHorizontalAnchorRight, VerticalAnchor:=acVerticalAnchorBottom)) ' Right=9.50, Bottom=4.45

    ' 6. FOOTER SUMMARY STRIP (T=5.00" - 5.40")
    Call gui_buildControl(FORM_NAME, "linFooterSeparator", acLine, _
        gui_buildControlSizeDef(Top:=5#, Left:=0.2, Width:=9.3, Height:=0.01, _
                                HorizontalAnchor:=acHorizontalAnchorBoth, VerticalAnchor:=acVerticalAnchorBottom)) ' Right=9.50, Bottom=5.01

    Call gui_buildControl(FORM_NAME, "lblRecordCount", acLabel, _
        gui_buildControlSizeDef(Top:=5.1, Left:=0.2, Width:=2.5, Height:=0.25, _
                                HorizontalAnchor:=acHorizontalAnchorLeft, VerticalAnchor:=acVerticalAnchorBottom)) ' Right=2.70, Bottom=5.35

    ' [FINALIZE]: Save & Release
    DoCmd.Close acForm, FORM_NAME, acSaveYes
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "BuildForm_frmProductGridView", Err.Description
End Sub

'--------------------------------------------------------------------
' 2.2 AESTHETICS & STYLING: frmProductGridView
'--------------------------------------------------------------------
Public Sub StyleForm_frmProductGridView()
    On Error GoTo ErrHandler
    Const FORM_NAME As String = "frmProductGridView"

    ' 1. OPEN IN DESIGN VIEW
    If Not CurrentProject.AllForms(FORM_NAME).IsLoaded Then
        DoCmd.OpenForm FORM_NAME, acDesign, , , , acHidden
    End If
    Dim frm As Access.Form: Set frm = Forms(FORM_NAME)

    ' 2. LEVEL 1 & 2 BASE STYLING (Theme Presets)
    Call gui_styleForm(FORM_NAME)
    Call gui_styleComboBox_Standard(FORM_NAME, "cboFilterCategory")
    Call gui_styleComboBox_Standard(FORM_NAME, "cboFilterLine")
    Call gui_styleLine_Divider(FORM_NAME, "linMainDivider")
    Call gui_styleSubform_Theater(FORM_NAME, "sfProductGrid")

    ' Action Buttons
    Call gui_styleButton_Power(FORM_NAME, "cmdNewProduct")
    Call gui_styleButton_Success(FORM_NAME, "cmdExport")

    ' Media Overlay Controls
    Call gui_styleImage_Showcase(FORM_NAME, "imgProduct")
    Call gui_styleLabel_Watermark(FORM_NAME, "LabelNoImage")
    Call gui_styleButton_Ghost(FORM_NAME, "cmdAddImage")
    Call gui_styleButton_Danger(FORM_NAME, "cmdClearImage")

    ' Footer Strip
    Call gui_styleLine_Divider(FORM_NAME, "linFooterSeparator")
    Call gui_styleLabel_Metadata(FORM_NAME, "lblRecordCount")

    ' 3. LEVEL 3 CONTEXTUAL OVERRIDES (Captions, Tooltips, Default States)
    With frm
        ' Form Chrome
        .Caption = "Product Grid Viewport"
        .Modal = False
        .MinMaxButtons = 0

        ' Connect Subform Grid Source
        .Controls("sfProductGrid").SourceObject = "frmProductList"

        ' Filter ComboBox Placeholders
        .Controls("cboFilterCategory").Format = "@;""[ All Categories ]"""
        .Controls("cboFilterLine").Format = "@;""[ All Marketing Lines ]"""

        ' Action Button Captions
        .Controls("cmdNewProduct").Caption = "+ NEW PRODUCT"
        .Controls("cmdExport").Caption = "EXPORT EXCEL"

        ' Media Captions & Default Visibility
        .Controls("LabelNoImage").Caption = "NO IMAGE"
        .Controls("cmdAddImage").Caption = "ADD IMAGE"
        .Controls("cmdClearImage").Caption = "CLEAR"
        .Controls("imgProduct").Visible = True
        .Controls("LabelNoImage").Visible = False
        .Controls("cmdAddImage").Visible = False
        .Controls("cmdClearImage").Visible = False

        ' Footer Labels
        .Controls("lblRecordCount").Caption = "Showing: 0 / 0 products"

        ' Contextual Tooltips
        .Controls("cboFilterCategory").ControlTipText = "Filter by Product Category"
        .Controls("cboFilterLine").ControlTipText = "Filter by Marketing Line"
        .Controls("cmdNewProduct").ControlTipText = "Create a new product record"
        .Controls("cmdExport").ControlTipText = "Export active products to Excel"
        .Controls("cmdAddImage").ControlTipText = "Select and attach image for active product"
        .Controls("cmdClearImage").ControlTipText = "Remove attached image from active product"
    End With

    ' 4. SAVE & CLOSE
    DoCmd.Close acForm, FORM_NAME, acSaveYes
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "StyleForm_frmProductGridView", Err.Description
End Sub

'====================================================================
' FORM 3: frmProductDetailView (Detail Viewport Page)
' SPEC  : 3-Column Detail + Media + Actions + Active Nav (10.98" x 5.50")
'====================================================================

Public Sub Generate_frmProductDetailView()
    ' PURPOSE: One-stop generator to build wireframe and apply style.
    On Error GoTo ErrHandler
    Application.Echo False

    Call BuildForm_frmProductDetailView
    Call StyleForm_frmProductDetailView

    Application.Echo True
    Exit Sub

ErrHandler:
    Application.Echo True
    Err.Raise Err.Number, "Generate_frmProductDetailView", Err.Description
End Sub

'--------------------------------------------------------------------
' 3.1 WIREFRAME BUILDER: frmProductDetailView
'--------------------------------------------------------------------
Public Sub BuildForm_frmProductDetailView()
    On Error GoTo ErrHandler
    Const FORM_NAME As String = "frmProductDetailView"

    ' 1. SIZE FORM CANVAS (10.9792" x 5.50" - Standard Viewport Canvas)
    Call gui_buildForm(FORM_NAME, gui_buildFormSizeDef(Width:=10.9792, Height:=5.5))

    ' 2. COLUMN 1: GENERAL & PACKAGING (Left Card)
    Call gui_buildControl(FORM_NAME, "boxGeneral", acRectangle, _
        gui_buildControlSizeDef(Top:=0.1, Left:=0.2, Width:=2.65, Height:=3#)) ' Right=2.85, Bottom=3.10

    Call gui_buildControl(FORM_NAME, "lblGeneralTitle", acLabel, _
        gui_buildControlSizeDef(Top:=0.15, Left:=0.25, Width:=2.55, Height:=0.22)) ' Right=2.80, Bottom=0.37

    Call gui_buildControl(FORM_NAME, "lblMarketingCategory", acLabel, _
        gui_buildControlSizeDef(Top:=0.45, Left:=0.25, Width:=0.9, Height:=0.22)) ' Right=1.15, Bottom=0.67

    Call gui_buildControl(FORM_NAME, "comboMarketingCategory", acComboBox, _
        gui_buildControlSizeDef(Top:=0.45, Left:=1.2, Width:=1.6, Height:=0.22)) ' Right=2.80, Bottom=0.67

    Call gui_buildControl(FORM_NAME, "lblMarketingLine", acLabel, _
        gui_buildControlSizeDef(Top:=0.72, Left:=0.25, Width:=0.9, Height:=0.22)) ' Right=1.15, Bottom=0.94

    Call gui_buildControl(FORM_NAME, "comboMarketingLine", acComboBox, _
        gui_buildControlSizeDef(Top:=0.72, Left:=1.2, Width:=1.6, Height:=0.22)) ' Right=2.80, Bottom=0.94

    Call gui_buildControl(FORM_NAME, "lblBarcode", acLabel, _
        gui_buildControlSizeDef(Top:=0.99, Left:=0.25, Width:=0.9, Height:=0.22)) ' Right=1.15, Bottom=1.21

    Call gui_buildControl(FORM_NAME, "txtBarcode", acTextBox, _
        gui_buildControlSizeDef(Top:=0.99, Left:=1.2, Width:=1.6, Height:=0.22)) ' Right=2.80, Bottom=1.21

    Call gui_buildControl(FORM_NAME, "lblHSCode", acLabel, _
        gui_buildControlSizeDef(Top:=1.26, Left:=0.25, Width:=0.9, Height:=0.22)) ' Right=1.15, Bottom=1.48

    Call gui_buildControl(FORM_NAME, "txtHSCode", acTextBox, _
        gui_buildControlSizeDef(Top:=1.26, Left:=1.2, Width:=1.6, Height:=0.22)) ' Right=2.80, Bottom=1.48

    Call gui_buildControl(FORM_NAME, "lblProductName_en", acLabel, _
        gui_buildControlSizeDef(Top:=1.53, Left:=0.25, Width:=0.9, Height:=0.22)) ' Right=1.15, Bottom=1.75

    Call gui_buildControl(FORM_NAME, "txtProductName_en", acTextBox, _
        gui_buildControlSizeDef(Top:=1.53, Left:=1.2, Width:=1.6, Height:=0.22)) ' Right=2.80, Bottom=1.75

    Call gui_buildControl(FORM_NAME, "lblProductName_vn", acLabel, _
        gui_buildControlSizeDef(Top:=1.8, Left:=0.25, Width:=0.9, Height:=0.22)) ' Right=1.15, Bottom=2.02

    Call gui_buildControl(FORM_NAME, "txtProductName_vn", acTextBox, _
        gui_buildControlSizeDef(Top:=1.8, Left:=1.2, Width:=1.6, Height:=0.22)) ' Right=2.80, Bottom=2.02

    Call gui_buildControl(FORM_NAME, "lblPackageType", acLabel, _
        gui_buildControlSizeDef(Top:=2.07, Left:=0.25, Width:=0.9, Height:=0.22)) ' Right=1.15, Bottom=2.29

    Call gui_buildControl(FORM_NAME, "cboPackageType", acComboBox, _
        gui_buildControlSizeDef(Top:=2.07, Left:=1.2, Width:=1.6, Height:=0.22)) ' Right=2.80, Bottom=2.29

    Call gui_buildControl(FORM_NAME, "lblShelfLife_month", acLabel, _
        gui_buildControlSizeDef(Top:=2.34, Left:=0.25, Width:=0.9, Height:=0.22)) ' Right=1.15, Bottom=2.56

    Call gui_buildControl(FORM_NAME, "txtShelfLife_month", acTextBox, _
        gui_buildControlSizeDef(Top:=2.34, Left:=1.2, Width:=1.6, Height:=0.22)) ' Right=2.80, Bottom=2.56

    Call gui_buildControl(FORM_NAME, "lblOuterPaperBox", acLabel, _
        gui_buildControlSizeDef(Top:=2.61, Left:=0.25, Width:=0.9, Height:=0.22)) ' Right=1.15, Bottom=2.83

    Call gui_buildControl(FORM_NAME, "blnOuterPaperBox", acCheckBox, _
        gui_buildControlSizeDef(Top:=2.61, Left:=1.2, Width:=0.3, Height:=0.22)) ' Right=1.50, Bottom=2.83

    ' 3. COLUMN 2: PRICING & LOGISTICS (Middle Card)
    Call gui_buildControl(FORM_NAME, "boxDimensions", acRectangle, _
        gui_buildControlSizeDef(Top:=0.1, Left:=2.95, Width:=2.65, Height:=3#)) ' Right=5.60, Bottom=3.10

    Call gui_buildControl(FORM_NAME, "lblDimensionsTitle", acLabel, _
        gui_buildControlSizeDef(Top:=0.15, Left:=3#, Width:=2.55, Height:=0.22)) ' Right=5.55, Bottom=0.37

    Call gui_buildControl(FORM_NAME, "lblCapacity", acLabel, _
        gui_buildControlSizeDef(Top:=0.45, Left:=3#, Width:=0.9, Height:=0.22)) ' Right=3.90, Bottom=0.67

    Call gui_buildControl(FORM_NAME, "txtCapacity", acTextBox, _
        gui_buildControlSizeDef(Top:=0.45, Left:=3.95, Width:=0.95, Height:=0.22)) ' Right=4.90, Bottom=0.67

    Call gui_buildControl(FORM_NAME, "cboCapacityUnit", acComboBox, _
        gui_buildControlSizeDef(Top:=0.45, Left:=4.95, Width:=0.6, Height:=0.22)) ' Right=5.55, Bottom=0.67

    Call gui_buildControl(FORM_NAME, "lblCapacity_oz", acLabel, _
        gui_buildControlSizeDef(Top:=0.72, Left:=3#, Width:=0.9, Height:=0.22)) ' Right=3.90, Bottom=0.94

    Call gui_buildControl(FORM_NAME, "txtCapacity_oz", acTextBox, _
        gui_buildControlSizeDef(Top:=0.72, Left:=3.95, Width:=1.6, Height:=0.22)) ' Right=5.55, Bottom=0.94

    Call gui_buildControl(FORM_NAME, "lblPrice_vnd", acLabel, _
        gui_buildControlSizeDef(Top:=0.99, Left:=3#, Width:=0.9, Height:=0.22)) ' Right=3.90, Bottom=1.21

    Call gui_buildControl(FORM_NAME, "txtPrice_vnd", acTextBox, _
        gui_buildControlSizeDef(Top:=0.99, Left:=3.95, Width:=1.6, Height:=0.22)) ' Right=5.55, Bottom=1.21

    Call gui_buildControl(FORM_NAME, "lblPrice_usd", acLabel, _
        gui_buildControlSizeDef(Top:=1.26, Left:=3#, Width:=0.9, Height:=0.22)) ' Right=3.90, Bottom=1.48

    Call gui_buildControl(FORM_NAME, "txtPrice_usd", acTextBox, _
        gui_buildControlSizeDef(Top:=1.26, Left:=3.95, Width:=1.6, Height:=0.22)) ' Right=5.55, Bottom=1.48

    Call gui_buildControl(FORM_NAME, "lblLength_cm", acLabel, _
        gui_buildControlSizeDef(Top:=1.53, Left:=3#, Width:=0.9, Height:=0.22)) ' Right=3.90, Bottom=1.75

    Call gui_buildControl(FORM_NAME, "txtLength_cm", acTextBox, _
        gui_buildControlSizeDef(Top:=1.53, Left:=3.95, Width:=1.6, Height:=0.22)) ' Right=5.55, Bottom=1.75

    Call gui_buildControl(FORM_NAME, "lblWidth_cm", acLabel, _
        gui_buildControlSizeDef(Top:=1.8, Left:=3#, Width:=0.9, Height:=0.22)) ' Right=3.90, Bottom=2.02

    Call gui_buildControl(FORM_NAME, "txtWidth_cm", acTextBox, _
        gui_buildControlSizeDef(Top:=1.8, Left:=3.95, Width:=1.6, Height:=0.22)) ' Right=5.55, Bottom=2.02

    Call gui_buildControl(FORM_NAME, "lblHeight_cm", acLabel, _
        gui_buildControlSizeDef(Top:=2.07, Left:=3#, Width:=0.9, Height:=0.22)) ' Right=3.90, Bottom=2.29

    Call gui_buildControl(FORM_NAME, "txtHeight_cm", acTextBox, _
        gui_buildControlSizeDef(Top:=2.07, Left:=3.95, Width:=1.6, Height:=0.22)) ' Right=5.55, Bottom=2.29

    Call gui_buildControl(FORM_NAME, "lblWeight_gram", acLabel, _
        gui_buildControlSizeDef(Top:=2.34, Left:=3#, Width:=0.9, Height:=0.22)) ' Right=3.90, Bottom=2.56

    Call gui_buildControl(FORM_NAME, "txtWeight_gram", acTextBox, _
        gui_buildControlSizeDef(Top:=2.34, Left:=3.95, Width:=1.6, Height:=0.22)) ' Right=5.55, Bottom=2.56

    ' 4. DOCUMENTATION AREA (Below Columns 1 & 2)
    Call gui_buildControl(FORM_NAME, "boxDocumentation", acRectangle, _
        gui_buildControlSizeDef(Top:=3.18, Left:=0.2, Width:=5.4, Height:=1.75, _
                                HorizontalAnchor:=acHorizontalAnchorLeft, VerticalAnchor:=acVerticalAnchorBoth)) ' Right=5.60, Bottom=4.93

    Call gui_buildControl(FORM_NAME, "lblUSP", acLabel, _
        gui_buildControlSizeDef(Top:=3.23, Left:=0.25, Width:=2.6, Height:=0.2, _
                                HorizontalAnchor:=acHorizontalAnchorLeft, VerticalAnchor:=acVerticalAnchorTop)) ' Right=2.85, Bottom=3.43

    Call gui_buildControl(FORM_NAME, "txtUSP", acTextBox, _
        gui_buildControlSizeDef(Top:=3.45, Left:=0.25, Width:=2.6, Height:=1.4, _
                                HorizontalAnchor:=acHorizontalAnchorLeft, VerticalAnchor:=acVerticalAnchorBoth)) ' Right=2.85, Bottom=4.85

    Call gui_buildControl(FORM_NAME, "lblHowToUse", acLabel, _
        gui_buildControlSizeDef(Top:=3.23, Left:=2.95, Width:=2.6, Height:=0.2, _
                                HorizontalAnchor:=acHorizontalAnchorLeft, VerticalAnchor:=acVerticalAnchorTop)) ' Right=5.55, Bottom=3.43

    Call gui_buildControl(FORM_NAME, "txtHowToUse", acTextBox, _
        gui_buildControlSizeDef(Top:=3.45, Left:=2.95, Width:=2.6, Height:=1.4, _
                                HorizontalAnchor:=acHorizontalAnchorLeft, VerticalAnchor:=acVerticalAnchorBoth)) ' Right=5.55, Bottom=4.85

    ' 5. COLUMN 3: IMAGE & INGREDIENTS
    ' Upper: Image Card (Stretches horizontally, stays top)
    Call gui_buildControl(FORM_NAME, "boxImage", acRectangle, _
        gui_buildControlSizeDef(Top:=0.1, Left:=5.7, Width:=3.75, Height:=2.35, _
                                HorizontalAnchor:=acHorizontalAnchorBoth, VerticalAnchor:=acVerticalAnchorTop)) ' Right=9.45, Bottom=2.45

    Call gui_buildControl(FORM_NAME, "lblImage", acLabel, _
        gui_buildControlSizeDef(Top:=0.15, Left:=5.75, Width:=3.65, Height:=0.2, _
                                HorizontalAnchor:=acHorizontalAnchorBoth, VerticalAnchor:=acVerticalAnchorTop)) ' Right=9.40, Bottom=0.35

    Call gui_buildControl(FORM_NAME, "imgProduct", acImage, _
        gui_buildControlSizeDef(Top:=0.4, Left:=5.75, Width:=3.65, Height:=1.6, _
                                HorizontalAnchor:=acHorizontalAnchorBoth, VerticalAnchor:=acVerticalAnchorTop)) ' Right=9.40, Bottom=2.00

    Call gui_buildControl(FORM_NAME, "LabelNoImage", acLabel, _
        gui_buildControlSizeDef(Top:=0.9, Left:=5.75, Width:=3.65, Height:=0.5, _
                                HorizontalAnchor:=acHorizontalAnchorBoth, VerticalAnchor:=acVerticalAnchorTop)) ' Right=9.40, Bottom=1.40

    Call gui_buildControl(FORM_NAME, "cmdAddImage", acCommandButton, _
        gui_buildControlSizeDef(Top:=2.05, Left:=5.75, Width:=1.75, Height:=0.32, _
                                HorizontalAnchor:=acHorizontalAnchorLeft, VerticalAnchor:=acVerticalAnchorTop)) ' Right=7.50, Bottom=2.37

    Call gui_buildControl(FORM_NAME, "cmdClearImage", acCommandButton, _
        gui_buildControlSizeDef(Top:=2.05, Left:=7.65, Width:=1.75, Height:=0.32, _
                                HorizontalAnchor:=acHorizontalAnchorRight, VerticalAnchor:=acVerticalAnchorTop)) ' Right=9.40, Bottom=2.37

    ' Lower: Ingredients Card (Stretches both horizontally and vertically)
    Call gui_buildControl(FORM_NAME, "boxIngredients", acRectangle, _
        gui_buildControlSizeDef(Top:=2.52, Left:=5.7, Width:=3.75, Height:=2.41, _
                                HorizontalAnchor:=acHorizontalAnchorBoth, VerticalAnchor:=acVerticalAnchorBoth)) ' Right=9.45, Bottom=4.93

    Call gui_buildControl(FORM_NAME, "lblIngredients", acLabel, _
        gui_buildControlSizeDef(Top:=2.57, Left:=5.75, Width:=3.65, Height:=0.2, _
                                HorizontalAnchor:=acHorizontalAnchorBoth, VerticalAnchor:=acVerticalAnchorTop)) ' Right=9.40, Bottom=2.77

    Call gui_buildControl(FORM_NAME, "txtIngredients", acTextBox, _
        gui_buildControlSizeDef(Top:=2.82, Left:=5.75, Width:=3.65, Height:=2.03, _
                                HorizontalAnchor:=acHorizontalAnchorBoth, VerticalAnchor:=acVerticalAnchorBoth)) ' Right=9.40, Bottom=4.85

    ' 6. COLUMN 4: SIDEBAR ACTIONS (Right Stack)
    Call gui_buildControl(FORM_NAME, "cmdNewProduct", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.1, Left:=9.6, Width:=1.18, Height:=0.4, _
                                HorizontalAnchor:=acHorizontalAnchorRight, VerticalAnchor:=acVerticalAnchorTop)) ' Right=10.78, Bottom=0.50

    Call gui_buildControl(FORM_NAME, "cmdExport", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.6, Left:=9.6, Width:=1.18, Height:=0.3, _
                                HorizontalAnchor:=acHorizontalAnchorRight, VerticalAnchor:=acVerticalAnchorTop)) ' Right=10.78, Bottom=0.90

    ' 7. FOOTER NAVIGATION STRIP (T=5.00" - 5.40")
    Call gui_buildControl(FORM_NAME, "linFooterSeparator", acLine, _
        gui_buildControlSizeDef(Top:=5#, Left:=0.2, Width:=9.25, Height:=0.01, _
                                HorizontalAnchor:=acHorizontalAnchorBoth, VerticalAnchor:=acVerticalAnchorBottom)) ' Right=9.45, Bottom=5.01

    Call gui_buildControl(FORM_NAME, "cmdPrev", acCommandButton, _
        gui_buildControlSizeDef(Top:=5.1, Left:=0.2, Width:=1#, Height:=0.3, _
                                HorizontalAnchor:=acHorizontalAnchorLeft, VerticalAnchor:=acVerticalAnchorBottom)) ' Right=1.20, Bottom=5.40

    Call gui_buildControl(FORM_NAME, "lblNavStatus", acLabel, _
        gui_buildControlSizeDef(Top:=5.1, Left:=1.3, Width:=1.6, Height:=0.3, _
                                HorizontalAnchor:=acHorizontalAnchorLeft, VerticalAnchor:=acVerticalAnchorBottom)) ' Right=2.90, Bottom=5.40

    Call gui_buildControl(FORM_NAME, "cmdNext", acCommandButton, _
        gui_buildControlSizeDef(Top:=5.1, Left:=3#, Width:=1#, Height:=0.3, _
                                HorizontalAnchor:=acHorizontalAnchorLeft, VerticalAnchor:=acVerticalAnchorBottom)) ' Right=4.00, Bottom=5.40

    ' [FINALIZE]: Save & Release
    DoCmd.Close acForm, FORM_NAME, acSaveYes
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "BuildForm_frmProductDetailView", Err.Description
End Sub

'--------------------------------------------------------------------
' 3.2 AESTHETICS & STYLING: frmProductDetailView
'--------------------------------------------------------------------
Public Sub StyleForm_frmProductDetailView()
    On Error GoTo ErrHandler
    Const FORM_NAME As String = "frmProductDetailView"

    ' 1. OPEN IN DESIGN VIEW
    If Not CurrentProject.AllForms(FORM_NAME).IsLoaded Then
        DoCmd.OpenForm FORM_NAME, acDesign, , , , acHidden
    End If
    Dim frm As Access.Form: Set frm = Forms(FORM_NAME)

    ' 2. LEVEL 1 & 2 BASE STYLING (Theme Presets)
    Call gui_styleForm(FORM_NAME)

    ' Column 1: General & Packaging
    Call gui_styleRectangle_Card(FORM_NAME, "boxGeneral")
    Call gui_styleLabel_SectionTitle(FORM_NAME, "lblGeneralTitle")
    Call gui_styleLabel(FORM_NAME, "lblMarketingCategory")
    Call gui_styleComboBox_Standard(FORM_NAME, "comboMarketingCategory")
    Call gui_styleLabel(FORM_NAME, "lblMarketingLine")
    Call gui_styleComboBox_Standard(FORM_NAME, "comboMarketingLine")
    Call gui_styleLabel(FORM_NAME, "lblBarcode")
    Call gui_styleTextBox_Machine(FORM_NAME, "txtBarcode")
    Call gui_styleLabel(FORM_NAME, "lblHSCode")
    Call gui_styleTextBox_Machine(FORM_NAME, "txtHSCode")
    Call gui_styleLabel(FORM_NAME, "lblProductName_en")
    Call gui_styleTextBox_Base(FORM_NAME, "txtProductName_en")
    Call gui_styleLabel(FORM_NAME, "lblProductName_vn")
    Call gui_styleTextBox_Base(FORM_NAME, "txtProductName_vn")
    Call gui_styleLabel(FORM_NAME, "lblPackageType")
    Call gui_styleComboBox_Standard(FORM_NAME, "cboPackageType")
    Call gui_styleLabel(FORM_NAME, "lblShelfLife_month")
    Call gui_styleTextBox_Logistics(FORM_NAME, "txtShelfLife_month")
    Call gui_styleLabel(FORM_NAME, "lblOuterPaperBox")
    Call gui_styleCheckBox_Standard(FORM_NAME, "blnOuterPaperBox")

    ' Column 2: Pricing & Logistics
    Call gui_styleRectangle_Card(FORM_NAME, "boxDimensions")
    Call gui_styleLabel_SectionTitle(FORM_NAME, "lblDimensionsTitle")
    Call gui_styleLabel(FORM_NAME, "lblCapacity")
    Call gui_styleTextBox_Logistics(FORM_NAME, "txtCapacity")
    Call gui_styleComboBox_Standard(FORM_NAME, "cboCapacityUnit")
    Call gui_styleLabel(FORM_NAME, "lblCapacity_oz")
    Call gui_styleTextBox_Logistics(FORM_NAME, "txtCapacity_oz")
    Call gui_styleLabel(FORM_NAME, "lblPrice_vnd")
    Call gui_styleTextBox_Value(FORM_NAME, "txtPrice_vnd")
    Call gui_styleLabel(FORM_NAME, "lblPrice_usd")
    Call gui_styleTextBox_Value(FORM_NAME, "txtPrice_usd")
    Call gui_styleLabel(FORM_NAME, "lblLength_cm")
    Call gui_styleTextBox_Logistics(FORM_NAME, "txtLength_cm")
    Call gui_styleLabel(FORM_NAME, "lblWidth_cm")
    Call gui_styleTextBox_Logistics(FORM_NAME, "txtWidth_cm")
    Call gui_styleLabel(FORM_NAME, "lblHeight_cm")
    Call gui_styleTextBox_Logistics(FORM_NAME, "txtHeight_cm")
    Call gui_styleLabel(FORM_NAME, "lblWeight_gram")
    Call gui_styleTextBox_Logistics(FORM_NAME, "txtWeight_gram")

    ' Documentation Area
    Call gui_styleRectangle_Card(FORM_NAME, "boxDocumentation")
    Call gui_styleLabel(FORM_NAME, "lblUSP")
    Call gui_styleTextBox_Document(FORM_NAME, "txtUSP")
    Call gui_styleLabel(FORM_NAME, "lblHowToUse")
    Call gui_styleTextBox_Document(FORM_NAME, "txtHowToUse")

    ' Column 3: Media & Ingredients
    Call gui_styleRectangle_Card(FORM_NAME, "boxImage")
    Call gui_styleLabel_SectionTitle(FORM_NAME, "lblImage")
    Call gui_styleImage_Showcase(FORM_NAME, "imgProduct")
    Call gui_styleLabel_Watermark(FORM_NAME, "LabelNoImage", 14)
    Call gui_styleButton_Ghost(FORM_NAME, "cmdAddImage")
    Call gui_styleButton_Danger(FORM_NAME, "cmdClearImage")

    Call gui_styleRectangle_Card(FORM_NAME, "boxIngredients")
    Call gui_styleLabel_SectionTitle(FORM_NAME, "lblIngredients")
    Call gui_styleTextBox_Document(FORM_NAME, "txtIngredients")

    ' Column 4: Actions Sidebar
    Call gui_styleButton_Power(FORM_NAME, "cmdNewProduct")
    Call gui_styleButton_Success(FORM_NAME, "cmdExport")

    ' Footer Strip
    Call gui_styleLine_Divider(FORM_NAME, "linFooterSeparator")
    Call gui_styleButton_Secondary(FORM_NAME, "cmdPrev")
    Call gui_styleLabel_Metadata(FORM_NAME, "lblNavStatus")
    Call gui_styleButton_Secondary(FORM_NAME, "cmdNext")

    ' 3. LEVEL 3 CONTEXTUAL OVERRIDES (Captions, Formats, Tooltips)
    With frm
        ' Form Chrome
        .Caption = "Product Detail Viewport"
        .Modal = False
        .MinMaxButtons = 0

        ' Section Titles
        .Controls("lblGeneralTitle").Caption = "GENERAL & PACKAGING"
        .Controls("lblDimensionsTitle").Caption = "PRICING & LOGISTICS"
        .Controls("lblImage").Caption = "PRODUCT IMAGE"
        .Controls("lblIngredients").Caption = "INGREDIENTS"

        ' Field Labels
        .Controls("lblMarketingCategory").Caption = "Category"
        .Controls("lblMarketingLine").Caption = "Marketing Line"
        .Controls("lblBarcode").Caption = "Barcode"
        .Controls("lblHSCode").Caption = "HS Code"
        .Controls("lblProductName_en").Caption = "Product (EN)"
        .Controls("lblProductName_vn").Caption = "Product (VN)"
        .Controls("lblPackageType").Caption = "Packaging"
        .Controls("lblShelfLife_month").Caption = "Shelf Life"
        .Controls("lblOuterPaperBox").Caption = "Outer Box"
        .Controls("lblCapacity").Caption = "Volume"
        .Controls("lblCapacity_oz").Caption = "Volume (Oz)"
        .Controls("lblPrice_vnd").Caption = "Price (VND)"
        .Controls("lblPrice_usd").Caption = "Price (USD)"
        .Controls("lblLength_cm").Caption = "Length (cm)"
        .Controls("lblWidth_cm").Caption = "Width (cm)"
        .Controls("lblHeight_cm").Caption = "Height (cm)"
        .Controls("lblWeight_gram").Caption = "Weight (g)"
        .Controls("lblUSP").Caption = "USP / Highlights"
        .Controls("lblHowToUse").Caption = "How To Use"

        ' Formats
        .Controls("txtBarcode").Format = "@"
        .Controls("txtHSCode").Format = "@"
        .Controls("txtProductName_en").Format = "@"
        .Controls("txtProductName_vn").Format = "@"
        .Controls("txtShelfLife_month").Format = "#,##0 ""months"""
        .Controls("txtCapacity").Format = "#,##0.00"
        .Controls("txtCapacity_oz").Format = "#,##0.00 ""oz"""
        .Controls("txtPrice_vnd").Format = "#,##0 ""VND"""
        .Controls("txtPrice_usd").Format = "$#,##0.00;($#,##0.00);""$0.00"""
        .Controls("txtLength_cm").Format = "#,##0.0 ""cm"""
        .Controls("txtWidth_cm").Format = "#,##0.0 ""cm"""
        .Controls("txtHeight_cm").Format = "#,##0.0 ""cm"""
        .Controls("txtWeight_gram").Format = "#,##0 ""g"""
        .Controls("txtUSP").Format = "@;""[Enter USP / Key Highlights]"""
        .Controls("txtHowToUse").Format = "@;""[Enter Usage Instructions]"""

        ' Media Block
        .Controls("LabelNoImage").Caption = "NO IMAGE"
        .Controls("cmdAddImage").Caption = "ADD IMAGE"
        .Controls("cmdClearImage").Caption = "CLEAR"

        ' Action Buttons
        .Controls("cmdNewProduct").Caption = "+ NEW PRODUCT"
        .Controls("cmdExport").Caption = "EXPORT EXCEL"

        ' Footer Navigation
        .Controls("cmdPrev").Caption = "< PREVIOUS"
        .Controls("lblNavStatus").Caption = "Product 0 of 0"
        .Controls("lblNavStatus").TextAlign = TextAlign_Center
        .Controls("cmdNext").Caption = "NEXT >"

        ' Tooltips
        .Controls("cmdPrev").ControlTipText = "Go to previous product record"
        .Controls("cmdNext").ControlTipText = "Go to next product record"
        .Controls("cmdNewProduct").ControlTipText = "Create a new product record"
        .Controls("cmdExport").ControlTipText = "Export active product to Excel"
        .Controls("cmdAddImage").ControlTipText = "Select and attach image for active product"
        .Controls("cmdClearImage").ControlTipText = "Remove attached image from active product"
    End With

    ' 4. SAVE & CLOSE
    DoCmd.Close acForm, FORM_NAME, acSaveYes
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "StyleForm_frmProductDetailView", Err.Description
End Sub

'====================================================================
' FORM 4: frmProductList (Datasheet Grid Viewport)
' SPEC  : Core Embedded Product Dataset Table
'====================================================================

Public Sub Generate_frmProductList()
    ' PURPOSE: One-stop generator to build wireframe and apply style.
    On Error GoTo ErrHandler
    Application.Echo False

    Call BuildForm_frmProductList
    Call StyleForm_frmProductList

    Application.Echo True
    Exit Sub

ErrHandler:
    Application.Echo True
    Err.Raise Err.Number, "Generate_frmProductList", Err.Description
End Sub

'--------------------------------------------------------------------
' 4.1 WIREFRAME BUILDER: frmProductList
'--------------------------------------------------------------------
Public Sub BuildForm_frmProductList()
    On Error GoTo ErrHandler
    Const FORM_NAME As String = "frmProductList"

    ' 1. SIZE FORM CANVAS (10.9792" x 6.0")
    Call gui_buildForm(FORM_NAME, gui_buildFormSizeDef(Width:=10.9792, Height:=6#))

    ' 2. GRID CONTROLS (T=0, H=0.25")
    Call gui_buildControl(FORM_NAME, "chkContextAction", acCheckBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=0.4, Height:=0.25)) ' Right=0.40, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtRowNum", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=0.4, Height:=0.25)) ' Right=0.40, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "comboCategory", acComboBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=1.2, Height:=0.25)) ' Right=1.20, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "comboMarketingLine", acComboBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=1.2, Height:=0.25)) ' Right=1.20, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtBarcode", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=1.2, Height:=0.25)) ' Right=1.20, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtHSCode", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=1.2, Height:=0.25)) ' Right=1.20, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtProductName", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=2.5, Height:=0.25)) ' Right=2.50, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "cboPackageType", acComboBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=1#, Height:=0.25)) ' Right=1.00, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtCapacity", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=0.8, Height:=0.25)) ' Right=0.80, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "cboCapacityUnit", acComboBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=0.6, Height:=0.25)) ' Right=0.60, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtCapacity_oz", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=0.8, Height:=0.25)) ' Right=0.80, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtPrice", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=0.9, Height:=0.25)) ' Right=0.90, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtLength_cm", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=0.7, Height:=0.25)) ' Right=0.70, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtWidth_cm", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=0.7, Height:=0.25)) ' Right=0.70, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtHeight_cm", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=0.7, Height:=0.25)) ' Right=0.70, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtWeight_gram", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=0.8, Height:=0.25)) ' Right=0.80, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtIngredients", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=2#, Height:=0.25)) ' Right=2.00, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtUSP", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=2#, Height:=0.25)) ' Right=2.00, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "txtHowToUse", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=2#, Height:=0.25)) ' Right=2.00, Bottom=0.25

    Call gui_buildControl(FORM_NAME, "blnOuterPaperBox", acCheckBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=0.4, Height:=0.2)) ' Right=0.40, Bottom=0.20

    Call gui_buildControl(FORM_NAME, "txtShelfLife_month", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=0.8, Height:=0.25)) ' Right=0.80, Bottom=0.25

    ' [FINALIZE]: Save & Release
    DoCmd.Close acForm, FORM_NAME, acSaveYes
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "BuildForm_frmProductList", Err.Description
End Sub

'--------------------------------------------------------------------
' 4.2 AESTHETICS & STYLING: frmProductList
'--------------------------------------------------------------------
Public Sub StyleForm_frmProductList()
    On Error GoTo ErrHandler
    Const FORM_NAME As String = "frmProductList"

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
    Call gui_styleComboBox_Standard(FORM_NAME, "comboCategory")
    Call gui_styleComboBox_Standard(FORM_NAME, "comboMarketingLine")
    Call gui_styleTextBox_Machine(FORM_NAME, "txtBarcode")
    Call gui_styleTextBox_Machine(FORM_NAME, "txtHSCode")
    Call gui_styleTextBox_HubTitle(FORM_NAME, "txtProductName")
    Call gui_styleComboBox_Standard(FORM_NAME, "cboPackageType")
    Call gui_styleTextBox_Logistics(FORM_NAME, "txtCapacity")
    Call gui_styleComboBox_Standard(FORM_NAME, "cboCapacityUnit")
    Call gui_styleTextBox_Logistics(FORM_NAME, "txtCapacity_oz")
    Call gui_styleTextBox_Value(FORM_NAME, "txtPrice")
    Call gui_styleTextBox_Logistics(FORM_NAME, "txtLength_cm")
    Call gui_styleTextBox_Logistics(FORM_NAME, "txtWidth_cm")
    Call gui_styleTextBox_Logistics(FORM_NAME, "txtHeight_cm")
    Call gui_styleTextBox_Logistics(FORM_NAME, "txtWeight_gram")
    Call gui_styleTextBox_Document(FORM_NAME, "txtIngredients")
    Call gui_styleTextBox_Document(FORM_NAME, "txtUSP")
    Call gui_styleTextBox_Document(FORM_NAME, "txtHowToUse")
    Call gui_styleCheckBox_Standard(FORM_NAME, "blnOuterPaperBox")
    Call gui_styleTextBox_Logistics(FORM_NAME, "txtShelfLife_month")

    ' 4. CONTEXTUAL OVERRIDES (Datasheet Captions & Formats)
    With frm
        .Controls("chkContextAction").Properties("DatasheetCaption") = "Selected"
        .Controls("txtRowNum").Properties("DatasheetCaption") = "No."
        .Controls("comboCategory").Properties("DatasheetCaption") = "Category"
        .Controls("comboMarketingLine").Properties("DatasheetCaption") = "Marketing Line"
        .Controls("txtBarcode").Properties("DatasheetCaption") = "Barcode"
        .Controls("txtHSCode").Properties("DatasheetCaption") = "HS Code"
        .Controls("txtProductName").Properties("DatasheetCaption") = "Product Name"
        .Controls("cboPackageType").Properties("DatasheetCaption") = "Packaging"
        .Controls("txtCapacity").Properties("DatasheetCaption") = "Volume"
        .Controls("cboCapacityUnit").Properties("DatasheetCaption") = "Unit"
        .Controls("txtCapacity_oz").Properties("DatasheetCaption") = "Oz"
        .Controls("txtPrice").Properties("DatasheetCaption") = "Price"
        .Controls("txtLength_cm").Properties("DatasheetCaption") = "Length"
        .Controls("txtWidth_cm").Properties("DatasheetCaption") = "Width"
        .Controls("txtHeight_cm").Properties("DatasheetCaption") = "Height"
        .Controls("txtWeight_gram").Properties("DatasheetCaption") = "Weight"
        .Controls("txtIngredients").Properties("DatasheetCaption") = "Ingredients"
        .Controls("txtUSP").Properties("DatasheetCaption") = "USP"
        .Controls("txtHowToUse").Properties("DatasheetCaption") = "How to Use"
        .Controls("blnOuterPaperBox").Properties("DatasheetCaption") = "Have Outer Box?"
        .Controls("txtShelfLife_month").Properties("DatasheetCaption") = "Shelf Life"

        ' Formats
        .Controls("txtBarcode").Format = "@"
        .Controls("txtHSCode").Format = "@"
        .Controls("txtPrice").Format = "#,##0 ""VND"""
        .Controls("txtCapacity").Format = "#,##0.00"
        .Controls("txtCapacity_oz").Format = "#,##0.00 ""oz"""
        .Controls("txtLength_cm").Format = "#,##0.0 ""cm"""
        .Controls("txtWidth_cm").Format = "#,##0.0 ""cm"""
        .Controls("txtHeight_cm").Format = "#,##0.0 ""cm"""
        .Controls("txtWeight_gram").Format = "#,##0 ""g"""
        .Controls("txtShelfLife_month").Format = "#,##0 ""months"""
    End With

    ' 5. SAVE & CLOSE
    DoCmd.Close acForm, FORM_NAME, acSaveYes
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "StyleForm_frmProductList", Err.Description
End Sub
