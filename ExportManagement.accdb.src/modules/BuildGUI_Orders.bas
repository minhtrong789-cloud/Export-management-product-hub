Attribute VB_Name = "BuildGUI_Orders"
Option Compare Database
Option Explicit

'====================================================================
' MODULE    : BuildGUI_Orders
' PURPOSE   : Layout Wireframe Builder for Orders Management Domain.
' DEPENDENCY: Extension_UIBuilder
'====================================================================

'====================================================================
' FORM: frmOrders
' SPEC: Export Document Architecture Master View
'====================================================================
Public Sub BuildForm_frmOrders()
    On Error GoTo ErrHandler
    Const FORM_NAME As String = "frmOrders"

    ' 1. SIZE FORM CANVAS (10.9792" x 6.0")
    Call gui_buildForm(FORM_NAME, gui_buildFormSizeDef(Width:=10.9792, Height:=6#))

    ' 2. BRANDING & HEADER
    Call gui_buildControl(FORM_NAME, "lblManifestTitle", acLabel, _
        gui_buildControlSizeDef(Top:=0.1, Left:=0.2, Width:=3#, Height:=0.3)) ' Right=3.20, Bottom=0.40

    Call gui_buildControl(FORM_NAME, "txtShipmentStatus", acTextBox, _
        gui_buildControlSizeDef(Top:=0.1, Left:=3.5, Width:=2#, Height:=0.3)) ' Right=5.50, Bottom=0.40

    ' 3. ADMIN DATA STRIPE
    Call gui_buildControl(FORM_NAME, "txtOrderNo", acTextBox, _
        gui_buildControlSizeDef(Top:=0.1, Left:=6#, Width:=1.5, Height:=0.3)) ' Right=7.50, Bottom=0.40

    Call gui_buildControl(FORM_NAME, "txtOrderDate", acTextBox, _
        gui_buildControlSizeDef(Top:=0.1, Left:=7.6, Width:=1.5, Height:=0.3)) ' Right=9.10, Bottom=0.40

    Call gui_buildControl(FORM_NAME, "txtLogisticsOperator", acTextBox, _
        gui_buildControlSizeDef(Top:=0.1, Left:=9.2, Width:=1.6, Height:=0.3)) ' Right=10.80, Bottom=0.40

    ' 4. PANE A: BUYER HUB (Left)
    Call gui_buildControl(FORM_NAME, "txtConsigneeName", acTextBox, _
        gui_buildControlSizeDef(Top:=0.6, Left:=0.2, Width:=5#, Height:=0.3)) ' Right=5.20, Bottom=0.90

    Call gui_buildControl(FORM_NAME, "txtConsigneeContact", acTextBox, _
        gui_buildControlSizeDef(Top:=1#, Left:=0.2, Width:=2.4, Height:=0.3)) ' Right=2.60, Bottom=1.30

    Call gui_buildControl(FORM_NAME, "txtConsigneeEmail", acTextBox, _
        gui_buildControlSizeDef(Top:=1#, Left:=2.8, Width:=2.4, Height:=0.3)) ' Right=5.20, Bottom=1.30

    Call gui_buildControl(FORM_NAME, "txtConsigneeAddress", acTextBox, _
        gui_buildControlSizeDef(Top:=1.4, Left:=0.2, Width:=5#, Height:=0.6)) ' Right=5.20, Bottom=2.00

    ' 5. PANE B: LOGISTICS SIDEBAR (Right)
    Call gui_buildControl(FORM_NAME, "txtPaymentTerm", acTextBox, _
        gui_buildControlSizeDef(Top:=0.6, Left:=5.5, Width:=2.5, Height:=0.3)) ' Right=8.00, Bottom=0.90

    Call gui_buildControl(FORM_NAME, "txtIncoterm", acTextBox, _
        gui_buildControlSizeDef(Top:=0.6, Left:=8.2, Width:=2.6, Height:=0.3)) ' Right=10.80, Bottom=0.90

    Call gui_buildControl(FORM_NAME, "txtExpectedETD", acTextBox, _
        gui_buildControlSizeDef(Top:=1#, Left:=5.5, Width:=2.5, Height:=0.3)) ' Right=8.00, Bottom=1.30

    Call gui_buildControl(FORM_NAME, "txtLoadingPort", acTextBox, _
        gui_buildControlSizeDef(Top:=1#, Left:=8.2, Width:=2.6, Height:=0.3)) ' Right=10.80, Bottom=1.30

    Call gui_buildControl(FORM_NAME, "txtDischargePort", acTextBox, _
        gui_buildControlSizeDef(Top:=1.4, Left:=5.5, Width:=2.5, Height:=0.3)) ' Right=8.00, Bottom=1.70

    Call gui_buildControl(FORM_NAME, "txtOrigin", acTextBox, _
        gui_buildControlSizeDef(Top:=1.4, Left:=8.2, Width:=1.2, Height:=0.3)) ' Right=9.40, Bottom=1.70

    Call gui_buildControl(FORM_NAME, "txtQuality", acTextBox, _
        gui_buildControlSizeDef(Top:=1.4, Left:=9.5, Width:=1.3, Height:=0.3)) ' Right=10.80, Bottom=1.70

    ' 6. SUBFORM VIEWPORT
    Call gui_buildControl(FORM_NAME, "sfOrderDetails", acSubform, _
        gui_buildControlSizeDef(Top:=2.2, Left:=0.2, Width:=10.6, Height:=3#)) ' Right=10.80, Bottom=5.20

    ' 7. SUMMARY FOOTER
    Call gui_buildControl(FORM_NAME, "txtTotalTally", acTextBox, _
        gui_buildControlSizeDef(Top:=5.4, Left:=0.2, Width:=2.5, Height:=0.3)) ' Right=2.70, Bottom=5.70

    Call gui_buildControl(FORM_NAME, "txtTotalGrossWeight", acTextBox, _
        gui_buildControlSizeDef(Top:=5.4, Left:=3#, Width:=2.5, Height:=0.3)) ' Right=5.50, Bottom=5.70

    Call gui_buildControl(FORM_NAME, "txtTotalMeasurement", acTextBox, _
        gui_buildControlSizeDef(Top:=5.4, Left:=5.8, Width:=2.5, Height:=0.3)) ' Right=8.30, Bottom=5.70

    Call gui_buildControl(FORM_NAME, "Command182", acCommandButton, _
        gui_buildControlSizeDef(Top:=5.4, Left:=8.6, Width:=1#, Height:=0.35)) ' Right=9.60, Bottom=5.75

    Call gui_buildControl(FORM_NAME, "Command183", acCommandButton, _
        gui_buildControlSizeDef(Top:=5.4, Left:=9.7, Width:=1#, Height:=0.35)) ' Right=10.70, Bottom=5.75

    ' [FINALIZE]: Save & Release
    DoCmd.Close acForm, FORM_NAME, acSaveYes
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "BuildForm_frmOrders", Err.Description
End Sub

'====================================================================
' FORM: frmOrdersList
' SPEC: Strategic Dashboard Hub (Summary KPIs + Global Filters)
'====================================================================
Public Sub BuildForm_frmOrdersList()
    On Error GoTo ErrHandler
    Const FORM_NAME As String = "frmOrdersList"

    ' 1. SIZE FORM CANVAS (10.9792" x 6.0")
    Call gui_buildForm(FORM_NAME, gui_buildFormSizeDef(Width:=10.9792, Height:=6#))

    ' 2. HEADER STRIP (T=0.21" - 0.63")
    Call gui_buildControl(FORM_NAME, "lblTitle", acLabel, _
        gui_buildControlSizeDef(Top:=0.21, Left:=0.21, Width:=2.22, Height:=0.42)) ' Right=2.43, Bottom=0.63

    Call gui_buildControl(FORM_NAME, "btnNewOrder", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.21, Left:=8.97, Width:=1.81, Height:=0.42)) ' Right=10.78, Bottom=0.63

    Call gui_buildControl(FORM_NAME, "lblStatTotal", acLabel, _
        gui_buildControlSizeDef(Top:=0.28, Left:=2.78, Width:=1.53, Height:=0.28)) ' Right=4.31, Bottom=0.56

    Call gui_buildControl(FORM_NAME, "lblStatPending", acLabel, _
        gui_buildControlSizeDef(Top:=0.28, Left:=4.51, Width:=1.53, Height:=0.28)) ' Right=6.04, Bottom=0.56

    Call gui_buildControl(FORM_NAME, "lblStatMonthlyValue", acLabel, _
        gui_buildControlSizeDef(Top:=0.28, Left:=6.25, Width:=2.08, Height:=0.28)) ' Right=8.33, Bottom=0.56

    ' 3. FILTERS RIBBON (T=0.69" - 1.25")
    Call gui_buildControl(FORM_NAME, "txtSearch", acTextBox, _
        gui_buildControlSizeDef(Top:=0.9, Left:=0.21, Width:=2.43, Height:=0.35)) ' Right=2.64, Bottom=1.25

    Call gui_buildControl(FORM_NAME, "lblFilterStatus", acLabel, _
        gui_buildControlSizeDef(Top:=0.69, Left:=2.85, Width:=1.39, Height:=0.21)) ' Right=4.24, Bottom=0.90
    Call gui_buildControl(FORM_NAME, "cboStatus", acComboBox, _
        gui_buildControlSizeDef(Top:=0.9, Left:=2.85, Width:=1.39, Height:=0.35)) ' Right=4.24, Bottom=1.25

    Call gui_buildControl(FORM_NAME, "lblFrom", acLabel, _
        gui_buildControlSizeDef(Top:=0.69, Left:=4.44, Width:=1.25, Height:=0.21)) ' Right=5.69, Bottom=0.90
    Call gui_buildControl(FORM_NAME, "dtpFrom", acTextBox, _
        gui_buildControlSizeDef(Top:=0.9, Left:=4.44, Width:=1.25, Height:=0.35)) ' Right=5.69, Bottom=1.25

    Call gui_buildControl(FORM_NAME, "lblTo", acLabel, _
        gui_buildControlSizeDef(Top:=0.69, Left:=5.9, Width:=1.25, Height:=0.21)) ' Right=7.15, Bottom=0.90
    Call gui_buildControl(FORM_NAME, "dtpTo", acTextBox, _
        gui_buildControlSizeDef(Top:=0.9, Left:=5.9, Width:=1.25, Height:=0.35)) ' Right=7.15, Bottom=1.25

    Call gui_buildControl(FORM_NAME, "btnRefresh", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.93, Left:=9.66, Width:=0.3, Height:=0.3)) ' Right=9.96, Bottom=1.23

    Call gui_buildControl(FORM_NAME, "btnExportExcel", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.93, Left:=10.03, Width:=0.3, Height:=0.3)) ' Right=10.33, Bottom=1.23

    Call gui_buildControl(FORM_NAME, "btnExportPDF", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.93, Left:=10.4, Width:=0.3, Height:=0.3)) ' Right=10.70, Bottom=1.23

    ' 4. VIEWPORT GRID
    Call gui_buildControl(FORM_NAME, "sfOrdersList", acSubform, _
        gui_buildControlSizeDef(Top:=1.46, Left:=0.21, Width:=10.56, Height:=4.03)) ' Right=10.77, Bottom=5.49

    ' 5. PAGINATION FOOTER
    Call gui_buildControl(FORM_NAME, "lblTotalRecords", acLabel, _
        gui_buildControlSizeDef(Top:=5.66, Left:=0.21, Width:=2.5, Height:=0.28)) ' Right=2.71, Bottom=5.94

    Call gui_buildControl(FORM_NAME, "btnPrev", acCommandButton, _
        gui_buildControlSizeDef(Top:=5.63, Left:=7.99, Width:=0.3, Height:=0.3)) ' Right=8.29, Bottom=5.93

    Call gui_buildControl(FORM_NAME, "lblPage", acLabel, _
        gui_buildControlSizeDef(Top:=5.66, Left:=8.36, Width:=1.25, Height:=0.28)) ' Right=9.61, Bottom=5.94

    Call gui_buildControl(FORM_NAME, "btnNext", acCommandButton, _
        gui_buildControlSizeDef(Top:=5.63, Left:=9.68, Width:=0.3, Height:=0.3)) ' Right=9.98, Bottom=5.93

    Call gui_buildControl(FORM_NAME, "cboPageSize", acComboBox, _
        gui_buildControlSizeDef(Top:=5.63, Left:=10.08, Width:=0.69, Height:=0.35)) ' Right=10.77, Bottom=5.98

    ' [FINALIZE]: Save & Release
    DoCmd.Close acForm, FORM_NAME, acSaveYes
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "BuildForm_frmOrdersList", Err.Description
End Sub

'====================================================================
' FORM: frmOrdersHub_Advanced
' SPEC: Next-Gen Order Blade Workspace (Flat-SaaS Matrix)
'====================================================================
Public Sub BuildForm_frmOrdersHub_Advanced()
    On Error GoTo ErrHandler
    Const FORM_NAME As String = "frmOrdersHub_Advanced"

    ' 1. SIZE FORM CANVAS (10.9792" x 6.0")
    Call gui_buildForm(FORM_NAME, gui_buildFormSizeDef(Width:=10.9792, Height:=6#))

    ' 2. GLOBAL HEADER (T=0 to 0.5")
    Call gui_buildControl(FORM_NAME, "boxHeaderBg", acRectangle, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=10.9792, Height:=0.5)) ' Right=10.98, Bottom=0.50

    Call gui_buildControl(FORM_NAME, "lblTitle", acLabel, _
        gui_buildControlSizeDef(Top:=0.1, Left:=0.2, Width:=2#, Height:=0.3)) ' Right=2.20, Bottom=0.40

    Call gui_buildControl(FORM_NAME, "txtSearch", acTextBox, _
        gui_buildControlSizeDef(Top:=0.1, Left:=2.3, Width:=2.5, Height:=0.3)) ' Right=4.80, Bottom=0.40

    ' 3. FILTER & KPI BAR (T=0.6" - 0.9")
    Call gui_buildControl(FORM_NAME, "cboCustomer", acComboBox, _
        gui_buildControlSizeDef(Top:=0.6, Left:=0.2, Width:=1.7, Height:=0.3)) ' Right=1.90, Bottom=0.90

    Call gui_buildControl(FORM_NAME, "cboStatus", acComboBox, _
        gui_buildControlSizeDef(Top:=0.6, Left:=2#, Width:=1.3, Height:=0.3)) ' Right=3.30, Bottom=0.90

    Call gui_buildControl(FORM_NAME, "txtDateFrom", acTextBox, _
        gui_buildControlSizeDef(Top:=0.6, Left:=3.4, Width:=1.2, Height:=0.3)) ' Right=4.60, Bottom=0.90

    Call gui_buildControl(FORM_NAME, "txtDateTo", acTextBox, _
        gui_buildControlSizeDef(Top:=0.6, Left:=4.7, Width:=1.2, Height:=0.3)) ' Right=5.90, Bottom=0.90

    Call gui_buildControl(FORM_NAME, "btnClearFilter", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.6, Left:=6.1, Width:=0.8, Height:=0.3)) ' Right=6.90, Bottom=0.90

    Call gui_buildControl(FORM_NAME, "btnKPIDraft", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.6, Left:=7#, Width:=0.8, Height:=0.3)) ' Right=7.80, Bottom=0.90

    Call gui_buildControl(FORM_NAME, "btnKPIShipped", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.6, Left:=7.9, Width:=0.8, Height:=0.3)) ' Right=8.70, Bottom=0.90

    Call gui_buildControl(FORM_NAME, "btnKPIAlert", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.6, Left:=8.8, Width:=0.8, Height:=0.3)) ' Right=9.60, Bottom=0.90

    ' 4. VIEWPORT AREA
    Call gui_buildControl(FORM_NAME, "linMainDivider", acLine, _
        gui_buildControlSizeDef(Top:=1#, Left:=0.2, Width:=9.4, Height:=0.01)) ' Right=9.60, Bottom=1.00

    Call gui_buildControl(FORM_NAME, "subOrderList", acSubform, _
        gui_buildControlSizeDef(Top:=1.1, Left:=0.2, Width:=9.4, Height:=4.325)) ' Right=9.60, Bottom=5.425

    ' 5. SIDEBAR ACTIONS
    Call gui_buildControl(FORM_NAME, "btnNewOrder", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.6, Left:=9.7, Width:=1.18, Height:=0.3)) ' Right=10.88, Bottom=0.90

    Call gui_buildControl(FORM_NAME, "btnExport", acCommandButton, _
        gui_buildControlSizeDef(Top:=1#, Left:=9.7, Width:=1.18, Height:=0.3)) ' Right=10.88, Bottom=1.30

    Call gui_buildControl(FORM_NAME, "btnPrintInvoice", acCommandButton, _
        gui_buildControlSizeDef(Top:=1.4, Left:=9.7, Width:=1.18, Height:=0.3)) ' Right=10.88, Bottom=1.70

    Call gui_buildControl(FORM_NAME, "btnPackingList", acCommandButton, _
        gui_buildControlSizeDef(Top:=1.8, Left:=9.7, Width:=1.18, Height:=0.3)) ' Right=10.88, Bottom=2.10

    Call gui_buildControl(FORM_NAME, "btnRefresh", acCommandButton, _
        gui_buildControlSizeDef(Top:=2.2, Left:=9.7, Width:=1.18, Height:=0.3)) ' Right=10.88, Bottom=2.50

    Call gui_buildControl(FORM_NAME, "togTrashbin", acToggleButton, _
        gui_buildControlSizeDef(Top:=5.525, Left:=9.7, Width:=1.18, Height:=0.3)) ' Right=10.88, Bottom=5.825

    ' 6. STATUS FOOTER
    Call gui_buildControl(FORM_NAME, "linFooterDivider", acLine, _
        gui_buildControlSizeDef(Top:=5.475, Left:=0.2, Width:=9.4, Height:=0.01)) ' Right=9.60, Bottom=5.475

    Call gui_buildControl(FORM_NAME, "lblRecordCount", acLabel, _
        gui_buildControlSizeDef(Top:=5.575, Left:=0.2, Width:=3#, Height:=0.25)) ' Right=3.20, Bottom=5.825

    Call gui_buildControl(FORM_NAME, "lblTotalAmount", acLabel, _
        gui_buildControlSizeDef(Top:=5.575, Left:=5.3, Width:=2.2, Height:=0.25)) ' Right=7.50, Bottom=5.825

    Call gui_buildControl(FORM_NAME, "txtSumAmount", acTextBox, _
        gui_buildControlSizeDef(Top:=5.525, Left:=7.6, Width:=2#, Height:=0.3)) ' Right=9.60, Bottom=5.825

    ' [FINALIZE]: Save & Release
    DoCmd.Close acForm, FORM_NAME, acSaveYes
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "BuildForm_frmOrdersHub_Advanced", Err.Description
End Sub

'====================================================================
' FORM: frmOrdersList_Sub
' SPEC: Continuous Form Grid for Order Rows
'====================================================================
Public Sub BuildForm_frmOrdersList_Sub()
    On Error GoTo ErrHandler
    Const SUB_FORM_NAME As String = "frmOrdersList_Sub"

    ' 1. SIZE CANVAS (10.5625" Width, Header=0.3125", Detail=0.3472")
    Call gui_buildForm(SUB_FORM_NAME, _
        gui_buildFormSizeDef(Width:=10.5625, Height:=0.6597, HeaderHeight:=0.3125))

    ' 2. HEADER LABELS (acHeader)
    Call gui_buildControl(SUB_FORM_NAME, "lblHdrOrderNo", acLabel, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=1.11, Height:=0.28), _
        sectionID:=acHeader) ' Right=1.11, Bottom=0.28

    Call gui_buildControl(SUB_FORM_NAME, "lblHdrDate", acLabel, _
        gui_buildControlSizeDef(Top:=0, Left:=1.11, Width:=0.83, Height:=0.28), _
        sectionID:=acHeader) ' Right=1.94, Bottom=0.28

    Call gui_buildControl(SUB_FORM_NAME, "lblHdrBuyer", acLabel, _
        gui_buildControlSizeDef(Top:=0, Left:=1.94, Width:=1.67, Height:=0.28), _
        sectionID:=acHeader) ' Right=3.61, Bottom=0.28

    Call gui_buildControl(SUB_FORM_NAME, "lblHdrStatus", acLabel, _
        gui_buildControlSizeDef(Top:=0, Left:=3.61, Width:=0.97, Height:=0.28), _
        sectionID:=acHeader) ' Right=4.58, Bottom=0.28

    Call gui_buildControl(SUB_FORM_NAME, "lblHdrETD", acLabel, _
        gui_buildControlSizeDef(Top:=0, Left:=4.58, Width:=0.83, Height:=0.28), _
        sectionID:=acHeader) ' Right=5.41, Bottom=0.28

    Call gui_buildControl(SUB_FORM_NAME, "lblHdrPOL", acLabel, _
        gui_buildControlSizeDef(Top:=0, Left:=5.41, Width:=1.04, Height:=0.28), _
        sectionID:=acHeader) ' Right=6.45, Bottom=0.28

    Call gui_buildControl(SUB_FORM_NAME, "lblHdrPOD", acLabel, _
        gui_buildControlSizeDef(Top:=0, Left:=6.45, Width:=1.04, Height:=0.28), _
        sectionID:=acHeader) ' Right=7.49, Bottom=0.28

    Call gui_buildControl(SUB_FORM_NAME, "lblHdrAmount", acLabel, _
        gui_buildControlSizeDef(Top:=0, Left:=7.49, Width:=0.97, Height:=0.28), _
        sectionID:=acHeader) ' Right=8.46, Bottom=0.28

    Call gui_buildControl(SUB_FORM_NAME, "lblHdrLogi", acLabel, _
        gui_buildControlSizeDef(Top:=0, Left:=8.46, Width:=0.83, Height:=0.28), _
        sectionID:=acHeader) ' Right=9.29, Bottom=0.28

    Call gui_buildControl(SUB_FORM_NAME, "lblHdrActions", acLabel, _
        gui_buildControlSizeDef(Top:=0, Left:=9.29, Width:=1.18, Height:=0.28), _
        sectionID:=acHeader) ' Right=10.47, Bottom=0.28

    ' 3. DETAIL ROW CONTROLS (acDetail)
    Call gui_buildControl(SUB_FORM_NAME, "txtOrderNo", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=0, Width:=1.11, Height:=0.31)) ' Right=1.11, Bottom=0.31

    Call gui_buildControl(SUB_FORM_NAME, "txtOrderDate", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=1.11, Width:=0.83, Height:=0.31)) ' Right=1.94, Bottom=0.31

    Call gui_buildControl(SUB_FORM_NAME, "txtBuyer", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=1.94, Width:=1.67, Height:=0.31)) ' Right=3.61, Bottom=0.31

    Call gui_buildControl(SUB_FORM_NAME, "txtStatus", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=3.61, Width:=0.97, Height:=0.31)) ' Right=4.58, Bottom=0.31

    Call gui_buildControl(SUB_FORM_NAME, "txtETD", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=4.58, Width:=0.83, Height:=0.31)) ' Right=5.41, Bottom=0.31

    Call gui_buildControl(SUB_FORM_NAME, "txtPOL", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=5.41, Width:=1.04, Height:=0.31)) ' Right=6.45, Bottom=0.31

    Call gui_buildControl(SUB_FORM_NAME, "txtPOD", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=6.45, Width:=1.04, Height:=0.31)) ' Right=7.49, Bottom=0.31

    Call gui_buildControl(SUB_FORM_NAME, "txtTotalAmount", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=7.49, Width:=0.97, Height:=0.31)) ' Right=8.46, Bottom=0.31

    Call gui_buildControl(SUB_FORM_NAME, "txtLogistics", acTextBox, _
        gui_buildControlSizeDef(Top:=0, Left:=8.46, Width:=0.83, Height:=0.31)) ' Right=9.29, Bottom=0.31

    Call gui_buildControl(SUB_FORM_NAME, "btnActionView", acCommandButton, _
        gui_buildControlSizeDef(Top:=0, Left:=9.29, Width:=0.38, Height:=0.31)) ' Right=9.67, Bottom=0.31

    Call gui_buildControl(SUB_FORM_NAME, "btnActionEdit", acCommandButton, _
        gui_buildControlSizeDef(Top:=0, Left:=9.71, Width:=0.38, Height:=0.31)) ' Right=10.09, Bottom=0.31

    Call gui_buildControl(SUB_FORM_NAME, "btnActionPrint", acCommandButton, _
        gui_buildControlSizeDef(Top:=0, Left:=10.13, Width:=0.38, Height:=0.31)) ' Right=10.51, Bottom=0.31

    ' [FINALIZE]: Save & Release
    DoCmd.Close acForm, SUB_FORM_NAME, acSaveYes
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "BuildForm_frmOrdersList_Sub", Err.Description
End Sub

'====================================================================
' FORM: frmOrdersVersion2
' SPEC: Scrollable Multi-Layer Order Life-cycle Hub
'====================================================================
Public Sub BuildForm_frmOrdersVersion2()
    On Error GoTo ErrHandler
    Const FORM_NAME As String = "frmOrdersVersion2"

    ' 1. SIZE CANVAS (10.9792" Width x 13.3333" Height = 15810 x 19200 twips)
    Call gui_buildForm(FORM_NAME, gui_buildFormSizeDef(Width:=10.9792, Height:=13.3333))

    ' 2. LAYER 0: PROGRESS TRACKER (T=0.04" to 0.34")
    Call gui_buildControl(FORM_NAME, "btnStDraft", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.04, Left:=0.2, Width:=2.06, Height:=0.3)) ' Right=2.26, Bottom=0.34

    Call gui_buildControl(FORM_NAME, "btnStConfirmed", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.04, Left:=2.33, Width:=2.06, Height:=0.3)) ' Right=4.39, Bottom=0.34

    Call gui_buildControl(FORM_NAME, "btnStReady", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.04, Left:=4.46, Width:=2.06, Height:=0.3)) ' Right=6.52, Bottom=0.34

    Call gui_buildControl(FORM_NAME, "btnStShipped", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.04, Left:=6.59, Width:=2.06, Height:=0.3)) ' Right=8.65, Bottom=0.34

    Call gui_buildControl(FORM_NAME, "btnStClosed", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.04, Left:=8.72, Width:=2.06, Height:=0.3)) ' Right=10.78, Bottom=0.34

    Call gui_buildControl(FORM_NAME, "divTracker", acRectangle, _
        gui_buildControlSizeDef(Top:=0.42, Left:=0.2, Width:=10.58, Height:=0.02)) ' Right=10.78, Bottom=0.44

    ' 3. LAYER 1: HEADER STRIP (T=0.50" to 0.98")
    Call gui_buildControl(FORM_NAME, "lblOrderNo", acLabel, _
        gui_buildControlSizeDef(Top:=0.5, Left:=0.2, Width:=2.42, Height:=0.19)) ' Right=2.62, Bottom=0.69
    Call gui_buildControl(FORM_NAME, "txtOrderNo", acTextBox, _
        gui_buildControlSizeDef(Top:=0.73, Left:=0.2, Width:=2.42, Height:=0.25)) ' Right=2.62, Bottom=0.98

    Call gui_buildControl(FORM_NAME, "lblDate", acLabel, _
        gui_buildControlSizeDef(Top:=0.5, Left:=2.92, Width:=2.42, Height:=0.19)) ' Right=5.34, Bottom=0.69
    Call gui_buildControl(FORM_NAME, "txtDate", acTextBox, _
        gui_buildControlSizeDef(Top:=0.73, Left:=2.92, Width:=2.42, Height:=0.25)) ' Right=5.34, Bottom=0.98

    Call gui_buildControl(FORM_NAME, "lblCustomerPO", acLabel, _
        gui_buildControlSizeDef(Top:=0.5, Left:=5.64, Width:=2.42, Height:=0.19)) ' Right=8.06, Bottom=0.69
    Call gui_buildControl(FORM_NAME, "txtCustomerPO", acTextBox, _
        gui_buildControlSizeDef(Top:=0.73, Left:=5.64, Width:=2.42, Height:=0.25)) ' Right=8.06, Bottom=0.98

    Call gui_buildControl(FORM_NAME, "lblOrderStatus", acLabel, _
        gui_buildControlSizeDef(Top:=0.5, Left:=8.36, Width:=2.42, Height:=0.19)) ' Right=10.78, Bottom=0.69
    Call gui_buildControl(FORM_NAME, "cboOrderStatus", acComboBox, _
        gui_buildControlSizeDef(Top:=0.73, Left:=8.36, Width:=2.42, Height:=0.25)) ' Right=10.78, Bottom=0.98

    Call gui_buildControl(FORM_NAME, "divHeader", acRectangle, _
        gui_buildControlSizeDef(Top:=1.06, Left:=0.2, Width:=10.58, Height:=0.02)) ' Right=10.78, Bottom=1.08

    ' 4. LAYER 2A & 2B: PARTIES & SHIPPING (T=1.15" - 4.90")
    Call gui_buildControl(FORM_NAME, "lblSecParties", acLabel, _
        gui_buildControlSizeDef(Top:=1.15, Left:=0.2, Width:=5.14, Height:=0.21)) ' Right=5.34, Bottom=1.36
    Call gui_buildControl(FORM_NAME, "lblConsignee", acLabel, _
        gui_buildControlSizeDef(Top:=1.44, Left:=0.2, Width:=5.14, Height:=0.19)) ' Right=5.34, Bottom=1.63
    Call gui_buildControl(FORM_NAME, "cboConsignee", acComboBox, _
        gui_buildControlSizeDef(Top:=1.67, Left:=0.2, Width:=5.14, Height:=0.25)) ' Right=5.34, Bottom=1.92
    Call gui_buildControl(FORM_NAME, "lblConsigneeAdd", acLabel, _
        gui_buildControlSizeDef(Top:=2.17, Left:=0.2, Width:=5.14, Height:=0.19)) ' Right=5.34, Bottom=2.36
    Call gui_buildControl(FORM_NAME, "txtConsigneeAdd", acTextBox, _
        gui_buildControlSizeDef(Top:=2.4, Left:=0.2, Width:=5.14, Height:=0.5)) ' Right=5.34, Bottom=2.90
    Call gui_buildControl(FORM_NAME, "lblNotifyParty", acLabel, _
        gui_buildControlSizeDef(Top:=3.15, Left:=0.2, Width:=5.14, Height:=0.19)) ' Right=5.34, Bottom=3.34
    Call gui_buildControl(FORM_NAME, "cboNotifyParty", acComboBox, _
        gui_buildControlSizeDef(Top:=3.38, Left:=0.2, Width:=5.14, Height:=0.25)) ' Right=5.34, Bottom=3.63
    Call gui_buildControl(FORM_NAME, "lblSecDocs", acLabel, _
        gui_buildControlSizeDef(Top:=4.13, Left:=0.2, Width:=5.14, Height:=0.21)) ' Right=5.34, Bottom=4.34
    Call gui_buildControl(FORM_NAME, "lblCOType", acLabel, _
        gui_buildControlSizeDef(Top:=4.42, Left:=0.2, Width:=5.14, Height:=0.19)) ' Right=5.34, Bottom=4.61
    Call gui_buildControl(FORM_NAME, "cboCOType", acComboBox, _
        gui_buildControlSizeDef(Top:=4.65, Left:=0.2, Width:=5.14, Height:=0.25)) ' Right=5.34, Bottom=4.90

    Call gui_buildControl(FORM_NAME, "lblSecShipping", acLabel, _
        gui_buildControlSizeDef(Top:=1.15, Left:=5.64, Width:=5.14, Height:=0.21)) ' Right=10.78, Bottom=1.36
    Call gui_buildControl(FORM_NAME, "lblIncoterms", acLabel, _
        gui_buildControlSizeDef(Top:=1.44, Left:=5.64, Width:=2.42, Height:=0.19)) ' Right=8.06, Bottom=1.63
    Call gui_buildControl(FORM_NAME, "cboIncoterms", acComboBox, _
        gui_buildControlSizeDef(Top:=1.67, Left:=5.64, Width:=2.42, Height:=0.25)) ' Right=8.06, Bottom=1.92
    Call gui_buildControl(FORM_NAME, "lblCurrency", acLabel, _
        gui_buildControlSizeDef(Top:=1.44, Left:=8.36, Width:=2.42, Height:=0.19)) ' Right=10.78, Bottom=1.63
    Call gui_buildControl(FORM_NAME, "cboCurrency", acComboBox, _
        gui_buildControlSizeDef(Top:=1.67, Left:=8.36, Width:=2.42, Height:=0.25)) ' Right=10.78, Bottom=1.92

    Call gui_buildControl(FORM_NAME, "lblPOL", acLabel, _
        gui_buildControlSizeDef(Top:=2.17, Left:=5.64, Width:=2.42, Height:=0.19)) ' Right=8.06, Bottom=2.36
    Call gui_buildControl(FORM_NAME, "cboPOL", acComboBox, _
        gui_buildControlSizeDef(Top:=2.4, Left:=5.64, Width:=2.42, Height:=0.25)) ' Right=8.06, Bottom=2.65
    Call gui_buildControl(FORM_NAME, "lblPOD", acLabel, _
        gui_buildControlSizeDef(Top:=2.17, Left:=8.36, Width:=2.42, Height:=0.19)) ' Right=10.78, Bottom=2.36
    Call gui_buildControl(FORM_NAME, "cboPOD", acComboBox, _
        gui_buildControlSizeDef(Top:=2.4, Left:=8.36, Width:=2.42, Height:=0.25)) ' Right=10.78, Bottom=2.65

    Call gui_buildControl(FORM_NAME, "lblETD", acLabel, _
        gui_buildControlSizeDef(Top:=2.9, Left:=5.64, Width:=2.42, Height:=0.19)) ' Right=8.06, Bottom=3.09
    Call gui_buildControl(FORM_NAME, "dtpETD", acTextBox, _
        gui_buildControlSizeDef(Top:=3.13, Left:=5.64, Width:=2.42, Height:=0.25)) ' Right=8.06, Bottom=3.38
    Call gui_buildControl(FORM_NAME, "lblETA", acLabel, _
        gui_buildControlSizeDef(Top:=2.9, Left:=8.36, Width:=2.42, Height:=0.19)) ' Right=10.78, Bottom=3.09
    Call gui_buildControl(FORM_NAME, "dtpETA", acTextBox, _
        gui_buildControlSizeDef(Top:=3.13, Left:=8.36, Width:=2.42, Height:=0.25)) ' Right=10.78, Bottom=3.38

    Call gui_buildControl(FORM_NAME, "lblVessel", acLabel, _
        gui_buildControlSizeDef(Top:=3.63, Left:=5.64, Width:=2.42, Height:=0.19)) ' Right=8.06, Bottom=3.82
    Call gui_buildControl(FORM_NAME, "txtVessel", acTextBox, _
        gui_buildControlSizeDef(Top:=3.85, Left:=5.64, Width:=2.42, Height:=0.25)) ' Right=8.06, Bottom=4.10
    Call gui_buildControl(FORM_NAME, "lblBookingNo", acLabel, _
        gui_buildControlSizeDef(Top:=3.63, Left:=8.36, Width:=2.42, Height:=0.19)) ' Right=10.78, Bottom=3.82
    Call gui_buildControl(FORM_NAME, "txtBookingNo", acTextBox, _
        gui_buildControlSizeDef(Top:=3.85, Left:=8.36, Width:=2.42, Height:=0.25)) ' Right=10.78, Bottom=4.10

    Call gui_buildControl(FORM_NAME, "lblBLNo", acLabel, _
        gui_buildControlSizeDef(Top:=4.35, Left:=5.64, Width:=2.42, Height:=0.19)) ' Right=8.06, Bottom=4.54
    Call gui_buildControl(FORM_NAME, "txtBLNo", acTextBox, _
        gui_buildControlSizeDef(Top:=4.58, Left:=5.64, Width:=2.42, Height:=0.25)) ' Right=8.06, Bottom=4.83
    Call gui_buildControl(FORM_NAME, "lblContainerType", acLabel, _
        gui_buildControlSizeDef(Top:=4.35, Left:=8.36, Width:=2.42, Height:=0.19)) ' Right=10.78, Bottom=4.54
    Call gui_buildControl(FORM_NAME, "cboContainerType", acComboBox, _
        gui_buildControlSizeDef(Top:=4.58, Left:=8.36, Width:=2.42, Height:=0.25)) ' Right=10.78, Bottom=4.83

    Call gui_buildControl(FORM_NAME, "divColumns", acRectangle, _
        gui_buildControlSizeDef(Top:=4.98, Left:=0.2, Width:=10.58, Height:=0.02)) ' Right=10.78, Bottom=5.00

    ' 5. LAYER 3: INTERNAL ROW (T=5.06" - 5.86")
    Call gui_buildControl(FORM_NAME, "lblSecInternal", acLabel, _
        gui_buildControlSizeDef(Top:=5.06, Left:=0.2, Width:=10.58, Height:=0.21)) ' Right=10.78, Bottom=5.27
    Call gui_buildControl(FORM_NAME, "lblFactory", acLabel, _
        gui_buildControlSizeDef(Top:=5.35, Left:=0.2, Width:=3.26, Height:=0.19)) ' Right=3.46, Bottom=5.54
    Call gui_buildControl(FORM_NAME, "cboFactory", acComboBox, _
        gui_buildControlSizeDef(Top:=5.58, Left:=0.2, Width:=3.26, Height:=0.25)) ' Right=3.46, Bottom=5.83
    Call gui_buildControl(FORM_NAME, "lblReadiness", acLabel, _
        gui_buildControlSizeDef(Top:=5.35, Left:=3.76, Width:=2.08, Height:=0.19)) ' Right=5.84, Bottom=5.54
    Call gui_buildControl(FORM_NAME, "txtReadinessDate", acTextBox, _
        gui_buildControlSizeDef(Top:=5.58, Left:=3.76, Width:=2.08, Height:=0.25)) ' Right=5.84, Bottom=5.83
    Call gui_buildControl(FORM_NAME, "lblLCNo", acLabel, _
        gui_buildControlSizeDef(Top:=5.35, Left:=6.15, Width:=4.63, Height:=0.19)) ' Right=10.78, Bottom=5.54
    Call gui_buildControl(FORM_NAME, "txtLCNo", acTextBox, _
        gui_buildControlSizeDef(Top:=5.58, Left:=6.15, Width:=4.63, Height:=0.25)) ' Right=10.78, Bottom=5.83

    Call gui_buildControl(FORM_NAME, "divInternal", acRectangle, _
        gui_buildControlSizeDef(Top:=5.92, Left:=0.2, Width:=10.58, Height:=0.02)) ' Right=10.78, Bottom=5.94

    ' 6. LAYER 4: ORDER LINE ITEMS (T=6.00" - 10.29")
    Call gui_buildControl(FORM_NAME, "lblSecGrid", acLabel, _
        gui_buildControlSizeDef(Top:=6#, Left:=0.2, Width:=10.58, Height:=0.21)) ' Right=10.78, Bottom=6.21
    Call gui_buildControl(FORM_NAME, "subOrderDetails", acSubform, _
        gui_buildControlSizeDef(Top:=6.29, Left:=0.2, Width:=10.58, Height:=4#)) ' Right=10.78, Bottom=10.29

    Call gui_buildControl(FORM_NAME, "divGrid", acRectangle, _
        gui_buildControlSizeDef(Top:=10.38, Left:=0.2, Width:=10.58, Height:=0.02)) ' Right=10.78, Bottom=10.40

    ' 7. LAYER 4.5: FINANCIAL SUMMARY (T=10.46" - 11.23")
    Call gui_buildControl(FORM_NAME, "lblSecFinancial", acLabel, _
        gui_buildControlSizeDef(Top:=10.46, Left:=0.2, Width:=10.58, Height:=0.21)) ' Right=10.78, Bottom=10.67

    Call gui_buildControl(FORM_NAME, "lblSumAmount", acLabel, _
        gui_buildControlSizeDef(Top:=10.75, Left:=0.2, Width:=2.42, Height:=0.19)) ' Right=2.62, Bottom=10.94
    Call gui_buildControl(FORM_NAME, "txtSumAmount", acTextBox, _
        gui_buildControlSizeDef(Top:=10.98, Left:=0.2, Width:=2.42, Height:=0.25)) ' Right=2.62, Bottom=11.23

    Call gui_buildControl(FORM_NAME, "lblSumGW", acLabel, _
        gui_buildControlSizeDef(Top:=10.75, Left:=2.92, Width:=2.42, Height:=0.19)) ' Right=5.34, Bottom=10.94
    Call gui_buildControl(FORM_NAME, "txtSumGW", acTextBox, _
        gui_buildControlSizeDef(Top:=10.98, Left:=2.92, Width:=2.42, Height:=0.25)) ' Right=5.34, Bottom=11.23

    Call gui_buildControl(FORM_NAME, "lblSumCBM", acLabel, _
        gui_buildControlSizeDef(Top:=10.75, Left:=5.64, Width:=2.42, Height:=0.19)) ' Right=8.06, Bottom=10.94
    Call gui_buildControl(FORM_NAME, "txtSumCBM", acTextBox, _
        gui_buildControlSizeDef(Top:=10.98, Left:=5.64, Width:=2.42, Height:=0.25)) ' Right=8.06, Bottom=11.23

    Call gui_buildControl(FORM_NAME, "lblSumCartons", acLabel, _
        gui_buildControlSizeDef(Top:=10.75, Left:=8.36, Width:=2.42, Height:=0.19)) ' Right=10.78, Bottom=10.94
    Call gui_buildControl(FORM_NAME, "txtSumCartons", acTextBox, _
        gui_buildControlSizeDef(Top:=10.98, Left:=8.36, Width:=2.42, Height:=0.25)) ' Right=10.78, Bottom=11.23

    Call gui_buildControl(FORM_NAME, "divFinancial", acRectangle, _
        gui_buildControlSizeDef(Top:=11.31, Left:=0.2, Width:=10.58, Height:=0.02)) ' Right=10.78, Bottom=11.33

    ' 8. LAYER 5: NOTES & REMARKS (T=11.40" - 12.69")
    Call gui_buildControl(FORM_NAME, "lblSecRemarks", acLabel, _
        gui_buildControlSizeDef(Top:=11.4, Left:=0.2, Width:=10.58, Height:=0.21)) ' Right=10.78, Bottom=11.61
    Call gui_buildControl(FORM_NAME, "txtRemarks", acTextBox, _
        gui_buildControlSizeDef(Top:=11.69, Left:=0.2, Width:=10.58, Height:=1#)) ' Right=10.78, Bottom=12.69

    Call gui_buildControl(FORM_NAME, "divRemarks", acRectangle, _
        gui_buildControlSizeDef(Top:=12.77, Left:=0.2, Width:=10.58, Height:=0.02)) ' Right=10.78, Bottom=12.79

    ' 9. LAYER 6: FOOTER (T=12.85" - 13.15")
    Call gui_buildControl(FORM_NAME, "btnPrevRecord", acCommandButton, _
        gui_buildControlSizeDef(Top:=12.85, Left:=0.2, Width:=1#, Height:=0.3)) ' Right=1.20, Bottom=13.15

    Call gui_buildControl(FORM_NAME, "lblRecNav", acLabel, _
        gui_buildControlSizeDef(Top:=12.88, Left:=1.35, Width:=2#, Height:=0.25)) ' Right=3.35, Bottom=13.13

    Call gui_buildControl(FORM_NAME, "btnNextRecord", acCommandButton, _
        gui_buildControlSizeDef(Top:=12.85, Left:=3.5, Width:=1#, Height:=0.3)) ' Right=4.50, Bottom=13.15

    Call gui_buildControl(FORM_NAME, "cmdSave", acCommandButton, _
        gui_buildControlSizeDef(Top:=12.85, Left:=4.83, Width:=1#, Height:=0.3)) ' Right=5.83, Bottom=13.15

    Call gui_buildControl(FORM_NAME, "cmdInvoice", acCommandButton, _
        gui_buildControlSizeDef(Top:=12.85, Left:=5.98, Width:=1.5, Height:=0.3)) ' Right=7.48, Bottom=13.15

    Call gui_buildControl(FORM_NAME, "cmdPacking", acCommandButton, _
        gui_buildControlSizeDef(Top:=12.85, Left:=7.63, Width:=1.5, Height:=0.3)) ' Right=9.13, Bottom=13.15

    Call gui_buildControl(FORM_NAME, "cmdPicking", acCommandButton, _
        gui_buildControlSizeDef(Top:=12.85, Left:=9.28, Width:=1.5, Height:=0.3)) ' Right=10.78, Bottom=13.15

    ' [FINALIZE]: Save & Release
    DoCmd.Close acForm, FORM_NAME, acSaveYes
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "BuildForm_frmOrdersVersion2", Err.Description
End Sub

'====================================================================
' FORM: frmOrdersTabbed
' SPEC: High-Density Multi-Tab Order Studio (10.98" x 6.0")
'====================================================================
Public Sub BuildForm_frmOrdersTabbed()
    On Error GoTo ErrHandler
    Const FORM_NAME As String = "frmOrdersTabbed"

    ' 1. SIZE CANVAS (10.9792" x 6.0")
    Call gui_buildForm(FORM_NAME, gui_buildFormSizeDef(Width:=10.9792, Height:=6#))

    ' 2. HEADER PROGRESS STRIP (T=0.08" to 0.38")
    Call gui_buildControl(FORM_NAME, "btnStDraft", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.08, Left:=0.2, Width:=1.25, Height:=0.29)) ' Right=1.45, Bottom=0.37

    Call gui_buildControl(FORM_NAME, "btnStConfirmed", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.08, Left:=1.5, Width:=1.25, Height:=0.29)) ' Right=2.75, Bottom=0.37

    Call gui_buildControl(FORM_NAME, "btnStReady", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.08, Left:=2.8, Width:=1.25, Height:=0.29)) ' Right=4.05, Bottom=0.37

    Call gui_buildControl(FORM_NAME, "btnStShipped", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.08, Left:=4.1, Width:=1.25, Height:=0.29)) ' Right=5.35, Bottom=0.37

    Call gui_buildControl(FORM_NAME, "btnStClosed", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.08, Left:=5.4, Width:=1.25, Height:=0.29)) ' Right=6.65, Bottom=0.37

    ' Action Buttons (Right Aligned)
    Call gui_buildControl(FORM_NAME, "cmdSave", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.08, Left:=6.81, Width:=0.96, Height:=0.29)) ' Right=7.77, Bottom=0.37

    Call gui_buildControl(FORM_NAME, "cmdInvoice", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.08, Left:=7.81, Width:=0.96, Height:=0.29)) ' Right=8.77, Bottom=0.37

    Call gui_buildControl(FORM_NAME, "cmdPacking", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.08, Left:=8.81, Width:=0.96, Height:=0.29)) ' Right=9.77, Bottom=0.37

    Call gui_buildControl(FORM_NAME, "cmdPicking", acCommandButton, _
        gui_buildControlSizeDef(Top:=0.08, Left:=9.81, Width:=0.96, Height:=0.29)) ' Right=10.77, Bottom=0.37

    Call gui_buildControl(FORM_NAME, "divTracker", acRectangle, _
        gui_buildControlSizeDef(Top:=0.46, Left:=0.2, Width:=10.58, Height:=0.02)) ' Right=10.78, Bottom=0.48

    ' 3. METADATA RIBBON (T=0.56" to 1.02")
    Call gui_buildControl(FORM_NAME, "lblOrderNo", acLabel, _
        gui_buildControlSizeDef(Top:=0.56, Left:=0.2, Width:=2.49, Height:=0.19)) ' Right=2.69, Bottom=0.75
    Call gui_buildControl(FORM_NAME, "txtOrderNo", acTextBox, _
        gui_buildControlSizeDef(Top:=0.77, Left:=0.2, Width:=2.49, Height:=0.25)) ' Right=2.69, Bottom=1.02

    Call gui_buildControl(FORM_NAME, "lblDate", acLabel, _
        gui_buildControlSizeDef(Top:=0.56, Left:=2.9, Width:=2.49, Height:=0.19)) ' Right=5.39, Bottom=0.75
    Call gui_buildControl(FORM_NAME, "txtDate", acTextBox, _
        gui_buildControlSizeDef(Top:=0.77, Left:=2.9, Width:=2.49, Height:=0.25)) ' Right=5.39, Bottom=1.02

    Call gui_buildControl(FORM_NAME, "lblCustomerPO", acLabel, _
        gui_buildControlSizeDef(Top:=0.56, Left:=5.59, Width:=2.49, Height:=0.19)) ' Right=8.08, Bottom=0.75
    Call gui_buildControl(FORM_NAME, "txtCustomerPO", acTextBox, _
        gui_buildControlSizeDef(Top:=0.77, Left:=5.59, Width:=2.49, Height:=0.25)) ' Right=8.08, Bottom=1.02

    Call gui_buildControl(FORM_NAME, "lblOrderStatus", acLabel, _
        gui_buildControlSizeDef(Top:=0.56, Left:=8.29, Width:=2.49, Height:=0.19)) ' Right=10.78, Bottom=0.75
    Call gui_buildControl(FORM_NAME, "cboOrderStatus", acComboBox, _
        gui_buildControlSizeDef(Top:=0.77, Left:=8.29, Width:=2.49, Height:=0.25)) ' Right=10.78, Bottom=1.02

    Call gui_buildControl(FORM_NAME, "divHeader", acRectangle, _
        gui_buildControlSizeDef(Top:=1.1, Left:=0.2, Width:=10.58, Height:=0.02)) ' Right=10.78, Bottom=1.12

    ' 4. TAB CONTAINER (T=1.21", H=4.71")
    Call gui_buildControl(FORM_NAME, "tabMain", 123, _
        gui_buildControlSizeDef(Top:=1.21, Left:=0.2, Width:=10.58, Height:=4.71)) ' Right=10.78, Bottom=5.92

    ' Configure 3 Tab Pages
    Dim frm As Access.Form: Set frm = Forms(FORM_NAME)
    Dim tbc As TabControl: Set tbc = frm.Controls("tabMain")
    While tbc.Pages.Count < 3
        Application.CreateControl FORM_NAME, 124, acDetail, "tabMain"
    Wend
    tbc.Pages(0).name = "pgLogistics":  tbc.Pages(0).Caption = " 1. PARTIES & LOGISTICS "
    tbc.Pages(1).name = "pgLineItems":  tbc.Pages(1).Caption = " 2. ORDER LINE ITEMS "
    tbc.Pages(2).name = "pgFinancials": tbc.Pages(2).Caption = " 3. FINANCIALS & REMARKS "

    ' PAGE 1: pgLogistics Controls
    Call gui_buildControl(FORM_NAME, "lblConsignee", acLabel, _
        gui_buildControlSizeDef(Top:=0.33, Left:=0.2, Width:=4.93, Height:=0.19), _
        parentName:="pgLogistics") ' Right=5.13, Bottom=0.52
    Call gui_buildControl(FORM_NAME, "cboConsignee", acComboBox, _
        gui_buildControlSizeDef(Top:=0.54, Left:=0.2, Width:=4.93, Height:=0.25), _
        parentName:="pgLogistics") ' Right=5.13, Bottom=0.79

    Call gui_buildControl(FORM_NAME, "lblIncoterms", acLabel, _
        gui_buildControlSizeDef(Top:=0.33, Left:=5.34, Width:=2.62, Height:=0.19), _
        parentName:="pgLogistics") ' Right=7.96, Bottom=0.52
    Call gui_buildControl(FORM_NAME, "cboIncoterms", acComboBox, _
        gui_buildControlSizeDef(Top:=0.54, Left:=5.34, Width:=2.62, Height:=0.25), _
        parentName:="pgLogistics") ' Right=7.96, Bottom=0.79

    Call gui_buildControl(FORM_NAME, "lblCurrency", acLabel, _
        gui_buildControlSizeDef(Top:=0.33, Left:=8.16, Width:=2.62, Height:=0.19), _
        parentName:="pgLogistics") ' Right=10.78, Bottom=0.52
    Call gui_buildControl(FORM_NAME, "cboCurrency", acComboBox, _
        gui_buildControlSizeDef(Top:=0.54, Left:=8.16, Width:=2.62, Height:=0.25), _
        parentName:="pgLogistics") ' Right=10.78, Bottom=0.79

    Call gui_buildControl(FORM_NAME, "lblConsigneeAdd", acLabel, _
        gui_buildControlSizeDef(Top:=1.06, Left:=0.2, Width:=4.93, Height:=0.19), _
        parentName:="pgLogistics") ' Right=5.13, Bottom=1.25
    Call gui_buildControl(FORM_NAME, "txtConsigneeAdd", acTextBox, _
        gui_buildControlSizeDef(Top:=1.27, Left:=0.2, Width:=4.93, Height:=0.98), _
        parentName:="pgLogistics") ' Right=5.13, Bottom=2.25

    Call gui_buildControl(FORM_NAME, "lblPOL", acLabel, _
        gui_buildControlSizeDef(Top:=1.06, Left:=5.34, Width:=2.62, Height:=0.19), _
        parentName:="pgLogistics") ' Right=7.96, Bottom=1.25
    Call gui_buildControl(FORM_NAME, "cboPOL", acComboBox, _
        gui_buildControlSizeDef(Top:=1.27, Left:=5.34, Width:=2.62, Height:=0.25), _
        parentName:="pgLogistics") ' Right=7.96, Bottom=1.52

    Call gui_buildControl(FORM_NAME, "lblPOD", acLabel, _
        gui_buildControlSizeDef(Top:=1.06, Left:=8.16, Width:=2.62, Height:=0.19), _
        parentName:="pgLogistics") ' Right=10.78, Bottom=1.25
    Call gui_buildControl(FORM_NAME, "cboPOD", acComboBox, _
        gui_buildControlSizeDef(Top:=1.27, Left:=8.16, Width:=2.62, Height:=0.25), _
        parentName:="pgLogistics") ' Right=10.78, Bottom=1.52

    Call gui_buildControl(FORM_NAME, "lblETD", acLabel, _
        gui_buildControlSizeDef(Top:=1.79, Left:=5.34, Width:=2.62, Height:=0.19), _
        parentName:="pgLogistics") ' Right=7.96, Bottom=1.98
    Call gui_buildControl(FORM_NAME, "dtpETD", acTextBox, _
        gui_buildControlSizeDef(Top:=2#, Left:=5.34, Width:=2.62, Height:=0.25), _
        parentName:="pgLogistics") ' Right=7.96, Bottom=2.25

    Call gui_buildControl(FORM_NAME, "lblETA", acLabel, _
        gui_buildControlSizeDef(Top:=1.79, Left:=8.16, Width:=2.62, Height:=0.19), _
        parentName:="pgLogistics") ' Right=10.78, Bottom=1.98
    Call gui_buildControl(FORM_NAME, "dtpETA", acTextBox, _
        gui_buildControlSizeDef(Top:=2#, Left:=8.16, Width:=2.62, Height:=0.25), _
        parentName:="pgLogistics") ' Right=10.78, Bottom=2.25

    Call gui_buildControl(FORM_NAME, "lblNotifyParty", acLabel, _
        gui_buildControlSizeDef(Top:=2.52, Left:=0.2, Width:=4.93, Height:=0.19), _
        parentName:="pgLogistics") ' Right=5.13, Bottom=2.71
    Call gui_buildControl(FORM_NAME, "cboNotifyParty", acComboBox, _
        gui_buildControlSizeDef(Top:=2.73, Left:=0.2, Width:=4.93, Height:=0.25), _
        parentName:="pgLogistics") ' Right=5.13, Bottom=2.98

    Call gui_buildControl(FORM_NAME, "lblVessel", acLabel, _
        gui_buildControlSizeDef(Top:=2.52, Left:=5.34, Width:=2.62, Height:=0.19), _
        parentName:="pgLogistics") ' Right=7.96, Bottom=2.71
    Call gui_buildControl(FORM_NAME, "txtVessel", acTextBox, _
        gui_buildControlSizeDef(Top:=2.73, Left:=5.34, Width:=2.62, Height:=0.25), _
        parentName:="pgLogistics") ' Right=7.96, Bottom=2.98

    Call gui_buildControl(FORM_NAME, "lblBookingNo", acLabel, _
        gui_buildControlSizeDef(Top:=2.52, Left:=8.16, Width:=2.62, Height:=0.19), _
        parentName:="pgLogistics") ' Right=10.78, Bottom=2.71
    Call gui_buildControl(FORM_NAME, "txtBookingNo", acTextBox, _
        gui_buildControlSizeDef(Top:=2.73, Left:=8.16, Width:=2.62, Height:=0.25), _
        parentName:="pgLogistics") ' Right=10.78, Bottom=2.98

    Call gui_buildControl(FORM_NAME, "lblCOType", acLabel, _
        gui_buildControlSizeDef(Top:=3.25, Left:=0.2, Width:=4.93, Height:=0.19), _
        parentName:="pgLogistics") ' Right=5.13, Bottom=3.44
    Call gui_buildControl(FORM_NAME, "cboCOType", acComboBox, _
        gui_buildControlSizeDef(Top:=3.46, Left:=0.2, Width:=4.93, Height:=0.25), _
        parentName:="pgLogistics") ' Right=5.13, Bottom=3.71

    Call gui_buildControl(FORM_NAME, "lblBLNo", acLabel, _
        gui_buildControlSizeDef(Top:=3.25, Left:=5.34, Width:=2.62, Height:=0.19), _
        parentName:="pgLogistics") ' Right=7.96, Bottom=3.44
    Call gui_buildControl(FORM_NAME, "txtBLNo", acTextBox, _
        gui_buildControlSizeDef(Top:=3.46, Left:=5.34, Width:=2.62, Height:=0.25), _
        parentName:="pgLogistics") ' Right=7.96, Bottom=3.71

    Call gui_buildControl(FORM_NAME, "lblContainerType", acLabel, _
        gui_buildControlSizeDef(Top:=3.25, Left:=8.16, Width:=2.62, Height:=0.19), _
        parentName:="pgLogistics") ' Right=10.78, Bottom=3.44
    Call gui_buildControl(FORM_NAME, "cboContainerType", acComboBox, _
        gui_buildControlSizeDef(Top:=3.46, Left:=8.16, Width:=2.62, Height:=0.25), _
        parentName:="pgLogistics") ' Right=10.78, Bottom=3.71

    ' PAGE 2: pgLineItems Subform
    Call gui_buildControl(FORM_NAME, "subOrderDetails", acSubform, _
        gui_buildControlSizeDef(Top:=0.33, Left:=0.2, Width:=10.58, Height:=4#), _
        parentName:="pgLineItems") ' Right=10.78, Bottom=4.33

    ' PAGE 3: pgFinancials Controls
    Call gui_buildControl(FORM_NAME, "lblFactory", acLabel, _
        gui_buildControlSizeDef(Top:=0.33, Left:=0.2, Width:=3.18, Height:=0.19), _
        parentName:="pgFinancials") ' Right=3.38, Bottom=0.52
    Call gui_buildControl(FORM_NAME, "cboFactory", acComboBox, _
        gui_buildControlSizeDef(Top:=0.54, Left:=0.2, Width:=3.18, Height:=0.25), _
        parentName:="pgFinancials") ' Right=3.38, Bottom=0.79

    Call gui_buildControl(FORM_NAME, "lblReadiness", acLabel, _
        gui_buildControlSizeDef(Top:=0.33, Left:=3.59, Width:=3.18, Height:=0.19), _
        parentName:="pgFinancials") ' Right=6.77, Bottom=0.52
    Call gui_buildControl(FORM_NAME, "txtReadinessDate", acTextBox, _
        gui_buildControlSizeDef(Top:=0.54, Left:=3.59, Width:=3.18, Height:=0.25), _
        parentName:="pgFinancials") ' Right=6.77, Bottom=0.79

    Call gui_buildControl(FORM_NAME, "lblLCNo", acLabel, _
        gui_buildControlSizeDef(Top:=0.33, Left:=6.98, Width:=3.8, Height:=0.19), _
        parentName:="pgFinancials") ' Right=10.78, Bottom=0.52
    Call gui_buildControl(FORM_NAME, "txtLCNo", acTextBox, _
        gui_buildControlSizeDef(Top:=0.54, Left:=6.98, Width:=3.8, Height:=0.25), _
        parentName:="pgFinancials") ' Right=10.78, Bottom=0.79

    Call gui_buildControl(FORM_NAME, "lblSumAmount", acLabel, _
        gui_buildControlSizeDef(Top:=1.06, Left:=0.2, Width:=2.49, Height:=0.19), _
        parentName:="pgFinancials") ' Right=2.69, Bottom=1.25
    Call gui_buildControl(FORM_NAME, "txtSumAmount", acTextBox, _
        gui_buildControlSizeDef(Top:=1.27, Left:=0.2, Width:=2.49, Height:=0.25), _
        parentName:="pgFinancials") ' Right=2.69, Bottom=1.52

    Call gui_buildControl(FORM_NAME, "lblSumGW", acLabel, _
        gui_buildControlSizeDef(Top:=1.06, Left:=2.9, Width:=2.49, Height:=0.19), _
        parentName:="pgFinancials") ' Right=5.39, Bottom=1.25
    Call gui_buildControl(FORM_NAME, "txtSumGW", acTextBox, _
        gui_buildControlSizeDef(Top:=1.27, Left:=2.9, Width:=2.49, Height:=0.25), _
        parentName:="pgFinancials") ' Right=5.39, Bottom=1.52

    Call gui_buildControl(FORM_NAME, "lblSumCBM", acLabel, _
        gui_buildControlSizeDef(Top:=1.06, Left:=5.59, Width:=2.49, Height:=0.19), _
        parentName:="pgFinancials") ' Right=8.08, Bottom=1.25
    Call gui_buildControl(FORM_NAME, "txtSumCBM", acTextBox, _
        gui_buildControlSizeDef(Top:=1.27, Left:=5.59, Width:=2.49, Height:=0.25), _
        parentName:="pgFinancials") ' Right=8.08, Bottom=1.52

    Call gui_buildControl(FORM_NAME, "lblSumCartons", acLabel, _
        gui_buildControlSizeDef(Top:=1.06, Left:=8.29, Width:=2.49, Height:=0.19), _
        parentName:="pgFinancials") ' Right=10.78, Bottom=1.25
    Call gui_buildControl(FORM_NAME, "txtSumCartons", acTextBox, _
        gui_buildControlSizeDef(Top:=1.27, Left:=8.29, Width:=2.49, Height:=0.25), _
        parentName:="pgFinancials") ' Right=10.78, Bottom=1.52

    Call gui_buildControl(FORM_NAME, "lblSecRemarks", acLabel, _
        gui_buildControlSizeDef(Top:=1.79, Left:=0.2, Width:=10.58, Height:=0.19), _
        parentName:="pgFinancials") ' Right=10.78, Bottom=1.98
    Call gui_buildControl(FORM_NAME, "txtRemarks", acTextBox, _
        gui_buildControlSizeDef(Top:=2#, Left:=0.2, Width:=10.58, Height:=2#), _
        parentName:="pgFinancials") ' Right=10.78, Bottom=4.00

    ' [FINALIZE]: Save & Release
    DoCmd.Close acForm, FORM_NAME, acSaveYes
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "BuildForm_frmOrdersTabbed", Err.Description
End Sub
