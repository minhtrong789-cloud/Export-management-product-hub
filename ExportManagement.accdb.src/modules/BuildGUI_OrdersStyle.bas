Attribute VB_Name = "BuildGUI_OrdersStyle"
Option Compare Database
Option Explicit

'====================================================================
' MODULE    : BuildGUI_OrdersStyle
' PURPOSE   : Styling Enforcement for Orders Hub (frmOrdersHub_Advanced).
'====================================================================

Public Sub EnforceAesthetics_OrderHub()
    On Error GoTo ErrHandler
    Const FORM_NAME As String = "frmOrdersHub_Advanced"

    ' 1. SHELL ENFORCEMENT
    Call EnforceAesthetics_frmOrdersHub_Advanced(FORM_NAME)

    ' 2. GLOBAL HEADER
    Call EnforceAesthetics_boxHeaderBg(FORM_NAME, "boxHeaderBg")
    Call EnforceAesthetics_lblTitle(FORM_NAME, "lblTitle")
    Call EnforceAesthetics_txtSearch(FORM_NAME, "txtSearch")

    ' 3. FILTER TOOLBAR
    Call EnforceAesthetics_cboCustomer(FORM_NAME, "cboCustomer")
    Call EnforceAesthetics_cboStatus(FORM_NAME, "cboStatus")
    Call EnforceAesthetics_txtDateFrom(FORM_NAME, "txtDateFrom")
    Call EnforceAesthetics_txtDateTo(FORM_NAME, "txtDateTo")
    Call EnforceAesthetics_btnClearFilter(FORM_NAME, "btnClearFilter")

    ' 4. KPI BUTTONS
    Call EnforceAesthetics_btnKPIDraft(FORM_NAME, "btnKPIDraft")
    Call EnforceAesthetics_btnKPIShipped(FORM_NAME, "btnKPIShipped")
    Call EnforceAesthetics_btnKPIAlert(FORM_NAME, "btnKPIAlert")

    ' 5. VIEWPORT & GRID
    Call EnforceAesthetics_linMainDivider(FORM_NAME, "linMainDivider")
    Call EnforceAesthetics_subOrderList(FORM_NAME, "subOrderList")

    ' 6. ACTION BUTTONS
    Call EnforceAesthetics_btnNewOrder(FORM_NAME, "btnNewOrder")
    Call EnforceAesthetics_btnExport(FORM_NAME, "btnExport")
    Call EnforceAesthetics_btnPrintInvoice(FORM_NAME, "btnPrintInvoice")
    Call EnforceAesthetics_btnPackingList(FORM_NAME, "btnPackingList")
    Call EnforceAesthetics_btnRefresh(FORM_NAME, "btnRefresh")

    ' 7. SUMMARY FOOTER
    Call EnforceAesthetics_linFooterDivider(FORM_NAME, "linFooterDivider")
    Call EnforceAesthetics_lblRecordCount(FORM_NAME, "lblRecordCount")
    Call EnforceAesthetics_lblTotalAmount(FORM_NAME, "lblTotalAmount")
    Call EnforceAesthetics_txtSumAmount(FORM_NAME, "txtSumAmount")
    Call EnforceAesthetics_togTrashbin(FORM_NAME, "togTrashbin")

    DoCmd.Close acForm, FORM_NAME, acSaveYes
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] BuildGUI_OrdersStyle.EnforceAesthetics_OrderHub: #" & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "EnforceAesthetics_OrderHub", Err.Description
End Sub

'--------------------------------------------------------------------
' FORM SHELL ENFORCER
'--------------------------------------------------------------------
Public Sub EnforceAesthetics_frmOrdersHub_Advanced(ByVal formName As String)
    On Error Resume Next
    Call gui_styleForm(formName)
    With Forms(formName)
        .Caption = "Advanced Orders Hub"
        .NavigationCaption = "Orders Explorer"
        .MinMaxButtons = MinMaxButtons_Both
    End With
End Sub

'--------------------------------------------------------------------
' ATOMIC CONTROL ENFORCERS (HEADER & STRUCTURAL)
'--------------------------------------------------------------------
Public Sub EnforceAesthetics_boxHeaderBg(ByVal formName As String, ByVal ControlName As String)
    On Error Resume Next
    Call gui_styleRectangle_BrandBar(formName, ControlName)
    Forms(formName).Controls(ControlName).BackColor = HexToRBGLong("#2F3C7E")
End Sub

Public Sub EnforceAesthetics_lblTitle(ByVal formName As String, ByVal ControlName As String)
    On Error Resume Next
    Call gui_styleLabel_Title(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .Caption = "ORDER HUB"
        .fontSize = 16
        .FontWeight = FontWeight_Bold
        .ForeColor = vbWhite
    End With
End Sub

Public Sub EnforceAesthetics_txtSearch(ByVal formName As String, ByVal ControlName As String)
    On Error Resume Next
    Call gui_styleTextBox_Magnifying(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .fontSize = 11
        .ForeColor = HexToRBGLong("#2F3C7E")
        .BorderStyle = BorderStyle_Transparent
        .TopMargin = 30
    End With
End Sub

'--------------------------------------------------------------------
' ATOMIC CONTROL ENFORCERS (FILTER TOOLBAR)
'--------------------------------------------------------------------
Public Sub EnforceAesthetics_cboCustomer(ByVal formName As String, ByVal ControlName As String)
    On Error Resume Next
    Call gui_styleComboBox_Standard(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .ColumnCount = 2
        '.ColumnWidths = "3";0""
        .ListWidth = "3"""
        .ForeColor = HexToRBGLong("#2F3C7E")
    End With
End Sub

Public Sub EnforceAesthetics_cboStatus(ByVal formName As String, ByVal ControlName As String)
    On Error Resume Next
    Call gui_styleComboBox_Standard(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .ColumnCount = 2
        '.ColumnWidths = "3";0""
        .ListWidth = "Auto"
        .ForeColor = HexToRBGLong("#2F3C7E")
    End With
End Sub

Public Sub EnforceAesthetics_txtDateFrom(ByVal formName As String, ByVal ControlName As String)
    On Error Resume Next
    Call gui_styleTextBox_Base(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .TextAlign = TextAlign_Center
        .ShowDatePicker = 1
    End With
End Sub

Public Sub EnforceAesthetics_txtDateTo(ByVal formName As String, ByVal ControlName As String)
    On Error Resume Next
    Call gui_styleTextBox_Base(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .TextAlign = TextAlign_Center
        .ShowDatePicker = 1
    End With
End Sub

Public Sub EnforceAesthetics_btnClearFilter(ByVal formName As String, ByVal ControlName As String)
    On Error Resume Next
    Call gui_styleButton_Ghost(formName, ControlName, "#C62828", "#FFEBEE")
    With Forms(formName).Controls(ControlName)
        .Caption = "RESET"
        .fontSize = 8
        .Shape = ButtonShape_Rounded
    End With
End Sub

'--------------------------------------------------------------------
' ATOMIC CONTROL ENFORCERS (KPIS)
'--------------------------------------------------------------------
Public Sub EnforceAesthetics_btnKPIDraft(ByVal formName As String, ByVal ControlName As String)
    On Error Resume Next
    Call gui_styleButton_Ghost(formName, ControlName, "#2F3C7E", "#E8EAF6")
    With Forms(formName).Controls(ControlName)
        .Caption = "DRAFTS: 12"
        .fontSize = 8
        .FontWeight = FontWeight_Bold
        .Shape = ButtonShape_Rounded
    End With
End Sub

Public Sub EnforceAesthetics_btnKPIShipped(ByVal formName As String, ByVal ControlName As String)
    On Error Resume Next
    Call gui_styleButton_Ghost(formName, ControlName, "#2E7D32", "#C8E6C9")
    With Forms(formName).Controls(ControlName)
        .Caption = "SHIPPED: 85"
        .fontSize = 8
        .FontWeight = FontWeight_Bold
        .BackColor = HexToRBGLong("#E8F5E9")
        .Shape = ButtonShape_Rounded
    End With
End Sub

Public Sub EnforceAesthetics_btnKPIAlert(ByVal formName As String, ByVal ControlName As String)
    On Error Resume Next
    Call gui_styleButton_Ghost(formName, ControlName, "#C62828", "#FFEBEE")
    With Forms(formName).Controls(ControlName)
        .Caption = "ALERTS: 3"
        .fontSize = 8
        .FontWeight = FontWeight_Bold
        .BackColor = HexToRBGLong("#FBEBEB")
        .Shape = ButtonShape_Rounded
    End With
End Sub

'--------------------------------------------------------------------
' ATOMIC CONTROL ENFORCERS (GRID & BLADE)
'--------------------------------------------------------------------
Public Sub EnforceAesthetics_linMainDivider(ByVal formName As String, ByVal ControlName As String)
    On Error Resume Next
    Call gui_styleLine_Divider(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .BorderWidth = 2
        .BorderColor = HexToRBGLong("#C7CCEA")
        .HorizontalAnchor = acHorizontalAnchorBoth
        .SpecialEffect = 0
    End With
End Sub

Public Sub EnforceAesthetics_subOrderList(ByVal formName As String, ByVal ControlName As String)
    On Error Resume Next
    Call gui_styleSubform_Theater(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .BorderWidth = 1
        .BorderColor = HexToRBGLong("#C7CCEA")
        .HorizontalAnchor = acHorizontalAnchorBoth
        .VerticalAnchor = acVerticalAnchorBoth
    End With
End Sub

Public Sub EnforceAesthetics_btnNewOrder(ByVal formName As String, ByVal ControlName As String)
    On Error Resume Next
    Call gui_styleButton_Power(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .Caption = "NEW ORDER"
        .HorizontalAnchor = acHorizontalAnchorRight
        .Shape = ButtonShape_Rounded
    End With
End Sub

Public Sub EnforceAesthetics_btnExport(ByVal formName As String, ByVal ControlName As String)
    On Error Resume Next
    Call gui_styleButton_Ghost(formName, ControlName, "#1B5E20", "#E8F5E9")
    With Forms(formName).Controls(ControlName)
        .Caption = "EXPORT"
        .HorizontalAnchor = acHorizontalAnchorRight
        .Shape = ButtonShape_Rounded
    End With
End Sub

Public Sub EnforceAesthetics_btnPrintInvoice(ByVal formName As String, ByVal ControlName As String)
    On Error Resume Next
    Call gui_styleButton_Ghost(formName, ControlName, "#2F3C7E", "#E8EAF6")
    With Forms(formName).Controls(ControlName)
        .Caption = "PRINT INVOICE"
        .HorizontalAnchor = acHorizontalAnchorRight
        .Shape = ButtonShape_Rounded
    End With
End Sub

Public Sub EnforceAesthetics_btnPackingList(ByVal formName As String, ByVal ControlName As String)
    On Error Resume Next
    Call gui_styleButton_Ghost(formName, ControlName, "#2F3C7E", "#E8EAF6")
    With Forms(formName).Controls(ControlName)
        .Caption = "PACKING LIST"
        .HorizontalAnchor = acHorizontalAnchorRight
        .Shape = ButtonShape_Rounded
    End With
End Sub

Public Sub EnforceAesthetics_btnRefresh(ByVal formName As String, ByVal ControlName As String)
    On Error Resume Next
    Call gui_styleButton_Ghost(formName, ControlName, "#757575", "#F5F5F5")
    With Forms(formName).Controls(ControlName)
        .Caption = "REFRESH"
        .HorizontalAnchor = acHorizontalAnchorRight
        .Shape = ButtonShape_Rounded
    End With
End Sub

'--------------------------------------------------------------------
' ATOMIC CONTROL ENFORCERS (FOOTER)
'--------------------------------------------------------------------
Public Sub EnforceAesthetics_linFooterDivider(ByVal formName As String, ByVal ControlName As String)
    On Error Resume Next
    Call gui_styleLine_Divider(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .BorderWidth = 1
        .BorderColor = HexToRBGLong("#E5E7EB")
        .HorizontalAnchor = acHorizontalAnchorBoth
        .VerticalAnchor = acVerticalAnchorBottom
    End With
End Sub

Public Sub EnforceAesthetics_lblRecordCount(ByVal formName As String, ByVal ControlName As String)
    On Error Resume Next
    Call gui_styleLabel_Metadata(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .Caption = "Showing: 0 / 0 orders"
        .ForeColor = HexToRBGLong("#757575")
        .VerticalAnchor = acVerticalAnchorBottom
    End With
End Sub

Public Sub EnforceAesthetics_lblTotalAmount(ByVal formName As String, ByVal ControlName As String)
    On Error Resume Next
    Call gui_styleLabel_Metadata(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .Caption = "Total Displayed:"
        .ForeColor = HexToRBGLong("#757575")
        .VerticalAnchor = acVerticalAnchorBottom
        .HorizontalAnchor = acHorizontalAnchorRight
        .TextAlign = TextAlign_Right
    End With
End Sub

Public Sub EnforceAesthetics_txtSumAmount(ByVal formName As String, ByVal ControlName As String)
    On Error Resume Next
    Call gui_styleTextBox_Value(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .fontSize = 12
        .FontWeight = FontWeight_Bold
        .ForeColor = HexToRBGLong("#2F3C7E")
        .HorizontalAnchor = acHorizontalAnchorRight
        .VerticalAnchor = acVerticalAnchorBottom
        .TextAlign = TextAlign_Right
        .BackStyle = BackStyle_Transparent
        .BorderStyle = BorderStyle_Transparent
    End With
End Sub

Public Sub EnforceAesthetics_togTrashbin(ByVal formName As String, ByVal ControlName As String)
    On Error Resume Next
    Call gui_styleToggleButton_Sentinel(formName, ControlName)
    With Forms(formName).Controls(ControlName)
        .Caption = "ARCHIVE"
        .HorizontalAnchor = acHorizontalAnchorRight
        .VerticalAnchor = acVerticalAnchorBottom
    End With
End Sub
