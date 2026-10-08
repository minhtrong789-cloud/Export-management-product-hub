Attribute VB_Name = "ExcelRender_Orders_CommercialInvoice"
Option Explicit
'@Folder "App.Products"
'805.89 points
'usable point: 795 points
'Header: 78.50 points
'Buyer: 195.00 points
'Footer: 213.50 points
'Fixed Load: 487 points
'Remaining for product: 318.89 points

'Total Page Height: 750'

' ====================================================================
' STATE VARIABLES (Excel Engine)
' ====================================================================
Private m_xlApp As Object
Private m_xlBook As Object
Private m_xlSheet As Object
Private m_pinGrandTotalRow As Object




' GLOBAL GEOMETRY CONSTANTS
' ====================================================================
Private Const LINE_UNIT As Double = 10.15           ' Standard 1-line text footprint
Private Const SPACER_07 As Double = 5               ' Strict 0.07" padding height
Private Const SPAN_COLS As Integer = 9              ' The 9-column invoice width
Private Const LAST_CELL_POSITION As Double = 750#       ' Absolute Physical Limit
Private Const DESC_ZONE_WIDTH As Double = 245#      ' Physical width of C:E span (245 pts)
Private Const MSG_FOC As String = "*** FREE OF CHARGE ITEMS (F.O.C.) ***"
' --- CALIBRATED PRODUCT CAPACITIES (Points) ---
Private Const MAX_CAP_SINGLE_PAGE As Double = 234#      ' Header + Big Footer
Private Const MAX_CAP_FIRST_PAGE  As Double = 444#      ' Header + No Footer
Private Const MAX_CAP_MIDDLE_PAGE As Double = 609#      ' Slim Header + No Footer
Private Const MAX_CAP_LAST_PAGE   As Double = 397#      ' Slim Header + Big Footer

' GEOMETRIC COSTS (Displacement Math)
Private Const GEOM_COST_FOC_HEADER As Double = 22#
Private Const GEOM_COST_CAT_SUMMARY As Double = 15#
' ----------------------------------------------


' ====================================================================
' DATA BACKBONE (Module-Level State)
' ====================================================================
' Zone 1: Shipper & Invoice Metadata
Private m_Shipper_Name As String, m_Shipper_Addr1 As String, m_Shipper_Addr2 As String
Private m_Shipper_Web As String, m_Shipper_Email As String, m_Shipper_Tel As String
Private m_Inv_Number As String, m_Inv_Date As Date

' Zone 2: Buyer & Shipping Info
Private m_SoldTo_Name As String, m_SoldTo_Addr1 As String, m_SoldTo_Addr2 As String
Private m_SoldTo_Contact As String, m_SoldTo_Email As String, m_SoldTo_Tel As String
Private m_ShipTo_Name As String, m_ShipTo_Addr1 As String, m_ShipTo_Addr2 As String
Private m_ShipTo_Contact As String, m_ShipTo_Email As String, m_ShipTo_Tel As String

' Zone 3: Logistics & Terms
Private m_Term_Incoterm As String, m_Term_PO As String, m_Term_POL As String
Private m_Term_Ref As String, m_Term_POD As String, m_Term_Payment As String
Private m_Term_Currency As String

' Zone 6: Settlement & Compliance
Private m_Sett_Quality As String, m_Sett_Origin As String
Private m_Sett_AmountInWords As String, m_Sett_PayableAmount As Double
Private m_Sign_ShipperName As String ' Signature line company name
Private m_Sign_BankInfo As String    ' Banking details for sign zone


'================================================================================
' PROCEDURE : RenderCommercialInvoice
' PURPOSE   : High-fidelity Excel-based Invoice Printing Engine.
' STATUS    : STABLE (Production Engine)
'================================================================================
Public Sub RenderCommercialInvoice(ByVal targetOrderID As Long, _
                                         Optional ByVal targetPath As String = "")
    On Error GoTo ErrHandler
    Set m_pinGrandTotalRow = Nothing

    ' [SECTION 1]: INITIALIZE EXCEL ENGINE
    Set m_xlApp = CreateObject("Excel.Application")
    m_xlApp.ScreenUpdating = False
    If targetPath = "" Then m_xlApp.Visible = False ' Show for debugging if no path given
    Set m_xlBook = m_xlApp.Workbooks.Add
    Set m_xlSheet = m_xlBook.Sheets(1)

    ' [SECTION 2]: DATA INITIALIZATION
    Dim ProductData As Variant
    ProductData = GetMock_ProductData()
    Call HydrateModuleMetadata ' Hydrate Module Backbone with Hardcoded Sample Data

    ' [PAGE SET UP & GEOMETRY]
    Call ConfigurePageSetup(m_xlSheet)

    ' 1. INITIALIZE DATA TANKS
    ' --------------------------------------------------------------------
    Dim pItems As New Collection
    Dim iLoop As Long
    For iLoop = 0 To 4
        pItems.Add Array(ProductData(iLoop, 0), ProductData(iLoop, 1), ProductData(iLoop, 2), ProductData(iLoop, 3), ProductData(iLoop, 4), ProductData(iLoop, 5), ProductData(iLoop, 6), ProductData(iLoop, 7), ProductData(iLoop, 8))
    Next iLoop

    ' HARDCODE: SECOND TANK (F.O.C. Items) - MULTIPAGE STRESS TEST
    ' Items must be 1D arrays matching Paint_TableRow parameter order:
    ' (0)=NO, (1)=HSCode, (2)=Description, (3)=Qty, (4)=Unit, (5)=Price, (6)=Barcode, (7)=LotNum, (8)=Expiry
    Dim pFocItems As New Collection
    Dim f As Integer
    For f = 5 To 9
        pFocItems.Add Array(ProductData(f, 0), ProductData(f, 1), ProductData(f, 2), ProductData(f, 3), ProductData(f, 4), ProductData(f, 5), ProductData(f, 6), ProductData(f, 7), ProductData(f, 8))
    Next f

    ' 2. CALCULATE DATA SET HEIGHT (Phantom Measurement Engine)
    ' --------------------------------------------------------------------
    Dim totalProductLoad As Double: totalProductLoad = 0
    Dim itemHeight_Paid As Double: itemHeight_Paid = 0
    Dim itemHeight_Foc As Double: itemHeight_Foc = 0
    Dim paidStackHeight As Double: paidStackHeight = 0
    Dim focStackHeight As Double: focStackHeight = 0

    If pItems.Count = 0 And pFocItems.Count = 0 Then
        totalProductLoad = 0 ' No products, fits on Page 1
    Else
        ' PRE-CALCULATION GEOMETRY: Create Burner Sheet (Deep Clone)
        m_xlSheet.Copy After:=m_xlSheet
        Dim pSheet As Object: Set pSheet = m_xlBook.Sheets(2)

        ' Run the REAL painter on the Mock sheet to get the mathematical TRUTH for BOTH tanks
        Dim pDrawingCell As Object, pStackRange As Object, focStackRange As Object

        Set pDrawingCell = pSheet.Range("A1")

        ' Build PHANTOM copies to protect originals from the Eater Engine
        Dim pPhantomPaid As New Collection, pPhantomFoc As New Collection
        Dim ph As Long
        For ph = 1 To pItems.Count:    pPhantomPaid.Add pItems(ph): Next ph
        For ph = 1 To pFocItems.Count: pPhantomFoc.Add pFocItems(ph): Next ph

        ' Measure Tank 1 (Paid) using Phantom copy
        If pPhantomPaid.Count > 0 Then
            Set pStackRange = Render_ProductStack(pDrawingCell, pPhantomPaid, 0, True)
            paidStackHeight = pStackRange.Height
            itemHeight_Paid = pStackRange.Rows("1:2").Height
            Set pDrawingCell = pStackRange.Offset(pStackRange.Rows.Count, 0).Cells(1, 1)
        End If

        ' Measure Tank 2 (F.O.C.) using Phantom copy
        If pPhantomFoc.Count > 0 Then
            Set focStackRange = Render_ProductStack(pDrawingCell, pPhantomFoc, 0, True)
            focStackHeight = focStackRange.Height
            itemHeight_Foc = focStackRange.Rows("1:2").Height
        End If

        totalProductLoad = paidStackHeight + focStackHeight


        ' THE MASTER MEMORY: Segmented Tanks for Paid vs F.O.C. synthesis
        Dim pPaidZones As New Collection, pFocZones As New Collection
        Dim targetZones As Collection

        ' 2. DEFINE THE SCENARIO (The "Brain" part)
        Dim hasPaid As Boolean:   hasPaid = (pItems.Count > 0)
        Dim hasFoc As Boolean:    hasFoc = (pFocItems.Count > 0)
        Dim isDualMode As Boolean: isDualMode = (hasPaid And hasFoc)

        Dim LabelFOCHeight As Double: LabelFOCHeight = 0
        If hasFoc Then LabelFOCHeight = GEOM_COST_FOC_HEADER

        Dim PrintMode As String
        Dim dataObject As Object
        Dim focLabel As Boolean: focLabel = False

        ' Assign the single-tank data object correctly.
        ' In Dual-Mode, this is irrelevant as each stage uses its own collection directly.
        If hasPaid And Not isDualMode Then
            Set dataObject = pItems
        ElseIf hasFoc And Not isDualMode Then
            focLabel = True
            Set dataObject = pFocItems
        End If

        ' 3. THE HIGH-LEVEL DECISION
        If isDualMode Then
            ' IF Dual, we FORCE Multi-Page mode (Your "Sectional" rule)
            PrintMode = "Multi-Page"
        Else
            PrintMode = IIf(totalProductLoad > (MAX_CAP_SINGLE_PAGE - LabelFOCHeight), "Multi-Page", "Single-Page")
            ' If single-tank, point the target dispatcher at the correct bucket once
            If hasPaid Then Set targetZones = pPaidZones Else Set targetZones = pFocZones
        End If

        Dim currentCursor As Object, NumberOfProducts As Long
        Dim remainingPaidHeight As Double, remainingFocHeight As Double
        Dim capFirst_Summary As Double, capMiddle_Summary As Double
        Dim capLast_FOC As Double, capMiddle_FOC As Double

        ' UI ANCHORS: GPS Pins for Synthesis
        Dim pinPaidSummary As Object, pinFocSummary As Object
        Dim SummaryPins As New Collection

        Select Case True
            Case PrintMode = "Single-Page"
                Dim selectedH As Double: selectedH = IIf(hasPaid, itemHeight_Paid, itemHeight_Foc)
                Call Layout_SinglePage(m_xlSheet.Range("A1"), dataObject, selectedH, targetZones, focLabel, pinFocSummary)
                If Not pinFocSummary Is Nothing Then
                    SummaryPins.Add pinFocSummary
                    Call BindCategorySummary(pinFocSummary, targetZones)
                End If

            Case PrintMode = "Multi-Page" And Not isDualMode
                Dim activeH As Double: activeH = IIf(hasPaid, itemHeight_Paid, itemHeight_Foc)

                ' Kick off page 1 at A1
                NumberOfProducts = Int((MAX_CAP_FIRST_PAGE - LabelFOCHeight) / activeH)
                Set currentCursor = Layout_FirstPage(m_xlSheet.Range("A1"), dataObject, NumberOfProducts, activeH, targetZones, focLabel, False)

                Do While ((dataObject.Count * activeH) / MAX_CAP_LAST_PAGE) > 1
                    ' Advance pen to next free row and draw Middle Page
                    NumberOfProducts = Int((MAX_CAP_MIDDLE_PAGE - LabelFOCHeight) / activeH)
                    Set currentCursor = currentCursor.Offset(currentCursor.Rows.Count + 1, 0).Cells(1, 1)
                    m_xlSheet.HPageBreaks.Add Before:=currentCursor
                    Set currentCursor = Layout_MiddlePage(currentCursor, dataObject, NumberOfProducts, activeH, targetZones, focLabel, False)
                Loop

                ' Advance pen again and draw Final Page
                NumberOfProducts = Int((MAX_CAP_LAST_PAGE - LabelFOCHeight) / activeH)
                Set currentCursor = currentCursor.Offset(currentCursor.Rows.Count + 1, 0).Cells(1, 1)
                m_xlSheet.HPageBreaks.Add Before:=currentCursor
                Set currentCursor = Layout_LastPage(currentCursor, dataObject, NumberOfProducts, activeH, targetZones, focLabel, pinFocSummary)
                If Not pinFocSummary Is Nothing Then
                    SummaryPins.Add pinFocSummary
                    Call BindCategorySummary(pinFocSummary, targetZones)
                End If

            Case PrintMode = "Multi-Page" And isDualMode
                ' ==========================================
                ' STAGE 1: COMMERCIAL DATA (pItems)
                ' ==========================================
                remainingPaidHeight = paidStackHeight ' Use Surveyor's Mathematical Truth

                ' Calculate capacities with the 15-pt Summary penalty
                capFirst_Summary = MAX_CAP_FIRST_PAGE - GEOM_COST_CAT_SUMMARY
                capMiddle_Summary = MAX_CAP_MIDDLE_PAGE - GEOM_COST_CAT_SUMMARY

                If remainingPaidHeight <= capFirst_Summary Then
                    ' Fits entirely on First Page. Trigger Summary!
                    NumberOfProducts = pItems.Count
                    Set currentCursor = Layout_FirstPage(m_xlSheet.Range("A1"), pItems, NumberOfProducts, itemHeight_Paid, pPaidZones, False, True, pinPaidSummary)
                    If Not pinPaidSummary Is Nothing Then SummaryPins.Add pinPaidSummary
                Else
                    ' Overflows. Standard First Page (No summary yet).
                    NumberOfProducts = Int(MAX_CAP_FIRST_PAGE / itemHeight_Paid)
                    Set currentCursor = Layout_FirstPage(m_xlSheet.Range("A1"), pItems, NumberOfProducts, itemHeight_Paid, pPaidZones, False, False)
                    remainingPaidHeight = remainingPaidHeight - (NumberOfProducts * itemHeight_Paid)

                    ' Loop Middle Pages for Paid Data
                    Do While remainingPaidHeight > capMiddle_Summary
                        NumberOfProducts = Int(MAX_CAP_MIDDLE_PAGE / itemHeight_Paid)
                        Set currentCursor = currentCursor.Offset(currentCursor.Rows.Count + 1, 0).Cells(1, 1)
                        m_xlSheet.HPageBreaks.Add Before:=currentCursor
                        Set currentCursor = Layout_MiddlePage(currentCursor, pItems, NumberOfProducts, itemHeight_Paid, pPaidZones, False, False)
                        remainingPaidHeight = remainingPaidHeight - (NumberOfProducts * itemHeight_Paid)
                    Loop

                    ' Final tail of Paid data on a Middle Page WITH Summary
                    If remainingPaidHeight > 0 Then
                        NumberOfProducts = Int(remainingPaidHeight / itemHeight_Paid)
                        Set currentCursor = currentCursor.Offset(currentCursor.Rows.Count + 1, 0).Cells(1, 1)
                        m_xlSheet.HPageBreaks.Add Before:=currentCursor
                        Set currentCursor = Layout_MiddlePage(currentCursor, pItems, NumberOfProducts, itemHeight_Paid, pPaidZones, False, True, pinPaidSummary)
                        If Not pinPaidSummary Is Nothing Then SummaryPins.Add pinPaidSummary
                    End If
                End If

                ' BINDING: Synthesis for Paid Items
                Call BindCategorySummary(pinPaidSummary, pPaidZones)

                ' ==========================================
                ' STAGE 2: F.O.C. DATA (pFocItems)
                ' ==========================================
                remainingFocHeight = focStackHeight ' Use Surveyor's Mathematical Truth

                ' Calculate capacities with the 20-pt FOC Header penalty
                capLast_FOC = MAX_CAP_LAST_PAGE - GEOM_COST_FOC_HEADER
                capMiddle_FOC = MAX_CAP_MIDDLE_PAGE - GEOM_COST_FOC_HEADER

                If remainingFocHeight <= capLast_FOC Then
                    ' Fits entirely on Last Page (The Tail of the document)
                    NumberOfProducts = pFocItems.Count
                    Set currentCursor = currentCursor.Offset(currentCursor.Rows.Count + 1, 0).Cells(1, 1)
                    m_xlSheet.HPageBreaks.Add Before:=currentCursor
                    Set currentCursor = Layout_LastPage(currentCursor, pFocItems, NumberOfProducts, itemHeight_Foc, pFocZones, True, pinFocSummary)
                    If Not pinFocSummary Is Nothing Then SummaryPins.Add pinFocSummary
                Else
                    ' Overflows. Loop Middle Pages first
                    Do While remainingFocHeight > capLast_FOC
                        NumberOfProducts = Int(capMiddle_FOC / itemHeight_Foc)
                        Set currentCursor = currentCursor.Offset(currentCursor.Rows.Count + 1, 0).Cells(1, 1)
                        m_xlSheet.HPageBreaks.Add Before:=currentCursor
                        Set currentCursor = Layout_MiddlePage(currentCursor, pFocItems, NumberOfProducts, itemHeight_Foc, pFocZones, True, False)
                        remainingFocHeight = remainingFocHeight - (NumberOfProducts * itemHeight_Foc)
                    Loop

                    ' Final piece goes on Last Page
                    If remainingFocHeight > 0 Then
                        NumberOfProducts = Int(remainingFocHeight / itemHeight_Foc)
                        Set currentCursor = currentCursor.Offset(currentCursor.Rows.Count + 1, 0).Cells(1, 1)
                        m_xlSheet.HPageBreaks.Add Before:=currentCursor
                        Set currentCursor = Layout_LastPage(currentCursor, pFocItems, NumberOfProducts, itemHeight_Foc, pFocZones, True, pinFocSummary)
                        If Not pinFocSummary Is Nothing Then SummaryPins.Add pinFocSummary
                    End If
                End If

                ' BINDING: Synthesis for F.O.C. Items
                Call BindCategorySummary(pinFocSummary, pFocZones)

        End Select

        ' [STAGE 7]: FINAL SYNTHESIS (Mathematical Glue)
        ' DIRECT SOURCE ANCHORING: We sum product blocks directly from the tanks.
        If Not m_pinGrandTotalRow Is Nothing Then
            Dim zone As Variant, gQty As String, gAmt As String
            gQty = "=SUM(": gAmt = "=SUM("

            ' A. QUANTITY SYNTHESIS (Paid + F.O.C. = Total Physical Volume)
            For Each zone In pPaidZones
                gQty = gQty & zone.Columns(6).Address(True, True) & ","
            Next
            For Each zone In pFocZones
                gQty = gQty & zone.Columns(6).Address(True, True) & ","
            Next

            ' B. AMOUNT SYNTHESIS (Purely Commercial Values)
            For Each zone In pPaidZones
                gAmt = gAmt & zone.Columns(9).Address(True, True) & ","
            Next

            ' Finalize Formulas (Cleanup terminal commas)
            If Len(gQty) > 5 Then gQty = Left(gQty, Len(gQty) - 1) & ")" Else gQty = "=0"
            If Len(gAmt) > 5 Then gAmt = Left(gAmt, Len(gAmt) - 1) & ")" Else gAmt = "=0"

            ' EXCLUSIVE INJECTION
            m_pinGrandTotalRow.Cells(1, 6).Formula = gQty
            m_pinGrandTotalRow.Cells(1, 9).Formula = gAmt
        End If

        ' Cleanup Burner Evidence
        m_xlApp.DisplayAlerts = False
        pSheet.Delete
        m_xlApp.DisplayAlerts = True
    End If


    ' [SECTION 6]: SHIP TO PRODUCTION (Save & Deliver)
    ' 6A. Restore Screen Output
    m_xlApp.ScreenUpdating = True
    m_xlApp.Visible = True

    ' 6B. Final Delivery (Save if path provided)
    If targetPath <> "" Then
        On Error Resume Next: Kill targetPath: On Error GoTo ErrHandler
        m_xlBook.SaveAs targetPath
        m_xlBook.Close
        m_xlApp.Quit
    End If

    ' 6C. Cleanup
    Set m_xlSheet = Nothing: Set m_xlBook = Nothing: Set m_xlApp = Nothing
    Exit Sub

ErrHandler:
    If Not m_xlApp Is Nothing Then m_xlApp.ScreenUpdating = True: m_xlApp.Visible = True
    Debug.Print "[ERROR] ExcelRender_Orders_CommercialInvoice: #" & Err.Number & " - " & Err.Description
    MsgBox "Invoice Production Failed: " & Err.Description, vbCritical
End Sub




' ====================================================================
' DATA HYDRATION (The Bridge)
' ====================================================================
' PURPOSE: Maps a database recordset to the Module Backbone.
' --------------------------------------------------------------------
' PURPOSE: Loads Mock Data for engine validation (Hardcoded for now).
' --------------------------------------------------------------------
Public Sub HydrateModuleMetadata(Optional ByRef rs As Object)
    ' Zone 1: Shipper
    m_Shipper_Name = "Y&B CO., LTD"
    m_Shipper_Addr1 = "No 1, Ly Tu Trong Str, Ben Nghe Ward"
    m_Shipper_Addr2 = "District 1, Ho Chi Minh City, Vietnam"
    m_Shipper_Web = "www.yandb-trading.vn"
    m_Shipper_Email = "info@yandb-trading.vn"
    m_Shipper_Tel = "+84 28 3823 4567"

    ' Zone 1: Invoice Header
    m_Inv_Number = "INV-2026-POC-001"
    m_Inv_Date = Date

    ' Zone 2: Buyer & Shipping
    m_SoldTo_Name = "GLOBAL LOGISTICS IMPORT EXPORT CORP"
    m_SoldTo_Addr1 = "77 Robinson Road, #13-00 Robinson 77"
    m_SoldTo_Addr2 = "Singapore 068896"
    m_SoldTo_Contact = "Mr. James Bond"
    m_SoldTo_Email = "j.bond@global-logistics.sg"
    m_SoldTo_Tel = "+65 6789 0123"

    m_ShipTo_Name = "PORT OF SINGAPORE WAREHOUSE 5"
    m_ShipTo_Addr1 = "10 Pasir Panjang Road"
    m_ShipTo_Addr2 = "MapleTree Business City, Singapore"
    m_ShipTo_Contact = "Receiving Dept"
    m_ShipTo_Email = "ops@psa.sg"
    m_ShipTo_Tel = "+65 6123 4567"

    ' Zone 3: Logistics
    m_Term_Incoterm = "FOB HO CHI MINH CITY (INCOTERMS 2020)"
    m_Term_PO = "PO-2026-XYZ-789"
    m_Term_POL = "CAT LAI PORT, VIETNAM"
    m_Term_Ref = "S/O NO. 99887766"
    m_Term_POD = "PORT OF SINGAPORE"
    m_Term_Payment = "T/T 100% IN ADVANCE"
    m_Term_Currency = "USD"

    ' Zone 6: Totals & Signature
    m_Sett_Quality = "AS PER CONTRACT NO. YB/2026/001"
    m_Sett_Origin = "VIETNAM"
    m_Sign_ShipperName = "Y&B CO., LTD"
    m_Sign_BankInfo = "BENEFICIARY: Y&B JOINT STOCK COMPANY" & vbCrLf & _
                      "BANK: JOINT STOCK COMMERCIAL BANK FOR FOREIGN TRADE OF VIETNAM - TAN DINH BRANCH" & vbCrLf & _
                      "ADD: 72 PHAM NGOC THACH STREET, XUAN HOA WARD, HO CHI MINH CITY" & vbCrLf & _
                      "ACC NO: 1039569999 (USD)"

    ' Math Placeholders
    m_Sett_PayableAmount = 0 ' Will be updated by summary math
    m_Sett_AmountInWords = "ZERO DOLLARS ONLY"
End Sub

' ====================================================================
' THE DATA SCISSOR (Modern Collection Engine)
' ====================================================================
' PURPOSE: Turns a 2D array into a dynamic Collection for the Eater Engine.
' --------------------------------------------------------------------
Public Function MapArrayToCollection(ByVal ProductData As Variant) As Collection
    Dim col As New Collection
    Dim i As Long, c As Long
    Dim itemData As Variant

    If Not IsArray(ProductData) Then
        Set MapArrayToCollection = col
        Exit Function
    End If

    For i = LBound(ProductData, 1) To UBound(ProductData, 1)
        ' Extract single item's data (9 columns) into a 1D array
        ReDim itemData(LBound(ProductData, 2) To UBound(ProductData, 2))
        For c = LBound(ProductData, 2) To UBound(ProductData, 2)
            itemData(c) = ProductData(i, c)
        Next c
        col.Add itemData
    Next i

    Set MapArrayToCollection = col
End Function



'================================================================================
' CORE SETUP: ConfigurePageSetup
' Purpose: Applies template-specific Excel Document properties, orientation, fonts,
'          and the Platinum Column Blueprint.
'================================================================================
Private Sub ConfigurePageSetup(ByVal targetSheet As Object)
    With targetSheet
        .name = "COMMERCIAL INVOICE"

        With .Cells.Font
            .name = "Arial": .size = 8
        End With

        With .PageSetup
            'PORTRAIT A4
            .PaperSize = 9 ' xlPaperA4
            .Orientation = 1 ' xlPortrait
            .Zoom = 100

            ' Precise Asymmetrical Margins (Extracted from Template)
            .LeftMargin = m_xlApp.InchesToPoints(0.4)    ' 0.40"
            .RightMargin = m_xlApp.InchesToPoints(0.3)   ' 0.30"
            .TopMargin = m_xlApp.InchesToPoints(0.2)     ' 0.20"
            .BottomMargin = m_xlApp.InchesToPoints(0.3)  ' 0.30"
            .HeaderMargin = 0
            .FooterMargin = m_xlApp.InchesToPoints(0.1)
            .RightFooter = "Page &P of &N"
            .CenterHorizontally = True
        End With

    ' 2. COLUMN GEOMETRY (The Platinum Grid Blueprint)
    Dim colBlueprint(1 To 9, 1 To 2) As Variant
    colBlueprint(1, 1) = "A": colBlueprint(1, 2) = 0.2917 ' NO. (21 pts)
    colBlueprint(2, 1) = "B": colBlueprint(2, 2) = 0.7292 ' HS CODE (52.5 pts)
    colBlueprint(3, 1) = "C": colBlueprint(3, 2) = 0.4306 ' SOLD TO (31 pts)
    colBlueprint(4, 1) = "D": colBlueprint(4, 2) = 2.5833 ' DESCRIPTION (186 pts)
    colBlueprint(5, 1) = "E": colBlueprint(5, 2) = 0.3889 ' UNIT BUFFER (28 pts)
    colBlueprint(6, 1) = "F": colBlueprint(6, 2) = 0.5833 ' QTY (42 pts)
    colBlueprint(7, 1) = "G": colBlueprint(7, 2) = 0.5833 ' UNIT (42 pts)
    colBlueprint(8, 1) = "H": colBlueprint(8, 2) = 0.8264 ' PRICE (59.5 pts)
    colBlueprint(9, 1) = "I": colBlueprint(9, 2) = 0.9444 ' AMOUNT (68 pts)

    Dim i As Integer
    For i = 1 To 9
        Call Precision_SetColWidth(targetSheet, CStr(colBlueprint(i, 1)), CDbl(colBlueprint(i, 2)))
    Next i

        ' Hide initial clutter
        m_xlApp.ActiveWindow.DisplayGridlines = False
    End With
End Sub





'--------------------------------------------------------------------------------
' LAYOUT: Layout_SinglePage
' PURPOSE: Orchestrates the high-authority single-page layout.
'--------------------------------------------------------------------------------
Private Function Layout_SinglePage(ByVal startCell As Object, _
                                Optional ByRef pItems As Collection, _
                                Optional ByVal ForceTargetHeight As Double = 0, _
                                Optional ByRef outProductRanges As Collection, _
                                Optional ByVal ForceFOCHeader As Boolean = False, _
                                Optional ByRef outSummaryRow As Object) As Object
    Dim HeightTracker As Double
    Dim DrawingCell As Object, DrawedRange As Object

    ' 1. DRAW HEADER
    Set DrawingCell = startCell
    Set DrawedRange = Graphic_RenderSpacerBlock(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Paint_Header(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Graphic_RenderDividerBlock(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Paint_BuyerInfo(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Graphic_RenderDividerBlock(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Paint_Terms(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Graphic_RenderSpacerBlock(DrawingCell)

    ' 2. DRAW PRODUCT TABLE HEADER
    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Paint_TableHeader(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Paint_TableSpacer(DrawingCell)

    If ForceFOCHeader Then
        Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
        Set DrawedRange = Paint_TableFOCHeader(DrawingCell)

        Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
        Set DrawedRange = Paint_TableSpacer(DrawingCell)
    End If

    ' 3. LOOP & DRAW PRODUCTS
    ' --------------------------------------------------------------------
    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Render_ProductStack(DrawingCell, pItems, , , ForceTargetHeight)

    ' LOG THE ZONE: Capture the product range for document-wide intelligence
    If Not outProductRanges Is Nothing Then outProductRanges.Add DrawedRange

    ' 4. THE ELASTIC BUFFER (Double-Armor System)
    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Dim ElasticBufferRange As Object
    Set ElasticBufferRange = Graphic_RenderElasticBuffer(DrawingCell)

    ' THE PRECISION GUARD
    Set DrawingCell = ElasticBufferRange.Offset(ElasticBufferRange.Rows.Count, 0).Cells(1, 1)
    Dim ElasticBufferRange_Guard As Object
    Set ElasticBufferRange_Guard = Graphic_RenderSpacerBlock(DrawingCell)
    ElasticBufferRange_Guard.RowHeight = 0

    ' 5. DRAW FOOTER STACK
    Set DrawingCell = ElasticBufferRange_Guard.Offset(ElasticBufferRange_Guard.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Paint_TableFooter(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set outSummaryRow = Paint_TableSummary(DrawingCell)
    Set DrawedRange = outSummaryRow

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Paint_FinalSettlement(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Graphic_RenderSpacerBlock(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Paint_BankAndSignature(DrawingCell)

    ' FINAL JUSTIFICATION (Geometric Truth Engine)
    ' --------------------------------------------------------------------
    ' Calculate the absolute Physical Truth of the entire painted zone
    HeightTracker = m_xlSheet.Range(startCell, DrawedRange).Height

    If LAST_CELL_POSITION > HeightTracker Then
        Dim totalGap As Double: totalGap = LAST_CELL_POSITION - HeightTracker
        Dim unitH As Double: unitH = Int((totalGap / ElasticBufferRange.Rows.Count) * 0.7)

        ' Distribute conservative buffer
        Dim rIdx As Long
        For rIdx = 1 To ElasticBufferRange.Rows.Count
            ElasticBufferRange.Rows(rIdx).RowHeight = unitH
        Next rIdx

        ' Final Sync before Guard Mopping
        HeightTracker = m_xlSheet.Range(startCell, DrawedRange).Height

        If LAST_CELL_POSITION > HeightTracker Then
            Call Precision_SetRowHeight(m_xlSheet, ElasticBufferRange_Guard.row, (LAST_CELL_POSITION - HeightTracker) / 72)
        End If
    End If

    ' RETURN CURSOR
    Set Layout_SinglePage = DrawedRange
End Function

'--------------------------------------------------------------------------------
' LAYOUT: Layout_FirstPage
' PURPOSE: Orchestrates the first page of a MULTI-PAGE document.
' DESIGN : Full Branding + Full Buyer/Terms + Products. (No Signature).
'--------------------------------------------------------------------------------
Private Function Layout_FirstPage(ByVal startCell As Object, _
                                ByRef pItems As Collection, _
                                ByVal NumberOfProducts As Long, _
                                ByVal ForceTargetHeight As Double, _
                                 Optional ByRef outProductRanges As Collection, _
                                 Optional ByVal ForceFOCHeader As Boolean = False, _
                                 Optional ByVal ForceTableSummary As Boolean = False, _
                                 Optional ByRef outSummaryRow As Object) As Object
    Dim HeightTracker As Double
    Dim DrawingCell As Object, DrawedRange As Object

    ' 1. DRAW HEADER (Z1 + Z2 + Z3)
    Set DrawingCell = startCell
    Set DrawedRange = Graphic_RenderSpacerBlock(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Paint_Header(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Graphic_RenderDividerBlock(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Paint_BuyerInfo(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Graphic_RenderDividerBlock(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Paint_Terms(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Graphic_RenderSpacerBlock(DrawingCell)

    ' 2. DRAW PRODUCT TABLE HEADER
    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Paint_TableHeader(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Paint_TableSpacer(DrawingCell)

    If ForceFOCHeader Then
        Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
        Set DrawedRange = Paint_TableFOCHeader(DrawingCell)

        Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
        Set DrawedRange = Paint_TableSpacer(DrawingCell)
    End If

    ' 3. LOOP & DRAW PRODUCTS
    ' --------------------------------------------------------------------
    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Render_ProductStack(DrawingCell, pItems, NumberOfProducts, False, ForceTargetHeight)
    If Not outProductRanges Is Nothing Then outProductRanges.Add DrawedRange

    ' 4. THE ELASTIC BUFFER (Double-Armor System)
    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Dim ElasticBufferRange As Object
    Set ElasticBufferRange = Graphic_RenderElasticBuffer(DrawingCell)

    ' THE PRECISION GUARD
    Set DrawingCell = ElasticBufferRange.Offset(ElasticBufferRange.Rows.Count, 0).Cells(1, 1)
    Dim ElasticBufferRange_Guard As Object
    Set ElasticBufferRange_Guard = Graphic_RenderSpacerBlock(DrawingCell)
    ElasticBufferRange_Guard.RowHeight = 0

    ' 5. DRAW FOOTER STACK
    Set DrawingCell = ElasticBufferRange_Guard.Offset(ElasticBufferRange_Guard.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Paint_TableFooter(DrawingCell)

    If ForceTableSummary Then
        Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
        Set outSummaryRow = Paint_TableSummary(DrawingCell)
        Set DrawedRange = outSummaryRow

    End If

    ' FINAL JUSTIFICATION (Geometric Truth Engine)
    ' --------------------------------------------------------------------
    HeightTracker = m_xlSheet.Range(startCell, DrawedRange).Height

    If LAST_CELL_POSITION > HeightTracker Then
        Dim totalGap_F As Double: totalGap_F = LAST_CELL_POSITION - HeightTracker
        Dim unitH_F As Double: unitH_F = Int((totalGap_F / ElasticBufferRange.Rows.Count) * 0.7)

        Dim rIdx_F As Long
        For rIdx_F = 1 To ElasticBufferRange.Rows.Count
            ElasticBufferRange.Rows(rIdx_F).RowHeight = unitH_F
        Next rIdx_F

        ' Final Sync before Guard Mopping
        HeightTracker = m_xlSheet.Range(startCell, DrawedRange).Height
        If LAST_CELL_POSITION > HeightTracker Then Call Precision_SetRowHeight(m_xlSheet, ElasticBufferRange_Guard.row, (LAST_CELL_POSITION - HeightTracker) / 72)
    End If

    Set Layout_FirstPage = DrawedRange
End Function


'product space:  636.5 pts

Private Function Layout_MiddlePage(ByVal startCell As Object, _
                                 ByRef pItems As Collection, _
                                 ByVal NumberOfProducts As Long, _
                                 ByVal ForceTargetHeight As Double, _
                                  Optional ByRef outProductRanges As Collection, _
                                  Optional ByVal ForceFOCHeader As Boolean = False, _
                                  Optional ByVal ForceTableSummary As Boolean = False, _
                                  Optional ByRef outSummaryRow As Object) As Object
    Dim HeightTracker As Double
    Dim DrawingCell As Object, DrawedRange As Object

    ' 1. DRAW HEADER
    Set DrawingCell = startCell
    Set DrawedRange = Graphic_RenderSpacerBlock(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Paint_Header(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Graphic_RenderDividerBlock(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Graphic_RenderSpacerBlock(DrawingCell)
    DrawedRange.RowHeight = 20

    ' 2. DRAW PRODUCT TABLE HEADER
    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Paint_TableHeader(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Paint_TableSpacer(DrawingCell)

    If ForceFOCHeader Then
        Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
        Set DrawedRange = Paint_TableFOCHeader(DrawingCell)

        Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
        Set DrawedRange = Paint_TableSpacer(DrawingCell)
    End If

    ' 3. LOOP & DRAW PRODUCTS
    ' --------------------------------------------------------------------
    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Render_ProductStack(DrawingCell, pItems, NumberOfProducts, False, ForceTargetHeight)
    If Not outProductRanges Is Nothing Then outProductRanges.Add DrawedRange

    ' 4. THE ELASTIC BUFFER (Double-Armor System)
    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Dim ElasticBufferRange As Object
    Set ElasticBufferRange = Graphic_RenderElasticBuffer(DrawingCell)

    ' THE PRECISION GUARD
    Set DrawingCell = ElasticBufferRange.Offset(ElasticBufferRange.Rows.Count, 0).Cells(1, 1)
    Dim ElasticBufferRange_Guard As Object
    Set ElasticBufferRange_Guard = Graphic_RenderSpacerBlock(DrawingCell)
    ElasticBufferRange_Guard.RowHeight = 0

    ' 5. DRAW FOOTER STACK
    Set DrawingCell = ElasticBufferRange_Guard.Offset(ElasticBufferRange_Guard.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Paint_TableFooter(DrawingCell)

    If ForceTableSummary Then
        Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
        Set outSummaryRow = Paint_TableSummary(DrawingCell)
        Set DrawedRange = outSummaryRow

    End If

    ' FINAL JUSTIFICATION (Geometric Truth Engine)
    ' --------------------------------------------------------------------
    HeightTracker = m_xlSheet.Range(startCell, DrawedRange).Height

    If LAST_CELL_POSITION > HeightTracker Then
        Dim totalGap_M As Double: totalGap_M = LAST_CELL_POSITION - HeightTracker
        Dim unitH_M As Double: unitH_M = Int((totalGap_M / ElasticBufferRange.Rows.Count) * 0.7)

        Dim rIdx_M As Long
        For rIdx_M = 1 To ElasticBufferRange.Rows.Count
            ElasticBufferRange.Rows(rIdx_M).RowHeight = unitH_M
        Next rIdx_M

        ' Final Sync before Guard Mopping
        HeightTracker = m_xlSheet.Range(startCell, DrawedRange).Height
        If LAST_CELL_POSITION > HeightTracker Then Call Precision_SetRowHeight(m_xlSheet, ElasticBufferRange_Guard.row, (LAST_CELL_POSITION - HeightTracker) / 72)
    End If

    ' RETURN CURSOR
    Set Layout_MiddlePage = DrawedRange
End Function

'productlist:  424
Private Function Layout_LastPage(ByVal startCell As Object, _
                               ByRef pItems As Collection, _
                               ByVal NumberOfProducts As Long, _
                               ByVal ForceTargetHeight As Double, _
                               Optional ByRef outProductRanges As Collection, _
                               Optional ByVal ForceFOCHeader As Boolean = False, _
                               Optional ByRef outSummaryRow As Object) As Object
    Dim HeightTracker As Double
    Dim DrawingCell As Object, DrawedRange As Object

    ' 1. DRAW HEADER (Z1 + Z2 + Z3)
    Set DrawingCell = startCell
    Set DrawedRange = Graphic_RenderSpacerBlock(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Paint_Header(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Graphic_RenderDividerBlock(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Graphic_RenderSpacerBlock(DrawingCell)
    DrawedRange.RowHeight = 20

    ' 2. DRAW PRODUCT TABLE HEADER
    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Paint_TableHeader(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Paint_TableSpacer(DrawingCell)

    If ForceFOCHeader Then
        Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
        Set DrawedRange = Paint_TableFOCHeader(DrawingCell)

        Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
        Set DrawedRange = Paint_TableSpacer(DrawingCell)
    End If

    ' 3. LOOP & DRAW PRODUCTS
    ' --------------------------------------------------------------------
    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Render_ProductStack(DrawingCell, pItems, NumberOfProducts, False, ForceTargetHeight)
    If Not outProductRanges Is Nothing Then outProductRanges.Add DrawedRange

    ' 4. THE ELASTIC BUFFER (Double-Armor System)
    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Dim ElasticBufferRange As Object
    Set ElasticBufferRange = Graphic_RenderElasticBuffer(DrawingCell)

    ' THE PRECISION GUARD
    Set DrawingCell = ElasticBufferRange.Offset(ElasticBufferRange.Rows.Count, 0).Cells(1, 1)
    Dim ElasticBufferRange_Guard As Object
    Set ElasticBufferRange_Guard = Graphic_RenderSpacerBlock(DrawingCell)
    ElasticBufferRange_Guard.RowHeight = 0

    ' 5. DRAW FOOTER STACK
    Set DrawingCell = ElasticBufferRange_Guard.Offset(ElasticBufferRange_Guard.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Paint_TableFooter(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set outSummaryRow = Paint_TableSummary(DrawingCell)
    Set DrawedRange = outSummaryRow

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Paint_FinalSettlement(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Graphic_RenderSpacerBlock(DrawingCell)

    Set DrawingCell = DrawedRange.Offset(DrawedRange.Rows.Count, 0).Cells(1, 1)
    Set DrawedRange = Paint_BankAndSignature(DrawingCell)

    ' FINAL JUSTIFICATION (Geometric Truth Engine)
    ' --------------------------------------------------------------------
    HeightTracker = m_xlSheet.Range(startCell, DrawedRange).Height

    If LAST_CELL_POSITION > HeightTracker Then
        Dim totalGap_L As Double: totalGap_L = LAST_CELL_POSITION - HeightTracker
        Dim unitH_L As Double: unitH_L = Int((totalGap_L / ElasticBufferRange.Rows.Count) * 0.7)

        Dim rIdx_L As Long
        For rIdx_L = 1 To ElasticBufferRange.Rows.Count
            ElasticBufferRange.Rows(rIdx_L).RowHeight = unitH_L
        Next rIdx_L

        ' Final Sync before Guard Mopping
        HeightTracker = m_xlSheet.Range(startCell, DrawedRange).Height
        If LAST_CELL_POSITION > HeightTracker Then Call Precision_SetRowHeight(m_xlSheet, ElasticBufferRange_Guard.row, (LAST_CELL_POSITION - HeightTracker) / 72)
    End If

    ' RETURN CURSOR
    Set Layout_LastPage = DrawedRange
End Function




'====================================================================
' PRIVATE RENDER ENGINES (The "Diamond" Standard)
'====================================================================

'--------------------------------------------------------------------------------
' PAINT: Paint_Header
' PURPOSE: Branding & Metadata Header (Rows 1-9)
' RETURNS: Range object of the header footprint
'--------------------------------------------------------------------------------
Private Function Paint_Header(ByVal startCell As Object) As Object
    'note on physical size: 25.3 points ~ 0.35 inches.
    Dim baseRow As Long, baseCol As Long
    baseRow = startCell.row
    baseCol = startCell.Column
    ' DESIGN HEIGHT
    Dim hManifest As Variant
    hManifest = Array(19, 10, 10, 11.5, 11.5, 11.5)

    With m_xlSheet 'Set Rows Height
    Dim i As Integer
    For i = 0 To UBound(hManifest)
            .Rows(baseRow + i).RowHeight = hManifest(i)
    Next i
    End With

    ' DESIGN LOGO SPACE
    With startCell

        ' Logo Injection
        Dim logoPath As String, shp As Object, relAnchor As Object
        logoPath = CurrentProject.Path & "\Resources\YandBLogo.png"

            If Dir(logoPath) <> "" Then
                Set relAnchor = .Range("A1:B6") ' [A1:B6]
            ' Vapourise existing shapes
            Precision_ClearZoneShapes relAnchor

            ' Pin picture using relative anchor coordinates
            Set shp = m_xlSheet.Shapes.AddPicture(logoPath, 0, 1, relAnchor.Left, relAnchor.Top, -1, -1)
            With shp
                .LockAspectRatio = -1 ' msoTrue
                .Width = 0.86 * 72    ' 0.86 inches
                .Left = relAnchor.Left + (relAnchor.Width - .Width) / 2
                .Top = relAnchor.Top + (relAnchor.Height - .Height) / 2
            End With
            End If

        'COMPANY/SHIPPER INFORMATION
        With .Range("C1")
            .value = m_Shipper_Name
            .Font.size = 10: .Font.Bold = True: .NumberFormat = "@"
        End With
        With .Range("C2"): .value = m_Shipper_Addr1: .Font.size = 8: .NumberFormat = "@": End With
        With .Range("C3"): .value = m_Shipper_Addr2: .Font.size = 8: .NumberFormat = "@": End With

        ' Contact Metadata (C=Label, D=Value)
        .Range("C4").value = "Web:":   .Range("C4").Font.Bold = False: .Range("C4").Font.size = 8
        With .Range("D4"): .value = m_Shipper_Web: .Font.size = 8: .NumberFormat = "@": End With

        .Range("C5").value = "Email:": .Range("C5").Font.Bold = False: .Range("C5").Font.size = 8
        With .Range("D5"): .value = m_Shipper_Email: .Font.size = 8: .NumberFormat = "@": End With

        .Range("C6").value = "Tel:":   .Range("C6").Font.Bold = False: .Range("C6").Font.size = 8
        With .Range("D6"): .value = m_Shipper_Tel: .Font.size = 8: .NumberFormat = "@": End With



        'COMMERCIAL INVOICE
        With .Range("F1:I2")
            .Merge: .value = "COMMERCIAL INVOICE": .Font.size = 14: .Font.Bold = True
            .HorizontalAlignment = -4108: .VerticalAlignment = -4108
        End With

        'INVOICE NO.
        With .Range("F4")
            .value = "INVOICE NO.:": .Font.Bold = True: .Font.size = 8
            .HorizontalAlignment = -4131: .VerticalAlignment = -4160
        End With

        With .Range("H4:I4")
            .Merge: .value = m_Inv_Number: .Font.size = 8: .NumberFormat = "@"
            .HorizontalAlignment = -4131: .VerticalAlignment = -4107
        End With

        'INVOICE DATE
        With .Range("F5")
            .value = "DATE:": .Font.Bold = True: .Font.size = 8
            .HorizontalAlignment = -4131: .VerticalAlignment = -4160
        End With

        With .Range("H5:I5")
            .Merge: .value = m_Inv_Date: .Font.size = 8: .Font.Bold = False: .NumberFormat = "mmmm d, yyyy"
            .HorizontalAlignment = -4131
        End With
    End With

    Set Paint_Header = startCell.Resize(6, 9)
End Function

'--------------------------------------------------------------------------------
' PAINT: Paint_BuyerInfo
' PURPOSE: Draws Sold To / Shipped To Info
' RETURNS: The Range covered by this zone.
'--------------------------------------------------------------------------------
Private Function Paint_BuyerInfo(ByVal startCell As Object) As Object
    ' Note on physical size : 89.5 points = 1.24 inches.
    ' 1. Anchor & Geometry Initialization
    Dim baseRow As Long: baseRow = startCell.row
    ' ROW HEIGHT - (8 rows content)
    Dim hManifest As Variant
    hManifest = Array(10.5, 11.5, 11.5, 11.5, 11.5, 11.5, 11.5, 10)

    ' Apply Vertical Geometry
    Dim i As Integer
    For i = 0 To UBound(hManifest)
        m_xlSheet.Rows(baseRow + i).RowHeight = hManifest(i)
    Next i

    ' FORMAT FIELDS
    With startCell
        'SOLD TO
        With .Range("A1:B1"): .Merge: .value = "SOLD TO:": .Font.Bold = True: .Font.size = 8: End With
        With .Range("A2"): .value = m_SoldTo_Name: .Font.Bold = False: .Font.size = 8: .NumberFormat = "@": End With
        With .Range("A3"): .value = m_SoldTo_Addr1: .Font.size = 8: .NumberFormat = "@": End With
        With .Range("A4"): .value = m_SoldTo_Addr2: .Font.size = 8: .NumberFormat = "@": End With

        ' Row 5: Contact
        With .Range("A5:B5"): .Merge: .value = "Contact:": .Font.Bold = False: .Font.size = 8: End With
        With .Range("C5"): .value = m_SoldTo_Contact: .Font.size = 8: .NumberFormat = "@": End With

        ' Row 6: Email
        With .Range("A6:B6"): .Merge: .value = "Email:": .Font.Bold = False: .Font.size = 8: End With
        With .Range("C6"): .value = m_SoldTo_Email: .Font.size = 8: .NumberFormat = "@": End With

        ' Row 7: Tel
        With .Range("A7:B7"):  .Merge: .value = "Tel:": .Font.Bold = False: .Font.size = 8: End With
        With .Range("C7"): .value = m_SoldTo_Tel: .Font.size = 8: .NumberFormat = "@": End With

        ' SHIPPED TO
        With .Range("E1:F1"): .Merge: .value = "SHIPPED TO:": .Font.Bold = True: .Font.size = 8: End With
        With .Range("E2"): .value = m_ShipTo_Name: .Font.Bold = False: .Font.size = 8: .NumberFormat = "@": End With
        With .Range("E3"): .value = m_ShipTo_Addr1: .Font.size = 8: .NumberFormat = "@": End With
        With .Range("E4"): .value = m_ShipTo_Addr2: .Font.size = 8: .NumberFormat = "@": End With

        ' Row 5: Contact
        With .Range("E5:F5"): .Merge: .value = "Contact:": .Font.Bold = False: .Font.size = 8: End With
        With .Range("G5"): .value = m_ShipTo_Contact: .Font.size = 8: .NumberFormat = "@": End With

        ' Row 6: Email
        With .Range("E6:F6"): .Merge: .value = "Email:":   .Font.Bold = False: .Font.size = 8: End With
        With .Range("G6"): .value = m_ShipTo_Email:   .Font.size = 8: .NumberFormat = "@": End With

        ' Row 7: Tel
        With .Range("E7:F7"): .Merge: .value = "Tel:":     .Font.Bold = False: .Font.size = 8: End With
        With .Range("G7"): .value = m_ShipTo_Tel:     .Font.size = 8: .NumberFormat = "@": End With
    End With

    ' Return Footprint: (8 rows)
    Set Paint_BuyerInfo = startCell.Resize(8, 9)
End Function

'--------------------------------------------------------------------------------
' PAINT: Paint_Terms
' PURPOSE: Draws Logistics / Terms Section.
' RETURNS: The Range covered by this zone.
'--------------------------------------------------------------------------------
Private Function Paint_Terms(ByVal startCell As Object) As Object
    'Note on physical size: 80.5 points  = 1.12 inches.
    ' 1. Anchor & Geometry Initialization
    Dim baseRow As Long: baseRow = startCell.row

    ' GEOMETRY: (7 content rows)
    Dim i As Integer
    For i = 0 To 6
        m_xlSheet.Rows(baseRow + i).RowHeight = 11.5
    Next i

    ' 2. RELATIVE MAPPING
    With startCell
        ' ROW 1: Delivery Term and PO Number
        .Range("A1").value = "DELIVERY TERM:": .Range("A1").Font.Bold = True: .Range("A1").Font.size = 8
        With .Range("D1:D2"): .Merge: .value = m_Term_Incoterm: .Font.size = 8: .NumberFormat = "@": .VerticalAlignment = -4160: End With

        With .Range("E1:F1"): .Merge: .value = "PO NO.:": .Font.Bold = True: .Font.size = 8: End With
        With .Range("G1:I2"): .Merge: .value = m_Term_PO: .Font.size = 8: .NumberFormat = "@": .VerticalAlignment = -4160: End With

        ' ROW 3: Port of Loading and Reference
        .Range("A3").value = "PORT OF LOADING:": .Range("A3").Font.Bold = True: .Range("A3").Font.size = 8
        With .Range("D3:D4"): .Merge: .value = m_Term_POL: .Font.size = 8: .NumberFormat = "@": .VerticalAlignment = -4160: End With

        With .Range("E3:F3"): .Merge: .value = "REFERENCE:": .Font.Bold = True: .Font.size = 8: End With
        With .Range("G3:I4"): .Merge: .value = m_Term_Ref: .Font.size = 8: .NumberFormat = "@": .VerticalAlignment = -4160: End With

        ' ROW 5: Port of Discharge and Payment Term
        .Range("A5").value = "PORT OF DISCHARGE:": .Range("A5").Font.Bold = True: .Range("A5").Font.size = 8
        With .Range("D5:D6"): .Merge: .value = m_Term_POD: .Font.size = 8: .NumberFormat = "@": .VerticalAlignment = -4160: End With

        With .Range("E5:F5"): .Merge: .value = "PAYMENT TERM:": .Font.Bold = True: .Font.size = 8: End With
        With .Range("G5:I6"): .Merge: .value = m_Term_Payment: .Font.size = 8: .NumberFormat = "@": .VerticalAlignment = -4160: End With

        ' ROW 7: Currency
        With .Range("E7:F7"): .Merge: .value = "CURRENCY:": .Font.Bold = True: .Font.size = 8: End With
        With .Range("G7:I7"): .Merge: .value = m_Term_Currency: .Font.size = 8: .NumberFormat = "@": End With
    End With

    ' Return Footprint: (7 rows)
    Set Paint_Terms = startCell.Resize(7, 9)
End Function

'--------------------------------------------------------------------------------
' RENDER: Render_ProductStack
Private Function Render_ProductStack(ByRef DrawingCell As Object, ByRef pItems As Collection, _
                                    Optional ByVal NumberOfProducts As Long = 0, _
                                    Optional ByVal TempSheetMode As Boolean = False, _
                                    Optional ByVal ForceTargetHeight As Double = 0) As Object
    If pItems Is Nothing Then Set Render_ProductStack = DrawingCell: Exit Function
    If pItems.Count = 0 Then
        Set Render_ProductStack = DrawingCell.Resize(1, 1)
        Render_ProductStack.RowHeight = 0
        Exit Function
    End If

    Dim iStartRow As Long: iStartRow = DrawingCell.row
    Dim currentRange As Object
    Dim StoreRange As Collection
    Set StoreRange = New Collection
    Dim mockData As Variant

    Dim targetDeckSize As Long

    ' Determine our exit condition BEFORE we start looping
    If NumberOfProducts <= 0 Then
        targetDeckSize = 0 ' We want to eat everything until the deck is empty
    Else
        targetDeckSize = pItems.Count - NumberOfProducts ' We want to eat a specific amount
    End If

    ' Ensure we don't try to eat more than exists
    If targetDeckSize < 0 Then targetDeckSize = 0

    Do While pItems.Count > targetDeckSize
        mockData = pItems(1)

        ' 1. Paint the Product Item
        Set currentRange = Paint_TableRow(DrawingCell, _
            CStr(mockData(LBound(mockData) + 0)), _
            CStr(mockData(LBound(mockData) + 1)), _
            CStr(mockData(LBound(mockData) + 2)), _
            CDbl(mockData(LBound(mockData) + 3)), _
            CStr(mockData(LBound(mockData) + 4)), _
            CDbl(mockData(LBound(mockData) + 5)), _
            CStr(mockData(LBound(mockData) + 6)), _
            CStr(mockData(LBound(mockData) + 7)), _
            CStr(mockData(LBound(mockData) + 8)))

        ' Apply forced target height instantly (No dummy printing stretch needed for Real Pages)
        If ForceTargetHeight > 0 Then
            ' Adjust Row 2 (Metadata) to absorb the difference between natural height and forced height
            If currentRange.Height < ForceTargetHeight Then
                currentRange.Rows(2).RowHeight = ForceTargetHeight - currentRange.Rows(1).RowHeight
            End If
        End If

        pItems.Remove 1
        StoreRange.Add currentRange
        ' Advance Pen
        Set DrawingCell = currentRange.Offset(currentRange.Rows.Count, 0).Cells(1, 1)
    Loop

    ' Only mathematically equalize if it's running in TempSheet mode (The Surveyor)
    If TempSheetMode Then
        Dim item As Object
        Dim maxHeight As Double: maxHeight = 0

        ' 1. Find the "Alpha" (Tallest Item) on this batch
        For Each item In StoreRange
            If item.Height > maxHeight Then
                maxHeight = item.Height
            End If
        Next item

        ' 2. "Inflate" all other items to match the Alpha height
        If maxHeight > 0 Then
            For Each item In StoreRange
                If item.Height < maxHeight Then
                    ' Adjust Row 2 (Metadata) to absorb the difference
                    item.Rows(2).RowHeight = maxHeight - item.Rows(1).RowHeight
                End If
            Next item
        End If
    End If

    ' Return the entire stack footprint
    If DrawingCell.row > iStartRow Then
        Set Render_ProductStack = DrawingCell.Worksheet.Rows(iStartRow & ":" & DrawingCell.row - 1)
    Else
        Set Render_ProductStack = DrawingCell ' Fallback if nothing was drawn
    End If
End Function

'--------------------------------------------------------------------------------
' PAINT: Paint_TableHeader
' PURPOSE: Draws the Steel Frame Horizontal Bar Header.
'--------------------------------------------------------------------------------
Private Function Paint_TableHeader(ByVal startCell As Object) As Object
    ' [PLATINUM ENGINE: STEEL HEADER]
    ' Note on physical size: 25.3 point - 0.35"

    Dim baseRow As Long: baseRow = startCell.row

    ' 1. GEOMETRY: (1 row Steel Header)
    m_xlSheet.Rows(baseRow).RowHeight = 25.3

    ' 2. THE STEEL FRAME MASTER HEADER
    With m_xlSheet.Range(m_xlSheet.Cells(baseRow, 1), m_xlSheet.Cells(baseRow, 9))
        .Interior.ColorIndex = -4142
        .Font.Color = 0
        .Font.Bold = True: .Font.size = 8: .Font.name = "Arial"
        .HorizontalAlignment = -4108: .VerticalAlignment = -4108: .WrapText = True

        .Cells(1, 1).value = "NO."
        .Cells(1, 2).value = "HS CODE"
        With m_xlSheet.Range(m_xlSheet.Cells(baseRow, 3), m_xlSheet.Cells(baseRow, 5))
            .Merge: .value = "DESCRIPTION OF GOODS"
        End With
        .Cells(1, 6).value = "QTY"
        .Cells(1, 7).value = "UNIT"
        .Cells(1, 8).value = "PRICE" & vbLf & "(USD)"
        .Cells(1, 9).value = "AMOUNT" & vbLf & "(USD)"

        .Borders.LineStyle = -4142
        With .Borders(8): .LineStyle = 1: .Weight = 4: End With
        With .Borders(9): .LineStyle = 1: .Weight = 4: End With
    End With

    Set Paint_TableHeader = startCell.Resize(1, 9)
End Function


'--------------------------------------------------------------------------------
' PAINT: Paint_TableFOCHeader
' PURPOSE: Draws a high-visibility divider for Free-of-Charge items.
'--------------------------------------------------------------------------------
Private Function Paint_TableFOCHeader(ByVal startCell As Object) As Object
    With startCell.Resize(1, 9)
        .Merge
        .value = MSG_FOC

        .Font.name = "Arial": .Font.Bold = True: .Font.size = 9
        .Font.Color = 0

        .HorizontalAlignment = -4108: .VerticalAlignment = -4108
        .RowHeight = 22
        .Borders.LineStyle = -4142 ' xlNone
        .Interior.ColorIndex = -4142 ' xlNone
    End With

    Set Paint_TableFOCHeader = startCell.Resize(1, 9)
End Function



'--------------------------------------------------------------------------------
' PAINT: Paint_TableRow
' PURPOSE: Draws the 2-row Primary Product Block.
' DESIGN : Row 1 (Main Data) + Row 2 (Trace Metadata).
'--------------------------------------------------------------------------------
Private Function Paint_TableRow(ByVal startCell As Object, _
                                        ByVal Numbering As Variant, _
                                        ByVal HSCode As String, _
                                        ByVal Description As String, _
                                        ByVal Qty As Double, _
                                        ByVal Unit As String, _
                                        ByVal Price As Double, _
                                        ByVal Barcode As String, _
                                        ByVal LotNum As String, _
                                        ByVal Expiry As String) As Object
    ' Visual Anchor: Each Block contains [Row 1: Main Data] and [Row 2: Metadata/Desc]

    '========================================================================
    ' PHASE 1: FRAME & FORMATTING (Data-Agnostic)
    '========================================================================
    ' Global Block Base
    With startCell.Resize(2, 9)
        .Font.name = "Arial": .Font.size = 8: .Font.Bold = False
        .VerticalAlignment = -4160 ' xlTop
    End With

    ' Row 1: Primary Dimensions & Alignment
    With startCell.Rows(1)
        .Cells(1, 1).HorizontalAlignment = -4108: .Cells(1, 1).NumberFormat = "0" ' NO. (Center Numeric)
        .Cells(1, 2).HorizontalAlignment = -4108: .Cells(1, 2).NumberFormat = "@" ' HS CODE (Center Text)
        With .Range("C1:E1")
            .Merge: .HorizontalAlignment = -4131: .WrapText = True: .NumberFormat = "@" ' DESCRIPTION
        End With
        .Cells(1, 6).HorizontalAlignment = -4108: .Cells(1, 6).NumberFormat = "#,##0" ' QTY (Center Whole Number)
        .Cells(1, 7).HorizontalAlignment = -4108: .Cells(1, 7).NumberFormat = "@"      ' UNIT (Center Text)
        .Cells(1, 8).HorizontalAlignment = -4152: .Cells(1, 8).NumberFormat = "#,##0.00" ' PRICE (Right Dec)
        .Cells(1, 9).HorizontalAlignment = -4152: .Cells(1, 9).NumberFormat = "#,##0.00" ' AMOUNT (Right Dec)
    End With

    ' Row 2: Metadata Dimensions & Typography
    startCell.Rows(2).RowHeight = 2 * LINE_UNIT
    With startCell.Rows(2).Range("C1:E1")
        .Merge: .NumberFormat = "@"
        .Font.size = 7: .IndentLevel = 1
        .HorizontalAlignment = -4131: .VerticalAlignment = -4160
    End With

    '========================================================================
    ' PHASE 2: DATA INJECTION (Normally pushed by Access DB Engine)
    '========================================================================
    startCell.Cells(1, 1).value = Numbering
    startCell.Cells(1, 2).value = HSCode
    startCell.Rows(1).Range("C1:E1").value = Description
    startCell.Cells(1, 6).value = Qty
    startCell.Cells(1, 7).value = Unit
    startCell.Cells(1, 8).value = Price
    startCell.Cells(1, 9).FormulaR1C1 = "=RC[-3]*RC[-1]"

    ' Row 2 Trace Data Header
    startCell.Rows(2).Range("C1:E1").value = "Barcode: " & Barcode & " | Lot: " & LotNum & " | Exp: " & Format(Expiry, "mmmm d, yyyy")

    '========================================================================
    ' PHASE 3: SMART ROW HEIGHT ADJUSTER (Phantom Cell Engine via Titanium Loop)
    '========================================================================
    Dim zCell As Object
    Dim targetWidth As Double

    ' 1. Set physical Width of C:E merge target from Master Blueprint
    targetWidth = DESC_ZONE_WIDTH

    ' 2. Target phantom cell and use Titanium Engine to set exact physical width
    Set zCell = startCell.Cells(1, 26) ' Column Z
    Call Precision_SetColWidth(startCell.Worksheet, "Z", targetWidth / 72)

    ' 3. Inject data and trigger layout physics
    zCell.value = Description
    zCell.Font.size = 8
    zCell.Font.name = "Arial"
    zCell.Font.Bold = False
    zCell.WrapText = True

    ' 4. AutoFit, Harvest Height, Enforce minimum, and Cleanup
    zCell.EntireRow.AutoFit

    Dim rowLines As Integer
    rowLines = -Int(-((zCell.Height - 2) / LINE_UNIT))

    startCell.Rows(1).RowHeight = rowLines * LINE_UNIT
    zCell.Clear

    ' Return Footprint (2 rows)
    Set Paint_TableRow = startCell.Resize(2, 9)
End Function

'--------------------------------------------------------------------------------
' PAINT: Paint_TableSpacer
' PURPOSE: Adds rhythmic vertical padding between rows.
'--------------------------------------------------------------------------------
Function Paint_TableSpacer(ByVal startCell As Object) As Object

    ' Anchor to 2 rows spanning the requested number of columns (Default A to I)
    Dim rngSpacer As Object
    Set rngSpacer = startCell.Resize(1, SPAN_COLS)

    ' 1. Resize top row (the padding before the line)
    rngSpacer.Rows(1).RowHeight = SPACER_07

    ' Return Footprint
    Set Paint_TableSpacer = rngSpacer
End Function

'--------------------------------------------------------------------------------
' PAINT: Paint_TableFooter
' PURPOSE: Draws the dividing line between the product stack and totals.
'--------------------------------------------------------------------------------
Function Paint_TableFooter(ByVal startCell As Object) As Object

    ' Anchor to 1 row spanning the requested number of columns (Default A to I)
    Dim rngDivider As Object
    Set rngDivider = startCell.Resize(1, SPAN_COLS)

    ' 1. Resize top row (the padding before the line)
    rngDivider.Rows(1).RowHeight = SPACER_07

    ' 2. Draw the line BETWEEN them (Bottom border of Row 1 spans continuously)
    With rngDivider.Rows(1).Borders(9) ' 9 = xlEdgeBottom
        .LineStyle = 1 ' xlContinuous
        .Weight = 2 ' 1=Hairline, 2=Thin, 4=Thick
        .Color = RGB(0, 0, 0) ' Pure Black
    End With

    ' Return Footprint
    Set Paint_TableFooter = rngDivider
End Function



'--------------------------------------------------------------------------------
' PAINT: Paint_TableSummary
' PURPOSE: Renders synthesized totals for the product table.
'--------------------------------------------------------------------------------
Public Function Paint_TableSummary(ByVal startCell As Object) As Object
    Dim rngMain As Object: Set rngMain = startCell.Resize(1, 9)

    With rngMain.Rows(1)
        .RowHeight = 1.5 * LINE_UNIT
        .VerticalAlignment = -4108 ' Center

        ' Quantity Pillar
        With .Range("D1:E1"): .Merge: .value = "TOTAL QUANTITY:": .Font.Bold = True: .HorizontalAlignment = -4152: End With
        With .Cells(1, 6): .HorizontalAlignment = -4108: .NumberFormat = "#,##0": End With

        ' Amount Pillar
        With .Cells(1, 8): .value = "TOTAL AMOUNT:": .Font.Bold = True: .HorizontalAlignment = -4152: End With
        With .Cells(1, 9): .NumberFormat = "#,##0.00": .HorizontalAlignment = -4152: End With
    End With

    Set Paint_TableSummary = rngMain
End Function

'--------------------------------------------------------------------------------
' PAINT: Paint_FinalSettlement
' PURPOSE: Draws the document's global final payable amount and verbal spelling.
'--------------------------------------------------------------------------------
Public Function Paint_FinalSettlement(ByVal startCell As Object) As Object
    ' Footprint: 4 Rows (Totals, Gap, 2-Row Verbal Spelling)
    Dim rngMain As Object: Set rngMain = startCell.Resize(4, 9)

    ' 1. THE BALANCE (Row 1)
    With rngMain.Rows(1)
        .RowHeight = 1.5 * LINE_UNIT
        .VerticalAlignment = -4108 ' Center

        ' No label on Column 1
        ' With .Cells(1, 1): .Value = "VII. TOTALS": .Font.Bold = True: End With

        ' Qty Pillar (Aligning with Column 6)
        With .Cells(1, 4).Resize(1, 2): .Merge: .value = "GRAND TOTAL QUANTITY:": .Font.Bold = True: .HorizontalAlignment = -4152: End With
        With .Cells(1, 6): .value = 0: .HorizontalAlignment = -4108: End With

        ' Amount Pillar (Column 8/9)
        With .Cells(1, 8): .value = "PAYABLE AMOUNT:": .Font.Bold = True: .HorizontalAlignment = -4152: End With
        With .Cells(1, 9): .value = m_Sett_PayableAmount: .NumberFormat = "#,##0.00": .HorizontalAlignment = -4152: End With

    End With

    ' CAPTURE: Save this location in the global pin for the final "Synthesis" logic
    Set m_pinGrandTotalRow = rngMain.Rows(1)

    ' 2. THE VISUAL GAP (Row 2)
    ' This provides a 0.07" breathable space between the math and the spelling.
    rngMain.Rows(2).RowHeight = SPACER_07

    ' 3 & 4. THE VERBAL BLOCK (Merged Area)
    ' We merge D to I over two rows to allow for long verbal spelling strings.
    With rngMain.Range("A3:C3")
        .Merge: .value = "SAY PAYABLE AMOUNT:": .Font.Bold = True: .HorizontalAlignment = -4131: .VerticalAlignment = -4160
    End With

    With rngMain.Range("D3:I4")
        .Merge: .value = m_Sett_AmountInWords: .NumberFormat = "@"
        .Font.Italic = True: .HorizontalAlignment = -4131: .VerticalAlignment = -4160
        .WrapText = True
    End With

    Set Paint_FinalSettlement = rngMain
End Function

'--------------------------------------------------------------------------------
' PAINT: Paint_BankAndSignature
' PURPOSE: Renders Certifications, Banking Vault, and the Signatory Block.
'--------------------------------------------------------------------------------
Public Function Paint_BankAndSignature(ByVal startCell As Object) As Object
    ' Footprint: 3 Static Rows
    Dim rngMain As Object: Set rngMain = startCell.Resize(3, 9)

    ' Clear existing shapes
    Call Precision_ClearZoneShapes(rngMain)

    '========================================================================
    ' PHASE 1: CERTIFICATIONS & COMPANY HEADER (Rows 1-2)
    '========================================================================
    With rngMain.Rows(1)
        .RowHeight = 15: .VerticalAlignment = -4107 ' xlBottom
        ' Left: Quality
        .Cells(1, 1).value = "QUALITY:": .Cells(1, 1).Font.Bold = True
        With .Cells(1, 3): .value = m_Sett_Quality: .HorizontalAlignment = -4131: .NumberFormat = "@": End With
        ' Right: Company Header (Center Aligned as requested)
        With .Range("G1:I1"): .Merge: .value = m_Sign_ShipperName: .Font.Bold = True: .Font.size = 10: .HorizontalAlignment = -4108: .NumberFormat = "@": End With
    End With

    With rngMain.Rows(2)
        .RowHeight = 15: .VerticalAlignment = -4107 ' xlBottom
        ' Left: Origin
        .Cells(1, 1).value = "ORIGIN:": .Cells(1, 1).Font.Bold = True
        With .Cells(1, 3): .value = m_Sett_Origin: .HorizontalAlignment = -4131: .NumberFormat = "@": End With
    End With

    '========================================================================
    ' PHASE 2: THE MAIN VAULT (Row 3 - 1.7" Chamber)
    ' This one tall row contains both the Bank Info and the Signatory Name
    ' sitting at the absolute bottom.
    '========================================================================
    With rngMain.Rows(3)
        .RowHeight = 122.4 ' Exactly 1.7 inches

        ' LEFT: Banking Pillar
        With .Range("A1:F1")
            .Merge: .NumberFormat = "@": .value = m_Sign_BankInfo
            .Font.size = 7: .WrapText = True
            .HorizontalAlignment = -4131: .VerticalAlignment = -4107 ' xlLeft, xlBottom
        End With

        ' RIGHT: Signature Gap + Name (The name sits at the base of the chamber)
        With .Range("G1:I1")
            .Merge: .NumberFormat = "@": .value = m_Sign_ShipperName
            .Font.Bold = True: .HorizontalAlignment = -4108: .VerticalAlignment = -4107 ' xlCenter, xlBottom
        End With
    End With

    Set Paint_BankAndSignature = rngMain
End Function




'HELPER
'--------------------------------------------------------------------------------
' PRECISION: Precision_SetColWidth
' PURPOSE: Iterative Refinement Engine for Physical Column Sizing.
'--------------------------------------------------------------------------------
Private Sub Precision_SetColWidth(ByVal targetSheet As Object, ByVal colName As String, ByVal targetInch As Double)

    Dim targetPoints As Double, currentPoints As Double
    Dim currentUnits As Double
    Dim ratio As Double
    Dim i As Integer

    targetPoints = targetInch * 72

    With targetSheet.Columns(colName)

        .ColumnWidth = targetPoints * (.ColumnWidth / .Width)

        For i = 1 To 3
            currentPoints = .Width
            currentUnits = .ColumnWidth

            If Abs(targetPoints - currentPoints) < 0.01 Then Exit For

            ' Adjust the Unit value based on the Point error
            .ColumnWidth = currentUnits * (targetPoints / currentPoints)
        Next i
    End With
End Sub


'--------------------------------------------------------------------------------
' PRECISION: Precision_SetRowHeight
' PURPOSE: Iterative Refinement Engine for Physical Row Sizing.
'--------------------------------------------------------------------------------
Private Sub Precision_SetRowHeight(ByVal targetSheet As Object, ByVal rowNum As Long, ByVal targetInch As Double)
    Dim targetPoints As Double, currentPoints As Double
    Dim currentUnits As Double
    Dim i As Integer

    targetPoints = targetInch * 72

    With targetSheet.Rows(rowNum)
        ' Initial assignment (RowHeight is point-based)
        .RowHeight = targetPoints

        ' Iterative refinement for Platinum precision
        For i = 1 To 3
            currentPoints = .Height
            currentUnits = .RowHeight

            If Abs(targetPoints - currentPoints) < 0.01 Then Exit For

            ' Adjust based on point error (Self-correcting loop)
            .RowHeight = currentUnits * (targetPoints / currentPoints)
        Next i
    End With
End Sub


' ======================================================================================
' UTILITY: Precision_ClearZoneShapes
' Purpose: Deletes any floating shapes (logos, signatures) touching the target range.
' ======================================================================================
Private Sub Precision_ClearZoneShapes(ByVal rngTarget As Object)
    Dim shp As Object
    On Error Resume Next ' Defensive
    For Each shp In rngTarget.Parent.Shapes
        ' If the Top-Left corner of the shape is within our anchor range, delete it.
        ' Must use .Application.Intersect because we are in Access, not Excel native.
        If Not rngTarget.Application.Intersect(shp.TopLeftCell, rngTarget) Is Nothing Then
            shp.Delete
        End If
    Next shp
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------------
' HELPER: BindCategorySummary
' PURPOSE: Binds a collection of product ranges to a summary pin via SUM formulas.
'--------------------------------------------------------------------------------
Private Sub BindCategorySummary(ByVal summaryPin As Object, ByVal productBlocks As Collection)
    If summaryPin Is Nothing Or productBlocks Is Nothing Then Exit Sub
    If productBlocks.Count = 0 Then Exit Sub

    Dim fQty As String, fAmt As String, i As Long
    fQty = "=SUM(": fAmt = "=SUM("

    For i = 1 To productBlocks.Count
        fQty = fQty & productBlocks(i).Columns(6).Address(False, False) & IIf(i = productBlocks.Count, "", ",")
        fAmt = fAmt & productBlocks(i).Columns(9).Address(False, False) & IIf(i = productBlocks.Count, "", ",")
    Next i

    fQty = fQty & ")": fAmt = fAmt & ")"

    summaryPin.Cells(1, 6).Formula = fQty
    summaryPin.Cells(1, 9).Formula = fAmt
End Sub

' ======================================================================================
' GRAPHIC UTILITY: Graphic_RenderDividerBlock
' Purpose: Renders a generic 2-row geometric block with a dividing line in between.
' Returns: A 2-row Range object representing the geometric footprint.
' ======================================================================================
Public Function Graphic_RenderDividerBlock(ByVal startCell As Object) As Object


    ' Anchor to 2 rows spanning the requested number of columns (Default A to I)
    Dim rngDivider As Object
    Set rngDivider = startCell.Resize(2, SPAN_COLS)

    ' 1. Resize top row (the padding before the line)
    rngDivider.Rows(1).RowHeight = SPACER_07

    ' 2. Resize bottom row (the padding after the line)
    rngDivider.Rows(2).RowHeight = SPACER_07

    ' 3. Draw the line BETWEEN them (Bottom border of Row 1 spans continuously)
    With rngDivider.Rows(1).Borders(9) ' 9 = xlEdgeBottom
        .LineStyle = 1 ' xlContinuous
        .Weight = 2 ' 1=Hairline, 2=Thin, 4=Thick
        .ColorIndex = 0 ' Pure Black
    End With

    ' Return Footprint
    Set Graphic_RenderDividerBlock = rngDivider
End Function

Function Graphic_RenderSpacerBlock(ByVal startCell As Object) As Object
    Dim rngSpacer As Object
    Set rngSpacer = startCell.Resize(1, SPAN_COLS)
    rngSpacer.Rows(1).RowHeight = SPACER_07
    Set Graphic_RenderSpacerBlock = rngSpacer
End Function

Public Function Graphic_RenderElasticBuffer(ByVal startCell As Object) As Object
    ' Multi-row zone (10 rows) to distribute justification load and bypass 409pt limit
    Dim rngBuffer As Object
    Set rngBuffer = startCell.Resize(3, SPAN_COLS)
    rngBuffer.RowHeight = 0
    Set Graphic_RenderElasticBuffer = rngBuffer
End Function

Private Function GetMock_ProductData() As Variant
    Dim data(0 To 9, 0 To 8) As Variant
    Dim i As Long, baseIdx As Long

    ' Template items to duplicate
    Dim tNames As Variant: tNames = Array( _
        "Coffee Body Polish (200ml)", _
        "Lip Balm with Coconut Oil (5g)", _
        "Pomelo Hair Tonic (140ml)", _
        "Winter Melon Cleansing Gel (500ml)", _
        "Turmeric Face Cream (50ml)" _
    )

    Dim tHS As Variant: tHS = Array("33049930", "33041000", "33059000", "34013000", "33049930")
    Dim tPrice As Variant: tPrice = Array(12.5, 4.75, 18.25, 145.5, 15.99)

    For i = 0 To 9
        baseIdx = i Mod 5
        data(i, 0) = i + 1                                     ' Item #
        data(i, 1) = tHS(baseIdx)                               ' HS Code

        If i < 5 Then
            data(i, 2) = tNames(baseIdx)                       ' Paid Label
            data(i, 5) = tPrice(baseIdx)                        ' Paid Price
        Else
            data(i, 2) = tNames(baseIdx) & " (FOC)"            ' FOC Label
            data(i, 5) = 0                                      ' FOC Price
        End If

        data(i, 3) = 100 + i                                    ' Qty
        data(i, 4) = "PCS"                                      ' Unit
        data(i, 6) = "8936217700" & Format(i, "000")            ' Barcode
        data(i, 7) = "BATCH-" & i                               ' Batch
        data(i, 8) = "01/01/2027"                               ' Date
    Next i

    GetMock_ProductData = data
End Function
