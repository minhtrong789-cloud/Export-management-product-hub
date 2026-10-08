Attribute VB_Name = "ExcelRender_Products_PrintProducts"
Option Compare Database
Option Explicit
'@Folder "App.Products"
' ==================================================================================
' Module:       ExcelRender_Products_PrintProducts
' Architecture: Orchestrator-Worker-Helper (Data-Driven Design)
' Purpose:      Programmatically builds the Excel Catalogue Template with Platinum Design
' ==================================================================================

' --- 1. CORE DATA STRUCTURES (Metadata) ---

Public Enum ExportDataType
    edtText = 0
    edtNumber = 1
    edtNumberAsText = 2
    edtImage = 3
    edtCurrency = 4
End Enum

Public Type ColumnMap
    ColumnIndex As Integer
    Caption As String
    fieldName As String
    Width As Double
    DataType As ExportDataType
End Type

' --- 2. ORCHESTRATOR (High-level API) ---

' Purpose: Bootstraps Excel and triggers the layout generation.
Public Sub GetTemplate_ProductList(ByVal rs As Object)
    On Error GoTo ErrHandler

    Dim xlApp As Object
    Dim wb As Object
    Dim ws As Object

    ' 1. Start Excel
    On Error GoTo CreateNewExcel
    Set xlApp = GetObject(, "Excel.Application")
    GoTo ExcelReady

CreateNewExcel:
    Err.Clear
    On Error GoTo ErrHandler
    Set xlApp = CreateObject("Excel.Application")

ExcelReady:
    On Error GoTo ErrHandler
    xlApp.Visible = False

    Dim visualName As String
    visualName = "Product_DataTemplate_" & Format(Now(), "yyyymmdd_hhnnss")
    Set wb = xlApp.Workbooks.Add
    Set ws = wb.Sheets(1)

    ' 2. Build the Shell (Static Layout for Import Compatibility)
    Dim pMap() As ColumnMap
    pMap = Layout_ProductList(ws)

    ' 3. Render Data
    Call Render_ProductData(ws, pMap, rs)
    rs.Close
    Set rs = Nothing

    ' 4. Finalize & Display
    ws.name = "Bulk Edit Worksheet"
    wb.Windows(1).Caption = visualName

    xlApp.Visible = True
    Exit Sub

ErrHandler:
    If Not xlApp Is Nothing Then xlApp.Visible = True
    Debug.Print "[ERROR] ExcelRender_Products_PrintProducts.GetTemplate_ProductList: #" & Err.Number & " - " & Err.Description
    MsgBox "System Error: " & Err.Description, vbCritical
End Sub

' Purpose: Orchestrates the dynamic export of the Product List based on provided Recordset.
Public Sub Print_ProductList(ByVal langID As Integer, ByVal rs As Object)
    On Error GoTo ErrHandler

    Dim xlApp As Object
    Dim wb As Object
    Dim ws As Object

    ' 1. Start Excel
    On Error GoTo CreateNewExcel
    Set xlApp = GetObject(, "Excel.Application")
    GoTo ExcelReady

CreateNewExcel:
    Err.Clear
    On Error GoTo ErrHandler
    Set xlApp = CreateObject("Excel.Application")

ExcelReady:
    On Error GoTo ErrHandler
    xlApp.Visible = False

    Dim visualName As String
    visualName = "Product_Catalogue_" & Format(Now(), "yyyymmdd_hhnnss")
    Set wb = xlApp.Workbooks.Add
    Set ws = wb.Sheets(1)

    ' 2. Build the Dynamic Shell
    Dim pMap() As ColumnMap
    pMap = Layout_DynamicProductList(ws, langID)

    ' 3. Render Data
    Call Render_ProductData(ws, pMap, rs)

    ' 4. Finalize & Display
    ws.name = "Product Catalogue"
    wb.Windows(1).Caption = visualName

    xlApp.Visible = True
    Exit Sub

ErrHandler:
    If Not xlApp Is Nothing Then xlApp.Visible = True
    Debug.Print "[ERROR] ExcelRender_Products_PrintProducts.Print_ProductList: #" & Err.Number & " - " & Err.Description
    MsgBox "System Error: " & Err.Description, vbCritical
End Sub


' --- 3. WORKERS (Mid-level Implementations) ---

' Purpose: Generates the STATIC ColumnMap metadata and paints the Excel layout shell.
Private Function Layout_ProductList(ws As Object) As ColumnMap()
    Dim pMap(1 To 21) As ColumnMap

    ' 1. Define Column Mapping
    pMap(1) = ColumnMap_createCol(1, "No.", "", 5, edtNumber)
    pMap(2) = ColumnMap_createCol(2, "Group / Line", "CategoryFull", 25)
    pMap(3) = ColumnMap_createCol(3, "Barcode", "Barcode", 15, edtNumberAsText)
    pMap(4) = ColumnMap_createCol(4, "HS Code", "HSCode", 12)
    pMap(5) = ColumnMap_createCol(5, "Name (VN)", "ProductName_vn", 25)
    pMap(6) = ColumnMap_createCol(6, "Name (EN)", "ProductName_en", 25)
    pMap(7) = ColumnMap_createCol(7, "Image", "ImagePath", 12, edtImage)
    pMap(8) = ColumnMap_createCol(8, "Package Type", "PackageType", 12)
    pMap(9) = ColumnMap_createCol(9, "Capacity", "CapacityFull", 15)
    pMap(10) = ColumnMap_createCol(10, "Price", "Price_vnd", 12, edtCurrency)
    pMap(11) = ColumnMap_createCol(11, "Price (+VAT 8%)", "Price_VAT", 12, edtCurrency)
    pMap(12) = ColumnMap_createCol(12, "Capacity (OZ)", "Capacity_oz", 12, edtNumber)
    pMap(13) = ColumnMap_createCol(13, "Length (CM)", "Length_cm", 8, edtNumber)
    pMap(14) = ColumnMap_createCol(14, "Width (CM)", "Width_cm", 8, edtNumber)
    pMap(15) = ColumnMap_createCol(15, "Height (CM)", "Height_cm", 8, edtNumber)
    pMap(16) = ColumnMap_createCol(16, "Weight (G)", "Weight_gram", 10, edtNumber)
    pMap(17) = ColumnMap_createCol(17, "Ingredients", "Ingredients", 50)
    pMap(18) = ColumnMap_createCol(18, "USP (VN)", "USP_vn", 35)
    pMap(19) = ColumnMap_createCol(19, "USP (EN)", "USP_en", 35)
    pMap(20) = ColumnMap_createCol(20, "Usage (VN)", "HowToUse_vn", 35)
    pMap(21) = ColumnMap_createCol(21, "Usage (EN)", "HowToUse_en", 35)

    ' 2. Paint Headers & Widths
    Dim i As Integer
    For i = LBound(pMap) To UBound(pMap)
        ws.Cells(1, i).value = pMap(i).Caption
        ws.Columns(i).ColumnWidth = pMap(i).Width
    Next i

    ' 3. Style Header Row
    With ws.Range("A1:U1")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(44, 62, 80)
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
        .RowHeight = 35
        .WrapText = True
    End With

    ' 4. Set Global Cell Alignment
    With ws.Range("A:U")
        .VerticalAlignment = -4108
        .Font.name = "Segoe UI"
        .Font.size = 10
    End With

    ' 5. Specific Column Formatting
    ws.Range("E:F,Q:U").WrapText = True
    ws.Range("J:K").NumberFormat = "#,##0"

    ' 6. Freeze Panes (Standard Platinum UI)
    ws.Activate
    ws.Range("A2").Select
    ws.Parent.Windows(1).FreezePanes = True

    ' 7. Borders for Header
    With ws.Range("A1:U1").Borders
        .LineStyle = 1
        .Color = RGB(0, 0, 0)
        .Weight = 2
    End With

    ' 8. Return Map to Caller
    Layout_ProductList = pMap
End Function

' Purpose: Generates a dynamic, schema-driven layout shell filtered by Language ID.
Private Function Layout_DynamicProductList(ws As Object, langID As Integer) As ColumnMap()
    ' 1. Define Column Mapping
    Dim pMap() As ColumnMap
    Dim i As Integer: i = 0

    Call ColumnMap_addCol(pMap, i, "No.", "", 5, edtNumber)
    Call ColumnMap_addCol(pMap, i, "Group / Line", "CategoryFull", 25)
    Call ColumnMap_addCol(pMap, i, "Barcode", "Barcode", 15, edtNumberAsText)
    Call ColumnMap_addCol(pMap, i, "HS Code", "HSCode", 12)
    Call ColumnMap_addCol(pMap, i, "Name (VN)", "ProductName_vn", 25)
    Call ColumnMap_addCol(pMap, i, "Name (EN)", "ProductName_en", 25)
    Call ColumnMap_addCol(pMap, i, "Image", "ImagePath", 12, edtImage)
    Call ColumnMap_addCol(pMap, i, "Package Type", "PackageType", 12)
    Call ColumnMap_addCol(pMap, i, "Capacity", "CapacityFull", 15)
    Call ColumnMap_addCol(pMap, i, "Price", "Price_vnd", 12, edtCurrency)
    Call ColumnMap_addCol(pMap, i, "Price (+VAT 8%)", "Price_VAT", 12, edtCurrency)
    Call ColumnMap_addCol(pMap, i, "Capacity (OZ)", "Capacity_oz", 12, edtNumber)
    Call ColumnMap_addCol(pMap, i, "Length (CM)", "Length_cm", 8, edtNumber)
    Call ColumnMap_addCol(pMap, i, "Width (CM)", "Width_cm", 8, edtNumber)
    Call ColumnMap_addCol(pMap, i, "Height (CM)", "Height_cm", 8, edtNumber)
    Call ColumnMap_addCol(pMap, i, "Weight (G)", "Weight_gram", 10, edtNumber)
    Call ColumnMap_addCol(pMap, i, "Ingredients", "Ingredients", 50)

    If (langID = 1 Or langID = 3) Then Call ColumnMap_addCol(pMap, i, "USP (VN)", "USP_vn", 35)
    If (langID = 2 Or langID = 3) Then Call ColumnMap_addCol(pMap, i, "USP (EN)", "USP_en", 35)
    If (langID = 1 Or langID = 3) Then Call ColumnMap_addCol(pMap, i, "Usage (VN)", "HowToUse_vn", 35)
    If (langID = 2 Or langID = 3) Then Call ColumnMap_addCol(pMap, i, "Usage (EN)", "HowToUse_en", 35)

    ' 2. Paint Headers & Widths
    Dim col As Integer
    For col = LBound(pMap) To UBound(pMap)
        ws.Cells(1, col).value = pMap(col).Caption
        ws.Columns(col).ColumnWidth = pMap(col).Width

        If pMap(col).DataType = edtCurrency Then
            ws.Columns(col).NumberFormat = "#,##0"
        End If

        If pMap(col).DataType = edtNumberAsText Then
            ws.Columns(col).NumberFormat = "@"
        End If

        If pMap(col).Width >= 25 Then
            ws.Columns(col).WrapText = True
        End If
    Next col

    ' 3. Style Header Row dynamically
    Dim MaxCol As Integer
    MaxCol = UBound(pMap)

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, MaxCol))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(44, 62, 80)
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4108
        .RowHeight = 35
        .WrapText = True
        .Borders.LineStyle = 1
        .Borders.Color = RGB(0, 0, 0)
        .Borders.Weight = 2
    End With

    ' 4. Set Global Cell Alignment
    With ws.Range(ws.Columns(1), ws.Columns(MaxCol))
        .VerticalAlignment = -4108
        .Font.name = "Segoe UI"
        .Font.size = 10
    End With

    ' 5. Freeze Panes (Standard Platinum UI)
    ws.Activate
    ws.Range("A2").Select
    ws.Parent.Windows(1).FreezePanes = True

    ' 6. Return Map to Caller
    Layout_DynamicProductList = pMap
End Function

' Purpose: The Worker that loops through the Recordset and renders data based on the schema map.
Private Sub Render_ProductData(ws As Object, pMap() As ColumnMap, rs As Object)
    ws.Activate ' Ensure sheet is active for In-Cell picture coordinate calculation

    Dim row As Integer: row = 2 ' Start data at row 2

    Do While Not rs.EOF
        ws.Cells(row, 1).value = row - 1

        Dim col As Integer
        For col = 2 To UBound(pMap)
            Dim fldName As String
            fldName = pMap(col).fieldName

            If fldName <> "" Then
                If Not IsNull(rs.Fields(fldName).value) Then
                    ' AUTHORITY FLOW: Explicit Target Cell Binding
                    Dim targetCell As Object
                    Set targetCell = ws.Cells(row, col)

                    Select Case pMap(col).DataType
                        Case edtImage
                            Call InsertProductImage(targetCell, rs.Fields(fldName).value)
                        Case edtNumberAsText
                            targetCell.value = "'" & rs.Fields(fldName).value
                        Case Else
                            Dim cellVal As Variant
                            cellVal = rs.Fields(fldName).value
                            targetCell.value = cellVal
                    End Select
                End If
            End If
        Next col

        row = row + 1
        rs.MoveNext
    Loop
End Sub


' --- 4. HELPERS & FACTORIES (Low-level Utilities) ---

Private Sub ColumnMap_addCol(ByRef pMap() As ColumnMap, ByRef cIdx As Integer, cap As String, fld As String, w As Double, Optional dType As ExportDataType = edtText)
    cIdx = cIdx + 1
    ReDim Preserve pMap(1 To cIdx)
    pMap(cIdx) = ColumnMap_createCol(cIdx, cap, fld, w, dType)
End Sub

Private Function ColumnMap_createCol(cIdx As Integer, cap As String, fld As String, w As Double, Optional dType As ExportDataType = edtText) As ColumnMap
    Dim c As ColumnMap
    c.ColumnIndex = cIdx
    c.Caption = cap
    c.fieldName = fld
    c.Width = w
    c.DataType = dType
    ColumnMap_createCol = c
End Function

' Purpose: Renders a physical image onto the Excel worksheet.
Private Sub InsertProductImage(cell As Object, imgPath As String)
    Dim img As Object
    Dim fullPath As String

    If imgPath = "" Then Exit Sub

    ' 1. Construct Path (Simple & Direct)
    fullPath = imgPath
    If InStr(imgPath, ":") = 0 Then fullPath = CurrentProject.Path & "\" & imgPath

    ' 2. Validation
    If Dir(fullPath) = "" Then Exit Sub

    ' 3. Prepare Environment (Clear + Focus)
    cell.value = ""
    ' AUTHORITY FLOW: Explicit Focus to ensure 'Select' works 100%
    cell.Worksheet.Parent.Activate ' Workbook
    cell.Worksheet.Activate        ' Sheet
    cell.Select                    ' Target Cell

    ' 4. Attempt "Insert into Cell" (Excel 365 Platinum Method)
    On Error GoTo Fallback
    ' Positional parameter for better compatibility with Late Binding
    cell.Application.Selection.InsertPictureInCell fullPath
    On Error GoTo 0
    Exit Sub

Fallback:
    ' 5. Fallback to High-Fidelity Shapes (Older Excel)
    Err.Clear
    On Error GoTo 0

    Set img = cell.Worksheet.Shapes.AddPicture2(fileName:=fullPath, _
                                   LinkToFile:=0, _
                                   SaveWithDocument:=-1, _
                                   Left:=cell.Left + 2, _
                                   Top:=cell.Top + 2, _
                                   Width:=-1, _
                                   Height:=-1, _
                                   Compress:=0)

    With img
        .LockAspectRatio = -1
        If .Height / .Width > cell.Height / cell.Width Then
            .Height = cell.Height - 4
        Else
            .Width = cell.Width - 4
        End If
        .Left = cell.Left + (cell.Width - .Width) / 2
        .Top = cell.Top + (cell.Height - .Height) / 2
        .Placement = 1
    End With
    On Error GoTo 0
End Sub
