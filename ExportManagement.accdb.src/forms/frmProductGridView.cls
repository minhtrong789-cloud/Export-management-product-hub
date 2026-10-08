Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'@Folder "App.Products"

'====================================================================
' FORM      : frmProductGridView
' ROLE      : Grid Workspace Page Viewport (Unbound Container)
'====================================================================

'--- [INTERNAL STATE & ANCHOR CACHE] ---
Private m_AnchorCache    As Object
Private m_SearchFilter   As String
Private m_CategoryFilter As String
Private m_LineFilter     As String

'====================================================================
' FORM LIFECYCLE EVENTS
'====================================================================

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrHandler

    ' 1. Capture anchor cache for safe runtime geometry resizing
    Call AnchorCache_Init

    ' 2. Reset combo filter selections
    Me.cboFilterCategory = Null
    Me.cboFilterLine = Null

    ' 3. Populate filter dropdown row sources
    Me.cboFilterCategory.RowSource = "SELECT MarketingCategory, MarketingCategoryID FROM MarketingCategories ORDER BY MarketingCategory"
    Me.cboFilterLine.RowSource = "SELECT MarketingLine, MarketingLineID FROM MarketingLines ORDER BY MarketingLine"

    ' 4. Image control is always visible (BUG IF TRY TO INVISIBLE)
    Me.imgProduct.Visible = True
'
    Exit Sub

ErrHandler:
    MsgBox "Error initializing Product Grid View: " & Err.Description, vbCritical
    Cancel = True
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler

    ' 1. Ensure subform container points to frmProductList
    If Me.sfProductGrid.SourceObject <> "frmProductList" Then
        Me.sfProductGrid.SourceObject = "frmProductList"
    End If

    ' 2. Initial binding to active products
    Me.sfProductGrid.Form.RecordSource = "SELECT * FROM Products_Linked WHERE IsActive = True ORDER BY ProductID;"

    ' 3. Apply baseline layout from Parent form state
    Dim isMax As Boolean
    Dim isMed As Boolean
    On Error Resume Next
    isMax = Me.Parent.isMaximized
    isMed = Me.Parent.IsMediaVisible
    On Error GoTo ErrHandler

    Call ApplyLayout(isMax, isMed)

    ' 4. Initial sync of product image and footer summary
    Call UpdateSummary

    Exit Sub

ErrHandler:
    MsgBox "Error loading Product Grid View: " & Err.Description, vbCritical
End Sub

'====================================================================
' PUBLIC VIEW HELPERS & GEOMETRY
'====================================================================

' Purpose: Safe geometry resize for 4 layout scenarios (Maximized x MediaVisible).
Public Sub ApplyLayout(ByVal isMaximized As Boolean, ByVal showMedia As Boolean)
    On Error GoTo ErrHandler
    Application.Echo False

    Const TWIPS_PER_INCH As Long = 1440
    Const formWidth As Double = 10.9792
    Const formHeight As Double = 5.5

    ' 1. Shift focus away from potential hidden controls to prevent Error #2165
    On Error Resume Next
    Me.cboFilterCategory.SetFocus
    On Error GoTo ErrHandler

    Dim showActions As Boolean: showActions = Not isMaximized
    'Me.cboFilterCategory.Visible = showActions
    Me.cboFilterLine.Visible = showActions
    Me.linMainDivider.Visible = showActions
    Me.cmdNewProduct.Visible = showActions
    Me.cmdExport.Visible = showActions
    Me.linFooterSeparator.Visible = showActions
    Me.lblRecordCount.Visible = showActions

    Me.cmdAddImage.Visible = showMedia
    Me.cmdClearImage.Visible = showMedia

    ' 2. Freeze anchors to TopLeft (Top-Left = 0)
    Call SetAnchorStateTopLeft

    ' 3. Assert canvas base bounds
    Me.Width = formWidth * TWIPS_PER_INCH
    Me.Section(acDetail).Height = formHeight * TWIPS_PER_INCH

    ' 4. Dynamic grid coordinates calculation (IIf style)
    Dim gTop As Single, gLeft As Single, gWidth As Single, gHeight As Single
    gTop = IIf(isMaximized, 0#, 0.55) * TWIPS_PER_INCH
    gLeft = IIf(isMaximized, 0#, 0.2) * TWIPS_PER_INCH
    gHeight = IIf(isMaximized, 5.5, 4.35) * TWIPS_PER_INCH
    gWidth = IIf(showMedia, IIf(isMaximized, 5.4, 4.9), IIf(isMaximized, 10.9692, 9.3)) * TWIPS_PER_INCH

    Me.sfProductGrid.Move gLeft, gTop, gWidth, gHeight

    ' Media block geometry
    If isMaximized And showMedia Then
        Me.imgProduct.Move 5.5 * TWIPS_PER_INCH, 0.1 * TWIPS_PER_INCH, 5.35 * TWIPS_PER_INCH, 4.8 * TWIPS_PER_INCH
        Me.LabelNoImage.Move 5.5 * TWIPS_PER_INCH, 2# * TWIPS_PER_INCH, 5.35 * TWIPS_PER_INCH, 1# * TWIPS_PER_INCH
        Me.cmdAddImage.Move 5.5 * TWIPS_PER_INCH, 5# * TWIPS_PER_INCH, 2.6 * TWIPS_PER_INCH, 0.35 * TWIPS_PER_INCH
        Me.cmdClearImage.Move 8.25 * TWIPS_PER_INCH, 5# * TWIPS_PER_INCH, 2.6 * TWIPS_PER_INCH, 0.35 * TWIPS_PER_INCH
    Else
        Me.imgProduct.Move 5.3 * TWIPS_PER_INCH, 0.55 * TWIPS_PER_INCH, 4.2 * TWIPS_PER_INCH, 3.5 * TWIPS_PER_INCH
        Me.LabelNoImage.Move 5.3 * TWIPS_PER_INCH, 1.55 * TWIPS_PER_INCH, 4.2 * TWIPS_PER_INCH, 0.67 * TWIPS_PER_INCH
        Me.cmdAddImage.Move 5.3 * TWIPS_PER_INCH, 4.15 * TWIPS_PER_INCH, 1.95 * TWIPS_PER_INCH, 0.3 * TWIPS_PER_INCH
        Me.cmdClearImage.Move 7.45 * TWIPS_PER_INCH, 4.15 * TWIPS_PER_INCH, 2.05 * TWIPS_PER_INCH, 0.3 * TWIPS_PER_INCH
    End If
'
    Me.Width = formWidth * TWIPS_PER_INCH
    Me.Section(acDetail).Height = formHeight * TWIPS_PER_INCH

    ' 5. Restore original anchor states
    Call ReturnAnchorState(m_AnchorCache)

    ' 6. Control visibility toggles and sync image
    If showMedia Then
        Me.imgProduct.Visible = True
        Call SyncProductImage
    Else
        Me.imgProduct.Visible = False
        Me.LabelNoImage.Visible = False
    End If

    Application.Echo True
    Exit Sub

ErrHandler:
    Application.Echo True
    Debug.Print "[ERROR] frmProductGridView.ApplyLayout: #" & Err.Number & " - " & Err.Description
End Sub

' Purpose: Synchronize image display for currently selected product in grid.
Public Sub SyncProductImage()
    On Error GoTo FallbackNoImage

    ' 1. Guard check subform & record availability
    Dim imgPath As String: imgPath = ""
    If Not Me.sfProductGrid.Form Is Nothing Then
        If Not (Me.sfProductGrid.Form.Recordset.BOF And Me.sfProductGrid.Form.Recordset.EOF) Then
            imgPath = Nz(Me.sfProductGrid.Form!ImagePath, "")
        End If
    End If

    ' 2. Build full file system path
    Dim fullPath As String: fullPath = ""
    If Len(imgPath) > 0 Then
        If Left$(imgPath, 1) = "\" Then
            fullPath = CurrentProject.Path & imgPath
        Else
            fullPath = CurrentProject.Path & "\" & imgPath
        End If
    End If

    ' 3. Validate existence & Load Picture
    If Len(fullPath) > 0 And Len(Dir(fullPath)) > 0 Then
        Me.imgProduct.Picture = fullPath
        Me.LabelNoImage.Visible = False
        Exit Sub
    End If

FallbackNoImage:
    ' Graceful fallback: clear image and display watermark label
    Me.imgProduct.Picture = ""
    Me.LabelNoImage.Visible = True
End Sub

' Purpose: Update record count on footer strip.
Public Sub UpdateSummary()
    On Error GoTo ErrHandler

    Dim recordCount As Long: recordCount = 0
    If Not Me.sfProductGrid.Form Is Nothing Then
        If Not Me.sfProductGrid.Form.Recordset Is Nothing Then
            recordCount = Me.sfProductGrid.Form.Recordset.recordCount
        End If
    End If

    Me.lblRecordCount.Caption = "Showing: " & Format$(recordCount, "#,##0") & " products"
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProductGridView.UpdateSummary: #" & Err.Number & " - " & Err.Description
End Sub

' Purpose: Synchronize grid RecordSource and column capabilities for Live vs Trashbin mode.
Public Sub SyncDataMode(ByVal isTrash As Boolean)
    On Error GoTo ErrHandler
    Application.Echo False

    If Not Me.sfProductGrid.Form Is Nothing Then
        Dim strSql As String
        strSql = "SELECT * FROM Products_Linked WHERE IsActive = " & IIf(isTrash, "False", "True") & " ORDER BY ProductID;"
        Me.sfProductGrid.Form.RecordSource = strSql

        On Error Resume Next
        Me.sfProductGrid.Form.Controls("chkContextAction").Properties("DatasheetCaption") = IIf(isTrash, "Restore", "Selected")
        Me.sfProductGrid.Form.AllowAdditions = Not isTrash
        Me.sfProductGrid.Form.AllowDeletions = Not isTrash
        On Error GoTo ErrHandler
    End If

    Call UpdateSummary
    Application.Echo True
    Exit Sub

ErrHandler:
    Application.Echo True
    Debug.Print "[ERROR] frmProductGridView.SyncDataMode: #" & Err.Number & " - " & Err.Description
End Sub

' Purpose: Propagates locale change to child product list grid.
Public Sub SyncLocale(ByVal CurrentLocale As String)
    On Error GoTo ErrHandler
    If Not Me.sfProductGrid.Form Is Nothing Then
        Call Me.sfProductGrid.Form.SyncLocale(CurrentLocale)
    End If
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProductGridView.SyncLocale: #" & Err.Number & " - " & Err.Description
End Sub

'====================================================================
' SEARCH & FILTER ORCHESTRATION
'====================================================================

Public Sub ApplySearch(ByVal keyword As String)
    On Error GoTo ErrHandler
    If Len(Trim(keyword)) = 0 Then
        m_SearchFilter = ""
    Else
        If Not Me.sfProductGrid.Form Is Nothing Then
            m_SearchFilter = BuildQuickSearchFilter(Me.sfProductGrid.Form.RecordsetClone, keyword)
        Else
            m_SearchFilter = ""
        End If
    End If
    Call RebuildFilter
    Exit Sub
ErrHandler:
    Debug.Print "[ERROR] frmProductGridView.ApplySearch: #" & Err.Number & " - " & Err.Description
End Sub

Public Sub RebuildFilter()
    On Error GoTo ErrHandler
    Dim fullFilter As String: fullFilter = ""

    If m_CategoryFilter <> "" Then fullFilter = fullFilter & IIf(fullFilter <> "", " AND ", "") & m_CategoryFilter
    If m_LineFilter <> "" Then fullFilter = fullFilter & IIf(fullFilter <> "", " AND ", "") & m_LineFilter
    If m_SearchFilter <> "" Then fullFilter = fullFilter & IIf(fullFilter <> "", " AND ", "") & "(" & m_SearchFilter & ")"

    If Not Me.sfProductGrid.Form Is Nothing Then
        If fullFilter <> "" Then
            Me.sfProductGrid.Form.Filter = fullFilter
            Me.sfProductGrid.Form.FilterOn = True
        Else
            Me.sfProductGrid.Form.Filter = ""
            Me.sfProductGrid.Form.FilterOn = False
        End If
    End If

    Call UpdateSummary
    Exit Sub
ErrHandler:
    Debug.Print "[ERROR] frmProductGridView.RebuildFilter: #" & Err.Number & " - " & Err.Description
End Sub

Public Sub ClearFilters()
    On Error GoTo ErrHandler
    Me.cboFilterCategory = Null
    Me.cboFilterLine = Null
    Me.cboFilterLine.RowSource = "SELECT MarketingLine, MarketingLineID FROM MarketingLines ORDER BY MarketingLine"
    m_CategoryFilter = ""
    m_LineFilter = ""
    m_SearchFilter = ""
    Call RebuildFilter
    Exit Sub
ErrHandler:
    Debug.Print "[ERROR] frmProductGridView.ClearFilters: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub cboFilterCategory_AfterUpdate()
    On Error GoTo ErrHandler
    Dim catID As Long: catID = Nz(Me.cboFilterCategory.Column(1), 0)
    If catID = 0 Or IsNull(Me.cboFilterCategory) Then
        m_CategoryFilter = ""
        Me.cboFilterLine.RowSource = "SELECT MarketingLine, MarketingLineID FROM MarketingLines ORDER BY MarketingLine"
    Else
        m_CategoryFilter = "MarketingCategoryID = " & catID
        Me.cboFilterLine.RowSource = "SELECT MarketingLine, MarketingLineID FROM MarketingLines WHERE MarketingCategoryID = " & catID & " ORDER BY MarketingLine"
    End If
    Call RebuildFilter
    Exit Sub
ErrHandler:
    Debug.Print "[ERROR] frmProductGridView.cboFilterCategory_AfterUpdate: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub cboFilterLine_AfterUpdate()
    On Error GoTo ErrHandler
    Dim lineID As Long: lineID = Nz(Me.cboFilterLine.Column(1), 0)
    If lineID = 0 Or IsNull(Me.cboFilterLine) Then
        m_LineFilter = ""
    Else
        m_LineFilter = "MarketingLineID = " & lineID
    End If
    Call RebuildFilter
    Exit Sub
ErrHandler:
    Debug.Print "[ERROR] frmProductGridView.cboFilterLine_AfterUpdate: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub cboFilterCategory_NotInList(NewData As String, Response As Integer)
    Response = acDataErrContinue
End Sub

Private Sub cboFilterLine_NotInList(NewData As String, Response As Integer)
    Response = acDataErrContinue
End Sub

'====================================================================
' CLUSTER 2: DATA MANIPULATION (NEW & SELECT ALL)
'====================================================================

Private Sub cmdNewProduct_Click()
    On Error GoTo ErrHandler
    Me.sfProductGrid.SetFocus
    DoCmd.RunCommand acCmdRecordsGoToNew
    Exit Sub
ErrHandler:
    Debug.Print "[ERROR] frmProductGridView.cmdNewProduct_Click: #" & Err.Number & " - " & Err.Description
End Sub

Public Sub SelectAll()
    ' PURPOSE: Toggle selection state exclusively for active records in the current window viewport.
    On Error GoTo ErrHandler

    ' 1. Guard clause: Do not execute in Trashbin mode
    Dim isArchive As Boolean: isArchive = False
    On Error Resume Next
    isArchive = Me.Parent.IsTrashbin
    On Error GoTo ErrHandler
    If isArchive Then Exit Sub

    ' 2. Build filter strictly for active window records
    Dim strFilter As String
    If Me.sfProductGrid.Form.FilterOn And Len(Trim(Me.sfProductGrid.Form.Filter)) > 0 Then
        strFilter = "(" & Me.sfProductGrid.Form.Filter & ") AND IsActive=True"
    Else
        strFilter = "IsActive=True"
    End If

    ' 3. Toggle decision: If at least one row is unselected -> Select All, otherwise Deselect All
    Dim unselectedCount As Long
    unselectedCount = DCount("*", "Products_Linked", strFilter & " AND chkContextAction=False")

    Dim targetState As String
    targetState = IIf(unselectedCount > 0, "True", "False")

    ' 4. Execute atomic update on Staging table
    Dim strSql As String
    strSql = "UPDATE [Products_Linked] SET [chkContextAction]=" & targetState & " WHERE " & strFilter & ";"

    Dim db As DAO.Database: Set db = CurrentDb
    Call db.Execute(strSql, dbFailOnError)

    ' 5. Refresh grid UI
    Me.sfProductGrid.Form.Requery
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProductGridView.SelectAll: #" & Err.Number & " - " & Err.Description
End Sub

'====================================================================
' CLUSTER 3: MEDIA ASSET OPERATIONS (ADD & CLEAR IMAGE)
'====================================================================

Public Sub cmdAddImage_Click()
    On Error GoTo ErrHandler

    Dim targetProductID As Long
    targetProductID = Nz(Me.sfProductGrid.Form!ProductID, 0)

    If targetProductID = 0 Or Me.sfProductGrid.Form.NewRecord Then
        MsgBox "Please save the product record before attaching an image.", vbInformation, "Record Not Saved"
        Exit Sub
    End If

    Dim fileDialog As Object
    Set fileDialog = Application.fileDialog(3) ' msoFileDialogFilePicker
    With fileDialog
        .Title = "Select Product Image"
        .Filters.Clear
        .Filters.Add "Images (*.jpg, *.jpeg, *.png, *.gif)", "*.jpg; *.jpeg; *.png; *.gif", 1
        If .Show = 0 Then Exit Sub
    End With

    If Not AppDataHub Is Nothing Then
        Dim selectedPath As String: selectedPath = fileDialog.SelectedItems(1)
        Dim targetBarcode As String: targetBarcode = Nz(Me.sfProductGrid.Form!Barcode, "IMG_" & targetProductID)
        Call AppDataHub.AddImage(targetProductID, selectedPath, targetBarcode)
    End If

    Me.sfProductGrid.Form.Refresh
    Call SyncProductImage
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProductGridView.cmdAddImage_Click: #" & Err.Number & " - " & Err.Description
End Sub

Public Sub cmdClearImage_Click()
    On Error GoTo ErrHandler

    If Me.sfProductGrid.Form.NewRecord Then Exit Sub

    Dim targetProductID As Long
    targetProductID = Nz(Me.sfProductGrid.Form!ProductID, 0)
    If targetProductID = 0 Then Exit Sub

    If Not AppDataHub Is Nothing Then
        Call AppDataHub.RemoveImage(targetProductID)
    End If

    Me.sfProductGrid.Form.Refresh
    Call SyncProductImage
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProductGridView.cmdClearImage_Click: #" & Err.Number & " - " & Err.Description
End Sub

'====================================================================
' CLUSTER 4: EXCEL EXPORT & TEMPLATE OPERATIONS
'====================================================================

Private Sub cmdExport_Click()
    On Error GoTo ErrHandler

    Me.cmdExport.Caption = "..."
    Me.cmdExport.Enabled = False
    DoEvents

    Dim rsSource As DAO.Recordset
    Dim selectedIDs As String: selectedIDs = ""
    Dim rs As DAO.Recordset
    Dim strSql As String
    Dim isConfirmed As Boolean
    Dim langID As Integer
    Const DLG_NAME As String = "dlgLanguagePicker"

    Set rsSource = Me.sfProductGrid.Form.RecordsetClone
    rsSource.Filter = "[chkContextAction] = True"
    Set rsSource = rsSource.OpenRecordset

    Do While Not rsSource.EOF
        selectedIDs = selectedIDs & rsSource!ProductID & ","
        rsSource.MoveNext
    Loop

    If selectedIDs = "" Then
        MsgBox "Please select at least one product checkbox to export.", vbExclamation, "No Selection"
        GoTo Cleanup
    End If

    selectedIDs = Left$(selectedIDs, Len(selectedIDs) - 1)

    strSql = "SELECT *, " & _
             "[MarketingCategory] & ' - ' & [MarketingLine] AS CategoryFull, " & _
             "[Capacity] & ' ' & [CapacityUnit] AS CapacityFull, " & _
             "[Price_vnd] * 1.08 AS Price_VAT " & _
             "FROM [Products_Linked] " & _
             "WHERE [ProductID] IN (" & selectedIDs & ");"

    Set rs = CurrentDb.OpenRecordset(strSql)

    DoCmd.OpenForm DLG_NAME, WindowMode:=acDialog
    isConfirmed = CurrentProject.AllForms(DLG_NAME).IsLoaded
    If Not isConfirmed Then GoTo Cleanup

    langID = Forms(DLG_NAME).fraLang.value
    DoCmd.Close acForm, DLG_NAME

    Call Print_ProductList(langID, rs)

Cleanup:
    Me.cmdExport.Caption = "EXPORT EXCEL"
    Me.cmdExport.Enabled = True
    If Not rs Is Nothing Then rs.Close: Set rs = Nothing
    If Not rsSource Is Nothing Then rsSource.Close: Set rsSource = Nothing
    Exit Sub

ErrHandler:
    Me.cmdExport.Caption = "EXPORT EXCEL"
    Me.cmdExport.Enabled = True
    Debug.Print "[ERROR] frmProductGridView.cmdExport_Click: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrHandler

    ' Save pending edits on child grid if dirty
    If Not Me.sfProductGrid.Form Is Nothing Then
        If Me.sfProductGrid.Form.Dirty Then Me.sfProductGrid.Form.Dirty = False
    End If

    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProductGridView.Form_Unload: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub Form_Close()
    On Error Resume Next
    Me.imgProduct.Picture = ""
    Set m_AnchorCache = Nothing
End Sub

'====================================================================
' INTERNAL GEOMETRIC LAYOUT & ANCHOR HELPERS
'====================================================================

Private Sub AnchorCache_Init()
    Set m_AnchorCache = CreateObject("Scripting.Dictionary")

    Dim ctrl As Access.Control
    For Each ctrl In Me.Controls
        Select Case ctrl.controlType
            Case acLabel, acTextBox, acComboBox, acListBox, _
                 acCommandButton, acToggleButton, acCheckBox, acOptionButton, acOptionGroup, _
                 acSubform, acImage, acRectangle, acLine, _
                 acTabCtl, acAttachment, acBoundObjectFrame, acObjectFrame
                m_AnchorCache(ctrl.name & "_H") = ctrl.HorizontalAnchor
                m_AnchorCache(ctrl.name & "_V") = ctrl.VerticalAnchor
            Case acSubform
        End Select
    Next ctrl
End Sub

Private Sub SetAnchorStateTopLeft()
    Dim ctrl As Access.Control
    For Each ctrl In Me.Controls
        Select Case ctrl.controlType
            Case acLabel, acTextBox, acComboBox, acListBox, _
                 acCommandButton, acToggleButton, acCheckBox, acOptionButton, acOptionGroup, _
                 acImage, acRectangle, acLine, _
                 acTabCtl, acAttachment, acBoundObjectFrame, acObjectFrame
                ctrl.HorizontalAnchor = 0
                ctrl.VerticalAnchor = 0
            Case acSubform
        End Select

    Next ctrl
End Sub

Private Sub ReturnAnchorState(ByRef anchorCache As Object)
    If anchorCache Is Nothing Then Exit Sub

    Dim ctrl As Access.Control
    For Each ctrl In Me.Controls
        Select Case ctrl.controlType
            Case acLabel, acTextBox, acComboBox, acListBox, _
                 acCommandButton, acToggleButton, acCheckBox, acOptionButton, acOptionGroup, _
                 acImage, acRectangle, acLine, _
                 acTabCtl, acAttachment, acBoundObjectFrame, acObjectFrame
                If anchorCache.exists(ctrl.name & "_H") Then
                    ctrl.HorizontalAnchor = anchorCache(ctrl.name & "_H")
                    ctrl.VerticalAnchor = anchorCache(ctrl.name & "_V")
                End If
            Case acSubform
        End Select
    Next ctrl
End Sub
