Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'@Folder "App.Products"


'====================================================================
' FORM      : frmProductDetailView
' ROLE      : Product Detail Viewport Page (3-Column Layout & Media)
'====================================================================

'====================================================================
' FORM LIFECYCLE EVENTS
'====================================================================

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrHandler

    ' 1. Populate dropdown row sources
    Me.comboMarketingCategory.RowSource = "SELECT MarketingCategory, MarketingCategoryID FROM MarketingCategories ORDER BY MarketingCategory"
    Me.comboMarketingLine.RowSource = "SELECT MarketingLine, MarketingLineID FROM MarketingLines ORDER BY MarketingLine"
    Me.cboPackageType.RowSource = "SELECT PackageType FROM PackagingTypes ORDER BY PackageType"
    Me.cboCapacityUnit.RowSource = "SELECT CapacityUnit FROM CapacityUnits ORDER BY CapacityUnit"

    ' 2. Image control is always visible (BUG IF TRY TO INVISIBLE)
    Me.imgProduct.Visible = True

    Exit Sub

ErrHandler:
    MsgBox "Error initializing Product Detail View: " & Err.Description, vbCritical
    Cancel = True
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler

    ' 1. Initial binding to active products
    Me.RecordSource = "SELECT * FROM Products_Linked WHERE IsActive = True ORDER BY ProductID"

    ' 2. Initial sync of product image and navigation status
    Call SyncProductImage
    Call UpdateNavigation

    ' 3. Initial sync locale from parent shell
    Dim parentLocale As String: parentLocale = "EN"
    On Error Resume Next
    parentLocale = Me.Parent.CurrentLocale
    On Error GoTo ErrHandler
    Call SyncLocale(parentLocale)

    Exit Sub

ErrHandler:
    MsgBox "Error loading Product Detail View: " & Err.Description, vbCritical
End Sub

Private Sub Form_Current()
    Call UpdateNavigation
    Call SyncProductImage
End Sub

'====================================================================
' PUBLIC VIEW HELPERS
'====================================================================

' Purpose: Synchronize multilingual field bindings and labels (EN vs VN).
Public Sub SyncLocale(ByVal CurrentLocale As String)
    On Error GoTo ErrHandler
    Dim isVN As Boolean: isVN = (CurrentLocale = "VN")

    If isVN Then
        Me.txtUSP.ControlSource = "USP_vn"
        Me.txtHowToUse.ControlSource = "HowToUse_vn"
        Me.lblUSP.Caption = "USP / Highlights (VN)"
        Me.lblHowToUse.Caption = "How To Use (VN)"
    Else
        Me.txtUSP.ControlSource = "USP_en"
        Me.txtHowToUse.ControlSource = "HowToUse_en"
        Me.lblUSP.Caption = "USP / Highlights (EN)"
        Me.lblHowToUse.Caption = "How To Use (EN)"
    End If
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProductDetailView.SyncLocale: #" & Err.Number & " - " & Err.Description
End Sub

' Purpose: Synchronize detail RecordSource for Live vs Trashbin mode.
Public Sub SyncDataMode(ByVal isTrash As Boolean)
    On Error GoTo ErrHandler
    Me.RecordSource = "SELECT * FROM Products_Linked WHERE IsActive = " & IIf(isTrash, "False", "True") & " ORDER BY ProductID"
    Me.AllowAdditions = Not isTrash
    Me.AllowDeletions = Not isTrash
    Call UpdateNavigation
    Call SyncProductImage
    Exit Sub
ErrHandler:
    Debug.Print "[ERROR] frmProductDetailView.SyncDataMode: #" & Err.Number & " - " & Err.Description
End Sub

' Purpose: Synchronize image display for active record.
Public Sub SyncProductImage()
    On Error GoTo FallbackNoImage

    ' 1. Guard check active record
    Dim imgPath As String: imgPath = ""
    If Not (Me.Recordset.BOF And Me.Recordset.EOF) Then
        imgPath = Nz(Me!ImagePath, "")
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

' Purpose: Update navigation record status and Prev/Next button states.
Public Sub UpdateNavigation()
    On Error GoTo ErrHandler

    Dim currRec As Long, totalRecs As Long
    Dim rs As DAO.Recordset

    Set rs = Me.RecordsetClone
    totalRecs = 0
    If Not rs Is Nothing Then
        If Not (rs.BOF And rs.EOF) Then
            rs.MoveLast
            totalRecs = rs.recordCount
        End If
    End If

    currRec = Me.CurrentRecord

    If totalRecs = 0 Then
        Me.lblNavStatus.Caption = "Product 0 of 0"
        Me.cmdPrev.Enabled = False
        Me.cmdNext.Enabled = False
    ElseIf Me.NewRecord Then
        Me.lblNavStatus.Caption = "New Product"
        Me.cmdPrev.Enabled = True
        Me.cmdNext.Enabled = False
    Else
        Me.lblNavStatus.Caption = "Product " & currRec & " of " & totalRecs
        Me.cmdPrev.Enabled = (currRec >= 2)
        Me.cmdNext.Enabled = (currRec < totalRecs)
    End If

    Set rs = Nothing
    Exit Sub

ErrHandler:
    Set rs = Nothing
End Sub

'====================================================================
' CONTROL EVENTS (NAVIGATION & DROPDOWNS)
'====================================================================

Private Sub cmdPrev_Click()
    On Error GoTo ErrHandler

    If Me.CurrentRecord >= 2 Or Me.NewRecord Then
        DoCmd.GoToRecord , , acPrevious
    End If
    Exit Sub

ErrHandler:
    If Err.Number <> 2105 Then
        Debug.Print "[ERROR] frmProductDetailView.cmdPrev_Click: #" & Err.Number & " - " & Err.Description
    End If
End Sub

Private Sub cmdNext_Click()
    On Error GoTo ErrHandler

    DoCmd.GoToRecord , , acNext
    Exit Sub

ErrHandler:
    If Err.Number <> 2105 Then
        Debug.Print "[ERROR] frmProductDetailView.cmdNext_Click: #" & Err.Number & " - " & Err.Description
    End If
End Sub

Private Sub cmdNewProduct_Click()
    On Error GoTo ErrHandler

    DoCmd.GoToRecord , , acNewRec
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProductDetailView.cmdNewProduct_Click: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub comboMarketingCategory_AfterUpdate()
    On Error GoTo ErrHandler

    Dim categoryVal As String: categoryVal = Replace(Nz(Me.comboMarketingCategory.value, ""), "'", "''")

    If categoryVal = "" Then
        Me.comboMarketingLine.RowSource = "SELECT MarketingLine FROM MarketingLines ORDER BY MarketingLine"
    Else
        Me.comboMarketingLine.RowSource = "SELECT ML.MarketingLine FROM MarketingLines AS ML " & _
                                          "INNER JOIN MarketingCategories AS MC ON ML.MarketingCategoryID = MC.MarketingCategoryID " & _
                                          "WHERE MC.MarketingCategory = '" & categoryVal & "' ORDER BY ML.MarketingLine"
    End If
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProductDetailView.comboMarketingCategory_AfterUpdate: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub cmdAddImage_Click()
    On Error GoTo ErrHandler

    Dim targetProductID As Long: targetProductID = Nz(Me!ProductID, 0)
    If targetProductID = 0 Or Me.NewRecord Then
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
        Dim targetBarcode As String: targetBarcode = Nz(Me!Barcode, "IMG_" & targetProductID)
        Call AppDataHub.AddImage(targetProductID, selectedPath, targetBarcode)
    End If

    Me.Refresh
    Call SyncProductImage
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProductDetailView.cmdAddImage_Click: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub cmdClearImage_Click()
    On Error GoTo ErrHandler

    If Me.NewRecord Then Exit Sub

    Dim targetProductID As Long: targetProductID = Nz(Me!ProductID, 0)
    If targetProductID = 0 Then Exit Sub

    If Not AppDataHub Is Nothing Then
        Call AppDataHub.RemoveImage(targetProductID)
    End If

    Me.Refresh
    Call SyncProductImage
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProductDetailView.cmdClearImage_Click: #" & Err.Number & " - " & Err.Description
End Sub



Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrHandler

    ' Save pending edits if record is dirty before view closes
    If Me.Dirty Then Me.Dirty = False

    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProductDetailView.Form_Unload: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub Form_Close()
    On Error Resume Next
    ' Release image resource
    Me.imgProduct.Picture = ""
End Sub
