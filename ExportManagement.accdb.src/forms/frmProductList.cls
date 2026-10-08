Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'@Folder "App.Products"

'====================================================================
' MODULE    : frmProductList
' PURPOSE   : High-fidelity Grid Viewport for the Product Domain.
' DESIGN    : AUTHORITY FLOW (Head-Heart-Hands)
'
' [FACTS]:
' - ARCHITECTURE : Active Broadcaster ("Remote Control" pattern).
' - PURPOSE      : Pure Viewport - Zero domain logic, only I/O.
' - BINDING      : Dynamic RecordSource binding via Products_Linked.
'====================================================================

'====================================================================
' SECTION 1: LIFECYCLE & VIEW REGISTRATION
'====================================================================

Private Sub Form_Open(Cancel As Integer)
    If Not AppDataHub Is Nothing Then Call AppDataHub.RegisterView(Me)

    ' Set initial row sources
    Me.comboCategory.RowSource = "SELECT MarketingCategory FROM MarketingCategories ORDER BY MarketingCategory;"
    Me.comboMarketingLine.RowSource = "SELECT MarketingLine FROM MarketingLines ORDER BY MarketingLine;"
End Sub

Private Sub Form_Close()
    If Not AppDataHub Is Nothing Then Call AppDataHub.UnregisterView(Me)
End Sub

Private Sub Form_Current()
    On Error Resume Next
    If Not Me.Parent Is Nothing Then
        Call Me.Parent.SyncProductImage
    End If
End Sub

'====================================================================
' SECTION 2: DATA MUTATION & DOMAIN ACTIONS
'====================================================================

Private Sub Form_AfterUpdate()
    On Error GoTo ErrHandler
    If Not AppDataHub Is Nothing Then
        Call AppDataHub.UpdateTimestamp(Me!ProductID)
    End If
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProductList.Form_AfterUpdate: #" & Err.Number & " - " & Err.Description
End Sub

'SoftDelete Feature
Private Sub Form_BeforeDelConfirm(Cancel As Integer, Response As Integer)
    ' Intercept hard-delete and cancel native Access prompt
    Cancel = True
End Sub

Private Sub Form_AfterDelConfirm(Status As Integer)
    On Error GoTo ErrHandler

    Dim idList As String, actualCount As Long
    Call GetSelectedIDs(idList, actualCount)

    If actualCount > 0 And Not AppDataHub Is Nothing Then
        Call AppDataHub.Deactivate(idList)
    End If

    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProductList.Form_AfterDelConfirm: #" & Err.Number & " - " & Err.Description
End Sub

'Restore feature
Private Sub chkContextAction_BeforeUpdate(Cancel As Integer)
    On Error GoTo ErrHandler

    ' Prevent ticking on new uncommitted row
    If Me.NewRecord Then
        Cancel = True: Me.Undo
        Exit Sub
    End If

    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProductList.chkContextAction_BeforeUpdate: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub chkContextAction_AfterUpdate()
    On Error GoTo ErrHandler

    If Me.Dirty Then Me.Dirty = False

    ' If record is inactive, checkbox triggers restoration
    If Nz(Me!isActive, True) = False Then
        If Not AppDataHub Is Nothing Then
            If AppDataHub.Restore(Me!ProductID) Then
                Me.Requery
            End If
        End If
    End If

    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProductList.chkContextAction_AfterUpdate: #" & Err.Number & " - " & Err.Description
End Sub

'====================================================================
' SECTION 3: IN-GRID CONTROL EVENTS
'====================================================================

Private Sub comboCategory_LostFocus()
    Dim categoryVal As String
    Dim sql As String

    categoryVal = Replace(Nz(Me.comboCategory.Text, ""), "'", "''")

    If categoryVal = "" Then
        sql = "SELECT MarketingLine FROM MarketingLines ORDER BY MarketingLine;"
    Else
        sql = "SELECT ML.MarketingLine " & _
              "FROM MarketingLines AS ML " & _
              "INNER JOIN MarketingCategories AS MC ON ML.MarketingCategoryID = MC.MarketingCategoryID " & _
              "WHERE MC.MarketingCategory = '" & categoryVal & "' " & _
              "ORDER BY ML.MarketingLine;"
    End If

    Me.comboMarketingLine.RowSource = sql
End Sub

Private Sub comboMarketingLine_AfterUpdate()
    ' Automatically auto-fill and enforce parent MarketingCategory from selected Line
    On Error GoTo ErrHandler

    Dim selectedLine As String
    selectedLine = Replace(Nz(Me.comboMarketingLine.value, ""), "'", "''")

    If Len(selectedLine) > 0 Then
        Dim catID As Variant
        catID = DLookup("MarketingCategoryID", "MarketingLines", "MarketingLine = '" & selectedLine & "'")

        If Not IsNull(catID) Then
            Dim catName As Variant
            catName = DLookup("MarketingCategory", "MarketingCategories", "MarketingCategoryID = " & catID)

            If Not IsNull(catName) Then
                Me.comboCategory.value = catName
            End If
        End If
    End If
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProductList.comboMarketingLine_AfterUpdate: #" & Err.Number & " - " & Err.Description
End Sub

'====================================================================
' SECTION 4: PRIVATE HELPERS
'====================================================================

Private Sub GetSelectedIDs(ByRef outIdList As String, ByRef outActualCount As Long)
    outIdList = ""
    outActualCount = 0

    Dim selectCount As Long: selectCount = Me.SelHeight
    If selectCount = 0 Then Exit Sub

    Dim sourceRs As DAO.Recordset: Set sourceRs = Me.RecordsetClone
    If sourceRs.BOF And sourceRs.EOF Then Exit Sub

    Dim selTopIndex As Long: selTopIndex = Me.SelTop - 1
    If selTopIndex < 0 Then selTopIndex = 0
    If selTopIndex >= sourceRs.recordCount Then Exit Sub

    sourceRs.AbsolutePosition = selTopIndex

    Dim idArray() As String
    Dim i As Long

    For i = 0 To selectCount - 1
        If sourceRs.EOF Then Exit For

        Dim currentID As Variant
        currentID = sourceRs.Fields("ProductID").value
        If Not IsNull(currentID) Then
            ReDim Preserve idArray(outActualCount)
            idArray(outActualCount) = CStr(currentID)
            outActualCount = outActualCount + 1
        End If
        sourceRs.MoveNext
    Next i

    If outActualCount > 0 Then
        outIdList = Join(idArray, ",")
    End If
End Sub

'====================================================================
' SECTION 5: LOCALE SYNCHRONIZATION
'====================================================================

Public Sub SyncLocale(ByVal CurrentLocale As String)
    On Error GoTo ErrHandler
    Dim isVN As Boolean: isVN = (CurrentLocale = "VN")

    If isVN Then
        Me.txtProductName.Properties("DatasheetCaption") = "Product Name (VN)"
        Me.txtUSP.Properties("DatasheetCaption") = "Unique Selling Point (VN)"
        Me.txtHowToUse.Properties("DatasheetCaption") = "How to Use (VN)"
        Me.txtPrice.Properties("DatasheetCaption") = "Price (VND)"
        Me.txtProductName.ControlSource = "ProductName_vn"
        Me.txtUSP.ControlSource = "USP_vn"
        Me.txtHowToUse.ControlSource = "HowToUse_vn"
        Me.txtPrice.ControlSource = "Price_vnd"
        Me.txtPrice.Format = "#,##0 ""VND"""
    Else
        Me.txtProductName.Properties("DatasheetCaption") = "Product Name (EN)"
        Me.txtUSP.Properties("DatasheetCaption") = "Unique Selling Point (EN)"
        Me.txtHowToUse.Properties("DatasheetCaption") = "How to Use (EN)"
        Me.txtPrice.Properties("DatasheetCaption") = "Price (USD)"
        Me.txtProductName.ControlSource = "ProductName_en"
        Me.txtUSP.ControlSource = "USP_en"
        Me.txtHowToUse.ControlSource = "HowToUse_en"
        Me.txtPrice.ControlSource = "Price_usd"
        Me.txtPrice.Format = "$#,##0.00"
    End If
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProductList.SyncLocale: #" & Err.Number & " - " & Err.Description
End Sub
