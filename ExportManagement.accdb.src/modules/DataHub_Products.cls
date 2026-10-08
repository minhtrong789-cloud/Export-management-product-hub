VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataHub_Products"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit
'@Folder "App.Products"

'====================================================================
' CLASS     : DataHub_Products
' PURPOSE   : Domain data orchestrator for Product Hub.
'             Coordinates UI Views, StagingBuffer, MediaStorage,
'             and Production Backend Synchronization.
'====================================================================

' INTERNAL INFRASTRUCTURE & SERVICES
Private Const EXCLUDE_FIELDS    As String = "RowNum,chkContextAction,ChangeDate,InsertDate,IsActive"
Private m_Buffer                As StagingBuffer
Private m_Media                 As MediaStorage
Private m_Views                 As Scripting.Dictionary
Private m_RecordSourceHistory   As Scripting.Dictionary
Private m_NumberOfModified      As Long
Private m_NumberOfNew           As Long

'====================================================================
' LIFECYCLE & VIEW MANAGEMENT
'====================================================================

Private Sub Class_Initialize()
    ' PURPOSE: Instantiates internal services (MediaStorage, Views dictionary).
    Set m_Media = New MediaStorage
    Set m_Views = New Scripting.Dictionary
    m_Views.CompareMode = TextCompare
    Set m_RecordSourceHistory = New Scripting.Dictionary
    m_RecordSourceHistory.CompareMode = TextCompare
End Sub

Public Sub Init(ByVal BufferInstance As StagingBuffer)
    ' PURPOSE: Injects shared StagingBuffer instance.
    Set m_Buffer = BufferInstance
End Sub

Private Sub Class_Terminate()
    ' PURPOSE: Releases internal service handles and object references.
    On Error Resume Next
    Set m_Views = Nothing
    Set m_RecordSourceHistory = Nothing
    Set m_Buffer = Nothing
    Set m_Media = Nothing
End Sub

Public Sub RegisterView(ByVal FormObject As Object)
    ' PURPOSE: Registers an active Form/Subform to be coordinated by this hub.
    If FormObject Is Nothing Then Exit Sub
    Set m_Views(CStr(ObjPtr(FormObject))) = FormObject
End Sub

Public Sub UnregisterView(ByVal FormObject As Object)
    ' PURPOSE: Unregisters a closing Form/Subform from active tracking.
    If FormObject Is Nothing Then Exit Sub
    Dim ptrKey As String: ptrKey = CStr(ObjPtr(FormObject))
    If m_Views.exists(ptrKey) Then m_Views.Remove ptrKey
    If m_RecordSourceHistory.exists(ptrKey) Then m_RecordSourceHistory.Remove ptrKey
End Sub

Public Sub TerminateClass()
    ' PURPOSE: Prompts user to save pending changes if dirty, then releases locks.
    On Error Resume Next
    Call TryCommit(Silent:=False)
    Call ReleaseLocks
    If Not m_Buffer Is Nothing Then
        Call m_Buffer.UnlinkTable("Products_Linked")
    End If
    Set m_Views = Nothing
    Set m_RecordSourceHistory = Nothing
    Set m_Buffer = Nothing
    Set m_Media = Nothing
End Sub

'====================================================================
' DATA PIPELINE & COMMIT WORKFLOW
'====================================================================

Public Sub RefreshData()
    ' PURPOSE: Reloads staging table from host query, configures UI columns/indexes,
    '          resets change baseline, and purges orphan media.
    On Error GoTo ErrHandler

    ' 1. DISCONNECT BOUND FORMS BEFORE SCHEMA MUTATIONS
    Call ReleaseLocks

    ' 2. RE-INITIALIZE STAGING TABLE AND COMPUTE BASELINE SNAPSHOT
    Call m_Buffer.InitializeStaging("Source_Products", "Products", "Products_Linked", EXCLUDE_FIELDS)

    ' 3. ADD PRODUCT-SPECIFIC UI COLUMN AND UNIQUE INDEX
    On Error Resume Next
    Call m_Buffer.Database.Execute("ALTER TABLE [Products] ADD COLUMN [chkContextAction] BIT;", dbFailOnError)
    Call m_Buffer.Database.Execute("CREATE UNIQUE INDEX idx_Barcode ON [Products] (Barcode);", dbFailOnError)
    On Error GoTo ErrHandler

    ' 4. RESET DIRTY COUNTERS
    m_NumberOfModified = 0
    m_NumberOfNew = 0

    ' 5. REBIND BOUND FORMS TO LATEST DATA
    Call RestoreLocks

    ' 6. 2-WAY ORPHAN MEDIA CLEANUP
    If Not m_Media Is Nothing Then
        Call m_Media.ClearOrphanImageFiles("Products_Linked", "ImagePath")
        Call m_Media.ClearOrphanImagePaths("Products_Linked", "ImagePath")
    End If
    Exit Sub

ErrHandler:
    Call RestoreLocks
    Debug.Print "[ERROR] DataHub_Product.RefreshData: #" & Err.Number & " - " & Err.Description
End Sub

Public Function TryCommit(Optional ByVal Silent As Boolean = False) As Boolean
    ' PURPOSE: Main gatekeeper for persisting pending changes to production database.
    On Error GoTo ErrHandler

    Call m_Buffer.UpdateAfterHash("Products_Linked", EXCLUDE_FIELDS)
    m_NumberOfNew = m_Buffer.IsDataNew("Products_Linked")
    m_NumberOfModified = m_Buffer.IsDataChanged("Products_Linked")

    Dim IsDirty As Boolean
    IsDirty = (m_NumberOfModified > 0 Or m_NumberOfNew > 0)

    If Not IsDirty Then
        TryCommit = True
        Exit Function
    End If

    If Silent Then
        Call MakeChanges
        TryCommit = True
        Exit Function
    End If

    Dim msgText As String
    msgText = "UNSAVED CHANGES DETECTED" & vbCrLf & vbCrLf & _
              "The following pending changes were found in your workspace:"
    If m_NumberOfModified > 0 Then msgText = msgText & vbCrLf & " * Modified: " & m_NumberOfModified & " record(s)"
    If m_NumberOfNew > 0 Then msgText = msgText & vbCrLf & " * New:      " & m_NumberOfNew & " record(s)"
    msgText = msgText & vbCrLf & vbCrLf & _
              "Would you like to SAVE these changes before proceeding?" & vbCrLf & vbCrLf & _
              "[Yes]    = Save and Continue" & vbCrLf & _
              "[No]     = Continue without saving" & vbCrLf & _
              "[Cancel] = Stay in current view"

    Dim userChoice As VbMsgBoxResult
    userChoice = MsgBox(msgText, vbYesNoCancel + vbQuestion, "Save Changes?")

    Select Case userChoice
        Case vbYes
            Call MakeChanges
            Call RefreshData
            TryCommit = True
        Case vbNo
            Call RefreshData
            TryCommit = True
        Case vbCancel
            TryCommit = False
    End Select
    Exit Function

ErrHandler:
    Debug.Print "[ERROR] DataHub_Product.TryCommit: #" & Err.Number & " - " & Err.Description
    TryCommit = False
End Function

Private Sub MakeChanges()
    ' PURPOSE: Executes transactional pipeline to persist parent lookups,
    '          modified records, and new records to master table.
    On Error GoTo ErrHandler

    Dim inTrans As Boolean: inTrans = False
    Dim hasWrittenChanges As Boolean: hasWrittenChanges = False

    DBEngine.BeginTrans: inTrans = True

    ' 1. RESOLVE AND SYNC FOREIGN KEY LOOKUPS (CATEGORIES, LINES, UNITS, PACKAGES)
    If (m_NumberOfModified > 0 Or m_NumberOfNew > 0) Then
        Call SyncParents("Products_Linked")
        hasWrittenChanges = True
    End If

    ' 2. SYNC MODIFIED PRODUCT RECORDS
    If m_NumberOfModified > 0 Then
        Call SyncModified("Products_Linked", "Products")
        hasWrittenChanges = True
    End If

    ' 3. SYNC NEWLY CREATED PRODUCT RECORDS
    If m_NumberOfNew > 0 Then
        Call SyncNew("Products_Linked", "Products")
        hasWrittenChanges = True
    End If

    ' 4. UPDATE BASELINE SNAPSHOT & COMMIT TRANSACTION
    If hasWrittenChanges Then
        Call m_Buffer.TakeSnapshot("Products_Linked", EXCLUDE_FIELDS)
        DBEngine.CommitTrans: inTrans = False
    Else
        DBEngine.Rollback: inTrans = False
    End If

    m_NumberOfModified = 0
    m_NumberOfNew = 0
    Exit Sub

ErrHandler:
    On Error Resume Next
    If inTrans Then DBEngine.Rollback
    On Error GoTo 0
    Debug.Print "[ERROR] DataHub_Product.MakeChanges: #" & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "DataHub_Product.MakeChanges", Err.Description
End Sub

Private Sub SyncParents(ByVal StagingTableName As String)
    ' PURPOSE: Resolves and syncs foreign key lookups (Categories, Lines, Units, Packages).
    On Error GoTo ErrHandler

    Dim sourceDb As DAO.Database, tableField As DAO.Field
    Set sourceDb = CurrentDb

    ' 1. SYNC HIERARCHICAL PARENTS (Category -> Line)
    Call SyncLookupRecord("MarketingCategories", StagingTableName, "MarketingCategory")
    Call SyncMapLookupID(StagingTableName, "MarketingCategoryID", "MarketingCategories", "MarketingCategory")

    Call SyncLookupRecord("MarketingLines", StagingTableName, "MarketingLine", "MarketingCategoryID")
    Call SyncMapLookupID(StagingTableName, "MarketingLineID", "MarketingLines", "MarketingLine", "MarketingCategoryID")

    ' 2. SYNC REMAINING LOOKUPS (Order Independent)
    For Each tableField In sourceDb.TableDefs(StagingTableName).Fields
        Select Case tableField.name
            Case "PackageTypeID"
                Call SyncLookupRecord("PackageTypes", StagingTableName, "PackageType")
                Call SyncMapLookupID(StagingTableName, "PackageTypeID", "PackageTypes", "PackageType")
            Case "CapacityUnitID"
                Call SyncLookupRecord("CapacityUnits", StagingTableName, "CapacityUnit")
                Call SyncMapLookupID(StagingTableName, "CapacityUnitID", "CapacityUnits", "CapacityUnit")
        End Select
    Next tableField
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] DataHub_Product.SyncParents: #" & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "DataHub_Product.SyncParents", Err.Description
End Sub

Private Sub SyncModified(ByVal StagingTableName As String, ByVal MasterTableName As String)
    ' PURPOSE: Updates modified field values from staging into master table.
    On Error GoTo ErrHandler

    Dim sourceDb As DAO.Database, tableField As DAO.Field
    Set sourceDb = CurrentDb

    For Each tableField In sourceDb.TableDefs(MasterTableName).Fields
        If Exist_Field(StagingTableName, tableField.name) Then
            Select Case tableField.name
                Case "ProductID", "InsertDate", "RowNum", "BeforeHash", "AfterHash"
                    ' Skip primary key and system tracking columns
                Case Else
                    Call SyncModifiedRecords(MasterTableName, StagingTableName, tableField.name, "ProductID")
            End Select
        End If
    Next tableField
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] DataHub_Product.SyncModified: #" & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "DataHub_Product.SyncModified", Err.Description
End Sub

Private Sub SyncNew(ByVal StagingTableName As String, ByVal MasterTableName As String)
    ' PURPOSE: Inserts newly created staging records into master table.
    On Error GoTo ErrHandler

    Dim sourceDb As DAO.Database, tableField As DAO.Field, fieldCount As Integer: fieldCount = 0
    Dim sqlColumnNames() As String, sqlInsertValues() As String
    Set sourceDb = CurrentDb

    ReDim sqlColumnNames(sourceDb.TableDefs(MasterTableName).Fields.Count)
    ReDim sqlInsertValues(sourceDb.TableDefs(MasterTableName).Fields.Count)

    For Each tableField In sourceDb.TableDefs(MasterTableName).Fields
        If Exist_Field(StagingTableName, tableField.name) Then
            Select Case tableField.name
                Case "ProductID", "IsActive", "ChangeDate", "RowNum", "BeforeHash", "AfterHash"
                    ' Skip auto-managed columns
                Case Else
                    sqlColumnNames(fieldCount) = tableField.name
                    If tableField.name = "InsertDate" Then
                        sqlInsertValues(fieldCount) = "Now()"
                    Else
                        sqlInsertValues(fieldCount) = tableField.name
                    End If
                    fieldCount = fieldCount + 1
            End Select
        End If
    Next tableField

    If fieldCount = 0 Then Exit Sub

    ReDim Preserve sqlColumnNames(fieldCount - 1)
    ReDim Preserve sqlInsertValues(fieldCount - 1)

    Dim sqlFields As String: sqlFields = Join(sqlColumnNames, ",")
    Dim sqlValues As String: sqlValues = Join(sqlInsertValues, ",")

    Call SyncNewRecords(MasterTableName, StagingTableName, sqlFields, sqlValues)
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] DataHub_Product.SyncNew: #" & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "DataHub_Product.SyncNew", Err.Description
End Sub

'====================================================================
' ROW-LEVEL DOMAIN OPERATIONS
'====================================================================

Public Function Deactivate(ByVal idList As String) As Boolean
    ' PURPOSE: Soft-deletes selected products (tombstones Barcode with _DEL_ and sets IsActive = False).
    On Error GoTo ErrHandler

    Dim cleanList As String: cleanList = Trim(idList)
    If Len(cleanList) = 0 Then Exit Function

    Dim idArray() As String: idArray = Split(cleanList, ",")
    Dim countSelected As Long: countSelected = UBound(idArray) - LBound(idArray) + 1
    If countSelected <= 0 Then Exit Function

    If MsgBox("Confirm: Deactivate " & countSelected & " selected record(s)?", _
              vbYesNo + vbQuestion, "Confirm Deactivation") <> vbYes Then
        Exit Function
    End If

    If m_Buffer.DeactivateRows("Products_Linked", "ProductID", cleanList, "Products", "Barcode") Then
        Call ReleaseLocks
        Call RestoreLocks
        Deactivate = True
    End If
    Exit Function

ErrHandler:
    Debug.Print "[ERROR] DataHub_Product.Deactivate: #" & Err.Number & " - " & Err.Description
    Deactivate = False
End Function

Public Function Restore(ByVal ProductID As Variant) As Boolean
    ' PURPOSE: Restores a soft-deleted product back to active live state.
    On Error GoTo ErrHandler

    If IsNull(ProductID) Or Len(CStr(ProductID)) = 0 Then Exit Function

    If m_Buffer.RestoreRow("Products_Linked", "ProductID", ProductID, "Products", "Barcode") Then
        Call ReleaseLocks
        Call RestoreLocks
        Restore = True
    End If
    Exit Function

ErrHandler:
    Debug.Print "[ERROR] DataHub_Product.Restore: #" & Err.Number & " - " & Err.Description
    Restore = False
End Function

Public Function UpdateTimestamp(ByVal ProductID As Variant) As Boolean
    ' PURPOSE: Updates ChangeDate timestamp for a modified product row.
    On Error GoTo ErrHandler

    If IsNull(ProductID) Or Len(CStr(ProductID)) = 0 Then Exit Function
    UpdateTimestamp = m_Buffer.TouchRow("Products_Linked", "ProductID", ProductID)
    Exit Function

ErrHandler:
    Debug.Print "[ERROR] DataHub_Product.UpdateTimestamp: #" & Err.Number & " - " & Err.Description
    UpdateTimestamp = False
End Function

'====================================================================
' MEDIA ASSET OPERATIONS
'====================================================================

Public Function AddImage(ByVal ProductID As Long, _
                         ByVal sourcePath As String, _
                         ByVal TargetName As String) As String
    ' PURPOSE: Ingests image into MediaStorage and updates product ImagePath in database.
    On Error GoTo ErrHandler

    If m_Media Is Nothing Or ProductID = 0 Or Len(Trim(TargetName)) = 0 Then Exit Function

    Dim newRelativePath As String
    newRelativePath = m_Media.CopyImageToLocal(sourcePath, "Products", Trim(TargetName))
    If Len(newRelativePath) = 0 Then Exit Function

    Call CurrentDb.Execute("UPDATE [Products_Linked] SET [ImagePath] = '" & _
                           Replace(newRelativePath, "'", "''") & "' " & _
                           "WHERE [ProductID] = " & ProductID & ";", dbFailOnError)

    Call UpdateTimestamp(ProductID)
    AddImage = newRelativePath
    Exit Function

ErrHandler:
    Debug.Print "[ERROR] DataHub_Product.AddImage: #" & Err.Number & " - " & Err.Description
    AddImage = ""
End Function

Public Sub RemoveImage(ByVal ProductID As Long)
    ' PURPOSE: Clears product ImagePath reference in database.
    On Error GoTo ErrHandler

    If ProductID = 0 Then Exit Sub

    Call CurrentDb.Execute("UPDATE [Products_Linked] SET [ImagePath] = Null " & _
                           "WHERE [ProductID] = " & ProductID & ";", dbFailOnError)

    Call UpdateTimestamp(ProductID)
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] DataHub_Product.RemoveImage: #" & Err.Number & " - " & Err.Description
End Sub

'====================================================================
' INTERNAL LOCK MANAGEMENT (Private)
'====================================================================

Private Sub ReleaseLocks()
    ' PURPOSE: Temporarily disconnects RecordSource on registered views to allow schema/DDL edits.
    On Error GoTo ErrHandler

    If m_Views Is Nothing Or m_RecordSourceHistory Is Nothing Then Exit Sub
    m_RecordSourceHistory.RemoveAll

    Dim key As Variant
    Dim frm As Object

    For Each key In m_Views.Keys
        Set frm = m_Views(key)
        If Not frm Is Nothing Then
            On Error Resume Next
            m_RecordSourceHistory(key) = frm.RecordSource
            frm.RecordSource = ""
            On Error GoTo ErrHandler
        End If
    Next key
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] DataHub_Product.ReleaseLocks: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub RestoreLocks()
    ' PURPOSE: Rebinds RecordSource on registered views to reload latest data.
    On Error GoTo ErrHandler

    If m_Views Is Nothing Or m_RecordSourceHistory Is Nothing Then Exit Sub

    Dim key As Variant
    Dim frm As Object

    For Each key In m_Views.Keys
        If m_RecordSourceHistory.exists(key) Then
            Set frm = m_Views(key)
            If Not frm Is Nothing Then
                On Error Resume Next
                frm.RecordSource = m_RecordSourceHistory(key)
                On Error GoTo ErrHandler
            End If
        End If
    Next key
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] DataHub_Product.RestoreLocks: #" & Err.Number & " - " & Err.Description
End Sub
