VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StagingBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'@Folder "Extensions.Data"
'====================================================================
' CLASS   : StagingBuffer
' PURPOSE : Primitive toolkit for staging table operations.
'           Self-manages TempDatabase.accdb file lifecycle.
'           Provides raw DDL, TableDef linking, and SHA-256
'           hash-based change tracking.
'           DataEngine orchestrates the calling sequence.
'====================================================================

' Internal state (self-managed file lifecycle)
Private m_TempDB        As DAO.Database
Private m_TempDBPath    As String
Private m_ExcludeFields As String

'====================================================================
' FILE LIFECYCLE (self-managed)
'====================================================================

Private Sub Class_Initialize()
    ' PURPOSE: Recreates fresh hidden TempDatabase.accdb for the app session.
    '          Sets m_TempDBPath, m_TempDB.
    m_TempDBPath = CurrentProject.Path & "\TempDatabase.accdb"

    If FileExists(m_TempDBPath) Then
        On Error Resume Next
        Call SetAttr(m_TempDBPath, vbNormal)
        Call Kill(m_TempDBPath)
        On Error GoTo 0
    End If

        Set m_TempDB = DBEngine.CreateDatabase(m_TempDBPath, dbLangGeneral)
        Call SetAttr(m_TempDBPath, vbHidden)

    ' Clean up any stale linked tables from prior sessions or crashes
    Call PurgeOrphans
End Sub

Private Sub Class_Terminate()
    ' PURPOSE: Closes TempDB handle. Deletes .accdb file from disk.
    On Error Resume Next
    If Not m_TempDB Is Nothing Then
        m_TempDB.Close
        Set m_TempDB = Nothing
    End If

    If FileExists(m_TempDBPath) Then
        Call SetAttr(m_TempDBPath, vbNormal)
        Call Kill(m_TempDBPath)
    End If
    On Error GoTo 0
End Sub

'====================================================================
' PROPERTY
'====================================================================

Public Property Get Database() As DAO.Database
    ' PURPOSE: Exposes TempDB for caller to run custom DDL
    '          (e.g. CREATE UNIQUE INDEX, SET DefaultValue)
    Set Database = m_TempDB
End Property

'====================================================================
' HIGH-LEVEL ORCHESTRATION
'====================================================================

Public Function InitializeStaging(ByVal SourceQueryName As String, _
                                  ByVal TargetTableName As String, _
                                  Optional ByVal linkName As String = "", _
                                  Optional ByVal ExcludeFields As String = "") As Boolean
    ' PURPOSE: High-level orchestrator for full staging table creation:
    '          Drop -> Import -> Add Tracking Columns -> Set Defaults -> Link -> Baseline Snapshot.
    On Error GoTo ErrHandler

    Dim hostDb As DAO.Database: Set hostDb = CurrentDb
    Dim strSql As String

    ' 1. Drop existing table in TempDB (Safe No-Op)
    On Error Resume Next
    If TableExists(m_TempDB.TableDefs, TargetTableName) Then
        Call m_TempDB.Execute("DROP TABLE [" & TargetTableName & "];", dbFailOnError)
        Call m_TempDB.TableDefs.Refresh
    End If
    On Error GoTo ErrHandler

    ' 2. Import data from host query to TempDB
    strSql = "SELECT * INTO [" & TargetTableName & "] IN '" & m_TempDBPath & "' FROM [" & SourceQueryName & "];"
    Call hostDb.Execute(strSql, dbFailOnError)
    Call m_TempDB.TableDefs.Refresh

    ' 3. Add BeforeHash and AfterHash tracking columns
    Call m_TempDB.Execute("ALTER TABLE [" & TargetTableName & "] ADD COLUMN [BeforeHash] CHAR(64);", dbFailOnError)
    Call m_TempDB.Execute("ALTER TABLE [" & TargetTableName & "] ADD COLUMN [AfterHash] CHAR(64);", dbFailOnError)
    Call m_TempDB.TableDefs.Refresh

    ' 4. Ensure IsActive defaults to True for new rows (safe if field absent)
    On Error Resume Next
    m_TempDB.TableDefs(TargetTableName).Fields("IsActive").DefaultValue = CStr(True)
    On Error GoTo ErrHandler

    ' 5 & 6. Create linked TableDef in CurrentDb and compute baseline hash
    m_ExcludeFields = ExcludeFields
    Dim actualLinkName As String
    actualLinkName = LinkTable(TargetTableName, linkName)
    InitializeStaging = TakeSnapshot(actualLinkName, ExcludeFields)
    Exit Function

ErrHandler:
    Debug.Print "[ERROR] StagingBuffer.InitializeStaging: #" & Err.Number & " - " & Err.Description
    InitializeStaging = False
End Function

'====================================================================
' TABLEDEF LINKING (TempDB <-> CurrentDb bridge)
'====================================================================

Public Function LinkTable(ByVal TableName As String, _
                          Optional ByVal linkName As String = "") As String
    ' PURPOSE: Creates a linked TableDef in CurrentDb pointing to
    '          [TableName] in TempDB.
    '          LinkName defaults to TableName & "_Linked" if omitted.
    ' RETURNS: Name of the linked table created in CurrentDb.
    On Error GoTo ErrHandler

    Dim targetLinkName As String
    If Len(linkName) = 0 Then
        targetLinkName = TableName & "_Linked"
    Else
        targetLinkName = linkName
    End If

    Dim hostDb As DAO.Database: Set hostDb = CurrentDb
    Dim connectStr As String: connectStr = ";DATABASE=" & m_TempDBPath
    Dim tblDef As DAO.TableDef

    If TableExists(hostDb.TableDefs, targetLinkName) Then
        Set tblDef = hostDb.TableDefs(targetLinkName)
        If tblDef.Connect <> connectStr Or tblDef.SourceTableName <> TableName Then
            hostDb.TableDefs.Delete targetLinkName
            Set tblDef = hostDb.CreateTableDef(targetLinkName)
            tblDef.Connect = connectStr
            tblDef.SourceTableName = TableName
            hostDb.TableDefs.Append tblDef
        Else
            tblDef.RefreshLink
        End If
    Else
        Set tblDef = hostDb.CreateTableDef(targetLinkName)
        tblDef.Connect = connectStr
        tblDef.SourceTableName = TableName
        hostDb.TableDefs.Append tblDef
    End If

    RefreshDatabaseWindow
    LinkTable = targetLinkName
    Exit Function

ErrHandler:
    Debug.Print "[ERROR] StagingBuffer.LinkTable: #" & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "StagingBuffer.LinkTable", Err.Description
End Function

Public Sub UnlinkTable(ByVal linkName As String)
    ' PURPOSE: Removes linked TableDef from CurrentDb.
    '          Safe no-op if link does not exist.
    On Error Resume Next
    Dim hostDb As DAO.Database: Set hostDb = CurrentDb
    Dim targetLinkName As String: targetLinkName = linkName

    If Not TableExists(hostDb.TableDefs, targetLinkName) Then
        If TableExists(hostDb.TableDefs, targetLinkName & "_Linked") Then
            targetLinkName = targetLinkName & "_Linked"
        End If
    End If

    If TableExists(hostDb.TableDefs, targetLinkName) Then
        hostDb.TableDefs.Delete targetLinkName
        RefreshDatabaseWindow
    End If

    PurgeOrphans
    On Error GoTo 0
End Sub

Public Sub PurgeOrphans()
    ' PURPOSE: 2-way garbage collector. Removes:
    '          1. Orphaned linked TableDefs in CurrentDb (pointing to non-existent TempDB tables).
    '          2. Orphaned tables in TempDB (having no corresponding linked TableDef in CurrentDb).
    On Error Resume Next
    Dim hostDb As DAO.Database: Set hostDb = CurrentDb
    Dim activeLinks As New Scripting.Dictionary
    activeLinks.CompareMode = TextCompare
    Dim i As Long
    Dim tblDef As DAO.TableDef

    ' 1. CLEAN CURRENTDB: Remove broken links and collect active source table names
    For i = hostDb.TableDefs.Count - 1 To 0 Step -1
        Set tblDef = hostDb.TableDefs(i)
        If InStr(tblDef.Connect, m_TempDBPath) > 0 Then
            If Not TableExists(m_TempDB.TableDefs, tblDef.SourceTableName) Then
                hostDb.TableDefs.Delete tblDef.name
            Else
                activeLinks(tblDef.SourceTableName) = True
            End If
        End If
    Next i

    ' 2. CLEAN TEMPDB: Remove tables that have no link pointing to them in CurrentDb
    For i = m_TempDB.TableDefs.Count - 1 To 0 Step -1
        Set tblDef = m_TempDB.TableDefs(i)
        If Not (tblDef.name Like "MSys*" Or tblDef.name Like "~*") Then
            If Not activeLinks.exists(tblDef.name) Then
                Call m_TempDB.Execute("DROP TABLE [" & tblDef.name & "];", dbFailOnError)
            End If
        End If
    Next i

    RefreshDatabaseWindow
    On Error GoTo 0
End Sub

'====================================================================
' HASH & CHANGE TRACKING
'====================================================================

Public Function TakeSnapshot(ByVal TableName As String, _
                             Optional ByVal ExcludeFields As String = "") As Boolean
    ' PURPOSE: Computes SHA-256 hash of each row and writes to BeforeHash.
    '          Called after data load (baseline) and after commit (reset).
    ' PARAM  : ExcludeFields - comma-delimited field names to skip when hashing.
    '          DataEngine decides which fields are non-hashable (domain policy).
    On Error GoTo ErrHandler

    Dim activeExclude As String: activeExclude = ExcludeFields
    If Len(Trim(activeExclude)) = 0 Then activeExclude = m_ExcludeFields

    Dim strSql As String
    strSql = SQL_UpdateHash(TableName, "BeforeHash", activeExclude)
    If Len(strSql) = 0 Then Exit Function

    Call CurrentDb.Execute(strSql, dbFailOnError)
    TakeSnapshot = True
    Exit Function

ErrHandler:
    Debug.Print "[ERROR] StagingBuffer.TakeSnapshot: #" & Err.Number & " - " & Err.Description
    TakeSnapshot = False
End Function

Public Function UpdateAfterHash(ByVal TableName As String, _
                                Optional ByVal ExcludeFields As String = "") As Boolean
    ' PURPOSE: Computes SHA-256 hash of each row and writes to AfterHash.
    ' PARAM  : ExcludeFields - same as TakeSnapshot.
    On Error GoTo ErrHandler

    Dim activeExclude As String: activeExclude = ExcludeFields
    If Len(Trim(activeExclude)) = 0 Then activeExclude = m_ExcludeFields

    Dim strSql As String
    strSql = SQL_UpdateHash(TableName, "AfterHash", activeExclude)
    If Len(strSql) = 0 Then Exit Function

    Call CurrentDb.Execute(strSql, dbFailOnError)
    UpdateAfterHash = True
    Exit Function

ErrHandler:
    Debug.Print "[ERROR] StagingBuffer.UpdateAfterHash: #" & Err.Number & " - " & Err.Description
    UpdateAfterHash = False
End Function

' NOTE (Architectural Decision - Pragmatic Trade-off):
' In pure layered architecture, domain filtering semantics (like [IsActive] soft-deletes)
' strictly belong to the DataHub/Domain layer. Placing IsDataChanged and IsDataNew here
' in StagingBuffer is a pragmatic decision to avoid boilerplate SQL count queries across multiple Hubs.

Public Function IsDataChanged(ByVal TableName As String) As Long
    ' PURPOSE: Returns count of rows where AfterHash <> BeforeHash.
    On Error GoTo ErrHandler

    Dim sqlWhere As String
    sqlWhere = " WHERE [AfterHash] <> [BeforeHash]"
    If FieldExists(TableName, "IsActive") Then
        sqlWhere = sqlWhere & " AND [IsActive] = True"
    End If

    Dim strSql As String
    strSql = "SELECT Count(*) FROM [" & TableName & "]" & sqlWhere & ";"

    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset(strSql, dbOpenForwardOnly)
    IsDataChanged = Nz(rs.Fields(0).value, 0)
    rs.Close
    Set rs = Nothing
    Exit Function

ErrHandler:
    Debug.Print "[ERROR] StagingBuffer.IsDataChanged: #" & Err.Number & " - " & Err.Description
    IsDataChanged = 0
End Function

Public Function IsDataNew(ByVal TableName As String) As Long
    ' PURPOSE: Returns count of rows where BeforeHash IS NULL AND AfterHash IS NOT NULL.
    On Error GoTo ErrHandler

    Dim sqlWhere As String
    sqlWhere = " WHERE [BeforeHash] IS NULL AND [AfterHash] IS NOT NULL"
    If FieldExists(TableName, "IsActive") Then
        sqlWhere = sqlWhere & " AND [IsActive] = True"
    End If

    Dim strSql As String
    strSql = "SELECT Count(*) FROM [" & TableName & "]" & sqlWhere & ";"

    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset(strSql, dbOpenForwardOnly)
    IsDataNew = Nz(rs.Fields(0).value, 0)
    rs.Close
    Set rs = Nothing
    Exit Function

ErrHandler:
    Debug.Print "[ERROR] StagingBuffer.IsDataNew: #" & Err.Number & " - " & Err.Description
    IsDataNew = 0
End Function

'====================================================================
' ROW OPERATIONS & SOFT DELETION
'====================================================================

Public Function DeactivateRows(ByVal TableName As String, _
                              ByVal idFieldName As String, _
                              ByVal idList As String, _
                              ByVal MasterTableName As String, _
                              Optional ByVal TombstoneField As String = "") As Boolean
    ' PURPOSE: Soft-deletes rows in Staging and Master tables, optionally appending a tombstone suffix to a unique field.
    On Error GoTo ErrHandler

    Dim cleanList As String: cleanList = Trim(idList)
    If Len(cleanList) = 0 Then Exit Function

    Dim sqlTombstone As String: sqlTombstone = ""
    If Len(Trim(TombstoneField)) > 0 Then
        sqlTombstone = ", [" & TombstoneField & "] = [" & TombstoneField & "] & '_DEL_' & [" & idFieldName & "]"
    End If

    Dim inTrans As Boolean: inTrans = False
    DBEngine.BeginTrans: inTrans = True

    ' 1. UPDATE STAGING TABLE
    Call CurrentDb.Execute("UPDATE [" & TableName & "] SET [IsActive] = False" & sqlTombstone & " " & _
                           "WHERE [" & idFieldName & "] IN (" & cleanList & ");", dbFailOnError)

    ' 2. UPDATE PRODUCTION MASTER TABLE
    If Len(Trim(MasterTableName)) > 0 Then
        Call CurrentDb.Execute("UPDATE [" & MasterTableName & "] SET [IsActive] = False" & sqlTombstone & " " & _
                               "WHERE [" & idFieldName & "] IN (" & cleanList & ");", dbFailOnError)
    End If

    DBEngine.CommitTrans: inTrans = False
    DeactivateRows = True
    Exit Function

ErrHandler:
    On Error Resume Next
    If inTrans Then DBEngine.Rollback
    On Error GoTo 0
    Debug.Print "[ERROR] StagingBuffer.DeactivateRows: #" & Err.Number & " - " & Err.Description
    DeactivateRows = False
End Function

Public Function RestoreRow(ByVal TableName As String, _
                           ByVal idFieldName As String, _
                           ByVal TargetID As Variant, _
                           ByVal MasterTableName As String, _
                           Optional ByVal TombstoneField As String = "") As Boolean
    ' PURPOSE: Restores a soft-deleted row in Master, removes tombstone suffix, and deletes from Staging.
    On Error GoTo ErrHandler

    If IsNull(TargetID) Or Len(CStr(TargetID)) = 0 Then Exit Function

    Dim targetVal As Long: targetVal = CLng(TargetID)
    Dim sqlTombstone As String: sqlTombstone = ""

    If Len(Trim(TombstoneField)) > 0 Then
        sqlTombstone = ", [" & TombstoneField & "] = IIf(InStr([" & TombstoneField & "], '_DEL_') > 0, " & _
                       "Left([" & TombstoneField & "], InStr([" & TombstoneField & "], '_DEL_') - 1), [" & TombstoneField & "])"
    End If

    Dim inTrans As Boolean: inTrans = False
    DBEngine.BeginTrans: inTrans = True

    ' 1. REACTIVATE IN MASTER TABLE AND REMOVE TOMBSTONE
    Call CurrentDb.Execute("UPDATE [" & MasterTableName & "] SET [IsActive] = True" & sqlTombstone & " " & _
                           "WHERE [" & idFieldName & "] = " & targetVal & ";", dbFailOnError)

    ' 2. DELETE RESTORED ROW FROM STAGING TABLE
    ' NOTE: Updating IsActive = True in Staging causes runtime bugs; row must be
    '       purged directly from Staging and reloaded via RefreshData when returning to Live.
    Call CurrentDb.Execute("DELETE FROM [" & TableName & "] " & _
                           "WHERE [" & idFieldName & "] = " & targetVal & ";", dbFailOnError)

    DBEngine.CommitTrans: inTrans = False
    RestoreRow = True
    Exit Function

ErrHandler:
    On Error Resume Next
    If inTrans Then DBEngine.Rollback
    On Error GoTo 0
    If Err.Number = 3022 Then
        MsgBox "RESTORE FAILED: A unique value (" & TombstoneField & ") is already in use by an active record.", _
               vbExclamation, "Conflict Detected"
    Else
        Debug.Print "[ERROR] StagingBuffer.RestoreRow: #" & Err.Number & " - " & Err.Description
    End If
    RestoreRow = False
End Function

Public Function TouchRow(ByVal TableName As String, _
                         ByVal idFieldName As String, _
                         ByVal TargetID As Variant) As Boolean
    ' PURPOSE: Updates ChangeDate modification timestamp (Now()) for a specific record.
    On Error GoTo ErrHandler

    If IsNull(TargetID) Or Len(CStr(TargetID)) = 0 Then Exit Function

    Dim targetVal As Long: targetVal = CLng(TargetID)
    Dim strSql As String
    strSql = "UPDATE [" & TableName & "] SET [ChangeDate] = Now() " & _
             "WHERE [" & idFieldName & "] = " & targetVal & ";"

    Call CurrentDb.Execute(strSql, dbFailOnError)
    TouchRow = True
    Exit Function

ErrHandler:
    Debug.Print "[ERROR] StagingBuffer.TouchRow: #" & Err.Number & " - " & Err.Description
    TouchRow = False
End Function

'====================================================================
' PRIVATE HELPERS
'====================================================================

Private Function BuildHashStringFromFields(ByVal FieldsCollection As Object) As String
    ' PURPOSE: Concatenates field expressions into a pipe-delimited string
    '          for SHA-256 input: CStr(Nz([Field1],"")) & "|" & ...
    Dim fieldCount As Integer: fieldCount = FieldsCollection.Count
    If fieldCount = 0 Then Exit Function

    Dim fieldsArray() As String
    ReDim fieldsArray(fieldCount - 1)

    Dim tableField As Object
    Dim i As Integer: i = 0
    For Each tableField In FieldsCollection
        fieldsArray(i) = "CStr(Nz([" & tableField.name & "],""""))"
        i = i + 1
    Next tableField

    BuildHashStringFromFields = Join(fieldsArray, " & ""|"" & ")
End Function

Private Function SQL_UpdateHash(ByVal TableName As String, _
                                ByVal HashColumnName As String, _
                                ByVal ExcludeFields As String) As String
    ' PURPOSE: Generates UPDATE SQL to compute SHA-256 for hashable fields.
    '          Always excludes BeforeHash/AfterHash internally.
    '          Additional exclusions come from ExcludeFields param (domain policy).
    On Error GoTo ErrHandler

    Dim excludeList As String
    excludeList = "," & LCase(Replace(ExcludeFields, " ", "")) & ",beforehash,afterhash,"

    Dim fieldsColl As New VBA.Collection
    Dim hostDb As DAO.Database: Set hostDb = CurrentDb
    Dim fieldDef As DAO.Field

    For Each fieldDef In hostDb.TableDefs(TableName).Fields
        If InStr(1, excludeList, "," & LCase(fieldDef.name) & ",") = 0 Then
            Call fieldsColl.Add(fieldDef)
        End If
    Next fieldDef

    Dim hashSource As String: hashSource = BuildHashStringFromFields(fieldsColl)
    If Len(hashSource) = 0 Then Exit Function

    Dim sqlWhere As String: sqlWhere = ""
    If FieldExists(TableName, "IsActive") Then sqlWhere = " WHERE IsActive = True"

    SQL_UpdateHash = "UPDATE [" & TableName & "] " & _
                     "SET [" & HashColumnName & "] = SHA256_CNG(" & hashSource & ")" & sqlWhere & ";"
    Exit Function

ErrHandler:
    Debug.Print "[ERROR] StagingBuffer.SQL_UpdateHash: #" & Err.Number & " - " & Err.Description
    SQL_UpdateHash = ""
End Function

Private Function FieldExists(ByVal TableName As String, ByVal fieldName As String) As Boolean
    ' PURPOSE: Checks if a field exists in a table in CurrentDb.
    On Error Resume Next
    Dim testField As DAO.Field
    Set testField = CurrentDb.TableDefs(TableName).Fields(fieldName)
    FieldExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function TableExists(ByRef TableDefsCollection As DAO.TableDefs, _
                             ByVal TableName As String) As Boolean
    ' PURPOSE: Checks if a table exists in the given TableDefs collection.
    On Error Resume Next
    Dim checkTable As DAO.TableDef
    Call TableDefsCollection.Refresh
    Err.Clear
    Set checkTable = TableDefsCollection(TableName)
    TableExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function FileExists(ByVal FilePath As String) As Boolean
    ' PURPOSE: Checks if a file exists on disk.
    On Error Resume Next
    FileExists = (Dir(FilePath, vbHidden Or vbNormal Or vbReadOnly) <> "")
    On Error GoTo 0
End Function
