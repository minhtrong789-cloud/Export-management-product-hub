VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MediaStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'@Folder "Extensions.Data"

'====================================================================
' CLASS   : MediaStorage
' PURPOSE : Service for physical media asset storage, file archiving,
'           soft-deletion (TrashBin), and 2-way orphan integrity.
'====================================================================

' Internal state
Private m_FSO          As Scripting.FileSystemObject
Private m_StorageRoot  As String
Private m_TrashBinRoot As String

'====================================================================
' LIFECYCLE
'====================================================================

Private Sub Class_Initialize()
    ' PURPOSE: Initializes paths and ensures MediaStorage & TrashBin exist.
    Set m_FSO = New Scripting.FileSystemObject
    m_StorageRoot = CurrentProject.Path & "\MediaStorage"
    m_TrashBinRoot = m_StorageRoot & "\TrashBin"

    If Not m_FSO.FolderExists(m_StorageRoot) Then Call m_FSO.CreateFolder(m_StorageRoot)
    If Not m_FSO.FolderExists(m_TrashBinRoot) Then Call m_FSO.CreateFolder(m_TrashBinRoot)
End Sub

Private Sub Class_Terminate()
    ' PURPOSE: Cleans up internal references.
    Set m_FSO = Nothing
End Sub


'====================================================================
' FILE INGESTION & ARCHIVING
'====================================================================

Public Function CopyImageToLocal(ByVal SourceImagePath As String, _
                                 ByVal TableName As String, _
                                 ByVal TargetName As String) As String
    ' PURPOSE: Copies an incoming image file to MediaStorage\<TableName>\<TargetName>.<ext>.
    '          Archives existing file with (1), (2) suffix if duplicate name.
    ' RETURNS: Relative path string e.g. "\MediaStorage\Products\P001.jpg"
    On Error GoTo ErrHandler

    If Not m_FSO.FileExists(SourceImagePath) Then Exit Function

    Dim targetFolder As String: targetFolder = m_StorageRoot & "\" & TableName
    If Not m_FSO.FolderExists(targetFolder) Then Call m_FSO.CreateFolder(targetFolder)

    Dim fileExt As String: fileExt = m_FSO.GetExtensionName(SourceImagePath)
    Dim canonicalName As String: canonicalName = TargetName & "." & fileExt
    Dim fullDestPath As String: fullDestPath = targetFolder & "\" & canonicalName

    ' Archive existing file to prevent destructive overwrite
    If m_FSO.FileExists(fullDestPath) Then
        Dim archiveIndex As Long: archiveIndex = 1
        Dim archivePath As String
        Do
            archivePath = targetFolder & "\" & TargetName & "(" & archiveIndex & ")." & fileExt
            archiveIndex = archiveIndex + 1
        Loop While m_FSO.FileExists(archivePath)

        Dim oldFile As Scripting.File
        Set oldFile = m_FSO.GetFile(fullDestPath)
        oldFile.name = m_FSO.GetFileName(archivePath)
    End If

    ' Execute file copy
    Call m_FSO.CopyFile(SourceImagePath, fullDestPath, True)
    CopyImageToLocal = "\MediaStorage\" & TableName & "\" & canonicalName
    Exit Function

ErrHandler:
    Debug.Print "[ERROR] MediaStorage.CopyImageToLocal: #" & Err.Number & " - " & Err.Description
    CopyImageToLocal = ""
End Function

'====================================================================
' SOFT DELETION & QUOTA MANAGEMENT
'====================================================================

Public Sub MoveToTrashBin(ByVal SourceFilePath As String)
    ' PURPOSE: Moves a file to MediaStorage\TrashBin with conflict resolution,
    '          then maintains the maximum file quota.
    On Error GoTo ErrHandler

    If Not m_FSO.FileExists(SourceFilePath) Then Exit Sub

    Dim fileName As String: fileName = m_FSO.GetFileName(SourceFilePath)
    Dim baseName As String: baseName = m_FSO.GetBaseName(SourceFilePath)
    Dim fileExt  As String: fileExt = m_FSO.GetExtensionName(SourceFilePath)
    Dim destPath As String: destPath = m_TrashBinRoot & "\" & fileName
    Dim suffixIndex As Long: suffixIndex = 1

    Do While m_FSO.FileExists(destPath)
        destPath = m_TrashBinRoot & "\" & baseName & "(" & suffixIndex & ")." & fileExt
        suffixIndex = suffixIndex + 1
    Loop

    Call m_FSO.MoveFile(SourceFilePath, destPath)
    Call RemoveExcessFiles
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] MediaStorage.MoveToTrashBin: #" & Err.Number & " - " & Err.Description
End Sub

Public Sub RemoveExcessFiles(Optional ByVal MaxFiles As Long = 200)
    ' PURPOSE: Deletes oldest files in TrashBin to maintain MaxFiles limit.
    On Error GoTo ErrHandler

    If Not m_FSO.FolderExists(m_TrashBinRoot) Then Exit Sub

    Dim trashFolder As Scripting.Folder
    Set trashFolder = m_FSO.GetFolder(m_TrashBinRoot)

    Dim totalFiles As Long: totalFiles = trashFolder.Files.Count
    If totalFiles <= MaxFiles Then Exit Sub

    Dim excessCount As Long: excessCount = totalFiles - MaxFiles
    Dim fileList() As Variant
    ReDim fileList(1 To totalFiles, 1 To 2)

    Dim f As Scripting.File, i As Long: i = 1
    For Each f In trashFolder.Files
        fileList(i, 1) = f.DateLastModified
        fileList(i, 2) = f.Path
        i = i + 1
    Next f

    ' Sort oldest first (ascending by creation date)
    Dim j As Long, tempDate As Date, tempPath As String
    For i = 1 To totalFiles - 1
        For j = i + 1 To totalFiles
            If fileList(i, 1) > fileList(j, 1) Then
                tempDate = fileList(i, 1): tempPath = fileList(i, 2)
                fileList(i, 1) = fileList(j, 1): fileList(i, 2) = fileList(j, 2)
                fileList(j, 1) = tempDate: fileList(j, 2) = tempPath
            End If
        Next j
    Next i

    ' Delete oldest excess files
    For i = 1 To excessCount
        On Error Resume Next
        Call m_FSO.DeleteFile(fileList(i, 2), True)
        On Error GoTo ErrHandler
    Next i
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] MediaStorage.RemoveExcessFiles: #" & Err.Number & " - " & Err.Description
End Sub

'====================================================================
' 2-WAY ORPHAN INTEGRITY (Disk <-> Database)
'====================================================================

Public Sub ClearOrphanImageFiles(ByVal TableName As String, _
                                 Optional ByVal ImagePathFieldName As String = "ImagePath")
    ' PURPOSE: Moves unreferenced physical files in MediaStorage\<TableName>\ to TrashBin.
    ' OPTIMIZATION: Loads all DB image paths into a Dictionary in 1 query (O(1) in-memory lookup).
    On Error GoTo ErrHandler

    Dim tableFolder As String: tableFolder = m_StorageRoot & "\" & TableName
    If Not m_FSO.FolderExists(tableFolder) Then Exit Sub

    ' 1. Load active database paths into Dictionary in 1 query
    Dim activePaths As New Scripting.Dictionary
    activePaths.CompareMode = TextCompare

    Dim rs As DAO.Recordset
    Dim strSql As String
    strSql = "SELECT [" & ImagePathFieldName & "] FROM [" & TableName & "] " & _
             "WHERE [" & ImagePathFieldName & "] IS NOT NULL AND [" & ImagePathFieldName & "] <> '';"
    Set rs = CurrentDb.OpenRecordset(strSql, dbOpenForwardOnly)

    Do While Not rs.EOF
        Dim rawPath As String: rawPath = Trim(Nz(rs.Fields(0).value, ""))
        If Len(rawPath) > 0 Then
            If Left(rawPath, 1) <> "\" Then rawPath = "\" & rawPath
            activePaths(LCase(rawPath)) = True
        End If
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing

    ' 2. Scan physical folder and move unreferenced files to TrashBin
    Dim folderObj As Scripting.Folder: Set folderObj = m_FSO.GetFolder(tableFolder)
    Dim f As Scripting.File

    For Each f In folderObj.Files
        ' Skip hidden and system files
        If (f.Attributes And 2) = 0 Then
            Dim relPath As String
            relPath = Replace(f.Path, CurrentProject.Path, "", , , vbTextCompare)
            If Left(relPath, 1) <> "\" Then relPath = "\" & relPath

            If Not activePaths.exists(LCase(relPath)) Then
                Call MoveToTrashBin(f.Path)
            End If
        End If
    Next f
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] MediaStorage.ClearOrphanImageFiles: #" & Err.Number & " - " & Err.Description
End Sub

Public Sub ClearOrphanImagePaths(ByVal TableName As String, _
                                 Optional ByVal ImagePathFieldName As String = "ImagePath")
    ' PURPOSE: Nullifies ImagePath in database if physical file is missing on disk.
    On Error GoTo ErrHandler

    Dim rs As DAO.Recordset
    Dim strSql As String
    strSql = "SELECT [" & ImagePathFieldName & "] FROM [" & TableName & "] " & _
             "WHERE [" & ImagePathFieldName & "] IS NOT NULL AND [" & ImagePathFieldName & "] <> '';"
    Set rs = CurrentDb.OpenRecordset(strSql, dbOpenDynaset)

    With rs
        Do While Not .EOF
            Dim dbPath As String: dbPath = Trim(Nz(.Fields(0).value, ""))
            If Len(dbPath) > 0 Then
                Dim fullPath As String
                If Left(dbPath, 1) = "\" Then
                    fullPath = CurrentProject.Path & dbPath
                Else
                    fullPath = CurrentProject.Path & "\" & dbPath
                End If

                If Not m_FSO.FileExists(fullPath) Then
                    .Edit
                    .Fields(0).value = Null
                    .Update
                End If
            End If
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] MediaStorage.ClearOrphanImagePaths: #" & Err.Number & " - " & Err.Description
End Sub
