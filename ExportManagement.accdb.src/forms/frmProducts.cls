Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'@Folder "App.Products"

'====================================================================
' FORM      : frmProducts
' ROLE      : Master Shell Router for Products Domain
'====================================================================

'--- [INTERNAL STATE VARIABLES] ---
Private m_CurrentView    As String   ' "GRID" or "DETAIL"
Private m_IsTrashbin     As Boolean  ' False: Active records | True: Archive records
Private m_IsMaximized    As Boolean  ' False: Standard view | True: Full-width workspace
Private m_IsMediaVisible As Boolean  ' False: Grid full | True: Split-view with image
Private m_CurrentLocale  As String   ' "EN" or "VN"

'--- [PUBLIC STATE ACCESSORS] ---
Public Property Get isMaximized() As Boolean
    isMaximized = m_IsMaximized
End Property

Public Property Get IsMediaVisible() As Boolean
    IsMediaVisible = m_IsMediaVisible
End Property

Public Property Get IsTrashbin() As Boolean
    IsTrashbin = m_IsTrashbin
End Property

Public Property Get CurrentLocale() As String
    CurrentLocale = m_CurrentLocale
End Property

'====================================================================
' FORM LIFECYCLE EVENTS
'====================================================================

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrHandler

    ' Initialize default state values
    m_CurrentView = "GRID"
    m_IsTrashbin = False
    m_IsMaximized = False
    m_IsMediaVisible = False
    m_CurrentLocale = "EN"

    Exit Sub

ErrHandler:
    MsgBox "Error initializing Products Hub: " & Err.Description, vbCritical
    Cancel = True
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler

    ' 1. Reset and initialize header controls
    Me.txtSearch = ""
    Me.cmdSave.Caption = "[ SAVE ]"
    Me.togTrashbin.Caption = "TRSH"
    Me.cmdLanguage.Caption = m_CurrentLocale
    Me.cmdMaximize.Caption = "MAXM"
    Me.cmdToggleView.Caption = "VIEW"
    Me.cmdViewImage.Caption = "IMGE"
    Me.cmdViewImage.Enabled = True
    Me.togTrashbin.Enabled = True
    Me.cmdLanguage.Enabled = True

    ' 2. Load default Grid View into workspace viewport
    If Me.sfWorkspace.SourceObject <> "frmProductGridView" Then
        Me.sfWorkspace.SourceObject = "frmProductGridView"
    End If

    Exit Sub

ErrHandler:
    MsgBox "Error loading Products Hub: " & Err.Description, vbCritical
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrHandler

    ' 1. Ensure child workspace commits any active editing before closing
    If Not Me.sfWorkspace.Form Is Nothing Then
        If Me.sfWorkspace.Form.Dirty Then Me.sfWorkspace.Form.Dirty = False
    End If

    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProducts.Form_Unload: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub Form_Close()
    On Error Resume Next
    ' Cleanup references
End Sub

'====================================================================
' HEADER BUTTON EVENTS (VIEW & NAVIGATION)
'====================================================================

Private Sub cmdSave_Click()
    On Error GoTo ErrHandler

    ' 1. Commit active row edits in any active subform level
    If Me.Dirty Then Me.Dirty = False
    If Not Me.sfWorkspace.Form Is Nothing Then
        If Me.sfWorkspace.Form.Dirty Then Me.sfWorkspace.Form.Dirty = False
        If m_CurrentView = "GRID" Then
            If Not Me.sfWorkspace.Form.Controls("sfProductGrid").Form Is Nothing Then
                If Me.sfWorkspace.Form.Controls("sfProductGrid").Form.Dirty Then
                    Me.sfWorkspace.Form.Controls("sfProductGrid").Form.Dirty = False
                End If
            End If
        End If
    End If

    ' 2. Persist staging changes to backend via DataHub
    If Not AppDataHub Is Nothing Then
        If Not AppDataHub.TryCommit(Silent:=True) Then Exit Sub
    End If

    ' 3. Visual feedback (DONE pulse effect)
    Dim bckColor As Long, hovColor As Long, successColor As Long
    bckColor = Me.cmdSave.BackColor
    hovColor = Me.cmdSave.HoverColor
    successColor = HexToRBGLong("#2ECC71")

    Me.cmdSave.Caption = "DONE"
    Me.cmdSave.BackColor = successColor
    Me.cmdSave.HoverColor = successColor
    Me.Repaint

    Dim endTime As Double
    endTime = Timer + 0.4
    Do While Timer < endTime
        DoEvents
    Loop

    Me.cmdSave.Caption = "[ SAVE ]"
    Me.cmdSave.BackColor = bckColor
    Me.cmdSave.HoverColor = hovColor
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProducts.cmdSave_Click: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub togTrashbin_Click()
    On Error GoTo ErrHandler

    ' 1. Check and save pending edits before switching mode
    If Not AppDataHub Is Nothing Then
        If Not AppDataHub.TryCommit(Silent:=False) Then Exit Sub
    End If

    ' 2. Toggle trashbin state and sync caption
    m_IsTrashbin = Not m_IsTrashbin
    Me.togTrashbin.Caption = IIf(m_IsTrashbin, "LIVE", "TRSH")

    ' 3. Reload live staging data when returning from trashbin
    If Not m_IsTrashbin And Not AppDataHub Is Nothing Then
        Call AppDataHub.RefreshData
    End If

    ' 4. Synchronize active child workspace viewport
    If Not Me.sfWorkspace.Form Is Nothing Then
        Call Me.sfWorkspace.Form.SyncDataMode(m_IsTrashbin)
    End If

    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProducts.togTrashbin_Click: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub cmdLanguage_Click()
    On Error GoTo ErrHandler

    ' 1. Check and save pending edits if dirty
    If Not AppDataHub Is Nothing Then
        If Not AppDataHub.TryCommit(Silent:=False) Then Exit Sub
    End If

    ' 2. Toggle locale state: EN <-> VN
    m_CurrentLocale = IIf(m_CurrentLocale = "EN", "VN", "EN")
    Me.cmdLanguage.Caption = m_CurrentLocale

    ' 3. Synchronize child workspace (Grid or Detail View)
    If Not Me.sfWorkspace.Form Is Nothing Then
        Call Me.sfWorkspace.Form.SyncLocale(m_CurrentLocale)
    End If

    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProducts.cmdLanguage_Click: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrHandler

    If KeyCode = vbKeyReturn Then
        Dim searchKeyword As String
        searchKeyword = Nz(Me.txtSearch.Text, "")
        Me.txtSearch.value = searchKeyword

        If m_CurrentView = "GRID" And Not Me.sfWorkspace.Form Is Nothing Then
            Call Me.sfWorkspace.Form.ApplySearch(searchKeyword)
        End If
        KeyCode = 0
    End If
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProducts.txtSearch_KeyDown: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub btnClearFilter_Click()
    On Error GoTo ErrHandler
    Me.txtSearch = Null

    If m_CurrentView = "GRID" And Not Me.sfWorkspace.Form Is Nothing Then
        Call Me.sfWorkspace.Form.ClearFilters
        Call Me.sfWorkspace.Form.ApplySearch("")
    End If
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProducts.btnClearFilter_Click: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub cmdRefresh_Click()
    On Error GoTo ErrHandler

    ' 1. Prompt user to commit/save pending edits before full reload
    If Not AppDataHub Is Nothing Then
        If Not AppDataHub.TryCommit(Silent:=False) Then Exit Sub
        Call AppDataHub.RefreshData
    End If

    ' 2. Re-synchronize current active workspace view
    If Not Me.sfWorkspace.Form Is Nothing Then
        Call Me.sfWorkspace.Form.SyncDataMode(m_IsTrashbin)
        Call Me.sfWorkspace.Form.SyncLocale(m_CurrentLocale)
    End If

    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProducts.cmdRefresh_Click: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub cmdSelectAll_Click()
    On Error GoTo ErrHandler
    If m_CurrentView = "GRID" And Not Me.sfWorkspace.Form Is Nothing Then
        Call Me.sfWorkspace.Form.SelectAll
    End If
    Exit Sub
ErrHandler:
    Debug.Print "[ERROR] frmProducts.cmdSelectAll_Click: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub cmdToggleView_Click()
    On Error GoTo ErrHandler

    ' 1. Check and save pending edits before swapping viewport
    If Not AppDataHub Is Nothing Then
        If Not AppDataHub.TryCommit(Silent:=False) Then Exit Sub
    End If

    ' 2. Toggle view state and swap subform viewport source
    If m_CurrentView = "GRID" Then
        m_CurrentView = "DETAIL"
        Me.cmdToggleView.Caption = "GRID"
        Me.cmdMaximize.Enabled = False
        Me.cmdViewImage.Enabled = False
        Me.togTrashbin.Enabled = False
        Me.cmdLanguage.Enabled = True
        Me.sfWorkspace.SourceObject = "frmProductDetailView"
    Else
        m_CurrentView = "GRID"
        Me.cmdToggleView.Caption = "DETL"
        Me.cmdMaximize.Enabled = True
        Me.cmdViewImage.Enabled = True
        Me.togTrashbin.Enabled = True
        Me.cmdLanguage.Enabled = True
        Me.sfWorkspace.SourceObject = "frmProductGridView"
    End If

    ' 3. Synchronize locale for newly active view
    If Not Me.sfWorkspace.Form Is Nothing Then
        Call Me.sfWorkspace.Form.SyncLocale(m_CurrentLocale)
    End If

    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProducts.cmdToggleView_Click: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub cmdMaximize_Click()
    On Error GoTo ErrHandler

    ' 1. Toggle maximize state
    m_IsMaximized = Not m_IsMaximized

    ' 2. Update button caption: MAXM <-> REST
    Me.cmdMaximize.Caption = IIf(m_IsMaximized, "REST", "MAXM")

    ' 3. Apply layout to active workspace view
    If m_CurrentView = "GRID" And Not Me.sfWorkspace.Form Is Nothing Then
        Call Me.sfWorkspace.Form.ApplyLayout(m_IsMaximized, m_IsMediaVisible)
    End If

    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProducts.cmdMaximize_Click: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub cmdViewImage_Click()
    On Error GoTo ErrHandler

    ' 1. Toggle media visibility state in Grid View
    If m_CurrentView = "GRID" And Not Me.sfWorkspace.Form Is Nothing Then
        m_IsMediaVisible = Not m_IsMediaVisible
        Call Me.sfWorkspace.Form.ApplyLayout(m_IsMaximized, m_IsMediaVisible)
    End If

    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmProducts.cmdViewImage_Click: #" & Err.Number & " - " & Err.Description
End Sub
