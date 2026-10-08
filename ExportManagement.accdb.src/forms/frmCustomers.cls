Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'@Folder "App.Customers"

'====================================================================
' FORM      : frmCustomers
' ROLE      : Master Shell Router for Customers Domain (Phase 2 Preview)
'====================================================================

'--- [INTERNAL STATE VARIABLES] ---
Private m_IsTrashbin  As Boolean
Private m_IsMaximized As Boolean

'--- [PUBLIC STATE ACCESSORS] ---
Public Property Get IsTrashbin() As Boolean
    IsTrashbin = m_IsTrashbin
End Property

Public Property Get isMaximized() As Boolean
    isMaximized = m_IsMaximized
End Property

'====================================================================
' FORM LIFECYCLE EVENTS
'====================================================================

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrHandler

    m_IsTrashbin = False
    m_IsMaximized = False
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmCustomers.Form_Open: #" & Err.Number & " - " & Err.Description
    Cancel = True
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler

    Me.txtSearch = ""
    Me.cmdSave.Caption = "[ SAVE ]"
    Me.togTrashbin.Caption = "TRSH"
    Me.cmdMaximize.Caption = "MAXM"

    If Me.sfWorkspace.SourceObject <> "frmCustomerGridView" Then
        Me.sfWorkspace.SourceObject = "frmCustomerGridView"
    End If
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmCustomers.Form_Load: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If Not Me.sfWorkspace.Form Is Nothing Then
        If Me.sfWorkspace.Form.Dirty Then Me.sfWorkspace.Form.Dirty = False
    End If
End Sub

Private Sub Form_Close()
    On Error Resume Next
End Sub

'====================================================================
' HEADER ACTIONS & STUBS (SCHEDULED FOR PHASE 2)
'====================================================================

Private Sub cmdSave_Click()
    ' TODO: SaveCustomerChanges(ByVal silent As Boolean) As Boolean
    ' Pipeline: 1. Commit active edits -> 2. Call DataHub_Customers.TryCommit -> 3. Visual feedback
    MsgBox "Customer Hub data mutation is scheduled for Phase 2.", vbInformation, "Portfolio Scope"
End Sub

Private Sub cmdRefresh_Click()
    ' TODO: RefreshCustomerData()
    If Not Me.sfWorkspace.Form Is Nothing Then
        Me.sfWorkspace.Form.Requery
    End If
End Sub

Private Sub cmdSelectAll_Click()
    ' TODO: SelectAllCustomers(ByVal selectAll As Boolean)
    If Not Me.sfWorkspace.Form Is Nothing Then
        Call Me.sfWorkspace.Form.SelectAll
    End If
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    ' TODO: FilterCustomersByKeyword(ByVal keyword As String)
    If KeyCode = vbKeyReturn Then
        Dim searchKeyword As String: searchKeyword = Nz(Me.txtSearch.Text, "")
        Me.txtSearch.value = searchKeyword
        If Not Me.sfWorkspace.Form Is Nothing Then
            Call Me.sfWorkspace.Form.ApplySearch(searchKeyword)
        End If
        KeyCode = 0
    End If
End Sub

Private Sub btnClearFilter_Click()
    ' TODO: ResetCustomerFilters()
    Me.txtSearch = Null
    If Not Me.sfWorkspace.Form Is Nothing Then
        Call Me.sfWorkspace.Form.ClearFilters
    End If
End Sub

Private Sub cmdMaximize_Click()
    m_IsMaximized = Not m_IsMaximized
    Me.cmdMaximize.Caption = IIf(m_IsMaximized, "REST", "MAXM")
    ' TODO: ApplyCustomerLayout(ByVal isMaximized As Boolean)
End Sub

Private Sub togTrashbin_Click()
    m_IsTrashbin = (Me.togTrashbin.value = -1)
    Me.togTrashbin.Caption = IIf(m_IsTrashbin, "LIVE", "TRSH")
    ' TODO: SyncCustomerDataMode(ByVal isTrash As Boolean)
End Sub
