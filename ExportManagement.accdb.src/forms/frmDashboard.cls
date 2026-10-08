Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'@Folder "App.Dashboard"

'====================================================================
' MODULE  : frmDashboard (Form)
' PURPOSE : Central navigation hub and system state monitor for the application.
'
' EVENT PROCESS (EXECUTION ORDER):
'   [1] Form_Open   -> Sync window title & Access window toggle icon.
'   [2] Form_Load   -> Auto-select Product Hub as the primary portfolio module.
'   [3] (Interactive: cmdCallAccessWindow_Click, RouteSelection)
'   [4] Form_Unload -> Verify save (AppDataHub.TryCommit) & cleanup.
'====================================================================

'====================================================================
' LOCAL UI ASSETS (Shared Image Gallery)
'====================================================================
Private Const ICON_GEAR As String = "gear_40x40"
Private Const ICON_FILE As String = "back_40x40"

'====================================================================
' FORM EVENTS
'====================================================================
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrHandler

    If Me.Caption <> AppTitle Then Me.Caption = AppTitle
    Me.cmdCallAccessWindow.Picture = IIf(Win_IsAccessVisible, ICON_FILE, ICON_GEAR)

    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmDashboard.Form_Open: #" & Err.Number & " - " & Err.Description
    MsgBox "Could not initialize Dashboard: " & Err.Description, vbCritical
    Cancel = True
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler
    Me.navbtnProducts.SetFocus
    SendKeys "{ENTER}", True
    Call RouteSelection
    DoCmd.BrowseTo acBrowseToForm, "frmProducts", "frmDashboard.sfDashboard"

    Exit Sub
ErrHandler:
    Debug.Print "[ERROR] frmDashboard.Form_Load: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub Form_Resize()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' 1. VERIFY AND COMMIT PENDING DATA BEFORE EXIT
    If Not AppDataHub Is Nothing Then
        If Not AppDataHub.TryCommit() Then
            Cancel = True
            Exit Sub
        End If
        AppDataHub.TerminateClass
        Set AppDataHub = Nothing
    End If

    ' 2. RELEASE GLOBAL BUFFER AND PURGE TEMPDB FILE ON EXIT
    Call ReleaseAppBuffer
End Sub

'====================================================================
' CONTROL EVENTS
'====================================================================
Private Sub cmdCallAccessWindow_Click()
    On Error GoTo ErrHandler

    ' TOGGLE ACCESS WINDOW VISIBILITY STATE & SYNC ICON
    If Win_IsAccessVisible Then
        Call Win_SetStandalone(Me.hWnd)
        Call Watchdog_Start
        Me.cmdCallAccessWindow.Picture = ICON_GEAR
    Else
        Call Win_RestoreAccess(Me.hWnd)
        Call Watchdog_Stop
        Me.cmdCallAccessWindow.Picture = ICON_FILE
    End If

    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] frmDashboard.cmdCallAccessWindow_Click: #" & Err.Number & " - " & Err.Description
    MsgBox "Failed to toggle Access window visibility: " & Err.Description, vbExclamation
End Sub

'====================================================================
' SECTION 4: NAVIGATION ROUTING
'====================================================================
Public Function RouteSelection()
    ' PURPOSE: Navigation router coordinating DataHub instances and Viewport forms.
    On Error GoTo ErrHandler

    Static currentHubName As String
    Dim hubName As String: hubName = Me.ActiveControl.Tag

    If hubName = currentHubName Then Exit Function

    If hubName <> "PRODUCTS" Then
        MsgBox "Scheduled for Phase 2.", vbInformation, "Portfolio Scope"
        Exit Function
    End If

    Application.Echo False

    ' 3. REBUILD DATA HUB (TRANSITION ONLY)
    If Not AppDataHub Is Nothing Then
        If Not AppDataHub.TryCommit() Then
            Me.sfDashboard.SetFocus
            Application.Echo True
            Exit Function
        End If
        AppDataHub.TerminateClass
        Set AppDataHub = Nothing
    End If

    ' Get Data For View
    Select Case hubName
        Case "PRODUCTS"
            Dim productHub As DataHub_Products
            Set productHub = New DataHub_Products
            Call productHub.Init(GetAppBuffer())
            Set AppDataHub = productHub
        Case Else
            Set AppDataHub = Nothing
    End Select

    If Not AppDataHub Is Nothing Then
        Call AppDataHub.RefreshData
    End If

    currentHubName = hubName

    Application.Echo True
    Exit Function

ErrHandler:
    Application.Echo True
    Debug.Print "[ERROR] frmDashboard.RouteSelection: #" & Err.Number & " - " & Err.Description
End Function
