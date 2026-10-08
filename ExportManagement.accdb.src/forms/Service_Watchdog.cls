Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'@Folder "Extensions"

'====================================================================
' MODULE  : Service_Watchdog (Background Service)
' PURPOSE : Monitors active forms and manages Access window recovery.
'====================================================================

Private Sub Form_Open(Cancel As Integer)
    Me.Visible = False
    Me.TimerInterval = 1000
End Sub

Private Sub Form_Timer()
    On Error GoTo ErrHandler

    If Win_IsAccessVisible Then Exit Sub

    Dim frm As Form
    Dim hasActiveBrowseForm As Boolean: hasActiveBrowseForm = False

    For Each frm In Forms
        If frm.name <> Me.name Then
            If Win_IsWindowVisible(frm.hWnd) Then hasActiveBrowseForm = True
        End If
    Next frm

    If Not hasActiveBrowseForm Then
        If InDevMode Then
            ' Rescue Access window for user and pause watchdog timer
            Call Win_SetAccessVisible(True)
            Me.TimerInterval = 0
        Else
            ' Development: Clean up resources and close Access to prevent zombie processes
            If Not AppDataHub Is Nothing Then Set AppDataHub = Nothing
             Call ReleaseAppBuffer
            Call Application.Quit(acQuitSaveNone)
        End If
    End If

    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] Service_Watchdog.Form_Timer: #" & Err.Number & " - " & Err.Description
End Sub
