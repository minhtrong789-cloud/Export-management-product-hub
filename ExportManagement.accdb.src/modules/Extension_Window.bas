Attribute VB_Name = "Extension_Window"
Option Compare Database
Option Explicit
'@Folder "Extensions"

'====================================================================
' MODULE    : Extension_Window
' PURPOSE   : Window Integration & State Management Library
'====================================================================

'============================ WIN32 API =============================
#If Win64 Then ' Office 64-bit (VBA7)
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongPtrW" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
    Public Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function IsZoomed Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function LockWindowUpdate Lib "user32" (ByVal hWndLock As LongPtr) As Long

    Private Type vbaHWND
        value As LongPtr
    End Type
#ElseIf VBA7 Then ' Office 32-bit
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
    Public Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function IsZoomed Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function LockWindowUpdate Lib "user32" (ByVal hWndLock As LongPtr) As Long

    Private Type vbaHWND
        value As LongPtr
    End Type
#Else ' Office 32-bit (Legacy Access 2007 and earlier)
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long

    Private Type vbaHWND
        value As Long
    End Type
#End If

'--- WINDOWS CONSTANTS ---
Public Enum SW_nCmdShow
    SW_HIDE = 0
    SW_NORMAL = 1
    SW_MINIMIZED = 2
    SW_MAXIMIZED = 3
    SW_SHOW = 5
    SW_RESTORE = 9
    SW_SHOWDEFAULT = 10
End Enum

Public Enum GWL_nIndex
    GWL_WNDPROC = -4
    GWL_HINSTANCE = -6
    GWL_HWNDPARENT = -8
    GWL_ID = -12
    GWL_STYLE = -16
    GWL_EXSTYLE = -20
    GWL_USERDATA = -21
End Enum

Public Enum SWL_EX_dwNewLong
    WS_EX_DLGMODALFRAME = &H1
    WS_EX_NOPARENTNOTIFY = &H4
    WS_EX_TOPMOST = &H8
    WS_EX_ACCEPTFILES = &H10
    WS_EX_TRANSPARENT = &H22
    WS_EX_MDICHILD = &H40
    WS_EX_TOOLWINDOW = &H80
    WS_EX_WINDOWEDGE = &H100
    WS_EX_CLIENTEDGE = &H200
    WS_EX_CONTEXTHELP = &H400
    WS_EX_RIGHT = &H1000
    WS_EX_RTLREADING = &H2000
    WS_EX_LEFTSCROLLBAR = &H4000
    WS_EX_CONTROLPARENT = &H10000
    WS_EX_STATICEDGE = &H20000
    WS_EX_APPWINDOW = &H40000
    WS_EX_LAYERED = &H80000
    WS_EX_NOINHERITLAYOUT = &H100000
    WS_EX_LAYOUTRTL = &H400000
    WS_EX_COMPOSITED = &H2000000
    WS_EX_NOACTIVATE = &H8000000
End Enum

'====================================================================
' IN USE: CORE STANDALONE & VISIBILITY MANAGEMENT
'====================================================================

Public Sub Win_SetStandalone(ByVal hWnd As LongPtr)
    ' 1. APPLY EXTENDED WINDOW STYLE (TASKBAR VISIBILITY)
    Dim Form_ExStyle As vbaHWND
    Form_ExStyle.value = GetWindowLong(hWnd, GWL_EXSTYLE)
    Form_ExStyle.value = Form_ExStyle.value Or WS_EX_APPWINDOW
    Form_ExStyle.value = SetWindowLong(hWnd, GWL_EXSTYLE, Form_ExStyle.value)

    ' 2. HIDE MAIN ACCESS WINDOW & SHOW FORM INDEPENDENTLY
    If IsWindowVisible(hWndAccessApp) Then Call ShowWindow(hWndAccessApp, SW_HIDE)
    Call ShowWindow(hWnd, SW_SHOW)
End Sub

'--------------------------------------------------------------------

Public Sub Win_RestoreAccess(ByVal hWnd As LongPtr)
    ' 1. REMOVE EXTENDED WINDOW STYLE
    Dim Form_ExStyle As vbaHWND
    Form_ExStyle.value = GetWindowLong(hWnd, GWL_EXSTYLE)
    Form_ExStyle.value = Form_ExStyle.value And Not WS_EX_APPWINDOW
    Form_ExStyle.value = SetWindowLong(hWnd, GWL_EXSTYLE, Form_ExStyle.value)

    ' 2. RESTORE ACCESS WINDOW
    Call ShowWindow(hWndAccessApp, SW_SHOW)
End Sub

'--------------------------------------------------------------------

Public Function Win_IsAccessVisible() As Boolean
    ' PURPOSE : Checks if the main MS Access application window is visible.
    Win_IsAccessVisible = (IsWindowVisible(hWndAccessApp) <> 0)
End Function

'--------------------------------------------------------------------

Public Function Win_IsWindowVisible(ByVal hWnd As LongPtr) As Boolean
    ' PURPOSE : Checks if a specific window handle is visible via Windows OS.
    Win_IsWindowVisible = (IsWindowVisible(hWnd) <> 0)
End Function

'--------------------------------------------------------------------

Public Sub Win_SetAccessVisible(ByVal isVisible As Boolean)
    ' PURPOSE : Shows or hides the main MS Access application window.
    Call ShowWindow(hWndAccessApp, IIf(isVisible, SW_SHOW, SW_HIDE))
End Sub

'====================================================================
' EXTENDED WINDOW CONTROLS & UTILITIES
'====================================================================

Public Function Win_IsMaximized(ByVal hWnd As LongPtr) As Boolean
    ' PURPOSE : Checks if the window is system-maximized (Zoomed).
    Win_IsMaximized = (IsZoomed(hWnd) <> 0)
End Function

'--------------------------------------------------------------------

Public Function Win_IsMinimized(ByVal hWnd As LongPtr) As Boolean
    ' PURPOSE : Checks if the window is system-minimized (Iconic).
    Win_IsMinimized = (IsIconic(hWnd) <> 0)
End Function

'--------------------------------------------------------------------

Public Sub Win_Maximize(ByVal hWnd As LongPtr)
    ' PURPOSE : Forces the window to a Maximized state.
    Call ShowWindow(hWnd, SW_MAXIMIZED)
End Sub

'--------------------------------------------------------------------

Public Sub Win_Minimize(ByVal hWnd As LongPtr)
    ' PURPOSE : Forces the window to a Minimized state.
    Call ShowWindow(hWnd, SW_MINIMIZED)
End Sub

'--------------------------------------------------------------------

Public Sub Win_Restore(ByVal hWnd As LongPtr)
    ' PURPOSE : Restores the window to its normal size.
    Call ShowWindow(hWnd, SW_RESTORE)
End Sub

'--------------------------------------------------------------------

Public Sub Win_LockRedraw(ByVal hWnd As LongPtr)
    ' PURPOSE  : Freezes the visual painting of a window and its frame.
    ' PROTOCOL : Must call Win_UnlockRedraw to release system pixels.
    Call LockWindowUpdate(hWnd)
End Sub

'--------------------------------------------------------------------

Public Sub Win_UnlockRedraw()
    ' PURPOSE : Resumes visual painting for all windows.
    Call LockWindowUpdate(0)
End Sub
