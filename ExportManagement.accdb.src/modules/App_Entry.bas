Attribute VB_Name = "App_Entry"

Option Compare Database
Option Explicit
'@Folder "App.Dashboard"

'====================================================================
' MODULE    : App_Entry
' PURPOSE   : Application Initialization and Bootstrapper Controller.
'====================================================================

'--- [SYSTEM CONFIGURATION] ---
Public Const InDevMode        As Boolean = False
Public Const AppTitle         As String = "Export Manager Software"
Public Const RELPATH_APP_ICON As String = "Export_Management.ico"

'--- [GLOBAL APPLICATION STATE] ---
Public AppPath       As String
Public AppDataHub    As Object
Private AppBuffer     As StagingBuffer

Public Function GetAppBuffer() As StagingBuffer
    If AppBuffer Is Nothing Then
        Set AppBuffer = New StagingBuffer
    End If
    Set GetAppBuffer = AppBuffer
End Function

Public Sub ReleaseAppBuffer()
    Set AppBuffer = Nothing
End Sub

Public Function EntryPoint() As Boolean
    On Error GoTo ErrHandler
    Call Application.Echo(False)

    Call Win_SetAccessVisible(False)

    AppPath = CurrentProject.Path

    Dim IconPath As String: IconPath = CurrentProject.Path & "\" & RELPATH_APP_ICON
    If Dir(IconPath) <> "" Then
        Call SetDBProperty("AppIcon", IconPath)
        Call SetDBProperty("UseAppIconForFrmRpt", True)
    End If

    Call SetDBProperty("AppTitle", AppTitle)
    Application.RefreshTitleBar

    If Not CurrentProject.AllForms("frmDashboard").IsLoaded Then
        DoCmd.OpenForm "frmDashboard"
    End If
    Call Win_SetStandalone(Forms("frmDashboard").hWnd)

    Call Watchdog_Start

    Call Application.Echo(True)
    Debug.Print "[SUCCESS] Application bootstrapper executed successfully."
    EntryPoint = True
    Exit Function

ErrHandler:
    Call Win_SetAccessVisible(True)
    Call Application.Echo(True)
    Debug.Print "[ERROR] App_Entry.EntryPoint: #" & Err.Number & " - " & Err.Description
    MsgBox "Critical Startup Error [" & Err.Number & "]: " & Err.Description, vbCritical, "Boot Failure"
    EntryPoint = False
End Function

'====================================================================
' INTERNAL HELPERS
'====================================================================

Private Function GetDBProperty(ByVal PropName As String) As Variant
    On Error Resume Next
    GetDBProperty = CurrentDb.Properties(PropName).value
    On Error GoTo 0
End Function

Private Sub SetDBProperty(ByVal PropName As String, ByVal PropValue As Variant)
    On Error GoTo CatchProp
    Dim db As DAO.Database: Set db = CurrentDb

    db.Properties(PropName).value = PropValue
    Exit Sub

CatchProp:
    If Err.Number = 3270 Then ' Property not found
        Err.Clear
        Dim PropType As DAO.DataTypeEnum
        Select Case VarType(PropValue)
        Case vbBoolean: PropType = dbBoolean
        Case vbLong, vbInteger: PropType = dbLong
        Case vbDouble, vbSingle: PropType = dbDouble
        Case Else: PropType = dbText
        End Select

        Dim prop As DAO.Property
        Set prop = db.CreateProperty(PropName, PropType, PropValue)
        db.Properties.Append prop
    End If
    Err.Clear
End Sub

Public Function GetProjectPath() As String
    If LenB(AppPath) = 0 Then AppPath = CurrentProject.Path
    GetProjectPath = AppPath
End Function

'====================================================================
' WATCHDOG SERVICE CONTROLLER
'====================================================================

Public Sub Watchdog_Start()
    On Error GoTo ErrHandler

    If CurrentProject.AllForms("Service_Watchdog").IsLoaded Then
        Forms("Service_Watchdog").TimerInterval = 1000
    Else
        DoCmd.OpenForm "Service_Watchdog", WindowMode:=acHidden
    End If

    Exit Sub

ErrHandler:
    Call Win_SetAccessVisible(True)
    Debug.Print "[ERROR] App_Entry.Watchdog_Start: #" & Err.Number & " - " & Err.Description
End Sub

Public Sub Watchdog_Stop()
    On Error GoTo ErrHandler

    If CurrentProject.AllForms("Service_Watchdog").IsLoaded Then
        Forms("Service_Watchdog").TimerInterval = 0
    End If

    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] App_Entry.Watchdog_Stop: #" & Err.Number & " - " & Err.Description
End Sub
