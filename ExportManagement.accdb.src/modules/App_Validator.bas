Attribute VB_Name = "App_Validator"
Option Compare Database
Option Explicit
'@Folder "App"
'====================================================================
' MODULE    : App_Validator
' PURPOSE   : Centralized existence checks for UI controls, DAO objects, and collections.
'====================================================================

'====================================================================
' ACCESS UI OBJECTS
'====================================================================

Public Function Exist_Form(ByVal formName As String) As Boolean
    On Error Resume Next
    Exist_Form = (CurrentProject.AllForms(formName).name = formName)
    On Error GoTo 0
End Function

Public Function Exist_Report(ByVal ReportName As String) As Boolean
    On Error Resume Next
    Exist_Report = (CurrentProject.AllReports(ReportName).name = ReportName)
    On Error GoTo 0
End Function

Public Function Exist_Control(ByRef frm As Form, ByVal ControlName As String) As Boolean
    On Error Resume Next
    Exist_Control = (Not frm.Controls(ControlName) Is Nothing)
    On Error GoTo 0
End Function


'====================================================================
' DATABASE OBJECTS (DAO)
'====================================================================

Public Function Exist_Table(ByVal TableName As String, Optional ByRef db As DAO.Database = Nothing) As Boolean
    If db Is Nothing Then Set db = CurrentDb

    On Error Resume Next
    Dim tdf As DAO.TableDef
    Set tdf = db.TableDefs(TableName)
    Exist_Table = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function Exist_Query(ByVal QueryName As String, Optional ByRef db As DAO.Database = Nothing) As Boolean
    If db Is Nothing Then Set db = CurrentDb

    On Error Resume Next
    Dim qdf As DAO.QueryDef
    Set qdf = db.QueryDefs(QueryName)
    Exist_Query = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function Exist_Field(ByVal TableName As String, ByVal fieldName As String, Optional ByRef db As DAO.Database = Nothing) As Boolean
    If db Is Nothing Then Set db = CurrentDb

    On Error Resume Next
    Dim fld As DAO.Field
    Set fld = db.TableDefs(TableName).Fields(fieldName)
    Exist_Field = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function Exist_Property(ByVal PropertyName As String, Optional ByVal Collection As Object = Nothing) As Boolean
    If Collection Is Nothing Then Set Collection = CurrentDb.Properties

    On Error Resume Next
    Dim prop As DAO.Property
    Set prop = Collection(PropertyName)
    Exist_Property = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function Property_Changed(ByVal NewValue As String, ByVal PropertyName As String) As Boolean
    On Error Resume Next
    Property_Changed = (CurrentDb.Properties(PropertyName).value <> NewValue)
    On Error GoTo 0
End Function

'====================================================================
' GENERIC COLLECTIONS
'====================================================================

Public Function In_Collection(ByVal ItemName As String, ByVal CollectionObject As Object) As Boolean
    On Error Resume Next
    Dim obj As Object
    Set obj = CollectionObject(ItemName)
    In_Collection = (Err.Number = 0 And Not obj Is Nothing)

    ' Double check with iteration if direct access fails (some collections don't support string keys nicely or error differently)
    If Err.Number <> 0 Then
        Err.Clear
        For Each obj In CollectionObject
            If obj.name = ItemName Then
                In_Collection = True
                Exit Function
            End If
        Next obj
    End If
    On Error GoTo 0
End Function
