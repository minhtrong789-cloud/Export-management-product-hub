Attribute VB_Name = "BuildSQL"
Option Compare Database
Option Explicit

'====================================================================
' MODULE    : BuildSQL
' PURPOSE   : Manages persistent QueryDefs used as RecordSources for Forms.
'====================================================================

Public Function RebuildUIQueries() As Boolean
    On Error GoTo ErrHandler

    Dim db As DAO.Database: Set db = CurrentDb

    Call BuildSQL_Customers.BuildCustomerQueries(db)
    Call BuildSQL_Products.BuildProductQueries(db)

    Application.RefreshDatabaseWindow

    Set db = Nothing
    RebuildUIQueries = True
    Exit Function

ErrHandler:
    Set db = Nothing
    Debug.Print "[ERROR] BuildSQL.RebuildUIQueries: #" & Err.Number & " - " & Err.Description
    RebuildUIQueries = False
End Function

Public Sub UpsertQueryDef(ByVal db As DAO.Database, _
                          ByVal QueryName As String, _
                          ByVal SQLText As String)
    On Error GoTo ErrHandler

    Dim exists As Boolean: exists = False
    Dim i As Integer
    For i = 0 To db.QueryDefs.Count - 1
        If db.QueryDefs(i).name = QueryName Then
            exists = True
            Exit For
        End If
    Next i

    Dim qdf As DAO.QueryDef
    If exists Then
        Set qdf = db.QueryDefs(QueryName)
        qdf.sql = SQLText
    Else
        Set qdf = db.CreateQueryDef(QueryName, SQLText)
    End If

    Set qdf = Nothing
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] BuildSQL.UpsertQueryDef: #" & Err.Number & " - " & Err.Description
End Sub
