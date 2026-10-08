Attribute VB_Name = "Datahub_DataSync"
Option Compare Database
Option Explicit
'@Folder "Extensions.Data"

'====================================================================
' MODULE  : Datahub_DataSync
' PURPOSE : SQL-based DML operations to synchronize staging buffer
'           records with backend tables.
'====================================================================

Public Sub SyncLookupRecord(ByVal targetTable As String, _
                            ByVal sourceTable As String, _
                            ParamArray FieldNames() As Variant)
    If UBound(FieldNames) < LBound(FieldNames) Then Exit Sub

    Dim sqlFields As String, sqlSelect As String, sqlJoin As String
    Dim i As Integer, fieldName As String

    For i = LBound(FieldNames) To UBound(FieldNames)
        fieldName = FieldNames(i)
        sqlFields = sqlFields & "[" & fieldName & "], "
        sqlSelect = sqlSelect & "Src.[" & fieldName & "], "
        sqlJoin = sqlJoin & "Src.[" & fieldName & "] = Tgt.[" & fieldName & "] AND "
    Next i

    sqlFields = Left(sqlFields, Len(sqlFields) - 2)
    sqlSelect = Left(sqlSelect, Len(sqlSelect) - 2)
    sqlJoin = Left(sqlJoin, Len(sqlJoin) - 5)

    Dim strSql As String
    strSql = "INSERT INTO [" & targetTable & "] (" & sqlFields & ") " & _
             "SELECT DISTINCT " & sqlSelect & " " & _
             "FROM [" & sourceTable & "] AS Src " & _
             "LEFT JOIN [" & targetTable & "] AS Tgt ON " & sqlJoin & " " & _
             "WHERE Tgt.[" & FieldNames(0) & "] IS NULL " & _
             "  AND (Src.BeforeHash <> Src.AfterHash OR Src.BeforeHash IS NULL) " & _
             "  AND Src.[" & FieldNames(0) & "] IS NOT NULL;"

    Call CurrentDb.Execute(strSql, dbFailOnError)
End Sub

Public Sub SyncMapLookupID(ByVal tempTableName As String, _
                           ByVal fkFieldName As String, _
                           ByVal lookupTableName As String, _
                           ParamArray FieldNames() As Variant)
    On Error GoTo ErrHandler

    If UBound(FieldNames) < LBound(FieldNames) Then Exit Sub

    Dim i As Integer
    Dim sql_join As String: sql_join = ""

    For i = LBound(FieldNames) To UBound(FieldNames)
        Dim field_name As String: field_name = FieldNames(i)
        sql_join = sql_join & "Nz([" & tempTableName & "].[" & field_name & "], 0) = Nz(Tgt.[" & field_name & "], 0) AND "
    Next i

    sql_join = Left(sql_join, Len(sql_join) - 5)

    Dim final_SQL As String
    final_SQL = " UPDATE [" & tempTableName & "] " & _
             "      INNER JOIN [" & lookupTableName & "] AS Tgt " & _
             "          ON " & sql_join & _
             "      SET [" & tempTableName & "].[" & fkFieldName & "] = Tgt.[" & fkFieldName & "] " & _
             "          WHERE ([" & tempTableName & "].[" & fkFieldName & "] IS NULL " & _
             "          OR [" & tempTableName & "].[" & fkFieldName & "] <> Tgt.[" & fkFieldName & "]);"

    Call CurrentDb.Execute(final_SQL, dbFailOnError)
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "DataSync.SyncMapLookupID", Err.Description
End Sub

Public Sub SyncModifiedRecords(ByVal targetTable As String, _
                               ByVal sourceTable As String, _
                               ByVal fieldName As String, _
                               ByVal idFieldName As String)
    Dim updateSql As String

    updateSql = "UPDATE [" & targetTable & "] " & _
                "INNER JOIN [" & sourceTable & "] AS Src " & _
                "  ON Src.[" & idFieldName & "] = [" & targetTable & "].[" & idFieldName & "] " & _
                "SET [" & targetTable & "].[" & fieldName & "] = Src.[" & fieldName & "] " & _
                "WHERE Src.BeforeHash <> Src.AfterHash " & _
                "  AND Nz([" & targetTable & "].[" & fieldName & "], """") <> Nz(Src.[" & fieldName & "], """");"

    Call CurrentDb.Execute(updateSql, dbFailOnError)
End Sub

Public Sub SyncNewRecords(ByVal targetTable As String, _
                          ByVal sourceTable As String, _
                          ByVal fieldList As String, _
                          ByVal valueList As String)
    Dim insertSql As String

    insertSql = "INSERT INTO [" & targetTable & "] (" & fieldList & ") " & _
                "SELECT " & valueList & " " & _
                "FROM [" & sourceTable & "] " & _
                "WHERE BeforeHash IS NULL AND IsActive = True;"

    Call CurrentDb.Execute(insertSql, dbFailOnError)
End Sub
