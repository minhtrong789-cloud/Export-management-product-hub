Attribute VB_Name = "App_Misc"
Option Compare Database
Option Explicit
'@Folder "App"

'====================================================================
' MODULE    : App_Misc
' PURPOSE   : General application utility functions and UI helpers.
'====================================================================

' Where : frmProductGridView, frmCustomerGridView
' Input : rs, "coffee", "CostPrice,Notes"
' Output: "InStr(1, Nz([Col], ''), ""coffee"") > 0 OR ..."
Public Function BuildQuickSearchFilter(ByVal rs As DAO.Recordset, _
                                      ByVal keyword As String, _
                                      Optional ByVal ignoredFields As String = "") As String
    ' 1. Guard clause: Fail fast on invalid inputs
    If rs Is Nothing Then Exit Function
    keyword = Trim(keyword)
    If Len(keyword) = 0 Then Exit Function

    On Error GoTo ErrHandler

    ' 2. Sanitize search keyword to prevent string literal breaks
    keyword = Replace(keyword, "'", "''")
    keyword = Replace(keyword, """", """""")

    Dim formattedIgnored As String
    If Len(ignoredFields) > 0 Then
        formattedIgnored = "," & Replace(ignoredFields, " ", "") & ","
    End If

    Const c_Quote As String = """"
    Dim currentField As DAO.Field
    Dim filterExpr As String

    ' 3. Iterate fields dynamically and construct InStr expressions
    For Each currentField In rs.Fields
        Dim isIgnored As Boolean
        isIgnored = False

        If Len(formattedIgnored) > 0 Then
            If InStr(1, formattedIgnored, "," & currentField.name & ",", vbTextCompare) > 0 Then
                isIgnored = True
            End If
        End If

        If Not isIgnored Then
            Select Case True
                Case currentField.name Like "*ID", _
                     currentField.name = "BeforeHash", _
                     currentField.name = "AfterHash", _
                     currentField.name = "ImagePath", _
                     currentField.name = "IsActive", _
                     currentField.name = "chkContextAction", _
                     currentField.name Like "*Date"
                    ' Skip system and internal audit fields

                Case Else
                    Select Case currentField.Type
                        Case dbText, dbMemo
                            filterExpr = filterExpr & "InStr(1, Nz([" & currentField.name & "], ''), " & c_Quote & keyword & c_Quote & ") > 0 OR "

                        Case dbLong, dbInteger, dbDouble, dbCurrency, dbSingle, dbByte, dbDecimal
                            filterExpr = filterExpr & "InStr(1, CStr(Nz([" & currentField.name & "], 0)), " & c_Quote & keyword & c_Quote & ") > 0 OR "
                    End Select
            End Select
        End If
    Next currentField

    ' 4. Finalize expression by removing trailing " OR "
    If Len(filterExpr) > 4 Then
        BuildQuickSearchFilter = Left(filterExpr, Len(filterExpr) - 4)
    End If

    Exit Function

ErrHandler:
    Debug.Print "[WARN] App_Misc.BuildQuickSearchFilter: #" & Err.Number & " - " & Err.Description
    BuildQuickSearchFilter = ""
End Function

' Where : form (continuous form row number ControlSource)
' Input : frm, "ProductID", [ProductID]
' Output: 1, 2, 3... or "*"
Public Function GetRowNumber(ByVal frm As Access.Form, _
                             ByVal keyFieldName As String, _
                             ByVal keyValue As Variant) As Variant
    ' 1. Guard clause: Handle new record or unattached form
    If frm Is Nothing Then Exit Function
    If IsNull(keyValue) Or IsEmpty(keyValue) Then
        GetRowNumber = "*"
        Exit Function
    End If
    If Len(Trim(keyFieldName)) = 0 Then Exit Function

    On Error GoTo ErrHandler

    ' 2. Acquire recordset clone from target form
    Dim rs As DAO.Recordset
    Set rs = frm.RecordsetClone
    If rs Is Nothing Then Exit Function

    ' 3. Format search criteria based on key data type
    Dim searchCriteria As String
    Select Case rs.Fields(keyFieldName).Type
        Case dbText, dbMemo
            searchCriteria = "[" & keyFieldName & "] = '" & Replace(CStr(keyValue), "'", "''") & "'"
        Case dbDate
            searchCriteria = "[" & keyFieldName & "] = #" & Format(keyValue, "yyyy-mm-dd hh:nn:ss") & "#"
        Case Else
            searchCriteria = "[" & keyFieldName & "] = " & keyValue
    End Select

    ' 4. Locate record position in clone
    rs.FindFirst searchCriteria
    If Not rs.NoMatch Then
        GetRowNumber = rs.AbsolutePosition + 1
    Else
        GetRowNumber = Null
    End If

    Exit Function

ErrHandler:
    ' Gracefully return Null during unpainted or unattached form transitions
    GetRowNumber = Null
End Function
