Attribute VB_Name = "BuildSQL_Customers"
Option Compare Database
Option Explicit
'@Folder "App.Customers"
'====================================================================
' MODULE    : BuildSQL_Customers
' PURPOSE   : DDL Table Schemas and QueryDefs for Customers Domain.
'====================================================================

'====================================================================
' 1. SCHEMA CREATION ENGINE & DDL STATEMENTS
'====================================================================

Public Sub BuildCustomerTables(Optional ByVal conn As Object = Nothing)
    ' PURPOSE: Execute DDL creation for all Customer domain tables in dependency order.
    On Error GoTo ErrHandler

    If conn Is Nothing Then Set conn = CurrentProject.Connection

    ' 1. Master Customer Table
    conn.Execute SQL_CREATE_Customers

    ' 2. Child Dependent Tables (Foreign Keys referencing Customers)
    conn.Execute SQL_CREATE_CustomerContacts
    conn.Execute SQL_CREATE_CustomerAddresses
    conn.Execute SQL_CREATE_CustomerLogs
    RefreshDatabaseWindow
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] BuildSQL_Customers.BuildCustomerTables: #" & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "BuildCustomerTables", Err.Description
End Sub

Public Function SQL_CREATE_Customers() As String
    ' PURPOSE : SQL to create the enterprise-grade Customers table.
    ' DESIGN  : Supports B2B management with audit timestamps and soft-delete.
    Dim strSql As String
    strSql = "CREATE TABLE Customers (" & vbCrLf
    strSql = strSql & "    CustomerID           AUTOINCREMENT PRIMARY KEY," & vbCrLf
    strSql = strSql & "    CompanyName          VARCHAR(190) WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    Country              VARCHAR(50) WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    Continent            VARCHAR(50) WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    TaxID                VARCHAR(50) WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    CompanyPhone         VARCHAR(50) WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    CompanyEmail         VARCHAR(150) WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    CompanyWebsite       VARCHAR(150) WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    CompanyAddress1      TEXT WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    CompanyAddress2      TEXT WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    Notes                TEXT WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    IsActive             BIT DEFAULT -1," & vbCrLf
    strSql = strSql & "    InsertDate           DATETIME DEFAULT Now()," & vbCrLf
    strSql = strSql & "    ChangeDate           DATETIME DEFAULT Now()" & vbCrLf
    strSql = strSql & ");"
    SQL_CREATE_Customers = strSql
End Function

Public Function SQL_CREATE_CustomerContacts() As String
    ' PURPOSE : SQL to create the CustomerContacts table for 1-to-many relationship.
    Dim strSql As String
    strSql = "CREATE TABLE CustomerContacts (" & vbCrLf
    strSql = strSql & "    ContactID            AUTOINCREMENT PRIMARY KEY," & vbCrLf
    strSql = strSql & "    CustomerID           LONG," & vbCrLf
    strSql = strSql & "    ContactPerson        VARCHAR(100) WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    JobTitle             VARCHAR(80) WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    ContactPersonEmail   VARCHAR(100) WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    ContactPersonPhone   VARCHAR(50) WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    IsPrimary            BIT DEFAULT -1," & vbCrLf
    strSql = strSql & "    CONSTRAINT FK_CustomerContacts_Customers" & vbCrLf
    strSql = strSql & "        FOREIGN KEY (CustomerID)" & vbCrLf
    strSql = strSql & "        REFERENCES Customers (CustomerID)" & vbCrLf
    strSql = strSql & "        ON DELETE CASCADE" & vbCrLf
    strSql = strSql & ");"
    SQL_CREATE_CustomerContacts = strSql
End Function

Public Function SQL_CREATE_CustomerAddresses() As String
    ' PURPOSE : SQL to create the CustomerAddresses table for secondary shipping locations.
    Dim strSql As String
    strSql = "CREATE TABLE CustomerAddresses (" & vbCrLf
    strSql = strSql & "    AddressID            AUTOINCREMENT PRIMARY KEY," & vbCrLf
    strSql = strSql & "    CustomerID           LONG," & vbCrLf
    strSql = strSql & "    AddressLabel         VARCHAR(100) WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    RecipientName        VARCHAR(190) WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    AddressLine1         TEXT WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    AddressLine2         TEXT WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    Country              VARCHAR(50) WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    CONSTRAINT FK_CustomerAddresses_Customers" & vbCrLf
    strSql = strSql & "        FOREIGN KEY (CustomerID)" & vbCrLf
    strSql = strSql & "        REFERENCES Customers (CustomerID)" & vbCrLf
    strSql = strSql & "        ON DELETE CASCADE" & vbCrLf
    strSql = strSql & ");"
    SQL_CREATE_CustomerAddresses = strSql
End Function

Public Function SQL_CREATE_CustomerLogs() As String
    ' PURPOSE : SQL to create CustomerLogs table for historical notes and CRM tracking.
    Dim strSql As String
    strSql = "CREATE TABLE CustomerLogs (" & vbCrLf
    strSql = strSql & "    LogID                AUTOINCREMENT PRIMARY KEY," & vbCrLf
    strSql = strSql & "    CustomerID           LONG," & vbCrLf
    strSql = strSql & "    LogDate              DATETIME DEFAULT Now()," & vbCrLf
    strSql = strSql & "    LogTitle             VARCHAR(150) WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    LogDetails           TEXT WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    CONSTRAINT FK_CustomerLogs_Customers" & vbCrLf
    strSql = strSql & "        FOREIGN KEY (CustomerID)" & vbCrLf
    strSql = strSql & "        REFERENCES Customers (CustomerID)" & vbCrLf
    strSql = strSql & "        ON DELETE CASCADE" & vbCrLf
    strSql = strSql & ");"
    SQL_CREATE_CustomerLogs = strSql
End Function

'====================================================================
' 2. UI QUERIES (QueryDefs)
'====================================================================

Public Sub BuildCustomerQueries(Optional ByVal db As DAO.Database = Nothing)
    ' PURPOSE: Upsert saved QueryDefs used by Customer Views.
    On Error GoTo ErrHandler

    If db Is Nothing Then Set db = CurrentDb

    Call BuildSQL.UpsertQueryDef(db, "Source_Customers", GetCustomerListSQL())
    Call BuildSQL.UpsertQueryDef(db, "View_Customers", GetCustomerViewSQL())
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] BuildSQL_Customers.BuildCustomerQueries: #" & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "BuildCustomerQueries", Err.Description
End Sub

Public Function GetCustomerListSQL() As String
    ' PURPOSE: SQL for master customer data source joining primary contact.
    Dim strSql As String

    strSql = "SELECT " & vbCrLf
    strSql = strSql & "    c.CustomerID, " & vbCrLf
    strSql = strSql & "    c.CompanyName, " & vbCrLf
    strSql = strSql & "    c.Country, " & vbCrLf
    strSql = strSql & "    c.Continent, " & vbCrLf
    strSql = strSql & "    cc.ContactPerson, " & vbCrLf
    strSql = strSql & "    cc.ContactPersonPhone, " & vbCrLf
    strSql = strSql & "    cc.ContactPersonEmail, " & vbCrLf
    strSql = strSql & "    c.CompanyPhone, " & vbCrLf
    strSql = strSql & "    c.CompanyEmail, " & vbCrLf
    strSql = strSql & "    c.CompanyWebsite, " & vbCrLf
    strSql = strSql & "    c.CompanyAddress1, " & vbCrLf
    strSql = strSql & "    c.CompanyAddress2, " & vbCrLf
    strSql = strSql & "    c.TaxID, " & vbCrLf
    strSql = strSql & "    c.Notes, " & vbCrLf
    strSql = strSql & "    c.IsActive, " & vbCrLf
    strSql = strSql & "    c.InsertDate, " & vbCrLf
    strSql = strSql & "    c.ChangeDate " & vbCrLf
    strSql = strSql & "FROM " & vbCrLf
    strSql = strSql & "    Customers AS c " & vbCrLf
    strSql = strSql & "    LEFT JOIN (SELECT * FROM CustomerContacts WHERE IsPrimary = True) AS cc " & vbCrLf
    strSql = strSql & "      ON c.CustomerID = cc.CustomerID;"

    GetCustomerListSQL = strSql
End Function

Public Function GetCustomerViewSQL() As String
    ' PURPOSE: SQL for filtered active records from linked snapshot.
    GetCustomerViewSQL = "SELECT * FROM Customers_Linked WHERE IsActive = TRUE;"
End Function
