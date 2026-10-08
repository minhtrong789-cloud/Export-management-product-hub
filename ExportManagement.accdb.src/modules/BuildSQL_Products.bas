Attribute VB_Name = "BuildSQL_Products"
Option Compare Database
Option Explicit
'@Folder "App.Products"
'====================================================================
' MODULE    : BuildSQL_Products
' PURPOSE   : DDL Table Schemas and QueryDefs for Products Domain.
'====================================================================

'====================================================================
' 1. SCHEMA CREATION ENGINE & DDL STATEMENTS
'====================================================================

Public Sub BuildProductTables(Optional ByVal conn As Object = Nothing)
    ' PURPOSE: Execute DDL creation for all Product domain tables in dependency order.
    On Error GoTo ErrHandler

    If conn Is Nothing Then Set conn = CurrentProject.Connection

    ' 1. Level 0: Lookups
    conn.Execute SQL_CREATE_CapacityUnits
    conn.Execute SQL_CREATE_PackageTypes
    conn.Execute SQL_CREATE_MarketingCategories

    ' 2. Level 1: Category Dependent
    conn.Execute SQL_CREATE_MarketingLines

    ' 3. Level 2: Master Product Table
    conn.Execute SQL_CREATE_Products
    RefreshDatabaseWindow
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] BuildSQL_Products.BuildProductTables: #" & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "BuildProductTables", Err.Description
End Sub

Public Function SQL_CREATE_CapacityUnits() As String
    Dim strSql As String
    strSql = "CREATE TABLE CapacityUnits (" & vbCrLf
    strSql = strSql & "    CapacityUnitID AUTOINCREMENT PRIMARY KEY," & vbCrLf
    strSql = strSql & "    CapacityUnit VARCHAR(10) WITH COMPRESSION UNIQUE" & vbCrLf
    strSql = strSql & ");"
    SQL_CREATE_CapacityUnits = strSql
End Function

Public Function SQL_CREATE_PackageTypes() As String
    Dim strSql As String
    strSql = "CREATE TABLE PackageTypes (" & vbCrLf
    strSql = strSql & "    PackageTypeID AUTOINCREMENT PRIMARY KEY," & vbCrLf
    strSql = strSql & "    PackageType VARCHAR(20) WITH COMPRESSION UNIQUE" & vbCrLf
    strSql = strSql & ");"
    SQL_CREATE_PackageTypes = strSql
End Function

Public Function SQL_CREATE_MarketingCategories() As String
    Dim strSql As String
    strSql = "CREATE TABLE MarketingCategories (" & vbCrLf
    strSql = strSql & "    MarketingCategoryID AUTOINCREMENT PRIMARY KEY," & vbCrLf
    strSql = strSql & "    MarketingCategory VARCHAR(50) WITH COMPRESSION UNIQUE" & vbCrLf
    strSql = strSql & ");"
    SQL_CREATE_MarketingCategories = strSql
End Function

Public Function SQL_CREATE_MarketingLines() As String
    Dim strSql As String
    strSql = "CREATE TABLE MarketingLines (" & vbCrLf
    strSql = strSql & "    MarketingLineID AUTOINCREMENT PRIMARY KEY," & vbCrLf
    strSql = strSql & "    MarketingLine VARCHAR(50) WITH COMPRESSION UNIQUE," & vbCrLf
    strSql = strSql & "    MarketingCategoryID LONG," & vbCrLf
    strSql = strSql & "    CONSTRAINT FK_MarketingLines_MarketingCategories " & vbCrLf
    strSql = strSql & "    FOREIGN KEY (MarketingCategoryID) " & vbCrLf
    strSql = strSql & "    REFERENCES MarketingCategories (MarketingCategoryID) " & vbCrLf
    strSql = strSql & "    ON UPDATE CASCADE" & vbCrLf
    strSql = strSql & ");"
    SQL_CREATE_MarketingLines = strSql
End Function

Public Function SQL_CREATE_Products() As String
    Dim strSql As String
    strSql = "CREATE TABLE Products (" & vbCrLf
    strSql = strSql & "    ProductID AUTOINCREMENT PRIMARY KEY," & vbCrLf
    strSql = strSql & "    MarketingLineID LONG," & vbCrLf
    strSql = strSql & "    Barcode VARCHAR(100) WITH COMPRESSION UNIQUE," & vbCrLf
    strSql = strSql & "    HSCode VARCHAR(20) WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    ProductName_en VARCHAR(80) WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    ProductName_vn VARCHAR(80) WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    PackageTypeID LONG," & vbCrLf
    strSql = strSql & "    Capacity DOUBLE," & vbCrLf
    strSql = strSql & "    CapacityUnitID LONG," & vbCrLf
    strSql = strSql & "    Capacity_oz DOUBLE," & vbCrLf
    strSql = strSql & "    Price_vnd CURRENCY," & vbCrLf
    strSql = strSql & "    Price_usd CURRENCY," & vbCrLf
    strSql = strSql & "    Length_cm DOUBLE," & vbCrLf
    strSql = strSql & "    Width_cm DOUBLE," & vbCrLf
    strSql = strSql & "    Height_cm DOUBLE," & vbCrLf
    strSql = strSql & "    Weight_gram DOUBLE," & vbCrLf
    strSql = strSql & "    Ingredients TEXT WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    USP_en TEXT WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    USP_vn TEXT WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    HowToUse_en TEXT WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    HowToUse_vn TEXT WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    ImagePath TEXT WITH COMPRESSION," & vbCrLf
    strSql = strSql & "    OuterPaperBox BIT DEFAULT 0," & vbCrLf
    strSql = strSql & "    ShelfLife_month INTEGER," & vbCrLf
    strSql = strSql & "    IsActive BIT DEFAULT -1," & vbCrLf
    strSql = strSql & "    InsertDate DATETIME DEFAULT Now()," & vbCrLf
    strSql = strSql & "    ChangeDate DATETIME DEFAULT Now()," & vbCrLf
    strSql = strSql & "    CONSTRAINT FK_Products_PackageTypes" & vbCrLf
    strSql = strSql & "        FOREIGN KEY (PackageTypeID)" & vbCrLf
    strSql = strSql & "        REFERENCES PackageTypes (PackageTypeID)" & vbCrLf
    strSql = strSql & "        ON UPDATE CASCADE," & vbCrLf
    strSql = strSql & "    CONSTRAINT FK_Products_MarketingLines" & vbCrLf
    strSql = strSql & "        FOREIGN KEY (MarketingLineID)" & vbCrLf
    strSql = strSql & "        REFERENCES MarketingLines (MarketingLineID)" & vbCrLf
    strSql = strSql & "        ON UPDATE CASCADE," & vbCrLf
    strSql = strSql & "    CONSTRAINT FK_Products_CapacityUnits" & vbCrLf
    strSql = strSql & "        FOREIGN KEY (CapacityUnitID)" & vbCrLf
    strSql = strSql & "        REFERENCES CapacityUnits (CapacityUnitID)" & vbCrLf
    strSql = strSql & "        ON UPDATE CASCADE" & vbCrLf
    strSql = strSql & ");"

    SQL_CREATE_Products = strSql
End Function


'====================================================================
' 2. UI QUERIES (QueryDefs)
'====================================================================

Public Sub BuildProductQueries(Optional ByVal db As DAO.Database = Nothing)
    ' PURPOSE: Upsert saved QueryDefs used by Product Views.
    On Error GoTo ErrHandler

    If db Is Nothing Then Set db = CurrentDb

    Call BuildSQL.UpsertQueryDef(db, "Source_Products", GetProductListSQL())
    Call BuildSQL.UpsertQueryDef(db, "View_Products", GetProductViewSQL())
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] BuildSQL_Products.BuildProductQueries: #" & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "BuildProductQueries", Err.Description
End Sub

Public Function GetProductListSQL() As String
    ' PURPOSE : SQL for master product data source.
    ' SOURCE  : Products (p) with 4 LEFT JOINs to lookup tables.
    Dim strSql As String

    strSql = "SELECT " & vbCrLf
    strSql = strSql & "    p.ProductID, " & vbCrLf
    strSql = strSql & "    p.MarketingLineID, " & vbCrLf
    strSql = strSql & "    mktLines.MarketingLine, " & vbCrLf
    strSql = strSql & "    mktLines.MarketingCategoryID, " & vbCrLf
    strSql = strSql & "    mktCats.MarketingCategory, " & vbCrLf
    strSql = strSql & "    p.Barcode, " & vbCrLf
    strSql = strSql & "    p.HSCode, " & vbCrLf
    strSql = strSql & "    p.ProductName_en, " & vbCrLf
    strSql = strSql & "    p.ProductName_vn, " & vbCrLf
    strSql = strSql & "    p.PackageTypeID, " & vbCrLf
    strSql = strSql & "    pt.PackageType, " & vbCrLf
    strSql = strSql & "    p.Capacity, " & vbCrLf
    strSql = strSql & "    p.CapacityUnitID, " & vbCrLf
    strSql = strSql & "    cunits.CapacityUnit, " & vbCrLf
    strSql = strSql & "    p.Capacity_oz, " & vbCrLf
    strSql = strSql & "    p.Price_vnd, " & vbCrLf
    strSql = strSql & "    p.Price_usd, " & vbCrLf
    strSql = strSql & "    p.Length_cm, " & vbCrLf
    strSql = strSql & "    p.Width_cm, " & vbCrLf
    strSql = strSql & "    p.Height_cm, " & vbCrLf
    strSql = strSql & "    p.Weight_gram, " & vbCrLf
    strSql = strSql & "    p.Ingredients, " & vbCrLf
    strSql = strSql & "    p.USP_en, " & vbCrLf
    strSql = strSql & "    p.USP_vn, " & vbCrLf
    strSql = strSql & "    p.HowToUse_en, " & vbCrLf
    strSql = strSql & "    p.HowToUse_vn, " & vbCrLf
    strSql = strSql & "    p.ImagePath, " & vbCrLf
    strSql = strSql & "    p.OuterPaperBox, " & vbCrLf
    strSql = strSql & "    p.ShelfLife_month, " & vbCrLf
    strSql = strSql & "    p.ChangeDate, " & vbCrLf
    strSql = strSql & "    p.InsertDate, " & vbCrLf
    strSql = strSql & "    p.IsActive " & vbCrLf
    strSql = strSql & "FROM " & vbCrLf
    strSql = strSql & "    (" & vbCrLf
    strSql = strSql & "      (" & vbCrLf
    strSql = strSql & "        (" & vbCrLf
    strSql = strSql & "          Products AS p" & vbCrLf
    strSql = strSql & "          LEFT JOIN PackageTypes AS pt ON p.PackageTypeID = pt.PackageTypeID" & vbCrLf
    strSql = strSql & "        )" & vbCrLf
    strSql = strSql & "        LEFT JOIN (MarketingLines AS mktLines " & vbCrLf
    strSql = strSql & "                   LEFT JOIN MarketingCategories AS mktCats ON mktLines.MarketingCategoryID = mktCats.MarketingCategoryID) " & vbCrLf
    strSql = strSql & "          ON p.MarketingLineID = mktLines.MarketingLineID" & vbCrLf
    strSql = strSql & "      )" & vbCrLf
    strSql = strSql & "      LEFT JOIN CapacityUnits AS cunits ON p.CapacityUnitID = cunits.CapacityUnitID" & vbCrLf
    strSql = strSql & "    )" & vbCrLf
    strSql = strSql & "ORDER BY " & vbCrLf
    strSql = strSql & "    p.ProductID;"

    GetProductListSQL = strSql
End Function

Public Function GetProductViewSQL() As String
    ' PURPOSE : SQL for a unified view of the Linked Product table.
    GetProductViewSQL = "SELECT * FROM Products_Linked;"
End Function
