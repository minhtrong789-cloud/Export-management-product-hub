Attribute VB_Name = "DataSeed_Products"
Option Compare Database
Option Explicit

'====================================================================
' MODULE    : DataSeed_Products
' PURPOSE   : Populate sample data for Product Hub (Lookups & 15 SKUs).
'====================================================================


Public Sub Seed_ProductHub()
    On Error GoTo ErrHandler
    Dim db As DAO.Database: Set db = CurrentDb

    ' 1. Seed lookup dependencies
    Call Seed_CapacityUnits(db)
    Call Seed_PackageTypes(db)
    Call Seed_MarketingCategories(db)
    Call Seed_MarketingLines(db)

    ' 2. Seed Master Products
    Call Seed_Products(db)

    Set db = Nothing
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] DataSeed_Products.Seed_ProductHub: #" & Err.Number & " - " & Err.Description
    Set db = Nothing
End Sub


Public Sub Wipe_ProductHub(Optional ByVal db As DAO.Database = Nothing)
    On Error GoTo ErrHandler
    Dim isLocalDb As Boolean
    If db Is Nothing Then
        Set db = CurrentDb
        isLocalDb = True
    End If

    ' 1. Drop Product Hub tables in reverse dependency order
    Call DropTableIfExists(db, "Products")
    Call DropTableIfExists(db, "MarketingLines")
    Call DropTableIfExists(db, "MarketingCategories")
    Call DropTableIfExists(db, "PackageTypes")
    Call DropTableIfExists(db, "CapacityUnits")

    ' 2. Release DAO reference before ADO DDL recreation
    If isLocalDb Then Set db = Nothing

    ' 3. Recreate fresh tables with AutoNumber starting at 1
    Call BuildSQL_Products.BuildProductTables

    DBEngine.Idle dbRefreshCache
    Debug.Print "[OK] Product Hub tables recreated fresh with AutoNumber reset to 1."
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] DataSeed_Products.Wipe_ProductHub: #" & Err.Number & " - " & Err.Description
    If isLocalDb Then Set db = Nothing
End Sub

Public Sub Seed_Products(Optional ByVal db As DAO.Database = Nothing)
    On Error GoTo ErrHandler
    Dim isLocalDb As Boolean
    If db Is Nothing Then
        Set db = CurrentDb
        isLocalDb = True
    End If

    If Not IsTableEmpty("Products") Then
        Debug.Print "[SKIP] Products: Table already has data."
        If isLocalDb Then Set db = Nothing
        Exit Sub
    End If

    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("Products", dbOpenDynaset)

        ' P01: Coffee Body Polish | Line=2 Pkg=Jar(2) Unit=g(2)
        Call SeedProduct(rs, 2, "893617304001", "3304.99", _
            "Dak Lak Coffee Body Polish", "Tay te bao chet ca phe Dak Lak", _
            2, 200, 2, 7.05, 115000, 5#, 7.5, 7.5, 8#, 260, _
            "Coffea Arabica Seed Powder, Cocos Nucifera Oil, Butyrospermum Parkii Butter, Glycerin, Aqua", _
            "Awaken skin with natural Dak Lak coffee. Exfoliates and energises.", _
            "Da sang min voi ca phe Dak Lak. Tay te bao va tao sinh luc.", _
            "Apply to damp skin. Massage circular motions 2 min. Rinse.", _
            "Thoa len da am. Massage vong tron 2 phut. Xa sach.", _
            0, 24, True)

        ' P02: Pomelo Hair Tonic | Line=3 Pkg=Bottle(1) Unit=ml(1)
        Call SeedProduct(rs, 3, "893617304002", "3305.10", _
            "Pomelo Hair Tonic", "Nuoc duong toc tinh dau buoi", _
            1, 140, 1, 4.73, 145000, 6.5, 4.5, 4.5, 17#, 175, _
            "Citrus Maxima Peel Oil, Panthenol, Biotin, Niacinamide, Aqua", _
            "Strengthens follicles with pure pomelo oil. Reduces hair fall in 4 weeks.", _
            "Tang cuong nang toc voi tinh dau buoi. Giam rung toc trong 4 tuan.", _
            "Apply 5-6 drops to scalp after washing. Massage 3 min. Do not rinse.", _
            "Nho 5-6 giot len da dau sau khi goi. Massage 3 phut. Khong xa lai.", _
            1, 36, True)

        ' P03: Winter Melon Gel Cleanser | Line=1 Pkg=Pump(5) Unit=ml(1)
        Call SeedProduct(rs, 1, "893617304003", "3304.99", _
            "Winter Melon Facial Gel Cleanser", "Gel rua mat bi dao", _
            5, 140, 1, 4.73, 175000, 7.5, 4.5, 4.5, 16.5, 185, _
            "Benincasa Hispida Extract, Sodium Hyaluronate, Niacinamide, Allantoin, Aqua", _
            "Gentle cleanser infused with winter melon. Brightens and hydrates.", _
            "Sua rua mat diu nhe tu bi dao. Lam sang va cap am cho da.", _
            "Wet face. Pump 1-2 times, lather, massage 30s. Rinse.", _
            "Lam uot mat. Bom 1-2 lan, tao bot, massage 30 giay. Xa sach.", _
            1, 36, True)

        ' P04: Turmeric Face Cream (INACTIVE) | Line=4 Pkg=Jar(2) Unit=g(2)
        Call SeedProduct(rs, 4, "893617304004", "3304.99", _
            "Hung Yen Turmeric Face Cream", "Kem nghe Hung Yen", _
            2, 30, 2, 1.06, 245000, 10.5, 6#, 6#, 5.5, 90, _
            "Curcuma Longa Root Extract, Centella Asiatica Extract, Shea Butter, Niacinamide", _
            "Rich overnight repair with Hung Yen turmeric. Fades dark spots.", _
            "Kem duong phuc hoi qua dem voi nghe Hung Yen. Mo tham nam.", _
            "Apply pea-sized to clean dry face at night. Pat until absorbed.", _
            "Lay luong bang hat dau thoa len mat sach kho. Vo nhe cho tham.", _
            1, 24, False)

        ' P05: Salicylic Acid 2% Toner | Line=5 Pkg=Pump(5) Unit=ml(1)
        Call SeedProduct(rs, 5, "893617304005", "3304.99", _
            "Salicylic Acid 2% Clarifying Toner", "Nuoc hoa hong BHA 2% lam sach lo chan long", _
            5, 100, 1, 3.38, 195000, 8.5, 5#, 5#, 15#, 140, _
            "Salicylic Acid 2%, Niacinamide 5%, Panthenol, Witch Hazel Extract, Aqua", _
            "Unclogs pores with 2% BHA. Reduces acne and blackheads in 2 weeks.", _
            "Thong thoang lo chan long voi BHA 2%. Giam mun va dau den trong 2 tuan.", _
            "Sweep over face with cotton pad after cleansing. AM and PM.", _
            "Lau khap mat bang bong tay sau khi rua mat. Sang va toi.", _
            1, 24, True)

        ' P06: Coffee Eye Cream | Line=2 Pkg=Tube(3) Unit=g(2)
        Call SeedProduct(rs, 2, "893617304006", "3304.99", _
            "Dak Lak Coffee Eye Cream", "Kem mat ca phe Dak Lak", _
            3, 15, 2, 0.53, 185000, 8#, 2.5, 2.5, 10#, 35, _
            "Coffea Arabica Seed Extract, Caffeine, Peptides, Hyaluronic Acid, Squalane", _
            "Depuffs and brightens under-eye area with concentrated coffee extract.", _
            "Giam bong mat va lam sang vung mat voi chiet xuat ca phe co dac.", _
            "Dab small amount under eyes using ring finger. Morning and night.", _
            "Cham mot luong nho duoi mat bang ngon ap ut. Sang va toi.", _
            1, 24, True)

        ' P07: Pomelo Shampoo | Line=3 Pkg=Pump(5) Unit=ml(1)
        Call SeedProduct(rs, 3, "893617304007", "3305.10", _
            "Pomelo Anti-Hair Loss Shampoo", "Dau goi buoi chong rung toc", _
            5, 310, 1, 10.48, 165000, 7#, 7#, 7#, 18#, 380, _
            "Citrus Maxima Peel Oil, Biotin, Caffeine, Keratin, Coco-Glucoside, Aqua", _
            "Sulfate-free formula fortified with pomelo oil. Reduces breakage 60%.", _
            "Cong thuc khong sulfate voi tinh dau buoi. Giam gay rung 60%.", _
            "Apply to wet hair. Lather and massage scalp 2 min. Rinse well.", _
            "Thoa len toc uot. Tao bot va massage da dau 2 phut. Xa sach.", _
            1, 24, True)

        ' P08: Pomelo Conditioner | Line=3 Pkg=Pump(5) Unit=ml(1)
        Call SeedProduct(rs, 3, "893617304008", "3305.10", _
            "Pomelo Strengthening Conditioner", "Dau xa buoi phuc hoi toc", _
            5, 310, 1, 10.48, 175000, 7.5, 7#, 7#, 18#, 395, _
            "Citrus Maxima Peel Oil, Argan Oil, Hydrolyzed Keratin, Panthenol, Cetyl Alcohol", _
            "Deep conditioning with pomelo and argan oil. Silky smooth in one wash.", _
            "Duong sau voi tinh dau buoi va argan. Toc muot ma sau 1 lan goi.", _
            "After shampoo, apply mid-length to ends. Wait 2-3 min. Rinse.", _
            "Sau khi goi, thoa tu giua den duoi toc. Cho 2-3 phut. Xa sach.", _
            1, 24, True)

        ' P09: Winter Melon Hydrating Toner | Line=1 Pkg=Pump(5) Unit=ml(1)
        Call SeedProduct(rs, 1, "893617304009", "3304.99", _
            "Winter Melon Hydrating Toner", "Nuoc hoa hong bi dao cap am", _
            5, 140, 1, 4.73, 185000, 8#, 4.5, 4.5, 16.5, 178, _
            "Benincasa Hispida Water, Hyaluronic Acid, Panthenol, Aloe Barbadensis Leaf Juice", _
            "Lightweight watery toner with 92% winter melon water. Instant hydration boost.", _
            "Toner dang nuoc nhe voi 92% nuoc bi dao. Cap am tuc thi.", _
            "Pour onto palms or cotton pad. Pat gently into skin after cleansing.", _
            "Do ra long ban tay hoac bong tay. Vo nhe len da sau khi rua mat.", _
            0, 36, True)

        ' P10: Winter Melon Brightening Serum | Line=1 Pkg=Bottle(1) Unit=ml(1)
        Call SeedProduct(rs, 1, "893617304010", "3304.99", _
            "Winter Melon Brightening Serum", "Serum bi dao lam sang da", _
            1, 30, 1, 1.01, 295000, 12.5, 3.5, 3.5, 12#, 65, _
            "Benincasa Hispida Fruit Extract 10%, Niacinamide 5%, Alpha Arbutin 2%, Hyaluronic Acid", _
            "Concentrated brightening serum. Fades dark spots in 14 days.", _
            "Serum co dac lam sang. Mo tham nam chi trong 14 ngay.", _
            "Apply 2-3 drops to face after toner. Pat in gently. Use AM/PM.", _
            "Nho 2-3 giot len mat sau toner. Vo nhe. Dung sang va toi.", _
            1, 24, True)

        ' P11: Rose Facial Mist | Line=6 Pkg=Spray(6) Unit=ml(1)
        Call SeedProduct(rs, 6, "893617304011", "3304.99", _
            "Bulgarian Rose Facial Mist", "Xit khoang hoa hong Bulgaria", _
            6, 130, 1, 4.39, 155000, 6.5, 4#, 4#, 16#, 165, _
            "Rosa Damascena Flower Water, Glycerin, Allantoin, Panthenol, Aqua", _
            "Ultra-fine mist with real Bulgarian rose water. Refreshes and sets makeup.", _
            "Xit sieu min tu nuoc hoa hong Bulgaria. Lam tuoi va giu lop trang diem.", _
            "Hold 15-20cm from face. Spray 2-3 times. Use over or under makeup.", _
            "Giu cach mat 15-20cm. Xit 2-3 lan. Dung tren hoac duoi lop makeup.", _
            1, 36, True)

        ' P12: Coconut Body Oil | Line=7 Pkg=Bottle(1) Unit=ml(1)
        Call SeedProduct(rs, 7, "893617304012", "3304.99", _
            "Ben Tre Virgin Coconut Body Oil", "Dau dua nguyen chat Ben Tre duong the", _
            1, 200, 1, 6.76, 135000, 5.5, 5#, 5#, 18#, 240, _
            "Cocos Nucifera Oil, Tocopheryl Acetate, Simmondsia Chinensis Seed Oil", _
            "Cold-pressed Ben Tre coconut oil. Deep moisture for body and hair.", _
            "Dau dua ep lanh Ben Tre. Duong am sau cho da va toc.", _
            "Warm a few drops in palms. Massage onto body or hair ends.", _
            "Lam am vai giot trong long ban tay. Massage len co the hoac duoi toc.", _
            0, 36, True)

        ' P13: Green Tea Gentle Cleanser | Line=8 Pkg=Tube(3) Unit=ml(1)
        Call SeedProduct(rs, 8, "893617304013", "3304.99", _
            "Green Tea Gentle Foam Cleanser", "Sua rua mat tra xanh diu nhe", _
            3, 100, 1, 3.38, 125000, 5.5, 4#, 4#, 14#, 135, _
            "Camellia Sinensis Leaf Extract, Aloe Barbadensis, Glycerin, Coco-Glucoside, Aqua", _
            "Ultra-gentle daily cleanser with antioxidant-rich green tea. pH 5.5 balanced.", _
            "Sua rua mat diu nhe hang ngay voi tra xanh giau chong oxy hoa. pH 5.5.", _
            "Squeeze small amount. Lather with water. Massage face 30s. Rinse.", _
            "Nan mot luong nho. Tao bot voi nuoc. Massage mat 30 giay. Xa sach.", _
            1, 24, True)

        ' P14: Coffee Lip Scrub | Line=2 Pkg=Jar(2) Unit=g(2)
        Call SeedProduct(rs, 2, "893617304014", "3304.99", _
            "Dak Lak Coffee Lip Scrub", "Tay te bao chet moi ca phe Dak Lak", _
            2, 15, 2, 0.53, 85000, 3.5, 3#, 3#, 3.5, 30, _
            "Coffea Arabica Seed Powder, Sucrose, Cocos Nucifera Oil, Tocopherol, Honey Extract", _
            "Gently buffs away dry lip skin with fine coffee grounds and honey.", _
            "Loai bo da chet moi nhe nhang voi bot ca phe min va mat ong.", _
            "Apply small amount to lips. Rub gently in circles. Wipe or lick off.", _
            "Thoa luong nho len moi. Xoa nhe theo vong tron. Lau hoac liem sach.", _
            0, 12, True)

        ' P15: Turmeric Brightening Serum (INACTIVE) | Line=4 Pkg=Bottle(1) Unit=ml(1)
        Call SeedProduct(rs, 4, "893617304015", "3304.99", _
            "Hung Yen Turmeric Brightening Serum", "Serum nghe Hung Yen lam sang da", _
            1, 30, 1, 1.01, 275000, 12#, 3.5, 3.5, 12#, 62, _
            "Curcuma Longa Root Extract 15%, Niacinamide 10%, Tranexamic Acid, Hyaluronic Acid", _
            "Potent brightening serum with 15% turmeric extract. Evens skin tone.", _
            "Serum lam sang manh voi 15% chiet xuat nghe. Deu mau da.", _
            "Apply 2-3 drops to clean face at night. Follow with moisturizer.", _
            "Nho 2-3 giot len mat sach vao buoi toi. Thoa kem duong sau do.", _
            1, 24, False)

    rs.Close
    Set rs = Nothing
    DBEngine.Idle dbRefreshCache
    Debug.Print "[OK] Products seeded: 15 SKUs."

    If isLocalDb Then Set db = Nothing
    Exit Sub

ErrHandler:
    Debug.Print "[ERROR] DataSeed_Products.Seed_Products: #" & Err.Number & " - " & Err.Description
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If isLocalDb Then Set db = Nothing
End Sub

Private Sub Seed_CapacityUnits(ByVal db As DAO.Database)
    If Not IsTableEmpty("CapacityUnits") Then Exit Sub

    db.Execute "INSERT INTO CapacityUnits (CapacityUnit) VALUES ('ml')", dbFailOnError
    db.Execute "INSERT INTO CapacityUnits (CapacityUnit) VALUES ('g')", dbFailOnError
    db.Execute "INSERT INTO CapacityUnits (CapacityUnit) VALUES ('pcs')", dbFailOnError
    db.Execute "INSERT INTO CapacityUnits (CapacityUnit) VALUES ('fl oz')", dbFailOnError
    DBEngine.Idle dbRefreshCache
End Sub

Private Sub Seed_PackageTypes(ByVal db As DAO.Database)
    If Not IsTableEmpty("PackageTypes") Then Exit Sub

    db.Execute "INSERT INTO PackageTypes (PackageType) VALUES ('Bottle')", dbFailOnError
    db.Execute "INSERT INTO PackageTypes (PackageType) VALUES ('Jar')", dbFailOnError
    db.Execute "INSERT INTO PackageTypes (PackageType) VALUES ('Tube')", dbFailOnError
    db.Execute "INSERT INTO PackageTypes (PackageType) VALUES ('Box')", dbFailOnError
    db.Execute "INSERT INTO PackageTypes (PackageType) VALUES ('Pump Bottle')", dbFailOnError
    db.Execute "INSERT INTO PackageTypes (PackageType) VALUES ('Spray Bottle')", dbFailOnError
    DBEngine.Idle dbRefreshCache
End Sub

Private Sub Seed_MarketingCategories(ByVal db As DAO.Database)
    If Not IsTableEmpty("MarketingCategories") Then Exit Sub

    db.Execute "INSERT INTO MarketingCategories (MarketingCategory) VALUES ('Skin Care')", dbFailOnError
    db.Execute "INSERT INTO MarketingCategories (MarketingCategory) VALUES ('Hair Care')", dbFailOnError
    db.Execute "INSERT INTO MarketingCategories (MarketingCategory) VALUES ('Body Care')", dbFailOnError
    db.Execute "INSERT INTO MarketingCategories (MarketingCategory) VALUES ('Lip Care')", dbFailOnError
    DBEngine.Idle dbRefreshCache
End Sub

Private Sub Seed_MarketingLines(ByVal db As DAO.Database)
    ' Marketing Lines (FK: 1=Skin Care, 2=Hair Care, 3=Body Care, 4=Lip Care)
    If Not IsTableEmpty("MarketingLines") Then Exit Sub

    db.Execute "INSERT INTO MarketingLines (MarketingLine, MarketingCategoryID) VALUES ('Winter Melon', 1)", dbFailOnError
    db.Execute "INSERT INTO MarketingLines (MarketingLine, MarketingCategoryID) VALUES ('Dak Lak Coffee', 3)", dbFailOnError
    db.Execute "INSERT INTO MarketingLines (MarketingLine, MarketingCategoryID) VALUES ('Pomelo', 2)", dbFailOnError
    db.Execute "INSERT INTO MarketingLines (MarketingLine, MarketingCategoryID) VALUES ('Hung Yen Turmeric', 1)", dbFailOnError
    db.Execute "INSERT INTO MarketingLines (MarketingLine, MarketingCategoryID) VALUES ('Salicylic Acid', 1)", dbFailOnError
    db.Execute "INSERT INTO MarketingLines (MarketingLine, MarketingCategoryID) VALUES ('Rose', 1)", dbFailOnError
    db.Execute "INSERT INTO MarketingLines (MarketingLine, MarketingCategoryID) VALUES ('Ben Tre Coconut', 3)", dbFailOnError
    db.Execute "INSERT INTO MarketingLines (MarketingLine, MarketingCategoryID) VALUES ('Green Tea', 1)", dbFailOnError
    DBEngine.Idle dbRefreshCache
End Sub


'====================================================================
' PRIVATE HELPERS
'====================================================================
Private Function IsTableEmpty(ByVal tblName As String) As Boolean
    On Error Resume Next
    Dim cnt As Long
    cnt = DCount("*", tblName)
    If Err.Number <> 0 Then cnt = 0: Err.Clear
    IsTableEmpty = (cnt = 0)
End Function

Private Sub SeedProduct( _
    ByVal rs As DAO.Recordset, _
    ByVal lineID As Long, ByVal Barcode As String, ByVal HSCode As String, _
    ByVal nameEn As String, ByVal nameVn As String, _
    ByVal pkgID As Long, ByVal cap As Double, ByVal unitID As Long, ByVal capOz As Double, _
    ByVal priceVnd As Currency, ByVal priceUsd As Currency, _
    ByVal lenCm As Double, ByVal widCm As Double, ByVal hgtCm As Double, ByVal wtGram As Double, _
    ByVal ingr As String, _
    ByVal uspEn As String, ByVal uspVn As String, _
    ByVal howEn As String, ByVal howVn As String, _
    ByVal outerBox As Long, ByVal shelfMo As Long, _
    ByVal isActive As Boolean)

    rs.AddNew
    rs!MarketingLineID = lineID
    rs!Barcode = Barcode
    rs!HSCode = HSCode
    rs!ProductName_en = nameEn
    rs!ProductName_vn = nameVn
    rs!PackageTypeID = pkgID
    rs!Capacity = cap
    rs!CapacityUnitID = unitID
    rs!Capacity_oz = capOz
    rs!Price_vnd = priceVnd
    rs!Price_usd = priceUsd
    rs!Length_cm = lenCm
    rs!Width_cm = widCm
    rs!Height_cm = hgtCm
    rs!Weight_gram = wtGram
    rs!Ingredients = ingr
    rs!USP_en = uspEn
    rs!USP_vn = uspVn
    rs!HowToUse_en = howEn
    rs!HowToUse_vn = howVn
    rs!OuterPaperBox = outerBox
    rs!ShelfLife_month = shelfMo
    rs!isActive = isActive
    rs!InsertDate = Now()
    rs!ChangeDate = Now()
    rs.Update
End Sub

Private Sub DropTableIfExists(ByVal db As DAO.Database, ByVal tblName As String)
    Dim tdf As DAO.TableDef
    Dim exists As Boolean

    On Error Resume Next
    Set tdf = db.TableDefs(tblName)
    exists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    If Not exists Then Exit Sub

    On Error GoTo ErrHandler
    Dim i As Long
    ' Delete relationships involving this table to avoid foreign key block
    For i = db.Relations.Count - 1 To 0 Step -1
        If StrComp(db.Relations(i).Table, tblName, vbTextCompare) = 0 Or _
           StrComp(db.Relations(i).ForeignTable, tblName, vbTextCompare) = 0 Then
            db.Relations.Delete db.Relations(i).name
        End If
    Next i

    db.Execute "DROP TABLE [" & tblName & "];", dbFailOnError
    Debug.Print "  [DROPPED] " & tblName
    Exit Sub

ErrHandler:
    Debug.Print "  [ERROR] DropTable [" & tblName & "]: #" & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "DropTableIfExists", Err.Description
End Sub
