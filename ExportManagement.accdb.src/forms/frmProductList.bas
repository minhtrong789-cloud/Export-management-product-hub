Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =2
    ScrollBars =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15810
    RowHeight =350
    DatasheetFontHeight =10
    ItemSuffix =212
    Left =3060
    Top =2200
    Right =16460
    Bottom =8470
    OnUnload ="[Event Procedure]"
    BeforeDelConfirm ="[Event Procedure]"
    AfterDelConfirm ="[Event Procedure]"
    RecSrcDt = Begin
        0xa881e5735788e640
    End
    OnCurrent ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Segoe UI"
    FilterOnLoad =0
    OrderByOnLoad =0
    SplitFormOrientation =1
    OrderByOnLoad =0
    SplitFormOrientation =1
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetGridlinesColor12 =0
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Segoe UI"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Segoe UI"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ListBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Segoe UI"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Segoe UI"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Visible = NotDefault
            Height =0
            Name ="FormHeader"
            AutoHeight =255
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackTint =20.0
        End
        Begin PageHeader
            DisplayWhen =1
            Height =0
            Name ="PageHeaderSection"
            AutoHeight =255
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin Section
            CanGrow = NotDefault
            Height =8640
            Name ="Detail"
            AlternateBackColor =16777215
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =93
                    TextAlign =1
                    TextFontFamily =49
                    IMESentenceMode =3
                    Top =1470
                    Width =1152
                    Height =360
                    ColumnWidth =2220
                    ColumnOrder =4
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtBarcode"
                    ControlSource ="Barcode"
                    Format ="@"
                    FontName ="Courier New"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =1470
                    LayoutCachedWidth =1152
                    LayoutCachedHeight =1830
                    DatasheetCaption ="Barcode"
                    RowStart =4
                    RowEnd =4
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =1
                    TextFontFamily =49
                    IMESentenceMode =3
                    Top =1840
                    Width =1152
                    Height =360
                    ColumnWidth =2220
                    ColumnOrder =5
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtHSCode"
                    ControlSource ="HSCode"
                    Format ="@"
                    FontName ="Courier New"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =1840
                    LayoutCachedWidth =1152
                    LayoutCachedHeight =2200
                    DatasheetCaption ="HS Code"
                    RowStart =5
                    RowEnd =5
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    BorderWidth =1
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Top =2200
                    Width =1152
                    Height =360
                    ColumnWidth =2220
                    ColumnOrder =6
                    FontSize =12
                    FontWeight =700
                    TabIndex =6
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =13421772
                    ForeColor =8272943
                    Name ="txtProductName"
                    ControlSource ="ProductName_en"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =2200
                    LayoutCachedWidth =1152
                    LayoutCachedHeight =2560
                    DatasheetCaption ="Product Name"
                    RowStart =6
                    RowEnd =6
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Top =2940
                    Width =1152
                    Height =360
                    ColumnWidth =2220
                    ColumnOrder =8
                    FontSize =10
                    TabIndex =8
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16775672
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtCapacity"
                    ControlSource ="Capacity"
                    Format ="Standard"
                    ValidationRule =">0"
                    ValidationText ="\"Caannot do this"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =2940
                    LayoutCachedWidth =1152
                    LayoutCachedHeight =3300
                    DatasheetCaption ="Volume"
                    RowStart =8
                    RowEnd =8
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Top =3680
                    Width =1152
                    Height =360
                    ColumnWidth =2220
                    ColumnOrder =10
                    FontSize =10
                    TabIndex =10
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16775672
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtCapacity_oz"
                    ControlSource ="Capacity_oz"
                    Format ="#,##0.00\" oz\""
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =3680
                    LayoutCachedWidth =1152
                    LayoutCachedHeight =4040
                    DatasheetCaption ="Oz"
                    RowStart =10
                    RowEnd =10
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =3
                    IMESentenceMode =3
                    Top =4050
                    Width =1152
                    Height =360
                    ColumnWidth =2220
                    ColumnOrder =11
                    FontSize =10
                    TabIndex =11
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtPrice"
                    ControlSource ="Price_usd"
                    Format ="#,##0\" VND\""
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =4050
                    LayoutCachedWidth =1152
                    LayoutCachedHeight =4410
                    DatasheetCaption ="Price"
                    RowStart =11
                    RowEnd =11
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Top =4420
                    Width =1152
                    Height =360
                    ColumnWidth =2220
                    ColumnOrder =12
                    FontSize =10
                    TabIndex =12
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16775672
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtLength_cm"
                    ControlSource ="Length_cm"
                    Format ="#,##0.0\" cm\""
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =4420
                    LayoutCachedWidth =1152
                    LayoutCachedHeight =4780
                    DatasheetCaption ="Length"
                    RowStart =12
                    RowEnd =12
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Top =4790
                    Width =1152
                    Height =360
                    ColumnWidth =2220
                    ColumnOrder =13
                    FontSize =10
                    TabIndex =13
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16775672
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtWidth_cm"
                    ControlSource ="Width_cm"
                    Format ="#,##0.0\" cm\""
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =4790
                    LayoutCachedWidth =1152
                    LayoutCachedHeight =5150
                    DatasheetCaption ="Width"
                    RowStart =13
                    RowEnd =13
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Top =5160
                    Width =1152
                    Height =360
                    ColumnWidth =2220
                    ColumnOrder =14
                    FontSize =10
                    TabIndex =14
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16775672
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtHeight_cm"
                    ControlSource ="Height_cm"
                    Format ="#,##0.0\" cm\""
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =5160
                    LayoutCachedWidth =1152
                    LayoutCachedHeight =5520
                    DatasheetCaption ="Height"
                    RowStart =14
                    RowEnd =14
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Top =5530
                    Width =1152
                    Height =360
                    ColumnWidth =2220
                    ColumnOrder =15
                    FontSize =10
                    TabIndex =15
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16775672
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtWeight_gram"
                    ControlSource ="Weight_gram"
                    Format ="#,##0\" g\""
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =5530
                    LayoutCachedWidth =1152
                    LayoutCachedHeight =5890
                    DatasheetCaption ="Weight"
                    RowStart =15
                    RowEnd =15
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Top =5900
                    Width =1152
                    Height =360
                    ColumnWidth =2220
                    ColumnOrder =16
                    FontSize =10
                    TabIndex =16
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16447992
                    BorderColor =14737632
                    ForeColor =3355443
                    Name ="txtIngredients"
                    ControlSource ="Ingredients"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =5900
                    LayoutCachedWidth =1152
                    LayoutCachedHeight =6260
                    DatasheetCaption ="Ingredients"
                    RowStart =16
                    RowEnd =16
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    ScrollBars =2
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Top =6270
                    Width =1152
                    Height =360
                    ColumnWidth =2560
                    ColumnOrder =17
                    FontSize =10
                    TabIndex =17
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16447992
                    BorderColor =14737632
                    ForeColor =3355443
                    Name ="txtUSP"
                    ControlSource ="USP_en"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =6270
                    LayoutCachedWidth =1152
                    LayoutCachedHeight =6630
                    DatasheetCaption ="USP"
                    RowStart =17
                    RowEnd =17
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    ScrollBars =2
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Top =6640
                    Width =1152
                    Height =360
                    ColumnWidth =3290
                    ColumnOrder =18
                    FontSize =10
                    TabIndex =18
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16447992
                    BorderColor =14737632
                    ForeColor =3355443
                    Name ="txtHowToUse"
                    ControlSource ="HowToUse_en"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =6640
                    LayoutCachedWidth =1152
                    LayoutCachedHeight =7000
                    DatasheetCaption ="How to Use"
                    RowStart =18
                    RowEnd =18
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    GroupTable =1
                End
                Begin CheckBox
                    OverlapFlags =215
                    OldBorderStyle =0
                    Top =7000
                    Width =1152
                    Height =288
                    ColumnWidth =2220
                    ColumnOrder =19
                    TabIndex =19
                    BorderColor =8272943
                    Name ="blnOuterPaperBox"
                    ControlSource ="OuterPaperBox"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =7000
                    LayoutCachedWidth =1152
                    LayoutCachedHeight =7288
                    RowStart =19
                    RowEnd =19
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    GroupTable =1
                    DatasheetCaption ="Have Outer Box?"
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Top =7300
                    Width =1152
                    Height =360
                    ColumnWidth =2220
                    ColumnOrder =20
                    FontSize =10
                    TabIndex =20
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16775672
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtShelfLife_month"
                    ControlSource ="ShelfLife_month"
                    Format ="#,##0\" months\""
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =7300
                    LayoutCachedWidth =1152
                    LayoutCachedHeight =7660
                    DatasheetCaption ="Shelf Life"
                    RowStart =20
                    RowEnd =20
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Width =1152
                    Height =360
                    ColumnWidth =2700
                    ColumnOrder =0
                    FontSize =10
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtRowNum"
                    ControlSource ="=GetRowNumber([Form],\"ProductID\",[ProductID])"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedWidth =1152
                    LayoutCachedHeight =360
                    DatasheetCaption ="No."
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    GroupTable =1
                End
                Begin ComboBox
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Top =1100
                    Width =1152
                    Height =360
                    ColumnWidth =2610
                    ColumnOrder =3
                    FontSize =10
                    TabIndex =3
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="comboMarketingLine"
                    ControlSource ="MarketingLine"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MarketingLine FROM MarketingLines ORDER BY MarketingLine; "
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LeftMargin =72
                    TopMargin =30

                    LayoutCachedTop =1100
                    LayoutCachedWidth =1152
                    LayoutCachedHeight =1460
                    DatasheetCaption ="Marketing Line"
                    RowStart =3
                    RowEnd =3
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    GroupTable =1
                End
                Begin ComboBox
                    BorderWidth =1
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Top =2570
                    Width =1152
                    Height =360
                    ColumnWidth =2220
                    ColumnOrder =7
                    FontSize =10
                    TabIndex =7
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="cboPackageType"
                    ControlSource ="PackageType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT PackageTypes.PackageTypeID, PackageTypes.PackageType FROM PackageTypes; "
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LeftMargin =72
                    TopMargin =30

                    LayoutCachedTop =2570
                    LayoutCachedWidth =1152
                    LayoutCachedHeight =2930
                    DatasheetCaption ="Packaging"
                    RowStart =7
                    RowEnd =7
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    GroupTable =1
                End
                Begin CheckBox
                    OverlapFlags =215
                    OldBorderStyle =0
                    Top =360
                    Width =1152
                    Height =360
                    ColumnOrder =1
                    TabIndex =1
                    BorderColor =8272943
                    Name ="chkContextAction"
                    ControlSource ="chkContextAction"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="False"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0

                    LayoutCachedTop =360
                    LayoutCachedWidth =1152
                    LayoutCachedHeight =720
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    BorderThemeColorIndex =-1
                    GroupTable =1
                    DatasheetCaption ="Selected"
                End
                Begin ComboBox
                    BorderWidth =1
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Top =3310
                    Width =1152
                    Height =360
                    ColumnOrder =9
                    FontSize =10
                    TabIndex =9
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="cboCapacityUnit"
                    ControlSource ="CapacityUnit"
                    RowSourceType ="Table/Query"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LeftMargin =72
                    TopMargin =30

                    LayoutCachedTop =3310
                    LayoutCachedWidth =1152
                    LayoutCachedHeight =3670
                    DatasheetCaption ="Unit"
                    RowStart =9
                    RowEnd =9
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    GroupTable =1
                End
                Begin ComboBox
                    BorderWidth =1
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Top =730
                    Width =1152
                    Height =360
                    ColumnOrder =2
                    FontSize =10
                    TabIndex =2
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="comboCategory"
                    ControlSource ="MarketingCategory"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MarketingCategory FROM MarketingCategories ORDER BY MarketingCategory; "
                    OnLostFocus ="[Event Procedure]"
                    GroupTable =1
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    LeftMargin =72
                    TopMargin =30

                    LayoutCachedTop =730
                    LayoutCachedWidth =1152
                    LayoutCachedHeight =1090
                    DatasheetCaption ="Category"
                    RowStart =2
                    RowEnd =2
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    GroupTable =1
                End
            End
        End
        Begin PageFooter
            DisplayWhen =1
            Height =0
            Name ="PageFooterSection"
            AutoHeight =255
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin FormFooter
            Visible = NotDefault
            Height =0
            Name ="FormFooter"
            AutoHeight =255
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
    End
End
CodeBehindForm
' See "frmProductList.cls"
