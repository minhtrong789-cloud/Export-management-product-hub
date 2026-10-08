Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15810
    DatasheetFontHeight =11
    ItemSuffix =63
    Left =5240
    Top =3040
    Right =13220
    Bottom =9490
    OnUnload ="[Event Procedure]"
    RecSrcDt = Begin
        0xb0943ecaf395e640
    End
    Caption ="Product Detail Viewport"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Segoe UI"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    NavigationCaption ="Product Detail Viewport"
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
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
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
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
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Segoe UI"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
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
        Begin Section
            Height =7920
            Name ="Detail"
            AutoHeight =1
            AlternateBackColor =16777215
            AlternateBackShade =95.0
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    BorderWidth =1
                    OverlapFlags =93
                    Left =288
                    Top =144
                    Width =3816
                    Height =4320
                    BackColor =16447992
                    BorderColor =14737632
                    Name ="boxGeneral"
                    LayoutCachedLeft =288
                    LayoutCachedTop =144
                    LayoutCachedWidth =4104
                    LayoutCachedHeight =4464
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =360
                    Top =216
                    Width =3672
                    Height =317
                    FontWeight =700
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =8272943
                    Name ="lblGeneralTitle"
                    Caption ="GENERAL & PACKAGING"
                    LayoutCachedLeft =360
                    LayoutCachedTop =216
                    LayoutCachedWidth =4032
                    LayoutCachedHeight =533
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =360
                    Top =648
                    Width =1296
                    Height =317
                    FontSize =10
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =3355443
                    Name ="lblMarketingCategory"
                    Caption ="Category"
                    LayoutCachedLeft =360
                    LayoutCachedTop =648
                    LayoutCachedWidth =1656
                    LayoutCachedHeight =965
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin ComboBox
                    BorderWidth =1
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1728
                    Top =648
                    Width =2304
                    Height =317
                    FontSize =10
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="comboMarketingCategory"
                    ControlSource ="MarketingCategory"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MarketingCategory FROM MarketingCategories ORDER BY MarketingCategory; "
                    AfterUpdate ="[Event Procedure]"
                    LeftMargin =72
                    TopMargin =30

                    LayoutCachedLeft =1728
                    LayoutCachedTop =648
                    LayoutCachedWidth =4032
                    LayoutCachedHeight =965
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =360
                    Top =1037
                    Width =1296
                    Height =317
                    FontSize =10
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =3355443
                    Name ="lblMarketingLine"
                    Caption ="Marketing Line"
                    LayoutCachedLeft =360
                    LayoutCachedTop =1037
                    LayoutCachedWidth =1656
                    LayoutCachedHeight =1354
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin ComboBox
                    BorderWidth =1
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1728
                    Top =1037
                    Width =2304
                    Height =317
                    FontSize =10
                    TabIndex =1
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="comboMarketingLine"
                    ControlSource ="MarketingLine"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MarketingLine FROM MarketingLines ORDER BY MarketingLine; "
                    LeftMargin =72
                    TopMargin =30

                    LayoutCachedLeft =1728
                    LayoutCachedTop =1037
                    LayoutCachedWidth =4032
                    LayoutCachedHeight =1354
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =360
                    Top =1426
                    Width =1296
                    Height =317
                    FontSize =10
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =3355443
                    Name ="lblBarcode"
                    Caption ="Barcode"
                    LayoutCachedLeft =360
                    LayoutCachedTop =1426
                    LayoutCachedWidth =1656
                    LayoutCachedHeight =1743
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =215
                    TextAlign =1
                    TextFontFamily =49
                    IMESentenceMode =3
                    Left =1728
                    Top =1426
                    Width =2304
                    Height =317
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtBarcode"
                    ControlSource ="Barcode"
                    Format ="@"
                    FontName ="Courier New"

                    LayoutCachedLeft =1728
                    LayoutCachedTop =1426
                    LayoutCachedWidth =4032
                    LayoutCachedHeight =1743
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =360
                    Top =1814
                    Width =1296
                    Height =317
                    FontSize =10
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =3355443
                    Name ="lblHSCode"
                    Caption ="HS Code"
                    LayoutCachedLeft =360
                    LayoutCachedTop =1814
                    LayoutCachedWidth =1656
                    LayoutCachedHeight =2131
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =215
                    TextAlign =1
                    TextFontFamily =49
                    IMESentenceMode =3
                    Left =1728
                    Top =1814
                    Width =2304
                    Height =317
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtHSCode"
                    ControlSource ="HSCode"
                    Format ="@"
                    FontName ="Courier New"

                    LayoutCachedLeft =1728
                    LayoutCachedTop =1814
                    LayoutCachedWidth =4032
                    LayoutCachedHeight =2131
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =360
                    Top =2203
                    Width =1296
                    Height =317
                    FontSize =10
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =3355443
                    Name ="lblProductName_en"
                    Caption ="Product (EN)"
                    LayoutCachedLeft =360
                    LayoutCachedTop =2203
                    LayoutCachedWidth =1656
                    LayoutCachedHeight =2520
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1728
                    Top =2203
                    Width =2304
                    Height =317
                    FontSize =10
                    TabIndex =4
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtProductName_en"
                    ControlSource ="ProductName_en"
                    Format ="@"

                    LayoutCachedLeft =1728
                    LayoutCachedTop =2203
                    LayoutCachedWidth =4032
                    LayoutCachedHeight =2520
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =360
                    Top =2592
                    Width =1296
                    Height =317
                    FontSize =10
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =3355443
                    Name ="lblProductName_vn"
                    Caption ="Product (VN)"
                    LayoutCachedLeft =360
                    LayoutCachedTop =2592
                    LayoutCachedWidth =1656
                    LayoutCachedHeight =2909
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1728
                    Top =2592
                    Width =2304
                    Height =317
                    FontSize =10
                    TabIndex =5
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtProductName_vn"
                    ControlSource ="ProductName_vn"
                    Format ="@"

                    LayoutCachedLeft =1728
                    LayoutCachedTop =2592
                    LayoutCachedWidth =4032
                    LayoutCachedHeight =2909
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =360
                    Top =2981
                    Width =1296
                    Height =317
                    FontSize =10
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =3355443
                    Name ="lblPackageType"
                    Caption ="Packaging"
                    LayoutCachedLeft =360
                    LayoutCachedTop =2981
                    LayoutCachedWidth =1656
                    LayoutCachedHeight =3298
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin ComboBox
                    BorderWidth =1
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1728
                    Top =2981
                    Width =2304
                    Height =317
                    FontSize =10
                    TabIndex =6
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="cboPackageType"
                    ControlSource ="PackageType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT PackageTypes.PackageTypeID, PackageTypes.PackageType FROM PackageTypes; "
                    LeftMargin =72
                    TopMargin =30

                    LayoutCachedLeft =1728
                    LayoutCachedTop =2981
                    LayoutCachedWidth =4032
                    LayoutCachedHeight =3298
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =360
                    Top =3370
                    Width =1296
                    Height =317
                    FontSize =10
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =3355443
                    Name ="lblShelfLife_month"
                    Caption ="Shelf Life"
                    LayoutCachedLeft =360
                    LayoutCachedTop =3370
                    LayoutCachedWidth =1656
                    LayoutCachedHeight =3687
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1728
                    Top =3370
                    Width =2304
                    Height =317
                    FontSize =10
                    TabIndex =7
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16775672
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtShelfLife_month"
                    ControlSource ="ShelfLife_month"
                    Format ="#,##0\" months\""

                    LayoutCachedLeft =1728
                    LayoutCachedTop =3370
                    LayoutCachedWidth =4032
                    LayoutCachedHeight =3687
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =360
                    Top =3758
                    Width =1296
                    Height =317
                    FontSize =10
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =3355443
                    Name ="lblOuterPaperBox"
                    Caption ="Outer Box"
                    LayoutCachedLeft =360
                    LayoutCachedTop =3758
                    LayoutCachedWidth =1656
                    LayoutCachedHeight =4075
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin CheckBox
                    OverlapFlags =215
                    OldBorderStyle =0
                    Left =1728
                    Top =3758
                    Width =432
                    Height =317
                    TabIndex =8
                    BorderColor =8272943
                    Name ="blnOuterPaperBox"
                    ControlSource ="OuterPaperBox"

                    LayoutCachedLeft =1728
                    LayoutCachedTop =3758
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =4075
                    BorderThemeColorIndex =-1
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    BorderWidth =1
                    OverlapFlags =93
                    Left =4248
                    Top =144
                    Width =3816
                    Height =4320
                    BackColor =16447992
                    BorderColor =14737632
                    Name ="boxDimensions"
                    LayoutCachedLeft =4248
                    LayoutCachedTop =144
                    LayoutCachedWidth =8064
                    LayoutCachedHeight =4464
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =4320
                    Top =216
                    Width =3672
                    Height =317
                    FontWeight =700
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =8272943
                    Name ="lblDimensionsTitle"
                    Caption ="PRICING & LOGISTICS"
                    LayoutCachedLeft =4320
                    LayoutCachedTop =216
                    LayoutCachedWidth =7992
                    LayoutCachedHeight =533
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =4320
                    Top =648
                    Width =1296
                    Height =317
                    FontSize =10
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =3355443
                    Name ="lblCapacity"
                    Caption ="Volume"
                    LayoutCachedLeft =4320
                    LayoutCachedTop =648
                    LayoutCachedWidth =5616
                    LayoutCachedHeight =965
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5688
                    Top =648
                    Width =1368
                    Height =317
                    FontSize =10
                    TabIndex =9
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16775672
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtCapacity"
                    ControlSource ="Capacity"
                    Format ="Standard"

                    LayoutCachedLeft =5688
                    LayoutCachedTop =648
                    LayoutCachedWidth =7056
                    LayoutCachedHeight =965
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin ComboBox
                    BorderWidth =1
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7128
                    Top =648
                    Width =864
                    Height =317
                    FontSize =10
                    TabIndex =10
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="cboCapacityUnit"
                    ControlSource ="CapacityUnit"
                    RowSourceType ="Table/Query"
                    LeftMargin =72
                    TopMargin =30

                    LayoutCachedLeft =7128
                    LayoutCachedTop =648
                    LayoutCachedWidth =7992
                    LayoutCachedHeight =965
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =4320
                    Top =1037
                    Width =1296
                    Height =317
                    FontSize =10
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =3355443
                    Name ="lblCapacity_oz"
                    Caption ="Volume (Oz)"
                    LayoutCachedLeft =4320
                    LayoutCachedTop =1037
                    LayoutCachedWidth =5616
                    LayoutCachedHeight =1354
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5688
                    Top =1037
                    Width =2304
                    Height =317
                    FontSize =10
                    TabIndex =11
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16775672
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtCapacity_oz"
                    ControlSource ="Capacity_oz"
                    Format ="#,##0.00\" oz\""

                    LayoutCachedLeft =5688
                    LayoutCachedTop =1037
                    LayoutCachedWidth =7992
                    LayoutCachedHeight =1354
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =4320
                    Top =1426
                    Width =1296
                    Height =317
                    FontSize =10
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =3355443
                    Name ="lblPrice_vnd"
                    Caption ="Price (VND)"
                    LayoutCachedLeft =4320
                    LayoutCachedTop =1426
                    LayoutCachedWidth =5616
                    LayoutCachedHeight =1743
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5688
                    Top =1426
                    Width =2304
                    Height =317
                    FontSize =10
                    TabIndex =12
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtPrice_vnd"
                    ControlSource ="Price_vnd"
                    Format ="#,##0\" VND\""

                    LayoutCachedLeft =5688
                    LayoutCachedTop =1426
                    LayoutCachedWidth =7992
                    LayoutCachedHeight =1743
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =4320
                    Top =1814
                    Width =1296
                    Height =317
                    FontSize =10
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =3355443
                    Name ="lblPrice_usd"
                    Caption ="Price (USD)"
                    LayoutCachedLeft =4320
                    LayoutCachedTop =1814
                    LayoutCachedWidth =5616
                    LayoutCachedHeight =2131
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5688
                    Top =1814
                    Width =2304
                    Height =317
                    FontSize =10
                    TabIndex =13
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtPrice_usd"
                    ControlSource ="Price_usd"
                    Format ="$#,##0.00;($#,##0.00);\"$0.00\""

                    LayoutCachedLeft =5688
                    LayoutCachedTop =1814
                    LayoutCachedWidth =7992
                    LayoutCachedHeight =2131
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =4320
                    Top =2203
                    Width =1296
                    Height =317
                    FontSize =10
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =3355443
                    Name ="lblLength_cm"
                    Caption ="Length (cm)"
                    LayoutCachedLeft =4320
                    LayoutCachedTop =2203
                    LayoutCachedWidth =5616
                    LayoutCachedHeight =2520
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5688
                    Top =2203
                    Width =2304
                    Height =317
                    FontSize =10
                    TabIndex =14
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16775672
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtLength_cm"
                    ControlSource ="Length_cm"
                    Format ="#,##0.0\" cm\""

                    LayoutCachedLeft =5688
                    LayoutCachedTop =2203
                    LayoutCachedWidth =7992
                    LayoutCachedHeight =2520
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =4320
                    Top =2592
                    Width =1296
                    Height =317
                    FontSize =10
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =3355443
                    Name ="lblWidth_cm"
                    Caption ="Width (cm)"
                    LayoutCachedLeft =4320
                    LayoutCachedTop =2592
                    LayoutCachedWidth =5616
                    LayoutCachedHeight =2909
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5688
                    Top =2592
                    Width =2304
                    Height =317
                    FontSize =10
                    TabIndex =15
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16775672
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtWidth_cm"
                    ControlSource ="Width_cm"
                    Format ="#,##0.0\" cm\""

                    LayoutCachedLeft =5688
                    LayoutCachedTop =2592
                    LayoutCachedWidth =7992
                    LayoutCachedHeight =2909
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =4320
                    Top =2981
                    Width =1296
                    Height =317
                    FontSize =10
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =3355443
                    Name ="lblHeight_cm"
                    Caption ="Height (cm)"
                    LayoutCachedLeft =4320
                    LayoutCachedTop =2981
                    LayoutCachedWidth =5616
                    LayoutCachedHeight =3298
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5688
                    Top =2981
                    Width =2304
                    Height =317
                    FontSize =10
                    TabIndex =16
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16775672
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtHeight_cm"
                    ControlSource ="Height_cm"
                    Format ="#,##0.0\" cm\""

                    LayoutCachedLeft =5688
                    LayoutCachedTop =2981
                    LayoutCachedWidth =7992
                    LayoutCachedHeight =3298
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =4320
                    Top =3370
                    Width =1296
                    Height =317
                    FontSize =10
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =3355443
                    Name ="lblWeight_gram"
                    Caption ="Weight (g)"
                    LayoutCachedLeft =4320
                    LayoutCachedTop =3370
                    LayoutCachedWidth =5616
                    LayoutCachedHeight =3687
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5688
                    Top =3370
                    Width =2304
                    Height =317
                    FontSize =10
                    TabIndex =17
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16775672
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtWeight_gram"
                    ControlSource ="Weight_gram"
                    Format ="#,##0\" g\""

                    LayoutCachedLeft =5688
                    LayoutCachedTop =3370
                    LayoutCachedWidth =7992
                    LayoutCachedHeight =3687
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    BorderWidth =1
                    OverlapFlags =93
                    Left =288
                    Top =4579
                    Width =7776
                    Height =2520
                    BackColor =16447992
                    BorderColor =14737632
                    Name ="boxDocumentation"
                    VerticalAnchor =2
                    LayoutCachedLeft =288
                    LayoutCachedTop =4579
                    LayoutCachedWidth =8064
                    LayoutCachedHeight =7099
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =360
                    Top =4651
                    Width =3744
                    Height =288
                    FontSize =10
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =3355443
                    Name ="lblUSP"
                    Caption ="USP / Highlights"
                    LayoutCachedLeft =360
                    LayoutCachedTop =4651
                    LayoutCachedWidth =4104
                    LayoutCachedHeight =4939
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    ScrollBars =2
                    BorderWidth =1
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =360
                    Top =4968
                    Width =3744
                    Height =2016
                    FontSize =10
                    TabIndex =18
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16447992
                    BorderColor =14737632
                    ForeColor =3355443
                    Name ="txtUSP"
                    ControlSource ="USP_en"
                    Format ="@;\"[Enter USP / Key Highlights]\""
                    VerticalAnchor =2

                    LayoutCachedLeft =360
                    LayoutCachedTop =4968
                    LayoutCachedWidth =4104
                    LayoutCachedHeight =6984
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =4248
                    Top =4651
                    Width =3744
                    Height =288
                    FontSize =10
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =3355443
                    Name ="lblHowToUse"
                    Caption ="How To Use"
                    LayoutCachedLeft =4248
                    LayoutCachedTop =4651
                    LayoutCachedWidth =7992
                    LayoutCachedHeight =4939
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    ScrollBars =2
                    BorderWidth =1
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4248
                    Top =4968
                    Width =3744
                    Height =2016
                    FontSize =10
                    TabIndex =19
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16447992
                    BorderColor =14737632
                    ForeColor =3355443
                    Name ="txtHowToUse"
                    ControlSource ="HowToUse_en"
                    Format ="@;\"[Enter Usage Instructions]\""
                    VerticalAnchor =2

                    LayoutCachedLeft =4248
                    LayoutCachedTop =4968
                    LayoutCachedWidth =7992
                    LayoutCachedHeight =6984
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    BorderWidth =1
                    OverlapFlags =93
                    Left =8208
                    Top =144
                    Width =5400
                    Height =3384
                    BackColor =16447992
                    BorderColor =14737632
                    Name ="boxImage"
                    HorizontalAnchor =2
                    LayoutCachedLeft =8208
                    LayoutCachedTop =144
                    LayoutCachedWidth =13608
                    LayoutCachedHeight =3528
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =8280
                    Top =216
                    Width =5256
                    Height =288
                    FontWeight =700
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =8272943
                    Name ="lblImage"
                    Caption ="PRODUCT IMAGE"
                    HorizontalAnchor =2
                    LayoutCachedLeft =8280
                    LayoutCachedTop =216
                    LayoutCachedWidth =13536
                    LayoutCachedHeight =504
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Image
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =1
                    Left =8280
                    Top =576
                    Width =5256
                    Height =2304
                    BorderColor =15387847
                    Name ="imgProduct"
                    HorizontalAnchor =1
                    VerticalAnchor =2

                    LayoutCachedLeft =8280
                    LayoutCachedTop =576
                    LayoutCachedWidth =13536
                    LayoutCachedHeight =2880
                    TabIndex =26
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    Left =8280
                    Top =1296
                    Width =5256
                    Height =720
                    FontSize =14
                    FontWeight =700
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =13421772
                    Name ="LabelNoImage"
                    Caption ="NO IMAGE"
                    HorizontalAnchor =2
                    LayoutCachedLeft =8280
                    LayoutCachedTop =1296
                    LayoutCachedWidth =13536
                    LayoutCachedHeight =2016
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =8280
                    Top =2952
                    Width =2520
                    Height =461
                    FontSize =9
                    FontWeight =600
                    TabIndex =20
                    ForeColor =8272943
                    Name ="cmdAddImage"
                    Caption ="ADD IMAGE"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Select and attach image for active product"

                    LayoutCachedLeft =8280
                    LayoutCachedTop =2952
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =3413
                    ForeThemeColorIndex =-1
                    Shape =0
                    Gradient =0
                    BackColor =16777215
                    BackThemeColorIndex =-1
                    BorderWidth =1
                    BorderColor =8272943
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    HoverColor =16447992
                    HoverThemeColorIndex =-1
                    PressedColor =14737632
                    PressedThemeColorIndex =-1
                    HoverForeColor =8272943
                    HoverForeThemeColorIndex =-1
                    PressedForeColor =8272943
                    PressedForeThemeColorIndex =-1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =11016
                    Top =2952
                    Width =2520
                    Height =461
                    FontSize =10
                    FontWeight =600
                    TabIndex =21
                    ForeColor =16777215
                    Name ="cmdClearImage"
                    Caption ="CLEAR"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Remove attached image from active product"

                    LayoutCachedLeft =11016
                    LayoutCachedTop =2952
                    LayoutCachedWidth =13536
                    LayoutCachedHeight =3413
                    ForeThemeColorIndex =-1
                    Shape =0
                    Gradient =0
                    BackColor =3488229
                    BackThemeColorIndex =-1
                    OldBorderStyle =0
                    BorderColor =3488229
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    HoverColor =5264367
                    HoverThemeColorIndex =-1
                    PressedColor =2631878
                    PressedThemeColorIndex =-1
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    Overlaps =1
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    BorderWidth =1
                    OverlapFlags =93
                    Left =8208
                    Top =3629
                    Width =5400
                    Height =3470
                    BackColor =16447992
                    BorderColor =14737632
                    Name ="boxIngredients"
                    HorizontalAnchor =2
                    VerticalAnchor =2
                    LayoutCachedLeft =8208
                    LayoutCachedTop =3629
                    LayoutCachedWidth =13608
                    LayoutCachedHeight =7099
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =8280
                    Top =3701
                    Width =5256
                    Height =288
                    FontWeight =700
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =8272943
                    Name ="lblIngredients"
                    Caption ="INGREDIENTS"
                    HorizontalAnchor =2
                    LayoutCachedLeft =8280
                    LayoutCachedTop =3701
                    LayoutCachedWidth =13536
                    LayoutCachedHeight =3989
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    ScrollBars =2
                    BorderWidth =1
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8280
                    Top =4061
                    Width =5256
                    Height =2923
                    FontSize =10
                    TabIndex =22
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16447992
                    BorderColor =14737632
                    ForeColor =3355443
                    Name ="txtIngredients"
                    ControlSource ="Ingredients"
                    HorizontalAnchor =2
                    VerticalAnchor =2

                    LayoutCachedLeft =8280
                    LayoutCachedTop =4061
                    LayoutCachedWidth =13536
                    LayoutCachedHeight =6984
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =13824
                    Top =144
                    Width =1699
                    Height =576
                    FontSize =9
                    FontWeight =700
                    TabIndex =23
                    ForeColor =16777215
                    Name ="cmdNewProduct"
                    Caption ="+ NEW PRODUCT"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Create a new product record"

                    LayoutCachedLeft =13824
                    LayoutCachedTop =144
                    LayoutCachedWidth =15523
                    LayoutCachedHeight =720
                    ForeThemeColorIndex =-1
                    Shape =0
                    Gradient =0
                    BackColor =6775289
                    BackThemeColorIndex =-1
                    OldBorderStyle =0
                    BorderColor =6775289
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    HoverColor =8354554
                    HoverThemeColorIndex =-1
                    PressedColor =5327840
                    PressedThemeColorIndex =-1
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                End
                Begin Line
                    BorderWidth =1
                    OverlapFlags =85
                    Left =288
                    Top =7200
                    Width =13320
                    Height =14
                    BorderColor =15460325
                    Name ="linFooterSeparator"
                    HorizontalAnchor =2
                    VerticalAnchor =1
                    LayoutCachedLeft =288
                    LayoutCachedTop =7200
                    LayoutCachedWidth =13608
                    LayoutCachedHeight =7214
                    BorderThemeColorIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =288
                    Top =7344
                    Height =432
                    FontSize =10
                    FontWeight =600
                    TabIndex =24
                    ForeColor =16777215
                    Name ="cmdPrev"
                    Caption ="< PREVIOUS"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Go to previous product record"

                    LayoutCachedLeft =288
                    LayoutCachedTop =7344
                    LayoutCachedWidth =1728
                    LayoutCachedHeight =7776
                    ForeThemeColorIndex =-1
                    Shape =0
                    Gradient =0
                    BackColor =6576709
                    BackThemeColorIndex =-1
                    OldBorderStyle =0
                    BorderColor =6576709
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    HoverColor =9141600
                    HoverThemeColorIndex =-1
                    PressedColor =3682854
                    PressedThemeColorIndex =-1
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1872
                    Top =7344
                    Width =2304
                    Height =432
                    FontSize =9
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =8947848
                    Name ="lblNavStatus"
                    Caption ="Product 0 of 0"
                    VerticalAnchor =1
                    LayoutCachedLeft =1872
                    LayoutCachedTop =7344
                    LayoutCachedWidth =4176
                    LayoutCachedHeight =7776
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4320
                    Top =7344
                    Height =432
                    FontSize =10
                    FontWeight =600
                    TabIndex =25
                    ForeColor =16777215
                    Name ="cmdNext"
                    Caption ="NEXT >"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Go to next product record"

                    LayoutCachedLeft =4320
                    LayoutCachedTop =7344
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =7776
                    ForeThemeColorIndex =-1
                    Shape =0
                    Gradient =0
                    BackColor =6576709
                    BackThemeColorIndex =-1
                    OldBorderStyle =0
                    BorderColor =6576709
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    HoverColor =9141600
                    HoverThemeColorIndex =-1
                    PressedColor =3682854
                    PressedThemeColorIndex =-1
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =8208
                    Top =7344
                    Width =5400
                    Height =432
                    FontSize =10
                    FontWeight =700
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =3355443
                    Name ="lblRecordSummary"
                    Caption ="Total CBM: 0.00 | Total Value: $0.00"
                    HorizontalAnchor =1
                    VerticalAnchor =1
                    LayoutCachedLeft =8208
                    LayoutCachedTop =7344
                    LayoutCachedWidth =13608
                    LayoutCachedHeight =7776
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
            End
        End
    End
End
CodeBehindForm
' See "frmProductDetailView.cls"
