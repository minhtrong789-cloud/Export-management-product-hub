Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
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
    ItemSuffix =26
    Left =5510
    Top =3030
    Right =18940
    Bottom =9490
    RecSrcDt = Begin
        0x1b81780d7184e640
    End
    Caption ="Advanced Orders Hub"
    DatasheetFontName ="Segoe UI"
    AllowDatasheetView =0
    FilterOnLoad =0
    NavigationCaption ="Orders Explorer"
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =255
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
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
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin ToggleButton
            FontName ="Segoe UI"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =2
            Bevel =1
            BackThemeColorIndex =4
            BackTint =60.0
            OldBorderStyle =0
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
            PressedForeThemeColorIndex =1
        End
        Begin Section
            Height =8640
            Name ="Detail"
            AlternateBackColor =16777215
            AlternateBackShade =95.0
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =10080
                    Top =864
                    Width =1152
                    Height =432
                    FontSize =8
                    FontWeight =700
                    ForeColor =8272943
                    Name ="btnKPIDraft"
                    Caption ="DRAFTS"

                    LayoutCachedLeft =10080
                    LayoutCachedTop =864
                    LayoutCachedWidth =11232
                    LayoutCachedHeight =1296
                    ForeThemeColorIndex =-1
                    Gradient =0
                    BackColor =16777215
                    BackThemeColorIndex =-1
                    BorderWidth =1
                    BorderColor =8272943
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    HoverColor =16181992
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
                    OverlapFlags =85
                    Left =11376
                    Top =864
                    Width =1152
                    Height =432
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    ForeColor =3308846
                    Name ="btnKPIShipped"
                    Caption ="SHIPPED"

                    LayoutCachedLeft =11376
                    LayoutCachedTop =864
                    LayoutCachedWidth =12528
                    LayoutCachedHeight =1296
                    ForeThemeColorIndex =-1
                    Gradient =0
                    BackColor =15332840
                    BackThemeColorIndex =-1
                    BorderWidth =1
                    BorderColor =3308846
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    HoverColor =13231816
                    HoverThemeColorIndex =-1
                    PressedColor =14737632
                    PressedThemeColorIndex =-1
                    HoverForeColor =3308846
                    HoverForeThemeColorIndex =-1
                    PressedForeColor =3308846
                    PressedForeThemeColorIndex =-1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12672
                    Top =864
                    Width =1152
                    Height =432
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    ForeColor =2631878
                    Name ="btnKPIAlert"
                    Caption ="ALERTS"

                    LayoutCachedLeft =12672
                    LayoutCachedTop =864
                    LayoutCachedWidth =13824
                    LayoutCachedHeight =1296
                    ForeThemeColorIndex =-1
                    Gradient =0
                    BackColor =15461371
                    BackThemeColorIndex =-1
                    BorderWidth =1
                    BorderColor =2631878
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    HoverColor =15657983
                    HoverThemeColorIndex =-1
                    PressedColor =14737632
                    PressedThemeColorIndex =-1
                    HoverForeColor =2631878
                    HoverForeThemeColorIndex =-1
                    PressedForeColor =2631878
                    PressedForeThemeColorIndex =-1
                    Overlaps =1
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =288
                    Top =864
                    Width =2448
                    Height =432
                    FontSize =10
                    TabIndex =4
                    BorderColor =14737632
                    ForeColor =8272943
                    Name ="cboCustomer"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="4320;0"
                    Format ="@;\"👥 Select Customer...\""
                    LeftMargin =72
                    TopMargin =72

                    LayoutCachedLeft =288
                    LayoutCachedTop =864
                    LayoutCachedWidth =2736
                    LayoutCachedHeight =1296
                    DisplayAsHyperlink =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2880
                    Top =864
                    Width =1872
                    Height =432
                    FontSize =10
                    TabIndex =5
                    BorderColor =14737632
                    ForeColor =8272943
                    Name ="cboStatus"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="4320;0"
                    Format ="@;\"📋 Select Status...\""
                    LeftMargin =72
                    TopMargin =72

                    LayoutCachedLeft =2880
                    LayoutCachedTop =864
                    LayoutCachedWidth =4752
                    LayoutCachedHeight =1296
                    DisplayAsHyperlink =1
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8784
                    Top =864
                    Width =1152
                    Height =432
                    FontSize =8
                    FontWeight =600
                    TabIndex =8
                    ForeColor =2631878
                    Name ="btnClearFilter"
                    Caption ="CLEAR"

                    LayoutCachedLeft =8784
                    LayoutCachedTop =864
                    LayoutCachedWidth =9936
                    LayoutCachedHeight =1296
                    ForeThemeColorIndex =-1
                    Gradient =0
                    BackColor =16777215
                    BackThemeColorIndex =-1
                    BorderWidth =1
                    BorderColor =2631878
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    HoverColor =15657983
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =14737632
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =2631878
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =2631878
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =13968
                    Top =1440
                    Width =1699
                    Height =432
                    FontSize =9
                    FontWeight =600
                    TabIndex =9
                    ForeColor =2121243
                    Name ="btnExport"
                    Caption ="EXPORT"
                    HorizontalAnchor =1

                    LayoutCachedLeft =13968
                    LayoutCachedTop =1440
                    LayoutCachedWidth =15667
                    LayoutCachedHeight =1872
                    ForeThemeColorIndex =-1
                    Gradient =0
                    BackColor =16777215
                    BackThemeColorIndex =-1
                    BorderWidth =1
                    BorderColor =2121243
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    HoverColor =15332840
                    HoverThemeColorIndex =-1
                    PressedColor =14737632
                    PressedThemeColorIndex =-1
                    HoverForeColor =2121243
                    HoverForeThemeColorIndex =-1
                    PressedForeColor =2121243
                    PressedForeThemeColorIndex =-1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =13968
                    Top =864
                    Width =1699
                    Height =432
                    FontSize =9
                    FontWeight =600
                    TabIndex =10
                    ForeColor =16777215
                    Name ="btnNewOrder"
                    Caption ="+ NEW ORDER"
                    HorizontalAnchor =1

                    LayoutCachedLeft =13968
                    LayoutCachedTop =864
                    LayoutCachedWidth =15667
                    LayoutCachedHeight =1296
                    ForeThemeColorIndex =-1
                    Gradient =0
                    BackColor =8272943
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    OldBorderStyle =0
                    BorderColor =8272943
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    HoverColor =11225401
                    HoverThemeColorIndex =-1
                    PressedColor =5974303
                    PressedThemeColorIndex =-1
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                    Overlaps =1
                End
                Begin Line
                    BorderWidth =1
                    OverlapFlags =85
                    Left =288
                    Top =1440
                    Width =13536
                    BorderColor =15387847
                    Name ="linMainDivider"
                    HorizontalAnchor =2
                    LayoutCachedLeft =288
                    LayoutCachedTop =1440
                    LayoutCachedWidth =13824
                    LayoutCachedHeight =1440
                    BorderThemeColorIndex =-1
                End
                Begin Subform
                    OverlapFlags =85
                    BorderWidth =1
                    Left =288
                    Top =1584
                    Width =13536
                    Height =6228
                    TabIndex =11
                    BorderColor =14737632
                    Name ="subOrderList"
                    HorizontalAnchor =2
                    VerticalAnchor =2

                    LayoutCachedLeft =288
                    LayoutCachedTop =1584
                    LayoutCachedWidth =13824
                    LayoutCachedHeight =7812
                    BorderThemeColorIndex =-1
                End
                Begin Line
                    BorderWidth =1
                    OverlapFlags =85
                    Left =288
                    Top =7884
                    Width =13536
                    BorderColor =15460325
                    Name ="linFooterDivider"
                    HorizontalAnchor =2
                    VerticalAnchor =1
                    LayoutCachedLeft =288
                    LayoutCachedTop =7884
                    LayoutCachedWidth =13824
                    LayoutCachedHeight =7884
                    BorderThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =288
                    Top =8028
                    Width =4320
                    Height =360
                    FontSize =10
                    BorderColor =16777215
                    ForeColor =7697781
                    Name ="lblRecordCount"
                    Caption ="Showing: 0 / 0 orders"
                    VerticalAnchor =1
                    LayoutCachedLeft =288
                    LayoutCachedTop =8028
                    LayoutCachedWidth =4608
                    LayoutCachedHeight =8388
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =7632
                    Top =8028
                    Width =3168
                    Height =360
                    FontSize =10
                    BorderColor =16777215
                    ForeColor =7697781
                    Name ="lblTotalAmount"
                    Caption ="Total Displayed:"
                    HorizontalAnchor =1
                    VerticalAnchor =1
                    LayoutCachedLeft =7632
                    LayoutCachedTop =8028
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =8388
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10944
                    Top =7956
                    Width =2880
                    Height =432
                    FontSize =12
                    FontWeight =700
                    TabIndex =12
                    LeftMargin =72
                    TopMargin =72
                    BorderColor =13421772
                    ForeColor =8272943
                    Name ="txtSumAmount"
                    Format ="@;\"$0.00\""
                    HorizontalAnchor =1
                    VerticalAnchor =1

                    LayoutCachedLeft =10944
                    LayoutCachedTop =7956
                    LayoutCachedWidth =13824
                    LayoutCachedHeight =8388
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Width =15810
                    BackColor =8272943
                    BorderColor =14737632
                    Name ="boxHeaderBg"
                    HorizontalAnchor =2
                    LayoutCachedWidth =15810
                    LayoutCachedHeight =720
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =13968
                    Top =2016
                    Width =1699
                    Height =432
                    FontSize =9
                    FontWeight =600
                    TabIndex =13
                    ForeColor =8272943
                    Name ="btnPrintInvoice"
                    Caption ="PRINT INVOICE"
                    HorizontalAnchor =1

                    LayoutCachedLeft =13968
                    LayoutCachedTop =2016
                    LayoutCachedWidth =15667
                    LayoutCachedHeight =2448
                    ForeThemeColorIndex =-1
                    Gradient =0
                    BackColor =16777215
                    BackThemeColorIndex =-1
                    BorderWidth =1
                    BorderColor =8272943
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    HoverColor =16181992
                    HoverThemeColorIndex =-1
                    PressedColor =14737632
                    PressedThemeColorIndex =-1
                    HoverForeColor =8272943
                    HoverForeThemeColorIndex =-1
                    PressedForeColor =8272943
                    PressedForeThemeColorIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =13968
                    Top =2592
                    Width =1699
                    Height =432
                    FontSize =9
                    FontWeight =600
                    TabIndex =14
                    ForeColor =8272943
                    Name ="btnPackingList"
                    Caption ="PACKING LIST"
                    HorizontalAnchor =1

                    LayoutCachedLeft =13968
                    LayoutCachedTop =2592
                    LayoutCachedWidth =15667
                    LayoutCachedHeight =3024
                    ForeThemeColorIndex =-1
                    Gradient =0
                    BackColor =16777215
                    BackThemeColorIndex =-1
                    BorderWidth =1
                    BorderColor =8272943
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    HoverColor =16181992
                    HoverThemeColorIndex =-1
                    PressedColor =14737632
                    PressedThemeColorIndex =-1
                    HoverForeColor =8272943
                    HoverForeThemeColorIndex =-1
                    PressedForeColor =8272943
                    PressedForeThemeColorIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =13968
                    Top =3168
                    Width =1699
                    Height =432
                    FontSize =9
                    FontWeight =600
                    TabIndex =15
                    ForeColor =7697781
                    Name ="btnRefresh"
                    Caption ="REFRESH"
                    HorizontalAnchor =1

                    LayoutCachedLeft =13968
                    LayoutCachedTop =3168
                    LayoutCachedWidth =15667
                    LayoutCachedHeight =3600
                    ForeThemeColorIndex =-1
                    Gradient =0
                    BackColor =16777215
                    BackThemeColorIndex =-1
                    BorderWidth =1
                    BorderColor =7697781
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    HoverColor =16119285
                    HoverThemeColorIndex =-1
                    PressedColor =14737632
                    PressedThemeColorIndex =-1
                    HoverForeColor =7697781
                    HoverForeThemeColorIndex =-1
                    PressedForeColor =7697781
                    PressedForeThemeColorIndex =-1
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =13968
                    Top =7956
                    Width =1699
                    Height =432
                    FontSize =9
                    FontWeight =600
                    TabIndex =16
                    ForeColor =16777215
                    Name ="togTrashbin"
                    Caption ="TRASH"
                    HorizontalAnchor =1
                    VerticalAnchor =1

                    LayoutCachedLeft =13968
                    LayoutCachedTop =7956
                    LayoutCachedWidth =15667
                    LayoutCachedHeight =8388
                    ForeThemeColorIndex =-1
                    Shape =0
                    Bevel =0
                    BackColor =2631878
                    BackThemeColorIndex =-1
                    BorderColor =2631878
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    HoverColor =3488229
                    HoverThemeColorIndex =-1
                    PressedColor =142
                    PressedThemeColorIndex =-1
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    PressedForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =288
                    Top =144
                    Width =2880
                    Height =432
                    FontSize =16
                    FontWeight =700
                    BorderColor =16777215
                    ForeColor =16777215
                    Name ="lblTitle"
                    Caption ="ORDER HUB"
                    LayoutCachedLeft =288
                    LayoutCachedTop =144
                    LayoutCachedWidth =3168
                    LayoutCachedHeight =576
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6768
                    Top =864
                    Width =1728
                    Height =432
                    FontSize =10
                    TabIndex =7
                    LeftMargin =72
                    TopMargin =72
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtDateTo"
                    Format ="dd/mm/yyyy;;;\"📅 To Date\""

                    LayoutCachedLeft =6768
                    LayoutCachedTop =864
                    LayoutCachedWidth =8496
                    LayoutCachedHeight =1296
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4896
                    Top =864
                    Width =1728
                    Height =432
                    FontSize =10
                    TabIndex =6
                    LeftMargin =72
                    TopMargin =72
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtDateFrom"
                    Format ="dd/mm/yyyy;;;\"📅 From Date\""

                    LayoutCachedLeft =4896
                    LayoutCachedTop =864
                    LayoutCachedWidth =6624
                    LayoutCachedHeight =1296
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3312
                    Top =144
                    Width =3600
                    Height =432
                    TabIndex =3
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =13421772
                    ForeColor =8272943
                    Name ="txtSearch"
                    Format ="@;\"🔍 Search PO, Customer, B/L...\""

                    LayoutCachedLeft =3312
                    LayoutCachedTop =144
                    LayoutCachedWidth =6912
                    LayoutCachedHeight =576
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
    End
End
