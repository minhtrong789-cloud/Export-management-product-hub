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
    ItemSuffix =16
    Left =-4650
    Top =4360
    Right =10900
    Bottom =12020
    OnUnload ="[Event Procedure]"
    RecSrcDt = Begin
        0x0af134a95e95e640
    End
    Caption ="Product Grid Viewport"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Segoe UI"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    NavigationCaption ="Product Grid Viewport"
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
        Begin Section
            Height =7920
            Name ="Detail"
            AutoHeight =1
            AlternateBackColor =16777215
            AlternateBackShade =95.0
            Begin
                Begin ComboBox
                    BorderWidth =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =288
                    Top =144
                    Width =2592
                    Height =432
                    FontSize =10
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="cboFilterCategory"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Filter by Product Category"
                    Format ="@;\"[ All Categories ]\""
                    LeftMargin =72
                    TopMargin =30

                    LayoutCachedLeft =288
                    LayoutCachedTop =144
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =576
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin ComboBox
                    BorderWidth =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3024
                    Top =144
                    Width =2592
                    Height =432
                    FontSize =10
                    TabIndex =1
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="cboFilterLine"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Filter by Marketing Line"
                    Format ="@;\"[ All Marketing Lines ]\""
                    LeftMargin =72
                    TopMargin =30

                    LayoutCachedLeft =3024
                    LayoutCachedTop =144
                    LayoutCachedWidth =5616
                    LayoutCachedHeight =576
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Line
                    BorderWidth =1
                    OverlapFlags =85
                    Left =288
                    Top =691
                    Width =13392
                    Height =14
                    BorderColor =15460325
                    Name ="linMainDivider"
                    HorizontalAnchor =2
                    LayoutCachedLeft =288
                    LayoutCachedTop =691
                    LayoutCachedWidth =13680
                    LayoutCachedHeight =705
                    BorderThemeColorIndex =-1
                End
                Begin Subform
                    OverlapFlags =247
                    BorderWidth =1
                    Left =288
                    Top =792
                    Width =13392
                    Height =6264
                    TabIndex =2
                    BorderColor =15387847
                    Name ="sfProductGrid"
                    SourceObject ="Form.frmProductList"
                    HorizontalAnchor =2
                    VerticalAnchor =2

                    LayoutCachedLeft =288
                    LayoutCachedTop =792
                    LayoutCachedWidth =13680
                    LayoutCachedHeight =7056
                    BorderThemeColorIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =13824
                    Top =144
                    Width =1699
                    Height =576
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    ForeColor =16777215
                    Name ="cmdNewProduct"
                    Caption ="+ NEW PRODUCT"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Create a new product record"
                    HorizontalAnchor =1

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
                Begin CommandButton
                    OverlapFlags =85
                    Left =13824
                    Top =864
                    Width =1699
                    Height =432
                    FontSize =10
                    FontWeight =600
                    TabIndex =4
                    ForeColor =16777215
                    Name ="cmdExport"
                    Caption ="EXPORT EXCEL"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Export active products to Excel"
                    HorizontalAnchor =1

                    LayoutCachedLeft =13824
                    LayoutCachedTop =864
                    LayoutCachedWidth =15523
                    LayoutCachedHeight =1296
                    ForeThemeColorIndex =-1
                    Shape =0
                    Gradient =0
                    BackColor =2121243
                    BackThemeColorIndex =-1
                    OldBorderStyle =0
                    BorderColor =2121243
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    HoverColor =3968568
                    HoverThemeColorIndex =-1
                    PressedColor =729354
                    PressedThemeColorIndex =-1
                    HoverForeColor =16777215
                    HoverForeThemeColorIndex =-1
                    PressedForeColor =16777215
                    PressedForeThemeColorIndex =-1
                End
                Begin Image
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =1
                    Left =7632
                    Top =792
                    Width =6048
                    Height =5040
                    BorderColor =15387847
                    Name ="imgProduct"
                    HorizontalAnchor =1
                    VerticalAnchor =2

                    LayoutCachedLeft =7632
                    LayoutCachedTop =792
                    LayoutCachedWidth =13680
                    LayoutCachedHeight =5832
                    TabIndex =7
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                End
                Begin Label
                    Visible = NotDefault
                    FontItalic = NotDefault
                    OverlapFlags =255
                    TextAlign =2
                    Left =7632
                    Top =2232
                    Width =6048
                    Height =965
                    FontSize =12
                    FontWeight =700
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =13421772
                    Name ="LabelNoImage"
                    Caption ="NO IMAGE"
                    HorizontalAnchor =1
                    VerticalAnchor =2
                    LayoutCachedLeft =7632
                    LayoutCachedTop =2232
                    LayoutCachedWidth =13680
                    LayoutCachedHeight =3197
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =7632
                    Top =5976
                    Width =2808
                    Height =432
                    FontSize =9
                    FontWeight =600
                    TabIndex =5
                    ForeColor =8272943
                    Name ="cmdAddImage"
                    Caption ="ADD IMAGE"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Select and attach image for active product"

                    LayoutCachedLeft =7632
                    LayoutCachedTop =5976
                    LayoutCachedWidth =10440
                    LayoutCachedHeight =6408
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
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =10728
                    Top =5976
                    Width =2952
                    Height =432
                    FontSize =10
                    FontWeight =600
                    TabIndex =6
                    ForeColor =16777215
                    Name ="cmdClearImage"
                    Caption ="CLEAR"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Remove attached image from active product"

                    LayoutCachedLeft =10728
                    LayoutCachedTop =5976
                    LayoutCachedWidth =13680
                    LayoutCachedHeight =6408
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
                Begin Line
                    BorderWidth =1
                    OverlapFlags =85
                    Left =288
                    Top =7200
                    Width =13392
                    Height =14
                    BorderColor =15460325
                    Name ="linFooterSeparator"
                    HorizontalAnchor =2
                    VerticalAnchor =1
                    LayoutCachedLeft =288
                    LayoutCachedTop =7200
                    LayoutCachedWidth =13680
                    LayoutCachedHeight =7214
                    BorderThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =288
                    Top =7344
                    Width =3600
                    Height =360
                    FontSize =9
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =8947848
                    Name ="lblRecordCount"
                    Caption ="Showing: 0 / 0 products"
                    VerticalAnchor =1
                    LayoutCachedLeft =288
                    LayoutCachedTop =7344
                    LayoutCachedWidth =3888
                    LayoutCachedHeight =7704
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =10224
                    Top =7344
                    Width =3456
                    Height =360
                    FontSize =10
                    FontWeight =700
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =3355443
                    Name ="lblTotalValue"
                    Caption ="Total CBM: 0.00 | Total Value: $0.00"
                    HorizontalAnchor =1
                    VerticalAnchor =1
                    LayoutCachedLeft =10224
                    LayoutCachedTop =7344
                    LayoutCachedWidth =13680
                    LayoutCachedHeight =7704
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
' See "frmProductGridView.cls"
