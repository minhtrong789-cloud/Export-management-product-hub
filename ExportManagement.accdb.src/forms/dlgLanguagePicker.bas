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
    Width =5040
    DatasheetFontHeight =11
    ItemSuffix =20
    Left =6000
    Top =3040
    Right =18940
    Bottom =9490
    RecSrcDt = Begin
        0xf11d858ad688e640
    End
    Caption ="Export Configuration"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Segoe UI"
    AllowDatasheetView =0
    FilterOnLoad =0
    NavigationCaption ="Export Configuration"
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
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
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
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =2880
            BackColor =16447477
            Name ="Detail"
            AutoHeight =1
            AlternateBackColor =16777215
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =288
                    Top =144
                    Width =4464
                    Height =432
                    FontWeight =700
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =5258796
                    Name ="lblTitle"
                    Caption ="SELECT EXPORT LANGUAGE"
                    LayoutCachedLeft =288
                    LayoutCachedTop =144
                    LayoutCachedWidth =4752
                    LayoutCachedHeight =576
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin OptionGroup
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =93
                    Left =288
                    Top =720
                    Width =4464
                    Height =1296
                    BorderColor =13091773
                    Name ="fraLang"
                    DefaultValue ="1"

                    LayoutCachedLeft =288
                    LayoutCachedTop =720
                    LayoutCachedWidth =4752
                    LayoutCachedHeight =2016
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    Begin
                        Begin OptionButton
                            OverlapFlags =87
                            BorderWidth =1
                            Left =504
                            Top =1224
                            Width =288
                            Height =288
                            BorderColor =13091773
                            Name ="optVN"

                            LayoutCachedLeft =504
                            LayoutCachedTop =1224
                            LayoutCachedWidth =792
                            LayoutCachedHeight =1512
                            BorderThemeColorIndex =-1
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            BorderWidth =1
                            Left =2016
                            Top =1224
                            Width =288
                            Height =288
                            BorderColor =13091773
                            Name ="optEN"

                            LayoutCachedLeft =2016
                            LayoutCachedTop =1224
                            LayoutCachedWidth =2304
                            LayoutCachedHeight =1512
                            BorderThemeColorIndex =-1
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            BorderWidth =1
                            Left =3528
                            Top =1224
                            Width =288
                            Height =288
                            BorderColor =13091773
                            Name ="optBI"

                            LayoutCachedLeft =3528
                            LayoutCachedTop =1224
                            LayoutCachedWidth =3816
                            LayoutCachedHeight =1512
                            BorderThemeColorIndex =-1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1008
                    Top =2160
                    Height =504
                    FontSize =10
                    FontWeight =600
                    TabIndex =1
                    ForeColor =16777215
                    Name ="btnOK"
                    Caption ="CONFIRM"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =1008
                    LayoutCachedTop =2160
                    LayoutCachedWidth =2448
                    LayoutCachedHeight =2664
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =2592
                    Top =2160
                    Height =504
                    FontSize =10
                    FontWeight =600
                    TabIndex =2
                    ForeColor =16777215
                    Name ="btnCancel"
                    Caption ="CANCEL"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2592
                    LayoutCachedTop =2160
                    LayoutCachedWidth =4032
                    LayoutCachedHeight =2664
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
                    OverlapFlags =215
                    TextAlign =1
                    Left =864
                    Top =1224
                    Width =922
                    Height =310
                    FontSize =9
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =3355443
                    Name ="lblVN"
                    Caption ="Tiếng Việt"
                    LayoutCachedLeft =864
                    LayoutCachedTop =1224
                    LayoutCachedWidth =1786
                    LayoutCachedHeight =1534
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =2376
                    Top =1224
                    Width =720
                    Height =288
                    FontSize =9
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =3355443
                    Name ="lblEN"
                    Caption ="English"
                    LayoutCachedLeft =2376
                    LayoutCachedTop =1224
                    LayoutCachedWidth =3096
                    LayoutCachedHeight =1512
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =3888
                    Top =1224
                    Width =802
                    Height =310
                    FontSize =9
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =3355443
                    Name ="lblBI"
                    Caption ="Bilingual"
                    LayoutCachedLeft =3888
                    LayoutCachedTop =1224
                    LayoutCachedWidth =4690
                    LayoutCachedHeight =1534
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
' See "dlgLanguagePicker.cls"
