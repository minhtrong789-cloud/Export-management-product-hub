Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
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
    ItemSuffix =9
    Top =710
    Right =18940
    Bottom =9820
    RecSrcDt = Begin
        0x8e588b9f7796e640
    End
    DatasheetFontName ="Segoe UI"
    AllowDatasheetView =0
    FilterOnLoad =0
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
                    Width =3168
                    Height =432
                    FontSize =10
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="cboFilterCountry"
                    RowSourceType ="Table/Query"
                    Format ="@;\"[ All Countries ]\""
                    LeftMargin =72
                    TopMargin =30

                    LayoutCachedLeft =288
                    LayoutCachedTop =144
                    LayoutCachedWidth =3456
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
                    Left =3600
                    Top =144
                    Width =2160
                    Height =432
                    FontSize =10
                    TabIndex =1
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="cboFilterActive"
                    RowSourceType ="Table/Query"
                    Format ="@;\"[ All Status ]\""
                    LeftMargin =72
                    TopMargin =30

                    LayoutCachedLeft =3600
                    LayoutCachedTop =144
                    LayoutCachedWidth =5760
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
                    BorderColor =15387847
                    Name ="linMainDivider"
                    HorizontalAnchor =2
                    LayoutCachedLeft =288
                    LayoutCachedTop =691
                    LayoutCachedWidth =13680
                    LayoutCachedHeight =705
                    BorderThemeColorIndex =-1
                End
                Begin Subform
                    OverlapFlags =85
                    BorderWidth =1
                    Left =288
                    Top =792
                    Width =13392
                    Height =6264
                    TabIndex =2
                    BorderColor =15387847
                    Name ="sfCustomerGrid"
                    SourceObject ="Form.frmCustomerList"
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
                    Name ="cmdNewCustomer"
                    Caption ="+ NEW CUSTOMER"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add a new customer entry"

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
                    ControlTipText ="Export customer dataset to Excel"

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
                Begin CommandButton
                    OverlapFlags =85
                    Left =13824
                    Top =1440
                    Width =1699
                    Height =432
                    FontSize =10
                    FontWeight =600
                    TabIndex =5
                    ForeColor =16777215
                    Name ="cmdImport"
                    Caption ="IMPORT"
                    ControlTipText ="Import customer records from external file"

                    LayoutCachedLeft =13824
                    LayoutCachedTop =1440
                    LayoutCachedWidth =15523
                    LayoutCachedHeight =1872
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
                Begin Line
                    BorderWidth =1
                    OverlapFlags =85
                    Left =288
                    Top =7200
                    Width =13392
                    Height =14
                    BorderColor =15387847
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
                    Width =5040
                    Height =403
                    FontSize =9
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =16777215
                    ForeColor =8947848
                    Name ="lblRecordCount"
                    Caption ="Showing: 0 / 0 customers"
                    VerticalAnchor =1
                    LayoutCachedLeft =288
                    LayoutCachedTop =7344
                    LayoutCachedWidth =5328
                    LayoutCachedHeight =7747
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
' See "frmCustomerGridView.cls"
