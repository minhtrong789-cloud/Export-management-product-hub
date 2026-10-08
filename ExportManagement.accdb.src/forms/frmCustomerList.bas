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
    DatasheetFontHeight =11
    ItemSuffix =15
    Left =3060
    Top =2200
    Right =16460
    Bottom =8470
    OnUnload ="[Event Procedure]"
    BeforeDelConfirm ="[Event Procedure]"
    AfterDelConfirm ="[Event Procedure]"
    RecSrcDt = Begin
        0x399e949f7796e640
    End
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Segoe UI"
    OnLoad ="[Event Procedure]"
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
        Begin Section
            Height =8640
            Name ="Detail"
            AutoHeight =1
            AlternateBackColor =16777215
            AlternateBackShade =95.0
            Begin
                Begin CheckBox
                    OverlapFlags =93
                    OldBorderStyle =0
                    Width =576
                    Height =360
                    ColumnWidth =1320
                    ColumnOrder =1
                    BorderColor =8272943
                    Name ="chkContextAction"

                    LayoutCachedWidth =576
                    LayoutCachedHeight =360
                    BorderThemeColorIndex =-1
                    DatasheetCaption ="Selected"
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =255
                    TextAlign =3
                    IMESentenceMode =3
                    Width =576
                    Height =360
                    ColumnWidth =630
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =1
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtRowNum"

                    LayoutCachedWidth =576
                    LayoutCachedHeight =360
                    DatasheetCaption ="No."
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    OldBorderStyle =0
                    BorderWidth =1
                    OverlapFlags =255
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Width =3600
                    Height =360
                    ColumnWidth =1930
                    ColumnOrder =4
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =13421772
                    ForeColor =8272943
                    Name ="txtCompanyName"

                    LayoutCachedWidth =3600
                    LayoutCachedHeight =360
                    DatasheetCaption ="Company Name"
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =255
                    TextAlign =1
                    TextFontFamily =49
                    IMESentenceMode =3
                    Width =1728
                    Height =360
                    ColumnWidth =1130
                    ColumnOrder =3
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtCountry"
                    FontName ="Courier New"

                    LayoutCachedWidth =1728
                    LayoutCachedHeight =360
                    DatasheetCaption ="Country"
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    ScrollBars =2
                    BorderWidth =1
                    OverlapFlags =255
                    TextAlign =1
                    IMESentenceMode =3
                    Width =2592
                    Height =360
                    ColumnWidth =1820
                    ColumnOrder =5
                    FontSize =10
                    TabIndex =4
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16447992
                    BorderColor =14737632
                    ForeColor =3355443
                    Name ="txtContactPerson"

                    LayoutCachedWidth =2592
                    LayoutCachedHeight =360
                    DatasheetCaption ="Contact Person"
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =255
                    TextAlign =1
                    TextFontFamily =49
                    IMESentenceMode =3
                    Width =1872
                    Height =360
                    ColumnWidth =1770
                    ColumnOrder =6
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtContactPersonPhone"
                    Format ="@"
                    FontName ="Courier New"

                    LayoutCachedWidth =1872
                    LayoutCachedHeight =360
                    DatasheetCaption ="Contact Phone"
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    ScrollBars =2
                    BorderWidth =1
                    OverlapFlags =255
                    TextAlign =1
                    IMESentenceMode =3
                    Width =2592
                    Height =360
                    ColumnWidth =1670
                    ColumnOrder =7
                    FontSize =10
                    TabIndex =6
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16447992
                    BorderColor =14737632
                    ForeColor =3355443
                    Name ="txtContactPersonEmail"

                    LayoutCachedWidth =2592
                    LayoutCachedHeight =360
                    DatasheetCaption ="Contact Email"
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =255
                    TextAlign =1
                    TextFontFamily =49
                    IMESentenceMode =3
                    Width =1872
                    Height =360
                    ColumnWidth =1610
                    ColumnOrder =8
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtCompanyPhone"
                    Format ="@"
                    FontName ="Courier New"

                    LayoutCachedWidth =1872
                    LayoutCachedHeight =360
                    DatasheetCaption ="Office Phone"
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    ScrollBars =2
                    BorderWidth =1
                    OverlapFlags =255
                    TextAlign =1
                    IMESentenceMode =3
                    Width =2592
                    Height =360
                    ColumnWidth =980
                    ColumnOrder =9
                    FontSize =10
                    TabIndex =8
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16447992
                    BorderColor =14737632
                    ForeColor =3355443
                    Name ="txtCompanyEmail"

                    LayoutCachedWidth =2592
                    LayoutCachedHeight =360
                    DatasheetCaption ="Office Email"
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    ScrollBars =2
                    BorderWidth =1
                    OverlapFlags =255
                    TextAlign =1
                    IMESentenceMode =3
                    Width =2592
                    Height =360
                    ColumnOrder =10
                    FontSize =10
                    TabIndex =9
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16447992
                    BorderColor =14737632
                    ForeColor =3355443
                    Name ="txtCompanyWebsite"

                    LayoutCachedWidth =2592
                    LayoutCachedHeight =360
                    DatasheetCaption ="Website"
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    ScrollBars =2
                    BorderWidth =1
                    OverlapFlags =255
                    TextAlign =1
                    IMESentenceMode =3
                    Width =3600
                    Height =360
                    ColumnOrder =11
                    FontSize =10
                    TabIndex =10
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16447992
                    BorderColor =14737632
                    ForeColor =3355443
                    Name ="txtCompanyAddress1"

                    LayoutCachedWidth =3600
                    LayoutCachedHeight =360
                    DatasheetCaption ="Address Line 1"
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    ScrollBars =2
                    BorderWidth =1
                    OverlapFlags =255
                    TextAlign =1
                    IMESentenceMode =3
                    Width =3600
                    Height =360
                    ColumnOrder =12
                    FontSize =10
                    TabIndex =11
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16447992
                    BorderColor =14737632
                    ForeColor =3355443
                    Name ="txtCompanyAddress2"

                    LayoutCachedWidth =3600
                    LayoutCachedHeight =360
                    DatasheetCaption ="Address Line 2"
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =255
                    TextAlign =1
                    TextFontFamily =49
                    IMESentenceMode =3
                    Width =1728
                    Height =360
                    ColumnOrder =13
                    FontSize =10
                    FontWeight =700
                    TabIndex =12
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtTaxID"
                    Format ="@"
                    FontName ="Courier New"

                    LayoutCachedWidth =1728
                    LayoutCachedHeight =360
                    DatasheetCaption ="Tax ID"
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    ScrollBars =2
                    BorderWidth =1
                    OverlapFlags =255
                    TextAlign =1
                    IMESentenceMode =3
                    Width =2880
                    Height =360
                    ColumnOrder =14
                    FontSize =10
                    TabIndex =13
                    LeftMargin =72
                    TopMargin =30
                    BackColor =16447992
                    BorderColor =14737632
                    ForeColor =3355443
                    Name ="txtNotes"

                    LayoutCachedWidth =2880
                    LayoutCachedHeight =360
                    DatasheetCaption ="Notes"
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin TextBox
                    BorderWidth =1
                    OverlapFlags =247
                    TextAlign =1
                    TextFontFamily =49
                    IMESentenceMode =3
                    Width =1728
                    Height =360
                    ColumnOrder =2
                    FontSize =10
                    FontWeight =700
                    TabIndex =14
                    LeftMargin =72
                    TopMargin =30
                    BorderColor =13421772
                    ForeColor =3355443
                    Name ="txtContinent"
                    ControlSource ="Continent"
                    FontName ="Courier New"

                    LayoutCachedWidth =1728
                    LayoutCachedHeight =360
                    DatasheetCaption ="Continent"
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
            End
        End
    End
End
CodeBehindForm
' See "frmCustomerList.cls"
