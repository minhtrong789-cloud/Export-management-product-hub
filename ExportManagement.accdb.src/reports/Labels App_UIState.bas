Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =20
    GridY =24
    Width =11316
    DatasheetFontHeight =11
    ItemSuffix =3
    RecSrcDt = Begin
        0x368d9ec11787e640
    End
    RecordSource ="App_UIState"
    DatasheetFontName ="Segoe UI"
    FilterOnLoad =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
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
        Begin TextBox
            AutoLabel = NotDefault
            AddColon = NotDefault
            FontItalic = NotDefault
            FontUnderline = NotDefault
            FELineBreak = NotDefault
            OldBorderStyle =0
            BorderLineStyle =0
            LabelX =-1800
            FontWeight =300
            FontName ="Arial"
            AsianLineBreak =1
            ShowDatePicker =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin Section
            KeepTogether = NotDefault
            Height =16390
            Name ="Detail"
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OverlapFlags =4
                    TextAlign =1
                    IMESentenceMode =3
                    Left =144
                    Top =120
                    Width =11028
                    Name ="Text2"
                    ControlSource ="= Trim( [ViewName] & [ViewName] & [IsActive] & [InsertDate])"

                    LayoutCachedLeft =144
                    LayoutCachedTop =120
                    LayoutCachedWidth =11172
                    LayoutCachedHeight =360
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
