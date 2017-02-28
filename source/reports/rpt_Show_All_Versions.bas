Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =124
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9360
    DatasheetFontHeight =10
    ItemSuffix =28
    Top =210
    DatasheetGridlinesColor =12632256
    Filter ="(version_key_number = 1)"
    RecSrcDt = Begin
        0xa0afbc4535fee240
    End
    RecordSource ="qry_version_key_by_number"
    Caption ="Version Key Listing"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000902400006801000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    Begin
        Begin Label
            BackStyle =0
            TextAlign =1
            FontSize =9
            FontWeight =700
            ForeColor =128
            FontName ="Arial"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
        End
        Begin Line
            BorderColor =128
        End
        Begin Image
            OldBorderStyle =0
            PictureAlignment =2
        End
        Begin CheckBox
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =18
            BackStyle =0
            FontSize =9
            FontName ="Times New Roman"
            AsianLineBreak =255
        End
        Begin ListBox
            TextFontFamily =18
            OldBorderStyle =0
            FontSize =9
            FontName ="Times New Roman"
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontFamily =18
            BackStyle =0
            FontSize =9
            FontName ="Times New Roman"
        End
        Begin Subform
            OldBorderStyle =0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="version_key_number"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="version_key_number"
        End
        Begin BreakLevel
            ControlSource ="SOP_number"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =960
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextFontFamily =18
                    Left =60
                    Top =60
                    Width =7155
                    Height =510
                    FontSize =20
                    FontWeight =900
                    Name ="Label14"
                    Caption ="Master Version Key Information Listing"
                    FontName ="Times New Roman"
                End
                Begin Line
                    BorderWidth =3
                    Top =60
                    Width =9360
                    BorderColor =0
                    Name ="Line17"
                End
                Begin Line
                    BorderWidth =3
                    Top =90
                    Width =9360
                    BorderColor =0
                    Name ="Line18"
                End
                Begin Line
                    BorderWidth =3
                    Top =930
                    Width =9360
                    BorderColor =0
                    Name ="Line19"
                End
                Begin Line
                    BorderWidth =3
                    Top =960
                    Width =9360
                    BorderColor =0
                    Name ="Line20"
                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =480
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2160
                    Top =180
                    Width =2640
                    Height =300
                    ColumnWidth =1815
                    FontSize =10
                    FontWeight =700
                    Name ="version_key_number"
                    ControlSource ="version_key_number"
                    StatusBarText ="Protocol version key number (maintained in SOP #10)"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            Left =60
                            Top =180
                            Width =2025
                            Height =300
                            FontSize =10
                            ForeColor =0
                            Name ="version_key_number_Label"
                            Caption ="Version Key Number"
                        End
                    End
                End
                Begin Line
                    BorderWidth =2
                    Top =60
                    Width =9360
                    BorderColor =0
                    Name ="Line27"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =1620
            BreakLevel =1
            Name ="GroupHeader1"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    IMESentenceMode =3
                    Left =2040
                    Top =60
                    Width =1035
                    Height =300
                    Name ="version_key_date"
                    ControlSource ="version_key_date"
                    Format ="Short Date"
                    StatusBarText ="Date of protocol version key number"
                    Begin
                        Begin Label
                            Left =60
                            Top =60
                            Width =1860
                            Height =285
                            Name ="version_key_date_Label"
                            Caption ="Version Key Date"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    IMESentenceMode =3
                    Left =2040
                    Top =480
                    Width =1020
                    Height =300
                    TabIndex =1
                    Name ="narrative_version"
                    ControlSource ="narrative_version"
                    Format ="Fixed"
                    StatusBarText ="Version of protocol narrative"
                    Begin
                        Begin Label
                            Left =60
                            Top =480
                            Width =1860
                            Height =285
                            Name ="narrative_version_Label"
                            Caption ="Narrative Version"
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    OldBorderStyle =1
                    IMESentenceMode =3
                    Left =2040
                    Top =900
                    Width =4560
                    Height =300
                    ColumnWidth =3615
                    TabIndex =2
                    Name ="version_comments"
                    ControlSource ="version_comments"
                    StatusBarText ="Comments regarding version, if any"
                    Begin
                        Begin Label
                            Left =60
                            Top =900
                            Width =1860
                            Height =285
                            Name ="version_comments_Label"
                            Caption ="Version Comments"
                        End
                    End
                End
                Begin Label
                    TextAlign =3
                    Left =1140
                    Top =1320
                    Width =1215
                    Height =255
                    Name ="SOP_number_Label"
                    Caption ="SOP Number"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =2580
                    Top =1320
                    Width =2145
                    Height =255
                    Name ="SOP_version_number_Label"
                    Caption ="SOP Version Number"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =4785
                    Top =1320
                    Width =1005
                    Height =255
                    Name ="active_flag_Label"
                    Caption ="Active"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =1140
                    Top =1320
                    Width =4650
                    Name ="Line21"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =1140
                    Top =1290
                    Width =4650
                    Name ="Line22"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =1140
                    Top =1575
                    Width =4650
                    Name ="Line23"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =1140
                    Top =1605
                    Width =4650
                    Name ="Line24"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =360
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1140
                    Top =60
                    Width =1215
                    Height =300
                    Name ="SOP_number"
                    ControlSource ="SOP_number"
                    StatusBarText ="SOP number"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2415
                    Top =60
                    Width =2310
                    Height =300
                    ColumnWidth =1785
                    TabIndex =1
                    Name ="SOP_version_number"
                    ControlSource ="SOP_version_number"
                    Format ="Fixed"
                    StatusBarText ="SOP version number"
                End
                Begin CheckBox
                    Left =5040
                    Top =60
                    TabIndex =2
                    Name ="active_flag"
                    ControlSource ="active_flag"
                    StatusBarText ="Yes indicates SOP is active"
                End
            End
        End
        Begin PageFooter
            Height =540
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =240
                    Width =4560
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Text15"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4740
                    Top =240
                    Width =4560
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="Text16"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                End
                Begin Line
                    BorderWidth =3
                    Width =9360
                    BorderColor =0
                    Name ="Line25"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =9360
                    BorderColor =0
                    Name ="Line26"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
