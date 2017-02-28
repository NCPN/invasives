Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9360
    DatasheetFontHeight =9
    ItemSuffix =15
    Left =270
    Top =210
    DatasheetGridlinesColor =12632256
    Filter ="(Unit_Code = 'CARE')"
    RecSrcDt = Begin
        0x611cfa7acc85e340
    End
    RecordSource ="qry_Sp_Rpt_by_Park"
    Caption ="rpt_Species_by_Park"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000902400008601000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    Begin
        Begin Label
            BackStyle =0
            TextAlign =1
            TextFontFamily =18
            FontSize =9
            FontWeight =700
            FontName ="Times New Roman"
        End
        Begin Rectangle
            BackStyle =0
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
            FontName ="Times New Roman"
            AsianLineBreak =255
        End
        Begin ListBox
            TextFontFamily =18
            OldBorderStyle =0
            FontName ="Times New Roman"
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontFamily =18
            BackStyle =0
            FontName ="Times New Roman"
        End
        Begin Subform
            OldBorderStyle =0
        End
        Begin BreakLevel
            ControlSource ="Unit_Code"
        End
        Begin BreakLevel
            ControlSource ="Plot_ID"
        End
        Begin BreakLevel
            ControlSource ="Master_Family"
        End
        Begin BreakLevel
            ControlSource ="Utah_Species"
        End
        Begin BreakLevel
            ControlSource ="Visit_Year"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1020
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =2160
                    Top =180
                    Width =4170
                    Height =600
                    FontSize =24
                    FontWeight =400
                    Name ="Label10"
                    Caption ="Species by Park"
                End
            End
        End
        Begin PageHeader
            Height =390
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =60
                    Top =60
                    Width =1080
                    Height =270
                    FontSize =10
                    Name ="Unit_Code_Label"
                    Caption ="Park Code"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =1365
                    Top =60
                    Width =735
                    Height =270
                    FontSize =10
                    Name ="Plot_ID_Label"
                    Caption ="Plot"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =5160
                    Top =60
                    Width =960
                    Height =270
                    FontSize =10
                    Name ="Utah_Species_Label"
                    Caption ="Species"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =8460
                    Top =60
                    Width =600
                    Height =270
                    FontSize =10
                    Name ="Year_Label"
                    Caption ="Year"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =2280
                    Top =60
                    Width =840
                    Height =270
                    FontSize =10
                    Name ="Master_Family_Label"
                    Caption ="Family"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =330
                    Width =9240
                    Name ="Line13"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =390
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =900
                    Height =270
                    Name ="Unit_Code"
                    ControlSource ="Unit_Code"
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1440
                    Top =60
                    Width =600
                    Height =270
                    TabIndex =1
                    Name ="Plot_ID"
                    ControlSource ="Plot_ID"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5160
                    Top =60
                    Width =3120
                    Height =270
                    ColumnWidth =2520
                    TabIndex =2
                    Name ="Utah_Species"
                    ControlSource ="Utah_Species"
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8460
                    Top =60
                    Width =600
                    Height =270
                    TabIndex =3
                    Name ="Year"
                    ControlSource ="Visit_Year"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2280
                    Top =60
                    Width =2640
                    Height =270
                    ColumnWidth =1395
                    TabIndex =4
                    Name ="Master_Family"
                    ControlSource ="Master_Family"
                End
            End
        End
        Begin PageFooter
            Height =510
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =240
                    Width =4560
                    Height =270
                    Name ="Text11"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4740
                    Top =240
                    Width =4560
                    Height =270
                    TabIndex =1
                    Name ="Text12"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                End
                Begin Line
                    Left =60
                    Top =240
                    Width =9240
                    Name ="Line14"
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
