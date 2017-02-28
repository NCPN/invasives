Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =126
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10080
    DatasheetFontHeight =9
    ItemSuffix =37
    Left =705
    Top =1035
    DatasheetGridlinesColor =12632256
    Filter ="([Unit_Code] = 'BLCA' AND [Plot_Id] = 141AND [Visit_Year] = '2009')"
    RecSrcDt = Begin
        0x04ae7a726ca8e340
    End
    RecordSource ="qry_sel_OT_Census_Report"
    Caption ="rpt_OT_Census"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x40020000d002000040020000d002000000000000602700008601000001000000 ,
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
            GroupHeader = NotDefault
            ControlSource ="Unit_Code"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Plot_ID"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Quad"
        End
        Begin BreakLevel
            ControlSource ="Tag_No"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =900
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =1440
                    Top =180
                    Width =7380
                    Height =600
                    FontSize =24
                    FontWeight =400
                    Name ="Label20"
                    Caption ="Overstory Census Revisit"
                End
                Begin Line
                    BorderWidth =2
                    Left =120
                    Top =900
                    Width =9792
                    Name ="Line34"
                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =405
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    IMESentenceMode =3
                    Left =1620
                    Height =405
                    FontSize =14
                    FontWeight =700
                    Name ="Unit_Code"
                    ControlSource ="Unit_Code"
                    StatusBarText ="Park Code."
                    Begin
                        Begin Label
                            Left =60
                            Width =1500
                            Height =405
                            FontSize =14
                            Name ="Unit_Code_Label"
                            Caption ="Park Code"
                        End
                    End
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =420
            BreakLevel =1
            Name ="GroupHeader1"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    IMESentenceMode =3
                    Left =1620
                    Height =390
                    FontSize =14
                    Name ="Plot_ID"
                    ControlSource ="Plot_ID"
                    StatusBarText ="Plot identifier"
                    Begin
                        Begin Label
                            Left =60
                            Width =1500
                            Height =390
                            FontSize =14
                            FontWeight =400
                            Name ="Plot_ID_Label"
                            Caption ="Plot"
                        End
                    End
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =720
            BreakLevel =2
            Name ="GroupHeader2"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    IMESentenceMode =3
                    Left =660
                    Top =60
                    Width =420
                    Height =300
                    FontSize =10
                    Name ="Quad"
                    ControlSource ="Quad"
                    StatusBarText ="Quadrat number"
                    Begin
                        Begin Label
                            Left =60
                            Top =60
                            Width =540
                            Height =300
                            FontSize =10
                            Name ="Quad_Label"
                            Caption ="Quad"
                        End
                    End
                End
                Begin Label
                    TextAlign =2
                    Left =60
                    Top =420
                    Width =540
                    Height =270
                    FontSize =10
                    Name ="Tag_No_Label"
                    Caption ="Tag #"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =660
                    Top =420
                    Width =2220
                    Height =270
                    FontSize =10
                    Name ="Species_Label"
                    Caption ="Species"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =2940
                    Top =420
                    Width =480
                    Height =270
                    FontSize =10
                    Name ="DBH_Label"
                    Caption ="DBH"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =3480
                    Top =420
                    Width =1155
                    Height =270
                    FontSize =10
                    Name ="Crown_Class_Label"
                    Caption ="Crown Class"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =4680
                    Top =420
                    Width =1320
                    Height =270
                    FontSize =10
                    Name ="Class_Description_Label"
                    Caption ="Crown Health"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =6360
                    Top =420
                    Width =3000
                    Height =270
                    FontSize =10
                    Name ="Notes_Label"
                    Caption ="Notes"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =360
                    Width =9900
                    Name ="Line35"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =60
                    Width =9900
                    Name ="Line36"
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
                    Width =540
                    Height =270
                    Name ="Tag_No"
                    ControlSource ="Tag_No"
                    StatusBarText ="Tag number"
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =660
                    Top =60
                    Width =540
                    Height =270
                    TabIndex =1
                    Name ="Species"
                    ControlSource ="Species"
                    StatusBarText ="Species code"
                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1260
                    Top =60
                    Width =1620
                    Height =270
                    ColumnWidth =3345
                    TabIndex =2
                    Name ="Utah_Species"
                    ControlSource ="Utah_Species"
                    StatusBarText ="Utah Species (Welsh et al 2003)"
                End
                Begin TextBox
                    DecimalPlaces =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2940
                    Top =60
                    Width =480
                    Height =270
                    TabIndex =3
                    Name ="DBH"
                    ControlSource ="DBH"
                    StatusBarText ="Diameter at breast height in centimeters"
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3480
                    Top =60
                    Width =1140
                    Height =270
                    TabIndex =4
                    Name ="Crown_Class"
                    ControlSource ="Crown_Class"
                    StatusBarText ="Crown class"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4680
                    Top =60
                    Width =1320
                    Height =270
                    ColumnWidth =1845
                    TabIndex =5
                    Name ="Class_Description"
                    ControlSource ="Class_Description"
                    StatusBarText ="Health class description"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6120
                    Top =60
                    Width =3840
                    Height =270
                    TabIndex =6
                    Name ="Notes"
                    ControlSource ="Notes"
                    StatusBarText ="Notes about any significant damage to a living tree"
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
                    Name ="Text21"
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
                    Name ="Text22"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                End
                Begin Line
                    Width =9360
                    Name ="Line32"
                End
                Begin Line
                    Top =30
                    Width =9360
                    Name ="Line33"
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
