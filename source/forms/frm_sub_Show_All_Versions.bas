Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    AllowEdits = NotDefault
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4440
    DatasheetFontHeight =10
    ItemSuffix =7
    Left =390
    Top =360
    Right =5910
    Bottom =3660
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x4f5eb3fe2dffe240
    End
    RecordSource ="tbl_SOP_version"
    Caption ="Show All Versions Subform"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin Tab
            BackStyle =0
        End
        Begin FormHeader
            Height =360
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Top =60
                    Width =1380
                    Height =240
                    FontWeight =700
                    Name ="SOP_number_Label"
                    Caption ="SOP Number"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =1620
                    Top =60
                    Width =1140
                    Height =240
                    FontWeight =700
                    Name ="SOP_version_number_Label"
                    Caption ="SOP Version"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3000
                    Top =60
                    Width =960
                    Height =240
                    FontWeight =700
                    Name ="active_flag_Label"
                    Caption ="Active"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            Height =360
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =180
                    Top =60
                    Width =1140
                    Height =255
                    ColumnWidth =900
                    Name ="SOP_number"
                    ControlSource ="SOP_number"
                    StatusBarText ="SOP number"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1680
                    Top =60
                    Width =1080
                    Height =255
                    ColumnWidth =900
                    TabIndex =1
                    Name ="SOP_version_number"
                    ControlSource ="SOP_version_number"
                    Format ="Fixed"
                    StatusBarText ="SOP version number"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3300
                    Top =60
                    TabIndex =2
                    Name ="active_flag"
                    ControlSource ="active_flag"
                    StatusBarText ="Yes indicates SOP is active"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =1020
                    Top =60
                    Width =480
                    ColumnWidth =1815
                    TabIndex =3
                    Name ="version_key_number"
                    ControlSource ="version_key_number"
                    StatusBarText ="Protocol version key number (maintained in SOP #10)"
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
