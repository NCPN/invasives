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
    Width =10080
    DatasheetFontHeight =9
    ItemSuffix =38
    Left =270
    Top =210
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa124e8882c61e340
    End
    RecordSource ="qry_sel_Presence_Report"
    Caption ="rpt_Species_Presence"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xd0020000d0020000d0020000d002000000000000602700006801000001000000 ,
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
            ControlSource ="Plot_ID"
        End
        Begin BreakLevel
            ControlSource ="Master_Family"
        End
        Begin BreakLevel
            ControlSource ="Utah_Species"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =840
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =240
                    Top =180
                    Width =5580
                    Height =600
                    FontSize =20
                    FontWeight =400
                    Name ="Label28"
                    Caption ="Species Presence by Plot"
                End
                Begin Line
                    BorderWidth =2
                    Top =60
                    Width =10020
                    Name ="Line31"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6180
                    Top =300
                    Width =3180
                    Height =360
                    ColumnWidth =3108
                    FontSize =14
                    Name ="ParkName"
                    ControlSource ="ParkName"
                    StatusBarText ="Full name of park where data were collected"
                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =780
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    IMESentenceMode =3
                    Left =600
                    Top =60
                    Width =660
                    Height =300
                    FontSize =12
                    FontWeight =700
                    Name ="Plot_ID"
                    ControlSource ="Plot_ID"
                    StatusBarText ="Plot_ID"
                    Begin
                        Begin Label
                            Left =60
                            Top =60
                            Width =540
                            Height =300
                            FontSize =12
                            Name ="Plot_ID_Label"
                            Caption ="Plot"
                        End
                    End
                End
                Begin Label
                    Left =60
                    Top =480
                    Width =1860
                    Height =270
                    FontSize =10
                    Name ="Master_Family_Label"
                    Caption ="Master Family"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =1980
                    Top =480
                    Width =1860
                    Height =270
                    FontSize =10
                    Name ="Utah_Species_Label"
                    Caption ="Utah Species"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =3900
                    Top =480
                    Width =540
                    Height =270
                    FontSize =10
                    Name ="L1"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =4500
                    Top =480
                    Width =540
                    Height =270
                    FontSize =10
                    Name ="L2"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =5100
                    Top =480
                    Width =540
                    Height =270
                    FontSize =10
                    Name ="L3"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =5700
                    Top =480
                    Width =540
                    Height =270
                    FontSize =10
                    Name ="L4"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =6300
                    Top =480
                    Width =540
                    Height =270
                    FontSize =10
                    Name ="L5"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =6900
                    Top =480
                    Width =540
                    Height =270
                    FontSize =10
                    Name ="L6"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =7500
                    Top =480
                    Width =540
                    Height =270
                    FontSize =10
                    Name ="L7"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =8100
                    Top =480
                    Width =540
                    Height =270
                    FontSize =10
                    Name ="L8"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =8700
                    Top =480
                    Width =540
                    Height =270
                    FontSize =10
                    Name ="L9"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =9300
                    Top =480
                    Width =540
                    Height =270
                    FontSize =10
                    Name ="L10"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Width =10020
                    Name ="Line37"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =360
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =1860
                    Height =270
                    Name ="Master_Family"
                    ControlSource ="Master_Family"
                    StatusBarText ="Master_Family"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1980
                    Top =60
                    Width =1860
                    Height =270
                    ColumnWidth =3345
                    TabIndex =1
                    Name ="Utah_Species"
                    ControlSource ="Utah_Species"
                    StatusBarText ="Utah Species (Welsh et al 2003)"
                End
                Begin CheckBox
                    Left =4080
                    Top =120
                    TabIndex =2
                    Name ="P1"
                    ControlSource ="P1"
                    StatusBarText ="10 yes/no columns will last 10 years"
                End
                Begin CheckBox
                    Left =4680
                    Top =120
                    TabIndex =3
                    Name ="P2"
                    ControlSource ="P2"
                End
                Begin CheckBox
                    Left =5280
                    Top =120
                    TabIndex =4
                    Name ="P3"
                    ControlSource ="P3"
                End
                Begin CheckBox
                    Left =5880
                    Top =120
                    TabIndex =5
                    Name ="P4"
                    ControlSource ="P4"
                End
                Begin CheckBox
                    Left =6480
                    Top =120
                    TabIndex =6
                    Name ="P5"
                    ControlSource ="P5"
                End
                Begin CheckBox
                    Left =7080
                    Top =120
                    TabIndex =7
                    Name ="P6"
                    ControlSource ="P6"
                End
                Begin CheckBox
                    Left =7680
                    Top =120
                    TabIndex =8
                    Name ="P7"
                    ControlSource ="P7"
                End
                Begin CheckBox
                    Left =8280
                    Top =120
                    TabIndex =9
                    Name ="P8"
                    ControlSource ="P8"
                End
                Begin CheckBox
                    Left =8880
                    Top =120
                    TabIndex =10
                    Name ="P9"
                    ControlSource ="P9"
                End
                Begin CheckBox
                    Left =9480
                    Top =120
                    TabIndex =11
                    Name ="P10"
                    ControlSource ="P10"
                End
                Begin Line
                    Left =60
                    Top =60
                    Width =9960
                    Name ="Line35"
                End
            End
        End
        Begin PageFooter
            Height =420
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =120
                    Width =2820
                    Height =270
                    Name ="Text29"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7860
                    Top =120
                    Width =2220
                    Height =270
                    TabIndex =1
                    Name ="Text30"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                End
                Begin Line
                    Width =10020
                    Name ="Line34"
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
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Report_Open(Cancel As Integer)

  Dim YearCount As Integer
  Dim ControlName As String
  Dim ControlYear As String
  Dim YearWork As String
  Dim FieldIndex As Integer

  YearWork = Me.OpenArgs
  FieldIndex = 1
  Do Until IsNull(YearWork) Or YearWork = ""
    ControlYear = Left(YearWork, 4)
    ControlName = "L" & FieldIndex
    Me.Controls(ControlName).Caption = ControlYear  ' set year in column headings
    If Len(YearWork) = 4 Then
      Exit Do
    Else
      YearWork = Right(YearWork, Len(YearWork) - 4)
      FieldIndex = FieldIndex + 1
    End If
  Loop
  FieldIndex = FieldIndex + 1
  Do Until FieldIndex > 10
    ControlName = "P" & FieldIndex
    Me.Controls(ControlName).Visible = False  ' dont show unused controls
    FieldIndex = FieldIndex + 1
  Loop
  
End Sub
