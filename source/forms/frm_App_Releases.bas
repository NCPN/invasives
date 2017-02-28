Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =8160
    DatasheetFontHeight =10
    ItemSuffix =21
    Left =7965
    Top =2640
    Right =16125
    Bottom =10665
    DatasheetGridlinesColor =12632256
    OrderBy ="tsys_App_Releases.Release_date DESC"
    RecSrcDt = Begin
        0x16cf15acd1cee240
    End
    RecordSource ="tsys_App_Releases"
    Caption =" Application Releases"
    BeforeInsert ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
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
        Begin Section
            CanGrow = NotDefault
            Height =8040
            BackColor =9677753
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1560
                    Top =120
                    Width =5340
                    Height =252
                    ColumnWidth =1440
                    FontSize =9
                    Name ="txtRelease_ID"
                    ControlSource ="Release_ID"
                    StatusBarText ="Unique identifier for the release"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =60
                            Top =120
                            Width =1032
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="labRelease_ID"
                            Caption ="Release ID"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =6840
                    Top =840
                    Width =1200
                    Height =252
                    ColumnWidth =1140
                    FontSize =9
                    TabIndex =3
                    Name ="txtRelease_date"
                    ControlSource ="Release_date"
                    Format ="Short Date"
                    StatusBarText ="Date of the release"
                    FontName ="Arial"
                    InputMask ="99/99/0000;0;_"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5640
                            Top =840
                            Width =1185
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="labRelease_date"
                            Caption ="Release date"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =840
                    Width =1557
                    Height =252
                    ColumnWidth =972
                    FontSize =9
                    TabIndex =2
                    Name ="txtVersion_number"
                    ControlSource ="Version_number"
                    Format ="General Number"
                    StatusBarText ="Version control number"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =840
                            Width =1452
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="labVersion_number"
                            Caption ="Version number"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =1200
                    Width =3300
                    Height =252
                    ColumnWidth =2568
                    FontSize =9
                    TabIndex =4
                    Name ="txtFile_name"
                    ControlSource ="File_name"
                    StatusBarText ="Filename, used to identify older versions of the database"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1200
                            Width =924
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="labFile_name"
                            Caption ="File name"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =5940
                    Top =1200
                    Width =2106
                    Height =252
                    ColumnWidth =2568
                    FontSize =9
                    TabIndex =5
                    ColumnInfo ="\"\";\"\";\"10\";\"0\""
                    Name ="cmbRelease_by"
                    ControlSource ="Release_by"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [Last_Name] & (\", \"+[First_Name]) AS FullName FROM tlu_Contacts ORDER B"
                        "Y Last_Name, First_Name; "
                    StatusBarText ="Person who made the release"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4920
                            Top =1200
                            Width =1044
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="labRelease_by"
                            Caption ="Release by"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =2280
                    Width =6480
                    Height =1020
                    ColumnWidth =3000
                    FontSize =9
                    TabIndex =10
                    Name ="txtRelease_notes"
                    ControlSource ="Release_notes"
                    StatusBarText ="Release notes, which may include a summary of revisions"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2280
                            Width =1332
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="labRelease_notes"
                            Caption ="Release notes"
                            FontName ="Arial"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =215
                    Left =120
                    Top =3480
                    Width =7920
                    Height =4380
                    TabIndex =11
                    Name ="subBugs"
                    SourceObject ="Form.fsub_Bug_Reports"
                    LinkChildFields ="Release_ID"
                    LinkMasterFields ="Release_ID"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Top =3240
                            Width =1116
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="labBugs"
                            Caption ="Known bugs"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =1560
                    Width =1380
                    FontSize =9
                    TabIndex =6
                    Name ="txtAuthor_phone"
                    ControlSource ="Author_phone"
                    StatusBarText ="Phone number of application author"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1560
                            Width =1215
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="Label17"
                            Caption ="Author phone"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =5460
                    Top =1560
                    Width =2580
                    FontSize =9
                    TabIndex =7
                    Name ="txtAuthor_email"
                    ControlSource ="Author_email"
                    StatusBarText ="Email address of application author"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4320
                            Top =1560
                            Width =1155
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="Label18"
                            Caption ="Author email"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =4980
                    Top =1920
                    Width =3060
                    FontSize =9
                    TabIndex =9
                    Name ="txtAuthor_org_name"
                    ControlSource ="Author_org_name"
                    StatusBarText ="Name of organization for author's place of work"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3420
                            Top =1920
                            Width =1545
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="Label20"
                            Caption ="Author org. name"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =1920
                    Width =1080
                    FontSize =9
                    TabIndex =8
                    Name ="cboAuthor_org"
                    ControlSource ="Author_org"
                    RowSourceType ="Table/Query"
                    StatusBarText ="Organization (NPS Unit code) for the author's place of work"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =1920
                            Width =1440
                            Height =270
                            FontSize =9
                            FontWeight =700
                            Name ="Label19"
                            Caption ="Author org code"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =480
                    Width =6480
                    ColumnWidth =2568
                    FontSize =9
                    TabIndex =1
                    Name ="cmbDatabase_title"
                    ControlSource ="Database_title"
                    StatusBarText ="Title of the database"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =480
                            Width =1272
                            Height =252
                            FontSize =9
                            FontWeight =700
                            Name ="labDatabase_title"
                            Caption ="Database title"
                            FontName ="Arial"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Create the GUID primary key value
    If IsNull(Me!Release_ID) Then
        If GetDataType("tsys_App_Releases", "Release_ID") = dbText Then
            Me.Release_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure


End Sub

Private Sub Form_Close()
Dim varLastReleaseDate As Variant

varLastReleaseDate = DMax("[Release_date]", "tsys_App_Releases")

If Not IsNull(varLastReleaseDate) Then
    Call AddAppProperty("AppTitle", dbText, DLookup("[Database_title]", "tsys_App_Releases", "[Release_date]=#" & varLastReleaseDate & "#"))
    Application.RefreshTitleBar
End If
End Sub
