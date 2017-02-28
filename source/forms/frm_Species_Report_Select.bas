Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    DatasheetFontHeight =9
    ItemSuffix =12
    Left =2535
    Top =300
    Right =9480
    Bottom =5790
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x1385341e7574e340
    End
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin ListBox
            SpecialEffect =2
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            FontName ="Tahoma"
        End
        Begin Section
            Height =5760
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =525
                    Left =4155
                    Top =960
                    Width =900
                    ColumnInfo ="\"\";\"\";\"10\";\"8\""
                    Name ="Park_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qry_Parks.Unit_Code FROM qry_Parks; "
                    ColumnWidths ="525"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1935
                            Top =960
                            Width =2100
                            Height =245
                            FontWeight =700
                            Name ="Select a park if desired_Label"
                            Caption ="Select a park if desired"
                            EventProcPrefix ="Select_a_park_if_desired_Label"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3135
                    Top =2940
                    Width =1334
                    Height =300
                    TabIndex =3
                    Name ="Button_Close"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =720
                    Left =4155
                    Top =1380
                    Width =900
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"3\";\"2\""
                    Name ="Visit_Date"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qry_Event_Date.Visit_Year FROM qry_Event_Date; "
                    ColumnWidths ="720"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1935
                            Top =1380
                            Width =2100
                            Height =245
                            FontWeight =700
                            Name ="Select a date if desired_Label"
                            Caption ="Select a year if desired"
                            EventProcPrefix ="Select_a_date_if_desired_Label"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1980
                    Top =300
                    Width =3060
                    Height =420
                    FontSize =14
                    FontWeight =700
                    Name ="Label6"
                    Caption ="Species Report"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3135
                    Top =2400
                    Width =1320
                    Height =345
                    TabIndex =4
                    Name ="Button_rpt_by_Park"
                    Caption ="Report by Park"
                    OnClick ="[Event Procedure]"
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4155
                    Top =1800
                    Width =720
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"3\";\"2\""
                    Name ="Plot"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Locations.Plot_ID FROM tbl_Locations WHERE (((tbl_Locations.Unit_Code"
                        ")=\"cany\")) ORDER BY tbl_Locations.Plot_ID; "
                    ColumnWidths ="420"
                    OnGotFocus ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1935
                            Top =1800
                            Width =2100
                            Height =245
                            FontWeight =700
                            Name ="Plot_Select_Label"
                            Caption ="Select a plot if desired"
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
Private Sub Button_Close_Click()
On Error GoTo Err_Button_Close_Click

    DoCmd.Close

Exit_Button_Close_Click:
    Exit Sub

Err_Button_Close_Click:
    MsgBox Err.Description
    Resume Exit_Button_Close_Click
End Sub

Private Sub Button_rpt_by_Park_Click()
On Error GoTo Err_Button_rpt_by_Park_Click

    Dim strWhere As String
    Dim stDocName As String

    stDocName = "rpt_Species_by_Park"
    ' Set where condition if needed
    If (IsNull(Me!Park_Code) + IsNull(Me!Visit_Date) + IsNull(Me!Plot)) > -3 Then
      If Not IsNull(Me!Park_Code) Then
        strWhere = "Unit_Code = '" & Me!Park_Code & "'"
        If Not IsNull(Me!Plot) Then
          strWhere = strWhere & " And Plot_ID = " & Me!Plot
        End If
        If Not IsNull(Me!Visit_Date) Then
          strWhere = strWhere & " AND Visit_Year = " & Me!Visit_Date
        End If
      Else
        strWhere = "Visit_Year = " & Me!Visit_Date
      End If
    End If
    DoCmd.OpenReport stDocName, acViewPreview, , strWhere

Exit_Button_rpt_by_Park_Click:
    Exit Sub

Err_Button_rpt_by_Park_Click:
    MsgBox Err.Description
    Resume Exit_Button_rpt_by_Park_Click
    
End Sub

Private Sub Park_Code_AfterUpdate()

  If Not IsNull(Me!Park_Code) Then
    Me!Plot.RowSource = "SELECT Plot_ID FROM tbl_Locations WHERE Unit_Code = '" & Me!Park_Code & "' ORDER BY Plot_ID"
    Me!Plot.Requery
  End If
  
End Sub

Private Sub Plot_GotFocus()
  If IsNull(Me!Park_Code) Then
    MsgBox "You must select a park first."
    Me!Park_Code.SetFocus
  End If
End Sub
