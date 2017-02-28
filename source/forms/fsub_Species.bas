Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7800
    DatasheetFontHeight =9
    ItemSuffix =47
    Left =-135
    Top =1155
    Right =6840
    Bottom =7560
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x43f03470521ee340
    End
    RecordSource ="tbl_Quadrat_Species"
    Caption ="fsub_Species"
    BeforeInsert ="[Event Procedure]"
    DatasheetFontName ="Arial"
    FilterOnLoad =255
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontWeight =700
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =540
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =2160
                    Width =2700
                    Height =240
                    Name ="Nested_Quad_Label"
                    Caption ="% Cover in Classes"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4920
                    Top =60
                    Width =840
                    Height =420
                    Name ="Percent_Cover_Label"
                    Caption ="AverageCover"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =60
                    Top =240
                    Width =960
                    Height =240
                    Name ="Label14"
                    Caption ="Species"
                End
                Begin Label
                    OverlapFlags =95
                    Left =2160
                    Top =240
                    Width =900
                    Height =240
                    Name ="Label23"
                    Caption ="Q1@0m"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =2940
                    Top =240
                    Width =915
                    Height =240
                    Name ="Label25"
                    Caption ="Q2@4.5m"
                    LayoutCachedLeft =2940
                    LayoutCachedTop =240
                    LayoutCachedWidth =3855
                    LayoutCachedHeight =480
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =3945
                    Top =240
                    Width =915
                    Height =240
                    Name ="Label26"
                    Caption ="Q3@9.5m"
                    LayoutCachedLeft =3945
                    LayoutCachedTop =240
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =480
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5820
                    Top =60
                    Width =840
                    Height =420
                    Name ="lblIsDead"
                    Caption ="Dead or Alive?"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5820
                    LayoutCachedTop =60
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =480
                End
            End
        End
        Begin Section
            Height =420
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =420
                    Height =255
                    ColumnWidth =2310
                    Name ="Species_ID"
                    ControlSource ="Species_ID"
                    StatusBarText ="Unique record identifier - primary key"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =600
                    Top =60
                    Width =360
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Transect_ID"
                    ControlSource ="Transect_ID"
                    StatusBarText ="Foreign key to tbl_Quadrat_Transect"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =1
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5040
                    Top =60
                    Width =540
                    Height =255
                    ColumnWidth =465
                    TabIndex =6
                    Name ="Average_Cover"
                    ControlSource ="Average_Cover"
                    Format ="General Number"
                    StatusBarText ="Percent cover in 10 m2 quadrat"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6960
                    Top =60
                    Width =705
                    Height =300
                    TabIndex =7
                    ForeColor =255
                    Name ="ButtonDelete"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6960
                    LayoutCachedTop =60
                    LayoutCachedWidth =7665
                    LayoutCachedHeight =360
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    ListRows =21
                    Left =2160
                    Top =60
                    Width =900
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Q1_hm"
                    ControlSource ="Q1_hm"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Percent cover Q1 @ 3m"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    ListRows =21
                    Left =3060
                    Top =60
                    Width =900
                    TabIndex =4
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Q2_5m"
                    ControlSource ="Q2_5m"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Percent cover Q2 @ 8m"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    ListRows =21
                    Left =3960
                    Top =60
                    Width =900
                    TabIndex =5
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    Name ="Q3_10m"
                    ControlSource ="Q3_10m"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Percent cover Q3 @ 13m"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"

                End
                Begin ComboBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4320
                    Left =60
                    Top =60
                    Width =1860
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"50\""
                    Name ="Plant_Code"
                    ControlSource ="Plant_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qry_sel_Species_Lookup.Master_PLANT_Code, qry_sel_Species_Lookup.LU_Code,"
                        " qry_sel_Species_Lookup.Utah_Species FROM qry_sel_Species_Lookup; "
                    ColumnWidths ="0;1728;2592"
                    BeforeUpdate ="[Event Procedure]"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =21
                    Left =5760
                    Top =60
                    Width =900
                    TabIndex =8
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"510\""
                    ConditionalFormat = Begin
                        0x010000008a000000020000000000000002000000000000000300000001000000 ,
                        0x00000000fff20000010000000000000004000000140000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x22002200000000005b006300620078004900730044006500610064005d003c00 ,
                        0x3e002200220000000000
                    End
                    Name ="cbxIsDead"
                    ControlSource ="IsDead"
                    RowSourceType ="Table/Query"
                    RowSource ="qry_IsDead_Plus_Flags"
                    ColumnWidths ="0;1440"
                    ControlTipText ="Indicate if species is alive or dead (or the appropriate missing data flag)"
                    LayoutCachedLeft =5760
                    LayoutCachedTop =60
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =300
                    ConditionalFormat14 = Begin
                        0x01000200000000000000020000000100000000000000fff20000020000002200 ,
                        0x2200000000000000000000000000000000000000000000010000000000000001 ,
                        0x00000000000000ffffff000f0000005b00630062007800490073004400650061 ,
                        0x0064005d003c003e002200220000000000000000000000000000000000000000 ,
                        0x0000
                    End
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
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public Function CalcAvgCover() As Single
' Calculate average cover for a species - 9/22/2010 - Russ DenBleyker
' Northern Colorado Plateau Network
    On Error GoTo Err_Handler
    
    Dim AvgCover As Single
    Dim TotCover As Single
   
    AvgCover = 0
    TotCover = 0
    If Not IsNull(Me!Q1_hm) Or Not IsNull(Me!Q2_5m) Or Not IsNull(Me!Q3_10m) Then
      If Not IsNull(Me!Q1_hm) Then
        TotCover = Me!Q1_hm
      End If
      If Not IsNull(Me!Q2_5m) Then
        TotCover = TotCover + Me!Q2_5m
      End If
      If Not IsNull(Me!Q3_10m) Then
        TotCover = TotCover + Me!Q3_10m
      End If
      AvgCover = TotCover / 3
    End If
    CalcAvgCover = AvgCover
Exit_Procedure_1M:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (CalcAvgCover)"
            Resume Exit_Procedure_1M
    End Select

End Function
Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo Err_Handler

    If IsNull(Me.Parent!Observer) Then
      MsgBox "You must enter Observer first."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
      GoTo Exit_Procedure
    End If
    ' Create the GUID primary key value
    If IsNull(Me!Species_ID) Then
        If GetDataType("tbl_Quadrat_Species", "species_ID") = dbText Then
            Me.Species_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub ButtonDelete_Click()
On Error GoTo Err_ButtonDelete_Click
  Dim Reply As Integer
  Reply = MsgBox("Are you sure you want to delete this record?", vbYesNo, "Species Delete")
  If Reply = 6 Then
    DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70
  End If
Exit_ButtonDelete_Click:
    Exit Sub

Err_ButtonDelete_Click:
    MsgBox Err.Description
    Resume Exit_ButtonDelete_Click
    
End Sub

Private Sub Plant_Code_BeforeUpdate(Cancel As Integer)
  If Not IsNull(Me!Plant_Code) Then
    If Not IsNull(DLookup("[Species_ID]", "tbl_Quadrat_Species", "[Transect_ID] = '" & Me!Transect_ID & "' AND [Plant_Code] = '" & Me!Plant_Code & "'")) Then
      MsgBox "Duplicate species for this quadrat."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
    End If
  End If
End Sub

Private Sub Q1_hm_AfterUpdate()
  Me!Average_Cover = CalcAvgCover
End Sub

Private Sub Q1_hm_BeforeUpdate(Cancel As Integer)
  If IsNull(Me!Plant_Code) Then
      MsgBox "You must enter species first."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
  End If
End Sub

Private Sub Q2_5m_AfterUpdate()
  Me!Average_Cover = CalcAvgCover
End Sub

Private Sub Q2_5m_BeforeUpdate(Cancel As Integer)
  If IsNull(Me!Plant_Code) Then
      MsgBox "You must enter species first."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
  End If
End Sub

Private Sub Q3_10m_AfterUpdate()
  Me!Average_Cover = CalcAvgCover
End Sub


Private Sub Q3_10m_BeforeUpdate(Cancel As Integer)
  If IsNull(Me!Plant_Code) Then
      MsgBox "You must enter species first."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
  End If
End Sub
