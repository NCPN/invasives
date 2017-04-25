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
    Left =2580
    Top =2235
    Right =11130
    Bottom =8640
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x43f03470521ee340
    End
    RecordSource ="tbl_Quadrat_Species"
    Caption ="fsub_Species"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
                    Name ="btnDelete"
                    Caption ="Delete"

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
Option Explicit

' =================================
' Form:         fsub_Species
' Level:        Application form
' Version:      1.02
' Basis:        -
'
' Description:  Species subform object related properties, functions & procedures for UI display
'
' Source/date:  Russ DenBleyker, Unknown - for NCPN tools
' References:   -
' Revisions:    RDB - Unknown - 1.00 - initial version
'               BLC - 3/8/2017 - 1.01 - added documentation, error handling
'               BLC - 4/21/2017 - 1.02 - added HasRecords, ParentForm properties
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
'Private WithEvents m_ParentForm As Form 'Form_frm_Quadrat_Transect
Private m_HasRecords As Boolean
Private m_HasRecordsQ1 As Boolean
Private m_HasRecordsQ2 As Boolean
Private m_HasRecordsQ3 As Boolean

'---------------------
' Event Declarations
'---------------------
Public Event InvalidHasRecords(Value As Boolean)
Public Event InvalidHasRecordsQ1(Value As Boolean)
Public Event InvalidHasRecordsQ2(Value As Boolean)
Public Event InvalidHasRecordsQ3(Value As Boolean)

'---------------------
' Properties
'---------------------
'Public Property Let ParentForm(value As Form) '_frm_Quadrat_Transect)
'    m_ParentForm = value
'End Property
'
'Public Property Get ParentForm() As Form '_frm_Quadrat_Transect
'    Set ParentForm = m_ParentForm
'End Property

Public Property Let HasRecords(Value As Boolean)
    If varType(Value) = vbBoolean Then
        m_HasRecords = Value
    Else
        RaiseEvent InvalidHasRecords(Value)
    End If
End Property

Public Property Get HasRecords() As Boolean
    HasRecords = m_HasRecords
End Property

Public Property Let HasRecordsQ1(Value As Boolean)
    If varType(Value) = vbBoolean Then
        m_HasRecordsQ1 = Value
    Else
        RaiseEvent InvalidHasRecordsQ1(Value)
    End If
End Property

Public Property Get HasRecordsQ1() As Boolean
    HasRecordsQ1 = m_HasRecordsQ1
End Property

Public Property Let HasRecordsQ2(Value As Boolean)
    If varType(Value) = vbBoolean Then
        m_HasRecordsQ2 = Value
    Else
        RaiseEvent InvalidHasRecordsQ2(Value)
    End If
End Property

Public Property Get HasRecordsQ2() As Boolean
    HasRecordsQ2 = m_HasRecordsQ2
End Property

Public Property Let HasRecordsQ3(Value As Boolean)
    If varType(Value) = vbBoolean Then
        m_HasRecordsQ3 = Value
    Else
        RaiseEvent InvalidHasRecordsQ3(Value)
    End If
End Property

Public Property Get HasRecordsQ3() As Boolean
    HasRecordsQ3 = m_HasRecordsQ3
End Property


'---------------------
' Methods
'---------------------

' ---------------------------------
' Sub:          Form_Open
' Description:  form opening actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 3/8/2017 - added documentation, error handling
'   BLC - 4/21/2017 - added setting HasRecordsQ1-3 properties
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler
    
    'defaults
    Me.HasRecords = False
    Me.HasRecordsQ1 = False
    Me.HasRecordsQ2 = False
    Me.HasRecordsQ3 = False

    'determine if Q1-3 have records
    If Me.Form.Recordset.RecordCount > 0 And Not IsNull(Me.Plant_Code) Then Me.HasRecords = True

'    'determine if any Q1-3 has values
'    If Not IsNull(Me.Q1_hm) Then
'        Debug.Print "Q1_hm: " & Q1_hm
'        HasRecordsQ1 = True
'    End If
'
'    If Not IsNull(Me.Q2_8m) Then
'        Debug.Print "Q2_8m: " & Q2_8m
'        HasRecordsQ2 = True
'    End If
'
'    If Not IsNull(Me.Q3_13m) Then
'        Debug.Print "Q3_13m: " & Q3_13m
'        HasRecordsQ3 = True
'    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[fsub_Species form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Current
' Description:  form current actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 21, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/21/2017 - initial version
' ---------------------------------
Private Sub Form_Current()
    On Error GoTo Err_Handler

    'defaults
    HasRecords = False
    HasRecordsQ1 = False
    HasRecordsQ2 = False
    HasRecordsQ3 = False
    
    If Me.Form.Recordset.RecordCount > 0 And Not IsNull(Me.Plant_Code) Then _
        Me.HasRecords = True
    
    'determine if any Q1-3 has values
    Debug.Print "Q1_hm: " & Q1_hm
    If Not IsNull(Me.Q1_hm) Then
        HasRecordsQ1 = True
    End If

    Debug.Print "Q2_8m: " & Q2_8m
    If Not IsNull(Me.Q2_8m) Then
        HasRecordsQ2 = True
    End If

    Debug.Print "Q3_13m: " & Q3_13m
    If Not IsNull(Me.Q3_13m) Then
        HasRecordsQ3 = True
    End If
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[fsub_Species form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_BeforeInsert
' Description:  form before insert actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 3/8/2017 - added documentation, error handling
' ---------------------------------
Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo Err_Handler

    If IsNull(Me.Parent!Observer) Then
      MsgBox "You must enter Observer first."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
      GoTo Exit_Handler
    End If
    ' Create the GUID primary key value
    If IsNull(Me!Species_ID) Then
        If GetDataType("tbl_Quadrat_Species", "species_ID") = dbText Then
            Me.Species_ID = fxnGUIDGen
        End If
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeInsert[fsub_Species form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          ParentForm_Current
' Description:  Parent form current actions
' Assumptions:  parent form is frm_Quadrat_Transect
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Harvey French, July 31, 2015
'   http://stackoverflow.com/questions/31611912/how-best-to-call-a-public-sub-routine-declared-in-a-form-used-as-the-source-obje
' Source/date:  Bonnie Campbell, April 21, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/21/2017 - initial version
' ---------------------------------
Private Sub ParentForm_Current(Cancel As Integer)
    On Error GoTo Err_Handler

'respond to parent form current action


Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ParentForm_Current[fsub_Species form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnDelete_Click
' Description:  Delete button actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 3/8/2017 - added documentation, error handling
' ---------------------------------
Private Sub btnDelete_Click()
On Error GoTo Err_Handler

  Dim Reply As Integer
  Reply = MsgBox("Are you sure you want to delete this record?", vbYesNo, "Species Delete")
  If Reply = 6 Then
    DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
    DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70
  End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnDelete_Click[fsub_Species form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Plant_Code_BeforeUdpate
' Description:  Plant_Code combobox actions before update
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 3/8/2017 - added documentation, error handling
' ---------------------------------
Private Sub Plant_Code_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

  If Not IsNull(Me!Plant_Code) Then
    If Not IsNull(DLookup("[Species_ID]", "tbl_Quadrat_Species", "[Transect_ID] = '" & Me!Transect_ID & "' AND [Plant_Code] = '" & Me!Plant_Code & "'")) Then
      MsgBox "Duplicate species for this quadrat."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
    End If
  End If
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Plant_Code_BeforeUpdate[fsub_Species form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Q1_hm_BeforeUpdate
' Description:  Q1_hm combobox actions before update
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 3/8/2017 - added documentation, error handling
' ---------------------------------
Private Sub Q1_hm_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

  If IsNull(Me!Plant_Code) Then
      MsgBox "You must enter species first."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
  End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Q1_hm_BeforeUpdate[fsub_Species form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Q2_5m_BeforeUpdate
' Description:  Q2_5m combobox actions before update
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 3/8/2017 - added documentation, error handling
' ---------------------------------
Private Sub Q2_5m_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

  If IsNull(Me!Plant_Code) Then
      MsgBox "You must enter species first."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
  End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Q2_5m_BeforeUpdate[fsub_Species form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Q3_10m_BeforeUpdate
' Description:  Q3_10m combobox actions before update
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 3/8/2017 - added documentation, error handling
' ---------------------------------
Private Sub Q3_10m_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

  If IsNull(Me!Plant_Code) Then
      MsgBox "You must enter species first."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
  End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Q3_10m_BeforeUpdate[fsub_Species form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Q1_hm_AfterUpdate
' Description:  Q1_hm combobox actions after update
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 3/8/2017 - added documentation, error handling
' ---------------------------------
Private Sub Q1_hm_AfterUpdate()
On Error GoTo Err_Handler

  Me!Average_Cover = CalcAvgCover

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Q1_hm_AfterUpdate[fsub_Species form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Q2_5m_AfterUpdate
' Description:  Q2_5m combobox actions after update
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 3/8/2017 - added documentation, error handling
' ---------------------------------
Private Sub Q2_5m_AfterUpdate()
On Error GoTo Err_Handler

  Me!Average_Cover = CalcAvgCover

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Q2_5m_AfterUpdate[fsub_Species form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Q3_10m_AfterUpdate
' Description:  Q3_10m combobox actions after update
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 3/8/2017 - added documentation, error handling
' ---------------------------------
Private Sub Q3_10m_AfterUpdate()
On Error GoTo Err_Handler

  Me!Average_Cover = CalcAvgCover

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Q3_10m_AfterUpdate[fsub_Species form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          CalcAvgCover
' Description:  Calculate the average cover
' Assumptions:  Percent cover controls for quadrats that are
'               not sampled are DISABLED in the user interface
'               if control enabled  --> include it in average
'                          disabled --> exclude from average
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Russ DenBleyker, 9/22/2010 - for Northern Colorado Plateau Network
' Adapted:      -
' Revisions:
'   RDB - 9/22/2010 - initial version
'   BLC - 3/8/2017  - added documentation, error handling,
'                     revised to address non-sampled quadrats
' ---------------------------------
Public Function CalcAvgCover() As Single
    On Error GoTo Err_Handler
    
    Dim AvgCover As Single
    Dim TotalCover As Single
    Dim Count As Integer, i As Integer
    Dim strControl As String, strPosition As String
   
    Count = 0
    AvgCover = 0
    TotalCover = 0
    
'    If Not IsNull(Me!Q1_hm) Or Not IsNull(Me!Q2_5m) Or Not IsNull(Me!Q3_10m) Then
'      If Not IsNull(Me!Q1_hm) Then
'        TotCover = Me!Q1_hm
'      End If
'      If Not IsNull(Me!Q2_5m) Then
'        TotCover = TotCover + Me!Q2_5m
'      End If
'      If Not IsNull(Me!Q3_10m) Then
'        TotCover = TotCover + Me!Q3_10m
'      End If
'      AvgCover = TotCover / 3
'    End If
    
    
    '---------------------------------------------------
        '---------------------------------------------------
    For i = 1 To 3
        'determine quadrat control
        Select Case i
            Case 1
                strPosition = "h"
            Case 2
                strPosition = "5"
            Case 3
                strPosition = "10"
        End Select
    
        strControl = "Q" & i & "_" & strPosition & "m"
    
        If Me.Controls(strControl).Enabled Then
            If Not IsNull(Me.Controls(strControl)) Then
                TotalCover = TotalCover + Me.Controls(strControl)
                Count = Count + 1
            End If
        End If
    Next
    
    If Count > 0 Then
        'calculate the average
        AvgCover = TotalCover / Count
    End If

    CalcAvgCover = AvgCover

Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CalcAvgCover[fsub_Species form])"
    End Select
    Resume Exit_Handler
End Function
