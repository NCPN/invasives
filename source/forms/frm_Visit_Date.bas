Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    TabularFamily =127
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    DatasheetFontHeight =9
    ItemSuffix =15
    Left =4965
    Top =5160
    Right =12165
    Bottom =8805
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x5bd611c7ad13e340
    End
    RecordSource ="qfrm_Visit_Date"
    Caption ="Select a Visit"
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
            Height =1140
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =180
                    Top =900
                    Width =1035
                    Height =240
                    FontWeight =700
                    Name ="Start_Date_Label"
                    Caption ="Visit Date"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =120
                    Top =120
                    Width =540
                    Height =240
                    FontWeight =700
                    Name ="Unit_Code_Label"
                    Caption ="Park"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =120
                    Top =540
                    Width =600
                    Height =240
                    FontWeight =700
                    Name ="Plot_ID_Label"
                    Caption ="Route"
                    Tag ="DetachedLabel"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6000
                    Top =120
                    Width =1020
                    Height =300
                    Name ="btnClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6000
                    LayoutCachedTop =120
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =420
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =720
                    Top =120
                    Width =600
                    Height =255
                    ColumnWidth =540
                    ColumnOrder =0
                    FontWeight =700
                    TabIndex =1
                    Name ="Unit_Code"
                    ControlSource ="Unit_Code"
                    StatusBarText ="Park Code."

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =780
                    Top =540
                    Width =3540
                    Height =255
                    ColumnWidth =600
                    ColumnOrder =1
                    FontWeight =700
                    TabIndex =2
                    Name ="Plot_ID"
                    ControlSource ="Plot_ID"
                    StatusBarText ="Plot identifier"

                End
                Begin Label
                    OverlapFlags =85
                    Left =2760
                    Top =120
                    Width =2640
                    Height =360
                    FontSize =12
                    FontWeight =700
                    Name ="Label12"
                    Caption ="Select a Visit to Edit"
                End
            End
        End
        Begin Section
            Height =360
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =660
                    Height =255
                    ColumnWidth =2310
                    Name ="Event_ID"
                    ControlSource ="Event_ID"
                    StatusBarText ="M. Event identifier (Event_ID)"

                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =840
                    Top =60
                    Width =690
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Location_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="M. Link to tbl_Locations (Loc_ID)"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    BackStyle =0
                    IMESentenceMode =3
                    Left =180
                    Top =60
                    Width =1035
                    Height =255
                    ColumnWidth =1035
                    TabIndex =2
                    Name ="Start_Date"
                    ControlSource ="Start_Date"
                    Format ="Short Date"
                    StatusBarText ="M. Starting date for the event (Start_Date)"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1860
                    Top =60
                    Width =1020
                    Height =300
                    TabIndex =3
                    Name ="btnEdit"
                    Caption ="Edit Visit"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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
' Form:         frm_Visit_Date
' Level:        Application form
' Version:      1.05
' Basis:        -
'
' Description:  Visit Date form object related properties, functions & procedures for UI display
'
' Source/date:  NCPN, Unknown - for NCPN tools
' References:   -
' Revisions:    NCPN - Unknown  - 1.00 - initial version
'               BLC - 7/12/2017 - 1.01 - added documentation, error handling, update usys_temp_transect
'                                        before opening form
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------

'---------------------
' Event Declarations
'---------------------

'---------------------
' Properties
'---------------------

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
'   BLC - 7/12/2017 - added documentation, error handling
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler
    
    'set hover colors
    Me.btnEdit.HoverColor = lngGreen
    Me.btnClose.HoverColor = lngGreen
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[frm_Visit_Date form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnEdit_Click
' Description:  Edit button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 7/12/2017 - added documentation, error handling,
'                     update of usys_temp_transect data table before opening form
' ---------------------------------
Private Sub btnEdit_Click()
    On Error GoTo Err_Handler

    Dim strCriteriaLoc As String
    Dim strCriteriaEvent As String

    strCriteriaLoc = GetCriteriaString("[Location_ID]=", "tbl_Locations", "Location_ID", Me.name, "Location_ID")
    strCriteriaEvent = GetCriteriaString("[Event_ID]=", "tbl_Events", "Event_ID", Me.name, "Event_ID")
    
    're-generate the temp table source
    DoCmd.SetWarnings False
    If TableExists("usys_temp_transect") Then
        DoCmd.DeleteObject acTable, "usys_temp_transect"
    End If
    DoCmd.OpenQuery "Create_usys_temp_transect"
    
    'move tables to Queries - Application group
    SetNavGroup "Queries - Application", "usys_temp_transect", "table"

    DoCmd.SetWarnings True
    
    ' Filter by location and event
    DoCmd.OpenForm "frm_Data_Entry", , , strCriteriaLoc & " AND " & strCriteriaEvent, , , strCriteriaEvent
    DoCmd.Close acForm, "frm_Visit_Date"
    DoCmd.SelectObject acForm, "frm_Data_Entry"

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnEdit_Click[frm_Visit_Date form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnClose_Click
' Description:  Close button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 7/12/2017 - added documentation, error handling
' ---------------------------------
Private Sub btnClose_Click()
On Error GoTo Err_Handler

    DoCmd.Close

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnClose_Click[frm_Visit_Date form])"
    End Select
    Resume Exit_Handler
End Sub
