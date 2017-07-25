Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
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
    Width =4440
    DatasheetFontHeight =9
    ItemSuffix =15
    Left =6630
    Top =2325
    Right =11070
    Bottom =5970
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
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =960
                    Top =1200
                    Width =1035
                    Height =240
                    FontWeight =700
                    Name ="lblStart_Date"
                    Caption ="Visit Date"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =960
                    LayoutCachedTop =1200
                    LayoutCachedWidth =1995
                    LayoutCachedHeight =1440
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =120
                    Top =480
                    Width =540
                    Height =240
                    FontWeight =700
                    Name ="lblUnit_Code"
                    Caption ="Park"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =120
                    LayoutCachedTop =480
                    LayoutCachedWidth =660
                    LayoutCachedHeight =720
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =120
                    Top =780
                    Width =600
                    Height =240
                    FontWeight =700
                    Name ="lblPlot_ID"
                    Caption ="Route"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =120
                    LayoutCachedTop =780
                    LayoutCachedWidth =720
                    LayoutCachedHeight =1020
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =840
                    Top =480
                    Width =600
                    Height =255
                    ColumnWidth =540
                    ColumnOrder =0
                    Name ="Unit_Code"
                    ControlSource ="Unit_Code"
                    StatusBarText ="Park Code."

                    LayoutCachedLeft =840
                    LayoutCachedTop =480
                    LayoutCachedWidth =1440
                    LayoutCachedHeight =735
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
                    Left =840
                    Top =780
                    Width =3540
                    Height =255
                    ColumnWidth =600
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Plot_ID"
                    ControlSource ="Plot_ID"
                    StatusBarText ="Plot identifier"

                    LayoutCachedLeft =840
                    LayoutCachedTop =780
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =85
                    Left =660
                    Top =60
                    Width =2640
                    Height =360
                    FontSize =12
                    FontWeight =700
                    Name ="lblTitle"
                    Caption ="Select a Visit to Edit"
                    LayoutCachedLeft =660
                    LayoutCachedTop =60
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =420
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3660
                    Top =60
                    Width =720
                    Height =294
                    FontSize =9
                    FontWeight =500
                    TabIndex =2
                    Name ="btnClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Close the data entry form"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =3660
                    LayoutCachedTop =60
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =354
                    ForeThemeColorIndex =0
                    UseTheme =1
                    BackColor =15921906
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =0
                    HoverColor =9434577
                    PressedColor =15921906
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeColor =16711680
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =22
                    QuickStyleMask =-53
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =2
                    WebImagePaddingRight =4
                    WebImagePaddingBottom =7
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
                    Left =840
                    Top =60
                    Width =660
                    Height =255
                    ColumnWidth =2310
                    Name ="Event_ID"
                    ControlSource ="Event_ID"
                    StatusBarText ="M. Event identifier (Event_ID)"

                    LayoutCachedLeft =840
                    LayoutCachedTop =60
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =315
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
                    Left =1620
                    Top =60
                    Width =690
                    Height =255
                    ColumnWidth =2310
                    TabIndex =1
                    Name ="Location_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="M. Link to tbl_Locations (Loc_ID)"

                    LayoutCachedLeft =1620
                    LayoutCachedTop =60
                    LayoutCachedWidth =2310
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    BackStyle =0
                    IMESentenceMode =3
                    Left =840
                    Top =60
                    Width =1035
                    Height =255
                    ColumnWidth =1035
                    TabIndex =2
                    Name ="Start_Date"
                    ControlSource ="Start_Date"
                    Format ="Short Date"
                    StatusBarText ="M. Starting date for the event (Start_Date)"

                    LayoutCachedLeft =840
                    LayoutCachedTop =60
                    LayoutCachedWidth =1875
                    LayoutCachedHeight =315
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2520
                    Top =30
                    Width =1019
                    Height =300
                    TabIndex =3
                    Name ="btnEdit"
                    Caption ="Edit Visit"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =2520
                    LayoutCachedTop =30
                    LayoutCachedWidth =3539
                    LayoutCachedHeight =330
                    ForeThemeColorIndex =0
                    UseTheme =1
                    BackColor =15921906
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =0
                    HoverColor =9434577
                    PressedColor =15921906
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeColor =16711680
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    QuickStyle =22
                    QuickStyleMask =-53
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =2
                    WebImagePaddingRight =4
                    WebImagePaddingBottom =7
                End
            End
        End
        Begin FormFooter
            Height =60
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
' Version:      1.03
' Basis:        -
'
' Description:  Visit Date form object related properties, functions & procedures for UI display
'
' Source/date:  NCPN, Unknown - for NCPN tools
' References:   -
' Revisions:    NCPN - Unknown  - 1.00 - initial version
'               BLC - 7/12/2017 - 1.01 - added documentation, error handling, update usys_temp_transect
'                                        before opening form
'               BLC - 7/14/2017 - 1.02 - renamed buttons
'               BLC - 7/18/2017 - 1.03 - revised to refresh temp tables (usys_temp_*)
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
'   BLC - 7/14/2017 - remove hovers
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler
    
  
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
'   BLC - 7/18/2017 - update of usys_temp_speciescover data table before opening form,
'                     use RefreshTempTable instead
' ---------------------------------
Private Sub btnEdit_Click()
    On Error GoTo Err_Handler

    Dim strCriteriaLoc As String
    Dim strCriteriaEvent As String

    strCriteriaLoc = GetCriteriaString("[Location_ID]=", "tbl_Locations", "Location_ID", Me.Name, "Location_ID")
    strCriteriaEvent = GetCriteriaString("[Event_ID]=", "tbl_Events", "Event_ID", Me.Name, "Event_ID")
    
    're-generate the temp table sources
    RefreshTempTable "usys_temp_transect"
    RefreshTempTable "usys_temp_speciescover"

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
