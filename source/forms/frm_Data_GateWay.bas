Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    ScrollBars =2
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10560
    DatasheetFontHeight =10
    ItemSuffix =37
    Left =1785
    Top =195
    Right =12345
    Bottom =6045
    DatasheetGridlinesColor =12632256
    Filter ="Unit_code = 'FOBU'"
    OrderBy ="Plot_ID DESC, Unit_Code"
    RecSrcDt = Begin
        0x29b5dcdf75fbe240
    End
    RecordSource ="qfrm_Data_Gateway"
    Caption ="Data Gateway"
    OnOpen ="[Event Procedure]"
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
        Begin FormHeader
            Height =1248
            BackColor =11056034
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5040
                    Top =1020
                    Width =1680
                    Height =228
                    Name ="labUpdated_Date"
                    Caption ="Entered/updated*"
                    FontName ="Arial"
                    OnDblClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =105
                    Top =1020
                    Width =795
                    Height =225
                    Name ="labUnit_code"
                    Caption ="Unit*"
                    FontName ="Arial"
                    OnDblClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9720
                    Top =120
                    Width =720
                    Height =354
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    Name ="cmdClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Close the data entry form"
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =120
                    Width =4860
                    Height =408
                    BackColor =16777215
                    ForeColor =0
                    Name ="labOverview"
                    Caption ="* Double-click on the field label to change sort order.  Double-click on a Plot "
                        "ID to open the Site form for that record."
                    FontName ="Arial"
                    ControlTipText ="View mode"
                End
                Begin ComboBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =660
                    Width =960
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"10\";\"8\""
                    Name ="selPark"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT tbl_Locations.Unit_Code FROM tbl_Locations ORDER BY tbl_Location"
                        "s.Unit_Code; "
                    StatusBarText ="Park code"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =120
                            Top =660
                            Width =1320
                            Height =228
                            Name ="labPark"
                            Caption ="Filter by:  Park"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ToggleButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =2760
                    Top =660
                    Width =480
                    Height =300
                    Name ="togFilterByPark"
                    AfterUpdate ="[Event Procedure]"
                    Caption ="Filter on"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Turn the park filter on or off"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1005
                    Top =1020
                    Width =675
                    Height =225
                    Name ="labPlot_ID"
                    Caption ="Route*"
                    FontName ="Arial"
                    OnDblClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7800
                    Top =120
                    Width =1680
                    FontSize =9
                    FontWeight =700
                    TabIndex =4
                    Name ="ButtonNewSite"
                    Caption ="Add a new Route"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6600
                    Top =120
                    Width =960
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    Name ="buttonRefresh"
                    Caption ="Refresh"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                End
            End
        End
        Begin Section
            Height =420
            BackColor =11056034
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5220
                    Top =60
                    Width =1680
                    ColumnWidth =1710
                    TabIndex =1
                    Name ="txtUpdated"
                    ControlSource ="Updated_Date"
                    Format ="yyyy mmm dd hh:nn"
                    StatusBarText ="Date on which data entry occurred"
                    FontName ="Arial"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =780
                    ColumnWidth =2310
                    Name ="txtUnit_code"
                    ControlSource ="Unit_Code"
                    StatusBarText ="Unit code"
                    FontName ="Arial"
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1920
                    Top =60
                    Width =420
                    TabIndex =2
                    Name ="txtLocation_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="Name of the location"
                    FontName ="Arial"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =1080
                    Top =60
                    Width =3840
                    TabIndex =3
                    Name ="txtPlot_ID"
                    ControlSource ="Plot_ID"
                    StatusBarText ="Plot identifier"
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =7200
                    Top =60
                    Width =1320
                    Height =300
                    TabIndex =4
                    Name ="ButtonVisitList"
                    Caption ="View Visits"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =7200
                    Top =60
                    Width =1319
                    Height =299
                    TabIndex =5
                    Name ="ButtonNewVisit"
                    Caption ="Add New Visit"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8700
                    Top =60
                    Width =1650
                    Height =300
                    TabIndex =6
                    Name ="ButtonSiteChar"
                    Caption ="Edit/Add Routes"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
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
' FORM NAME:    frm_Data_Gateway
' Description:
' Data source:  qfrm_Data_Gateway
' Data access:  view and delete records (delete by cmdDeleteRec)
' Pages:        none
' Functions:    fxnSortRecords
' References:   none
' Source/date:  John R. Boetsch, June 7, 2006
' Revisions:    Simon Kingston, Sept. 2006 - added CorrectText calls where strings were being used in criteria
'                                          - updated cmdDeleteRec_Click() event to use appropriate criteria depending on primary key
' =================================

Dim strSortField As String    ' Keeps track of current sort settings
Dim strSortOrder As String
Dim strSortFieldLabel As String

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    Dim varReturn As Variant

    ' On opening the form, set the initial sort order
    strSortFieldLabel = "labPlot_ID"
    varReturn = fxnSortRecords("Unit_Code", "Plot_ID")
    ' Set the filter
    If fxnSwitchboardIsOpen Then
        Me.selPark = Forms!frm_Switchboard.cPark
        Me.Filter = "Unit_code = " & CorrectText(Me.selPark)
        Me.FilterOn = True
        Me.labPark.FontBold = True
        Me.togFilterByPark = True
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub labPlot_ID_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    fxnSortRecords ("Plot_ID")

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub



Private Sub selPark_AfterUpdate()
    On Error GoTo Err_Handler

    Me.Filter = "Unit_code = " & CorrectText(Me.selPark)
    If togFilterByPark Then
      Me.Filter = "Unit_code = " & CorrectText(Me.selPark)
      Me.FilterOn = True
      Me.labPark.FontBold = True
    End If
Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub togFilterByPark_AfterUpdate()
    On Error GoTo Err_Handler

    If Me.ActiveControl Then
      If Not IsNull(Me!selPark) Then
        Me.Filter = "Unit_code = " & CorrectText(Me.selPark)
        Me.FilterOn = True
        Me.labPark.FontBold = True
      End If
    Else
        Me.FilterOn = False
        Me.labPark.FontBold = False
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmdClose_Click()
    On Error GoTo Err_Handler

    DoCmd.Close , , acSaveNo

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' =================================
' The next several procedures re-sort the records if the user
'   double-clicks on a field label


Private Sub labUnit_code_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    fxnSortRecords ("Unit_code")

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub labUpdated_Date_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    fxnSortRecords ("Updated_Date")

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' =================================
' FUNCTION:     fxnSortRecords
' Description:  Sorts the records by the indicated field
' Parameters:   strFieldName
' Returns:      none
' Throws:       none
' References:   strFieldName, strSortOrder, strSortFieldLabel
'               (form-level variables)
' Source/date:  John R. Boetsch, May 5, 2006
' Revisions:    <name, date, desc - add lines as you go>
' =================================

Private Function fxnSortRecords(ByVal strFieldName As String, _
    Optional ByVal strField2Name As String)
    On Error GoTo Err_Handler

    Dim strOrderBy As String

    ' If already sorting in ascending order by this field, sort descending
    If strFieldName = strSortField And strSortOrder = "" Then
        strSortOrder = " DESC"
    Else: strSortOrder = ""
    End If
    ' Create the order by string and activate the filter
    strOrderBy = strFieldName & strSortOrder
    If strField2Name <> "" Then
        strOrderBy = strField2Name & " DESC, " & strOrderBy
    End If
    strSortField = strFieldName
    Me.Form.OrderBy = strOrderBy
    Me.Form.OrderByOn = True

    ' Change the label format to indicate the sorted field
    Me.Controls.item(strSortFieldLabel).FontItalic = False
    Me.Controls.item(strSortFieldLabel).FontBold = False
    strSortFieldLabel = "lab" & strFieldName
    Me.Controls.item(strSortFieldLabel).FontItalic = True
    Me.Controls.item(strSortFieldLabel).FontBold = True

Exit_Procedure:
    Exit Function

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
        "Error encountered (fxnSortRecords)"
    Resume Exit_Procedure

End Function

Private Sub ButtonNewSite_Click()
On Error GoTo Err_ButtonNewSite_Click

    DoCmd.OpenForm "frm_Locations", , , , acFormAdd, , "New record"

Exit_ButtonNewSite_Click:
    Exit Sub

Err_ButtonNewSite_Click:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_ButtonNewSite_Click
    
End Sub

Private Sub ButtonVisitList_Click()
On Error GoTo Err_ButtonVisitList_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Visit_Date"
    
    stLinkCriteria = "[Location_ID]=" & "'" & Me![txtLocation_ID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonVisitList_Click:
    Exit Sub

Err_ButtonVisitList_Click:
    MsgBox Err.Description
    Resume Exit_ButtonVisitList_Click
    
End Sub
Private Sub ButtonRefresh_Click()
On Error GoTo Err_ButtonRefresh_Click

    Me.Requery

Exit_ButtonRefresh_Click:
    Exit Sub

Err_ButtonRefresh_Click:
    MsgBox Err.Description
    Resume Exit_ButtonRefresh_Click
    
End Sub
Private Sub ButtonNewVisit_Click()
On Error GoTo Err_ButtonNewVisit_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Plot_Revisit"
    
    stLinkCriteria = "[Location_ID]=" & "'" & Me![txtLocation_ID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
'    DoCmd.Close acForm, "frm_Data_Gateway"

Exit_ButtonNewVisit_Click:
    Exit Sub

Err_ButtonNewVisit_Click:
    MsgBox Err.Description
    Resume Exit_ButtonNewVisit_Click
    
End Sub
Private Sub ButtonSiteChar_Click()
On Error GoTo Err_ButtonSiteChar_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Locations"
    
    stLinkCriteria = "[Location_ID]=" & "'" & Me![txtLocation_ID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonSiteChar_Click:
    Exit Sub

Err_ButtonSiteChar_Click:
    MsgBox Err.Description
    Resume Exit_ButtonSiteChar_Click
    
End Sub
