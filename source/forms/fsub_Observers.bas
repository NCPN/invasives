Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =3660
    DatasheetFontHeight =9
    ItemSuffix =6
    Left =10605
    Top =1695
    Right =14640
    Bottom =3150
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xde0128929108e340
    End
    RecordSource ="xref_Event_Contacts"
    Caption =" Observers"
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
            Height =312
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =996
                    Height =252
                    FontWeight =700
                    Name ="labContact_ID"
                    Caption ="Contact"
                End
                Begin Label
                    OverlapFlags =85
                    Left =2040
                    Top =60
                    Width =960
                    Height =252
                    FontWeight =700
                    Name ="labObserver_notes"
                    Caption ="Role"
                End
            End
        End
        Begin Section
            Height =366
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =16
                    Left =60
                    Top =60
                    Width =1923
                    Height =252
                    ColumnWidth =2268
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"0\""
                    Name ="cmbContact_ID"
                    ControlSource ="Contact_ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Contact_ID, [Last_Name] & (\", \"+[First_Name]) AS FullName FROM tlu_Cont"
                        "acts ORDER BY Last_Name, First_Name; "
                    ColumnWidths ="0;2880"
                    StatusBarText ="Observer identifier"
                    OnNotInList ="[Event Procedure]"
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2040
                    Top =60
                    Width =1503
                    Height =252
                    ColumnWidth =2376
                    TabIndex =1
                    Name ="txtObserver_notes"
                    ControlSource ="Contact_Role"
                    StatusBarText ="Comments about the observer specific to this sampling event"
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
' FORM NAME:    fsub_Observers
' Description:  Data entry form for observers associated with sampling events
' Data source:  tbl_Observers
' Data access:  edit, add and delete
' Pages:        none
' Functions:    none
' References:   none
' Source/date:  John R. Boetsch, June 2006
' Revisions:    <name, date, desc - add lines as you go>
' =================================

Private Sub cmbContact_ID_NotInList(NewData As String, Response As Integer)
    On Error GoTo Err_Handler

    Dim ctl As Control

    Set ctl = Me!cmbContact_ID
    ' Prompt user to verify they wish to add new value
    If MsgBox("The contact is not in list. Would you like to add this name?", vbYesNo) = vbYes Then
        Response = acDataErrContinue
        ctl.Undo
        DoCmd.OpenForm "frm_Contacts", , , , , , "new"
    Else
    ' Suppress error message and undo changes
        Response = acDataErrContinue
        ctl.Undo
    End If

Exit_Procedure:
    On Error Resume Next
    Set ctl = Nothing
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub
