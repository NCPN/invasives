Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =126
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5760
    DatasheetFontHeight =9
    ItemSuffix =6
    Left =2100
    Top =2430
    Right =7605
    Bottom =7200
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x3b0b2d760f50e340
    End
    OnOpen ="[Event Procedure]"
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
        Begin Section
            Height =5040
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1440
                    Top =480
                    Width =2880
                    Height =480
                    FontSize =14
                    FontWeight =700
                    Name ="Label0"
                    Caption ="Report Menu"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2340
                    Top =4020
                    Width =1035
                    Height =405
                    TabIndex =4
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1800
                    Top =1800
                    Width =2160
                    TabIndex =1
                    Name ="ButtonPE"
                    Caption ="Plot Revisit Data Sheet"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =1800
                    Top =2880
                    Width =2160
                    TabIndex =3
                    Name ="ButtonNP"
                    Caption ="Not Present in Park Report"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1800
                    Top =2340
                    Width =2160
                    TabIndex =2
                    Name ="ButtonSpeciesPresence"
                    Caption ="Species Presence by Plot"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1800
                    Top =1260
                    Width =2160
                    Name ="ButtonOverstory"
                    Caption ="Overstory Plot Revisit"
                    OnClick ="[Event Procedure]"
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

Private Sub ButtonClose_Click()
On Error GoTo Err_ButtonClose_Click


    DoCmd.Close

Exit_ButtonClose_Click:
    Exit Sub

Err_ButtonClose_Click:
    MsgBox Err.Description
    Resume Exit_ButtonClose_Click
    
End Sub
Private Sub ButtonPE_Click()
On Error GoTo Err_ButtonPE_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Select_Plot_Establishment"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "frm_Report_Menu"
Exit_ButtonPE_Click:
    Exit Sub

Err_ButtonPE_Click:
    MsgBox Err.Description
    Resume Exit_ButtonPE_Click
    
End Sub
Private Sub ButtonNP_Click()
On Error GoTo Err_ButtonNP_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Select_Not_Present"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "frm_Report_Menu"

Exit_ButtonNP_Click:
    Exit Sub

Err_ButtonNP_Click:
    MsgBox Err.Description
    Resume Exit_ButtonNP_Click
    
End Sub

Private Sub Form_Open(Cancel As Integer)
  DoCmd.Restore
End Sub
Private Sub ButtonSpeciesPresence_Click()
On Error GoTo Err_ButtonSpeciesPresence_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Species_Report_Select"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonSpeciesPresence_Click:
    Exit Sub

Err_ButtonSpeciesPresence_Click:
    MsgBox Err.Description
    Resume Exit_ButtonSpeciesPresence_Click
    
End Sub
Private Sub ButtonOverstory_Click()
On Error GoTo Err_ButtonOverstory_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Select_Overstory_Revisit"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "frm_Report_Menu"
Exit_ButtonOverstory_Click:
    Exit Sub

Err_ButtonOverstory_Click:
    MsgBox Err.Description
    Resume Exit_ButtonOverstory_Click
    
End Sub
