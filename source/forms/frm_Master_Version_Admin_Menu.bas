Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DefaultView =0
    TabularFamily =124
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5100
    DatasheetFontHeight =10
    ItemSuffix =7
    Left =3720
    Top =1485
    Right =8820
    Bottom =4335
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x6a97492f33fee240
    End
    Caption ="Master Version Admin Menu"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
            Height =2865
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =1440
                    Top =1980
                    Width =2280
                    Height =405
                    TabIndex =2
                    Name ="ButtonClose"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =120
                    Top =60
                    Width =4860
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Label2"
                    Caption ="Master Version Administration Menu"
                    FontName ="MS Sans Serif"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1440
                    Top =660
                    Width =2280
                    Height =480
                    Name ="ButtonShowAll"
                    Caption ="Show All Master Version Keys"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1440
                    Top =1320
                    Width =2279
                    Height =405
                    TabIndex =1
                    Name ="ButtonInitialize"
                    Caption ="Initialize Protocol"
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



Private Sub ButtonShowAll_Click()

On Error GoTo Err_ButtonShowAll_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Version_List"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonShowAll_Click:
    Exit Sub

Err_ButtonShowAll_Click:
    MsgBox Err.Description
    Resume Exit_ButtonShowAll_Click

End Sub


Private Sub ButtonInitialize_Click()
On Error GoTo Err_ButtonInitialize_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Show_All_Versions"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_ButtonInitialize_Click:
    Exit Sub

Err_ButtonInitialize_Click:
    MsgBox Err.Description
    Resume Exit_ButtonInitialize_Click
    
End Sub
