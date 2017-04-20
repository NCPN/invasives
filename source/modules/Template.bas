Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Template
' Level:        Framework class
' Version:      1.00
'
' Description:  Template object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 10/4/2016
' References:   -
' Revisions:    BLC - 10/4/2016 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Long

Private m_EventID As Long

Private m_TemplateName As String '255
Private m_Context As String '255
Private m_Syntax As String '10
Private m_TemplateSQL As String 'memo
Private m_Params As String '255
Private m_Version As String '10
Private m_IsSupported As Integer
Private m_Remarks As String '255
Private m_EffectiveDate As Date 'date
Private m_RetireDate As Date 'date

'creator/modifier
Private m_ContactID As Long

'---------------------
' Events
'---------------------
Public Event InvalidTemplateSQL(value As String)
Public Event InvalidSyntax(value As String)

'---------------------
' Properties
'---------------------
Public Property Let ID(value As Long)
    m_ID = value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let TemplateName(value As String)
    m_TemplateName = value
End Property

Public Property Get TemplateName() As String
    TemplateName = m_TemplateName
End Property

Public Property Let Context(value As String)
    m_Context = value
End Property

Public Property Get Context() As String
    Context = m_Context
End Property

Public Property Let TemplateSQL(value As String)
    m_TemplateSQL = value
    
    'set params property
    If Len(Me.Syntax) <> Len(Replace(Me.Syntax, "SQL", "")) Then
        Me.params = GetParamsFromSQL(Me.TemplateSQL)
    End If
    
End Property

Public Property Get TemplateSQL() As String
    TemplateSQL = m_TemplateSQL
End Property

Public Property Let Syntax(value As String)
    m_Syntax = value
End Property

Public Property Get Syntax() As String
    Syntax = m_Syntax
End Property

Public Property Let params(value As String)
    m_Params = value
End Property

Public Property Get params() As String
    params = m_Params
End Property

Public Property Let Version(value As String)
    m_Version = value
End Property

Public Property Get Version() As String
    Version = m_Version
End Property

Public Property Let Remarks(value As String)
    m_Remarks = value
End Property

Public Property Get Remarks() As String
    Remarks = m_Remarks
End Property

Public Property Let IsSupported(value As Integer)
    m_IsSupported = value
End Property

Public Property Get IsSupported() As Integer
    IsSupported = m_IsSupported
End Property

Public Property Let EffectiveDate(value As Date)
    m_EffectiveDate = Format(value, "mm/dd/yyyy")
End Property

Public Property Get EffectiveDate() As Date
    EffectiveDate = m_EffectiveDate
End Property

Public Property Let RetireDate(value As Date)
    m_RetireDate = Format(value, "mm/dd/yyyy")
End Property

Public Property Get RetireDate() As Date
    RetireDate = m_RetireDate
End Property

Public Property Let ContactID(value As Long)
    m_ContactID = value
End Property

Public Property Get ContactID() As Long
    ContactID = m_ContactID
End Property

'---------------------
' Methods
'---------------------

'======== Standard Methods ===========

' ---------------------------------
' SUB:          Class_Initialize
' Description:  Initialize the class
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  -
' Adapted:      Bonnie Campbell, April 4, 2016 - for NCPN tools
' Revisions:
'   BLC - 4/4/2016 - initial version
' ---------------------------------
Private Sub Class_Initialize()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Class_Initialize[cls_Template])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          Class_Terminate
' Description:  -
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/4/2016 - for NCPN tools
' Revisions:
'   BLC, 4/4/2016 - initial version
'---------------------------------------------------------------------------------------
Private Sub Class_Terminate()
On Error GoTo Err_Handler

    'Set m_ID = 0

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Class_Terminate[cls_Template])"
    End Select
    Resume Exit_Handler
End Sub

'======== Custom Methods ===========
'---------------------------------------------------------------------------------------
' SUB:          SaveToDb
' Description:  -
' Parameters:   -
' Returns:      -
' Throws:       -
' References:
'   Fionnuala, February 2, 2009
'   David W. Fenton, October 27, 2009
'   http://stackoverflow.com/questions/595132/how-to-get-id-of-newly-inserted-record-using-excel-vba
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/4/2016 - for NCPN tools
' Revisions:
'   BLC, 4/4/2016 - initial version
'   BLC, 8/8/2016 - added update parameter to identify if this is an update vs. an insert
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler

    Dim Template As String
    
    Template = "i_template"
    
    Dim params(0 To 12) As Variant

    With Me
        params(0) = Template
        params(1) = .TemplateName
        params(2) = .Context
        params(3) = .TemplateSQL
        params(4) = .Remarks
        params(5) = .EffectiveDate
        params(6) = .ContactID
        params(7) = .params
        params(8) = .Syntax
        params(9) = .Version
        params(10) = .IsSupported
        params(11) = IIf(IsDate(.RetireDate), _
                     IIf(.RetireDate = #12:00:00 AM#, Null, .RetireDate), Null)
    
        If IsUpdate Then
            Template = "u_template"
            params(12) = .ID
        End If

        .ID = SetRecord(Template, params)
    End With
    
    'after template is saved, refresh global Template dictionary
    GetTemplates
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case 457 'key already element --> template already exists
            MsgBox _
                "Template " & Me.TemplateName & " is a duplicate. Please contact " _
                & "a data manager to fix this for you. " _
                & vbCrLf & "If you are a data manager, oops." _
                & vbCrLf & "Remove the duplicate template from tsys_Db_Templates " _
                & "and try again." _
                & vbCrLf & vbCrLf & "Error #" & Err.Description _
                & "Error encountered (#" & Err.Number & " - SaveToDb[cls_Template])", _
                vbCritical, "Duplicate Template"
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[cls_Template])"
    End Select
    Resume Exit_Handler
End Sub