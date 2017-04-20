Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Surface
' Level:        Framework class
' Version:      1.03
'
' Description:  Surface (microhabitat) object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, 4/17/2017
' References:   -
' Revisions:    BLC - 4/17/2017 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Long
Private m_SurfaceID As Long
Private m_SfcName As String
Private m_SfcDescription As String
Private m_OrigColumnName As String

'---------------------
' Events
'---------------------
Public Event InvalidID(value As Long)
Public Event InvalidSfcID(value As Long)
Public Event InvalidSfcName(value As String)
Public Event InvalidSfcDescription(value As String)
Public Event InvalidOrigColumnName(value As String)

'---------------------
' Properties
'---------------------
Public Property Let ID(value As Long)
    If varType(value) = vbLong Then
        m_ID = value
        'also set surfaceID value
        m_SurfaceID = value
    Else
        RaiseEvent InvalidID(value)
    End If
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let SurfaceID(value As Long)
    If varType(value) = vbLong Then
        m_SurfaceID = value
    Else
        RaiseEvent InvalidSfcID(value)
    End If
End Property

Public Property Get SurfaceID() As Long
    SurfaceID = m_SurfaceID
End Property

Public Property Let SfcName(value As String)
    'valid length varchar(25) or ZLS
    If IsBetween(Len(value), 1, 25, True) Then
        m_SfcName = value
    Else
        RaiseEvent InvalidSfcName(value)
    End If
End Property

Public Property Get SfcName() As String
    SfcName = m_SfcName
End Property

Public Property Let SfcDescription(value As String)
    'valid length varchar(255) or ZLS
    If IsBetween(Len(value), 1, 255, True) Then
        m_SfcDescription = value
    Else
        RaiseEvent InvalidSfcDescription(value)
    End If
End Property

Public Property Get SfcDescription() As String
    SfcDescription = m_SfcDescription
End Property

Public Property Let OrigColumnName(value As String)
    'valid length varchar(25) or ZLS
    If IsBetween(Len(value), 1, 25, True) Then
        m_OrigColumnName = value
    Else
        RaiseEvent InvalidOrigColumnName(value)
    End If
End Property

Public Property Get OrigColumnName() As String
    OrigColumnName = m_OrigColumnName
End Property

'---------------------
' Methods
'---------------------

' ---------------------------------
' Sub:          Class_Initialize
' Description:  Class initialization (starting) event
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 30, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/30/2015 - initial version
' ---------------------------------
Private Sub Class_Initialize()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Initialize[Surface class])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Class_Terminate
' Description:  Class termination (closing) event
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 30, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/30/2015 - initial version
' ---------------------------------
Private Sub Class_Terminate()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Terminate[cls_WoodyCanopy])"
    End Select
    Resume Exit_Handler
End Sub

'======== Custom Methods ===========
'---------------------------------------------------------------------------------------
' SUB:          Init
' Description:  Lookup surface based on surface/microhabitat ID
' Parameters:   ID - identifier for surface/microhabitat record (long)
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/17/2017 - for NCPN tools
' Revisions:
'   BLC, 4/17/2017 - initial version
'---------------------------------------------------------------------------------------
Public Sub Init(ID As Long)
On Error GoTo Err_Handler
    
    Dim rs As DAO.Recordset
    
    'set ID for parameters
    SetTempVar "SurfaceID", ID
    
    Set rs = GetRecords("s_surface_by_ID")
    If Not (rs.EOF And rs.BOF) Then
        With rs
            Me.ID = Nz(.Fields("ID"), 0)
            Me.SfcName = Nz(.Fields("Surface"), "")
            Me.SfcDescription = Nz(.Fields("Description"), "")
            Me.OrigColumnName = Nz(.Fields("ColName"), "")
        End With
    Else
        RaiseEvent InvalidID(ID)
    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Init[Surface class])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          SaveToDb
' Description:  Save surface/microhabitat based to database
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/17/2017 - for NCPN tools
' Revisions:
'   BLC, 4/17/2017 - initial version
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler
    
    Dim Template As String
    
    Template = "i_surface"
    
    Dim params(0 To 5) As Variant

    With Me
        params(0) = "Surface"
        params(1) = .SfcName
        params(2) = .SfcDescription
        params(3) = .OrigColumnName
        
        If IsUpdate Then
            Template = "u_surface"
            params(4) = .ID
        End If
        
        .ID = SetRecord(Template, params)
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[Surface class])"
    End Select
    Resume Exit_Handler
End Sub