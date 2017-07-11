Option Compare Database
Option Explicit


Public Function FormCheck(frm As Form) As Boolean
Dim ctl As Control
Dim booHasData As Boolean

For Each ctl In frm.Controls
    If InStr(ctl.tag, "<data>") > 0 Then
        If Not IsNothing(ctl.Value) Then
            booHasData = True
            Exit For
        End If
    End If
Next
FormCheck = booHasData
End Function


Public Sub UpdateControl(varOpenArgs As Variant)
Dim strFormName As String
Dim strControlName As String

On Error Resume Next

If Not IsNothing(varOpenArgs) Then
    strFormName = XML_Read("FormFrom", CStr(varOpenArgs))
    strControlName = XML_Read("ControlFrom", CStr(varOpenArgs))
    
    If Len(strFormName) > 0 And Len(strControlName) > 0 Then
        If IsLoaded(strFormName) Then
            Forms(strFormName)(strControlName).Requery
        End If
    End If
End If

End Sub