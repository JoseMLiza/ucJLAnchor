Attribute VB_Name = "mdlAnchor"
Public Const ST_KEYDIGVAL = "0123456789."
Public CtrlParent As Object

Public Function GetControlIndex(objControl As Object) As String
    On Error GoTo ErrorFunction
    '---
    If objControl.Index >= 0 Then
        GetControlIndex = objControl.Index
    Else
        GoTo Fin
    End If
    Exit Function
    '---
ErrorFunction:
Fin:
    GetControlIndex = ""
End Function

Public Function GetControlScaleWidth(objControl As Object) As Long
    On Error GoTo FunctionError
    GetControlScaleWidth = objControl.ScaleWidth
    Exit Function
FunctionError:
    GetControlScaleWidth = objControl.Width
End Function

Public Function GetControlScaleHeight(objControl As Object) As Long
    On Error GoTo FunctionError
    GetControlScaleHeight = objControl.ScaleHeight
    Exit Function
FunctionError:
    GetControlScaleHeight = objControl.Height
End Function

Public Function GetControlHwnd(objControl As Object) As Long
    On Error GoTo FunctionError
    GetControlHwnd = objControl.Hwnd
    Exit Function
FunctionError:
    GetControlHwnd = 0
End Function

Public Function GetControlsParent(objParent As Object) As Long
    On Error GoTo ErrorFunction
    Dim c As Long
    For i = 0 To objParent.Controls.Count - 1
        If Len(objParent.Controls(i).Name) > 0 Then
            If objParent.Controls(i).Name <> "ucJLAnchor" Then
                c = c + 1
            End If
        End If
    Next
    GetControlsParent = c - 1
    Exit Function
ErrorFunction:
    GetControlsParent = c - 1
End Function

Public Function GetControlsCount(objControl As Object) As Integer
    On Error GoTo ErrorFunction
    GetControlsCount = objControl.Controls.Count
ErrorFunction:
    GetControlsCount = 0
End Function

Public Function ValidarKey(txtObject As TextBox, ValKey As Integer) As Integer
    Dim carKey As String
    '--
    ValidarKey = ValKey
    carKey = Chr(ValKey)
    '--
    If InStr(ST_KEYDIGVAL, carKey) = 0 Then
        ValidarKey = 0
    End If
End Function

Public Sub SafeRange(Value, Min, Max)
    If Value < Min Then Value = Min
    If Value > Max Then Value = Max
End Sub

