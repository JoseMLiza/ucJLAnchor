Attribute VB_Name = "mdlAnchor"
'-----------------------------
'Autor: Jose Liza
'ClassName: clsControl
'Date: 12/06/2022
'Version: 0.0.1
'-----------------------------
Option Explicit
Public Const ST_KEYDIGVAL = "0123456789."
Public Const OBJ_EXCLUDED = "/Error/Menu/ucJLAnchor/Form/Toolbar/StatusBar"
Public CtrlParent As Object

Dim i As Long

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

Public Function GetUseScaleWidth(objControl As Object) As Boolean
    On Error GoTo FunctionError
    Dim lW As Long
    lW = objControl.ScaleWidth
    GetUseScaleWidth = True
    Exit Function
FunctionError:
    GetUseScaleWidth = False
End Function

Public Function GetControlScaleWidth(objControl As Object) As Long
    On Error GoTo FunctionError
    GetControlScaleWidth = objControl.ScaleWidth
    Exit Function
FunctionError:
    GetControlScaleWidth = objControl.Width
End Function

Public Function GetUseScaleHeight(objControl As Object) As Boolean
    On Error GoTo FunctionError
    Dim lW As Long
    lW = objControl.ScaleHeight
    GetUseScaleHeight = True
    Exit Function
FunctionError:
    GetUseScaleHeight = False
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
    GetControlHwnd = objControl.hWnd
    Exit Function
FunctionError:
    GetControlHwnd = 0
End Function

Public Function GetControlsParent(objParent As Object) As Long
    On Error GoTo ErrorFunction
    Dim c As Long
    For i = 0 To objParent.Controls.Count - 1
        If Len(objParent.Controls(i).name) > 0 Then
            If objParent.Controls(i).name <> "ucJLAnchor" Then
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
    Exit Function
ErrorFunction:
    GetControlsCount = 0
End Function

Public Function GetControlInContainer(objControl As Object) As Boolean
On Error GoTo ErrorFunction
'---
    Dim name As String
    name = TypeName(objControl.Container)
    GetControlInContainer = True
    Exit Function
'---
ErrorFunction:
    GetControlInContainer = False
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

Public Function GetMinWidthControl(ByVal objControl As Object) As Long
    Dim lTemp As Long
    '--
    lTemp = objControl.Width
    objControl.Width = 0
    GetMinWidthControl = objControl.Width
    objControl.Width = lTemp
    '--
End Function

Public Function GetMinHeightControl(ByVal objControl As Object) As Long
    Dim lTemp As Long
    '--
    lTemp = objControl.Height
    objControl.Height = 0
    GetMinHeightControl = objControl.Height
    objControl.Height = lTemp
    '--
End Function

Public Function GetContainerTypeName(ByVal objControl As Object) As String
On Error GoTo FunctionError
    GetContainerTypeName = TypeName(objControl.Container)
    Exit Function
FunctionError:
    GetContainerTypeName = "Error"
End Function

Public Function GetParentScaleMode(ByVal objControl As Object) As ScaleModeConstants
On Error GoTo FunctionError
    GetParentScaleMode = objControl.Container.ScaleMode
    Exit Function
FunctionError:
    GetParentScaleMode = vbTwips
End Function

Public Function GetParentUseScaleMode(ByVal objControl As Object) As Boolean
On Error GoTo FunctionError
    Dim sMode As ScaleModeConstants
    sMode = objControl.Container.ScaleMode
    GetParentUseScaleMode = True
    Exit Function
FunctionError:
    GetParentUseScaleMode = False
End Function


Public Function BytesLength(abBytes() As Byte) As Long
    On Error Resume Next
    BytesLength = UBound(abBytes) - LBound(abBytes) + 1
End Function

Public Sub DrawBackAreaIcon(obj As Object)
    Dim i As Long, j As Long, x As Long
    '--
    For j = -1 To obj.ScaleHeight Step 5
        x = IIf(x = -1, 4, -1)
        For i = x To obj.ScaleWidth Step 10
            obj.Line (i, j)-(i + 4, j + 4), &HF2F2F2, BF
        Next
    Next
    '--
    obj.Line (0, 0)-(obj.ScaleWidth - 1, obj.ScaleHeight - 1), vbButtonShadow, B
    obj = obj.Image
End Sub

