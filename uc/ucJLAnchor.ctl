VERSION 5.00
Begin VB.UserControl ucJLAnchor 
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   Picture         =   "ucJLAnchor.ctx":0000
   PropertyPages   =   "ucJLAnchor.ctx":0CCA
   ScaleHeight     =   480
   ScaleWidth      =   480
   ToolboxBitmap   =   "ucJLAnchor.ctx":0CDB
End
Attribute VB_Name = "ucJLAnchor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'USER32
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function RedrawWindow Lib "user32.dll" (ByVal Hwnd As Long, ByRef lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function UpdateWindow Lib "user32.dll" (ByVal Hwnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
'KERNEL32
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
                                                                         
Private Type RECT
   Left   As Long
   Top    As Long
   Right  As Long
   Bottom As Long
End Type

Private Type Settings
   RECT             As RECT
   ToAffectToLefts  As Boolean
   ToAffectToTops   As Boolean
   ContainerWidth   As Long
   ContainerHeight  As Long
End Type

Private Const WM_SETREDRAW As Long = &HB&
Private Const RDW_ALLCHILDREN As Long = &H80
Private Const RDW_INVALIDATE As Long = &H1
'--
Private Const WMSZ_BOTTOM = 6           'Borde inferior
Private Const WMSZ_BOTTOMLEFT = 7       'Esquina inferior izquierda
Private Const WMSZ_TOPLEFT = 4          'Esquina superior izquierda
Private Const WMSZ_LEFT = 1             'Borde izquierdo
Private Const WMSZ_BOTTOMRIGHT = 8      'Ezquina inferior derecha
Private Const WMSZ_RIGHT = 2            'Borde derecho
Private Const WMSZ_TOP = 3              'Borde superior
Private Const WMSZ_TOPRIGHT = 5         'Esquina superior derecha
'--
Private Const WMSZ_RESTORED = 0
Private Const WMSZ_MINIMIZED = 1
Private Const WMSZ_MAXIMIZED = 2
Private Const WMSZ_MAXSHOW = 3
Private Const WMSZ_MAXHIDE = 4
'--
Private Const GWL_STYLE = (-16)
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_THICKFRAME = &H40000
'--
Private Const WM_CREATE = &H1&
Private Const WM_SIZE = &H5&
Private Const WM_SIZING = &H214
Private Const WM_ENTERSIZEMOVE = &H231&
Private Const WM_PAINT As Long = &HF&
Private Const WM_SHOWWINDOW = &H18&
Private Const WM_CHILDACTIVATE = &H22&
'--
Public Controls             As New clsControls
Private m_Control           As clsControl
'--
Private m_FormMinWidth      As Long
Private m_FormMinHeight     As Long
Private m_FormMaxWidth      As Long
Private m_FormMaxHeight     As Long
'--
Private m_Hwnd              As Long
Private m_Container         As Object
Private m_ControlsCount     As Long
Private m_Tag               As String
Private m_Settings          As Settings
Private cSubClass           As clsSubClass

Dim WithEvents frmParent As Form
Attribute frmParent.VB_VarHelpID = -1
Dim i As Long

'm_FormMinWidth         As Long
Public Property Get FormMinWidth() As Long
    FormMinWidth = m_FormMinWidth
End Property
Public Property Let FormMinWidth(ByVal newValue As Long)
    If newValue < 0 Then newValue = 0
    If newValue > 0 And newValue > m_FormMaxWidth And m_FormMaxWidth > 0 Then
        MsgBox "FormMinWidth cannot be larger than FormMaxWidth", vbExclamation + vbOKOnly, UserControl.Name
        Exit Property
    End If
    m_FormMinWidth = newValue
    PropertyChanged "FormMinWidth"
End Property

'm_FormMinHeight        As Long
Public Property Get FormMinHeight() As Long
    FormMinHeight = m_FormMinHeight
End Property
Public Property Let FormMinHeight(ByVal newValue As Long)
    If newValue < 0 Then newValue = 0
    If newValue > 0 And newValue > m_FormMaxHeight And m_FormMaxHeight > 0 Then
        MsgBox "FormMinHeight cannot be larger than FormMaxHeight", vbExclamation + vbOKOnly, UserControl.Name
        Exit Property
    End If
    m_FormMinHeight = newValue
    PropertyChanged "FormMinHeight"
End Property

'm_FormMaxWidth         As Long
Public Property Get FormMaxWidth() As Long
    FormMaxWidth = m_FormMaxWidth
End Property
Public Property Let FormMaxWidth(ByVal newValue As Long)
    If newValue < 0 Then newValue = 0
    If newValue > 0 And newValue < m_FormMinWidth And m_FormMinWidth > 0 Then
        MsgBox "FormMaxWidth cannot be smaller than FormMinWidth", vbExclamation + vbOKOnly, UserControl.Name
        Exit Property
    End If
    If newValue > Screen.Width Then
        MsgBox "FormMaxWidth can not be greater than the limit of the screen." & vbCrLf & _
               "Maximum width (twips): " & Screen.Width & vbCrLf & _
               "Maximum width (pixels): " & ScaleX(Screen.Width, vbTwips, vbPixels), vbExclamation + vbOKOnly, UserControl.Name
        Exit Property
    End If
    m_FormMaxWidth = newValue
    PropertyChanged "FormMaxWidth"
End Property

'm_FormMaxHeight        As Long
Public Property Get FormMaxHeight() As Long
    FormMaxHeight = m_FormMaxHeight
End Property
Public Property Let FormMaxHeight(ByVal newValue As Long)
    If newValue < 0 Then newValue = 0
    If newValue > 0 And newValue < FormMinHeight And FormMinHeight > 0 Then
        MsgBox "FormMaxHeight cannot be smaller than FormMinHeight", vbExclamation + vbOKOnly, UserControl.Name
        Exit Property
    End If
    If newValue > Screen.Height Then
        MsgBox "FormMaxHeight can not be greater than the limit of the screen." & vbCrLf & _
               "Maximum height (twips): " & Screen.Height & vbCrLf & _
               "Maximum height (pixels): " & ScaleY(Screen.Height, vbTwips, vbPixels), vbExclamation + vbOKOnly, UserControl.Name
        Exit Property
    End If
    m_FormMaxHeight = newValue
    PropertyChanged "FormMaxHeight"
End Property

'm_Hwnd                 As Long
Public Property Get Hwnd() As Long
    Hwnd = UserControl.Hwnd
End Property

'm_Container            As Object
Public Property Get Container() As Object
    Set Container = Extender.Container
End Property

'm_ControlsCount        As Long
Public Property Get ControlsCount() As Long
    ControlsCount = m_ControlsCount
End Property

'm_Tag                  As String
Public Property Get Tag() As String
    Tag = m_Tag
End Property
Public Property Let Tag(ByVal newValue As String)
    m_Tag = newValue
    PropertyChanged "Tag"
End Property

'RunMode                As Boolean
Private Property Get RunMode() As Boolean
   On Error Resume Next
   RunMode = Ambient.UserMode
   RunMode = Extender.Parent.RunMode
End Property

Private Sub frmParent_Activate()
    Dim lStyle As Long
    '---
    If (m_FormMaxWidth < Screen.Width And m_FormMaxWidth > 0) Or (m_FormMaxHeight < Screen.Height And m_FormMaxHeight > 0) Then
        frmParent.WindowState = vbNormal
        lStyle = GetWindowLong(frmParent.Hwnd, GWL_STYLE)
        lStyle = lStyle And Not WS_MAXIMIZEBOX
        Call SetWindowLong(frmParent.Hwnd, GWL_STYLE, lStyle)
    End If
    'Resize Form MinWidth or MinHeight
    If frmParent.Width < m_FormMinWidth And m_FormMinWidth > 0 Then frmParent.Width = m_FormMinWidth
    If frmParent.Height < m_FormMinHeight And m_FormMinHeight > 0 Then frmParent.Height = m_FormMinHeight
    If frmParent.Width > m_FormMaxWidth And m_FormMaxWidth > 0 Then frmParent.Width = m_FormMaxWidth
    If frmParent.Height > m_FormMaxHeight And m_FormMaxHeight > 0 Then frmParent.Height = m_FormMaxHeight
    'AutoResize
    DoResize
End Sub

Private Sub frmParent_Resize()
    If ((m_FormMaxWidth < Screen.Width And m_FormMaxWidth > 0) Or (m_FormMaxHeight < Screen.Height And m_FormMaxHeight > 0)) And frmParent.WindowState = vbMaximized Then
        frmParent.WindowState = vbNormal
        If m_FormMaxWidth > 0 Then
            frmParent.Width = m_FormMaxWidth
        Else
            frmParent.Width = Screen.Width
        End If
        '--
        If m_FormMaxHeight > 0 Then
            frmParent.Height = m_FormMaxHeight
        Else
            frmParent.Height = Screen.Height
        End If
    End If
End Sub

Private Sub UserControl_InitProperties()
    m_ControlsCount = 0
    m_FormMinWidth = 0
    m_FormMinHeight = 0
    m_FormMaxWidth = 0
    m_FormMaxHeight = 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '---
    If RunMode Then
        If Not Extender.Visible Then
            Extender.Visible = True
        End If
    End If
    With PropBag
        m_FormMinWidth = .ReadProperty("FormMinWidth", 0)
        m_FormMinHeight = .ReadProperty("FormMinHeight", 0)
        m_FormMaxWidth = .ReadProperty("FormMaxWidth", 0)
        m_FormMaxHeight = .ReadProperty("FormMaxHeight", 0)
        'Settings
        m_Settings.ContainerWidth = .ReadProperty("ContainerWidth", 0)
        m_Settings.ContainerHeight = .ReadProperty("ContainerHeight", 0)
        '---
        m_ControlsCount = .ReadProperty("ControlsCount", 0)
        If m_ControlsCount > 0 Then
            For i = 1 To m_ControlsCount
                Set m_Control = .ReadProperty("Control_" & i, Nothing)
                If m_Control Is Nothing Then
                    Set m_Control = New clsControl
                End If
                Controls.Add m_Control
            Next
        End If
    End With
    '---
End Sub

Private Sub UserControl_Resize()
    UserControl.Size 32 * Screen.TwipsPerPixelX, 32 * Screen.TwipsPerPixelY
End Sub

Private Sub UserControl_Show()
    LoadControls
    If RunMode Then
        Extender.Left = -1000 'Ocultar control
        Set cSubClass = New clsSubClass
        With cSubClass
            If .ssc_Subclass(UserControl.ContainerHwnd, , , Me) Then
                .ssc_AddMsg UserControl.ContainerHwnd, WM_PAINT
            End If
        End With
        Set frmParent = Extender.Parent
    End If
End Sub

Private Sub UserControl_Hide()
    '---
    If RunMode Then
        cSubClass.ssc_UnSubclass UserControl.ContainerHwnd
        Set cSubClass = Nothing
    End If
    '---
End Sub

Private Sub UserControl_Terminate()
    Set Controls = Nothing
    Set Controls = Nothing
    If Not frmParent Is Nothing Then Set frmParent = Nothing
    If Not CtrlParent Is Nothing Then Set CtrlParent = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '---
    LoadControls
    With PropBag
        m_ControlsCount = Controls.Count
        '---
        .WriteProperty "FormMinWidth", m_FormMinWidth, 0
        .WriteProperty "FormMinHeight", m_FormMinHeight, 0
        .WriteProperty "ControlsCount", m_ControlsCount, 0
        .WriteProperty "FormMaxWidth", m_FormMaxWidth, 0
        .WriteProperty "FormMaxHeight", m_FormMaxHeight, 0
        '---
        If m_ControlsCount > 0 Then
            For i = 1 To m_ControlsCount
                If Controls.Item(i).TypeName <> "ucJLAnchor" Then
                    .WriteProperty "Control_" & i, Controls.Item(i), Nothing
                End If
            Next
        End If
    End With
    '---
End Sub

Private Sub LoadControls()
    On Error Resume Next
    Dim cControl As New clsControl
    Dim obj As Control
    Dim cCount As Long, cIni As Long, cCountP As Long
    Dim IsExists As Boolean, IsChange As Boolean
    Dim sTemp As String
    'If CtrlParent Is Nothing Then Set CtrlParent = Extender.Container
    Set CtrlParent = Extender.Container
    '--
    cCount = m_ControlsCount  'Controls.Count
    cCountP = GetControlsParent(CtrlParent)
    cIni = 1
    'Buscar controles huerfanos para eliminar.
    If cCount > 0 Then
        For i = cIni To cCount
            IsExists = False
            For Each obj In CtrlParent.Controls
                'If i <= cCountP Then
                    If TypeName(obj) & obj.Name & GetControlIndex(obj) = Controls.Item(i).TypeName & Controls.Item(i).Name & Controls.Item(i).ControlIndex Then
                        IsExists = True
                        Exit For
                    End If
                'Else
                '    IsExists = False
                '    Exit For
                'End If
            Next
            If Not IsExists And TypeName(obj) <> "ucJLAnchor" Then
                Controls.Remove i
                IsChange = True
            End If
        Next
    End If
    cCount = Controls.Count
    '---
    'Buscar nuevos controles para agregar.
    For Each obj In CtrlParent.Controls
        IsExists = False
        With obj
            If Not cCount > 0 Then
                If TypeName(obj) <> "ucJLAnchor" Then
                    cControl.ParentTypeName = TypeName(obj.Container)
                    cControl.ParentName = obj.Container.Name
                    cControl.ParentIndex = GetControlIndex(obj.Container.Index)
                    cControl.TypeName = TypeName(obj)
                    cControl.Name = .Name
                    cControl.ControlIndex = GetControlIndex(obj)
                    cControl.Hwnd = GetControlHwnd(obj)
                    cControl.Left = .Left
                    cControl.Top = .Top
                    cControl.Right = GetControlScaleWidth(obj.Container) - (obj.Left + obj.Width)
                    cControl.Bottom = GetControlScaleHeight(obj.Container) - (obj.Top + obj.Height)
                    '--
                    cControl.LeftPercent = (obj.Left * 100) / GetControlScaleWidth(obj.Container)
                    cControl.TopPercent = (obj.Top * 100) / GetControlScaleHeight(obj.Container)
                    cControl.WidthPercent = (obj.Width * 100) / GetControlScaleWidth(obj.Container)
                    cControl.HeightPercent = (obj.Height * 100) / GetControlScaleHeight(obj.Container)
                    '--
                    Controls.Add cControl
                    '--
                    IsChange = True
                End If
            Else
                For i = cIni To cCount
                    If TypeName(obj) <> "ucJLAnchor" Then
                        If TypeName(obj) & .Name & GetControlIndex(obj) = Controls.Item(i).TypeName & Controls.Item(i).Name & Controls.Item(i).ControlIndex Then
                            IsExists = True
                            Exit For
                        End If
                    End If
                Next
                If TypeName(obj) <> "ucJLAnchor" Then
                    If Not IsExists Then
                        cControl.ParentTypeName = TypeName(obj.Container)
                        cControl.ParentName = obj.Container.Name
                        cControl.ParentIndex = GetControlIndex(obj.Container.Index)
                        cControl.TypeName = TypeName(obj)
                        cControl.Name = .Name
                        cControl.ControlIndex = GetControlIndex(obj)
                        cControl.Hwnd = GetControlHwnd(obj)
                        cControl.Left = .Left
                        cControl.Top = .Top
                        cControl.Right = GetControlScaleWidth(obj.Container) - (obj.Left + obj.Width)
                        cControl.Bottom = GetControlScaleHeight(obj.Container) - (obj.Top + obj.Height)
                        '--
                        cControl.LeftPercent = (obj.Left * 100) / GetControlScaleWidth(obj.Container)
                        cControl.TopPercent = (obj.Top * 100) / GetControlScaleHeight(obj.Container)
                        cControl.WidthPercent = (obj.Width * 100) / GetControlScaleWidth(obj.Container)
                        cControl.HeightPercent = (obj.Height * 100) / GetControlScaleHeight(obj.Container)
                        '--
                        Controls.Add cControl
                        '--
                        IsChange = True
                    Else
                        If Not Ambient.UserMode Then
                            If Controls.Item(i).ParentTypeName <> TypeName(obj.Container) Then
                                Controls.Item(i).ParentTypeName = TypeName(obj.Container)
                            End If
                            If Controls.Item(i).ParentName <> obj.Container.Name Then
                                Controls.Item(i).ParentName = obj.Container.Name
                            End If
                            If Controls.Item(i).ParentIndex <> GetControlIndex(obj) Then
                                Controls.Item(i).ParentIndex = GetControlIndex(obj)
                            End If
                            If Controls.Item(i).ControlIndex <> GetControlIndex(obj) Then
                                Controls.Item(i).ControlIndex = GetControlIndex(obj)
                                IsChange = True
                            End If
                            'If Controls.Item(i).hWnd <> GetControlHwnd(obj) Then cControl.hWnd = GetControlHwnd(obj)
                            If Controls.Item(i).Left <> .Left Then
                                Controls.Item(i).Left = .Left
                                IsChange = True
                            End If
                            If Controls.Item(i).Top <> .Top Then
                                Controls.Item(i).Top = .Top
                                IsChange = True
                            End If
                            If Controls.Item(i).Right <> GetControlScaleWidth(obj.Container) - (obj.Left + obj.Width) Then
                                Controls.Item(i).Right = GetControlScaleWidth(obj.Container) - (obj.Left + obj.Width)
                                IsChange = True
                            End If
                            If Controls.Item(i).Bottom <> GetControlScaleHeight(obj.Container) - (obj.Top + obj.Height) Then
                                Controls.Item(i).Bottom = GetControlScaleHeight(obj.Container) - (obj.Top + obj.Height)
                                IsChange = True
                            End If
                            '--> LeftPercent
                            If Controls.Item(i).LeftPercent <> (obj.Left * 100) / GetControlScaleWidth(obj.Container) Then
                                Controls.Item(i).LeftPercent = (obj.Left * 100) / GetControlScaleWidth(obj.Container)
                            End If
                            '--> TopPercent
                            If Controls.Item(i).TopPercent <> (obj.Top * 100) / GetControlScaleHeight(obj.Container) Then
                                Controls.Item(i).TopPercent = (obj.Top * 100) / GetControlScaleHeight(obj.Container)
                            End If
                            '--> WidthPercent
                            If Controls.Item(i).WidthPercent <> (obj.Width * 100) / GetControlScaleWidth(obj.Container) Then
                                Controls.Item(i).WidthPercent = (obj.Width * 100) / GetControlScaleWidth(obj.Container)
                            End If
                            '--> HeightPercent
                            If Controls.Item(i).HeightPercent <> (obj.Height * 100) / GetControlScaleHeight(obj.Container) Then
                                Controls.Item(i).HeightPercent = (obj.Height * 100) / GetControlScaleHeight(obj.Container)
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Next
    '--
    If IsChange Then sTemp = m_Tag: Tag = "Edit": Tag = sTemp
    '--
End Sub

Private Sub DoResize()
    Dim a As Long, objControl As Control
    Dim idx As String
    Dim sWidth As Long, sHeight As Long
    'Resize Controls
    Call SendMessage(frmParent.Hwnd, WM_SETREDRAW, 0&, 0&)
    '--
    For Each objControl In frmParent
        '--
        idx = GetControlIndex(objControl)
        '--
        sWidth = GetControlScaleWidth(objControl.Container)
        sHeight = GetControlScaleHeight(objControl.Container)
        '--
        For a = 1 To Controls.Count
            If Controls.Item(a).TypeName = TypeName(objControl) Then
                If Controls.Item(a).Name & Controls.Item(a).ControlIndex = objControl.Name & idx Then
                    With Controls.Item(a)
                        If .AnchorRight Then
                            If .AnchorLeft Then
                                objControl.Width = IIf(sWidth - (.Left + .Right) > 0, sWidth - (.Left + .Right), 0)
                            Else
                                'If Not .LeftPercent > 0 Then
                                If Not .UseLeftPercent Then
                                    objControl.Left = sWidth - (objControl.Width + .Right)
                                Else
                                    objControl.Left = sWidth * (.LeftPercent / 100)
                                End If
                            End If
                        Else
                            'If Not .AnchorLeft And Not .LeftPercent > 0 Then
                            If Not .AnchorLeft And Not .UseLeftPercent Then
                                If m_Settings.ToAffectToLefts Then
                                    objControl.Left = .Left
                                End If
                            'ElseIf Not .AnchorLeft And .LeftPercent > 0 Then
                            ElseIf Not .AnchorLeft And .UseLeftPercent Then
                                objControl.Left = sWidth * (.LeftPercent / 100)
                            End If
                            '--
                            'If .WidthPercent > 0 Then objControl.Width = sWidth * (.WidthPercent / 100)
                            If .UseWidthPercent Then objControl.Width = sWidth * (.WidthPercent / 100)
                        End If
                        If .AnchorBottom Then
                            If .AnchorTop Then
                                objControl.Height = IIf(sHeight - (.Top + .Bottom) > 0, sHeight - (.Top + .Bottom), 0)
                            Else
                                'If Not .TopPercent > 0 Then
                                If Not .UseTopPercent Then
                                    objControl.Top = sHeight - (objControl.Height + .Bottom)
                                Else
                                    objControl.Top = sHeight * (.TopPercent / 100)
                                End If
                            End If
                        Else
                            'If Not .AnchorTop And Not .TopPercent > 0 Then
                            If Not .AnchorTop And Not .UseTopPercent Then
                                If m_Settings.ToAffectToTops Then
                                    objControl.Top = .Top
                                End If
                            'ElseIf Not .AnchorTop And .TopPercent > 0 Then
                            ElseIf Not .AnchorTop And .UseTopPercent Then
                                objControl.Top = sHeight * (.TopPercent / 100)
                            End If
                            '--
                            'If .HeightPercent > 0 Then objControl.Height = sHeight * (.HeightPercent / 100)
                            If .UseHeightPercent Then objControl.Height = sHeight * (.HeightPercent / 100)
                        End If
                    End With
                End If
            End If
        Next
    Next
    '---
    Call SendMessage(frmParent.Hwnd, WM_SETREDRAW, 1&, 0&)
    RedrawWindow frmParent.Hwnd, ByVal &H0, 0, RDW_INVALIDATE Or RDW_ALLCHILDREN
    '--
End Sub

Private Sub WndProc(ByVal bBefore As Boolean, _
                    ByRef bHandled As Boolean, _
                    ByRef lReturn As Long, _
                    ByVal Hwnd As Long, _
                    ByVal uMsg As Long, _
                    ByVal wParam As Long, _
                    ByVal lParam As Long, _
                    ByRef lParamUser As Long)
                    
    Dim r As RECT
    Dim sMinWidth As Long, sMinHeight As Long
    Dim sMaxWidth As Long, sMaxHeight As Long
    '---
    Select Case uMsg
        Case WM_PAINT
            If UserControl.ContainerHwnd Then
                cSubClass.ssc_DelMsg UserControl.ContainerHwnd, WM_PAINT
                cSubClass.ssc_AddMsg UserControl.ContainerHwnd, WM_SIZE, MSG_BEFORE_AFTER
                cSubClass.ssc_AddMsg UserControl.ContainerHwnd, WM_ENTERSIZEMOVE, MSG_AFTER
                cSubClass.ssc_AddMsg UserControl.ContainerHwnd, WM_SIZING, MSG_BEFORE
                '---
            End If
        Case WM_ENTERSIZEMOVE
            If Ambient.UserMode Then
                m_Settings.ContainerWidth = Extender.Container.ScaleWidth
                m_Settings.ContainerHeight = Extender.Container.ScaleHeight
            End If
        Case WM_SIZING
            sMinWidth = ScaleX(m_FormMinWidth, vbTwips, vbPixels)
            sMinHeight = ScaleY(m_FormMinHeight, vbTwips, vbPixels)
            sMaxWidth = ScaleX(m_FormMaxWidth, vbTwips, vbPixels)
            sMaxHeight = ScaleY(m_FormMaxHeight, vbTwips, vbPixels)
            '---
            Call CopyMemory(r, ByVal lParam, Len(r))
            '/////////
            '* WIDTH *
            '/////////
            'MinWidth
            If (r.Right - r.Left < sMinWidth) And m_FormMinWidth > 0 Then
                Select Case wParam
                    Case WMSZ_TOPLEFT, WMSZ_LEFT, WMSZ_BOTTOMLEFT 'Left Part
                        r.Left = r.Right - sMinWidth
                        m_Settings.ToAffectToLefts = True
                    Case WMSZ_TOPRIGHT, WMSZ_RIGHT, WMSZ_BOTTOMRIGHT 'Right Part
                        r.Right = r.Left + sMinWidth
                        m_Settings.ToAffectToLefts = False
                End Select
            Else
                Select Case wParam
                    Case WMSZ_TOPLEFT, WMSZ_LEFT, WMSZ_BOTTOMLEFT 'Left Part
                        m_Settings.ToAffectToLefts = True
                    Case WMSZ_TOPRIGHT, WMSZ_RIGHT, WMSZ_BOTTOMRIGHT 'Right Part
                        m_Settings.ToAffectToLefts = False
                End Select
            End If
            'MaxWidth
            If (r.Right - r.Left > sMaxWidth) And m_FormMaxWidth > 0 Then
                Select Case wParam
                    Case WMSZ_LEFT, WMSZ_BOTTOMLEFT, WMSZ_TOPLEFT
                        r.Left = r.Right - sMaxWidth
                        m_Settings.ToAffectToLefts = True
                    Case WMSZ_RIGHT, WMSZ_BOTTOMRIGHT, WMSZ_TOPRIGHT
                        r.Right = r.Left + sMaxWidth
                        m_Settings.ToAffectToLefts = False
                End Select
            Else
                Select Case wParam
                    Case WMSZ_TOPLEFT, WMSZ_LEFT, WMSZ_BOTTOMLEFT 'Left Part
                        m_Settings.ToAffectToLefts = True
                    Case WMSZ_TOPRIGHT, WMSZ_RIGHT, WMSZ_BOTTOMRIGHT 'Right Part
                        m_Settings.ToAffectToLefts = False
                End Select
            End If
            '//////////
            '* HEIGHT *
            '//////////
            'MinHeight
            If (r.Bottom - r.Top < sMinHeight) And m_FormMinHeight > 0 Then
                Select Case wParam
                    Case WMSZ_TOPLEFT, WMSZ_TOP, WMSZ_TOPRIGHT 'Top Part
                        r.Top = r.Bottom - sMinHeight
                        m_Settings.ToAffectToTops = True
                    Case WMSZ_BOTTOMLEFT, WMSZ_BOTTOM, WMSZ_BOTTOMRIGHT 'Bottom Part
                        r.Bottom = r.Top + sMinHeight
                        m_Settings.ToAffectToTops = False
                End Select
            Else
                Select Case wParam
                    Case WMSZ_TOPLEFT, WMSZ_TOP, WMSZ_TOPRIGHT 'Top Part
                        m_Settings.ToAffectToTops = True
                    Case WMSZ_BOTTOMLEFT, WMSZ_BOTTOM, WMSZ_BOTTOMRIGHT 'Bottom Part
                        m_Settings.ToAffectToTops = False
                End Select
            End If
            'MaxHeight
            If (r.Bottom - r.Top > sMaxHeight) And m_FormMaxHeight > 0 Then
                Select Case wParam
                    Case WMSZ_TOP, WMSZ_TOPLEFT, WMSZ_TOPRIGHT
                        r.Top = r.Bottom - sMaxHeight
                        m_Settings.ToAffectToTops = True
                    Case WMSZ_BOTTOM, WMSZ_BOTTOMLEFT, WMSZ_BOTTOMRIGHT
                        r.Bottom = r.Top + sMaxHeight
                        m_Settings.ToAffectToTops = False
                End Select
            Else
                Select Case wParam
                    Case WMSZ_TOP, WMSZ_TOPLEFT, WMSZ_TOPRIGHT
                        m_Settings.ToAffectToTops = True
                    Case WMSZ_BOTTOM, WMSZ_BOTTOMLEFT, WMSZ_BOTTOMRIGHT
                        m_Settings.ToAffectToTops = False
                End Select
            End If
            '----------------------------------------------------------------------
            Call CopyMemory(ByVal lParam, r, Len(r))
        Case WM_SIZE
            m_Settings.ContainerWidth = Extender.Container.ScaleWidth
            m_Settings.ContainerHeight = Extender.Container.ScaleHeight
            Call DoResize
        Case WM_CHILDACTIVATE
            m_Settings.ContainerWidth = Extender.Container.ScaleWidth
            m_Settings.ContainerHeight = Extender.Container.ScaleHeight
            Call DoResize
    End Select
End Sub
