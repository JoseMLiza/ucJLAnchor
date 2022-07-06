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
'-----------------------------
'Autor: Jose Liza
'Date: 12/06/2022
'Version: 0.0.1
'Thanks: Leandro Ascierto (www.leandroascierto.com) And Latin Group of VB6
'-----------------------------
Option Explicit
'USER32
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByRef lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'--
Private Declare Function CreateIconFromResourceEx Lib "user32" (ByRef presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

'KERNEL32
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

'OLE32
Private Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)

'OLEAUT32
Private Declare Function OleCreatePictureIndirect Lib "OleAut32" (lpPictDesc As Any, riid As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

'GDIPLUS
Private Declare Function GdiplusStartup Lib "GdiPlus" (token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "GdiPlus" (ByVal token As Long)
'--
Private Declare Function GdipDrawImageRect Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long
Private Declare Function GdipSetInterpolationMode Lib "GdiPlus" (ByVal graphics As Long, ByVal InterpolationMode As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "GdiPlus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "GdiPlus" (ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, ByVal Format As Long, ByRef Scan0 As Any, ByRef Bitmap As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "GdiPlus" (ByVal Stream As Any, ByRef Image As Long) As Long
Private Declare Function GdipGetImageBounds Lib "GdiPlus" (ByVal mImage As Long, ByRef mSrcRect As RECTF, ByRef mSrcUnit As Long) As Long
Private Declare Function GdipCreateHICONFromBitmap Lib "GdiPlus" (ByVal Bitmap As Long, hbmReturn As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GdiPlus" (ByVal mHbm As Long, ByVal mhPal As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "GdiPlus" (ByVal mHicon As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipDisposeImage Lib "GdiPlus" (ByVal Image As Long) As Long

'GDI32
Private Declare Function GetObjectType Lib "gdi32" (ByVal hGDIObj As Long) As Long

'MSVBVM60
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long

Private Type GDIPlusStartupInput
    GdiPlusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Type RECTF
    Left        As Single
    Top         As Single
    Width       As Single
    Height      As Single
End Type

Private Type Settings
    RECT             As RECT
    ToAffectToLefts  As Boolean
    ToAffectToTops   As Boolean
    ContainerWidth   As Long
    ContainerHeight  As Long
End Type

Private Type IconHeader
    ihReserved          As Integer
    ihType              As Integer
    ihCount             As Integer
End Type

Private Type IconEntry
    ieWidth             As Byte
    ieHeight            As Byte
    ieColorCount        As Byte
    ieReserved          As Byte
    iePlanes            As Integer
    ieBitCount          As Integer
    ieBytesInRes        As Long
    ieImageOffset       As Long
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
Private Const UnitPixel                     As Long = &H2&
Private Const OBJ_BITMAP                    As Long = 7
'--
Private Const InterpolationModeHighQuality  As Long = &H2
Private Const IconVersion                   As Long = &H30000
Private Const PixelFormat32bppARGB          As Long = &H26200A
Private Const ICON_JUMBO                    As Long = 256
Private Const ICON_BIG                      As Long = 1
Private Const ICON_SMALL                    As Long = 0
Private Const WM_SETICON                    As Long = &H80
'--
Public Controls             As New clsControls
Private m_Control           As clsControl
'--
Private m_IconPresent       As Boolean
Private m_FormIcon()        As Byte
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
Dim hIcon As Long

'm_IconPresent          As Boolean
Public Property Get IconPresent() As Boolean
    IconPresent = m_IconPresent
End Property
Public Property Let IconPresent(ByVal newValue As Boolean)
    m_IconPresent = newValue
    PropertyChanged "IconPresent"
End Property

'm_FormIcon             As Byte
Public Property Get FormIcon() As Variant
    FormIcon = m_FormIcon
End Property
Public Property Let FormIcon(ByRef newValue As Variant)
    If m_IconPresent Then
        m_FormIcon = newValue
        PropertyChanged "FormIcon"
    End If
End Property


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
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
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
        lStyle = GetWindowLong(frmParent.hWnd, GWL_STYLE)
        lStyle = lStyle And Not WS_MAXIMIZEBOX
        Call SetWindowLong(frmParent.hWnd, GWL_STYLE, lStyle)
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
        m_IconPresent = .ReadProperty("IconPresent", False)
        If m_IconPresent Then m_FormIcon = .ReadProperty("FormIcon")
        '--
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
    'Load icon
    'If m_IconPresent And App.LogMode Then
    If m_IconPresent Then
        'Set frmParent.Icon = Nothing
        Call SetIconForm(UserControl.Parent.hWnd, m_FormIcon, 32, 32)
        'IsDrawIcon = True
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
    DestroyIcon hIcon
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '---
    LoadControls
    With PropBag
        m_ControlsCount = Controls.Count
        '---
        .WriteProperty "IconPresent", m_IconPresent
        If m_IconPresent Then
            .WriteProperty "FormIcon", m_FormIcon
        Else
            .WriteProperty "FormIcon", 0
        End If
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
    Dim cCount As Long, cIni As Long, cCountP As Long, lTemp As Long
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
                If TypeName(obj) & obj.Name & GetControlIndex(obj) = Controls.Item(i).TypeName & Controls.Item(i).Name & Controls.Item(i).ControlIndex Then
                    IsExists = True
                    Exit For
                End If
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
                    cControl.hWnd = GetControlHwnd(obj)
                    cControl.Left = .Left
                    cControl.Top = .Top
                    cControl.Right = GetControlScaleWidth(obj.Container) - (obj.Left + obj.Width)
                    cControl.Bottom = GetControlScaleHeight(obj.Container) - (obj.Top + obj.Height)
                    '--
                    cControl.MinWidth = GetMinWidthControl(obj)
                    cControl.MinHeight = GetMinHeightControl(obj)
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
                        cControl.hWnd = GetControlHwnd(obj)
                        cControl.Left = .Left
                        cControl.Top = .Top
                        cControl.Right = GetControlScaleWidth(obj.Container) - (obj.Left + obj.Width)
                        cControl.Bottom = GetControlScaleHeight(obj.Container) - (obj.Top + obj.Height)
                        '--
                        cControl.MinWidth = GetMinWidthControl(obj)
                        cControl.MinHeight = GetMinHeightControl(obj)
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
                        '--> Validar cambio de resolucion
                        If obj.Left <= Screen.Width And (obj.Width) > (Screen.Width - obj.Left) Then
                            Debug.Print UserControl.Parent.Name & ": Width of control " & obj.Name & " exceeds that of the screen, the positions have been updated, verify."
                            obj.Width = GetControlScaleWidth(obj.Container) - (obj.Left + (Controls.Item(i).Right))
                            IsChange = True
                        End If
                        If obj.Top <= Screen.Height And (obj.Height) > (Screen.Height - obj.Top) Then
                            Debug.Print UserControl.Parent.Name & ": Height of control " & obj.Name & " exceeds that of the screen, the positions have been updated, verify."
                            obj.Height = GetControlScaleHeight(obj.Container) - (obj.Top + (Controls.Item(i).Bottom))
                            IsChange = True
                        End If
                        '--< Fin Validar cambio de resolucion
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
                            '--> MinSize Controls
                            Controls.Item(i).MinWidth = GetMinWidthControl(obj)
                            Controls.Item(i).MinHeight = GetMinHeightControl(obj)
                            '--> LeftPercent
                            If Controls.Item(i).LeftPercent <> (obj.Left * 100) / GetControlScaleWidth(obj.Container) Then
                                If Not CBool(Controls.Item(i).UseModePercent) Then
                                    Controls.Item(i).LeftPercent = (obj.Left * 100) / GetControlScaleWidth(obj.Container)
                                    Controls.Item(i).LeftPercentStatic = 0
                                Else
                                    Controls.Item(i).LeftPercentStatic = obj.Left - (GetControlScaleWidth(obj.Container) * (Controls.Item(i).LeftPercent / 100))
                                End If
                            End If
                            '--> TopPercent
                            If Controls.Item(i).TopPercent <> (obj.Top * 100) / GetControlScaleHeight(obj.Container) Then
                                If Not CBool(Controls.Item(i).UseModePercent) Then
                                    Controls.Item(i).TopPercent = (obj.Top * 100) / GetControlScaleHeight(obj.Container)
                                    Controls.Item(i).TopPercentStatic = 0
                                Else
                                    Controls.Item(i).TopPercentStatic = obj.Top - (GetControlScaleHeight(obj.Container) * (Controls.Item(i).TopPercent / 100))
                                End If
                            End If
                            '--> WidthPercent
                            If Controls.Item(i).WidthPercent <> (obj.Width * 100) / GetControlScaleWidth(obj.Container) Then
                                If Not CBool(Controls.Item(i).UseModePercent) Then
                                    Controls.Item(i).WidthPercent = (obj.Width * 100) / GetControlScaleWidth(obj.Container)
                                    Controls.Item(i).RightPercentStatic = 0
                                Else
                                    Controls.Item(i).RightPercentStatic = (obj.Width + (IIf(Controls.Item(i).UseLeftPercent, Controls.Item(i).LeftPercentStatic, obj.Left))) - (GetControlScaleWidth(obj.Container) * (Controls.Item(i).WidthPercent / 100))
                                End If
                            End If
                            '--> HeightPercent
                            If Controls.Item(i).HeightPercent <> (obj.Height * 100) / GetControlScaleHeight(obj.Container) Then
                                If Not CBool(Controls.Item(i).UseModePercent) Then
                                    Controls.Item(i).HeightPercent = (obj.Height * 100) / GetControlScaleHeight(obj.Container)
                                    Controls.Item(i).BottomPercentStatic = 0
                                Else
                                    Controls.Item(i).BottomPercentStatic = (obj.Height + (IIf(Controls.Item(i).UseTopPercent, Controls.Item(i).TopPercentStatic, obj.Top))) - (GetControlScaleHeight(obj.Container) * (Controls.Item(i).HeightPercent / 100))
                                End If
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
    Dim objControl As Control
    Dim idx As String
    Dim sWidth As Long, sHeight As Long, a As Long, lTemp As Long
    Dim cRect As RECT
    'Resize Controls
    Call SendMessage(frmParent.hWnd, WM_SETREDRAW, 0&, 0&)
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
                                    If .UseModePercent Then
                                        'AQUIIIIII
                                        objControl.Left = sWidth * (.LeftPercent / 100) + .LeftPercentStatic
                                    Else
                                        objControl.Left = sWidth * (.LeftPercent / 100)
                                    End If
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
                                If .UseModePercent Then
                                    'AQUIIIIII
                                    objControl.Left = sWidth * (.LeftPercent / 100) + .LeftPercentStatic
                                Else
                                    objControl.Left = sWidth * (.LeftPercent / 100)
                                End If
                            End If
                            '--
                            'If .WidthPercent > 0 Then objControl.Width = sWidth * (.WidthPercent / 100)
                            If .UseWidthPercent Then
                                If .UseModePercent Then
                                    'AQUIIIIII
                                    objControl.Width = IIf(sWidth * (.WidthPercent / 100) + .RightPercentStatic - .LeftPercentStatic > 0, sWidth * (.WidthPercent / 100) + .RightPercentStatic - .LeftPercentStatic, 0)
                                Else
                                    objControl.Width = sWidth * (.WidthPercent / 100)
                                End If
                            End If
                        End If
                        If .AnchorBottom Then
                            If .AnchorTop Then
                                objControl.Height = IIf(sHeight - (.Top + .Bottom) > 0, sHeight - (.Top + .Bottom), 0)
                            Else
                                'If Not .TopPercent > 0 Then
                                If Not .UseTopPercent Then
                                    objControl.Top = sHeight - (objControl.Height + .Bottom)
                                Else
                                    If .UseModePercent Then
                                        'AQUIIIIII
                                        objControl.Top = sHeight * (.TopPercent / 100) + .TopPercentStatic
                                    Else
                                        objControl.Top = sHeight * (.TopPercent / 100)
                                    End If
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
                                If .UseModePercent Then
                                    'AQUIIIIII
                                    objControl.Top = sHeight * (.TopPercent / 100) + .TopPercentStatic
                                Else
                                    objControl.Top = sHeight * (.TopPercent / 100)
                                End If
                            End If
                            '--
                            'If .HeightPercent > 0 Then objControl.Height = sHeight * (.HeightPercent / 100)
                            If .UseHeightPercent Then
                                If .UseModePercent Then
                                    'AQUIIIIII
                                    objControl.Height = IIf(sHeight * (.HeightPercent / 100) + .BottomPercentStatic - .TopPercentStatic > 0, sHeight * (.HeightPercent / 100) + .BottomPercentStatic - .TopPercentStatic, 0)
                                Else
                                    objControl.Height = sHeight * (.HeightPercent / 100)
                                End If
                            End If
                        End If
                    End With
                End If
            End If
        Next
    Next
    '---
    Call SendMessage(frmParent.hWnd, WM_SETREDRAW, 1&, 0&)
    RedrawWindow frmParent.hWnd, ByVal &H0, 0, RDW_INVALIDATE Or RDW_ALLCHILDREN
    '--
End Sub

Private Sub SetIconForm(hWnd As Long, ByRef iconData() As Byte, ByVal cx As Long, ByVal cy As Long)
On Local Error GoTo RutinaError
    '--
    Dim gToken      As Long
    Dim hBitmap     As Long
    Dim IconPic     As StdPicture
    '--
    If Not IsArrayDim(VarPtrArray(iconData)) Then Exit Sub
    '--
    If iconData(2) = vbResIcon Or iconData(2) = vbResCursor Then
        Dim tIconHeader     As IconHeader
        Dim tIconEntry()    As IconEntry
        Dim MaxBitCount     As Long
        Dim MaxSize         As Long
        Dim Aproximate      As Long
        Dim IconID          As Long
        Dim i               As Long
        '--
        Call CopyMemory(tIconHeader, iconData(0), Len(tIconHeader))
        '--
        If tIconHeader.ihCount >= 1 Then
            ReDim tIconEntry(tIconHeader.ihCount - 1)
            Call CopyMemory(tIconEntry(0), iconData(Len(tIconHeader)), Len(tIconEntry(0)) * tIconHeader.ihCount)
            IconID = -1
            '--
            For i = 0 To tIconHeader.ihCount - 1
                If tIconEntry(i).ieBitCount > MaxBitCount Then MaxBitCount = tIconEntry(i).ieBitCount
            Next
            '--
            For i = 0 To tIconHeader.ihCount - 1
                If MaxBitCount = tIconEntry(i).ieBitCount Then
                    MaxSize = CLng(tIconEntry(i).ieWidth) + CLng(tIconEntry(i).ieHeight)
                    If MaxSize > Aproximate And MaxSize <= (cx + cy) Then
                        Aproximate = MaxSize
                        IconID = i
                    End If
                End If
            Next
            '--
            If IconID = -1 Then
                For i = 0 To tIconHeader.ihCount - 1
                    If MaxBitCount = tIconEntry(i).ieBitCount Then
                        If (tIconEntry(i).ieWidth) > 0 And (tIconEntry(i).ieHeight > 0) Then
                            IconID = i
                        End If
                    End If
                Next
            End If
            '--
            With tIconEntry(IconID)
                hIcon = CreateIconFromResourceEx(iconData(.ieImageOffset), .ieBytesInRes, 1, IconVersion, cx, cy, &H0)
                If hIcon <> 0 Then
                    SendMessage hWnd, WM_SETICON, ICON_JUMBO, ByVal hIcon
                    SendMessage hWnd, WM_SETICON, ICON_BIG, ByVal hIcon
                    SendMessage hWnd, WM_SETICON, ICON_SMALL, ByVal hIcon
                End If
            End With
        End If
    Else
        Dim TR          As RECTF
        Dim ResizeBmp   As Long
        Dim ResizeGra   As Long
        Dim IStream As IUnknown
        '--
        Call CreateStreamOnHGlobal(iconData(0), 0&, IStream)
        If Not IStream Is Nothing Then
            Dim GDIsi       As GDIPlusStartupInput
            '--
            GDIsi.GdiPlusVersion = 1&
            '--
            If GdiplusStartup(gToken, GDIsi) = 0 Then
                If GdipLoadImageFromStream(IStream, hBitmap) = 0 Then
                    Call GdipGetImageBounds(hBitmap, TR, UnitPixel)
                    '--
                    If cx <> TR.Width Or cy <> TR.Height Then
                        If GdipCreateBitmapFromScan0(cx, cy, 0&, PixelFormat32bppARGB, ByVal 0&, ResizeBmp) = 0 Then
                            If GdipGetImageGraphicsContext(ResizeBmp, ResizeGra) = 0 Then
                                GdipSetInterpolationMode ResizeGra, InterpolationModeHighQuality
                                If GdipDrawImageRect(ResizeGra, hBitmap, 0, 0, cx, cy) = 0 Then
                                    If GdipCreateHICONFromBitmap(ResizeBmp, hIcon) = 0 Then
                                        SendMessage hWnd, WM_SETICON, ICON_JUMBO, ByVal hIcon
                                        SendMessage hWnd, WM_SETICON, ICON_BIG, ByVal hIcon
                                        SendMessage hWnd, WM_SETICON, ICON_SMALL, ByVal hIcon
                                        'DestroyIcon hIcon
                                    End If
                                End If
                                Call GdipDeleteGraphics(ResizeGra)
                            End If
                            Call GdipDisposeImage(ResizeBmp)
                        End If
                    Else
                        If GdipCreateHICONFromBitmap(hBitmap, hIcon) = 0 Then
                            SendMessage hWnd, WM_SETICON, ICON_JUMBO, ByVal hIcon
                            SendMessage hWnd, WM_SETICON, ICON_BIG, ByVal hIcon
                            SendMessage hWnd, WM_SETICON, ICON_SMALL, ByVal hIcon
                            'DestroyIcon hIcon
                        End If
                    End If
                End If
                '--
                GdiplusShutdown gToken: gToken = 0
            End If
        End If
    End If
    '--
RutinaError:
    If gToken Then GdiplusShutdown gToken
    Debug.Print Err.Description
End Sub

Private Function IsArrayDim(ByVal lpArray As Long) As Boolean
    Dim lAddress As Long
    Call CopyMemory(lAddress, ByVal lpArray, &H4)
    IsArrayDim = Not (lAddress = 0)
End Function

Private Sub WndProc(ByVal bBefore As Boolean, _
                    ByRef bHandled As Boolean, _
                    ByRef lReturn As Long, _
                    ByVal hWnd As Long, _
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
