VERSION 5.00
Begin VB.UserControl ucProgressCircular 
   BackStyle       =   0  'Transparent
   ClientHeight    =   2130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2205
   ClipBehavior    =   0  'None
   PropertyPages   =   "ucProgressCircular.ctx":0000
   ScaleHeight     =   142
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   147
   Windowless      =   -1  'True
   Begin VB.Timer Timer1 
      Left            =   960
      Top             =   840
   End
End
Attribute VB_Name = "ucProgressCircular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------
'Module Name: ucProgressCircular
'Autor:  Leandro Ascierto
'Web: www.leandroascierto.com
'Date: 27/04/2020
'Version: 1.0.0
'-----------------------------------------------
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
'Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
'Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
'Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
'Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
'Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
'Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
'Private Declare Function CreateWindowExA Lib "user32.dll" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
'Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
'Private Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal HDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal HDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GdipCreatePen2 Lib "GdiPlus.dll" (ByVal mBrush As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipFlattenPath Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mMatrix As Long, ByVal mFlatness As Single) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipGetPointCount Lib "GdiPlus.dll" (ByVal mPath As Long, ByRef mCount As Long) As Long
Private Declare Function GdipRotatePathGradientTransform Lib "GdiPlus.dll" (ByVal mBrush As Long, ByVal mAngle As Single, ByVal mOrder As MatrixOrder) As Long
Private Declare Function GdipTranslatePathGradientTransform Lib "GdiPlus.dll" (ByVal mBrush As Long, ByVal mDx As Single, ByVal mDy As Single, ByVal mOrder As MatrixOrder) As Long
Private Declare Function GdipTranslateTextureTransform Lib "GdiPlus.dll" (ByVal mBrush As Long, ByVal mDx As Single, ByVal mDy As Single, ByVal mOrder As MatrixOrder) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal HDC As Long, ByRef graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mSmoothingMode As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal mBrush As Long) As Long
Private Declare Function GdipFillEllipseI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByVal mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipAddPathEllipseI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipSetPathGradientSurroundColorsWithCount Lib "GdiPlus.dll" (ByVal mBrush As Long, ByRef mColor As Long, ByRef mCount As Long) As Long
Private Declare Function GdipCreatePathGradientFromPath Lib "GdiPlus.dll" (ByVal mPath As Long, ByRef mPolyGradient As Long) As Long
Private Declare Function GdipSetPenEndCap Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mEndCap As LineCap) As Long
Private Declare Function GdipSetPenStartCap Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mStartCap As LineCap) As Long
Private Declare Function GdipDrawEllipse Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipSaveGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByRef mState As Long) As Long
Private Declare Function GdipRestoreGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mState As Long) As Long
Private Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "GdiPlus.dll" (ByRef mRect As RectL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mAngle As Single, ByVal mIsAngleScalable As Long, ByVal mWrapMode As WrapMode, ByRef mLineGradient As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal ARGB As Long, ByRef Brush As Long) As Long
Private Declare Function GdipCreateFont Lib "GdiPlus.dll" (ByVal mFontFamily As Long, ByVal mEmSize As Single, ByVal mStyle As Long, ByVal mUnit As Long, ByRef mFont As Long) As Long
Private Declare Function GdipDeleteFont Lib "GdiPlus.dll" (ByVal mFont As Long) As Long
Private Declare Function GdipDrawString Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RectF, ByVal mStringFormat As Long, ByVal mBrush As Long) As Long
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal Name As Long, ByVal fontCollection As Long, fontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As Long
Private Declare Function GdipGetGenericFontFamilySansSerif Lib "GdiPlus.dll" (ByRef mNativeFamily As Long) As Long
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As Long
Private Declare Function GdipSetStringFormatFlags Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mFlags As StringFormatFlags) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As StringAlignment) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mAlign As StringAlignment) As Long
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal Angle As Single, ByVal Order As Long) As Long
Private Declare Function GdipDeleteStringFormat Lib "GdiPlus.dll" (ByVal mFormat As Long) As Long
Private Declare Function GdipDrawArc Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal HDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal HDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal HDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "oleaut32" (lpPictDesc As PICTDESC, riid As Any, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal HDC As Long, ByRef lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal HDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Private Type PICTDESC
    lSize               As Long
    lType               As Long
    hBmp                As Long
    hPal                As Long
End Type

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type




Private Enum LineCap
    LineCapFlat = &H0
    LineCapSquare = &H1
    LineCapRound = &H2
    LineCapTriangle = &H3
    LineCapNoAnchor = &H10
    LineCapSquareAnchor = &H11
    LineCapRoundAnchor = &H12
    LineCapDiamondAnchor = &H13
    LineCapArrowAnchor = &H14
    LineCapCustom = &HFF
    LineCapAnchorMask = &HF0
End Enum

Private Enum MatrixOrder
    MatrixOrderPrepend = &H0
    MatrixOrderAppend = &H1
End Enum

Private Type RectF
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type

Private Type RectL
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Type GDIPlusStartupInput
    GdiPlusVersion              As Long
    DebugEventCallback          As Long
    SuppressBackgroundThread    As Long
    SuppressExternalCodecs      As Long
End Type

Public Enum CaptionAlignmentH
    cLeft
    cCenter
    cRight
End Enum

Public Enum CaptionAlignmentV
    cTop
    cMiddle
    cBottom
End Enum

Private Enum GDIPLUS_FONTSTYLE
    FontStyleRegular = 0
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum
  
Private Enum StringAlignment
    StringAlignmentNear = &H0
    StringAlignmentCenter = &H1
    StringAlignmentFar = &H2
End Enum

Private Enum StringFormatFlags
    StringFormatFlagsNone = &H0
    StringFormatFlagsDirectionRightToLeft = &H1
    StringFormatFlagsDirectionVertical = &H2
    StringFormatFlagsNoFitBlackBox = &H4
    StringFormatFlagsDisplayFormatControl = &H20
    StringFormatFlagsNoFontFallback = &H400
    StringFormatFlagsMeasureTrailingSpaces = &H800
    StringFormatFlagsNoWrap = &H1000
    StringFormatFlagsLineLimit = &H2000
    StringFormatFlagsNoClip = &H4000
End Enum

Private Enum WrapMode
    WrapModeTile = &H0
    WrapModeTileFlipX = &H1
    WrapModeTileFlipy = &H2
    WrapModeTileFlipXY = &H3
    WrapModeClamp = &H4
End Enum

Private Const GWL_WNDPROC               As Long = -4
Private Const GW_OWNER                  As Long = 4
Private Const WS_CHILD                  As Long = &H40000000
Private Const UnitPixel                 As Long = &H2&
Private Const LOGPIXELSX                As Long = 88
Private Const LOGPIXELSY                As Long = 90
Private Const SmoothingModeAntiAlias    As Long = 4
Private Const GDIP_OK                   As Long = &H0

Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Public Event PrePaint(hdc As Long)
'Public Event PostPaint(ByVal hdc As Long)
'Public Event KeyPress(KeyAscii As Integer)
'Public Event KeyUp(KeyCode As Integer, Shift As Integer)
'Public Event KeyDown(KeyCode As Integer, Shift As Integer)
'Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
'Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

Dim c_lhWnd As Long
Dim nScale As Single
Dim StartAngleAnimation As Single
Dim PF_ColorsCount As Long
Dim m_PenGradient As Long
Dim m_Caption1 As String
Dim m_Caption1_ForeColor As OLE_COLOR
Dim m_Caption1_Font As StdFont
Dim m_Caption1_OffsetY As Long
Dim m_Caption2 As String
Dim m_Caption2_ForeColor As OLE_COLOR
Dim m_Caption2_Font As StdFont
Dim m_Caption2_OffsetY As Long
Dim m_StepSpaceSize   As Long
Dim m_PF_Width As Long
Dim m_PF_Steps As Long
Dim m_PB_Color1 As OLE_COLOR
Dim m_PB_Color1Opacity As Long
Dim m_PB_Color2 As OLE_COLOR
Dim m_PB_Color2Opacity As Long
Dim m_PB_ColorGradient As Boolean
Dim m_PB_Width  As Long
Dim m_PB_Steps As Long
Dim m_PB_Border  As Boolean
Dim m_PB_BorderColor  As OLE_COLOR
Dim m_PB_BorderWidth As Long
Dim m_PB_BorderColorOpacity  As Long
Dim m_Min As Single
Dim m_Max As Single
Dim m_Value As Single
Dim m_Angle As Single
Dim m_StartAngle As Single
Dim m_CenterGradient As Boolean
Dim m_GradientAngle As Single
Dim m_CenterColor1 As OLE_COLOR
Dim m_CenterColor1Opacity As Long
Dim m_CenterColor2 As OLE_COLOR
Dim m_CenterColor2Opacity As Long
Dim m_CenterVisible As Boolean
Dim m_RoundStartStyle As Boolean
Dim m_RoundEndStyle As Boolean
Dim m_DisplayInPercent As Boolean
Dim m_ShowAnimation As Boolean
Dim m_PF_ForeColor As OLE_COLOR
Dim m_PF_ForeColorOpacity  As Long
Dim m_AnimationInterval As Long
Dim m_PF_Colors()   As Long
Dim GdipToken As Long

Public Property Get Image(Optional ByVal Width As Long, Optional ByVal Height As Long) As IPicture
    Dim lDC As Long
    Dim TempDC As Long
    Dim hBmp As Long, OldBmp As Long
    Dim hBrush As Long
    Dim Rect As Rect
    Dim lColor As Long
    Dim uDesc As PICTDESC
    Dim aInput(3) As Long
    
    If Width = 0 Then Width = UserControl.ScaleWidth
    If Height = 0 Then Height = UserControl.ScaleHeight
    
    lDC = GetDC(0&)
    TempDC = CreateCompatibleDC(lDC)
    hBmp = CreateCompatibleBitmap(lDC, Width, Height)
    OldBmp = SelectObject(TempDC, hBmp)
    Rect.Right = Width
    Rect.Bottom = Height
    
    'cambie el color manualmente, puse blanco por la impresora.
    lColor = vbWhite 'm_BackColor
    'If (lColor And &H80000000) Then lColor = GetSysColor(lColor And &HFF&)
    
    hBrush = CreateSolidBrush(lColor)
    FillRect TempDC, Rect, hBrush
    DeleteObject hBrush
    
    Draw TempDC, Width, Height
    
    With uDesc
        .lSize = Len(uDesc)
        .lType = vbPicTypeBitmap
        .hBmp = hBmp
    End With
    
    aInput(0) = &H7BF80980
    aInput(1) = &H101ABF32
    aInput(2) = &HAA00BB8B
    aInput(3) = &HAB0C3000

    Call OleCreatePictureIndirect(uDesc, aInput(0), 1, Image)
    
    ReleaseDC 0&, lDC
    SelectObject TempDC, OldBmp
    DeleteDC TempDC
    
End Property

Public Property Get Caption1() As String
    Caption1 = m_Caption1
End Property

Public Property Let Caption1(ByVal New_Value As String)
    m_Caption1 = New_Value
    PropertyChanged "Caption1"
    Refresh
End Property

Public Property Get Caption1_ForeColor() As OLE_COLOR
    Caption1_ForeColor = m_Caption1_ForeColor
End Property

Public Property Let Caption1_ForeColor(ByVal New_Value As OLE_COLOR)
    m_Caption1_ForeColor = New_Value
    PropertyChanged "Caption1_ForeColor"
    Refresh
End Property

Public Property Get Caption1_Font() As StdFont
    Set Caption1_Font = m_Caption1_Font
End Property

Public Property Set Caption1_Font(New_Font As StdFont)
    With m_Caption1_Font
        .Name = New_Font.Name
        .Size = New_Font.Size
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
        .Weight = New_Font.Weight
        .Charset = New_Font.Charset
    End With
    PropertyChanged "Caption1_Font"
    Refresh
End Property

Public Property Get Caption1_OffsetY() As Long
    Caption1_OffsetY = m_Caption1_OffsetY
End Property

Public Property Let Caption1_OffsetY(ByVal New_Value As Long)
    m_Caption1_OffsetY = New_Value
    PropertyChanged "Caption1_OffsetY"
    Refresh
End Property

Public Property Get Caption2() As String
    Caption2 = m_Caption2
End Property

Public Property Let Caption2(ByVal New_Value As String)
    m_Caption2 = New_Value
    PropertyChanged "Caption2"
    Refresh
End Property

Public Property Get Caption2_ForeColor() As OLE_COLOR
    Caption2_ForeColor = m_Caption2_ForeColor
End Property

Public Property Let Caption2_ForeColor(ByVal New_Value As OLE_COLOR)
    m_Caption2_ForeColor = New_Value
    PropertyChanged "Caption2_ForeColor"
    Refresh
End Property

Public Property Get Caption2_Font() As StdFont
    Set Caption2_Font = m_Caption2_Font
End Property

Public Property Set Caption2_Font(New_Font As StdFont)
    With m_Caption2_Font
        .Name = New_Font.Name
        .Size = New_Font.Size
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
        .Weight = New_Font.Weight
        .Charset = New_Font.Charset
    End With

    PropertyChanged "Caption2_Font"
    Refresh
End Property

Public Property Get Caption2_OffsetY() As Long
    Caption2_OffsetY = m_Caption2_OffsetY
End Property

Public Property Let Caption2_OffsetY(ByVal New_Value As Long)
    m_Caption2_OffsetY = New_Value
    PropertyChanged "Caption2_OffsetY"
    Refresh
End Property

Public Property Get StepSpaceSize() As Long
    StepSpaceSize = m_StepSpaceSize
End Property

Public Property Let StepSpaceSize(ByVal New_Value As Long)
    m_StepSpaceSize = New_Value
    PropertyChanged "StepSpaceSize"
    CleanPen
    Refresh
End Property

Public Property Get PF_Width() As Long
    PF_Width = m_PF_Width
End Property

Public Property Let PF_Width(ByVal New_Value As Long)
    m_PF_Width = SafeRange(New_Value, 1, 1000)
    PropertyChanged "PF_Width"
    CleanPen
    Refresh
End Property

Public Property Get PF_Steps() As Long
    PF_Steps = m_PF_Steps
End Property

Public Property Let PF_Steps(ByVal New_Value As Long)
    m_PF_Steps = SafeRange(New_Value, 1, 100)
    PropertyChanged "PF_Steps"
    CleanPen
    Refresh
End Property

Public Property Get PB_Color1() As OLE_COLOR
    PB_Color1 = m_PB_Color1
End Property

Public Property Let PB_Color1(ByVal New_Value As OLE_COLOR)
    m_PB_Color1 = New_Value
    PropertyChanged "PB_Color1"
    Refresh
End Property

Public Property Get PB_Color1Opacity() As Long
    PB_Color1Opacity = m_PB_Color1Opacity
End Property

Public Property Let PB_Color1Opacity(ByVal New_Value As Long)
    m_PB_Color1Opacity = SafeRange(New_Value, 0, 100)
    PropertyChanged "PB_Color1Opacity"
    Refresh
End Property

Public Property Get PB_Color2() As OLE_COLOR
    PB_Color2 = m_PB_Color2
End Property

Public Property Let PB_Color2(ByVal New_Value As OLE_COLOR)
    m_PB_Color2 = New_Value
    PropertyChanged "PB_Color2"
    Refresh
End Property

Public Property Get PB_Color2Opacity() As Long
    PB_Color2Opacity = m_PB_Color2Opacity
End Property

Public Property Let PB_Color2Opacity(ByVal New_Value As Long)
    m_PB_Color2Opacity = SafeRange(New_Value, 0, 100)
    PropertyChanged "PB_Color2Opacity"
    Refresh
End Property

Public Property Get PB_ColorGradient() As Boolean
    PB_ColorGradient = m_PB_ColorGradient
End Property

Public Property Let PB_ColorGradient(ByVal New_Value As Boolean)
    m_PB_ColorGradient = New_Value
    PropertyChanged "PB_ColorGradient"
    Refresh
End Property

Public Property Get PB_Width() As Long
    PB_Width = m_PB_Width
End Property

Public Property Let PB_Width(ByVal New_Value As Long)
    m_PB_Width = New_Value
    PropertyChanged "PB_Width"
    Refresh
End Property

Public Property Get PB_Steps() As Long
    PB_Steps = m_PB_Steps
End Property

Public Property Let PB_Steps(ByVal New_Value As Long)
    m_PB_Steps = SafeRange(New_Value, 1, 100)
    PropertyChanged "PB_Steps"
    Refresh
End Property

Public Property Get PB_Border() As Boolean
    PB_Border = m_PB_Border
End Property

Public Property Let PB_Border(ByVal New_Value As Boolean)
    m_PB_Border = New_Value
    PropertyChanged "PB_Border"
    Refresh
End Property

Public Property Get PB_BorderColor() As OLE_COLOR
    PB_BorderColor = m_PB_BorderColor
End Property

Public Property Let PB_BorderColor(ByVal New_Value As OLE_COLOR)
    m_PB_BorderColor = New_Value
    PropertyChanged "PB_BorderColor"
    Refresh
End Property

Public Property Get PB_BorderWidth() As Long
    PB_BorderWidth = m_PB_BorderWidth
End Property

Public Property Let PB_BorderWidth(ByVal New_Value As Long)
    m_PB_BorderWidth = SafeRange(New_Value, 1, 100)
    PropertyChanged "PB_BorderWidth"
    Refresh
End Property

Public Property Get PB_BorderColorOpacity() As Long
    PB_BorderColorOpacity = m_PB_BorderColorOpacity
End Property

Public Property Let PB_BorderColorOpacity(ByVal New_Value As Long)
    m_PB_BorderColorOpacity = SafeRange(New_Value, 0, 100)
    PropertyChanged "PB_BorderColorOpacity"
    Refresh
End Property

Public Property Get Min() As Single
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Value As Single)
    m_Min = New_Value
    PropertyChanged "Min"
    Refresh
End Property

Public Property Get Max() As Single
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Value As Single)
    m_Max = New_Value
    PropertyChanged "Max"
    Refresh
End Property

Public Property Get Value() As Single
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Single)
    m_Value = New_Value
    PropertyChanged "Value"
    Refresh
End Property

Public Property Get Angle() As Single
    Angle = m_Angle
End Property

Public Property Let Angle(ByVal New_Value As Single)
    m_Angle = New_Value
    PropertyChanged "Angle"
    CleanPen
    Refresh
End Property

Public Property Get StartAngle() As Single
    StartAngle = m_StartAngle
End Property

Public Property Let StartAngle(ByVal New_Value As Single)
    m_StartAngle = New_Value
    PropertyChanged "StartAngle"
    CleanPen
    Refresh
End Property

Public Property Get CenterGradient() As Boolean
    CenterGradient = m_CenterGradient
End Property

Public Property Let CenterGradient(ByVal New_Value As Boolean)
    m_CenterGradient = New_Value
    PropertyChanged "CenterGradient"
    Refresh
End Property

Public Property Get GradientAngle() As Single
    GradientAngle = m_GradientAngle
End Property

Public Property Let GradientAngle(ByVal New_Value As Single)
    m_GradientAngle = New_Value
    PropertyChanged "GradientAngle"
    Refresh
End Property

Public Property Get CenterColor1() As OLE_COLOR
    CenterColor1 = m_CenterColor1
End Property

Public Property Let CenterColor1(ByVal New_Value As OLE_COLOR)
    m_CenterColor1 = New_Value
    PropertyChanged "CenterColor1"
    Refresh
End Property

Public Property Get CenterColor1Opacity() As Long
    CenterColor1Opacity = m_CenterColor1Opacity
End Property

Public Property Let CenterColor1Opacity(ByVal New_Value As Long)
    m_CenterColor1Opacity = SafeRange(New_Value, 0, 100)
    PropertyChanged "CenterColor1Opacity"
    Refresh
End Property

Public Property Get CenterColor2() As OLE_COLOR
    CenterColor2 = m_CenterColor2
End Property

Public Property Let CenterColor2(ByVal New_Value As OLE_COLOR)
    m_CenterColor2 = New_Value
    PropertyChanged "CenterColor2"
    Refresh
End Property

Public Property Get CenterColor2Opacity() As Long
    CenterColor2Opacity = m_CenterColor2Opacity
End Property

Public Property Let CenterColor2Opacity(ByVal New_Value As Long)
    m_CenterColor2Opacity = SafeRange(New_Value, 0, 100)
    PropertyChanged "CenterColor2Opacity"
    Refresh
End Property

Public Property Get CenterVisible() As Boolean
    CenterVisible = m_CenterVisible
End Property

Public Property Let CenterVisible(ByVal New_Value As Boolean)
    m_CenterVisible = New_Value
    PropertyChanged "CenterVisible"
    Refresh
End Property

Public Property Get RoundStartStyle() As Boolean
    RoundStartStyle = m_RoundStartStyle
End Property

Public Property Let RoundStartStyle(ByVal New_Value As Boolean)
    m_RoundStartStyle = New_Value
    PropertyChanged "RoundStartStyle"
    CleanPen
    Refresh
End Property

Public Property Get RoundEndStyle() As Boolean
    RoundEndStyle = m_RoundEndStyle
End Property

Public Property Let RoundEndStyle(ByVal New_Value As Boolean)
    m_RoundEndStyle = New_Value
    PropertyChanged "RoundEndStyle"
    CleanPen
    Refresh
End Property

Public Property Get DisplayInPercent() As Boolean
    DisplayInPercent = m_DisplayInPercent
End Property

Public Property Let DisplayInPercent(ByVal New_Value As Boolean)
    m_DisplayInPercent = New_Value
    PropertyChanged "DisplayInPercent"
    Refresh
End Property

Public Property Get ShowAnimation() As Boolean
    ShowAnimation = m_ShowAnimation
End Property

Public Property Let ShowAnimation(ByVal New_Value As Boolean)
    m_ShowAnimation = New_Value
    If New_Value Then
        If Ambient.UserMode Then
            Timer1.Interval = m_AnimationInterval
        End If
    Else
        Timer1.Interval = 0
    End If
    PropertyChanged "ShowAnimation"
    CleanPen
    Refresh
End Property

Public Property Get PF_ForeColor() As OLE_COLOR
    PF_ForeColor = m_PF_ForeColor
End Property

Public Property Let PF_ForeColor(ByVal New_Value As OLE_COLOR)
    m_PF_ForeColor = New_Value
    m_PF_Colors(0) = RGBtoARGB(m_PF_ForeColor, m_PF_ForeColorOpacity)
    PropertyChanged "PF_ForeColor"
    CleanPen
    Refresh
End Property

Public Property Get PF_ForeColorOpacity() As Long
    PF_ForeColorOpacity = m_PF_ForeColorOpacity
End Property

Public Property Let PF_ForeColorOpacity(ByVal New_Value As Long)
    m_PF_ForeColorOpacity = SafeRange(New_Value, 0, 100)
    m_PF_Colors(0) = RGBtoARGB(m_PF_ForeColor, m_PF_ForeColorOpacity)
    PropertyChanged "PF_ForeColorOpacity"
    CleanPen
    Refresh
End Property

Public Property Get AnimationInterval() As Long
    AnimationInterval = m_AnimationInterval
End Property

Public Property Let AnimationInterval(ByVal New_Value As Long)
    m_AnimationInterval = SafeRange(New_Value, 1, 1000)
    PropertyChanged "AnimationInterval"
    If m_ShowAnimation Then
        If Ambient.UserMode Then
            Timer1.Interval = m_AnimationInterval
        End If
    End If
End Property

Public Function AddPaletteColors(lPalette() As Long)
    PF_ColorsCount = UBound(lPalette) + 1
    m_PF_Colors = lPalette
    PropertyChanged "PF_ColorsCount"
    PropertyChanged "PF_Colors"
    CleanPen
    Refresh
End Function

Public Function GetPaletteColors() As Long()
    GetPaletteColors = m_PF_Colors
End Function
Private Sub Timer1_Timer()
    If m_PF_Steps > 1 Then
        StartAngleAnimation = StartAngleAnimation + m_Angle / m_PF_Steps
    Else
        StartAngleAnimation = StartAngleAnimation + 36
    End If
    If StartAngleAnimation >= 360 Then StartAngleAnimation = 0
    Refresh
End Sub

Private Sub UserControl_Hide()
    Timer1.Interval = 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim bColors() As Byte
    
    c_lhWnd = UserControl.ContainerHwnd
  
    With PropBag
        m_Caption1 = .ReadProperty("Caption1", vbNullString)
        m_Caption1_ForeColor = .ReadProperty("Caption1_ForeColor", Ambient.ForeColor)
        Set m_Caption1_Font = .ReadProperty("Caption1_Font", Ambient.Font)
        m_Caption1_OffsetY = .ReadProperty("Caption1_OffsetY", 0)
        m_Caption2 = .ReadProperty("Caption2", vbNullString)
        m_Caption2_ForeColor = .ReadProperty("Caption2_ForeColor", vb3DShadow)
        Set m_Caption2_Font = .ReadProperty("Caption2_Font", Ambient.Font)
        m_Caption2_OffsetY = .ReadProperty("Caption2_OffsetY", 0)
        m_StepSpaceSize = .ReadProperty("StepSpaceSize", 4)
        m_PF_Width = .ReadProperty("PF_Width", 10)
        m_PF_Steps = .ReadProperty("PF_Steps", 1)
        m_PB_Color1 = .ReadProperty("PB_Color1", vbButtonShadow)
        m_PB_Color1Opacity = .ReadProperty("PB_Color1Opacity", 100)
        m_PB_Color2 = .ReadProperty("PB_Color2", vbButtonFace)
        m_PB_Color2Opacity = .ReadProperty("PB_Color2Opacity", 100)
        m_PB_ColorGradient = .ReadProperty("PB_ColorGradient", False)
        m_PB_Width = .ReadProperty("PB_Width", 10)
        m_PB_Steps = .ReadProperty("PB_Steps", 1)
        m_PB_Border = .ReadProperty("PB_Border", False)
        m_PB_BorderColor = .ReadProperty("PB_BorderColor", vbActiveBorder)
        m_PB_BorderWidth = .ReadProperty("PB_BorderWidth", 2)
        m_PB_BorderColorOpacity = .ReadProperty("PB_BorderColorOpacity", 100)
        m_Min = .ReadProperty("Min", 0)
        m_Max = .ReadProperty("Max", 100)
        m_Value = .ReadProperty("Value", 50)
        m_Angle = .ReadProperty("Angle", 360)
        m_StartAngle = .ReadProperty("StartAngle", 0)
        m_CenterGradient = .ReadProperty("CenterGradient", False)
        m_GradientAngle = .ReadProperty("GradientAngle", 45)
        m_CenterColor1 = .ReadProperty("CenterColor1", vbWindowBackground)
        m_CenterColor1Opacity = .ReadProperty("CenterColor1Opacity", 100)
        m_CenterColor2 = .ReadProperty("CenterColor2", vbButtonShadow)
        m_CenterColor2Opacity = .ReadProperty("CenterColor2Opacity", 100)
        m_CenterVisible = .ReadProperty("CenterVisible", False)
        m_RoundStartStyle = .ReadProperty("RoundStartStyle", False)
        m_RoundEndStyle = .ReadProperty("RoundEndStyle", False)
        m_DisplayInPercent = .ReadProperty("DisplayInPercent", True)
        m_ShowAnimation = .ReadProperty("ShowAnimation", False)
        m_PF_ForeColor = .ReadProperty("PF_ForeColor", vbHighlight)
        m_PF_ForeColorOpacity = .ReadProperty("PF_ForeColorOpacity", 100)
        m_AnimationInterval = .ReadProperty("AnimationInterval", 100)
        PF_ColorsCount = .ReadProperty("PF_ColorsCount", 1)
        
        ReDim m_PF_Colors(PF_ColorsCount - 1)
        If PF_ColorsCount > 1 Then
            bColors() = .ReadProperty("PF_Colors", 0)
            CopyMemory m_PF_Colors(0), bColors(0), PF_ColorsCount * 4&
        Else
            m_PF_Colors(0) = RGBtoARGB(m_PF_ForeColor, m_PF_ForeColorOpacity)
        End If
        
    End With
End Sub

Private Sub UserControl_Resize()
    CleanPen
End Sub

Private Sub UserControl_Terminate()
    CleanPen
    Call GdiplusShutdown(GdipToken)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim bColors() As Byte
    
    With PropBag
        .WriteProperty "Caption1", m_Caption1, vbNullString
        .WriteProperty "Caption1_ForeColor", m_Caption1_ForeColor, Ambient.ForeColor
        .WriteProperty "Caption1_Font", m_Caption1_Font, 0
        .WriteProperty "Caption1_OffsetY", m_Caption1_OffsetY
        .WriteProperty "Caption2", m_Caption2, vbNullString
        .WriteProperty "Caption2_ForeColor", m_Caption2_ForeColor, vb3DShadow
        .WriteProperty "Caption2_Font", m_Caption2_Font, Ambient.Font
        .WriteProperty "Caption2_OffsetY", m_Caption2_OffsetY, 0
        .WriteProperty "StepSpaceSize", m_StepSpaceSize, 4
        .WriteProperty "PF_Width", m_PF_Width, 10
        .WriteProperty "PF_Steps", m_PF_Steps, 1
        .WriteProperty "PB_Color1", m_PB_Color1, vbButtonShadow
        .WriteProperty "PB_Color1Opacity", m_PB_Color1Opacity, 100
        .WriteProperty "PB_Color2", m_PB_Color2, vbButtonFace
        .WriteProperty "PB_Color2Opacity", m_PB_Color2Opacity, 100
        .WriteProperty "PB_ColorGradient", m_PB_ColorGradient, False
        .WriteProperty "PB_Width", m_PB_Width, 10
        .WriteProperty "PB_Steps", m_PB_Steps, 1
        .WriteProperty "PB_Border", m_PB_Border, False
        .WriteProperty "PB_BorderColor", m_PB_BorderColor, vbActiveBorder
        .WriteProperty "PB_BorderWidth", m_PB_BorderWidth, 2
        .WriteProperty "PB_BorderColorOpacity", m_PB_BorderColorOpacity, 100
        .WriteProperty "Min", m_Min, 0
        .WriteProperty "Max", m_Max, 100
        .WriteProperty "Value", m_Value, 50
        .WriteProperty "Angle", m_Angle, 360
        .WriteProperty "StartAngle", m_StartAngle, 0
        .WriteProperty "CenterGradient", m_CenterGradient, False
        .WriteProperty "GradientAngle", m_GradientAngle, 45
        .WriteProperty "CenterColor1", m_CenterColor1, vbWindowBackground
        .WriteProperty "CenterColor1Opacity", m_CenterColor1Opacity, 100
        .WriteProperty "CenterColor2", m_CenterColor2, vbButtonShadow
        .WriteProperty "CenterColor2Opacity", m_CenterColor2Opacity, 100
        .WriteProperty "CenterVisible", m_CenterVisible, False
        .WriteProperty "RoundStartStyle", m_RoundStartStyle, False
        .WriteProperty "RoundEndStyle", m_RoundEndStyle, False
        .WriteProperty "DisplayInPercent", m_DisplayInPercent, True
        .WriteProperty "ShowAnimation", m_ShowAnimation, False
        .WriteProperty "PF_ForeColor", m_PF_ForeColor, vbHighlight
        .WriteProperty "PF_ForeColorOpacity", m_PF_ForeColorOpacity, 100
        .WriteProperty "AnimationInterval", m_AnimationInterval
        .WriteProperty "PF_ColorsCount", PF_ColorsCount, 1
        
        If PF_ColorsCount > 1 Then
            ReDim bColors(PF_ColorsCount * 4)
            CopyMemory bColors(0), m_PF_Colors(0), PF_ColorsCount * 4&
            .WriteProperty "PF_Colors", bColors, 0
        End If
        CleanPen
    End With
End Sub

Private Sub UserControl_InitProperties()
    m_Caption1 = vbNullString
    m_Caption1_ForeColor = Ambient.ForeColor
    Set m_Caption1_Font = Ambient.Font
    m_Caption1_Font.Size = 14
    m_Caption1_OffsetY = 0
    m_Caption2 = vbNullString
    m_Caption2_ForeColor = vb3DShadow
    Set m_Caption2_Font = Ambient.Font
    m_Caption2_OffsetY = 0
    m_StepSpaceSize = 4
    m_PF_Width = 10
    m_PF_Steps = 1
    m_PB_Color1 = vbButtonShadow
    m_PB_Color1Opacity = 100
    m_PB_Color2 = vbButtonFace
    m_PB_Color2Opacity = 100
    m_PB_ColorGradient = False
    m_PB_Steps = 1
    m_PB_Border = False
    m_PB_BorderColor = vbActiveBorder
    m_PB_BorderColorOpacity = 100
    m_PB_BorderWidth = 2
    m_PB_Width = 10
    m_Min = 0
    m_Max = 100
    m_Value = 50
    m_Angle = 360
    m_StartAngle = 0
    m_CenterGradient = False
    m_GradientAngle = 45
    m_CenterColor1 = vbWindowBackground
    m_CenterColor1Opacity = 100
    m_CenterColor2 = vbButtonShadow
    m_CenterColor2Opacity = 100
    m_DisplayInPercent = True
    m_PF_ForeColor = vbHighlight
    m_PF_ForeColorOpacity = 100
    m_AnimationInterval = 100
    PF_ColorsCount = 1
    
    ReDim m_PF_Colors(0)
    m_PF_Colors(0) = RGBtoARGB(m_PF_ForeColor, m_PF_ForeColorOpacity)

    c_lhWnd = UserControl.ContainerHwnd

End Sub


Public Function DrawGradientArc(ByVal hGraphics As Long, _
                                ByRef lColors() As Long, _
                                ByVal X As Single, ByVal Y As Single, _
                                ByVal Width As Single, ByVal Height As Single, _
                                ByVal PenWidth As Single, ByVal ColorsAngle As Long, _
                                ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Boolean
  

    Dim mPath As Long
    Dim hBrush As Long
    Dim hPen As Long
    Dim pColors() As Long
    Dim lCount As Long
    Dim num_pts  As Long
    Dim Index As Integer
    Dim i As Integer, j As Integer, N As Integer, t As Integer
    Dim PenM As Single
    Dim hState As Long
    
    If UBound(lColors) = 0 Then
        
        If GdipCreatePen1(lColors(0), PenWidth, UnitPixel, hPen) = GDIP_OK Then
            If RoundStartStyle Then GdipSetPenStartCap hPen, LineCapRound
            If RoundEndStyle Then GdipSetPenEndCap hPen, LineCapRound
            GdipDrawArc hGraphics, hPen, X, Y, Width, Height, mStartAngle, mSweepAngle
            GdipDeletePen hPen
        End If
        Exit Function
    End If
    
    
    
    If m_PenGradient = 0 Then
        PenM = PenWidth / 2
        
        
        If GdipCreatePath(&H0, mPath) <> GDIP_OK Then Exit Function
        If ColorsAngle <> 0 Then
            GdipAddPathEllipseI mPath, PenWidth, PenWidth, -(Width + PenWidth * 2), -(Height + PenWidth * 2)
        Else
            GdipAddPathEllipseI mPath, X - PenWidth, Y - PenWidth, Width + PenWidth * 2, Height + PenWidth * 2
        End If
        GdipFlattenPath mPath, 0, 0.1
        GdipCreatePathGradientFromPath mPath, hBrush
        GdipGetPointCount mPath, lCount
        GdipDeletePath mPath
        
        If hBrush = 0 Or Count = 0 Then Exit Function
        
        num_pts = (lCount - 1) / UBound(lColors)
        ReDim pColors(lCount - 1)
        
        For i = 0 To UBound(lColors) - 1
            If i < UBound(lColors) - 1 Then N = (i + 1) * num_pts Else N = lCount - 1
            t = N - Index
            For j = 0 To t
                pColors(Index) = ShiftColor(lColors(i + 1), lColors(i), j * 255& / t)
                Index = Index + 1
            Next
        Next
        
        GdipSetPathGradientSurroundColorsWithCount hBrush, pColors(0), lCount - 1
        
        If ColorsAngle <> 0 Then
            GdipRotatePathGradientTransform hBrush, ColorsAngle + 180, MatrixOrderAppend
            GdipTranslatePathGradientTransform hBrush, Width / 2, Height / 2, MatrixOrderPrepend
        End If
        
        If GdipCreatePen2(hBrush, PenWidth, UnitPixel, m_PenGradient) = GDIP_OK Then
            If RoundStartStyle Then GdipSetPenStartCap m_PenGradient, LineCapRound
            If RoundEndStyle Then GdipSetPenEndCap m_PenGradient, LineCapRound
            DrawGradientArc = True
            GdipDeleteBrush hBrush
        Else
            GdipDeleteBrush hBrush
            Exit Function
        End If
    End If
      
    If m_PenGradient = 0 Then Exit Function
      
    If ColorsAngle <> 0 Then
        'GdipSaveGraphics hGraphics, hState
        GdipTranslateWorldTransform hGraphics, X + Width / 2, Y + Height / 2, MatrixOrderPrepend
        GdipDrawArc hGraphics, m_PenGradient, -Width / 2, -Height / 2, Width, Height, mStartAngle, mSweepAngle
        GdipTranslateWorldTransform hGraphics, -(X + Width / 2), -(Y + Height / 2), MatrixOrderPrepend
        'GdipRestoreGraphics hGraphics, hState
    Else
        GdipDrawArc hGraphics, m_PenGradient, X, Y, Width, Height, mStartAngle, mSweepAngle
    End If
    
    
    

  
End Function


Private Function DrawText(ByVal hGraphics As Long, ByVal text As String, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal oFont As StdFont, ByVal ForeColor As Long, Optional HAlign As CaptionAlignmentH, Optional VAlign As CaptionAlignmentV, Optional bWordWrap As Boolean) As Long
    Dim hBrush As Long
    Dim hFontFamily As Long
    Dim hFormat As Long
    Dim layoutRect As RectF
    Dim lFontSize As Long
    Dim lFontStyle As GDIPLUS_FONTSTYLE
    Dim hFont As Long
    Dim HDC As Long

  
    If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), 0, hFontFamily) <> GDIP_OK Then
        If GdipGetGenericFontFamilySansSerif(hFontFamily) <> GDIP_OK Then Exit Function
        'If GdipGetGenericFontFamilySerif(hFontFamily) Then Exit Function
    End If
    
    If GdipCreateStringFormat(0, 0, hFormat) = GDIP_OK Then
        If Not bWordWrap Then GdipSetStringFormatFlags hFormat, StringFormatFlagsNoWrap
        'GdipSetStringFormatFlags hFormat, HotkeyPrefixShow
        GdipSetStringFormatAlign hFormat, HAlign
        GdipSetStringFormatLineAlign hFormat, VAlign
    End If
        
    If oFont.Bold Then lFontStyle = lFontStyle Or FontStyleBold
    If oFont.Italic Then lFontStyle = lFontStyle Or FontStyleItalic
    If oFont.Underline Then lFontStyle = lFontStyle Or FontStyleUnderline
    If oFont.Strikethrough Then lFontStyle = lFontStyle Or FontStyleStrikeout
        

    HDC = GetDC(0&)
    lFontSize = MulDiv(oFont.Size, GetDeviceCaps(HDC, LOGPIXELSY), 72)
    ReleaseDC 0&, HDC

    layoutRect.Left = X: layoutRect.Top = Y
    layoutRect.Width = Width: layoutRect.Height = Height

    If GdipCreateSolidFill(ForeColor, hBrush) = GDIP_OK Then
        If GdipCreateFont(hFontFamily, lFontSize, lFontStyle, UnitPixel, hFont) = GDIP_OK Then
            GdipDrawString hGraphics, StrPtr(text), -1, hFont, layoutRect, hFormat, hBrush
            GdipDeleteFont hFont
        End If
        GdipDeleteBrush hBrush
    End If
    
    If hFormat Then GdipDeleteStringFormat hFormat
    GdipDeleteFontFamily hFontFamily


End Function

Public Function GetWindowsDPI() As Double
    Dim HDC As Long, LPX  As Double
    HDC = GetDC(0)
    LPX = CDbl(GetDeviceCaps(HDC, LOGPIXELSX))
    ReleaseDC 0, HDC

    If (LPX = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = LPX / 96#
    End If
End Function

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    If UserControl.Enabled Then
        HitResult = vbHitResultHit
    End If
End Sub

Private Sub UserControl_Initialize()
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
    nScale = GetWindowsDPI
End Sub

'Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'    RaiseEvent KeyDown(KeyCode, Shift)
'End Sub
'
'Private Sub UserControl_KeyPress(KeyAscii As Integer)
'     RaiseEvent KeyPress(KeyAscii)
'End Sub
'
'Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
'    RaiseEvent KeyUp(KeyCode, Shift)
'End Sub
'
'Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
'    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
'End Sub
'
'Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
'    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
'End Sub

Private Sub UserControl_Paint()
    Draw UserControl.HDC, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

Public Sub Refresh()
    UserControl.Refresh
End Sub

Private Sub UserControl_Show()
    Me.Refresh
    If m_ShowAnimation Then
        If Ambient.UserMode Then
            Timer1.Interval = m_AnimationInterval
        End If
    End If
End Sub


Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Function SafeRange(Value, Min, Max)
    
    If Value < Min Then
        SafeRange = Min
    ElseIf Value > Max Then
        SafeRange = Max
    Else
        SafeRange = Value
    End If
End Function


Public Function RGBtoARGB(ByVal RGBColor As Long, ByVal Opacity As Long) As Long
    'By LaVople
    ' GDI+ color conversion routines. Most GDI+ functions require ARGB format vs standard RGB format
    ' This routine will return the passed RGBcolor to RGBA format
    ' Passing VB system color constants is allowed, i.e., vbButtonFace
    ' Pass Opacity as a value from 0 to 255

    If (RGBColor And &H80000000) Then RGBColor = GetSysColor(RGBColor And &HFF&)
    RGBtoARGB = (RGBColor And &HFF00&) Or (RGBColor And &HFF0000) \ &H10000 Or (RGBColor And &HFF) * &H10000
    Opacity = CByte((Abs(Opacity) / 100) * 255)
    If Opacity < 128 Then
        If Opacity < 0& Then Opacity = 0&
        RGBtoARGB = RGBtoARGB Or Opacity * &H1000000
    Else
        If Opacity > 255& Then Opacity = 255&
        RGBtoARGB = RGBtoARGB Or (Opacity - 128&) * &H1000000 Or &H80000000
    End If
    
End Function

Private Function ShiftColor(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal Proportion As Long) As Long
  
    Dim clrFore(3) As Byte
    Dim clrBack(3) As Byte
    Dim cResult(3) As Byte
    Dim R As Long
    
    clrFore(0) = (clrFirst And &HFF)
    clrFore(1) = (clrFirst And &HFF00&) \ &H100
    clrFore(2) = (clrFirst And &HFF0000) \ &H10000
    clrFore(3) = (clrFirst And &HFF000000) \ &H1000000 And &HFF

    clrBack(0) = (clrSecond And &HFF)
    clrBack(1) = (clrSecond And &HFF00&) \ &H100
    clrBack(2) = (clrSecond And &HFF0000) \ &H10000
    clrBack(3) = (clrSecond And &HFF000000) \ &H1000000 And &HFF

    R = (&HFF - Proportion)
    cResult(0) = (clrFore(0) * Proportion + clrBack(0) * R) / &HFF
    cResult(1) = (clrFore(1) * Proportion + clrBack(1) * R) / &HFF
    cResult(2) = (clrFore(2) * Proportion + clrBack(2) * R) / &HFF
    cResult(3) = (clrFore(3) * Proportion + clrBack(3) * R) / &HFF
    
    If cResult(3) < 128& Then
        ShiftColor = cResult(3) * &H1000000
    Else
        ShiftColor = (cResult(3) - 128&) * &H1000000 Or &H80000000
    End If
    ShiftColor = ShiftColor Or CLng(cResult(2)) * &H10000 Or CLng(cResult(1)) * &H100 Or cResult(0)

End Function


Private Sub Draw(HDC As Long, ScaleWidth As Long, ScaleHeight As Long)
    Dim Size As Long
    Dim RectL As RectL
    Dim X As Long, Y As Long
    Dim hPen As Long, hGraphics As Long
    Dim hBrush As Long
    Dim i As Single
    Dim iStep As Single
    Dim Percent As Single
    Dim Range As Single
    Dim Portion As Single
    Dim Rest As Single
    Dim sDiplay As String
    Dim StaAng As Single, FixPenCurve As Single
    Dim PenW As Long
    Dim PBW As Long, PBBW As Long, PFW As Long
    
    
    PBW = m_PB_Width * nScale
    PBBW = m_PB_BorderWidth * nScale
    PFW = m_PF_Width * nScale

    PenW = PBW
    If m_PB_Border Then PenW = PenW + PBBW
    If PFW > PenW Then PenW = PFW
   
    If ScaleWidth > ScaleHeight Then
        Size = ScaleHeight - PenW - 2
        X = ((ScaleWidth - ScaleHeight) / 2) + PenW / 2 + 1
        Y = PenW / 2 + 1
    Else
        Size = ScaleWidth - PenW - 2
        X = PenW / 2 + 1
        Y = ((ScaleHeight - ScaleWidth) / 2) + PenW / 2 + 1
    End If
    
    If GdipCreateFromHDC(HDC, hGraphics) <> GDIP_OK Then Exit Sub
    
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias


    With RectL
        .Left = X + PenW / 2 - 1
        .Top = Y + PenW / 2 - 1
        .Width = Size - PenW + 2
        .Height = Size - PenW + 2
        If m_CenterVisible Then
            If m_CenterGradient Then
                GdipCreateLineBrushFromRectWithAngleI RectL, _
                    RGBtoARGB(m_CenterColor1, m_CenterColor1Opacity), _
                    RGBtoARGB(m_CenterColor2, m_CenterColor2Opacity), _
                    90 - m_GradientAngle, 0, WrapModeTileFlipXY, hBrush
            Else
                GdipCreateSolidFill RGBtoARGB(m_CenterColor1, m_CenterColor1Opacity), hBrush
            End If
            If hBrush Then
                GdipFillEllipseI hGraphics, hBrush, .Left, .Top, .Width, .Height
                GdipDeleteBrush hBrush
            End If
        End If
    End With
    


    If m_PB_Border Then
        
        If GdipCreatePen1(RGBtoARGB(m_PB_BorderColor, m_PB_BorderColorOpacity), PBW + PBBW, UnitPixel, hPen) = GDIP_OK Then
            If RoundStartStyle Then GdipSetPenStartCap hPen, LineCapRound
            If RoundEndStyle Then GdipSetPenEndCap hPen, LineCapRound
    
            If m_PB_Steps = 1 Then
                GdipDrawArc hGraphics, hPen, X, Y, Size, Size, m_StartAngle - 90!, m_Angle
            Else
                iStep = m_Angle / m_PB_Steps
                For i = 0 To m_Angle - iStep + m_StepSpaceSize Step iStep
                    GdipDrawArc hGraphics, hPen, X, Y, Size, Size, m_StartAngle - 90! + i + (m_StepSpaceSize / 2) - (PBBW / (4 * nScale)), iStep - m_StepSpaceSize + PBBW / (2 * nScale)
                Next
            End If
            GdipDeletePen hPen
        End If
    End If
    
    If m_PB_ColorGradient Then
        GdipCreateLineBrushFromRectWithAngleI RectL, _
            RGBtoARGB(m_PB_Color1, m_PB_Color2Opacity), _
            RGBtoARGB(m_PB_Color2, m_PB_Color1Opacity), _
            90 - m_GradientAngle, 0, WrapModeTileFlipXY, hBrush
        If hBrush Then
            GdipCreatePen2 hBrush, PBW, UnitPixel, hPen
            GdipDeleteBrush hBrush
        End If
    Else
        GdipCreatePen1 RGBtoARGB(m_PB_Color1, m_PB_Color1Opacity), PBW, UnitPixel, hPen
    End If
    If hPen Then
        If RoundStartStyle Then GdipSetPenStartCap hPen, LineCapRound
        If RoundEndStyle Then GdipSetPenEndCap hPen, LineCapRound
        
        If m_PB_Steps = 1 Then
            GdipDrawArc hGraphics, hPen, X, Y, Size, Size, m_StartAngle - 90!, m_Angle
        Else
            iStep = m_Angle / m_PB_Steps
            For i = 0 To m_Angle - iStep + m_StepSpaceSize Step iStep
                GdipDrawArc hGraphics, hPen, X, Y, Size, Size, m_StartAngle - 90! + i + m_StepSpaceSize / 2, iStep - m_StepSpaceSize
            Next
        End If
        GdipDeletePen hPen
    End If
    
    Range = m_Max - m_Min
    Portion = m_Value - m_Min
    Percent = Portion * m_Angle / Range
    
    If m_ShowAnimation = False Then
        FixPenCurve = 0
        StaAng = m_StartAngle
    Else
        'Percent = 90
        CleanPen
        FixPenCurve = PFW / 4
        StaAng = StartAngleAnimation
    End If
    
    If m_PF_Steps = 1 Then
        DrawGradientArc hGraphics, m_PF_Colors, X, Y, Size, Size, PFW, StaAng - 90 - FixPenCurve, StaAng - 90, Percent
    Else
        iStep = m_Angle / m_PF_Steps
        Rest = Percent Mod iStep
        For i = 0 To Percent - Rest - 1 Step iStep
            DrawGradientArc hGraphics, m_PF_Colors, X, Y, Size, Size, PFW, m_StartAngle - 90, StaAng - 90! + i + m_StepSpaceSize / 2, iStep - m_StepSpaceSize
        Next

        If Rest Then
            If Rest > iStep - m_StepSpaceSize Then Rest = iStep - m_StepSpaceSize
            DrawGradientArc hGraphics, m_PF_Colors, X, Y, Size, Size, PFW, m_StartAngle - 90, StaAng - 90! + i + m_StepSpaceSize / 2, Rest
        End If
    End If
    
    If Len(m_Caption1) Or m_ShowAnimation Then
        sDiplay = m_Caption1
    Else
        If m_DisplayInPercent Then
            sDiplay = CStr(Portion * 100 \ Range) & "%"
        Else
            sDiplay = CLng(m_Value)
        End If
    End If
    
    With RectL
        DrawText hGraphics, sDiplay, .Left, (m_Caption1_OffsetY * nScale) + .Top, .Width, .Height, m_Caption1_Font, RGBtoARGB(m_Caption1_ForeColor, 100), cCenter, cMiddle, True
        DrawText hGraphics, m_Caption2, .Left, (m_Caption2_OffsetY * nScale) + .Top, .Width, .Height, m_Caption2_Font, RGBtoARGB(m_Caption2_ForeColor, 100), cCenter, cMiddle, True
    End With
    GdipDeleteGraphics hGraphics
End Sub

Private Function CleanPen()
    If m_PenGradient <> 0 Then
        GdipDeletePen m_PenGradient
        m_PenGradient = 0
    End If
End Function
