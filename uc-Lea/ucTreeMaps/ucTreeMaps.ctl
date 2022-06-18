VERSION 5.00
Begin VB.UserControl ucTreeMaps 
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipBehavior    =   0  'None
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Windowless      =   -1  'True
   Begin VB.Timer tmrMOUSEOVER 
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Left            =   1560
      Top             =   1560
   End
End
Attribute VB_Name = "ucTreeMaps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------
'Module Name: ucTreeMaps
'Autor:  Leandro Ascierto
'Web: www.leandroascierto.com
'Date: 17/08/2020
'Version: 1.0.0
'-----------------------------------------------
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTL) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTL) As Long
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
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal HDC As Long, ByRef graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mSmoothingMode As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal mBrush As Long) As Long
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
Private Declare Function GdipDeleteStringFormat Lib "GdiPlus.dll" (ByVal mFormat As Long) As Long
Private Declare Function GdipDrawArc Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)
Private Declare Function GdipDrawLine Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX1 As Single, ByVal mY1 As Single, ByVal mX2 As Single, ByVal mY2 As Single) As Long
Private Declare Function GdipAddPathLineI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipAddPathCurveI Lib "GdiPlus.dll" (ByVal mPath As Long, ByRef mPoints As POINTL, ByVal mCount As Long) As Long
Private Declare Function GdipSetCompositingQuality Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mCompositingQuality As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipAddPathCurve2I Lib "GdiPlus.dll" (ByVal mPath As Long, ByRef mPoints As POINTL, ByVal mCount As Long, ByVal mTension As Single) As Long
Private Declare Function GdipFillRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipDrawLineI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipDrawRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipAddPathLine2I Lib "GdiPlus.dll" (ByVal mPath As Long, ByRef mPoints As POINTL, ByVal mCount As Long) As Long
Private Declare Function GdipDrawLinesI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByRef mPoints As POINTL, ByVal mCount As Long) As Long
Private Declare Function GdipDrawCurveI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByRef mPoints As POINTL, ByVal mCount As Long) As Long
Private Declare Function GdipFillClosedCurveI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByRef mPoints As POINTL, ByVal mCount As Long) As Long
Private Declare Function GdipDrawEllipseI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipFillEllipseI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipCreateLineBrushI Lib "GdiPlus.dll" (ByRef mPoint1 As POINTL, ByRef mPoint2 As POINTL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mWrapMode As WrapMode, ByRef mLineGradient As Long) As Long
Private Declare Function GdipSetClipRectI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mCombineMode As CombineMode) As Long
Private Declare Function GdipResetClip Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipMeasureString Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RectF, ByVal mStringFormat As Long, ByRef mBoundingBox As RectF, ByRef mCodepointsFitted As Long, ByRef mLinesFilled As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipAddPathString Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFamily As Long, ByVal mStyle As Long, ByVal mEmSize As Single, ByRef mLayoutRect As RectF, ByVal mFormat As Long) As Long
Private Declare Function GdipClosePathFigure Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As Long) As Long
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal Angle As Single, ByVal Order As Long) As Long
Private Declare Function GdipResetWorldTransform Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Function TlsGetValue Lib "kernel32.dll" (ByVal dwTlsIndex As Long) As Long
Private Declare Function TlsSetValue Lib "kernel32.dll" (ByVal dwTlsIndex As Long, ByVal lpTlsValue As Long) As Long
Private Declare Function TlsFree Lib "kernel32.dll" (ByVal dwTlsIndex As Long) As Long
Private Declare Function TlsAlloc Lib "kernel32.dll" () As Long
Private Declare Function GdipSetStringFormatTrimming Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mTrimming As StringTrimming) As Long
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

Private Enum StringTrimming
    StringTrimmingNone = &H0
    StringTrimmingCharacter = &H1
    StringTrimmingWord = &H2
    StringTrimmingEllipsisCharacter = &H3
    StringTrimmingEllipsisWord = &H4
    StringTrimmingEllipsisPath = &H5
End Enum

Private Enum CombineMode
    CombineModeReplace = &H0
    CombineModeIntersect = &H1
    CombineModeUnion = &H2
    CombineModeXor = &H3
    CombineModeExclude = &H4
    CombineModeComplement = &H5
End Enum

Private Type POINTL
    X As Long
    Y As Long
End Type

Private Type POINTF
    X As Single
    Y As Single
End Type

Private Type SIZEF
    Width As Single
    Height As Single
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

Public Enum ucTM_TextAlignmentH
    cLeft
    cCenter
    cRight
End Enum

Private Enum TextAlignmentV
    cTop
    cMiddle
    cBottom
End Enum

Public Enum eLabelsPositions
    P_LeftTop
    P_LeftMiddle
    P_LeftBottom
    P_CenterTop
    P_CenterMiddle
    P_CenterBottom
    P_RightTop
    P_RightMiddle
    P_RightBottom
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
Private Const TLS_MINIMUM_AVAILABLE     As Long = 64

Public Event Click()
Public Event ItemClick(Key As Variant)
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

Public Enum ChartStyle
    CS_GroupedColumn
    CS_StackedBars
    CS_StackedBarsPercent
End Enum

Public Enum ChartOrientation
    CO_Vertical
    CO_Horizontal
End Enum

Public Enum ucTM_LegendAlign
    LA_LEFT
    LA_TOP
    LA_RIGHT
    LA_BOTTOM
End Enum

Private Type tSerie
    SerieName As String
    TextWidth As Long
    TextHeight As Long
    SeireColor As Long
    Labels As Collection
    IconsFonts As Collection
    keys As Collection
    CustomColors As Collection
    Values() As Single
    Rects() As RectF
    LegendRect As RectF
End Type

Dim c_lhWnd As Long
Dim nScale As Single
Dim m_Title As String
Dim m_BackColor As OLE_COLOR
Dim m_BackColorOpacity As Long
Dim m_ForeColor As OLE_COLOR
Dim m_LinesColor As OLE_COLOR
Dim m_FillOpacity As Long
Dim m_Border As Boolean
Dim m_BorderColor As OLE_COLOR
Dim m_LabelsVisible As Boolean
Dim m_LegendAlign As ucTM_LegendAlign
Dim m_LegendVisible As Boolean
Dim m_DrawTilteSerie As Boolean
Dim m_TitleFont As StdFont
Dim m_TitleForeColor As OLE_COLOR
Dim m_LabelsPositions As eLabelsPositions
Dim m_IconsPositions As eLabelsPositions
Dim m_LabelsFormats As String
Dim m_CornerRound As Long
Dim m_IconFont As StdFont
Dim m_ShowToolTips As Boolean
Dim m_Serie() As tSerie
Dim m_SeriesTotals() As Single
Dim m_SeriesRects() As RectF
Dim m_CursorPos As POINTF
Dim SerieCount As Long
Dim mHotSerie As Long
Dim mHotBar As Long
Dim hFontCollection As Long
Dim m_PT As POINTL
Dim m_Left As Long
Dim m_Top As Long
Dim GdipToken As Long
'*-

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
    
    lColor = m_BackColor
    If (lColor And &H80000000) Then lColor = GetSysColor(lColor And &HFF&)
    
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

Public Sub Clear()
    mHotBar = -1
    mHotSerie = -1
    Erase m_SeriesTotals
    Erase m_SeriesRects
    Erase m_Serie
    SerieCount = 0
    Me.Refresh
End Sub

Public Function AddLineSeries(ByVal SerieName As String, SerieColor As Long, cValues As Collection, Optional cLabels As Collection, Optional cIconsFonts As Collection, Optional cKeys As Collection, Optional cCustomColors As Collection)
    Dim i As Long, j As Long
    Dim vTemp As Variant
    Dim Total As Single
    Dim TempSerie As tSerie
    
    ReDim Preserve m_Serie(SerieCount)
    ReDim Preserve m_SeriesTotals(SerieCount)
 
    
    For i = 1 To cValues.Count - 1
        For j = i + 1 To cValues.Count
            If cValues(i) < cValues(j) Then
            
                vTemp = cValues(j)
                cValues.Remove j
                cValues.Add vTemp, , i
                
                If Not cLabels Is Nothing Then
                    vTemp = cLabels(j)
                    cLabels.Remove j
                    cLabels.Add vTemp, , i
                End If
                
                If Not cIconsFonts Is Nothing Then
                    vTemp = cIconsFonts(j)
                    cIconsFonts.Remove j
                    cIconsFonts.Add vTemp, , i
                End If
                
                If Not cKeys Is Nothing Then
                    vTemp = cKeys(j)
                    cKeys.Remove j
                    cKeys.Add vTemp, , i
                End If
                
                If Not cCustomColors Is Nothing Then
                    vTemp = cCustomColors(j)
                    cCustomColors.Remove j
                    cCustomColors.Add vTemp, , i
                End If
                
            End If
        Next j
    Next i
    
    With m_Serie(SerieCount)
        .SerieName = SerieName
        .SeireColor = SerieColor
        Set .Labels = cLabels
        Set .IconsFonts = cIconsFonts
        Set .keys = cKeys
        Set .CustomColors = cCustomColors
        ReDim .Values(cValues.Count - 1)

        For i = 1 To cValues.Count
            If cValues(i) <= 0 Then
                .Values(i - 1) = 0.0001
            Else
                .Values(i - 1) = cValues(i)
            End If
            m_SeriesTotals(SerieCount) = m_SeriesTotals(SerieCount) + cValues(i)
        Next
        If m_SeriesTotals(SerieCount) = 0 Then m_SeriesTotals(SerieCount) = 0.0001
    End With
    
    'Sort Series
    For i = 0 To SerieCount - 1
        If m_SeriesTotals(SerieCount) >= m_SeriesTotals(i) Then
            Total = m_SeriesTotals(SerieCount)
            TempSerie = m_Serie(SerieCount)
            For j = SerieCount To i + 1 Step -1
                m_SeriesTotals(j) = m_SeriesTotals(j - 1)
                m_Serie(j) = m_Serie(j - 1)
            Next
            m_SeriesTotals(i) = Total
            m_Serie(i) = TempSerie
        End If
    Next
    
    SerieCount = SerieCount + 1
    

End Function
 
Private Sub Timer1_Timer()
    Me.Refresh
    Timer1.Interval = 0
End Sub

Private Sub tmrMOUSEOVER_Timer()
    Dim PT As POINTL
    Dim Rect As RectL
  
    GetCursorPos PT
    ScreenToClient c_lhWnd, PT
 
    With Rect
        .Left = m_PT.X - (m_Left - ScaleX(Extender.Left, vbContainerSize, UserControl.ScaleMode))
        .Top = m_PT.Y - (m_Top - ScaleY(Extender.Top, vbContainerSize, UserControl.ScaleMode))
        .Width = UserControl.ScaleWidth
        .Height = UserControl.ScaleHeight
    End With

    If PtInRectL(Rect, PT.X, PT.Y) = 0 Then
        mHotBar = -1
        mHotSerie = -1
        tmrMOUSEOVER.Interval = 0
        UserControl.Refresh
        'RaiseEvent MouseLeave
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim i As Long, j As Long
    For i = 0 To SerieCount - 1
        With m_Serie(i)
            If PtInRectF(.LegendRect, X, Y) Then
                m_CursorPos.X = X: m_CursorPos.Y = Y
                If i <> mHotSerie Then
                    mHotSerie = i
                    mHotBar = -1
                    If m_ShowToolTips Then
                        Timer1.Interval = 150
                    Else
                        Me.Refresh
                    End If
                End If
                Exit Sub
            End If

            For j = 0 To UBound(.Values)
                If PtInRectF(.Rects(j), X, Y) Then
                    m_CursorPos.X = X: m_CursorPos.Y = Y
                    If j <> mHotBar Then
                        mHotSerie = i
                        mHotBar = j
                        If m_ShowToolTips Then
                            Timer1.Interval = 150
                        Else
                            Me.Refresh
                        End If
                    Else
                        If mHotSerie <> i Then
                            mHotBar = -1
                        End If
                    End If
                    Exit Sub
                End If
            Next
        End With
    Next
    
    If mHotSerie <> -1 Then
        mHotBar = -1
        mHotSerie = -1
        Me.Refresh
    End If

End Sub

Private Function PtInRectF(Rect As RectF, X As Single, Y As Single) As Boolean
    With Rect
        If X >= .Left And X <= .Left + .Width And Y >= .Top And Y <= .Top + .Height Then
            PtInRectF = True
        End If
    End With
End Function

Private Function PtInRectL(Rect As RectL, ByVal X As Single, ByVal Y As Single) As Boolean
    With Rect
        If X >= .Left And X <= .Left + .Width And Y >= .Top And Y <= .Top + .Height Then
            PtInRectL = True
        End If
    End With
End Function

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Title(ByVal New_Value As String)
    m_Title = New_Value
    PropertyChanged "Title"
    Refresh
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_Value As OLE_COLOR)
    m_BackColor = New_Value
    PropertyChanged "BackColor"
    Refresh
End Property

Public Property Get BackColorOpacity() As Long
    BackColorOpacity = m_BackColorOpacity
End Property

Public Property Let BackColorOpacity(ByVal New_Value As Long)
    m_BackColorOpacity = New_Value
    PropertyChanged "BackColorOpacity"
    Refresh
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_Value As OLE_COLOR)
    m_ForeColor = New_Value
    PropertyChanged "ForeColor"
    Refresh
End Property

Public Property Get LinesColor() As OLE_COLOR
    LinesColor = m_LinesColor
End Property

Public Property Let LinesColor(ByVal New_Value As OLE_COLOR)
    m_LinesColor = New_Value
    PropertyChanged "LinesColor"
    Refresh
End Property

Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

Public Property Let Font(ByVal New_Value As StdFont)
    Set UserControl.Font = New_Value
    PropertyChanged "Font"
    Refresh
End Property

Public Property Set Font(New_Font As StdFont)
    With UserControl.Font
        .Name = New_Font.Name
        .Size = New_Font.Size
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
        .Weight = New_Font.Weight
        .Charset = New_Font.Charset
    End With
    PropertyChanged "Font"
    Refresh
End Property

Public Property Get FillOpacity() As Long
    FillOpacity = m_FillOpacity
End Property

Public Property Let FillOpacity(ByVal New_Value As Long)
    m_FillOpacity = New_Value
    PropertyChanged "FillOpacity"
    Refresh
End Property

Public Property Get Border() As Boolean
    Border = m_Border
End Property

Public Property Let Border(ByVal New_Value As Boolean)
    m_Border = New_Value
    PropertyChanged "Border"
    Refresh
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_Value As OLE_COLOR)
    m_BorderColor = New_Value
    PropertyChanged "BorderColor"
    Refresh
End Property

Public Property Get LabelsVisible() As Boolean
    LabelsVisible = m_LabelsVisible
End Property

Public Property Let LabelsVisible(ByVal New_Value As Boolean)
    m_LabelsVisible = New_Value
    PropertyChanged "LabelsVisible"
    Refresh
End Property

Public Property Get LegendAlign() As ucTM_LegendAlign
    LegendAlign = m_LegendAlign
End Property

Public Property Let LegendAlign(ByVal New_Value As ucTM_LegendAlign)
    m_LegendAlign = New_Value
    PropertyChanged "LegendAlign"
    Refresh
End Property

Public Property Get LegendVisible() As Boolean
    LegendVisible = m_LegendVisible
End Property

Public Property Let LegendVisible(ByVal New_Value As Boolean)
    m_LegendVisible = New_Value
    PropertyChanged "LegendVisible"
    Refresh
End Property

Public Property Get DrawTilteSerie() As Boolean
    DrawTilteSerie = m_DrawTilteSerie
End Property

Public Property Let DrawTilteSerie(ByVal New_Value As Boolean)
    m_DrawTilteSerie = New_Value
    PropertyChanged "DrawTilteSerie"
    Refresh
End Property

Public Property Get TitleFont() As StdFont
    Set TitleFont = m_TitleFont
End Property

Public Property Let TitleFont(ByVal New_Value As StdFont)
    Set m_TitleFont = New_Value
    PropertyChanged "TitleFont"
    Refresh
End Property

Public Property Set TitleFont(New_Font As StdFont)
    With m_TitleFont
        .Name = New_Font.Name
        .Size = New_Font.Size
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
        .Weight = New_Font.Weight
        .Charset = New_Font.Charset
    End With
    PropertyChanged "TitleFont"
    Refresh
End Property

Public Property Get TitleForeColor() As OLE_COLOR
    TitleForeColor = m_TitleForeColor
End Property

Public Property Let TitleForeColor(ByVal New_Value As OLE_COLOR)
    m_TitleForeColor = New_Value
    PropertyChanged "TitleForeColor"
    Refresh
End Property

Public Property Get LabelsPositions() As eLabelsPositions
    LabelsPositions = m_LabelsPositions
End Property

Public Property Let LabelsPositions(ByVal New_Value As eLabelsPositions)
    m_LabelsPositions = New_Value
    PropertyChanged "LabelsPositions"
    Refresh
End Property

Public Property Get IconsPositions() As eLabelsPositions
    IconsPositions = m_IconsPositions
End Property

Public Property Let IconsPositions(ByVal New_Value As eLabelsPositions)
    m_IconsPositions = New_Value
    PropertyChanged "IconsPositions"
    Refresh
End Property

Public Property Get LabelsFormats() As String
    LabelsFormats = m_LabelsFormats
End Property

Public Property Let LabelsFormats(ByVal New_Value As String)
    m_LabelsFormats = New_Value
    PropertyChanged "LabelsFormats"
    Refresh
End Property

Public Property Get CornerRound() As Long
    CornerRound = m_CornerRound
End Property

Public Property Let CornerRound(ByVal New_Value As Long)
    m_CornerRound = New_Value
    PropertyChanged "CornerRound"
    Refresh
End Property

Public Property Get IconFont() As StdFont
    Set IconFont = m_IconFont
End Property

Public Property Set IconFont(New_Font As StdFont)
    With m_IconFont
        .Name = New_Font.Name
        .Size = New_Font.Size
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
        .Weight = New_Font.Weight
        .Charset = New_Font.Charset
    End With
    PropertyChanged "IconFont"
    Refresh
End Property


Public Property Get ShowToolTips() As Boolean
    ShowToolTips = m_ShowToolTips
End Property

Public Property Let ShowToolTips(ByVal New_Value As Boolean)
    m_ShowToolTips = New_Value
    PropertyChanged "ShowToolTips"
    Refresh
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long, j As Long
       
    For i = 0 To SerieCount - 1
        With m_Serie(i)
            If Not m_Serie(i).keys Is Nothing Then
                For j = 0 To UBound(.Values)
                    If PtInRectF(.Rects(j), X, Y) Then
                        RaiseEvent ItemClick(m_Serie(i).keys(j + 1))
                        Exit Sub
                    End If
                Next
            End If
        End With
    Next
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    hFontCollection = ReadValue(&HFC)
    c_lhWnd = UserControl.ContainerHwnd

    
    With PropBag
        m_Title = .ReadProperty("Title", Ambient.DisplayName)
        m_BackColor = .ReadProperty("BackColor", vbWhite)
        m_BackColorOpacity = .ReadProperty("BackColorOpacity", 100)
        m_ForeColor = .ReadProperty("ForeColor", vbBlack)
        m_LinesColor = .ReadProperty("LinesColor", vbWhite)
        Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
        m_FillOpacity = .ReadProperty("FillOpacity", 50)
        m_Border = .ReadProperty("Border", False)
        m_BorderColor = .ReadProperty("BorderColor", &HF2F2F2)
        m_LabelsVisible = .ReadProperty("LabelsVisible", True)
        m_LegendAlign = .ReadProperty("LegendAlign", LA_RIGHT)
        m_LegendVisible = .ReadProperty("LegendVisible", True)
        m_DrawTilteSerie = .ReadProperty("DrawTilteSerie", False)
        Set m_TitleFont = .ReadProperty("TitleFont", Ambient.Font)
        m_TitleForeColor = .ReadProperty("TitleForeColor", Ambient.ForeColor)
        m_LabelsPositions = .ReadProperty("LabelsPositions", P_LeftTop)
        m_IconsPositions = .ReadProperty("IconsPositions", P_CenterMiddle)
        m_LabelsFormats = .ReadProperty("LabelsFormats", "{V}")
        m_CornerRound = .ReadProperty("CornerRound", 0)
        Set m_IconFont = .ReadProperty("IconFont", UserControl.Ambient.Font)
        m_ShowToolTips = .ReadProperty("ShowToolTips", True)
    End With
    
    
    If Not Ambient.UserMode Then Call Example
End Sub

Private Sub UserControl_Terminate()
    Call GdiplusShutdown(GdipToken)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Title", m_Title, Ambient.DisplayName
        .WriteProperty "BackColor", m_BackColor, vbWhite
        .WriteProperty "BackColorOpacity", m_BackColorOpacity, 100
        .WriteProperty "ForeColor", m_ForeColor, vbBlack
        .WriteProperty "LinesColor", m_LinesColor, vbWhite
        .WriteProperty "Font", UserControl.Font, Ambient.Font
        .WriteProperty "FillOpacity", m_FillOpacity, 50
        .WriteProperty "Border", m_Border, False
        .WriteProperty "BorderColor", m_BorderColor, &HF2F2F2
        .WriteProperty "LabelsVisible", m_LabelsVisible, True
        .WriteProperty "LegendAlign", m_LegendAlign, LA_RIGHT
        .WriteProperty "LegendVisible", m_LegendVisible, True
        .WriteProperty "DrawTilteSerie", m_DrawTilteSerie, False
        .WriteProperty "TitleFont", m_TitleFont, Ambient.Font
        .WriteProperty "TitleForeColor", m_TitleForeColor, Ambient.ForeColor
        .WriteProperty "LabelsPositions", m_LabelsPositions, P_LeftTop
        .WriteProperty "IconsPositions", m_IconsPositions, P_CenterMiddle
        .WriteProperty "LabelsFormats", m_LabelsFormats, "{V}"
        .WriteProperty "CornerRound", m_CornerRound, 0
        .WriteProperty "IconFont", m_IconFont, UserControl.Ambient.Font
        .WriteProperty "ShowToolTips", m_ShowToolTips, True
        
        
    End With
End Sub

Private Sub UserControl_InitProperties()
    m_Title = Ambient.DisplayName
    m_BackColor = vbWhite
    m_BackColorOpacity = 100
    m_ForeColor = vbBlack
    m_LinesColor = vbWhite
    Set UserControl.Font = Ambient.Font
    m_FillOpacity = 50
    m_Border = False
    m_BorderColor = &HF2F2F2
    'm_LinesWidth = 1
    m_LegendAlign = LA_RIGHT
    m_LegendVisible = True
    m_DrawTilteSerie = False
    m_TitleFont.Name = UserControl.Font.Name
    m_TitleFont.Size = UserControl.Font.Size + 8
    m_LabelsPositions = P_LeftTop
    m_IconsPositions = P_CenterMiddle
    m_LabelsFormats = "{V}"
    m_CornerRound = 0
    Set m_IconFont = UserControl.Ambient.Font
    m_ShowToolTips = True
    
    c_lhWnd = UserControl.ContainerHwnd

    hFontCollection = ReadValue(&HFC)
    If Not Ambient.UserMode Then Call Example
End Sub


Private Function GetTextSize(ByVal hGraphics As Long, ByVal text As String, ByVal Width As Long, ByVal Height As Long, ByVal oFont As StdFont, ByVal bWordWrap As Boolean, ByRef SZ As SIZEF) As Long
    Dim hBrush As Long
    Dim hFontFamily As Long
    Dim hFormat As Long
    Dim layoutRect As RectF
    Dim lFontSize As Long
    Dim lFontStyle As GDIPLUS_FONTSTYLE
    Dim hFont As Long
    Dim BB As RectF, CF As Long, LF As Long
    
    If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), 0, hFontFamily) Then
        If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
    End If
    
    If GdipCreateStringFormat(0, 0, hFormat) = 0 Then
        If Not bWordWrap Then GdipSetStringFormatFlags hFormat, StringFormatFlagsNoWrap
    End If
        
    If oFont.Bold Then lFontStyle = lFontStyle Or FontStyleBold
    If oFont.Italic Then lFontStyle = lFontStyle Or FontStyleItalic
    If oFont.Underline Then lFontStyle = lFontStyle Or FontStyleUnderline
    If oFont.Strikethrough Then lFontStyle = lFontStyle Or FontStyleStrikeout
        

    lFontSize = MulDiv(oFont.Size, GetDeviceCaps(UserControl.HDC, LOGPIXELSY), 72)

    layoutRect.Width = Width * nScale: layoutRect.Height = Height * nScale

    Call GdipCreateFont(hFontFamily, lFontSize, lFontStyle, UnitPixel, hFont)

    GdipMeasureString hGraphics, StrPtr(text), -1, hFont, layoutRect, hFormat, BB, CF, LF

    SZ.Width = BB.Width
    SZ.Height = BB.Height
    
    GdipDeleteFont hFont
    GdipDeleteStringFormat hFormat
    GdipDeleteFontFamily hFontFamily
End Function
  
Private Function DrawText(ByVal hGraphics As Long, text As Variant, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, oFont As StdFont, ByVal ForeColor As Long, Optional HAlign As Long, Optional VAlign As Long, Optional bWordWrap As Boolean, Optional Angle As Single) As Long
    Dim hBrush As Long
    Dim hFontFamily As Long
    Dim hFormat As Long
    Dim layoutRect As RectF
    Dim lFontSize As Long
    Dim lFontStyle As GDIPLUS_FONTSTYLE
    Dim hFont As Long
    Dim HDC As Long
    Dim W As Single, H As Single
    Dim hPath As Long
    
    W = Width
    H = Height
    If Angle <> 0 Then
        Select Case Angle
            Case Is <= 90: W = Width + Angle * (Height - Width) / 90
            Case Is < 180: W = Width + (180 - Angle) * (Height - Width) / 90
            Case Is < 270: W = Width + (Angle Mod 90) * (Height - Width) / 90
            Case Else: W = Width + (360 - Angle) * (Height - Width) / 90
         End Select
         
        X = X - ((W - Width) / 2)
        Width = W

        GdipTranslateWorldTransform hGraphics, X + Width / 2, Y + Height / 2, 0
        GdipRotateWorldTransform hGraphics, Angle, 0
        GdipTranslateWorldTransform hGraphics, -(X + Width / 2), -(Y + Height / 2), 0
    End If

    If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), 0, hFontFamily) Then
        If hFontCollection Then
            If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), hFontCollection, hFontFamily) Then
                If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
            End If
        Else
            If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
        End If
    End If
    
    If GdipCreateStringFormat(0, 0, hFormat) = 0 Then
        'If Not bWordWrap Then GdipSetStringFormatFlags hFormat, StringFormatFlagsNoWrap
        GdipSetStringFormatFlags hFormat, StringFormatFlagsNoWrap
        GdipSetStringFormatTrimming hFormat, StringTrimmingEllipsisWord
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

    If GdipCreateSolidFill(ForeColor, hBrush) = 0 Then
        If GdipCreateFont(hFontFamily, lFontSize, lFontStyle, UnitPixel, hFont) = 0 Then
        
            If GdipCreatePath(&H0, hPath) = 0 Then
               
                If VarType(text) = vbLong Then
                    GdipAddPathString hPath, StrPtr(ChrW2(text)), -1, hFontFamily, lFontStyle, lFontSize, layoutRect, hFormat
                    'GdipDrawString hGraphics, StrPtr(ChrW2(text)), -1, hFont, layoutRect, hFormat, hBrush
                Else
                    'GdipAddPathString hPath, StrPtr(text), -1, hFontFamily, lFontStyle, lFontSize, layoutRect, hFormat
                    GdipDrawString hGraphics, StrPtr(text), -1, hFont, layoutRect, hFormat, hBrush
                End If
                GdipFillPath hGraphics, hBrush, hPath
                GdipDeletePath hPath
            End If
        
            'GdipDrawString hGraphics, StrPtr(text), -1, hFont, layoutRect, hFormat, hBrush
            
            GdipDeleteFont hFont
        End If
        GdipDeleteBrush hBrush
    End If
    
    If hFormat Then GdipDeleteStringFormat hFormat
    GdipDeleteFontFamily hFontFamily
    If Angle <> 0 Then GdipResetWorldTransform hGraphics

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
        If Ambient.UserMode Then
            Dim PT As POINTL

            If tmrMOUSEOVER.Interval = 0 Then
                GetCursorPos PT
                ScreenToClient c_lhWnd, PT
                m_PT.X = PT.X - X
                m_PT.Y = PT.Y - Y
    
                m_Left = ScaleX(Extender.Left, vbContainerSize, UserControl.ScaleMode)
                m_Top = ScaleY(Extender.Top, vbContainerSize, UserControl.ScaleMode)
 
              
                tmrMOUSEOVER.Interval = 1
                'RaiseEvent MouseEnter
            End If

        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
    nScale = GetWindowsDPI
    Set m_TitleFont = New StdFont

    mHotBar = -1
    mHotSerie = -1
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

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Paint()
    Draw UserControl.HDC, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

Public Sub Refresh()
    UserControl.Refresh
End Sub

Private Sub UserControl_Show()
    Me.Refresh
End Sub


'Funcion para combinar dos colores
Private Function ShiftColor(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
  
    Dim clrFore(3)         As Byte
    Dim clrBack(3)         As Byte
  
    OleTranslateColor clrFirst, 0, VarPtr(clrFore(0))
    OleTranslateColor clrSecond, 0, VarPtr(clrBack(0))
  
    clrFore(0) = (clrFore(0) * lAlpha + clrBack(0) * (255 - lAlpha)) / 255
    clrFore(1) = (clrFore(1) * lAlpha + clrBack(1) * (255 - lAlpha)) / 255
    clrFore(2) = (clrFore(2) * lAlpha + clrBack(2) * (255 - lAlpha)) / 255
  
    CopyMemory ShiftColor, clrFore(0), 4
  
End Function

Private Function SafeRange(Value, Min, Max)
    
    If Value < Min Then
        SafeRange = Min
    ElseIf Value > Max Then
        SafeRange = Max
    Else
        SafeRange = Value
    End If
End Function


Public Function RGBtoARGB(ByVal RGBColor As Long, Optional ByVal Opacity As Long = 100) As Long
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

Private Sub GetMinMax(Values() As Single, Min As Single, Max As Single)
    Dim i As Long
    For i = 0 To UBound(Values)
       If Values(i) < Min Then Min = Values(i)
       If Values(i) > Max Then Max = Values(i)
    Next
End Sub
'*1
Private Sub Draw(HDC As Long, ScaleWidth As Long, ScaleHeight As Long)
    Dim hGraphics As Long, hPath As Long
    Dim hBrush As Long, hPen As Long
    Dim mRect As RectL
    Dim Min As Single, Max As Single
    Dim nVal As Single
    Dim i As Single, j As Long
    Dim mHeight As Single
    Dim mWidth As Single
    Dim mPenWidth As Single
    Dim MarginLeft As Single
    Dim MarginRight As Single
    Dim TopHeader As Single
    Dim Footer As Single
    Dim TextWidth As Single
    Dim TextHeight As Single
    Dim lForeColor As Long
    Dim LW As Long
    Dim lColor As Long
    Dim LabelsRect As RectL
    Dim PT16 As Single
    Dim TitleSize As SIZEF
    Dim sDisplay As String
    Dim TextMargin As Single
    Dim Color1 As Long, Color2 As Long
    Dim ColRow As Long
    Dim VAlign As Long, HAling As Long
    
    If GdipCreateFromHDC(HDC, hGraphics) Then Exit Sub
  
    Call GdipSetSmoothingMode(hGraphics, SmoothingModeAntiAlias)
    Call GdipSetCompositingQuality(hGraphics, &H3) 'CompositingQualityGammaCorrected
    
    
    PT16 = (ScaleWidth + ScaleHeight) * 2.5 / 100
    
 
    mPenWidth = 1 * nScale
    LW = 1 * nScale 'm_LinesWidth
    lForeColor = RGBtoARGB(m_ForeColor, 100)
        
    If m_LegendVisible Then
        For i = 0 To SerieCount - 1
            m_Serie(i).TextHeight = UserControl.TextHeight(m_Serie(i).SerieName) * 1.5
            m_Serie(i).TextWidth = UserControl.TextWidth(m_Serie(i).SerieName) * 1.5 + m_Serie(i).TextHeight
        Next
    End If

    If Len(m_Title) Then
        Call GetTextSize(hGraphics, m_Title, ScaleWidth, 0, m_TitleFont, True, TitleSize)
    End If
    
    MarginRight = PT16
    TopHeader = PT16 + TitleSize.Height
    MarginLeft = PT16
    Footer = PT16
    
    mWidth = ScaleWidth - MarginLeft - MarginRight
    mHeight = ScaleHeight - TopHeader - Footer
    
    If m_LegendVisible Then
        ColRow = 1
        Select Case m_LegendAlign
            Case LA_RIGHT, LA_LEFT
                With LabelsRect
                    TextWidth = 0
                    TextHeight = 0
                    For i = 0 To SerieCount - 1
                        If TextHeight + m_Serie(i).TextHeight > mHeight Then
                            .Width = .Width + TextWidth
                            ColRow = ColRow + 1
                            TextWidth = 0
                            TextHeight = 0
                        End If
    
                        TextHeight = TextHeight + m_Serie(i).TextHeight
                        .Height = .Height + m_Serie(i).TextHeight
    
                        If TextWidth < m_Serie(i).TextWidth Then
                            TextWidth = m_Serie(i).TextWidth '+ PT16
                        End If
                    Next
                    .Width = .Width + TextWidth
                    If m_LegendAlign = LA_LEFT Then
                        MarginLeft = MarginLeft + .Width
                    Else
                        MarginRight = MarginRight + .Width
                    End If
                    mWidth = mWidth - .Width
                End With
    
            Case LA_BOTTOM, LA_TOP
                With LabelsRect
             
                    .Height = m_Serie(0).TextHeight + PT16 / 2
                    TextWidth = 0
                    For i = 0 To SerieCount - 1
                        If TextWidth + m_Serie(i).TextWidth > mWidth Then
                            .Height = .Height + m_Serie(i).TextHeight
                            ColRow = ColRow + 1
                            TextWidth = 0
                        End If
                        TextWidth = TextWidth + m_Serie(i).TextWidth
                        .Width = .Width + m_Serie(i).TextWidth
                    Next
                    If m_LegendAlign = LA_TOP Then
                        TopHeader = TopHeader + .Height
                    End If
                    mHeight = mHeight - .Height
                End With
        End Select
    End If

    'Background
    If m_BackColorOpacity > 0 Then
        GdipCreateSolidFill RGBtoARGB(m_BackColor, m_BackColorOpacity), hBrush
        GdipFillRectangleI hGraphics, hBrush, 0, 0, ScaleWidth, ScaleHeight
        GdipDeleteBrush hBrush
    End If
    
    'Border
    If m_Border Then
        Call GdipCreatePen1(RGBtoARGB(m_BorderColor, 100), mPenWidth, &H2, hPen)
        GdipDrawRectangleI hGraphics, hPen, mPenWidth / 2, mPenWidth / 2, ScaleWidth - mPenWidth, ScaleHeight - mPenWidth
        GdipDeletePen hPen
    End If

    'HORIZONTAL LINES AND vertical axis
    If SerieCount Then
        Squarified m_SeriesTotals, m_SeriesRects, MarginLeft, TopHeader, mWidth, mHeight
    End If
    
    For i = 0 To SerieCount - 1
        With m_SeriesRects(i)
            If m_DrawTilteSerie Then
                TextHeight = UserControl.TextHeight(m_Serie(0).SerieName)
                Squarified m_Serie(i).Values, m_Serie(i).Rects, .Left, .Top + TextHeight, .Width, .Height - TextHeight
            Else
                Squarified m_Serie(i).Values, m_Serie(i).Rects, .Left, .Top, .Width, .Height
            End If
        End With
    Next
    
    For i = 0 To SerieCount - 1
    
        With m_Serie(i)
            If m_DrawTilteSerie Then
                lColor = ShiftColor(m_Serie(i).SeireColor, vbBlack, 200)
    
                With m_SeriesRects(i)
                    Dim RectF As RectF
                    TextHeight = UserControl.TextHeight(m_Serie(0).SerieName)
                    RectF = m_SeriesRects(i)
                    RectF.Height = TextHeight
                    RoundRect hGraphics, RectF, RGBtoARGB(lColor), RGBtoARGB(m_LinesColor), m_CornerRound * nScale
                    If IsDarkColor(lColor) Then lColor = vbWhite Else lColor = vbBlack
                    DrawText hGraphics, m_Serie(i).SerieName, .Left, .Top, .Width, TextHeight, UserControl.Font, RGBtoARGB(lColor, 100), cCenter, cMiddle, True
                End With
            End If
            GetMinMax .Values, Min, Max
            For j = 0 To UBound(.Values)
                If Not .CustomColors Is Nothing Then
                    lColor = .CustomColors(j + 1)
                    If (mHotSerie = i And mHotBar = -1) Or (mHotSerie = i And mHotBar = j) Then
                        lColor = ShiftColor(lColor, vbWhite, 130)
                    End If
                Else
                    Color1 = m_Serie(i).SeireColor
                    Color2 = ShiftColor(m_Serie(i).SeireColor, vbWhite, 50)
                    If (mHotSerie = i And mHotBar = -1) Or (mHotSerie = i And mHotBar = j) Then
                        Color1 = ShiftColor(Color1, vbWhite, 130)
                        Color2 = ShiftColor(Color2, vbWhite, 130)
                    End If
                    lColor = ShiftColor(Color1, Color2, m_Serie(i).Values(j) * 255 / Max)
                End If
      
                RoundRect hGraphics, .Rects(j), RGBtoARGB(lColor), RGBtoARGB(m_LinesColor), m_CornerRound * nScale
                
                With .Rects(j)
                    If IsDarkColor(lColor) Then lColor = vbWhite Else lColor = vbBlack
                    TextMargin = (m_CornerRound / 8 + 2) * nScale
                    
                    If Not m_Serie(i).IconsFonts Is Nothing Then
                        Dim mSize As Long
                        mSize = m_IconFont.Size
                        
                        m_IconFont.Size = IIf(.Height < .Width, .Height, .Width) / 3 / nScale
                        
                        GetTextAlign m_IconsPositions, VAlign, HAling
                        DrawText hGraphics, CLng(m_Serie(i).IconsFonts(j + 1)), .Left + TextMargin, .Top + TextMargin, .Width - TextMargin * 2, .Height - TextMargin * 2, m_IconFont, RGBtoARGB(lColor, 70), VAlign, HAling, False
                        
                        m_IconFont.Size = mSize
                    End If
                    
                    GetTextAlign m_LabelsPositions, VAlign, HAling
                    If m_Serie(i).Labels Is Nothing Then
                        DrawText hGraphics, m_Serie(i).Values(j), .Left + TextMargin, .Top + TextMargin, .Width - TextMargin * 2, .Height - TextMargin * 2, UserControl.Font, RGBtoARGB(lColor, 100), VAlign, HAling, True
                    Else
                        'Dim TextWidth As Single
                        TextWidth = UserControl.TextWidth(m_Serie(i).Labels(j + 1))
                        If TextWidth > .Width - TextMargin * 2 And TextWidth < .Height - TextMargin * 2 Then
                            DrawText hGraphics, m_Serie(i).Labels(j + 1), .Left + TextMargin, .Top + TextMargin, .Width - TextMargin * 2, .Height - TextMargin * 2, UserControl.Font, RGBtoARGB(lColor, 100), 0, 1, True, 270
                        Else
                            DrawText hGraphics, m_Serie(i).Labels(j + 1), .Left + TextMargin, .Top + TextMargin, .Width - TextMargin * 2, .Height - TextMargin * 2, UserControl.Font, RGBtoARGB(lColor, 100), VAlign, HAling, True
                        End If
                    End If
                    
                End With
            Next
        End With
        
        If m_LegendVisible Then
            Select Case m_LegendAlign
                Case LA_RIGHT, LA_LEFT
                    With LabelsRect
                        TextWidth = 0
                        
                        If .Left = 0 Then
                            TextHeight = 0
                            If m_LegendAlign = LA_LEFT Then
                                .Left = PT16
                            Else
                                .Left = MarginLeft + mWidth + PT16
                            End If
                            If ColRow = 1 Then
                                .Top = TopHeader + mHeight / 2 - .Height / 2
                            Else
                                .Top = TopHeader
                            End If
                        End If
                        
                        If TextWidth < m_Serie(i).TextWidth Then
                            TextWidth = m_Serie(i).TextWidth '+ PT16
                        End If
        
                        If TextHeight + m_Serie(i).TextHeight > mHeight Then
                             If i > 0 Then .Left = .Left + TextWidth
                            .Top = TopHeader
                             TextHeight = 0
                        End If
                        m_Serie(i).LegendRect.Left = .Left
                        m_Serie(i).LegendRect.Top = .Top
                        m_Serie(i).LegendRect.Width = m_Serie(i).TextWidth
                        m_Serie(i).LegendRect.Height = m_Serie(i).TextHeight
                        
                        GdipCreateSolidFill RGBtoARGB(m_Serie(i).SeireColor, 100), hBrush
                        GdipFillRectangleI hGraphics, hBrush, .Left, .Top + m_Serie(i).TextHeight / 4, m_Serie(i).TextHeight / 2, m_Serie(i).TextHeight / 2
                        GdipDeleteBrush hBrush
                        
                        DrawText hGraphics, m_Serie(i).SerieName, .Left + m_Serie(i).TextHeight / 1.5, .Top, m_Serie(i).TextWidth, m_Serie(i).TextHeight, UserControl.Font, lForeColor, cLeft, cMiddle
                        TextHeight = TextHeight + m_Serie(i).TextHeight
                        .Top = .Top + m_Serie(i).TextHeight
                        
                    End With
                    
                Case LA_BOTTOM, LA_TOP
                    With LabelsRect
                        If .Left = 0 Then
                            If ColRow = 1 Then
                                .Left = MarginLeft + mWidth / 2 - .Width / 2
                            Else
                                .Left = MarginLeft
                            End If
                            If m_LegendAlign = LA_TOP Then
                                .Top = PT16 + TitleSize.Height + PT16 / 4
                            Else
                                .Top = TopHeader + mHeight + PT16
                            End If
                        End If
        
                        If .Left + m_Serie(i).TextWidth - MarginLeft > mWidth Then
                            .Left = MarginLeft
                            .Top = .Top + m_Serie(i).TextHeight
                        End If
        
                        GdipCreateSolidFill RGBtoARGB(m_Serie(i).SeireColor, 100), hBrush
                        GdipFillRectangleI hGraphics, hBrush, .Left, .Top + m_Serie(i).TextHeight / 4, m_Serie(i).TextHeight / 2, m_Serie(i).TextHeight / 2
                        GdipDeleteBrush hBrush
                        m_Serie(i).LegendRect.Left = .Left
                        m_Serie(i).LegendRect.Top = .Top
                        m_Serie(i).LegendRect.Width = m_Serie(i).TextWidth
                        m_Serie(i).LegendRect.Height = m_Serie(i).TextHeight
                        
                        DrawText hGraphics, m_Serie(i).SerieName, .Left + m_Serie(i).TextHeight / 1.5, .Top, m_Serie(i).TextWidth, m_Serie(i).TextHeight, UserControl.Font, lForeColor, cLeft, cMiddle
                        .Left = .Left + m_Serie(i).TextWidth '+ m_Serie(i).TextHeight / 1.5
                    End With
            End Select
        End If

    Next
 
    'Title
    If Len(m_Title) Then
        DrawText hGraphics, m_Title, 0, PT16 / 2, ScaleWidth, TopHeader, m_TitleFont, RGBtoARGB(m_TitleForeColor, 100), cCenter, cTop, True
    End If


    If m_ShowToolTips Then fShowToolTips hGraphics
    
    Call GdipDeleteGraphics(hGraphics)
    

End Sub

Private Function GetTextAlign(LblPos As eLabelsPositions, HAlign As Long, VAlign As Long)
    Select Case LblPos
        Case P_LeftTop: HAlign = 0: VAlign = 0
        Case P_LeftMiddle: HAlign = 0: VAlign = 1
        Case P_LeftBottom: HAlign = 0: VAlign = 2
        Case P_CenterTop: HAlign = 1: VAlign = 0
        Case P_CenterMiddle: HAlign = 1: VAlign = 1
        Case P_CenterBottom: HAlign = 1: VAlign = 2
        Case P_RightTop: HAlign = 2: VAlign = 0
        Case P_RightMiddle: HAlign = 2: VAlign = 1
        Case P_RightBottom: HAlign = 2: VAlign = 2
    End Select
End Function

Private Sub fShowToolTips(hGraphics As Long)
    Dim i As Long, j As Long
    Dim sDisplay As String
    Dim bBold As Boolean
    Dim RectF As RectF
    Dim LW As Single
    Dim lForeColor As Long
    Dim TM As Single
    Dim SZ As SIZEF
    
    TM = UserControl.TextHeight("ÁJ") / 3
    lForeColor = RGBtoARGB(m_ForeColor, 100)
    LW = 1 * nScale
    
    If mHotSerie > -1 Then
       
        For i = 0 To SerieCount - 1
            For j = 0 To UBound(m_Serie(i).Values)
                
                Dim sText As String
                If mHotSerie = i Then
                    If mHotBar = j Then
                        sDisplay = Replace(m_LabelsFormats, "{V}", m_Serie(i).Values(j))
                        sDisplay = Replace(sDisplay, "{LF}", vbLf)
                        
                        If Not m_Serie(i).Labels Is Nothing Then
                            If Len(m_Serie(i).SerieName) Then
                                sText = m_Serie(i).SerieName & vbCrLf
                            End If
                            sText = sText & m_Serie(i).Labels(j + 1) & ": " & sDisplay
                        Else
                            sText = sText & m_Serie(i).SerieName & ": " & sDisplay
                        End If
                        
                        GetTextSize hGraphics, sText, 0, 0, UserControl.Font, False, SZ

                        With RectF
                          
                            .Left = m_CursorPos.X + 16 * nScale
                            .Top = m_CursorPos.Y + 16 * nScale
                            .Width = SZ.Width + TM * 2
                            .Height = SZ.Height + TM * 2
                            
                            If .Left < 0 Then .Left = LW
                            If .Top < 0 Then .Top = LW
                            
                            If .Left + .Width >= UserControl.ScaleWidth - LW Then .Left = UserControl.ScaleWidth - .Width - LW
                            If .Top + .Height >= UserControl.ScaleHeight - LW Then .Top = UserControl.ScaleHeight - .Height - LW
                            
                        End With
                        
                        RoundRect hGraphics, RectF, RGBtoARGB(vbWhite, 85), RGBtoARGB(m_Serie(i).SeireColor), TM

                        
                        
                        With RectF
                            .Left = .Left + TM
                            
                            
                            If Not m_Serie(i).Labels Is Nothing Then
                                
                                If Len(m_Serie(i).SerieName) Then
                                    .Top = .Top + TM
                                    bBold = UserControl.Font.Bold
                                    UserControl.Font.Bold = True
                                    DrawText hGraphics, m_Serie(i).SerieName, .Left, .Top, .Width, 0, UserControl.Font, RGBtoARGB(m_Serie(i).SeireColor, 100), cLeft, cTop
                                    UserControl.Font.Bold = False
                                    GetTextSize hGraphics, m_Serie(i).SerieName, 0, 0, UserControl.Font, False, SZ
                                    .Height = SZ.Height
                                    .Top = .Top + SZ.Height
                                End If
                                GetTextSize hGraphics, m_Serie(i).Labels(j + 1) & ": ", 0, 0, UserControl.Font, False, SZ
                                'TextWidth = UserControl.TextWidth(m_Serie(i).Labels(j + 1) & ": ")
                                DrawText hGraphics, m_Serie(i).Labels(j + 1) & ": ", .Left, .Top, .Width, .Height, UserControl.Font, lForeColor, cLeft, cMiddle
                                UserControl.Font.Bold = True
                                DrawText hGraphics, sDisplay, .Left + SZ.Width, .Top, .Width, .Height, UserControl.Font, lForeColor, cLeft, cMiddle
                                UserControl.Font.Bold = bBold
                            Else
                                GetTextSize hGraphics, m_Serie(i).SerieName & ": ", 0, 0, UserControl.Font, False, SZ
                                DrawText hGraphics, m_Serie(i).SerieName & ": ", .Left, .Top, .Width, .Height, UserControl.Font, lForeColor, cLeft, cMiddle
                                bBold = UserControl.Font.Bold
                                UserControl.Font.Bold = True
                                DrawText hGraphics, sDisplay, .Left + SZ.Width, .Top, .Width, .Height, UserControl.Font, lForeColor, cLeft, cMiddle
                                UserControl.Font.Bold = bBold
                            End If
                        End With
                    ElseIf mHotBar = -1 Then
                        
                        sDisplay = Replace(m_LabelsFormats, "{V}", m_SeriesTotals(i))
                        sDisplay = Replace(sDisplay, "{LF}", vbLf)
                        sText = m_Serie(i).SerieName & ": " & sDisplay

                        GetTextSize hGraphics, sText, 0, 0, UserControl.Font, False, SZ
                      
                        With RectF
                          
                            .Left = m_CursorPos.X + 16 * nScale
                            .Top = m_CursorPos.Y + 16 * nScale
                            .Width = SZ.Width + TM * 2
                            .Height = SZ.Height + TM * 2
                            
                            If .Left < 0 Then .Left = LW
                            If .Top < 0 Then .Top = LW
                            If .Left + .Width >= UserControl.ScaleWidth - LW Then .Left = UserControl.ScaleWidth - .Width - LW
                            If .Top + .Height >= UserControl.ScaleHeight - LW Then .Top = UserControl.ScaleHeight - .Height - LW
                        End With
                        
                        RoundRect hGraphics, RectF, RGBtoARGB(vbWhite, 85), RGBtoARGB(m_Serie(i).SeireColor), TM

                        With RectF
                            .Left = .Left + TM
                            GetTextSize hGraphics, m_Serie(i).SerieName & ": ", 0, 0, UserControl.Font, False, SZ
                            DrawText hGraphics, m_Serie(i).SerieName & ": ", .Left, .Top, .Width, .Height, UserControl.Font, lForeColor, cLeft, cMiddle
                            bBold = UserControl.Font.Bold
                            UserControl.Font.Bold = True
                            DrawText hGraphics, sDisplay, .Left + SZ.Width, .Top, .Width, .Height, UserControl.Font, lForeColor, cLeft, cMiddle
                            UserControl.Font.Bold = bBold
                           
                        End With
                    End If
                End If
            Next
        Next
    End If
End Sub


Private Sub Example()
    Dim Value As Collection
    Set Value = New Collection
    
    Value.Add "2018"
    Value.Add "2019"
    Value.Add "2020"

    
    Set Value = New Collection
    With Value
        .Add 10
        .Add 15
        .Add 5
    End With
    Me.AddLineSeries "Juan", vbRed, Value
    Set Value = New Collection
    With Value
        .Add 8
        .Add 4
        .Add 12
    End With
    Me.AddLineSeries "Pedro", vbBlue, Value
End Sub


Private Sub RoundRect(ByVal hGraphics As Long, Rect As RectF, ByVal BackColor As Long, ByVal BorderColor As Long, ByVal Round As Single)
    Dim hPen As Long, hBrush As Long
    Dim mPath As Long
    
    GdipCreateSolidFill BackColor, hBrush
    GdipCreatePen1 BorderColor, &H1 * nScale, &H2, hPen

    If Round = 0 Then
        With Rect
            GdipFillRectangleI hGraphics, hBrush, .Left, .Top, .Width, .Height
            GdipDrawRectangleI hGraphics, hPen, .Left, .Top, .Width, .Height
        End With
    Else
        If GdipCreatePath(&H0, mPath) = 0 Then
            Round = Round * 2
            With Rect
                GdipAddPathArcI mPath, .Left, .Top, Round, Round, 180, 90
                GdipAddPathArcI mPath, .Left + .Width - Round, .Top, Round, Round, 270, 90
                GdipAddPathArcI mPath, .Left + .Width - Round, .Top + .Height - Round, Round, Round, 0, 90
                GdipAddPathArcI mPath, .Left, .Top + .Height - Round, Round, Round, 90, 90
                GdipClosePathFigure mPath
            End With
            GdipFillPath hGraphics, hBrush, mPath
            GdipDrawPath hGraphics, hPen, mPath
            Call GdipDeletePath(mPath)
        End If
    End If
        
    Call GdipDeleteBrush(hBrush)
    Call GdipDeletePen(hPen)
    
End Sub


Private Sub Squarified(Values() As Single, Rects() As RectF, ByVal Left As Single, ByVal Top As Single, ByVal Width As Single, ByVal Height As Single)
    Dim Aspect  As Single, Aspect1 As Single, Aspect2 As Single, MinAsp As Single
    Dim A As Single, LastW As Single, LastH As Single
    Dim i As Long, j As Long
    Dim Sum As Single
    Dim P As Single
    Dim N As Single
    Dim Tot As Single
    Dim Cont As Long
    
    Dim X As Single, Y As Single
    Dim R As RectF
    Dim TempSum As Single
    Dim bDrawVert As Boolean
    
    ReDim Rects(UBound(Values))
    
    With R
        .Left = Left
        .Top = Top
        .Width = Width
        .Height = Height
    End With

    For i = 0 To UBound(Values)
        Tot = Tot + Values(i)
    Next
    
    For i = 0 To UBound(Values)
        Sum = 0
        Cont = 0
        MinAsp = 0
        TempSum = 0
        bDrawVert = DrawVertically(R.Width, R.Height)

        For j = i To UBound(Values)
            
            If bDrawVert Then
                Sum = Sum + Values(j)
                A = Sum / R.Height
                P = Tot / Sum
                Aspect1 = Max(P / A, A / P)
                If MinAsp = 0 Then
                    P = Tot / Sum
                Else
                    P = Sum / Values(j)
                End If
                Aspect2 = Max(P / A, A / P)
                Aspect = Max(Aspect1, Aspect2)
                If Aspect < MinAsp Or MinAsp = 0 Then
                    TempSum = TempSum + Values(j)
                    MinAsp = Aspect
                    Cont = Cont + 1
                    LastW = (Sum * R.Width / Tot)
                    If j = UBound(Values) Then N = j
                Else
                    N = j - 1
                    Exit For
                End If
            Else
                Sum = Sum + Values(j)
                A = Sum / R.Width
                If MinAsp = 0 Then
                    P = Tot / Sum
                Else
                    P = Values(j) / A
                End If
                
                Aspect = Max(P / A, A / P)
                
                If Aspect < MinAsp Or MinAsp = 0 Then
                    TempSum = TempSum + Values(j)
                    MinAsp = Aspect
                    Cont = Cont + 1
                    LastH = (Sum * R.Height / Tot)
                    If j = UBound(Values) Then N = j
                Else
                    N = j - 1
                    Exit For
                End If
            End If
        Next
        
        Sum = 0
        Y = R.Top
        X = R.Left
        
        For j = i To N
            If bDrawVert Then
                Sum = Sum + Values(j)
                P = Values(j) * R.Height / TempSum
                
                With Rects(j)
                    .Left = X
                    .Top = Y
                    .Width = LastW
                    .Height = P
                End With

                Y = Y + P
                If j = N Then
                    R.Width = R.Width - LastW
                    R.Left = R.Left + LastW
                End If
            Else
                Sum = Sum + Values(j)
                P = Values(j) * R.Width / TempSum
                
                With Rects(j)
                    .Left = X
                    .Top = Y
                    .Width = P
                    .Height = LastH
                End With
                
                X = X + P
                If j = N Then
                    R.Height = R.Height - LastH
                    R.Top = R.Top + LastH
                End If
            End If
        Next
        i = N
        Tot = Tot - Sum
    Next
End Sub

Private Function Max(Val1 As Single, Val2 As Single) As Single
    If Val1 > Val2 Then
        Max = Val1
    Else
        Max = Val2
    End If
End Function

Private Function DrawVertically(Width As Single, Height As Single) As Boolean
    DrawVertically = Width > Height
End Function

Private Function IsDarkColor(ByVal Color As Long) As Boolean
    Dim BGRA(0 To 3) As Byte
    OleTranslateColor Color, 0, VarPtr(Color)
    CopyMemory BGRA(0), Color, 4&
  
    IsDarkColor = ((CLng(BGRA(0)) + (CLng(BGRA(1) * 3)) + CLng(BGRA(2))) / 2) < 382

End Function

Public Function ChrW2(ByVal CharCode As Long) As String
  Const POW10 As Long = 2 ^ 10
  If CharCode <= &HFFFF& Then ChrW2 = ChrW$(CharCode) Else _
                              ChrW2 = ChrW$(&HD800& + (CharCode And &HFFFF&) \ POW10) & _
                                      ChrW$(&HDC00& + (CharCode And (POW10 - 1)))
End Function


'Autor: Cobein
Private Function ReadValue(ByVal lProp As Long, Optional Default As Long) As Long
    Dim i       As Long
    For i = 0 To TLS_MINIMUM_AVAILABLE - 1
        If TlsGetValue(i) = lProp Then
            ReadValue = TlsGetValue(i + 1)
            Exit Function
        End If
    Next
    ReadValue = Default
End Function
