VERSION 5.00
Begin VB.UserControl LabelPlus 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   ClipBehavior    =   0  'None
   PropertyPages   =   "LabelPlus.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   318
   Windowless      =   -1  'True
   Begin VB.Timer tmrMOUSEOVER 
      Left            =   1320
      Top             =   1800
   End
End
Attribute VB_Name = "LabelPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------
'Module Name: LabelPlus
'Autor:  Leandro Ascierto
'Web: www.leandroascierto.com
'Published: 14/02/2020
'LastUpdate: 20/12/2021
'Version: 1.5.4
'Based on: FirenzeLabel Project :http://www.vbforums.com/showthread.php?845221-VB6-FIRENZE-LABEL-label-control-with-so-many-functions
           'Martin Vartiak, powered by Cairo Graphics and vbRichClient-Framework.
'Special thanks to: All members of the VB6 Latin group (www.leandroacierto.com/foro), vbforum.com and activevb.net
'-----------------------------------------------
'History
'    1.5.4   change function Image()
'            remove safe gdi+, is not necesary in post vb6 IDE  Windows XP
'------------------------------------------------

Private Declare Function TlsGetValue Lib "kernel32.dll" (ByVal dwTlsIndex As Long) As Long
Private Declare Function TlsSetValue Lib "kernel32.dll" (ByVal dwTlsIndex As Long, ByVal lpTlsValue As Long) As Long
Private Declare Function TlsFree Lib "kernel32.dll" (ByVal dwTlsIndex As Long) As Long
Private Declare Function TlsAlloc Lib "kernel32.dll" () As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function PtInRect Lib "user32.dll" (ByRef lpRect As Rect, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SetRect Lib "user32" (lpRect As Any, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal HDC As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As Any, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal HDC As Long, ByVal nIndex As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal HDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)
Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Any, ByRef Image As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal HDC As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal CallbackData As Long = 0) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imageattr As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef imageattr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imageattr As Long, ByVal ColorAdjust As Long, ByVal EnableFlag As Boolean, ByRef MatrixColor As COLORMATRIX, ByRef MatrixGray As COLORMATRIX, ByVal flags As Long) As Long
Private Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal Image As Long, ByVal rfType As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal ARGB As Long, ByRef Brush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal Brush As Long) As Long
Private Declare Function GdipSetSolidFillColor Lib "gdiplus" (ByVal Brush As Long, ByVal ARGB As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDrawLineI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipFillPolygonI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByRef mPoints As Any, ByVal mCount As Long, ByVal mFillMode As Long) As Long
Private Declare Function GdipDrawPolygonI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByRef mPoints As Any, ByVal mCount As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipAddPathLineI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipCreateTexture Lib "GdiPlus.dll" (ByVal mImage As Long, ByVal mWrapMode As Long, ByRef mTexture As Long) As Long
Private Declare Function GdipClosePathFigures Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipBitmapUnlockBits Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mLockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapLockBits Lib "GdiPlus.dll" (ByVal mBitmap As Long, ByRef mRect As RectL, ByVal mFlags As ImageLockMode, ByVal mPixelFormat As Long, ByRef mLockedBitmapData As BitmapData) As Long
Private Declare Function GdipGetImageHeight Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mHeight As Long) As Long
Private Declare Function GdipGetImageWidth Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mWidth As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "GdiPlus.dll" (ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStride As Long, ByVal mPixelFormat As Long, ByVal mScan0 As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipMeasureString Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RectF, ByVal mStringFormat As Long, ByRef mBoundingBox As RectF, ByRef mCodepointsFitted As Long, ByRef mLinesFilled As Long) As Long
Private Declare Function GdipCreateFont Lib "GdiPlus.dll" (ByVal mFontFamily As Long, ByVal mEmSize As Single, ByVal mStyle As Long, ByVal mUnit As Long, ByRef mFont As Long) As Long
Private Declare Function GdipDeleteFont Lib "GdiPlus.dll" (ByVal mFont As Long) As Long
Private Declare Function GdipDrawString Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RectF, ByVal mStringFormat As Long, ByVal mBrush As Long) As Long
Private Declare Function GdipSetPenMode Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mPenMode As PenAlignment) As Long
Private Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "GdiPlus.dll" (ByRef mRect As RectL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mAngle As Single, ByVal mIsAngleScalable As Long, ByVal mWrapMode As WrapMode, ByRef mLineGradient As Long) As Long
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal Name As Long, ByVal fontCollection As Long, fontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As Long
Private Declare Function GdipAddPathString Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFamily As Long, ByVal mStyle As Long, ByVal mEmSize As Single, ByRef mLayoutRect As RectF, ByVal mFormat As Long) As Long
Private Declare Function GdipGetGenericFontFamilySerif Lib "gdiplus" (ByRef nativeFamily As Long) As Long
Private Declare Function GdipGetGenericFontFamilySansSerif Lib "GdiPlus.dll" (ByRef mNativeFamily As Long) As Long
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As Long
Private Declare Function GdipSetStringFormatFlags Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mFlags As StringFormatFlags) As Long
Private Declare Function GdipSetStringFormatHotkeyPrefix Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mHotkeyPrefix As HotkeyPrefix) As Long
Private Declare Function GdipSetStringFormatTrimming Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mTrimming As StringTrimming) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As StringAlignment) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mAlign As StringAlignment) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As Long) As Long
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal Angle As Single, ByVal Order As Long) As Long
Private Declare Function GdipResetWorldTransform Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipDeleteStringFormat Lib "GdiPlus.dll" (ByVal mFormat As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GdiPlus.dll" (ByVal mHbm As Long, ByVal mhPal As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipCreateEffect Lib "gdiplus" (ByVal dwCid1 As Long, ByVal dwCid2 As Long, ByVal dwCid3 As Long, ByVal dwCid4 As Long, ByRef Effect As Long) As Long
Private Declare Function GdipSetEffectParameters Lib "gdiplus" (ByVal Effect As Long, ByRef params As Any, ByVal Size As Long) As Long
Private Declare Function GdipDeleteEffect Lib "gdiplus" (ByVal Effect As Long) As Long
Private Declare Function GdipDrawImageFX Lib "gdiplus" (ByVal graphics As Long, ByVal Image As Long, ByRef Source As RectF, ByVal xForm As Long, ByVal Effect As Long, ByVal imageAttributes As Long, ByVal srcUnit As Long) As Long
Private Declare Function GdipDrawPie Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipDrawArc Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipSetClipRectI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mCombineMode As Long) As Long
Private Declare Function GdipResetClip Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipResetPath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal HDC As Long, ByRef lpRect As Rect, ByVal hBrush As Long) As Long

Private Type BlurParams
    Radius As Single
    ExpandEdge As Long
End Type

Private Type RectF
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type

Private Type COLORMATRIX
    M(0 To 4, 0 To 4)           As Single
End Type

Private Type PicBmp
    Size As Long
    type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type RectL
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Type POINTL
    X As Long
    Y As Long
End Type

Private Type BitmapData
    Width                       As Long
    Height                      As Long
    stride                      As Long
    PixelFormat                 As Long
    Scan0Ptr                    As Long
    ReservedPtr                 As Long
End Type

Private Type GDIPlusStartupInput
    GdiPlusVersion              As Long
    DebugEventCallback          As Long
    SuppressBackgroundThread    As Long
    SuppressExternalCodecs      As Long
End Type

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public Enum eCallOutPosition
    coLeft
    coTop
    coRight
    coBottom
End Enum

Public Enum eCallOutAlign
    coFirstCorner
    coMidle
    coSecondCorner
    coCustomPosition
End Enum

Public Enum eBorderPosition
    bpInside
    bpCenter
    bpOutside
End Enum

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

Public Enum PictureAlignmentH
    pLeft
    pCenter
    pRight
End Enum

Public Enum PictureAlignmentV
    pTop
    pMiddle
    pBottom
End Enum

Public Enum HotLinePosition
    hlLeft
    hlTop
    hlRight
    hlBottom
End Enum

Private Enum GDIPLUS_FONTSTYLE
    FontStyleRegular = 0
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum
  
Public Enum StringAlignment
    StringAlignmentNear = &H0
    StringAlignmentCenter = &H1
    StringAlignmentFar = &H2
End Enum
  
Public Enum StringTrimming
    StringTrimmingNone = &H0
    StringTrimmingCharacter = &H1
    StringTrimmingWord = &H2
    StringTrimmingEllipsisCharacter = &H3
    StringTrimmingEllipsisWord = &H4
    StringTrimmingEllipsisPath = &H5
End Enum

Public Enum StringFormatFlags
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

Private Enum HotkeyPrefix
    HotkeyPrefixNone = &H0
    HotkeyPrefixShow = &H1
    HotkeyPrefixHide = &H2
End Enum

Private Enum WrapMode
    WrapModeTile = &H0
    WrapModeTileFlipX = &H1
    WrapModeTileFlipy = &H2
    WrapModeTileFlipXY = &H3
    WrapModeClamp = &H4
End Enum

Private Enum ImageLockMode
    ImageLockModeRead = &H1
    ImageLockModeWrite = &H2
    ImageLockModeUserInputBuf = &H4
End Enum
 
Private Enum PenAlignment
    PenAlignmentCenter = &H0
    PenAlignmentInset = &H1
End Enum

Private Const TLS_MINIMUM_AVAILABLE     As Long = 64
Private Const IDC_HAND                  As Long = 32649
Private Const GWL_WNDPROC               As Long = -4
Private Const GW_OWNER                  As Long = 4
Private Const WS_CHILD                  As Long = &H40000000
Private Const WM_MOUSEMOVE              As Long = &H200
Private Const UnitPixel                 As Long = &H2&
Private Const LOGPIXELSX                As Long = 88
Private Const LOGPIXELSY                As Long = 90
Private Const PixelFormat32bppPARGB     As Long = &HE200B
Private Const PixelFormat32bppARGB      As Long = &H26200A
Private Const SmoothingModeAntiAlias    As Long = 4
Private Const CombineModeExclude        As Long = &H4

Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseEnter()
Public Event MouseLeave()
Public Event MouseOver()
Public Event MouseOut()
Public Event PrePaint(HDC As Long, X As Long, Y As Long)
Public Event PostPaint(ByVal HDC As Long)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Public Event PictureDownloadProgress(BytesMax As Long, BytesLeidos As Long)
Public Event PictureDownloadComplete()
Public Event PictureDownloadError()

Dim m_CallOut As Boolean
Dim m_CallOutPosicion As eCallOutPosition
Dim m_CallOutAlign As eCallOutAlign
Dim m_coLen As Long
Dim m_coWidth As Long
Dim m_coCustomPos As Long
Dim m_coRightTriangle As Boolean
Dim m_BackColor As OLE_COLOR
Dim m_BackColorOpacity As Integer
Dim m_BackAcrylicBlur As Boolean
Dim m_BackShadow As Boolean
Dim m_Border As Boolean
Dim m_BorderColor As OLE_COLOR
Dim m_BorderColorOpacity As Integer
Dim m_BorderPosition As eBorderPosition
Dim m_BorderCornerLeftTop As Integer
Dim m_BorderCornerRightTop As Integer
Dim m_BorderCornerBottomLeft As Integer
Dim m_BorderCornerBottomRight As Integer
Dim m_BorderWidth As Integer
Dim hImgShadow As Long
Dim m_ShadowSize As Integer
Dim m_ShadowColor As OLE_COLOR
Dim m_ShadowOffsetX As Integer
Dim m_ShadowOffsetY As Integer
Dim m_ShadowColorOpacity As Integer
Dim hImgCaptionShadow As Long
Dim m_Caption() As Byte
Dim m_CaptionAlignmentH As CaptionAlignmentH
Dim m_CaptionAlignmentV As CaptionAlignmentV
Dim m_CaptionPaddingX As Integer
Dim m_CaptionPaddingY As Integer
Dim m_CaptionTriming As StringTrimming
Dim m_CaptionBorderWidth As Integer
Dim m_CaptionBorderColor As OLE_COLOR
Dim m_CaptionShadow As Boolean
Dim m_CaptionAngle As Integer
Dim m_CaptionShowPrefix As Boolean
Dim m_AutoSize As Boolean
Dim m_MousePointerHands As Boolean
Dim m_Font As StdFont
Attribute m_Font.VB_VarHelpID = -1
Dim m_ForeColor As OLE_COLOR
Dim m_ForeColorOpacity As Integer
Dim m_Gradient As Boolean
Dim m_GradientAngle As Integer
Dim m_GradientColor1 As OLE_COLOR
Dim m_GradientColor1Opacity As Integer
Dim m_GradientColor2 As OLE_COLOR
Dim m_GradientColor2Opacity As Integer
Dim m_PictureAngle As Integer
Dim m_PictureAlignmentH As PictureAlignmentH
Dim m_PictureAlignmentV As PictureAlignmentV
Dim m_PicturePaddingX As Integer
Dim m_PicturePaddingY As Integer
Dim m_PictureRealWidth As Long
Dim m_PictureRealHeight As Long
Dim m_PictureSetWidth As Long
Dim m_PictureSetHeight As Long
Dim m_PictureArr()  As Byte
Dim m_PicturePresent As Boolean
Dim m_PictureGraysScale As Boolean
Dim m_PictureContrast As Integer
Dim m_PictureBrightness As Integer
Dim m_PictureOpacity As Integer
Dim m_PictureColor As OLE_COLOR
Dim m_PictureColorize As Boolean
Dim m_PictureShadow As Boolean
Dim m_MouseToParent As Boolean
Dim m_HotLine As Boolean
Dim m_HotLineColor As OLE_COLOR
Dim m_HotLineColorOpacity As Integer
Dim m_HotLineWidth As Long
Dim m_HotLinePosition As HotLinePosition
Dim m_IconFont As StdFont
Dim m_IconCharCode As Long
Dim m_IconForeColor As Long
Dim m_IconPaddingX As Integer
Dim m_IconPaddingY As Integer
Dim m_IconAlignmentH As CaptionAlignmentH
Dim m_IconAlignmentV As CaptionAlignmentV
Dim m_IconOpacity As Integer
Dim m_WordWrap As Boolean
Dim m_Redraw As Boolean
Dim hCur As Long
Dim c_lhWnd As Long
Dim nScale As Single
Dim hDCMemory As Long
Dim hBmp As Long
Dim OldhBmp As Long
Dim bRecreateShadowCaption As Boolean
Dim m_DrawProgress As Boolean
Dim c_AsyncProp As AsyncProperty
Dim m_PictureBrush As Long
Dim hFontCollection As Long
Dim bIntercept As Boolean
Dim m_Enter As Boolean
Dim m_Over As Boolean
Dim m_PT As POINTAPI
Dim m_Left As Long
Dim m_Top As Long
Dim GdipToken As Long

Public Sub Draw(ByVal HDC As Long, ByVal hGraphics As Long, ByVal PosX As Long, PosY As Long)
    Dim hPath As Long
    Dim hBrush As Long
    Dim hPen As Long
    Dim X As Long, Y As Long
    Dim XX As Long, YY As Long
    Dim lWidth As Long, lHeight As Long
    Dim WW As Long, HH As Long
    Dim ShadowSize As Integer
    Dim ShadowOffsetX As Integer
    Dim ShadowOffsetY As Integer
    Dim BorderWidth As Integer

    ShadowSize = m_ShadowSize * nScale
    ShadowOffsetX = m_ShadowOffsetX * nScale
    ShadowOffsetY = m_ShadowOffsetY * nScale
    BorderWidth = m_BorderWidth * nScale
    
    If m_BackAcrylicBlur Then
        BitBlt hDCMemory, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.HDC, 0, 0, vbSrcCopy
    End If
     
    If hGraphics = 0 Then GdipCreateFromHDC HDC, hGraphics
    
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    GdipTranslateWorldTransform hGraphics, PosX, PosY, &H1
    
    If m_BorderPosition = bpInside Then
        lWidth = UserControl.ScaleWidth
        lHeight = UserControl.ScaleHeight
    ElseIf m_BorderPosition = bpCenter Then
        X = (BorderWidth \ 2)
        Y = (BorderWidth \ 2)
        lWidth = UserControl.ScaleWidth - BorderWidth
        lHeight = UserControl.ScaleHeight - BorderWidth
    Else
        X = BorderWidth
        Y = BorderWidth
        lWidth = UserControl.ScaleWidth - (BorderWidth * 2)
        lHeight = UserControl.ScaleHeight - (BorderWidth * 2)
    End If
    
    If hImgShadow Then
        XX = IIf(ShadowOffsetX > 0, ShadowOffsetX, 0) '+ PosX
        YY = IIf(ShadowOffsetY > 0, ShadowOffsetY, 0) '+ PosY
        GdipDrawImageRectI hGraphics, hImgShadow, XX, YY, UserControl.ScaleWidth - Abs(ShadowOffsetX), UserControl.ScaleHeight - Abs(ShadowOffsetY)
    End If
    
    If m_BackShadow = True And m_ShadowSize > 0 Then
        X = X + ShadowSize + IIf(ShadowOffsetX < 0, Abs(ShadowOffsetX), 0) '+ PosX
        Y = Y + ShadowSize + IIf(ShadowOffsetY < 0, Abs(ShadowOffsetY), 0) '+ PosY
        lWidth = lWidth - (ShadowSize * 2) - Abs(ShadowOffsetX)
        lHeight = lHeight - (ShadowSize * 2) - Abs(ShadowOffsetY)
    End If
  
    XX = X:         YY = Y
    WW = lWidth:    HH = lHeight

    hPath = RoundRectangle(XX, YY, WW, HH)
    
    If m_BackAcrylicBlur Then
        DrawAcrylicBlur hGraphics, hPath
    End If

    If m_Gradient Then
        Dim RectL As RectL
        SetRect RectL, X, Y, lWidth, lHeight
        GdipCreateLineBrushFromRectWithAngleI RectL, ConvertColor(m_GradientColor1, m_GradientColor1Opacity), _
                                                    ConvertColor(m_GradientColor2, m_GradientColor2Opacity), _
                                                    m_GradientAngle + 90, 0, WrapModeTileFlipXY, hBrush
    Else
        GdipCreateSolidFill ConvertColor(m_BackColor, m_BackColorOpacity), hBrush
    End If
        
    GdipFillPath hGraphics, hBrush, hPath
    GdipDeleteBrush hBrush

    If m_PicturePresent Then
        If m_PictureBrush = 0 Then m_PictureBrush = CreateBrushTexture(hGraphics, hPath, XX, YY, WW, HH)
        GdipFillPath hGraphics, m_PictureBrush, hPath
    End If

    If Not c_AsyncProp Is Nothing Then
        If c_AsyncProp.BytesMax = c_AsyncProp.BytesRead Then
            Set c_AsyncProp = Nothing
        Else
            DrawProgress hGraphics, XX, YY, WW, HH, c_AsyncProp.BytesRead, c_AsyncProp.BytesMax
        End If
    End If
    
    If m_CaptionShadow Then
        CreateCaptionShadow XX, YY, WW, HH
        
        If hImgCaptionShadow <> 0 Then
            X = XX - ShadowSize + IIf(ShadowOffsetX > 0, ShadowOffsetX, ShadowOffsetX * 2) + PosX
            Y = YY - ShadowSize + IIf(ShadowOffsetY > 0, ShadowOffsetY, ShadowOffsetY * 2) + PosY
            GdipDrawImageRectI hGraphics, hImgCaptionShadow, X, Y, WW + ShadowSize * 2, HH + ShadowSize * 2
        End If
    End If
    
    If m_HotLine Then
        DrawHotLine hGraphics, hPath ', PosX, PosY
    End If
    
    GDIP_AddPathString hGraphics, XX, YY, WW, HH

    If m_Border And BorderWidth > 0 Then
        GdipCreatePen1 ConvertColor(m_BorderColor, m_BorderColorOpacity), BorderWidth, UnitPixel, hPen
        
        If m_BorderPosition = bpInside Then
            GdipSetPenMode hPen, PenAlignmentInset
        ElseIf m_BorderPosition = bpOutside Then
    
            GdipDeletePath hPath
            X = (BorderWidth / 2) + PosX
            Y = (BorderWidth / 2) + PosY
            lWidth = UserControl.ScaleWidth - BorderWidth
            lHeight = UserControl.ScaleHeight - BorderWidth
            hPath = RoundRectangle(X, Y, lWidth, lHeight, True)
        End If
        
        GdipDrawPath hGraphics, hPen, hPath
        GdipDeletePen hPen
    End If
    
    GdipDeletePath hPath
    If HDC <> 0 Then GdipDeleteGraphics hGraphics
End Sub

Private Function GDIP_AddPathString(ByVal hGraphics As Long, X As Long, Y As Long, Width As Long, Height As Long, Optional ForShadow As Boolean, Optional GetMeasureString As Boolean) As Boolean
    Dim hPath As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim hFontFamily As Long
    Dim hFormat As Long
    Dim layoutRect As RectF
    Dim lFontSize As Long
    Dim lFontStyle As GDIPLUS_FONTSTYLE
    Dim hFont As Long
    Dim HDC As Long

    If GdipCreatePath(&H0, hPath) = 0 Then
    
        If GdipCreateStringFormat(0, 0, hFormat) = 0 Then
            If Not m_WordWrap Then GdipSetStringFormatFlags hFormat, StringFormatFlagsNoWrap
            If m_CaptionShowPrefix Then GdipSetStringFormatHotkeyPrefix hFormat, HotkeyPrefixShow
            GdipSetStringFormatTrimming hFormat, m_CaptionTriming
            GdipSetStringFormatAlign hFormat, m_CaptionAlignmentH
            GdipSetStringFormatLineAlign hFormat, m_CaptionAlignmentV
        End If

        GetFontStyleAndSize m_Font, lFontStyle, lFontSize
        
        If GdipCreateFontFamilyFromName(StrPtr(m_Font.Name), 0, hFontFamily) Then
            If hFontCollection Then
                If GdipCreateFontFamilyFromName(StrPtr(m_Font.Name), hFontCollection, hFontFamily) Then
                    If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
                End If
            Else
                If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
            End If
        End If
        
        If GetMeasureString Then
            Dim BB As RectF, CF As Long, LF As Long
            
            With layoutRect
                .Left = X: .Top = Y
                .Width = Width: .Height = Height
            End With
            
            Call GdipCreateFont(hFontFamily, lFontSize, lFontStyle, UnitPixel, hFont)
            GdipMeasureString hGraphics, StrPtr(m_Caption), -1, hFont, layoutRect, hFormat, BB, CF, LF
            GdipDeleteFont hFont
           
            X = BB.Left
            Y = BB.Top
            Width = BB.Width
            Height = BB.Height
            GdipDeleteFontFamily hFontFamily
        Else
            With layoutRect
                .Left = X + m_CaptionPaddingX * nScale: .Width = Width - (m_CaptionPaddingX * nScale) * 2
                .Top = Y + m_CaptionPaddingY * nScale: .Height = Height - (m_CaptionPaddingY * nScale) * 2
            End With
            
            If m_CaptionAngle <> 0 Then
                If ForShadow Then
                    layoutRect.Left = layoutRect.Left - (Width / 2)
                    layoutRect.Top = layoutRect.Top - (Height / 2)
                    Call GdipTranslateWorldTransform(hGraphics, (Width / 2), (Height / 2), 0)
                Else
                    layoutRect.Left = layoutRect.Left - (UserControl.ScaleWidth / 2)
                    layoutRect.Top = layoutRect.Top - (UserControl.ScaleHeight / 2)
                    Call GdipTranslateWorldTransform(hGraphics, (UserControl.ScaleWidth / 2), (UserControl.ScaleHeight / 2), 0)
                End If
                Call GdipRotateWorldTransform(hGraphics, m_CaptionAngle, 0)
            End If
            
            GdipAddPathString hPath, StrPtr(m_Caption), -1, hFontFamily, lFontStyle, lFontSize, layoutRect, hFormat
            GdipDeleteStringFormat hFormat
        
            GdipCreateSolidFill ConvertColor(m_ForeColor, IIf(ForShadow, m_ShadowColorOpacity, m_ForeColorOpacity)), hBrush
            GdipFillPath hGraphics, hBrush, hPath
            GdipDeleteBrush hBrush
            
            If m_CaptionBorderWidth > 0 Then
               GdipCreatePen1 ConvertColor(m_CaptionBorderColor, IIf(ForShadow, m_ShadowColorOpacity, 100)), m_CaptionBorderWidth, UnitPixel, hPen
               GdipDrawPath hGraphics, hPen, hPath
               GdipDeletePen hPen
            End If
            
            If m_CaptionAngle <> 0 Then GdipResetWorldTransform hGraphics
            
            GdipDeleteFontFamily hFontFamily
            
            If m_IconCharCode Then
                
                If GdipCreateFontFamilyFromName(StrPtr(m_IconFont.Name), 0, hFontFamily) Then
                    If GdipCreateFontFamilyFromName(StrPtr(m_IconFont.Name), hFontCollection, hFontFamily) Then
                        GdipDeletePath hPath
                        Exit Function
                    End If
                End If
                
                With layoutRect
                    .Left = X + m_IconPaddingX * nScale: .Width = Width - (m_IconPaddingX * nScale) * 2
                    .Top = Y + m_IconPaddingY * nScale: .Height = Height - (m_IconPaddingY * nScale) * 2
                End With
                
                If m_CaptionAngle <> 0 Then
                    If ForShadow Then
                        layoutRect.Left = layoutRect.Left - (Width / 2)
                        layoutRect.Top = layoutRect.Top - (Height / 2)
                        Call GdipTranslateWorldTransform(hGraphics, (Width / 2), (Height / 2), 0)
                    Else
                        layoutRect.Left = layoutRect.Left - (UserControl.ScaleWidth / 2)
                        layoutRect.Top = layoutRect.Top - (UserControl.ScaleHeight / 2)
                        Call GdipTranslateWorldTransform(hGraphics, (UserControl.ScaleWidth / 2), (UserControl.ScaleHeight / 2), 0)
                    End If
                    Call GdipRotateWorldTransform(hGraphics, m_CaptionAngle, 0)
                End If
                GetFontStyleAndSize m_IconFont, lFontStyle, lFontSize
                
                If GdipCreateStringFormat(0, 0, hFormat) = 0 Then
                    GdipSetStringFormatAlign hFormat, m_IconAlignmentH
                    GdipSetStringFormatLineAlign hFormat, m_IconAlignmentV
                End If
                                
                GdipResetPath hPath
                GdipAddPathString hPath, StrPtr(ChrW2(m_IconCharCode)), -1, hFontFamily, lFontStyle, lFontSize, layoutRect, hFormat
                GdipDeleteStringFormat hFormat
            
                GdipCreateSolidFill ConvertColor(m_IconForeColor, IIf(ForShadow, m_ShadowColorOpacity, m_IconOpacity)), hBrush
                GdipFillPath hGraphics, hBrush, hPath
                GdipDeleteBrush hBrush
                
                If m_CaptionBorderWidth > 0 Then
                   GdipCreatePen1 ConvertColor(m_CaptionBorderColor, IIf(ForShadow, m_ShadowColorOpacity, 100)), m_CaptionBorderWidth, UnitPixel, hPen
                   GdipDrawPath hGraphics, hPen, hPath
                   GdipDeletePen hPen
                End If
                
                GdipDeleteFontFamily hFontFamily
                
                If m_CaptionAngle <> 0 Then GdipResetWorldTransform hGraphics
            End If
        End If
        
        GdipDeletePath hPath
    End If

End Function

Private Function GetFontStyleAndSize(oFont As StdFont, lFontStyle As Long, lFontSize As Long)
        Dim HDC As Long
        lFontStyle = 0
        If oFont.Bold Then lFontStyle = lFontStyle Or FontStyleBold
        If oFont.Italic Then lFontStyle = lFontStyle Or FontStyleItalic
        If oFont.Underline Then lFontStyle = lFontStyle Or FontStyleUnderline
        If oFont.Strikethrough Then lFontStyle = lFontStyle Or FontStyleStrikeout
        
        HDC = GetDC(0&)
        lFontSize = MulDiv(oFont.Size, GetDeviceCaps(HDC, LOGPIXELSY), 72)
        ReleaseDC 0&, HDC
End Function

Function DrawText(ByVal HDC As Long, ByVal text As String, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal oFont As StdFont, ByVal ForeColor As OLE_COLOR, Optional ByVal ColorOpacity As Integer = 100, Optional HAlign As CaptionAlignmentH, Optional VAlign As CaptionAlignmentV, Optional bWordWrap As Boolean) As Long
    Dim hBrush As Long
    Dim hFontFamily As Long
    Dim hFormat As Long
    Dim layoutRect As RectF
    Dim lFontSize As Long
    Dim lFontStyle As GDIPLUS_FONTSTYLE
    Dim hFont As Long
    Dim hGraphics As Long
    
    SafeRange ColorOpacity, 0, 100
    
    GdipCreateFromHDC HDC, hGraphics
  
    If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), 0, hFontFamily) Then
        If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
        'If GdipGetGenericFontFamilySerif(hFontFamily) Then Exit Function
    End If
    
    If GdipCreateStringFormat(0, 0, hFormat) = 0 Then
        If Not bWordWrap Then GdipSetStringFormatFlags hFormat, StringFormatFlagsNoWrap
        'GdipSetStringFormatFlags hFormat, HotkeyPrefixShow
        GdipSetStringFormatAlign hFormat, HAlign
        GdipSetStringFormatLineAlign hFormat, VAlign
    End If
        
    If oFont.Bold Then lFontStyle = lFontStyle Or FontStyleBold
    If oFont.Italic Then lFontStyle = lFontStyle Or FontStyleItalic
    If oFont.Underline Then lFontStyle = lFontStyle Or FontStyleUnderline
    If oFont.Strikethrough Then lFontStyle = lFontStyle Or FontStyleStrikeout
        

    lFontSize = MulDiv(oFont.Size, GetDeviceCaps(HDC, LOGPIXELSY), 72)

    layoutRect.Left = X * nScale: layoutRect.Top = Y * nScale
    layoutRect.Width = Width * nScale: layoutRect.Height = Height * nScale

    GdipCreateSolidFill ConvertColor(ForeColor, ColorOpacity), hBrush
            
    'GdipSetTextRenderingHint hGraphics, TextRenderingHintClearTypeGridFit
    'GdipSetTextContrast hGraphics, 12
    Call GdipCreateFont(hFontFamily, lFontSize, lFontStyle, UnitPixel, hFont)
    GdipDrawString hGraphics, StrPtr(text), -1, hFont, layoutRect, hFormat, hBrush
    
    Dim BB As RectF, CF As Long, LF As Long

    Call GdipCreateFont(hFontFamily, lFontSize, lFontStyle, UnitPixel, hFont)
    GdipMeasureString hGraphics, StrPtr(text), -1, hFont, layoutRect, hFormat, BB, CF, LF

    
    If bWordWrap Then
        DrawText = BB.Height / nScale
    Else
        DrawText = BB.Width / nScale
    End If
    
    GdipDeleteFont hFont
    GdipDeleteBrush hBrush
    GdipDeleteStringFormat hFormat
    GdipDeleteFontFamily hFontFamily
    GdipDeleteGraphics hGraphics

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

Private Sub CreateShadow()
    Dim hImage As Long
    Dim hGraphics As Long
    Dim hPath As Long
    Dim hBrush As Long, hPen As Long
    Dim lWidth As Long, lHeight As Long
    Dim ShadowSize As Integer
    
    If hImgShadow Then GdipDisposeImage hImgShadow: hImgShadow = 0
    If m_BackShadow = False Then Exit Sub
    
    bRecreateShadowCaption = True
        
    If m_ShadowSize = 0 Then Exit Sub
    If m_BackColorOpacity = 0 And m_Border = False Then Exit Sub
    If m_ShadowColorOpacity = 0 Then Exit Sub
   
    ShadowSize = m_ShadowSize * nScale
    lWidth = UserControl.ScaleWidth - (ShadowSize * 2)
    lHeight = UserControl.ScaleHeight - (ShadowSize * 2)
    
    GdipCreateBitmapFromScan0 lWidth, lHeight, 0&, PixelFormat32bppPARGB, ByVal 0&, hImage
    GdipGetImageGraphicsContext hImage, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    
    hPath = RoundRectangle(0, 0, lWidth - 0, lHeight - 0, True, True)
    
    If m_BackColorOpacity > 0 Then
        GdipCreateSolidFill ConvertColor(m_ShadowColor, m_ShadowColorOpacity), hBrush
        GdipFillPath hGraphics, hBrush, hPath
        GdipDeleteBrush hBrush
    Else
        GdipCreatePen1 ConvertColor(m_ShadowColor, m_ShadowColorOpacity), (m_BorderWidth * nScale) * 2, UnitPixel, hPen
        GdipDrawPath hGraphics, hPen, hPath
        GdipDeletePen hPen
    End If
        
    hImgShadow = CreateBlurShadowImage(hImage, m_ShadowColor, ShadowSize, 0, 0, lWidth, lHeight)
    
    GdipDeletePath hPath
    GdipDeleteGraphics hGraphics
    GdipDisposeImage hImage
    
End Sub

Private Sub CreateCaptionShadow(ByVal X As Long, ByVal Y As Long, ByVal lWidth As Long, ByVal lHeight As Long)
    Dim hGraphics As Long
    Dim hPath As Long
    Dim hBrush As Long, hPen As Long
    Dim hImage As Long
    Dim RecL As RectL
    
    If bRecreateShadowCaption = False Then Exit Sub
    If hImgCaptionShadow Then GdipDisposeImage hImgCaptionShadow: hImgCaptionShadow = 0
    If m_ShadowSize = 0 Then Exit Sub
    If UBound(m_Caption) <= 0 Then Exit Sub
    
    GdipCreateBitmapFromScan0 lWidth, lHeight, 0&, PixelFormat32bppPARGB, ByVal 0&, hImage
    GdipGetImageGraphicsContext hImage, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias

    GDIP_AddPathString hGraphics, 0, 0, lWidth, lHeight, True

    hImgCaptionShadow = CreateBlurShadowImage(hImage, m_ShadowColor, m_ShadowSize * nScale, 0, 0, lWidth, lHeight)
    bRecreateShadowCaption = False
    
    GdipDeletePath hPath
    GdipDeleteGraphics hGraphics
    GdipDisposeImage hImage
    
End Sub

Private Function DrawHotLine(hGraphics As Long, hPath As Long) ', ByVal PosX As Long, ByVal PosY As Long)
    Dim hBrush As Long
    Dim X As Long, Y As Long
    Dim WW As Long, HH As Long
    Dim BW As Long
    Dim LW As Long
    Dim CL As Long
    Dim SS As Long

    
    Select Case m_BorderPosition
        Case bpOutside: BW = BorderWidth * nScale
        Case bpCenter: BW = BorderWidth / 2 * nScale
        Case bpInside
    End Select
    
    If m_BorderPosition = bpOutside Then
        If m_HotLinePosition = hlLeft Or m_HotLinePosition = hlTop Then
            X = BW
            Y = BW
            WW = UserControl.ScaleWidth - BW
            HH = UserControl.ScaleHeight - BW
        Else
            WW = UserControl.ScaleWidth - BW
            HH = UserControl.ScaleHeight - BW
        End If
    Else
        X = BW
        Y = BW
        WW = UserControl.ScaleWidth - BW * 2
        HH = UserControl.ScaleHeight - BW * 2
    End If
    
    If m_BackShadow Then
        SS = ShadowSize * nScale
        If m_HotLinePosition = hlRight Or m_HotLinePosition = hlBottom Then
            X = X - SS
            Y = Y - SS
        Else
            X = X + SS
            Y = Y + SS
        End If
        If ShadowOffsetX < 0 Then X = X + Abs(ShadowOffsetX * nScale)
        If ShadowOffsetY < 0 Then Y = Y + Abs(ShadowOffsetY * nScale)
    End If
    
    LW = m_HotLineWidth * nScale
    Select Case m_HotLinePosition
        Case hlLeft: X = X + LW
        Case hlTop: Y = Y + LW
        Case hlRight: WW = WW - LW
        Case hlBottom: HH = HH - LW
    End Select
    
    If m_CallOut Then
        CL = m_coLen * nScale
        Select Case m_CallOutPosicion
            Case coLeft: If m_HotLinePosition = hlLeft Then X = X + CL
            Case coTop: If m_HotLinePosition = hlTop Then Y = Y + CL
            Case coRight: If m_HotLinePosition = hlRight Then WW = WW - CL
            Case coBottom: If m_HotLinePosition = hlBottom Then HH = HH - CL
        End Select
    End If
        
    GdipSetClipRectI hGraphics, X, Y, WW, HH, CombineModeExclude
    GdipCreateSolidFill ConvertColor(m_HotLineColor, m_HotLineColorOpacity), hBrush
    GdipFillPath hGraphics, hBrush, hPath
    GdipDeleteBrush hBrush
    GdipResetClip hGraphics
        
End Function

Private Sub DrawProgress(hGraphics As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Value As Long, ByVal Max As Long)
    Dim hPen As Long
    Dim ReqWidth As Long, ReqHeight As Long
    Dim HScale As Double, VScale As Double
    Dim MyScale As Double
    Dim imgWidth As Long
    Dim imgHeight As Long
    Dim nSize As Long
    Static Angle As Long

    If m_PictureSetWidth = 0 Then ReqWidth = UserControl.ScaleWidth \ 2 Else ReqWidth = m_PictureSetWidth \ 2
    If m_PictureSetHeight = 0 Then ReqHeight = UserControl.ScaleHeight \ 2 Else ReqHeight = m_PictureSetHeight \ 2

    MyScale = IIf(ReqHeight >= ReqWidth, ReqWidth, ReqHeight)

    ReqWidth = MyScale * nScale
    ReqHeight = MyScale * nScale

        '----------------
    If m_PictureAlignmentH = pLeft Then X = X + (m_PicturePaddingX * nScale)
    If m_PictureAlignmentH = pCenter Then X = X + (Width \ 2) - (ReqWidth \ 2) + (m_PicturePaddingX * nScale)
    If m_PictureAlignmentH = pRight Then X = X + Width - ReqWidth - (m_PicturePaddingX * nScale)
    If m_PictureAlignmentV = pTop Then Y = Y + (m_PicturePaddingY * nScale)
    If m_PictureAlignmentV = pMiddle Then Y = Y + (Height \ 2) - (ReqHeight \ 2) + (m_PicturePaddingY * nScale)
    If m_PictureAlignmentV = pBottom Then Y = Y + Height - ReqHeight - (m_PicturePaddingY * nScale)
    
    GdipCreatePen1 ConvertColor(vbBlack, 50), 3 * nScale, UnitPixel, hPen
    GdipDrawArc hGraphics, hPen, X, Y, ReqWidth, ReqHeight, 0, 360
    GdipDeletePen hPen
    GdipCreatePen1 ConvertColor(&HFFCC00, 50), 3 * nScale, UnitPixel, hPen
    '
    If Max = 0 Then
        Angle = Angle + 36
        GdipDrawArc hGraphics, hPen, X, Y, ReqWidth, ReqHeight, -90 + Angle, 60
    Else
        GdipDrawArc hGraphics, hPen, X, Y, ReqWidth, ReqHeight, -90, 360 * Value / Max
    End If
    GdipDeletePen hPen

End Sub


Private Sub DrawAcrylicBlur(hGraphics As Long, hPath As Long)
    Dim hBrush As Long
    Dim hImage As Long
    Dim hGraphics2 As Long, hImage2 As Long
    Dim lEffect As Long
    Dim bp As BlurParams
    Dim rcSource As RectF

    If GdipCreateBitmapFromHBITMAP(hBmp, 0, hImage) = 0 Then
        Call GdipCreateBitmapFromScan0(UserControl.ScaleWidth, UserControl.ScaleHeight, 0&, PixelFormat32bppARGB, ByVal 0&, hImage2)
        Call GdipGetImageGraphicsContext(hImage2, hGraphics2)

        bp.ExpandEdge = 0
        bp.Radius = 25
              
        Call GdipCreateEffect(&H633C80A4, &H482B1843, &H28BEF29E, &HD4FDC534, lEffect)
        Call GdipSetEffectParameters(lEffect, bp, Len(bp))
              
        rcSource.Width = UserControl.ScaleWidth
        rcSource.Height = UserControl.ScaleHeight
    
        Call GdipDrawImageFX(hGraphics2, hImage, rcSource, 0, lEffect, 0, UnitPixel)
        
        GdipCreateTexture hImage2, &H0, hBrush
        GdipFillPath hGraphics, hBrush, hPath
        
        'Cleanup
        Call GdipDeleteBrush(hBrush)
        Call GdipDeleteEffect(lEffect)
        Call GdipDeleteGraphics(hGraphics2)
        Call GdipDisposeImage(hImage2)
        Call GdipDisposeImage(hImage)
    End If
End Sub

Private Function CreateBrushTexture(hGraphics As Long, hPath As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
    Dim hBrush As Long
    Dim hImage As Long
    Dim hGraphics2 As Long, hImage2 As Long
    Dim tMatrixColor    As COLORMATRIX, tMatrixGray    As COLORMATRIX
    Dim hAttributes As Long
    Dim ReqWidth As Long, ReqHeight As Long
    Dim HScale As Double, VScale As Double
    Dim MyScale As Double
    Dim imgWidth As Long
    Dim imgHeight As Long

    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0

    If LoadImageFromStream(m_PictureArr, hImage) Then
 
        Call GdipCreateImageAttributes(hAttributes)
        
        With tMatrixColor
            If m_PictureColorize Then
                Dim R As Byte, G As Byte, B As Byte

                B = ((m_PictureColor \ &H10000) And &HFF)
                G = ((m_PictureColor \ &H100) And &HFF)
                R = (m_PictureColor And &HFF)
                
                .M(0, 0) = R / 255
                .M(1, 0) = G / 255
                .M(2, 0) = B / 255
                .M(0, 4) = R / 255
                .M(1, 4) = G / 255
                .M(2, 4) = B / 255
            Else
                .M(0, 0) = 1
                .M(1, 1) = 1
                .M(2, 2) = 1

            End If
            .M(3, 3) = m_PictureOpacity / 100
            .M(4, 4) = 1
 
            If Not m_PictureContrast = 0 Then
                .M(0, 0) = 1 + m_PictureContrast
                .M(1, 1) = .M(0, 0)
                .M(2, 2) = .M(0, 0)
                .M(0, 4) = 0.5 * -m_PictureContrast
                .M(1, 4) = .M(0, 4)
                .M(2, 4) = .M(0, 4)
            End If
            
            If m_PictureBrightness <> 0 Then
                .M(0, 4) = .M(0, 4) + m_PictureBrightness / 100
                .M(1, 4) = .M(1, 4) + m_PictureBrightness / 100
                .M(2, 4) = .M(2, 4) + m_PictureBrightness / 100
            End If
            
            If m_PictureGraysScale Then
                .M(0, 0) = 0.299
                .M(1, 0) = 0.299
                .M(2, 0) = 0.299
                .M(0, 1) = 0.587
                .M(1, 1) = 0.587
                .M(2, 1) = 0.587
                .M(0, 2) = 0.114
                .M(1, 2) = 0.114
                .M(2, 2) = 0.114
            End If
        End With

        Call GdipCreateBitmapFromScan0(UserControl.ScaleWidth, UserControl.ScaleHeight, 0&, PixelFormat32bppARGB, ByVal 0&, hImage2)
        Call GdipGetImageGraphicsContext(hImage2, hGraphics2)
        GdipSetSmoothingMode hGraphics2, SmoothingModeAntiAlias
        
        If m_PictureSetWidth = 0 Then ReqWidth = m_PictureRealWidth Else ReqWidth = m_PictureSetWidth
        If m_PictureSetHeight = 0 Then ReqHeight = m_PictureRealHeight Else ReqHeight = m_PictureSetHeight

        HScale = ReqWidth / m_PictureRealWidth
        VScale = ReqHeight / m_PictureRealHeight
        
        MyScale = IIf(VScale >= HScale, HScale, VScale)

        ReqWidth = m_PictureRealWidth * MyScale * nScale
        ReqHeight = m_PictureRealHeight * MyScale * nScale

        If m_PictureAlignmentH = pLeft Then X = X + (m_PicturePaddingX * nScale)
        If m_PictureAlignmentH = pCenter Then X = X + (Width / 2) - (ReqWidth / 2) + (m_PicturePaddingX * nScale)
        If m_PictureAlignmentH = pRight Then X = X + Width - ReqWidth - (m_PicturePaddingX * nScale)
        If m_PictureAlignmentV = pTop Then Y = Y + (m_PicturePaddingY * nScale)
        If m_PictureAlignmentV = pMiddle Then Y = Y + (Height / 2) - (ReqHeight / 2) + (m_PicturePaddingY * nScale)
        If m_PictureAlignmentV = pBottom Then Y = Y + Height - ReqHeight - (m_PicturePaddingY * nScale)

        If m_PictureShadow = True And m_ShadowSize > 0 And m_ShadowColorOpacity > 0 Then
            Dim hPictureShadow As Long
            Dim ShadowSize As Integer
            Dim W As Long, H As Long
            
            ShadowSize = m_ShadowSize * nScale
            hPictureShadow = CreateBlurShadowImage(hImage, m_ShadowColor, ShadowSize, 0, 0, m_PictureRealWidth, m_PictureRealHeight)
            tMatrixColor.M(3, 3) = m_ShadowColorOpacity / 100
            GdipSetImageAttributesColorMatrix hAttributes, &H0, True, tMatrixColor, tMatrixGray, &H0
            If m_PictureAngle <> 0 Then
                W = ReqWidth + ShadowSize * 2
                H = ReqHeight + ShadowSize * 2
                Call GdipRotateWorldTransform(hGraphics2, m_PictureAngle + 180, 0)
                Call GdipTranslateWorldTransform(hGraphics2, X + (W \ 2) - ShadowSize + m_ShadowOffsetX, Y + (H \ 2) - ShadowSize + m_ShadowOffsetY, 1)
                GdipDrawImageRectRectI hGraphics2, hPictureShadow, W \ 2, H \ 2, -W, -H, 0, 0, m_PictureRealWidth + ShadowSize * 2, m_PictureRealHeight + ShadowSize * 2, UnitPixel, hAttributes
                GdipResetWorldTransform hGraphics2
            Else
                GdipDrawImageRectRectI hGraphics2, hPictureShadow, X - ShadowSize + m_ShadowOffsetX, Y - ShadowSize + m_ShadowOffsetY, ReqWidth + ShadowSize * 2, ReqHeight + ShadowSize * 2, 0, 0, m_PictureRealWidth + ShadowSize * 2, m_PictureRealHeight + ShadowSize * 2, UnitPixel, hAttributes
            End If
            GdipDisposeImage hPictureShadow
            
            tMatrixColor.M(3, 3) = m_PictureOpacity / 100
        End If
                
        GdipSetImageAttributesColorMatrix hAttributes, &H0, True, tMatrixColor, tMatrixGray, &H0

        If m_PictureAngle <> 0 Then
            Call GdipRotateWorldTransform(hGraphics2, m_PictureAngle + 180, 0)
            Call GdipTranslateWorldTransform(hGraphics2, X + (ReqWidth \ 2), Y + (ReqHeight \ 2), 1)
            GdipDrawImageRectRectI hGraphics2, hImage, ReqWidth \ 2, ReqHeight \ 2, -ReqWidth, -ReqHeight, 0, 0, m_PictureRealWidth, m_PictureRealHeight, UnitPixel, hAttributes
        Else
            GdipDrawImageRectRectI hGraphics2, hImage, X, Y, ReqWidth, ReqHeight, 0, 0, m_PictureRealWidth, m_PictureRealHeight, UnitPixel, hAttributes
        End If
                
        GdipDisposeImage hImage
        Call GdipDisposeImageAttributes(hAttributes)
        
        GdipCreateTexture hImage2, &H0, hBrush
        
        CreateBrushTexture = hBrush

        Call GdipDeleteGraphics(hGraphics2)
        Call GdipDisposeImage(hImage2)
        
    End If
End Function

Private Function GetSafeRound(Angle As Integer, Width As Long, Height As Long) As Integer
    Dim lRet As Integer
    lRet = Angle
    If lRet * 2 > Height Then lRet = Height \ 2
    If lRet * 2 > Width Then lRet = Width \ 2
    GetSafeRound = lRet
End Function


Private Function CreateBlurShadowImage(ByVal hImage As Long, ByVal Color As Long, blurDepth As Integer, _
                                        Optional ByVal Left As Long, Optional ByVal Top As Long, _
                                        Optional ByVal Width As Long, Optional ByVal Height As Long) As Long
                                        
    Dim REC As RectL
    Dim X As Long, Y As Long
    Dim hImgShadow As Long
    Dim bmpData1 As BitmapData
    Dim bmpData2 As BitmapData
    Dim t2xBlur As Long
    Dim R As Long, G As Long, B As Long
    Dim Alpha As Byte
    Dim lSrcAlpha As Long, lDestAlpha As Long
    Dim dBytes() As Byte
    Dim srcBytes() As Byte
    Dim vTally() As Long
    Dim tAlpha As Long, tColumn As Long, tAvg As Long
    Dim initY As Long, initYstop As Long, initYstart As Long
    Dim initX As Long, initXstop As Long
    
    If hImage = 0& Then Exit Function
 
    If Width = 0& Then Call GdipGetImageWidth(hImage, Width)
    If Height = 0& Then Call GdipGetImageHeight(hImage, Height)
 
    t2xBlur = blurDepth * 2
 
    R = Color And &HFF
    G = (Color \ &H100&) And &HFF
    B = (Color \ &H10000) And &HFF
 
    SetRect REC, Left, Top, Width, Height
 
    ReDim srcBytes(REC.Width * 4 - 1&, REC.Height - 1&)
  
    With bmpData1
        .Scan0Ptr = VarPtr(srcBytes(0&, 0&))
        .stride = 4& * REC.Width
    End With
   
    Call GdipBitmapLockBits(hImage, REC, ImageLockModeUserInputBuf Or ImageLockModeRead, PixelFormat32bppPARGB, bmpData1)
 
    SetRect REC, Left, Top, Width + t2xBlur, Height + t2xBlur
    
    Call GdipCreateBitmapFromScan0(REC.Width, REC.Height, 0&, PixelFormat32bppPARGB, ByVal 0&, hImgShadow)

    ReDim dBytes(REC.Width * 4 - 1&, REC.Height - 1&)
    
    With bmpData2
        .Scan0Ptr = VarPtr(dBytes(0&, 0&))
        .stride = 4& * REC.Width
    End With
    
    Call GdipBitmapLockBits(hImgShadow, REC, ImageLockModeUserInputBuf Or ImageLockModeRead Or ImageLockModeWrite, PixelFormat32bppPARGB, bmpData2)
 
    R = Color And &HFF
    G = (Color \ &H100&) And &HFF
    B = (Color \ &H10000) And &HFF
    
    tAvg = (t2xBlur + 1) * (t2xBlur + 1)    ' how many pixels are being blurred
    
    ReDim vTally(0 To t2xBlur)              ' number of blur columns per pixel
    
    For Y = 0 To Height + t2xBlur - 1     ' loop thru shadow dib
    
        FillMemory vTally(0), (t2xBlur + 1) * 4, 0  ' reset column totals
        
        If Y < t2xBlur Then         ' y does not exist in source
            initYstart = 0          ' use 1st row
        Else
            initYstart = Y - t2xBlur ' start n blur rows above y
        End If
        ' how may source rows can we use for blurring?
        If Y < Height Then initYstop = Y Else initYstop = Height - 1
        
        tAlpha = 0  ' reset alpha sum
        tColumn = 0    ' reset column counter
        
        ' the first n columns will all be zero
        ' only the far right blur column has values; tally them
        For initY = initYstart To initYstop
            tAlpha = tAlpha + srcBytes(3, initY)
        Next
        ' assign the right column value
        vTally(t2xBlur) = tAlpha
        
        For X = 3 To (Width - 2) * 4 - 1 Step 4
            ' loop thru each source pixel's alpha
            
            ' set shadow alpha using blur average
            dBytes(X, Y) = tAlpha \ tAvg
            ' and set shadow color
            Select Case dBytes(X, Y)
            Case 255
                dBytes(X - 1, Y) = R
                dBytes(X - 2, Y) = G
                dBytes(X - 3, Y) = B
            Case 0
            Case Else
                dBytes(X - 1, Y) = R * dBytes(X, Y) \ 255
                dBytes(X - 2, Y) = G * dBytes(X, Y) \ 255
                dBytes(X - 3, Y) = B * dBytes(X, Y) \ 255
            End Select
            ' remove the furthest left column's alpha sum
            tAlpha = tAlpha - vTally(tColumn)
            ' count the next column of alphas
            vTally(tColumn) = 0&
            For initY = initYstart To initYstop
                vTally(tColumn) = vTally(tColumn) + srcBytes(X + 4, initY)
            Next
            ' add the new column's sum to the overall sum
            tAlpha = tAlpha + vTally(tColumn)
            ' set the next column to be recalculated
            tColumn = (tColumn + 1) Mod (t2xBlur + 1)
        Next
        
        ' now to finish blurring from right edge of source
        For X = X To (Width + t2xBlur - 1) * 4 - 1 Step 4
            dBytes(X, Y) = tAlpha \ tAvg
            Select Case dBytes(X, Y)
            Case 255
                dBytes(X - 1, Y) = R
                dBytes(X - 2, Y) = G
                dBytes(X - 3, Y) = B
            Case 0
            Case Else
                dBytes(X - 1, Y) = R * dBytes(X, Y) \ 255
                dBytes(X - 2, Y) = G * dBytes(X, Y) \ 255
                dBytes(X - 3, Y) = B * dBytes(X, Y) \ 255
            End Select
            ' remove this column's alpha sum
            tAlpha = tAlpha - vTally(tColumn)
            ' set next column to be removed
            tColumn = (tColumn + 1) Mod (t2xBlur + 1)
        Next
    Next
 
    Call GdipBitmapUnlockBits(hImage, bmpData1)
    Call GdipBitmapUnlockBits(hImgShadow, bmpData2)
    
    CreateBlurShadowImage = hImgShadow
End Function

Public Property Let AutoSize(ByVal NewValue As Boolean)
    Dim hGraphics As Long, hImage As Long
    Dim lWidth As Long, lHeight As Long
    Dim lDif As Long
    
    m_AutoSize = NewValue
    If m_AutoSize = False Then Exit Property
    
    GdipCreateBitmapFromScan0 UserControl.ScaleWidth, UserControl.ScaleHeight, 0&, PixelFormat32bppPARGB, ByVal 0&, hImage
    GdipGetImageGraphicsContext hImage, hGraphics
    
    lDif = ((m_BorderWidth * 2) + (m_CaptionPaddingX * 2))
    If m_BackShadow Then lDif = lDif + (m_ShadowSize * 2)
    If m_CallOut = True And m_CallOutPosicion = coLeft Or m_CallOutPosicion = coRight Then
        lDif = lDif + m_coLen
    End If
    lDif = lDif * nScale
    
    If m_WordWrap Then
        lWidth = UserControl.ScaleWidth - lDif
    Else
        lWidth = Screen.Width
    End If
    
    GDIP_AddPathString hGraphics, 0, 0, lWidth, lHeight, False, True
    lWidth = lWidth + lDif + 1 'NO SE QUE FALLA QUE DEVO SUMAR 1
    lDif = ((m_BorderWidth * 2) + (m_CaptionPaddingY * 2))
    If m_BackShadow Then lDif = lDif + (m_ShadowSize * 2)
    If m_CallOut = True And m_CallOutPosicion = coTop Or m_CallOutPosicion = coBottom Then
        lDif = lDif + m_coLen
    End If
    lDif = lDif * nScale
    lHeight = lHeight + lDif
    
    UserControl.Size (lWidth + 1) * Screen.TwipsPerPixelX, (lHeight + 1) * Screen.TwipsPerPixelY
    
    GdipDeleteGraphics hGraphics
    GdipDisposeImage hImage
End Property

Public Property Get AutoSize() As Boolean
    AutoSize = m_AutoSize
End Property

Public Property Get PictureColorize() As Boolean
    PictureColorize = m_PictureColorize
End Property

Public Property Let PictureColorize(ByVal New_Value As Boolean)
    m_PictureColorize = New_Value
    PropertyChanged "PictureColorize"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get PictureShadow() As Boolean
    PictureShadow = m_PictureShadow
End Property

Public Property Let PictureShadow(ByVal New_Value As Boolean)
    m_PictureShadow = New_Value
    PropertyChanged "PictureShadow"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get PictureColor() As OLE_COLOR
    PictureColor = m_PictureColor
End Property

Public Property Let PictureColor(ByVal New_Value As OLE_COLOR)
    m_PictureColor = New_Value
    PropertyChanged "PictureColor"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get PictureSetHeight() As Long
    PictureSetHeight = m_PictureSetHeight
End Property

Public Property Let PictureSetHeight(ByVal New_Value As Long)
    m_PictureSetHeight = New_Value
    PropertyChanged "PictureSetHeight"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get PictureGetHeight() As Long
    PictureGetHeight = m_PictureRealHeight
End Property

Public Property Get PictureSetWidth() As Long
    PictureSetWidth = m_PictureSetWidth
End Property

Public Property Let PictureSetWidth(ByVal New_Value As Long)
    m_PictureSetWidth = New_Value
    PropertyChanged "PictureSetWidth"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get PictureGetWidth() As Long
    PictureGetWidth = m_PictureRealWidth
End Property

Public Property Get PictureContrast() As Long
    PictureContrast = m_PictureContrast
End Property

Public Property Let PictureContrast(ByVal New_Value As Long)
    m_PictureContrast = New_Value
    SafeRange m_PictureContrast, -100, 100
    PropertyChanged "PictureContrast"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get PictureBrightness() As Long
    PictureBrightness = m_PictureBrightness
End Property

Public Property Let PictureBrightness(ByVal New_Value As Long)
    m_PictureBrightness = New_Value
    SafeRange m_PictureBrightness, -100, 100
    PropertyChanged "PictureBrightness"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get PictureGrayScale() As Boolean
    PictureGrayScale = m_PictureGraysScale
End Property

Public Property Let PictureGrayScale(ByVal New_Value As Boolean)
    m_PictureGraysScale = New_Value
    PropertyChanged "PictureGrayScale"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get CallOutCustomPosition() As Long
    CallOutCustomPosition = m_coCustomPos
End Property

Public Property Let CallOutCustomPosition(ByVal New_Value As Long)
    m_coCustomPos = New_Value
    PropertyChanged "CallOutCustomPosition"
    CreateShadow
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get CallOutRightTriangle() As Boolean
    CallOutRightTriangle = m_coRightTriangle
End Property

Public Property Let CallOutRightTriangle(ByVal New_Value As Boolean)
    m_coRightTriangle = New_Value
    PropertyChanged "CallOutRightTriangle"
    CreateShadow
    Refresh
End Property

Public Property Get CallOut() As Boolean
    CallOut = m_CallOut
End Property

Public Property Let CallOut(ByVal New_Value As Boolean)
    m_CallOut = New_Value
    PropertyChanged "CallOut"
    CreateShadow
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property
    
Public Property Get CallOutWidth() As Integer
    CallOutWidth = m_coWidth
End Property

Public Property Let CallOutWidth(ByVal New_Value As Integer)
    m_coWidth = New_Value
    PropertyChanged "CallOutWidth"
    CreateShadow
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property
    
Public Property Get CallOutLen() As Integer
    CallOutLen = m_coLen
End Property

Public Property Let CallOutLen(ByVal New_Value As Integer)
    m_coLen = New_Value
    PropertyChanged "CallOutLen"
    CreateShadow
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get CallOutPosicion() As eCallOutPosition
    CallOutPosicion = m_CallOutPosicion
End Property

Public Property Let CallOutPosicion(ByVal New_Value As eCallOutPosition)
    m_CallOutPosicion = New_Value
    PropertyChanged "CallOutPosicion"
    CreateShadow
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get CallOutAlign() As eCallOutAlign
    CallOutAlign = m_CallOutAlign
End Property

Public Property Let CallOutAlign(ByVal New_Value As eCallOutAlign)
    m_CallOutAlign = New_Value
    PropertyChanged "CallOutAlign"
    CreateShadow
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get CaptionTriming() As StringTrimming
    CaptionTriming = m_CaptionTriming
End Property

Public Property Let CaptionTriming(ByVal New_Value As StringTrimming)
    m_CaptionTriming = New_Value
    PropertyChanged "CaptionTriming"
    Refresh
End Property

Public Property Get ShadowOffsetX() As Integer
    ShadowOffsetX = m_ShadowOffsetX
End Property

Public Property Let ShadowOffsetX(ByVal New_Value As Integer)
    m_ShadowOffsetX = New_Value
    PropertyChanged "ShadowOffsetX"
    If (m_PictureBrush <> 0) And (m_PictureShadow = True) Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get ShadowOffsetY() As Integer
    ShadowOffsetY = m_ShadowOffsetY
End Property

Public Property Let ShadowOffsetY(ByVal New_Value As Integer)
    m_ShadowOffsetY = New_Value
    PropertyChanged "ShadowOffsetY"
    If (m_PictureBrush <> 0) And (m_PictureShadow = True) Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get ShadowColor() As OLE_COLOR
    ShadowColor = m_ShadowColor
End Property

Public Property Let ShadowColor(ByVal New_Value As OLE_COLOR)
    m_ShadowColor = New_Value
    PropertyChanged "ShadowColor"
    If (m_PictureBrush <> 0) And (m_PictureShadow = True) Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    CreateShadow
    Refresh
End Property

Public Property Get ShadowColorOpacity() As Integer
    ShadowColorOpacity = m_ShadowColorOpacity
End Property

Public Property Let ShadowColorOpacity(ByVal New_Value As Integer)
    m_ShadowColorOpacity = New_Value
    SafeRange m_ShadowColorOpacity, 0, 100
    PropertyChanged "ShadowColorOpacity"
    If (m_PictureBrush <> 0) And (m_PictureShadow = True) Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    CreateShadow
    Refresh
End Property

Public Property Get ShadowSize() As Integer
    ShadowSize = m_ShadowSize
End Property

Public Property Let ShadowSize(ByVal New_Value As Integer)
    m_ShadowSize = New_Value
    SafeRange m_ShadowSize, 0, 100
    PropertyChanged "ShadowSize"
    CreateShadow
    If (m_PictureBrush <> 0) And (m_PictureShadow = True) Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get BorderCornerLeftTop() As Integer
    BorderCornerLeftTop = m_BorderCornerLeftTop
End Property

Public Property Let BorderCornerLeftTop(ByVal New_Value As Integer)
    m_BorderCornerLeftTop = New_Value
    PropertyChanged "BorderCornerLeftTop"
    CreateShadow
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get BorderCornerRightTop() As Integer
    BorderCornerRightTop = m_BorderCornerRightTop
End Property

Public Property Let BorderCornerRightTop(ByVal New_Value As Integer)
    m_BorderCornerRightTop = New_Value
    PropertyChanged "BorderCornerRightTop"
    CreateShadow
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get BorderCornerBottomLeft() As Integer
    BorderCornerBottomLeft = m_BorderCornerBottomLeft
End Property

Public Property Let BorderCornerBottomLeft(ByVal New_Value As Integer)
    m_BorderCornerBottomLeft = New_Value
    PropertyChanged "BorderCornerBottomLeft"
    CreateShadow
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get BorderCornerBottomRight() As Integer
    BorderCornerBottomRight = m_BorderCornerBottomRight
End Property

Public Property Let BorderCornerBottomRight(ByVal New_Value As Integer)
    m_BorderCornerBottomRight = New_Value
    PropertyChanged "BorderCornerBottomRight"
    CreateShadow
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    Refresh
End Property

Public Property Get BackColorOpacity() As Integer
    BackColorOpacity = m_BackColorOpacity
End Property

Public Property Let BackColorOpacity(ByVal New_BackColorOpacity As Integer)
    m_BackColorOpacity = New_BackColorOpacity
    SafeRange m_BackColorOpacity, 0, 100
    PropertyChanged "BackColorOpacity"
    CreateShadow
    Refresh
End Property

Public Property Get BackAcrylicBlur() As Boolean
    BackAcrylicBlur = m_BackAcrylicBlur
End Property

Public Property Let BackAcrylicBlur(ByVal New_Value As Boolean)
    m_BackAcrylicBlur = New_Value
    PropertyChanged "BackAcrylicBlur"
    
    If New_Value Then
        CreateBuffer
    Else
        If OldhBmp Then DeleteObject SelectObject(hDCMemory, OldhBmp): OldhBmp = 0
        If hDCMemory Then DeleteDC hDCMemory: hDCMemory = 0
    End If
    CreateShadow
    Refresh
End Property

Public Property Get BackShadow() As Boolean
    BackShadow = m_BackShadow
End Property

Public Property Let BackShadow(ByVal New_Value As Boolean)
    m_BackShadow = New_Value
    PropertyChanged "BackShadow"
    CreateShadow
    Refresh
End Property

Public Property Get Border() As Boolean
    Border = m_Border
End Property

Public Property Let Border(ByVal New_Border As Boolean)
    m_Border = New_Border
    PropertyChanged "Border"
    CreateShadow
    Refresh
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    m_BorderColor = New_BorderColor
    PropertyChanged "BorderColor"
    Refresh
End Property

Public Property Get BorderColorOpacity() As Integer
    BorderColorOpacity = m_BorderColorOpacity
End Property

Public Property Let BorderColorOpacity(ByVal New_BorderColorOpacity As Integer)
    m_BorderColorOpacity = New_BorderColorOpacity
    SafeRange m_BorderColorOpacity, 0, 100
    PropertyChanged "BorderColorOpacity"
    Refresh
End Property

Public Property Get BorderPosition() As eBorderPosition
    BorderPosition = m_BorderPosition
End Property

Public Property Let BorderPosition(ByVal New_BorderPosition As eBorderPosition)
    m_BorderPosition = New_BorderPosition
    PropertyChanged "BorderPosition"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    CreateShadow
    Refresh
End Property

Public Property Get BorderWidth() As Integer
    BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Integer)
    m_BorderWidth = New_BorderWidth
    PropertyChanged "BorderWidth"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    CreateShadow
    Refresh
End Property

Public Property Get CaptionShadow() As Boolean
    CaptionShadow = m_CaptionShadow
End Property

Public Property Let CaptionShadow(ByVal New_Value As Boolean)
    m_CaptionShadow = New_Value
    PropertyChanged "CaptionShadow"
    If hImgCaptionShadow Then GdipDisposeImage hImgCaptionShadow: hImgCaptionShadow = 0
    bRecreateShadowCaption = True
    Refresh
End Property


Public Property Get CaptionShowPrefix() As Boolean
    CaptionShowPrefix = m_CaptionShowPrefix
End Property

Public Property Let CaptionShowPrefix(ByVal New_Value As Boolean)
    m_CaptionShowPrefix = New_Value
    PropertyChanged "CaptionShowPrefix"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get CaptionAngle() As Integer
    CaptionAngle = m_CaptionAngle
End Property

Public Property Let CaptionAngle(ByVal New_Value As Integer)
    m_CaptionAngle = New_Value
    PropertyChanged "CaptionAngle"
    'If hImgCaptionShadow Then GdipDisposeImage hImgCaptionShadow: hImgCaptionShadow = 0
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get CaptionBorderColor() As OLE_COLOR
    CaptionBorderColor = m_CaptionBorderColor
End Property

Public Property Let CaptionBorderColor(ByVal New_Value As OLE_COLOR)
    m_CaptionBorderColor = New_Value
    PropertyChanged "CaptionBorderColor"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get CaptionBorderWidth() As Integer
    CaptionBorderWidth = m_CaptionBorderWidth
End Property

Public Property Let CaptionBorderWidth(ByVal New_Value As Integer)
    m_CaptionBorderWidth = New_Value
    If m_CaptionBorderWidth < 0 Then m_CaptionBorderWidth = 0
    PropertyChanged "CaptionBorderWidth"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByRef New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    If m_AutoSize Then Me.AutoSize = True
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get CaptionAlignmentH() As CaptionAlignmentH
    CaptionAlignmentH = m_CaptionAlignmentH
End Property

Public Property Let CaptionAlignmentH(ByVal New_CaptionAlignmentH As CaptionAlignmentH)
    m_CaptionAlignmentH = New_CaptionAlignmentH
    PropertyChanged "CaptionAlignmentH"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get CaptionAlignmentV() As CaptionAlignmentV
    CaptionAlignmentV = m_CaptionAlignmentV
End Property

Public Property Let CaptionAlignmentV(ByVal New_CaptionAlignmentV As CaptionAlignmentV)
    m_CaptionAlignmentV = New_CaptionAlignmentV
    PropertyChanged "CaptionAlignmentV"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get CaptionPaddingX() As Integer
    CaptionPaddingX = m_CaptionPaddingX
End Property

Public Property Let CaptionPaddingX(ByVal New_CaptionPaddingX As Integer)
    m_CaptionPaddingX = New_CaptionPaddingX
    PropertyChanged "CaptionPaddingX"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get CaptionPaddingY() As Integer
    CaptionPaddingY = m_CaptionPaddingY
End Property

Public Property Let CaptionPaddingY(ByVal New_CaptionPaddingY As Integer)
    m_CaptionPaddingY = New_CaptionPaddingY
    PropertyChanged "CaptionPaddingY"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    UserControl.Enabled = NewValue
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As StdFont
    Set Font = m_Font
End Property

Public Property Set Font(New_Font As StdFont)
    With m_Font
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
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    Refresh
End Property

Public Property Get ForeColorOpacity() As Integer
    ForeColorOpacity = m_ForeColorOpacity
End Property

Public Property Let ForeColorOpacity(ByVal New_ForeColorOpacity As Integer)
    m_ForeColorOpacity = New_ForeColorOpacity
    SafeRange m_ForeColorOpacity, 0, 100
    PropertyChanged "ForeColorOpacity"
    Refresh
End Property

Public Property Get Gradient() As Boolean
    Gradient = m_Gradient
End Property

Public Property Let Gradient(ByVal New_Gradient As Boolean)
    m_Gradient = New_Gradient
    PropertyChanged "Gradient"
    Refresh
End Property

Public Property Get GradientAngle() As Integer
    GradientAngle = m_GradientAngle
End Property

Public Property Let GradientAngle(ByVal New_GradientAngle As Integer)
    m_GradientAngle = New_GradientAngle
    SafeRange m_GradientAngle, 0, 359
    PropertyChanged "GradientAngle"
    Refresh
End Property

Public Property Get GradientColor1() As OLE_COLOR
    GradientColor1 = m_GradientColor1
End Property

Public Property Let GradientColor1(ByVal New_GradientColor1 As OLE_COLOR)
    m_GradientColor1 = New_GradientColor1
    PropertyChanged "GradientColor1"
    Refresh
End Property

Public Property Get GradientColor1Opacity() As Integer
    GradientColor1Opacity = m_GradientColor1Opacity
End Property

Public Property Let GradientColor1Opacity(ByVal New_GradientColor1Opacity As Integer)
    m_GradientColor1Opacity = New_GradientColor1Opacity
    SafeRange m_GradientColor1Opacity, 0, 100
    PropertyChanged "GradientColor1Opacity"
    Refresh
End Property

Public Property Get GradientColor2() As OLE_COLOR
    GradientColor2 = m_GradientColor2
End Property

Public Property Let GradientColor2(ByVal New_GradientColor2 As OLE_COLOR)
    m_GradientColor2 = New_GradientColor2
    PropertyChanged "GradientColor2"
    Refresh
End Property

Public Property Get GradientColor2Opacity() As Integer
    GradientColor2Opacity = m_GradientColor2Opacity
End Property

Public Property Let GradientColor2Opacity(ByVal New_GradientColor2Opacity As Integer)
    m_GradientColor2Opacity = New_GradientColor2Opacity
    SafeRange m_GradientColor2Opacity, 0, 100
    PropertyChanged "GradientColor2Opacity"
    Refresh
End Property

Public Property Get PictureAngle() As Integer
    PictureAngle = m_PictureAngle
End Property

Public Property Let PictureAngle(ByVal New_Value As Integer)
    m_PictureAngle = New_Value
    PropertyChanged "PictureAngle"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get PictureAlignmentH() As PictureAlignmentH
    PictureAlignmentH = m_PictureAlignmentH
End Property

Public Property Let PictureAlignmentH(ByVal New_PictureAlignmentH As PictureAlignmentH)
    m_PictureAlignmentH = New_PictureAlignmentH
    PropertyChanged "PictureAlignmentH"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get PictureAlignmentV() As PictureAlignmentV
    PictureAlignmentV = m_PictureAlignmentV
End Property

Public Property Let PictureAlignmentV(ByVal New_PictureAlignmentV As PictureAlignmentV)
    m_PictureAlignmentV = New_PictureAlignmentV
    PropertyChanged "PictureAlignmentV"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get PictureOpacity() As Integer
    PictureOpacity = m_PictureOpacity
End Property

Public Property Let PictureOpacity(ByVal New_PictureOpacity As Integer)
    m_PictureOpacity = New_PictureOpacity
    SafeRange m_PictureOpacity, 0, 100
    PropertyChanged "PictureOpacity"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get PicturePaddingY() As Integer
    PicturePaddingY = m_PicturePaddingY
End Property

Public Property Let PicturePaddingY(ByVal New_PicturePaddingY As Integer)
    m_PicturePaddingY = New_PicturePaddingY
    PropertyChanged "PicturePaddingY"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get PicturePaddingX() As Integer
    PicturePaddingX = m_PicturePaddingX
End Property

Public Property Let PicturePaddingX(ByVal New_PicturePaddingX As Integer)
    m_PicturePaddingX = New_PicturePaddingX
    PropertyChanged "PicturePaddingX"
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Refresh
End Property

Public Property Get WordWrap() As Boolean
    WordWrap = m_WordWrap
End Property

Public Property Let WordWrap(ByVal New_WordWrap As Boolean)
    m_WordWrap = New_WordWrap
    PropertyChanged "WordWrap"
    Refresh
End Property

Public Function PictureGetStream() As Byte()
    PictureGetStream = m_PictureArr
End Function

Public Function PictureDelete()
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Erase m_PictureArr
    m_PicturePresent = False
    Call PropertyChanged("PicturePresent")
    Call PropertyChanged("PictureArr")
    Refresh
End Function

Public Property Get PictureExist() As Boolean
    PictureExist = m_PicturePresent
End Property

Public Property Get MouseToParent() As Boolean
    MouseToParent = m_MouseToParent
End Property

Public Property Let MouseToParent(ByVal New_Value As Boolean)
    m_MouseToParent = New_Value
    PropertyChanged "MouseToParent"
End Property

Public Property Let OLEDropMode(ByVal New_Value As OLEDropConstants)
    UserControl.OLEDropMode = New_Value
End Property
Public Property Get OLEDropMode() As OLEDropConstants
    OLEDropMode = UserControl.OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Public Property Get HotLine() As Boolean
    HotLine = m_HotLine
End Property

Public Property Let HotLine(ByVal New_Value As Boolean)
    m_HotLine = New_Value
    PropertyChanged "HotLine"
    Refresh
End Property

Public Property Get HotLineColor() As OLE_COLOR
    HotLineColor = m_HotLineColor
End Property

Public Property Let HotLineColor(ByVal New_Value As OLE_COLOR)
    m_HotLineColor = New_Value
    PropertyChanged "HotLineColor "
    Refresh
End Property

Public Property Get HotLineColorOpacity() As Integer
    HotLineColorOpacity = m_HotLineColorOpacity
End Property

Public Property Let HotLineColorOpacity(ByVal New_Value As Integer)
    m_HotLineColorOpacity = New_Value
    SafeRange m_HotLineColorOpacity, 0, 100
    PropertyChanged "HotLineColorOpacity "
    Refresh
End Property

Public Property Get HotLineWidth() As Integer
    HotLineWidth = m_HotLineWidth
End Property

Public Property Let HotLineWidth(ByVal New_Value As Integer)
    m_HotLineWidth = New_Value
    PropertyChanged "HotLineWidth"
    Refresh
End Property

Public Property Get HotLinePosition() As HotLinePosition
    HotLinePosition = m_HotLinePosition
End Property

Public Property Let HotLinePosition(ByVal New_Value As HotLinePosition)
    m_HotLinePosition = New_Value
    PropertyChanged "HotLinePosition"
    Refresh
End Property

Public Property Get Redraw() As Boolean
    Redraw = m_Redraw
End Property

Public Property Let Redraw(ByVal New_Value As Boolean)
    m_Redraw = New_Value
End Property

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    On Error Resume Next

    If UserControl.Enabled Then
        If Not MouseToParent Then
            HitResult = vbHitResultHit
        ElseIf Not Ambient.UserMode Then
            HitResult = vbHitResultHit
        End If
        If Ambient.UserMode Then
            Dim PT As POINTAPI
            Dim lHwnd As Long
            GetCursorPos PT
            lHwnd = WindowFromPoint(PT.X, PT.Y)
            
            If m_Enter = False Then

                ScreenToClient c_lhWnd, PT
                m_PT.X = PT.X - X
                m_PT.Y = PT.Y - Y
    
                m_Left = ScaleX(Extender.Left, vbContainerSize, UserControl.ScaleMode)
                m_Top = ScaleY(Extender.Top, vbContainerSize, UserControl.ScaleMode)
 
                m_Enter = True
                tmrMOUSEOVER.Interval = 1
                 RaiseEvent MouseEnter
            End If
        
            bIntercept = True
            
            If lHwnd = c_lhWnd Then
                If m_Over = False Then
                    m_Over = True
                    RaiseEvent MouseOver
                End If
            Else
                If m_Over = True Then
                    m_Over = False
                    RaiseEvent MouseOut
                End If
            End If
        End If
    ElseIf Not Ambient.UserMode Then
        HitResult = vbHitResultHit
    End If
End Sub

Private Sub UserControl_Initialize()
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
    nScale = GetWindowsDPI
    m_Redraw = True
End Sub

Private Sub UserControl_InitProperties()
    hFontCollection = ReadValue(&HFC)
    m_BackColor = Ambient.BackColor
    m_BackColorOpacity = 100
    m_BorderColor = vbActiveBorder
    m_BorderColorOpacity = 100
    m_BorderPosition = bpCenter
    m_Caption = Ambient.DisplayName
    m_CaptionBorderColor = vbHighlightText
    Set m_Font = UserControl.Ambient.Font
    m_ForeColor = vbButtonText
    m_ForeColorOpacity = 100
    m_GradientColor1 = &HD3A042
    m_GradientColor1Opacity = 100
    m_GradientColor2 = &HE96E9B
    m_GradientColor2Opacity = 100
    m_PictureOpacity = 100
    m_HotLineColor = vbHighlight
    m_HotLineColorOpacity = 100
    m_HotLineWidth = 5&
    m_HotLinePosition = hlBottom
    m_WordWrap = True
    Set m_IconFont = UserControl.Ambient.Font
    c_lhWnd = UserControl.ContainerHwnd
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
     RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub UserControl_Paint()
    Dim lHdc As Long
    Dim X As Long, Y As Long
    lHdc = UserControl.HDC
    RaiseEvent PrePaint(lHdc, X, Y)
    Call Draw(lHdc, 0, X, Y)
    RaiseEvent PostPaint(UserControl.HDC)
End Sub

Public Sub Refresh()
    If m_Redraw Then
        UserControl.Refresh
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    hFontCollection = ReadValue(&HFC)
    c_lhWnd = UserControl.ContainerHwnd
    
    With PropBag
        m_BackColor = .ReadProperty("BackColor", Ambient.BackColor)
        m_BackColorOpacity = .ReadProperty("BackColorOpacity", 100)
        m_BackAcrylicBlur = .ReadProperty("BackAcrylicBlur", False)
        m_BackShadow = .ReadProperty("BackShadow", True)
        m_Border = .ReadProperty("Border", False)
        m_BorderColor = .ReadProperty("BorderColor", vbActiveBorder)
        m_BorderColorOpacity = .ReadProperty("BorderColorOpacity", 100)
        m_BorderCornerLeftTop = .ReadProperty("BorderCornerLeftTop", 0)
        m_BorderCornerRightTop = .ReadProperty("BorderCornerRightTop", 0)
        m_BorderCornerBottomRight = .ReadProperty("BorderCornerBottomRight", 0)
        m_BorderCornerBottomLeft = .ReadProperty("BorderCornerBottomLeft", 0)
        m_BorderPosition = .ReadProperty("BorderPosition", bpCenter)
        m_BorderWidth = .ReadProperty("BorderWidth", 0)
        m_CaptionAlignmentH = .ReadProperty("CaptionAlignmentH", cLeft)
        m_CaptionAlignmentV = .ReadProperty("CaptionAlignmentV", cTop)
        m_Caption = .ReadProperty("Caption", Ambient.DisplayName)
        m_CaptionPaddingX = .ReadProperty("CaptionPaddingX", 0)
        m_CaptionPaddingY = .ReadProperty("CaptionPaddingY", 0)
        m_CaptionTriming = .ReadProperty("CaptionTriming", StringTrimmingNone)
        m_CaptionBorderWidth = .ReadProperty("CaptionBorderWidth", 0)
        m_CaptionBorderColor = .ReadProperty("CaptionBorderColor", vbHighlightText)
        m_CaptionShadow = .ReadProperty("CaptionShadow", False)
        m_CaptionAngle = .ReadProperty("CaptionAngle", 0)
        m_CaptionShowPrefix = .ReadProperty("CaptionShowPrefix", False)
        UserControl.Enabled = .ReadProperty("Enabled", True)
        Set m_Font = .ReadProperty("Font", UserControl.Ambient.Font)
        m_ForeColor = .ReadProperty("ForeColor", vbButtonText)
        m_ForeColorOpacity = .ReadProperty("ForeColorOpacity", 100)
        m_Gradient = .ReadProperty("Gradient", False)
        m_GradientAngle = .ReadProperty("GradientAngle", 0)
        m_GradientColor1 = .ReadProperty("GradientColor1", &HD3A042)
        m_GradientColor1Opacity = .ReadProperty("GradientColor1Opacity", 100)
        m_GradientColor2 = .ReadProperty("GradientColor2", &HE96E9B)
        m_GradientColor2Opacity = .ReadProperty("GradientColor2Opacity", 100)
        m_PictureAngle = .ReadProperty("PictureAngle", 0)
        m_PictureAlignmentH = .ReadProperty("PictureAlignmentH", pLeft)
        m_PictureAlignmentV = .ReadProperty("PictureAlignmentV", pTop)
        m_PictureOpacity = .ReadProperty("PictureOpacity", 100)
        m_PictureBrightness = .ReadProperty("PictureBrightness", 0)
        m_PictureContrast = .ReadProperty("PictureContrast", 0)
        m_PictureGraysScale = .ReadProperty("PictureGraysScale", False)
        m_PicturePaddingX = .ReadProperty("PicturePaddingX", 0)
        m_PicturePaddingY = .ReadProperty("PicturePaddingY", 0)
        m_PictureSetWidth = .ReadProperty("PictureSetWidth", 0)
        m_PictureSetHeight = .ReadProperty("PictureSetHeight", 0)
        m_WordWrap = .ReadProperty("WordWrap", True)
        m_ShadowSize = .ReadProperty("ShadowSize", 0)
        m_ShadowColor = .ReadProperty("ShadowColor", vbBlack)
        m_ShadowOffsetX = .ReadProperty("ShadowOffsetX", 0)
        m_ShadowOffsetY = .ReadProperty("ShadowOffsetY", 0)
        m_ShadowColorOpacity = .ReadProperty("ShadowColorOpacity", 50)
        m_CallOutAlign = .ReadProperty("CallOutAlign", coMidle)
        m_CallOutPosicion = .ReadProperty("CallOutPosicion", coLeft)
        m_coWidth = .ReadProperty("CallOutWidth", 10)
        m_coLen = .ReadProperty("CallOutLen", 10)
        m_CallOut = .ReadProperty("CallOut", False)
        m_coCustomPos = .ReadProperty("CallOutCustomPosition", 0)
        m_coRightTriangle = .ReadProperty("CallOutRightTriangle", False)
        m_PictureColorize = .ReadProperty("PictureColorize", False)
        m_PictureShadow = .ReadProperty("PictureShadow", False)
        m_PictureColor = .ReadProperty("PictureColor", vbBlack)
        UserControl.MousePointer = .ReadProperty("MousePointer", vbArrow)
        UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
        m_MousePointerHands = .ReadProperty("MousePointerHands", False)
        m_MouseToParent = .ReadProperty("MouseToParent", False)
        UserControl.OLEDropMode = .ReadProperty("OLEDropMode", 0&)
        m_HotLine = .ReadProperty("HotLine", False)
        m_HotLineColor = .ReadProperty("HotLineColor", vbHighlight)
        m_HotLineColorOpacity = .ReadProperty("HotLineColorOpacity", 100)
        m_HotLineWidth = .ReadProperty("HotLineWidth", 5&)
        m_HotLinePosition = .ReadProperty("HotLinePosition", hlBottom)
        Set m_IconFont = .ReadProperty("IconFont", UserControl.Ambient.Font)
        m_IconCharCode = .ReadProperty("IconCharCode", 0)
        m_IconForeColor = .ReadProperty("IconForeColor", vbButtonText)
        m_IconPaddingX = .ReadProperty("IconPaddingX", 0)
        m_IconPaddingY = .ReadProperty("IconPaddingY", 0)
        m_IconAlignmentH = .ReadProperty("IconAlignmentH", 0)
        m_IconAlignmentV = .ReadProperty("IconAlignmentV", 0)
        m_IconOpacity = .ReadProperty("IconOpacity", 100)
        
        If m_MousePointerHands Then
            If Ambient.UserMode Then
                UserControl.MousePointer = vbCustom
                UserControl.MouseIcon = GetSystemHandCursor
            End If
        End If
    
        If CBool(.ReadProperty("PicturePresent", False)) Then
            m_PictureArr() = .ReadProperty("PictureArr")
            Call PictureFromStream(m_PictureArr)
        End If
        bRecreateShadowCaption = True
        CreateShadow
        If m_BackAcrylicBlur Then CreateBuffer
    End With
End Sub

Private Sub UserControl_Show()
    Me.Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("BackColor", m_BackColor, Ambient.BackColor)
        Call .WriteProperty("BackColorOpacity", m_BackColorOpacity, 100)
        Call .WriteProperty("BackAcrylicBlur", m_BackAcrylicBlur, False)
        Call .WriteProperty("BackShadow", m_BackShadow, True)
        Call .WriteProperty("Border", m_Border, False)
        Call .WriteProperty("BorderColor", m_BorderColor, vbActiveBorder)
        Call .WriteProperty("BorderColorOpacity", m_BorderColorOpacity, 100)
        Call .WriteProperty("BorderCornerLeftTop", m_BorderCornerLeftTop, 0)
        Call .WriteProperty("BorderCornerRightTop", m_BorderCornerRightTop, 0)
        Call .WriteProperty("BorderCornerBottomRight", m_BorderCornerBottomRight, 0)
        Call .WriteProperty("BorderCornerBottomLeft", m_BorderCornerBottomLeft, 0)
        Call .WriteProperty("BorderPosition", m_BorderPosition, bpCenter)
        Call .WriteProperty("BorderWidth", m_BorderWidth, 0)
        Call .WriteProperty("CaptionAlignmentH", m_CaptionAlignmentH, cLeft)
        Call .WriteProperty("CaptionAlignmentV", m_CaptionAlignmentV, cTop)
        Call .WriteProperty("Caption", m_Caption, Ambient.DisplayName)
        Call .WriteProperty("CaptionPaddingX", m_CaptionPaddingX, 0)
        Call .WriteProperty("CaptionPaddingY", m_CaptionPaddingY, 0)
        Call .WriteProperty("CaptionBorderWidth", m_CaptionBorderWidth, 0)
        Call .WriteProperty("CaptionBorderColor", m_CaptionBorderColor, vbHighlightText)
        Call .WriteProperty("CaptionShadow", m_CaptionShadow, False)
        Call .WriteProperty("CaptionAngle", m_CaptionAngle, 0)
        Call .WriteProperty("CaptionTriming", m_CaptionTriming, StringTrimmingNone)
        Call .WriteProperty("CaptionShowPrefix", m_CaptionShowPrefix, False)
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("Font", m_Font, UserControl.Ambient.Font)
        Call .WriteProperty("ForeColor", m_ForeColor, vbButtonText)
        Call .WriteProperty("ForeColorOpacity", m_ForeColorOpacity, 100)
        Call .WriteProperty("Gradient", m_Gradient, False)
        Call .WriteProperty("GradientAngle", m_GradientAngle, 0)
        Call .WriteProperty("GradientColor1", m_GradientColor1, &HD3A042)
        Call .WriteProperty("GradientColor1Opacity", m_GradientColor1Opacity, 100)
        Call .WriteProperty("GradientColor2", m_GradientColor2, &HE96E9B)
        Call .WriteProperty("GradientColor2Opacity", m_GradientColor2Opacity, 100)
        Call .WriteProperty("PictureAngle", m_PictureAngle, 0)
        Call .WriteProperty("PictureAlignmentH", m_PictureAlignmentH, pLeft)
        Call .WriteProperty("PictureAlignmentV", m_PictureAlignmentV, pTop)
        Call .WriteProperty("PictureOpacity", m_PictureOpacity, 100)
        Call .WriteProperty("PictureBrightness", m_PictureBrightness, 0)
        Call .WriteProperty("PictureContrast", m_PictureContrast, 0)
        Call .WriteProperty("PictureGraysScale", m_PictureGraysScale, False)
        Call .WriteProperty("PicturePaddingX", m_PicturePaddingX, 0)
        Call .WriteProperty("PicturePaddingY", m_PicturePaddingY, 0)
        Call .WriteProperty("PictureSetWidth", m_PictureSetWidth, 0)
        Call .WriteProperty("PictureSetHeight", m_PictureSetHeight, 0)
        Call .WriteProperty("WordWrap", m_WordWrap, True)
        Call .WriteProperty("ShadowSize", m_ShadowSize, 0)
        Call .WriteProperty("ShadowColor", m_ShadowColor, vbBlack)
        Call .WriteProperty("ShadowOffsetX", m_ShadowOffsetX, 0)
        Call .WriteProperty("ShadowOffsetY", m_ShadowOffsetY, 0)
        Call .WriteProperty("ShadowColorOpacity", m_ShadowColorOpacity, 50)
        Call .WriteProperty("CallOutAlign", m_CallOutAlign, coMidle)
        Call .WriteProperty("CallOutPosicion", m_CallOutPosicion, coLeft)
        Call .WriteProperty("CallOutWidth", m_coWidth, 10)
        Call .WriteProperty("CallOutLen", m_coLen, 10)
        Call .WriteProperty("CallOut", m_CallOut, False)
        Call .WriteProperty("CallOutCustomPosition", m_coCustomPos, 0)
        Call .WriteProperty("CallOutRightTriangle", m_coRightTriangle, 0)
        Call .WriteProperty("PictureColorize", m_PictureColorize, False)
        Call .WriteProperty("PictureShadow", m_PictureShadow, False)
        Call .WriteProperty("PictureColor", m_PictureColor, False)
        Call .WriteProperty("MousePointer", UserControl.MousePointer, vbArrow)
        Call .WriteProperty("MouseIcon", UserControl.MouseIcon, Nothing)
        Call .WriteProperty("MousePointerHands", m_MousePointerHands, False)
        Call .WriteProperty("MouseToParent", m_MouseToParent, False)
        Call .WriteProperty("OLEDropMode", UserControl.OLEDropMode, 0&)
        Call .WriteProperty("HotLine", m_HotLine, False)
        Call .WriteProperty("HotLineColor", m_HotLineColor, vbHighlight)
        Call .WriteProperty("HotLineColorOpacity", m_HotLineColorOpacity, 100)
        Call .WriteProperty("HotLineWidth", m_HotLineWidth, 5)
        Call .WriteProperty("HotLinePosition", m_HotLinePosition, hlBottom)
        Call .WriteProperty("IconFont", m_IconFont, UserControl.Ambient.Font)
        Call .WriteProperty("IconCharCode", m_IconCharCode, 0)
        Call .WriteProperty("IconForeColor", m_IconForeColor, vbButtonText)
        Call .WriteProperty("IconPaddingX", m_IconPaddingX, 0)
        Call .WriteProperty("IconPaddingY", m_IconPaddingY, 0)
        Call .WriteProperty("IconAlignmentH", m_IconAlignmentH, 0)
        Call .WriteProperty("IconAlignmentV", m_IconAlignmentV, 0)
        Call .WriteProperty("IconOpacity", m_IconOpacity, 100)
        
        Call .WriteProperty("PicturePresent", m_PicturePresent, False)
        If m_PicturePresent Then
            Call .WriteProperty("PictureArr", m_PictureArr, 0)
        Else
            Call .WriteProperty("PictureArr", 0)
        End If
        
    End With

End Sub

Private Sub UserControl_Resize()
    If m_BackAcrylicBlur Then CreateBuffer
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    CreateShadow
End Sub

Private Sub UserControl_Terminate()
    If m_PictureBrush Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    If hImgShadow Then GdipDisposeImage hImgShadow
    If hImgCaptionShadow Then GdipDisposeImage hImgCaptionShadow
    If hCur Then DestroyCursor hCur
    If OldhBmp Then DeleteObject SelectObject(hDCMemory, OldhBmp): OldhBmp = 0
    If hDCMemory Then DeleteDC hDCMemory: hDCMemory = 0
    Call GdiplusShutdown(GdipToken)
End Sub

Private Function ConvertColor(ByVal Color As Long, ByVal Opacity As Long) As Long
    Dim BGRA(0 To 3) As Byte
    OleTranslateColor Color, 0, VarPtr(Color)
  
    BGRA(3) = CByte((Abs(Opacity) / 100) * 255)
    BGRA(0) = ((Color \ &H10000) And &HFF)
    BGRA(1) = ((Color \ &H100) And &HFF)
    BGRA(2) = (Color And &HFF)
    CopyMemory ConvertColor, BGRA(0), 4&
End Function

Private Function RoundRectangle(X As Long, Y As Long, Width As Long, Height As Long, Optional Inflate As Boolean, Optional nn As Boolean) As Long
    Dim mPath As Long
    Dim BCLT As Integer
    Dim BCRT As Integer
    Dim BCBR As Integer
    Dim BCBL As Integer
    Dim XX As Long, YY As Long
    Dim MidBorder As Long
    Dim coLen As Long
    Dim coWidth As Long
    Dim lMax As Long
    Dim coAngle  As Long

    Width = Width - 1 'Antialias pixel
    Height = Height - 1 'Antialias pixel
        
    coWidth = m_coWidth * nScale
    coLen = m_coLen * nScale
    coAngle = IIf(m_coRightTriangle, 0, coWidth / 2)

    If nn Then
        If m_BorderPosition = bpCenter Then
            coWidth = coWidth + m_BorderWidth * nScale / 2
        ElseIf m_BorderPosition = bpOutside Then
            coWidth = coWidth + m_BorderWidth * nScale
        ElseIf m_BorderPosition = bpInside Then
            coWidth = coWidth - m_BorderWidth * nScale / 2
        End If
    End If
    

    If Inflate Then MidBorder = m_BorderWidth / 2
    BCLT = GetSafeRound((m_BorderCornerLeftTop + MidBorder) * nScale, Width, Height)
    BCRT = GetSafeRound((m_BorderCornerRightTop + MidBorder) * nScale, Width, Height)
    BCBR = GetSafeRound((m_BorderCornerBottomRight + MidBorder) * nScale, Width, Height)
    BCBL = GetSafeRound((m_BorderCornerBottomLeft + MidBorder) * nScale, Width, Height)

    If m_CallOut Then
        Select Case m_CallOutPosicion
            Case coLeft
                X = X + coLen
                Width = Width - coLen
                lMax = Height - BCLT - BCBL
                If coWidth > lMax Then coWidth = lMax
            Case coTop
                Y = Y + coLen
                Height = Height - coLen
                lMax = Width - BCLT - BCBL
                If coWidth > lMax Then coWidth = lMax
            Case coRight
                Width = Width - coLen
                lMax = Height - BCRT - BCBR
                If coWidth > lMax Then coWidth = lMax
            Case coBottom
                Height = Height - coLen
                lMax = Width - BCBL - BCBR
                If coWidth > lMax Then coWidth = lMax
        End Select
    End If

    Call GdipCreatePath(&H0, mPath)
                    
                    
    If BCLT Then GdipAddPathArcI mPath, X, Y, BCLT * 2, BCLT * 2, 180, 90

    If m_CallOutPosicion = coTop And m_CallOut Then
        Select Case m_CallOutAlign
            Case coFirstCorner: XX = X + BCLT
            Case coMidle: XX = X + BCLT + ((Width - BCLT - BCRT) \ 2) - (coWidth \ 2)
            Case coSecondCorner: XX = X + Width - coWidth - BCRT
            Case coCustomPosition: XX = X + (m_coCustomPos * nScale)
        End Select
        
        If (XX > Width / 2) And coAngle = 0 Then
            GdipAddPathLineI mPath, XX, Y, XX + coWidth, Y - coLen
            GdipAddPathLineI mPath, XX + coWidth, Y - coLen, XX + coWidth, Y
        Else
            If BCLT = 0 Then GdipAddPathLineI mPath, X, Y, X, Y
            GdipAddPathLineI mPath, XX, Y, XX + coAngle, Y - coLen
            GdipAddPathLineI mPath, XX + coAngle, Y - coLen, XX + coWidth, Y
        End If
    Else
        If BCLT = 0 Then GdipAddPathLineI mPath, X, Y, X + Width - BCRT, Y
    End If


    If BCRT Then GdipAddPathArcI mPath, X + Width - BCRT * 2, Y, BCRT * 2, BCRT * 2, 270, 90

    If m_CallOutPosicion = coRight And m_CallOut Then
        Select Case m_CallOutAlign
            Case coFirstCorner: YY = Y + BCRT
            Case coMidle: YY = Y + BCRT + ((Height - BCRT - BCBR) \ 2) - (coWidth \ 2)
            Case coSecondCorner: YY = Y + Height - coWidth - BCBR
            Case coCustomPosition: YY = Y + (m_coCustomPos * nScale)
        End Select
        XX = X + Width
        If (YY > Height / 2) And coAngle = 0 Then
            GdipAddPathLineI mPath, XX, YY, XX + coLen, YY + coWidth
            GdipAddPathLineI mPath, XX + coLen, YY + coWidth, XX, YY + coWidth
            
        Else
            If BCRT = 0 Then GdipAddPathLineI mPath, X + Width, Y, X + Width, Y
            GdipAddPathLineI mPath, XX, YY, XX + coLen, YY + coAngle
            GdipAddPathLineI mPath, XX + coLen, YY + coAngle, XX, YY + coWidth
        End If
    Else
        If BCRT = 0 Then GdipAddPathLineI mPath, X + Width, Y, X + Width, Y + Height - BCBR
    End If

    If BCBR Then GdipAddPathArcI mPath, X + Width - BCBR * 2, Y + Height - BCBR * 2, BCBR * 2, BCBR * 2, 0, 90


    If m_CallOutPosicion = coBottom And m_CallOut Then
        Select Case m_CallOutAlign
            Case coFirstCorner: XX = X + BCBL
            Case coMidle: XX = X + BCBL + ((Width - BCBR - BCBL) \ 2) - (coWidth \ 2)
            Case coSecondCorner: XX = X + Width - coWidth - BCBR
            Case coCustomPosition: XX = X + (m_coCustomPos * nScale)
        End Select
        
        YY = Y + Height
        If (XX > Width / 2) And coAngle = 0 Then
            GdipAddPathLineI mPath, XX + coWidth, YY, XX + coWidth, YY + coLen
            GdipAddPathLineI mPath, XX + coWidth, YY + coLen, XX, YY
        Else
            If BCBR = 0 Then GdipAddPathLineI mPath, X + Width, Y + Height, X + Width, Y + Height
            GdipAddPathLineI mPath, XX + coWidth, YY, XX + coAngle, YY + coLen
            GdipAddPathLineI mPath, XX + coAngle, YY + coLen, XX, YY
        End If
    Else
        If BCBR = 0 Then GdipAddPathLineI mPath, X + Width, Y + Height, X + BCBL, Y + Height
    End If

    If BCBL Then GdipAddPathArcI mPath, X, Y + Height - BCBL * 2, BCBL * 2, BCBL * 2, 90, 90
    
    If m_CallOutPosicion = coLeft And m_CallOut Then
        Select Case m_CallOutAlign
            Case coFirstCorner: YY = Y + BCLT
            Case coMidle: YY = Y + BCLT + ((Height - BCBL - BCLT) \ 2) - (coWidth \ 2)
            Case coSecondCorner: YY = Y + Height - coWidth - BCBL
            Case coCustomPosition: YY = Y + (m_coCustomPos * nScale)
        End Select
        
        If (YY > Height / 2) And coAngle = 0 Then
            GdipAddPathLineI mPath, X, YY + coWidth, X - coLen, YY + coWidth
            GdipAddPathLineI mPath, X - coLen, YY + coWidth, X, YY
        Else
            If BCBL = 0 Then GdipAddPathLineI mPath, X, Y + Height, X, Y + Height
            GdipAddPathLineI mPath, X, YY + coWidth, X - coLen, YY + coAngle
            GdipAddPathLineI mPath, X - coLen, YY + coAngle, X, YY
        End If
    Else
        If BCBL = 0 Then GdipAddPathLineI mPath, X, Y + Height, X, Y + BCLT
    End If

    GdipClosePathFigures mPath
  
    RoundRectangle = mPath

End Function

Private Function LoadImageFromStream(ByRef bvData() As Byte, ByRef hImage As Long) As Boolean
    On Local Error GoTo LoadImageFromStream_Error
    
    Dim IStream     As IUnknown
    If Not IsArrayDim(VarPtrArray(bvData)) Then Exit Function
    
    Call CreateStreamOnHGlobal(bvData(0), 0&, IStream)
    If Not IStream Is Nothing Then
        If GdipLoadImageFromStream(IStream, hImage) = 0 Then
            LoadImageFromStream = True
        End If
    End If

    Set IStream = Nothing
    
LoadImageFromStream_Error:
End Function

Private Function IsArrayDim(ByVal lpArray As Long) As Boolean
    Dim lAddress As Long
    Call CopyMemory(lAddress, ByVal lpArray, &H4)
    IsArrayDim = Not (lAddress = 0)
End Function

Public Function PictureFromURL(ByVal sUrl As String, Optional ByVal UseCache As Boolean, Optional ByVal DrawProgress As Boolean = True) As Boolean
    On Error Resume Next
    m_DrawProgress = DrawProgress
    UserControl.CancelAsyncRead "URL"
    Err.Clear
    Call AsyncRead(sUrl, vbAsyncTypeByteArray, "URL", IIf(UseCache, 0, vbAsyncReadForceUpdate))
    PictureFromURL = Err.Number = 0
End Function

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
    On Error Resume Next
    If m_DrawProgress Then
        If c_AsyncProp Is Nothing Then Set c_AsyncProp = AsyncProp
        Refresh
    End If
    RaiseEvent PictureDownloadProgress(AsyncProp.BytesMax, AsyncProp.BytesRead)
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    On Error GoTo PropErr
    
    If PictureFromStream(AsyncProp.Value) Then
        RaiseEvent PictureDownloadComplete
    Else
        RaiseEvent PictureDownloadError
    End If
    
    Set c_AsyncProp = Nothing
    Exit Sub
PropErr:
    RaiseEvent PictureDownloadError
End Sub


Public Function PictureFromStream(ByRef bvStream() As Byte) As Boolean
    Dim hImage As Long

    If LoadImageFromStream(bvStream, hImage) Then
        GdipGetImageWidth hImage, m_PictureRealWidth
        GdipGetImageHeight hImage, m_PictureRealHeight
        GdipDisposeImage hImage
        m_PictureArr() = bvStream
        PictureFromStream = True
        m_PicturePresent = True
    Else
        Erase m_PictureArr
        m_PicturePresent = False
    End If
    If m_PictureBrush <> 0 Then GdipDeleteBrush m_PictureBrush: m_PictureBrush = 0
    Call PropertyChanged("PicturePresent")
    Call PropertyChanged("PictureArr")
    Refresh
End Function

Public Sub tmrMOUSEOVER_Timer()
    Dim PT As POINTAPI
    Dim Left As Long, Top As Long
    Dim Rect As Rect
  
    GetCursorPos PT
    ScreenToClient c_lhWnd, PT
    
    Left = ScaleX(Extender.Left, vbContainerSize, UserControl.ScaleMode)
    Top = ScaleY(Extender.Top, vbContainerSize, UserControl.ScaleMode)

    With Rect
        .Left = m_PT.X - (m_Left - Left)
        .Top = m_PT.Y - (m_Top - Top)
        .Right = .Left + UserControl.ScaleWidth
        .Bottom = .Top + UserControl.ScaleHeight
    End With
    
    bIntercept = False
    SendMessage c_lhWnd, WM_MOUSEMOVE, 0&, ByVal PT.X Or PT.Y * &H10000
    
    If bIntercept = False Then
        If m_Over = True Then
            m_Over = False
            RaiseEvent MouseOut
        End If
    End If
    
    
    If PtInRect(Rect, PT.X, PT.Y) = 0 Then
        'WriteValue &H10, 0
        m_Enter = False
        tmrMOUSEOVER.Interval = 0
        RaiseEvent MouseLeave
    End If
    
End Sub

Public Function DrawLine(ByVal HDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, Optional ByVal oColor As OLE_COLOR = vbBlack, Optional ByVal Opacity As Integer = 100, Optional ByVal PenWidth As Integer = 1) As Boolean
    Dim hGraphics As Long, hPen As Long
    
    GdipCreateFromHDC HDC, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    GdipCreatePen1 ConvertColor(oColor, Opacity), PenWidth * nScale, UnitPixel, hPen
    DrawLine = GdipDrawLineI(hGraphics, hPen, X1 * nScale, Y1 * nScale, X2 * nScale, Y2 * nScale) = 0
    GdipDeletePen hPen
    GdipDeleteGraphics hGraphics
End Function


Public Function Polygon(ByVal HDC As Long, ByVal PenWidth As Long, ByVal oColor As OLE_COLOR, ByVal Opacity As Integer, ParamArray vPoints() As Variant) As Boolean
    Dim hGraphics As Long, hBrush As Long, hPen As Long
    Dim lPoints() As Long
    Dim lCount As Long
    Dim i As Long
    
    If UBound(vPoints) = 1 Then
        lCount = vPoints(1)
        ReDim lPoints(lCount - 1)
        CopyMemory lPoints(0), ByVal CLng(vPoints(0)), lCount * 4
    Else
        lCount = UBound(vPoints) + 1
        ReDim lPoints(lCount - 1)
        For i = 0 To lCount - 1
            lPoints(i) = vPoints(i) * nScale
        Next
    End If
    GdipCreateFromHDC HDC, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    
    If PenWidth > 0 Then
        GdipCreatePen1 ConvertColor(oColor, Opacity), PenWidth, UnitPixel, hPen
        Call GdipDrawPolygonI(hGraphics, hPen, lPoints(0), lCount / 2)
        GdipDeletePen hPen
    Else
        GdipCreateSolidFill ConvertColor(oColor, Opacity), hBrush
        Call GdipFillPolygonI(hGraphics, hBrush, lPoints(0), lCount / 2, &H1)
        GdipDeleteBrush hBrush
    End If
    
    GdipDeleteGraphics hGraphics
End Function

Public Property Get IsMouseOver() As Boolean
    IsMouseOver = m_Over
End Property

Public Property Get IsMouseEnter() As Boolean
    IsMouseEnter = m_Enter
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If hCur Then SetCursor hCur
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Debug.Print X, Y
    
    If hCur Then SetCursor hCur
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Property Let MousePointer(ByVal NewValue As MousePointerConstants)
    UserControl.MousePointer = NewValue
    PropertyChanged "MousePointer"
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Set MouseIcon(ByVal NewValue As IPictureDisp)
    UserControl.MouseIcon = NewValue
    PropertyChanged "MouseIcon"
End Property

Public Property Get MouseIcon() As IPictureDisp
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Let MousePointerHands(ByVal NewValue As Boolean)
    m_MousePointerHands = NewValue
    If NewValue Then
        If Ambient.UserMode Then
            UserControl.MousePointer = vbCustom
            UserControl.MouseIcon = GetSystemHandCursor
        End If
    Else
        If hCur Then DestroyCursor hCur: hCur = 0
        UserControl.MousePointer = vbDefault
        UserControl.MouseIcon = Nothing
    End If
    PropertyChanged "MousePointerHands"
End Property

Public Property Get MousePointerHands() As Boolean
    MousePointerHands = m_MousePointerHands
End Property

Public Property Get Image() As IPicture
    Dim DC As Long, TempHdc As Long
    Dim hBmp As Long, OldhBmp As Long
    Dim Pic As PicBmp, IID_IDispatch As GUID
    Dim lColor As Long, hBrush As Long
    Dim Rect As Rect
    
    DC = GetDC(0)
    TempHdc = CreateCompatibleDC(0)
    hBmp = CreateCompatibleBitmap(DC, UserControl.ScaleWidth, UserControl.ScaleHeight)
    ReleaseDC 0&, DC
    OldhBmp = SelectObject(TempHdc, hBmp)
    
    OleTranslateColor m_BackColor, 0, VarPtr(lColor)
    Rect.Right = UserControl.ScaleWidth
    Rect.Bottom = UserControl.ScaleHeight
    hBrush = CreateSolidBrush(lColor)
    FillRect TempHdc, Rect, hBrush
    DeleteObject hBrush
    
    'BitBlt TempHdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.HDC, 0, 0, vbSrcCopy
    Draw TempHdc, 0, 0, 0

    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
  
    With Pic
        .Size = Len(Pic)
        .type = vbPicTypeBitmap
        .hBmp = hBmp
        .hPal = 0
    End With

    If OldhBmp Then Call SelectObject(hDCMemory, OldhBmp): OldhBmp = 0
    If TempHdc Then DeleteDC TempHdc: TempHdc = 0

    Call OleCreatePictureIndirect(Pic, IID_IDispatch, 1, Image)

End Property

Public Function GetSystemHandCursor() As Picture
    Dim Pic As PicBmp, ipic As IPicture, GUID(0 To 3) As Long
    
    If hCur Then DestroyCursor hCur: hCur = 0
    
    hCur = LoadCursor(ByVal 0&, IDC_HAND)
     
    GUID(0) = &H7BF80980
    GUID(1) = &H101ABF32
    GUID(2) = &HAA00BB8B
    GUID(3) = &HAB0C3000
 
    With Pic
        .Size = Len(Pic)
        .type = vbPicTypeIcon
        .hBmp = hCur
        .hPal = 0
    End With
 
    Call OleCreatePictureIndirect(Pic, GUID(0), 1, ipic)
 
    Set GetSystemHandCursor = ipic
    
End Function

Public Property Get IconCharCode() As String
    IconCharCode = "&H" & Hex(m_IconCharCode)
End Property

Public Property Let IconCharCode(ByVal New_IconCharCode As String)
    New_IconCharCode = UCase(Replace(New_IconCharCode, Space(1), vbNullString))
    New_IconCharCode = UCase(Replace(New_IconCharCode, "U+", "&H"))
    If Not Left(New_IconCharCode, 2) = "&H" And Not IsNumeric(New_IconCharCode) Then
        m_IconCharCode = "&H" & New_IconCharCode
    Else
        m_IconCharCode = New_IconCharCode
    End If
    PropertyChanged "IconCharCode"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get IconForeColor() As OLE_COLOR
    IconForeColor = m_IconForeColor
End Property

Public Property Let IconForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_IconForeColor = New_ForeColor
    PropertyChanged "IconForeColor"
    Refresh
End Property

Public Property Get IconAlignmentH() As CaptionAlignmentH
    IconAlignmentH = m_IconAlignmentH
End Property

Public Property Let IconAlignmentH(ByVal New_IconAlignmentH As CaptionAlignmentH)
    m_IconAlignmentH = New_IconAlignmentH
    PropertyChanged "IconAlignmentH"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get IconAlignmentV() As CaptionAlignmentV
    IconAlignmentV = m_IconAlignmentV
End Property

Public Property Let IconAlignmentV(ByVal New_IconAlignmentV As CaptionAlignmentV)
    m_IconAlignmentV = New_IconAlignmentV
    PropertyChanged "IconAlignmentV"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get IconOpacity() As Integer
    IconOpacity = m_IconOpacity
End Property

Public Property Let IconOpacity(ByVal New_IconOpacity As Integer)
    m_IconOpacity = New_IconOpacity
    SafeRange m_IconOpacity, 0, 100
    PropertyChanged "IconOpacity"
    Refresh
End Property

Public Property Get IconPaddingY() As Integer
    IconPaddingY = m_IconPaddingY
End Property

Public Property Let IconPaddingY(ByVal New_IconPaddingY As Integer)
    m_IconPaddingY = New_IconPaddingY
    PropertyChanged "IconPaddingY"
    bRecreateShadowCaption = True
    Refresh
End Property

Public Property Get IconPaddingX() As Integer
    IconPaddingX = m_IconPaddingX
End Property

Public Property Let IconPaddingX(ByVal New_IconPaddingX As Integer)
    m_IconPaddingX = New_IconPaddingX
    PropertyChanged "IconPaddingX"
    bRecreateShadowCaption = True
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
    bRecreateShadowCaption = True
    Refresh
End Property

Private Sub SafeRange(Value, Min, Max)
    If Value < Min Then Value = Min
    If Value > Max Then Value = Max
End Sub

Private Sub CreateBuffer() 'Acrylic buffer
    Dim DC As Long
    If OldhBmp Then DeleteObject SelectObject(hDCMemory, OldhBmp): OldhBmp = 0
    If hDCMemory Then DeleteDC hDCMemory: hDCMemory = 0
 
    DC = GetDC(0)
    hDCMemory = CreateCompatibleDC(0)
    hBmp = CreateCompatibleBitmap(DC, UserControl.ScaleWidth, UserControl.ScaleHeight)
    ReleaseDC 0&, DC
    OldhBmp = SelectObject(hDCMemory, hBmp)
End Sub

'las dos funciones a continuacion son de cobein y con algunas modificaciones mias,
'las he utilizado para crear una bandera publica sin tener que agregar un modulo publico.
Private Function WriteValue(ByVal lProp As Long, ByVal lValue As Long) As Boolean
    Dim lFlagIndex As Long
    Dim i       As Long
    Dim lIndex  As Long: lIndex = -1
    
    For i = 0 To TLS_MINIMUM_AVAILABLE - 1
        If TlsGetValue(i) = lProp Then
            lIndex = i + 1
            Exit For
        End If
    Next

    If lIndex = -1 Then
        Do
            lFlagIndex = TlsAlloc '// Find two consecutive slots
            lIndex = TlsAlloc
            If lIndex >= TLS_MINIMUM_AVAILABLE Then Exit Function
        Loop While Not lFlagIndex + 1 = lIndex
        Call TlsSetValue(lFlagIndex, lProp)
        Call TlsSetValue(lIndex, lValue)
        WriteValue = True
    Else
        Call TlsSetValue(lIndex, lValue)
    End If
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

Public Function ChrW2(ByVal CharCode As Long) As String
  Const POW10 As Long = 2 ^ 10
  If CharCode <= &HFFFF& Then ChrW2 = ChrW$(CharCode) Else _
                              ChrW2 = ChrW$(&HD800& + (CharCode And &HFFFF&) \ POW10) & _
                                      ChrW$(&HDC00& + (CharCode And (POW10 - 1)))
End Function

Private Function ObjFromPtr(ByVal pObj As Long) As Object
    Dim obj As Object
    ' force the value of the pointer into the temporary object variable
    CopyMemory obj, pObj, 4
    ' assign to the result (this increments the ref counter)
    Set ObjFromPtr = obj
    ' manually destroy the temporary object variable
    ' (if you omit this step you'll get a GPF!)
    CopyMemory obj, 0&, 4
End Function


