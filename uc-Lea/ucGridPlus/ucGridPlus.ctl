VERSION 5.00
Begin VB.UserControl ucGridPlus 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ucGridPlus.ctx":0000
   Begin VB.Timer Timer1 
      Left            =   2520
      Top             =   1920
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin Proyecto1.ucScrollbar ucScrollbarH 
      Height          =   210
      Left            =   840
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   370
      SmallChange     =   3
      LargeChange     =   1
      Orientation     =   1
      DisableMouseWheelSupport=   -1  'True
      SmoothScrollFactor=   1
      WheelChange     =   3
      BeginProperty ThumbTooltipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Thumbsize_min   =   8
   End
   Begin Proyecto1.ucScrollbar ucScrollbarV 
      Height          =   2535
      Left            =   3720
      Top             =   720
      Visible         =   0   'False
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   4471
      LargeChange     =   10
      SmoothScrollFactor=   1
      WheelChange     =   27
      BeginProperty ThumbTooltipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Thumbsize_min   =   8
   End
End
Attribute VB_Name = "ucGridPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'--------------------------
'Autor: Leandro Ascierto
'Web: www.leandroascierto.com
'Date: 26/09/2021
'Version:   1.0.6
'Base on LynxGrid, IGrid, VBFlexGrid
'--------------------------
'29/09/21 V:1.0.1
    'Fixed Column sort, only left button, no sort when draw column.
    'Fixed when a row was resize the scrollbar did not show last rows.
    'Fixed error column sort when not rows
'01/10/21 V:1.0.2 'Thank you Elihu
    'Scroll Bugs if value > max on resize, internal scroll add direct call becaus,vbTimer is freeze
    'Selection with right button, and no deselect if cursor is in selection range + Rigth click
    'added AllowRowsResize property
    'added parameters 'Button', 'Shift' in the events (CellClick,LabelMouseDown,LabelMouseUp,ImgMouseDown,ImgMouseUp)
    'added Property BorderWidth and remove BorderVisible, fixel BorderBolor
    'added Property ShowHotColumn (highlight column under cursor)
'11/10/21 V:1.0.3
    'Added a lot of changes and fix, I can't remember which ones anymore
    'Added all events and property referring to Drag And Drop
    'Improvements in reading bmp and Icon images with 32-bit alpha channel
    'improved the function AutoWidthColumn
    'Added ColSort and ColSortOrder property to be able to sort by sql
'15/11/21 V:1.0.4
    'changes made by jpbro (vbforum), behavior in the text box when moving the keyboard arrows,
    'and added BeforeEdit event (Thanks)
'23/12/2021
    'fixed if the fixedrow row has a greater height than the rest of the rows, when clicking on a cell, the one in the next row is highlighted
'----------------------
'30/12/2021 V:1.0.5
    'The CtrlEdit property was added, this returns the textbox control of the grid,
    'with this you can manipulate the textbox at will and intercept all the events
'29/01/2022  V:1.0.6
    'Added Propertys ShowHotRow, HeaderTextWordBreak, GradientStyle, GetTopRow and GetVisibleRows

'Pending
'ColDelete
'HeaderMenu
'Selection for drag


#Const OCX_VERSION = False

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub CreateStreamOnHGlobal Lib "Ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long
Private Declare Function PathIsURL Lib "shlwapi.dll" Alias "PathIsURLA" (ByVal pszPath As String) As Long
Private Declare Function CryptStringToBinaryA Lib "crypt32.dll" (ByVal pszString As String, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByVal pcbBinary As Long, ByVal pdwSkip As Long, ByVal pdwFlags As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextW" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As Rect, ByVal wFormat As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "GDI32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "GDI32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "GDI32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetObjectType Lib "gdi32.dll" (ByVal hgdiobj As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, ByRef Brush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal Brush As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipSetPenMode Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mPenMode As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipFillRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipDrawRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipClosePathFigure Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "GdiPlus.dll" (ByRef mFilename As Long, ByRef mImage As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imageattr As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef imageattr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imageattr As Long, ByVal ColorAdjust As Long, ByVal EnableFlag As Boolean, ByRef MatrixColor As COLORMATRIX, ByRef MatrixGray As COLORMATRIX, ByVal Flags As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstWidth As Long, ByVal DstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipGetImageHeight Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mHeight As Long) As Long
Private Declare Function GdipGetImageWidth Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mWidth As Long) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "GdiPlus.dll" (ByVal mHicon As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GdiPlus.dll" (ByVal mHbm As Long, ByVal mhPal As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As IUnknown, ByRef Image As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "GdiPlus.dll" (ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStride As Long, ByVal mPixelFormat As Long, ByVal mScan0 As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipImageRotateFlip Lib "GdiPlus.dll" (ByVal mImage As Long, ByVal mRfType As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipSetClipPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPath As Long, ByVal mCombineMode As Long) As Long
Private Declare Function GdipResetClip Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
Private Declare Function PtInRect Lib "user32.dll" (ByRef lpRect As Rect, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ExtCreatePen Lib "gdi32.dll" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, ByRef lplb As LOGBRUSH, ByVal dwStyleCount As Long, ByRef lpStyle As Long) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetTextColor Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function GetObject Lib "GDI32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long
Private Declare Function CopyImage Lib "user32.dll" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function GradientFill Lib "msimg32" (ByVal hdc As Long, ByRef Vertex As TRIVERTEX, ByVal nVertex As Long, ByRef Mesh As GRADIENT_RECT, ByVal nMesh As Long, ByVal Mode As Long) As Long

'/>Jose Liza - FocusRect
Private Declare Function CreatePen Lib "GDI32" (ByVal nPenStyle As PEN_STYLE, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As Rect) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hdc As Long, lpRect As Rect, ByVal hBrush As Long) As Long

Private Enum PEN_STYLE
    PS_SOLID = 0
    PS_DASH = 1
    PS_DOT = 2
    PS_DASHDOT = 3
    PS_DASHDOTDOT = 4
    PS_NULL = 5
    PS_INSIDEFRAME = 6
End Enum
'/<Jose Liza - FocusRect

Private Type TRIVERTEX
    PxX As Long
    PxY As Long
    RedLow As Byte
    Red As Byte
    GreenLow As Byte
    Green As Byte
    BlueLow As Byte
    Blue As Byte
    AlphaLow As Byte
    Alpha As Byte
End Type

Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Private Type ICONINFO
    fIcon           As Long
    xHotspot        As Long
    yHotspot        As Long
    hbmMask         As Long
    hbmColor        As Long
End Type

Private Type BITMAP
  bmType                    As Long
  bmWidth                   As Long
  bmHeight                  As Long
  bmWidthBytes              As Long
  bmPlanes                  As Integer
  bmBitsPixel               As Integer
  bmBits                    As Long
End Type

Private Type GDIPlusStartupInput
    GdiPlusVersion                      As Long
    DebugEventCallback                  As Long
    SuppressBackgroundThread            As Long
    SuppressExternalCodecs              As Long
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

Private Type COLORMATRIX
    M(0 To 4, 0 To 4)           As Single
End Type

Private Const SmoothingModeAntiAlias    As Long = 4
Private Const UnitPixel                 As Long = 2&
Private Const PixelFormat32bppPARGB     As Long = &HE200B
Private Const PixelFormat32bppARGB      As Long = &H26200A
Private Const CombineModeIntersect      As Long = &H1
Private Const PenAlignmentInset         As Long = &H1
Private Const RotateNoneFlipY           As Long = &H6

Private Type tColumn
    Text As String
    Width As Long
    MinWidth As Long
    Align As eGridAlign
    DataType As eDataType
    TempWidth As Long
    BackColor As Long
    ForeColor As Long
    HeaderForeColor As Long
    Font As StdFont
    WordBreak As Boolean
    nSortOrder As lgSortTypeEnum
    SizeLocked As Boolean
    EditionLocked As Boolean
    ImgListWidth As Long
    ImgListHeight  As Long
    ColImgList As Collection
    ImgAlign As eGridAlign
    ImagesRadius As Integer
    ImagesMonocrome As Boolean
    LabelsEvents As Boolean
    ImagesEvents As Boolean
    TextHide As Boolean
    Format As String
    IconIndex As Long
    Tag As Variant
End Type

Private Type tCell
    Value As Variant
    Tag As Variant
    BackColor As Long
    ForeColor As Long
    Font As StdFont
    Align As eGridAlign
    WordBreak As Boolean
    IconIndex As Integer
    EditionLocked As Boolean
End Type

Private Type tRow
    Height As Long
    TempHeight As Long
    Cells() As tCell
    BackColor As Long
    ForeColor As Long
    Font As StdFont
    Align As eGridAlign
    WordBreak As Boolean
    IsGroup As Boolean
    IsFullRow As Boolean
    IsGroupExpanded As Boolean
    Ident As Byte
    Checked As Boolean
    Tag As Variant
    '/> Jose Liza - ParentGroup
    RowParent As Long
    '/< Jose Liza - ParentGroup
End Type

Private Type tRowColResize
    HotColumn As Long
    CurColumn As Long
    HotRow As Long
    CurRow As Long
    Left As Long
End Type

Private Type tColDrag
    SrcCol  As Long
    DestCol As Long
    X As Long
    Left As Long
End Type

Private Type tSelectionRange
    Start As POINTAPI
    End As POINTAPI
End Type

Private Const DT_LEFT As Long = &H0
Private Const DT_CENTER As Long = &H1
Private Const DT_RIGHT As Long = &H2
Private Const DT_TOP As Long = &H0
Private Const DT_VCENTER        As Long = &H4
Private Const DT_BOTTOM As Long = &H8
Private Const DT_WORDBREAK As Long = &H10
Private Const DT_SINGLELINE     As Long = &H20
Private Const DT_WORD_ELLIPSIS  As Long = &H40000
Private Const DT_NOPREFIX As Long = &H800
Private Const DT_EXPANDTABS As Long = &H40
Private Const DT_CALCRECT As Long = &H400

Private Const IMAGE_BITMAP As Long = 0
Private Const OBJ_BITMAP As Long = 7
Private Const LR_CREATEDIBSECTION As Long = &H2000
                    

Private Const CLR_NONE          As Long = &HFFFFFFFF
Private Const LOGPIXELSX        As Long = 88
Private Const IDC_HAND          As Long = 32649

Private Const PS_COSMETIC As Long = &H0
Private Const PS_ALTERNATE As Long = 8

Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Private Const C_NULL_RESULT As Long = -1

Public Enum lgSortTypeEnum
   lgSTAscending = 1
   lgSTDescending = 2
   lgSTNormal = 0
End Enum

Private Enum enuIdIconFont
    IIF_SortAsc
    IIF_SortDes
    IIF_Edit
    IIF_TreeExpanded
    IIF_TreeColapsed
    IIF_GroupExpanded
    IIF_GroupColapsed
    IIF_Asterisk
    IIF_ArrowRight
End Enum

Public Enum eDataType
    GP_STRING
    GP_NUMERIC
    GP_CURRENCY
    GP_DATE
    GP_BOOLEAN
    GP_CUSTOM
End Enum

Public Enum eGridAlign
    TopLeft = DT_TOP Or DT_LEFT
    TopCenter = DT_TOP Or DT_CENTER
    TopRight = DT_TOP Or DT_RIGHT
    CenterLeft = DT_VCENTER Or DT_LEFT
    CenterCenter = DT_VCENTER Or DT_CENTER
    CenterRight = DT_VCENTER Or DT_RIGHT
    BottomLeft = DT_BOTTOM Or DT_LEFT
    BottomCenter = DT_BOTTOM Or DT_CENTER
    BottomRight = DT_BOTTOM Or DT_RIGHT
End Enum

Public Enum eHeaderAlign
    HA_DefaultColumn = -1
    HA_Left = DT_LEFT
    HA_Center = DT_CENTER
    HA_Right = DT_RIGHT
End Enum

Public Enum eSelectionMode
    GP_SelFree
    GP_SelBySingleCell
    GP_SelByMultiRow
    GP_SelBySingleRow
    GP_SelByCol
End Enum

Private Enum EnuSelectBy
    SelectByNone = -1
    SelectByCells = 0
    SelectByColumns = 1
    SelectByRow = 2
    SelectBtCornerLeftTop = 3
End Enum

Public Enum GridPlusOLEDropModeConstants
    OLEDropModeNone = vbOLEDropNone
    OLEDropModeManual = vbOLEDropManual
End Enum

'/> Lizano Dias - Lynxgrid
Public Enum lgFocusRectModeEnum
   lgNone = 0
   lgRow = 1
   lgCol = 2
End Enum

Public Enum lgFocusRectStyleEnum
   lgFRLight = 0
   lgFRMedium = 1
   lgFRHeavy = 2
End Enum
'/< Lizano Dias - Lynxgrid


'EVENTOS
Public Event AfterEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal vOldValue As Variant)
Public Event BeforeEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal vValue As Variant, ByRef bCancel As Boolean)
Public Event BeforeSorted()
Public Event AfterSorted()
Public Event ColumnClick(ByVal lCol As Long, Button As Integer, Shift As Integer)
Public Event CurCellChange(ByVal lNewRowIfAny As Long, ByVal lNewColIfAny As Long, ByVal lOldRowIfAny As Long, ByVal lOldColIfAny As Long)
Public Event RowPostPaint(ByVal hdc As Long, ByVal Row As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
Public Event CellPostPaint(ByVal hdc As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
Public Event VerticalScroll()
Public Event HorizontalScroll()
Public Event EditKeyPress(KeyAscii As Integer)
Public Event EditKeyDown(KeyCode As Integer, Shift As Integer)
Public Event CellClick(ByVal lRow As Long, ByVal lCol As Long, Button As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event ImgMouseDown(ByVal lRow As Long, ByVal lCol As Long, Button As Integer, Shift As Integer)
Public Event ImgMouseUp(ByVal lRow As Long, ByVal lCol As Long, Button As Integer, Shift As Integer)
Public Event LabelMouseDown(ByVal lRow As Long, ByVal lCol As Long, Button As Integer, Shift As Integer)
Public Event LabelMouseUp(ByVal lRow As Long, ByVal lCol As Long, Button As Integer, Shift As Integer)
Public Event DblClick()
Public Event ColumnUserResize(ByVal lCol As Long)
Public Event RowUserResize(ByVal lRow As Long)
Public Event VerticalScrollEnd()
Public Event OnScrollPaint(ByVal Vertical As Boolean, ByVal lhDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal ePart As sbOnPaintPartCts, ByVal estate As sbOnPaintPartStateCts)
Public Event CustomSort(bAscending As Boolean, lSortColumn As Long, Value1 As Variant, Value2 As Variant, bSwap As Boolean)
'Public Event HorizontalScrollEnd()
Public Event HotCellChange(lRow As Long, lCol As Long)
Public Event MouseLeave()
Public Event MouseEnter()
Public Event BeforeColumnDrag(ByVal lCol As Long, Cancel As Boolean)
Public Event OnColumnDrag(ByVal SrcCol As Long, ByVal DestCol As Long, Cancel As Boolean)
Public Event AfterColumnDrag(ByVal lCol As Long)
Public Event OLECompleteDrag(Effect As Long)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, state As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)


Private mCol() As tColumn
Private mRow() As tRow
Private PtrCol() As Long
Private PtrRow() As Long

Dim m_GroupsTreeStyle As Boolean
Dim m_ColumnsAutoFit As Boolean
Dim m_BorderRadius As Long
Dim m_FixedColumns As Long
Dim m_FixedRows As Long
Dim mColDrag As tColDrag
Dim mSortColumn As Long
Dim mSortSubColumn As Long
Dim m_AllowColumnSort As Boolean
Dim m_AllowColumnDrag As Boolean
Dim m_AllowRowsResize As Boolean
Dim m_ShowHotColumn As Boolean
Dim m_AllowEdit As Boolean
Dim mCellEdit As POINTAPI
Dim m_Font As StdFont
Dim m_ParentBackColor As OLE_COLOR
Dim m_RowsBackColor As OLE_COLOR
Dim m_RowsBackColorAlt As OLE_COLOR
Dim m_BorderColor As OLE_COLOR
Dim m_ShowHotRow As Boolean
'Dim m_BorderVisible As Boolean
Dim m_Redraw As Boolean
Dim m_HeaderTextWordBreak As Boolean
Dim m_HeaderTextAlign As eHeaderAlign
Dim m_HeaderImageAlign As eHeaderAlign
Dim m_HeaderFont As Font
Dim m_ColsCount As Long
Dim m_RowsCount As Long
Dim m_ColsWidth As Long
Dim m_HeaderBackColor As Long
Dim m_HeaderHeight As Long
Dim m_RowsHeight As Long
Dim m_RowSelectorWidth As Long
Dim m_RowSelectorBkColor As Long
Dim m_SelectionMode As eSelectionMode
Dim m_SelectionColor As OLE_COLOR
Dim m_LinesHorizontalColor As OLE_COLOR
Dim m_LinesHorizontalWidth As Long
Dim m_LinesVerticalColor As OLE_COLOR
Dim m_LinesVerticalWidth As Long
'Dim m_HeaderLinesHorizontalColor As OLE_COLOR
Dim m_HeaderLinesHorizontalWidth As Long
'Dim m_HeaderLinesVerticalColor As OLE_COLOR
Dim m_HeaderLinesVerticalWidth As Long
Dim m_GradientStyle As Boolean
Dim SizeColumn As tRowColResize
Dim SelCel As POINTAPI
Dim SelRange As tSelectionRange
Dim Margin As Long
Dim DpiF As Single
Dim mHotCell As POINTAPI, HotRow As Long
Dim mHotCol As Long, HotCol As Long
Dim mHotPart As POINTAPI, HotPartIsImage As Boolean
Dim isSegoeFontInstaled As Boolean
Dim m_SegoeFont As StdFont
Dim m_Wingdings2 As StdFont
Dim m_Wingdings3 As StdFont
Dim m_BorderWidth As Long
Dim mHideRowsCount As Long
Dim VScrollPos As Long
Dim m_CheckStyle As Boolean
Dim m_AllRowAreCheked As Boolean
Dim hCurHands As Long
Dim GdipToken As Long
Dim cHeaderImageList As Collection
Dim m_HeaderImgLstWidth As Long
Dim m_HeaderImgLstHeight As Long
Dim m_LastRowIsFooter As Boolean
Dim mSelectAllOnFocus As Boolean
Dim eSelBy As EnuSelectBy

'/> Lizano Dias - Lynxgrid
Dim m_FocusRectColor            As Long
Dim m_FocusRectMode         As lgFocusRectModeEnum
Dim m_FocusRectStyle        As lgFocusRectStyleEnum
'/< Lizano Dias - Lynxgrid

'/>Jose Liza - FocusRect
Dim m_CellFocusRect As POINTAPI
Dim m_CellRectInfo  As Rect

Private Enum CellOrientationFocus
    All
    Horizontal
    Vertical
End Enum
'/<Jose Liza - FocusRect


'*1

'Esto puede no ser muy confiable dependiendo factores como fixedrow y gruprows etc.
Public Property Get GetTopRow() As Long
    GetTopRow = VScrollPos
End Property

Public Property Get GetVisibleRows() As Long
    GetVisibleRows = GetNroRowsInScreen
End Property

Public Property Get RunMode() As Boolean
On Error Resume Next
    RunMode = True
    RunMode = Ambient.UserMode
    RunMode = Extender.Parent.RunMode
End Property

Private Function IsPointInSelRange(X As Long, Y As Long) As Boolean
    If Y >= SelRange.Start.Y And Y <= SelRange.End.Y And _
        X >= SelRange.Start.X And X <= SelRange.End.X Then
        IsPointInSelRange = True
    End If
End Function

Private Function GetCellBackColor(Row As Long, Col As Long, Index As Long) As Long
    If mRow(Row).Cells(Col).BackColor <> CLR_NONE Then
        GetCellBackColor = mRow(Row).Cells(Col).BackColor
    ElseIf mRow(Row).BackColor <> CLR_NONE Then
        GetCellBackColor = mRow(Row).BackColor
    ElseIf mCol(Col).BackColor <> CLR_NONE Then
        GetCellBackColor = mCol(Col).BackColor
    Else
        GetCellBackColor = IIF(Index Mod 2 <> 0, m_RowsBackColorAlt, m_RowsBackColor)
    End If
End Function

Private Function GetCellFont(Row As Long, Col As Long) As StdFont
    If Not mRow(Row).Cells(Col).Font Is Nothing Then
        Set GetCellFont = mRow(Row).Cells(Col).Font
    ElseIf Not mRow(Row).Font Is Nothing Then
        Set GetCellFont = mRow(Row).Font
    ElseIf Not mCol(Col).Font Is Nothing Then
        Set GetCellFont = mCol(Col).Font
    Else
        Set GetCellFont = m_Font
    End If
End Function

Private Function GetCellAlign(Row As Long, Col As Long) As eGridAlign
    If mRow(Row).Cells(Col).Align <> CenterLeft Then
        GetCellAlign = mRow(Row).Cells(Col).Align
    ElseIf mRow(Row).Align <> CenterLeft Then
        GetCellAlign = mRow(Row).Align
    ElseIf mCol(Col).Align <> CenterLeft Then
        GetCellAlign = mCol(Col).Align
    Else
        GetCellAlign = CenterLeft
    End If
End Function

Private Function GetCellWordBreak(Row As Long, Col As Long) As Long
    If mRow(Row).Cells(Col).WordBreak Then
        GetCellWordBreak = DT_WORDBREAK
    ElseIf mRow(Row).WordBreak Then
        GetCellWordBreak = DT_WORDBREAK
    ElseIf mCol(Col).WordBreak Then
        GetCellWordBreak = DT_WORDBREAK
    Else
        GetCellWordBreak = DT_SINGLELINE Or DT_WORD_ELLIPSIS
    End If
End Function

Private Function GetCellForeColor(Row As Long, Col As Long, BkColor As Long) As Long
    If mRow(Row).Cells(Col).ForeColor <> CLR_NONE Then
        GetCellForeColor = mRow(Row).Cells(Col).ForeColor
    ElseIf mRow(Row).ForeColor <> CLR_NONE Then
        GetCellForeColor = mRow(Row).ForeColor
    ElseIf mCol(Col).ForeColor <> CLR_NONE Then
        GetCellForeColor = mCol(Col).ForeColor
    Else
        GetCellForeColor = IIF(IsDarkColor(BkColor), vbWhite, vbBlack)
    End If
End Function

Private Function GetColumnsWidth() As Long
    Dim Col As Long
    GetColumnsWidth = m_RowSelectorWidth
    For Col = 0 To m_ColsCount - 1
        If mCol(Col).TempWidth = 0 Then
            GetColumnsWidth = GetColumnsWidth + mCol(Col).Width
        End If
    Next
End Function

Private Sub DrawTreeViewLines(hdc As Long, Row As Long, pRow As Long, R As Rect)
    Dim R2 As Rect, i As Long, J As Long
    Dim bDraw As Boolean

    For i = 0 To mRow(pRow).Ident - 2
        bDraw = True
        For J = Row To m_RowsCount - 1
            If mRow(PtrRow(J)).IsGroup Then
                  If mRow(PtrRow(J)).Ident = i Then
                    bDraw = False
                    Exit For
                 ElseIf mRow(PtrRow(J)).Ident = i + 1 Then
                    Exit For
                End If
            End If
        Next
        If J = m_RowsCount Then bDraw = False
        
        If bDraw Then
            With R2
                .Left = R.Left + Margin + (20 * DpiF * (i)) + 8 * DpiF
                .Top = R.Top
                .Right = .Left
                .Bottom = R.Bottom
                If i = mRow(pRow).Ident - 2 Then
                    If Row = m_RowsCount - 1 Then
                        .Bottom = R.Bottom - (R.Bottom - R.Top) / 2
                    End If
                End If

                DrawLine hdc, .Left, .Top, .Right, .Bottom, PS_COSMETIC Or PS_ALTERNATE, ForeColor, 1
            End With
        End If
    Next

    If mRow(pRow).Ident > 0 Then
        With R2
            .Left = R.Left + Margin + (20 * DpiF * (mRow(pRow).Ident - 1)) + 8 * DpiF
            If Not mRow(pRow).IsGroup Then .Left = .Left - 20 * DpiF
            .Top = R.Bottom - (R.Bottom - R.Top) / 2
            .Right = .Left + 16 * DpiF
            .Bottom = .Top
            DrawLine hdc, .Left, .Top, .Right, .Bottom, PS_COSMETIC Or PS_ALTERNATE, ForeColor, 1
        End With
    End If
    
    With R2
        .Top = R.Top
        .Right = .Left
        If Row < m_RowsCount - 1 Then
            If mRow(pRow).IsGroup Then
                For J = Row + 1 To m_RowsCount - 1
                    If mRow(PtrRow(J)).IsGroup Then
                        If mRow(PtrRow(J)).Ident = mRow(pRow).Ident Then
                            .Bottom = R.Bottom
                            Exit For
                        ElseIf mRow(PtrRow(J)).Ident < mRow(pRow).Ident Then
                            .Bottom = R.Bottom - (R.Bottom - R.Top) / 2
                            Exit For
                        End If
                    End If
                Next
            Else
                If mRow(PtrRow(Row + 1)).IsGroup Then
                    .Bottom = R.Bottom - (R.Bottom - R.Top) / 2
                Else
                    .Bottom = R.Bottom
                End If
            End If
        End If
        DrawLine hdc, .Left, .Top, .Right, .Bottom, PS_COSMETIC Or PS_ALTERNATE, ForeColor, 1
    End With
    
    If mRow(pRow).IsGroup And mRow(pRow).IsGroupExpanded Then
        With R2
            .Left = R.Left + Margin + (20 * DpiF * (mRow(pRow).Ident)) + 8 * DpiF
            .Top = R.Bottom - (R.Bottom - R.Top) / 2 + 4 * DpiF
            .Right = .Left
            .Bottom = R.Bottom
            DrawLine hdc, .Left, .Top, .Right, .Bottom, PS_COSMETIC Or PS_ALTERNATE, ForeColor, 1
        End With
    End If
End Sub

Private Sub GetCellRectsParts(ByVal Row As Long, ByVal Col As Long, _
                    FirstCol As Boolean, Flags As Long, Align As Long, _
                    RCell As Rect, RCheckRow As Rect, RText As Rect, RImage As Rect)
                    
    Dim CheckSize As Long
    Dim RLabel As Rect
    Dim TextHeight As Long, TextWidth As Long
    Dim ImgW As Long, ImgH As Long
    Dim CalculateImage As Boolean
    CheckSize = 16 * DpiF

    With RLabel
        .Left = RCell.Left + Margin
        .Top = RCell.Top + Margin / 2
        .Right = RCell.Right - Margin - m_LinesVerticalWidth * DpiF
        .Bottom = RCell.Bottom - Margin / 2 - m_LinesHorizontalWidth * DpiF
        
        If FirstCol And mRow(Row).Ident > 0 Then
            .Left = .Left + Margin + ((Margin + CheckSize) * (mRow(Row).Ident - 1))
        End If

        If m_CheckStyle And FirstCol Then

            RCheckRow.Left = .Left
            RCheckRow.Top = .Top + (.Bottom - .Top) / 2 - CheckSize / 2
            RCheckRow.Right = RCheckRow.Left + CheckSize
            RCheckRow.Bottom = RCheckRow.Top + CheckSize
            .Left = .Left + CheckSize + Margin
        End If
        
        If mCol(Col).ImgListWidth > 0 Then
            ImgW = mCol(Col).ImgListWidth
            ImgH = mCol(Col).ImgListHeight
            CalculateImage = True
        ElseIf mRow(Row).Cells(Col).IconIndex > 0 And m_HeaderImgLstWidth > 0 Then
            ImgW = m_HeaderImgLstWidth
            ImgH = m_HeaderImgLstHeight
            CalculateImage = True
        End If
        
        If CalculateImage Then
            Select Case mCol(Col).ImgAlign
                Case TopLeft
                    RImage.Left = .Left
                    RImage.Top = .Top
                    .Left = .Left + ImgW + Margin
                Case TopCenter
                    RImage.Left = .Left + (.Right - .Left) / 2 - ImgW / 2
                    RImage.Top = .Top
                Case TopRight
                    RImage.Left = .Right - ImgW
                    RImage.Top = .Top
                    .Right = .Right - ImgW - Margin
                Case CenterLeft
                    RImage.Left = .Left
                    RImage.Top = .Top + (.Bottom - .Top) / 2 - ImgH / 2
                    .Left = .Left + ImgW + Margin
                Case CenterCenter
                    RImage.Left = .Left + (.Right - .Left) / 2 - ImgW / 2
                    RImage.Top = .Top + (.Bottom - .Top) / 2 - ImgH / 2
                Case CenterRight
                    RImage.Left = .Right - ImgW
                    RImage.Top = .Top + (.Bottom - .Top) / 2 - ImgH / 2
                    .Right = .Right - ImgW - Margin
                Case BottomLeft
                    RImage.Left = .Left
                    RImage.Top = .Bottom - ImgH
                    .Left = .Left + ImgW + Margin
                Case BottomCenter
                    RImage.Left = .Left + (.Right - .Left) / 2 - ImgW / 2
                    RImage.Top = .Bottom - ImgH
                Case BottomRight
                    RImage.Left = .Right - ImgW
                    RImage.Top = .Bottom - ImgH
                    .Right = .Right - ImgW - Margin
            End Select
                RImage.Right = RImage.Left + ImgW
                RImage.Bottom = RImage.Top + ImgH
        End If

  
        RText = RLabel
        
        If mCol(Col).DataType = GP_BOOLEAN Then
            RText.Right = .Left + CheckSize
            RText.Bottom = .Top + CheckSize
        Else
            DrawText UserControl.hdc, StrPtr(GetCellText(Row, Col)), -1, RText, DT_CALCRECT Or Flags
        End If
        
        With RText
            TextHeight = .Bottom - .Top
            TextWidth = .Right - .Left
        
            Select Case Align
                Case TopLeft
                    .Left = RLabel.Left
                    .Top = RLabel.Top
                Case TopCenter
                    .Left = RLabel.Left + (RLabel.Right - RLabel.Left) / 2 - TextWidth / 2
                    .Top = RLabel.Top
                Case TopRight
                    .Left = RLabel.Right - TextWidth
                    .Top = RLabel.Top
                Case CenterLeft
                    .Left = RLabel.Left
                    .Top = RLabel.Top + (RLabel.Bottom - RLabel.Top) / 2 - TextHeight / 2
                Case CenterCenter
                    .Left = RLabel.Left + (RLabel.Right - RLabel.Left) / 2 - TextWidth / 2
                    .Top = RLabel.Top + (RLabel.Bottom - RLabel.Top) / 2 - TextHeight / 2
                Case CenterRight
                    .Left = RLabel.Right - TextWidth
                    .Top = RLabel.Top + (RLabel.Bottom - RLabel.Top) / 2 - TextHeight / 2
                Case BottomLeft
                    .Left = RLabel.Left
                    .Top = RLabel.Bottom - TextHeight
                Case BottomCenter
                    .Left = RLabel.Left + (RLabel.Right - RLabel.Left) / 2 - TextWidth / 2
                    .Top = RLabel.Bottom - TextHeight
                Case BottomRight
                    .Left = RLabel.Right - TextWidth
                    .Top = RLabel.Bottom - TextHeight
            End Select
            
            .Bottom = .Top + TextHeight
            .Right = .Left + TextWidth
            
            
            If .Top < RLabel.Top Then .Top = RLabel.Top
            If .Left < RLabel.Left Then .Left = RLabel.Left
            If .Right > RLabel.Right Then .Right = RLabel.Right
            If .Bottom > RLabel.Bottom Then .Bottom = RLabel.Bottom
            
            
        End With
        
        'If mCol(Col).DataType = GP_BOOLEAN Then RCheck = RText
        
        'FillRectangle UserControl.hdc, RCheckRow, vbRed
    End With
End Sub

'*1
Private Sub Draw()
    Dim FlagFooter As Boolean
    Dim X As Long, Y As Long, i As Long
    Dim R As Rect, R2 As Rect, R3 As Rect
    Dim lTop As Long
    Dim VLineColor As Long, HLineColors As Long
    Dim BkColor As Long
    Dim lForeColor As Long
    Dim oFont As iFont
    Dim Align As Long, Flags As Long
    Dim Row As Long, Col As Long
    Dim LVW As Long, LHW As Long
    Dim HLVW As Long, HLHW As Long
    Dim FixedColWidth As Long
    Dim FixedRowHeight As Long
    Dim PaintPart As Long
    Dim lStart1  As Long, lEnd1 As Long
    Dim lStart2  As Long, lEnd2 As Long
    Dim ColumsWidth As Long
    Dim IdIconFont As enuIdIconFont
    Dim IconColors As Long
    Dim RCheck As Rect, RText As Rect, RImage As Rect
    Dim hImage As Long
    Dim bDrawSelLine As Boolean
    
    If m_Redraw = False Then Exit Sub
    
    If m_FixedRows > m_RowsCount Then Exit Sub
    If m_FixedColumns > m_ColsCount Then Exit Sub
    
    LVW = m_LinesVerticalWidth * DpiF
    LHW = m_LinesHorizontalWidth * DpiF
    HLVW = m_HeaderLinesVerticalWidth * DpiF
    HLHW = m_HeaderLinesHorizontalWidth * DpiF
    
    If ColumsWidth = 0 Then ColumsWidth = GetColumnsWidth
    
    
    If m_HeaderFont Is Nothing Then Exit Sub
    
    If ucScrollbarH.Visible = False Or ucScrollbarV.Visible = False Or ucScrollbarH.Value = ucScrollbarH.Max Or ucScrollbarV.Value = ucScrollbarV.Max Then
        UserControl.Cls
    End If

    Set oFont = m_Font
    SelectObject UserControl.hdc, oFont.hFont

    For i = 0 To m_FixedColumns - 1
        FixedColWidth = FixedColWidth + mCol(PtrCol(i)).Width
    Next

    For i = 0 To m_FixedRows - 1
        FixedRowHeight = FixedRowHeight + mRow(PtrRow(i)).Height
    Next
    
    '/> Jose Liza
    If SelCel.Y = -1 And SelCel.X = -1 Then
        SelCel.Y = 0: SelCel.X = 0
        SelRange.Start = SelCel: SelRange.End = SelCel
    End If
    '/< Jose Liza

DrawFixed:
    
    If PaintPart = 0 Or PaintPart = 1 Then
        lStart1 = VScrollPos + m_FixedRows
        lEnd1 = m_RowsCount - 1
        lTop = m_HeaderHeight + FixedRowHeight
    ElseIf PaintPart = 2 Or PaintPart = 3 Then
        lStart1 = 0
        lEnd1 = m_FixedRows - 1
        lTop = m_HeaderHeight
    End If
    
    '/>Jose Liza - FocusRect
    If m_RowsCount > 0 And m_CellFocusRect.X = -1 And m_CellFocusRect.Y = -1 Then m_CellFocusRect.X = 0: m_CellFocusRect.Y = 0
    '/<Jose Liza - FocusRect

    'Draw Cells
    
    For Row = lStart1 To lEnd1
                
Draw_Footer:
        Y = PtrRow(Row)

        If PaintPart = 0 Or PaintPart = 2 Then
            lStart2 = m_FixedColumns
            lEnd2 = m_ColsCount - 1
            R.Right = -ucScrollbarH.Value + m_RowSelectorWidth + FixedColWidth
        ElseIf PaintPart = 1 Or PaintPart = 3 Then
            lStart2 = 0
            lEnd2 = m_FixedColumns - 1
            R.Right = m_RowSelectorWidth
        End If
'*1
        If mRow(Y).IsGroup Or mRow(Y).IsFullRow = True Then
        
            If mRow(Y).TempHeight = 0 Then

                SetRect R3, m_RowSelectorWidth - ucScrollbarH.Value, lTop, ColumsWidth - ucScrollbarH.Value, lTop + mRow(Y).Height
                
                BkColor = GetCellBackColor(Y, 0, Row)
                Set oFont = GetCellFont(Y, 0)
                Align = GetCellAlign(Y, 0)
                Flags = GetCellWordBreak(Y, 0) Or DT_NOPREFIX Or DT_EXPANDTABS
                
                
                If SelCel.Y = Row Then
                    BkColor = m_SelectionColor
                    VLineColor = ShiftColor(m_RowsBackColor, m_SelectionColor, 50)
                    HLineColors = VLineColor
                    bDrawSelLine = True
                ElseIf IsPointInSelRange(0, Row) And Text1.Visible = False Then
                    BkColor = ShiftColor(BkColor, m_SelectionColor, 150)
                    VLineColor = ShiftColor(m_LinesVerticalColor, m_SelectionColor, 120)
                    HLineColors = VLineColor
                    bDrawSelLine = True
                Else
                    VLineColor = m_LinesVerticalColor
                    HLineColors = m_LinesHorizontalColor
                End If
    
                If m_LinesVerticalWidth > 1 Then VLineColor = m_LinesVerticalColor
                If m_LinesHorizontalWidth > 1 Then HLineColors = m_LinesHorizontalColor
                
                lForeColor = GetCellForeColor(Y, 0, BkColor)
                SetTextColor hdc, lForeColor
                
                FillRectangle UserControl.hdc, R3, BkColor
                
                '/>Jose Liza - FocusRect
                If Me.FocusRectMode <> lgNone Then
                    If m_CellFocusRect.Y = Row Then
                        m_CellFocusRect.X = 0
                        With m_CellRectInfo
                            .Left = R3.Left + 1
                            .Top = R3.Top - IIF(Row > 0, 1, 0)
                            .Right = R3.Right
                            .Bottom = R3.Bottom
                        End With
                        '---
                    End If
                End If
                '/<Jose Liza - FocusRect

                Call SelectObject(UserControl.hdc, oFont.hFont)
                
                With R2
                  .Left = R3.Left + Margin + (20 * DpiF * mRow(Y).Ident)
                  .Top = R3.Top + Margin
                  .Right = R3.Right - Margin - LVW
                  .Bottom = R3.Bottom - Margin - LHW
                End With
                    
                If mRow(Y).IsGroup Then
                    R2.Left = R2.Left + Margin
                    
                    If m_GroupsTreeStyle Then
                        DrawTreeViewLines hdc, Row, Y, R3 '------------TreeLines
                        '---------------
                        
                        If mRow(Y).IsGroupExpanded Then
                            DrawIconFont hdc, IIF_TreeExpanded, m_Font.Size, False, R2, CenterLeft
                        Else
                            DrawIconFont hdc, IIF_TreeColapsed, m_Font.Size, False, R2, CenterLeft
                        End If
                    Else
                        If mRow(Y).IsGroupExpanded Then
                            DrawIconFont hdc, IIF_GroupExpanded, m_Font.Size, False, R2, CenterLeft
                        Else
                            DrawIconFont hdc, IIF_GroupColapsed, m_Font.Size, False, R2, CenterLeft
                        End If
                    End If
                    R2.Left = R2.Left + Margin + 16 * DpiF
                End If
                
                If m_CheckStyle Then
                    R2.Bottom = R2.Bottom + Margin / 2
                    R2.Top = R2.Top - Margin / 2
                    DrawCheckBox UserControl.hdc, mRow(Y).Checked, BkColor = m_SelectionColor, R2, CenterLeft ' Align
                    R2.Bottom = R2.Bottom - Margin / 2
                    R2.Top = R2.Top + Margin / 2
                    R2.Left = R2.Left + 20 * DpiF
                End If
                
                DrawText UserControl.hdc, StrPtr(GetCellText(Y, 0)), -1, R2, Flags Or Align 'Or DT_WORDBREAK
                
                
                If LVW > 0 Then DrawLine2 hdc, R3.Right - LVW, R3.Top, R3.Right, R3.Bottom, VLineColor
                If LHW > 0 Then DrawLine2 hdc, R3.Left, R3.Bottom - LHW, R3.Right, R3.Bottom, HLineColors
                
                If FlagFooter = True Then
                    If LHW > 0 Then DrawLine2 hdc, R3.Left, R3.Top - LHW * 2, R3.Right, R3.Top, HLineColors
                End If
                
                If bDrawSelLine Then
                    If LVW > 0 Then DrawLine2 hdc, R3.Left - LVW, R3.Top, R3.Left, R3.Bottom, VLineColor
                    If LHW > 0 Then DrawLine2 hdc, R3.Left - LVW, R3.Top - LHW, R3.Right, R3.Top, HLineColors
                End If
                
                If (PaintPart = 2 Or PaintPart = 3) And Row = m_FixedRows - 1 Then
                    DrawLine2 hdc, R.Left, R.Bottom - LHW * 2, R.Right, R.Bottom, HLineColors
                End If
                
            End If
        Else '===================DRAW CELLS=================================================

            For Col = lStart2 To lEnd2
                X = PtrCol(Col)
                bDrawSelLine = False
                SetRect R, R.Right, lTop, R.Right + mCol(X).Width, lTop + mRow(Y).Height
                
                If R.Right > 0 And mRow(Y).TempHeight = 0 And mCol(X).TempWidth = 0 Then
                    
                    BkColor = GetCellBackColor(Y, X, Row)
                    Set oFont = GetCellFont(Y, X)
                    Align = GetCellAlign(Y, X)
                    Flags = GetCellWordBreak(Y, X) Or DT_NOPREFIX Or DT_EXPANDTABS
                    
                    If m_SelectionMode = GP_SelBySingleRow And SelCel.Y = Row Then
                        BkColor = m_SelectionColor
                        VLineColor = ShiftColor(m_RowsBackColor, m_SelectionColor, 50)
                        HLineColors = VLineColor
                        bDrawSelLine = True
                        
                    ElseIf SelCel.X = Col And SelCel.Y = Row And Text1.Visible = False Then
                        BkColor = m_SelectionColor
                        VLineColor = ShiftColor(m_RowsBackColor, m_SelectionColor, 50)
                        HLineColors = VLineColor
                        bDrawSelLine = True
                        
                    ElseIf IsPointInSelRange(Col, Row) And Text1.Visible = False Then
                        If m_SelectionMode = GP_SelByMultiRow Then
                            BkColor = m_SelectionColor
                            VLineColor = ShiftColor(m_RowsBackColor, m_SelectionColor, 50)
                            HLineColors = VLineColor
                        Else
                            BkColor = ShiftColor(BkColor, m_SelectionColor, 150) '     &HF9EAD8
                            VLineColor = ShiftColor(m_LinesVerticalColor, m_SelectionColor, 120)
                            HLineColors = VLineColor
                        End If
                        bDrawSelLine = True
                    Else
                        VLineColor = m_LinesVerticalColor
                        HLineColors = m_LinesHorizontalColor
                    End If
                    
                    'HOT ROWS
                    If m_ShowHotRow Then
                        If (mHotCell.Y = Row) Or (HotRow = Row) Then
                            BkColor = ShiftColor(m_SelectionColor, BkColor, 30)
                        End If
                    End If
                    
                    If m_LinesVerticalWidth > 1 Then VLineColor = m_LinesVerticalColor
                    If m_LinesHorizontalWidth > 1 Then HLineColors = m_LinesHorizontalColor
                    
                    lForeColor = GetCellForeColor(Y, X, BkColor)
                    SetTextColor hdc, lForeColor
               
                    FillRectangle UserControl.hdc, R, BkColor
                    
                    '/>Jose Liza - FocusRect
                    If Me.FocusRectMode <> lgNone Then
                        If m_CellFocusRect.Y = Row And m_CellFocusRect.X = Col Then
                            lForeColor = mRow(Row).Cells(Col).ForeColor
                            SetTextColor hdc, lForeColor
                            '---
                            Select Case Me.FocusRectMode
                                Case lgCol
                                    With m_CellRectInfo
                                        .Left = R.Left + IIF(Col = 0, 1, 0) - IIF(m_ShowHotColumn, 1, 0)
                                        .Top = R.Top - IIF(Row > 0, 1, 0)
                                        .Right = R.Right
                                        .Bottom = R.Bottom
                                    End With
                            End Select
                            '---
                            If m_FocusRectMode = lgCol Then FillRectangle hdc, m_CellRectInfo, IIF(Row Mod 2, m_RowsBackColor, m_RowsBackColorAlt)
                        End If
                    End If
                    '/<Jose Liza - FocusRect
                    
                    '------------TreeLines
                    If m_GroupsTreeStyle Then
                        If Col = 0 And mRow(Y).Ident > 0 Then DrawTreeViewLines hdc, Row, Y, R
                    End If
                    
                    Call SelectObject(UserControl.hdc, oFont.hFont)
                    GetCellRectsParts Y, X, CBool(X = 0), Flags, Align, R, RCheck, RText, RImage
                    
                    If mHotPart.Y = Row And mHotPart.X = Col Then
                        If HotPartIsImage Then
                            RoundRectPlus UserControl.hdc, RImage.Left - Margin / 2, RImage.Top - Margin / 2, RImage.Right - RImage.Left + Margin, RImage.Bottom - RImage.Top + Margin, RGBtoARGB(vbButtonFace, 50), RGBtoARGB(vbButtonShadow, 50), 2 * DpiF
                        Else
                            RoundRectPlus UserControl.hdc, RText.Left - Margin / 2, RText.Top - Margin / 2, RText.Right - RText.Left + Margin, RText.Bottom - RText.Top + Margin, RGBtoARGB(vbButtonFace, 50), RGBtoARGB(vbButtonShadow, 50), 2 * DpiF
                        End If
                    End If

                    If m_CheckStyle = True And X = 0 Then
                        DrawCheckBox UserControl.hdc, mRow(Y).Checked, BkColor = m_SelectionColor, RCheck, CenterLeft ' Align
                    End If
                    
                    
                    If mRow(Y).Cells(X).IconIndex > 0 Then
                        Dim ImgW As Long, ImgH As Long
                        
                        If mCol(X).ImagesMonocrome Then IconColors = lForeColor Else IconColors = CLR_NONE
                        
                        If mCol(X).ImgListWidth > 0 Then
                             hImage = mCol(X).ColImgList(mRow(Y).Cells(X).IconIndex)
                             ImgW = mCol(X).ImgListWidth
                             ImgH = mCol(X).ImgListHeight
                        Else
                             hImage = cHeaderImageList(mRow(Y).Cells(X).IconIndex)
                             ImgW = m_HeaderImgLstWidth
                             ImgH = m_HeaderImgLstHeight
                        End If
                        ImageListDraw hImage, hdc, RImage.Left, RImage.Top, ImgW, ImgH, IconColors
                     End If

 
                    If mCol(X).DataType = GP_BOOLEAN Then
                        DrawCheckBox UserControl.hdc, mRow(Y).Cells(X).Value, BkColor = m_SelectionColor, RText, CenterCenter
                    Else
                        If Not (mCellEdit.X = Col And mCellEdit.Y = Row) Then
                            If Not mCol(X).TextHide Then
                                DrawText UserControl.hdc, StrPtr(GetCellText(Y, X)), -1, RText, Flags 'Or CenterCenter 'Or DT_WORDBREAK
                            End If
                        End If
                    End If
                    
                    RaiseEvent CellPostPaint(hdc, Row, Col, R.Left, R.Top, R.Right, R.Bottom)
                    
                    If LVW > 0 Then DrawLine2 hdc, R.Right - LVW, R.Top, R.Right, R.Bottom, VLineColor
                    If LHW > 0 Then DrawLine2 hdc, R.Left, R.Bottom - LHW, R.Right, R.Bottom, HLineColors


                    If bDrawSelLine Then
                        If LVW > 0 Then DrawLine2 hdc, R.Left - LVW, R.Top, R.Left, R.Bottom, VLineColor
                        If LHW > 0 Then DrawLine2 hdc, R.Left - LVW, R.Top - LHW, R.Right, R.Top, HLineColors
                    End If

                    If FlagFooter = True Then
                        If LHW > 0 Then DrawLine2 hdc, R.Left, R.Top - LHW * 2, R.Right, R.Top, HLineColors
                    End If

                    If (PaintPart = 2 Or PaintPart = 3) And Row = m_FixedRows - 1 Then
                        DrawLine2 hdc, R.Left, R.Bottom - LHW * 2, R.Right, R.Bottom, HLineColors
                    End If

                    If (PaintPart = 1 Or PaintPart = 3) And Col = m_FixedColumns - 1 Then
                        DrawLine2 hdc, R.Right - LVW * 2, R.Top, R.Right, R.Bottom, HLineColors
                    End If

                    If mColDrag.X > -1 And mColDrag.DestCol > -1 Then
                        If PtrCol(mColDrag.DestCol) = X Then
                            If mColDrag.DestCol > mColDrag.SrcCol Then
                                DrawLine2 UserControl.hdc, R.Right - 2 * DpiF, 0, R.Right, R.Bottom, m_SelectionColor
                            Else
                                DrawLine2 UserControl.hdc, R.Left, 0, R.Left + 2 * DpiF, R.Bottom, m_SelectionColor
                            End If
                        End If
                    End If

                    If mCellEdit.Y = Row And mCellEdit.X = Col Then
                        UserControl.ForeColor = m_SelectionColor
                        FillRectangle UserControl.hdc, R, Text1.BackColor
                        Rectangle hdc, R.Left, R.Top, R.Right - LVW, R.Bottom - LHW
                    End If
                End If
                If R.Left > ucScrollbarV.Left Then Exit For
            Next
        End If
   
        SetRect R2, m_RowSelectorWidth - ucScrollbarH.Value, lTop, ColumsWidth - ucScrollbarH.Value, lTop + mRow(Y).Height
        RaiseEvent RowPostPaint(hdc, Row, R2.Left, R2.Top, R2.Right, R2.Bottom)

        lTop = lTop + mRow(Y).Height
        If R.Top > ucScrollbarH.Top Then Exit For
    Next
    
    If m_LastRowIsFooter Then
        If R.Bottom > UserControl.ScaleHeight Then
    
            If FlagFooter = False Then
                FlagFooter = True
                R.Right = m_RowSelectorWidth
                Row = m_RowsCount - 1
                If ucScrollbarH.Visible Then
                    lTop = ucScrollbarH.Top - mRow(PtrRow(Row)).Height
                Else
                    lTop = UserControl.ScaleHeight - mRow(PtrRow(Row)).Height
                End If
                GoTo Draw_Footer
    
            End If
        End If
        FlagFooter = False
    End If
    
    VLineColor = m_LinesVerticalColor
    HLineColors = m_LinesHorizontalColor
    If PaintPart = 3 Then GoTo EndDrawFixed
    '==================================
    'Draw COLUMN
    '==================================
    If (PaintPart <> 2) And (m_HeaderHeight > 0) Then

        Set oFont = m_HeaderFont
        SelectObject UserControl.hdc, oFont.hFont

        If PaintPart = 0 Then
            lStart2 = m_FixedColumns
            lEnd2 = m_ColsCount - 1
            R.Right = -ucScrollbarH.Value + m_RowSelectorWidth + FixedColWidth
        ElseIf PaintPart = 1 Then
            lStart2 = 0
            lEnd2 = m_FixedColumns - 1
            R.Right = m_RowSelectorWidth
        End If
                
        For Col = lStart2 To lEnd2
            X = PtrCol(Col)
            SetRect R, R.Right, 0, R.Right + mCol(X).Width, m_HeaderHeight
            If R.Right > 0 And mCol(X).Width > 0 Then
                If Col = mHotCol And m_ShowHotColumn Then
                    If IsDarkColor(m_SelectionColor) Then
                        BkColor = ShiftColor(m_SelectionColor, m_HeaderBackColor, 25)
                    Else
                        BkColor = ShiftColor(m_SelectionColor, m_HeaderBackColor, 50)
                    End If
                Else
                    BkColor = m_HeaderBackColor
                End If
                
                If m_GradientStyle Then
                    FillGradientRect UserControl.hdc, R, ShiftColor(vbWhite, BkColor, 50), ShiftColor(vbBlack, BkColor, 50), True
                Else
                    FillRectangle UserControl.hdc, R, BkColor
                End If
                
                SetRect R2, R.Left + Margin, R.Top + Margin, R.Right - Margin - HLVW, R.Bottom - Margin - HLHW

                If mCol(X).HeaderForeColor = CLR_NONE Then
                    
                    SetTextColor hdc, IIF(IsDarkColor(m_HeaderBackColor), vbWhite, vbBlack)
                Else
                    lForeColor = mCol(X).HeaderForeColor
                    If (lForeColor And &H80000000) Then lForeColor = GetSysColor(lForeColor And &HFF&)
                    SetTextColor hdc, lForeColor
                End If

                IdIconFont = IIF(mCol(X).nSortOrder = lgSTAscending, IIF_SortAsc, IIF_SortDes)
                
                If mSortColumn = X Then
                    DrawIconFont hdc, IdIconFont, 9, False, R2, CenterRight
                    R2.Right = R2.Right - 16 * DpiF - Margin
                ElseIf mSortSubColumn = X Then
                    DrawIconFont hdc, IdIconFont, 5, False, R2, CenterRight
                    R2.Right = R2.Right - 16 * DpiF - Margin
                End If
                
                If m_CheckStyle And X = 0 Then
                    DrawCheckBox UserControl.hdc, m_AllRowAreCheked, False, R2, CenterLeft
                    R2.Left = R2.Left + 16 * DpiF + Margin
                End If

                
                If mCol(X).IconIndex > 0 Then
                    If Not cHeaderImageList Is Nothing Then
                        Align = IIF(m_HeaderImageAlign = HA_DefaultColumn, mCol(X).ImgAlign, m_HeaderImageAlign)
                        hImage = cHeaderImageList(mCol(X).IconIndex)
                        If (Align And DT_RIGHT) = DT_RIGHT Then
                            R2.Right = R2.Right - m_HeaderImgLstWidth '+ Margin * 2
                            ImageListDraw hImage, hdc, R2.Right, R2.Top + (R2.Bottom - R2.Top) / 2 - m_HeaderImgLstHeight / 2, m_HeaderImgLstWidth, m_HeaderImgLstHeight
                            R2.Right = R2.Right - Margin
                        ElseIf (Align And DT_CENTER) = DT_CENTER Then
                            ImageListDraw hImage, hdc, R2.Left + (R2.Right - R2.Left) / 2 - m_HeaderImgLstWidth / 2, R2.Top + (R2.Bottom - R2.Top) / 2 - m_HeaderImgLstHeight / 2, m_HeaderImgLstWidth, m_HeaderImgLstHeight
                        Else
                            ImageListDraw hImage, hdc, R2.Left, R2.Top + (R2.Bottom - R2.Top) / 2 - m_HeaderImgLstHeight / 2, m_HeaderImgLstWidth, m_HeaderImgLstHeight
                            R2.Left = R2.Left + m_HeaderImgLstWidth + Margin
                        End If
                        
                    End If
                End If
                
                '*-
                Align = GetColTextAlign(X)
                
                If m_HeaderTextWordBreak Then
                    Dim RT As Rect
                    RT = R2
                    DrawText UserControl.hdc, StrPtr(mCol(X).Text), -1, RT, DT_WORDBREAK Or DT_CALCRECT
                    R2.Top = R2.Top + (R2.Bottom - R2.Top) / 2 - (RT.Bottom - RT.Top) / 2
                    DrawText UserControl.hdc, StrPtr(mCol(X).Text), -1, R2, Align Or DT_VCENTER Or DT_WORDBREAK
                Else
                    DrawText UserControl.hdc, StrPtr(mCol(X).Text), -1, R2, DT_SINGLELINE Or Align Or DT_VCENTER Or DT_WORD_ELLIPSIS
                End If
                
                If HLVW > 0 Then DrawLine2 hdc, R.Right - HLVW, R.Top, R.Right, R.Bottom, VLineColor
                If HLHW > 0 Then DrawLine2 hdc, R.Left, R.Bottom - HLHW, R.Right, R.Bottom, HLineColors
                
                If PaintPart = 1 And Col = m_FixedColumns - 1 Then
                    DrawLine2 hdc, R.Right - LVW * 2, R.Top, R.Right, R.Bottom, HLineColors
                End If
    
                If mColDrag.X > -1 And mColDrag.DestCol > -1 Then
                    If PtrCol(mColDrag.DestCol) = X Then
                        If mColDrag.DestCol > mColDrag.SrcCol Then
                            DrawLine2 UserControl.hdc, R.Right - 2 * DpiF, 0, R.Right, R.Bottom, m_SelectionColor
                        Else
                            DrawLine2 UserControl.hdc, R.Left, 0, R.Left + 2 * DpiF, R.Bottom, m_SelectionColor
                        End If
                    End If
                End If
    
            End If
            If R.Left > ucScrollbarV.Left Then Exit For
        Next
    End If
    
    '/>Jose Liza - FocusRect
    Dim FocusW As Integer
    If Me.FocusRectMode <> lgNone Then
        '---
        If PaintPart = 0 Or PaintPart = 1 Then
            lTop = m_HeaderHeight + FixedRowHeight
        ElseIf PaintPart = 2 Or PaintPart = 3 Then
            lTop = m_HeaderHeight
        End If
        '---
        For Row = lStart1 To lEnd1
            Y = PtrRow(Row)
            For Col = lStart2 To lEnd2
                If m_CellFocusRect.Y = Row And m_CellFocusRect.X = Col Then
                    Select Case m_FocusRectMode
                        Case lgRow
                            SetRect m_CellRectInfo, m_RowSelectorWidth - ucScrollbarH.Value + 1, lTop, ColumsWidth - ucScrollbarH.Value, lTop + mRow(Y).Height
                    End Select
                    '---
                    Select Case m_FocusRectStyle
                        Case lgFRLight
                            DrawFocusCell hdc, m_CellRectInfo, IIF(Row Mod 2, m_RowsBackColor, RowsBackColorAlt)
                            DrawFocusRect hdc, m_CellRectInfo
                        Case lgFRHeavy, lgFRMedium
                            DrawFocusCell hdc, m_CellRectInfo, m_FocusRectColor, Abs(m_FocusRectStyle)
                    End Select
                End If
            Next
            lTop = lTop + mRow(Y).Height
        Next
    End If
    '/<Jose Liza - FocusRect
    
    If R.Right < UserControl.ScaleWidth And PaintPart = 0 Then
        SetRect R, R.Right, 0, UserControl.ScaleWidth, m_HeaderHeight
        'FillRectangle UserControl.Hdc, R, m_HeaderBackColor
        If m_GradientStyle Then
            FillGradientRect UserControl.hdc, R, ShiftColor(vbWhite, m_HeaderBackColor, 50), ShiftColor(vbBlack, m_HeaderBackColor, 50), True
        Else
            FillRectangle UserControl.hdc, R, m_HeaderBackColor
        End If
    
        
        If HLHW > 0 Then DrawLine2 hdc, R.Left, R.Bottom - HLHW, R.Right, R.Bottom, HLineColors
    End If

    If m_RowSelectorWidth > 0 And PaintPart <> 1 Then
        SetTextColor hdc, IIF(IsDarkColor(m_RowSelectorBkColor), vbWhite, vbBlack)
        
        If PaintPart = 0 Then
            lStart1 = VScrollPos + m_FixedRows
            lEnd1 = m_RowsCount - 1
            lTop = m_HeaderHeight + FixedRowHeight
        ElseIf PaintPart = 2 Then
            lStart1 = 0
            lEnd1 = m_FixedRows - 1
            lTop = m_HeaderHeight
        End If

        For Row = lStart1 To lEnd1
            Y = PtrRow(Row)
          
            SetRect R, 0, lTop, m_RowSelectorWidth, lTop + mRow(Y).Height
   
            
            If m_ShowHotRow And HotRow = Row Then
                BkColor = ShiftColor(m_SelectionColor, m_RowSelectorBkColor, 30)
            Else
                BkColor = m_RowSelectorBkColor
            End If
            
            If m_GradientStyle Then
                FillGradientRect UserControl.hdc, R, ShiftColor(vbBlack, BkColor, 50), ShiftColor(vbWhite, BkColor, 50), False
            Else
                FillRectangle UserControl.hdc, R, BkColor
            End If
            
            If LVW > 0 Then DrawLine2 hdc, R.Right - LVW, R.Top, R.Right, R.Bottom, VLineColor
            If LHW > 0 Then DrawLine2 hdc, R.Left, R.Bottom - LHW, R.Right, R.Bottom, HLineColors
            If PaintPart = 2 And Row = m_FixedRows - 1 Then
                DrawLine2 hdc, R.Left, R.Bottom - LHW * 2, R.Right, R.Bottom, HLineColors
            End If
                        
            If mCellEdit.Y = Row Then
                DrawIconFont hdc, IIF_Edit, m_Font.Size, False, R, CenterCenter
            ElseIf Row = m_RowsCount - 1 Then
                DrawIconFont hdc, IIF_Asterisk, m_Font.Size, False, R, CenterCenter
            ElseIf Row = SelCel.Y Then
                DrawIconFont hdc, IIF_ArrowRight, m_Font.Size, False, R, CenterCenter
            End If
            
            lTop = lTop + mRow(Y).Height
            If R.Top > ucScrollbarH.Top Then Exit For
        Next


    End If
    
    '---------------------

    If m_FixedRows > 0 Then
        If PaintPart = 0 Then
            PaintPart = 2
            GoTo DrawFixed
        End If
    End If
    
    If m_FixedColumns > 0 Then
        If PaintPart = 0 Or PaintPart = 2 Then
            PaintPart = 1
            GoTo DrawFixed
        End If
    End If
    
    If m_FixedRows > 0 And m_FixedColumns > 0 Then
        If PaintPart <> 3 Then
            PaintPart = 3
            GoTo DrawFixed
        End If
    End If
    
EndDrawFixed:
    '----------------------
    If m_RowSelectorWidth > 0 Then
        'Corner TopLeft
        SetRect R, 0, 0, m_RowSelectorWidth, m_HeaderHeight
        'FillRectangle UserControl.Hdc, R, m_HeaderBackColor

        If m_GradientStyle Then
            FillGradientRect UserControl.hdc, R, ShiftColor(vbWhite, m_HeaderBackColor, 50), ShiftColor(vbBlack, m_HeaderBackColor, 50), True
        Else
            FillRectangle UserControl.hdc, R, m_HeaderBackColor
        End If



        If HLVW > 0 Then DrawLine2 hdc, R.Right - HLVW, R.Top, R.Right, R.Bottom, VLineColor
        If HLHW > 0 Then DrawLine2 hdc, R.Left, R.Bottom - HLHW, R.Right, R.Bottom, HLineColors
    End If
    
    If mColDrag.X > -1 Then
        SetRect R, mColDrag.Left - mColDrag.X, 0, mColDrag.Left - mColDrag.X + mCol(PtrCol(mColDrag.SrcCol)).Width, m_HeaderHeight
        FillRectangle UserControl.hdc, R, ShiftColor(m_HeaderBackColor, m_SelectionColor, 150)
        SetRect R2, R.Left + Margin, R.Top + Margin, R.Right - Margin, R.Bottom - Margin
        Align = GetColTextAlign(PtrCol(mColDrag.SrcCol))
        
        If m_HeaderTextWordBreak Then
            RT = R2
            DrawText UserControl.hdc, StrPtr(mCol(PtrCol(mColDrag.SrcCol)).Text), -1, RT, DT_WORDBREAK Or DT_CALCRECT
            R2.Top = R2.Top + (R2.Bottom - R2.Top) / 2 - (RT.Bottom - RT.Top) / 2
            DrawText UserControl.hdc, StrPtr(mCol(PtrCol(mColDrag.SrcCol)).Text), -1, R2, Align Or DT_VCENTER Or DT_WORDBREAK
        Else
            DrawText UserControl.hdc, StrPtr(mCol(PtrCol(mColDrag.SrcCol)).Text), -1, R2, DT_SINGLELINE Or Align Or DT_VCENTER Or DT_WORD_ELLIPSIS
        End If
        
        
        
        'DrawText UserControl.hdc, StrPtr(mCol(PtrCol(mColDrag.SrcCol)).Text), -1, R2, DT_SINGLELINE Or Align Or DT_VCENTER
        UserControl.ForeColor = m_SelectionColor
        Rectangle hdc, R.Left, R.Top + DpiF, R.Right, R.Bottom
    End If

    'Corner BottomRight
    If ucScrollbarV.Visible And ucScrollbarH.Visible Then
        SetRect R, ucScrollbarV.Left, ucScrollbarH.Top, UserControl.ScaleWidth, UserControl.ScaleHeight
        If ucScrollbarV.Style = sGoogle Then
            FillRectangle UserControl.hdc, R, m_RowsBackColor
        Else
            FillRectangle UserControl.hdc, R, vbButtonFace
        End If
    End If
  '*1
    If m_BorderWidth > 0 Then

        If m_BorderRadius = 0 Then
            UserControl.ForeColor = m_BorderColor
            UserControl.DrawWidth = m_BorderWidth * 2
            Rectangle hdc, 0, 0, UserControl.ScaleWidth + 1, UserControl.ScaleHeight + 1
        Else
            RoundBorders hdc, m_BorderWidth / 3, m_BorderWidth / 3, UserControl.ScaleWidth - m_BorderWidth - FixLine(m_BorderWidth), UserControl.ScaleHeight - m_BorderWidth - FixLine(m_BorderWidth), RGBtoARGB(m_ParentBackColor, 100), RGBtoARGB(m_BorderColor, 100), m_BorderRadius * DpiF
        End If
    End If
    
    UserControl.Refresh
End Sub

Private Function GetColTextAlign(ByVal Col As Long) As Long
    If m_HeaderTextAlign = HA_DefaultColumn Then
        If (mCol(Col).Align And DT_RIGHT) = DT_RIGHT Then
            GetColTextAlign = DT_RIGHT
        ElseIf (mCol(Col).Align And DT_CENTER) = DT_CENTER Then
            GetColTextAlign = DT_CENTER
        Else
            GetColTextAlign = DT_LEFT
        End If
    Else
        GetColTextAlign = m_HeaderTextAlign
    End If
End Function


Private Function FixLine(ByVal lWidth As Long) As Long
    If lWidth = 2 Then FixLine = 1
End Function

Private Function GetCellText(Row As Long, Col As Long) As String
    If Not IsNull(mRow(Row).Cells(Col).Value) Then
        If mCol(Col).DataType = GP_CURRENCY Then
            If IsNumeric(mRow(Row).Cells(Col).Value) Then
                GetCellText = FormatNumber(mRow(Row).Cells(Col).Value, , vbTrue, vbTrue)
            Else
                GetCellText = mRow(Row).Cells(Col).Value
            End If
            If Len(mCol(Col).Format) Then GetCellText = Format(GetCellText, mCol(Col).Format)
        ElseIf Len(mCol(Col).Format) Then
            GetCellText = Format(mRow(Row).Cells(Col).Value, mCol(Col).Format)
        Else
            GetCellText = mRow(Row).Cells(Col).Value
        End If
    End If
End Function

Private Sub RoundBorders(ByVal hdc As Long, ByVal Left As Long, ByVal Top As Long, _
                       ByVal Width As Long, ByVal Height As Long, ByVal BackColor As Long, ByVal BorderColor As Long, Radius As Long)
    Dim hPen As Long, hBrush As Long
    Dim mPath As Long, hGraphics As Long
    
    GdipCreateFromHDC hdc, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    
    If BackColor <> 0 Then GdipCreateSolidFill BackColor, hBrush
    If BorderColor <> 0 Then GdipCreatePen1 BorderColor, m_BorderWidth, &H2, hPen
    
    If Radius = 0 Then
      
        If hBrush Then GdipFillRectangleI hGraphics, hBrush, Left, Top, Width, Height
        If hPen Then GdipDrawRectangleI hGraphics, hPen, Left, Top, Width, Height
        
    Else
        If GdipCreatePath(&H0, mPath) = 0 Then
            
            GdipAddPathArcI mPath, Left, Top, Radius, Radius, 180, 90
            GdipAddPathArcI mPath, Left + Width - Radius, Top, Radius, Radius, 270, 90
            GdipAddPathArcI mPath, Left + Width - Radius, Top + Height - Radius, Radius, Radius, 0, 90
            GdipAddPathArcI mPath, Left, Top + Height - Radius, Radius, Radius, 90, 90
            GdipClosePathFigure mPath
            GdipSetClipPath hGraphics, mPath, &H3
            GdipFillRectangleI hGraphics, hBrush, Left - m_BorderWidth, Top - m_BorderWidth, Width + m_BorderWidth * 2, Height + m_BorderWidth * 2
            GdipResetClip hGraphics
            If hPen Then
                'GdipSetPenMode hPen, PenAlignmentInset
                GdipDrawPath hGraphics, hPen, mPath
            End If
            Call GdipDeletePath(mPath)
        End If
    End If
        
    If hBrush Then Call GdipDeleteBrush(hBrush)
    If hPen Then Call GdipDeletePen(hPen)
    GdipDeleteGraphics hGraphics
End Sub

Private Sub Text1_GotFocus()
    If mSelectAllOnFocus Then
        mSelectAllOnFocus = False
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim bUpdate As Boolean
    Dim mPT As POINTAPI
    Dim bCancel As Boolean

    RaiseEvent EditKeyDown(KeyCode, Shift)
    
    mPT = SelCel
    
    Select Case KeyCode
        Case vbKeyRight: If Text1.SelStart = Len(Text1.Text) Then If Text1.SelLength = 0 Then mPT.X = mPT.X + 1: bUpdate = True
        Case vbKeyLeft: If Text1.SelStart = 0 Then If Text1.SelLength = 0 Then mPT.X = mPT.X - 1: bUpdate = True
        Case vbKeyDown: mPT.Y = mPT.Y + 1: bUpdate = True
        Case vbKeyUp: mPT.Y = mPT.Y - 1: bUpdate = True
    End Select

    If bUpdate Then
        If mPT.Y > -1 And mPT.Y < m_RowsCount And mPT.X > -1 And mPT.X < m_ColsCount Then
            SelCel = mPT
            m_Redraw = False
            CellSaveEdit
            
            If Not mCol(PtrCol(mPT.X)).DataType = GP_BOOLEAN And Not mRow(PtrRow(mPT.Y)).Cells(PtrCol(mPT.X)).EditionLocked = True And Not mCol(PtrCol(mPT.X)).EditionLocked = True Then
                If mCol(PtrCol(mPT.X)).TempWidth = 0 Then
                    RaiseEvent BeforeEdit(SelCel.Y, SelCel.X, CellValue(SelCel.Y, SelCel.X), bCancel)
                    If Not bCancel Then
                        mSelectAllOnFocus = True
                        CellStartEdit SelCel.Y, SelCel.X
                        m_Redraw = True
                        Draw
                    End If
                    Exit Sub
                End If
            End If
            m_Redraw = True
            Draw
        End If
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

   RaiseEvent EditKeyPress(KeyAscii)
    
    
    If KeyAscii = vbKeyReturn Then
        Text1.Enabled = False 'unitextbox
        CellSaveEdit
        UserControl.SetFocus
        Text1.Enabled = True
        KeyAscii = 0
    End If
    
    If KeyAscii = vbKeyEscape Then
        Text1.Enabled = False 'unitextbox
        Text1.Visible = False
        EmptyPoint mCellEdit
        UserControl.SetFocus
        Text1.Enabled = True
        KeyAscii = 0
        Draw
    End If
End Sub

Private Sub Timer1_Timer()
    Dim PT As POINTAPI

    GetCursorPos PT
    ScreenToClient UserControl.hwnd, PT

    If GetKeyState(vbLeftButton) < 0 Then
        UserControl_MouseMove vbLeftButton, 0, CSng(PT.X), CSng(PT.Y)
    End If
End Sub

Private Sub ucScrollbarV_Change()
    Dim lCount As Long
    
    If m_RowsCount = 0 Then
        Draw
        Exit Sub
    End If
    
    VScrollPos = 0
    Do While lCount < ucScrollbarV.Value / 10
        If mRow(PtrRow(VScrollPos)).TempHeight = 0 Then
            lCount = lCount + 1
        End If
        VScrollPos = VScrollPos + 1
    Loop
    

    Text1.Visible = False
    EmptyPoint mCellEdit
    Draw

    RaiseEvent VerticalScroll
    If ucScrollbarV.Value = ucScrollbarV.Max Then
        RaiseEvent VerticalScrollEnd
    End If
End Sub


Private Sub ucScrollbarH_Change()
    Text1.Visible = False
    EmptyPoint mCellEdit
    Draw
    RaiseEvent HorizontalScroll
End Sub


Public Sub AutoWidthColumn(ByVal Col As Long)
    Dim MaxW As Long, Row As Long, R As Rect
    Dim oFont As iFont
    Dim Y As Long
    
    Col = PtrCol(Col)
    Set oFont = m_HeaderFont
    SelectObject UserControl.hdc, oFont.hFont
    
    DrawText UserControl.hdc, StrPtr(mCol(Col).Text), -1, R, DT_SINGLELINE Or DT_NOPREFIX Or DT_EXPANDTABS Or DT_CALCRECT
    MaxW = R.Right + 16 + DpiF + Margin
    If mCol(Col).DataType = GP_BOOLEAN Then
        If MaxW < 16 * DpiF Then MaxW = 16 * DpiF
    Else
        For Y = 0 To m_RowsCount - 1
            Row = PtrRow(Y)
            Set oFont = GetCellFont(Row, Col)
            SelectObject UserControl.hdc, oFont.hFont
            DrawText UserControl.hdc, StrPtr(GetCellText(Row, Col)), -1, R, DT_SINGLELINE Or DT_NOPREFIX Or DT_EXPANDTABS Or DT_CALCRECT
            If R.Right > MaxW Then MaxW = R.Right
        Next
    End If
    
    MaxW = MaxW + Margin * 2 + m_LinesVerticalWidth * DpiF
    
    If Col = 0 And m_CheckStyle = True Then
        MaxW = MaxW + 16 * DpiF
    End If
    
    If Not mCol(Col).ColImgList Is Nothing Then
        If Not (mCol(Col).ImgAlign And DT_CENTER) = DT_CENTER Then
            MaxW = MaxW + mCol(Col).ImgListWidth + Margin * 2
        End If
    End If
    
    If mCol(Col).TempWidth > 0 Then
        mCol(Col).TempWidth = MaxW
    Else
        mCol(Col).Width = MaxW
    End If
    If m_Redraw Then Me.Refresh
End Sub

Public Sub AutoWidthAllColumns()
    Dim i As Long, bRedraw As Boolean
    bRedraw = m_Redraw
    m_Redraw = False
    For i = 0 To m_ColsCount - 1
        AutoWidthColumn i
    Next
    m_Redraw = bRedraw
    If m_Redraw Then Me.Refresh
End Sub

Public Sub AutoHeightRow(ByVal Row As Long)
    Dim MaxH As Long, X As Long, Col As Long, R As Rect
    Dim Flags As Long
    Dim oFont As iFont

    
    Row = PtrRow(Row)
    
    If mRow(Row).IsFullRow Or mRow(Row).IsGroup Then
        Col = PtrCol(0)
        Set oFont = GetCellFont(Row, Col)
        Flags = GetCellWordBreak(Row, Col) Or DT_NOPREFIX Or DT_EXPANDTABS
        Call SelectObject(UserControl.hdc, oFont.hFont)
        
        R.Right = Me.ColLeft(PtrCol(m_ColsCount - 1)) + mCol(PtrCol(m_ColsCount - 1)).Width

        DrawText UserControl.hdc, StrPtr(GetCellText(Row, Col)), -1, R, Flags Or DT_CALCRECT
        MaxH = R.Bottom
        
        If Not mCol(Col).ColImgList Is Nothing Then
            If mCol(Col).ImgListHeight + Margin > MaxH Then
                MaxH = mCol(Col).ImgListHeight + Margin
            End If
        End If
        
    Else
        For X = 0 To m_ColsCount - 1
            Col = PtrCol(X)
            Set oFont = GetCellFont(Row, Col)
            Flags = GetCellWordBreak(Row, Col) Or DT_NOPREFIX Or DT_EXPANDTABS
            Call SelectObject(UserControl.hdc, oFont.hFont)
            R.Right = mCol(Col).Width
            DrawText UserControl.hdc, StrPtr(GetCellText(Row, Col)), -1, R, Flags Or DT_CALCRECT
            If R.Bottom > MaxH Then MaxH = R.Bottom
            
            If Not mCol(Col).ColImgList Is Nothing Then
                If mCol(Col).ImgListHeight + Margin > MaxH Then
                    MaxH = mCol(Col).ImgListHeight + Margin
                End If
            End If
        Next
    End If

    If mRow(Row).TempHeight > 0 Then
        mRow(Row).TempHeight = MaxH + Margin + m_LinesHorizontalWidth
    Else
        mRow(Row).Height = MaxH + Margin * 2 + m_LinesHorizontalWidth
    End If
    If m_Redraw Then Me.Refresh
End Sub

Public Sub AutoHeightAllRows()
    Dim i As Long, bRedraw As Boolean
    bRedraw = m_Redraw
    m_Redraw = False
    For i = 0 To m_RowsCount - 1
        AutoHeightRow i
    Next
    m_Redraw = bRedraw
    If m_Redraw Then Me.Refresh
End Sub

Private Sub ucScrollbarV_ContainerMouseEnter()
    RaiseEvent MouseEnter
End Sub

Private Sub ucScrollbarV_ContainerMouseLeave()
    Dim bUpdate As Boolean
    If mHotCell.X <> -1 Or mHotCell.Y <> -1 Then
        EmptyPoint mHotCell
        RaiseEvent HotCellChange(mHotCell.Y, mHotCell.X)
    End If
    If m_ShowHotColumn Then
        If mHotCol <> -1 Then
            mHotCol = -1
            
            bUpdate = True
        End If
    End If
    
    If m_ShowHotRow Then
        If HotRow <> -1 Then
            HotRow = -1
            bUpdate = True
        End If
    End If

    If Not IsPointEmpty(mHotPart) Then
        EmptyPoint mHotPart
        bUpdate = True
    End If
    If bUpdate Then Draw
    RaiseEvent MouseLeave
End Sub

Private Sub ucScrollbarV_OnPaint(ByVal lhDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal ePart As sbOnPaintPartCts, ByVal estate As sbOnPaintPartStateCts)
    RaiseEvent OnScrollPaint(True, hdc, X1, Y1, X2, Y2, ePart, estate)
End Sub

Private Sub ucScrollbarH_OnPaint(ByVal lhDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal ePart As sbOnPaintPartCts, ByVal estate As sbOnPaintPartStateCts)
    RaiseEvent OnScrollPaint(False, hdc, X1, Y1, X2, Y2, ePart, estate)
End Sub

Private Sub UserControl_DblClick()
    Dim bCancel As Boolean
    
    If SizeColumn.HotColumn <> -1 Then
        AutoWidthColumn SizeColumn.HotColumn
        UserControl.MousePointer = vbDefault
        RaiseEvent ColumnUserResize(SizeColumn.HotColumn)
        GoTo EventDblClick
    End If
    
    If SizeColumn.HotRow <> -1 Then
        AutoHeightRow SizeColumn.HotRow
        UserControl.MousePointer = vbDefault
        GoTo EventDblClick
    End If

    If Not IsPointEmpty(SelCel) And eSelBy = SelectByCells Then
        If mCol(PtrCol(SelCel.X)).DataType = GP_BOOLEAN Then GoTo EventDblClick
        If mRow(PtrRow(SelCel.Y)).IsGroup = True Then GoTo EventDblClick
        If mRow(PtrRow(SelCel.Y)).IsFullRow = True Then GoTo EventDblClick
        If Not IsPointEmpty(mHotPart) Then GoTo EventDblClick
        If mCol(PtrCol(SelCel.X)).EditionLocked Then GoTo EventDblClick
        If mRow(PtrRow(SelCel.Y)).Cells(PtrCol(SelCel.X)).EditionLocked Then GoTo EventDblClick

        RaiseEvent BeforeEdit(SelCel.Y, SelCel.X, CellValue(SelCel.Y, SelCel.X), bCancel)
        
        If Not bCancel Then CellStartEdit SelCel.Y, SelCel.X
    End If
    
EventDblClick:
    RaiseEvent DblClick
End Sub

Private Sub CellSaveEdit()
    Dim mPT As POINTAPI
    Dim vOldValue As Variant
    vOldValue = mRow(PtrRow(mCellEdit.Y)).Cells(PtrCol(mCellEdit.X)).Value
    mRow(PtrRow(mCellEdit.Y)).Cells(PtrCol(mCellEdit.X)).Value = Text1.Text
    Text1.Visible = False
    mPT = mCellEdit
    EmptyPoint mCellEdit
    
    RaiseEvent AfterEdit(mPT.Y, mPT.X, vOldValue)
    Draw
End Sub

Public Sub CellStartEdit(ByVal Row As Long, ByVal Col As Long)
    Dim Flags As Long
    Dim Align As eGridAlign
    Dim oFont As iFont
    Dim RCheck As Rect, RText As Rect, RImage As Rect, R As Rect
    
    If m_AllowEdit = False Then Exit Sub
    
    EnsureCellVisible Row, Col

    If PvGetCellRect(Row, Col, R) Then
        EmptyPoint SelRange.Start
        EmptyPoint SelRange.End
        mCellEdit.Y = Row
        mCellEdit.X = Col
        Row = PtrRow(Row)
        Col = PtrCol(Col)

        Set oFont = GetCellFont(Row, Col)
        Align = GetCellAlign(Row, Col)
        Flags = GetCellWordBreak(Row, Col) Or DT_NOPREFIX Or DT_EXPANDTABS
    
        Call SelectObject(UserControl.hdc, oFont.hFont)

        GetCellRectsParts Row, Col, Col = 0, Flags, Align, R, RCheck, RText, RImage
        'solo para la version ocx
        '------------------------------
        Dim oText As Object
        Set oText = Text1
        If TypeName(oText) = "UniTextBox" Then
            If (Flags And DT_WORDBREAK) = DT_WORDBREAK Then
                oText.MultiLine = True
            Else
                oText.MultiLine = False
            End If
        End If
        '------------------------------
        With Text1
            .Text = MyString(mRow(Row).Cells(Col).Value)
            Set .Font = oFont

            .BackColor = GetCellBackColor(Row, Col, mCellEdit.Y)
            .ForeColor = GetCellForeColor(Row, Col, .BackColor)

            If RImage.Left < RText.Right Then RText.Right = R.Right - Margin
            If RImage.Right = 0 Then
                If RCheck.Right = 0 Then
                    RText.Left = R.Left + Margin
                Else
                    RText.Left = RCheck.Right + Margin
                End If
            End If

            .Move RText.Left, RText.Top, RText.Right - RText.Left, RText.Bottom - RText.Top

            Select Case Align
                Case TopLeft, CenterLeft, BottomLeft
                    .Alignment = 0
                Case TopCenter, CenterCenter, BottomCenter
                    .Alignment = 2
                Case TopRight, CenterRight, BottomRight
                    .Alignment = 1
            End Select

            .SelStart = 0
            .SelLength = Len(.Text)
            .Visible = True
            .SetFocus
            Draw
        End With
    End If
End Sub

Private Sub UserControl_EnterFocus()
    #If OCX_VERSION Then
    Call modIOleInPlaceActiveObject.SetIPAO(Me) 'OCX VERSION
    #End If
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim bUpdate As Boolean, NroRowsInScreen As Long
    Dim xCel As POINTAPI
    Dim mPos As Long
    Dim OldSelCel As POINTAPI
    
    If (Shift And vbShiftMask) = vbShiftMask Then
   
        If IsPointEmpty(SelRange.End) Then SelRange.End = SelCel
        If IsPointEmpty(SelRange.Start) Then SelRange.Start = SelCel
    
        If SelCel.X > SelRange.Start.X Then
            xCel.X = SelRange.Start.X
        Else
            xCel.X = SelRange.End.X
        End If
        If SelCel.Y > SelRange.Start.Y Then
            xCel.Y = SelRange.Start.Y
        Else
            xCel.Y = SelRange.End.Y
        End If
    Else
        xCel = SelCel
    End If

    Select Case KeyCode
        Case vbKeyDelete
            UserControl_KeyPress -1    ' Delete contents on edit
            
        Case vbKeyF2
            UserControl_KeyPress 0  ' Send KeyAscii 0 to start cell editing when it is allowed
            
        Case vbKeyRight
            If Not Text1.Visible Then
                If (Shift And vbCtrlMask) = vbCtrlMask Then
                    xCel.X = m_ColsCount - 1
                    Do While mCol(PtrCol(xCel.X)).Width = 0
                        xCel.X = xCel.X - 1
                        If xCel.X = 0 Then Exit Do
                    Loop
                    bUpdate = True
                Else
                    mPos = xCel.X
                    Do While xCel.X < m_ColsCount - 1
                        xCel.X = xCel.X + 1
                        If mCol(PtrCol(xCel.X)).Width > 0 Then Exit Do
                    Loop
                    If mCol(PtrCol(xCel.X)).Width = 0 Then xCel.X = mPos
                    bUpdate = True
                End If
            End If
        Case vbKeyLeft
            If Not Text1.Visible Then
                If (Shift And vbCtrlMask) = vbCtrlMask Then
                    xCel.X = 0
                    Do While mCol(PtrCol(xCel.X)).Width = 0
                        xCel.X = xCel.X + 1
                        If xCel.X = m_ColsCount - 1 Then Exit Do
                    Loop
                    bUpdate = True
                Else
                    If xCel.X = -1 Then Exit Sub
                    mPos = xCel.X
                    Do While xCel.X > 0
                        xCel.X = xCel.X - 1
                        If mCol(PtrCol(xCel.X)).Width > 0 Then Exit Do
                    Loop
                    If mCol(PtrCol(xCel.X)).Width = 0 Then xCel.X = mPos
                    bUpdate = True
                End If
            End If
        Case vbKeyDown
            If (Shift And vbCtrlMask) = vbCtrlMask Then
                UserControl_KeyDown vbKeyEnd, 0
                Exit Sub
            Else
                mPos = xCel.Y
                Do While xCel.Y < m_RowsCount - 1
                    xCel.Y = xCel.Y + 1
                    If mRow(PtrRow(xCel.Y)).Height > 0 Then Exit Do
                Loop
                If mRow(PtrRow(xCel.Y)).Height = 0 Then xCel.Y = mPos
                bUpdate = True
            End If
        Case vbKeyUp
            If (Shift And vbCtrlMask) = vbCtrlMask Then
                UserControl_KeyDown vbKeyHome, 0
                Exit Sub
            Else
                If xCel.Y = -1 Then Exit Sub
                mPos = xCel.Y
                Do While xCel.Y > 0
                    xCel.Y = xCel.Y - 1
                    If mRow(PtrRow(xCel.Y)).Height > 0 Then Exit Do
                Loop
                If mRow(PtrRow(xCel.Y)).Height = 0 Then xCel.Y = mPos
                bUpdate = True
            End If
        Case vbKeyHome
            xCel.Y = 0
            Do While mRow(PtrRow(xCel.Y)).Height = 0
                xCel.Y = xCel.Y + 1
                If xCel.Y = m_RowsCount - 1 Then Exit Do
            Loop
            bUpdate = True
        Case vbKeyEnd
            xCel.Y = m_RowsCount - 1
            Do While mRow(PtrRow(xCel.Y)).Height = 0
                xCel.Y = xCel.Y - 1
                If xCel.Y = 0 Then Exit Do
            Loop
            bUpdate = True
        Case vbKeyPageUp
            NroRowsInScreen = GetNroRowsInScreen
            xCel.Y = xCel.Y - NroRowsInScreen
            If xCel.Y < 0 Then
                xCel.Y = 0
                Do While mRow(PtrRow(xCel.Y)).Height = 0
                    xCel.Y = xCel.Y + 1
                    If xCel.Y = m_RowsCount - 1 Then Exit Do
                Loop
            Else
                Do While mRow(PtrRow(xCel.Y)).Height = 0
                    xCel.Y = xCel.Y - 1
                    If xCel.Y = 0 Then Exit Do
                Loop
            End If
            bUpdate = True
        Case vbKeyPageDown
            NroRowsInScreen = GetNroRowsInScreen
            xCel.Y = xCel.Y + NroRowsInScreen
            If xCel.Y > m_RowsCount - 1 Then
                xCel.Y = m_RowsCount - 1
                Do While mRow(PtrRow(xCel.Y)).Height = 0
                    xCel.Y = xCel.Y - 1
                    If xCel.Y = 0 Then Exit Do
                Loop
            Else
                Do While mRow(PtrRow(xCel.Y)).Height = 0
                    xCel.Y = xCel.Y + 1
                    If xCel.Y = m_RowsCount - 1 Then Exit Do
                Loop
            End If
            bUpdate = True
    End Select
    
    '/>Jose Liza - FocusRect
    Call UpdateFocusRect(xCel, All)
    '/>Jose Liza - FocusRect
    
    If bUpdate Then
        Dim R As Rect
        
        If (Shift And vbShiftMask) = vbShiftMask Then
            PvSelect xCel.X, xCel.Y
        Else
            OldSelCel = SelCel
            SelCel = xCel
          
            If SelCel.Y > m_RowsCount - 1 Then SelCel.Y = m_RowsCount - 1
            If SelCel.X > m_ColsCount - 1 Then SelCel.X = m_ColsCount - 1
            
            Select Case m_SelectionMode
                Case GP_SelFree
                    'Free Selection
                    SelRange.Start = SelCel
                    SelRange.End = SelCel
                Case GP_SelByMultiRow, GP_SelBySingleRow
                    'Sel by Row
                    With SelRange
                        .Start.Y = SelCel.Y
                        .Start.X = 0
                        .End.Y = SelCel.Y
                        .End.X = m_ColsCount - 1
                    End With
                Case GP_SelByCol
                    'Sel by Col
                    With SelRange
                        .Start.Y = 0
                        .Start.X = SelCel.X
                        .End.Y = m_RowsCount - 1
                        .End.X = SelCel.X
                    End With
                Case GP_SelBySingleCell
                    EmptyPoint SelRange.Start
                    EmptyPoint SelRange.End
            End Select
            
            RaiseEvent CurCellChange(SelCel.Y, SelCel.X, OldSelCel.Y, OldSelCel.X)

        End If

        EnsureCellVisible xCel.Y, xCel.X
        Draw
    End If
    
    RaiseEvent KeyDown(KeyCode, Shift)

End Sub


Private Function GetNroRowsInScreen() As Long
    Dim lTop As Long, lRow As Long, lCount As Long
    Dim mBootom As Long

    mBootom = IIF(ucScrollbarH.Visible, ucScrollbarH.Top, UserControl.ScaleHeight)
    If m_LastRowIsFooter Then mBootom = mBootom - mRow(m_RowsCount - 1).Height
    
    lTop = m_HeaderHeight
    For lRow = VScrollPos To m_RowsCount - 1
        lTop = lTop + mRow(PtrRow(lRow)).Height
        If lTop > mBootom Then Exit For
        lCount = lCount + 1
    Next
    GetNroRowsInScreen = lCount - 1
End Function

Public Sub EnsureCellVisible(ByVal Row As Long, ByVal Col As Long)
    Dim lLeft As Long, lCol As Long, NroRowsInScreen As Long
    Dim lDif As Long, bRedraw As Boolean
    Dim R As Rect
    Dim mBootom As Long
    Dim mRight As Long
    Dim mLeft As Long
    
    If Row >= m_RowsCount - 1 Then
        If m_LastRowIsFooter = False Then
            ucScrollbarV.Value = ucScrollbarV.Max
        End If
    End If

    If Col >= m_ColsCount - 1 Then
        ucScrollbarH.Value = ucScrollbarH.Max
    End If

    mRight = IIF(ucScrollbarV.Visible, ucScrollbarV.Left, UserControl.ScaleWidth)
    mBootom = IIF(ucScrollbarH.Visible, ucScrollbarH.Top, UserControl.ScaleHeight)
   
   If Col < m_FixedColumns Then
        ucScrollbarH.Value = 0
    Else
        For lCol = 0 To m_FixedColumns - 1
            mLeft = mLeft + mCol(PtrCol(lCol)).Width
        Next
    End If
    
    If m_LastRowIsFooter Then mBootom = mBootom - mRow(PtrRow(m_RowsCount - 1)).Height

    NroRowsInScreen = GetNroRowsInScreen

    If Row < VScrollPos + m_FixedRows Then
        If Row - m_FixedRows > -1 Then
            ucScrollbarV.Value = (Row - mHideRowsCount - m_FixedRows) * 10
            VScrollPos = Row - m_FixedRows
        Else
            ucScrollbarV.Value = (Row - mHideRowsCount) * 10
            VScrollPos = Row
        End If
        If VScrollPos < 0 Then VScrollPos = 0
    End If
    
    If Row > VScrollPos + NroRowsInScreen Then
        If mHideRowsCount = 0 Then
            ucScrollbarV.Value = (Row - NroRowsInScreen) * 10
            VScrollPos = Row - NroRowsInScreen
        Else
            bRedraw = Me.Redraw
            m_Redraw = False
            ucScrollbarV.Value = (Row - mHideRowsCount) * 10
            VScrollPos = Row - NroRowsInScreen
            Call PvGetCellRect(Row, 0, R, True)
    
            Do While R.Bottom > mBootom
                ucScrollbarV.Value = ucScrollbarV.Value + 10
                PvGetCellRect Row, 0, R
                If ucScrollbarV.Value = ucScrollbarV.Max Then
                    Exit Do
                End If
            Loop
            m_Redraw = bRedraw
        End If
    End If
    
    For lCol = 0 To m_ColsCount - 1
        If lCol = Col Then
            If lLeft < ucScrollbarH.Value + mLeft Then ucScrollbarH.Value = lLeft - mLeft
            lDif = mRight - lLeft + mCol(PtrCol(lCol)).Width - ucScrollbarH.Value + m_RowSelectorWidth
            If lLeft + mCol(PtrCol(lCol)).Width - ucScrollbarH.Value + m_RowSelectorWidth > mRight Then
                ucScrollbarH.Value = lLeft - mRight + mCol(PtrCol(lCol)).Width + m_RowSelectorWidth
            End If
            Exit Sub
        End If
        lLeft = lLeft + mCol(lCol).Width
    Next
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    Dim bCancel As Boolean
    
    If KeyAscii <> 0 Then RaiseEvent KeyPress(KeyAscii)
    
    If Not IsPointEmpty(SelCel) Then
        If mCol(SelCel.X).DataType = GP_BOOLEAN Then 'if is Boolean click
            Select Case KeyAscii
            Case vbKeySpace, vbKeyReturn
                RaiseEvent BeforeEdit(SelCel.Y, SelCel.X, mRow(SelCel.Y).Cells(SelCel.X).Value, bCancel)
                If Not bCancel Then
                    mRow(SelCel.Y).Cells(SelCel.X).Value = Not mRow(SelCel.Y).Cells(SelCel.X).Value
                    RaiseEvent AfterEdit(SelCel.Y, SelCel.X, Not mRow(SelCel.Y).Cells(SelCel.X).Value)
                    Draw
                End If
            End Select
            Exit Sub
        Else
            If KeyAscii = vbKeyReturn Then Exit Sub 'Filter Key Enter
            If KeyAscii = vbKeyBack Then KeyAscii = 0
            If KeyAscii > 0 And KeyAscii < 32 Then Exit Sub
        End If
        If mRow(PtrRow(SelCel.Y)).IsGroup = True Then Exit Sub
        If mRow(PtrRow(SelCel.Y)).IsFullRow = True Then Exit Sub
        
        If mCol(PtrCol(SelCel.X)).EditionLocked = False Then
            If mRow(PtrRow(SelCel.Y)).Cells(PtrCol(SelCel.X)).EditionLocked = False Then
                RaiseEvent BeforeEdit(SelCel.Y, SelCel.X, CellValue(SelCel.Y, SelCel.X), bCancel)
                
                If Not bCancel Then
                    CellStartEdit SelCel.Y, SelCel.X
                    Select Case KeyAscii
                    Case Is > 0
                        Text1.Text = Chr(KeyAscii)
                        Text1.SelStart = Len(Text1.Text)
                    Case -1
                        ' Delete contents
                        Text1.Text = ""
                    End Select
                End If
            End If
        End If
    End If

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim R As Rect, R2 As Rect, vOldValue As Variant
    Dim bCheckRow As Boolean, bText As Boolean, bImage As Boolean
    Dim i As Long
    Dim mTemp As Long
    Dim HotCell As POINTAPI
    Dim Cel As POINTAPI
    Dim bCancel As Boolean
    
'    If Button <> vbLeftButton Then
'        RaiseEvent MouseUp(Button, Shift, X, Y)
'        Exit Sub
'    End If
    
    SizeColumn.CurColumn = -1
    SizeColumn.CurRow = -1
    Timer1.Interval = 0
    
    
    If mColDrag.X > -1 Then
        mColDrag.X = -1
        If mColDrag.DestCol > -1 Then
            mTemp = PtrCol(mColDrag.SrcCol)
            If mColDrag.DestCol < mColDrag.SrcCol Then
                For i = mColDrag.SrcCol To mColDrag.DestCol + 1 Step -1
                    PtrCol(i) = PtrCol(i - 1)
                Next
            Else
                For i = mColDrag.SrcCol To mColDrag.DestCol - 1
                    PtrCol(i) = PtrCol(i + 1)
                Next
            End If
            PtrCol(mColDrag.DestCol) = mTemp
            If m_SelectionMode <> GP_SelByMultiRow And m_SelectionMode <> GP_SelBySingleRow Then
                SelCel.X = mColDrag.DestCol
                SelRange.Start.X = mColDrag.DestCol
                SelRange.End.X = mColDrag.DestCol
            End If
            
            '/>Jose Liza - FocusRect
            If m_FocusRectMode <> lgNone Then
               m_CellFocusRect.X = mColDrag.DestCol
               SelCel.X = mColDrag.DestCol
            End If
            '/<Jose Liza - FocusRect
            
            Draw
            RaiseEvent AfterColumnDrag(mColDrag.DestCol)
            GoTo EventMouseUp
        End If
    End If
    
    'Cell checkbox
    If PvGetHotCell(X, Y, HotCell, R) Then
        IsMouseInCellPart HotCell.Y, HotCell.X, X, Y, bCheckRow, bText, bImage
        
        If bCheckRow Then
            SetCursor hCurHands
            mRow(PtrRow(HotCell.Y)).Checked = Not mRow(PtrRow(HotCell.Y)).Checked
            m_AllRowAreCheked = True
            For i = 0 To m_RowsCount - 1
                If mRow(PtrRow(i)).Checked = False Then
                    m_AllRowAreCheked = False
                    Exit For
                End If
            Next
            Draw
            GoTo EventMouseUp
        End If
        
        If (bText And mCol(PtrCol(HotCell.X)).LabelsEvents) Or (bText And mCol(PtrCol(HotCell.X)).DataType = GP_BOOLEAN) Then
            
            If mCol(PtrCol(HotCell.X)).DataType = GP_BOOLEAN And Not mCol(PtrCol(HotCell.X)).EditionLocked And m_AllowEdit Then
                RaiseEvent BeforeEdit(HotCell.Y, HotCell.X, mRow(PtrRow(HotCell.Y)).Cells(PtrCol(HotCell.X)).Value, bCancel)
                If bCancel Then GoTo EventMouseUp
                
                If mRow(PtrRow(HotCell.Y)).Cells(PtrCol(HotCell.X)).EditionLocked = False Then
                    SetCursor hCurHands
                    With mRow(PtrRow(HotCell.Y)).Cells(PtrCol(HotCell.X))
                        vOldValue = .Value
                        If IsNull(.Value) Then
                            .Value = True
                        Else
                            .Value = Not .Value
                        End If
                    End With
                    
                    RaiseEvent AfterEdit(HotCell.Y, HotCell.X, vOldValue)
                    Draw
                End If
            Else
                SetCursor hCurHands
                RaiseEvent LabelMouseUp(HotCell.Y, HotCell.X, Button, Shift)
            End If
        End If
        
        If bImage And mCol(PtrCol(HotCell.X)).ImagesEvents Then
            SetCursor hCurHands
            RaiseEvent ImgMouseUp(HotCell.Y, HotCell.X, Button, Shift)
        End If
    End If
    
    If Y < m_HeaderHeight And m_HeaderHeight > 0 Then
        If IsHotColCheckBox(HotCol, X, Y) Then
            m_AllRowAreCheked = Not m_AllRowAreCheked
            For i = 0 To m_RowsCount - 1
                mRow(PtrRow(i)).Checked = m_AllRowAreCheked
            Next
            Draw
            GoTo EventMouseUp
        End If
    
         '// Sort requested from Column Header click
         If (HotCol <> C_NULL_RESULT) And (SizeColumn.HotColumn = C_NULL_RESULT) And Button = vbLeftButton Then
            If m_AllowColumnSort Then
               If (Shift And vbCtrlMask) And Not (mSortColumn = -1) Then
'                  If Not (mSortSubColumn = PtrCol(HotCol)) Then
'                     mCol(PtrCol(HotCol)).nSortOrder = lgSTNormal
'                  End If

                  'mSortSubColumn = HotCol
                  Sort mSortColumn, mCol(mSortColumn).nSortOrder, PtrCol(HotCol) ', mCol(PtrCol(mSortColumn)).nSortOrder

               Else
                  If Not (mSortColumn = PtrCol(HotCol)) Then
                     mCol(PtrCol(HotCol)).nSortOrder = lgSTNormal
                     mSortSubColumn = -1
                  End If

                  'mSortColumn = HotCol
                  If Not (mSortSubColumn = -1) Then
                     Sort , , , mCol(PtrCol(mSortSubColumn)).nSortOrder
                  Else
                     Sort PtrCol(HotCol)
                  End If
               End If
               
               '/>Jose Liza - FocusRect
                If m_FocusRectMode <> lgNone Then
                   m_CellFocusRect.X = HotCol
                   SelCel.X = HotCol
                End If
                '/<Jose Liza - FocusRect

               Draw
            End If
         End If
         If HotCol <> -1 Then
            RaiseEvent ColumnClick(HotCol, Button, Shift)
         End If
         
    End If
    
    'Click in Group Button
    If PvGetHotCell(X, Y, Cel, R) And eSelBy = SelectByCells Then
        RaiseEvent CellClick(Cel.Y, Cel.X, Button, Shift)
    End If

    
    If PvGetHotCell(m_RowSelectorWidth + 1, Y, Cel, R) Then
        If Button = vbLeftButton Then
            If mRow(PtrRow(Cel.Y)).IsGroup = True Then
                With R2
                    .Left = Margin + (20 * DpiF * mRow(PtrRow(Cel.Y)).Ident)
                    .Top = (R.Bottom - R.Top) / 2 - 8 * DpiF
                    .Right = .Left + 16 * DpiF
                    .Bottom = .Top + 16 * DpiF
                End With
                
                If PtInRect(R2, X - R.Left, Y - R.Top) Then
                    If mRow(PtrRow(Cel.Y)).IsGroupExpanded Then
                        GroupColapse Cel.Y
                    Else
                        GroupExpand Cel.Y
                    End If
                End If
            End If
        End If
    End If
    
EventMouseUp:
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Function IsHotColCheckBox(HotCol As Long, ByVal X As Long, ByVal Y As Long) As Boolean
    Dim i As Long, lLeft As Long, lTop As Long
    If HotCol = -1 Then Exit Function
    If m_CheckStyle And PtrCol(HotCol) = 0 Then
        
        For i = 0 To HotCol - 1
            lLeft = lLeft + mCol(PtrCol(i)).Width
        Next
        
        lLeft = lLeft + m_RowSelectorWidth + Margin
        lTop = m_HeaderHeight / 2 - 8 * DpiF

        If X > lLeft And X < lLeft + 16 * DpiF And Y > lTop And Y < lTop + 16 * DpiF Then
            IsHotColCheckBox = True
        End If
    End If
End Function

Public Sub GroupExpand(ByVal Row As Long)
    Dim i As Long
    mRow(PtrRow(Row)).IsGroupExpanded = True
    
    For i = Row + 1 To m_RowsCount - 1
        With mRow(PtrRow(i))
            If Not .IsGroup Then
                If .Height = 0 Then
                    .Height = .TempHeight
                    .TempHeight = 0
                    mHideRowsCount = mHideRowsCount - 1
                End If
            ElseIf .Ident > mRow(PtrRow(Row)).Ident Then
                If .Height = 0 Then
                    .Height = .TempHeight
                    .TempHeight = 0
                    mHideRowsCount = mHideRowsCount - 1
                End If
                If .IsGroupExpanded = False Then
                    Do
                        i = i + 1
                        If i > m_RowsCount - 1 Then Exit Do
                        If mRow(PtrRow(i)).IsGroup Then i = i - 1: Exit Do
                    Loop
                End If
            Else
                Exit For
            End If
        End With
    Next
    UserControl_Resize
End Sub

Public Sub GroupColapse(ByVal Row As Long)
    Dim i As Long
    
    mRow(PtrRow(Row)).IsGroupExpanded = False
    
    For i = Row + 1 To m_RowsCount - 1
        With mRow(PtrRow(i))
            If Not .IsGroup Then
                If .Height > 0 Then
                    .TempHeight = .Height
                    .Height = 0
                    mHideRowsCount = mHideRowsCount + 1
                End If
            ElseIf .Ident > mRow(PtrRow(Row)).Ident Then
                If .Height > 0 Then
                    .TempHeight = .Height
                    .Height = 0
                    mHideRowsCount = mHideRowsCount + 1
                End If
            Else
                Exit For
            End If
        End With
    Next
    UserControl_Resize
End Sub

Private Function GetColumnsRight()
    Dim Col As Long
    Dim Right As Long
    Right = m_RowSelectorWidth
    For Col = 0 To m_ColsCount - 1
        Right = Right + mCol(PtrCol(Col)).Width
    Next
    GetColumnsRight = Right
End Function


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim R As Rect, HotCell As POINTAPI
    Dim OldSelCel As POINTAPI
    
    Text1.Visible = False
    EmptyPoint mCellEdit
    OldSelCel = SelCel

    If (Button <> vbLeftButton And Button <> vbRightButton) Then
        RaiseEvent MouseDown(Button, Shift, X, Y)
        Exit Sub
    End If
    
    'CORNER LEFT
    If Y < m_HeaderHeight And X < m_RowSelectorWidth Then
        eSelBy = SelectBtCornerLeftTop
        SelCel.X = 0: SelCel.Y = 0
        If m_SelectionMode = GP_SelByMultiRow Or m_SelectionMode = GP_SelFree Then
            With SelRange
                .Start.X = 0
                .Start.Y = 0
                .End.X = m_ColsCount - 1
                .End.Y = m_RowsCount - 1
            End With
        End If
        Draw
        RaiseEvent CurCellChange(SelCel.Y, SelCel.X, OldSelCel.Y, OldSelCel.X)
        GoTo EventMouseDown
    End If
    
    'COLUMN HEADER
    If Y < m_HeaderHeight And X < GetColumnsRight Then
        If IsHotColCheckBox(HotCol, X, Y) Then
            eSelBy = SelectByNone
            GoTo EventMouseDown
        End If
    
        If SizeColumn.HotColumn <> -1 Then
            SizeColumn.CurColumn = SizeColumn.HotColumn
            eSelBy = SelectByNone
            GoTo EventMouseDown
        End If

        eSelBy = SelectByColumns
        If m_SelectionMode = GP_SelByCol Or m_SelectionMode = GP_SelFree Then
            
            SelCel.X = HotCol: SelCel.Y = 0

            With SelRange
                .Start.X = HotCol
                .Start.Y = 0
                .End.X = HotCol
                .End.Y = m_RowsCount - 1
            End With
            Draw
            RaiseEvent CurCellChange(SelCel.Y, SelCel.X, OldSelCel.Y, OldSelCel.X)
        '/>Jose Liza - FocusRect
        Else
            If Me.FocusRectMode <> lgNone Then
                m_CellFocusRect.Y = 0
                m_CellFocusRect.X = HotCol
            End If
        '/<Jose Liza - FocusRect
        End If
        GoTo EventMouseDown
    End If

    'ROW SELECTOR
    If X < m_RowSelectorWidth Then
        If SizeColumn.HotRow <> -1 Then
            SizeColumn.CurRow = SizeColumn.HotRow
            eSelBy = SelectByNone
            GoTo EventMouseDown
        End If
        
        eSelBy = SelectByRow

        If m_SelectionMode = GP_SelByMultiRow Or m_SelectionMode = GP_SelFree Or m_SelectionMode = GP_SelBySingleRow Or m_SelectionMode = GP_SelBySingleCell Then
            
            SelCel.X = 0: SelCel.Y = HotRow
            
            '/>Jose Liza - FocusRect
            If Me.FocusRectMode <> lgNone Then
                m_CellFocusRect.Y = HotRow
                m_CellFocusRect.X = 0
            End If
            '/<Jose Liza - FocusRect
           
            If m_SelectionMode <> GP_SelBySingleCell Then
                 With SelRange
                     .Start.X = 0
                     .Start.Y = HotRow
                     .End.X = m_ColsCount - 1
                     .End.Y = HotRow
                 End With
             End If
             Draw

             RaiseEvent CurCellChange(SelCel.Y, SelCel.X, OldSelCel.Y, OldSelCel.X)
             GoTo EventMouseDown
        End If

    End If
    
    'CELLS
    If PvGetHotCell(X, Y, HotCell, R) Then
        Dim bCheckRow As Boolean, bText As Boolean, bImage As Boolean

        IsMouseInCellPart HotCell.Y, HotCell.X, X, Y, bCheckRow, bText, bImage
        
        '/>Jose Liza - FocusRect
        Call UpdateFocusRect(HotCell, All)
        '/<Jose Liza - FocusRect
         
        If bCheckRow Then
            SetCursor hCurHands
            GoTo EventMouseDown
        End If

        If (bText And mCol(PtrCol(HotCell.X)).LabelsEvents) Or (bText And mCol(PtrCol(HotCell.X)).DataType = GP_BOOLEAN And m_AllowEdit) Then
            SetCursor hCurHands
            RaiseEvent LabelMouseDown(HotCell.Y, HotCell.X, Button, Shift)
            GoTo EventMouseDown
        End If

        If bImage And mCol(PtrCol(HotCell.X)).ImagesEvents Then
            SetCursor hCurHands
            RaiseEvent ImgMouseDown(HotCell.Y, HotCell.X, Button, Shift)
            GoTo EventMouseDown
        End If
        
        If Button = vbRightButton Then
            If HotCell.X >= SelRange.Start.X And HotCell.X <= SelRange.End.X And _
                HotCell.Y >= SelRange.Start.Y And HotCell.Y <= SelRange.End.Y Then
                    GoTo EventMouseDown
            End If
        End If

        SelCel = HotCell
        eSelBy = SelectByCells
        
        Select Case m_SelectionMode
            Case GP_SelFree
                'Free Selection
                SelRange.Start = SelCel
                SelRange.End = SelCel
            Case GP_SelByMultiRow, GP_SelBySingleRow
                'Sel by Row
                With SelRange
                    .Start.Y = SelCel.Y
                    .Start.X = 0
                    .End.Y = SelCel.Y
                    .End.X = m_ColsCount - 1
                End With
            Case GP_SelByCol
                'Sel by Col
                With SelRange
                    .Start.Y = 0
                    .Start.X = SelCel.X
                    .End.Y = m_RowsCount - 1
                    .End.X = SelCel.X
                End With
            Case GP_SelBySingleCell
                EmptyPoint SelRange.Start
                EmptyPoint SelRange.End
        End Select

    
        If m_LastRowIsFooter = False Or SelCel.Y < m_RowsCount - 1 Then
            EnsureCellVisible SelCel.Y, SelCel.X
        End If

        Draw
   
        RaiseEvent CurCellChange(SelCel.Y, SelCel.X, OldSelCel.Y, OldSelCel.X)
    Else
        'ver si esto es necesario
        EmptyPoint SelRange.Start
        EmptyPoint SelRange.End
        EmptyPoint SelCel
        '--
        eSelBy = SelectByNone
        Draw

    End If
    
EventMouseDown:


    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Function IsPointEmpty(PT As POINTAPI) As Boolean
    If PT.X = -1 And PT.Y = -1 Then IsPointEmpty = True
End Function

Private Sub EmptyPoint(PT As POINTAPI)
    PT.X = -1: PT.Y = -1
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long, lLeft As Long, lTop As Long, bUpdate As Boolean
    Dim HotCell As POINTAPI
    Dim R As Rect
    Dim mBootom As Long
    Dim bCancel As Boolean
    
    HotCol = -1
    
    mBootom = IIF(ucScrollbarH.Visible, ucScrollbarH.Top, UserControl.ScaleHeight)
    
    mColDrag.Left = X
    
   If Button = vbLeftButton Then
   
        'Resize Column
        If SizeColumn.CurColumn <> -1 Then
            
            With mCol(PtrCol(SizeColumn.CurColumn))
                .Width = X - SizeColumn.Left
                If .Width < Margin Then .Width = Margin
                If .MinWidth > 0 And .Width < .MinWidth Then .Width = .MinWidth
            End With
            UserControl_Resize
            RaiseEvent ColumnUserResize(SizeColumn.CurColumn)
            GoTo EventMouseMove
        End If
        
        'Resize Row
        If m_AllowRowsResize Then
            If SizeColumn.CurRow <> -1 Then
                With mRow(PtrRow(SizeColumn.CurRow))
                    .Height = Y - SizeColumn.Left
                    If .Height < Margin Then .Height = Margin
                End With
                UserControl_Resize
                RaiseEvent RowUserResize(SizeColumn.CurRow)
                GoTo EventMouseMove
            End If
        End If
    End If
    
    'Select Out cell Range
    If (Button = vbLeftButton Or Button = vbRightButton) And _
        (eSelBy = SelectByCells Or eSelBy = SelectByColumns Or SelectByRow) Then

        If Y > mBootom And (eSelBy = SelectByCells Or eSelBy = SelectByRow) And Not ((m_SelectionMode = GP_SelBySingleCell) Or (m_SelectionMode = GP_SelBySingleRow)) Then
            If SelRange.End.Y < m_RowsCount - 1 Then
                SelRange.End.Y = SelRange.End.Y + 1
                EnsureCellVisible SelRange.End.Y, SelRange.End.X
                bUpdate = True
            End If
        End If
        
        If X > ucScrollbarV.Left Then
            If mColDrag.X = -1 Then
                If SelRange.End.X < m_ColsCount Then
                    SelRange.End.X = SelRange.End.X + 1
                    EnsureCellVisible SelRange.End.Y, SelRange.End.X
                    If SelRange.End.X > m_ColsCount - 1 Then SelRange.End.X = m_ColsCount - 1 'sino no muestra el ultimo
                    bUpdate = True
                End If
            Else    'drag column move scroll
                If ucScrollbarH.Value < ucScrollbarH.Max Then
                    ucScrollbarH.Value = ucScrollbarH.Value + m_ColsWidth / 10
                    Timer1.Interval = 40
                End If
            End If
           
        End If
        If Y < m_HeaderHeight And (eSelBy = SelectByCells Or eSelBy = SelectByRow) And Not ((m_SelectionMode = GP_SelBySingleCell) Or (m_SelectionMode = GP_SelBySingleRow)) Then
            If SelRange.Start.Y > 0 Then
                SelRange.Start.Y = SelRange.Start.Y - 1
                EnsureCellVisible SelRange.Start.Y, SelRange.Start.X
                bUpdate = True
            End If
        End If
        If X < m_RowSelectorWidth And eSelBy <> SelectByRow Then
            If SelRange.Start.X > -1 Then
                If mColDrag.X = -1 Or m_AllowColumnDrag = False Then
                    SelRange.Start.X = SelRange.Start.X - 1
                    EnsureCellVisible SelRange.Start.Y, SelRange.Start.X
                    If SelRange.Start.X = -1 Then SelRange.Start.X = 0 'sino no seleciona el 0
                    bUpdate = True
                Else  'drag column move scroll
                    If ucScrollbarH.Value > 0 Then
                        ucScrollbarH.Value = ucScrollbarH.Value - m_ColsWidth / 10
                        Timer1.Interval = 40
                    End If
                End If
            End If
        End If
        
        If bUpdate Then
            Draw
            Timer1.Interval = 30
        End If
    End If
    
    'HotCell
    If Y >= m_HeaderHeight And X >= m_RowSelectorWidth Then
        If (UserControl.MousePointer = vbSizeWE Or UserControl.MousePointer = vbSizeNS) Then
            UserControl.MousePointer = vbDefault
        End If

        
        If mColDrag.X = -1 And PvGetHotCell(X, Y, HotCell, R) Then
            Dim bCheckRow As Boolean, bText As Boolean, bImage As Boolean
            
            SizeColumn.HotColumn = -1
            SizeColumn.HotRow = -1
            If (Button = vbLeftButton Or Button = vbRightButton) Then
                If mHotCell.X <> HotCell.X Or mHotCell.Y <> HotCell.Y Then
                    mHotCell = HotCell
                    
                    RaiseEvent HotCellChange(HotCell.Y, HotCell.X)
                    IsMouseInCellPart HotCell.Y, HotCell.X, X, Y, bCheckRow, bText, bImage
                    
                    If Not (bText And mCol(PtrCol(HotCell.X)).LabelsEvents) And Not (bImage And mCol(PtrCol(HotCell.X)).ImagesEvents) Then
                        If eSelBy <> SelectByNone Then
                            Call PvSelect(HotCell.X, HotCell.Y)
                        End If
                    End If
                End If
            Else
                If mHotCell.X <> HotCell.X Or mHotCell.Y <> HotCell.Y Then
                    mHotCell = HotCell
                    HotRow = HotCell.Y
                    If m_ShowHotRow Then Draw
                    RaiseEvent HotCellChange(HotCell.Y, HotCell.X)
                End If
                IsMouseInCellPart HotCell.Y, HotCell.X, X, Y, bCheckRow, bText, bImage
                 
                If bCheckRow Then
                    SetCursor hCurHands
                End If
                
                If bText And (mCol(PtrCol(HotCell.X)).DataType = GP_BOOLEAN) And m_AllowEdit Then
                    SetCursor hCurHands
                    If IsPointEmpty(mHotPart) Then
                        mHotPart = HotCell
                        HotPartIsImage = False
                        Draw
                    End If
                    GoTo EventMouseMove
                End If
                
                If bText And mCol(PtrCol(HotCell.X)).LabelsEvents Then
                    SetCursor hCurHands
                    If IsPointEmpty(mHotPart) Then
                        mHotPart = HotCell
                        HotPartIsImage = False
                        Draw
                    End If
                    GoTo EventMouseMove
                End If
                
                If bImage And mCol(PtrCol(HotCell.X)).ImagesEvents Then
                    SetCursor hCurHands
                    If IsPointEmpty(mHotPart) Then
                        mHotPart = HotCell
                        HotPartIsImage = True
                        Draw
                    End If
                    GoTo EventMouseMove
                End If
                
                If Not IsPointEmpty(mHotPart) Then
                    EmptyPoint mHotPart
                    Draw
                End If
            End If
            GoTo EventMouseMove
        End If
    End If
    
    If mHotCell.X <> -1 Or mHotCell.Y <> -1 Then
        EmptyPoint mHotCell
        HotRow = -1
        If m_ShowHotRow Then Draw
        RaiseEvent HotCellChange(mHotCell.Y, mHotCell.X)
    End If
    
        
    'EmptyPoint mHotCell
    
   
    If Y < m_HeaderHeight Or mColDrag.X > -1 Then
        If m_ShowHotRow And HotRow <> -1 Then
            HotRow = -1
            Draw
        End If
        
        lLeft = -ucScrollbarH.Value + m_RowSelectorWidth
        
        For i = 0 To m_ColsCount - 1
            If X > lLeft And X <= lLeft + mCol(PtrCol(i)).Width Then
                HotCol = i
                If m_ShowHotColumn Then
                    If mHotCol <> HotCol And Button = 0 Then
                        mHotCol = HotCol
                        Draw
                    End If
                End If
                If Button = vbLeftButton And mColDrag.X = -1 Then
                    If eSelBy = SelectByColumns Then
                        If m_SelectionMode = GP_SelFree Then
                            Call PvSelect(HotCol, m_RowsCount - 1)
                        End If
                    End If
                End If
                
                'DRAG COLUMN
                If Button = vbLeftButton And eSelBy = SelectByColumns Then
                    If m_AllowColumnDrag Then
                        If mColDrag.X = -1 Then
                            bCancel = False
                            RaiseEvent BeforeColumnDrag(HotCol, bCancel)
                            If bCancel = False Then
                                mColDrag.SrcCol = HotCol
                                mColDrag.X = X - lLeft
                                mColDrag.DestCol = -1
                                
                            Else
                                eSelBy = SelectByNone
                            End If
                        Else
                            If X - lLeft <> mColDrag.X Then
                                bCancel = False
                                RaiseEvent OnColumnDrag(mColDrag.SrcCol, HotCol, bCancel)
                                If bCancel = False Then
                                    mColDrag.DestCol = HotCol
                                    mHotCol = HotCol
                                End If
                                Draw
                            End If
                        End If
                    End If
                End If
            End If
    
            lLeft = lLeft + mCol(PtrCol(i)).Width
            If X < lLeft + Margin And X > lLeft - Margin Then
                If mCol(PtrCol(i)).SizeLocked = False Then
                    UserControl.MousePointer = vbSizeWE
                    SizeColumn.HotColumn = i
                    SizeColumn.Left = lLeft - mCol(PtrCol(i)).Width
                End If
                Exit For
            Else
                UserControl.MousePointer = vbDefault
                If Button <> vbLeftButton Then SizeColumn.HotColumn = -1
                If lLeft > X Then Exit For
            End If
        Next
    Else
        If UserControl.MousePointer = vbSizeWE Then UserControl.MousePointer = vbDefault
    End If
    
    'RowSelector
    If X < m_RowSelectorWidth Then
        lTop = m_HeaderHeight
        
        For i = VScrollPos To m_RowsCount - 1
            If Y > lTop And Y <= lTop + mRow(PtrRow(i)).Height Then
                HotRow = i
                If Button = vbLeftButton Then
                    If eSelBy = SelectByRow Then
                        If mColDrag.X = -1 Then
                            Call PvSelect(m_ColsCount - 1, HotRow)
                        End If
                    End If
                Else
                    If m_ShowHotRow Then Draw
                End If
            End If
            
            
            
            lTop = lTop + mRow(PtrRow(i)).Height
                            
           
            If Y < lTop + Margin And Y > lTop - Margin Then
                If m_AllowRowsResize Then
                    UserControl.MousePointer = vbSizeNS
                    SizeColumn.HotRow = i
                    SizeColumn.Left = lTop - mRow(PtrRow(i)).Height
                    Exit For
                End If
            Else
                UserControl.MousePointer = vbDefault
                If Button <> vbLeftButton Then SizeColumn.HotRow = -1
                If lTop > Y Then Exit For
            End If
        Next
    Else
        If UserControl.MousePointer = vbSizeNS Then UserControl.MousePointer = vbDefault
    End If
        
EventMouseMove:
    If m_ShowHotColumn Then
        If HotCol = -1 And mHotCol <> -1 Then
            mHotCol = -1
            Draw
        End If
    End If
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub


Private Function IsMouseInCellPart(ByVal Row As Long, ByVal Col As Long, _
                                    ByVal X As Long, ByVal Y As Long, _
                                    bCheck As Boolean, bText As Boolean, bImage As Boolean)
                                    

   
    Dim Flags As Long
    Dim Align As eGridAlign
    Dim oFont As iFont
    Dim RCheck As Rect, RText As Rect, RImage As Rect, R As Rect
    Dim IsFirstColumn As Boolean

    PvGetCellRect Row, Col, R
    
    Row = PtrRow(Row)
    Col = PtrCol(Col)
    IsFirstColumn = Col = 0
    Set oFont = GetCellFont(Row, Col)
    Align = GetCellAlign(Row, Col)
    Flags = GetCellWordBreak(Row, Col) Or DT_NOPREFIX Or DT_EXPANDTABS
    
    Call SelectObject(UserControl.hdc, oFont.hFont)
    GetCellRectsParts Row, Col, IsFirstColumn, Flags, Align, R, RCheck, RText, RImage
    
    bCheck = PtInRect(RCheck, X, Y)
    bText = PtInRect(RText, X, Y)
    bImage = PtInRect(RImage, X, Y)
End Function

Public Function GetCellRect(ByVal Row As Long, ByVal Col As Long, ByVal ptrRect As Long) As Boolean
    Dim R As Rect
    GetCellRect = PvGetCellRect(Row, Col, R, True)
    CopyMemory ByVal ptrRect, R, 16
End Function

Private Function PvGetCellRect(ByVal Y As Long, ByVal X As Long, Rect As Rect, Optional AllGrid As Boolean) As Boolean
    Dim lTop As Long, lCol As Long, lRow As Long, R As Rect
    Dim iPart As Long
    Dim lStart1  As Long, lEnd1 As Long
    Dim lStart2  As Long, lEnd2 As Long
    Dim i As Long
    Dim bIsFullRow As Boolean
    Dim FixedColWidth As Long, FixedRowHeight As Long
    
    
    For i = 0 To m_FixedColumns - 1
        FixedColWidth = FixedColWidth + mCol(PtrCol(i)).Width
    Next

    For i = 0 To m_FixedRows - 1
        FixedRowHeight = FixedRowHeight + mRow(PtrRow(i)).Height
    Next
    
    bIsFullRow = RowIsFullRow(Y) Or Me.RowIsGroup(Y)
    
    If m_LastRowIsFooter And Y = m_RowsCount - 1 Then
        If ucScrollbarH.Visible Then
            lTop = ucScrollbarH.Top - mRow(PtrRow(m_RowsCount - 1)).Height
        Else
            lTop = UserControl.ScaleHeight - mRow(PtrRow(m_RowsCount - 1)).Height
        End If
 
        If ucScrollbarV.Visible And ucScrollbarV.Value < ucScrollbarV.Max Then
         
            R.Right = -ucScrollbarH.Value + m_RowSelectorWidth
            lRow = m_RowsCount - 1
            
            If bIsFullRow Then
                SetRect Rect, m_RowSelectorWidth, lTop, m_RowSelectorWidth, lTop + mRow(PtrRow(lRow)).Height
                For lCol = 0 To m_ColsCount - 1
                    Rect.Right = Rect.Right + mCol(PtrCol(lCol)).Width
                Next
            Else
                For lCol = 0 To m_ColsCount - 1
                    SetRect R, R.Right, lTop, R.Right + mCol(PtrCol(lCol)).Width, lTop + mRow(PtrRow(lRow)).Height
                    If lCol = X Then
                        Rect = R
                        PvGetCellRect = True
                        Exit Function
                    End If
                    If AllGrid = False Then
                        If R.Left > ucScrollbarV.Left Then Exit For
                    End If
                Next
            End If
        End If
    End If

    If Y < m_FixedRows And X < m_FixedColumns Then
        iPart = 0
    Else
        If Y < m_FixedRows Then
            iPart = 1
        ElseIf X < m_FixedColumns Then
            iPart = 2
        Else
            iPart = 3
        End If
    End If

    For i = iPart To 3
        R.Top = 0: R.Left = 0
        
        
        If i = 0 Or i = 1 Then
            lStart1 = 0
            lEnd1 = m_FixedRows - 1
            lTop = m_HeaderHeight
        Else
            lStart1 = VScrollPos + m_FixedRows
            lEnd1 = m_RowsCount - 1
            lTop = m_HeaderHeight + FixedRowHeight
        End If
        
        For lRow = lStart1 To lEnd1
            If i = 0 Or i = 2 Then
                R.Right = m_RowSelectorWidth
                lStart2 = 0
                lEnd2 = m_FixedColumns - 1
            Else
                R.Right = -ucScrollbarH.Value + m_RowSelectorWidth
                lStart2 = 0
                lEnd2 = m_ColsCount - 1
            End If
            
            If bIsFullRow Then
                If lRow = Y Then
                    SetRect Rect, m_RowSelectorWidth, lTop, m_RowSelectorWidth, lTop + mRow(PtrRow(lRow)).Height
                    For lCol = lStart2 To lEnd2
                        Rect.Right = Rect.Right + mCol(PtrCol(lCol)).Width
                    Next
                    PvGetCellRect = True
                    Exit Function
                End If
            Else
                For lCol = lStart2 To lEnd2
                    SetRect R, R.Right, lTop, R.Right + mCol(PtrCol(lCol)).Width, lTop + mRow(PtrRow(lRow)).Height
                    If lCol = X And lRow = Y Then
                        Rect = R
                        PvGetCellRect = True
                        Exit Function
                    End If
                    If R.Left > ucScrollbarV.Left Then Exit For
                Next
            End If
            lTop = lTop + mRow(PtrRow(lRow)).Height
            If AllGrid = False Then
                If R.Top > ucScrollbarH.Top Then Exit For
            End If
        Next
    Next
End Function

Public Function GetHotCell(ByVal X As Long, ByVal Y As Long, Row As Long, Col As Long, ptrRect As Long) As Boolean
    Dim PT As POINTAPI, R As Rect
    GetHotCell = PvGetHotCell(X, Y, PT, R)
    Row = PT.Y
    Col = PT.X
    CopyMemory ByVal ptrRect, R, 16
End Function

Private Function PvGetHotCell(ByVal X As Long, ByVal Y As Long, PT As POINTAPI, Rect As Rect) As Boolean
    Dim lTop As Long, lCol As Long, lRow As Long, R As Rect
    Dim i As Long
    Dim iPart As Long
    Dim lStart1  As Long, lEnd1 As Long
    Dim lStart2  As Long, lEnd2 As Long
    Dim FixedColWidth As Long, FixedRowHeight As Long
    
    EmptyPoint PT
    
    For i = 0 To m_FixedColumns - 1
        FixedColWidth = FixedColWidth + mCol(PtrCol(i)).Width
    Next

    For i = 0 To m_FixedRows - 1
        FixedRowHeight = FixedRowHeight + mRow(PtrRow(i)).Height
    Next
    
    If m_LastRowIsFooter Then
        If ucScrollbarH.Visible Then
            lTop = ucScrollbarH.Top - mRow(PtrRow(m_RowsCount - 1)).Height
        Else
            lTop = UserControl.ScaleHeight - mRow(PtrRow(m_RowsCount - 1)).Height
        End If
        If Y > lTop Then
            If ucScrollbarV.Visible And ucScrollbarV.Value < ucScrollbarV.Max Then
                
                R.Right = -ucScrollbarH.Value + m_RowSelectorWidth
                lRow = m_RowsCount - 1
                For lCol = 0 To m_ColsCount - 1
                    SetRect R, R.Right, lTop, R.Right + mCol(PtrCol(lCol)).Width, lTop + mRow(PtrRow(lRow)).Height
                    If X > R.Left And X <= R.Right And Y > R.Top And Y <= R.Bottom Then
                        PT.X = lCol
                        PT.Y = lRow
                        Rect = R
                        PvGetHotCell = True
                        Exit Function
                    End If
                    If R.Left > ucScrollbarV.Left Then Exit For
                Next
                
                
            End If
        End If
    End If

    If m_FixedRows > 0 And m_FixedColumns > 0 Then
        iPart = 0
    Else
        If m_FixedRows > 0 Then
            iPart = 1
        ElseIf m_FixedColumns > 0 Then
            iPart = 2
        Else
            iPart = 3
        End If
    End If

    For i = iPart To 3
        R.Top = 0: R.Left = 0
        
        If i = 0 Or i = 1 Then
            lTop = m_HeaderHeight
            lStart1 = 0
            lEnd1 = m_FixedRows - 1
        Else
            lTop = m_HeaderHeight + FixedRowHeight
            lStart1 = VScrollPos + m_FixedRows
            lEnd1 = m_RowsCount - 1
        End If
        
        For lRow = lStart1 To lEnd1
            If i = 0 Or i = 2 Then
                R.Right = m_RowSelectorWidth
                lStart2 = 0
                lEnd2 = m_FixedColumns - 1
            Else
                R.Right = -ucScrollbarH.Value + m_RowSelectorWidth
                lStart2 = 0
                lEnd2 = m_ColsCount - 1
            End If
            
            For lCol = lStart2 To lEnd2
                SetRect R, R.Right, lTop, R.Right + mCol(PtrCol(lCol)).Width, lTop + mRow(PtrRow(lRow)).Height
                If X > R.Left And X <= R.Right And Y > R.Top And Y <= R.Bottom Then
                    PT.X = lCol
                    PT.Y = lRow
                    Rect = R
                    PvGetHotCell = True
                    Exit Function
                End If
    
                If R.Left > ucScrollbarV.Left Then Exit For
            Next
    
            lTop = lTop + mRow(PtrRow(lRow)).Height
            If R.Top > ucScrollbarH.Top Then Exit For
        Next
    Next
End Function

Private Sub PvSelect(lCol As Long, lRow As Long)
    Select Case m_SelectionMode
        Case GP_SelFree
            If lCol >= SelCel.X Then
                SelRange.Start.X = SelCel.X
                SelRange.End.X = lCol
            Else
                SelRange.End.X = SelCel.X
                SelRange.Start.X = lCol
            End If
            
            If lRow >= SelCel.Y Then
                SelRange.Start.Y = SelCel.Y
                SelRange.End.Y = lRow
            Else
                SelRange.End.Y = SelCel.Y
                SelRange.Start.Y = lRow
            End If
        Case GP_SelByMultiRow
            'Sel by Row
            With SelRange
                If lRow >= SelCel.Y Then
                    .Start.Y = SelCel.Y
                    .End.Y = lRow
                Else
                    .Start.Y = lRow
                    .End.Y = SelCel.Y
                End If
                .Start.X = 0
                .End.X = m_ColsCount - 1
            End With
        Case GP_SelByCol
            'Sel By Column
            With SelRange
                If lCol >= SelCel.X Then
                    .Start.X = SelCel.X
                    .End.X = lCol
                Else
                    .Start.X = lCol
                    .End.X = SelCel.X
                End If
                .Start.Y = 0
                .End.Y = m_RowsCount - 1
            End With
    End Select
    Draw
End Sub


Private Sub UserControl_Resize()
    Dim lMax As Long, lHeight As Long, i As Long, ListHeight As Long
    Dim SH As Long, SV As Long, RowsInRectArea As Long, lTop As Long, lLeft As Long
    
    If m_ColsCount = 0 Then
        Draw 'border
        Exit Sub
    End If
    
    If m_ColumnsAutoFit = False Then
        lMax = m_RowSelectorWidth
        For i = 0 To m_ColsCount - 1
            lMax = lMax + mCol(i).Width
        Next
        
        If ucScrollbarV.Visible Then lMax = lMax + ucScrollbarV.Width
        If lMax > UserControl.ScaleWidth Then SH = ucScrollbarH.Height
        
    End If
    
    ListHeight = UserControl.ScaleHeight - m_HeaderHeight - SH
    
    For i = m_RowsCount - 1 To 0 Step -1
        lHeight = lHeight + mRow(PtrRow(i)).Height
        If lHeight <= ListHeight Then
            If mRow(i).TempHeight = 0 Then
                RowsInRectArea = RowsInRectArea + 1
            End If
        Else
            SV = ucScrollbarV.Width
            Exit For
        End If
    Next

    If RowsInRectArea = 0 Then
        If m_RowsCount = 0 Then
            Draw
            If m_ColumnsAutoFit Then AjustColumnsToClient
        End If
        Exit Sub
    End If

    With ucScrollbarH
        .Visible = SH
        If m_RowSelectorWidth = 0 Then
            lLeft = m_BorderWidth + (m_BorderRadius / 2) * DpiF
        Else
            lLeft = m_RowSelectorWidth
        End If
        If SV = 0 And m_BorderRadius > 0 Then SV = (m_BorderRadius / 2) * DpiF
        
        If lMax - UserControl.ScaleWidth + DpiF < .Value Then .Value = lMax - UserControl.ScaleWidth + DpiF
        .Max = lMax - UserControl.ScaleWidth + DpiF
        If lMax > 0 Then .LargeChange = .Max * UserControl.ScaleWidth / lMax
        .SmallChange = 10
        .Move lLeft, UserControl.ScaleHeight - .Height - m_BorderWidth, UserControl.ScaleWidth - SV - lLeft - m_BorderWidth, .Height
        
    End With

    With ucScrollbarV
        .Visible = lHeight >= ListHeight 'SV
        If m_HeaderHeight = 0 Then
            lTop = m_BorderWidth + (m_BorderRadius / 2) * DpiF
        Else
            lTop = m_HeaderHeight
        End If
        If SH = 0 And m_BorderRadius > 0 Then SH = (m_BorderRadius / 2) * DpiF
        lMax = (m_RowsCount - RowsInRectArea - mHideRowsCount) * 10
        If lMax < .Value Then .Value = lMax
        .Max = lMax
        If .Max > 0 Then
            .LargeChange = .Max * (ListHeight \ (ListHeight \ RowsInRectArea)) \ (m_RowsCount - mHideRowsCount)
            .WheelChange = 10 * 3
        End If
        .Move UserControl.ScaleWidth - .Width - m_BorderWidth, lTop, .Width, UserControl.ScaleHeight - SH - m_BorderWidth - lTop

    End With
    
    If m_ColumnsAutoFit Then AjustColumnsToClient
    ucScrollbarV_Change
End Sub

Private Sub AjustColumnsToClient()
    Dim Width As Long
    Dim TotalWidth
    Dim TempCols() As Single
    Dim i As Long
    
    Width = IIF(ucScrollbarV.Visible, ucScrollbarV.Left, UserControl.ScaleWidth) - m_RowSelectorWidth

    ReDim TempCols(m_ColsCount - 1)
        
    For i = 0 To m_ColsCount - 1
        If mCol(PtrCol(i)).SizeLocked Then
            Width = Width - mCol(i).Width
        Else
            TotalWidth = TotalWidth + mCol(i).Width
        End If
    Next
    
    For i = 0 To m_ColsCount - 1
        If mCol(PtrCol(i)).SizeLocked = False Then
            TempCols(i) = mCol(i).Width * 100 / TotalWidth
        End If
    Next
    
    For i = 0 To m_ColsCount - 1
        If mCol(PtrCol(i)).SizeLocked = False Then
            If mCol(i).TempWidth = 0 Then '<--ColHiden
                mCol(i).Width = Width * TempCols(i) / 100
                If mCol(i).Width < mCol(i).MinWidth Then mCol(i).Width = mCol(i).MinWidth
                If mCol(i).Width < Margin Then mCol(i).Width = Margin
            End If
        End If
    Next
End Sub

Private Sub FillRectangle(hdc As Long, Rect As Rect, ByVal Color As OLE_COLOR)
    Dim hBrush As Long
    If (Color And &H80000000) Then Color = GetSysColor(Color And &HFF&)
    hBrush = CreateSolidBrush(Color)
    FillRect hdc, Rect, hBrush
    Call DeleteObject(hBrush)
End Sub

Private Function FillGradientRect(hdc As Long, R As Rect, Color1 As Long, Color2 As Long, Vertical As Boolean)
    Dim TRIVERTEX(0 To 1) As TRIVERTEX
    Dim GRADIENT_RECT As GRADIENT_RECT
    
    If (Color1 And &H80000000) Then Color1 = GetSysColor(Color1 And &HFF&)
    If (Color2 And &H80000000) Then Color2 = GetSysColor(Color2 And &HFF&)
    
    With GRADIENT_RECT
        .UpperLeft = 0
        .LowerRight = 1
    End With
    
    With TRIVERTEX(0)
        .PxX = R.Left
        .PxY = R.Top
        .Red = Color1 And &HFF&
        .Green = (Color1 And &HFF00&) \ &H100&
        .Blue = Color1 \ &H10000
    End With
    
    With TRIVERTEX(1)
        .PxX = R.Right
        .PxY = R.Bottom
        .Red = Color2 And &HFF&
        .Green = (Color2 And &HFF00&) \ &H100&
        .Blue = Color2 \ &H10000
    End With
    
    GradientFill hdc, TRIVERTEX(0), 2, GRADIENT_RECT, 1, IIF(Vertical, &H1, &H0)
End Function

Private Sub DrawLine2(hdc As Long, X As Long, Y As Long, X2 As Long, Y2 As Long, Color As Long)
    Dim R      As Rect
    Dim hBrush    As Long
 
    If (Color And &H80000000) Then Color = GetSysColor(Color And &HFF&)
    
    hBrush = CreateSolidBrush(Color)
    With R
        .Left = X
        .Top = Y
        .Right = X2
        .Bottom = Y2
    End With
    FillRect hdc, R, hBrush
    Call DeleteObject(hBrush)
End Sub

Private Sub DrawLine(lpDC As Long, X As Long, Y As Long, X2 As Long, Y2 As Long, PenStyle As Long, Color As Long, ByVal Width As Long)
    Dim PT      As POINTAPI
    Dim hPen    As Long
    Dim hPenOld As Long
    If (Color And &H80000000) Then Color = GetSysColor(Color And &HFF&)
    Dim logBR  As LOGBRUSH

    With logBR
        .lbColor = Color
        .lbStyle = 0
        .lbHatch = 0&
    End With
    
    hPen = ExtCreatePen(PenStyle, Width, logBR, 0, ByVal 0&)
    hPenOld = SelectObject(lpDC, hPen)
    Call MoveToEx(lpDC, X, Y, PT)
    Call LineTo(lpDC, X2, Y2)
    Call SelectObject(lpDC, hPenOld)
    Call DeleteObject(hPen)
End Sub

Public Function GetWindowsDPI() As Double
    Dim hdc As Long, LPX  As Double
    hdc = GetDC(0)
    LPX = CDbl(GetDeviceCaps(hdc, LOGPIXELSX))
    ReleaseDC 0, hdc

    If (LPX = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = LPX / 96#
    End If
End Function

Private Function IsDarkColor(ByVal Color As Long) As Boolean
    Dim BGRA(0 To 3) As Byte
    If (Color And &H80000000) Then Color = GetSysColor(Color And &HFF&)
    CopyMemory BGRA(0), Color, 4&
    IsDarkColor = ((CLng(BGRA(0)) + (CLng(BGRA(1) * 3)) + CLng(BGRA(2))) / 2) < 382
End Function

'Funcion para combinar dos colores
Private Function ShiftColor(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
  
    Dim clrFore(3)         As Byte
    Dim clrBack(3)         As Byte
    
    If (clrFirst And &H80000000) Then clrFirst = GetSysColor(clrFirst And &HFF&)
    If (clrSecond And &H80000000) Then clrSecond = GetSysColor(clrSecond And &HFF&)
  
    CopyMemory clrFore(0), clrFirst, 4
    CopyMemory clrBack(0), clrSecond, 4
  
    clrFore(0) = (clrFore(0) * lAlpha + clrBack(0) * (255 - lAlpha)) / 255
    clrFore(1) = (clrFore(1) * lAlpha + clrBack(1) * (255 - lAlpha)) / 255
    clrFore(2) = (clrFore(2) * lAlpha + clrBack(2) * (255 - lAlpha)) / 255
  
    CopyMemory ShiftColor, clrFore(0), 4
  
End Function

'================================ROWS=============================
'#################################################################
'=================================================================
Public Property Get RowsCount() As Long
    RowsCount = m_RowsCount
End Property

Public Property Let RowsCount(ByVal New_Value As Long)
    Dim PrevCount As Long, lRow As Long, lCol As Long
    
    If New_Value <= 0 Then
        Clear
        Exit Property
    End If
    
    PrevCount = m_RowsCount
    m_RowsCount = New_Value
    
    ReDim Preserve mRow(m_RowsCount - 1)
    ReDim Preserve PtrRow(m_RowsCount - 1)
    
    If m_ColsCount > 0 Then
    
        For lRow = PrevCount To m_RowsCount - 1
            PtrRow(lRow) = lRow
            With mRow(lRow)
                .Height = m_RowsHeight
                .BackColor = CLR_NONE
                .ForeColor = CLR_NONE
                .Align = CenterLeft
            
                ReDim Preserve .Cells(m_ColsCount - 1)
                For lCol = 0 To m_ColsCount - 1
                    With mCol(lCol)
                        If .Width = 0 And .TempWidth = 0 Then
                            .Width = m_ColsWidth
                            .BackColor = CLR_NONE
                            .ForeColor = CLR_NONE
                            .HeaderForeColor = CLR_NONE
                            .Align = CenterLeft
                        End If
                    End With
                
                    With .Cells(lCol)
                        .BackColor = CLR_NONE
                        .ForeColor = CLR_NONE
                        .Align = CenterLeft
                    End With
                Next
            End With
        Next
    End If
    
    PropertyChanged "RowsCount"
    UserControl_Resize
End Property

Public Property Get RowRef(ByVal Row As Long) As Long
    If Row = -1 Then
        RowRef = -1
    Else
        RowRef = PtrRow(Row)
    End If
End Property

Public Property Get RowHeight(ByVal Row As Long) As Long
    RowHeight = mRow(PtrRow(Row)).Height / DpiF
End Property

Public Property Let RowHeight(ByVal Row As Long, ByVal New_Value As Long)
    mRow(PtrRow(Row)).Height = New_Value * DpiF
    If m_Redraw Then Refresh
End Property

Public Property Get RowsHeight() As Long
    RowsHeight = m_RowsHeight / DpiF
End Property

Public Property Let RowsHeight(ByVal New_Value As Long)
    m_RowsHeight = New_Value * DpiF
    PropertyChanged "RowsHeight"
    If m_Redraw Then Refresh
End Property

Public Property Get RowSelectorWidth() As Long
    RowSelectorWidth = m_RowSelectorWidth
End Property

Public Property Let RowSelectorWidth(ByVal New_Value As Long)
    m_RowSelectorWidth = New_Value
    If m_RowSelectorWidth < 0 Then m_RowSelectorWidth = 0
    PropertyChanged "RowSelectorWidth"
    If m_Redraw Then Refresh
End Property


Public Property Get RowHidden(ByVal Row As Long) As Boolean
    RowHidden = mRow(PtrRow(Row)).TempHeight > 0
End Property

Public Property Let RowHidden(ByVal Row As Long, ByVal New_Value As Boolean)
    Row = PtrRow(Row)
    If New_Value Then
        If mRow(Row).TempHeight = 0 Then
            mRow(Row).TempHeight = mRow(Row).Height
            mRow(Row).Height = 0
            mHideRowsCount = mHideRowsCount + 1
        End If
    Else
        If mRow(Row).TempHeight > 0 Then
            mRow(Row).Height = mRow(Row).TempHeight
            mRow(Row).TempHeight = 0
            mHideRowsCount = mHideRowsCount - 1
        End If
    End If
    If m_Redraw Then Refresh
End Property

Public Property Get RowBackColor(ByVal Row As Long) As Long
    RowBackColor = mRow(PtrRow(Row)).BackColor
End Property

Public Property Let RowBackColor(ByVal Row As Long, ByVal New_Value As Long)
    mRow(PtrRow(Row)).BackColor = New_Value
    Draw
End Property

Public Property Get RowForeColor(ByVal Row As Long) As Long
    RowForeColor = mRow(PtrRow(Row)).ForeColor
End Property

Public Property Let RowForeColor(ByVal Row As Long, ByVal New_Value As Long)
    mRow(PtrRow(Row)).ForeColor = New_Value
    Draw
End Property

Public Property Get RowFont(ByVal Row As Long) As Font
    Row = PtrRow(Row)
    If mRow(Row).Font Is Nothing Then
        Set mRow(Row).Font = CloneFont(m_Font)
    End If
    Set RowFont = mRow(Row).Font
End Property

Public Property Let RowFont(ByVal Row As Long, ByVal New_Value As Font)
    Set mRow(PtrRow(Row)).Font = New_Value
    Draw
End Property

Public Property Get RowAlign(ByVal Row As Long) As eGridAlign
    RowAlign = mRow(PtrRow(Row)).Align
End Property

Public Property Let RowAlign(ByVal Row As Long, ByVal New_Value As eGridAlign)
    mRow(PtrRow(Row)).Align = New_Value
    Draw
End Property

Public Property Get RowWordBreak(ByVal Row As Long) As Boolean
    RowWordBreak = mRow(PtrRow(Row)).WordBreak
End Property

Public Property Let RowWordBreak(ByVal Row As Long, ByVal New_Value As Boolean)
    mRow(PtrRow(Row)).WordBreak = New_Value
    Draw
End Property

Public Property Get RowTag(ByVal Row As Long) As Variant
    RowTag = mRow(PtrRow(Row)).Tag
End Property

Public Property Let RowTag(ByVal Row As Long, ByVal Value As Variant)
     mRow(PtrRow(Row)).Tag = Value
End Property

Public Property Let RowIsGroup(ByVal Row As Long, ByVal New_Value As Boolean)
    If New_Value Then
        mRow(PtrRow(Row)).IsGroupExpanded = True
    Else
        If mRow(PtrRow(Row)).IsGroup Then
            If mRow(PtrRow(Row)).IsGroupExpanded = False Then
                GroupExpand Row
            End If
        End If
    End If
    mRow(PtrRow(Row)).IsGroup = New_Value
    
    Draw
End Property

Public Property Get RowIsGroup(ByVal Row As Long) As Boolean
    RowIsGroup = mRow(PtrRow(Row)).IsGroup
End Property

Public Property Let RowIsFullRow(ByVal Row As Long, ByVal New_Value As Boolean)
    mRow(PtrRow(Row)).IsFullRow = New_Value
    Draw
End Property

Public Property Get RowIsFullRow(ByVal Row As Long) As Boolean
    RowIsFullRow = mRow(PtrRow(Row)).IsFullRow
End Property

Public Property Get RowChecked(ByVal Row As Long) As Boolean
    RowChecked = mRow(PtrRow(Row)).Checked
End Property

Public Property Let RowChecked(ByVal Row As Long, ByVal New_Value As Boolean)
    Dim i  As Long
    mRow(PtrRow(Row)).Checked = New_Value
    m_AllRowAreCheked = True
    For i = 0 To m_RowsCount - 1
        If mRow(PtrRow(i)).Checked = False Then
            m_AllRowAreCheked = False
        End If
    Next
    Draw
End Property


Public Function RowDelete(ByVal Row As Long)
    Dim i As Long
    Dim ptr As Long
    
    ptr = PtrRow(Row)
 
    If RowHidden(Row) Then mHideRowsCount = mHideRowsCount - 1
    
    For i = ptr To m_RowsCount - 2
        mRow(i) = mRow(i + 1)
    Next
    
    For i = Row To m_RowsCount - 2
        PtrRow(i) = PtrRow(i + 1)
    Next

    For i = 0 To m_RowsCount - 1
       If PtrRow(i) > ptr Then
          PtrRow(i) = PtrRow(i) - 1
       End If
    Next

    m_RowsCount = m_RowsCount - 1
    If m_RowsCount = 0 Then
        Clear
    Else
        ReDim Preserve mRow(m_RowsCount - 1)
        ReDim Preserve PtrRow(m_RowsCount - 1)
        If m_Redraw Then UserControl_Resize
    End If
    
End Function

Public Property Get RowIdent(ByVal Row As Long) As Long
    RowIdent = mRow(PtrRow(Row)).Ident
End Property

Public Property Let RowIdent(ByVal Row As Long, ByVal Value As Long)
     mRow(PtrRow(Row)).Ident = Value
End Property

Public Function Clear(Optional ClearAllCols As Boolean)
    If ClearAllCols Then
        m_ColsCount = 0
        ReDim mCols(0)
    End If
    m_RowsCount = 0
    ReDim mRow(0)
    mColDrag.X = -1
    mSortColumn = -1
    mSortSubColumn = -1
    SizeColumn.CurColumn = -1
    SizeColumn.CurRow = -1
    HotRow = -1
    HotCol = -1
    EmptyPoint mHotCell
    EmptyPoint SelRange.Start
    EmptyPoint SelRange.End
    EmptyPoint SelCel
    EmptyPoint mCellEdit
    EmptyPoint mHotPart
    'm_Redraw = True
    ucScrollbarV.Visible = False
    'ucScrollbarH.Visible = False
    UserControl.Cls
    If m_Redraw Then Me.Refresh
    
End Function

Public Property Get CurRow() As Long
    CurRow = SelCel.Y
End Property

Public Property Let CurRow(ByVal NewValue As Long)
    SelCel.Y = NewValue
    Draw
End Property

'*-
'================================COLS=============================
'#################################################################
'=================================================================

Public Property Get ColsCount() As Long
    ColsCount = m_ColsCount
End Property

Public Property Let ColsCount(ByVal New_Value As Long)
    Dim PrevCount As Long, lRow As Long, lCol As Long
    
    If New_Value <= 0 Then Exit Property
    
    PrevCount = m_ColsCount
    m_ColsCount = New_Value
    
    ReDim Preserve mCol(m_ColsCount - 1)
    ReDim Preserve PtrCol(m_ColsCount - 1)
    If m_ColsCount > 0 Then
        For lRow = 0 To m_RowsCount - 1
            With mRow(lRow)
                .Height = m_RowsHeight
                .BackColor = CLR_NONE
                .ForeColor = CLR_NONE
                .Align = CenterLeft
                ReDim Preserve .Cells(m_ColsCount - 1)
                For lCol = PrevCount To m_ColsCount - 1
                    With .Cells(lCol)
                        .BackColor = CLR_NONE
                        .ForeColor = CLR_NONE
                        .Align = CenterLeft
                    End With
                Next
            End With
        Next
    End If
    
    For lCol = PrevCount To m_ColsCount - 1
        PtrCol(lCol) = lCol
        With mCol(lCol)
            .Width = m_ColsWidth
            .BackColor = CLR_NONE
            .ForeColor = CLR_NONE
            .HeaderForeColor = CLR_NONE
            .Align = CenterLeft
        End With
    Next
    PropertyChanged "ColsCount"
    UserControl_Resize
End Property

Public Property Get ColRef(ByVal Col As Long) As Long
    If Col = -1 Then
        ColRef = -1
    Else
        ColRef = PtrCol(Col)
    End If
End Property

Public Property Get ColHidden(ByVal Col As Long) As Boolean
    ColHidden = mCol(PtrCol(Col)).Width = 0
End Property


Public Property Let ColHidden(ByVal Col As Long, ByVal New_Value As Boolean)
    Col = PtrCol(Col)
    If New_Value Then
        mCol(Col).TempWidth = mCol(Col).Width
        mCol(Col).Width = 0
    Else
        If mCol(Col).TempWidth > 0 Then mCol(Col).Width = mCol(Col).TempWidth
        mCol(Col).TempWidth = 0
    End If
    If m_Redraw Then Refresh
End Property

Public Property Get ColDataType(ByVal Col As Long) As eDataType
    ColDataType = mCol(PtrCol(Col)).DataType
End Property

Public Property Let ColDataType(ByVal Col As Long, ByVal New_Value As eDataType)
    mCol(PtrCol(Col)).DataType = New_Value
    Draw
End Property

Public Property Get ColumnText(ByVal Col As Long) As String
    ColumnText = mCol(PtrCol(Col)).Text
End Property

Public Property Let ColumnText(ByVal Col As Long, ByVal Text As String)
    mCol(PtrCol(Col)).Text = Text
    Draw
End Property

Public Property Get ColsWidth() As Long
    ColsWidth = m_ColsWidth / DpiF
End Property

Public Property Let ColsWidth(ByVal New_Value As Long)
    m_ColsWidth = New_Value * DpiF
    PropertyChanged "ColsWidth"
    If m_Redraw Then Refresh
End Property

Public Property Get HeaderHeight() As Long
    HeaderHeight = m_HeaderHeight / DpiF
End Property

Public Property Let HeaderHeight(ByVal New_Value As Long)
    m_HeaderHeight = New_Value * DpiF
    PropertyChanged "HeaderHeight"
    If m_Redraw Then Refresh
End Property

Public Property Get ColBackColor(ByVal Col As Long) As Long
    ColBackColor = mCol(PtrCol(Col)).BackColor
End Property

Public Property Let ColBackColor(ByVal Col As Long, ByVal New_Value As Long)
    mCol(PtrCol(Col)).BackColor = New_Value
    Draw
End Property

Public Property Get ColForeColor(ByVal Col As Long) As Long
    ColForeColor = mCol(PtrCol(Col)).ForeColor
End Property

Public Property Let ColForeColor(ByVal Col As Long, ByVal New_Value As Long)
    mCol(PtrCol(Col)).ForeColor = New_Value
    Draw
End Property

Public Property Get ColFont(ByVal Col As Long) As Font
    Col = PtrCol(Col)
    If mCol(Col).Font Is Nothing Then
        Set mCol(Col).Font = CloneFont(m_Font)
    End If
    Set ColFont = mCol(Col).Font
End Property

Public Property Let ColFont(ByVal Col As Long, ByVal New_Value As Font)
    Set mCol(PtrCol(Col)).Font = New_Value
    Draw
End Property

Public Property Get ColTag(ByVal Col As Long) As Variant
    ColTag = mCol(PtrCol(Col)).Tag
End Property

Public Property Let ColTag(ByVal Col As Long, ByVal Value As Variant)
     mCol(PtrCol(Col)).Tag = Value
End Property

Public Property Get ColAlign(ByVal Col As Long) As eGridAlign
    ColAlign = mCol(PtrCol(Col)).Align
End Property

Public Property Let ColAlign(ByVal Col As Long, ByVal New_Value As eGridAlign)
    mCol(PtrCol(Col)).Align = New_Value
    Draw
End Property

Public Property Get ColImgAlign(ByVal Col As Long) As eGridAlign
    ColImgAlign = mCol(PtrCol(Col)).ImgAlign
End Property

Public Property Let ColImgAlign(ByVal Col As Long, ByVal New_Value As eGridAlign)
    mCol(PtrCol(Col)).ImgAlign = New_Value
    Draw
End Property

Public Property Get ColImgMonocrome(ByVal Col As Long) As Boolean
    ColImgMonocrome = mCol(PtrCol(Col)).ImagesMonocrome
End Property

Public Property Let ColImgMonocrome(ByVal Col As Long, ByVal New_Value As Boolean)
    mCol(PtrCol(Col)).ImagesMonocrome = New_Value
    Draw
End Property

Public Property Get ColWordBreak(ByVal Col As Long) As Boolean
    ColWordBreak = mCol(PtrCol(Col)).WordBreak
End Property

Public Property Let ColWordBreak(ByVal Col As Long, ByVal New_Value As Boolean)
    mCol(PtrCol(Col)).WordBreak = New_Value
    Draw
End Property

Public Property Get ColTextHidde(ByVal Col As Long) As Boolean
    ColTextHidde = mCol(PtrCol(Col)).TextHide
End Property

Public Property Let ColTextHidde(ByVal Col As Long, ByVal New_Value As Boolean)
    mCol(PtrCol(Col)).TextHide = New_Value
    Draw
End Property

Public Property Let ColMinWidth(ByVal Col As Long, ByVal New_Value As Long)
    mCol(PtrCol(Col)).MinWidth = New_Value * DpiF
    If m_Redraw Then Refresh
End Property

Public Property Get ColMinWidth(ByVal Col As Long) As Long
    ColMinWidth = mCol(PtrCol(Col)).MinWidth / DpiF
End Property

Public Property Let ColWidth(ByVal Col As Long, ByVal New_Value As Long)
    mCol(PtrCol(Col)).Width = New_Value * DpiF
    If m_Redraw Then Refresh
End Property

Public Property Get ColWidth(ByVal Col As Long) As Long
    ColWidth = mCol(PtrCol(Col)).Width / DpiF
End Property

Public Property Get ColUserResizeLocked(ByVal Col As Long) As Boolean
    ColUserResizeLocked = mCol(PtrCol(Col)).SizeLocked
End Property

Public Property Let ColUserResizeLocked(ByVal Col As Long, ByVal New_Value As Boolean)
    mCol(PtrCol(Col)).SizeLocked = New_Value
End Property

Public Property Get ColEditionLocked(ByVal Col As Long) As Boolean
    ColEditionLocked = mCol(PtrCol(Col)).EditionLocked
End Property

Public Property Let ColEditionLocked(ByVal Col As Long, ByVal New_Value As Boolean)
    mCol(PtrCol(Col)).EditionLocked = New_Value
End Property

Public Property Get ColLabelsEvents(ByVal Col As Long) As Boolean
    ColLabelsEvents = mCol(PtrCol(Col)).LabelsEvents
End Property

Public Property Let ColLabelsEvents(ByVal Col As Long, ByVal New_Value As Boolean)
    mCol(PtrCol(Col)).LabelsEvents = New_Value
End Property

Public Property Get ColHeaderImgIndex(ByVal Col As Long) As Integer
    ColHeaderImgIndex = mCol(PtrCol(Col)).IconIndex
End Property

Public Property Let ColHeaderImgIndex(ByVal Col As Long, ByVal New_Value As Integer)
    mCol(PtrCol(Col)).IconIndex = New_Value
    Draw
End Property

Public Property Get ColHeaderForeColor(ByVal Col As Long) As Long
    ColHeaderForeColor = mCol(PtrCol(Col)).HeaderForeColor
End Property

Public Property Let ColHeaderForeColor(ByVal Col As Long, ByVal New_Value As Long)
    mCol(PtrCol(Col)).HeaderForeColor = New_Value
    Draw
End Property


Public Property Get ColImagesEvents(ByVal Col As Long) As Boolean
    ColImagesEvents = mCol(PtrCol(Col)).ImagesEvents
End Property

Public Property Let ColImagesEvents(ByVal Col As Long, ByVal New_Value As Boolean)
    mCol(PtrCol(Col)).ImagesEvents = New_Value
End Property

Public Property Let ColFormat(ByVal Col As Long, ByVal New_Value As String)
     mCol(PtrCol(Col)).Format = New_Value
    Draw
End Property

Public Property Get ColFormat(ByVal Col As Long) As String
    ColFormat = mCol(PtrCol(Col)).Format
End Property

Public Property Get ColLeft(ByVal Col As Long) As Long
    Dim i As Long, lLeft As Long
    lLeft = m_RowSelectorWidth
    For i = 0 To Col - 1
        lLeft = lLeft + mCol(PtrCol(i)).Width
    Next
    ColLeft = lLeft
End Property

Public Property Get CurCol() As Long
    CurCol = SelCel.X
End Property

Public Property Let CurCol(ByVal NewValue As Long)
    SelCel.X = NewValue
    Draw
End Property

Public Property Get ColSortOrder(ByVal Col As Long) As lgSortTypeEnum
    ColSortOrder = mCol(PtrCol(Col)).nSortOrder
End Property

Public Property Let ColSortOrder(ByVal Col As Long, ByVal NewValue As lgSortTypeEnum)
    mCol(PtrCol(Col)).nSortOrder = NewValue
End Property

Public Property Get ColSort() As Long
    ColSort = mSortColumn
End Property

Public Property Let ColSort(ByVal Col As Long)
    mSortColumn = Col
    Draw
End Property

'================================CELLS=============================
'#################################################################
'=================================================================

Public Property Get CellValue(ByVal Row As Long, ByVal Col As Long) As Variant
    CellValue = mRow(PtrRow(Row)).Cells(PtrCol(Col)).Value
End Property

Public Property Let CellValue(ByVal Row As Long, ByVal Col As Long, ByVal Value As Variant)
     mRow(PtrRow(Row)).Cells(PtrCol(Col)).Value = Value
     Draw
End Property

Public Property Get CellTag(ByVal Row As Long, ByVal Col As Long) As Variant
    CellTag = mRow(PtrRow(Row)).Cells(PtrCol(Col)).Tag
End Property

Public Property Let CellTag(ByVal Row As Long, ByVal Col As Long, ByVal Value As Variant)
     mRow(PtrRow(Row)).Cells(PtrCol(Col)).Tag = Value
End Property

Public Property Get CellBackColor(ByVal Row As Long, ByVal Col As Long) As Long
    CellBackColor = mRow(PtrRow(Row)).Cells(PtrCol(Col)).BackColor
End Property

Public Property Let CellBackColor(ByVal Row As Long, ByVal Col As Long, ByVal New_Value As Long)
    mRow(PtrRow(Row)).Cells(PtrCol(Col)).BackColor = New_Value
    Draw
End Property

Public Property Get CellForeColor(ByVal Row As Long, ByVal Col As Long) As Long
    CellForeColor = mRow(PtrRow(Row)).Cells(Col).ForeColor
End Property

Public Property Let CellForeColor(ByVal Row As Long, ByVal Col As Long, ByVal New_Value As Long)
    mRow(PtrRow(Row)).Cells(PtrCol(Col)).ForeColor = New_Value
    Draw
End Property

Public Property Get CellFont(ByVal Row As Long, ByVal Col As Long) As StdFont
    Col = PtrCol(Col)
    Row = PtrRow(Row)
    If mRow(Row).Cells(Col).Font Is Nothing Then
        Set mRow(Row).Cells(Col).Font = CloneFont(m_Font)
    End If
    Set CellFont = mRow(Row).Cells(Col).Font
End Property

Public Property Let CellFont(ByVal Row As Long, ByVal Col As Long, ByVal New_Value As StdFont)
    Set mRow(PtrRow(Row)).Cells(PtrCol(Col)).Font = New_Value
    Draw
End Property

Public Property Get CellAlign(ByVal Row As Long, ByVal Col As Long) As eGridAlign
    CellAlign = mRow(PtrRow(Row)).Cells(PtrCol(Col)).Align
End Property

Public Property Let CellAlign(ByVal Row As Long, ByVal Col As Long, ByVal New_Value As eGridAlign)
    mRow(PtrRow(Row)).Cells(PtrCol(Col)).Align = New_Value
    Draw
End Property

Public Property Get CellWordBreak(ByVal Row As Long, ByVal Col As Long) As Boolean
    CellWordBreak = mRow(PtrRow(Row)).Cells(PtrCol(Col)).WordBreak
End Property

Public Property Let CellWordBreak(ByVal Row As Long, ByVal Col As Long, ByVal New_Value As Boolean)
     mRow(PtrRow(Row)).Cells(PtrCol(Col)).WordBreak = New_Value
    Draw
End Property

Public Property Get CellEditionLocked(ByVal Row As Long, ByVal Col As Long) As Boolean
    CellEditionLocked = mRow(PtrRow(Row)).Cells(PtrCol(Col)).EditionLocked
End Property

Public Property Let CellEditionLocked(ByVal Row As Long, ByVal Col As Long, ByVal New_Value As Boolean)
     mRow(PtrRow(Row)).Cells(PtrCol(Col)).EditionLocked = New_Value
End Property

Public Property Get CellImageIndex(ByVal Row As Long, ByVal Col As Long) As Integer
    CellImageIndex = mRow(PtrRow(Row)).Cells(PtrCol(Col)).IconIndex
End Property

Public Property Let CellImageIndex(ByVal Row As Long, ByVal Col As Long, ByVal New_Value As Integer)
    mRow(PtrRow(Row)).Cells(PtrCol(Col)).IconIndex = New_Value
    Draw
End Property

Public Function SetCurCell(ByVal Row As Long, ByVal Col As Long, Optional bEnsureVisible As Boolean)
    SelCel.Y = Row
    SelCel.X = Col
    SelRange.Start = SelCel
    SelRange.End = SelCel
    If bEnsureVisible Then
        EnsureCellVisible Row, Col
    Else
        Draw
    End If
End Function

'================================OTHERS===========================
'#################################################################
'=================================================================

Public Sub Refresh()
    UserControl_Resize
End Sub

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Get CtrlEdit() As Object
    Set CtrlEdit = Text1
End Property


Public Property Let SetMargin(ByVal New_Value As Long)
    Margin = New_Value * DpiF
    If m_Redraw Then Me.Refresh
End Property

Public Property Let AllowEdit(ByVal New_Value As Boolean)
    m_AllowEdit = New_Value
    PropertyChanged "AllowEdit"
    Draw
End Property

Public Property Get AllowEdit() As Boolean
    AllowEdit = m_AllowEdit
End Property

Public Property Let AllowColumnDrag(ByVal New_Value As Boolean)
    m_AllowColumnDrag = New_Value
    PropertyChanged "AllowColumnDrag"
    Draw
End Property

Public Property Get AllowColumnDrag() As Boolean
    AllowColumnDrag = m_AllowColumnDrag
End Property

Public Property Let AllowColumnSort(ByVal New_Value As Boolean)
    m_AllowColumnSort = New_Value
    PropertyChanged "AllowColumnSort"
    Draw
End Property

Public Property Get AllowColumnSort() As Boolean
    AllowColumnSort = m_AllowColumnSort
End Property

Public Property Let AllowRowsResize(ByVal New_Value As Boolean)
    m_AllowRowsResize = New_Value
    PropertyChanged "AllowRowsResize"
End Property

Public Property Get AllowRowsResize() As Boolean
    AllowRowsResize = m_AllowRowsResize
End Property

Public Property Let ShowHotColumn(ByVal New_Value As Boolean)
    m_ShowHotColumn = New_Value
    PropertyChanged "ShowHotColumn"
    Draw
End Property

Public Property Get ShowHotColumn() As Boolean
    ShowHotColumn = m_ShowHotColumn
End Property


Public Property Let CheckStyle(ByVal New_Value As Boolean)
    m_CheckStyle = New_Value
    PropertyChanged "CheckStyle"
    Draw
End Property

Public Property Get CheckStyle() As Boolean
    CheckStyle = m_CheckStyle
End Property

Public Property Get Redraw() As Boolean
    Redraw = m_Redraw
End Property

Public Property Let Redraw(ByVal New_Value As Boolean)
    If New_Value = True And m_Redraw <> New_Value Then
        m_Redraw = New_Value
        Refresh
    Else
        m_Redraw = New_Value
    End If
End Property

Public Property Get SelectionMode() As eSelectionMode
    SelectionMode = m_SelectionMode
End Property

Public Property Let SelectionMode(ByVal New_Value As eSelectionMode)
    m_SelectionMode = New_Value
    PropertyChanged "SelectionMode"
    Draw
End Property

Public Property Get SelectionColor() As OLE_COLOR
    SelectionColor = m_SelectionColor
End Property

Public Property Let SelectionColor(ByVal New_Value As OLE_COLOR)
    m_SelectionColor = New_Value
    PropertyChanged "SelectionColor"
    Draw
End Property

Public Property Get Font() As StdFont
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Value As StdFont)
    Set m_Font = New_Value
 
End Property

Public Property Let Font(ByVal New_Value As StdFont)
    Set m_Font = New_Value
    PropertyChanged "Font"
    Draw
End Property

Public Property Get BorderRadius() As Long
    BorderRadius = m_BorderRadius
End Property

Public Property Let BorderRadius(ByVal New_Value As Long)
    m_BorderRadius = New_Value
    PropertyChanged "BorderRadius"
    If m_Redraw Then Refresh
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_Value As OLE_COLOR)
    m_BorderColor = New_Value
    PropertyChanged "BorderColor"
    Draw
End Property

Public Property Get BorderWidth() As Long
    BorderWidth = m_BorderWidth / DpiF
End Property

Public Property Let BorderWidth(ByVal New_Value As Long)
    m_BorderWidth = New_Value * DpiF
    PropertyChanged "BorderWidth"
    If m_Redraw Then Refresh
End Property

Public Property Get LinesHorizontalColor() As OLE_COLOR
    LinesHorizontalColor = m_LinesHorizontalColor
End Property

Public Property Let LinesHorizontalColor(ByVal New_Value As OLE_COLOR)
    m_LinesHorizontalColor = New_Value
    PropertyChanged "LinesHorizontalColor"
    Draw
End Property

Public Property Get LinesHorizontalWidth() As Long
    LinesHorizontalWidth = m_LinesHorizontalWidth
End Property

Public Property Let LinesHorizontalWidth(ByVal New_Value As Long)
    m_LinesHorizontalWidth = New_Value
    PropertyChanged "LinesHorizontalWidth"
    Draw
End Property

Public Property Get LinesVerticalColor() As OLE_COLOR
    LinesVerticalColor = m_LinesVerticalColor
End Property

Public Property Let LinesVerticalColor(ByVal New_Value As OLE_COLOR)
    m_LinesVerticalColor = New_Value
    PropertyChanged "LinesVerticalColor"
    Draw
End Property


Public Property Get LinesVerticalWidth() As Long
    LinesVerticalWidth = m_LinesVerticalWidth
End Property

Public Property Let LinesVerticalWidth(ByVal New_Value As Long)
    m_LinesVerticalWidth = New_Value
    PropertyChanged "LinesVerticalWidth"
    Draw
End Property


Public Property Get HeaderLinesVerticalWidth() As Long
    HeaderLinesVerticalWidth = m_HeaderLinesVerticalWidth
End Property

Public Property Let HeaderLinesVerticalWidth(ByVal New_Value As Long)
    m_HeaderLinesVerticalWidth = New_Value
    PropertyChanged "HeaderLinesVerticalWidth"
    Draw
End Property

Public Property Get HeaderLinesHorizontalWidth() As Long
    HeaderLinesHorizontalWidth = m_HeaderLinesHorizontalWidth
End Property

Public Property Let HeaderLinesHorizontalWidth(ByVal New_Value As Long)
    m_HeaderLinesHorizontalWidth = New_Value
    PropertyChanged "HeaderLinesHorizontalWidth"
    Draw
End Property


Public Property Get ParentBackColor() As OLE_COLOR
    ParentBackColor = m_ParentBackColor
End Property

Public Property Let ParentBackColor(ByVal New_Value As OLE_COLOR)
    m_ParentBackColor = New_Value
    PropertyChanged "ParentBackColor"
    Draw
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_Value As OLE_COLOR)
    UserControl.BackColor = New_Value
    ucScrollbarV.StyleBackColor = New_Value
    ucScrollbarH.StyleBackColor = New_Value
    PropertyChanged "BackColor"
    Draw
End Property

Public Property Get RowsBackColor() As OLE_COLOR
    RowsBackColor = m_RowsBackColor
End Property

Public Property Let RowsBackColor(ByVal New_Value As OLE_COLOR)
    m_RowsBackColor = New_Value
    PropertyChanged "RowsBackColor"
    Draw
End Property

Public Property Get RowsBackColorAlt() As OLE_COLOR
    RowsBackColorAlt = m_RowsBackColorAlt
End Property

Public Property Let RowsBackColorAlt(ByVal New_Value As OLE_COLOR)
    m_RowsBackColorAlt = New_Value
    PropertyChanged "RowsBackColorAlt"
    Draw
End Property

Public Property Get RowSelectorBkColor() As OLE_COLOR
    RowSelectorBkColor = m_RowSelectorBkColor
End Property

Public Property Let RowSelectorBkColor(ByVal New_Value As OLE_COLOR)
    m_RowSelectorBkColor = New_Value
    PropertyChanged "RowSelectorBkColor"
    Draw
End Property

Public Property Get ShowHotRow() As Boolean
    ShowHotRow = m_ShowHotRow
End Property

Public Property Let ShowHotRow(ByVal New_Value As Boolean)
    m_ShowHotRow = New_Value
    PropertyChanged "ShowHotRow"
    Draw
End Property



Public Property Get HeaderBackColor() As OLE_COLOR
    HeaderBackColor = m_HeaderBackColor
End Property

Public Property Let HeaderBackColor(ByVal New_Value As OLE_COLOR)
    m_HeaderBackColor = New_Value
    PropertyChanged "HeaderBackColor"
    Draw
End Property

Public Property Get HeaderTextAlign() As eHeaderAlign
    HeaderTextAlign = m_HeaderTextAlign
End Property

Public Property Let HeaderTextAlign(ByVal New_Value As eHeaderAlign)
    m_HeaderTextAlign = New_Value
    PropertyChanged "HeaderTextAlign"
    Draw
End Property

Public Property Get HeaderTextWordBreak() As Boolean
    HeaderTextWordBreak = m_HeaderTextWordBreak
End Property

Public Property Let HeaderTextWordBreak(ByVal New_Value As Boolean)
    m_HeaderTextWordBreak = New_Value
    PropertyChanged "HeaderTextWordBreak"
    Draw
End Property

Public Property Get HeaderImageAlign() As eHeaderAlign
    HeaderImageAlign = m_HeaderImageAlign
End Property

Public Property Let HeaderImageAlign(ByVal New_Value As eHeaderAlign)
    m_HeaderImageAlign = New_Value
    PropertyChanged "HeaderImageAlign"
    Draw
End Property

Public Property Get HeaderFont() As StdFont
    Set HeaderFont = m_HeaderFont
End Property

Public Property Set HeaderFont(ByVal New_Value As StdFont)
    Set m_HeaderFont = New_Value
    PropertyChanged "HeaderFont"
    Draw
End Property

Public Property Let HeaderFont(ByVal New_Value As StdFont)
    Set HeaderFont = New_Value
End Property

Public Property Get ColumnsAutoFit() As Boolean
    ColumnsAutoFit = m_ColumnsAutoFit
End Property

Public Property Let ColumnsAutoFit(ByVal New_Value As Boolean)
    m_ColumnsAutoFit = New_Value
    PropertyChanged "ColumnsAutoFit"
    If m_Redraw Then Refresh
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Value As Boolean)
    UserControl.Enabled = New_Value
    PropertyChanged "Enabled"
End Property

Public Property Get LastRowIsFooter() As Boolean
    LastRowIsFooter = m_LastRowIsFooter
End Property

Public Property Let LastRowIsFooter(ByVal New_Value As Boolean)
    m_LastRowIsFooter = New_Value
    PropertyChanged "LastRowIsFooter"
    Draw
End Property

Public Property Get FixedColumns() As Long
    FixedColumns = m_FixedColumns
End Property

Public Property Let FixedColumns(ByVal New_Value As Long)
    m_FixedColumns = New_Value
    PropertyChanged "FixedColumns"
    If m_Redraw Then Me.Refresh
End Property

Public Property Get FixedRows() As Long
    FixedRows = m_FixedRows
End Property

Public Property Let FixedRows(ByVal New_Value As Long)
    m_FixedRows = New_Value
    PropertyChanged "FixedRows"
    If m_Redraw Then Me.Refresh
End Property

Public Property Get DragIcon() As IPictureDisp
Set DragIcon = Extender.DragIcon
End Property

Public Property Let DragIcon(ByVal Value As IPictureDisp)
    Extender.DragIcon = Value
End Property

Public Property Set DragIcon(ByVal Value As IPictureDisp)
    Set Extender.DragIcon = Value
End Property

Public Property Get DragMode() As Integer
    DragMode = Extender.DragMode
End Property

Public Property Let DragMode(ByVal Value As Integer)
    Extender.DragMode = Value
End Property

Public Property Get OLEDropMode() As GridPlusOLEDropModeConstants
    OLEDropMode = UserControl.OLEDropMode
End Property

'*1
Public Property Let OLEDropMode(ByVal Value As GridPlusOLEDropModeConstants)
    Select Case Value
        Case OLEDropModeNone, OLEDropModeManual
            UserControl.OLEDropMode = Value
        Case Else
            Err.Raise 380
    End Select
    UserControl.PropertyChanged "OLEDropMode"
End Property


Public Sub Drag(Optional ByRef Action As Variant)
    If IsMissing(Action) Then Extender.Drag Else Extender.Drag Action
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


Public Property Get GradientStyle() As Boolean
    GradientStyle = m_GradientStyle
End Property

Public Property Let GradientStyle(ByVal NewValue As Boolean)
    m_GradientStyle = NewValue
    PropertyChanged "GradientStyle"
    If m_Redraw Then Me.Refresh
End Property

Public Property Get ScrollBarStyle() As sbStyleCts
    ScrollBarStyle = ucScrollbarV.Style
End Property

Public Property Let ScrollBarStyle(ByVal New_Value As sbStyleCts)
    ucScrollbarV.Style = New_Value
    ucScrollbarH.Style = New_Value
    If New_Value = sGoogle Then
        ucScrollbarV.Width = 8 * DpiF
        ucScrollbarH.Height = 8 * DpiF
        ucScrollbarV.ShowButtons = False
        ucScrollbarH.ShowButtons = False
    Else
        ucScrollbarV.Width = 14 * DpiF
        ucScrollbarH.Height = 14 * DpiF
        ucScrollbarV.ShowButtons = True
        ucScrollbarH.ShowButtons = True
    End If
    
    PropertyChanged "ScrollBarStyle"
    
    If m_Redraw Then Me.Refresh
End Property

'/>Lizano Días - Lynxgrid
Public Property Get FocusRectStyle() As lgFocusRectStyleEnum
   FocusRectStyle = m_FocusRectStyle
End Property

Public Property Let FocusRectStyle(ByVal vNewValue As lgFocusRectStyleEnum)
   m_FocusRectStyle = vNewValue
   PropertyChanged "FocusRectStyle"
   Call DisplayChange
Redraw = True
End Property

Public Property Get FocusRectMode() As lgFocusRectModeEnum
   FocusRectMode = m_FocusRectMode
End Property

Public Property Let FocusRectMode(ByVal vNewValue As lgFocusRectModeEnum)
   m_FocusRectMode = vNewValue
   PropertyChanged "FocusRectMode"
   Call DisplayChange
End Property

Private Sub DisplayChange()
   If Redraw Then
      Call Refresh
   End If
End Sub

Public Property Get FocusRectColor() As OLE_COLOR
   FocusRectColor = m_FocusRectColor
End Property

Public Property Let FocusRectColor(ByVal vNewValue As OLE_COLOR)
   m_FocusRectColor = vNewValue
   PropertyChanged "FocusRectColor"
End Property
'/<Lizano Días - Lynxgrid

Public Sub GetSelectionRange(StartRow As Long, StartCol As Long, EndRow As Long, EndCol As Long)
    StartRow = SelRange.Start.Y
    StartCol = SelRange.Start.X
    EndRow = SelRange.End.Y
    EndCol = SelRange.End.X
End Sub

Public Function SetSelectionRange(ByVal StartRow As Long, ByVal StartCol As Long, ByVal EndRow As Long, ByVal EndCol As Long)
     SelRange.Start.Y = StartRow
     SelRange.Start.X = StartCol
     SelRange.End.Y = EndRow
     SelRange.End.X = EndCol
     Draw
End Function

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition))
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, state As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, UserControl.ScaleX(X, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition), state)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Public Sub OLEDrag()
    UserControl.OLEDrag
End Sub

Public Sub FillFromRS(rs As Object)
    Dim Row As Long, Col As Long
    Dim bRedraw As Boolean
    If rs.state = 0 Then Exit Sub
    rs.MoveFirst
    With Me
        bRedraw = .Redraw
        .Redraw = False
        .ColsCount = rs.Fields.Count
        .RowsCount = rs.RecordCount
      
        For Col = 0 To rs.Fields.Count - 1
            .ColumnText(Col) = rs.Fields(Col).Name
            Select Case rs.Fields(Col).Type
                Case 3 'Number
                    .ColDataType(Col) = GP_NUMERIC
                    .ColAlign(Col) = CenterRight
                Case 6
                    .ColDataType(Col) = GP_CURRENCY
                    .ColAlign(Col) = CenterRight
                Case 7
                    .ColDataType(Col) = GP_DATE
                    .ColAlign(Col) = CenterRight
                Case 11
                    .ColDataType(Col) = GP_BOOLEAN
                    .ColAlign(Col) = CenterCenter
                Case Else '203, 204
                    .ColDataType(Col) = GP_STRING
                    .ColAlign(Col) = CenterLeft
            End Select
        Next
        
  
        Do While Not rs.EOF
            For Col = 0 To rs.Fields.Count - 1
                .CellValue(Row, Col) = rs.Fields(Col)
            Next
            rs.MoveNext
            Row = Row + 1
        Loop
        .AutoWidthAllColumns
        .Redraw = bRedraw
    End With
End Sub


Public Sub SwapRow(ByVal SrcRow As Long, ByVal DestRow As Long)
    SwapLng PtrRow(SrcRow), PtrRow(DestRow)
    Draw
End Sub

Public Sub SwapCol(ByVal SrcCol As Long, ByVal DestCol As Long)
    SwapLng PtrCol(SrcCol), PtrCol(DestCol)
    Draw
End Sub

Public Sub RowMoveTo(ByVal SrcRow As Long, ByVal DestPos As Long)
    Dim Row As Long
    If SrcRow < DestPos Then
        For Row = SrcRow To DestPos - 1
            SwapLng PtrRow(Row), PtrRow(Row + 1)
        Next
    Else
        For Row = SrcRow - 1 To DestPos Step -1
            SwapLng PtrRow(Row), PtrRow(Row + 1)
        Next
    End If
    Draw
End Sub

Public Sub ColMoveTo(ByVal SrcCol As Long, ByVal DestPos As Long)
    Dim Col As Long
    If SrcCol < DestPos Then
        For Col = SrcCol To DestPos - 1
            SwapLng PtrCol(Col), PtrCol(Col + 1)
        Next
    Else
        For Col = SrcCol - 1 To DestPos Step -1
            SwapLng PtrCol(Col), PtrCol(Col + 1)
        Next
    End If
    Draw
End Sub

'================================END PROPERTY=====================
'#################################################################
'=================================================================


Private Sub UserControl_Initialize()
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)

    DpiF = GetWindowsDPI

    Call LoadFontsIcons
    
    m_AllowColumnSort = True
    mColDrag.X = -1
    mSortColumn = -1
    mSortSubColumn = -1
    SizeColumn.CurColumn = -1
    SizeColumn.CurRow = -1
    m_SelectionColor = vbHighlight
    HotRow = -1
    HotCol = -1
    mHotCol = -1
    EmptyPoint mHotCell
    EmptyPoint SelRange.Start
    EmptyPoint SelRange.End
    EmptyPoint SelCel
    EmptyPoint mCellEdit
    EmptyPoint mHotPart
    eSelBy = SelectByNone
    Margin = 4 * DpiF
    m_BorderWidth = 1 * DpiF
    m_Redraw = True
    
    '/>Jose Liza - FocusRect
    m_CellFocusRect.X = -1
    m_CellFocusRect.Y = -1
    '/<Jose Liza - FocusRect
    
    hCurHands = LoadCursor(ByVal 0&, IDC_HAND)
    #If OCX_VERSION Then
        Call modIOleInPlaceActiveObject.InitIPAO  'OCX VERSION
    #End If
End Sub

Private Sub UserControl_InitProperties()
    m_ShowHotColumn = True
    m_ColsWidth = 100
    m_AllowRowsResize = True
    m_HeaderHeight = 50
    m_RowsHeight = 20
    m_RowSelectorWidth = 50
    m_SelectionMode = GP_SelFree
    m_SelectionColor = vbHighlight
    m_LinesHorizontalColor = vb3DLight
    m_LinesVerticalColor = vb3DLight
    Set m_Font = Ambient.Font
    m_RowSelectorBkColor = vbButtonFace
    m_HeaderBackColor = vbButtonFace
    m_BorderColor = vb3DLight
    m_HeaderTextWordBreak = True
    m_HeaderTextAlign = HA_DefaultColumn
    m_HeaderImageAlign = HA_DefaultColumn
    Set m_HeaderFont = Ambient.Font
    m_HeaderFont.Bold = True
    m_RowsBackColor = vbWindowBackground
    m_RowsBackColorAlt = vbWindowBackground
    m_LinesHorizontalWidth = 1
    m_LinesVerticalWidth = 1
    m_HeaderLinesHorizontalWidth = 1
    m_HeaderLinesVerticalWidth = 1
    m_AllowEdit = True
    m_ParentBackColor = Ambient.BackColor
    
    '/> Lizano Dias - LynxGrid
    m_FocusRectMode = lgFocusRectModeEnum.lgNone
    m_FocusRectStyle = lgFocusRectStyleEnum.lgFRHeavy
    m_FocusRectColor = &H40C0&
    '/< Lizano Dias - LynxGrid
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        m_AllowColumnDrag = .ReadProperty("AllowColumnDrag", False)
        m_AllowColumnSort = .ReadProperty("AllowColumnSort", True)
        m_AllowRowsResize = .ReadProperty("AllowRowsResize", True)
        m_ShowHotColumn = .ReadProperty("ShowHotColumn", True)
        m_ColumnsAutoFit = .ReadProperty("ColumnsAutoFit", False)
        m_ColsWidth = .ReadProperty("ColsWidth", 100) * DpiF
        m_HeaderHeight = .ReadProperty("HeaderHeight", 40) * DpiF
        m_RowsHeight = .ReadProperty("RowsHeight", 20) * DpiF
        m_RowSelectorWidth = .ReadProperty("RowSelectorWidth", 50) * DpiF
        m_SelectionMode = .ReadProperty("SelectionMode", GP_SelFree)
        m_SelectionColor = .ReadProperty("SelectionColor", vbHighlight)
        m_LinesHorizontalColor = .ReadProperty("LinesHorizontalColor", vb3DLight)
        m_LinesVerticalColor = .ReadProperty("LinesVerticalColor", vb3DLight)
        m_ParentBackColor = .ReadProperty("ParentBackColor", Ambient.BackColor)
        UserControl.BackColor = .ReadProperty("BackColor", vbWindowBackground)
        m_RowsBackColor = .ReadProperty("RowsBackColor", vbWindowBackground)
        m_RowsBackColorAlt = .ReadProperty("RowsBackColorAlt", vbWindowBackground)
        Set m_Font = .ReadProperty("Font", Ambient.Font)
        m_RowSelectorBkColor = .ReadProperty("RowSelectorBkColor", vbButtonFace)
        m_HeaderBackColor = .ReadProperty("HeaderBackColor", vbButtonFace)
        m_BorderColor = .ReadProperty("BorderColor", vb3DLight)
        m_HeaderTextWordBreak = .ReadProperty("HeaderTextWordBreak", True)
        m_HeaderTextAlign = .ReadProperty("HeaderTextAlign", HA_DefaultColumn)
        m_HeaderImageAlign = .ReadProperty("HeaderImageAlign", HA_DefaultColumn)
        Set m_HeaderFont = .ReadProperty("HeaderFont", Ambient.Font)
        m_LinesHorizontalWidth = .ReadProperty("LinesHorizontalWidth", 1)
        m_HeaderLinesVerticalWidth = .ReadProperty("HeaderLinesVerticalWidth", 1)
        m_HeaderLinesHorizontalWidth = .ReadProperty("HeaderLinesHorizontalWidth", 1)
        m_LinesVerticalWidth = .ReadProperty("LinesVerticalWidth", 1)
        m_CheckStyle = .ReadProperty("CheckStyle", False)
        m_BorderRadius = .ReadProperty("BorderRadius", 0)
        m_AllowEdit = .ReadProperty("AllowEdit", True)
        Me.ColsCount = .ReadProperty("ColsCount", 0)
        Me.RowsCount = .ReadProperty("RowsCount", 0)
        UserControl.Enabled = .ReadProperty("Enabled", UserControl.Enabled)
        m_FixedRows = .ReadProperty("FixedRows", 0)
        m_FixedColumns = .ReadProperty("FixedColumns", 0)
        m_LastRowIsFooter = .ReadProperty("LastRowIsFooter", False)
        Me.ScrollBarStyle = .ReadProperty("ScrollBarStyle", sFlat)
        m_BorderWidth = .ReadProperty("BorderWidth", 1) * DpiF
        Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
        UserControl.MousePointer = .ReadProperty("MousePointer", vbDefault)
        UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
        m_GradientStyle = .ReadProperty("GradientStyle", False)
        m_ShowHotRow = .ReadProperty("ShowHotRow", False)
        
        '/> Lizano Dias - LynxGrid
        m_FocusRectColor = .ReadProperty("FocusRectColor", &H40C0&)
        m_FocusRectMode = .ReadProperty("FocusRectMode", lgFocusRectModeEnum.lgNone)
        m_FocusRectStyle = .ReadProperty("FocusRectStyle", lgFocusRectStyleEnum.lgFRHeavy)
        '/< Lizano Dias - LynxGrid
        
    End With

    'If m_BorderVisible Then m_BorderWidth = 7 * DpiF Else m_BorderWidth = 0
    ucScrollbarH.WheelChange = m_ColsWidth / 2
    ucScrollbarV.TrackMouseWheelOnHwnd UserControl.hwnd
    ucScrollbarV.AttachHorizontalScrollBar ucScrollbarH

    UserControl_Resize

End Sub

Private Sub UserControl_Terminate()
    Dim i As Long, J As Long
    DestroyCursor hCurHands
    If Not cHeaderImageList Is Nothing Then
        For i = 1 To cHeaderImageList.Count
            GdipDisposeImage CLng(cHeaderImageList(i))
        Next
    End If
    
    For i = 0 To m_ColsCount - 1
        If Not mCol(i).ColImgList Is Nothing Then
            For J = 1 To mCol(i).ColImgList.Count
                GdipDisposeImage CLng(mCol(i).ColImgList(J))
            Next
        End If
    Next
        
    Call GdiplusShutdown(GdipToken)
    
    #If OCX_VERSION Then
        Call modIOleInPlaceActiveObject.TerminateIPAO  'OCX VERSION
    #End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "AllowColumnDrag", m_AllowColumnDrag, False
        .WriteProperty "AllowColumnSort", m_AllowColumnSort, True
        .WriteProperty "AllowRowsResize", m_AllowRowsResize, True
        .WriteProperty "ShowHotColumn", m_ShowHotColumn, True
        .WriteProperty "ColumnsAutoFit", m_ColumnsAutoFit, False
        .WriteProperty "ColsCount", m_ColsCount, 0
        .WriteProperty "RowsCount", m_RowsCount, 0
        .WriteProperty "ColsWidth", m_ColsWidth / DpiF, 100 / DpiF
        .WriteProperty "HeaderHeight", m_HeaderHeight / DpiF, 40 / DpiF
        .WriteProperty "RowsHeight", m_RowsHeight / DpiF, 20 / DpiF
        .WriteProperty "RowSelectorWidth", m_RowSelectorWidth / DpiF, 50 / DpiF
        .WriteProperty "SelectionMode", m_SelectionMode, GP_SelFree
        .WriteProperty "SelectionColor", m_SelectionColor, vbHighlight
        .WriteProperty "LinesHorizontalColor", m_LinesHorizontalColor, vb3DLight
        .WriteProperty "LinesVerticalColor", m_LinesVerticalColor, vb3DLight
        .WriteProperty "ParentBackColor", m_ParentBackColor, Ambient.BackColor
        .WriteProperty "BackColor", UserControl.BackColor, vbWindowBackground
        .WriteProperty "RowsBackColor", m_RowsBackColor, vbWindowBackground
        .WriteProperty "RowsBackColorAlt", m_RowsBackColorAlt, vbWindowBackground
        .WriteProperty "RowSelectorBkColor", m_RowSelectorBkColor, vbButtonFace
        .WriteProperty "HeaderBackColor", m_HeaderBackColor, vbButtonFace
        .WriteProperty "Font", m_Font, Ambient.Font
        .WriteProperty "BorderColor", m_BorderColor, vb3DLight
        .WriteProperty "HeaderTextWordBreak", m_HeaderTextWordBreak, True
        .WriteProperty "HeaderTextAlign", m_HeaderTextAlign, HA_DefaultColumn
        .WriteProperty "HeaderImageAlign", m_HeaderImageAlign, HA_DefaultColumn
        .WriteProperty "HeaderFont", m_HeaderFont, Ambient.Font
        .WriteProperty "LinesHorizontalWidth", m_LinesHorizontalWidth, 1
        .WriteProperty "LinesVerticalWidth", m_LinesVerticalWidth, 1
        .WriteProperty "HeaderLinesHorizontalWidth", m_HeaderLinesHorizontalWidth, 1
        .WriteProperty "HeaderLinesVerticalWidth", m_HeaderLinesVerticalWidth, 1
        .WriteProperty "CheckStyle", m_CheckStyle, False
        .WriteProperty "BorderRadius", m_BorderRadius, 0
        .WriteProperty "AllowEdit", m_AllowEdit, True
        .WriteProperty "Enabled", UserControl.Enabled, True
        .WriteProperty "FixedRows", m_FixedRows, 0
        .WriteProperty "FixedColumns", m_FixedColumns, 0
        .WriteProperty "LastRowIsFooter", m_LastRowIsFooter, False
        .WriteProperty "ScrollBarStyle", ucScrollbarV.Style, sFlat
        .WriteProperty "BorderWidth", m_BorderWidth / DpiF, 1
        .WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
        .WriteProperty "MousePointer", UserControl.MousePointer, vbDefault
        .WriteProperty "MouseIcon", UserControl.MouseIcon, Nothing
        .WriteProperty "GradientStyle", m_GradientStyle, False
        .WriteProperty "ShowHotRow", m_ShowHotRow, False
        
        '/> Lizano Dias - LynxGrid
        .WriteProperty "FocusRectMode", m_FocusRectMode, lgFocusRectModeEnum.lgNone
        .WriteProperty "FocusRectColor", m_FocusRectColor, &H40C0&
        .WriteProperty "FocusRectStyle", m_FocusRectStyle, lgFocusRectStyleEnum.lgFRHeavy
        '/< Lizano Dias - LynxGrid
        
    End With
End Sub

Private Sub LoadFontsIcons()
    Set m_SegoeFont = New StdFont
    m_SegoeFont.Name = "Segoe MDL2 Assets"
    
    If m_SegoeFont.Name = "Segoe MDL2 Assets" Then
        isSegoeFontInstaled = True
    Else
        'abandoned
        Set m_SegoeFont = Nothing
        Set m_Wingdings2 = New StdFont
        Set m_Wingdings3 = New StdFont
        m_Wingdings2.Name = "Wingdings 2"
        m_Wingdings3.Name = "Wingdings 3"
        isSegoeFontInstaled = False
    End If
End Sub

'ByEduardo
Private Function CloneFont(nOrigFont As iFont) As StdFont 'By Eduardo
    If nOrigFont Is Nothing Then Exit Function
    nOrigFont.Clone CloneFont
End Function

Public Function ChrW2(ByVal CharCode As Long) As String
  Const POW10 As Long = 2 ^ 10
  If CharCode <= &HFFFF& Then ChrW2 = ChrW$(CharCode) Else _
                              ChrW2 = ChrW$(&HD800& + (CharCode And &HFFFF&) \ POW10) & _
                                      ChrW$(&HDC00& + (CharCode And (POW10 - 1)))
End Function

Private Sub DrawIconFont(hdc As Long, Id As enuIdIconFont, Size As Long, ByVal bHighlight As Boolean, Rect As Rect, Align As eGridAlign)
    Dim iFont As iFont
    Dim hPrevFont As Long
    Dim PrevForeColor As Long
    Dim sChr As String
    

    PrevForeColor = GetTextColor(hdc)
    If bHighlight Then SetTextColor hdc, m_SelectionColor
     
    If isSegoeFontInstaled Then
        
        Select Case Id
            Case IIF_SortAsc
                sChr = ChrW2(&HE96D)
            Case IIF_SortDes
                sChr = ChrW2(&HE96E)
            Case IIF_Edit
                sChr = ChrW2(&HE1C2)
            Case IIF_GroupColapsed
                sChr = ChrW2(&HE970)
            Case IIF_GroupExpanded
                sChr = ChrW2(&HE96E)
            Case IIF_TreeColapsed
                sChr = ChrW2(&HF164)
            Case IIF_TreeExpanded
                sChr = ChrW2(&HF166)
            Case IIF_Asterisk
                sChr = ChrW2(&HEA38)
            Case IIF_ArrowRight
                sChr = ChrW2(&HE937)
        End Select

        m_SegoeFont.Size = Size
        Set iFont = m_SegoeFont

    Else 'abandoned
        m_Wingdings2.Size = Size
        Set iFont = m_Wingdings2
    End If

 
    hPrevFont = SelectObject(hdc, iFont.hFont)
    DrawText hdc, StrPtr(sChr), 1, Rect, DT_SINGLELINE Or Align
    SetTextColor hdc, PrevForeColor
    Call SelectObject(hdc, hPrevFont)

End Sub

Private Sub DrawCheckBox(hdc As Long, ByVal Value As Variant, ByVal bHighlight As Boolean, Rect As Rect, Align As eGridAlign)
    Dim iFont As iFont
    Dim hPrevFont As Long
    Dim PrevForeColor As Long
    Dim sChr As String
    Dim Color As Long
    Color = m_SelectionColor

    If (Color And &H80000000) Then Color = GetSysColor(Color And &HFF&)
    PrevForeColor = GetTextColor(hdc)
    
    If isSegoeFontInstaled Then
        m_SegoeFont.Size = 10.66 '<--=16/1.5
        Set iFont = m_SegoeFont
        If IsNull(Value) Or IsEmpty(Value) Then
            sChr = ChrW2(&HF16D)
        ElseIf Value = True Then
            If Not bHighlight Then
                SetTextColor hdc, Color
            End If
            sChr = ChrW2(&HE005)
        Else
            sChr = ChrW2(&HE003)
        End If
    Else
        m_Wingdings2.Size = 14
        Set iFont = m_Wingdings2
        If Value Then
            If Not bHighlight Then
                PrevForeColor = GetTextColor(hdc)
                SetTextColor hdc, Color
            End If
            sChr = ChrW2(&H52)
        Else
            sChr = ChrW2(&HA3)
        End If
    End If

    hPrevFont = SelectObject(hdc, iFont.hFont)
    DrawText hdc, StrPtr(sChr), 1, Rect, DT_SINGLELINE Or Align
    SetTextColor hdc, PrevForeColor
    Call SelectObject(hdc, hPrevFont)
    
End Sub

'*-
'routine extracted from VB FlexGrid, bubble sort was changed to quicksort
Public Sub Sort(Optional ByVal vCol1 As Long = C_NULL_RESULT, _
                Optional ByVal vCol1SortType As lgSortTypeEnum = C_NULL_RESULT, _
                Optional ByVal vCol2 As Long = C_NULL_RESULT, _
                Optional ByVal vCol2SortType As lgSortTypeEnum = C_NULL_RESULT)

   '// Purpose: Sort Grid based on current Sort Columns.
  Dim lCount    As Long
  Dim i As Long, lStart As Long, lEnd As Long
  
   If Not m_RowsCount < 1 Then '// Error Prevention
      'If UpdateCell() Then
      
        
        RaiseEvent BeforeSorted
        
         '// Set new Columns if specified
         If Not (vCol1 = C_NULL_RESULT) Then
            mSortColumn = vCol1
         End If
   
         If Not (vCol2 = C_NULL_RESULT) Then
            mSortSubColumn = vCol2
         End If
   
         '// Validate Sort Columns
         If mSortColumn = C_NULL_RESULT And Not (mSortSubColumn = C_NULL_RESULT) Then
            mSortColumn = mSortSubColumn
            mSortSubColumn = C_NULL_RESULT
   
         ElseIf mSortColumn = mSortSubColumn Then
            mSortSubColumn = C_NULL_RESULT
         End If
   
         '// Fix column number in case column order was changed
         If vCol1 = C_NULL_RESULT Then
            mSortColumn = PtrCol(mSortColumn)
         End If
         
         If vCol2 = C_NULL_RESULT Then
            If Not mSortSubColumn = C_NULL_RESULT Then
               mSortSubColumn = PtrCol(mSortSubColumn)
            End If
         End If
         
         '// Set Sort Order if specified - otherwise inverse last Sort Order
         With mCol(mSortColumn)
            If vCol1SortType = C_NULL_RESULT Then
   
               Select Case .nSortOrder
                Case lgSTNormal
                   .nSortOrder = lgSTDescending
    
                Case lgSTAscending
                   '//.nSortOrder = lgSTNormal
                   .nSortOrder = lgSTDescending
                Case lgSTDescending
                   .nSortOrder = lgSTAscending
               End Select
   
            Else
               .nSortOrder = vCol1SortType
               '.nSortOrder = lgSTAscending
            End If
         End With
   
         If Not (mSortSubColumn = C_NULL_RESULT) Then
            With mCol(mSortSubColumn)
               If vCol2SortType = C_NULL_RESULT Then
                  '.nSortOrder = mCol(mSortColumn).nSortOrder
                  
                  
                    Select Case .nSortOrder
                        Case lgSTNormal
                           .nSortOrder = lgSTDescending
                        
                        Case lgSTAscending
                           '//.nSortOrder = lgSTNormal
                           .nSortOrder = lgSTDescending
                        Case lgSTDescending
                           .nSortOrder = lgSTAscending
                    End Select
                  
               Else
                  .nSortOrder = vCol2SortType
               End If
            End With
         End If
   
         '// Note previously selected Row
'         If Not (mRow = C_NULL_RESULT) Then
'            lRowIndex = PtrRow(mRow)
'         End If
   
         If mCol(mSortColumn).nSortOrder = lgSTNormal Then
 
            For lCount = 0 To m_RowsCount - 1
               PtrRow(lCount) = lCount
            Next lCount
   
            mSortColumn = C_NULL_RESULT
            mSortSubColumn = C_NULL_RESULT
   
         Else
            lStart = m_FixedRows
            lEnd = m_RowsCount - 1
            If m_LastRowIsFooter Then lEnd = lEnd - 1
            '----------------------
            For lCount = 0 To m_RowsCount - 1
                If mRow(PtrRow(lCount)).IsGroup Then
                    lStart = lCount + 1
                    For i = lStart To m_RowsCount - 1
                        If mRow(PtrRow(i)).IsGroup Then Exit For
                    Next
                    lEnd = i - 1
                    Call SortArray(lStart, lEnd, mSortColumn, mCol(mSortColumn).nSortOrder)
                End If
            Next
            If lStart = m_FixedRows Then
                Call SortArray(lStart, lEnd, mSortColumn, mCol(mSortColumn).nSortOrder)
            End If
            '----------------------
            Call SortSubList
         End If
   
'         For lCount = 0 To m_RowsCount-1
'            If PtrRow(lCount) = lRowIndex Then
'               RowColSet lCount '// keep selected row visible
'               Exit For
'            End If
'         Next lCount
   
         'RaiseEvent SortComplete
      'End If
      
      DoEvents '// added to give the system time to update
      
      RaiseEvent AfterSorted
   End If
   Draw
   
End Sub

Private Sub SortArray(ByVal lFirst As Long, ByVal lLast As Long, ByVal lSortColumn As Long, ByVal nSortType As Integer)

   '// Purpose: A simple data-type aware quick-sort method to Sort Grid Rows
   Select Case mCol(lSortColumn).DataType
   Case GP_BOOLEAN
      QSortArrayBool lFirst, lLast, lSortColumn, nSortType

   Case GP_DATE
      QSortArrayDate lFirst, lLast, lSortColumn, nSortType

   Case GP_NUMERIC, GP_CURRENCY
      QSortArrayNumeric lFirst, lLast, lSortColumn, nSortType
      
   Case GP_CUSTOM
      QSortArrayCustom lFirst, lLast, lSortColumn, nSortType

   Case Else
      QSortArrayString lFirst, lLast, lSortColumn, nSortType
   End Select

End Sub


Private Sub QSortArrayCustom(ByVal First As Long, ByVal Last As Long, ByVal lSortColumn As Long, ByVal nSortType As Integer)
    Dim Low As Long, High As Long
    Dim MidValue As Variant
    Dim Item As Long
    Dim bSwap As Boolean
    
    Low = First
    High = Last
    MidValue = mRow(PtrRow((First + Last) \ 2)).Cells(lSortColumn).Value
    
    
    If nSortType = lgSTDescending Then
        
        Do
            Do
                bSwap = False
                RaiseEvent CustomSort(False, lSortColumn, mRow(PtrRow(Low)).Cells(lSortColumn).Value, MidValue, bSwap)
                If bSwap = False Then Exit Do
                Low = Low + 1
            Loop
 
            Do
                bSwap = False
                RaiseEvent CustomSort(False, lSortColumn, MidValue, mRow(PtrRow(High)).Cells(lSortColumn).Value, bSwap)
                If bSwap = False Then Exit Do
                High = High - 1

            Loop
            
            If Low <= High Then
                Item = PtrRow(Low)
                PtrRow(Low) = PtrRow(High)
                PtrRow(High) = Item
               
                Low = Low + 1
                High = High - 1
            End If
        Loop While Low <= High
        
    Else
        Do
            Do
                bSwap = False
                RaiseEvent CustomSort(True, lSortColumn, mRow(PtrRow(Low)).Cells(lSortColumn).Value, MidValue, bSwap)
                If bSwap = False Then Exit Do
                Low = Low + 1
            Loop
 
            Do
                bSwap = False
                RaiseEvent CustomSort(True, lSortColumn, MidValue, mRow(PtrRow(High)).Cells(lSortColumn).Value, bSwap)
                If bSwap = False Then Exit Do
                High = High - 1
            Loop
        

            If Low <= High Then
                Item = PtrRow(Low)
                PtrRow(Low) = PtrRow(High)
                PtrRow(High) = Item

                Low = Low + 1
                High = High - 1
            End If
        Loop While Low <= High
        
    End If
    If First < High Then QSortArrayCustom First, High, lSortColumn, nSortType
    If Low < Last Then QSortArrayCustom Low, Last, lSortColumn, nSortType
End Sub

Private Function MyBool(Value As Variant) As Boolean
    If IsNull(Value) Then
        MyBool = False
    Else
        MyBool = CBool(Value)
    End If
End Function

Private Sub QSortArrayBool(ByVal First As Long, ByVal Last As Long, ByVal lSortColumn As Long, ByVal nSortType As Integer)
    Dim Low As Long, High As Long
    Dim MidValue As Boolean
    Dim Item As Long
    
    Low = First
    High = Last
    MidValue = MyBool(mRow(PtrRow((First + Last) \ 2)).Cells(lSortColumn).Value)
    
    
    If nSortType = lgSTDescending Then
        
        Do
            While (MyBool(mRow(PtrRow(Low)).Cells(lSortColumn).Value) > MidValue)
                Low = Low + 1
            Wend
 
            While (MidValue > MyBool(mRow(PtrRow(High)).Cells(lSortColumn).Value))
                High = High - 1
            Wend
            
            If Low <= High Then
                Item = PtrRow(Low)
                PtrRow(Low) = PtrRow(High)
                PtrRow(High) = Item
               
                Low = Low + 1
                High = High - 1
            End If
        Loop While Low <= High
        
    Else
        Do
            While (MyBool(mRow(PtrRow(Low)).Cells(lSortColumn).Value) < MidValue)
                Low = Low + 1
            Wend

            
            While (MidValue < MyBool(mRow(PtrRow(High)).Cells(lSortColumn).Value))
                High = High - 1
            Wend

            If Low <= High Then
                Item = PtrRow(Low)
                PtrRow(Low) = PtrRow(High)
                PtrRow(High) = Item

                Low = Low + 1
                High = High - 1
            End If
        Loop While Low <= High
        
    End If
    If First < High Then QSortArrayBool First, High, lSortColumn, nSortType
    If Low < Last Then QSortArrayBool Low, Last, lSortColumn, nSortType
End Sub


Private Sub QSortArrayDate(ByVal First As Long, ByVal Last As Long, ByVal lSortColumn As Long, ByVal nSortType As Integer)
    Dim Low As Long, High As Long
    Dim MidValue As Variant 'As Date
    Dim Item As Long

    Low = First
    High = Last
    MidValue = MyDate(mRow(PtrRow((First + Last) \ 2)).Cells(lSortColumn).Value)
    
    If nSortType = lgSTDescending Then
        
        Do
            While (MyDate(mRow(PtrRow(Low)).Cells(lSortColumn).Value) > MidValue)
                Low = Low + 1
            Wend

            While (MidValue > MyDate(mRow(PtrRow(High)).Cells(lSortColumn).Value))
                High = High - 1
            Wend
            
            If Low <= High Then
                Item = PtrRow(Low)
                PtrRow(Low) = PtrRow(High)
                PtrRow(High) = Item
               
                Low = Low + 1
                High = High - 1
            End If
        Loop While Low <= High
        
    Else
        Do
            While (MyDate(mRow(PtrRow(Low)).Cells(lSortColumn).Value) < MidValue)
                Low = Low + 1
            Wend
            
            While (MidValue < MyDate(mRow(PtrRow(High)).Cells(lSortColumn).Value))
                High = High - 1
            Wend

            If Low <= High Then
                Item = PtrRow(Low)
                PtrRow(Low) = PtrRow(High)
                PtrRow(High) = Item

                Low = Low + 1
                High = High - 1
            End If
        Loop While Low <= High
        
    End If
    If First < High Then QSortArrayDate First, High, lSortColumn, nSortType
    If Low < Last Then QSortArrayDate Low, Last, lSortColumn, nSortType
End Sub


Private Sub QSortArrayNumeric(ByVal First As Long, ByVal Last As Long, ByVal lSortColumn As Long, ByVal nSortType As Integer)
    Dim Low As Long, High As Long
    Dim MidValue As Double
    Dim Item As Long
    
    Low = First
    High = Last
    MidValue = rVal(mRow(PtrRow((First + Last) \ 2)).Cells(lSortColumn).Value)
    
    
    If nSortType = lgSTDescending Then
        Do
            While (rVal(mRow(PtrRow(Low)).Cells(lSortColumn).Value) > MidValue)
                Low = Low + 1
            Wend
 
            While (MidValue > rVal(mRow(PtrRow(High)).Cells(lSortColumn).Value))
                High = High - 1
            Wend
            
            If Low <= High Then
                Item = PtrRow(Low)
                PtrRow(Low) = PtrRow(High)
                PtrRow(High) = Item
               
                Low = Low + 1
                High = High - 1
            End If
        Loop While Low <= High
        
    Else
        Do
            While (rVal(mRow(PtrRow(Low)).Cells(lSortColumn).Value) < MidValue)
                Low = Low + 1
            Wend
 
            While (MidValue < rVal(mRow(PtrRow(High)).Cells(lSortColumn).Value))
                High = High - 1
            Wend

            If Low <= High Then
                Item = PtrRow(Low)
                PtrRow(Low) = PtrRow(High)
                PtrRow(High) = Item

                Low = Low + 1
                High = High - 1
            End If
        Loop While Low <= High
        
    End If
    If First < High Then QSortArrayNumeric First, High, lSortColumn, nSortType
    If Low < Last Then QSortArrayNumeric Low, Last, lSortColumn, nSortType
End Sub

Private Function MyDate(Value As Variant) As Date
    If IsNull(Value) Or Not IsDate(Value) Then
        MyDate = CDate("01/01/1000")
    Else
        MyDate = CDate(Value)
    End If
End Function


Private Function MyString(Value As Variant) As String
    If IsNull(Value) Then
        MyString = vbNullString
    Else
        MyString = CStr(Value)
    End If
End Function

 
Private Sub QSortArrayString(ByVal First As Long, ByVal Last As Long, ByVal lSortColumn As Long, ByVal nSortType As Integer)
    Dim Low As Long, High As Long
    Dim MidValue As Variant 'As String
    Dim Item As Long

    Low = First
    High = Last
    MidValue = MyString(mRow(PtrRow((First + Last) \ 2)).Cells(lSortColumn).Value)
    
    
    If nSortType = lgSTDescending Then
        Do
            While StrComp(MidValue, MyString(mRow(PtrRow(Low)).Cells(lSortColumn).Value), vbTextCompare) = -1
                Low = Low + 1
            Wend
            
            While StrComp(MyString(mRow(PtrRow(High)).Cells(lSortColumn).Value), MidValue, vbTextCompare) = -1
                High = High - 1
            Wend
            
            If Low <= High Then
                Item = PtrRow(Low)
                PtrRow(Low) = PtrRow(High)
                PtrRow(High) = Item
               
                Low = Low + 1
                High = High - 1
            End If
        Loop While Low <= High
    Else
        Do
            While StrComp(MyString(mRow(PtrRow(Low)).Cells(lSortColumn).Value), MidValue, vbTextCompare) = -1
                Low = Low + 1
            Wend
            
            While StrComp(MidValue, MyString(mRow(PtrRow(High)).Cells(lSortColumn).Value), vbTextCompare) = -1
                High = High - 1
            Wend
            
            If Low <= High Then
                Item = PtrRow(Low)
                PtrRow(Low) = PtrRow(High)
                PtrRow(High) = Item
               
                Low = Low + 1
                High = High - 1
            End If
        Loop While Low <= High
    End If
    If First < High Then QSortArrayString First, High, lSortColumn, nSortType
    If Low < Last Then QSortArrayString Low, Last, lSortColumn, nSortType
End Sub

Private Sub SortSubList()

   '// Purpose: Used to sort by a secondary Column after a Sort
  Dim lCount     As Long
  Dim lStartSort As Long
  Dim bDifferent As Boolean
  Dim sMajorSort As Variant

   If mSortSubColumn > C_NULL_RESULT Then
      '// Re-Sort the Items by a secondary column, preserving the sort sequence of the primary sort
      lStartSort = m_FixedRows

      For lCount = 0 To m_RowsCount - 1
         If Not IsNull(mRow(PtrRow(lCount)).Cells(mSortColumn).Value) Then
            bDifferent = Not (mRow(PtrRow(lCount)).Cells(mSortColumn).Value = sMajorSort)
         End If
         If bDifferent Or lCount = m_RowsCount - 1 Then
            If lCount > 1 Then
               If lCount - lStartSort > 1 Then
                  If lCount = m_RowsCount - 1 And Not bDifferent Then
                     SortArray lStartSort, lCount, mSortSubColumn, mCol(mSortSubColumn).nSortOrder
                  Else
                     SortArray lStartSort, lCount - 1, mSortSubColumn, mCol(mSortSubColumn).nSortOrder
                  End If

               End If
               lStartSort = lCount
            End If

            sMajorSort = mRow(PtrRow(lCount)).Cells(mSortColumn).Value
         End If

      Next lCount
   End If

End Sub

Private Sub SwapLng(ByRef Value1 As Long, ByRef Value2 As Long)

  Dim lTemp As Long

   lTemp = Value1
   Value1 = Value2
   Value2 = lTemp

End Sub

Private Function rVal(ByVal vString As String) As Double

   '// Returns the numbers contained in a string as a numeric value
   '// VB's Val function recognizes only the period (.) as a valid decimal separator.
   '// VB's CDbl errors on empty strings or values containing non-numeric values

  Dim lngI     As Long
  Dim lngS     As Long
  Dim bytAscV  As Byte
  Dim strTemp  As String
  
  On Error Resume Next

   vString = Trim$(UCase$(vString))
   If LenB(vString) Then
   
      Select Case Left$(vString, 2)          '// Hex or Octal?
      Case Is = "&H", Is = "&O"
         lngS = 3
         strTemp = Left$(vString, 2)
      Case Else
         lngS = 1
      End Select
      
      For lngI = lngS To Len(vString)
         bytAscV = AscW(Mid$(vString, lngI, 1))
         Select Case bytAscV
         Case 48 To 57, 69 '// 1234567890E
            strTemp = strTemp & Mid$(vString, lngI, 1)
         
         Case 44, 45, 46 '// , - .
            strTemp = strTemp & Mid$(vString, lngI, 1)
         
         Case 36, 163, 32 '// $
            '// Ignore
            
         Case Is > 57, Is < 44
            If Left$(strTemp, 2) = "&H" Then '// Hex Values ?
               Select Case bytAscV
               Case 65 To 70 '// ABCDEF
                  strTemp = strTemp & Mid$(vString, lngI, 1)
               Case Else
                  Exit For
               End Select
            Else
               Exit For
            End If
         End Select
      Next lngI
      
      If LenB(strTemp) Then
         rVal = CDbl(strTemp)
         If rVal = 0 Then
            strTemp = Replace$(strTemp, ".", ",")
            rVal = CDbl(strTemp)
         End If
      
      Else '// Check for boolean text (True or False)
         '// VB's CBool errors on empty or invalid strings (not True or False)
         '// Check for valid boolean
         rVal = CBool(vString)
      End If
   
   Else
      rVal = 0
   End If
   
Exit_Here:
   On Error GoTo 0

End Function

Public Sub InsertRow(Optional ByVal Before As Long = -1)
    Dim i As Long, lCol As Long

    ReDim Preserve mRow(m_RowsCount)
    ReDim Preserve PtrRow(m_RowsCount)
    ReDim Preserve mRow(m_RowsCount).Cells(m_ColsCount - 1)
    With mRow(m_RowsCount)
        .Height = m_RowsHeight
        .BackColor = CLR_NONE
        .ForeColor = CLR_NONE
        .Align = CenterLeft
    
        ReDim Preserve .Cells(m_ColsCount - 1)
        For lCol = 0 To m_ColsCount - 1
            With .Cells(lCol)
                .BackColor = CLR_NONE
                .ForeColor = CLR_NONE
                .Align = CenterLeft
            End With
        Next
    End With

    If Before <> -1 Then
        For i = m_RowsCount To Before Step -1
            PtrRow(i) = PtrRow(i - 1)
        Next
        PtrRow(Before - 1) = m_RowsCount
    Else
        PtrRow(m_RowsCount) = m_RowsCount
    End If
    
    m_RowsCount = m_RowsCount + 1
    
    If m_Redraw Then Me.Refresh
End Sub


Public Sub GroupByColumn(ByVal Col As Long, _
                        Optional ByVal BackColor As OLE_COLOR = vbButtonFace, _
                        Optional ByVal SubOrder As Long = 0, _
                        Optional ByVal RowHeight As Long = 0, _
                        Optional ByVal HiddenColumn As Boolean, _
                        Optional ByVal bExpanded As Boolean = True, _
                        Optional ByVal SortType As lgSortTypeEnum = lgSTAscending)
                        
    Dim lRow As Long
    Dim bRedraw As Boolean
    Dim Value As String
    '/> Jose Liza - ParentGroup
    Dim lRowParent As Long
    '/< Jose Liza - ParentGroup
        
    bRedraw = Me.Redraw
    Me.Redraw = False
    
    
    
    Sort PtrCol(Col), SortType

    
    Do While lRow < m_RowsCount
        
        If mRow(PtrRow(lRow)).Cells(PtrCol(Col)).Value <> Value And Not mRow(PtrRow(lRow)).IsGroup Then
            Value = mRow(PtrRow(lRow)).Cells(PtrCol(Col)).Value
            InsertRow lRow + 1
            RowIsGroup(lRow) = True
            With mRow(PtrRow(lRow))
                .Ident = SubOrder
                '/> Jose Liza - ParentGroup
                lRowParent = lRow
                '/< Jose Liza - ParentGroup
                .BackColor = BackColor
                .Cells(0).Value = StrConv(Format(Value, mCol(PtrCol(Col)).Format), vbProperCase)
                If RowHeight > 0 Then .Height = RowHeight
            End With
            If bExpanded = False Then
                GroupColapse lRow
            End If
        Else
            If Not mRow(PtrRow(lRow)).IsGroup Then
                mRow(PtrRow(lRow)).Ident = SubOrder + 2
                '/> Jose Liza - ParentGroup
                mRow(PtrRow(lRow)).RowParent = lRowParent
                '/< Jose Liza - ParentGroup
            Else
                Value = vbNullString
                
            End If
        End If
        lRow = lRow + 1
    Loop
    ColHidden(Col) = HiddenColumn
    Me.Redraw = bRedraw
End Sub

Public Sub UnGroup()
    Dim bRedraw As Boolean

    Dim i As Long
    bRedraw = Me.Redraw
    Me.Redraw = False
    For i = m_RowsCount - 1 To 0 Step -1
        If mRow(PtrRow(i)).IsGroup Then
            RowDelete i
        Else
            If mRow(PtrRow(i)).TempHeight > 0 Then RowHidden(i) = False
            mRow(PtrRow(i)).Ident = 0
        End If
    Next
    
    For i = 0 To m_ColsCount - 1
        ColHidden(i) = False
    Next
        
    Me.Redraw = bRedraw
    ucScrollbarV.Value = 0
    UserControl_Resize
End Sub
'*-

Public Function HeaderInitImgList(Optional ByVal ImageWidth As Integer = 16, Optional ByVal ImageHeight As Integer = 16) As Boolean
    m_HeaderImgLstWidth = ImageWidth * DpiF
    m_HeaderImgLstHeight = ImageHeight * DpiF
    HeaderImgListClear
    HeaderInitImgList = True
End Function

Public Property Get HeaderImgListWidth() As Long
   HeaderImgListWidth = m_HeaderImgLstWidth
End Property

Public Property Get HeaderImgListHeight() As Long
   HeaderImgListHeight = m_HeaderImgLstHeight
End Property

Public Function HeaderImgListAddImage(SrcImg As Variant, Optional Key As Variant, Optional Before, Optional After) As Long
    Dim hImage As Long
    hImage = LoadPictureEx(SrcImg, m_HeaderImgLstWidth, m_HeaderImgLstHeight)
    If hImage <> 0& Then
        cHeaderImageList.Add hImage, Key, Before, After
        HeaderImgListAddImage = cHeaderImageList.Count
    End If
End Function

Public Function HeaderImgListRemoveImage(Index As Variant) As Long
    GdipDisposeImage cHeaderImageList(Index)
    cHeaderImageList.Remove Index
    HeaderImgListRemoveImage = cHeaderImageList.Count
End Function

Public Sub HeaderImgListClear()
    Dim i As Long
    If Not cHeaderImageList Is Nothing Then
        For i = 1 To cHeaderImageList.Count
            GdipDisposeImage cHeaderImageList(i)
        Next
    End If
    Set cHeaderImageList = New Collection
End Sub

Public Property Get HeaderImgListCount() As Long
    HeaderImgListCount = cHeaderImageList.Count
End Property

Public Function ColInitImgList(ByVal Col As Long, Optional ByVal ImageWidth As Integer = 16, Optional ByVal ImageHeight As Integer = 16, Optional ByVal ImgAlign As eGridAlign, Optional ByVal Radius As Long, Optional ByVal ImagesMonocrome As Boolean) As Boolean
    With mCol(PtrCol(Col))
        .ImgListWidth = ImageWidth * DpiF
        .ImgListHeight = ImageHeight * DpiF
        .ImagesRadius = Radius
        .ImagesMonocrome = ImagesMonocrome
        .ImgAlign = ImgAlign
    End With
    ColImgListClear Col
    ColInitImgList = True
End Function

Public Property Get ColImgListWidth(ByVal Col As Long) As Long
   ColImgListWidth = mCol(PtrCol(Col)).ImgListWidth
End Property

Public Property Get ColImgListHeight(ByVal Col As Long) As Long
   ColImgListHeight = mCol(PtrCol(Col)).ImgListHeight
End Property

Public Function ColImgListAddImage(ByVal Col As Long, SrcImg As Variant, Optional Key, Optional Before, Optional After) As Long
    Dim hImage As Long
    hImage = LoadPictureEx(SrcImg, mCol(PtrCol(Col)).ImgListWidth, mCol(PtrCol(Col)).ImgListHeight, mCol(PtrCol(Col)).ImagesRadius)
    If hImage <> 0& Then
        mCol(PtrCol(Col)).ColImgList.Add hImage, Key, Before, After
        ColImgListAddImage = mCol(PtrCol(Col)).ColImgList.Count
    Else
        ColImgListAddImage = -1
    End If
End Function

Public Function ColImgListRemoveImage(ByVal Col As Long, Index As Variant) As Long
    GdipDisposeImage mCol(PtrCol(Col)).ColImgList.Item(Index)
    mCol(PtrCol(Col)).ColImgList.Remove Index
    ColImgListRemoveImage = mCol(PtrCol(Col)).ColImgList.Count
End Function

Public Sub ColImgListClear(ByVal Col As Long)
    Dim i As Long
    If Not mCol(PtrCol(Col)).ColImgList Is Nothing Then
        For i = 1 To mCol(PtrCol(Col)).ColImgList.Count
            GdipDisposeImage mCol(PtrCol(Col)).ColImgList(i)
        Next
    End If
    Set mCol(PtrCol(Col)).ColImgList = New Collection
End Sub

Public Property Get ColImageListCount(ByVal Col As Long) As Long
    ColImageListCount = mCol(PtrCol(Col)).ColImgList.Count
End Property

Private Sub RoundRectPlus(ByVal hdc As Long, ByVal Left As Long, ByVal Top As Long, _
                       ByVal Width As Long, ByVal Height As Long, ByVal BackColor As Long, ByVal BorderColor As Long, Radius As Long)
    Dim hPen As Long, hBrush As Long
    Dim mPath As Long, hGraphics As Long
    
    GdipCreateFromHDC hdc, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    
    If BackColor <> 0 Then GdipCreateSolidFill BackColor, hBrush
    If BorderColor <> 0 Then GdipCreatePen1 BorderColor, 1 * DpiF, &H2, hPen
    
    If Radius = 0 Then
      
        If hBrush Then GdipFillRectangleI hGraphics, hBrush, Left, Top, Width, Height
        If hPen Then GdipDrawRectangleI hGraphics, hPen, Left, Top, Width, Height
        
    Else
        If GdipCreatePath(&H0, mPath) = 0 Then
            
            GdipAddPathArcI mPath, Left, Top, Radius, Radius, 180, 90
            GdipAddPathArcI mPath, Left + Width - Radius, Top, Radius, Radius, 270, 90
            GdipAddPathArcI mPath, Left + Width - Radius, Top + Height - Radius, Radius, Radius, 0, 90
            GdipAddPathArcI mPath, Left, Top + Height - Radius, Radius, Radius, 90, 90
            GdipClosePathFigure mPath
            
            If hBrush Then GdipFillPath hGraphics, hBrush, mPath
            If hPen Then GdipDrawPath hGraphics, hPen, mPath
            Call GdipDeletePath(mPath)
        End If
    End If
        
    If hBrush Then Call GdipDeleteBrush(hBrush)
    If hPen Then Call GdipDeletePen(hPen)
    GdipDeleteGraphics hGraphics
End Sub

Private Function LoadPictureEx(SrcImg As Variant, Optional ByVal Width As Long, Optional ByVal Height As Long, Optional ByVal ImagesRadius As Integer) As Long
    Dim hImage1 As Long, hImage2 As Long, hGraphics As Long
    Dim DataArr() As Byte
    Dim lPictureRealWidth As Long, lPictureRealHeight As Long
    Dim X As Long, Y As Long, cx As Long, cy As Long
    Dim sngRatio1 As Single, sngRatio2 As Single
    Dim mPath As Long, Radius As Long, hPen As Long
    
    Select Case VarType(SrcImg)
        Case vbString
            If PathIsURL(SrcImg) Then
            
                If Left$(LCase(SrcImg), 5) = "data:" Then
                    Base64Decode Split(SrcImg, ",")(1), DataArr
                    Call LoadImageFromArray(DataArr, hImage1)
                Else
                    Dim oXMLHTTP As Object
                    Set oXMLHTTP = CreateObject("Microsoft.XMLHTTP")
                    
                    oXMLHTTP.Open "GET", SrcImg, True
                    oXMLHTTP.send
                    While oXMLHTTP.readyState <> 4
                        DoEvents
                    Wend
                    If oXMLHTTP.Status = 200 Then
                        DataArr() = oXMLHTTP.responseBody
                        Call LoadImageFromArray(DataArr, hImage1)
                    End If
                End If
            Else
                Call GdipLoadImageFromFile(ByVal StrPtr(SrcImg), hImage1)
            End If
        Case vbLong
            Dim hBmp As Long
            Dim IIF As ICONINFO
            Dim tBmp As BITMAP
            
            If GetObjectType(SrcImg) = OBJ_BITMAP Then
                If GetObject(SrcImg, Len(tBmp), tBmp) Then
                    
                    If tBmp.bmBitsPixel = 32 And tBmp.bmBits > 0 Then
                        GdipCreateBitmapFromScan0 tBmp.bmWidth, tBmp.bmHeight, tBmp.bmWidthBytes, PixelFormat32bppPARGB, tBmp.bmBits, hImage1
                        GdipImageRotateFlip hImage1, RotateNoneFlipY
                    Else
                        Call GdipCreateBitmapFromHBITMAP(SrcImg, 0, hImage1)
                    End If
                End If
            Else
                If TypeName(SrcImg) = "Long" Then
                    GetIconInfo SrcImg, IIF
                    hBmp = CopyImage(IIF.hbmColor, IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION Or &H2)
                    Call GetObject(hBmp, Len(tBmp), tBmp)
                    GdipCreateBitmapFromScan0 tBmp.bmWidth, tBmp.bmHeight, tBmp.bmWidthBytes, PixelFormat32bppPARGB, tBmp.bmBits, hImage1
                    GdipImageRotateFlip hImage1, RotateNoneFlipY
                    DeleteObject hBmp
                    DeleteObject IIF.hbmColor
                    DeleteObject IIF.hbmMask
                Else
                    GdipCreateBitmapFromHICON SrcImg, hImage1
                End If
            End If
            
        Case vbDataObject
            Call GdipLoadImageFromStream(SrcImg, hImage1)
            
        Case (vbArray Or vbByte)
            DataArr() = SrcImg
            Call LoadImageFromArray(DataArr, hImage1)
    End Select

    If hImage1 <> 0 Then
        GdipGetImageWidth hImage1, lPictureRealWidth
        GdipGetImageHeight hImage1, lPictureRealHeight
        If Width = 0 Then Width = lPictureRealWidth
        If Height = 0 Then Height = lPictureRealHeight
        
        sngRatio1 = Width / lPictureRealWidth
        sngRatio2 = Height / lPictureRealHeight
        If sngRatio1 > sngRatio2 Then sngRatio1 = sngRatio2
        cx = lPictureRealWidth * sngRatio1: cy = lPictureRealHeight * sngRatio1
        X = (Width - cx) \ 2: Y = (Height - cy) \ 2

        GdipCreateBitmapFromScan0 Width, Height, 0&, PixelFormat32bppPARGB, ByVal 0&, hImage2
        GdipGetImageGraphicsContext hImage2, hGraphics
        GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
        
        If ImagesRadius > 0 Then
            Radius = ImagesRadius * DpiF
            If GdipCreatePath(&H0, mPath) = 0 Then
                X = X + DpiF: Y = Y + DpiF
                cx = cx - DpiF * 2: cy = cy - DpiF * 2
                GdipAddPathArcI mPath, X, Y, Radius, Radius, 180, 90
                GdipAddPathArcI mPath, X + cx - Radius, Y, Radius, Radius, 270, 90
                GdipAddPathArcI mPath, X + cx - Radius, Y + cy - Radius, Radius, Radius, 0, 90
                GdipAddPathArcI mPath, X, Y + cy - Radius, Radius, Radius, 90, 90
                GdipClosePathFigure mPath
                GdipSetClipPath hGraphics, mPath, CombineModeIntersect
            End If
        End If
        
        GdipDrawImageRectRectI hGraphics, hImage1, X, Y, cx, cy, 0, 0, lPictureRealWidth, lPictureRealHeight, UnitPixel, 0&, 0&, 0&
        
        If mPath Then
            GdipCreatePen1 RGBtoARGB(vbButtonFace, 100), 1 * DpiF, &H2, hPen
            GdipResetClip hGraphics
            GdipDrawPath hGraphics, hPen, mPath
            GdipDeletePen hPen
            GdipDeletePath mPath
        End If
        
        GdipDeleteGraphics hGraphics
        GdipDisposeImage hImage1
        LoadPictureEx = hImage2
    End If
End Function

Private Function ImageListDraw(ByVal hImage As Long, ByVal hdc As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, Optional PictureColor As Long = CLR_NONE, Optional Disabled As Boolean) As Boolean
    Dim tMatrixColor As COLORMATRIX, tMatrixGray As COLORMATRIX
    Dim hAttributes As Long, hGraphics As Long

    Call GdipCreateImageAttributes(hAttributes)
    
    With tMatrixColor
        If PictureColor <> CLR_NONE Then
            If (PictureColor And &H80000000) Then PictureColor = GetSysColor(PictureColor And &HFF&)
            Dim R As Byte, G As Byte, b As Byte

            b = ((PictureColor \ &H10000) And &HFF)
            G = ((PictureColor \ &H100) And &HFF)
            R = (PictureColor And &HFF)

            .M(0, 0) = R / 255
            .M(1, 0) = G / 255
            .M(2, 0) = b / 255
            .M(0, 4) = R / 255
            .M(1, 4) = G / 255
            .M(2, 4) = b / 255
        Else
            .M(0, 0) = 1
            .M(1, 1) = 1
            .M(2, 2) = 1
        End If
        .M(3, 3) = 1
        .M(4, 4) = 1

        If Disabled Then
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
    
    GdipCreateFromHDC hdc, hGraphics
    GdipSetImageAttributesColorMatrix hAttributes, &H0, True, tMatrixColor, tMatrixGray, &H0
    ImageListDraw = GdipDrawImageRectRectI(hGraphics, hImage, Left, Top, Width, Height, 0, 0, Width, Height, UnitPixel, hAttributes) = 0&
    Call GdipDeleteGraphics(hGraphics)
    Call GdipDisposeImageAttributes(hAttributes)
End Function

Private Function LoadImageFromArray(ByRef bvData() As Byte, ByRef hImage As Long) As Boolean
    On Local Error GoTo LoadImageFromArray_Error
    Dim IStream     As IUnknown
    If Not IsArrayDim(VarPtrArray(bvData)) Then Exit Function
    
    Call CreateStreamOnHGlobal(bvData(0), 0&, IStream)
    If Not IStream Is Nothing Then
        If GdipLoadImageFromStream(IStream, hImage) = 0 Then
            LoadImageFromArray = True
        End If
    End If
    Set IStream = Nothing
    
LoadImageFromArray_Error:
End Function

Private Function Base64Decode(ByVal sIn As String, ByRef bvOut() As Byte) As Boolean 'By Cocus
    Dim lLenOut As Long
    Const CRYPT_STRING_BASE64 = &H1
    Call CryptStringToBinaryA(sIn, Len(sIn), CRYPT_STRING_BASE64, 0, VarPtr(lLenOut), 0, 0)
    If lLenOut = 0 Then Exit Function
    ReDim bvOut(lLenOut - 1)
    Call CryptStringToBinaryA(sIn, Len(sIn), CRYPT_STRING_BASE64, VarPtr(bvOut(0)), VarPtr(lLenOut), 0, 0)
    Base64Decode = True
End Function

Private Function IsArrayDim(ByVal lpArray As Long) As Boolean
    Dim lAddress As Long
    Call CopyMemory(lAddress, ByVal lpArray, &H4)
    IsArrayDim = Not (lAddress = 0)
End Function

Public Function RGBtoARGB(ByVal RGBColor As Long, ByVal Opacity As Long) As Long 'By LaVople
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

'/>Jose Liza - FocusRect
Private Sub UpdateFocusRect(CellPoint As POINTAPI, Orientation As CellOrientationFocus)
    If m_FocusRectMode <> lgNone Then
        If Orientation = All Then
            m_CellFocusRect.X = CellPoint.X
            m_CellFocusRect.Y = CellPoint.Y
        ElseIf Orientation = Horizontal Then
            m_CellFocusRect.X = CellPoint.X
        ElseIf Orientation = Vertical Then
            m_CellFocusRect.Y = CellPoint.Y
        End If
    End If
End Sub

Private Sub DrawFocusCell(hdc As Long, cRect As Rect, ByVal Color As OLE_COLOR, Optional BorderWidth As Integer = 1)
    'Dim hBrush As Long
    Dim hPen As Long
    Dim hOldPen As Long
    '---
    If (Color And &H80000000) Then Color = GetSysColor(Color And &HFF&)
    
    'hBrush = CreateSolidBrush(Color)   'Reemplazado por hPen
    'FrameRect hdc, cRect, hBrush        'Reemplazado por Rectangle
    'Call DeleteObject(hBrush)
    
    hPen = CreatePen(PS_SOLID, BorderWidth, Color)
    hOldPen = SelectObject(hdc, hPen)
    Rectangle hdc, cRect.Left, cRect.Top, cRect.Right, cRect.Bottom
    SelectObject hdc, hOldPen
    DeleteObject hPen
End Sub
'/>Jose Liza - FocusRect
