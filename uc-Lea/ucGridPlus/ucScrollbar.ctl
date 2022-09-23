VERSION 5.00
Begin VB.UserControl ucScrollbar 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   CanGetFocus     =   0   'False
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1260
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H8000000F&
   ScaleHeight     =   177
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   84
   ToolboxBitmap   =   "ucScrollbar.ctx":0000
   Begin VB.Timer timHideTooltip 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   780
      Top             =   270
   End
   Begin VB.Timer TimSmoothChange 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   30
      Top             =   1890
   End
End
Attribute VB_Name = "ucScrollbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================================================================
' User control:  ucScrollbar.ctl
' Author:        Carles P.V. - 2005 (*)
' Dependencies:  None
' Last revision: 22/07/21
' Version:       1.0.15 (Shagratt 2021)
' Version:       1.0.5 (Jason James Newland 2007)
' Version:       1.0.4 (Carles P.V 2005)
'----------------------------------------------------------------------------------------
'
' (*) 1. Self-Subclassing UserControl template (IDE safe) by Paul Caton:
'
'        Self-subclassing Controls/Forms - NO dependencies
'        http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=54117&lngWId=1
'
'     2. pvCheckEnvironment() and pvIsLuna() routines by Paul Caton
'
'     3. Flat button fxs code extracted from (see pvDrawFlatButton() routine):
'        Special flat Cool Scrollbars version 1.2 by James Brown
'        http://www.catch22.net/tuts/coolscroll.asp
'----------------------------------------------------------------------------------------
'
' History:
'
'   * 1.0.0: - First release.
'   * 1.0.1: - Flat style *properly* painted:
'              * Hot thumb appearance = Pressed thumb appearance.
'              * Pressed/hot buttons using correct system colors.
'              Is there a default? For example, ListView with flat-scrollbars flag set,
'              preserves pressed buttons with 1-pixel edge using 'shadow' color and
'              their background is filled using color black instead of 'dark shadow'.
'   * 1.0.2: - Added Refresh method: only for custom-draw purposes.
'   * 1.0.3: - Fixed control on m_bHasTrack and m_bHasNullTrack flags.
'   * 1.0.4: - Fixed thumb rendering (classic style). DrawFrameControl->DrawEdge.
'   * 1.0.5: - Added theme support for Vista in theme mode style (JJN)
'   * 1.0.6: (Shagratt)
'            - Added MouseWhell support on Hover (no focus needed)
'            - Added MouseWhell support over another object (on demand)
'            - Added Public Methods to create object on demand (by code)
'            - Added Smooth Scrolling (can be adjusted.value=1 Disabled)
'            - MouseWheel scroll value can be adjusted
'            - Changed Thumb size formula
'            - Continous TL/BR button scrolling
'            - Thumb scrolling raise Scroll and Change events
'            - Changed default props: Flat, scroll rates,etc.
'            - Changed default props: Flat, scroll rates,etc.
'            - Updated to Paul Caton Self Sub v1.6
'   * 1.0.7: - Added Horizontal Scroll with Shift Pressed (.AttachHorizontalScrollBar() )
'              Thanks dseaman@vbforums for the idea!
'   * 1.0.7b:- Changed dragging thumb to smooth scroll instead of direct to value (smooth drawing too)
'   * 1.0.8: - Changed dragging the thumb is direct visually (only value change smooth)
'            - Fixed lot of errors on thumb draws/calculation introduced in 1.0.7b
'   * 1.0.9: - Fixed Themed Style
'            - Change Themed hover on arrows to emphasize both buttons
'   * 1.0.9b:- Fixed thumb formula when value to scroll was too small (finally!)
'            - Exposed ZOrder
'   * v1.0.9c (28/02/21)
'           - Fixed usScrollbar unsubclassing (crash on form unload from a button)
'   * v1.0.9d (05/07/21)
'           - Changing largChange also change Wheel Scroll distance (you can override it
'             changing "WheelChange" later
'   * v1.0.10 (20/07/21)
'           - Added ThumbTooltip:
'              When dragged from the thumb it will call the event
'              showThumbTooltip(byref ThumbTooltipText) were the tooltip text can be set.
'             property ThumbTooltipEnable: Enable/Disable this function
'             property ThumbTooltipFont: Allow to change the font to use
'   * v1.0.11 (22/07/21)
'           - Added soft shadow to the ThumbTooltip (removed, caused problems)
'           - Added "Google" style (created by Leandro Ascierto)
'           - "Google" Style Radio width, background and thumb color configurable (-1 = auto)
'           - Fixed Instant scroll when scrollbar was still moving
'   * v1.0.12 (29/08/21)
'           - Added error controls on SubclassingStart and SubclassEnd
'           - Changed UserControl.Parent.hwnd for UserControl.ContainerHwnd
'   *v1.0.13 (09/09/21)
'           - Included AMV changes on pvGetThumbSize() and pvGetThumbPos()
'           - Can take int values for SmoothScrollFactor (Values >1 are divided by 100)
'   *v1.0.14 (09/09/21)
'           - Included Javier Balkenende changes on pvGetThumbSize()
'           - Thumb Min/Max size customizable between 0-100% (Never be less than 8px)
'   *v1.0.15 (12/09/21)
'           - Included Leandro Ascierto changes to Add Horizontal scroll with mousewheel over
'             Horizontal scrollbar
'           - Fixed Left/Right keys for scroll movement
'           - Forward KeyDown,KeyPress and KeyUp events
'           - Changed how ThumbTooltip event work. Now you pass the textObject from the Form/UC
'             receiving the event. This allow to maintaing the Scrollbar with CanGetFocus=False
'----------------------------------------------------------------------------------------
'
' Notes:
'
'   * Restriction: Max >= Min
'----------------------------------------------------------------------------------------
'
' Known issues:
'   - TabStop not working: Solution Set CanGetFocus=True
'   - If you modify the scrollbar with "CanGetFocus=true" and use it inside another UC the
'      KeyPreview will not work (you dont Key Events:  KeyUp,KeyPress,KeyDown)
'     With this code you can get the focus back in your usercontrol and get them.
'
'        'Hack to recover focus from the scrollbars
'        Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'            UCScrollX.Enabled = False: UCScrollY.Enabled = False
'            DoEvents
'            UserControl.SetFocus
'            UCScrollX.Enabled = True: UCScrollY.Enabled = True
'        End Sub
'
'
'
'========================================================================================



Option Explicit

Private Const VERSION_INFO As String = "v1.0.15"

'-Selfsub declarations----------------------------------------------------------------------------
Private Enum eMsgWhen                                                       'When to callback
  MSG_BEFORE = 1                                                            'Callback before the original WndProc
  MSG_AFTER = 2                                                             'Callback after the original WndProc
  MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                                'Callback before and after the original WndProc
End Enum


Private Type tSubData                                                         'Subclass data type
    hwnd                   As Long                                            'Handle of the window being subclassed
    nAddrSub               As Long                                            'The address of our new WndProc (allocated memory).
    nAddrOrig              As Long                                            'The address of the pre-existing WndProc
    nMsgCntA               As Long                                            'Msg after table entry count
    nMsgCntB               As Long                                            'Msg before table entry count
    aMsgTblA()             As Long                                            'Msg after table array
    aMsgTblB()             As Long                                            'Msg Before table array
End Type

Private Const ALL_MESSAGES  As Long = -1                                    'All messages callback
Private Const MSG_ENTRIES   As Long = 32                                    'Number of msg table entries
Private Const WNDPROC_OFF   As Long = &H38                                  'Thunk offset to the WndProc execution address
Private Const GWL_WNDPROC   As Long = -4                                    'SetWindowsLong WndProc index
Private Const IDX_SHUTDOWN  As Long = 1                                     'Thunk data index of the shutdown flag
Private Const IDX_HWND      As Long = 2                                     'Thunk data index of the subclassed hWnd
Private Const IDX_WNDPROC   As Long = 9                                     'Thunk data index of the original WndProc
Private Const IDX_BTABLE    As Long = 11                                    'Thunk data index of the Before table
Private Const IDX_ATABLE    As Long = 12                                    'Thunk data index of the After table
Private Const IDX_PARM_USER As Long = 13                                    'Thunk data index of the User-defined callback parameter data index

Private z_ScMem             As Long                                         'Thunk base address
Private z_Sc(64)            As Long                                         'Thunk machine-code initialised here
Private z_Funk              As Collection                                   'hWnd/thunk-address collection

Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
'Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

'========================================================================================
' UserControl API declarations
'========================================================================================
Private Const SM_CXVSCROLL  As Long = 2
Private Const SM_CYHSCROLL  As Long = 3
Private Const SM_CYVSCROLL  As Long = 20
Private Const SM_CXHSCROLL  As Long = 21
Private Const SM_SWAPBUTTON As Long = 23

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SPI_GETKEYBOARDDELAY As Long = 22
Private Const SPI_GETKEYBOARDPREF  As Long = 68

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Type Rect
    X1 As Long
    Y1 As Long
    X2 As Long
    Y2 As Long
End Type

Private Declare Function SetRect Lib "user32" (lpRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As Rect) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As Rect, lpSourceRect As Rect) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As Rect, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As Rect, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As Rect) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As Rect, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
'Shagratt 29/11/19
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As KeyCodeConstants) As Integer
Private Const KEY_DOWN As Integer = &H8000

Private ExtWheelHwnd&

Private Const DFC_SCROLL          As Long = 3
Private Const DFCS_SCROLLUP       As Long = &H0
Private Const DFCS_SCROLLDOWN     As Long = &H1
Private Const DFCS_SCROLLLEFT     As Long = &H2
Private Const DFCS_SCROLLRIGHT    As Long = &H3
Private Const DFCS_INACTIVE       As Long = &H100
Private Const DFCS_PUSHED         As Long = &H200
Private Const DFCS_FLAT           As Long = &H4000
Private Const DFCS_MONO           As Long = &H8000

Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As Rect, ByVal un1 As Long, ByVal un2 As Long) As Long

Private Const BDR_RAISED As Long = &H5
Private Const BF_RECT    As Long = &HF

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Long

Private Const COLOR_BTNFACE     As Long = 15
Private Const COLOR_3DSHADOW    As Long = 16
Private Const COLOR_BTNTEXT     As Long = 18
Private Const COLOR_3DHIGHLIGHT As Long = 20
Private Const COLOR_3DDKSHADOW  As Long = 21

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function SetTextColor Lib "GDI32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "GDI32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private Const WHITE_BRUSH As Long = 0
Private Const BLACK_BRUSH As Long = 4

Private Declare Function GetStockObject Lib "GDI32" (ByVal nIndex As Long) As Long
    
Private Const MOUSEEVENTF_LEFTDOWN As Long = &H2

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateBitmap Lib "GDI32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Integer) As Long
Private Declare Function CreatePatternBrush Lib "GDI32" (ByVal hBitmap As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hdc As Long) As Long
 
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Private Type PAINTSTRUCT
    hdc             As Long
    fErase          As Long
    rcPaint         As Rect
    fRestore        As Long
    fIncUpdate      As Long
    rgbReserved(32) As Byte
End Type
Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENTTYPE) As Long

Private Const WM_SIZE           As Long = &H5
Private Const WM_PAINT          As Long = &HF
Private Const WM_SYSCOLORCHANGE As Long = &H15
Private Const WM_CANCELMODE     As Long = &H1F
Private Const WM_TIMER          As Long = &H113
Private Const WM_MOUSEMOVE      As Long = &H200
Private Const WM_LBUTTONDOWN    As Long = &H201
Private Const WM_LBUTTONUP      As Long = &H202
Private Const WM_LBUTTONDBLCLK  As Long = &H203
Private Const WM_THEMECHANGED   As Long = &H31A
Private Const MK_LBUTTON        As Long = &H1
Private Const WM_MOUSEWHEEL     As Long = &H20A
Private Const WM_KEYDOWN        As Long = &H100
Private Const WM_MOUSELEAVE     As Long = &H2A3

Private Const TME_LEAVE         As Long = &H2&

Private Type TRACKMOUSEEVENTTYPE
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function GetCurrentThemeName Lib "uxtheme" (ByVal pszThemeFileName As Long, ByVal cchMaxNameChars As Long, ByVal pszColorBuff As Long, ByVal cchMaxColorChars As Long, ByVal pszSizeBuff As Long, ByVal cchMaxSizeChars As Long) As Long
Private Declare Function GetThemeDocumentationProperty Lib "uxtheme" (ByVal pszThemeName As Long, ByVal pszPropertyName As Long, ByVal pszValueBuff As Long, ByVal cchMaxValChars As Long) As Long
Private Declare Function OpenThemeData Lib "uxtheme" (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As Rect, pClipRect As Rect) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_EXSTYLE As Long = (-20)
Private Const GWL_STYLE As Long = (-16)
Private Const WS_THICKFRAME = &H40000
Private Const WS_EX_TOOLWINDOW = &H80

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const HWND_TOPMOST              As Long = -1
Private Const HWND_NOTOPMOST            As Long = -2
Private Const SWP_NOSIZE                As Long = &H1
Private Const SWP_NOMOVE                As Long = &H2
Private Const SWP_NOZORDER              As Long = &H4
Private Const SWP_NOACTIVATE            As Long = &H10
Private Const SWP_FRAMECHANGED          As Long = &H20
Private Const SWP_SHOWWINDOW            As Long = &H40
Private Const SWP_HIDEWINDOW            As Long = &H80
Private Const SWP_ASYNCWINDOWPOS        As Long = &H4000

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
'*2
'TEST
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Type GDIPlusStartupInput
    GdiPlusVersion                      As Long
    DebugEventCallback                  As Long
    SuppressBackgroundThread            As Long
    SuppressExternalCodecs              As Long
End Type
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function CreatePen Lib "GDI32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal Brush As Long) As Long
Private Declare Function GdipClosePathFigure Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, ByRef Brush As Long) As Long
Private Declare Function GdipFillRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipDrawRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long

Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private GdipToken&
Private DpiFactor!



' Class name
Private Const SB_THEME As String = "Scrollbar"

' [UxThemeSCROLLBARParts]
Private Const SBP_ARROWBTN = 1
Private Const SBP_THUMBBTNHORZ = 2
Private Const SBP_THUMBBTNVERT = 3
Private Const SBP_LOWERTRACKHORZ = 4
Private Const SBP_UPPERTRACKHORZ = 5
Private Const SBP_LOWERTRACKVERT = 6
Private Const SBP_UPPERTRACKVERT = 7
Private Const SBP_GRIPPERHORZ = 8
Private Const SBP_GRIPPERVERT = 9
Private Const SBP_SIZEBOX = 10

' [UxThemeARROWBTNStates]
Private Const ABS_UPNORMAL = 1
Private Const ABS_UPHOT = 2
Private Const ABS_UPPRESSED = 3
Private Const ABS_UPDISABLED = 4
Private Const ABS_DOWNNORMAL = 5
Private Const ABS_DOWNHOT = 6
Private Const ABS_DOWNPRESSED = 7
Private Const ABS_DOWNDISABLED = 8
Private Const ABS_LEFTNORMAL = 9
Private Const ABS_LEFTHOT = 10
Private Const ABS_LEFTPRESSED = 11
Private Const ABS_LEFTDISABLED = 12
Private Const ABS_RIGHTNORMAL = 13
Private Const ABS_RIGHTHOT = 14
Private Const ABS_RIGHTPRESSED = 15
Private Const ABS_RIGHTDISABLED = 16

' [UxThemeHorzScrollStates]
Private Const HSS_NORMAL = 1
Private Const HSS_HOT = 2
Private Const HSS_PUSHED = 3
Private Const HSS_DISABLED = 4

' [UxThemeHorzThumbStates]
Private Const HTS_NORMAL = 1
Private Const HTS_HOT = 2
Private Const HTS_PUSHED = 3
Private Const HTS_DISABLED = 4

' [UxThemeVertScrollStates]
Private Const VSS_NORMAL = 1
Private Const VSS_HOT = 2
Private Const VSS_PUSHED = 3
Private Const VSS_DISABLED = 4

' [UxThemeVertThumbStates]
Private Const VTS_NORMAL = 1
Private Const VTS_HOT = 2
Private Const VTS_PUSHED = 3
Private Const VTS_DISABLED = 4

'========================================================================================
' UserControl enums., variables and constants
'========================================================================================
'-- Public enums.:
Public Enum sbOrientationCts
    [oVertical] = 0
    [oHorizontal] = 1
End Enum

Public Enum sbStyleCts
    [sClassic] = 0
    [sFlat] = 1
    [sThemed] = 2
    [sCustomDraw] = 3
    [sGoogle] = 4
End Enum

Public Enum sbOnPaintPartCts
    [ppTLButton] = 0
    [ppBRButton] = 1
    [ppTLTrack] = 2
    [ppBRTrack] = 3
    [ppNullTrack] = 4
    [ppThumb] = 5
End Enum

Public Enum sbOnPaintPartStateCts
    [ppsNormal] = 0
    [ppsPressed] = 1
    [ppsHot] = 2
    [ppsDisabled] = 3
End Enum

'-- Private enums.:
Private Enum eFlatButtonStateCts
    [fbsNormal] = 0
    [fbsSelected] = 1
    [fbsHot] = 2
End Enum

'-- Private constants:
Private Const HT_NOTHING          As Long = 0
Private Const HT_TLBUTTON         As Long = 1
Private Const HT_BRBUTTON         As Long = 2
Private Const HT_TLTRACK          As Long = 3
Private Const HT_BRTRACK          As Long = 4
Private Const HT_THUMB            As Long = 5

Private Const TIMERID_CHANGE1     As Long = 1
Private Const TIMERID_CHANGE2     As Long = 2
Private Const TIMERID_HOT         As Long = 3

Private Const CHANGEDELAY_MIN     As Long = 0
Private Const CHANGEFREQUENCY_MIN As Long = 25
Private Const TIMERDT_HOT         As Long = 25

Private Const Thumbsize_AbsMin    As Long = 8
Private Const GRIPPERSIZE_MIN     As Long = 16

'-- Private variables:
Private m_bHasTrack               As Boolean
Private m_bHasNullTrack           As Boolean
Private m_uRctNullTrack           As Rect

Private m_uRctTLButton            As Rect
Private m_uRctBRButton            As Rect
Private m_uRctTLTrack             As Rect
Private m_uRctBRTrack             As Rect
Private m_uRctThumb               As Rect
Private m_lThumbOffset            As Long
Private m_uRctDrag                As Rect

Private m_bTLButtonPressed        As Boolean
Private m_bBRButtonPressed        As Boolean
Private m_bTLTrackPressed         As Boolean
Private m_bBRTrackPressed         As Boolean
Private m_bThumbPressed           As Boolean

Private m_bTLButtonHot            As Boolean
Private m_bBRButtonHot            As Boolean
Private m_bThumbHot               As Boolean

Private m_lAbsRange               As Long
Private m_lThumbPos               As Long
Private m_lThumbSize              As Long
Private m_eHitTest                As Long
Private m_eHitTestHot             As Long
Private m_x                       As Long
Private m_y                       As Long
Private m_lValueStartDrag         As Long

Private m_hPatternBrush           As Long

'-- Property variables:
Private m_lChangeDelay            As Long
Private m_lChangeFrequency        As Long
Private m_lMax                    As Long
Private m_lMin                    As Long
Private m_lValue                  As Long
Private m_lSmallChange            As Long
Private m_lLargeChange            As Long
Private m_eOrientation            As sbOrientationCts
Private m_eStyle                  As sbStyleCts
Private m_bShowButtons            As Boolean

Private m_bIsXP                   As Boolean ' RO
Private m_bIsLuna                 As Boolean ' RO

' Variable to hold 'DisableMouseWheelSupport' property value
Private m_bDisableMouseWheelSupport As Boolean
Private m_TargetValue As Long 'For SmoothScroll
Private m_SinSmoothScrollFactor As Single 'For SmoothScroll
' Variable to hold 'WheelChange' property value
Private m_LonWheelChange As Long
Private m_ucScrollbarH As ucScrollbar

Private bReadPropertiesDone As Boolean 'to not process for a second time
Private bSubclassed As Boolean
Private txtToolTip As Object

'-- Default property values:
Private Const ENABLED_DEF         As Boolean = True
Private Const MIN_DEF             As Long = 0
Private Const MAX_DEF             As Long = 100
Private Const VALUE_DEF           As Long = MIN_DEF
Private Const SMALLCHANGE_DEF     As Long = 10 '1
Private Const LARGECHANGE_DEF     As Long = 40 '10
Private Const CHANGEDELAY_DEF     As Long = 250
Private Const CHANGEFREQUENCY_DEF As Long = 75
Private Const ORIENTATION_DEF     As Long = [oVertical]
Private Const STYLE_DEF           As Long = [sFlat]
Private Const SHOWBUTTONS_DEF     As Boolean = True

'-- Events:
Public Event Change()
Public Event Scroll()
Public Event ThemeChanged()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OnPaint(ByVal lhDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal ePart As sbOnPaintPartCts, ByVal estate As sbOnPaintPartStateCts)
Public Event showThumbTooltip(ByRef ThumbTooltipText As String, ByRef TxtControl As Object)
'Forwarded events
Public Event KeyDown(KeyCode As Integer, Shift As Integer, ByRef DontProcess As Boolean)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event ContainerMouseLeave()
Public Event ContainerMouseEnter()

Private m_bThumbTooltipEnable As Boolean
Private m_StdThumbTooltipFont As StdFont
Private m_StyleThumbColor As OLE_COLOR
Private m_StyleBackColor As OLE_COLOR
Private m_StyleCurveRadius As Long
Private m_Thumbsize_min As Long
Private m_Thumbsize_max As Long
Private m_bMouseIn As Boolean
Public Property Get Thumbsize_max() As Long
Attribute Thumbsize_max.VB_Description = "Maximum size in % of the Thumb (0-100%)"
    Thumbsize_max = m_Thumbsize_max
End Property
Public Property Let Thumbsize_max(ByVal LonValue As Long)
    If (LonValue < 0) Then LonValue = 0
    If (LonValue > 100) Then LonValue = 100
    m_Thumbsize_max = LonValue
    PropertyChanged "Thumbsize_max"
End Property

Public Property Get Thumbsize_min() As Long
Attribute Thumbsize_min.VB_Description = "Minumum size in % of the Thumb (0-100%). Never be less than 8px."
    Thumbsize_min = m_Thumbsize_min
End Property
Public Property Let Thumbsize_min(ByVal LonValue As Long)
    If (LonValue < 0) Then LonValue = 0
    If (LonValue > 100) Then LonValue = 100
    m_Thumbsize_min = LonValue
    PropertyChanged "Thumbsize_min"
End Property

Public Property Get StyleCurveRadius() As Long
Attribute StyleCurveRadius.VB_Description = "(For Style 4) -1 (FFFFFFFF) = auto"
    StyleCurveRadius = m_StyleCurveRadius
End Property
Public Property Let StyleCurveRadius(ByVal LonValue As Long)
    m_StyleCurveRadius = LonValue
    Me.Refresh
    PropertyChanged "StyleCurveRadius"
End Property

Public Property Get StyleBackColor() As OLE_COLOR
Attribute StyleBackColor.VB_Description = "(For Style 4)  -1 (FFFFFFFF) = take color from container. "
    StyleBackColor = m_StyleBackColor
End Property
Public Property Let StyleBackColor(ByVal OLEValue As OLE_COLOR)
    m_StyleBackColor = OLEValue
    If (m_StyleBackColor = -1) Then
        UserControl.BackColor = Ambient.BackColor ' UserControl.Parent.BackColor
    Else
        UserControl.BackColor = OLEValue
    End If
    Me.Refresh
    PropertyChanged "StyleBackColor"
End Property

Public Property Get StyleThumbColor() As OLE_COLOR
Attribute StyleThumbColor.VB_Description = "(For Style 4) -1 (FFFFFFFF) = auto"
    StyleThumbColor = m_StyleThumbColor
End Property
Public Property Let StyleThumbColor(ByVal OLEValue As OLE_COLOR)
    m_StyleThumbColor = OLEValue
    Me.Refresh
    PropertyChanged "StyleThumbColor"
End Property


Public Property Get ThumbTooltipFont() As StdFont
    Set ThumbTooltipFont = m_StdThumbTooltipFont
End Property
Public Property Set ThumbTooltipFont(ByVal StdValue As StdFont)
    Set m_StdThumbTooltipFont = StdValue
    
    If Not (txtToolTip Is Nothing) Then Set txtToolTip.Font = m_StdThumbTooltipFont
    Set UserControl.Font = m_StdThumbTooltipFont
    
    PropertyChanged "ThumbTooltipFont"
End Property




'========================================================================================
' UserControl initialization/termination
'========================================================================================
Private Sub UserControl_Initialize()
    Call pvCreatePatternBrush
    DpiFactor = GetWindowsDPI
    m_StyleCurveRadius = -1
    m_Thumbsize_min = 5
    m_Thumbsize_max = 95
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Dim DontProcess As Boolean
    
    RaiseEvent KeyDown(KeyCode, Shift, DontProcess)
    
    If (DontProcess) Then Exit Sub
    
    Select Case (KeyCode)
        Case 33: Scroll_UP 1 'PGUP
        Case 34: Scroll_DOWN 1 'PGDOWN
        Case 36: Scroll_UP 2 'HOME
        Case 35: Scroll_DOWN 2 'END
        Case 37, 38: Scroll_UP 0 'LEFT,UP ARROW
        Case 39, 40: Scroll_DOWN 0 'RIGHT,DOWN ARROW
    End Select
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Sub UserControl_Terminate()
On Error Resume Next
    '-- In any case...
    Call pvKillTimer(TIMERID_HOT)
    Call pvKillTimer(TIMERID_CHANGE1)
    Call pvKillTimer(TIMERID_CHANGE2)
    
    '-- Stop subclassing
    Call SubclassEnd
    '-- Clean up
    Call DeleteObject(m_hPatternBrush)
    
    If (GdipToken <> 0) Then GdiplusShutdown (GdipToken)
    
End Sub

'========================================================================================
' Only in design-mode
'========================================================================================
Private Sub UserControl_Resize()
    On Error Resume Next
    If (Ambient.UserMode = False) Or (Not bSubclassed) Then
        Call pvOnSize
    End If
End Sub

Private Sub UserControl_Paint()
    If (Ambient.UserMode = False) Or (Not bSubclassed) Then
        Call pvOnPaint(UserControl.hdc)
    End If
End Sub

Public Property Get ThumbTooltipEnable() As Boolean
Attribute ThumbTooltipEnable.VB_Description = "Enable event showThumbTooltip in wich text can be set."
    ThumbTooltipEnable = m_bThumbTooltipEnable
End Property
Public Property Let ThumbTooltipEnable(ByVal bValue As Boolean)
    m_bThumbTooltipEnable = bValue
    PropertyChanged "ThumbTooltipEnable"
End Property

Public Property Get Parent() As Object
    Set Parent = UserControl.Extender.Parent
End Property
Public Property Set Parent(ByVal ObjValue As Object)
    Set UserControl.Extender.Parent = ObjValue
End Property

Public Property Get SmoothScrollFactor() As Single
    SmoothScrollFactor = m_SinSmoothScrollFactor
End Property
Public Property Let SmoothScrollFactor(ByVal SinValue As Single)
    If (m_SinSmoothScrollFactor > 1) Then m_SinSmoothScrollFactor = m_SinSmoothScrollFactor / 100
    m_SinSmoothScrollFactor = SinValue
    PropertyChanged "SmoothScrollFactor"
End Property

Public Property Get Extender() As Object
    Set Extender = UserControl.Extender
End Property

Public Property Get Visible() As Boolean
    Visible = UserControl.Extender.Visible
End Property
Public Property Let Visible(ByVal bValue As Boolean)
    UserControl.Extender.Visible = bValue
End Property

Public Property Get Top() As Long
    Top = UserControl.Extender.Top
End Property
Public Property Let Top(ByVal LonValue As Long)
    UserControl.Extender.Top = LonValue
End Property

Public Property Get Left() As Long
    Left = UserControl.Extender.Left
End Property
Public Property Let Left(ByVal LonValue As Long)
    UserControl.Extender.Left = LonValue
End Property

Public Property Get DisableMouseWheelSupport() As Boolean
    DisableMouseWheelSupport = m_bDisableMouseWheelSupport
End Property
Public Property Let DisableMouseWheelSupport(ByVal bValue As Boolean)
    m_bDisableMouseWheelSupport = bValue
    PropertyChanged "DisableMouseWheelSupport"
End Property

'========================================================================================
' Methods
'========================================================================================
Public Sub Refresh()
    '-- Force a complete paint
    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
End Sub

Public Function ZOrder(Optional LonValue As Long = 0)
    UserControl.Extender.ZOrder LonValue
End Function

Public Sub AttachHorizontalScrollBar(ucScrollbarH As ucScrollbar)
    
    'Sanity check controls
    If (m_eOrientation = oHorizontal) Then
        Debug.Print UserControl.Extender.Name & " AttachHorizontalScrollBar() can only be called on Vertical Scrollbars"
        Exit Sub
    End If
    
    'If (ucScrollbarH Is Nothing) Then
    '    Set m_ucScrollbarH = ucScrollbarH
    '    Exit Sub
    'End If
    
    If (ucScrollbarH.Orientation = oVertical) Then
        Debug.Print UserControl.Extender.Name & " AttachHorizontalScrollBar() parameter was another Vertical Scrollbar, you need to pass a Horizontal SB."
        Exit Sub
    End If
        
    Set m_ucScrollbarH = ucScrollbarH
    
End Sub

'======================================
'Do a scrolling top/left
'(depending on the type of scrollbar)
'======================================
Public Sub WheelScrollTopLeft()
    m_eHitTest = HT_NOTHING
    Call pvScrollPosDec(m_LonWheelChange, True)
    Call pvKillTimer(TIMERID_CHANGE1)
    Call pvSetTimer(TIMERID_CHANGE1, 150)
End Sub

'======================================
'Do a scrolling bottom/right
'(depending on the type of scrollbar)
'======================================
Public Sub WheelScrollBotRight()

    m_eHitTest = HT_NOTHING
    Call pvScrollPosInc(m_LonWheelChange, True)
    Call pvKillTimer(TIMERID_CHANGE1)
    Call pvSetTimer(TIMERID_CHANGE1, 150)
End Sub

'========================================================================================
' Messages response
'========================================================================================
Private Sub pvOnSize()
    Call pvSizeButtons
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
End Sub

Private Sub pvOnPaint(ByVal lhDC As Long)
On Error GoTo Err:
    Dim lfHorz As Long
    '
    lfHorz = -CLng(m_eOrientation = [oHorizontal])
    Select Case True
        Case m_eStyle = [sClassic] Or (m_eStyle = [sThemed] And m_bIsLuna = False)
            If (UserControl.Enabled) Then
                '-- Buttons
                If (m_bTLButtonPressed) Then
                    Call DrawFrameControl(lhDC, m_uRctTLButton, DFC_SCROLL, DFCS_SCROLLUP + (2 * lfHorz) Or DFCS_FLAT Or DFCS_PUSHED)
                Else
                    Call DrawFrameControl(lhDC, m_uRctTLButton, DFC_SCROLL, DFCS_SCROLLUP + (2 * lfHorz))
                End If
                If (m_bBRButtonPressed) Then
                    Call DrawFrameControl(lhDC, m_uRctBRButton, DFC_SCROLL, DFCS_SCROLLDOWN + (2 * lfHorz) Or DFCS_FLAT Or DFCS_PUSHED)
                Else
                    Call DrawFrameControl(lhDC, m_uRctBRButton, DFC_SCROLL, DFCS_SCROLLDOWN + (2 * lfHorz))
                End If
                '-- Track + thumb
                If (m_bHasTrack) Then
                    '-- Top-Left track part
                    If (m_bTLTrackPressed) Then
                        Call FillRect(lhDC, m_uRctTLTrack, GetStockObject(BLACK_BRUSH))
                    Else
                        Call FillRect(lhDC, m_uRctTLTrack, m_hPatternBrush)
                    End If
                    '-- Right-Bottom track part
                    If (m_bBRTrackPressed) Then
                        Call FillRect(lhDC, m_uRctBRTrack, GetStockObject(BLACK_BRUSH))
                    Else
                        Call FillRect(lhDC, m_uRctBRTrack, m_hPatternBrush)
                    End If
                    '-- Thumb
                    Call FillRect(lhDC, m_uRctThumb, GetSysColorBrush(COLOR_BTNFACE))
                    Call DrawEdge(lhDC, m_uRctThumb, BDR_RAISED, BF_RECT)
                End If
                If (m_bHasNullTrack) Then
                    Call FillRect(lhDC, m_uRctNullTrack, m_hPatternBrush)
                End If
            Else
                '-- Draw all disabled
                Call DrawFrameControl(lhDC, m_uRctTLButton, DFC_SCROLL, DFCS_SCROLLUP + (2 * lfHorz) Or DFCS_INACTIVE)
                Call DrawFrameControl(lhDC, m_uRctBRButton, DFC_SCROLL, DFCS_SCROLLDOWN + (2 * lfHorz) Or DFCS_INACTIVE)
                If (m_bHasTrack) Then
                    Call FillRect(lhDC, m_uRctTLTrack, m_hPatternBrush)
                    Call FillRect(lhDC, m_uRctBRTrack, m_hPatternBrush)
                    Call DrawFrameControl(lhDC, m_uRctThumb, 0, 0)
                End If
                If (m_bHasNullTrack) Then
                    Call FillRect(lhDC, m_uRctNullTrack, m_hPatternBrush)
                End If
            End If
        
        'Style created by Leandro Ascierto
        Case m_eStyle = [sGoogle]
            Dim Rect As Rect, estate As sbOnPaintPartStateCts
            Dim Radius As Long, lColor As Long, lBackColor&, tmpThumbColor&
            Dim mDC As Long, hBmp As Long, OldBmp As Long
            
            If (m_StyleBackColor = -1) Then
                lBackColor& = Ambient.BackColor 'UserControl.Parent.BackColor
            Else
                lBackColor& = m_StyleBackColor
            End If
            
            With m_uRctThumb
            
                If (m_StyleCurveRadius <> -1) Then
                    Radius = m_StyleCurveRadius
                Else
                    If (m_eOrientation = oVertical) Then
                        Radius = .X2 - .X1 - 1
                    Else
                        Radius = .Y2 - .Y1 - 1
                    End If
                End If
                estate = IIf(m_bThumbHot, [ppsHot], IIf(m_bThumbPressed, [ppsPressed], [ppsNormal]))
                
                If (m_StyleThumbColor = -1) Then
                    tmpThumbColor& = lBackColor&
                Else
                    tmpThumbColor& = m_StyleThumbColor
                End If
                
                If IsDarkColor(tmpThumbColor&) Then
                    If estate = ppsHot Then
                        lColor = ShiftColor(tmpThumbColor&, vbWhite, 180)
                    ElseIf estate = ppsPressed Then
                        lColor = ShiftColor(tmpThumbColor&, vbWhite, 150)
                    Else
                        lColor = ShiftColor(tmpThumbColor&, vbWhite, 200)
                    End If
                Else
                    If estate = ppsHot Then
                        lColor = ShiftColor(tmpThumbColor&, vbBlack, 220)
                    ElseIf estate = ppsPressed Then
                        lColor = ShiftColor(tmpThumbColor&, vbBlack, 180)
                    Else
                        lColor = ShiftColor(tmpThumbColor&, vbBlack, 200)
                    End If
                    
                End If
            
                mDC = CreateCompatibleDC(0)
                hBmp = CreateCompatibleBitmap(lhDC, .X2 - .X1, .Y2 - .Y1)
                OldBmp = SelectObject(mDC, hBmp)
                
                'Fill the rest of the toolbar to clear movement traces
                FillRectangle lhDC, m_uRctTLTrack, lBackColor&
                FillRectangle lhDC, m_uRctBRTrack, lBackColor&
                'Idem with buttons
                If (m_bShowButtons) Then
                    FillRectangle lhDC, m_uRctTLButton, lBackColor&
                    FillRectangle lhDC, m_uRctBRButton, lBackColor&
                End If
                
                'Background fill
                RoundRectPlus mDC, -DpiFactor, -DpiFactor, .X2 - .X1 + DpiFactor, .Y2 - .Y1 + DpiFactor, RGBtoARGB(lBackColor&, 100), 0, 0
                'Curve
                RoundRectPlus mDC, 0, 0, .X2 - .X1 - DpiFactor, .Y2 - .Y1 - DpiFactor, RGBtoARGB(lColor, 100), 0, Radius
                BitBlt lhDC, .X1, .Y1, .X2 - .X1, .Y2 - .Y1, mDC, 0, 0, vbSrcCopy
                DeleteObject SelectObject(mDC, OldBmp)
                DeleteDC mDC
            End With
            
        
        Case m_eStyle = [sFlat]
            If (UserControl.Enabled) Then
                '-- Buttons
                If (m_bTLButtonHot) Then
                    Call pvDrawFlatButton(lhDC, m_uRctTLButton, DFCS_SCROLLUP + (2 * lfHorz), [fbsHot])
                Else
                    If (m_bTLButtonPressed) Then
                        Call pvDrawFlatButton(lhDC, m_uRctTLButton, DFCS_SCROLLUP + (2 * lfHorz), [fbsSelected])
                      Else
                        Call pvDrawFlatButton(lhDC, m_uRctTLButton, DFCS_SCROLLUP + (2 * lfHorz), [fbsNormal])
                    End If
                End If
                If (m_bBRButtonHot) Then
                    Call pvDrawFlatButton(lhDC, m_uRctBRButton, DFCS_SCROLLDOWN + (2 * lfHorz), [fbsHot])
                Else
                    If (m_bBRButtonPressed) Then
                        Call pvDrawFlatButton(lhDC, m_uRctBRButton, DFCS_SCROLLDOWN + (2 * lfHorz), [fbsSelected])
                    Else
                        Call pvDrawFlatButton(lhDC, m_uRctBRButton, DFCS_SCROLLDOWN + (2 * lfHorz), [fbsNormal])
                    End If
                End If
                '-- Track + thumb
                If (m_bHasTrack) Then
                    '-- Top-Left track part
                    If (m_bTLTrackPressed) Then
                        Call FillRect(lhDC, m_uRctTLTrack, GetStockObject(BLACK_BRUSH))
                    Else
                        Call FillRect(lhDC, m_uRctTLTrack, m_hPatternBrush)
                    End If
                    '-- Right-Bottom track part
                    If (m_bBRTrackPressed) Then
                        Call FillRect(lhDC, m_uRctBRTrack, GetStockObject(BLACK_BRUSH))
                    Else
                        Call FillRect(lhDC, m_uRctBRTrack, m_hPatternBrush)
                    End If
                    '-- Thumb
                    If (m_bThumbHot) Then
                        Call FillRect(lhDC, m_uRctThumb, GetSysColorBrush(COLOR_3DSHADOW))
                    Else
                        If (m_bThumbPressed) Then
                            Call FillRect(lhDC, m_uRctThumb, GetSysColorBrush(COLOR_3DSHADOW))
                        Else
                            Call DrawFrameControl(lhDC, m_uRctThumb, 0, DFCS_FLAT)
                        End If
                    End If
                End If
                If (m_bHasNullTrack) Then
                    Call FillRect(lhDC, m_uRctNullTrack, m_hPatternBrush)
                End If
            Else
                '-- Draw all disabled
                Call DrawFrameControl(lhDC, m_uRctTLButton, DFC_SCROLL, DFCS_SCROLLUP + (2 * lfHorz) Or DFCS_FLAT Or DFCS_INACTIVE)
                Call DrawFrameControl(lhDC, m_uRctBRButton, DFC_SCROLL, DFCS_SCROLLDOWN + (2 * lfHorz) Or DFCS_FLAT Or DFCS_INACTIVE)
                If (m_bHasTrack) Then
                    Call FillRect(lhDC, m_uRctTLTrack, m_hPatternBrush)
                    Call FillRect(lhDC, m_uRctBRTrack, m_hPatternBrush)
                    Call DrawFrameControl(lhDC, m_uRctThumb, 0, DFCS_FLAT)
                End If
                If (m_bHasNullTrack) Then
                    Call FillRect(lhDC, m_uRctNullTrack, m_hPatternBrush)
                End If
            End If
        Case m_eStyle = [sThemed]
            If (UserControl.Enabled) Then
                '-- Buttons
                If (m_bTLButtonHot) Or (m_bBRButtonHot) Then
                    Call pvDrawThemePart(lhDC, SB_THEME, SBP_ARROWBTN, ABS_UPHOT + (8 * lfHorz), m_uRctTLButton)
                Else
                    If (m_bTLButtonPressed) Then
                        Call pvDrawThemePart(lhDC, SB_THEME, SBP_ARROWBTN, ABS_UPPRESSED + (8 * lfHorz), m_uRctTLButton)
                    Else
                        Call pvDrawThemePart(lhDC, SB_THEME, SBP_ARROWBTN, ABS_UPNORMAL + (8 * lfHorz), m_uRctTLButton)
                    End If
                End If
                If (m_bTLButtonHot) Or (m_bBRButtonHot) Then
                    Call pvDrawThemePart(lhDC, SB_THEME, SBP_ARROWBTN, ABS_DOWNHOT + (8 * lfHorz), m_uRctBRButton)
                Else
                    If (m_bBRButtonPressed) Then
                        Call pvDrawThemePart(lhDC, SB_THEME, SBP_ARROWBTN, ABS_DOWNPRESSED + (8 * lfHorz), m_uRctBRButton)
                    Else
                        Call pvDrawThemePart(lhDC, SB_THEME, SBP_ARROWBTN, ABS_DOWNNORMAL + (8 * lfHorz), m_uRctBRButton)
                    End If
                End If
                '-- Track + thumb
                If (m_bHasTrack) Then
                    '-- Top-Left track part
                    If (m_bTLTrackPressed) Then
                        Call pvDrawThemePart(lhDC, SB_THEME, SBP_UPPERTRACKVERT - (2 * lfHorz), HSS_PUSHED, m_uRctTLTrack)
                    Else
                        Call pvDrawThemePart(lhDC, SB_THEME, SBP_UPPERTRACKVERT - (2 * lfHorz), HSS_NORMAL, m_uRctTLTrack)
                    End If
                    '-- Right-Bottom track part
                    If (m_bBRTrackPressed) Then
                        Call pvDrawThemePart(lhDC, SB_THEME, SBP_LOWERTRACKVERT - (2 * lfHorz), HSS_PUSHED, m_uRctBRTrack)
                    Else
                        Call pvDrawThemePart(lhDC, SB_THEME, SBP_LOWERTRACKVERT - (2 * lfHorz), HSS_NORMAL, m_uRctBRTrack)
                    End If
                    '-- Thumb
                    If (m_bThumbHot) Then
                        Call pvDrawThemePart(lhDC, SB_THEME, SBP_THUMBBTNVERT - (lfHorz), VSS_HOT, m_uRctThumb)
                        If (m_lThumbSize >= GRIPPERSIZE_MIN) Then
                            Call pvDrawThemePart(lhDC, SB_THEME, SBP_GRIPPERVERT - (lfHorz), VSS_HOT, m_uRctThumb)
                        End If
                    Else
                        If (m_bThumbPressed) Then
                            Call pvDrawThemePart(lhDC, SB_THEME, SBP_THUMBBTNVERT - (lfHorz), VSS_PUSHED, m_uRctThumb)
                            If (m_lThumbSize >= GRIPPERSIZE_MIN) Then
                                Call pvDrawThemePart(lhDC, SB_THEME, SBP_GRIPPERVERT - (lfHorz), VSS_PUSHED, m_uRctThumb)
                            End If
                        Else
                            Call pvDrawThemePart(lhDC, SB_THEME, SBP_THUMBBTNVERT - (lfHorz), VSS_NORMAL, m_uRctThumb)
                            If (m_lThumbSize >= GRIPPERSIZE_MIN) Then
                               Call pvDrawThemePart(lhDC, SB_THEME, SBP_GRIPPERVERT - (lfHorz), VSS_NORMAL, m_uRctThumb)
                            End If
                        End If
                    End If
                End If
                If (m_bHasNullTrack) Then
                    Call pvDrawThemePart(lhDC, SB_THEME, SBP_UPPERTRACKVERT - (2 * lfHorz), HSS_NORMAL, m_uRctNullTrack)
                End If
            Else
                '-- Draw all disabled
                Call pvDrawThemePart(lhDC, SB_THEME, SBP_ARROWBTN, ABS_UPDISABLED + (8 * lfHorz), m_uRctTLButton)
                Call pvDrawThemePart(lhDC, SB_THEME, SBP_ARROWBTN, ABS_DOWNDISABLED + (8 * lfHorz), m_uRctBRButton)
                Call pvDrawThemePart(lhDC, SB_THEME, SBP_UPPERTRACKVERT + (2 * lfHorz), HSS_DISABLED, m_uRctTLTrack)
                If (m_bHasTrack) Then
                    Call pvDrawThemePart(lhDC, SB_THEME, SBP_LOWERTRACKVERT + (2 * lfHorz), HSS_DISABLED, m_uRctBRTrack)
                    Call pvDrawThemePart(lhDC, SB_THEME, SBP_THUMBBTNVERT - (lfHorz), VSS_DISABLED, m_uRctThumb)
                    If (m_lThumbSize >= GRIPPERSIZE_MIN) Then
                        Call pvDrawThemePart(lhDC, SB_THEME, SBP_GRIPPERVERT - (lfHorz), VSS_DISABLED, m_uRctThumb)
                    End If
                End If
                If (m_bHasNullTrack) Then
                    Call pvDrawThemePart(lhDC, SB_THEME, SBP_UPPERTRACKVERT - (2 * lfHorz), HSS_DISABLED, m_uRctNullTrack)
                End If
            End If
            
        Case m_eStyle = [sCustomDraw]
            With m_uRctTLButton
                RaiseEvent OnPaint(lhDC, .X1, .Y1, .X2, .Y2, [ppTLButton], IIf(m_bTLButtonHot, [ppsHot], IIf(m_bTLButtonPressed, [ppsPressed], [ppsNormal])))
            End With
            With m_uRctBRButton
                RaiseEvent OnPaint(lhDC, .X1, .Y1, .X2, .Y2, [ppBRButton], IIf(m_bBRButtonHot, [ppsHot], IIf(m_bBRButtonPressed, [ppsPressed], [ppsNormal])))
            End With
            If (m_bHasTrack) Then
                With m_uRctTLTrack
                    RaiseEvent OnPaint(lhDC, .X1, .Y1, .X2, .Y2, [ppTLTrack], IIf(m_bTLTrackPressed, [ppsPressed], [ppsNormal]))
                End With
                With m_uRctBRTrack
                    RaiseEvent OnPaint(lhDC, .X1, .Y1, .X2, .Y2, [ppBRTrack], IIf(m_bBRTrackPressed, [ppsPressed], [ppsNormal]))
                End With
                With m_uRctThumb
                    RaiseEvent OnPaint(lhDC, .X1, .Y1, .X2, .Y2, [ppThumb], IIf(m_bThumbHot, [ppsHot], IIf(m_bThumbPressed, [ppsPressed], [ppsNormal])))
                End With
            End If
            If (m_bHasNullTrack) Then
                With m_uRctNullTrack
                    RaiseEvent OnPaint(lhDC, .X1, .Y1, .X2, .Y2, [ppNullTrack], [ppsNormal])
                End With
            End If
    End Select

    Exit Sub
Err:
    Debug.Print "pvOnPaint Err: " & Err.Description
End Sub

Private Sub SetThumbTooltip(Enable As Boolean)
Dim c&, d&, p As POINTAPI, hx&, hy&, sTooltipText$
Dim TxtObj As Object
    If (Enable = False) Then
        txtToolTip.Visible = False
        timHideTooltip.Enabled = False
        
        'Not needed
        'SetParent txtToolTip.hWnd, UserControl.hWnd
        Exit Sub
    End If
    
    'Raise event so instance can set the tooltip as needed
    RaiseEvent showThumbTooltip(sTooltipText, TxtObj)
    
    'Exit if no object passed
    If (TxtObj Is Nothing) Then Exit Sub
    
    'Prepare textbox for tooltip
    If (txtToolTip Is Nothing) Then
        Set txtToolTip = TxtObj
        txtToolTip.Appearance = 0
    
        c = GetParent(txtToolTip.hwnd)
        d = GetDesktopWindow()
        If (c <> d) Then
            SetParent txtToolTip.hwnd, d
            pOnTop txtToolTip.hwnd, True, True
            'To hide it from taskbar
            SetWindowLong txtToolTip.hwnd, GWL_EXSTYLE, GetWindowLong(txtToolTip.hwnd, GWL_EXSTYLE) Or WS_EX_TOOLWINDOW
            'To add shadow (sadly only work inside VB6 IDE.On exe add too much space on top and flicker)
            'SetWindowLong txtToolTip.hWnd, GWL_STYLE, GetWindowLong(txtToolTip.hWnd, GWL_STYLE) Or WS_THICKFRAME
        End If
    
    End If
    
    'Always se the font so 1 txt can be recycled for all Scrollbars
    Set txtToolTip.Font = m_StdThumbTooltipFont
    
    'if blank text is passed then tooltip is disabled
    If (sTooltipText = "") Then
        If (txtToolTip.Visible = True) Then txtToolTip.Visible = False
        Exit Sub
    End If
    
    txtToolTip.Text = sTooltipText
    txtToolTip.Width = (UserControl.TextWidth(sTooltipText) + 5) * Screen.TwipsPerPixelX '+12
    'Also process height to support multiline
    txtToolTip.Height = (UserControl.TextHeight(sTooltipText)) * Screen.TwipsPerPixelY  '+9
    
    'Uncomment to use with Shadow
    'txtToolTip.Width = txtToolTip.Width + (20 * Screen.TwipsPerPixelX)
    'txtToolTip.Height = txtToolTip.Height + (10 * Screen.TwipsPerPixelY)
    
    
    hx = (txtToolTip.Width / 2) / Screen.TwipsPerPixelX
    hy = txtToolTip.Height / Screen.TwipsPerPixelY
    'Position it 15 pixels on top of the mouse
    GetCursorPos p
    SetWindowPos txtToolTip.hwnd, HWND_TOPMOST, (p.X - hx), (p.Y - hy) - 15, 0, 0, SWP_NOSIZE Or SWP_NOACTIVATE
    timHideTooltip.Enabled = False
    txtToolTip.Visible = True
    timHideTooltip.Enabled = True

End Sub

'Ensure Thumbtooltip is hidden when focus is changed (alt-tab) and thumbtooltip was visible
Private Sub timHideTooltip_Timer()
    If Not (txtToolTip Is Nothing) Then
        If (txtToolTip.Visible = True) Then
            If (m_bThumbPressed = False) Then
                txtToolTip.Visible = False
            Else
                'Update the ThumbtooltipValue
                SetThumbTooltip (True)
            End If
        Else
            timHideTooltip.Enabled = False
        End If
    Else
        timHideTooltip.Enabled = False
    End If
End Sub

Private Sub pvOnMouseDown(ByVal wParam As Long, ByVal lParam As Long)
    If (wParam And MK_LBUTTON = MK_LBUTTON) Then
        Call pvMakePoints(lParam, m_x, m_y)
        m_eHitTest = pvHitTest(m_x, m_y)
        Select Case m_eHitTest
            Case HT_THUMB
            
                If (m_bThumbTooltipEnable) Then SetThumbTooltip True
            
                Select Case m_eOrientation
                    Case [oVertical]
                        m_lThumbOffset = m_uRctThumb.Y1 - m_y
                    Case [oHorizontal]
                        m_lThumbOffset = m_uRctThumb.X1 - m_x
                End Select
                m_bThumbPressed = True
                m_bThumbHot = False
                m_lValueStartDrag = m_lValue
                Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
            Case HT_TLBUTTON
                m_bTLButtonPressed = True
                m_bTLButtonHot = False
                Call pvScrollPosDec(m_lSmallChange, True)
                Call pvKillTimer(TIMERID_CHANGE2)
                Call pvSetTimer(TIMERID_CHANGE2, m_lChangeDelay)
            Case HT_BRBUTTON
                m_bBRButtonPressed = True
                m_bBRButtonHot = False
                Call pvScrollPosInc(m_lSmallChange, True)
                Call pvKillTimer(TIMERID_CHANGE2)
                Call pvSetTimer(TIMERID_CHANGE2, m_lChangeDelay)
            Case HT_TLTRACK
                m_bTLTrackPressed = True
                Call pvScrollPosDec(m_lLargeChange)
                Call pvKillTimer(TIMERID_CHANGE1)
                Call pvSetTimer(TIMERID_CHANGE1, m_lChangeDelay)
            Case HT_BRTRACK
                m_bBRTrackPressed = True
                Call pvScrollPosInc(m_lLargeChange)
                Call pvKillTimer(TIMERID_CHANGE1)
                Call pvSetTimer(TIMERID_CHANGE1, m_lChangeDelay)
        End Select
    End If
End Sub

Private Sub pvOnMouseMove(ByVal wParam As Long, ByVal lParam As Long)
On Error GoTo Err:
Dim auxl&
    Dim lValuePrev As Long, lThumbPosPrev As Long, bPressed As Boolean, bHot As Boolean
    '
    Call pvMakePoints(lParam, m_x, m_y)
    If (wParam And MK_LBUTTON = MK_LBUTTON) Then
        Select Case m_eHitTest
            Case HT_THUMB
            
                If (m_bThumbTooltipEnable) Then SetThumbTooltip True
                            
                lValuePrev = m_lValue
                lThumbPosPrev = m_lThumbPos
                If (PtInRect(m_uRctDrag, m_x, m_y)) Then
                    Select Case m_eOrientation
                        Case [oVertical]
                            m_lThumbPos = m_y + m_lThumbOffset
                            If (m_lThumbPos < m_uRctTLButton.Y2) Then
                                m_lThumbPos = m_uRctTLButton.Y2
                            End If
                            If (m_lThumbPos + m_lThumbSize > m_uRctBRButton.Y1) Then
                                m_lThumbPos = m_uRctBRButton.Y1 - m_lThumbSize
                            End If
                        Case [oHorizontal]
                            m_lThumbPos = m_x + m_lThumbOffset
                            If (m_lThumbPos < m_uRctTLButton.X2) Then
                                m_lThumbPos = m_uRctTLButton.X2
                            End If
                            If (m_lThumbPos + m_lThumbSize > m_uRctBRButton.X1) Then
                                m_lThumbPos = m_uRctBRButton.X1 - m_lThumbSize
                            End If
                    End Select
                    '*2-
                    'Shagratt 08/12/19
                    'For instant replace all and set m_lValue=pvGetScrollPos()
                    auxl& = pvGetScrollPos()
                    If (auxl& <> lValuePrev) Then
                        If (TimSmoothChange.Enabled) Then TimSmoothChange.Enabled = False
                        m_TargetValue = auxl&
                        TimSmoothChange.Enabled = True
                        TimSmoothChange_Timer 'BY LeandroA
                    End If
                Else
                    'Shagratt removed from original Carles P.V. code
                    'This is when you drag the thumb outside the Scrollbar area... not needed
                    'm_lValue = m_lValueStartDrag
                    'm_lThumbPos = pvGetThumbPos()
                End If
                If (m_lThumbPos <> lThumbPosPrev) Then
                    Call pvSizeTrack
                    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                    If (m_lValue <> lValuePrev) Then
                        RaiseEvent Scroll
                    End If
                End If
            Case HT_TLBUTTON
                If (m_bTLButtonPressed) Then
                    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                End If
            Case HT_BRBUTTON
                If (m_bBRButtonPressed) Then
                    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                End If
        End Select
    Else
        m_eHitTestHot = pvHitTest(m_x, m_y)
        Select Case m_eHitTestHot
            Case HT_TLBUTTON
                bHot = (PtInRect(m_uRctTLButton, m_x, m_y) <> 0)
                If (m_bTLButtonHot Xor bHot) Then
                    m_bTLButtonHot = True
                    m_bBRButtonHot = False
                    m_bThumbHot = False
                    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                    Call pvKillTimer(TIMERID_HOT)
                    Call pvSetTimer(TIMERID_HOT, TIMERDT_HOT)
                End If
            Case HT_BRBUTTON
                bHot = (PtInRect(m_uRctBRButton, m_x, m_y) <> 0)
                If (m_bBRButtonHot Xor bHot) Then
                    m_bTLButtonHot = False
                    m_bBRButtonHot = True
                    m_bThumbHot = False
                    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                    Call pvKillTimer(TIMERID_HOT)
                    Call pvSetTimer(TIMERID_HOT, TIMERDT_HOT)
                End If
            Case HT_THUMB
                bHot = (PtInRect(m_uRctThumb, m_x, m_y) <> 0)
                If (m_bThumbHot Xor bHot) Then
                    m_bTLButtonHot = False
                    m_bBRButtonHot = False
                    m_bThumbHot = True
                    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                    Call pvKillTimer(TIMERID_HOT)
                    Call pvSetTimer(TIMERID_HOT, TIMERDT_HOT)
                End If
        End Select
    End If
Err:
End Sub

Private Sub pvOnMouseUp()
    Call pvKillTimer(TIMERID_HOT)
    Call pvKillTimer(TIMERID_CHANGE1)
    Call pvKillTimer(TIMERID_CHANGE2)
    
    If Not (txtToolTip Is Nothing) Then
        If (txtToolTip.Visible) Then SetThumbTooltip False
    End If
'    If (m_eHitTest = HT_THUMB) Then
'        If (m_lValue <> m_lValueStartDrag) Then
'            RaiseEvent Change
'        End If
'    End If
    'Shagratt this is to keep the thumb on the position where we release the mouse
    If (m_eHitTest <> 5) Then
        m_eHitTest = HT_NOTHING
    End If
    m_bTLButtonPressed = False
    m_bBRButtonPressed = False
    m_bThumbPressed = False
    m_bTLTrackPressed = False
    m_bBRTrackPressed = False
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
End Sub

Private Sub pvOnTimer(ByVal wParam As Long)
    Dim uPt As POINTAPI
    Select Case wParam
        Case TIMERID_CHANGE1
            Call pvKillTimer(TIMERID_CHANGE1)
            Call pvSetTimer(TIMERID_CHANGE2, m_lChangeFrequency)
        Case TIMERID_CHANGE2
            Select Case m_eHitTest
                Case HT_TLBUTTON
                    'Shagratt 30/11/19
                    pvScrollPosDec (m_lSmallChange)
                    Call pvKillTimer(TIMERID_CHANGE2)
                    Call pvSetTimer(TIMERID_CHANGE2, 50)
                Case HT_BRBUTTON
                    'Shagratt 30/11/19
                    pvScrollPosInc (m_lSmallChange)
                    Call pvKillTimer(TIMERID_CHANGE2)
                    Call pvSetTimer(TIMERID_CHANGE2, 50)
                Case HT_TLTRACK
                    Select Case m_eOrientation
                        Case [oVertical]
                            If (m_lThumbPos > m_y) Then
                                m_bTLTrackPressed = True
                                Call pvScrollPosDec(m_lLargeChange)
                            Else
                                m_bTLTrackPressed = False
                                Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                            End If
                        Case [oHorizontal]
                            If (m_lThumbPos > m_x) Then
                                m_bTLTrackPressed = True
                                Call pvScrollPosDec(m_lLargeChange)
                            Else
                                m_bTLTrackPressed = False
                                Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                            End If
                    End Select
                Case HT_BRTRACK
                    Select Case m_eOrientation
                        Case [oVertical]
                            If (m_lThumbPos + m_lThumbSize < m_y) Then
                                m_bBRTrackPressed = True
                                Call pvScrollPosInc(m_lLargeChange)
                            Else
                                m_bBRTrackPressed = False
                                Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                            End If
                        Case [oHorizontal]
                            If (m_lThumbPos + m_lThumbSize < m_x) Then
                                m_bBRTrackPressed = True
                                Call pvScrollPosInc(m_lLargeChange)
                            Else
                                m_bBRTrackPressed = False
                                Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                            End If
                    End Select
           End Select
        Case TIMERID_HOT
            Call GetCursorPos(uPt)
            Call ScreenToClient(hwnd, uPt)
            Select Case True
                Case m_bTLButtonHot
                    If (PtInRect(m_uRctTLButton, uPt.X, uPt.Y) = 0) Then
                        m_bTLButtonHot = False
                        Call pvKillTimer(TIMERID_HOT)
                        Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                    End If
                Case m_bBRButtonHot
                    If (PtInRect(m_uRctBRButton, uPt.X, uPt.Y) = 0) Then
                        m_bBRButtonHot = False
                        Call pvKillTimer(TIMERID_HOT)
                        Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                    End If
                Case m_bThumbHot
                    If (PtInRect(m_uRctThumb, uPt.X, uPt.Y) = 0) Then
                        m_bThumbHot = False
                        Call pvKillTimer(TIMERID_HOT)
                        Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                    End If
            End Select
                        
    End Select
End Sub

Private Sub pvOnSysColorChange()
    '-- Repaint all
    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
End Sub

Private Sub pvOnThemeChanged()
    '-- Check OS
    Call pvCheckEnvironment
    RaiseEvent ThemeChanged
    '-- Repaint all
    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
End Sub

'========================================================================================
' Private
'========================================================================================
'----------------------------------------------------------------------------------------
' Sizing
'----------------------------------------------------------------------------------------
Private Sub pvSizeButtons()
    Dim uRct As Rect, lButtonSize As Long
    '
    Call GetClientRect(hwnd, uRct)
    m_bHasTrack = False
    m_bHasNullTrack = False
    Select Case m_eOrientation
        Case [oVertical]
            '-- Size buttons
            lButtonSize = GetSystemMetrics(SM_CYVSCROLL) * -CLng(m_bShowButtons)
            With uRct
                If (2 * lButtonSize + Thumbsize_AbsMin > .Y2) Then
                    If (2 * lButtonSize < .Y2) Then
                        Call SetRect(m_uRctTLButton, 0, 0, .X2, lButtonSize)
                        Call SetRect(m_uRctBRButton, 0, .Y2 - lButtonSize, .X2, .Y2)
                        m_bHasNullTrack = True
                        Call SetRect(m_uRctNullTrack, 0, lButtonSize, .X2, .Y2 - lButtonSize)
                    Else
                        Call SetRect(m_uRctTLButton, 0, 0, .X2, .Y2 \ 2)
                        Call SetRect(m_uRctBRButton, 0, .Y2 \ 2 + (.Y2 Mod 2), .X2, .Y2)
                        m_bHasNullTrack = CBool(.Y2 Mod 2)
                        If (m_bHasNullTrack) Then
                            Call SetRect(m_uRctNullTrack, 0, .Y2 \ 2, .X2, .Y2 \ 2 + 1)
                        End If
                    End If
                Else
                    m_bHasTrack = True
                    Call SetRect(m_uRctTLButton, 0, 0, .X2, lButtonSize)
                    Call SetRect(m_uRctBRButton, 0, .Y2 - lButtonSize, .X2, .Y2)
                End If
            End With
            '-- Get max. drag area
            Call CopyRect(m_uRctDrag, uRct)
            Call InflateRect(m_uRctDrag, 250, 25)
        Case [oHorizontal]
            '-- Size buttons
            lButtonSize = GetSystemMetrics(SM_CXHSCROLL) * -CLng(m_bShowButtons)
            With uRct
                If (2 * lButtonSize + Thumbsize_AbsMin > .X2) Then
                    If (2 * lButtonSize < .X2) Then
                        Call SetRect(m_uRctTLButton, 0, 0, lButtonSize, .Y2)
                        Call SetRect(m_uRctBRButton, .X2 - lButtonSize, 0, .X2, .Y2)
                        m_bHasNullTrack = True
                        Call SetRect(m_uRctNullTrack, lButtonSize, 0, .X2 - lButtonSize, .Y2)
                    Else
                        Call SetRect(m_uRctTLButton, 0, 0, .X2 \ 2, .Y2)
                        Call SetRect(m_uRctBRButton, .X2 \ 2 + (.X2 Mod 2), 0, .X2, .Y2)
                        m_bHasNullTrack = CBool(.X2 Mod 2)
                        If (m_bHasNullTrack) Then
                            Call SetRect(m_uRctNullTrack, .X2 \ 2, 0, .X2 \ 2 + 1, .Y2)
                        End If
                    End If
                Else
                    m_bHasTrack = True
                    Call SetRect(m_uRctTLButton, 0, 0, lButtonSize, .Y2)
                    Call SetRect(m_uRctBRButton, .X2 - lButtonSize, 0, .X2, .Y2)
                End If
            End With
            '-- Get max. drag area
            Call CopyRect(m_uRctDrag, uRct)
            Call InflateRect(m_uRctDrag, 25, 250)
    End Select
    '-- No track: avoid pvSizeTrack() calcs.
    If (m_bHasTrack = False) Then
        Call SetRectEmpty(m_uRctTLTrack)
        Call SetRectEmpty(m_uRctBRTrack)
        Call SetRectEmpty(m_uRctThumb)
    End If
End Sub

Private Sub pvSizeTrack()
    If (m_bHasTrack) Then
        '-- Tracks and thumbs exist
        Select Case m_eOrientation
            Case [oVertical]
                '-- Size both track parts and thumb
                Call SetRect(m_uRctTLTrack, 0, m_uRctTLButton.Y2, m_uRctTLButton.X2, m_lThumbPos)
                Call SetRect(m_uRctBRTrack, 0, m_lThumbPos + m_lThumbSize, m_uRctBRButton.X2, m_uRctBRButton.Y1)
                Call SetRect(m_uRctThumb, 0, m_lThumbPos, m_uRctBRButton.X2, m_lThumbPos + m_lThumbSize)
            Case [oHorizontal]
                '-- Size both track parts and thumb
                Call SetRect(m_uRctTLTrack, m_uRctTLButton.X2, 0, m_lThumbPos, m_uRctTLButton.Y2)
                Call SetRect(m_uRctBRTrack, m_lThumbPos + m_lThumbSize, 0, m_uRctBRButton.X1, m_uRctBRButton.Y2)
                Call SetRect(m_uRctThumb, m_lThumbPos, 0, m_lThumbPos + m_lThumbSize, m_uRctBRButton.Y2)
        End Select
    End If
End Sub

Private Function pvGetThumbSize() As Long
On Error Resume Next
Dim aux&
    'Shagratt At last found the problem with my formula :)
    If (m_lLargeChange > m_lAbsRange) Then
        m_lLargeChange = m_lAbsRange
    End If
    Select Case m_eOrientation
        Case [oVertical]
            'Original: pvGetThumbSize = (m_uRctBRButton.y1 - m_uRctTLButton.y2) \ (m_lAbsRange \ m_lLargeChange + 1)
            
            '// By AMV 09/08/2021 ================================================================================
            '// original: pvGetThumbSize = (m_uRctBRButton.Y1 - m_uRctTLButton.Y2) * ((m_lLargeChange / m_lAbsRange))
            If m_lAbsRange > 0 Then
              pvGetThumbSize = (m_uRctBRButton.Y1 - m_uRctTLButton.Y2) * ((m_lLargeChange / m_lAbsRange))
            Else
              pvGetThumbSize = (m_uRctBRButton.Y1 - m_uRctTLButton.Y2) ' * ((m_lLargeChange / m_lAbsRange))
            End If
            '// ================|================================================================|================
           
            ' Sugerencia de JB
            aux& = (m_uRctBRButton.Y1 - m_uRctTLButton.Y2) * (m_Thumbsize_min / 100)
            If (pvGetThumbSize < aux) Then pvGetThumbSize = aux
            aux& = (m_uRctBRButton.Y1 - m_uRctTLButton.Y2) * (m_Thumbsize_max / 100)
            If (pvGetThumbSize > aux) Then pvGetThumbSize = aux
            ' Fin Sugerencia de JB
           
            'Absolute minimum no matter what
            If (pvGetThumbSize < Thumbsize_AbsMin) Then pvGetThumbSize = Thumbsize_AbsMin
            
        Case [oHorizontal]
            'Original: pvGetThumbSize = (m_uRctBRButton.x1 - m_uRctTLButton.x2) \ (m_lAbsRange \ m_lLargeChange + 1)
            pvGetThumbSize = CLng((m_uRctBRButton.X1 - m_uRctTLButton.X2) * ((m_lLargeChange / m_lAbsRange)))
            
            ' Sugerencia de JB
            aux = (m_uRctBRButton.X1 - m_uRctTLButton.X2) * (m_Thumbsize_min / 100)
            If (pvGetThumbSize < aux) Then pvGetThumbSize = aux
            aux = (m_uRctBRButton.X1 - m_uRctTLButton.X2) * (m_Thumbsize_max / 100)
            If (pvGetThumbSize > aux) Then pvGetThumbSize = aux
            ' Fin Sugerencia de JB
            
            'Absolute minimum no matter what
            If (pvGetThumbSize < Thumbsize_AbsMin) Then pvGetThumbSize = Thumbsize_AbsMin
    End Select
    On Error GoTo 0
End Function

'----------------------------------------------------------------------------------------
' Controling value
'----------------------------------------------------------------------------------------
'Animate to value
Public Function ScrollToValue(lTarget&)
    m_eHitTest = HT_NOTHING
    m_TargetValue = lTarget&
    TimSmoothChange.Enabled = True
End Function

Public Property Get TargetValue()
Attribute TargetValue.VB_Description = "Value that the scrollbar is going to reach after animation ended"
    TargetValue = m_TargetValue
End Property

Private Sub TimSmoothChange_Timer()
Dim SmoothStep&, vstep&
    If (m_lValue = m_TargetValue) Then
        TimSmoothChange.Enabled = False
        Exit Sub
    End If

    vstep& = m_TargetValue - m_lValue
    SmoothStep& = Abs(vstep) * m_SinSmoothScrollFactor '0.3
    If (SmoothStep& < 1) Then SmoothStep& = 1
    
    If (vstep& > 0) Then
        pvScrollPosInc2 SmoothStep, True
    Else
        pvScrollPosDec2 SmoothStep, True
    End If
End Sub

Private Function pvScrollPosDec(ByVal lSteps As Long, Optional ByVal bForceRepaint As Boolean = False) As Boolean
    'If it was animating the scroll with add to the final value, and not the actual value
    If (TimSmoothChange.Enabled) Then
        m_TargetValue = m_TargetValue - lSteps
    Else
        m_TargetValue = m_lValue - lSteps
    End If
    If (m_TargetValue < m_lMin) Then m_TargetValue = m_lMin
    TimSmoothChange.Enabled = True
End Function

Private Function pvScrollPosDec2(ByVal lSteps As Long, Optional ByVal bForceRepaint As Boolean = False) As Boolean
    Dim bChange As Boolean, lValuePrev As Long

    lValuePrev = m_lValue
    m_lValue = m_lValue - lSteps
    If (m_lValue < m_lMin) Then
        m_lValue = m_lMin
        m_TargetValue = m_lValue
    End If
    If (m_lValue <> lValuePrev) Then
        m_lThumbPos = pvGetThumbPos()
        Call pvSizeTrack
        bChange = True
    End If
    If (bChange Or bForceRepaint) Then
        Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
        If (bChange) Then
            RaiseEvent Change
        End If
    End If
    pvScrollPosDec2 = bChange
End Function


Private Function pvScrollPosInc(ByVal lSteps As Long, Optional ByVal bForceRepaint As Boolean = False) As Boolean
    'If it was animating the scroll with add to the final value, and not the actual value
    If (TimSmoothChange.Enabled) Then
        m_TargetValue = m_TargetValue + lSteps
    Else
        m_TargetValue = m_lValue + lSteps
    End If
    If (m_TargetValue > m_lMax) Then m_TargetValue = m_lMax
    TimSmoothChange.Enabled = True
End Function

Private Function pvScrollPosInc2(ByVal lSteps As Long, Optional ByVal bForceRepaint As Boolean = False) As Boolean
    Dim bChange As Boolean, lValuePrev As Long
    '
    lValuePrev = m_lValue
    m_lValue = m_lValue + lSteps
    If (m_lValue > m_lMax) Then
        m_lValue = m_lMax
        m_TargetValue = m_lValue
    End If
    If (m_lValue <> lValuePrev) Then
        m_lThumbPos = pvGetThumbPos()
        Call pvSizeTrack
        bChange = True
    End If
    If (bChange Or bForceRepaint) Then
        Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
        If (bChange) Then
            RaiseEvent Change
        End If
    End If
    pvScrollPosInc2 = bChange
End Function

'----------------------------------------------------------------------------------------
' Positioning thumb from Value (not thumb position!)
'----------------------------------------------------------------------------------------
Private Function pvGetThumbPos() As Long
    On Error Resume Next
    Select Case m_eOrientation
        Case [oVertical]
            'Shagratt if dragged from the thumb set intantly the thumb position
            If (m_eHitTest = 5) Then
                pvGetThumbPos = m_uRctTLButton.Y2 + (m_uRctBRButton.Y1 - m_uRctTLButton.Y2 - m_lThumbSize) / m_lAbsRange * (m_TargetValue - m_lMin)
            Else
                '// By AMV 09/08/2021 ================================================================================
                '// original: pvGetThumbPos = m_uRctTLButton.Y2 + (m_uRctBRButton.Y1 - m_uRctTLButton.Y2 - m_lThumbSize) / m_lAbsRange * (m_lValue - m_lMin)
                If m_lAbsRange > 0 Then
                  pvGetThumbPos = m_uRctTLButton.Y2 + (m_uRctBRButton.Y1 - m_uRctTLButton.Y2 - m_lThumbSize) / m_lAbsRange * (m_lValue - m_lMin)
                Else
                  pvGetThumbPos = m_uRctTLButton.Y2 + (m_uRctBRButton.Y1 - m_uRctTLButton.Y2 - m_lThumbSize) '/ m_lAbsRange * (m_lValue - m_lMin)
                End If
                '// ================|================================================================|================

            End If
        Case [oHorizontal]
            If (m_eHitTest = 5) Then
                pvGetThumbPos = m_uRctTLButton.X2 + (m_uRctBRButton.X1 - m_uRctTLButton.X2 - m_lThumbSize) / m_lAbsRange * (m_TargetValue - m_lMin)
            Else
                If m_lAbsRange > 0 Then
                    pvGetThumbPos = m_uRctTLButton.X2 + (m_uRctBRButton.X1 - m_uRctTLButton.X2 - m_lThumbSize) / m_lAbsRange * (m_lValue - m_lMin)
                Else
                    pvGetThumbPos = m_uRctTLButton.X2 + (m_uRctBRButton.X1 - m_uRctTLButton.X2 - m_lThumbSize) '/ m_lAbsRange * (m_lValue - m_lMin)
                End If
            End If
    End Select
    
    On Error GoTo 0
End Function

'Get value from the thumb pos
Private Function pvGetScrollPos() As Long
    On Error Resume Next
    Select Case m_eOrientation
        Case [oVertical]
            pvGetScrollPos = m_lMin + (m_lThumbPos - m_uRctTLButton.Y2) / (m_uRctBRButton.Y1 - m_uRctTLButton.Y2 - m_lThumbSize) * m_lAbsRange
        Case [oHorizontal]
            pvGetScrollPos = m_lMin + (m_lThumbPos - m_uRctTLButton.X2) / (m_uRctBRButton.X1 - m_uRctTLButton.X2 - m_lThumbSize) * m_lAbsRange
    End Select
    On Error GoTo 0
End Function

'----------------------------------------------------------------------------------------
' Hit-Test
'----------------------------------------------------------------------------------------
Private Function pvHitTest(ByVal X As Long, ByVal Y As Long) As Long
    Select Case True
        Case PtInRect(m_uRctTLButton, X, Y)
            pvHitTest = HT_TLBUTTON
        Case PtInRect(m_uRctBRButton, X, Y)
            pvHitTest = HT_BRBUTTON
        Case PtInRect(m_uRctTLTrack, X, Y)
            pvHitTest = HT_TLTRACK
        Case PtInRect(m_uRctBRTrack, X, Y)
            pvHitTest = HT_BRTRACK
        Case PtInRect(m_uRctThumb, X, Y)
            pvHitTest = HT_THUMB
    End Select
End Function

Private Sub pvMakePoints(ByVal lPoint As Long, X As Long, Y As Long)
    If (lPoint And &H8000&) Then
        X = &H8000 Or (lPoint And &H7FFF&)
    Else
        X = lPoint And &HFFFF&
    End If
    If (lPoint And &H80000000) Then
        Y = (lPoint \ &H10000) - 1
    Else
        Y = lPoint \ &H10000
    End If
End Sub

'----------------------------------------------------------------------------------------
' Timing
'----------------------------------------------------------------------------------------
Private Sub pvSetTimer(ByVal lTimerID As Long, ByVal ldT As Long)
    Call SetTimer(UserControl.hwnd, lTimerID, ldT, 0)
End Sub

Private Sub pvKillTimer(ByVal lTimerID As Long)
    Call KillTimer(UserControl.hwnd, lTimerID)
    m_eHitTestHot = HT_NOTHING
End Sub

'----------------------------------------------------------------------------------------
' Painting
'----------------------------------------------------------------------------------------
Private Sub pvDrawFlatButton(ByVal hdc As Long, uRct As Rect, ByVal lfArrowDirection As Long, ByVal estate As eFlatButtonStateCts)
    Dim uRctMem As Rect, hDCMem1 As Long, hDCMem2 As Long
    Dim hBmp1 As Long, hBmp2 As Long, hBmpOld1 As Long, hBmpOld2 As Long
    Dim clrBkOld As Long, clrTextOld As Long
    '
    With uRct
        '-- Monochrome bitmap to convert the arrow to black/white mask
        hDCMem1 = CreateCompatibleDC(hdc)
        hBmp1 = CreateBitmap(.X2 - .X1, .Y2 - .Y1, 1, 1, ByVal 0)
        hBmpOld1 = SelectObject(hDCMem1, hBmp1)
        '-- Normal bitmap to draw the arrow into
        hDCMem2 = CreateCompatibleDC(hdc)
        hBmp2 = CreateCompatibleBitmap(hdc, .X2 - .X1, .Y2 - .Y1)
        hBmpOld2 = SelectObject(hDCMem2, hBmp2)
        '-- Draw frame normaly
        Call CopyRect(uRctMem, uRct)
        Call OffsetRect(uRctMem, -.X1, -.Y1)
        Call DrawFrameControl(hDCMem2, uRctMem, DFC_SCROLL, DFCS_FLAT Or lfArrowDirection)
        Select Case estate
            Case [fbsNormal]
                '-- Nothing to do
                Call BitBlt(hdc, .X1, .Y1, .X2 - .X1, .Y2 - .Y1, hDCMem2, 0, 0, vbSrcCopy)
            Case [fbsSelected]
                '-- Invert
                Call InvertRect(hDCMem2, uRctMem)
                Call BitBlt(hdc, .X1, .Y1, .X2 - .X1, .Y2 - .Y1, hDCMem2, 0, 0, vbSrcCopy)
            Case [fbsHot]
                '-- Mask glyph
                Call SetBkColor(hDCMem2, GetSysColor(COLOR_BTNTEXT))
                Call BitBlt(hDCMem1, 0, 0, .X2 - .X1, .Y2 - .Y1, hDCMem2, 0, 0, vbSrcCopy)
                clrBkOld = SetBkColor(hdc, GetSysColor(COLOR_3DHIGHLIGHT))
                clrTextOld = SetTextColor(hdc, GetSysColor(COLOR_3DSHADOW))
                Call BitBlt(hdc, .X1, .Y1, .X2 - .X1, .Y2 - .Y1, hDCMem1, 0, 0, vbSrcCopy)
                Call SetBkColor(hdc, clrBkOld)
                Call SetTextColor(hdc, clrTextOld)
        End Select
    End With
    '-- Clean up
    Call DeleteObject(SelectObject(hDCMem1, hBmpOld1))
    Call DeleteObject(SelectObject(hDCMem2, hBmpOld2))
    Call DeleteDC(hDCMem1)
    Call DeleteDC(hDCMem2)
End Sub

Private Function pvDrawThemePart(ByVal lhDC As Long, ByVal sClass As String, ByVal lPart As Long, ByVal lState As Long, lpRect As Rect) As Boolean
    Dim hTheme As Long
    On Error GoTo Catch
    '
    hTheme = OpenThemeData(UserControl.hwnd, StrPtr(sClass))
    If (hTheme <> 0) Then
        pvDrawThemePart = (DrawThemeBackground(hTheme, lhDC, lPart, lState, lpRect, lpRect) = 0)
    End If
Catch:
    On Error GoTo 0
End Function

'----------------------------------------------------------------------------------------
' Misc.
'----------------------------------------------------------------------------------------
'-- Creating a pattern bitmap (track)
Private Sub pvCreatePatternBrush()
    Dim hBitmap As Long, nPattern(1 To 8) As Integer
    '
    '-- Brush pattern (8x8)
    nPattern(1) = &HAA
    nPattern(2) = &H55
    nPattern(3) = &HAA
    nPattern(4) = &H55
    nPattern(5) = &HAA
    nPattern(6) = &H55
    nPattern(7) = &HAA
    nPattern(8) = &H55
    '-- Create brush from bitmap
    hBitmap = CreateBitmap(8, 8, 1, 1, nPattern(1))
    m_hPatternBrush = CreatePatternBrush(hBitmap)
    Call DeleteObject(hBitmap)
End Sub

'-- Checking environment and Windows theming
Private Sub pvCheckEnvironment()
    'modified by Jason James Newland 2007
    Dim uOSV As OSVERSIONINFO
    '
    m_bIsXP = False
    m_bIsLuna = False
    
    With uOSV
        .dwOSVersionInfoSize = Len(uOSV)
        Call GetVersionEx(uOSV)
        If (.dwPlatformId = 2) Then
            If (.dwMajorVersion >= 5) Then    ' NT based
                If (.dwMinorVersion > 0) Then ' XP
                    m_bIsXP = True
                    m_bIsLuna = pvIsLuna()
                End If
            End If
        End If
    End With
    
End Sub

Private Function pvIsLuna() As Boolean
    'modified by Jason James Newland 2007
    Dim hLib   As Long, lPos  As Long, sTheme As String, sName As String
    '-- Be sure that the theme dll is present
    hLib = LoadLibrary("uxtheme.dll")
    If (hLib <> 0) Then
        '-- Get the theme file name
        sTheme = String$(255, 0)
        Call GetCurrentThemeName(StrPtr(sTheme), Len(sTheme), 0, 0, 0, 0)
        lPos = InStr(1, sTheme, Chr$(0))
        If (lPos > 0) Then
            '-- Get the canonical theme name
            sTheme = Left$(sTheme, lPos - 1)
            sName = String$(255, 0)
            Call GetThemeDocumentationProperty(StrPtr(sTheme), StrPtr("ThemeName"), StrPtr(sName), Len(sName))
            lPos = InStr(1, sName, Chr$(0))
            If (lPos > 0) Then
                '-- Is it Luna or Areo?
                sName = Left$(sName, lPos - 1)
                pvIsLuna = IIf(LenB(sName) <> 0, True, False)
            End If
        End If
        Call FreeLibrary(hLib)
    End If
End Function

'*///////////////////////////////////////////////
'#REGION: "Google" style by Leandro Ascierto
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

Private Function IsDarkColor(ByVal Color As Long) As Boolean
    Dim BGRA(0 To 3) As Byte
    If (Color And &H80000000) Then Color = GetSysColor(Color And &HFF&)
    CopyMemory BGRA(0), Color, 4&
    IsDarkColor = ((CLng(BGRA(0)) + (CLng(BGRA(1) * 3)) + CLng(BGRA(2))) / 2) < 382
End Function

Private Sub FillRectangle(hdc As Long, Rect As Rect, ByVal Color As OLE_COLOR)
    Dim hBrush As Long
    If (Color And &H80000000) Then Color = GetSysColor(Color And &HFF&)
    hBrush = CreateSolidBrush(Color)
    FillRect hdc, Rect, hBrush
    Call DeleteObject(hBrush)
End Sub

Private Sub RoundRectPlus(ByVal hdc As Long, ByVal Left As Long, ByVal Top As Long, _
                       ByVal Width As Long, ByVal Height As Long, ByVal BackColor As Long, ByVal BorderColor As Long, Radius As Long)
    Dim hPen As Long, hBrush As Long
    Dim mPath As Long, hGraphics As Long
    
    GdipCreateFromHDC hdc, hGraphics
    GdipSetSmoothingMode hGraphics, 4 'SmoothingModeAntiAlias
    
    If BackColor <> 0 Then GdipCreateSolidFill BackColor, hBrush
    If BorderColor <> 0 Then GdipCreatePen1 BorderColor, 1 * DpiFactor, &H2, hPen
    
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



Public Function GetWindowsDPI() As Double
    Dim hdc As Long, LPX  As Double
    hdc = GetDC(0)
    LPX = CDbl(GetDeviceCaps(hdc, 88)) ' 88=LOGPIXELSX
    ReleaseDC 0, hdc

    If (LPX = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = LPX / 96#
    End If
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
'*2

Private Sub StartGDIPlus()
    If (GdipToken = 0) Then
        Dim GdipStartupInput As GDIPlusStartupInput
        GdipStartupInput.GdiPlusVersion = 1&
        Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
    End If
End Sub


'#END REGION
'*///////////////////////////////////////////////


'======================================
' Pone en "ALWAYS ON TOP" una ventana
'======================================
Public Sub pOnTop(lhWnd&, state As Boolean, Show As Boolean)
Dim v&
    v = SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
    If (Show) Then
        v = v Or SWP_SHOWWINDOW
    Else
        v = v Or SWP_HIDEWINDOW
    End If
    If (state = False) Then
        Call SetWindowPos(lhWnd&, HWND_NOTOPMOST, 0, 0, 0, 0, v)
    Else
        Call SetWindowPos(lhWnd&, HWND_TOPMOST, 0, 0, 0, 0, v)
    End If
End Sub

'========================================================================================
' UserControl persistent properties
'========================================================================================
Private Sub UserControl_InitProperties()
    '-- Initialization default values
    Let m_lChangeDelay = CHANGEDELAY_DEF
    Let m_lChangeFrequency = CHANGEFREQUENCY_DEF
    Let m_lMin = MIN_DEF
    Let m_lMax = MAX_DEF
    Let m_lValue = VALUE_DEF
    Let m_lSmallChange = SMALLCHANGE_DEF
    Let m_lLargeChange = LARGECHANGE_DEF
    Let m_eOrientation = ORIENTATION_DEF
    Let m_eStyle = STYLE_DEF
    Let m_bShowButtons = SHOWBUTTONS_DEF
    m_bDisableMouseWheelSupport = False
    
    '-- Initialize rectangles
    Let m_lAbsRange = m_lMax - m_lMin
    Call pvSizeButtons
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    m_SinSmoothScrollFactor = 0.15
    m_LonWheelChange = m_lLargeChange
    m_bThumbTooltipEnable = True
    UserControl.BackColor = UserControl.Parent.BackColor
    m_StyleThumbColor = -1
    m_StyleBackColor = -1
End Sub

'*3
'=============================================
'Shagratt: Needed cause if we add from code
' the readproperties is never called
'=============================================
Public Sub AddedFromCodeINIT_AND_DEFAULTS()
   
   If (bReadPropertiesDone) Then Exit Sub
   
    '-- Check OS and Luna theme
    Call pvCheckEnvironment
    
    Call SubclassingStart
    
End Sub

'Shagratt 29/11/19
Public Function TrackMouseWheelOnHwnd(lhWnd&)
    'Keep the handle so we can check if its over its updated position
    ExtWheelHwnd& = lhWnd
End Function

'Shagratt 29/11/19
Public Function TrackMouseWheelOnHwndStop()
    ExtWheelHwnd& = 0
End Function


'==================================
'Scroll Up
'==================================
'(Also for external call)
Public Sub Scroll_UP(Distance_Type&)
Dim dist&

    Select Case (Distance_Type&)
        Case 0: dist = m_lSmallChange
        Case 1: dist = m_lLargeChange
        Case 2: dist = 99999999
    End Select

    Call pvScrollPosDec(dist&)
    Call pvKillTimer(TIMERID_CHANGE1)
    Call pvSetTimer(TIMERID_CHANGE1, m_lChangeDelay)
End Sub

'==================================
'Scroll Down
'==================================
'(Also for external call)
Public Sub Scroll_DOWN(Distance_Type&)
Dim dist&

    Select Case (Distance_Type&)
        Case 0: dist = m_lSmallChange
        Case 1: dist = m_lLargeChange
        Case 2: dist = 99999999
    End Select

    Call pvScrollPosInc(dist&)
    Call pvKillTimer(TIMERID_CHANGE1)
    Call pvSetTimer(TIMERID_CHANGE1, m_lChangeDelay)
End Sub

'Convert a single/double value in a string no matter the regional configuration for decimals
'Used on ReadProperties
Private Function FixDec(s$) As Double
    FixDec = Val(Replace(s$, ",", "."))
End Function


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    bReadPropertiesDone = True
    '-- Bag properties
    With PropBag
        '-- Read inherently-stored properties
        Let UserControl.Enabled = .ReadProperty("Enabled", ENABLED_DEF)
        '-- Read 'memory' properties
        Let m_lMin = .ReadProperty("Min", MIN_DEF)
        Let m_lMax = .ReadProperty("Max", MAX_DEF)
        Let m_lValue = .ReadProperty("Value", VALUE_DEF)
        Let m_lSmallChange = .ReadProperty("SmallChange", SMALLCHANGE_DEF)
        Let m_lLargeChange = .ReadProperty("LargeChange", LARGECHANGE_DEF)
        Let m_lChangeDelay = .ReadProperty("ChangeDelay", CHANGEDELAY_DEF)
        Let m_lChangeFrequency = .ReadProperty("ChangeFrequency", CHANGEFREQUENCY_DEF)
        Let m_eOrientation = .ReadProperty("Orientation", ORIENTATION_DEF)
        Let m_eStyle = .ReadProperty("Style", STYLE_DEF)
            If (m_eStyle = sGoogle) Then StartGDIPlus
        Let m_bShowButtons = .ReadProperty("ShowButtons", SHOWBUTTONS_DEF)
        Let m_bDisableMouseWheelSupport = .ReadProperty("DisableMouseWheelSupport", False)
    End With
    m_SinSmoothScrollFactor = FixDec(PropBag.ReadProperty("SmoothScrollFactor", 0.3))
        If (m_SinSmoothScrollFactor > 1) Then m_SinSmoothScrollFactor = m_SinSmoothScrollFactor / 100
    m_LonWheelChange = PropBag.ReadProperty("WheelChange", LARGECHANGE_DEF)
    '-- Initialize rectangles

    Let m_lAbsRange = m_lMax - m_lMin
    Call pvSizeButtons
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    '-- Run-time?

    If (RunMode) Then
        '-- Check OS and Luna theme
        Call pvCheckEnvironment
        
        Call SubclassingStart
    
    End If
    m_bThumbTooltipEnable = PropBag.ReadProperty("ThumbTooltipEnable", True)
    Set m_StdThumbTooltipFont = PropBag.ReadProperty("ThumbTooltipFont", Ambient.Font) 'UserControl.Font
    
    Set UserControl.Font = m_StdThumbTooltipFont
    m_StyleThumbColor = PropBag.ReadProperty("StyleThumbColor", -1)
    m_StyleBackColor = PropBag.ReadProperty("StyleBackColor", -1)
    m_StyleCurveRadius = PropBag.ReadProperty("StyleCurveRadius", -1)
    If (m_StyleBackColor <> -1) Then
        UserControl.BackColor = m_StyleBackColor
    End If
        
    m_Thumbsize_min = PropBag.ReadProperty("Thumbsize_min", 5)
    m_Thumbsize_max = PropBag.ReadProperty("Thumbsize_max", 95)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Enabled", UserControl.Enabled, ENABLED_DEF)
        Call .WriteProperty("Min", m_lMin, MIN_DEF)
        Call .WriteProperty("Max", m_lMax, MAX_DEF)
        Call .WriteProperty("Value", m_lValue, VALUE_DEF)
        Call .WriteProperty("SmallChange", m_lSmallChange, SMALLCHANGE_DEF)
        Call .WriteProperty("LargeChange", m_lLargeChange, LARGECHANGE_DEF)
        Call .WriteProperty("ChangeDelay", m_lChangeDelay, CHANGEDELAY_DEF)
        Call .WriteProperty("ChangeFrequency", m_lChangeFrequency, CHANGEFREQUENCY_DEF)
        Call .WriteProperty("Orientation", m_eOrientation, ORIENTATION_DEF)
        Call .WriteProperty("Style", m_eStyle, STYLE_DEF)
        Call .WriteProperty("ShowButtons", m_bShowButtons, SHOWBUTTONS_DEF)
        Call .WriteProperty("DisableMouseWheelSupport", m_bDisableMouseWheelSupport, False)
    End With
    Call PropBag.WriteProperty("SmoothScrollFactor", m_SinSmoothScrollFactor, 0.3)
    Call PropBag.WriteProperty("WheelChange", m_LonWheelChange, LARGECHANGE_DEF)
    Call PropBag.WriteProperty("ThumbTooltipEnable", m_bThumbTooltipEnable, True)
    Call PropBag.WriteProperty("ThumbTooltipFont", m_StdThumbTooltipFont, Ambient.Font)
    
    Call PropBag.WriteProperty("StyleThumbColor", m_StyleThumbColor, -1)
    Call PropBag.WriteProperty("StyleBackColor", m_StyleBackColor, -1)
    Call PropBag.WriteProperty("StyleCurveRadius", m_StyleCurveRadius, -1)
    Call PropBag.WriteProperty("Thumbsize_min", m_Thumbsize_min, 5)
    Call PropBag.WriteProperty("Thumbsize_max", m_Thumbsize_max, 95)
End Sub

Public Property Get WheelChange() As Long
    WheelChange = m_LonWheelChange
End Property
Public Property Let WheelChange(ByVal LonValue As Long)
    m_LonWheelChange = LonValue
    PropertyChanged "WheelChange"
End Property


Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enable As Boolean)
    UserControl.Enabled = New_Enable
    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
End Property

Public Property Get Max() As Long
    Max = m_lMax
End Property
Public Property Let Max(ByVal New_Max As Long)
    If (New_Max < m_lMin) Then
        New_Max = m_lMin
    End If
    m_lMax = New_Max
    m_lAbsRange = m_lMax - m_lMin
    If (m_lValue > m_lMax) Then
        m_lValue = m_lMax
    End If
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
End Property

Public Property Get Min() As Long
    Min = m_lMin
End Property
Public Property Let Min(ByVal New_Min As Long)
    If (New_Min > m_lMax) Then
        New_Min = m_lMax
    End If
    m_lMin = New_Min
    m_lAbsRange = m_lMax - m_lMin
    If (m_lValue < m_lMin) Then
        m_lValue = m_lMin
    End If
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
End Property

Public Property Get Value() As Long
    Value = m_lValue
End Property
Public Property Let Value(ByVal New_Value As Long)
    Dim lValuePrev As Long
    
    m_TargetValue = New_Value
    TimSmoothChange.Enabled = False
    '
    If (New_Value < m_lMin) Then
        New_Value = m_lMin
    ElseIf (New_Value > m_lMax) Then
        New_Value = m_lMax
    End If
    lValuePrev = m_lValue
    m_lValue = New_Value
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
    '
    If (m_lValue <> lValuePrev) Then
        RaiseEvent Change
    End If
End Property

Public Property Get SmallChange() As Long
    SmallChange = m_lSmallChange
End Property
Public Property Let SmallChange(ByVal New_SmallChange As Long)
    If (New_SmallChange < 1) Then
        New_SmallChange = 1
    End If
    m_lSmallChange = New_SmallChange
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
End Property

Public Property Get LargeChange() As Long
    LargeChange = m_lLargeChange
End Property
Public Property Let LargeChange(ByVal New_LargeChange As Long)
    If (New_LargeChange < 1) Then
        New_LargeChange = 1
    End If
    m_lLargeChange = New_LargeChange
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    m_LonWheelChange = m_lLargeChange
    Call pvSizeTrack
    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
End Property

Public Property Get ChangeDelay() As Long
    ChangeDelay = m_lChangeDelay
End Property
Public Property Let ChangeDelay(ByVal New_ChangeDelay As Long)
    If (New_ChangeDelay < CHANGEDELAY_MIN) Then
        New_ChangeDelay = CHANGEDELAY_MIN
    End If
    m_lChangeDelay = New_ChangeDelay
End Property

Public Property Get ChangeFrequency() As Long
    ChangeFrequency = m_lChangeFrequency
End Property

Public Property Let ChangeFrequency(ByVal New_ChangeFrequency As Long)
    If (New_ChangeFrequency < CHANGEFREQUENCY_MIN) Then
        New_ChangeFrequency = CHANGEFREQUENCY_MIN
    End If
    m_lChangeFrequency = New_ChangeFrequency
End Property

Public Property Get Orientation() As sbOrientationCts
    Orientation = m_eOrientation
End Property

Public Property Let Orientation(ByVal New_Orientation As sbOrientationCts)
    If (New_Orientation < [oVertical]) Then
        New_Orientation = [oVertical]
    ElseIf (New_Orientation > [oHorizontal]) Then
        New_Orientation = [oHorizontal]
    End If
    m_eOrientation = New_Orientation
    Call pvOnSize
End Property

Public Property Get Style() As sbStyleCts
Attribute Style.VB_Description = "Only on runtime. On IDE you will see standard scrollbars. Google style look better with ShowButtons = False"
    Style = m_eStyle
End Property
Public Property Let Style(ByVal New_Style As sbStyleCts)
    If (New_Style < [sClassic]) Then
        New_Style = [sClassic]
    ElseIf (New_Style > [sGoogle]) Then
        New_Style = [sGoogle]
    End If
    m_eStyle = New_Style
        If (m_eStyle = sGoogle) Then StartGDIPlus
    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
End Property

Public Property Get ShowButtons() As Boolean
Attribute ShowButtons.VB_Description = "Only on runtime. On IDE you will see standard scrollbars"
    ShowButtons = m_bShowButtons
End Property
Public Property Let ShowButtons(ByVal New_ShowButtons As Boolean)
    m_bShowButtons = New_ShowButtons
    Call pvOnSize
End Property

'// Runtime read only
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Get IsXP() As Boolean
    IsXP = m_bIsXP
End Property

Public Property Get IsThemed() As Boolean
    IsThemed = m_bIsLuna
End Property



'========================================================================================
' About
'========================================================================================
Public Sub About()
    Call VBA.MsgBox("ucScrollbar " & VERSION_INFO & " - Shagratt MOD (Original by Carles P.V. 2005)", , "About")
End Sub

'====================================================
'Detect usermode even if its used inside another UC
'====================================================
Public Property Get RunMode() As Boolean
Attribute RunMode.VB_MemberFlags = "400"
On Error Resume Next
    RunMode = True
    RunMode = Ambient.UserMode
    RunMode = Extender.Parent.RunMode
End Property

'Replace with favourite subclassing
Private Sub SubclassingStart()
On Error Resume Next
    If Not (RunMode) Then Exit Sub
    
    bSubclassed = True
    'Paul Caton
    Call sc_Subclass(UserControl.hwnd) 'Optional pass and object with ObjPtr(pic(0))
    Call sc_AddMsg(UserControl.hwnd, WM_PAINT, MSG_BEFORE)
    Call sc_AddMsg(UserControl.hwnd, WM_SIZE, MSG_BEFORE)
    Call sc_AddMsg(UserControl.hwnd, WM_CANCELMODE)
    Call sc_AddMsg(UserControl.hwnd, WM_MOUSEMOVE)
    Call sc_AddMsg(UserControl.hwnd, WM_LBUTTONDOWN)
    Call sc_AddMsg(UserControl.hwnd, WM_LBUTTONUP)
    Call sc_AddMsg(UserControl.hwnd, WM_LBUTTONDBLCLK)
    Call sc_AddMsg(UserControl.hwnd, WM_TIMER)
    Call sc_AddMsg(UserControl.hwnd, WM_SYSCOLORCHANGE)
    Call sc_AddMsg(UserControl.hwnd, WM_MOUSEWHEEL, MSG_BEFORE)
    'Call sc_AddMsg(UserControl.hwnd, WM_KEYDOWN, MSG_BEFORE)

    If (m_bIsXP) Then
        Call sc_AddMsg(UserControl.hwnd, WM_THEMECHANGED)
    End If
    
    'To receive mousewheel and process when is over control
    'Call sc_Subclass(UserControl.Parent.hwnd)
    Call sc_Subclass(UserControl.ContainerHwnd) 'Change suggested by Leandro Ascierto
    Call sc_AddMsg(UserControl.ContainerHwnd, WM_MOUSEWHEEL, MSG_BEFORE)
    Call sc_AddMsg(UserControl.ContainerHwnd, WM_MOUSEMOVE, MSG_AFTER)
    Call sc_AddMsg(UserControl.ContainerHwnd, WM_MOUSELEAVE, MSG_AFTER)

End Sub

Private Sub SubclassEnd()
On Error Resume Next
    If Not (bSubclassed) Then Exit Sub
    
    'sc_UnSubclass UserControl.Parent.hwnd
    sc_UnSubclass UserControl.ContainerHwnd
    sc_UnSubclass UserControl.hwnd
End Sub


'*-

'-SelfSub code------------------------------------------------------------------------------------
'==================================================================================================
' Paul_Caton@hotmail.com
' Copyright free, use and abuse as you see fit.
'
'* v1.0 Re-write of the SelfSub/WinSubHook-2 submission to Planet Source Code............ 20060322
'* v1.1 VirtualAlloc memory to prevent Data Execution Prevention faults on Win64......... 20060324
'* v1.2 Thunk redesigned to handle unsubclassing and memory release...................... 20060325
'* v1.3 Data array scrapped in favour of property accessors.............................. 20060405
'* v1.4 Optional IDE protection added
'*      User-defined callback parameter added
'*      All user routines that pass in a hWnd get additional validation
'*      End removed from zError.......................................................... 20060411
'* v1.5 Added nOrdinal parameter to sc_Subclass
'*      Switched machine-code array from Currency to Long................................ 20060412
'* v1.6 Added an optional callback target object
'*      Added an IsBadCodePtr on the callback address in the thunk prior to callback..... 20060413
'*************************************************************************************************
Private Function sc_Subclass(ByVal lng_hWnd As Long, _
                    Optional ByVal lParamUser As Long = 0, _
                    Optional ByVal nOrdinal As Long = 1, _
                    Optional ByVal oCallback As Object = Nothing, _
                    Optional ByVal bIdeSafety As Boolean = True) As Boolean 'Subclass the specified window handle
'*************************************************************************************************
'* lng_hWnd   - Handle of the window to subclass
'* lParamUser - Optional, user-defined callback parameter
'* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
'* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
'* bIdeSafety - Optional, enable/disable IDE safety measures. NB: you should really only disable IDE safety in a UserControl for design-time subclassing
'*************************************************************************************************
Const CODE_LEN      As Long = 260                                           'Thunk length in bytes
Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))            'Bytes to allocate per thunk, data + code + msg tables
Const PAGE_RWX      As Long = &H40&                                         'Allocate executable memory
Const MEM_COMMIT    As Long = &H1000&                                       'Commit allocated memory
Const MEM_RELEASE   As Long = &H8000&                                       'Release allocated memory flag
Const IDX_EBMODE    As Long = 3                                             'Thunk data index of the EbMode function address
Const IDX_CWP       As Long = 4                                             'Thunk data index of the CallWindowProc function address
Const IDX_SWL       As Long = 5                                             'Thunk data index of the SetWindowsLong function address
Const IDX_FREE      As Long = 6                                             'Thunk data index of the VirtualFree function address
Const IDX_BADPTR    As Long = 7                                             'Thunk data index of the IsBadCodePtr function address
Const IDX_OWNER     As Long = 8                                             'Thunk data index of the Owner object's vTable address
Const IDX_CALLBACK  As Long = 10                                            'Thunk data index of the callback method address
Const IDX_EBX       As Long = 16                                            'Thunk code patch index of the thunk data
Const SUB_NAME      As String = "sc_Subclass"                               'This routine's name
  Dim nAddr         As Long
  Dim nID           As Long
  Dim nMyID         As Long
  
  If IsWindow(lng_hWnd) = 0 Then                                            'Ensure the window handle is valid
    zError SUB_NAME, "Invalid window handle"
    Exit Function
  End If

  nMyID = GetCurrentProcessId                                               'Get this process's ID
  GetWindowThreadProcessId lng_hWnd, nID                                    'Get the process ID associated with the window handle
  If nID <> nMyID Then                                                      'Ensure that the window handle doesn't belong to another process
    zError SUB_NAME, "Window handle belongs to another process"
    Exit Function
  End If
  
  If oCallback Is Nothing Then                                              'If the user hasn't specified the callback owner
    Set oCallback = Me                                                      'Then it is me
  End If
  
  nAddr = zAddressOf(oCallback, nOrdinal)                                   'Get the address of the specified ordinal method
  If nAddr = 0 Then                                                         'Ensure that we've found the ordinal method
    zError SUB_NAME, "Callback method not found"
    Exit Function
  End If
    
  If z_Funk Is Nothing Then                                                 'If this is the first time through, do the one-time initialization
    Set z_Funk = New Collection                                             'Create the hWnd/thunk-address collection
    z_Sc(14) = &HD231C031: z_Sc(15) = &HBBE58960: z_Sc(17) = &H4339F631: z_Sc(18) = &H4A21750C: z_Sc(19) = &HE82C7B8B: z_Sc(20) = &H74&: z_Sc(21) = &H75147539: z_Sc(22) = &H21E80F: z_Sc(23) = &HD2310000: z_Sc(24) = &HE8307B8B: z_Sc(25) = &H60&: z_Sc(26) = &H10C261: z_Sc(27) = &H830C53FF: z_Sc(28) = &HD77401F8: z_Sc(29) = &H2874C085: z_Sc(30) = &H2E8&: z_Sc(31) = &HFFE9EB00: z_Sc(32) = &H75FF3075: z_Sc(33) = &H2875FF2C: z_Sc(34) = &HFF2475FF: z_Sc(35) = &H3FF2473: z_Sc(36) = &H891053FF: z_Sc(37) = &HBFF1C45: z_Sc(38) = &H73396775: z_Sc(39) = &H58627404
    z_Sc(40) = &H6A2473FF: z_Sc(41) = &H873FFFC: z_Sc(42) = &H891453FF: z_Sc(43) = &H7589285D: z_Sc(44) = &H3045C72C: z_Sc(45) = &H8000&: z_Sc(46) = &H8920458B: z_Sc(47) = &H4589145D: z_Sc(48) = &HC4836124: z_Sc(49) = &H1862FF04: z_Sc(50) = &H35E30F8B: z_Sc(51) = &HA78C985: z_Sc(52) = &H8B04C783: z_Sc(53) = &HAFF22845: z_Sc(54) = &H73FF2775: z_Sc(55) = &H1C53FF28: z_Sc(56) = &H438D1F75: z_Sc(57) = &H144D8D34: z_Sc(58) = &H1C458D50: z_Sc(59) = &HFF3075FF: z_Sc(60) = &H75FF2C75: z_Sc(61) = &H873FF28: z_Sc(62) = &HFF525150: z_Sc(63) = &H53FF2073: z_Sc(64) = &HC328&

    z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")                    'Store CallWindowProc function address in the thunk data
    z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")                     'Store the SetWindowLong function address in the thunk data
    z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")                     'Store the VirtualFree function address in the thunk data
    z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr")                  'Store the IsBadCodePtr function address in the thunk data
  End If
  
  z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)                  'Allocate executable memory

  If z_ScMem <> 0 Then                                                      'Ensure the allocation succeeded
    On Error GoTo CatchDoubleSub                                            'Catch double subclassing
      z_Funk.Add z_ScMem, "h" & lng_hWnd                                    'Add the hWnd/thunk-address to the collection
    On Error GoTo 0
  
    If bIdeSafety Then                                                      'If the user wants IDE protection
      z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode")                          'Store the EbMode function address in the thunk data
    End If
    
    z_Sc(IDX_EBX) = z_ScMem                                                 'Patch the thunk data address
    z_Sc(IDX_HWND) = lng_hWnd                                               'Store the window handle in the thunk data
    z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                                   'Store the address of the before table in the thunk data
    z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data
    z_Sc(IDX_OWNER) = ObjPtr(oCallback)                                     'Store the callback owner's object address in the thunk data
    z_Sc(IDX_CALLBACK) = nAddr                                              'Store the callback address in the thunk data
    z_Sc(IDX_PARM_USER) = lParamUser                                        'Store the lParamUser callback parameter in the thunk data
    
    nAddr = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)    'Set the new WndProc, return the address of the original WndProc
    If nAddr = 0 Then                                                       'Ensure the new WndProc was set correctly
      zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
      GoTo ReleaseMemory
    End If
        
    z_Sc(IDX_WNDPROC) = nAddr                                               'Store the original WndProc address in the thunk data
    RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                        'Copy the thunk code/data to the allocated memory
    sc_Subclass = True                                                      'Indicate success
  Else
    zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
  End If
  
  Exit Function                                                             'Exit sc_Subclass

CatchDoubleSub:
  zError SUB_NAME, "Window handle is already subclassed"
  
ReleaseMemory:
  VirtualFree z_ScMem, 0, MEM_RELEASE                                       'sc_Subclass has failed after memory allocation, so release the memory
End Function

'Terminate all subclassing
Private Sub sc_Terminate()
  Dim i As Long

  If Not (z_Funk Is Nothing) Then                                           'Ensure that subclassing has been started
    With z_Funk
      For i = .Count To 1 Step -1                                           'Loop through the collection of window handles in reverse order
        z_ScMem = .Item(i)                                                  'Get the thunk address
        If IsBadCodePtr(z_ScMem) = 0 Then                                   'Ensure that the thunk hasn't already released its memory
          sc_UnSubclass zData(IDX_HWND)                                     'UnSubclass
        End If
      Next i                                                                'Next member of the collection
    End With
    Set z_Funk = Nothing                                                    'Destroy the hWnd/thunk-address collection
  End If
End Sub

'UnSubclass the specified window handle
Private Sub sc_UnSubclass(ByVal lng_hWnd As Long)
  If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
    zError "sc_UnSubclass", "Window handle isn't subclassed"
  Else
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                           'Ensure that the thunk hasn't already released its memory
      zData(IDX_SHUTDOWN) = -1                                              'Set the shutdown indicator
      zDelMsg ALL_MESSAGES, IDX_BTABLE                                      'Delete all before messages
      zDelMsg ALL_MESSAGES, IDX_ATABLE                                      'Delete all after messages
    End If
    z_Funk.Remove "h" & lng_hWnd                                            'Remove the specified window handle from the collection
  End If
End Sub

'Add the message value to the window handle's specified callback table
Private Sub sc_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    If When And MSG_BEFORE Then                                             'If the message is to be added to the before original WndProc table...
      zAddMsg uMsg, IDX_BTABLE                                              'Add the message to the before table
    End If
    If When And MSG_AFTER Then                                              'If message is to be added to the after original WndProc table...
      zAddMsg uMsg, IDX_ATABLE                                              'Add the message to the after table
    End If
  End If
End Sub

'Delete the message value from the window handle's specified callback table
Private Sub sc_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    If When And MSG_BEFORE Then                                             'If the message is to be deleted from the before original WndProc table...
      zDelMsg uMsg, IDX_BTABLE                                              'Delete the message from the before table
    End If
    If When And MSG_AFTER Then                                              'If the message is to be deleted from the after original WndProc table...
      zDelMsg uMsg, IDX_ATABLE                                              'Delete the message from the after table
    End If
  End If
End Sub

'Call the original WndProc
Private Function sc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    sc_CallOrigWndProc = _
        CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
  End If
End Function

'Get the subclasser lParamUser callback parameter
Private Property Get sc_lParamUser(ByVal lng_hWnd As Long) As Long
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    sc_lParamUser = zData(IDX_PARM_USER)                                    'Get the lParamUser callback parameter
  End If
End Property

'Let the subclasser lParamUser callback parameter
Private Property Let sc_lParamUser(ByVal lng_hWnd As Long, ByVal NewValue As Long)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    zData(IDX_PARM_USER) = NewValue                                         'Set the lParamUser callback parameter
  End If
End Property

'-The following routines are exclusively for the sc_ subclass routines----------------------------

'Add the message to the specified table of the window handle
Private Sub zAddMsg(ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim nBase  As Long                                                        'Remember z_ScMem
  Dim i      As Long                                                        'Loop index

  nBase = z_ScMem                                                            'Remember z_ScMem so that we can restore its value on exit
  z_ScMem = zData(nTable)                                                    'Map zData() to the specified table

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
    nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
  Else
    nCount = zData(0)                                                       'Get the current table entry count
    If nCount >= MSG_ENTRIES Then                                           'Check for message table overflow
      zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
      GoTo Bail
    End If

    For i = 1 To nCount                                                     'Loop through the table entries
      If zData(i) = 0 Then                                                  'If the element is free...
        zData(i) = uMsg                                                     'Use this element
        GoTo Bail                                                           'Bail
      ElseIf zData(i) = uMsg Then                                           'If the message is already in the table...
        GoTo Bail                                                           'Bail
      End If
    Next i                                                                  'Next message table entry

    nCount = i                                                              'On drop through: i = nCount + 1, the new table entry count
    zData(nCount) = uMsg                                                    'Store the message in the appended table entry
  End If

  zData(0) = nCount                                                         'Store the new table entry count
Bail:
  z_ScMem = nBase                                                           'Restore the value of z_ScMem
End Sub

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim nBase  As Long                                                        'Remember z_ScMem
  Dim i      As Long                                                        'Loop index

  nBase = z_ScMem                                                           'Remember z_ScMem so that we can restore its value on exit
  z_ScMem = zData(nTable)                                                   'Map zData() to the specified table

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
    zData(0) = 0                                                            'Zero the table entry count
  Else
    nCount = zData(0)                                                       'Get the table entry count
    
    For i = 1 To nCount                                                     'Loop through the table entries
      If zData(i) = uMsg Then                                               'If the message is found...
        zData(i) = 0                                                        'Null the msg value -- also frees the element for re-use
        GoTo Bail                                                           'Bail
      End If
    Next i                                                                  'Next message table entry
    
    zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
  End If
  
Bail:
  z_ScMem = nBase                                                           'Restore the value of z_ScMem
End Sub

'Error handler
Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String)
  App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
  zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                   'Get the specified procedure address
  Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
End Function

'Map zData() to the thunk address for the specified window handle
Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long
  If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
    zError "zMap_hWnd", "Subclassing hasn't been started"
  Else
    On Error GoTo Catch                                                     'Catch unsubclassed window handles
    z_ScMem = z_Funk("h" & lng_hWnd)                                        'Get the thunk address
    zMap_hWnd = z_ScMem
  End If
  
  Exit Function                                                             'Exit returning the thunk address

Catch:
  zError "zMap_hWnd", "Window handle isn't subclassed"
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim i     As Long                                                         'Loop index
  Dim J     As Long                                                         'Loop limit
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, i, bSub) Then                                 'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, i, bSub) Then                              'Probe for a Form method
      If Not zProbe(nAddr + &H7A4, i, bSub) Then                            'Probe for a UserControl method
        Exit Function                                                       'Bail...
      End If
    End If
  End If
  
  i = i + 4                                                                 'Bump to the next entry
  J = i + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
  Do While i < J
    RtlMoveMemory VarPtr(nAddr), i, 4                                       'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    
    i = i + 4                                                             'Next vTable entry
  Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Function                                                       'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
End Function

Private Property Get zData(ByVal nIndex As Long) As Long
  RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4
End Property

Private Property Let zData(ByVal nIndex As Long, ByVal nValue As Long)
  RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4
End Property

'*-
'-Subclass callback, usually ordinal #1, the last method in this source file----------------------
Private Sub zWndProc1(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lhWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Object)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter
'*************************************************************************************************
On Error GoTo Err
Dim uPS As PAINTSTRUCT

    Select Case lhWnd
        Case UserControl.hwnd
            Select Case uMsg
            Case WM_PAINT
                Call BeginPaint(lhWnd, uPS)
                Call pvOnPaint(uPS.hdc)
                Call EndPaint(lhWnd, uPS)
                bHandled = True: lReturn = 0
            Case WM_SIZE
                Call pvOnSize
                bHandled = True: lReturn = 0
            Case WM_LBUTTONDOWN
                Call pvOnMouseDown(wParam, lParam)
            Case WM_MOUSEMOVE
                Call pvOnMouseMove(wParam, lParam)
            Case WM_LBUTTONUP, WM_CANCELMODE
                Call pvOnMouseUp
            Case WM_LBUTTONDBLCLK
                Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            Case WM_TIMER
                Call pvOnTimer(wParam)
            Case WM_SYSCOLORCHANGE
                Call pvOnSysColorChange
            Case WM_THEMECHANGED
                Call pvOnThemeChanged
            'Case WM_MOUSEWHEEL
            '    Not needed. We get the MouseWheel from the parent. This way we dont need focus on the scrollbar
            
            Case WM_KEYDOWN
                'PGUP
                If (wParam = 33) Then
                    Scroll_UP 1
                'PGDW
                ElseIf (wParam = 34) Then
                    Scroll_DOWN 1
                'HOME
                ElseIf (wParam = 36) Then
                    Scroll_UP 2
                'END
                ElseIf (wParam = 35) Then
                    Scroll_DOWN 2
                End If
                
            End Select
    Case Else
        
'        'sLog UserControl.hwnd & "," & lhWnd & "-[" & Hex(uMsg) & "] "
'        'Dim pt As POINTAPI, r As RECT
'        If (uMsg = WM_KEYDOWN) Then
'             Debug.Print ("SC->KD ")
''            'PGUP
''            If (wParam = 33) Then
''                Scroll_UP 1
''            'PGDOWN
''            ElseIf (wParam = 34) Then
''                Scroll_DOWN 1
''            'HOME
''            ElseIf (wParam = 36) Then
''                Scroll_UP 2
''            'END
''            ElseIf (wParam = 35) Then
''                Scroll_DOWN 2
''            End If
'
''            GetCursorPos pt
''            GetWindowRect UserControl.hwnd, r
''            'If is used over the scrollbar (any part) process it
''            'If Mouse Wheel is used over the scrollbar (any part) process it
''            If (PtInRect(r, pt.x, pt.y)) Then
''                bHandled = True: lReturn = 0
''                'PGUP
''                If (wParam = 33) Then
''                    Call pvScrollPosDec(m_lLargeChange)
''                    Call pvKillTimer(TIMERID_CHANGE1)
''                    Call pvSetTimer(TIMERID_CHANGE1, m_lChangeDelay)
''
''                'PGDW
''                ElseIf (wParam = 34) Then
''                    Call pvScrollPosInc(m_lLargeChange)
''                    Call pvKillTimer(TIMERID_CHANGE1)
''                    Call pvSetTimer(TIMERID_CHANGE1, m_lChangeDelay)
''
''                'HOME
''                ElseIf (wParam = 36) Then
''                    Call pvScrollPosDec(999999)
''                    Call pvKillTimer(TIMERID_CHANGE1)
''                    Call pvSetTimer(TIMERID_CHANGE1, m_lChangeDelay)
''
''                'END
''                ElseIf (wParam = 35) Then
''                    Call pvScrollPosInc(999999)
''                    Call pvKillTimer(TIMERID_CHANGE1)
''                    Call pvSetTimer(TIMERID_CHANGE1, m_lChangeDelay)
''
''                Else
''                    bHandled = False
''                End If
''
''            End If
'        End If

        Select Case uMsg
            Case WM_MOUSEMOVE
                
                Dim ET As TRACKMOUSEEVENTTYPE
                'initialize structure
                ET.cbSize = Len(ET)
                ET.hwndTrack = lhWnd
                ET.dwFlags = TME_LEAVE
                'start the tracking
                TrackMouseEvent ET
                If m_bMouseIn = False Then
                    m_bMouseIn = True
                    RaiseEvent ContainerMouseEnter
                End If
            Case WM_MOUSELEAVE
           
                If m_bMouseIn Then
                    m_bMouseIn = False
                   
                    RaiseEvent ContainerMouseLeave
                End If
        
        
            Case WM_MOUSEWHEEL
                Dim PT As POINTAPI, R As Rect
                GetCursorPos PT
                GetWindowRect UserControl.hwnd, R
                'If Mouse Wheel is used over the scrollbar (any part) process it
                If (PtInRect(R, PT.X, PT.Y) And (UserControl.Extender.Visible)) Then
                    'sLog UserControl.hwnd & " IN OVER"
                    bHandled = True: lReturn = 0
                
                'If its tracking Mouse Wheel over another object
                ElseIf (ExtWheelHwnd& <> 0) Then
                    GetWindowRect ExtWheelHwnd&, R
                    'GetWindowRect UserControl.Parent.hwnd, r
                    'Check if mouse if over the obj tracking area
                    If (PtInRect(R, PT.X, PT.Y)) Then
                        'sLog UserControl.hwnd & " EXT OVER EXT"
                        bHandled = True: lReturn = 0
                    End If
                End If
                
                If (bHandled) Then
                
                    'Have a Horizontal Scrollbar attached? check for SHIFT key pressed
                    If Not (m_ucScrollbarH Is Nothing) Then
                        GetWindowRect m_ucScrollbarH.hwnd, R ' by LeandroA
                        If GetKeyState(vbKeyShift) And KEY_DOWN Or PtInRect(R, PT.X, PT.Y) Then                        'Relay the scroll to the Horizontal scrollbar
                            If (wParam > 0) Then
                                m_ucScrollbarH.WheelScrollTopLeft
                            Else
                                m_ucScrollbarH.WheelScrollBotRight
                            End If
                            
                            Exit Sub
                        End If
                    End If
                
                    If (wParam > 0) Then
                        'Top/Left ...scroll up
                        WheelScrollTopLeft
                    Else
                        'Bottom/Right ...scroll down
                        WheelScrollBotRight
                    End If
                End If
        
        End Select
    End Select

Exit Sub
Err:
    'Debug.Print Err.Number
End Sub







