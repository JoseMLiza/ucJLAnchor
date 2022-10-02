VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------
'Autor: Jose Liza
'ClassName: clsControl
'Date: 12/06/2022
'Version: 0.0.1
'-----------------------------
Option Explicit
'Datos del control
Private m_ParentTypeName        As String
Private m_ParentName            As String
Private m_ParentIndex           As String
Private m_TypeName              As String
Private m_Name                  As String
Private m_ControlIndex          As String
Private m_Index                 As Long
Private m_Hwnd                  As Long
Private m_ScaleModeParent       As ScaleModeConstants
Private m_Left                  As Long
Private m_Top                   As Long
Private m_Right                 As Long
Private m_Bottom                As Long
'--
Private m_MinWidth              As Long
Private m_MinHeight             As Long
'--
Private m_UseModePercent        As Integer
'--
Private m_UseLeftPercent        As Boolean
Private m_LeftPercent           As Single
Private m_LeftPercentStatic     As Long
'--
Private m_UseTopPercent         As Boolean
Private m_TopPercent            As Single
Private m_TopPercentStatic      As Long
'--
Private m_UseWidthPercent       As Boolean
Private m_WidthPercent          As Single
Private m_RightPercentStatic    As Long
'--
Private m_UseHeightPercent      As Boolean
Private m_HeightPercent         As Single
Private m_BottomPercentStatic   As Long
'--
Private m_AnchorLeft            As Boolean
Private m_AnchorTop             As Boolean
Private m_AnchorRight           As Boolean
Private m_AnchorBottom          As Boolean

'm_ParentTypeName           As String
Public Property Get ParentTypeName() As String
    ParentTypeName = m_ParentTypeName
End Property
Public Property Let ParentTypeName(ByVal newValue As String)
    m_ParentTypeName = newValue
    PropertyChanged "ParentTypeName"
End Property

'm_ParentName               As String
Public Property Get ParentName() As String
    ParentName = m_ParentName
End Property
Public Property Let ParentName(ByVal newValue As String)
    m_ParentName = newValue
    PropertyChanged "ParentName"
End Property

'm_ParentIndex              As String
Public Property Get ParentIndex() As String
    ParentIndex = m_ParentIndex
End Property
Public Property Let ParentIndex(ByVal newValue As String)
    m_ParentIndex = newValue
    PropertyChanged "ParentIndex"
End Property

'm_TypeName                 As String
Public Property Get TypeName() As String
    TypeName = m_TypeName
End Property
Public Property Let TypeName(ByVal newValue As String)
    m_TypeName = newValue
    PropertyChanged "TypeName"
End Property

'm_Name                     As String
Public Property Get name() As String
    name = m_Name
End Property
Public Property Let name(ByVal newValue As String)
    m_Name = newValue
    PropertyChanged "Name"
End Property

'm_ControlIndex             As String
Public Property Get ControlIndex() As String
    ControlIndex = m_ControlIndex
End Property
Public Property Let ControlIndex(ByVal newValue As String)
    m_ControlIndex = newValue
    PropertyChanged "ControlIndex"
End Property

'm_Index                    As Long
Public Property Get Index() As Long
    Index = m_Index
End Property
Public Property Let Index(ByVal newValue As Long)
    m_Index = newValue
    PropertyChanged "Index"
End Property

'm_hWnd                     As Long
Public Property Get hWnd() As Long
    hWnd = m_Hwnd
End Property
Public Property Let hWnd(ByVal newValue As Long)
    m_Hwnd = newValue
    PropertyChanged "hWnd"
End Property

'm_ScaleModeParent          As ScaleModeConstants
Public Property Get ScaleModeParent() As ScaleModeConstants
    ScaleModeParent = m_ScaleModeParent
End Property
Public Property Let ScaleModeParent(ByVal newValue As ScaleModeConstants)
    m_ScaleModeParent = newValue
    PropertyChanged "ScaleModeParent"
End Property

'm_Left                     As Long
Public Property Get Left() As Long
    Left = m_Left
End Property
Public Property Let Left(ByVal newValue As Long)
    m_Left = newValue
    PropertyChanged "Left"
End Property

'm_Top                      As Long
Public Property Get Top() As Long
    Top = m_Top
End Property
Public Property Let Top(ByVal newValue As Long)
    m_Top = newValue
    PropertyChanged "Top"
End Property

'm_Right                    As Long
Public Property Get Right() As Long
    Right = m_Right
End Property
Public Property Let Right(ByVal newValue As Long)
    m_Right = newValue
    PropertyChanged "Right"
End Property

'm_Bottom                   As Long
Public Property Get Bottom() As Long
    Bottom = m_Bottom
End Property
Public Property Let Bottom(ByVal newValue As Long)
    m_Bottom = newValue
    PropertyChanged "Bottom"
End Property

'm_MinWidth                 As Long
Public Property Get MinWidth() As Long
    MinWidth = m_MinWidth
End Property
Public Property Let MinWidth(ByVal newValue As Long)
    m_MinWidth = newValue
    PropertyChanged "MinWidth"
End Property

'm_MinHeight                As Long
Public Property Get MinHeight() As Long
    MinHeight = m_MinHeight
End Property
Public Property Let MinHeight(ByVal newValue As Long)
    m_MinHeight = newValue
    PropertyChanged "MinHeight"
End Property

'm_UseModePercent           As Integer
Public Property Get UseModePercent() As Integer
    UseModePercent = m_UseModePercent
End Property
Public Property Let UseModePercent(ByVal newValue As Integer)
    m_UseModePercent = newValue
    PropertyChanged "UseModePercent"
End Property

'm_UseLeftPercent           As boolean
Public Property Get UseLeftPercent() As Boolean
    UseLeftPercent = m_UseLeftPercent
End Property
Public Property Let UseLeftPercent(ByVal newValue As Boolean)
    m_UseLeftPercent = newValue
    PropertyChanged "UseLeftPercent"
End Property

'm_LeftPercent              As Single
Public Property Get LeftPercent() As Single
    LeftPercent = m_LeftPercent
End Property
Public Property Let LeftPercent(ByVal newValue As Single)
    m_LeftPercent = newValue
    PropertyChanged "LeftPercent"
End Property

'm_LeftPercentStatic        As Long
Public Property Get LeftPercentStatic() As Long
    LeftPercentStatic = m_LeftPercentStatic
End Property
Public Property Let LeftPercentStatic(ByVal newValue As Long)
    m_LeftPercentStatic = newValue
    PropertyChanged "LeftPercentStatic"
End Property

'm_UseTopPercent            As boolean
Public Property Get UseTopPercent() As Boolean
    UseTopPercent = m_UseTopPercent
End Property
Public Property Let UseTopPercent(ByVal newValue As Boolean)
    m_UseTopPercent = newValue
    PropertyChanged "UseTopPercent"
End Property

'm_TopPercent               As Single
Public Property Get TopPercent() As Single
    TopPercent = m_TopPercent
End Property
Public Property Let TopPercent(ByVal newValue As Single)
    m_TopPercent = newValue
    PropertyChanged "TopPercent"
End Property

'm_TopPercentStatic         As Long
Public Property Get TopPercentStatic() As Long
    TopPercentStatic = m_TopPercentStatic
End Property
Public Property Let TopPercentStatic(ByVal newValue As Long)
    m_TopPercentStatic = newValue
    PropertyChanged "TopPercentStatic"
End Property

'm_UseWidthPercent          As boolean
Public Property Get UseWidthPercent() As Boolean
    UseWidthPercent = m_UseWidthPercent
End Property
Public Property Let UseWidthPercent(ByVal newValue As Boolean)
    m_UseWidthPercent = newValue
    PropertyChanged "UseWidthPercent"
End Property

'm_WidthPercent             As Single
Public Property Get WidthPercent() As Single
    WidthPercent = m_WidthPercent
End Property
Public Property Let WidthPercent(ByVal newValue As Single)
    m_WidthPercent = newValue
    PropertyChanged "WidthPercent"
End Property

'm_RightPercentStatic       As Long
Public Property Get RightPercentStatic() As Long
    RightPercentStatic = m_RightPercentStatic
End Property
Public Property Let RightPercentStatic(ByVal newValue As Long)
    m_RightPercentStatic = newValue
    PropertyChanged "RightPercentStatic"
End Property

'm_UseHeightPercent         As boolean
Public Property Get UseHeightPercent() As Boolean
    UseHeightPercent = m_UseHeightPercent
End Property
Public Property Let UseHeightPercent(ByVal newValue As Boolean)
    m_UseHeightPercent = newValue
    PropertyChanged "UseHeightPercent"
End Property

'm_HeightPercent            As Single
Public Property Get HeightPercent() As Single
    HeightPercent = m_HeightPercent
End Property
Public Property Let HeightPercent(ByVal newValue As Single)
    m_HeightPercent = newValue
    PropertyChanged "HeightPercent"
End Property

'm_BottomPercentStatic      As Long
Public Property Get BottomPercentStatic() As Long
    BottomPercentStatic = m_BottomPercentStatic
End Property
Public Property Let BottomPercentStatic(ByVal newValue As Long)
    m_BottomPercentStatic = newValue
    PropertyChanged "BottomPercentStatic"
End Property

'm_AnchorLeft               As Boolean
Public Property Get AnchorLeft() As Boolean
    AnchorLeft = m_AnchorLeft
End Property
Public Property Let AnchorLeft(ByVal newValue As Boolean)
    m_AnchorLeft = newValue
    PropertyChanged "AnchorLeft"
End Property

'm_AnchorTop                As Boolean
Public Property Get AnchorTop() As Boolean
    AnchorTop = m_AnchorTop
End Property
Public Property Let AnchorTop(ByVal newValue As Boolean)
    m_AnchorTop = newValue
    PropertyChanged "AnchorTop"
End Property

'm_AnchorRight              As Boolean
Public Property Get AnchorRight() As Boolean
    AnchorRight = m_AnchorRight
End Property
Public Property Let AnchorRight(ByVal newValue As Boolean)
    m_AnchorRight = newValue
    PropertyChanged "AnchorRight"
End Property

'm_AnchorBottom             As Boolean
Public Property Get AnchorBottom() As Boolean
    AnchorBottom = m_AnchorBottom
End Property
Public Property Let AnchorBottom(ByVal newValue As Boolean)
    m_AnchorBottom = newValue
    PropertyChanged "AnchorBottom"
End Property

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_ParentTypeName = .ReadProperty("ParentTypeName", "")
        m_ParentName = .ReadProperty("ParentName", "")
        m_ParentIndex = .ReadProperty("ParentIndex", "")
        m_TypeName = .ReadProperty("TypeName", "")
        m_Name = .ReadProperty("Name", "")
        m_ControlIndex = .ReadProperty("ControlIndex", "")
        m_Index = .ReadProperty("Index", 0)
        m_Hwnd = .ReadProperty("hWnd", 0)
        m_ScaleModeParent = .ReadProperty("ScaleModeParent", 0)
        m_Left = .ReadProperty("Left", 0)
        m_Top = .ReadProperty("Top", 0)
        m_Right = .ReadProperty("Right", 0)
        m_Bottom = .ReadProperty("Bottom", 0)
        '--
        m_MinWidth = .ReadProperty("MinWidth", 0)
        m_MinHeight = .ReadProperty("MinHeight", 0)
        '--
        m_UseModePercent = .ReadProperty("UseModePercent", 0)
        '--
        m_UseLeftPercent = .ReadProperty("UseLeftPercent", 0)
        m_LeftPercent = .ReadProperty("LeftPercent", 0)
        m_LeftPercentStatic = .ReadProperty("LeftPercentStatic", 0)
        '--
        m_UseTopPercent = .ReadProperty("UseTopPercent", 0)
        m_TopPercent = .ReadProperty("TopPercent", 0)
        m_TopPercentStatic = .ReadProperty("TopPercentStatic", 0)
        '--
        m_UseWidthPercent = .ReadProperty("UseWidthPercent", 0)
        m_WidthPercent = .ReadProperty("WidthPercent", 0)
        m_RightPercentStatic = .ReadProperty("RightPercentStatic", 0)
        '--
        m_UseHeightPercent = .ReadProperty("UseHeightPercent", 0)
        m_HeightPercent = .ReadProperty("HeightPercent", 0)
        m_BottomPercentStatic = .ReadProperty("BottomPercentStatic", 0)
        '--
        m_AnchorLeft = .ReadProperty("AnchorLeft", False)
        m_AnchorTop = .ReadProperty("AnchorTop", False)
        m_AnchorRight = .ReadProperty("AnchorRight", False)
        m_AnchorBottom = .ReadProperty("AnchorBottom", False)
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "ParentTypeName", m_ParentTypeName, ""
        .WriteProperty "ParentName", m_ParentName, ""
        .WriteProperty "ParentIndex", m_ParentIndex, ""
        .WriteProperty "TypeName", m_TypeName, ""
        .WriteProperty "Name", m_Name, ""
        .WriteProperty "ControlIndex", m_ControlIndex, ""
        .WriteProperty "Index", m_Index, 0
        .WriteProperty "hWnd", m_Hwnd, 0
        .WriteProperty "ScaleModeParent", m_ScaleModeParent, 0
        .WriteProperty "Left", m_Left, 0
        .WriteProperty "Top", m_Top, 0
        .WriteProperty "Right", m_Right, 0
        .WriteProperty "Bottom", m_Bottom, 0
        '--
        .WriteProperty "MinWidth", m_MinWidth, 0
        .WriteProperty "MinHeight", m_MinHeight, 0
        '--
        .WriteProperty "UseModePercent", m_UseModePercent, 0
        '--
        .WriteProperty "UseLeftPercent", m_UseLeftPercent, 0
        .WriteProperty "LeftPercent", m_LeftPercent, 0
        .WriteProperty "LeftPercentStatic", m_LeftPercentStatic, 0
        '--
        .WriteProperty "UseTopPercent", m_UseTopPercent, 0
        .WriteProperty "TopPercent", m_TopPercent, 0
        .WriteProperty "TopPercentStatic", m_TopPercentStatic, 0
        '--
        .WriteProperty "UseWidthPercent", m_UseWidthPercent, 0
        .WriteProperty "WidthPercent", m_WidthPercent, 0
        .WriteProperty "RightPercentStatic", m_RightPercentStatic, 0
        '--
        .WriteProperty "UseHeightPercent", m_UseHeightPercent, 0
        .WriteProperty "HeightPercent", m_HeightPercent, 0
        .WriteProperty "BottomPercentStatic", m_BottomPercentStatic, 0
        '--
        .WriteProperty "AnchorLeft", m_AnchorLeft, False
        .WriteProperty "AnchorTop", m_AnchorTop, False
        .WriteProperty "AnchorRight", m_AnchorRight, False
        .WriteProperty "AnchorBottom", m_AnchorBottom, False
    End With
End Sub
