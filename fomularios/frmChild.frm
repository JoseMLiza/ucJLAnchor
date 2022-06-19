VERSION 5.00
Object = "{62ED814B-A922-424D-833A-57C22F97CFE7}#3.5#0"; "JLAnchor.ocx"
Begin VB.Form frmChild 
   Caption         =   "Child"
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4425
   ScaleWidth      =   8970
   Begin JLAnchor.ucJLAnchor ucJLAnchor1 
      Height          =   495
      Left            =   8400
      TabIndex        =   6
      Top             =   120
      Width           =   495
      _ExtentX        =   847
      _ExtentY        =   847
      ControlsCount   =   7
      BeginProperty Control_1 {F01A6FF4-B13C-432C-BCBF-2B9C8BFB03DE} 
         ParentTypeName  =   "Form"
         ParentName      =   "frmChild"
         TypeName        =   "Frame"
         Name            =   "Frame1"
         Object.Index           =   1
         hWnd            =   9112978
         Object.Left            =   4485
         Object.Top             =   840
         Bottom          =   105
         UseLeftPercent  =   -1  'True
         LeftPercent     =   50
         TopPercent      =   18.983
         UseWidthPercent =   -1  'True
         WidthPercent    =   50
         HeightPercent   =   78.644
         AnchorTop       =   -1  'True
         AnchorBottom    =   -1  'True
      EndProperty
      BeginProperty Control_2 {F01A6FF4-B13C-432C-BCBF-2B9C8BFB03DE} 
         ParentTypeName  =   "Frame"
         ParentName      =   "Frame1"
         TypeName        =   "TextBox"
         Name            =   "Text4"
         Object.Index           =   2
         hWnd            =   2229188
         Object.Left            =   480
         Object.Top             =   1200
         Right           =   1470
         Bottom          =   1785
         LeftPercent     =   50
         TopPercent      =   18.983
         WidthPercent    =   50
         HeightPercent   =   18.983
         AnchorLeft      =   -1  'True
         AnchorTop       =   -1  'True
         AnchorRight     =   -1  'True
      EndProperty
      BeginProperty Control_3 {F01A6FF4-B13C-432C-BCBF-2B9C8BFB03DE} 
         ParentTypeName  =   "Frame"
         ParentName      =   "Frame1"
         TypeName        =   "TextBox"
         Name            =   "Text3"
         Object.Index           =   3
         hWnd            =   7147782
         Object.Left            =   480
         Object.Top             =   480
         Right           =   1470
         Bottom          =   2385
         LeftPercent     =   50
         TopPercent      =   18.983
         WidthPercent    =   50
         HeightPercent   =   18.983
         AnchorLeft      =   -1  'True
         AnchorTop       =   -1  'True
         AnchorRight     =   -1  'True
      EndProperty
      BeginProperty Control_4 {F01A6FF4-B13C-432C-BCBF-2B9C8BFB03DE} 
         ParentTypeName  =   "Form"
         ParentName      =   "frmChild"
         TypeName        =   "PictureBox"
         Name            =   "Picture1"
         Object.Index           =   4
         hWnd            =   1839316
         Object.Top             =   840
         Right           =   4485
         Bottom          =   90
         UseLeftPercent  =   -1  'True
         TopPercent      =   18.983
         UseWidthPercent =   -1  'True
         WidthPercent    =   50
         HeightPercent   =   78.983
         AnchorTop       =   -1  'True
         AnchorBottom    =   -1  'True
      EndProperty
      BeginProperty Control_5 {F01A6FF4-B13C-432C-BCBF-2B9C8BFB03DE} 
         ParentTypeName  =   "Form"
         ParentName      =   "frmChild"
         TypeName        =   "TextBox"
         Name            =   "Text1"
         Object.Index           =   5
         hWnd            =   2493862
         Object.Left            =   1440
         Object.Top             =   120
         Right           =   675
         Bottom          =   3810
         LeftPercent     =   16.054
         TopPercent      =   2.712
         WidthPercent    =   76.421
         HeightPercent   =   11.186
         AnchorLeft      =   -1  'True
         AnchorTop       =   -1  'True
         AnchorRight     =   -1  'True
      EndProperty
      BeginProperty Control_6 {F01A6FF4-B13C-432C-BCBF-2B9C8BFB03DE} 
         ParentTypeName  =   "Form"
         ParentName      =   "frmChild"
         TypeName        =   "CommandButton"
         Name            =   "Command1"
         Object.Index           =   6
         hWnd            =   5902796
         Object.Left            =   120
         Object.Top             =   120
         Right           =   7635
         Bottom          =   3810
         LeftPercent     =   1.338
         TopPercent      =   2.712
         WidthPercent    =   13.545
         HeightPercent   =   11.186
         AnchorLeft      =   -1  'True
         AnchorTop       =   -1  'True
      EndProperty
      BeginProperty Control_7 {F01A6FF4-B13C-432C-BCBF-2B9C8BFB03DE} 
         ParentTypeName  =   "PictureBox"
         ParentName      =   "Picture1"
         TypeName        =   "LabelPlus"
         Name            =   "LabelPlus1"
         Object.Index           =   7
         Object.Left            =   240
         Object.Top             =   480
         Right           =   1410
         Bottom          =   2460
         LeftPercent     =   5.424
         TopPercent      =   13.974
         WidthPercent    =   62.712
         HeightPercent   =   14.41
         AnchorLeft      =   -1  'True
         AnchorTop       =   -1  'True
         AnchorRight     =   -1  'True
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3480
      Left            =   4485
      TabIndex        =   3
      Top             =   840
      Width           =   4485
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   480
         TabIndex        =   5
         Text            =   "Text4"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Height          =   615
         Left            =   480
         TabIndex        =   4
         Text            =   "Text3"
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   0
      ScaleHeight     =   3435
      ScaleWidth      =   4425
      TabIndex        =   2
      Top             =   840
      Width           =   4485
      Begin Proyecto1.LabelPlus LabelPlus1 
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         BackColor       =   15296031
         Border          =   -1  'True
         BorderColor     =   16232825
         BorderCornerLeftTop=   5
         BorderCornerRightTop=   5
         BorderCornerBottomRight=   5
         BorderCornerBottomLeft=   5
         BorderWidth     =   1
         CaptionAlignmentH=   1
         CaptionAlignmentV=   1
         Caption         =   "frmChild.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Sans"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         ShadowSize      =   2
         ShadowColor     =   16232825
         ShadowColorOpacity=   100
         CallOutAlign    =   0
         CallOutWidth    =   0
         CallOutLen      =   0
         MousePointer    =   0
         BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IconForeColor   =   0
         IconOpacity     =   0
         PictureArr      =   0
      End
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   6855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
