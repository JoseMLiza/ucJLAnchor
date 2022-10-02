VERSION 5.00
Object = "{917DED51-9D95-4D2C-9431-281FAC2C6FF4}#2.0#0"; "JLAnchor.ocx"
Begin VB.Form frmTest 
   Caption         =   "Test ScaleMode"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   280
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   416
   Begin JLAnchor.ucJLAnchor ucJLAnchor1 
      Height          =   480
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      IconPresent     =   -1  'True
      FormIcon        =   "frmTest.frx":0000
      ControlsCount   =   7
      BeginProperty Control_1 {881E793B-2C4F-4BDE-963C-25AE290F2EA6} 
         ParentTypeName  =   "Form"
         ParentName      =   "frmTest"
         TypeName        =   "SSTabEx"
         Name            =   "SSTabEx1"
         Object.Index           =   1
         hWnd            =   7013636
         ScaleModeParent =   3
         Object.Left            =   8
         Object.Top             =   8
         Right           =   8
         Bottom          =   7
         MinWidth        =   2
         MinHeight       =   2
         LeftPercent     =   1.923
         TopPercent      =   2.857
         WidthPercent    =   96.154
         HeightPercent   =   94.643
         AnchorLeft      =   -1  'True
         AnchorTop       =   -1  'True
         AnchorRight     =   -1  'True
         AnchorBottom    =   -1  'True
      EndProperty
      BeginProperty Control_2 {881E793B-2C4F-4BDE-963C-25AE290F2EA6} 
         ParentTypeName  =   "SSTabEx"
         ParentName      =   "SSTabEx1"
         TypeName        =   "ucText"
         Name            =   "ucText1"
         Object.Index           =   2
         hWnd            =   1772620
         ScaleModeParent =   1
         Object.Left            =   240
         Object.Top             =   600
         Right           =   225
         Bottom          =   3000
         MinWidth        =   30
         MinHeight       =   30
         LeftPercent     =   4
         TopPercent      =   15.094
         WidthPercent    =   92.25
         HeightPercent   =   9.434
         AnchorLeft      =   -1  'True
         AnchorTop       =   -1  'True
         AnchorRight     =   -1  'True
      EndProperty
      BeginProperty Control_3 {881E793B-2C4F-4BDE-963C-25AE290F2EA6} 
         ParentTypeName  =   "SSTabEx"
         ParentName      =   "SSTabEx1"
         TypeName        =   "PictureBox"
         Name            =   "Picture1"
         Object.Index           =   3
         hWnd            =   1379206
         ScaleModeParent =   1
         Object.Left            =   -74760
         Object.Top             =   1560
         Right           =   75225
         Bottom          =   360
         MinWidth        =   15
         MinHeight       =   15
         LeftPercent     =   -1246
         TopPercent      =   39.245
         WidthPercent    =   92.25
         HeightPercent   =   51.698
         AnchorLeft      =   -1  'True
         AnchorTop       =   -1  'True
         AnchorRight     =   -1  'True
         AnchorBottom    =   -1  'True
      EndProperty
      BeginProperty Control_4 {881E793B-2C4F-4BDE-963C-25AE290F2EA6} 
         ParentTypeName  =   "SSTabEx"
         ParentName      =   "SSTabEx1"
         TypeName        =   "ucGridPlus"
         Name            =   "ucGridPlus1"
         Object.Index           =   4
         hWnd            =   6817052
         ScaleModeParent =   1
         Object.Left            =   -74640
         Object.Top             =   1440
         Right           =   75225
         Bottom          =   360
         MinWidth        =   30
         MinHeight       =   30
         LeftPercent     =   -1244
         TopPercent      =   36.226
         WidthPercent    =   90.25
         HeightPercent   =   54.717
         AnchorLeft      =   -1  'True
         AnchorTop       =   -1  'True
         AnchorRight     =   -1  'True
         AnchorBottom    =   -1  'True
      EndProperty
      BeginProperty Control_5 {881E793B-2C4F-4BDE-963C-25AE290F2EA6} 
         ParentTypeName  =   "SSTabEx"
         ParentName      =   "SSTabEx1"
         TypeName        =   "CommandButton"
         Name            =   "Command1"
         Object.Index           =   5
         hWnd            =   8258854
         ScaleModeParent =   1
         Object.Left            =   -74640
         Object.Top             =   600
         Right           =   75225
         Bottom          =   2640
         MinWidth        =   75
         MinHeight       =   195
         LeftPercent     =   -1244
         TopPercent      =   15.094
         WidthPercent    =   90.25
         HeightPercent   =   18.491
         AnchorLeft      =   -1  'True
         AnchorTop       =   -1  'True
         AnchorRight     =   -1  'True
      EndProperty
      BeginProperty Control_6 {881E793B-2C4F-4BDE-963C-25AE290F2EA6} 
         ParentTypeName  =   "SSTabEx"
         ParentName      =   "SSTabEx1"
         TypeName        =   "Label"
         Name            =   "Label1"
         Object.Index           =   6
         ScaleModeParent =   1
         Object.Left            =   240
         Object.Top             =   1080
         Right           =   3105
         Bottom          =   480
         MinWidth        =   15
         MinHeight       =   15
         UseLeftPercent  =   -1  'True
         LeftPercent     =   4
         TopPercent      =   27.17
         UseWidthPercent =   -1  'True
         WidthPercent    =   44.25
         HeightPercent   =   60.755
         AnchorTop       =   -1  'True
         AnchorBottom    =   -1  'True
      EndProperty
      BeginProperty Control_7 {881E793B-2C4F-4BDE-963C-25AE290F2EA6} 
         ParentTypeName  =   "SSTabEx"
         ParentName      =   "SSTabEx1"
         TypeName        =   "Label"
         Name            =   "Label2"
         Object.Index           =   7
         ScaleModeParent =   1
         Object.Left            =   3120
         Object.Top             =   1080
         Right           =   225
         Bottom          =   480
         MinWidth        =   15
         MinHeight       =   15
         UseLeftPercent  =   -1  'True
         LeftPercent     =   52
         TopPercent      =   27.17
         UseWidthPercent =   -1  'True
         WidthPercent    =   44.25
         HeightPercent   =   60.755
         AnchorTop       =   -1  'True
         AnchorBottom    =   -1  'True
      EndProperty
   End
   Begin Proyecto1.SSTabEx SSTabEx1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   7011
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tab             =   2
      TabHeight       =   563
      AutoTabHeight   =   -1  'True
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   2
      Tab(0).Control(0)=   "ucGridPlus1"
      Tab(0).Control(1)=   "Command1"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   1
      Tab(1).Control(0)=   "Picture1"
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   3
      Tab(2).Control(0)=   "ucText1"
      Tab(2).Control(1)=   "Label2"
      Tab(2).Control(2)=   "Label1"
      Begin Proyecto1.ucText ucText1 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         Text            =   "frmTest.frx":04C9
         ImgLeft         =   "frmTest.frx":04F7
         ImgRight        =   "frmTest.frx":050F
         RightButtonStyle=   0
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFC0&
         Height          =   2055
         Left            =   -74760
         ScaleHeight     =   1995
         ScaleWidth      =   5475
         TabIndex        =   3
         Top             =   1560
         Width           =   5535
      End
      Begin Proyecto1.ucGridPlus ucGridPlus1 
         Height          =   2175
         Left            =   -74640
         TabIndex        =   2
         Top             =   1440
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   3836
         HeaderHeight    =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   735
         Left            =   -74640
         TabIndex        =   1
         Top             =   600
         Width           =   5415
      End
      Begin VB.Label Label2 
         BackColor       =   &H000080FF&
         Caption         =   "Label1"
         Height          =   2415
         Left            =   3120
         TabIndex        =   7
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H008080FF&
         Caption         =   "Label1"
         Height          =   2415
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
