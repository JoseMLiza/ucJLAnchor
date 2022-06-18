VERSION 5.00
Object = "{62ED814B-A922-424D-833A-57C22F97CFE7}#3.5#0"; "JLAnchor.ocx"
Begin VB.Form frmNoChild 
   Caption         =   "Not Child"
   ClientHeight    =   3255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   3390
   Begin JLAnchor.ucJLAnchor ucJLAnchor1 
      Height          =   480
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      ControlsCount   =   4
      BeginProperty Control_1 {F01A6FF4-B13C-432C-BCBF-2B9C8BFB03DE} 
         ParentTypeName  =   "Form"
         ParentName      =   "frmNoChild"
         TypeName        =   "CommandButton"
         Name            =   "Command1"
         Object.Index           =   1
         hWnd            =   264470
         Object.Left            =   120
         Object.Top             =   120
         Right           =   2055
         Bottom          =   2640
         LeftPercent     =   3.54
         TopPercent      =   3.687
         WidthPercent    =   35.841
         HeightPercent   =   15.207
         AnchorLeft      =   -1  'True
         AnchorTop       =   -1  'True
      EndProperty
      BeginProperty Control_2 {F01A6FF4-B13C-432C-BCBF-2B9C8BFB03DE} 
         ParentTypeName  =   "Form"
         ParentName      =   "frmNoChild"
         TypeName        =   "TextBox"
         Name            =   "Text1"
         Object.Index           =   2
         hWnd            =   264474
         Object.Left            =   1440
         Object.Top             =   120
         Right           =   735
         Bottom          =   2640
         LeftPercent     =   42.478
         TopPercent      =   3.687
         WidthPercent    =   35.841
         HeightPercent   =   15.207
         AnchorLeft      =   -1  'True
         AnchorTop       =   -1  'True
         AnchorRight     =   -1  'True
      EndProperty
      BeginProperty Control_3 {F01A6FF4-B13C-432C-BCBF-2B9C8BFB03DE} 
         ParentTypeName  =   "Form"
         ParentName      =   "frmNoChild"
         TypeName        =   "PictureBox"
         Name            =   "Picture1"
         Object.Index           =   3
         hWnd            =   329958
         Object.Left            =   120
         Object.Top             =   840
         Right           =   135
         Bottom          =   120
         LeftPercent     =   3.54
         TopPercent      =   25.806
         WidthPercent    =   92.478
         HeightPercent   =   70.507
         AnchorLeft      =   -1  'True
         AnchorTop       =   -1  'True
         AnchorRight     =   -1  'True
         AnchorBottom    =   -1  'True
      EndProperty
      BeginProperty Control_4 {F01A6FF4-B13C-432C-BCBF-2B9C8BFB03DE} 
         ParentTypeName  =   "PictureBox"
         ParentName      =   "Picture1"
         TypeName        =   "TextBox"
         Name            =   "Text2"
         Object.Index           =   4
         hWnd            =   526604
         Object.Left            =   240
         Object.Top             =   240
         Right           =   540
         Bottom          =   1620
         LeftPercent     =   7.805
         TopPercent      =   10.738
         WidthPercent    =   74.634
         HeightPercent   =   16.779
         AnchorLeft      =   -1  'True
         AnchorTop       =   -1  'True
         AnchorRight     =   -1  'True
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2235
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   840
      Width           =   3135
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Text            =   "Text2"
         Top             =   240
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmNoChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
