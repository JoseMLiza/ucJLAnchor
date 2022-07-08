VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{917DED51-9D95-4D2C-9431-281FAC2C6FF4}#1.0#0"; "JLAnchor.ocx"
Begin VB.Form frmNoChild 
   Caption         =   "No Child"
   ClientHeight    =   3255
   ClientLeft      =   120
   ClientTop       =   765
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
      IconPresent     =   -1  'True
      FormIcon        =   "frmNoChild.frx":0000
      ControlsCount   =   5
      BeginProperty Control_1 {1820B3AA-C9AA-4184-8C03-110DA90D2ECB} 
         ParentTypeName  =   "Form"
         ParentName      =   "frmNoChild"
         TypeName        =   "CommandButton"
         Name            =   "Command1"
         Object.Index           =   1
         hWnd            =   7868802
         Object.Left            =   120
         Object.Top             =   120
         Right           =   2055
         Bottom          =   2640
         MinWidth        =   75
         MinHeight       =   195
         LeftPercent     =   3.54
         TopPercent      =   3.687
         WidthPercent    =   35.841
         HeightPercent   =   15.207
         AnchorLeft      =   -1  'True
         AnchorTop       =   -1  'True
      EndProperty
      BeginProperty Control_2 {1820B3AA-C9AA-4184-8C03-110DA90D2ECB} 
         ParentTypeName  =   "Form"
         ParentName      =   "frmNoChild"
         TypeName        =   "TextBox"
         Name            =   "Text1"
         Object.Index           =   2
         hWnd            =   4786524
         Object.Left            =   1440
         Object.Top             =   120
         Right           =   735
         Bottom          =   2640
         MinWidth        =   150
         MinHeight       =   285
         LeftPercent     =   42.478
         TopPercent      =   3.687
         WidthPercent    =   35.841
         HeightPercent   =   15.207
         AnchorLeft      =   -1  'True
         AnchorTop       =   -1  'True
         AnchorRight     =   -1  'True
      EndProperty
      BeginProperty Control_3 {1820B3AA-C9AA-4184-8C03-110DA90D2ECB} 
         ParentTypeName  =   "Form"
         ParentName      =   "frmNoChild"
         TypeName        =   "PictureBox"
         Name            =   "Picture1"
         Object.Index           =   3
         hWnd            =   9505864
         Object.Left            =   120
         Object.Top             =   840
         Right           =   135
         Bottom          =   120
         MinWidth        =   15
         MinHeight       =   15
         LeftPercent     =   3.54
         TopPercent      =   25.806
         WidthPercent    =   92.478
         HeightPercent   =   70.507
         AnchorLeft      =   -1  'True
         AnchorTop       =   -1  'True
         AnchorRight     =   -1  'True
         AnchorBottom    =   -1  'True
      EndProperty
      BeginProperty Control_4 {1820B3AA-C9AA-4184-8C03-110DA90D2ECB} 
         ParentTypeName  =   "PictureBox"
         ParentName      =   "Picture1"
         TypeName        =   "TextBox"
         Name            =   "Text2"
         Object.Index           =   4
         hWnd            =   3148000
         Object.Left            =   240
         Object.Top             =   240
         Right           =   540
         Bottom          =   1620
         MinWidth        =   150
         MinHeight       =   285
         LeftPercent     =   7.805
         TopPercent      =   10.738
         WidthPercent    =   74.634
         HeightPercent   =   16.779
         AnchorLeft      =   -1  'True
         AnchorTop       =   -1  'True
         AnchorRight     =   -1  'True
      EndProperty
      BeginProperty Control_5 {1820B3AA-C9AA-4184-8C03-110DA90D2ECB} 
         ParentTypeName  =   "PictureBox"
         ParentName      =   "Picture1"
         TypeName        =   "ListView"
         Name            =   "ListView1"
         Object.Index           =   5
         Object.Left            =   240
         Object.Top             =   720
         Right           =   540
         Bottom          =   180
         MinWidth        =   30
         MinHeight       =   30
         LeftPercent     =   7.805
         TopPercent      =   32.215
         WidthPercent    =   74.634
         HeightPercent   =   59.732
         AnchorLeft      =   -1  'True
         AnchorTop       =   -1  'True
         AnchorRight     =   -1  'True
         AnchorBottom    =   -1  'True
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   1335
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   2355
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Text            =   "Text2"
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Menu mnPrueba 
      Caption         =   "Menu"
   End
End
Attribute VB_Name = "frmNoChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
