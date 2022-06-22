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
      TabIndex        =   4
      Top             =   120
      Width           =   495
      _ExtentX        =   847
      _ExtentY        =   847
      ControlsCount   =   8
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
         ParentTypeName  =   "Form"
         ParentName      =   "frmChild"
         TypeName        =   "PictureBox"
         Name            =   "Picture1"
         Object.Index           =   2
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
      BeginProperty Control_3 {F01A6FF4-B13C-432C-BCBF-2B9C8BFB03DE} 
         ParentTypeName  =   "Form"
         ParentName      =   "frmChild"
         TypeName        =   "TextBox"
         Name            =   "Text1"
         Object.Index           =   3
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
      BeginProperty Control_4 {F01A6FF4-B13C-432C-BCBF-2B9C8BFB03DE} 
         ParentTypeName  =   "Form"
         ParentName      =   "frmChild"
         TypeName        =   "CommandButton"
         Name            =   "Command1"
         Object.Index           =   4
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
      BeginProperty Control_5 {F01A6FF4-B13C-432C-BCBF-2B9C8BFB03DE} 
         ParentTypeName  =   "Frame"
         ParentName      =   "Frame1"
         TypeName        =   "TextBox"
         Name            =   "Text2"
         Object.Index           =   5
         hWnd            =   3868374
         Object.Left            =   120
         Object.Top             =   360
         Right           =   150
         Bottom          =   1785
         LeftPercent     =   2.676
         UseTopPercent   =   -1  'True
         TopPercent      =   10.345
         WidthPercent    =   93.98
         UseHeightPercent=   -1  'True
         HeightPercent   =   38.362
         AnchorLeft      =   -1  'True
         AnchorRight     =   -1  'True
      EndProperty
      BeginProperty Control_6 {F01A6FF4-B13C-432C-BCBF-2B9C8BFB03DE} 
         ParentTypeName  =   "Frame"
         ParentName      =   "Frame1"
         TypeName        =   "TextBox"
         Name            =   "Text3"
         Object.Index           =   6
         hWnd            =   4734948
         Object.Left            =   120
         Object.Top             =   1800
         Right           =   150
         Bottom          =   105
         LeftPercent     =   2.676
         UseTopPercent   =   -1  'True
         TopPercent      =   51.724
         WidthPercent    =   93.98
         UseHeightPercent=   -1  'True
         HeightPercent   =   45.259
         AnchorLeft      =   -1  'True
         AnchorRight     =   -1  'True
      EndProperty
      BeginProperty Control_7 {F01A6FF4-B13C-432C-BCBF-2B9C8BFB03DE} 
         ParentTypeName  =   "PictureBox"
         ParentName      =   "Picture1"
         TypeName        =   "TextBox"
         Name            =   "Text4"
         Object.Index           =   7
         hWnd            =   463752
         Bottom          =   1717
         UseLeftPercent  =   -1  'True
         UseTopPercent   =   -1  'True
         UseWidthPercent =   -1  'True
         WidthPercent    =   100
         UseHeightPercent=   -1  'True
         HeightPercent   =   50.015
      EndProperty
      BeginProperty Control_8 {F01A6FF4-B13C-432C-BCBF-2B9C8BFB03DE} 
         ParentTypeName  =   "PictureBox"
         ParentName      =   "Picture1"
         TypeName        =   "TextBox"
         Name            =   "Text5"
         Object.Index           =   8
         hWnd            =   529218
         Object.Top             =   1718
         Bottom          =   -1
         UseLeftPercent  =   -1  'True
         UseTopPercent   =   -1  'True
         TopPercent      =   50.015
         UseWidthPercent =   -1  'True
         WidthPercent    =   100
         UseHeightPercent=   -1  'True
         HeightPercent   =   50.015
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3480
      Left            =   4485
      TabIndex        =   3
      Top             =   840
      Width           =   4485
      Begin VB.TextBox Text3 
         Height          =   1575
         Left            =   120
         TabIndex        =   6
         Text            =   "Text3"
         Top             =   1800
         Width           =   4215
      End
      Begin VB.TextBox Text2 
         Height          =   1335
         Left            =   120
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   360
         Width           =   4215
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
      Begin VB.TextBox Text5 
         Height          =   1718
         Left            =   0
         TabIndex        =   8
         Text            =   "Text5"
         Top             =   1718
         Width           =   4425
      End
      Begin VB.TextBox Text4 
         Height          =   1718
         Left            =   0
         TabIndex        =   7
         Text            =   "Text4"
         Top             =   0
         Width           =   4425
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
