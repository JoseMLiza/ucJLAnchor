VERSION 5.00
Object = "{917DED51-9D95-4D2C-9431-281FAC2C6FF4}#1.0#0"; "JLAnchor.ocx"
Begin VB.Form frmChild2 
   Caption         =   "fmrChild2 - Percent Static"
   ClientHeight    =   4305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4305
   ScaleWidth      =   6270
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00C0C0FF&
      Height          =   1095
      Left            =   3240
      ScaleHeight     =   1035
      ScaleWidth      =   2835
      TabIndex        =   6
      Top             =   3000
      Width           =   2895
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00C0FFC0&
      Height          =   1215
      Left            =   3240
      ScaleHeight     =   1155
      ScaleWidth      =   2835
      TabIndex        =   5
      Top             =   1560
      Width           =   2895
   End
   Begin JLAnchor.ucJLAnchor ucJLAnchor1 
      Height          =   480
      Left            =   5760
      TabIndex        =   2
      Top             =   0
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      IconPresent     =   -1  'True
      FormIcon        =   "frmChild2.frx":0000
      ControlsCount   =   6
      BeginProperty Control_1 {881E793B-2C4F-4BDE-963C-25AE290F2EA6} 
         ParentTypeName  =   "Form"
         ParentName      =   "frmChild2"
         TypeName        =   "PictureBox"
         Name            =   "Picture2"
         Object.Index           =   1
         hWnd            =   2428132
         Object.Left            =   120
         Object.Top             =   1560
         Right           =   3255
         Bottom          =   1530
         MinWidth        =   15
         MinHeight       =   15
         UseModePercent  =   1
         LeftPercent     =   1.887
         LeftPercentStatic=   2
         UseTopPercent   =   -1  'True
         TopPercent      =   33.333
         TopPercentStatic=   125
         UseWidthPercent =   -1  'True
         WidthPercent    =   50
         RightPercentStatic=   -120
         UseHeightPercent=   -1  'True
         HeightPercent   =   33.333
         BottomPercentStatic=   -95
         AnchorLeft      =   -1  'True
      EndProperty
      BeginProperty Control_2 {881E793B-2C4F-4BDE-963C-25AE290F2EA6} 
         ParentTypeName  =   "Form"
         ParentName      =   "frmChild2"
         TypeName        =   "PictureBox"
         Name            =   "Picture1"
         Object.Index           =   2
         hWnd            =   1773700
         Object.Left            =   120
         Object.Top             =   120
         Right           =   3255
         Bottom          =   2970
         MinWidth        =   15
         MinHeight       =   15
         UseModePercent  =   1
         LeftPercent     =   1.887
         LeftPercentStatic=   2
         UseTopPercent   =   -1  'True
         TopPercentStatic=   120
         UseWidthPercent =   -1  'True
         WidthPercent    =   50
         RightPercentStatic=   -120
         UseHeightPercent=   -1  'True
         HeightPercent   =   33.333
         BottomPercentStatic=   -100
         AnchorLeft      =   -1  'True
      EndProperty
      BeginProperty Control_3 {881E793B-2C4F-4BDE-963C-25AE290F2EA6} 
         ParentTypeName  =   "Form"
         ParentName      =   "frmChild2"
         TypeName        =   "PictureBox"
         Name            =   "Picture3"
         Object.Index           =   3
         hWnd            =   2623590
         Object.Left            =   120
         Object.Top             =   3000
         Right           =   3255
         Bottom          =   210
         MinWidth        =   15
         MinHeight       =   15
         UseModePercent  =   1
         LeftPercent     =   1.887
         LeftPercentStatic=   2
         UseTopPercent   =   -1  'True
         TopPercent      =   66.666
         TopPercentStatic=   130
         UseWidthPercent =   -1  'True
         WidthPercent    =   50
         RightPercentStatic=   -120
         UseHeightPercent=   -1  'True
         HeightPercent   =   33.333
         BottomPercentStatic=   -210
         AnchorLeft      =   -1  'True
      EndProperty
      BeginProperty Control_4 {881E793B-2C4F-4BDE-963C-25AE290F2EA6} 
         ParentTypeName  =   "Form"
         ParentName      =   "frmChild2"
         TypeName        =   "PictureBox"
         Name            =   "Picture4"
         Object.Index           =   4
         hWnd            =   7606934
         Object.Left            =   3240
         Object.Top             =   120
         Right           =   135
         Bottom          =   2970
         MinWidth        =   15
         MinHeight       =   15
         UseModePercent  =   1
         UseLeftPercent  =   -1  'True
         LeftPercent     =   50
         LeftPercentStatic=   105
         UseTopPercent   =   -1  'True
         TopPercentStatic=   120
         UseWidthPercent =   -1  'True
         WidthPercent    =   50
         RightPercentStatic=   -135
         UseHeightPercent=   -1  'True
         HeightPercent   =   33.333
         BottomPercentStatic=   -100
      EndProperty
      BeginProperty Control_5 {881E793B-2C4F-4BDE-963C-25AE290F2EA6} 
         ParentTypeName  =   "Form"
         ParentName      =   "frmChild2"
         TypeName        =   "PictureBox"
         Name            =   "Picture5"
         Object.Index           =   5
         hWnd            =   4721266
         Object.Left            =   3240
         Object.Top             =   1560
         Right           =   135
         Bottom          =   1530
         MinWidth        =   15
         MinHeight       =   15
         UseModePercent  =   1
         UseLeftPercent  =   -1  'True
         LeftPercent     =   50
         LeftPercentStatic=   105
         UseTopPercent   =   -1  'True
         TopPercent      =   33.333
         TopPercentStatic=   125
         UseWidthPercent =   -1  'True
         WidthPercent    =   50
         RightPercentStatic=   -135
         UseHeightPercent=   -1  'True
         HeightPercent   =   33.333
         BottomPercentStatic=   -95
      EndProperty
      BeginProperty Control_6 {881E793B-2C4F-4BDE-963C-25AE290F2EA6} 
         ParentTypeName  =   "Form"
         ParentName      =   "frmChild2"
         TypeName        =   "PictureBox"
         Name            =   "Picture6"
         Object.Index           =   6
         hWnd            =   4800488
         Object.Left            =   3240
         Object.Top             =   3000
         Right           =   135
         Bottom          =   210
         MinWidth        =   15
         MinHeight       =   15
         UseModePercent  =   1
         UseLeftPercent  =   -1  'True
         LeftPercent     =   50
         LeftPercentStatic=   105
         UseTopPercent   =   -1  'True
         TopPercent      =   66.666
         TopPercentStatic=   130
         UseWidthPercent =   -1  'True
         WidthPercent    =   50
         RightPercentStatic=   -135
         UseHeightPercent=   -1  'True
         HeightPercent   =   33.333
         BottomPercentStatic=   -210
      EndProperty
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00C0C0FF&
      Height          =   1215
      Left            =   3240
      ScaleHeight     =   1155
      ScaleWidth      =   2835
      TabIndex        =   4
      Top             =   120
      Width           =   2895
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0C0FF&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1035
      ScaleWidth      =   2835
      TabIndex        =   3
      Top             =   3000
      Width           =   2895
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0FFC0&
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   2835
      TabIndex        =   1
      Top             =   1560
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0FF&
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   2835
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmChild2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
