VERSION 5.00
Object = "{917DED51-9D95-4D2C-9431-281FAC2C6FF4}#1.0#0"; "JLAnchor.ocx"
Begin VB.Form frmChild 
   Caption         =   "Child"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4410
   ScaleWidth      =   9060
   Begin JLAnchor.ucJLAnchor ucJLAnchor1 
      Height          =   480
      Left            =   8400
      TabIndex        =   7
      Top             =   120
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      IconPresent     =   -1  'True
      FormIcon        =   "frmChild.frx":0000
      ControlsCount   =   12
      BeginProperty Control_1 {1820B3AA-C9AA-4184-8C03-110DA90D2ECB} 
         ParentTypeName  =   "Form"
         ParentName      =   "frmChild"
         TypeName        =   "Frame"
         Name            =   "Frame1"
         Object.Index           =   1
         hWnd            =   2637912
         Object.Left            =   4485
         Object.Top             =   840
         Right           =   90
         Bottom          =   90
         MinWidth        =   15
         MinHeight       =   15
         UseLeftPercent  =   -1  'True
         LeftPercent     =   49.503
         TopPercent      =   19.048
         UseWidthPercent =   -1  'True
         WidthPercent    =   49.503
         HeightPercent   =   78.912
         AnchorTop       =   -1  'True
         AnchorBottom    =   -1  'True
      EndProperty
      BeginProperty Control_2 {1820B3AA-C9AA-4184-8C03-110DA90D2ECB} 
         ParentTypeName  =   "Frame"
         ParentName      =   "Frame1"
         TypeName        =   "TextBox"
         Name            =   "Text3"
         Object.Index           =   2
         hWnd            =   2296264
         Object.Left            =   120
         Object.Top             =   1680
         Right           =   150
         Bottom          =   180
         MinWidth        =   150
         MinHeight       =   285
         UseModePercent  =   1
         UseLeftPercent  =   -1  'True
         LeftPercentStatic=   120
         UseTopPercent   =   -1  'True
         TopPercent      =   50
         TopPercentStatic=   -60
         UseWidthPercent =   -1  'True
         WidthPercent    =   100
         RightPercentStatic=   -150
         UseHeightPercent=   -1  'True
         HeightPercent   =   50
         BottomPercentStatic=   -180
      EndProperty
      BeginProperty Control_3 {1820B3AA-C9AA-4184-8C03-110DA90D2ECB} 
         ParentTypeName  =   "Frame"
         ParentName      =   "Frame1"
         TypeName        =   "TextBox"
         Name            =   "Text2"
         Object.Index           =   3
         hWnd            =   2953380
         Object.Left            =   120
         Object.Top             =   240
         Right           =   150
         Bottom          =   1860
         MinWidth        =   150
         MinHeight       =   285
         UseModePercent  =   1
         UseLeftPercent  =   -1  'True
         LeftPercentStatic=   120
         UseTopPercent   =   -1  'True
         TopPercentStatic=   240
         UseWidthPercent =   -1  'True
         WidthPercent    =   100
         RightPercentStatic=   -150
         UseHeightPercent=   -1  'True
         HeightPercent   =   50
         BottomPercentStatic=   -120
      EndProperty
      BeginProperty Control_4 {1820B3AA-C9AA-4184-8C03-110DA90D2ECB} 
         ParentTypeName  =   "Form"
         ParentName      =   "frmChild"
         TypeName        =   "PictureBox"
         Name            =   "Picture1"
         Object.Index           =   4
         hWnd            =   6623644
         Object.Top             =   840
         Right           =   4575
         Bottom          =   75
         MinWidth        =   15
         MinHeight       =   15
         UseLeftPercent  =   -1  'True
         TopPercent      =   19.048
         UseWidthPercent =   -1  'True
         WidthPercent    =   49.503
         HeightPercent   =   79.252
         AnchorTop       =   -1  'True
         AnchorBottom    =   -1  'True
      EndProperty
      BeginProperty Control_5 {1820B3AA-C9AA-4184-8C03-110DA90D2ECB} 
         ParentTypeName  =   "Form"
         ParentName      =   "frmChild"
         TypeName        =   "TextBox"
         Name            =   "Text1"
         Object.Index           =   5
         hWnd            =   7932554
         Object.Left            =   1440
         Object.Top             =   120
         Right           =   765
         Bottom          =   3795
         MinWidth        =   150
         MinHeight       =   285
         LeftPercent     =   15.894
         TopPercent      =   2.721
         WidthPercent    =   75.662
         HeightPercent   =   11.224
         AnchorLeft      =   -1  'True
         AnchorTop       =   -1  'True
         AnchorRight     =   -1  'True
      EndProperty
      BeginProperty Control_6 {1820B3AA-C9AA-4184-8C03-110DA90D2ECB} 
         ParentTypeName  =   "Form"
         ParentName      =   "frmChild"
         TypeName        =   "CommandButton"
         Name            =   "Command1"
         Object.Index           =   6
         hWnd            =   38340526
         Object.Left            =   120
         Object.Top             =   120
         Right           =   7725
         Bottom          =   3795
         MinWidth        =   75
         MinHeight       =   195
         LeftPercent     =   1.325
         TopPercent      =   2.721
         WidthPercent    =   13.411
         HeightPercent   =   11.224
         AnchorLeft      =   -1  'True
         AnchorTop       =   -1  'True
      EndProperty
      BeginProperty Control_7 {1820B3AA-C9AA-4184-8C03-110DA90D2ECB} 
         ParentTypeName  =   "PictureBox"
         ParentName      =   "Picture1"
         TypeName        =   "LabelPlus"
         Name            =   "LabelPlus1"
         Object.Index           =   7
         Object.Left            =   120
         Object.Top             =   120
         Right           =   2010
         Bottom          =   2820
         MinWidth        =   30
         MinHeight       =   30
         LeftPercent     =   2.712
         TopPercent      =   3.493
         WidthPercent    =   51.864
         HeightPercent   =   14.41
         AnchorLeft      =   -1  'True
         AnchorTop       =   -1  'True
         AnchorRight     =   -1  'True
      EndProperty
      BeginProperty Control_8 {1820B3AA-C9AA-4184-8C03-110DA90D2ECB} 
         TypeName        =   "Timer"
         Name            =   "Timer1"
         Object.Index           =   8
      EndProperty
      BeginProperty Control_9 {1820B3AA-C9AA-4184-8C03-110DA90D2ECB} 
         ParentTypeName  =   "PictureBox"
         ParentName      =   "Picture1"
         TypeName        =   "ucGridPlus"
         Name            =   "ucGridPlus1"
         Object.Index           =   9
         hWnd            =   987862
         Object.Left            =   120
         Object.Top             =   720
         Right           =   1050
         Bottom          =   780
         MinWidth        =   30
         MinHeight       =   30
         LeftPercent     =   2.712
         TopPercent      =   20.961
         WidthPercent    =   73.559
         HeightPercent   =   56.332
         AnchorLeft      =   -1  'True
         AnchorTop       =   -1  'True
         AnchorRight     =   -1  'True
         AnchorBottom    =   -1  'True
      EndProperty
      BeginProperty Control_10 {1820B3AA-C9AA-4184-8C03-110DA90D2ECB} 
         ParentTypeName  =   "ucGridPlus"
         ParentName      =   "ucGridPlus1"
         TypeName        =   "TextBox"
         Name            =   "Text4"
         Object.Index           =   10
         hWnd            =   3089264
         Object.Left            =   960
         Object.Top             =   600
         Right           =   1200
         Bottom          =   840
         MinWidth        =   150
         MinHeight       =   285
         LeftPercent     =   29.493
         TopPercent      =   31.008
         WidthPercent    =   33.641
         HeightPercent   =   25.581
         AnchorLeft      =   -1  'True
         AnchorTop       =   -1  'True
         AnchorRight     =   -1  'True
      EndProperty
      BeginProperty Control_11 {1820B3AA-C9AA-4184-8C03-110DA90D2ECB} 
         ParentTypeName  =   "PictureBox"
         ParentName      =   "Picture1"
         ParentIndex     =   "0"
         TypeName        =   "ucGridPlus"
         Name            =   "ucGridPlus1"
         ControlIndex    =   "0"
         Object.Index           =   11
         hWnd            =   3285872
         Object.Left            =   120
         Object.Top             =   720
         Right           =   1050
         Bottom          =   780
         MinWidth        =   30
         MinHeight       =   30
         LeftPercent     =   2.712
         TopPercent      =   20.961
         WidthPercent    =   73.559
         HeightPercent   =   56.332
         AnchorLeft      =   -1  'True
         AnchorTop       =   -1  'True
         AnchorRight     =   -1  'True
         AnchorBottom    =   -1  'True
      EndProperty
      BeginProperty Control_12 {1820B3AA-C9AA-4184-8C03-110DA90D2ECB} 
         ParentTypeName  =   "ucGridPlus"
         ParentName      =   "ucGridPlus1"
         ParentIndex     =   "0"
         TypeName        =   "ucText"
         Name            =   "ucText1"
         ControlIndex    =   "0"
         Object.Index           =   12
         hWnd            =   3345658
         Object.Left            =   120
         Object.Top             =   120
         Right           =   2040
         Bottom          =   1440
         MinWidth        =   30
         MinHeight       =   30
         LeftPercent     =   3.687
         TopPercent      =   6.202
         WidthPercent    =   33.641
         HeightPercent   =   19.38
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
      Begin VB.TextBox Text3 
         Height          =   1620
         Left            =   120
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox Text2 
         Height          =   1380
         Left            =   120
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   240
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
      Begin Proyecto1.ucGridPlus ucGridPlus1 
         Height          =   1935
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   3413
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
         Begin Proyecto1.ucText ucText1 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
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
            Text            =   "frmChild.frx":56C36
            ImgLeft         =   "frmChild.frx":56C64
            ImgRight        =   "frmChild.frx":56C7C
            RightButtonStyle=   0
         End
      End
      Begin Proyecto1.LabelPlus LabelPlus1 
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         BackColor       =   15296031
         BorderCornerLeftTop=   5
         BorderCornerRightTop=   5
         BorderCornerBottomRight=   5
         BorderCornerBottomLeft=   5
         CaptionAlignmentH=   1
         CaptionAlignmentV=   1
         Caption         =   "frmChild.frx":56C94
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
