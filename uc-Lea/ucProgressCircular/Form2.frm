VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ucProgressCircular"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6630
   LinkTopic       =   "Form2"
   ScaleHeight     =   8415
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   1320
      Max             =   100
      TabIndex        =   24
      Top             =   8040
      Width           =   3975
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1335
      Index           =   23
      Left            =   5040
      TabIndex        =   23
      Top             =   6480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2355
      Caption1_ForeColor=   16724889
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   0
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PB_Color1       =   13369344
      PB_Steps        =   64
      Value           =   60
      AnimationInterval=   100
      PF_ColorsCount  =   3
      PF_Colors       =   "Form2.frx":0000
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1305
      Index           =   19
      Left            =   3720
      TabIndex        =   19
      Top             =   4920
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   2302
      Caption1        =   " "
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   0
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PF_Width        =   5
      PB_Color1       =   15527149
      PB_Width        =   5
      Value           =   60
      StartAngle      =   180
      GradientAngle   =   0
      RoundEndStyle   =   -1  'True
      PF_ForeColor    =   3355596
      AnimationInterval=   100
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1305
      Index           =   18
      Left            =   3600
      TabIndex        =   18
      Top             =   4920
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   2302
      Caption1        =   " "
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   0
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PF_Width        =   5
      PB_Color1       =   15527149
      PB_Width        =   5
      Value           =   40
      StartAngle      =   180
      GradientAngle   =   0
      RoundEndStyle   =   -1  'True
      PF_ForeColor    =   52224
      AnimationInterval=   100
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1305
      Index           =   17
      Left            =   3480
      TabIndex        =   17
      Top             =   4920
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   2302
      Caption1        =   " "
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   0
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PF_Width        =   5
      PB_Color1       =   15527149
      PB_Width        =   5
      StartAngle      =   180
      GradientAngle   =   0
      RoundEndStyle   =   -1  'True
      PF_ForeColor    =   13382400
      AnimationInterval=   100
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1215
      Index           =   16
      Left            =   1920
      TabIndex        =   16
      Top             =   6480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2143
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   0
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PF_Width        =   5
      PB_ColorGradient=   -1  'True
      PB_Width        =   5
      Value           =   100
      AnimationInterval=   100
      PF_ColorsCount  =   7
      PF_Colors       =   "Form2.frx":002D
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1215
      Index           =   15
      Left            =   3720
      TabIndex        =   15
      Top             =   3360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2143
      Caption1        =   " "
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   0
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PB_Color1Opacity=   0
      Value           =   75
      RoundEndStyle   =   -1  'True
      ShowAnimation   =   -1  'True
      AnimationInterval=   50
      PF_ColorsCount  =   2
      PF_Colors       =   "Form2.frx":006A
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1575
      Index           =   14
      Left            =   5040
      TabIndex        =   14
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2778
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   -5
      Caption2        =   "Danger"
      Caption2_ForeColor=   153
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2_OffsetY=   12
      PB_Color1       =   15527149
      PB_BorderWidth  =   1
      PB_BorderColorOpacity=   0
      Max             =   360
      Value           =   360
      Angle           =   270
      StartAngle      =   225
      CenterColor1    =   13421772
      CenterColor1Opacity=   50
      RoundStartStyle =   -1  'True
      RoundEndStyle   =   -1  'True
      DisplayInPercent=   0   'False
      AnimationInterval=   100
      PF_ColorsCount  =   5
      PF_Colors       =   "Form2.frx":0093
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1575
      Index           =   13
      Left            =   1920
      TabIndex        =   13
      Top             =   4920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2778
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   -5
      Caption2        =   "Danger"
      Caption2_ForeColor=   153
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2_OffsetY=   12
      PB_Color1       =   15527149
      PB_Width        =   7
      PB_Border       =   -1  'True
      PB_BorderWidth  =   10
      PB_BorderColorOpacity=   0
      Max             =   360
      Value           =   351
      CenterColor1    =   13421772
      CenterColor1Opacity=   50
      CenterVisible   =   -1  'True
      DisplayInPercent=   0   'False
      AnimationInterval=   100
      PF_ColorsCount  =   4
      PF_Colors       =   "Form2.frx":00C8
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1335
      Index           =   11
      Left            =   240
      TabIndex        =   12
      Top             =   6360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2355
      Caption1        =   "Loading..."
      Caption1_ForeColor=   16777215
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   0
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StepSpaceSize   =   2
      PF_Steps        =   64
      PB_Color1       =   0
      PB_Color2       =   10066329
      PB_ColorGradient=   -1  'True
      PB_Border       =   -1  'True
      PB_BorderColor  =   3355443
      PB_BorderWidth  =   3
      Value           =   45
      CenterGradient  =   -1  'True
      GradientAngle   =   0
      CenterColor1    =   10066329
      CenterColor2    =   0
      CenterVisible   =   -1  'True
      PF_ForeColor    =   16777215
      PF_ForeColorOpacity=   80
      AnimationInterval=   50
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1335
      Index           =   10
      Left            =   240
      TabIndex        =   11
      Top             =   4920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2355
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   0
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PB_ColorGradient=   -1  'True
      Value           =   10
      CenterGradient  =   -1  'True
      CenterVisible   =   -1  'True
      PF_ForeColor    =   16711884
      AnimationInterval=   100
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1695
      Index           =   12
      Left            =   4920
      TabIndex        =   10
      Top             =   4680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2990
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   0
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PF_Width        =   7
      PF_Steps        =   8
      PB_Color1       =   16777215
      PB_Width        =   7
      PB_Steps        =   8
      PB_Border       =   -1  'True
      PB_BorderColor  =   3368703
      PF_ForeColor    =   39423
      AnimationInterval=   100
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1335
      Index           =   9
      Left            =   4920
      TabIndex        =   9
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2355
      Caption1_ForeColor=   16777215
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   0
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PF_Width        =   7
      PB_Color1       =   13080638
      PB_Border       =   -1  'True
      PB_BorderColor  =   13080638
      PB_BorderWidth  =   7
      Value           =   66
      CenterColor1    =   13080638
      CenterVisible   =   -1  'True
      RoundStartStyle =   -1  'True
      RoundEndStyle   =   -1  'True
      PF_ForeColor    =   16777215
      AnimationInterval=   100
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1695
      Index           =   8
      Left            =   1800
      TabIndex        =   8
      Top             =   3240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   2990
      Caption1        =   "time remaining"
      Caption1_ForeColor=   10066329
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   -10
      Caption2        =   "07:14"
      Caption2_ForeColor=   10210816
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Light"
         Size            =   15.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2_OffsetY=   10
      PF_Width        =   7
      PB_Color1       =   4143675
      PB_Width        =   7
      Value           =   70
      CenterColor1    =   4143675
      CenterVisible   =   -1  'True
      PF_ForeColor    =   10210816
      AnimationInterval=   100
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1455
      Index           =   7
      Left            =   3360
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2566
      Caption1_ForeColor=   16777215
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Impact"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   8
      Caption2        =   "EXCEL"
      Caption2_ForeColor=   16777215
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2_OffsetY=   -8
      PB_Color1       =   13421772
      PB_Border       =   -1  'True
      PB_BorderColor  =   39168
      Value           =   25
      CenterColor1    =   26112
      CenterVisible   =   -1  'True
      PF_ForeColor    =   39168
      AnimationInterval=   100
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1335
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2355
      Caption1_ForeColor=   11565097
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   0
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StepSpaceSize   =   3
      PF_Steps        =   20
      PB_Color1       =   11565097
      PB_Color1Opacity=   10
      PB_Steps        =   20
      Value           =   70
      PF_ForeColor    =   11565097
      AnimationInterval=   100
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1455
      Index           =   5
      Left            =   1800
      TabIndex        =   5
      Top             =   1680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2566
      Caption1_ForeColor=   16737792
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   0
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PB_Color1       =   10066329
      PB_Width        =   5
      Value           =   65
      RoundStartStyle =   -1  'True
      RoundEndStyle   =   -1  'True
      PF_ForeColor    =   16737792
      AnimationInterval=   100
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1455
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2566
      Caption1_ForeColor=   6710886
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   0
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PB_Color1       =   255
      PB_Color1Opacity=   10
      Value           =   25
      RoundStartStyle =   -1  'True
      RoundEndStyle   =   -1  'True
      PF_ForeColor    =   255
      AnimationInterval=   100
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1455
      Index           =   3
      Left            =   4920
      TabIndex        =   3
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2566
      Caption1_ForeColor=   10156473
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   0
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StepSpaceSize   =   2
      PF_Width        =   5
      PF_Steps        =   100
      PB_Color1       =   15527149
      PB_Width        =   5
      PB_Steps        =   100
      Value           =   60
      PF_ForeColor    =   10156473
      AnimationInterval=   100
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1455
      Index           =   2
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2566
      Caption1_ForeColor=   16777215
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   0
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PB_Color1Opacity=   0
      PB_Border       =   -1  'True
      PB_BorderWidth  =   10
      PB_BorderColorOpacity=   0
      Value           =   40
      CenterColor1    =   6645093
      CenterVisible   =   -1  'True
      PF_ForeColor    =   16737792
      AnimationInterval=   100
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1335
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2355
      Caption1_ForeColor=   13892210
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   0
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StepSpaceSize   =   2
      PF_Width        =   5
      PF_Steps        =   36
      PB_Color1       =   15527149
      PB_Width        =   5
      Value           =   60
      PF_ForeColor    =   11682539
      AnimationInterval=   100
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1455
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2566
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Light"
         Size            =   21.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   0
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PF_Width        =   5
      PB_Width        =   5
      PF_ForeColor    =   52224
      AnimationInterval=   100
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1335
      Index           =   22
      Left            =   3360
      TabIndex        =   20
      Top             =   6480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2355
      Caption1_ForeColor=   16777215
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   -10
      Caption2        =   "Loading"
      Caption2_ForeColor=   16777215
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Algerian"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2_OffsetY=   10
      PF_Width        =   5
      PB_Color1       =   4013373
      PB_Color2       =   18318
      PB_Width        =   5
      Value           =   20
      CenterGradient  =   -1  'True
      CenterColor1    =   27605
      CenterColor2    =   18061
      CenterVisible   =   -1  'True
      PF_ForeColor    =   12648447
      AnimationInterval=   100
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1575
      Index           =   21
      Left            =   3340
      TabIndex        =   22
      Top             =   6330
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2778
      Caption1        =   " "
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   0
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PB_Color1       =   16777215
      PB_Color2       =   6710886
      PB_ColorGradient=   -1  'True
      PB_Border       =   -1  'True
      PB_BorderColor  =   10066329
      PB_BorderWidth  =   1
      Value           =   0
      Angle           =   45
      StartAngle      =   120
      GradientAngle   =   225
      AnimationInterval=   100
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1455
      Index           =   20
      Left            =   3240
      TabIndex        =   21
      Top             =   6405
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   2566
      Caption1        =   " "
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   0
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PB_Color1       =   16777215
      PB_Color2       =   6710886
      PB_ColorGradient=   -1  'True
      PB_Border       =   -1  'True
      PB_BorderColor  =   10066329
      PB_BorderWidth  =   1
      Value           =   0
      Angle           =   120
      StartAngle      =   270
      AnimationInterval=   100
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub HScroll1_Change()
    Dim i As Long
    
    For i = 0 To ucProgressCircular1.Count - 1
        With ucProgressCircular1(i)
            .Value = (.Max - .Min) * HScroll1.Value / 100
        End With
    Next
End Sub

Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub

Private Sub ucProgressCircular1_Click(Index As Integer)
    ucProgressCircular1(Index).ShowAnimation = Not ucProgressCircular1(Index).ShowAnimation
End Sub
