VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FDFDFD&
   Caption         =   "Print"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4200
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H00E6C29B&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   0
      ScaleHeight     =   6855
      ScaleWidth      =   1815
      TabIndex        =   1
      Top             =   0
      Width           =   1815
      Begin VB.CheckBox ChkSeparatorLine 
         BackColor       =   &H00E6C29B&
         Caption         =   "Separator Line"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   3360
         Width           =   1575
      End
      Begin VB.CommandButton CmdAnimate 
         Caption         =   "Animate 360°"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   3720
         Width           =   1575
      End
      Begin VB.CheckBox ChkLegendVisible 
         BackColor       =   &H00E6C29B&
         Caption         =   "Legend Visible"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Text            =   "{P}%"
         Top             =   2280
         Width           =   1575
      End
      Begin VB.ComboBox CboStyle 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   120
         List            =   "Form1.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ComboBox CboLegendAlign 
         Height          =   315
         ItemData        =   "Form1.frx":0020
         Left            =   120
         List            =   "Form1.frx":0030
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox cboLabelsPositions 
         Height          =   315
         ItemData        =   "Form1.frx":0066
         Left            =   120
         List            =   "Form1.frx":0073
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3000
         Width           =   1575
      End
      Begin VB.CheckBox ChkLabelsVisible 
         BackColor       =   &H00E6C29B&
         Caption         =   "Labels Visible"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2640
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "(Drag the Graphic)"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Labels Format, {V} = Value, {P}=Percent"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Chart Style"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1545
      End
   End
   Begin Proyecto1.ucPieChart ucPieChart1 
      Height          =   3255
      Left            =   3360
      TabIndex        =   0
      Top             =   960
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FillGradient    =   -1  'True
      LabelsVisible   =   -1  'True
      ChartStyle      =   1
      LegendAlign     =   1
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleForeColor  =   0
      SeparatorLine   =   0   'False
      SeparatorLineWidth=   3
      SeparatorLineColor=   15921906
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MemAngle As Single
  
Private Sub CmdPrint_Click()
    Dim oPic As StdPicture
    Printer.ScaleMode = vbPixels
    Set oPic = ucPieChart1.Image(Printer.ScaleWidth / 10, Printer.ScaleHeight / 10 / 2)
    Printer.PaintPicture oPic, 0, 0, Printer.ScaleWidth, Printer.ScaleHeight / 2
    Printer.EndDoc
End Sub

Private Sub Form_Load()
    Dim Palette() As String
    
    Me.cboLabelsPositions.ListIndex = 0
    Me.CboLegendAlign.ListIndex = 0
    Me.CboStyle.ListIndex = 0
    
    Palette = Split("&H004744E3,&H003DB0EF,&H00ABA56C,&H0048BDBF,&H004D91F4,&H007450,&H0050C187", ",")

    Dim i As Long
    Randomize Timer
    For i = 0 To 6
        ucPieChart1.AddItem "2000" + i, Rnd, CLng(Palette(i)), i = 3
    Next
    
End Sub

Private Function Random(Min!, Max!)
    Random = Int((Max - Min + 1) * Rnd + Min)
End Function

Private Sub cboLabelsPositions_Click()
    ucPieChart1.LabelsPositions = cboLabelsPositions.ListIndex
End Sub

Private Sub CboLegendAlign_Click()
    ucPieChart1.LegendAlign = CboLegendAlign.ListIndex
End Sub

Private Sub ChkLabelsVisible_Click()
    ucPieChart1.LabelsVisible = ChkLabelsVisible.Value
End Sub

Private Sub ChkLegendVisible_Click()
    ucPieChart1.LegendVisible = ChkLegendVisible.Value
End Sub

Private Sub CboStyle_Click()
    ucPieChart1.ChartStyle = CboStyle.ListIndex
End Sub

Private Sub ChkSeparatorLine_Click()
    ucPieChart1.SeparatorLine = ChkSeparatorLine.Value
End Sub

Private Sub CmdAnimate_Click()
    Dim i As Long
    For i = ucPieChart1.Rotation To ucPieChart1.Rotation + 360
        ucPieChart1.Rotation = i
        DoEvents
    Next
End Sub

Private Sub Form_Resize()
    ucPieChart1.Move Picture1.Width, 0, Me.ScaleWidth - Picture1.Width, Me.ScaleHeight
End Sub

Private Sub Text1_Change()
    ucPieChart1.LabelsFormats = Text1.text
End Sub

Private Sub ucPieChart1_ItemClick(Index As Long)
    ucPieChart1.Special(Index) = Not ucPieChart1.Special(Index)
End Sub

Private Sub ucPieChart1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ret As Single, A As Single
    Dim XX As Single, YY As Single
    
    If Button <> vbLeftButton Then Exit Sub
    
    ucPieChart1.GetCenterPie XX, YY
    ret = MATH_GetAngle(XX, Y, X, YY)
    
    If ret <= 0 Then
        A = Round(180 + (180 - (-ret)))
    Else
        A = Round(ret)
    End If

    MemAngle = A - ucPieChart1.Rotation
End Sub

Private Sub ucPieChart1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ret As Single, A As Single
    Dim XX As Single, YY As Single
    
    If Button <> vbLeftButton Then Exit Sub
    
    ucPieChart1.GetCenterPie XX, YY
    ret = MATH_GetAngle(XX, Y, X, YY)
    
    If ret <= 0 Then
        A = Round(180 + (180 - (-ret)))
    Else
        A = Round(ret)
    End If
     ucPieChart1.Rotation = A - MemAngle
End Sub
