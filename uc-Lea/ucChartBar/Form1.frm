VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F2F2F2&
   Caption         =   "ucChartBar1"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   10185
   StartUpPosition =   3  'Windows Default
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
      Height          =   7545
      Left            =   0
      ScaleHeight     =   7545
      ScaleWidth      =   1815
      TabIndex        =   1
      Top             =   0
      Width           =   1815
      Begin VB.CommandButton Command3 
         Caption         =   "Colors"
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
         TabIndex        =   23
         Top             =   7080
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
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
         TabIndex        =   22
         Top             =   6600
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
         TabIndex        =   21
         Top             =   5040
         Width           =   1455
      End
      Begin VB.ComboBox cboLabelsPositions 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   120
         List            =   "Form1.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   5280
         Width           =   1575
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   120
         Max             =   359
         TabIndex        =   19
         Top             =   5760
         Width           =   1575
      End
      Begin VB.CheckBox ChkHorizontalLines 
         BackColor       =   &H00E6C29B&
         Caption         =   "Horizontal Lines"
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
         TabIndex        =   18
         Top             =   2400
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox ChkVerticalLines 
         BackColor       =   &H00E6C29B&
         Caption         =   "Vertical Lines"
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
         TabIndex        =   17
         Top             =   2160
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox ChkAxisY 
         BackColor       =   &H00E6C29B&
         Caption         =   "AxisY Visible"
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
         TabIndex        =   16
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.ComboBox CboLegendAlign 
         Height          =   315
         ItemData        =   "Form1.frx":0031
         Left            =   120
         List            =   "Form1.frx":0041
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   840
         Width           =   1575
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E6C29B&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   1575
         TabIndex        =   12
         Top             =   1800
         Width           =   1575
         Begin VB.OptionButton Option2 
            BackColor       =   &H00E6C29B&
            Caption         =   "Horiz."
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
            Index           =   1
            Left            =   600
            TabIndex        =   14
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00E6C29B&
            Caption         =   "Vert."
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
            Index           =   0
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Form1.frx":0077
         Left            =   120
         List            =   "Form1.frx":0084
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3240
         Width           =   1575
      End
      Begin VB.ComboBox CboNroSeries 
         Height          =   315
         ItemData        =   "Form1.frx":00BB
         Left            =   120
         List            =   "Form1.frx":00CE
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   4680
         Width           =   1575
      End
      Begin VB.CheckBox ChkAxisX 
         BackColor       =   &H00E6C29B&
         Caption         =   "AxisX visible"
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
         TabIndex        =   6
         Top             =   1200
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
         TabIndex        =   5
         Text            =   "${V}"
         Top             =   4080
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
         TabIndex        =   4
         Top             =   480
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox ChkNegative 
         BackColor       =   &H00E6C29B&
         Caption         =   "Negatives Value"
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
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Random"
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
         TabIndex        =   2
         Top             =   6120
         Width           =   1575
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
         TabIndex        =   11
         Top             =   3000
         Width           =   1545
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Series"
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
         Top             =   4440
         Width           =   1320
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Labels Format, {V} = Value"
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
         TabIndex        =   8
         Top             =   3600
         Width           =   1215
      End
   End
   Begin Proyecto1.ucChartBar ucChartBar1 
      CausesValidation=   0   'False
      Height          =   3975
      Left            =   2040
      TabIndex        =   0
      Top             =   720
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7011
      ForeColor       =   5855577
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FillOpacity     =   100
      VerticalLines   =   -1  'True
      ChartStyle      =   2
      LegendAlign     =   0
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
      LabelsPositions =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ChkLabelsVisible_Click()
    ucChartBar1.LabelsVisible = ChkLabelsVisible.Value
End Sub

Private Sub ChkNegative_Click()
    ucChartBar1.Clear
    Form_Load
    ucChartBar1.Refresh
End Sub

Private Sub ChkLegendVisible_Click()
    ucChartBar1.LegendVisible = ChkLegendVisible.Value
End Sub

Private Sub CboLegendAlign_Click()
    ucChartBar1.LegendAlign = CboLegendAlign.ListIndex
End Sub

Private Sub ChkAxisX_Click()
    ucChartBar1.AxisXVisible = ChkAxisX.Value
End Sub

Private Sub ChkAxisy_Click()
    ucChartBar1.AxisYVisible = ChkAxisY.Value
End Sub

Private Sub CboLabelsPositions_Click()
    ucChartBar1.LabelsPositions = cboLabelsPositions.ListIndex
End Sub

Private Sub Command2_Click()
Dim oPic As StdPicture
    Printer.ScaleMode = vbPixels
    Set oPic = ucChartBar1.Image(Printer.ScaleWidth / 10, Printer.ScaleHeight / 10 / 2)
    Printer.PaintPicture oPic, 0, 0, Printer.ScaleWidth, Printer.ScaleHeight / 2
    Printer.EndDoc
End Sub

Private Sub Command3_Click()
    Form2.Show
End Sub

Private Sub Option2_Click(Index As Integer)
    ucChartBar1.ChartOrientation = Index
End Sub

Private Sub ChkHorizontalLines_Click()
    ucChartBar1.HorizontalLines = ChkHorizontalLines.Value
End Sub

Private Sub ChkVerticalLines_Click()
    ucChartBar1.VerticalLines = ChkVerticalLines.Value
End Sub

Private Sub Combo2_Click()
    ucChartBar1.ChartStyle = Combo2.ListIndex
End Sub

Private Sub Text1_Change()
    ucChartBar1.LabelsFormats = Text1.text
End Sub

Private Sub CboNroSeries_Click()
    ucChartBar1.Clear
    Form_Load
    ucChartBar1.Refresh
End Sub

Private Sub HScroll1_Change()
    Dim Value As Collection
    Dim Users() As String
    Dim i As Long
        
    Users = Split("Jhon,Michael,Julia,McMartins,Matiu,Emili,Smit", ",")
    
    Set Value = New Collection
    For i = 0 To UBound(Users)
        Value.Add Users(i)
    Next
    ucChartBar1.AddAxisItems Value, False, HScroll1.Value, cCenter
    ucChartBar1.Refresh
End Sub

Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub


Private Sub Command1_Click()
    Dim Value As Collection
    Dim i As Long, j As Long
    Dim Min As Single, Max As Single
    
    If ChkNegative.Value Then Min = -200
    Max = 500
    
    Set Value = New Collection
    For i = 0 To 6
        Value.Add Random(Min, Max)
    Next
    ucChartBar1.UpdateSerie 0, "2000", &HFF8D11, Value

    ucChartBar1.Refresh

End Sub

Private Function Random(ByVal Min!, ByVal Max!)
    Random = Int((Max - Min + 1) * Rnd + Min)
End Function

Private Sub Form_Load()
    Dim Value As Collection
    Dim i As Long, j As Long
    Dim Palette() As String
    Dim Users() As String
    Dim Min As Single
    Dim Max As Single
    

    If CboNroSeries.ListIndex = -1 Then
        Combo2.ListIndex = 0
        CboLegendAlign.ListIndex = 0
        cboLabelsPositions.ListIndex = 0
        CboNroSeries.ListIndex = 1:
        Exit Sub
    End If
    
    Randomize Timer
    
    Palette = Split("&HFF8D11,&HA744E0,&H376CE6,&H40AB1A,&H7B006B", ",")
    Users = Split("Jhon,Michael,Julia,McMartins,Matiu,Emili,Smit", ",")
    
    Set Value = New Collection
    For i = 0 To UBound(Users)
        Value.Add Users(i)
    Next
    ucChartBar1.AddAxisItems Value, False, HScroll1.Value, cCenter
    
    If ChkNegative.Value Then Min = -200
    Max = 200
    
    For i = 0 To CboNroSeries.ListIndex
        Set Value = New Collection
        For j = 0 To 6
            Value.Add Random(Min, Max * (j + 1))
        Next
        ucChartBar1.AddSerie "200" & i, CLng(Palette(i)), Value
    Next
   
End Sub

Private Sub Form_Resize()
    ucChartBar1.Move Picture1.Width, 0, Me.ScaleWidth - Picture1.Width, Me.ScaleHeight
End Sub


