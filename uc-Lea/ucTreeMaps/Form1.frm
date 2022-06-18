VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   8700
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
      TabIndex        =   15
      Top             =   4680
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
      Height          =   6030
      Left            =   0
      ScaleHeight     =   6030
      ScaleWidth      =   1815
      TabIndex        =   0
      Top             =   0
      Width           =   1815
      Begin VB.CommandButton Command2 
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
         TabIndex        =   11
         Top             =   4200
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E6C29B&
         Caption         =   "Round Corners"
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
         TabIndex        =   10
         Top             =   120
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
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
         TabIndex        =   9
         Top             =   480
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E6C29B&
         Caption         =   "Aling Left"
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
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E6C29B&
         Caption         =   "Aling Top"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E6C29B&
         Caption         =   "Aling Right"
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
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E6C29B&
         Caption         =   "Aling Bottom"
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
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   1575
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
         TabIndex        =   4
         Text            =   "${V}"
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00E6C29B&
         Caption         =   "Draw Title Serie"
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
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Iconos And CustomColors"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   5160
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   120
         List            =   "Form1.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   3720
         Width           =   1575
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
         TabIndex        =   13
         Top             =   2640
         Width           =   1215
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
         TabIndex        =   12
         Top             =   3480
         Width           =   1320
      End
   End
   Begin Proyecto1.FontMemRes FontMemRes1 
      Left            =   6840
      Top             =   2880
      _ExtentX        =   1270
      _ExtentY        =   1270
      bvData          =   "Form1.frx":0048
      bData           =   -1  'True
   End
   Begin Proyecto1.ucTreeMaps ucTreeMaps1 
      Height          =   3375
      Left            =   2160
      TabIndex        =   14
      Top             =   600
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5953
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelsVisible   =   0   'False
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
      LabelsFormats   =   "${V}"
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
Private Sub Check1_Click()
    ucTreeMaps1.CornerRound = IIf(Check1.Value, 10, 0)
End Sub

Private Sub Check2_Click()
    ucTreeMaps1.LegendVisible = Check2.Value
End Sub

Private Function Random(ByVal Min!, ByVal Max!)
    Random = Int((Max - Min + 1) * Rnd + Min)
End Function

Private Sub Check3_Click()
    ucTreeMaps1.DrawTilteSerie = Check3.Value
End Sub

Private Sub CmdPrint_Click()
    Dim oPic As StdPicture
    Printer.ScaleMode = vbPixels
    Set oPic = ucTreeMaps1.Image(Printer.ScaleWidth / 10, Printer.ScaleHeight / 10 / 2)
    Printer.PaintPicture oPic, 0, 0, Printer.ScaleWidth, Printer.ScaleHeight / 2
    Printer.EndDoc
End Sub

Private Sub Combo1_Click()
    ucTreeMaps1.Clear
    Form_Load
    ucTreeMaps1.Refresh
End Sub

Private Sub Command1_Click()
    Form2.Show
End Sub

Private Sub Command2_Click()
    ucTreeMaps1.Clear
    Form_Load
    ucTreeMaps1.Refresh
End Sub

Private Sub Form_Load()
    Dim Value As Collection
    Dim Lables As Collection
    Dim Icons As Collection
    Dim Palette() As String
    Dim Users() As String
    
    If Combo1.ListIndex = -1 Then Combo1.ListIndex = 4: Exit Sub
    
    Dim i As Long, j As Long
    
    Randomize Timer
    Palette = Split("&HFF8D11,&HA744E0,&H376CE6,&H40AB1A,&H7B006B", ",")
    Users = Split("Jhon,Michael,Julia,McMartins ,Matiu,Emili,Smit", ",")
    For i = 0 To Combo1.ListIndex
        Set Value = New Collection
        Set Lables = New Collection
        For j = 0 To 6
            Value.Add Random(5, 50 * (i + 1))
            Lables.Add Users(j)
        Next
        
        ucTreeMaps1.AddLineSeries "200" & i, CLng(Palette(i)), Value, Lables
    Next
 

End Sub

Private Sub Form_Resize()
    ucTreeMaps1.Move Picture1.Width, 0, Me.ScaleWidth - Picture1.Width, Me.ScaleHeight
End Sub

Private Sub Option1_Click(Index As Integer)
    ucTreeMaps1.LegendAlign = Index
End Sub

Private Sub Text1_Change()
    ucTreeMaps1.LabelsFormats = Text1.Text
End Sub
