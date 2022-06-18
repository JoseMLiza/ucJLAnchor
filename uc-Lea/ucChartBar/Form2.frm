VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FAF6F5&
   Caption         =   "Colors, Fonts, And Round Borders"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9825
   LinkTopic       =   "Form2"
   ScaleHeight     =   6855
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin Proyecto1.ucChartBar ucChartBar1 
      Height          =   3135
      Index           =   3
      Left            =   5160
      TabIndex        =   3
      Top             =   3480
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5530
      Title           =   "ucChartBar1"
      ForeColor       =   10132122
      LinesColor      =   15132391
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FillOpacity     =   80
      Border          =   -1  'True
      VerticalLines   =   -1  'True
      HorizontalLines =   0   'False
      ChartOrientation=   1
      LegendVisible   =   0   'False
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleForeColor  =   0
      BorderRound     =   10
   End
   Begin Proyecto1.ucChartBar ucChartBar1 
      Height          =   3135
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5530
      Title           =   "ucChartBar2"
      BackColor       =   4008231
      ForeColor       =   10132122
      LinesColor      =   5255974
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FillOpacity     =   20
      VerticalLines   =   -1  'True
      FillGradient    =   -1  'True
      HorizontalLines =   0   'False
      ChartOrientation=   1
      LegendVisible   =   0   'False
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleForeColor  =   10526880
      BorderRound     =   10
   End
   Begin Proyecto1.ucChartBar ucChartBar1 
      Height          =   3135
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5530
      Title           =   "ucChartBar2"
      BackColor       =   4008231
      ForeColor       =   10132122
      LinesColor      =   5255974
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FillOpacity     =   20
      FillGradient    =   -1  'True
      LegendVisible   =   0   'False
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleForeColor  =   10526880
      BorderRound     =   10
   End
   Begin Proyecto1.ucChartBar ucChartBar1 
      Height          =   3135
      Index           =   2
      Left            =   5160
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5530
      Title           =   "ucChartBar1"
      ForeColor       =   10132122
      LinesColor      =   15132391
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FillOpacity     =   80
      Border          =   -1  'True
      LegendVisible   =   0   'False
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleForeColor  =   0
      BorderRound     =   10
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00281E1E&
      FillStyle       =   0  'Solid
      Height          =   7215
      Left            =   0
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim i As Long, j As Long
    Dim Value As Collection
    Dim cPalette As Collection

    Randomize Timer
    


    Set Value = New Collection
    For i = 0 To 5
        Value.Add Random(0, 1000)
    Next
    
    ucChartBar1(0).AddSerie "Serie 1", vbBlue, Value
    ucChartBar1(1).AddSerie "Serie 1", vbRed, Value
    
    Set cPalette = New Collection
    With cPalette
        .Add &HFF8D11
        .Add &HA744E0
        .Add &H376CE6
        .Add &H40AB1A
        .Add &H5CD9FB
        .Add &H7B006B
    End With
    
    For i = 2 To ucChartBar1.Count - 1
        ucChartBar1(i).AddSerie "Serie 1", vbRed, Value, cPalette
    Next
End Sub


Private Function Random(ByVal Min!, ByVal Max!)
    Random = Int((Max - Min + 1) * Rnd + Min)
End Function

Private Sub Form_Resize()
    Dim i As Long
    If (Me.ScaleWidth < 2000) Or (Me.ScaleHeight < 2000) Then Exit Sub
    For i = 0 To 3
        ucChartBar1(i).Font.Size = Me.ScaleHeight / 15 / 50
        ucChartBar1(i).TitleFont.Size = Me.ScaleHeight / 15 / 32
    Next
    Shape1.Move 0, 0, Me.ScaleWidth / 2, Me.ScaleHeight
    ucChartBar1(0).Move 100, 100, Me.ScaleWidth / 2 - 200, Me.ScaleHeight / 2 - 100
    ucChartBar1(1).Move 100, Me.ScaleHeight / 2 + 100, Me.ScaleWidth / 2 - 200, Me.ScaleHeight / 2 - 200
    ucChartBar1(2).Move Me.ScaleWidth / 2 + 100, 100, Me.ScaleWidth / 2 - 200, Me.ScaleHeight / 2 - 100
    ucChartBar1(3).Move Me.ScaleWidth / 2 + 100, Me.ScaleHeight / 2 + 100, Me.ScaleWidth / 2 - 200, Me.ScaleHeight / 2 - 200

End Sub
