VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   5280
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14280
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   14280
   StartUpPosition =   3  'Windows Default
   Begin Proyecto1.ucChartArea ucChartArea1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   6376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LinesCurve      =   -1  'True
      VerticalLines   =   -1  'True
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
   End
   Begin VB.Menu MnuPrint 
      Caption         =   "Print"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim Value As Collection
    Set Value = New Collection
    
    Value.Add "Enero"
    Value.Add "Febrero"
    Value.Add "Marzo"
    Value.Add "Abril"
    Value.Add "Mayo"
    Value.Add "Junio"
    ucChartArea1.AddAxisItems Value
    
    Set Value = New Collection
    With Value
        .Add 2
        .Add 5
        .Add 7
        .Add -10
        .Add 5
        .Add 10
    End With
    ucChartArea1.AddLineSeries "2007", Value, vbRed
    'Exit Sub
    Set Value = New Collection
    With Value
        .Add 8
        .Add 4
        .Add 45
        .Add -15
        .Add 9
        .Add 14
    End With
    ucChartArea1.AddLineSeries "2008", Value, vbBlue
    Set Value = New Collection
    With Value
        .Add 14
        .Add 8
        .Add 16
        .Add 4
        .Add 24
        .Add 3
    End With
    ucChartArea1.AddLineSeries "2009", Value, vbGreen
    ucChartArea1.Refresh
    
End Sub

Private Sub Form_Resize()
    ucChartArea1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub MnuPrint_Click()
    Dim oPic As StdPicture
    Printer.ScaleMode = vbPixels
    Set oPic = ucChartArea1.Image(Printer.ScaleWidth / 10, Printer.ScaleHeight / 10 / 2)
    Printer.PaintPicture oPic, 0, 0, Printer.ScaleWidth, Printer.ScaleHeight / 2
    Printer.EndDoc
End Sub
