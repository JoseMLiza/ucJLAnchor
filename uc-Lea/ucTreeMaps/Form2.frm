VERSION 5.00
Begin VB.Form Form2 
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6750
   LinkTopic       =   "Form2"
   ScaleHeight     =   5700
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin Proyecto1.ucTreeMaps ucTreeMaps1 
      Height          =   3255
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelsVisible   =   0   'False
      LegendVisible   =   0   'False
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
      CornerRound     =   10
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Font Awesome 5 Brands Regular"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowToolTips    =   0   'False
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim Value As Collection
    Dim Lables As Collection
    Dim Icons As Collection
    Dim keys As Collection
    Dim CustomColors As Collection
    Dim i As Long

    Set Value = New Collection
    With Value
        .Add 100
        .Add 80
        .Add 30
        .Add 25
        .Add 90
        .Add 60
        .Add 70
        .Add 20
    End With
    
    Set Lables = New Collection
    With Lables
        .Add "Facebook"
        .Add "Intagram"
        .Add "Wikipedia"
        .Add "Pinterest"
        .Add "WhatsApp"
        .Add "Twiter"
        .Add "Youtube"
        .Add "Vimeo"
    End With
    
    'Icons font "Font Awesome 5 Brands"
    Set Icons = New Collection
    With Icons
        .Add &HF082
        .Add &HF16D
        .Add &HF266
        .Add &HF231
        .Add &HF232
        .Add &HF099
        .Add &HF167
        .Add &HF194
    End With

   ' For Events ItemClick
    Set keys = New Collection
    With keys
        .Add "Facebook"
        .Add "Intagram"
        .Add "Wikipedia"
        .Add "Pinterest"
        .Add "WhatsApp"
        .Add "Twiter"
        .Add "Youtube"
        .Add "Vimeo"
    End With
    
    'Customize Item Color
    Set CustomColors = New Collection
    With CustomColors
        .Add &H9E2312
        .Add &H7B006B
        .Add &HB3D9&
        .Add &HA744E0
        .Add &H40AB1A
        .Add &HFF8D11
        .Add &H376CE6
        .Add &H7B6B
    End With
    
    ucTreeMaps1.AddLineSeries vbNullString, 0, Value, Lables, Icons, keys, CustomColors
End Sub

Private Sub Form_Resize()
    ucTreeMaps1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub ucTreeMaps1_ItemClick(Key As Variant)
    MsgBox Key
End Sub
