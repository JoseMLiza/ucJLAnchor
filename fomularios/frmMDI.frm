VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   6900
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13425
   Icon            =   "frmMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuTools 
      Caption         =   "&Abrir"
      Begin VB.Menu mnuForms 
         Caption         =   "Child"
         Index           =   0
      End
      Begin VB.Menu mnuForms 
         Caption         =   "Not Child"
         Index           =   1
      End
      Begin VB.Menu mnuForms 
         Caption         =   "Controles Lea"
         Index           =   2
      End
      Begin VB.Menu mnuForms 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuForms 
         Caption         =   "Child 2"
         Index           =   4
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuForms_Click(Index As Integer)
    Select Case Index
        Case 0 'Child
            frmChild.Show
        Case 1 'Not Child
            frmNoChild.Show
        Case 2 'Controles Lea
            frmControlsLea.Show
        Case 3 'Separator
        Case 4 'Child 2
            frmChild2.Show
    End Select
End Sub
