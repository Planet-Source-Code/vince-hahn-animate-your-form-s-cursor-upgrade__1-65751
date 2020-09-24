VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()
    
    ' the first two aren't required, but make it better for the animated cursor I've included.
    
    Me.Show
    Me.BackColor = RGB(0, 0, 0)
    
    'ok now this thing is required
    SetAniCursor Form1, App.Path & "\DTwannabe.ani"          'the "Me" thingy will also work in place of the form's actual name

End Sub

Private Sub Form_Unload(Cancel As Integer)

    KillAniCursor Form1                     ' the "Me" thingy will also work in place of the form's actual name

End Sub
