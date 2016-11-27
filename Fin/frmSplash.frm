VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5715
   ClientLeft      =   4185
   ClientTop       =   3315
   ClientWidth     =   9555
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   5715
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timsplah 
      Interval        =   1000
      Left            =   240
      Top             =   120
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Keypress(keyascii As Integer)
    Unload Me
    Frmenu.Show
End Sub

Private Sub frmsplash_Click()
    Unload Me
    Frmenu.Show
End Sub

Private Sub timsplah_Timer()
    Frmenu.Show
    Unload Me
End Sub


'all this loads the menu form
