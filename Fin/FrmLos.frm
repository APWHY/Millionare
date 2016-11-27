VERSION 5.00
Begin VB.Form FrmLos 
   Caption         =   "You Are Slighly more stupid than that guy..."
   ClientHeight    =   5790
   ClientLeft      =   4200
   ClientTop       =   3450
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "FrmLos.frx":0000
   ScaleHeight     =   5790
   ScaleWidth      =   9600
End
Attribute VB_Name = "FrmLos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Keypress(keyascii As Integer)
 If keyascii = 13 Then
    Frmenu.Show
    Unload Me 'load menu when enter is pressed
End If
 
End Sub
