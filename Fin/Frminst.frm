VERSION 5.00
Begin VB.Form Frminst 
   Caption         =   "Burn after Reading"
   ClientHeight    =   5790
   ClientLeft      =   4200
   ClientTop       =   3450
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   Picture         =   "Frminst.frx":0000
   ScaleHeight     =   5790
   ScaleWidth      =   9600
End
Attribute VB_Name = "frminst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Keypress(keyascii As Integer)
    If keyascii = 13 Then
        frmqe.Show
        Unload Me 'load question form when enter is pressed
    End If
End Sub

