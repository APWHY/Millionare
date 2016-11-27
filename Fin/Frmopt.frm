VERSION 5.00
Begin VB.Form Frmopt 
   Caption         =   "My failed options"
   ClientHeight    =   5790
   ClientLeft      =   4200
   ClientTop       =   3450
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   9600
   Begin VB.Label Label1 
      Caption         =   "I am the options Enter to go back. This is because there is no time for me to do options.. the game"
      Height          =   1815
      Left            =   2880
      TabIndex        =   0
      Top             =   1920
      Width           =   4335
   End
End
Attribute VB_Name = "Frmopt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub From_Load()

End Sub
Private Sub Form_Keypress(keyascii As Integer)
If keyascii = 13 Then
    Frmenu.Show 'loads menu when enter is pressed
    Me.Hide
End If
End Sub

