VERSION 5.00
Begin VB.Form Frmenu 
   Caption         =   "This is a menu. There is nothing more to say."
   ClientHeight    =   5790
   ClientLeft      =   4200
   ClientTop       =   3450
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Frmenu.frx":0000
   ScaleHeight     =   5790
   ScaleWidth      =   9600
   Begin VB.Label butquit 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   4080
      Width           =   5535
   End
   Begin VB.Label butoptions 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   3120
      Width           =   5535
   End
   Begin VB.Label butStartgame 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   2040
      TabIndex        =   0
      Top             =   2280
      Width           =   5535
   End
End
Attribute VB_Name = "Frmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub butoptions_Click()
    Me.Hide
    Frmopt.Show 'If the label is clicked show approiate form, same for all label click functions
End Sub

Private Sub butquit_Click()
    Unload Me
End Sub

Private Sub butStartgame_Click()
    Me.Hide
    frminst.Show
End Sub

Private Sub Form_Keypress(keyascii As Integer)
If keyascii = 115 Then 'If correct keys are pressed than load correct form
    Me.Hide
    frminst.Show
End If
If keyascii = 111 Then
    Me.Hide
    Frmopt.Show
End If
If keyascii = 113 Then
    Unload Me
End If
If keyascii = 13 Then
    Me.Hide
    frmtest.Show
End If
End Sub

