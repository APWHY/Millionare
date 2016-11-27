VERSION 5.00
Begin VB.Form frmtest 
   Caption         =   "Why are you here? This is meant to be hidden!"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   Begin VB.Label t 
      Caption         =   "Enter to go back..."
      Height          =   1095
      Left            =   2400
      TabIndex        =   0
      Top             =   1080
      Width           =   4335
   End
End
Attribute VB_Name = "frmtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Open App.Path & "/fq.txt" For Input As #1
    For Counter = 1 To 51 '(this is the number of questions)
        Input #1, Question(Counter)
        Input #1, Answer1(Counter) 'This is a test form, Left here as an easter egg
        Input #1, Answer2(Counter)
        Input #1, Answer3(Counter)
        Input #1, Answer4(Counter)
        Input #1, Correct(Counter)
    Next
Close #1
End Sub
Private Sub Form_Keypress(keyascii As Integer)
If keyascii = 13 Then
 Frmenu.Show 'go back to menu when enter is pressed
 Unload Me
End If

t.Caption = Question(10)
Print Answer1(10)
Print Answer2(10) 'testing answer printing
Print Answer3(10)
Print Answer4(10)
Print Correct(10)
End Sub


