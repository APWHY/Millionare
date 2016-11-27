VERSION 5.00
Begin VB.Form Frmwin 
   Caption         =   "Good work! Now go and write this yourself!"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Frmwin.frx":0000
   ScaleHeight     =   5790
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lbldollar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   735
      Left            =   2520
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Lblwine 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   735
      Left            =   3360
      TabIndex        =   0
      Top             =   3120
      Width           =   6975
   End
End
Attribute VB_Name = "Frmwin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Keypress(Keyascoo As Integer)
 If Keyascoo = 13 Then
    Frmenu.Show
    Unload Me
 End If 'if enter is pressed load menu
End Sub
Private Sub Form_Load()
Select Case Quedonecount
    Case Is = 1
        Lblwine.Caption = "NOTHING" 'shows what you won
        lbldollar.Caption = ""
    Case Is = 2
        Lblwine.Caption = "100"
    Case Is = 3
        Lblwine.Caption = "200"
    Case Is = 4
        Lblwine.Caption = "300"
    Case Is = 5
        Lblwine.Caption = "500"
    Case Is = 6
        Lblwine.Caption = "1,000"
    Case Is = 7
        Lblwine.Caption = "2,000"
    Case Is = 8
        Lblwine.Caption = "4,000"
    Case Is = 9
        Lblwine.Caption = "8,000"
    Case Is = 10
        Lblwine.Caption = "16,000"
    Case Is = 11
        Lblwine.Caption = "32,000"
    Case Is = 12
        Lblwine.Caption = "64,000"
    Case Is = 13
        Lblwine.Caption = "125,000"
    Case Is = 14
        Lblwine.Caption = "250,000"
    Case Is = 15
        Lblwine.Caption = "500,000"
    Case Is = 16
        Lblwine.Caption = "MILLIONARE"
        lbldollar.Caption = ""
End Select
        
    

End Sub

