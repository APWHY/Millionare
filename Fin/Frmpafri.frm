VERSION 5.00
Begin VB.Form Frmpafri 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5820
   ClientLeft      =   4470
   ClientTop       =   3435
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frmpafri.frx":0000
   ScaleHeight     =   5820
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timpho 
      Interval        =   1000
      Left            =   7920
      Top             =   840
   End
   Begin VB.Label lblgoback 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label lblwr 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   3
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label lblphoney 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   2
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label lblfriendadvice 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Shape Circle 
      BackColor       =   &H00404080&
      BorderColor     =   &H0000FF00&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   14
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   480
      Width           =   375
   End
   Begin VB.Shape Circle 
      BackColor       =   &H00404080&
      BorderColor     =   &H0000FF00&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   13
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   840
      Width           =   375
   End
   Begin VB.Shape Circle 
      BackColor       =   &H00404080&
      BorderColor     =   &H0000FF00&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   12
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   375
   End
   Begin VB.Shape Circle 
      BackColor       =   &H00404080&
      BorderColor     =   &H0000FF00&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   11
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   2760
      Width           =   375
   End
   Begin VB.Shape Circle 
      BackColor       =   &H00404080&
      BorderColor     =   &H0000FF00&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   10
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape Circle 
      BackColor       =   &H00404080&
      BorderColor     =   &H0000FF00&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   9
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   4200
      Width           =   375
   End
   Begin VB.Shape Circle 
      BackColor       =   &H00404080&
      BorderColor     =   &H0000FF00&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   8
      Left            =   3360
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   375
   End
   Begin VB.Shape Circle 
      BackColor       =   &H00404080&
      BorderColor     =   &H0000FF00&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   7
      Left            =   5520
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   375
   End
   Begin VB.Shape Circle 
      BackColor       =   &H00404080&
      BorderColor     =   &H0000FF00&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   6
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   4320
      Width           =   375
   End
   Begin VB.Shape Circle 
      BackColor       =   &H00404080&
      BorderColor     =   &H0000FF00&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   5
      Left            =   6720
      Shape           =   3  'Circle
      Top             =   3600
      Width           =   375
   End
   Begin VB.Shape Circle 
      BackColor       =   &H00404080&
      BorderColor     =   &H0000FF00&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   4
      Left            =   6960
      Shape           =   3  'Circle
      Top             =   2760
      Width           =   375
   End
   Begin VB.Shape Circle 
      BackColor       =   &H00404080&
      BorderColor     =   &H0000FF00&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   3
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   375
   End
   Begin VB.Shape Circle 
      BackColor       =   &H00404080&
      BorderColor     =   &H0000FF00&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   2
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   840
      Width           =   375
   End
   Begin VB.Shape Circle 
      BackColor       =   &H00404080&
      BorderColor     =   &H0000C000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   5400
      Shape           =   3  'Circle
      Top             =   480
      Width           =   375
   End
   Begin VB.Shape Circle 
      BackColor       =   &H00404080&
      BorderColor     =   &H0000FF00&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   15
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   360
      Width           =   375
   End
   Begin VB.Label lbltim 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   3240
      TabIndex        =   0
      Top             =   2160
      Width           =   2775
   End
End
Attribute VB_Name = "Frmpafri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Timer As Byte
Dim Fakeans As String
Dim Realans As String
Dim RndPhon As Byte
Dim RndFrie As Byte
Private Sub Form_Load()
Timer = 5
    AnsA = 1
    AnsB = 2
    AnsC = 3
    AnsD = 4
    Select Case Correct(Random)
        Case Is = 1
        AnsA = 4
        AnsD = 1
        Case Is = 2
        AnsB = 4
        AnsD = 2
        Case Is = 3
        AnsC = 4
        AnsD = 3
    End Select
End Sub
Private Sub Form_Keypress(Keyascpp As Integer)
If Keyascpp = 13 Then
Unload Me
End If
End Sub
Private Sub timpho_Timer()
If Timer > 1 Then
Timer = Timer - 1
Else: timpho.Enabled = False
    lblgoback.Caption = "Press Return to go back"
    Randomize
    RndPhon = Int(Rnd * Quedonecount * 5) + 1
    Phoney = 100 - RndPhon
    If Phoney < ((Quedonecount * 5) + 10) Then 'choses if you get a wrong or right answer
      Randomize
      RndFrie = Int(Rnd * 3) + 1
          Select Case RndFrie
              Case AnsA = RndFrie: Fakeans = "A" 'generates wrong answer
              Case AnsB = RndFrie: Fakeans = "B"
              Case AnsC = RndFrie: Fakeans = "C"
              Case AnsD = RndFrie: Fakeans = "D"
          End Select
          If Fiftychange = True Then
              If Fakeans = Random50 Or Fakeans = Random250 Then
                  Do
                      Randomize
                      RndFrie = Int(Rnd * 3) + 1
                      Select Case RndFrie
                          Case 1: Fakeans = "A" 'generates another wrong answer if one of the answers is 50/50'd
                          Case 2: Fakeans = "B"
                          Case 3: Fakeans = "C"
                          Case 4: Fakeans = "D"
                      End Select
                  Loop Until RndFrie <> Random50 Or RndFrie <> Random250
              End If
           End If
                 lblfriendadvice.Caption = "Your Friend is        & sure that        is the answer"
                 lblphoney.Caption = Phoney
                 lblwr.Caption = Fakeans
    Else
       Select Case AnsD 'note AnsD holds the value for the correct Answer
          Case 1: Realans = "A"
          Case 2: Realans = "B"
          Case 3: Realans = "C"
          Case 4: Realans = "D"
      End Select
      lblfriendadvice.Caption = "Your Friend is       % sure that         is the answer"
      lblphoney.Caption = Phoney 'shows what your friend chose
      lblwr.Caption = Realans
End If
End If
lbltim.Caption = Timer 'shows time in seconds
End Sub
