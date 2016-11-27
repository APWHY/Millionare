VERSION 5.00
Begin VB.Form Frmqe 
   Caption         =   "You shouldn't be reading this..."
   ClientHeight    =   5775
   ClientLeft      =   4200
   ClientTop       =   3450
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmque.frx":0000
   ScaleHeight     =   5775
   ScaleWidth      =   9570
   Begin VB.Line lnpoll2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   2080
      X2              =   1080
      Y1              =   2760
      Y2              =   3260
   End
   Begin VB.Line lnpoll1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   1080
      X2              =   2080
      Y1              =   2760
      Y2              =   3260
   End
   Begin VB.Line lnphone2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   2080
      X2              =   1080
      Y1              =   2040
      Y2              =   2540
   End
   Begin VB.Line lnphone1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   1080
      X2              =   2080
      Y1              =   2040
      Y2              =   2540
   End
   Begin VB.Line ln50502 
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   2005
      X2              =   1005
      Y1              =   1425
      Y2              =   1925
   End
   Begin VB.Line ln5050 
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   1005
      X2              =   2005
      Y1              =   1425
      Y2              =   1925
   End
   Begin VB.Image Prize 
      Height          =   165
      Index           =   14
      Left            =   6525
      Picture         =   "frmque.frx":B4F44
      Top             =   863
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image Prize 
      Height          =   135
      Index           =   13
      Left            =   6525
      Picture         =   "frmque.frx":B6008
      Top             =   1028
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image Prize 
      Height          =   120
      Index           =   12
      Left            =   6525
      Picture         =   "frmque.frx":B6DCC
      Top             =   1163
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image Prize 
      Height          =   150
      Index           =   11
      Left            =   6525
      Picture         =   "frmque.frx":B7A10
      Top             =   1283
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image Prize 
      Height          =   135
      Index           =   10
      Left            =   6525
      Picture         =   "frmque.frx":B8954
      Top             =   1433
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image Prize 
      Height          =   120
      Index           =   9
      Left            =   6525
      Picture         =   "frmque.frx":B9718
      Top             =   1568
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image Prize 
      Height          =   135
      Index           =   8
      Left            =   6525
      Picture         =   "frmque.frx":BA35C
      Top             =   1688
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image Prize 
      Height          =   120
      Index           =   7
      Left            =   6525
      Picture         =   "frmque.frx":BB120
      Top             =   1823
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image Prize 
      Height          =   150
      Index           =   6
      Left            =   6525
      Picture         =   "frmque.frx":BBD64
      Top             =   1943
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image Prize 
      Height          =   135
      Index           =   5
      Left            =   6525
      Picture         =   "frmque.frx":BCCA8
      Top             =   2093
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image Prize 
      Height          =   135
      Index           =   4
      Left            =   6525
      Picture         =   "frmque.frx":BDA6C
      Top             =   2228
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image Prize 
      Height          =   105
      Index           =   3
      Left            =   6525
      Picture         =   "frmque.frx":BE830
      Top             =   2365
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image Prize 
      Height          =   135
      Index           =   2
      Left            =   6525
      Picture         =   "frmque.frx":BF2F4
      Top             =   2468
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image Prize 
      Height          =   150
      Index           =   1
      Left            =   6520
      Picture         =   "frmque.frx":C00B8
      Top             =   2603
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image Prize 
      Height          =   165
      Index           =   0
      Left            =   6525
      Picture         =   "frmque.frx":C0FFC
      Top             =   2760
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Label lblwalk 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   960
      TabIndex        =   13
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblphoney 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   1060
      TabIndex        =   12
      Top             =   2010
      Width           =   1095
   End
   Begin VB.Label lblpercen 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   3
      Left            =   5160
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblpercen 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   2
      Left            =   4440
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblpercen 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblpercen 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   4
      Left            =   5760
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Bar 
      Height          =   1500
      Index           =   2
      Left            =   4500
      Picture         =   "frmque.frx":C20C0
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Bar 
      Height          =   1500
      Index           =   3
      Left            =   5160
      Picture         =   "frmque.frx":C47DC
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Bar 
      Height          =   1500
      Index           =   4
      Left            =   5760
      Picture         =   "frmque.frx":C6EF8
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Bar 
      Height          =   1500
      Index           =   1
      Left            =   3900
      Picture         =   "frmque.frx":C9614
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgraph 
      Height          =   2130
      Left            =   3720
      Picture         =   "frmque.frx":CBD30
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.Label lblpolaud 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   555
      Left            =   1080
      TabIndex        =   7
      Top             =   2640
      Width           =   1005
   End
   Begin VB.Label lbl5050 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   960
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblfirnexque 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CLICK ME To start the game"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1095
      Left            =   2040
      TabIndex        =   5
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label lblque 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   3720
      Width           =   6135
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblans 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   4
      Top             =   4800
      Width           =   2415
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblans 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   5640
      TabIndex        =   3
      Top             =   4440
      Width           =   2415
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblans 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   2
      Top             =   4440
      Width           =   2415
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblans 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   4
      Left            =   5640
      TabIndex        =   1
      Top             =   4800
      Width           =   2415
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmqe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Wrong As Boolean
Dim Counter As Byte
Dim Audw(4) As Byte
Dim RanPaa(4) As Byte
Dim Pollaud As Boolean
Dim Deletegraph As Boolean
Dim Ning(51) As Boolean
Dim Flag As Boolean



Private Sub Form_Load()
Questno = 1 'What question are we up to?
Fiftyfifty = False
Pollaud = False
Phoney = False
Questans = False 'Has the question already been answered?
Quedonecount = 1
For Counter = 1 To 51 'This defines all Ning's for later on. There are 51 Ning's, because there are 51 possible values for Ning
    Ning(Counter) = False
Next

'defining variables for lifelines and if questions are answered
Open App.Path & "/fq.txt" For Input As #1
    For Counter = 1 To 51 '(this is the number of questions)
        Input #1, Question(Counter)
        Input #1, Answer1(Counter)
        Input #1, Answer2(Counter)
        Input #1, Answer3(Counter)
        Input #1, Answer4(Counter)
        Input #1, Correct(Counter)
    Next
Close #1

End Sub

Private Sub Form_Keypress(Keyascnn As Integer)
If Questno = 1 Then Wrong = False



Select Case Keyascnn 'keying in answer
    Case Is = 97
    If lblans(1) = "" Then 'If the user selects a 5050'd option
        MsgBox "Idiot"
    Else
        Guess = 1
    End If
    Case Is = 98
    If lblans(2) = "" Then
        MsgBox "Idiot"
    Else
        Guess = 2
    End If
    Case Is = 99
    If lblans(3) = "" Then
        MsgBox "Idiot"
    Else
        Guess = 3
    End If
    Case Is = 100
    If lblans(4) = "" Then
        MsgBox "Idiot"
    Else
        Guess = 4
    End If
    Case Else
        Guess = 6
End Select

If Fiftychange = True Then
    Fiftychange = False
End If
If Deletegraph = True Then 'gets rid of the graph that is bought up by poll the audience
    imgraph.Visible = False
    Bar(1).Visible = False
    Bar(2).Visible = False
    Bar(3).Visible = False
    Bar(4).Visible = False
    lblpercen(1).Visible = False
    lblpercen(3).Visible = False
    lblpercen(2).Visible = False
    lblpercen(4).Visible = False
End If
Select Case Guess 'check for correctness
    Case Is = Correct(Random)
        MsgBox "Correct!", vbOKOnly, "Win..."
        Quedonecount = Quedonecount + 1
    Case Is <> Correct(Random) 'if wrong
            If Guess < 5 Then
            MsgBox "You have failed. You don't even deserve a full stop at the end of this sentence", vbOKOnly, "Fail"
                    If Quedonecount = 1 Then
                FrmLos.Show
                Unload Me
            End If
            Select Case Quedonecount 'changes dollar amount to last checkpoint
                Case 1 To 5
                    Quedonecount = 1
                Case 6 To 10
                    Quedonecount = 6
                Case 11 To 15
                    Quedonecount = 11
            End Select
        Frmwin.Show
        Unload Me
            Else
            MsgBox "Are you Retarded?", vbOKOnly, "???" 'if the person pressed a key other than a,b,c or d
        End If
End Select
If Quedonecount > 15 Then 'check for winner
Frmwin.Show
Unload Me
End If
End Sub


Private Sub lblphoney_Click()
If Phone = False Then 'loads the phone-a-friend form
Frmpafri.Show
lnphone1.Visible = True 'loads crosses
lnphone2.Visible = True
Phone = True
Else
MsgBox "Don't Cheat!"
End If
End Sub

Private Sub lbl5050_Click()
If Fiftyfifty = False Then
    Fiftyfifty = True
    Fiftychange = True
    ln5050.Visible = True 'loads crosses
    ln50502.Visible = True
        Do
        Randomize
        Random50 = Int(Rnd * 3) + 1
        Loop Until Random50 <> Correct(Random)
        Do
        Randomize
        Random250 = Int(Rnd * 3) + 1
        Loop Until Random50 <> Random250 And Random250 <> Correct(Random) 'Picks two random numbers that are different and not correct answer
    lblans(Random50).Caption = ""
    lblans(Random250).Caption = ""
Else
    MsgBox "Don't Cheat"
End If

End Sub

Private Sub lblfirnexque_Click()
If Questno <> Quedonecount Then 'Has the question been answered?
   MsgBox "Don't Cheat!"
Else
        Randomize
        Flag = False
        Do
            Random = Int((Rnd * 17) + 1)
                Select Case Questno 'Changing Random to suit the diffculty level
                    Case 6 To 10: Random = Random + 17
                    Case 11 To 15: Random = Random + 34
                End Select
            If Ning(Random) = False Then 'See line 20 of code, stops same question from being shown twice
                Flag = True
                Ning(Random) = True
            End If
        Loop Until Flag = True
    Select Case Questno 'how many questions have been answered?
        Case Is <= 5:
            lblque.Caption = Question(Random) 'Loads random question easy
            lblans(1).Caption = Answer1(Random)
            lblans(2).Caption = Answer2(Random)
            lblans(3).Caption = Answer3(Random)
            lblans(4).Caption = Answer4(Random)
            Questno = Questno + 1
        Case 6 To 10:
            lblque.Caption = Question(Random) 'Loads random question medium
            lblans(1).Caption = Answer1(Random)
            lblans(2).Caption = Answer2(Random)
            lblans(3).Caption = Answer3(Random)
            lblans(4).Caption = Answer4(Random)
            Questno = Questno + 1
        Case 11 To 15:
            lblque.Caption = Question(Random) 'Loads random question hard
            lblans(1).Caption = Answer1(Random)
            lblans(2).Caption = Answer2(Random)
            lblans(3).Caption = Answer3(Random)
            lblans(4).Caption = Answer4(Random)
            Questno = Questno + 1
    End Select
    lblfirnexque.Caption = "Click me for next question NOW" 'Changes the original caption
End If
Prize(Questno - 2).Visible = True 'Show the Prize the person is up to
End Sub

Private Sub lblpolaud_Click()
    If Pollaud = False Then
        lnpoll1.Visible = True
        lnpoll2.Visible = True
        Pollaud = True
    AnsA = 1
    AnsB = 2
    AnsC = 3
    AnsD = 4
    Select Case Correct(Random) '4 holds the correct answer, used later to assign the bar that is most likely to
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
    Audw(4) = 75
        Randomize
        RanPaa(1) = Int(Rnd * Quedonecount * 4) 'assigns values to A,B,c,D in graph
        Audw(4) = Audw(4) - RanPaa(1)
        Audw(1) = RanPaa(1) + 15
        Randomize
        RanPaa(2) = Int(Rnd * Audw(1))
        Audw(2) = Audw(1) - RanPaa(2)
        Audw(1) = Audw(1) - Audw(2)
        Randomize
        RanPaa(3) = Int(Rnd * Audw(2))
        Audw(3) = Audw(2) - RanPaa(3)
        Audw(2) = Audw(2) - Audw(3)
        Randomize
        RanPaa(4) = Int(Rnd * Audw(3))
        Audw(4) = 100 - Audw(3) - Audw(2) - Audw(1)

Bar(AnsA).Height = Audw(1) * 15 'sets height of bars
Bar(AnsB).Height = Audw(2) * 15
Bar(AnsC).Height = Audw(3) * 15
Bar(4).Height = Audw(Correct(Random)) * 15
Bar(AnsD).Height = Audw(4) * 15
lblpercen(AnsA).Caption = Audw(1) 'shows height of bars
lblpercen(AnsB).Caption = Audw(2)
lblpercen(AnsC).Caption = Audw(3)
lblpercen(4).Caption = Audw(Correct(Random))
lblpercen(AnsD).Caption = Audw(4)
If Fiftychange = True Then 'If 5050 was used with this graph
Bar(Random50).Height = 0
Bar(Random250).Height = 0
lblpercen(Random50).Caption = "0"
lblpercen(Random250).Caption = "0"
End If
lblpercen(1).Visible = True 'show the graph
lblpercen(3).Visible = True
lblpercen(2).Visible = True
lblpercen(4).Visible = True
imgraph.Visible = True
Bar(1).Visible = True
Bar(2).Visible = True
Bar(3).Visible = True
Bar(4).Visible = True
Deletegraph = True
Else
MsgBox "Don't Cheat!"
End If
End Sub

Private Sub lblwalk_Click()
If Quedonecount = 1 Then 'loads the winning form
    FrmLos.Show
    Unload Me
Else
    Frmwin.Show
    Unload Me
End If
End Sub
