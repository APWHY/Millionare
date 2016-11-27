Attribute VB_Name = "questionselect"
Public Sub Form_Keypress(Keyascnn As Integer)
If Questans = False Then
    If Keyascnn = 97 Then 'If answer A is selected
        If Answer1(Random) = Correct(Random) Then 'If Correct
            Wrong = False
            Questans = True
        Else
            Wrong = True 'If Wrong
            Questans = True
    End If
    End If
    If Keyascnn = 98 Then 'If answer B is selected
        If Answer1(Random) = Correct(Random) Then 'If Correct
            Wrong = False
            Questans = True
        Else
            Wrong = True 'If Wrong
            Questans = True
    End If
    End If
    If Keyascnn = 99 Then 'If answer C is selected
        If Answer1(Random) = Correct(Random) Then 'If Correct
            Wrong = False
            Questans = True
        Else
            Wrong = True 'If Wrong
            Questans = True
    End If
    End If
    If Keyascnn = 100 Then 'If answer D is selected
        If Answer1(Random) = Correct(Random) Then 'If Correct
            Wrong = False
            Questans = True
        Else
            Wrong = True 'If Wrong
            Questans = True
    End If
    End If
End If
End Sub

