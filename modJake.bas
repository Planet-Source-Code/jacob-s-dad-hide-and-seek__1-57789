Attribute VB_Name = "modJake"
Option Explicit

'<!--                                                               __                          -->
'<!--              /\                                            /\(..)/\             _    ___    -->
'<!--             //\\                 |                        //\\)(//\\    -   -  |*\  |        -->
'<!--created by: //==\\    /\  //  -|- |_ () /\  // \|   |).   //  \  /  \\  (.) (.) |_/  |___      -->
'<!--           //    \\  //\\//    |  | |  //\\//   |         \|   \/   |/   '/     | \  |         -->
'<!--           \|    |/ //  \/            //  \/   _|                         \/    |  \ |___     -->
'<!--     on                                                                  (___)'              -->
'<!--  4/5/04 for my wonderful handsome son and sun JACOB MALAKAI MOORE                         -->
'______________________________________________________________________________________________________
'------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------
'______________________________________________________________________________________________________


Public Jake As IAgentCtlCharacterEx
Dim HidePlace As Integer
Dim NoRepeat As Integer

Public Sub Spell(Words As String)
On Error Resume Next
Jake.Speak Words
End Sub

Public Sub moveMe(X As Long, Y As Long)
Jake.MoveTo X, Y, 4
End Sub

Public Sub PointAt(X As Long, Y As Long)
Jake.GestureAt X, Y
End Sub

Public Sub act(Action As String)
Jake.Play (Action)
End Sub

Public Sub HideAway()
On Error Resume Next
Randomize
'If HidePlace = 0 Then
 '   HidePlace = Int(Rnd() * 100 + 1)
'End If
Do Until HidePlace < 10
    HidePlace = Int(Rnd() * 100 + 1)
Loop
If NoRepeat = HidePlace Then
    HidePlace = Int(Rnd() * 100 + 1)
        Do Until HidePlace < 10
            HidePlace = Int(Rnd() * 100 + 1)
        Loop
End If
Select Case HidePlace
    Case 10:
        Jake.Hide
            Call moveMe(frmFindMe.Shape1.Left, frmFindMe.Shape1.Top)
    Case 1:
        Jake.Hide
            Call moveMe(frmFindMe.Label2.Left, frmFindMe.Label2.Top)
    Case 2:
        Jake.Hide
            Call moveMe(frmFindMe.Label5.Left, frmFindMe.Label5.Top)
    Case 3:
        Jake.Hide
            Call moveMe(frmFindMe.Label1.Left, frmFindMe.Label1.Top)
    Case 4:
        Jake.Hide
            Call moveMe(frmFindMe.Label3.Left, frmFindMe.Label3.Top)
    Case 5:
        Jake.Hide
            Call moveMe(frmFindMe.Label4.Left, frmFindMe.Label4.Top)
    Case 6:
        Jake.Hide
            Call moveMe(frmFindMe.Label6.Left, frmFindMe.Label6.Top)
    Case 7:
        Jake.Hide
            Call moveMe(frmFindMe.Shape2.Left, frmFindMe.Shape2.Top)
    Case 8:
        Jake.Hide
            Call moveMe(frmFindMe.Label9.Left + 50, frmFindMe.Label9.Top + 25)
End Select
NoRepeat = HidePlace
End Sub
