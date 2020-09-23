VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Begin VB.Form frmFindMe 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   9075
   ClientLeft      =   555
   ClientTop       =   705
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   ScaleHeight     =   605
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   749
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   5400
      Top             =   4320
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Height          =   1935
      Left            =   7200
      TabIndex        =   8
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   975
      Left            =   4560
      TabIndex        =   7
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Height          =   975
      Left            =   1440
      TabIndex        =   6
      Top             =   3000
      Width           =   855
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00404000&
      BorderWidth     =   40
      X1              =   536
      X2              =   592
      Y1              =   320
      Y2              =   320
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00404000&
      BorderWidth     =   40
      X1              =   592
      X2              =   568
      Y1              =   312
      Y2              =   280
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00404000&
      BorderWidth     =   50
      X1              =   536
      X2              =   560
      Y1              =   312
      Y2              =   280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00404000&
      BorderWidth     =   25
      X1              =   504
      X2              =   616
      Y1              =   328
      Y2              =   328
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404000&
      BorderWidth     =   25
      X1              =   616
      X2              =   560
      Y1              =   328
      Y2              =   248
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404000&
      BorderWidth     =   25
      X1              =   504
      X2              =   560
      Y1              =   328
      Y2              =   248
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1110
      Left            =   9720
      TabIndex        =   5
      Top             =   7920
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   1110
      Left            =   6000
      TabIndex        =   4
      Top             =   6360
      Width           =   660
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1110
      Left            =   2760
      TabIndex        =   3
      Top             =   7560
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1110
      Left            =   9840
      TabIndex        =   2
      Top             =   360
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1110
      Left            =   5400
      TabIndex        =   1
      Top             =   1680
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1110
      Left            =   2040
      TabIndex        =   0
      Top             =   0
      Width           =   660
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   1440
      Top             =   3000
      Width           =   855
   End
   Begin VB.Menu mnuName 
      Caption         =   "&Change Name"
   End
End
Attribute VB_Name = "frmFindMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim uName$
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


Private Sub Form_Load()
Agent1.Characters.Load "Jacob", ("C:\WINDOWS\msagent\chars\merlin.acs")
Set Jake = Agent1.Characters("Jacob")
Jake.Balloon.Visible = True
Jake.Balloon.Style = 0
uName = GetSetting("Hidey", "User", "Name", "")
If uName = vbNullString Then
    uName = InputBox("What is your name?", "Name", "Merlin")
    SaveSetting "Hidey", "User", "Name", uName
End If
Jake.Show
Spell "Hi " & uName & " I want to play hide and seek with you."
Spell "I will go hide " & uName & "\pau=150\Try to find me!"
HideAway
End Sub

Private Sub Form_Unload(Cancel As Integer)
Jake.Show
Spell "Goodbye then " & uName & ". I'll play with you again."
Jake.Stop
End Sub

Private Sub Label1_Click()
If Jake.Left = Label1.Left And Jake.Top = Label1.Top Then
    Jake.Show
        Spell "Yaay..You found me! " & uName
        Spell "I was hiding behind the Yellow Letter Ayye!!!"
        Spell "Here I go!"
        HideAway
End If
End Sub

Private Sub Label2_Click()
If Jake.Left = Label2.Left And Jake.Top = Label2.Top Then
    Jake.Show
        Spell "Oh man..You found me!"
        Spell "I was hiding behind the White Letter Bee!!!"
        Spell "I'll hide better now!" & uName
        HideAway
End If
End Sub

Private Sub Label3_Click()
If Jake.Left = Label3.Left And Jake.Top = Label3.Top Then
    Jake.Show
        Spell "Here I am! " & uName
        Spell "behind the green Letter Sea!!!"
        Spell "Where will I hide next?"
        HideAway
End If
End Sub

Private Sub Label4_Click()
If Jake.Left = Label4.Left And Jake.Top = Label4.Top Then
    Jake.Show
        Spell "Can't fool you! " & uName
        Spell "I was hiding behind the red Letter Ex!!!"
        Spell "How quick can you find me now?!"
        HideAway
End If
End Sub

Private Sub Label5_Click()
If Jake.Left = Label5.Left And Jake.Top = Label5.Top Then
    Jake.Show
        Spell "Yep, " & uName & " You found me!"
        Spell "I was hiding behind the Pink Letter Why!!!"
        Spell "I'm hiding again!"
        HideAway
End If
End Sub

Private Sub Label6_Click()
If Jake.Left = Label6.Left And Jake.Top = Label6.Top Then
    Jake.Show
        Spell "Oh man..You found me! You are too good " & uName & "."
        Spell "I was hiding behind the Grey Letter Zee!!!"
        Spell "I'll go hide!"
        HideAway
End If
End Sub

Private Sub Label7_Click()
If Jake.Left = Shape1.Left And Jake.Top = Shape1.Top Then
    Jake.Show
        Spell "I was hiding behind the blue square!!!"
        Spell "You hide now, " & uName & " there you are at the computer!"
        Spell "I will go hide again!"
        HideAway
End If
End Sub

Private Sub Label8_Click()
If Jake.Left = Shape2.Left And Jake.Top = Shape2.Top Then
    Jake.Show
        Spell "You are so good " & uName & "!"
        Spell "I was hiding behind the Red Circle!!!"
        Spell "Find me now!"
        HideAway
End If
End Sub

Private Sub Label9_Click()
If Jake.Left = Label9.Left + 50 And Jake.Top = Label9.Top + 25 Then
    Jake.Show
        Spell "Yahoo, You found me, " & uName & "!"
        Spell "I was hiding behind the Blue Triangle!!!"
        Spell "Can't hide from you!"
        HideAway
End If
End Sub

Private Sub mnuName_Click()
    uName = InputBox("What is your name?", "Name", "Merlin")
    SaveSetting "Hidey", "User", "Name", uName
End Sub
