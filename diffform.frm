VERSION 5.00
Begin VB.Form diffform 
   BorderStyle     =   0  'None
   Caption         =   "Select Difficulty"
   ClientHeight    =   1935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2895
   Icon            =   "diffform.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton playbutt 
      Caption         =   "Play"
      Height          =   975
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.OptionButton diff 
      Caption         =   "Hard"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.OptionButton diff 
      Caption         =   "Normal"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton diff 
      Caption         =   "Easy"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label desbox 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Opponent considers all possible moves up to 2 moves deep."
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2655
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "diffform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private dif As Byte
Private selected As Boolean
Private Sub diff_Click(Index As Integer)
dif = Index
Select Case Index
Case 0: desbox.Caption = "Opponent picks the highest number regardless of where it leads."
Case 1: desbox.Caption = "Opponent considers all possible moves up to 2 moves deep."
Case 2: desbox.Caption = "Opponent considers all possible moves up to 4 moves deep."
End Select
End Sub

Private Sub Form_Terminate()
On Error Resume Next
Unload diffform
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Unload diffform
End Sub

Private Sub playbutt_Click()
selected = True
End Sub

Public Function GetLevel() As Byte
diffform.Show
selected = False
Select Case Current_Board.Player(1)
Case 1: diff(0).Value = True
Case 2: diff(1).Value = True
Case 4: diff(2).Value = True
End Select
playbutt.SetFocus
Do
    DoEvents
Loop Until selected = True
Select Case dif
Case 0: dif = 1
Case 1: dif = 2
Case 2: dif = 4
End Select
GetLevel = dif
diffform.Hide
End Function
