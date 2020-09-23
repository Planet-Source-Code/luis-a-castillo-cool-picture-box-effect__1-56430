VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   3465
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Shape1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   600
      ScaleHeight     =   270
      ScaleWidth      =   870
      TabIndex        =   1
      Top             =   480
      Width           =   900
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   2040
      Top             =   720
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2040
      Top             =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make Copies"
      Height          =   615
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2040
      Top             =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public n As Integer


Private Sub Command1_Click()
    For X = 1 To 5
        Load Shape1(X)
            Shape1(X).Left = Shape1(X - 1).Left
            Shape1(X).Top = Shape1(X - 1).Top + 500
            Shape1(X).Visible = True
            Command1.Enabled = False
    Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False
Timer2.Enabled = True
End Sub

Private Sub Shape1_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = True
Timer2.Enabled = False
n = index
End Sub

Private Sub Timer1_Timer()
For i = 1 To 10
If Shape1(n).Height < 1000 Or Shape1(n).Width < 1000 Then
Shape1(n).Left = Shape1(n).Left - (i / 2)
Shape1(n).Top = Shape1(n).Top - (i / 2)
Shape1(n).Height = Shape1(n).Height + i
Shape1(n).Width = Shape1(n).Width + i
Else
Timer1.Enabled = False
End If
Next
End Sub

Private Sub Timer2_Timer()
For i = 1 To 10
If Shape1(n).Height > 300 Or Shape1(n).Width > 900 Then
Shape1(n).Left = Shape1(n).Left + (i / 2)
Shape1(n).Top = Shape1(n).Top + (i / 2)
Shape1(n).Height = Shape1(n).Height - i
Shape1(n).Width = Shape1(n).Width - i
Else
Timer2.Enabled = False
End If
Next
End Sub


Private Sub Timer3_Timer()
For i = 1 To 10
For sh = 1 To Shape1.Count
If n <> (sh - 1) Then
If Shape1(sh - 1).Height > 300 Or Shape1(sh - 1).Width > 900 Then
Shape1(sh - 1).Left = Shape1(sh - 1).Left + (i / 2)
Shape1(sh - 1).Top = Shape1(sh - 1).Top + (i / 2)
Shape1(sh - 1).Height = Shape1(sh - 1).Height - i
Shape1(sh - 1).Width = Shape1(sh - 1).Width - i
End If
End If
Next
Next
End Sub
