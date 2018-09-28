VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Searching..."
   ClientHeight    =   495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3840
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3360
      Top             =   0
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   1680
      Top             =   120
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   -120
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Shape2.Left = Shape1.Left
Shape2.Width = 1
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
If Shape2.Width < Shape1.Width Then
    Shape2.Width = Shape2.Width + 50
Else
    Timer2.Enabled = True
    Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
If Shape2.Width > 51 Then
    Shape2.Width = Shape2.Width - 50
    Shape2.Left = Shape2.Left + 50
Else
    Shape2.Left = Shape1.Left
    Shape2.Width = 1
    Timer1.Enabled = True
    Timer2.Enabled = False
End If
End Sub
