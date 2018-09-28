VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compare"
   ClientHeight    =   12465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15330
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12465
   ScaleWidth      =   15330
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   11760
      Left            =   7680
      TabIndex        =   1
      Top             =   480
      Width           =   7575
   End
   Begin VB.ListBox List1 
      Height          =   11760
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   7575
   End
   Begin VB.Label Label2 
      Caption         =   "Results"
      Height          =   255
      Left            =   7680
      TabIndex        =   3
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Originals:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    teller = 0
    Do
        List1.AddItem Form1.List1.List(teller)
        teller2 = 0
        found = 0
        Do
            If Form3.List1.List(teller2) = Form1.List1.List(teller) Then
                List2.AddItem Form3.List1.List(teller2)
                found = 1
            End If
            teller2 = teller2 + 1
        Loop Until teller2 = Form3.List1.ListCount
        If found = 0 Then
            List2.AddItem "------------------------------------------------------------------------"
        End If
        teller = teller + 1
    Loop Until teller = Form1.List1.ListCount
End Sub

Private Sub List1_Click()
List2.Selected(List1.ListIndex) = True
End Sub

Private Sub List2_Click()
List1.Selected(List2.ListIndex) = True
End Sub
