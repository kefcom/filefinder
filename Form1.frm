VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Find"
   ClientHeight    =   10845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7410
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10845
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "query"
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      Top             =   10440
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Export to file"
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   10440
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   5715
      Left            =   0
      TabIndex        =   7
      Top             =   4680
      Width           =   7335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find!"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   4200
      Width           =   7455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   3840
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   "*.jpg *.gif"
      Top             =   3480
      Width           =   6375
   End
   Begin VB.DirListBox Dir1 
      Height          =   3015
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   7335
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   10440
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "find in file or folder name:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Extention(s):"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   3480
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

On Error Resume Next
Kill "c:\temp\zoek.bat"
On Error Resume Next
Kill "c:\temp\output.txt"
On Error Resume Next
Kill "c:\temp\done.txt"

Open ("c:\temp\zoek.bat") For Output As #1
    Print #1, "cd /d " & Dir1.Path
    Print #1, "dir /b /s " & Text1.Text & " >> c:\temp\output.txt"
    Print #1, "dir >> c:\temp\done.txt"
Close #1

Form2.Show

Shell ("c:\temp\zoek.bat"), vbHide

Do
    On Error Resume Next
    Open ("c:\temp\done.txt") For Input As #1
    
    On Error Resume Next
    Input #1, testja
    
    If testja <> "" Then
        Close #1
        GoTo endloop
    End If
    
    On Error Resume Next
    Close #1
    
    DoEvents
Loop

endloop:

'file openen en inlezen wat nodig is
Open ("c:\temp\output.txt") For Input As #1
    Do Until EOF(1) = True
        Input #1, lijntje
        If Len(Text2.Text) > 0 Then
            If InStr(1, lijntje, Text2.Text, vbTextCompare) > 0 Then
                List1.AddItem lijntje
            End If
        Else
            List1.AddItem lijntje
        End If
    Loop
Close #1

'zoekvenster sluiten
Form2.Visible = False

'files wissen
Kill "c:\temp\output.txt"
Kill "c:\temp\done.txt"
Kill "c:\temp\zoek.bat"
    
Label3.Caption = "" & List1.ListCount & " results"
    

End Sub

Private Sub Command2_Click()
vraag = InputBox("file path+name?", "where to save?", "c:\temp\list.txt")
If Len(vraag) > 0 Then
    Open vraag For Output As #1
        teller = 0
        Do
            Print #1, List1.List(teller)
            teller = teller + 1
        Loop Until teller = List1.ListCount
    Close #1
    MsgBox "saved!", vbInformation
End If
End Sub

Private Sub Command3_Click()
Form4.Show
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

