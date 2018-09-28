VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Building query structure..."
   ClientHeight    =   270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4275
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   270
   ScaleWidth      =   4275
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox folders_result 
      Height          =   4155
      Left            =   0
      TabIndex        =   8
      Top             =   4560
      Width           =   4455
   End
   Begin VB.ListBox files_result 
      Height          =   4155
      Left            =   4560
      TabIndex        =   7
      Top             =   4560
      Width           =   4455
   End
   Begin VB.ListBox extentions_result 
      Height          =   4155
      Left            =   9120
      TabIndex        =   6
      Top             =   4560
      Width           =   4455
   End
   Begin VB.ListBox filesizes_result 
      Height          =   4155
      Left            =   13680
      TabIndex        =   5
      Top             =   4560
      Width           =   4455
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3600
      Top             =   120
   End
   Begin VB.ListBox filesizes 
      Height          =   4155
      Left            =   13680
      TabIndex        =   4
      Top             =   360
      Width           =   4455
   End
   Begin VB.ListBox extentions 
      Height          =   4155
      Left            =   9120
      TabIndex        =   3
      Top             =   360
      Width           =   4455
   End
   Begin VB.ListBox files 
      Height          =   4155
      Left            =   4560
      TabIndex        =   2
      Top             =   360
      Width           =   4455
   End
   Begin VB.ListBox folders 
      Height          =   4155
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Hang on..."
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim fullname As String
Dim filename As String
Dim fileextention As String
Dim filesize As String
Dim enkelfolder As String


'remove previous query folder structure
Open "c:\temp\tmp.bat" For Output As #1
    Print #1, "rmdir c:\temp\query /S /Q"
Close #1
Shell ("c:\temp\tmp.bat"), vbHide
DoEvents

'grab file informations
teller = 0
Do
    fullname = Form1.List1.List(teller)
    filename = Right(fullname, (Len(fullname) - InStrRev(fullname, "\", , vbTextCompare)))
    fileextention = Right(filename, (Len(filename) - InStrRev(filename, ".", , vbTextCompare)))
    filename = Left(filename, (Len(filename) - 4))
    filesize = 0
    On Error Resume Next
    filesize = FileLen(fullname)

    enkelfolder = Replace(fullname, filename & "." & fileextention, "", 1, -1, vbTextCompare)
    
    folders.AddItem enkelfolder
    files.AddItem filename
    extentions.AddItem fileextention
    filesizes.AddItem filesize

teller = teller + 1
Loop Until teller = Form1.List1.ListCount

Form4.Hide
Form3.Show
End Sub

Private Sub Timer1_Timer()
Me.Hide
Timer1.Enabled = False
End Sub
