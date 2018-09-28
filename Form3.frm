VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Query"
   ClientHeight    =   10275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14370
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10275
   ScaleWidth      =   14370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "compare with original"
      Height          =   495
      Left            =   2760
      TabIndex        =   17
      Top             =   9720
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear queries"
      Height          =   495
      Left            =   13080
      TabIndex        =   15
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Export list to file"
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   9720
      Width           =   2655
   End
   Begin VB.ListBox List1 
      Height          =   5130
      ItemData        =   "Form3.frx":6852
      Left            =   0
      List            =   "Form3.frx":6854
      TabIndex        =   9
      Top             =   4560
      Width           =   14295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Run queries"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   2655
   End
   Begin VB.ListBox queries 
      Height          =   2010
      Left            =   0
      TabIndex        =   7
      Top             =   1920
      Width           =   14295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add query"
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   14295
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "Form3.frx":6856
         Left            =   960
         List            =   "Form3.frx":6860
         TabIndex        =   11
         Text            =   "..."
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add query"
         Height          =   255
         Left            =   12600
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   6480
         TabIndex        =   5
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Form3.frx":6878
         Left            =   4920
         List            =   "Form3.frx":688E
         TabIndex        =   4
         Text            =   "..."
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form3.frx":68BA
         Left            =   2640
         List            =   "Form3.frx":68CD
         TabIndex        =   3
         Text            =   "..."
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "File size is expressed in bytes"
         Height          =   255
         Left            =   9000
         TabIndex        =   16
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "/!\ warning, this query may be resource intensive!"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   9000
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "WHERE"
         Height          =   255
         Left            =   1920
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Select from "
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label lbl_results 
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
      Left            =   12360
      TabIndex        =   10
      Top             =   9720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Query builder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   14175
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "Form3.frx":6906
      Top             =   120
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   14415
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
If Combo1.Text = "Filecontent" Then
    Label5.Visible = True
    Exit Sub
End If
End Sub

Private Sub Combo1_Click()
If Combo1.Text = "Filecontent" Then
    Label5.Visible = True
    Exit Sub
End If
End Sub

Private Sub Command1_Click()
'check for false input
If Combo1.Text = "..." Then
    MsgBox "please fill in all the values.", vbInformation, "Error"
    Exit Sub
End If
If Combo2.Text = "..." Then
    MsgBox "please fill in all the values.", vbInformation, "Error"
    Exit Sub
End If
If Combo3.Text = "..." Then
    MsgBox "please fill in all the values.", vbInformation, "Error"
    Exit Sub
End If


Label5.Visible = False


queries.AddItem "SELECT from '" & Combo3.Text & "' WHERE " & Combo1.Text & " " & Combo2.Text & " '" & Text1.Text & "'"
End Sub

Private Sub Command2_Click()
Form6.Show
DoEvents

DoEvents
DoEvents
DoEvents
DoEvents
DoEvents
DoEvents

Dim lijn As String
teller_query = 0
Do
    querystring = Replace(queries.List(teller_query), "SELECT from '", "", 1, -1, vbTextCompare)
    q0 = Left(querystring, (InStr(1, querystring, "'", vbTextCompare) - 1))
    
    
    querystring = Right(querystring, Len(querystring) - InStr(1, querystring, "WHERE", vbTextCompare))
    querystring = Right(querystring, Len(querystring) - InStr(1, querystring, " ", vbTextCompare))
    
    q1 = Left(querystring, (InStr(1, querystring, " ", vbTextCompare) - 1))
    
    querystring = Right(querystring, Len(querystring) - InStr(1, querystring, " ", vbTextCompare))
    q2 = Left(querystring, (InStr(1, querystring, "'", vbTextCompare) - 2))
    
    querystring = Right(querystring, Len(querystring) - InStr(1, querystring, "'", vbTextCompare))
    q3 = Left(querystring, (InStr(1, querystring, "'", vbTextCompare) - 1))


    'Q1
    'filename
    'Dirname
    'Filesize
    'Extention
    'Filecontent
''''''''''''''''
    'Q2
    'IS
    'IS NOT
    'IS >
    'IS <
    'LIKE
    'NOT LIKE
''''''''''''''''


    Select Case q0
    Case "Originals"
    
    Select Case q1
        Case "Filename"
            Select Case q2
                Case "IS"
                    teller = 0
                    Do
                        If Form4.files.List(teller) = q3 Then
                            Form4.folders_result.AddItem Form4.folders.List(teller)
                            Form4.files_result.AddItem Form4.files.List(teller)
                            Form4.extentions_result.AddItem Form4.extentions.List(teller)
                            Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                        End If
                        teller = teller + 1
                    Loop Until teller >= Form4.files.ListCount
                    
                    GoTo nextquery
                
                Case "IS NOT"
                    teller = 0
                    Do
                        If Form4.files.List(teller) <> q3 Then
                            Form4.folders_result.AddItem Form4.folders.List(teller)
                            Form4.files_result.AddItem Form4.files.List(teller)
                            Form4.extentions_result.AddItem Form4.extentions.List(teller)
                            Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                        End If
                        teller = teller + 1
                    Loop Until teller >= Form4.files.ListCount
                    GoTo nextquery
                    
                Case "IS >"
                    teller = 0
                    Do
                        If Form4.files.List(teller) > q3 Then
                            Form4.folders_result.AddItem Form4.folders.List(teller)
                            Form4.files_result.AddItem Form4.files.List(teller)
                            Form4.extentions_result.AddItem Form4.extentions.List(teller)
                            Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                        End If
                        teller = teller + 1
                    Loop Until teller >= Form4.files.ListCount
                    GoTo nextquery
                    
                Case "IS <"
                    teller = 0
                    Do
                        If Form4.files.List(teller) < q3 Then
                            Form4.folders_result.AddItem Form4.folders.List(teller)
                            Form4.files_result.AddItem Form4.files.List(teller)
                            Form4.extentions_result.AddItem Form4.extentions.List(teller)
                            Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                        End If
                        teller = teller + 1
                    Loop Until teller >= Form4.files.ListCount
                    GoTo nextquery
                                                
                Case "LIKE"
                    teller = 0
                    Do
                        If InStr(1, Form4.files.List(teller), q3, vbTextCompare) > 0 Then
                            Form4.folders_result.AddItem Form4.folders.List(teller)
                            Form4.files_result.AddItem Form4.files.List(teller)
                            Form4.extentions_result.AddItem Form4.extentions.List(teller)
                            Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                        End If
                        teller = teller + 1
                    Loop Until teller >= Form4.files.ListCount
                    GoTo nextquery
                    
                Case "NOT LIKE"
                    teller = 0
                    Do
                        If InStr(1, Form4.files.List(teller), q3, vbTextCompare) = 0 Then
                            Form4.folders_result.AddItem Form4.folders.List(teller)
                            Form4.files_result.AddItem Form4.files.List(teller)
                            Form4.extentions_result.AddItem Form4.extentions.List(teller)
                            Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                        End If
                        teller = teller + 1
                    Loop Until teller >= Form4.files.ListCount
                    GoTo nextquery
            End Select
            
            
            
            
        Case "Dirname"
            Select Case q2
                Case "IS"
                    teller = 0
                    Do
                        If Form4.folders.List(teller) = q3 Then
                            Form4.folders_result.AddItem Form4.folders.List(teller)
                            Form4.files_result.AddItem Form4.files.List(teller)
                            Form4.extentions_result.AddItem Form4.extentions.List(teller)
                            Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                        End If
                        teller = teller + 1
                    Loop Until teller >= Form4.folders.ListCount
                    
                    GoTo nextquery
                
                Case "IS NOT"
                    teller = 0
                    Do
                        If Form4.folders.List(teller) <> q3 Then
                            Form4.folders_result.AddItem Form4.folders.List(teller)
                            Form4.files_result.AddItem Form4.files.List(teller)
                            Form4.extentions_result.AddItem Form4.extentions.List(teller)
                            Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                        End If
                        teller = teller + 1
                    Loop Until teller >= Form4.folders.ListCount
                    GoTo nextquery
                    
                Case "IS >"
                    teller = 0
                    Do
                        If Form4.folders.List(teller) > q3 Then
                            Form4.folders_result.AddItem Form4.folders.List(teller)
                            Form4.files_result.AddItem Form4.files.List(teller)
                            Form4.extentions_result.AddItem Form4.extentions.List(teller)
                            Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                        End If
                        teller = teller + 1
                    Loop Until teller >= Form4.folders.ListCount
                    GoTo nextquery
                    
                Case "IS <"
                    teller = 0
                    Do
                        If Form4.folders.List(teller) < q3 Then
                            Form4.folders_result.AddItem Form4.folders.List(teller)
                            Form4.files_result.AddItem Form4.files.List(teller)
                            Form4.extentions_result.AddItem Form4.extentions.List(teller)
                            Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                        End If
                        teller = teller + 1
                    Loop Until teller >= Form4.folders.ListCount
                    GoTo nextquery
                                                
                Case "LIKE"
                    teller = 0
                    Do
                        If InStr(1, Form4.folders.List(teller), q3, vbTextCompare) > 0 Then
                            Form4.folders_result.AddItem Form4.folders.List(teller)
                            Form4.files_result.AddItem Form4.files.List(teller)
                            Form4.extentions_result.AddItem Form4.extentions.List(teller)
                            Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                        End If
                        teller = teller + 1
                    Loop Until teller >= Form4.folders.ListCount
                    GoTo nextquery
                    
                Case "NOT LIKE"
                    teller = 0
                    Do
                        If InStr(1, Form4.folders.List(teller), q3, vbTextCompare) = 0 Then
                            Form4.folders_result.AddItem Form4.folders.List(teller)
                            Form4.files_result.AddItem Form4.files.List(teller)
                            Form4.extentions_result.AddItem Form4.extentions.List(teller)
                            Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                        End If
                        teller = teller + 1
                    Loop Until teller >= Form4.folders.ListCount
                    GoTo nextquery
            End Select
            
            
        
        
        
        
        Case "Filesize"
            Select Case q2
                Case "IS"
                    teller = 0
                    Do
                        If Form4.filesizes.List(teller) = q3 Then
                            Form4.folders_result.AddItem Form4.folders.List(teller)
                            Form4.files_result.AddItem Form4.files.List(teller)
                            Form4.extentions_result.AddItem Form4.extentions.List(teller)
                            Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                        End If
                        teller = teller + 1
                    Loop Until teller >= Form4.filesizes.ListCount
                    
                    GoTo nextquery
                
                Case "IS NOT"
                    teller = 0
                    Do
                        If Form4.filesizes.List(teller) <> q3 Then
                            Form4.folders_result.AddItem Form4.folders.List(teller)
                            Form4.files_result.AddItem Form4.files.List(teller)
                            Form4.extentions_result.AddItem Form4.extentions.List(teller)
                            Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                        End If
                        teller = teller + 1
                    Loop Until teller >= Form4.filesizes.ListCount
                    GoTo nextquery
                    
                Case "IS >"
                    teller = 0
                    Do
                        If Form4.filesizes.List(teller) > q3 Then
                            Form4.folders_result.AddItem Form4.folders.List(teller)
                            Form4.files_result.AddItem Form4.files.List(teller)
                            Form4.extentions_result.AddItem Form4.extentions.List(teller)
                            Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                        End If
                        teller = teller + 1
                    Loop Until teller >= Form4.filesizes.ListCount
                    GoTo nextquery
                    
                Case "IS <"
                    teller = 0
                    Do
                        If Form4.filesizes.List(teller) < q3 Then
                            Form4.folders_result.AddItem Form4.folders.List(teller)
                            Form4.files_result.AddItem Form4.files.List(teller)
                            Form4.extentions_result.AddItem Form4.extentions.List(teller)
                            Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                        End If
                        teller = teller + 1
                    Loop Until teller >= Form4.filesizes.ListCount
                    GoTo nextquery
                                                
                Case "LIKE"
                    teller = 0
                    Do
                        If InStr(1, Form4.filesizes.List(teller), q3, vbTextCompare) > 0 Then
                            Form4.folders_result.AddItem Form4.folders.List(teller)
                            Form4.files_result.AddItem Form4.files.List(teller)
                            Form4.extentions_result.AddItem Form4.extentions.List(teller)
                            Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                        End If
                        teller = teller + 1
                    Loop Until teller >= Form4.filesizes.ListCount
                    GoTo nextquery
                    
                Case "NOT LIKE"
                    teller = 0
                    Do
                        If InStr(1, Form4.filesizes.List(teller), q3, vbTextCompare) = 0 Then
                            Form4.folders_result.AddItem Form4.folders.List(teller)
                            Form4.files_result.AddItem Form4.files.List(teller)
                            Form4.extentions_result.AddItem Form4.extentions.List(teller)
                            Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                        End If
                        teller = teller + 1
                    Loop Until teller >= Form4.filesizes.ListCount
                    GoTo nextquery
            End Select
            
            
            
            
        Case "Extention"
            Select Case q2
                Case "IS"
                    teller = 0
                    Do
                        If Form4.extentions.List(teller) = q3 Then
                            Form4.folders_result.AddItem Form4.folders.List(teller)
                            Form4.files_result.AddItem Form4.files.List(teller)
                            Form4.extentions_result.AddItem Form4.extentions.List(teller)
                            Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                        End If
                        teller = teller + 1
                    Loop Until teller >= Form4.extentions.ListCount
                    
                    GoTo nextquery
                
                Case "IS NOT"
                    teller = 0
                    Do
                        If Form4.extentions.List(teller) <> q3 Then
                            Form4.folders_result.AddItem Form4.folders.List(teller)
                            Form4.files_result.AddItem Form4.files.List(teller)
                            Form4.extentions_result.AddItem Form4.extentions.List(teller)
                            Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                        End If
                        teller = teller + 1
                    Loop Until teller >= Form4.extentions.ListCount
                    GoTo nextquery
                    
                Case "IS >"
                    teller = 0
                    Do
                        If Form4.extentions.List(teller) > q3 Then
                            Form4.folders_result.AddItem Form4.folders.List(teller)
                            Form4.files_result.AddItem Form4.files.List(teller)
                            Form4.extentions_result.AddItem Form4.extentions.List(teller)
                            Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                        End If
                        teller = teller + 1
                    Loop Until teller >= Form4.extentions.ListCount
                    GoTo nextquery
                    
                Case "IS <"
                    teller = 0
                    Do
                        If Form4.extentions.List(teller) < q3 Then
                            Form4.folders_result.AddItem Form4.folders.List(teller)
                            Form4.files_result.AddItem Form4.files.List(teller)
                            Form4.extentions_result.AddItem Form4.extentions.List(teller)
                            Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                        End If
                        teller = teller + 1
                    Loop Until teller >= Form4.extentions.ListCount
                    GoTo nextquery
                                                
                Case "LIKE"
                    teller = 0
                    Do
                        If InStr(1, Form4.extentions.List(teller), q3, vbTextCompare) > 0 Then
                            Form4.folders_result.AddItem Form4.folders.List(teller)
                            Form4.files_result.AddItem Form4.files.List(teller)
                            Form4.extentions_result.AddItem Form4.extentions.List(teller)
                            Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                        End If
                        teller = teller + 1
                    Loop Until teller >= Form4.extentions.ListCount
                    GoTo nextquery
                    
                Case "NOT LIKE"
                    teller = 0
                    Do
                        If InStr(1, Form4.extentions.List(teller), q3, vbTextCompare) = 0 Then
                            Form4.folders_result.AddItem Form4.folders.List(teller)
                            Form4.files_result.AddItem Form4.files.List(teller)
                            Form4.extentions_result.AddItem Form4.extentions.List(teller)
                            Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                        End If
                        teller = teller + 1
                    Loop Until teller >= Form4.extentions.ListCount
                    GoTo nextquery
                End Select
            
            
            
            
            
            Case "Filecontent"
            Select Case q2
                Case "IS"
                    teller = 0
                    Do
                        On Error Resume Next
                        Open (Form4.folders.List(teller) & Form4.files.List(teller) & "." & Form4.extentions.List(teller)) For Input As #1
                            On Error Resume Next
                            Input #1, lijn
                                If lijn = q3 Then
                                    Form4.folders_result.AddItem Form4.folders.List(teller)
                                    Form4.files_result.AddItem Form4.files.List(teller)
                                    Form4.extentions_result.AddItem Form4.extentions.List(teller)
                                    Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                                End If
                        Close #1
                        teller = teller + 1
                    Loop Until teller >= Form4.files.ListCount
                    
                    GoTo nextquery
                
                Case "IS NOT"
                    teller = 0
                    Do
                        On Error Resume Next
                        Open (Form4.folders.List(teller) & Form4.files.List(teller) & "." & Form4.extentions.List(teller)) For Input As #1
                        On Error Resume Next
                            Input #1, lijn
                                If lijn <> q3 Then
                                    Form4.folders_result.AddItem Form4.folders.List(teller)
                                    Form4.files_result.AddItem Form4.files.List(teller)
                                    Form4.extentions_result.AddItem Form4.extentions.List(teller)
                                    Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                                End If
                        Close #1
                        teller = teller + 1
                    Loop Until teller >= Form4.files.ListCount
                    GoTo nextquery
                    
                Case "IS >"
                    teller = 0
                    Do
                        On Error Resume Next
                        Open (Form4.folders.List(teller) & Form4.files.List(teller) & "." & Form4.extentions.List(teller)) For Input As #1
                        On Error Resume Next
                            Input #1, lijn
                                If lijn > q3 Then
                                    Form4.folders_result.AddItem Form4.folders.List(teller)
                                    Form4.files_result.AddItem Form4.files.List(teller)
                                    Form4.extentions_result.AddItem Form4.extentions.List(teller)
                                    Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                                End If
                        Close #1
                        teller = teller + 1
                    Loop Until teller >= Form4.files.ListCount
                    GoTo nextquery
                    
                Case "IS <"
                    teller = 0
                    Do
                        On Error Resume Next
                        Open (Form4.folders.List(teller) & Form4.files.List(teller) & "." & Form4.extentions.List(teller)) For Input As #1
                        On Error Resume Next
                            Input #1, lijn
                                If lijn < q3 Then
                                    Form4.folders_result.AddItem Form4.folders.List(teller)
                                    Form4.files_result.AddItem Form4.files.List(teller)
                                    Form4.extentions_result.AddItem Form4.extentions.List(teller)
                                    Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                                End If
                        Close #1
                        teller = teller + 1
                    Loop Until teller >= Form4.files.ListCount
                    GoTo nextquery
                                                
                Case "LIKE"
                    teller = 0
                    Do
                        On Error Resume Next
                        Open (Form4.folders.List(teller) & Form4.files.List(teller) & "." & Form4.extentions.List(teller)) For Input As #1
                        Do
                        On Error Resume Next
                            Input #1, lijn
                                If InStr(1, lijn, q3, vbTextCompare) > 0 Then
                                    Form4.folders_result.AddItem Form4.folders.List(teller)
                                    Form4.files_result.AddItem Form4.files.List(teller)
                                    Form4.extentions_result.AddItem Form4.extentions.List(teller)
                                    Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                                End If
                        Loop Until EOF(1)
                        Close #1
                        teller = teller + 1
                    Loop Until teller >= Form4.files.ListCount
                    GoTo nextquery
                    
                Case "NOT LIKE"
                    teller = 0
                    Do
                        On Error Resume Next
                        Open (Form4.folders.List(teller) & Form4.files.List(teller) & "." & Form4.extentions.List(teller)) For Input As #1
                        Do
                        On Error Resume Next
                            Input #1, lijn
                                If InStr(1, lijn, q3, vbTextCompare) = 0 Then
                                    Form4.folders_result.AddItem Form4.folders.List(teller)
                                    Form4.files_result.AddItem Form4.files.List(teller)
                                    Form4.extentions_result.AddItem Form4.extentions.List(teller)
                                    Form4.filesizes_result.AddItem Form4.filesizes.List(teller)
                                End If
                        Loop Until EOF(1)
                        Close #1
                        teller = teller + 1
                    Loop Until teller >= Form4.files.ListCount
                    GoTo nextquery
                End Select
            
            
            
            
            
            
   
    
    
    
    
    
    
    
    End Select
    
    
Case Else






    Select Case q1
        Case "Filename"
            Select Case q2
                Case "IS"
                    teller = 0
                    Do
                        If Form4.files_result.List(teller) = q3 Then
                            'do nothing
                        teller = teller + 1
                        Else
                            Form4.folders_result.RemoveItem (teller)
                            Form4.files_result.RemoveItem (teller)
                            Form4.extentions_result.RemoveItem (teller)
                            Form4.filesizes_result.RemoveItem (teller)
                        End If
                    Loop Until teller >= Form4.files_result.ListCount
                    
                    GoTo nextquery
                
                Case "IS NOT"
                    teller = 0
                    Do
                        If Form4.files_result.List(teller) <> q3 Then
                            'do nothing
                        teller = teller + 1
                        Else
                            Form4.folders_result.RemoveItem (teller)
                            Form4.files_result.RemoveItem (teller)
                            Form4.extentions_result.RemoveItem (teller)
                            Form4.filesizes_result.RemoveItem (teller)
                        End If
                    Loop Until teller >= Form4.files_result.ListCount
                    GoTo nextquery
                    
                Case "IS >"
                    teller = 0
                    Do
                        If Form4.files_result.List(teller) > q3 Then
                            'do nothing
                            teller = teller + 1
                        Else
                            Form4.folders_result.RemoveItem (teller)
                            Form4.files_result.RemoveItem (teller)
                            Form4.extentions_result.RemoveItem (teller)
                            Form4.filesizes_result.RemoveItem (teller)
                        End If
                    Loop Until teller >= Form4.files_result.ListCount
                    GoTo nextquery
                    
                Case "IS <"
                    teller = 0
                    Do
                        If Form4.files_result.List(teller) < q3 Then
                            'do nothing
                        teller = teller + 1
                        Else
                            Form4.folders_result.RemoveItem (teller)
                            Form4.files_result.RemoveItem (teller)
                            Form4.extentions_result.RemoveItem (teller)
                            Form4.filesizes_result.RemoveItem (teller)
                        End If
                    Loop Until teller >= Form4.files_result.ListCount
                    GoTo nextquery
                                                
                Case "LIKE"
                    teller = 0
                    Do
                        If InStr(1, Form4.files_result.List(teller), q3, vbTextCompare) > 0 Then
                            'do nothing
                        teller = teller + 1
                        Else
                            Form4.folders_result.RemoveItem (teller)
                            Form4.files_result.RemoveItem (teller)
                            Form4.extentions_result.RemoveItem (teller)
                            Form4.filesizes_result.RemoveItem (teller)
                        End If
                    Loop Until teller >= Form4.files_result.ListCount
                    GoTo nextquery
                    
                Case "NOT LIKE"
                    teller = 0
                    Do
                        If InStr(1, Form4.files_result.List(teller), q3, vbTextCompare) = 0 Then
                            'do nothing
                        teller = teller + 1
                        Else
                            Form4.folders_result.RemoveItem (teller)
                            Form4.files_result.RemoveItem (teller)
                            Form4.extentions_result.RemoveItem (teller)
                            Form4.filesizes_result.RemoveItem (teller)
                        End If
                    Loop Until teller >= Form4.files_result.ListCount
                    GoTo nextquery
            End Select
            
            
            
            
        Case "Dirname"
            Select Case q2
                Case "IS"
                    teller = 0
                    Do
                        If Form4.folders_result.List(teller) = q3 Then
                            'do nothing
                        teller = teller + 1
                        Else
                            Form4.folders_result.RemoveItem (teller)
                            Form4.files_result.RemoveItem (teller)
                            Form4.extentions_result.RemoveItem (teller)
                            Form4.filesizes_result.RemoveItem (teller)
                        End If
                    Loop Until teller >= Form4.folders_result.ListCount
                    
                    GoTo nextquery
                
                Case "IS NOT"
                    teller = 0
                    Do
                        If Form4.folders_result.List(teller) <> q3 Then
                            'do nothing
                        teller = teller + 1
                        Else
                            Form4.folders_result.RemoveItem (teller)
                            Form4.files_result.RemoveItem (teller)
                            Form4.extentions_result.RemoveItem (teller)
                            Form4.filesizes_result.RemoveItem (teller)
                        End If
                    Loop Until teller >= Form4.folders_result.ListCount
                    GoTo nextquery
                    
                Case "IS >"
                    teller = 0
                    Do
                        If Form4.folders_result.List(teller) > q3 Then
                            'do nothing
                            teller = teller + 1
                        Else
                            Form4.folders_result.RemoveItem (teller)
                            Form4.files_result.RemoveItem (teller)
                            Form4.extentions_result.RemoveItem (teller)
                            Form4.filesizes_result.RemoveItem (teller)
                        End If
                    Loop Until teller >= Form4.folders_result.ListCount
                    GoTo nextquery
                    
                Case "IS <"
                    teller = 0
                    Do
                        If Form4.folders_result.List(teller) < q3 Then
                            'do nothing
                        teller = teller + 1
                        Else
                            Form4.folders_result.RemoveItem (teller)
                            Form4.files_result.RemoveItem (teller)
                            Form4.extentions_result.RemoveItem (teller)
                            Form4.filesizes_result.RemoveItem (teller)
                        End If
                    Loop Until teller >= Form4.folders_result.ListCount
                    GoTo nextquery
                                                
                Case "LIKE"
                    teller = 0
                    Do
                        If InStr(1, Form4.folders_result.List(teller), q3, vbTextCompare) > 0 Then
                            'do nothing
                        teller = teller + 1
                        Else
                            Form4.folders_result.RemoveItem (teller)
                            Form4.files_result.RemoveItem (teller)
                            Form4.extentions_result.RemoveItem (teller)
                            Form4.filesizes_result.RemoveItem (teller)
                        End If
                    Loop Until teller >= Form4.folders_result.ListCount
                    GoTo nextquery
                    
                Case "NOT LIKE"
                    teller = 0
                    Do
                        If InStr(1, Form4.folders_result.List(teller), q3, vbTextCompare) = 0 Then
                            'do nothing
                        teller = teller + 1
                        Else
                            Form4.folders_result.RemoveItem (teller)
                            Form4.files_result.RemoveItem (teller)
                            Form4.extentions_result.RemoveItem (teller)
                            Form4.filesizes_result.RemoveItem (teller)
                        End If
                    Loop Until teller >= Form4.folders_result.ListCount
                    GoTo nextquery
            End Select
            
            
        
        
        
        
        Case "Filesize"
            Select Case q2
                Case "IS"
                    teller = 0
                    Do
                        If Form4.filesizes_result.List(teller) = q3 Then
                            'do nothing
                        teller = teller + 1
                        Else
                            Form4.folders_result.RemoveItem (teller)
                            Form4.files_result.RemoveItem (teller)
                            Form4.extentions_result.RemoveItem (teller)
                            Form4.filesizes_result.RemoveItem (teller)
                        End If
                    Loop Until teller >= Form4.filesizes_result.ListCount
                    
                    GoTo nextquery
                
                Case "IS NOT"
                    teller = 0
                    Do
                        If Form4.filesizes_result.List(teller) <> q3 Then
                            'do nothing
                        teller = teller + 1
                        Else
                            Form4.folders_result.RemoveItem (teller)
                            Form4.files_result.RemoveItem (teller)
                            Form4.extentions_result.RemoveItem (teller)
                            Form4.filesizes_result.RemoveItem (teller)
                        End If
                    Loop Until teller >= Form4.filesizes_result.ListCount
                    GoTo nextquery
                    
                Case "IS >"
                    teller = 0
                    Do
                        If Form4.filesizes_result.List(teller) > q3 Then
                            'do nothing
                        teller = teller + 1
                        Else
                            Form4.folders_result.RemoveItem (teller)
                            Form4.files_result.RemoveItem (teller)
                            Form4.extentions_result.RemoveItem (teller)
                            Form4.filesizes_result.RemoveItem (teller)
                        End If
                    Loop Until teller >= Form4.filesizes_result.ListCount
                    GoTo nextquery
                    
                Case "IS <"
                    teller = 0
                    Do
                        If Form4.filesizes_result.List(teller) < q3 Then
                            'do nothing
                        teller = teller + 1
                        Else
                            Form4.folders_result.RemoveItem (teller)
                            Form4.files_result.RemoveItem (teller)
                            Form4.extentions_result.RemoveItem (teller)
                            Form4.filesizes_result.RemoveItem (teller)
                        End If
                    Loop Until teller >= Form4.filesizes_result.ListCount
                    GoTo nextquery
                                                
                Case "LIKE"
                    teller = 0
                    Do
                        If InStr(1, Form4.filesizes_result.List(teller), q3, vbTextCompare) > 0 Then
                            'do nothing
                        teller = teller + 1
                        Else
                            Form4.folders_result.RemoveItem (teller)
                            Form4.files_result.RemoveItem (teller)
                            Form4.extentions_result.RemoveItem (teller)
                            Form4.filesizes_result.RemoveItem (teller)
                        End If
                    Loop Until teller >= Form4.filesizes_result.ListCount
                    GoTo nextquery
                    
                Case "NOT LIKE"
                    teller = 0
                    Do
                        If InStr(1, Form4.filesizes_result.List(teller), q3, vbTextCompare) = 0 Then
                            'do nothing
                        teller = teller + 1
                        Else
                            Form4.folders_result.RemoveItem (teller)
                            Form4.files_result.RemoveItem (teller)
                            Form4.extentions_result.RemoveItem (teller)
                            Form4.filesizes_result.RemoveItem (teller)
                        End If
                    Loop Until teller >= Form4.filesizes_result.ListCount
                    GoTo nextquery
            End Select
            
            
            
            
        Case "Extention"
            Select Case q2
                Case "IS"
                    teller = 0
                    Do
                        If Form4.extentions_result.List(teller) = q3 Then
                            'do nothing
                        teller = teller + 1
                        Else
                            Form4.folders_result.RemoveItem (teller)
                            Form4.files_result.RemoveItem (teller)
                            Form4.extentions_result.RemoveItem (teller)
                            Form4.filesizes_result.RemoveItem (teller)
                        End If
                    Loop Until teller >= Form4.extentions_result.ListCount
                    
                    GoTo nextquery
                
                Case "IS NOT"
                    teller = 0
                    Do
                        If Form4.extentions_result.List(teller) <> q3 Then
                            'do nothing
                        teller = teller + 1
                        Else
                            Form4.folders_result.RemoveItem (teller)
                            Form4.files_result.RemoveItem (teller)
                            Form4.extentions_result.RemoveItem (teller)
                            Form4.filesizes_result.RemoveItem (teller)
                        End If
                    Loop Until teller >= Form4.extentions_result.ListCount
                    GoTo nextquery
                    
                Case "IS >"
                    teller = 0
                    Do
                        If Form4.extentions_result.List(teller) > q3 Then
                            'do nothing
                        teller = teller + 1
                        Else
                            Form4.folders_result.RemoveItem (teller)
                            Form4.files_result.RemoveItem (teller)
                            Form4.extentions_result.RemoveItem (teller)
                            Form4.filesizes_result.RemoveItem (teller)
                        End If
                    Loop Until teller >= Form4.extentions_result.ListCount
                    GoTo nextquery
                    
                Case "IS <"
                    teller = 0
                    Do
                        If Form4.extentions_result.List(teller) < q3 Then
                            'do nothing
                        teller = teller + 1
                        Else
                            Form4.folders_result.RemoveItem (teller)
                            Form4.files_result.RemoveItem (teller)
                            Form4.extentions_result.RemoveItem (teller)
                            Form4.filesizes_result.RemoveItem (teller)
                        End If
                    Loop Until teller >= Form4.extentions_result.ListCount
                    GoTo nextquery
                                                
                Case "LIKE"
                    teller = 0
                    Do
                        If InStr(1, Form4.extentions_result.List(teller), q3, vbTextCompare) > 0 Then
                            'do nothing
                        teller = teller + 1
                        Else
                            Form4.folders_result.RemoveItem (teller)
                            Form4.files_result.RemoveItem (teller)
                            Form4.extentions_result.RemoveItem (teller)
                            Form4.filesizes_result.RemoveItem (teller)
                        End If
                    Loop Until teller >= Form4.extentions_result.ListCount
                    GoTo nextquery
                    
                Case "NOT LIKE"
                    teller = 0
                    Do
                        If InStr(1, Form4.extentions_result.List(teller), q3, vbTextCompare) = 0 Then
                            'do nothing
                        teller = teller + 1
                        Else
                            Form4.folders_result.RemoveItem (teller)
                            Form4.files_result.RemoveItem (teller)
                            Form4.extentions_result.RemoveItem (teller)
                            Form4.filesizes_result.RemoveItem (teller)
                        End If
                    Loop Until teller >= Form4.extentions_result.ListCount
                    GoTo nextquery
            End Select
            
            
            
            
            
            Case "Filecontent"
            Select Case q2
                Case "IS"
                    teller = 0
                    Do
                        On Error Resume Next
                        Open (Form4.folders_result.List(teller) & Form4.files_result.List(teller) & "." & Form4.extentions_result.List(teller)) For Input As #1
                            On Error Resume Next
                            Input #1, lijn
                                If lijn = q3 Then
                                    'do nothing
                                    teller = teller + 1
                                Else
                                    Form4.folders_result.RemoveItem (teller)
                                    Form4.files_result.RemoveItem (teller)
                                    Form4.extentions_result.RemoveItem (teller)
                                    Form4.filesizes_result.RemoveItem (teller)
                                End If
                        Close #1
                    Loop Until teller >= Form4.files_result.ListCount
                    
                    GoTo nextquery
                
                Case "IS NOT"
                    teller = 0
                    Do
                        On Error Resume Next
                        Open (Form4.folders_result.List(teller) & Form4.files_result.List(teller) & "." & Form4.extentions_result.List(teller)) For Input As #1
                        On Error Resume Next
                            Input #1, lijn
                                If lijn <> q3 Then
                                    'do nothing
                                    teller = teller + 1
                                Else
                                    Form4.folders_result.RemoveItem (teller)
                                    Form4.files_result.RemoveItem (teller)
                                    Form4.extentions_result.RemoveItem (teller)
                                    Form4.filesizes_result.RemoveItem (teller)
                                End If
                        Close #1
                    Loop Until teller >= Form4.files_result.ListCount
                    GoTo nextquery
                    
                Case "IS >"
                    teller = 0
                    Do
                        On Error Resume Next
                        Open (Form4.folders_result.List(teller) & Form4.files_result.List(teller) & "." & Form4.extentions_result.List(teller)) For Input As #1
                        On Error Resume Next
                            Input #1, lijn
                                If lijn > q3 Then
                                    'do nothing
                                    teller = teller + 1
                                Else
                                    Form4.folders_result.RemoveItem (teller)
                                    Form4.files_result.RemoveItem (teller)
                                    Form4.extentions_result.RemoveItem (teller)
                                    Form4.filesizes_result.RemoveItem (teller)
                                End If
                        Close #1
                    Loop Until teller >= Form4.files_result.ListCount
                    GoTo nextquery
                    
                Case "IS <"
                    teller = 0
                    Do
                        On Error Resume Next
                        Open (Form4.folders_result.List(teller) & Form4.files_result.List(teller) & "." & Form4.extentions_result.List(teller)) For Input As #1
                        On Error Resume Next
                            Input #1, lijn
                                If lijn < q3 Then
                                    'do nothing
                                    teller = teller + 1
                                Else
                                    Form4.folders_result.RemoveItem (teller)
                                    Form4.files_result.RemoveItem (teller)
                                    Form4.extentions_result.RemoveItem (teller)
                                    Form4.filesizes_result.RemoveItem (teller)
                                End If
                        Close #1
                    Loop Until teller >= Form4.files_result.ListCount
                    GoTo nextquery
                                                
                Case "LIKE"
                    teller = 0
                    Do
                        On Error Resume Next
                        Open (Form4.folders_result.List(teller) & Form4.files_result.List(teller) & "." & Form4.extentions_result.List(teller)) For Input As #1
                        Do
                        On Error Resume Next
                            Input #1, lijn
                                If InStr(1, lijn, q3, vbTextCompare) > 0 Then
                                    'do nothing
                                    teller = teller + 1
                                Else
                                    Form4.folders_result.RemoveItem (teller)
                                    Form4.files_result.RemoveItem (teller)
                                    Form4.extentions_result.RemoveItem (teller)
                                    Form4.filesizes_result.RemoveItem (teller)
                                End If
                        Loop Until EOF(1)
                        Close #1
                    Loop Until teller >= Form4.files_result.ListCount
                    GoTo nextquery
                    
                Case "NOT LIKE"
                    teller = 0
                    Do
                        On Error Resume Next
                        Open (Form4.folders_result.List(teller) & Form4.files_result.List(teller) & "." & Form4.extentions_result.List(teller)) For Input As #1
                        Do
                        On Error Resume Next
                            Input #1, lijn
                                If InStr(1, lijn, q3, vbTextCompare) = 0 Then
                                    'do nothing
                                    teller = teller + 1
                                Else
                                    Form4.folders_result.RemoveItem (teller)
                                    Form4.files_result.RemoveItem (teller)
                                    Form4.extentions_result.RemoveItem (teller)
                                    Form4.filesizes_result.RemoveItem (teller)
                                End If
                        Loop Until EOF(1)
                        Close #1
                    Loop Until teller >= Form4.files_result.ListCount
                    GoTo nextquery
                End Select
            







    End Select
End Select
    
                                                                    
nextquery:
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    
    teller_query = teller_query + 1
Loop Until teller_query >= queries.ListCount


teller = 0
If Form4.files_result.ListCount > 0 Then
    Do
        List1.AddItem Form4.folders_result.List(teller) & Form4.files_result.List(teller) & "." & Form4.extentions_result.List(teller)
        teller = teller + 1
    Loop Until teller >= Form4.folders_result.ListCount
End If
lbl_results.Caption = "Found " & teller & " results."
    
    
Unload Form6
End Sub

Private Sub Command3_Click()
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

Private Sub Command4_Click()
queries.Clear
End Sub

Private Sub Command5_Click()
If List1.ListCount > 1000 Then
    vraag = MsgBox("This operation could take a while, continue?", vbQuestion + vbYesNo, "problem?")
    If vraag = vbYes Then
        Form5.Show
    End If
Else
    Form5.Show
End If
End Sub

Private Sub Form_Load()
Form4.Hide
End Sub

