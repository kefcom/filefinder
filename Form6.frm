VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "running query..."
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5730
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Caption         =   "running query,... please wait..."
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "Form6.frx":0000
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
