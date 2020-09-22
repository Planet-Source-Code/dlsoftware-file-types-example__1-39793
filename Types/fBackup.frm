VERSION 5.00
Begin VB.Form fBackup 
   Caption         =   "File Types Example"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   Icon            =   "fBackup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   6195
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      ToolTipText     =   "Browse File"
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      ToolTipText     =   "Browse File"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      ToolTipText     =   "Exit"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Backup file for the file type:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Backup file for the file extention:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "fBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rc1 As Long

Private Sub Command1_Click()
   On Error Resume Next
   Unload fBackup
   Set fBackup = Nothing
End Sub

Private Sub Command2_Click()
   On Error Resume Next
   rc1 = Shell("notepad.exe " & Text1.Text, vbNormalFocus)
End Sub

Private Sub Command3_Click()
   On Error Resume Next
   rc1 = Shell("notepad.exe " & Text2.Text, vbNormalFocus)
End Sub
