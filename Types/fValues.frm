VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fValues 
   Caption         =   "File Types Example"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7110
   Icon            =   "fValues.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7110
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   5040
      TabIndex        =   13
      Top             =   960
      Width           =   1935
   End
   Begin VB.PictureBox pLarge 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   3720
      Width           =   480
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2160
      TabIndex        =   10
      Top             =   360
      Width           =   4815
   End
   Begin VB.PictureBox pSmall 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   840
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      Top             =   3840
      Width           =   240
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   6855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      ToolTipText     =   "Exit"
      Top             =   3800
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin MSComctlLib.ListView lvwFile 
      Height          =   1650
      Left            =   120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1920
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   2910
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Action"
         Text            =   "Action"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Command"
         Text            =   "Action Command"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label5 
      Caption         =   "Default Action"
      Height          =   255
      Left            =   5040
      TabIndex        =   14
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "File Type"
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "File Extension"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Default Icon File"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Registry Key Value"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Context Type MIME"
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "fValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
   On Error Resume Next
   Unload fValues
   Set fValues = Nothing
End Sub
