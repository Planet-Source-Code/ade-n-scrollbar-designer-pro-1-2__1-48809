VERSION 5.00
Begin VB.Form frmCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Code - Scrollbar Designer Pro"
   ClientHeight    =   4200
   ClientLeft      =   7425
   ClientTop       =   1245
   ClientWidth     =   6360
   Icon            =   "frmCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6360
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read CSS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      ToolTipText     =   "If you have entered CSS code into the textbox, the colors in it will be read and put into the scrollbar designer"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      ToolTipText     =   "Copy the code to the clipboard"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   45
      Width           =   6255
   End
   Begin VB.Label lblStatus 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   3255
   End
End
Attribute VB_Name = "frmCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------
' Scrollbar Designer Pro
'--------------------------
' Welcome to the sourcecode of Scrollbar Designer Pro.
' There's not alot of comments, I hope you find your way.
' Thanks for voting; http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=48809&lngWId=1
'
' / aDe
'
' http://www.ade.se

Private Sub cmdCopy_Click()
Clipboard.Clear
Clipboard.SetText txtCode.Text

End Sub

Private Sub cmdRead_Click()
lblStatus.Caption = frmMain.ReadCSS(txtCode.Text) & " of 7 possible colors read."
End Sub
