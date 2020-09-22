VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scrollbar Designer Pro"
   ClientHeight    =   3945
   ClientLeft      =   4650
   ClientTop       =   3660
   ClientWidth     =   9705
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   9705
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD 
      Left            =   5040
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import colors from file.."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5760
      TabIndex        =   204
      ToolTipText     =   "Attempt to find the scrollbar colors in a HTML / CSS textfile."
      Top             =   3540
      Width           =   2220
   End
   Begin VB.CommandButton cmdToCSS 
      Caption         =   "Edit/View CSS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6675
      TabIndex        =   203
      ToolTipText     =   "Click here to view the CSS code you need."
      Top             =   3195
      Width           =   1305
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5760
      TabIndex        =   202
      ToolTipText     =   "Click here to see a preview of the scrollbar in a Internet Explorer window."
      Top             =   3195
      Width           =   900
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8850
      TabIndex        =   189
      Top             =   3195
      Width           =   735
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset all"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7995
      TabIndex        =   188
      Top             =   3195
      Width           =   840
   End
   Begin VB.Frame frameEdit 
      Caption         =   "Edit property..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5760
      TabIndex        =   37
      Top             =   0
      Width           =   3855
      Begin VB.CheckBox chkAutoselect 
         Caption         =   "Autoselect colors"
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
         Left            =   2280
         TabIndex        =   198
         ToolTipText     =   "With this checked, you only need to pick one color and the corresponding colors will be generated for the whole scrollbar."
         Top             =   0
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.OptionButton optEdit 
         Caption         =   "Highlight"
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
         Index           =   6
         Left            =   2280
         TabIndex        =   44
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optEdit 
         Caption         =   "Track"
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
         Index           =   5
         Left            =   1440
         TabIndex        =   43
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optEdit 
         Caption         =   "Dark shadow"
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
         Index           =   4
         Left            =   120
         TabIndex        =   42
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optEdit 
         Caption         =   "Shadow"
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
         Index           =   3
         Left            =   2640
         TabIndex        =   41
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optEdit 
         Caption         =   "Face"
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
         Index           =   2
         Left            =   1920
         TabIndex        =   40
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optEdit 
         Caption         =   "Arrow"
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
         Index           =   1
         Left            =   1080
         TabIndex        =   39
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optEdit 
         Caption         =   "3d Light"
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
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.PictureBox picScrollbar 
      Height          =   3735
      Left            =   120
      Picture         =   "frmMain.frx":0E42
      ScaleHeight     =   3675
      ScaleWidth      =   1035
      TabIndex        =   1
      Top             =   120
      Width           =   1095
      Begin VB.PictureBox scr_Track 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   0
         Left            =   0
         ScaleHeight     =   480
         ScaleWidth      =   960
         TabIndex        =   34
         Tag             =   "Track"
         Top             =   2160
         Width           =   960
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   127
            Left            =   900
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   187
            Tag             =   "Arrow"
            Top             =   420
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   126
            Left            =   840
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   186
            Tag             =   "Arrow"
            Top             =   420
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   125
            Left            =   780
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   185
            Tag             =   "Arrow"
            Top             =   420
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   124
            Left            =   720
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   184
            Tag             =   "Arrow"
            Top             =   420
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   123
            Left            =   660
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   183
            Tag             =   "Arrow"
            Top             =   420
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   122
            Left            =   600
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   182
            Tag             =   "Arrow"
            Top             =   420
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   121
            Left            =   540
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   181
            Tag             =   "Arrow"
            Top             =   420
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   120
            Left            =   480
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   180
            Tag             =   "Arrow"
            Top             =   420
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   119
            Left            =   420
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   179
            Tag             =   "Arrow"
            Top             =   420
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   118
            Left            =   360
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   178
            Tag             =   "Arrow"
            Top             =   420
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   117
            Left            =   300
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   177
            Tag             =   "Arrow"
            Top             =   420
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   116
            Left            =   240
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   176
            Tag             =   "Arrow"
            Top             =   420
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   115
            Left            =   180
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   175
            Tag             =   "Arrow"
            Top             =   420
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   114
            Left            =   120
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   174
            Tag             =   "Arrow"
            Top             =   420
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   113
            Left            =   60
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   173
            Tag             =   "Arrow"
            Top             =   420
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   112
            Left            =   0
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   172
            Tag             =   "Arrow"
            Top             =   420
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   111
            Left            =   900
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   171
            Tag             =   "Arrow"
            Top             =   360
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   110
            Left            =   840
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   170
            Tag             =   "Arrow"
            Top             =   360
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   109
            Left            =   780
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   169
            Tag             =   "Arrow"
            Top             =   360
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   108
            Left            =   720
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   168
            Tag             =   "Arrow"
            Top             =   360
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   107
            Left            =   660
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   167
            Tag             =   "Arrow"
            Top             =   360
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   106
            Left            =   600
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   166
            Tag             =   "Arrow"
            Top             =   360
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   105
            Left            =   540
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   165
            Tag             =   "Arrow"
            Top             =   360
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   104
            Left            =   480
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   164
            Tag             =   "Arrow"
            Top             =   360
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   103
            Left            =   420
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   163
            Tag             =   "Arrow"
            Top             =   360
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   102
            Left            =   360
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   162
            Tag             =   "Arrow"
            Top             =   360
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   101
            Left            =   300
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   161
            Tag             =   "Arrow"
            Top             =   360
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   100
            Left            =   240
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   160
            Tag             =   "Arrow"
            Top             =   360
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   99
            Left            =   180
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   159
            Tag             =   "Arrow"
            Top             =   360
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   98
            Left            =   120
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   158
            Tag             =   "Arrow"
            Top             =   360
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   97
            Left            =   60
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   157
            Tag             =   "Arrow"
            Top             =   360
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   96
            Left            =   0
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   156
            Tag             =   "Arrow"
            Top             =   360
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   95
            Left            =   900
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   155
            Tag             =   "Arrow"
            Top             =   300
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   94
            Left            =   840
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   154
            Tag             =   "Arrow"
            Top             =   300
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   93
            Left            =   780
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   153
            Tag             =   "Arrow"
            Top             =   300
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   92
            Left            =   720
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   152
            Tag             =   "Arrow"
            Top             =   300
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   91
            Left            =   660
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   151
            Tag             =   "Arrow"
            Top             =   300
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   90
            Left            =   600
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   150
            Tag             =   "Arrow"
            Top             =   300
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   89
            Left            =   540
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   149
            Tag             =   "Arrow"
            Top             =   300
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   88
            Left            =   480
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   148
            Tag             =   "Arrow"
            Top             =   300
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   87
            Left            =   420
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   147
            Tag             =   "Arrow"
            Top             =   300
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   86
            Left            =   360
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   146
            Tag             =   "Arrow"
            Top             =   300
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   85
            Left            =   300
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   145
            Tag             =   "Arrow"
            Top             =   300
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   84
            Left            =   240
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   144
            Tag             =   "Arrow"
            Top             =   300
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   83
            Left            =   180
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   143
            Tag             =   "Arrow"
            Top             =   300
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   82
            Left            =   120
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   142
            Tag             =   "Arrow"
            Top             =   300
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   81
            Left            =   60
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   141
            Tag             =   "Arrow"
            Top             =   300
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   80
            Left            =   0
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   140
            Tag             =   "Arrow"
            Top             =   300
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   79
            Left            =   900
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   139
            Tag             =   "Arrow"
            Top             =   240
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   78
            Left            =   840
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   138
            Tag             =   "Arrow"
            Top             =   240
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   77
            Left            =   780
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   137
            Tag             =   "Arrow"
            Top             =   240
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   76
            Left            =   720
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   136
            Tag             =   "Arrow"
            Top             =   240
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   75
            Left            =   660
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   135
            Tag             =   "Arrow"
            Top             =   240
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   74
            Left            =   600
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   134
            Tag             =   "Arrow"
            Top             =   240
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   73
            Left            =   540
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   133
            Tag             =   "Arrow"
            Top             =   240
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   72
            Left            =   480
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   132
            Tag             =   "Arrow"
            Top             =   240
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   71
            Left            =   420
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   131
            Tag             =   "Arrow"
            Top             =   240
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   70
            Left            =   360
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   130
            Tag             =   "Arrow"
            Top             =   240
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   69
            Left            =   300
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   129
            Tag             =   "Arrow"
            Top             =   240
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   68
            Left            =   240
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   128
            Tag             =   "Arrow"
            Top             =   240
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   67
            Left            =   180
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   127
            Tag             =   "Arrow"
            Top             =   240
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   66
            Left            =   120
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   126
            Tag             =   "Arrow"
            Top             =   240
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   65
            Left            =   60
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   125
            Tag             =   "Arrow"
            Top             =   240
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   64
            Left            =   0
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   124
            Tag             =   "Arrow"
            Top             =   240
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   63
            Left            =   900
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   123
            Tag             =   "Arrow"
            Top             =   180
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   62
            Left            =   840
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   122
            Tag             =   "Arrow"
            Top             =   180
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   61
            Left            =   780
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   121
            Tag             =   "Arrow"
            Top             =   180
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   60
            Left            =   720
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   120
            Tag             =   "Arrow"
            Top             =   180
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   59
            Left            =   660
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   119
            Tag             =   "Arrow"
            Top             =   180
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   58
            Left            =   600
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   118
            Tag             =   "Arrow"
            Top             =   180
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   57
            Left            =   540
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   117
            Tag             =   "Arrow"
            Top             =   180
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   56
            Left            =   480
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   116
            Tag             =   "Arrow"
            Top             =   180
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   55
            Left            =   420
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   115
            Tag             =   "Arrow"
            Top             =   180
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   54
            Left            =   360
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   114
            Tag             =   "Arrow"
            Top             =   180
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   53
            Left            =   300
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   113
            Tag             =   "Arrow"
            Top             =   180
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   52
            Left            =   240
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   112
            Tag             =   "Arrow"
            Top             =   180
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   51
            Left            =   180
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   111
            Tag             =   "Arrow"
            Top             =   180
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   50
            Left            =   120
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   110
            Tag             =   "Arrow"
            Top             =   180
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   49
            Left            =   60
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   109
            Tag             =   "Arrow"
            Top             =   180
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   48
            Left            =   0
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   108
            Tag             =   "Arrow"
            Top             =   180
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   47
            Left            =   900
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   107
            Tag             =   "Arrow"
            Top             =   120
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   46
            Left            =   840
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   106
            Tag             =   "Arrow"
            Top             =   120
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   45
            Left            =   780
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   105
            Tag             =   "Arrow"
            Top             =   120
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   44
            Left            =   720
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   104
            Tag             =   "Arrow"
            Top             =   120
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   43
            Left            =   660
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   103
            Tag             =   "Arrow"
            Top             =   120
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   42
            Left            =   600
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   102
            Tag             =   "Arrow"
            Top             =   120
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   41
            Left            =   540
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   101
            Tag             =   "Arrow"
            Top             =   120
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   40
            Left            =   480
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   100
            Tag             =   "Arrow"
            Top             =   120
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   39
            Left            =   420
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   99
            Tag             =   "Arrow"
            Top             =   120
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   38
            Left            =   360
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   98
            Tag             =   "Arrow"
            Top             =   120
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   37
            Left            =   300
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   97
            Tag             =   "Arrow"
            Top             =   120
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   36
            Left            =   240
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   96
            Tag             =   "Arrow"
            Top             =   120
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   35
            Left            =   180
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   95
            Tag             =   "Arrow"
            Top             =   120
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   34
            Left            =   120
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   94
            Tag             =   "Arrow"
            Top             =   120
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   33
            Left            =   60
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   93
            Tag             =   "Arrow"
            Top             =   120
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   32
            Left            =   0
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   92
            Tag             =   "Arrow"
            Top             =   120
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   31
            Left            =   900
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   91
            Tag             =   "Arrow"
            Top             =   60
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   30
            Left            =   840
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   90
            Tag             =   "Arrow"
            Top             =   60
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   29
            Left            =   780
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   89
            Tag             =   "Arrow"
            Top             =   60
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   28
            Left            =   720
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   88
            Tag             =   "Arrow"
            Top             =   60
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   27
            Left            =   660
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   87
            Tag             =   "Arrow"
            Top             =   60
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   26
            Left            =   600
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   86
            Tag             =   "Arrow"
            Top             =   60
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   25
            Left            =   540
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   85
            Tag             =   "Arrow"
            Top             =   60
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   24
            Left            =   480
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   84
            Tag             =   "Arrow"
            Top             =   60
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   23
            Left            =   420
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   83
            Tag             =   "Arrow"
            Top             =   60
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   22
            Left            =   360
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   82
            Tag             =   "Arrow"
            Top             =   60
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   21
            Left            =   300
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   81
            Tag             =   "Arrow"
            Top             =   60
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   20
            Left            =   240
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   80
            Tag             =   "Arrow"
            Top             =   60
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   19
            Left            =   180
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   79
            Tag             =   "Arrow"
            Top             =   60
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   18
            Left            =   120
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   78
            Tag             =   "Arrow"
            Top             =   60
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   17
            Left            =   60
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   77
            Tag             =   "Arrow"
            Top             =   60
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   16
            Left            =   0
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   76
            Tag             =   "Arrow"
            Top             =   60
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   15
            Left            =   900
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   75
            Tag             =   "Arrow"
            Top             =   0
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   14
            Left            =   840
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   74
            Tag             =   "Arrow"
            Top             =   0
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   13
            Left            =   780
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   73
            Tag             =   "Arrow"
            Top             =   0
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   12
            Left            =   720
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   72
            Tag             =   "Arrow"
            Top             =   0
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   11
            Left            =   660
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   71
            Tag             =   "Arrow"
            Top             =   0
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   10
            Left            =   600
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   70
            Tag             =   "Arrow"
            Top             =   0
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   9
            Left            =   540
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   69
            Tag             =   "Arrow"
            Top             =   0
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   8
            Left            =   480
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   68
            Tag             =   "Arrow"
            Top             =   0
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   7
            Left            =   420
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   67
            Tag             =   "Arrow"
            Top             =   0
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   6
            Left            =   360
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   66
            Tag             =   "Arrow"
            Top             =   0
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   5
            Left            =   300
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   65
            Tag             =   "Arrow"
            Top             =   0
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   4
            Left            =   240
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   64
            Tag             =   "Arrow"
            Top             =   0
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   3
            Left            =   180
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   63
            Tag             =   "Arrow"
            Top             =   0
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   2
            Left            =   120
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   62
            Tag             =   "Arrow"
            Top             =   0
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   1
            Left            =   60
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   61
            Tag             =   "Arrow"
            Top             =   0
            Width           =   60
         End
         Begin VB.PictureBox scr_TrackGrid 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   60
            Index           =   0
            Left            =   0
            ScaleHeight     =   60
            ScaleWidth      =   60
            TabIndex        =   60
            Tag             =   "Arrow"
            Top             =   0
            Width           =   60
         End
      End
      Begin VB.PictureBox scr_Arrow 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   7
         Left            =   420
         ScaleHeight     =   60
         ScaleWidth      =   60
         TabIndex        =   33
         Tag             =   "Arrow"
         Top             =   3180
         Width           =   60
      End
      Begin VB.PictureBox scr_Arrow 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   6
         Left            =   360
         ScaleHeight     =   60
         ScaleWidth      =   180
         TabIndex        =   32
         Tag             =   "Arrow"
         Top             =   3120
         Width           =   185
      End
      Begin VB.PictureBox scr_Arrow 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   5
         Left            =   300
         ScaleHeight     =   60
         ScaleWidth      =   300
         TabIndex        =   31
         Tag             =   "Arrow"
         Top             =   3060
         Width           =   305
      End
      Begin VB.PictureBox scr_Arrow 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   4
         Left            =   240
         ScaleHeight     =   60
         ScaleWidth      =   420
         TabIndex        =   30
         Tag             =   "Arrow"
         Top             =   3000
         Width           =   415
      End
      Begin VB.PictureBox scr_Arrow 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   3
         Left            =   420
         ScaleHeight     =   60
         ScaleWidth      =   60
         TabIndex        =   29
         Tag             =   "Arrow"
         Top             =   360
         Width           =   60
      End
      Begin VB.PictureBox scr_Arrow 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   2
         Left            =   360
         ScaleHeight     =   60
         ScaleWidth      =   180
         TabIndex        =   28
         Tag             =   "Arrow"
         Top             =   420
         Width           =   185
      End
      Begin VB.PictureBox scr_Arrow 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   1
         Left            =   300
         ScaleHeight     =   60
         ScaleWidth      =   300
         TabIndex        =   27
         Tag             =   "Arrow"
         Top             =   480
         Width           =   305
      End
      Begin VB.PictureBox scr_Arrow 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   0
         Left            =   240
         ScaleHeight     =   60
         ScaleWidth      =   420
         TabIndex        =   26
         Tag             =   "Arrow"
         Top             =   540
         Width           =   415
      End
      Begin VB.PictureBox scr_Face 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   960
         Index           =   0
         Left            =   120
         ScaleHeight     =   960
         ScaleWidth      =   720
         TabIndex        =   25
         Tag             =   "Face"
         Top             =   1080
         Width           =   715
      End
      Begin VB.PictureBox scr_Shadow 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   5
         Left            =   60
         ScaleHeight     =   60
         ScaleWidth      =   780
         TabIndex        =   24
         Tag             =   "Shadow"
         Top             =   3480
         Width           =   780
      End
      Begin VB.PictureBox scr_Shadow 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   4
         Left            =   60
         ScaleHeight     =   60
         ScaleWidth      =   780
         TabIndex        =   23
         Tag             =   "Shadow"
         Top             =   2040
         Width           =   780
      End
      Begin VB.PictureBox scr_Shadow 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   3
         Left            =   60
         ScaleHeight     =   60
         ScaleWidth      =   780
         TabIndex        =   22
         Tag             =   "Shadow"
         Top             =   840
         Width           =   780
      End
      Begin VB.PictureBox scr_Shadow 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   840
         Index           =   2
         Left            =   840
         ScaleHeight     =   840
         ScaleWidth      =   60
         TabIndex        =   21
         Tag             =   "Shadow"
         Top             =   2700
         Width           =   60
      End
      Begin VB.PictureBox scr_Shadow 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   1075
         Index           =   1
         Left            =   840
         ScaleHeight     =   1080
         ScaleWidth      =   60
         TabIndex        =   20
         Tag             =   "Shadow"
         Top             =   1020
         Width           =   60
      End
      Begin VB.PictureBox scr_Shadow 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   840
         Index           =   0
         Left            =   840
         ScaleHeight     =   840
         ScaleWidth      =   60
         TabIndex        =   19
         Tag             =   "Shadow"
         Top             =   60
         Width           =   60
      End
      Begin VB.PictureBox scr_Highlight 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   965
         Index           =   5
         Left            =   60
         ScaleHeight     =   960
         ScaleWidth      =   60
         TabIndex        =   18
         Tag             =   "Highlight"
         Top             =   1080
         Width           =   60
      End
      Begin VB.PictureBox scr_Highlight 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   715
         Index           =   4
         Left            =   60
         ScaleHeight     =   720
         ScaleWidth      =   60
         TabIndex        =   17
         Tag             =   "Highlight"
         Top             =   2760
         Width           =   60
      End
      Begin VB.PictureBox scr_Highlight 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   715
         Index           =   3
         Left            =   60
         ScaleHeight     =   720
         ScaleWidth      =   60
         TabIndex        =   16
         Tag             =   "Highlight"
         Top             =   120
         Width           =   60
      End
      Begin VB.PictureBox scr_Darkshadow 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   965
         Index           =   4
         Left            =   900
         ScaleHeight     =   960
         ScaleWidth      =   60
         TabIndex        =   15
         Tag             =   "Dark shadow"
         Top             =   2640
         Width           =   60
      End
      Begin VB.PictureBox scr_Highlight 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   2
         Left            =   60
         ScaleHeight     =   60
         ScaleWidth      =   780
         TabIndex        =   14
         Tag             =   "Highlight"
         Top             =   2700
         Width           =   780
      End
      Begin VB.PictureBox scr_Highlight 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   1
         Left            =   60
         ScaleHeight     =   60
         ScaleWidth      =   780
         TabIndex        =   13
         Tag             =   "Highlight"
         Top             =   1020
         Width           =   780
      End
      Begin VB.PictureBox scr_Highlight 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   0
         Left            =   60
         ScaleHeight     =   60
         ScaleWidth      =   780
         TabIndex        =   12
         Tag             =   "Highlight"
         Top             =   60
         Width           =   780
      End
      Begin VB.PictureBox scr_Darkshadow 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   3
         Left            =   0
         ScaleHeight     =   60
         ScaleWidth      =   900
         TabIndex        =   11
         Tag             =   "Dark shadow"
         Top             =   3540
         Width           =   900
      End
      Begin VB.PictureBox scr_Darkshadow 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   2
         Left            =   0
         ScaleHeight     =   60
         ScaleWidth      =   900
         TabIndex        =   10
         Tag             =   "Dark shadow"
         Top             =   2100
         Width           =   900
      End
      Begin VB.PictureBox scr_Darkshadow 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   1
         Left            =   0
         ScaleHeight     =   60
         ScaleWidth      =   900
         TabIndex        =   9
         Tag             =   "Dark shadow"
         Top             =   900
         Width           =   900
      End
      Begin VB.PictureBox scr_Darkshadow 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   2160
         Index           =   0
         Left            =   900
         ScaleHeight     =   2160
         ScaleWidth      =   60
         TabIndex        =   8
         Tag             =   "Dark shadow"
         Top             =   0
         Width           =   60
      End
      Begin VB.PictureBox scr_3dlight 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   1080
         Index           =   5
         Left            =   0
         ScaleHeight     =   1080
         ScaleWidth      =   60
         TabIndex        =   7
         Tag             =   "3d Light"
         Top             =   1020
         Width           =   60
      End
      Begin VB.PictureBox scr_3dlight 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   840
         Index           =   4
         Left            =   0
         ScaleHeight     =   840
         ScaleWidth      =   60
         TabIndex        =   6
         Tag             =   "3d Light"
         Top             =   2700
         Width           =   60
      End
      Begin VB.PictureBox scr_3dlight 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   840
         Index           =   3
         Left            =   0
         ScaleHeight     =   840
         ScaleWidth      =   60
         TabIndex        =   5
         Tag             =   "3d Light"
         Top             =   60
         Width           =   60
      End
      Begin VB.PictureBox scr_3dlight 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   2
         Left            =   0
         ScaleHeight     =   60
         ScaleWidth      =   900
         TabIndex        =   4
         Tag             =   "3d Light"
         Top             =   2640
         Width           =   900
      End
      Begin VB.PictureBox scr_3dlight 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   1
         Left            =   0
         ScaleHeight     =   60
         ScaleWidth      =   900
         TabIndex        =   3
         Tag             =   "3d Light"
         Top             =   960
         Width           =   900
      End
      Begin VB.PictureBox scr_3dlight 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   0
         Left            =   0
         ScaleHeight     =   60
         ScaleWidth      =   900
         TabIndex        =   2
         Tag             =   "3d Light"
         Top             =   0
         Width           =   900
      End
      Begin VB.PictureBox scr_Face 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   730
         Index           =   1
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   720
         TabIndex        =   35
         Tag             =   "Face"
         Top             =   105
         Width           =   715
      End
      Begin VB.PictureBox scr_Face 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   730
         Index           =   2
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   720
         TabIndex        =   36
         Tag             =   "Face"
         Top             =   2760
         Width           =   715
      End
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00808080&
      Height          =   3705
      Left            =   1320
      Picture         =   "frmMain.frx":C284
      ScaleHeight     =   3645
      ScaleWidth      =   4260
      TabIndex        =   0
      Top             =   120
      Width           =   4320
   End
   Begin VB.Frame frameColor 
      Caption         =   "Color"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   5760
      TabIndex        =   45
      Top             =   960
      Width           =   3855
      Begin VB.TextBox txtSaturation 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3390
         TabIndex        =   200
         Text            =   "0"
         Top             =   1560
         Width           =   405
      End
      Begin VB.TextBox txtHue 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3390
         TabIndex        =   199
         Text            =   "0"
         Top             =   1320
         Width           =   405
      End
      Begin VB.Frame frameApply 
         Height          =   855
         Left            =   2640
         TabIndex        =   195
         Top             =   0
         Width           =   1215
         Begin VB.CommandButton cmdApply 
            Caption         =   "Apply"
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
            Left            =   165
            TabIndex        =   196
            ToolTipText     =   "Apply the current color to the selected surface"
            Top             =   165
            Width           =   855
         End
         Begin VB.CheckBox chkAuto 
            Caption         =   "Auto"
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
            Left            =   165
            TabIndex        =   197
            ToolTipText     =   "Automatically apply the selected color when changed"
            Top             =   540
            Value           =   1  'Checked
            Width           =   975
         End
      End
      Begin VB.PictureBox picCurHue 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   510
         ScaleHeight     =   180
         ScaleWidth      =   360
         TabIndex        =   194
         Top             =   1350
         Width           =   390
      End
      Begin VB.HScrollBar scrLight 
         Height          =   255
         LargeChange     =   16
         Left            =   960
         Max             =   240
         TabIndex        =   192
         Top             =   1800
         Width           =   2415
      End
      Begin VB.HScrollBar scrSat 
         Height          =   255
         LargeChange     =   16
         Left            =   960
         Max             =   240
         TabIndex        =   57
         Top             =   1560
         Width           =   2415
      End
      Begin VB.HScrollBar scrHue 
         Height          =   255
         LargeChange     =   16
         Left            =   960
         Max             =   239
         TabIndex        =   190
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CheckBox chkNoColor 
         Height          =   240
         Left            =   3420
         Picture         =   "frmMain.frx":41914
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "No color - click here to toggle wheter the selected surface should be colored"
         Top             =   960
         Value           =   1  'Checked
         Width           =   315
      End
      Begin VB.TextBox txtBlue 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   55
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtGreen 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   53
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtRed 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         MaxLength       =   3
         TabIndex        =   51
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtHexColor 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         MaxLength       =   6
         TabIndex        =   47
         Text            =   "000000"
         Top             =   240
         Width           =   1575
      End
      Begin VB.PictureBox imgCurColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         ScaleHeight     =   225
         ScaleWidth      =   1545
         TabIndex        =   46
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtLightness 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3390
         TabIndex        =   201
         Text            =   "0"
         Top             =   1800
         Width           =   405
      End
      Begin VB.Label lblGui 
         Caption         =   "Lightness:"
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
         Index           =   8
         Left            =   120
         TabIndex        =   193
         Top             =   1800
         Width           =   840
      End
      Begin VB.Label lblGui 
         Caption         =   "Hue:"
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
         Index           =   6
         Left            =   120
         TabIndex        =   191
         Top             =   1320
         Width           =   840
      End
      Begin VB.Label lblGui 
         Caption         =   "#"
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
         Index           =   7
         Left            =   840
         TabIndex        =   58
         Top             =   280
         Width           =   375
      End
      Begin VB.Label lblGui 
         Caption         =   "Saturation:"
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
         Index           =   5
         Left            =   120
         TabIndex        =   56
         Top             =   1560
         Width           =   840
      End
      Begin VB.Label lblGui 
         Caption         =   "Blue"
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
         Index           =   4
         Left            =   2160
         TabIndex        =   54
         Top             =   990
         Width           =   360
      End
      Begin VB.Label lblGui 
         Caption         =   "Green"
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
         Index           =   3
         Left            =   1080
         TabIndex        =   52
         Top             =   990
         Width           =   480
      End
      Begin VB.Label lblGui 
         Caption         =   "Red"
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
         Index           =   2
         Left            =   120
         TabIndex        =   50
         Top             =   990
         Width           =   360
      End
      Begin VB.Label lblGui 
         Caption         =   "Hex:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblGui 
         Caption         =   "Preview:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   48
         Top             =   615
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmMain"
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


Dim clrSelect As Boolean, lColor As Long, refColor As Boolean
Dim iEdit As Integer, objColor(0 To 6) As String, clrObject(0 To 7) As Object
Dim strCaption
Dim sTempFile

Sub CleanUp()
  If Dir(sTempFile) <> "" Then Kill sTempFile
End Sub

Sub RefreshColor(Optional skipGroup As Integer)
Dim R As Integer, G As Integer, B As Integer, nHSL As HSL
Dim iOldEdit As Integer, lCColor As Long, nRGB As RGB, nLum As Integer
refColor = True
  
  R = (lColor And &HFF&) 'red
  G = (lColor And &HFF00&) / &H100& ' green
  B = (lColor And &HFF0000) / &H10000 ' blue
  nHSL = RGBtoHSL(R, G, B)
  
 
  If skipGroup <> grpHEX Then
    txtHexColor.Text = MakeHex(lColor)
  End If
  
  If skipGroup <> grpRGB Then
    txtRed.Text = R
    txtBlue.Text = B
    txtGreen.Text = G
  End If

  If skipGroup <> grpHSL Then
    scrHue.Value = nHSL.Hue
    scrSat.Value = nHSL.Saturation
    scrLight.Value = nHSL.Luminance
  End If
  
  If skipGroup <> grpHSLText Then
    txtHue = scrHue.Value
    txtSaturation = scrSat.Value
    txtLightness = scrLight.Value
  End If

  imgCurColor.BackColor = lColor
  If (skipGroup <> grpHSL) And (skipGroup <> grpHSLText) Then picCurHue.BackColor = HSL(nHSL.Hue, 240, 120)
  refColor = False
  
  If chkAutoselect.Value Then
    iOldEdit = iEdit
    SetEditItem numFace
    cmdApply_Click
    SetEditItem num3dLight
    cmdApply_Click
    nLum = Format(nHSL.Luminance * 1.5, "0")
    If nLum < 128 Then nLum = 150
    nRGB = HSLtoRGB(nHSL.Hue, nHSL.Saturation, nLum)
    lColor = RGB(nRGB.Red, nRGB.Green, nRGB.Blue)
    SetEditItem numHighlight
    cmdApply_Click
    SetEditItem numTrack
    cmdApply_Click
    
    nRGB = HSLtoRGB(nHSL.Hue, nHSL.Saturation, Format(nHSL.Luminance * 0.66667, "0"))
    lColor = RGB(nRGB.Red, nRGB.Green, nRGB.Blue)
    SetEditItem numShadow
    cmdApply_Click
    
    lColor = RGB(0, 0, 0)
    SetEditItem numDarkshadow
    cmdApply_Click
    SetEditItem numArrow
    cmdApply_Click
    
    SetEditItem iOldEdit
  Else
    If chkAuto.Value Then cmdApply_Click
  End If
End Sub

Sub refHSL()
  picCurHue.BackColor = HSL(scrHue.Value, 240, 120)
  lColor = HSL(scrHue.Value, scrSat.Value, scrLight.Value)
  RefreshColor grpHSL
End Sub

Sub refHSLText()
  picCurHue.BackColor = HSL(txtHue.Text, 240, 120)
  lColor = HSL(txtHue.Text, txtSaturation.Text, txtLightness.Text)
  RefreshColor grpHSLText
End Sub

Sub refRGB()
  lColor = RGB(Val(txtRed.Text), Val(txtGreen.Text), Val(txtBlue.Text))
  RefreshColor grpRGB
End Sub

Sub SetEditItem(Index As Integer)
iEdit = Index
If objColor(iEdit) <> "/" Then
  chkNoColor.Value = 0
Else
  chkNoColor.Value = 1
End If
End Sub

Private Sub chkAuto_Click()
If chkAuto.Value Then
  cmdApply.Enabled = False
Else
  cmdApply.Enabled = True
End If
End Sub

Private Sub chkAutoselect_Click()
  If chkAutoselect.Value Then
    For i = 0 To optEdit.UBound
      optEdit(i).Enabled = False
    Next
    cmdApply.Enabled = False
    chkAuto.Enabled = False
    chkNoColor.Enabled = False
  Else
    For i = 0 To optEdit.UBound
      optEdit(i).Enabled = True
    Next
    If chkAuto.Value = 0 Then cmdApply.Enabled = True
    chkAuto.Enabled = True
    chkNoColor.Enabled = True
  End If
End Sub

Private Sub chkNoColor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If chkNoColor.Value Then
  objColor(iEdit) = "/"
  If iEdit = numTrack Then ShowTrack False
  ResetColor iEdit
Else
  cmdApply_Click
  If iEdit = numTrack Then ShowTrack True
End If

RefreshTrack
End Sub
Sub ShowTrack(bShow As Boolean)
  For i = 0 To 127
    scr_TrackGrid(i).Visible = Not bShow
  Next
End Sub
Private Sub cmdAbout_Click()
Dim sMsg As String
sMsg = "Scrollbar Designer Pro version " & App.Major & "." & App.Minor & " by ade" & vbCrLf
sMsg = sMsg & "Thanks to Siebe Tolsma for some hex/rgb code," & vbCrLf
sMsg = sMsg & "and thanks to Andrew Gray for the HSL code." & vbCrLf
sMsg = sMsg & "Comments / suggestions to ade@ade.se" & vbCrLf
sMsg = sMsg & "If you want the source (for free), please contact me" & vbCrLf
sMsg = sMsg & "Enjoy"

MsgBox sMsg, vbOKOnly + vbInformation, "About"
End Sub

Private Sub cmdApply_Click()
Dim EditObj As Object
Select Case iEdit
  Case 0 '3d light
    Set EditObj = scr_3dlight
  Case 1 'arrow
    Set EditObj = scr_Arrow
  Case 2 'face
    Set EditObj = scr_Face
  Case 3 'shadow
    Set EditObj = scr_Shadow
  Case 4 'dark shadow
    Set EditObj = scr_Darkshadow
  Case 5 'track
    Set EditObj = scr_Track
    ShowTrack True
  Case 6 'highlight
    Set EditObj = scr_Highlight
End Select

With EditObj
  For i = 0 To .UBound
    With .Item(i)
      .BackColor = lColor
    End With
  Next
End With

objColor(iEdit) = lColor
If chkNoColor.Value Then chkNoColor.Value = 0
RefreshTrack

End Sub
Sub RefreshTrack()
If objColor(numTrack) = "/" Then
  If objColor(numFace) <> "/" Then
    For i = 0 To scr_TrackGrid.UBound - 1 Step 2
      If (i = 16) Or (i = 48) Or (i = 80) Or (i = 112) Then
        X = 1
      ElseIf (i = 32) Or (i = 64) Or (i = 96) Then
        X = 0
      End If
      scr_TrackGrid(i + X).BackColor = objColor(numFace)
    Next
  Else
    For i = 0 To scr_TrackGrid.UBound - 1 Step 2
      If (i = 16) Or (i = 48) Or (i = 80) Or (i = 112) Then
        X = 1
      ElseIf (i = 32) Or (i = 64) Or (i = 96) Then
        X = 0
      End If
      scr_TrackGrid(i + X).BackColor = vb3DFace
    Next
  End If
  If objColor(numHighlight) <> "/" Then
    For i = 0 To scr_TrackGrid.UBound Step 2
      If (i = 16) Or (i = 48) Or (i = 80) Or (i = 112) Then
        X = 0
      ElseIf (i = 32) Or (i = 64) Or (i = 96) Then
        X = 1
      End If
      scr_TrackGrid(i + X).BackColor = objColor(numHighlight)
    Next
  Else
    For i = 0 To scr_TrackGrid.UBound Step 2
      If (i = 16) Or (i = 48) Or (i = 80) Or (i = 112) Then
        X = 0
      ElseIf (i = 32) Or (i = 64) Or (i = 96) Then
        X = 1
      End If
      scr_TrackGrid(i + X).BackColor = vb3DHighlight
    Next
  End If
End If
End Sub
Sub ResetColor(Index As Integer)
Select Case Index
  Case 0 '3d light
    With scr_3dlight
      For i = 0 To .UBound
        With .Item(i)
          .BackColor = hexToRGB("D4D0C8")
        End With
      Next
    End With
  Case 1 'arrow
    With scr_Arrow
      For i = 0 To .UBound
        With .Item(i)
          .BackColor = hexToRGB("000000")
        End With
      Next
    End With
  Case 2 'face
    With scr_Face
      For i = 0 To .UBound
        With .Item(i)
          .BackColor = hexToRGB("D4D0C8")
        End With
      Next
    End With
  Case 3 'shadow
    With scr_Shadow
      For i = 0 To .UBound
        With .Item(i)
          .BackColor = hexToRGB("808080")
        End With
      Next
    End With
  Case 4 'dark shadow
    With scr_Darkshadow
      For i = 0 To .UBound
        With .Item(i)
          .BackColor = hexToRGB("404040")
        End With
      Next
    End With
  Case 5 'track
    scr_Track(0).BackColor = RGB(0, 0, 0)
    For i = 0 To scr_TrackGrid.UBound - 1 Step 2
      If (i = 16) Or (i = 48) Or (i = 80) Or (i = 112) Then
        X = 1
      ElseIf (i = 32) Or (i = 64) Or (i = 96) Then
        X = 0
      End If
      scr_TrackGrid(i + X).BackColor = vb3DFace
    Next
    For i = 0 To scr_TrackGrid.UBound Step 2
      If (i = 16) Or (i = 48) Or (i = 80) Or (i = 112) Then
        X = 0
      ElseIf (i = 32) Or (i = 64) Or (i = 96) Then
        X = 1
      End If
      scr_TrackGrid(i + X).BackColor = vb3DHighlight
    Next
  Case 6 'highlight
    With scr_Highlight
      For i = 0 To .UBound
        With .Item(i)
          .BackColor = hexToRGB("FFFFFF")
        End With
      Next
    End With
  
End Select


End Sub

Private Sub cmdImport_Click()
On Error GoTo Ennd
CD.CancelError = True
CD.Filter = "HTML/CSS Files (*.htm; *.html; *.css; *.asp; *.php)|*.htm;*.html;*.css;*.asp;*.php|All files (*.*)|*.*"
CD.ShowOpen
If Len(CD.FileName) = 0 Then Exit Sub


MsgBox ReadCSS(sGetFile(CD.FileName)) & " of 7 possible colors imported."
Ennd:
End Sub

Function FindColorInText(sColor As String, sData As String)
Dim sTmp As String, sTClr As String
On Error GoTo nd
sData = LCase(sData)
sTmp = AfterFirst(sData, "scrollbar-" & sColor & "-color")
sTmp = AfterFirst(sTmp, ":")
sTmp = BeforeFirst(sTmp, ";")
sTmp = Trim(sTmp)
sTmp = Replace(sTmp, """", "")
sTmp = Replace(sTmp, "'", "")
sTmp = Replace(sTmp, "#", "")
If sTmp Like "[0123456789abcdef][0123456789abcdef][0123456789abcdef][0123456789abcdef][0123456789abcdef][0123456789abcdef]" Then FindColorInText = sTmp
nd:
End Function
Function ReadCSS(strData As String) As Integer
On Error GoTo Ennd
Dim sColor(0 To 6) As String, i As Integer, nFound As Integer

sColor(numFace) = FindColorInText("face", strData)
sColor(numHighlight) = FindColorInText("highlight", strData)
sColor(num3dLight) = FindColorInText("3dlight", strData)
sColor(numTrack) = FindColorInText("track", strData)
sColor(numArrow) = FindColorInText("arrow", strData)
sColor(numShadow) = FindColorInText("shadow", strData)
sColor(numDarkshadow) = FindColorInText("darkshadow", strData)

ResetAllColors

For i = 0 To 6
  If Len(sColor(i)) Then
    SetEditItem i
    lColor = hexToRGB(UCase(sColor(i)))
    cmdApply_Click
    nFound = nFound + 1
  End If
Next
Ennd:
ReadCSS = nFound
End Function
Function sGetFile(sFile)
On Error GoTo nd
Open sFile For Input As #1
Do While Not EOF(1)
  Line Input #1, InputData
  aData = aData & vbCrLf & InputData
Loop
Close #1
sGetFile = Right(aData, Len(aData) - 2)
Exit Function

nd:
sGetFile = ""
End Function

Private Sub cmdPreview_Click()
Dim sOut As String

sOut = "<html>" & vbCrLf
sOut = sOut & MakeCSS(True) & vbCrLf
sOut = sOut & "<body>" & vbCrLf
sOut = sOut & "<span style=""font-family:verdana,arial;font-size:14px;""><b>Scrollbar Designer Pro</b></span><br>"
sOut = sOut & "<span style=""font-family:tahoma,verdana,arial;font-size:11px;"">Preview page</span><br>"
sOut = sOut & "<br><br><br><br><br><br><br><br><br><br><br><br><br>"
sOut = sOut & "</body></html>"

Open sTempFile For Output As #1
Print #1, sOut
Close #1

frmPreview.Show
frmPreview.WebBrowser1.Navigate "file://" & sTempFile

End Sub

Private Sub cmdReset_Click()

If MsgBox("Reset all colors?", vbOKCancel + vbExclamation, "Reset") = vbOK Then
  ResetAllColors
End If

End Sub

Sub ResetAllColors()
Dim i As Integer
  For i = 0 To 6
    ResetColor i
    objColor(i) = "/"
  Next
  ShowTrack False
  RefreshTrack
  chkNoColor.Value = 1
End Sub
Private Sub cmdToCSS_Click()

  frmCode.Show
  frmCode.txtCode.Text = MakeCSS(True)

End Sub
Function MakeCSS(includeHTMLTag As Boolean)
Dim sOut As String
If includeHTMLTag Then sOut = "<STYLE TYPE=""text/css"">" & vbCrLf
sOut = sOut & "BODY " & vbTab & "{ " & vbCrLf
If objColor(numHighlight) <> "/" Then sOut = sOut & vbTab & "scrollbar-highlight-color:""#" & MakeHex(objColor(numHighlight)) & """;" & vbCrLf
If objColor(numFace) <> "/" Then sOut = sOut & vbTab & "scrollbar-face-color:""#" & MakeHex(objColor(numFace)) & """;" & vbCrLf
If objColor(num3dLight) <> "/" Then sOut = sOut & vbTab & "scrollbar-3dlight-color:""#" & MakeHex(objColor(num3dLight)) & """;" & vbCrLf
If objColor(numTrack) <> "/" Then sOut = sOut & vbTab & "scrollbar-track-color:""#" & MakeHex(objColor(numTrack)) & """;" & vbCrLf
If objColor(numShadow) <> "/" Then sOut = sOut & vbTab & "scrollbar-shadow-color:""#" & MakeHex(objColor(numShadow)) & """;" & vbCrLf
If objColor(numDarkshadow) <> "/" Then sOut = sOut & vbTab & "scrollbar-darkshadow-color:""#" & MakeHex(objColor(numDarkshadow)) & """;" & vbCrLf
If objColor(numArrow) <> "/" Then sOut = sOut & vbTab & "scrollbar-arrow-color:""#" & MakeHex(objColor(numArrow)) & """;" & vbCrLf
sOut = sOut & vbTab & "}" & vbCrLf
If includeHTMLTag Then sOut = sOut & "</STYLE>"

MakeCSS = sOut
End Function



Private Sub Form_Load()
Dim sToolTip As String, o As Integer
Dim i As Integer

For i = 0 To 6
  ResetColor i
  objColor(i) = "/"
Next

strCaption = "Scrollbar Designer Pro " & App.Major & "." & App.Minor
sTempFile = App.Path & "\scrdstmp.html"
sToolTip = "Leftclick to edit, rightclick to copy color"

Caption = strCaption

Set clrObject(0) = scr_Face
Set clrObject(1) = scr_Highlight
Set clrObject(2) = scr_3dlight
Set clrObject(3) = scr_Shadow
Set clrObject(4) = scr_Darkshadow
Set clrObject(5) = scr_Track
Set clrObject(6) = scr_Arrow
Set clrObject(7) = scr_TrackGrid

For o = 0 To 6
With clrObject(o)
  For i = 0 To .UBound
    With .Item(i)
      .ToolTipText = sToolTip
      .MousePointer = 2
    End With
  Next
End With
Next

chkAutoselect_Click

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Caption = strCaption
End Sub

Private Sub Form_Unload(Cancel As Integer)
CleanUp
End Sub

Private Sub optEdit_Click(Index As Integer)
If chkAutoselect.Value = 0 Then SetEditItem Index
End Sub

Private Sub picColors_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
clrSelect = True
picColors_MouseMove Button, Shift, X, Y
End Sub

Private Sub picColors_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If picColors.Point(X, Y) <> -1 Then
  If clrSelect Then
    lColor = picColors.Point(X, Y)
    RefreshColor
  End If
End If

Caption = strCaption
End Sub

Private Sub picColors_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
clrSelect = False
End Sub

Private Sub scr_3dlight_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
  lColor = scr_3dlight(Index).BackColor
  RefreshColor
Else
  optEdit(num3dLight).Value = True
End If
End Sub

Private Sub scr_3dlight_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Caption = scr_3dlight(Index).Tag
End Sub

Private Sub scr_Arrow_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
  lColor = scr_Arrow(Index).BackColor
  RefreshColor
Else
  optEdit(numArrow).Value = True
End If
End Sub

Private Sub scr_Arrow_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Caption = scr_Arrow(Index).Tag
End Sub

Private Sub scr_Darkshadow_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
  lColor = scr_Darkshadow(Index).BackColor
  RefreshColor
Else
  optEdit(numDarkshadow).Value = True
End If
End Sub

Private Sub scr_Darkshadow_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Caption = scr_Darkshadow(Index).Tag
End Sub

Private Sub scr_Face_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
  lColor = scr_Face(Index).BackColor
  RefreshColor
Else
  optEdit(numFace).Value = True
End If
End Sub

Private Sub scr_Face_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Caption = scr_Face(Index).Tag
End Sub

Private Sub scr_Highlight_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
  lColor = scr_Highlight(Index).BackColor
  RefreshColor
Else
  optEdit(numHighlight).Value = True
End If
End Sub

Private Sub scr_Highlight_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Caption = scr_Highlight(Index).Tag
End Sub

Private Sub scr_Shadow_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
  lColor = scr_Shadow(Index).BackColor
  RefreshColor
Else
  optEdit(numShadow).Value = True
End If
End Sub

Private Sub scr_Shadow_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Caption = scr_Shadow(Index).Tag
End Sub

Private Sub scr_Track_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
  lColor = scr_Track(Index).BackColor
  RefreshColor
Else
  optEdit(numTrack).Value = True
End If
End Sub

Private Sub scr_Track_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Caption = scr_Track(Index).Tag
End Sub

Private Sub scr_TrackGrid_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> vbRightButton Then optEdit(numTrack).Value = True
End Sub

Private Sub scr_TrackGrid_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Caption = scr_Track(0).Tag
End Sub

Private Sub scrHue_Change()
  If Not refColor Then refHSL
End Sub



Private Sub scrHue_Scroll()
scrHue_Change
End Sub

Private Sub scrLight_Change()
If Not refColor Then refHSL
End Sub

Private Sub scrLight_Scroll()
scrLight_Change
End Sub

Private Sub scrSat_Change()
If Not refColor Then refHSL
End Sub

Private Sub scrSat_Scroll()
scrSat_Change
End Sub

Private Sub txtBlue_Change()
If (Not refColor) And Len(txtBlue) Then refRGB
End Sub


Private Sub txtGreen_Change()
If (Not refColor) And Len(txtGreen) Then refRGB
End Sub

Private Sub txtHexColor_Change()
If Not refColor Then
If Len(txtHexColor.Text) = 6 Then
  lColor = hexToRGB(UCase(txtHexColor.Text))
  RefreshColor
End If
End If
End Sub

Private Sub txtHue_Change()
If (Not refColor) And (Len(txtHue)) Then refHSLText
End Sub

Private Sub txtLightness_Change()
If (Not refColor) And (Len(txtLightness)) Then refHSLText
End Sub

Private Sub txtRed_Change()
If (Not refColor) And Len(txtRed) Then refRGB
End Sub

Private Sub txtSaturation_Change()
If (Not refColor) And (Len(txtSaturation)) Then refHSLText
End Sub
