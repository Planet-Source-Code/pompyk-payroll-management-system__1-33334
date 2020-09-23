VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00404040&
   Caption         =   "PAYROLL SYSTEM"
   ClientHeight    =   6420
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10455
   Icon            =   "payrollmng.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   10425
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      Begin VB.CommandButton cmdcharts 
         Caption         =   "charts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   7
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8640
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmddetails 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9960
         TabIndex        =   5
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdpayroll 
         Caption         =   "payroll"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   4
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdleaveinfo 
         Caption         =   "leave info"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "employee info"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdcompanyinforma 
         Caption         =   "company info"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Menu popmeup 
      Caption         =   "popme"
      Begin VB.Menu query 
         Caption         =   "query employees"
         Index           =   0
      End
      Begin VB.Menu query 
         Caption         =   "direct queries employees"
         Index           =   1
      End
      Begin VB.Menu query 
         Caption         =   "direct queries for payroll"
         Index           =   2
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'email: gangulysomdutt@yahoo.com
'address: no 6,chandrodaya apt,bhaikaka nagar
'thaltej, ahmedabad, gujarat, india - 380059
'year: 2001-2002
'status: TY BCA from CPICA College - Gujarat university
'note: plz don't modify the source code...and pretend that
'u r the author since this is not a good practice for
'any programmer.....i appreciate if u make this program
'better.
'i appreciate feed backs...thx
Private Sub cmdcharts_Click()
Form10.Visible = True

End Sub

Private Sub cmdcompanyinforma_Click()

Form1.Visible = True

End Sub

Private Sub cmddetails_Click()
PopupMenu popmeup
End Sub

Private Sub cmdexit_Click()
Unload Me
End

End Sub

Private Sub cmdleaveinfo_Click()
Form3.Visible = True

End Sub

Private Sub cmdpayroll_Click()
Form5.Visible = True

End Sub

Private Sub Command1_Click()
Form2.Visible = True

End Sub

Private Sub MDIForm_Load()
Form1.Visible = True

End Sub

Private Sub query_Click(Index As Integer)
If Index = 0 Then

Form7.Visible = True
End If
If Index = 1 Then
Form8.Visible = True
End If
If Index = 2 Then
Form9.Visible = True
End If

End Sub
