VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "COMPANY INFO"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   Icon            =   "payrollmng1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5565
   ScaleWidth      =   6735
   Begin VB.CommandButton cmdclose 
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   5640
      TabIndex        =   27
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmdmodify 
      Caption         =   "MODIFY"
      Height          =   375
      Left            =   3480
      TabIndex        =   26
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "DELETE"
      Height          =   375
      Left            =   2400
      TabIndex        =   25
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "UPDATE"
      Height          =   375
      Left            =   1320
      TabIndex        =   24
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "ADD"
      Height          =   375
      Left            =   480
      TabIndex        =   23
      Top             =   4680
      Width           =   735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7435
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   4210752
      TabCaption(0)   =   "GENERAL"
      TabPicture(0)   =   "payrollmng1.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtcname"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtadd1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtadd2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtcity"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtstate"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtcontry"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "CONTACTS"
      TabPicture(1)   =   "payrollmng1.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtwebsite"
      Tab(1).Control(1)=   "txtemail"
      Tab(1).Control(2)=   "txtfax"
      Tab(1).Control(3)=   "txtphone"
      Tab(1).Control(4)=   "Label2(3)"
      Tab(1).Control(5)=   "Label2(2)"
      Tab(1).Control(6)=   "Label2(1)"
      Tab(1).Control(7)=   "Label2(0)"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "NOTE"
      TabPicture(2)   =   "payrollmng1.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2(4)"
      Tab(2).Control(1)=   "txtnote"
      Tab(2).ControlCount=   2
      Begin VB.TextBox txtnote 
         Appearance      =   0  'Flat
         DataField       =   "note"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   -73800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   720
         Width           =   4455
      End
      Begin VB.TextBox txtwebsite 
         Appearance      =   0  'Flat
         DataField       =   "website"
         DataSource      =   "Adodc1"
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
         Left            =   -72840
         TabIndex        =   16
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtemail 
         Appearance      =   0  'Flat
         DataField       =   "email"
         DataSource      =   "Adodc1"
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
         Left            =   -72840
         TabIndex        =   15
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtfax 
         Appearance      =   0  'Flat
         DataField       =   "fax"
         DataSource      =   "Adodc1"
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
         Left            =   -72840
         TabIndex        =   14
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtphone 
         Appearance      =   0  'Flat
         DataField       =   "phone"
         DataSource      =   "Adodc1"
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
         Left            =   -72840
         TabIndex        =   13
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtcontry 
         Appearance      =   0  'Flat
         DataField       =   "country"
         DataSource      =   "Adodc1"
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
         TabIndex        =   11
         Top             =   3720
         Width           =   2175
      End
      Begin VB.TextBox txtstate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "state"
         DataSource      =   "Adodc1"
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
         TabIndex        =   5
         Top             =   3240
         Width           =   2175
      End
      Begin VB.TextBox txtcity 
         Appearance      =   0  'Flat
         DataField       =   "city"
         DataSource      =   "Adodc1"
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
         TabIndex        =   4
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox txtadd2 
         Appearance      =   0  'Flat
         DataField       =   "address2"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox txtadd1 
         Appearance      =   0  'Flat
         DataField       =   "address1"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtcname 
         Appearance      =   0  'Flat
         DataField       =   "companyname"
         DataSource      =   "Adodc1"
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
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "NOTE:"
         Height          =   375
         Index           =   4
         Left            =   -74760
         TabIndex        =   22
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "WEBSITE:"
         Height          =   375
         Index           =   3
         Left            =   -74760
         TabIndex        =   20
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "E-MAIL:"
         Height          =   375
         Index           =   2
         Left            =   -74760
         TabIndex        =   19
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "FAX:"
         Height          =   375
         Index           =   1
         Left            =   -74760
         TabIndex        =   18
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "PHONE:"
         Height          =   375
         Index           =   0
         Left            =   -74760
         TabIndex        =   17
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "COUNTRY:"
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
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "STATE:"
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
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CITY:"
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
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS 2:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS 1:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "COMPANY NAME:"
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
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1800
      Top             =   5160
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=payroll.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=payroll.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "companyinfo"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
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
Private Sub cmdadd_Click()

If Adodc1.Recordset.RecordCount = 1 Then
MsgBox "u have already added the details of your company..u can update them by pressing update button"
Exit Sub
End If
Adodc1.Refresh
Adodc1.Recordset.AddNew

End Sub

Private Sub cmdclose_Click()
Me.Visible = False

End Sub

Private Sub cmddelete_Click()
On Error Resume Next
Adodc1.Refresh
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "db is empty"
End If
If Adodc1.Recordset.RecordCount = 1 Then
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Delete
Adodc1.Recordset.Update

txtadd1 = ""
txtadd2 = ""
txtcity = ""
txtcname = ""
txtcontry = ""
txtemail = ""
txtfax = ""
txtnote = ""
txtphone = ""
txtstate = ""
txtwebsite = ""
End If

End Sub

Private Sub cmdmodify_Click()
On Error Resume Next
Adodc1.Recordset.Update

End Sub

Private Sub cmdupdate_Click()
On Error Resume Next
Adodc1.Recordset.Update

End Sub

Private Sub Form_Load()
Form1.Width = 6855
Form1.Height = 5970
End Sub
