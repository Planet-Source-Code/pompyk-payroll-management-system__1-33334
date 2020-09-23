VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00404040&
   Caption         =   "EMPLOYEE INFO"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10725
   Icon            =   "payrollemp.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6075
   ScaleWidth      =   10725
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Height          =   1695
      Left            =   9000
      TabIndex        =   67
      Top             =   4320
      Width           =   1695
      Begin VB.CommandButton cmdlast 
         Caption         =   "last"
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
         Left            =   840
         TabIndex        =   71
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton cmdfirst 
         Caption         =   "first"
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
         TabIndex        =   70
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton cmdprevious 
         Caption         =   "previous"
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
         Left            =   360
         TabIndex        =   69
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   "next"
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
         Left            =   360
         TabIndex        =   68
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "CLOSE"
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
      Left            =   9240
      TabIndex        =   66
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "DELETE"
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
      Left            =   9240
      TabIndex        =   65
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "UPDATE"
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
      Left            =   9240
      TabIndex        =   64
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "ADD"
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
      Left            =   9240
      TabIndex        =   63
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ListBox lstinactive 
      Height          =   2010
      ItemData        =   "payrollemp.frx":030A
      Left            =   240
      List            =   "payrollemp.frx":030C
      TabIndex        =   60
      Top             =   3840
      Width           =   1575
   End
   Begin VB.ListBox lstactive 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2010
      ItemData        =   "payrollemp.frx":030E
      Left            =   240
      List            =   "payrollemp.frx":0310
      TabIndex        =   59
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtlname 
      Appearance      =   0  'Flat
      DataField       =   "lastname"
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
      Left            =   5520
      TabIndex        =   4
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtmname 
      Appearance      =   0  'Flat
      DataField       =   "middlename"
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
      Height          =   405
      Left            =   5520
      TabIndex        =   3
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtfname 
      Appearance      =   0  'Flat
      DataField       =   "firstname"
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
      Left            =   5520
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtempno 
      Appearance      =   0  'Flat
      DataField       =   "empno"
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
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6600
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "employee"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   2040
      TabIndex        =   0
      Top             =   2280
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   4210752
      TabCaption(0)   =   "GENERAL INFO"
      TabPicture(0)   =   "payrollemp.frx":0312
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(4)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(5)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(6)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(7)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(22)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(21)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtaddress"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtcity"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtstate"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtcountry"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtdob"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtspouse"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "CONTACTS"
      TabPicture(1)   =   "payrollemp.frx":032E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtemergency"
      Tab(1).Control(1)=   "txttelephone"
      Tab(1).Control(2)=   "txtemail"
      Tab(1).Control(3)=   "Label1(10)"
      Tab(1).Control(4)=   "Label1(9)"
      Tab(1).Control(5)=   "Label1(8)"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "ACADEMIC"
      TabPicture(2)   =   "payrollemp.frx":034A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtexperience"
      Tab(2).Control(1)=   "txteducation"
      Tab(2).Control(2)=   "Label1(12)"
      Tab(2).Control(3)=   "Label1(11)"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "COMPANY"
      TabPicture(3)   =   "payrollemp.frx":0366
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1(13)"
      Tab(3).Control(1)=   "Label1(14)"
      Tab(3).Control(2)=   "Label1(15)"
      Tab(3).Control(3)=   "Label1(16)"
      Tab(3).Control(4)=   "Label1(17)"
      Tab(3).Control(5)=   "Label1(18)"
      Tab(3).Control(6)=   "Label1(19)"
      Tab(3).Control(7)=   "txtdesignation"
      Tab(3).Control(8)=   "txtsalary"
      Tab(3).Control(9)=   "txtdept"
      Tab(3).Control(10)=   "txtnormal"
      Tab(3).Control(11)=   "txtsick"
      Tab(3).Control(12)=   "txtcasual"
      Tab(3).Control(13)=   "txthiredate"
      Tab(3).ControlCount=   14
      TabCaption(4)   =   "STATUS"
      TabPicture(4)   =   "payrollemp.frx":0382
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtstatus"
      Tab(4).Control(1)=   "Label1(25)"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "TERMINATION"
      TabPicture(5)   =   "payrollemp.frx":039E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txtterminate"
      Tab(5).Control(1)=   "txtdot"
      Tab(5).Control(2)=   "Label1(24)"
      Tab(5).Control(3)=   "Label1(23)"
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "LEAVE INFO"
      TabPicture(6)   =   "payrollemp.frx":03BA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "txtsickid"
      Tab(6).Control(1)=   "txtcasualid"
      Tab(6).Control(2)=   "txtnormalid"
      Tab(6).Control(3)=   "Label1(28)"
      Tab(6).Control(4)=   "Label1(27)"
      Tab(6).Control(5)=   "Label1(26)"
      Tab(6).ControlCount=   6
      TabCaption(7)   =   "NOTE"
      TabPicture(7)   =   "payrollemp.frx":03D6
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "txtnote"
      Tab(7).Control(1)=   "Label1(20)"
      Tab(7).ControlCount=   2
      Begin VB.TextBox txtsickid 
         Appearance      =   0  'Flat
         DataField       =   "sickleaveid"
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
         Left            =   -72240
         TabIndex        =   55
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtcasualid 
         Appearance      =   0  'Flat
         DataField       =   "casualleaveid"
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
         Left            =   -72240
         TabIndex        =   54
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtnormalid 
         Appearance      =   0  'Flat
         DataField       =   "normalleaveid"
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
         Left            =   -72240
         TabIndex        =   53
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtspouse 
         Appearance      =   0  'Flat
         DataField       =   "spouse"
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
         Left            =   2760
         TabIndex        =   50
         Top             =   2760
         Width           =   2055
      End
      Begin VB.TextBox txtdob 
         Appearance      =   0  'Flat
         DataField       =   "dob"
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
         Left            =   2760
         TabIndex        =   49
         Top             =   3120
         Width           =   2055
      End
      Begin VB.TextBox txthiredate 
         Appearance      =   0  'Flat
         DataField       =   "doj"
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
         Left            =   -72600
         TabIndex        =   37
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtstatus 
         Appearance      =   0  'Flat
         DataField       =   "status"
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
         Left            =   -72720
         TabIndex        =   23
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox txtterminate 
         Appearance      =   0  'Flat
         DataField       =   "terminationreason"
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
         Height          =   1335
         Left            =   -72720
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox txtdot 
         Appearance      =   0  'Flat
         DataField       =   "dot"
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
         Left            =   -72720
         TabIndex        =   21
         Top             =   1440
         Width           =   2415
      End
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
         Height          =   2295
         Left            =   -73320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   960
         Width           =   4575
      End
      Begin VB.TextBox txtcasual 
         Appearance      =   0  'Flat
         DataField       =   "totalcasual"
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
         Left            =   -72600
         TabIndex        =   19
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox txtsick 
         Appearance      =   0  'Flat
         DataField       =   "totalsick"
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
         Left            =   -72600
         TabIndex        =   18
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox txtnormal 
         Appearance      =   0  'Flat
         DataField       =   "totalnormal"
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
         Left            =   -72600
         TabIndex        =   17
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox txtdept 
         Appearance      =   0  'Flat
         DataField       =   "department"
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
         Left            =   -72600
         TabIndex        =   16
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox txtsalary 
         Appearance      =   0  'Flat
         DataField       =   "salary"
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
         Left            =   -72600
         TabIndex        =   15
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtdesignation 
         Appearance      =   0  'Flat
         DataField       =   "designation"
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
         Left            =   -72600
         TabIndex        =   14
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtexperience 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "experience"
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
         Height          =   975
         Left            =   -72720
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   2040
         Width           =   3015
      End
      Begin VB.TextBox txteducation 
         Appearance      =   0  'Flat
         DataField       =   "education"
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
         Height          =   975
         Left            =   -72720
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txtemergency 
         Appearance      =   0  'Flat
         DataField       =   "emergencyphone"
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
         Left            =   -72360
         TabIndex        =   11
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox txttelephone 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "telephone"
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
         Left            =   -72360
         TabIndex        =   10
         Top             =   1680
         Width           =   2535
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
         Left            =   -72360
         TabIndex        =   9
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtcountry 
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
         Left            =   2760
         TabIndex        =   8
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox txtstate 
         Appearance      =   0  'Flat
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
         Left            =   2760
         TabIndex        =   7
         Top             =   2040
         Width           =   2055
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
         Left            =   2760
         TabIndex        =   6
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtaddress 
         Appearance      =   0  'Flat
         DataField       =   "address"
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
         Height          =   975
         Left            =   2760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SICK LEAVE ID:"
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
         Index           =   28
         Left            =   -74400
         TabIndex        =   58
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CASUAL LEAVE ID:"
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
         Index           =   27
         Left            =   -74400
         TabIndex        =   57
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "NORMAL LEAVE ID:"
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
         Index           =   26
         Left            =   -74400
         TabIndex        =   56
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SPOUSE:"
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
         Index           =   21
         Left            =   840
         TabIndex        =   52
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DATE OF BIRTH:"
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
         Index           =   22
         Left            =   840
         TabIndex        =   51
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "STATUS:"
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
         Index           =   25
         Left            =   -74640
         TabIndex        =   48
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "TERMINATION REASON:"
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
         Index           =   24
         Left            =   -74760
         TabIndex        =   47
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "TERMINATION DATE:"
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
         Index           =   23
         Left            =   -74760
         TabIndex        =   46
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "NOTE:"
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
         Index           =   20
         Left            =   -74880
         TabIndex        =   45
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL CASUAL LEAVES:"
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
         Index           =   19
         Left            =   -74760
         TabIndex        =   44
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL SICK LEAVES:"
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
         Index           =   18
         Left            =   -74760
         TabIndex        =   43
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL NORMAL LEAVES:"
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
         Index           =   17
         Left            =   -74760
         TabIndex        =   42
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DEPARTMENT:"
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
         Index           =   16
         Left            =   -74760
         TabIndex        =   41
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SALARY:"
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
         Index           =   15
         Left            =   -74760
         TabIndex        =   40
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DESIGNATION:"
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
         Index           =   14
         Left            =   -74760
         TabIndex        =   39
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "HIRE DATE:"
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
         Index           =   13
         Left            =   -74760
         TabIndex        =   38
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "EXPERIENCE:"
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
         Index           =   12
         Left            =   -74400
         TabIndex        =   36
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "EDUCATION:"
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
         Index           =   11
         Left            =   -74400
         TabIndex        =   35
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "EMERGENCY:"
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
         Index           =   10
         Left            =   -74040
         TabIndex        =   34
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "TELEPHONE:"
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
         Index           =   9
         Left            =   -74040
         TabIndex        =   33
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "E-MAIL:"
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
         Index           =   8
         Left            =   -74040
         TabIndex        =   32
         Top             =   1440
         Width           =   1335
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
         Index           =   7
         Left            =   840
         TabIndex        =   31
         Top             =   2520
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
         Index           =   6
         Left            =   840
         TabIndex        =   30
         Top             =   2160
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
         Index           =   5
         Left            =   840
         TabIndex        =   29
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS:"
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
         Left            =   840
         TabIndex        =   28
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "INACTIVE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   62
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ACTIVE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   61
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "LAST NAME:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   4080
      TabIndex        =   27
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MIDDLE NAME:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   26
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FIRST NAME:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   25
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "EMP NO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   24
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadd_Click()
On Error GoTo x
Adodc1.Recordset.AddNew
x:
End Sub

Private Sub cmdclose_Click()
Me.Visible = False

End Sub

Private Sub cmddelete_Click()
On Error GoTo x
Adodc1.Recordset.Delete
Adodc1.Recordset.Update
txtaddress = ""
txtcasual = ""
txtcasualid = ""
txtcity = ""
txtcountry = ""
txtdept = ""
txtdesignation = ""
txtdob = ""
txtdot = ""
txteducation = ""
txtemail = ""
txtemergency = ""
txtempno = ""
txtexperience = ""
txtfname = ""
txthiredate = ""
txtlname = ""
txtmname = ""
txtnormal = ""
txtnormalid = ""
txtnote = ""
txtsalary = ""
txtsick = ""
txtsickid = ""
txtspouse = ""
txtstate = ""
txtstatus = ""
txttelephone = ""
txtterminate = ""
lstactive.Clear
lstinactive.Clear
Adodc1.Refresh
Form_Load
x:
End Sub

Private Sub cmdfirst_Click()
On Error Resume Next

Adodc1.Recordset.MoveFirst
End Sub

Private Sub cmdlast_Click()
On Error Resume Next

Adodc1.Recordset.MoveLast

End Sub

Private Sub cmdnext_Click()
On Error Resume Next
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
Adodc1.Recordset.MoveLast

End If
End Sub

Private Sub cmdprevious_Click()
On Error Resume Next
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF = True Then
Adodc1.Recordset.MoveFirst

End If
End Sub

Private Sub cmdupdate_Click()
On Error GoTo cancelupdate
Adodc1.Recordset.Update
lstactive.Clear
lstinactive.Clear
Adodc1.Refresh
Form_Load
Exit Sub
cancelupdate:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
Form2.Width = 10845
Form2.Height = 6480
On Error GoTo x
Dim i As Integer
Adodc1.Refresh
Adodc1.Recordset.MoveLast
Adodc1.Recordset.MoveFirst
For i = 1 To Adodc1.Recordset.RecordCount
lstactive.AddItem Adodc1.Recordset.Fields(2)
Adodc1.Recordset.MoveNext
Next
Adodc1.Recordset.MoveFirst
x:
End Sub

Private Sub lstactive_Click()
Adodc1.Refresh
Dim l As String
Dim no As Integer
Dim count As Integer
count = 0
no = lstactive.ListIndex
'l = lstactive.Text
'Adodc1.Recordset.Find "firstname='" & l & "' "
Adodc1.Recordset.MoveFirst
While Adodc1.Recordset.EOF <> True
If count = no Then
Exit Sub
End If
count = count + 1
Adodc1.Recordset.MoveNext
Wend

End Sub
