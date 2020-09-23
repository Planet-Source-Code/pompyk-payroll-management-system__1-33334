VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form7 
   BackColor       =   &H00404040&
   Caption         =   "query form"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "payrollquery.frx":0000
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   Begin VB.CheckBox chkwhere 
      BackColor       =   &H00404040&
      Caption         =   "Enable condition"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   8040
      TabIndex        =   47
      Top             =   240
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5520
      Top             =   1080
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      RecordSource    =   "select * from employee"
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
   Begin VB.CommandButton cmdquery 
      Caption         =   "QUERY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9960
      TabIndex        =   45
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtdept 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   44
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtsalary 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   42
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtdoj 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   40
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtdob 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   38
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtstate 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   36
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtcity 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   34
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtfname 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   32
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtlname 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   30
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "TICK THOSE WHICH R TO BE DISPLAYED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   11535
      Begin VB.CheckBox chkall 
         BackColor       =   &H00808080&
         Caption         =   "all"
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
         Left            =   8520
         TabIndex        =   46
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CheckBox Check33 
         BackColor       =   &H00808080&
         Caption         =   "terminationreason"
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
         Left            =   6720
         TabIndex        =   28
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox Check32 
         BackColor       =   &H00808080&
         Caption         =   "dot"
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
         Left            =   6720
         TabIndex        =   27
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox Check31 
         BackColor       =   &H00808080&
         Caption         =   "totalcasual"
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
         Left            =   6720
         TabIndex        =   26
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox Check30 
         BackColor       =   &H00808080&
         Caption         =   "totalsick"
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
         Left            =   5280
         TabIndex        =   25
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox Check29 
         BackColor       =   &H00808080&
         Caption         =   "totalnormal"
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
         Left            =   5280
         TabIndex        =   24
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox Check28 
         BackColor       =   &H00808080&
         Caption         =   "status"
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
         Left            =   5280
         TabIndex        =   23
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox Check27 
         BackColor       =   &H00808080&
         Caption         =   "casualleaveid"
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
         Left            =   5280
         TabIndex        =   22
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox Check26 
         BackColor       =   &H00808080&
         Caption         =   "sickleaveid"
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
         Left            =   5280
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox Check25 
         BackColor       =   &H00808080&
         Caption         =   "normalleaveid"
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
         Left            =   3840
         TabIndex        =   20
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox Check24 
         BackColor       =   &H00808080&
         Caption         =   "department"
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
         Left            =   3840
         TabIndex        =   19
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox Check23 
         BackColor       =   &H00808080&
         Caption         =   "salary"
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
         Left            =   3840
         TabIndex        =   18
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox Check22 
         BackColor       =   &H00808080&
         Caption         =   "designation"
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
         Left            =   3840
         TabIndex        =   17
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox Check21 
         BackColor       =   &H00808080&
         Caption         =   "doj"
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
         Left            =   3840
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox Check20 
         BackColor       =   &H00808080&
         Caption         =   "dob"
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
         Left            =   2520
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CheckBox Check19 
         BackColor       =   &H00808080&
         Caption         =   "spouse"
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
         Left            =   2520
         TabIndex        =   14
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox Check18 
         BackColor       =   &H00808080&
         Caption         =   "experience"
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
         Left            =   2520
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox Check17 
         BackColor       =   &H00808080&
         Caption         =   "education"
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
         Left            =   2520
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox Check16 
         BackColor       =   &H00808080&
         Caption         =   "note"
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
         Left            =   2520
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox Check15 
         BackColor       =   &H00808080&
         Caption         =   "emergency"
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
         Left            =   1320
         TabIndex        =   10
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox Check14 
         BackColor       =   &H00808080&
         Caption         =   "telephone"
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
         Left            =   1320
         TabIndex        =   9
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox Check13 
         BackColor       =   &H00808080&
         Caption         =   "email"
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
         Left            =   1320
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox Check12 
         BackColor       =   &H00808080&
         Caption         =   "country"
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
         Left            =   1320
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox Check11 
         BackColor       =   &H00808080&
         Caption         =   "state"
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
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox Check10 
         BackColor       =   &H00808080&
         Caption         =   "city"
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
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox Check9 
         BackColor       =   &H00808080&
         Caption         =   "address"
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
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H00808080&
         Caption         =   "firstname"
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
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00808080&
         Caption         =   "lastname"
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
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "payrollquery.frx":030A
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7011
      _Version        =   393216
      BackColor       =   16512
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "specify =,<= or >= sign in case of salary, dob and doj    EG: >=1200"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   615
      Left            =   6720
      TabIndex        =   48
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "dept:"
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
      Index           =   7
      Left            =   5520
      TabIndex        =   43
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "salary:"
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
      Index           =   6
      Left            =   5520
      TabIndex        =   41
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "doj:"
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
      Index           =   5
      Left            =   3000
      TabIndex        =   39
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "dob:"
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
      Index           =   4
      Left            =   3000
      TabIndex        =   37
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "state:"
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
      Index           =   3
      Left            =   3000
      TabIndex        =   35
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "city:"
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
      Index           =   2
      Left            =   240
      TabIndex        =   33
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "first name:"
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
      TabIndex        =   31
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "last name:"
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
      TabIndex        =   29
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check10_Click()
If Check10.Value = 1 Then
Check10.Tag = "," & Check10.Caption
Else
Check10.Tag = ""
End If
End Sub

Private Sub Check11_Click()
If Check11.Value = 1 Then
Check11.Tag = "," & Check11.Caption
Else
Check11.Tag = ""
End If
End Sub

Private Sub Check12_Click()
If Check12.Value = 1 Then
Check12.Tag = "," & Check12.Caption
Else
Check12.Tag = ""
End If
End Sub

Private Sub Check13_Click()
If Check13.Value = 1 Then
Check13.Tag = "," & Check13.Caption
Else
Check13.Tag = ""
End If
End Sub

Private Sub Check14_Click()
If Check14.Value = 1 Then
Check14.Tag = "," & Check14.Caption
Else
Check14.Tag = ""
End If
End Sub

Private Sub Check15_Click()
If Check15.Value = 1 Then
Check15.Tag = "," & Check15.Caption & "phone"

Else
Check15.Tag = ""
End If
End Sub

Private Sub Check16_Click()
If Check16.Value = 1 Then
Check16.Tag = "," & Check16.Caption
Else
Check16.Tag = ""
End If
End Sub

Private Sub Check17_Click()
If Check17.Value = 1 Then
Check17.Tag = "," & Check17.Caption
Else
Check17.Tag = ""
End If
End Sub

Private Sub Check18_Click()
If Check18.Value = 1 Then
Check18.Tag = "," & Check18.Caption
Else
Check18.Tag = ""
End If
End Sub

Private Sub Check19_Click()
If Check19.Value = 1 Then
Check19.Tag = "," & Check19.Caption
Else
Check19.Tag = ""
End If
End Sub

Private Sub Check20_Click()
If Check20.Value = 1 Then
Check20.Tag = "," & Check20.Caption
Else
Check20.Tag = ""
End If
End Sub

Private Sub Check21_Click()
If Check21.Value = 1 Then
Check21.Tag = "," & Check21.Caption
Else
Check21.Tag = ""
End If
End Sub

Private Sub Check22_Click()
If Check22.Value = 1 Then
Check22.Tag = "," & Check22.Caption
Else
Check22.Tag = ""
End If
End Sub

Private Sub Check23_Click()
If Check23.Value = 1 Then
Check23.Tag = "," & Check23.Caption
Else
Check23.Tag = ""
End If
End Sub

Private Sub Check24_Click()
If Check24.Value = 1 Then
Check24.Tag = "," & Check24.Caption
Else
Check24.Tag = ""
End If
End Sub

Private Sub Check25_Click()
If Check25.Value = 1 Then
Check25.Tag = "," & Check25.Caption
Else
Check25.Tag = ""
End If
End Sub

Private Sub Check26_Click()
If Check26.Value = 1 Then
Check26.Tag = "," & Check26.Caption
Else
Check26.Tag = ""
End If
End Sub

Private Sub Check27_Click()
If Check27.Value = 1 Then
Check27.Tag = "," & Check27.Caption
Else
Check27.Tag = ""
End If
End Sub

Private Sub Check28_Click()
If Check28.Value = 1 Then
Check28.Tag = "," & Check28.Caption
Else
Check28.Tag = ""
End If
End Sub

Private Sub Check29_Click()
If Check29.Value = 1 Then
Check29.Tag = "," & Check29.Caption
Else
Check29.Tag = ""
End If
End Sub

Private Sub Check30_Click()
If Check30.Value = 1 Then
Check30.Tag = "," & Check30.Caption
Else
Check30.Tag = ""
End If
End Sub

Private Sub Check31_Click()
If Check31.Value = 1 Then
Check31.Tag = "," & Check31.Caption
Else
Check31.Tag = ""
End If
End Sub

Private Sub Check32_Click()
If Check32.Value = 1 Then
Check32.Tag = "," & Check32.Caption
Else
Check32.Tag = ""
End If
End Sub

Private Sub Check33_Click()
If Check33.Value = 1 Then
Check33.Tag = "," & Check33.Caption
Else
Check33.Tag = ""
End If
End Sub

Private Sub Check7_Click()
If Check7.Value = 1 Then
Check7.Tag = "," & Check7.Caption

Else
Check7.Tag = ""
End If
End Sub

Private Sub Check8_Click()
If Check8.Value = 1 Then
Check8.Tag = "," & Check8.Caption
Else
Check8.Tag = ""
End If
End Sub

Private Sub Check9_Click()
If Check9.Value = 1 Then
Check9.Tag = "," & Check9.Caption
Else
Check9.Tag = ""
End If
End Sub

Private Sub chkall_Click()
If chkall.Value = 1 Then
chkall.Tag = " * "
Else
chkall.Tag = ""
End If
End Sub

Private Sub cmdquery_Click()
On Error GoTo x
If chkall.Value = 1 Then
sqlbaby = " select * "
Else
sqlbaby = sqlbaby & Check7.Tag & Check8.Tag & Check9.Tag & Check10.Tag & Check11.Tag & Check12.Tag & Check13.Tag & Check14.Tag & Check15.Tag & Check16.Tag & Check17.Tag & Check18.Tag & Check19.Tag & Check20.Tag & Check21.Tag & Check22.Tag & Check23.Tag & Check24.Tag & Check25.Tag & Check26.Tag & Check27.Tag & Check28.Tag & Check29.Tag & Check30.Tag & Check31.Tag & Check32.Tag & Check33.Tag
End If

sqlbaby = sqlbaby & " from employee"
If chkwhere.Value = 1 Then
sqlbaby = sqlbaby & " where "
sqlbaby = sqlbaby & "lastname='" & txtlname.Text & "'"
sqlbaby = sqlbaby & " or firstname='" & txtfname.Text & "'"
sqlbaby = sqlbaby & " or city='" & txtcity.Text & "'"
sqlbaby = sqlbaby & " or state='" & txtstate.Text & "'"
sqlbaby = sqlbaby & " or department='" & txtdept.Text & "'"

If txtsalary.Text <> "" Then
sqlbaby = sqlbaby & " or salary" & txtsalary.Text
End If
If txtdob.Text <> "" Then
sqlbaby = sqlbaby & " or dob" & txtdob.Text
End If
If txtdoj.Text <> "" Then
sqlbaby = sqlbaby & " or doj" & txtdoj.Text
End If
End If
Adodc1.RecordSource = sqlbaby

Adodc1.Refresh
Form_Load
Exit Sub
x:
MsgBox Err.Description

End Sub

Private Sub Form_Load()
On Error Resume Next
Form7.Width = 12000
Form7.Height = 9000
sqlbaby = "Select empno "
chkwhere.Value = 0

End Sub
