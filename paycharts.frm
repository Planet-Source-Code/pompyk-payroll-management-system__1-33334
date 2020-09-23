VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form10 
   BackColor       =   &H00404040&
   Caption         =   " With MSChart1.Plot.Backdrop"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "paycharts.frx":0000
   LinkTopic       =   "Form10"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   Begin VB.CommandButton cmdexperience 
      Caption         =   "EXP"
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
      Left            =   10680
      TabIndex        =   17
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdages 
      Caption         =   "AGE"
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
      Left            =   10680
      TabIndex        =   16
      Top             =   600
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   3480
   End
   Begin VB.CommandButton cmdstop 
      BackColor       =   &H00808080&
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmdautorotate 
      BackColor       =   &H00808080&
      Caption         =   "Auto Rotate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4200
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   3480
   End
   Begin VB.CommandButton cmdsolid 
      Caption         =   "Hatched"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdsolid 
      Caption         =   "Solid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdborder 
      Caption         =   "Without Border"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdborder 
      Caption         =   "With Border"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdshadow 
      Caption         =   "Without Shadow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdshadow 
      Caption         =   "With Shadow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdch 
      Caption         =   "3d Pie"
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
      Left            =   2160
      TabIndex        =   7
      Top             =   6960
      Width           =   855
   End
   Begin VB.CommandButton cmdch 
      Caption         =   "2d Pie"
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
      Left            =   2160
      TabIndex        =   6
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton cmdch 
      Caption         =   "3d Line"
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
      Left            =   1200
      TabIndex        =   5
      Top             =   6960
      Width           =   855
   End
   Begin VB.CommandButton cmdch 
      Caption         =   "2d Line"
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
      Left            =   1200
      TabIndex        =   4
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton cmdch 
      Caption         =   "3d Bar"
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
      Left            =   240
      TabIndex        =   3
      Top             =   6960
      Width           =   855
   End
   Begin VB.CommandButton cmdch 
      Caption         =   "2d Bar"
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
      TabIndex        =   2
      Top             =   6600
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   7800
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
   Begin VB.CommandButton cmdage 
      Caption         =   "SALARY"
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
      Left            =   10680
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin MSChart20Lib.MSChart chartman 
      Bindings        =   "paycharts.frx":030A
      Height          =   6255
      Left            =   1080
      OleObjectBlob   =   "paycharts.frx":031F
      TabIndex        =   0
      Top             =   240
      Width           =   8895
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ii As Double
Dim e As Double
Dim r As Double
Dim iii As Double


Private Sub cmdage_Click()
'chartman.DataSource = Adodc1
'chartman.Legend = True
Adodc1.RecordSource = "select empno,firstname, salary from employee"
Adodc1.Refresh
End Sub

Private Sub cmdages_Click()
Adodc1.RecordSource = "select empno,firstname, (Date()-dob)/365 from employee"
Adodc1.Refresh
End Sub

Private Sub cmdautorotate_Click()
e = chartman.Plot.View3d.Elevation
r = chartman.Plot.View3d.Rotation
Timer1.Enabled = True
End Sub



Private Sub cmdborder_Click(Index As Integer)
With chartman.Plot.Backdrop
    Select Case Index
      Case 0
        .Frame.Style = VtFrameStyleThickOuter
      ' Set style to show a shadow.
      Case 1
        .Frame.Style = VtFrameStyleNull
'        .Shadow.Style = VtShadowStyleDrop
    End Select
  End With
End Sub

Private Sub cmdch_Click(Index As Integer)
'optdef.Value = False
  'chartman.ShowLegend = True
  Select Case Index
    Case 0
      chartman.chartType = VtChChartType2dBar
    Case 1
      chartman.chartType = VtChChartType3dBar
    Case 2
      chartman.chartType = VtChChartType2dPie
    Case 3
     chartman.chartType = VtChChartType2dLine
    Case 4
      chartman.chartType = VtChChartType3dLine
    Case 5
      chartman.chartType = VtChChartType3dStep
  End Select
End Sub





Private Sub cmdexperience_Click()
Adodc1.RecordSource = "select empno,firstname, (Date()-doj)/365 from employee"
Adodc1.Refresh
End Sub

Private Sub cmdshadow_Click(Index As Integer)
With chartman.Plot.Backdrop
    Select Case Index
      Case 0
        .Shadow.Style = VtShadowStyleDrop
      ' Set style to show a shadow.
      Case 1
        .Shadow.Style = VtShadowStyleNull
    End Select
  End With
End Sub

Private Sub cmdsolid_Click(Index As Integer)
 With chartman.Plot
    Select Case Index
      Case 0
        ' Set the style to solid.
        .Wall.Brush.Style = VtBrushStyleHatched
      Case 1
        .Wall.Brush.Style = VtBrushStyleSolid
    End Select
    ' Set the color to white.
    .Wall.Brush.FillColor.Set 255, 255, 255
  End With
End Sub

Private Sub cmdstop_Click()
Timer1.Enabled = False
Timer2.Enabled = False
e = chartman.Plot.View3d.Elevation
r = chartman.Plot.View3d.Rotation
ii = 1
iii = 1
End Sub

Private Sub Timer1_Timer()
chartman.Plot.View3d.Set r + ii, e
  ii = ii + 1
End Sub

Private Sub Timer2_Timer()
chartman.Plot.View3d.Set r, e + iii
  iii = iii + 1
End Sub
