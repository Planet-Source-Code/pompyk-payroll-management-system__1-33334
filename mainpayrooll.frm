VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00404040&
   Caption         =   "payroll info"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   DrawStyle       =   1  'Dash
   Icon            =   "mainpayrooll.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "mainpayrooll.frx":030A
   ScaleHeight     =   5685
   ScaleWidth      =   5775
   Begin VB.CommandButton Command2 
      Caption         =   "DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3480
      TabIndex        =   30
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   28
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdtotal 
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   27
      Top             =   4200
      Width           =   375
   End
   Begin VB.CommandButton cmdrebateothertax 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   25
      ToolTipText     =   "grant rebate due to some policies"
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton cmdrebatetax 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   24
      ToolTipText     =   "grantrebate due to some policies"
      Top             =   2760
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   1320
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      RecordSource    =   "chequedetail"
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   5160
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
   Begin VB.TextBox txtleave 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc2"
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
      TabIndex        =   23
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdleaveaddition 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   22
      Top             =   2760
      Width           =   375
   End
   Begin VB.CommandButton cmdpay 
      Caption         =   "GRANT PAY"
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
      TabIndex        =   21
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox txtnet 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc2"
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
      TabIndex        =   12
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtothertax 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc2"
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
      TabIndex        =   11
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txttax 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc2"
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
      TabIndex        =   10
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtchequeno 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc2"
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
      TabIndex        =   9
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txtpaymode 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc2"
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
      TabIndex        =   8
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox txtdop 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc2"
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
      TabIndex        =   7
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtsalary 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc2"
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
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtempno 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc2"
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
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
   Begin VB.ListBox lstunpaid 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2010
      ItemData        =   "mainpayrooll.frx":0F4C
      Left            =   120
      List            =   "mainpayrooll.frx":0F4E
      TabIndex        =   1
      Top             =   2640
      Width           =   1575
   End
   Begin VB.ListBox lstpaid 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      ItemData        =   "mainpayrooll.frx":0F50
      Left            =   120
      List            =   "mainpayrooll.frx":0F52
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PAYMENT FOR THE CURRENT MONTH.........."
      BeginProperty Font 
         Name            =   "Zurich Cn BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   855
      Left            =   2160
      TabIndex        =   29
      Top             =   4680
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extras:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   1800
      TabIndex        =   26
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Net:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   1800
      TabIndex        =   20
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Other tax:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   1800
      TabIndex        =   19
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   18
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque no:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   17
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pay mode:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   16
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Pay:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1800
      TabIndex        =   15
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Salary:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   14
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Emp no:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "UNPAID EMP:"
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
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PAID EMP:"
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub cmdleaveaddition_Click()
Dim total As Long
total = Val(txtsalary.Text) * 12
If total > 50000 And total <= 60000 Then
total = total - 50000
total = (total / 100) * 10
total = total / 12
txttax.Text = total
ElseIf total > 60000 And total <= 150000 Then
total = total - 60000
total = (total / 100) * 20
total = total / 12
txttax.Text = total
ElseIf total > 150000 Then
total = total - 150000
total = (total / 100) * 30
total = total / 12
txttax.Text = total
Else
txttax.Text = "0"
End If

End Sub

Private Sub cmdpay_Click()
On Error Resume Next

If txtempno = "" Or txtsalary = "" Or txtnet = "" Or txtdop = "" Or txtpaymode = "" Then
MsgBox "one or more fields r left blank..so try again"
Exit Sub
End If
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = Val(txtempno.Text)
Adodc2.Recordset.Fields(1) = Val(txtsalary.Text)
Adodc2.Recordset.Fields(2) = txtdop.Text
Adodc2.Recordset.Fields(3) = txtpaymode.Text
Adodc2.Recordset.Fields(4) = txtchequeno.Text
Adodc2.Recordset.Fields(5) = Val(txttax.Text)
Adodc2.Recordset.Fields(6) = Val(txtothertax.Text)
Adodc2.Recordset.Fields(7) = Val(txtnet.Text)
Adodc2.Recordset.Fields(8) = Val(txtleave.Text)
Adodc2.Recordset.Update
MsgBox "records r updated"
lstpaid.Clear
lstunpaid.Clear
Form_Load



End Sub

Private Sub cmdrebateothertax_Click()
txtothertax.Text = "0"

End Sub

Private Sub cmdrebatetax_Click()
txttax.Text = "0"

End Sub

Private Sub cmdtotal_Click()
txtnet.Text = Val(txtsalary.Text) + Val(txttax.Text) + Val(txtothertax.Text) + Val(txtleave.Text)

End Sub

Private Sub Command1_Click()
txtleave.Text = "0"

End Sub

Private Sub Command2_Click()
Form11.Visible = True

End Sub

Private Sub Form_Load()
Form5.Width = 6105
Form5.Height = 6090
On Error GoTo x
Dim i As Integer
Dim j As Integer
Dim k As Long

On Error Resume Next
Adodc1.Refresh
Adodc1.Recordset.MoveLast
Adodc1.Recordset.MoveFirst
While Adodc1.Recordset.EOF <> True
k = Adodc1.Recordset.Fields(0)
Adodc2.Recordset.MoveFirst
While Adodc2.Recordset.EOF <> True
On Error Resume Next
If k = Adodc2.Recordset.Fields(0) And Left(Adodc2.Recordset.Fields(2), 2) = Left(Date, 2) Then
lstpaid.AddItem Adodc1.Recordset.Fields(2) & "  " & Adodc1.Recordset.Fields(0)

GoTo m
End If
Adodc2.Recordset.MoveNext
Wend

lstunpaid.AddItem Adodc1.Recordset.Fields(2) & "  " & Adodc1.Recordset.Fields(0)
m:
Adodc1.Recordset.MoveNext
k = ""

Wend
Adodc1.Recordset.MoveFirst
Adodc2.Recordset.MoveFirst

x:
End Sub

Private Sub lstpaid_Click()
Adodc2.Refresh
'Adodc2.Recordset.Find "empno=" & Val(Mid(lstpaid.Text, InStr(lstpaid.Text, " ")))
Adodc2.Recordset.MoveFirst
While Adodc2.Recordset.EOF <> True
 If Adodc2.Recordset.Fields(0) = Val(Mid(lstpaid.Text, InStr(lstpaid.Text, " "))) Then
 txtempno = Adodc2.Recordset.Fields(0)
 txtsalary = Adodc2.Recordset.Fields(1)
 txtdop = Adodc2.Recordset.Fields(2)
 txtpaymode = Adodc2.Recordset.Fields(3)
 txtchequeno = Adodc2.Recordset.Fields(4)
 txttax = Adodc2.Recordset.Fields(5)
 txtothertax = Adodc2.Recordset.Fields(6)
 txtleave = Adodc2.Recordset.Fields(8)
 txtnet = Adodc2.Recordset.Fields(7)
 
 Exit Sub
 End If
Adodc2.Recordset.MoveNext
Wend
End Sub

Private Sub lstunpaid_Click()
Dim m As Long
Dim n As Long
txtempno.Text = ""
txtchequeno = ""
txtdop = ""
txtpaymode = ""
txtsalary = ""
txtnet = ""
txtothertax = ""
txttax = ""
txtleave = ""
m = Val(Mid(lstunpaid.Text, InStr(lstunpaid.Text, " ")))
txtempno.Text = m
Adodc1.Recordset.MoveFirst
While Adodc1.Recordset.EOF <> True
If m = Adodc1.Recordset.Fields(0) Then
n = Adodc1.Recordset.Fields(18)
txtsalary = n
txtdop = Date

Exit Sub
End If
Adodc1.Recordset.MoveNext

Wend
End Sub
