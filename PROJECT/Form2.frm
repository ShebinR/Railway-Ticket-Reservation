VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   9435
   ClientLeft      =   2790
   ClientTop       =   1485
   ClientWidth     =   15420
   LinkTopic       =   "Form2"
   ScaleHeight     =   9435
   ScaleWidth      =   15420
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1320
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=MSDASQL.1;Password=shebin;Persist Security Info=True;User ID=SYSTEM;Data Source=my_pro"
      OLEDBString     =   "Provider=MSDASQL.1;Password=shebin;Persist Security Info=True;User ID=SYSTEM;Data Source=my_pro"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "PNR"
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
   Begin VB.CommandButton Command2 
      Caption         =   "BACK TO MAIN MENU"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      TabIndex        =   12
      Top             =   8280
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CHECK ANOTHER STATUS"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      Picture         =   "Form2.frx":0000
      TabIndex        =   11
      Top             =   8280
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Caption         =   "DETAILS"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   5160
      TabIndex        =   0
      Top             =   1680
      Width           =   5175
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   16
         Top             =   4800
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2760
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2760
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2760
         TabIndex        =   3
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2760
         TabIndex        =   2
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2760
         TabIndex        =   1
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000016&
         Caption         =   "SEAT NO"
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   15
         Top             =   4920
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000016&
         Caption         =   "PASS ID"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000016&
         Caption         =   "TRIAN CODE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   9
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000016&
         Caption         =   "PNR NO."
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000016&
         Caption         =   "AMOUNT"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000016&
         Caption         =   "STATUS"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   4080
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9495
      Left            =   0
      Picture         =   "Form2.frx":C483
      ScaleHeight     =   9435
      ScaleWidth      =   15435
      TabIndex        =   13
      Top             =   0
      Width           =   15495
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PNR STATUS"
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   14
         Top             =   960
         Width           =   3135
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conv As ADODB.Connection
Dim rsv As ADODB.Recordset
Dim cmdv As ADODB.Command
Private Sub Command1_Click()
Form2.Text1.Text = ""
Form2.Text2.Text = ""
Form2.Text4.Text = ""
Form2.Text5.Text = ""
Form2.Text6.Text = ""
Form2.Text3.Text = ""
Dim VIEW As Integer
VIEW = InputBox("ENTER THE PNR NUMBER")
rsv.Open "select COUNT(*) AS C from ticket where pnr_no='" & VIEW & "' ", conv, adOpenDynamic, adLockOptimistic
C = rsv.Fields("C")
rsv.Close
rsv.Open "select *from ticket where pnr_no='" & VIEW & "' ", conv, adOpenDynamic, adLockOptimistic
If (C = 0) Then
MsgBox ("INVALID PNR NUMBER...!!!")
ElseIf (C = 1) Then
Text1.Text = rsv.Fields("PNR_NO")
Text2.Text = rsv.Fields("TRAIN_NO")
Text4.Text = rsv.Fields("PASS_ID")
Text5.Text = rsv.Fields("amount")
Text6.Text = rsv.Fields("status")
Text3.Text = rsv.Fields("seatno")
End If
If (rsv.State = 1) Then
rsv.Close
End If

End Sub

Private Sub Command2_Click()
MENU.Show
Me.Hide
End Sub

Private Sub Form_Load()
Set conv = New ADODB.Connection
Set rsv = New ADODB.Recordset
Set cmdv = New ADODB.Command
conv.Open "Provider=MSDASQL.1;Password=shebin;Persist Security Info=True;User ID=SYSTEM;Data Source=my_pro"
End Sub

