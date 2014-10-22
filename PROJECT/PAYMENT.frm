VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form PAYMENT 
   BackColor       =   &H00FFFFFF&
   Caption         =   "PAYMENT"
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15420
   LinkTopic       =   "Form2"
   ScaleHeight     =   8490
   ScaleWidth      =   15420
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PAYMENT DETAILS"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   8400
      TabIndex        =   2
      Top             =   1920
      Width           =   6375
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
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   3360
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   3480
         Width           =   1815
      End
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
         Left            =   3360
         TabIndex        =   12
         Top             =   2520
         Width           =   1815
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
         Height          =   495
         Left            =   3360
         TabIndex        =   11
         Top             =   1560
         Width           =   1815
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
         Height          =   495
         Left            =   3360
         TabIndex        =   10
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "BANK NAME"
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
         Left            =   720
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CARD TYPE"
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
         Left            =   720
         TabIndex        =   5
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CARD NO."
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
         Left            =   720
         TabIndex        =   4
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "PIN NUMBER"
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
         Left            =   720
         TabIndex        =   3
         Top             =   3600
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   960
      Top             =   840
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "PAYMENT"
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
      Caption         =   "CANCEL"
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
      Left            =   10440
      TabIndex        =   1
      Top             =   7320
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PAY TICKET"
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
      Left            =   2640
      TabIndex        =   0
      Top             =   7320
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   10005
      Left            =   0
      Picture         =   "PAYMENT.frx":0000
      ScaleHeight     =   9945
      ScaleWidth      =   15600
      TabIndex        =   7
      Top             =   0
      Width           =   15660
      Begin VB.CommandButton Command3 
         Caption         =   "RESET VALUES"
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
         Left            =   6600
         TabIndex        =   9
         Top             =   7320
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "PAYMENT DETAILS"
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
         Left            =   9600
         TabIndex        =   8
         Top             =   1200
         Width           =   3615
      End
   End
End
Attribute VB_Name = "PAYMENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conv As ADODB.Connection
Dim rsv As ADODB.Recordset
Dim cmdv As ADODB.Command

Private Sub Command1_Click()
Dim PID, MI As Integer
PID = BOOK_TRAIN.pas
MI = BOOK_TRAIN.pnr
Dim bname, ctype As String
Dim cno As Single
bname = Text1.Text
ctype = Text2.Text

rsv.Open "insert into payment values ('" & MI & "', '" & PID & "', '" & bname & "','" & ctype & "','" & Text3.Text & "')", conv, adOpenDynamic, adLockOptimistic
If (rsv.State = 1) Then
rsv.Close
End If
Form1.Text1.Text = ""
Form1.Text2.Text = ""
Form1.Text3.Text = ""
Form1.Text4.Text = ""
Form1.Text5.Text = ""
Form1.Text6.Text = ""
Form1.Text7.Text = ""
Form1.Text8.Text = ""
Form1.Text9.Text = ""
Form1.Text10.Text = ""
Form1.Text11.Text = ""
Form1.Text12.Text = ""

Dim X As Integer
X = BOOK_TRAIN.pnr
rsv.Open "select *from ticket where pnr_no='" & X & "' ", conv, adOpenDynamic, adLockOptimistic
Dim T As Integer
T = rsv.Fields("TRAIN_NO")
Form1.Text1.Text = rsv.Fields("PNR_NO")
Form1.Text2.Text = rsv.Fields("TRAIN_NO")
Form1.Text3.Text = BOOK_TRAIN.da
Form1.Text4.Text = rsv.Fields("PASS_ID")
Form1.Text5.Text = rsv.Fields("amount")
Form1.Text6.Text = rsv.Fields("status")
Form1.Text7.Text = rsv.Fields("SEATNO")
If (rsv.State = 1) Then
rsv.Close
End If

rsv.Open "SELECT *FROM TRAIN WHERE TRAIN_NO=' " & T & "' ", conv, adOpenDynamic, adLockOptimistic
Form1.Text8.Text = rsv.Fields("TRAIN_NAME")
Form1.Text9.Text = rsv.Fields("STARTING_POINT")
Form1.Text10.Text = rsv.Fields("DESTINATION")
Form1.Text11.Text = rsv.Fields("DEPARTURE_TIME")
Form1.Text12.Text = rsv.Fields("DESTINATION_TIME")
Form1.Show
Me.Hide
MsgBox ("Your Ticket Has Been Successfully Booked.......")

End Sub

Private Sub Command2_Click()
Dim S As String
Dim SE As Integer

rsv.Open "DELETE FROM TICKET WHERE PASS_ID='" & BOOKING.p & "' ", conv, adOpenDynamic, adLockOptimistic
If (rsv.State = 1) Then
rsv.Close
End If

rsv.Open "DELETE FROM PASSENGER WHERE PASS_ID='" & BOOKING.p & "' ", conv, adOpenDynamic, adLockOptimistic
If (rsv.State = 1) Then
rsv.Close
End If



Dim T As Integer
T = BOOK_TRAIN.tno


rsv.Open "SELECT * FROM AVAILABILITY WHERE TRAIN_NO = '" & T & "' AND T_DATE = '" & BOOK_TRAIN.da & "' ", conv, adOpenDynamic, adLockOptimistic
S = rsv.Fields("STATUS")
SE = rsv.Fields("SEATS")
SE = SE + 1
rsv.Close

If (S = "AV") Then
rsv.Open "UPDATE AVAILABILITY SET SEATS=SEATS+1 WHERE TRAIN_NO='" & BOOK_TRAIN.tno & "' AND T_DATE='" & BOOK_TRAIN.da & "' ", conv, adOpenDynamic, adLockOptimistic
If (rsv.State = 1) Then
rsv.Close
End If

Else
rsv.Open "UPDATE AVAILABILITY SET SEATS=SEATS-1 WHERE TRAIN_NO='" & BOOK_TRAIN.tno & "' AND T_DATE='" & BOOK_TRAIN.da & "' ", conv, adOpenDynamic, adLockOptimistic
If (rsv.State = 1) Then
rsv.Close
End If
End If


MENU.Show
Me.Hide
End Sub

Private Sub Label4_Click()

End Sub

Private Sub Command3_Click()
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
End Sub

Private Sub Form_Load()
Set conv = New ADODB.Connection
Set rsv = New ADODB.Recordset
Set cmdv = New ADODB.Command
conv.Open "Provider=MSDASQL.1;Password=shebin;Persist Security Info=True;User ID=SYSTEM;Data Source=my_pro"
End Sub

