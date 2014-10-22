VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form MENU 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   9435
   ClientLeft      =   7695
   ClientTop       =   4155
   ClientWidth     =   15420
   FillColor       =   &H00C0FFFF&
   LinkTopic       =   "Form1"
   Picture         =   "MENU.frx":0000
   ScaleHeight     =   9435
   ScaleWidth      =   15420
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      TabIndex        =   5
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      Caption         =   "TRAIN DETAILS"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   4
      Top             =   7920
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "PNR STATUS"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      TabIndex        =   3
      Top             =   6120
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CHECK AVAILABILITY"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      TabIndex        =   2
      Top             =   4320
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL TICKET"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   1
      Top             =   6120
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000E&
      Caption         =   "BOOK TICKET"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   4320
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   11175
      Left            =   0
      Picture         =   "MENU.frx":1D216
      ScaleHeight     =   11115
      ScaleMode       =   0  'User
      ScaleWidth      =   15915
      TabIndex        =   6
      Top             =   0
      Width           =   15975
      Begin MSAdodcLib.Adodc MENU 
         Height          =   615
         Left            =   600
         Top             =   6360
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
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
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TO"
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
         Left            =   6840
         TabIndex        =   10
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WELCOME "
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
         Left            =   5160
         TabIndex        =   9
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "CHOOSE AMOUNG THE FOLLOWING OPTIONS......."
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   8
         Top             =   2760
         Width           =   6015
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "RAILWAY TICKET RESERVATION SYSTEM"
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
         Left            =   3600
         TabIndex        =   7
         Top             =   1920
         Width           =   7815
      End
   End
End
Attribute VB_Name = "MENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conv As ADODB.Connection
Dim rsv As ADODB.Recordset
Dim cmdv As ADODB.Command
Private Sub Command1_Click()
BOOKING.Text1.Text = ""
BOOKING.Text2.Text = ""
BOOKING.Text3.Text = ""
BOOKING.Text4.Text = ""
BOOKING.Text5.Text = ""
BOOKING.Text6.Text = ""
BOOKING.Text7.Text = ""
BOOKING.Option1.Value = False
BOOKING.Option2.Value = False
BOOKING.Show
Me.Hide
End Sub

Private Sub Command2_Click()
CANCEL.Text1.Text = ""
CANCEL.Text2.Text = ""
CANCEL.Text3.Text = ""
CANCEL.Text4.Text = ""
CANCEL.Show
Me.Hide
End Sub

Private Sub Command3_Click()
AVAILABILITY.Show
AVAILABILITY.Text1.Text = ""
AVAILABILITY.Text2.Text = ""
AVAILABILITY.Text3.Text = ""
AVAILABILITY.Text4.Text = ""
AVAILABILITY.Text5.Text = ""
AVAILABILITY.Text6.Text = ""
AVAILABILITY.Text7.Text = ""
Dim view1, view2, view3, view4 As String
view1 = InputBox("ENTER THE STARTING PLACE")
view2 = InputBox("ENTER THE DESTINATION PLACE")
view3 = InputBox("ENTER THE DATE OF JOURNEY")
Dim C, D As Integer
rsv.Open "select COUNT(*) AS C from train where starting_point ='" & view1 & "' and destination = '" & view2 & "' ", conv, adOpenDynamic, adLockOptimistic
C = rsv.Fields("C")
rsv.Close
If (C = 0) Then
    MsgBox ("NO RECORD FOUND..!!!")
Else
    rsv.Open "select * from train where starting_point ='" & view1 & "' and destination = '" & view2 & "' ", conv, adOpenDynamic, adLockOptimistic
    view4 = rsv.Fields("TRAIN_NO")
    rsv.Close
    rsv.Open "select COUNT(*) AS D from availability where t_date = '" & view3 & "' and train_no= '" & view4 & "' ", conv, adOpenDynamic, adLockOptimistic
    D = rsv.Fields("D")
    rsv.Close
    If (D = 0) Then
        MsgBox ("NO RECORD FOUND..!!!")
    Else
        rsv.Open "select * from train where starting_point ='" & view1 & "' and destination = '" & view2 & "' ", conv, adOpenDynamic, adLockOptimistic
        AVAILABILITY.Text1.Text = rsv.Fields("TRAIN_NO")
        AVAILABILITY.Text2.Text = rsv.Fields("TRAIN_NAME")
        AVAILABILITY.Text4.Text = rsv.Fields("STARTING_POINT")
        AVAILABILITY.Text5.Text = rsv.Fields("DESTINATION")
        AVAILABILITY.Text6.Text = view3
        tnam = rsv.Fields("TRAIN_NAME")
        rsv.Close
        rsv.Open "select * from availability where t_date = '" & view3 & "' and train_no= '" & view4 & "' ", conv, adOpenDynamic, adLockOptimistic
        AVAILABILITY.Text3.Text = rsv.Fields("STATUS")
        AVAILABILITY.Text7.Text = rsv.Fields("SEATS")
        If (rsv.State = 1) Then
            rsv.Close
        End If
    End If
End If
Me.Hide
End Sub

Private Sub Command4_Click()
Form2.Text1.Text = ""
Form2.Text2.Text = ""
Form2.Text4.Text = ""
Form2.Text5.Text = ""
Form2.Text6.Text = ""
Form2.Text3.Text = ""
Form2.Show
Dim VIEW, C As Integer
VIEW = InputBox("ENTER THE PNR NUMBER")
rsv.Open "select COUNT(*) AS C from ticket where pnr_no='" & VIEW & "' ", conv, adOpenDynamic, adLockOptimistic
C = rsv.Fields("C")
rsv.Close
rsv.Open "select *from ticket where pnr_no='" & VIEW & "' ", conv, adOpenDynamic, adLockOptimistic
If (C = 0) Then
MsgBox ("INVALID PNR NUMBER...!!!")
Else
Form2.Text1.Text = rsv.Fields("PNR_NO")
Form2.Text2.Text = rsv.Fields("TRAIN_NO")
Form2.Text4.Text = rsv.Fields("PASS_ID")
Form2.Text5.Text = rsv.Fields("amount")
Form2.Text6.Text = rsv.Fields("status")
Form2.Text3.Text = rsv.Fields("seatno")
End If
If (rsv.State = 1) Then
rsv.Close
End If
Me.Hide
End Sub

Private Sub Command5_Click()
DETAILS.Text1.Text = ""
DETAILS.Text2.Text = ""
DETAILS.Text3.Text = ""
DETAILS.Text4.Text = ""
DETAILS.Text5.Text = ""
DETAILS.Text6.Text = ""
DETAILS.Text7.Text = ""
DETAILS.Text8.Text = ""
DETAILS.Show
Dim VIEW As String
VIEW = InputBox("ENTER THE TRAIN NAME")
Dim C As Integer
rsv.Open "SELECT COUNT(*) AS C FROM TRAIN WHERE TRAIN_NAME='" & VIEW & "' ", conv, adOpenDynamic, adLockOptimistic
C = rsv.Fields("C")
rsv.Close
If (C = 0) Then
MsgBox ("NO RECORD FOUND FOR THE ENTERED TRIAN NAME...!!!")
Else
rsv.Open "select * from train where train_name='" & VIEW & "' ", conv, adOpenDynamic, adLockOptimistic
DETAILS.Text1.Text = rsv.Fields("TRAIN_NO")
DETAILS.Text2.Text = rsv.Fields("TRAIN_NAME")
DETAILS.Text3.Text = rsv.Fields("STARTING_POINT")
DETAILS.Text4.Text = rsv.Fields("DESTINATION")
DETAILS.Text5.Text = rsv.Fields("DEPARTURE_TIME")
DETAILS.Text6.Text = rsv.Fields("DESTINATION_TIME")
DETAILS.Text7.Text = rsv.Fields("DURATION")
DETAILS.Text8.Text = rsv.Fields("AMOUNT")
If (rsv.State = 1) Then
rsv.Close
End If
End If
Me.Hide
End Sub

Private Sub Command6_Click()
Form3.Show
Me.Hide
End Sub

Private Sub Form_Load()
Set conv = New ADODB.Connection
Set rsv = New ADODB.Recordset
Set cmdv = New ADODB.Command
conv.Open "Provider=MSDASQL.1;Password=shebin;Persist Security Info=True;User ID=SYSTEM;Data Source=my_pro"

End Sub

