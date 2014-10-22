VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form BOOK_TRAIN 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form1"
   ClientHeight    =   9435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   15420
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Left            =   960
      Picture         =   "BOOK_TRAIN.frx":0000
      ScaleHeight     =   4395
      ScaleWidth      =   4995
      TabIndex        =   7
      Top             =   2400
      Width           =   5055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   600
      Top             =   600
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "TRAIN DETAILS"
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
      Caption         =   "EXIT TO MAIN MENU"
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
      Left            =   10560
      TabIndex        =   6
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROCEED"
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
      Left            =   2880
      TabIndex        =   5
      Top             =   7920
      Width           =   2415
   End
   Begin VB.Frame Frame3 
      Caption         =   "TRAIN DETAILS"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   7440
      TabIndex        =   0
      Top             =   2040
      Width           =   5655
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "BOOK_TRAIN.frx":E85E
         Left            =   2520
         List            =   "BOOK_TRAIN.frx":E871
         TabIndex        =   13
         Top             =   3960
         Width           =   2415
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         IntegralHeight  =   0   'False
         ItemData        =   "BOOK_TRAIN.frx":E8B1
         Left            =   2520
         List            =   "BOOK_TRAIN.frx":E8C7
         Style           =   1  'Simple Combo
         TabIndex        =   12
         Top             =   1680
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "BOOK_TRAIN.frx":E931
         Left            =   2520
         List            =   "BOOK_TRAIN.frx":E947
         TabIndex        =   11
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label9 
         Caption         =   "TRIAN NAME"
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
         TabIndex        =   4
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "TRAIN NO"
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
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "DATE"
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
         TabIndex        =   2
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label17 
         Caption         =   "FORMAT : DD-MM-YYYY"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   1
         Top             =   4680
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   10935
      Left            =   0
      Picture         =   "BOOK_TRAIN.frx":E975
      ScaleHeight     =   10875
      ScaleWidth      =   16035
      TabIndex        =   8
      Top             =   0
      Width           =   16095
      Begin VB.CommandButton Command3 
         Caption         =   "RESET ALL VALUES"
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
         Left            =   6720
         TabIndex        =   10
         Top             =   7920
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TICKET BOOKING"
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
         Left            =   4800
         TabIndex        =   9
         Top             =   1200
         Width           =   3615
      End
   End
End
Attribute VB_Name = "BOOK_TRAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conv As ADODB.Connection
Dim rsv As ADODB.Recordset
Dim cmdv As ADODB.Command

Public pnr As Integer
Public tno, am, pas As Integer
Public av, tname, da As String
Private Sub Command1_Click()

Dim ne As Integer
Dim count, pn As Integer
pn = 2435
rsv.Open "select count(*) as c from ticket ", conv, adOpenDynamic, adLockOptimistic
count = rsv.Fields("c")
rsv.Close

rsv.Open "SELECT *FROM TICKET order by PNR_NO desc ", conv, adOpenDynamic, adLockOptimistic
If (count = 0) Then
pnr = pn + 1 + BOOKING.A
Else
pnr = rsv.Fields("PNR_NO") + 1
End If


rsv.Close
tno = Combo1.Text
tname = Combo2.Text
da = Combo3.Text

rsv.Open "select amount from train where train_no = '" & tno & "' ", conv, adOpenDynamic, adLockOptimistic
am = rsv.Fields("amount")
If (rsv.State = 1) Then
    rsv.Close
End If
 
pas = BOOKING.p

rsv.Open "select * from availability where train_no ='" & tno & "' and t_date ='" & da & "' ", conv, adOpenDynamic, adLockOptimistic
seat = rsv.Fields("seats")
av = rsv.Fields("status")
If (rsv.State = 1) Then
    rsv.Close
End If

Dim N, H As Integer
If (av = "AV") Then
    rsv.Open "SELECT COUNT(*) AS Y FROM TICKET WHERE TRAIN_NO='" & tno & "' AND T_DATE='" & da & "' ", conv, adOpenDynamic, adLockOptimistic
    N = rsv.Fields("Y")
    ST = "CONF"
    seat = N + 1
    H = 1
    If (rsv.State = 1) Then
        rsv.Close
    End If
Else
    MsgBox ("TRAIN IS FULL...SORRY...CANNOT BOOK THE TICKET...!!!")
End If

If (H = 1) Then
    rsv.Open "insert into ticket values('" & pnr & "','" & tno & "','" & pas & "','" & da & "' ,'" & am & "','" & ST & "','" & seat & "')", conv, adOpenDynamic, adLockOptimistic
    If (rsv.State = 1) Then
        rsv.Close
    End If
    rsv.Open "SELECT *FROM AVAILABILITY WHERE TRAIN_NO='" & tno & "' AND T_DATE='" & da & "' ", conv, adOpenDynamic, adLockOptimistic
    Dim X As Integer
    X = rsv.Fields("SEATS")
    rsv.Close
    If (X = 1) Then
        rsv.Open "UPDATE AVAILABILITY SET SEATS = 0 WHERE TRAIN_NO='" & tno & "' AND T_DATE='" & da & "' ", conv, adOpenDynamic, adLockOptimistic
        If (rsv.State = 1) Then
            rsv.Close
        End If
        rsv.Open "UPDATE AVAILABILITY SET STATUS ='FULL' WHERE TRAIN_NO='" & tno & "' AND T_DATE='" & da & "' ", conv, adOpenDynamic, adLockOptimistic
        If (rsv.State = 1) Then
            rsv.Close
        End If
    Else
        rsv.Open "UPDATE AVAILABILITY SET SEATS = SEATS-1 WHERE TRAIN_NO='" & tno & "' AND T_DATE='" & da & "' ", conv, adOpenDynamic, adLockOptimistic
        If (rsv.State = 1) Then
            rsv.Close
        End If
    End If
    PAYMENT.Text1.Text = ""
    PAYMENT.Text2.Text = ""
    PAYMENT.Text3.Text = ""
    PAYMENT.Text4.Text = ""
    PAYMENT.Show
    Me.Hide
Else
    MENU.Show
    Me.Hide
End If

End Sub

Private Sub Command2_Click()
Dim pass As Integer
pass = BOOKING.p
rsv.Open "delete from passenger where pass_id = '" & pass & "' ", conv, adOpenDynamic, adLockOptimistic
If (rsv.State = 1) Then
rsv.Close
End If
MENU.Show
Me.Hide
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub Form_Load()
Set conv = New ADODB.Connection
Set rsv = New ADODB.Recordset
Set cmdv = New ADODB.Command
conv.Open "Provider=MSDASQL.1;Password=shebin;Persist Security Info=True;User ID=SYSTEM;Data Source=my_pro"
End Sub

