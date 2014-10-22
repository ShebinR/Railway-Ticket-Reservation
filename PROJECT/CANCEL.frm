VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form CANCEL 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000003&
   Caption         =   "Form1"
   ClientHeight    =   9435
   ClientLeft      =   2790
   ClientTop       =   1680
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   Picture         =   "CANCEL.frx":0000
   ScaleHeight     =   9435
   ScaleWidth      =   15420
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   720
      Top             =   720
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
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
      Caption         =   "CANCEL TICKET"
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
      Height          =   735
      Left            =   8040
      TabIndex        =   10
      Top             =   8040
      Width           =   2775
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
      Left            =   4080
      TabIndex        =   9
      Top             =   8040
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "TICKET DETAILS"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   4680
      TabIndex        =   0
      Top             =   2520
      Width           =   6135
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
         Height          =   405
         Left            =   3000
         TabIndex        =   8
         Top             =   3600
         Width           =   2295
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
         Height          =   375
         Left            =   3000
         TabIndex        =   7
         Top             =   2760
         Width           =   2295
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
         Height          =   405
         Left            =   3000
         TabIndex        =   6
         Top             =   1800
         Width           =   2295
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
         Height          =   405
         Left            =   3000
         TabIndex        =   5
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label4 
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
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label3 
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
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label2 
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
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label1 
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
         TabIndex        =   1
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   9135
      Left            =   -240
      Picture         =   "CANCEL.frx":210CF
      ScaleHeight     =   9075
      ScaleWidth      =   15915
      TabIndex        =   11
      Top             =   0
      Width           =   15975
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         BackStyle       =   0  'Transparent
         Caption         =   "TICKET CANCELLATION"
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   5760
         TabIndex        =   12
         Top             =   960
         Width           =   4215
      End
   End
End
Attribute VB_Name = "CANCEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conv As ADODB.Connection
Dim rsv As ADODB.Recordset
Dim cmdv As ADODB.Command

Private Sub Command1_Click()
Dim pnr As Integer
pnr = Val(Text1.Text)
Dim pas As Integer
pas = Val(Text2.Text)
Dim tno As Integer
tno = Val(Text3.Text)
Dim da As String
da = Text4.Text

rsv.Open "delete from payment where pnr_no='" & pnr & "' ", conv, adOpenDynamic, adLockOptimistic
If (rsv.State = 1) Then
rsv.Close
End If

rsv.Open "delete from ticket where pnr_no='" & pnr & "' and pass_id='" & pas & "' ", conv, adOpenDynamic, adLockOptimistic
If (rsv.State = 1) Then
rsv.Close
End If

Dim S As String
Dim SE As Integer
rsv.Open "SELECT * FROM AVAILABILITY WHERE TRAIN_NO = '" & tno & "' AND T_DATE = '" & da & "' ", conv, adOpenDynamic, adLockOptimistic

S = rsv.Fields("STATUS")
SE = rsv.Fields("SEATS")
rsv.Close

If (S = "FULL") Then
    rsv.Open "UPDATE AVAILABILITY SET SEATS=SEATS+1 WHERE TRAIN_NO='" & tno & "' AND T_DATE='" & da & "' ", conv, adOpenDynamic, adLockOptimistic
    If (rsv.State = 1) Then
        rsv.Close
    End If
    rsv.Open "UPDATE AVAILABILITY SET STATUS='AV' WHERE TRAIN_NO='" & tno & "' AND T_DATE='" & da & "' ", conv, adOpenDynamic, adLockOptimistic
    If (rsv.State = 1) Then
        rsv.Close
    End If
Else
    rsv.Open "UPDATE AVAILABILITY SET SEATS=SEATS+1 WHERE TRAIN_NO='" & tno & "' AND T_DATE='" & da & "' ", conv, adOpenDynamic, adLockOptimistic
    If (rsv.State = 1) Then
        rsv.Close
    End If
End If

MsgBox ("YOUR TICKET HAS BEEN SUCCESSFULLY CANCELLED.....")

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

