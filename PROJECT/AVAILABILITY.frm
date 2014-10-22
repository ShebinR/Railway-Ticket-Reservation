VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form AVAILABILITY 
   BackColor       =   &H00FFFFC0&
   ClientHeight    =   9435
   ClientLeft      =   2985
   ClientTop       =   1065
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   15420
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   600
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "AVAILABILITY"
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
   Begin VB.CommandButton Command3 
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
      Left            =   8520
      TabIndex        =   13
      Top             =   8160
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CHECK ANOTHER TRAIN"
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
      TabIndex        =   12
      Top             =   8160
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      Caption         =   "TRAIN DETAILS"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   8880
      TabIndex        =   1
      Top             =   2520
      Width           =   5295
      Begin VB.TextBox Text7 
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
         Left            =   2760
         TabIndex        =   18
         Top             =   3360
         Width           =   2055
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
         Left            =   2760
         TabIndex        =   17
         Top             =   2400
         Width           =   2055
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
         Height          =   375
         Left            =   2760
         TabIndex        =   8
         Top             =   1560
         Width           =   2055
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
         Left            =   2760
         TabIndex        =   7
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "NO OF SEATS"
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
         TabIndex        =   16
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label6 
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
         Height          =   375
         Left            =   600
         TabIndex        =   15
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "TRAIN NAME"
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
         TabIndex        =   6
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label4 
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
         TabIndex        =   5
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "TRAVEL DETAILS"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   1320
      TabIndex        =   0
      Top             =   2520
      Width           =   5535
      Begin VB.TextBox Text6 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
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
         TabIndex        =   11
         Top             =   3000
         Width           =   1815
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
         Height          =   375
         Left            =   3000
         TabIndex        =   10
         Top             =   1920
         Width           =   1815
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
         Height          =   405
         Left            =   3000
         TabIndex        =   9
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "FORMAT: (DD-MM-YYYY)"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd.MM.yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
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
         Left            =   3000
         TabIndex        =   14
         Top             =   3720
         Width           =   1815
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
         Left            =   360
         TabIndex        =   4
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "DESTINATION   PLACE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "STARTING PLACE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   10935
      Left            =   0
      Picture         =   "AVAILABILITY.frx":0000
      ScaleHeight     =   10875
      ScaleWidth      =   15915
      TabIndex        =   19
      Top             =   0
      Width           =   15975
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CHECK AVAILATBILITY"
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
         Height          =   375
         Left            =   6120
         TabIndex        =   20
         Top             =   840
         Width           =   3615
      End
   End
End
Attribute VB_Name = "AVAILABILITY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conv As ADODB.Connection
Dim rsv As ADODB.Recordset
Dim cmdv As ADODB.Command

Private Sub Command1_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Dim view1, view2, view3, view4 As String
view1 = InputBox("ENTER THE STARTING PLACE")
view2 = InputBox("ENTER THE DESTIONATION")
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
        Text1.Text = rsv.Fields("TRAIN_NO")
        Text2.Text = rsv.Fields("TRAIN_NAME")
        Text4.Text = rsv.Fields("STARTING_POINT")
        Text5.Text = rsv.Fields("DESTINATION")
        Text6.Text = view3
        tnam = rsv.Fields("TRAIN_NAME")
        rsv.Close
        rsv.Open "select * from availability where t_date = '" & view3 & "' and train_no= '" & view4 & "' ", conv, adOpenDynamic, adLockOptimistic
        Text3.Text = rsv.Fields("STATUS")
        Text7.Text = rsv.Fields("SEATS")
        If (rsv.State = 1) Then
            rsv.Close
        End If
    End If
End If
End Sub

Private Sub Command3_Click()
MENU.Show
Me.Hide
End Sub
Private Sub Form_Load()
Set conv = New ADODB.Connection
Set rsv = New ADODB.Recordset
Set cmdv = New ADODB.Command
conv.Open "Provider=MSDASQL.1;Password=shebin;Persist Security Info=True;User ID=SYSTEM;Data Source=my_pro"
End Sub

