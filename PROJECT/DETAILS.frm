VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form DETAILS 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   9435
   ClientLeft      =   2790
   ClientTop       =   1275
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   Picture         =   "DETAILS.frx":0000
   ScaleHeight     =   9435
   ScaleWidth      =   15420
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
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
      Height          =   615
      Left            =   8280
      Picture         =   "DETAILS.frx":9098
      TabIndex        =   10
      Top             =   8400
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   1080
      TabIndex        =   1
      Top             =   2640
      Width           =   5535
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
         Left            =   2760
         TabIndex        =   9
         Top             =   3720
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
         Height          =   495
         Left            =   2760
         TabIndex        =   7
         Top             =   2760
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
         Height          =   495
         Left            =   2760
         TabIndex        =   5
         Top             =   1680
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
         Height          =   495
         Left            =   2760
         TabIndex        =   3
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "DESTINATION"
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
         Left            =   480
         TabIndex        =   8
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "STARTING POINT"
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
         Left            =   480
         TabIndex        =   6
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "TRAIN NAME "
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
         Left            =   480
         TabIndex        =   4
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
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
         Left            =   480
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CHECK ANOTHER TRAIN"
      Default         =   -1  'True
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
      TabIndex        =   0
      Top             =   8400
      Width           =   3255
   End
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
   Begin VB.PictureBox Picture1 
      Height          =   9135
      Left            =   0
      Picture         =   "DETAILS.frx":12130
      ScaleHeight     =   9075
      ScaleWidth      =   15915
      TabIndex        =   11
      Top             =   0
      Width           =   15975
      Begin VB.Frame Frame2 
         BackColor       =   &H80000003&
         Height          =   4935
         Left            =   8400
         TabIndex        =   13
         Top             =   2640
         Width           =   5895
         Begin VB.TextBox Text8 
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
            Left            =   3240
            TabIndex        =   21
            Top             =   3840
            Width           =   1935
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
            Height          =   495
            Left            =   3240
            TabIndex        =   16
            Top             =   600
            Width           =   1935
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
            Height          =   495
            Left            =   3240
            TabIndex        =   15
            Top             =   1560
            Width           =   1935
         End
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
            Height          =   495
            Left            =   3240
            TabIndex        =   14
            Top             =   2760
            Width           =   1935
         End
         Begin VB.Label Label9 
            BackColor       =   &H80000003&
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
            TabIndex        =   20
            Top             =   3840
            Width           =   2055
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFC0C0&
            Caption         =   "DEPARTURE TIME"
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
            TabIndex        =   19
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFC0C0&
            Caption         =   "DESTINATION TIME"
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
            TabIndex        =   18
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FFC0C0&
            Caption         =   "DURATION"
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
            TabIndex        =   17
            Top             =   2760
            Width           =   2055
         End
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "TRAIN DETAILS"
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
         Left            =   6480
         TabIndex        =   12
         Top             =   1320
         Width           =   2175
      End
   End
End
Attribute VB_Name = "DETAILS"
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
Text8.Text = ""
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
Text1.Text = rsv.Fields("TRAIN_NO")
Text2.Text = rsv.Fields("TRAIN_NAME")
Text3.Text = rsv.Fields("STARTING_POINT")
Text4.Text = rsv.Fields("DESTINATION")
Text5.Text = rsv.Fields("DEPARTURE_TIME")
Text6.Text = rsv.Fields("DESTINATION_TIME")
Text7.Text = rsv.Fields("DURATION")
Text8.Text = rsv.Fields("AMOUNT")
If (rsv.State = 1) Then
rsv.Close
End If
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

