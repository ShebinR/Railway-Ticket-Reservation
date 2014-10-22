VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   Picture         =   "fina.frx":0000
   ScaleHeight     =   9435
   ScaleWidth      =   15420
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5640
      TabIndex        =   11
      Top             =   8040
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   11280
      Top             =   600
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "TICKET"
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
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   5175
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
         Height          =   375
         Left            =   2760
         TabIndex        =   24
         Top             =   1920
         Width           =   1815
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
         Height          =   375
         Left            =   2760
         TabIndex        =   14
         Top             =   4800
         Width           =   1815
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
         Height          =   405
         Left            =   2760
         TabIndex        =   10
         Top             =   4080
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
         Height          =   405
         Left            =   2760
         TabIndex        =   9
         Top             =   3360
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
         Left            =   2760
         TabIndex        =   8
         Top             =   2640
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
         Height          =   405
         Left            =   2760
         TabIndex        =   7
         Top             =   1200
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
         Height          =   405
         Left            =   2760
         TabIndex        =   6
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label13 
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
         Left            =   360
         TabIndex        =   23
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "SEAT NO/WL NO"
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
         Left            =   360
         TabIndex        =   15
         Top             =   4920
         Width           =   1815
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
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label5 
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
         Left            =   360
         TabIndex        =   4
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label4 
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
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label2 
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
         Left            =   360
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label11 
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
         Left            =   360
         TabIndex        =   1
         Top             =   2640
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   9135
      Left            =   0
      Picture         =   "fina.frx":CAC1
      ScaleHeight     =   9075
      ScaleWidth      =   15915
      TabIndex        =   12
      Top             =   0
      Width           =   15975
      Begin VB.Frame Frame2 
         Caption         =   "TICKET DETAILS"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5535
         Left            =   8760
         TabIndex        =   16
         Top             =   1800
         Width           =   5775
         Begin VB.TextBox Text12 
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
            Left            =   3480
            TabIndex        =   28
            Top             =   4440
            Width           =   1695
         End
         Begin VB.TextBox Text11 
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
            Left            =   3480
            TabIndex        =   27
            Top             =   3480
            Width           =   1695
         End
         Begin VB.TextBox Text10 
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
            Left            =   3480
            TabIndex        =   26
            Top             =   2520
            Width           =   1695
         End
         Begin VB.TextBox Text9 
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
            Left            =   3480
            TabIndex        =   25
            Top             =   1560
            Width           =   1695
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
            Height          =   405
            Left            =   3480
            TabIndex        =   21
            Top             =   600
            Width           =   1695
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
            Left            =   480
            TabIndex        =   22
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label12 
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
            Height          =   495
            Left            =   480
            TabIndex        =   20
            Top             =   4440
            Width           =   2055
         End
         Begin VB.Label Label10 
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
            Height          =   495
            Left            =   480
            TabIndex        =   19
            Top             =   3480
            Width           =   2055
         End
         Begin VB.Label Label9 
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
            TabIndex        =   18
            Top             =   2520
            Width           =   2055
         End
         Begin VB.Label Label8 
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
            TabIndex        =   17
            Top             =   1560
            Width           =   2055
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "--------YOUR TICKET--------"
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   4440
         TabIndex        =   13
         Top             =   600
         Width           =   5295
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conv As ADODB.Connection
Dim rsv As ADODB.Recordset
Dim cmdv As ADODB.Command

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

