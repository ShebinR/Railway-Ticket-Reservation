VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   8445
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15465
   LinkTopic       =   "Form3"
   ScaleHeight     =   8445
   ScaleWidth      =   15465
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   8535
      Left            =   0
      Picture         =   "EXIT.frx":0000
      ScaleHeight     =   8475
      ScaleWidth      =   15195
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "CLICK TO LEAVE......"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         Picture         =   "EXIT.frx":47D06
         TabIndex        =   7
         Top             =   7440
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "SHEBIN R............2011103534"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   5160
         TabIndex        =   6
         Top             =   5280
         Width           =   5175
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "PRAVEEN SK.......2011103527"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   5160
         TabIndex        =   5
         Top             =   4680
         Width           =   5175
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "VIGNESH NT......2011103544"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   5160
         TabIndex        =   4
         Top             =   4080
         Width           =   5175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "PROJECT DONE BY  :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   4080
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "RAILWAY TICKET RESERVATION SYSTEM"
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   1920
         TabIndex        =   2
         Top             =   3000
         Width           =   9615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "THANK YOU FOR USING"
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   3960
         TabIndex        =   1
         Top             =   2280
         Width           =   6135
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
