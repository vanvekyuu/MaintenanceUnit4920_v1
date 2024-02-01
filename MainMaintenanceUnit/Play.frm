VERSION 5.00
Begin VB.Form Play 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Assembly Station"
   ClientHeight    =   5805
   ClientLeft      =   9570
   ClientTop       =   2505
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   9570
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton cmdPlayHelp 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Help"
      Height          =   615
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   120
   End
   Begin VB.Image placedR4 
      Height          =   1095
      Left            =   4080
      Picture         =   "Play.frx":0000
      Stretch         =   -1  'True
      Top             =   3240
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image placedR2 
      Height          =   1095
      Left            =   3360
      Picture         =   "Play.frx":1293
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image placedR3 
      Height          =   1095
      Left            =   1560
      Picture         =   "Play.frx":2535
      Stretch         =   -1  'True
      Top             =   3240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image placedR1 
      Height          =   1095
      Left            =   2160
      Picture         =   "Play.frx":388D
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image cpnR4 
      DragMode        =   1  'Automatic
      Height          =   1455
      Left            =   7560
      Picture         =   "Play.frx":4BDB
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   780
   End
   Begin VB.Image cpnR3 
      DragMode        =   1  'Automatic
      Height          =   1455
      Left            =   6360
      Picture         =   "Play.frx":5E6E
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   780
   End
   Begin VB.Image cpnR2 
      DragMode        =   1  'Automatic
      Height          =   1455
      Left            =   7560
      Picture         =   "Play.frx":71C6
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   780
   End
   Begin VB.Image cpnR1 
      DragMode        =   1  'Automatic
      Height          =   1455
      Left            =   6360
      Picture         =   "Play.frx":8468
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   780
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Components"
      BeginProperty Font 
         Name            =   "Source Code Pro Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   6000
      TabIndex        =   2
      Top             =   600
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   3915
      Left            =   600
      Picture         =   "Play.frx":97B6
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   4905
   End
   Begin VB.Label Label_MaintenanceUnit 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Maintenance Unit #4920"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Automatic Takeover System"
      BeginProperty Font 
         Name            =   "Source Code Pro Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5775
   End
End
Attribute VB_Name = "Play"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InputData
Dim LEAF As String



Private Sub Command1_Click()

Play.Hide
Main.Show

End Sub

Private Sub cmdPlayHelp_Click()

MsgBox ("Unable to display help, going back to Main Room")
Main.Show
Play.Hide
Main.TextBox.Text = ""
Main.TextBox.Tag = act2
Main.cmdNext.Visible = True

End Sub


Private Sub Form_Load()

LEAF = "c:\MainMaintenanceUnit\text\left.txt"

Open LEAF For Input As #3

End Sub
Private Sub Form_Unload(Cancel As Integer)
Close #3
End Sub

Private Sub placedR4_Click()

Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()

On Error Resume Next
Line Input #3, InputData
Play.Left = InputData

End Sub
