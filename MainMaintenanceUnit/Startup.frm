VERSION 5.00
Begin VB.Form Startup 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Loading..."
   ClientHeight    =   4245
   ClientLeft      =   900
   ClientTop       =   3240
   ClientWidth     =   7755
   Icon            =   "Startup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Startup.frx":1CF61
   MousePointer    =   4  'Icon
   Moveable        =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7755
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Left            =   720
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   120
   End
   Begin VB.TextBox SpeechBox 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "complete"
      Top             =   3840
      Width           =   7455
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3120
      Left            =   120
      Picture         =   "Startup.frx":490D0
      Stretch         =   -1  'True
      Top             =   600
      Width           =   7485
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
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   7470
   End
End
Attribute VB_Name = "Startup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InputData
Dim LOADING As String
Dim X As Integer

Private Sub Form_Load()

LOADING = "c:\MainMaintenanceUnit\text\loading.txt"

Open LOADING For Input As #1

End Sub

Private Sub Option1_Click()

Form1.Show


End Sub


Private Sub Timer1_Timer()



On Error Resume Next
Line Input #1, InputData
SpeechBox.Text = InputData

If SpeechBox.Text = "complete" Then

Close #1
Timer2.Enabled = True
Timer2.Interval = 5000
Main.Show
Startup.Hide
Timer2.Enabled = False
Unload Startup

End If




End Sub

