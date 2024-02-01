VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9720
   FillColor       =   &H00FFC0FF&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   9720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Help"
      Height          =   735
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
      Width           =   2295
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Next"
      Height          =   855
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox TextBox 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   6975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
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
      Left            =   7200
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
   Begin VB.Image Image5 
      Height          =   2520
      Left            =   4800
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1680
   End
   Begin VB.Image Image4 
      Height          =   2520
      Left            =   4800
      Picture         =   "Form1.frx":2A67
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1680
   End
   Begin VB.Image Image3 
      Height          =   2640
      Left            =   3120
      Picture         =   "Form1.frx":505F
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1530
   End
   Begin VB.Image Image2 
      Height          =   1635
      Left            =   360
      Picture         =   "Form1.frx":7771
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   2010
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3360
      Left            =   120
      Picture         =   "Form1.frx":81DB
      Stretch         =   -1  'True
      Top             =   600
      Width           =   7005
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
      Top             =   0
      Width           =   9375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Image3_Click()

Play.Show
Form1.Hide

End Sub

