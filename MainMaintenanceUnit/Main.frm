VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Main 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Hello, Engineer"
   ClientHeight    =   5805
   ClientLeft      =   -30
   ClientTop       =   2505
   ClientWidth     =   9570
   FillColor       =   &H00FFC0FF&
   FillStyle       =   0  'Solid
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   9570
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.TextBox PasswordBox 
      Height          =   285
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   13
      Top             =   3000
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Help"
      Height          =   735
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Next"
      Height          =   735
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   6
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
      Height          =   1605
      Left            =   120
      TabIndex        =   5
      Tag             =   "act1"
      Top             =   4080
      Width           =   6975
   End
   Begin VB.CommandButton Option4 
      BackColor       =   &H00FFC0FF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "3"
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Option3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "2"
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Option2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "1"
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Option1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      MaskColor       =   &H00FFC0FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "0"
      Top             =   1680
      Width           =   615
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   0
      Top             =   4200
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label LabelReCal 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Index           =   3
      Left            =   4680
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label LabelReCal 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Index           =   2
      Left            =   3960
      TabIndex        =   11
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label LabelReCal 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Index           =   1
      Left            =   3240
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label LabelReCal 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Index           =   0
      Left            =   2520
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image5 
      Height          =   2520
      Left            =   4800
      Picture         =   "Main.frx":1CF61
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1680
   End
   Begin VB.Image Image4 
      Height          =   2520
      Left            =   4800
      Picture         =   "Main.frx":1F9C8
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1680
   End
   Begin VB.Image Image3 
      Height          =   2640
      Left            =   3120
      Picture         =   "Main.frx":21FC0
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1530
   End
   Begin VB.Image Image2 
      Height          =   1635
      Left            =   360
      Picture         =   "Main.frx":246D2
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   2010
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
      TabIndex        =   8
      Top             =   600
      Width           =   2175
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3360
      Left            =   120
      Picture         =   "Main.frx":2513C
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
      Top             =   120
      Width           =   9255
   End
   Begin WMPLibCtl.WindowsMediaPlayer mainend1 
      Height          =   4335
      Left            =   0
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   7215
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   12726
      _cy             =   7646
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InputData
Dim InputData2
Dim InputData3
Dim InputData4
Dim START1 As String
Dim PLAYERR As String
Dim ACTIONS As String


Private Sub cmdHelp_Click()

MsgBox ("Unable to display help")

End Sub

Private Sub Command1_Click()

Main.Hide

End Sub

Private Sub Form_Load()

START1 = "c:\MainMaintenanceUnit\text\start1.txt"
PLAYERR = "c:\MainMaintenanceUnit\text\playerror.txt"
ACTIONS = "c:\MainMaintenanceUnit\text\actions.txt"

Open START1 For Input As #2


If TextBox.Tag = act2 Then
Close #2
End If


End Sub
Private Sub cmdNext_Click()

On Error Resume Next
Line Input #2, InputData
TextBox.Text = InputData
If TextBox.Text = "we shall start when you are ready" Then
'StartUpPosition = Manual
'Left = -150
End If

If TextBox.Tag = act2 Then
Open PLAYERR For Input As #4
On Error Resume Next
Line Input #4, InputData2
TextBox.Text = InputData2


If TextBox.Text = "simply press the buttons on the right in the order shown on the screen" Then
cmdNext.Visible = False
LabelReCal(0).Visible = True
LabelReCal(1).Visible = True
LabelReCal(2).Visible = True
LabelReCal(3).Visible = True
LabelReCal(0).Caption = Int(Rnd * 3)
LabelReCal(1).Caption = Int(Rnd * 3)
LabelReCal(2).Caption = Int(Rnd * 3)
LabelReCal(3).Caption = Int(Rnd * 3)

Open ACTIONS For Output As #5
Print #5, LabelReCal(0).Caption + LabelReCal(1).Caption + LabelReCal(2).Caption + LabelReCal(3).Caption
Close #5

'Open LABEL0 For Output As #5
'Print #5, LabelReCal(0).Caption
'MsgBox InputData + "" + Image6(Index).Tag
'Close #5

'Open LABELA For Output As #6
'Print #6, LabelReCal(1).Caption
'MsgBox InputData + "" + Image6(Index).Tag
'Close #6

'Open LABEL2 For Output As #7
'Print #7, LabelReCal(2).Caption
'MsgBox InputData + "" + Image6(Index).Tag
'Close #7

'Open LABEL3 For Output As #8
'Print #8, LabelReCal(3).Caption
'MsgBox InputData + "" + Image6(Index).Tag
'Close #8

End If
End If
'If cmdNext.Tag = infected Then
'RECAL = "c:\MainMaintenanceUnit\text\afterRecal.txt"
'Open RECAL For Input As #7
'On Error Resume Next
'Line Input #7, InputData4
'TextBox.Text = InputData4
'End If

If TextBox.Text = "great job! we can now get back into" Then
End1.Show
End1.end1timer.Enabled = True
'ShakeTimer.Enabled = True
'cmdNext.Visible
'Close #4
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Close #2
Close #4
End Sub

Private Sub Image2_Click()

Laptop.Show

End Sub

Private Sub Image3_Click()

'MsgBox "Error"
Play.Visible = True
Main.Visible = False

End Sub
Private Sub Option1_Click()

PasswordBox = PasswordBox & 0

End Sub

Private Sub Option2_Click()

PasswordBox = PasswordBox & 1

End Sub

Private Sub Option3_Click()

PasswordBox = PasswordBox & 2

End Sub

Private Sub Option4_Click()

PasswordBox = PasswordBox & 3

End Sub

Private Sub PasswordBox_Change()

If Len(PasswordBox) = 4 Then
Open ACTIONS For Input As #6
Line Input #6, InputData3
If PasswordBox.Text = InputData3 Then
MsgBox ("Maintenance Unit Recalibrated")
Close #6
cmdNext.Visible = True
LabelReCal(0).Visible = False
LabelReCal(1).Visible = False
LabelReCal(2).Visible = False
LabelReCal(3).Visible = False
Else
MsgBox ("Error, retry")
PasswordBox = ""
Close #6
End If

End If

End Sub

Private Sub ShakeTimer_Timer()


'Dim InputLeft
'Dim InputTop
'Dim SHAKEX As String
'Dim SHAKEY As String

'SHAKEX = "c:\MainMaintenanceUnit\text\left.txt"
'SHAKEY = "c:\MainMaintenanceUnit\text\top.txt"

'Open SHAKEX For Input As #7 'error file already open or sumthin ah im gna sleep
'Open SHAKEY For Input As #8
'On Error Resume Next
'Line Input #7, InputLeft
'Main.Left = InputLeft
'Line Input #8, InputTop
'Main.Top = InputTop
'Close #7

'Close #8



End Sub
