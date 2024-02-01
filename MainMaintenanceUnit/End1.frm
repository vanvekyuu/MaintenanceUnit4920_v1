VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form End1 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   0  'None
   Caption         =   "msg"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
   Enabled         =   0   'False
   FillColor       =   &H00FFC0FF&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Timer closeTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   120
   End
   Begin VB.Timer end1timer 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   120
      Top             =   120
   End
   Begin VB.TextBox msgbox1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   4815
   End
   Begin WMPLibCtl.WindowsMediaPlayer file1 
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4815
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
      stretchToFit    =   -1  'True
      windowlessVideo =   -1  'True
      enabled         =   0   'False
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   8493
      _cy             =   4895
   End
End
Attribute VB_Name = "End1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InputMessage
Dim MESSAGE As String

Private Sub closeTimer_Timer()

End

End Sub

Private Sub Form_Load()

MESSAGE = "c:\MainMaintenanceUnit\text\finalmsg.txt"
Open MESSAGE For Input As #9

file1.URL = "c:\MainMaintenanceUnit\video\finalmessage.mp4"

Main.mainend1.URL = "c:\MainMaintenanceUnit\video\endmain.mp4"
Main.mainend1.Visible = True
Main.mainend1.Controls.Play
Main.BackColor = &H0&
Main.Shape12.Visible = True
Main.TextBox = "error"
Main.Label1.Visible = False
Main.Option1.Visible = False
Main.Option2.Visible = False
Main.Option3.Visible = False
Main.Option4.Visible = False
Main.cmdNext.Visible = False
Main.cmdHelp.Visible = False

If end1timer.Enabled = True Then
file1.Controls.Play
End If

End Sub

Private Sub end1timer_Timer()

On Error Resume Next
Line Input #9, InputMessage
msgbox1.Text = InputMessage

If msgbox1.Text = "PLEASE EJECT DEVICE FROM BODY" Then
closeTimer.Enabled = True
End If

End Sub
