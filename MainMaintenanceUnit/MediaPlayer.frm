VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form MediaPlayer 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Media Player"
   ClientHeight    =   3015
   ClientLeft      =   15225
   ClientTop       =   390
   ClientWidth     =   3855
   Icon            =   "MediaPlayer.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3855
   Visible         =   0   'False
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3375
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
      _cx             =   5953
      _cy             =   4471
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFF80&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   2775
      Left            =   120
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "MediaPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub WindowsMediaPlayer1_OpenStateChange(ByVal NewState As Long)

End Sub
