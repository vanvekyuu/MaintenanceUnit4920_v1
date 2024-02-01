VERSION 5.00
Begin VB.Form Laptop 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "[ID:192939P] You have successfully logged in..."
   ClientHeight    =   3945
   ClientLeft      =   10110
   ClientTop       =   855
   ClientWidth     =   9030
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Laptop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   9030
   Visible         =   0   'False
   Begin VB.Label label_Player 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Media Player"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3300
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label label_Document 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Documents"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1530
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label label_Library 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Library"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   600
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image imgPlayer 
      Height          =   495
      Left            =   3480
      Picture         =   "Laptop.frx":1CF61
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgDoc 
      Height          =   495
      Left            =   1680
      Picture         =   "Laptop.frx":1F470
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgLibrary 
      Height          =   495
      Left            =   600
      Picture         =   "Laptop.frx":20389
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3720
      Left            =   120
      Picture         =   "Laptop.frx":214B7
      Stretch         =   -1  'True
      Top             =   120
      Width           =   8805
   End
End
Attribute VB_Name = "Laptop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub imgDoc_Click()

Documents.Show

End Sub

Private Sub imgLibrary_Click()

Library.Show

End Sub

Private Sub imgPlayer_Click()

MediaPlayer.Show

End Sub
