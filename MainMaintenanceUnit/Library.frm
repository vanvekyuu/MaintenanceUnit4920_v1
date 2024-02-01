VERSION 5.00
Begin VB.Form Library 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Library"
   ClientHeight    =   3015
   ClientLeft      =   12735
   ClientTop       =   390
   ClientWidth     =   3855
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Library.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3855
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   5
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   4
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   3
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   2
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   1
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   3375
   End
   Begin VB.CommandButton btnPicture 
      BackColor       =   &H00FFFF80&
      Caption         =   "Picture"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton btnVideo 
      BackColor       =   &H00FFFF80&
      Caption         =   "Video"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton btnDocument 
      BackColor       =   &H00FFFF80&
      Caption         =   "Documents"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1215
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
Attribute VB_Name = "Library"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
