VERSION 5.00
Object = "*\AProject1.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2325
   ClientLeft      =   1020
   ClientTop       =   1815
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   ScaleHeight     =   2325
   ScaleWidth      =   4305
   Begin Project1.VUAudioMeter VUAudioMeter1 
      Height          =   1050
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   1852
      BorderStyle     =   0
      Picture         =   "Form1.frx":0000
   End
   Begin Project1.VUAudioMeter VUAudioMeter1 
      Height          =   1050
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   1852
      BorderStyle     =   1
      Picture         =   "Form1.frx":13CC6
   End
   Begin Project1.VUAudioMeter VUAudioMeter1 
      Height          =   1050
      Index           =   2
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   1852
      BorderStyle     =   2
      Picture         =   "Form1.frx":2798C
   End
   Begin Project1.VUAudioMeter VUAudioMeter1 
      Height          =   1050
      Index           =   3
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   1852
      BorderStyle     =   3
      Picture         =   "Form1.frx":3B652
   End
   Begin Project1.VUAudioMeter VUAudioMeter1 
      Height          =   1050
      Index           =   4
      Left            =   2880
      TabIndex        =   5
      Top             =   120
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   1852
      Picture         =   "Form1.frx":4F318
   End
   Begin Project1.VUAudioMeter VUAudioMeter1 
      Height          =   1050
      Index           =   5
      Left            =   3480
      TabIndex        =   6
      Top             =   120
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   1852
      BorderStyle     =   5
      Picture         =   "Form1.frx":62FDE
   End
   Begin VB.Label Label1 
      Caption         =   "If this does not work, Your sound card does not support the VU Meter on your mixer control. Sorry, For any inconvenience."
      Height          =   915
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   4050
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub HScroll1_Scroll()
'ProgressMeter1.Value = HScroll1.Value
End Sub

Private Sub Form_Load()
Dim i As Integer
For i = 0 To 5
    VUAudioMeter1(i).Enabled = True
Next
End Sub
