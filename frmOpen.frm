VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmOpen 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Initializing Flip!"
   ClientHeight    =   1620
   ClientLeft      =   1710
   ClientTop       =   390
   ClientWidth     =   9525
   ControlBox      =   0   'False
   FontTransparent =   0   'False
   Icon            =   "frmOpen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmOpen.frx":030A
   ScaleHeight     =   1620
   ScaleWidth      =   9525
   Begin VB.PictureBox Picture11 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   1200
      Picture         =   "frmOpen.frx":0750
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   10
      Top             =   480
      Width           =   735
   End
   Begin VB.PictureBox Picture10 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   480
      Picture         =   "frmOpen.frx":1586
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   9
      Top             =   480
      Width           =   735
   End
   Begin VB.PictureBox Picture9 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   5520
      Picture         =   "frmOpen.frx":23BC
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   8
      Top             =   480
      Width           =   735
   End
   Begin VB.PictureBox Picture8 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   300
      Left            =   1680
      Picture         =   "frmOpen.frx":31F2
      ScaleHeight     =   300
      ScaleWidth      =   3525
      TabIndex        =   7
      Top             =   1320
      Width           =   3525
   End
   Begin VB.PictureBox Picture7 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   315
      Left            =   2040
      Picture         =   "frmOpen.frx":48A4
      ScaleHeight     =   315
      ScaleWidth      =   2670
      TabIndex        =   6
      Top             =   120
      Width           =   2670
   End
   Begin VB.PictureBox Picture6 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   735
      Left            =   4080
      Picture         =   "frmOpen.frx":5BAA
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.PictureBox Picture5 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   735
      Left            =   1920
      Picture         =   "frmOpen.frx":69E0
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   735
      Left            =   4800
      Picture         =   "frmOpen.frx":7816
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   735
      Left            =   3360
      Picture         =   "frmOpen.frx":864C
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   735
      Left            =   2640
      Picture         =   "frmOpen.frx":9482
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1320
      Left            =   6600
      Picture         =   "frmOpen.frx":A2B8
      ScaleHeight     =   1320
      ScaleWidth      =   2655
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   4200
      TabIndex        =   11
      Top             =   600
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Form_Activate()
frmOpen.Top = 0
frmOpen.Left = (Screen.Width / 2) - (frmOpen.Width / 2)
Flip.Show
End Sub

Private Sub Form_Load()
With frmOpen.MMControl1
        .Shareable = False
        .Visible = False
        .DeviceType = "Sequencer"
        .Command = "close"
        .FileName = "0.MID"
        .Command = "Open"
        .Command = "play"
        End With
End Sub


