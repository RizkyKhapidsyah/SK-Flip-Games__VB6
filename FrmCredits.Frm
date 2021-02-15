VERSION 5.00
Begin VB.Form FrmAbout 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flip!"
   ClientHeight    =   2790
   ClientLeft      =   2100
   ClientTop       =   1845
   ClientWidth     =   5250
   ControlBox      =   0   'False
   Icon            =   "FrmCredits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2790
   ScaleWidth      =   5250
   Begin VB.CommandButton Command1 
      Caption         =   "&J"
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      X1              =   240
      X2              =   5040
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Jason Bennison   ""JaySoft"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   890
      Width           =   5055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Technical Consultants: Sammy 'n Steve"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1800
      Width           =   5295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Graphics (Portions) by:  Belinda Bennison"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   5295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      X1              =   240
      X2              =   5040
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Flip! for Windows"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Programming and Musical Adaptations by:"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5055
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
FrmAbout.Hide
End Sub

Private Sub Command1_Click()
Dim tit
Dim mes
Dim dia
Dim res
mes = "Hey Jason;  Stay Kool!"
tit = "Message from Me to Me!"
dia = vbOK + vbInformation
res = MsgBox(mes, dia, tit)
End Sub


Private Sub Form_Activate()
FrmAbout.Left = (Screen.Width / 2) - (FrmAbout.Width / 2)
FrmAbout.Top = (Screen.Height / 2) - (FrmAbout.Height / 2)
End Sub

