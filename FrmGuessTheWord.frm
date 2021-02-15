VERSION 5.00
Begin VB.Form FrmGuessTheWord 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   ClientHeight    =   600
   ClientLeft      =   2370
   ClientTop       =   6720
   ClientWidth     =   9495
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   600
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      Height          =   735
      Left            =   7080
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox TxtName 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   480
      Left            =   5280
      TabIndex        =   0
      Top             =   90
      Width           =   3015
   End
   Begin VB.Label lblGuess 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Guess the Word:"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "FrmGuessTheWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub cmdOK_Click()
GWD = TxtName.Text
TxtName.Text = UCase(Trim(""))
FrmGuessTheWord.Hide
End Sub


Private Sub Form_Activate()
FrmGuessTheWord.Left = Flip.Left + 50
FrmGuessTheWord.Top = (Flip.Top + Flip.Height) - FrmGuessTheWord.Height
FrmGuessTheWord.Width = Flip.Width - 100
End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)
Dim VX
VX = Chr(KeyAscii)
KeyAscii = Asc(UCase(VX))
End Sub


