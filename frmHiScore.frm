VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmHiScore 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Flip! Best Scores"
   ClientHeight    =   7275
   ClientLeft      =   1905
   ClientTop       =   795
   ClientWidth     =   7095
   ControlBox      =   0   'False
   Icon            =   "frmHiScore.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7275
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtBtime1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5880
      MultiLine       =   -1  'True
      TabIndex        =   32
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtTname2 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   31
      Top             =   3360
      Width           =   2895
   End
   Begin VB.TextBox txtBTime2 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5880
      MultiLine       =   -1  'True
      TabIndex        =   30
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtTname3 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   29
      Top             =   3720
      Width           =   2895
   End
   Begin VB.TextBox txtBtime3 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5880
      MultiLine       =   -1  'True
      TabIndex        =   28
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtTname1 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   27
      Top             =   3000
      Width           =   2895
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Text            =   "Intermediate:"
      Top             =   3360
      Width           =   2775
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      TabIndex        =   25
      Text            =   "Expert:"
      Top             =   3720
      Width           =   2775
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Text            =   "Beginner:"
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00008000&
      Caption         =   "Best Times"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1575
      Left            =   120
      TabIndex        =   23
      Top             =   2640
      Width           =   6855
   End
   Begin VB.PictureBox a2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   22
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox A3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   6240
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   21
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox A1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   2880
      ScaleHeight     =   735
      ScaleWidth      =   1215
      TabIndex        =   20
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   56
      Left            =   1200
      Top             =   360
   End
   Begin VB.TextBox txtwsc1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5880
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox txtwsc2 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5880
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox txtwsc3 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5880
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox txtbeg2 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   16
      Top             =   4680
      Width           =   2895
   End
   Begin VB.TextBox txtInt2 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   15
      Top             =   5040
      Width           =   2895
   End
   Begin VB.TextBox txtExp2 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   14
      Top             =   5400
      Width           =   2895
   End
   Begin VB.TextBox txtBsc3 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5880
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtBsc2 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5880
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtExp1 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox txtInt1 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox txtBsc1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5880
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox TxtBeg1 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Text            =   "Expert:"
      Top             =   2040
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Text            =   "Expert:"
      Top             =   5400
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Text            =   "Intermediate:"
      Top             =   5040
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Text            =   "Intermediate:"
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00008000&
      Caption         =   "Wall of Shame"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   6855
      Begin VB.TextBox Text2 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Text            =   "Beginner:"
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "Beginner:"
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Hall of Fame"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   6855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   6120
      Width           =   6855
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1920
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   86
      ImageHeight     =   42
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmHiScore.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmHiScore.frx":15CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmHiScore.frx":288E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmHiScore.frx":3B50
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmHiScore.frx":4E12
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmHiScore.frx":60D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmHiScore.frx":7396
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmHiScore.frx":8658
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmHiScore.frx":991A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmHiScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim S1 As Integer
Dim BS As Integer

Private Sub cmdClear_Click()
Dim AA, AB, AC, AD
AA = vbYesNo + vbQuestion
AB = "Clear Score Table"
AC = "Are you sure?"
AD = MsgBox(AC, AA, AB)
If AD = vbNo Then GoTo 15
If AD = vbYes Then GoTo 10
10
LB = "JaySoft"
LC = "JaySoft"
LD = "JaySoft"
LE = 100
LF = 200
LG = 300
LH = "Anon"
LI = "Anon"
LJ = "Anon"
LK = 101
LL = 201
LM = 301
JL = 999
JM = 999
JN = 999
LO = "JaySoft"
LP = "JaySoft"
LQ = "JaySoft"
TxtBeg1.Text = LB
txtInt1.Text = LC
txtExp1.Text = LD
txtBsc1.Text = Str(LE)
txtBsc2.Text = Str(LF)
txtBsc3.Text = Str(LG)
txtbeg2.Text = LH
txtInt2.Text = LI
txtExp2.Text = LJ
txtwsc1.Text = Str(LK)
txtwsc2.Text = Str(LL)
txtwsc3.Text = Str(LM)
txtTname1.Text = LO
txtTname2.Text = LP
txtTname3.Text = LQ
TxtBtime1.Text = Str(JL)
txtBTime2.Text = Str(JM)
txtBtime3.Text = Str(JN)
15
End Sub


Private Sub cmdOK_Click()

End Sub

Private Sub Form_Load()
Timer1.Enabled = True
S1 = 1
BS = 0
a2.Picture = ImageList1.ListImages(9).Picture
A3.Picture = ImageList1.ListImages(9).Picture
Text7.ForeColor = RGB(0, 255, 0)
Text8.ForeColor = RGB(0, 255, 0)
Text9.ForeColor = RGB(0, 255, 0)
Text6.ForeColor = RGB(0, 255, 0)
Text3.ForeColor = RGB(0, 255, 0)
Text1.ForeColor = RGB(0, 255, 0)
Text2.ForeColor = RGB(0, 255, 0)
Text4.ForeColor = RGB(0, 255, 0)
Text5.ForeColor = RGB(0, 255, 0)
Frame1.ForeColor = RGB(255, 255, 0)
Frame2.ForeColor = RGB(255, 255, 0)
Frame3.ForeColor = RGB(255, 255, 0)
txtBsc1.ForeColor = RGB(255, 255, 255)
txtBsc2.ForeColor = RGB(255, 255, 255)
txtBsc3.ForeColor = RGB(255, 255, 255)
TxtBeg1.ForeColor = RGB(255, 255, 255)
txtInt1.ForeColor = RGB(255, 255, 255)
txtExp1.ForeColor = RGB(255, 255, 255)
txtTname1.ForeColor = RGB(255, 255, 255)
txtTname2.ForeColor = RGB(255, 255, 255)
txtTname3.ForeColor = RGB(255, 255, 255)
TxtBtime1.ForeColor = RGB(255, 255, 255)
txtBTime2.ForeColor = RGB(255, 255, 255)
txtBtime3.ForeColor = RGB(255, 255, 255)
txtbeg2.ForeColor = RGB(255, 255, 255)
txtInt2.ForeColor = RGB(255, 255, 255)
txtExp2.ForeColor = RGB(255, 255, 255)
txtwsc1.ForeColor = RGB(255, 255, 255)
txtwsc2.ForeColor = RGB(255, 255, 255)
txtwsc3.ForeColor = RGB(255, 255, 255)
TxtBeg1.Text = LB
txtInt1.Text = LC
txtExp1.Text = LD
txtBsc1.Text = Str(LE)
txtBsc2.Text = Str(LF)
txtBsc3.Text = Str(LG)
txtbeg2.Text = LH
txtInt2.Text = LI
txtExp2.Text = LJ
txtwsc1.Text = Str(LK)
txtwsc2.Text = Str(LL)
txtwsc3.Text = Str(LM)
txtTname1.Text = LO
txtTname2.Text = LP
txtTname3.Text = LQ
TxtBtime1.Text = Str(JL)
txtBTime2.Text = Str(JM)
txtBtime3.Text = Str(JN)
End Sub


Private Sub Text10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Label1_Click()
frmHiScore.Hide
End Sub

Private Sub Timer1_Timer()
A1.Picture = ImageList1.ListImages(S1).Picture
If BS = 0 Then
S1 = S1 + 1
End If
If S1 = 9 Then
BS = 1
End If
If BS = 1 Then
S1 = S1 - 1
End If
If S1 = 1 Then
BS = 0
End If
End Sub

