VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form frmMiniChessBoard 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PC thinking"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picChessBoard 
      AutoSize        =   -1  'True
      Height          =   2790
      Left            =   0
      Picture         =   "frmMiniChessBoard.frx":0000
      ScaleHeight     =   2730
      ScaleWidth      =   2745
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2805
   End
   Begin PicClip.PictureClip pclWhitePictureClip 
      Left            =   0
      Top             =   3360
      _ExtentX        =   3651
      _ExtentY        =   1244
      _Version        =   393216
      Rows            =   2
      Cols            =   6
      Picture         =   "frmMiniChessBoard.frx":115A
   End
   Begin PicClip.PictureClip pclBlackPictureClip 
      Left            =   0
      Top             =   4080
      _ExtentX        =   3651
      _ExtentY        =   1244
      _Version        =   393216
      Rows            =   2
      Cols            =   6
      Picture         =   "frmMiniChessBoard.frx":1560
   End
   Begin VB.Image imgWhitePiece 
      Height          =   375
      Left            =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgBlackPiece 
      Height          =   375
      Left            =   480
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Na potezi je beli"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4920
      Width           =   1335
   End
End
Attribute VB_Name = "frmMiniChessBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intMiniMoveCount As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If the user pressed a key while the
    '"Engine Thinking" window has the focus,
    'pass the key pressed and the focus to
    'the "Chess Board" window.
    Call frmChessBoard.Form_KeyPress(KeyAscii)
    frmChessBoard.SetFocus
    
End Sub

Private Sub Form_Load()
    'Reset the "Engine Thinking" setting...
    
    bolBlackTurn = False
    intMiniMoveCount = 1
    aryMiniCaptPiecesCount(0) = 0
    aryMiniCaptPiecesCount(1) = 0
    
End Sub
