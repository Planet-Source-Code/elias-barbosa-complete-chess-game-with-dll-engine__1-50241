VERSION 5.00
Begin VB.Form frmMoveList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Move list"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4215
   Begin VB.TextBox txtMoveList 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmMoveList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If the user pressed a key while the
    '"Move List" window has the focus,
    'pass the key pressed and the focus to
    'the "Chess Board" window.
    Call frmChessBoard.Form_KeyPress(KeyAscii)
    frmChessBoard.SetFocus
    
End Sub

Private Sub txtMoveList_Change()
    'Keep the text cursor at the
    'end of the TextBox...
    txtMoveList.SelStart = Len(txtMoveList.Text)
    
End Sub

Private Sub txtMoveList_KeyPress(KeyAscii As Integer)
    'If the user pressed a key while the
    '"txtMoveList" TextBox has the focus,
    'pass the key pressed and the focus to
    'the "Chess Board" window.
    Call frmChessBoard.Form_KeyPress(KeyAscii)
    frmChessBoard.SetFocus
    
End Sub
