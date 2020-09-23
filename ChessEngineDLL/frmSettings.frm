VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.HScrollBar srbPlayingStyle 
      Height          =   255
      Left            =   240
      Max             =   10
      TabIndex        =   1
      Top             =   1560
      Width           =   4215
   End
   Begin VB.HScrollBar srbGameLevel 
      Height          =   255
      Left            =   240
      Max             =   5
      Min             =   1
      TabIndex        =   0
      Top             =   600
      Value           =   1
      Width           =   4215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Attack:"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Defence"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblDepthLevel 
      Caption         =   "Depth level:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Me.Hide
    
End Sub

Private Sub srbGameLevel_Change()
    lblDepthLevel.Caption = "Game Level: " & srbGameLevel.Value
    
End Sub

Private Sub srbGameLevel_Scroll()
    lblDepthLevel.Caption = "Game Level: " & srbGameLevel.Value
    
End Sub
