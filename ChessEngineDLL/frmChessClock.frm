VERSION 5.00
Begin VB.Form frmChessClock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clock"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   4980
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2280
      Top             =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Black:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "White:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   660
   End
   Begin VB.Label lblBlacksClock 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblWhitesClock 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmChessClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Whites As Clock
Dim Blacks As Clock

Public Sub subStartChessClock()
    Timer1.Enabled = True
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    Call frmChessBoard.Form_KeyPress(KeyAscii)
    frmChessBoard.SetFocus
    
End Sub

Private Sub Timer1_Timer()
    'kateremu se štejejo sekunde?
    
    'If it is Blacks' turn, start its clock.
    If (bolBlackTurn = True) Then
        
        'prištejemo sekundo
        
        'Calculate time for Blacks...
        With Blacks
            .Second = .Second + 1
            
            If (.Second = 60) Then
                .Second = 0
                .Minute = .Minute + 1
            End If
            
            If (.Minute = 60) Then
                .Minute = 0
                .Hour = .Hour + 1
            End If
            
        End With
        
    'If it is Whites' turn, start its clock.
    Else
    
        'Calculate time for Whites...
        With Whites
            .Second = .Second + 1
            
            If (.Second = 60) Then
                .Second = 0
                .Minute = .Minute + 1
            End If
            
            If (.Minute = 60) Then
                .Minute = 0
                .Hour = .Hour + 1
            End If
                
        End With
        
    End If
    
    'izpišemo oba èasa

    lblWhitesClock.Caption = fntFormatTime(Whites.Hour, Whites.Minute, Whites.Second)
    lblBlacksClock.Caption = fntFormatTime(Blacks.Hour, Blacks.Minute, Blacks.Second)
    
End Sub

'This function will format the time
'provided so it looks neat on the
'Chess Clock...
Private Function fntFormatTime(ByVal bytHours As Byte, ByVal bytMinutes As Byte, ByVal bytSecunds As Byte) As String
    
    If (bytHours < 10) Then
        fntFormatTime = fntFormatTime & "0" & bytHours & ":"
        
    Else
        fntFormatTime = fntFormatTime & bytHours & ":"
        
    End If
        
    If (bytMinutes < 10) Then
        fntFormatTime = fntFormatTime & "0" & bytMinutes & ":"
        
    Else
        fntFormatTime = fntFormatTime & bytMinutes & ":"
        
    End If
    
    If (bytSecunds < 10) Then
        fntFormatTime = fntFormatTime & "0" & bytSecunds
        
    Else
        fntFormatTime = fntFormatTime & bytSecunds
        
    End If
    
End Function

'No comments :)
Public Sub subStopClock()
    Timer1.Enabled = False
    
End Sub

'Reset both Clocks...
Public Sub subResetClocks()
    lblWhitesClock.Caption = "00:00:00"
    lblBlacksClock.Caption = "00:00:00"
    
    Blacks.Minute = 0
    Blacks.Hour = 0
    Blacks.Second = 0
    
    Whites.Minute = 0
    Whites.Second = 0
    Whites.Hour = 0
    
End Sub
