VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form frmChessBoard 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chess board"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   8775
   Begin PicClip.PictureClip pclWhitePictureClip 
      Left            =   3120
      Top             =   2040
      _ExtentX        =   6059
      _ExtentY        =   2249
      _Version        =   393216
      Rows            =   2
      Cols            =   6
      Picture         =   "frmChessBoard.frx":0000
   End
   Begin PicClip.PictureClip pclBlackPictureClip 
      Left            =   2640
      Top             =   3720
      _ExtentX        =   6059
      _ExtentY        =   2249
      _Version        =   393216
      Rows            =   2
      Cols            =   6
      Picture         =   "frmChessBoard.frx":515A
   End
   Begin VB.PictureBox picChessBoard 
      AutoSize        =   -1  'True
      Height          =   6420
      Left            =   0
      Picture         =   "frmChessBoard.frx":A2B4
      ScaleHeight     =   6360
      ScaleWidth      =   6360
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6420
   End
   Begin VB.Image imgBlackPiece 
      Height          =   855
      Left            =   6720
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label lblWhoIsMoving 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "White moves"
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label lblMove 
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
      Left            =   1920
      TabIndex        =   1
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Image imgWhitePiece 
      Height          =   1095
      Left            =   6600
      Top             =   3840
      Width           =   2055
   End
End
Attribute VB_Name = "frmChessBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'sahovnica = frmChessBoard
'msahovnica = picChessBoard
'miniSahovnica = frmMiniChessBoard
'Nastavitve = frmSettings
'MDIForm1 = frmMain
'Poteze = frmMoveList
'statusPoteze = strMoveListStatus
'ministatusPoteze = strMiniMoveListStatus
'pozrte = frmCapturedPieces
'ura = frmChessClock
'sestavipolozaje = fntBuildPositString
'rez = strPosition
'crnarosadamogocalevo = bolBlackCanCastleLeft
'belarosadamogocalevo = bolWhiteCanCastleLeft
'crnarosadamogocadesno = bolBlackCanCastleRight
'belarosadamogocadesno = bolWhiteCanCastleRight
'minicrnarosadamogocalevo = bolMiniBlackCanCastleLeft
'minibelarosadamogocalevo = bolMiniWhiteCanCastleLeft
'minicrnarosadamogocadesno = bolMiniBlackCanCastleRight
'minibelarosadamogocadesno = bolMiniWhiteCanCastleRight
'zapStPot = intMoveCount
'minizapStPot = intMiniMoveCount
'startajUro = subStartChessClock
'movepiece = subMovePiece
'minimovepiece = subMiniMovePiece
'Crnimisli = subStartChessEngine
'mot = objEngine
'stNivojev = srbGameLevel
'Število nivojev: = Game Level:
'obrNap = srbPlayingStyle
'cpoteza = strMove
'ustaviUro = subStopClock
'ComputerMove = bolChessEngineTurn
'poteza = strMove
'minidisplayEmptyBoard = subMiniEmptyBoard
'polozajin = strPosition
'polozaji = aryGamePosition
'minipolozaji = aryMiniGamePosition
'minidrawPiece = subMiniDrawPiece
'katero = strPiece
'kje = strCoordinate
'vrniZnakLinije = funGetCoordinLetter
'vrniZnakVrste = funGetCoordinNumber
'vrednost = intCoordin
'cas = Clock
'sekunde = Second
'ure = Hour
'belcas = Whites
'crncas= Blacks
'uraBeli = lblWhitesClock
'uracrni = lblBlacksClock
'ur = bytHours
'minut = bytMinutes
'sekund = bytSecunds
'formatirajCas = fntFormatTime
'resetirajUro = subResetClocks
'vrstica = strVariable
'lahkoPokaze = bolValue
'ploscax = intBoardPosX
'ploscay = intBoardPosY
'pcx = intMiniBoardPosX
'pcy = intMiniBoardPosY
'potx = intMoveListPosX
'poty = intMoveListPosY
'urax = intClockPosX
'uray = intClockPosY
'pozx = intCapturPiecesPosX
'pozy = intCapturPiecesPosY
'preberiVrstico = subProcessData
'DisplayCompleteBoard = subResetChessBoard
'displaymoves = subSetMoveListWindow
'displaycaptured = subSetCapturPiecesWindow
'miniDisplayCompleteBoard = subResetMiniChessBoard
'displayWatch = subSetClockWindow
'game_new = mnuNewGame
'pozrtih = aryCaptPiecesCount
'minipozrtih = aryMiniCaptPiecesCount
'view_pc = mnuEngineThinking
'view_moves = mnuMoveList
'view_captured = mnuCaptPieces
'view_clock = mnuChessClock
'Seznam = txtMoveList
'preracunajKoordinate = subSquareCoord
'kolicnik = strCharacter
'minipreracunajKoordinate = subMiniSquareCoord
'kam = strCoordinate
'faktor = cnsFactor
'minifaktor = cnsMiniFactor
'drawPiece = subDrawPiece
'katero = strPiece
'kje = strCoordinate
'orgX = sngOrgX
'orgY = sngOrgY
'figindex = intPieceIndex
'figurac = imgBlackPiece
'figurab = imgWhitePiece
'figc = pclBlackPictureClip
'figb = pclWhitePictureClip
'displayEmptyBoard = subEmptyChessBoard
'pol = strPosition
'deletePiece = subDeletePiece
'minideletePiece = subMiniDeletePiece
'prvapot = strOrigin
'drugapot = strDestination
'linija = intASCLetter
'vrsta = intASCNumber
'ime = strPiece
'drawcaptured = subDrawCaptured
'pozrta = strCaptured
'pozindex = intColorIndex
'kaj = strPiece
'tekst = strText
'x = intWindowPosX
'y = intWindowPosY
'lahko = bolVisible
'pol = intTemp

Dim WithEvents objEngine As chess.engine
Attribute objEngine.VB_VarHelpID = -1

Private Sub Form_Load()
    bolBlackTurn = False
    intMoveCount = 1
    aryCaptPiecesCount(0) = 0
    aryCaptPiecesCount(1) = 0
    
End Sub

Private Sub Form_Click()
    'mot.Postavi
    
End Sub

Public Sub Form_KeyPress(KeyAscii As Integer)
    Dim intKeyCode As Integer
    
    'If this Sub is called by a move made by the
    'computer and not by a KeyPress event, we
    'should skip the first phase of this Sub...
    If (KeyAscii = 0) Then
        GoTo CheckCurrMove
        
    End If
    
    'Check if it is your turn to move...
    If (bolChessEngineTurn = True) Then
        MsgBox "It's not your turn!", vbCritical, "Wait a moment..."
        Exit Sub
    End If
    
    'Check if the user pressed the Backspace key...
    If (KeyAscii = 8) Then
            
        'If the TextBox where the user is supposed
        'to type the moves is empty, just Exit the Sub...
        If (lblMove.Caption = "") Then
            Exit Sub
        End If
           
        'When you write your move and you type the
        'first coordinate, a dash "-" character is
        'automatically added to facilitate the writing
        'of the coordinates. However, when you press
        'the "BackSpace" key, this dash should be
        'automatically eliminated too. That's why
        'the following "If Then" statement is for.
        If (Len(lblMove.Caption) = 3) Then
            lblMove.Caption = Left(lblMove.Caption, 1)
            Exit Sub
            
        'In any other case, just delete the
        'leftmost character.
        Else
            lblMove.Caption = Left(lblMove.Caption, Len(lblMove.Caption) - 1)
            
        End If
        
        Exit Sub
        
    End If
    
    
    'If KeyAscii = 0 Then GoTo izvrsipotezo
    'If KeyAscii < 49 Then Exit Sub
    
    'prvi in tretji znak MORA biti èrka!
    
    'If the player is typing the first letter
    'of the origin coordinate or the first
    'letter of the destination coordinate,
    'make sure that it is in the range of
    'allowed letters. Otherwise, exit sub.
    If (Len(lblMove.Caption) = 0) _
    Or (Len(lblMove.Caption) = 3) Then
        'It is just a complicated way to
        'make sure that the typed letter
        'is an Upercase letter...
        intKeyCode = Asc(UCase(Chr(KeyAscii)))
        
        If (intKeyCode < 65) _
        Or (intKeyCode > 72) Then
            Exit Sub
        End If
        
    End If
    
    'Just making sure that the second
    'character of the origin coordinate
    'and the second character of the
    'destination coordinate is a number
    'withing the allowed range...
    If (Len(lblMove.Caption) = 1) _
    Or (Len(lblMove.Caption) = 4) Then
        
        If (KeyAscii < 49) _
        Or (KeyAscii > 56) Then
            Exit Sub
        End If
        
    End If
    
    
    'sprejetje poteze = MoveList!

'On this phase of the Sub we will get in touch
CheckCurrMove:
    Dim ChessEngine As chess.Parser
    Set ChessEngine = CreateObject("Chess.Parser")
    Dim strPosition As String
    
    lblMove.Caption = lblMove.Caption & UCase(Chr(KeyAscii))
    
    'First of all, check the first coordinate
    'to make sure that the user is actually
    'typing a coordinate where one of his
    'pieces is seating in.
    If (Len(lblMove.Caption) = 2) Then
        'Tell the Engine which color's turn
        'we are in.
        ChessEngine.BlackTurn = bolBlackTurn
        'Send the current coordinate to
        'the Chess Engine.
        ChessEngine.Move = lblMove.Caption
        'Call the Function that will create
        'a string representing the current
        'Position of the pieces on the Board.
        strPosition = fntBuildPositString
        'Send the string with the current
        'position to the Chess Engine.
        ChessEngine.Parse (strPosition)
        
        'The Chess Engine will check if
        'there actually is a piece on
        'the coordinate that the player
        'has provided. If there isn't
        'any, it will generate an error
        'message.
        If (ChessEngine.ErrorNumber > 0) Then
            strPosition = ChessEngine.ErrorText
            MsgBox "Error:" & strPosition, vbCritical, "ERROR!!!"
            lblMove.Caption = ""
            GoTo endsub
        End If
        
        'Automatically add a dash "-" between
        'the origin and destination coordinates.
        lblMove.Caption = lblMove.Caption & "-"
        
    End If
    
    'If the player has finished typing
    'the destination coordinate, proceed
    'with the analyses of the move...
    If (Len(lblMove.Caption) = 5) Then
        'poteza! je bil bel?
        'najprej jo je treba sparsati in potem prikazati...
        
        'Tell the Engine which color's turn
        'we are in.
        ChessEngine.BlackTurn = bolBlackTurn
        'Send the current coordinate to
        'the Chess Engine.
        ChessEngine.Move = lblMove.Caption
        
        'Dim linija As Integer, vrsta As Integer
        
        'Call the Function that will create
        'a string representing the current
        'Position of the pieces on the Board.
        strPosition = fntBuildPositString
        
        'Inform the Chess Engine whether it
        'is possible to castle on either
        'side with any of the Kings.
        ChessEngine.BlackCanCastleLeft = bolBlackCanCastleLeft
        ChessEngine.WhiteCanCastleLeft = bolWhiteCanCastleLeft
        ChessEngine.BlackCanCastleRight = bolBlackCanCastleRight
        ChessEngine.WhiteCanCastleRight = bolWhiteCanCastleRight
        
        'Send the string with the current
        'position to the Chess Engine.
        ChessEngine.Parse (strPosition)
        
        'The Chess Engine will check if
        'the move made by the player is
        'legal. If it isn't, an error
        'message will be generated.
        If (ChessEngine.ErrorNumber > 0) Then
            strPosition = ChessEngine.ErrorText
            MsgBox "ERROR:" & strPosition, vbCritical, "ERROR!!!"
            lblMove.Caption = ""
            GoTo endsub
        End If
        
        'opravimo premik
        
        'If this is the first move, start
        'the ChessClock...
        If (intMoveCount < 2) Then
            frmChessClock.subStartChessClock
            
        End If
        
        
        'Retrieve the status of the Kings in
        'relation to their Castling rights.
        bolBlackCanCastleRight = ChessEngine.BlackCanCastleRight
        bolWhiteCanCastleRight = ChessEngine.WhiteCanCastleRight
        bolBlackCanCastleLeft = ChessEngine.BlackCanCastleLeft
        bolWhiteCanCastleLeft = ChessEngine.WhiteCanCastleLeft
        
        'Make the actual move on the Board.
        Call subMovePiece(lblMove.Caption)
        DoEvents
        
        'Make the same move on the little
        'board that shows what the Chess
        'Engine is thinking.
        Call subMiniMovePiece(lblMove.Caption)
        DoEvents
        
        'Check with the Chess Engine if the move
        'made by the player is a White Castling.
        'If so, complete the move by setting the
        'Rook into its proper square.
        If (ChessEngine.WhiteCastled = True) Then
        
            'kam je bila rošada(katero trdnjavo premaknemo)
            
            'Determine whether it was a Castling
            'on the Queen side or on the King side...
            If (Right(lblMove.Caption, 2) = "G1") Then
                subMovePiece "H1-F1"
                subMiniMovePiece "H1-F1"
            Else
                subMovePiece "A1-D1"
                subMiniMovePiece "A1-D1"
            End If
        End If
        
        'Check with the Chess Engine if the move
        'made by the player is a Black Castling.
        'If so, complete the move by setting the
        'Rook into its proper square.
        If ChessEngine.BlackCastled = True Then
        
            'kam je bila rošada(katero trdnjavo premaknemo)
            
            'Determine whether it was a Castling
            'on the Queen side or on the King side...
            If Right(lblMove.Caption, 2) = "G8" Then
                subMovePiece "H8-F8"
                subMiniMovePiece "H8-F8"
            Else
                subMovePiece "A8-D8"
                subMiniMovePiece "A8-D8"
            End If
        End If
        
        'Check with the Chess Engine if the move
        'made by the player ended up with the
        'Capture of a piece. If so, automatically
        'add a "X" to the end of the move
        'description at the MoveList window.
        If (ChessEngine.capture = True) Then
            strMoveListStatus = "X"
            
        End If
        
        'Check with the Chess Engine if the move
        'made by the player ended up with a
        'Check to the King. If so, automatically
        'add a Plus Sing "+" to the end of the
        'move description at the MoveList window.
        If (ChessEngine.Check = True) Then
            strMoveListStatus = strMoveListStatus & "+"
        End If
        
        'Send the move to the MoveList window
        'accordingly...
        If (bolBlackTurn = False) Then
            bolBlackTurn = True
            
            frmMoveList.txtMoveList = frmMoveList.txtMoveList & intMoveCount & ". " & lblMove.Caption & strMoveListStatus
            lblMove.Caption = ""
            lblWhoIsMoving.Caption = "Black moves"
            GoTo endsub
            
        Else
            bolBlackTurn = False
            frmMoveList.txtMoveList = frmMoveList.txtMoveList & " ... " & lblMove.Caption & strMoveListStatus & vbCrLf
            lblMove.Caption = ""
            lblWhoIsMoving.Caption = "White moves"
            intMoveCount = intMoveCount + 1
            
        End If
    End If
                
endsub:
    'Release computer resources...
    Set ChessEngine = Nothing
    
    'If it is Black's turn, start
    'Chess Engine...
    If (bolBlackTurn = True) Then
        bolChessEngineTurn = True
        Call subStartChessEngine
        
    End If
    
End Sub

'Time to make the Chess Engine think!!
Public Sub subStartChessEngine()
    Set objEngine = CreateObject("Chess.engine")
    
    strMoveListStatus = ""
    
    'Send the Catling status of each King...
    objEngine.BlackCanCastleLeft = bolBlackCanCastleLeft
    objEngine.BlackCanCastleRight = bolBlackCanCastleRight
    objEngine.WhiteCanCastleLeft = bolWhiteCanCastleLeft
    objEngine.WhiteCanCastleRight = bolWhiteCanCastleRight
    
    'Get the Game Level and the Playing
    'Style form the Settings window.
    objEngine.Levels = frmSettings.srbGameLevel.Value
    objEngine.StyleOfPlay = frmSettings.srbPlayingStyle.Value
    
    'Send the current position and the
    'Move Count to the Engine.
    objEngine.think fntBuildPositString(), intMoveCount
    
End Sub

Private Sub objEngine_FoundMove(ByVal strMove As String, ByVal statuspo As String)
    
    If (strMove = "00-00") Then
        frmChessClock.subStopClock
        Exit Sub  'konec igre!
        
    End If
    
    bolChessEngineTurn = False
    
    'Paste move made by the Chess
    'Engine to the Label.
    lblMove.Caption = UCase(strMove)
    
    'Have no idea!!
    strMoveListStatus = statuspo
    
    'I was trying to understand what
    'the statuspo was... No luck so far...
    If statuspo <> "" Then
        statuspo = statuspo
    End If
    
    'Call the Form_KeyPress Sub to
    'force it to check the lblMove
    'again. The lblMove, now, has
    'the Chess Engine move.
    Call Form_KeyPress(0)
    
End Sub

Private Sub objEngine_DrawMove(ByVal strMove As String)
    'Dim sfig As String
    'Dim cfig As String
    
    Call subMiniMovePiece(strMove)
    
End Sub

Private Sub objEngine_drawcastle(ByVal strMove As String)
    
    If (strMove = "E8-G8") Then
        subMiniMovePiece (strMove)
        subMiniMovePiece ("H8-F8")
        
    ElseIf (strMove = "E1-G1") Then
        subMiniMovePiece (strMove)
        subMiniMovePiece ("H1-F1")
        
    ElseIf (strMove = "E1-C1") Then
        subMiniMovePiece (strMove)
        subMiniMovePiece ("A1-D1")
        
    ElseIf (strMove = "E8-C8") Then
        subMiniMovePiece (strMove)
        subMiniMovePiece ("A8-D8")
        
    End If
    
    'MsgBox "moja ideja!"
    
    'If (strMove = "E8-G8") Then
    '    subMiniMovePiece ("G8-E8")
    '    subMiniMovePiece ("F8-H8")
    '
    'End If
    
    'If (strMove = "E1-G1") Then
    '    subMiniMovePiece ("G1-E1")
    '    subMiniMovePiece ("F1-H1")
    '
    'End If
    
    'If (strMove = "E1-C1") Then
    '    subMiniMovePiece ("C1-E1")
    '    subMiniMovePiece ("D1-A1")
    '
    'End If
    
    'If (strMove = "E8-C8") Then
    '    subMiniMovePiece ("C8-E8")
    '    subMiniMovePiece ("D8-A8")
    '
    'End If

End Sub

'Draw the original board on the little
'chess board that displays the Chess
'Engine thinking.
Private Sub objEngine_RestoreBoard(ByVal strPosition As String)
    'potrebno restavrirati mini plošèo na poteze = Move List
    Dim n As Integer, m As Integer
    
    'First, we have to cleanup the board...
    Call subMiniEmptyBoard   'ozris prazne plošèe
    
    'napolnimo jo!
    
    'Now, it's time to redraw the position
    'sent by the ChessEngine to the little
    'chess board.
    For n = 1 To 8
        For m = 1 To 8
            If (strPosition = "") Then
                Call subMiniDrawPiece(aryGamePosition(n, m), funGetCoordinLetter(n) & funGetCoordinNumber(m))
                aryMiniGamePosition(n, m) = aryGamePosition(n, m)
                
            Else
                Call subMiniDrawPiece(Mid(strPosition, 4, 2), Left(strPosition, 2))
                aryMiniGamePosition(m, n) = Mid(strPosition, 4, 2)
                strPosition = Mid(strPosition, 7)
                
            End If
        Next
    Next
    
    Call frmChessBoard.SetFocus
    
End Sub
