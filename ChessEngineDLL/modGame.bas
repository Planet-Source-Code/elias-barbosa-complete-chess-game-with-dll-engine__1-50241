Attribute VB_Name = "modGame"
Option Explicit
Const cnsFactor = 740
Global aryGamePosition(9, 9) As String * 2
Global bolBlackTurn As Boolean
Global aryCaptPiecesCount(2) As Integer
Global strMoveListStatus As String
Global bolWhiteCanCastleLeft As Boolean
Global bolBlackCanCastleLeft As Boolean
Global bolWhiteCanCastleRight As Boolean
Global bolBlackCanCastleRight As Boolean

'When this sub is called, a coordinate
'is given by the calling sub. For
'example, when the sub is called
'and the coordinate given is "A2",
'the sub will determine the exact
'coordinates of the correspondent
'square on the bitmap picture of
'the chess board.
Private Sub subSquareCoord(ByRef x As Single, ByRef y As Single, ByVal strCoordinate As String)
    Dim strCharacter As Single
    
    'najprej èrka
    
    'This sub will update the x & y "ByRef"
    'variables provided by the calling sub.
    strCharacter = Asc(UCase(Left(strCoordinate, 1))) - 65
    x = x + (strCharacter * cnsFactor)
    
    'potem številka
    'strCharacter = Asc(UCase(Right(strCoordinate, 1))) - 48
    
    strCharacter = Asc(Right(strCoordinate, 1)) - 48
    y = y - (strCharacter * cnsFactor) 'dvigamo se
    
End Sub

'When you provide the piece that
'you want to draw and the coordinate
'on the board where you want the
'piece to be drawn to, this sub
'will actually do the work for you!!
Private Sub subDrawPiece(ByVal strPiece As String, ByVal strCoordinate As String)
    Dim sngOrgX As Single
    Dim sngOrgY As Single
    Dim intPieceIndex As Integer
    
    sngOrgX = 300
    sngOrgY = frmChessBoard.picChessBoard.Height - 200 'dno šahovnice
    Call subSquareCoord(sngOrgX, sngOrgY, strCoordinate)
    
    Select Case UCase(Right(strPiece, 1))
        'If the piece is a Pawn...
        Case "P"
            'kmet
            intPieceIndex = 0
            
        'If the piece is a Rook...
        Case "T"
            'trdnjava
            intPieceIndex = 1
            
        'If the piece is a Knight...
        Case "S"
            'skakaè
            intPieceIndex = 2
            
        'If the piece is a Bishop...
        Case "L"
            'lovec
            intPieceIndex = 3
            
        'If the piece is a Queen...
        Case "Q"
            'kraljica
            intPieceIndex = 4
            
        'If the piece is a King...
        Case "K"
            'kralj
            intPieceIndex = 5
            
    End Select
    
    'If the piece is Black...
    If (UCase(Left(strPiece, 1)) = "C") Then
        intPieceIndex = intPieceIndex + 6
        
    End If
    
    'Temporarily copy the piece image
    'form the PictureClip into an
    'Image control...
    frmChessBoard.imgBlackPiece.Picture = frmChessBoard.pclBlackPictureClip.GraphicCell(intPieceIndex)
    frmChessBoard.imgWhitePiece.Picture = frmChessBoard.pclWhitePictureClip.GraphicCell(intPieceIndex)
    
    'pa jo narišimo
    
    'Paint the piece image that was
    'temporarily stored into an Image
    'control at the right positon onto
    'the chess board that was painted
    'on the frmChessBoard...
    frmChessBoard.PaintPicture frmChessBoard.imgWhitePiece.Picture, sngOrgX, sngOrgY, , , , , , , vbSrcAnd
    frmChessBoard.PaintPicture frmChessBoard.imgBlackPiece.Picture, sngOrgX, sngOrgY, , , , , , , vbSrcPaint
    DoEvents
    
End Sub

'This sub is not currently being
'used by any other procedure...
Private Sub subEmptyChessBoard()
    Dim n As Integer, m As Integer
    
    'Reset the Chess Board on
    'the frmChessBoard form...
    frmChessBoard.Width = frmChessBoard.picChessBoard.Width
    frmChessBoard.Height = frmChessBoard.picChessBoard.Height + 1000
    frmChessBoard.PaintPicture frmChessBoard.picChessBoard.Picture, 0, 0
    frmChessBoard.Left = intBoardPosX
    frmChessBoard.Top = intBoardPosY
    
    'izpraznemo šahovnico
    
    'Reset the Castling parameters...
    bolBlackCanCastleRight = True
    bolWhiteCanCastleRight = True
    bolBlackCanCastleLeft = True
    bolWhiteCanCastleLeft = True
    
    'Empty the aryGamePosition and
    'the aryMiniGamePosition arrays...
    For n = 1 To 8
        For m = 1 To 8
            aryGamePosition(n, m) = "  "
            aryMiniGamePosition(n, m) = "  "
        Next
    Next
    
    'Show the Chess Board window...
    frmChessBoard.Show
    DoEvents
    
End Sub

Public Sub subResetChessBoard()
    Dim n As Integer
    Dim m As Integer 'izpraznemo vsebino
    
    'Reset the Chess Board on
    'the frmChessBoard form...
    frmChessBoard.Width = frmChessBoard.picChessBoard.Width
    frmChessBoard.Height = frmChessBoard.picChessBoard.Height + 1000
    frmChessBoard.PaintPicture frmChessBoard.picChessBoard.Picture, 0, 0
    frmChessBoard.Left = intBoardPosX
    frmChessBoard.Top = intBoardPosY
    
    'Reset the Castling parameters...
    bolBlackCanCastleRight = True
    bolWhiteCanCastleRight = True
    bolBlackCanCastleLeft = True
    bolWhiteCanCastleLeft = True
    
    'Empty the aryGamePosition and
    'the aryMiniGamePosition arrays...
    For n = 1 To 8
        For m = 1 To 8
            aryGamePosition(n, m) = "  "
            aryMiniGamePosition(n, m) = "  "
        Next
    Next
    
    'beli kmetje
    
    'Draw the white pawns
    'on the chess board...
    subDrawPiece "BP", "A2"
    subDrawPiece "BP", "B2"
    subDrawPiece "BP", "c2"
    subDrawPiece "BP", "d2"
    subDrawPiece "BP", "e2"
    subDrawPiece "BP", "f2"
    subDrawPiece "BP", "g2"
    subDrawPiece "BP", "h2"
    
    'še zafilamo tabelo
    
    'Update the arrays with the
    'pawns locations...
    For n = 1 To 8
        aryGamePosition(n, 2) = "BP"
        'aryMiniGamePosition(n, 2) = "BP"
    Next
    
    'èrni kmetje
    
    'Draw the black pawns
    'on the chess board...
    subDrawPiece "cP", "A7"
    subDrawPiece "cP", "B7"
    subDrawPiece "cP", "c7"
    subDrawPiece "cP", "d7"
    subDrawPiece "cP", "e7"
    subDrawPiece "cP", "f7"
    subDrawPiece "cP", "g7"
    subDrawPiece "cP", "h7"
    
    'Update the arrays with the
    'pawns locations...
    For n = 1 To 8
        aryGamePosition(n, 7) = "CP"
        'aryMiniGamePosition(n, 7) = "CP"
    Next
    
    'beli trdnjavi
    'èrni trdnjavi
    'bela konja
    'èrna konja
    'bela lovca
    'crna lovca
    'bela kraljica
    'crna kraljica
    'beli kralj
    'èrni kralj
    
    'Draw the white pieces
    'on the chess board...
    subDrawPiece "bt", "A1"
    subDrawPiece "bs", "B1"
    subDrawPiece "bl", "C1"
    subDrawPiece "bq", "D1"
    subDrawPiece "bk", "E1"
    subDrawPiece "bl", "F1"
    subDrawPiece "bs", "G1"
    subDrawPiece "bt", "H1"
    
    'Draw the black pieces
    'on the chess board...
    subDrawPiece "ct", "A8"
    subDrawPiece "cs", "B8"
    subDrawPiece "cl", "C8"
    subDrawPiece "cq", "D8"
    subDrawPiece "ck", "E8"
    subDrawPiece "cl", "F8"
    subDrawPiece "cs", "G8"
    subDrawPiece "ct", "H8"
    
    'Update the array with the
    'pieces locations...
    aryGamePosition(1, 1) = "BT"
    aryGamePosition(2, 1) = "BS"
    aryGamePosition(3, 1) = "BL"
    aryGamePosition(4, 1) = "BQ"
    aryGamePosition(5, 1) = "BK"
    aryGamePosition(6, 1) = "BL"
    aryGamePosition(7, 1) = "BS"
    aryGamePosition(8, 1) = "BT"
    
    aryGamePosition(1, 8) = "CT"
    aryGamePosition(2, 8) = "CS"
    aryGamePosition(3, 8) = "CL"
    aryGamePosition(4, 8) = "CQ"
    aryGamePosition(5, 8) = "CK"
    aryGamePosition(6, 8) = "CL"
    aryGamePosition(7, 8) = "CS"
    aryGamePosition(8, 8) = "CT"
    
    'aryMiniGamePosition(1, 1) = "BT"
    
    frmChessBoard.Show
    DoEvents
    
End Sub

'Set chess clock window to
'its proper position...
Public Sub subSetClockWindow()
    frmChessClock.Top = intClockPosY
    frmChessClock.Left = intClockPosX
    frmChessClock.Show
    
End Sub

'Set move list window to
'its proper position...
Public Sub subSetMoveListWindow()
    frmMoveList.Left = intMoveListPosX
    frmMoveList.Top = intMoveListPosY
    frmMoveList.Show
    
End Sub

'This function will create a
'string that will represent the
'current game position.
Public Function fntBuildPositString() As String
    'tu moramo iz polozajev sestaviti string:
    Dim x As Integer, y As Integer
    Dim strPosition As String
    
    strPosition = ""
    
    For y = 1 To 8
        For x = 1 To 8
            'If the coordinate of the array
            'aryGamePosition is empty put
            'two spaces in it...
            If (Trim(Left(aryGamePosition(x, y), 1)) = Chr(0)) Then
                aryGamePosition(x, y) = "  "
            End If
            
            'For each square of the board put
            'the number of the square, add a
            'column ":" and then the initial
            'of the piece. To complete the
            'description of the current square,
            'add a strait bar to the end "|"...
            strPosition = strPosition & Chr(x + 64) & Trim(Str(y)) & ":" & aryGamePosition(x, y) & "|"
        Next
    Next
    
    'Return the string to the
    'calling procedure...
    fntBuildPositString = strPosition
    
End Function

'This sub will clear a square.
'It is called in each move...
Private Sub subDeletePiece(ByVal strCoordinate As String)
    'izbrišemo zahtevan položaj
    Dim sngOrgX As Single
    Dim sngOrgY As Single
    
    sngOrgX = 300
    sngOrgY = frmChessBoard.picChessBoard.Height - 200 'dno šahovnice
    
    'This sub will calculate the correct
    'sngOrgX and sngOrgY and, then, will
    'alter these variables that are set as
    '"By Reference".
    'This is not the most elegant way to
    'return information to a calling procedure.
    'I didn't want to alter it at the moment, though...
    Call subSquareCoord(sngOrgX, sngOrgY, strCoordinate)
    
    
    'With all the information gathered,
    'it's time to do the actual painting...
    frmChessBoard.PaintPicture frmChessBoard.picChessBoard.Picture, sngOrgX, sngOrgY, , , sngOrgX, sngOrgY, frmChessBoard.imgWhitePiece.Width, frmChessBoard.imgWhitePiece.Height
    DoEvents
    
End Sub


Public Sub subMovePiece(ByVal strMove As String)
    Dim strPiece As String, strOrigin As String, strDestination As String
    Dim intASCLetter As Integer, intASCNumber As Integer
    
    strMoveListStatus = ""
    
    'najprej izbrišemo figuro
    
    'Separate the origin and destination
    'coordinates from the move string that
    'was provided by the calling procedure...
    strOrigin = Left(strMove, 2)
    strDestination = Right(strMove, 2)
    
    'Cleanup the origin and destination squares...
    Call subDeletePiece(strOrigin) 'izpraznemo štart
    Call subDeletePiece(strDestination) 'in cilj
    
    'vzamemo ime figure
    
    'Find out the ASCII code of the letter and
    'the number that make up the origin coordinate...
    intASCLetter = Asc(UCase(Left(strOrigin, 1))) - 64
    intASCNumber = Asc(UCase(Right(strOrigin, 1))) - 48
    
    'Find out what piece is moving...
    strPiece = aryGamePosition(intASCLetter, intASCNumber)
    
    'Inform the aryGamePosition array
    'that there is no piece there anymore...
    aryGamePosition(intASCLetter, intASCNumber) = "  " 'izpraznemo polozaj
    
    
    'Find out the ASCII code of the letter and
    'the number that make up the destination coordinate...
    intASCLetter = Asc(UCase(Left(strDestination, 1))) - 64
    intASCNumber = Asc(UCase(Right(strDestination, 1))) - 48
    
    'pogledati je treba, ali se na destinaciji nahaja kaka figura!
    'èe se , potem jo je treba narisati v 'POŽRTE' oknu
    
    'Find out if this move involves
    'a piece capture...
    If (Trim(aryGamePosition(intASCLetter, intASCNumber)) <> "") Then
        Dim strCaptured As String
        
        'katero figuro bomo pojedli?
        strCaptured = aryGamePosition(intASCLetter, intASCNumber)
        
        'Draw the captured piece to the
        'frmCapturedPieces form...
        Call subDrawCaptured(strCaptured)
        
        'This is just to track the
        'number of pieces of each color...
        If (bolBlackTurn = False) Then
            aryCaptPiecesCount(0) = aryCaptPiecesCount(0) + 1
        Else
            aryCaptPiecesCount(1) = aryCaptPiecesCount(1) + 1
        End If
            
    End If
    
    'Update the aryGamePosition array with
    'the new position of the piece...
    aryGamePosition(intASCLetter, intASCNumber) = strPiece
    
    'Finally, draw the piece to its
    'destination place...
    Call subDrawPiece(strPiece, strDestination)
    
End Sub

'Set captured pieces window to
'its proper position...
Public Sub subSetCapturPiecesWindow()
    frmCapturedPieces.Left = intCapturPiecesPosX
    frmCapturedPieces.Top = intCapturPiecesPosY
    frmCapturedPieces.Show
    
End Sub

'This sub will draw the captured
'piece to the frmCapturedPieces form...
Private Sub subDrawCaptured(ByVal strPiece As String)
    Dim sngOrgX As Single
    Dim sngOrgY As Single
    Dim intPieceIndex As Integer, intColorIndex As Integer
    
    'Find out the color index for the array...
    If (bolBlackTurn = False) Then
        intColorIndex = 0
        
    Else
        intColorIndex = 1
        
    End If
    
    'This "Original X" will hold the multiplication
    'between the number of pieces already taken that
    'have the same color and the width of the pieces...
    sngOrgX = aryCaptPiecesCount(intColorIndex) * frmChessBoard.imgWhitePiece.Width
    
    'v bistvu je treba prezrcaliti barvi, kajti figure gredo nasprotniku!
    
    'The "Original Y" will be different
    'depending on the color of the piece...
    If (bolBlackTurn = False) Then
        sngOrgY = frmChessBoard.imgWhitePiece.Height * 2 'èe je figuro vzel beli, jo je treba
        
    Else        'narisati èrnemu
        sngOrgY = 300
        
    End If
    
    
    Select Case UCase(Right(strPiece, 1))
        'If it is a Pawn...
        Case "P"
            'kmet
            intPieceIndex = 0
            
        'If it is a Rook...
        Case "T"
            'trdnjava
            intPieceIndex = 1
                    
        'If it is a Knight...
        Case "S"
            'skakaè
            intPieceIndex = 2
               
        'If it is a Bishop...
        Case "L"
            'lovec
            intPieceIndex = 3
            
        'If it is a Queen...
        Case "Q"
            'kraljica
            intPieceIndex = 4
            
        'If it is a King...
        Case "K"
            'kralj
            intPieceIndex = 5
            
    End Select
    
    'If the piece is black, add 6
    'to the intPieceIndex...
    If (UCase(Left(strPiece, 1)) = "C") Then
        intPieceIndex = intPieceIndex + 6
    End If
    
    'Capture the image of the piece
    'that has been taken into a
    'temporary Image object...
    frmChessBoard.imgBlackPiece.Picture = frmChessBoard.pclBlackPictureClip.GraphicCell(intPieceIndex)
    frmChessBoard.imgWhitePiece.Picture = frmChessBoard.pclWhitePictureClip.GraphicCell(intPieceIndex)
    
    'pa jo narišimo
    'Paste the image stored in the
    'Image object onto the
    'frmCapturedPieces form...
    frmCapturedPieces.PaintPicture frmChessBoard.imgWhitePiece.Picture, sngOrgX, sngOrgY, , , , , , , vbSrcAnd
    frmCapturedPieces.PaintPicture frmChessBoard.imgBlackPiece.Picture, sngOrgX, sngOrgY, , , , , , , vbSrcPaint
    
End Sub
