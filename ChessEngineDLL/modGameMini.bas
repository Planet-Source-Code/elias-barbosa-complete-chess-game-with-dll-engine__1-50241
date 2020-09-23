Attribute VB_Name = "modGameMini"
Option Explicit

Const cnsMiniFactor = 335
Global aryMiniGamePosition(9, 9) As String * 2
Global bolMiniBlackTurn As Boolean
Global aryMiniCaptPiecesCount(2) As Integer
Global strMiniMoveListStatus As String
Global bolMiniWhiteCanCastleLeft As Boolean
Global bolMiniBlackCanCastleLeft As Boolean
Global bolMiniWhiteCanCastleRight As Boolean
Global bolMiniBlackCanCastleRight As Boolean

'When this sub is called, a coordinate
'is given by the calling sub. For
'example, when the sub is called
'and the coordinate given is "A2",
'the sub will determine the exact
'coordinates of the correspondent
'square on the bitmap picture of
'the chess board.
Private Sub subMiniSquareCoord(ByRef x As Single, ByRef y As Single, ByVal strCoordinate As String)
    Dim strCharacter As Single
    
    'najprej èrka
    
    'This sub will update the x & y "ByRef"
    'variables provided by the calling sub.
    strCharacter = Asc(UCase(Left(strCoordinate, 1))) - 65
    x = x + (strCharacter * cnsMiniFactor)
    
    'potem številka
    'strCharacter = Asc(UCase(Right(strCoordinate, 1))) - 48
    
    strCharacter = Asc(Right(strCoordinate, 1)) - 48
    y = y - (strCharacter * cnsMiniFactor) 'dvigamo se
    
End Sub

'When you provide the piece that
'you want to draw and the coordinate
'on the board where you want the
'piece to be drawn to, this sub
'will actually do the work for you!!
Sub subMiniDrawPiece(ByVal strPiece As String, ByVal strCoordinate As String)
    Dim sngOrgX As Single
    Dim sngOrgY As Single
    Dim intPieceIndex As Integer
    
    sngOrgX = 50
    sngOrgY = frmMiniChessBoard.picChessBoard.Height - 80 'dno šahovnice
    Call subMiniSquareCoord(sngOrgX, sngOrgY, strCoordinate)
    
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
        
        Case Else
            Exit Sub
            
    End Select
    
    'If the piece is Black...
    If (UCase(Left(strPiece, 1)) = "C") Then
        intPieceIndex = intPieceIndex + 6
        
    End If
    
    'Temporarily copy the piece image
    'form the PictureClip into an
    'Image control...
    frmMiniChessBoard.imgBlackPiece.Picture = frmMiniChessBoard.pclBlackPictureClip.GraphicCell(intPieceIndex)
    frmMiniChessBoard.imgWhitePiece.Picture = frmMiniChessBoard.pclWhitePictureClip.GraphicCell(intPieceIndex)
            
    
    'pa jo narišimo
    
    'Paint the piece image that was
    'temporarily stored into an Image
    'control at the right positon onto
    'the chess board that was painted
    'on the frmChessBoard...
    frmMiniChessBoard.PaintPicture frmMiniChessBoard.imgWhitePiece.Picture, sngOrgX, sngOrgY, , , , , , , vbSrcAnd
    frmMiniChessBoard.PaintPicture frmMiniChessBoard.imgBlackPiece.Picture, sngOrgX, sngOrgY, , , , , , , vbSrcPaint
    DoEvents

End Sub

'This sub is not currently being
'used by any other procedure...
Public Sub subMiniEmptyBoard()
    Dim n As Integer, m As Integer
    
    'Reset the Chess Board on
    'the frmMiniChessBoard form...
    frmMiniChessBoard.Width = frmMiniChessBoard.picChessBoard.Width
    frmMiniChessBoard.Height = frmMiniChessBoard.picChessBoard.Height + 320
    frmMiniChessBoard.PaintPicture frmMiniChessBoard.picChessBoard.Picture, 0, 0
    frmMiniChessBoard.Left = intMiniBoardPosX
    frmMiniChessBoard.Top = intMiniBoardPosY
    
    
    'izpraznemo šahovnico
    
    'Reset the Castling parameters...
    bolMiniBlackCanCastleRight = True
    bolMiniWhiteCanCastleRight = True
    bolMiniBlackCanCastleLeft = True
    bolMiniWhiteCanCastleLeft = True
    
    'Empty the aryGamePosition and
    'the aryMiniGamePosition arrays...
    For n = 1 To 8
        For m = 1 To 8
            aryMiniGamePosition(n, m) = "  "
        Next
    Next
    
    'Show the Mini Chess Board window...
    frmMiniChessBoard.Show
End Sub

Sub subResetMiniChessBoard()
    Dim n As Integer
    Dim m As Integer 'izpraznemo vsebino
    
    'Reset the Chess Board on
    'the frmMiniChessBoard form...
    frmMiniChessBoard.Width = frmMiniChessBoard.picChessBoard.Width
    frmMiniChessBoard.Height = frmMiniChessBoard.picChessBoard.Height + 320
    frmMiniChessBoard.PaintPicture frmMiniChessBoard.picChessBoard.Picture, 0, 0
    frmMiniChessBoard.Left = intMiniBoardPosX
    frmMiniChessBoard.Top = intMiniBoardPosY
    
    'Reset the Castling parameters...
    bolMiniBlackCanCastleRight = True
    bolMiniWhiteCanCastleRight = True
    bolMiniBlackCanCastleLeft = True
    bolMiniWhiteCanCastleLeft = True
    
    'Empty the aryMiniGamePosition arrays...
    For n = 1 To 8
        For m = 1 To 8
            aryMiniGamePosition(n, m) = "  "
        Next
    Next
    
    'beli kmetje
    
    'Draw the white pawns
    'on the chess board...
    subMiniDrawPiece "BP", "A2"
    subMiniDrawPiece "BP", "B2"
    subMiniDrawPiece "BP", "c2"
    subMiniDrawPiece "BP", "d2"
    subMiniDrawPiece "BP", "e2"
    subMiniDrawPiece "BP", "f2"
    subMiniDrawPiece "BP", "g2"
    subMiniDrawPiece "BP", "h2"
    
    'še zafilamo tabelo
    
    'Update the arrays with the
    'pawns locations...
    For n = 1 To 8
        aryMiniGamePosition(n, 2) = "BP"
    Next
    
    'èrni kmetje
    
    'Draw the black pawns
    'on the chess board...
    subMiniDrawPiece "cP", "A7"
    subMiniDrawPiece "cP", "B7"
    subMiniDrawPiece "cP", "c7"
    subMiniDrawPiece "cP", "d7"
    subMiniDrawPiece "cP", "e7"
    subMiniDrawPiece "cP", "f7"
    subMiniDrawPiece "cP", "g7"
    subMiniDrawPiece "cP", "h7"
    
    'Update the arrays with the
    'pawns locations...
    For n = 1 To 8
        aryMiniGamePosition(n, 7) = "CP"
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
    subMiniDrawPiece "bt", "A1"
    subMiniDrawPiece "bs", "B1"
    subMiniDrawPiece "bl", "C1"
    subMiniDrawPiece "bq", "D1"
    subMiniDrawPiece "bk", "E1"
    subMiniDrawPiece "bl", "F1"
    subMiniDrawPiece "bs", "G1"
    subMiniDrawPiece "bt", "H1"
    
    'Draw the black pieces
    'on the chess board...
    subMiniDrawPiece "ct", "A8"
    subMiniDrawPiece "cs", "B8"
    subMiniDrawPiece "cl", "C8"
    subMiniDrawPiece "cq", "D8"
    subMiniDrawPiece "ck", "E8"
    subMiniDrawPiece "cl", "F8"
    subMiniDrawPiece "cs", "G8"
    subMiniDrawPiece "ct", "H8"
    
    'Update the array with the
    'pieces locations...
    aryMiniGamePosition(1, 1) = "BT"
    aryMiniGamePosition(2, 1) = "BS"
    aryMiniGamePosition(3, 1) = "BL"
    aryMiniGamePosition(4, 1) = "BQ"
    aryMiniGamePosition(5, 1) = "BK"
    aryMiniGamePosition(6, 1) = "BL"
    aryMiniGamePosition(7, 1) = "BS"
    aryMiniGamePosition(8, 1) = "BT"
    
    aryMiniGamePosition(1, 8) = "CT"
    aryMiniGamePosition(2, 8) = "CS"
    aryMiniGamePosition(3, 8) = "CL"
    aryMiniGamePosition(4, 8) = "CQ"
    aryMiniGamePosition(5, 8) = "CK"
    aryMiniGamePosition(6, 8) = "CL"
    aryMiniGamePosition(7, 8) = "CS"
    aryMiniGamePosition(8, 8) = "CT"
    
    frmMiniChessBoard.Show
    
End Sub

'THIS FUNCTION IS NOT BEEN CALLED BY ANY OTHER PROCEDURE...
'This function will create a
'string that will represent the
'current game position.
Public Function fntMiniBuildPositString() As String
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
    fntMiniBuildPositString = strPosition
    
End Function

'This sub will clear a square.
'It is called in each move...
Sub subMiniDeletePiece(ByVal strCoordinate As String)
    'izbrišemo zahtevan položaj
    Dim sngOrgX As Single
    Dim sngOrgY As Single
    
    sngOrgX = 50
    sngOrgY = frmMiniChessBoard.picChessBoard.Height - 80 'dno šahovnice
    
    'This sub will calculate the correct
    'sngOrgX and sngOrgY and, then, will
    'alter these variables that are set as
    '"By Reference".
    'This is not the most elegant way to
    'return information to a calling procedure.
    'I didn't want to alter it at the moment, though...
    subMiniSquareCoord sngOrgX, sngOrgY, strCoordinate
    
    'With all the information gathered,
    'it's time to do the actual painting...
    frmMiniChessBoard.PaintPicture frmMiniChessBoard.picChessBoard.Picture, sngOrgX, sngOrgY, , , sngOrgX, sngOrgY, frmMiniChessBoard.imgWhitePiece.Width, frmMiniChessBoard.imgWhitePiece.Height
    DoEvents
    
End Sub

Public Sub subMiniMovePiece(ByVal strMove As String)
    Dim strPiece As String, strOrigin As String, strDestination As String
    Dim intASCLetter As Integer, intASCNumber As Integer
    
    'MsgBox "izvajam minipremik " & strMove
    'najprej izbrišemo figuro
    
    'Separate the origin and destination
    'coordinates from the move string that
    'was provided by the calling procedure...
    strOrigin = Left(strMove, 2)
    strDestination = Right(strMove, 2)
    
    'Cleanup the origin and destination squares...
    Call subMiniDeletePiece(strOrigin)  'izpraznemo štart
    Call subMiniDeletePiece(strDestination)  'in cilj
    
    'vzamemo ime figure
    
    'Find out the ASCII code of the letter and
    'the number that make up the origin coordinate...
    intASCLetter = Asc(UCase(Left(strOrigin, 1))) - 64
    intASCNumber = Asc(UCase(Right(strOrigin, 1))) - 48
    
    'Find out what piece is moving...
    strPiece = aryMiniGamePosition(intASCLetter, intASCNumber)
    
    'Inform the aryGamePosition array
    'that there is no piece there anymore...
    aryMiniGamePosition(intASCLetter, intASCNumber) = "  " 'izpraznemo polozaj
    
    'Find out the ASCII code of the letter and
    'the number that make up the destination coordinate...
    intASCLetter = Asc(UCase(Left(strDestination, 1))) - 64
    intASCNumber = Asc(UCase(Right(strDestination, 1))) - 48
    
    'Update the aryGamePosition array with
    'the new position of the piece...
    aryMiniGamePosition(intASCLetter, intASCNumber) = strPiece
    
    'Finally, draw the piece to its
    'destination place...
    Call subMiniDrawPiece(strPiece, strDestination)

End Sub
