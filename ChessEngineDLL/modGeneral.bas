Attribute VB_Name = "modGeneral"
Option Explicit

Global intCapturPiecesPosX As Integer, intCapturPiecesPosY As Integer
Global intMiniBoardPosX As Integer, intMiniBoardPosY As Integer
Global intClockPosX As Integer, intClockPosY As Integer
Global intMoveListPosX As Integer, intMoveListPosY As Integer
Global intBoardPosX As Integer, intBoardPosY As Integer
Global intMoveCount As Integer
Global bolChessEngineTurn As Boolean

Type Clock
    Second As Byte
    Minute As Byte
    Hour As Byte
End Type

'Global tipka As Boolean

'This Function is not been used by any procedure...
Function vrniVrednostKoordinate(ByVal strCoordinate As String) As Integer
    'se gre za številko?
    If Asc(strCoordinate) < 57 Then
        'ja!
        vrniVrednostKoordinate = Asc(strCoordinate) - 48
    Else
        'ne, je èrka
        vrniVrednostKoordinate = Asc(strCoordinate) - 64
    End If
    
End Function

'If you provide a number from 1 to 8,
'this function will return a letter from A to H...
Function funGetCoordinLetter(ByVal intCoordin As Integer) As String
    funGetCoordinLetter = Chr(intCoordin + 64)
    
End Function

'It might sound a little dumb but if
'you provide a number from 1 to 8,
'this function will return a number
'from 1 to 8...
Function funGetCoordinNumber(ByVal intCoordin As Integer) As String
    funGetCoordinNumber = Chr(intCoordin + 48)
    
End Function

'This sub will process the information
'retrieved form the INI file...
Public Sub subProcessData(ByVal strText As String, ByRef intWindowPosX As Integer, ByRef intWindowPosY As Integer, ByRef bolVisible As Boolean)
    Dim intTemp As Integer
    'Dim pola As Integer
    
    intTemp = InStr(strText, ":")
    strText = Mid(strText, intTemp + 1)
    
    intTemp = InStr(strText, ",")
    intWindowPosX = CInt(Left(strText, intTemp - 1))
    
    strText = Mid(strText, intTemp + 1)
    
    'še za drugo vejico
    intTemp = InStr(strText, ",")
    
    If (intTemp = 0) Then
        intWindowPosY = CInt(strText)
        Exit Sub
    End If
        
    'vejica je bila, se pravi...
    intWindowPosY = CInt(Left(strText, intTemp - 1))
    strText = Mid(strText, intTemp + 1)
    
    'še true/false
    strText = UCase(strText)
    
    If (Left(strText, 1) = "R") _
    Or (Left(strText, 1) = "T") Then
        bolVisible = True
        
    End If
    
    If (Left(strText, 1) = "N") _
    Or (Left(strText, 1) = "F") Then
        bolVisible = False
        
    End If
    
End Sub
