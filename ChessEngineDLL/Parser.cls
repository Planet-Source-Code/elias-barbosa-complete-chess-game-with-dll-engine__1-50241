VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Parser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarBlackTurn As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarErrorNumber As Integer 'local copy
Private mvarErrorText As String 'local copy
'local variable(s) to hold property value(s)
Private mvarMove As String 'local copy
'local variable(s) to hold property value(s)
'local variable(s) to hold property value(s)
'local variable(s) to hold property value(s)
Private mvarCapture As Boolean 'local copy
'local variable(s) to hold property value(s)
'local variable(s) to hold property value(s)
Private mvarCheck As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarWhiteCanCastleLeft As Boolean 'local copy
Private mvarBlackCanCastleLeft As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarBlackCanCastleRight As Boolean 'local copy
Private mvarWhiteCanCastleRight As Boolean 'local copy
'local variable(s) to hold property value(s)
'local variable(s) to hold property value(s)
Private mvarMoveSearch As Boolean 'local copy
Public Property Let MoveSearch(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MoveSearch = 5
    mvarMoveSearch = vData
    test = vData
End Property

Public Property Get BlackCastled() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BlackCastled
    BlackCastled = mvarBlackCastled
End Property


Public Property Get WhiteCastled() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WhiteCastled
    WhiteCastled = mvarWhiteCastled
End Property

Public Property Let WhiteCanCastleRight(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WhiteCanCastleright = 5
    mvarWhiteCanCastleRight = vData
    belroslahkodesno = vData
End Property


Public Property Get WhiteCanCastleRight() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WhiteCanCastleright
    WhiteCanCastleRight = belroslahkodesno
    
End Property



Public Property Let BlackCanCastleRight(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BlackCanCastleRight = 5
    mvarBlackCanCastleRight = vData
    crnroslahkodesno = vData
End Property


Public Property Get BlackCanCastleRight() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BlackCanCastleRight
    BlackCanCastleRight = crnroslahkodesno
End Property



Public Property Let BlackCanCastleLeft(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BlackCanCastle = 5
    
    crnroslahkolevo = vData
End Property


Public Property Get BlackCanCastleLeft() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BlackCanCastle
    BlackCanCastleLeft = crnroslahkolevo
End Property



Public Property Let WhiteCanCastleLeft(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WhiteCanCastle = 5
    
    belroslahkolevo = vData
End Property


Public Property Get WhiteCanCastleLeft() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WhiteCanCastle
    WhiteCanCastleLeft = belroslahkolevo
End Property


Public Property Get Check() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Check
    Check = mvarCheck
End Property

Public Property Get Capture() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Capture
    Capture = mvarCapture
End Property

Public Property Let Move(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Move = 5
    mvarMove = vData
End Property

Public Property Get ErrorText() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ErrorText
    ErrorText = mvarErrorText
End Property

Public Property Get ErrorNumber() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ErrorNumber
    ErrorNumber = mvarErrorNumber
End Property



Public Property Let BlackTurn(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BlackTurn = 5
    mvarBlackTurn = vData
End Property
Public Sub Parse(Position As String)

polozaji = Position 'globalna kopija polo�ajev

Dim xkoord As Integer, ykoord As Integer
Dim znak As String * 1  '�rka oz. �t.
Dim figura As String * 2
Dim pol As Integer


mvarErrorNumber = 0
mvarErrorText = ""



pol = InStr(Position, Left(mvarMove, 2) & ":")

'pa poglejmo

figura = Mid(Position, pol + 3, 2)
If Trim(Mid(Position, pol + 3, 2)) = "" Then
    'ni nobene figure!
    mvarErrorText = "There is no piece to move!"
    mvarErrorNumber = 1
    Exit Sub
    End If
    
'sedaj ugotvimo, ali je figura v skladu z barvo
'dobimo prvo �rko, ki je barva polo�aja
znak = Mid(Position, pol + 3, 1)
If znak = "C" And mvarBlackTurn = False Then
    'napa�na barva figure!
    mvarErrorText = "Wrong piece color!"
    mvarErrorNumber = 2
    Exit Sub
    End If
    
If znak = "B" And mvarBlackTurn = True Then
    'napa�na barva figure!
    mvarErrorText = "Wrong piece color!"
    mvarErrorNumber = 2
    Exit Sub
    End If

'pravzaprav smo izbrani polo�aj s tem preverili, in je veljaven
If Len(mvarMove) = 2 Then Exit Sub

'==================================================
'najprej bi preverili, ali se kdo ho�e premakniti
'na isto mesto
If Left(mvarMove, 2) = Right(mvarMove, 2) Then
    mvarErrorText = "You have to move piece!"
    mvarErrorNumber = 4
    Exit Sub
    End If


'potem preverimo, ali je to veljaven premik, hkrati pa �e, �e je pot odprta
znak = Mid(Position, pol + 4, 1)    'tip figure
Select Case znak
Case "P"
    If preverikmeta(mvarBlackTurn, mvarMove, figura) = False Then
    'nedovoljen premik kmeta!
    mvarErrorText = "Error in pawn's move!"
    mvarErrorNumber = 5
    Exit Sub
    End If
        
Case "T"
    If preveriTrdnjavo(mvarBlackTurn, mvarMove, figura) = False Then
    mvarErrorText = "Error in rook'2 move!"
    mvarErrorNumber = 6
    Exit Sub
    End If
    
        
Case "S"
    If preveriSkakaca(mvarBlackTurn, mvarMove, figura) = False Then
    mvarErrorText = "Error in knight's move!"
    mvarErrorNumber = 7
    Exit Sub
    End If
    
    
Case "L"
    If preveriLovca(mvarBlackTurn, mvarMove, figura) = False Then
    mvarErrorText = "Error in bishop's move!"
    mvarErrorNumber = 8
    Exit Sub
    End If
    
    
Case "Q"
    If preveriKraljico(mvarBlackTurn, mvarMove, figura) = False Then
    mvarErrorText = "Error in queen's move!"
    mvarErrorNumber = 9
    Exit Sub
    End If
    

Case "K"
    Dim kajkralj As Boolean
    
    kajkralj = preveriKralja(mvarBlackTurn, mvarMove, figura)
    If kajkralj = False Then
    mvarErrorText = "Error in king's move!"
    mvarErrorNumber = 10
    Exit Sub
    End If
    
    'sedaj pa �e preverimo, ali bo igral�ev
    'kralj po premiku v �ahu...
    If aliJeSah(Mid(mvarMove, 4, 2), mvarBlackTurn) = True Then
        mvarErrorNumber = 11
        mvarErrorText = "That move would put your king in chess!"
        Exit Sub
        End If

End Select



'sedaj pa bomo preverili, kaj je na polju!

pol = InStr(Position, Right(mvarMove, 2) & ":") 'ciljno polje
'If Trim(Mid(Position, pol + 3, 2)) <>"" then goto


znak = Mid(Position, pol + 3, 1)
If znak = "B" And mvarBlackTurn = False Then
    'napa�na barva figure;figura je iste barve kot plenilec
    mvarErrorText = "You cannot capture youre own piece!"
    mvarErrorNumber = 3
    Exit Sub
    End If

If znak = "C" And mvarBlackTurn = True Then
    'napa�na barva figure;figura je iste barve kot plenilec
    mvarErrorText = "You cannot capture youre own piece!"
    mvarErrorNumber = 3
    Exit Sub
    End If


'smo v �ah spravili nasprotnega kralja?
    'poi��emo njegov polo�aj

    Dim krPol As String
    
    If mvarBlackTurn = False Then
        'i��emo polo�aj crnega kralja; igral je beli
        pol = InStr(polozaji, "CK")
        krPol = Mid(polozaji, pol - 3, 2)
        Else
        'i��emo belega; igral je �rni
        pol = InStr(polozaji, "BK")
        krPol = Mid(polozaji, pol - 3, 2)
        End If
    
    'polo�aj imamo.
    Dim ciljfig As String, startfig As String
    'sedaj pa moramo na ciljno polje navidezno prestaviti na�o figuro
    ciljfig = vrniFiguro(polozaji, Right(mvarMove, 2))
    startfig = vrniFiguro(polozaji, Left(mvarMove, 2))
        
    'in na polo�aj prestavimo na�o
    postaviFiguro polozaji, Right(mvarMove, 2), vrniFiguro(polozaji, Left(mvarMove, 2))
    'pa izpraznemo �tartni polo�aj
    postaviFiguro polozaji, Left(mvarMove, 2), "  "
    
    'poi��emo, ali je tuji kralj v �ahu
    Dim crni As Boolean
    If mvarBlackTurn = True Then
        crni = False
        Else
        crni = True
        End If
        
    If aliJeSah(krPol, crni) = True Then
        's to potezo smo spro�ili �ah!
        mvarCheck = True
        End If
    
    'sedaj pa bomo pogledali, ali je NA� kralj po potezi v �ahu?
    
    If mvarBlackTurn = True Then
        'i��emo polo�aj �rnega kralja; igral je �rni
        pol = InStr(polozaji, "CK")
        
        krPol = Mid(polozaji, pol - 3, 2)
        Else
        'i��emo belega; igral je beli
        pol = InStr(polozaji, "BK")
        krPol = Mid(polozaji, pol - 3, 2)
        End If
    
    
    If aliJeSah(krPol, mvarBlackTurn) = True Then
        'kljub potezi je kralj v �ahu!
        mvarErrorText = "King is still in chess!"
        mvarErrorNumber = 12
        'restavriramo polo�aja
        postaviFiguro polozaji, Right(mvarMove, 2), ciljfig
        postaviFiguro polozaji, Left(mvarMove, 2), startfig
        Exit Sub
        End If
    
    
    'restavriramo polo�aja
    postaviFiguro polozaji, Right(mvarMove, 2), ciljfig
    postaviFiguro polozaji, Left(mvarMove, 2), startfig
    
    'poglejmo glede ro�ade!
    If test = True Then GoTo chkros
    
    If Left(mvarMove, 2) = "A1" Then
        'premik bele trdnjave; ro�ada ni ve� mogo�a!
        belroslahkolevo = False
        End If
    
    If Left(mvarMove, 2) = "A8" Then
        'premik bele trdnjave; ro�ada ni ve� mogo�a!
        crnroslahkolevo = False
        End If
    
    If Left(mvarMove, 2) = "H1" Then
        'premik bele trdnjave; ro�ada ni ve� mogo�a!
        belroslahkodesno = False
        End If
    
    If Left(mvarMove, 2) = "H8" Then
        'premik bele trdnjave; ro�ada ni ve� mogo�a!
        crnroslahkodesno = False
        End If
chkros:
If znak = "C" And mvarBlackTurn = False Then
    'zajetje figure!
    mvarCapture = True
    Exit Sub
    End If

If znak = "B" And mvarBlackTurn = True Then
    'zajetje figure
    mvarCapture = True
    Exit Sub
    End If



End Sub

Private Sub Class_Initialize()
mvarErrorNumber = 0
mvarBlackTurn = False
mvarCapture = False
mvarCheck = False
mvarBlackCastled = False
mvarWhiteCastled = False
test = False    'po defaultu se igra
End Sub
