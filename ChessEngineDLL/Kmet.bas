Attribute VB_Name = "Kmet"
Option Explicit

Function preverikmeta(ByVal crninavrsti As Boolean, ByVal poteza As String, ByVal figura As String) As Boolean
Dim bel As Boolean
preverikmeta = True 'zaenkrat nimamo razloga za paniko

'kak�ne barve je?
If Left(figura, 1) = "B" Then
    bel = True
    Else
    bel = False
    End If

'ali je razlika med premikom v Y osi 2?
Dim razy As Integer, razX As Integer

vrniRazliko poteza, razX, razy


If razy < 0 And crninavrsti = False Then
    'premik belega nazaj
    preverikmeta = False
    Exit Function
    End If
    
If razy > 0 And crninavrsti = True Then
    'premik �rnega naprej
    preverikmeta = False
    Exit Function
    End If

'za koliko se je premaknil?
If Abs(razX) > 1 Then
    'kmet je zavil ve� kot eno polje vstran
    preverikmeta = False
    Exit Function
    End If

If Abs(razy) > 2 Then
    'premik naprej/nazaj za ve� kot dve polji. Napaka
    preverikmeta = False
    Exit Function
    End If

'torej, kmet je �el v pravo smer za najve� 2 polji in lahko da je zavil.

'najprej poglejmo, ali je �el naprej za dve polji
'to lahko stori samo, �e �tarta iz osnovne vrste
If Abs(razy) = 2 Then
    '�e je beli, je moral iz vrste 2
    If crninavrsti = False And Mid(poteza, 2, 1) <> "2" Then
        '�el je iz druge vrste, kot bi smel pri premiku za dve polji!
        preverikmeta = False
        Exit Function
        End If
        
    'kaj pa  �rni?
    If crninavrsti = True And Mid(poteza, 2, 1) <> "7" Then
        '�el je iz druge vrste kot bi smel!
        preverikmeta = False
        Exit Function
        End If
        
    'pravilno je �el dve polji naprej in je zavil?
    If Abs(razX) > 0 Then
        'zavil je. Ne bi smel!
        preverikmeta = False
        Exit Function
        End If
        
    'je vmes kaka figura?
    'pri crnem?
    If crninavrsti = True And Trim(vrniFiguro(polozaji, Left(poteza, 1) & "6")) <> "" Then
        'figura je vmes!
        preverikmeta = False
        Exit Function
        End If
        
    'pri belem
    If crninavrsti = False And Trim(vrniFiguro(polozaji, Left(poteza, 1) & "3")) <> "" Then
        'figura je vmes!
        preverikmeta = False
        Exit Function
        End If
    
    'ali je kon�no polje prosto?
    If Trim(vrniFiguro(polozaji, Mid(poteza, 4, 2))) <> "" Then
        'na kon�nem polju je figura. Napaka!
        preverikmeta = False
        Exit Function
        End If
    
    End If
    
    'tu je moral iti za eno polje in lahko da je zavil
    
    'je na kon�nem polju kaka figura
    Dim fig As String
    fig = Trim(vrniFiguro(polozaji, Mid(poteza, 4, 2)))
    
    
    If Abs(razX) = 1 Then
        'zavil je!
        If Abs(razy) <> 1 Then
            '�reti ho�e vstran
            preverikmeta = False
            Exit Function
            End If
        
        
        'naskakuje lastno figuro?
        If crninavrsti = False And Left(fig, 1) = "B" Then
            'po�reti ho�e svojo figuro!
            preverikmeta = False
            Exit Function
            End If
            
        If crninavrsti = True And Left(fig, 1) = "C" Then
            'po�reti ho�e lastno figuro
            preverikmeta = False
            Exit Function
            End If
        'pa je kaj za po�reti?
        
        If fig = "" Then
            'ni ni�esar!
            preverikmeta = False
            Exit Function
            End If
        
        End If
        
        
    'tu je lahko �el samo �e za eno polje naprej. je tam kaka figura?
    If Abs(razX) = 0 And fig <> "" Then
        'tam je figura
        preverikmeta = False
        Exit Function
    End If
        
End Function
