Attribute VB_Name = "Splosno"
Option Explicit
Global polozaji As String
Global belroslahkolevo As Boolean
Global crnroslahkolevo As Boolean
Global belroslahkodesno As Boolean
Global crnroslahkodesno As Boolean
Global mvarWhiteCastled As Boolean 'local copy
Global mvarBlackCastled As Boolean  'local copy
Global mvarLevels As Integer 'local copy
Global test As Boolean
Global nivo As Integer
Global tipka As Boolean

Function vrniFiguro(ByVal polozaji As String, ByVal poteza As String) As String
'funkcija sprejme dvoznakovno koordinato; vrne pa, kar
'je na tistem mestu

Dim pol As Integer, razX As Integer
pol = InStr(polozaji, poteza & ":")
Dim Cfigura As String
vrniFiguro = Mid(polozaji, pol + 3, 2)
End Function

Sub vrniRazliko(ByVal poteza As String, ByRef razX As Integer, ByRef razy As Integer)
razX = Asc(Mid(poteza, 4, 1)) - Asc(Mid(poteza, 1, 1))
razy = Val(Mid(poteza, 5, 1)) - Val(Mid(poteza, 2, 1))
End Sub

Function vrniVrednostKoordinate(ByVal kje As String) As Integer
'se gre za številko?
If Asc(kje) < 57 Then
    'ja!
    vrniVrednostKoordinate = Asc(kje) - 48
    Else
    'ne, je èrka
    vrniVrednostKoordinate = Asc(kje) - 64
    End If
End Function
Function vrniZnakLinije(ByVal vrednost As Integer) As String
vrniZnakLinije = Chr(vrednost + 64)
End Function
Function vrniZnakVrste(ByVal vrednost As Integer) As String
vrniZnakVrste = Chr(vrednost + 48)
End Function

Function diagonalaProsta(ByVal poteza As String) As Boolean
'funkcija preveri, ali je med zaèetnim in ciljnim poljem prosta pot
diagonalaProsta = True  'zaenkrat je prosto...

Dim xKor As Integer, yKor As Integer
Dim razX As Integer, razy As Integer

'ugotoviti moramo, ali je pot do cilja prosta
'zaèetna X koordinata je linija izvornega polja
xKor = vrniVrednostKoordinate(Left(poteza, 1))

'zaèetna Y koordinata pa vrsta
yKor = vrniVrednostKoordinate(Mid(poteza, 2, 1))

vrniRazliko poteza, razX, razy

Dim polozaj As String
Do
polozaj = vrniZnakLinije(xKor) & vrniZnakVrste(yKor)

If Trim(vrniFiguro(polozaji, polozaj)) <> "" Then
    'položaj je zaseden! diagonala ni prosta!
    'je to morda zadnji položaj?
    If xKor = vrniVrednostKoordinate(Left(poteza, 1)) And _
        yKor = vrniVrednostKoordinate(Mid(poteza, 2, 1)) Then GoTo nochk 'to je izhodišèni
                                                                        'položaj
    
    If xKor = vrniVrednostKoordinate(Mid(poteza, 4, 1)) And _
        yKor = vrniVrednostKoordinate(Mid(poteza, 5, 1)) Then Exit Do   'to je zadnji
                                                                        'položaj diagonale
    'tu pa ni! napaka!
    diagonalaProsta = False
    Exit Function
    End If

If xKor = vrniVrednostKoordinate(Mid(poteza, 4, 1)) And _
        yKor = vrniVrednostKoordinate(Mid(poteza, 5, 1)) Then Exit Do   'to je BIL zadnji
                                                                        'položaj diagonale
'zaenkrat je ok!

nochk:
xKor = xKor + Sgn(razX) 'prištej/odštej X koordinato
yKor = yKor + Sgn(razy) 'in Y


Loop

End Function

Function aliJeSah(ByVal cpolje As String, ByVal potezacrnega As Boolean) As Boolean
'fukcija preveri, ali na polje 'deluje'
'katera od nasprotnikovih figur
Dim xoff As Integer, yoff As Integer
Dim figura As String, ovirnopolje As String



aliJeSah = False 'zaenkrat ni šaha

'najprej bomo preverili naprej
figura = Trim(pojdiDoKoncaVrste(0, 1, cpolje, potezacrnega, ovirnopolje))
If preveriFiguroprav(figura, cpolje, ovirnopolje) = True Then
    aliJeSah = True
    Exit Function
    End If

'preverjamo nazaj
figura = Trim(pojdiDoKoncaVrste(0, -1, cpolje, potezacrnega, ovirnopolje))
If preveriFiguroprav(figura, cpolje, ovirnopolje) = True Then
    aliJeSah = True
    Exit Function
    End If


'preverjamo levo
figura = Trim(pojdiDoKoncaVrste(-1, 0, cpolje, potezacrnega, ovirnopolje))
If preveriFiguroprav(figura, cpolje, ovirnopolje) = True Then
    aliJeSah = True
    Exit Function
    End If


'preverjamo desno
figura = Trim(pojdiDoKoncaVrste(1, 0, cpolje, potezacrnega, ovirnopolje))
If preveriFiguroprav(figura, cpolje, ovirnopolje) = True Then
    aliJeSah = True
    Exit Function
    End If


'preverjamo levogor
figura = Trim(pojdiDoKoncaVrste(-1, 1, cpolje, potezacrnega, ovirnopolje))
If preveriFigurodiag(figura, cpolje, ovirnopolje, potezacrnega) = True Then
    aliJeSah = True
    Exit Function
    End If


'preverjamo desnogor
figura = Trim(pojdiDoKoncaVrste(1, 1, cpolje, potezacrnega, ovirnopolje))
If preveriFigurodiag(figura, cpolje, ovirnopolje, potezacrnega) = True Then
    aliJeSah = True
    Exit Function
    End If


'preverjamo levodol
figura = Trim(pojdiDoKoncaVrste(-1, -1, cpolje, potezacrnega, ovirnopolje))
If preveriFigurodiag(figura, cpolje, ovirnopolje, potezacrnega) = True Then
    aliJeSah = True
    Exit Function
    End If


'preverjamo desnodol
figura = Trim(pojdiDoKoncaVrste(1, -1, cpolje, potezacrnega, ovirnopolje))
If preveriFigurodiag(figura, cpolje, ovirnopolje, potezacrnega) = True Then
    aliJeSah = True
    Exit Function
    End If

'kaže, da ni niè narobe;

'RAZEN KONJA!!!!! Konja moramo preveriti posebej

    Dim a As Boolean
    If preveriSahSkakaca(polozaji, cpolje, potezacrnega) = True Then
            'kralj je v šahu po konjevi strani!
            aliJeSah = True
            End If
      
End Function

Function pojdiDoKoncaVrste(ByVal xoff As Integer, ByVal yoff As Integer, ByVal strtPolje As String, ByVal crniigra As Boolean, ByRef oviraPolje As String) As String
'funkcija se 'sprehodi' od danega polja v smeri odmika xoff in yoff; èe je na poti
'nasprotna figura, vrne njen položaj
'najprej dobimo vrednosti koordinat
Dim x As Integer, y As Integer
Dim figura As String, testpos As String

pojdiDoKoncaVrste = "  "  'zaenkrat smo prosti

x = vrniVrednostKoordinate(Left(strtPolje, 1))
y = vrniVrednostKoordinate(Right(strtPolje, 1))

Do
'poveèamo x in y
x = x + xoff
y = y + yoff

If x > 8 Or y > 8 Or x < 1 Or y < 1 Then
    'prišli do konca diagonale
    pojdiDoKoncaVrste = "  "
    Exit Function
    End If
'preverimo polje

testpos = vrniZnakLinije(x) & vrniZnakVrste(y)

figura = Trim(vrniFiguro(polozaji, testpos))

    If figura <> "" Then
    'nekaj smo našli!
    'preveimo, kakšne barve figura je!
    If crniigra = False And Left(figura, 1) = "C" Then
        'poteza belega; naletel na èrno figuro; šah
        oviraPolje = testpos
        pojdiDoKoncaVrste = figura
        Exit Function
        End If
        
    If crniigra = True And Left(figura, 1) = "B" Then
        'poteza èrnega; naletel na belo figuro; šah!
        oviraPolje = testpos
        pojdiDoKoncaVrste = figura
        Exit Function
        End If

    If crniigra = True And Left(figura, 1) = "C" Then
        'poteza èrnega; naletel na èrno figuro; ni šah
        Exit Function
        End If
        
    If crniigra = False And Left(figura, 1) = "B" Then
        'poteza belega; naletel na belo figuro; ni šah
        Exit Function
        End If
    
    End If
'polje je prosto
Loop

'èe je v èrti stoji figura, ki šaha ne daje, potem tiste za njo itak ne moreja dati šaha...

End Function

Sub postaviFiguro(ByRef polozaji As String, ByVal poteza As String, ByVal figura As String)
'poišèemo položaj

Dim pol As Integer, razX As Integer

pol = InStr(polozaji, poteza & ":")
Mid(polozaji, pol + 3, 2) = figura
End Sub

Function preveriSahSkakaca(ByRef mpolozaji As String, ByVal cpolje As String, ByVal potezacrnega As Boolean) As Boolean
'najlaže je, da preverimo vseh 8 smeri

'potezacrnega = Not (potezacrnega)
Dim x As Integer, y As Integer
Dim figura As String
Dim polozaj As String

'If cpolje = "D5" Then MsgBox "!"

preveriSahSkakaca = False

x = vrniVrednostKoordinate(Left(cpolje, 1))
y = vrniVrednostKoordinate(Right(cpolje, 1))

'najprej gledamo gor v obe smeri
'èe smo prek roba...
If (y + 2) > 8 Then GoTo preverigl
'ne, preveri oba stranska položaja; najprej desnega
If (x + 1) > 8 Then GoTo preverigl

'preverjamo zgoraj desno
polozaj = vrniZnakLinije(x + 1) & vrniZnakVrste(y + 2)
figura = vrniFiguro(mpolozaji, polozaj)
If Right(figura, 1) <> "S" Then GoTo preverigl
'figura je skakaè
If Left(figura, 1) = "B" And potezacrnega = True Then
    'konj je bel, igralec èrn;šah
    preveriSahSkakaca = True
    Exit Function
    End If
    
If Left(figura, 1) = "C" And potezacrnega = False Then
    'konj je èrn, igralec bel;šah
    preveriSahSkakaca = True
    Exit Function
    End If

preverigl:
'sedaj zgoraj levo
If (x - 1) < 1 Then GoTo preveridol
polozaj = vrniZnakLinije(x - 1) & vrniZnakVrste(y + 2)
figura = vrniFiguro(mpolozaji, polozaj)

If Right(figura, 1) <> "S" Then GoTo preveridol
'figura je skakaè
If Left(figura, 1) = "B" And potezacrnega = True Then
    'konj je bel, igralec èrn;šah
    preveriSahSkakaca = True
    Exit Function
    End If
If Left(figura, 1) = "C" And potezacrnega = False Then
    'konj je èrn, igralec bel;šah
    preveriSahSkakaca = True
    Exit Function
    End If

preveridol:
'najprej gledamo dol v obe smeri
'èe smo prek roba...
If (y - 2) < 1 Then GoTo preveridl
'ne, preveri oba stranska položaja; najprej desnega
If (x + 1) > 8 Then GoTo preveridl

'gledamo spodaj desno
polozaj = vrniZnakLinije(x + 1) & vrniZnakVrste(y - 2)
figura = vrniFiguro(mpolozaji, polozaj)
If Right(figura, 1) <> "S" Then GoTo preveridl
'figura je skakaè
If Left(figura, 1) = "B" And potezacrnega = True Then
    'konj je bel, igralec èrn;šah
    preveriSahSkakaca = True
    Exit Function
    End If
If Left(figura, 1) = "C" And potezacrnega = False Then
    'konj je èrn, igralec bel;šah
    preveriSahSkakaca = True
    Exit Function
    End If

preveridl:
'sedaj spodaj levega
If (x - 1) < 1 Then GoTo preverilevo
polozaj = vrniZnakLinije(x - 1) & vrniZnakVrste(y - 2)
figura = vrniFiguro(mpolozaji, polozaj)
If Right(figura, 1) <> "S" Then GoTo preverilevo
'figura je skakaè
If Left(figura, 1) = "B" And potezacrnega = True Then
    'konj je bel, igralec èrn;šah
    preveriSahSkakaca = True
    Exit Function
    End If
If Left(figura, 1) = "C" And potezacrnega = False Then
    'konj je èrn, igralec bel;šah
    preveriSahSkakaca = True
    Exit Function
    End If

preverilevo:
'èe smo prek roba...
If (y + 1) > 8 Then GoTo preverild
'ne, preveri oba stranska položaja; najprej desnega
If (x - 2) < 1 Then GoTo preverild

'najprej levo gor
polozaj = vrniZnakLinije(x - 2) & vrniZnakVrste(y + 1)
figura = vrniFiguro(mpolozaji, polozaj)
If Right(figura, 1) <> "S" Then GoTo preverild
'figura je skakaè
If Left(figura, 1) = "B" And potezacrnega = True Then
    'konj je bel, igralec èrn;šah
    preveriSahSkakaca = True
    Exit Function
    End If
If Left(figura, 1) = "C" And potezacrnega = False Then
    'konj je èrn, igralec bel;šah
    preveriSahSkakaca = True
    Exit Function
    End If

preverild:
If (y - 1) < 1 Then GoTo preveridesno
'potem levo dol
polozaj = vrniZnakLinije(x - 2) & vrniZnakVrste(y - 1)
figura = vrniFiguro(mpolozaji, polozaj)
If Right(figura, 1) <> "S" Then GoTo preveridesno
'figura je skakaè
If Left(figura, 1) = "B" And potezacrnega = True Then
    'konj je bel, igralec èrn;šah
    preveriSahSkakaca = True
    Exit Function
    End If
If Left(figura, 1) = "C" And potezacrnega = False Then
    'konj je èrn, igralec bel;šah
    preveriSahSkakaca = True
    Exit Function
    End If


preveridesno:
'èe smo prek roba...
If (y + 1) > 8 Then GoTo preveridd
'ne, preveri oba stranska položaja; najprej desnega
If (x + 2) > 8 Then GoTo preveridd

'najprej desno gor
polozaj = vrniZnakLinije(x + 2) & vrniZnakVrste(y + 1)
figura = vrniFiguro(mpolozaji, polozaj)
If Right(figura, 1) <> "S" Then GoTo preveridd
'figura je skakaè
If Left(figura, 1) = "B" And potezacrnega = True Then
    'konj je bel, igralec èrn;šah
    preveriSahSkakaca = True
    Exit Function
    End If
If Left(figura, 1) = "C" And potezacrnega = False Then
    'konj je èrn, igralec bel;šah
    preveriSahSkakaca = True
    Exit Function
    End If

preveridd:
If (y - 1) < 1 Then GoTo preverikonec
'potem desno dol
polozaj = vrniZnakLinije(x + 2) & vrniZnakVrste(y - 1)
figura = vrniFiguro(mpolozaji, polozaj)
If Right(figura, 1) <> "S" Then GoTo preverikonec
'figura je skakaè
If Left(figura, 1) = "B" And potezacrnega = True Then
    'konj je bel, igralec èrn;šah
    preveriSahSkakaca = True
    Exit Function
    End If
If Left(figura, 1) = "C" And potezacrnega = False Then
    'konj je èrn, igralec bel;šah
    preveriSahSkakaca = True
    Exit Function
    End If

preverikonec:
End Function

Function preveriRosado(ByRef polozaji As String, ByVal poteza As String, ByVal potezacrnega As Boolean) As Boolean

preveriRosado = True 'zaenkrat je OK

If poteza = "E1-G1" Then
    
    If test = True Then GoTo nog1e1
    If belroslahkodesno = False Then
        'rošada ni veè mogoèa!
        preveriRosado = False
        Exit Function
        End If
        
nog1e1:
    'preverimo polji med trdnjavo in kraljem, èe sta zasedeni
    If Trim(vrniFiguro(polozaji, "F1")) <> "" Then
        'na F1 je figura...
        preveriRosado = False
        Exit Function
        End If
        
    If Trim(vrniFiguro(polozaji, "G1")) <> "" Then
        'na G1 je figura...
        preveriRosado = False
        Exit Function
        End If
    
    'kar se prostora tièe, je OK. kaj pa polja med trdnjavo in kraljem? so pod šahom?
    If aliJeSah("E1", False) = True Then
        'kralj je pod šahom!
        preveriRosado = False
        Exit Function
        End If
    
    If aliJeSah("F1", False) = True Then
        'F1 je pod šahom!
        preveriRosado = False
        Exit Function
        End If
    
    If aliJeSah("G1", False) = True Then
        'G1 je pod šahom!
        preveriRosado = False
        Exit Function
        End If
    
    'rošada v to smer je mogoèa.
    mvarWhiteCastled = True
    preveriRosado = True
    End If
    
    
If poteza = "E8-G8" Then

    If test = True Then GoTo noe8g8
    If crnroslahkodesno = False Then
        'rošada ni veè mogoèa!
        preveriRosado = False
        Exit Function
        End If
        
        
noe8g8:
    'preverimo polji med trdnjavo in kraljem, èe sta zasedeni
    If Trim(vrniFiguro(polozaji, "F8")) <> "" Then
        'na F1 je figura...
        preveriRosado = False
        Exit Function
        End If
        
    If Trim(vrniFiguro(polozaji, "G8")) <> "" Then
        'na G1 je figura...
        preveriRosado = False
        Exit Function
        End If
    
    'kar se prostora tièe, je OK. kaj pa polja med trdnjavo in kraljem? so pod šahom?
    If aliJeSah("E8", True) = True Then
        'kralj je pod šahom!
        preveriRosado = False
        Exit Function
        End If
    
    If aliJeSah("F8", True) = True Then
        'F1 je pod šahom!
        preveriRosado = False
        Exit Function
        End If
    
    If aliJeSah("G8", True) = True Then
        'G1 je pod šahom!
        preveriRosado = False
        Exit Function
        End If
    
    'rošada v to smer je mogoèa.
    mvarBlackCastled = True
    preveriRosado = True
    
    End If

'=============================================
'desna stran je gotova. Še leva!
If poteza = "E1-C1" Then
    
    If test = True Then GoTo noe1c1
    If belroslahkolevo = False Then
        'rošada že od prej onemogoèena
        preveriRosado = False
        Exit Function
        End If
        
noe1c1:
    'preverimo polja med kraljem in trdnjavo
    
    If Trim(vrniFiguro(polozaji, "B1")) <> "" Then
        'na B1 je figura...
        preveriRosado = False
        Exit Function
        End If

    If Trim(vrniFiguro(polozaji, "C1")) <> "" Then
        'na C1 je figura...
        preveriRosado = False
        Exit Function
        End If

    If Trim(vrniFiguro(polozaji, "D1")) <> "" Then
        'na D1 je figura...
        preveriRosado = False
        Exit Function
        End If

 'do sem OK, preveri še polja, èe so pod šahom

    If aliJeSah("E1", False) = True Then
        'Kralj je pod šahom!
        preveriRosado = False
        Exit Function
        End If
    
    If aliJeSah("D1", False) = True Then
        'D1 je pod šahom!
        preveriRosado = False
        Exit Function
        End If

    If aliJeSah("C1", False) = True Then
        'Kralj je pod šahom!
        preveriRosado = False
        Exit Function
        End If
    'rokada v to smer je mogoèa
    mvarWhiteCastled = True
    preveriRosado = True
    End If
    
    
If poteza = "E8-C8" Then
    If test = True Then GoTo noe8c8
    If crnroslahkolevo = False Then
        'rošada že od prej onemogoèena
        preveriRosado = False
        Exit Function
        End If
        
noe8c8:
    'preverimo polja med kraljem in trdnjavo
    If Trim(vrniFiguro(polozaji, "B8")) <> "" Then
        'na B1 je figura...
        preveriRosado = False
        Exit Function
        End If

    If Trim(vrniFiguro(polozaji, "C8")) <> "" Then
        'na C1 je figura...
        preveriRosado = False
        Exit Function
        End If

    If Trim(vrniFiguro(polozaji, "D8")) <> "" Then
        'na D1 je figura...
        preveriRosado = False
        Exit Function
        End If

 'do sem OK, preveri še polja, èe so pod šahom

    If aliJeSah("E8", True) = True Then
        'Kralj je pod šahom!
        preveriRosado = False
        Exit Function
        End If
    
    If aliJeSah("D8", True) = True Then
        'D1 je pod šahom!
        preveriRosado = False
        Exit Function
        End If

    If aliJeSah("C8", True) = True Then
        'C8 je pod šahom!
        preveriRosado = False
        Exit Function
        End If
    'rokada v to smer je mogoèa
    mvarBlackCastled = True
    preveriRosado = True
    End If
    
End Function

Function preveriFiguroprav(ByVal figura As String, ByVal cpolje As String, ByVal ovirnopolje As String) As Boolean

preveriFiguroprav = False   'zaenkrat je vse OK

If Trim(figura) = "" Then Exit Function 'èe figure ni...

Dim razX As Integer, razy As Integer

Select Case Right(figura, 1)
    Case "Q"
    'kraljica;to je šah
    preveriFiguroprav = True
    Exit Function
    
    Case "T"
    'trdnjava, to je šah
    preveriFiguroprav = True
    Exit Function
    
    Case "K"
    'kralj; razlika ne sme biti veè kot 1 polje
    vrniRazliko cpolje & "-" & ovirnopolje, razX, razy
    If Abs(razX) < 2 And Abs(razy) < 2 Then
        preveriFiguroprav = True
        Exit Function
        End If
    
    'druge figure pa niso nevarne.
    End Select
    
End Function


Function preveriFigurodiag(ByVal figura As String, ByVal cpolje As String, ByVal ovirnopolje As String, ByVal potezacrnega As Boolean) As Boolean

preveriFigurodiag = False
If Trim(figura) = "" Then Exit Function


Dim razX As Integer, razy As Integer
Select Case Right(figura, 1)
    Case "Q"
    'kraljica; to je šah!
    preveriFigurodiag = True
    Exit Function
    
    Case "L"
    'lovec; to je šah
    preveriFigurodiag = True
    Exit Function
    
    Case "K"
    'kralj; razlika ne sme biti veè kot 1 polje
     vrniRazliko cpolje & "-" & ovirnopolje, razX, razy
    If Abs(razX) < 2 And Abs(razy) < 2 Then
        preveriFigurodiag = True
        Exit Function
        End If
    
    Case "P"
    'kmet! sedaj pa katere barve je igralec in kje stoji
    vrniRazliko cpolje & "-" & ovirnopolje, razX, razy
    If potezacrnega = False Then
        'igralec je bele barve, kar pomeni, da mora kmet biti eno polje diagonalno zadaj
        'abs(razx)=1, razy=1
        If Abs(razX) = 1 And razy = 1 Then
            'kmet je spredaj v diagonali;šah
            preveriFigurodiag = True
            Exit Function
            End If
        
        Else
        'èe pa je igralec èrne barve, pomeni, da mora kmet biti eno polje diagonalno zadaj
        'abs(razx)=1, razy=-1
        If Abs(razX) = 1 And razy = -1 Then
            'kmet je zadaj diagonalno;šah
            preveriFigurodiag = True
            Exit Function
            End If
        End If
        
        'druge figure nam pa niso nevarne
        End Select
End Function

'Function sestavipolozaje() As String
'tu moramo iz polozajev sestaviti string:
'Dim x As Integer, y As Integer
'Dim pol As String
'pol = ""
'
'For y = 1 To 8
'    For x = 1 To 8
'    If Trim(Left(polozaji(x, y), 1)) = Chr(0) Then polozaji(x, y) = "  "
'
'    pol = pol & Chr(x + 64) & Trim(Str(y)) & ":" & polozaji(x, y) & "|"
'    Next
'Next
'sestavipolozaje = pol
'End Function

Function aliJeSahIskanja(ByRef kpolozaji As String, cpolje As String, ByVal potezacrnega As Boolean) As String
'fukcija preveri, ali na polje 'deluje'
'katera od nasprotnikovih figur in vrne vse napadalna polja
Dim xoff As Integer, yoff As Integer
Dim figura As String, ovirnopolje As String

aliJeSahIskanja = "" 'zaenkrat ni šaha

'najprej bomo preverili naprej
figura = Trim(pojdiDoKoncaVrsteIskanja(kpolozaji, 0, 1, cpolje, potezacrnega, ovirnopolje))
If preveriFiguroprav(figura, cpolje, ovirnopolje) = True Then
    aliJeSahIskanja = aliJeSahIskanja & ovirnopolje
    End If

'preverjamo nazaj
figura = Trim(pojdiDoKoncaVrsteIskanja(kpolozaji, 0, -1, cpolje, potezacrnega, ovirnopolje))
If preveriFiguroprav(figura, cpolje, ovirnopolje) = True Then
    aliJeSahIskanja = aliJeSahIskanja & ovirnopolje
    End If


'preverjamo levo
figura = Trim(pojdiDoKoncaVrsteIskanja(kpolozaji, -1, 0, cpolje, potezacrnega, ovirnopolje))
If preveriFiguroprav(figura, cpolje, ovirnopolje) = True Then
    aliJeSahIskanja = aliJeSahIskanja & ovirnopolje
    End If


'preverjamo desno
figura = Trim(pojdiDoKoncaVrsteIskanja(kpolozaji, 1, 0, cpolje, potezacrnega, ovirnopolje))
If preveriFiguroprav(figura, cpolje, ovirnopolje) = True Then
    aliJeSahIskanja = aliJeSahIskanja & ovirnopolje
    End If


'preverjamo levogor
figura = Trim(pojdiDoKoncaVrsteIskanja(kpolozaji, -1, 1, cpolje, potezacrnega, ovirnopolje))
If preveriFigurodiag(figura, cpolje, ovirnopolje, potezacrnega) = True Then
    aliJeSahIskanja = aliJeSahIskanja & ovirnopolje
    End If


'preverjamo desnogor
figura = Trim(pojdiDoKoncaVrsteIskanja(kpolozaji, 1, 1, cpolje, potezacrnega, ovirnopolje))
If preveriFigurodiag(figura, cpolje, ovirnopolje, potezacrnega) = True Then
    aliJeSahIskanja = aliJeSahIskanja & ovirnopolje
    End If


'preverjamo levodol
figura = Trim(pojdiDoKoncaVrsteIskanja(kpolozaji, -1, -1, cpolje, potezacrnega, ovirnopolje))
If preveriFigurodiag(figura, cpolje, ovirnopolje, potezacrnega) = True Then
    aliJeSahIskanja = aliJeSahIskanja & ovirnopolje
    End If


'preverjamo desnodol
figura = Trim(pojdiDoKoncaVrsteIskanja(kpolozaji, 1, -1, cpolje, potezacrnega, ovirnopolje))
If preveriFigurodiag(figura, cpolje, ovirnopolje, potezacrnega) = True Then
    aliJeSahIskanja = aliJeSahIskanja & ovirnopolje
    End If

'kaže, da ni niè narobe;

'RAZEN KONJA!!!!! Konja moramo preveriti posebej

    aliJeSahIskanja = aliJeSahIskanja & preveriSahSkakacaIskanja(kpolozaji, cpolje, potezacrnega)
      
End Function


Function pojdiDoKoncaVrsteIskanja(ByRef posicija As String, xoff As Integer, ByVal yoff As Integer, ByVal strtPolje As String, ByVal crniigra As Boolean, ByRef oviraPolje As String) As String
'funkcija se 'sprehodi' od danega polja v smeri odmika xoff in yoff; èe je na poti
'nasprotna figura, vrne njen položaj
'najprej dobimo vrednosti koordinat
Dim x As Integer, y As Integer
Dim figura As String, testpos As String

pojdiDoKoncaVrsteIskanja = "  "  'zaenkrat smo prosti

x = vrniVrednostKoordinate(Left(strtPolje, 1))
y = vrniVrednostKoordinate(Right(strtPolje, 1))

Do
'poveèamo x in y
x = x + xoff
y = y + yoff

If x > 8 Or y > 8 Or x < 1 Or y < 1 Then
    'prišli do konca diagonale
    pojdiDoKoncaVrsteIskanja = "  "
    Exit Function
    End If
'preverimo polje

testpos = vrniZnakLinije(x) & vrniZnakVrste(y)

figura = Trim(vrniFiguro(posicija, testpos))

    If figura <> "" Then
    'nekaj smo našli!
    'preveimo, kakšne barve figura je!
    If crniigra = False And Left(figura, 1) = "C" Then
        'poteza belega; naletel na èrno figuro; šah
        oviraPolje = testpos
        pojdiDoKoncaVrsteIskanja = figura
        Exit Function
        End If
        
    If crniigra = True And Left(figura, 1) = "B" Then
        'poteza èrnega; naletel na belo figuro; šah!
        oviraPolje = testpos
        pojdiDoKoncaVrsteIskanja = figura
        Exit Function
        End If

    If crniigra = True And Left(figura, 1) = "C" Then
        'poteza èrnega; naletel na èrno figuro; ni šah
        Exit Function
        End If
        
    If crniigra = False And Left(figura, 1) = "B" Then
        'poteza belega; naletel na belo figuro; ni šah
        Exit Function
        End If
    
    End If
'polje je prosto
Loop

'èe je v èrti stoji figura, ki šaha ne daje, potem tiste za njo itak ne moreja dati šaha...

End Function

Function preveriSahSkakacaIskanja(ByRef mpolozaji As String, ByVal cpolje As String, ByVal potezacrnega As Boolean) As String
'najlaže je, da preverimo vseh 8 smeri
Dim x As Integer, y As Integer
Dim figura As String
Dim polozaj As String

preveriSahSkakacaIskanja = ""

x = vrniVrednostKoordinate(Left(cpolje, 1))
y = vrniVrednostKoordinate(Right(cpolje, 1))

'najprej gledamo gor v obe smeri
'èe smo prek roba...
If (y + 2) > 8 Then GoTo ipreverigl
'ne, preveri oba stranska položaja; najprej desnega
If (x + 1) > 8 Then GoTo ipreverigl

'preverjamo zgoraj desno
polozaj = vrniZnakLinije(x + 1) & vrniZnakVrste(y + 2)
figura = vrniFiguro(mpolozaji, polozaj)
If Right(figura, 1) <> "S" Then GoTo ipreverigl
'figura je skakaè
If Left(figura, 1) = "B" And potezacrnega = True Then
    'konj je bel, igralec èrn;šah
    preveriSahSkakacaIskanja = preveriSahSkakacaIskanja & polozaj
    End If
    
If Left(figura, 1) = "C" And potezacrnega = False Then
    'konj je èrn, igralec bel;šah
    preveriSahSkakacaIskanja = preveriSahSkakacaIskanja & polozaj
    End If

ipreverigl:
'sedaj zgoraj levega
If (x - 1) < 1 Then GoTo ipreveridol
polozaj = vrniZnakLinije(x - 1) & vrniZnakVrste(y + 2)
figura = vrniFiguro(mpolozaji, polozaj)

If Right(figura, 1) <> "S" Then GoTo ipreveridol
'figura je skakaè
If Left(figura, 1) = "B" And potezacrnega = True Then
    'konj je bel, igralec èrn;šah
    preveriSahSkakacaIskanja = preveriSahSkakacaIskanja & polozaj
    End If
If Left(figura, 1) = "C" And potezacrnega = False Then
    'konj je èrn, igralec bel;šah
    preveriSahSkakacaIskanja = preveriSahSkakacaIskanja & polozaj
    End If

ipreveridol:
'najprej gledamo dol v obe smeri
'èe smo prek roba...
If (y - 2) < 1 Then GoTo ipreveridl
'ne, preveri oba stranska položaja; najprej desnega
If (x + 1) > 8 Then GoTo ipreveridl

'gledamo spodaj desno
polozaj = vrniZnakLinije(x + 1) & vrniZnakVrste(y - 2)
figura = vrniFiguro(mpolozaji, polozaj)
If Right(figura, 1) <> "S" Then GoTo ipreveridl
'figura je skakaè
If Left(figura, 1) = "B" And potezacrnega = True Then
    'konj je bel, igralec èrn;šah
    preveriSahSkakacaIskanja = preveriSahSkakacaIskanja & polozaj
    End If
If Left(figura, 1) = "C" And potezacrnega = False Then
    'konj je èrn, igralec bel;šah
    preveriSahSkakacaIskanja = preveriSahSkakacaIskanja & polozaj
    End If

ipreveridl:
'sedaj spodaj levega
If (x - 1) < 1 Then GoTo ipreverilevo
polozaj = vrniZnakLinije(x - 1) & vrniZnakVrste(y - 2)
figura = vrniFiguro(mpolozaji, polozaj)
If Right(figura, 1) <> "S" Then GoTo ipreverilevo
'figura je skakaè
If Left(figura, 1) = "B" And potezacrnega = True Then
    'konj je bel, igralec èrn;šah
    preveriSahSkakacaIskanja = preveriSahSkakacaIskanja & polozaj
    End If
If Left(figura, 1) = "C" And potezacrnega = False Then
    'konj je èrn, igralec bel;šah
    preveriSahSkakacaIskanja = preveriSahSkakacaIskanja & polozaj
    End If

ipreverilevo:
'èe smo prek roba...
If (y + 1) > 8 Then GoTo ipreverild
'ne, preveri oba stranska položaja; najprej desnega
If (x - 2) < 1 Then GoTo ipreverild

'najprej levo gor
polozaj = vrniZnakLinije(x - 2) & vrniZnakVrste(y + 1)
figura = vrniFiguro(mpolozaji, polozaj)
If Right(figura, 1) <> "S" Then GoTo ipreverild
'figura je skakaè
If Left(figura, 1) = "B" And potezacrnega = True Then
    'konj je bel, igralec èrn;šah
    preveriSahSkakacaIskanja = preveriSahSkakacaIskanja & polozaj
    End If
If Left(figura, 1) = "C" And potezacrnega = False Then
    'konj je èrn, igralec bel;šah
    preveriSahSkakacaIskanja = preveriSahSkakacaIskanja & polozaj
    End If

ipreverild:
If (y - 1) < 1 Then GoTo ipreveridesno
'potem levo dol
polozaj = vrniZnakLinije(x - 2) & vrniZnakVrste(y - 1)
figura = vrniFiguro(mpolozaji, polozaj)
If Right(figura, 1) <> "S" Then GoTo ipreveridesno
'figura je skakaè
If Left(figura, 1) = "B" And potezacrnega = True Then
    'konj je bel, igralec èrn;šah
    preveriSahSkakacaIskanja = preveriSahSkakacaIskanja & polozaj
    End If
If Left(figura, 1) = "C" And potezacrnega = False Then
    'konj je èrn, igralec bel;šah
    preveriSahSkakacaIskanja = preveriSahSkakacaIskanja & polozaj
    End If


ipreveridesno:
'èe smo prek roba...
If (y + 1) > 8 Then GoTo ipreveridd
'ne, preveri oba stranska položaja; najprej desnega
If (x + 2) > 8 Then GoTo ipreveridd

'najprej desno gor
polozaj = vrniZnakLinije(x + 2) & vrniZnakVrste(y + 1)
figura = vrniFiguro(mpolozaji, polozaj)
If Right(figura, 1) <> "S" Then GoTo ipreveridd
'figura je skakaè
If Left(figura, 1) = "B" And potezacrnega = True Then
    'konj je bel, igralec èrn;šah
    preveriSahSkakacaIskanja = preveriSahSkakacaIskanja & polozaj
    End If
If Left(figura, 1) = "C" And potezacrnega = False Then
    'konj je èrn, igralec bel;šah
    preveriSahSkakacaIskanja = preveriSahSkakacaIskanja & polozaj
    End If

ipreveridd:
If (y - 1) < 1 Then GoTo ipreverikonec
'potem desno dol
polozaj = vrniZnakLinije(x + 2) & vrniZnakVrste(y - 1)
figura = vrniFiguro(mpolozaji, polozaj)
If Right(figura, 1) <> "S" Then GoTo ipreverikonec
'figura je skakaè
If Left(figura, 1) = "B" And potezacrnega = True Then
    'konj je bel, igralec èrn;šah
    preveriSahSkakacaIskanja = preveriSahSkakacaIskanja & polozaj
    End If
If Left(figura, 1) = "C" And potezacrnega = False Then
    'konj je èrn, igralec bel;šah
    preveriSahSkakacaIskanja = preveriSahSkakacaIskanja & polozaj
    End If

ipreverikonec:
End Function
