Attribute VB_Name = "Eigschaftsverr"
Option Explicit

Private AntrScheibe As Boolean
Private UmlScheibe As Boolean
Private Achsabstand As Double
Private Awert As Double 'gilt modulweit, wird übergeben
Private Nwert As Double 'gilt modulweit, wird übergeben
Private R1 As Integer
Private R2 As Integer


Public Sub Verrechnung(Elt As Integer, Eig As Integer, ByVal i As Double, ByVal j As Double)
    'i,indecs,merk,gabsnich waren's früher
    'element, eigenschaft, alte einstellung, neue einstellung
    'wenn ok, dann alte einstellung über bord
    'nwert beinhaltet die eingabe, bis sie erst am ende nach überprüfung übernommen wird
    
    Dim M1 As Double, M2 As Double
    Dim N As Integer, M As Integer, L As Integer, Abbruch As Boolean, abruch As Boolean
    
    
    AnlageRefresh = False
    Nwert = i
    Awert = j
    
    Abbruch = False
    Mother.H = ""
    
   
    'test auf maximalwerte
    If Nwert < El(Eig).Minimum Then 'zu kleine Werte abfangen
        Beep
        Mother.H = Lang_Res(803) & El(Eig).Minimum  'kleinstmöglicher wert
        abruch = True
        Exit Sub
    End If
    If Nwert > El(Eig).Maximum Then 'zu große Werte abfangen
        Beep
        Mother.H = Lang_Res(804) & El(Eig).Maximum  'Größtmöglicher Wert
        Abbruch = True
        Exit Sub
    End If
    
    Select Case Eig
        Case 1 'Auflegedehnung
            For N = 1 To Maxelementindex
                If Sys(N).E(54) <> 0 Then
                    Mother.H = Lang_Res(805) & N  'erst alle vorgeg. Spannkräfte entfernen. Element:
                    Abbruch = True
                    Exit Sub
                End If
            Next N
            If Nwert = 0 Then
                Mother.H = Lang_Res(807) 'Auflegedehnungsvorgabe entfernt
            End If
        Case 2 'durchmesser
            Sys(Elt).E(2) = Nwert
            If Sys(Elt).E(3) = 0 Then
                Sys(Elt).E(3) = Nwert - 10 '5mm wandstärke
            End If
            If Nwert <= Sys(Elt).E(3) + 10 Then
                Sys(Elt).E(3) = Nwert - 10 '5mm wandstärke
                Mother.H = Lang_Res(808)  'Innendurchmesser wurde an Außendurchmesser angepaßt'
            End If
            If Sys(Elt).E(3) < 0 Then Sys(Elt).E(3) = 0
            Call Interner_Abgleich(Elt, 2) 'erstmal sich selbst in Ordnung bringen
            Call Geschwindigkeitsanpassung(Elt) 'alle anderen anpassen
            'Call Bandmindestlängenberechnung(Eig)
        Case 3 'innendurchmesser
            If Nwert >= Sys(Elt).E(2) Then
                Mother.H = Lang_Res(809) 'Innendurchmesser = Außendurchmesser gesetzt
                Nwert = Sys(Elt).E(2) - 10
                If Nwert < 0 Then Nwert = 0
                'ausendurchmesser nicht verändern, sonst sind weit mehr bestandteile betroffen
            End If
        Case 4 'scheibenbreite
            If Nwert < Sys(1).E(34) And Nwert > 0 Then 'ausschalten wird zugelassen'dann eben keine spitzenlast
                Mother.H = Lang_Res(810) & Sys(1).E(34) 'Bandbreite =
                Abbruch = True
            End If
        Case 9 'tragrollendurchmesser
            'muß größer innendurchmesser bleiben
            If Nwert > 10 Then
                If Nwert < Sys(Elt).E(96) + 4 Then '2023 aenderung, innendurchmesser anpassen
                    'Mother.H = Lang_Res(808) 'Außendurchmesser wurde an Innendurchmesser angepaßt'
                    Sys(Elt).E(96) = Nwert - 6
                End If
            End If
            If Sys(Elt).E(96) = 0 Then Sys(Elt).E(96) = Nwert - 6 'innendurchmesser ueberhaupt integrieren
            If Sys(Elt).E(96) < 0 Then Sys(Elt).E(96) = 0
            
            If Nwert > Sys(Elt).E(43) Then
                Sys(Elt).E(43) = Nwert
                If Sys(Elt).E(22) = 0 Then 'trägerlänge anpassen
                    'Sys(Elt).E(22) = Sys(Elt).E(43) * (nwert - 1)'bloß nicht, sonst ärger mit streckenlasten
                Else
                    '9 = tragrollendurchmesser
                    '22 = foerderlaenge
                    '43 = tragrollenachsabstand
                    '62 = Tragrollenanzahl
                    
                    M = Int(Sys(Elt).E(22) / Sys(Elt).E(43))
                    Sys(Elt).E(43) = Sys(Elt).E(22) / M 'nicht beliebige Achsabstände zulassen
                    Mother.H = Lang_Res(811)  'Tragrollenachsabstand wurde angepaßt
                    Sys(Elt).E(62) = CInt(Sys(Elt).E(22) / Sys(Elt).E(43) + 1) 'tragrollenanzahl korrigieren
                End If
            End If
            If Sys(Elt).E(10) = 0 Then Sys(Elt).E(10) = Nwert

        Case 10 'andruckrollendurchmesser
        Case 11 'eindringtiefe bei rollenbahnen
            If Sys(Elt).E(71) > 0 Then
                Sys(Elt).E(71) = 0
                Mother.H = Lang_Res(812) 'Andruckkraft bei federbelasteten Andruckrollen auf 0 gesetzt
            End If
        Case 13 'umschlingung
            Sys(Elt).E(13) = Nwert 'gibt keinen grund mehr, die eingabe zu sabotieren
        Case 16 'neigungswinkel von trägern, förderhöhe neu einstellen
            '22 = foerderlaenge
            '16 = Neigungswinkel
            '31 = forderhoehe
            Sys(Elt).E(31) = Sys(Elt).E(22) * Sin(Nwert * PI / 180)
            AnlageRefresh = True
        Case 17 'umfangskraft

            Sys(Elt).E(17) = Nwert 'neueingabe
            Call Interner_Abgleich(Elt, 17) 'nur sich selbst in ordnung bringen
        Case 18 'drehmoment
            Sys(Elt).E(18) = Nwert 'neueingabe
            Call Interner_Abgleich(Elt, 18)
        Case 19 'leistung
            Sys(Elt).E(19) = Nwert 'neueingabe
            Call Interner_Abgleich(Elt, 19)
        Case 20 'geschwindigkeit
            Sys(Elt).E(20) = Nwert 'neueingabe
            Call Geschwindigkeitsanpassung(Elt) 'alle anderen anpassen
        Case 21 'drehzahl aller elemente anpassen
            Sys(Elt).E(21) = Nwert 'neueingabe
            Call Interner_Abgleich(Elt, 21) 'sich selbst in ordnung bringen
            Call Geschwindigkeitsanpassung(Elt) 'alle anderen anpassen
        
        Case 22 'trägerlänge darf nur direkt und nicht durch eine andere eigenschaft geändert werden
            
            '22 = foerderlaenge
            '16 = Neigungswinkel
            '31 = forderhoehe
            
            'getragene elemente dürfen nicht beeinträchtigt werden
            For L = 10 To Maxelementindex
                If Sys(L).Zugehoerigkeit = Elt Then 'beide gehören zum gleichen träger
                    If Sys(L).E(25) > Nwert Or Sys(L).E(46) > Nwert Then
                        Abbruch = True
                        Mother.H = Lang_Res(845) 'entfernen Sie erst das rechte Element auf dem Träger
                        Exit Sub
                    End If
                End If
            Next L
            
            'förderhöhe neu berechnen
            If Sys(Elt).E(31) <> 0 Then
                Sys(Elt).E(31) = Sys(Elt).E(31) / Sys(Elt).E(22) * Nwert
            End If
            
            'Massen an konstante streckenlast anpassen
            For L = 9 To Maxelementindex
                If Sys(L).Zugehoerigkeit = Elt And Sys(L).Tag = "201" Then
                    Sys(L).E(28) = Sys(L).E(29) / 1000 * Nwert
                    Mother.H = Lang_Res(813)  'Transportmassen wurden angepaßt
                End If
            Next L
            
            'max. belegte förderlänge, betrifft nur rollenbahn
            If Sys(Eig).Tag = "103" Then
                If Sys(Eig).E(65) > Nwert Then Sys(Eig).E(65) = Nwert
            End If
            
            AnlageRefresh = True
            
            'noch tragrollenangaben korrigieren
                '9 = tragrollendurchmesser
                '22 = foerderlaenge
                '43 = tragrollenachsabstand
                '62 = Tragrollenanzahl

            If Sys(Elt).Tag = "103" Or Sys(Elt).Tag = "102" And Nwert > 0 Then
                If Sys(Elt).E(43) > 0 Then Sys(Elt).E(62) = CInt(Nwert / Sys(Elt).E(43) + 1) 'zuerst anzahl verändern, sonst ärger mit zu großem tragrollendurchmesser
                If Sys(Elt).E(62) > 1 Then Sys(Elt).E(43) = Nwert / (Sys(Elt).E(62) - 1)
                If Sys(Elt).E(9) > Sys(Elt).E(43) Then Sys(Elt).E(9) = Sys(Elt).E(43)
                If Nwert < Sys(Elt).E(65) Then Sys(Elt).E(65) = Nwert 'belegte länge notfalls anpassen
                'If nwert < Sys(Elt).E(68) Then Sys(Elt).E(68) = nwert 'staulänge notfalls anpassen
            End If
            
        Case 23 'staumasse muß kleiner der Gesamtmasse sein
            L = 0 'masse wird addiert
            For j = 1 To Maxelementindex
                If Sys(j).Tag = "201" And Sys(j).Zugehoerigkeit = Sys(Elt).Zugehoerigkeit Then L = L + Sys(j).E(28)
                If Sys(j).Tag = "204" And Sys(j).Zugehoerigkeit = Sys(Elt).Zugehoerigkeit And Elt <> j Then L = L - Sys(j).E(23) 'andere staumassen abziehen
            Next j
            If L < Nwert Then
                Mother.H = Lang_Res(814)  'soviel Transportgut haben Sie nicht definiert
                Abbruch = True
                Exit Sub
            End If
        Case 24 'freies trum zum element links gibts nicht mehr, frei
        Case 25 'position linkes ende
            AnlageRefresh = True
            M1 = Sys(Elt).E(25)
            M2 = Sys(Elt).E(46)
            Sys(Elt).E(25) = Nwert 'damit die Unterroutine funktioniert
            If Sys(Elt).E(25) > Sys(Elt).E(46) Then Sys(Elt).E(46) = Sys(Elt).E(25) 'bei abweisern immer
            Call Trägeraufteilung(Elt)
            If Abbruch = True Then 'abbruch, originalzustand wieder herstellen
                Sys(Elt).E(25) = M1 'wird unten wieder richtig gesetzt
                Sys(Elt).E(46) = M2 'wird unten wieder richtig gesetzt
            End If
        Case 28 'transportgutmasse
            If Nwert < Sys(Elt).E(32) Then
                Mother.H = Lang_Res(815) 'weniger Masse als ein Transportgutstück?
                Abbruch = True
                Exit Sub
            End If
            L = 0 'masse wird addiert, prüfung auf zuviel staumasse
            For j = 1 To Maxelementindex
                If Sys(j).Tag = "201" And Sys(j).Zugehoerigkeit = Sys(Elt).Zugehoerigkeit And Elt <> j Then L = L + Sys(j).E(28)
                If Sys(j).Tag = "204" And Sys(j).Zugehoerigkeit = Sys(Elt).Zugehoerigkeit Then L = L - Sys(j).E(23) 'andere staumassen abziehen
            Next j
            If L + Nwert < 0 Then
                Beep
                Mother.H = Lang_Res(816)  'setzen Sie erst die Staumasse herab
                Abbruch = True
                Exit Sub
            End If
            
            L = 0
            Do
                L = L + 1
            Loop Until Sys(Elt).Zugehoerigkeit = L
            If Sys(L).E(22) > 0 Then
                Sys(Elt).E(29) = Nwert * 1000 / Sys(L).E(22)
                Mother.H = Lang_Res(817) 'die Streckenlast wurde angepaßt
            End If
        Case 29 'streckenlast
            N = 0
            Do
                N = N + 1
            Loop Until Sys(Elt).Zugehoerigkeit = N
            If Sys(N).E(22) = 0 Then
                Abbruch = True
                Mother.H = Lang_Res(818)  'Geben Sie erst eine Förderlänge an
                Exit Sub
            End If
            L = 0
            For j = 1 To Maxelementindex
                If Sys(j).Tag = "201" And Sys(j).Zugehoerigkeit = Sys(Elt).Zugehoerigkeit And Elt <> j Then L = L + Sys(j).E(28)
                If Sys(j).Tag = "204" And Sys(j).Zugehoerigkeit = Sys(Elt).Zugehoerigkeit Then L = L - Sys(j).E(23) 'andere staumassen abziehen
            Next j
            If L + Sys(N).E(22) * Nwert / 1000 < 0 Then
                Mother.H = Lang_Res(819)  'setzen Sie erst die Staumasse herab
                Abbruch = True
                Exit Sub
            End If
            If Sys(N).E(22) > 0 Then 'Transportgutmasse noch verändern
                Sys(Elt).E(28) = Sys(N).E(22) * Nwert / 1000
                If Sys(Elt).E(32) > Sys(Elt).E(28) Then Sys(Elt).E(32) = Sys(Elt).E(28)
                Mother.H = Lang_Res(820)  'die Transportmasse wurde angepaßt
            End If
        Case 31 'förderhöhe
            '22 = foerderlaenge
            '16 = Neigungswinkel
            '31 = forderhoehe
            If Abs(Nwert) >= Sys(Elt).E(22) Then
                Abbruch = True
                Mother.H = Lang_Res(821)  'Förderhöhe darf nicht > Förderlänge sein
                Exit Sub
            End If
            AnlageRefresh = True
            
            'ersatz vom arcussinus,arcsin=(h/l)*180/pi
            Sys(Elt).E(16) = Atn((Nwert / Sys(Elt).E(22)) / Sqr(-(Nwert / Sys(Elt).E(22)) * (Nwert / Sys(Elt).E(22)) + 1)) * 180 / PI
        Case 32 'max masse eines Stückes
            If Nwert > Sys(Elt).E(28) Then
                Beep
                Mother.H = Lang_Res(814)  'So viel Transportgut haben Sie nicht definiert
                Abbruch = True
                Exit Sub
                'Sys(Elt).E(28) = nwert
            End If
        Case 33 'bandlänge gibts nicht mehr Call Bandmindestlängenberechnung(Eig)
        Case 34 'bandbreite weiterverarbeiten
            
            'und andere Scheiben verbreitern, wenn nötig
            L = 0
            Do
                L = L + 1
                M = 0 'die richtige elementzuordnung finden
                Do
                    M = M + 1
                Loop Until El(0).Eig(M) = Sys(L).Tag 'ruhig durchzählen, da alle elemente getestet werden
                If Sys(L).Element <> "" And Sys(L).Verb(1, 1) > 0 And El(4).Eig(M) <> "" Then
                    If Sys(L).E(4) < Nwert Or Init_B_Rex_Scheibenbreitefixieren = 1 Then
                        M1 = (2 * Init_B_Rex_ScheibenbreitenueberhangProz + 100) / 100
                        If M1 < 1 Then M1 = 1
                        Sys(L).E(4) = Nwert * M1
                        If (Sys(L).E(4) - Nwert) / 2 > Init_B_Rex_ScheibenbreitenUeberhangGrenze Then
                            Sys(L).E(4) = Nwert + 2 * Init_B_Rex_ScheibenbreitenUeberhangGrenze
                        End If
                    End If
                End If
            Loop Until L = Maxelementezahl
            
        Case 43 'tragrollenachsabstand
                '9 = tragrollendurchmesser
                '22 = foerderlaenge
                '43 = tragrollenachsabstand
                '62 = Tragrollenanzahl

            If Nwert > 0 And Nwert < Sys(Elt).E(9) Then
                Abbruch = True
                Mother.H = Lang_Res(823) & Sys(Elt).E(9)  'Tragrollendurchmesser =
                Exit Sub
            End If
            If Sys(Elt).E(22) > 0 And Nwert > 0 Then
                M = Int(Sys(Elt).E(22) / Nwert)
                If M > 0 Then
                    Nwert = Sys(Elt).E(22) / M 'nicht beliebige Achsabstände zulassen
                    Sys(Elt).E(62) = Sys(Elt).E(22) / Nwert + 1
                End If
            End If
            'If Sys(Elt).E(22) = 0 Then Sys(Elt).E(22) = (Sys(Elt).E(62) - 1) * Sys(Elt).E(43)
        Case 45 'messerkantendurchmesser
        Case 46 'abstand vom fördererende/rechts
            AnlageRefresh = True
            M1 = Sys(Elt).E(25)
            M2 = Sys(Elt).E(46)
            Sys(Elt).E(46) = Nwert 'damit die Unterroutine funktioniert
            If Sys(Elt).E(25) > Sys(Elt).E(46) Then Sys(Elt).E(25) = Sys(Elt).E(46) 'bei abweisern immer
            Call Trägeraufteilung(Elt)
            If Abbruch = True Then 'abbruch, originalzustand wieder herstellen
                Sys(Elt).E(25) = M1 'wird unten wieder richtig gesetzt
                Sys(Elt).E(46) = M2 'wird unten wieder richtig gesetzt
            End If
        Case 47 'staulänge rechts / unbenutzt
        Case 54 'spannkraft auf scheibe
            If Sys(1).E(1) <> 0 Then
                Sys(1).E(1) = 0
                Mother.H = Lang_Res(826)  'vorgegebene Auflegedehnung auf 0 gesetzt
            End If
            For N = 9 To Maxelementindex
                If Sys(N).E(54) > 0 And N <> Elt Then
                    Abbruch = True
                    Mother.H = Lang_Res(846)  'Nur eine Scheibe kann gewichts-/federbelastet sein.
                    Sys(Elt).E(54) = 0 'vorsichtshalber
                    Sys(Elt).E(55) = 0
                    Exit Sub
                End If
            Next N
            Sys(Elt).E(55) = Nwert / 9.81 'entsprechendes gewicht
            
            'winkel trum von rechts und von links gleichmäßig auf die umschlingung verteilen
            If Sys(Elt).E(56) = 0 And Sys(Elt).E(57) = 0 Then
                Sys(Elt).E(56) = 180 - (Sys(Elt).E(13) / 2 + 90) 'die 90 berücksichtigt die umkehr der orientierung
                If Sys(Elt).E(56) < 0 Then Sys(Elt).E(56) = 360 - Abs(Sys(Elt).E(56))
                Sys(Elt).E(57) = 180 + Sys(Elt).E(13) / 2 + 90
                If Sys(Elt).E(57) > 360 Then Sys(Elt).E(57) = Sys(Elt).E(57) - 360
            End If
        Case 55 'gewicht auf scheibe
            If Sys(1).E(1) <> 0 Then
                Sys(1).E(1) = 0
                Mother.H = Lang_Res(826) 'vorgegebene Auflegedehnung auf 0 gesetzt
            End If
            For N = 9 To Maxelementindex
                If Sys(N).E(54) > 0 And N <> Elt Then
                    Abbruch = True
                    Mother.H = Lang_Res(846)  'Nur eine Scheibe kann gewichts-/federbelastet sein.
                    Sys(Elt).E(54) = 0 'vorsichtshalber
                    Sys(Elt).E(55) = 0
                    Exit Sub
                End If
            Next N
            Sys(Elt).E(54) = Nwert * 9.81
            
            'winkel trum von rechts und von links gleichmäßig auf die umschlingung verteilen
            If Sys(Elt).E(56) = 0 And Sys(Elt).E(57) = 0 Then
                Sys(Elt).E(56) = 180 - (Sys(Elt).E(13) / 2 + 90)
                If Sys(Elt).E(56) < 0 Then Sys(Elt).E(56) = 360 - Abs(Sys(Elt).E(56))
                Sys(Elt).E(57) = 180 + Sys(Elt).E(13) / 2 + 90
                If Sys(Elt).E(57) > 360 Then Sys(Elt).E(57) = Sys(Elt).E(57) - 360
            End If
        Case 56 'winkel trum von links
        Case 57 ' winkel trum von rechts
        Case 59 'Überlastvorgabe manuell
            Sys(Elt).E(60) = 22 'überlastvorgabe durch auswahl löschen
            Mother.H = "automatische Überlast wurde auf 0 gesetzt" 'englisch
        Case 62 'tragrollenanzahl
            If Nwert > 0 Then
                Nwert = CInt(Nwert)
                If Sys(Elt).E(22) = 0 And Sys(Elt).E(43) > 0 Then Sys(Elt).E(22) = Sys(Elt).E(43) * (Nwert - 1)
                If Nwert > 1 Then
                    If Sys(Elt).E(22) / (Nwert - 1) < Sys(Elt).E(9) Then
                        Sys(Elt).E(9) = Sys(Elt).E(22) / (Nwert - 1)
                        Mother.H = Lang_Res(828)  'Tragrollendurchmesser mußte verändert werden
                    End If
                    Sys(Elt).E(43) = Sys(Elt).E(22) / (Nwert - 1)
                End If
            End If
        Case 63 'anz trag- pro andruckrollen
            Nwert = CInt(Nwert)
        Case 65 'max mit transportgut belegte Rollenbahnlänge
            If Nwert > Sys(Elt).E(22) Then
                Beep
                Mother.H = Lang_Res(829)  'belegte Länge kann nicht größer Förderlänge sein
                Abbruch = True
                Exit Sub
            End If
        Case 68 'beschleunigung
            If Nwert > 0 Then
                Sys(1).E(95) = Sys(1).E(20) / Nwert 't=v/a
            Else
                Sys(1).E(95) = 0
            End If
        Case 71 'Andruckkraft bei federbelasteten Andruckrollen
            If Sys(Elt).E(11) > 0 Then
                Sys(Elt).E(11) = 0
                Sys(Elt).E(12) = 0
                Mother.H = Lang_Res(830)  'Eindringtiefe der Andruck- in die Tragrollenreihe auf 0 gesetzt
            End If
        Case 73 'achsabstand von antriebsscheibe
            AnlageRefresh = True
        Case 74
            AnlageRefresh = True
        Case 82 'Fw
            SystemTyp.KraftdehnungMode = 4 'selbst gewaehlt
            SystemTyp.Kraftdehnung = Nwert 'ausnahmsweise wird hier ein zusaetzlicher wert befuellt, weil es mehrere quellen gibt
            Sys(1).E(117) = Nwert / 2 ' * 1.2
        Case 95 'hochlaufzeit
            If Nwert > 0 Then
                Sys(1).E(68) = Sys(1).E(20) / Nwert 'a=v/t
            Else
                Sys(1).E(68) = 0
            End If
        Case 96 'Tragrolleninnendurchmesser
            
            If Nwert > Sys(Elt).E(9) Then
                Mother.H = Lang_Res(809)  'Innendurchmesser = Außendurchmesser gesetzt
                Nwert = Sys(Elt).E(9) - 2
                If Nwert > 0 Then Nwert = 0
            End If
        Case 117 'k1 wert eingabe, fw-wert nachziehen und systemtyp.kraftdehnung
            Sys(1).E(82) = Nwert / 1.2 * 2 '1.2 von k1 nach sd, 2 von sd nach fw
            SystemTyp.Kraftdehnung = Sys(1).E(82)
            SystemTyp.KraftdehnungMode = 4

    End Select
    
    'bei allen, also auch bei denen ohne besondere behandlung weiter oben
        If Abbruch = False Then Sys(Elt).E(Eig) = Nwert
    
    'massenträgheitsmoment für alle ausrechnen
        Call Massenträgheitsmoment
    
    If Abbruch = False Then
        Call Bandmindestlängenberechnung(Eig) 'geht erst hier, weil sys(elt) mit neuem wert versehen ist
        Call CodeCalc.Rechnungssteuerung("VC") 'falls vollst element verändert wurde, wird anlagerefresh = true, also neudarstellung
'''        If AnlageRefresh = True Then Call CodeDraw.Alleelementeverbinden
        Call Dateiverwaltung.Undo(0) 'regelt undo und aktualität
    End If
    
    
End Sub
Public Sub Massenträgheitsmoment()
'der einfachheit für alle
Dim L, M As Integer
Dim Rechnen As Boolean
    L = 0
    Do
        L = L + 1
        Rechnen = True
        If Sys(L).Element <> "" Then 'vorhanden und einzelstehend
            If left(Sys(L).Tag, 1) = "0" Then 'einzelne scheiben
                If Sys(L).E(2) = 0 Then Rechnen = False 'durchmesser
                'innendurchmesser (3) brauchen wir nicht, 0 ist ok
                If Sys(L).E(4) = 0 Then Rechnen = False 'scheibenbreite
                If Sys(L).E(47) = 0 Then Rechnen = False 'materialhinweis
                'If Sys(L).E(114) > 0 Then Rechnen = False
                If Rechnen = True Then
                    Sys(L).E(5) = PI / 4 * (Sys(L).E(4) / 100) * ((Sys(L).E(2) / 100) ^ 2 - (Sys(L).E(3) / 100) ^ 2) 'volumen in dm^3 oder Liter, zwischenrechnung
                    Sys(L).E(5) = Sys(L).E(5) * Kst(Sys(L).E(47)).Einstellung  'm^3 * kg/dm^3 = kg! 'masse in kg
                    Sys(L).E(8) = Sys(L).E(5) * 0.5 * ((Sys(L).E(2) / 2000) ^ 2 + (Sys(L).E(3) / 2000) ^ 2) 'massenträgheitsmoment, Nms^2
                    Sys(L).E(99) = PI * (Sys(L).E(2) ^ 4 - Sys(L).E(3) ^ 4) / 64 'flächenmoment für durchbiegung
                Else
                    Sys(L).E(5) = 0 'masse = 0
                    Sys(L).E(8) = 0 'massenträgheitsmoment = 0
                    Sys(L).E(99) = 0 'flächenmoment = 0
                End If
                
                'es gibt ein manuelles traegheitsmoment, das bevorzugt behandelt werden muss
                    'If Sys(L).E(114) > 0 Then Sys(L).E(8) = 0
            End If
            If left(Sys(L).Tag, 1) = "1" Then 'rollen- und tragrollenbahnen, tisch fällt sowieso durchs rost
                If Sys(L).E(9) = 0 Then Rechnen = False 'Tragrollendurchmesser
                'innendurchmesser brauchen wir nicht, 0 ist ok
                If Sys(L).E(4) = 0 Then Rechnen = False 'scheibenbreite
                If Sys(L).E(47) = 0 Then Rechnen = False 'materialhinweis
                If Rechnen = True Then
                    Sys(L).E(5) = PI / 4 * (Sys(L).E(4) / 100) * ((Sys(L).E(9) / 100) ^ 2 - (Sys(L).E(96) / 100) ^ 2) 'volumen in dm^3 oder Liter
                    Sys(L).E(5) = Sys(L).E(5) * Kst(Sys(L).E(47)).Einstellung  'm^3 * kg/dm^3 = kg! 'masse in kg
                    'einer einzigen Scheibe!
                    Sys(L).E(8) = Sys(L).E(5) * 0.5 * ((Sys(L).E(9) / 1000) ^ 2 + (Sys(L).E(96) / 1000) ^ 2) 'massenträgheitsmoment, Nms^2
                Else
                    Sys(L).E(5) = 0 'masse = 0
                    Sys(L).E(8) = 0 'massenträgheitsmoment = 0
                End If
            End If

        End If
    Loop Until L = Maxelementindex

End Sub
Private Sub Geschwindigkeitsanpassung(ByVal Elt As Integer)
'Elt ist die Referenz, das gerade veränderte element, an das die anderen angepasst werden müssen
Dim L, M As Integer
    L = 0
    Do
        L = L + 1
        If Sys(L).Element <> "" And left(Sys(L).Tag, 1) <> "2" Then 'vorhanden und einzelstehend
            M = 7 'die richtige elementzuordnung finden
            Do
                M = M + 1
            Loop Until El(0).Eig(M) = Sys(L).Tag 'ruhig durchzählen, da alle elemente getestet werden
            
            Call Interner_Abgleich(L, 20)
            'sich selbst gleicht er gleich mit ab, das ändert ja auch nix
        End If
    Loop Until L = Maxelementindex
    
    'und dann wäre da noch die neue hochlaufzeit nach einer v-aenderung, a wird konstant gelassen
    If Sys(1).E(68) > 0 Then Sys(1).E(95) = Sys(1).E(20) / Sys(1).E(68) 't=v/a

End Sub
Private Sub Interner_Abgleich(ByVal L As Integer, ByVal Indecs As Integer)
    'element L soll an die neue geschw./Drehzahl angepaßt werden
    Select Case Indecs
        Case 2 'durchmesser
            If Sys(L).E(2) = 0 Then
                Sys(L).E(21) = 0
                Sys(L).E(18) = 0
            Else
                Sys(L).E(18) = Sys(L).E(17) * Sys(L).E(2) / 2000
                'wenn schon eine drehzahl da war, ist jetzt diese scheibe geschwindigkeitsbestimmend
                'statt wie früher eine drehzahl nur zuzulassen, wenn schon ein durchmesser da war
                'If Sys(L).E(21) > 0 Then
                '    Sys(1).E(20) = Sys(L).E(21) * Sys(L).E(2) * PI / 60000
                'Else
                '    Sys(L).E(21) = Sys(1).E(20) * 60000 / (PI * Sys(L).E(2))
                'End If
                
                'anderer durchmesser verändert drehzahl, nicht gleich die ganze bandgeschwindigkeit
                '21 drehzahl
                '20 v
                '2 durchmesser
                If Sys(1).E(20) > 0 Then
                    Sys(L).E(21) = Sys(1).E(20) * 60000 / (PI * Sys(L).E(2))
                Else
                    Sys(1).E(20) = Sys(L).E(21) * Sys(L).E(2) * PI / 60000
                End If

            End If
        Case 17 'kraft
            If Sys(L).E(17) = 0 Then
                Sys(L).E(18) = 0
                Sys(L).E(19) = 0
            Else
                Sys(L).E(18) = Sys(L).E(17) * Sys(L).E(2) / 2000
                Sys(L).E(19) = Sys(L).E(17) * Sys(1).E(20) / 1000
            End If
        Case 18 'drehmoment
            If Sys(L).E(2) = 0 Then
                Mother.H = Lang_Res(831)  'erst einen Durchmesser eingeben
                Abbruch = True
                Exit Sub
            End If
            If Sys(L).E(18) = 0 Then
                Sys(L).E(17) = 0
                Sys(L).E(19) = 0
            Else
                Sys(L).E(17) = Sys(L).E(18) / (Sys(L).E(2) / 2000)
                Sys(L).E(19) = Sys(L).E(17) * Sys(1).E(20) / 1000
            End If
        Case 19 'leistung
            If Sys(1).E(20) = 0 Then
                Mother.H = Lang_Res(832)  'erst eine Geschwindigkeit eingeben
                Abbruch = True
                Exit Sub
            End If
            If Sys(L).E(19) = 0 Then
                Sys(L).E(17) = 0
                Sys(L).E(18) = 0
            Else
                Sys(L).E(17) = Sys(L).E(19) * 1000 / Sys(1).E(20)
                Sys(L).E(18) = Sys(L).E(17) * Sys(L).E(2) / 2000
            End If
        Case 20 'bandgeschwindigkeit
            If Sys(1).E(20) = 0 Then
                Sys(L).E(21) = 0
                Sys(L).E(19) = 0
            Else
                If Sys(L).E(2) > 0 Then Sys(L).E(21) = Sys(1).E(20) * 60000 / (PI * Sys(L).E(2))
                Sys(L).E(19) = Sys(L).E(17) * Sys(1).E(20) / 1000
            End If
        Case 21 'drehzahl
            If Sys(L).E(21) = 0 Then
                Sys(1).E(20) = 0
                Sys(L).E(19) = 0
            Else
                Sys(1).E(20) = Sys(L).E(21) * Sys(L).E(2) * PI / 60000
                Sys(L).E(19) = Sys(L).E(17) * Sys(1).E(20) / 1000
            End If
    End Select
End Sub
Public Sub Bandmindestlängenberechnung(ByVal Eig As Integer)
    'indecs ist die eigenschaftsnummer zur letzten eingabe, nwert der eingegebene wert
    Dim N As Integer, L As Integer
    
    Call Zwei_Scheiben
    
    If Eig = 73 Then Achsabstand = Nwert 'nur bei 2-scheiben-lösung, bei allen anderen indecs, wird er oben übertragen
    
    If Zweischeiben = True Then 'dann isses auch extremultus
        'r1 immer systemnummer antriebsscheibe, r2 immer umlenkscheibe
        Call Bandlänge_Achsabstand(Eig)
        
        If Eig = 74 Then
            If Sys(1).E(74) = 0 Then Exit Sub
            'Achsabstand = 1000000 'muß von oben kommen
            N = 0
            Do
                N = N + 1
                Achsabstand = Achsabstand * Nwert / Sys(1).E(74)
                Call Bandlänge_Achsabstand(74)
            Loop Until (Sys(1).E(74) / Nwert > 0.999 And Sys(1).E(74) / Nwert < 1.001) Or N = 100
            If N > 999 Then
                Abbruch = True
                Mother.H = Lang_Res(833)  'Diese Bandlänge ist unmöglich
                Exit Sub
            End If
        End If
        
        Sys(1).E(33) = 0 'die unechte bandlänge ist dann ja wohl obsolet
        
        If Eig = 33 Then
            Nwert = 0 'alle anderen daten werden so in ordnung gebracht
            Mother.H = Lang_Res(834)  '2-Scheiben-Lösungen besitzen nur eine exakte Bandlänge
        End If
    Else
        'grobe bandlängenberechnung
        If Eig = 74 Or Eig = 73 Then 'tats. bandlängenberechnung
            Abbruch = True
            Mother.H = Lang_Res(835)  'Angabe nur bei 2 Scheiben - Antrieben möglich
            Exit Sub
        End If
        
        Sys(1).E(74) = 0 'tatsächliche bandlänge gibts nur bei 2-Scheiben
        Sys(1).E(33) = 0 'zurückstellen, denn sie wird neu gebildet
        
        For L = 9 To Maxelementindex
            'freie trümer zwischen den elementen (immer nur zum höherwertigen, sonst doppeltzählung)
            If Sys(L).Verb(1, 1) > L Then Sys(1).E(33) = Sys(1).E(33) + Sys(L).Verb(1, 3)
            If Sys(L).Verb(2, 1) > L Then Sys(1).E(33) = Sys(1).E(33) + Sys(L).Verb(2, 3)
            
            'und die anteile an den elementen:
            
            'förderer
            If left(Sys(L).Tag, 1) = "1" Then
                Sys(1).E(33) = Sys(1).E(33) + Sys(L).E(22) 'förderlänge
            End If
            
            'Einzelstehend
            '33 = bandlaenge
            '45 = Messerkantendurchmesser
            '2 = Durchmesser
            '13 = umschlingungswinkel
            
            If left(Sys(L).Tag, 1) = "0" Then
                If left(Sys(L).Tag, 3) = "005" Then 'Messerkante
                   Sys(1).E(33) = Sys(1).E(33) + PI * Sys(L).E(45) * Sys(L).E(13) / 360 'durchmesser
                Else
                   Sys(1).E(33) = Sys(1).E(33) + PI * Sys(L).E(2) * Sys(L).E(13) / 360 'durchmesser
                End If
            End If
        Next L
    End If
End Sub
Private Sub Bandlänge_Achsabstand(ByVal Eig As Integer)
Dim B As Double
Dim i As Integer

    'achsabstand mind radius+radius?
    If Achsabstand < Sys(R1).E(2) / 2 + Sys(R2).E(2) / 2 Then Achsabstand = Sys(R1).E(2) / 2 + Sys(R2).E(2) / 2 + 0.1
    Sys(R2).E(73) = Achsabstand
    If Eig = 73 Then Nwert = Achsabstand
    '73 = achsabstand von der antriebsscheibe, 2 = durchmesser
    'laenge durch satz von pythagoras, radien voneinander abziehen
    Sys(R1).E(24) = Sqr(Abs(Sys(R2).E(73) ^ 2 - (Sys(R2).E(2) / 2 - Sys(R1).E(2) / 2) ^ 2))
    Sys(R2).E(24) = Sys(R1).E(24) 'freie trümer zum element links
    Sys(1).E(74) = 2 * Sys(R1).E(24)
    
    Sys(R1).Verb(1, 3) = Sys(R1).E(24) 'die neuen achsabstände
    Sys(R1).Verb(2, 3) = Sys(R1).E(24)
    Sys(R2).Verb(1, 3) = Sys(R1).E(24)
    Sys(R2).Verb(2, 3) = Sys(R1).E(24)
    
    'winkel
    If Abs(Sys(R2).E(2) - Sys(R1).E(2) = 0) Or Sys(R2).E(73) = 0 Then
        B = 0
    Else
        B = Abs(Sys(R2).E(2) / 2 - Sys(R1).E(2) / 2) / Sys(R2).E(73) 'bruch in handlicher form für die nächste zeile
        If B >= 0.9999 Then B = 0.9999
        B = Atn(B / Sqr(-B * B + 1)) 'ersatz für arkussinus, ergebnis im bogenmaß
    End If
    If Sys(R1).E(2) > Sys(R2).E(2) Then
        Sys(R1).E(13) = (PI + 2 * B) * 180 / PI
        Sys(R2).E(13) = (PI - 2 * B) * 180 / PI
    Else
        Sys(R1).E(13) = (PI - 2 * B) * 180 / PI
        Sys(R2).E(13) = (PI + 2 * B) * 180 / PI
    End If
    
    '54 = spannkraft auf scheibe
    '13 = umschlingung
    If Sys(R1).E(54) > 0 Then 'neue Winkel
        Sys(R1).E(56) = 180 + Sys(R1).E(13)  ' - 180
        If Sys(R1).E(56) > 360 Then Sys(R1).E(56) = Sys(R1).E(56) - 360
        Sys(R1).E(57) = 180 - Sys(R1).E(13)  ' - 180
        If Sys(R1).E(57) < 0 Then Sys(R1).E(57) = Sys(R1).E(57) + 360
    End If
    If Sys(R2).E(54) > 0 Then
        Sys(R2).E(56) = 180 + Sys(R2).E(13)  ' - 180
        If Sys(R2).E(56) > 360 Then Sys(R2).E(56) = Sys(R2).E(56) - 360
        Sys(R2).E(57) = 180 - Sys(R2).E(13)  ' - 180
        If Sys(R2).E(57) < 0 Then Sys(R2).E(57) = Sys(R2).E(57) + 360
    End If
    
    
    'geaendert 11_05_2005, frueher radius inkl Banddicke (79)
    'Sys(1).E(74) = Sys(1).E(74) + PI * (Sys(R1).E(2) + Sys(1).E(79)) * Sys(R1).E(13) / 360 + PI * (Sys(R2).E(2) + Sys(1).E(79)) * Sys(R2).E(13) / 360
    'heute ohne, also bandinnenseite
    Sys(1).E(74) = Sys(1).E(74) + PI * Sys(R1).E(2) * Sys(R1).E(13) / 360 + PI * Sys(R2).E(2) * Sys(R2).E(13) / 360
    
    
    If Eig = 13 Then
        Abbruch = True
        Beep
        Nwert = Sys(i).E(13)
        Mother.H = Lang_Res(840)  'mit Achsabstand und Durchmesser auch Winkel definiert
    End If
End Sub
Public Sub Zwei_Scheiben()
Dim N As Integer

'rausfinden, obs eine extremultus anlage mit nur zwei scheiben ist
    'r1 enthält systemnumer der antriebsscheibe, r2 die der umlenkscheibe bei 2 elementen
    Achsabstand = 0 'wird in der umlenkscheibe und in dieser variable bevorratet
    UmlScheibe = False
    AntrScheibe = False
    Zweischeiben = True
    N = 9
    Do
        N = N + 1
            If Sys(N).Tag = "001" Then
                AntrScheibe = True
                'mehr als eine kann ja nicht da sein
                R1 = N
            End If
            If Sys(N).Tag = "003" Then
                If UmlScheibe = True Then
                    Zweischeiben = False
                Else
                    UmlScheibe = True
                End If
                Achsabstand = Sys(N).E(73)
                R2 = N
            End If
        If Sys(N).Tag <> "" And Sys(N).Tag <> "001" And Sys(N).Tag <> "003" Then Zweischeiben = False 'dann ists irgendein anderes element aus der fördertechnik
    Loop Until N > Maxelementindex Or Zweischeiben = False
    If AntrScheibe = False Or UmlScheibe = False Then Zweischeiben = False 'emperie
    
'und die eigenschaften der beiden beteiligten scheiben entsprechend hinbiegen
    
    'Band
    N = 6 'die richtige elementzuordnung für antriebsscheibe finden
    Do
        N = N + 1
    Loop Until El(0).Eig(N) = "301" 'ruhig durchzählen, da alle elemente getestet werden
    If Zweischeiben = True Then 'alles einrichten für extremultus
        El(33).Eig(N) = 0 'ca. Bandlänge ausschalten
        El(74).Eig(N) = 2 'tats. Bandlänge ins kann
    Else
        El(33).Eig(N) = 3 'ca. Bandlänge ansehbar
        El(74).Eig(N) = 0 'tats. Bandlänge ausschalten
    End If
    
    'Antriebsscheibe
        N = 7 'die richtige elementzuordnung für antriebsscheibe finden
        Do
            N = N + 1
        Loop Until El(0).Eig(N) = "001" 'ruhig durchzählen, da alle elemente getestet werden
        If Zweischeiben = True Then 'alles einrichten für extremultus
            El(21).Eig(N) = Replace(El(21).Eig(N), "2", "1") 'drehzahl
        Else
            El(21).Eig(N) = Replace(El(21).Eig(N), "1", "2") 'drehzahl
        End If
    
    'umlenkscheibe
        N = 7 'die richtige elementzuordnung für umlenkscheibe finden
        Do
            N = N + 1
        Loop Until El(0).Eig(N) = "003" 'ruhig durchzählen, da alle elemente getestet werden
        If Zweischeiben = True Then 'alles einrichten für extremultus
            El(73).Eig(N) = 1 'achsabstand ins muß
            El(21).Eig(N) = Replace(El(21).Eig(N), "2", "1") 'drehzahl
            El(17).Eig(N) = Replace(El(17).Eig(N), "2", "1") 'Umfangskraft
            El(18).Eig(N) = Replace(El(18).Eig(N), "2", "1") 'Drehmoment
            El(19).Eig(N) = Replace(El(19).Eig(N), "2", "1") 'Leistung
        Else 'alles einrichten für transilon
            El(73).Eig(N) = 0 'achsabstand aus
            El(21).Eig(N) = Replace(El(21).Eig(N), "1", "2") 'drehzahl
            El(17).Eig(N) = Replace(El(17).Eig(N), "1", "2") 'Umfangskraft
            El(18).Eig(N) = Replace(El(18).Eig(N), "1", "2") 'Drehmoment
            El(19).Eig(N) = Replace(El(19).Eig(N), "1", "2") 'Leistung
        End If
End Sub
Public Sub Trägeraufteilung(ByVal Elt As Integer)
    Dim L As Integer
    Dim N, P As Double
    
    Abbruch = False
    If Sys(Elt).Tag = "201" Then Exit Sub 'transportgut aufhalten, falls es sich bis hierhin geschummelt hat
    
    N = Sys(Elt).E(25) 'linke grenze
    P = Sys(Elt).E(46) 'rechte grenze
    
    If Sys(Elt).E(46) > Sys(Sys(Elt).Zugehoerigkeit).E(22) Or Sys(Elt).E(25) > Sys(Sys(Elt).Zugehoerigkeit).E(22) Then
        Abbruch = True
        Mother.H = Lang_Res(841)  'Förderer zu kurz / an dieser Stelle belegt
        Exit Sub
    End If
    
    For L = 9 To Maxelementindex
        If Sys(L).Element <> "" And Sys(L).Zugehoerigkeit <> 0 Then
            If Sys(L).Zugehoerigkeit = Sys(Elt).Zugehoerigkeit And L <> Elt Then 'beide gehören zum gleichen träger
                If Sys(L).Tag <> "201" Then
                    If Sys(L).E(25) >= N And Sys(L).E(25) <= P Then Abbruch = True 'l/links inmitten elt
                    If Sys(L).E(46) >= N And Sys(L).E(46) <= P Then Abbruch = True 'l/rechts inmitten elt
                    If Sys(L).E(25) <= N And Sys(L).E(46) >= N Then Abbruch = True 'l verdeckt elt/links
                    If Sys(L).E(25) <= P And Sys(L).E(46) >= P Then Abbruch = True 'l verdeckt elt/rechts
                    If Abbruch = True Then
                        Mother.H = Lang_Res(843)  'Staus / Staubereiche / Abweiser würden sich überschneiden
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next L
End Sub

