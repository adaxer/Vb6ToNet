Imports System.Diagnostics.Metrics
Imports System.Runtime.Intrinsics.X86

Module CodeCalc
    'anlagenparameter, nur einmal durchführen
    Private Messervorhanden As Boolean 'hinweis zur richtigen bandspannung an messerbandanlagen
    Private Antriebsscheibe As Integer 'enthält nummer der antriebsscheibe, die jetzt ja irgendwo sein kann
    Private Startelement As Integer 'erstes element in richtung ablaufendes trum, legt rechenreihenfolge fest
    Private ScheibeFedGew As Integer 'diese eine scheibe hat eine feder/gewichtsbelastung
    Private FwScheibeFedGew As Double 'und diese wellenbelastung hat sie nach der rechnung, daraufhin iteration
    Private ScheibeFedGewNormalFu As Double 'verschiebung der normallast bei Feder/gewicht
    Private ScheibeFedGewSpitzeFu As Double 'verschiebung der spitzenlast bei Feder/gewicht
    Private maxFoerdererlaenge As Double 'für eine fehlermeldung zur plausiblen bandlänge
    Private Bandfuehrungvorhanden As Boolean 'um ein übereinanderlaufen des bandes zu verhindern
    Private Spitzenlastvorhanden As Boolean

    Private Auflegemodus As Integer 'feder, vorgabe, oder durch B_Rex

    Private LetzteBerechnung As Boolean

    'extremwerte der anlage
    Private Fumin As Double
    Private FuminSp As Double
    Private Fumax As Double 'hat immer den höchsten der Fu-Normalkurve parat
    Private FumaxSp As Double 'ebenso bei spitzenlasten

    'fehlerverwaltung'erst in string, damit schnell
    Private Fehler$ 'enthält die fehlertexte, werden zum schluß in die liste geschrieben (variable statt objekt = zeitgewinn)
    Private Datenfehler$
    Private SchwLongFehler$
    Private SchwTransFehler$

    Private Rechengenauigkeit As Integer 'max 40
    Private Fehlerverlauf(4, 42) As Double
    '1 fehlerpunkte,
    '2 bei der kraft,
    '3 und der dehnung,
    '4 fehlerpunkte in sachen schwingungen

    'rechnungssteuerung:
    Private Zeichnen As Boolean

    'verwaltung der bandparameter ab datenbank
    Private MinTrumKraft As Single 'enthält die min. aufltrumkraft
    Private MaxTrumKraft As Single 'enthält die max. dehn./aufltrumkraft je nach einestellung von:
    Private Dehnung$ 'enthält den text

    Private AuflTrumkraftSp As Double 'enthält die im letzten durchlauf ermittelte auflegetrumkraft bei spitzenlast
    Private AuflTK_Sp_N_Diff As Double 'differenz zwischen dem mittel fuer Trumkraft im Normal und im spitzenlastzustand aus letzten durchlauf

    Private DurchschFaktor As Double
    Private Staumasse As Double
    Private Staulänge As Double
    Private Fu As Double
    Private Fuletztes As Double
    Private FuletztesSp As Double
    Private Fuerstes As Double

    Private Schwingungen_berechnen As Boolean

    Private Fusteig As Double 'für stau
    Private Fusteig1 As Double 'für trägergeb. umfangskraft
    Private FuFwSpitze As Double



    Public Sub Rechnungssteuerung(Mode As String)
        'instr:
        'E = endlosprüfung
        'V = vollstaendig alle?
        'B = vollstaendig band?
        'C = neuauslegung

        Dim parameters = ReadBrexDump("C:\temp\vb6dump.json")
        CompareWithParameters(parameters, "C:\temp\comp.txt")

        Dim M As Double
        Dim i As Integer, P As Integer, j As Integer, K As Integer

        Dim Masse As Double, Memo As Double, H As Double, Errfreq As Double
        Dim FuerstesMerk As Double
        Dim mue As Double, OptDehn(2, 2) As Double 'fehlerwerte,position
        Dim Datenaenderung$ = String.Empty



        Fehlerwert = -100 'heisst so viel wie nicht drueber nachgedacht, weil e noch angaben fehlen


        'Endlospruefung
        If Mode.Contains("E") Then
            Endlos = False
            K = 9
            Do
                K = K + 1
            Loop Until Sys(K).Element <> "" Or K > Maxelementindex
            j = K 'erstes element merken
            If Sys(K).Verb(1, 1) > 0 Then
                P = 1
            Else
                P = 2
            End If
            Do
                i = K ' altes element merken
                K = Sys(K).Verb(P, 1) 'und nächstes bestimmen
                If Sys(K).Verb(1, 1) = i Then 'voreinstellungen für neuen durchlauf
                    P = 2
                Else
                    P = 1
                End If
            Loop Until K = j Or Sys(K).Verb(1, 1) = 0 Or Sys(K).Verb(2, 1) = 0
            If K = j Then Endlos = True  'einmal ohne unterbrechnung rum
        End If

        'Vollstaendigkeitskontrolle
        If Mode.Contains("V") Or Mode.Contains("B") Then
            Dim ElVollst As Boolean
            Vollstaendig = True
            K = Maxelementindex
            If Mode.Contains("B") Then K = 1 'nur das band
            For i = 1 To K 'elemente zählen
                If i = 1 Or i > 9 Then
                    If Sys(i).Element <> "" Then  '2 speichert die aus der datenbank übernommenen bandeigenschaften
                        ElVollst = True

                        M = Elementnummer(Sys(i).Tag)

                        'wenns als transportgut auf einem tisch liegt
                        If Sys(i).Tag = "201" Then
                            If Sys(Sys(i).Zugehoerigkeit).Tag = "101" Then
                                'sind nicht soviele infos erforderlich
                                El(32).Eig(M) = 0
                                El(36).Eig(M) = 0
                            Else 'rollenbahnen und tragrollenbahnen aber schon
                                El(32).Eig(M) = 1
                                El(36).Eig(M) = 1
                            End If
                        End If

                        P = 0
                        Do
                            P = P + 1
                            If El(P).Eig(M).Contains("1") Then
                                If CDbl(Sys(i).E(P)) = 0 Then
                                    ElVollst = False
                                    j = i
                                    If Sys(i).Zugehoerigkeit > 0 Then j = Sys(i).Zugehoerigkeit 'huckepacks auch nur erfassen, wenn sie zur anlage gehören
                                    If Sys(j).Verb(1, 1) <> 0 Or Sys(j).Verb(2, 1) <> 0 Or i = 1 Then 'nur wenn's auch am band ist, ist anlage unVollstaendig
                                        'vielleicht gehört eines der elemente nur nicht zur anlage
                                        Vollstaendig = False 'gilt natürlich auch fürs ganze system
                                    End If
                                End If
                            End If
                        Loop Until P = Eigenschaftszahl Or ElVollst = False 'eigenschaften zählen
                        'If I = 1 Then Stop
                        If Sys(i).Vollstaendig <> ElVollst Then AnlageRefresh = True 'neudarstellung erzwingen
                        Sys(i).Vollstaendig = ElVollst 'und festhalten
                    End If
                End If
            Next i
        End If

        If Vollstaendig = False Or Endlos = False Then  'ruft sich wohl nicht selbst auf
            If B_Rex_AutoLauf = False Then
                '''        Mother.Statusverwaltung(0)
                '''If B_Rex.FuKurve.Visible = True Then B_Rex.FuKurve.Cls 'Visible = False 'fukurve raus
                '''If B_Rex.Fehlerliste.Visible = True Then B_Rex.Fehlerliste = "" '.Visible = False  'fehlerliste raus
            End If
            Exit Sub
        End If

        If Not Mode.Contains("C") Then Exit Sub 'auftrag auch so schon erledigt


        'grundeinstellungen
        Rechengenauigkeit = 40
        If B_Rex_AutoLauf = True Then Rechengenauigkeit = 15 'spart rechenzeit beim autolauf, sollte trotzdem reichen
        Fumax = 0
        FumaxSp = 0
        Zeichnen = False
        Fuerstes = 0
        ScheibeFedGew = 0
        FwScheibeFedGew = 0
        maxFoerdererlaenge = 0
        Bandfuehrungvorhanden = False
        Spitzenlastvorhanden = False

        'Schätzungen inbezug auf das band werden unter query verlagert

        'biegeleistungskennwert schätzen, aber erstmal keinen fehler anzeigen
        If Sys(1).E(80) = 0 Then Sys(1).E(80) = Sys(1).E(79) / 33 + (Sys(1).E(79) / 10) ^ 2

        'bandgewicht wird nicht erzwungen, falls nötig, aber geschätzt
        If Sys(1).E(81) <= 0 Then
            Datenfehler$ = Lang_Res(616) & Environment.NewLine  '-Bandgewicht geschätzt
            If Sys(1).E(81) = 0 Then Sys(1).E(81) = Sys(1).E(79)
        End If

        'das gesamte bandgewicht ermitteln und wenn nötig, die kraft, das band zu beschleunigen
        If Sys(1).E(33) = 0 Then 'extremultus
            H = Sys(1).E(74) 'nach tats. Bandlänge
        Else 'transilon
            H = Sys(1).E(33) 'nach der ca. Bandlänge
        End If
        Sys(1).E(30) = (H / 1000 * Sys(1).E(34) / 1000) * Sys(1).E(81) 'nach der ca. Bandlänge
        Sys(1).E(98) = Sys(1).E(30) * Sys(1).E(68) 'F = m*a
        Sys(1).FusteigSp = Sys(1).E(98) / H 'diese steigung durch die beschleunigung der Bandmasse in der spitzenlastkurve entlang des gesamten Bandes

        'maximale dehnung oder max. auflegedehnung
        'max. aufldehn. ist überkommener unsinn, daher wenn möglich max. dehn. nehmen
        Dehnung$ = ""
        MaxTrumKraft = 0
        If Sys(1).E(84) > 0 Then 'max. auflegedehnung, am liebsten diese nicht
            Dehnung$ = Lang_Res(646)  'max. zul. Auflegedehn.
            MaxTrumKraft = Abs(Sys(1).E(84) / 2 * SystemTyp.Kraftdehnung * Sys(1).E(34))
            Fuerstes = Abs(Sys(1).E(84) / 7) * Abs(SystemTyp.Kraftdehnung / 2) * (Sys(1).E(34)) 'Dehnung im Leertrum
            'MaxTrumKraft enthält eine Kraft
        End If
        If Sys(1).E(85) > 0 Then 'maxdehn, schon viel besser, wenn möglich hier
            Dehnung$ = Lang_Res(647)  'max. zul. Dehnung
            MaxTrumKraft = Abs(Sys(1).E(85) / 2 * SystemTyp.Kraftdehnung * Sys(1).E(34))
            Fuerstes = Abs(Sys(1).E(85) / 11) * Abs(SystemTyp.Kraftdehnung / 2) * (Sys(1).E(34))  'Dehnung im Leertrum
        End If
        If Dehnung$ = "" Then Fehler$ = Fehler$ & Lang_Res(671) & Environment.NewLine  'Keine Angabe 'maximal zulässige Dehnung' gefunden
        'hier die grundeinstellung für den Rechnungsstart
        If Fuerstes = 0 Then Fuerstes = 0.2 * Abs(SystemTyp.Kraftdehnung / 2) * (Sys(1).E(34)) 'bevor garnichts drin steht, starten wir eben mit 0.2

        'minimale trumkraft
        MinTrumKraft = Abs(Sys(1).E(94) / 2 * SystemTyp.Kraftdehnung * Sys(1).E(34)) 'F = Fw*e*bo/2 '2 eigentlich nicht, ist nur die umrechnung pro trum

        If Sys(1).E(68) > 0 Then Spitzenlastvorhanden = True


        'paar grundsätzliche feststellungen, müssen nur einmal pro anlage gemacht werden
        For M = 9 To Maxelementindex

            'richtung festlegen (erstes element ist antrscheibe, in richtung des ablaufenden trums das zweite, dann der reihe nach)
            If Sys(M).Tag = "001" Then 'ist erster Durchlauf (Antriebsscheibe) und geschlossen

                'Startposition der rechnung
                Antriebsscheibe = M
                i = Sys(M).Verb(1, 2) 'anschluß festhalten

                'startrichtung (in richtung ablaufendes trum)
                '4 möglichkeiten
                If Reversieren = False Then
                    If i = 1 Or i = 3 Or i = 5 Or i = 7 Then 'eben die ins leertrum ablaufenden teile
                        Startelement = Sys(M).Verb(1, 1)
                    Else
                        Startelement = Sys(M).Verb(2, 1)
                    End If
                Else
                    If i = 2 Or i = 4 Or i = 6 Or i = 8 Then
                        Startelement = Sys(M).Verb(1, 1)
                    Else
                        Startelement = Sys(M).Verb(2, 1)
                    End If
                End If

            End If

            'fliehkraftbehandlung/Wölbhöhe/Spitzenlast
            If Sys(M).E(21) > 0 And Sys(M).E(2) > 0 Then 'drehzahl, durchmesser

                'gesamte bandmasse rund um die scheibe
                Masse = PI * Sys(M).E(2) / 1000 * Sys(1).E(34) / 1000 * Sys(1).E(81)

                'gesamte fliehkraft, um die scheibe ent- und Band belastet wird, jedes trum zur hälfte dieses wertes
                '0.4 gedachter anteil der überhaupt zum tragen kommenden masse'schätzwert, vorsicht, zweimal pflegen
                Sys(M).E(51) = 2 * PI ^ 2 * (Sys(M).E(21) / 60) ^ 2 * (Sys(M).E(2) / 1000) * Masse * 0.4 * (Sin(((Sys(M).E(13) * PI / 180) / 2 - PI / 2)) + 1) / 2
                'Fliehkraftsumme = Fliehkraftsumme + Sys(M).E(51) / 2 'jetzt in nur einem trum!

                'wölbhöhe empfehlen
                Sys(M).E(7) = 0
                'abgeschaltet 201910, das war ne eigene formel
                Select Case LCase(Sys(1).S(5))

                    Case "extremultus"
                        'die elastischen etwas anders
                        If SystemTyp.Name.Contains("0U") And SystemTyp.Name.Contains("FDA") Then 'die elastischen 20U 40U 60U + FDA ist Hinweis auf elastisch
                            If Sys(M).E(2) > 1000 Then Sys(M).E(7) = 1.2
                            If Sys(M).E(2) < 1000 Then Sys(M).E(7) = 1
                            If Sys(M).E(2) < 600 Then Sys(M).E(7) = 0.6
                            If Sys(M).E(2) < 300 Then Sys(M).E(7) = 0.5
                            If Sys(M).E(2) < 200 Then Sys(M).E(7) = 0.4
                        Else
                            '201910 ab jetzt weiter mit norm ISO 22 fuer extremultus
                            If Sys(1).E(34) > 250 Then
                                Sys(M).E(7) = 2.5
                                If Sys(M).E(2) <= 1500 Then Sys(M).E(7) = 2 'eigentlich 1600, aber was, wenn groesser?
                                If Sys(M).E(2) <= 1120 Then Sys(M).E(7) = 1.5
                            Else
                                Sys(M).E(7) = 1.8
                                If Sys(M).E(2) <= 1500 Then Sys(M).E(7) = 1.5 'eigentlich 1600, aber was, wenn groesser?
                                If Sys(M).E(2) <= 1120 Then Sys(M).E(7) = 1.2
                            End If
                            If Sys(M).E(2) <= 1120 Then Sys(M).E(7) = 1.2
                            If Sys(M).E(2) <= 800 Then Sys(M).E(7) = 1.2
                            If Sys(M).E(2) <= 560 Then Sys(M).E(7) = 1
                            If Sys(M).E(2) <= 315 Then Sys(M).E(7) = 0.8
                            If Sys(M).E(2) <= 250 Then Sys(M).E(7) = 0.6
                            If Sys(M).E(2) <= 200 Then Sys(M).E(7) = 0.7
                            If Sys(M).E(2) <= 160 Then Sys(M).E(7) = 0.4
                            If Sys(M).E(2) <= 125 Then Sys(M).E(7) = 0.3
                        End If
                    Case "transilon"
                        If LCase(Left(SystemTyp.Name, 2)) = "el" Then 'die elastischen
                            If Sys(M).E(2) > 1000 Then Sys(M).E(7) = 1.2
                            If Sys(M).E(2) < 1000 Then Sys(M).E(7) = 1
                            If Sys(M).E(2) < 600 Then Sys(M).E(7) = 0.6
                            If Sys(M).E(2) < 300 Then Sys(M).E(7) = 0.5
                            If Sys(M).E(2) < 200 Then Sys(M).E(7) = 0.4
                        Else
                            If SystemTyp.Name.Contains("/1") Then 'einlagig
                                Sys(M).E(7) = 1
                                If Sys(M).E(2) <= 500 Then Sys(M).E(7) = 0.8
                                If Sys(M).E(2) <= 200 Then Sys(M).E(7) = 0.5
                            End If
                            If SystemTyp.Name.Contains("/2") Or SystemTyp.Name.Contains("/M") Or SystemTyp.Name.Contains("NOVO") Then
                                Sys(M).E(7) = 1.5
                                If Sys(M).E(2) <= 500 Then Sys(M).E(7) = 1.3
                                If Sys(M).E(2) <= 200 Then Sys(M).E(7) = 0.7
                            End If
                            If SystemTyp.Name.Contains("/3") Then
                                Sys(M).E(7) = 2
                                If Sys(M).E(2) <= 500 Then Sys(M).E(7) = 1.6
                                If Sys(M).E(2) <= 200 Then Sys(M).E(7) = 1
                            End If
                        End If
                    Case Else
                        'auch hier koennte man noch einer norm folgen, unter knowhow, woelbhoehe
                        If SystemTyp.Kraftdehnung > 0 Then Sys(M).E(7) = (Sys(M).E(2) ^ 0.27 + 0.393) / (15.7 / SystemTyp.Kraftdehnung + PI) 'eigene formel
                End Select
                'nix drin, dann eigene formel
                If Sys(M).E(7) = 0 Then If SystemTyp.Kraftdehnung > 0 Then Sys(M).E(7) = (Sys(M).E(2) ^ 0.27 + 0.393) / (15.7 / SystemTyp.Kraftdehnung + PI)


                'wölbhöhe vorhanden?
                If Sys(M).E(106) > 0 Then Bandfuehrungvorhanden = True

                'spitzenlast?
                If Sys(M).E(60) <> 22 And Sys(M).E(60) <> 0 Then Spitzenlastvorhanden = True
                If Sys(M).E(59) > 0 Then Spitzenlastvorhanden = True

                'fehlen kennwerte?
                If Sys(M).E(109) = 0 Then Sys(M).E(109) = 1 'faktor von funenn zw. 1 und 1,25
                If Sys(M).E(110) = 0 Then Sys(M).E(110) = 1 'funenn bei dieser dehnung angegeben

            End If

            'fliehkraftbehandlung messer
            If Sys(M).Tag = "005" And Sys(M).E(45) > 0 Then 'messerkante

                'hinweis zur ermittlung der richtigen bandspannung
                Messervorhanden = True

                'imaginäre messerkantendrehzahl, wird ja keine eingegeben
                Sys(M).E(21) = Sys(1).E(20) * 60000 / (PI * Sys(M).E(45))

                'sonst wie unter scheiben
                Masse = PI * Sys(M).E(45) / 1000 * Sys(1).E(34) / 1000 * Sys(1).E(81)
                Sys(M).E(51) = 2 * PI ^ 2 * (Sys(M).E(21) / 60) ^ 2 * (Sys(M).E(45) / 1000) * Masse * 0.4 * (Sin(((Sys(M).E(13) * PI / 180) / 2 - PI / 2)) + 1) / 2
                'Fliehkraftsumme = Fliehkraftsumme + Sys(M).E(51) / 2

            End If

            'eine scheibe gewichts- oder federbelastet?
            'auflegedehnungsvorgabe steht immer in sys(1).e(1)
            If Sys(M).E(54) > 0 Then
                ScheibeFedGew = M
            End If

            If Left(Sys(M).Tag, 1) = "1" Then
                If maxFoerdererlaenge < Sys(M).E(22) Then maxFoerdererlaenge = Sys(M).E(22)
            End If

            'alten ergebnisse der schwingungen loeschen
            Sys(M).Verb(1, 4) = 0
            Sys(M).Verb(2, 4) = 0
            Sys(M).E(112) = 0
            Sys(M).E(113) = 0
        Next M

        Schwingungen_berechnen = False
        If Init_B_Rex_Schw_alle = 1 Then Schwingungen_berechnen = True
        If Init_B_Rex_Schw_nur_Ex = 1 And Zweischeiben = True Then Schwingungen_berechnen = True

        SchwTransFehler$ = ""
        If Schwingungen_berechnen = True Then
            If Sys(1).E(20) < 10 Then
                Schwingungen_berechnen = False
                SchwTransFehler$ = "- no vibration calculation up to 10 m/s" & vbCrLf
            End If
        End If


        'longitudinalschwingungsberechnung
        'ist ein anlagenspezifischer zustand, hat nichts mit dehnung und geschwindigkeit zu tun
        'also reicht pro auslegung eine einmalige betrachtung hier an dieser stelle
        'gibts bloss bei zweischeiben
        Sys(1).E(116) = 0 'systemeigenfrequenz longitudinal
        FehlerwertLongSchwing = 0
        SchwLongFehler$ = ""


        If 1 = 2 Then 'Schwingungen_berechnen = True Then 'longitudinalschwingungen zum ausprobieren eingeschaltet
            Dim Federkonstante As Double
            '''            Dim Federkonstante2 As Double
            Dim KleinerRadius As Double
            Dim MTMotorseite As Double 'kombiniert aus scheibe und motor, teils autom. oder manuell
            Dim MTreduziert As Double
            Dim MTscheibe As Double

            '202003 angeblich das hier einsetzen, was ich fuer unfug halte, ist band individuel
            ' E-dyn. PA  = 4300 N/mm^2
            ' E-dyn.PES = 15000 N/mm^2
            ' E-dyn.a = 89000 N/mm^2'aramide


            'erstmal ein zwischenwert, fw * bo/(0,02 * freie laenge)
            'eines trums, gleich der des anderen trums entgegen landläufiger meinung
            Federkonstante = SystemTyp.Kraftdehnung * Sys(1).E(34) / (0.02 * (Sys(Startelement).Verb(1, 3) / 1000))
            'federkonstante des einen trums bei zweischeiben gleich der des anderen

            'weiter zur systemfederkonstante
            KleinerRadius = Sys(Startelement).E(2) / 2000
            If Sys(Antriebsscheibe).E(2) / 2000 < KleinerRadius Then KleinerRadius = Sys(Antriebsscheibe).E(2) / 2000
            Federkonstante = 2 * Federkonstante * KleinerRadius ^ 2

            'reduziertes massentraegheitsmoment der scheibe
            'i = durchmessergetrieben/durchmessertreibend
            '21 = drehzahl
            '8 = errechnetes massentraegheitsmoment
            MTscheibe = Sys(Startelement).E(114) 'manuelles scheibe gegeben?
            If MTscheibe = 0 Then MTscheibe = MassentraegheitsErmittlung(Startelement) 'dann automatisches scheibe bevorzugen
            MTscheibe = MTscheibe + Sys(Startelement).E(115) 'kommt noch das der maschine hinzu
            If Sys(Startelement).E(21) > 0 And Sys(Antriebsscheibe).E(21) > 0 Then
                MTreduziert = MTscheibe * (Sys(Antriebsscheibe).E(2) / Sys(Startelement).E(2)) ^ 2
            End If
            'wenn nichts eingegeben wurde, dann eben das errechnete vom system

            'massentraegheitsmoment motorseite
            MTMotorseite = Sys(Startelement).E(114) 'manuelles scheibe gegeben?
            If MTMotorseite = 0 Then MTMotorseite = MassentraegheitsErmittlung(Startelement) 'dann automatisches scheibe bevorzugen
            MTMotorseite = MTMotorseite + Sys(Startelement).E(115) 'kommt noch das der maschine hinzu

            'die eigenfrequenz des systems ermitteln
            If MTMotorseite > 0 And MTreduziert > 0 Then
                Sys(1).E(116) = 1 / (2 * PI) * Sqr(Federkonstante * (MTMotorseite + MTreduziert) / (MTMotorseite * MTreduziert))
            End If

            'entgültige fehlersuche und einstufung
            If Sys(Antriebsscheibe).E(111) > 0 Then
                Errfreq = Sys(Antriebsscheibe).E(111) * Sys(Antriebsscheibe).E(21) / 60
                If Errfreq > 0.8 * Sys(1).E(116) And Errfreq < 1.2 * Sys(1).E(116) Then
                    'ganz nah dran auf 20%
                    SchwLongFehler$ = SchwLongFehler$ & "-Antriebsscheibe: kritische Eigenfrequenzanregung longitudinal" & vbCrLf 'unterschreitet mindestdurchmesser
                    FehlerwertLongSchwing = FehlerwertLongSchwing + 100
                Else
                    If Errfreq > 0.7 * Sys(1).E(116) And Errfreq < 1.3 * Sys(1).E(116) Then
                        'in der naehe, 30%
                        SchwLongFehler$ = SchwLongFehler$ & "-Antriebsscheibe: Eigenfrequenzanregung longitudinal" & vbCrLf 'unterschreitet mindestdurchmesser
                        FehlerwertLongSchwing = FehlerwertLongSchwing + 10
                    End If
                End If
            End If

            If Sys(Startelement).E(111) > 0 Then
                Errfreq = Sys(Startelement).E(111) * Sys(Startelement).E(21) / 60
                If Errfreq > 0.8 * Sys(1).E(116) And Errfreq < 1.2 * Sys(1).E(116) Then
                    'ganz nah dran auf 20%
                    SchwLongFehler$ = SchwLongFehler$ & "-getriebene Scheibe: kritische Eigenfrequenzanregung longitudinal" & vbCrLf 'unterschreitet mindestdurchmesser
                    FehlerwertLongSchwing = FehlerwertLongSchwing + 100
                Else
                    If Errfreq > 0.7 * Sys(1).E(116) And Errfreq < 1.3 * Sys(1).E(116) Then
                        'in der naehe, 30%
                        SchwLongFehler$ = SchwLongFehler$ & "-getriebene Scheibe: Eigenfrequenzanregung longitudinal" & vbCrLf 'unterschreitet mindestdurchmesser
                        FehlerwertLongSchwing = FehlerwertLongSchwing + 10
                    End If
                End If
            End If
        End If



        'fliehkraft hebt das gesamte niveau, nicht das einzelner elemente
        'fliehkraftsummme bezieht sich auf ein trum

        FuerstesMerk = Fuerstes 'falls was schiefgeht, eine passende antwort
        Auflegemodus = 1 'normal
        'iteration gewaehlte auflegedehnung
        'üblich wären höchstens 5 iterationen
        If Sys(1).E(1) > 0 And ScheibeFedGew <= 0 Then
            Berechnung() 'einmal, damit sys(1).e(53) <> 0 ist
            If Sys(1).E(53) = 0 Then Sys(1).E(53) = 0.5 'nur zur sicherheit
            i = 0
            Do
                i = i + 1
                'neue einst = letzte einst * anzust. Wert / IstWert
                Fuerstes = Fuerstes * Sys(1).E(1) / Sys(1).E(53)
                Berechnung()
            Loop Until (Sys(1).E(1) / Sys(1).E(53) > 0.99 And Sys(1).E(1) / Sys(1).E(53) < 1.01) Or i > 50
            If i > 49 Then 'notbremse, läuft ins nirvana
                Fuerstes = FuerstesMerk 'zurückstellen
                '''Mother.H = Lang_Res(672)  'Bei der vorgegebenen Auflegedehnung wäre keine Dehnung im Leertrum
                'also raus damit
                Sys(1).E(1) = 0
                Berechnung()
            Else
                Auflegemodus = 3 'durch auflegedehungsvorgabe
                FuerstesMerk = Fuerstes 'da müssen wir hin, user will es so
            End If
        End If





        'iteration feder-/gewichtsbelastete scheibe
        'es können schon mal 30 iterationen werden
        If ScheibeFedGew > 0 Then
            'Sys(ScheibeFedGew).E(54) da soll er hin
            'FwScheibeFedGew das hier hat die letzte rechnung ergeben

            Berechnung() 'sonst unten nulldivision
            i = 0
            Do
                i = i + 1
                'neue einst = letzte einst * anzust. Wert / IstWert
                If FwScheibeFedGew >= 0 Then Fuerstes = Fuerstes * Sys(ScheibeFedGew).E(54) / FwScheibeFedGew
                Berechnung()
                'e(54) = spannkraft auf scheibe
            Loop Until (Sys(ScheibeFedGew).E(54) / FwScheibeFedGew > 0.95 And Sys(ScheibeFedGew).E(54) / FwScheibeFedGew < 1.05) Or i > 100
            'Berechnung
            If i > 99 Then 'notbremse, läuft ins nirvana
                Fuerstes = FuerstesMerk 'zurückstellen
                '''Mother.H = Lang_Res(694)  'Gewichte/Federn der Spannscheibe reichen nicht und bleiben unberücksichtigt
                Berechnung()
            Else
                Auflegemodus = 4 ' durch feder/gewicht
                FuerstesMerk = Fuerstes 'da müssen wir hin, anlage will es so
            End If
        End If

        'wenn noch keine berechnung war, dann jetzt eine zur orientierung
        If Fumax = 0 Then Berechnung

        'jetzt zur farblichen anzeige noch n paar spielereien
        Memo = 1.5 * MaxTrumKraft / Rechengenauigkeit 'maximalangabe ist immer die beste
        If Memo = 0 Then Memo = 2 * Fumax / Rechengenauigkeit 'sonst einfach den erkenntnissen aus dem ersten durchlauf folgen

        OptDehn(1, 1) = 32000 'grundeinstellung viel zu hoch
        j = 0
        Do
            'bei messerkanten können kleine memo-unterschiede riesige stufen bei aufltrumkraft hervorrufen
            'dann mechanismus einbauen, der kleine stufen wählt
            'die farbliche kennzeichnung der Fukurve stimmt dann nicht mehr, weil nur eine stufe von 40, die ampel unten aber schon
            Fuerstes = 0 + j * Memo
            Berechnung()
            Kontrollrechnungen(K, M, mue, Errfreq) 'um die beurteilung komplett zu machen
            If AuflTrumKraft < 0 Then AuflTrumKraft = 0
            Fehlerverlauf(1, j) = AuflTrumKraft 'erf. auflegedehnung als kraft
            Fehlerverlauf(2, j) = Fehlerwert 'und der dazugehörige fehlerwert
            Fehlerverlauf(3, j) = Sys(1).E(53) 'AuflTrumKraft * 2 / (Systemty.Kraftdehnung * Sys(1).E(34)) 'entspr Dehnung mitprotokollieren
            Fehlerverlauf(4, j) = FehlerwertSchwingungen + FehlerwertLongSchwing 'und der dazugehörige fehlerwert
            '***wahrscheinlich ist ein grossteil unabhängig von der dehnung, daher müssen wesentliche teile der berechnung garnicht immer wiederholt werden, das checken


            'anfang des optimalen bereiches(idealerweise grüner bereich)
            If OptDehn(1, 1) > Fehlerwert And Fehlerwert < 100 Then 'aus rot muß er schon raus sein
                'immer, wenn ein noch besserer bereich gefunden wird, alles bisherige kippen:
                OptDehn(1, 1) = Fehlerwert 'untere begrenzung
                OptDehn(1, 2) = Fuerstes
                OptDehn(2, 1) = Fehlerwert 'obere begrenzung
                OptDehn(2, 2) = Fuerstes
            End If
            'ausdehnung des optimalen bereiches, optdehn(2,2) markiert obere grenze
            If OptDehn(1, 1) = Fehlerwert Then OptDehn(2, 2) = Fuerstes
            j = j + 1
        Loop Until j = Rechengenauigkeit 'Or (BO = True And Fehlerwert >= 100) 'zweiter roter bereich mus nicht weiter untersucht werden

        'entgültige entscheidung, wo's hingeht
        If Auflegemodus = 1 Then  'keine vorgaben, comp wählt aus
            If OptDehn(1, 1) < 100 Then 'optimaler bereich, also die dehnung genau mittenrein setzen
                Auflegemodus = 2
                H = MaxTrumKraft
                If Dehnung$ = Lang_Res(646) Then H = H * 1.4  'max. zul. Auflegedehn.
                'ein zehntel der maximalen dehnung in den grünen bereich hinein, nicht mehr, um die anlage nicht zu belasten
                If OptDehn(1, 2) + H / 15 < (OptDehn(1, 2) + OptDehn(2, 2)) / 2 And H > 0 Then
                    Fuerstes = OptDehn(1, 2) + H / 15
                Else 'geht nicht, also mittenrein
                    Fuerstes = (OptDehn(1, 2) + OptDehn(2, 2)) / 2
                End If
            Else 'kein optimaler bereich, also kann er machen, was er will, das band ist e mist
                Fuerstes = FuerstesMerk
            End If
        Else 'vorgaben von user und anlage ansteuern, die oben ermittelt wurden
            Fuerstes = FuerstesMerk
        End If

        'und entgültige rechnung
        Berechnung 'ab in die entgültige position

        'liegt die auswahl in der nähe eines roten bereichs?
        j = 0
        Do
            i = 0
            If Fehlerverlauf(2, j) >= 100 Then i = 1
            If Fehlerverlauf(2, j + 1) >= 100 Then i = i + 2
            If i = 1 Then 'dann ist es der untere rand
                'oftmals gibt's unten keine roten werte, hinweis
                If Sys(1).E(53) > Fehlerverlauf(3, j) And Sys(1).E(53) < Fehlerverlauf(3, j + 2) Then
                    Fehler$ = Fehler$ & Lang_Res(709) & Environment.NewLine  '"- Hinweis: Ihre Anlage enthält kaum Sicherheit."
                End If
            End If
            If i = 2 And j > 0 Then 'dann ist es der obere rand
                If Sys(1).E(53) > Fehlerverlauf(3, j - 1) And Sys(1).E(53) < Fehlerverlauf(3, j + 1) Then
                    Fehler$ = Fehler$ & Lang_Res(709) & Environment.NewLine  '"- Hinweis: Ihre Anlage enthält kaum Sicherheit."
                End If
            End If
            j = j + 1
        Loop Until j = Rechengenauigkeit



        'so, die rechnung stimmt, jetzt die restlichen aufgaben einfach erledigen

        'n paar bandeigenschaften, die eigentlich keiner wissen will:
        Sys(1).E(91) = Fumax
        Sys(1).E(92) = Fumin
        Sys(1).E(89) = Fumax * 2 / (SystemTyp.Kraftdehnung * Sys(1).E(34))
        Sys(1).E(90) = Fumin * 2 / (SystemTyp.Kraftdehnung * Sys(1).E(34))
        Sys(1).E(87) = Sys(1).E(89) - Sys(1).E(90)
        'der von der fliehkraft verursachte anteil an der dehnung

        If B_Rex_AutoLauf = False Then 'B_Rex.FuKurve.Visible = True
            '''Destination = B_Rex.FuKurve
            LetzteBerechnung = True
            ''' If Auflegemodus = 4 Then Grafik(False) 'feder/gewicht, spitzenlast braucht die info zum einregeln, einmal mehr ausführen
            '''Destination.Cls 'hier wird immer fukurve aufgebaut, nicht der printer oder sie seitenvorschau, also cls hier
            '''   Grafik(False)
            LetzteBerechnung = False
            'dort wird zwar auch die auflegedehnung neu gemacht, aber mit zeichnen und genau
            'den werten des letzten berechnungsdurchlaufs oben
            'wichtig für versatz bei feder/gewicht oder richtiges legen der spitzenlastkurve

        End If

        If AuflTrumKraft = 0 Then AuflTrumKraft = 0.0001 'sonst gefahr nulldivision

        Kontrollrechnungen(K, M, mue, Errfreq) 'noch fehlerwert und texte der gewählten konf. ermitteln

        'datenänderung ist immer leer, wenn das programm hier ankommt
        'änderungen an den Banddaten? Nicht durchgehen lassen, für meinen persönlichen Schlaf
        'der zweite teil lautet immer:...wurde gegüber der datenbank verändert
        If Abs(Sys(1).E(77)) <> Abs(Sys(2).E(77)) Then Datenaenderung$ = Datenaenderung$ & Lang_Res(683) & Lang_Res(682) & Environment.NewLine
        If Abs(Sys(1).E(78)) <> Abs(Sys(2).E(78)) Then Datenaenderung$ = Datenaenderung$ & Lang_Res(684) & Lang_Res(682) & Environment.NewLine
        If Abs(Sys(1).E(79)) <> Abs(Sys(2).E(79)) Then Datenaenderung$ = Datenaenderung$ & Lang_Res(685) & Lang_Res(682) & Environment.NewLine
        If Abs(Sys(1).E(81)) <> Abs(Sys(2).E(81)) Then Datenaenderung$ = Datenaenderung$ & Lang_Res(686) & Lang_Res(682) & Environment.NewLine

        If Left(Sys(1).S(2), 1) = "9" Then 'transilon

            Select Case SystemTyp.KraftdehnungMode
                Case 4 'selbstgewaehlt
                    Datenaenderung$ = Datenaenderung$ & "- force-stretch-value choosed by user" & vbCrLf 'fw-wert wurde veraendert
            End Select
        Else 'extremultus
            If Abs(Sys(1).E(83)) <> Abs(Sys(2).E(83)) Then Datenaenderung$ = Datenaenderung$ & Lang_Res(688) & Lang_Res(682) & Environment.NewLine  'sd-wert wurde veraendert (geht garnicht)
        End If
        If Abs(Sys(1).E(84)) <> Abs(Sys(2).E(84)) Then Datenaenderung$ = Datenaenderung$ & Lang_Res(689) & Lang_Res(682) & Environment.NewLine
        If Abs(Sys(1).E(85)) <> Abs(Sys(2).E(85)) Then Datenaenderung$ = Datenaenderung$ & Lang_Res(690) & Lang_Res(682) & Environment.NewLine
        If Abs(Sys(1).E(86)) <> Abs(Sys(2).E(86)) Then Datenaenderung$ = Datenaenderung$ & Lang_Res(691) & Lang_Res(682) & Environment.NewLine

        'ca bandlänge nicht plausibel, bandlänge reicht nicht für den rückweg
        If maxFoerdererlaenge > Sys(1).E(33) / 2 Then
            Fehler$ = Fehler$ & Lang_Res(695) & Environment.NewLine  '-In Ihrer Anlage fehlen Trumlängen (Trum anklicken, Länge eingeben, >enter< druecken)
        End If

        'gibts da n element, das garnicht mit der anlage verbunden ist?
        For K = 9 To Maxelementindex
            If Sys(K).Element <> "" And (Left(Sys(K).Tag, 1) = "0" Or Left(Sys(K).Tag, 1) = "1") Then
                If Sys(K).Verb(1, 1) = 0 Or Sys(K).Verb(2, 1) = 0 Then
                    Fehler$ = Fehler$ & Lang_Res(710) & K & Lang_Res(711) & Environment.NewLine   '- hinweis: element...ist nicht korrekt in ihre anlage eingebunden
                End If
            End If
        Next K




        'fehlerverwaltung und anzeige
        '''If B_Rex_AutoLauf = False Then
        '''    B_Rex.Fehlerliste = Datenfehler$ & Datenaenderung$ & Fehler$ & SchwTransFehler$ & SchwLongFehler$
        '''End If
        Exit Sub
    End Sub

    Private Sub Kontrollrechnungen(k As Integer, m As Double, mue As Double, errfreq As Double)
        'todo: ...'
    End Sub

    Private Sub Berechnung()
        'todo: ...'
    End Sub

    Function MassentraegheitsErmittlung(Element As Integer) As Double
        'manuelle Alterung ueberschreibt errechnete, neu 202006
        '8 errechnet
        '114 manuell
        MassentraegheitsErmittlung = Sys(Element).E(8)
        If Sys(Element).E(114) > 0 Then MassentraegheitsErmittlung = Sys(Element).E(114)
    End Function


End Module
