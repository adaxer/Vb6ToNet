Attribute VB_Name = "CodeCalc"
Option Explicit

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

'fehlerwerte
'10  Fw/Fu ev. Probleme bei der Kraftübertragung
'100 band hebt aufgrund der fliehkraft von der scheibe ab
'10  rho-wert > 0,08
'100 durchmesser reicht zur übertragung der umfangskraft nicht aus
'100 scheibe unterschreitet mindestdurchmesser
'100 keine restspannung im Leertrum
'100 keine auflegedehnung
'10  auflegedehnung etwas überschritten
'100 auflegedehnung deutlich überschritten
'100 maximale dehnung überschreitet max zul dehn, ermittelt aus auflegedehnung
'10  max dehnung überschreitet max zul dehn etwas
'100 max dehnung überschreitet max zul dehn deutlich
'100 messerkante untershreitet mindestdurchmesser
'100 messerkantentemp > 200
'***deaktiviert 10  transportgut könnte rutschen
'***deaktiviert 100 transportgut wird rutschen (bei Profilen (Elevator) eben nicht)
'10  abweiser läßt wahrscheinlich kein transportgut durch
'100 tragrollendurchmesser < mindestdurchmesser
'100 andruckrollendurchmesser < mindestdurchmesser
'100 band berührt tragrollen nicht
'10  transportgut könnte auf der rollenbahn rutschen
'100 transportgut wird auf der rollenbahn rutschen
'100 andruckkraft band-tragrolle reicht für mitnahme nicht aus
'10 messerbänder mir einer auflegedehnung unterhalb von 0,2% können schrumpfen

'Spitzenlastberechnung
    'die beschleunigung der bandmasse wird erst bei Aussgabe der Zeichnung auf jeden bandteil dazuaddiert
    'sys(x).e(98) enthält die zusätzliche Kraft zur Beschleunigung dieses elements
    'sys(x).FuSteigSp die dadurch verursachte steigung bei der fu-kurve,
        'beim band gültig für die ganze kurve
        'bei förderern bezogen auf nicht gestaute masse auf nicht gestauter Länge
    'sys(x).FuSteigSpRoll bei Rollen/Tragrollenbahnen die durch all die zu beschleunigenden Rollen zusätzliche Kraft
    
Public Sub Rechnungssteuerung(Mode As String)
'instr:
'E = endlosprüfung
'V = vollstaendig alle?
'B = vollstaendig band?
'C = neuauslegung

Dim M As Double
Dim i As Integer, P As Integer, j As Integer, K As Integer

Dim Masse As Double, Memo As Double, H As Double, Errfreq As Double
Dim FuerstesMerk As Double
Dim mue As Double, OptDehn(2, 2) As Double 'fehlerwerte,position
Dim Datenaenderung$
Dim BO As Boolean



Fehlerwert = -100 'heisst so viel wie nicht drueber nachgedacht, weil e noch angaben fehlen


'Endlospruefung
    If InStr(Mode, "E") > 0 Then
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
    If InStr(Mode, "V") > 0 Or InStr(Mode, "B") > 0 Then
        Dim ElVollst As Boolean
        Vollstaendig = True
        K = Maxelementindex
        If InStr(Mode, "B") > 0 Then K = 1 'nur das band
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
                        If InStr(El(P).Eig(M), "1") > 0 Then
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
'''        Call Mother.Statusverwaltung(0)
        If B_Rex.FuKurve.Visible = True Then B_Rex.FuKurve.Cls 'Visible = False 'fukurve raus
        If B_Rex.Fehlerliste.Visible = True Then B_Rex.Fehlerliste = "" '.Visible = False  'fehlerliste raus
    End If
    Exit Sub
End If

If InStr(Mode, "C") = 0 Then Exit Sub 'auftrag auch so schon erledigt


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
        Datenfehler$ = Lang_Res(616) & Chr$(13) & Chr$(10)  '-Bandgewicht geschätzt
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
    If Dehnung$ = "" Then Fehler$ = Fehler$ & Lang_Res(671) & Chr$(13) & Chr$(10)  'Keine Angabe 'maximal zulässige Dehnung' gefunden
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
                    If InStr(SystemTyp.Name, "0U") > 0 And InStr(SystemTyp.Name, "FDA") > 0 Then 'die elastischen 20U 40U 60U + FDA ist Hinweis auf elastisch
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
                    If LCase(left(SystemTyp.Name, 2)) = "el" Then 'die elastischen
                        If Sys(M).E(2) > 1000 Then Sys(M).E(7) = 1.2
                        If Sys(M).E(2) < 1000 Then Sys(M).E(7) = 1
                        If Sys(M).E(2) < 600 Then Sys(M).E(7) = 0.6
                        If Sys(M).E(2) < 300 Then Sys(M).E(7) = 0.5
                        If Sys(M).E(2) < 200 Then Sys(M).E(7) = 0.4
                    Else
                        If InStr(SystemTyp.Name, "/1") > 0 Then 'einlagig
                            Sys(M).E(7) = 1
                            If Sys(M).E(2) <= 500 Then Sys(M).E(7) = 0.8
                            If Sys(M).E(2) <= 200 Then Sys(M).E(7) = 0.5
                        End If
                        If InStr(SystemTyp.Name, "/2") > 0 Or InStr(SystemTyp.Name, "/M") > 0 Or InStr(SystemTyp.Name, "NOVO") > 0 Then
                            Sys(M).E(7) = 1.5
                            If Sys(M).E(2) <= 500 Then Sys(M).E(7) = 1.3
                            If Sys(M).E(2) <= 200 Then Sys(M).E(7) = 0.7
                        End If
                        If InStr(SystemTyp.Name, "/3") > 0 Then
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

        If left(Sys(M).Tag, 1) = "1" Then
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
            Call Berechnung 'einmal, damit sys(1).e(53) <> 0 ist
            If Sys(1).E(53) = 0 Then Sys(1).E(53) = 0.5 'nur zur sicherheit
            i = 0
            Do
                i = i + 1
                'neue einst = letzte einst * anzust. Wert / IstWert
                Fuerstes = Fuerstes * Sys(1).E(1) / Sys(1).E(53)
                Call Berechnung
            Loop Until (Sys(1).E(1) / Sys(1).E(53) > 0.99 And Sys(1).E(1) / Sys(1).E(53) < 1.01) Or i > 50
            If i > 49 Then 'notbremse, läuft ins nirvana
                Fuerstes = FuerstesMerk 'zurückstellen
                Mother.H = Lang_Res(672)  'Bei der vorgegebenen Auflegedehnung wäre keine Dehnung im Leertrum
                'also raus damit
                Sys(1).E(1) = 0
                Call Berechnung
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
            
            Call Berechnung 'sonst unten nulldivision
            i = 0
            Do
                i = i + 1
                'neue einst = letzte einst * anzust. Wert / IstWert
                If FwScheibeFedGew >= 0 Then Fuerstes = Fuerstes * Sys(ScheibeFedGew).E(54) / FwScheibeFedGew
                Call Berechnung
                'e(54) = spannkraft auf scheibe
            Loop Until (Sys(ScheibeFedGew).E(54) / FwScheibeFedGew > 0.95 And Sys(ScheibeFedGew).E(54) / FwScheibeFedGew < 1.05) Or i > 100
            'Call Berechnung
            If i > 99 Then 'notbremse, läuft ins nirvana
                Fuerstes = FuerstesMerk 'zurückstellen
                Mother.H = Lang_Res(694)  'Gewichte/Federn der Spannscheibe reichen nicht und bleiben unberücksichtigt
                Call Berechnung
            Else
                Auflegemodus = 4 ' durch feder/gewicht
                FuerstesMerk = Fuerstes 'da müssen wir hin, anlage will es so
            End If
        End If
    
    'wenn noch keine berechnung war, dann jetzt eine zur orientierung
    If Fumax = 0 Then Call Berechnung
    
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
            Call Berechnung
            Call Kontrollrechnungen(K, M, mue, Errfreq) 'um die beurteilung komplett zu machen
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
        Call Berechnung 'ab in die entgültige position
    
    'liegt die auswahl in der nähe eines roten bereichs?
        j = 0
        Do
            i = 0
            If Fehlerverlauf(2, j) >= 100 Then i = 1
            If Fehlerverlauf(2, j + 1) >= 100 Then i = i + 2
            If i = 1 Then 'dann ist es der untere rand
                'oftmals gibt's unten keine roten werte, hinweis
                If Sys(1).E(53) > Fehlerverlauf(3, j) And Sys(1).E(53) < Fehlerverlauf(3, j + 2) Then
                    Fehler$ = Fehler$ & Lang_Res(709) & Chr$(13) & Chr$(10)  '"- Hinweis: Ihre Anlage enthält kaum Sicherheit."
                End If
            End If
            If i = 2 And j > 0 Then 'dann ist es der obere rand
                If Sys(1).E(53) > Fehlerverlauf(3, j - 1) And Sys(1).E(53) < Fehlerverlauf(3, j + 1) Then
                    Fehler$ = Fehler$ & Lang_Res(709) & Chr$(13) & Chr$(10)  '"- Hinweis: Ihre Anlage enthält kaum Sicherheit."
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
        Set Destination = B_Rex.FuKurve
        LetzteBerechnung = True
        If Auflegemodus = 4 Then Call Grafik(False) 'feder/gewicht, spitzenlast braucht die info zum einregeln, einmal mehr ausführen
        Destination.Cls 'hier wird immer fukurve aufgebaut, nicht der printer oder sie seitenvorschau, also cls hier
        Call Grafik(False)
        LetzteBerechnung = False
        'dort wird zwar auch die auflegedehnung neu gemacht, aber mit zeichnen und genau
        'den werten des letzten berechnungsdurchlaufs oben
        'wichtig für versatz bei feder/gewicht oder richtiges legen der spitzenlastkurve
        
    End If
    
    If AuflTrumKraft = 0 Then AuflTrumKraft = 0.0001 'sonst gefahr nulldivision
    
    Call Kontrollrechnungen(K, M, mue, Errfreq) 'noch fehlerwert und texte der gewählten konf. ermitteln
        
    'datenänderung ist immer leer, wenn das programm hier ankommt
    'änderungen an den Banddaten? Nicht durchgehen lassen, für meinen persönlichen Schlaf
    'der zweite teil lautet immer:...wurde gegüber der datenbank verändert
    If Abs(Sys(1).E(77)) <> Abs(Sys(2).E(77)) Then Datenaenderung$ = Datenaenderung$ & Lang_Res(683) & Lang_Res(682) & Chr$(13) & Chr$(10)
    If Abs(Sys(1).E(78)) <> Abs(Sys(2).E(78)) Then Datenaenderung$ = Datenaenderung$ & Lang_Res(684) & Lang_Res(682) & Chr$(13) & Chr$(10)
    If Abs(Sys(1).E(79)) <> Abs(Sys(2).E(79)) Then Datenaenderung$ = Datenaenderung$ & Lang_Res(685) & Lang_Res(682) & Chr$(13) & Chr$(10)
    If Abs(Sys(1).E(81)) <> Abs(Sys(2).E(81)) Then Datenaenderung$ = Datenaenderung$ & Lang_Res(686) & Lang_Res(682) & Chr$(13) & Chr$(10)
    
    If left(Sys(1).S(2), 1) = "9" Then 'transilon
        
        Select Case SystemTyp.KraftdehnungMode
            Case 4 'selbstgewaehlt
                Datenaenderung$ = Datenaenderung$ & "- force-stretch-value choosed by user" & vbCrLf 'fw-wert wurde veraendert
        End Select
    Else 'extremultus
        If Abs(Sys(1).E(83)) <> Abs(Sys(2).E(83)) Then Datenaenderung$ = Datenaenderung$ & Lang_Res(688) & Lang_Res(682) & Chr$(13) & Chr$(10)  'sd-wert wurde veraendert (geht garnicht)
    End If
    If Abs(Sys(1).E(84)) <> Abs(Sys(2).E(84)) Then Datenaenderung$ = Datenaenderung$ & Lang_Res(689) & Lang_Res(682) & Chr$(13) & Chr$(10)
    If Abs(Sys(1).E(85)) <> Abs(Sys(2).E(85)) Then Datenaenderung$ = Datenaenderung$ & Lang_Res(690) & Lang_Res(682) & Chr$(13) & Chr$(10)
    If Abs(Sys(1).E(86)) <> Abs(Sys(2).E(86)) Then Datenaenderung$ = Datenaenderung$ & Lang_Res(691) & Lang_Res(682) & Chr$(13) & Chr$(10)
    
    'ca bandlänge nicht plausibel, bandlänge reicht nicht für den rückweg
    If maxFoerdererlaenge > Sys(1).E(33) / 2 Then
        Fehler$ = Fehler$ & Lang_Res(695) & Chr$(13) & Chr$(10)  '-In Ihrer Anlage fehlen Trumlängen (Trum anklicken, Länge eingeben, >enter< druecken)
    End If
    
    'gibts da n element, das garnicht mit der anlage verbunden ist?
        For K = 9 To Maxelementindex
            If Sys(K).Element <> "" And (left(Sys(K).Tag, 1) = "0" Or left(Sys(K).Tag, 1) = "1") Then
                If Sys(K).Verb(1, 1) = 0 Or Sys(K).Verb(2, 1) = 0 Then
                    Fehler$ = Fehler$ & Lang_Res(710) & K & Lang_Res(711) & Chr$(13) & Chr$(10)   '- hinweis: element...ist nicht korrekt in ihre anlage eingebunden
                End If
            End If
        Next K
    


    
    'fehlerverwaltung und anzeige
    If B_Rex_AutoLauf = False Then
        B_Rex.Fehlerliste = Datenfehler$ & Datenaenderung$ & Fehler$ & SchwTransFehler$ & SchwLongFehler$
    End If
Exit Sub
End Sub

Public Sub Kontrollrechnungen(K As Integer, M As Double, mue As Double, Errfreq As Double)
    'bei iterationen kann man sich das hier sparen, daher auslagerung hierher
    'KI muss hier für jede rechnungsversion einmal rein, um anlage zu beurteilen
    Dim Flaechenmoment As Double
    Dim WellenbelastungFwFu As Double
    Dim Q As Double
    'kontrollrechnungen nach fertiger berechnung
    For K = 9 To Maxelementindex
        If Sys(K).E(13) > 5 Then  'umschlingung, also nur durchmesser
            Flaechenmoment = Sys(K).E(99)
            WellenbelastungFwFu = 0
            If Sys(K).E(120) > 0 Then Flaechenmoment = Sys(K).E(120)
            
            'erstmal alle wellenbelastungen ermitteln
                
                '1. dynamische Wellenbelastung/fliehkraftkontrolle
                
                    'an jeder scheibe dehnt sich das band etwas entspr. der fliehkraft dort
                    'es bildet sich als summe über das gesamte band eine zus. dehnung,
                    'um die das band an jeder stelle belastet und jede scheibe entlastet wird.
                    'die summe ist also an jeder scheibe relevant, nicht der dort entstehende anteil
                    Q = Sys(K).Furein ' - 2 * Fliehkraftsumme
                    M = Sys(K).Furaus ' - 2 * Fliehkraftsumme
                    Sys(K).E(52) = Sqr(Q ^ 2 + M ^ 2 - 2 * Q * M * Cos(Sys(K).E(13) / 180 * PI)) - Sys(K).E(51) '201507 fliehkraft von wellenbelastung abziehen, das ist kritische stelle, band dehnt sich nicht (bisher falsch angenommen)
                    
                    '*2 wegen belastung es bandes durch fliehkraft + entlastung scheibe, also 2mal
                    'auf jeden fall ist das ergebnis so überaus schlüssig
                    'und gewährleistet nur zustände, in denen nach abzug der fliehkraft im band eine restdehnung bleibt
                    'bsp: 0,8 im trum, davon 0,6 fliehkraft, also 0,2 max. auflegedehnung, diese summe stimmt nur mit der 2*
                   
                    '202004 ueber 180 grad nicht die kleiner werdende wellenbelastung verwenden, sondern summer F1 F2
                    WellenbelastungFwFu = Sys(K).E(52)
                    If Sys(K).E(13) > 180 Then WellenbelastungFwFu = Q + M 'besser wirds nicht
                    
                    'kleiner kunstgriff abseits wissenschaftlicher pfade,
                    'damit bei 0-umschlingung auch nichts an der scheibe hängenbleibt:
                    If Sys(K).E(13) < 90 And Sys(K).E(17) > 0 Then Sys(K).E(52) = Sys(K).E(52) * Sin(Sys(K).E(13) / 180 * PI) - Sys(K).E(51)
                    
                    'falls es null ist, was eigentlich nicht geht ;-)
                    'hebts aufgrund der fliehkraft von der scheibe ab
                    'zumindest einseitig bleibt keine restdehnung
                    If Sys(K).E(52) <= 0 Or Q <= 0 Or M <= 0 Then '52 = dynamische Wellenbelastung
                        If Sys(1).E(20) > 10 Then 'erst ab ernstzunehmenden geschwindikeiten zuschalten
                            Sys(K).E(52) = 1 'negative wellenbelastung kann's nicht geben
                            Fehler$ = Fehler$ & Lang_Res(676) & " (" & K & ")." & Chr$(13) & Chr$(10)   '!Das Band hebt aufgrund der Fliehkraft von der Scheibe ab (x).
                            Fehlerwert = Fehlerwert + 100
                        End If
                    End If
                    
                    '99 flaechenmoment
                    '100 durchbiegung statisch, spielt hier keine rolle
                    '101 dynamisch
                    '104 spitzenlast
                    '106 woelbhoehe verwendet
                    '7 woelbhoehe nach siegling
                    
                    'dazu passend die dynamische durchbiegung
                        Q = Sqr(Sys(K).E(52) ^ 2 + (9.81 * Sys(K).E(5)) ^ 2) 'kraft mit schwerkraft der scheibe ergänzen, geht einstweilen won waagerecht aus, also phytagoras :-)
                        If Flaechenmoment > 0 And Sys(K).E(103) > 0 Then Sys(K).E(101) = (5 * Q * Sys(K).E(4) ^ 3) / (384 * Flaechenmoment * Kst(Sys(K).E(103)).Einstellung)
                        
                        If Init_B_Rex_WoelbDurchb = 1 And Bandfuehrungvorhanden = True Then 'will er sie ueberhaupt kontrolliert haben?
                            
                            'zylindrische sollen sich nicht zu weit biegen, sonst läuft das Band übereinander
                                If Sys(K).E(106) > 0 Then 'hat eine wölbhöhe angegeben, also werden wir die auch kontrollieren
                                    If Sys(K).E(101) > Sys(K).E(106) * 0.6 Then 'wenn er mehr als 0.8 wegbiegt, 202004 auf 0.6 geaendert
                                        Fehler$ = Fehler$ & Lang_Res(696) & " (" & K & ")." & Chr$(13) & Chr$(10)   '- Durch die Durchbiegung beim Lauf geht die Bandführung an Scheibe ... verloren
                                        Fehlerwert = Fehlerwert + 10
                                    End If
                                End If
                            'End If
                        End If
                
                '2. Wellenbelastung bei spitzenlast:
            
                    If Spitzenlastvorhanden = False Then
                        Sys(K).E(105) = Sys(K).E(52) 'Wellenbelastung gleich der dynamischen
                        Sys(K).E(104) = Sys(K).E(101) 'durchbiegung auch
                    Else
                        Q = Sys(K).FureinSp ' - 2 * Fliehkraftsumme
                        M = Sys(K).FurausSp ' - 2 * Fliehkraftsumme
                        Sys(K).E(105) = Sqr(Q ^ 2 + M ^ 2 - 2 * Q * M * Cos(Sys(K).E(13) / 180 * PI))
                        If Sys(K).E(13) < 90 And Sys(K).E(17) > 0 Then Sys(K).E(105) = Sys(K).E(105) * Sin(Sys(K).E(13) / 180 * PI)
                        
                        If Sys(K).E(105) <= 0 Or Q <= 0 Or M <= 0 Then
                            Sys(K).E(105) = 0 'negative wellenbelastung kann's nicht geben
                            'lassen wir die kirche erstmal im dorf
                                'Fehler$ = Fehler$ & Lang_Res(676 ) & K & Lang_Res(677) & Chr$(13) & Chr$(10) '!Das Band hebt aufgrund der Fliehkraft von der Scheibe...ab
                                'Fehlerwert = Fehlerwert + 100
                        End If
                        
                        'dazu passend die durchbiegung
                        If Sys(K).FureinSp <= 0 Or Sys(K).FurausSp <= 0 Then 'natürlich nur, wenn es keinen slipstick an dieser scheibe gibt
                            Sys(K).E(104) = 0
                        Else
                            Q = Sqr(Sys(K).E(105) ^ 2 + (9.81 * Sys(K).E(5)) ^ 2) 'kraft mit schwerkraft der scheibe ergänzen, geht einstweilen won waagerecht aus, also phytagoras :-)
                            If Flaechenmoment > 0 And Sys(K).E(103) > 0 Then Sys(K).E(104) = (5 * Q * Sys(K).E(4) ^ 3) / (384 * Flaechenmoment * Kst(Sys(K).E(103)).Einstellung)
                        End If
                        
                        If Init_B_Rex_WoelbDurchb = 1 Then
                            'hat eine wölbhöhe angegeben, also werden wir die auch kontrollieren
                            If Sys(K).E(106) > 0 And Sys(1).E(95) > 3 Then 'aber einer hochlaufzeit von 5 sekunden
                                If Sys(K).E(104) > Sys(K).E(106) * 0.5 Then
                                    Fehler$ = Fehler$ & Lang_Res(700) & " (" & K & ")." & Chr$(13) & Chr$(10)   '- Durch die Durchbiegung beim Anfahren geht die Bandführung an Scheibe... verloren
                                    Fehlerwert = Fehlerwert + 10
                                End If
                            End If
                            
                            'zylindrische sollen sich auch nicht biegen, sonst läuft das Band übereinander
                            If Bandfuehrungvorhanden = True And Sys(K).E(106) = 0 Then
                                If Sys(K).E(104) > Sys(K).E(7) Then 'einfache wölbhöhenempfehlung
                                    Fehler$ = Fehler$ & Lang_Res(701) & " (" & K & ")." & Chr$(13) & Chr$(10)   '- Durch die Durchbiegung beim Anfahren kann das Band an Scheibe übereinanderlaufen
                                    Fehlerwert = Fehlerwert + 10
                                End If
                            End If
                        End If
                    End If
            
                '3. statische wellenbelastung ermitteln
                
                    If AuflTrumKraft > 0 Then
                        Sys(K).E(75) = Sin(Sys(K).E(13) / 2 / 180 * PI) * AuflTrumKraft + Sin(Sys(K).E(13) / 2 / 180 * PI) * AuflTrumKraft
                    Else
                        Sys(K).E(75) = 0
                    End If
                    
                    'dazu passend die statische durchbiegung nur ausrechnen, nicht kontrollieren, wie soll das band auch verlaufen?
                        Q = Sqr(Sys(K).E(75) ^ 2 + (9.81 * Sys(K).E(5)) ^ 2) 'kraft mit schwerkraft der scheibe ergänzen, geht einstweilen won waagerecht aus, also phytagoras :-)
                        If Flaechenmoment > 0 And Sys(K).E(103) > 0 Then Sys(K).E(100) = (5 * Q * Sys(K).E(4) ^ 3) / (384 * Flaechenmoment * Kst(Sys(K).E(103)).Einstellung)
            
            'faktor Fw/Fu /rho-wertkontrolle, zugeständnis an die alten druiden
                If Sys(K).Tag <> "005" And Sys(K).E(50) > 0 Then 'messerkanten brauchen das nicht, ergeben nulldivision
                    
                    'faktor 1,7?, kontrolle
                        'neu 202004, siehe oben
                        Sys(K).E(93) = WellenbelastungFwFu / Abs(Sys(K).E(50)) 'wellenbelastung dynamisch / umfangskraft gesamt normalbetrieb
                        
                        'aber nur wenn der benutzer von den grundeinstellungen her will, bei extremultus ab okt 2004 immer (bloedsinn, f&e idee)
                        'lt. dr meyer/kaemper ab juli 2006 aufgrund kundenbeschwerde wieder aufgehoben :-)
                        If Init_B_Rex_FwFu_Fehler = 1 Then 'Or Left(Sys(1).S(2), 1) = "8" Then
                            If Sys(K).E(93) > 0 Then  '(fw/fu-kennwert)
                                'nah dran
                                If Sys(K).E(93) > Fw_Fu_Winkelabh(Sys(K).E(13)) And Sys(K).E(93) < Fw_Fu_Winkelabh(Sys(K).E(13)) + 0.2 Then
                                    Fehler$ = Fehler$ & Lang_Res(674) & K & ") = " & Int(Sys(K).E(93) * 100) / 100 & Lang_Res(675) & Chr$(13) & Chr$(10)  'Fw/Fu von Scheibe(...;ev. Probleme bei der Kraftübertragung
                                    Fehlerwert = Fehlerwert + 10
                                End If
                                'gerissen
                                If Sys(K).E(93) <= Fw_Fu_Winkelabh(Sys(K).E(13)) Then
                                    Fehler$ = Fehler$ & "!" & Lang_Res(674) & K & ") = " & Int(Sys(K).E(93) * 100) / 100 & "; Kraftuebertragungsprobleme" & Chr$(13) & Chr$(10)  'Fw/Fu von Scheibe(...;  **** Kraftübertragungsprobleme
                                    Fehlerwert = Fehlerwert + 100 'zappenduster
                                End If
                            End If
                            '201606 auch fuer spitzenlast einbauen
                            If Spitzenlastvorhanden = True Then
                                If Abs(Sys(K).E(118)) > 0 And Sys(K).E(105) > 0 Then
                                    FuFwSpitze = Sys(K).E(105) / Abs(Sys(K).E(118))
                                    If FuFwSpitze <= Fw_Fu_Winkelabh(Sys(K).E(13)) + 0.2 Then
                                        Fehler$ = Fehler$ & Lang_Res(713) & K & ") = " & Int(FuFwSpitze * 100) / 100 & "; " & Lang_Res(714) & Chr$(13) & Chr$(10)  'Fw/Fu von Scheibe(...;  **** Kraftübertragungsprobleme
                                        'Fehlerwert = Fehlerwert + 100 'nicht spitzenlast bestrafen, das geht vorbei...
                                    End If
                                End If
                            End If
                        
                        End If
                        
                    
                    'abgleich funenn
                        'der groesste schwachsinn, leider politisch von den wirrkoepfen und angsthasen durchgesetzt
                        'am besten verkaufen wir kein extremultus, es koennte ja schiefgehen
                        'bei mitarbeitern (userstatus 2) bleibts weiter freigestellt
                        'funenn (108) bei funennauflegedehnung (110)
                        If left(Sys(1).S(2), 1) = "8" Then
                            If Init_B_Rex_FuNenn_Fehler = 1 Or (Init_B_Rex_FuNenn_Fehler = 0 And UserStatus <> 2) Then
                                If Sys(1).E(108) > 0 And Sys(1).E(109) > 0 And Sys(1).E(110) > 0 Then 'funenn vorhanden?
                                    Q = (Sys(K).Furein + Sys(K).Furaus) / 2 * 2 / (SystemTyp.Kraftdehnung * Sys(1).E(34)) 'dehnung band an dieser scheibe
                                    'fuelement/dehnung>funenn*funennmax*bandbreite/funennangabebeidehnung..
                                    If Sys(K).E(50) / Q > Sys(1).E(108) * Sys(1).E(109) * Sys(1).E(34) / Sys(1).E(110) Then
                                        Fehler$ = Fehler$ & "!FuNenn dieses Typs an Scheibe (" & K & ") ueberschritten" & Chr$(13) & Chr$(10) 'Fw/Fu von Scheibe(...;  **** Kraftübertragungsprobleme
                                        Fehlerwert = Fehlerwert + 100 'zappenduster
                                    End If
                                End If
                            End If
                        End If
                    
                    
                    'rho-Wert berechnen und bewerten
                        Sys(K).E(76) = Sys(K).E(50) * 360 / (PI * Sys(K).E(2) * Sys(1).E(34) * Sys(K).E(13))
                        If Init_B_Rex_rho_Wert_Fehler = 1 Then 'aber nur wenn der benutzer von den grundeinstellungen her will
                            Select Case left(LCase(Sys(1).S(5)), 3) 'transilon, geht zwar, ribbelt aber anders als extremultus früh drauflos
                                Case "tra"
                                    'gummiert oder nicht?
                                    If Sys(Antriebsscheibe).E(14) = 17 Then 'gummiert
                                        If Sys(K).E(76) > 0.045 Then
                                            Fehler$ = Fehler$ & Lang_Res(678) & K & Lang_Res(679) & Chr$(13) & Chr$(10)   '- rho-wert von Element...too high; ev. Probleme bei der Kraftübertragung
                                            Fehlerwert = Fehlerwert + 10
                                        End If
                                    Else 'blank
                                        If Sys(K).E(76) > 0.03 Then
                                            Fehler$ = Fehler$ & Lang_Res(678) & K & Lang_Res(679) & Chr$(13) & Chr$(10)   '- rho-wert von Element...too high; ev. Probleme bei der Kraftübertragung
                                            Fehlerwert = Fehlerwert + 10
                                        End If
                                    End If
                                Case "ext" 'extremultus
                                    'PA 0,08
                                    'PET 0,1 (B81 & B82)
                                    'A 0,15 (B81 & B82)
                                    If SystemTyp.rho = 0 Then SystemTyp.rho = 0.08
                                    If Sys(K).E(76) > SystemTyp.rho Then '0.08 Then  '202004 rho wert unterschieden nach zugträger
                                        Fehler$ = Fehler$ & Lang_Res(678) & K & Lang_Res(679) & " (>" & SystemTyp.rho & ")" & Chr$(13) & Chr$(10)   '- rho-wert von Element...too high; ev. Probleme bei der Kraftübertragung
                                        Fehlerwert = Fehlerwert + 10
                                    End If
                            End Select
                        End If
                End If
               
            'durchmesserkontrolle nach meiner formel
                If Sys(K).E(2) > 0 Then 'And Sys(K).E(17) > 20 Then '17 umfangskraft messerkanten also nicht
                    '(ist ein durchmesser und soll übertragen,
                    'messerkanten ausgeschlossen
                    
                    'Mue ermitteln zwischen Band und scheibe
                
                    If Sys(K).E(41) = 1 Then
                        mue = Sys(1).E(78) 'bandreibungszahl lauf
                    Else
                        mue = Sys(1).E(77) 'mit tragseite zum messer, auch wenn er damit seine anlage abfackelt
                        
                        If Sys(1).E(77) = 0 Then
                            Fehler$ = Fehler$ & Lang_Res(702) & K & Lang_Res(703) & Chr$(13) & Chr$(10)   '- Element ... Reibungszahl Tragseite/Antriebsseite des Bandes wurde gebraucht, aber nicht gefunden
                        End If
                    End If
                    'negatives mue sind geschaetzte werte aus der datenbank
                    'If Mue < 0 Then Fehler$ = Fehler$ & Lang_Res(640) & K & Lang_Res(641) & Chr$(13) & Chr$(10)   '-Reibungszahl Band-Scheibe ... geschätzt
                    
                    mue = ReibungszahlBerechnung(mue, 2, Sys(K).E(14), True)
                    
                    
                    'erforderlicher Durchmesser dynamisch, statisch gibts ja nix zum mitnehmen
                    'Sys(K).E(72) = (Abs(Sys(K).E(50)) / (Sys(1).E(34) * Abs(Sys(K).Furaus + Sys(K).Furein)) + 0.001) / (0.26 * Mue - 0.017) * 180 / Sys(K).E(13) * 1000 * 1.3
                    'Sys(K).E(72) = (Abs(Sys(K).E(50)) / (Sys(1).E(34) * Abs(Sys(K).Furaus + Sys(K).Furein)) + 0.0002) / (0.26 * Mue - 0.017) * 230 / Sys(K).E(13) * 1000 * 1.3
                    '1,3 als sicherheit für meinen persönlichen ruhigen Schlaf, formel noch nicht so richtig wahr
                    
                    If mue > 0 Then Sys(K).E(72) = Abs(Sys(K).E(50)) ^ 1.6 * 25000 / (Sys(1).E(34) * Sys(K).E(52) * mue * Sys(K).E(13))
                    'erf. Durchmesser = Fu^1.6*25000/(bo*Fw*Mue*umschlingung)
                    
                    '-Durchmesser Element () reicht zur Übertragung der Umfangskraft nicht aus
                    If Init_B_Rex_KraftUebertrkontr = 1 Then
                        If Sys(K).E(72) > Sys(K).E(2) And Sys(K).E(17) > 0 Then
                            Fehler$ = Fehler$ & Lang_Res(663) & K & Lang_Res(664) & Chr$(13) & Chr$(10)
                            Fehlerwert = Fehlerwert + 100
                        End If
                    End If
                    
                    'erf. durchmesser auf mindestdurchmesser erhöhen, falls nötig'201902
                    If Sys(K).E(41) = 1 Then If Sys(K).E(72) < Abs(Sys(1).E(86)) Then Sys(K).E(72) = Abs(Sys(1).E(86)) 'laufseite, also wert ohne gegenbiegung
                    If Sys(K).E(41) = 2 Then If Sys(K).E(72) < Abs(Sys(1).E(119)) Then Sys(K).E(72) = Abs(Sys(1).E(119)) 'tragseite, also wert mit gegenbiegung
                                                            
                    'überarbeiten bei spitzenlast:
                    'erforderlicher Durchmesser im statischen zustand zusätzlich
                    'Memo = (Abs(Sys(K).E(50)) / (Sys(1).E(34) * Abs(2 * AuflTrumKraft)) + 0.001) / (0.26 * Mue - 0.017) * 180 / Sys(K).E(13) * 1000 * 1.3
                    'If Memo > Sys(K).E(72) Then Sys(K).E(72) = Memo 'statischer zustand beim anfahren erfordert größeren durchmesser als dynamischer
                    '1,3 als sicherheit für meinen persönlichen ruhigen Schlaf, formel noch nicht so richtig wahr
                    'If Sys(K).E(72) > Sys(K).E(2) And Sys(K).E(17) > 0 Then Fehler$ = Fehler$ & "-Durchmesser Element (" & K & ") reicht zur Übertragung der Umfangskraft beim Anfahren nicht aus." & Chr$(13) & Chr$(10)
                End If
                
            'mindestdurchmesserkontrolle scheibe, messerkante besitzt eine eigene und durchläuft diese schleife nicht
                If Sys(K).Tag <> "005" And Init_B_Rex_Minddurchmkontr = 1 Then  'messerkanten nicht
                    If Sys(K).E(41) = 1 Then
                        If Sys(K).E(2) < Abs(Sys(1).E(86)) Then 'laufseite
                            Fehler$ = Fehler$ & "!" & Lang_Res(680) & K & Lang_Res(681) & "(" & Sys(1).E(86) & "mm)" & vbCrLf  'unterschreitet mindestdurchmesser
                            Fehlerwert = Fehlerwert + 100
                        End If
                    End If
                    If Sys(K).E(41) = 2 Then 'tragseite
                        If Sys(K).E(2) < Abs(Sys(1).E(119)) Then
                            Fehler$ = Fehler$ & "!" & Lang_Res(680) & K & Lang_Res(681) & "(" & Sys(1).E(119) & "mm)" & vbCrLf  'unterschreitet mindestdurchmesser
                            Fehlerwert = Fehlerwert + 100
                        End If
                    End If

                End If
        
            'eigenfrequenzkontrolle
                            
                If Sys(K).Tag <> "005" And Schwingungen_berechnen = True Then 'also keine messerkanten, sonst alle
                    
                    If Sys(K).E(111) = 0 Then Sys(K).E(111) = 1  '201909, irgend ne unwucht hat doch jeder :-)
                    
                    'abgleich mit erregungen pro umdrehung
                        'If Sys(K).E(111) > 0 Then 'nur wenn eine erregung vorgegeben ist
                            'erste verbindungF
                            '111 anregungen
                            '21 drehzahl in umin
                            
                            Errfreq = Sys(K).E(111) * Sys(K).E(21) / 60
                            If Errfreq > 0.8 * Sys(K).Verb(1, 4) And Errfreq < 1.2 * Sys(K).Verb(1, 4) Then '(1,4) eigenfrequenz dieses trumstueckchens
                                'ganz nah dran auf 20%
                                Fehler$ = Fehler$ & Lang_Res(185) & K & Lang_Res(186) & Round(Sys(K).Verb(1, 3)) & "mm. " & vbCrLf '-Element ( // "): kritische Eigenfrequenzanregung transversal des Trums mit Länge "
                                FehlerwertSchwingungen = FehlerwertSchwingungen + 100
                            Else
                                If Errfreq > 0.75 * Sys(K).Verb(1, 4) And Errfreq < 1.25 * Sys(K).Verb(1, 4) Then
                                    'in der naehe, 30%
                                    Fehler$ = Fehler$ & Lang_Res(185) & K & Lang_Res(186) & Round(Sys(K).Verb(1, 3)) & "mm. " & vbCrLf
                                    FehlerwertSchwingungen = FehlerwertSchwingungen + 10
                                End If
                            End If
                            'zweite verbindung
                            If Errfreq > 0.8 * Sys(K).Verb(2, 4) And Errfreq < 1.2 * Sys(K).Verb(2, 4) Then
                                'ganz nah dran auf 20%
                                Fehler$ = Fehler$ & Lang_Res(185) & K & Lang_Res(186) & Round(Sys(K).Verb(2, 3)) & "mm. " & vbCrLf
                                FehlerwertSchwingungen = FehlerwertSchwingungen + 100
                            Else
                                If Errfreq > 0.75 * Sys(K).Verb(2, 4) And Errfreq < 1.25 * Sys(K).Verb(2, 4) Then
                                    'in der naehe, 30%
                                    Fehler$ = Fehler$ & Lang_Res(185) & K & Lang_Res(186) & Round(Sys(K).Verb(2, 3)) & "mm. " & vbCrLf
                                    FehlerwertSchwingungen = FehlerwertSchwingungen + 10
                                End If
                            End If
                            
                            
                        'End If
                    
                End If
        End If 'umschlingung > grenzwert
    Next K
    
    'messerbänder gegen Schrumpf schützen
        If Sys(1).E(53) < 0.2 And Messervorhanden = True Then
            Fehler$ = Fehler$ & Lang_Res(707) & Chr$(13) & Chr$(10)  'Messerbänder mir einer Auflegedehnung unterhalb von 0,2% können schrumpfen.
            Fehlerwert = Fehlerwert + 10
        End If
        
    'keine dehnung?
        If Fuerstes < 0 Then
            'Zulässig = False
            Fehlerwert = Fehlerwert + 100
            Fehler$ = Fehler$ & Lang_Res(605) & Chr$(13) & Chr$(10) '!-Keine Restspannung im Leertrum gewährleistet
        End If
        If Sys(1).E(53) < 0 Then
            'Zulässig = False
            Fehlerwert = Fehlerwert + 100
            Fehler$ = Fehler$ & Lang_Res(606) & Chr$(13) & Chr$(10)  '!-keine Auflegedehnung?
        End If
    
    'maximale Dehnung/Auflegedehnung eingehalten?
        If Dehnung$ = Lang_Res(646) Then  '"max. zul. Auflegedehn. "'AuflTrumKraft und MaxTrumKraft enthalten kräfte
            If Sys(1).E(84) >= 1.4 Then 'bloss bei aramid nicht diese meldung bringen
                If Sys(1).E(53) < 0.3 Then Fehler$ = Fehler$ & Lang_Res(708) & Chr$(13) & Chr$(10) '- Hinweis: unterhalb von 0,3% Auflegedehnung schwankt das Kraft-Dehnungsverhalten von Förderbändern/ Antriebsriemen aus Kunststoff erheblich (Spannweg?).
            End If
            If AuflTrumKraft > MaxTrumKraft And AuflTrumKraft <= 1.2 * MaxTrumKraft Then
                Fehler$ = Fehler$ & Lang_Res(668) & Chr$(13) & Chr$(10) 'die auflegedehnung ist etwas zu hoch
                Fehlerwert = Fehlerwert + 10
            End If
            If AuflTrumKraft > MaxTrumKraft * 1.2 Then
                Fehler$ = Fehler$ & Lang_Res(652) & Chr$(13) & Chr$(10)  '"!-Die Auflegedehnung ist unzulässig hoch."
                Fehlerwert = Fehlerwert + 100
                'Zulässig = False
            End If
            If Fumax > MaxTrumKraft * 1.5 Then
                Fehler$ = Fehler$ & Lang_Res(656) & Chr$(13) & Chr$(10)  '"!-Die maximale Dehnung (dicke, durchgezogene Linie) ist unzulässig hoch."
                Fehlerwert = Fehlerwert + 100
                'Zulässig = False
            End If
            If FumaxSp > MaxTrumKraft * 1.7 Then
                If Sys(1).E(95) > 0 Then 'beschleunigung vorhanden
                    Fehler$ = Fehler$ & Lang_Res(704) & Chr$(13) & Chr$(10)  '"!-Das Dehnungsmaximum beim Anfahren ist unzulässig hoch
                    Fehlerwert = Fehlerwert + 100
                End If
                'Zulässig = False
            End If
        End If
        
        If Dehnung$ = Lang_Res(647) Then  '"max. zul. Dehnung "'AuflTrumKraft und MaxTrumKraft enthalten kräfte
            If Sys(1).E(85) >= 2 Then 'bloss bei aramid nicht diese meldung bringen
                If Sys(1).E(53) < 0.3 Then Fehler$ = Fehler$ & Lang_Res(708) & Chr$(13) & Chr$(10)  '& Lang_Res(708 )
            End If

            'AuflTrumKraft wird hier völlig neu definiert, hat auch mit fliehkraft nichts tun
            'H = Fumax 'Abs(Fumax * 2 / (Systemtyp.Kraftdehnung * sys(1).e(34)))
            If Fumax > MaxTrumKraft And Fumax <= 1.2 * MaxTrumKraft Then
                Fehler$ = Fehler$ & Lang_Res(655) & Chr$(13) & Chr$(10) '"-Die maximale Dehnung ist etwas zu hoch."
                Fehlerwert = Fehlerwert + 10
            End If
            If Fumax > 1.2 * MaxTrumKraft Then
                Fehler$ = Fehler$ & Lang_Res(656) & Chr$(13) & Chr$(10)  '"!-Die maximale Dehnung (dicke, durchgezogene Linie) ist unzulässig hoch."
                Fehlerwert = Fehlerwert + 100
                'Zulässig = False
            End If
            If FumaxSp > 1.4 * MaxTrumKraft Then 'fumaxsp-fumax
                Fehler$ = Fehler$ & Lang_Res(704) & Chr$(13) & Chr$(10) '!- das Dehnungsmaximum beim Anfahren ist unzulässig hoch
                Fehlerwert = Fehlerwert + 100
                'Zulässig = False
            End If
        End If
        
    'slip-stick-effekt beim anfahren
        If FuminSp < 0 Then
            Fehler$ = Fehler$ & Lang_Res(705) & Chr$(13) & Chr$(10)  '- keine Restdehnung im Leertrum beim Anfahren = Slip-Stick-Effekt
            Fehlerwert = Fehlerwert + 10
            'Zulässig = False
        End If
        
    'minimale Auflegedehnung auch noch verwursteln, meistens allerdings fehlt diese Angabe und ist damit 0
        If MinTrumKraft > 0 Then
            If AuflTrumKraft < MinTrumKraft Then
                Fehler$ = Fehler$ & Lang_Res(712) & Chr$(13) & Chr$(10)  '"!- Die minimale Auflegedehnung wurde unterschritten.
                'Fehlerwert = Fehlerwert + 100'2017 raus damit, hinweis reicht, dann wird sie eben unterschritten. einmal an der schraube drehen, schon gehts wieder
            End If
        End If

End Sub

Private Sub Berechnung()
Dim Masse As Double, Restmasse As Double, Restlaenge As Double, MaxMasse As Double, Memo As Double, Memo1 As Double
Dim DurchschMasse As Double, M As Double, Q As Double, P As Double 'eine Stückchens (bei tragrollen-, rollenbahnen)
Dim K As Integer, IaltK As Integer, j As Integer, H As Integer, B As Integer
Dim mue As Double, Mue1 As Double, Mue2 As Double, Mue3 As Double
Dim Tragseite_Mue_verwendet As Boolean
Dim i As Integer
    
    Fehlerwert = 0
    FehlerwertSchwingungen = 0
    
    Fehler$ = ""
    'Zulässig = True
    'leertrum spannen, nur wenn keine Feder/gewichtspannstation
    'vorher die fliehkräfte festlegen
    'Überlastvorgabe = 1 'nämlich garkeine
    
    K = 1 'ist hier der Zähler
        
    Fuletztes = Fuerstes 'nur merken für die jeweils nächste schleife Rechnung
    FuletztesSp = Fuerstes
    Fumin = Fuerstes 'minimalwert festhalten zur späteren skalierung
    Fumax = Fuerstes 'maximalwert festhalten zur späteren skalierung
    FumaxSp = Fuerstes
    
    K = Startelement 'wird oben in richtiger richtung festgelegt
    IaltK = Antriebsscheibe 'von da kommt er, da soll er nicht gleich wieder hin
    Sys(Antriebsscheibe).Lraus = 0 'sonst wird immer die doppelte länge aufaddiert
    
    Do
        Fu = 0
        Staumasse = 0
        Tragseite_Mue_verwendet = False
        
        Sys(K).Furein = Fuletztes
        
        'den vorherigen bandabschnitt im kreislauf ermitteln
        'abstände werden immer in beiden beteilgten elementen protokolliert
        If Sys(K).Verb(1, 1) = IaltK Then
            Sys(K).Lrein = Sys(K).Verb(1, 3) + Sys(IaltK).Lraus 'trumlänge +bisher protokoll. länge
        Else
            Sys(K).Lrein = Sys(K).Verb(2, 3) + Sys(IaltK).Lraus
        End If
        
        Select Case Sys(K).Tag
            '001-antriebsscheibe wird nicht berechnet
            Case "002" 'antriebsscheibe, sekundäre
                Fu = Sys(K).E(17) 'vorgegebene umfangskraft
                
                'biegeleistung
                Sys(K).E(66) = (Sys(1).E(34) * (4 + Sys(1).E(20)) * Sqr(Sys(K).E(13))) / Sys(K).E(2) * Sys(1).E(80)
                Sys(K).E(50) = Abs(-Fu + Sys(K).E(66))
                      
                'Spitzenlast aus beschleunigung der Masse:
                Sys(K).E(98) = Sys(1).E(68) * MassentraegheitsErmittlung(K) / (Sys(K).E(2) / 2000) ^ 2 'J*a/r^2  wenn eins fehlt, ist alles zusammen e null
                          
                Sys(K).Furaus = Sys(K).Furein - Fu + Sys(K).E(66)
                Sys(K).Lraus = Sys(K).Lrein + PI * Sys(K).E(2) / 360 * Sys(K).E(13)
            Case "003" 'umlenkscheibe
               
                'biegeleistungsanteil
                '66 = biegeleistungsanteil
                '34 = bandbreite
                '20 = bandgeschwindigkeit
                '13 = umschlingung
                '2 = durchmesser
                '80 = biegeleistungskennwert
                
                
                Sys(K).E(66) = (Sys(1).E(34) * (4 + Sys(1).E(20)) * Sqr(Sys(K).E(13))) / Sys(K).E(2) * Sys(1).E(80)
                
                Fu = Sys(K).E(17) + Sys(K).E(66)
                Sys(K).E(50) = Fu
                
                'Spitzenlast aus beschleunigung der Masse:
                Sys(K).E(98) = Sys(1).E(68) * MassentraegheitsErmittlung(K) / (Sys(K).E(2) / 2000) ^ 2 'J*a/r^2  wenn eins fehlt, ist alles zusammen e null
                If Kst(Sys(K).E(60)).Einstellung <> 22 And Kst(Sys(K).E(60)).Einstellung <> 0 Then
                    Sys(K).E(98) = Sys(K).E(98) + Fu / 100 * Kst(Sys(K).E(60)).Einstellung 'automatische überlastvorgabe
                Else
                    Sys(K).E(98) = Sys(K).E(98) + Fu / 100 * Sys(K).E(59) 'manuelle überlastvorgabe
                End If
                
                Sys(K).Furaus = Sys(K).Furein + Fu
                Sys(K).Lraus = Sys(K).Lrein + PI * Sys(K).E(2) / 360 * Sys(K).E(13)
                'I = I
            Case "005" 'messerkante
                
                'eigene, keine ausgelagerte durchmesserkontrolle
                'später halt mal auslagern
                If Sys(K).E(45) < Abs(Sys(1).E(86)) And Init_B_Rex_Minddurchmkontr = 1 Then
                    'Zulässig = False
                    '201902, nicht nach gegenbiegung oder ohne unterscheiden, weil e nur die laufseite beim messer in frage kommt
                    Fehlerwert = Fehlerwert + 100
                    Fehler$ = Fehler$ & Lang_Res(610) & K & Lang_Res(609) & Sys(1).E(86) & "mm)" & Chr$(13) & Chr$(10)  '-Messerkante ... unterschreitet Mindestdurchmesser (
                End If
                
                If Sys(K).E(41) = 1 Then
                    mue = Sys(1).E(78)
                Else
                    mue = Sys(1).E(77) 'mit tragseite zum messer, auch wenn er damit seine anlage abfackelt
                    Tragseite_Mue_verwendet = True
                End If
                
                'If mue < 0 Then Fehler$ = Fehler$ & Lang_Res(611) & K & Lang_Res(612) & Chr$(13) & Chr$(10)   '-Reibungszahl Band-Messerkante ... geschätzt
                
                mue = ReibungszahlBerechnung(mue, 1, Sys(K).E(48), True)

                Fu = (mue * 0.001 * Sys(K).Furein / Sys(1).E(34) * Sys(K).E(13) * (Sys(K).E(20) + 2.578 * Sys(K).E(45) / 2 + 17.88)) * Sys(1).E(34) 'pro mm gerechnet, umgerechnet auf breite
                If Fu < 0 Then Fu = 0 'aus keiner kraft darf messerkante auch keine negative machen
                Sys(K).E(49) = (mue * 0.55 * Sys(K).Furein / Sys(1).E(34) * Sys(K).E(13) * (Sys(K).E(20) + 0.08 * Sys(K).E(45) / 2 + 0.32)) + 23 'temperatur an der messerkante
                If Sys(K).E(49) > 200 Then 'temperatur > 200 glaubt mir sowieso keiner, aber put ist put
                    'Zulässig = False
                    Sys(K).E(49) = 200
                    Fehlerwert = Fehlerwert + 100
                    Fehler$ = Fehler$ & Lang_Res(613) & K & Lang_Res(614) & Chr$(13) & Chr$(10)   '-Messerkantentemperatur ... >200°C, Überhitzung
                End If
                
                'biegeleistung messerkante, ist aber oben schon berücksichtigt (das ist die frage, wahrscheinlich kann mans getrost obendraufschlagen), wird nur extra ausgewiesen
                Sys(K).E(66) = (Sys(1).E(34) * (4 + Sys(1).E(20)) * Sqr(Sys(K).E(13))) / Sys(K).E(45) * Sys(1).E(80)

                Sys(K).E(50) = Fu
                Sys(K).Furaus = Sys(K).Furein + Fu
                Sys(K).Lraus = Sys(K).Lrein + PI * Sys(K).E(13) / 360 * Sys(K).E(45)
            Case "101" 'tisch, träger
                
                'Mue reibungszahl zwischen Band und Gleittischunterlage
                'Mue1 reibungszahl zwischen band und transportgut
                    If Sys(K).E(41) = 1 Then
                        mue = Sys(1).E(78) 'Laufseite zum Tisch
                        Mue1 = Sys(1).E(77) 'Tragseite zum Transportgut
                    Else
                        mue = Sys(1).E(77) 'Tragseite zum Tisch
                        Mue1 = Sys(1).E(78) 'Laufseite zum Transportgut
                        Tragseite_Mue_verwendet = True
                    End If
                
                mue = ReibungszahlBerechnung(mue, 1, Sys(K).E(15), True)
                
               
                'transportierte masse ermitteln
                'trägerRestlaenge ermitteln
                'repräsentative reibungszahl ermitteln (p,q) bei transportgut
                'für geschwindigkeitsgewinn alles in einer schleife
                    P = 0
                    Q = 0
                    Masse = 0 'nur die bewegte masse
                    MaxMasse = 0
                    Restlaenge = Sys(K).E(22)
                    B = 0
                    For j = 9 To Maxelementindex
                        If Sys(j).Zugehoerigkeit = K Then
                            If Sys(j).Tag = "201" Then 'transportgut
                                If Sys(j).E(32) > MaxMasse Then MaxMasse = Sys(j).E(32)
                                Masse = Masse + Sys(j).E(28)
                                P = P + Sys(j).E(28) * Kst(Sys(j).E(36)).Einstellung
                                Q = Q + Sys(j).E(28)
                            End If
                            If Sys(j).Tag = "204" Then 'stau
                                Masse = Masse - Sys(j).E(23) 'staumassen abziehen
                                Restlaenge = Restlaenge - (Sys(j).E(46) - Sys(j).E(25)) 'rechts -linke grenze
                                Tragseite_Mue_verwendet = True
                            End If
                            If Sys(j).Tag = "205" Then 'abweiser, einer reicht
                                B = j 'position merken
                                Tragseite_Mue_verwendet = True
                            End If
                       End If
                    Next j
                    If P > 0 And Q > 0 Then 'wenn garnichts gestaut wurde
                        Mue1 = Mue1 * (P / Q) 'durchschnittliche reibung wird errechnet
                        If B > 0 Then Mue3 = P / Q * (Kst(Sys(B).E(15)).Einstellung * 0.2)
                    End If
                    P = 0 'stehen wieder zur Verfügung
                    Q = 0
                    
                    'Mue1, Mue2 = reibungszahl zw. transportgut und band
                    'Mue1 möglichst hoch, wenn's Gut rutschen muß wegen stau
                    'Mue2 mögl. niedrig, wenn's Gut rutschen könnte wegen Steigung
                    Mue2 = Mue1
                    Mue1 = ReibungszahlBerechnung(Mue1, 1, 0, False) 'false = ohne materialpaarung
                    Mue2 = ReibungszahlBerechnung(Mue2, 2, 0, False)
                    
                    
                    
                Restmasse = Masse
                If Restmasse < 0 Then
                    Fehler$ = Fehler$ & Lang_Res(625) & K & Lang_Res(615) & Chr$(13) & Chr$(10)  '-Förderer: .... Stau und Abweiser mangels Transportgut nicht berechnet.
                End If
                
                If Masse < 0 Then Masse = 0 'kann passieren, wenn stau ohne transportgut da ist
                                
                'noch ohne bandmasse, denn deren potentialsumme ist 0, weil band endlos
                
                If Sys(K).Rechts = True Then
                    
                    Fu = Fu + Masse * 9.81 * Sin(Sys(K).E(16) * PI / 180) 'm*g*sin alpha 'fu durch potentialunterschiede mit dieser masse berechnen
                Else
                    Fu = Fu + Masse * 9.81 * Sin(-Sys(K).E(16) * PI / 180) 'm*g*sin alpha 'fu durch potentialunterschiede mit dieser masse berechnen
                End If
                
                'jetzt bandmasse dazu, denn sie wird ja auch über den tisch geschliffen
                Masse = Masse + (Sys(K).E(22) / 1000 * Sys(1).E(34) / 1000) * Sys(1).E(81) 'bandmasse dazuzählen
                'bandmasse wird nur über nichtgestaute bereiche verteilt.
                'das ist zwar falsch, der fehler ist aber gering und hinnehmbar
                
                'massenkorrektur entsprechend der Steigung:
                Masse = Masse * Cos(Sys(K).E(16) * PI / 180) 'masse bezieht sich jetzt nur auf das zur zeit bewegte gut und auf fu durch reibung, nicht vom reversieren abh.
                
                'kräfte durch reibung addieren, DAS ist die eigentliche rechnung, die steigung steckt in der masse
                Fu = Fu + 9.81 * Masse * mue 'fu durch reibung, steigungsausgleich ist in der masse enthalten
                
                If Restlaenge = 0 Then Restlaenge = 0.1 'falls staulänge mit trägerlänge übereinstimmt
                Sys(K).Fusteig = Fu / Restlaenge 'steigung ermitteln
                
                'm*g*cosa=m*g*Mue*sina'transportgut rutscht?
                If Sys(K).E(16) <> 0 Then 'steigungswinkel
                    If Mue2 <> 0 Then
                        If Tan(Abs(Sys(K).E(16)) * PI / 180) > Mue2 And Tan(Abs(Sys(K).E(16)) * PI / 180) <= Mue2 * 1.4 Then 'irgendwo zw. haft- und gleitreibung
                            'Fehlerwert = Fehlerwert + 10
                            Fehler$ = Fehler$ & Lang_Res(619) & Chr$(13) & Chr$(10) '-Das Transportgut könnte auf dem Band rutschen
                        End If
                        If Tan(Abs(Sys(K).E(16)) * PI / 180) > Mue2 * 1.4 Then '1.4 Haftreibungsverhältnis zu gleitreibungsverhältnis
                            'Zulässig = False
                            'Fehlerwert = Fehlerwert + 100
                            Fehler$ = Fehler$ & Lang_Res(620) & Chr$(13) & Chr$(10) '!-Das Transportgut wird auf dem Band rutschen
                        End If
                    End If
                    Tragseite_Mue_verwendet = True
                End If
                
                'stau, abweiser, freie umfangskraft
                For j = 9 To Maxelementindex
                    If Sys(j).Zugehoerigkeit = K Then
                        
                        'stau
                        If Sys(j).Tag = "204" And Restmasse >= 0 Then
                            Sys(j).E(50) = 9.81 * (Sys(j).E(23) * Cos(Sys(K).E(16) * PI / 180)) * (Mue1 + mue) 'reibung band-tisch ist in der steigung enthalten
                            
                            'stau hat nur reibung, keine potentialunterschiede
                            
                            Fu = Fu + Sys(j).E(50)
                        End If
                        
                        'abweiser
                        If Sys(j).Tag = "205" And Restmasse >= 0 Then
                                
                            If Restmasse > MaxMasse Then 'ist überhaupt nichtgestaute masse übrig
                                'beschaffenheit transportgut in Mue1 durch p/q enthalten
                                'merk ist die kraft, mit der das Fördergut gegen abweiser drückt, wenn ein stau wäre, nur reibung!, druck durch schwerkraft entfällt gegenüber stau
                                Merk = 9.81 * MaxMasse * Cos(Sys(K).E(16) * PI / 180) * (Mue1) 'auf tisch rutscht das gut rechnerisch oben, cos für steigung, aus reibung
                                If Mue3 * 1.2 > Tan((90 - Sys(j).E(26)) * PI / 180) Then '1.2 selbsteingebauter angstfaktor
                                    Fehler$ = Fehler$ & Lang_Res(621) & j & Lang_Res(622) & Chr$(13) & Chr$(10)  '-Der Abweiser ("") läßt wahrscheinlich kein Transportgut durch (Berechnung ähnlich Stau)
                                    Fehlerwert = Fehlerwert + 10
                                    'so ähnlich wie stau, nur ohne reibung band-tisch, merk von oben wird verwendet
                                Else
                                    Merk = Merk * Mue3 * Sin(Sys(j).E(26) * PI / 180) 'nur die reibung am abweiser (merk aus gewicht und reibung)
                                    'auch wenn die schwerkraft abwärts dazukommt, mit mehr als merk kann das band nicht drücken!
                                End If
                                Sys(j).E(50) = Merk
                                Fu = Fu + Sys(j).E(50)
                            Else
                                Fehler$ = Fehler$ & Lang_Res(623) & j & Lang_Res(624) & Chr$(13) & Chr$(10)   '-keine Abweiserberechnung ("") durchgeführt, da zuviel Masse im Stau
                                Sys(j).E(50) = 0
                            End If
                        End If
                        
                        'trägergebundene_Umfangskraft
                        If Sys(j).Tag = "206" Then
                            Sys(j).E(50) = Sys(j).E(44) + Sys(K).Fusteig * (Sys(j).E(46) - Sys(j).E(25))
                            Fu = Fu + Sys(j).E(44) 'eingeleitete Kraft
                            'steigung für Spitzenlast wird unten bei der ermittlung der auflegedehnung ausgerechnet
                        End If
                    End If
                Next j
                
                'spitzenlast bei beschleunigung der masse
                If Restmasse > 0.1 And Restlaenge > 0.1 Then
                    Sys(K).E(98) = Restmasse * Sys(1).E(68) 'F = m*a
                    Sys(K).FusteigSp = Sys(K).E(98) / Restlaenge 'diese steigung durch die beschleunigung der ungestauten transportmasse
                Else
                    Sys(K).E(98) = 0
                    Sys(K).FusteigSp = 0
                End If
                Sys(K).FusteigSpRoll = 0 'nur zur Vorsicht, falls noch von früher was drinsteht
                                
                Sys(K).E(50) = Fu 'alles auf den träger
                Sys(K).Furaus = Sys(K).Furein + Fu
                Sys(K).Lraus = Sys(K).Lrein + Sys(K).E(22)
            Case "103" ' Rollenbahn, Träger
                DurchschMasse = 0
                DurchschFaktor = 1
                Mue3 = 1
                MaxMasse = 0
                Masse = 0 'masse enthält bewegte masse
                Staumasse = 0
                Staulänge = 0
                Restlaenge = Sys(K).E(22)
                Q = 0
                M = 0
                H = 0
                P = 0
                
                '28 = Masse dieser Transportgutart
                
                For j = 9 To Maxelementindex
                    If Sys(j).Zugehoerigkeit = K Then
                        If Sys(j).Tag = "201" Then 'transportgut
                            'hier wird ein durchschnittlicher Korrekturfaktor ermittelt,
                            If Sys(j).E(32) > MaxMasse Then MaxMasse = Sys(j).E(32)
                            Masse = Masse + Sys(j).E(28)
                            H = H + 1
                            DurchschMasse = DurchschMasse + Sys(j).E(32) 'max masse eines stückes
                            M = M + Sys(j).E(28) * Kst(Sys(j).E(36)).Einstellung 'für abweiser
                            P = P + Sys(j).E(28) * Kst(Sys(j).E(61)).Einstellung 'reibung Masse gegen tragrollen
                            Q = Q + Sys(j).E(28)
                        End If
                        If Sys(j).Tag = "204" Then 'stau
                            Masse = Masse - Sys(j).E(23) 'staumassen abziehen
                            Staumasse = Staumasse + Sys(j).E(23)
                            Restlaenge = Restlaenge - (Sys(j).E(46) - Sys(j).E(25)) 'rechts -linke grenze
                        End If
                    End If
                Next j
                Staulänge = Sys(K).E(22) - Restlaenge
                If P > 0 And Q > 0 Then
                    DurchschFaktor = (P / Q) 'durchschnittlicher faktor für lager/rollreibung transportgut gegen tragrollen
                End If
                If M > 0 And Q > 0 Then
                    '1 ist referenz, *1 in paarung ergibt keine reibung, korrektur um *0,2
                    Mue3 = (M / Q) * 0.2 'durchschnittlicher faktor für abweiser
                End If

                Restmasse = Masse 'für abweiserberechnung und letzte notbremse, falls b_rex-hauptfenster versagt
                If Restmasse < 0 Then
                    Fehler$ = Fehler$ & Lang_Res(625) & K & Lang_Res(615) & Chr$(13) & Chr$(10)   '-Förderer: ... Stau und Abweiser mangels Transportgut nicht berechnet."
                End If
                
                'durchschnitt aus den durchschnittsmassen der einzelnen transportgutarten :-)
                If H > 0 Then DurchschMasse = DurchschMasse / H
                
                'kann passieren, wenn stau ohne transportgut da ist, eigentlich garnicht mehr
                If Masse < 0 Then Masse = 0
                'masse kennt nur bewegte masse
                
                'potentialunterschied durch steigung herausarbeiten
                If Sys(K).Rechts = False Then 'weils bei der rollenbahn nun mal anders herum ist
                    Fu = Fu + Masse * 9.81 * Sin(Sys(K).E(16) * PI / 180) 'm*g*sin alpha 'fu durch potentialunterschiede mit dieser masse berechnen
                Else
                    Fu = Fu + Masse * 9.81 * Sin(-Sys(K).E(16) * PI / 180) 'm*g*sin alpha 'fu durch potentialunterschiede mit dieser masse berechnen
                End If
                
                'massenkorrektur entsprechend der Steigung:
                Masse = Masse * Cos(Sys(K).E(16) * PI / 180) 'masse bezieht sich jetzt nur auf das zur zeit bewegte gut und auf fu durch reibung, nicht vom reversieren abh.
                DurchschMasse = DurchschMasse * Cos(Sys(K).E(16) * PI / 180) 'ditto
                
                'lager und rollreibung transportgut
                Sys(K).E(67) = 0
                Memo = (Sys(K).E(65) - Staulänge) / Sys(K).E(43) '(belege Länge -Staulänge)/tragrollenachsabstand, also bewegte rollenzahl
                If Memo > 0 Then Sys(K).E(67) = Sys(1).E(20) * (DurchschFaktor / 6.66) * Memo * (Masse / Memo + 3)
         
                'bandbiegeanteil, erst winkel, zur eindringtiefe die banddicke addieren
                If Sys(K).E(71) = 0 Then
                    P = Sys(K).E(11) ' eindringtiefe
                Else
                    P = (Sys(K).E(43) / 2) * (Sys(K).E(71) / 2) / (1.5 * Sys(K).Furein) ' + Sys(K).E(67) * 1.2) 'gedachte eindringtiefe, damit die rechnung normal weitergehen kann, nur eine seite der andruckrolle
                    'per verhältnisgleichung; Stau und Biegeleistung durch faktor 1.2 erfaßt, spart approximation und soll erst mal reichen
                End If
                Sys(K).E(12) = Atn((2 * (P + Sys(1).E(79)) / Sys(K).E(43))) * 180 / PI 'bandwinkel
                
                'mindestdurchmesserkontrolle trag/andruckrollen
                If Sys(K).E(12) > 8 And Init_B_Rex_Minddurchmkontr = 1 Then 'nur wenn der winkel >8° ist
                    If Sys(K).E(9) < Abs(Sys(1).E(86)) Then
                        'Zulässig = False
                        Fehlerwert = Fehlerwert + 100
                        Fehler$ = Fehler$ & Lang_Res(626) & Sys(1).E(86) & "mm)" & Chr$(13) & Chr$(10)  '!-Tragrollendurchmesser unterschreitet Mindestdurchmesser (
                    End If
                    If Sys(K).E(10) < Abs(Sys(1).E(86)) Then
                        'Zulässig = False
                        Fehlerwert = Fehlerwert + 100
                        Fehler$ = Fehler$ & Lang_Res(627) & Sys(1).E(86) & "mm)" & Chr$(13) & Chr$(10)  '!-Andruckrollendurchmesser unterschreitet Mindestdurchmesser (
                    End If
                End If
                
                If Sys(K).E(12) > 0 Then 'tragwinkel zwischen trag-/ und andruckrolle
                    'tragrollen
                    Sys(K).E(66) = (Sys(1).E(34) * Sys(K).E(62) * (4 + Sys(1).E(20)) * Sqr(Sys(K).E(12))) / Sys(K).E(9) * Sys(1).E(80)
                    'andruckrollen
                    Sys(K).E(66) = Sys(K).E(66) + ((Sys(1).E(34) * Sys(K).E(62) * (4 + Sys(1).E(20)) * Sqr(Sys(K).E(12))) / Sys(K).E(10) * Sys(1).E(80)) / Sys(K).E(63)
                Else
                    Sys(K).E(66) = 0 'band berührt tragrollen garnicht
                    Fehlerwert = Fehlerwert + 10
                    Fehler$ = Fehler$ & Lang_Res(669) & Chr$(13) & Chr$(10) '!-Kein Kontakt Band-Tragrollen
                End If
                
                If Restlaenge = 0 Then Restlaenge = 0.1 'falls staulänge mit trägerlänge übereinstimmt
                
                
                Sys(K).Fusteig = (Fu + Sys(K).E(67)) / Restlaenge 'steigung durch lager/rollreibung und potentialunterschied auf nichtgestauter strecke
                Sys(K).Fusteig = Sys(K).Fusteig + (Sys(K).E(66) / Sys(K).E(22) * Restlaenge) / Restlaenge '+ anteilig biegeleistung ohne die der staustrecke
                
                'Sys(K).Fusteig = (Fu + Sys(K).E(67)) / Restlaenge 'steigung durch lager/rollreibung und potentialunterschied auf nichtgestauter strecke
                'Sys(K).Fusteig = Sys(K).Fusteig + Sys(K).E(66) / Sys(K).E(22) 'steigung durch bandbiegung am gesamten träger
                'Fusteig = Sys(K).E(66) / Sys(K).E(22) 'noch beim stau dazurechnen, bandbiegung auch hier
                'Fusteig1 = (Sys(K).E(66) + Sys(K).E(67)) / Sys(K).E(22) 'anteil für freie umfangskraft
                Fu = Fu + Sys(K).Fusteig * Restlaenge
                   
                'trägerRestlaenge ermitteln
                
                'm*g*cosa=m*g*Mue*sina'transportgut rutscht? 0,22 grundsätzlich voraussetzen
                If Tan(Abs(Sys(K).E(16)) * PI / 180) > 0.22 * DurchschFaktor And Tan(Abs(Sys(K).E(16)) * PI / 180) <= 0.22 * DurchschFaktor * 1.4 Then '1.4 Haftreibungsverhältnis zu gleitreibungsverhältnis
                    Fehler$ = Fehler$ & Lang_Res(629) & Chr$(13) & Chr$(10)  '-Das Transportgut könnte auf der Rollenbahn rutschen
                    'Fehlerwert = Fehlerwert + 10
                End If
                If Tan(Abs(Sys(K).E(16)) * PI / 180) > 0.22 * DurchschFaktor * 1.4 Then '1.4 Haftreibungsverhältnis zu gleitreibungsverhältnis
                    'Zulässig = False
                    'Fehlerwert = Fehlerwert + 100
                    Fehler$ = Fehler$ & Lang_Res(630) & Chr$(13) & Chr$(10)  '!-Das Transportgut wird auf der Rollenbahn rutschen
                End If
                
                'für stau und kontrollrechnung "mitnahme" unten
                mue = Sys(1).E(77) 'andere seite Mue
                If Sys(K).E(42) = 1 Then mue = Sys(1).E(78)
                
                'If mue < 0 Then Fehler$ = Fehler$ & Lang_Res(631) & Chr$(13) & Chr$(10)  '-Reibungszahl Band-Tragrollen geschätzt
                'mue = mue 'voraussetzung:tragrollen aus stahl
                'If mue < 0.35 Then mue = ((0.35 - mue) / 2 + mue) 'alterungskorrektur
                
                mue = ReibungszahlBerechnung(mue, 1, 0, False)
                
                '***in die kontrollrechnungen aufnehmen und raus hier (einfach also am ende und nicht jedesmal, ist ja kein fehler
                If Staumasse > 0 Then
                    Memo = (9.81 * Staumasse * Cos(Sys(K).E(16) * PI / 180)) * 0.22 * DurchschFaktor 'würde bei reibung tragrolle gegen transportgut entstehen
                    'memo kraft durch stau transportgut-tragrollen
                    'memo zum abschätzen der furein/furaus für berechnung reibung tragrollen gegen band
                    'und zum vergleich mit daraus folgendem ergebnis memo1
                    'Mue finden
                    Memo1 = (Sys(K).Furein + 2 * Sys(K).Furein) / 2 'abschätzung mittlere trumkraft unter der rollenbahn, iteration verhindern
                    Memo1 = Sin(Sys(K).E(12) / 180 * PI) * Memo1 * 2 ' anpreßkraft an eine andruckrolle aus beiden trums
                    If Sys(K).E(71) > 0 Then Memo1 = Sys(K).E(71) 'bei federbelasteten andruckrollen
                    Memo1 = mue * ((Staulänge) / Sys(K).E(43) - 1) / Sys(K).E(63) * Memo1
                    'memo1 ist jetzt ne kraft, staulänge unabhängig von laufrichtung
                    If Memo1 < 0 Then Memo1 = 0 'falls negative eindringtiefe gewählt wurde oder furein negativ reinkommt
                    If Memo1 > Memo Then
                        Fehler$ = Fehler$ & Lang_Res(632) & Lang_Res(633) & Chr$(13) & Chr$(10)    '-Stau ...: Tragrollen reiben gegen das Transportgut
                    Else
                        Fehler$ = Fehler$ & Lang_Res(632) & Lang_Res(634) & Chr$(13) & Chr$(10)    '-Stau ...: Tragrollen reiben gegen das Band
                    End If
                End If
                
                'anzahl der andruckrollen im staubereich geht, da sie genauso viele kräfte wie die meist mehr tragrollen aufnehmen
                'memo oder memo1, wer kleiner ist, da rutscht es
                'stau:
                'memo1 größer: nach länge berechnen, tragrolle reibt gegen transportgut, mit memo weiterrechnen
                'memo größer: band reibt gegen tragrolle, mit memo1 wird weitergerechnet
                'anhand anteil staumasse an gesamtmasse auf die staulänge verteilen
                
                'bis hierher für rollenbahn gültig
                For j = 9 To Maxelementindex
                    If Sys(j).Zugehoerigkeit = K Then
                        'staumasse ist nicht vorab definiert
                        If Sys(j).Tag = "204" Then  'vom reversieren noch nicht abhängig, stau
                            If Memo1 >= Memo Then
                                'die anteilige staumasse für diesen stau
                                Sys(j).E(50) = Memo / Staumasse * Sys(j).E(23) 'rauf oder runter egal
                            Else
                                Sys(j).E(50) = Memo1 / Staumasse * Sys(j).E(23) 'sieht genauso aus
                                'berechnung unter memo1, wird proportional zur masse unter den staus aufgeteilt
                            End If
                            If Sys(j).E(46) = Sys(j).E(25) Then
                                Fehler$ = Fehler$ & Lang_Res(706) & Chr$(13) & Chr$(10)
                            Else
                                Sys(j).E(50) = Sys(j).E(50) + Sys(K).E(66) / Sys(K).E(22) * (Sys(j).E(46) - Sys(j).E(25))  'biegeleistung anteilig über die staulänge
                                Fu = Fu + Sys(j).E(50)
                            End If
                        End If
                        If Sys(j).Tag = "205" Then 'abweiser
                            If Restmasse > MaxMasse Then 'ist überhaupt nichtgestaute masse übrig
                                Mue3 = Mue3 * Kst(Sys(j).E(15)).Einstellung 'reibungszahl gegen abweiser
                                '0.2 bringt faktor von ca 1 auf reibungszahl zurück
                                'merk ist die kraft, mit der das Fördergut gegen abweiser drückt, wenn ein stau wäre, nur reibung!, druck durch schwerkraft entfällt
                                Merk = 9.81 * MaxMasse * Cos(Sys(K).E(16) * PI / 180) * 0.22 * DurchschFaktor 'auf tragrollen rutscht das gut rechnerisch oben, cos für steigung, aus reibung
                                'der einfacheit halber: gut rutscht immer auf tragrollen
                                If Mue3 * 1.2 > Tan((90 - Sys(j).E(26)) * PI / 180) Then '1.2 selbsteingebauter angstfaktor
                                    Fehlerwert = Fehlerwert + 10
                                    Fehler$ = Fehler$ & Lang_Res(621) & j & Lang_Res(622) & Chr$(13) & Chr$(10)  '-Der Abweiser ... läßt wahrscheinlich kein Transportgut durch (Berechnung ähnlich Stau)
                                    'so ähnlich wie stau, nur ohne reibung band-tisch, merk von oben wird verwendet
                                Else
                                    Merk = Merk * Mue3 * Sin(Sys(j).E(26) * PI / 180) 'nur die reibung am abweiser (merk aus gewicht und reibung)
                                    'auch wenn die schwerkraft abwärts dazukommt, mit mehr als merk kann das band nicht drücken!
                                    'reibung auf tragrollen entfällt (merk wird umgewandelt)
                                End If
                                Sys(j).E(50) = Merk
                                Fu = Fu + Sys(j).E(50)
                            Else
                                Fehler$ = Fehler$ & Lang_Res(623) & j & Lang_Res(624) & Chr$(13) & Chr$(10)   '-keine Abweiserberechnung ... durchgeführt, da zuviel Masse im Stau
                                Sys(j).E(50) = 0
                            End If
                        End If
                        If Sys(j).Tag = "206" Then 'trägergebundene_Umfangskraft
                            Sys(j).E(50) = Sys(j).E(44) + Sys(K).Fusteig * (Sys(j).E(46) - Sys(j).E(25))
                            Fu = Fu + Sys(j).E(44)  'eingeleitete Kraft
                            'steigung für Spitzenlast wird unten bei der ermittlung der auflegedehnung ausgerechnet
                        End If
                    End If
                Next j
                    
                Sys(K).E(50) = Fu 'alles auf den träger
                Sys(K).Furaus = Sys(K).Furein + Fu
                Sys(K).Lraus = Sys(K).Lrein + Sys(K).E(22)
                
                'spitzenlast bei beschleunigung der transportmasse
                    If Restmasse > 0.1 And Restlaenge > 0.1 Then
                        Sys(K).E(98) = Restmasse * Sys(1).E(68) 'F = m*a
                        Sys(K).FusteigSp = Sys(K).E(98) / Restlaenge 'diese steigung durch die beschleunigung der ungestauten transportmasse
                    Else
                        Sys(K).E(98) = 0
                        Sys(K).FusteigSp = 0
                    End If
                    'und über massenträgheit der tragrollen
                    P = Sys(K).E(62) * (Sys(1).E(68) * MassentraegheitsErmittlung(K) / (Sys(K).E(9) / 1000) ^ 2) 'J*a/r^2  wenn eins fehlt, ist alles zusammen e null
                    Sys(K).FusteigSpRoll = P / Sys(K).E(22) 'in eine steigung übertragen
                    Sys(K).E(98) = Sys(K).E(98) + P
                    
                'andruck vorn/hinten als service anbieten
                P = Sys(K).E(63)
                If P > 2 Then P = 2 'mehr als 2 tragrollen können von einer anpreßrolle nicht berührt werden
                If Sys(K).E(71) > 0 Then 'anpreßkraft festgelegt
                    Sys(K).E(69) = Sys(K).E(71) / P 'tatsächliche kraft auf eine Scheibe
                    M = Sys(K).E(71) / Sys(K).E(63) 'gemittelte kraft auf eine scheibe, später noch zur berechnung, ob andruck ausreicht
                    Sys(K).E(70) = Sys(K).E(69) 'also vorne und hinten gleich
                Else
                    Sys(K).E(69) = (2 * Sin(Sys(K).E(12) / 180 * PI) * Sys(K).Furein) / P 'vorne die 2, weils ja 2 stränge sind
                    Sys(K).E(70) = (2 * Sin(Sys(K).E(12) / 180 * PI) * Sys(K).Furaus) / P
                    M = (2 * Sin(Sys(K).E(12) / 180 * PI) * Sys(K).Furein) / Sys(K).E(63) 'vorne die 2, weils ja 2 stränge sind
                End If
                
                'kontrolle, ob auch die mitnahme gewährleistet ist
                Memo = Sys(K).E(64) / Sys(K).E(43) 'rollenzahl unter einem durchschn. Transportgut
                Memo1 = Memo
                If Reversieren = True Then 'weils bei der rollenbahn nun mal anders herum ist
                    Memo = Sys(1).E(20) * (DurchschFaktor / 6.66) * Memo * (DurchschMasse / Memo + 3) + DurchschMasse * 9.81 * Sin(Sys(K).E(16) * PI / 180) 'm*g*sin alpha'+potentialunterschied
                Else
                    Memo = Sys(1).E(20) * (DurchschFaktor / 6.66) * Memo * (DurchschMasse / Memo + 3) + DurchschMasse * 9.81 * Sin(-Sys(K).E(16) * PI / 180) 'm*g*sin alpha 'fu durch potentialunterschiede mit dieser masse berechnen
                End If
                Memo = Abs(Memo) 'enthält verformungsarbeit am transportgut
                'Memo = Sys(1).E(20) * (DurchschFaktor / 6.66) * Memo * (DurchschMasse / Memo + 3) 'umfangskraft durch biegung
                Memo1 = mue * M * Memo1 'Mue Band-Tragrollen, m gemittelte kraft auf eine rolle, memo1 anzahl rollen
                If Memo1 < Memo Then
                    Fehler$ = Fehler$ & Lang_Res(662) & Chr$(13) & Chr$(10)  '-Andruckkraft Band-Tragrolle reicht für Mitnahme nicht aus
                    Fehlerwert = Fehlerwert + 10
                End If
            Case "102" 'tragrollenbahn, träger
                Fu = 0
                Masse = 0
                Staulänge = 0
                Restlaenge = Sys(K).E(22)
                MaxMasse = 0
                P = 0
                Q = 0
                M = 0
                Memo = 1
                DurchschFaktor = 1
                For j = 9 To Maxelementindex
                    If Sys(j).Zugehoerigkeit = K Then
                        If Sys(j).Tag = "201" Then
                            If Sys(j).E(32) > MaxMasse Then MaxMasse = Sys(j).E(32)
                            Masse = Masse + Sys(j).E(28)
                            P = P + Sys(j).E(28) * Kst(Sys(j).E(36)).Einstellung 'angaben zur reibung
                            M = M + Sys(j).E(28) * Kst(Sys(j).E(61)).Einstellung 'angaben zur biegeleistung
                            Q = Q + Sys(j).E(28) 'ist nicht identisch masse, weil stau abgezogen wird
                        End If
                        If Sys(j).Tag = "204" Then
                            Masse = Masse - Sys(j).E(23) 'andere staumassen abziehen
                            Staulänge = Staulänge + (Sys(j).E(46) - Sys(j).E(25))
                            Restlaenge = Restlaenge - (Sys(j).E(46) - Sys(j).E(25))
                            If Sys(K).E(41) = 1 Then Tragseite_Mue_verwendet = True '1 heißt laufseite zu den rollen, also tragseite zum transportgut
                        End If
                        If Sys(j).Tag = "205" Then
                            If Sys(K).E(41) = 1 Then Tragseite_Mue_verwendet = True '1 heißt laufseite zu den rollen, also tragseite zum transportgut
                        End If
                    End If
                Next j
                If M > 0 And Q > 0 Then
                    DurchschFaktor = (M / Q) 'durchschnittlicher faktor für reibung gegen tragrollen
                End If
                
                Restmasse = Masse 'nur für abweiserberechnung
                
                If Restlaenge <= 0 Then Restlaenge = 0.1 'falls staulänge mit trägerlänge übereinstimmt
                
                If Restmasse < 0 Then
                    Fehler$ = Fehler$ & Lang_Res(625) & K & Lang_Res(615) & Chr$(13) & Chr$(10)   '-Förderer: ... Stau und Abweiser mangels Transportgut nicht berechnet.
                End If
                If Masse < 0 Then Masse = 0 'kann passieren, wenn stau ohne transportgut da ist
                                
                If Sys(K).Rechts = True Then
                    Fu = Fu + Masse * 9.81 * Sin(Sys(K).E(16) * PI / 180) 'm*g*sin alpha 'fu durch potentialunterschiede mit dieser masse berechnen
                Else
                    Fu = Fu + Masse * 9.81 * Sin(-Sys(K).E(16) * PI / 180) 'm*g*sin alpha 'fu durch potentialunterschiede mit dieser masse berechnen
                End If
                
                'bandmasse ist bei potentialunterschied nicht dabei, denn die befindet sich im kreislauf
                Masse = Masse + (Sys(K).E(22) / 1000 * Sys(1).E(34) / 1000) * Sys(1).E(81) 'bandmasse dazuzählen
                
                'massenkorrektur entsprechend der Steigung, wir später zur biegeleistung verarbeitet
                Masse = Masse * Cos(Sys(K).E(16) * PI / 180) 'masse bezieht sich jetzt nur auf das zur zeit bewegtes gut und auf fu durch reibung entspr der neigung, nicht vom reversieren abh.
                
                'stau und abweiser berechnen, Mue1 reibwert zwischen band und transportgut
                If Sys(K).E(41) = 1 Then
                    Mue1 = Sys(1).E(77) 'tragseite nach oben
                Else
                    Mue1 = Sys(1).E(78) 'laufseite nach oben
                End If
         
                'P und Q bei dieser trägerberechnung nicht mehr verändern
                Mue2 = Mue1 'Mue1 möglichst hoch, wenn's Gut rutschen muß wegen stau, Mue2 mögl. niedrig, wenn's Gut rutscht wegen Steigung
                Mue1 = ReibungszahlBerechnung(Mue1, 1, 0, False)
                Mue2 = ReibungszahlBerechnung(Mue2, 2, 0, False)

                
                'm*g*cosa=m*g*Mue*sina'transportgut rutscht?
                If Sys(K).E(16) <> 0 Then 'steigungswinkel
                    If Mue2 <> 0 Then
                        If Tan(Abs(Sys(K).E(16)) * PI / 180) > Mue2 And Tan(Abs(Sys(K).E(16)) * PI / 180) <= Mue2 * 1.4 Then '1.3 Haftreibungsverhältnis zu gleitreibungsverhältnis
                            Fehler$ = Fehler$ & Lang_Res(619) & Chr$(13) & Chr$(10)  '-Das Transportgut könnte auf dem Band rutschen
                            'Fehlerwert = Fehlerwert + 10
                        End If
                        If Tan(Abs(Sys(K).E(16)) * PI / 180) > Mue2 * 1.4 Then '1.4 Haftreibungsverhältnis zu gleitreibungsverhältnis
                            'Fehlerwert = Fehlerwert + 100
                            'Zulässig = False
                            Fehler$ = Fehler$ & Lang_Res(620) & Chr$(13) & Chr$(10)  '!-Das Transportgut wird auf dem Band rutschen
                        End If
                    End If
                    Tragseite_Mue_verwendet = True
                End If
                
                'rollreibung durch transportgut und bandgewicht
                '67 = lager/Rollreibung
                '20 = Bandgeschwindigkeit
                '22 = Foederlaenge
                '43 = Tragrollenachsabstand
                'DurchschFaktor = durchschnittlicher Faktor fuer tragrollenreibung
                
                Sys(K).E(67) = 0
                Memo = Sys(K).E(22) / Sys(K).E(43) 'bewegte rollenzahl
                Sys(K).E(67) = Sys(1).E(20) * (DurchschFaktor / 6.66) * Memo * (Masse / Memo + 3)
                
                Fu = Fu + Sys(K).E(67)
                Sys(K).Fusteig = Fu / Restlaenge 'steigung ermitteln
                
                For j = 9 To Maxelementindex
                    If Sys(j).Zugehoerigkeit = K Then
                        If Sys(j).Tag = "204" Then  'vom reversieren noch nicht abhängig, stau
                            Sys(j).E(50) = 9.81 * (Sys(j).E(23) * Cos(Sys(K).E(16) * PI / 180)) * Mue1  'reibung band-tisch ist in der steigung enthalten
                            Fu = Fu + Sys(j).E(50)
                        End If
                        If Sys(j).Tag = "205" Then 'abweiser
                            If Restmasse > MaxMasse Then 'ist überhaupt nichtgestaute masse übrig
                                Mue3 = Memo * (Kst(Sys(j).E(15)).Einstellung * 0.2) 'reibungszahl gegen abweiser
                                'beschaffenheit transportgut in Mue1 durch p/q enthalten
                                'merk ist die kraft, mit der das Fördergut gegen abweiser drückt, wenn ein stau wäre, nur reibung!, druck durch schwerkraft entfällt
                                Merk = 9.81 * MaxMasse * Cos(Sys(K).E(16) * PI / 180) * (Mue1) 'auf tisch rutscht das gut rechnerisch oben, cos für steigung, aus reibung
                                If Mue3 * 1.2 > Tan((90 - Sys(j).E(26)) * PI / 180) Then '1.2 selbsteingebauter angstfaktor
                                    Fehlerwert = Fehlerwert + 10
                                    Fehler$ = Fehler$ & Lang_Res(621) & j & Lang_Res(622) & Chr$(13) & Chr$(10)   '-Der Abweiser ... läßt wahrscheinlich kein Transportgut durch (Berechnung ähnlich Stau)
                                    'so ähnlich wie stau, nur ohne reibung band-tisch, merk von oben wird verwendet
                                Else
                                    Merk = Merk * Mue3 * Sin(Sys(j).E(26) * PI / 180) 'nur die reibung am abweiser (merk aus gewicht und reibung)
                                    'auch wenn die schwerkraft abwärts dazukommt, mit mehr als merk kann das band nicht drücken!
                                End If
                                Sys(j).E(50) = Merk
                                Fu = Fu + Sys(j).E(50)
                            Else
                                Fehler$ = Fehler$ & Lang_Res(623) & j & Lang_Res(624) & Chr$(13) & Chr$(10) '-keine Abweiserberechnung ... durchgeführt, da zuviel Masse im Stau
                                Sys(j).E(50) = 0
                            End If
                        End If
                        If Sys(j).Tag = "206" Then 'trägergebundene_Umfangskraft
                            Sys(j).E(50) = Sys(j).E(44) + Sys(K).Fusteig * (Sys(j).E(46) - Sys(j).E(25))
                            Fu = Fu + Sys(j).E(44) 'eingeleitete Kraft
                            'steigung für Spitzenlast wird unten bei der ermittlung der auflegedehnung ausgerechnet
                        End If
                    End If
                Next j
                Sys(K).E(50) = Fu 'alles auf den träger
                Sys(K).Furaus = Sys(K).Furein + Fu
                Sys(K).Lraus = Sys(K).Lrein + Sys(K).E(22)
                
                'spitzenlast bei beschleunigung der transportmasse
                    If Restmasse > 0.1 And Restlaenge > 0.1 Then
                        Sys(K).E(98) = Restmasse * Sys(1).E(68) 'F = m*a
                        Sys(K).FusteigSp = Sys(K).E(98) / Restlaenge 'diese steigung durch die beschleunigung der ungestauten transportmasse
                    Else
                        Sys(K).E(98) = 0
                        Sys(K).FusteigSp = 0
                    End If
                    'und über massenträgheit der tragrollen
                    P = Sys(K).E(62) * (Sys(1).E(68) * MassentraegheitsErmittlung(K) / (Sys(K).E(9) / 2000) ^ 2) 'J*a/r^2  wenn eins fehlt, ist alles zusammen e null
                    Sys(K).FusteigSpRoll = P / Sys(K).E(22) 'in eine steigung übertragen
                    Sys(K).E(98) = Sys(K).E(98) + P

            Case "104" 'freie_Umfangskraft,träger
                j = K 'nach trägergebundene_Umfangskraft suchen
                Fu = Sys(K).E(44)
                Sys(K).Fusteig = Sys(K).E(44) / Sys(K).E(22) 'steigung ermitteln
                For j = 9 To Maxelementindex
                    If Sys(j).Zugehoerigkeit = K Then
                        If Sys(j).Tag = "206" Then 'trägergebundene_Umfangskraft
                            Sys(j).E(50) = Sys(j).E(44) + Sys(K).Fusteig * (Sys(j).E(46) - Sys(j).E(25))
                            Fu = Fu + Sys(j).E(44)  'eingeleitete Kraft
                        End If
                    End If
                Next j
                Sys(K).FusteigSp = Sys(K).E(98) / Sys(K).E(22)
                Sys(K).E(50) = Fu 'alles auf den träger
                Sys(K).Furaus = Sys(K).Furein + Fu
                Sys(K).Lraus = Sys(K).Lrein + Sys(K).E(22)
        End Select
      
        Fuletztes = Sys(K).Furaus
        If Fumax < Sys(K).Furaus Then Fumax = Sys(K).Furaus 'speichert die höchste Kraft
        If Fumin > Sys(K).Furaus Then
            Fumin = Sys(K).Furaus
        End If
        
        'wurde gebraucht, ist aber nicht in der Datenbank, dann darauf hinweisen
        'würde man die selten gebrauchte Tragseitenreibungszahl zur Pflicht machen, ginge die automatische Bandauslegung schief
        If Tragseite_Mue_verwendet = True And Sys(1).E(77) = 0 Then
            Fehler$ = Fehler$ & Lang_Res(702) & K & Lang_Res(703) & Chr$(13) & Chr$(10)   '- Element ... Reibungszahl Tragseite/Antriebsseite des Bandes wurde gebraucht, aber nicht gefunden
        End If
        
        'GoSub Wellenbelastung_abgleichen
        If K = ScheibeFedGew Then Call Wellenbelastung_abgleichen(K, Memo, Q, M, i, IaltK, P) 'falls die antriebsscheibe betroffen ist
        
        'einstellungen für den nächsten durchlauf, nicht wieder zum alten element zurück
        'k ist die naechste, ialtk die letzte scheibe
        If Sys(K).Verb(1, 1) = IaltK Then 'voreinstellungen für neuen durchlauf
            IaltK = K
            K = Sys(K).Verb(2, 1)
        Else
            IaltK = K
            K = Sys(K).Verb(1, 1)
        End If
        
        'eigenfrequenz/erregerfrequenzberechnung
            'hier ist die reihenfolge der elemente bekannt, auch furaus in die richtige richtung,
            'drum trumeigenfrequenzberechnung ab ialtk zu k (neu) hier,
            'eintragen in beide, bewertung findet unter kontrollrechnung statt,
            'wo das mit stoerfrequenzen von scheiben abgeglichen wird
            Call Eigenfrequenzberechnung(K, IaltK, False)
        
            
    Loop Until K = Antriebsscheibe Or K >= Maxelementindex + 1 'einmal rum oder es ist was schiefgegangen
    

    
    'letzte entfernung zur antriebsscheibe protokollieren
    If Sys(Antriebsscheibe).Verb(1, 1) = IaltK Then
        Sys(Antriebsscheibe).Lrein = Sys(Antriebsscheibe).Verb(1, 3) + Sys(IaltK).Lraus
    Else
        Sys(Antriebsscheibe).Lrein = Sys(Antriebsscheibe).Verb(2, 3) + Sys(IaltK).Lraus
    End If
    Sys(Antriebsscheibe).Lraus = Sys(Antriebsscheibe).Lrein + PI * Sys(Antriebsscheibe).E(2) / 360 * Sys(Antriebsscheibe).E(13)


    'kontrollrechnungen aus den iterationen raushalten = zeitgewinn
    
    'hier noch die berechnung der antriebsscheibe, wenn die nicht gegeben, einbeziehen in länge, schlupfausgleich
        Sys(Antriebsscheibe).Furaus = Fuerstes 'enthält fliehkraft
        Sys(Antriebsscheibe).E(50) = Fuletztes - Fuerstes
        
    Call Eigenfrequenzberechnung(K, Startelement, True) 'die erste strecke kommt sonst zu kurz, furaus mus erst da sein
        
        
        'Spitzenlast (beschleunigung der eigenen masse):
            Sys(Antriebsscheibe).E(98) = Sys(1).E(68) * MassentraegheitsErmittlung(Antriebsscheibe) / (Sys(Antriebsscheibe).E(2) / 2000) ^ 2 'J*a/r^2  wenn eins fehlt, ist alles zusammen e null
            
        If Fuerstes < 0 Then Sys(Antriebsscheibe).E(50) = Fuletztes 'bei neg. Dehnung im Leertrum
        Sys(Antriebsscheibe).Furein = Fuletztes 'enthält fliehkraft
    
    'nur noch bei antriebsscheibe, wenn furaus ermittelt:
        If K = ScheibeFedGew Then Call Wellenbelastung_abgleichen(K, Memo, Q, M, i, IaltK, P) 'falls die antriebsscheibe betroffen ist
    
    
    'alle furein/furaus Muessen bekannt sein
    Call Auflegedehnung_ermitteln(False) 'ohne zeichnen, das kommt erst zum schluß
        
    'kurve verschieben, um vorgabedehnung an spannscheibe zu treffen
    'ab ermittlung der auflegedehnung steht fest, um wieviel verschoben werden muss
    'die veränderung geht voll in die kontrollrechnungen ein
        
        'If Auflegemodus = 4 Then 'feder/gewicht
        '    Q = (ScheibeFedGewNormalFu - AuflTrumKraft)
        '    For J = 10 To Maxelementindex 'ganze normalkurve heben/senken und dann damit rechnen
        '        Sys(J).Furaus = Sys(J).Furaus - Q
        '        Sys(J).Furein = Sys(J).Furein - Q
        '    Next J
        '    Fumin = Fumin - Q
        '    Fumax = Fumax - Q
        'End If
    
    'biegeleistung, geht nur in die antriebsscheibe ein, die bringt auch gleich selbst die kraft auf, geht nicht ins band
    Sys(Antriebsscheibe).E(66) = (Sys(1).E(34) * (4 + Sys(1).E(20)) * Sqr(Sys(Antriebsscheibe).E(13))) / Sys(Antriebsscheibe).E(2) * Sys(1).E(80)
    '******!!!!!!!spitzenlast in der antriebsscheibe noch ergänzen, die muß mit der leistung mitgewuppt werden
    Sys(Antriebsscheibe).E(50) = Sys(Antriebsscheibe).E(50) + Sys(Antriebsscheibe).E(66)
           
    'fu, M, P
    Sys(Antriebsscheibe).E(17) = Sys(Antriebsscheibe).E(50)
    Sys(Antriebsscheibe).E(18) = Sys(Antriebsscheibe).E(50) * Sys(Antriebsscheibe).E(2) / 2000 'drehmoment Antriebsscheibe
    Sys(Antriebsscheibe).E(19) = Sys(Antriebsscheibe).E(50) * Sys(1).E(20) / 1000 'leistung antriebsscheibe
    
    'fu bei spitzenlast+
    'biegeleistung+
    'massebeschleunigung der antriebsscheibe+
    'massebeschleunigung dieses bandstückchens 'eigentlich pillepalle, aber so hab ich wenigstens ne passende antwort auf blöde fragen
    'nur vorübergehend in leistung gespeichert, weil die variable gerade frei ist
    Sys(Antriebsscheibe).E(118) = FuletztesSp - Fuerstes + Sys(Antriebsscheibe).E(66) + Sys(Antriebsscheibe).E(98) + (Sys(Antriebsscheibe).Lraus - Sys(Antriebsscheibe).Lrein) * Sys(1).FusteigSp
    Sys(Antriebsscheibe).E(107) = Sys(Antriebsscheibe).E(118) * Sys(Antriebsscheibe).E(2) / 2000 'fu zu drehmoment
    Sys(Antriebsscheibe).E(102) = Sys(Antriebsscheibe).E(118) * Sys(1).E(20) / 1000 'fu zu leistung
    

End Sub

Private Sub Wellenbelastung_abgleichen(K As Integer, Memo As Double, ByRef Q As Double, ByRef M As Double, ByRef i As Integer, ByRef IaltK As Integer, ByRef P As Double)
        'son zirkus, weil die stellung zum gewicht geklärt werden muß
        'das wird nur wichtig, wenn die scheibe auch noch umfangskräfte hat
        'sonst haben die beiden trums e etwa gleiche kräfte
            Memo = Sys(K).E(56) 'winkel zum höheren element
            Merk = Sys(K).E(57)
            Q = Sys(K).Furein ' - Fliehkraftsumme 'immer kleiner
            M = Sys(K).Furaus ' - Fliehkraftsumme 'immer größer
            
            'trum passend
            i = Sys(K).Verb(1, 1)
            If Sys(K).Verb(2, 1) > Sys(K).Verb(1, 1) Then i = Sys(K).Verb(2, 1)
            
            'alte variante, ausgetauscht sep 2009
                'i enthält höheres der beiden verb. elemente
                If i = IaltK Then 'das band kommt vom höh. element
                    P = Abs(Cos(Memo / 180 * PI) * Q) + Abs(Cos(Merk / 180 * PI) * M) 'sin bei 90° = 1, cos = 0
                Else 'geht zum höheren element
                    P = Abs(Cos(Memo / 180 * PI) * M) + Abs(Cos(Merk / 180 * PI) * Q)
                End If
            
            'weg damit ab 2011, wieder die variante 2009 inkraft gesetzt
                FwScheibeFedGew = P - Sys(K).E(51) 'merken, um es mit sys(...).e(54) bei der iteration zu vergleichen
End Sub

Private Sub Eigenfrequenzberechnung(ByVal K As Integer, ByVal IaltK As Integer, Antriebsscheibendurchlauf As Boolean)
Dim ZW As Boolean
Dim Furaus


Furaus = Sys(IaltK).Furaus
If Zweischeiben = True And Antriebsscheibendurchlauf = True Then
    ZW = True
    Furaus = Sys(Antriebsscheibe).Furaus
End If

    If Sys(1).E(81) = 0 Then Exit Sub
    If Furaus <= 0 Then Exit Sub
    'problem: bei zweischeibe sind beide zweimal miteinander verbunden,
    'deswegen werden beide ergebnisse immer in nur eine verbindung eingetragen
    
    If Sys(K).Verb(1, 1) = IaltK And ZW = False Then  'die erste der beiden verbindungen
        'f = 1/l*sqr(F/(4*G)) laenge in m, drum statt 1 die 1000
        '34 bandbreite
        '81 bandmasse in kg/m^2
        'verb 1,3 und 2,3 sind die trumlaengen zwischen den elementen
        If Sys(K).Verb(1, 3) > 0 Then 'vielleicht gibts da kein frei schwingendes trum
            Sys(K).Verb(1, 4) = 1000 / Sys(K).Verb(1, 3) * Sqr(Furaus / (4 * Sys(1).E(81) * Sys(1).E(34) / 1000)) 'G muss erst auf kg/m umgerechnet werden, auch wenns dann eiheitenmaessig nicht hinkommt.
            Sys(K).E(112) = Sys(K).Verb(1, 4)
            If Sys(IaltK).Verb(1, 1) = K Then 'auch auf der anderen seite das ergebnis eintragen
                Sys(IaltK).Verb(1, 4) = Sys(K).Verb(1, 4)
                Sys(IaltK).E(112) = Sys(K).Verb(1, 4) 'nur zum merken, damit der kunde auch was sieht
            Else
                Sys(IaltK).Verb(2, 4) = Sys(K).Verb(1, 4)
                Sys(IaltK).E(113) = Sys(K).Verb(1, 4)
            End If
        End If
    Else 'die zweite
        If Sys(K).Verb(2, 3) > 0 Then
            Sys(K).Verb(2, 4) = 1000 / Sys(K).Verb(2, 3) * Sqr(Furaus / (4 * Sys(1).E(81) * Sys(1).E(34) / 1000)) 'G muss erst auf kg/m umgerechnet werden, auch wenns dann eiheitenmaessig nicht hinkommt.
            Sys(K).E(113) = Sys(K).Verb(2, 4)
            If Sys(IaltK).Verb(1, 1) = K And ZW = False Then 'auch auf der anderen seite das ergebnis eintragen
                Sys(IaltK).Verb(1, 4) = Sys(K).Verb(2, 4)
                Sys(IaltK).E(112) = Sys(K).Verb(2, 4)
            Else
                Sys(IaltK).Verb(2, 4) = Sys(K).Verb(2, 4)
                Sys(IaltK).E(113) = Sys(K).Verb(2, 4)
            End If
        End If
    End If
'Exit Sub
'Errorhandler:
'Stop
End Sub
Public Sub Grafik(ByVal Drucken As Boolean)
'wird immer erst am ende der rechnung ausgerufen

Dim Tickx As Double, Ticky As Double, L As Double, i As Double, j As Double, Fuvolumen As Double
Dim K As Double, P As Double, Q As Double, H As Double, B As Double
Dim V As Double
Dim Ymax, Dumy As Double
Dim Dummy$
Dim Strichbreite As Integer
Dim BL As Boolean

If Vollstaendig = False Then
    B_Rex.FuKurve.Cls
    Exit Sub
End If

On Error Resume Next
Ymax = Fumax
If FumaxSp > Ymax Then Ymax = FumaxSp

If Drucken = False Then
    Set Destination = B_Rex.FuKurve
    B_Rex.FuKurve.FontSize = 8
    Strichbreite = 2
Else
    Strichbreite = 10
End If

    L = Sys(Antriebsscheibe).Lraus 'ist die bandlänge
    
    If Fumin > 0 Then Fumin = 0
    
    Destination.ForeColor = QBColor(0)
    Fuvolumen = 0
    
    Destination.Scale (0, Ymax)-(L, Fumin) 'nur zum vermessen der buchstaben
    
    'linke grenze
    H = Destination.TextWidth(Str(Int(Ymax)) & "----")
    If Destination.TextWidth(Str(Int(Fumin)) & "----") > H Then
        H = Destination.TextWidth(Str(Int(Fumin)) & "----")
    End If
    
    'rechte grenze ermitteln
    j = Destination.TextWidth("000.00---")
    
    'untere grenze ermitteln,maße für druck festhalten
        P = Abs(Destination.TextHeight("0"))
        FuScaleX1 = -H
        If Schwingungen_berechnen = True Then
            FuScaleX2 = L + 1.5 * j '1 vorher, jetzt ein bisschen platz fuer die eigenfrequenzskala
        Else
            FuScaleX2 = L + 1 * j '1 vorher, jetzt ein bisschen platz fuer die eigenfrequenzskala
        End If
        FuScaleY1 = Ymax + 1.4 * P
        FuScaleY2 = Fumin - 3 * P
    
    Dim ObenLinksX As Double
    Dim ObenLinksY As Double
    Dim UntenRechtsX As Double
    Dim UntenRechtsY As Double
    
    'die kurve
    
    If Drucken = True Then BL = True
    
    If BL = True Then
        'sonst reicht der platz nicht
        'FuScaleY2 = Fumin - (Fumax - Fumin) / 2
        'FuScaleY1 = Fumax + (Fumax - Fumin) / 2
        
        'aus der mathematischen 2-Punkte - Form einer Geraden
            'X = ((X2 - X1) / (Y2 - Y1)) * (Y - Y1) + X1
            'Y = ((Y2 - Y1) / (X2 - X1)) * (X - X1) + Y1
            'idee: 2 punkte habe ich, die koordinaten des dritten werden berechenbar
            
        B = (FuScaleY2 - FuScaleY1) / (Drucky + 800 - Drucky)
        Destination.ScaleTop = B * (0 - Drucky) + FuScaleY1 'mit dem selben maßstab die y-achse. oben, falls er nicht oben begonnen hat
        Destination.ScaleHeight = B * 2960 ' - Destination.ScaleTop  'mit dem selben maßstab die y-achse
        
        B = (FuScaleX2 - FuScaleX1) / (1340 - 235) '1340 statt 1390 (235 statt 195) ist unsauber, aber als provisorium erstaml ausreichend
        Destination.ScaleLeft = B * (0 - 235) + FuScaleX1
        Destination.ScaleWidth = B * 2100
        
        H = Destination.TextWidth("000.00---")
        P = Abs(Destination.TextHeight("0"))

    Else 'herkömmlich
        Destination.Scale (FuScaleX1, FuScaleY1)-(FuScaleX2, FuScaleY2) 'nur zum vermessen der buchstaben
    End If
    
    'elemente kenntlich machen
    'stau, freie Umfangskraft
    Destination.DrawWidth = 1
    'L = Sys(Antriebsscheibe).Lraus 'ist die bandlänge, zur vorsicht nochmal
    
    'elemente in der Fu-Kurve kennzeichnen
    For i = 9 To Maxelementindex
        If Sys(i).Element <> "" And left(Sys(i).Tag, 1) <> "2" Then
            Destination.Line (L - Sys(i).Lrein, Fumin + 1.4 * P)-(L - Sys(i).Lraus, Fumin), &H80000003, BF
            Destination.DrawStyle = 2
            Destination.Line (L - Sys(i).Lrein, Fumin)-(L - Sys(i).Lrein, Ymax)
            Destination.Line (L - Sys(i).Lraus, Fumin)-(L - Sys(i).Lraus, Ymax)
            Destination.DrawStyle = 0
            Destination.CurrentX = L - (Sys(i).Lraus + Sys(i).Lrein) / 2 - Destination.TextWidth(i) / 2
            Destination.CurrentY = Fumin + 1.2 * P
            If Drucken = False Then Destination.ForeColor = QBColor(15)
            If Sys(i).Lraus - Sys(i).Lrein > Destination.TextWidth(CStr(i)) Then
                Destination.Print CStr(i) 'nur wenns reinpaßt
            End If
            Destination.ForeColor = QBColor(0)
        End If
        Destination.DrawStyle = 0
        If Sys(i).Tag = "204" Then
            Destination.Line (Sys(i).Lrein, Fumin + 1.4 * P)-(Sys(i).Lraus, Fumin + 1.8 * P), vbRed, BF
            Destination.DrawStyle = 2
            Destination.Line (Sys(i).Lrein, Fumin + 1.4 * P)-(Sys(i).Lrein, Sys(i).Furein), vbRed
            Destination.Line (Sys(i).Lraus, Fumin + 1.4 * P)-(Sys(i).Lraus, Sys(i).Furaus), vbRed
        End If
        If Sys(i).Tag = "206" Then
            Destination.Line (Sys(i).Lrein, Fumin + 1.4 * P)-(Sys(i).Lraus, Fumin + 1.8 * P), vbBlue, BF
            Destination.DrawStyle = 2
            Destination.Line (Sys(i).Lrein, Fumin + 1.4 * P)-(Sys(i).Lrein, Sys(i).Furein), vbBlue
            Destination.Line (Sys(i).Lraus, Fumin + 1.4 * P)-(Sys(i).Lraus, Sys(i).Furaus), vbBlue
        End If
    Next i
    
    Destination.DrawWidth = Strichbreite
    Destination.Line (0, Fumin)-(L, Fumin) 'x-Achse
    Destination.Line (0, Fumin)-(0, Ymax) 'y-Achse 1
    Destination.Line (L, Fumin)-(L, Ymax) 'y Achse 2
    
    'x-Achse beschriften
        Tickx = Int(L / 5)
        If Tickx = 0 Then Tickx = 1
        Dummy$ = CStr(Tickx)
        If L / Tickx > 7 Then Tickx = Tickx * 2
        For i = 2 To Len(Dummy$)
            Mid(Dummy$, i, 1) = "0"
        Next i
        Tickx = Val(Dummy$)
        K = 0
        Do
            Destination.Line (K, Fumin)-(K, Fumin - P / 2)
            Destination.CurrentY = Fumin - P / 2
            Destination.CurrentX = K - Destination.TextWidth(K) / 2
            Destination.Print K
            K = K + Tickx
        Loop Until K > L
    
    'y-Achse beschriften
        If Fumin >= 0 Then
            Ticky = Int(Ymax / 5)
        Else
            Ticky = Int((Ymax + Abs(Fumin)) / 5)
        End If
        If Ticky = 0 Then Ticky = 1
        Dummy$ = CStr(Ticky) 'cstr ohne voranstehendes leerzeichen
        For i = 2 To Len(Dummy$)
            Mid(Dummy$, i, 1) = "0"
        Next i
        Ticky = Val(Dummy$)
        If Ymax / Ticky > 7 Then Ticky = Ticky * 2
        K = 0
        Do While K > Fumin
            K = K - Ticky
        Loop
        If K < 0 Then K = K + Ticky
        Q = Destination.TextWidth("0")
        
        Destination.DrawWidth = 1
        Do
            Destination.Line (0, K)-(-Q, K) 'fu
            'Destination.Line (L, K)-(L + Q, K) 'dehn
            Destination.DrawStyle = 2
            Destination.Line (0, K)-(L, K) 'dehn
            Destination.DrawStyle = 0
            Destination.CurrentX = -Destination.TextWidth(Str(K) & "  ")
            Destination.CurrentY = K - Destination.TextHeight(K) / 2
            Destination.Print K 'fu
            Destination.CurrentX = L + 1.5 * Q
            Destination.CurrentY = K - Destination.TextHeight(K) / 2
            Destination.Print Format(K * 2 / (SystemTyp.Kraftdehnung * Sys(1).E(34)), "#####0.00") 'auflegedehnung
            K = K + Ticky
        Loop Until K > Ymax
        
        'fehlerwerte eintragen
                Destination.Line (L, 0)-(L + Q, (Fehlerverlauf(1, 0) + Fehlerverlauf(1, 1)) / 2), vbRed, BF 'QBColor(12), BF 'dehn am anfang sowieso rot
            
            i = 1
            Do Until i >= Rechengenauigkeit - 2 Or Fehlerverlauf(1, i) >= Ymax
                'If Drucken = False Then
                    K = B_Rex_Yellow '&H80FFFF '14 'gelb
                    If Fehlerverlauf(2, i) = 0 Then K = B_Rex_Green '&H80FF80 '10 'grün
                    If Fehlerverlauf(2, i) >= 100 Then K = vbRed '&H8080FF '12 'rot
                Dumy = (Fehlerverlauf(1, i) + Fehlerverlauf(1, i + 1)) / 2
                If Dumy > Ymax Then
                    Dumy = Ymax
                    'Stop
                End If
                Destination.Line (L, (Fehlerverlauf(1, i) + Fehlerverlauf(1, i - 1)) / 2)-(L + Q, Dumy), K, BF 'dehn
                i = i + 1
            Loop
            'leiste auffüllen
            Dumy = (Fehlerverlauf(1, i - 1) + Fehlerverlauf(1, i)) / 2
            If Dumy > Ymax Then
                Dumy = Ymax
                'Stop
            End If
            Destination.Line (L, Dumy)-(L + Q, Ymax), K, BF         'dehn
            Destination.Line (L, 0)-(L + Q, Ymax), vbBlack, B 'rahmen

        'fehlerwerte eintragen schwingungen
            If Schwingungen_berechnen = True Then
                V = L * 1.033 '+ Destination.TextWidth("0000")
                If B_Rex.Konstruktion.Width > 10000 Then V = L * 1.015 'bei zu grossen bildschirmen verschwindet das sonst rechts aus dem bild, 201909
                Destination.Line (V, 0)-(V + Q, (Fehlerverlauf(1, 0) + Fehlerverlauf(1, 1)) / 2), vbRed, BF 'frequenz
                
                i = 1
                Do Until i >= Rechengenauigkeit - 2 Or Fehlerverlauf(1, i) >= Ymax
                    K = B_Rex_Yellow '&H80FFFF '14 'gelb
                    If Fehlerverlauf(4, i) = 0 Then K = B_Rex_Green '&H80FF80 '10 'grün
                    If Fehlerverlauf(4, i) >= 100 Then K = vbRed '&H8080FF '12 'rot
                    Dumy = (Fehlerverlauf(1, i) + Fehlerverlauf(1, i + 1)) / 2
                    If Dumy > Ymax Then
                        Dumy = Ymax
                        'Stop
                    End If
                    Destination.Line (V, (Fehlerverlauf(1, i) + Fehlerverlauf(1, i - 1)) / 2)-(V + Q, Dumy), K, BF 'dehn
                    i = i + 1
                Loop
                'leiste auffüllen
                Dumy = (Fehlerverlauf(1, i - 1) + Fehlerverlauf(1, i)) / 2
                If Dumy > Ymax Then
                    Dumy = Ymax
                    'Stop
                End If
                Destination.Line (L, Dumy)-(L + Q, Ymax), K, BF         'dehn
                Destination.Line (V, 0)-(V + Q, Ymax), vbBlack, B 'rahmen
            End If
        
        Destination.DrawWidth = Strichbreite
        
        Destination.CurrentX = L / 2 - Destination.TextWidth(Lang_Res(643) & Sys(1).E(34) & "mm") / 2  'Bandposition ab primärer Antriebsscheibe [mm], bo=
        Destination.CurrentY = Fumin - 1.7 * P
        Destination.Print Lang_Res(643) & Sys(1).E(34) & "mm"  'Bandposition ab primärer Antriebsscheibe [mm], bo=
        
        Destination.CurrentX = -0.8 * H
        Destination.CurrentY = Ymax + 1.3 * P
        Destination.Print Lang_Res(692)  'Trumkr.[N]
        
        Destination.CurrentX = L + j - 1.2 * Destination.TextWidth(Lang_Res(0))
        Destination.CurrentY = Ymax + 1.3 * P
        Destination.Print Lang_Res(693) & "   f"  'Dehn.[%]
        
    'trumkraftkurve zeichnen
        Zeichnen = True
        Call Auflegedehnung_ermitteln(Drucken) 'dieses eine letzte mal mit zeichnen
        Zeichnen = False
    
    Destination.DrawStyle = 0
    Destination.DrawWidth = 2
    
    'maximale trumkaft/dehnung oder max. zul aufldehn/kraft eintragen einzeichnen
    If MaxTrumKraft < Fumax And MaxTrumKraft > 0 Then 'paßt sonst nicht aufs diagramm
        Destination.Line (0, MaxTrumKraft)-(L, MaxTrumKraft), QBColor(12)
        Destination.CurrentX = L / 100
        Destination.CurrentY = MaxTrumKraft - 0.1 * P
        Destination.ForeColor = QBColor(12)
        Destination.Print Dehnung$ & Format(MaxTrumKraft * 2 / (SystemTyp.Kraftdehnung * Sys(1).E(34)), "#####0.00")
        Destination.ForeColor = QBColor(0)
    End If
        
    'auflegedehnung einzeichnen
        'erst die gerade
        Destination.Line (0, AuflTrumKraft)-(L + H, AuflTrumKraft), QBColor(1)
        Destination.CurrentY = AuflTrumKraft - 0.1 * P
        Destination.ForeColor = QBColor(1)
        
        Select Case Auflegemodus
            Case 4 'feder/gewicht, hat geklappt
                Destination.CurrentX = L / 50
                Destination.Print Lang_Res(649) & Format(Sys(1).E(53), "#####0.00") & " (" & Sys(1).S(1) & ")"  'Auflegedehnung durch Feder/Gewicht
            Case 3 'dehnungsvorgabe
                Destination.CurrentX = L / 50
                Destination.Print Lang_Res(650) & Format(Sys(1).E(53), "#####0.00") & " (" & Sys(1).S(1) & ")"  'gewählte Auflegedehnung
            Case 1, 2 'comp-optimiert oder in abh mx dehnung
                Destination.CurrentX = L / 50
                Destination.Print Lang_Res(651) & Format(Sys(1).E(53), "#####0.00") & " (" & Sys(1).S(1) & ")"  '"erf. Auflegedehnung
        End Select
        
    Destination.ForeColor = QBColor(0)
End Sub
Private Sub ALD(ByRef X1 As Double, ByRef X2 As Double, ByRef Y1 As Double, ByRef Y2 As Double, ByRef Fuvolumen As Double, ByRef Memo As Double, ByRef YS1 As Double, ByRef YS2 As Double, ByRef YS1memo As Double, ByRef YS2memo As Double, ByRef FuVolumenSp As Double) 'überdeckte Fu-fläche ermitteln, so, als wäre alles im positiven bereich
        'zum Zeichnen
            
        
        If Zeichnen = True Then 'nur beim letzten durchlauf vom modul grafik aus
            Destination.DrawStyle = 0
            Destination.DrawWidth = 2
            Destination.Line (X1, Y1)-(X2, Y2)
        End If
        'und fürs Protokoll zur Ermittlung der Auflegedehnung
        Merk = Y1 'merk ist die kleinere von beiden
        If Y2 < Merk Then Merk = Y2
        Merk = Merk + Abs(Fumin) 'in den positiven bereich
        Fuvolumen = Fuvolumen + Merk * Abs((X1 - X2)) 'der viereckige bereich
        Memo = Y1 'memo ist die größere von beiden
        If Y2 > Memo Then Memo = Y2
        Memo = Memo + Abs(Fumin) 'in den positiven bereich
        Fuvolumen = Fuvolumen + ((Memo - Merk) * Abs(X1 - X2)) / 2 'der dreiecksbereich oben drauf
    
    'spitzenlast
        'zum zeichnen
        If Zeichnen = True Then  'nur beim letzten durchlauf vom modul grafik aus
            Destination.DrawStyle = 1
            Destination.DrawWidth = 1
            'hier wird die spitzenlastlinie noch verschoben, so dass sie genau um die auflegetrumkraft angeordnet wird
            
            
            If Auflegemodus = 4 Then
                'erst wird die spitzenlastkurve verschoben, so dass sie dasselbe mittel wie die normalkurve besitzt
                'dann wird die kurve bei federbelastung genau auf die normale bei der federbelasteten scheibe geschoben
                YS1memo = YS1 + ScheibeFedGewNormalFu - ScheibeFedGewSpitzeFu
                YS2memo = YS2 + ScheibeFedGewNormalFu - ScheibeFedGewSpitzeFu
            Else 'FwScheibeFedGew
                YS1memo = YS1 - AuflTK_Sp_N_Diff
                YS2memo = YS2 - AuflTK_Sp_N_Diff
            End If
            Destination.Line (X1, YS1memo)-(X2, YS2memo)
        End If
        
        
        
        'unbedingt die ys1 und ys2 beurteilen, die memos sind verfälscht von summanten, die erst zum ergebnis gehören
            Merk = YS1 'merk ist die kleinere von beiden
            If Merk < FuminSp Then FuminSp = Merk
            If YS2 < Merk Then Merk = YS2
            Merk = Merk + Abs(Fumin) 'in den positiven bereich, auch bei spitzenlast gültig, weil die kurve erst zum schluss abgesenkt wird
            FuVolumenSp = FuVolumenSp + Merk * Abs((X1 - X2)) 'der viereckige bereich
            Memo = YS1 'memo ist die größere von beiden
            If Memo < FuminSp Then FuminSp = Memo
            If YS2 > Memo Then Memo = YS2
            If Memo > FumaxSp Then FumaxSp = Memo 'nur für die skalierung
            Memo = Memo + Abs(Fumin) 'in den positiven bereich
            FuVolumenSp = FuVolumenSp + ((Memo - Merk) * Abs(X1 - X2)) / 2 'der dreiecksbereich oben drauf

End Sub

Public Sub Auflegedehnung_ermitteln(ByVal Drucken As Boolean)
'auflegedehnung ermitteln und ev. kurve zeichnen, wenn es aus grafik aufgerufen wird
Dim Memo As Double, Merk As Double, Furein As Double, Lrein As Double, L As Double, Fuvolumen As Double, FuVolumenSp As Double
Dim Rechts As Double, Links As Double
Dim X1 As Double, X2 As Double, Y1 As Double, Y2 As Double
Dim YS1 As Double, YS2 As Double 'spitzenlasten, die x- koordinaten sind identisch
Dim YS1memo As Double, YS2memo As Double 'zwischenspeichern

'Dim ImFörderer As Boolean
Dim K As Integer, IaltK As Integer, j As Integer, i As Integer, LetztHP As Integer
'L soll die tatsächliche rechenreihenfolge (start im Leertrum) tarnen
    
    FuVolumenSp = 0
    Fuvolumen = 0
    
    
    FumaxSp = 0
    FuminSp = 1000000
    
    
    L = Sys(Antriebsscheibe).Lraus
    K = Startelement 'wird oben in richtiger richtung festgelegt
    If Sys(K).Element = "" Then Exit Sub 'keine anlage da, hier droht absturz
    'If K = 0 Then Exit Sub
    IaltK = Antriebsscheibe 'von da kommt er, da soll er nicht gleich wieder hin

    X1 = L
    Y1 = Sys(K).Furein
    X2 = L - Sys(K).Lrein
    Y2 = Sys(K).Furein
    'If Drucken = True Then Stop
    YS1 = Y1
    YS2 = Y2 + (X1 - X2) * Sys(1).FusteigSp  'nur die bandbeschleunigung
    'GoSub ALD
    Call ALD(X1, X2, Y1, Y2, Fuvolumen, Memo, YS1, YS2, YS1memo, YS2memo, FuVolumenSp)  'ist schneller gefunden, wenn es in derselben prozedur ist
    
    
    'die kurve zum aktuellen element muß hier schon gezeichnet sein
    Do
        'festhalten zur berechnung der durchbiegung bei scheiben, nullen, falls es anderen elementen zugeordnet wird
            Sys(K).FureinSp = 0
            Sys(K).FurausSp = 0
        
        If left(Sys(K).Tag, 1) = "0" Then 'einfach stehendes
            'scheiben
            X1 = L - Sys(K).Lrein
            Y1 = Sys(K).Furein
            X2 = L - Sys(K).Lraus
            Y2 = Sys(K).Furaus
            
            'spitzenlast
                YS1 = YS2
                'anfang der strecke +kraft aus normalbetrieb +spitzenlast durch beschleunigung dieses bandstückchens +spitzenlast durch beschleunigung der scheibenmasse
                YS2 = YS2 + (Y2 - Y1) + (X1 - X2) * Sys(1).FusteigSp + Sys(K).E(98)  'und dann noch irgendwas
                
                'festhalten zur berechnung der durchbiegung bei scheiben
                'eigentlich nur für den letzten durchlauf erforderlich
                'If Zeichnen = True Then'immer, weil sie trotz abgeschaltetem zeichnen erfasst werden muss
                    Sys(K).FureinSp = YS1 '- (AuflTrumkraftSp - AuflTrumKraft)
                    Sys(K).FurausSp = YS2 '- (AuflTrumkraftSp - AuflTrumKraft)
                'End If
                
                'abweichung mitprotokollieren zur korrektur der spitzenlast bei feder/gewicht
                'nur beim letzten mal nicht, sonst gehts durcheinander
                'hier, weil k hier eindeutig zur scheibe und nicht zur strecke zum nächsten element gehört
                'wiederholung ganz unten bei schlupfausgleich an der antriebsscheibe
                If K = ScheibeFedGew Then
                    ScheibeFedGewSpitzeFu = (YS1 + YS2) / 2 '- (AuflTrumkraftSp - AuflTrumKraft) 'letzte klammer ist korrektur, wos hingezeichnet wird
                    ScheibeFedGewNormalFu = (Y1 + Y2) / 2
                End If
            
            Call ALD(X1, X2, Y1, Y2, Fuvolumen, Memo, YS1, YS2, YS1memo, YS2memo, FuVolumenSp) 'kurve im element, wenn's kein träger ist
            'ImFörderer = False
        Else 'förderer
            Lrein = 0 'enthält Abstand vom Trägeranfang
            Furein = Sys(K).Furein 'protokolliert fuverlauf der huckepacks
            'M = 0 'akt. Position auf Träger
            If Sys(K).Rechts = True Then 'wie unter B_Rex1: reversieren = false
                'also zuerst die linken teile (e(25) als maßstab in allen fällen, denn den haben alle huckepacks
                Rechts = Sys(K).E(22) + 1 'nur einmal pro förderer und dann von rechts nach links
                Links = -1
                Do
                    j = 0
                    For i = 9 To Maxelementindex
                        If Sys(i).Zugehoerigkeit = K And Sys(i).Tag <> "201" Then 'kein transportgut
                            If Sys(i).E(25) > Links And Sys(i).E(25) < Rechts Then
                                j = i
                                Rechts = Sys(i).E(25) 'neue linke grenze
                            End If
                        End If
                    Next i
                    Links = Sys(j).E(25) 'damit die linken nicht immer wieder drankommen
                    Rechts = Sys(K).E(22) + 1
                    If j > 0 Then 'er hat noch ein unbehandeltes gefunden
                        If Sys(j).Tag = "204" Or Sys(j).Tag = "206" Then 'stau oder trägergebundene_Umfangskraft
                            
                            'zum huckepack
                                X1 = L - (Sys(K).Lrein + Lrein) 'zum huckepack
                                Y1 = Furein
                                X2 = L - (Sys(K).Lrein + Sys(j).E(25))
                                Y2 = (X1 - X2) * Sys(K).Fusteig + Furein
                                
                                'spitzenlast
                                    YS1 = YS2
                                    'anfang der strecke
                                    'kraftanstieg aus normalbetrieb
                                    'spitzenlast durch beschleunigung dieses bandstückchens
                                    'spitzenlast durch beschleunigung des transportgutes
                                    'spitzenlast durch beschleunigung der rollen, sofern welche da sind
                                    YS2 = YS2 + (Y2 - Y1) + (X1 - X2) * (Sys(1).FusteigSp + Sys(K).FusteigSp + Sys(K).FusteigSpRoll) 'und dann noch irgendwas
                                
                                Call ALD(X1, X2, Y1, Y2, Fuvolumen, Memo, YS1, YS2, YS1memo, YS2memo, FuVolumenSp)
                            
                            'durchs huckepack
                                X1 = X2
                                Y1 = Y2
                                X2 = L - (Sys(K).Lrein + Sys(j).E(46))
                                Y2 = Y1 + Sys(j).E(50)
                                Lrein = Sys(j).E(46)
                                
                                'spitzenlast
                                    YS1 = YS2
                                    'anfang der strecke
                                    'kraftanstieg aus normalbetrieb
                                    'spitzenlast durch beschleunigung dieses bandstückchens
                                    'spitzenlast durch beschleunigung des transportgutes
                                    'spitzenlast durch beschleunigung der rollen, sofern welche da sind
                                    'If Zeichnen = True Then Stop
                                    If Sys(j).Tag = "204" Then 'stau
                                        YS2 = YS2 + (Y2 - Y1) + (X1 - X2) * (Sys(1).FusteigSp + Sys(K).FusteigSpRoll) 'ohne beschleunigung transportgut
                                    End If
                                    If Sys(j).Tag = "206" Then 'freie förderergebundene umfangskraft
                                        If (Sys(j).E(46) - Sys(j).E(25)) > 0 Then
                                            Sys(j).FusteigSp = Sys(j).E(98) / (Sys(j).E(46) - Sys(j).E(25)) 'anstieg über strecke
                                        Else
                                            YS2 = YS2 + Sys(j).E(98) 'dann eben punktueller anstieg
                                        End If
                                        YS2 = YS2 + (Y2 - Y1) + (X1 - X2) * (Sys(1).FusteigSp + Sys(K).FusteigSp + Sys(K).FusteigSpRoll + Sys(j).FusteigSp) 'und dann noch irgendwas
                                    End If
                                    
                        End If
                        If Sys(j).Tag = "205" Then 'abweiser
                            
                            'zum abweiser
                                X1 = L - (Sys(K).Lrein + Lrein) 'k enthält den träger
                                Y1 = Furein
                                X2 = L - (Sys(K).Lrein + Sys(j).E(25))
                                Y2 = (X1 - X2) * Sys(K).Fusteig + Furein
                                
                                'spitzenlast
                                    YS1 = YS2
                                    'anfang der strecke
                                    'kraftanstieg aus normalbetrieb
                                    'spitzenlast durch beschleunigung dieses bandstückchens
                                    'spitzenlast durch beschleunigung des transportgutes
                                    'spitzenlast durch beschleunigung der rollen, sofern welche da sind
                                    YS2 = YS2 + (Y2 - Y1) + (X1 - X2) * (Sys(1).FusteigSp + Sys(K).FusteigSp + Sys(K).FusteigSpRoll) 'und dann noch irgendwas
                                
                                Call ALD(X1, X2, Y1, Y2, Fuvolumen, Memo, YS1, YS2, YS1memo, YS2memo, FuVolumenSp)
                            
                            'durch den abweiser (ist nur n Punkt, keine Strecke)
                                X1 = X2
                                Y1 = Y2
                                Y2 = Y1 + Sys(j).E(50)
                                Lrein = Sys(j).E(25)
                                'spitzenlast
                                    YS1 = YS2
                                    'anfang der strecke
                                    'kraftanstieg aus normalbetrieb
                                    YS2 = YS2 + (Y2 - Y1)
                        End If
                        Call ALD(X1, X2, Y1, Y2, Fuvolumen, Memo, YS1, YS2, YS1memo, YS2memo, FuVolumenSp) 'im stau/Abweiser
                        Sys(j).Lrein = X1
                        Sys(j).Lraus = X2
                        Sys(j).Furein = Y1
                        Sys(j).Furaus = Y2
                        'aktualisieren
                        Furein = Y2
                    End If
                Loop Until j = 0 'kein stau/abweiser/freie umfangskraft mehr gefunden
            Else 'wie unter B_rex1, reversieren = true
                Rechts = Sys(K).E(22) + 1 'nur einmal pro förderer und dann von rechts nach links
                Links = -1
                Do
                    j = 0
                    For i = 9 To Maxelementindex
                        If Sys(i).Zugehoerigkeit = K And Sys(i).Tag <> "201" Then 'kein transportgut
                            If Sys(i).E(25) > Links And Sys(i).E(25) < Rechts Then
                                j = i
                                Links = Sys(i).E(25) 'neue rechte grenze
                            End If
                        End If
                    Next i
                    Links = -1
                    Rechts = Sys(j).E(25)
                    
                    If j > 0 Then 'er hat noch ein unbehandeltes gefunden
                        If Sys(j).Tag = "204" Or Sys(j).Tag = "206" Then 'stau oder trägergebundene_Umfangskraft
                            'zum huckpack
                                X1 = L - (Sys(K).Lrein + Lrein) 'lrein immer innerhalb des trägers
                                Y1 = Furein
                                X2 = L - (Sys(K).Lrein + (Sys(K).E(22) - Sys(j).E(46)))
                                Y2 = (X1 - X2) * Sys(K).Fusteig + Furein
                                    
                                'spitzenlast
                                    YS1 = YS2
                                    'anfang der strecke
                                    'kraftanstieg aus normalbetrieb
                                    'spitzenlast durch beschleunigung dieses bandstückchens
                                    'spitzenlast durch beschleunigung des transportgutes
                                    'spitzenlast durch beschleunigung der rollen, sofern welche da sind
                                    YS2 = YS2 + (Y2 - Y1) + (X1 - X2) * (Sys(1).FusteigSp + Sys(K).FusteigSp + Sys(K).FusteigSpRoll) 'und dann noch irgendwas

                                Call ALD(X1, X2, Y1, Y2, Fuvolumen, Memo, YS1, YS2, YS1memo, YS2memo, FuVolumenSp) 'zum stau
                            
                            'durchs huckepack
                                X1 = X2
                                Y1 = Y2
                                X2 = L - (Sys(K).Lrein + (Sys(K).E(22) - Sys(j).E(25)))
                                Y2 = Y2 + Sys(j).E(50)
                                    
                                'spitzenlast
                                    YS1 = YS2
                                    'anfang der strecke
                                    'kraftanstieg aus normalbetrieb
                                    'spitzenlast durch beschleunigung dieses bandstückchens
                                    'spitzenlast durch beschleunigung des transportgutes
                                    'spitzenlast durch beschleunigung der rollen, sofern welche da sind
                                    'If Zeichnen = True Then Stop
                                    If Sys(j).Tag = "204" Then 'stau
                                        YS2 = YS2 + (Y2 - Y1) + (X1 - X2) * (Sys(1).FusteigSp + Sys(K).FusteigSpRoll) 'ohne beschleunigung transportgut
                                    End If
                                    If Sys(j).Tag = "206" Then 'freie förderergebundene umfangskraft
                                        If (Sys(j).E(46) - Sys(j).E(25)) > 0 Then
                                            Sys(j).FusteigSp = Sys(j).E(98) / (Sys(j).E(46) - Sys(j).E(25)) 'anstieg über strecke
                                        Else
                                            YS2 = YS2 + Sys(j).E(98) 'dann eben punktueller anstieg
                                        End If
                                        YS2 = YS2 + (Y2 - Y1) + (X1 - X2) * (Sys(1).FusteigSp + Sys(K).FusteigSp + Sys(K).FusteigSpRoll + Sys(j).FusteigSp) 'und dann noch irgendwas
                                    End If

                        End If
                        If Sys(j).Tag = "205" Then
                            'zum abweiser
                                X1 = L - (Sys(K).Lrein + Lrein)
                                Y1 = Furein
                                X2 = L - (Sys(K).Lrein + (Sys(K).E(22) - Sys(j).E(25)))
                                Y2 = (X1 - X2) * Sys(K).Fusteig + Furein
                                
                                'spitzenlast
                                    YS1 = YS2
                                    'anfang der strecke
                                    'kraftanstieg aus normalbetrieb
                                    'spitzenlast durch beschleunigung dieses bandstückchens
                                    'spitzenlast durch beschleunigung des transportgutes
                                    'spitzenlast durch beschleunigung der rollen, sofern welche da sind
                                    YS2 = YS2 + (Y2 - Y1) + (X1 - X2) * (Sys(1).FusteigSp + Sys(K).FusteigSp + Sys(K).FusteigSpRoll) 'und dann noch irgendwas
                                
                                Call ALD(X1, X2, Y1, Y2, Fuvolumen, Memo, YS1, YS2, YS1memo, YS2memo, FuVolumenSp)
                            
                            'durch den abweiser (ist nur n Punkt, keine Strecke)
                                X1 = X2
                                Y1 = Y2
                                Y2 = Y1 + Sys(j).E(50)
                                'spitzenlast
                                    YS1 = YS2
                                    'anfang der strecke
                                    'kraftanstieg aus normalbetrieb
                                    YS2 = YS2 + (Y2 - Y1)

                        End If
                        Call ALD(X1, X2, Y1, Y2, Fuvolumen, Memo, YS1, YS2, YS1memo, YS2memo, FuVolumenSp) 'im stau/abweiser
                        Sys(j).Lrein = X1
                        Sys(j).Lraus = X2
                        Sys(j).Furein = Y1
                        Sys(j).Furaus = Y2
                        Furein = Y2
                        Lrein = Sys(K).E(22) - Sys(j).E(25)
                   End If
                Loop Until j = 0
            End If
           
            'restlänge bis trägerende
                X1 = L - (Sys(K).Lrein + Lrein)
                Y1 = Furein
                X2 = L - Sys(K).Lraus
                Y2 = Sys(K).Furaus
                
                'spitzenlast
                    YS1 = YS2
                    'anfang der strecke
                    'kraftanstieg aus normalbetrieb
                    'spitzenlast durch beschleunigung dieses bandstückchens
                    'spitzenlast durch beschleunigung des transportgutes
                    'spitzenlast durch beschleunigung der rollen, sofern welche da sind
                    YS2 = YS2 + (Y2 - Y1) + (X1 - X2) * (Sys(1).FusteigSp + Sys(K).FusteigSp + Sys(K).FusteigSpRoll) 'und dann noch irgendwas
    
                Call ALD(X1, X2, Y1, Y2, Fuvolumen, Memo, YS1, YS2, YS1memo, YS2memo, FuVolumenSp)
        End If
        
        'nächstes element entlang des bandes ermitteln
        If Sys(K).Verb(1, 1) = IaltK Then 'voreinstellungen für neuen durchlauf
            IaltK = K
            K = Sys(K).Verb(2, 1)
        Else
            IaltK = K
            K = Sys(K).Verb(1, 1)
        End If
        
        'linie zum nächsten element ziehen
            X1 = L - Sys(IaltK).Lraus
            Y1 = Y2
            X2 = L - Sys(K).Lrein
            Y2 = Sys(K).Furein
            
            'spitzenlast
                YS1 = YS2
                'anfang der strecke + kraft aus normalbetrieb +spitzenlast durch beschleunigung dieses bandstückchens
                YS2 = YS2 + (Y2 - Y1) + (X1 - X2) * Sys(1).FusteigSp  'und dann noch irgendwas
            
            Call ALD(X1, X2, Y1, Y2, Fuvolumen, Memo, YS1, YS2, YS1memo, YS2memo, FuVolumenSp) 'kurve zum nächsten element, wird immer durchgeführt
        
    Loop Until K = Antriebsscheibe Or K >= Maxelementindex + 1 'einmal rum oder es ist was schiefgegangen
    
    
    'Schlupfausgleich an der Antriebsscheibe
        X1 = L - Sys(K).Lrein
        Y1 = Sys(K).Furein
        X2 = 0
        Y2 = Sys(Startelement).Furein
        
        'spitzenlast
            YS1 = YS2
            FuletztesSp = YS1 'hieraus wird später die leistung an der Antriebsscheibe berechnet
            YS2 = Y2 'zurück an den ursprung
        
            'If Zeichnen = True Then 'spitzenlastanteil für durchbiegung festhalten
                Sys(K).FureinSp = YS1 '- (AuflTrumkraftSp - AuflTrumKraft)
                Sys(K).FurausSp = YS2 '- (AuflTrumkraftSp - AuflTrumKraft)
            'End If
        
            'abweichung mitprotokollieren zur korrektur der spitzenlast bei feder/gewicht
                'nur beim letzten mal nicht, sonst gehts durcheinander
                'hier, weil k hier eindeutig zur scheibe und nicht zur strecke zum nächsten element gehört
                If K = ScheibeFedGew Then
                    ScheibeFedGewSpitzeFu = (YS1 + YS2) / 2 '- (AuflTrumkraftSp - AuflTrumKraft) 'letzte klammer ist korrektur, wos hingezeichnet wird
                    ScheibeFedGewNormalFu = (Y1 + Y2) / 2
                End If
        
        Call ALD(X1, X2, Y1, Y2, Fuvolumen, Memo, YS1, YS2, YS1memo, YS2memo, FuVolumenSp) 'schlupfausgleich über den durchmesser der antriebsscheibe
    
    'If Zeichnen = True Then Stop
    'zuletzt das ziel der reise, die auflegedehnung zu dieser kurve ermitteln
    Fuvolumen = Fuvolumen - Abs(Fumin) * L 'fumin wurde addiert, um im positiven bereich zu rechnen
    FuVolumenSp = FuVolumenSp - Abs(Fumin) * L 'fumin wurde addiert, um im positiven bereich zu rechnen
    
    'in der trumkraft ist fliehkraft enthalten, deren summe oben ermittelt wurde.
    'die wieder abziehen, denn auflegedehnung ist bei stehender anlage ohne fliehkraft.
    'aufltrumkraft durchschnitt aller werte ohne fliehkraftanteil
    
    If Zeichnen = False Then
        'würde bei diesem letzten rechengang ohnehin nicht verändert, weil identisch mit der letzten vollen berechnung
        'bei feder/gewicht würde aber die niveaukorrektur wieder zunichte gemacht, also finger wech
        If Auflegemodus = 4 Then
            AuflTrumKraft = ScheibeFedGewNormalFu
            
            'AuflTrumkraftSp = ScheibeFedGewSpitzeFu
            'egal, wo die ist, es gibt keine weitere auswertung
            
        Else
            AuflTrumKraft = Fuvolumen / L ' - Fliehkraftsumme
            AuflTrumkraftSp = FuVolumenSp / L ' - Fliehkraftsumme 'na schön, stimmt nur zum ende der beschleunigung, aber da ists eben maximal
            'differenz der gemittelten dehnung von normal und spitzenlastkurve vom letzten durchgang
            'benoetigt eigentlich erst beim letzten durchlauf, der 2mal stattfindet, weil eben diese differenz fuer den letzten durhlauf eigentlich bekant sein muß
            AuflTK_Sp_N_Diff = (AuflTrumkraftSp - AuflTrumKraft)
        
        End If
    End If
    
    'If Zeichnen = True Then Stop
    Sys(1).E(53) = AuflTrumKraft * 2 / (SystemTyp.Kraftdehnung * Sys(1).E(34)) 'prozent aus kraft
    'die noch absenken, weil sie versetzt erfaßt wurde

    'noch hinbiegen, weil es eben verfälscht erfaßt wurde
    If Auflegemodus = 4 Then
        FumaxSp = FumaxSp + ScheibeFedGewNormalFu - ScheibeFedGewSpitzeFu
        FuminSp = FuminSp + ScheibeFedGewNormalFu - ScheibeFedGewSpitzeFu
    Else
        FumaxSp = FumaxSp - AuflTK_Sp_N_Diff
        FuminSp = FuminSp - AuflTK_Sp_N_Diff
    End If
End Sub

Function ReibungszahlBerechnung(M As Double, Alterung As Integer, Einstellung As Double, Materialpaarung As Boolean) As Double
Dim Ampmiser As Boolean
    'Alterung, neu 201602
        '0 = ohne Alterung
        '1 = Alterung nur bei unter 0,35 Reibungszahl
        '2 = Alterung nur bei ueber 0,35 Reibungszahl
        'beispiel  mue = ReibungszahlBerechnung(mue, 1, Sys(K).E(14), true)
        If InStr(Sys(1).S(1), "TX0") > 0 Then Ampmiser = True
        If Init_B_Rex_Aging = 1 Then
            If Alterung = 1 And M < 0.35 Then M = (Abs(M) + 0.35) / 2
            If Alterung = 2 And M > 0.35 Then M = (Abs(M) + 0.35) / 2
        End If
        
        If Materialpaarung = True And Init_B_Rex_Pairing = 1 Then M = M * Kst(Einstellung).Einstellung 'gleittischunterlage als faktor
        
        '202408 bei stahl und verzinktem stahl Ampmiser einen vorteil verschaffen
        If Ampmiser = True And Init_B_Rex_Aging = 1 And Alterung = 1 And M < 0.35 Then
            Select Case Kst(Einstellung).ID
                Case 5
                    M = 0.18
                Case 124
                    M = 0.24
            End Select
        End If
        
        ReibungszahlBerechnung = M
End Function
Function MassentraegheitsErmittlung(Element As Integer) As Double
    'manuelle Alterung ueberschreibt errechnete, neu 202006
        '8 errechnet
        '114 manuell
        MassentraegheitsErmittlung = Sys(Element).E(8)
        If Sys(Element).E(114) > 0 Then MassentraegheitsErmittlung = Sys(Element).E(114)
End Function
Function Fw_Fu_Winkelabh(Winkel As Double) As Double
    Fw_Fu_Winkelabh = 1.9
    If Winkel > 30 Then Fw_Fu_Winkelabh = 2.1
    If Winkel > 60 Then Fw_Fu_Winkelabh = 2.2
    If Winkel > 90 Then Fw_Fu_Winkelabh = 2.2
    If Winkel > 120 Then Fw_Fu_Winkelabh = 2.2
    If Winkel > 150 Then Fw_Fu_Winkelabh = 2.1
    If Winkel > 180 Then Fw_Fu_Winkelabh = 1.9
    If Winkel > 210 Then Fw_Fu_Winkelabh = 1.7
    If Winkel > 240 Then Fw_Fu_Winkelabh = 1.5
    If Winkel > 270 Then Fw_Fu_Winkelabh = 1.3
    If Winkel > 300 Then Fw_Fu_Winkelabh = 1.1
    If Winkel > 330 Then Fw_Fu_Winkelabh = 1
End Function

