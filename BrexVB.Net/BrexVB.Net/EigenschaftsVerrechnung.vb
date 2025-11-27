Option Explicit On
Module EigenschaftsVerrechnung
    Private AntrScheibe As Boolean
    Private UmlScheibe As Boolean
    Private Achsabstand As Double
    Private Awert As Double 'gilt modulweit, wird übergeben
    Private Nwert As Double 'gilt modulweit, wird übergeben
    Private R1 As Integer
    Private R2 As Integer


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
End Module
