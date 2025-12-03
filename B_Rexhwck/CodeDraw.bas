Attribute VB_Name = "CodeDraw"
Option Explicit

Private LYStrang As Integer
Private LYStrangx As Integer
Private LYStrangy As Integer
Private X2 As Integer
Private Y2 As Integer
Private Ausgabeobjekt As Integer

Public Sub Alleelementeverbinden(Optional ByVal x As Integer, Optional ByVal y As Integer)
    Dim i, M, P, K As Integer
    Dim L, Q As Single
    Dim Aufwärts As Boolean
    
    X2 = x
    Y2 = y
    LYStrang = 0
    LYStrangx = 0
    LYStrangy = 0
    Ausgabeobjekt = 0 'sonst wird cls nicht ausgefuehrt, wenn man einmal in kundenunterlagen war
    If AnlageRefresh = True Then ModusCalc = ""
    
    Destination.FontSize = 7
    If Druck = True Then
        Set Destination = Printer
    Else
        If B_Rex.Visible = True Then
            Set Destination = B_Rex.Konstruktion
        End If
    End If
    Destination.FontTransparent = True 'sonst bleibt von der schrift nur ein schwarzer balken (betriebssystem setzt back=frontfarbe)
    Destination.FillStyle = 1
    
    'löschen, clear, clr
    If ModusCalc <> "liniensuchen" Then
        If Druck = False And Ausgabeobjekt <> 1 Then B_Rex.Konstruktion.Cls
    End If
    
    'bei zweischeibentriebmode nicht die anlage zeichnen
    If B_Rex.Button(0).Tag = "E" Then
        Call Zweischeibentrieb_zeichnen
        Exit Sub
    End If
    
    AnlageRefresh = False
    
    'sieglinglogo nur für kunden zuunterst'10% der breite abstand von allem anderen
    If Druck = False And UserStatus <> 2 Then
        'forbomovement
    End If
    
    'werden neu gewählt und optimiert, daher zunächst gelöscht, damit auch 2 elemente 2mal verbunden werden können
    For i = 9 To Maxelementindex
        Sys(i).Verb(1, 2) = 0
        Sys(i).Verb(2, 2) = 0
    Next i
        
    'so weit so gut, aber vielleicht gehts noch besser, zumindest bei scheiben
    For i = 9 To Maxelementindex
        If Sys(i).Verb(1, 1) <> Sys(i).Verb(2, 1) Then 'nur wenns einzelstehende elemente sind
            If left(Sys(i).Tag, 1) = "0" Then Call Pfadoptimieren(i) 'nur bei den scheiben
        End If
    Next i

    If Endlos = False Then
        For i = 9 To Maxelementindex
            If Sys(i).Element <> "" Then
                'immer nur die verbindung zum höheren element aufbauen, sonst wirds ja doppelt gezeichnet
                If Sys(i).Verb(1, 1) > i Then
                    Call Elementeverbinden(i, Sys(i).Verb(1, 1))
                End If
                If Sys(i).Verb(2, 1) > i Then
                    Call Elementeverbinden(i, Sys(i).Verb(2, 1))
                End If
            End If
        Next i
    Else 'geschlossene anlage an den elementen entlang abhandeln
        i = 9 'antriebsscheibe suchen
        Do
            i = i + 1
        Loop Until Sys(i).Tag = "001" Or i = Maxelementindex
        
        'minimalste entfernung finden, dort wird angefangen
        K = 10000 'enthält den Maximalwert, der abgesenkt wird
        L = Sqr(Abs(Sys(Sys(i).Verb(1, 1)).left - Sys(i).left) ^ 2 + Abs(Sys(Sys(i).Verb(1, 1)).Top - Sys(i).Top) ^ 2)
        If L < K Then
            K = L
            P = 1
        End If
        L = Sqr(Abs(Sys(Sys(i).Verb(1, 1)).left + Sys(Sys(i).Verb(1, 1)).Width - Sys(i).left) ^ 2 + Abs(Sys(Sys(i).Verb(1, 1)).Top - Sys(i).Top) ^ 2)
        If L < K Then
            K = L
            P = 1
        End If
        L = Sqr(Abs(Sys(Sys(i).Verb(2, 1)).left - Sys(i).left) ^ 2 + Abs(Sys(Sys(i).Verb(2, 1)).Top - Sys(i).Top) ^ 2)
        If L < K Then
            K = L
            P = 2
        End If
        L = Sqr(Abs(Sys(Sys(i).Verb(2, 1)).left + Sys(Sys(i).Verb(2, 1)).Width - Sys(i).left) ^ 2 + Abs(Sys(Sys(i).Verb(2, 1)).Top - Sys(i).Top) ^ 2)
        If L < K Then
            K = L
            P = 2
        End If
        Do
            K = i ' altes element merken, los gehts mit antriebsscheibe
            i = Sys(K).Verb(P, 1) 'und nächstes bestimmen
            If Sys(i).Verb(1, 1) = K Then 'voreinstellungen für neuen durchlauf
                P = 2
            Else
                P = 1
            End If
            'Destination.FontTransparent = True
            Call Elementeverbinden(K, i)
        Loop Until Sys(i).Tag = "001" Or Sys(i).Tag = ""  'ist einmal rum'oder es ist ein fehler im kreislauf
    End If
    
    'elemente zeichnen
    'horizontal/vertikalbalken anlegen
    If ModusCalc <> "liniensuchen" Then
        Anlbreite = 0
        Anlhöhe = 0
        'Destination.FontBold = True
        
        For i = 9 To Maxelementindex
            If Sys(i).Element <> "" Then
                'huckepacks werden automatisch in 'element_aufbauen' mit aufgebaut, beschriftet werden sie allerdings hier
                If left(Sys(i).Tag, 1) <> "2" Then Call Element_aufbauen(i, Sys(i).left, Sys(i).Top)
                If Sys(i).left + Sys(i).Width > Anlbreite Then Anlbreite = Sys(i).left + Sys(i).Width
                If Sys(i).Top + Sys(i).Height > Anlhöhe Then Anlhöhe = Sys(i).Top + Sys(i).Height
            End If
        Next i
         
        'beschriftung
        For i = 9 To Maxelementindex
            If Sys(i).Element <> "" Then
                Destination.CurrentX = Sys(i).left + Sys(i).Width / 2
                If left(Sys(i).Tag, 1) = "1" Then Destination.CurrentX = Sys(i).left + Sys(i).Width
                Destination.CurrentY = Sys(i).Top - Destination.TextHeight("8") * 1.2
                Destination.Print "(" & i & ")"
            End If
        Next i
        
        'Destination.FontBold = False
        'mindestens jedoch die breite des fensters
        If Druck = False Then
            If Anlbreite < B_Rex.Konstruktion.Width / Screen.TwipsPerPixelX Then Anlbreite = B_Rex.Konstruktion.Width / Screen.TwipsPerPixelX
            If Anlhöhe < B_Rex.Konstruktion.Height / Screen.TwipsPerPixelX Then Anlhöhe = B_Rex.Konstruktion.Height / Screen.TwipsPerPixelY
            'plus zuschlag für vergrösserungen
            Anlhöhe = Anlhöhe ' + 200
            Anlbreite = Anlbreite ' + 200
            B_Rex.Vertikal.Max = Anlhöhe
            B_Rex.Vertikal.LargeChange = 400
            B_Rex.Horizontal.Max = Anlbreite
            B_Rex.Horizontal.LargeChange = 400
        End If
    End If
    
    'Call Eigschaftsverr.Zwei_Scheiben 'damit Vollstaendige anlagen auch so angezeigt werden (nicht falsch für extremultus gehalten)
    
    For M = 1 To Maxelementindex
        'förderrichtung einzeichnen
        Aufwärts = False 'voreinstellung

        If left(Sys(M).Tag, 1) = "1" And Endlos = True And Sys(M).Verb(1, 1) <> 0 Then  'richtungspfeile für Träger einzeichnen
            'schwarze linie über träger
            Destination.Line (Sys(M).left, Sys(M).Top - 32)-(Sys(M).left + Sys(M).Width, Sys(M).Top - 34), QBColor(0), BF
            Destination.DrawWidth = 2
            
            If left(Sys(M).Tag, 3) <> "103" Then 'tisch, tragrollenbahn, freie_Umfangskraft
                If Sys(M).Rechts = False Then
                    Destination.Line (Sys(M).left, Sys(M).Top - 34)-(Sys(M).left + 8, Sys(M).Top - 38), QBColor(0)
                    If Sys(M).E(16) > 0 Then Aufwärts = True 'abwärts
                Else
                    Destination.Line (Sys(M).left + Sys(M).Width, Sys(M).Top - 34)-(Sys(M).left + Sys(M).Width - 8, Sys(M).Top - 38), QBColor(0)
                    If Sys(M).E(16) < 0 Then Aufwärts = True
                End If
            Else 'rollenbahn fördert gegen bandrichtung
                If Sys(M).Rechts = True Then
                    Destination.Line (Sys(M).left, Sys(M).Top - 34)-(Sys(M).left + 8, Sys(M).Top - 38), QBColor(0)
                    If Sys(M).E(16) > 0 Then Aufwärts = True
                Else
                    Destination.Line (Sys(M).left + Sys(M).Width, Sys(M).Top - 34)-(Sys(M).left + Sys(M).Width - 8, Sys(M).Top - 38), QBColor(0)
                    If Sys(M).E(16) < 0 Then Aufwärts = True
                End If
            End If
            
            'aufwärts/abwärts-Pfeil eintragen
            If Sys(M).E(16) <> 0 Then 'abwärts
                Destination.DrawWidth = 3
                Destination.Line (Sys(M).Width / 2 + Sys(M).left, -14 + Sys(M).Top)-(Sys(M).Width / 2 + Sys(M).left, 22 + Sys(M).Top), QBColor(0)
                If Aufwärts = True Then
                    Destination.Line (Sys(M).Width / 2 + Sys(M).left, 22 + Sys(M).Top)-(Sys(M).Width / 2 + Sys(M).left - 4, 15 + Sys(M).Top), QBColor(0)
                    Destination.Line (Sys(M).Width / 2 + Sys(M).left, 22 + Sys(M).Top)-(Sys(M).Width / 2 + Sys(M).left + 4, 15 + Sys(M).Top), QBColor(0)
                Else
                    Destination.Line (Sys(M).Width / 2 + Sys(M).left, -14 + Sys(M).Top)-(Sys(M).Width / 2 + Sys(M).left - 4, -7 + Sys(M).Top), QBColor(0)
                    Destination.Line (Sys(M).Width / 2 + Sys(M).left, -14 + Sys(M).Top)-(Sys(M).Width / 2 + Sys(M).left + 4, -7 + Sys(M).Top), QBColor(0)
                End If
            End If
             
            'staus eintragen als rote linien innerhalb der schwarzen trägerlinie
            Destination.DrawWidth = 1
            For i = 9 To Maxelementindex
                If M <> i And Sys(i).Zugehoerigkeit = M And Sys(i).Tag <> "201" And left(Sys(i).Tag, 1) <> "201" And Sys(M).E(22) > 0 Then
                    L = Sys(i).E(25) / Sys(M).E(22) * Sys(M).Width + Sys(M).left
                    Q = Sys(i).E(46) / Sys(M).E(22) * Sys(M).Width + Sys(M).left
                    Destination.Line (L, Sys(M).Top - 32)-(Q, Sys(M).Top - 34), QBColor(12), BF
                End If
            Next i
            
        End If 'bis hierher mußte es ein träger sein
        
        'fördererlängen einsetzen
        If left(Sys(M).Tag, 1) = "1" And Sys(M).E(22) > 0 Then
            Destination.CurrentX = (Sys(M).Width - Destination.TextWidth(Sys(M).E(22))) / 2 + Sys(M).left
            Destination.CurrentY = Sys(M).Top + 17
            Destination.Print Sys(M).E(22)
            i = Printer.DrawMode
        End If
      
        'fragezeichen setzen an unVollstaendigen elementen
        If Sys(M).Element <> "" And Sys(M).Vollstaendig = False Then
            B_Rex.Elbilder.ClipHeight = 15
            B_Rex.Elbilder.ClipWidth = 15
            B_Rex.Elbilder.ClipX = 355
            B_Rex.Elbilder.ClipY = 105
            P = Sys(M).left
            If Sys(M).Tag = "001" Then P = P + 17
            i = Sys(M).Top
            If left(Sys(M).Tag, 1) = "2" Then i = i + 10
            If M > 9 Then Destination.PaintPicture B_Rex.Elbilder.Clip, P, i, 15, 15 'an den elementen
            If M = 1 And LYStrangx <> 0 Then Destination.PaintPicture B_Rex.Elbilder.Clip, LYStrangx, LYStrangy, 15, 15 'am band
        End If
    Next M
    

End Sub
Public Sub Zweischeibentrieb_zeichnen()
Dim Pic$
Dim Breite As Double, Hoehe As Double, x As Single, Br As Single, i As Integer
Dim Pos10 As Single, Pos11 As Single
If Druck = False And Ausgabeobjekt = 0 Then Destination.Cls 'dann isses die konstruktion


If Sys(10).E(2) > Sys(11).E(2) Then
    'antrieb>angetrieben
    'enthält nur die position fuer die beschriftung
    Pic$ = App.Path & "\drawing\zweischeiben_linksgetrieben.jpg"
    Pos10 = (B_Rex.Konstruktion.Width / Screen.TwipsPerPixelX) / 6 * 1.5 'antriebsscheibe nach links
    Pos11 = (B_Rex.Konstruktion.Width / Screen.TwipsPerPixelX) / 6 * 4 'angetriebene nach rechts
Else
    Pic$ = App.Path & "\drawing\zweischeiben_rechtsgetrieben.jpg"
    Pos11 = (B_Rex.Konstruktion.Width / Screen.TwipsPerPixelX) / 6 * 1.5 'antriebsscheibe nach links
    Pos10 = (B_Rex.Konstruktion.Width / Screen.TwipsPerPixelX) / 6 * 4 'angetriebene nach rechts
End If
    
If FileExist(Pic$) = False Then Exit Sub
    
 Dim MyPic As StdPicture
 Set MyPic = LoadPicture(Pic$)
 
 Breite = CLng(B_Rex.Konstruktion.ScaleX(MyPic.Width, vbHimetric, vbPixels))
 Hoehe = CLng(B_Rex.Konstruktion.ScaleY(MyPic.Height, vbHimetric, vbPixels))
 x = (B_Rex.Konstruktion.Width / Screen.TwipsPerPixelX) / 6 'bei 1/6 beginnen, bis 5/6
 Br = (B_Rex.Konstruktion.Width / Screen.TwipsPerPixelX) / 6 * 4
 Hoehe = Hoehe / Breite * (Br)
 B_Rex.Konstruktion.Tag = Hoehe 'merken, damit darunter die informationen eingezeichnet werden koennen (TabelleEig_ausfuellen in brex)
 
 Destination.PaintPicture LoadPicture(Pic$), x, 0, Br, Hoehe

 i = Hoehe 'B_Rex.Konstruktion.Tag 'nach dem bild einfuegen wird hier die hoehe desselben uebergeben
 Destination.CurrentY = i
 Destination.CurrentX = Pos11
 Destination.Print "d (11) = " & Sys(11).E(2) & " mm";
 Destination.CurrentX = Pos10
 Destination.Print "d (10) = " & Sys(10).E(2) & " mm"
 Destination.CurrentX = Pos11
 Destination.Print "n = " & Int(Sys(11).E(21)) & " 1/min";
 Destination.CurrentX = Pos10
 Destination.Print "n = " & Int(Sys(10).E(21)) & " 1/min"
 Destination.CurrentX = Pos11
 Destination.Print "ß = " & Int(Sys(11).E(13)) & " °";
 Destination.CurrentX = Pos10
 Destination.Print "ß = " & Int(Sys(10).E(13)) & " °"
 
 Destination.CurrentX = (B_Rex.Konstruktion.Width / Screen.TwipsPerPixelX) / 6 * 1.5 'linke Pos
 Destination.Print Lang_Res(183) & Int(Sys(11).E(73)) & Lang_Res(184)   ' " mm, Zeichnung nicht maßstäblich."'"Achsabstand = "
    
 
   
End Sub
Private Sub Pfadoptimieren(ByVal i As Integer) 'hier werden die beiden elemente vergeben
Dim j As Integer
Dim K As Integer
Dim H As Integer
Dim V$(2)
    If Sys(i).Verb(1, 1) <> 0 And Sys(i).Verb(2, 1) <> 0 Then 'scheiben bei zwei anschlüssen optimieren
        
        'erst die hauptsächliche Richtung bestimmen
        For j = 1 To 2
            H = Sys(i).Top - Sys(Sys(i).Verb(j, 1)).Top
            K = (Sys(Sys(i).Verb(j, 1)).left + Sys(Sys(i).Verb(j, 1)).Width / 2) - (Sys(i).left + Sys(i).Width / 2)
            If H > 0 And Abs(H) > Abs(K) Then V$(j) = "o"
            If H < 0 And Abs(H) > Abs(K) Then V$(j) = "u"
            If K > 0 And Abs(K) > Abs(H) Then V$(j) = "r"
            If K < 0 And Abs(K) > Abs(H) Then V$(j) = "l"
        Next j
        Select Case V$(1) & V$(2)
            Case "or"
                Sys(i).Verb(1, 2) = 2
                Sys(i).Verb(2, 2) = 5
            Case "ro"
                Sys(i).Verb(1, 2) = 5
                Sys(i).Verb(2, 2) = 2
            Case "ru"
                Sys(i).Verb(1, 2) = 4
                Sys(i).Verb(2, 2) = 7
            Case "ur"
                Sys(i).Verb(1, 2) = 7
                Sys(i).Verb(2, 2) = 4
            Case "ul"
                Sys(i).Verb(1, 2) = 6
                Sys(i).Verb(2, 2) = 1
            Case "lu"
                Sys(i).Verb(1, 2) = 1
                Sys(i).Verb(2, 2) = 6
            Case "lo"
                Sys(i).Verb(1, 2) = 8
                Sys(i).Verb(2, 2) = 3
            Case "ol"
                Sys(i).Verb(1, 2) = 3
                Sys(i).Verb(2, 2) = 8
        End Select
    End If
End Sub
Private Sub Elementeverbinden(ByVal E1 As Integer, ByVal E2 As Integer) 'hier werden die beiden elemente vergeben
    'dieses unterprogramm wird immer nur einmal aufgerufen
    'ev. alte verbindungen zw. e1 und e2 müssen vorher getrennt werden
    Dim i As Integer, j As Integer
    Dim N As Single, M As Single, P As Single, Q As Single
    
    Dim MaxLaenge As Double 'eben dort wird dann die länge in mm eingezeichnet
    Dim MaxL1 As Single 'halten die position fest
    Dim MaxL2 As Single
    Dim MaxL3 As Single
    
    Dim Verb(4, 8) As Long 'xE1, yE1, xE2, yE2 aller 8 anschlüsse, die variable hat immer den ursprung, nicht die aktuelle position
    Dim x As Integer 'zum eckenzeichnen
    Dim y As Integer 'zum eckenzeichnen
    Dim X5 As Integer 'koordinaten kommend von e1
    Dim Y5 As Integer 'koordinaten kommend von e1
    Dim X6 As Integer 'ditto e2
    Dim Y6 As Integer 'ditto e2
    Dim X7 As Integer 'zum linienzeichnen, start
    Dim Y7 As Integer 'zum linienzeichnen, start
    Dim X8 As Integer 'zum linienzeichnen, ziel
    Dim Y8 As Integer 'zum linienzeichnen, ziel
    Dim RS$
    Dim R$(2)
    Dim XAusgleich As Boolean
    Dim E2Ausgleich As Boolean
    Dim Ausgleichen As Boolean
    Dim ErstY As Boolean
    Dim XBruecke As Boolean
    
    'verb auffüllen mit allen xy-anschluss-koordinaten
    'XBruecke = False
    MaxLaenge = 0
    XAusgleich = False
    E2Ausgleich = False
    ErstY = False
    
    'positionen aller anschlüsse an beiden elementen bestimmen, jeweils x (1,3) und y(2,4)
    For i = 1 To 4 Step 2
        j = E1
        If i = 3 Then j = E2
        Verb(i, 1) = Sys(j).left
        Verb(i + 1, 1) = Sys(j).Top + 8
        Verb(i, 4) = Sys(j).left + Sys(j).Width
        Verb(i + 1, 4) = Sys(j).Top + 8
        If left(Sys(j).Tag, 1) <> "1" Then 'förderer besitzen diese Anschlüsse nicht
            Verb(i, 2) = Sys(j).left + 8
            Verb(i + 1, 2) = Sys(j).Top
            Verb(i, 3) = Sys(j).left + 22
            Verb(i + 1, 3) = Sys(j).Top
            Verb(i, 5) = Sys(j).left + Sys(j).Width
            Verb(i + 1, 5) = Sys(j).Top + 24
            Verb(i, 6) = Sys(j).left + 22
            Verb(i + 1, 6) = Sys(j).Top + Sys(j).Width
            Verb(i, 7) = Sys(j).left + 8
            Verb(i + 1, 7) = Sys(j).Top + Sys(j).Width
            Verb(i, 8) = Sys(j).left
            Verb(i + 1, 8) = Sys(j).Top + 24
        End If
    Next i
    
    'vorhandene anschlüsse suchen und auf 0 setzen, also zur wiederverwendung ausschließen
    
    For i = 1 To 3 Step 2
        j = E1
        If i = 3 Then j = E2
        P = 0
        If Sys(j).Verb(1, 2) > 0 Then P = Sys(j).Verb(1, 2)
        If Sys(j).Verb(2, 2) > 0 Then P = Sys(j).Verb(2, 2)
        If P <> 0 Then
            
            'die bestehende verbindung selbst ausschließen
            'aber nicht wenn vorab schon beide verbindungen optimal bestimmt wurden
            If Sys(j).Verb(1, 2) = 0 Or Sys(j).Verb(2, 2) = 0 Then Verb(i, P) = 0
            
            Select Case P
                Case 1
                    Verb(i, 2) = 0
                    Verb(i, 3) = 0
                    Verb(i, 5) = 0
                    Verb(i, 7) = 0
                    Verb(i, 8) = 0
                        'diese Zeilen verhindern tangential angesteuerte scheiben
                        'optionalisieren
                        'If Left(Sys(J).Tag, 1) <> "1" Then Verb(I, 4) = 0
                Case 2
                    Verb(i, 1) = 0
                    Verb(i, 3) = 0
                    Verb(i, 4) = 0
                    Verb(i, 6) = 0
                    Verb(i, 8) = 0
                        'If Left(Sys(J).Tag, 1) <> "1" Then Verb(I, 7) = 0
                Case 3
                    Verb(i, 1) = 0
                    Verb(i, 2) = 0
                    Verb(i, 4) = 0
                    Verb(i, 5) = 0
                    Verb(i, 7) = 0
                       'If Left(Sys(J).Tag, 1) <> "1" Then Verb(I, 6) = 0
                Case 4
                    Verb(i, 2) = 0
                    Verb(i, 3) = 0
                    Verb(i, 5) = 0
                    Verb(i, 6) = 0
                    Verb(i, 8) = 0
                       'If Left(Sys(J).Tag, 1) <> "1" Then Verb(I, 1) = 0
                Case 5
                    Verb(i, 1) = 0
                    Verb(i, 3) = 0
                    Verb(i, 4) = 0
                    Verb(i, 6) = 0
                    Verb(i, 7) = 0
                       'If Left(Sys(J).Tag, 1) <> "1" Then Verb(I, 8) = 0
                Case 6
                    Verb(i, 2) = 0
                    Verb(i, 4) = 0
                    Verb(i, 5) = 0
                    Verb(i, 7) = 0
                    Verb(i, 8) = 0
                       'If Left(Sys(J).Tag, 1) <> "1" Then Verb(I, 3) = 0
                Case 7
                    Verb(i, 1) = 0
                    Verb(i, 3) = 0
                    Verb(i, 5) = 0
                    Verb(i, 6) = 0
                    Verb(i, 8) = 0
                       'If Left(Sys(J).Tag, 1) <> "1" Then Verb(I, 2) = 0
                 Case 8
                    Verb(i, 1) = 0
                    Verb(i, 2) = 0
                    Verb(i, 4) = 0
                    Verb(i, 6) = 0
                    Verb(i, 7) = 0
                       'If Left(Sys(J).Tag, 1) <> "1" Then Verb(I, 5) = 0
            End Select
        End If
    Next i
    
    'restliche Anschlüsse auf kürzeste distanz prüfen
    P = 10000 'viel zu hoch ansetzen, wird runtergesetzt
    For i = 1 To 8 'i enthält nummer des ersten Punktes
        For j = 1 To 8
            If Verb(1, i) <> 0 And Verb(3, j) <> 0 Then 'wenn x-koord 0 sind,also ohne Anschluß
                Q = Sqr(Abs(Verb(1, i) - Verb(3, j)) ^ 2 + Abs(Verb(2, i) - Verb(4, j)) ^ 2) 'pythagoras, abstand berechnen
                If Q < P Then
                    P = Q 'die Anschlüsse mit den kleinsten Entfernungen suchen
                    M = i
                    N = j
                End If
            End If
        Next j
    Next i
   
    'ab hier vorsicht, wenn 2 elemente 2mal miteinander verbunden sind
    'verbindungsdaten im Element abspeichern
    
    If Sys(E1).Verb(1, 1) = Sys(E1).Verb(2, 1) Then
        'zweischeibensystem muß anders behandelt werden
        If Sys(E1).Verb(1, 2) = 0 Then
            Sys(E1).Verb(1, 2) = M 'verbindung muß zum verbundenen element passen
        Else
            Sys(E1).Verb(2, 2) = M
        End If
        If Sys(E2).Verb(1, 2) = 0 Then
            Sys(E2).Verb(1, 2) = N
        Else
            Sys(E2).Verb(2, 2) = N
        End If
        'hier gibts nix zu optimieren
    Else
        If Sys(E1).Verb(1, 1) = E2 Then 'folgende bedingungen, weil es vielleicht vorab schon festgelegt wurde
            If Sys(E1).Verb(1, 2) = 0 Then
                Sys(E1).Verb(1, 2) = M 'es wurde vorab noch keine günstigere verbindung erstellt (pfadoptimieren)
            Else
                M = Sys(E1).Verb(1, 2) 'die günstigere verbindung nehmen (aus pfadoptimieren)
            End If
        Else
            If Sys(E1).Verb(2, 2) = 0 Then
                Sys(E1).Verb(2, 2) = M
            Else
                M = Sys(E1).Verb(2, 2)
            End If
        End If
        If Sys(E2).Verb(1, 1) = E1 Then
            If Sys(E2).Verb(1, 2) = 0 Then
                Sys(E2).Verb(1, 2) = N
            Else
                N = Sys(E2).Verb(1, 2)
            End If
        Else
            If Sys(E2).Verb(2, 2) = 0 Then
                Sys(E2).Verb(2, 2) = N
            Else
                N = Sys(E2).Verb(2, 2)
           End If
        End If
    End If
    
    'hier ist genug bekannt, um die Richtung festzulegen
    If Endlos = True Then
        If Sys(E1).Tag = "001" Then 'ist erster Durchlauf (Antriebsscheibe) und geschlossen
            If Reversieren = False Then 'antrieb links drehend
                If M = 1 Or M = 3 Or M = 5 Or M = 7 Then 'dreht auf E2 zu
                    RE2 = True 'richtung e2
                Else
                    RE2 = False 'richtung e1
                End If
            Else 'antrieb rechts drehend
                If M = 1 Or M = 3 Or M = 5 Or M = 7 Then 'dreht auf E1 zu
                    RE2 = False
                Else
                    RE2 = True
                End If
            End If
        End If
        
        'bandlaufrichtung (<>Förderrichtung) im förderer ermitteln
        If left(Sys(E2).Tag, 1) = "1" Then
            Sys(E2).Rechts = False
            If RE2 = True And N = 1 Then Sys(E2).Rechts = True
            If RE2 = False And N = 4 Then Sys(E2).Rechts = True
        End If
    End If
    
Eineliniezeichnen: 'markierung wird bei einfärbung eines trumstückchens nochmal angesprungen
    
    '5 enthält den derzeitigen aufenthaltsort des striches von e1, 6 das Wertepaar von e2
    'striche zunächst von den elementen wegführen
    'nutzen, daß es noch in verb(,) abgespeichert ist
    X5 = Verb(1, M) 'x-koord e1
    Y5 = Verb(2, M) 'y-koord e1
    X6 = Verb(3, N) 'x-koord e2
    Y6 = Verb(4, N) 'y-koord e2
    
    'raus aus dem ersten element
    Destination.ForeColor = QBColor(0) 'schwarz
    i = 13 'muß 13 sein, sonst werden aufsteigende linien durch die elemente verdeckt
    j = 13
    'If Abs(Y6 - Y5) < 20 And Abs(Y6 - Y5) > 0 Then I = 26 'sonst kommts zum winzigen y-ausgleich, also lieber dran vorbei
    'If Abs(X6 - X5) < 20 And Abs(Y6 - Y5) > 0 Then J = 26 'sonst kommts zum winzigen x-ausgleich, also lieber dran vorbei
    If M = 2 Or M = 3 Then 'nach oben raus
        X7 = X5
        X8 = X5
        Y7 = Y5
        Y8 = Y5 - i
        GoSub Linienzeichnen
        Y5 = Y5 - (i - 1)
        R$(1) = "o"
    End If
    If M = 4 Or M = 5 Then 'nach rechts
        X7 = X5
        X8 = X5 + j
        Y7 = Y5
        Y8 = Y5
        GoSub Linienzeichnen
        X5 = X5 + (j - 1)
        R$(1) = "r"
    End If
    If M = 6 Or M = 7 Then 'nach unten
        X7 = X5
        X8 = X5
        Y7 = Y5
        Y8 = Y5 + i
        GoSub Linienzeichnen
        Y5 = Y5 + (i - 1)
        R$(1) = "u"
    End If
    If M = 8 Or M = 1 Then 'nach links
        X7 = X5
        X8 = X5 - j
        Y7 = Y5
        Y8 = Y5
        GoSub Linienzeichnen
        X5 = X5 - (j - 1)
        R$(1) = "l"
    End If
    
    'raus aus dem 2. element
    j = 13
    i = 13
    
    'verb, weil die ursprungswerte zum vergleich herangezoegen werden müssen
    'If Abs(Y6 - Verb(2, M)) < 20 And Abs(Y6 - Verb(2, M)) > 0 Then I = 26 'sonst kommts zum winzigen y-ausgleich, also lieber dran vorbei
    'If Abs(X6 - Verb(1, M)) < 20 And Abs(Y6 - Verb(2, M)) > 0 Then J = 26 'sonst kommts zum winzigen y-ausgleich, also lieber dran vorbei
    If N = 2 Or N = 3 Then 'nach oben raus
        X7 = X6
        X8 = X6
        Y7 = Y6
        Y8 = Y6 - i
        GoSub Linienzeichnen
        Y6 = Y6 - (i - 1)
        R$(2) = "o"
    End If
    If N = 4 Or N = 5 Then 'nach rechts
        X7 = X6
        X8 = X6 + j
        Y7 = Y6
        Y8 = Y6
        GoSub Linienzeichnen
        X6 = X6 + (j - 1)
        R$(2) = "r"
    End If
    If N = 6 Or N = 7 Then 'nach unten
        X7 = X6
        X8 = X6
        Y7 = Y6
        Y8 = Y6 + i
        GoSub Linienzeichnen
        Y6 = Y6 + (i - 1)
        R$(2) = "u"
    End If
    If N = 8 Or N = 1 Then 'nach links
        X7 = X6
        X8 = X6 - j
        Y7 = Y6
        Y8 = Y6
        GoSub Linienzeichnen
        X6 = X6 - (j - 1)
        R$(2) = "l"
    End If
    j = X5 'merken
    
    'schönheitskorrektur
    If (M = 2 Or M = 3 Or M = 6 Or M = 7) And (N = 1 Or N = 4 Or N = 5 Or N = 8) Then 'erspart n paar biegungen
        If Abs(Verb(2, M) - (Verb(2, M) + Verb(4, N)) / 2) > Abs(Y5 - (Verb(2, M) + Y5) / 2) Then XBruecke = True 'aber nicht, wenn sie zu nah beieinander sind in y
        If Abs(Verb(1, M) - (Verb(1, M) + Verb(3, N)) / 2) > Abs(X5 - (Verb(1, M) + X5) / 2) Then XBruecke = True 'aber nicht, wenn sie zu nah beieinander sind in x
    End If
    
    'von e1 in richtung e2, x-richtung ausgleichen
    Ausgleichen = True
    If (M = 6 Or M = 7) And (N = 1 Or N = 4 Or N = 5 Or N = 8) And Y5 < Y6 Then Ausgleichen = False
    If (M = 2 Or M = 3) And (N = 1 Or N = 4 Or N = 5 Or N = 8) And Y5 > Y6 Then Ausgleichen = False
    
        'nicht ausgleichen, wären zwei biegungen zuviel
        'einschränkung, um überflüssige Ecken zu vermeiden (verhindert unter Umständen 2 extrabiegungen, die nix tun)
    If Ausgleichen = True Then
        If Abs(X5 - X6) > 0 Then 'einen x-Ausgleich schlucken, wenn die elemente zu nah beieinander (oder gleich) liegen
            'soll nicht entgegengesetzt starten
            If X5 > X6 Then
                R$(1) = R$(1) & "l" 'nach links
            Else
                R$(1) = R$(1) & "r" 'nach rechts
            End If
            If R$(1) <> "lr" And R$(1) <> "rl" Then 'sonst würde der zweite strich zurück durchs element führen
                X7 = X5 'braucht man e
                Y7 = Y5
                Y8 = Y5
                If (Y5 > Y6 And R$(2) = "o") Or (Y5 < Y6 And R$(2) = "u") Then 'nicht ganz ausgleichen, sonst trifft sich alles im element
                    X8 = (X6 + X5) / 2
                    GoSub Linienzeichnen
                    y = Y5
                    x = X5
                    RS$ = R$(1)
                    GoSub Eckeeinbauen
                    X5 = (X6 + X5) / 2
                Else
                    X8 = X6
                    GoSub Linienzeichnen
                    y = Y5
                    x = X5
                    RS$ = R$(1)
                    GoSub Eckeeinbauen
                    X5 = X6
                End If
                XAusgleich = True
            Else
                R$(1) = left(R$(1), 1) 'den neuen eintrag wieder rückgängig machen
            End If
        End If
    End If
    
    'richtungsumkehr, weil aus der sicht der y-gerade alles anders herum aussieht
    R$(1) = Right(R$(1), 1)
    If R$(1) = "l" Then
        R$(1) = "r"
    Else
        If R$(1) = "r" Then R$(1) = "l"
    End If
    
    'von e2 in richtung e1, x-richtung ausgleichen, falls noch etwas übrig ist
    i = 0 'und zwar immer und sei es, um nur an die ecke zu kommen
    If X6 > j Then
        R$(2) = R$(2) & "l" 'nach links
        i = -2
    Else
        R$(2) = R$(2) & "r" 'nach rechts
        i = 2
    End If
    If X6 <> j And X5 <> X6 Then 'j ist ein oben gemerktes x5
        'soll nicht entgegengesetzt starten
        If R$(2) <> "lr" And R$(2) <> "rl" Then 'sonst würde der zweite strich zurück durchs element führen
            'folgende unterscheidung, damit nicht x aus e2 durch e1 hindurchgeführt wird
            If ((Right$(R$(1), 1) = "r" And Right(R$(2), 1) = "l") Or (Right$(R$(1), 1) = "l" And Right(R$(2), 1) = "r")) And Y5 = Y6 Then 'beide y nach unten wegziehen
                X7 = X6
                X8 = X6
                Y7 = Y6 'y nach unten bewegen auf einer seite, die andere wird unten aut. hinzugefügt
                y = Y6
                Y6 = Y6 + 26
                Y8 = Y6
                GoSub Linienzeichnen
                x = X6
                R$(2) = Right(R$(2), 1) + "u"
                RS$ = R$(2)
                GoSub Eckeeinbauen
                If X5 > X6 Then R$(2) = "lo"
                If X5 < X6 Then R$(2) = "ro"
                ErstY = True
            End If
            X7 = X6
            X8 = X5
            Y7 = Y6
            Y8 = Y6
            GoSub Linienzeichnen
            y = Y6
            If X5 = X6 Then y = Y5 'der ganze x-ausgleich war schon, wir brauchen nur noch die ecke
            
            x = X6
            RS$ = R$(2)
            GoSub Eckeeinbauen
            
            If XAusgleich = False Then
                x = X5
                If X5 > X6 Then
                    RS$ = R$(1) & "l"
                Else
                    RS$ = R$(1) & "r"
                End If
                GoSub Eckeeinbauen
            End If
                        
            XAusgleich = True
            E2Ausgleich = True
            X6 = X5
           
        Else
            R$(2) = left(R$(2), 1) 'den neuen eintrag wieder rückgängig machen
            E2Ausgleich = False
        End If
    Else 'kein xausgleich ab e2 und kein yausgleich, also ecke nachzeichnen
        If Y5 = Y6 And XAusgleich = True Then
            y = Y6
            x = X6
            RS$ = R$(2)
            GoSub Eckeeinbauen
            XAusgleich = True
            X6 = X5
        End If
        R$(2) = left(R$(2), 1) 'den neuen eintrag wieder rückgängig machen
    End If
    
    'richtungsumkehr, weil aus der sicht der y-gerade alles anders herum aussieht
    R$(2) = Right(R$(2), 1)
    If R$(2) = "l" Then
        R$(2) = "r"
    Else
        If R$(2) = "r" Then R$(2) = "l"
    End If
    
    
    '2 y-strahlen nach oben erzwingen
    If Y5 = Y6 And XAusgleich = False Then 'band würde durch das element zurücklaufen
        
        X7 = X5
        X8 = X5
        Y7 = Y5
        Y8 = Y5 - 30
        GoSub Linienzeichnen
        If left(R$(1), 1) = "r" Then
            RS$ = "lo"
        Else
            RS$ = "ro"
        End If
        y = Y5
        x = X5
        GoSub Eckeeinbauen
        Y5 = Y5 - 30
        
        X7 = X6
        X8 = X6
        Y7 = Y6
        Y8 = Y6 - 30
        GoSub Linienzeichnen
        If left(R$(2), 1) = "r" Then
            RS$ = "lo"
        Else
            RS$ = "ro"
        End If
        y = Y6
        x = X6
        GoSub Eckeeinbauen
        Y6 = Y6 - 30
        
        If X5 > X6 Then
            RS$ = "ol"
        Else
            RS$ = "or"
        End If
        y = Y5
        x = X5
        GoSub Eckeeinbauen
        
        If X6 > X5 Then
            RS$ = "ol"
        Else
            RS$ = "or"
        End If
        y = Y6
        x = X6
        GoSub Eckeeinbauen
    End If
    
YAusgleich:
    'y wird immer ausgleichen
    If Abs(Y5 - Y6) > 5 Then
        If Y5 < Y6 Then 'sonst bleibt ein pixel lücke, was weiß ich warum
            i = 2
        Else
            i = -2
        End If
        If Y5 < Y6 Then
            R$(1) = Right(R$(1), 1) & "o"
            R$(2) = Right(R$(2), 1) & "u"
        Else
            R$(1) = Right(R$(1), 1) & "u"
            R$(2) = Right(R$(2), 1) & "o"
        End If
        If XAusgleich = True Then 'Or XBruecke = True Then 'wenigstens ein x-ausgleich hat schon stattgefunden, kann normal weitergehen
            X7 = X5
            X8 = X5
            Y7 = Y5
            Y8 = Y6 + i
            GoSub Linienzeichnen
            x = X5
            y = Y5
            'o und u müssen unter y-ausgleich immer links stehen
            'und x muß ein bißchen versetzt werden
            If Right(R$(1), 1) = "o" Or Right(R$(1), 1) = "u" Then R$(1) = Right(R$(1), 1) & left(R$(1), 1)
            RS$ = R$(1)
            GoSub Eckeeinbauen
            y = Y6
            R$(2) = Right(R$(2), 1) & left(R$(2), 1)   'umdrehen
            RS$ = R$(2)
            
            'If XBruecke = True Then 'empirisch
                If left(R$(1), 1) = "u" Then
                    R$(1) = "o" & Right(R$(1), 1)
                Else
                    R$(1) = "u" & Right(R$(1), 1)
                End If
            'End If
            
            If E2Ausgleich = False And (N = 1 Or N = 8 Or N = 4 Or N = 5) Then
                GoSub Eckeeinbauen
            Else
                If ErstY = True Then RS$ = R$(1)
                If E2Ausgleich = True Then GoSub Eckeeinbauen
            End If
            Y5 = Y6
        Else 'gab keinen x-ausgleich, also ist r$(1/2) auch nur 1 zeichen lang
            'dann eben erst beide y bis zur Hälfte ausgleichen
            X7 = X5
            X8 = X5
            Y7 = Y5
            Y8 = (Y6 + Y5 + i) / 2
            GoSub Linienzeichnen
            x = X5
            y = Y5
            RS$ = R$(1)
            RS$ = Right(RS$, 1) & left(RS$, 1) 'umdrehen
            GoSub Eckeeinbauen
            y = (Y5 + Y6) / 2
            If Right(R$(1), 1) = "u" Then
                R$(1) = "o"
            Else
                R$(1) = "u"
            End If
            If X6 < X5 Then
                R$(1) = R$(1) & "l"
            Else
                R$(1) = R$(1) & "r"
            End If
            RS$ = R$(1)
            If X6 <> X5 Then GoSub Eckeeinbauen
            
            'y aus zweiter richtung
            X7 = X6
            X8 = X6
            Y7 = Y6
            Y8 = (Y6 + Y5) / 2
            GoSub Linienzeichnen
            x = X6
            y = Y6
            RS$ = R$(2)
            RS$ = Right(RS$, 1) & left(RS$, 1) 'umdrehen
            GoSub Eckeeinbauen
            y = (Y5 + Y6) / 2
            If Right(R$(2), 1) = "u" Then
                R$(2) = "o"
            Else
                R$(2) = "u"
            End If
            If X6 < X5 Then
                R$(2) = R$(2) & "r"
            Else
                R$(2) = R$(2) & "l"
            End If
            RS$ = R$(2)
            If X6 <> X5 Then GoSub Eckeeinbauen
            Y5 = (Y5 + Y6) / 2
        End If
    End If
    
    'falls es oben keinen x-ausgleich gab, hier nun der entgültige zusammenschluß
    If X5 <> X6 Then
        If XBruecke = False Then
            If X6 > X5 Then
                X6 = X6 - 10
                X5 = X5 + 10
            Else
                X6 = X6 + 10
                X5 = X5 - 10
            End If
        End If
        X7 = X5
        X8 = X6
        Y7 = Y5
        Y8 = Y5
        GoSub Linienzeichnen
    End If
    
    'und noch die längen eintragen, optional
    If ModusCalc <> "liniensuchen" Then
        If Sys(E1).Verb(1, 1) = E2 And Sys(E1).Verb(1, 2) = M Then
            RS$ = Str(Int(Sys(E1).Verb(1, 3)))
        Else
            RS$ = Str(Int(Sys(E1).Verb(2, 3)))
        End If
        If MaxL1 < 0 Then 'y-strecke beschriften
            Destination.CurrentY = ((Abs(MaxL1) + Abs(MaxL2)) / 2) - Destination.TextHeight("A") / 2
            Destination.CurrentX = MaxL3 + 3
        Else 'x-strecke beschriften
            Destination.CurrentX = ((MaxL1 + MaxL2) / 2) - Destination.TextWidth(Str(RS$)) / 2
            Destination.CurrentY = MaxL3 + 5 'darunter
        End If
        If RS$ > 0 Then
            Destination.Print RS$
            'Destination.Print Replace(RS$, "0", "O") 'weil er nun mal nullen durch lange schwarze balken ersetzt. danke, bill!
        End If
    End If
Exit Sub
    
Eckeeinbauen:
    If ModusCalc = "liniensuchen" And E3 <> E1 And E4 <> E2 Then Return
    B_Rex.Elbilder.ClipHeight = 15
    B_Rex.Elbilder.ClipWidth = 15
    B_Rex.Elbilder.ClipY = 105
    'If E3 > 0 And M = EA3 And N = EA4 Then Elbilder.ClipY = 105 + 16 'markierte ecken
    If E3 = E1 And E4 = E2 Then B_Rex.Elbilder.ClipY = 105 + 16
    If RS$ = "ol" Or RS$ = "ru" Then 'ok
        B_Rex.Elbilder.ClipX = 291
        Destination.PaintPicture B_Rex.Elbilder.Clip, x - 10, y - 3, 15, 15
    End If
    If RS$ = "lu" Or RS$ = "or" Then 'ok
        B_Rex.Elbilder.ClipX = 307
        Destination.PaintPicture B_Rex.Elbilder.Clip, x - 3, y - 3, 15, 15
    End If
    If RS$ = "ur" Or RS$ = "lo" Then 'ok
        B_Rex.Elbilder.ClipX = 323
        Destination.PaintPicture B_Rex.Elbilder.Clip, x - 3, y - 10, 15, 15
    End If
    If RS$ = "ro" Or RS$ = "ul" Then
        B_Rex.Elbilder.ClipX = 339
        Destination.PaintPicture B_Rex.Elbilder.Clip, x - 10, y - 10, 15, 15
    End If
Return

Linienzeichnen:
    'immer von x7,Y7 nach x8,y8
    If ModusCalc = "liniensuchen" And E3 = 0 Then
        'eine linie wurde getroffen?
        If Abs(Y7 - (Y7 + Y8) / 2) + 6 > Abs(Y2 - (Y7 + Y8) / 2) Then
            If Abs(X7 - (X7 + X8) / 2) + 6 > Abs(X2 - (X7 + X8) / 2) Then
                E3 = E1
                E4 = E2
                EA3 = M 'speichern die anschlüsse, um diese ein strecke nochmal zu zeichnen
                EA4 = N
                GoTo Eineliniezeichnen 'nochmal durchlaufen, aber mit zeichnen und ohne diese kontrolle
            End If
        End If
        Return 'war nix, müssen wir auch nicht nochmal zeichnen
    End If
        
    If ModusCalc = "liniensuchen" And E3 <> E1 And E4 <> E2 Then Return
        
    If X7 = X8 Then 'y-linie ziehen
        If E3 = E1 And E4 = E2 Then
            Destination.Line (X7 - 2, Y7)-(X7 + 3, Y8), QBColor(2), BF 'dunkelgrün
        Else
            Destination.Line (X7 - 2, Y7)-(X7 + 3, Y8), QBColor(10), BF
        End If
        Destination.Line (X7 - 3, Y7)-(X7 - 3, Y8) ', QBColor(0)
        Destination.Line (X7 + 4, Y7)-(X7 + 4, Y8) ', QBColor(0)
        If LYStrang < Abs(Y8 - Y7) Then
            LYStrang = Abs(Y8 - Y7)
            LYStrangx = X7 + 5
            LYStrangy = (Y7 + Y8) / 2 - 8
        End If
    End If
    If Y7 = Y8 Then 'x-linie ziehen
        If E3 = E1 And E4 = E2 Then
            Destination.Line (X7, Y7 - 2)-(X8, Y7 + 3), QBColor(2), BF
        Else
            Destination.Line (X7, Y7 - 2)-(X8, Y7 + 3), QBColor(10), BF
        End If
        Destination.Line (X7, Y7 - 3)-(X8, Y7 - 3) ', QBColor(0)
        Destination.Line (X7, Y7 + 4)-(X8, Y7 + 4) ', QBColor(0)
        
        'richtungspfeile anzeigen'auf den x-Strängen, die ab der antriebsscheibe kommen
        If Endlos = True And Abs(X7 - X8) > 65 Then '65 ist genug platz
            'If X5 = X7 Or X6 = X7 Then 'x-linie kommt tatsächlich von e1
                B_Rex.Elbilder.ClipHeight = 12
                B_Rex.Elbilder.ClipWidth = 46
                B_Rex.Elbilder.ClipY = 442 'nach links
                If X6 = X7 Then B_Rex.Elbilder.ClipY = 456 'nach rechts
                B_Rex.Elbilder.ClipX = 0
                If E3 = E1 And E4 = E2 Then B_Rex.Elbilder.ClipX = 47 'markiert
                If X7 < X8 Then 'antriebsscheibe links
                    'RE2 = In Richtung Element 2
                    If RE2 = True Then 'nach r2 also nach rechts
                        B_Rex.Elbilder.ClipY = 456
                        If X6 = X7 Then B_Rex.Elbilder.ClipY = 442
                    End If
                    Destination.PaintPicture B_Rex.Elbilder.Clip, (X7 + X8) / 2 - 23, Y7 - 6, 46, 12 'breite und höhe sind zwar optional, aber ohne diese angaben wirds nicht skaliert
                Else
                    If RE2 = False Then
                        B_Rex.Elbilder.ClipY = 456
                        If X6 = X7 Then B_Rex.Elbilder.ClipY = 442
                    End If
                End If
                Destination.PaintPicture B_Rex.Elbilder.Clip, (X7 + X8) / 2 - 23, Y7 - 6, 46, 12
            'End If
        End If
    End If
    
    'die längste Strecke merken, um die länge daneben zu schreiben
    If Abs(X7 - X8) > MaxLaenge Then
        MaxLaenge = Abs(X7 - X8)
        MaxL1 = X7
        MaxL2 = X8
        MaxL3 = Y7
    End If
    If Abs(Y7 - Y8) > MaxLaenge Then
        MaxLaenge = Abs(X7 - X8)
        MaxL1 = -Y7 ',mit minus wird die y-richtung sichtbar gemacht
        MaxL2 = -Y8
        MaxL3 = X7
    End If
Return

End Sub
Private Sub Element_aufbauen(ByVal Element As Single, ByVal x As Long, ByVal y As Long)
    Dim i As Integer
    Dim j As Integer
    Dim P$
    
    B_Rex.Elbilder.ClipWidth = 32
    'If Left(Sys(Element).Tag, 1) = "2" Then 'der träger wird statt des huckepacks gezeichnet
    '    X = Sys(Sys(Element).Zugehoerigkeit).Left
    '    Y = Sys(Sys(Element).Zugehoerigkeit).Top
    '    Element = Sys(Element).Zugehoerigkeit 'statt huckepack lieber den Träger zeichnen
    'End If
    For i = 32 To Sys(Element).Width Step 32 'bild aufbauen
        If left(Sys(Element).Tag, 1) = "0" Then
            B_Rex.Elbilder.ClipHeight = 32
            B_Rex.Elbilder.ClipX = 0
            B_Rex.Elbilder.ClipY = 6 * 35 + (Right(Sys(Element).Tag, 1) - 1) * 33
            'messerkantenbild fällt aus der systematik
            If Sys(Element).Tag = "005" Then B_Rex.Elbilder.ClipY = 342 - 33
            
            'ein Anschluß vorhanden
            If (Sys(Element).Verb(1, 1) = 0 And Sys(Element).Verb(2, 1) <> 0) Or (Sys(Element).Verb(2, 1) = 0 And Sys(Element).Verb(1, 1) <> 0) Then
                j = Sys(Element).Verb(1, 2)
                If j = 0 Then j = Sys(Element).Verb(2, 2)
                If j = 8 Or j = 5 Then B_Rex.Elbilder.ClipX = B_Rex.Elbilder.ClipX + 165
                If j = 2 Or j = 7 Then B_Rex.Elbilder.ClipX = B_Rex.Elbilder.ClipX + 198
                If j = 3 Or j = 6 Then B_Rex.Elbilder.ClipX = B_Rex.Elbilder.ClipX + 231
                If j = 1 Or j = 4 Then B_Rex.Elbilder.ClipX = B_Rex.Elbilder.ClipX + 264
            End If
            
            'zwei anschlüsse vorhanden
            If Sys(Element).Verb(1, 1) <> 0 And Sys(Element).Verb(2, 1) <> 0 Then
                P$ = Format(Sys(Element).Verb(1, 2)) & Format(Sys(Element).Verb(2, 2)) '(verbindungsstelle, mit element..)
                If P$ = "74" Or P$ = "47" Then B_Rex.Elbilder.ClipX = B_Rex.Elbilder.ClipX + 33
                If P$ = "61" Or P$ = "16" Then B_Rex.Elbilder.ClipX = B_Rex.Elbilder.ClipX + 66
                If P$ = "83" Or P$ = "38" Then B_Rex.Elbilder.ClipX = B_Rex.Elbilder.ClipX + 99
                If P$ = "25" Or P$ = "52" Then B_Rex.Elbilder.ClipX = B_Rex.Elbilder.ClipX + 132
                If P$ = "85" Or P$ = "58" Then B_Rex.Elbilder.ClipX = B_Rex.Elbilder.ClipX + 165
                If P$ = "27" Or P$ = "72" Then B_Rex.Elbilder.ClipX = B_Rex.Elbilder.ClipX + 198
                If P$ = "36" Or P$ = "63" Then B_Rex.Elbilder.ClipX = B_Rex.Elbilder.ClipX + 231
                If P$ = "14" Or P$ = "41" Then B_Rex.Elbilder.ClipX = B_Rex.Elbilder.ClipX + 264
            End If
            Destination.PaintPicture B_Rex.Elbilder.Clip, x, y, 32, 32
        End If
        If left(Sys(Element).Tag, 1) = "1" Then 'förderer
            B_Rex.Elbilder.ClipHeight = 34
            B_Rex.Elbilder.ClipY = (Right(Sys(Element).Tag, 1) - 1) * 35
            If i = 32 Then 'linker Teil
                B_Rex.Elbilder.ClipX = 0
            End If
            If i = Sys(Element).Width Then 'rechter Teil
                B_Rex.Elbilder.ClipX = 64
            End If
            If i > 32 And i < Sys(Element).Width Then 'alle anderen Teile
                B_Rex.Elbilder.ClipX = 32
            End If
            For j = 1 To Maxelementindex 'nach Transportgut suchen
                If Sys(j).Zugehoerigkeit = Element And Sys(j).Tag = "201" Then
                    If Abs(Sys(j).left - x) < 10 Then ' das richtige Trägerstückechen abpassen
                        B_Rex.Elbilder.ClipX = B_Rex.Elbilder.ClipX + 97
                    End If
                End If
            Next j
            If Sys(Element).Verb(1, 1) > 0 Or Sys(Element).Verb(2, 1) > 0 Then B_Rex.Elbilder.ClipX = B_Rex.Elbilder.ClipX + 97 * 2
            Destination.PaintPicture B_Rex.Elbilder.Clip, x, y - 11, 32, 34 'mit der 11 auf die richtige Höhe bringen
            For j = 9 To Maxelementindex 'nach weiteren huckepacks suchen
                If Sys(j).Zugehoerigkeit = Element And Sys(j).Tag <> "201" Then
                    If Abs(Sys(j).left - x) < 10 Then ' das richtige Trägerstückechen abpassen
                        B_Rex.Elbilder.ClipHeight = 32
                        If Sys(j).Tag = "204" Then B_Rex.Elbilder.ClipX = 129
                        If Sys(j).Tag = "205" Then B_Rex.Elbilder.ClipX = 129 - 32
                        If Sys(j).Tag = "206" Then B_Rex.Elbilder.ClipX = 129 + 32
                        B_Rex.Elbilder.ClipY = 105
                        Destination.PaintPicture B_Rex.Elbilder.Clip, x, y - 32, 32, 32
                    End If
                End If
            Next j
        End If
        x = x + 32
    Next i    'transportgut auf verändertem Förderer wiederherstellen
    DoEvents
End Sub

