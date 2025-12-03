Attribute VB_Name = "Dateiverwaltung"
Option Explicit

Private ActAnlage As Integer
Public Sub Undo(Mode As Integer)
'mode ab 0: b_rex
'mode ab 20: prolink2

Static Aktiv As Boolean
Static Rueckwaerts(2) As Integer
Static Vorwaerts(2) As Integer
'1 kennzeichnet den status von b_rex
'2 kennzeichnet den status von prolink2
If Aktiv = True Then Exit Sub 'soll sich nicht selbst aufrufen
Aktiv = True
Dim i As Integer

    Select Case Mode
        'jeweils die beiden knöpfe, die vorwärts- und die rückwärtsschritte verwalten
        'und actanlage von 0 bis 10 zirkulieren lassen
        Case 0 '0= einfach nur den status aufnehmen und weiterzählen
            Gespeichert = False
            ActAnlage = ActAnlage + 1
            If ActAnlage > 10 Then ActAnlage = 0 'dann eben wieder von vorne
            Vorwaerts(1) = 0
            Call Einlesen 'in den string nämlich
            Rueckwaerts(1) = Rueckwaerts(1) + 1
            If Rueckwaerts(1) > 10 Then Rueckwaerts(1) = 10
        Case 1 '1= den vorhergehenden Status aktivieren
            Rueckwaerts(1) = Rueckwaerts(1) - 1
            If Rueckwaerts(1) <= 0 And left(B_Rex.Datei(10).Tag, 1) = "E" Then 'weiter zurück gehts nicht
                Call Mother.Knopfverwaltung(10, "GrosserKnopf", "Button", "B_Rex")
            End If
            If Rueckwaerts(1) < 0 Then
                Rueckwaerts(1) = 0  ' nichts gespeichertes mehr da
            Else
                ActAnlage = ActAnlage - 1
                If ActAnlage < 0 Then ActAnlage = 10 'dann eben wieder von vorne
                Call Auslesen
                Call Wiederherstellen
                Vorwaerts(1) = Vorwaerts(1) + 1
                
            End If
        Case 2 '2= den nächsten status wiederherstellen, falls vorher mit 1 rückgängig gemacht wurde, nie mehr Schritte als diese
            Vorwaerts(1) = Vorwaerts(1) - 1
            If Vorwaerts(1) <= 0 And left(B_Rex.Datei(11).Tag, 1) = "E" Then 'weiter vorwärts gehts nicht
                Call Mother.Knopfverwaltung(11, "GrosserKnopf", "Button", "B_Rex")
            End If
            If Vorwaerts(1) < 0 Then
                Vorwaerts(1) = 0 ' nichts gespeichertes mehr da
            Else
                ActAnlage = ActAnlage + 1
                If ActAnlage > 10 Then ActAnlage = 0 'dann eben wieder von vorne
                Call Auslesen
                Call Wiederherstellen
                Rueckwaerts(1) = Rueckwaerts(1) + 1
                If left(B_Rex.Datei(10).Tag = "A", 1) Then Call Mother.Knopfverwaltung(10, "GrosserKnopf", "Button", "B_Rex")
                
            End If
            
        Case 3 'aktuelle anlage in die datenbank aufnehmen
            On Local Error Resume Next
            If Sam(19) = "" Then
                i = MsgBox("Tragen Sie zuerst einen Projektnamen im B_Rex-ID ein", vbOKOnly) 'Unter diesem Namen ist bereits eine andere Anlage abgespeichert. Wollen Sie diese Anlage überschreiben?
                Exit Sub
            End If
            i = MsgBox("Anlage wird als >" & Sam(19) & "< in die Datenbank aufgenommen", vbOKCancel)
            If i = vbCancel Then Exit Sub
            Call Einlesen
            Artikeldaten.Execute "insert into beispielanlagen (bezeichnung, anlage) values ('" & Sam(19) & "', '" & Anlage(ActAnlage) & "')"
            Mother.H = "aufgenommen"
            Call Beispielanlagen_einrichten
        Case 4 'eine beispielanlage wurde gewählt
            'If Gespeichert = False Then I = MsgBox(Lang_Res(20), vbOKCancel) '"ihre bisherige arbeit geht verloren
            'If I = vbCancel Then Exit Sub
            Anlage(ActAnlage) = Bspanlagen("anlage")
            Call Auslesen(1)
            Call Wiederherstellen
            Mother.H = Lang_Res(21)
            If left(B_Rex.Datei(5).Tag, 1) = "A" Then Call B_Rex.Datei_Click(5) 'zeichnung
            If left(B_Rex.Datei(6).Tag, 1) = "A" Then Call B_Rex.Datei_Click(6) 'eigenschaften
            If left(B_Rex.Datei(7).Tag, 1) = "A" Then Call B_Rex.Datei_Click(7) 'fu kurve
            B_Rex.Konstruktion.SetFocus
        Case 5 'aktuelle anlage aus datenbank löschen, falls vorhanden
            On Local Error Resume Next
            i = MsgBox("Anlage >" & B_Rex.Beispielanlagen & "< wird aus den Beispielanlagen gelöscht", vbOKCancel)
            If i = vbCancel Then Exit Sub
            Artikeldaten.Execute "delete * from beispielanlagen where bezeichnung = '" & B_Rex.Beispielanlagen & "'"
            Mother.H = "gelöscht"
            Call Beispielanlagen_einrichten
    End Select
Aktiv = False
End Sub
Private Sub Wiederherstellen()
    On Local Error Resume Next
    If Sys(1).S(2) <> SystemTyp.Artnr Then
        Call Code1.SystemTyp_Set(Sys(1).S(2))
    End If
    
    Call Eigschaftsverr.Zwei_Scheiben
    DoEvents '
'''    Call CodeDraw.Alleelementeverbinden
    DoEvents
    Call CodeCalc.Rechnungssteuerung("EVC")
    
    'den rev - knopf noch richtig rum drehen
        If Reversieren = True Then
            If B_Rex.Datei(9).Tag = "A" Then Call Mother.Knopfverwaltung(9, "GrosserKnopf", "Button", "B_Rex")
        Else
            If B_Rex.Datei(9).Tag = "E" Then Call Mother.Knopfverwaltung(9, "GrosserKnopf", "Button", "B_Rex")
        End If
    
    Lastaktel = 0
    Call B_Rex.Tabelle_ausfuellen(0)
    Call Code1.B_Rex_Uebersetzen

    Gespeichert = True

End Sub

Private Sub Einlesen() 'anlagenauslegung
Dim a$, B$, C$
Dim i As Integer, j As Integer

    a$ = "#EOS" & vbCrLf
    
    C$ = ""
    
    C$ = C$ & "[Version]" & a$
        C$ = C$ & Vers_B_Rex & a$ & vbCrLf
    
    C$ = C$ & "[Head]" & a$
        C$ = C$ & "B_Rex, Program by" & a$
        C$ = C$ & "Dirk Wiederholt, Dep. IT" & a$
        C$ = C$ & "SIEGLING GmbH, Lilienthalstraße 6/8, Hannover 30179, Germany" & a$
        C$ = C$ & "dirk.wiederholt@siegling.com" & a$ & vbCrLf
    
    C$ = C$ & "[Type]" & a$
        If Dateioffen = "" Then C$ = C$ & "first creation date " & Date & a$
        C$ = C$ & "conveyor/drive unit calculation" & a$ & vbCrLf
    
        For i = 1 To Maxelementindex 'anzahl der zu speichernden Datensätze feststellen
            If Sys(i).Element <> "" Then
                C$ = C$ & "#BOT[Content]" & a$
                C$ = C$ & "(a)number=" & i & a$
                C$ = C$ & "(b)element=" & Sys(i).Element & a$
                C$ = C$ & "(c)height=" & Sys(i).Height & a$
                C$ = C$ & "(d)width=" & Sys(i).Width & a$
                C$ = C$ & "(e)top=" & Sys(i).Top & a$
                C$ = C$ & "(f)left=" & Sys(i).left & a$
                C$ = C$ & "(g)tag=" & Sys(i).Tag & a$
                C$ = C$ & "(h)Zugehoerigkeit=" & Sys(i).Zugehoerigkeit & a$
                
                
                C$ = C$ & "(i)1.Verbindung mit=" & Sys(i).Verb(1, 1) & a$
                C$ = C$ & "(j)2.Verbindung mit=" & Sys(i).Verb(2, 1) & a$
                C$ = C$ & "(k)Länge zu 1=" & Sys(i).Verb(1, 3) & a$
                C$ = C$ & "(l)Länge zu 2=" & Sys(i).Verb(2, 3) & a$
                C$ = C$ & "(o)f zu 1=" & Sys(i).Verb(1, 4) & a$
                C$ = C$ & "(p)f zu 2=" & Sys(i).Verb(2, 4) & a$
                
                'meistens ist es false, dann garnichts tun
                If Sys(i).Rechts = True Then
                    C$ = C$ & "(m)rechts=" & 1 & a$
                Else
                    C$ = C$ & "(m)rechts=" & 0 & a$
                End If
                If Sys(i).Vollstaendig = True Then
                    C$ = C$ & "(n)vollstaendig=" & 1 & a$
                Else
                    C$ = C$ & "(n)vollstaendig=" & 0 & a$
                End If
                
                For j = 1 To Eigenschaftszahl 'betrifft die zahleneigenschaften
                    If Sys(i).E(j) <> 0 Then
                        'texte einsparen um datei zu verkleinern
                        'C$ = C$ & "(" & J & ")" & El(J).Eigenschaft & "=" & Sys(I).E(J)
                        C$ = C$ & "(" & j & ")" & "=" & Sys(i).E(j)
                        If Sys(i).B(j) = True Then
                            C$ = C$ & " #*" & a$
                        Else
                            C$ = C$ & a$
                        End If
                    End If
                Next j
                For j = TextEigenschaftszahl To 1 Step -1  'betrifft die texteigenschaften
                    If Sys(i).S(j) <> "" Then
                        'texte einsparen um datei zu verkleinern'beeinträchtigt nicht die lesefunktion
                        'C$ = C$ & "(" & -J & ")" & El(-J).Eigenschaft & "=" & Sys(I).S(J)
                        C$ = C$ & "(" & -j & ")" & "=" & Sys(i).S(j)
                        If Sys(i).B(-j) = True Then
                            C$ = C$ & " #*" & a$
                        Else
                            C$ = C$ & a$
                        End If
                    End If
                Next j
                C$ = C$ & vbCrLf
            End If
        Next i
            
    C$ = C$ & "#BOT[General properties]" & a$
        If Reversieren = True Then
            C$ = C$ & "(a)Reversieren=" & 1 & a$
        Else
            C$ = C$ & "(a)Reversieren=" & 0 & a$
        End If
        If Endlos = True Then
            C$ = C$ & "(b)Endlos=" & 1 & a$
        Else
            C$ = C$ & "(b)Endlos=" & 0 & a$
        End If
        C$ = C$ & "(c)=" & Init_B_Rex_rho_Wert_Fehler & a$
        C$ = C$ & "(d)=" & Init_B_Rex_FwFu_Fehler & a$
        C$ = C$ & "(e)=" & Init_B_Rex_KraftUebertrkontr & a$
        C$ = C$ & "(f)=" & Init_B_Rex_WoelbDurchb & a$
        C$ = C$ & "(g)=" & Init_B_Rex_Minddurchmkontr & a$
        C$ = C$ & "(h)=" & B_Rex.Button(0).Tag & a$
        C$ = C$ & "(i)=" & SystemTyp.Kraftdehnung & a$
        C$ = C$ & "(j)=" & SystemTyp.KraftdehnungMode & a$
        C$ = C$ & "(k)=" & Init_B_Rex_Aging & a$
        C$ = C$ & "(l)=" & Init_B_Rex_Pairing & a$
        C$ = C$ & vbCrLf

            
    C$ = C$ & "#BOT[Customer/User]" & a$
        For i = 0 To 20 'benutzerdaten
            C$ = C$ & "(" & i & ")C/U" & "=" & Sam(i) & a$
        Next i
        C$ = C$ & vbCrLf

    Anlage(ActAnlage) = C$

End Sub
Private Sub Auslesen(Optional Mode As Integer) 'anlagenauslegung
    '0 = alles einlesen
    '1 = benutzer/userdaten nicht einlesen, bei beispielanlagen werden dann nicht immer die benutzerdaten ueberschrieben
    
Dim i As Double, j As Double, L As Double, M As Double, N As Double, ElNummer As Integer
Dim a$, Klammer$, Text$, Wert$
    'erst klarschiff
        For i = 0 To 20 'userdaten löschen
            Sam(i) = ""
        Next i
        B_Rex.GrKl(0).Visible = False
        B_Rex.GrKl(1).Visible = False
        For i = 0 To Maxelementezahl 'sicher ist sicher
            Sys(i) = Del 'falls jemand sys(0) verhunzt hat, das eigentlich hierfür ist
        Next i
        Sys(1).Element = "Band" 'damits aufgerufen wird, ob was drinsteht oder nicht
        Sys(2).Element = "Band" 'damits mit gespeichert wird
        Sys(1).Tag = "301" 'damits aufgerufen wird, ob was drinsteht oder nicht
        
        SystemTyp.Kraftdehnung = 0 'kann sd, fw, k1 oder selbstgewaehlt sein
        SystemTyp.KraftdehnungMode = 0
        Init_B_Rex_Pairing = 1 'alte anlagen sind immer noch mit pairing

    
    Maxelementindex = 10
    i = 1
    j = 1
    Do
        ElNummer = 0
        i = InStr(i, Anlage(ActAnlage), "#BOT[")
        If i = 0 Then Exit Sub ' das wars, mehr gibts net
        j = InStr(i + 1, Anlage(ActAnlage), "]")
        a$ = LCase(Mid(Anlage(ActAnlage), i + 5, j - i - 5))
        L = InStr(j, Anlage(ActAnlage), "#BOT[") 'hier findet sich der nächste eintrag
        'Sys(ElNummer).Rechts = False
        Select Case a$
            Case "content"
                Do
                    'If ElNummer = 12 Then Stop
                    i = InStr(j + 1, Anlage(ActAnlage), "(")
                    j = InStr(i + 1, Anlage(ActAnlage), "#EOS")
                    a$ = Mid(Anlage(ActAnlage), i, j - i)
                    GoSub String_auslesen
                    Select Case Val(Klammer$)
                        Case 0 'angaben ohne felder
                            Select Case Klammer$
                                Case "a"
                                    ElNummer = Val(Wert$) 'wird immer zuerst ausgelesen
                                Case "b"
                                    Sys(ElNummer).Element = Wert$
                                Case "c"
                                    Sys(ElNummer).Height = Val(Wert$)
                                Case "d"
                                    Sys(ElNummer).Width = Val(Wert$)
                                Case "e"
                                    Sys(ElNummer).Top = Val(Wert$)
                                Case "f"
                                    Sys(ElNummer).left = Val(Wert$)
                                Case "g"
                                    Sys(ElNummer).Tag = Wert$
                                Case "h"
                                    Sys(ElNummer).Zugehoerigkeit = Wert$
                                Case "i"
                                    Sys(ElNummer).Verb(1, 1) = CDBLVAL(Wert$)
                                Case "j"
                                    Sys(ElNummer).Verb(2, 1) = CDBLVAL(Wert$)
                                Case "k"
                                    Sys(ElNummer).Verb(1, 3) = CDBLVAL(Wert$)
                                Case "l"
                                    Sys(ElNummer).Verb(2, 3) = CDBLVAL(Wert$)
                                Case "o"
                                    Sys(ElNummer).Verb(1, 4) = CDBLVAL(Wert$)
                                Case "p"
                                    Sys(ElNummer).Verb(2, 4) = CDBLVAL(Wert$)
                                Case "m"
                                    Sys(ElNummer).Rechts = CBool(Val(Wert$))
                                Case "n"
                                    Sys(ElNummer).Vollstaendig = CBool(Val(Wert$))
                            End Select
                        Case Is > 0 'felder zahlen
                            Sys(ElNummer).E(Val(Klammer$)) = CDBLVAL(Wert$)
                            If InStr(1, Wert$, "#*") > 0 Then Sys(ElNummer).B(Val(Klammer$)) = True
                            If Val(Klammer$) = 82 Then
                                SystemTyp.Kraftdehnung = CDBLVAL(Wert$) 'kann sd, fw, k1 oder selbstgewaehlt sein
                                SystemTyp.KraftdehnungMode = 2
                            End If
                        Case Is < 0 'felder texte
                            Sys(ElNummer).S(Abs(Val(Klammer$))) = Wert$
                            If InStr(1, Wert$, "#*") > 0 Then Sys(ElNummer).B(Val(Klammer$)) = True
                    End Select
                Loop Until InStr(j + 1, Anlage(ActAnlage), "(") > L Or InStr(j + 1, Anlage(ActAnlage), "(") = 0 'kommt nix mehr
                'Stop
            Case "general properties"
                Do
                    i = j + 6
                    j = InStr(i + 1, Anlage(ActAnlage), "#EOS")
                    a$ = Mid(Anlage(ActAnlage), i, j - i)
                    GoSub String_auslesen
                    Select Case LCase(Klammer$)
                        Case "a" 'angaben ohne felder
                            Reversieren = CBool(Val(Wert$))
                        Case "b"
                            Endlos = CBool(Val(Wert$))
                        Case "c"
                            Init_B_Rex_rho_Wert_Fehler = Val(Wert$)
                        Case "d"
                            Init_B_Rex_FwFu_Fehler = Val(Wert$)
                        Case "e"
                            Init_B_Rex_KraftUebertrkontr = Val(Wert$)
                        Case "f"
                            Init_B_Rex_WoelbDurchb = Val(Wert$)
                        Case "g"
                            Init_B_Rex_Minddurchmkontr = Val(Wert$)
                        Case "h"
                        Case "i"
                            SystemTyp.Kraftdehnung = CDBLVAL(Wert$)
                        Case "j"
                            SystemTyp.KraftdehnungMode = Val(Wert$)
                        Case "k"
                            Init_B_Rex_Aging = Val(Wert$)
                        Case "l"
                            Init_B_Rex_Pairing = Val(Wert$)
                    End Select
                Loop Until InStr(j + 1, Anlage(ActAnlage), "(") > L Or InStr(j + 1, Anlage(ActAnlage), "(") = 0 'kommt nix mehr
            Case "customer/user"
                Do
                    'If ElNummer = 12 Then Stop
                    i = j + 5
                    j = InStr(i + 1, Anlage(ActAnlage), "#EOS")
                    a$ = Mid(Anlage(ActAnlage), i, j - i)
                    GoSub String_auslesen
                    If Mode = 0 Then Sam(Val(Klammer$)) = Wert$
                Loop Until InStr(j + 1, Anlage(ActAnlage), "(") = 0 'kommt nix mehr
            Case Else
                i = i + 1
        End Select
        If ElNummer > Maxelementindex Then Maxelementindex = ElNummer

    Loop Until InStr(j, Anlage(ActAnlage), "#BOT[") = 0 'end of file

    'A = CDbl("123,234E-04")
Exit Sub

String_auslesen:
    N = InStr(1, a$, "(")
    M = InStr(1, a$, ")")
    If N > 0 Then
        Klammer$ = Mid(a$, N + 1, M - N - 1) 'gibts überhaupt ne klammer?
    Else
        M = 0
    End If
    N = M
    M = InStr(N + 1, a$, "=")
    Text$ = Mid(a$, N + 1, M - N - 1)
    N = M
    M = Len(a$)
    Wert$ = Mid(a$, N + 1, M - N)
    Return
End Sub
Public Sub Beispielanlagen_einrichten()
Static ZweitesMal As Boolean 'ab dem zweiten mal erst schliessen vor öffnen
    On Error Resume Next
    If ZweitesMal = True Then Bspanlagen.Close
    B_Rex.Beispielanlagen.Clear
    
    Bspanlagen.Open "select * from beispielanlagen order by bezeichnung", Artikeldaten, adOpenStatic, adLockPessimistic
    
    ZweitesMal = True
    Bspanlagen.MoveFirst
    Do
        B_Rex.Beispielanlagen.AddItem Bspanlagen("bezeichnung")
        Bspanlagen.MoveNext
    Loop Until Bspanlagen.EOF
End Sub







