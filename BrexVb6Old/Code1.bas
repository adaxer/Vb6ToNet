Attribute VB_Name = "Code1"
Option Explicit

Private Elementinfos As New ADODB.Recordset
Private Konstanten As New ADODB.Recordset
Private D1 As New Scripting.Dictionary

Public Sub DatbaSteuerung1(Mode As Integer)
    Select Case Mode
        Case 1 'nikita datenbank
            Call Artikeldaten_open
        Case 10 'alle schliessen

            If Artikeldaten.State = 1 Then Artikeldaten.Close
    End Select
End Sub


Private Sub Artikeldaten_open()
Dim a As Integer
Dim Connected As Boolean
    On Local Error GoTo Errorhandler
    

            If FileExist(App.Path & "\pd2000.mdb") And Artikeldaten.State = 0 Then
                Artikeldaten.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\pd2000.mdb" & ";Persist Security Info=False"
                Connected = True
            End If
    
    If Connected = True Then Call Code1.Datenbank_einlesen
    
    Exit Sub

Errorhandler:
    MsgBox ("Fehler beim einlesen der Datenbank:" & App.Path & "\pd2000.mdb ")

End Sub



Private Sub Objekt_Uebersetzen(O As Object, Mode As Integer)
'0 caption wird uebersetzt aufgrund einer zahl in der tag-info (labels)
'1 tooltiptext wird uebersetzt aufgrund einer zahl in der tag-info (bsp: "#03xxxxx", x steht fuer zahl)
'GEHT NICHT MIT COMBOBOXEN
Dim i As Integer, j As Double
Dim a$
On Error Resume Next
    For i = 0 To O.Count + 6 'zur sicherheit, es fehlen wohl immer ein paar
       a$ = ""
       Select Case Mode
        Case 0
            a$ = Lang_Res(Val(O(i).Tag))
            If a$ <> "" Then O(i).Caption = Imperial_Metric_unit(a$)
        Case 1
            j = 0
            If InStr(O(i).Tag, "#03") Then
                j = Val(IstDrin(O(i).Tag, "", 13))
            Else
                j = Val(O(i).Tag)
            End If
            a$ = Lang_Res(j)
            If a$ <> "" Then O(i).ToolTipText = Imperial_Metric_unit(a$) 'wird dort schon fertig uebersetzt
        End Select
    Next i

End Sub

Public Sub B_Rex_Uebersetzen()
Dim i, j As Integer
    With B_Rex
        
        Call Objekt_Uebersetzen(.Datei, 1) 'wegen E in arbeitsoeberflaeche noch nicht moeglich
        
        'tag bei element fuer was anderes gebraucht
        .Element(1).ToolTipText = Lang_Res(115) 'primäre Antriebsscheibe
        .Element(2).ToolTipText = Lang_Res(116) 'Antriebsscheibe
        .Element(3).ToolTipText = Lang_Res(117) 'Umlenkscheibe
        .Element(4).ToolTipText = Lang_Res(118) 'Messerkante
        .Element(5).ToolTipText = Lang_Res(119) 'passiver Abweiser
        .Element(6).ToolTipText = Lang_Res(120) 'Stau
        .Element(7).ToolTipText = Lang_Res(121) 'Transportgut
        .Element(8).ToolTipText = Lang_Res(122) 'trägergebundene Umfangskraft, freidefinierbar
        .Element(9).ToolTipText = Lang_Res(123) 'Tisch
        .Element(10).ToolTipText = Lang_Res(124)  'Tragrollenbahn
        .Element(11).ToolTipText = Lang_Res(125) 'Rollenbahn
        .Element(12).ToolTipText = Lang_Res(126) 'freidefinierbare Umfangskraft, Trägereigenschaften
        
        
        Call Objekt_Uebersetzen(.EigButton, 1) 'wegen E in arbeitsoeberflaeche noch nicht moeglich

        
        .Trumlängeneinheit = Lang_Res(181) 'Truml. [mm]
        
        .Beispielanlagen.ToolTipText = Lang_Res(179) 'beispiele
        Call Dateiverwaltung.Beispielanlagen_einrichten


    End With
    
    Sys(1).Element = Lang_Res(168) 'Band
    Sys(2).Element = Lang_Res(168) 'Band
    For i = 10 To Maxelementezahl
        Select Case Sys(i).Tag
            Case "001" 'primAntriebsscheibe
                Sys(i).Element = Lang_Res(115)
            Case "002" 'Antriebsscheibe
                Sys(i).Element = Lang_Res(116)
            Case "003" 'Umlenkscheibe
                Sys(i).Element = Lang_Res(117)
            Case "005" 'Messerkante
                Sys(i).Element = Lang_Res(118)
            Case "205" 'Abweiser
                Sys(i).Element = Lang_Res(119)
            Case "204" 'Stau
                Sys(i).Element = Lang_Res(120)
            Case "201" 'Transportgut
                Sys(i).Element = Lang_Res(121)
            Case "206" 'trägergebundene_Umfangskraft
                Sys(i).Element = Lang_Res(122)
            Case "101" 'Tisch
                Sys(i).Element = Lang_Res(123)
            Case "102" 'Tragrollenbahn
                Sys(i).Element = Lang_Res(124)
            Case "103" 'Rollenbahn
                Sys(i).Element = Lang_Res(125)
            Case "104" 'freie_Umfangskraft
                Sys(i).Element = Lang_Res(126)
        End Select
    Next i
    Call Datenbank_einlesen
End Sub
Public Sub Datenbank_einlesen()
Dim j As Integer
Dim i As Integer
    On Error Resume Next 'meist wegen fehlender englischer datenbankeinträge
    'eigentlich nur nötig, wenn b_rex freigeschaltet wäre:
    'elementinformationen einlesen, sprachabhängig, darum hier
    
    Elementinfos.Open "Elemente_a", Artikeldaten, adOpenStatic, adLockReadOnly
    
    Elementinfos.MoveFirst 'sonst kriegt er recordcount nicht raus
    Elementinfos.MoveLast 'sonst kriegt er recordcount nicht raus
    Elementinfos.MoveFirst
    For j = 0 To Elementinfos.RecordCount - 1
        'Die "negativen Datenbankeinträge werden automatisch mit eingelesen"
        
        If Elementinfos("eigenschaft") <> "" Then El(Elementinfos("nummer")).Eigenschaft = Elementinfos("eigenschaft")
        
        If Elementinfos("einheit") <> "" Then El(Elementinfos("nummer")).Einheit = Elementinfos("einheit")
        If Elementinfos("Feldart") <> "" Then El(Elementinfos("nummer")).Feldart = Elementinfos("Feldart")
        If Elementinfos("minimum") <> 0 Then El(Elementinfos("nummer")).Minimum = Elementinfos("minimum")
        If Elementinfos("maximum") <> 0 Then El(Elementinfos("nummer")).Maximum = Elementinfos("maximum")
        For i = 6 To 27
            If Elementinfos(i) <> "" Then El(Elementinfos("nummer")).Eig(i) = Elementinfos(i)
        Next i
        Elementinfos.MoveNext
    Next j
    Elementinfos.Close 'sind gespeichert, damit kann das programm zügiger arbeiten
    Call Eigschaftsverr.Zwei_Scheiben 'variable zweischeiben in ordnung bringen, die wird durch das einlesen der elemente verstellt
    
    Konstanten.Open "konstanten order by nummer", Artikeldaten, adOpenStatic, adLockReadOnly
    
    Konstanten.MoveFirst
    Konstanten.MoveLast 'sonst kriegt er recordcount nicht raus
    Konstanten.MoveFirst
    
    For j = 1 To Konstanten.RecordCount
        Kst(Konstanten("Nummer")).zuEigenschaft = Konstanten("zu_Eigenschaft")
        Kst(Konstanten("Nummer")).ID = Konstanten("Nummer")
        
        Kst(Konstanten("Nummer")).Bezeichnung = Konstanten("bezeichnung")
        Kst(Konstanten("Nummer")).Einstellung = Konstanten("einstellung")
        Konstanten.MoveNext
    Next j
    Konstanten.Close 'sind gespeichert, damit kann das programm zügig arbeiten
    
    'datenbanksprachänderungen
        'gibts überhaupt schon ne datenbank?
        Mother.H = "Hier später die Pick_It Daten Laden"
            Sys(1).S(1) = "Typ('typ')"
            Sys(1).S(3) = "Typ('beschtsfs')" & " " & "Typ('oberfltsfs')"
            Sys(1).S(4) = "Typ('beschlsas')" & " " & "Typ('oberfllsas')"
End Sub

Public Function CDBLVAL(a$) As Double
    If a$ = "" Then
        CDBLVAL = 0
    Else
        a$ = Replace(a$, ",", ".")
        CDBLVAL = Val(a$)
        'cdbl schlägt aus unerfindlichen gründen nicht selten fehl
    End If
End Function
Public Function StringDot(a As Double) As String

StringDot = Str(a)
StringDot = Replace(StringDot, " ", "")
StringDot = Replace(StringDot, ",", ".")

End Function

Public Function StringzuZahl(a$) As Double  'mit punkt drin
    StringzuZahl = CDbl(a$)
End Function


Public Function Selfround(ByVal a As Double) As Double
'vb6 round geht oft schief, also sicherheit einbauen

Selfround = Round(a)
If Abs(Selfround - a) > 1 Then
    Selfround = a 'wenn scheisse passiert ist, den wert unveraendert zurueckschicken
End If


End Function

Public Function InstrBrex(Search$, SearchChar$) As Boolean
Dim i As Integer
    For i = 1 To Len(Search$)
        If LCase(Mid$(Search$, i, 1)) = LCase(SearchChar$) Then 'immer nur ein zeichen wird verglichen
            InstrBrex = True
            Exit For
        End If
    Next
End Function
Public Function Elementnummer(a$) As Integer
    Elementnummer = 5 'neuverwendung m
    Do 'ab 5 beginnen erst die elemente
        Elementnummer = Elementnummer + 1
    Loop Until El(0).Eig(Elementnummer) = a$ '0 enthält die elementnamen
End Function

Public Function IstDrin(a$, B$, Mode As Integer) As String
Dim L As Integer, K As Integer
On Error Resume Next
    Select Case Mode
        Case 0 'nach eingruppierung suchen
            L = InStr(1, a$, "#01") + 2
            K = InStr(L, a$, "#")
            If InStr(1, a$, B$) > L And InStr(1, a$, B$) < K Then IstDrin = "J"
        Case 1 'nach a$ irgendwo da drin suchen 'etwa markiert oder datenblatt
            If InStr(1, a$, B$) > 0 Then IstDrin = "J"
        Case 2 'bereich zwischen () übergeben, also feldnamen
            L = InStr(1, a$, "#02") + 2
            K = InStr(L, a$, "#")
            IstDrin = Mid(a$, L + 1, K - L - 1)
        Case 3 'bezeichnung aufgrund der ressourcennummer, die steht hinter der #03
            L = InStr(1, a$, "#03") + 3
            K = InStr(L, a$, "#")
            IstDrin = Lang_Res(1100 + Mid(a$, L, K - L))
        Case 4 'bereich zwischen [] übergeben
            L = InStr(1, a$, "#04")
            If L > 0 Then
                K = InStr(L + 2, a$, "#")
                IstDrin = Mid(a$, L + 3, K - L - 3)
            End If
        Case 5 'gruppenZugehoerigkeit zurückgeben
            L = InStr(1, a$, "#01") + 2
            K = InStr(L, a$, "#")
            If InStr(1, a$, B$) > L And InStr(1, a$, B$) < K Then
                IstDrin = Mid(a$, InStr(1, a$, B$), 3) 'nämlich gruppe mit nummer, zB 'G01'
            End If
        Case 6 '?
            L = InStr(1, a$, "#SU") + 2
            K = InStr(L, a$, "#")
            If L > 2 Then IstDrin = Mid(a$, L + 1, K - L - 1)
        Case 7 ' wenn '06' drin, dann koerper, nicht kopf, nur kt
            L = InStr(1, a$, "#06") + 2
            If L = 2 Then
            Else
                K = InStr(L, a$, "#")
                If L > 1 Then IstDrin = Mid(a$, L + 1, K - L - 1)
            End If
        Case 9 'gruppenZugehoerigkeit zurückgeben
            L = InStr(1, a$, "#01") + 2
            K = InStr(L, a$, "#")
            If InStr(1, a$, "G") > L And InStr(1, a$, "G") < K Then
                IstDrin = Mid(a$, InStr(1, a$, "G") + 1, 2) 'nämlich nur nummer, zB '01'
            End If
        Case 10 'einfach bereich hinter b$ wiedergeben
            If InStr(1, a$, B$) > 0 Then
                L = InStr(1, a$, B$) + 2
                'I = Len(A$)
                K = InStr(L, a$, "#")
                IstDrin = Mid(a$, L + 1, K - L - 1)
            End If
        Case 11
            'bereich zwischen () übergeben, also feldnamen
            L = InStr(1, a$, "#07") + 2
            K = InStr(L, a$, "#")
            IstDrin = Mid(a$, L + 1, K - L - 1)
            If L = 2 Then IstDrin = ""
        Case 12 'bezeichnung aufgrund der ressourcennummer, die steht hinter der #12
            L = InStr(1, a$, "#03") + 3
            K = InStr(L, a$, "#")
            IstDrin = Lang_Res(Mid(a$, L, K - L))
        Case 13 'numer hinter der 12
            L = InStr(1, a$, "#03") + 3
            K = InStr(L, a$, "#")
            IstDrin = Mid(a$, L, K - L)
        Case 14 'juli 2010, stellen zwischen ##...## entfernen (vererben rollenbemerkungen)
            L = InStr(1, a$, "##") + 2
            IstDrin = a$
            If L > 0 Then
                K = InStr(L, a$, "##")
                If K > 0 Then
                    IstDrin = Replace(a$, Mid(a$, L - 2, K - L + 4), "")
                    'IstDrin = Replace(IstDrin, "####", "")
                    
                End If
            End If


    End Select

End Function


Public Function FileExist(Dateiname$) As Boolean
On Error GoTo Errorhandler

    If Dateiname$ = "" Then
        FileExist = False
    Else
        FileExist = Dir$(Dateiname$) <> ""
    End If


Exit Function


Errorhandler:

    FileExist = False
    Resume Next

End Function

Public Function Lang_Res(i As Double) As String
Static Second As Boolean
Dim a$
Dim B$
Dim L As Integer
Dim K As Integer

If Second = False Then
    On Local Error Resume Next
    Second = True
    Call Artikeldaten_open 'datenbank oeffnen, falls noch nicht geschehen
    rsLang_Res.Open "select * from language_ress", Artikeldaten, adOpenForwardOnly, adLockReadOnly 'so gehts rasend schnell :-)
    K = rsLang_Res.RecordCount
    If K = 0 Then Exit Function
    Do Until rsLang_Res.EOF
        L = L + 1
        a$ = ""
        B$ = "de" & rsLang_Res("ID")
        If IsNull(rsLang_Res("german")) = False Then
            If rsLang_Res("german") <> "" Then 'nur ausgefuellte felder
                a$ = rsLang_Res("german")
                D1.Add B$, a$ 'zahlen/Datenfelder als schluessel nicht akzeptiert
            End If
        End If
        rsLang_Res.MoveNext
    Loop
    rsLang_Res.Close
End If

On Local Error GoTo Errorhandler

Lang_Res = D1.Item("de" & i)
Exit Function

Errorhandler:
'Stop
End Function


Public Sub SystemTyp_Set(Typ$) '202008 zentralisieren
'artnr wird uebergeben
    Typ$ = Replace(Typ$, " ", "") '202408 weil sich auch mal ein fuehrendes space einschleicht
    SystemTyp.Artnr = Replace(Typ$, vbCrLf, "") 'nur fuer den fall
    SystemTyp.Netto = left(SystemTyp.Artnr, 6) 'wobei das eigentlich nicht mehr vorkommt
    'hier wuerde eigentlich noch bandauflegen zur aufnahme der physikalischen daten hingehoeren, gelegentlich
End Sub
Public Function Imperial_Metric_unit(a$, Optional Mode As Integer) As String
'mode 1 feet zulaessig, sonst bleibts bei in (laenge, nicht aber breite)

    If InStr(a$, "[mm]") = 0 And InStr(a$, "[in]") = 0 And InStr(a$, "[ft]") = 0 Then 'da gibts nix zu aendern
        Imperial_Metric_unit = a$
        Exit Function
    End If

    Imperial_Metric_unit = Replace(a$, " [mm]", " ")
    Imperial_Metric_unit = Replace(Imperial_Metric_unit, " [in]", " ")
    Imperial_Metric_unit = Replace(Imperial_Metric_unit, " [ft]", " ")

Dim Target As Integer
Target = Init_Imperial
If Mode = 0 And Target = 2 Then Target = 1

    Select Case Target
        Case 0
            Imperial_Metric_unit = Imperial_Metric_unit & "[mm]"
        Case 1
            Imperial_Metric_unit = Imperial_Metric_unit & "[in]"
        Case 2
            Imperial_Metric_unit = Imperial_Metric_unit & "[ft]"
    End Select

End Function
