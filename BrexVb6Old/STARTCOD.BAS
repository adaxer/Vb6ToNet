Attribute VB_Name = "startcode"
Option Explicit

'versionsverwaltungsvariablen
Public Const Vers_B_Rex As String = "3.3.0340" '20250702  letzte umstellung
Public Const B_Rex_Yellow As Single = &H80FFFF 'mittelgelb
Public Const B_Rex_Green As Single = &HC0FFC0 'pastellgrün
Public rsLang_Res As New ADODB.Recordset
Public SystemZeile As Integer 'und da steht er in der tabelle
Public Systemartikel_bewertet As Boolean ' registrieren, um bei änderungen wieder zu löschen

Public Type Styp 'zusammensetzung aus 2 datenbanken, deswegen leider redundant
    Artnr As String  'nur hier einstellen, dann werden gleich ein paar aufgaben mehr erledigt' Call Code1.SystemTyp_Set(ArtListe.TextMatrix(1, 4))
    Typenreihe As String
    Netto As String 'artnr kann schon mal mehr als 6 zeichen umfassen, dann sind hier auf jeden fall nur 6 drin
    Name As String
    ACX As String
    Gewicht As Single
    Dicke As Single
    'Lagen As Single'unbenutzt
    MinLng As Double
    MinBrt As Double
    Kraftdehnung As Single
    KraftdehnungMode As Single
        '1 = sd
        '2 = fw
        '3 = k1
        '4 = selbstgewaehlt
    Zugtraeger As String
        'PA
        'Aramid
        'PET
    rho As Double
    Zahnabstand As String
    RZ As String
    TGR As String 'manschettenregeln, wird derzeit nicht benutzt. gibt gruppenzugehörigkeit bei extremultus an, da erst sind die manschetten drin
    Verbindung01 As String
    Fach5101 As String 'artikelspezifisches fach der japaner, einstweilen wird das nicht gebraucht
    ProjektXPeriod_aut As Integer 'posps spinnerei, unbenutzt
    ProjektXPeriod_man As Integer 'posps spinnerei, unbenutzt
End Type
Public SystemTyp As Styp


'sprachenvariablen (deutsch, englisch, etc...)
'    Public Sp As Double 'bedienersprache 0=deutsch, 5000=Englisch
    Public Abfragesprache As String 'also "en" oder "jp" oder "de"
    Public Datenblattsprache As String 'also "en" oder "jp" oder "de"

'Datenbank
    Public Artikeldaten As New ADODB.Connection
    Public Typ As New ADODB.Recordset
    Public Bspanlagen As New ADODB.Recordset

    
    Public Datenoperation_ok As Boolean
    Public DataSheetNew_ok As Boolean
    
    Type Tag
        AE As String 'an/aus A oder E ganz vorne
        ToolTipLang As Double 'beginnt mit 03
        Lang As String ' beginnt mit 20
    End Type
    
'cheaten
Public Admin As Boolean


'b_rex
Public Anlbreite As Integer
Public Anlhöhe As Integer
Public ModusCalc As String 'merkt sich den aktuellen Modus
Public AktEl As Integer
Public Lastaktel As Integer
Public NeuEl As Integer
Public E3 As Integer 'speichert e1 bei markierter Verbindung
Public E4 As Integer
Public EA3 As Integer 'speichert den Anschluß der markierten verbindung
Public EA4 As Integer
Public RE2 As Boolean 'gibt die orientierung der laufrichtungspfeile vor
Public Anlage(10) As String 'string, der die anlage enthält, zum laden, zum speichern, zum undo und nochmal do, siehe dateiverwaltung

'Fukurve
'public, um sie für die maus im hauptfenster abtastbar zu machen
Public FuScaleX1, FuScaleX2, FuScaleY1, FuScaleY2 As Double
Public AuflTrumKraft As Double 'enthält die im letzten durchlauf ermittelte auflegetrumkraft
Public Fehlerwert As Integer 'addiert die fehler auf zur Bewertung der Anlage, s. liste unter in codecalc, enthält fehler der aktuel aufgelegten anlage
Public FehlerwertSchwingungen As Integer 'addiert die fehler auf zur Bewertung der Anlage, s. liste unter in codecalc
Public FehlerwertLongSchwing As Integer 'anlagenspezifisch, wird daher nur einmal pro rechnung erfasst

'organisation, drucken, kundenverwaltung
Public Destination As Object 'hierein wird gedruckt
Public Druck As Boolean
Public B_Rex_AutoLauf As Boolean

Public Drucky As Single
Public STSchrift As Single
Public Geraetschrift As Single  'unveränderliche schrift für den datenblattaufbau
Public AktSeite As Single 'wird gerade "durchdacht"
Public GezSeite As Single 'wird gerade "gedruckt"

'allgemeine anlagenangaben
Public Zweischeiben As Boolean
Public AnlageRefresh As Boolean

Public Endlos As Boolean
Public Vollstaendig As Boolean 'wenn alle elemente Vollstaendig sind, endlos wird extra gewertet


Public Und As Boolean 'und oder oder bei der abfrage

Type El 'liest elementsteuerdaten aus datenbank ein
    Eigenschaft As String
    Einheit As String
    Feldart As String 'liste oder text
    Minimum As Single
    Maximum As Single
    Eig(30) As String
        '1 sind Muß-Eingaben
        '2 sind Kann-Eingaben
        '3 sind Ergebnisse
        '4 sind Muß-Eingaben, die aber nicht ausgefüllt werden müssen
        '5 sind spitzenlastbezogene Angaben
        '6 sind durchbiegungsbezogene angaben
        '7 frequenzberechnete angaben
        '8 auch frequenzberechnete angaben?
End Type

'speicherung aufbau der datenblätter
Public Type Inhalt
    'Datenfeld As Single
    Feldname As String
    Feldinhalt As String
    Bezeichnung As String
    Tag As String
    Top As Integer
    Height As Integer
End Type
Public MVorl(2, 30) As String '1 enthält die bezeichnung des inhalts, 2 enthält die codierte zusammenstellung des inhalts

Public Const Feldinfogröße As Integer = 76
Public Inhalt(Feldinfogröße) As Inhalt 'zum aktuel dargestellten typen die aus mvorl(,) ausgelesene vorlage

Public Feldinfo(Feldinfogröße) As String

Type Kst
    ID As Integer
    zuEigenschaft As Integer
    Bezeichnung As String
    Einstellung As Single
    
End Type
Public Kst(300) As Kst

'rund um die elementinformationen
Public Const Eigenschaftszahl As Integer = 120
Public El(-10 To Eigenschaftszahl) As El
Public Const TextEigenschaftszahl As Integer = 10

Type s
    Height As Integer
    Width As Integer
    Top As Integer
    left As Integer
    'Index As Integer 'verweis auf element per zahl
    Element As String
    Tag As String
    Rechts As Boolean 'richtung bei förderern
    Vollstaendig As Boolean
    Zugehoerigkeit As Integer 'Huckepacks zum Träger (Elementindex)
    E(Eigenschaftszahl) As Double
    s(TextEigenschaftszahl) As String 'dient der speicherung der info-felder
    b(-TextEigenschaftszahl To Eigenschaftszahl) As Boolean
    Verb(2, 4) As Double
        '1,1 verbundenes element
        '1,2 pos.der verb., wird bei jedem bildschirmaufbau neu erzeugt
        '1,3 Länge des freien Trums;
        '1,4 eigenfrequenz dieses trumstueckchens
        '2,x entsprechend mit dem zweiten element
        'alle Angaben finden sich jeweils in beiden beteiligten Elementen
    Anzverb As Single
    Lrein As Single
    Lraus As Double
    Furein As Double
    Furaus As Double
    FureinSp As Double 'bloss bei scheiben zur durchbiegung
    FurausSp As Double 'bloss bei scheiben zur durchbiegung
    Fusteig As Double
    FusteigSp As Double 'enthält die Spitzenlast
    FusteigSpRoll As Double 'enthält die Spitzenlast
End Type
Public Const Maxelementezahl As Integer = 150 'soviele elemente darf die anlage besitzen
Public Maxelementindex As Integer 'höchster benutzter index, damit nicht komplett durchgezählt werden muß
'0 leeres feld, um informationen zwischenzuspeichern, wenn das programm mit den bandeigenschaften spielt
'1 reserviert für band, darf verändert werden, hinweis, wenn verändert
'2 reserviert für band, darf nicht verändert werden, dient zum vergleich mit 2
'3-9 bleiben leer für was auch immer
'del ist zum löschen
Public Sys(Maxelementezahl) As s
Public Del As s 'nur zum löschen der anderen

Public BarCodeFont As String

Public Ursprung$
Public Hilfetext$
Public Sam(30) 'beinhaltet userdaten
Public Const PI As Single = 3.141592654


Public Reversieren As Boolean
Public Abbruch As Boolean
Public Merk As Single
Public Gespeichert As Boolean

Public Berechtigungsstatus As String

Public Dateioffen As String
Public Dateipfad As String

Public cUserName As String
Public cPCName As String
Public NeueUebersicht As Boolean 'nikita, damit beim Übersichtsaufruf nicht immer wieder aufgebaut wird

'Initialisierung
        
    'alle, da drin sind die einstellungen fuer das registrierungsverzeichnis
        Public Const Init_SettingDir As String = "B_Rex_31"
        Public Init_Location As String
        Public Init_Imperial As Integer
        
    'b_rex
        Public Init_B_Rex_Scheibenbreitefixieren As Integer
        Public Init_B_Rex_ScheibenbreitenueberhangProz As Single
        Public Init_B_Rex_ScheibenbreitenUeberhangGrenze As Single
        
        Public Init_B_Rex_rho_Wert_Fehler As Integer
        Public Init_B_Rex_FwFu_Fehler As Integer
        Public Init_B_Rex_FuNenn_Fehler As Integer
        Public Init_B_Rex_KraftUebertrkontr As Integer
        Public Init_B_Rex_WoelbDurchb As Integer
        Public Init_B_Rex_Minddurchmkontr As Integer
        Public Init_B_Rex_Schw_nur_Ex As Integer
        Public Init_B_Rex_Schw_alle As Integer
        Public Init_B_Rex_Aging As Integer
        Public Init_B_Rex_Pairing As Integer
        
'tagverzeichnis der inhaltsfelder
    
    '#01 Die Großbuchstaben in der Tag-Eigenschaft sind Kürzel mit folgender Bedeutung:
        'Y = darf nicht geaendert werden
        'X = sekundäres feld, mehrere datensaetze sind betroffen, nicht ohne weiteres freischalten
        'S = feld ist zur suche zugelassen
        'Z = zahlenfeld, mit kommata
        'V = zahlenfeld, ohne kommata
        'E = sprachabhaengiges feld
        'D = darf bei bedarf aufs datenblatt
        'O = Boolean
        'R = aus anderen feldern errechnet, keine aenderungsmoeglichkeit
        'E = sprachrelevantes Feld,kann englisch, franzzösisch, spanisch sein
        'I = diese feld wird nicht von fertigen algorithmen gepflegt, es ist modulspezifisch und wird extra gepflegt
        
        'Nikita speziell
            'S = sonstige Daten
            'G = gruppe zugehoerig (immer gefolgt von zahl)
            'P = Physikdaten
            'A = Aufbaudaten
            'B = B_Rex-Daten
            'L = lims daten
            'M = schweizer daten
            'N = Daten werden nicht direkt aufs Datenblatt übertragen, höchstens angehängt an eine andere Information
            'X = penta daten
            'T = technische Daten (für Datenblattaufbau)
        
    '#02 feldname in der access-datenbank
    '#03 verständliche bezeichnung in der tabelle, ersetzen durch zeiger auf ressource
    '#04 einheit, wenn vorhanden
    '#05 übersetzung von 03
    '#06 Kopf oder Körper, derzeit nur KT
    '#07 Tabellenname, aus dem das feld nach #02 entnommen ist

    Public UserStatus As Integer 'kunde 1 oder mitarbeiter 2


Sub Main()
    Call Code1.DatbaSteuerung1(1)
    Load B_Rex
    B_Rex.Show
    DumpData
End Sub

Public Sub DumpData()
    Dim fn As Integer
    fn = FreeFile
    Open "C:\temp\vb6dump.json" For Output As #fn

    Print #fn, "{"

    Call DumpSys(fn)
    Print #fn, ","
    Call DumpEl(fn)
    Print #fn, ","
    Call DumpKst(fn)

    Print #fn, "}"
    Close #fn

    MsgBox "Dump complete."
End Sub

Private Sub DumpSys(ByVal fn As Integer)
    Dim i As Long, j As Long, k As Long, x As Long, y As Long
    Dim firstObj As Boolean, firstInner As Boolean

    Print #fn, "  ""Sys"": ["

    firstObj = True
    For i = 0 To Maxelementezahl
        'wenn du nur “benutzte” Elemente willst, hier ggf. filtern
        If Not firstObj Then
            Print #fn, "    ,"
        End If
        firstObj = False

        With Sys(i)
            Print #fn, "    {"
            Print #fn, "      ""Index"": " & CStr(i) & ","
            Print #fn, "      ""Height"": " & JsonNumber(.Height) & ","
            Print #fn, "      ""Width"": " & JsonNumber(.Width) & ","
            Print #fn, "      ""Top"": " & JsonNumber(.Top) & ","
            Print #fn, "      ""Left"": " & JsonNumber(.left) & ","
            Print #fn, "      ""Element"": """ & JsonEscape(.Element) & ""","
            Print #fn, "      ""Tag"": """ & JsonEscape(.Tag) & ""","
            Print #fn, "      ""Rechts"": " & JsonBool(.Rechts) & ","
            Print #fn, "      ""Vollstaendig"": " & JsonBool(.Vollstaendig) & ","
            Print #fn, "      ""Zugehoerigkeit"": " & JsonNumber(.Zugehoerigkeit) & ","

            'E() als Objekt mit Index-Keys
            Print #fn, "      ""E"": {"
            firstInner = True
            For j = 1 To Eigenschaftszahl
                If Not firstInner Then Print #fn, ","
                firstInner = False
                Print #fn, "        """ & CStr(j) & """: " & JsonNumber(.E(j))
            Next j
            Print #fn, ""
            Print #fn, "      },"

            'S() als Objekt mit Index-Keys
            Print #fn, "      ""S"": {"
            firstInner = True
            For j = 1 To TextEigenschaftszahl
                If Not firstInner Then Print #fn, ","
                firstInner = False
                Print #fn, "        """ & CStr(j) & """: """ & JsonEscape(.s(j)) & """"
            Next j
            Print #fn, ""
            Print #fn, "      },"

            'B() mit negativen Indizes ? Objekt
            Print #fn, "      ""B"": {"
            firstInner = True
            For j = -TextEigenschaftszahl To Eigenschaftszahl
                If Not firstInner Then Print #fn, ","
                firstInner = False
                Print #fn, "        """ & CStr(j) & """: " & JsonBool(.b(j))
            Next j
            Print #fn, ""
            Print #fn, "      },"

            'Verb(2,4) ? Objekt von Objekten
            Print #fn, "      ""Verb"": {"
            For x = 1 To 2
                Print #fn, "        """ & CStr(x) & """: {"
                For y = 1 To 4
                    If y < 4 Then
                        Print #fn, "          """ & CStr(y) & """: " & JsonNumber(.Verb(x, y)) & ","
                    Else
                        Print #fn, "          """ & CStr(y) & """: " & JsonNumber(.Verb(x, y))
                    End If
                Next y
                If x < 2 Then
                    Print #fn, "        },"
                Else
                    Print #fn, "        }"
                End If
            Next x
            Print #fn, "      },"

            'Restliche Felder
            Print #fn, "      ""Anzverb"": " & JsonNumber(.Anzverb) & ","
            Print #fn, "      ""Lrein"": " & JsonNumber(.Lrein) & ","
            Print #fn, "      ""Lraus"": " & JsonNumber(.Lraus) & ","
            Print #fn, "      ""Furein"": " & JsonNumber(.Furein) & ","
            Print #fn, "      ""Furaus"": " & JsonNumber(.Furaus) & ","
            Print #fn, "      ""FureinSp"": " & JsonNumber(.FureinSp) & ","
            Print #fn, "      ""FurausSp"": " & JsonNumber(.FurausSp) & ","
            Print #fn, "      ""Fusteig"": " & JsonNumber(.Fusteig) & ","
            Print #fn, "      ""FusteigSp"": " & JsonNumber(.FusteigSp) & ","
            Print #fn, "      ""FusteigSpRoll"": " & JsonNumber(.FusteigSpRoll)

            Print #fn, "    }"
        End With
    Next i

    Print #fn, "  ]"
End Sub

Private Sub DumpEl(ByVal fn As Integer)
    Dim i As Long, j As Long
    Dim firstObj As Boolean, firstInner As Boolean

    Print #fn, "  ""El"": ["

    firstObj = True
    For i = -10 To Eigenschaftszahl
        If Not firstObj Then
            Print #fn, "    ,"
        End If
        firstObj = False

        With El(i)
            Print #fn, "    {"
            Print #fn, "      ""Index"": " & CStr(i) & ","
            Print #fn, "      ""Eigenschaft"": """ & JsonEscape(.Eigenschaft) & ""","
            Print #fn, "      ""Einheit"": """ & JsonEscape(.Einheit) & ""","
            Print #fn, "      ""Feldart"": """ & JsonEscape(.Feldart) & ""","
            Print #fn, "      ""Minimum"": " & JsonNumber(.Minimum) & ","
            Print #fn, "      ""Maximum"": " & JsonNumber(.Maximum) & ","

            Print #fn, "      ""Eig"": {"
            firstInner = True
            For j = 1 To 30
                If Not firstInner Then Print #fn, ","
                firstInner = False
                Print #fn, "        """ & CStr(j) & """: """ & JsonEscape(.Eig(j)) & """"
            Next j
            Print #fn, ""
            Print #fn, "      }"

            Print #fn, "    }"
        End With
    Next i

    Print #fn, "  ]"
End Sub

Private Sub DumpKst(ByVal fn As Integer)
    Dim i As Long
    Dim firstObj As Boolean

    Print #fn, "  ""Kst"": ["

    firstObj = True
    For i = 1 To 300
        If Not firstObj Then
            Print #fn, "    ,"
        End If
        firstObj = False

        With Kst(i)
            Print #fn, "    {"
            Print #fn, "      ""Index"": " & CStr(i) & ","
            Print #fn, "      ""ID"": " & JsonNumber(.ID) & ","
            Print #fn, "      ""zuEigenschaft"": " & JsonNumber(.zuEigenschaft) & ","
            Print #fn, "      ""Bezeichnung"": """ & JsonEscape(.Bezeichnung) & ""","
            Print #fn, "      ""Einstellung"": " & JsonNumber(.Einstellung)
            Print #fn, "    }"
        End With
    Next i

    Print #fn, "  ]"
End Sub

Private Function JsonEscape(ByVal s As String) As String
    Dim r As String
    r = s
    r = Replace$(r, "\", "\\")
    r = Replace$(r, """", "\""")
    r = Replace$(r, vbCrLf, "\n")
    r = Replace$(r, vbCr, "\n")
    r = Replace$(r, vbLf, "\n")
    r = Replace$(r, vbTab, "\t")
    JsonEscape = r
End Function

Private Function JsonBool(ByVal b As Boolean) As String
    If b Then
        JsonBool = "true"
    Else
        JsonBool = "false"
    End If
End Function

Private Function JsonNumber(ByVal v As Double) As String
    'Stellt sicher, dass Dezimalpunkt benutzt wird, nicht Komma
    JsonNumber = Replace$(CStr(v), ",", ".")
End Function






