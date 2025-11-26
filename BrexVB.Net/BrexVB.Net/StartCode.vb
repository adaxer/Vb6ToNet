Option Explicit On
Imports System.Diagnostics.Metrics
Imports System.Security.Principal
Imports System.Text.Json
Imports Microsoft.Data.SqlClient

Module StartCode

    'versionsverwaltungsvariablen
    Public Const Vers_B_Rex As String = "3.3.0340" '20250702  letzte umstellung
    Public Const B_Rex_Yellow As Single = &H80FFFF 'mittelgelb
    Public Const B_Rex_Green As Single = &HC0FFC0 'pastellgrün
    'Public rsLang_Res As New ADODB.Recordset
    Public SystemZeile As Integer 'und da steht er in der tabelle
    Public Systemartikel_bewertet As Boolean ' registrieren, um bei änderungen wieder zu löschen

    Public Class Styp 'zusammensetzung aus 2 datenbanken, deswegen leider redundant
        Public Artnr As String  'nur hier einstellen, dann werden gleich ein paar aufgaben mehr erledigt' Call Code1.SystemTyp_Set(ArtListe.TextMatrix(1, 4))
        Public Typenreihe As String
        'Public Netto As String 'artnr kann schon mal mehr als 6 zeichen umfassen, dann sind hier auf jeden fall nur 6 drin
        Public Name As String
        Public ACX As String
        Public Gewicht As Single
        Public Dicke As Single
        'Lagen As Single'unbenutzt
        Public MinLng As Double
        Public MinBrt As Double
        Public Kraftdehnung As Single
        Public KraftdehnungMode As Single
        '1 = sd
        '2 = fw
        '3 = k1
        '4 = selbstgewaehlt
        Public Zugtraeger As String
        'PA
        'Aramid
        'PET
        Public rho As Double
        Public Zahnabstand As String
        Public RZ As String
        'Public TGR As String 'manschettenregeln, wird derzeit nicht benutzt. gibt gruppenzugehörigkeit bei extremultus an, da erst sind die manschetten drin
        Public Verbindung01 As String
        'Public Fach5101 As String 'artikelspezifisches fach der japaner, einstweilen wird das nicht gebraucht
        'Public ProjektXPeriod_aut As Integer 'posps spinnerei, unbenutzt
        'Public ProjektXPeriod_man As Integer 'posps spinnerei, unbenutzt
        Public OberflächenbeschaffenheitTop As String
        Public OberflächenbeschaffenheitButton As String
        Public BeschichtungTop As String
        Public BeschichtungButton As String
    End Class
    Public SystemTyp As Styp


    'sprachenvariablen (deutsch, englisch, etc...)
    '    Public Sp As Double 'bedienersprache 0=deutsch, 5000=Englisch
    Public Abfragesprache As String 'also "en" oder "jp" oder "de"
    Public Datenblattsprache As String 'also "en" oder "jp" oder "de"

    'Datenbank
    'Public Artikeldaten As New ADODB.Connection
    'Public Typ As New ADODB.Recordset
    'Public Bspanlagen As New ADODB.Recordset


    Public Datenoperation_ok As Boolean
    Public DataSheetNew_ok As Boolean

    Public Class Tag
        Public AE As String 'an/aus A oder E ganz vorne
        Public ToolTipLang As Double 'beginnt mit 03
        Public Lang As String ' beginnt mit 20
    End Class

    Public Class AnlagenTemplate
        Public Bezeichnung As String
        Public Anlage As String
    End Class
    Public Templates(12) As AnlagenTemplate

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
    Public AnlageRefresh As Boolean

    Public Endlos As Boolean
    Public Vollstaendig As Boolean 'wenn alle elemente Vollstaendig sind, endlos wird extra gewertet


    Public Und As Boolean 'und oder oder bei der abfrage

    Public Class ElModel 'liest elementsteuerdaten aus datenbank ein
        Public Eigenschaft As String
        Public Einheit As String
        Public Feldart As String 'liste oder text
        Public Minimum As Single
        Public Maximum As Single
        Public Eig(30) As String
        '1 sind Muß-Eingaben
        '2 sind Kann-Eingaben
        '3 sind Ergebnisse
        '4 sind Muß-Eingaben, die aber nicht ausgefüllt werden müssen
        '5 sind spitzenlastbezogene Angaben
        '6 sind durchbiegungsbezogene angaben
        '7 frequenzberechnete angaben
        '8 auch frequenzberechnete angaben?
    End Class

    'speicherung aufbau der datenblätter
    'Public Class InhaltClass
    '    'Datenfeld As Single
    '    Public Feldname As String
    '    Public Feldinhalt As String
    '    Public Bezeichnung As String
    '    Public Tag As String
    '    Public Top As Integer
    '    Public Height As Integer
    'End Class
    'Public Inhalt(Feldinfogröße) As InhaltClass 'zum aktuel dargestellten typen die aus mvorl(,) ausgelesene vorlage
    'Public MVorl(2, 30) As String '1 enthält die bezeichnung des inhalts, 2 enthält die codierte zusammenstellung des inhalts
    'Public Const Feldinfogröße As Integer = 76
    'Public Feldinfo(Feldinfogröße) As String


    Public Class KstModel
        Public ID As Integer
        Public zuEigenschaft As Integer
        Public Bezeichnung As String
        Public Einstellung As Single

    End Class
    Public Kst(300) As KstModel

    'rund um die elementinformationen
    Public Const Eigenschaftszahl As Integer = 120
    Public El As New IndexedArray(Of ElModel)(-10, Eigenschaftszahl)
    Public Const TextEigenschaftszahl As Integer = 10

    Public Class S
        Public Height As Integer
        Public Width As Integer
        Public Top As Integer
        Public left As Integer
        'Index As Integer 'verweis auf element per zahl
        Public Element As String
        Public Tag As String
        Public Rechts As Boolean 'richtung bei förderern
        Public Vollstaendig As Boolean
        Public Zugehoerigkeit As Integer 'Huckepacks zum Träger (Elementindex)
        Public E(Eigenschaftszahl) As Double
        Public S(TextEigenschaftszahl) As String 'dient der speicherung der info-felder
        Public B As New IndexedArray(Of Boolean)(-TextEigenschaftszahl, Eigenschaftszahl)
        Public Verb(2, 4) As Double
        '1,1 verbundenes element
        '1,2 pos.der verb., wird bei jedem bildschirmaufbau neu erzeugt
        '1,3 Länge des freien Trums;
        '1,4 eigenfrequenz dieses trumstueckchens
        '2,x entsprechend mit dem zweiten element
        'alle Angaben finden sich jeweils in beiden beteiligten Elementen
        Public Anzverb As Single
        Public Lrein As Single
        Public Lraus As Double
        Public Furein As Double
        Public Furaus As Double
        Public FureinSp As Double 'bloss bei scheiben zur durchbiegung
        Public FurausSp As Double 'bloss bei scheiben zur durchbiegung
        Public Fusteig As Double
        Public FusteigSp As Double 'enthält die Spitzenlast
        Public FusteigSpRoll As Double 'enthält die Spitzenlast
    End Class
    Public Const Maxelementezahl As Integer = 150 'soviele elemente darf die anlage besitzen
    Public Maxelementindex As Integer 'höchster benutzter index, damit nicht komplett durchgezählt werden muß
    '0 leeres feld, um informationen zwischenzuspeichern, wenn das programm mit den bandeigenschaften spielt
    '1 reserviert für band, darf verändert werden, hinweis, wenn verändert
    '2 reserviert für band, darf nicht verändert werden, dient zum vergleich mit 2
    '3-9 bleiben leer für was auch immer
    'del ist zum löschen
    Public Sys As S() = New S(Maxelementezahl) {}

    Public Del As S 'nur zum löschen der anderen

    Public BarCodeFont As String

    Public Ursprung$
    Public Hilfetext$
    Public Sam(30) 'beinhaltet userdaten
    Public Const PI As Single = 3.1415926535


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

    Public Sub ReadData()
        EnsureArraysInitialized()
        Dim connectionString = "Data Source=(localdb)\mssqllocaldb;Initial Catalog=BrexAccess;Integrated Security=True"
        If WindowsIdentity.GetCurrent().Name.Equals("F01\mdehwck") Then 'ToDo...
            connectionString = "Data Source=.;Initial Catalog=BrexAccess;Integrated Security=True;TrustServerCertificate=True"
        End If
        Using connection = New SqlConnection(connectionString)
            connection.Open()
            ReadElemente(connection)
            ZweiScheiben()
            ReadKonstanten(connection)
            ReadBeispielAnlagen(connection)
            ReadArtikeldaten(connection, "900025")
        End Using
        Console.WriteLine("Hier später die Pick_It Daten Laden")
        Sys(1).S(1) = "Typ('typ')"
        Sys(1).S(3) = "Typ('beschtsfs')" & " " & "Typ('oberfltsfs')"
        Sys(1).S(4) = "Typ('beschlsas')" & " " & "Typ('oberfllsas')"
    End Sub

    Private Sub ReadArtikeldaten(connection As SqlConnection, artikelNr As String)
        Dim command = New SqlCommand("SELECT * from Artikel where  ArtNr like '" & artikelNr & "'", connection) ' "SELECT * FROM art_de_en WHERE (hersteller = 'damussnureinANDrein' or (hersteller like '%siegling%' and (typenreihe = 'transilon' or typenreihe = 'transvent'))  or (hersteller like '%siegling%' and (typenreihe = 'extremultus'))  ) ORDER BY artnr"

        Using reader = command.ExecuteReader()
            reader.Read()
            SystemTyp.Artnr = reader("Artnr")
            SystemTyp.Typenreihe = reader("Typenreihe")
            SystemTyp.Gewicht = reader("Gewicht")
            SystemTyp.Dicke = reader("Dicke")
            SystemTyp.Kraftdehnung = 0.0 'Wird im laufe des Programms berechnet oder gesetzt
            SystemTyp.KraftdehnungMode = 0.0 'Wird im laufe des Programms berechnet oder gesetzt
            SystemTyp.Name = "E  8/2 U0/V5 grün -> Value nur zum testen"

            'Mal schauen ob wir die brauchen:
            'SystemTyp.ACX = reader("ACX")
            'SystemTyp.MinLng = reader("MinLng")
            'SystemTyp.MinBrt = reader("MinBrt")
            'SystemTyp.Zugtraeger = reader("Zugtraeger")
            'SystemTyp.rho = reader("rho")
            'SystemTyp.Zahnabstand = reader("Zahnabstand")
            'SystemTyp.RZ = reader("RZ")
            'SystemTyp.Verbindung01 = reader("Verbindung01")

            SystemTyp.BeschichtungTop = "Polyvinylchlorid (0.5 mm) -> Value nur zum testen" 'reader("beschtsfs")
            SystemTyp.BeschichtungButton = "Polyurethan-Imprägnierung -> Value nur zum testen"  'reader("beschlsas")
            SystemTyp.OberflächenbeschaffenheitTop = "Glatt (0.5 mm) -> Value nur zum testen"  'reader("oberfltsfs")
            SystemTyp.OberflächenbeschaffenheitButton = "Gewebe -> Value nur zum testen"  'reader("oberfllsas")

            Sys(1).S(1) = SystemTyp.Name
            Sys(1).S(3) = SystemTyp.BeschichtungTop & " " & SystemTyp.OberflächenbeschaffenheitTop
            Sys(1).S(4) = SystemTyp.BeschichtungButton & " " & SystemTyp.OberflächenbeschaffenheitButton

        End Using
    End Sub

    Private Sub ReadKonstanten(connection As SqlConnection)
        Dim command = New SqlCommand("SELECT * from Konstanten ORDER BY nummer", connection)
        Using reader = command.ExecuteReader()
            Dim i As Integer
            While reader.Read()
                i = reader("nummer")
                Kst(i).ID = i
                Kst(i).zuEigenschaft = reader("zu_Eigenschaft")
                Kst(i).Bezeichnung = reader("bezeichnung")
                Kst(i).Einstellung = reader("einstellung")
            End While
        End Using
    End Sub

    Private Sub ReadBeispielAnlagen(connection As SqlConnection)
        Dim command = New SqlCommand("SELECT * from Beispielanlagen ORDER BY Bezeichnung", connection)
        Using reader = command.ExecuteReader()
            Dim index As Integer = 1
            While reader.Read()
                Templates(index).Bezeichnung = reader("Bezeichnung")
                Templates(index).Anlage = reader("Anlage")
                index += 1
            End While
        End Using
    End Sub

    Public Sub ZweiScheiben()
        ' todo!
    End Sub


    Private Sub ReadElemente(connection As SqlConnection)
        Dim command = New SqlCommand("SELECT * from Elemente_a", connection)
        Using reader = command.ExecuteReader()
            Dim i As Integer
            While reader.Read()
                i = reader("nummer")
                If IsDefined(reader("eigenschaft")) Then El(i).Eigenschaft = reader("eigenschaft")
                If IsDefined(reader("einheit")) Then El(i).Einheit = reader("einheit")
                If IsDefined(reader("Feldart")) Then El(i).Feldart = reader("Feldart")
                If IsDefined(reader("minimum")) Then El(i).Minimum = reader("minimum")
                If IsDefined(reader("maximum")) Then El(i).Maximum = reader("maximum")
                For i = 6 To 27
                    If IsDefined(reader(i)) Then El(reader("nummer")).Eig(i) = reader(i)
                Next i
            End While
        End Using
    End Sub

    Private Sub EnsureArraysInitialized()
        If El(1) IsNot Nothing Then Exit Sub
        For i = -10 To Eigenschaftszahl
            El(i) = New ElModel()
        Next
        For i = 1 To 300
            Kst(i) = New KstModel()
        Next
        For i = 1 To Maxelementezahl
            Sys(i) = New S()
        Next
        For i = 1 To 12
            Templates(i) = New AnlagenTemplate()
        Next
        SystemTyp = New Styp()
    End Sub

    Public Function IsDefined(value As Object) As Boolean
        If value Is Nothing Then Return False
        If value Is DBNull.Value Then Return False
        If TypeOf value Is String Then
            Return Not String.IsNullOrEmpty(CStr(value))
        End If
        Return True
    End Function

End Module
