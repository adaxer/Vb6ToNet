Imports System.IO
Imports System.Text
Imports System.Text.Json

Module InputOutput
    Public Class BrexDumpDto
        Public Property Sys As SDump()
        Public Property El As ElDump()
        Public Property Kst As KstDump()
    End Class

    Public Class SDump
        Public Property Index As Integer
        Public Property Height As Integer
        Public Property Width As Integer
        Public Property Top As Integer
        Public Property Left As Integer
        Public Property Element As String
        Public Property Tag As String
        Public Property Rechts As Boolean
        Public Property Vollstaendig As Boolean
        Public Property Zugehoerigkeit As Integer

        ' hier kommen die früher als JSON-Objekte ausgegebenen Arrays:
        Public Property E As Dictionary(Of Integer, Double)
        Public Property S As Dictionary(Of Integer, String)
        Public Property B As Dictionary(Of Integer, Boolean)
        Public Property Verb As Dictionary(Of Integer, Dictionary(Of Integer, Double))

        Public Property Anzverb As Double
        Public Property Lrein As Double
        Public Property Lraus As Double
        Public Property Furein As Double
        Public Property Furaus As Double
        Public Property FureinSp As Double
        Public Property FurausSp As Double
        Public Property Fusteig As Double
        Public Property FusteigSp As Double
        Public Property FusteigSpRoll As Double
    End Class

    Public Class ElDump
        Public Property Index As Integer
        Public Property Eigenschaft As String
        Public Property Einheit As String
        Public Property Feldart As String
        Public Property Minimum As Double
        Public Property Maximum As Double
        Public Property Eig As Dictionary(Of Integer, String)
    End Class

    Public Class KstDump
        Public Property Index As Integer
        Public Property ID As Integer
        Public Property zuEigenschaft As Integer
        Public Property Bezeichnung As String
        Public Property Einstellung As Double
    End Class

    Public Function ReadBrexDump(filePath As String) As BrexParameters
        Dim options = New JsonSerializerOptions With {.PropertyNameCaseInsensitive = True}
        Dim enc = Encoding.GetEncoding(Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ANSICodePage)
        Dim dump = JsonSerializer.Deserialize(Of BrexDumpDto)(File.ReadAllText(filePath, enc), options)
        Return ApplyDumpToParameters(dump)
    End Function

    Public Function ApplyDumpToParameters(dump As BrexDumpDto) As BrexParameters
        Const ElOffset As Integer = 10

        Dim result As New BrexParameters With {
        .Sys = New S(Maxelementezahl) {},
        .Kst = New KstModel(300) {},
        .El = New ElModel(Eigenschaftszahl + ElOffset) {}
    }

        ' Instanzen anlegen
        Dim i As Integer
        For i = 0 To Maxelementezahl
            result.Sys(i) = New S()
        Next

        For i = 0 To 300
            result.Kst(i) = New KstModel()
        Next

        For i = 0 To Eigenschaftszahl + ElOffset
            result.El(i) = New ElModel()
        Next

        ' ===== Sys =====
        If dump.Sys IsNot Nothing Then
            For Each sDump In dump.Sys
                If sDump Is Nothing Then Continue For
                If sDump.Index < 0 OrElse sDump.Index > Maxelementezahl Then Continue For

                Dim s = result.Sys(sDump.Index)

                s.Height = sDump.Height
                s.Width = sDump.Width
                s.Top = sDump.Top
                s.left = sDump.Left
                s.Element = sDump.Element
                s.Tag = sDump.Tag
                s.Rechts = sDump.Rechts
                s.Vollstaendig = sDump.Vollstaendig
                s.Zugehoerigkeit = sDump.Zugehoerigkeit

                ' E()
                If sDump.E IsNot Nothing Then
                    For Each kv In sDump.E
                        If kv.Key >= 0 AndAlso kv.Key <= Eigenschaftszahl Then
                            s.E(kv.Key) = kv.Value
                        End If
                    Next
                End If

                ' S()
                If sDump.S IsNot Nothing Then
                    For Each kv In sDump.S
                        Dim idx = Math.Abs(kv.Key)
                        If idx >= 0 AndAlso idx <= TextEigenschaftszahl Then
                            s.S(idx) = kv.Value
                        End If
                    Next
                End If

                ' B() – IndexedArray, Indizes sind wie in VB6
                If sDump.B IsNot Nothing Then
                    For Each kv In sDump.B
                        ' Wir vertrauen auf die Bounds-Prüfung in IndexedArray
                        s.B(kv.Key) = kv.Value
                    Next
                End If

                ' Verb(2,4)
                If sDump.Verb IsNot Nothing Then
                    For Each x In sDump.Verb.Keys
                        If x < 1 OrElse x > 2 Then Continue For
                        Dim inner = sDump.Verb(x)
                        If inner Is Nothing Then Continue For
                        For Each y In inner.Keys
                            If y < 1 OrElse y > 4 Then Continue For
                            s.Verb(x, y) = inner(y)
                        Next
                    Next
                End If

                s.Anzverb = CSng(sDump.Anzverb)
                s.Lrein = CSng(sDump.Lrein)
                s.Lraus = sDump.Lraus
                s.Furein = sDump.Furein
                s.Furaus = sDump.Furaus
                s.FureinSp = sDump.FureinSp
                s.FurausSp = sDump.FurausSp
                s.Fusteig = sDump.Fusteig
                s.FusteigSp = sDump.FusteigSp
                s.FusteigSpRoll = sDump.FusteigSpRoll
            Next
        End If

        ' ===== El =====
        If dump.El IsNot Nothing Then
            For Each eDump In dump.El
                If eDump Is Nothing Then Continue For
                If eDump.Index < -10 OrElse eDump.Index > Eigenschaftszahl Then Continue For

                Dim arrIndex = eDump.Index + ElOffset
                If arrIndex < 0 OrElse arrIndex > result.El.Length - 1 Then Continue For

                Dim elModel = result.El(arrIndex)

                elModel.Eigenschaft = eDump.Eigenschaft
                elModel.Einheit = eDump.Einheit
                elModel.Feldart = eDump.Feldart
                elModel.Minimum = CSng(eDump.Minimum)
                elModel.Maximum = CSng(eDump.Maximum)

                If eDump.Eig IsNot Nothing Then
                    For Each kv In eDump.Eig
                        If kv.Key >= 0 AndAlso kv.Key <= 30 Then
                            elModel.Eig(kv.Key) = kv.Value
                        End If
                    Next
                End If
            Next
        End If

        ' ===== Kst =====
        If dump.Kst IsNot Nothing Then
            For Each kDump In dump.Kst
                If kDump Is Nothing Then Continue For
                If kDump.Index < 0 OrElse kDump.Index > 300 Then Continue For

                Dim k = result.Kst(kDump.Index)
                k.ID = kDump.ID
                k.zuEigenschaft = kDump.zuEigenschaft
                k.Bezeichnung = kDump.Bezeichnung
                k.Einstellung = CSng(kDump.Einstellung)
            Next
        End If

        Return result
    End Function

    ' Vergleicht globale Arrays (Sys, El, Kst) mit einem BrexParameters-Snapshot.
    ' Nur Unterschiede werden nach filePath geschrieben.
    Public Sub CompareWithParameters(params As BrexParameters, filePath As String)
        Using w As New StreamWriter(filePath, append:=False)
            CompareSysArrays(w, params)
            CompareElArrays(w, params)
            CompareKstArrays(w, params)
        End Using
    End Sub

    ' ===== Helper: Toleranzen =====

    Private Function NearlyEqual(a As Double, b As Double, Optional relTol As Double = 0.001) As Boolean
        Dim da = Math.Abs(a)
        Dim db = Math.Abs(b)
        Dim denom = Math.Max(da, db)
        If denom = 0 Then Return True ' beide 0
        Return Math.Abs(a - b) / denom <= relTol
    End Function

    Private Function NearlyEqualSingle(a As Single, b As Single, Optional relTol As Double = 0.001) As Boolean
        Return NearlyEqual(CDbl(a), CDbl(b), relTol)
    End Function

    ' ===== Sys-Vergleich =====

    Private Sub CompareSysArrays(w As StreamWriter, params As BrexParameters)
        If params.Sys Is Nothing Then
            w.WriteLine("Sys: parameters.Sys is Nothing")
            Return
        End If

        Dim maxIdx = Math.Min(Sys.Length - 1, params.Sys.Length - 1)

        For i = 0 To maxIdx
            Dim live = Sys(i)
            Dim snap = params.Sys(i)
            If live Is Nothing AndAlso snap Is Nothing Then Continue For
            If live Is Nothing Xor snap Is Nothing Then
                w.WriteLine($"Sys[{i}]: live IsNothing={live Is Nothing}, snapshot IsNothing={snap Is Nothing}")
                Continue For
            End If

            ' Skalare
            CompareInt(w, live.Height, snap.Height, $"Sys[{i}].Height")
            CompareInt(w, live.Width, snap.Width, $"Sys[{i}].Width")
            CompareInt(w, live.Top, snap.Top, $"Sys[{i}].Top")
            CompareInt(w, live.left, snap.left, $"Sys[{i}].Left")
            CompareString(w, live.Element, snap.Element, $"Sys[{i}].Element")
            CompareString(w, live.Tag, snap.Tag, $"Sys[{i}].Tag")
            CompareBool(w, live.Rechts, snap.Rechts, $"Sys[{i}].Rechts")
            CompareBool(w, live.Vollstaendig, snap.Vollstaendig, $"Sys[{i}].Vollstaendig")
            CompareInt(w, live.Zugehoerigkeit, snap.Zugehoerigkeit, $"Sys[{i}].Zugehoerigkeit")

            CompareSingle(w, live.Anzverb, snap.Anzverb, $"Sys[{i}].Anzverb")
            CompareSingle(w, live.Lrein, snap.Lrein, $"Sys[{i}].Lrein")
            CompareDouble(w, live.Lraus, snap.Lraus, $"Sys[{i}].Lraus")
            CompareDouble(w, live.Furein, snap.Furein, $"Sys[{i}].Furein")
            CompareDouble(w, live.Furaus, snap.Furaus, $"Sys[{i}].Furaus")
            CompareDouble(w, live.FureinSp, snap.FureinSp, $"Sys[{i}].FureinSp")
            CompareDouble(w, live.FurausSp, snap.FurausSp, $"Sys[{i}].FurausSp")
            CompareDouble(w, live.Fusteig, snap.Fusteig, $"Sys[{i}].Fusteig")
            CompareDouble(w, live.FusteigSp, snap.FusteigSp, $"Sys[{i}].FusteigSp")
            CompareDouble(w, live.FusteigSpRoll, snap.FusteigSpRoll, $"Sys[{i}].FusteigSpRoll")

            ' E()
            If live.E IsNot Nothing AndAlso snap.E IsNot Nothing Then
                Dim lenE = Math.Min(live.E.Length - 1, snap.E.Length - 1)
                For j = 0 To lenE
                    Dim path = $"Sys[{i}].E[{j}]"
                    If Not NearlyEqual(live.E(j), snap.E(j)) Then
                        w.WriteLine($"{path}: live={live.E(j)} snapshot={snap.E(j)}")
                    End If
                Next
            End If

            ' S()
            If live.S IsNot Nothing AndAlso snap.S IsNot Nothing Then
                Dim lenS = Math.Min(live.S.Length - 1, snap.S.Length - 1)
                For j = 0 To lenS
                    CompareString(w, live.S(j), snap.S(j), $"Sys[{i}].S[{j}]")
                Next
            End If

            ' B: IndexedArray(-TextEigenschaftszahl..Eigenschaftszahl)
            If live.B IsNot Nothing AndAlso snap.B IsNot Nothing Then
                Dim low = Math.Max(live.B.LowerBound, snap.B.LowerBound)
                Dim up = Math.Min(live.B.UpperBound, snap.B.UpperBound)
                For idx = low To up
                    Dim path = $"Sys[{i}].B[{idx}]"
                    If live.B(idx) <> snap.B(idx) Then
                        w.WriteLine($"{path}: live={live.B(idx)} snapshot={snap.B(idx)}")
                    End If
                Next
            End If

            ' Verb(2,4)
            If live.Verb IsNot Nothing AndAlso snap.Verb IsNot Nothing Then
                For x = 1 To 2
                    For y = 1 To 4
                        Dim path = $"Sys[{i}].Verb[{x},{y}]"
                        If Not NearlyEqual(live.Verb(x, y), snap.Verb(x, y)) Then
                            w.WriteLine($"{path}: live={live.Verb(x, y)} snapshot={snap.Verb(x, y)}")
                        End If
                    Next
                Next
            End If
        Next
    End Sub

    ' ===== El-Vergleich =====
    ' Mapping: global El(-10..Eigenschaftszahl) <-> params.El(index + 10)

    Private Sub CompareElArrays(w As StreamWriter, params As BrexParameters)
        If params.El Is Nothing Then
            w.WriteLine("El: parameters.El is Nothing")
            Return
        End If

        Const ElOffset As Integer = 10

        For origIndex = -10 To Eigenschaftszahl
            Dim snapIndex = origIndex + ElOffset
            If snapIndex < 0 OrElse snapIndex >= params.El.Length Then Continue For

            Dim live = El(origIndex)
            Dim snap = params.El(snapIndex)

            If live Is Nothing AndAlso snap Is Nothing Then Continue For
            If live Is Nothing Xor snap Is Nothing Then
                w.WriteLine($"El[{origIndex}]: live IsNothing={live Is Nothing}, snapshot IsNothing={snap Is Nothing}")
                Continue For
            End If

            Dim basePath = $"El[{origIndex}]"

            CompareString(w, live.Eigenschaft, snap.Eigenschaft, basePath & ".Eigenschaft")
            CompareString(w, live.Einheit, snap.Einheit, basePath & ".Einheit")
            CompareString(w, live.Feldart, snap.Feldart, basePath & ".Feldart")
            CompareSingle(w, live.Minimum, snap.Minimum, basePath & ".Minimum")
            CompareSingle(w, live.Maximum, snap.Maximum, basePath & ".Maximum")

            If live.Eig IsNot Nothing AndAlso snap.Eig IsNot Nothing Then
                Dim lenEig = Math.Min(live.Eig.Length - 1, snap.Eig.Length - 1)
                For j = 0 To lenEig
                    CompareString(w, live.Eig(j), snap.Eig(j), $"{basePath}.Eig[{j}]")
                Next
            End If
        Next
    End Sub

    ' ===== Kst-Vergleich =====

    Private Sub CompareKstArrays(w As StreamWriter, params As BrexParameters)
        If params.Kst Is Nothing Then
            w.WriteLine("Kst: parameters.Kst is Nothing")
            Return
        End If

        Dim maxIdx = Math.Min(Kst.Length - 1, params.Kst.Length - 1)

        For i = 0 To maxIdx
            Dim live = Kst(i)
            Dim snap = params.Kst(i)

            If live Is Nothing AndAlso snap Is Nothing Then Continue For
            If live Is Nothing Xor snap Is Nothing Then
                w.WriteLine($"Kst[{i}]: live IsNothing={live Is Nothing}, snapshot IsNothing={snap Is Nothing}")
                Continue For
            End If

            Dim basePath = $"Kst[{i}]"
            CompareInt(w, live.ID, snap.ID, basePath & ".ID")
            CompareInt(w, live.zuEigenschaft, snap.zuEigenschaft, basePath & ".zuEigenschaft")
            CompareString(w, live.Bezeichnung, snap.Bezeichnung, basePath & ".Bezeichnung")
            CompareSingle(w, live.Einstellung, snap.Einstellung, basePath & ".Einstellung")
        Next
    End Sub

    ' ===== primitive Compare-Helper =====

    Private Sub CompareInt(w As StreamWriter, live As Integer, snap As Integer, path As String)
        If live <> snap Then
            w.WriteLine($"{path}: live={live} snapshot={snap}")
        End If
    End Sub

    Private Sub CompareDouble(w As StreamWriter, live As Double, snap As Double, path As String)
        If Not NearlyEqual(live, snap) Then
            w.WriteLine($"{path}: live={live} snapshot={snap}")
        End If
    End Sub

    Private Sub CompareSingle(w As StreamWriter, live As Single, snap As Single, path As String)
        If Not NearlyEqualSingle(live, snap) Then
            w.WriteLine($"{path}: live={live} snapshot={snap}")
        End If
    End Sub

    Private Sub CompareBool(w As StreamWriter, live As Boolean, snap As Boolean, path As String)
        If live <> snap Then
            w.WriteLine($"{path}: live={live} snapshot={snap}")
        End If
    End Sub

    Private Sub CompareString(w As StreamWriter, live As String, snap As String, path As String)
        If Not String.Equals(live, snap, StringComparison.Ordinal) And Not (String.IsNullOrWhiteSpace(live) And String.IsNullOrWhiteSpace(snap)) Then
            ' optional: Leer/Null und "nur Spaces" gleich behandeln:
            'If String.IsNullOrWhiteSpace(live) AndAlso String.IsNullOrWhiteSpace(snap) Then Return
            w.WriteLine($"{path}: live=""{live}"" snapshot=""{snap}""")
        End If
    End Sub


End Module
