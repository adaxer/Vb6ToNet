Option Strict On
Option Explicit On
Public Module VBCompat
    ''' <summary>
    ''' Provides an array-like container with arbitrary lower and upper integer bounds and a default indexer.
    ''' This allows existing VB6-style access patterns like <c>El(-2)</c> to be preserved after migration.
    ''' </summary>
    Public Class IndexedArray(Of T)
        Private ReadOnly _lower As Integer
        Private ReadOnly _upper As Integer
        Private ReadOnly _data() As T

        Public Sub New(lowerBound As Integer, upperBound As Integer)
            If upperBound < lowerBound Then
                Throw New ArgumentException("upperBound must be >= lowerBound", NameOf(upperBound))
            End If
            _lower = lowerBound
            _upper = upperBound
            ' create storage for (upper - lower + 1) elements
            ReDim _data(_upper - _lower)
        End Sub

        Default Public Property Item(index As Integer) As T
            Get
                Dim idx = index - _lower
                If idx < 0 OrElse idx >= _data.Length Then
                    Throw New IndexOutOfRangeException($"Index {index} out of bounds ({_lower}..{_upper})")
                End If
                Return _data(idx)
            End Get
            Set(value As T)
                Dim idx = index - _lower
                If idx < 0 OrElse idx >= _data.Length Then
                    Throw New IndexOutOfRangeException($"Index {index} out of bounds ({_lower}..{_upper})")
                End If
                _data(idx) = value
            End Set
        End Property

        Public ReadOnly Property LowerBound As Integer
            Get
                Return _lower
            End Get
        End Property

        Public ReadOnly Property UpperBound As Integer
            Get
                Return _upper
            End Get
        End Property

        Public ReadOnly Property Length As Integer
            Get
                Return _data.Length
            End Get
        End Property
    End Class

    ''' <summary>
    ''' VB6-kompatible InStr-Implementierung (1-basierte Rückgabe, 0 = nicht gefunden).
    ''' Start ist 1-basiert wie in VB6. Vergleichsart kann per StringComparison gesteuert werden.
    ''' </summary>
    Public Function InStr(start As Integer, source As String, value As String, Optional comparison As StringComparison = StringComparison.Ordinal) As Integer
        If String.IsNullOrEmpty(source) OrElse String.IsNullOrEmpty(value) Then
            Return 0
        End If

        ' VB6 beginnt bei 1 — clamp auf mindestens 1
        Dim s As Integer = Math.Max(1, start)

        ' .NET IndexOf ist 0-basiert, daher s-1 übergeben
        Dim idx As Integer = source.IndexOf(value, s - 1, comparison)

        If idx >= 0 Then
            Return idx + 1 ' zurück zu 1-basiert
        Else
            Return 0
        End If
    End Function

    Public Function CDBLVAL(a$) As Double
        If a$ = "" Then
            CDBLVAL = 0
        Else
            a$ = Replace(a$, ",", ".")
            CDBLVAL = Val(a$)
            'cdbl schlägt aus unerfindlichen gründen nicht selten fehl
        End If
    End Function
End Module
