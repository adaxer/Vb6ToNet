Module Code1

    Public Function CDBLVAL(value As String) As Double
        If value = "" Then
            CDBLVAL = 0
        Else
            value = Replace(value, ",", ".")
            CDBLVAL = Val(value)
            'cdbl schlägt aus unerfindlichen gründen nicht selten fehl
        End If
    End Function

    Public Function Elementnummer(value As String) As Integer
        Elementnummer = 5 'neuverwendung m
        Do 'ab 5 beginnen erst die elemente
            Elementnummer = Elementnummer + 1
        Loop Until El(0).Eig(Elementnummer) = value '0 enthält die elementnamen
    End Function


End Module
