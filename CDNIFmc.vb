Public Class CDNI
    'creado por Fernando Mendiz
    '29/11/2012 
    '12:09
    'v1.00
    'cambios : inicial
    'Para usar 
    'utilizar la funcion esvalido 
    Private _letra As String
    Private _dni As Integer



    Sub New(ByVal dni As Integer)
        _dni = dni
        Calculaletra()
    End Sub
    Sub New()

    End Sub

    Public Function DameLetra() As String
        Return _letra
    End Function
    Function esValido(ByVal DNI As String) as boolean
        If DNI.Length = 9 Then
            If Char.IsLetter(DNI.Chars(8)) Then 'si la ULTIMA posicion es letra 
                If IsNumeric(DNI.Substring(0, 8)) Then ' SI el RESTO ES NUMERO
                    Calculaletra(CInt(DNI.Substring(0, 8)))
                    If DameLetra() = CStr(DNI.Chars(8)) Then
                        Return True
                    End If
                End If
            End If
        End If
        Return False
    End Function
    Public Sub Calculaletra()
        Dim Resto As Integer
        Resto = CInt(_dni) Mod 23
        Select Case (Resto)
            Case 0
                _letra = "T"
            Case 1
                _letra = "R"
            Case 2
                _letra = "W"
            Case 3
                _letra = "A"
            Case 4
                _letra = "G"
            Case 5
                _letra = "M"
            Case 6
                _letra = "Y"
            Case 7
                _letra = "F"
            Case 8
                _letra = "P"
            Case 9
                _letra = "D"
            Case 10
                _letra = "X"
            Case 11
                _letra = "B"
            Case 12
                _letra = "N"
            Case 13
                _letra = "J"
            Case 14
                _letra = "Z"
            Case 15
                _letra = "S"
            Case 16
                _letra = "Q"
            Case 17
                _letra = "V"
            Case 18
                _letra = "H"
            Case 19
                _letra = "L"
            Case 20
                _letra = "C"
            Case 21
                _letra = "K"
            Case 22
                _letra = "E"
        End Select
    End Sub

    Public Sub Calculaletra(ByVal dni As Integer)
        Dim Resto As Integer
        Resto = CInt(dni) Mod 23
        Select Case (Resto)
            Case 0
                _letra = "T"
            Case 1
                _letra = "R"
            Case 2
                _letra = "W"
            Case 3
                _letra = "A"
            Case 4
                _letra = "G"
            Case 5
                _letra = "M"
            Case 6
                _letra = "Y"
            Case 7
                _letra = "F"
            Case 8
                _letra = "P"
            Case 9
                _letra = "D"
            Case 10
                _letra = "X"
            Case 11
                _letra = "B"
            Case 12
                _letra = "N"
            Case 13
                _letra = "J"
            Case 14
                _letra = "Z"
            Case 15
                _letra = "S"
            Case 16
                _letra = "Q"
            Case 17
                _letra = "V"
            Case 18
                _letra = "H"
            Case 19
                _letra = "L"
            Case 20
                _letra = "C"
            Case 21
                _letra = "K"
            Case 22
                _letra = "E"
        End Select
    End Sub

End Class
