Attribute VB_Name = "Convertir"
Public Function De_Txt_a_Num_01(ByVal sTexto As String, _
                                   Optional ByVal nDecimales As Integer = 3, _
                                   Optional ByVal sP_Formato_Decimal As String = "") As Double
    '-------------------------------------------------§§§----'
    ' FUNCION PARA CONVERTIR UN TEXTO EN NUMERO DECIMAL
    '-------------------------------------------------§§§----'
    '
    Dim bCte2 As Boolean
    '
    Dim nContador1 As Integer
    Dim nContador2 As Integer
    Dim nLong_Total As Integer
    Dim nPos_Punto As Integer
    Dim nCte1 As Integer
    Dim nDecimal As Integer
    '
    Dim lNumeruco As Double
    '
    Dim sNumero As String
    Dim sL_Aux_01 As String
    '
    Dim sL_Array_Pto_01() As String
    Dim sL_Array_Coma_01() As String
    '
    On Error GoTo Error_Numero
    '
    '-------------------------------------------------§§§----'
    Select Case sP_Formato_Decimal
        Case "."    ' USAMOS "." COMO SEPARADOR DE DECIMALES
                    ' Y LA "," LA ELIMINAMOS
            sL_Array_Pto_01 = Split(sTexto, ".")
            sL_Array_Coma_01 = Split(sTexto, ",")
            '
            sL_Aux_01 = ""
            For nContador1 = LBound(sL_Array_Coma_01) To UBound(sL_Array_Coma_01)
                sL_Aux_01 = sL_Aux_01 & sL_Array_Coma_01(nContador1)
                ''
            Next nContador1
            '
            sTexto = sL_Aux_01
            ''
        Case ","    ' USAMOS "," COMO SEPARADOR DE DECIMALES
                    ' Y EL "." LE ELIMINAMOS
            sL_Array_Pto_01 = Split(sTexto, ".")
            sL_Array_Coma_01 = Split(sTexto, ",")
            '
            sL_Aux_01 = ""
            For nContador1 = LBound(sL_Array_Pto_01) To UBound(sL_Array_Pto_01)
                sL_Aux_01 = sL_Aux_01 & sL_Array_Pto_01(nContador1)
                ''
            Next nContador1
            '
            sTexto = sL_Aux_01
            ''
    End Select
    '-------------------------------------------------§§§----'
    '
    lNumeruco = 0
    '
    If nDecimales >= 0 Then
        nDecimal = nDecimales
        ''
    Else
        nDecimal = 3
        ''
    End If
    '
    sTexto = Trim(sTexto)
    '
    If InStr(1, sTexto, "-") > 0 Then
        'Es un numero negativo
        bCte2 = True
        sTexto = Mid$(sTexto, 2)
        ''
    ElseIf InStr(1, sTexto, "+") > 0 Then
        'Es un numero positivo (con signo)
        bCte2 = False
        sTexto = Mid$(sTexto, 2)
        ''
    Else
        'Es un numero positivo
        bCte2 = False
        ''
    End If
    '
    nLong_Total = Len(sTexto)
    '
    For nContador1 = 1 To nLong_Total
        If Mid(sTexto, nContador1, 1) = "," Then Mid(sTexto, nContador1, 1) = "."
        ''
    Next nContador1
    '
    If InStr(1, sTexto, ".") <= 0 Then sTexto = sTexto & ".0"
    '
    nPos_Punto = InStr(1, sTexto, ".")
    '
    nContador2 = 0
    For nContador1 = 1 To nLong_Total
        If Mid$(sTexto, nContador1, 1) <> "." Then
            'No estamos en el caracte "."
            If nContador1 < nPos_Punto And nPos_Punto <> 0 Then
                nCte1 = 1
                ''
            Else
                nContador2 = nContador2 + 1
                nCte1 = 0
                ''
            End If
            '
            sNumero = Mid$(sTexto, nContador1, 1)
            '
            If nContador2 > nDecimal Then
                If sNumero > 5 Then lNumeruco = lNumeruco + (CSng(1) * (10 ^ (nPos_Punto - nContador1 - nCte1 + 1)))
                nContador1 = nLong_Total
                ''
            Else
                lNumeruco = lNumeruco + (CSng(sNumero) * (10 ^ (nPos_Punto - nContador1 - nCte1)))
                ''
            End If
            ''
        End If
        ''
    Next nContador1
    '
    If bCte2 = True Then
        De_Txt_a_Num_01 = (-1) * lNumeruco
        ''
    Else
        De_Txt_a_Num_01 = (1) * lNumeruco
        ''
    End If
    '
    If (nDecimales >= 0) Then De_Txt_a_Num_01 = Round(De_Txt_a_Num_01, nDecimales)
    '
Exit Function
'
Error_Numero:
    '
    '-------------------------------------------------§§§----'
    ' ERROR DE NUMERO
    '-------------------------------------------------§§§----'
    '
    De_Txt_a_Num_01 = -1.75E+308
    ''
End Function

Public Function De_Num_a_Tx_01(ByVal lNumero As Double, _
                               Optional ByVal bEntero As Boolean = False, _
                               Optional ByVal nDecimales As Integer = 3) As String
    '-------------------------------------------------§§§----'
    ' FUNCION PARA CONVERTIR UN NUMERO EN TEXTO
    '-------------------------------------------------§§§----'
    '
    On Error GoTo Fin
    '
    Dim sNumero As String
    Dim nLong1 As Integer
    Dim nCont1 As Integer
    '
    If bEntero = True Then
        sNumero = CStr(Format(lNumero, "########0"))
        ''
    Else
        Select Case nDecimales
            Case -1: sNumero = CStr(Format(lNumero, "########0.#########"))
            Case 1: sNumero = CStr(Format(lNumero, "########0.#"))
            Case 2: sNumero = CStr(Format(lNumero, "########0.0#"))
            Case 3: sNumero = CStr(Format(lNumero, "########0.00#"))
            Case 4: sNumero = CStr(Format(lNumero, "########0.000#"))
            Case 5: sNumero = CStr(Format(lNumero, "########0.0000#"))
            Case 6: sNumero = CStr(Format(lNumero, "########0.00000#"))
            Case 7: sNumero = CStr(Format(lNumero, "########0.000000#"))
            Case 8: sNumero = CStr(Format(lNumero, "########0.0000000#"))
            Case 9: sNumero = CStr(Format(lNumero, "########0.00000000#"))
            Case 9: sNumero = CStr(Format(lNumero, "########0.00000000#"))
            Case 10: sNumero = CStr(Format(lNumero, "########0.000000000#"))
            Case 11: sNumero = CStr(Format(lNumero, "########0.0000000000#"))
            Case 12: sNumero = CStr(Format(lNumero, "########0.00000000000#"))
            Case Else: sNumero = CStr(Format(lNumero, "########0.00#"))
        End Select
        ''
    End If
    '
    nLong1 = Len(sNumero)
    '
    For nCont1 = 1 To nLong1
        If Mid$(sNumero, nCont1, 1) = "," Then Mid(sNumero, nCont1, 1) = "."
        ''
    Next nCont1
    '
    If bEntero = True Then
        De_Num_a_Tx_01 = sNumero
        ''
    ElseIf InStr(sNumero, ".") > 0 Then
        If (Len(sNumero) = InStr(sNumero, ".")) And (nDecimales = -1) Then
            De_Num_a_Tx_01 = Mid$(sNumero, 1, InStr(sNumero, ".") - 1)
            ''
        Else
            De_Num_a_Tx_01 = sNumero
            ''
        End If
        ''
    Else
        De_Num_a_Tx_01 = sNumero & ".0"
        ''
    End If
    '
Exit Function
'
Fin:
    De_Num_a_Tx_01 = "###.###"
    ''
End Function



