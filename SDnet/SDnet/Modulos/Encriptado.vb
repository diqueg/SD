Module Encriptado

    'NOTA DE MIGRACION A VB.NET: TOdas estas funciones en la versión VB6 estaban definidas para trabajar con la matriz PLAYFAIR_MATRIX
    ' de dos dimensiones 4x4 pero con base = 1. En VB.net la base standard es base 0 y es la única que se admite, de manera que
    ' tuve que modificar todo el código para que trabaje en base cero. No solo se cambió el rango de dimensinoes de la matiz ("1 to 4" por "4") sino también
    ' varios FOR NEXT donde hubo que cambiar los limites "entre 1 y 4" por "entre 0 y 3". Si cometí un algún error en esto estas funcones van a fallar.
    'Atento con esto!!.


    Dim KEY As String
    Dim ROTH As Boolean
    Dim ROTL As Boolean
    Dim XCHG As Boolean
    Dim XCHG_ROT As Boolean
    Dim MATRIX As Integer

    '* GENERADOR DE CLAVE ALEATORIA *
    '********************************

    Public Function Endata_encriptar_genKEY() As String
        '//  - Parte ALTA multiplo de 2 --> ROTO PARTE ALTA
        '//  - Parte ALTA multiplo de 3 --> ROTO PARTE BAJA
        '//  > Parta ALTA multiplo de 6 --> ROTO PARTE ALTA y BAJA
        '//
        '//  - Parte BAJA multiplo de 2 --> INTERCAMBIO PARTE ALTA y BAJA
        '//  - Parte BAJA multiplo de 3 --> INTERCAMBIAR y ROTAR (sino, al revez)
        '//
        '//  - Parte BAJA y ALTA multiplos de 5 --> MATRIZ 4
        '//  - Parte BAJA multiplos de 5 ---------> MATRIZ 3
        '//  - Parte ALTA multiplos de 5 ---------> MATRIZ 2
        '//  - NINGUNO multiplos de 5 ------------> MATRIZ 1
        Dim aleatH As Integer
        Dim aleatL As Integer
        Dim clave As Integer

        aleatH = Aleatorio(0, 255)
        aleatL = Aleatorio(0, 255)

        clave = 0                ' Reseteo Clave

        ' ROTAR PARTE ALTA
        If ((aleatH <> 0) And ((aleatH Mod 2) = 0)) Then
            ROTH = True
            clave = clave + 64
        Else
            ROTH = False
        End If

        ' ROTAR PARTE BAJA
        If ((aleatH <> 0) And ((aleatH Mod 3) = 0)) Then
            ROTL = True
            clave = clave + 32
        Else
            ROTL = False
        End If

        ' INTERCAMBIAR PARTE BAJA y ALTA
        If ((aleatL <> 0) And ((aleatL Mod 2) = 0)) Then
            XCHG = True
            clave = clave + 16
        Else
            XCHG = False
        End If

        ' INTERCAMBIAR Y DESPUES ROTAR
        If ((aleatL <> 0) And ((aleatL Mod 3) = 0)) Then
            XCHG_ROT = True
            clave = clave + 128
        Else
            XCHG_ROT = False
        End If

        ' MATRIZ
        If ((aleatH <> 0) And (aleatL <> 0) And ((aleatH Mod 5) = 0) And ((aleatL Mod 5) = 0)) Then
            MATRIX = 4
            clave = clave + 8

        ElseIf ((aleatH <> 0) And (aleatL <> 0) And ((aleatH Mod 5) = 0) And ((aleatL Mod 5) <> 0)) Then
            MATRIX = 3
            clave = clave + 4

        ElseIf ((aleatH <> 0) And (aleatL <> 0) And ((aleatH Mod 5) <> 0) And ((aleatL Mod 5) = 0)) Then
            MATRIX = 2
            clave = clave + 2

        Else
            MATRIX = 1
            clave = clave + 1
        End If

        Endata_encriptar_genKEY = DecToHex(clave)

        Endata_encriptar_KEY(Endata_encriptar_genKEY)
    End Function

    '* ENCRIPTAR UN DATO *
    '*********************
    Public Function Endata_Encriptar(dato As String) As String
        Dim HIGH As String
        Dim LOW As String
        Dim temp As String

        HIGH = Left(dato, 1)
        LOW = Right(dato, 1)

        If (XCHG_ROT And XCHG) Then
            ' Intercambio parte ALTA y BAJA y despues roto
            temp = HIGH
            HIGH = LOW
            LOW = temp
        End If

        If ROTH Then
            ' Roto parte Alta
            HIGH = Rotar_4bits(HIGH)
        End If

        If ROTL Then
            ' Roto parte Baja
            LOW = Rotar_4bits(LOW)
        End If

        If ((Not (XCHG_ROT)) And XCHG) Then
            ' Roto y despues intercambio parte ALTA y BAJA
            temp = HIGH
            HIGH = LOW
            LOW = temp
        End If

        dato = HIGH & LOW                ' Rearmo dato

        Endata_Encriptar = Playfair_Encriptar(dato, MATRIX)
    End Function


    '* DESENCRIPTAR UN DATO *
    '************************
    Public Function Endata_Desencriptar(dato As String) As String
        Dim HIGH As String
        Dim LOW As String
        Dim temp As String

        dato = Playfair_Desencriptar(dato, MATRIX)

        HIGH = Left(dato, 1)
        LOW = Right(dato, 1)

        If ((Not (XCHG_ROT)) And XCHG) Then
            ' Roto y despues intercambio parte ALTA y BAJA
            temp = HIGH
            HIGH = LOW
            LOW = temp
        End If

        If ROTH Then
            ' Roto parte Alta
            HIGH = Rotar_4bits(HIGH)
        End If

        If ROTL Then
            ' Roto parte Baja
            LOW = Rotar_4bits(LOW)
        End If

        If (XCHG_ROT And XCHG) Then
            ' Intercambio parte ALTA y BAJA y despues roto
            temp = HIGH
            HIGH = LOW
            LOW = temp
        End If

        Endata_Desencriptar = HIGH & LOW                ' Rearmo dato
    End Function



    '* ENCRIPTAR UN BUFFER DE STRINGS *
    '**********************************
    Public Function Endata_Encriptar_Buffer(clave As String, buffer As String) As String
        Dim N As Integer
        Dim dato As String
        Dim i As Integer

        Call Endata_encriptar_KEY(clave)

        N = Len(buffer)                         ' Largo del buffer
        Endata_Encriptar_Buffer = ""            ' Limpio Resultado Encriptacion
        For i = 1 To N Step 2
            dato = Mid(buffer, i, 2)            ' Tomo 2 bytes
            Endata_Encriptar_Buffer = Endata_Encriptar_Buffer & Endata_Encriptar(dato)
        Next i
    End Function


    '* DESENCRIPTAR UN BUFFER DE STRINGS *
    '*************************************
    Public Function Endata_Desencriptar_Buffer(buffer As String) As String
        Dim N As Integer
        Dim clave As String
        Dim dato As String
        Dim i As Integer

        clave = Mid(buffer, 1, 2)               ' Extraigo Clave
        Call Endata_encriptar_KEY(clave)

        N = Len(buffer)                         ' Largo del buffer
        Endata_Desencriptar_Buffer = ""            ' Limpio Resultado Encriptacion
        For i = 3 To N Step 2
            dato = Mid(buffer, i, 2)            ' Tomo 2 bytes
            Endata_Desencriptar_Buffer = Endata_Desencriptar_Buffer & Endata_Desencriptar(dato)
        Next i
    End Function


    '******************************************************************************************************
    '*** FUNCIONES INTERNAS ***
    '**************************

    '* DETERMINAR METODO DE ENCRIPTACION *
    '*************************************
    Private Sub Endata_encriptar_KEY(clave As String)
        Dim int_clave As Integer

        KEY = clave
        int_clave = HexToDec(clave)

        ' ROTAR PARTE ALTA
        If ((int_clave And 64) / 64 = 1) Then
            ROTH = True
        Else
            ROTH = False
        End If

        ' ROTAR PARTE BAJA
        If ((int_clave And 32) / 32 = 1) Then
            ROTL = True
        Else
            ROTL = False
        End If

        ' INTERCAMBIAR PARTE BAJA y ALTA
        If ((int_clave And 16) / 16 = 1) Then
            XCHG = True
        Else
            XCHG = False
        End If

        ' INTERCAMBIAR Y DESPUES ROTAR
        If ((int_clave And 128) / 128 = 1) Then
            XCHG_ROT = True
        Else
            XCHG_ROT = False
        End If

        ' MATRIZ
        If ((int_clave And 8) / 8 = 1) Then
            MATRIX = 4
        ElseIf ((int_clave And 4) / 4 = 1) Then
            MATRIX = 3
        ElseIf ((int_clave And 2) / 2 = 1) Then
            MATRIX = 2
        ElseIf ((int_clave And 1) / 1 = 1) Then
            MATRIX = 1
        End If
    End Sub


    '* GENERADOR DE NUMERO ALEATORIO *
    '*********************************
    Private Function Aleatorio(Minimo As Long, Maximo As Long) As Long
        Randomize()
        Aleatorio = CLng((Minimo - Maximo) * Rnd + Maximo)
    End Function

    '* ROTAR 4 bits *
    '****************
    Private Function Rotar_4bits(dato As String) As String
        Dim int_dato As Integer
        Dim result As Integer

        result = 0                      ' Limpio Resultado
        int_dato = HexToDec(dato)       ' Convierto Hex->Dec

        If ((int_dato And 1) = 1) Then
            result = result Or 8
        End If

        If ((int_dato And 2) / 2 = 1) Then
            result = result Or 4
        End If

        If ((int_dato And 4) / 4 = 1) Then
            result = result Or 2
        End If

        If ((int_dato And 8) / 8 = 1) Then
            result = result Or 1
        End If

        Rotar_4bits = Right(DecToHex(result), 1) ' Convierto Dec->Hex
    End Function



    Dim PLAYFAIR_MATRIX(4, 4) As String

    '* ENCRIPTAR DATO *
    '******************
    Public Function Playfair_Encriptar(dato As String, matriz As Integer) As String
        Dim dato1 As String                 ' Parte alta del byte: DATO1
        Dim dato2 As String                 ' Parte baja del byte: DATO2
        Dim fil1 As Integer                 ' Fila del DATO 1
        Dim fil2 As Integer                 ' Fila del DATO 2
        Dim col1 As Integer                 ' Columna del DATO 1
        Dim col2 As Integer                 ' Columna del DATO 2
        Dim dato_encriptado As String       ' Resultado de la Encriptacion

        Call Playfair_init(matriz)          ' Crea Matriz de Playfair segun el numero indicado

        dato1 = Left(dato, 1)               ' Extraigo Parte alta del byte
        dato2 = Right(dato, 1)              ' Extrigo Parte baja del byte

        If dato1 <> dato2 Then
            ' Si son distintos, el meto de encriptacion depende de la fila y columna
            ' de dato1 y dato2
            fil1 = Playfair_determineFILA(dato1)
            col1 = Playfair_determineCOLUMNA(dato1)

            fil2 = Playfair_determineFILA(dato2)
            col2 = Playfair_determineCOLUMNA(dato2)

            If fil1 = fil2 Then
                ' Estan en la misma FILA
                dato_encriptado = Playfair_Encriptar_sameFILA(fil1, col1, col2)

            ElseIf col1 = col2 Then
                ' Estan en la misma COLUMNA
                dato_encriptado = Playfair_Encriptar_sameCOLUMNA(col1, fil1, fil2)

            Else
                ' Estan en distinta FILA y COLUMNA
                dato_encriptado = Playfair_Encriptar_NOTsame(col1, fil1, col2, fil2)
            End If
        Else
            dato_encriptado = dato              ' Si son iguales, no los encripto
        End If

        Playfair_Encriptar = dato_encriptado    ' Devuelvo dato encriptado
    End Function


    '* DESENCRIPTAR DATO *
    '*********************
    Public Function Playfair_Desencriptar(dato As String, matriz As Integer) As String
        Dim dato1 As String                 ' Parte alta del byte: DATO1
        Dim dato2 As String                 ' Parte baja del byte: DATO2
        Dim fil1 As Integer                 ' Fila del DATO 1
        Dim fil2 As Integer                 ' Fila del DATO 2
        Dim col1 As Integer                 ' Columna del DATO 1
        Dim col2 As Integer                 ' Columna del DATO 2
        Dim dato_desencriptado As String    ' Resultado de la Desencriptacion

        Call Playfair_init(matriz)          ' Crea Matriz de Playfair segun el numero indicado

        dato1 = Left(dato, 1)               ' Extraigo Parte alta del byte
        dato2 = Right(dato, 1)              ' Extrigo Parte baja del byte

        If dato1 <> dato2 Then
            ' Si son distintos, el meto de desencriptacion depende de la fila y columna
            ' de dato1 y dato2
            fil1 = Playfair_determineFILA(dato1)
            col1 = Playfair_determineCOLUMNA(dato1)

            fil2 = Playfair_determineFILA(dato2)
            col2 = Playfair_determineCOLUMNA(dato2)

            If fil1 = fil2 Then
                ' Estan en la misma FILA
                dato_desencriptado = Playfair_Desencriptar_sameFILA(fil1, col1, col2)

            ElseIf col1 = col2 Then
                ' Estan en la misma COLUMNA
                dato_desencriptado = Playfair_Desencriptar_sameCOLUMNA(col1, fil1, fil2)

            Else
                ' Estan en distinta FILA y COLUMNA
                dato_desencriptado = Playfair_Encriptar_NOTsame(col1, fil1, col2, fil2)
            End If
        Else
            dato_desencriptado = dato              ' Si son iguales, no los desencripto
        End If

        Playfair_Desencriptar = dato_desencriptado ' Devuelvo dato desencriptado
    End Function








    '****************************************************************************************
    '***** FUNCIONES INTERNAS *****
    '******************************

    '* INICIALIZAR METODO PLAYFAIR *
    '*******************************
    Private Sub Playfair_init(matriz As Integer)
        ' Crea matriz de Playfair
        If (matriz = 1) Then
            PLAYFAIR_MATRIX(0, 0) = "B"
            PLAYFAIR_MATRIX(1, 0) = "4"
            PLAYFAIR_MATRIX(2, 0) = "5"
            PLAYFAIR_MATRIX(3, 0) = "9"

            PLAYFAIR_MATRIX(0, 1) = "C"
            PLAYFAIR_MATRIX(1, 1) = "F"
            PLAYFAIR_MATRIX(2, 1) = "0"
            PLAYFAIR_MATRIX(3, 1) = "7"

            PLAYFAIR_MATRIX(0, 2) = "2"
            PLAYFAIR_MATRIX(1, 2) = "E"
            PLAYFAIR_MATRIX(2, 2) = "1"
            PLAYFAIR_MATRIX(3, 2) = "A"

            PLAYFAIR_MATRIX(0, 3) = "3"
            PLAYFAIR_MATRIX(1, 3) = "D"
            PLAYFAIR_MATRIX(2, 3) = "8"
            PLAYFAIR_MATRIX(3, 3) = "6"

        ElseIf (matriz = 2) Then
            PLAYFAIR_MATRIX(0, 0) = "3"
            PLAYFAIR_MATRIX(1, 0) = "1"
            PLAYFAIR_MATRIX(2, 0) = "0"
            PLAYFAIR_MATRIX(3, 0) = "2"

            PLAYFAIR_MATRIX(0, 1) = "6"
            PLAYFAIR_MATRIX(1, 1) = "4"
            PLAYFAIR_MATRIX(2, 1) = "7"
            PLAYFAIR_MATRIX(3, 1) = "5"

            PLAYFAIR_MATRIX(0, 2) = "A"
            PLAYFAIR_MATRIX(1, 2) = "9"
            PLAYFAIR_MATRIX(2, 2) = "B"
            PLAYFAIR_MATRIX(3, 2) = "8"

            PLAYFAIR_MATRIX(0, 3) = "C"
            PLAYFAIR_MATRIX(1, 3) = "E"
            PLAYFAIR_MATRIX(2, 3) = "D"
            PLAYFAIR_MATRIX(3, 3) = "F"

        ElseIf (matriz = 3) Then
            PLAYFAIR_MATRIX(0, 0) = "9"
            PLAYFAIR_MATRIX(1, 0) = "F"
            PLAYFAIR_MATRIX(2, 0) = "0"
            PLAYFAIR_MATRIX(3, 0) = "1"

            PLAYFAIR_MATRIX(0, 1) = "2"
            PLAYFAIR_MATRIX(1, 1) = "E"
            PLAYFAIR_MATRIX(2, 1) = "8"
            PLAYFAIR_MATRIX(3, 1) = "A"

            PLAYFAIR_MATRIX(0, 2) = "5"
            PLAYFAIR_MATRIX(1, 2) = "3"
            PLAYFAIR_MATRIX(2, 2) = "C"
            PLAYFAIR_MATRIX(3, 2) = "B"

            PLAYFAIR_MATRIX(0, 3) = "7"
            PLAYFAIR_MATRIX(1, 3) = "D"
            PLAYFAIR_MATRIX(2, 3) = "6"
            PLAYFAIR_MATRIX(3, 3) = "4"

        ElseIf (matriz = 4) Then
            PLAYFAIR_MATRIX(0, 0) = "0"
            PLAYFAIR_MATRIX(1, 0) = "4"
            PLAYFAIR_MATRIX(2, 0) = "8"
            PLAYFAIR_MATRIX(3, 0) = "C"

            PLAYFAIR_MATRIX(0, 1) = "1"
            PLAYFAIR_MATRIX(1, 1) = "5"
            PLAYFAIR_MATRIX(2, 1) = "9"
            PLAYFAIR_MATRIX(3, 1) = "D"

            PLAYFAIR_MATRIX(0, 2) = "E"
            PLAYFAIR_MATRIX(1, 2) = "A"
            PLAYFAIR_MATRIX(2, 2) = "6"
            PLAYFAIR_MATRIX(3, 2) = "2"

            PLAYFAIR_MATRIX(0, 3) = "F"
            PLAYFAIR_MATRIX(1, 3) = "B"
            PLAYFAIR_MATRIX(2, 3) = "7"
            PLAYFAIR_MATRIX(3, 3) = "3"
        End If
    End Sub


    '* DETERMINACION DE FILA *
    '*************************
    Private Function Playfair_determineFILA(dato As String) As Integer
        ' Determina la fila a la que pertence el dato, dentro de la matriz
        ' de Playfair
        Dim i As Integer
        Dim j As Integer
        Dim fila As Integer

        For i = 0 To 3
            For j = 0 To 3
                If (PLAYFAIR_MATRIX(j, i) = dato) Then
                    fila = i                ' Fila donde esta el dato
                End If
            Next j
        Next i

        Playfair_determineFILA = fila       ' Devuelvo Fila
    End Function


    '* DETERMINACION DE COLUMNA *
    '****************************
    Private Function Playfair_determineCOLUMNA(dato As String) As Integer
        ' Determina la columna a la que pertence el dato, dentro de la
        ' matriz de Playfair
        Dim i As Integer
        Dim j As Integer
        Dim COLUMNA As Integer

        For i = 0 To 3
            For j = 0 To 3
                If (PLAYFAIR_MATRIX(j, i) = dato) Then
                    COLUMNA = j             ' Columna donde esta el dato
                End If
            Next j
        Next i

        Playfair_determineCOLUMNA = COLUMNA ' Devuelvo Columna
    End Function


    '* ENCRIPTAR EN MISMA FILA *
    '***************************
    Private Function Playfair_Encriptar_sameFILA(fila As Integer, col1 As Integer, col2 As Integer) As String
        Dim dat1_encriptado As String
        Dim dat2_encriptado As String

        If col1 = 3 Then
            col1 = 0                        ' Primer Columna
            col2 = col2 + 1                 ' Siguiente Columna
        ElseIf col2 = 3 Then
            col1 = col1 + 1                 ' Siguiente Columna
            col2 = 0                        ' Primer Columna
        Else
            col1 = col1 + 1                 ' Siguiente Columna
            col2 = col2 + 1                 ' Siguiente Columna
        End If

        dat1_encriptado = PLAYFAIR_MATRIX(col1, fila)   ' Encripto dato 1
        dat2_encriptado = PLAYFAIR_MATRIX(col2, fila)   ' Encripto dato 2

        Playfair_Encriptar_sameFILA = dat1_encriptado & dat2_encriptado ' Concateno datos encriptados
    End Function


    '* ENCRIPTAR EN MISMA COLUMANA *
    '*******************************
    Private Function Playfair_Encriptar_sameCOLUMNA(col As Integer, fil1 As Integer, fil2 As Integer) As String
        Dim dat1_encriptado As String
        Dim dat2_encriptado As String

        If fil1 = 3 Then
            fil1 = 0                        ' Primer Fila
            fil2 = fil2 + 1                 ' Siguiente Fila
        ElseIf fil2 = 3 Then
            fil1 = fil1 + 1                 ' Siguiente Fila
            fil2 = 0                        ' Primer Fila
        Else
            fil1 = fil1 + 1                 ' Siguiente Fila
            fil2 = fil2 + 1                 ' Siguiente Fila
        End If

        dat1_encriptado = PLAYFAIR_MATRIX(col, fil1)   ' Encripto dato 1
        dat2_encriptado = PLAYFAIR_MATRIX(col, fil2)   ' Encripto dato 2

        Playfair_Encriptar_sameCOLUMNA = dat1_encriptado & dat2_encriptado ' Concateno datos encriptados
    End Function


    '* DESENCRIPTAR EN MISMA FILA *
    '******************************
    Private Function Playfair_Desencriptar_sameFILA(fila As Integer, col1 As Integer, col2 As Integer) As String
        Dim dat1_desencriptado As String
        Dim dat2_desencriptado As String

        If col1 = 0 Then
            col1 = 3                        ' Ultima Columna
            col2 = col2 - 1                 ' Columna Anterior
        ElseIf col2 = 0 Then
            col1 = col1 - 1                 ' Columna Anterior
            col2 = 3                        ' Ultima Columna
        Else
            col1 = col1 - 1                 ' Columna Anterior
            col2 = col2 - 1                 ' Columna Anterior
        End If

        dat1_desencriptado = PLAYFAIR_MATRIX(col1, fila)   ' Desencripto dato 1
        dat2_desencriptado = PLAYFAIR_MATRIX(col2, fila)   ' Desencripto dato 2

        Playfair_Desencriptar_sameFILA = dat1_desencriptado & dat2_desencriptado ' Concateno datos desencriptados
    End Function


    '* DESENCRIPTAR EN MISMA COLUMANA *
    '**********************************
    Private Function Playfair_Desencriptar_sameCOLUMNA(col As Integer, fil1 As Integer, fil2 As Integer) As String
        Dim dat1_desencriptado As String
        Dim dat2_desencriptado As String

        If fil1 = 0 Then
            fil1 = 3                        ' Ultima Fila
            fil2 = fil2 - 1                 ' Fila Anterior
        ElseIf fil2 = 0 Then
            fil1 = fil1 - 1                 ' Fila Anterior
            fil2 = 3                        ' Ultima Fila
        Else
            fil1 = fil1 - 1                 ' Fila Anterior
            fil2 = fil2 - 1                 ' Fila Anterior
        End If

        dat1_desencriptado = PLAYFAIR_MATRIX(col, fil1)   ' Encripto dato 1
        dat2_desencriptado = PLAYFAIR_MATRIX(col, fil2)   ' Encripto dato 2

        Playfair_Desencriptar_sameCOLUMNA = dat1_desencriptado & dat2_desencriptado ' Concateno datos encriptados
    End Function



    '* ENCRIPTAR/DESENCRIPTAR EN DISTINTA FILA y COLUMNA *
    '*****************************************************
    Private Function Playfair_Encriptar_NOTsame(col1 As Integer, fil1 As Integer, col2 As Integer, fil2 As Integer) As String
        Dim dat1_encriptado As String
        Dim dat2_encriptado As String

        dat1_encriptado = PLAYFAIR_MATRIX(col2, fil1)   ' Encripto dato 1
        dat2_encriptado = PLAYFAIR_MATRIX(col1, fil2)   ' Encripto dato 2

        Playfair_Encriptar_NOTsame = dat1_encriptado & dat2_encriptado ' Concateno datos encriptados
    End Function


End Module
