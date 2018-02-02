Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.Logging
Imports vb = Microsoft.VisualBasic
Imports System.Data.SqlClient

Public Module Comunes

    Public TiempoEsperaDespenalizacion As Integer

    Public Class Posicion
        Public Fecha As Date?
        Public Latitud As Single?
        Public Longitud As Single?
        Public Velocidad As Integer?
        Public Rumbo As String
    End Class


    Public Function DecToHex(ByVal intValorDec As Short) As String

        'Convierte Decimales en hexadecimales

        DecToHex = Hex(intValorDec)

        If DecToHex.Length = 1 Then
            DecToHex = "0" & DecToHex
        End If

    End Function


    Public Function HexToDec(ByVal strValorHex As String) As Integer

        Dim strValorL As String
        Dim strValorH As String

        strValorL = vb.Right(strValorHex, 1)

        If Len(strValorHex) = 1 Then
            strValorH = "0"
        Else
            strValorH = vb.Left(strValorHex, 1)
        End If

        HexToDec = CInt("&H" & strValorH) * 16 + CInt("&H" & strValorL)

    End Function

    Public Function HexToDec2(ByVal strValorHex As String) As Long

        HexToDec2 = CLng("&H" & strValorHex)

    End Function

    Public Function HexToDec3(ByVal strValorHex As String) As Long


        Dim strValorL As String
        Dim strValorH As String
        Dim strvalorM As String

        strValorL = Right(strValorHex, 2)
        strValorH = Left(strValorHex, 2)
        strvalorM = Mid(strValorHex, 3, 2)

        HexToDec3 = CLng(HexToDec(strValorH)) * 65536 + CInt(HexToDec(strvalorM)) * 256 + CInt(HexToDec(strValorL))

    End Function

    Public Function DecToHexLarga(ByVal intValorDec As Long) As String

        DecToHexLarga = Hex(intValorDec)

    End Function


    Public Function FormatDate(ByVal Fecha As Date) As String

        Try
            FormatDate = Format(Fecha, "MM/dd/yyyy HH:MM:SS")

        Catch ex As Exception
            FormatDate = Nothing
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

    End Function


    Public Function FormatTime(ByVal DateTime As Date) As String

        Try
            FormatTime = Format(DateTime, "HH:MM")

        Catch ex As Exception
            FormatTime = Nothing
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

    End Function

    'Public Function ArmarDireccion(ByVal lngCodCalle As Integer, ByVal strNombreCalle As String, ByVal lngNumero As Integer, ByVal strUbicacion As String, ByVal intCodLocalidad As Short, ByVal strNombreLocalidad As String) As String


    Public Function CIFNull(ByVal Expresion As Object, ByVal ResultIf As Object) As Object
        CIFNull = IIf(IsNothing(Expresion), ResultIf, Expresion)
    End Function


    Public Function DecriptPassword(ByVal EncriptedPassword As String) As String

        Dim strDecripted As String
        Dim strEncripted As String
        Dim intI As Short
        Dim intPwdLen As Short
        Dim intCodeOffset As Short
        Dim intFirstChar As Short
        Dim intLastchar As Short
        Dim intPosOffset As Short
        Dim strNewChar As String
        Dim intNewPos As Short
        Dim intControl As Short

        Try
            strEncripted = EncriptedPassword

            intControl = 0

            For intI = 0 To 48
                intControl = intControl Or Asc(strEncripted.Substring(intI, 1))
            Next intI

            If Chr(intControl) <> strEncripted.Substring(49, 1) Then
                DecriptPassword = ""
                Exit Function
            End If

            intPwdLen = Asc(strEncripted.Substring(45, 1)) - 30
            intCodeOffset = (Asc(strEncripted.Substring(46, 1)) - 30) * 100 + (Asc(strEncripted.Substring(48, 1)) - 30)
            intPosOffset = Asc(strEncripted.Substring(48, 1)) - 30

            strDecripted = ""

            For intI = 0 To intPwdLen - 1
                intNewPos = 45 - (intPosOffset + (2 * intI) - 1)
                intFirstChar = Asc(strEncripted.Substring(intNewPos, 1)) - 70
                intLastchar = Asc(strEncripted.Substring(intNewPos + 1, 1)) - 70
                strNewChar = Chr(intFirstChar * 100 + intLastchar - intCodeOffset)
                strDecripted += strNewChar
            Next intI

            DecriptPassword = strDecripted
            Exit Function

        Catch ex As Exception
            DecriptPassword = ""
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)

        End Try

    End Function


    Public Function EmpaquetarNumero(ByVal Numero As String) As String

        Try

            Static CantNumeros As Integer
            Static Caracter As String
            Static NuevoNumero As String


            CantNumeros = Numero.Length

            If Not EsPar(CantNumeros) Then
                'el número tiene una cantidad impar de dígitos
                CantNumeros = CantNumeros - 1
                EmpaquetarNumero = Chr(Val((vb.Left(Numero, 1))) + 192)
            Else
                'el número tiene una cantidad par de dígitos
                CantNumeros = CantNumeros - 2
                Caracter = (Val(vb.Left(Numero, 2)))
                EmpaquetarNumero = Chr(Int(Caracter / 10) * 16 + (Caracter - Int(Caracter / 10) * 10))
            End If

            NuevoNumero = vb.Right(Numero, CantNumeros)

            For i As Integer = 1 To CantNumeros - 1 Step 2
                Caracter = Val(Mid(NuevoNumero, i, 2))
                EmpaquetarNumero += Chr(Int(Caracter / 10) * 16 + (Caracter - Int(Caracter / 10) * 10))
            Next

        Catch ex As Exception
            EmpaquetarNumero = Nothing
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

    End Function


    Public Function QuitarCaracterEx(ByVal sValor As String, ByVal sCaracter As String, _
        Optional ByVal sPoner As String = Nothing) As String

        Try

            '----------------------------------------------------------
            ' Cambiar/Quitar caracteres                     (17/Sep/97)
            ' Si se especifica sPoner, se cambiará por ese carácter
            '
            'Esta versión permite cambiar los caracteres    (17/Sep/97)
            'y sustituirlos por el/los indicados
            'a diferencia de QuitarCaracter, no se buscan uno a uno,
            'sino todos juntos
            '
            'Última revisión:           (11/Jun/98)
            '----------------------------------------------------------
            Dim i As Long
            Dim sCh As String
            Dim bPoner As Boolean
            Dim iLen As Integer
            sCh = ""


            bPoner = False
            If Not IsNothing(sPoner) Then
                sCh = sPoner
                bPoner = True
            End If
            iLen = Len(sCaracter)
            If iLen = 0 Then
                QuitarCaracterEx = sValor
                Exit Function
            End If

            'Si el caracter a quitar/cambiar es chr(0), usar otro método
            If Asc(sCaracter) = 0 Then
                'Quitar todos los chr(0) del final
                Do While vb.Right(sValor, 1) = Chr(0)
                    sValor = vb.Left(sValor, Len(sValor) - 1)
                    If Len(sValor) = 0 Then Exit Do
                Loop
                iLen = 1
                Do
                    i = InStr(iLen, CStr(sValor), CStr(sCaracter))
                    If i Then
                        If bPoner Then
                            sValor = vb.Left(sValor, i - 1) & sCh & Mid(sValor, i + 1)
                        Else
                            sValor = vb.Left(sValor, i - 1) & Mid(sValor, i + 1)
                        End If
                        iLen = i
                    Else
                        'ya no hay más, salir del bucle
                        Exit Do
                    End If
                Loop
            Else
                i = 1
                Do While i <= Len(sValor)

                    If vb.Mid(sValor, i, iLen) = sCaracter Then
                        If bPoner Then
                            sValor = vb.Left(sValor, i - 1) & sCh & vb.Mid(sValor, i + iLen)
                            i = i - 1
                            'Si lo que hay que poner está incluido en
                            'lo que se busca, incrementar el puntero
                            '                                   (11/Jun/98)
                            If InStr(sCh, sCaracter) Then
                                i = i + 1
                            End If
                        Else
                            sValor = vb.Left(sValor, i - 1) & vb.Mid(sValor, i + iLen)
                        End If
                    End If

                    i = i + 1
                Loop
            End If

            QuitarCaracterEx = sValor

        Catch ex As Exception
            QuitarCaracterEx = Nothing
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

    End Function

    Public Function EsPar(ByVal Numero As Integer) As Boolean
        EsPar = IIf(Numero - (Int(Numero / 2) * 2) = 0, True, False)
    End Function


    Public Function ConvertirHoraGMT(ByVal fechaHora As Date, ByVal GMTDif As Integer) As Date
        ConvertirHoraGMT = DateAdd(DateInterval.Hour, GMTDif, fechaHora)
        ConvertirHoraGMT = fechaHora
    End Function


    Public Function ByteArrayToString(ByVal bdata As Byte()) As String
        Dim hex As System.Text.StringBuilder = New System.Text.StringBuilder(bdata.Length * 2)
        For Each b As Byte In bdata
            hex.AppendFormat("{0:x2}", b)
        Next
        Return hex.ToString()
    End Function

    Public Function QuitarAcentos(ByVal texto As String) As String

        'Cambia los caracteres con tildes, ñ, ç por letras sin tilde o N o C
        'Elimina las tilde spor completo, puede ser necesario al pasar datos
        'desde UTF-8 a sistemas con ISO-8859-1


        If Not IsNothing(texto) Then
            texto = Replace(texto, "¡", "", 1, Len(texto), 1)
            texto = Replace(texto, "¿", "", 1, Len(texto), 1)
            texto = Replace(texto, "'", "", 1, Len(texto), 1)

            texto = Replace(texto, vbCrLf, "", 1, Len(texto), 1)
            texto = Replace(texto, "Nº", "Nro", 1, Len(texto), 1)
            texto = Replace(texto, "nº", "Nro", 1, Len(texto), 1)


            texto = Replace(texto, "á", "a", 1, Len(texto), 1)
            texto = Replace(texto, "é", "e", 1, Len(texto), 1)
            texto = Replace(texto, "í", "i", 1, Len(texto), 1)
            texto = Replace(texto, "ó", "o", 1, Len(texto), 1)
            texto = Replace(texto, "ú", "u", 1, Len(texto), 1)
            texto = Replace(texto, "ñ", "n", 1, Len(texto), 1)
            texto = Replace(texto, "ç", "c", 1, Len(texto), 1)

            texto = Replace(texto, "Á", "A", 1, Len(texto), 1)
            texto = Replace(texto, "É", "E", 1, Len(texto), 1)
            texto = Replace(texto, "Í", "I", 1, Len(texto), 1)
            texto = Replace(texto, "Ó", "O", 1, Len(texto), 1)
            texto = Replace(texto, "Ú", "U", 1, Len(texto), 1)
            texto = Replace(texto, "Ñ", "N", 1, Len(texto), 1)
            texto = Replace(texto, "Ç", "C", 1, Len(texto), 1)

            texto = Replace(texto, "à", "a", 1, Len(texto), 1)
            texto = Replace(texto, "è", "e", 1, Len(texto), 1)
            texto = Replace(texto, "ì", "i", 1, Len(texto), 1)
            texto = Replace(texto, "ò", "o", 1, Len(texto), 1)
            texto = Replace(texto, "ù", "u", 1, Len(texto), 1)

            texto = Replace(texto, "À", "A", 1, Len(texto), 1)
            texto = Replace(texto, "È", "E", 1, Len(texto), 1)
            texto = Replace(texto, "Ì", "I", 1, Len(texto), 1)
            texto = Replace(texto, "Ò", "O", 1, Len(texto), 1)
            texto = Replace(texto, "Ù", "U", 1, Len(texto), 1)

            texto = Replace(texto, "ä", "a", 1, Len(texto), 1)
            texto = Replace(texto, "ë", "e", 1, Len(texto), 1)
            texto = Replace(texto, "ï", "i", 1, Len(texto), 1)
            texto = Replace(texto, "ö", "o", 1, Len(texto), 1)
            texto = Replace(texto, "ü", "u", 1, Len(texto), 1)

            texto = Replace(texto, "Ä", "A", 1, Len(texto), 1)
            texto = Replace(texto, "Ë", "E", 1, Len(texto), 1)
            texto = Replace(texto, "Ï", "I", 1, Len(texto), 1)
            texto = Replace(texto, "Ö", "O", 1, Len(texto), 1)
            texto = Replace(texto, "Ü", "U", 1, Len(texto), 1)

            texto = Replace(texto, "â", "a", 1, Len(texto), 1)
            texto = Replace(texto, "ê", "e", 1, Len(texto), 1)
            texto = Replace(texto, "î", "i", 1, Len(texto), 1)
            texto = Replace(texto, "ô", "o", 1, Len(texto), 1)
            texto = Replace(texto, "û", "u", 1, Len(texto), 1)

            texto = Replace(texto, "Â", "A", 1, Len(texto), 1)
            texto = Replace(texto, "Ê", "E", 1, Len(texto), 1)
            texto = Replace(texto, "Î", "I", 1, Len(texto), 1)
            texto = Replace(texto, "Ô", "O", 1, Len(texto), 1)
            texto = Replace(texto, "Û", "U", 1, Len(texto), 1)
        Else
            texto = ""
        End If

        Return texto


    End Function

    Public Function IsNull(Of T)(ByVal Value As T, Optional ByVal DefaultValue As T = Nothing) As T

        If Value Is Nothing OrElse IsDBNull(Value) Then
            Return DefaultValue
        Else
            Return Value
        End If

    End Function

End Module