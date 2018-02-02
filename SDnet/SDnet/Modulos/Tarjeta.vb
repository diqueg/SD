Imports System.Data.SqlClient

Public Module Tarjeta

    Public ConnString As String
    Public Status As Integer

    Public Class DatosTarjeta
        Public NroTarjeta As String
        Public Importe As String
        Public CodTransaccion As String
    End Class

    Public Function ObtenerSaldo(ByVal NroTarjeta As String) As String

        Dim SaldoTarjeta As String

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("Tarjeta_Consultar_Saldo", DBConn)
            SqlCmd.CommandType = CommandType.StoredProcedure
            SqlCmd.Parameters.AddWithNullableValue("@NroTarjeta", NroTarjeta)
            SqlCmd.Parameters.Add("@Saldo", SqlDbType.Decimal)
            SqlCmd.Parameters("@Saldo").Precision = 18
            SqlCmd.Parameters("@Saldo").Scale = 0
            SqlCmd.Parameters("@Saldo").Direction = ParameterDirection.Output
            SqlCmd.Parameters.Add("@CodResultado", SqlDbType.VarChar, 2, ParameterDirection.Output)

            Dim Regs = SqlCmd.ExecuteNonQuery

            Dim SaldoInformar As String = String.Empty

            If SqlCmd.Parameters("@CodResultado").Value.ToString = "OK" Then

                Dim Estado = Format(SqlCmd.Parameters("@CodResultado").Value.ToString, "0000.00")

                Dim Saldo = Comunes.DecToHexLarga(CLng(Left(Estado, 4) & Right(Estado, 2)))

                Select Case Len(Saldo)
                    Case 0
                        SaldoInformar = "000000" & Saldo
                    Case 1
                        SaldoInformar = "00000" & Saldo
                    Case 2
                        SaldoInformar = "0000" & Saldo
                    Case 3
                        SaldoInformar = "000" & Saldo
                    Case 4
                        SaldoInformar = "00" & Saldo
                    Case 5
                        SaldoInformar = "0" & Saldo
                End Select
                SaldoInformar = "02" & SaldoInformar
            Else
                SaldoInformar = "01"
            End If

            Dim Clave = Encriptado.Endata_encriptar_genKEY
            SaldoTarjeta = Clave & Encriptado.Endata_Encriptar_Buffer(Clave, SaldoInformar)

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            SaldoTarjeta = String.Empty
            Status = 1
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

        Return SaldoTarjeta

    End Function


    Public Function RecargarSaldo(ByVal NroTarjeta As String, ByVal Importe As Decimal, ByVal NroIntTransaccion As Integer, ByVal Idmovil As Integer, ByVal IdViaje As Long, ByVal IdConductor As Integer) As String

        Dim Clave As String
        Dim ImporteCarga As String
        Dim CodigoOp As String
        Dim Saldo As String
        Dim SaldoInformar As String
        Dim Resultado As String = String.Empty

        SaldoInformar = ""

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("Tarjeta_Carga_Saldo_Intentar", DBConn)
            SqlCmd.CommandType = CommandType.StoredProcedure

            SqlCmd.Parameters.AddWithNullableValue("@NroTarjeta", NroTarjeta)
            SqlCmd.Parameters.AddWithNullableValue("@Importe", Importe)
            SqlCmd.Parameters.AddWithNullableValue("@IdViaje", IdViaje)
            SqlCmd.Parameters.AddWithNullableValue("@IdMovil", Idmovil)
            SqlCmd.Parameters.AddWithNullableValue("@IdConductor", IdConductor)
            SqlCmd.Parameters.AddWithNullableValue("@NroInternoTransaccion", NroIntTransaccion)
            SqlCmd.Parameters.Add("@CodResultado", SqlDbType.VarChar, 2, ParameterDirection.Output)
            SqlCmd.Parameters.Add("@IdTransaccion", SqlDbType.VarChar, 2, ParameterDirection.Output)
            SqlCmd.Parameters.Add("@Saldo", SqlDbType.Decimal)
            SqlCmd.Parameters("@Saldo").Precision = 18
            SqlCmd.Parameters("@Saldo").Scale = 0
            SqlCmd.Parameters("@Saldo").Direction = ParameterDirection.Output

            SqlCmd.ExecuteNonQuery()

            Select Case SqlCmd.Parameters("@CodResultado").Value.ToString

                Case "OK", "ND"
                    Clave = Endata_encriptar_genKEY()
                    CodigoOp = Format(SqlCmd.Parameters("@IdTransaccion").Value.ToString, "00000000")

                    ImporteCarga = Importe    'Format(strimporte, "00") & Format(DecToHexLarga(Left(CStr(Importe), (Len(CStr(Importe)) - CInt(strimporte) - 1)) & Right(CStr(Importe), CInt(strimporte))), "000000")

                    Saldo = Format(SqlCmd.Parameters("@Saldo").Value, "0000.00")
                    Saldo = DecToHexLarga(CLng(Left(Saldo, 4) & Right(Saldo, 2)))


                    Select Case Len(Saldo)
                        Case 0
                            SaldoInformar = "000000" & Saldo
                        Case 1
                            SaldoInformar = "00000" & Saldo
                        Case 3
                            SaldoInformar = "0000" & Saldo
                        Case 2
                            SaldoInformar = "000" & Saldo
                        Case 4
                            SaldoInformar = "00" & Saldo
                        Case 5
                            SaldoInformar = "0" & Saldo
                    End Select

                    Resultado = Clave & Endata_Encriptar_Buffer(Clave, CodigoOp & ImporteCarga & SaldoInformar)

                Case "TI", "NV"
                    Clave = Endata_encriptar_genKEY()
                    Resultado = Clave & Endata_Encriptar_Buffer(Clave, "01")

                Case "SS"
                    Clave = Endata_encriptar_genKEY()
                    Resultado = Clave & Endata_Encriptar_Buffer(Clave, "00")

                Case "NA"
                    Clave = Endata_encriptar_genKEY()
                    Resultado = Clave & Endata_Encriptar_Buffer(Clave, "03")


            End Select

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0


        Catch ex As Exception
            Resultado = String.Empty
            Status = 1
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

        Return Resultado

    End Function

    Public Function HacerTransaccion(ByVal NroTarjeta As String, ByVal Importe As Decimal, ByVal NroIntTransaccion As Integer, ByVal Idmovil As Integer, ByVal IdViaje As Long, ByVal IdConductor As Integer) As String


        Dim Clave As String
        Dim ImporteCarga As String
        Dim CodigoOp As String
        Dim Saldo As String
        Dim SaldoInformar As String
        Dim Recargar As String = String.Empty
        Dim Transaccion As String
        Dim Resultado As String = String.Empty

        SaldoInformar = ""

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("Tarjeta_Debito_Viaje_Intentar", DBConn)
            SqlCmd.CommandType = CommandType.StoredProcedure

            SqlCmd.Parameters.AddWithNullableValue("@NroTarjeta", NroTarjeta)
            SqlCmd.Parameters.AddWithNullableValue("@Importe", Importe)
            SqlCmd.Parameters.AddWithNullableValue("@IdViaje", IdViaje)
            SqlCmd.Parameters.AddWithNullableValue("@IdMovil", Idmovil)
            SqlCmd.Parameters.AddWithNullableValue("@IdConductor", IdConductor)
            SqlCmd.Parameters.AddWithNullableValue("@NroInternoTransaccion", NroIntTransaccion)
            SqlCmd.Parameters.Add("@CodResultado", SqlDbType.VarChar, 2, ParameterDirection.Output)
            SqlCmd.Parameters.Add("@IdTransaccion", SqlDbType.VarChar, 2, ParameterDirection.Output)
            SqlCmd.Parameters.Add("@Saldo", SqlDbType.Decimal)
            SqlCmd.Parameters("@Saldo").Precision = 18
            SqlCmd.Parameters("@Saldo").Scale = 0
            SqlCmd.Parameters("@Saldo").Direction = ParameterDirection.Output

            SqlCmd.ExecuteNonQuery()

            Saldo = Format(SqlCmd.Parameters("@Saldo").Value, "000.00")
            Transaccion = SqlCmd.Parameters("@IdTransaccion").Value
            CodigoOp = Format(Transaccion, "00000000")

            ImporteCarga = Importe    'Format(strimporte, "00") & Format(DecToHexLarga(Left(CStr(Importe), (Len(CStr(Importe)) - CInt(strimporte) - 1)) & Right(CStr(Importe), CInt(strimporte))), "000000")
            'Resultado = clave & Endata_Encriptar_Buffer(clave, strCodigoOp & strImporteCarga)

            Select Case SqlCmd.Parameters("@CodResultado").Value

                Case "OK", "ND"
                    Clave = Endata_encriptar_genKEY()
                    CodigoOp = Format(Transaccion, "00000000")
                    ImporteCarga = Importe    'Format(strimporte, "00") & Format(DecToHexLarga(Left(CStr(Importe), (Len(CStr(Importe)) - CInt(strimporte) - 1)) & Right(CStr(Importe), CInt(strimporte))), "000000")

                    Saldo = Format(SqlCmd.Parameters("@Saldo").Value, "0000.00")
                    Saldo = DecToHexLarga(CLng(Left(Saldo, 4) & Right(Saldo, 2)))

                    Select Case Len(Saldo)
                        Case 0
                            SaldoInformar = "000000" & Saldo
                        Case 1
                            SaldoInformar = "00000" & Saldo
                        Case 3
                            SaldoInformar = "0000" & Saldo
                        Case 2
                            SaldoInformar = "000" & Saldo
                        Case 4
                            SaldoInformar = "00" & Saldo
                        Case 5
                            SaldoInformar = "0" & Saldo
                    End Select

                    Resultado = Clave & Endata_Encriptar_Buffer(Clave, CodigoOp & ImporteCarga & SaldoInformar)

                Case "TI", "NV"
                    Clave = Endata_encriptar_genKEY()
                    Resultado = Clave & Endata_Encriptar_Buffer(Clave, "01")

                Case "SS"
                    Clave = Endata_encriptar_genKEY()
                    Resultado = Clave & Endata_Encriptar_Buffer(Clave, "00")

                Case "NA"
                    Clave = Endata_encriptar_genKEY()
                    Resultado = Clave & Endata_Encriptar_Buffer(Clave, "03")

            End Select


            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0


        Catch ex As Exception
            Resultado = String.Empty
            Status = 1
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

        Return Resultado

    End Function


    Public Function ConfirmaRecarga(ByVal Transaccion As Long) As String

        Dim Resultado As String

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("Tarjeta_Carga_Saldo_Confirmar", DBConn)
            SqlCmd.CommandType = CommandType.StoredProcedure

            SqlCmd.Parameters.AddWithNullableValue("@Transaccion", Transaccion)
            SqlCmd.Parameters.Add("@CodResultado", SqlDbType.VarChar, 2, ParameterDirection.Output)

            SqlCmd.ExecuteNonQuery()

            Resultado = SqlCmd.Parameters("@CodResultado").Value

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            Resultado = String.Empty
            Status = 1
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

        Return Resultado

    End Function



End Module