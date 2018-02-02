Imports System.Data.SqlClient
Public Module SMS
    Public ConnString As String
    Public Status As Integer
    Public Sub PrepararMensajePosicionMovilAsignado(ByVal Telefono As String, ByVal idMovil As Integer, ByVal latitud As Single, ByVal longitud As Single)

        Try
            If latitud = 0 Then Exit Try

            Dim TextoAEnviar As String = "LatLngMsg:" & latitud & "," & longitud

            EnviarSMS(3, Telefono, idMovil, TextoAEnviar, 100)

            If Status <> 0 Then Exit Try

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

    End Sub

    Public Sub EnviarSMS(ByVal Tipo As Integer, ByVal Telefono As String, ByVal Movil As Integer, ByVal Texto As String, ByVal Distancia As Integer)
        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("SMS_Inserta_Movil", DBConn)
            SqlCmd.CommandType = CommandType.StoredProcedure

            SqlCmd.Parameters.AddWithNullableValue("@Telefono", Telefono)
            SqlCmd.Parameters.AddWithNullableValue("@Tipo", Tipo)
            SqlCmd.Parameters.AddWithNullableValue("@Movil", Movil)
            SqlCmd.Parameters.AddWithNullableValue("@Direccion", Texto)
            SqlCmd.Parameters.AddWithNullableValue("@Distancia", Distancia)

            SqlCmd.ExecuteNonQuery()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            Status = 1
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

    End Sub

End Module
