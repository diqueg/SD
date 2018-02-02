Imports System.Data.SqlClient

Public Module MensajeOper

    Public ConnString As String
    Public Status As Integer

    Public Sub BorrarMensajesTodos()

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("delete from mensajes_oper ", DBConn)

            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteNonQuery

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

    Public Sub BorrarMensajesAntiguos()

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim HoraLimpieza = DateAdd("n", -40, Now)

            Dim SqlCmd As New SqlCommand("delete from mensajes_oper where fecha_hora_mensaje < @0", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@0", HoraLimpieza)

            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteNonQuery

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

    Public Sub EnviarMensaje(ByVal NroFlota As Integer, ByVal NroMovil As Integer, ByVal TipoMensaje As String, ByVal TextoMensaje As String)


        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("insert into mensajes_oper (nro_mensaje, nro_movil, nro_flota, tipo_mensaje, fecha_hora_mensaje, texto, estado_mensaje) values (@ProxNroMensaje, @NroMovil, @NroFlota, @TipoMensaje, getdate(), @TextoMensaje, 'P')", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@ProxNroMensaje", UltimoNumero.ObtenerProximo("mensoper"))
            SqlCmd.Parameters.AddWithNullableValue("@NroMovil", NroMovil)
            SqlCmd.Parameters.AddWithNullableValue("@NroFlota", NroFlota)
            SqlCmd.Parameters.AddWithNullableValue("@TipoMensaje", TipoMensaje)
            SqlCmd.Parameters.AddWithNullableValue("@TextoMensaje", TextoMensaje)

            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteNonQuery

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
