Imports System.Data.SqlClient

Public Module MensajeMovil

    Public ConnString As String
    Public Status As Integer


    Public Sub EnviarMensaje(ByVal NroFlota As Integer, ByVal NroMovil As Integer, ByVal TipoMensaje As String, ByVal TextoMensaje As String)


        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("insert into mensajes_moviles (nro_mensaje, texto, nro_movil, estado_mensaje, fecha_hora_envio) values (@ProxNroMensaje ,@TextoMensaje, @NroMovil,  'P' ,getdate()", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@ProxNroMensaje", UltimoNumero.ObtenerProximo("mensmovil"))
            SqlCmd.Parameters.AddWithNullableValue("@TextoMensaje", UCase(TextoMensaje))
            SqlCmd.Parameters.AddWithNullableValue("@NroMovil", NroMovil)
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


    Public Sub CancelarMensajesPendientes(ByVal NroMovil As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("update mensajes_moviles set estado_mensaje = 'R' where nro_movil = @NroMovil and estado_mensaje = 'P'", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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