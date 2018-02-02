Imports System.Data.SqlClient

Public Module Alarma

    Public ConnString As String
    Public Status As Integer



    Public Sub Activar(ByVal ObjMovil As Movil.DatosMovil, ByVal Causa As String, ByVal MovilPosicion As Comunes.Posicion)

        Dim IdAlarma As Integer

        Try


            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("Select IdAlarma from Alarmas where Nro_movil = @NroMovil and Estado_Alarma = 'A'", DBConn)
            SqlCmd.CommandType = CommandType.Text

            SqlCmd.Parameters.AddWithNullableValue("@NrMovil", ObjMovil.NroMovil)
            Dim Regs = SqlCmd.ExecuteReader()

            If Regs.HasRows Then

                Regs.Read()
                IdAlarma = Regs("IdAlarma")

            Else

                Dim SqlInsert1 As New SqlCommand("Alarma_Grabar", DBConn)
                SqlInsert1.CommandType = CommandType.StoredProcedure

                SqlInsert1.Parameters.AddWithNullableValue("@FechaHora", Now)
                SqlInsert1.Parameters.AddWithNullableValue("@Movil", ObjMovil.NroMovil)
                SqlInsert1.Parameters.AddWithNullableValue("@Conductor", ObjMovil.NroConductor)
                SqlInsert1.Parameters.AddWithNullableValue("@Estado", "A")
                SqlCmd.Parameters("ReturnValue").Direction = ParameterDirection.ReturnValue

                SqlCmd.ExecuteScalar()

                IdAlarma = SqlCmd.Parameters("ReturnValue").Value

            End If


            Dim SqlInsert2 As New SqlCommand("Insert into ALARMAS_EVENTOS (IdAlarma, Fecha_Hora_Evento, Nro_Movil, Codigo_Evento, IdOperador, Latitud, Longitud, Velocidad, Rumbo, Estado_Operativo)  values (@IdAlarma, getdate(), @NroMovil , 'AC' ,@NroInternoUsuario , @lLatitud , @Longitud, @Velocidad , @Rumbo,@EstadoOper", DBConn)
            SqlInsert2.CommandType = CommandType.Text

            SqlInsert2.Parameters.AddWithNullableValue("@IdAlarma", IdAlarma)
            SqlInsert2.Parameters.AddWithNullableValue("@NrInternoUsuario", Nothing)
            SqlInsert2.Parameters.AddWithNullableValue("@Latitud", MovilPosicion.Latitud)
            SqlInsert2.Parameters.AddWithNullableValue("@Longitud", MovilPosicion.Longitud)
            SqlInsert2.Parameters.AddWithNullableValue("@Velocidad", MovilPosicion.Velocidad)
            SqlInsert2.Parameters.AddWithNullableValue("@EstadoOper", MovilPosicion.Rumbo)
            SqlInsert2.Parameters.AddWithNullableValue("@NrMovil", ObjMovil.NroMovil)

            SqlInsert2.ExecuteNonQuery()


            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()
            Status = 0

        Catch ex As Exception
            Status = 1
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

        Return

    End Sub
End Module
