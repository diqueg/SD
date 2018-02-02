Imports System.Data.SqlClient

Public Module Turno

    Public ConnString As String
    Public Status As Integer


    Public Class DatosTurno

        Public NroMovil As Int16
        Public FechaHoraInicio As Date?
        Public FechaHoraFin As Date?
        Public NroConductor As Integer?
        Public CantExcesosVelocidad As Long?
        Public KmRecorridosOcupado As Long?
        Public KmRecorridosTotales As Long?
        Public CantTicketsEmitidos As Long?
        Public CantDesconexionesReloj As Long?
        Public CantViajesPorSensor As Long?
        Public CantViajesTarifa1 As Long?
        Public CantViajesTarifa2 As Long?
        Public CantViajesTarifa3 As Long?
        Public CantViajesTarifa4 As Long?
        Public CantRecaudadaTarifa1 As Decimal?
        Public CantRecaudadaTarifa2 As Decimal?
        Public CantRecaudadaTarifa3 As Decimal?
        Public CantRecaudadaTarifa4 As Decimal?
        Public NroCorrelativoViaje As Int16?
        Public FinTurnoOK As Boolean?
        Public Liquidado As Int16?
        Public Correccion As Int16?
        Public NroInternoUsuario As Int16?
        Public LatitudInicio As Single?
        Public LongitudInicio As Single?
        Public LatitudFin As Single?
        Public LongitudFin As Single?
        
    End Class

    Public Sub IniciarTurno(ByVal NroMovil As Integer, ByVal NroConductor As Integer, ByVal FechaHora As Date, ByVal Latitud As Single, ByVal Longitud As Single)

        Dim ProxNroSec As Integer

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("select top 1 nro_secuencia_turno from historia_turnos where nro_movil = @NroMovil  order by nro_movil, nro_secuencia_turno desc", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@NroMovil", NroMovil)
            SqlCmd.CommandType = CommandType.Text

            Dim UltNroSec = SqlCmd.ExecuteScalar

            If Not IsNothing(UltNroSec) Then
                ProxNroSec = UltNroSec + 1
            Else
                ProxNroSec = 1
            End If

            UltNroSec.close()


            Dim SqlInsert1 As New SqlCommand("insert into historia_turnos (nro_movil, nro_secuencia_turno, fecha_hora_inicio, fecha_hora_fin, nro_conductor, cant_excesos_vel, km_rec_ocupado, km_rec_totales, cant_tickets_emitidos, cant_desc_reloj, cant_viajes_por_sensor, cant_viajes_t1, cant_recaud_t1, cant_viajes_t2, cant_recaud_t2, cant_viajes_t3, cant_recaud_t3, cant_viajes_t4, cant_recaud_t4, nro_correlativo_viaje, ind_fin_turno_ok, liquidado, correccion, NroInternoUsuario, latitud_inicio, longitud_inicio, latitud_fin, longitud_fin) values (@NroMovil, @ProxNroSecuencia,@FechaHoraInicio, Null, @NroConductor, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 'FALSE', 0, 0, 0,@LatitudInicio, @LongitudInicio, 0, 0)", DBConn)

            SqlInsert1.Parameters.AddWithNullableValue("@ProxNroSecuencia", ProxNroSec)
            SqlInsert1.Parameters.AddWithNullableValue("@FechaHoraInicio", FechaHora)
            SqlInsert1.Parameters.AddWithNullableValue("@Latitud", Latitud)
            SqlInsert1.Parameters.AddWithNullableValue("@Longitud", Longitud)
            SqlInsert1.Parameters.AddWithNullableValue("@NroConductor", NroConductor)
            SqlInsert1.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlInsert1.CommandType = CommandType.Text

            SqlInsert1.ExecuteNonQuery()

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

    Public Sub LogOff(ByVal NroMovil As Integer, ByVal FechaHoraFin As Date, ByVal NroConductor As Integer, ByVal CantExcesosVel As Integer, ByVal KmRecOcupado As Single, ByVal KmRecTotales As Single, ByVal CantTicketsEmitidos As Integer, ByVal CantDescReloj As Integer, ByVal CantViajesPorSensor As Integer, ByVal CantViajesT1 As Integer, ByVal CantRecaudT1 As Decimal, ByVal CantViajesT2 As Integer, ByVal CantRecaudT2 As Decimal, ByVal CantViajesT3 As Integer, ByVal CantRecaudT3 As Decimal, ByVal CantViajesT4 As Integer, ByVal CantRecaudT4 As Decimal, ByVal NroCorrelativoViaje As Integer, ByVal LatitudFin As Single, ByVal LongitudFin As Single)


        Dim ProxNroSec As Integer

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("select * from historia_turnos where nro_movil = @NroMovil and fecha_hora_fin is null order by fecha_hora_inicio desc", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@NroMovil", NroMovil)
            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            If Not Regs.HasRows Then

                ' Si no encuentra el un turno abierto para cerrar, crea uno con la misma hora de inicio que la de fin
                ' de turno que se cierra y con los datos del cierre ya cargados

                Dim SqlCmd1 As New SqlCommand("select top 1 nro_secuencia_turno from historia_turnos where nro_movil = @NroMovil  order by nro_movil, nro_secuencia_turno desc", DBConn)

                SqlCmd1.Parameters.AddWithNullableValue("@NroMovil", NroMovil)
                SqlCmd1.CommandType = CommandType.Text

                Dim UltNroSec = SqlCmd1.ExecuteScalar

                If Not IsNothing(UltNroSec) Then
                    ProxNroSec = UltNroSec + 1
                Else
                    ProxNroSec = 1
                End If


                Dim SqlInsert1 As New SqlCommand("insert into historia_turnos (nro_movil, nro_secuencia_turno, fecha_hora_inicio, fecha_hora_fin, nro_conductor, cant_excesos_vel, km_rec_ocupado, km_rec_totales, cant_tickets_emitidos, cant_desc_reloj, cant_viajes_por_sensor, cant_viajes_t1, cant_recaud_t1, cant_viajes_t2, cant_recaud_t2, cant_viajes_t3, cant_recaud_t3, cant_viajes_t4, cant_recaud_t4, nro_correlativo_viaje, ind_fin_turno_ok, liquidado, correccion, NroInternoUsuario, latitud_inicio, longitud_inicio, latitud_fin, longitud_fin) values (@NroMovil, @NroSecuenciaTurno, @FechaHoraInicio, @FechaHoraFin, @NroConductor, @CantExecesosVel,@KmRecOcupado, @KmRecTotales,@CantTicketsEmitidos, @CantDescRelos, @CantViajesPorSensor, @CantViajesT1, @CantRecaudT1, @CantViajesT2, @CantRecaudT2, @CantViajesT3, @CantRecaudT3, @CantViajesT4, @CantRecaudT4, @NroCorrelativoViaje, @IndFinTurnoOK, @Liquidado, @Correcion, @NroInternoUsuario, @LatitudInicio, @LongitudInicio, @LatitudFin, @LongitudFin)", DBConn)

                SqlInsert1.Parameters.AddWithNullableValue("@NroMovil", NroMovil)
                SqlInsert1.Parameters.AddWithNullableValue("@NroSecuenciaTurno", ProxNroSec)
                SqlInsert1.Parameters.AddWithNullableValue("@FechaHoraInicio", FechaHoraFin)
                SqlInsert1.Parameters.AddWithNullableValue("@FechaHoraFin", FechaHoraFin)
                SqlInsert1.Parameters.AddWithNullableValue("@NroConductor", NroConductor)
                SqlInsert1.Parameters.AddWithNullableValue("@CantExcesosVel", CantExcesosVel)
                SqlInsert1.Parameters.AddWithNullableValue("@KmRecOcupado", KmRecOcupado)
                SqlInsert1.Parameters.AddWithNullableValue("@KmRecTotales", KmRecTotales)
                SqlInsert1.Parameters.AddWithNullableValue("@CantTicketsEmitidos", CantTicketsEmitidos)
                SqlInsert1.Parameters.AddWithNullableValue("@CantDescReloj", CantDescReloj)
                SqlInsert1.Parameters.AddWithNullableValue("@CantViajesPorSensor", CantViajesPorSensor)
                SqlInsert1.Parameters.AddWithNullableValue("@CantViajesT1", CantViajesT1)
                SqlInsert1.Parameters.AddWithNullableValue("@CantRecaudT1", CantRecaudT1)
                SqlInsert1.Parameters.AddWithNullableValue("@CantViajesT2", CantViajesT2)
                SqlInsert1.Parameters.AddWithNullableValue("@CantRecaudT2", CantRecaudT2)
                SqlInsert1.Parameters.AddWithNullableValue("@CantViajesT3", CantViajesT3)
                SqlInsert1.Parameters.AddWithNullableValue("@CantRecaudT3", CantRecaudT3)
                SqlInsert1.Parameters.AddWithNullableValue("@CantViajesT4", CantViajesT4)
                SqlInsert1.Parameters.AddWithNullableValue("@CantRecaudT4", CantRecaudT4)
                SqlInsert1.Parameters.AddWithNullableValue("@NroCorrelativoViaje", NroCorrelativoViaje)
                SqlInsert1.Parameters.AddWithNullableValue("@IndFinTurnoOK", True)
                SqlInsert1.Parameters.AddWithNullableValue("@Liquidado", 0)
                SqlInsert1.Parameters.AddWithNullableValue("@Correcion", 0)
                SqlInsert1.Parameters.AddWithNullableValue("@NroInternoUsuario", 0)
                SqlInsert1.Parameters.AddWithNullableValue("@LatitudInicio", LatitudFin)
                SqlInsert1.Parameters.AddWithNullableValue("@LongitudInicio", LongitudFin)
                SqlInsert1.Parameters.AddWithNullableValue("@LatitudFin", LatitudFin)
                SqlInsert1.Parameters.AddWithNullableValue("@LongitudFin", LongitudFin)

                SqlInsert1.CommandType = CommandType.Text

                SqlInsert1.ExecuteNonQuery()


            Else


                Dim SqlUpdate1 As New SqlCommand("update historia_turnos set fecha_hora_fin = @FechaHoraFin , cant_excesos_vel = @CantExcesosVel, km_rec_ocupado = @KmRecOcupado, km_rec_totales = @KmRecTotales, cant_tickets_emitidos = @CantTicketsEmitidos, cant_desc_reloj = @CantDescReloj, cant_viajes_por_sensor = @CantViajesPorSensor, cant_viajes_t1 = @CantViajesT1, cant_viajes_t2 = @CantViajesT2, cant_viajes_t3 = @CantViajesT3, cant_viajes_t4 = @CantViajes4, cant_recaud_t1 = @CantRecaudT1, cant_recaud_t2 = @CantRecaudT2, cant_recaud_t3= @CantRecaudT3, cant_recaud_t4 = @CantRecaudT4, nro_correlativo_viaje = @NroCorrelativoViaje, ind_fin_turno_ok = @IndFinTurnoOK, latitud_fin = @LatitudFin , longitud_fin = @LongitudFin where nro_movil = @NroMovil  and fecha_hora_inicio = @FechaHoraInicio", DBConn)

                SqlUpdate1.Parameters.AddWithNullableValue("@NroMovil", NroMovil)
                SqlUpdate1.Parameters.AddWithNullableValue("@NroSecuenciaTurno", ProxNroSec)
                SqlUpdate1.Parameters.AddWithNullableValue("@FechaHoraInicio", FechaHoraFin)
                SqlUpdate1.Parameters.AddWithNullableValue("@FechaHoraFin", FechaHoraFin)
                SqlUpdate1.Parameters.AddWithNullableValue("@NroConductor", NroConductor)
                SqlUpdate1.Parameters.AddWithNullableValue("@CantExcesosVel", CantExcesosVel)
                SqlUpdate1.Parameters.AddWithNullableValue("@KmRecOcupado", KmRecOcupado)
                SqlUpdate1.Parameters.AddWithNullableValue("@KmRecTotales", KmRecTotales)
                SqlUpdate1.Parameters.AddWithNullableValue("@CantTicketsEmitidos", CantTicketsEmitidos)
                SqlUpdate1.Parameters.AddWithNullableValue("@CantDescReloj", CantDescReloj)
                SqlUpdate1.Parameters.AddWithNullableValue("@CantViajesPorSensor", CantViajesPorSensor)
                SqlUpdate1.Parameters.AddWithNullableValue("@CantViajesT1", CantViajesT1)
                SqlUpdate1.Parameters.AddWithNullableValue("@CantRecaudT1", CantRecaudT1)
                SqlUpdate1.Parameters.AddWithNullableValue("@CantViajesT2", CantViajesT2)
                SqlUpdate1.Parameters.AddWithNullableValue("@CantRecaudT2", CantRecaudT2)
                SqlUpdate1.Parameters.AddWithNullableValue("@CantViajesT3", CantViajesT3)
                SqlUpdate1.Parameters.AddWithNullableValue("@CantRecaudT3", CantRecaudT3)
                SqlUpdate1.Parameters.AddWithNullableValue("@CantViajesT4", CantViajesT4)
                SqlUpdate1.Parameters.AddWithNullableValue("@CantRecaudT4", CantRecaudT4)
                SqlUpdate1.Parameters.AddWithNullableValue("@NroCorrelativoViaje", NroCorrelativoViaje)
                SqlUpdate1.Parameters.AddWithNullableValue("@IndFinTurnoOK", True)
                SqlUpdate1.Parameters.AddWithNullableValue("@LatitudFin", LatitudFin)
                SqlUpdate1.Parameters.AddWithNullableValue("@LongitudFin", LongitudFin)

                SqlUpdate1.CommandType = CommandType.Text

                SqlUpdate1.ExecuteNonQuery()



                Dim SqlUpdate2 As New SqlCommand("update historia_turnos_temp set fecha_hora_fin = @FechaHoraFin , cant_excesos_vel = @CantExcesosVel, km_rec_ocupado = @KmRecOcupado, km_rec_totales = @KmRecTotales, cant_tickets_emitidos = @CantTicketsEmitidos, cant_desc_reloj = @CantDescReloj, cant_viajes_por_sensor = @CantViajesPorSensor, cant_viajes_t1 = @CantViajesT1, cant_viajes_t2 = @CantViajesT2, cant_viajes_t3 = @CantViajesT3, cant_viajes_t4 = @CantViajes4, cant_recaud_t1 = @CantRecaudT1, cant_recaud_t2 = @CantRecaudT2, cant_recaud_t3= @CantRecaudT3, cant_recaud_t4 = @CantRecaudT4, nro_correlativo_viaje = @NroCorrelativoViaje, ind_fin_turno_ok = @IndFinTurnoOK, latitud_fin = @LatitudFin , longitud_fin = @LongitudFin where nro_movil = @NroMovil  and fecha_hora_inicio = @FechaHoraInicio", DBConn)

                SqlUpdate2.Parameters.AddWithNullableValue("@NroMovil", NroMovil)
                SqlUpdate2.Parameters.AddWithNullableValue("@NroSecuenciaTurno", ProxNroSec)
                SqlUpdate2.Parameters.AddWithNullableValue("@FechaHoraInicio", FechaHoraFin)
                SqlUpdate2.Parameters.AddWithNullableValue("@FechaHoraFin", FechaHoraFin)
                SqlUpdate2.Parameters.AddWithNullableValue("@NroConductor", NroConductor)
                SqlUpdate2.Parameters.AddWithNullableValue("@CantExcesosVel", CantExcesosVel)
                SqlUpdate2.Parameters.AddWithNullableValue("@KmRecOcupado", KmRecOcupado)
                SqlUpdate2.Parameters.AddWithNullableValue("@KmRecTotales", KmRecTotales)
                SqlUpdate2.Parameters.AddWithNullableValue("@CantTicketsEmitidos", CantTicketsEmitidos)
                SqlUpdate2.Parameters.AddWithNullableValue("@CantDescReloj", CantDescReloj)
                SqlUpdate2.Parameters.AddWithNullableValue("@CantViajesPorSensor", CantViajesPorSensor)
                SqlUpdate2.Parameters.AddWithNullableValue("@CantViajesT1", CantViajesT1)
                SqlUpdate2.Parameters.AddWithNullableValue("@CantRecaudT1", CantRecaudT1)
                SqlUpdate2.Parameters.AddWithNullableValue("@CantViajesT2", CantViajesT2)
                SqlUpdate2.Parameters.AddWithNullableValue("@CantRecaudT2", CantRecaudT2)
                SqlUpdate2.Parameters.AddWithNullableValue("@CantViajesT3", CantViajesT3)
                SqlUpdate2.Parameters.AddWithNullableValue("@CantRecaudT3", CantRecaudT3)
                SqlUpdate2.Parameters.AddWithNullableValue("@CantViajesT4", CantViajesT4)
                SqlUpdate2.Parameters.AddWithNullableValue("@CantRecaudT4", CantRecaudT4)
                SqlUpdate2.Parameters.AddWithNullableValue("@NroCorrelativoViaje", NroCorrelativoViaje)
                SqlUpdate2.Parameters.AddWithNullableValue("@IndFinTurnoOK", True)
                SqlUpdate2.Parameters.AddWithNullableValue("@LatitudFin", LatitudFin)
                SqlUpdate2.Parameters.AddWithNullableValue("@LongitudFin", LongitudFin)

                SqlUpdate2.CommandType = CommandType.Text

                SqlUpdate2.ExecuteNonQuery()

            End If

            Regs.Close()

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
