Imports System.Data.SqlClient

Public Module Zona
    Public ConnString As String
    Public Status As Integer
    Public Zonas As List(Of DatosZona)
    Public CodHorarioActual As Integer
    
    Public Class CentroZona
        Public Latitud As Single
        Public Longitud As Single
    End Class

    Public Class DatosZona
        Public LatitudCentro As Single
        Public LongitudCentro As Single
        Public AdmiteFuturos As Boolean
        Public Radios(4, 5) As Single
    End Class

    Public Function ObtenerMaximaCantidadZonas() As Integer

        Dim Valor As Object

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("select max(cod_zona) as maxzona from zonas where cod_zona <= 60", DBConn)
            SqlCmd.CommandType = CommandType.Text

            Valor = SqlCmd.ExecuteScalar

            If IsNothing(Valor) Then
                Valor = 0
            End If

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            Valor = Nothing
            Status = 1
        End Try

        Return Valor

    End Function

    Public Function ObtenerZona(ByVal Latitud As Single, ByVal Longitud As Single) As Integer

        Dim Valor As Integer

        Try

            If Latitud <> 0 And Longitud <> 0 Then

                Dim DBConn = New SqlConnection(ConnString)
                DBConn.Open()

                Dim SqlCmd As New SqlCommand("Zona_Obtener1", DBConn)
                SqlCmd.CommandType = CommandType.StoredProcedure

                SqlCmd.Parameters.AddWithNullableValue("@Latitud", Latitud)
                SqlCmd.Parameters.AddWithNullableValue("@Longitud", Longitud)

                SqlCmd.Parameters.Add("ReturnValue", SqlDbType.Int)
                SqlCmd.Parameters("ReturnValue").Direction = ParameterDirection.ReturnValue

                SqlCmd.Parameters.Add("@Zona", SqlDbType.Int)
                SqlCmd.Parameters("@Zona").Direction = ParameterDirection.Output

                SqlCmd.ExecuteScalar()

                Valor = SqlCmd.Parameters("@Zona").Value

                If IsNothing(Valor) Then
                    Valor = 0
                End If

                If DBConn.State = ConnectionState.Open Then
                    DBConn.Close()
                End If

                DBConn.Dispose()

            Else
                Valor = 0
            End If

            Status = 0

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            Valor = 0
            Status = 1
        End Try

        Return Valor

    End Function

    Public Function ObtenerCoordenadasCentro(ByVal CodZona As Integer) As CentroZona

        Dim Centro = New CentroZona

        Centro.Latitud = 0
        Centro.Longitud = 0

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("Select latitud_centro, longitud_centro from Zonas where IdZona = @CodZona", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@CodZona", CodZona)
            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            Do While Regs.Read
                Centro.Latitud = Regs("latitud_centro")
                Centro.Longitud = Regs("longitud_centro")
                Exit Do
            Loop

            Regs.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            Status = 1
        End Try

        Return Centro

    End Function

    Public Function CalcularDemora(ByVal TipoServicio As Integer, ByVal CodZonaOrigen As Integer, ByVal CodZonaDestino As Integer) As Integer

        Dim Demora As Integer = 0

        Try

            Dim CodHorario = Parametro.Leer("cod_horario_actual")

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("select tiempo_promedio from promedios where tipo_servicio = @TipoServicio and cod_horario = @CodHorario  and cod_zona_origen = @CodZonaOrigen and cod_zona_destino = @CodZonaDestino", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@TipoServicio", TipoServicio)
            SqlCmd.Parameters.AddWithNullableValue("@CodHorario", CodHorario)
            SqlCmd.Parameters.AddWithNullableValue("@CodZonaOrigen", CodZonaOrigen)
            SqlCmd.Parameters.AddWithNullableValue("@CodZonaDestino", CodZonaDestino)

            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            If Regs.HasRows Then
                Regs.Read()
                Demora = Regs("tiempo_promedio")
            Else
                Demora = 60
            End If

            Regs.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            Status = 1
            Demora = 0
        End Try

        Return Demora

    End Function

    Public Sub ActulizarPromedios(ByVal CodZonaOrigen As Integer, ByVal CodZonaDestino As Integer, ByVal TipoServicio As Integer, ByVal FechaInicioViaje As Date, ByVal FechaFinViaje As Date)

        Dim Demora As Integer
        Dim DemoraMinima As Integer
        Dim DemoraMaxima As Integer
        Dim DemoraPromedio As Integer
        Dim CantCasos As Integer
        Dim SumatoriaTiempos As Integer

        Try

            Call ActualizarDatosHorario()

            Demora = DateDiff("n", FechaInicioViaje, FechaFinViaje)

            If Demora > 60 Or Demora < 0 Then
                Exit Sub
            End If

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("select * from promedios where tipo_servicio = @TipoServicio and cod_horario = @CodHorarioActual  and cod_zona_origen = @CodZonaOrigen  and cod_zona_destino = @CodZonaDestino", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@TipoServicio", TipoServicio)
            SqlCmd.Parameters.AddWithNullableValue("@CodHOrarioActual", CodHorarioActual)
            SqlCmd.Parameters.AddWithNullableValue("@CodZonaOrigen", CodZonaOrigen)
            SqlCmd.Parameters.AddWithNullableValue("@CodZonaDestino", CodZonaDestino)

            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            If Regs.HasRows Then

                Regs.Read()

                CantCasos = Regs("cant_casos") + 1

                SumatoriaTiempos = Regs("sumatoria_tiempos") + Demora

                If Regs("tiempo_minimo") = 0 Then
                    DemoraMinima = Demora
                Else
                    If Demora < Regs("tiempo_minimo") Then
                        DemoraMinima = Demora
                    Else
                        DemoraMinima = Regs("tiempo_minimo")
                    End If
                End If

                If Demora > Regs("tiempo_maximo") Then
                    DemoraMaxima = Demora
                Else
                    DemoraMaxima = Regs("tiempo_maximo")
                End If

                If CantCasos > 0 Then
                    DemoraPromedio = SumatoriaTiempos / CantCasos
                Else
                    DemoraPromedio = Regs("tiempo_promedio")
                End If

                Dim SqlCmdUpdate As New SqlCommand("update promedios set tiempo_minimo = @DemoraMinima, tiempo_maximo = @DemoraMaxima, tiempo_promedio = @DemoraPromedio, sumatoria_tiempos =  @SumatoriaTiempos, cant_casos = @CantCasos  where tipo_servicio = @TipoServicio  and cod_horario = @CodHorarioActual  and cod_zona_origen = @CodZonaOrigen  and cod_zona_destino = @CodZonaDestino", DBConn)

                SqlCmdUpdate.Parameters.AddWithNullableValue("@DemoraMinima", DemoraMinima)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@DemoraMaxima", DemoraMaxima)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@DemoraPromedio", DemoraPromedio)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@SumatoriaTiempos", SumatoriaTiempos)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@CantCasos", CantCasos)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@TipoServicio", TipoServicio)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@CodHOrarioActual", CodHorarioActual)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@CodZonaOrigen", CodZonaOrigen)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@CodZonaDestino", CodZonaDestino)

                SqlCmdUpdate.CommandType = CommandType.Text
                SqlCmdUpdate.ExecuteNonQuery()

            Else

                Dim SqlCmdInsert As New SqlCommand("insert into promedios (tipo_servicio, cod_horario, cod_zona_origen, cod_zona_destino, cant_casos, sumatoria_tiempos, tiempo_maximo, tiempo_minimo, tiempo_promedio) values (@TipoServicio, @CodHorarioActual, @CodZonaOrigen, @CodZonaDestino, 1, @Demora, @Demora, @Demora, @Demora)", DBConn)

                SqlCmdInsert.Parameters.AddWithNullableValue("@Demora", Demora)
                SqlCmdInsert.Parameters.AddWithNullableValue("@TipoServicio", TipoServicio)
                SqlCmdInsert.Parameters.AddWithNullableValue("@CodHorarioActual", CodHorarioActual)
                SqlCmdInsert.Parameters.AddWithNullableValue("@CodZonaOrigen", CodZonaOrigen)
                SqlCmdInsert.Parameters.AddWithNullableValue("@CodZonaDestino", CodZonaDestino)

                SqlCmdInsert.CommandType = CommandType.Text
                SqlCmdInsert.ExecuteNonQuery()

            End If

            Regs.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            Status = 1
        End Try


    End Sub

    Public Sub ActualizarDatosHorario()

        Dim CodHorario = Parametro.Leer("cod_horario_actual")

        If CodHorarioActual <> CodHorario Then
            CodHorarioActual = CodHorario
        End If

    End Sub

    Public Function CargarCentrosZonas() As List(Of DatosZona)

        Dim Zonas As New List(Of DatosZona)

        Zonas.Clear()

        Try

            Call ActualizarDatosHorario()

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd1 As New SqlCommand("select cod_zona, latitud_centro, longitud_centro, ind_viajes_futuros from zonas", DBConn)
            SqlCmd1.CommandType = CommandType.Text

            Dim RegZonas = SqlCmd1.ExecuteReader()

            Do While RegZonas.Read

                Dim NuevaZona = New DatosZona

                NuevaZona.LatitudCentro = RegZonas("Latitud_centro")
                NuevaZona.LongitudCentro = RegZonas("Longitud_centro")
                NuevaZona.AdmiteFuturos = RegZonas("ind_viajes_futuros")

                Dim SqlCmd2 As New SqlCommand("select * from zonas_horarios where cod_zona = @Zona and cod_horario = @CodHorarioActual", DBConn)
                SqlCmd2.Parameters.AddWithNullableValue("@CodHorarioActual", CodHorarioActual)
                SqlCmd2.Parameters.AddWithNullableValue("@Zona", RegZonas("cod_zona"))

                SqlCmd2.CommandType = CommandType.Text

                Dim Horarios = SqlCmd2.ExecuteReader

                Do While Horarios.Read
                    NuevaZona.Radios(Horarios("tipo_servicio"), 0) = Horarios("radio_zona")
                    NuevaZona.Radios(Horarios("tipo_servicio"), 1) = Horarios("tiempo_anticip_lanzamiento_age")
                    NuevaZona.Radios(Horarios("tipo_servicio"), 3) = Horarios("tiempo_anticip_invitacion_agen")
                    NuevaZona.Radios(Horarios("tipo_servicio"), 4) = Horarios("RadioBusqBase")
                    NuevaZona.Radios(Horarios("tipo_servicio"), 5) = Horarios("TiempoAnticFuturo")
                Loop

                Horarios.Close()

                'Carga los radios de zona

                Zonas.Add(NuevaZona)
            Loop

            RegZonas.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()
            Status = 0

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            Status = 1
        End Try

        Return Zonas


    End Function

    Public Sub ActualizaHorarioSistema()

        Try

            Dim CodHorario = Parametro.Leer("cod_horario_actual")

            Dim FechaActual = Now
            Dim HoraActual = TimeOfDay

            Dim intdiasemana = Weekday(FechaActual)
            Dim bolEsHabil As Boolean

            If intdiasemana >= 2 And intdiasemana <= 6 Then
                bolEsHabil = EsHabil(FechaActual)
            Else
                bolEsHabil = False
            End If

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("select * from Horarios_Sistema", DBConn)

            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            Do While Regs.Read

                If (Regs("ind_domingo") = True And intdiasemana = 1) Or _
                    (Regs("ind_Lunes") = True And intdiasemana = 2) Or _
                    (Regs("ind_Martes") = True And intdiasemana = 3) Or _
                    (Regs("ind_Miercoles") = True And intdiasemana = 4) Or _
                    (Regs("ind_Jueves") = True And intdiasemana = 5) Or _
                    (Regs("ind_Viernes") = True And intdiasemana = 6) Or _
                    (Regs("ind_Sabado") = True And intdiasemana = 7) And _
                    (Regs("ind_habil") = True And bolEsHabil = True Or _
                    Regs("ind_feriado") = True And bolEsHabil = False) And _
                    (Regs("hora_desde") <= HoraActual And Regs("hora_hasta") >= HoraActual) Then

                    'Cambia el horario del sistema

                    Dim SqlCmd1 As New SqlCommand("update config_sis set cod_horario_actual = @CodHorarioActual Where cod_registro = 1 ", DBConn)

                    SqlCmd1.CommandType = CommandType.Text
                    SqlCmd1.Parameters.AddWithNullableValue("@CodHorarioActual", Regs("cod_horario"))

                    SqlCmd1.ExecuteNonQuery()

                End If
            Loop

            Regs.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            Status = 1
        End Try

    End Sub

    Public Function ObtieneRadioBusquedaZona(ByVal IdZona As Integer, ByVal TipoServicio As Integer) As Long

        Dim Radio As Integer = 0

        Try

            'Dim CodHorario = Parametro.Leer("cod_horario_actual")

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("Select Radio_Zona from Zonas_Horarios where Tipo_Servicio = @TipoServicio and cod_zona = @IdZona", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@TipoServicio", TipoServicio)
            SqlCmd.Parameters.AddWithNullableValue("@IdZona", IdZona)
        
            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            If Regs.HasRows Then
                Regs.Read()
                Radio = Regs("Radio_Zona") * 10
            Else
                Radio = 2000
            End If

            Regs.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            Status = 1
            Radio = 2000
        End Try

        Return Radio

    End Function

    Public Function ObtieneRadioBusquedaBase(ByVal IdZona As Integer, ByVal TipoServicio As Integer) As Long

        Dim Radio As Integer = 0

        Try

            'Dim CodHorario = Parametro.Leer("cod_horario_actual")

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("select RadioBusqBase from zonas_horarios where Tipo_Servicio = @TipoServicio and Cod_Zona = @IdZona", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@TipoServicio", TipoServicio)
            SqlCmd.Parameters.AddWithNullableValue("@IdZona", IdZona)

            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            If Regs.HasRows Then
                Regs.Read()
                Radio = Regs("RadioBusqBase")
            Else
                Radio = 1000
            End If

            Regs.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            Status = 1
            Radio = 1000
        End Try

        Return Radio

    End Function

    Public Function ObtieneTiempoBusquedaFuturo(ByVal IdZona As Integer, ByVal TipoServicio As Integer) As Long

        Dim Radio As Integer = 0

        Try

            'Dim CodHorario = Parametro.Leer("cod_horario_actual")

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("Select TiempoAnticFuturo from Zonas_Horarios where Tipo_Servicio = @TipoServicio and Cod_Zona = @IdZona", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@TipoServicio", TipoServicio)
            SqlCmd.Parameters.AddWithNullableValue("@IdZona", IdZona)

            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            If Regs.HasRows Then
                Regs.Read()
                Radio = Regs("TiempoAnticFuturo")
            Else
                Radio = 2
            End If

            Regs.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            Status = 1
            Radio = 2
        End Try

        Return Radio

    End Function


End Module