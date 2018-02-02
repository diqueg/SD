Imports System.Data.SqlClient

Public Module Movil

    Public ConnString As String
    Public Status As Integer

    Public Class DatosMovil
        Public NroMovil As Int16
        Public NroInternoMovil As Int16?
        Public NroTitular As Int16?
        Public Modelo As String
        Public Version As String
        Public Categoria As String
        Public NroMotor As String
        Public NroChasis As String
        Public NroPatente As String
        Public NroOficRNPA As String
        Public FecVtoCedulaVerde As Date?
        Public NroFlota As Int16?
        Public TipoServicio As Int16?
        Public EstadoAdmin As String
        Public NroConductor As Int16?
        Public ClaveConductor As Integer?
        Public NroCanal As Int16?
        Public CodZonaActual As Int16?
        Public CodZonaFutura As Int16?
        Public EstadoOper As String
        Public FechaEstadoOper As Date?
        Public EstadoFuturo As String
        Public FechaEstadoFuturo As Date?
        Public EstadoCom As String
        Public FechaUltContacto As Date?
        Public EstadoReloj As String
        Public EstadoAlarma As String
        Public CantViajes As Integer?
        Public ViajesNormales As Boolean?
        Public ViajesCorta As Boolean?
        Public ViajesMedia As Boolean?
        Public ViajesLarga As Boolean?
        Public EmpresaSeguro As String
        Public NroPoliza As String
        Public FechaInicioPoliza As Date?
        Public FechaFinPoliza As Date?
        Public HistoriaIncidentes As String
        Public SaldoAnterior As Decimal?
        Public FechaSaldoAnterior As Decimal?
        Public SaldoActual As Decimal
        Public NroViajeActual As Integer?
        Public NroViajeFuturo As Integer?
        Public FechaLiberacion As Date?
        Public HoraInicioTurno As Date?
        Public HoraFinTurno As Date?
        Public Latitud As Single?
        Public Longitud As Single?
        Public Velocidad As Int16?
        Public Rumbo As String
        Public FechaHoraPosicion As Date?
        Public CantMinutosSector0 As Integer?
        Public LugarLiq As Int16?
        Public CodParada As Int16?
        Public FechaIngresoParada As Date?
        Public NroVisorTDM As Integer?
    End Class


    Public Function ObtenerLibresPorZona(ByVal NroCanal As Integer) As Dictionary(Of Integer, Integer)

        Dim Result = New Dictionary(Of Integer, Integer)


        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("select cod_zona_actual as cod_zona, count(nro_movil) as cantidad from moviles where estado_operativo = 'L' and estado_administrativo = 'H' and cod_zona_actual > 0 and nro_canal_comunicaciones = @NroCanal group by cod_zona_actual order by cod_zona_actual", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@NroCanal", NroCanal)

            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            While Regs.Read

                Dim CantMoviles As Integer

                If Regs("cantidad") < 10 Then
                    CantMoviles = 10
                Else
                    CantMoviles = Regs("cantidad")
                End If

                Result.Add(Regs("cod_zona"), CantMoviles)

            End While

            Regs.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            Result = Nothing
            Status = 1
        End Try

        Return Result

    End Function


    Public Sub ConsultarEstadoYZonaDemorados(ByVal NroCanal As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("select * from moviles where (estado_operativo <> 'F' and estado_operativo <> 'S') and nro_canal_comunicaciones = @NroCanal", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@NroCanal", NroCanal)

            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            While Regs.Read

                If Regs("estado_operativo") = "D" Or Regs("estado_operativo") = "Y" Then

                    'consulta el movil el estado y zona

                    Comando.EnviarComando(NroCanal, Regs("nro_movil"), "17", "", False, 3, 15)

                End If

            End While

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

    Public Sub ResolverPenalizados(ByVal NroCanal As Integer, ByVal TiempoEsperaRespuestaInvitacion As Integer, ByVal TiempoPenalizacion As Integer)

        Try
            Dim DatMovil As New DatosMovil

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("Select * from moviles where (estado_operativo = 'D' or estado_futuro = 'D' or estado_operativo = 'I' or estado_futuro = 'I' )  and nro_canal_comunicaciones = @NroCanal", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@NroCanal", NroCanal)

            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            While Regs.Read

                Dim TiempoPenalizado As Integer
                Dim MaxEspera As Integer

                If Regs("estado_operativo") = "I" Then
                    TiempoPenalizado = DateDiff("s", CType(Regs("fecha_hora_estado_operativo"), Date), Now)
                    MaxEspera = TiempoEsperaRespuestaInvitacion
                End If

                If Regs("estado_futuro") = "I" Then
                    TiempoPenalizado = DateDiff("s", CType(Regs("fecha_hora_estado_futuro"), Date), Now)
                    MaxEspera = TiempoEsperaRespuestaInvitacion
                End If

                If Regs("estado_operativo") = "D" Then
                    TiempoPenalizado = DateDiff("s", CType(Regs("fecha_hora_estado_operativo"), Date), Now)
                    MaxEspera = TiempoPenalizacion
                End If

                If Regs("estado_futuro") = "D" Then
                    TiempoPenalizado = DateDiff("s", CType(Regs("fecha_hora_estado_futuro"), Date), Now)
                    MaxEspera = TiempoPenalizacion
                End If

                DatMovil = ObtenerDatos(Regs("nro_movil"))

                'Si se venció el tiempo de penalización...
                If TiempoPenalizado > MaxEspera Then

                    'Si está en consulta el viaje actual...

                    If Regs("estado_operativo") = "I" Then
                        Movil.DemorarActual(Regs("nro_movil"))
                        Viaje.Redespachar(DatMovil.NroViajeActual)
                    End If

                    'Si está en consulta el viaje futuro...

                    If Regs("estado_futuro") = "I" Then
                        Movil.DemorarFuturo(Regs("nro_movil"))
                    End If

                    'Si está penalizado el estado actual...

                    If Regs("estado_operativo") <> "D" Then
                        Movil.DespenalizarEstadoFuturo(Regs("nro_movil"))
                    Else
                        If TiempoPenalizado <= TiempoEsperaDespenalizacion * 60 Then
                            If EsPar(Fix(TiempoPenalizado / 10)) Then
                                Comando.EnviarComando(NroCanal, Regs("nro_movil"), "17", "", False, 3, 15)
                            Else
                                'Pone al movil Fuera de Servicio
                                Movil.SacarDeServicio(Regs("nro_movil"))
                            End If
                        End If

                    End If
                End If
            End While

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

    Public Sub DemorarActual(ByVal NroMovil As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("Update Moviles set estado_operativo = 'D' , fecha_hora_estado_operativo = @Now, nro_viaje_actual = 0  Where nro_movil = @NroMovil", DBConn)
            SqlCmd.Parameters.AddWithNullableValue("@Now", Now)
            SqlCmd.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmd.CommandType = CommandType.Text
            Dim Regs = SqlCmd.ExecuteNonQuery

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


    Public Sub DemorarFuturo(ByVal NroMovil As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("Update Moviles set estado_futuro = 'D' , fecha_hora_estado_futuro = getdate() , nro_viaje_futuro = 0  Where nro_movil = @NroMovil", DBConn)
            SqlCmd.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmd.CommandType = CommandType.Text
            Dim Regs = SqlCmd.ExecuteNonQuery

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


    Public Sub DespenalizarEstadoFuturo(ByVal NroMovil As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("Update Moviles set estado_futuro = 'L' , fecha_hora_estado_futuro = getdate(), nro_viaje_futuro = 0  Where nro_movil = @NroMovil", DBConn)
            SqlCmd.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmd.CommandType = CommandType.Text
            Dim Regs = SqlCmd.ExecuteNonQuery

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

    Public Function ObtenerDatos(ByVal NroMovil As Integer) As DatosMovil

        Dim DatosMovil As New DatosMovil

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("Select * from moviles where nro_movil = @NroMovil", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            If Regs.HasRows Then

                Regs.Read()

                DatosMovil.NroMovil = IsNull(Regs("nro_movil"))
                DatosMovil.NroInternoMovil = IsNull(Regs("nro_interno_movil"))
                DatosMovil.NroTitular = IsNull(Regs("nro_titular"))
                DatosMovil.Modelo = IsNull(Regs("modelo"))
                DatosMovil.Version = IsNull(Regs("Version"))
                DatosMovil.Categoria = IsNull(Regs("categ_movil"))
                DatosMovil.NroMotor = IsNull(Regs("nro_motor"))
                DatosMovil.NroChasis = IsNull(Regs("nro_chasis"))
                DatosMovil.NroPatente = IsNull(Regs("nro_patente"))
                DatosMovil.NroOficRNPA = IsNull(Regs("nro_oficina_rnpa"))
                DatosMovil.FecVtoCedulaVerde = IsNull(Regs("fecha_vto_cedula_verde"))
                DatosMovil.NroFlota = IsNull(Regs("nro_flota"))
                DatosMovil.TipoServicio = IsNull(Regs("tipo_servicio"))
                DatosMovil.EstadoAdmin = IsNull(Regs("estado_administrativo"))
                DatosMovil.NroConductor = IsNull(Regs("nro_conductor"))
                DatosMovil.NroCanal = IsNull(Regs("nro_canal_comunicaciones"))
                DatosMovil.CodZonaActual = IsNull(Regs("cod_zona_actual"))
                DatosMovil.CodZonaFutura = IsNull(Regs("cod_zona_futura"))
                DatosMovil.EstadoOper = IsNull(Regs("estado_operativo"))
                DatosMovil.FechaEstadoOper = IsNull(Regs("fecha_hora_estado_operativo"))
                DatosMovil.EstadoFuturo = IsNull(Regs("estado_futuro"))
                DatosMovil.FechaEstadoFuturo = IsNull(Regs("fecha_hora_estado_futuro"))
                DatosMovil.EstadoCom = IsNull(Regs("estado_comunicacion"))
                DatosMovil.FechaUltContacto = IsNull(Regs("fecha_hora_ultimo_contacto"))
                DatosMovil.EstadoReloj = IsNull(Regs("estado_reloj"))
                DatosMovil.EstadoAlarma = IsNull(Regs("estado_alarma"))
                DatosMovil.CantViajes = IsNull(Regs("cant_viajes_realizados"))
                DatosMovil.ViajesNormales = CType(IsNull(Regs("ind_acepta_viajes_normales")), Boolean)
                DatosMovil.ViajesCorta = CType(IsNull(Regs("ind_acepta_viajes_corta")), Boolean)
                DatosMovil.ViajesMedia = CType(IsNull(Regs("ind_acepta_viajes_media")), Boolean)
                DatosMovil.ViajesLarga = CType(IsNull(Regs("ind_acepta_viajes_larga")), Boolean)
                DatosMovil.EmpresaSeguro = IsNull(Regs("empresa_poliza_seguro"))
                'DatosMovil.NroPoliza = Isnull(Regs("nro_poliza_seguro"))
                'DatosMovil.FechaInicioPoliza = Isnull(Regs("fec_ini_poliza_seguro"))
                'DatosMovil.FechaFinPoliza = Isnull(Regs("fec_fin_poliza_seguro"))
                'DatosMovil.HistoriaIncidentes = Isnull(Regs("historia_incidentes"))
                DatosMovil.SaldoAnterior = IsNull(Regs("saldo_anterior_ctacte"))
                'DatosMovil.FechaSaldoAnterior = Isnull(Regs("fecha_saldo_anterior_ctacte"))
                DatosMovil.SaldoActual = IsNull(Regs("saldo_actual_ctacte"))
                DatosMovil.NroViajeActual = IsNull(Regs("nro_viaje_actual"))
                DatosMovil.NroViajeFuturo = IsNull(Regs("nro_viaje_futuro"))
                DatosMovil.FechaLiberacion = IsNull(Regs("fecha_hora_prevista_liberacion"))
                DatosMovil.HoraInicioTurno = IsNull(Regs("hora_inicio_turno"))
                DatosMovil.HoraFinTurno = IsNull(Regs("hora_fin_turno"))
                DatosMovil.HoraFinTurno = IsNull(Regs("hora_fin_turno"))
                DatosMovil.Latitud = IsNull(Regs("latitud"))
                DatosMovil.Longitud = IsNull(Regs("longitud"))
                DatosMovil.Velocidad = IsNull(Regs("velocidad"))
                DatosMovil.Rumbo = IsNull(Regs("rumbo"))
                DatosMovil.LugarLiq = IsNull(Regs("Lugar_liq"))
                DatosMovil.CodParada = IsNull(Regs("cod_parada"))
                DatosMovil.FechaIngresoParada = IsNull(Regs("fecha_hora_entrada_parada"))
                DatosMovil.NroVisorTDM = IsNull(Regs("nro_visor_tdm"))

                Regs.Close()

                If DBConn.State = ConnectionState.Open Then
                    DBConn.Close()
                End If

                DBConn.Dispose()

                Status = 0

            End If

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            DatosMovil.NroMovil = 0
            Status = 1
        End Try

        Return DatosMovil

    End Function

    Public Sub CambiarZonaActual(ByVal NroMovil As Integer, CodZona As Integer, ByVal Posicion As Comunes.Posicion)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()


            Dim SqlCmd As New SqlCommand("Update Moviles set cod_zona_actual = @CodZona, fecha_hora_estado_operativo = getdate(), estado_comunicacion = 'C', fecha_hora_ultimo_contacto = getdate(), latitud = @Latitud, longitud = @Longitud, velocidad = @Velocidad, rumbo = @Rumbo Where nro_movil = @NroMovil", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@CodZona", CodZona)
            SqlCmd.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
            SqlCmd.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
            SqlCmd.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
            SqlCmd.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
            SqlCmd.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmd.CommandType = CommandType.Text

            Dim RegAfect = SqlCmd.ExecuteNonQuery()

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


    Public Sub PasarALibre(ByVal NroMovil As Integer, ByVal Posicion As Comunes.Posicion)

        Dim NuevaZonaActual As Integer
        Dim NuevoEstadoOPerativo As String
        Dim NuevoViajeActual As Integer

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("select * from moviles where nro_movil = @NroMovil", DBConn)
            SqlCmd.Parameters.AddWithNullableValue("@NroMovil", NroMovil)
            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            While Regs.Read

                Select Case Regs("estado_operativo")

                    Case "O", "C", "A", "I"

                        'If Regs("cod_zona_futura") <> 0 Then
                        '   NuevaZonaActual = Regs("cod_zona_futura")
                        'Else
                        '   NuevaZonaActual = Regs("cod_zona_actual")
                        'End If

                        If Regs("estado_futuro") = "A" Or Regs("estado_futuro") = "I" Then
                            NuevoEstadoOPerativo = IsNull(Regs("estado_futuro"))
                            NuevoViajeActual = IsNull(Regs("nro_viaje_futuro"))
                        Else
                            NuevoEstadoOPerativo = "L"
                            NuevoViajeActual = 0
                        End If

                        Dim SqlCmdUpdate As New SqlCommand("Update Moviles set estado_operativo = @NuevoEstadoOperativo " & _
                                  ", fecha_hora_estado_operativo = getdate()" & _
                                  ", nro_viaje_actual = @NuevoViajeActual " & _
                                  ", cod_zona_actual =  @CodZonaActual " & _
                                  ", estado_futuro = 'L' " & _
                                  ", fecha_hora_estado_futuro = getdate() " & _
                                  ", nro_viaje_futuro = 0 " & _
                                  ", cod_zona_futura =  0 " & _
                                  ", estado_comunicacion = 'C' " & _
                                  ", fecha_hora_ultimo_contacto = getdate() " & _
                                  ", estado_reloj = 'C'" & _
                                  ", latitud = @Latitud " & _
                                  ", longitud = @Longitud " & _
                                  ", velocidad =  @Velocidad " & _
                                  ", rumbo = @Rumbo " & _
                                  "  Where nro_movil = @NroMovil", DBConn)

                        SqlCmdUpdate.Parameters.AddWithNullableValue("@NuevoEstadoOperativo", NuevoEstadoOPerativo)
                        SqlCmdUpdate.Parameters.AddWithNullableValue("@NuevoViajeActual", NuevoViajeActual)
                        SqlCmdUpdate.Parameters.AddWithNullableValue("@CodZonaActual", NuevaZonaActual)
                        SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
                        SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
                        SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
                        SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
                        SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

                        SqlCmdUpdate.CommandType = CommandType.Text
                        SqlCmdUpdate.ExecuteNonQuery()


                    Case "F", "S"


                        'Si el movil está fuera de servicio y se recibe un PASE A LIBRE, corresponde
                        'grabar una entrada en el log de salidas de servicio, indicando que entró nuevamente en servicio

                        If Regs("estado_operativo") = "F" Then
                            Call EntrarEnServicio(NroMovil)
                        End If


                        Dim SqlCmdUpdate As New SqlCommand("Update Moviles set estado_operativo = 'L' " & _
                                  ", fecha_hora_estado_operativo = getdate() " & _
                                  ", nro_viaje_actual = 0 " & _
                                  ", cod_zona_actual =  @CodZonaActual " & _
                                  ", estado_futuro = 'L' " & _
                                  ", fecha_hora_estado_futuro = getdate() " & _
                                  ", nro_viaje_futuro = 0 " & _
                                  ", cod_zona_futura =  0 " & _
                                  ", estado_comunicacion = 'C' " & _
                                  ", fecha_hora_ultimo_contacto = getdate() " & _
                                  ", estado_reloj = 'C'" & _
                                  ", latitud = @Latitud " & _
                                  ", longitud = @Longitud " & _
                                  ", velocidad =  @Velocidad " & _
                                  ", rumbo = @Rumbo " & _
                                  "  Where nro_movil = @NroMovil", DBConn)

                        SqlCmdUpdate.Parameters.AddWithNullableValue("@CodZonaActual", NuevaZonaActual)
                        SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
                        SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
                        SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
                        SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
                        SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

                        SqlCmdUpdate.CommandType = CommandType.Text
                        SqlCmdUpdate.ExecuteNonQuery()

                End Select

                Exit While

            End While

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
    Public Sub PasarACobrando(ByVal NroMovil As Integer, ByVal CodZonaActual As Integer, ByVal CodZonaFutura As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("select * from moviles where nro_movil = @NroMovil", DBConn)
            SqlCmd.Parameters.AddWithNullableValue("@NroMovil", NroMovil)
            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            While Regs.Read

                If Regs("estado_operativo") = "F" Then
                    Call EntrarEnServicio(NroMovil)
                End If

                Dim SqlCmdUpdate As New SqlCommand("Update Moviles set" & _
                                            " estado_operativo = 'C'" & _
                                            ", fecha_hora_estado_operativo = GETDATE()" & _
                                            ", cod_zona_futura = @CodZonaFutura" & _
                                            ", cod_zona_actual = @CodZonaActual" & _
                                            ", cod_parada = 0" & _
                                            "  Where nro_movil = @NroMovil", DBConn)

                SqlCmdUpdate.Parameters.AddWithNullableValue("@CodZonaFutura", CodZonaFutura)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@CodZonaActual", CodZonaActual)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

                SqlCmdUpdate.CommandType = CommandType.Text
                SqlCmdUpdate.ExecuteNonQuery()


            End While

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

    Public Sub EntrarEnServicio(ByVal NroMovil As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("select * from moviles where nro_movil = @NroMovil", DBConn)
            SqlCmd.Parameters.AddWithNullableValue("@NroMovil", NroMovil)
            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            While Regs.Read

                Dim SqlCmdUpdate As New SqlCommand("insert into historia_salidas_servicio (nro_chofer, fecha_hora_evento, tipo_evento, nro_movil) values (@NroConductor, getdate(), 'E', @NroMovil)", DBConn)

                SqlCmdUpdate.Parameters.AddWithNullableValue("@NroCondutcotr", Regs("nro_conductor"))
                SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", Regs("nro_movil"))

                SqlCmdUpdate.CommandType = CommandType.Text
                SqlCmdUpdate.ExecuteNonQuery()

                Exit While

            End While

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


    Public Sub CancelarViaje(ByVal NroMovil As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()


            Dim SqlCmdUpdate As New SqlCommand("Update Moviles set estado_operativo = 'D', fecha_hora_estado_operativo = getdate(), nro_viaje_actual = 0 Where nro_movil = @NroMovil", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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

    Public Sub ActivarAlarma(ByVal NroMovil As Integer, ByVal Posicion As Comunes.Posicion)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("Update Moviles set estado_alarma = 'A' " & _
                                       ", estado_comunicacion = 'C' " & _
                                       ", fecha_hora_ultimo_contacto = getdate() " & _
                                       ", latitud = @Latitud " & _
                                       ", longitud = @Longitud " & _
                                       ", velocidad =  @Velocidad " & _
                                       ", rumbo = @Rumbo " & _
                                       "  Where nro_movil = @NroMovil", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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

    Public Sub DesconectarReloj(ByVal NroMovil As Integer, ByVal Posicion As Comunes.Posicion)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("Update Moviles set " & _
                                                "estado_operativo = 'D'" & _
                                                ", fecha_hora_estado_operativo = getdate() " & _
                                                ", nro_viaje_actual = 0" & _
                                                ", estado_futuro = 'L'" & _
                                                ", fecha_hora_estado_futuro = getdate()" & _
                                                ", estado_reloj = 'D'" & _
                                                ", fecha_hora_ultimo_contacto = getdate() " & _
                                                ", estado_comunicacion = 'C' " & _
                                                ", fecha_hora_ultimo_contacto = getdate() " & _
                                                ", latitud = @Latitud " & _
                                                ", longitud = @Longitud " & _
                                                ", velocidad =  @Velocidad " & _
                                                ", rumbo = @Rumbo " & _
                                                "  Where nro_movil = @NroMovil", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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


    Public Sub NoCambiarZonaActual(ByVal NroMovil As Integer, ByVal Posicion As Comunes.Posicion)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("Update Moviles set " & _
                                                "estado_operativo = 'L'" & _
                                                ", fecha_hora_estado_operativo = getdate() " & _
                                                ", estado_comunicacion = 'C' " & _
                                                ", fecha_hora_ultimo_contacto = getdate() " & _
                                                ", latitud = @Latitud " & _
                                                ", longitud = @Longitud " & _
                                                ", velocidad =  @Velocidad " & _
                                                ", rumbo = @Rumbo " & _
                                                "  Where nro_movil = @NroMovil", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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


    Public Sub CambiarZonaFutura(ByVal NroMovil As Integer, ByVal Posicion As Comunes.Posicion, ByVal EstadoFuturo As String, ByVal CodZonaFutura As Integer, ByVal FechaHoraLiberacion As Date)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            If EstadoFuturo = "D" Then

                Dim SqlCmdUpdate As New SqlCommand("Update Moviles set " & _
                                    "estado_futuro = 'L'" & _
                                    ", cod_zona_futura =  @CodZonaFutura " & _
                                    ", fecha_hora_estado_futuro = getdate() " & _
                                    ", fecha_hora_prevista_liberacion = @FechaHoraLiberacion " & _
                                    ", estado_comunicacion = 'C' " & _
                                    ", fecha_hora_ultimo_contacto = getdate() " & _
                                    ", latitud = @Latitud " & _
                                    ", longitud = @Longitud " & _
                                    ", velocidad =  @Velocidad " & _
                                    ", rumbo = @Rumbo " & _
                                    "  Where nro_movil = @NroMovil", DBConn)

                SqlCmdUpdate.Parameters.AddWithNullableValue("@CodZonaFutura", CodZonaFutura)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@FechaHoraLiberacion", FechaHoraLiberacion)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

                SqlCmdUpdate.CommandType = CommandType.Text
                SqlCmdUpdate.ExecuteNonQuery()


            Else

                Dim SqlCmdUpdate As New SqlCommand("Update Moviles set " & _
                                    " cod_zona_futura =  @CodZonaFutura " & _
                                    ", fecha_hora_estado_futuro = getdate() " & _
                                    ", fecha_hora_prevista_liberacion = @FechaHoraLiberacion " & _
                                    ", estado_comunicacion = 'C' " & _
                                    ", fecha_hora_ultimo_contacto = getdate() " & _
                                    ", latitud = @Latitud " & _
                                    ", longitud = @Longitud " & _
                                    ", velocidad =  @Velocidad " & _
                                    ", rumbo = @Rumbo " & _
                                    "  Where nro_movil = @NroMovil", DBConn)

                SqlCmdUpdate.Parameters.AddWithNullableValue("@CodZonaFutura", CodZonaFutura)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@FechaHoraLiberacion", FechaHoraLiberacion)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

                SqlCmdUpdate.CommandType = CommandType.Text
                SqlCmdUpdate.ExecuteNonQuery()

            End If


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

    Public Sub AceptarViaje(ByVal NroMovil As Integer, ByVal Posicion As Comunes.Posicion, ByVal EstadoOperativo As String, ByVal EstadoFuturo As String)


        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            If EstadoOperativo = "I" Then


                Dim SqlCmdUpdate As New SqlCommand("Update Moviles set " & _
                                    "estado_operativo = 'A'" & _
                                    ", fecha_hora_estado_operativo = getdate() " & _
                                    ", estado_comunicacion = 'C' " & _
                                    ", fecha_hora_ultimo_contacto = getdate() " & _
                                    ", latitud = @Latitud " & _
                                    ", longitud = @Longitud " & _
                                    ", velocidad =  @Velocidad " & _
                                    ", rumbo = @Rumbo " & _
                                    "  Where nro_movil = @NroMovil", DBConn)

                SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

                SqlCmdUpdate.CommandType = CommandType.Text
                SqlCmdUpdate.ExecuteNonQuery()


            Else

                If EstadoFuturo = "I" Then

                    Dim SqlCmdUpdate As New SqlCommand("Update Moviles set " & _
                                                            "estado_futuro = 'A'" & _
                                                            ", fecha_hora_estado_futuro = getdate() " & _
                                                            ", estado_comunicacion = 'C' " & _
                                                            ", fecha_hora_ultimo_contacto = getdate() " & _
                                                            ", latitud = @Latitud " & _
                                                            ", longitud = @Longitud " & _
                                                            ", velocidad =  @Velocidad " & _
                                                            ", rumbo = @Rumbo " & _
                                                            "  Where nro_movil = @NroMovil", DBConn)


                    SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
                    SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
                    SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
                    SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
                    SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

                    SqlCmdUpdate.CommandType = CommandType.Text
                    SqlCmdUpdate.ExecuteNonQuery()

                End If

            End If


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

    Public Sub SalirDeServicio(ByVal NroMovil As Integer, ByVal Posicion As Comunes.Posicion)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("Update Moviles set " & _
                                                "estado_operativo = 'F'" & _
                                                ", fecha_hora_estado_operativo = getdate() " & _
                                                ", estado_futuro = 'L'" & _
                                                ", fecha_hora_estado_futuro = getdate() " & _
                                                ", nro_viaje_futuro = 0 " & _
                                                ", estado_comunicacion = 'C' " & _
                                                ", fecha_hora_ultimo_contacto = getdate() " & _
                                                ", latitud = @Latitud " & _
                                                ", longitud = @Longitud " & _
                                                ", velocidad =  @Velocidad " & _
                                                ", rumbo = @Rumbo " & _
                                                "  Where nro_movil = @NroMovil", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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

    Public Sub ActualizarEstadoZona(ByVal NroMovil As Integer, ByVal Posicion As Comunes.Posicion, ByVal EstadoOperativo As String, ByVal CodZonaActual As String)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("Update Moviles set " & _
                                                "estado_operativo = @EstadoOperativo" & _
                                                ", fecha_hora_estado_operativo = getdate() " & _
                                                ", cod_zona_actual = @CodZonaActual" & _
                                                ", estado_comunicacion = 'C' " & _
                                                ", fecha_hora_ultimo_contacto = getdate() " & _
                                                ", latitud = @Latitud " & _
                                                ", longitud = @Longitud " & _
                                                ", velocidad =  @Velocidad " & _
                                                ", rumbo = @Rumbo " & _
                                                "  Where nro_movil = @NroMovil", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@EstdoOperativo", EstadoOperativo)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@CodZonaActual", CodZonaActual)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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

    Public Sub DesactivarAlarma(ByVal NroMovil As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("Update Moviles set estado_alarma = 'N' where nro_movil = @NroMovil", DBConn)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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

    Public Sub SacarDeServicio(ByVal NroMovil As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("Update Moviles set " & _
                                                "estado_operativo = 'S'" & _
                                                ", fecha_hora_estado_operativo = getdate() " & _
                                                ", nro_viaje_actual = 0 " & _
                                                ", estado_futuro = 'L'" & _
                                                ", fecha_hora_estado_futuro = getdate() " & _
                                                ", nro_viaje_futuro = 0 " & _
                                                ", estado_comunicacion = 'C' " & _
                                                ", fecha_hora_ultimo_contacto = getdate() " & _
                                                "  Where nro_movil = @NroMovil", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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

    Public Sub Inhabilitar(ByVal NroMovil As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("Update Moviles set estado_administrativo = 'I' where nro_movil = @NroMovil", DBConn)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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

    Public Sub Habilitar(ByVal NroMovil As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("Update Moviles set estado_administrativo = 'H' where nro_movil = @NroMovil", DBConn)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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

    Public Sub Incomunicar(ByVal NroMovil As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("Update Moviles set estado_comunicacion = 'I' where nro_movil = @NroMovil", DBConn)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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

    Public Sub PonerEnServicio(ByVal NroMovil As Integer, ByVal Posicion As Comunes.Posicion, ByVal EstadoFuturo As String, ByVal NroViajeFuturo As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            If EstadoFuturo = "L" Then

                Dim SqlCmdUpdate As New SqlCommand("Update Moviles set " & _
                                    "estado_operativo = 'L'" & _
                                    ", fecha_hora_estado_operativo = getdate() " & _
                                    ", nro_viaje_actual = 0 " & _
                                    ", estado_comunicacion = 'C' " & _
                                    ", fecha_hora_ultimo_contacto = getdate() " & _
                                    ", latitud = @Latitud " & _
                                    ", longitud = @Longitud " & _
                                    ", velocidad =  @Velocidad " & _
                                    ", rumbo = @Rumbo " & _
                                    "  Where nro_movil = @NroMovil", DBConn)

                SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

                SqlCmdUpdate.CommandType = CommandType.Text
                SqlCmdUpdate.ExecuteNonQuery()


            Else

                Dim SqlCmdUpdate As New SqlCommand("Update Moviles set " & _
                                                        "estado_operativo = @EstadoFuturo" & _
                                                        ", fecha_hora_estado_operativo = getdate() " & _
                                                        ", nro_viaje_actual = @NroViajeFuturo " & _
                                                        ", estado_comunicacion = 'C' " & _
                                                        ", fecha_hora_ultimo_contacto = getdate() " & _
                                                        ", latitud = @Latitud " & _
                                                        ", longitud = @Longitud " & _
                                                        ", velocidad =  @Velocidad " & _
                                                        ", rumbo = @Rumbo " & _
                                                        "  Where nro_movil = @NroMovil", DBConn)

                SqlCmdUpdate.Parameters.AddWithNullableValue("@EstadoOperativo", EstadoFuturo)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@NroViajeFuturo", NroViajeFuturo)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

                SqlCmdUpdate.CommandType = CommandType.Text
                SqlCmdUpdate.ExecuteNonQuery()

            End If


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

    Public Sub InicializarMovilesCanal(ByVal NroCanal As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("Update Moviles set" & _
                          " cod_zona_actual = 0" & _
                          ", cod_zona_futura = 0" & _
                          ", estado_operativo = 'Y'" & _
                          ", fecha_hora_estado_operativo = getdate() " & _
                          ", estado_futuro = 'L'" & _
                          ", fecha_hora_estado_futuro = getdate()" & _
                          ", estado_comunicacion = 'I'" & _
                          ", fecha_hora_ultimo_contacto = NULL" & _
                          ", estado_reloj = 'C'" & _
                          ", estado_alarma = 'N'" & _
                          ", cant_viajes_realizados = 0" & _
                          ", nro_viaje_actual = 0" & _
                          ", nro_viaje_futuro = 0" & _
                          ", latitud = 0 " & _
                          ", longitud = 0 " & _
                          ", velocidad = 0 " & _
                          ", rumbo = Null " & _
                          "  where  nro_canal_comunicaciones = @NroCanal ", DBConn)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroCanal", NroCanal)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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

    Public Sub InvitarViajeActual(ByVal NroMovil As Integer, ByVal NroViajeActual As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("Update Moviles set" & _
                                             " estado_operativo = 'I'" & _
                                             ", fecha_hora_estado_operativo = getdate() " & _
                                             ", nro_viaje_actual = @NroViajeActual " & _
                                             "  Where nro_movil =  @NroMovil ", DBConn)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroViajeActual", NroViajeActual)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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

    Public Sub InvitarViajeFuturo(ByVal NroMovil As Integer, ByVal NroViajeFuturo As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("Update Moviles set" & _
                                             " Estado_Futuro = 'I'" & _
                                             ", fecha_hora_estado_futuro = getdate() " & _
                                             ", nro_viaje_futuro = @NroViajeFuturo " & _
                                             "  Where nro_movil =  @NroMovil ", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroViajeFuturo", NroViajeFuturo)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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



    Public Sub IniciarViajeAsignado(ByVal NroMovil As Integer, ByVal CodZonaActual As Integer, ByVal Latitud As Single, ByVal Longitud As Single, ByVal Velocidad As Integer, ByVal Rumbo As String)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim DatosMovil = Movil.ObtenerDatos(NroMovil)

            If DatosMovil.EstadoOper = "F" Then
                Movil.EntrarEnServicio(NroMovil)
            End If

            Dim SqlCmdUpdate As New SqlCommand("Update Moviles set" & _
                                                " estado_operativo = 'O'" & _
                                                ", fecha_hora_prevista_liberacion = Null " & _
                                                ", fecha_hora_estado_operativo = getdate() " & _
                                                ", cant_viajes_realizados = cant_viajes_realizados + 1" & _
                                                ", estado_comunicacion = 'C'" & _
                                                ", fecha_hora_ultimo_contacto = getdate() " & _
                                                ", estado_reloj = 'C'" & _
                                                ", cod_zona_actual = @CodZonaActual " & _
                                                ", cod_parada = 0 " & _
                                                ", latitud = @Latitud " & _
                                                ", longitud = @Longitud " & _
                                                ", velocidad = @Velocidad " & _
                                                ", rumbo = @Rumbo" & _
                                                "  Where nro_movil = @NroMovil", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@CodZonaActual", CodZonaActual)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Latitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Longitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Velocidad)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Rumbo)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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

    Public Sub IniciarViajePublico(ByVal NroMovil As Integer, ByVal NroViajeActual As Integer, ByVal CodZonaActual As Integer, ByVal Latitud As Single, ByVal Longitud As Single, ByVal Velocidad As Integer, ByVal Rumbo As String)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("Update Moviles set" & _
                                                " estado_operativo = 'O'" & _
                                                ", fecha_hora_prevista_liberacion = Null " & _
                                                ", fecha_hora_estado_operativo = getdate() " & _
                                                ", cant_viajes_realizados = cant_viajes_realizados + 1" & _
                                                ", nro_viaje_actual = @NroViajeActual " & _
                                                ", estado_comunicacion = 'C'" & _
                                                ", fecha_hora_ultimo_contacto = getdate() " & _
                                                ", estado_reloj = 'C'" & _
                                                ", cod_zona_actual = @CodZonaActual " & _
                                                ", cod_parada = 0 " & _
                                                ", latitud = @Latitud " & _
                                                ", longitud = @Longitud " & _
                                                ", velocidad = @Velocidad " & _
                                                ", rumbo = @Rumbo" & _
                                                "  Where nro_movil = @NroMovil", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroViajeActual", NroViajeActual)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@CodZonaActual", CodZonaActual)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Latitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Longitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Velocidad)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Rumbo)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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

    Public Sub PTT(ByVal NroMovil As Integer, ByVal Posicion As Comunes.Posicion)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("update moviles set " & _
                                               " estado_comunicacion = 'C'" & _
                                               ", fecha_hora_ultimo_contacto = getdate() " & _
                                               ", cod_zona_actual =  @CodZonaActual " & _
                                               ", estado_reloj = 'C'" & _
                                               ", estado_comunicacion = 'C' " & _
                                               ", fecha_hora_ultimo_contacto = getdate() " & _
                                               ", latitud = @Latitud " & _
                                               ", longitud = @Longitud " & _
                                               ", velocidad =  @Velocidad " & _
                                               ", rumbo = @Rumbo " & _
                                               "  Where nro_movil =  @NroMovil ", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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

    Public Sub RtaConsultaFueraServicio(ByVal NroMovil As Integer, ByVal Posicion As Comunes.Posicion, ByVal CodZonaActual As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("update moviles set " & _
                                                " estado_operativo = 'F'" & _
                                                ", fecha_hora_estado_operativo = getdate() " & _
                                                ", cod_zona_actual =  @CodZonaActual " & _
                                                ", estado_reloj = 'C'" & _
                                                ", estado_comunicacion = 'C' " & _
                                                ", fecha_hora_ultimo_contacto = getdate() " & _
                                                ", latitud = @Latitud " & _
                                                ", longitud = @Longitud " & _
                                                ", velocidad =  @Velocidad " & _
                                                ", rumbo = @Rumbo " & _
                                                "  Where nro_movil =  @NroMovil ", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@CodZonaActual", CodZonaActual)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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

    Public Sub RtaConsultaOcupado(ByVal NroMovil As Integer, ByVal Posicion As Comunes.Posicion, ByVal CodZonaActual As Integer, ByVal NroViajeActual As Integer)


        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("update moviles set " & _
                                                " estado_operativo = 'O'" & _
                                                ", fecha_hora_estado_operativo = getdate() " & _
                                                ", nro_viaje_actual =  @NroViajeActual " & _
                                                ", cod_zona_actual =  @CodZonaActual " & _
                                                ", estado_reloj = 'C'" & _
                                                ", estado_comunicacion = 'C' " & _
                                                ", fecha_hora_ultimo_contacto = getdate() " & _
                                                ", latitud = @Latitud " & _
                                                ", longitud = @Longitud " & _
                                                ", velocidad =  @Velocidad " & _
                                                ", rumbo = @Rumbo " & _
                                                "  Where nro_movil =  @NroMovil ", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@CodZonaActual", CodZonaActual)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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

    Public Sub RtaConsultaCobrando(ByVal NroMovil As Integer, ByVal Posicion As Comunes.Posicion, ByVal CodZonaActual As Integer, ByVal NroViajeActual As Integer)


        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("update moviles set " & _
                                                " estado_operativo = 'C'" & _
                                                ", fecha_hora_estado_operativo = getdate() " & _
                                                ", cod_zona_actual =  @CodZonaActual " & _
                                                ", estado_reloj = 'C'" & _
                                                ", estado_comunicacion = 'C' " & _
                                                ", fecha_hora_ultimo_contacto = getdate() " & _
                                                ", latitud = @Latitud " & _
                                                ", longitud = @Longitud " & _
                                                ", velocidad =  @Velocidad " & _
                                                ", rumbo = @Rumbo " & _
                                                "  Where nro_movil =  @NroMovil ", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@CodZonaActual", CodZonaActual)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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

    Public Sub RtaConsultaLibre(ByVal NroMovil As Integer, ByVal Posicion As Comunes.Posicion, ByVal CodZonaActual As Integer, ByVal NroViajeActual As Integer)


        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("update moviles set " & _
                                                " estado_operativo = 'L'" & _
                                                ", fecha_hora_estado_operativo = getdate() " & _
                                                ", cod_zona_actual =  @CodZonaActual " & _
                                                 ", estado_futuro = 'L'" & _
                                                ", fecha_hora_estado_futuro = getdate() " & _
                                                ", nro_viaje_futuro =  0" & _
                                                ", estado_reloj = 'C'" & _
                                                ", estado_comunicacion = 'C' " & _
                                                ", fecha_hora_ultimo_contacto = getdate() " & _
                                                ", latitud = @Latitud " & _
                                                ", longitud = @Longitud " & _
                                                ", velocidad =  @Velocidad " & _
                                                ", rumbo = @Rumbo " & _
                                                "  Where nro_movil =  @NroMovil ", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@CodZonaActual", CodZonaActual)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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

    Public Sub RtaConsultaRelojDesc(ByVal NroMovil As Integer, ByVal Posicion As Comunes.Posicion, ByVal CodZonaActual As Integer, ByVal NroViajeActual As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("update moviles set " & _
                                                " estado_operativo = 'D'" & _
                                                ", fecha_hora_estado_operativo = getdate() " & _
                                                ", cod_zona_actual =  @CodZonaActual " & _
                                                ", nro_viaje_actual =  0" & _
                                                ", estado_reloj = 'D'" & _
                                                ", estado_futuro = 'L'" & _
                                                ", fecha_hora_estado_futuro = getdate() " & _
                                                ", nro_viaje_futuro =  0" & _
                                                ", estado_comunicacion = 'C' " & _
                                                ", fecha_hora_ultimo_contacto = getdate() " & _
                                                ", latitud = @Latitud " & _
                                                ", longitud = @Longitud " & _
                                                ", velocidad =  @Velocidad " & _
                                                ", rumbo = @Rumbo " & _
                                                "  Where nro_movil =  @NroMovil ", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@CodZonaActual", CodZonaActual)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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

    Public Sub DespenalizarEstadoFuturo(ByVal NroMovil As Integer, ByVal Posicion As Comunes.Posicion)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("update moviles set " & _
                                                ", estado_futuro = 'L'" & _
                                                ", fecha_hora_estado_futuro = getdate() " & _
                                                ", nro_viaje_futuro =  0" & _
                                                ", estado_comunicacion = 'C' " & _
                                                ", fecha_hora_ultimo_contacto = getdate() " & _
                                                ", latitud = @Latitud " & _
                                                ", longitud = @Longitud " & _
                                                ", velocidad =  @Velocidad " & _
                                                ", rumbo = @Rumbo " & _
                                                "  Where nro_movil =  @NroMovil ", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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


    Public Sub CambiarConductor(ByVal NroMovil As Integer, ByVal Posicion As Comunes.Posicion, ByVal NroConductor As Integer, ByVal CodZonaActual As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("update moviles set " & _
                                                " nro_conductor =  @NroConductor " & _
                                                ", cod_zona_actual =  @CodZonaActual " & _
                                                ", estado_comunicacion = 'C' " & _
                                                ", fecha_hora_ultimo_contacto = getdate() " & _
                                                ", latitud = @Latitud " & _
                                                ", longitud = @Longitud " & _
                                                ", velocidad =  @Velocidad " & _
                                                ", rumbo = @Rumbo " & _
                                                "  Where nro_movil =  @NroMovil ", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@CodZonaActual", CodZonaActual)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroConductor", NroConductor)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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

    Public Sub CancelarAlarma(ByVal NroMovil As Integer, ByVal Posicion As Comunes.Posicion)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("update moviles set " & _
                                                ", estado_alarma = 'N' " & _
                                                ", estado_comunicacion = 'C' " & _
                                                ", fecha_hora_ultimo_contacto = getdate() " & _
                                                ", latitud = @Latitud " & _
                                                ", longitud = @Longitud " & _
                                                ", velocidad =  @Velocidad " & _
                                                ", rumbo = @Rumbo " & _
                                                "  Where nro_movil =  @NroMovil ", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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

    Public Sub ControlarTurnos()

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("update moviles set " & _
                                                " estado_operativo = 'S'" & _
                                                ", fecha_hora_estado_operativo = getdate() " & _
                                                ", nro_viaje_actual =  0" & _
                                                ", estado_futuro = 'L'" & _
                                                ", fecha_hora_estado_futuro = getdate() " & _
                                                ", nro_viaje_futuro =  0" & _
                                                ", estado_comunicacion = 'C' " & _
                                                ", fecha_hora_ultimo_contacto = getdate() " & _
                                                " where estado_operativo = 'L' and (hora_inicio_turno > @Time  or hora_fin_turno < @Time)", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@Time", Now)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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


    Public Function IniciarTurno(ByVal NroMovil As Integer, ByVal NroConductor As Integer, ByVal ClaveConductor As String, ByVal EstadoOper As String, ByVal CodZonaActual As Integer) As Integer

        Dim Result As Integer

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("select nro_conductor, clave_conductor, liq_pend, liq_pend_max from conductores where nro_conductor = @NroConductor", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@NroConductor", NroConductor)

            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            If Regs.HasRows Then

                Regs.Read()

                If Format(ClaveConductor, "0000") = DecriptPassword(Regs("clave_conductor")) Then

                    If Regs("liq_pend") > Regs("liq_pend_max") Then
                        Result = 3240
                    Else

                        If EstadoOper = "S" Or EstadoOper = "F" Then

                            Dim SqlCmdUpdate As New SqlCommand("Update Moviles set" & _
                                                         " fecha_hora_estado_operativo = getdate()" & _
                                                         ", cod_zona_actual = @CodZonaActual " & _
                                                         ", estado_futuro = 'L' " & _
                                                         ", fecha_hora_estado_futuro = getdate() " & _
                                                         ", nro_viaje_futuro = 0" & _
                                                         ", cod_zona_futura = 0" & _
                                                         ", estado_comunicacion = 'C'" & _
                                                         ", fecha_hora_ultimo_contacto = getdate()" & _
                                                         ", estado_reloj = 'C'" & _
                                                         ", nro_conductor = @NroConductor" & _
                                                         "  Where nro_movil = @NroMovil", DBConn)

                            SqlCmdUpdate.Parameters.AddWithNullableValue("@CodZonaActual", CodZonaActual)
                            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroConductor", NroConductor)
                            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

                            SqlCmdUpdate.CommandType = CommandType.Text
                            SqlCmdUpdate.ExecuteNonQuery()

                        Else
                            Result = 1407  'el movil ya está en servicio
                        End If
                    End If
                Else
                    Result = 1406  'clave de conductor incorrecta
                End If
            Else
                Result = 1405  'nro de conductor inexistente
            End If

            Regs.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            Result = 0
            Status = 1
        End Try

        Return Result

    End Function
    Private Function EncriptedPassword(Password As String) As String

        Dim Result As String = Nothing

        Dim Encripted As String
        Dim I As Integer
        Dim PwdLen As Integer
        Dim CodeOffset As Integer
        Dim CharSep As Integer
        Dim PosOffset As Integer
        Dim NewChar As String
        Dim NewPos As Integer
        Dim Random As Integer
        Dim Control As Integer

        Try

            Randomize()

            Encripted = ""

            For I = 1 To 49
                Random = (254 * Rnd() + 1)
                Encripted = Encripted & Chr(Random)
            Next I

            PwdLen = Len(Password)
            Randomize()
            CodeOffset = (100 * Rnd() + 70)
            Randomize()
            CharSep = (5 * Rnd() + 1)
            Randomize()
            PosOffset = (10 * Rnd() + 1)


            For I = 1 To PwdLen
                NewChar = Chr(Asc(Mid(Password, I, 1)) + CodeOffset)
                NewPos = PosOffset + (CharSep * (I - 1))
                Mid(Encripted, NewPos, 1) = NewChar
            Next I

            Mid(Encripted, 46, 1) = Chr(PwdLen)
            Mid(Encripted, 47, 1) = Chr(CodeOffset)
            Mid(Encripted, 48, 1) = Chr(CharSep)
            Mid(Encripted, 49, 1) = Chr(PosOffset)

            Control = 0

            For I = 1 To 49
                Control = Control Or Asc(Mid(Encripted, I, 1))
            Next I

            Result = Encripted & Chr(Control)

            Status = 0

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            Result = Nothing
            Status = 1
        End Try

        Return Result

    End Function

    Public Sub RegistrarPosicion(ByVal NroMovil As Integer, ByVal Posicion As Comunes.Posicion, ByVal Causa As String, ByVal EstadoOperativo As String)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("insert into MOVILES_POSICIONES (nro_movil, fecha_hora, latitud, longitud,  velocidad, rumbo, estado_operativo, causa) values (@NroMovil, @FechaHoraPosicion, @Latitud,  @Longitud, @Velocidad, @Rumbo, @EstadoOper, @Causa)", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@FechaHoraPosicion", Posicion.Fecha)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@EstadoOper", EstadoOperativo)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Causa", Causa)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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

    Public Sub ActualizarPosicion(ByVal NroMovil As Integer, ByVal Posicion As Comunes.Posicion, ByVal CodZonaActual As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("update moviles set " & _
                                                " cod_zona_actual = @CodZonaActual " & _
                                                ", estado_comunicacion = 'C' " & _
                                                ", fecha_hora_ultimo_contacto = @Fecha " & _
                                                ", latitud = @Latitud " & _
                                                ", longitud = @Longitud " & _
                                                ", velocidad =  @Velocidad " & _
                                                ", rumbo = @Rumbo " & _
                                                "  Where nro_movil =  @NroMovil ", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@CodZonaActual", CodZonaActual)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Fecha", Posicion.Fecha)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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


    Public Sub PasarASector0(ByVal NroMovil As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("update moviles set " & _
                                                " estado_operativo = 'Z' " & _
                                                ", fecha_hora_estado_operativo = getdate() " & _
                                                "  Where nro_movil =  @NroMovil ", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

            Dim SqlCmd As New SqlCommand("select * from moviles where nro_movil = @NroMovil", DBConn)
            SqlCmd.Parameters.AddWithNullableValue("@NroMovil", NroMovil)
            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            While Regs.Read

                Dim SqlCmdInsert As New SqlCommand("insert into historia_salidas_servicio (nro_chofer, fecha_hora_evento, tipo_evento, nro_movil) values (@NroConductor , getdate(), 'S', @NroMovil)", DBConn)

                SqlCmdInsert.Parameters.AddWithNullableValue("@NroConductor", Regs("nro_conductor"))
                SqlCmdInsert.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

                SqlCmdInsert.CommandType = CommandType.Text
                SqlCmdInsert.ExecuteNonQuery()

                Exit While

            End While

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


    Public Sub SalirDeSector0(ByVal NroMovil As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("update moviles set " & _
                                                " estado_operativo = 'L' " & _
                                                ", fecha_hora_estado_operativo = getdate() " & _
                                                ", cant_minutos_sector_0 = 0" & _
                                                "  Where nro_movil =  @NroMovil ", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

            Dim SqlCmd As New SqlCommand("select * from moviles where nro_movil = @NroMovil", DBConn)
            SqlCmd.Parameters.AddWithNullableValue("@NroMovil", NroMovil)
            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            While Regs.Read

                Dim SqlCmdInsert As New SqlCommand("insert into historia_salidas_servicio (nro_chofer, fecha_hora_evento, tipo_evento, nro_movil) values (@NroConductor , getdate(), 'E', @NroMovil)", DBConn)

                SqlCmdInsert.Parameters.AddWithNullableValue("@NroConductor", Regs("nro_conductor"))
                SqlCmdInsert.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

                SqlCmdInsert.CommandType = CommandType.Text
                SqlCmdInsert.ExecuteNonQuery()

                Exit While

            End While

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


    Public Function ValidarPassword(ByVal NroConductor As Integer, ByVal ClaveConductor As Integer) As Integer


        Dim Result As Integer

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("select nro_conductor, clave_conductor from conductores where nro_conductor = @NroConductor", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@NroConductor", NroConductor)

            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            If Regs.HasRows Then

                Regs.Read()

                If Format(ClaveConductor, "0000") = Regs("clave_conductor") Then
                    Result = 0  'clave correcta
                Else
                    Result = 1406  'clave de conductor incorrecta
                End If
            Else
                Result = 1405  'nro de conductor inexistente
            End If

            Regs.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            Result = 0
            Status = 1
        End Try

        Return Result

    End Function

    Public Sub EntraParada(ByVal NroMovil As Integer, ByVal Posicion As Comunes.Posicion, ByVal CodParada As Integer, ByVal CodZonaActual As Integer)

        Try


            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("update moviles set " & _
                                                ", cod_parada = @CodParada " & _
                                                ", fecha_hora_entrada_parada = getdate() " & _
                                                ", estado_reloj = 'C' " & _
                                                ", fecha_hora_ultimo_contacto = getdate() " & _
                                                ", cod_zona_actual = @CodZonaActual " & _
                                                ", latitud = @Latitud " & _
                                                ", longitud = @Longitud " & _
                                                ", velocidad =  @Velocidad " & _
                                                ", rumbo = @Rumbo " & _
                                                "  Where nro_movil =  @NroMovil ", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@CodParada", CodParada)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@CodZonaActual", CodZonaActual)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Latitud", Posicion.Latitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Longitud", Posicion.Longitud)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Velocidad", Posicion.Velocidad)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Rumbo", Posicion.Rumbo)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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


    Public Sub SaleParada(ByVal NroMovil As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()


            Dim SqlCmdUpdate As New SqlCommand("update moviles set cod_parada = 0 where nro_movil = @NroMovil", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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


    Public Sub ActualizarEstadoMovilVisor(ByVal NroMovil As Integer, ByVal Estado As String)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("update Moviles set  estado_reloj = @Estado where Nro_Movil = @NroMovil", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@Estado", Estado)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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

    Public Sub ActualizarVersion(ByVal NroMovil As Integer, ByVal NroCanal As Integer, ByVal Version As String, ByVal HoraApagado As String)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlUpdate As New SqlCommand("update Moviles set VersionFirmware = @Version where Nro_Movil = @NroMovil", DBConn)

            SqlUpdate.Parameters.AddWithNullableValue("@Version", Version)
            SqlUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlUpdate.CommandType = CommandType.Text

            SqlUpdate.ExecuteNonQuery()

            If HoraApagado <> "00:00:00" Then

                Dim SqlCmd As New SqlCommand("select cod_evento from log_event_movil where nro_movil = @NroMovil and fecha_hora_evento = @HoraApagado", DBConn)

                SqlCmd.Parameters.AddWithNullableValue("@HoraApagado", HoraApagado.Replace(".", ""))
                SqlCmd.Parameters.AddWithNullableValue("@NroMovil", NroMovil)
                SqlCmd.CommandType = CommandType.Text

                Dim Regs = SqlCmd.ExecuteReader

                If Not Regs.HasRows Then

                    Dim SqlInsert1 As New SqlCommand("insert into LOG_EVENT_MOVIL (nro_movil, fecha_hora_evento, cod_evento, nro_conductor, cod_zona_actual, cod_zona_futura, nro_viaje_actual, nro_canal, latitud, longitud) values (@NroMovil, @HoraApagado, 'AP', Null , Null , Null , Null , @NroCanal , 0, 0)", DBConn)

                    SqlInsert1.Parameters.AddWithNullableValue("@HoraApagado", HoraApagado.Replace(".", ""))
                    SqlInsert1.Parameters.AddWithNullableValue("@NroMovil", NroMovil)
                    SqlInsert1.Parameters.AddWithNullableValue("@NroCanal", NroCanal)

                    SqlInsert1.CommandType = CommandType.Text

                    SqlInsert1.ExecuteNonQuery()

                    Dim SqlInsert2 As New SqlCommand("insert into LOG_EVENT_MOVIL (nro_movil, fecha_hora_evento, cod_evento, nro_conductor, cod_zona_actual, cod_zona_futura, nro_viaje_actual, nro_canal, latitud, longitud) values (@NroMovil, getdate(), 'EN', Null , Null , Null , Null , @NroCanal , 0, 0)", DBConn)

                    SqlInsert2.Parameters.AddWithNullableValue("@HoraApagado", HoraApagado.Replace(".", ""))
                    SqlInsert2.Parameters.AddWithNullableValue("@NroMovil", NroMovil)
                    SqlInsert2.Parameters.AddWithNullableValue("@NroCanal", NroCanal)

                    SqlInsert2.CommandType = CommandType.Text

                    SqlInsert2.ExecuteNonQuery()

                End If

                Regs.Close()

            End If



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


    Public Sub ActualizarUltimoContacto(ByVal NroMovil As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmdUpdate As New SqlCommand("Update Moviles set estado_comunicacion = 'C', fecha_hora_ultimo_contacto = getdate() Where nro_movil = @NroMovil", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@NroMovil", NroMovil)

            SqlCmdUpdate.CommandType = CommandType.Text
            SqlCmdUpdate.ExecuteNonQuery()

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

End Module
