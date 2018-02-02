Imports System.Data.SqlClient

Module Viaje

    Public ConnString As String
    Public Status As Integer

    Public Class DatosViaje

        Public NroViaje As Integer?
        Public NroMovilReq As Int16?
        Public NroUsuarioTelef As Integer?
        Public FechaEntrada As Date?
        Public NroCliente As Integer?
        Public CodCalle As Integer?
        Public NombreCalle As String
        Public Numero As Integer?
        Public Ubicacion As String
        Public CodLocalidad As Int16?
        Public Localidad As String
        Public Latitud As Single?
        Public Longitud As Single?
        Public CodCalle1 As Integer?
        Public NombreCalle1 As String
        Public CodCalle2 As Integer?
        Public NombreCalle2 As String
        Public NroTelefono As Decimal?
        Public Comentario As String
        Public Direccion As String
        Public CodZona As Int16?
        Public CodZonaOrigen As Int16?
        Public TipoServicio As Int16?
        Public CategoriaMovil As String
        Public TipoViaje As String
        Public EstadoRespectoCliente As String
        Public FechaInicioViaje As Date?
        Public LatitudInicioViaje As Single?
        Public LongitudInicioViaje As Single?
        Public FechaFinViaje As Date?
        Public LatitudFinViaje As Single?
        Public LongitudFinViaje As Single?
        Public CodMotivoCancel As Int16?
        Public EstadoRespectoMovil As String
        Public FechaInvitacion As Date?
        Public LatitudInvitacion As Single?
        Public LongitudInvitacion As Single?
        Public FechaResolucion As Date?
        Public LatitudResolucion As Single?
        Public LongitudResolucion As Single?
        Public NroMovilAsignado As Int16?
        Public NroFlotaAsignada As Int16?
        Public Importe As Decimal?
        Public ImportePactado As Decimal?
        Public FinAnormal As Boolean?
        Public ViajeUrgente As Boolean?
        Public NroConductor As Int16?
        Public NroCupon As Integer?
        Public CodZonaFutura As Int16?
        Public KmRecorridosLibre As Int16?
        Public VelocidadMaxLibre As Int16?
        Public NroCorrLibre As String
        Public KmRecorridosOcupado As Int16?
        Public VelocidadMaxOcupado As Int16?
        Public MinCobradosEspera As Integer?
        Public TipoTarifa As Int16?
        Public NroCorrOcupado As String
        Public Nombre As String
        Public FechaHoraRequeridaContacto As Date?
        Public NroUsuarioDespachador As Int16?
        Public Observaciones As String
        Public ImportePactadoTaxi As Decimal?
        Public CodParada As Int16?
        Public NroCtaCte As Integer?
        Public NroSubCta As Integer?

    End Class

    Public Class DatosZona
        Public NombreZona As String
        Public Demora As Integer
    End Class

    Public Class DatosDestino
        Public NroViaje As Integer
        Public NroDestino As Integer
        Public CodCalle As Long
        Public NombreCalle As String
        Public Numero As Long
        Public Ubicacion As String
        Public CodLocalidad As Integer
        Public CodZona As Integer
        Public NombreLocalidad As String
        Public Latitud As Single
        Public Longitud As Single
        Public CodCalle1 As Long
        Public NombreCalle1 As String
        Public CodCalle2 As Long
        Public NombreCalle2 As String
        Public NroTelefono As Decimal
        Public Comentario As String
        Public Direccion As String
        Public FechaHoraContacto As Date
        Public EstadoDestino As String
        Public Nombre As String
        Public Observaciones As String

    End Class

    Public Class DatosAdicional
        Public CodAdicional As Integer
        Public CantAdicional As Integer
        Public ValorAdicional As Decimal
    End Class

    Public Class MovilOptimo
        Public NroMovilOptimo As Integer
        Public TipoServicio As Integer
        Public FactorOptimo As Single
        Public Distancia As Integer
    End Class

    Public Sub BorrarViajesFactores()

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("delete from viajes_factores ", DBConn)

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

    Public Function CrearNuevo(ByVal NroMovil As Integer, ByVal ObjViaje As DatosViaje) As Integer


        Dim Result As Integer

        Try

            Dim NroViaje = UltimoNumero.ObtenerProximo("viaje")

            ObjViaje.Direccion = "<" & NroViaje.ToString & "> " & ObjViaje.Direccion

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()


            Dim SqlCmdInsert As New SqlCommand("insert into viajes (nro_viaje, nro_movil_requerido, nro_usuario_telefonista, fecha_hora_entrada, nro_cliente, cod_calle, nombre_calle, numero,  ubicacion,   cod_localidad, localidad,  latitud,  longitud,  cod_calle1, nombre_calle1, cod_calle2, nombre_calle2, nro_telefono, comentario,  direccion,  cod_zona, cod_zona_origen, tipo_servicio_requerido , categ_movil,     tipo_viaje, estado_respecto_cliente, fecha_hora_inicio_viaje, latitud_inicio_viaje, longitud_inicio_viaje, fecha_hora_fin_viaje, latitud_fin_viaje, longitud_fin_viaje, cod_motivo_cancelacion, estado_respecto_movil, fecha_hora_invitacion, latitud_invitacion, longitud_invitacion, fecha_hora_resolucion, latitud_resolucion, longitud_resolucion, nro_movil_asignado, nro_flota_asignada, importe, importe_pactado,  ind_fin_anormal, ind_viaje_urgente, nro_conductor, nro_cupon, cod_zona_futura, km_recorridos_libre, velocidad_maxima_libre, nro_correlativo_libre, km_recorridos_ocupado, velocidad_maxima_ocupado, minutos_cobrados_espera, tipo_tarifa, nro_correlativo_ocupado, nombre,  fecha_hora_requerida_contacto, nro_usuario_despachador, observaciones, Pactado_Taxi,         cod_parada, nro_ctacte,  nro_subcta) " & _
                                               "values             (@NroViaje, @NroMovilReq, @NroUsuarioTelef, @FechaEntrada, @NroCliente, @CodCalle, @NombreCalle, @Numero, @Ubicacion,  @CodLocalidad, @Localidad, @Latitud, @Longitud, @CodCalle1, @NombreCalle1, @CodCalle2, @NombreCalle2, @NroTelefono, @Comentario, @Direccion, @CodZona, @CodZonaOrigen,  @TipoServicio,            @CategoriaMovil, @TipoViaje, @EstadoRespectoCliente,  @FechaInicioViaje,       @LatitudInicioViaje,  @LongitudInicioViaje,  @FechaFinViaje,       @LatitudFinViaje,  @LongitudFinViaje,  @CodMotivoCancel,       @EstadoRespectoMovil,  @FechaInvitacion,      @LatitudInvitacion, @LongitudInvitacion, @FechaResolucion,      @LatitudResolucion, @LongitudResolucion, @NroMovilAsignado,  @NroFlotaAsignada,  @Importe, @ImportePactado, @FinAnormal,     @ViajeUrgente,     @NroConductor, @NroCupon, @CodZonaFutura,  @KmRecorridosLibre,  @VelocidadMaxLibre,     @NroCorrLibre,         @KmRecorridosOcupado,  @VelocidadMaxOcupado,     @MinCobradosEspera,      @TipoTarifa, @NroCorrOcupado,         @Nombre, @FechaHoraRequeridaContacto,   @NroUsuarioDespachador,  @Observaciones, @ImportePactadoTaxi, @CodParada, @NroCtaCte, @NroSubCta)", DBConn)

            SqlCmdInsert.Parameters.AddWithNullableValue("@NroViaje", NroViaje)
            SqlCmdInsert.Parameters.AddWithNullableValue("@NroMovilReq", ObjViaje.NroMovilReq)
            SqlCmdInsert.Parameters.AddWithNullableValue("@NroUsuarioTelef", ObjViaje.NroUsuarioTelef)
            SqlCmdInsert.Parameters.AddWithNullableValue("@FechaEntrada", ObjViaje.FechaEntrada)
            SqlCmdInsert.Parameters.AddWithNullableValue("@NroCliente", ObjViaje.NroCliente)
            SqlCmdInsert.Parameters.AddWithNullableValue("@CodCalle", ObjViaje.CodCalle)
            SqlCmdInsert.Parameters.AddWithNullableValue("@NombreCalle", ObjViaje.NombreCalle)
            SqlCmdInsert.Parameters.AddWithNullableValue("@Numero", ObjViaje.Numero)
            SqlCmdInsert.Parameters.AddWithNullableValue("@Ubicacion", ObjViaje.Ubicacion)
            SqlCmdInsert.Parameters.AddWithNullableValue("@CodLocalidad", ObjViaje.CodLocalidad)
            SqlCmdInsert.Parameters.AddWithNullableValue("@Localidad", ObjViaje.Localidad)
            SqlCmdInsert.Parameters.AddWithNullableValue("@Latitud", ObjViaje.Latitud)
            SqlCmdInsert.Parameters.AddWithNullableValue("@Longitud", ObjViaje.Longitud)
            SqlCmdInsert.Parameters.AddWithNullableValue("@CodCalle1", ObjViaje.CodCalle1)
            SqlCmdInsert.Parameters.AddWithNullableValue("@NombreCalle1", ObjViaje.NombreCalle1)
            SqlCmdInsert.Parameters.AddWithNullableValue("@CodCalle2", ObjViaje.CodCalle2)
            SqlCmdInsert.Parameters.AddWithNullableValue("@NombreCalle2", ObjViaje.NombreCalle2)
            SqlCmdInsert.Parameters.AddWithNullableValue("@NroTelefono", ObjViaje.NroTelefono)
            SqlCmdInsert.Parameters.AddWithNullableValue("@Comentario", ObjViaje.Comentario)
            SqlCmdInsert.Parameters.AddWithNullableValue("@Direccion", ObjViaje.Direccion)
            SqlCmdInsert.Parameters.AddWithNullableValue("@CodZona", ObjViaje.CodZona)
            SqlCmdInsert.Parameters.AddWithNullableValue("@CodZonaOrigen", ObjViaje.CodZonaOrigen)
            SqlCmdInsert.Parameters.AddWithNullableValue("@TipoServicio", ObjViaje.TipoServicio)
            SqlCmdInsert.Parameters.AddWithNullableValue("@CategoriaMovil", ObjViaje.CategoriaMovil)
            SqlCmdInsert.Parameters.AddWithNullableValue("@TipoViaje", ObjViaje.TipoViaje)
            SqlCmdInsert.Parameters.AddWithNullableValue("@EstadoRespectoCliente", ObjViaje.EstadoRespectoCliente)
            SqlCmdInsert.Parameters.AddWithNullableValue("@FechaInicioViaje", ObjViaje.FechaInicioViaje)
            SqlCmdInsert.Parameters.AddWithNullableValue("@LatitudInicioViaje", ObjViaje.LatitudInicioViaje)
            SqlCmdInsert.Parameters.AddWithNullableValue("@LongitudInicioViaje", ObjViaje.LongitudInicioViaje)
            SqlCmdInsert.Parameters.AddWithNullableValue("@FechaFinViaje", ObjViaje.FechaFinViaje)
            SqlCmdInsert.Parameters.AddWithNullableValue("@LatitudFinViaje", ObjViaje.LatitudFinViaje)
            SqlCmdInsert.Parameters.AddWithNullableValue("@LongitudFinViaje", ObjViaje.LongitudFinViaje)
            SqlCmdInsert.Parameters.AddWithNullableValue("@CodMotivoCancel", ObjViaje.CodMotivoCancel)
            SqlCmdInsert.Parameters.AddWithNullableValue("@EstadoRespectoMovil", ObjViaje.EstadoRespectoMovil)
            SqlCmdInsert.Parameters.AddWithNullableValue("@FechaInvitacion", ObjViaje.FechaInvitacion)
            SqlCmdInsert.Parameters.AddWithNullableValue("@LatitudInvitacion", ObjViaje.LatitudInvitacion)
            SqlCmdInsert.Parameters.AddWithNullableValue("@LongitudInvitacion", ObjViaje.LongitudInvitacion)
            SqlCmdInsert.Parameters.AddWithNullableValue("@FechaResolucion", ObjViaje.FechaResolucion)
            SqlCmdInsert.Parameters.AddWithNullableValue("@LatitudResolucion", ObjViaje.LatitudResolucion)
            SqlCmdInsert.Parameters.AddWithNullableValue("@LongitudResolucion", ObjViaje.LongitudResolucion)
            SqlCmdInsert.Parameters.AddWithNullableValue("@NroMovilAsignado", ObjViaje.NroMovilAsignado)
            SqlCmdInsert.Parameters.AddWithNullableValue("@NroFlotaAsignada", ObjViaje.NroFlotaAsignada)
            SqlCmdInsert.Parameters.AddWithNullableValue("@Importe", ObjViaje.Importe)
            SqlCmdInsert.Parameters.AddWithNullableValue("@ImportePactado", ObjViaje.ImportePactado)
            SqlCmdInsert.Parameters.AddWithNullableValue("@ImportePactadoTaxi", ObjViaje.ImportePactadoTaxi)
            SqlCmdInsert.Parameters.AddWithNullableValue("@FinAnormal", ObjViaje.FinAnormal)
            SqlCmdInsert.Parameters.AddWithNullableValue("@ViajeUrgente", ObjViaje.ViajeUrgente)
            SqlCmdInsert.Parameters.AddWithNullableValue("@NroConductor", ObjViaje.NroConductor)
            SqlCmdInsert.Parameters.AddWithNullableValue("@NroCupon", ObjViaje.NroCupon)
            SqlCmdInsert.Parameters.AddWithNullableValue("@CodZonaFutura", ObjViaje.CodZonaFutura)
            SqlCmdInsert.Parameters.AddWithNullableValue("@KmRecorridosLibre", ObjViaje.KmRecorridosLibre)
            SqlCmdInsert.Parameters.AddWithNullableValue("@VelocidadMaxLibre", ObjViaje.VelocidadMaxLibre)
            SqlCmdInsert.Parameters.AddWithNullableValue("@NroCorrLibre", ObjViaje.NroCorrLibre)
            SqlCmdInsert.Parameters.AddWithNullableValue("@KmRecorridosOcupado", ObjViaje.KmRecorridosOcupado)
            SqlCmdInsert.Parameters.AddWithNullableValue("@VelocidadMaxOcupado", ObjViaje.VelocidadMaxOcupado)
            SqlCmdInsert.Parameters.AddWithNullableValue("@MinCobradosEspera", ObjViaje.MinCobradosEspera)
            SqlCmdInsert.Parameters.AddWithNullableValue("@TipoTarifa", ObjViaje.TipoTarifa)
            SqlCmdInsert.Parameters.AddWithNullableValue("@NroCorrOcupado", ObjViaje.NroCorrOcupado)
            SqlCmdInsert.Parameters.AddWithNullableValue("@Nombre", ObjViaje.Nombre)
            SqlCmdInsert.Parameters.AddWithNullableValue("@FechaHoraRequeridaContacto", ObjViaje.FechaHoraRequeridaContacto)
            SqlCmdInsert.Parameters.AddWithNullableValue("@NroUsuarioDespachador", ObjViaje.NroUsuarioDespachador)
            SqlCmdInsert.Parameters.AddWithNullableValue("@Observaciones", ObjViaje.Observaciones)
            SqlCmdInsert.Parameters.AddWithNullableValue("@CodParada", ObjViaje.CodParada)
            SqlCmdInsert.Parameters.AddWithNullableValue("@NroCtaCte", ObjViaje.NroCtaCte)
            SqlCmdInsert.Parameters.AddWithNullableValue("@NroSubCta", ObjViaje.NroSubCta)

            SqlCmdInsert.CommandType = CommandType.Text
            SqlCmdInsert.ExecuteNonQuery()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Result = NroViaje
            Status = 0

        Catch ex As Exception
            Result = 0
            Status = 1
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try


        Return Result


    End Function

    Public Function ObtenerDatos(ByVal NroViaje As Integer) As DatosViaje

        Dim DatosViaje As New DatosViaje

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("Select * from viajes where nro_viaje = @NroViaje", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            If Regs.HasRows Then

                Regs.Read()

                DatosViaje.NroViaje = IsNull(Regs("nro_viaje"))
                DatosViaje.NroMovilReq = IsNull(Regs("nro_movil_requerido"))
                DatosViaje.NroUsuarioTelef = IsNull(Regs("nro_usuario_telefonista"))
                DatosViaje.FechaEntrada = IsNull(Regs("fecha_hora_entrada"))
                DatosViaje.NroCliente = IsNull(Regs("nro_cliente"))
                DatosViaje.CodCalle = IsNull(Regs("cod_calle"))
                DatosViaje.NombreCalle = IsNull(Regs("nombre_calle"))
                DatosViaje.Numero = IsNull(Regs("numero"))
                DatosViaje.Ubicacion = IsNull(Regs("ubicacion"))
                DatosViaje.CodLocalidad = IsNull(Regs("cod_localidad"))
                DatosViaje.Localidad = IsNull(Regs("localidad"))
                DatosViaje.Latitud = IsNull(Regs("latitud"))
                DatosViaje.Longitud = IsNull(Regs("longitud"))
                DatosViaje.CodCalle1 = IsNull(Regs("cod_calle1"))
                DatosViaje.NombreCalle1 = IsNull(Regs("nombre_calle1"))
                DatosViaje.CodCalle2 = IsNull(Regs("cod_calle2"))
                DatosViaje.NombreCalle2 = IsNull(Regs("nombre_calle2"))
                DatosViaje.NroTelefono = IsNull(Regs("nro_telefono"))
                DatosViaje.Comentario = IsNull(Regs("comentario"))
                DatosViaje.Direccion = IsNull(Regs("direccion"))
                DatosViaje.CodZona = IsNull(Regs("cod_zona"))
                DatosViaje.CodZonaOrigen = IsNull(Regs("cod_zona_origen"))
                DatosViaje.TipoServicio = IsNull(Regs("tipo_servicio_requerido"))
                DatosViaje.CategoriaMovil = IsNull(Regs("categ_movil"))
                DatosViaje.TipoViaje = IsNull(Regs("tipo_viaje"))
                DatosViaje.EstadoRespectoCliente = IsNull(Regs("estado_respecto_cliente"))
                DatosViaje.FechaInicioViaje = IsNull(Regs("fecha_hora_inicio_viaje"))
                DatosViaje.LatitudInicioViaje = IsNull(Regs("latitud_inicio_viaje"))
                DatosViaje.LongitudInicioViaje = IsNull(Regs("longitud_inicio_viaje"))
                DatosViaje.FechaFinViaje = IsNull(Regs("fecha_hora_fin_viaje"))
                DatosViaje.LatitudFinViaje = IsNull(Regs("latitud_fin_viaje"))
                DatosViaje.LongitudFinViaje = IsNull(Regs("longitud_fin_viaje"))
                DatosViaje.CodMotivoCancel = IsNull(Regs("cod_motivo_cancelacion"))
                DatosViaje.EstadoRespectoMovil = IsNull(Regs("estado_respecto_movil"))
                DatosViaje.FechaInvitacion = IsNull(Regs("fecha_hora_invitacion"))
                DatosViaje.LatitudInvitacion = IsNull(Regs("latitud_invitacion"))
                DatosViaje.LongitudInvitacion = IsNull(Regs("longitud_invitacion"))
                DatosViaje.FechaResolucion = IsNull(Regs("fecha_hora_resolucion"))
                DatosViaje.LatitudResolucion = IsNull(Regs("latitud_resolucion"))
                DatosViaje.LongitudResolucion = IsNull(Regs("longitud_resolucion"))
                DatosViaje.NroMovilAsignado = IsNull(Regs("nro_movil_asignado"))
                DatosViaje.NroFlotaAsignada = IsNull(Regs("nro_flota_asignada"))
                DatosViaje.Importe = IsNull(Regs("importe"))
                DatosViaje.ImportePactado = IsNull(Regs("importe_pactado"))
                DatosViaje.FinAnormal = CType(IsNull(Regs("ind_fin_anormal")), Boolean)
                DatosViaje.ViajeUrgente = CType(IsNull(Regs("ind_viaje_urgente")), Boolean)
                DatosViaje.NroConductor = IsNull(Regs("nro_conductor"))
                DatosViaje.NroCupon = IsNull(Regs("nro_cupon"))
                DatosViaje.CodZonaFutura = IsNull(Regs("cod_zona_futura"))
                DatosViaje.KmRecorridosLibre = IsNull(Regs("km_recorridos_libre"))
                DatosViaje.VelocidadMaxLibre = IsNull(Regs("velocidad_maxima_libre"))
                DatosViaje.NroCorrLibre = IsNull(Regs("nro_correlativo_libre"))
                DatosViaje.KmRecorridosOcupado = IsNull(Regs("km_recorridos_ocupado"))
                DatosViaje.VelocidadMaxOcupado = IsNull(Regs("velocidad_maxima_ocupado"))
                DatosViaje.MinCobradosEspera = IsNull(Regs("minutos_cobrados_espera"))
                DatosViaje.TipoTarifa = IsNull(Regs("tipo_tarifa"))
                DatosViaje.NroCorrOcupado = IsNull(Regs("nro_correlativo_ocupado"))
                DatosViaje.Nombre = IsNull(Regs("nombre"))
                DatosViaje.FechaHoraRequeridaContacto = IsNull(Regs("fecha_hora_requerida_contacto"))
                DatosViaje.NroUsuarioDespachador = IsNull(Regs("nro_usuario_despachador"))
                DatosViaje.Observaciones = IsNull(Regs("observaciones"))
                DatosViaje.ImportePactadoTaxi = IsNull(Regs("Pactado_Taxi"))
                DatosViaje.CodParada = IsNull(Regs("cod_parada"))
                DatosViaje.NroCtaCte = IsNull(Regs("nro_ctacte"))

                Status = 0

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

        Return DatosViaje

    End Function

    Public Sub IniciarViaje(ByVal NroViaje As Integer, ByVal KmRecorridosLibre As Integer, ByVal VelocidadMaxLibre As Integer, ByVal NroCorrLibre As Integer, ByVal LatitudInicioViaje As Single, ByVal LongitudInicioViaje As Single)


        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("update viajes set estado_respecto_cliente = 'I' , fecha_hora_inicio_viaje = getdate()  , km_recorridos_libre = @KmRecorridosLibre , velocidad_maxima_libre = @VelocidadMaxLibre , nro_correlativo_libre = @NroCorrLibre , latitud_inicio_viaje =  @LatitudInicioViaje , longitud_inicio_viaje = @LongitudInicioViaje where nro_viaje = @NroViaje", DBConn)
            SqlCmd.Parameters.AddWithNullableValue("@KmRecorridosLibre", KmRecorridosLibre)
            SqlCmd.Parameters.AddWithNullableValue("@VelocidadMaxLibre", VelocidadMaxLibre)
            SqlCmd.Parameters.AddWithNullableValue("@NroCorrLibre", NroCorrLibre)
            SqlCmd.Parameters.AddWithNullableValue("@LatitudInicioViaje", LatitudInicioViaje)
            SqlCmd.Parameters.AddWithNullableValue("@LongitudInicioViaje", LongitudInicioViaje)
            SqlCmd.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

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


    Public Sub TerminarNormal(ByVal NroViaje As Integer, ByVal LatitudFinViaje As Single, ByVal LongitudFinViaje As Single)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim sqlTran As SqlTransaction = DBConn.BeginTransaction()

            Try

                Dim SqlUpdate As New SqlCommand("update viajes set ind_fin_anormal = 0 , estado_respecto_cliente = 'T' , estado_respecto_movil = 'T' , fecha_hora_fin_viaje = getdate() ,  latitud_fin_viaje =  @LatitudFinViaje , longitud_fin_viaje = @LongitudFinViaje where nro_viaje = @NroViaje", DBConn)
                SqlUpdate.Parameters.AddWithNullableValue("@LatitudFinViaje", LatitudFinViaje)
                SqlUpdate.Parameters.AddWithNullableValue("@LongitudFinViaje", LongitudFinViaje)
                SqlUpdate.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

                SqlUpdate.CommandType = CommandType.Text
                SqlUpdate.Transaction = sqlTran
                Dim Regs = SqlUpdate.ExecuteNonQuery


                Dim SqlInsert As New SqlCommand("insert into viajes_historicos (nro_viaje , nro_movil_requerido , nro_usuario_telefonista , fecha_hora_entrada , nro_cliente , cod_calle , nombre_calle , numero , ubicacion , cod_localidad , localidad , latitud , longitud , cod_calle1 , nombre_calle1 , cod_calle2 , nombre_calle2 , nro_telefono , comentario , cod_zona , cod_zona_origen , tipo_servicio_requerido  , categ_movil , tipo_viaje , estado_respecto_cliente , fecha_hora_inicio_viaje , latitud_inicio_viaje , longitud_inicio_viaje , fecha_hora_fin_viaje , latitud_fin_viaje , longitud_fin_viaje , cod_motivo_cancelacion , estado_respecto_movil , fecha_hora_invitacion , latitud_invitacion , longitud_invitacion , fecha_hora_resolucion , latitud_resolucion , longitud_resolucion , nro_movil_asignado , nro_flota_asignada , importe , importe_pactado , ind_fin_anormal , ind_viaje_urgente , nro_conductor , nro_cupon , cod_zona_futura , km_recorridos_libre , velocidad_maxima_libre , nro_correlativo_libre , km_recorridos_ocupado , velocidad_maxima_ocupado , minutos_cobrados_espera , tipo_tarifa , nro_correlativo_ocupado , nombre , fecha_hora_requerida_contacto , nro_usuario_despachador , observaciones , Pactado_Taxi , nro_ctacte , nro_subcta , lugar_liq , zona_final , cod_parada ) " & _
                                                "                        select nro_viaje , nro_movil_requerido , nro_usuario_telefonista , fecha_hora_entrada , nro_cliente , cod_calle , nombre_calle , numero , ubicacion , cod_localidad , localidad , latitud , longitud , cod_calle1 , nombre_calle1 , cod_calle2 , nombre_calle2 , nro_telefono , comentario , cod_zona , cod_zona_origen , tipo_servicio_requerido  , categ_movil , tipo_viaje , estado_respecto_cliente , fecha_hora_inicio_viaje , latitud_inicio_viaje , longitud_inicio_viaje , fecha_hora_fin_viaje , latitud_fin_viaje , longitud_fin_viaje , cod_motivo_cancelacion , estado_respecto_movil , fecha_hora_invitacion , latitud_invitacion , longitud_invitacion , fecha_hora_resolucion , latitud_resolucion , longitud_resolucion , nro_movil_asignado , nro_flota_asignada , importe , importe_pactado , ind_fin_anormal , ind_viaje_urgente , nro_conductor , nro_cupon , cod_zona_futura , km_recorridos_libre , velocidad_maxima_libre , nro_correlativo_libre , km_recorridos_ocupado , velocidad_maxima_ocupado , minutos_cobrados_espera , tipo_tarifa , nro_correlativo_ocupado , nombre , fecha_hora_requerida_contacto , nro_usuario_despachador , observaciones , Pactado_Taxi  ,nro_ctacte , nro_subcta , lugar_liq , zona_final , cod_parada from viajes where nro_viaje =  @NroViaje", DBConn)
                SqlInsert.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

                SqlInsert.CommandType = CommandType.Text
                SqlInsert.Transaction = sqlTran
                Dim Regs2 = SqlInsert.ExecuteNonQuery

                Dim SqlDelete As New SqlCommand("delete from viajes where nro_viaje =  @NroViaje", DBConn)
                SqlDelete.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

                SqlDelete.Transaction = sqlTran
                SqlDelete.CommandType = CommandType.Text
                Dim Regs3 = SqlDelete.ExecuteNonQuery

                sqlTran.Commit()

                Status = 0

            Catch ex As Exception
                sqlTran.Rollback()
                Status = 1
                Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            End Try

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            Status = 1
        End Try

    End Sub

    Public Sub TerminarAnormal(ByVal NroViaje As Integer, ByVal LatitudFinViaje As Single, ByVal LongitudFinViaje As Single, ByVal CodMotivoCancelacion As Integer, ByVal Importe As Decimal)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim sqlTran As SqlTransaction = DBConn.BeginTransaction()

            Try

                Dim SqlUpdate As New SqlCommand("update viajes set ind_fin_anormal = -1  , estado_respecto_cliente = 'T' , estado_respecto_movil = 'T' , fecha_hora_fin_viaje = getdate() , nro_usuario_despachador = @NroUsuarioDespachador , latitud_fin_viaje =  @LatitudFinViaje , longitud_fin_viaje = @LongitudFinViaje , cod_motivo_cancelacion =  @CodMotivoCancel , importe = @Importe where nro_viaje = @NroViaje", DBConn)
                SqlUpdate.Parameters.AddWithNullableValue("@LatitudFinViaje", LatitudFinViaje)
                SqlUpdate.Parameters.AddWithNullableValue("@LongitudFinViaje", LongitudFinViaje)
                SqlUpdate.Parameters.AddWithNullableValue("@CodMotivoCancelacion", CodMotivoCancelacion)
                SqlUpdate.Parameters.AddWithNullableValue("@Importe", Importe)
                SqlUpdate.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

                SqlUpdate.CommandType = CommandType.Text
                SqlUpdate.Transaction = sqlTran
                Dim Regs = SqlUpdate.ExecuteNonQuery


                Dim SqlInsert1 As New SqlCommand("insert into viajes_historicos (nro_viaje , nro_movil_requerido , nro_usuario_telefonista , fecha_hora_entrada , nro_cliente , cod_calle , nombre_calle , numero , ubicacion , cod_localidad , localidad , latitud , longitud , cod_calle1 , nombre_calle1 , cod_calle2 , nombre_calle2 , nro_telefono , comentario , cod_zona , cod_zona_origen , tipo_servicio_requerido  , categ_movil , tipo_viaje , estado_respecto_cliente , fecha_hora_inicio_viaje , latitud_inicio_viaje , longitud_inicio_viaje , fecha_hora_fin_viaje , latitud_fin_viaje , longitud_fin_viaje , cod_motivo_cancelacion , estado_respecto_movil , fecha_hora_invitacion , latitud_invitacion , longitud_invitacion , fecha_hora_resolucion , latitud_resolucion , longitud_resolucion , nro_movil_asignado , nro_flota_asignada , importe , importe_pactado , ind_fin_anormal , ind_viaje_urgente , nro_conductor , nro_cupon , cod_zona_futura , km_recorridos_libre , velocidad_maxima_libre , nro_correlativo_libre , km_recorridos_ocupado , velocidad_maxima_ocupado , minutos_cobrados_espera , tipo_tarifa , nro_correlativo_ocupado , nombre , fecha_hora_requerida_contacto , nro_usuario_despachador , observaciones , Pactado_Taxi , nro_ctacte , nro_subcta , lugar_liq , zona_final , cod_parada ) " & _
                                                 "                        select nro_viaje , nro_movil_requerido , nro_usuario_telefonista , fecha_hora_entrada , nro_cliente , cod_calle , nombre_calle , numero , ubicacion , cod_localidad , localidad , latitud , longitud , cod_calle1 , nombre_calle1 , cod_calle2 , nombre_calle2 , nro_telefono , comentario , cod_zona , cod_zona_origen , tipo_servicio_requerido  , categ_movil , tipo_viaje , estado_respecto_cliente , fecha_hora_inicio_viaje , latitud_inicio_viaje , longitud_inicio_viaje , fecha_hora_fin_viaje , latitud_fin_viaje , longitud_fin_viaje , cod_motivo_cancelacion , estado_respecto_movil , fecha_hora_invitacion , latitud_invitacion , longitud_invitacion , fecha_hora_resolucion , latitud_resolucion , longitud_resolucion , nro_movil_asignado , nro_flota_asignada , importe , importe_pactado , ind_fin_anormal , ind_viaje_urgente , nro_conductor , nro_cupon , cod_zona_futura , km_recorridos_libre , velocidad_maxima_libre , nro_correlativo_libre , km_recorridos_ocupado , velocidad_maxima_ocupado , minutos_cobrados_espera , tipo_tarifa , nro_correlativo_ocupado , nombre , fecha_hora_requerida_contacto , nro_usuario_despachador , observaciones , Pactado_Taxi  ,nro_ctacte , nro_subcta , lugar_liq , zona_final , cod_parada from viajes where nro_viaje =  @NroViaje", DBConn)

                SqlInsert1.CommandType = CommandType.Text
                SqlInsert1.Transaction = sqlTran

                SqlInsert1.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

                Dim Regs1 = SqlInsert1.ExecuteNonQuery


                Dim SqlInsert2 As New SqlCommand("insert into viajes_historicos_destinos(nro_viaje , nro_correlativo_destino , cod_calle , nombre_calle , numero , ubicacion , cod_localidad , localidad , latitud , longitud , cod_calle1 , nombre_calle1 , cod_calle2 , nombre_calle2 , nro_telefono , comentario , cod_zona , fecha_hora_contacto , estado , nombre , observaciones) select nro_viaje , nro_correlativo_destino , cod_calle , nombre_calle , numero , ubicacion , cod_localidad , localidad , latitud , longitud , cod_calle1 , nombre_calle1 , cod_calle2 , nombre_calle2 , nro_telefono , comentario , cod_zona , fecha_hora_contacto , estado , nombre , observaciones from viajes_destinos where nro_viaje =  @NroViaje", DBConn)

                SqlInsert2.CommandType = CommandType.Text
                SqlInsert2.Transaction = sqlTran

                SqlInsert2.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

                Dim Regs2 = SqlInsert2.ExecuteNonQuery

                Dim SqlInsert3 As New SqlCommand("insert into viajes_historicos_tarjetas(nro_viaje , tarjeta , nro_tarjeta , cod_seguridad , vencimiento , cod_autorizacion , importe) select nro_viaje , tarjeta , nro_tarjeta , cod_seguridad , vencimiento , cod_autorizacion , importe from viajes_tarjetas where nro_viaje =  @NroViaje", DBConn)

                SqlInsert3.CommandType = CommandType.Text
                SqlInsert3.Transaction = sqlTran

                SqlInsert3.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

                Dim Regs3 = SqlInsert3.ExecuteNonQuery

                Dim SqlInsert4 As New SqlCommand("insert into viajes_historicos_adicionales(nro_viaje , cod_adicional , cantidad , valor_adicional) select nro_viaje , cod_adicional , cantidad , valor_adicional from viajes_adicionales where nro_viaje =  @NroViaje", DBConn)

                SqlInsert4.CommandType = CommandType.Text
                SqlInsert4.Transaction = sqlTran

                SqlInsert4.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

                Dim Regs4 = SqlInsert4.ExecuteNonQuery


                Dim SqlDelete1 As New SqlCommand("delete from viajes_adicionales where nro_viaje = @NroViaje", DBConn)

                SqlDelete1.Transaction = sqlTran
                SqlDelete1.CommandType = CommandType.Text

                SqlDelete1.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

                Dim Regs5 = SqlDelete1.ExecuteNonQuery

                Dim SqlDelete2 As New SqlCommand("delete from viajes_destinos where nro_viaje =  @NroViaje", DBConn)

                SqlDelete2.Transaction = sqlTran
                SqlDelete2.CommandType = CommandType.Text

                SqlDelete2.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

                Dim Regs6 = SqlDelete2.ExecuteNonQuery

                Dim SqlDelete3 As New SqlCommand("delete from viajes_tarjetas where nro_viaje =  @NroViaje", DBConn)

                SqlDelete3.Transaction = sqlTran
                SqlDelete3.CommandType = CommandType.Text

                SqlDelete3.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

                Dim Regs7 = SqlDelete3.ExecuteNonQuery

                Dim SqlDelete4 As New SqlCommand("delete from viajes_factores where nro_viaje = @NroViaje", DBConn)

                SqlDelete4.Transaction = sqlTran
                SqlDelete4.CommandType = CommandType.Text

                SqlDelete4.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

                Dim Regs8 = SqlDelete4.ExecuteNonQuery

                Dim SqlDelete5 As New SqlCommand("delete from viajes where nro_viaje =  @NroViaje", DBConn)

                SqlDelete5.Transaction = sqlTran
                SqlDelete5.CommandType = CommandType.Text

                SqlDelete5.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

                Dim Regs9 = SqlDelete5.ExecuteNonQuery

                sqlTran.Commit()

                Status = 0
            Catch ex As Exception
                sqlTran.Rollback()
                Status = 1
                Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            End Try

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            Status = 1
        End Try


    End Sub

    Public Sub Cobrar(ByVal NroViaje As Integer, ByVal KmRecorridosOcupado As String, ByVal VelocidadMaxOcupado As Integer, ByVal TipoTarifa As Integer, ByVal Importe As String, ByVal NroCorrLibre As Integer, ByVal MinCobradosEspera As Integer)


        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("update viajes set km_recorridos_ocupado = @KmRecorridosOcupado , velocidad_maxima_ocupado = @VelocidadMaxOcupado , tipo_tarifa = @TipoTarifa , importe = @Importe , nro_correlativo_ocupado = @NroCorrLibre , minutos_cobrados_espera = @MinCobradosEspera where nro_viaje = @NroViaje", DBConn)
            SqlCmd.Parameters.AddWithNullableValue("@KmRecorridosOcupado", KmRecorridosOcupado)
            SqlCmd.Parameters.AddWithNullableValue("@VelocidadMaxOcupado", VelocidadMaxOcupado)
            SqlCmd.Parameters.AddWithNullableValue("@TipoTarifa", TipoTarifa)
            SqlCmd.Parameters.AddWithNullableValue("@Importe", Importe)
            SqlCmd.Parameters.AddWithNullableValue("@NroCorrLibre", NroCorrLibre)
            SqlCmd.Parameters.AddWithNullableValue("@MinCobradosEspera", MinCobradosEspera)
            SqlCmd.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

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

    Public Sub EliminarViajePreIngresado(ByVal NroViaje As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim sqlTran As SqlTransaction = DBConn.BeginTransaction()

            Try

                Dim SqlDelete1 As New SqlCommand("delete from viajes_adicionales where nro_viaje = @NroViaje", DBConn)

                SqlDelete1.Transaction = sqlTran
                SqlDelete1.CommandType = CommandType.Text

                SqlDelete1.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

                Dim Regs5 = SqlDelete1.ExecuteNonQuery

                Dim SqlDelete2 As New SqlCommand("delete from viajes_destinos where nro_viaje =  @NroViaje", DBConn)

                SqlDelete2.Transaction = sqlTran
                SqlDelete2.CommandType = CommandType.Text

                SqlDelete2.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

                Dim Regs6 = SqlDelete2.ExecuteNonQuery

                Dim SqlDelete3 As New SqlCommand("delete from viajes_tarjetas where nro_viaje =  @NroViaje", DBConn)

                SqlDelete3.Transaction = sqlTran
                SqlDelete3.CommandType = CommandType.Text

                SqlDelete3.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

                Dim Regs7 = SqlDelete3.ExecuteNonQuery

                Dim SqlDelete4 As New SqlCommand("delete from viajes_factores where nro_viaje = @NroViaje", DBConn)

                SqlDelete4.Transaction = sqlTran
                SqlDelete4.CommandType = CommandType.Text

                SqlDelete4.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

                Dim Regs8 = SqlDelete4.ExecuteNonQuery

                Dim SqlDelete5 As New SqlCommand("delete from viajes where nro_viaje =  @NroViaje", DBConn)

                SqlDelete5.Transaction = sqlTran
                SqlDelete5.CommandType = CommandType.Text

                SqlDelete5.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

                Dim Regs9 = SqlDelete5.ExecuteNonQuery


                sqlTran.Commit()

                Status = 0

            Catch ex As Exception
                sqlTran.Rollback()
                Status = 1
                Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            End Try

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

        Catch ex As Exception
            Status = 1
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

    End Sub


    Public Sub Redespachar(ByVal NroViaje As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd1 As New SqlCommand("select cod_parada from  viajes where nro_viaje = @NroViaje", DBConn)
            SqlCmd1.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

            SqlCmd1.CommandType = CommandType.Text
            Dim CodZona = SqlCmd1.ExecuteScalar

            Dim SqlUpdate As New SqlCommand("update viajes set  nro_movil_requerido = 0, nromesaDespacho = 0, estado_respecto_movil = 'P', nro_movil_asignado = 0, fecha_hora_invitacion = null, cod_zona_origen = 0, nro_usuario_despachador = 0 , latitud_invitacion = 0, longitud_invitacion = 0 , nro_cupon = 0, cod_parada = @CodZona where nro_viaje = @NroViaje", DBConn)
            SqlUpdate.Parameters.AddWithNullableValue("@CodZona", CodZona)
            SqlUpdate.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

            SqlUpdate.CommandType = CommandType.Text
            SqlUpdate.ExecuteNonQuery()

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


    Public Sub AceptarViaje(ByVal NroViaje As Integer, ByVal NroCupon As Integer, ByVal LatitudResolucion As Single, ByVal LongitudResolucion As Single, ByVal NroConductor As Integer)


        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlUpdate As New SqlCommand("update viajes set  estado_respecto_movil = 'M'  , nro_cupon = @NroCupon , latitud_resolucion = @LatitudResolucion , longitud_resolucion = @LongitudResolucion , nro_conductor = @NroConductor  where nro_viaje =  @NroViaje", DBConn)
            SqlUpdate.Parameters.AddWithNullableValue("@NroCupon", NroCupon)
            SqlUpdate.Parameters.AddWithNullableValue("@LatitudResolucion", LatitudResolucion)
            SqlUpdate.Parameters.AddWithNullableValue("@LongitudResolucion", LongitudResolucion)
            SqlUpdate.Parameters.AddWithNullableValue("@NroConductor", NroConductor)
            SqlUpdate.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

            SqlUpdate.CommandType = CommandType.Text
            Dim Regs = SqlUpdate.ExecuteNonQuery


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


    Public Sub TerminarViajesCanal(ByVal NroCanal As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim sqlTran As SqlTransaction = DBConn.BeginTransaction()

            Try

                Dim SqlUpdate As New SqlCommand("update viajes set ind_fin_anormal = -1  , estado_respecto_cliente = 'T' , estado_respecto_movil = 'T' , fecha_hora_fin_viaje = getdate() , importe = 0  where estado_respecto_cliente <> 'T' and nro_movil_asignado in (select nro_movil from moviles where nro_canal_comunicaciones = @NroCanal", DBConn)
                SqlUpdate.Parameters.AddWithNullableValue("@NroCanal", NroCanal)

                SqlUpdate.CommandType = CommandType.Text
                SqlUpdate.Transaction = sqlTran
                Dim Regs = SqlUpdate.ExecuteNonQuery


                sqlTran.Commit()

                Status = 0

            Catch ex As Exception
                sqlTran.Rollback()
                Status = 1
                Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            End Try

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

        Catch ex As Exception
            Status = 1
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

    End Sub


    Public Sub PonerEnConsulta(ByVal NroViaje As Integer, ByVal NroMovilAsignado As Integer, ByVal NroFlotaAsignada As Integer, ByVal LatitudInvitacion As Single, ByVal LongitudInvitacion As Single, ByVal CodZonaOrigen As Integer, ByVal NroCupon As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlUpdate As New SqlCommand("update viajes set estado_respecto_movil = 'C' , fecha_hora_invitacion = getdate() , nro_movil_asignado = @NroMovilAsignado , cod_zona_origen = @CodZonaOrigen ,  latitud_invitacion = @LatitudInvitacion , longitud_invitacion = @LongitudInvitacion , nro_cupon = @NroCupon , nro_flota_asignada =  @NroFlotaAsignada where nro_viaje = @NroViaje", DBConn)
            SqlUpdate.Parameters.AddWithNullableValue("@NroMovilAsignado", NroMovilAsignado)
            SqlUpdate.Parameters.AddWithNullableValue("@NroFlotaAsignada", NroFlotaAsignada)
            SqlUpdate.Parameters.AddWithNullableValue("@LatitudInvitacion", LatitudInvitacion)
            SqlUpdate.Parameters.AddWithNullableValue("@LongitudInvitacion", LongitudInvitacion)
            SqlUpdate.Parameters.AddWithNullableValue("@CodZonaOrigen", CodZonaOrigen)
            SqlUpdate.Parameters.AddWithNullableValue("@NroCupon", NroCupon)
            SqlUpdate.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

            SqlUpdate.CommandType = CommandType.Text
            Dim Regs = SqlUpdate.ExecuteNonQuery

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

    Public Sub CambiarZonaFutura(ByVal NroViaje As Integer, ByVal CodZonaFutura As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlUpdate As New SqlCommand("update viajes set cod_zona_futura = @CodZonaFutura where nro_viaje = @NroViaje", DBConn)
            SqlUpdate.Parameters.AddWithNullableValue("@CodZonaFutura", CodZonaFutura)
            SqlUpdate.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

            SqlUpdate.CommandType = CommandType.Text
            Dim Regs = SqlUpdate.ExecuteNonQuery

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


    Public Sub Marcar(ByVal NroViaje As Integer, ByVal NroUsuarioDespachador As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlUpdate As New SqlCommand("update viajes set cod_motivo_cancelacion = 9 , nro_usuario_despachador = @NroUsuarioDespachador where nro_viaje = @NroViaje", DBConn)
            SqlUpdate.Parameters.AddWithNullableValue("@NroUsuarioDespachador", NroUsuarioDespachador)
            SqlUpdate.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

            SqlUpdate.CommandType = CommandType.Text
            Dim Regs = SqlUpdate.ExecuteNonQuery

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

    Public Sub AsignarCupon(ByVal NroViaje As Integer, ByVal NroCupon As Integer, ByVal NroCliente As Integer, ByVal NroUsuarioDespachador As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlUpdate As New SqlCommand("update viajes set nro_cupon = @NroCupon  , nro_cliente = @NroCliente , nro_usuario_despachador = @NroUsuarioDespachador  where nro_viaje = @NroViaje", DBConn)
            SqlUpdate.Parameters.AddWithNullableValue("@NroCupon", NroCupon)
            SqlUpdate.Parameters.AddWithNullableValue("@NroCliente", NroCliente)
            SqlUpdate.Parameters.AddWithNullableValue("@NroUsuarioDespachador", NroUsuarioDespachador)
            SqlUpdate.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

            SqlUpdate.CommandType = CommandType.Text
            Dim Regs = SqlUpdate.ExecuteNonQuery

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


    Public Sub AgregarDestino(ByVal NroViaje As Integer, ByVal NroDestino As Integer, ByVal CodCalle As Integer?, ByVal NombreCalle As String, ByVal Numero As Integer?, ByVal Ubicacion As String, ByVal CodLocalidad As Integer?, ByVal NombreLocalidad As String, ByVal Latitud As Single?, ByVal Longitud As Single?, ByVal CodCalle1 As Integer, ByVal NombreCalle1 As String, ByVal CodCalle2 As Integer, ByVal NombreCalle2 As String, ByVal NroTelefono As Int64, ByVal Comentario As String, ByVal Direccion As String, ByVal CodZona As Integer, ByVal FechaHoraContacto As Date, ByVal EstadoDestino As String, ByVal Nombre As String, ByVal Observaciones As String)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlInsert As New SqlCommand("insert into viajes_destinos (nro_viaje, nro_correlativo_destino, cod_calle, nombre_calle, numero, ubicacion, cod_localidad, localidad, latitud, longitud, cod_calle1, nombre_calle1, cod_calle2, nombre_calle2, nro_telefono, comentario, direccion, cod_zona, fecha_hora_contacto, estado, nombre, observaciones) values (@NroViaje, @NroDestino, @CodCalle, @NombreCalle, @Numero, @Ubicacion, @CodLocalidad, @NombreLocalidad, @Latitud, @Longitud, @CodCalle1, @NombreCalle1, @CodCalle2, @NombreCalle2, @NroTelefono, @Comentario, @Direccion, @CodZona, @FechaHoraContacto, @EstadoDestino, @Nombre, @Observaciones)", DBConn)
            SqlInsert.CommandType = CommandType.Text

            SqlInsert.Parameters.AddWithNullableValue("@NroViaje", NroViaje)
            SqlInsert.Parameters.AddWithNullableValue("@NroDestino", NroDestino)
            SqlInsert.Parameters.AddWithNullableValue("@CodCalle", CodCalle)
            SqlInsert.Parameters.AddWithNullableValue("@NombreCalle", NombreCalle)
            SqlInsert.Parameters.AddWithNullableValue("@Numero", Numero)
            SqlInsert.Parameters.AddWithNullableValue("@Ubicacion", Ubicacion)
            SqlInsert.Parameters.AddWithNullableValue("@CodLocalidad", CodLocalidad)
            SqlInsert.Parameters.AddWithNullableValue("@NombreLocalidad", NombreLocalidad)
            SqlInsert.Parameters.AddWithNullableValue("@Latitud", Latitud)
            SqlInsert.Parameters.AddWithNullableValue("@Longitud", Longitud)
            SqlInsert.Parameters.AddWithNullableValue("@CodCalle1", CodCalle1)
            SqlInsert.Parameters.AddWithNullableValue("@NombreCalle1", NombreCalle1)
            SqlInsert.Parameters.AddWithNullableValue("@CodCalle2", CodCalle2)
            SqlInsert.Parameters.AddWithNullableValue("@NombreCalle2", NombreCalle2)
            SqlInsert.Parameters.AddWithNullableValue("@NroTelefono", NroTelefono)
            SqlInsert.Parameters.AddWithNullableValue("@Comentario", Comentario)
            SqlInsert.Parameters.AddWithNullableValue("@Direccion", Direccion)
            SqlInsert.Parameters.AddWithNullableValue("@CodZona", CodZona)
            SqlInsert.Parameters.AddWithNullableValue("@FechaHoraContacto", FechaHoraContacto)
            SqlInsert.Parameters.AddWithNullableValue("@EstadoDestino", EstadoDestino)
            SqlInsert.Parameters.AddWithNullableValue("@Nombre", Nombre)
            SqlInsert.Parameters.AddWithNullableValue("@Observaciones", Observaciones)

            Dim Regs = SqlInsert.ExecuteNonQuery

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

    Public Function ObtenerDatosDestino(ByVal NroViaje As Integer, ByVal NroDestino As Integer) As DatosDestino


        Dim DatosDestino As New DatosDestino

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("Select * from viajes_destinos where nro_viaje = @NroViaje and  and nro_correlativo_destino = @NroDestino", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@NroViaje", NroViaje)
            SqlCmd.Parameters.AddWithNullableValue("@NroDestino", NroDestino)

            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            If Regs.HasRows Then

                Regs.Read()

                DatosDestino.NroViaje = IsNull(Regs("nro_viaje"))
                DatosDestino.CodCalle = IsNull(Regs("cod_calle"))
                DatosDestino.NombreCalle = IsNull(Regs("nombre_calle"))
                DatosDestino.Numero = IsNull(Regs("numero"))
                DatosDestino.Ubicacion = IsNull(Regs("ubicacion"))
                DatosDestino.CodLocalidad = IsNull(Regs("cod_localidad"))
                DatosDestino.NombreLocalidad = IsNull(Regs("localidad"))
                DatosDestino.Latitud = IsNull(Regs("latitud"))
                DatosDestino.Longitud = IsNull(Regs("longitud"))
                DatosDestino.CodCalle1 = IsNull(Regs("cod_calle1"))
                DatosDestino.NombreCalle1 = IsNull(Regs("nombre_calle1"))
                DatosDestino.CodCalle2 = IsNull(Regs("cod_calle2"))
                DatosDestino.NombreCalle2 = IsNull(Regs("nombre_calle2"))
                DatosDestino.NroTelefono = IsNull(Regs("nro_telefono"))
                DatosDestino.Comentario = IsNull(Regs("comentario"))
                DatosDestino.Direccion = IsNull(Regs("direccion"))
                DatosDestino.CodZona = IsNull(Regs("cod_zona"))
                DatosDestino.Nombre = IsNull(Regs("nombre"))
                DatosDestino.Observaciones = IsNull(Regs("observaciones"))

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

        Return DatosDestino

    End Function

    Public Sub AsignarMovilPosible(ByVal NroViaje As Integer, ByVal NroMovilAsignado As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlUpdate As New SqlCommand("update viajes set nro_movil_asignado = @NroMovilAsignado  where nro_viaje = @NroViaje", DBConn)
            SqlUpdate.Parameters.AddWithNullableValue("@NroMovilAsignado", NroMovilAsignado)
            SqlUpdate.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

            SqlUpdate.CommandType = CommandType.Text
            Dim Regs = SqlUpdate.ExecuteNonQuery

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


    Public Sub PonerEnBusqueda(ByVal NroViaje As Integer, ByVal FechaInvitacion As Date)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlUpdate As New SqlCommand("update viajes set nro_movil_asignado = @NroMovilAsignado, fecha_hora_invitacion = @FechaInvitacion  where nro_viaje = @NroViaje", DBConn)
            SqlUpdate.Parameters.AddWithNullableValue("@FechaInviatacion", FechaInvitacion)
            SqlUpdate.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

            SqlUpdate.CommandType = CommandType.Text
            Dim Regs = SqlUpdate.ExecuteNonQuery

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


    Public Sub SacarDeBusqueda(ByVal NroViaje As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlUpdate As New SqlCommand("update viajes set nro_movil_asignado = 0, fecha_hora_invitacion = NULL  where nro_viaje = @NroViaje", DBConn)
            SqlUpdate.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

            SqlUpdate.CommandType = CommandType.Text
            Dim Regs = SqlUpdate.ExecuteNonQuery

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

    Public Function ProximoNumeroDestino(ByVal NroViaje As Integer) As Integer

        Dim ProxNroDestino As Integer

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("select max(nro_correlativo_destino) as UltNroDestino from viajes_destinos where nro_viaje = @NroViaje", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            If Regs.HasRows Then

                Regs.Read()

                If IsDBNull(Regs("UltNroDestino")) Then
                    ProxNroDestino = 1
                Else
                    ProxNroDestino = CType(Regs("UltNroDestino"), Integer) + 1
                End If

            Else
                ProxNroDestino = 1
            End If

            Regs.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            ProxNroDestino = 0
            Status = 1
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

        Return ProximoNumeroDestino


    End Function


    Public Sub CrearImporteAdicional(ByVal NroViaje As Integer, ByVal CodAdicional As Integer, ByVal CantAdicional As Integer, ByVal ValorAdicional As Decimal)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlInsert As New SqlCommand("insert into viajes_adicionales (nro_viaje as integer, cod_adicional as integer, cantidad as integer, valor_adicional) values (@NroViaje,@CodAdicional, @CantAdicional,@ValorAdicional", DBConn)
            SqlInsert.CommandType = CommandType.Text

            SqlInsert.Parameters.AddWithNullableValue("@NroViaje", NroViaje)
            SqlInsert.Parameters.AddWithNullableValue("@CodAdicional", CodAdicional)
            SqlInsert.Parameters.AddWithNullableValue("@CantAdicional", CantAdicional)
            SqlInsert.Parameters.AddWithNullableValue("@ValorAdicional", ValorAdicional)

            Dim Regs = SqlInsert.ExecuteNonQuery

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

    Public Function ObtenerDatosAdicional(ByVal NroViaje As Integer, ByVal CodAdicional As Integer) As DatosAdicional


        Dim DatosADic As New DatosAdicional

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("select * from viajes_adicionales where nro_viaje = @NroViaje  and cod_adicional = @CodAdicional", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@NroViaje", NroViaje)
            SqlCmd.Parameters.AddWithNullableValue("@CodAdicional", CodAdicional)

            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            If Regs.HasRows Then

                Regs.Read()

                DatosADic.CodAdicional = CodAdicional
                DatosADic.CantAdicional = IsNull(Regs("cantidad"))
                DatosADic.ValorAdicional = IsNull(Regs("valor_adicional"))
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

        Return DatosADic


    End Function

    Public Function LeerMovilOptimo(ByVal NroViaje As Integer, ByVal CategoriaMovil As String, ByVal TipoServicio As Integer) As MovilOptimo


        Dim MovilOpt As New MovilOptimo

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("select * from viajes_factores where nro_viaje = @NroViaje  order by nro_viaje as integer, tipo_servicio desc as integer, factor", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

            SqlCmd.CommandType = CommandType.Text

            Dim RegsFact = SqlCmd.ExecuteReader


            MovilOpt.NroMovilOptimo = 0
            MovilOpt.TipoServicio = 0
            MovilOpt.Distancia = 0

            Dim SqlCmdAppend As String

            If CategoriaMovil = "5" Or CategoriaMovil = "*" Then
                SqlCmdAppend = ""
            Else
                SqlCmdAppend = " and categ_movil <= '" & CategoriaMovil & "'"
            End If

            While RegsFact.Read

                Dim SqlCmd2 As New SqlCommand("select nro_movil as integer, estado_operativo as integer, tipo_servicio from moviles where nro_movil = @NroMovil " & SqlCmdAppend, DBConn)
                SqlCmd2.Parameters.AddWithNullableValue("@NroMovil", RegsFact("nro_movil"))

                SqlCmd2.CommandType = CommandType.Text

                Dim RegsMoviles = SqlCmd.ExecuteReader

                If RegsMoviles.HasRows Then

                    RegsMoviles.Read()

                    If (RegsMoviles("estado_operativo") = "L" And (RegsMoviles("tipo_servicio") = TipoServicio Or TipoServicio = 3)) Then

                        MovilOpt.NroMovilOptimo = RegsMoviles("nro_movil")
                        MovilOpt.TipoServicio = RegsMoviles("tipo_servicio")
                        MovilOpt.FactorOptimo = RegsMoviles("factor")
                        MovilOpt.Distancia = RegsMoviles("factor")

                        Exit While

                    End If

                End If

                RegsMoviles.Close()


            End While

            RegsFact.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            Status = 1
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

        Return MovilOpt



    End Function

    Public Sub BorrarFactores(ByVal NroViaje As Integer, ByVal NroMovilOptimo As Integer)

        Dim DBConn = New SqlConnection(ConnString)
        DBConn.Open()

        Try

            Dim SqlDelete As New SqlCommand("delete from viajes_factores where nro_viaje = @NroViaje  or  nro_movil = @NroMovilOptimo", DBConn)
            SqlDelete.Parameters.AddWithNullableValue("@NroViaje", NroViaje)
            SqlDelete.Parameters.AddWithNullableValue("@NroMovilOptimo", NroMovilOptimo)

            SqlDelete.CommandType = CommandType.Text
            Dim Regs3 = SqlDelete.ExecuteNonQuery



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

    Public Sub GrabarFactor(ByVal NroViajeRed As Integer, ByVal TipoServicio As Integer, ByVal FactorOptimo As Single, ByVal NroMovilOptimo As Integer, ByVal Distancia As Integer)

        Dim MovilOpt As New MovilOptimo

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("select nro_viaje as integer, tipo_servicio_requerido from viajes where nro_viaje like '%" & Format(NroViajeRed, "0000"), DBConn)

            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            If Regs.HasRows Then

                Regs.Read()

                If Regs("tipo_servicio_requerido") = TipoServicio Or Regs("tipo_servicio_requerido") = 3 Then

                    Dim SqlInsert As New SqlCommand("insert into viajes_factores (nro_viaje as integer, tipo_servicio as integer, factor as integer, nro_movil as integer, distancia) values (@NroViaje, @TipoServicio, @FactorOptimo,@NroMovilOptimo,@Distancia )", DBConn)
                    SqlInsert.CommandType = CommandType.Text

                    SqlInsert.Parameters.AddWithNullableValue("@NroViaje", Regs("nro_viaje"))
                    SqlInsert.Parameters.AddWithNullableValue("@TipoServicio", TipoServicio)
                    SqlInsert.Parameters.AddWithNullableValue("@NroMovilOptimo", NroMovilOptimo)
                    SqlInsert.Parameters.AddWithNullableValue("@Distancia", Distancia)

                    SqlInsert.ExecuteNonQuery()

                End If

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

    Public Function CumplirDestino(ByVal NroViaje As Integer) As DatosDestino

        Dim UltimoEstadoT As Integer
        Dim HayPendientes As Boolean
        Dim HoraUltimoEstadoT As Date
        Dim NroDestinoAEnviar As Integer
        Dim EsReintento As Boolean

        Dim Dest = New DatosDestino

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("select *  from viajes_destinos where nro_viaje = @NroViaje order by nro_correlativo_destino", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

            SqlCmd.CommandType = CommandType.Text

            Dim RegsDest = SqlCmd.ExecuteReader

            If RegsDest.HasRows Then

                UltimoEstadoT = 0
                HayPendientes = False

                While RegsDest.Read

                    Select Case RegsDest("Estado")

                        Case "T"

                            UltimoEstadoT = RegsDest("nro_correlativo_destino")
                            HoraUltimoEstadoT = RegsDest("fecha_hora_contacto")

                        Case "P"

                            HayPendientes = True

                            If UltimoEstadoT = 0 Then
                                NroDestinoAEnviar = RegsDest("nro_correlativo_destino")
                                EsReintento = False
                            Else
                                If DateDiff("n", HoraUltimoEstadoT, Now) > 2 Then
                                    NroDestinoAEnviar = RegsDest("nro_correlativo_destino")
                                    EsReintento = False
                                Else
                                    NroDestinoAEnviar = UltimoEstadoT
                                    EsReintento = True
                                End If
                            End If

                            Exit While

                    End Select

                End While

                'Verifica si hube algún registro en estado "P"
                'Si no lo hubo y el tiempo venció, entonces no hay más destinos
                'Si no lo hubo y el tiempo no venció se asume que es un reintento

                If UltimoEstadoT > 0 Then

                    If HayPendientes = False Then

                        If DateDiff("n", HoraUltimoEstadoT, Now) > 2 Then
                            NroDestinoAEnviar = 0
                        Else
                            NroDestinoAEnviar = UltimoEstadoT
                            EsReintento = True
                        End If
                    End If

                End If

            Else
                NroDestinoAEnviar = 0
            End If

            'Porcesa los datos de acuerdo a lo analizado antes.

            If NroDestinoAEnviar <> 0 Then

                If EsReintento = False Then

                    Dim SqlUpdate As New SqlCommand("update viajes_destinos set estado = 'T' as integer, fecha_hora_contacto  = getdate() where nro_viaje = @NroViaje  and  nro_correlativo_destino = @NroDestinoAEnviar", DBConn)
                    SqlUpdate.Parameters.AddWithNullableValue("@NroViaje", NroViaje)
                    SqlUpdate.Parameters.AddWithNullableValue("@NroDestinoAEnviar", NroDestinoAEnviar)

                    SqlUpdate.CommandType = CommandType.Text
                    Dim CantRegs = SqlUpdate.ExecuteNonQuery

                    If CantRegs > 0 Then

                        Dim SqlCmd2 As New SqlCommand("select * from viajes_destinos where nro_viaje = @NroViaje and nro_correlativo_destino = @NroDestinoAEnviar", DBConn)

                        SqlCmd2.Parameters.AddWithNullableValue("@NroViaje", NroViaje)
                        SqlCmd2.Parameters.AddWithNullableValue("@NroDestinoAEnviar", NroDestinoAEnviar)

                        SqlCmd2.CommandType = CommandType.Text

                        Dim RegsDest2 = SqlCmd2.ExecuteReader


                        If RegsDest2.HasRows Then

                            RegsDest2.Read()

                            Dest.NroViaje = RegsDest2("nro_viaje")
                            Dest.NroDestino = RegsDest2("nro_correlativo_destino")
                            Dest.CodCalle = RegsDest2("cod_calle")
                            Dest.NombreCalle = RegsDest2("nombre_calle")
                            Dest.Numero = RegsDest2("numero")
                            Dest.Ubicacion = RegsDest2("ubicacion")
                            Dest.CodLocalidad = RegsDest2("cod_localidad")
                            Dest.Latitud = RegsDest2("latitud")
                            Dest.Longitud = RegsDest2("longitud")
                            Dest.CodCalle1 = RegsDest2("cod_calle1")
                            Dest.NombreCalle1 = RegsDest2("nombre_calle1")
                            Dest.CodCalle2 = RegsDest2("cod_calle2")
                            Dest.NombreCalle2 = RegsDest2("nombre_calle2")
                            Dest.NroTelefono = RegsDest2("nro_telefono")
                            Dest.Comentario = RegsDest2("comentario")
                            Dest.Direccion = RegsDest2("direccion")
                            Dest.CodZona = RegsDest2("cod_zona")
                            Dest.FechaHoraContacto = RegsDest2("fecha_hora_contacto")
                            Dest.EstadoDestino = RegsDest2("Estado")
                            Dest.Nombre = RegsDest2("nombre")
                            Dest.Observaciones = RegsDest2("Observaciones")

                        End If

                        RegsDest2.Close()

                    Else

                    End If
                Else

                    Dim SqlCmd2 As New SqlCommand("select * from viajes_destinos where nro_viaje = @NroViaje and nro_correlativo_destino = @NroDestinoAEnviar", DBConn)

                    SqlCmd2.Parameters.AddWithNullableValue("@NroViaje", NroViaje)
                    SqlCmd2.Parameters.AddWithNullableValue("@NroDestinoAEnviar", NroDestinoAEnviar)

                    SqlCmd2.CommandType = CommandType.Text

                    Dim RegsDest2 = SqlCmd2.ExecuteReader


                    If RegsDest2.HasRows Then

                        RegsDest2.Read()

                        Dest.NroViaje = RegsDest2("nro_viaje")
                        Dest.NroDestino = RegsDest2("nro_correlativo_destino")
                        Dest.CodCalle = RegsDest2("cod_calle")
                        Dest.NombreCalle = RegsDest2("nombre_calle")
                        Dest.Numero = RegsDest2("numero")
                        Dest.Ubicacion = RegsDest2("ubicacion")
                        Dest.CodLocalidad = RegsDest2("cod_localidad")
                        Dest.Latitud = RegsDest2("latitud")
                        Dest.Longitud = RegsDest2("longitud")
                        Dest.CodCalle1 = RegsDest2("cod_calle1")
                        Dest.NombreCalle1 = RegsDest2("nombre_calle1")
                        Dest.CodCalle2 = RegsDest2("cod_calle2")
                        Dest.NombreCalle2 = RegsDest2("nombre_calle2")
                        Dest.NroTelefono = RegsDest2("nro_telefono")
                        Dest.Comentario = RegsDest2("comentario")
                        Dest.Direccion = RegsDest2("direccion")
                        Dest.CodZona = RegsDest2("cod_zona")
                        Dest.FechaHoraContacto = RegsDest2("fecha_hora_contacto")
                        Dest.EstadoDestino = RegsDest2("Estado")
                        Dest.Nombre = RegsDest2("nombre")
                        Dest.Observaciones = RegsDest2("Observaciones")

                    End If

                    RegsDest2.Close()

                End If

            Else

            End If

            RegsDest.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            Status = 1
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

        Return Dest

    End Function

    Public Sub ConfirmarDireccion(ByVal NroViaje As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlUpdate As New SqlCommand("update viajes set estado_respecto_movil = 'A' where nro_viaje = @NroViaje", DBConn)
            SqlUpdate.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

            SqlUpdate.CommandType = CommandType.Text
            Dim Regs = SqlUpdate.ExecuteNonQuery

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

    Public Sub PasarAEstadoX(ByVal NroViaje As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlUpdate As New SqlCommand("update viajes set estado_respecto_movil = 'X', estado_respecto_cliente = 'X' where nro_viaje = @NroViaje", DBConn)
            SqlUpdate.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

            SqlUpdate.CommandType = CommandType.Text
            Dim Regs = SqlUpdate.ExecuteNonQuery

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

    Public Sub SalirDeEstadoX(ByVal NroViaje As Integer)


        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlUpdate As New SqlCommand("update viajes set estado_respecto_movil = 'P', estado_respecto_cliente = 'P' where nro_viaje = @NroViaje", DBConn)
            SqlUpdate.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

            SqlUpdate.CommandType = CommandType.Text
            Dim Regs = SqlUpdate.ExecuteNonQuery

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

    Public Sub ModificarDatos(ByVal NroViaje As Integer, ByVal TipoServicio As Integer, ByVal NroMovilRequerido As Integer, ByVal NroFlotaAsignada As Integer, ByVal CategoriaMovil As String, ByVal FechaHoraRequeridaContacto As Date)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlUpdate As New SqlCommand("update viajes set  nro_movil_requerido = @NroMovilReq " & _
                                             ", tipo_servicio_requerido = @TipoServicio" & _
                                             ", nro_flota_asignada = @NroFlotaAsignada " & _
                                             ", nro_movil_asignado = 0 " & _
                                             ", categ_movil = @CategoriaMovil " & _
                                             ", fecha_hora_requerida_contacto = @FechaHoraRequeridaContacto " & _
                                             ", estado_respecto_movil = 'P' " & _
                                             ", estado_respecto_cliente = 'P' " & _
                                             " where nro_viaje = @NroViaje", DBConn)

            SqlUpdate.Parameters.AddWithNullableValue("@NroMovilReq", NroMovilRequerido)
            SqlUpdate.Parameters.AddWithNullableValue("@TipoServicio", TipoServicio)
            SqlUpdate.Parameters.AddWithNullableValue("@NroFlotaAsignada", NroFlotaAsignada)
            SqlUpdate.Parameters.AddWithNullableValue("@CategoriaMovil", CategoriaMovil)
            SqlUpdate.Parameters.AddWithNullableValue("@FechaHoraRequeridaContacto", FechaHoraRequeridaContacto)
            SqlUpdate.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

            SqlUpdate.CommandType = CommandType.Text
            Dim Regs = SqlUpdate.ExecuteNonQuery

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
    Public Sub ModificarDatosViaje(ByVal NroViaje As Integer, ByVal NroCliente As Integer, ByVal CodCalle As Integer, ByVal NombreCalle As String, ByVal Numero As Integer, ByVal Ubicacion As String, ByVal Latitud As Single, ByVal Longitud As Single, ByVal CodCalle1 As Integer, ByVal NombreCalle1 As String, ByVal CodCalle2 As Integer, ByVal NombreCalle2 As String, ByVal Comentario As String, ByVal ImportePactado As Decimal, ByVal ViajeUrgente As Boolean, ByVal Nombre As String, ByVal Observaciones As String, ByVal Direccion As String, ByVal CodParada As Integer, ByVal NroCtaCte As Integer, ByVal TipoServicio As Integer, ByVal CategoriaMovil As String, ByVal FechaHoraRequeridaContacto As Date)


        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlUpdate As New SqlCommand("update viajes set  nro_movil_requerido = @NroMovilReq" & _
                                                  ",nro_cliente =  @NroCliente" & _
                                                  ",cod_calle = @CodCalle" & _
                                                  ",nombre_calle = @NombreCalle" & _
                                                  ",numero = @Numero" & _
                                                  ",ubicacion = @Ubicacion" & _
                                                  ",latitud = @Latitud" & _
                                                  ",longitud = @Longitud" & _
                                                  ",cod_calle1 = @CodCalle1" & _
                                                  ",nombre_calle1 = @NombreCalle1" & _
                                                  ",cod_calle2 = @CodCalle2" & _
                                                  ",nombre_calle2 = @NombreCalle2" & _
                                                  ",comentario = @Comentario" & _
                                                  ",importe_pactado = @ImportePactado" & _
                                                  ",ind_viaje_urgente = @ViajeUrgente" & _
                                                  ",nombre = @Nombre" & _
                                                  ",observaciones = @Observaciones" & _
                                                  ",direccion = @Direccion" & _
                                                  ",cod_parada = @CodParada" & _
                                                  ",nro_ctacte = @NroCtaCte" & _
                                                  ",tipo_servicio_requerido = @TipoServicio" & _
                                                  ",nro_flota_asignada = @TipoServicio" & _
                                                  ", nro_movil_asignado = 0 " & _
                                                  ", categ_movil = @CategoriaMovil" & _
                                                  ",fecha_hora_requerida_contacto = @FechaHoraRequeridaContacto" & _
                                                  ",estado_respecto_movil = 'P' " & _
                                                  ", estado_respecto_cliente = 'P' " & _
                                                  " where nro_viaje =  @NroViaje", DBConn)

            SqlUpdate.Parameters.AddWithNullableValue("@NroCliente", NroCliente)
            SqlUpdate.Parameters.AddWithNullableValue("@CodCalle", CodCalle)
            SqlUpdate.Parameters.AddWithNullableValue("@NombreCalle", NombreCalle)
            SqlUpdate.Parameters.AddWithNullableValue("@Numero", Numero)
            SqlUpdate.Parameters.AddWithNullableValue("@Ubicacion", Ubicacion)
            SqlUpdate.Parameters.AddWithNullableValue("@Latitud", Latitud)
            SqlUpdate.Parameters.AddWithNullableValue("@Longitud", Longitud)
            SqlUpdate.Parameters.AddWithNullableValue("@CodCalle1", CodCalle1)
            SqlUpdate.Parameters.AddWithNullableValue("@NombreCalle1", NombreCalle1)
            SqlUpdate.Parameters.AddWithNullableValue("@CodCalle2", CodCalle2)
            SqlUpdate.Parameters.AddWithNullableValue("@NombreCalle2", NombreCalle2)
            SqlUpdate.Parameters.AddWithNullableValue("@Comentario", Comentario)
            SqlUpdate.Parameters.AddWithNullableValue("@ImportePactado", ImportePactado)
            SqlUpdate.Parameters.AddWithNullableValue("@ViajeUrgente", ViajeUrgente)
            SqlUpdate.Parameters.AddWithNullableValue("@Nombre", Nombre)
            SqlUpdate.Parameters.AddWithNullableValue("@Observacion", Observaciones)
            SqlUpdate.Parameters.AddWithNullableValue("@Direccion", Direccion)
            SqlUpdate.Parameters.AddWithNullableValue("@CodParada", CodParada)
            SqlUpdate.Parameters.AddWithNullableValue("@NroCtaCte", NroCtaCte)
            SqlUpdate.Parameters.AddWithNullableValue("@TipoServicio", TipoServicio)
            SqlUpdate.Parameters.AddWithNullableValue("@CategoriaMovil", CategoriaMovil)
            SqlUpdate.Parameters.AddWithNullableValue("@FechaHoraRequeridaContacto", FechaHoraRequeridaContacto)
            SqlUpdate.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

            SqlUpdate.CommandType = CommandType.Text
            Dim Regs = SqlUpdate.ExecuteNonQuery

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

    Public Sub CompletarPedido(ByVal NroViaje As Integer, ByVal NroMovilReq As Integer, ByVal NroUsuarioTelef As Integer, ByVal FechaEntrada As Date, ByVal NroCliente As Integer, ByVal CodCalle As Integer, ByVal NombreCalle As String, ByVal Numero As Integer, ByVal Ubicacion As String, ByVal CodLocalidad As Integer, ByVal Localidad As String, ByVal Latitud As Single, ByVal Longitud As Single, ByVal CodCalle1 As Integer, ByVal NombreCalle1 As String, ByVal CodCalle2 As Integer, ByVal NombreCalle2 As String, ByVal NroTelefono As Int64, ByVal Comentario As String, ByVal Direccion As String, ByVal CodZona As Integer, ByVal TipoServicio As Integer, ByVal CategoriaMovil As String, ByVal TipoViaje As Integer, ByVal EstadoRespectoCliente As String, ByVal EstadoRespectoMovil As String, ByVal ImportePactado As Decimal, ByVal ViajeUrgente As Boolean, ByVal Nombre As String, ByVal FechaHoraRequeridaContacto As Date, ByVal Observaciones As String, ByVal NroCtaCte As Integer, ByVal CodParada As Integer)


        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlUpdate As New SqlCommand("update viajes set nro_movil_requerido = @NroMovilReq " & _
                                            ",nro_usuario_telefonista = @NroUsuarioTelef" & _
                                            ",fecha_hora_entrada = @FechaEntrada " & _
                                            ",nro_cliente = @NroCliente " & _
                                            ",cod_calle = @CodCalle" & _
                                            ",nombre_calle = @NombreCalle " & _
                                            ",numero = @Numero " & _
                                            ",ubicacion = @Ubicacion" & _
                                            ",cod_localidad = @CodLocalidad " & _
                                            ",localidad = @Localidad " & _
                                            ",latitud = @Latitud" & _
                                            ",longitud = @Longitud " & _
                                            ",cod_calle1 = @CodCalle1 " & _
                                            ",nombre_calle1 = @NombreCalle1" & _
                                            ",cod_calle2 = @CodCalle2 " & _
                                            ",nombre_calle2 = @NombreCalle2 " & _
                                            ",nro_telefono = @NroTelefono" & _
                                            ",comentario = @Comentario " & _
                                            ",direccion = @Direccion " & _
                                            ",cod_zona = @CodZona" & _
                                            ",tipo_servicio_requerido = @TipoServicio " & _
                                            ",categ_movil = @CategoriaMovil" & _
                                            ",tipo_viaje = @TipoViaje " & _
                                            ",estado_respecto_cliente = @EstadoRespectoCliente" & _
                                            ",estado_respecto_movil = @EstadoRespectoMovil " & _
                                            ",importe_pactado = @ImportePactado" & _
                                            ",ind_viaje_urgente = @ViajeUrgente " & _
                                            ",nombre = @Nombre " & _
                                            ",fecha_hora_requerida_contacto = @FechaHoraRequeridaContacto" & _
                                            ",observaciones = @Observaciones " & _
                                            ",nro_subcta = 0 " & _
                                            ",lugar_liq = 0 " & _
                                            ",nro_ctacte =@NroCtaCte " & _
                                            ",cod_parada= @CodParada " & _
                                            ",where nro_viaje = @NroViaje", DBConn)

            SqlUpdate.Parameters.AddWithNullableValue("@NroMovilReq", NroMovilReq)
            SqlUpdate.Parameters.AddWithNullableValue("@NroUsuarioTelef", NroUsuarioTelef)
            SqlUpdate.Parameters.AddWithNullableValue("@FechaEntrada", FechaEntrada)
            SqlUpdate.Parameters.AddWithNullableValue("@NroCliente", NroCliente)
            SqlUpdate.Parameters.AddWithNullableValue("@CodCalle", CodCalle)
            SqlUpdate.Parameters.AddWithNullableValue("@NombreCalle", NombreCalle)
            SqlUpdate.Parameters.AddWithNullableValue("@Numero", Numero)
            SqlUpdate.Parameters.AddWithNullableValue("@Ubicacion", Ubicacion)
            SqlUpdate.Parameters.AddWithNullableValue("@CodLocalidad", CodLocalidad)
            SqlUpdate.Parameters.AddWithNullableValue("@Localidad", Localidad)
            SqlUpdate.Parameters.AddWithNullableValue("@Latitud", Latitud)
            SqlUpdate.Parameters.AddWithNullableValue("@Longitud", Longitud)
            SqlUpdate.Parameters.AddWithNullableValue("@CodCalle1", CodCalle1)
            SqlUpdate.Parameters.AddWithNullableValue("@NombreCalle1", NombreCalle1)
            SqlUpdate.Parameters.AddWithNullableValue("@CodCalle2", CodCalle2)
            SqlUpdate.Parameters.AddWithNullableValue("@NombreCalle2", NombreCalle2)
            SqlUpdate.Parameters.AddWithNullableValue("@NroTelefono", NroTelefono)
            SqlUpdate.Parameters.AddWithNullableValue("@Comentario", Comentario)
            SqlUpdate.Parameters.AddWithNullableValue("@Direccion", "<" & NroViaje.ToString & ">" & Direccion)
            SqlUpdate.Parameters.AddWithNullableValue("@CodZona", CodZona)
            SqlUpdate.Parameters.AddWithNullableValue("@TipoServicio", TipoServicio)
            SqlUpdate.Parameters.AddWithNullableValue("@CategoriaMovil", CategoriaMovil)
            SqlUpdate.Parameters.AddWithNullableValue("@TipoViaje", TipoViaje)
            SqlUpdate.Parameters.AddWithNullableValue("@EstadoRespectoCliente", EstadoRespectoCliente)
            SqlUpdate.Parameters.AddWithNullableValue("@EstadoRespectoMovil", EstadoRespectoMovil)
            SqlUpdate.Parameters.AddWithNullableValue("@ImportePactado", ImportePactado)
            SqlUpdate.Parameters.AddWithNullableValue("@ViajeUrgente", ViajeUrgente)
            SqlUpdate.Parameters.AddWithNullableValue("@Nombre", Nombre)
            SqlUpdate.Parameters.AddWithNullableValue("@FechaHoraRequeridaContacto", FechaHoraRequeridaContacto)
            SqlUpdate.Parameters.AddWithNullableValue("@Observaciones", Observaciones)
            SqlUpdate.Parameters.AddWithNullableValue("@NroCtaCte", NroCtaCte)
            SqlUpdate.Parameters.AddWithNullableValue("@CodParada", CodParada)
            SqlUpdate.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

            SqlUpdate.CommandType = CommandType.Text
            Dim Regs = SqlUpdate.ExecuteNonQuery

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
    Public Sub LlegoaDestino(ByVal NroViaje As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlUpdate As New SqlCommand("update viajes set estado_respecto_movil = 'D' where nro_viaje = @NroViaje", DBConn)
            SqlUpdate.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

            SqlUpdate.CommandType = CommandType.Text
            Dim Regs = SqlUpdate.ExecuteNonQuery

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


    Public Sub RegistrarRechazo(ByVal NroViaje As Integer, ByVal NroMovil As Integer, ByVal Causa As String, ByVal NroConductor As Integer, ByVal Latitud As Single, ByVal Longitud As Single)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlInsert As New SqlCommand("insert into VIAJES_RECHAZOS (IdMovil, FechaHoraRechazo, Latitud , longitud, IDConductor, IdViaje, Causa) values ( @NroMovil, getdate(), @Latitud, @Longitud, @NroConductor, @NroViaje, @Causa)", DBConn)
            SqlInsert.CommandType = CommandType.Text

            SqlInsert.Parameters.AddWithNullableValue("@NroViaje", NroViaje)
            SqlInsert.Parameters.AddWithNullableValue("@NroMovil", NroMovil)
            SqlInsert.Parameters.AddWithNullableValue("@NroConductor", NroConductor)
            SqlInsert.Parameters.AddWithNullableValue("@Latitud", Latitud)
            SqlInsert.Parameters.AddWithNullableValue("@Longitud", longitud)
            SqlInsert.Parameters.AddWithNullableValue("@Causa", Causa)

            Dim Regs = SqlInsert.ExecuteNonQuery

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


    Public Sub ActualizarNroCupon(ByVal NroViaje As Integer, ByVal NroCupon As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlUpdate As New SqlCommand("update Viajes set nro_subcta = @NroCupon where nro_viaje = @NroViaje", DBConn)
            SqlUpdate.Parameters.AddWithNullableValue("@NroViaje", NroViaje)
            SqlUpdate.Parameters.AddWithNullableValue("@NroSubCtaCte", NroCupon)

            SqlUpdate.CommandType = CommandType.Text
            Dim Regs = SqlUpdate.ExecuteNonQuery

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

    Public Function CrearViajeDeCalle(ByVal NroMovil As Integer?, ByVal FechaEntrada As Date?, ByVal Latitud As Single?, ByVal Longitud As Single?, ByVal CodZona As Integer?, ByVal TipoServicio As Integer?, ByVal Categoria As String, ByVal FechaHoraInicio As Date?, ByVal LatitudInicio As Single?, ByVal LongitudInicio As Single?, ByVal NroConductor As Integer?, ByVal KmrecorridosLibre As Integer?, ByVal VelocidadMaxLibre As Integer?, ByVal NroCorrelativoViaje As Integer?) As Integer

        Dim NroViaje As Integer

        Try

            Dim NuevoViaje = New DatosViaje


            NuevoViaje.NroViaje = 0
            NuevoViaje.NroMovilReq = 0
            NuevoViaje.NroUsuarioTelef = 0
            NuevoViaje.FechaEntrada = FechaEntrada   'Now
            NuevoViaje.NroCliente = 0
            NuevoViaje.CodCalle = 0
            NuevoViaje.NombreCalle = ""
            NuevoViaje.Numero = 0
            NuevoViaje.Ubicacion = ""
            NuevoViaje.CodLocalidad = 0
            NuevoViaje.Localidad = ""
            NuevoViaje.Latitud = Latitud
            NuevoViaje.Longitud = Longitud
            NuevoViaje.CodCalle1 = 0
            NuevoViaje.NombreCalle1 = ""
            NuevoViaje.CodCalle2 = 0
            NuevoViaje.NombreCalle2 = ""
            NuevoViaje.NroTelefono = 0
            NuevoViaje.Comentario = ""
            NuevoViaje.Direccion = ""
            NuevoViaje.CodZona = CodZona
            NuevoViaje.CodZonaOrigen = CodZona
            NuevoViaje.TipoServicio = TipoServicio
            NuevoViaje.CategoriaMovil = Categoria
            NuevoViaje.TipoViaje = "P"
            NuevoViaje.EstadoRespectoCliente = "I"
            NuevoViaje.FechaInicioViaje = FechaHoraInicio   'Now
            NuevoViaje.LatitudInicioViaje = LatitudInicio
            NuevoViaje.LongitudInicioViaje = LongitudInicio
            NuevoViaje.FechaFinViaje = Nothing
            NuevoViaje.LatitudFinViaje = 0
            NuevoViaje.LongitudFinViaje = 0
            NuevoViaje.CodMotivoCancel = 0
            NuevoViaje.EstadoRespectoMovil = "A"
            NuevoViaje.FechaInvitacion = Nothing
            NuevoViaje.LatitudInvitacion = 0
            NuevoViaje.LongitudInvitacion = 0
            NuevoViaje.FechaResolucion = Nothing
            NuevoViaje.LatitudResolucion = 0
            NuevoViaje.LongitudResolucion = 0
            NuevoViaje.NroMovilAsignado = NroMovil
            NuevoViaje.NroFlotaAsignada = TipoServicio
            NuevoViaje.Importe = 0
            NuevoViaje.ImportePactado = 0
            NuevoViaje.ImportePactadoTaxi = 0
            NuevoViaje.FinAnormal = False
            NuevoViaje.ViajeUrgente = False
            NuevoViaje.NroConductor = NroConductor
            NuevoViaje.NroCupon = 0
            NuevoViaje.CodZonaFutura = 0
            NuevoViaje.KmRecorridosLibre = KmrecorridosLibre
            NuevoViaje.VelocidadMaxLibre = VelocidadMaxLibre
            NuevoViaje.NroCorrLibre = NroCorrelativoViaje
            NuevoViaje.KmRecorridosOcupado = 0
            NuevoViaje.VelocidadMaxOcupado = 0
            NuevoViaje.MinCobradosEspera = 0
            NuevoViaje.TipoTarifa = 0
            NuevoViaje.NroCorrOcupado = 0
            NuevoViaje.Nombre = ""
            NuevoViaje.FechaHoraRequeridaContacto = Nothing
            NuevoViaje.NroUsuarioDespachador = 0

            NroViaje = CrearNuevo(NroMovil, NuevoViaje)

        Catch ex As Exception
            NroViaje = 0
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

        Return NroViaje

    End Function


    Public Function ObtenerViajeDesvinculado(ByVal Nromovil As Integer) As Integer?

        Dim NroViaje As Integer? = 0

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand(" select nro_viaje from viajes where estado_respecto_movil <> 'T' and nro_movil_asignado = @NroMovil", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@NroMovil", Nromovil)
            SqlCmd.CommandType = CommandType.Text

            NroViaje = SqlCmd.ExecuteScalar

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            Status = 1
            NroViaje = Nothing
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

        Return NroViaje


    End Function



    Public Function ObtenerViajeDesvinculadoHistorico(ByVal Nromovil As Integer) As Integer?

        Dim NroViaje As Integer? = 0

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("select top 1 nro_viaje from viajes_historicos where Importe = 0 and nro_movil_asignado = @NroMovil", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@NroMovil", Nromovil)
            SqlCmd.CommandType = CommandType.Text

            NroViaje = SqlCmd.ExecuteScalar

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            Status = 1
            NroViaje = Nothing
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

        Return NroViaje


    End Function

    Public Sub ActualizarViajeDesvinculadoHistorico(ByVal NroViaje As Integer, ByVal Importe As Decimal, ByVal KmRecorridosOcupado As Single, ByVal VelocidadMaximaOcupado As Integer, ByVal TipoTarifa As Integer, ByVal NroCorrelativoViaje As Integer, ByVal MinCobradosEspera As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlUpdate As New SqlCommand("update viajes_historicos set importe = @Importe, ind_fin_anormal = 0,  cod_motivo_cancelacion = 0, km_recorridos_ocupado = @KmRecorridosOcupado, velocidad_maxima_ocupado = @VelocMaximaOcupado, tipo_tarifa = @TipoTarifa, nro_correlativo_ocupado = @NroCorrelativoViaje, minutos_cobrados_espera = @MinCobradosEspera  where nro_viaje = @NroViaje", DBConn)
            SqlUpdate.Parameters.AddWithNullableValue("@NroViaje", NroViaje)
            SqlUpdate.Parameters.AddWithNullableValue("@Importe", Importe)
            SqlUpdate.Parameters.AddWithNullableValue("@KmRecorridosOcupado", KmRecorridosOcupado)
            SqlUpdate.Parameters.AddWithNullableValue("@VelocMaximaOcupado", VelocidadMaximaOcupado)
            SqlUpdate.Parameters.AddWithNullableValue("@TipoTarifa", TipoTarifa)
            SqlUpdate.Parameters.AddWithNullableValue("@NroCorrelativoViaje", NroCorrelativoViaje)
            SqlUpdate.Parameters.AddWithNullableValue("@MinCobradosEspera", MinCobradosEspera)

            SqlUpdate.CommandType = CommandType.Text
            Dim Regs = SqlUpdate.ExecuteNonQuery

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

    Public Function ArmarDireccion(ByRef DatosViaje As DatosViaje) As String
        Try

            ArmarDireccion = ""

            If Not String.IsNullOrEmpty(DatosViaje.NombreCalle) Then
                ArmarDireccion = DatosViaje.NombreCalle
            End If

            If DatosViaje.Numero <> 0 Then
                Dim NroAltura As Long = Int(DatosViaje.Numero / 100) * 100

                ArmarDireccion = DatosViaje.NombreCalle & " " & Str(NroAltura)
            End If

            If Not String.IsNullOrEmpty(DatosViaje.Ubicacion) Then
                ArmarDireccion = ArmarDireccion & " " & DatosViaje.Ubicacion
            End If

        Catch ex As Exception
            ArmarDireccion = Nothing
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

        Return ArmarDireccion
    End Function

    Public Function ArmarDireccionPush(ByRef DatosViaje As DatosViaje) As String
        Try

            ArmarDireccionPush = DatosViaje.NroViaje & ","

            If Not String.IsNullOrEmpty(DatosViaje.NombreCalle) And DatosViaje.NombreCalle <> " " Then
                ArmarDireccionPush = ArmarDireccionPush & Replace(DatosViaje.NombreCalle, ",", "")
            End If

            If DatosViaje.Numero <> 0 And DatosViaje.Numero <> -1 And DatosViaje.Numero <> -10 Then
                ArmarDireccionPush = ArmarDireccionPush & " " & DatosViaje.Numero
            Else
                If Not String.IsNullOrEmpty(DatosViaje.NombreCalle1) And DatosViaje.NombreCalle1 <> " " Then
                    ArmarDireccionPush = ArmarDireccionPush & " Y " & Replace(DatosViaje.NombreCalle1, ",", "")
                End If
            End If

            If Not String.IsNullOrEmpty(DatosViaje.Ubicacion) And DatosViaje.Ubicacion <> "" Then
                ArmarDireccionPush = ArmarDireccionPush & " - " & Replace(DatosViaje.Ubicacion, ",", "")
            End If

            If Not String.IsNullOrEmpty(DatosViaje.Comentario) And DatosViaje.Comentario <> "" Then
                ArmarDireccionPush = ArmarDireccionPush & " - " & Replace(DatosViaje.Comentario, ",", "") & ","
            Else
                ArmarDireccionPush = ArmarDireccionPush & ","
            End If

            'Inserta esquinas en comentario
            If DatosViaje.Numero <> 0 And DatosViaje.Numero <> -1 And DatosViaje.Numero <> -10 And DatosViaje.CodCalle = 0 Then
                ArmarDireccionPush = ArmarDireccionPush & " entre " & Replace(DatosViaje.NombreCalle1, ",", "") & " Y " & Replace(DatosViaje.NombreCalle2, ",", "")
            End If

            If Not IsNothing(DatosViaje.Nombre) And DatosViaje.Nombre <> "" Then
                ArmarDireccionPush = ArmarDireccionPush & "," & Replace(DatosViaje.Nombre, ",", "")
            Else
                ArmarDireccionPush = ArmarDireccionPush & ", -------"
            End If

            If Not IsNothing(DatosViaje.FechaHoraRequeridaContacto) And DatosViaje.FechaHoraRequeridaContacto <> "00:00:00" Then
                ArmarDireccionPush = ArmarDireccionPush & "," & Format(DatePart("H", DatosViaje.FechaHoraRequeridaContacto), "00") & ":" & Format(DatePart("n", DatosViaje.FechaHoraRequeridaContacto), "00")
            Else
                ArmarDireccionPush = ArmarDireccionPush & "," & Format(DatePart("H", DatosViaje.FechaEntrada), "00") & ":" & Format(DatePart("n", DatosViaje.FechaEntrada), "00")
            End If

            If Not IsNothing(DatosViaje.NroCtaCte) And DatosViaje.NroCtaCte <> 0 Then
                ArmarDireccionPush = ArmarDireccionPush & "," & DatosViaje.NroCtaCte & "-" & Replace(CuentaCorriente.ObtenerNombre(DatosViaje.NroCtaCte), ",", "")
            Else
                ArmarDireccionPush = ArmarDireccionPush & ", -------"
            End If

            If Not IsNothing(DatosViaje.ImportePactado) And DatosViaje.ImportePactado <> 0 Then
                ArmarDireccionPush = ArmarDireccionPush & "," & Replace(DatosViaje.ImportePactado, ",", ".")
            Else
                ArmarDireccionPush = ArmarDireccionPush & ", -------"
            End If

            ArmarDireccionPush = ArmarDireccionPush & ",------," & DatosViaje.Latitud & "," & DatosViaje.Longitud
            ArmarDireccionPush = ArmarDireccionPush & ",1"


        Catch ex As Exception
            ArmarDireccionPush = Nothing
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

        Return ArmarDireccionPush
    End Function


End Module
