Imports System.Configuration
Imports System.Web
Imports System.Xml
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic.Strings
Imports System.IO
Imports vb = Microsoft.VisualBasic

Public Class frmControl

    Friend ConnString As String
    Friend ConnStringMQTT As String

    Dim ConfigLogPath As String
    Dim ConfigLogFile As String
    Dim ConfigLogOverwrite As Boolean
    Dim ConfigNroCanalGPRS As Integer
    Dim ConfigServerGPRS As String
    Dim ConfigUserGPRS As String
    Dim ConfigPasswordGPRS As String
    Dim ConfigSincronizarHora As Boolean
    Dim ConfigTipoInforme As Boolean
    Dim ConfigGMTDif As Integer

    Dim gDatosZona As List(Of Zona.DatosZona)


    Dim ConfigAutoRun As Boolean

    Dim ServicioActivo As Boolean = False
    Dim DisplayStoped As Boolean = False

    Dim MaxEsperaMovilCandidato As Integer
    Dim ViajesPendInterval As Integer
    Dim MaxTiempoEnConsulta As Integer
    Dim ViajeComplejoFuturo As Integer
    Dim ModoSelecMoviles As Integer
    Dim TiempoAnticipViajeOcupado As Integer
    Dim TiempoEsperaFase10 As Integer
    Dim TiempoPausaViajesGPS As Integer
    Dim ViajesAgendadosInterval As Integer
    Dim FiltroVip As Integer
    Dim ViajeClientePreferencial As Integer
    Dim ViajeNormal As Integer

    Dim StopFlag As Boolean
    Dim EstaActivoTmrPendientes1 As Boolean
    Dim EstaActivoTmrPendientes2 As Boolean
    Dim EstaActivoTmrDespenalizar As Boolean
    Dim EstaActivoTmrDesactivar As Boolean
    Dim EstaActivoTmrAgendados As Boolean

    Dim SessionID As String
    Dim wsResponse As String

    Dim DBConn As SqlConnection
    Const LineaPuntos As String = "------------------------------------------------------------------------------------------"
    Dim ModoTransparente As Boolean

    Dim WebServiceURL As String
    Dim WebServiceUser As String
    Dim WebServicePassword As String

    Dim TiempoEsperaRespuestaInvitacion As Integer
    Dim TiempoPenalizacion As Integer

    Dim frecActualizacionRelojSistema As Integer
    Dim frecTimerActualizarPosicion As Integer
    Dim frecTimerComandosPend As Integer
    Dim frecTimerPenalizados As Integer
    Dim frecTimerComandosMC As Integer
    Dim frecTimerAlarmasCanceladas As Integer
    Dim frecTimerConsultaEstadoZona As Integer

    Private Sub frmControl_Load(sender As Object, e As EventArgs) Handles Me.Load

        Call LeerArchivoConfig()

        If String.IsNullOrEmpty(ConnString) Then
            MsgBox("ERROR: No se ha especificado la cadena de conexion con la base Autoflot. No puede iniciarse el programa.", MsgBoxStyle.Critical)
            End
        End If

        '-- asigna la cadena de conexión a las clases estáticas

        Alarma.ConnString = ConnString
        Base.ConnString = ConnString
        Comando.ConnString = ConnString
        CuentaCorriente.ConnString = ConnString
        Localidad.ConnString = ConnString
        MensajeMovil.ConnString = ConnString
        MensajeOper.ConnString = ConnString
        Movil.ConnString = ConnString
        Parametro.ConnString = ConnString
        Tarjeta.ConnString = ConnString
        UltimoNumero.ConnString = ConnString
        Viaje.ConnString = ConnString
        ViajeHistorico.ConnString = ConnString
        Zona.ConnString = ConnString
        AlarmaMovil.ConnString = ConnString
        MovilPush.ConnString = ConnString
        ViajeAgendado.ConnString = ConnString
        Despacho.ConnString = ConnString
        Cliente.ConnString = ConnString

        InicializarLog(ConfigLogFile, ConfigLogPath, Not (ConfigLogOverwrite))

        Me.Text = "Servidor de Despachos"
        Me.cmbModoDespacho.Enabled = True
        Me.chkConBases.Enabled = True
        Me.chkConBases.Checked = True
        Me.chkInicializarCanales.Enabled = True
        Me.cmbModoDespacho.SelectedIndex = 0
        Me.botAceptar.Text = "Iniciar Servicio"
        Me.botCancelar.Enabled = True
        Me.cmbModoDespacho.Focus()
        Me.chkStopList.Enabled = False
        Me.chkGrabarLog.Checked = True

        ServicioActivo = False

        'Leo la configuracion de Zonas Horario y Zonas

        gDatosZona = Zona.CargarCentrosZonas()

        If ConfigAutoRun = True Then
            ActualizarDisplay(0, "Serv DespachoAbierto por AutoRun", "Serv DespachoAbierto por AutoRun")
            Call IniciarServicio()
        End If

    End Sub

    Private Sub IniciarServicio()

        Try

            Me.botAceptar.Enabled = False
            Me.botCancelar.Enabled = False

            Me.Cursor = Cursors.WaitCursor()

            Call ActualizarDisplay(0, "Iniciando el servicio...", "Iniciando el servicio...")

            Call LeerConfiguraciones()

            '---- Activa los timers

            tmrViajesPendientes.Interval = ViajesPendInterval
            tmrViajesAgendados.Interval = 1
            tmrHoraAutoflot.Interval = 1
            tmrAlarmasCanceladas.Interval = 20000

            EstaActivoTmrPendientes1 = False
            EstaActivoTmrPendientes2 = False
            EstaActivoTmrDespenalizar = False
            EstaActivoTmrDesactivar = False
            EstaActivoTmrAgendados = False


            tmrViajesPendientes.Enabled = True
            tmrViajesAgendados.Enabled = True
            tmrHoraAutoflot.Enabled = True
            tmrAlarmasCanceladas.Enabled = True

            ActualizarDisplay(0, "Servicio iniciado ", "Servicio iniciado")

            Me.Text = "Servidor de Despacho - Activo - " & Now

            Me.chkInicializarCanales.Enabled = False
            Me.chkStopList.Enabled = True

            ServicioActivo = True

            Me.botAceptar.Text = "Detener Servicio"
            Me.botAceptar.Enabled = True
            Me.botCancelar.Enabled = True

            Me.Cursor = Cursors.Default

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            Me.chkInicializarCanales.Enabled = True
        End Try

    End Sub

    Private Sub FinalizarServicio()

        Try

            Me.Cursor = Cursors.WaitCursor()

            Me.botAceptar.Enabled = False
            Me.botCancelar.Enabled = False

            StopFlag = True

            Call ActualizarDisplay(0, "Deteniendo el servicio...", "")
            Me.tmrViajesPendientes.Enabled = False
            Me.tmrViajesPendientes2.Enabled = False
            Me.tmrHoraAutoflot.Enabled = False
            Me.tmrAlarmasCanceladas.Enabled = False

            Me.tmrDesactivar.Interval = 500
            Me.tmrDesactivar.Enabled = True

        Catch ex As Exception

            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            
        End Try

    End Sub

    Private Sub LeerConfiguraciones()


        ViajeComplejoFuturo = CType(Parametro.Leer("ViajeComplejoFut"), Integer)
        frecTimerConsultaEstadoZona = CType(Parametro.Leer("frec_cons_estado_zona"), Integer)
        frecTimerPenalizados = CType(Parametro.Leer("frec_control_penalizados"), Integer)
        MaxEsperaMovilCandidato = CType(Parametro.Leer("Tiem_Espera_Movil_Candidato"), Integer)
        TiempoEsperaFase10 = CType(Parametro.Leer("Tiem_Pausa_Desp_Viajes_GPS"), Integer)
        ViajesPendInterval = CType(Parametro.Leer("Frec_Exp_Viajes_Pend"), Integer)
        ViajesAgendadosInterval = CType(Parametro.Leer("Frec_Exp_Agenda_Viajes"), Integer)
        TiempoEsperaRespuestaInvitacion = CType(Parametro.Leer("tiem_esp_resp_invitacion"), Integer)
        ModoSelecMoviles = CType(Parametro.Leer("modo_seleccion_moviles"), Integer)
        frecActualizacionRelojSistema = CType(Parametro.Leer("frec_act_reloj_sistema"), Integer)
        TiempoAnticipViajeOcupado = CType(Parametro.Leer("tiem_anticip_viaje_ocupado"), Integer)
        TiempoPausaViajesGPS = CType(Parametro.Leer("Tiem_Pausa_Desp_Viajes_GPS"), Integer)

        'Borra e Inicializa Comandos 

        Comando.BorrarComandosTodos()
        MensajeOper.BorrarMensajesTodos()
        Viaje.BorrarViajesFactores()

       
    End Sub

    Public Sub ActualizarDisplay(ByVal NroMovil As Integer, ByVal DatosAsc As String, ByVal DatosHex As String)

        Dim strNroMovil As String
        Dim Hora As String = Format(Now, "hh:mm:ss  ")

        Select Case NroMovil
            Case -1
                strNroMovil = "[CEN ]"
            Case 0
                strNroMovil = "      "
            Case -2
                strNroMovil = "[ERR ]"
            Case Else
                strNroMovil = "[" & Format(NroMovil, "0000") & "]"

        End Select

        '-- si no se suministran los datos hexa, toma los datos asc para mostrar 

        If DatosHex = "" Then
            DatosHex = DatosAsc
        End If

        Dim LineaASC As String = Hora & strNroMovil & DatosAsc
        Dim LineaHEX As String = Hora & strNroMovil & DatosHex

        '-- graba el log si corresponde
        '-- los mensajes que tiene movil 0 se graban si o si en log porque son mensajes del programa y no de actividad del canal

        If GrabarLog Or NroMovil = 0 Then
            Log.Grabar(strNroMovil & ";" & DatosAsc & ";" & DatosHex)
        End If

        '-- si no está detenido el display actualiza las ventanas

        If Not (Me.chkStopList.Checked) Then

            If Me.lstDisplay.Items.Count > 150 Then
                Me.lstDisplay.Items.RemoveAt(0)
                Me.lstDisplay.Items.RemoveAt(0)
            End If

            Me.lstDisplay.Items.Add(LineaASC)
            Me.lstDisplay.TopIndex = Me.lstDisplay.TopIndex + 1
            Me.lstDisplay.SelectedIndex = lstDisplay.Items.Count - 1

        End If

    End Sub


    Private Sub LeerArchivoConfig()

        Try
            Dim xmldoc As New XmlDocument

            Dim fs As New FileStream(My.Application.Info.DirectoryPath & "\Archivos\config.xml", FileMode.Open, FileAccess.Read)
            xmldoc.Load(fs)

            ConnString = xmldoc.SelectSingleNode("/settings/setting[@key='ConnectionString']").Attributes("value").Value
            ConnStringMQTT = xmldoc.SelectSingleNode("/settings/setting[@key='ConnectionMQTT']").Attributes("value").Value
            ConfigLogPath = xmldoc.SelectSingleNode("/settings/setting[@key='LogPath']").Attributes("value").Value
            ConfigLogFile = xmldoc.SelectSingleNode("/settings/setting[@key='LogFile']").Attributes("value").Value
            ConfigLogOverwrite = xmldoc.SelectSingleNode("/settings/setting[@key='LogOverwrite']").Attributes("value").Value
            ConfigNroCanalGPRS = xmldoc.SelectSingleNode("/settings/setting[@key='NroCanalGPRS']").Attributes("value").Value
            ConfigServerGPRS = xmldoc.SelectSingleNode("/settings/setting[@key='ServerGPRS']").Attributes("value").Value
            ConfigUserGPRS = xmldoc.SelectSingleNode("/settings/setting[@key='UserGPRS']").Attributes("value").Value
            ConfigPasswordGPRS = xmldoc.SelectSingleNode("/settings/setting[@key='PasswordGPRS']").Attributes("value").Value
            ConfigSincronizarHora = xmldoc.SelectSingleNode("/settings/setting[@key='SincronizarHora']").Attributes("value").Value
            ConfigGMTDif = xmldoc.SelectSingleNode("/settings/setting[@key='GMTDif']").Attributes("value").Value
            ConfigTipoInforme = xmldoc.SelectSingleNode("/settings/setting[@key='TipoInforme']").Attributes("value").Value
            ConfigAutoRun = xmldoc.SelectSingleNode("/settings/setting[@key='AutoRun']").Attributes("value").Value


        Catch ex As Exception
            MsgBox("Error al leer el archivo CONFIG.XML" & vbNewLine & vbNewLine & ex.Message, MsgBoxStyle.Critical, "ERROR DE SISTEMA")

        End Try


    End Sub

    Private Sub botCancelar_Click(sender As Object, e As EventArgs) Handles botCancelar.Click
        Me.Close()
    End Sub

    Private Sub tmrViajesAgendados_Tick(sender As Object, e As EventArgs) Handles tmrViajesAgendados.Tick

        Static Minutos As Integer
        Dim IniPeriodo As Date = Now
        Dim FinPeriodo As Date = DateAdd("n", 120, IniPeriodo)

        If StopFlag = True Then
            Me.tmrViajesAgendados.Enabled = False
            Exit Sub
        End If

        EstaActivoTmrAgendados = True

        Me.tmrViajesAgendados.Enabled = False


        If Minutos >= ViajesAgendadosInterval Then

            Dim Dia = Weekday(Now)

        
            Try

                Call ActualizarDisplay(0, "Procesando viajes agendados... ", "")

                Dim DBConn = New SqlConnection(ConnString)
                DBConn.Open()

                Dim SqlCmd1 As New SqlCommand("Select * from  Viajes_Agendados left join Viajes_agendados_Items " & _
                                              " on Viajes_agendados.nro_viaje_agendado =Viajes_agendados_items.nro_viaje_agendado " & _
                                              "where Viajes_agendados_items.fecha_hora_viaje >= @IniPeriodo and " & _
                                              "Viajes_agendados_items.fecha_hora_viaje < @FinPeriodo and nro_viaje_generado =0", DBConn)

                SqlCmd1.Parameters.AddWithNullableValue("@IniPeriodo", IniPeriodo)
                SqlCmd1.Parameters.AddWithNullableValue("@FinPeriodo", FinPeriodo)
                SqlCmd1.CommandType = CommandType.Text

                Dim ViajesAgendados = SqlCmd1.ExecuteReader()

                Do While ViajesAgendados.Read

                    Dim bolCrearViaje As Boolean = False
                    Dim intDiaSemana As Integer = Weekday(ViajesAgendados("fecha_hora_viaje"))
                    Dim RegAgendados = New DatosViajeAgendado

                    Select Case ViajesAgendados("Opcion_Agenda")

                        Case 1  'Solo esta Vez

                            bolCrearViaje = True

                        Case 2  'Dias Habiles

                            If intDiaSemana >= 2 And intDiaSemana <= 6 And ViajeAgendado.EsHabil(ViajesAgendados("fecha_hora_viaje")) Then
                                bolCrearViaje = True
                            End If

                        Case 3  'Lunes A Viernes
                            If intDiaSemana >= 2 And intDiaSemana <= 6 Then
                                bolCrearViaje = True
                            End If

                        Case 4  'Todos Los Dias
                            bolCrearViaje = True

                        Case 5  'OTROS

                            If (ViajesAgendados("ind_Domingo") = True And intDiaSemana = 1) Or _
                                (ViajesAgendados("ind_Lunes") = True And intDiaSemana = 2) Or _
                                (ViajesAgendados("ind_Martes") = True And intDiaSemana = 3) Or _
                                (ViajesAgendados("ind_Miercoles") = True And intDiaSemana = 4) Or _
                                (ViajesAgendados("ind_Jueves") = True And intDiaSemana = 5) Or _
                                (ViajesAgendados("ind_Viernes") = True And intDiaSemana = 6) Or _
                                (ViajesAgendados("ind_Sabado") = True And intDiaSemana = 7) Then

                                bolCrearViaje = True

                            End If

                    End Select

                    'Determinamos la zona del viaje

                    Dim intZona As Integer

                    intZona = Zona.ObtenerZona(ViajesAgendados("Latitud"), ViajesAgendados("Longitud"))

                    'Determinamos el Tiempo de Anticipacion para la zona en ese momento

                    Dim TiempoAnticipLanzaAgendados = gDatosZona(intZona).Radios(3, 1)

                    If TiempoAnticipLanzaAgendados < DateDiff("n", Now, ViajesAgendados("Fecha_hora_Viaje")) Then
                        bolCrearViaje = False
                    End If

                    If bolCrearViaje = True Then
                        'Creo Los Viajes segun la cantidad
                        For i As Short = 1 To ViajesAgendados("Cant_Moviles")

                            Dim DatosViajeG As New DatosViaje

                            DatosViajeG.NroCliente = ViajesAgendados("Nro_Cliente")
                            DatosViajeG.CodCalle = ViajesAgendados("Cod_Calle")
                            DatosViajeG.NombreCalle = ViajesAgendados("Nombre_Calle")
                            DatosViajeG.Numero = ViajesAgendados("Numero")
                            DatosViajeG.Ubicacion = IsNull(ViajesAgendados("ubicacion"))
                            DatosViajeG.CodLocalidad = ViajesAgendados("cod_Localidad")
                            DatosViajeG.Localidad = IsNull(ViajesAgendados("Localidad"))
                            DatosViajeG.Latitud = ViajesAgendados("Latitud")
                            DatosViajeG.Longitud = ViajesAgendados("Longitud")
                            DatosViajeG.CodCalle1 = ViajesAgendados("Cod_Calle1")
                            DatosViajeG.NombreCalle1 = IsNull(ViajesAgendados("Nombre_Calle1"))
                            DatosViajeG.CodCalle2 = ViajesAgendados("Cod_Calle2")
                            DatosViajeG.NombreCalle2 = IsNull(ViajesAgendados("Nombre_Calle2"))
                            DatosViajeG.NroTelefono = ViajesAgendados("Nro_Telefono")
                            DatosViajeG.Comentario = IsNull(ViajesAgendados("Comentario"))
                            DatosViajeG.Direccion = IsNull(ViajesAgendados("Direccion"))
                            DatosViajeG.CodZona = intZona
                            DatosViajeG.TipoServicio = ViajesAgendados("Tipo_Servicio_Requerido")
                            DatosViajeG.ImportePactado = ViajesAgendados("Importe_Pactado")
                            DatosViajeG.NroFlotaAsignada = ViajesAgendados("Tipo_Servicio_Requerido")
                            DatosViajeG.Nombre = IsNull(ViajesAgendados("Nombre"))
                            DatosViajeG.FechaHoraRequeridaContacto = ViajesAgendados("Fecha_hora_viaje")
                            DatosViajeG.NroUsuarioTelef = ViajesAgendados("Nro_Usuario_Telefonista")
                            'Aca Habria que recalcular la Base
                            DatosViajeG.CodParada = intZona
                            DatosViajeG.NroCtaCte = ViajesAgendados("Nro_CtaCte")
                            DatosViajeG.NroSubCta = ViajesAgendados("Nro_SubCta")
                            DatosViajeG.NroMovilReq = ViajesAgendados("Nro_Movil_Requerido")
                            'Datos Fijos
                            DatosViajeG.FechaEntrada = Now
                            DatosViajeG.CodZonaOrigen = 0
                            DatosViajeG.TipoViaje = "O"
                            DatosViajeG.EstadoRespectoCliente = "P"
                            DatosViajeG.EstadoRespectoMovil = "P"
                            DatosViajeG.CategoriaMovil = "5"
                            DatosViajeG.FechaInicioViaje = Nothing
                            DatosViajeG.LatitudInicioViaje = 0
                            DatosViajeG.LongitudInicioViaje = 0
                            DatosViajeG.FechaFinViaje = Nothing
                            DatosViajeG.LatitudFinViaje = 0
                            DatosViajeG.LongitudFinViaje = 0
                            DatosViajeG.FechaInicioViaje = Nothing
                            DatosViajeG.LatitudInicioViaje = 0
                            DatosViajeG.CodMotivoCancel = 0
                            DatosViajeG.FechaInvitacion = Nothing
                            DatosViajeG.LatitudInvitacion = 0
                            DatosViajeG.LongitudInvitacion = 0
                            DatosViajeG.FechaResolucion = Nothing
                            DatosViajeG.LatitudResolucion = 0
                            DatosViajeG.LongitudResolucion = 0
                            DatosViajeG.NroMovilAsignado = -99
                            DatosViajeG.Importe = 0
                            DatosViajeG.ImportePactadoTaxi = 0
                            DatosViajeG.FinAnormal = False
                            DatosViajeG.ViajeUrgente = True
                            DatosViajeG.NroConductor = 0
                            DatosViajeG.CodZonaFutura = 0
                            DatosViajeG.NroCupon = 0
                            DatosViajeG.VelocidadMaxLibre = 0
                            DatosViajeG.VelocidadMaxOcupado = 0
                            DatosViajeG.NroCorrLibre = 0
                            DatosViajeG.NroCorrOcupado = 0
                            DatosViajeG.KmRecorridosLibre = 0
                            DatosViajeG.KmRecorridosOcupado = 0
                            DatosViajeG.MinCobradosEspera = 0
                            DatosViajeG.TipoTarifa = 0
                            DatosViajeG.NroUsuarioDespachador = 0
                            DatosViajeG.Observaciones = ""

                            Dim IdViaje As Long = Viaje.CrearNuevo(ViajesAgendados("Nro_Movil_requerido"), DatosViajeG)

                            'Falta Marcar el viaje como Procesado

                            ViajeAgendado.MarcarComoGenerado(IdViaje, ViajesAgendados("nro_viaje_agendado"), ViajesAgendados("Fecha_hora_Viaje"))

                            Call ActualizarDisplay(0, "Disparo Viaje Agendado:" & ViajesAgendados("nro_viaje_agendado") & " " & ViajesAgendados("Direccion") & " - IdViaje Generado:" & IdViaje, "")


                        Next i

                    End If
                    Application.DoEvents()
                Loop
                ViajesAgendados.Close()

                If DBConn.State = ConnectionState.Open Then
                    DBConn.Close()
                End If

                DBConn.Dispose()

                Call ActualizarDisplay(0, "Fin de Proceso de viajes Agendados", "")
                Minutos = -1

            Catch ex As Exception
                Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            End Try

        End If

        Minutos = Minutos + 1
        EstaActivoTmrAgendados = False
        Me.tmrViajesAgendados.Enabled = True
        Me.tmrViajesAgendados.Interval = 60000
        Application.DoEvents()

    End Sub

    Private Sub tmrViajesPendientes_Tick(sender As Object, e As EventArgs) Handles tmrViajesPendientes.Tick

        Dim bolDispararViaje As Boolean
        Dim bolEsViajeDiferido As Boolean
        Dim intDesfasaComando As Integer = 0
        Dim MovilElegido As New MejorMovil
        Me.tmrViajesPendientes.Enabled = False
        Dim DBConn = New SqlConnection(ConnString)

        Try

            DBConn.Open()

            Call ActualizarDisplay(0, "Procesando viajes Pendientes... ", "")

            Dim SqlCmd1 As New SqlCommand("Select * from  Viajes where estado_respecto_movil = 'P' and estado_respecto_cliente = 'P' " & _
                                                                  "and (nro_movil_asignado =0 or nro_movil_asignado =-99) " & _
                                                                  " order by ind_fin_anormal, nro_viaje", DBConn)

            SqlCmd1.CommandType = CommandType.Text

            Dim ViajesPendientes = SqlCmd1.ExecuteReader()

            Do While ViajesPendientes.Read
                'verifica si corresponde Disparar el Viaje
                bolEsViajeDiferido = False

                If Not IsDate(ViajesPendientes("fecha_hora_requerida_contacto")) Then
                    bolDispararViaje = True
                Else
                    ' es un viaje comun
                    If CDate(ViajesPendientes("fecha_hora_requerida_contacto")) < DateAdd("n", CInt(gDatosZona(ViajesPendientes("Cod_Zona")).Radios(3, 2)), Now) Then
                        bolDispararViaje = True
                    Else
                        bolDispararViaje = False
                        bolEsViajeDiferido = True
                    End If
                End If

                If bolDispararViaje = True Then
                    'Se Trata de un viajeNO Diferido

                    ' Si el Nro_movil_asigando = 0 entonces es un viaje pendiente
                    ' Si el Nro_movil_asigando = 0 entonces es un viaje Agendado, que procesa primero Futuros
                    Debug.Print(ViajesPendientes("nro_viaje"))
                    Select Case ViajesPendientes("nro_movil_asignado")

                        Case 0
                            'Busca un Movil libre en la Parada

                            'Dim RadioBusqBase = gDatosZona(ViajesPendientes("cod_zona")).Radios(3, 4)
                            Dim RadioBusqBase = Zona.ObtieneRadioBusquedaBase(ViajesPendientes("cod_zona"), ViajesPendientes("tipo_servicio_requerido"))
                            MovilElegido = Despacho.BuscarMovilLibreParada(ViajesPendientes("Latitud"), ViajesPendientes("Longitud"), ViajesPendientes("cod_Parada"), RadioBusqBase, ViajesPendientes("categ_movil"), ViajesPendientes("tipo_servicio_requerido"), 0)

                            If MovilElegido.IdMovilOptimo <> 0 Then
                                'Encontro Movil en Base
                                Dim DatosMovil = Movil.ObtenerDatos(MovilElegido.IdMovilOptimo)

                                If DatosMovil.EstadoOper = "L" Then
                                    'Pone al movil en Consulta
                                    Viaje.PonerEnConsulta(ViajesPendientes("nro_viaje"), MovilElegido.IdMovilOptimo, DatosMovil.NroFlota, DatosMovil.Latitud, DatosMovil.Longitud, DatosMovil.CodZonaActual, MovilElegido.DistanciaMovilOptimo)
                                    Movil.InvitarViajeActual(MovilElegido.IdMovilOptimo, ViajesPendientes("nro_viaje"))
                                    Call ActualizarDisplay(0, "+++ Viaje " & ViajesPendientes("nro_viaje") & "en Consulta Libre Parada " & ViajesPendientes("cod_Parada") & "Movil:" & MovilElegido.IdMovilOptimo, "")
                                    'Envia al Movil el Comando de Invitacion
                                    Comando.EnviarComando(DatosMovil.NroCanal, MovilElegido.IdMovilOptimo, "2A", ViajesPendientes("nro_viaje"), False, 3, 8)

                                End If

                            Else
                                'No se Encontro Movil en la Parada, Buscamos el movil mas cercano

                                'Dim DistanciaBusqueda = CLng(gDatosZona(ViajesPendientes("cod_zona")).Radios(3, 0)) * 10
                                Dim DistanciaBusqueda = Zona.ObtieneRadioBusquedaZona(ViajesPendientes("cod_zona"), ViajesPendientes("tipo_servicio_requerido"))
                                MovilElegido = Despacho.BuscarMovilLibreDistancia(ViajesPendientes("Latitud"), ViajesPendientes("Longitud"), DistanciaBusqueda, 0, ViajesPendientes("categ_movil"), ViajesPendientes("tipo_servicio_requerido"))
                                'SACAR
                                MovilElegido.IdMovilOptimo = 0
                                If MovilElegido.IdMovilOptimo <> 0 Then
                                    'Encontro Movil Libre dentro del radio buscado
                                    Dim DatosMovil = Movil.ObtenerDatos(MovilElegido.IdMovilOptimo)

                                    If DatosMovil.EstadoOper = "L" And DatosMovil.EstadoAdmin = "H" Then
                                        'Pone al movil en Consulta
                                        Viaje.PonerEnConsulta(ViajesPendientes("nro_viaje"), MovilElegido.IdMovilOptimo, DatosMovil.NroFlota, DatosMovil.Latitud, DatosMovil.Longitud, DatosMovil.CodZonaActual, MovilElegido.DistanciaMovilOptimo)
                                        Movil.InvitarViajeActual(MovilElegido.IdMovilOptimo, ViajesPendientes("nro_viaje"))
                                        Call ActualizarDisplay(0, "+++ Viaje " & ViajesPendientes("nro_viaje") & "en Consulta Libre - Movil:" & MovilElegido.IdMovilOptimo, "")
                                        'Envia al Movil el Comando de Invitacion
                                        Comando.EnviarComando(DatosMovil.NroCanal, MovilElegido.IdMovilOptimo, "2A", ViajesPendientes("nro_viaje"), False, 3, 8)
                                    Else
                                        'Aca habia elegido movil pero no estaba libre o no estaba habilitado
                                        Call ActualizarDisplay(0, "+++ Viaje " & ViajesPendientes("nro_viaje") & " Elegido Movil:" & MovilElegido.IdMovilOptimo & " No despachado estado:" & DatosMovil.EstadoOper, "")
                                    End If
                                Else
                                    ' No encontro por Distancia
                                    'Busco Por Futuros que se dirijen a la zona
                                    'Dim TiempoAnticipacion = CInt(gDatosZona(ViajesPendientes("cod_zona")).Radios(3, 5))
                                    Dim TiempoAnticipacion = Zona.ObtieneTiempoBusquedaFuturo(ViajesPendientes("cod_zona"), ViajesPendientes("tipo_servicio_requerido"))
                                    MovilElegido = BuscarMovilOcupadoAZona(ViajesPendientes("Latitud"), ViajesPendientes("Longitud"), 0, ViajesPendientes("categ_movil"), ViajesPendientes("tipo_servicio_requerido"), TiempoAnticipacion, ViajesPendientes("Cod_Zona"))

                                    If MovilElegido.IdMovilOptimo <> 0 Then
                                        'Encontro Movil Ocupado
                                        Dim DatosMovil = Movil.ObtenerDatos(MovilElegido.IdMovilOptimo)

                                        If DatosMovil.EstadoOper = "O" And DatosMovil.EstadoAdmin = "H" Then
                                            'Pone al movil en Consulta
                                            Viaje.PonerEnConsulta(ViajesPendientes("nro_viaje"), MovilElegido.IdMovilOptimo, DatosMovil.NroFlota, DatosMovil.Latitud, DatosMovil.Longitud, DatosMovil.CodZonaActual, MovilElegido.DistanciaMovilOptimo)
                                            Movil.InvitarViajeFuturo(MovilElegido.IdMovilOptimo, ViajesPendientes("nro_viaje"))
                                            Call ActualizarDisplay(0, "+++ Viaje " & ViajesPendientes("nro_viaje") & "en Consulta Futuro - Movil:" & MovilElegido.IdMovilOptimo, "")
                                            'Envia al Movil el Comando de Invitacion
                                            Comando.EnviarComando(DatosMovil.NroCanal, MovilElegido.IdMovilOptimo, "2A", ViajesPendientes("nro_viaje"), False, 3, 8)
                                        Else
                                            'Aca habia elegido movil pero no estaba libre o no estaba habilitado
                                            Call ActualizarDisplay(0, "+++ Viaje " & ViajesPendientes("nro_viaje") & " Elegido Movil Futuro:" & MovilElegido.IdMovilOptimo & " No despachado estado:" & DatosMovil.EstadoOper, "")
                                        End If

                                    Else
                                        'No encontro por Futuro
                                        'Busco por Zonas Adyacentes

                                        Dim intDemora = DateDiff("s", ViajesPendientes("fecha_hora_entrada"), Now)

                                        If intDemora > 60 Then

                                            MovilElegido = BuscarMovilZonaAdyacentes(ViajesPendientes("Cod_Zona"), 0, ViajesPendientes("categ_movil"), ViajesPendientes("tipo_servicio_requerido"))
                                            'SACAR
                                            MovilElegido.IdMovilOptimo = 0
                                            If MovilElegido.IdMovilOptimo <> 0 Then
                                                'Encontro Movil Libre en Zona Propia o Adyacente
                                                Dim DatosMovil = Movil.ObtenerDatos(MovilElegido.IdMovilOptimo)

                                                If DatosMovil.EstadoOper = "L" And DatosMovil.EstadoAdmin = "H" Then
                                                    'Pone al movil en Consulta
                                                    Viaje.PonerEnConsulta(ViajesPendientes("nro_viaje"), MovilElegido.IdMovilOptimo, DatosMovil.NroFlota, DatosMovil.Latitud, DatosMovil.Longitud, DatosMovil.CodZonaActual, MovilElegido.DistanciaMovilOptimo)
                                                    Movil.InvitarViajeActual(MovilElegido.IdMovilOptimo, ViajesPendientes("nro_viaje"))
                                                    Call ActualizarDisplay(0, "+++ Viaje " & ViajesPendientes("nro_viaje") & "en Consulta Ady - Movil:" & MovilElegido.IdMovilOptimo, "")
                                                    'Envia al Movil el Comando de Invitacion
                                                    Comando.EnviarComando(DatosMovil.NroCanal, MovilElegido.IdMovilOptimo, "2A", ViajesPendientes("nro_viaje"), False, 3, 8)
                                                Else
                                                    'Aca habia elegido movil pero no estaba libre o no estaba habilitado
                                                    Call ActualizarDisplay(0, "+++ Viaje " & ViajesPendientes("nro_viaje") & " Elegido Movil Ady:" & MovilElegido.IdMovilOptimo & " No despachado estado:" & DatosMovil.EstadoOper, "")
                                                End If
                                            Else
                                                ' No encontro Moviles por ninguna busqueda
                                                'Busca Postulantes a Viajes Pendientes

                                                MovilElegido = Despacho.BuscarMovilPostulante(ViajesPendientes("nro_viaje"), ViajesPendientes("latitud"), ViajesPendientes("Longitud"), 0, ViajesPendientes("categ_movil"), ViajesPendientes("tipo_servicio_requerido"))

                                                If MovilElegido.IdMovilOptimo <> 0 Then
                                                    'Encontro Movil Libre dentro del radio buscado
                                                    Dim DatosMovil = Movil.ObtenerDatos(MovilElegido.IdMovilOptimo)

                                                    If DatosMovil.EstadoOper = "L" And DatosMovil.EstadoAdmin = "H" Then
                                                        'Pone al movil en Consulta
                                                        Viaje.PonerEnConsulta(ViajesPendientes("nro_viaje"), MovilElegido.IdMovilOptimo, DatosMovil.NroFlota, DatosMovil.Latitud, DatosMovil.Longitud, DatosMovil.CodZonaActual, MovilElegido.DistanciaMovilOptimo)
                                                        Movil.InvitarViajeActual(MovilElegido.IdMovilOptimo, ViajesPendientes("nro_viaje"))
                                                        Call ActualizarDisplay(0, "+++ Viaje " & ViajesPendientes("nro_viaje") & "en Consulta Postulante - Movil:" & MovilElegido.IdMovilOptimo, "")
                                                        'Envia al Movil el Comando de Invitacion
                                                        Comando.EnviarComando(DatosMovil.NroCanal, MovilElegido.IdMovilOptimo, "2A", ViajesPendientes("nro_viaje"), False, 3, 8)
                                                    Else
                                                        'Aca habia elegido movil pero no estaba libre o no estaba habilitado
                                                        Call ActualizarDisplay(0, "+++ Viaje " & ViajesPendientes("nro_viaje") & " Elegido Movil Postulante:" & MovilElegido.IdMovilOptimo & " No despachado estado:" & DatosMovil.EstadoOper, "")
                                                    End If


                                                Else
                                                    Call ActualizarDisplay(0, "+++ Viaje " & ViajesPendientes("nro_viaje") & " Sin Moviles Disponibles", "")

                                                End If
                                            End If
                                        End If   'busqueda de Adyacentes

                                    End If     'futuros
                                End If         'Distancia
                            End If             'Base
                    End Select

                End If
                Application.DoEvents()
            Loop

            ViajesPendientes.Close()

            Call ActualizarDisplay(0, "Fin Proceso viajes Pendientes... ", "")

            'Ahora verifico los viajes que estan en Consulta
            Call ActualizarDisplay(0, "Procesando viajes En Consulta... ", "")

            Dim SqlCmd2 As New SqlCommand("Select * from Viajes where estado_respecto_movil = 'C'", DBConn)

            SqlCmd2.CommandType = CommandType.Text

            Dim ViajesConsulta = SqlCmd2.ExecuteReader()

            Do While ViajesConsulta.Read

                Dim TiempoEnConsulta = DateDiff("s", ViajesConsulta("fecha_hora_invitacion"), Now)

                If TiempoEnConsulta > MaxTiempoEnConsulta Then

                    Dim DatosMovil = Movil.ObtenerDatos(ViajesConsulta("nro_movil_asignado"))

                    If DatosMovil.EstadoOper = "I" Then
                        Movil.DemorarActual(DatosMovil.NroMovil)
                        'Registrar Rechazo Viaje en tabla Movil Rechazo
                        Viaje.RegistrarRechazo(ViajesConsulta("nro_viaje"), DatosMovil.NroMovil, "D", DatosMovil.NroConductor, DatosMovil.Latitud, DatosMovil.Longitud)
                    Else
                        If DatosMovil.EstadoFuturo = "I" Then
                            Movil.DemorarFuturo(DatosMovil.NroMovil)

                            'Registrar Rechazo Viaje en tabla Movil Rechazo
                        End If
                    End If
                    'redespacha el viaje
                    Viaje.Redespachar(ViajesConsulta("nro_viaje"))
                    Call ActualizarDisplay(0, "Viaje " & ViajesConsulta("nro_viaje") & " Redespachado por Falta de Respuesta del Movil ", "")
                End If
                Application.DoEvents()
            Loop

            ViajesConsulta.Close()

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

        If DBConn.State = ConnectionState.Open Then
            DBConn.Close()
        End If

        DBConn.Dispose()

        Me.tmrViajesPendientes.Enabled = True

    End Sub

    Private Sub tmrHoraAutoflot_Tick(sender As Object, e As EventArgs) Handles tmrHoraAutoflot.Tick

        Static SegundosReloj As Integer
        Static SegundosHorario As Integer
        Dim DBConn = New SqlConnection(ConnString)
        Me.tmrHoraAutoflot.Enabled = False

        Try
            If SegundosReloj > frecActualizacionRelojSistema Then
                'Deberia sincronizar Hora con el Server
                SegundosReloj = 0
                Call ActualizarDisplay(0, "Se Actualizo la Hora del Sistema (OJO)... ", "")
            End If

            If SegundosHorario = 60 Then

                Zona.ActualizaHorarioSistema()
                Call ActualizarDisplay(0, "Se Actualizo El Codigo de Horario de Sistema... ", "")
                SegundosHorario = 0
            End If

            SegundosReloj = SegundosReloj + 1
            SegundosHorario = SegundosHorario + 1

            'Ahora Atiende Los Pedidos de IVR

            Call ActualizarDisplay(0, "Procesando Pedidos IVR... ", "")

            DBConn.Open()
            Dim SqlCmd1 As New SqlCommand("Select * from  IVRInterfaz where indTomado =0", DBConn)

            SqlCmd1.CommandType = CommandType.Text

            Dim ViajesIVR = SqlCmd1.ExecuteReader()

            Do While ViajesIVR.Read

                'levanto los datos del Cliente
                Dim DatosCliente = Cliente.ObtenerDatosxTelefono(CType(ViajesIVR("Telefono"), String))

                'Con esto Creo el viaje

                Dim DatosViajeG As New DatosViaje

                DatosViajeG.NroCliente = DatosCliente.NroCliente
                DatosViajeG.CodCalle = DatosCliente.CodCalle
                DatosViajeG.NombreCalle = DatosCliente.NombreCalle
                DatosViajeG.Numero = DatosCliente.Numero
                DatosViajeG.Ubicacion = DatosCliente.Ubicacion
                DatosViajeG.CodLocalidad = DatosCliente.CodLocalidad
                DatosViajeG.Localidad = DatosCliente.Localidad
                DatosViajeG.Latitud = DatosCliente.Latitud
                DatosViajeG.Longitud = DatosCliente.Longitud
                DatosViajeG.CodCalle1 = DatosCliente.CodCalle1
                DatosViajeG.NombreCalle1 = DatosCliente.NombreCalle1
                DatosViajeG.CodCalle2 = DatosCliente.CodCalle2
                DatosViajeG.NombreCalle2 = DatosCliente.NombreCalle2
                DatosViajeG.NroTelefono = DatosCliente.Telefono
                DatosViajeG.Comentario = ""
                DatosViajeG.Direccion = Format(Hour(TimeOfDay), "00") & ":" & Format(Minute(TimeOfDay), "00") & " " & DatosCliente.Domicilio
                DatosViajeG.CodZona = DatosCliente.CodZona
                DatosViajeG.TipoServicio = 3
                DatosViajeG.ImportePactado = 0
                DatosViajeG.NroFlotaAsignada = 3
                DatosViajeG.Nombre = DatosCliente.Nombre
                DatosViajeG.FechaHoraRequeridaContacto = Nothing
                DatosViajeG.NroUsuarioTelef = 99
                'Aca Habria que recalcular la Base
                DatosViajeG.CodParada = DatosCliente.CodZona 'DatosCliente.CodBase 
                DatosViajeG.NroCtaCte = 0   'DatosCliente.NroCtaCte
                DatosViajeG.NroSubCta = 0   'DatosCliente.NroSubCta
                DatosViajeG.NroMovilReq = 0
                'Datos Fijos
                DatosViajeG.FechaEntrada = Now
                DatosViajeG.CodZonaOrigen = 0
                DatosViajeG.TipoViaje = "O"
                DatosViajeG.EstadoRespectoCliente = "P"
                DatosViajeG.EstadoRespectoMovil = "P"
                DatosViajeG.FechaInicioViaje = Nothing
                DatosViajeG.LatitudInicioViaje = 0
                DatosViajeG.LongitudInicioViaje = 0
                DatosViajeG.FechaFinViaje = Nothing
                DatosViajeG.LatitudFinViaje = 0
                DatosViajeG.LongitudFinViaje = 0
                DatosViajeG.FechaInicioViaje = Nothing
                DatosViajeG.LatitudInicioViaje = 0
                DatosViajeG.CodMotivoCancel = 0
                DatosViajeG.FechaInvitacion = Nothing
                DatosViajeG.LatitudInvitacion = 0
                DatosViajeG.LongitudInvitacion = 0
                DatosViajeG.FechaResolucion = Nothing
                DatosViajeG.LatitudResolucion = 0
                DatosViajeG.LongitudResolucion = 0
                DatosViajeG.NroMovilAsignado = 0
                DatosViajeG.Importe = 0
                DatosViajeG.ImportePactadoTaxi = 0
                DatosViajeG.FinAnormal = False
                DatosViajeG.ViajeUrgente = True
                DatosViajeG.NroConductor = 0
                DatosViajeG.CodZonaFutura = 0
                DatosViajeG.NroCupon = 0
                DatosViajeG.VelocidadMaxLibre = 0
                DatosViajeG.VelocidadMaxOcupado = 0
                DatosViajeG.NroCorrLibre = 0
                DatosViajeG.NroCorrOcupado = 0
                DatosViajeG.KmRecorridosLibre = 0
                DatosViajeG.KmRecorridosOcupado = 0
                DatosViajeG.MinCobradosEspera = 0
                DatosViajeG.TipoTarifa = 0
                DatosViajeG.NroUsuarioDespachador = 0
                DatosViajeG.Observaciones = ""
                DatosViajeG.CategoriaMovil = "5"

                Dim IdViaje As Long = Viaje.CrearNuevo(0, DatosViajeG)

                'Falta Marcar el viaje como Procesado

                ViajeAgendado.MarcarIVRComoGenerado(ViajesIVR("idIVR"), IdViaje)

                Call ActualizarDisplay(0, "Inserta Viaje IVR :" & DatosCliente.Domicilio & " IdIvr;" & ViajesIVR("idIVR") & "- IdViaje Generado:" & IdViaje, "")
                Application.DoEvents()
            Loop

        Catch ex As Exception

        End Try

        Me.tmrHoraAutoflot.Enabled = True
        Me.tmrHoraAutoflot.Interval = 2000

    End Sub

    Private Sub tmrDesactivar_Tick(sender As Object, e As EventArgs) Handles tmrDesactivar.Tick

        tmrDesactivar.Enabled = False

        Call DesactivarServidor()

        End

        'tmrDesactivar.Enabled = True
    End Sub

    Private Sub botAceptar_Click(sender As Object, e As EventArgs) Handles botAceptar.Click
        tmrDesactivar.Enabled = False
    End Sub

    Private Sub DesactivarServidor()

        Dim n As String = MsgBox("Realmente Desea Finalizar la Aplicacion?", MsgBoxStyle.YesNo, "Finalizar Aplicacion")
        If n = vbYes Then
            ActualizarDisplay(0, "Cerrando el Servidor...", "")

            Me.tmrHoraAutoflot.Interval = 99999999
            Me.tmrViajesAgendados.Interval = 99999999
            Me.tmrViajesPendientes2.Interval = 99999999
            Me.tmrAlarmasCanceladas.Interval = 99999999

            Me.tmrHoraAutoflot.Stop()
            Me.tmrViajesAgendados.Stop()
            Me.tmrViajesPendientes2.Stop()
            Me.tmrAlarmasCanceladas.Stop()

            tmrHoraAutoflot.Dispose()
            Me.tmrViajesAgendados.Dispose()
            Me.tmrViajesPendientes2.Dispose()
            Me.tmrAlarmasCanceladas.Dispose()

            Me.tmrDesactivar.Interval = 5000
            Me.tmrDesactivar.Start()

        End If


    End Sub


    Private Sub tmrAlarmasCanceladas_Tick(sender As Object, e As EventArgs) Handles tmrAlarmasCanceladas.Tick

        Me.tmrAlarmasCanceladas.Enabled = False

        Try

        
        Dim DBConn = New SqlConnection(ConnString)

        DBConn.Open()

        Using DBConn

                Call ActualizarDisplay(0, "Procesando Alarmas Moviles... ", "")

                Dim SqlCmd1 As New SqlCommand("Select * from Alarmas where idAlarma in (Select Max (idAlarma) from Alarmas where " & _
                                              " Fecha_Hora_Fin > dateadd(s, -600, getdate()) " & _
                                              " and Estado_alarma = 'N' group by nro_movil)", DBConn)
                SqlCmd1.CommandType = CommandType.Text

                Dim Alarmas = SqlCmd1.ExecuteReader()

                Do While Alarmas.Read

                    'procesa cada movil y decide qué acción tomar

                    Dim DatosMovil = Movil.ObtenerDatos(Alarmas("nro_Movil"))

                    'Consulta Posicion
                    Comando.EnviarComando(DatosMovil.NroCanal, Alarmas("nro_Movil"), "17", "", False, 1, 15)

                    Call ActualizarDisplay(0, "ALARMA C Movil " & Alarmas("nro_Movil") & ".", "")

                    Application.DoEvents()

                Loop

                Alarmas.Close()

                Dim SqlCmd2 As New SqlCommand("Select * from Alarmas where idAlarma in (Select Max (idAlarma) from Alarmas where " & _
                                              " Fecha_Hora_Fin > dateadd(s, -600, getdate()) and Estado_alarma = 'A' group by nro_movil)", DBConn)
                SqlCmd2.CommandType = CommandType.Text

                Dim Alarmas1 = SqlCmd1.ExecuteReader()

                Do While Alarmas1.Read

                    'procesa cada movil y decide qué acción tomar

                    Dim DatosMovil = Movil.ObtenerDatos(Alarmas("nro_Movil"))

                    'Consulta Posicion
                    Comando.EnviarComando(DatosMovil.NroCanal, Alarmas("nro_Movil"), "50", "", False, 1, 15)

                    Call ActualizarDisplay(0, "ALARMA NC Movil " & Alarmas("nro_Movil") & ".", "")

                    Application.DoEvents()

                Loop

                Alarmas1.Close()

                'Actualizar WatchDog

                Dim SqlCmd3 As New SqlCommand("update Config_Sis set WatchDogSD = getdate() where cod_registro = 1 ", DBConn)

                SqlCmd3.CommandType = CommandType.Text

                SqlCmd3.ExecuteNonQuery()


                DBConn.Close()

            End Using

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try


        Me.tmrAlarmasCanceladas.Enabled = True
    End Sub
End Class
