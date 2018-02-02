Imports System.Data.SqlClient

Public Module Despacho

    Public ConnString As String
    Public Status As Integer


    Public Class MejorMovil
        Public IdMovilOptimo As Integer
        Public DistanciaMovilOptimo As Long
        Public TipoServicio As Integer
    End Class

    Public Function BuscarMovilLibreParada(ByVal latitud As Single, ByVal Longitud As Single, ByVal idParada As Single, DistanciaMaxima As Long, ByVal Categoria As String, ByVal TipoServicio As Integer, ByVal MovilExcluido As Integer) As MejorMovil

        'Busca primero en la Base que le indica el viaje
        'Si no encuentra Busca en las bases que estan en un radio max de busqueda

        Dim MovilElegido As New MejorMovil

        Try
            Dim strFiltroCategoria As String
            Dim strFiltroTipoServicio As String

            If Categoria = "5" Or Categoria = "*" Or Categoria = "" Then
                strFiltroCategoria = ""
            Else
                strFiltroCategoria = " and categ_movil like '" & Categoria & "'"
            End If

            If TipoServicio = 3 Then
                strFiltroTipoServicio = ""
            Else
                strFiltroTipoServicio = " and Tipo_Servicio = " & CStr(TipoServicio)
            End If


            Zona.ActualizarDatosHorario()

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            'busca el movil que entro primero a la base

            Dim strSQL = " Select Nro_movil, Tipo_Servicio from Moviles where cod_parada = @IdParada and Estado_Administrativo ='H' " & _
                                " and Estado_Operativo = 'L'  " & strFiltroCategoria & strFiltroCategoria & " order by fecha_hora_entrada_parada"

            Dim SqlCmd1 As New SqlCommand(strSQL, DBConn)

            SqlCmd1.Parameters.AddWithNullableValue("@IdParada", idParada)
            SqlCmd1.CommandType = CommandType.Text

            Dim regMovilElegido = SqlCmd1.ExecuteReader()

            If regMovilElegido.Read() Then
                MovilElegido.IdMovilOptimo = regMovilElegido("nro_movil")
                MovilElegido.TipoServicio = regMovilElegido("Tipo_Servicio")
                MovilElegido.DistanciaMovilOptimo = 4
            Else
                MovilElegido.IdMovilOptimo = 0
                MovilElegido.TipoServicio = 0
                MovilElegido.DistanciaMovilOptimo = 0
            End If

            Status = 0

            regMovilElegido.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()
            Status = 0

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            Status = 1
            MovilElegido.IdMovilOptimo = 0
        End Try


        Return MovilElegido

    End Function

    Public Function BuscarMovilLibreDistancia(ByVal LatitudViaje As Single, ByVal LongitudViaje As Single, ByVal DistanciaBusqueda As Long, ByVal MovilExcluido As Integer, ByVal Categoria As String, ByVal TipoServicio As Integer) As MejorMovil

        Dim MovilElegido As New MejorMovil

        Try

            Zona.ActualizarDatosHorario()

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            'busca el movil que entro primero a la base

            Dim Ancho = DistanciaBusqueda / 1853 / 60
            Dim LatitudPlus = LatitudViaje + Ancho
            Dim LatitudMinus = LatitudViaje - Ancho
            Dim LongitudPlus = LongitudViaje + Ancho
            Dim LongitudMinus = LongitudViaje - Ancho

            Dim strSQL = " Select top 1 Nro_movil, SQRT (SQUARE(Latitud -@Latitud) + SQUARE(Longitud - @Longitud)) * 60 * 1853 as Distancia " & _
             "from moviles where Latitud < @LatitudPlus and Latitud > @LatitudMinus and Longitud < @LongitudPlus and Longitud > @LongitudMinus and (estado_operativo = 'L' or estado_operativo = 'R') " & _
            " and estado_administrativo = 'H' and fecha_hora_ultimo_contacto > DATEADD(s, -150, GETDATE()) "

            'If MovilExcluido <> 0 Then strSQL = strSQL & " and IdMovil <>  " & MovilExcluido

            If Categoria <> "5" And Categoria <> "*" And Categoria <> "" Then
                strSQL = strSQL & " and categ_Movil like '" & Categoria & "'"
            End If

            If TipoServicio <> 3 Then
                strSQL = strSQL & " and Tipo_Servicio = " & CStr(TipoServicio)
            End If

            strSQL = strSQL & " and nro_canal_comunicaciones in (Select NroCanal FROM Mesas where NroMesa = 5) "

            'strSQL = strSQL & " and nro_flota = 3 "
            strSQL = strSQL & "  order by distancia"
            Dim SqlCmd1 As New SqlCommand(strSQL, DBConn)

            SqlCmd1.Parameters.AddWithNullableValue("@Latitud", LatitudViaje)
            SqlCmd1.Parameters.AddWithNullableValue("@Longitud", LongitudViaje)
            SqlCmd1.Parameters.AddWithNullableValue("@LatitudPlus", LatitudPlus)
            SqlCmd1.Parameters.AddWithNullableValue("@LatitudMinus", LatitudMinus)
            SqlCmd1.Parameters.AddWithNullableValue("@LongitudPlus", LongitudPlus)
            SqlCmd1.Parameters.AddWithNullableValue("@LongitudMinus", LongitudMinus)

            SqlCmd1.CommandType = CommandType.Text

            Dim regMovilElegido = SqlCmd1.ExecuteReader()
            'regMovilElegido.Read()

            If regMovilElegido.Read() Then
                MovilElegido.IdMovilOptimo = regMovilElegido("nro_movil")
                MovilElegido.TipoServicio = 3
                MovilElegido.DistanciaMovilOptimo = regMovilElegido("Distancia")
            Else
                MovilElegido.IdMovilOptimo = 0
                MovilElegido.TipoServicio = 0
                MovilElegido.DistanciaMovilOptimo = 0
            End If

            Status = 0

            regMovilElegido.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()
            Status = 0

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            Status = 1
            MovilElegido.IdMovilOptimo = 0
        End Try


        Return MovilElegido

    End Function

    Public Function BuscarMovilOcupadoAZona(ByVal LatitudViaje As Single, ByVal LongitudViaje As Single, ByVal MovilExcluido As Integer, ByVal Categoria As String, ByVal TipoServicio As Integer, ByVal TiempoAnticipacion As Integer, ByVal IdZona As Integer) As MejorMovil

        Dim MovilElegido As New MejorMovil

        Try

            Zona.ActualizarDatosHorario()

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            'busca el movil que entro primero a la base


            '            Dim strSQL = " Select Nro_movil from Moviles where cod_zona_futura in " & _
            '                       " (Select cod_zona from Zonas where SQRT (SQUARE(LatitudCentro -@Latitud) + SQUARE(LongitudCentro - @Longitud)) * 60 * 1853 < @DistanciaBusqueda " & _
            '           "and Estado_Operativo = 'O' and Estado_Administrativo = 'H' and Estado_Futuro = 'L' and fecha_hora_prevista_liberacion < DATEADD(minute, 2, GETDATE()) "

            Dim strSQL = " Select top 1 Nro_movil  from Moviles where cod_zona_futura = @IdZona " & _
                                 " and Estado_Operativo = 'O' and Estado_Administrativo = 'H' and Estado_Futuro = 'L' and fecha_hora_prevista_liberacion < DATEADD(minute, " & CStr(TiempoAnticipacion) & ", GETDATE()) "


            'If MovilExcluido <> 0 Then strSQL = strSQL & " and IdMovil <>  " & MovilExcluido

            If Categoria <> "5" And Categoria <> "*" And Categoria <> "" Then
                strSQL = strSQL & " and categ_Movil like '" & Categoria & "'"
            End If

            If TipoServicio <> 3 Then
                strSQL = strSQL & " and Tipo_Servicio = " & CStr(TipoServicio)
            End If

            strSQL = strSQL & " and nro_canal_comunicaciones in (Select NroCanal FROM Mesas where NroMesa = 5) "

            'strSQL = strSQL & " and nro_flota = 3 "
            'strSQL = strSQL & "  order by distancia"

            Dim SqlCmd1 As New SqlCommand(strSQL, DBConn)

            SqlCmd1.Parameters.AddWithNullableValue("@Latitud", LatitudViaje)
            SqlCmd1.Parameters.AddWithNullableValue("@Longitud", LongitudViaje)
            SqlCmd1.Parameters.AddWithNullableValue("@IdZona", IdZona)

            SqlCmd1.CommandType = CommandType.Text

            Dim regMovilElegido = SqlCmd1.ExecuteReader()
            'regMovilElegido.Read()

            If regMovilElegido.Read() Then
                MovilElegido.IdMovilOptimo = regMovilElegido("nro_movil")
                MovilElegido.TipoServicio = 3
                MovilElegido.DistanciaMovilOptimo = 2 'Indica Futuro
            Else
                MovilElegido.IdMovilOptimo = 0
                MovilElegido.TipoServicio = 0
                MovilElegido.DistanciaMovilOptimo = 0
            End If

            Status = 0

            regMovilElegido.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()
            Status = 0

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            Status = 1
            MovilElegido.IdMovilOptimo = 0
        End Try


        Return MovilElegido

    End Function

    Public Function BuscarMovilZonaAdyacentes(ByVal IdZona As Integer, ByVal MovilExcluido As Integer, ByVal Categoria As String, ByVal TipoServicio As Integer) As MejorMovil

        Dim MovilElegido As New MejorMovil

        Try

            Zona.ActualizarDatosHorario()

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim strSQL = " Select cod_zona_aled from Zonas_aled where cod_zona_pivot = @IdZona order by prioridad_zona"

            Dim SqlCmd1 As New SqlCommand(strSQL, DBConn)

            SqlCmd1.Parameters.AddWithNullableValue("@IdZona", IdZona)

            SqlCmd1.CommandType = CommandType.Text

            Dim ZonasAled = SqlCmd1.ExecuteReader()
            'regMovilElegido.Read()

            Do While ZonasAled.Read

                strSQL = " Select Top 1 nro_movil from Moviles where cod_zona_actual = @ZonaALed" & _
                                " and (estado_operativo = 'L' or estado_operativo = 'R') " & _
                                " and estado_administrativo = 'H' and fecha_hora_ultimo_contacto > DATEADD(s, -250, GETDATE()) "
                'If MovilExcluido <> 0 Then strSQL = strSQL & " and IdMovil <>  " & MovilExcluido

                If Categoria <> "5" And Categoria <> "*" And Categoria <> "" Then
                    strSQL = strSQL & " and categ_Movil like '" & Categoria & "'"
                End If

                If TipoServicio <> 3 Then
                    strSQL = strSQL & " and Tipo_Servicio = " & CStr(TipoServicio)
                End If

                strSQL = strSQL & " and nro_canal_comunicaciones in (Select NroCanal FROM Mesas where NroMesa = 5) "

                'strSQL = strSQL & " and nro_flota = 3 "

                Dim SqlCmd2 As New SqlCommand(strSQL, DBConn)

                SqlCmd2.Parameters.AddWithNullableValue("@ZonaALed", ZonasAled("cod_zona_aled"))

                SqlCmd2.CommandType = CommandType.Text

                Dim MovilesAled = SqlCmd2.ExecuteReader()

                If MovilesAled.Read() Then
                    MovilElegido.IdMovilOptimo = MovilesAled("nro_movil")
                    MovilElegido.TipoServicio = 3
                    MovilElegido.DistanciaMovilOptimo = 1 'Indica Adyacente
                    MovilesAled.Close()
                    Exit Do
                Else
                    MovilElegido.IdMovilOptimo = 0
                    MovilElegido.TipoServicio = 0
                    MovilElegido.DistanciaMovilOptimo = 0
                End If
                MovilesAled.Close()
            Loop

            ZonasAled.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()
            Status = 0

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            Status = 1
            MovilElegido.IdMovilOptimo = 0
        End Try


        Return MovilElegido

    End Function

    Public Function BuscarMovilPostulante(ByVal IdViaje As Long, ByVal LatitudViaje As Single, ByVal LongitudViaje As Single, ByVal MovilExcluido As Integer, ByVal Categoria As String, ByVal TipoServicio As Integer) As MejorMovil

        Dim MovilElegido As New MejorMovil

        Try

            'Zona.ActualizarDatosHorario()

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            'busca el movil que entro primero a la base

            
            Dim StrsqlMovil = "Select Nro_Movil from Moviles where (estado_operativo = 'L' or estado_operativo = 'R') " & _
            " and estado_administrativo = 'H' and fecha_hora_ultimo_contacto > DATEADD(s, -150, GETDATE()) "

            'If MovilExcluido <> 0 Then strSQL = strSQL & " and IdMovil <>  " & MovilExcluido

            If Categoria <> "5" And Categoria <> "*" And Categoria <> "" Then
                StrsqlMovil = StrsqlMovil & " and categ_Movil like '" & Categoria & "'"
            End If

            If TipoServicio <> 3 Then
                StrsqlMovil = StrsqlMovil & " and Tipo_Servicio = " & CStr(TipoServicio)
            End If

            StrsqlMovil = StrsqlMovil & " and nro_canal_comunicaciones in (Select NroCanal FROM Mesas where NroMesa = 5) "

            Dim strSQL = " Select  IdMovil, SQRT (SQUARE(Latitud -@Latitud) + SQUARE(Longitud - @Longitud)) * 60 * 1853 as Distancia " & _
             "from Viajes_Postulantes where idViaje = @IdViaje and indTomado = 'False'  and IdMovil in (" & StrsqlMovil & ")"

            'strSQL = strSQL & " and nro_flota = 3 "
            strSQL = strSQL & "  order by distancia"
            Dim SqlCmd1 As New SqlCommand(strSQL, DBConn)


            SqlCmd1.Parameters.AddWithNullableValue("@IdViaje", IdViaje)
            SqlCmd1.Parameters.AddWithNullableValue("@Latitud", LatitudViaje)
            SqlCmd1.Parameters.AddWithNullableValue("@Longitud", LongitudViaje)
            
            SqlCmd1.CommandType = CommandType.Text

            Dim regMovilElegido = SqlCmd1.ExecuteReader()
            'regMovilElegido.Read()

            If regMovilElegido.Read() Then
                MovilElegido.IdMovilOptimo = regMovilElegido("IdMovil")
                MovilElegido.TipoServicio = 3
                MovilElegido.DistanciaMovilOptimo = regMovilElegido("Distancia")
                MarcaPostulacionUtilizada(IdViaje, MovilElegido.IdMovilOptimo, MovilElegido.DistanciaMovilOptimo)
            Else
                MovilElegido.IdMovilOptimo = 0
                MovilElegido.TipoServicio = 0
                MovilElegido.DistanciaMovilOptimo = 0
            End If

            Status = 0

            regMovilElegido.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()
            Status = 0

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            Status = 1
            MovilElegido.IdMovilOptimo = 0
        End Try


        Return MovilElegido

    End Function

    Public Sub MarcaPostulacionUtilizada(ByVal IdViaje As Long, ByVal IdMovil As Integer, ByVal Distancia As Single)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()


            Dim SqlCmdUpdate As New SqlCommand("Update Viajes_Postulantes set indTomado = 'True' , Dist = @Distancia, FechaHoraTomado= getdate() Where IdMovil = @IdMovil and IdViaje =@IdViaje", DBConn)

            SqlCmdUpdate.Parameters.AddWithNullableValue("@IdMovil", IdMovil)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@IdViaje", IdViaje)
            SqlCmdUpdate.Parameters.AddWithNullableValue("@Distancia", Distancia)

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
