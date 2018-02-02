Imports System.Data.SqlClient

Public Module Comando

    Public ConnString As String
    Public Status As Integer

    Public Sub BorrarComandosTodos()

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("Delete from COMANDOS_CM ", DBConn)

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


    Public Sub BorrarComandosCanal(ByVal Nrocanal As Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("Delete from COMANDOS_CM WHERE nro_canal = @0", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@0", Nrocanal)

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

    Public Sub EnviarComando(ByVal NroCanal As Integer, ByVal NroMovil As Integer, ByVal CodComando As String, ByVal DatosComando As String, ByVal EsperarRespuesta As Boolean, ByVal MaxCantReintentos As Integer, ByVal EsperaReintentos As Integer)

        Try


            '---  Convierte el comando a una secuencia de caracteres no binarios

            Dim MensajeConv As String = ""
            Dim Parte1 As String
            Dim Parte2 As String
            Dim Parte3 As String

            For i = 1 To Len(DatosComando)
                MensajeConv = MensajeConv & DecToHex(Asc(Mid(DatosComando, i, 1)))
            Next

            Select Case Len(MensajeConv)

                Case 0
                    Parte1 = ""
                    Parte2 = ""
                    Parte3 = ""

                Case 1 To 254
                    Parte1 = Mid(MensajeConv, 1, Len(MensajeConv))
                    Parte2 = ""
                    Parte3 = ""

                Case 255 To 508
                    Parte1 = Mid(MensajeConv, 1, 254)
                    Parte2 = Mid(MensajeConv, 255, Len(MensajeConv))
                    Parte3 = ""

                Case 509 To 765
                    Parte1 = Mid(MensajeConv, 1, 254)
                    Parte2 = Mid(MensajeConv, 255, 254)
                    Parte3 = Mid(MensajeConv, 509, Len(MensajeConv))

                Case Else
                    Parte1 = Mid(MensajeConv, 1, 254)
                    Parte2 = Mid(MensajeConv, 255, 254)
                    Parte3 = Mid(MensajeConv, 509, 254)

            End Select

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("insert into comandos_CM (nro_canal, nro_secuencia, fecha_hora_comando, codigo_comando, datos_comando_1, datos_comando_2, datos_comando_3, nro_movil, estado_comando, max_cant_reintentos, espera_reintentos, hora_prox_reintento, cant_reintentos) values (@0, @1, @2, @3, @4, @5, @6, @7, @8, @9, @10, @11, @12)", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@0", NroCanal)
            SqlCmd.Parameters.AddWithNullableValue("@1", 0)
            SqlCmd.Parameters.AddWithNullableValue("@2", Now)
            SqlCmd.Parameters.AddWithNullableValue("@3", CodComando)
            SqlCmd.Parameters.AddWithNullableValue("@4", Parte1)
            SqlCmd.Parameters.AddWithNullableValue("@5", Parte2)
            SqlCmd.Parameters.AddWithNullableValue("@6", Parte3)
            SqlCmd.Parameters.AddWithNullableValue("@7", NroMovil)
            SqlCmd.Parameters.AddWithNullableValue("@8", "P")
            SqlCmd.Parameters.AddWithNullableValue("@9", MaxCantReintentos)
            SqlCmd.Parameters.AddWithNullableValue("@10", EsperaReintentos)
            SqlCmd.Parameters.AddWithNullableValue("@11", Now)
            SqlCmd.Parameters.AddWithNullableValue("@12", 0)

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

    Public Sub CancelarComandosNoRespondidos(ByVal NroMovil As Integer, ByVal ListaComandos As String)

        '-- Lista comandos es una lista de strings separada por comas. Ejemplo:    ListaComandos = "01,AB,4A,03,10,....,FF"

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            '-- la siguiente sentencia usa CHARINDEX como un truco para detectar todos los comandos donde "codigo_comando" contenga alguno de los valores listados en @ListaComandos

            Dim SqlCmd As New SqlCommand("delete from comandos_cm where nro_movil = @NroMovil and CHARINDEX( ',' + codigo_comando + ',', ',' + @ListaComandos + ',' ) > 0 ", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@NroMovil", NroMovil)
            SqlCmd.Parameters.AddWithNullableValue("@ListaComandos", ListaComandos)

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

    Public Sub MarcarComoRespondido(ByVal NroMovil As Integer, ByVal ListaComandos As String)


        '-- Lista comandos es una lista de strings separada por comas. Ejemplo:    ListaComandos = "01,AB,4A,03,10,....,FF"

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            '-- la siguiente sentencia usa CHARINDEX como un truco para detectar todos los comandos donde "codigo_comando" contenga alguno de los valores listados en @ListaComandos

            Dim SqlCmd As New SqlCommand("update comandos_cm set estadoComando = 'R' where nro_movil = @NroMovil and estado_Comando = 'E' and CHARINDEX( ',' + codigo_comando + ',', ',' + @ListaComandos + ',' ) > 0 ", DBConn)

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

End Module
