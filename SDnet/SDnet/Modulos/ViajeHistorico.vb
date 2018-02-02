Imports System.Data.SqlClient

Public Module ViajeHistorico

    Public ConnString As String
    Public Status As Integer

    Public Function ObtenerDuracionViajesPorZona(ByVal Desde As Date, ByVal Hasta As Date) As Dictionary(Of Integer, Integer)

        Dim Result = New Dictionary(Of Integer, Integer)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("select cod_zona, (60/count(nro_viaje)) as cantidad from viajes_historicos where fecha_hora_inicio_viaje >= @0  and fecha_hora_inicio_viaje <= @1  group by cod_zona order by cod_zona ", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@0", Desde)
            SqlCmd.Parameters.AddWithNullableValue("@1", Hasta)

            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            While Regs.Read

                Dim CantViajes As Integer

                If Regs("cantidad") > 0 And Regs("cantidad") <= 15 Then
                    CantViajes = Regs("cantidad")
                Else
                    CantViajes = 15
                End If

                Result.Add(Regs("cod_zona"), CantViajes)

            End While

            Regs.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            Result = Nothing
            Status = 1
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

        Return Result

    End Function

End Module
