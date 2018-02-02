Imports System.Data.SqlClient

Public Module AlarmaMovil

    Public ConnString As String
    Public Status As Integer


    Public Sub Crear(ByVal NroMovil As Integer, ByVal NroViaje As Long, ByVal IdCausa As String, ByVal Causa As String)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim NuevoViaje = New DatosViaje

            Dim SqlInsert As New SqlCommand("insert into Alarma_Movil (IdMovil, IdCausa, Causa, IdViaje) values (@NroMovil, @IdCausa, @Causa, @NroViaje)", DBConn)

            SqlInsert.Parameters.AddWithNullableValue("@NroMovil", NroMovil)
            SqlInsert.Parameters.AddWithNullableValue("@IdCausa", IdCausa)
            SqlInsert.Parameters.AddWithNullableValue("@Causa", Causa)
            SqlInsert.Parameters.AddWithNullableValue("@NroViaje", NroViaje)

            SqlInsert.CommandType = CommandType.Text

            SqlInsert.ExecuteNonQuery()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

    End Sub



End Module
