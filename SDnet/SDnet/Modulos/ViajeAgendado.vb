Imports System.Data.SqlClient


Public Module ViajeAgendado

    Public ConnString As String
    Public Status As Integer

    Public Class DatosViajeAgendado

        Public NroViajeAgendado As Integer?
        Public NroCliente As Integer?
        Public Lunes As Boolean?
        Public Martes As Boolean?
        Public Miercoles As Boolean?
        Public Jueves As Boolean?
        Public Viernes As Boolean?
        Public Sabado As Boolean?
        Public Domingo As Boolean?
        Public OpcionAgenda As Integer?
        Public HoraViaje As Date
        Public FechaInicioVigencia As Date?
        Public FechaFinVigencia As Date?
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
        Public CantMoviles As Int16?
        Public ImportePactado As Decimal?
        Public CategoriaMovil As String
        Public TipoViaje As String
        Public NroMovilReq As Int16?
        Public TipoServicio As Int16?
        Public Nombre As String
        Public Observaciones As String
        Public CodZona As Int16?
        Public NrousuarioTelefonista As Integer
        Public NroCtaCte As Integer?
        Public NroSubCta As Integer?
        Public FechaHoraIngreso As Date?

    End Class


    Public Function EsHabil(ByVal Fecha As Date) As Boolean

        Return False

    End Function

    Public Sub MarcarComoGenerado(ByVal IdViaje As Long, ByVal IdViajeAgendado As Long, ByVal FechaHoraAgendado As Date)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("update Viajes_Agendados_items set nro_viaje_generado = @IdViaje where nro_viaje_agendado = @IdViajeAgendado and Fecha_hora_Viaje = @FechaHoraAgendado ", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@IdViaje", IdViaje)
            SqlCmd.Parameters.AddWithNullableValue("@IdViajeAgendado", IdViajeAgendado)
            SqlCmd.Parameters.AddWithNullableValue("@FechaHoraAgendado", FechaHoraAgendado)

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

    Public Sub MarcarIVRComoGenerado(ByVal IdInterfaz As Long, ByVal IdViajeGenerado As Long)

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("update IVRInterfaz set indTomado = 'True' where idIVR = @IdInterfaz", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@IdInterfaz", IdInterfaz)
            'SqlCmd.Parameters.AddWithNullableValue("@IdViajeAgendado", IdViajeAgendado)

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
