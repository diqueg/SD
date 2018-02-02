Imports System.Data.SqlClient

Public Module Localidad

    Public ConnString As String
    Public Status As Integer

    Public Function ObtenerCodigo(ByVal IdLocalidad As Integer) As String

        Dim Nombre As String = Nothing

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("Select cod_localidad from LOCALIDADES where IdLocalidad = @IdLocalidad", DBConn)
            SqlCmd.CommandType = CommandType.Text
            SqlCmd.Parameters.AddWithNullableValue("@IdLocalidad", IdLocalidad)

            Dim Regs = SqlCmd.ExecuteReader()

            Do While Regs.Read
                Nombre = Regs("cod_localidad")
                Exit Do
            Loop

            Regs.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            Nombre = Nothing
            Status = 1
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

        Return Nombre

    End Function

    Public Function ObtenerNombreAbreviado(ByVal IdLocalidad As Integer) As String

        Dim Nombre As String = Nothing

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("Select nombre_abreviado from LOCALIDADES where IdLocalidad = @IdLocalidad", DBConn)
            SqlCmd.CommandType = CommandType.Text
            SqlCmd.Parameters.AddWithNullableValue("@IdLocalidad", IdLocalidad)

            Dim Regs = SqlCmd.ExecuteReader()

            Do While Regs.Read
                Nombre = Regs("nombre_abreviado")
                Exit Do
            Loop

            Regs.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            Nombre = Nothing
            Status = 1
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

        Return Nombre

    End Function
End Module