Imports System.Data.SqlClient

Public Module CuentaCorriente

    Public ConnString As String
    Public Status As Integer

    Public Function ObtenerNombre(ByVal NroCtaCte As Integer) As String

        Dim Nombre As String = Nothing

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("Select nombre from CTACTE where nro_ctacte = @NroCtaCte", DBConn)
            SqlCmd.CommandType = CommandType.Text

            SqlCmd.Parameters.AddWithNullableValue("@NroCtaCte", NroCtaCte)

            Dim Regs = SqlCmd.ExecuteReader()

            Do While Regs.Read
                Nombre = Regs("nombre")
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
