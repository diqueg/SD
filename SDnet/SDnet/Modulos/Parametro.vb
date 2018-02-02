Imports System.Data.SqlClient

Public Module Parametro

    Public ConnString As String
    Public Status As Integer


    Friend Function Leer(ByVal NombreParametro As String) As String

        Dim Valor As String = Nothing

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("Select TOP 1 * from CONFIG_SIS", DBConn)
            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader()

            Do While Regs.Read
                Valor = Regs(NombreParametro)
                Exit Do
            Loop

            Regs.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            Valor = Nothing
            Status = 1
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

        Return Valor

    End Function

End Module

