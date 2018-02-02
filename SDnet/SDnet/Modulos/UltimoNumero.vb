Imports System.Data.SqlClient

Public Module UltimoNumero

    Public ConnString As String
    Public Status As Integer

    Public Function ObtenerProximo(ByVal Codigo As String) As Integer

        Dim ProximoNumero As Integer

        Dim DBConn = New SqlConnection(ConnString)
        DBConn.Open()

        Dim sqlTran As SqlTransaction = DBConn.BeginTransaction()

        Try

            Dim SqlCmdRead As New SqlCommand("select * from UltimosNumeros where CodigoItem = @CodigoItem", DBConn)
            SqlCmdRead.CommandType = CommandType.Text
            SqlCmdRead.Parameters.AddWithNullableValue("@CodigoItem", Codigo)
            SqlCmdRead.Transaction = sqlTran

            Dim Regs = SqlCmdRead.ExecuteReader()

            ProximoNumero = 0

            While Regs.Read

                ProximoNumero = Regs("ultimonumero") + 1

                Dim SqlCmdUpdate As New SqlCommand("update UltimosNumeros set ultimonumero = @ProxNum where CodigoItem = @CodigoItem", DBConn)
                SqlCmdUpdate.CommandType = CommandType.Text
                SqlCmdUpdate.Parameters.AddWithNullableValue("@ProxNum", ProximoNumero)
                SqlCmdUpdate.Parameters.AddWithNullableValue("@CodigoItem", Codigo)
                SqlCmdUpdate.Transaction = sqlTran
                SqlCmdUpdate.ExecuteNonQuery()

                Exit While

            End While

            Regs.Close()

            sqlTran.Commit()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0
        Catch ex As Exception
            sqlTran.Rollback()
            ProximoNumero = 0
            Status = 1
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

        Return ProximoNumero

    End Function

    Public Function Establecer(ByVal Codigo As String, ByVal Valor As Integer) As Boolean

        Dim Result As Boolean

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("update UltimosNumeros set ultimonumero = @UltNumero where CodigoItem = @CodigoItem", DBConn)
            SqlCmd.CommandType = CommandType.Text
            SqlCmd.Parameters.AddWithNullableValue("@UltNumero", Valor)
            SqlCmd.Parameters.AddWithNullableValue("@CodigoItem", Codigo)
            SqlCmd.ExecuteNonQuery()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Result = True

            Status = 0

        Catch ex As Exception
            Result = False
            Status = 1
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
        End Try

        Return Result

    End Function

End Module
