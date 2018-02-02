
Imports System.Data.SqlClient

Public Module Base

    Public ConnString As String
    Public Status As Integer
    Public Bases As List(Of DatosBase)

    Public Class CentroBase
        Public Latitud As Single
        Public Longitud As Single
    End Class

    Public Class DatosBase
        Public LatitudCentro As Single
        Public LongitudCentro As Single
        Public IdBase As Integer
        Public DistanciaABase As Long
    End Class

    Public Function VerificarEntradaBase(ByVal Latitud As Single, ByVal Longitud As Single) As DatosBase

        Dim ValorBase As Integer
        Dim Distancia As Single
        Dim DatosBase As New DatosBase


        Try

            If Latitud <> 0 And Longitud <> 0 Then

                Dim DBConn = New SqlConnection(ConnString)
                DBConn.Open()

                Dim SqlCmd As New SqlCommand("Base_Obtener2", DBConn)
                SqlCmd.CommandType = CommandType.StoredProcedure

                SqlCmd.Parameters.AddWithNullableValue("@Latitud", Latitud)
                SqlCmd.Parameters.AddWithNullableValue("@Longitud", Longitud)

                SqlCmd.Parameters.Add("ReturnValue", SqlDbType.Int)
                SqlCmd.Parameters("ReturnValue").Direction = ParameterDirection.ReturnValue

                SqlCmd.Parameters.Add("@Base", SqlDbType.Int)
                SqlCmd.Parameters("@Base").Direction = ParameterDirection.Output

                SqlCmd.Parameters.Add("@Distancia", SqlDbType.Real)
                SqlCmd.Parameters("@Distancia").Direction = ParameterDirection.Output

                SqlCmd.ExecuteScalar()

                ValorBase = SqlCmd.Parameters("@Base").Value
                Distancia = SqlCmd.Parameters("@Distancia").Value

                DatosBase.IdBase = ValorBase
                DatosBase.DistanciaABase = Distancia

                If IsNothing(DatosBase.IdBase) Then
                    DatosBase.IdBase = 0
                    DatosBase.DistanciaABase = 0
                End If

                If DBConn.State = ConnectionState.Open Then
                    DBConn.Close()
                End If

                DBConn.Dispose()

            Else
                DatosBase.IdBase = 0
                DatosBase.DistanciaABase = 0
            End If

            Status = 0

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            DatosBase.IdBase = 0
            DatosBase.DistanciaABase = 0
            Status = 1
        End Try

        Return DatosBase

    End Function

    Public Function ObtenerBase(ByVal Latitud As Single, ByVal Longitud As Single) As Integer

        Dim Valor As Integer

        Try

            If Latitud <> 0 And Longitud <> 0 Then

                Dim DBConn = New SqlConnection(ConnString)
                DBConn.Open()

                Dim SqlCmd As New SqlCommand("Base_Obtener1", DBConn)
                SqlCmd.CommandType = CommandType.StoredProcedure

                SqlCmd.Parameters.AddWithNullableValue("@Latitud", Latitud)
                SqlCmd.Parameters.AddWithNullableValue("@Longitud", Longitud)

                SqlCmd.Parameters.Add("ReturnValue", SqlDbType.Int)
                SqlCmd.Parameters("ReturnValue").Direction = ParameterDirection.ReturnValue

                SqlCmd.Parameters.Add("@Zona", SqlDbType.Int)
                SqlCmd.Parameters("@Zona").Direction = ParameterDirection.Output

                SqlCmd.ExecuteScalar()

                Valor = SqlCmd.Parameters("@Zona").Value

                If IsNothing(Valor) Then
                    Valor = 0
                End If

                If DBConn.State = ConnectionState.Open Then
                    DBConn.Close()
                End If

                DBConn.Dispose()

            Else
                Valor = 0
            End If

            Status = 0

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            Valor = 0
            Status = 1
        End Try

        Return Valor

    End Function



End Module
