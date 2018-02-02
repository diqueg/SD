Imports System.Runtime.CompilerServices

Module SqlParametersCollectionExtensions

    <Extension()>
    Public Sub AddWithNullableValue(ByVal SqlParamCol As System.Data.SqlClient.SqlParameterCollection, ParameterName As String, ByVal Value As Object)

        Try

            If Not IsNothing(Value) Then
                SqlParamCol.AddWithValue(ParameterName, Value)
            Else
                SqlParamCol.AddWithValue(ParameterName, DBNull.Value)
            End If

        Catch ex As Exception

        End Try

    End Sub


End Module
