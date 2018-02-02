Imports System.Data.SqlClient

Public Class DatosCliente
    
    Public NroCliente As Long
    Public Nombre As String
    Public CodCalle As Long
    Public NombreCalle As String
    Public Numero As Long
    Public Ubicacion As String
    Public CodLocalidad As Integer
    Public Localidad As String
    Public Latitud As Single
    Public Longitud As Single
    Public CodCalle1 As Long
    Public NombreCalle1 As String
    Public CodCalle2 As Long
    Public NombreCalle2 As String
    Public Telefono As Double
    Public Calificacion As Integer
    Public Observaciones As String
    Public Caracteristicas As String
    Public NroCtaCte As Integer
    Public NroSubCta As Integer
    Public CZ As String
    Public ReferenciaPC As String
    Public FechaAlta As Date
    Public ViajesAnoAnteultimo As Integer
    Public ViajesAnoUltimo As Integer
    Public ViajesAnoActual As Integer
    Public AnoActual As Integer

    Public Domicilio As String

    Public CodZona As Integer
    Public CodBase As Integer
    Public bolLugarEspecial As Boolean
    Public bolEsquina As Boolean
    Public bolDomicilioLibre As Boolean
    Public bolLocalidad As Boolean

End Class

Module Cliente

    Public ConnString As String
    Public Status As Integer

    Public Function ObtenerDatosxTelefono(ByVal Telefono As String) As DatosCliente

        Dim DatosCliente As New DatosCliente

        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("Select * from Clientes where nro_telefono = @Telefono", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@Telefono", Telefono)

            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            If Regs.HasRows Then

                Regs.Read()

                DatosCliente.NroCliente = IsNull(Regs("nro_cliente"))
                DatosCliente.Nombre = IsNull(Regs("Nombre"))
                DatosCliente.CodCalle = IsNull(Regs("Cod_Calle"))
                DatosCliente.NombreCalle = IsNull(Regs("Nombre_Calle"))
                DatosCliente.Numero = IsNull(Regs("Numero"))
                DatosCliente.Ubicacion = IsNull(Regs("Ubicacion"))
                DatosCliente.CodLocalidad = IsNull(Regs("Cod_Localidad"))
                DatosCliente.Localidad = IsNull(Regs("Localidad"))
                DatosCliente.Latitud = IsNull(Regs("Latitud"))
                DatosCliente.Longitud = IsNull(Regs("Longitud"))
                DatosCliente.CodCalle1 = IsNull(Regs("Cod_Calle1"))
                DatosCliente.NombreCalle1 = IsNull(Regs("Nombre_Calle1"))
                DatosCliente.CodCalle2 = IsNull(Regs("Cod_Calle2"))
                DatosCliente.NombreCalle2 = IsNull(Regs("Nombre_Calle2"))
                DatosCliente.Telefono = IsNull(Regs("Nro_Telefono"))
                DatosCliente.Calificacion = IsNull(Regs("Calificacion"))
                DatosCliente.Observaciones = IsNull(Regs("Observaciones"))
                DatosCliente.NroCtaCte = IsNull(Regs("Nro_CtaCte"))
                DatosCliente.NroSubCta = IsNull(Regs("Nro_SubCta"))
                DatosCliente.CZ = IsNull(Regs("CZ"))
                DatosCliente.ReferenciaPC = IsNull(Regs("Referencia_PC"))
                DatosCliente.FechaAlta = IsNull(Regs("Fecha_Alta"))
                DatosCliente.ViajesAnoAnteultimo = IsNull(Regs("viajes_ano_anteultimo"))
                DatosCliente.ViajesAnoUltimo = IsNull(Regs("viajes_ano_ultimo"))
                DatosCliente.ViajesAnoActual = IsNull(Regs("viajes_ano_actual"))

                Select Case DatosCliente.CZ
                    Case "L"
                        DatosCliente.bolLugarEspecial = True
                    Case "E"
                        DatosCliente.bolEsquina = True
                    Case "DL"
                        DatosCliente.bolDomicilioLibre = True
                    Case "C"
                        DatosCliente.bolLocalidad = True
                End Select

                If DatosCliente.bolDomicilioLibre = True Then
                    DatosCliente.Domicilio = DatosCliente.Ubicacion
                Else
                    If DatosCliente.bolLocalidad = True Then
                        DatosCliente.Domicilio = " EN " & DatosCliente.Localidad & " - " & DatosCliente.Ubicacion
                    Else
                        If Not IsNothing(DatosCliente.NombreCalle) And DatosCliente.NombreCalle <> " " Then
                            DatosCliente.Domicilio = DatosCliente.NombreCalle
                        End If
                        If (Not IsNothing(DatosCliente.Numero) And DatosCliente.Numero <> 0) And DatosCliente.bolEsquina = False And DatosCliente.bolDomicilioLibre = False Then
                            DatosCliente.Domicilio = DatosCliente.Domicilio + " " + CStr(DatosCliente.Numero)
                        Else
                            If DatosCliente.NombreCalle1 <> "" And DatosCliente.NombreCalle1 <> " " Then
                                DatosCliente.Domicilio = DatosCliente.Domicilio & " Y " & DatosCliente.NombreCalle1
                            Else
                                If DatosCliente.NombreCalle1 <> "" And DatosCliente.NombreCalle1 <> " " Then DatosCliente.Domicilio = DatosCliente.Domicilio & " Y " & DatosCliente.NombreCalle2
                            End If
                        End If
                        If Not IsNothing(DatosCliente.Ubicacion) And DatosCliente.Ubicacion <> "" Then DatosCliente.Domicilio = DatosCliente.Domicilio + " " + DatosCliente.Ubicacion

                    End If

                End If
            End If

            'Obtengo la Basde y la Zona

            DatosCliente.CodZona = Zona.ObtenerZona(DatosCliente.Latitud, DatosCliente.Longitud)
            DatosCliente.CodBase = Base.ObtenerBase(DatosCliente.Latitud, DatosCliente.Longitud)


            Regs.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            DatosCliente.NroCliente = 0
            Status = 1
        End Try

        Return DatosCliente

    End Function

End Module
