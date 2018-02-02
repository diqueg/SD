Imports System.Data.SqlClient
Imports System.Net.HttpWebRequest
Imports System.Net
Imports System.Text
Imports System
Imports System.IO

Public Module MovilPush
    Public ConnString As String
    Public Status As Integer

    Public Class DatosMovilPush
        Public IdMovil As Int16
        Public Imei As String
        Public IP As String
    End Class

    Public Function BuscarIdMovilPorImei(ByVal imei As String) As Integer
        Dim NroMovil As Integer
        Try

            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("Select* from Moviles where imei = @Imei", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@Imei", imei)
            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            Do While Regs.Read
                NroMovil = Regs("nro_movil")
                Exit Do
            Loop

            Regs.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            Status = 1
        End Try

        Return NroMovil

    End Function

    Public Function ObtenerDatos(idMovil As Integer) As DatosMovilPush

        Dim DatosPush As New DatosMovilPush

        Try
            Dim DBConn = New SqlConnection(ConnString)
            DBConn.Open()

            Dim SqlCmd As New SqlCommand("Select* from Moviles where nro_movil = @Movil", DBConn)

            SqlCmd.Parameters.AddWithNullableValue("@Movil", idMovil)
            SqlCmd.CommandType = CommandType.Text

            Dim Regs = SqlCmd.ExecuteReader

            Do While Regs.Read
                DatosPush.IdMovil = Regs("nro_movil")
                DatosPush.Imei = IsNull(Regs("imei"))
                DatosPush.IP = IsNull(Regs("IP"))
                Exit Do
            Loop

            Regs.Close()

            If DBConn.State = ConnectionState.Open Then
                DBConn.Close()
            End If

            DBConn.Dispose()

            Status = 0

        Catch ex As Exception
            Log.Grabar("[ERR ];Error en " & System.Reflection.MethodInfo.GetCurrentMethod().ToString & ":" & ex.Message & " en " & ex.StackTrace)
            Status = 1
        End Try

        Return DatosPush


    End Function

    Public Function EnviarComandoMqtt(imei As String, Mensaje As String, Topico As String) As String

        Dim Retorno As String

        Try
            '//creamos la forma como se enviara los datos, en este caso “POST”

            'Dim url As String = "http://138.121.160.232:62907/mqtt/sendto"
            Dim url As String = "http://www.crayonweb.com.ar/taxivoy/pushall.php"

            Dim request As System.Net.WebRequest = System.Net.WebRequest.Create(url)
            request.Method = "POST"

            '//variable para tomar los datos del parámetro data que llega a la funcion
            Dim postData = "imei=" & imei & "&mensaje=" & Mensaje & "&topico=" & Topico

            '//arreglo de bytes codificados con UTF8 para enviar los datos
            Dim byteArray As Byte() = Encoding.UTF8.GetBytes(postData)
            request.ContentType = "application/x-www-form-urlencoded"
            request.ContentLength = byteArray.Length

            Dim dataStream As Stream = request.GetRequestStream()
            dataStream.Write(byteArray, 0, byteArray.Length)
            dataStream.Close()

            '//con esto manejamos las respuestas del servidor
            Dim response As WebResponse = request.GetResponse()
            dataStream = response.GetResponseStream()
            Dim reader As New StreamReader(dataStream)
            Dim responseFromServer As String = reader.ReadToEnd()
            reader.Close()
            dataStream.Close()
            response.Close()

            Retorno = responseFromServer
        Catch ex As Exception
            Dim error1 As String = ErrorToString()
            'If error1 = "Direccion Invalida!!: el formato no ha sido determinado." Then
            'MsgBox("ERROR! Debe de tener  http:// antes de la URL")
            'Else
            'MsgBox(error1)
            'End If
            Retorno = "ERROR:" & error1
        End Try
        If Retorno <> "OK" Then
            frmControl.ActualizarDisplay(0, "Error en TX Push", Retorno)
        End If
        Return Retorno
    End Function


End Module
