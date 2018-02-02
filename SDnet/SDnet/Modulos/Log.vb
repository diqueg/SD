Public Module Log

    Public GrabarLog As Boolean = True


    Public Sub InicializarLog(ByVal LogFileName As String, ByVal LogFilePath As String, ByVal Append As Boolean)

        My.Application.Log.DefaultFileLogWriter.Location = Logging.LogFileLocation.Custom
        My.Application.Log.DefaultFileLogWriter.BaseFileName = LogFileName
        My.Application.Log.DefaultFileLogWriter.CustomLocation = LogFilePath
        My.Application.Log.DefaultFileLogWriter.LogFileCreationSchedule = Logging.LogFileCreationScheduleOption.Daily
        My.Application.Log.DefaultFileLogWriter.Append = Append
        My.Application.Log.DefaultFileLogWriter.AutoFlush = True

    End Sub


    Public Sub Grabar(ByVal Datos As String)

        Try
            My.Application.Log.DefaultFileLogWriter.WriteLine(Now.ToString & "; " & Datos)
            My.Application.Log.DefaultFileLogWriter.Flush()
        Catch ex As Exception

        End Try

        Application.DoEvents()

    End Sub

  
End Module
