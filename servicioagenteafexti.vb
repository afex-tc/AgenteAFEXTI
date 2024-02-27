Imports System.Configuration
Imports AgenteAFEXTI.ReglasLog.ReglasLog

Public Class servicioagenteafexti
    Private WithEvents TimerAgente As New Timers.Timer

    Protected Overrides Sub OnStart(ByVal args() As String)
        Me.TimerAgente = New Timers.Timer(ConfigurationManager.AppSettings("IntervaloTimer").ToString)
        Me.TimerAgente.AutoReset = True
        Me.TimerAgente.Start()

        InsertarLog("Servicio Iniciado", IESB_AFEX_ServicioLogger.LogDatadogStatus.Info, "Servicio Windows")

    End Sub

    Protected Sub TimerAgente_Elapsed(ByVal sender As Object, e As EventArgs) Handles TimerAgente.Elapsed


        'Me.TimerAgente.Stop()

        Ejecutar()

    End Sub

    Protected Sub Ejecutar()
        Dim _agente As New ReglasAgente.ReglasAgente
        InsertarLog("Servicio (tick)", IESB_AFEX_ServicioLogger.LogDatadogStatus.Info, "Servicio Windows")
        Try
            _agente.VerificarColasPegasdas()

        Catch ex As Exception
            InsertarLog("Servicio (tick) - ERROR: " & ex.Source.ToString() & ". " & ex.Message.ToString(),
                        IESB_AFEX_ServicioLogger.LogDatadogStatus.Error, "Servicio Windows")
        End Try

        _agente.VerificarProgramacionHOST()
    End Sub

    Protected Overrides Sub OnStop()
        InsertarLog("Servicio Detenido", IESB_AFEX_ServicioLogger.LogDatadogStatus.Info, "Servicio Windows")
    End Sub

End Class
