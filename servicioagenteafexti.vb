Imports System.ComponentModel
Imports System.Configuration
Imports System.Linq.Expressions
Imports System.Windows.Forms
Imports AgenteAFEXTI.ReglasLog.ReglasLog
Imports LibreriaClases.General

Public Class servicioagenteafexti
    Private WithEvents TimerAgente As New Timers.Timer

    Protected Overrides Sub OnStart(ByVal args() As String)
        Me.TimerAgente = New Timers.Timer(ConfigurationManager.AppSettings("IntervaloTimer").ToString)
        Me.TimerAgente.AutoReset = True
        Me.TimerAgente.Start()

        InsertarLog("Servicio Iniciado", IESB_AFEX_ServicioLogger.LogDatadogStatus.Info, "Iicio/Detención")

    End Sub

    Protected Sub TimerAgente_Elapsed(ByVal sender As Object, e As EventArgs) Handles TimerAgente.Elapsed

        'Me.TimerAgente.Stop()

        Ejecutar()
    End Sub

    Public Shared Sub Ejecutar()
        Dim _agente As New ReglasAgente.ReglasAgente
        Dim _verbosidadLogServicio As Integer
        Dim _consultarBBDD As Boolean = False
        Dim _contexto As String = ""
        Dim _servicio As String = "TICK"

        Try
            _verbosidadLogServicio = ReglasAgente.ReglasAgente.ValorConfiguracionServicio(ReglasAgente.ReglasAgente.TipoConfiguracionServicio.verbosidadlogservicio) 'ConfigurationManager.AppSettings("VerbosidadLogServicio").ToString()

            _contexto = "Validando si se consulta a BBDD"
            CrearLog(_servicio, _contexto & " | Inicio", TipoLog.Info, _verbosidadLogServicio, 2)

            _consultarBBDD = ReglasAgente.ReglasAgente.ValorConfiguracionServicio(ReglasAgente.ReglasAgente.TipoConfiguracionServicio.consultarbbdd)

        Catch ex As Exception
            CrearLog(_servicio, _contexto & " | Error : " & ex.Message &
                " - " & ex.StackTrace, TipoLog.Err, 1, 1)

        Finally
            CrearLog(_servicio, _contexto & " | Fin (resultado : " & _consultarBBDD.ToString() & ")", TipoLog.Info, _verbosidadLogServicio, 1)
        End Try

        If Not _consultarBBDD Then GoTo Salir

        Try
            _contexto = "Buscar procesos pendientes para ejecutar"
            CrearLog(_servicio, _contexto & " | Inicio (BBDD) ", TipoLog.Info, _verbosidadLogServicio, 2)

            _agente.VerificarColasPegasdas()

            _agente.VerificarProgramacionHOST()

        Catch ex As Exception
            CrearLog(_servicio, _contexto & " | Error : " & ex.Message &
                " - " & ex.StackTrace, TipoLog.Err, _verbosidadLogServicio, 1)
        Finally
            CrearLog(_servicio, _contexto & " | Fin (BBDD) ", TipoLog.Info, _verbosidadLogServicio, 2)
        End Try

Salir:
    End Sub

    Protected Overrides Sub OnStop()
        InsertarLog("Servicio Detenido", IESB_AFEX_ServicioLogger.LogDatadogStatus.Info, "Iicio/Detención")
    End Sub

End Class
