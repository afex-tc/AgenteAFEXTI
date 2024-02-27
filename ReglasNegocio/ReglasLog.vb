Imports System.Net

Namespace ReglasLog

    Public Class ReglasLog

        ''' <summary>
        ''' Se comunica con el servicio de log para la inserción
        ''' </summary>
        ''' <param name="Contenido"></param>
        ''' <param name="TipoLog">1: Info; 2: Warning; 3: Error</param>
        ''' <param name="VerbosidadLog"></param>
        ''' <param name="VerbosidadRequerida"></param>
        Public Shared Sub CrearLog(Servicio As String, Contenido As String, TipoLog As TipoLog, VerbosidadLog As Integer, VerbosidadRequerida As Integer)

            If Not CorrespondeLog(VerbosidadLog, VerbosidadRequerida) Then Exit Sub

            Select Case TipoLog
                Case TipoLog.Info  ' info
                    InsertarLog(Contenido, IESB_AFEX_ServicioLogger.LogDatadogStatus.Info, Servicio)

                Case TipoLog.Warning  ' warning
                    InsertarLog(Contenido, IESB_AFEX_ServicioLogger.LogDatadogStatus.Warn, Servicio)

                Case TipoLog.Err  ' error
                    InsertarLog(Contenido, IESB_AFEX_ServicioLogger.LogDatadogStatus.Error, Servicio)

            End Select

        End Sub

        ''' <summary>
        ''' Inserta log en DATADOG a través de ESB_AFEX
        ''' </summary>
        ''' <param name="Descripcion"></param>
        ''' <param name="Estado"></param>
        Public Shared Sub InsertarLog(Descripcion As String, Estado As IESB_AFEX_ServicioLogger.LogDatadogStatus, Servicio As String)

            Try
                Dim _log As New IESB_AFEX_ServicioLogger.IESB_AFEX_ServicioLogger
                Dim _logData As New IESB_AFEX_ServicioLogger.LogDataDog
                _logData.Contenido = Descripcion
                _logData.Duration = 0
                _logData.FechaInicio = Now
                _logData.FechaInicio = Now
                _logData.LogDatadogStatus = Estado
                _logData.Servicio = Servicio
                _logData.Source = IESB_AFEX_ServicioLogger.LogDatadogSource.DEFAULT

                _log.InsertarLogDataDogHOST(_logData, "AgenteAFEXTI (v1)", Dns.GetHostName())

            Catch ex As Exception

            End Try
        End Sub

        ''' <summary>
        ''' Devuelve el tipo de origen que corresponde para DATADOG
        ''' </summary>
        ''' <param name="Sitio"></param>
        ''' <returns></returns>
        Public Shared Function DevolverSourceDD(Sitio As String) As IESB_AFEX_ServicioLogger.LogDatadogSource
            Dim _source As IESB_AFEX_ServicioLogger.LogDatadogSource = IESB_AFEX_ServicioLogger.LogDatadogSource.DEFAULT

            Select Case Sitio.ToUpper
                Case "AFEXCHANGEWEB"
                    _source = IESB_AFEX_ServicioLogger.LogDatadogSource.AFEXCHANGEWEB
                Case "EAFEX"
                    _source = IESB_AFEX_ServicioLogger.LogDatadogSource.EAFEX
                Case "EAFEXMONEYGRAM"
                    _source = IESB_AFEX_ServicioLogger.LogDatadogSource.EAFEX
                Case "MONEYGRAM"
                    _source = IESB_AFEX_ServicioLogger.LogDatadogSource.MONEYGRAM
                Case "AFEX.ENTERPRISESERVICEBUS"
                    _source = IESB_AFEX_ServicioLogger.LogDatadogSource.AFEXENTERPRISESERVICEBUS
                Case "ESB_AFEX"
                    _source = IESB_AFEX_ServicioLogger.LogDatadogSource.ESB_AFEX
            End Select

            Return _source

        End Function

        ''' <summary>
        ''' Verifica si corresponde grabar log según la verboidad
        ''' </summary>
        ''' <param name="VerbosidadLog"></param>
        ''' <param name="VerbosidadRequerida"></param>
        ''' <returns></returns>
        Public Shared Function CorrespondeLog(VerbosidadLog As Integer, VerbosidadRequerida As Integer) As Boolean
            Dim _resultado As Boolean = False

            If VerbosidadLog >= VerbosidadRequerida Then _resultado = True

            Return _resultado
        End Function


        Public Enum TipoLog
            Info = 1
            Warning = 2
            Err = 3
        End Enum

    End Class

End Namespace