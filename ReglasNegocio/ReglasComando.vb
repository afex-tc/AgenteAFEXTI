Imports LibreriaClases.General
Imports System.Net
Imports AgenteAFEXTI.ReglasCorreo.ReglasCorreo
Imports AgenteAFEXTI.ReglasProgramacion
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.CodeDom.Compiler
Imports System.IO
Imports LibreriaClases.IO

Public Class ReglasComando

    Public Shared Function EjecutarComando(Programacion As ReglasProgramacion.Programacion, Optional Contexto As String = "") As ResponseObject
        Dim _resultado As New ResponseObject
        Dim _mensajeError As String = ""
        Dim _cuerpoCorreo As String = ""
        Dim _proceso As String = Programacion.DescripcionProgramacion
        Dim _fecha As String = Right("00" + Now.Day.ToString(), 2) & Right("00" + Now.Month.ToString(), 2) &
                                    Now.Year.ToString() ' fecha de hoy DDMMYYYY

        Try
            ' lo rpimero es crear un BAT y len el bat agrega el comando y luego lo ejecuta
            ReglasLog.ReglasLog.CrearLog(_proceso, Contexto & " | Ejecutar comando | Inicio", ReglasLog.ReglasLog.TipoLog.Info,
                                             Programacion.VerbosidadLog, 2)

            _resultado = CrearBAT(Programacion.DescripcionProgramacion, Contexto, Programacion.Comando, Programacion.VerbosidadLog)

            If _resultado.IsError Then

                ReglasLog.ReglasLog.CrearLog(_proceso, Contexto & " | Ejecutar comando | ERROR: " &
                                                 _resultado.ErrorInfo.Origen & ". " & _resultado.ErrorInfo.Mensaje & ". " &
                                                 _resultado.ErrorInfo.Detalle, ReglasLog.ReglasLog.TipoLog.Err,
                                             Programacion.VerbosidadLog, 1)
            End If

        Catch ex As Exception
            ReglasLog.ReglasLog.CrearLog(_proceso, Contexto & " | Ejecutar comando | " & ex.Source & ". " & ex.Message & ". " & ex.StackTrace, 3, Programacion.VerbosidadLog, 1)

            _cuerpoCorreo = "<table><tr>" &
                                "<td>Error al procesar " & Programacion.DescripcionProgramacion & "</td></tr>" &
                                "<tr><td>Origen</td><td>" & ex.Source & "</td></tr>" &
                                "<tr><td>Detalle</td><td>" & ex.Message & "</td></tr>" &
                                "<tr><td>Stack</td><td>" & ex.StackTrace & "</td></tr>" &
                                "<tr></table>"

            EnviarCorreo(Programacion.CorreoConfirmacion, Programacion.CorreoError, "AgenteAFEXTI - " & Dns.GetHostName() & " - " & Programacion.DescripcionProgramacion & " - CON ERRORES ",
                             _cuerpoCorreo)
        End Try
Fin:
        Return _resultado
    End Function

    Public Shared Function CrearBAT(ByVal Proceso As String, ByVal Contexto As String, ByVal pComando As Comando, ByVal VerbosidadLog As Integer) As ResponseObject
        Dim _resultado As New ResponseObject
        Dim _contextoInterno As String = ""

        Try
            _contextoInterno = "Crear bat"
            ReglasLog.ReglasLog.CrearLog(Proceso, Contexto & " | " & _contextoInterno & " | Inicio", ReglasLog.ReglasLog.TipoLog.Info,
                                             VerbosidadLog, 2)

            ' crea el bat
            Dim _bat As Object
            Dim _archivo As Object = CreateObject("Scripting.FileSystemObject")
            Dim _rutaArchivo As String = AppDomain.CurrentDomain.BaseDirectory() & "\comandoagenteafexti.bat"

            _bat = _archivo.CreateTextFile(_rutaArchivo)

            Dim _scriptComando As String = pComando.ScriptComando
            _scriptComando = _scriptComando.Replace("<<PCREMOTO>>", pComando.NombreHostRemoto)
            Dim _lineasComando As String() = _scriptComando.Split(";")
            For i = 0 To _lineasComando.Count() - 1
                _bat.WriteLine(_lineasComando(i))
            Next
            _bat.close()

            ReglasLog.ReglasLog.CrearLog(Proceso, Contexto & " | " & _contextoInterno & " | Fin", ReglasLog.ReglasLog.TipoLog.Info,
                                             VerbosidadLog, 2)

            _contextoInterno = "Ejecutar bat"
            ReglasLog.ReglasLog.CrearLog(Proceso, Contexto & " | " & _contextoInterno & " | Inicio", ReglasLog.ReglasLog.TipoLog.Info,
                                             VerbosidadLog, 2)

            Dim _proceso As New ProcessStartInfo(_rutaArchivo)
            _proceso.RedirectStandardError = True
            _proceso.RedirectStandardOutput = True
            _proceso.CreateNoWindow = False
            _proceso.WindowStyle = ProcessWindowStyle.Hidden
            _proceso.UseShellExecute = False
            Dim _ejecutorproceso As Process = Process.Start(_proceso)


        Catch ex As Exception
            ReglasLog.ReglasLog.CrearLog(Proceso, Contexto & " | " & _contextoInterno & " | Error : " & ex.Message & " - " & ex.StackTrace, ReglasLog.ReglasLog.TipoLog.Err, 1, 1)

        Finally
            ReglasLog.ReglasLog.CrearLog(Proceso, Contexto & " | " & _contextoInterno & " | Fin", ReglasLog.ReglasLog.TipoLog.Info,
                                             VerbosidadLog, 2)
        End Try

        Return _resultado
    End Function

#Region "Estructuras"

    Public Structure Comando
        Dim IdComando As Integer
        Dim TipoComando As String
        Dim DescripcionComando As String
        Dim ScriptComando As String
        Dim NombreHostRemoto As String
    End Structure

#End Region


End Class
