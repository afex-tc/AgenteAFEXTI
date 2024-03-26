Imports LibreriaClases.Datos.SQL
Imports LibreriaClases.General
Imports LibreriaClases.Errores
Imports System.Configuration
Imports AgenteAFEXTI.ReglasCorreo.ReglasCorreo
Imports AgenteAFEXTI.ReglasProgramacion
Imports AgenteAFEXTI.ReglasComando
Imports System.Net
Imports System.IO
Imports System.Xml

Namespace ReglasAgente

    Public Class ReglasAgente

        ''' <summary>
        ''' Verifica si existen colas pegadas para informar y que se puedan liberar
        ''' </summary>
        ''' <returns></returns>
        Public Function VerificarColasPegasdas() As ResponseObject
            Dim _resultado As New ResponseObject
            Dim _conexion As CapaDatosSql

            Try
                ReglasLog.ReglasLog.CrearLog("VerificaColasPegadas", "Inicio", ReglasLog.ReglasLog.TipoLog.Info, 1, 2)

                Dim _comando As New SqlClient.SqlCommand("Programaciones.MostrarColasPegadas")
                _comando.CommandType = CommandType.StoredProcedure

                _conexion = New CapaDatosSql(ConfigurationManager.ConnectionStrings("Programaciones").ToString, False)

                _conexion.conectar()

                Dim _reader As SqlClient.SqlDataReader
                _reader = _conexion.EjecutarComandoReader(_comando)

                If _reader.Read Then
                    ReglasLog.ReglasLog.CrearLog("VerificaColasPegadas", "Existen colas pegadas, se debe verificar en BBDD", ReglasLog.ReglasLog.TipoLog.Err,
                                                 1, 1)
                    EnviarCorreo(ConfigurationManager.AppSettings("CorreosNotificacionErrores"), "", "AFEXAgenteTI - CON ERRORES - " & Dns.GetHostName() &
                             " - VerificaColasPegadas", "Existen colas pegadas, se debe verificar en BBDD")
                End If
                _conexion.desconectar()

            Catch ex As Exception
                _resultado = New ResponseObject(New ErrorInfo(ex.Message, ex.StackTrace, 0, ex.Source))
                ReglasLog.ReglasLog.CrearLog("VerificaColasPegadas", _resultado.ErrorInfo.Origen & ". " &
                                             _resultado.ErrorInfo.Mensaje & ". " & _resultado.ErrorInfo.Detalle,
                                             ReglasLog.ReglasLog.TipoLog.Err, 1, 1)

                EnviarCorreo(ConfigurationManager.AppSettings("CorreosNotificacionErrores"), "", "AFEXAgenteTI - CON ERRORES - " & Dns.GetHostName() &
                             " - VerificaColasPegadas", ex.Source & ". " & ex.Message & ". " & ex.StackTrace)

            Finally
                _conexion.desconectar()
            End Try

            ReglasLog.ReglasLog.CrearLog("VerificaColasPegadas", "Fin", 1, 1, 2)

            Return _resultado
        End Function

        ''' <summary>
        ''' Verifica qué acción debe ejecutar en este momento el HOST donde se encunetra instalado el AGENTE
        ''' </summary>
        ''' <returns></returns>
        Public Function VerificarProgramacionHOST() As ResponseObject
            Dim _resultado As New ResponseObject
            Dim _conexion As CapaDatosSql
            Dim _contexto As String = ""
            Dim _servicio As String = "VerificaProgramacion"

            Try
                ReglasLog.ReglasLog.CrearLog(_servicio, "Inicio", ReglasLog.ReglasLog.TipoLog.Info, 1, 2)

                Dim _comando As New SqlClient.SqlCommand("Programaciones.IniciarColaProgramacionDisponible")
                _comando.CommandType = CommandType.StoredProcedure

                _conexion = New CapaDatosSql(ConfigurationManager.ConnectionStrings("Programaciones").ToString, False)

                _comando.Parameters.Add("iniciar", SqlDbType.Int)
                _comando.Parameters(0).Value = 1

                _conexion.conectar()

                Dim _programacion As New ReglasProgramacion.Programacion(_conexion.EjecutarComandoReader(_comando))

                _conexion.desconectar()

                If Not _programacion Is Nothing Then
                    If _programacion.IdColaProgramacion > 0 Then
                        Try

                            ReglasLog.ReglasLog.CrearLog(_programacion.DescripcionProgramacion, "Inicio cola programación: " & _programacion.IdColaProgramacion.ToString(),
                                                          ReglasLog.ReglasLog.TipoLog.Info, _programacion.VerbosidadLog, 2)


                            Select Case _programacion.IdTipoProgramacion
                                Case TipoProgramacion.CopiadoArchivos ' copiar
                                    ReglasLog.ReglasLog.CrearLog(_programacion.DescripcionProgramacion, "Copiar Archivos | Inicio ",
                                                                 ReglasLog.ReglasLog.TipoLog.Info, _programacion.VerbosidadLog, 2)

                                    _resultado = CopiarArchivos(_programacion, "Copiar Archivos")

                                    If _resultado.IsError Then
                                        ReglasLog.ReglasLog.CrearLog(_programacion.DescripcionProgramacion, "Copiar Archivos | ERROR : " &
                                                                     _resultado.ErrorInfo.Origen & ". " & _resultado.ErrorInfo.Mensaje & ". " & _resultado.ErrorInfo.Detalle,
                                                                 ReglasLog.ReglasLog.TipoLog.Err, _programacion.VerbosidadLog, 1)

                                    End If

                                    ReglasLog.ReglasLog.CrearLog(_programacion.DescripcionProgramacion, "Copiar Archivos | Fin ",
                                                                 ReglasLog.ReglasLog.TipoLog.Info, _programacion.VerbosidadLog, 2)


                                Case TipoProgramacion.Publicacion ' publicar
                                    Select Case _programacion.TipoColaProgramacion
                                        Case TipoColaProgramacion.Validar

                                            ReglasLog.ReglasLog.CrearLog(_programacion.DescripcionProgramacion,
                                                                         "VALIDAR PUBLICACIONES PENDIENTES | Inicio",
                                                                         ReglasLog.ReglasLog.TipoLog.Info,
                                                                         _programacion.VerbosidadLog, 2)

                                            _resultado = VerificarRutaOrigen(_programacion, "VALIDAR PUBLICACIONES PENDIENTES")

                                            If Not _resultado.IsError And _resultado.Objeto = True Then
                                                ' crea una cola de copia programada para su posterior autorización
                                                _comando = New SqlClient.SqlCommand("Programaciones.InsertarCopiaProgramada")
                                                _comando.CommandType = CommandType.StoredProcedure

                                                _conexion = New CapaDatosSql(ConfigurationManager.ConnectionStrings("Programaciones").ToString, False)

                                                _comando.Parameters.Add("idcolaprogramacion", SqlDbType.Int)
                                                _comando.Parameters(0).Value = _programacion.IdColaProgramacion

                                                _conexion.conectar()

                                                _conexion.EjecutarComandoReader(_comando)

                                                ReglasLog.ReglasLog.CrearLog(_programacion.DescripcionProgramacion, "VALIDAR PUBLICACIONES PENDIENTES | Se crea programación para publicación pendiente.",
                                                                             ReglasLog.ReglasLog.TipoLog.Info, _programacion.VerbosidadLog, 1)

                                                EnviarCorreo(_programacion.CorreoConfirmacion, "", "AgenteAFEXTI - " & Dns.GetHostName() & " - " &
                                                            _programacion.DescripcionProgramacion,
                                                             "Se ha encontrado una nueva publicación. <br> Favor de autorizar el proceso.")

                                                _conexion.desconectar()

                                            ElseIf Not _resultado.IsError And _resultado.Objeto = 0 Then ' no existe el origen
                                                ReglasLog.ReglasLog.CrearLog(_programacion.DescripcionProgramacion, "VALIDAR PUBLICACIONES PENDIENTES | Sin publicación pendiente.",
                                                                             ReglasLog.ReglasLog.TipoLog.Info, _programacion.VerbosidadLog, 2)

                                            ElseIf _resultado.IsError Then
                                                ReglasLog.ReglasLog.CrearLog(_programacion.DescripcionProgramacion, "VALIDAR PUBLICACIONES PENDIENTES | ERROR: " &
                                                                             _resultado.ErrorInfo.Origen & ". " & _resultado.ErrorInfo.Mensaje & ". " & _resultado.ErrorInfo.Detalle,
                                                                  ReglasLog.ReglasLog.TipoLog.Err, _programacion.VerbosidadLog, 1)

                                            End If

                                        Case TipoColaProgramacion.CopiaProgramada
                                            ReglasLog.ReglasLog.CrearLog(_programacion.DescripcionProgramacion, "PUBLICACION PROGRAMADA | Inicio ", ReglasLog.ReglasLog.TipoLog.Info, _programacion.VerbosidadLog, 2)

                                            _resultado = CopiarArchivos(_programacion, " (publicando)")

                                            ' cambia el nombre de la ruta origen para no publicarla
                                            If Not _resultado.IsError And _programacion.CambiarNombreOrigen Then
                                                ReglasLog.ReglasLog.CrearLog(_programacion.DescripcionProgramacion, "PUBLICACION PROGRAMADA - Cambio de nombre origen ",
                                                                             ReglasLog.ReglasLog.TipoLog.Info, _programacion.VerbosidadLog, 2)


                                                Dim _nuevoNombre As String = _programacion.RutaOrigen & "_publicado_" & Year(Now) & Month(Now) & Day(Now)

                                                _resultado = ReglasCopia.ReglasCopia.CambiarNombreDirectorio(_programacion.DescripcionProgramacion,
                                                                _programacion.RutaOrigen, _nuevoNombre, _programacion.VerbosidadLog)

                                                If Not _resultado.IsError Then
                                                    ReglasLog.ReglasLog.CrearLog(_programacion.DescripcionProgramacion, "PUBLICACION PROGRAMADA - Cambio de nombre origen realizado",
                                                                             ReglasLog.ReglasLog.TipoLog.Info, _programacion.VerbosidadLog, 2)

                                                Else
                                                    ReglasLog.ReglasLog.CrearLog(_programacion.DescripcionProgramacion, "PUBLICACION PROGRAMADA - Cambio de nombre origen | ERROR: " &
                                                                             _resultado.ErrorInfo.Origen & ". " & _resultado.ErrorInfo.Mensaje & ". " & _resultado.ErrorInfo.Detalle,
                                                                  ReglasLog.ReglasLog.TipoLog.Err, _programacion.VerbosidadLog, 1)

                                                End If

                                            ElseIf _resultado.IsError Then
                                                ReglasLog.ReglasLog.CrearLog(_programacion.DescripcionProgramacion, "PUBLICACION PROGRAMADA | ERROR: " &
                                                                             _resultado.ErrorInfo.Origen & ". " & _resultado.ErrorInfo.Mensaje & ". " & _resultado.ErrorInfo.Detalle,
                                                                  ReglasLog.ReglasLog.TipoLog.Err, _programacion.VerbosidadLog, 1)
                                            End If


                                            ReglasLog.ReglasLog.CrearLog(_programacion.DescripcionProgramacion, "PUBLICACION PROGRAMADA | Fin", 1, _programacion.VerbosidadLog, 1)

                                    End Select

                                Case TipoProgramacion.EjecucionComando ' ejecución de comando
                                    _contexto = "Ejecución de comandos"
                                    ReglasLog.ReglasLog.CrearLog(_programacion.DescripcionProgramacion, _contexto & " | Inicio ",
                                                                 ReglasLog.ReglasLog.TipoLog.Info, _programacion.VerbosidadLog, 2)

                                    _resultado = EjecutarComando(_programacion, _contexto)

                                    If _resultado.IsError Then
                                        ReglasLog.ReglasLog.CrearLog(_programacion.DescripcionProgramacion, _contexto & " | ERROR : " &
                                                                     _resultado.ErrorInfo.Origen & ". " & _resultado.ErrorInfo.Mensaje & ". " & _resultado.ErrorInfo.Detalle,
                                                                 ReglasLog.ReglasLog.TipoLog.Err, _programacion.VerbosidadLog, 1)

                                    End If

                                    ReglasLog.ReglasLog.CrearLog(_programacion.DescripcionProgramacion, _contexto & " | Fin ",
                                                                 ReglasLog.ReglasLog.TipoLog.Info, _programacion.VerbosidadLog, 1)


                                Case TipoProgramacion.EliminarInstaladorInicio

                                    _contexto = "Eliminando bat"
                                    ReglasLog.ReglasLog.CrearLog(_programacion.DescripcionProgramacion, _contexto & " | Inicio ",
                                                                 ReglasLog.ReglasLog.TipoLog.Info, _programacion.VerbosidadLog, 2)

                                    _resultado = EliminarInstaladorInicio(_programacion)

                                    If _resultado.IsError Then
                                        ReglasLog.ReglasLog.CrearLog(_programacion.DescripcionProgramacion, _contexto & " | ERROR : " &
                                                                     _resultado.ErrorInfo.Origen & ". " & _resultado.ErrorInfo.Mensaje & ". " & _resultado.ErrorInfo.Detalle,
                                                                 ReglasLog.ReglasLog.TipoLog.Err, _programacion.VerbosidadLog, 1)

                                    End If

                                    ReglasLog.ReglasLog.CrearLog(_programacion.DescripcionProgramacion, _contexto & " | Fin ",
                                                                 ReglasLog.ReglasLog.TipoLog.Info, _programacion.VerbosidadLog, 1)



                                Case Else
                                    ReglasLog.ReglasLog.CrearLog(_programacion.DescripcionProgramacion, "Se encuentra programación sin tipo conocido " & _programacion.IdColaProgramacion.ToString(),
                                                                  ReglasLog.ReglasLog.TipoLog.Err, _programacion.VerbosidadLog, 1)


                            End Select

                            ' termina la cola
                            _comando = New SqlClient.SqlCommand("Programaciones.TerminarColaProgramacion")
                            _comando.CommandType = CommandType.StoredProcedure
                            _conexion = New CapaDatosSql(ConfigurationManager.ConnectionStrings("Programaciones").ToString, False)
                            _comando.Parameters.AddWithValue("IdColaProgramacion", _programacion.IdColaProgramacion)
                            _conexion.conectar()
                            _conexion.EjecutarComandoReader(_comando)
                            _conexion.desconectar()

                            ReglasLog.ReglasLog.CrearLog(_programacion.DescripcionProgramacion, "Fin | cola de programación : " & _programacion.IdColaProgramacion.ToString(),
                                                 ReglasLog.ReglasLog.TipoLog.Info, _programacion.VerbosidadLog, 2)


                        Catch ex As Exception
                            ReglasLog.ReglasLog.CrearLog(_programacion.DescripcionProgramacion, "Inicio | cola de programación : " & _programacion.IdColaProgramacion.ToString() &
                                         " | ERROR " & ex.Source & ". " & ex.Message & ". " & ex.StackTrace, ReglasLog.ReglasLog.TipoLog.Err,
                                         _programacion.VerbosidadLog, 1)

                            EnviarCorreo(ConfigurationManager.AppSettings("CorreosNotificacionErrores"), "", "AFEXAgenteTI - CON ERRORES - " &
                                         Dns.GetHostName() & " - " & _programacion.DescripcionProgramacion, ex.Source & ". " & ex.Message & ". " &
                                         ex.StackTrace)
                        End Try

                    Else
                        ReglasLog.ReglasLog.CrearLog(_servicio, "Sin programación para ejecutar", ReglasLog.ReglasLog.TipoLog.Info, 1, 2)
                    End If
                End If


            Catch ex As Exception
                _resultado = New ResponseObject(New ErrorInfo(ex.Message, ex.StackTrace, 0, ex.Source))
                ReglasLog.ReglasLog.CrearLog(_servicio, _resultado.ErrorInfo.Origen & ". " &
                                             _resultado.ErrorInfo.Mensaje & ". " & _resultado.ErrorInfo.Detalle,
                                             ReglasLog.ReglasLog.TipoLog.Err, 1, 1)

                EnviarCorreo(ConfigurationManager.AppSettings("CorreosNotificacionErrores"), "", "AFEXAgenteTI - CON ERRORES - " & Dns.GetHostName() &
                             " - VerificaProgramacion", ex.Source & ". " & ex.Message & ". " & ex.StackTrace)

            Finally
                _conexion.desconectar()

                ReglasLog.ReglasLog.CrearLog(_servicio, "Fin", 1, 1,
                                             ValorConfiguracionServicio(TipoConfiguracionServicio.verbosidadlogservicio))
            End Try

            Return _resultado
        End Function

        ''' <summary>
        ''' Copia archivos de una ruta a otra
        ''' </summary>
        ''' <param name="Programacion"></param>
        ''' <returns></returns>
        Public Function CopiarArchivos(Programacion As ReglasProgramacion.Programacion, Optional Etiqueta As String = "") As ResponseObject
            Dim _resultado As New ResponseObject
            Dim _mensajeError As String
            Dim _proceso As String = Programacion.DescripcionProgramacion
            Dim _estadoCorreo As String = "SIN ERRORES"
            Dim _cuerpoCorreo As String = ""
            Dim _intentosCopia As Integer = 1 ' cantidad de veces que tratará de copiar en caso que al final del proceso se encuentren diferencias
            Dim _archivosCopiados As Integer
            Dim _archivosLargos As Integer
            Dim _directoriosLargos As Integer
            Dim _archivosExcluidos As Integer
            Dim _directoriosExcluidos As Integer
            Dim _fecha As String = Right("00" + Now.Day.ToString(), 2) & Right("00" + Now.Month.ToString(), 2) &
                                    Now.Year.ToString() ' fecha de hoy DDMMYYYY

            Try
                ' antes de copiar se valida que existan archivos en el origen
                ReglasLog.ReglasLog.CrearLog(_proceso, Etiqueta & " | Validar Archivos Origen | Inicio", ReglasLog.ReglasLog.TipoLog.Info,
                                             Programacion.VerbosidadLog, 2)

                _resultado = ReglasCopia.ReglasCopia.ExistenArchivos(Programacion.DescripcionProgramacion,
                                                                Etiqueta, Programacion.RutaOrigen, Programacion.VerbosidadLog)

                If _resultado.IsError Then

                    ReglasLog.ReglasLog.CrearLog(_proceso, Etiqueta & " | Validar Archivos Origen | ERROR: " &
                                                 _resultado.ErrorInfo.Origen & ". " & _resultado.ErrorInfo.Mensaje & ". " &
                                                 _resultado.ErrorInfo.Detalle, ReglasLog.ReglasLog.TipoLog.Err,
                                             Programacion.VerbosidadLog, 1)

                ElseIf _resultado.Objeto = False Then
                    ReglasLog.ReglasLog.CrearLog(_proceso, Etiqueta & " | Validar Archivos Origen (no hay archivos para copiar)  | Fin",
                                                 ReglasLog.ReglasLog.TipoLog.Info, Programacion.VerbosidadLog, 2)
                    GoTo Fin

                Else
                    ReglasLog.ReglasLog.CrearLog(_proceso, Etiqueta & " | Validar Archivos Origen (si hay archivos para copiar) | Fin",
                                                 ReglasLog.ReglasLog.TipoLog.Info, Programacion.VerbosidadLog, 2)
                End If


                ' respaldo
                If Not Programacion.RutaRespaldo Is Nothing And Programacion.RutaRespaldo <> "" Then
                    ReglasLog.ReglasLog.CrearLog(_proceso, Etiqueta & " | Respaldar | Inicio", ReglasLog.ReglasLog.TipoLog.Info,
                                                 Programacion.VerbosidadLog, 2)

                    Dim _rutaValida As ResponseObject = ReglasCopia.ReglasCopia.ExisteRuta(_proceso, Etiqueta & " | Respaldar ", Programacion.RutaDestino,
                                                          Programacion.VerbosidadLog)
                    If Not _rutaValida.IsError And _rutaValida.Objeto = True Then
                        _archivosCopiados = 0
                        _archivosLargos = 0
                        _directoriosLargos = 0
                        _archivosExcluidos = 0
                        _directoriosExcluidos = 0

                        _cuerpoCorreo = _cuerpoCorreo &
                                    "<tr><td><br></td></tr>" &
                                    "<tr><td><b>Eliminar respaldo</b></td></tr>"

                        _mensajeError = "Eliminar respaldo"

                        ' elimina respaldos
                        Dim _rutaDestinoRespaldo As String = Programacion.RutaRespaldo

                        _resultado = ReglasCopia.ReglasCopia.EliminarRuta(_proceso, Etiqueta & " | Respaldar ", _rutaDestinoRespaldo,
                                                                              Programacion.VerbosidadLog, Programacion.DiasRespaldo)

                        If _resultado.IsError Then
                            ReglasLog.ReglasLog.CrearLog(_proceso, Etiqueta & " Respaldar | ERROR al Eliminar Redpaldo " &
                                                         _resultado.ErrorInfo.Origen & ". " & _resultado.ErrorInfo.Mensaje & ". " &
                                                         _resultado.ErrorInfo.Detalle, ReglasLog.ReglasLog.TipoLog.Err, Programacion.VerbosidadLog, 1)

                            _estadoCorreo = "CON ERRORES"
                            _cuerpoCorreo = _cuerpoCorreo &
                            "<tr><td>Error en el proceso</td></tr>" &
                            "<tr><td>Origen</td><td>" & _resultado.ErrorInfo.Origen & "</td></tr>" &
                            "<tr><td>Detalle</td><td>" & _resultado.ErrorInfo.Mensaje & "</td></tr>" &
                            "<tr><td>Stack</td><td>" & _resultado.ErrorInfo.Detalle & "</td></tr>"
                        Else
                            ReglasLog.ReglasLog.CrearLog(_proceso, Etiqueta & " | Respaldar | Respaldo eliminados ",
                                                         ReglasLog.ReglasLog.TipoLog.Info, Programacion.VerbosidadLog, 2)

                            _cuerpoCorreo = _cuerpoCorreo &
                               "<tr><td>Proceso correcto</td></tr>" &
                               "<tr><td>Ruta respaldos eliminados</td><td>" & _rutaDestinoRespaldo.ToString() & "</td></tr>"

                            ' respalda
                            _cuerpoCorreo = _cuerpoCorreo &
                                    "<tr><td><br></td></tr>" &
                                     "<tr><td><b>Respaldar</b></td></tr>"

                            _mensajeError = "Respaldar"
                            _rutaDestinoRespaldo += "\" & _fecha
                            _resultado = ReglasCopia.ReglasCopia.ProcesarArchivos(_proceso, Etiqueta & " | Respaldar ", ReglasCopia.ReglasCopia.TipoAccionArchivos.Copiar,
                                                                          Programacion.RutaDestino, _rutaDestinoRespaldo,
                                                                     Programacion.VerbosidadLog, _archivosCopiados, "", "", "",
                                                                     _archivosLargos, _directoriosLargos, _archivosExcluidos, _directoriosExcluidos)

                            If _resultado.IsError Then
                                ReglasLog.ReglasLog.CrearLog(_proceso, Etiqueta & " | Respaldar | FIN | ERROR " & _archivosCopiados.ToString() &
                                                 " archivos copiados | " & _resultado.ErrorInfo.Origen & ". " &
                                                 _resultado.ErrorInfo.Mensaje & ". " & _resultado.ErrorInfo.Detalle,
                                                  ReglasLog.ReglasLog.TipoLog.Err, Programacion.VerbosidadLog, 1)

                                _estadoCorreo = "CON ERRORES"
                                _cuerpoCorreo = _cuerpoCorreo &
                                "<tr><td>Error en el proceso</td></tr>" &
                                "<tr><td>Origen</td><td>" & _resultado.ErrorInfo.Origen & "</td></tr>" &
                                "<tr><td>Detalle</td><td>" & _resultado.ErrorInfo.Mensaje & "</td></tr>" &
                                "<tr><td>Stack</td><td>" & _resultado.ErrorInfo.Detalle & "</td></tr>"

                            Else
                                ReglasLog.ReglasLog.CrearLog(_proceso, Etiqueta & " | Respaldar | Fin | " & _archivosCopiados.ToString() &
                                                 " archivos copiados", ReglasLog.ReglasLog.TipoLog.Info, Programacion.VerbosidadLog, 2)

                                _cuerpoCorreo = _cuerpoCorreo &
                               "<tr><td>Proceso correcto</td></tr>" &
                               "<tr><td>Archivos copiados</td><td>" & _archivosCopiados.ToString() & "</td></tr>" &
                               "<tr><td>Archivos largos</td><td>" & _archivosLargos.ToString() & "</td></tr>" &
                               "<tr><td>Directorios largos</td><td>" & _directoriosLargos.ToString() & "</td></tr>" &
                               "<tr><td>Archivos excluidos</td><td>" & _archivosExcluidos.ToString() & "(" & Programacion.ArchivosExcluidos & ")</td></tr>" &
                               "<tr><td>Directorios excluidos</td><td>" & _directoriosExcluidos.ToString() & "(" & Programacion.DirectoriosExcluidos & ")</td></tr>"
                            End If
                        End If
                    Else
                        ReglasLog.ReglasLog.CrearLog(_proceso, Etiqueta & " | Respaldar | No existe información para respaldar.", 1, Programacion.VerbosidadLog, 2)
                    End If
                    ReglasLog.ReglasLog.CrearLog(_proceso, Etiqueta & " | Respaldar | Fin ", 1, Programacion.VerbosidadLog, 1)
                End If

                ' eliminar el destino antes de copiar
                If Not _resultado.IsError Then
                    If Programacion.DirectorioFecha Then
                        ReglasLog.ReglasLog.CrearLog(_proceso, Etiqueta & " | Eliminar Destino | Inicio",
                                                    ReglasLog.ReglasLog.TipoLog.Info, Programacion.VerbosidadLog, 2)

                        _resultado = ReglasCopia.ReglasCopia.EliminarRuta(_proceso, Etiqueta & " | Eliminar Destino ",
                                                                          Programacion.RutaDestino,
                                                                          Programacion.VerbosidadLog,
                                                                          Programacion.DiasCopiasAnteriores)

                        If Not _resultado.IsError Then
                            ReglasLog.ReglasLog.CrearLog(_proceso, Etiqueta & " | Eliminar Destino | Correcto",
                                                    ReglasLog.ReglasLog.TipoLog.Info, Programacion.VerbosidadLog, 2)
                        Else
                            ReglasLog.ReglasLog.CrearLog(_proceso, Etiqueta & " | Eliminar Destino | ERROR: " &
                                                         _resultado.ErrorInfo.Origen & ". " & _resultado.ErrorInfo.Mensaje & ". " &
                                                         _resultado.ErrorInfo.Detalle, ReglasLog.ReglasLog.TipoLog.Info,
                                                                                       Programacion.VerbosidadLog, 1)
                        End If

                        ReglasLog.ReglasLog.CrearLog(_proceso, Etiqueta & " | Eliminar Destino | Fin",
                                                    ReglasLog.ReglasLog.TipoLog.Info, Programacion.VerbosidadLog, 2)
                    End If
                End If

                ' copia
Copiar:
                If Not _resultado.IsError Then
                    _archivosCopiados = 0
                    _archivosLargos = 0
                    _directoriosLargos = 0
                    _archivosExcluidos = 0
                    _directoriosExcluidos = 0

                    _cuerpoCorreo = _cuerpoCorreo &
                                    "<tr><td><br></td></tr>" &
                                    "<tr><td><b>Copiar (intento " & _intentosCopia.ToString() & "</b></td></tr>"

                    ReglasLog.ReglasLog.CrearLog(_proceso, Etiqueta & " | Copiar | Inicio (intento " & _intentosCopia.ToString() & ")",
                                                ReglasLog.ReglasLog.TipoLog.Info, Programacion.VerbosidadLog, 2)

                    ' copia los archivos
                    Dim _rutaDestino As String = Programacion.RutaDestino
                    If Programacion.DirectorioFecha Then _rutaDestino += "\" & _fecha

                    _resultado = ReglasCopia.ReglasCopia.ProcesarArchivos(_proceso, Etiqueta & " | Copiar ",
                                                                          ReglasCopia.ReglasCopia.TipoAccionArchivos.Copiar,
                                                                          Programacion.RutaOrigen, _rutaDestino,
                                                                     Programacion.VerbosidadLog, _archivosCopiados,
                                                                     Programacion.ArchivosExcluidos,
                                                                     Programacion.DirectoriosExcluidos, "",
                                                                     _archivosLargos, _directoriosLargos,
                                                                     _archivosExcluidos, _directoriosExcluidos,
                                                                     Programacion.EliminarOrigen, Programacion.RutaRetencionOrigen)

                    If _resultado.IsError Then
                        ReglasLog.ReglasLog.CrearLog(_proceso, Etiqueta & " | Copiar | Fin (intento " & _intentosCopia.ToString() & ") | ERROR " & _archivosCopiados.ToString() &
                                                 " archivos copiados | " & _resultado.ErrorInfo.Origen & ". " &
                                                 _resultado.ErrorInfo.Mensaje & ". " & _resultado.ErrorInfo.Detalle,
                                                  ReglasLog.ReglasLog.TipoLog.Err, Programacion.VerbosidadLog, 1)

                        _estadoCorreo = "CON ERRORES"
                        _cuerpoCorreo = _cuerpoCorreo &
                                "<tr><td>Error en el proceso</td></tr>" &
                                "<tr><td>Origen</td><td>" & _resultado.ErrorInfo.Origen & "</td></tr>" &
                                "<tr><td>Detalle</td><td>" & _resultado.ErrorInfo.Mensaje & "</td></tr>" &
                                "<tr><td>Stack</td><td>" & _resultado.ErrorInfo.Detalle & "</td></tr>"

                    Else
                        ReglasLog.ReglasLog.CrearLog(_proceso, Etiqueta & " | Copiar | Fin (intento " & _intentosCopia.ToString() & ") | " & _archivosCopiados.ToString() &
                                                 " archivos copiados", ReglasLog.ReglasLog.TipoLog.Info, Programacion.VerbosidadLog, 1)

                        _cuerpoCorreo = _cuerpoCorreo &
                               "<tr><td>Proceso correcto</td></tr>" &
                               "<tr><td>Archivos copiados</td><td>" & _archivosCopiados.ToString() & "</td></tr>" &
                               "<tr><td>Archivos largos</td><td>" & _archivosLargos.ToString() & "</td></tr>" &
                               "<tr><td>Directorios largos</td><td>" & _directoriosLargos.ToString() & "</td></tr>" &
                               "<tr><td>Archivos excluidos</td><td>" & _archivosExcluidos.ToString() & "(" & Programacion.ArchivosExcluidos & ")</td></tr>" &
                               "<tr><td>Directorios excluidos</td><td>" & _directoriosExcluidos.ToString() & "(" & Programacion.DirectoriosExcluidos & ")</td></tr>"
                    End If
                End If

                ' valida los totales
                Dim _pesoOrigen As Integer, _directoriosOrigen As Integer, _archivosOrigen As Integer
                Dim _error As LibreriaClases.Errores.ErrorInfo

                _cuerpoCorreo = _cuerpoCorreo &
                                    "<tr><td><br></td></tr>" &
                                     "<tr><td><b>Validar Totales</b></td></tr>"

                _error = ReglasCopia.ReglasCopia.ValidarTotales(Programacion.DescripcionProgramacion, Programacion.RutaOrigen,
                                                       _pesoOrigen, _directoriosOrigen, _archivosOrigen, Programacion.VerbosidadLog,
                                                         Programacion.ArchivosExcluidos, Programacion.DirectoriosExcluidos)

                If Not _error.Detalle <> "" Then
                    ReglasLog.ReglasLog.CrearLog(Programacion.DescripcionProgramacion, Etiqueta & " Verificar Totales Origen | Archivos: " & _archivosOrigen.ToString() &
                                                " | Directorios: " & _directoriosOrigen.ToString() &
                                                " | Peso: " & _pesoOrigen.ToString, ReglasLog.ReglasLog.TipoLog.Info,
                                                 Programacion.VerbosidadLog, 2)

                    _cuerpoCorreo = _cuerpoCorreo &
                               "<tr><td><b>ORIGEN</b></td></tr>" &
                               "<tr><td>Proceso correcto</td></tr>" &
                               "<tr><td>Archivos</td><td>" & _archivosOrigen.ToString() & "</td></tr>" &
                               "<tr><td>Directorios</td><td>" & _directoriosOrigen.ToString() & "</td></tr>" &
                               "<tr><td>Peso</td><td>" & _pesoOrigen.ToString() & "</td></tr>"

                Else

                    ReglasLog.ReglasLog.CrearLog(Programacion.DescripcionProgramacion,
                                                  Etiqueta & " Verificar Totales Origen | ERROR: " & _error.Origen & " - " & _error.Mensaje & " - " & _error.Detalle,
                                                  ReglasLog.ReglasLog.TipoLog.Err, Programacion.VerbosidadLog, 1)

                    _estadoCorreo = "CON ERRORES"
                    _cuerpoCorreo = _cuerpoCorreo &
                            "<tr><td><b>ORIGEN</b></td></tr>" &
                            "<tr><td>Error en el proceso</td></tr>" &
                            "<tr><td>Origen</td><td>" & _resultado.ErrorInfo.Origen & "</td></tr>" &
                            "<tr><td>Detalle</td><td>" & _resultado.ErrorInfo.Mensaje & "</td></tr>" &
                            "<tr><td>Stack</td><td>" & _resultado.ErrorInfo.Detalle & "</td></tr>"
                End If

                ' valida los totales del destino
                Dim _pesoDestino As Integer, _directoriosDestino As Integer, _archivosDestino As Integer
                ReglasCopia.ReglasCopia.ValidarTotales(Programacion.DescripcionProgramacion, Programacion.RutaDestino,
                                                        _pesoDestino, _directoriosDestino, _archivosDestino, Programacion.VerbosidadLog,
                                                        Programacion.ArchivosExcluidos, Programacion.DirectoriosExcluidos)
                If Not _error.Detalle <> "" Then
                    ReglasLog.ReglasLog.CrearLog(Programacion.DescripcionProgramacion, Etiqueta & " Verificar Totales Destino | Archivos: " & _archivosDestino.ToString() &
                                                " | Directorios: " & _directoriosDestino.ToString() &
                                                " | Peso: " & _pesoDestino.ToString, ReglasLog.ReglasLog.TipoLog.Info,
                                                Programacion.VerbosidadLog, 2)

                    _cuerpoCorreo = _cuerpoCorreo &
                               "<tr><td><b>DESTINO</b></td></tr>" &
                               "<tr><td>Proceso correcto</td></tr>" &
                               "<tr><td>Archivos</td><td>" & _archivosDestino.ToString() & "</td></tr>" &
                               "<tr><td>Directorios</td><td>" & _directoriosDestino.ToString() & "</td></tr>" &
                               "<tr><td>Peso</td><td>" & _pesoOrigen.ToString() & "</td></tr>"

                Else
                    ReglasLog.ReglasLog.CrearLog(Programacion.DescripcionProgramacion,
                                              Etiqueta & " Verificar Totales Destino | ERROR: " & _error.Origen & " - " & _error.Mensaje & " - " & _error.Detalle,
                                              ReglasLog.ReglasLog.TipoLog.Err, Programacion.VerbosidadLog, 1)

                    _estadoCorreo = "CON ERRORES"
                    _cuerpoCorreo = _cuerpoCorreo &
                            "<tr><td><b>DESTINORIGEN</b></td></tr>" &
                            "<tr><td>Error en el proceso</td></tr>" &
                            "<tr><td>Origen</td><td>" & _resultado.ErrorInfo.Origen & "</td></tr>" &
                            "<tr><td>Detalle</td><td>" & _resultado.ErrorInfo.Mensaje & "</td></tr>" &
                            "<tr><td>Stack</td><td>" & _resultado.ErrorInfo.Detalle & "</td></tr>"
                End If

                ' si hay diferencias intentará copiar nuevamente
                If _archivosOrigen > _archivosDestino And _intentosCopia < 2 Then
                    _resultado = New ResponseObject
                    _intentosCopia += 1
                    GoTo Copiar
                End If

                ' verifica si hay archivos en el origen que no estén en el destino
                If _archivosOrigen <> _archivosDestino Or _directoriosOrigen <> _directoriosDestino Then
                    Dim _archivosFaltantes As Integer

                    _cuerpoCorreo = _cuerpoCorreo &
                                    "<tr><td><br></td></tr>" &
                                    "<tr><td><b>Encontrar Faltantes</b></td></tr>"

                    _resultado = ReglasCopia.ReglasCopia.ProcesarArchivos(Programacion.DescripcionProgramacion, Etiqueta, ReglasCopia.ReglasCopia.TipoAccionArchivos.BuscarFaltantes,
                                                                           Programacion.RutaOrigen, Programacion.RutaDestino,
                                                                      Programacion.VerbosidadLog, _archivosFaltantes,
                                                                      Programacion.ArchivosExcluidos, Programacion.DirectoriosExcluidos)
                    If _resultado.IsError Then
                        ReglasLog.ReglasLog.CrearLog(Programacion.DescripcionProgramacion,
                                                  Etiqueta & " EntontrarFaltantes | Archivos faltantes: " & _archivosFaltantes.ToString() & " | " &
                                                 _resultado.ErrorInfo.Origen & ". " & _resultado.ErrorInfo.Mensaje & ". " &
                                                 _resultado.ErrorInfo.Detalle, ReglasLog.ReglasLog.TipoLog.Err, Programacion.VerbosidadLog, 1)

                        _estadoCorreo = "CON ERRORES"
                        _cuerpoCorreo = _cuerpoCorreo &
                                "<tr><td>Error en el proceso</td></tr>" &
                                "<tr><td>Origen</td><td>" & _resultado.ErrorInfo.Origen & "</td></tr>" &
                                "<tr><td>Detalle</td><td>" & _resultado.ErrorInfo.Mensaje & "</td></tr>" &
                                "<tr><td>Stack</td><td>" & _resultado.ErrorInfo.Detalle & "</td></tr>"

                    Else
                        ReglasLog.ReglasLog.CrearLog(Programacion.DescripcionProgramacion,
                                                  Etiqueta & " EntontrarFaltantes | Archivos faltantes: " & _archivosFaltantes.ToString(),
                                                 IIf(_archivosFaltantes > 0, ReglasLog.ReglasLog.TipoLog.Err, ReglasLog.ReglasLog.TipoLog.Info),
                                                 Programacion.VerbosidadLog, IIf(_archivosFaltantes > 0, 1, 2))

                        If _archivosFaltantes > 0 Then _estadoCorreo = "CON ERRORES"
                        _cuerpoCorreo = _cuerpoCorreo &
                               "<tr><td>Archivos faltantes</td><td>" & _archivosFaltantes.ToString() & "</td></tr>"
                    End If
                End If

                Dim _contactoError As String = ""
                If _estadoCorreo.ToUpper = "CON ERRORES" Then _contactoError = Programacion.CorreoError

                _cuerpoCorreo = "<table><tr><td>Proceso " & Programacion.DescripcionProgramacion & "</td></tr>" &
                                _cuerpoCorreo &
                                "</table>"

                EnviarCorreo(Programacion.CorreoConfirmacion, _contactoError, "AgenteAFEXTI - " & _estadoCorreo & " - " & Dns.GetHostName() & " - " &
                         Programacion.DescripcionProgramacion, _cuerpoCorreo)

            Catch ex As Exception
                ReglasLog.ReglasLog.CrearLog(_proceso, Etiqueta & " CopiarArchivos | " & ex.Source & ". " & ex.Message & ". " & ex.StackTrace,
                                              ReglasLog.ReglasLog.TipoLog.Err, Programacion.VerbosidadLog, 1)

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

        ''' <summary>
        ''' Verifica si la ruta origen existe
        ''' </summary>
        ''' <param name="Programacion"></param>
        ''' <returns></returns>
        Public Function VerificarRutaOrigen(Programacion As ReglasProgramacion.Programacion, Etiqueta As String) As ResponseObject
            Dim _resultado As New ResponseObject
            Dim _mensajeError As String
            Dim _proceso As String = Programacion.DescripcionProgramacion

            Try
                _mensajeError = "Validar"
                ReglasLog.ReglasLog.CrearLog(_proceso, Etiqueta & " | " & "Verificar Ruta Origen | Inicio ", ReglasLog.ReglasLog.TipoLog.Info, Programacion.VerbosidadLog, 1)

                _resultado = ReglasCopia.ReglasCopia.ExisteRuta(_proceso, Etiqueta, Programacion.RutaOrigen, Programacion.VerbosidadLog)
                If _resultado.IsError Then
                    ReglasLog.ReglasLog.CrearLog(_proceso, Etiqueta & " | " & "Verificar Ruta Origen | Fin | ERROR " & _resultado.ErrorInfo.Origen & ". " &
                                                     _resultado.ErrorInfo.Mensaje & ". " &
                                                _resultado.ErrorInfo.Detalle, ReglasLog.ReglasLog.TipoLog.Err, Programacion.VerbosidadLog, 1)
                Else
                    ReglasLog.ReglasLog.CrearLog(_proceso, Etiqueta & " | " & "Verificar Ruta Origen | La ruta existe", ReglasLog.ReglasLog.TipoLog.Info, Programacion.VerbosidadLog, 1)

                    ReglasLog.ReglasLog.CrearLog(_proceso, Etiqueta & " | " & "Verificar Ruta Origen | Fin ", ReglasLog.ReglasLog.TipoLog.Info, Programacion.VerbosidadLog, 1)
                End If

            Catch ex As Exception
                ReglasLog.ReglasLog.CrearLog(_proceso, Etiqueta & " | " & "Verificar Ruta Origen | " & ex.Source & ". " & ex.Message & ". " & ex.StackTrace, 3, Programacion.VerbosidadLog, 1)


            End Try
Fin:
            Return _resultado
        End Function

        ''' <summary>
        ''' Elimina el archivo BAT de instalación o actualización del mismo agente que queda en el menú de inicio 
        ''' cada vez que se instala o actualiza
        ''' </summary>
        ''' <returns></returns>
        Public Function EliminarInstaladorInicio(ByVal Programacion As Programacion) As ResponseObject
            Dim _resultado As New ResponseObject
            Try
                ReglasLog.ReglasLog.CrearLog(Programacion.DescripcionProgramacion, "Eliminando | Inicio ", ReglasLog.ReglasLog.TipoLog.Info,
                    Programacion.VerbosidadLog, 2)

                'elimina el archivo
                File.Delete(Programacion.RutaOrigen)

            Catch ex As Exception
                ReglasLog.ReglasLog.CrearLog(Programacion.DescripcionProgramacion, "Eliminando | ERROR: " & ex.Source & ". " & ex.Message & ". " & ex.StackTrace,
                                              ReglasLog.ReglasLog.TipoLog.Err, Programacion.VerbosidadLog, 1)


            Finally
                ReglasLog.ReglasLog.CrearLog(Programacion.DescripcionProgramacion, "Eliminando | Fin ", ReglasLog.ReglasLog.TipoLog.Info,
                    Programacion.VerbosidadLog, 1)
            End Try

            Return _resultado
        End Function

        ''' <summary>
        ''' Devuelve el valor de una key de la configuración secundaria del agente
        ''' </summary>
        ''' <param name="TipoConfiguracion"></param>
        ''' <returns></returns>
        Public Shared Function ValorConfiguracionServicio(ByVal TipoConfiguracion As TipoConfiguracionServicio) As String
            Dim _resultado As String
            Dim _xml As New XmlDocument

            Try
                Dim _archivoconfig As String = AppDomain.CurrentDomain.BaseDirectory & "configuracionagente.xml"
                _xml.Load(_archivoconfig)

                Dim _node As XmlNode = _xml.GetElementsByTagName(TipoConfiguracion.ToString())(0)
                _resultado = _node.Attributes.GetNamedItem("value").Value

            Catch ex As Exception
                Throw ex
            Finally
            End Try

            Return _resultado
        End Function

#Region "Enumeraciones"
        Public Enum TipoConfiguracionServicio
            verbosidadlogservicio = 1
            consultarbbdd = 2
        End Enum

#End Region

    End Class

End Namespace
