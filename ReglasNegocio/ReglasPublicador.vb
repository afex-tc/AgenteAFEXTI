Imports System.IO
Imports System.Xml
Imports System.Configuration
Imports AgenteAFEXTI.ReglasLog.ReglasLog
Imports AgenteAFEXTI.ReglasCorreo.ReglasCorreo
Imports LibreriaClases.General

Namespace ReglasPublicador

    Public Class ReglasPublicador

        ''' <summary>
        ''' Valida los archivos que se publicarán, y considera solo los que sean de una fecha menor entre el destino y el origen
        ''' </summary>
        Public Sub ValidarPublicacion()
            Dim _asunto As String = "Publicaciones automatizadas: validación publicación de hoy " & Now.ToShortDateString
            Dim _cuerpoCorreo As String = ""
            Dim _estadoPublicacion As String = ""

            Dim _archivosExcluidos As String = ConfigurationManager.AppSettings("ArchivosExcluidos").ToString
            Dim _directoriosExcluidos As String = ConfigurationManager.AppSettings("DirectoriosExcluidos").ToString

            Try
                Dim _xmlLog As New XmlDocument
                _xmlLog.Load(ConfigurationManager.AppSettings("LogPublicacion").ToString)
                Dim _nodoLogPublicacion As XmlNode = _xmlLog.FirstChild

                Dim _xml As New XmlDocument
                _xml.Load(ConfigurationManager.AppSettings("XMLPublicacion").ToString)

                For Each _nodo As XmlNode In _xml.ChildNodes
                    For Each _sitio As XmlNode In _nodo.ChildNodes
                        Dim _nombreSitio As String = _sitio.FirstChild.Value
                        Dim _origen As String = ""
                        Dim _destino As String = ""
                        Dim _respaldo As String = ""
                        Dim _etiqueta As String = ""
                        Dim _ambiente As String = ""

                        For Each _hijoSitio As XmlNode In _sitio.ChildNodes
                            Select Case _hijoSitio.Name.ToUpper
                                Case "ORIGEN"
                                    _origen = _hijoSitio.InnerText
                                Case "DESTINO"
                                    _destino = _hijoSitio.InnerText
                                Case "ETIQUETA"
                                    _etiqueta = _hijoSitio.InnerText
                                Case "AMBIENTE"
                                    _ambiente = _hijoSitio.InnerText
                            End Select
                        Next

                        Dim _resultado As ResponseObject
                        Dim _estado As String = ""

                        _origen = _origen & "\" & _nombreSitio
                        _destino = _destino & "\" & _nombreSitio

                        ' valida
                        InsertarLog("Validar Publicación - Inicio - ORIGEN: " & _origen & " | DESTINO : " & _destino, IESB_AFEX_ServicioLogger.LogDatadogStatus.Info, DevolverSourceDD(_nombreSitio))
                        _resultado = ReglasCopia.ReglasCopia.ProcesarArchivos("Validar Publicacion", ReglasCopia.ReglasCopia.TipoAccionArchivos.ValidarpreCopia,
                                                                              _origen, _destino, 3, _archivosExcluidos, _directoriosExcluidos)
                        If Not _resultado.IsError Then
                            _estado = "VALIDADO"
                            If _resultado.Objeto = 0 Then
                                InsertarLog("Validar Publicación - Fin - ORIGEN: " & _origen &
                                        " | DESTINO : " & _destino & " | Archivos a publicar: No se encontraron archivos para publicar", IESB_AFEX_ServicioLogger.LogDatadogStatus.Info, DevolverSourceDD(_nombreSitio))
                            Else
                                InsertarLog("Validar Publicación - Fin - ORIGEN: " & _origen &
                                        " | DESTINO : " & _destino & " | Archivos a publicar: " & _resultado.Objeto.ToString(), IESB_AFEX_ServicioLogger.LogDatadogStatus.Info, DevolverSourceDD(_nombreSitio))

                            End If

                        Else
                            _estado = "ERROR al validar Publicación: " & _resultado.ErrorInfo.Mensaje
                            InsertarLog("Validar Publicación - Fin - ORIGEN: " & _origen & " | DESTINO : " & _destino & " | " & _resultado.ErrorInfo.Mensaje, IESB_AFEX_ServicioLogger.LogDatadogStatus.Error, DevolverSourceDD(_nombreSitio))

                        End If

                        _cuerpoCorreo += "<tr><td>" & _ambiente & "</td><td>" & _nombreSitio & "<td>" & _origen & "</td><td>" & _destino & "</td><td>" & _estado & "</td></tr>"
                        _cuerpoCorreo += "<tr><td colspan=""5"">" & _resultado.Objeto.ToString().Replace(";", "<br />") & "</td></tr>"

                    Next
                Next

                If _cuerpoCorreo.ToUpper.IndexOf("ERROR") > 0 Then
                    _estadoPublicacion = " con ERRORES"
                Else
                    _estadoPublicacion = " COMPLETADA"
                End If

                _cuerpoCorreo = "<table><tr style=""background-color: #c5d3dd; font-family: Arial Black;""><td>AMBIENTE</td><td>SITIO<td>ORIGEN</td><td>DESTINO</td><td>ESTADO</td></tr>" &
                                 _cuerpoCorreo & "</table>"

                _xmlLog.Save(ConfigurationManager.AppSettings("LogPublicacion"))
            Catch ex As Exception
                _estadoPublicacion = " con ERRORES"
                _cuerpoCorreo = "Error al validar la publicación de hoy. " & ex.Source & ". " & ex.Message
            End Try

            _asunto += " " & _estadoPublicacion
            EnviarCorreo(ConfigurationManager.AppSettings("correosnotificacion").ToString, "", _asunto, _cuerpoCorreo)
        End Sub

        Public Sub BuscarPublicacionCertificacion(ByVal Publicacion As Publicacion)
            Dim _asunto As String = "Publicaciones automatizadas: verificar pendientes CERTIFICACION " & Now.ToShortDateString
            Dim _cuerpoCorreo As String = ""
            Dim _estadoPublicacion As String = "sin ERRORES"
            Dim _contactosCorreo As String = Publicacion.Programacion.CorreoConfirmacion
            Dim _servicio As String = "Publicación CERTIFICACIÓN " & Publicacion.Sistema

            Try
                ReglasLog.ReglasLog.CrearLog(_servicio, "Verificar pendientes | Inicio",
                                             1, Publicacion.Programacion.VerbosidadLog, 1)

                ' verifica si existe una nueva publicación para certificación
                Dim _resultado As ResponseObject = ReglasCopia.ReglasCopia.ProcesarArchivos(_servicio, ReglasCopia.ReglasCopia.TipoAccionArchivos.ValidarpreCopia,
                                                                                            Publicacion.Programacion.RutaOrigen,
                                                                                            Publicacion.Programacion.RutaDestino,
                                                                                            Publicacion.Programacion.VerbosidadLog,
                                                                                            Publicacion.Programacion.ArchivosExcluidos,
                                                                                            Publicacion.Programacion.DirectoriosExcluidos)

                If Not _resultado.IsError Then
                    If IsNumeric(_resultado) = 1 Then
                        ' existen archivos pendientes de publicar
                        _cuerpoCorreo = "Existen publicaciones pendientes para CERTIFICACIÓN. <br /><br />" &
                                        "Debe autorizar la publicación."

                        ReglasLog.ReglasLog.CrearLog(_servicio, "Verificar pendientes | Existen versiones pendientes para publicar", 1, Publicacion.Programacion.VerbosidadLog,
                                            1)

                    End If
                Else
                    ReglasLog.ReglasLog.CrearLog(_servicio, "Verificar pendientes | " & _resultado.ErrorInfo.Mensaje, 3, Publicacion.Programacion.VerbosidadLog, 1)
                End If

                ReglasLog.ReglasLog.CrearLog(_servicio, "Verificar pendientes | Fin", 1, Publicacion.Programacion.VerbosidadLog, 1)

            Catch ex As Exception
                _estadoPublicacion = " con ERRORES"
                _cuerpoCorreo = "Error al verificar publicaciones pendientes para CERTIFICACIÓN. " & ex.Source & ". " & ex.Message
                _contactosCorreo = Publicacion.Programacion.CorreoError

                ReglasLog.ReglasLog.CrearLog(_servicio, "Verificar pendientes | " & ex.Source & ". " & ex.Message, 3, Publicacion.Programacion.VerbosidadLog, 1)
            End Try

            _asunto += " " & _estadoPublicacion
            EnviarCorreo(_contactosCorreo, "", _asunto, _cuerpoCorreo)
        End Sub

        ''' <summary>
        ''' Realiza la publicación configurada en el XML de publicación
        ''' </summary>
        Public Sub RealizarPublicacion()
            Dim _asunto As String = "Publicaciones automatizadas: publicación de hoy " & Now.ToShortDateString
            Dim _cuerpoCorreo As String = ""
            Dim _estadoPublicacion As String = ""

            Dim _archivosExcluidos As String = ConfigurationManager.AppSettings("ArchivosExcluidos").ToString
            Dim _directoriosExcluidos As String = ConfigurationManager.AppSettings("DirectoriosExcluidos").ToString

            Try
                Dim _xmlLog As New XmlDocument
                _xmlLog.Load(ConfigurationManager.AppSettings("LogPublicacion").ToString)
                Dim _nodoLogPublicacion As XmlNode = _xmlLog.FirstChild

                Dim _xml As New XmlDocument
                _xml.Load(ConfigurationManager.AppSettings("XMLPublicacion").ToString)

                Dim _ambienteAnterior As String = ""

                For Each _nodo As XmlNode In _xml.ChildNodes
                    Dim _cantidadRespaldo As Integer = 0
                    For Each _sitio As XmlNode In _nodo.ChildNodes
                        _cantidadRespaldo += 1
                        Dim _nombreSitio As String = _sitio.FirstChild.Value
                        Dim _origen As String = ""
                        Dim _destino As String = ""
                        Dim _respaldo As String = ""
                        Dim _etiqueta As String = ""
                        Dim _ambiente As String = ""

                        For Each _hijoSitio As XmlNode In _sitio.ChildNodes
                            Select Case _hijoSitio.Name.ToUpper
                                Case "ORIGEN"
                                    _origen = _hijoSitio.InnerText
                                Case "DESTINO"
                                    _destino = _hijoSitio.InnerText
                                Case "RESPALDO"
                                    _respaldo = _hijoSitio.InnerText
                                Case "ETIQUETA"
                                    _etiqueta = _hijoSitio.InnerText
                                Case "AMBIENTE"
                                    _ambiente = _hijoSitio.InnerText
                            End Select
                        Next

                        Dim _resultado As ResponseObject
                        Dim _estado As String = ""

                        _origen = _origen & "\" & _nombreSitio
                        _destino = _destino & "\" & _nombreSitio
                        _respaldo = _respaldo & "\" & _ambiente

                        Dim _descripcionLog As String = _ambiente & " - " & _respaldo

                        ' eliminar respaldos anteriores, ingresa solo 1 vez por ambiente
                        If _ambienteAnterior <> _ambiente Then
                            ' elimina solo 1 vez por ambiente
                            InsertarLog("Eliminar Respaldos - Inicio - " & _descripcionLog, IESB_AFEX_ServicioLogger.LogDatadogStatus.Info, DevolverSourceDD(_nombreSitio))
                            _resultado = New ResponseObject(EliminarRespaldosAnteriores(_respaldo))
                            If _resultado.IsError Then
                                _cuerpoCorreo += "<tr><td colspan=""5"">ERROR al eliminar respaldos: " & _resultado.ErrorInfo.Detalle & "</td></tr>"
                                InsertarLog("Eliminar Respaldos - Fin - " & _descripcionLog & " | " & _resultado.ErrorInfo.Detalle, IESB_AFEX_ServicioLogger.LogDatadogStatus.Error, DevolverSourceDD(_nombreSitio))
                            Else
                                _cuerpoCorreo += "<tr><td colspan=""5"">Respaldos anteriores eliminados.</td></tr>"
                                InsertarLog("Eliminar Respaldos - Fin - " & _descripcionLog, IESB_AFEX_ServicioLogger.LogDatadogStatus.Info, DevolverSourceDD(_nombreSitio))
                            End If
                            _ambienteAnterior = _ambiente
                        End If

                        ' respaldo, primero crea el directorio de fecha
                        _respaldo = _respaldo & "\" & Now.ToShortDateString.Replace("-", "")
                        _descripcionLog = _nombreSitio & " | ORIGEN: " & _destino & " | DESTINO: " & _respaldo & "\" & _nombreSitio & "_" & _cantidadRespaldo.ToString()
                        InsertarLog("Publicación | Crear Respaldos - Inicio - " & _descripcionLog, IESB_AFEX_ServicioLogger.LogDatadogStatus.Info, DevolverSourceDD(_nombreSitio))
                        _resultado = ReglasCopia.ReglasCopia.ProcesarArchivos("Publicador", True, "Publicación | Crear Respaldos", "", _respaldo, 3, 0)
                        If Not _resultado.IsError Then
                            _respaldo = _respaldo & "\" & _nombreSitio & "_" & _cantidadRespaldo.ToString()

                            _resultado = ReglasCopia.ReglasCopia.ProcesarArchivos("Publicador", True, "Publicación | Copiar Respaldos", _destino, _respaldo, 3, 0)
                            If Not _resultado.IsError Then
                                InsertarLog("Crear Respaldos - Fin - " & _descripcionLog, IESB_AFEX_ServicioLogger.LogDatadogStatus.Info, DevolverSourceDD(_nombreSitio))

                                ' publica
                                _descripcionLog = "ORIGEN: " & _origen & " | DESTINO : " & _destino
                                InsertarLog("Publicación | Publicar - Inicio - " & _descripcionLog, IESB_AFEX_ServicioLogger.LogDatadogStatus.Info, DevolverSourceDD(_nombreSitio))
                                Dim _archivosActualizados As Integer
                                _resultado = ReglasCopia.ReglasCopia.ProcesarArchivos("Publicador", True, "Publicación | Publicar ", _origen, _destino, 3, _archivosExcluidos, _directoriosExcluidos)
                                If Not _resultado.IsError Then
                                    _estado = "COPIADO"

                                    InsertarLog("Publicar - Fin - " & _descripcionLog & " | Archivos actualizados: " & _archivosActualizados.ToString, IESB_AFEX_ServicioLogger.LogDatadogStatus.Info, DevolverSourceDD(_nombreSitio))

                                Else
                                    _estado = "ERROR al publicar: " & _resultado.ErrorInfo.Mensaje
                                    InsertarLog("Publicar - Fin - " & _descripcionLog & " | " & _resultado.ErrorInfo.Mensaje, IESB_AFEX_ServicioLogger.LogDatadogStatus.Error, DevolverSourceDD(_nombreSitio))

                                End If
                            Else
                                _estado = "ERROR al respaldar: " & _resultado.ErrorInfo.Mensaje
                                InsertarLog("Crear Respaldos - Fin - " & _descripcionLog & " | " & _resultado.ErrorInfo.Mensaje, IESB_AFEX_ServicioLogger.LogDatadogStatus.Error, DevolverSourceDD(_nombreSitio))
                            End If
                        Else
                            _estado = "ERROR al respaldar: " & _resultado.ErrorInfo.Mensaje
                            InsertarLog("Crear Respaldos - Fin - " & _descripcionLog & " | " & _resultado.ErrorInfo.Mensaje, IESB_AFEX_ServicioLogger.LogDatadogStatus.Error, DevolverSourceDD(_nombreSitio))
                        End If
                        _cuerpoCorreo += "<tr><td>" & _ambiente & "</td><td>" & _nombreSitio & "<td>" & _origen & "</td><td>" & _destino & "</td><td>" & _respaldo & "</td><td>" & _estado & "</td></tr>"
                        _cuerpoCorreo += "<tr><td colspan=""6"">" & _resultado.Objeto.ToString().Replace(";", "<br />") & "</td></tr>"

                        Try

                            Dim _elemento As XmlElement = _xmlLog.CreateElement("log")
                            Dim _atributo As XmlAttribute = _xmlLog.CreateAttribute("fecha")
                            _atributo.Value = Now
                            _elemento.Attributes.Append(_atributo)

                            _atributo = _xmlLog.CreateAttribute("ambiente")
                            _atributo.Value = _ambiente
                            _elemento.Attributes.Append(_atributo)

                            _atributo = _xmlLog.CreateAttribute("sitio")
                            _atributo.Value = _nombreSitio
                            _elemento.Attributes.Append(_atributo)

                            _atributo = _xmlLog.CreateAttribute("origen")
                            _atributo.Value = _origen
                            _elemento.Attributes.Append(_atributo)

                            _atributo = _xmlLog.CreateAttribute("destino")
                            _atributo.Value = _destino
                            _elemento.Attributes.Append(_atributo)

                            _atributo = _xmlLog.CreateAttribute("respaldo")
                            _atributo.Value = _respaldo
                            _elemento.Attributes.Append(_atributo)

                            _atributo = _xmlLog.CreateAttribute("etiqueta")
                            _atributo.Value = _etiqueta
                            _elemento.Attributes.Append(_atributo)

                            _atributo = _xmlLog.CreateAttribute("estado")
                            _atributo.Value = _estado
                            _elemento.Attributes.Append(_atributo)

                            _nodoLogPublicacion.AppendChild(_elemento)

                        Catch ex As Exception
                            _cuerpoCorreo += "<tr><td colspan=""5"">ERROR al grabar log. " & ex.Source & ". " & ex.Message & "</td></tr>"
                        End Try

                    Next
                Next

                If _cuerpoCorreo.ToUpper.IndexOf("ERROR") > 0 Then
                    _estadoPublicacion = " con ERRORES"
                Else
                    _estadoPublicacion = " COMPLETADA"
                End If

                _cuerpoCorreo = "<table><tr style=""background-color: #c5d3dd; font-family: Arial Black;""><td>AMBIENTE</td><td>SITIO<td>ORIGEN</td><td>DESTINO</td><td>RESPALDO</td><td>ESTADO</td></tr>" &
                                 _cuerpoCorreo & "</table>"

                _xmlLog.Save(ConfigurationManager.AppSettings("LogPublicacion"))
            Catch ex As Exception
                _estadoPublicacion = " con ERRORES"
                _cuerpoCorreo = "Error al realizar la publicación de hoy. " & ex.Source & ". " & ex.Message
            End Try

            _asunto += " " & _estadoPublicacion
            EnviarCorreo(ConfigurationManager.AppSettings("correosnotificacion").ToString, "", _asunto, _cuerpoCorreo)
        End Sub

        ''' <summary>
        ''' Elimina los directorios de una ruta especifica
        ''' </summary>
        ''' <param name="Ruta"></param>
        ''' <returns></returns>
        Public Function EliminarRespaldosAnteriores(Ruta As String) As String
            Dim _resultado As String = ""
            Dim _mensaje As String = ""
            Try

                For Each _directorio As String In Directory.GetDirectories(Ruta)
                    Dim _fechaCreacion As Date = Directory.GetCreationTime(_directorio)
                    If _fechaCreacion.Date < Now.Date.AddDays(-2) Then
                        For Each _directorio2 As String In Directory.GetDirectories(_directorio)
                            _mensaje = _directorio2
                            Directory.Delete(_directorio2, True)
                        Next
                        Directory.Delete(_directorio, True)
                    End If
                Next
            Catch ex As Exception
                _resultado = "ERROR: " & ex.Source & ". " & ex.Message & ". " & _mensaje
            End Try

            Return _resultado
        End Function

    End Class

    Public Class Publicacion
        Private _programacion As ReglasProgramacion.Programacion
        Private _ambiente As AmbientePublicacion
        Private _sistema As String

        Public Property Programacion As ReglasProgramacion.Programacion
            Set(value As ReglasProgramacion.Programacion)
                Me._programacion = value
            End Set
            Get
                Return Me._programacion
            End Get
        End Property

        Public Property Ambiente As AmbientePublicacion
            Set(value As AmbientePublicacion)
                Me._ambiente = value
            End Set
            Get
                Return Me._ambiente
            End Get
        End Property

        Public Property Sistema()
            Set(value)
                Me._sistema = value
            End Set
            Get
                Return Me._sistema
            End Get
        End Property
    End Class

    Public Enum AmbientePublicacion
        PreCertificacion = 1
        Certificacion = 2
        Produccion = 3
    End Enum

End Namespace
