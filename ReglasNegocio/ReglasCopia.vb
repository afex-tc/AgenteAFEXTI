Imports System.IO
Imports System.Runtime.InteropServices
Imports AgenteAFEXTI.ReglasLog.ReglasLog
Imports LibreriaClases.General
Imports LibreriaClases.Errores


Namespace ReglasCopia

    Public Class ReglasCopia

        ''' <summary>
        ''' Copia o Valida los arhivos entre una ruta de origen y otra de destino, solo valida o copia lo diferente
        ''' </summary>
        ''' <param name="Servicio"></param>
        ''' <param name="Accion"></param>
        ''' <param name="DirectorioOrigen"></param>
        ''' <param name="DirectorioDestino"></param>
        ''' <param name="VerbosidadLog"></param>
        ''' <param name="CantidadArchivos"></param>
        ''' <param name="ArchivosExcluidos"></param>
        ''' <param name="DirectoriosExcluidos"></param>
        ''' <param name="RutaDestinoRaiz"></param>
        ''' <param name="ArchivosLargos"></param>
        ''' <param name="DirectoriosLargos"></param>
        ''' <param name="CantidadArchivosExcluidos"></param>
        ''' <param name="CantidadDirectoriosExcluidos"></param>
        ''' <returns></returns>
        Public Shared Function ProcesarArchivos(Servicio As String, Etiqueta As String, Accion As TipoAccionArchivos, ByVal DirectorioOrigen As String,
                                                ByVal DirectorioDestino As String, ByVal VerbosidadLog As Integer,
                                                ByRef CantidadArchivos As Integer,
                                                Optional ArchivosExcluidos As String = "",
                                                Optional DirectoriosExcluidos As String = "",
                                                Optional ByVal RutaDestinoRaiz As String = "",
                                                Optional ByRef ArchivosLargos As Integer = 0,
                                                Optional ByRef DirectoriosLargos As Integer = 0,
                                                Optional ByRef CantidadArchivosExcluidos As Integer = 0,
                                                Optional ByRef CantidadDirectoriosExcluidos As Integer = 0,
                                                Optional ByVal EliminarOrigen As Boolean = False,
                                                Optional ByVal RutaRetencionOrigen As String = "") As ResponseObject

            Dim _resultado As New ResponseObject
            Dim _linea As Integer = 1

            Dim _directorioLargo As Boolean = False
            Dim _nombreDirectorioLargo As String = ""

            If ArchivosExcluidos Is Nothing Then ArchivosExcluidos = ""
            If DirectoriosExcluidos Is Nothing Then DirectoriosExcluidos = ""

            If RutaDestinoRaiz = "" Then RutaDestinoRaiz = DirectorioDestino

            Dim _rutaDirectoriosLargos = Path.Combine(RutaDestinoRaiz, "@AL@")

            Dim _encabezadoLog As String = Etiqueta & " " & Accion.ToString()

            ' verifica si existe el origen, solo procesará si existe
            Dim _rutaValida As ResponseObject = ExisteRuta(Servicio, _encabezadoLog, DirectorioOrigen, VerbosidadLog)
            If Not _rutaValida.IsError And _rutaValida.Objeto = True Then
                Try

                    _directorioLargo = ProcesaRutaDestino(TipoElementoRuta.Directorio, DirectorioDestino, RutaDestinoRaiz, _rutaDirectoriosLargos, DirectorioDestino,
                                                _nombreDirectorioLargo, Servicio, _encabezadoLog, VerbosidadLog)
                    If _directorioLargo Then DirectoriosLargos += 1

                    ' si no existe el directorio en el destino, lo crea
                    If Accion = TipoAccionArchivos.Copiar And Not Directory.Exists(DirectorioDestino) And Not ValidarExcluido(Servicio, _encabezadoLog, DirectoriosExcluidos, DirectorioOrigen, VerbosidadLog) Then

                        My.Computer.FileSystem.CreateDirectory(DirectorioDestino)
                        ReglasLog.ReglasLog.CrearLog(Servicio, _encabezadoLog & " | Directorio creado : " & DirectorioDestino, 1, VerbosidadLog, 2)

                        If _directorioLargo Then

                            '' escribe en un txt la ruta del directorio largo para que se conozca su ruta original
                            LogRutaLarga(2, DirectorioDestino, DirectorioOrigen, _rutaDirectoriosLargos, Servicio, _encabezadoLog, VerbosidadLog)

                        End If
                    End If

                    ' recorre los archivos del directorioorigen
                    If Not ValidarExcluido(Servicio, _encabezadoLog, DirectoriosExcluidos, DirectorioOrigen, VerbosidadLog) Then

                        For Each _archivo As String In Directory.GetFiles(DirectorioOrigen)
                            Dim _archivoLargo As Boolean = False

                            Dim _nombreArchivo As String = Path.GetFileName(_archivo)

                            ReglasLog.ReglasLog.CrearLog(Servicio, _encabezadoLog & " | Archivo listado : " & _archivo & " | directorio | " &
                                                     Directory.GetParent(_archivo).ToString(), TipoLog.Info, VerbosidadLog, 2)

                            If Not ValidarExcluido(Servicio, _encabezadoLog, ArchivosExcluidos, _nombreArchivo, VerbosidadLog) Then

                                Dim _archivoDestino As String

                                _archivoLargo = ProcesaRutaDestino(TipoElementoRuta.Archivo, DirectorioDestino, RutaDestinoRaiz, _rutaDirectoriosLargos,
                                                _archivoDestino, _nombreArchivo, Servicio, _encabezadoLog, VerbosidadLog)

                                If _archivoLargo Then ArchivosLargos += 1

                                Dim _fechaModificacionOrigen As Date = File.GetLastWriteTime(_archivo)
                                Dim _fechaModificacionDestino As Date = File.GetLastWriteTime(_archivoDestino)

                                ReglasLog.ReglasLog.CrearLog(Servicio, _encabezadoLog & " | Archivo listado | fecha origen : " &
                                                         _archivo & " " & _fechaModificacionOrigen.ToString & " | fecha destino : " & _archivoDestino & " " &
                                                         _fechaModificacionDestino.ToString, TipoLog.Info, VerbosidadLog, 3)
                                Dim _errorCopiado As Boolean = False

                                ' valida o copia por fecha
                                If _fechaModificacionOrigen > _fechaModificacionDestino Or
                                LargoMaximoRuta(TipoElementoRuta.Archivo) <= _archivo.ToString().Length Then
                                    ReglasLog.ReglasLog.CrearLog(Servicio, _encabezadoLog & " | Archivo a copiar : " & _archivo, 1, VerbosidadLog, 2)

                                    Dim _descripcionAccion As String = ""

                                    Try

                                        Select Case Accion
                                            Case TipoAccionArchivos.Copiar ' copiar
                                                _descripcionAccion = "copiado"

                                                If _archivoLargo Then
                                                    '' escribe en un txt la ruta del directorio largo para que se conozca su ruta original
                                                    If Not Directory.Exists(_rutaDirectoriosLargos) Then Directory.CreateDirectory(_rutaDirectoriosLargos)

                                                    LogRutaLarga(1, _archivoDestino, _archivo, _rutaDirectoriosLargos, Servicio, _encabezadoLog, VerbosidadLog)
                                                End If

                                                File.Copy(_archivo, _archivoDestino, True)

                                                ' valida si tiene ruta para retener el archivo origen para copiar antes de eliminar
                                                If RutaRetencionOrigen <> "" Then
                                                    Dim _archivoRetencionOrigen As String = ""
                                                    _archivoRetencionOrigen = RutaRetencionOrigen & "\" & _nombreArchivo
                                                    File.Copy(_archivo, _archivoRetencionOrigen, True)
                                                End If


                                            Case TipoAccionArchivos.ValidarpreCopia ' valida diferencias
                                                _descripcionAccion = "validado"

                                        End Select

                                        ReglasLog.ReglasLog.CrearLog(Servicio, _encabezadoLog & " | Archivo " & _descripcionAccion & ": " & _archivoDestino, TipoLog.Info,
                                                                         VerbosidadLog, 2)

                                    Catch ex As Exception
                                        _errorCopiado = True
                                        ReglasLog.ReglasLog.CrearLog(Servicio, _encabezadoLog & " | Archivo NO " & _descripcionAccion & " | Origen : " & _archivo & " | Destino : " &
                                                                         _archivoDestino & " | ERROR : " & ex.Source & ". " & ex.Message & ". " & ex.StackTrace, TipoLog.Err,
                                                                          VerbosidadLog, 1)
                                    End Try
                                    If Accion <> TipoAccionArchivos.BuscarFaltantes Then CantidadArchivos += 1
                                End If

                                ' verifica si elimina
                                If Accion = TipoAccionArchivos.Copiar And EliminarOrigen And Not _errorCopiado Then
                                    ReglasLog.ReglasLog.CrearLog(Servicio, _encabezadoLog & " | Eliminar Archivo Origen : " & _archivo, TipoLog.Info,
                                                                         VerbosidadLog, 2)
                                    File.Delete(_archivo)

                                    ReglasLog.ReglasLog.CrearLog(Servicio, _encabezadoLog & " | Archivo Origen Eliminado : " & _archivo, TipoLog.Info,
                                                                         VerbosidadLog, 2)
                                End If

                                ' acciones independientes de su fecha de actualización
                                If Accion = TipoAccionArchivos.BuscarFaltantes Then
                                    If Not File.Exists(_archivoDestino) Then
                                        CantidadArchivos += 1
                                        ReglasLog.ReglasLog.CrearLog(Servicio, _encabezadoLog & " | EncontrarFaltantes | Archivo no existe : " &
                                                                 _archivoDestino & " | Origen : " & _archivo, TipoLog.Err,
                                                                 VerbosidadLog, 1)
                                    End If
                                End If
                            Else
                                CantidadArchivosExcluidos += 1
                            End If

Prueba:
                        Next
                    End If

                    ' valida los subdirectorios
                    For Each _directorio As String In Directory.GetDirectories(DirectorioOrigen)

                        If Not ValidarExcluido(Servicio, _encabezadoLog, DirectoriosExcluidos, _directorio, VerbosidadLog) Then
                            Dim _directorioDestino As String = Path.Combine(DirectorioDestino, Path.GetFileName(_directorio))
                            ReglasLog.ReglasLog.CrearLog(Servicio, _encabezadoLog & " | SubDirectorio : " & _directorio & " - Eliminar: " & EliminarOrigen.ToString(),
                                                         TipoLog.Info, VerbosidadLog, 3)
                            Dim _resultadoInterno As ResponseObject =
                            ReglasCopia.ProcesarArchivos(Servicio, Etiqueta, Accion, _directorio, _directorioDestino, VerbosidadLog, CantidadArchivos,
                                                         ArchivosExcluidos, DirectoriosExcluidos, RutaDestinoRaiz, ArchivosLargos,
                                                         DirectoriosLargos, 0, 0, EliminarOrigen, RutaRetencionOrigen)

                            If _resultadoInterno.IsError And _resultadoInterno.ErrorInfo.Codigo = -2 Then
                                Throw New Exception(_resultadoInterno.ErrorInfo.Origen & ". " &
                                                _resultadoInterno.ErrorInfo.Mensaje & ". " &
                                                _resultadoInterno.ErrorInfo.Detalle)
                            End If

                        Else
                            CantidadDirectoriosExcluidos += 1
                        End If

                        ' verifica si elimina
                        If EliminarOrigen Then
                            ReglasLog.ReglasLog.CrearLog(Servicio, _encabezadoLog & " | Eliminar Directorio Origen : " &
                                                         _directorio, TipoLog.Info, VerbosidadLog, 2)
                            Directory.Delete(_directorio, True)

                            ReglasLog.ReglasLog.CrearLog(Servicio, _encabezadoLog & " | Directorio Eliminado : " &
                                                         DirectorioOrigen, TipoLog.Info, VerbosidadLog, 2)
                        End If

                    Next

                Catch ex As Exception
                    Dim _codigoError As Integer = 0

                    ' errores que deben provocar el termino completo del proceso
                    If ex.Message.ToUpper.IndexOf("ESPACIO EN DISCO INSUFICIENTE") > -1 Or
                    ex.Message.ToUpper.IndexOf("Error relacionado con la red") > -1 Then
                        _codigoError = -2
                    End If

                    _resultado = New ResponseObject(New ErrorInfo("(" & _linea.ToString() & ") " & ex.Source & ". " &
                                                            ex.Message, ex.StackTrace, _codigoError, "ProcesarArchivos"))

                    ReglasLog.ReglasLog.CrearLog(Servicio, _encabezadoLog & " | " & _resultado.ErrorInfo.Origen &
                                             ". " & _resultado.ErrorInfo.Mensaje & ". " & _resultado.ErrorInfo.Detalle,
                                              TipoLog.Err, VerbosidadLog, 1)
                End Try

            Else
                Dim _tipoLog As Integer = TipoLog.Err
                If TipoAccionArchivos.ValidarpreCopia Then _tipoLog = TipoLog.Info

                ReglasLog.ReglasLog.CrearLog(Servicio, _encabezadoLog & "La ruta origen:" & DirectorioOrigen & " no existe.",
                                             _tipoLog, VerbosidadLog, 1)

                _resultado = New ResponseObject("Origen no encontrado.", "La ruta origen:" & DirectorioOrigen & " no existe.", 10000, "ReglasCopia.ProcesarArchivos")
            End If

            Return _resultado
        End Function

        ''' <summary>
        ''' Valida si un directorio o archivo se encuentra dentro del listado de los excluidos
        ''' </summary>
        ''' <param name="EncabezadoLog">descripción para el log</param>
        ''' <param name="ListaExcluidos">Lista de elementos exluidos, deben estar separados por ;</param>
        ''' <param name="Elemento">Elemento que se revisará denrro de los excluidos</param>
        ''' <returns>True en caso de que ELEMENTO esté excluido </returns>
        Public Shared Function ValidarExcluido(Servicio As String, EncabezadoLog As String, ListaExcluidos As String, Elemento As String, VerbosidadLog As Integer) As Boolean
            Dim _resultado As Boolean = False

            Dim _excluidos As String() = Split(ListaExcluidos, ";")
            Dim _excluido As Boolean = False
            For i As Integer = 0 To UBound(_excluidos) - 1
                ReglasLog.ReglasLog.CrearLog(Servicio, EncabezadoLog & " | ValidarExcluido | Elemento : " & Elemento & " | Excluidos " & ListaExcluidos &
                         " | excluido : " & _excluidos(i) & " (" & UBound(_excluidos).ToString & " - " & i.ToString & ")", 1, VerbosidadLog, 3)

                If Elemento.ToUpper.IndexOf(_excluidos(i).ToUpper) > -1 Then
                    _resultado = True
                    Exit For
                End If
            Next

            Return _resultado
        End Function

        Public Shared Function ValidarTotales(ByVal Servicio As String, ByVal Origen As String, ByRef PesoTotal As Integer, ByRef CantidadDirectorios As Integer,
                                              ByRef CantidadArchivos As Integer, ByVal VerbosidadLog As Integer,
                                              Optional ByVal ArchivosExcluidos As String = "",
                                              Optional ByVal DirectoriosExcluidos As String = "") As LibreriaClases.Errores.ErrorInfo

            Dim _resultado As New LibreriaClases.Errores.ErrorInfo()

            Try
                If Not ValidarExcluido(Servicio, "ValidarTotales", DirectoriosExcluidos, Origen, VerbosidadLog) Then
                    ' recorre los archivos de los directorios para sacar su tamaño
                    For Each _archivo In My.Computer.FileSystem.GetFiles(Origen)
                        If Not ValidarExcluido(Servicio, "ValidarTotales", ArchivosExcluidos, _archivo, VerbosidadLog) Then
                            CantidadArchivos += 1
                            PesoTotal += _archivo.Length
                        End If
                    Next

                    ' recorre los subdirectorios
                    For Each _directorio In My.Computer.FileSystem.GetDirectories(Origen)
                        If Not ValidarExcluido(Servicio, "ValidarTotales", DirectoriosExcluidos, _directorio, VerbosidadLog) Then
                            CantidadDirectorios += 1
                            _resultado = ValidarTotales(Servicio, _directorio, PesoTotal, CantidadDirectorios, CantidadArchivos, VerbosidadLog,
                                                        ArchivosExcluidos, DirectoriosExcluidos)
                        End If
                    Next
                End If

            Catch ex As Exception
                _resultado = New LibreriaClases.Errores.ErrorInfo(ex.Source & " - " & ex.Message, ex.StackTrace, 0, "ReglasCopia.ValidarTotales")
            End Try

            Return _resultado
        End Function

        ''' <summary>
        '''Toma una ruta de directorio o archivo y devuelve la nueva ruta validando si la ruta sobrepasa el máximo permitido
        '''devuelve también el nombre del directorio o de archivo en caso de exceder el máximo
        ''' </summary>
        ''' <param name="Tipo">1: archivo; 2: directorio</param>
        ''' <param name="Ruta"></param>
        ''' <param name="RutaRaiz"></param>
        ''' <param name="RutaElementosLargos"></param>
        ''' <param name="NuevaRuta"></param>
        ''' <param name="NombreElemento">si se valida la ruta de un archivo, aquí debe venir el nombre del archivo, de lo contrario vacio</param>
        ''' <param name="Servicio"></param>
        ''' <param name="EncabezadoLog"></param>
        ''' <param name="VerbosidadLog"></param>
        ''' <returns></returns>
        Public Shared Function ProcesaRutaDestino(ByVal Tipo As TipoElementoRuta, ByVal Ruta As String, ByVal RutaRaiz As String, ByVal RutaElementosLargos As String,
                                     ByRef NuevaRuta As String, ByRef NombreElemento As String, ByVal Servicio As String,
                                     ByVal EncabezadoLog As String, ByVal VerbosidadLog As Integer) As Boolean

            Dim _resultado As Boolean = False
            Dim _largoMaximo As Integer = LargoMaximoRuta(Tipo)

            ' si la ruta recibida sobrepasa el máximo se crea una nueva quitando caracteres menos importantes, además se el nombre por
            ' separado en caso de que la ruta exceda el máximo
            ' en caso de que la ruta supere los máximos, se quita la ruta raiz para crearla nueva ruta sin alterar la raiz

            CrearLog(Servicio, EncabezadoLog & " | " & Tipo.ToString() &
                         " a procesar | Ruta : " & Ruta & " | NuevaRuta : " & NuevaRuta & " | nombreelemento : " & NombreElemento, 1, VerbosidadLog, 3)

            If (Ruta.ToString() & "\" & NombreElemento.ToString()).Length >= _largoMaximo Then
                NuevaRuta = Ruta.ToString().Replace(RutaRaiz, "")
                NuevaRuta = NuevaRuta.ToString().Replace(" ", "").Replace(".", "").Replace(",", "").Replace("-", "").Replace("_", "")
                NuevaRuta = NuevaRuta.ToString().Replace("\", "")

                If Tipo = TipoElementoRuta.Archivo Then
                    NombreElemento = NombreElemento.ToString().Replace(" ", "").Replace(".", "").Replace(",", "").Replace("-", "").Replace("_", "")
                    NombreElemento.ToString().Replace("\", "")
                    NombreElemento = "__" & NombreElemento
                Else
                    NombreElemento = ""
                End If

                NombreElemento = NuevaRuta.ToString() & NombreElemento.ToString()

                ' vuelve a validar el largo del elemento para evitar que el LOG de ruta larga supere también el largo máximo
                Dim _largoArchivo As Integer = (RutaElementosLargos.ToString() & "\" & NombreElemento.ToString() & "_LOG_" &
                                            Tipo.ToString() & ".txt").ToString().Length

                If _largoArchivo > _largoMaximo Then
                    NombreElemento = Mid(NombreElemento.ToString(), (_largoArchivo + 1) - _largoMaximo)
                    ReglasLog.ReglasLog.CrearLog(Servicio, EncabezadoLog & " | " &
                                                 Tipo.ToString() & " extra largo | " &
                                                 NombreElemento & " | largo = " & _largoArchivo.ToString(), TipoLog.Info,
                                                 VerbosidadLog, 3)
                End If

                NuevaRuta = RutaElementosLargos.ToString() & "\" & NombreElemento

                CrearLog(Servicio, EncabezadoLog & " | " & Tipo.ToString() &
                         " largo | Ruta : " & Ruta & " | NuevaRuta : " & NuevaRuta, TipoLog.Info, VerbosidadLog, 3)

                _resultado = True
            Else

                NuevaRuta = Path.Combine(Ruta, IIf(Tipo = TipoElementoRuta.Archivo, NombreElemento, ""))
            End If

            Return _resultado
        End Function

        ''' <summary>
        ''' Escribe en un txt la ruta original del elemento largo, para que se conozca su ruta original
        ''' </summary>
        ''' <param name="Tipo">1: archivo; 2: directorio</param>
        ''' <param name="NombreElemento"></param>
        ''' <param name="RutaOrigen"></param>
        ''' <param name="Servicio"></param>
        ''' <param name="EncabezadoLog"></param>
        ''' <param name="VerbosidadLog"></param>
        Public Shared Sub LogRutaLarga(ByVal Tipo As TipoElementoRuta, ByVal NombreElemento As String, ByVal RutaOrigen As String,
                                       ByVal RutaElementosLargos As String, ByVal Servicio As String, ByVal EncabezadoLog As String,
                                       ByVal VerbosidadLog As Integer)

            Dim _descripcionElemento As String = Tipo.ToString().ToUpper()
            Dim _nombreLog As String = Tipo.ToString().ToUpper
            Dim _nombreTXTLog As String = NombreElemento & "_LOG_" & _descripcionElemento & ".txt"

            Dim _txt As New StreamWriter(Path.Combine(RutaElementosLargos, _nombreTXTLog))
            _txt.WriteLine("El " & IIf(Tipo = 1, " archivo ", " directorio ") & RutaOrigen & " es muy largo por lo que ha sido puesto en esta carpeta" & vbCrLf)
            _txt.Close()

            ReglasLog.ReglasLog.CrearLog(Servicio, EncabezadoLog & " | LOG " & _descripcionElemento &
                                         " largo creado : " & _nombreTXTLog, TipoLog.Info, VerbosidadLog, 3)

        End Sub

        ''' <summary>
        ''' Devuelve el largo máximo que puede tener una ruta para un archivo o para un directorio
        ''' </summary>
        ''' <param name="Tipo"></param>
        ''' <returns></returns>
        Public Shared Function LargoMaximoRuta(ByVal Tipo As TipoElementoRuta) As Integer
            If Tipo = TipoElementoRuta.Archivo Then
                Return 259
            Else
                Return 239
            End If
        End Function

        ''' <summary>
        ''' Valida criterios del archivo o directorio para permitir procesar o no
        ''' </summary>
        ''' <param name="Tipo"></param>
        ''' <param name="Origen"></param>
        ''' <param name="Destino"></param>
        ''' <param name="ArchivosExcluidos"></param>
        ''' <param name="DirectoriosExcluidos"></param>
        ''' <param name="Servicio"></param>
        ''' <param name="EncabezadoLog"></param>
        ''' <param name="VerbosidadLog"></param>
        ''' <returns></returns>
        Public Shared Function CorrespondeCopiar(ByVal Tipo As TipoElementoRuta, ByVal Origen As String, ByVal Destino As String,
                                          ByVal ArchivosExcluidos As String, ByVal DirectoriosExcluidos As String,
                                          ByVal Servicio As String, ByVal EncabezadoLog As String, ByVal VerbosidadLog As String) As Boolean
            Dim _resultado As Boolean = False

            If Tipo = TipoElementoRuta.Archivo Then
                If Not ValidarExcluido(Servicio, EncabezadoLog, ArchivosExcluidos, Origen, VerbosidadLog) Then
                    If LargoMaximoRuta(Tipo) < Origen.Length Then
                        ' si el largo es mayor se auoriza la copia sin validar fechas (al ser muy largo la fecha no se devuelve bien)
                        _resultado = True
                    Else

                        Dim _fechaModificacionOrigen As Date = File.GetLastWriteTime(Origen)
                        Dim _fechaModificacionDestino As Date = File.GetLastWriteTime(Destino)

                        ' valida o copia por fecha
                        If _fechaModificacionOrigen > _fechaModificacionDestino Then
                            _resultado = True
                        End If

                    End If

                Else
                    _resultado = False
                End If
            End If

            ReglasLog.ReglasLog.CrearLog(Servicio, EncabezadoLog & " | " & IIf(_resultado, "", "NO ") & "Corresponde procesar elemento | Origen : " &
                                             Origen & " | Destino : " & Destino, TipoLog.Info, VerbosidadLog, 1)

            Return _resultado
        End Function

        ''' <summary>
        ''' Actualiza el nombre de un directorio
        ''' </summary>
        ''' <param name="Servicio">Nombre del servicio que solicita el cambio de nombre</param>
        ''' <param name="Directorio">Ruta completa del directorio al que se cambia el nombnre</param>
        ''' <param name="NuevoNombre">Ruta completa del directorio con el nuevo nombre</param>
        ''' <param name="VerbosidadLog">Nivel de verbosidad para LOG's</param>
        ''' <returns></returns>
        Public Shared Function CambiarNombreDirectorio(Servicio As String, Directorio As String, NuevoNombre As String, VerbosidadLog As Integer) As ResponseObject
            Dim _resultado As New ResponseObject

            Try
                ReglasLog.ReglasLog.CrearLog(Servicio, " Actualizará nombre directorio: " & Directorio & " a " & NuevoNombre,
                                             TipoLog.Info, VerbosidadLog, 3)

                Directory.Move(Directorio, NuevoNombre)

                ReglasLog.ReglasLog.CrearLog(Servicio, " Actualizado nombre directorio: " & Directorio & " a " & NuevoNombre,
                                            TipoLog.Info, VerbosidadLog, 3)

            Catch ex As Exception
                ReglasLog.ReglasLog.CrearLog(Servicio, " ERROR al actualizar nombre: " & ex.Source.ToString() & ". " & ex.Message.ToString() & ". " & ex.StackTrace.ToString(),
                                             TipoLog.Err, VerbosidadLog, 1)

            End Try

            Return _resultado
        End Function

        ''' <summary>
        ''' Verifica si existe una ruta
        ''' </summary>
        ''' <param name="Ruta">Ruta a validar</param>
        ''' <returns></returns>
        Public Shared Function ExisteRuta(Servicio As String, Etiqueta As String, Ruta As String, VerbosidadLog As Integer) As ResponseObject
            Dim _resultado As New ResponseObject

            Try
                ReglasLog.ReglasLog.CrearLog(Servicio, Etiqueta & " | Validar si existe ruta | Inicio :" & Ruta,
                                              TipoLog.Info, VerbosidadLog, 2)

                _resultado = New ResponseObject(File.Exists(Ruta))
                If Not _resultado.IsError And _resultado.Objeto = False Then _resultado = New ResponseObject(Directory.Exists(Ruta))

                ReglasLog.ReglasLog.CrearLog(Servicio, Etiqueta & " | Validar si existe ruta | Fin :" & Ruta,
                                              TipoLog.Info, VerbosidadLog, 2)

            Catch ex As Exception
                _resultado = New ResponseObject("Error al validar ruta. " & ex.Message, ex.StackTrace, 0, ex.Source)

                ReglasLog.ReglasLog.CrearLog(Servicio, Etiqueta & " | ERROR :" & ex.Source & ". " &
                                             ex.Message & ". " & ex.StackTrace, TipoLog.Err, VerbosidadLog, 1)
            End Try

            Return _resultado
        End Function

        ''' <summary>
        ''' Verifica si existen archivos en la ruta consultada o en sus sub directorios
        ''' </summary>
        ''' <param name="Servicio"></param>
        ''' <param name="Etiqueta"></param>
        ''' <param name="Ruta"></param>
        ''' <param name="VerbosidadLog"></param>
        ''' <returns></returns>
        Public Shared Function ExistenArchivos(Servicio As String, Etiqueta As String, Ruta As String, VerbosidadLog As Integer) As ResponseObject
            Dim _resultado As New ResponseObject
            Dim _existen As Boolean = False

            Try
                ReglasLog.ReglasLog.CrearLog(Servicio, Etiqueta & " | Validar si existen archivos | Inicio : " & Ruta,
                                              TipoLog.Info, VerbosidadLog, 2)

                If Directory.Exists(Ruta) Then
                    ' verifica si hay archivos en la ruta
                    For Each _archivo As String In Directory.GetFiles(Ruta)
                        _existen = True
                        Exit For
                    Next

                    ' si no hay harchivos verifica si hay sub-carpetas para validar
                    If Not _existen Then
                        For Each _directorio As String In Directory.GetDirectories(Ruta)
                            _resultado = ExistenArchivos(Servicio, Etiqueta, _directorio, VerbosidadLog)
                            If Not _resultado.IsError And _resultado.Objeto = True Then
                                _existen = True
                                Exit For
                            ElseIf _resultado.IsError Then
                                Exit For
                            End If
                        Next
                    End If
                End If

                If Not _resultado.IsError Then
                    _resultado = New ResponseObject(_existen)
                Else
                    Throw New Exception(_resultado.ErrorInfo.Origen & ". " & _resultado.ErrorInfo.Mensaje & ". " &
                        _resultado.ErrorInfo.Detalle)
                End If

            Catch ex As Exception
                _resultado = New ResponseObject("Error al validar existencia de archivos. " & ex.Message, ex.StackTrace, 0, ex.Source)

                ReglasLog.ReglasLog.CrearLog(Servicio, Etiqueta & " | Validar si existen archivos | ERROR :" & ex.Source & ". " &
                                             ex.Message & ". " & ex.StackTrace, TipoLog.Err, VerbosidadLog, 1)
            End Try

            ReglasLog.ReglasLog.CrearLog(Servicio, Etiqueta & " | Validar si existen archivos | Fin : " & Ruta,
                                              TipoLog.Info, VerbosidadLog, 2)
            Return _resultado
        End Function

        ''' <summary>
        ''' Elimina la información dentro de una ruta
        ''' </summary>
        ''' <param name="Ruta"></param>
        ''' <returns></returns>
        Public Shared Function EliminarRuta(Servicio As String, Etiqueta As String, Ruta As String,
                                            VerbosidadLog As Integer, DiasMantener As Integer) As ResponseObject
            Dim _resultado As New ResponseObject
            Dim _mensaje As String = ""

            Try

                ReglasLog.ReglasLog.CrearLog(Servicio, Etiqueta & " | Eliminar ruta | Inicio : " & Ruta,
                                              TipoLog.Info, VerbosidadLog, 2)
                If Directory.Exists(Ruta) Then
                    If DiasMantener > 0 Then
                        ' elimina los archivos de la ruta
                        For Each _archivo As String In Directory.GetFiles(Ruta)
                            Dim _fechaCreacion As Date = File.GetCreationTime(_archivo)
                            If _fechaCreacion.Date < Now.Date.AddDays(-DiasMantener) Then
                                File.Delete(_archivo)
                            End If

                            ReglasLog.ReglasLog.CrearLog(Servicio, Etiqueta & " | Eliminando : " & _archivo, '_directorio,
                                                  TipoLog.Info, VerbosidadLog, 3)
                        Next

                        ' elimina los directorios de la ruta
                        For Each _directorio As String In Directory.GetDirectories(Ruta)
                            Dim _fechaCreacion As Date = Directory.GetCreationTime(_directorio)
                            If _fechaCreacion.Date < Now.Date.AddDays(-DiasMantener) Then
                                Directory.Delete(_directorio, True)
                            End If

                            ReglasLog.ReglasLog.CrearLog(Servicio, Etiqueta & " | Eliminando : " & _directorio, '_directorio,
                                                  TipoLog.Info, VerbosidadLog, 3)
                        Next

                    Else
                        Directory.Delete(Ruta, True)
                    End If

                Else
                    ReglasLog.ReglasLog.CrearLog(Servicio, Etiqueta & " | Eliminar ruta | La ruta no existe : " & Ruta,
                                              TipoLog.Info, VerbosidadLog, 2)
                End If

                ReglasLog.ReglasLog.CrearLog(Servicio, Etiqueta & " | Eliminar ruta | Fin : " & Ruta,
                                              TipoLog.Info, VerbosidadLog, 2)

            Catch ex As Exception
                _resultado = New ResponseObject("Error al eliminar ruta. " & ex.Message, ex.StackTrace, 0, ex.Source)

                ReglasLog.ReglasLog.CrearLog(Servicio, Etiqueta & " | ERROR Eliminar ruta : " & ex.Source.ToString() & ". " &
                                             ex.Message.ToString() & ". " & ex.StackTrace, TipoLog.Err, VerbosidadLog, 1)
            End Try

            Return _resultado
        End Function

#Region "Enumeraciones"
        Public Enum TipoAccionArchivos
            Copiar = 1
            ValidarpreCopia = 2
            BuscarFaltantes = 3
        End Enum

        Public Enum TipoElementoRuta
            Archivo = 1
            Directorio = 2
        End Enum
#End Region

    End Class

End Namespace
