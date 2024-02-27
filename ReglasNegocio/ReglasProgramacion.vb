Namespace ReglasProgramacion
    Public Class Programacion
        Private _idColaProgramacion As String
        Private _tipoProgramacion As Integer
        Private _correoConfirmacion As String
        Private _correoError As String
        Private _rutaOrigen As String
        Private _rutaDestino As String
        Private _rutaRespaldo As String
        Private _rutaRetencionOrigen As String
        Private _archivosExcluidos As String
        Private _directoriosExcluidos As String
        Private _validarCopia As Boolean
        Private _eliminarOrigen As Boolean
        Private _verbosidadLog As Integer
        Private _descripcionProgramacion As String
        Private _tipoColaProgramacion As TipoColaProgramacion
        Private _cambiarNombreOrigen As Boolean
        Private _diasCopiasAnteriores As Integer
        Private _directorioFecha As Boolean
        Private _diasRespaldo As Integer
        Public Property IdColaProgramacion()
            Set(value)
                Me._idColaProgramacion = value
            End Set
            Get
                Return Me._idColaProgramacion
            End Get
        End Property
        Public Property TipoProgramacion()
            Set(value)
                Me._tipoProgramacion = value
            End Set
            Get
                Return Me._tipoProgramacion
            End Get
        End Property

        Public Property CorreoConfirmacion()
            Set(value)
                Me._correoConfirmacion = value
            End Set
            Get
                Return Me._correoConfirmacion
            End Get
        End Property

        Public Property RutaOrigen()
            Set(value)
                Me._rutaOrigen = value
            End Set
            Get
                Return Me._rutaOrigen
            End Get
        End Property

        Public Property RutaDestino()
            Set(value)
                Me._rutaDestino = value
            End Set
            Get
                Return Me._rutaDestino
            End Get
        End Property

        Public Property RutaRetencionOrigen()
            Set(value)
                Me._rutaRetencionOrigen = value
            End Set
            Get
                Return Me._rutaRetencionOrigen
            End Get
        End Property

        Public Property RutaRespaldo()
            Set(value)
                Me._rutaRespaldo = value
            End Set
            Get
                Return Me._rutaRespaldo
            End Get
        End Property

        Public Property ArchivosExcluidos()
            Set(value)
                Me._archivosExcluidos = value
            End Set
            Get
                Return Me._archivosExcluidos
            End Get
        End Property

        Public Property DirectoriosExcluidos()
            Set(value)
                Me._directoriosExcluidos = value
            End Set
            Get
                Return Me._directoriosExcluidos
            End Get
        End Property

        Public Property ValidarCopia()
            Set(value)
                Me._validarCopia = value
            End Set
            Get
                Return Me._validarCopia
            End Get
        End Property

        Public Property EliminarOrigen()
            Set(value)
                Me._eliminarOrigen = value
            End Set
            Get
                Return Me._eliminarOrigen
            End Get
        End Property

        Public Property VerbosidadLog()
            Set(value)
                Me._verbosidadLog = value
            End Set
            Get
                Return Me._verbosidadLog
            End Get
        End Property

        Public Property DescripcionProgramacion()
            Set(value)
                Me._descripcionProgramacion = value
            End Set
            Get
                Return Me._descripcionProgramacion
            End Get
        End Property

        Public Property CorreoError()
            Set(value)
                Me._correoError = value
            End Set
            Get
                Return Me._correoError
            End Get
        End Property

        Public Property TipoColaProgramacion()
            Set(value)
                Me._tipoColaProgramacion = value
            End Set
            Get
                Return Me._tipoColaProgramacion
            End Get
        End Property

        Public Property CambiarNombreOrigen()
            Set(value)
                Me._cambiarNombreOrigen = value
            End Set
            Get
                Return Me._cambiarNombreOrigen
            End Get
        End Property

        Public Property DiasCopiasAnteriores()
            Set(value)
                Me._diasCopiasAnteriores = value
            End Set
            Get
                Return Me._diasCopiasAnteriores
            End Get
        End Property

        Public Property DirectorioFecha()
            Set(value)
                Me._directorioFecha = value
            End Set
            Get
                Return Me._directorioFecha
            End Get
        End Property

        Public Property DiasRespaldo()
            Set(value)
                Me._diasRespaldo = value
            End Set
            Get
                Return Me._diasRespaldo
            End Get
        End Property

        Public Sub New(Reader As SqlClient.SqlDataReader)
            If Reader.Read Then

                If Reader("idcolaprogramacion") <> 0 Then
                    Me._idColaProgramacion = Reader("idcolaprogramacion")
                    Me._tipoProgramacion = Reader("tipoprogramacion")
                    Me._descripcionProgramacion = Reader("descripcionprogramacion")
                    Me._verbosidadLog = Reader("verbosidadlog")
                    Me._idColaProgramacion = Reader("idColaProgramacion")
                    Me._correoConfirmacion = Reader("correoConfirmacion")
                    Me._correoError = Reader("correoError")
                    Me._rutaOrigen = Reader("rutaOrigen")
                    Me._rutaDestino = Reader("rutaDestino")
                    If Not IsDBNull(Reader("rutaRespaldo")) Then Me._rutaRespaldo = Reader("rutaRespaldo")
                    If Not IsDBNull(Reader("rutaRetencionOrigen")) Then Me._rutaRetencionOrigen = Reader("rutaRetencionOrigen")
                    If Not IsDBNull(Reader("archivosExcluidos")) Then Me._archivosExcluidos = Reader("archivosExcluidos")
                    If Not IsDBNull(Reader("directoriosExcluidos")) Then Me._directoriosExcluidos = Reader("directoriosExcluidos")
                    Me._validarCopia = Reader("validarCopia")
                    Me._eliminarOrigen = Reader("eliminarOrigen")
                    Me._tipoColaProgramacion = Reader("tipocolaprogramacion")
                    Me._cambiarNombreOrigen = Reader("cambiarnombreorigen")
                    Me._diasCopiasAnteriores = Reader("diascopiasanteriores")
                    Me._directorioFecha = Reader("directoriofecha")
                    Me._diasRespaldo = Reader("diasrespaldo")
                End If
            End If
        End Sub
    End Class

    Public Enum TipoColaProgramacion
        Copiar = 1
        Validar = 2
        CopiaProgramada = 3
    End Enum

End Namespace

