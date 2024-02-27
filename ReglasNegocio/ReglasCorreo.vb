
Namespace ReglasCorreo

    Public Class ReglasCorreo

        ''' <summary>
        ''' Envía correo a través de ESB_AFEX
        ''' </summary>
        ''' <param name="Para"></param>
        ''' <param name="Copia"></param>
        ''' <param name="Asunto"></param>
        ''' <param name="Cuerpo"></param>
        Public Shared Sub EnviarCorreo(Para As String, Copia As String, Asunto As String, Cuerpo As String)
            Try
                Dim _mensaje As New IESB_AV_ServicioMensajes.IESB_AV_ServicioMensajes
                Dim _response As Object = _mensaje.EnvioCorreoElectronico_EstandarSQL(Para, Copia, Asunto, Cuerpo, "HTML", "")
            Catch ex As Exception

            End Try

        End Sub


    End Class

End Namespace