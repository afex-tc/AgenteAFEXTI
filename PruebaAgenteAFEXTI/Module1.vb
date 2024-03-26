
Imports AgenteAFEXTI

Module Module1

    Sub Main(args As String())
        Dim _agente As New AgenteAFEXTI.ReglasAgente.ReglasAgente

        servicioagenteafexti.Ejecutar()
        _agente.VerificarProgramacionHOST()
    End Sub

End Module
