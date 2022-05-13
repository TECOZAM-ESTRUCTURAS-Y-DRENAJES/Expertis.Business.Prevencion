Public Class dataCopia

End Class

<Serializable()> _
Public Class dataVisita
    Public IdVisitas As Integer
    Public idObra As Integer
    Public nvisita As Integer
    Public fecha As Date
    Public Tecnico As String
    Public Comunicado As String
    Public Observaciones As String
    Public snHistorico As Integer

    Public Sub New()

    End Sub

End Class
<Serializable()> _
Public Class dataInformacion
    Public nombre As String
    Public idObra As Integer
    Public puesto As String
    Public telefono As String
    Public email As String
    Public Observaciones As String
    Public snHistorico As Integer

    Public Sub New()

    End Sub

End Class
<Serializable()> _
Public Class dataRec
    Public rec As String
    Public idObra As Integer
    Public finicio As Date
    Public ffin As Date
    Public codop As String
    Public operario As String
    Public snHistorico As Integer

    Public Sub New()

    End Sub

End Class
<Serializable()> _
Public Class dataDoc
    Public fecha As Date
    Public idObra As Integer
    Public envio As String
    Public observaciones As String
    Public solcita As String
    Public solicitante As String
    Public snHistorico As Integer

    Public Sub New()

    End Sub

End Class

<Serializable()> _
Public Class dataDireccion
    Public idDireccion As Integer
    Public IdObra As Integer
    Public ctipo As String
    Public Nombre As String
    Public FInicio As Date
    Public FFin As Date
    Public snHistorico As Integer


    Public Sub New()

    End Sub
End Class
