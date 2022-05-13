Imports Solmicro.Expertis.Engine.BE.BusinessProcesses
Imports Solmicro.Expertis
Imports Solmicro.Expertis.Engine.DAL
Imports Solmicro.Expertis.Engine.BE

Public Class Funciones
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "vFrmPrevObrasHoras"

    Public Function RecogeDatos(ByVal sql As String) As DataTable

        Dim dt As DataTable

        dt = AdminData.GetData(sql)

        Return dt
    End Function

End Class