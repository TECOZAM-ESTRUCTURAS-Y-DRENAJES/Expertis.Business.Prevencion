Imports Solmicro.Expertis.Engine.BE.BusinessProcesses
Imports Solmicro.Expertis
Imports Solmicro.Expertis.Engine.DAL
Imports Solmicro.Expertis.Engine.BE

Public Class PrevencionFormacionTipo
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbOperarioFormacionTipo"

#Region "RegisterAddNewTasks"
    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarValoresPredeterminados, data, services)
        'ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarCentroGestion, data, services)
        'ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarContador, data, services)
    End Sub

    <Task()> Public Shared Sub AsignarValoresPredeterminados(ByVal data As DataRow, ByVal services As ServiceProvider)
        'data("IDFormacion") = AdminData.GetAutoNumeric

    End Sub
#End Region

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCamposObligatorios)
        'validateProcess.AddTask(Of DataRow)(AddressOf ValidarEstadoPresupuesto)
    End Sub

    <Task()> Public Shared Sub ValidarCamposObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        'If Solmicro.Expertis.Engine.Length(data("IDFormacion")) = 0 Then ApplicationService.GenerateError("El EPI es un dato obligatorio.")
        'If Length(data("IDOperario")) = 0 Then ApplicationService.GenerateError("El código de la aportación es un dato obligatorio.")

    End Sub
#End Region

#Region "Eventos RegisterDeleteTasks"
    Protected Overrides Sub RegisterDeleteTasks(ByVal DeleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(DeleteProcess)
        '    addnewProcess.AddTask(Of DataRow)(AddressOf AsignarContadorPorDefecto)
        '    'addnewProcess.AddTask(Of DataRow)(AddressOf AsignarTipoDoc)
    End Sub
#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarValoresPredeterminados)
    End Sub
#End Region

#Region "Funciones"
    Public Function UpdateForm(ByVal sTipoFormacion As String, ByVal itipof As Integer, ByVal snombreVista As String)
        ' Recibe la fila del formulario y la trata
        Try
            Dim strsql As String
            Dim dtRdo As DataTable
            ' Localizar el tipo de formación a guardar
            Dim frdo As New Filter
            Dim DE As New DataEngine
            frdo.Add("ClaseFormacion", itipof)
            frdo.Add("DescClaseFormacion", sTipoFormacion)
            strsql = "SELECT ClaseFormacion,DescClaseFormacion FROM tbOperarioFormacionTipo "
            dtRdo = DE.Filter("tbOperarioFormacionTipo", frdo, "ClaseFormacion,DescClaseFormacion")
            'dtRdo = AdminData.Filter(strsql, , "ClaseFormacion =  " & itipof & " and DescClaseFormacion= '" & sTipoFormacion & "'", , False)

            ' Controlar resultado
            If dtRdo.Rows.Count <= 0 Then
                Dim dtTipoFor = MyBase.AddNewForm
                dtTipoFor.Rows(0)("ClaseFormacion") = itipof
                dtTipoFor.Rows(0)("DescClaseFormacion") = sTipoFormacion
                Me.Update(dtTipoFor)
            End If
        Catch ex As Exception
            ApplicationService.GenerateError("Error al crear fila de tipos de formación." & ex.Message)
        End Try
    End Function
#End Region

End Class

