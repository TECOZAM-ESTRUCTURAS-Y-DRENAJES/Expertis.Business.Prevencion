Imports Solmicro.Expertis.Engine.BE.BusinessProcesses
Imports Solmicro.Expertis
Imports Solmicro.Expertis.Engine.DAL
Imports Solmicro.Expertis.Engine.BE

Public Class PrevencionComboNotifSancion
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbOperarioNotificacionSan"

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
        data("IDSancion") = AdminData.GetAutoNumeric

    End Sub
#End Region

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCamposObligatorios)
        'validateProcess.AddTask(Of DataRow)(AddressOf ValidarEstadoPresupuesto)
    End Sub

    <Task()> Public Shared Sub ValidarCamposObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Solmicro.Expertis.Engine.Length(data("IDSancion")) = 0 Then ApplicationService.GenerateError("El código del tipo de epi es un dato obligatorio.")
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


    Friend Sub AdminSancion(ByVal sSancion As String)
        Try
            Dim DtAdmin As DataTable
            Dim FAdmin As New Filter
            Dim De As New DataEngine
            FAdmin.Add("DescSancion", sSancion.ToUpper)
            DtAdmin = De.Filter("tbOperarioNotificacionSan", FAdmin, "DescSancion", False)
            'DtAdmin = AdminData.Filter("select DescEpi from tbOperarioEpisTipo", , "DescEpi = '" & sEpi.ToUpper & "'", , False)


            If IsNothing(DtAdmin) OrElse DtAdmin.Rows.Count = 0 Then
                DtAdmin = MyBase.AddNewForm
                DtAdmin.Rows(0)("IDSancion") = AdminData.GetAutoNumeric()
                DtAdmin.Rows(0)("DescSancion") = sSancion.ToUpper

                Me.Update(DtAdmin)
            End If
        Catch ex As Exception
            ApplicationService.GenerateError("Error al crear fila de sanciones." & ex.Message)
        End Try
    End Sub

End Class
