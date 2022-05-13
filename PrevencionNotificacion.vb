Imports Solmicro.Expertis.Engine.BE.BusinessProcesses
Imports Solmicro.Expertis
Imports Solmicro.Expertis.Engine.DAL
Imports Solmicro.Expertis.Engine.BE

Public Class PrevencionNotificacion
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbOperarioNotificacion"

#Region "RegisterAddNewTasks"
    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        'ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarValoresPredeterminados, data, services)
        'ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarCentroGestion, data, services)
        'ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarContador, data, services)
    End Sub

    <Task()> Public Shared Sub AsignarValoresPredeterminados(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("IDNotificacion") = AdminData.GetAutoNumeric

    End Sub
#End Region

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCamposObligatorios)
        'validateProcess.AddTask(Of DataRow)(AddressOf ValidarEstadoPresupuesto)
    End Sub

    <Task()> Public Shared Sub ValidarCamposObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Solmicro.Expertis.Engine.Length(data("IDNotificacion")) = 0 Then ApplicationService.GenerateError("La notificacion es un dato obligatorio.")
        If Length(data("IDOperario")) = 0 Then ApplicationService.GenerateError("El código del operario es un dato obligatorio.")

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
        'updateProcess.AddTask(Of DataRow)(AddressOf AsignarValoresPredeterminados)
    End Sub
#End Region

    Public Overridable Overloads Function Update(ByVal dttSource As System.Data.DataTable) As System.Data.DataTable
        If Not dttSource Is Nothing AndAlso dttSource.Rows.Count > 0 Then
            For Each dr As DataRow In dttSource.Rows
                If dr.RowState = DataRowState.Added Or dr.RowState = DataRowState.Modified Then
                    If dr.RowState = DataRowState.Added Then
                        If Length(dr("IDNotificacion")) = 0 Then
                            dr("IDNotificacion") = AdminData.GetAutoNumeric
                        End If

                        If Length(dr("IDOperario")) = 0 Then
                            ApplicationService.GenerateError("No se ha indicado el operario")
                        End If
                    End If
                End If
            Next

            MyBase.Update(dttSource)
        End If
        Return dttSource
    End Function

#Region "Funciones"
    Public Function UpdateForm(ByVal dr As DataRow, Optional ByVal iHistorico As Int16 = 0) As System.Data.DataTable
        ' Recibe la fila del formulario y la trata
        Try
            Dim scolumna() As String
            Dim dtNotif = SelOnPrimaryKey(dr(cnEntidad & "_idnotificacion"))
            If IsNothing(dtNotif) OrElse dtNotif.Rows.Count = 0 Then
                dtNotif = MyBase.AddNewForm
                dtNotif.Rows(0)("idnotificacion") = AdminData.GetAutoNumeric()
            End If
            ' Control de historizar
            If iHistorico = 1 Then
                dtNotif.Rows(0)("snHistorico") = iHistorico
            End If
            ' Recorrer las columnas de la fila y asignar valores
            dtNotif.Rows(0)("idOperario") = dr("idOperario") ' <-- El operario de la cabecera
            For Each columna As DataColumn In dr.Table.Columns
                scolumna = columna.ColumnName.Split("_")
                ' Controlar si tiene varias partes y la primera es igual a la entidad
                If scolumna.GetLength(0) > 1 Then
                    If scolumna.GetValue(0) = cnEntidad And columna.ColumnName <> cnEntidad & "_idnotificacion" Then
                        dtNotif.Rows(0)(scolumna.GetValue(1)) = dr(columna.ColumnName)
                        ' Si es Historico
                        If iHistorico = 1 Then
                            dr(columna.ColumnName) = DBNull.Value
                        End If
                    End If
                End If
            Next
            ' Control de Encargado y Jefe
            Dim PEJ As New PrevencionEncargaJefe
            If Not IsDBNull(dr(cnEntidad & "_encargado")) Then
                PEJ.AdminEncargaJefe(dr(cnEntidad & "_encargado"), "E")
            End If
            If Not IsDBNull(dr(cnEntidad & "_JefeProd")) Then
                PEJ.AdminEncargaJefe(dr(cnEntidad & "_JefeProd"), "J")
            End If
            PEJ = Nothing
            ' Control de combo de tipo Sanción
            Dim PCS As New PrevencionCombonotifSancion
            If Not IsDBNull(dr(cnEntidad & "_DescSancion")) Then
                PCS.AdminSancion(dr(cnEntidad & "_DescSancion"))
            End If
            PCS = Nothing
            ' Update
            Me.Update(dtNotif)
            Return dtNotif
        Catch ex As Exception
            ApplicationService.GenerateError("Error al crear fila de accidentes." & ex.Message)
        End Try
    End Function

#End Region

End Class
