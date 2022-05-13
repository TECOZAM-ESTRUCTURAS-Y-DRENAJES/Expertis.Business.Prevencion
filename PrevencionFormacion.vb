Imports Solmicro.Expertis.Engine.BE.BusinessProcesses
Imports Solmicro.Expertis
Imports Solmicro.Expertis.Engine.DAL
Imports Solmicro.Expertis.Engine.BE

Public Class PrevencionFormacion
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbOperarioFormacion"

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
        data("IDFormacion") = AdminData.GetAutoNumeric

    End Sub
#End Region

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCamposObligatorios)
        'validateProcess.AddTask(Of DataRow)(AddressOf ValidarEstadoPresupuesto)
    End Sub

    <Task()> Public Shared Sub ValidarCamposObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Solmicro.Expertis.Engine.Length(data("IDFormacion")) = 0 Then ApplicationService.GenerateError("El EPI es un dato obligatorio.")
        If Length(data("IDOperario")) = 0 Then ApplicationService.GenerateError("El código de la aportación es un dato obligatorio.")

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
    Public Function UpdateForm(ByVal dr As DataRow, ByVal snombreVista As String, Optional ByVal iHistorico As Int16 = 0) As System.Data.DataTable
        ' Recibe la fila del formulario y la trata
        Try
            Dim itipof As Integer
            Select Case snombreVista
                Case "vFrmPrevencionFormacionI"
                    itipof = 1
                Case "vFrmPrevencionFormacionE"
                    itipof = 2
                Case "vFrmPrevencionFormacionTPC"
                    itipof = 3
                Case "vFrmPrevencionFormacion8"
                    itipof = 8
                Case "vFrmPrevencionFormacion20"
                    itipof = 20
                Case "vFrmPrevencionFormacion50"
                    itipof = 50
                Case Else
                    itipof = 2
            End Select
            Dim scolumna() As String
            Dim dtAc = SelOnPrimaryKey(dr(snombreVista & "_idFormacion"))
            If IsNothing(dtAc) OrElse dtAc.Rows.Count = 0 Then
                dtAc = MyBase.AddNewForm
                dtAc.Rows(0)("idFormacion") = AdminData.GetAutoNumeric()
            End If
            ' Control de historizar
            If iHistorico = 1 Then
                dtAc.Rows(0)("snHistorico") = iHistorico
            End If
            ' Recorrer las columnas de la fila y asignar valores
            dtAc.Rows(0)("idOperario") = dr("idOperario") ' <-- El operario de la cabecera
            For Each columna As DataColumn In dr.Table.Columns
                scolumna = columna.ColumnName.Split("_")
                ' Controlar si tiene varias partes y la primera es igual a la entidad
                If scolumna.GetLength(0) > 1 Then
                    If scolumna.GetValue(0) = snombreVista And columna.ColumnName <> snombreVista & "_idFormacion" Then
                        dtAc.Rows(0)(scolumna.GetValue(1)) = dr(columna.ColumnName)
                        ' Control Columna DescClaseFormacion guardar en tabla aux
                        If columna.ColumnName = snombreVista & "_DescClaseFormacion" And Not IsDBNull(dr(columna.ColumnName)) Then
                            Dim PFT As New PrevencionFormacionTipo
                            PFT.UpdateForm(dr(columna.ColumnName), itipof, snombreVista)
                            PFT = Nothing
                        End If
                        ' Si es Historico
                        If iHistorico = 1 Then
                            dr(columna.ColumnName) = DBNull.Value
                        End If
                    End If
                End If
            Next
            '' Grabamos también el tipo 1 vez
            dtAc.Rows(0)("claseformacion") = itipof

            '' Control de cursos de 6 o 20
            Select Case itipof
                Case 8
                    dtAc.rows(0)("DescClaseFormacion") = "Formación 8h."
                Case 20
                    If IsDBNull(dtAc.Rows(0)("horas")) Then
                        dtAc.Rows(0)("horas") = 6
                    End If
            End Select
            ' Update
            Me.Update(dtAc)
            Return dtAc
        Catch ex As Exception
            ApplicationService.GenerateError("Error al crear fila de Formación." & ex.Message)
        End Try
    End Function
#End Region

End Class
