Imports Solmicro.Expertis.Engine.BE.BusinessProcesses
Imports Solmicro.Expertis
Imports Solmicro.Expertis.Engine.DAL
Imports Solmicro.Expertis.Engine.BE

Public Class PrevencionAlertas
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbOperarioAlertas"

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
        data("IDAlertas") = AdminData.GetAutoNumeric

    End Sub
#End Region

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCamposObligatorios)
        'validateProcess.AddTask(Of DataRow)(AddressOf ValidarEstadoPresupuesto)
    End Sub

    <Task()> Public Shared Sub ValidarCamposObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Solmicro.Expertis.Engine.Length(data("IDAlertas")) = 0 Then ApplicationService.GenerateError("El código de la aportación es un dato obligatorio.")
        If Length(data("IDOperario")) = 0 Then ApplicationService.GenerateError("El código del Operario es un dato obligatorio.")

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
    Public Function UpdateForm(ByVal dr As DataRow, Optional ByVal iHistorico As Int16 = 0) As System.Data.DataTable
        ' Recibe la fila del formulario y la trata
        Try
            Dim scolumna() As String
            Dim dtAler = SelOnPrimaryKey(dr(cnEntidad & "_IDAlertas"))
            If IsNothing(dtAler) OrElse dtAler.Rows.Count = 0 Then
                dtAler = MyBase.AddNewForm
                dtAler.Rows(0)("IDAlertas") = AdminData.GetAutoNumeric()
            End If
            ' Control de historizar
            If iHistorico = 1 Then
                dtAler.Rows(0)("snHistorico") = iHistorico
            End If
            ' Recorrer las columnas de la fila y asignar valores
            dtAler.Rows(0)("idOperario") = dr("idOperario") ' <-- El operario de la cabecera
            For Each columna As DataColumn In dr.Table.Columns
                scolumna = columna.ColumnName.Split("_")
                ' Controlar si tiene varias partes y la primera es igual a la entidad
                If scolumna.GetLength(0) > 1 Then
                    If scolumna.GetValue(0) = cnEntidad And columna.ColumnName <> cnEntidad & "_IDAlertas" Then
                        dtAler.Rows(0)(scolumna.GetValue(1)) = dr(columna.ColumnName)
                        ' Si es Historico
                        If iHistorico = 1 Then
                            dr(columna.ColumnName) = DBNull.Value
                        End If
                    End If
                End If
            Next

            ' Control fechas
            If IsDate(dtAler.Rows(0)("fBaja")) And IsDate(dtAler.Rows(0)("fAlta")) Then
                dtAler.Rows(0)("nDias") = DateDiff(DateInterval.Day, dtAler.Rows(0)("fBaja"), dtAler.Rows(0)("fAlta"))
            End If
            ' Update
            Me.Update(dtAler)
            Return dtAler
        Catch ex As Exception
            ApplicationService.GenerateError("Error al crear fila de bajas." & ex.Message)
        End Try
    End Function
    Public Function ActualizaInfo(ByVal dibaja As String, ByVal idacci As String)
        Dim strDel As String = " update tbOperarioAlertas"
        strDel &= " SET ndias = ('" & dibaja & "')"
        strDel &= " WHERE IDAlertas = ('" & idacci & "')"
        Try
            AdminData.Execute(strDel)
        Catch Ex As Exception
            MsgBox("Ha habido un error " & Ex.Message)
        End Try
    End Function
#End Region
End Class
