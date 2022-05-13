Imports Solmicro.Expertis.Engine.BE.BusinessProcesses
Imports Solmicro.Expertis
Imports Solmicro.Expertis.Engine.DAL
Imports Solmicro.Expertis.Engine.BE

Public Class PrevencionReconocimiento
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbOperarioReconocimiento"

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
        data("IDReconocimiento") = AdminData.GetAutoNumeric

    End Sub
#End Region

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCamposObligatorios)
        'validateProcess.AddTask(Of DataRow)(AddressOf ValidarEstadoPresupuesto)
    End Sub

    <Task()> Public Shared Sub ValidarCamposObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Solmicro.Expertis.Engine.Length(data("IDReconocimiento")) = 0 Then ApplicationService.GenerateError("El código del reconocimiento es un dato obligatorio.")
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
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarValoresPredeterminados)
        'updateProcess.AddTask(Of DataRow)(AddressOf ValidarCamposObligatorios)
    End Sub
#End Region

#Region "Funciones"
    'Public Function UpdateProceso(ByVal dttSource As System.Data.DataTable) As System.Data.DataTable
    '    ' Cargar el esquema y generar la sentencia de comandos.
    '    dttSource.TableName = cnEntidad
    '    ConobjR.Actualizar(dttSource)
    '    Return dttSource
    'End Function

    Public Function UpdateForm(ByVal dr As DataRow, Optional ByVal iHistorico As Int16 = 0) As System.Data.DataTable
        ' Recibe la fila del formulario y la trata
        Try
            Dim scolumna() As String
            Dim dtReco = SelOnPrimaryKey(dr(cnEntidad & "_IDReconocimiento"))
            If IsNothing(dtReco) OrElse dtReco.Rows.Count = 0 Then
                dtReco = MyBase.AddNewForm
                dtReco.Rows(0)("IDReconocimiento") = AdminData.GetAutoNumeric()
            End If
            ' Control de historizar
            If iHistorico = 1 Then
                dtReco.Rows(0)("snHistorico") = iHistorico
            End If
            ' Recorrer las columnas de la fila y asignar valores
            dtReco.Rows(0)("idOperario") = dr("idOperario") ' <-- El operario de la cabecera
            For Each columna As DataColumn In dr.Table.Columns
                scolumna = columna.ColumnName.Split("_")
                ' Controlar si tiene varias partes y la primera es igual a la entidad
                If scolumna.GetLength(0) > 1 Then
                    If scolumna.GetValue(0) = cnEntidad And columna.ColumnName <> cnEntidad & "_IDReconocimiento" Then
                        dtReco.Rows(0)(scolumna.GetValue(1)) = dr(columna.ColumnName)
                        ' Si es Historico
                        If iHistorico = 1 Then
                            dr(columna.ColumnName) = DBNull.Value
                        End If
                    End If
                End If
            Next
            ' Update
            Me.Update(dtReco)
            Return dtReco
        Catch ex As Exception
            ApplicationService.GenerateError("Error al crear fila de reconocimiento." & ex.Message)
        End Try
    End Function
    'Friend Function GenerarRemesaEnvio(ByVal dfereco As Date, Optional ByRef ConobjR As ClaseConexion = Nothing) As System.Data.DataTable
    '    Try
    '        Dim sWhere As String
    '        Dim Dt As DataTable
    '        ' Compongo el Where con la fecha pasada. Menores o iguales a la fecha seguimiento pasada, campo ndiasaviso
    '        sWhere = "nDiasAviso <= '" & dfereco & "' and idRemesa is null"
    '        If IsNothing(ConobjR) Then
    '            Dt = AdminData.Filter("select * from " & cnEntidad, , sWhere, , False).DataTable
    '        Else
    '            Dt = ConobjR.Filtrar("select * from " & cnEntidad & " WHERE " & sWhere)
    '            Dt.TableName = cnEntidad
    '        End If
    '        Return Dt
    '    Catch ex As Exception
    '        Console.WriteLine(Now & " - Error al obtener los reconocimientos que van a caducar." & ex.Message)
    '    End Try
    'End Function
#End Region



End Class
