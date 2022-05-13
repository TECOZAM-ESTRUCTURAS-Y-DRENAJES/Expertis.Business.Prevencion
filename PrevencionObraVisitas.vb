Imports Solmicro.Expertis.Engine.BE.BusinessProcesses
Imports Solmicro.Expertis
Imports Solmicro.Expertis.Engine.DAL
Imports Solmicro.Expertis.Engine.BE

Public Class PrevencionObraVisitas
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbObraPrevVisitas"

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
        data("IDVisitas") = AdminData.GetAutoNumeric

    End Sub
#End Region

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCamposObligatorios)
        'validateProcess.AddTask(Of DataRow)(AddressOf ValidarEstadoPresupuesto)
    End Sub

    <Task()> Public Shared Sub ValidarCamposObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Solmicro.Expertis.Engine.Length(data("IDVisitas")) = 0 Then ApplicationService.GenerateError("El código de la visita es un dato obligatorio.")
        'If Length(data("IDVisitas")) = 0 Then ApplicationService.GenerateError("El código del operario es un dato obligatorio.")

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
        updateProcess.AddTask(Of DataRow)(AddressOf ValidarCamposObligatorios)
    End Sub
#End Region

#Region "Funciones"
    Public Function Contvisitas(ByVal idObra As Integer) As Short
        Try
            Dim dtr As DataTable
            Dim sselect As String

            Dim DE As New DataEngine

            sselect = "SELECT max(nvisita) AS nvisita FROM tbObraPrevVisitas WHERE (idObra = " & idObra & ")"
            dtr = AdminData.GetData(sselect, , , , False)

            If IsNothing(dtr) Then
                Return 1
            ElseIf dtr.Rows.Count > 0 Then
                If IsDBNull(dtr.Rows(0)("nvisita")) Then
                    Return 1
                Else
                    Return dtr.Rows(0)("nvisita") + 1
                End If
            End If

        Catch ex As Exception
            ApplicationService.GenerateError("Error al obtener contador de visitas." & ex.Message)
        End Try
    End Function

    'Public Function UpdateForm(ByVal dr As DataRow, Optional ByVal iHistorico As Int16 = 0) As System.Data.DataTable
    '    ' Recibe la fila del formulario y la trata
    '    Try
    '        Dim scolumna() As String
    '        Dim dtVisita = SelOnPrimaryKey(dr("IDVisitas"))
    '        If IsNothing(dtVisita) OrElse dtVisita.Rows.Count = 0 Then
    '            dtVisita = MyBase.AddNewForm
    '            dtVisita.Rows(0)("IDVisitas") = AdminData.GetAutoNumeric()
    '        End If
    '        ' Control de historizar
    '        If iHistorico = 1 Then
    '            dtVisita.Rows(0)("snHistorico") = iHistorico
    '        End If
    '        ' Recorrer las columnas de la fila y asignar valores
    '        dtVisita.Rows(0)("idObra") = dr("idObra") ' <-- La obra de la cabecera
    '        For Each columna As DataColumn In dr.Table.Columns
    '            scolumna = columna.ColumnName.Split("_")
    '            ' Controlar si tiene varias partes y la primera es igual a la entidad
    '            If scolumna.GetLength(0) > 1 Then
    '                If scolumna.GetValue(0) = cnEntidad And columna.ColumnName <> cnEntidad & "_IDVisitas" Then
    '                    dtVisita.Rows(0)(scolumna.GetValue(1)) = dr(columna.ColumnName)
    '                    ' Si es Historico
    '                    If iHistorico = 1 Then
    '                        dr(columna.ColumnName) = DBNull.Value
    '                    End If
    '                End If
    '            End If
    '        Next
    '        ' Validar al final el contador de obras
    '        If IsNothing(dtVisita.Rows(0)("nvisita")) Or IsDBNull(dtVisita.Rows(0)("nvisita")) Then
    '            dtVisita.Rows(0)("nvisita") = Contvisitas(dr("idObra"))
    '        End If
    '        ' Update
    '        Me.Update(dtVisita)
    '        Return dtVisita
    '    Catch ex As Exception
    '        ApplicationService.GenerateError("Error al crear fila de reconocimiento." & ex.Message)
    '    End Try
    'End Function

    Public Function ValidarVisita(ByVal idObra As Integer, ByVal nvisita As Integer)
        Dim dtr As DataTable
        Dim sselect As String

        Dim DE As New DataEngine

        sselect = "SELECT nvisita FROM tbObraPrevVisitas WHERE (idObra = " & idObra & ") and nvisita =" & nvisita
        dtr = AdminData.GetData(sselect, , , , False)

        If dtr.Rows.Count > 0 Then
            Return Contvisitas(idObra)
        Else
            Return 1
        End If
    End Function

    <Task()> Public Shared Sub GenerarNuevaVisita(ByVal data As dataVisita, ByVal services As ServiceProvider)
        Dim NV As New PrevencionObraVisitas
        Dim dtNV As DataTable = NV.AddNewForm
        Dim drNV As DataRow = dtNV.NewRow

        dtNV.Rows(0)("IDVisitas") = AdminData.GetAutoNumeric
        dtNV.Rows(0)("idObra") = data.idObra
        'Validar al final el contador de obras
        If IsNothing(data.nvisita) Or IsDBNull(data.nvisita) Then
            dtNV.Rows(0)("nvisita") = NV.Contvisitas(data.nvisita)
        Else
            dtNV.Rows(0)("nvisita") = NV.ValidarVisita(data.idObra, data.nvisita)
        End If
        'dtNV.Rows(0)("nvisita") = data.nvisita
        dtNV.Rows(0)("fecha") = data.fecha
        dtNV.Rows(0)("Tecnico") = data.Tecnico
        dtNV.Rows(0)("Comunicado") = data.Comunicado
        dtNV.Rows(0)("Observaciones") = data.Observaciones
        dtNV.Rows(0)("snHistorico") = data.snHistorico

        Dim upg As New UpdatePackage
        upg.Add(dtNV)
        NV.Update(upg)
    End Sub

    <Task()> Public Shared Sub UpdateVisita(ByVal data As dataVisita, ByVal services As ServiceProvider)

    End Sub
#End Region

End Class
