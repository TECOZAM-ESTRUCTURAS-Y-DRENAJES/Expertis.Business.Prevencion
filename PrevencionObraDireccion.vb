Imports Solmicro.Expertis.Engine.BE.BusinessProcesses
Imports Solmicro.Expertis
Imports Solmicro.Expertis.Engine.DAL
Imports Solmicro.Expertis.Engine.BE

Public Class PrevencionObraDireccion
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbObraPrevDireccion"

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
        data("IDDireccion") = AdminData.GetAutoNumeric

    End Sub
#End Region

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCamposObligatorios)
        'validateProcess.AddTask(Of DataRow)(AddressOf ValidarEstadoPresupuesto)
    End Sub

    <Task()> Public Shared Sub ValidarCamposObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Solmicro.Expertis.Engine.Length(data("IDDireccion")) = 0 Then ApplicationService.GenerateError("La Direccion de obra es un dato obligatorio.")
        'If Length(data("IDOperario")) = 0 Then ApplicationService.GenerateError("El código del operario es un dato obligatorio.")

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
    'Descomentado David Velasco 03/22
    Public Function UpdateForm2(ByVal dr As DataRow, ByVal ctipo As Char, ByVal nombre As String, ByVal fini As String, ByVal ffin As String, Optional ByVal iHistorico As Int16 = 0) As System.Data.DataTable
        ' Recibe la fila del formulario y la trata
        Try
            Dim scolumna() As String
            Dim svista As String
            svista = "vFrmPrevObraDireccion" & ctipo

            Dim dtDireccion = SelOnPrimaryKey(dr("IDDireccion"))
            'Dim dtDireccion = SelOnPrimaryKey(dr(svista & "_IDDireccion"))
            If IsNothing(dtDireccion) Or dtDireccion.Rows.Count = 0 Then
                dtDireccion = MyBase.AddNewForm
                dtDireccion.Rows(0)("IDDireccion") = AdminData.GetAutoNumeric()
            End If
            ' Control de historizar
            If iHistorico = 1 Then
                dtDireccion.Rows(0)("snHistorico") = iHistorico
            End If
            ' Recorrer las columnas de la fila y asignar valores
            dtDireccion.Rows(0)("idObra") = dr("idObra") ' <-- La obra de la cabecera
            Try
                dtDireccion.Rows(0)("Nombre") = nombre
                dtDireccion.Rows(0)("FInicio") = fini
                dtDireccion.Rows(0)("FFin") = ffin
                dtDireccion.Rows(0)("Tipo") = ctipo
            Catch ex As Exception

            End Try

            ' Update
            Me.Update(dtDireccion)
            Return dtDireccion
        Catch ex As Exception
            ApplicationService.GenerateError("Error al crear fila de reconocimiento." & ex.Message)
        End Try
    End Function

    Public Function UpdateForm(ByVal dr As DataRow, ByVal ctipo As Char, Optional ByVal iHistorico As Int16 = 0) As System.Data.DataTable
        ' Recibe la fila del formulario y la trata
        Try
            Dim scolumna() As String
            Dim svista As String
            svista = "vFrmPrevObraDireccion" & ctipo

            Dim dtDireccion = SelOnPrimaryKey(dr("IDDireccion"))
            'Dim dtDireccion = SelOnPrimaryKey(dr(svista & "_IDDireccion"))
            If IsNothing(dtDireccion) Or dtDireccion.Rows.Count = 0 Then
                dtDireccion = MyBase.AddNewForm
                dtDireccion.Rows(0)("IDDireccion") = AdminData.GetAutoNumeric()
            End If
            ' Control de historizar
            If iHistorico = 1 Then
                dtDireccion.Rows(0)("snHistorico") = iHistorico
            End If
            ' Recorrer las columnas de la fila y asignar valores
            dtDireccion.Rows(0)("idObra") = dr("idObra") ' <-- La obra de la cabecera


            'For Each columna As DataColumn In dr.Table.Columns
            '    scolumna = columna.ColumnName.Split("_")
            '    ' Controlar si tiene varias partes y la primera es igual a la entidad
            '    If scolumna.GetLength(0) > 1 Then
            '        If scolumna.GetValue(0) = svista And columna.ColumnName <> svista & "_IDDireccion" Then
            '            dtDireccion.Rows(0)(scolumna.GetValue(1)) = dr(columna.ColumnName)
            '            ' Si es Historico
            '            If iHistorico = 1 Then
            '                dr(columna.ColumnName) = DBNull.Value
            '            End If
            '        End If
            '    End If
            'Next
            ' Poner el tipo que se pasa
            dtDireccion.Rows(0)("Tipo") = ctipo
            ' Update
            Me.Update(dtDireccion)
            Return dtDireccion
        Catch ex As Exception
            ApplicationService.GenerateError("Error al crear fila de reconocimiento." & ex.Message)
        End Try
    End Function

    <Task()> Public Shared Sub GenerarNuevaDireccion(ByVal data As dataDireccion, ByVal services As ServiceProvider)
        Dim ND As New PrevencionObraDireccion
        Dim dtND As DataTable = ND.AddNewForm
        Dim drND As DataRow = dtND.NewRow

        dtND.Rows(0)("IDDireccion") = AdminData.GetAutoNumeric
        dtND.Rows(0)("idObra") = data.IdObra
        dtND.Rows(0)("Nombre") = data.Nombre
        dtND.Rows(0)("Tipo") = data.ctipo
        dtND.Rows(0)("Finicio") = data.FInicio
        dtND.Rows(0)("FFin") = data.FFin
        dtND.Rows(0)("snHistorico") = data.snHistorico

        Dim upg As New UpdatePackage
        upg.Add(dtND)
        ND.Update(upg)
    End Sub

#End Region

End Class
