Imports Solmicro.Expertis.Engine.BE.BusinessProcesses
Imports Solmicro.Expertis.Engine.BE
Imports Solmicro.Expertis
Imports Solmicro.Expertis.Engine.DAL

Public Class EstadisticasPreven
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbObraModControl"

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

    '<Task()> Public Shared Sub AsignarValoresPredeterminados(ByVal data As DataRow, ByVal services As ServiceProvider)
    '    data("IDPresup") = AdminData.GetAutoNumeric
    '    data("codAportacion") = ""
    'End Sub
#End Region

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCamposObligatorios)
        'validateProcess.AddTask(Of DataRow)(AddressOf ValidarEstadoPresupuesto)
    End Sub

    <Task()> Public Shared Sub ValidarCamposObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        'If Solmicro.Expertis.Engine.Length(data("IDOperario")) = 0 Then ApplicationService.GenerateError("El código de la aportación es un dato obligatorio.")
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
        'updateProcess.AddTask(Of DataRow)(AddressOf AsignarValoresPredeterminados)
    End Sub
#End Region

#Region "Funciones"
    Public Function NumTrabaMesObra(ByVal idObra As Integer, ByVal pMes As Integer, ByVal pyear As Int16, Optional ByVal b20 As Boolean = True) As String
        Try
            Dim ddesde, dHasta As DateTime
            Dim sretorno As String = ""
            Dim imax, isumTrab As Integer
            Dim fMedia As Decimal
            ' Mes de 20 a < 21,13/09/2012, O Natural 1 al < 1
            If b20 Then
                ddesde = CDate("21/" & pMes & "/" & pyear)
                ddesde = DateAdd(DateInterval.Month, -1, ddesde)
                dHasta = CDate("20/" & pMes & "/" & pyear)

            Else
                ddesde = CDate("01/" & pMes & "/" & pyear)
                dHasta = DateAdd(DateInterval.Month, 1, ddesde)
            End If

            Dim sSelect, sWhere, sOrder As String
            Dim Dt As DataTable
            ' 20 mes anterior al 20 mes selec usuario,13/09/2012
            sSelect = "select count(distinct idoperario) as TrabaDia,fechainicio,idobra from tbobraModControl"
            sWhere = " where idobra = " & idObra & " and fechaInicio >='" & ddesde & "'" & " and fechaInicio < '" & dHasta & "' group by fechaInicio,fechaInicio,idObra"
            Dt = AdminData.GetData(sSelect & sWhere)

            If Not IsNothing(Dt) Then
                If Dt.Rows.Count > 0 Then
                    imax = 0
                    For Each ifila As DataRow In Dt.Rows
                        If ifila(0) > imax Then imax = ifila(0)
                        isumTrab += ifila("TrabaDia")
                    Next
                    ' La media
                    fMedia = isumTrab / Dt.Rows.Count
                    sretorno &= imax & "-" & Decimal.Round(fMedia, 3)
                End If
            End If
            ' Retornar resultado
            If sretorno.Length = 0 Then sretorno = "0-0"
            Return sretorno
        Catch ex As Exception
            ApplicationService.GenerateError("Error al obtener la media de trabajadores en obra.Estadistica 1." & ex.Message)
        End Try
    End Function

    Public Function MediaHorasObra(ByVal idObra As Integer, ByVal pMes As Integer, ByVal pyear As Int16, Optional ByVal b20 As Boolean = True) As Decimal
        Try
            Dim ddesde, dHasta As DateTime
            Dim idiastrab As Integer
            Dim dsumHoras As Decimal
            ' Mes de 20 a < 21,13/09/2012, O Natural 1 al < 1
            If b20 Then
                ddesde = CDate("21/" & pMes & "/" & pyear)
                ddesde = DateAdd(DateInterval.Month, -1, ddesde)
                dHasta = CDate("20/" & pMes & "/" & pyear)

            Else
                ddesde = CDate("01/" & pMes & "/" & pyear)
                dHasta = DateAdd(DateInterval.Month, 1, ddesde)
            End If


            Dim sSelect, sWhere, sOrder As String
            Dim Dt As DataTable
            ' Obtener Horas
            sSelect = "select sum(horasRealMod) As Horas from tbobraModControl"
            sWhere = " where idobra = " & idObra & " and fechaInicio >='" & ddesde & "'" & " and fechaInicio < '" & dHasta & "'"
            Dt = AdminData.GetData(sSelect & sWhere)
            If Not IsNothing(Dt) Then
                If Dt.Rows.Count > 0 And Not IsDBNull(Dt.Rows(0)(0)) Then
                    dsumHoras = Dt.Rows(0)(0)
                Else
                    ' No hay horas retornar media 0
                    Return 0
                End If
            End If
            ' Obtener Nº días trabajados por trabajador
            sSelect = "select idoperario,fechainicio from tbobraModControl"
            sWhere = " where idobra = " & idObra & " and fechaInicio >='" & ddesde & "'" & " and fechaInicio < '" & dHasta & "' group by idoperario,fechainicio"
            Dt = AdminData.GetData(sSelect & sWhere)
            If Not IsNothing(Dt) Then
                If Dt.Rows.Count > 0 Then
                    idiastrab = Dt.Rows.Count
                End If
            End If
            ' Retornar resultado
            Return Decimal.Round(dsumHoras / idiastrab, 3)
        Catch ex As Exception
            ApplicationService.GenerateError("Error al obtener la media de horas en obra.Estadistica 2." & ex.Message)
        End Try
    End Function

    ' Punto 3,4
    Public Function AccidentesDias(ByVal idObra As Integer, ByVal pMes As Integer, ByVal pyear As Int16, Optional ByVal b20 As Boolean = True) As String
        Try
            Dim ddesde, dHasta As DateTime
            Dim iacc, idiasBaja As Integer
            ' Mes de 20 a < 21,13/09/2012, O Natural 1 al < 1
            If b20 Then
                ddesde = CDate("21/" & pMes & "/" & pyear)
                ddesde = DateAdd(DateInterval.Month, -1, ddesde)
                dHasta = CDate("20/" & pMes & "/" & pyear)

            Else
                ddesde = CDate("01/" & pMes & "/" & pyear)
                dHasta = DateAdd(DateInterval.Month, 1, ddesde)
            End If

            Dim sSelect, sWhere, sOrder As String
            Dim Dt As DataTable
            ' Obtener Accidentes
            sSelect = "SELECT COUNT(idOperario) AS Accidentes FROM  vFrmPrevAccidentesObra "
            sWhere = " where idobra = " & idObra & " and fAccidente >='" & ddesde & "'" & " and fAccidente < '" & dHasta & "'"
            Dt = AdminData.GetData(sSelect & sWhere)
            If Not IsNothing(Dt) Then
                If Dt.Rows.Count > 0 And Not IsDBNull(Dt.Rows(0)(0)) Then
                    iacc = Dt.Rows(0)(0)
                Else
                    iacc = 0
                End If
            End If
            ' Obtener Nº días perdidos
            sSelect = "SELECT  sum((datediff(dd, CASE WHEN FBaja > '" & ddesde & "' THEN fBaja ELSE '" & ddesde & "' END, CASE WHEN Falta < '" & dHasta & "' THEN falta ELSE '" & dHasta & "' END))) AS Dias FROM tbOperarioAccidentes "
            sWhere = " where codobra = " & idObra
            ' F. Alta null y F.Baja <= ddesde
            sWhere &= " and ( (falta is null) and (fbaja <= '" & ddesde & "') "
            ' o F.Baja dentro del desde hasta
            sWhere &= " or ( fbaja between '" & ddesde & "' and '" & dHasta & "') )"
            Dt = AdminData.GetData(sSelect & sWhere)
            If Not IsNothing(Dt) Then
                If Dt.Rows.Count > 0 And Not IsDBNull(Dt.Rows(0)(0)) Then
                    idiasBaja = Dt.Rows(0)(0)
                Else
                    idiasBaja = 0
                End If
            End If
            ' Retornar resultado
            Return iacc & "-" & idiasBaja
        Catch ex As Exception
            ApplicationService.GenerateError("Error al obtener datos estadísticas 3 y 4." & ex.Message)
        End Try
    End Function

    ' Punto 5,6 
    Public Function AccidentesMil(ByVal idObra As Integer, ByVal pMes As Integer, ByVal pyear As Int16, Optional ByVal b20 As Boolean = True) As String
        Try
            Dim ddesde, dHasta As DateTime
            Dim daccObra, dAccEmpr, dTotalAcc, dDiasBaja As Decimal

            ' Mes de 20 a < 21,13/09/2012, O Natural 1 al < 1
            If b20 Then
                ddesde = CDate("21/" & pMes & "/" & pyear)
                ddesde = DateAdd(DateInterval.Month, -1, ddesde)
                dHasta = CDate("20/" & pMes & "/" & pyear)

            Else
                ddesde = CDate("01/" & pMes & "/" & pyear)
                dHasta = DateAdd(DateInterval.Month, 1, ddesde)
            End If


            Dim sSelect, sWhere, sOrder As String
            Dim Dt As DataTable
            ' Obtener Accidentes, 100000,13/09/2012
            ' hasta el 20 del mes
            sSelect = "SELECT  count(idOperario) As TotalAcc, COUNT(idOperario)*100000 /(select sum(horasRealMod) from tbobraModControl where (idobra = " & idObra & ") AND (tbObraMODControl.FechaInicio < '" & dHasta & "' )) As a100000 FROM   vFrmPrevAccidentesObra "
            sWhere = " where idobra = " & idObra & " and (fAccidente < '" & dHasta & "')"
            Dt = AdminData.GetData(sSelect & sWhere)
            If Not IsNothing(Dt) Then
                If Dt.Rows.Count > 0 And Not IsDBNull(Dt.Rows(0)("a100000")) Then
                    daccObra = Dt.Rows(0)("a100000")
                    daccObra = Decimal.Round(daccObra, 5)
                    'dTotalAcc = Nz(Dt.Rows(0)("TotalAcc"), 0)
                    dTotalAcc = iif(IsDBNull(Dt.Rows(0)("TotalAcc")), 0,Dt.Rows(0)("TotalAcc"))
                Else
                    daccObra = 0
                    dTotalAcc = 0
                End If
            End If
            ' Días de trabajo perdidos en el total de la obra 24/09/2012
            sSelect = "SELECT  sum(nDiasBaja) As snDiasBajas FROM vFrmPrevAccidentesObra "
            sWhere = " where idobra = " & idObra '' Total & " and (fAccidente < '" & dHasta & "')"
            Dt = AdminData.GetData(sSelect & sWhere)
            If Not IsNothing(Dt) Then
                If Dt.Rows.Count > 0 And Not IsDBNull(Dt.Rows(0)("snDiasBajas")) Then
                    dDiasBaja = IIf(IsDBNull(Dt.Rows(0)("snDiasBajas")), 0, Dt.Rows(0)("snDiasBajas"))
                Else
                    dDiasBaja = 0
                End If
            End If
            ' Obtener Nº días perdidos, obras activas,13/09/2012
            ' hasta el 20 del mes, 18/09/2012
            ' Porque hago la segunda div?? sSelect = "SELECT COUNT(idOperario)*100000 /(select sum(horasRealMod) from tbobraModControl)/(select count(distinct idobra) from tbobraModControl) As afull "
            ' Sólo sumo los operarios de obras activas 22/09/2012,comprobar idobra vista igual idobra tbobracabecera.
            sSelect = "SELECT COUNT(idOperario)*100000 /(select sum(horasRealMod) from tbobraModControl,tbObraCabecera  where tbObraMODControl.IDObra = tbObraCabecera.IDObra AND tbObraCabecera.FFinPrev IS NULL AND (tbObraMODControl.FechaInicio < '" & dHasta & "') ) As afull "
            sWhere = " FROM vFrmPrevAccidentesObra,tbObraCabecera " & " WHERE (fAccidente < '" & dHasta & "') and (vFrmPrevAccidentesObra.idObra = tbObraCabecera.idObra) AND (tbObraCabecera.FFinPrev is Null) "

            Dt = AdminData.GetData(sSelect & sWhere)
            If Not IsNothing(Dt) Then
                If Dt.Rows.Count > 0 And Not IsDBNull(Dt.Rows(0)("afull")) Then
                    dAccEmpr = Dt.Rows(0)("afull")
                    dAccEmpr = Decimal.Round(dAccEmpr, 5)
                Else
                    dAccEmpr = 0
                End If
            End If
            ' Retornar resultado
            Return daccObra & "-" & dAccEmpr & "-" & dTotalAcc & "-" & dDiasBaja
        Catch ex As Exception
            ApplicationService.GenerateError("Error al obtener datos estadísticas 5 y 6." & ex.Message)
        End Try
    End Function

    ' Punto 7,8
    Public Function DayLost(ByVal idObra As Integer, ByVal pMes As Integer, ByVal pyear As Int16, Optional ByVal b20 As Boolean = True) As String
        Try
            Dim ddesde, dHasta As DateTime
            Dim lostObra, LostEmpr As Decimal

            ' Mes de 20 a < 21,13/09/2012, O Natural 1 al < 1
            If b20 Then
                ddesde = CDate("21/" & pMes & "/" & pyear)
                ddesde = DateAdd(DateInterval.Month, -1, ddesde)
                dHasta = CDate("20/" & pMes & "/" & pyear)

            Else
                ddesde = CDate("01/" & pMes & "/" & pyear)
                dHasta = DateAdd(DateInterval.Month, 1, ddesde)
            End If

            Dim sSelect, sWhere, sOrder As String
            Dim Dt As DataTable
            ' Obtener Accidentes,
            ' Control hasta, 22/09/2012
            ' Sólo sumo los operarios de obras activas 22/09/2012,comprobar idobra vista igual idobra tbobracabecera.
            sSelect = "SELECT (sum(vFrmPrevAccidentesObra.nDiasBaja)/(select sum(horasrealmod) from tbobramodcontrol where (idobra= " & idObra & ") AND (tbObraMODControl.FechaInicio < '" & dHasta & "' ) )) * 1000 As LostObra FROM vFrmPrevAccidentesObra "
            sWhere = " where idobra = " & idObra & " and (fAccidente < '" & dHasta & "')"
            Dt = AdminData.GetData(sSelect & sWhere)
            If Not IsNothing(Dt) Then
                If Dt.Rows.Count > 0 And Not IsDBNull(Dt.Rows(0)(0)) Then
                    lostObra = Dt.Rows(0)(0)
                    lostObra = Decimal.Round(lostObra, 5)
                Else
                    lostObra = 0
                End If
            End If
            ' Obtener Nº días perdidos
            ' de las que no tienen fecha fin,13/09/2012
            ' Vuelvo a dividir por la formula pasada: SUM diasper obras act. / Nº obras activas * 1000
            sSelect = "SELECT ((sum(vFrmPrevAccidentesObra.nDiasBaja)/(select sum(horasrealmod) from tbobramodcontrol,tbObraCabecera  where tbObraMODControl.IDObra = tbObraCabecera.IDObra AND tbObraCabecera.FFinPrev IS NULL AND (tbObraMODControl.FechaInicio < '" & dHasta & "' ))) / (select count(distinct tbObraCabecera.idobra) from tbobramodcontrol,tbObraCabecera  where tbObraMODControl.IDObra = tbObraCabecera.IDObra AND tbObraCabecera.FFinPrev IS NULL AND (tbObraMODControl.FechaInicio < '" & dHasta & "' ) )) * 1000 As LostFull "
            sWhere = " FROM vFrmPrevAccidentesObra,tbObraCabecera " & " WHERE (fAccidente < '" & dHasta & "') and (vFrmPrevAccidentesObra.idObra = tbObraCabecera.idObra AND tbObraCabecera.FFinPrev is Null) "

            Dt = AdminData.GetData(sSelect & sWhere)
            If Not IsNothing(Dt) Then
                If Dt.Rows.Count > 0 And Not IsDBNull(Dt.Rows(0)(0)) Then
                    LostEmpr = Dt.Rows(0)(0)
                    LostEmpr = Decimal.Round(LostEmpr, 5)
                Else
                    LostEmpr = 0
                End If
            End If
            ' Retornar resultado
            Return lostObra & "-" & LostEmpr
        Catch ex As Exception
            ApplicationService.GenerateError("Error al obtener datos estadísticas 7 y 8." & ex.Message)
        End Try
    End Function

    ' Punto 11,12
    Public Function JornadasJob(ByVal idObra As Integer, ByVal pMes As Integer, ByVal pyear As Int16, Optional ByVal b20 As Boolean = True) As String
        Try
            Dim ddesde, dHasta As DateTime
            Dim JorMes, JorObra As Decimal

            ' Mes de 20 a < 21,13/09/2012, O Natural 1 al < 1
            If b20 Then
                ddesde = CDate("21/" & pMes & "/" & pyear)
                ddesde = DateAdd(DateInterval.Month, -1, ddesde)
                dHasta = CDate("20/" & pMes & "/" & pyear)

            Else
                ddesde = CDate("01/" & pMes & "/" & pyear)
                dHasta = DateAdd(DateInterval.Month, 1, ddesde)
            End If

            Dim sSelect, sWhere, sOrder As String
            Dim Dt As DataTable
            ' Obtener Jornadas del mes,
            ' Control hasta, 22/09/2012
            ' Sólo sumo los operarios de obras activas 22/09/2012,comprobar idobra vista igual idobra tbobracabecera.
            sSelect = "SELECT count(DISTINCT IDOperario + CONVERT(VARCHAR(8), FechaInicio, 112)) AS Jornada FROM tbObraMODControl "
            sWhere = " where idobra = " & idObra & " and (FechaInicio  >= '" & ddesde & "')" & " and (FechaInicio  <= '" & dHasta & "')"
            Dt = AdminData.GetData(sSelect & sWhere)
            If Not IsNothing(Dt) Then
                If Dt.Rows.Count > 0 And Not IsDBNull(Dt.Rows(0)(0)) Then
                    JorMes = Dt.Rows(0)(0)
                Else
                    JorMes = 0
                End If
            End If
            ' Obtener jornadas obra
            ' de las que no tienen fecha fin,13/09/2012
            sSelect = "SELECT count(DISTINCT IDOperario + CONVERT(VARCHAR(8), FechaInicio, 112)) AS Jornada FROM tbObraMODControl "
            sWhere = " where idobra = " & idObra

            Dt = AdminData.GetData(sSelect & sWhere)
            If Not IsNothing(Dt) Then
                If Dt.Rows.Count > 0 And Not IsDBNull(Dt.Rows(0)(0)) Then
                    JorObra = Dt.Rows(0)(0)
                Else
                    JorObra = 0
                End If
            End If
            ' Retornar resultado
            Return JorMes & "-" & JorObra
        Catch ex As Exception
            ApplicationService.GenerateError("Error al obtener datos estadísticas 11 y 12." & ex.Message)
        End Try
    End Function

    ' Empresa 1,2
    Public Function HorasNumTrabaE(ByVal pMes As Integer, ByVal pyear As Int16) As String
        Try
            Dim ddesde, dHasta As DateTime
            Dim dSumHoras, dNumOp As Decimal

            ddesde = CDate("01/" & pMes & "/" & pyear)
            dHasta = DateAdd(DateInterval.Month, 1, ddesde)


            Dim sSelect, sWhere, sOrder As String
            Dim Dt As DataTable
            ' 
            sSelect = "select sum(horasRealMod) As Horas,count(distinct idoperario) As NumOp from tbobraModControl "
            sWhere = " where fechaInicio >='" & ddesde & "'" & " and fechaInicio < '" & dHasta & "'"
            Dt = AdminData.GetData(sSelect & sWhere)
            If Not IsNothing(Dt) Then
                If Dt.Rows.Count > 0 And Not IsDBNull(Dt.Rows(0)(0)) Then
                    dSumHoras = Dt.Rows(0)(0)
                    If IsDBNull(Dt.Rows(0)(1)) Then
                        dNumOp = 0
                    Else
                        dNumOp = Dt.Rows(0)(1)
                    End If

                Else
                    dSumHoras = 0
                    dNumOp = 0
                End If
            End If
            ' Redondeo
            dSumHoras = Decimal.Round(dSumHoras, 5)
            dNumOp = Decimal.Round(dNumOp, 5)
            ' Retornar resultado
            Return dSumHoras & "-" & dNumOp
        Catch ex As Exception
            ApplicationService.GenerateError("Error al obtener datos estadísticas 1 y 2 Empresa." & ex.Message)
        End Try
    End Function

    ' Empresa 3
    Public Function MedTrabaEmp(ByVal pMes As Integer, ByVal pyear As Int16) As Decimal
        Try
            Dim ddesde, dHasta As DateTime
            Dim dActivos, dMediaBaja As Decimal

            ddesde = CDate("01/" & pMes & "/" & pyear)
            dHasta = DateAdd(DateInterval.Month, 1, ddesde)

            Dim sSelect, sWhere As String
            Dim Dt As DataTable
            ' Obtener Trabajadores Activos
            sSelect = "SELECT COUNT(1) AS TrabAct FROM  vMaestroOperarioCompleta "
            sWhere = " where (fecha_baja is null )"
            Dt = AdminData.GetData(sSelect & sWhere)
            If Not IsNothing(Dt) Then
                If Dt.Rows.Count > 0 And Not IsDBNull(Dt.Rows(0)(0)) Then
                    dActivos = Dt.Rows(0)(0)
                Else
                    dActivos = 0
                End If
            End If

            ' Trabajadores con baja mes
            sSelect = "SELECT IDOperario, DATEDIFF(dd, '" & ddesde & "', Fecha_Baja) AS DiasAct,fecha_baja FROM vMaestroOperarioCompleta"
            sWhere = " where (fecha_baja >='" & ddesde & "')" & " and (fecha_baja < '" & dHasta & "')"
            Dt = AdminData.GetData(sSelect & sWhere)

            If Not IsNothing(Dt) Then
                If Dt.Rows.Count > 0 Then
                    dMediaBaja = 0
                    For Each ifila As DataRow In Dt.Rows
                        dMediaBaja += ifila("DiasAct") / 30
                    Next
                End If
            Else
                dMediaBaja = 0
            End If
            ' Redondeo
            dMediaBaja = Decimal.Round(dMediaBaja, 5)
            ' Retornar resultado
            Return dActivos + dMediaBaja
        Catch ex As Exception
            ApplicationService.GenerateError("Error al obtener datos estadística 3 Empresa." & ex.Message)
        End Try
    End Function

    ' Empresa 4
    Public Function MediaHorasEmp(ByVal pMes As Integer, ByVal pyear As Int16) As Decimal
        Try
            Dim ddesde, dHasta As DateTime
            Dim idiastrab As Integer
            Dim dsumHoras As Decimal

            ddesde = CDate("01/" & pMes & "/" & pyear)
            dHasta = DateAdd(DateInterval.Month, 1, ddesde)

            Dim sSelect, sWhere, sOrder As String
            Dim Dt As DataTable
            ' Obtener Horas
            sSelect = "select sum(horasRealMod) As Horas from tbobraModControl"
            sWhere = " where fechaInicio >='" & ddesde & "'" & " and fechaInicio < '" & dHasta & "'"
            Dt = AdminData.GetData(sSelect & sWhere)
            If Not IsNothing(Dt) Then
                If Dt.Rows.Count > 0 And Not IsDBNull(Dt.Rows(0)(0)) Then
                    dsumHoras = Dt.Rows(0)(0)
                Else
                    ' No hay horas retornar media 0
                    Return 0
                End If
            End If
            ' Obtener Nº días trabajados por trabajador
            sSelect = "select idoperario,fechainicio from tbobraModControl"
            sWhere = " where fechaInicio >='" & ddesde & "'" & " and fechaInicio < '" & dHasta & "' group by idoperario,fechainicio"
            Dt = AdminData.GetData(sSelect & sWhere)
            If Not IsNothing(Dt) Then
                If Dt.Rows.Count > 0 Then
                    idiastrab = Dt.Rows.Count
                End If
            End If
            ' Retornar resultado
            Return Decimal.Round(dsumHoras / idiastrab, 5)
        Catch ex As Exception
            ApplicationService.GenerateError("Error al obtener la media de horas.Estadistica 4 empresa." & ex.Message)
        End Try
    End Function

    ' Empresa 5
    Public Function AccidentesMilEmpresa(ByVal pMes As Integer, ByVal pyear As Int16, Optional ByVal b20 As Boolean = True) As String
        Try
            Dim ddesde, dHasta As DateTime
            Dim daccEmp As Decimal

            ' Mes de 20 a < 21,13/09/2012, O Natural 1 al < 1
            If b20 Then
                ddesde = CDate("21/" & pMes & "/" & pyear)
                ddesde = DateAdd(DateInterval.Month, -1, ddesde)
                dHasta = CDate("20/" & pMes & "/" & pyear)

            Else
                ddesde = CDate("01/" & pMes & "/" & pyear)
                dHasta = DateAdd(DateInterval.Month, 1, ddesde)
            End If


            Dim sSelect, sWhere, sOrder As String
            Dim Dt As DataTable
            ' Obtener Accidentes, 100000,25/09/2012
            ' hasta el 20 del mes
            sSelect = "SELECT  COUNT(idOperario)*100000 /(select sum(horasRealMod) from tbobraModControl where (tbObraMODControl.FechaInicio < '" & dHasta & "' )) As a100000 FROM   vFrmPrevAccidentesObra "
            sWhere = " where (fAccidente < '" & dHasta & "')"
            Dt = AdminData.GetData(sSelect & sWhere)
            If Not IsNothing(Dt) Then
                If Dt.Rows.Count > 0 And Not IsDBNull(Dt.Rows(0)("a100000")) Then
                    daccEmp = Dt.Rows(0)("a100000")
                    daccEmp = Decimal.Round(daccEmp, 5)
                Else
                    daccEmp = 0
                End If
            End If
            ' Retornar resultado
            Return daccEmp
        Catch ex As Exception
            ApplicationService.GenerateError("Error al obtener datos estadísticas 5 y 6." & ex.Message)
        End Try
    End Function

    ' Empresa días perdidos
    Public Function DiasPerdidosEmp(ByRef dtDias As DataTable)
        Try
            Dim ddesde, dHasta, dcursor As DateTime

            Dim sSelect, sWhere, sfiltro As String
            Dim DtProc As DataTable
            Dim dFilaDias As DataRow()
            ' Reseteo días
            dtDias = AdminData.GetData("select * from vFrmPrevAccidentesDias")

            ' Obtener Horas
            sSelect = "select fbaja,falta from tbOperarioAccidentes "
            sWhere = " where (fbaja >= DATEADD([year], - 1, getDate()))"
            DtProc = AdminData.GetData(sSelect & sWhere)
            If Not IsNothing(DtProc) Then
                ' Recorrer las filas tratando datos
                For Each dfila As DataRow In DtProc.Rows
                    ddesde = dfila("fBaja")
                    dHasta = CDate(IIf(IsDBNull(dfila("fAlta")), Now, dfila("fAlta"))) ' Ojo cuenta los de fecha nulo el mes en curso
                    ' Suma Días
                    dcursor = ddesde
                    ' Mover cursor
                    dcursor = DateAdd(DateInterval.Month, 1, dcursor)
                    dcursor = "01/" & dcursor.Month & "/" & dcursor.Year
                    While dcursor < dHasta
                        ' Coger la fila del mes a sumar datos, siempre es la primera
                        sfiltro = "MesYear = '" & ddesde.Year & "/" & Format(ddesde.Month, "00") & "'"
                        dFilaDias = dtDias.Select(sfiltro)
                        If dFilaDias.Length > 0 Then
                            dFilaDias(0)("DiasBaja") += DateDiff(DateInterval.Day, ddesde, dcursor)
                        End If
                        '' Resetear el desde con cursor
                        ddesde = dcursor
                        ' Mover cursor
                        dcursor = DateAdd(DateInterval.Month, 1, dcursor)
                        dcursor = "01/" & dcursor.Month & "/" & dcursor.Year
                        ' Si está en el mes
                        If dcursor.Month >= dHasta.Month Or dcursor.Year >= dHasta.Year Then
                            ' Días última fila
                            dcursor = dHasta
                            sfiltro = "MesYear = '" & dcursor.Year & "/" & Format(dcursor.Month, "00") & "'"
                            dFilaDias = dtDias.Select(sfiltro)
                            If dFilaDias.Length > 0 Then
                                dFilaDias(0)("DiasBaja") += DateDiff(DateInterval.Day, ddesde, dHasta)
                            End If
                        End If
                    End While
                    ' Control mismo Mes
                    If dcursor > dHasta Then
                        ' Días última fila
                        dcursor = dHasta
                        sfiltro = "MesYear = '" & dcursor.Year & "/" & Format(dcursor.Month, "00") & "'"
                        dFilaDias = dtDias.Select(sfiltro)
                        If dFilaDias.Length > 0 Then
                            dFilaDias(0)("DiasBaja") += DateDiff(DateInterval.Day, ddesde, dHasta)
                        End If
                    End If
                Next

            End If
            ' Retornar resultado
            Return dtDias
        Catch ex As Exception
            ApplicationService.GenerateError("Error al obtener los días perdidos. Empresa 5." & ex.Message)
        End Try
    End Function


#End Region

End Class
