Imports System.Threading
Imports System.Data.SqlClient
Imports MySql.Data.MySqlClient


Public Class ServicioActualizacionION
    Const StrNombreAplicacion As String = "Actualiza farmacia-ION."

    Dim elogLogEventos As New EventLog

    Private Property TrdProcesaActualizacionFarmacia As Thread

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.

        IniciarServicio()
    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
    End Sub

#Region "Métodos"
    Private Sub IniciarServicio()

        Try
            If (Not EventLog.SourceExists(StrNombreAplicacion)) Then
                EventLog.CreateEventSource(StrNombreAplicacion, StrNombreAplicacion)
            End If

            elogLogEventos.Source = StrNombreAplicacion
            elogLogEventos.WriteEntry("Inicio del servicio ""Actualización de claves de farmacia"" V1", EventLogEntryType.Information)


            TrdProcesaActualizacionFarmacia = New Thread(AddressOf ActualizacionClavesFarmacia)
            TrdProcesaActualizacionFarmacia.IsBackground = True
            TrdProcesaActualizacionFarmacia.Start()

            elogLogEventos.WriteEntry("Servicio iniciado correctamente", EventLogEntryType.Information)
        Catch exExcepcion As Exception
            elogLogEventos.WriteEntry("No puedo iniciarse el servicio, " & exExcepcion.ToString, EventLogEntryType.FailureAudit)
            elogLogEventos.WriteEntry("Cierre inesperado", EventLogEntryType.FailureAudit)
            Me.Stop()
        End Try

    End Sub

    Private Sub ActualizacionClavesFarmacia()
        While True
            ProcesaActualizacionMedicamentosPendiente()
            'Thread.Sleep(60000) '10 min
            Thread.Sleep(6000) ' 1 min
        End While
    End Sub

    Private Sub ProcesaActualizacionMedicamentosPendiente()
        'Base en SION
        'Ambiente de pruebas
        Conexiones.Catalog = "BDSION"
        Conexiones.ID = "sa"
        Conexiones.Pwd = "s4_password"
        Conexiones.Source = "70.35.194.173"

        'Ambiente producción
        'Conexiones.Catalog = "BDSIONProduccion"



        'Base en Farmacia

        'Ambiente de pruebas
        Conexiones.CatalogMySql = "farmacia"
        Conexiones.SourceMySql = "70.35.194.173"
        Conexiones.PortMySql = 3306
        Conexiones.IDMySql = "root"
        Conexiones.PwdMySql = "s3rv3rS1nu3"

        'Ambiente producción
        'Conexiones.CatalogMySql = "farmaciaprod"



        Dim blnExito As Boolean = False
        Try
            elogLogEventos.WriteEntry("Iniciada la actualización por minuto de claves y usuarios de farmacia", EventLogEntryType.Information)
            ActualizacionClaves()
            elogLogEventos.WriteEntry("Actualización completa exitosa: " & Now.ToString, EventLogEntryType.Information)

        Catch ex As Exception
            elogLogEventos.WriteEntry("Error: " & ex.ToString, EventLogEntryType.Error)
        End Try

    End Sub

    Private Sub ActualizacionClaves()
        'Obtiene claves a procesar desde Mysql
        Dim conexion As MySqlConnection = Nothing
        Try
            conexion = Conexiones.ConexionMySQL
            conexion.Open()
            elogLogEventos.WriteEntry("Conectado al servidor MySql")
        Catch ex As MySqlException
            elogLogEventos.WriteEntry("No se ha podido conectar al servidor")
        Finally
            If conexion IsNot Nothing Then
                conexion.Close()
                conexion.Dispose()
            End If
        End Try


        Dim MiConexion As SqlConnection = Nothing
        Dim dtClavesNoProcesadas As New DataTable

        Try
            MiConexion = Conexiones.Conexion
            MiConexion.Open()

            'Acciones obtener en Base ion
            elogLogEventos.WriteEntry("Conectado al servidor Sql")
        Catch ex As Exception
            elogLogEventos.WriteEntry("Error al obtener usuarios que existen en la base de ION: " & ex.ToString, EventLogEntryType.Error)
        Finally
            If MiConexion IsNot Nothing Then
                MiConexion.Close()
                MiConexion.Dispose()
            End If
        End Try



        'Inserta claves y precios nuevos en BDSION
        Dim sqlconConexion As SqlConnection = Nothing
        Dim sqltranTransaccion As SqlTransaction = Nothing
        Try
            sqlconConexion = Conexiones.Conexion
            sqlconConexion.Open()
            sqltranTransaccion = sqlconConexion.BeginTransaction(IsolationLevel.ReadUncommitted)

            'Acciones insert, update en base ion

            sqltranTransaccion.Commit()
        Catch ex As Exception
            If sqltranTransaccion IsNot Nothing Then
                sqltranTransaccion.Rollback()
            End If
            elogLogEventos.WriteEntry("Insertar claves y actualizar precios en base ION. Error: " & ex.ToString, EventLogEntryType.Error)
        Finally
            If sqlconConexion IsNot Nothing Then
                sqlconConexion.Close()
                sqlconConexion.Dispose()
            End If
        End Try
    End Sub


    'Private Sub ActualizaUMPoblacionAbierta(ByVal dtClavesMedicamentosUM As DataTable, ByVal ExpedientePA As String, ByRef trans As SqlTransaction, ByRef conex As SqlConnection)
    '    If dtClavesMedicamentosUM.Rows.Count > 0 And ExpedientePA <> "" Then
    '        For Each row In dtClavesMedicamentosUM.Rows
    '            'Verificar precio
    '            Dim DatosSalida As CLConsultaPreciosSigaf.obtenerPrecioVentaMedicamentosDataTypeOut
    '            DatosSalida = ObtienePrecioSigaf(ExpedientePA, row!ClaveUM)

    '            Try
    '                Dim i As Integer = 1
    '                While i <= 7 'Niveles 1,2,3,4,5,6,7
    '                    Dim spActualizaPrecio As New SP_BBDD.ActualizaPreciosMedicamentos_UpdateInsert
    '                    spActualizaPrecio.ActualizaPreciosMedicamentos_UpdateInsert(conex, trans, Nothing, Redondear((DatosSalida.returnValue.expedientes(0).precios(0).precioVenta) / row!Factor), "UM" & row!ClaveUM, i)
    '                    i += 1
    '                End While
    '            Catch ex As Exception
    '                'volver a lanzar exception
    '                Throw (ex)
    '            End Try

    '        Next
    '    Else
    '        elogLogEventos.WriteEntry("Datos incompletos para realizar la actualización de claves UM. Población abierta")
    '    End If
    'End Sub

    'Private Function ObtienePrecioSigaf(ByVal Expediente As String, ByVal clave As String) As obtenerPrecioVentaMedicamentosDataTypeOut
    '    Dim ConsumoServicio As New CLConsultaPreciosSigaf.ConsultaPrecioService
    '    Dim edmEncabezadoDisponbilidad As New CLConsultaPreciosSigaf.ExpedienteProductosType
    '    'solicitar precio de cada clave
    '    edmEncabezadoDisponbilidad.expediente = Expediente
    '    edmEncabezadoDisponbilidad.productos = {clave}

    '    Dim odmtdiEntradaObtenerDisponibilidad As New obtenerPrecioVentaMedicamentosDataTypeIn
    '    odmtdiEntradaObtenerDisponibilidad.expedientes = {edmEncabezadoDisponbilidad}
    '    Return ConsumoServicio.obtenerPrecioVentaMedicamentos(odmtdiEntradaObtenerDisponibilidad)

    'End Function

    Private Function Redondear(ByVal Precio As String) As Decimal
        Dim rounded As Decimal = Decimal.Round(System.Convert.ToDecimal(Precio), 6)
        Return rounded
    End Function

#End Region

End Class
