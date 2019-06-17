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
            elogLogEventos.WriteEntry("Inicio del servicio ""Actualización de claves y usuarios de farmacia"" V2 Pruebas.", EventLogEntryType.Information)


            TrdProcesaActualizacionFarmacia = New Thread(AddressOf ActualizacionClavesFarmacia)
            TrdProcesaActualizacionFarmacia.IsBackground = True
            TrdProcesaActualizacionFarmacia.Start()

            elogLogEventos.WriteEntry("Servicio iniciado correctamente", EventLogEntryType.Information)

            'Para escribir en el registro de Aplicación de windows.
            EventLog.WriteEntry("Se inicia servicio de act de ION-Farmacia", EventLogEntryType.FailureAudit)

        Catch exExcepcion As Exception
            EventLog.WriteEntry("No puedo iniciarse el servicio, " & exExcepcion.ToString, EventLogEntryType.FailureAudit)
            EventLog.WriteEntry("Cierre inesperado", EventLogEntryType.FailureAudit)
            Me.Stop()
        End Try

    End Sub

    Private Sub ActualizacionClavesFarmacia()
        While True
            ProcesaActualizacionMedicamentosPendiente()
            'Thread.Sleep(60000) '10 min pruebas
            Thread.Sleep(30000) '5 min 
            'Thread.Sleep(6000) ' 1 min
        End While
    End Sub

    Private Sub ProcesaActualizacionMedicamentosPendiente()
        'Base en SION
        'Ambiente de pruebas
        'Conexiones.Catalog = "BDSION"
        Conexiones.ID = "sa"
        Conexiones.Pwd = "s4_password"
        Conexiones.Source = "70.35.194.173"

        'Ambiente producción
        Conexiones.Catalog = "BDSIONProduccion"



        'Base en Farmacia

        'Ambiente de pruebas
        'Conexiones.CatalogMySql = "farmacia"
        Conexiones.SourceMySql = "70.35.194.173"
        Conexiones.PortMySql = 3306
        Conexiones.IDMySql = "root"
        Conexiones.PwdMySql = "s3rv3rS1nu3"

        'Ambiente producción
        Conexiones.CatalogMySql = "farmaciaprod"



        Dim blnExito As Boolean = False
        Try
            'Las nuevas claves de farmacia (Procesado=0) con sus precios se deben guardar en la base ion CALL usp_catalogo_obtenerProductosNuevos();
            'y actualizar estatus de Procesado=1
            elogLogEventos.WriteEntry("Iniciada la actualización por minuto de claves", EventLogEntryType.Information)
            ActualizacionClaves()
            'Se procesan los usuarios que tengan 
            elogLogEventos.WriteEntry("Iniciada la actualización por minuto de usuarios de farmacia", EventLogEntryType.Information)
            ActualizacionUsuariosFarmacia()

            elogLogEventos.WriteEntry("Actualización completa exitosa: " & Now.ToString, EventLogEntryType.Information)

        Catch ex As Exception
            elogLogEventos.WriteEntry("Error: " & ex.ToString, EventLogEntryType.Error)
        End Try

    End Sub

    Private Sub ActualizacionClaves()
        'Obtiene claves a procesar desde Mysql
        Dim dt As New DataTable
        Dim conexion As MySqlConnection = Nothing
        Try
            conexion = Conexiones.ConexionMySQL
            conexion.Open()
            elogLogEventos.WriteEntry("Conectado al servidor MySql para obtener claves....")
            Dim sp As New SP_BBDD.ObtenerClavesNuevas_select
            sp.ObtenerClavesNuevas_select(conexion, Nothing, dt)

            elogLogEventos.WriteEntry("Obtuve :" & dt.Rows.Count & " claves sin procesar")


        Catch ex As MySqlException
            elogLogEventos.WriteEntry("Error obteniendo claves nuevas para procesar")
            Throw ex
        Finally
            If conexion IsNot Nothing Then
                conexion.Close()
                conexion.Dispose()
            End If
        End Try

        If dt.Rows.Count > 0 Then
            Dim MiConexion As SqlConnection = Nothing
            Dim tran As SqlTransaction = Nothing
            Dim dtClavesNoProcesadas As New DataTable
            Dim strXmlClaves As String = "<data><items>"
            Try
                MiConexion = Conexiones.Conexion
                MiConexion.Open()
                tran = MiConexion.BeginTransaction(IsolationLevel.ReadUncommitted)

                'Inserta nuevas claves en BD de ion
                Dim spInserta As New SP_BBDD.TCCB_Procedimientos_insert
                Dim i As Integer = 0
                For Each row In dt.Rows
                    spInserta.TCCB_Procedimientos_insert(MiConexion, tran, Nothing, row!codigo, row!nombre, row!impuesto_fk, row!precio_venta, row!codigo_barras, row!activo)
                    strXmlClaves = strXmlClaves & "<row catalogo_k = """ & row!catalogo_k & """/>"
                    'El exit es para realizar pruebas 1 x1
                    'Exit For
                    i = i + 1
                    If i = 50 Then
                        Exit For
                    End If
                Next
                strXmlClaves = strXmlClaves & "</items></data>"

                tran.Commit()
                elogLogEventos.WriteEntry("Se insertaron las claves" & strXmlClaves)
            Catch ex As Exception
                If tran IsNot Nothing Then
                    tran.Rollback()
                End If
                elogLogEventos.WriteEntry("Error insertando claves nuevas. Error: " & ex.ToString, EventLogEntryType.Error)
                Throw ex
            Finally
                If MiConexion IsNot Nothing Then
                    MiConexion.Close()
                    MiConexion.Dispose()
                End If
            End Try



            'Actualizar claves procesadas en Farmacia
            Dim MyCon As MySqlConnection = Nothing
            Dim success As Boolean = False

            Try
                MyCon = Conexiones.ConexionMySQL
                MyCon.Open()

                'Inserta nuevas claves en BD de ion
                Dim spAct As New SP_BBDD.ActualizaPreciosMedicamentos_UpdateInsert
                spAct.ActualizaPreciosMedicamentos_UpdateInsert(MyCon, Nothing, Nothing, strXmlClaves, success)
                If success Then
                    elogLogEventos.WriteEntry("Se actualizaron estatus de claves procesadas." & strXmlClaves)
                Else
                    elogLogEventos.WriteEntry("No se procesaron correctamente las claves: " & strXmlClaves)
                End If

            Catch ex As Exception
                elogLogEventos.WriteEntry("Error actualizando claves procesadas. Error: " & ex.ToString, EventLogEntryType.Error)
                Throw ex
            Finally
                If MyCon IsNot Nothing Then
                    MyCon.Close()
                    MyCon.Dispose()
                End If
            End Try
        Else
            elogLogEventos.WriteEntry("No hay claves para procesar.")
        End If
    End Sub

    Private Sub ActualizacionUsuariosFarmacia()
        'Obtiene claves de usuario desde SQL
        Dim SqlConexion As SqlConnection = Nothing
        Dim SqlTran As SqlTransaction = Nothing
        Dim dtClavesUsNoProcesadas As New DataTable
        Try
            SqlConexion = Conexiones.Conexion
            SqlConexion.Open()
            SqlTran = SqlConexion.BeginTransaction(IsolationLevel.ReadUncommitted)

            'Obtiene las claves de usuario no procesadas
            Dim spClavesUs As New SP_BBDD.ObtienePendientesProcesaFarm_select
            spClavesUs.ObtienePendientesProcesaFarm_select(SqlConexion, SqlTran, dtClavesUsNoProcesadas)

            elogLogEventos.WriteEntry("Obtuve " & dtClavesUsNoProcesadas.Rows.Count & "usuarios para procesar")

        Catch ex As Exception
            elogLogEventos.WriteEntry("Error obteniendo claves de usuario no procesadas. Error: " & ex.ToString, EventLogEntryType.Error)
            Throw ex
        Finally
            If SqlConexion IsNot Nothing Then
                SqlConexion.Close()
                SqlConexion.Dispose()
            End If
        End Try

        If dtClavesUsNoProcesadas.Rows.Count > 0 Then
            'Enviar claves de usuario a insertar en mysql
            Dim MySqlCon As MySqlConnection = Nothing
            Dim strXmlClavesUs As String = "<data><items>"
            Dim successUs As Boolean = False
            Dim usuariosNoCreados As String = False

            Try
                MySqlCon = Conexiones.ConexionMySQL
                MySqlCon.Open()
                elogLogEventos.WriteEntry("Conectado al servidor MySql para enviar claves de usuario a crear....")
                Dim spCreaUsuario As New SP_BBDD.usp_users_insert()


                Dim i As Integer = 0
                For Each row In dtClavesUsNoProcesadas.Rows
                    Dim Psw As String = row!claveUsuario + "_ion381"
                    Dim PswEncriptado = CodeShare.Cryptography.Encriptacion.GenerateSHA256String(Psw)
                    strXmlClavesUs = strXmlClavesUs & "<row curp = """ & row!claveUsuario & """"
                    strXmlClavesUs = strXmlClavesUs & " nombres = """ & row!nombre & """"
                    strXmlClavesUs = strXmlClavesUs & " apellido_paterno = """ & row!aPaterno & """"
                    strXmlClavesUs = strXmlClavesUs & " apellido_materno = """ & row!aMaterno & """"
                    strXmlClavesUs = strXmlClavesUs & " telefono_casa = """""
                    strXmlClavesUs = strXmlClavesUs & " telefono_celular = """""
                    strXmlClavesUs = strXmlClavesUs & " telefono_oficina = """""
                    strXmlClavesUs = strXmlClavesUs & " direccion = """""
                    strXmlClavesUs = strXmlClavesUs & " login = """ & row!claveUsuario & """"
                    strXmlClavesUs = strXmlClavesUs & " contrasenya = """ & PswEncriptado & """"
                    strXmlClavesUs = strXmlClavesUs & " almacen_fk = """ & 1 & """"
                    strXmlClavesUs = strXmlClavesUs & " rol_fk = """ & 1 & """/>"
                Next

                strXmlClavesUs = strXmlClavesUs & "</items></data>"

                spCreaUsuario.usp_users_insert(MySqlCon, Nothing, Nothing, strXmlClavesUs, successUs, usuariosNoCreados)

                elogLogEventos.WriteEntry("Se procesaron los usuarios faltantes" & strXmlClavesUs & " ya existen en farmacia los usuarios: " & usuariosNoCreados)
            Catch ex As MySqlException
                elogLogEventos.WriteEntry("Error procesando nuevos usuarios" & strXmlClavesUs)
                Throw ex
            Finally
                If MySqlCon IsNot Nothing Then
                    MySqlCon.Close()
                    MySqlCon.Dispose()
                End If
            End Try





            If successUs Then
                'Actualizar claves de usuario procesados en SQL
                SqlConexion = Nothing
                SqlTran = Nothing
                Dim dtActualizarClaves As New DataTable

                Try
                    SqlConexion = Conexiones.Conexion
                    SqlConexion.Open()
                    SqlTran = SqlConexion.BeginTransaction(IsolationLevel.ReadUncommitted)

                    'Inserta nuevas claves en BD de ion
                    Dim spActualiza As New SP_BBDD.ActualizaPendientesProcesaFarm_update
                    Dim i As Integer = 0
                    For Each row In dtClavesUsNoProcesadas.Rows
                        spActualiza.ActualizaPendientesProcesaFarm_update(SqlConexion, SqlTran, Nothing, row!claveUsuario)

                    Next

                    SqlTran.Commit()
                    elogLogEventos.WriteEntry("Se actualizó estatus de usuarios en base ION" & strXmlClavesUs)
                Catch ex As Exception
                    If SqlTran IsNot Nothing Then
                        SqlTran.Rollback()
                    End If
                    elogLogEventos.WriteEntry("Error actualizando estatus de claves de usuario procesadas. Error: " & ex.ToString, EventLogEntryType.Error)
                    Throw ex
                Finally
                    If SqlConexion IsNot Nothing Then
                        SqlConexion.Close()
                        SqlConexion.Dispose()
                    End If
                End Try
            Else
                elogLogEventos.WriteEntry("No se realizó la actualización de claves de usuario en ION." & strXmlClavesUs & " Sucess:" & successUs)
            End If
        Else
            elogLogEventos.WriteEntry("No hay claves de usuario pendientes para procesar")
        End If
    End Sub



#End Region

End Class
