Imports System.Data
Imports MySql.Data.MySqlClient

Namespace SP_BBDD

    Public Class usp_users_insert

        Public Sub New()
            MyBase.New()
        End Sub

        Public Overridable Sub usp_users_insert(ByVal connection As MySqlConnection,
                                                                              ByVal transaccion As MySqlTransaction,
                                                                              ByVal table As DataTable,
                                                                            ByVal _xmlInput As String,
                                                                            ByRef _success As Boolean
                                                                            )
            Dim cmd As MySqlCommand = Nothing
            Dim reader As MySqlDataReader = Nothing
            If (connection Is Nothing) Then
                Throw New ArgumentException("El objeto conection no puede ser null")
            Else
                'Comentario
                If (transaccion Is Nothing) Then
                    If (connection.State = System.Data.ConnectionState.Closed) Then
                        connection.Open()
                        'Abrir Conexion
                    End If
                    cmd = New MySqlCommand("usp_users_insert", connection)
                    cmd.CommandType = CommandType.StoredProcedure
                Else
                    cmd = New MySqlCommand("usp_users_insert", connection, transaccion)
                    cmd.CommandType = CommandType.StoredProcedure
                End If

                Dim p As New MySqlParameter("?_xmlInput", MySql.Data.MySqlClient.MySqlDbType.VarChar)
                p.Value = _xmlInput
                p.Direction = ParameterDirection.Input
                cmd.Parameters.Add(p)

                Dim q As New MySqlParameter("?_success", MySql.Data.MySqlClient.MySqlDbType.Bit)
                q.Value = _success
                q.Direction = ParameterDirection.Output
                cmd.Parameters.Add(q)


                If (Not (table) Is Nothing) Then
                    reader = cmd.ExecuteReader
                Else
                    cmd.ExecuteNonQuery()
                End If
                If ((Not (table) Is Nothing) _
                            AndAlso (Not (reader) Is Nothing)) Then
                    table.Clear()
                    table.Columns.Clear()
                    Dim i As Integer = 0
                    Do While (i < reader.FieldCount)
                        Dim __type As System.Type
                        Dim __name As String
                        __type = reader.GetFieldType(i)
                        __name = reader.GetName(i)
                        table.Columns.Add(__name, __type)
                        i = (i + 1)
                    Loop

                    Do While reader.Read
                        Dim row As System.Data.DataRow = table.NewRow
                        Dim rowdata((reader.FieldCount) - 1) As Object
                        reader.GetValues(rowdata)
                        row.ItemArray = rowdata
                        table.Rows.Add(row)

                    Loop
                    reader.Close()
                End If
                'Cachar el valor de success
                _success = CType(q.Value, Boolean)
                If (transaccion Is Nothing) Then
                    connection.Close()
                End If

            End If
        End Sub
    End Class
End Namespace
