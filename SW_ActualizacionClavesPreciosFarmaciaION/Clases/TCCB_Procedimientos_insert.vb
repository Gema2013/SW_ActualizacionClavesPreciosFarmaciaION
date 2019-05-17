Imports System.Data
Imports System.Data.SqlClient

Namespace SP_BBDD

    Public Class TCCB_Procedimientos_insert

        Public Sub New()
            MyBase.New()
        End Sub

        Public Overridable Function TCCB_Procedimientos_insert(ByVal connection As System.Data.SqlClient.SqlConnection, ByVal transaccion As System.Data.SqlClient.SqlTransaction, ByVal table As System.Data.DataTable,
                                                        ByVal Clave As String,
                                                        ByVal Descripcion As String,
                                                        ByVal TasaIVA As Integer,
                                                        ByVal Precio As Double,
                                                        ByVal CodigoBarras As String,
                                                        ByVal Activo As Boolean)
            Dim RETURN_VALUE As Integer = 0
            Dim cmd As System.Data.SqlClient.SqlCommand = Nothing
            Dim reader As System.Data.SqlClient.SqlDataReader = Nothing
            If (connection Is Nothing) Then
                Throw New System.ArgumentException("El objeto conection no puede ser null")
            Else
                'Comentario
                If (transaccion Is Nothing) Then
                    If (connection.State = System.Data.ConnectionState.Closed) Then
                        connection.Open()
                        'Abrir Conexion
                    End If
                    cmd = New System.Data.SqlClient.SqlCommand("TCCB_Procedimientos_insert", connection)
                    cmd.CommandType = System.Data.CommandType.StoredProcedure
                Else
                    cmd = New System.Data.SqlClient.SqlCommand("TCCB_Procedimientos_insert", connection, transaccion)
                    cmd.CommandType = System.Data.CommandType.StoredProcedure
                End If
                cmd.Parameters.Add("@RETURN_VALUE", System.Data.SqlDbType.Int, 0)
                cmd.Parameters("@RETURN_VALUE").Direction = System.Data.ParameterDirection.ReturnValue
                cmd.Parameters("@RETURN_VALUE").Value = RETURN_VALUE


                cmd.Parameters.Add("@ClaveProcedimiento", System.Data.SqlDbType.VarChar)
                cmd.Parameters("@ClaveProcedimiento").Direction = System.Data.ParameterDirection.Input
                cmd.Parameters("@ClaveProcedimiento").Value = Clave
                cmd.Parameters.Add("@DescripcionProcedimiento", System.Data.SqlDbType.VarChar)
                cmd.Parameters("@DescripcionProcedimiento").Direction = System.Data.ParameterDirection.Input
                cmd.Parameters("@DescripcionProcedimiento").Value = Descripcion

                cmd.Parameters.Add("@TasaIVA", System.Data.SqlDbType.Int)
                cmd.Parameters("@TasaIVA").Direction = System.Data.ParameterDirection.Input
                cmd.Parameters("@TasaIVA").Value = TasaIVA

                cmd.Parameters.Add("@Precio", System.Data.SqlDbType.Decimal)
                cmd.Parameters("@Precio").Direction = System.Data.ParameterDirection.Input
                cmd.Parameters("@Precio").Value = Precio

                cmd.Parameters.Add("@CodigoBarras", System.Data.SqlDbType.VarChar)
                cmd.Parameters("@CodigoBarras").Direction = System.Data.ParameterDirection.Input
                cmd.Parameters("@CodigoBarras").Value = CodigoBarras

                cmd.Parameters.Add("@Activo", System.Data.SqlDbType.Bit)
                cmd.Parameters("@Activo").Direction = System.Data.ParameterDirection.Input
                cmd.Parameters("@Activo").Value = Activo


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
                If (transaccion Is Nothing) Then
                    connection.Close()
                End If
                RETURN_VALUE = CType(cmd.Parameters("@RETURN_VALUE").Value, Integer)
                Return RETURN_VALUE
            End If
        End Function
    End Class
End Namespace
