Option Strict Off
Option Explicit On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text

Namespace SP_BBDD

    Public Class ObtieneClavesMedicamentosFM_Select

        Public Sub New()
            MyBase.New()
        End Sub

        Public Overridable Sub ObtieneClavesMedicamentosFM_Select(ByVal connection As System.Data.SqlClient.SqlConnection, _
               ByVal transaccion As System.Data.SqlClient.SqlTransaction, _
              ByVal MiDataTable As System.Data.DataTable)
            Dim cmd As System.Data.SqlClient.SqlCommand = Nothing
            If (connection Is Nothing) Then
                Throw New System.ArgumentException("El objeto conection no puede ser null")
            Else
                'Comentario
                If (transaccion Is Nothing) Then
                    If (connection.State = System.Data.ConnectionState.Closed) Then
                        connection.Open()
                        'Abrir Conexion
                    End If
                    cmd = New System.Data.SqlClient.SqlCommand("ObtieneClavesMedicamentosFM_Select", connection)
                    cmd.CommandType = System.Data.CommandType.StoredProcedure
                Else
                    cmd = New System.Data.SqlClient.SqlCommand("ObtieneClavesMedicamentosFM_Select", connection, transaccion)
                    cmd.CommandType = System.Data.CommandType.StoredProcedure
                End If


                If (Not (MiDataTable) Is Nothing) Then
                    Dim dr As System.Data.SqlClient.SqlDataReader
                    dr = cmd.ExecuteReader
                    MiDataTable.Load(dr)
                Else
                    cmd.ExecuteScalar()
                End If
                If (transaccion Is Nothing) Then
                    connection.Close()
                End If
            End If
        End Sub
    End Class
End Namespace
