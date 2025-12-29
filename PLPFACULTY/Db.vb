' DbModule.vb - Add these methods
Imports MySql.Data.MySqlClient

Module DbModule
    Public conn As MySqlConnection
    Public connectionString As String = "Server=localhost;Database=plp_db;Uid=root;Pwd=;"
    Public currentTransaction As MySqlTransaction = Nothing

    ' Open the connection if not already open
    Public Sub Connect()
        If conn Is Nothing Then
            conn = New MySqlConnection(connectionString)
        End If

        If conn.State <> ConnectionState.Open Then
            Try
                conn.Open()
            Catch ex As Exception
                MessageBox.Show("Database connection failed: " & ex.Message, "DB Error")
            End Try
        End If
    End Sub

    ' Close the connection if needed
    Public Sub Disconnect()
        If conn IsNot Nothing AndAlso conn.State = ConnectionState.Open Then
            ' Rollback any active transaction
            If currentTransaction IsNot Nothing Then
                Try
                    currentTransaction.Rollback()
                Catch ex As Exception
                    ' Ignore rollback errors during disconnect
                End Try
                currentTransaction = Nothing
            End If
            conn.Close()
        End If
    End Sub

    ' Execute SELECT queries and return results
    Public Function ExecuteQuery(query As String) As DataTable
        Dim dt As New DataTable()

        Try
            Connect()
            Dim da As New MySqlDataAdapter(query, conn)
            da.Fill(dt)
        Catch ex As Exception
            MessageBox.Show("Query failed: " & ex.Message, "Query Error")
        End Try

        Return dt
    End Function

    ' Execute INSERT, UPDATE, DELETE
    Public Sub ExecuteNonQuery(query As String)
        Try
            Connect()
            Dim cmd As New MySqlCommand(query, conn)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show("Action failed: " & ex.Message, "Action Error")
        End Try
    End Sub

    ' Begin a transaction
    Public Sub BeginTransaction()
        Connect()
        If currentTransaction Is Nothing Then
            currentTransaction = conn.BeginTransaction()
        Else
            Throw New InvalidOperationException("A transaction is already active")
        End If
    End Sub

    ' Commit the current transaction
    Public Sub CommitTransaction()
        If currentTransaction IsNot Nothing Then
            currentTransaction.Commit()
            currentTransaction = Nothing
        Else
            Throw New InvalidOperationException("No active transaction to commit")
        End If
    End Sub

    ' Rollback the current transaction
    Public Sub RollbackTransaction()
        If currentTransaction IsNot Nothing Then
            currentTransaction.Rollback()
            currentTransaction = Nothing
        Else
            Throw New InvalidOperationException("No active transaction to rollback")
        End If
    End Sub

    ' Execute non-query with parameters (for use with transactions)
    Public Sub ExecuteNonQueryWithParams(query As String, ParamArray parameters() As MySqlParameter)
        Try
            Connect()
            Dim cmd As New MySqlCommand(query, conn)
            If currentTransaction IsNot Nothing Then
                cmd.Transaction = currentTransaction
            End If

            cmd.Parameters.AddRange(parameters)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show("Action failed: " & ex.Message, "Action Error")
            Throw ' Re-throw to allow handling in calling code
        End Try
    End Sub
End Module