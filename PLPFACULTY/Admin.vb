Imports MySql.Data.MySqlClient

' Module for admin-related DB operations
Module Admin

    Public CurrentAdmin As AdminModel

    ' Insert an admin into the database
    Public Function InsertAdmin(firstName As String, lastName As String,
                            middleInitial As String, suffix As String,
                            email As String, username As String, password As String) As Boolean
        Try
            Connect()

            Dim query As String = "INSERT INTO admins (first_name, last_name, middle_initial, suffix, email, username, password) " &
                              "VALUES (@first_name, @last_name, @middle_initial, @suffix, @email, @username, @password)"

            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@first_name", firstName)
                cmd.Parameters.AddWithValue("@last_name", lastName)
                cmd.Parameters.AddWithValue("@middle_initial", middleInitial)
                cmd.Parameters.AddWithValue("@suffix", suffix)
                cmd.Parameters.AddWithValue("@email", email)
                cmd.Parameters.AddWithValue("@username", username)
                cmd.Parameters.AddWithValue("@password", password)

                Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                Return rowsAffected > 0
            End Using

        Catch ex As Exception
            MessageBox.Show("Insert failed: " & ex.Message, "DB Error")
            Return False
        End Try
    End Function

    ' Update an admin in the database
    Public Function UpdateAdmin(id As Integer, firstName As String, lastName As String,
                            middleInitial As String, suffix As String,
                            email As String, username As String, password As String) As Boolean
        Try
            Connect()

            Dim query As String = "UPDATE admins SET " &
                              "first_name = @first_name, " &
                              "last_name = @last_name, " &
                              "middle_initial = @middle_initial, " &
                              "suffix = @suffix, " &
                              "email = @email, " &
                              "username = @username, " &
                              "password = @password " &
                              "WHERE id = @id"

            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@first_name", firstName)
                cmd.Parameters.AddWithValue("@last_name", lastName)
                cmd.Parameters.AddWithValue("@middle_initial", middleInitial)
                cmd.Parameters.AddWithValue("@suffix", suffix)
                cmd.Parameters.AddWithValue("@email", email)
                cmd.Parameters.AddWithValue("@username", username)
                cmd.Parameters.AddWithValue("@password", password)
                cmd.Parameters.AddWithValue("@id", id)

                Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                Return rowsAffected > 0
            End Using

        Catch ex As Exception
            MessageBox.Show("Update failed: " & ex.Message, "DB Error")
            Return False
        End Try
    End Function
    Public Function DeleteAdmin(id As Integer) As Boolean
        Try
            Connect()

            Dim query As String = "DELETE FROM admins WHERE id = @id"

            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@id", id)

                Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                Return rowsAffected > 0
            End Using

        Catch ex As Exception
            MessageBox.Show("Delete failed: " & ex.Message, "DB Error")
            Return False
        End Try
    End Function


    ' Login using username and password
    Public Function LoginAdmin(username As String, password As String) As AdminModel
        Try
            Connect()

            Dim query As String = "SELECT * FROM admins WHERE username = @username AND password = @password"

            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@username", username)
                cmd.Parameters.AddWithValue("@password", password)

                Using reader As MySqlDataReader = cmd.ExecuteReader()
                    If reader.Read() Then
                        ' Create and return the AdminModel object
                        Dim admin As New AdminModel With {
                            .Id = reader("id"),
                            .FirstName = reader("first_name").ToString(),
                            .LastName = reader("last_name").ToString(),
                            .MiddleInitial = reader("middle_initial").ToString(),
                            .Suffix = reader("suffix").ToString(),
                            .Email = reader("email").ToString(),
                            .Username = reader("username").ToString()
                        }
                        Return admin
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Login failed: " & ex.Message, "Login Error")
        End Try

        Return Nothing
    End Function
    ' Retrieve a list of admins, optionally filtered by surname
    Public Function GetAdminsByLastName(Optional lastNameFilter As String = "") As List(Of AdminModel)
        Dim adminList As New List(Of AdminModel)
        Try
            Connect()

            Dim query As String = "SELECT * FROM admins"
            If Not String.IsNullOrWhiteSpace(lastNameFilter) Then
                query &= " WHERE last_name LIKE @filter"
            End If

            Using cmd As New MySqlCommand(query, conn)
                If Not String.IsNullOrWhiteSpace(lastNameFilter) Then
                    cmd.Parameters.AddWithValue("@filter", lastNameFilter & "%")
                End If

                Using reader As MySqlDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim admin As New AdminModel With {
                        .Id = reader("id"),
                        .FirstName = reader("first_name").ToString(),
                        .LastName = reader("last_name").ToString(),
                        .MiddleInitial = reader("middle_initial").ToString(),
                        .Suffix = reader("suffix").ToString(),
                        .Email = reader("email").ToString(),
                        .Username = reader("username").ToString(),
                        .Password = reader("password").ToString()
                    }
                        adminList.Add(admin)
                    End While
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Failed to fetch admins: " & ex.Message, "Database Error")
        End Try

        Return adminList
    End Function

    Public Function GetAdminById(adminId As Integer) As DataRow
        Try
            Connect()
            Dim query As String = "SELECT * FROM admins WHERE id = @id"
            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@id", adminId)
                Using adapter As New MySqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    adapter.Fill(dt)
                    If dt.Rows.Count > 0 Then
                        Return dt.Rows(0)
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Failed to fetch admin data: " & ex.Message, "DB Error")
        Finally
            Disconnect()
        End Try
        Return Nothing
    End Function
    Public Function GetAdminByEmail(email As String) As AdminModel
        Dim query As String = $"SELECT * FROM admins WHERE email = '{MySqlHelper.EscapeString(email)}' LIMIT 1"
        Dim dt As DataTable = ExecuteQuery(query)

        If dt.Rows.Count = 0 Then
            Return Nothing
        End If

        Dim row As DataRow = dt.Rows(0)
        Dim admin As New AdminModel With {
        .Id = Convert.ToInt32(row("id")),
        .FirstName = row("first_name").ToString(),
        .LastName = row("last_name").ToString(),
        .MiddleInitial = row("middle_initial").ToString(),
        .Suffix = row("suffix").ToString(),
        .Email = row("email").ToString(),
        .Username = row("username").ToString(),
        .Password = row("password").ToString()
    }

        Return admin
    End Function


End Module
' Model class representing an admin
Public Class AdminModel
    Public Property Id As Integer
    Public Property FirstName As String
    Public Property LastName As String
    Public Property MiddleInitial As String
    Public Property Suffix As String
    Public Property Email As String
    Public Property Username As String

    Public Property Password As String
End Class
