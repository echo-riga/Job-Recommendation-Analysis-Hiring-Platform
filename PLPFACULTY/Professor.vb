Imports MySql.Data.MySqlClient

' Module for professor-related DB operations
Module Professor

    Public CurrentProfessor As ProfessorModel

    ' Insert a professor into the database
    Public Function InsertProfessor(lastName As String, firstName As String,
                                    middleInitial As String, suffix As String,
                                    email As String, username As String, password As String) As Boolean
        Try
            Connect()

            Dim query As String = "INSERT INTO professors (first_name, last_name, middle_initial, suffix, email, username, password) " &
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

    ' Login using username and password
    Public Function LoginProfessor(username As String, password As String) As ProfessorModel
        Try
            Connect()

            Dim query As String = "SELECT * FROM professors WHERE username = @username AND password = @password"

            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@username", username)
                cmd.Parameters.AddWithValue("@password", password)

                Using reader As MySqlDataReader = cmd.ExecuteReader()
                    If reader.Read() Then
                        ' Create and return the ProfessorModel object
                        Dim prof As New ProfessorModel With {
                            .Id = reader("id"),
                            .FirstName = reader("first_name").ToString(),
                            .LastName = reader("last_name").ToString(),
                            .MiddleInitial = reader("middle_initial").ToString(),
                            .Suffix = reader("suffix").ToString(),
                            .Email = reader("email").ToString(),
                            .Username = reader("username").ToString()
                        }
                        Return prof
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Login failed: " & ex.Message, "Login Error")
        End Try

        Return Nothing
    End Function

    ' Returns all professors as a DataTable
    ' Returns all professors as a DataTable using shared DbModule
    Public Function GetAllProfessors() As DataTable
        Dim query As String = "SELECT * FROM professors WHERE isHidden = 0 ORDER BY last_name ASC, first_name ASC"
        Return ExecuteQuery(query)
    End Function


    Public Function DeleteProfessor(professorId As Integer) As Boolean
        Try
            Connect()

            Dim query As String = "DELETE FROM professors WHERE id = @id"

            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@id", professorId)

                Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                Return rowsAffected > 0
            End Using

        Catch ex As Exception
            MessageBox.Show("Delete failed: " & ex.Message, "DB Error")
            Return False
        End Try
    End Function


    Public Function UpdateProfessor(professorId As Integer, lastName As String, firstName As String,
                                middleInitial As String, suffix As String,
                                email As String, username As String, password As String) As Boolean
        Try
            Connect()

            Dim query As String = "UPDATE professors SET " &
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
                cmd.Parameters.AddWithValue("@id", professorId) ' 🔁 fixed name

                Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                Return rowsAffected > 0
            End Using

        Catch ex As Exception
            MessageBox.Show("Update failed: " & ex.Message, "DB Error")
            Return False
        End Try
    End Function
    Public Function GetFilteredProfessors(surname As String) As DataTable
        Try
            Connect()

            Dim query As String
            Dim cmd As New MySqlCommand()

            If String.IsNullOrWhiteSpace(surname) Then
                ' No filter - return all non-hidden professors
                query = "SELECT * FROM professors WHERE isHidden = 0"
                cmd = New MySqlCommand(query, conn)
            Else
                ' Filter by surname and only non-hidden professors
                query = "SELECT * FROM professors WHERE last_name LIKE @surname AND isHidden = 0"
                cmd = New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@surname", "%" & surname & "%")
            End If

            Dim dt As New DataTable()
            Using adapter As New MySqlDataAdapter(cmd)
                adapter.Fill(dt)
            End Using

            Return dt

        Catch ex As Exception
            MessageBox.Show("Failed to retrieve professors: " & ex.Message, "Query Error")
            Return New DataTable()
        End Try
    End Function
    Public Function GetProfessorByEmail(email As String) As ProfessorModel
        Dim query As String = $"SELECT * FROM professors WHERE email = @Email LIMIT 1"
        Dim prof As ProfessorModel = Nothing

        Try
            Connect()
            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@Email", email)
                Using reader As MySqlDataReader = cmd.ExecuteReader()
                    If reader.Read() Then
                        prof = New ProfessorModel() With {
                        .Id = Convert.ToInt32(reader("id")),
                        .FirstName = reader("first_name").ToString(),
                        .LastName = reader("last_name").ToString(),
                        .MiddleInitial = reader("middle_initial").ToString(),
                        .Suffix = reader("suffix").ToString(),
                        .Email = reader("email").ToString(),
                        .Username = reader("username").ToString(),
                        .Password = reader("password").ToString()
                    }
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error retrieving professor: " & ex.Message, "DB Error")
        End Try

        Return prof
    End Function
    Public Function GetProfessorById(professorId As Integer) As DataRow
        Try
            Connect()
            Dim query As String = "SELECT * FROM professors WHERE id = @id"
            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@id", professorId)
                Using adapter As New MySqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    adapter.Fill(dt)
                    If dt.Rows.Count > 0 Then
                        Return dt.Rows(0)
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Failed to fetch professor: " & ex.Message, "Database Error")
        End Try
        Return Nothing
    End Function

End Module

' Model class representing a professor
Public Class ProfessorModel
    Public Property Id As Integer
    Public Property FirstName As String
    Public Property LastName As String
    Public Property MiddleInitial As String
    Public Property Suffix As String
    Public Property Email As String
    Public Property Username As String
    Public Property Password As String
End Class
