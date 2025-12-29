Imports MySql.Data.MySqlClient

' Module for student-related DB operations
Module Student

    Public CurrentStudent As StudentModel

    ' Insert a student into the database
    ' -----------------------------
    ' InsertStudent Function
    ' -----------------------------
    Public Function InsertStudent(studentNumber As String, firstName As String, lastName As String,
                              middleInitial As String, suffix As String, email As String,
                              sectionId As Integer, status As String) As Boolean
        Try
            Connect()

            Dim query As String = "INSERT INTO students (student_number, first_name, last_name, middle_initial, suffix, email, section, status) " &
                              "VALUES (@student_number, @first_name, @last_name, @middle_initial, @suffix, @email, @section, @status)"

            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@student_number", studentNumber)
                cmd.Parameters.AddWithValue("@first_name", firstName)
                cmd.Parameters.AddWithValue("@last_name", lastName)
                cmd.Parameters.AddWithValue("@middle_initial", middleInitial)
                cmd.Parameters.AddWithValue("@suffix", suffix)
                cmd.Parameters.AddWithValue("@email", email)
                cmd.Parameters.AddWithValue("@section", sectionId)
                cmd.Parameters.AddWithValue("@status", status)

                Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                Return rowsAffected > 0
            End Using

        Catch ex As Exception
            MessageBox.Show("Insert failed: " & ex.Message, "DB Error")
            Return False
        End Try
    End Function

    Public Function GetStudentByEmail(email As String) As StudentModel
        Dim student As StudentModel = Nothing
        Dim query As String = "SELECT * FROM students WHERE email = @Email"

        Try
            Connect()

            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@Email", email)

                Using reader As MySqlDataReader = cmd.ExecuteReader()
                    If reader.Read() Then
                        student = New StudentModel With {
                        .Id = Convert.ToInt32(reader("id")),
                        .StudentNumber = reader("student_number").ToString(),
                        .FirstName = reader("first_name").ToString(),
                        .LastName = reader("last_name").ToString(),
                        .MiddleInitial = If(IsDBNull(reader("middle_initial")), "", reader("middle_initial").ToString()),
                        .Suffix = If(IsDBNull(reader("suffix")), "", reader("suffix").ToString()),
                        .Email = reader("email").ToString(),
                        .Section = If(IsDBNull(reader("section")), "", reader("section").ToString()),
                        .Status = If(IsDBNull(reader("status")), "", reader("status").ToString())
                    }
                    End If
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show("Error retrieving student: " & ex.Message, "Error")
        End Try

        Return student
    End Function


    Public Function DeleteStudent(selectedStudentId As Integer) As Boolean
        Try
            Connect()

            Dim query As String = "DELETE FROM students WHERE id = @id"

            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@id", selectedStudentId)
                Return cmd.ExecuteNonQuery() > 0
            End Using

        Catch ex As Exception
            MessageBox.Show("Delete failed: " & ex.Message, "DB Error")
            Return False
        End Try
    End Function


    Public Function MarkStudentsAsGraduated(studentIds As List(Of Integer), Optional status As String = Nothing) As Boolean
        Try
            Connect()

            ' Calculate year_graduated (e.g., "2024-2025")
            Dim currentYear As Integer = DateTime.Now.Year
            Dim prevYear As Integer = currentYear - 1
            Dim yearGraduated As String = $"{prevYear}-{currentYear}"

            ' Create parameterized query for multiple IDs
            Dim idParameters As New List(Of String)()
            For i As Integer = 0 To studentIds.Count - 1
                idParameters.Add($"@id{i}")
            Next

            Dim idList As String = String.Join(", ", idParameters)

            ' Build the update query dynamically
            Dim updateQuery As String = $"UPDATE students SET isGraduate = 1, year_graduated = @year_graduated"

            ' Only add status update if a specific status is provided
            If status IsNot Nothing Then
                updateQuery &= ", status = @status"
            End If

            updateQuery &= $" WHERE id IN ({idList})"

            Using updateCmd As New MySqlCommand(updateQuery, conn)
                updateCmd.Parameters.AddWithValue("@year_graduated", yearGraduated)

                ' Only add status parameter if a specific status is provided
                If status IsNot Nothing Then
                    updateCmd.Parameters.AddWithValue("@status", status)
                End If

                ' Add parameters for each student ID
                For i As Integer = 0 To studentIds.Count - 1
                    updateCmd.Parameters.AddWithValue($"@id{i}", studentIds(i))
                Next

                Dim rowsAffected As Integer = updateCmd.ExecuteNonQuery()

                ' Verify that all selected students were updated
                If rowsAffected <> studentIds.Count Then
                    MessageBox.Show($"Warning: {rowsAffected} out of {studentIds.Count} students were archived. Some students may not exist anymore.", "Partial Success")
                End If
            End Using

            Return True

        Catch ex As Exception
            MessageBox.Show("Archive failed: " & ex.Message, "DB Error")
            Return False
        Finally
            Disconnect()
        End Try
    End Function

    Public Function UpdateStudent(originalStudentNumber As String, newStudentNumber As String, firstName As String, lastName As String,
                              middleInitial As String, suffix As String, email As String,
                              sectionId As Integer, status As String) As Boolean
        Try
            Connect()

            Dim query As String = "UPDATE students SET " &
                              "student_number = @new_student_number, " &
                              "first_name = @first_name, " &
                              "last_name = @last_name, " &
                              "middle_initial = @middle_initial, " &
                              "suffix = @suffix, " &
                              "email = @email, " &
                              "section = @section, " &
                              "status = @status " &
                              "WHERE student_number = @original_student_number"

            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@new_student_number", newStudentNumber)
                cmd.Parameters.AddWithValue("@first_name", firstName)
                cmd.Parameters.AddWithValue("@last_name", lastName)
                cmd.Parameters.AddWithValue("@middle_initial", middleInitial)
                cmd.Parameters.AddWithValue("@suffix", suffix)
                cmd.Parameters.AddWithValue("@email", email)
                cmd.Parameters.AddWithValue("@section", sectionId)
                cmd.Parameters.AddWithValue("@status", status)
                cmd.Parameters.AddWithValue("@original_student_number", originalStudentNumber)

                Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                Return rowsAffected > 0
            End Using

        Catch ex As Exception
            MessageBox.Show("Update failed: " & ex.Message, "DB Error")
            Return False
        Finally
            Disconnect()

        End Try
    End Function

    ' Get student record by student number
    Public Function GetStudentByNumber(studentNumber As String) As StudentModel
        Try
            Connect()

            Dim query As String = "
            SELECT s.id, s.student_number, s.first_name, s.last_name, s.middle_initial, s.suffix,
                   s.email, sec.section AS section_name, s.status
            FROM students s
            INNER JOIN sections sec ON s.section = sec.id
            WHERE s.student_number = @student_number"

            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@student_number", studentNumber)

                Using reader As MySqlDataReader = cmd.ExecuteReader()
                    If reader.Read() Then
                        Dim student As New StudentModel With {
                        .Id = reader("id"),
                        .StudentNumber = reader("student_number").ToString(),
                        .FirstName = reader("first_name").ToString(),
                        .LastName = reader("last_name").ToString(),
                        .MiddleInitial = reader("middle_initial").ToString(),
                        .Suffix = reader("suffix").ToString(),
                        .Email = reader("email").ToString(),
                        .Section = reader("section_name").ToString(), ' ← Section name instead of ID
                        .Status = reader("status").ToString()
                    }
                        Return student
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Failed to retrieve student: " & ex.Message, "DB Error")
        End Try

        Return Nothing
    End Function


    Public Function DoesStudentExist(studentNumber As String) As Boolean
        Try
            Connect()
            Dim query As String = "SELECT COUNT(*) FROM students WHERE student_number = @student_number"
            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@student_number", studentNumber)
                Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                Return count > 0
            End Using
        Catch ex As Exception
            MessageBox.Show("Error checking student existence: " & ex.Message, "DB Error")
            Return False
        End Try
    End Function


    ' Get list of students matching last name filter and optional section ID
    Public Function SearchStudentsByLastNameAndSection(lastNameFilter As String, Optional sectionId As Integer? = Nothing) As List(Of StudentModel)
        Dim students As New List(Of StudentModel)

        Try
            Connect()

            Dim query As String = "
        SELECT s.id, s.student_number, s.first_name, s.last_name, s.middle_initial, s.suffix,
               s.email, sec.section AS section_name, s.status
        FROM students s
        INNER JOIN sections sec ON s.section = sec.id
        WHERE s.last_name LIKE @lastName
        AND s.isGraduate = 0" ' ADDED THIS LINE

            If sectionId.HasValue Then
                query &= " AND s.section = @sectionId"
            End If

            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@lastName", "%" & lastNameFilter & "%")
                If sectionId.HasValue Then
                    cmd.Parameters.AddWithValue("@sectionId", sectionId.Value)
                End If

                Using reader As MySqlDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        students.Add(New StudentModel With {
                    .Id = reader("id"),
                    .StudentNumber = reader("student_number").ToString(),
                    .FirstName = reader("first_name").ToString(),
                    .LastName = reader("last_name").ToString(),
                    .MiddleInitial = reader("middle_initial").ToString(),
                    .Suffix = reader("suffix").ToString(),
                    .Email = reader("email").ToString(),
                    .Section = reader("section_name").ToString(),
                    .Status = reader("status").ToString()
                })
                    End While
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Search failed: " & ex.Message, "DB Error")
        End Try

        Return students
    End Function


End Module

' Model class representing a student
Public Class StudentModel
    Public Property Id As Integer
    Public Property StudentNumber As String
    Public Property FirstName As String
    Public Property LastName As String
    Public Property MiddleInitial As String
    Public Property Suffix As String
    Public Property Email As String
    Public Property Section As String
    Public Property Status As String
End Class
