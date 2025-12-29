Imports MySql.Data.MySqlClient

Module Archive
    ' Returns archived student records with section name, filtered by optional last name and year
    Public Function SearchArchives(Optional lastNameFilter As String = "", Optional yearGraduatedFilter As String = "") As DataTable
        Dim query As String = "
        SELECT 
            s.id,
            s.student_number,
            s.first_name,
            s.last_name,
            s.middle_initial,
            s.suffix,
            s.email,
            sec.section AS section_name,
            s.status,
            s.year_graduated
        FROM 
            students s
        LEFT JOIN 
            sections sec ON s.section = sec.id
        WHERE 
            s.isGraduate = 1" ' CHANGED: Filter for graduated students only

        ' Add filters if provided
        If Not String.IsNullOrWhiteSpace(lastNameFilter) Then
            query &= " AND s.last_name LIKE @lastName"
        End If

        If Not String.IsNullOrWhiteSpace(yearGraduatedFilter) Then
            query &= " AND s.year_graduated = @yearGrad"
        End If

        Dim dt As New DataTable()

        Try
            Connect()
            Dim cmd As New MySqlCommand(query, conn)

            If Not String.IsNullOrWhiteSpace(lastNameFilter) Then
                cmd.Parameters.AddWithValue("@lastName", "%" & lastNameFilter & "%")
            End If

            If Not String.IsNullOrWhiteSpace(yearGraduatedFilter) Then
                cmd.Parameters.AddWithValue("@yearGrad", yearGraduatedFilter)
            End If

            Dim adapter As New MySqlDataAdapter(cmd)
            adapter.Fill(dt)
        Catch ex As Exception
            MessageBox.Show("Archive search failed: " & ex.Message, "Search Error")
        Finally
            Disconnect()
        End Try

        Return dt
    End Function

    Public Function UpdateGraduatedStudent(id As Integer, studentNumber As String, firstName As String, lastName As String,
                                  middleInitial As String, suffix As String, email As String,
                                  sectionId As Integer, status As String, yearGraduated As String) As Boolean
        Try
            Connect()
            Dim query As String = "
        UPDATE students 
        SET student_number = @student_number, 
            first_name = @first_name, 
            last_name = @last_name, 
            middle_initial = @middle_initial, 
            suffix = @suffix, 
            email = @email, 
            section = @section, 
            status = @status, 
            year_graduated = @year_graduated,
            isGraduate = 1
        WHERE id = @id"

            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@student_number", studentNumber)
                cmd.Parameters.AddWithValue("@first_name", firstName)
                cmd.Parameters.AddWithValue("@last_name", lastName)
                cmd.Parameters.AddWithValue("@middle_initial", middleInitial)
                cmd.Parameters.AddWithValue("@suffix", suffix)
                cmd.Parameters.AddWithValue("@email", email)
                cmd.Parameters.AddWithValue("@section", sectionId)
                cmd.Parameters.AddWithValue("@status", status)
                cmd.Parameters.AddWithValue("@year_graduated", yearGraduated)
                cmd.Parameters.AddWithValue("@id", id)

                Return cmd.ExecuteNonQuery() > 0
            End Using
        Catch ex As Exception
            MessageBox.Show("Update failed: " & ex.Message, "DB Error")
            Return False
        Finally
            Disconnect()
        End Try
    End Function


    Public Function InsertGraduatedStudent(studentNumber As String, firstName As String, lastName As String,
                                  middleInitial As String, suffix As String, email As String,
                                  sectionId As Integer, status As String, yearGraduated As String) As Boolean
        Try
            Connect()
            Dim query As String = "
        INSERT INTO students (
            student_number, first_name, last_name, middle_initial, suffix,
            email, section, status, isGraduate, year_graduated
        ) VALUES (
            @student_number, @first_name, @last_name, @middle_initial, @suffix,
            @email, @section, @status, 1, @year_graduated
        )"

            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@student_number", studentNumber)
                cmd.Parameters.AddWithValue("@first_name", firstName)
                cmd.Parameters.AddWithValue("@last_name", lastName)
                cmd.Parameters.AddWithValue("@middle_initial", middleInitial)
                cmd.Parameters.AddWithValue("@suffix", suffix)
                cmd.Parameters.AddWithValue("@email", email)
                cmd.Parameters.AddWithValue("@section", sectionId)
                cmd.Parameters.AddWithValue("@status", status)
                cmd.Parameters.AddWithValue("@year_graduated", yearGraduated)

                Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                Return rowsAffected > 0
            End Using
        Catch ex As Exception
            MessageBox.Show("Insert failed: " & ex.Message, "DB Error")
            Return False
        Finally
            Disconnect()
        End Try
    End Function


    Public Function UnarchiveStudents(studentIds As List(Of Integer),
                                  Optional firstName As String = Nothing,
                                  Optional lastName As String = Nothing,
                                  Optional middleInitial As String = Nothing,
                                  Optional suffix As String = Nothing,
                                  Optional email As String = Nothing,
                                  Optional sectionId As Integer? = Nothing,
                                  Optional status As String = Nothing) As Boolean
        Try
            Connect()

            ' Create parameterized query for multiple IDs
            Dim idParameters As New List(Of String)()
            For i As Integer = 0 To studentIds.Count - 1
                idParameters.Add($"@id{i}")
            Next

            Dim idList As String = String.Join(", ", idParameters)
            Dim query As String = $"UPDATE students SET isGraduate = 0, year_graduated = NULL"

            ' Add optional field updates if provided (for single student)
            If firstName IsNot Nothing Then query &= ", first_name = @first_name"
            If lastName IsNot Nothing Then query &= ", last_name = @last_name"
            If middleInitial IsNot Nothing Then query &= ", middle_initial = @middle_initial"
            If suffix IsNot Nothing Then query &= ", suffix = @suffix"
            If email IsNot Nothing Then query &= ", email = @email"
            If sectionId.HasValue Then query &= ", section = @section"
            If status IsNot Nothing Then query &= ", status = @status"

            query &= $" WHERE id IN ({idList}) AND isGraduate = 1"

            Using cmd As New MySqlCommand(query, conn)
                ' Add parameters for optional fields if provided
                If firstName IsNot Nothing Then cmd.Parameters.AddWithValue("@first_name", firstName)
                If lastName IsNot Nothing Then cmd.Parameters.AddWithValue("@last_name", lastName)
                If middleInitial IsNot Nothing Then cmd.Parameters.AddWithValue("@middle_initial", middleInitial)
                If suffix IsNot Nothing Then cmd.Parameters.AddWithValue("@suffix", suffix)
                If email IsNot Nothing Then cmd.Parameters.AddWithValue("@email", email)
                If sectionId.HasValue Then cmd.Parameters.AddWithValue("@section", sectionId.Value)
                If status IsNot Nothing Then cmd.Parameters.AddWithValue("@status", status)

                ' Add parameters for each student ID
                For i As Integer = 0 To studentIds.Count - 1
                    cmd.Parameters.AddWithValue($"@id{i}", studentIds(i))
                Next

                Dim rowsAffected As Integer = cmd.ExecuteNonQuery()

                ' Verify that all selected students were updated
                If rowsAffected <> studentIds.Count Then
                    MessageBox.Show($"Warning: {rowsAffected} out of {studentIds.Count} students were unarchived. Some students may not be graduated or may not exist.", "Partial Success")
                End If

                Return rowsAffected > 0
            End Using

        Catch ex As Exception
            MessageBox.Show("Unarchive failed: " & ex.Message, "DB Error")
            Return False
        Finally
            Disconnect()
        End Try
    End Function


End Module
