Imports MySql.Data.MySqlClient

Module Report
    Public Sub InsertReport(studentId As Integer,
                            reasonId As Integer,
                            message As String,
                            consultationDate As Date,
                            timeIn As TimeSpan,
                            timeOut As TimeSpan,
                            Optional professorId As Integer? = Nothing)

        Dim query As String

        ' Construct SQL with or without professorId
        If professorId.HasValue Then
            query = $"INSERT INTO reports (student_id, professor_id, reason_id, message, consultation_date, time_in, time_out) " &
                    $"VALUES ({studentId}, {professorId.Value}, {reasonId}, '{MySqlHelper.EscapeString(message)}', '{consultationDate:yyyy-MM-dd}', '{timeIn}', '{timeOut}')"
        Else
            query = $"INSERT INTO reports (student_id, professor_id, reason_id, message, consultation_date, time_in, time_out) " &
                    $"VALUES ({studentId}, NULL, {reasonId}, '{MySqlHelper.EscapeString(message)}', '{consultationDate:yyyy-MM-dd}', '{timeIn}', '{timeOut}')"
        End If

        ExecuteNonQuery(query)
    End Sub
    Public Function GetFormattedReports(fromDate As Date, toDate As Date, Optional professorId As Integer = 0) As DataTable
        Dim query As String = "
SELECT
    r.id,
    s.student_number,
    CONCAT(
        s.last_name, ', ',
        s.first_name,
        IF(s.middle_initial IS NOT NULL AND s.middle_initial != '', CONCAT(' ', s.middle_initial, '.'), ''),
        IF(s.suffix IS NOT NULL AND s.suffix != '', CONCAT(' ', s.suffix), '')
    ) AS student_name,
    sec.section,
    CONCAT(
        p.last_name, ', ',
        p.first_name,
        IF(p.middle_initial IS NOT NULL AND p.middle_initial != '', CONCAT(' ', p.middle_initial, '.'), ''),
        IF(p.suffix IS NOT NULL AND p.suffix != '', CONCAT(' ', p.suffix), '')
    ) AS professor_name,
    rs.reason,
    r.message,
    DATE_FORMAT(r.consultation_date, '%M %d, %Y') AS consultation_date,
    DATE_FORMAT(r.time_in, '%h:%i %p') AS time_in,
    DATE_FORMAT(r.time_out, '%h:%i %p') AS time_out
FROM reports r
JOIN students s ON r.student_id = s.id
JOIN sections sec ON s.section = sec.id
LEFT JOIN professors p ON r.professor_id = p.id
JOIN reasons rs ON r.reason_id = rs.id
WHERE r.consultation_date BETWEEN @fromDate AND @toDate
"

        ' Apply filtering based on professorId:
        ' 0  = All professors
        ' 99 = "Special Reasons" → Only records with NULL professor_id
        If professorId > 0 AndAlso professorId <> 99 Then
            query &= " AND r.professor_id = @professorId"
        ElseIf professorId = 99 Then
            query &= " AND (r.professor_id IS NULL OR r.professor_id = 0)"
        End If

        query &= " ORDER BY r.consultation_date DESC, r.time_in DESC"

        Dim dt As New DataTable()
        Try
            Connect()
            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@fromDate", fromDate.ToString("yyyy-MM-dd"))
                cmd.Parameters.AddWithValue("@toDate", toDate.ToString("yyyy-MM-dd"))

                If professorId > 0 AndAlso professorId <> 99 Then
                    cmd.Parameters.AddWithValue("@professorId", professorId)
                End If

                Using da As New MySqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Query failed: " & ex.Message, "Query Error")
        End Try

        Return dt
    End Function



    Public Function GetFormattedReportsByProfessor(profId As Integer, fromDate As Date, toDate As Date, Optional sectionId As Integer? = Nothing) As DataTable
        Dim query As String = "
SELECT
    s.student_number,
    CONCAT(
        s.last_name, ', ',
        s.first_name,
        IF(s.middle_initial IS NOT NULL AND s.middle_initial != '', CONCAT(' ', s.middle_initial, '.'), ''),
        IF(s.suffix IS NOT NULL AND s.suffix != '', CONCAT(' ', s.suffix), '')
    ) AS student_name,
    sec.section AS section,
    rs.reason,
    r.message,
    DATE_FORMAT(r.consultation_date, '%M %d, %Y') AS consultation_date,
    DATE_FORMAT(r.time_in, '%h:%i %p') AS time_in,
    DATE_FORMAT(r.time_out, '%h:%i %p') AS time_out
FROM reports r
JOIN students s ON r.student_id = s.id
JOIN sections sec ON s.section = sec.id
JOIN reasons rs ON r.reason_id = rs.id
WHERE r.professor_id = @profId
  AND r.consultation_date BETWEEN @fromDate AND @toDate"

        ' Add section filter only if passed
        If sectionId.HasValue Then
            query &= " AND s.section = @sectionId"
        End If

        query &= " ORDER BY r.consultation_date DESC, r.time_in DESC"

        Dim dt As New DataTable()
        Try
            Connect()
            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@profId", profId)
                cmd.Parameters.AddWithValue("@fromDate", fromDate.ToString("yyyy-MM-dd"))
                cmd.Parameters.AddWithValue("@toDate", toDate.ToString("yyyy-MM-dd"))
                If sectionId.HasValue Then
                    cmd.Parameters.AddWithValue("@sectionId", sectionId.Value)
                End If

                Using da As New MySqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Query failed: " & ex.Message, "Query Error")
        End Try

        Return dt
    End Function




    Public Sub UpdateReport(reportId As Integer,
                        reasonId As Integer,
                        message As String,
                        consultationDate As Date,
                        timeIn As TimeSpan,
                        timeOut As TimeSpan,
                        Optional professorId As Integer? = Nothing,
                        Optional studentNumber As String = Nothing)

        Try
            Connect()

            ' Step 1: Get student_id using student_number
            Dim studentId As Integer? = Nothing

            If Not String.IsNullOrWhiteSpace(studentNumber) Then
                Dim studentQuery As String = "SELECT id FROM students WHERE student_number = @student_number"
                Using studentCmd As New MySqlCommand(studentQuery, conn)
                    studentCmd.Parameters.AddWithValue("@student_number", studentNumber)
                    Dim result = studentCmd.ExecuteScalar()
                    If result IsNot Nothing Then
                        studentId = Convert.ToInt32(result)
                    Else
                        MessageBox.Show("Student number not found. Please enter a valid one.", "Validation Error")
                        Exit Sub
                    End If
                End Using
            End If

            ' Step 2: Prepare the update query
            Dim query As String = "UPDATE reports SET " &
                              "reason_id = @reasonId, " &
                              "professor_id = @professorId, " &
                              "message = @message, " &
                              "consultation_date = @consultationDate, " &
                              "time_in = @timeIn, " &
                              "time_out = @timeOut, " &
                              "student_id = @studentId " &
                              "WHERE id = @reportId"

            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@reasonId", reasonId)
                cmd.Parameters.AddWithValue("@professorId", If(professorId.HasValue, professorId.Value, DBNull.Value))
                cmd.Parameters.AddWithValue("@message", MySqlHelper.EscapeString(message))
                cmd.Parameters.AddWithValue("@consultationDate", consultationDate.ToString("yyyy-MM-dd"))
                cmd.Parameters.AddWithValue("@timeIn", timeIn)
                cmd.Parameters.AddWithValue("@timeOut", timeOut)
                cmd.Parameters.AddWithValue("@studentId", studentId)
                cmd.Parameters.AddWithValue("@reportId", reportId)

                cmd.ExecuteNonQuery()
            End Using

        Catch ex As Exception
            MessageBox.Show("Error updating report: " & ex.Message, "Database Error")
        End Try
    End Sub


    Public Sub DeleteReport(reportId As Integer)
        Dim query As String = $"DELETE FROM reports WHERE id = {reportId}"
        ExecuteNonQuery(query)
    End Sub

End Module
