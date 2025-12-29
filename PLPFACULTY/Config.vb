Module Config
    ' Get the latest graduation date
    Public Function GetLatestGraduationDate() As Date
        Dim query As String = "SELECT date FROM config ORDER BY id DESC LIMIT 1"
        Dim dt As DataTable = ExecuteQuery(query)

        If dt.Rows.Count > 0 Then
            Return Convert.ToDateTime(dt.Rows(0)("date"))
        Else
            Return Date.Today
        End If
    End Function

    ' Insert new graduation date
    Public Sub SetGraduationDate(graduationDate As Date)
        Dim query As String = "INSERT INTO config (date) VALUES ('" & graduationDate.ToString("yyyy-MM-dd") & "')"
        ExecuteNonQuery(query)
    End Sub
End Module
