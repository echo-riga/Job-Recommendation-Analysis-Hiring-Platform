Module Reason
    ' Get all reasons
    Public Function GetAllReasons() As DataTable
        Return ExecuteQuery("SELECT * FROM reasons")
    End Function

    ' Insert a new reason
    Public Sub InsertReason(reasonText As String, isSpecial As Boolean)
        Dim isSpecialVal As Integer = If(isSpecial, 1, 0)
        Dim query As String = "INSERT INTO reasons (reason, is_special) VALUES ('" & reasonText.Replace("'", "''") & "', " & isSpecialVal & ")"
        ExecuteNonQuery(query)
    End Sub

    ' Update an existing reason by ID
    Public Sub UpdateReason(id As Integer, reasonText As String, isSpecial As Boolean)
        Dim isSpecialVal As Integer = If(isSpecial, 1, 0)
        Dim query As String = "UPDATE reasons SET reason = '" & reasonText.Replace("'", "''") & "', is_special = " & isSpecialVal & " WHERE id = " & id
        ExecuteNonQuery(query)
    End Sub

    ' Delete a reason by ID
    Public Sub DeleteReason(id As Integer)
        Dim query As String = "DELETE FROM reasons WHERE id = " & id
        ExecuteNonQuery(query)
    End Sub

    ' Filter by reason text
    Public Function GetReasons(Optional keyword As String = "", Optional isSpecial As Nullable(Of Boolean) = Nothing) As DataTable
        Dim query As String = "SELECT * FROM reasons WHERE 1=1"

        ' Filter by keyword if provided
        If Not String.IsNullOrWhiteSpace(keyword) Then
            query &= " AND reason LIKE '%" & keyword.Replace("'", "''") & "%'"
        End If

        ' Filter by isSpecial if provided
        If isSpecial.HasValue Then
            Dim specialVal As Integer = If(isSpecial.Value, 1, 0)
            query &= " AND is_special = " & specialVal
        End If

        Return ExecuteQuery(query)
    End Function


End Module
