Imports MySql.Data.MySqlClient

Module Section
    ' Get all sections
    Public Function GetAllSections() As DataTable
        Dim query As String = "SELECT * FROM sections"

        Try
            Return ExecuteQuery(query)
        Catch ex As Exception
            MessageBox.Show("Failed to retrieve sections: " & ex.Message, "SectionModel Error")
            Return New DataTable()
        End Try
    End Function

    ' Get sections filtered by section name
    Public Function GetSections(Optional keyword As String = "") As DataTable
        Dim query As String = "SELECT * FROM sections WHERE 1=1"

        If Not String.IsNullOrWhiteSpace(keyword) Then
            query &= " AND section LIKE '%" & keyword.Replace("'", "''") & "%'"
        End If

        Try
            Return ExecuteQuery(query)
        Catch ex As Exception
            MessageBox.Show("Failed to filter sections: " & ex.Message, "SectionModel Error")
            Return New DataTable()
        End Try
    End Function
    Public Function GetSectionIdByName(sectionName As String) As Integer
        Dim query As String = $"SELECT id FROM sections WHERE section = @section LIMIT 1"
        Try
            Connect()
            Using cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@section", sectionName)
                Using reader = cmd.ExecuteReader()
                    If reader.Read() Then
                        Return Convert.ToInt32(reader("id"))
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Failed to look up section: " & ex.Message)
        End Try
        Return -1
    End Function

    ' Insert new section
    Public Sub InsertSection(sectionName As String)
        Dim query As String = "INSERT INTO sections (section) VALUES ('" & sectionName.Replace("'", "''") & "')"

        Try
            ExecuteNonQuery(query)
        Catch ex As Exception
            MessageBox.Show("Failed to insert section: " & ex.Message, "SectionModel Error")
        End Try
    End Sub

    ' Update existing section
    Public Sub UpdateSection(id As Integer, sectionName As String)
        Dim query As String = "UPDATE sections SET section = '" & sectionName.Replace("'", "''") & "' WHERE id = " & id

        Try
            ExecuteNonQuery(query)
        Catch ex As Exception
            MessageBox.Show("Failed to update section: " & ex.Message, "SectionModel Error")
        End Try
    End Sub

    ' Delete section
    Public Sub DeleteSection(id As Integer)
        Dim query As String = "DELETE FROM sections WHERE id = " & id

        Try
            ExecuteNonQuery(query)
        Catch ex As Exception
            MessageBox.Show("Failed to delete section: " & ex.Message, "SectionModel Error")
        End Try
    End Sub
End Module
