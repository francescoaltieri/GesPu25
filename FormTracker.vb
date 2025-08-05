Imports Microsoft.Data.SqlClient

Public Class FormTracker
    Private connectionString As String = "Server=TUO_SERVER;Database=TUO_DB;Trusted_Connection=True;"

    Public Sub LoadFormState(form As Form)
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim cmd As New SqlCommand("SELECT PosX, PosY, Width, Height, WindowState FROM FormSettings WHERE FormName = @FormName", conn)
            cmd.Parameters.AddWithValue("@FormName", form.Name)

            Using reader = cmd.ExecuteReader()
                If reader.Read() Then
                    form.Location = New Point(reader.GetInt32(0), reader.GetInt32(1))
                    form.Size = New Size(reader.GetInt32(2), reader.GetInt32(3))
                    form.WindowState = CType(reader.GetInt32(4), FormWindowState)
                End If
            End Using
        End Using
    End Sub

    Public Sub SaveFormState(form As Form)
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim cmd As New SqlCommand("
                MERGE INTO FormSettings AS target
                USING (SELECT @FormName AS FormName) AS source
                ON (target.FormName = source.FormName)
                WHEN MATCHED THEN
                    UPDATE SET PosX=@PosX, PosY=@PosY, Width=@Width, Height=@Height, WindowState=@WindowState
                WHEN NOT MATCHED THEN
                    INSERT (FormName, PosX, PosY, Width, Height, WindowState)
                    VALUES (@FormName, @PosX, @PosY, @Width, @Height, @WindowState);", conn)

            cmd.Parameters.AddWithValue("@FormName", form.Name)
            cmd.Parameters.AddWithValue("@PosX", form.Location.X)
            cmd.Parameters.AddWithValue("@PosY", form.Location.Y)
            cmd.Parameters.AddWithValue("@Width", form.Size.Width)
            cmd.Parameters.AddWithValue("@Height", form.Size.Height)
            cmd.Parameters.AddWithValue("@WindowState", CInt(form.WindowState))

            cmd.ExecuteNonQuery()
        End Using
    End Sub

End Class
