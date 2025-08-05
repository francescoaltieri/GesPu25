Imports Microsoft.Data.SqlClient

Module GestioneStatoForm

    Public Sub CaricaStato(form As Form)
        Dim query = "SELECT X, Y, Width, Height, WindowsState FROM Sys_Form WHERE FormName = @nome"

        Using conn As New SqlConnection(ConnString)
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@nome", form.Name)
                conn.Open()

                Using reader = cmd.ExecuteReader()
                    If reader.Read() Then
                        Dim stato = reader("WindowsState")?.ToString()?.ToLower()
                        Select Case stato
                            Case "normal" : form.WindowState = FormWindowState.Normal
                            Case "maximized" : form.WindowState = FormWindowState.Maximized
                            Case "minimized" : form.WindowState = FormWindowState.Minimized
                        End Select

                        If form.WindowState = FormWindowState.Normal Then
                            If Not reader.IsDBNull(0) AndAlso Not reader.IsDBNull(1) Then
                                form.StartPosition = FormStartPosition.Manual
                                form.Location = New Point(reader.GetInt32(0), reader.GetInt32(1))
                            End If

                            If Not reader.IsDBNull(2) AndAlso Not reader.IsDBNull(3) Then
                                form.Size = New Size(reader.GetInt32(2), reader.GetInt32(3))
                            End If
                        End If
                    End If
                End Using
            End Using
        End Using
    End Sub

    Public Sub SalvaStato(form As Form)
        Dim statoFinestra = form.WindowState.ToString()
        Dim posizione = If(form.WindowState = FormWindowState.Normal, form.Location, form.RestoreBounds.Location)
        Dim dimensione = If(form.WindowState = FormWindowState.Normal, form.Size, form.RestoreBounds.Size)
        Dim queryCheck = "SELECT COUNT(*) FROM Sys_Form WHERE FormName = @nome"
        Dim queryInsert = "
            INSERT INTO Sys_Form (FormName, X, Y, Width, Height, WindowsState) 
            VALUES (@nome, @x, @y, @w, @h, @stato)"
        Dim queryUpdate = "
            UPDATE Sys_Form SET X = @x, Y = @y, Width = @w, Height = @h, WindowsState = @stato 
            WHERE FormName = @nome"

        Using conn As New SqlConnection(ConnString)
            Using cmdCheck As New SqlCommand(queryCheck, conn)
                cmdCheck.Parameters.AddWithValue("@nome", form.Name)
                conn.Open()
                Dim exists = Convert.ToInt32(cmdCheck.ExecuteScalar()) > 0
                conn.Close()

                Dim cmd As New SqlCommand(If(exists, queryUpdate, queryInsert), conn)
                cmd.Parameters.AddWithValue("@nome", form.Name)
                cmd.Parameters.AddWithValue("@x", posizione.X)
                cmd.Parameters.AddWithValue("@y", posizione.Y)
                cmd.Parameters.AddWithValue("@w", dimensione.Width)
                cmd.Parameters.AddWithValue("@h", dimensione.Height)
                cmd.Parameters.AddWithValue("@stato", statoFinestra)

                conn.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

End Module

