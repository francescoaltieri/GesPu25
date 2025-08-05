
Imports Microsoft.Data.SqlClient

Public Class PermessiUtenteForm

    Private nomeUtente As String
    Private connectionString As String = ConnString
    Private dgvPermessi As New DataGridView With {.Location = New Point(20, 60), .Size = New Size(440, 250), .AllowUserToAddRows = False}
    Private modifichePendenti As Boolean = False
    Private btnSalva As Button

    Public Sub New(utente As String)
        InitializeComponent()
        nomeUtente = utente
        Me.Text = $"Autorizzazioni per: {nomeUtente}"
        Me.Controls.Add(dgvPermessi)
        CostruisciGriglia()

        btnSalva = New Button With {
    .Text = "Salva",
    .Location = New Point(20, 330),
    .Width = 100,
    .Enabled = True
}
        AddHandler btnSalva.Click, AddressOf SalvaPermessi
        Me.Controls.Add(btnSalva)

        btnSalva = New Button With {.Text = "Salva", .Location = New Point(20, 330), .Width = 100, .Enabled = False}
        AddHandler btnSalva.Click, AddressOf SalvaPermessi
        Me.Controls.Add(btnSalva)

        AddHandler dgvPermessi.CellValueChanged, AddressOf OnCellChanged

        Dim btnChiudi As New Button With {.Text = "Chiudi", .Location = New Point(140, 330), .Width = 100}
        AddHandler btnChiudi.Click, Sub() Me.Close()
        Me.Controls.Add(btnChiudi)

        Me.Size = New Size(500, 400)

        CaricaPermessi()
        dgvPermessi.AllowUserToAddRows = True

    End Sub

    Private Sub OnCellChanged(sender As Object, e As DataGridViewCellEventArgs)
        modifichePendenti = True
        btnSalva.Enabled = True
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        If modifichePendenti Then
            Dim result = MessageBox.Show("Hai modifiche non salvate. Chiudere comunque?", "Attenzione", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If result = DialogResult.No Then
                e.Cancel = True
            End If
        End If
        MyBase.OnFormClosing(e)
    End Sub

    Private Sub CostruisciGriglia()
        dgvPermessi.Columns.Clear()
        dgvPermessi.Columns.Add("FormName", "Nome Form")

        Dim colView As New DataGridViewCheckBoxColumn With {.Name = "CanView", .HeaderText = "Visualizza"}
        Dim colInsert As New DataGridViewCheckBoxColumn With {.Name = "CanInsert", .HeaderText = "Inserisci"}
        Dim colUpdate As New DataGridViewCheckBoxColumn With {.Name = "CanUpdate", .HeaderText = "Modifica"}
        Dim colDelete As New DataGridViewCheckBoxColumn With {.Name = "CanDelete", .HeaderText = "Cancella"}

        dgvPermessi.Columns.Add(colView)
        dgvPermessi.Columns.Add(colInsert)
        dgvPermessi.Columns.Add(colUpdate)
        dgvPermessi.Columns.Add(colDelete)
    End Sub

    Private Sub CaricaPermessi()
        dgvPermessi.Rows.Clear()

        Using conn As New SqlConnection(connectionString)
            conn.Open()

            Dim cmdAut As New SqlCommand("SELECT Form, CanView, CanInsert, CanUpdate, CanDelete FROM Sys_Autorizzazioni WHERE NomeUtente = @nome", conn)
            cmdAut.Parameters.AddWithValue("@nome", nomeUtente)

            Dim permessi As New Dictionary(Of String, PermessiForm)

            Using reader = cmdAut.ExecuteReader()
                While reader.Read()
                    Dim formName As String = reader.GetString(reader.GetOrdinal("Form"))

                    Dim p As New PermessiForm With {
                    .CanView = reader.GetBoolean(reader.GetOrdinal("CanView")),
                    .CanInsert = reader.GetBoolean(reader.GetOrdinal("CanInsert")),
                    .CanUpdate = reader.GetBoolean(reader.GetOrdinal("CanUpdate")),
                    .CanDelete = reader.GetBoolean(reader.GetOrdinal("CanDelete"))
                }

                    permessi(formName) = p
                End While
            End Using

            ' Riempie la griglia con i permessi letti
            For Each kvp In permessi
                Dim r As Integer = dgvPermessi.Rows.Add()
                dgvPermessi.Rows(r).Cells("FormName").Value = kvp.Key
                dgvPermessi.Rows(r).Cells("CanView").Value = kvp.Value.CanView
                dgvPermessi.Rows(r).Cells("CanInsert").Value = kvp.Value.CanInsert
                dgvPermessi.Rows(r).Cells("CanUpdate").Value = kvp.Value.CanUpdate
                dgvPermessi.Rows(r).Cells("CanDelete").Value = kvp.Value.CanDelete
            Next
        End Using
    End Sub

    Private Sub SalvaPermessi()
        Using conn As New SqlConnection(connectionString)
            conn.Open()

            For Each row As DataGridViewRow In dgvPermessi.Rows
                If row.IsNewRow Then Continue For

                Dim formName = row.Cells("FormName").Value?.ToString()
                If String.IsNullOrWhiteSpace(formName) Then Continue For

                Dim canView = Convert.ToBoolean(row.Cells("CanView").Value)
                Dim canInsert = Convert.ToBoolean(row.Cells("CanInsert").Value)
                Dim canUpdate = Convert.ToBoolean(row.Cells("CanUpdate").Value)
                Dim canDelete = Convert.ToBoolean(row.Cells("CanDelete").Value)

                ' Verifica se già esiste
                Dim queryCheck = "SELECT COUNT(*) FROM Sys_Autorizzazioni WHERE NomeUtente = @nome AND Form = @form"
                Dim esiste As Boolean
                Using cmdCheck As New SqlCommand(queryCheck, conn)
                    cmdCheck.Parameters.AddWithValue("@nome", nomeUtente)
                    cmdCheck.Parameters.AddWithValue("@form", formName)
                    esiste = Convert.ToInt32(cmdCheck.ExecuteScalar()) > 0
                End Using

                If esiste Then
                    ' UPDATE
                    Dim queryUpdate = "UPDATE Sys_Autorizzazioni SET CanView = @v, CanInsert = @i, CanUpdate = @u, CanDelete = @d " &
                          "WHERE NomeUtente = @nome AND Form = @form"
                    Using cmd As New SqlCommand(queryUpdate, conn)
                        cmd.Parameters.AddWithValue("@v", canView)
                        cmd.Parameters.AddWithValue("@i", canInsert)
                        cmd.Parameters.AddWithValue("@u", canUpdate)
                        cmd.Parameters.AddWithValue("@d", canDelete)
                        cmd.Parameters.AddWithValue("@nome", nomeUtente)
                        cmd.Parameters.AddWithValue("@form", formName)
                        cmd.ExecuteNonQuery()
                    End Using
                Else
                    ' INSERT
                    Dim queryInsert = "INSERT INTO Sys_Autorizzazioni (NomeUtente, Form, CanView, CanInsert, CanUpdate, CanDelete) " &
                          "VALUES (@nome, @form, @v, @i, @u, @d)"
                    Using cmd As New SqlCommand(queryInsert, conn)
                        cmd.Parameters.AddWithValue("@nome", nomeUtente)
                        cmd.Parameters.AddWithValue("@form", formName)
                        cmd.Parameters.AddWithValue("@v", canView)
                        cmd.Parameters.AddWithValue("@i", canInsert)
                        cmd.Parameters.AddWithValue("@u", canUpdate)
                        cmd.Parameters.AddWithValue("@d", canDelete)
                        cmd.ExecuteNonQuery()
                    End Using
                End If
            Next

            MessageBox.Show("Permessi salvati correttamente.", "Autorizzazioni", MessageBoxButtons.OK, MessageBoxIcon.Information)
            modifichePendenti = False
            btnSalva.Enabled = False

        End Using
    End Sub

End Class