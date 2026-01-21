Imports System.Data.SqlClient
Imports System.Text
Imports System.Threading.Tasks
Imports Microsoft.Data.SqlClient
Imports Microsoft.SqlServer.Management.Common
Imports Microsoft.SqlServer.Management.Smo

Public Class SysNuovoDatabase

    Private Sub btnCreaNuovoDatabase_Click(sender As Object, e As EventArgs) Handles btnCreaNuovoDatabase.Click
        Dim newDbName As String = txtNomeNuovoDatabase.Text.Trim()
        If String.IsNullOrEmpty(newDbName) Then
            MessageBox.Show("Inserisci il nome del nuovo database.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim sourceDbName As String = "GesPu"

        'Try
        Using sqlConn As New SqlConnection(ConnStringSa)
                sqlConn.Open()
                Dim srvConn As New ServerConnection(sqlConn)
                Dim server As New Server(srvConn)

                If Not server.Databases.Contains(sourceDbName) Then
                    MessageBox.Show($"Database sorgente '{sourceDbName}' non trovato.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                ' Se esiste già, chiedi conferma e rimuovi
                If server.Databases.Contains(newDbName) Then
                    If MessageBox.Show($"Il database '{newDbName}' esiste già. Sovrascrivere?", "Conferma", MessageBoxButtons.YesNo) <> DialogResult.Yes Then
                        Return
                    End If
                    server.KillDatabase(newDbName)
                End If

                Dim sourceDb As Database = server.Databases(sourceDbName)
                Dim destDb As New Database(server, newDbName)
                destDb.Create()

            ' --- Transfer: schema + dati ---
            Dim xfr As New Transfer(sourceDb) With {
    .CopySchema = True,
    .CopyData = True,
    .CopyAllTables = True,
    .DestinationDatabase = newDbName,
    .DestinationServer = server.Name,
    .DestinationLoginSecure = False,
    .DestinationLogin = "sa",
    .DestinationPassword = "p1w5r0d0G"
}
            xfr.Options.WithDependencies = True
            xfr.Options.ContinueScriptingOnError = True
            xfr.Options.DriAll = True

            xfr.TransferData()


            ' Ricrea utenti DB e mapping (semplice)
            RecreateDatabaseUsers(server, sourceDb, newDbName)

                MessageBox.Show($"Copia completata: {sourceDbName} -> {newDbName}", "Fatto", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Using
        'Catch ex As Exception
        '    MessageBox.Show("Errore: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'End Try
    End Sub

    Private Sub RecreateDatabaseUsers(server As Server, sourceDb As Database, targetDbName As String)
        Dim sb As New StringBuilder()
        For Each u As User In sourceDb.Users
            If u.IsSystemObject Then Continue For

            ' Escape semplice dei nomi (migliorare se necessario)
            Dim userName As String = u.Name.Replace("]", "]]")
            Dim loginName As String = If(String.IsNullOrEmpty(u.Login), Nothing, u.Login.Replace("]", "]]"))

            sb.AppendLine($"USE [{targetDbName}];")
            sb.AppendLine($"IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = N'{userName}')")
            If Not String.IsNullOrEmpty(loginName) Then
                sb.AppendLine($"    CREATE USER [{userName}] FOR LOGIN [{loginName}];")
                sb.AppendLine($"    ALTER ROLE db_owner ADD MEMBER [{userName}];")
            Else
                sb.AppendLine($"    CREATE USER [{userName}];")
            End If
        Next

        If sb.Length = 0 Then Return

        ' Costruisci una connection string SQL Auth verso il DB di destinazione
        ' Usa le credenziali SQL valide per il server 10.8.0.1
        Dim sqlServer As String = "10.8.0.1,1433" ' o server.Name se appropriato
        Dim saUser As String = "sa"
        Dim saPassword As String = "p1w5r0d0G" ' NON hardcodare in produzione; usa secure store
        Dim connString As String = $"Server={sqlServer};Database={targetDbName};User Id={saUser};Password={saPassword};TrustServerCertificate=True;"

        Using conn As New SqlConnection(connString)
            conn.Open()
            Using cmd As New SqlCommand(sb.ToString(), conn)
                cmd.CommandTimeout = 0 ' disabilita timeout se script lungo
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub


    Private Sub AggiornaStato(msg As String)
        Try
            lblStato.Text = msg
            lblStato.Refresh()
            Application.DoEvents()
        Catch
        End Try
    End Sub

    Private Sub btnAnnulla_Click(sender As Object, e As EventArgs) Handles btnAnnulla.Click
        Me.Close()
    End Sub

End Class