Imports System.IO
Imports Microsoft.Data.SqlClient

Public Class SceltaVideo

    Public Property RevisioneSelezionata As RevisioneParametri
    Private videoFormDestinazione As VideoFBF

    Public Sub New(destinazione As VideoFBF)
        InitializeComponent()
        videoFormDestinazione = destinazione
    End Sub

    Private Sub SceltaVideo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CaricaRevisioni()
    End Sub

    Private Sub CaricaRevisioni()
        Dim nomeUtente As String = SessioneUtente.NomeUtenteCorrente
        Dim dt As New DataTable()

        Using conn As New SqlConnection(ConnString)
            Dim query As String = "
        SELECT 
            R.RevisioneID,
            R.DataRevisione,
            V.VideoID,
            V.Titolo AS TitoloVideo,
            R.Autore,
            R.NumRetake,
            R.Stato,
            R.Approvato,
            R.Note
        FROM Mov_Revisioni R
        INNER JOIN Mov_ConsegneScene V ON R.VideoID = V.VideoID
        INNER JOIN Mov_RevisioniUtente UR ON R.RevisioneID = UR.RevisioneID
        WHERE UR.NomeUtente = @NomeUtente
        ORDER BY V.Titolo, R.DataRevisione ASC;"

            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@NomeUtente", nomeUtente)
                conn.Open()
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using

        ' Aggiungi colonna calcolata: NumeroRevisione
        dt.Columns.Add("NumeroRevisione", GetType(String))
        For Each row As DataRow In dt.Rows
            Dim revisioneID As Integer = CInt(row("RevisioneID"))
            row("NumeroRevisione") = $"Revisione_{revisioneID:000}"
        Next

        dgvRevisioni.DataSource = dt
        dgvRevisioni.Columns("VideoID").Visible = False
        dgvRevisioni.Columns("RevisioneID").Visible = False
        dgvRevisioni.Columns("NumeroRevisione").DisplayIndex = 0

        For Each col As DataGridViewColumn In dgvRevisioni.Columns
            If col.Visible Then
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            End If
        Next

        dgvRevisioni.ReadOnly = True

    End Sub

    Private Sub dgvRevisioni_KeyDown(sender As Object, e As KeyEventArgs) Handles dgvRevisioni.KeyDown
        If e.KeyCode = Keys.Delete Then
            If dgvRevisioni.SelectedRows.Count = 0 Then Exit Sub

            Dim row = dgvRevisioni.SelectedRows(0)
            Dim revisioneID = CInt(row.Cells("RevisioneID").Value)

            If Not RevisioneCancellabile(revisioneID) Then
                MDIMessageBox.Show("La Revisione " & revisioneID & " non può essere cancellata perché esistono altre revisioni per questo video.", Me.MdiParent, MessageBoxButtons.OK, "Operazione non consentita")
                Exit Sub
            End If

            If Not ModuloAutorizzazioni.UtenteAutorizzato("SceltaVideo", "delete", SessioneUtente.NomeUtenteCorrente) Then
                MDIMessageBox.Show("Non hai i permessi per cancellare questa revisione.", Me.MdiParent, MessageBoxButtons.OK, "Accesso negato")
                Exit Sub
            End If

            If MDIMessageBox.Show("Vuoi davvero eliminare questa revisione?", Me.MdiParent, MessageBoxButtons.YesNo, "Conferma 1") <> DialogResult.Yes Then Exit Sub
            If MDIMessageBox.Show("Conferma definitiva: la revisione sarà eliminata in modo permanente.", Me.MdiParent, MessageBoxButtons.YesNo, "Conferma 2") <> DialogResult.Yes Then Exit Sub


            CancellaRevisione(revisioneID)

            CaricaRevisioni()
        End If
    End Sub

    Private Sub CancellaRevisione(revisioneID As Integer)

        Dim titoloVideo As String = ""
        Dim percorsoRevisione As String = ""
        Dim percorsoVideo As String = ""

        Using conn As New SqlConnection(ConnString)
            conn.Open()

            ' Recupera VideoID e Titolo
            Dim queryInfo As String = "
            SELECT V.Titolo
            FROM Mov_Revisioni R
            INNER JOIN Mov_ConsegneScene V ON R.VideoID = V.VideoID
            WHERE R.RevisioneID = @RevisioneID"
            Using cmd As New SqlCommand(queryInfo, conn)
                cmd.Parameters.AddWithValue("@RevisioneID", revisioneID)
                Using reader = cmd.ExecuteReader()
                    If reader.Read() Then
                        titoloVideo = reader("Titolo").ToString()
                    End If
                End Using
            End Using

            ' Elimina record dipendenti
            Dim queryDipendenti As String = "DELETE FROM Mov_RevisioniUtente WHERE RevisioneID = @RevisioneID"
            Using cmdDip As New SqlCommand(queryDipendenti, conn)
                cmdDip.Parameters.AddWithValue("@RevisioneID", revisioneID)
                cmdDip.ExecuteNonQuery()
            End Using

            ' Elimina revisione
            Dim queryDelete As String = "DELETE FROM Mov_Revisioni WHERE RevisioneID = @RevisioneID"
            Using cmdDel As New SqlCommand(queryDelete, conn)
                cmdDel.Parameters.AddWithValue("@RevisioneID", revisioneID)
                cmdDel.ExecuteNonQuery()
            End Using
        End Using

        ' Percorsi
        percorsoVideo = Path.Combine("C:\VideoEditor\Frames", titoloVideo)
        percorsoRevisione = Path.Combine(percorsoVideo, $"Revisione_{revisioneID:000}")

        ' Elimina cartella della revisione
        If Directory.Exists(percorsoRevisione) Then
            Try
                Directory.Delete(percorsoRevisione, True)
            Catch ex As Exception
                MDIMessageBox.Show($"Impossibile eliminare la cartella della revisione: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK, "Attenzione")
            End Try
        End If

        ' Se revisione 0, controlla se la cartella del video è vuota
        If revisioneID = 0 AndAlso Directory.Exists(percorsoVideo) Then
            If Directory.GetDirectories(percorsoVideo).Length = 0 AndAlso Directory.GetFiles(percorsoVideo).Length = 0 Then
                Try
                    Directory.Delete(percorsoVideo, True)
                Catch ex As Exception
                    MDIMessageBox.Show($"La revisione è stata cancellata, ma non è stato possibile eliminare la cartella del video: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK, "Attenzione")
                End Try
            End If
        End If

        ' Chiudi VideoFBF se la revisione cancellata è quella attiva
        For Each f As Form In Application.OpenForms
            If TypeOf f Is VideoFBF Then
                Dim videoForm = DirectCast(f, VideoFBF)
                If videoForm.lblRevAttiva.Text = revisioneID.ToString() Then
                    VideoFBF.picFrame.Image = Nothing
                    Exit For
                End If
            End If
        Next

        MDIMessageBox.Show("Revisione eliminata correttamente.", Me.MdiParent, MessageBoxButtons.OK, "Operazione completata")
    End Sub

    Private Function RevisioneCancellabile(revisioneID As Integer) As Boolean
        Using conn As New SqlConnection(ConnString)
            conn.Open()

            ' Recupera VideoID e Titolo della revisione
            Dim videoID As Integer = 0
            Dim titoloVideo As String = ""
            Dim queryInfo As String = "
            SELECT V.VideoID, V.Titolo
            FROM Mov_Revisioni R
            INNER JOIN Mov_ConsegneScene V ON R.VideoID = V.VideoID
            WHERE R.RevisioneID = @RevisioneID"
            Using cmd As New SqlCommand(queryInfo, conn)
                cmd.Parameters.AddWithValue("@RevisioneID", revisioneID)
                Using reader = cmd.ExecuteReader()
                    If reader.Read() Then
                        videoID = CInt(reader("VideoID"))
                        titoloVideo = reader("Titolo").ToString()
                    Else
                        Return False ' Revisione non trovata
                    End If
                End Using
            End Using

            ' Verifica che non esistano revisioni successive
            Dim querySucc = "
            SELECT COUNT(*) 
            FROM Mov_Revisioni 
            WHERE VideoID = @VideoID AND RevisioneID > @RevisioneID"
            Using cmdSucc As New SqlCommand(querySucc, conn)
                cmdSucc.Parameters.AddWithValue("@VideoID", videoID)
                cmdSucc.Parameters.AddWithValue("@RevisioneID", revisioneID)
                Dim countSucc = CInt(cmdSucc.ExecuteScalar())
                If countSucc > 0 Then Return False
            End Using

            Return True
        End Using
    End Function

    Private Function OttieniPercorsoCartellaRevisione(revisioneID As Integer) As String
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query As String = "
            SELECT V.Titolo, R.RevisioneID
            FROM Mov_Revisioni R
            INNER JOIN Mov_ConsegneScene V ON R.VideoID = V.VideoID
            WHERE R.RevisioneID = @RevisioneID"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@RevisioneID", revisioneID)
                Using reader = cmd.ExecuteReader()
                    If reader.Read() Then
                        Dim titolo = reader("Titolo").ToString()
                        Dim numero = CInt(reader("RevisioneID")) ' Usa direttamente RevisioneID
                        Return Path.Combine("C:\VideoEditor\Frames", titolo, $"Revisione_{numero:000}")
                    End If
                End Using
            End Using
        End Using
        Return ""
    End Function

    Private Function CalcolaNumeroRevisionexxxxx(revisioneID As Integer) As Integer
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query As String = "
            SELECT ROW_NUMBER() OVER (PARTITION BY VideoID ORDER BY DataRevisione ASC) - 1 AS Numero
            FROM Mov_Revisioni
            WHERE RevisioneID = @RevisioneID"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@RevisioneID", revisioneID)
                Return CInt(cmd.ExecuteScalar())
            End Using
        End Using
    End Function

    Private Sub dgvRevisioni_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvRevisioni.CellDoubleClick
        If e.RowIndex < 0 Then Exit Sub

        Dim row = dgvRevisioni.Rows(e.RowIndex)
        Dim videoID = CInt(row.Cells("VideoID").Value)
        Dim revisioneID = CInt(row.Cells("RevisioneID").Value)
        Dim autore = row.Cells("Autore").Value.ToString()
        Dim note = row.Cells("Note").Value.ToString()
        Dim stato = row.Cells("Stato").Value.ToString()
        Dim dataRevisione = Convert.ToDateTime(row.Cells("DataRevisione").Value)
        Dim approvato = If(Convert.IsDBNull(row.Cells("Approvato").Value), False, Convert.ToBoolean(row.Cells("Approvato").Value))

        Dim parametri = New RevisioneParametri(videoID, revisioneID, autore, note, stato, dataRevisione, approvato)

        Dim videoForm As VideoFBF = Nothing
        For Each f As Form In Application.OpenForms
            If TypeOf f Is VideoFBF Then
                videoForm = CType(f, VideoFBF)
                Exit For
            End If
        Next

        If videoForm Is Nothing Then
            videoForm = New VideoFBF()
            videoForm.MdiParent = Me.MdiParent
            videoForm.Show()
        End If

        videoForm.Parametri = parametri
        videoForm.AggiornaRevisioneAttiva()
        videoForm.CaricaRevisione(videoID, revisioneID)
        videoForm.AggiornaNoteDaDatabase(revisioneID)
        videoForm.AggiornaUtentiCondivisi(Int(videoForm.lblRevAttiva.Text))
        videoForm.AggiornaFrameCorrente(videoForm.TrackFrame.Value)

        Dim overlay = Me.Controls.Find("OverlayNotePanel", True).FirstOrDefault()
        overlay?.Invalidate()

        Me.Close()
    End Sub

    Private Sub txtFiltro_TextChanged(sender As Object, e As EventArgs) Handles txtFiltro.TextChanged
        Dim filtro As String = txtFiltro.Text.Trim().ToLower()

        Dim dv As DataView

        If TypeOf dgvRevisioni.DataSource Is DataTable Then
            dv = New DataView(CType(dgvRevisioni.DataSource, DataTable))
        ElseIf TypeOf dgvRevisioni.DataSource Is DataView Then
            dv = CType(dgvRevisioni.DataSource, DataView)
        Else
            Exit Sub
        End If

        dv.RowFilter = $"TitoloVideo LIKE '%{filtro}%' OR Stato LIKE '%{filtro}%' OR Note LIKE '%{filtro}%'"
        dgvRevisioni.DataSource = dv
    End Sub

End Class
