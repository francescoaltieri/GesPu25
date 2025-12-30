Imports System.IO
Imports System.Data
Imports System.Linq
Imports System.Threading
Imports Microsoft.Data.SqlClient

Public Class SceltaVideo

    'Public Property RevisioneSelezionata As RevisioneParametri
    Private videoFormDestinazione As VideoFBF

    ' Cache per PercorsoFrames e PercrsoTempFolder
    Private _cachedPercorsoFrames As String = Nothing
    Private _cachedPercorsoFramesLoaded As Boolean = False
    Private _cachedPercorsoTemp As String = Nothing
    Private _cachedPercorsoTempLoaded As Boolean = False
    Private ReadOnly _cacheLock As New Object()

    Public Sub New(destinazione As VideoFBF)
        InitializeComponent()
        videoFormDestinazione = destinazione
    End Sub

    Private Sub SceltaVideo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CaricaRevisioni()
    End Sub

    Private Sub CaricaRevisioni()
        Cursor.Current = Cursors.WaitCursor
        Application.DoEvents()

        Dim nomeUtente As String = SessioneUtente.NomeUtenteCorrente
        Dim dt As New DataTable()

        Using conn As New SqlConnection(ConnString)
            ' Query base
            Dim query As String =
            "SELECT " &
            "    R.RevisioneID, " &
            "    R.DataRevisione, " &
            "    V.VideoID, " &
            "    V.Titolo AS TitoloVideo, " &
            "    R.Autore, " &
            "    R.NumRetake, " &
            "    R.Stato, " &
            "    R.Approvato, " &
            "    R.Note " &
            "FROM Mov_Revisioni R " &
            "INNER JOIN Mov_ConsegneScene V ON R.VideoID = V.VideoID " &
            "INNER JOIN Mov_RevisioniUtente UR ON R.RevisioneID = UR.RevisioneID " &
            "WHERE (UR.NomeUtente = @NomeUtente OR R.Supervisore = @NomeUtente) "

            ' Se il checkbox esiste e è selezionato, aggiungo il filtro "da approvare"
            If ChkDaApprovare IsNot Nothing AndAlso ChkDaApprovare.Checked Then
                query &= "AND (R.Approvato = 0 OR R.Approvato IS NULL) "
            End If

            query &= "ORDER BY V.Titolo, R.DataRevisione ASC;"

            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.Add("@NomeUtente", SqlDbType.NVarChar, 256).Value = nomeUtente
                conn.Open()
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using

        If Not dt.Columns.Contains("NumeroRevisione") Then
            dt.Columns.Add("NumeroRevisione", GetType(String))
        End If

        For Each row As DataRow In dt.Rows
            Dim revisioneID As Integer = CInt(row("RevisioneID"))
            row("NumeroRevisione") = $"Lavorazione_{revisioneID:0000}"
        Next

        dgvRevisioni.DataSource = dt
        If dgvRevisioni.Columns.Contains("VideoID") Then dgvRevisioni.Columns("VideoID").Visible = False
        If dgvRevisioni.Columns.Contains("RevisioneID") Then dgvRevisioni.Columns("RevisioneID").Visible = False
        If dgvRevisioni.Columns.Contains("NumeroRevisione") Then dgvRevisioni.Columns("NumeroRevisione").DisplayIndex = 0

        For Each col As DataGridViewColumn In dgvRevisioni.Columns
            If col.Visible Then col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next

        dgvRevisioni.ReadOnly = True

        Cursor.Current = Cursors.Default
        Application.DoEvents()
    End Sub


    Private Sub dgvRevisioni_KeyDown(sender As Object, e As KeyEventArgs) Handles dgvRevisioni.KeyDown
        If e.KeyCode = Keys.Delete Then
            If dgvRevisioni.SelectedRows.Count = 0 Then Exit Sub

            Dim row = dgvRevisioni.SelectedRows(0)
            Dim revisioneID = CInt(row.Cells("RevisioneID").Value)

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

    ''' <summary>
    ''' Cancella la revisione dal DB e la cartella associata usando la strategia: sposta in temp -> elimina DB in transazione -> elimina temp.
    ''' Se il DB fallisce, ripristina la cartella.
    ''' </summary>
    Private Sub CancellaRevisione(revisioneID As Integer)
        Dim titoloVideo As String = ""
        Dim percorsoVideo As String = ""
        Dim percorsoRevisione As String = ""

        ' Recupera titolo e costruisci percorsi
        Using conn As New SqlConnection(ConnString)
            conn.Open()

            Dim queryInfo As String = "
                                        SELECT V.Titolo
                                        FROM Mov_Revisioni R
                                        INNER JOIN Mov_ConsegneScene V ON R.VideoID = V.VideoID
                                        WHERE R.RevisioneID = @RevisioneID"
            Using cmd As New SqlCommand(queryInfo, conn)
                cmd.Parameters.AddWithValue("@RevisioneID", revisioneID)
                Using reader = cmd.ExecuteReader()
                    If reader.Read() Then
                        titoloVideo = reader("Titolo").ToString().Trim()
                    End If
                End Using
            End Using
        End Using

        Dim baseFrames As String = GetPercorsoFramesCached()
        percorsoVideo = Path.Combine(baseFrames, titoloVideo)
        percorsoRevisione = Path.Combine(percorsoVideo, $"Revisione_{revisioneID:000}")

        ' Se la cartella esiste, prova a spostarla in temp (quarantena)
        Dim tempFolder As String = Nothing
        Dim moved As Boolean = False
        Try
            If Directory.Exists(percorsoRevisione) Then
                Dim tempBase = GetPercorsoTempCached()
                ' crea una sottocartella unica in tempBase
                tempFolder = Path.Combine(tempBase, "VideoFBF_RevisioneBackup_" & Guid.NewGuid().ToString("N"))
                ' assicurati che tempFolder non esista
                If Directory.Exists(tempFolder) Then
                    tempFolder = Path.Combine(tempBase, "VideoFBF_RevisioneBackup_" & Guid.NewGuid().ToString("N"))
                End If
                moved = MoveDirectorySafe(percorsoRevisione, tempFolder)
            End If
        Catch ex As Exception
            moved = False
            Try
                File.AppendAllText(Path.Combine(GetPercorsoTempCached(), "VideoFBF_delete.log"), $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Move to temp failed for {percorsoRevisione}: {ex.Message}{Environment.NewLine}")
            Catch : End Try
        End Try

        ' Esegui cancellazioni DB in transazione
        Dim dbDeleted As Boolean = False
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Using tx = conn.BeginTransaction()
                Try
                    ' Elimina record dipendenti
                    Using cmdDip As New SqlCommand("DELETE FROM Mov_RevisioniUtente WHERE RevisioneID = @RevisioneID", conn, tx)
                        cmdDip.Parameters.AddWithValue("@RevisioneID", revisioneID)
                        cmdDip.ExecuteNonQuery()
                    End Using

                    ' Elimina note associate
                    Using cmdNote As New SqlCommand("DELETE FROM Mov_FrameNote WHERE RevisioneID = @RevisioneID", conn, tx)
                        cmdNote.Parameters.AddWithValue("@RevisioneID", revisioneID)
                        cmdNote.ExecuteNonQuery()
                    End Using

                    ' Elimina revisione
                    Using cmdDel As New SqlCommand("DELETE FROM Mov_Revisioni WHERE RevisioneID = @RevisioneID", conn, tx)
                        cmdDel.Parameters.AddWithValue("@RevisioneID", revisioneID)
                        cmdDel.ExecuteNonQuery()
                    End Using

                    tx.Commit()
                    dbDeleted = True
                Catch ex As Exception
                    Try
                        tx.Rollback()
                    Catch : End Try
                    dbDeleted = False
                    Try
                        File.AppendAllText(Path.Combine(GetPercorsoTempCached(), "VideoFBF_delete.log"), $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - DB delete failed for RevisioneID {revisioneID}: {ex.Message}{Environment.NewLine}")
                    Catch : End Try
                End Try
            End Using
        End Using

        ' Se DB cancellato: elimina definitivamente la cartella temporanea (se spostata) o la cartella originale
        If dbDeleted Then
            Try
                If moved AndAlso Not String.IsNullOrWhiteSpace(tempFolder) AndAlso Directory.Exists(tempFolder) Then
                    SafeDeleteDirectoryRecursive(tempFolder)
                ElseIf Directory.Exists(percorsoRevisione) Then
                    SafeDeleteDirectoryRecursive(percorsoRevisione)
                End If

                ' Se la cartella del video è vuota, prova a rimuoverla
                If Directory.Exists(percorsoVideo) Then
                    If Directory.GetDirectories(percorsoVideo).Length = 0 AndAlso Directory.GetFiles(percorsoVideo).Length = 0 Then
                        Try
                            Directory.Delete(percorsoVideo, True)
                        Catch ex As Exception
                            Try
                                File.AppendAllText(Path.Combine(GetPercorsoTempCached(), "VideoFBF_delete.log"), $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Could not delete empty video folder {percorsoVideo}: {ex.Message}{Environment.NewLine}")
                            Catch : End Try
                        End Try
                    End If
                End If

                ' Aggiorna UI: chiudi VideoFBF se necessario
                For Each f As Form In Application.OpenForms
                    If TypeOf f Is VideoFBF Then
                        Dim videoForm = DirectCast(f, VideoFBF)
                        If videoForm.lblRevAttiva.Text = revisioneID.ToString() Then
                            Try
                                If videoForm.picFrame.Image IsNot Nothing Then
                                    videoForm.picFrame.Image.Dispose()
                                    videoForm.picFrame.Image = Nothing
                                End If
                            Catch : End Try
                            Exit For
                        End If
                    End If
                Next

                MDIMessageBox.Show("Revisione eliminata correttamente.", Me.MdiParent, MessageBoxButtons.OK, "Operazione completata")
                Return
            Catch ex As Exception
                Try
                    File.AppendAllText(Path.Combine(GetPercorsoTempCached(), "VideoFBF_delete.log"), $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Delete folder failed after DB delete for RevisioneID {revisioneID}: {ex.Message}{Environment.NewLine}")
                Catch : End Try
                MDIMessageBox.Show("La revisione è stata rimossa dal database, ma non è stato possibile eliminare completamente la cartella. Controlla i log e rimuovi manualmente: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK, "Attenzione")
                Return
            End Try
        Else
            ' DB non cancellato: se abbiamo spostato la cartella, proviamo a ripristinarla
            If moved AndAlso Not String.IsNullOrWhiteSpace(tempFolder) Then
                Try
                    If Not Directory.Exists(percorsoRevisione) Then
                        Directory.Move(tempFolder, percorsoRevisione)
                    Else
                        ' se la destinazione esiste, sposta in fallback
                        Dim fallback = Path.Combine(GetPercorsoTempCached(), "VideoFBF_RevisioneRestore_" & Guid.NewGuid().ToString("N"))
                        Directory.Move(tempFolder, fallback)
                        Try
                            File.AppendAllText(Path.Combine(GetPercorsoTempCached(), "VideoFBF_delete.log"), $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Temp moved to fallback {fallback}{Environment.NewLine}")
                        Catch : End Try
                    End If
                Catch ex As Exception
                    Try
                        File.AppendAllText(Path.Combine(GetPercorsoTempCached(), "VideoFBF_delete.log"), $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Restore move failed for {tempFolder}: {ex.Message}{Environment.NewLine}")
                    Catch : End Try
                End Try
            End If

            MDIMessageBox.Show("Impossibile eliminare la revisione dal database. Operazione annullata.", Me.MdiParent, MessageBoxButtons.OK, "Errore")
            Return
        End If
    End Sub

    ''' <summary>
    ''' Sposta una directory in modo sicuro anche tra volumi diversi.
    ''' Tenta Directory.Move; se fallisce per radici diverse, copia ricorsivamente e poi elimina la sorgente.
    ''' </summary>
    Public Function MoveDirectorySafe(sourceDir As String, destDir As String, Optional maxAttempts As Integer = 6) As Boolean
        If String.IsNullOrWhiteSpace(sourceDir) OrElse String.IsNullOrWhiteSpace(destDir) Then Return False
        If Not Directory.Exists(sourceDir) Then Return False

        Try
            ' Se la destinazione esiste, fallisci (non sovrascrivere)
            If Directory.Exists(destDir) Then
                Return False
            End If

            ' Primo tentativo: Directory.Move (veloce se nello stesso volume)
            Try
                Directory.Move(sourceDir, destDir)
                Return True
            Catch
                ' prosegui con copia se necessario
            End Try

            ' Controllo radici
            Dim rootSrc = Path.GetPathRoot(Path.GetFullPath(sourceDir))
            Dim rootDst = Path.GetPathRoot(Path.GetFullPath(destDir))
            If String.Equals(rootSrc, rootDst, StringComparison.OrdinalIgnoreCase) Then
                ' stesso volume: riprova con retry
                Dim attempts = 0
                While attempts < maxAttempts
                    Try
                        Directory.Move(sourceDir, destDir)
                        Return True
                    Catch
                        attempts += 1
                        Thread.Sleep(100)
                    End Try
                End While
                Return False
            End If

            ' Radici diverse: copia ricorsiva e poi elimina sorgente
            CopyDirectoryRecursive(sourceDir, destDir, maxAttempts)

            If Not Directory.Exists(destDir) Then Return False

            SafeDeleteDirectoryRecursive(sourceDir, maxAttempts)
            Return True
        Catch ex As Exception
            Try
                File.AppendAllText(Path.Combine(GetPercorsoTempCached(), "VideoFBF_move.log"), $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - MoveDirectorySafe unexpected error for {sourceDir} -> {destDir}: {ex.Message}{Environment.NewLine}")
            Catch : End Try
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Copia ricorsivamente una directory preservando struttura; tenta retry su file bloccati.
    ''' </summary>
    Public Sub CopyDirectoryRecursive(sourceDir As String, destDir As String, Optional maxAttempts As Integer = 6)
        If String.IsNullOrWhiteSpace(sourceDir) OrElse Not Directory.Exists(sourceDir) Then Throw New DirectoryNotFoundException($"Source not found: {sourceDir}")
        If String.IsNullOrWhiteSpace(destDir) Then Throw New ArgumentException("destDir is empty")

        Directory.CreateDirectory(destDir)

        For Each filePath In Directory.GetFiles(sourceDir)
            Dim fileName = Path.GetFileName(filePath)
            Dim destFile = Path.Combine(destDir, fileName)
            Dim attempts = 0
            While attempts < maxAttempts
                Try
                    File.Copy(filePath, destFile, True)
                    Exit While
                Catch
                    attempts += 1
                    Thread.Sleep(100)
                    If attempts >= maxAttempts Then Throw
                End Try
            End While
        Next

        For Each dirPath In Directory.GetDirectories(sourceDir)
            Dim dirName = Path.GetFileName(dirPath)
            Dim destSub = Path.Combine(destDir, dirName)
            CopyDirectoryRecursive(dirPath, destSub, maxAttempts)
        Next
    End Sub

    ''' <summary>
    ''' Cancella ricorsivamente una directory con retry su file bloccati.
    ''' </summary>
    Public Sub SafeDeleteDirectoryRecursive(dir As String, Optional maxAttempts As Integer = 6)
        If String.IsNullOrWhiteSpace(dir) OrElse Not Directory.Exists(dir) Then Return

        Try
            For Each f In Directory.GetFiles(dir, "*", SearchOption.AllDirectories)
                Dim attempts = 0
                While attempts < maxAttempts
                    Try
                        File.SetAttributes(f, FileAttributes.Normal)
                        File.Delete(f)
                        Exit While
                    Catch
                        attempts += 1
                        Thread.Sleep(100)
                    End Try
                End While
            Next

            Dim attemptsDir = 0
            While attemptsDir < maxAttempts
                Try
                    Directory.Delete(dir, True)
                    Exit While
                Catch
                    attemptsDir += 1
                    Thread.Sleep(200)
                End Try
            End While

            If Directory.Exists(dir) Then
                Try
                    File.AppendAllText(Path.Combine(GetPercorsoTempCached(), "VideoFBF_delete.log"), $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - SafeDeleteDirectoryRecursive: directory still exists {dir}{Environment.NewLine}")
                Catch : End Try
            End If
        Catch ex As Exception
            Try
                File.AppendAllText(Path.Combine(GetPercorsoTempCached(), "VideoFBF_delete.log"), $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - SafeDeleteDirectoryRecursive error for {dir}: {ex.Message}{Environment.NewLine}")
            Catch : End Try
        End Try
    End Sub

    ' Lettura parametro da Sys_Parametri
    Private Function GetSysParametro(descrizione As String) As String
        Try
            Using conn As New SqlConnection(ConnString)
                conn.Open()
                Using cmd As New SqlCommand("SELECT Valore FROM Sys_Parametri WHERE Descrizione = @Descrizione", conn)
                    cmd.Parameters.AddWithValue("@Descrizione", descrizione)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not Convert.IsDBNull(result) Then
                        Return result.ToString().Trim()
                    End If
                End Using
            End Using
        Catch ex As Exception
            ' Log su temp (fallback)
            Try
                Dim fallback = Path.GetTempPath()
                File.AppendAllText(Path.Combine(fallback, "VideoFBF_params.log"), $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - GetSysParametro error for '{descrizione}': {ex.Message}{Environment.NewLine}")
            Catch : End Try
        End Try
        Return String.Empty
    End Function

    ' PercorsoFrames con fallback
    Private Function GetPercorsoFrames() As String
        Dim valore = GetSysParametro("PercorsoFrames")
        If String.IsNullOrWhiteSpace(valore) Then
            Return "C:\VideoEditor\Frames"
        End If
        Return valore
    End Function

    ' Versione cached per evitare query ripetute
    Private Function GetPercorsoFramesCached() As String
        If _cachedPercorsoFramesLoaded Then
            Return _cachedPercorsoFrames
        End If

        SyncLock _cacheLock
            If Not _cachedPercorsoFramesLoaded Then
                Try
                    _cachedPercorsoFrames = GetPercorsoFrames()
                    _cachedPercorsoFramesLoaded = True
                Catch
                    _cachedPercorsoFrames = "C:\VideoEditor\Frames"
                    _cachedPercorsoFramesLoaded = True
                End Try
            End If
        End SyncLock

        Return _cachedPercorsoFrames
    End Function

    ' Percorso temporaneo parametrizzato (Descrizione = PercrsoTempFolder)
    Private Function GetPercorsoTemp() As String
        ' Nota: usa la stringa esatta che hai indicato nel DB: "PercrsoTempFolder"
        Dim valore = GetSysParametro("PercrsoTempFolder")
        If String.IsNullOrWhiteSpace(valore) Then
            Return Path.GetTempPath()
        End If
        Return valore
    End Function

    ' Cached version for temp path
    Private Function GetPercorsoTempCached() As String
        If _cachedPercorsoTempLoaded Then
            Return _cachedPercorsoTemp
        End If

        SyncLock _cacheLock
            If Not _cachedPercorsoTempLoaded Then
                Try
                    Dim p = GetPercorsoTemp()
                    ' assicurati che la cartella esista; se non esiste prova a crearla
                    If Not String.IsNullOrWhiteSpace(p) Then
                        Try
                            If Not Directory.Exists(p) Then Directory.CreateDirectory(p)
                            _cachedPercorsoTemp = p
                        Catch
                            _cachedPercorsoTemp = Path.GetTempPath()
                        End Try
                    Else
                        _cachedPercorsoTemp = Path.GetTempPath()
                    End If
                Catch
                    _cachedPercorsoTemp = Path.GetTempPath()
                End Try
                _cachedPercorsoTempLoaded = True
            End If
        End SyncLock

        Return _cachedPercorsoTemp
    End Function

    Private Sub dgvRevisioni_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvRevisioni.CellDoubleClick
        Cursor.Current = Cursors.WaitCursor
        Application.DoEvents()

        If e.RowIndex < 0 Then Exit Sub

        Dim row = dgvRevisioni.Rows(e.RowIndex)
        Dim videoID = CInt(row.Cells("VideoID").Value)
        Dim revisioneID = CInt(row.Cells("RevisioneID").Value)
        Dim autore = If(row.Cells("Autore").Value IsNot Nothing, row.Cells("Autore").Value.ToString(), String.Empty)
        Dim note = If(row.Cells("Note").Value IsNot Nothing, row.Cells("Note").Value.ToString(), String.Empty)
        Dim stato = If(row.Cells("Stato").Value IsNot Nothing, row.Cells("Stato").Value.ToString(), String.Empty)
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
        videoForm.AggiornaUtentiCondivisi(CInt(videoForm.lblRevAttiva.Text))
        videoForm.AggiornaFrameCorrente(videoForm.TrackFrame.Value)

        Dim overlay = Me.Controls.Find("OverlayNotePanel", True).FirstOrDefault()
        overlay?.Invalidate()

        Cursor.Current = Cursors.Default
        Application.DoEvents()
        Me.Close()
    End Sub

    Private Sub txtFiltro_TextChanged(sender As Object, e As EventArgs) Handles txtFiltroLavorazione.TextChanged
        Dim filtro = txtFiltroLavorazione.Text.Trim.ToLower
        Dim dv As DataView = Nothing

        If TypeOf dgvRevisioni.DataSource Is DataTable Then
            dv = New DataView(CType(dgvRevisioni.DataSource, DataTable))
        ElseIf TypeOf dgvRevisioni.DataSource Is DataView Then
            dv = CType(dgvRevisioni.DataSource, DataView)
        Else
            Exit Sub
        End If

        dv.RowFilter = $"NumeroRevisione LIKE '%{filtro}%' OR Stato LIKE '%{filtro}%' OR Note LIKE '%{filtro}%' OR Autore LIKE '%{filtro}%' OR TitoloVideo LIKE '%{filtro}%'"
        dgvRevisioni.DataSource = dv
    End Sub

    Private Sub ChkDaApprovare_CheckedChanged(sender As Object, e As EventArgs) Handles ChkDaApprovare.CheckedChanged
        Try
            Cursor.Current = Cursors.WaitCursor
            Application.DoEvents()
            CaricaRevisioni()
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub
End Class
