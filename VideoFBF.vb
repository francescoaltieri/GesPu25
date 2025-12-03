Imports System.IO
Imports Microsoft.Data.SqlClient
Imports System.Threading.Tasks
Imports System.Diagnostics

Public Class VideoFBF
    ' --- Campi di stato e UI ---
    Dim editor As VideoEditor
    Dim isDrawing As Boolean = False
    Dim lastPoint As Point
    Dim autoScrollActive As Boolean = False
    Dim autoScrollDirection As String = "" ' "forward" o "backward"
    Dim testoAvantiOriginale As String = "Avanti Veloce"
    Dim testoIndietroOriginale As String = "Indietro Veloce"
    Dim notaPosizione As Point = Point.Empty
    Dim colorePennino As Color = Color.Red
    Dim spessorePennino As Integer = 5
    Dim disegnoAttivo As Boolean = False
    Private FrameConNote As New List(Of Integer)
    Private aggiornamentoInCorso As Boolean = False
    Private hasUnsavedChanges As Boolean = False

    Public Property Parametri As RevisioneParametri

    ' --- Costruttori e Load ---
    Public Sub New(parametri As RevisioneParametri)
        InitializeComponent()
        Me.Parametri = parametri
    End Sub

    Public Sub New()
        InitializeComponent()
        Me.Parametri = Nothing
    End Sub

    Private Sub VideoFBF_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RipristinaPosizioneForm(Me)
        lblRevAttiva.Text = "Nessuna revisione attiva"
        CaricaUtentiDisponibili()
        TrackFrame.Enabled = False
        btnSuccessivo.Enabled = False
        btnPrecedente.Enabled = False
        btnPrimoFrame.Enabled = False
        btnUltimoFrame.Enabled = False
        btnAvantiVeloce.Enabled = False
        btnIndietroVeloce.Enabled = False
    End Sub

    ' --- Utility TrackBar sicura ---
    Private Sub SafeSetTrackFrameValue(desired As Integer)
        If TrackFrame Is Nothing Then Return
        If TrackFrame.Minimum > TrackFrame.Maximum Then
            TrackFrame.Minimum = 0
            TrackFrame.Maximum = 0
        End If
        If TrackFrame.Maximum < TrackFrame.Minimum Then
            TrackFrame.Enabled = False
            Return
        End If
        Dim safeValue As Integer = Math.Min(Math.Max(desired, TrackFrame.Minimum), TrackFrame.Maximum)
        RemoveHandler TrackFrame.Scroll, AddressOf trackFrame_Scroll
        Try
            TrackFrame.Value = safeValue
        Finally
            AddHandler TrackFrame.Scroll, AddressOf trackFrame_Scroll
        End Try

        UpdateFrameLabels()

    End Sub

    ' --- Caricamento NUOVO VIDEO: con ID revisione reale e backup NomeVideo_0000 ---
    Private Async Sub btnCaricaVideo_Click(sender As Object, e As EventArgs) Handles btnCaricaVideo.Click
        Dim previousUseWait As Boolean = Application.UseWaitCursor
        Dim previousCursor As Cursor = Me.Cursor

        Try
            OpenFileDialog1.Filter = "Video Files|*.mp4;*.mov"
            If OpenFileDialog1.ShowDialog() <> DialogResult.OK Then Exit Sub

            btnCaricaVideo.Enabled = False
            lblRevAttiva.Text = "Caricamento in corso..."
            Application.DoEvents()

            Application.UseWaitCursor = True
            Me.Cursor = Cursors.WaitCursor
            Cursor.Current = Cursors.WaitCursor

            Dim videoPath = OpenFileDialog1.FileName
            Dim nomeVideo = Path.GetFileNameWithoutExtension(videoPath)
            Dim baseDir = Path.Combine(OttieniPercorsoFrames(), nomeVideo)
            Directory.CreateDirectory(baseDir)

            Dim approvato As Boolean = False

            ' 1) Registra/recupera VideoID e garantisci Revisione 0 una sola volta
            Dim videoID = Await Task.Run(Function()
                                             Dim id = OttieniVideoID(nomeVideo)
                                             If id = -1 Then
                                                 id = InserisciVideo(nomeVideo, videoPath)
                                                 If Not RevisioneZeroEsiste(id) Then
                                                     InserisciRevisioneZero(id, nomeVideo)
                                                 End If
                                             End If
                                             Return id
                                         End Function)

            ' 2) Crea revisione: ottieni ID reale dal DB
            Dim newRevisioneID As Integer = Await Task.Run(Function()
                                                               Return InserisciRevisione_RitornaID(videoID, SessioneUtente.NomeUtenteCorrente, "bozza")
                                                           End Function)

            ' 3) Cartella della revisione reale
            Dim revisioneDir = Path.Combine(baseDir, $"Revisione_{newRevisioneID:D4}")
            Directory.CreateDirectory(revisioneDir)

            ' 4) Estrai i frame nella cartella revisione
            Dim tempEditor As New VideoEditor(videoPath, revisioneDir)
            Try
                Await Task.Run(Sub() tempEditor.ExtractFrames())
            Catch exInner As Exception
                MessageBox.Show("Errore durante l'estrazione dei frame: " & exInner.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End Try

            ' 5) Verifica frame estratti
            Dim filesEstratti = Directory.GetFiles(revisioneDir)
            If filesEstratti Is Nothing OrElse filesEstratti.Length = 0 Then
                MessageBox.Show($"Estrazione completata ma nessun frame trovato in: {revisioneDir}", "Nessun frame", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If

            ' 6) CREA/POPOLA Revisione_0000 in background (non blocca UI)
            Try
                Await Task.Run(Sub()
                                   EnsureOriginalBackupFolder(nomeVideo, revisioneDir)
                               End Sub)
            Catch ex As Exception
                MessageBox.Show("Attenzione: impossibile creare backup Revisione_0000: " & ex.Message)
            End Try

            ' 7) Inserisci permesso utente con FK coerente
            Await Task.Run(Sub() InserisciPermessoUtente(newRevisioneID, SessioneUtente.NomeUtenteCorrente))

            MessageBox.Show($"Video caricato, frame estratti, Revisione_{newRevisioneID:D4} registrata e backup {nomeVideo}_0000 creato.", "Operazione completata", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' 8) Aggiorna UI
            lblRevAttiva.Text = newRevisioneID.ToString()

            Dim parametri = New RevisioneParametri(
                videoID,
                newRevisioneID,
                SessioneUtente.NomeUtenteCorrente,
                $"Revisione {newRevisioneID}",
                "visualizza",
                DateTime.Now,
                approvato
            )

            editor = tempEditor
            editor.CurrentIndex = 0
            picFrame.Image = editor.LoadFrame(0)

            RemoveHandler TrackFrame.Scroll, AddressOf trackFrame_Scroll
            TrackFrame.Minimum = 0
            TrackFrame.Maximum = Math.Max(0, editor.FrameList.Count - 1)
            SafeSetTrackFrameValue(0)
            AddHandler TrackFrame.Scroll, AddressOf trackFrame_Scroll

            TrackFrame.Enabled = True
            btnSuccessivo.Enabled = True
            btnPrecedente.Enabled = True
            btnPrimoFrame.Enabled = True
            btnUltimoFrame.Enabled = True
            btnAvantiVeloce.Enabled = True
            btnIndietroVeloce.Enabled = True

            Me.Parametri = parametri
            Me.AggiornaRevisioneAttiva()

        Catch ex As Exception
            MessageBox.Show("Errore durante il caricamento: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            lblRevAttiva.Text = "Errore"
        Finally
            Application.UseWaitCursor = previousUseWait
            Me.Cursor = previousCursor
            Cursor.Current = previousCursor
            btnCaricaVideo.Enabled = True
            Application.DoEvents()
        End Try
    End Sub

    ' Crea/Popola la cartella di backup fissa Revisione_0000 con i frame originali
    Private Sub EnsureOriginalBackupFolder(nomeVideo As String, revisioneDir As String)
        Dim baseDir = Path.Combine(OttieniPercorsoFrames(), nomeVideo)
        Dim originalDir = Path.Combine(baseDir, "Revisione_0000")

        ' Crea la cartella se non esiste
        Directory.CreateDirectory(originalDir)

        ' Se è già popolata (almeno un file) non ricopiare
        Dim existing = Directory.GetFiles(originalDir)
        If existing IsNot Nothing AndAlso existing.Length > 0 Then
            Return
        End If

        ' Prendi i frame estratti nella revisione corrente
        If Not Directory.Exists(revisioneDir) Then
            Throw New DirectoryNotFoundException($"Cartella revisione non trovata: {revisioneDir}")
        End If

        Dim framesEstratti = Directory.GetFiles(revisioneDir, "frame_*.*", SearchOption.TopDirectoryOnly).
                                OrderBy(Function(f) f).ToArray()

        If framesEstratti.Length = 0 Then
            ' niente da copiare
            Return
        End If

        For Each src In framesEstratti
            Dim dst = Path.Combine(originalDir, Path.GetFileName(src))
            Try
                ' copia senza sovrascrivere file già presenti
                If Not File.Exists(dst) Then
                    File.Copy(src, dst, overwrite:=False)
                End If
            Catch ex As Exception
                ' ignora errori singoli ma continua; puoi loggare se vuoi
            End Try
        Next
    End Sub

    ' Ripristina il frame corrente dall’originale cancellando il file overlay
    Private Sub RestoreCurrentFrameFromOriginal()
        If editor Is Nothing Then Return

        Dim idx = editor.CurrentIndex
        If idx < 0 OrElse idx >= editor.FrameList.Count Then
            MessageBox.Show("Indice frame non valido.", "Ripristino frame", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim dst = editor.FrameList(idx)
        Dim baseNameNoExt = Path.GetFileNameWithoutExtension(dst)
        Dim overlayPath = Path.Combine(Path.GetDirectoryName(dst), baseNameNoExt & "_overlay.png")

        Dim logPath = Path.Combine(Path.GetTempPath(), "VideoFBF_restore.log")
        Try
            ' 1) Rilascia riferimenti UI/editor che potrebbero lockare il file overlay
            Try
                If picFrame.Image IsNot Nothing Then
                    Try : picFrame.Image.Dispose() : Catch : End Try
                    picFrame.Image = Nothing
                End If
            Catch
            End Try

            Try
                If editor.DrawingBitmap IsNot Nothing Then
                    Try : editor.DrawingBitmap.Dispose() : Catch : End Try
                    editor.DrawingBitmap = Nothing
                End If
            Catch
            End Try

            Try
                If editor.UndoStack IsNot Nothing Then
                    For Each b In editor.UndoStack
                        Try : b.Dispose() : Catch : End Try
                    Next
                    editor.UndoStack.Clear()
                End If
            Catch
            End Try

            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()

            ' 2) Elimina il file overlay se esiste (retry semplice)
            If File.Exists(overlayPath) Then
                Dim attempts As Integer = 0
                Dim deleted As Boolean = False
                Dim lastEx As Exception = Nothing
                While attempts < 6 AndAlso Not deleted
                    attempts += 1
                    Try
                        File.Delete(overlayPath)
                        deleted = True
                    Catch ex As IOException
                        lastEx = ex
                        Threading.Thread.Sleep(100 * attempts)
                    Catch ex As Exception
                        lastEx = ex
                        Exit While
                    End Try
                End While

                If Not deleted Then
                    Try : File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Warning: unable to delete overlay {overlayPath}: {If(lastEx IsNot Nothing, lastEx.Message, "unknown")}{Environment.NewLine}") : Catch : End Try
                    MessageBox.Show("Impossibile eliminare il file overlay. Riprova più tardi.", "Ripristino", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    ' procediamo comunque a pulire lo stato in memoria e ricaricare il frame base
                End If
            End If

            ' 3) Pulisci annotazioni in memoria per questo frame
            Try
                editor.ClearFrameAnnotations(idx)
            Catch
                Try
                    If editor.FrameNote IsNot Nothing AndAlso editor.FrameNote.ContainsKey(idx) Then
                        editor.FrameNote.Remove(idx)
                    End If
                Catch
                End Try
            End Try

            If FrameConNote IsNot Nothing AndAlso FrameConNote.Contains(idx) Then
                Try : FrameConNote.Remove(idx) : Catch : End Try
            End If

            ' 4) Aggiorna UI delle note (se la revisione è gestita dal DB, ricarica; altrimenti refresh in memoria)
            Try
                Dim revID As Integer
                If TryParseRevisioneID(lblRevAttiva.Text, revID) Then
                    AggiornaNoteDaDatabase(revID)
                Else
                    RefreshLstNoteFrame()
                End If
            Catch
            End Try

            ' 5) Ricarica il frame (ora il file base è il frame "originale" della revisione attiva)
            Try
                If editor IsNot Nothing Then
                    If picFrame.Image IsNot Nothing Then
                        Try : picFrame.Image.Dispose() : Catch : End Try
                        picFrame.Image = Nothing
                    End If
                    picFrame.Image = editor.LoadFrame(idx)
                Else
                    picFrame.Image = LoadImageWithoutLock(dst)
                End If
            Catch ex As Exception
                Try
                    If picFrame.Image IsNot Nothing Then
                        Try : picFrame.Image.Dispose() : Catch : End Try
                        picFrame.Image = Nothing
                    End If
                    picFrame.Image = LoadImageWithoutLock(dst)
                Catch
                End Try
            End Try

            hasUnsavedChanges = False
            Try : editor.HasUnsavedChanges = False : Catch : End Try

            Dim overlayCtrlFinal = Me.Controls.Find("OverlayNotePanel", True).FirstOrDefault()
            If overlayCtrlFinal IsNot Nothing Then overlayCtrlFinal.Invalidate()

            MessageBox.Show("Frame ripristinato: overlay eliminato e appunti rimossi in memoria.", "Ripristino completato", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            Try : File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - ERRORE: {ex.Message}{Environment.NewLine}{ex.StackTrace}{Environment.NewLine}") : Catch : End Try
            MessageBox.Show("Errore durante il ripristino del frame: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Try : UpdateRipristinaButton() : Catch : End Try
            Try
                Dim revID As Integer
                If TryParseRevisioneID(lblRevAttiva.Text, revID) Then
                    AggiornaNoteDaDatabase(revID)
                Else
                    RefreshLstNoteFrame()
                End If
            Catch
            End Try
        End Try
    End Sub

    ' Aggiorna lblFrameAttivo e lblTotFrames in modo thread-safe
    Private Sub UpdateFrameLabels(Optional currentIndex As Integer = -1)
        Try
            If Me.InvokeRequired Then
                Me.Invoke(New Action(Of Integer)(AddressOf UpdateFrameLabels), currentIndex)
                Return
            End If

            Dim total As Integer = 0
            Try
                If editor IsNot Nothing AndAlso editor.FrameList IsNot Nothing Then
                    total = editor.FrameList.Count
                End If
            Catch
                total = 0
            End Try

            ' Se non è stato passato un indice, prendi l'indice corrente dall'editor
            If currentIndex < 0 Then
                Try
                    If editor IsNot Nothing Then currentIndex = editor.CurrentIndex
                Catch
                    currentIndex = -1
                End Try
            End If

            ' Visualizziamo 1-based per l'utente; se non ci sono frame mostriamo 0/0 o "—"
            If total <= 0 Then
                lblFrameAttivo.Text = "0"
                lblTotFrames.Text = "0"
            Else
                Dim displayIndex As Integer = 0
                If currentIndex >= 0 AndAlso currentIndex < total Then
                    displayIndex = currentIndex + 1
                Else
                    ' se indice non valido, mostra 1 come fallback
                    displayIndex = 1
                End If
                lblFrameAttivo.Text = displayIndex.ToString()
                lblTotFrames.Text = total.ToString()
            End If
        Catch
            ' non interrompere l'app per un problema di UI
        End Try
    End Sub

    ' Carica un'immagine da file senza tenere il file lockato (usa MemoryStream)
    Private Function LoadImageWithoutLock(path As String) As Image
        ' Carica l'immagine in memoria e restituisce una copia indipendente
        Using fs As New FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read)
            Using ms As New MemoryStream()
                fs.CopyTo(ms)
                ms.Position = 0
                Using tmpImg As Image = Image.FromStream(ms)
                    Dim result As Image = New Bitmap(tmpImg)
                    Return result
                End Using
            End Using
        End Using
    End Function

    ' --- Inserire revisione ritornando ID reale (OUTPUT INSERTED) ---
    Private Function InserisciRevisione_RitornaID(videoId As Integer, nomeUtente As String, stato As String) As Integer
        Dim nomeRevisione = $"Revisione - {DateTime.Now:dd/MM/yyyy}"
        Dim NumRetake = CalcolaRetake(videoId)

        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Using tran = conn.BeginTransaction()
                Try
                    Dim query = "
                        INSERT INTO Mov_Revisioni (VideoID, Autore, DataRevisione, NumRetake, Note, Stato)
                        OUTPUT INSERTED.RevisioneID
                        VALUES (@VideoID, @Autore, @Data, @NumRetake, @Note, @Stato)"
                    Using cmd As New SqlCommand(query, conn, tran)
                        cmd.Parameters.Add("@VideoID", SqlDbType.Int).Value = videoId
                        cmd.Parameters.Add("@Autore", SqlDbType.NVarChar, 100).Value = nomeUtente
                        cmd.Parameters.Add("@Data", SqlDbType.DateTime).Value = DateTime.Now
                        cmd.Parameters.Add("@NumRetake", SqlDbType.Int).Value = NumRetake
                        cmd.Parameters.Add("@Note", SqlDbType.NVarChar, 500).Value = nomeRevisione
                        cmd.Parameters.Add("@Stato", SqlDbType.NVarChar, 50).Value = stato
                        Dim newIdObj = cmd.ExecuteScalar()
                        Dim newId = CInt(newIdObj)
                        tran.Commit()
                        Return newId
                    End Using
                Catch
                    tran.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Function

    ' --- UTENTI condivisi: cambi di tabella (Mov_RevisioniUtente) ---
    Private Sub CaricaUtentiDisponibili()
        lstUtentiCondivisi.Items.Clear()
        Using conn As New SqlConnection(ConnString)
            Dim cmd As New SqlCommand("
                SELECT NomeUtente, Generalità 
                FROM Sys_Utenti 
                WHERE IsActive = 1 AND (Amministratore = 1 OR Supervisore = 1)
                ORDER BY Generalità", conn)
            conn.Open()
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim nomeUtente = reader("NomeUtente").ToString()
                    Dim generalita = reader("Generalità").ToString()
                    If nomeUtente <> SessioneUtente.NomeUtenteCorrente Then
                        lstUtentiCondivisi.Items.Add(New KeyValuePair(Of String, String)(nomeUtente, generalita), False)
                    End If
                End While
            End Using
        End Using
    End Sub

    Private Sub lstUtentiCondivisi_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles lstUtentiCondivisi.ItemCheck
        If aggiornamentoInCorso Then Exit Sub
        Me.BeginInvoke(Sub()
                           Dim item = CType(lstUtentiCondivisi.Items(e.Index), KeyValuePair(Of String, String))
                           Dim nomeUtente = item.Key
                           Dim revisioneID As Integer
                           If Not TryParseRevisioneID(lblRevAttiva.Text, revisioneID) Then
                               MessageBox.Show("Revisione non valida.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
                               Return
                           End If
                           If e.NewValue = CheckState.Checked Then
                               AggiungiCondivisioneUtente(revisioneID, nomeUtente)
                           ElseIf e.NewValue = CheckState.Unchecked Then
                               RimuoviCondivisioneUtente(revisioneID, nomeUtente)
                           End If
                       End Sub)
    End Sub

    Public Sub AggiornaUtentiCondivisi(revisioneID As Integer)
        aggiornamentoInCorso = True
        For i = 0 To lstUtentiCondivisi.Items.Count - 1
            lstUtentiCondivisi.SetItemChecked(i, False)
        Next
        Using conn As New SqlConnection(ConnString)
            Dim cmd As New SqlCommand("
                SELECT NomeUtente 
                FROM Mov_RevisioniUtente 
                WHERE RevisioneID = @RevID", conn)
            cmd.Parameters.Add("@RevID", SqlDbType.Int).Value = revisioneID
            conn.Open()
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim nome = reader("NomeUtente").ToString()
                    For i = 0 To lstUtentiCondivisi.Items.Count - 1
                        Dim item = CType(lstUtentiCondivisi.Items(i), KeyValuePair(Of String, String))
                        If item.Key = nome Then
                            lstUtentiCondivisi.SetItemChecked(i, True)
                            Exit For
                        End If
                    Next
                End While
            End Using
        End Using
        aggiornamentoInCorso = False
    End Sub

    Private Sub AggiungiCondivisioneUtente(revisioneID As Integer, nomeUtente As String)
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim checkCmd As New SqlCommand("
                SELECT COUNT(*) 
                FROM Mov_RevisioniUtente 
                WHERE RevisioneID = @RevID AND NomeUtente = @NomeUtente", conn)
            checkCmd.Parameters.Add("@RevID", SqlDbType.Int).Value = revisioneID
            checkCmd.Parameters.Add("@NomeUtente", SqlDbType.NVarChar, 100).Value = nomeUtente
            Dim esiste = Convert.ToInt32(checkCmd.ExecuteScalar()) > 0
            If esiste Then Exit Sub
            Dim insertCmd As New SqlCommand("
                INSERT INTO Mov_RevisioniUtente (RevisioneID, NomeUtente, Permesso)
                VALUES (@RevID, @NomeUtente, 'visualizza')", conn)
            insertCmd.Parameters.Add("@RevID", SqlDbType.Int).Value = revisioneID
            insertCmd.Parameters.Add("@NomeUtente", SqlDbType.NVarChar, 100).Value = nomeUtente
            insertCmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub RimuoviCondivisioneUtente(revisioneID As Integer, nomeUtente As String)
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim cmd As New SqlCommand("
                DELETE FROM Mov_RevisioniUtente 
                WHERE RevisioneID = @RevID AND NomeUtente = @NomeUtente", conn)
            cmd.Parameters.Add("@RevID", SqlDbType.Int).Value = revisioneID
            cmd.Parameters.Add("@NomeUtente", SqlDbType.NVarChar, 100).Value = nomeUtente
            cmd.ExecuteNonQuery()
        End Using
    End Sub
    ' --- FormClosing e salvataggio posizione ---
    Private Sub VideoFBF_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If hasUnsavedChanges Then
            Dim proceed = ConfirmSaveChanges()
            If Not proceed Then
                e.Cancel = True
                Return
            End If
        End If
        SalvaPosizioneForm(Me)
    End Sub

    Private Sub SalvaPosizioneForm2()
        Dim stato = If(Me.WindowState = FormWindowState.Maximized, "Maximized", "Normal")
        Dim x = If(Me.WindowState = FormWindowState.Normal, Me.Location.X, RestoreBounds.X)
        Dim y = If(Me.WindowState = FormWindowState.Normal, Me.Location.Y, RestoreBounds.Y)
        Dim w = If(Me.WindowState = FormWindowState.Normal, Me.Size.Width, RestoreBounds.Width)
        Dim h = If(Me.WindowState = FormWindowState.Normal, Me.Size.Height, RestoreBounds.Height)

        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query = "
                IF EXISTS (SELECT 1 FROM Sys_Form WHERE FormName = @FormName)
                    UPDATE Sys_Form SET X = @X, Y = @Y, Width = @Width, Height = @Height, WindowsState = @WindowsState WHERE FormName = @FormName
                ELSE
                    INSERT INTO Sys_Form (FormName, X, Y, Width, Height, WindowsState) VALUES (@FormName, @X, @Y, @Width, @Height, @WindowsState)"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.Add("@FormName", SqlDbType.NVarChar, 200).Value = Me.Name
                cmd.Parameters.Add("@X", SqlDbType.Int).Value = x
                cmd.Parameters.Add("@Y", SqlDbType.Int).Value = y
                cmd.Parameters.Add("@Width", SqlDbType.Int).Value = w
                cmd.Parameters.Add("@Height", SqlDbType.Int).Value = h
                cmd.Parameters.Add("@WindowsState", SqlDbType.NVarChar, 50).Value = stato
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ' --- Parametri e DB helpers ---
    Private Function OttieniPercorsoFrames() As String
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query As String = "SELECT Valore FROM Sys_Parametri WHERE Descrizione = @DescPar"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.Add("@DescPar", SqlDbType.NVarChar, 200).Value = "PercorsoFrames"
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing Then
                    Return result.ToString()
                End If
            End Using
        End Using
        Return ""
    End Function

    Private Function OttieniVideoID(nomeVideo As String) As Integer
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query = "SELECT VideoID FROM Mov_ConsegneScene WHERE Titolo = @Titolo"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.Add("@Titolo", SqlDbType.NVarChar, 200).Value = nomeVideo
                Dim result = cmd.ExecuteScalar()
                Return If(result IsNot Nothing, CInt(result), -1)
            End Using
        End Using
    End Function

    Private Function InserisciVideo(nomeVideo As String, percorsoFile As String) As Integer
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim checkQuery = "SELECT VideoID FROM Mov_ConsegneScene WHERE Titolo = @Titolo"
            Using checkCmd As New SqlCommand(checkQuery, conn)
                checkCmd.Parameters.Add("@Titolo", SqlDbType.NVarChar, 200).Value = nomeVideo
                Dim result = checkCmd.ExecuteScalar()
                If result IsNot Nothing Then
                    Return CInt(result)
                End If
            End Using
            Dim insertQuery = "
                INSERT INTO Mov_ConsegneScene (Titolo, CreatoDa, FileScena, DataCreazione)
                OUTPUT INSERTED.VideoID
                VALUES (@Titolo, @CreatoDa, @FileScena, @DataCreazione)"
            Using insertCmd As New SqlCommand(insertQuery, conn)
                insertCmd.Parameters.Add("@Titolo", SqlDbType.NVarChar, 200).Value = nomeVideo
                insertCmd.Parameters.Add("@CreatoDa", SqlDbType.NVarChar, 100).Value = SessioneUtente.NomeUtenteCorrente
                insertCmd.Parameters.Add("@FileScena", SqlDbType.NVarChar, 1000).Value = percorsoFile
                insertCmd.Parameters.Add("@DataCreazione", SqlDbType.DateTime).Value = DateTime.Now
                Return CInt(insertCmd.ExecuteScalar())
            End Using
        End Using
    End Function

    Private Function RevisioneZeroEsiste(videoID As Integer) As Boolean
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query = "
                SELECT COUNT(*) 
                FROM Mov_Revisioni 
                WHERE VideoID = @VideoID AND Note LIKE 'Revisione 0%'"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.Add("@VideoID", SqlDbType.Int).Value = videoID
                Return CInt(cmd.ExecuteScalar()) > 0
            End Using
        End Using
    End Function

    Private Function InserisciRevisioneZero(videoID As Integer, nomeVideo As String) As Integer
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query = "
                INSERT INTO Mov_Revisioni (RevisioneID, VideoID, Autore, DataRevisione, Note, Stato)
                VALUES (@RevisioneID, @VideoID, @Autore, @Data, @Note, @Stato)"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.Add("@RevisioneID", SqlDbType.Int).Value = 0
                cmd.Parameters.Add("@VideoID", SqlDbType.Int).Value = videoID
                cmd.Parameters.Add("@Autore", SqlDbType.NVarChar, 100).Value = SessioneUtente.NomeUtenteCorrente
                cmd.Parameters.Add("@Data", SqlDbType.DateTime).Value = DateTime.Now
                cmd.Parameters.Add("@Note", SqlDbType.NVarChar, 500).Value = "Revisione 0 - " & nomeVideo
                cmd.Parameters.Add("@Stato", SqlDbType.NVarChar, 50).Value = "visualizza"
                cmd.ExecuteNonQuery()
            End Using
        End Using
        Return 0
    End Function

    Private Sub InserisciPermessoUtente(revisioneID As Integer, nomeUtente As String)
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query = "
                INSERT INTO Mov_RevisioniUtente (RevisioneID, NomeUtente, Permesso)
                VALUES (@RevID, @NomeUtente, 'visualizza')"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.Add("@RevID", SqlDbType.Int).Value = revisioneID
                cmd.Parameters.Add("@NomeUtente", SqlDbType.NVarChar, 100).Value = nomeUtente
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Function CalcolaRetake(videoID As Integer) As Integer
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query As String = "SELECT COUNT(*) FROM Mov_Revisioni WHERE VideoID = @VideoID AND RevisioneID > 0"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.Add("@VideoID", SqlDbType.Int).Value = videoID
                Return CInt(cmd.ExecuteScalar())
            End Using
        End Using
    End Function

    Private Function OttieniProssimoRevisioneIDPerVideo(videoID As Integer) As Integer
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query = "SELECT ISNULL(MAX(RevisioneID), 0) + 1 FROM Mov_Revisioni WHERE VideoID = @VideoID"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.Add("@VideoID", SqlDbType.Int).Value = videoID
                Return CInt(cmd.ExecuteScalar())
            End Using
        End Using
    End Function

    Private Function OttieniProssimoRevisioneID() As Integer
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query = "SELECT ISNULL(MAX(RevisioneID), 0) + 1 FROM Mov_Revisioni"
            Using cmd As New SqlCommand(query, conn)
                Return CInt(cmd.ExecuteScalar())
            End Using
        End Using
    End Function

    Public Function OttieniNumeroRevisione() As Integer
        Dim numero As Integer = 0
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query As String = "SELECT ISNULL(MAX(RevisioneID), -1) FROM Mov_Revisioni"
            Using cmd As New SqlCommand(query, conn)
                numero = CInt(cmd.ExecuteScalar())
            End Using
        End Using
        Return numero
    End Function

    Public Function IsRevisioneModificabile(revisioneID As Integer) As Boolean
        Dim modificabile As Boolean = False
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query As String = "SELECT Stato FROM Mov_Revisioni WHERE RevisioneID = @ID"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.Add("@ID", SqlDbType.Int).Value = revisioneID
                Dim statoObj = cmd.ExecuteScalar()
                If statoObj IsNot Nothing Then
                    Dim stato = CStr(statoObj)
                    modificabile = (stato = "bozza" OrElse stato = "modifica")
                End If
            End Using
        End Using
        Return modificabile
    End Function

    Private Function EsistonoRevisioniAttive(videoID As Integer) As Boolean
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query As String = "SELECT COUNT(*) FROM Mov_Revisioni WHERE VideoID = @VideoID"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.Add("@VideoID", SqlDbType.Int).Value = videoID
                Dim count As Integer = CInt(cmd.ExecuteScalar())
                Return count > 0
            End Using
        End Using
    End Function

    Public Sub AggiornaRevisioneAttiva()
        If Parametri Is Nothing Then
            lblRevAttiva.Text = "Nessuna revisione attiva"
            Return
        End If
        lblRevAttiva.Text = Parametri.RevisioneID.ToString()
    End Sub

    Private Function VerificaCreazioneRevisione(videoID As Integer, revisioneCorrente As Integer) As Boolean
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim revisioniEsistenti As New List(Of Integer)
            Dim query As String = "SELECT RevisioneID FROM Mov_Revisioni WHERE VideoID = @VideoID ORDER BY RevisioneID ASC"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.Add("@VideoID", SqlDbType.Int).Value = videoID
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        revisioniEsistenti.Add(CInt(reader("RevisioneID")))
                    End While
                End Using
            End Using
            Dim revisioneSuccessiva = revisioneCorrente + 1
            If revisioniEsistenti.Contains(revisioneSuccessiva) Then
                MDIMessageBox.Show($"La Revisione {revisioneSuccessiva} esiste già. Per ricrearla, devi prima cancellarla.", Me.MdiParent, MessageBoxButtons.OK, "Revisione già presente")
                Return False
            End If
            Dim revisioniSuperiori = revisioniEsistenti.Where(Function(r) r > revisioneSuccessiva).ToList()
            If revisioniSuperiori.Any() Then
                Dim elenco = String.Join(", ", revisioniSuperiori)
                MDIMessageBox.Show($"Non è possibile ricreare la Revisione {revisioneSuccessiva} perché esistono revisioni successive ({elenco})." & vbCrLf & "Elimina prima tutte le revisioni superiori.", Me.MdiParent, MessageBoxButtons.OK, "Catena incoerente")
                Return False
            End If
            Return True
        End Using
    End Function

    Private Function TryParseRevisioneID(text As String, ByRef id As Integer) As Boolean
        If String.IsNullOrWhiteSpace(text) Then
            id = -1
            Return False
        End If
        Return Integer.TryParse(text, id)
    End Function

    ' --- Caricamento revisione esistente ---
    Public Sub CaricaRevisione(videoID As Integer, revisioneID As Integer)
        Dim videoPath As String = ""
        Dim nomeVideo = OttieniNomeVideo(videoID)
        Dim basePath = OttieniPercorsoFrames()
        Dim frameDir = Path.Combine(basePath, nomeVideo, $"Revisione_{revisioneID:0000}")

        Using conn As New SqlConnection(ConnString)
            Dim query As String = "SELECT FileScena FROM Mov_ConsegneScene WHERE VideoID = @VideoID"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.Add("@VideoID", SqlDbType.Int).Value = videoID
                conn.Open()
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing Then
                    videoPath = result.ToString()
                Else
                    MessageBox.Show("Video non trovato.")
                    Exit Sub
                End If
            End Using
        End Using

        If Not Directory.Exists(frameDir) OrElse Directory.GetFiles(frameDir).Length = 0 Then
            MessageBox.Show("Frame non trovati per la revisione selezionata.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        editor = New VideoEditor(videoPath, frameDir)
        editor.CurrentIndex = 0

        AggiornaNoteDaDatabase(revisioneID)

        Dim overlay = Me.Controls.Find("OverlayNotePanel", True).FirstOrDefault()
        overlay?.Invalidate()

        RemoveHandler TrackFrame.Scroll, AddressOf trackFrame_Scroll
        TrackFrame.Minimum = 0
        TrackFrame.Maximum = Math.Max(0, editor.FrameList.Count - 1)
        SafeSetTrackFrameValue(0)
        AddHandler TrackFrame.Scroll, AddressOf trackFrame_Scroll

        TrackFrame.Enabled = True
        btnSuccessivo.Enabled = True
        btnPrecedente.Enabled = True
        btnPrimoFrame.Enabled = True
        btnUltimoFrame.Enabled = True
        btnAvantiVeloce.Enabled = True
        btnIndietroVeloce.Enabled = True

        picFrame.Image = editor.LoadFrame(0)
        txtNote.Text = RecuperaNotaDaDatabase(revisioneID, 0)

        hasUnsavedChanges = False
        editor.HasUnsavedChanges = False
    End Sub

    Private Function OttieniNomeVideo(videoID As Integer) As String
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query = "SELECT Titolo FROM Mov_ConsegneScene WHERE VideoID = @ID"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.Add("@ID", SqlDbType.Int).Value = videoID
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing Then
                    Return result.ToString()
                Else
                    Throw New Exception($"Titolo non trovato per VideoID {videoID}")
                End If
            End Using
        End Using
    End Function

    ' --- Trackbar e pulsanti navigazione ---
    Private Sub trackFrame_Scroll(sender As Object, e As EventArgs) Handles TrackFrame.Scroll
        If editor Is Nothing Then Return
        If Not ConfirmSaveChanges() Then
            SafeSetTrackFrameValue(editor.CurrentIndex)
            Return
        End If
        Dim idx = TrackFrame.Value
        editor.CurrentIndex = idx
        picFrame.Image = editor.LoadFrame(idx)
        AggiornaFrameCorrente(idx)
        'UpdateFrameLabels()
    End Sub

    Private Sub btnSuccessivo_Click(sender As Object, e As EventArgs) Handles btnSuccessivo.Click
        If editor Is Nothing Then Return
        If editor.CurrentIndex < editor.FrameList.Count - 1 Then
            If Not ConfirmSaveChanges() Then Return
            Dim nuovoIndex = editor.CurrentIndex + 1
            editor.CurrentIndex = nuovoIndex
            SafeSetTrackFrameValue(nuovoIndex)
            picFrame.Image = editor.LoadFrame(nuovoIndex)
            AggiornaFrameCorrente(nuovoIndex)
        End If
    End Sub

    Private Sub btnPrecedente_Click(sender As Object, e As EventArgs) Handles btnPrecedente.Click
        If editor Is Nothing Then Return
        If editor.CurrentIndex > 0 Then
            If Not ConfirmSaveChanges() Then Return
            Dim nuovoIndex = editor.CurrentIndex - 1
            editor.CurrentIndex = nuovoIndex
            SafeSetTrackFrameValue(nuovoIndex)
            picFrame.Image = editor.LoadFrame(nuovoIndex)
            AggiornaFrameCorrente(nuovoIndex)
        End If
    End Sub

    Private Sub btnPrimoFrame_Click(sender As Object, e As EventArgs) Handles btnPrimoFrame.Click
        If editor Is Nothing Then Return
        If Not ConfirmSaveChanges() Then Return
        editor.CurrentIndex = 0
        SafeSetTrackFrameValue(0)
        picFrame.Image = editor.LoadFrame(0)
        Dim revID As Integer
        If TryParseRevisioneID(lblRevAttiva.Text, revID) Then
            txtNote.Text = RecuperaNotaDaDatabase(revID, 0)
        Else
            txtNote.Text = ""
        End If
    End Sub

    Private Sub btnUltimoFrame_Click(sender As Object, e As EventArgs) Handles btnUltimoFrame.Click
        If editor Is Nothing Then Return
        If editor.FrameList Is Nothing OrElse editor.FrameList.Count = 0 Then
            MessageBox.Show("Nessun frame disponibile.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Information)
            TrackFrame.Enabled = False
            btnSuccessivo.Enabled = False
            btnPrecedente.Enabled = False
            btnPrimoFrame.Enabled = False
            btnUltimoFrame.Enabled = False
            Return
        End If
        If Not ConfirmSaveChanges() Then Return
        Dim ultimoIndex = editor.FrameList.Count - 1
        editor.CurrentIndex = ultimoIndex
        TrackFrame.Minimum = 0
        TrackFrame.Maximum = Math.Max(0, editor.FrameList.Count - 1)
        SafeSetTrackFrameValue(ultimoIndex)
        picFrame.Image = editor.LoadFrame(ultimoIndex)
        Dim revID As Integer
        If TryParseRevisioneID(lblRevAttiva.Text, revID) Then
            txtNote.Text = RecuperaNotaDaDatabase(revID, ultimoIndex)
        Else
            txtNote.Text = ""
        End If
    End Sub

    Private Sub btnAvantiVeloce_Click(sender As Object, e As EventArgs) Handles btnAvantiVeloce.Click
        If editor Is Nothing Then
            MessageBox.Show("Impossibile caricare il frame.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If autoScrollActive AndAlso autoScrollDirection = "forward" Then
            autoScrollActive = False
            btnAvantiVeloce.Text = testoAvantiOriginale
        Else
            autoScrollActive = True
            autoScrollDirection = "forward"
            btnAvantiVeloce.Text = "Stop"
            btnIndietroVeloce.Text = testoIndietroOriginale
            StartAutoScroll()
        End If
    End Sub

    Private Sub btnIndietroVeloce_Click(sender As Object, e As EventArgs) Handles btnIndietroVeloce.Click
        If editor Is Nothing Then
            MessageBox.Show("Impossibile caricare il frame.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If autoScrollActive AndAlso autoScrollDirection = "backward" Then
            autoScrollActive = False
            btnIndietroVeloce.Text = testoIndietroOriginale
        Else
            autoScrollActive = True
            autoScrollDirection = "backward"
            btnIndietroVeloce.Text = "Stop"
            btnAvantiVeloce.Text = testoAvantiOriginale
            StartAutoScroll()
            AggiornaFrameCorrente(TrackFrame.Value)
        End If
    End Sub

    Private Async Sub StartAutoScroll()
        While autoScrollActive
            If hasUnsavedChanges Then
                If Not ConfirmSaveChanges() Then
                    autoScrollActive = False
                    btnAvantiVeloce.Text = testoAvantiOriginale
                    btnIndietroVeloce.Text = testoIndietroOriginale
                    Exit While
                End If
            End If

            Dim nuovoIndex As Integer = editor.CurrentIndex
            If autoScrollDirection = "forward" AndAlso editor.CurrentIndex < editor.FrameList.Count - 1 Then
                nuovoIndex += 1
                picFrame.Image = editor.LoadFrame(nuovoIndex)
            ElseIf autoScrollDirection = "backward" AndAlso editor.CurrentIndex > 0 Then
                nuovoIndex -= 1
                picFrame.Image = editor.LoadFrame(nuovoIndex)
            Else
                autoScrollActive = False
                If autoScrollDirection = "forward" Then
                    btnAvantiVeloce.Text = testoAvantiOriginale
                ElseIf autoScrollDirection = "backward" Then
                    btnIndietroVeloce.Text = testoIndietroOriginale
                End If
            End If

            editor.CurrentIndex = nuovoIndex
            SafeSetTrackFrameValue(nuovoIndex)

            Dim revID As Integer
            If TryParseRevisioneID(lblRevAttiva.Text, revID) Then
                txtNote.Text = RecuperaNotaDaDatabase(revID, nuovoIndex)
            Else
                txtNote.Text = ""
            End If

            Application.DoEvents()
            Await Task.Delay(60)
        End While
    End Sub

    ' --- Disegno note sul frame ---
    Private Sub picFrame_MouseDown(sender As Object, e As MouseEventArgs) Handles picFrame.MouseDown
        If picFrame.Image Is Nothing Then Return
        isDrawing = True
        lastPoint = e.Location
        editor.SaveState() ' salva stato per undo
    End Sub

    Private Sub picFrame_MouseMove(sender As Object, e As MouseEventArgs) Handles picFrame.MouseMove
        If picFrame.Image Is Nothing Then Return
        If isDrawing Then
            Using g As Graphics = Graphics.FromImage(editor.DrawingBitmap)
                Using penna As New Pen(colorePennino, spessorePennino)
                    g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                    g.DrawLine(penna, lastPoint, e.Location)
                End Using
            End Using
            picFrame.Image = CType(editor.DrawingBitmap.Clone(), Bitmap)
            lastPoint = e.Location
        End If
    End Sub

    Private Sub picFrame_MouseUp(sender As Object, e As MouseEventArgs) Handles picFrame.MouseUp
        If picFrame.Image Is Nothing Then Return
        isDrawing = False
        hasUnsavedChanges = True
        editor.HasUnsavedChanges = True
    End Sub

    Private Sub picFrame_MouseClick(sender As Object, e As MouseEventArgs) Handles picFrame.MouseClick
        notaPosizione = e.Location
    End Sub

    Private Sub btnColorePennino_Click(sender As Object, e As EventArgs) Handles btnColorePennino.Click
        If colorDialogPennino.ShowDialog = DialogResult.OK Then
            colorePennino = colorDialogPennino.Color
            btnColorePennino.BackColor = colorePennino
        End If
    End Sub

    Private Sub numSpessorePennino_ValueChanged(sender As Object, e As EventArgs) Handles numSpessorePennino.ValueChanged
        spessorePennino = CInt(numSpessorePennino.Value)
    End Sub

    Private Sub btnAnnulla_Click(sender As Object, e As EventArgs) Handles btnAnnulla.Click
        If picFrame.Image Is Nothing Then Return
        Dim ancoraModifiche = editor.Undo()
        picFrame.Image = CType(editor.DrawingBitmap.Clone(), Bitmap)
        hasUnsavedChanges = editor.HasUnsavedChanges
    End Sub

    Private Sub SalvaFrame()
        editor.SaveFrame()
        hasUnsavedChanges = False
        editor.HasUnsavedChanges = False

        RefreshLstNoteFrame()

    End Sub

    Private Sub btnSalvaVideo_Click(sender As Object, e As EventArgs) Handles btnSalvaVideo.Click
        If picFrame.Image Is Nothing Then
            MessageBox.Show("Caricare prima i Frames", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Dim outputPath = "C:\VideoEditor\output.mp4"
        editor.RebuildVideo(outputPath)
        MessageBox.Show("Video salvato in: " & outputPath)
    End Sub

    ' --- Lista annotazioni ---
    Private Function RecuperaNotaDaDatabase(revisioneID As Integer, frameIndex As Integer) As String
        Using conn As New SqlConnection(ConnString)
            Dim cmd As New SqlCommand("
                SELECT TestoNota 
                FROM Mov_FrameNote 
                WHERE RevisioneID = @RevID AND FrameIndex = @FrameIndex", conn)
            cmd.Parameters.Add("@RevID", SqlDbType.Int).Value = revisioneID
            cmd.Parameters.Add("@FrameIndex", SqlDbType.Int).Value = frameIndex
            conn.Open()
            Dim nota = cmd.ExecuteScalar()
            Return If(nota IsNot Nothing, nota.ToString(), "")
        End Using
    End Function

    Public Sub AggiornaFrameCorrente(index As Integer)
        If editor Is Nothing Then Exit Sub
        If index < 0 OrElse index >= editor.FrameList.Count Then Exit Sub
        editor.CurrentIndex = index
        picFrame.Image = editor.LoadFrame(index)

        Dim revisioneID As Integer
        If Not TryParseRevisioneID(lblRevAttiva.Text, revisioneID) Then
            txtNote.Text = ""
            lblAutore.Text = ""
            lblDataNota.Text = ""
            Return
        End If

        Using conn As New SqlConnection(ConnString)
            Dim cmd As New SqlCommand("
                SELECT TestoNota, NomeUtente, DataNota 
                FROM Mov_FrameNote 
                WHERE RevisioneID = @RevID AND FrameIndex = @FrameIndex", conn)
            cmd.Parameters.Add("@RevID", SqlDbType.Int).Value = revisioneID
            cmd.Parameters.Add("@FrameIndex", SqlDbType.Int).Value = index
            conn.Open()
            Using reader = cmd.ExecuteReader()
                If reader.Read() Then
                    txtNote.Text = reader("TestoNota").ToString()
                    lblAutore.Text = reader("NomeUtente").ToString()
                    lblDataNota.Text = $"{Convert.ToDateTime(reader("DataNota")):dd/MM/yyyy HH:mm}"
                Else
                    txtNote.Text = ""
                    lblAutore.Text = ""
                    lblDataNota.Text = ""
                End If
            End Using
        End Using

        hasUnsavedChanges = False
        editor.HasUnsavedChanges = False
    End Sub

    Public Sub AggiornaNoteDaDatabase(revisioneID As Integer)
        Try
            FrameConNote.Clear()

            ' Assicurati che lstNoteFrame sia un ListView
            If LstNoteFrame Is Nothing Then Return
            LstNoteFrame.BeginUpdate()
            LstNoteFrame.Items.Clear()

            Using conn As New SqlConnection(ConnString)
                Dim query As String = "
                SELECT FrameIndex, TestoNota, NomeUtente, DataNota
                FROM Mov_FrameNote
                WHERE RevisioneID = @RevID
                ORDER BY FrameIndex"
                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.Add("@RevID", SqlDbType.Int).Value = revisioneID
                    conn.Open()
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim frameIndex = Convert.ToInt32(reader("FrameIndex"))
                            Dim testo = reader("TestoNota").ToString()
                            Dim autore = reader("NomeUtente").ToString()
                            Dim data = Convert.ToDateTime(reader("DataNota"))

                            ' Registra che il frame ha note (per overlay)
                            If Not FrameConNote.Contains(frameIndex) Then FrameConNote.Add(frameIndex)

                            ' Prepara DTO
                            Dim info As New NotaFrameInfo With {
                            .FrameIndex = frameIndex,
                            .TestoNota = testo,
                            .Autore = autore,
                            .DataNota = data
                        }

                            ' Anteprima e testo visualizzato
                            Dim anteprima = If(testo.Length > 60, testo.Substring(0, 60) & " ...", testo)
                            Dim voce = $"Frame {frameIndex + 1} : {anteprima}"

                            ' Crea ListViewItem e assegna Tag
                            Dim item As New ListViewItem(voce) With {
                            .Tag = info,
                            .ToolTipText = $"Autore: {autore}{Environment.NewLine}Data: {data:dd/MM/yyyy HH:mm}"
                        }

                            LstNoteFrame.Items.Add(item)
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ' Ignora errori di popolamento per non interrompere il flusso
        Finally
            Try
                LstNoteFrame.EndUpdate()
            Catch
            End Try
            ' Aggiorna overlay/segnalini
            Dim overlay = Me.Controls.Find("OverlayNotePanel", True).FirstOrDefault()
            If overlay IsNot Nothing Then overlay.Invalidate()
        End Try
    End Sub

    Private Sub lstNoteFrame_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LstNoteFrame.SelectedIndexChanged
        Try
            If LstNoteFrame.SelectedItems.Count = 0 Then Return
            Dim item = LstNoteFrame.SelectedItems(0)
            Dim info = TryCast(item.Tag, NotaFrameInfo)
            If info Is Nothing Then Return

            If Not ConfirmSaveChanges() Then
                ' Ripristina selezione precedente: cerca l'item corrispondente all'indice corrente
                For Each it As ListViewItem In LstNoteFrame.Items
                    Dim inf = TryCast(it.Tag, NotaFrameInfo)
                    If inf IsNot Nothing AndAlso inf.FrameIndex = editor.CurrentIndex Then
                        it.Selected = True
                        Exit For
                    End If
                Next
                Return
            End If

            Dim frameIndex = info.FrameIndex
            If editor Is Nothing Then Return
            If frameIndex < 0 OrElse frameIndex >= editor.FrameList.Count Then Return

            editor.CurrentIndex = frameIndex
            SafeSetTrackFrameValue(frameIndex)
            picFrame.Image = editor.LoadFrame(frameIndex)

            txtNote.Text = info.TestoNota
            lblAutore.Text = info.Autore
            lblDataNota.Text = $"{info.DataNota:dd/MM/yyyy HH:mm}"

            hasUnsavedChanges = False
            editor.HasUnsavedChanges = False
        Catch ex As Exception
            ' ignora errori minori
        End Try
    End Sub

    ' Aggiorna la ListBox lstNoteFrame per il frame corrente in modo thread-safe
    Private Sub RefreshLstNoteFrame()
        If Me.IsDisposed Then Return

        Dim action = Sub()
                         Try
                             If LstNoteFrame Is Nothing Then Return
                             LstNoteFrame.BeginUpdate()
                             LstNoteFrame.Items.Clear()

                             If editor Is Nothing Then Return
                             Dim idx As Integer = editor.CurrentIndex

                             ' Se le note sono caricate in memoria in FrameNote (VideoEditor), usale
                             If editor.FrameNote IsNot Nothing AndAlso editor.FrameNote.ContainsKey(idx) Then
                                 Dim fn = editor.FrameNote(idx)
                                 Dim display = $"{fn.Data:yyyy-MM-dd HH:mm} - {fn.Autore}: {If(fn.Testo.Length > 80, fn.Testo.Substring(0, 80) & " ...", fn.Testo)}"
                                 Dim info As New NotaFrameInfo With {
                                 .FrameIndex = idx,
                                 .TestoNota = fn.Testo,
                                 .Autore = fn.Autore,
                                 .DataNota = fn.Data
                             }
                                 Dim item As New ListViewItem(display) With {
                                 .Tag = info,
                                 .ToolTipText = $"Autore: {info.Autore}{Environment.NewLine}Data: {info.DataNota:dd/MM/yyyy HH:mm}"
                             }
                                 LstNoteFrame.Items.Add(item)
                             End If
                         Catch
                             ' ignora
                         Finally
                             Try : LstNoteFrame.EndUpdate() : Catch : End Try
                         End Try
                     End Sub

        If Me.InvokeRequired Then
            Me.BeginInvoke(action)
        Else
            action()
        End If
    End Sub

    Private Sub EliminaNotaSelezionata()
        If LstNoteFrame.SelectedItems.Count = 0 Then Return
        Dim item = LstNoteFrame.SelectedItems(0)
        Dim info = TryCast(item.Tag, NotaFrameInfo)
        If info Is Nothing Then Return

        ' Parse revisione
        Dim revisioneID As Integer
        If Not TryParseRevisioneID(lblRevAttiva.Text, revisioneID) Then
            MessageBox.Show("Revisione non valida.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim frameIndex = info.FrameIndex
        Dim testoNota = info.TestoNota

        ' 1) Conferma eliminazione nota dal DB
        Dim confermaNota = MessageBox.Show("Vuoi davvero eliminare questa nota?", "Conferma eliminazione", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        If confermaNota <> DialogResult.Yes Then Return

        ' Elimina la nota dal DB
        Try
            Using conn As New SqlConnection(ConnString)
                conn.Open()
                Using cmd As New SqlCommand("
                DELETE FROM Mov_FrameNote
                WHERE RevisioneID = @RevID AND FrameIndex = @FrameIndex AND TestoNota = @TestoNota", conn)
                    cmd.Parameters.Add("@RevID", SqlDbType.Int).Value = revisioneID
                    cmd.Parameters.Add("@FrameIndex", SqlDbType.Int).Value = frameIndex
                    cmd.Parameters.Add("@TestoNota", SqlDbType.NVarChar, 2000).Value = testoNota
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Errore durante l'eliminazione della nota: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try

        ' Determina percorso overlay in modo robusto
        Dim dstBase As String = Nothing
        Try
            If editor IsNot Nothing AndAlso frameIndex >= 0 AndAlso frameIndex < editor.FrameList.Count Then
                dstBase = editor.FrameList(frameIndex)
            End If
        Catch ex As Exception
            ' log temporaneo per debug
            Try : File.AppendAllText(Path.Combine(Path.GetTempPath(), "VideoFBF_debug.log"), $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Errore determinazione dstBase: {ex.Message}{Environment.NewLine}") : Catch : End Try
            dstBase = Nothing
        End Try

        If String.IsNullOrWhiteSpace(dstBase) Then
            ' Non abbiamo il percorso del frame: aggiorniamo comunque UI e usciamo
            Try : AggiornaNoteDaDatabase(revisioneID) : Catch : End Try
            txtNote.Clear()
            lblAutore.Text = ""
            lblDataNota.Text = ""
            Return
        End If

        If Microsoft.VisualBasic.Strings.Right(dstBase, 12) = "_overlay.png" Then
            dstBase = Microsoft.VisualBasic.Strings.Left(dstBase, Len(dstBase) - 12)
        End If

        Dim overlayPath = Path.Combine(Path.GetDirectoryName(dstBase), Path.GetFileNameWithoutExtension(dstBase) & "_overlay.png")

        ' Se esiste overlay, chiedi conferma esplicita per eliminarlo
        If File.Exists(overlayPath) Then
            Dim confermaOverlay = MessageBox.Show("Vuoi eliminare anche gli appunti sull'immagine (file overlay)?", "Elimina overlay", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If confermaOverlay = DialogResult.Yes Then
                ' Prova a cancellare con retry
                Dim deleted = SafeDeleteFile(overlayPath, maxAttempts:=6)
                If Not deleted Then
                    MessageBox.Show("Non è stato possibile eliminare il file overlay. Riprova più tardi.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            End If
        End If

        ' Ricarica note dal DB per coerenza
        Try
            AggiornaNoteDaDatabase(revisioneID)
        Catch
            Try : RefreshLstNoteFrame() : Catch : End Try
        End Try

        ' Se non ci sono più note per quel frame, rimuovi l'indice da FrameConNote
        Try
            If FrameConNote.Contains(frameIndex) Then
                Dim hasStill As Boolean = False
                Try
                    Using conn As New SqlConnection(ConnString)
                        conn.Open()
                        Using cmd As New SqlCommand("SELECT COUNT(*) FROM Mov_FrameNote WHERE RevisioneID = @RevID AND FrameIndex = @FrameIndex", conn)
                            cmd.Parameters.Add("@RevID", SqlDbType.Int).Value = revisioneID
                            cmd.Parameters.Add("@FrameIndex", SqlDbType.Int).Value = frameIndex
                            Dim cnt = Convert.ToInt32(cmd.ExecuteScalar())
                            hasStill = (cnt > 0)
                        End Using
                    End Using
                Catch
                    hasStill = False
                End Try

                If Not hasStill Then
                    Try : FrameConNote.Remove(frameIndex) : Catch : End Try
                End If
            End If
        Catch
        End Try

        ' Forza reload del frame corrente per riflettere la cancellazione dell'overlay
        Try
            If editor IsNot Nothing Then
                If picFrame.Image IsNot Nothing Then
                    Try : picFrame.Image.Dispose() : Catch : End Try
                    picFrame.Image = Nothing
                End If
                picFrame.Image = editor.LoadFrame(frameIndex)
            Else
                picFrame.Image = LoadImageWithoutLock(dstBase)
            End If
        Catch
            ' fallback silenzioso
        End Try

        ' Aggiorna overlay/segnalini e pannello
        Try
            Dim overlayCtrl = Me.Controls.Find("OverlayNotePanel", True).FirstOrDefault()
            If overlayCtrl IsNot Nothing Then
                overlayCtrl.Invalidate()
                overlayCtrl.Refresh()
            End If
        Catch
        End Try

        ' Pulisce dettagli visualizzati
        txtNote.Clear()
        lblAutore.Text = ""
        lblDataNota.Text = ""
    End Sub

    ' ----------------------------
    ' SafeDeleteFile (cancellazione con retry, restituisce True se cancellato)
    ' ----------------------------
    Private Function SafeDeleteFile(path As String, Optional maxAttempts As Integer = 6) As Boolean
        If String.IsNullOrWhiteSpace(path) Then Return False
        If Not File.Exists(path) Then Return True

        Dim attempts As Integer = 0
        Dim lastEx As Exception = Nothing
        While attempts < maxAttempts
            attempts += 1
            Try
                File.Delete(path)
                Return True
            Catch ex As IOException
                lastEx = ex
                Threading.Thread.Sleep(100 * attempts)
            Catch ex As UnauthorizedAccessException
                lastEx = ex
                Threading.Thread.Sleep(100 * attempts)
            Catch ex As Exception
                lastEx = ex
                Exit While
            End Try
        End While

        Try

        Catch
        End Try

        Return False
    End Function

    ' --- Conferma salvataggio modifiche ---
    Private Function ConfirmSaveChanges() As Boolean
        If Not hasUnsavedChanges Then Return True
        Dim result = MessageBox.Show("Ci sono modifiche non salvate sul frame corrente. Vuoi salvare prima di procedere?", "Salvare modifiche?", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            Try
                Dim revisioneID As Integer
                If Not TryParseRevisioneID(lblRevAttiva.Text, revisioneID) Then
                    MessageBox.Show("Revisione non valida.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return False
                End If
                editor.SaveFrame()
                hasUnsavedChanges = False
                editor.HasUnsavedChanges = False

                Dim frameIndex = editor.CurrentIndex
                Dim nota = txtNote.Text.Trim()
                Using conn As New SqlConnection(ConnString)
                    conn.Open()
                    Using cmd As New SqlCommand("
                        MERGE INTO Mov_FrameNote AS target
                        USING (SELECT @RevID AS RevisioneID, @FrameIndex AS FrameIndex) AS source
                        ON target.RevisioneID = source.RevisioneID AND target.FrameIndex = source.FrameIndex
                        WHEN MATCHED THEN
                            UPDATE SET TestoNota = @TestoNota, NomeUtente = @NomeUtente, DataNota = @DataNota
                        WHEN NOT MATCHED THEN
                            INSERT (RevisioneID, FrameIndex, TestoNota, NomeUtente, DataNota)
                            VALUES (@RevID, @FrameIndex, @TestoNota, @NomeUtente, @DataNota);", conn)
                        cmd.Parameters.Add("@RevID", SqlDbType.Int).Value = revisioneID
                        cmd.Parameters.Add("@FrameIndex", SqlDbType.Int).Value = frameIndex
                        cmd.Parameters.Add("@TestoNota", SqlDbType.NVarChar, 2000).Value = nota
                        cmd.Parameters.Add("@NomeUtente", SqlDbType.NVarChar, 100).Value = NomeUtenteCorrente
                        cmd.Parameters.Add("@DataNota", SqlDbType.DateTime).Value = DateTime.Now
                        cmd.ExecuteNonQuery()
                    End Using
                End Using

                AggiornaNoteDaDatabase(revisioneID)
                Return True
            Catch ex As Exception
                MessageBox.Show("Errore durante il salvataggio: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try
        ElseIf result = DialogResult.No Then
            Try
                Dim idx = editor.CurrentIndex
                picFrame.Image = editor.LoadFrame(idx)
                hasUnsavedChanges = False
                editor.HasUnsavedChanges = False
            Catch ex As Exception
                MessageBox.Show("Errore durante il ripristino del frame: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try
            Return True
        Else
            Return False
        End If
    End Function

    ' --- Overlay note sulla trackbar ---
    Private Sub DisegnaSegnaliniNote(sender As Object, e As PaintEventArgs)
        If FrameConNote Is Nothing OrElse FrameConNote.Count = 0 Then Exit Sub
        Dim totale = TrackFrame.Maximum - TrackFrame.Minimum
        If totale <= 0 Then Exit Sub
        Dim larghezza = TrackFrame.Width - 28
        For Each index In FrameConNote
            If index < TrackFrame.Minimum OrElse index > TrackFrame.Maximum Then Continue For
            Dim percentuale = (index - TrackFrame.Minimum) / totale
            Dim x = CInt(percentuale * larghezza)
            e.Graphics.FillRectangle(Brushes.Red, x + 11, 0, 5, 10)
        Next
    End Sub

    ' --- Attivazione form: ricarica revisione se presente ---
    Private Sub VideoFBF_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        If Me.Tag IsNot Nothing Then
            Dim param = CType(Me.Tag, Object)
            Dim videoID = param.VideoID
            Dim revisioneID = param.RevisioneID
            Dim permesso = param.Permesso
            CaricaRevisione(videoID, revisioneID)
            AggiornaNoteDaDatabase(revisioneID)
        End If
    End Sub

    ' --- Gestione click per il pulsante Ripristina Frame ---
    Private Sub BtnRipristinaFrame_Click(sender As Object, e As EventArgs) Handles BtnRipristinaFrame.Click
        If editor Is Nothing Then
            MessageBox.Show("Nessun video caricato.", "Ripristino frame", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Verifica che esista la cartella di backup Revisione_0000
        Dim nomeVideo = Path.GetFileNameWithoutExtension(editor.VideoPath)
        Dim baseDir = Path.Combine(OttieniPercorsoFrames(), nomeVideo)
        Dim originalDir = Path.Combine(baseDir, "Revisione_0000")
        If Not Directory.Exists(originalDir) Then
            MessageBox.Show("Cartella di backup Revisione_0000 non trovata.", "Ripristino frame", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Controllo indice valido
        Dim idx = editor.CurrentIndex
        If idx < 0 OrElse idx >= editor.FrameList.Count Then
            MessageBox.Show("Indice frame non valido.", "Ripristino frame", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Se ci sono modifiche non salvate chiedi conferma
        If hasUnsavedChanges OrElse (editor IsNot Nothing AndAlso GetEditorHasUnsavedChanges(editor)) Then
            Dim res = MessageBox.Show("Ci sono modifiche non salvate. Vuoi comunque ripristinare il frame dall'originale? Le modifiche locali andranno perse.", "Conferma ripristino", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If res <> DialogResult.Yes Then
                Return
            End If
        Else
            Dim res = MessageBox.Show("Sei sicuro di voler ripristinare il frame corrente dall'originale?", "Conferma ripristino", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If res <> DialogResult.Yes Then
                Return
            End If
        End If

        ' Esegui il ripristino (metodo già presente nella classe)
        Try
            RestoreCurrentFrameFromOriginal()
            ' Aggiorna stato UI e flag
            hasUnsavedChanges = False
            If editor IsNot Nothing Then
                Try
                    editor.HasUnsavedChanges = False
                Catch
                    ' se la proprietà non esiste o fallisce, ignoriamo
                End Try
            End If
            AggiornaOverlayEStato() ' chiamata opzionale per aggiornare overlay/segnalini
        Catch ex As Exception
            MessageBox.Show("Errore durante il ripristino: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Helper per verificare se l'editor segnala modifiche (se la proprietà esiste)
    Private Function GetEditorHasUnsavedChanges(ed As VideoEditor) As Boolean
        Try
            Return ed.HasUnsavedChanges
        Catch
            Return False
        End Try
    End Function

    ' Metodo opzionale per aggiornare overlay o stato UI dopo il ripristino
    Private Sub AggiornaOverlayEStato()
        ' Aggiorna overlay se presente
        Dim overlay = Me.Controls.Find("OverlayNotePanel", True).FirstOrDefault()
        If overlay IsNot Nothing Then
            overlay.Invalidate()
        End If

        ' Aggiorna abilitazione del pulsante ripristina (se vuoi disabilitarlo quando non c'è backup)
        'UpdateRipristinaButton()
    End Sub

    ' Abilita/disabilita BtnRipristinaFrame in base alla presenza della cartella Revisione_0000 e dello stato editor
    Private Sub UpdateRipristinaButton()
        Try
            If editor Is Nothing OrElse String.IsNullOrEmpty(editor.VideoPath) Then
                BtnRipristinaFrame.Enabled = False
                Return
            End If
            Dim nomeVideo = Path.GetFileNameWithoutExtension(editor.VideoPath)
            Dim baseDir = Path.Combine(OttieniPercorsoFrames(), nomeVideo)
            Dim originalDir = Path.Combine(baseDir, "Revisione_0000")
            BtnRipristinaFrame.Enabled = Directory.Exists(originalDir)
        Catch
            BtnRipristinaFrame.Enabled = False
        End Try
    End Sub

    ' Nuovo handler integrato per aprire la form di scelta revisione
    Private Sub btnCaricaRevisione_Click(sender As Object, e As EventArgs) Handles btnCaricaRevisione.Click
        Try
            If Me.MdiParent Is Nothing Then
                MessageBox.Show("Form principale non disponibile.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            If TypeOf Me.MdiParent Is GesPu25 Then
                Dim mainForm As GesPu25 = CType(Me.MdiParent, GesPu25)
                Dim scelta As New SceltaVideo(Me)

                ' Se vuoi aprire come MDI child:
                scelta.MdiParent = mainForm
                scelta.Show()

            Else
                MessageBox.Show("Form principale non è del tipo atteso.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show("Errore durante l'apertura della finestra di scelta: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' --- DTO ListView Note ---
    Public Class NotaFrameInfo
        Public Property FrameIndex As Integer
        Public Property TestoNota As String
        Public Property Autore As String
        Public Property DataNota As DateTime
    End Class

    Private Sub btnSalvaNote_Click(sender As Object, e As EventArgs) Handles btnSalvaNote.Click
        If picFrame.Image Is Nothing Then Return

        Dim frameIndex = editor.CurrentIndex
        Dim nota = If(txtNote.Text, String.Empty).Trim()
        Dim nomeUtente = NomeUtenteCorrente
        Dim revisioneID As Integer
        If Not TryParseRevisioneID(lblRevAttiva.Text, revisioneID) Then
            MessageBox.Show("Revisione non valida.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' Percorso overlay per il frame corrente
        Dim dstBase As String = Nothing
        If editor IsNot Nothing AndAlso frameIndex >= 0 AndAlso frameIndex < editor.FrameList.Count Then
            dstBase = editor.FrameList(frameIndex)
        End If
        Dim overlayPath As String = Nothing
        If Not String.IsNullOrWhiteSpace(dstBase) Then
            overlayPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(dstBase), System.IO.Path.GetFileNameWithoutExtension(dstBase) & "_overlay.png")
        End If

        ' Controlli preliminari: drawing vuoto? overlay esistente? overlay identico?
        Dim drawingEmpty As Boolean = True
        Try
            drawingEmpty = IsBitmapFullyTransparent(editor.DrawingBitmap)
        Catch
            drawingEmpty = True
        End Try

        Dim overlayExists As Boolean = (Not String.IsNullOrWhiteSpace(overlayPath) AndAlso System.IO.File.Exists(overlayPath))
        Dim overlayIdentical As Boolean = False
        If overlayExists Then
            Try
                Dim fileChecksum = ComputeFileChecksum(overlayPath)
                Dim bmpChecksum = ComputeBitmapChecksum(editor.DrawingBitmap)
                If Not String.IsNullOrEmpty(fileChecksum) AndAlso Not String.IsNullOrEmpty(bmpChecksum) AndAlso fileChecksum = bmpChecksum Then
                    overlayIdentical = True
                End If
            Catch
                overlayIdentical = False
            End Try
        End If

        ' Caso 1: niente da salvare (nessun testo, drawing vuoto, nessun overlay)
        If String.IsNullOrWhiteSpace(nota) AndAlso drawingEmpty AndAlso Not overlayExists Then
            MessageBox.Show("Nessuna modifica da salvare.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        ' Caso 2: testo vuoto, drawing vuoto, overlay esistente identico -> niente da salvare
        If String.IsNullOrWhiteSpace(nota) AndAlso drawingEmpty AndAlso overlayExists AndAlso overlayIdentical Then
            MessageBox.Show("Nessuna modifica da salvare.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Try
            ' Se testo vuoto e drawing vuoto ma esisteva una nota nel DB -> cancellala
            If String.IsNullOrWhiteSpace(nota) AndAlso drawingEmpty Then
                Dim hadNote As Boolean = False
                Try
                    Using conn As New SqlConnection(ConnString)
                        conn.Open()
                        Using cmd As New SqlCommand("SELECT COUNT(*) FROM Mov_FrameNote WHERE RevisioneID = @RevID AND FrameIndex = @FrameIndex", conn)
                            cmd.Parameters.Add("@RevID", SqlDbType.Int).Value = revisioneID
                            cmd.Parameters.Add("@FrameIndex", SqlDbType.Int).Value = frameIndex
                            Dim cnt = Convert.ToInt32(cmd.ExecuteScalar())
                            hadNote = (cnt > 0)
                        End Using

                        If hadNote Then
                            Using del As New SqlCommand("DELETE FROM Mov_FrameNote WHERE RevisioneID = @RevID AND FrameIndex = @FrameIndex", conn)
                                del.Parameters.Add("@RevID", SqlDbType.Int).Value = revisioneID
                                del.Parameters.Add("@FrameIndex", SqlDbType.Int).Value = frameIndex
                                del.ExecuteNonQuery()
                            End Using
                        End If
                    End Using
                Catch ex As Exception
                    ' se la cancellazione fallisce, log e prosegui (non bloccare l'utente)
                    Try : System.IO.File.AppendAllText(System.IO.Path.Combine(System.IO.Path.GetTempPath(), "VideoFBF_save.log"), $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Errore DELETE nota: {ex.Message}{Environment.NewLine}") : Catch : End Try
                    Throw
                End Try

                ' Se non ci sono appunti da salvare, aggiorna UI e esci
                If drawingEmpty Then
                    ' Se esisteva un overlay e lo vogliamo rimuovere quando non ci sono appunti, cancellalo (opzionale)
                    If overlayExists Then
                        Try
                            ' prova a rimuovere l'overlay (senza forzare se bloccato)
                            SafeWriteBitmapAtomic(New System.Drawing.Bitmap(1, 1), overlayPath) ' sovrascrive con 1x1 trasparente come fallback
                            ' oppure usare SafeDeleteFileEnhanced se preferisci cancellare
                        Catch
                            ' ignora errori di cancellazione
                        End Try
                    End If

                    AggiornaNoteDaDatabase(revisioneID)
                    UpdateFrameLabels(frameIndex)
                    MessageBox.Show("Nota rimossa (nessun contenuto).", "Operazione completata", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If
            End If

            ' Se ci sono appunti non trasparenti, salva overlay
            If Not drawingEmpty Then
                editor.SaveFrame()
                hasUnsavedChanges = False
                editor.HasUnsavedChanges = False
            ElseIf overlayExists AndAlso Not overlayIdentical Then
                ' Se drawing vuoto ma overlay su disco differente, sovrascrivi con overlay trasparente per coerenza
                Try
                    Dim baseWidth As Integer = 0, baseHeight As Integer = 0
                    Using fs As New System.IO.FileStream(dstBase, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                        Using ms As New System.IO.MemoryStream()
                            fs.CopyTo(ms)
                            ms.Position = 0
                            Using tmpImg As System.Drawing.Image = System.Drawing.Image.FromStream(ms)
                                baseWidth = tmpImg.Width
                                baseHeight = tmpImg.Height
                            End Using
                        End Using
                    End Using
                    Using emptyBmp As New System.Drawing.Bitmap(Math.Max(1, baseWidth), Math.Max(1, baseHeight), System.Drawing.Imaging.PixelFormat.Format32bppArgb)
                        Using g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(emptyBmp)
                            g.Clear(System.Drawing.Color.Transparent)
                        End Using
                        SafeWriteBitmapAtomic(emptyBmp, overlayPath)
                    End Using
                Catch
                    ' fallback silenzioso
                End Try
            End If

            ' Salva/aggiorna la nota nel DB solo se il testo non è vuoto
            If Not String.IsNullOrWhiteSpace(nota) Then
                Using conn As New SqlConnection(ConnString)
                    conn.Open()
                    Using cmd As New SqlCommand("
                    MERGE INTO Mov_FrameNote AS target
                    USING (SELECT @RevID AS RevisioneID, @FrameIndex AS FrameIndex) AS source
                    ON target.RevisioneID = source.RevisioneID AND target.FrameIndex = source.FrameIndex
                    WHEN MATCHED THEN
                        UPDATE SET TestoNota = @TestoNota, NomeUtente = @NomeUtente, DataNota = @DataNota
                    WHEN NOT MATCHED THEN
                        INSERT (RevisioneID, FrameIndex, TestoNota, NomeUtente, DataNota)
                        VALUES (@RevID, @FrameIndex, @TestoNota, @NomeUtente, @DataNota);", conn)
                        cmd.Parameters.Add("@RevID", SqlDbType.Int).Value = revisioneID
                        cmd.Parameters.Add("@FrameIndex", SqlDbType.Int).Value = frameIndex
                        cmd.Parameters.Add("@TestoNota", SqlDbType.NVarChar, 2000).Value = nota
                        cmd.Parameters.Add("@NomeUtente", SqlDbType.NVarChar, 100).Value = nomeUtente
                        cmd.Parameters.Add("@DataNota", SqlDbType.DateTime).Value = DateTime.Now
                        cmd.ExecuteNonQuery()
                    End Using
                End Using
            End If

            ' Aggiorna UI e lista note
            AggiornaNoteDaDatabase(revisioneID)
            UpdateFrameLabels(frameIndex)
            MDIMessageBox.Show("Operazione completata", GesPu25, MessageBoxButtons.OK)
        Catch ex As Exception
            MDIMessageBox.Show("Errore durante il salvataggio: " & ex.Message, GesPu25, MessageBoxButtons.OK)
        End Try
    End Sub

    Private Function ComputeBitmapChecksum(bmp As System.Drawing.Bitmap) As String
        Using ms As New System.IO.MemoryStream()
            bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
            ms.Position = 0
            Using md5 As System.Security.Cryptography.MD5 = System.Security.Cryptography.MD5.Create()
                Dim hash = md5.ComputeHash(ms)
                Return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant()
            End Using
        End Using
    End Function

    Private Function ComputeFileChecksum(path As String) As String
        If Not System.IO.File.Exists(path) Then Return String.Empty
        Using fs As New System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
            Using md5 As System.Security.Cryptography.MD5 = System.Security.Cryptography.MD5.Create()
                Dim hash = md5.ComputeHash(fs)
                Return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant()
            End Using
        End Using
    End Function

    Private Function IsBitmapFullyTransparent(bmp As System.Drawing.Bitmap) As Boolean
        If bmp Is Nothing Then Return True
        Dim pf = bmp.PixelFormat
        If pf <> System.Drawing.Imaging.PixelFormat.Format32bppArgb AndAlso pf <> System.Drawing.Imaging.PixelFormat.Format32bppPArgb Then
            Using tmp As New System.Drawing.Bitmap(bmp.Width, bmp.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
                Using g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(tmp)
                    g.DrawImage(bmp, 0, 0)
                End Using
                Return IsBitmapFullyTransparent(tmp)
            End Using
        End If

        Dim rect = New System.Drawing.Rectangle(0, 0, bmp.Width, bmp.Height)
        Dim data = bmp.LockBits(rect, System.Drawing.Imaging.ImageLockMode.ReadOnly, bmp.PixelFormat)
        Try
            Dim bytesPerPixel = System.Drawing.Image.GetPixelFormatSize(bmp.PixelFormat) \ 8
            Dim stride = Math.Abs(data.Stride)
            Dim total = stride * bmp.Height
            Dim raw(total - 1) As Byte
            System.Runtime.InteropServices.Marshal.Copy(data.Scan0, raw, 0, total)
            For i As Integer = 0 To raw.Length - bytesPerPixel Step bytesPerPixel
                Dim alpha As Byte = raw(i + 3)
                If alpha <> 0 Then
                    Return False
                End If
            Next
            Return True
        Finally
            bmp.UnlockBits(data)
        End Try
    End Function

    Private Sub SafeWriteBitmapAtomic(bmp As System.Drawing.Bitmap, dstPath As String)
        Dim dir = System.IO.Path.GetDirectoryName(dstPath)
        If Not System.IO.Directory.Exists(dir) Then System.IO.Directory.CreateDirectory(dir)

        Dim tmp = System.IO.Path.Combine(dir, System.IO.Path.GetFileNameWithoutExtension(dstPath) & "_tmp" & System.IO.Path.GetExtension(dstPath))
        Try
            If System.IO.File.Exists(tmp) Then
                Try : System.IO.File.Delete(tmp) : Catch : End Try
            End If

            bmp.Save(tmp, System.Drawing.Imaging.ImageFormat.Png)

            If System.IO.File.Exists(dstPath) Then
                Try
                    System.IO.File.Replace(tmp, dstPath, Nothing)
                Catch ex As PlatformNotSupportedException
                    If System.IO.File.Exists(dstPath) Then System.IO.File.Delete(dstPath)
                    System.IO.File.Move(tmp, dstPath)
                End Try
            Else
                System.IO.File.Move(tmp, dstPath)
            End If
        Finally
            If System.IO.File.Exists(tmp) Then
                Try : System.IO.File.Delete(tmp) : Catch : End Try
            End If
        End Try
    End Sub


    Private Sub btnApprovazione_Click(sender As Object, e As EventArgs) Handles btnApprovazione.Click

    End Sub

    Private Sub btnRetake_Click(sender As Object, e As EventArgs) Handles btnRetake.Click

    End Sub

    Private Sub LstNoteFrame_KeyDown(sender As Object, e As KeyEventArgs) Handles LstNoteFrame.KeyDown
        If e.KeyCode = Keys.Delete Then
            EliminaNotaSelezionata()
        End If
    End Sub
End Class
