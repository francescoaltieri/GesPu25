Imports System.IO
Imports Microsoft.Data.SqlClient
Imports System.Threading.Tasks
Imports System.Diagnostics
Imports DocumentFormat.OpenXml.Vml.Office

Public Class VideoFBF
    ' --- Campi di stato e UI ---
    Dim editor As VideoEditor
    Dim isDrawing As Boolean = False
    Dim lastPoint As Point
    Dim autoScrollActive As Boolean = False
    Dim autoScrollDirection As String = "" ' "forward" o "backward"
    Dim testoAvantiOriginale As String = "Avanti Veloce ⟩⟩⟩"
    Dim testoIndietroOriginale As String = "⟨⟨⟨ Indietro Veloce"
    Dim notaPosizione As Point = Point.Empty
    Dim colorePennino As Color = Color.Red
    Dim spessorePennino As Integer = 5
    Dim disegnoAttivo As Boolean = False
    Private FrameConNote As New List(Of Integer)
    Private aggiornamentoInCorso As Boolean = False
    Private hasUnsavedChanges As Boolean = False

    Public Property Parametri As RevisioneParametri

    ' Strumenti disponibili
    Private Enum ToolType
        None = 0
        PointTool = 1
        LineTool = 2
        EllipseTool = 3
        RectangleTool = 4
    End Enum

    Private currentTool As ToolType = ToolType.None

    ' Per il disegno temporaneo (rubberband)
    Private isDragging As Boolean = False
    Private dragStart As Point = Point.Empty
    Private dragCurrent As Point = Point.Empty

    ' rubberband overlay
    Private overlayRubber As OverlayPanel = Nothing
    Private isRubberDragging As Boolean = False
    Private rubberStart As Point = Point.Empty
    Private rubberCurrent As Point = Point.Empty


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

        PenColor.BackColor = colorePennino

        ' COLLEGA l’evento Paint una sola volta
        Try
            RemoveHandler TrackFrame.Paint, AddressOf DisegnaSegnaliniNote
        Catch
        End Try
        AddHandler TrackFrame.Paint, AddressOf DisegnaSegnaliniNote

        InitLstNoteFrameColumns()
        InitStrumenti()
        InitOverlayRubber()
        ' Inizializza tooltip e handler per la ListView
        Try
            If LstNoteFrame IsNot Nothing Then
                ' ToolTip per mostrare la nota completa al passaggio del mouse
                Dim lvToolTip As New ToolTip()
                lvToolTip.AutoPopDelay = 10000
                lvToolTip.InitialDelay = 300
                lvToolTip.ReshowDelay = 100
                lvToolTip.ShowAlways = True
                LstNoteFrame.Tag = lvToolTip

                AddHandler LstNoteFrame.MouseMove, AddressOf LstNoteFrame_MouseMove_ShowTooltip
                AddHandler LstNoteFrame.MouseLeave, AddressOf LstNoteFrame_MouseLeave_HideTooltip
                AddHandler LstNoteFrame.DoubleClick, AddressOf LstNoteFrame_DoubleClick_OpenViewer
            End If
            SettaPicFrame("Zoom")
        Catch
        End Try

    End Sub

    Private Sub LstNoteFrame_MouseMove_ShowTooltip(sender As Object, e As MouseEventArgs)
        Try
            Dim lv = TryCast(sender, ListView)
            If lv Is Nothing Then Return
            Dim tt = TryCast(lv.Tag, ToolTip)
            If tt Is Nothing Then Return

            Dim info = lv.HitTest(e.Location)
            If info Is Nothing OrElse info.Item Is Nothing Then
                tt.Hide(lv)
                Return
            End If

            Dim item = info.Item
            Dim noteInfo = TryCast(item.Tag, NotaFrameInfo)
            Dim testoCompleto As String = String.Empty
            If noteInfo IsNot Nothing Then
                testoCompleto = noteInfo.TestoNota
                If String.IsNullOrWhiteSpace(testoCompleto) Then
                    testoCompleto = "(nota vuota)"
                End If
            Else
                ' fallback: usa subitem Nota se presente
                If item.SubItems.Count > 1 Then
                    testoCompleto = item.SubItems(1).Text
                End If
            End If

            ' Mostra tooltip vicino al mouse
            If Not String.IsNullOrEmpty(testoCompleto) Then
                tt.Show(testoCompleto, lv, e.Location.X + 15, e.Location.Y + 15, 10000)
            Else
                tt.Hide(lv)
            End If
        Catch
            ' ignore
        End Try
    End Sub

    ' Aggiorna il double-click per passare anche il numero revisione al viewer
    Private Sub LstNoteFrame_DoubleClick_OpenViewer(sender As Object, e As EventArgs)
        Try
            Dim lv = TryCast(sender, ListView)
            If lv Is Nothing OrElse lv.SelectedItems.Count = 0 Then Return
            Dim item = lv.SelectedItems(0)
            Dim info = TryCast(item.Tag, NotaFrameInfo)

            Dim frameIndex As Integer = -1
            Dim testo As String = String.Empty
            Dim autore As String = String.Empty
            Dim dataNota As DateTime = DateTime.MinValue

            If info IsNot Nothing Then
                frameIndex = info.FrameIndex
                testo = info.TestoNota
                autore = info.Autore
                dataNota = info.DataNota
            Else
                Integer.TryParse(item.SubItems(0).Text, frameIndex)
                frameIndex = Math.Max(1, frameIndex) - 1
                If item.SubItems.Count > 1 Then testo = item.SubItems(1).Text
                If item.SubItems.Count > 2 Then autore = item.SubItems(2).Text
            End If

            ' Ottieni numero revisione dalla label (fallback -1 se non parseable)
            Dim revisioneNum As Integer = -1
            Integer.TryParse(lblRevAttiva.Text, revisioneNum)

            ' Apri form modale con i dettagli e numero revisione
            Using viewer As New NoteViewerForm(revisioneNum, frameIndex, testo, autore, dataNota)
                viewer.StartPosition = FormStartPosition.CenterParent
                viewer.ShowDialog(Me)
            End Using
        Catch
            ' ignore
        End Try
    End Sub


    ' NoteViewerForm esteso per mostrare anche il numero revisione
    Public Class NoteViewerForm
        Inherits Form

        Private txtFullNote As TextBox
        Private lblInfo As Label
        Private lblRevisione As Label
        Private btnClose As Button

        ' Nuovo costruttore: revisioneNumber può essere -1 se non disponibile
        Public Sub New(revisioneNumber As Integer, frameIndex As Integer, testo As String, autore As String, dataNota As DateTime)
            Me.Text = If(frameIndex >= 0, $"Nota Frame {frameIndex + 1}", "Nota")
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ClientSize = New Size(760, 380)

            ' Label revisione in alto a sinistra
            lblRevisione = New Label() With {
            .AutoSize = False,
            .Location = New Point(10, 8),
            .Size = New Size(300, 20),
            .Font = New Font("Segoe UI", 9, FontStyle.Bold),
            .Text = If(revisioneNumber >= 0, $"Revisione: {revisioneNumber}", "Revisione: -")
        }
            Me.Controls.Add(lblRevisione)

            ' Info autore/data a destra
            lblInfo = New Label() With {
            .AutoSize = False,
            .Location = New Point(320, 8),
            .Size = New Size(420, 20),
            .TextAlign = ContentAlignment.MiddleRight,
            .Text = $"Autore: {If(String.IsNullOrEmpty(autore), "-", autore)}    Data: {If(dataNota = DateTime.MinValue, "-", dataNota.ToString("dd/MM/yyyy HH:mm"))}"
        }
            Me.Controls.Add(lblInfo)

            txtFullNote = New TextBox() With {
            .Multiline = True,
            .ReadOnly = True,
            .ScrollBars = ScrollBars.Vertical,
            .Location = New Point(10, 36),
            .Size = New Size(740, 290),
            .Font = New Font("Segoe UI", 10),
            .Text = If(String.IsNullOrEmpty(testo), "(nota vuota)", testo)
        }
            Me.Controls.Add(txtFullNote)

            btnClose = New Button() With {
            .Text = "Chiudi",
            .DialogResult = DialogResult.OK,
            .Size = New Size(100, 30),
            .Location = New Point(Me.ClientSize.Width - 110, Me.ClientSize.Height - 40)
        }
            Me.Controls.Add(btnClose)

            Me.AcceptButton = btnClose
        End Sub
    End Class



    Private Sub LstNoteFrame_MouseLeave_HideTooltip(sender As Object, e As EventArgs)
        Try
            Dim lv = TryCast(sender, ListView)
            If lv Is Nothing Then Return
            Dim tt = TryCast(lv.Tag, ToolTip)
            If tt IsNot Nothing Then tt.Hide(lv)
        Catch
        End Try
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
                MDIMessageBox.Show("Errore durante l'estrazione dei frame: " & exInner.Message, Me.MdiParent, MessageBoxButtons.OK)
                Exit Sub
            End Try

            ' 5) Verifica frame estratti
            Dim filesEstratti = GetFrameFiles(revisioneDir)
            If filesEstratti Is Nothing OrElse filesEstratti.Length = 0 Then
                MDIMessageBox.Show($"Estrazione completata ma nessun frame trovato in: {revisioneDir}", Me.MdiParent, MessageBoxButtons.OK)
                Exit Sub
            End If

            ' 6) CREA/POPOLA Revisione_0000 in background (non blocca UI)
            Try
                Await Task.Run(Sub()
                                   EnsureOriginalBackupFolder(nomeVideo, revisioneDir)
                               End Sub)
            Catch ex As Exception
                MDIMessageBox.Show("Attenzione: impossibile creare backup Revisione_0000: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
            End Try

            ' 7) Inserisci permesso utente con FK coerente
            Await Task.Run(Sub() InserisciPermessoUtente(newRevisioneID, SessioneUtente.NomeUtenteCorrente))

            MDIMessageBox.Show($"Video caricato, frame estratti, Revisione_{newRevisioneID:D4} registrata e backup {nomeVideo}_0000 creato.", Me.MdiParent, MessageBoxButtons.OK)

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
            MDIMessageBox.Show("Errore durante il caricamento: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
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
        Dim existing = GetFrameFiles(originalDir)
        If existing IsNot Nothing AndAlso existing.Length > 0 Then
            Return
        End If

        ' Prendi i frame estratti nella revisione corrente
        If Not Directory.Exists(revisioneDir) Then
            Throw New DirectoryNotFoundException($"Cartella revisione non trovata: {revisioneDir}")
        End If

        ' Prendi i frame estratti nella revisione corrente (escludi overlay)
        Dim framesEstratti = GetFrameFiles(revisioneDir)

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
                    Try
                        File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Warning: unable to delete overlay {overlayPath}: {If(lastEx IsNot Nothing, lastEx.Message, "unknown")}{Environment.NewLine}")
                    Catch
                    End Try
                    MDIMessageBox.Show("Impossibile eliminare il file overlay. Riprova più tardi.", Me.MdiParent, MessageBoxButtons.OK)
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
                    UpdateSegnaliniNote()
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

            MDIMessageBox.Show("Frame ripristinato: overlay eliminato e appunti rimossi in memoria.", Me.MdiParent, MessageBoxButtons.OK)

        Catch ex As Exception
            Try : File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - ERRORE: {ex.Message}{Environment.NewLine}{ex.StackTrace}{Environment.NewLine}")
            Catch
            End Try
            MDIMessageBox.Show("Errore durante il ripristino del frame: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
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
                FROM Tab_Utenti 
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
                               MDIMessageBox.Show("Revisione non valida", Me.MdiParent, MessageBoxButtons.OK)
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
                    MDIMessageBox.Show("Video non trovato", Me.MdiParent, MessageBoxButtons.OK)
                    Exit Sub
                End If
            End Using
        End Using

        If Not Directory.Exists(frameDir) OrElse GetFrameFiles(frameDir).Length = 0 Then
            MDIMessageBox.Show("Nessun Frame trovato per la revisione selezionata", Me.MdiParent, MessageBoxButtons.OK)
            Exit Sub
        End If

        editor = New VideoEditor(videoPath, frameDir)
        editor.CurrentIndex = 0

        AggiornaNoteDaDatabase(revisioneID)
        UpdateSegnaliniNote(revisioneID)

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
        UpdateFrameLabels()
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
            MDIMessageBox.Show("Nessun Frame disponibile", Me.MdiParent, MessageBoxButtons.OK)
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
            MDIMessageBox.Show("Impossibile caricare il frame", Me.MdiParent, MessageBoxButtons.OK)
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
            MDIMessageBox.Show("Impossibile caricare il frame", Me.MdiParent, MessageBoxButtons.OK)
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

    Private Sub picFrame_MouseClick(sender As Object, e As MouseEventArgs) Handles picFrame.MouseClick
        notaPosizione = e.Location
    End Sub

    Private Sub btnColorePennino_Click(sender As Object, e As EventArgs) Handles btnColorePennino.Click
        If colorDialogPennino.ShowDialog = DialogResult.OK Then
            colorePennino = colorDialogPennino.Color
            PenColor.BackColor = colorePennino
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
        editor.SaveFrame(editor.CurrentIndex, editor.DrawingBitmap)
        hasUnsavedChanges = False
        editor.HasUnsavedChanges = False
        RefreshLstNoteFrame()
    End Sub

    Private Sub btnSalvaVideo_Click(sender As Object, e As EventArgs) Handles btnSalvaVideo.Click
        If picFrame.Image Is Nothing Then
            MDIMessageBox.Show("Caricare prima i Frame", Me.MdiParent, MessageBoxButtons.OK)
            Return
        End If
        Dim outputPath = "C:\VideoEditor\output.mp4"
        editor.RebuildVideo(outputPath)
        MDIMessageBox.Show("Video salvato in: " & outputPath, Me.MdiParent, MessageBoxButtons.OK)
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
            If FrameConNote Is Nothing Then FrameConNote = New List(Of Integer)()
            FrameConNote.Clear()

            If LstNoteFrame Is Nothing Then Return

            ' Thread-safe invoke
            If Me.InvokeRequired Then
                Me.BeginInvoke(New Action(Of Integer)(AddressOf AggiornaNoteDaDatabase), revisioneID)
                Return
            End If

            ' Assicurati colonne
            If LstNoteFrame.Columns.Count < 3 Then InitLstNoteFrameColumns()

            LstNoteFrame.BeginUpdate()
            Try
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
                                Dim testo = If(reader("TestoNota") Is DBNull.Value, String.Empty, reader("TestoNota").ToString())
                                Dim autore = If(reader("NomeUtente") Is DBNull.Value, String.Empty, reader("NomeUtente").ToString())
                                Dim data = If(reader("DataNota") Is DBNull.Value, DateTime.MinValue, Convert.ToDateTime(reader("DataNota")))

                                If Not FrameConNote.Contains(frameIndex) Then FrameConNote.Add(frameIndex)

                                Dim displayNota = If(testo.Length > 120, testo.Substring(0, 120) & " ...", testo)
                                Dim item As New ListViewItem((frameIndex + 1).ToString()) ' colonna Frame (1-based)
                                item.SubItems.Add(displayNota)                            ' colonna Nota (anteprima)
                                item.SubItems.Add(autore)                                 ' colonna Utente
                                item.Tag = New NotaFrameInfo With {
                                .FrameIndex = frameIndex,
                                .TestoNota = testo,
                                .Autore = autore,
                                .DataNota = data
                            }
                                LstNoteFrame.Items.Add(item)
                            End While
                        End Using
                    End Using
                End Using
            Finally
                Try : LstNoteFrame.EndUpdate() : Catch : End Try
            End Try

            ' Ridisegna overlay/segnalini
            Try
                Dim overlay = Me.Controls.Find("OverlayNotePanel", True).FirstOrDefault()
                If overlay IsNot Nothing Then overlay.Invalidate()
            Catch
            End Try

        Catch ex As Exception
            Try
                Dim logPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "VideoFBF_debug_notes.log")
                System.IO.File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - AggiornaNoteDaDatabase ERROR: {ex.Message}{Environment.NewLine}")
            Catch : End Try
        End Try
    End Sub

    ' Chiamare una volta in Load per impostare la ListView in Details e creare le colonne
    Private Sub InitLstNoteFrameColumns()
        If LstNoteFrame Is Nothing Then Return
        LstNoteFrame.View = View.Details
        LstNoteFrame.FullRowSelect = True
        LstNoteFrame.GridLines = True
        LstNoteFrame.Columns.Clear()
        LstNoteFrame.Columns.Add("Frame", 70, HorizontalAlignment.Left)
        LstNoteFrame.Columns.Add("Nota", 420, HorizontalAlignment.Left)
        LstNoteFrame.Columns.Add("Utente", 150, HorizontalAlignment.Left)
    End Sub

    ' Aggiorna la ListBox lstNoteFrame per il frame corrente in modo thread-safe
    Private Sub RefreshLstNoteFrame()
        If Me.IsDisposed Then Return

        Dim action = Sub()
                         Try
                             If LstNoteFrame Is Nothing Then Return
                             If LstNoteFrame.Columns.Count < 3 Then InitLstNoteFrameColumns()

                             LstNoteFrame.BeginUpdate()
                             Try
                                 LstNoteFrame.Items.Clear()
                                 If editor Is Nothing Then Return

                                 Dim idx As Integer = editor.CurrentIndex

                                 ' Se ci sono note in memoria per il frame corrente
                                 If editor.FrameNote IsNot Nothing AndAlso editor.FrameNote.ContainsKey(idx) Then
                                     Dim fn = editor.FrameNote(idx)
                                     Dim displayNota = If(fn.Testo.Length > 120, fn.Testo.Substring(0, 120) & " ...", fn.Testo)
                                     Dim item As New ListViewItem((idx + 1).ToString())
                                     item.SubItems.Add(displayNota)
                                     item.SubItems.Add(fn.Autore)
                                     item.Tag = New NotaFrameInfo With {
                                     .FrameIndex = idx,
                                     .TestoNota = fn.Testo,
                                     .Autore = fn.Autore,
                                     .DataNota = fn.Data
                                 }
                                     LstNoteFrame.Items.Add(item)
                                 Else
                                     ' Nessuna nota in memoria per il frame corrente: lascia vuoto
                                 End If
                             Finally
                                 Try : LstNoteFrame.EndUpdate() : Catch : End Try
                             End Try
                         Catch
                             Try : LstNoteFrame.EndUpdate() : Catch : End Try
                         End Try
                     End Sub

        If Me.InvokeRequired Then
            Me.BeginInvoke(action)
        Else
            action()
        End If
    End Sub

    Private Sub LstNoteFrame_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LstNoteFrame.SelectedIndexChanged
        Try
            If LstNoteFrame.SelectedItems.Count = 0 Then Return
            Dim item = LstNoteFrame.SelectedItems(0)
            Dim info = TryCast(item.Tag, NotaFrameInfo)
            If info Is Nothing Then
                ' Se Tag non è presente, ricava dai SubItems (fallback)
                Dim frameStr = item.SubItems(0).Text
                Dim frameIndex As Integer = 0
                Integer.TryParse(frameStr, frameIndex)
                frameIndex = Math.Max(1, frameIndex) - 1
                Dim testo = If(item.SubItems.Count > 1, item.SubItems(1).Text, String.Empty)
                Dim autore = If(item.SubItems.Count > 2, item.SubItems(2).Text, String.Empty)
                txtNote.Text = testo
                lblAutore.Text = autore
                lblDataNota.Text = ""
                ' Aggiorna immagine/frame
                If editor IsNot Nothing AndAlso frameIndex >= 0 AndAlso frameIndex < editor.FrameList.Count Then
                    editor.CurrentIndex = frameIndex
                    SafeSetTrackFrameValue(frameIndex)
                    picFrame.Image = editor.LoadFrame(frameIndex)
                End If
                Return
            End If

            ' Se ci sono modifiche non salvate, chiedi conferma (mantieni la logica esistente)
            If Not ConfirmSaveChanges() Then
                ' ripristina selezione precedente: cerca l'item corrispondente all'indice corrente
                For Each it As ListViewItem In LstNoteFrame.Items
                    Dim inf = TryCast(it.Tag, NotaFrameInfo)
                    If inf IsNot Nothing AndAlso inf.FrameIndex = editor.CurrentIndex Then
                        it.Selected = True
                        Exit For
                    End If
                Next
                Return
            End If

            ' Applica selezione: carica frame e dettagli
            Dim frameIndexSel = info.FrameIndex
            If editor Is Nothing Then Return
            If frameIndexSel < 0 OrElse frameIndexSel >= editor.FrameList.Count Then Return

            editor.CurrentIndex = frameIndexSel
            SafeSetTrackFrameValue(frameIndexSel)
            picFrame.Image = editor.LoadFrame(frameIndexSel)

            txtNote.Text = info.TestoNota
            lblAutore.Text = info.Autore
            lblDataNota.Text = If(info.DataNota = DateTime.MinValue, "", $"{info.DataNota:dd/MM/yyyy HH:mm}")

            hasUnsavedChanges = False
            Try : editor.HasUnsavedChanges = False : Catch : End Try
        Catch ex As Exception
            ' ignora errori minori
        End Try
    End Sub

    Private Sub EliminaNotaSelezionata()
        If LstNoteFrame.SelectedItems.Count = 0 Then Return
        Dim item = LstNoteFrame.SelectedItems(0)
        Dim info = TryCast(item.Tag, NotaFrameInfo)
        If info Is Nothing Then Return

        ' Parse revisione
        Dim revisioneID As Integer
        If Not TryParseRevisioneID(lblRevAttiva.Text, revisioneID) Then
            MDIMessageBox.Show("Revisione non valida", Me.MdiParent, MessageBoxButtons.OK)
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
            MDIMessageBox.Show("Errore durante l'eliminazione della nota: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
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
                    MDIMessageBox.Show("Non è stato possibile eliminare il file overlay. Riprova più tardi.", Me.MdiParent, MessageBoxButtons.OK)
                End If
            End If
        End If

        ' Ricarica note dal DB per coerenza
        Try
            AggiornaNoteDaDatabase(revisioneID)
            UpdateSegnaliniNote(revisioneID)
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
                    MDIMessageBox.Show("Revisione non valida", Me.MdiParent, MessageBoxButtons.OK)
                    Return False
                End If
                editor.SaveFrame(editor.CurrentIndex, editor.DrawingBitmap)
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
                MDIMessageBox.Show("Errore durante il salvataggio: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
                Return False
            End Try
            Return True
        Else
            Return False
        End If
    End Function

    ' --- Attivazione form: ricarica revisione se presente ---
    Private Sub VideoFBF_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        If Me.Tag IsNot Nothing Then
            Dim param = CType(Me.Tag, Object)
            Dim videoID = param.VideoID
            Dim revisioneID = param.RevisioneID
            Dim permesso = param.Permesso
            CaricaRevisione(videoID, revisioneID)
            AggiornaNoteDaDatabase(revisioneID)
            UpdateSegnaliniNote(revisioneID)
        End If
    End Sub

    ' --- Gestione click per il pulsante Ripristina Frame ---
    Private Sub BtnRipristinaFrame_Click(sender As Object, e As EventArgs) Handles BtnRipristinaFrame.Click
        If editor Is Nothing Then
            MDIMessageBox.Show("Nessun video caricato", Me.MdiParent, MessageBoxButtons.OK)
            Return
        End If

        ' Verifica che esista la cartella di backup Revisione_0000
        Dim nomeVideo = Path.GetFileNameWithoutExtension(editor.VideoPath)
        Dim baseDir = Path.Combine(OttieniPercorsoFrames(), nomeVideo)
        Dim originalDir = Path.Combine(baseDir, "Revisione_0000")
        If Not Directory.Exists(originalDir) Then
            MDIMessageBox.Show("Cartella di backup Revisione_0000 non trovato", Me.MdiParent, MessageBoxButtons.OK)
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
            MDIMessageBox.Show("Errore durante il ripristino: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
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
        Cursor.Current = Cursors.WaitCursor
        Application.DoEvents()
        Try
            If Me.MdiParent Is Nothing Then
                MDIMessageBox.Show("Form Principale non disponibile", Me.MdiParent, MessageBoxButtons.OK)
                Return
            End If

            If TypeOf Me.MdiParent Is GesPu25 Then
                Dim mainForm As GesPu25 = CType(Me.MdiParent, GesPu25)
                Dim scelta As New SceltaVideo(Me)
                scelta.MdiParent = mainForm
                scelta.Show()
            Else
                MDIMessageBox.Show("Form Principale non è del tipo atteso", Me.MdiParent, MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            MDIMessageBox.Show("Errore durante l'apertura della finestra di scelta: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
        End Try
        Cursor.Current = Cursors.Default
        Application.DoEvents()
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

            ' Per Format32bpp* l'ordine in memoria è B G R A (alpha all'offset +3)
            For y As Integer = 0 To bmp.Height - 1
                Dim rowStart = y * stride
                For x As Integer = 0 To bmp.Width - 1
                    Dim i = rowStart + x * bytesPerPixel
                    Dim alpha As Byte = raw(i + 3)
                    If alpha <> 0 Then
                        Return False
                    End If
                Next
            Next
            Return True
        Finally
            bmp.UnlockBits(data)
        End Try
    End Function

    Private Function BitmapsAreIdentical(bmpA As System.Drawing.Bitmap, bmpB As System.Drawing.Bitmap) As Boolean
        If bmpA Is Nothing AndAlso bmpB Is Nothing Then Return True
        If bmpA Is Nothing OrElse bmpB Is Nothing Then Return False
        If bmpA.Width <> bmpB.Width OrElse bmpA.Height <> bmpB.Height Then Return False

        Dim rect = New System.Drawing.Rectangle(0, 0, bmpA.Width, bmpA.Height)
        Dim dataA = bmpA.LockBits(rect, System.Drawing.Imaging.ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
        Dim dataB = bmpB.LockBits(rect, System.Drawing.Imaging.ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
        Try
            Dim strideA = Math.Abs(dataA.Stride)
            Dim strideB = Math.Abs(dataB.Stride)
            If strideA <> strideB Then Return False

            Dim total = strideA * bmpA.Height
            Dim rawA(total - 1) As Byte
            Dim rawB(total - 1) As Byte
            System.Runtime.InteropServices.Marshal.Copy(dataA.Scan0, rawA, 0, total)
            System.Runtime.InteropServices.Marshal.Copy(dataB.Scan0, rawB, 0, total)

            For i As Integer = 0 To total - 1
                If rawA(i) <> rawB(i) Then
                    Return False
                End If
            Next
            Return True
        Finally
            bmpA.UnlockBits(dataA)
            bmpB.UnlockBits(dataB)
        End Try
    End Function

    Private Function IsOverlayEmpty(ed As VideoEditor, frameIndex As Integer) As Boolean
        Try
            If ed Is Nothing Then Return True
            If frameIndex < 0 OrElse frameIndex >= ed.FrameList.Count Then Return True

            Dim basePath = ed.FrameList(frameIndex)
            If String.IsNullOrWhiteSpace(basePath) Then Return True

            ' Carica il base in memoria senza lock
            Dim baseBmp As System.Drawing.Bitmap = Nothing
            Using fs As New System.IO.FileStream(basePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                Using ms As New System.IO.MemoryStream()
                    fs.CopyTo(ms)
                    ms.Position = 0
                    Using tmpImg As System.Drawing.Image = System.Drawing.Image.FromStream(ms)
                        baseBmp = New System.Drawing.Bitmap(tmpImg)
                    End Using
                End Using
            End Using

            ' Se drawingBitmap è nulla o 1x1 trasparente consideralo vuoto
            If ed.DrawingBitmap Is Nothing Then
                baseBmp.Dispose()
                Return True
            End If

            ' Se le dimensioni differiscono, normalizza: crea copia del drawing con le dimensioni del base
            Dim drawingCopy As System.Drawing.Bitmap = Nothing
            If ed.DrawingBitmap.Width = baseBmp.Width AndAlso ed.DrawingBitmap.Height = baseBmp.Height Then
                drawingCopy = CType(ed.DrawingBitmap.Clone(), System.Drawing.Bitmap)
            Else
                drawingCopy = New System.Drawing.Bitmap(baseBmp.Width, baseBmp.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
                Using g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(drawingCopy)
                    g.Clear(System.Drawing.Color.Transparent)
                    g.DrawImage(ed.DrawingBitmap, 0, 0, Math.Min(ed.DrawingBitmap.Width, baseBmp.Width), Math.Min(ed.DrawingBitmap.Height, baseBmp.Height))
                End Using
            End If

            Dim identical = BitmapsAreIdentical(baseBmp, drawingCopy)

            drawingCopy.Dispose()
            baseBmp.Dispose()

            Return identical ' True => overlay vuoto (nessuna differenza)
        Catch
            ' In caso di errore, per sicurezza considera overlay non vuoto (evita cancellare)
            Return False
        End Try
    End Function


    Private Sub btnApprovazione_Click(sender As Object, e As EventArgs) Handles btnApprovazione.Click
        If picFrame.Image Is Nothing Then
            MessageBox.Show("Caricare prima i Frames", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ' Ignora il click se non ci sono frame
            Return
        End If
        Dim revisioneID As Integer = Int(Me.lblRevAttiva.Text)

        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query As String = "
            UPDATE Mov_Revisione
            SET Stato = 'Conforme', Approvato = 1
            WHERE RevisioneID = @RevisioneID"

            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@RevisioneID", revisioneID)
                cmd.ExecuteNonQuery()
            End Using
        End Using

    End Sub

    Private Sub btnRetake_Click(sender As Object, e As EventArgs) Handles btnRetake.Click
        If picFrame.Image Is Nothing Then
            MessageBox.Show("Caricare prima i Frames", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ' Ignora il click se non ci sono frame
            Return
        End If
        Dim revisioneID As Integer = Int(Me.lblRevAttiva.Text)

        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query As String = "
            UPDATE Mov_Revisione
            SET Stato = 'Non Conforme', Approvato = 0
            WHERE RevisioneID = @RevisioneID"

            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@RevisioneID", revisioneID)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Private Sub LstNoteFrame_KeyDown(sender As Object, e As KeyEventArgs) Handles LstNoteFrame.KeyDown
        If e.KeyCode = Keys.Delete Then
            EliminaNotaSelezionata()
        End If
    End Sub

    Private Sub btnSalvaNote_Click(sender As Object, e As EventArgs) Handles btnSalvaNote.Click
        If picFrame.Image Is Nothing OrElse editor Is Nothing Then
            MDIMessageBox.Show("Caricare prima i Frame", Me.MdiParent, MessageBoxButtons.OK)
            Return
        End If

        Dim frameIndex = editor.CurrentIndex
        Dim nota = If(txtNote.Text, String.Empty).Trim()
        Dim nomeUtente = NomeUtenteCorrente
        Dim revisioneID As Integer
        If Not TryParseRevisioneID(lblRevAttiva.Text, revisioneID) Then
            MDIMessageBox.Show("Revisione non valida", Me.MdiParent, MessageBoxButtons.OK)
            Return
        End If

        ' Percorso base e overlay
        Dim dstBase As String = Nothing
        If editor IsNot Nothing AndAlso frameIndex >= 0 AndAlso frameIndex < editor.FrameList.Count Then
            dstBase = editor.FrameList(frameIndex)
        End If

        Dim overlayPath As String = Nothing
        If Not String.IsNullOrWhiteSpace(dstBase) Then
            overlayPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(dstBase), System.IO.Path.GetFileNameWithoutExtension(dstBase) & "_overlay.png")
        End If

        ' Determina se il drawing è vuoto (se non è possibile valutare, consideralo vuoto per sicurezza)
        Dim drawingEmpty As Boolean = True
        Try
            drawingEmpty = IsOverlayEmpty(editor, frameIndex)
        Catch
            ' se non possiamo valutare, assumiamo vuoto per evitare salvataggi indesiderati
            drawingEmpty = True
        End Try

        ' NUOVA REGOLA: se non c'è testo e non c'è disegno -> NON SALVARE
        If String.IsNullOrWhiteSpace(nota) AndAlso drawingEmpty Then
            MDIMessageBox.Show("Nessuna modifica da salvare: inserire testo o disegnare sull'immagine prima di salvare.", Me.MdiParent, MessageBoxButtons.OK)
            Return
        End If

        ' Altrimenti procedi con il salvataggio coerente:
        Try
            ' 1) Se c'è disegno, salva overlay (anche se testo vuoto)
            If Not drawingEmpty Then
                Try
                    ' usa overload compatibile (wrapper SaveFrame() o SaveFrame(frameIndex, bitmap))
                    If editor IsNot Nothing Then
                        ' preferiamo chiamare l'overload esplicito se disponibile
                        Try
                            editor.SaveFrame(frameIndex, editor.DrawingBitmap)
                        Catch
                            ' fallback alla wrapper senza parametri
                            editor.SaveFrame(frameIndex, editor.DrawingBitmap)
                        End Try
                    End If
                Catch ex As Exception
                    Throw New Exception("Impossibile salvare l'overlay: " & ex.Message)
                End Try

                ' azzera flag solo dopo salvataggio overlay
                hasUnsavedChanges = False
                Try : editor.HasUnsavedChanges = False : Catch : End Try
            End If

            ' 2) Salva/aggiorna la nota nel DB (se testo non vuoto oppure se c'è disegno vogliamo comunque registrare la nota)
            If Not String.IsNullOrWhiteSpace(nota) OrElse Not drawingEmpty Then
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
                        cmd.Parameters.Add("@TestoNota", SqlDbType.NVarChar, 2000).Value = If(String.IsNullOrEmpty(nota), String.Empty, nota)
                        cmd.Parameters.Add("@NomeUtente", SqlDbType.NVarChar, 100).Value = nomeUtente
                        cmd.Parameters.Add("@DataNota", SqlDbType.DateTime).Value = DateTime.Now
                        cmd.ExecuteNonQuery()
                    End Using
                End Using
            End If

            ' 3) Aggiorna UI e segnalini
            If Me.InvokeRequired Then
                Me.BeginInvoke(New Action(Sub()
                                              AggiornaNoteDaDatabase(revisioneID)
                                              RefreshLstNoteFrame()
                                              UpdateSegnaliniNote(revisioneID)
                                          End Sub))
            Else
                AggiornaNoteDaDatabase(revisioneID)
                RefreshLstNoteFrame()
                UpdateSegnaliniNote(revisioneID)
            End If

            ' 4) Assicura che i flag siano azzerati dopo salvataggio
            hasUnsavedChanges = False
            Try : editor.HasUnsavedChanges = False : Catch : End Try

            MDIMessageBox.Show("Operazione completata: nota e/o overlay salvati.", Me.MdiParent, MessageBoxButtons.OK)
        Catch ex As Exception
            MDIMessageBox.Show("Errore durante il salvataggio: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
        End Try
    End Sub


    ' Disegna i segnalini rossi sopra la TrackBar
    Private Sub DisegnaSegnaliniNote(sender As Object, e As PaintEventArgs)
        Try
            If FrameConNote Is Nothing OrElse FrameConNote.Count = 0 Then Return
            If TrackFrame Is Nothing Then Return

            Dim minVal = TrackFrame.Minimum
            Dim maxVal = TrackFrame.Maximum
            ' se non ci sono range validi, esci
            If maxVal <= minVal Then Return

            ' larghezza utile: togli spazio per thumb e bordi (aggiusta se necessario)
            Dim trackClientWidth = Math.Max(1, TrackFrame.Width - 28)
            Dim totale = maxVal - minVal

            Dim g = e.Graphics
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias

            For Each index In FrameConNote
                If index < minVal OrElse index > maxVal Then Continue For
                Dim percentuale As Double = (index - minVal) / CDbl(totale)
                Dim x = CInt(percentuale * trackClientWidth)
                Dim rectX = x + 11 ' offset per allineare con il centro della thumb
                Dim rect = New System.Drawing.Rectangle(rectX, 0, 6, 10)
                Using b As New System.Drawing.SolidBrush(System.Drawing.Color.Red)
                    g.FillRectangle(b, rect)
                End Using
            Next
        Catch ex As Exception
            ' non bloccare UI: loggare se serve
            Try
                Dim logPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "VideoFBF_draw.log")
                System.IO.File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - DisegnaSegnaliniNote error: {ex.Message}{Environment.NewLine}")
            Catch : End Try
        End Try
    End Sub

    ' Aggiorna i segnalini: ricarica note e forza ridisegno overlay e trackbar
    Private Sub UpdateSegnaliniNote(Optional revisioneID As Integer = -1)
        Try
            Dim revID As Integer = revisioneID
            If revID < 0 Then
                If Not TryParseRevisioneID(lblRevAttiva.Text, revID) Then revID = -1
            End If

            If revID >= 0 Then
                AggiornaNoteDaDatabase(revID)
            Else
                RefreshLstNoteFrame()
            End If
        Catch
            ' ignora errori di popolamento
        End Try

        ' Forza creazione/posizionamento del pannello overlay e ridisegno
        Try
            Dim overlay = Me.Controls.Find("OverlayNotePanel", True).FirstOrDefault()
            If overlay Is Nothing Then
                Dim nuovoOverlay As New Panel With {
                .Width = TrackFrame.Width,
                .Height = 10,
                .Location = New Point(TrackFrame.Left, TrackFrame.Top - 10),
                .BackColor = System.Drawing.Color.Transparent,
                .Name = "OverlayNotePanel"
            }
                Me.Controls.Add(nuovoOverlay)
                AddHandler nuovoOverlay.Paint, AddressOf DisegnaSegnaliniNote
                nuovoOverlay.BringToFront()
            Else
                ' aggiorna dimensione/posizione in caso di resize o spostamento
                overlay.Width = TrackFrame.Width
                overlay.Location = New Point(TrackFrame.Left, TrackFrame.Top - overlay.Height)
                overlay.Invalidate()
                overlay.BringToFront()
            End If
        Catch
        End Try

        Try
            TrackFrame.Invalidate()
            TrackFrame.Update()
        Catch
        End Try
    End Sub

    ' Restituisce solo i file "frame" escludendo overlay, tmp e file nascosti
    Public Shared Function GetFrameFiles(dir As String) As String()
        If String.IsNullOrWhiteSpace(dir) OrElse Not Directory.Exists(dir) Then
            Return New String() {}
        End If

        Try
            Dim all = Directory.GetFiles(dir, "*.png")
            Dim list = New List(Of String)

            For Each f In all
                Dim name = Path.GetFileName(f)

                ' Escludi overlay anywhere nel nome (case-insensitive)
                If name.IndexOf("_overlay", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Continue For
                End If

                ' Escludi file temporanei convenzionali
                If name.StartsWith("_tmp", StringComparison.OrdinalIgnoreCase) OrElse name.StartsWith(".", StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If

                ' Escludi file con attributo Hidden
                Try
                    Dim attr = File.GetAttributes(f)
                    If (attr And FileAttributes.Hidden) = FileAttributes.Hidden Then
                        Continue For
                    End If
                Catch
                    ' ignore attribute errors
                End Try

                list.Add(f)
            Next

            ' Ordina in modo stabile: prova a estrarre l'ultimo numero nel nome, fallback alfabetico
            Return list.OrderBy(Function(p)
                                    Dim noExt = Path.GetFileNameWithoutExtension(p)
                                    Dim m = System.Text.RegularExpressions.Regex.Match(noExt, "(\d+)(?!.*\d)")
                                    If m.Success Then
                                        Dim n As Integer
                                        If Integer.TryParse(m.Value, n) Then Return n
                                    End If
                                    Return Integer.MaxValue
                                End Function).
                    ThenBy(Function(p) Path.GetFileName(p).ToLowerInvariant()).
                    ToArray()
        Catch
            Return New String() {}
        End Try
    End Function



    ' ---------------------------
    ' SaveOverlayAnchored helper (fallback)
    ' ---------------------------
    Private Sub SaveOverlayAnchored(drawingBmp As Bitmap, baseFramePath As String, overlayPath As String)
        If drawingBmp Is Nothing Then Return
        If String.IsNullOrWhiteSpace(baseFramePath) Then Return

        Dim baseWidth As Integer = 1, baseHeight As Integer = 1
        Try
            Using fs As New FileStream(baseFramePath, FileMode.Open, FileAccess.Read, FileShare.Read)
                Using ms As New MemoryStream()
                    fs.CopyTo(ms)
                    ms.Position = 0
                    Using tmpImg As Image = Image.FromStream(ms)
                        baseWidth = Math.Max(1, tmpImg.Width)
                        baseHeight = Math.Max(1, tmpImg.Height)
                    End Using
                End Using
            End Using
        Catch
            baseWidth = Math.Max(1, drawingBmp.Width)
            baseHeight = Math.Max(1, drawingBmp.Height)
        End Try

        Using bmpToSave As New Bitmap(baseWidth, baseHeight, Imaging.PixelFormat.Format32bppArgb)
            Using g As Graphics = Graphics.FromImage(bmpToSave)
                g.Clear(Color.Transparent)
                Dim drawW = Math.Min(drawingBmp.Width, baseWidth)
                Dim drawH = Math.Min(drawingBmp.Height, baseHeight)
                g.DrawImage(drawingBmp, 0, 0, drawW, drawH)
            End Using
            SafeWriteBitmapAtomic(bmpToSave, overlayPath)
        End Using
    End Sub

    ' ---------------------------
    ' SafeDeleteFileEnhanced (VideoFBF)
    ' ---------------------------
    Private Function SafeDeleteFileEnhanced(path As String, Optional maxAttempts As Integer = 6) As Boolean
        If String.IsNullOrWhiteSpace(path) Then Return False
        If Not System.IO.File.Exists(path) Then Return True

        Dim attempts As Integer = 0
        Dim lastEx As Exception = Nothing

        While attempts < maxAttempts
            attempts += 1
            Try
                System.IO.File.Delete(path)
                Return True
            Catch ex As System.IO.IOException
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

        ' Prova a sovrascrivere con 1x1 PNG trasparente e poi cancellare
        Try
            Dim dir = System.IO.Path.GetDirectoryName(path)
            If String.IsNullOrWhiteSpace(dir) Then dir = System.IO.Path.GetTempPath()
            If Not System.IO.Directory.Exists(dir) Then System.IO.Directory.CreateDirectory(dir)

            Dim tmp = System.IO.Path.Combine(dir, System.IO.Path.GetFileNameWithoutExtension(path) & "_tmp" & System.IO.Path.GetExtension(path))
            Try
                Using emptyBmp As New System.Drawing.Bitmap(1, 1, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
                    Using g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(emptyBmp)
                        g.Clear(System.Drawing.Color.Transparent)
                    End Using
                    emptyBmp.Save(tmp, System.Drawing.Imaging.ImageFormat.Png)
                End Using

                If System.IO.File.Exists(path) Then
                    Try
                        System.IO.File.Replace(tmp, path, Nothing)
                    Catch ex As PlatformNotSupportedException
                        Try : System.IO.File.Delete(path) : Catch : End Try
                        System.IO.File.Move(tmp, path)
                    End Try
                Else
                    System.IO.File.Move(tmp, path)
                End If

                Try
                    System.IO.File.Delete(path)
                    Return True
                Catch
                    ' fallthrough: tenteremo altre strategie/logging
                End Try
            Finally
                If System.IO.File.Exists(tmp) Then
                    Try : System.IO.File.Delete(tmp) : Catch : End Try
                End If
            End Try
        Catch ex As Exception
            lastEx = ex
        End Try

        Try
            Dim logPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "VideoFBF_delete.log")
            Dim msg = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - SafeDeleteFileEnhanced failed for {path} : {If(lastEx IsNot Nothing, lastEx.Message, "unknown")}{Environment.NewLine}"
            System.IO.File.AppendAllText(logPath, msg)
        Catch
            ' ignore logging errors
        End Try

        Return False
    End Function


    ' ---------------------------
    ' SafeWriteBitmapAtomic (VideoFBF)
    ' ---------------------------
    Private Sub SafeWriteBitmapAtomic(bmp As System.Drawing.Bitmap, dstPath As String)
        If bmp Is Nothing OrElse String.IsNullOrWhiteSpace(dstPath) Then Return
        Dim dir = Path.GetDirectoryName(dstPath)
        If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
        Dim tmp = Path.Combine(dir, Path.GetFileNameWithoutExtension(dstPath) & "_tmp" & Path.GetExtension(dstPath))
        Try
            If File.Exists(tmp) Then
                Try : File.Delete(tmp) : Catch : End Try
            End If
            bmp.Save(tmp, Imaging.ImageFormat.Png)
            If File.Exists(dstPath) Then
                Try
                    File.Replace(tmp, dstPath, Nothing)
                Catch ex As PlatformNotSupportedException
                    If File.Exists(dstPath) Then
                        Try : File.Delete(dstPath) : Catch : End Try
                    End If
                    File.Move(tmp, dstPath)
                End Try
            Else
                File.Move(tmp, dstPath)
            End If
        Finally
            If File.Exists(tmp) Then
                Try : File.Delete(tmp) : Catch : End Try
            End If
        End Try
    End Sub

    Private Sub ChkZoom_CheckedChanged(sender As Object, e As EventArgs) Handles ChkZoom.CheckedChanged
        If ChkZoom.Checked Then
            SettapicFrame("AutoSize")
        Else
            SettapicFrame("Zoom")
        End If
    End Sub

    Private Sub SettaPicFrame(Settaggio As String)
        picFrame.Width = panelViewer.Width - 30
        picFrame.Height = panelViewer.Height - 30
        If Settaggio = "Zoom" Then
            picFrame.SizeMode = PictureBoxSizeMode.Zoom
            picFrame.Enabled = False
            picFrame.Cursor = Cursors.Default
            numSpessorePennino.Enabled = False
            CmbStrumento.Enabled = False
            btnColorePennino.Enabled = False
        Else
            picFrame.SizeMode = PictureBoxSizeMode.AutoSize
            picFrame.Enabled = True
            picFrame.Cursor = Cursors.Cross
            numSpessorePennino.Enabled = True
            CmbStrumento.Enabled = True
            btnColorePennino.Enabled = True
        End If
    End Sub

    ' Popola cmbStrumento (chiamare una sola volta)
    Private Sub InitStrumenti()
        CmbStrumento.Items.Clear()
        CmbStrumento.Items.Add("Nessuno")        ' index 0
        CmbStrumento.Items.Add("Punto")          ' index 1
        CmbStrumento.Items.Add("Linea")          ' index 2
        CmbStrumento.Items.Add("Ellisse")        ' index 3
        CmbStrumento.Items.Add("Rettangolo")     ' index 4
        CmbStrumento.SelectedIndex = 0
        AddHandler CmbStrumento.SelectedIndexChanged, AddressOf cmbStrumento_SelectedIndexChanged
    End Sub

    Private Sub cmbStrumento_SelectedIndexChanged(sender As Object, e As EventArgs)
        Select Case CmbStrumento.SelectedIndex
            Case 0 : currentTool = ToolType.None
            Case 1 : currentTool = ToolType.PointTool
            Case 2 : currentTool = ToolType.LineTool
            Case 3 : currentTool = ToolType.EllipseTool
            Case 4 : currentTool = ToolType.RectangleTool
            Case Else : currentTool = ToolType.None
        End Select
        ' Aggiorna cursore e abilitazioni
        If currentTool = ToolType.None Then
            picFrame.Cursor = Cursors.Default
        Else
            picFrame.Cursor = Cursors.Cross
        End If
    End Sub

    Private Sub picFrame_MouseDown(sender As Object, e As MouseEventArgs) Handles picFrame.MouseDown
        Try
            ' Protezioni base
            If picFrame Is Nothing Then Return
            If picFrame.Image Is Nothing Then Return
            If editor Is Nothing Then
                LogDebug("picFrame_MouseDown: editor = Nothing")
                MDIMessageBox.Show("Editor non inizializzato.", Me.MdiParent, MessageBoxButtons.OK)
                Return
            End If

            ' Linea/Ellisse/Rettangolo: inizia rubberband (coordinate immagine)
            If currentTool = ToolType.LineTool OrElse currentTool = ToolType.EllipseTool OrElse currentTool = ToolType.RectangleTool Then
                Dim imgPt As Point = ConvertMouseToImagePoint(e.Location)
                isRubberDragging = True
                rubberStart = imgPt
                rubberCurrent = imgPt
                If overlayRubber IsNot Nothing Then overlayRubber.Invalidate()
                editor.SaveState()
                Return
            End If

            ' Punto: disegna subito (coordinate immagine)
            If currentTool = ToolType.PointTool Then
                editor.SaveState()
                Dim imgPt As Point = ConvertMouseToImagePoint(e.Location)
                Using g As Graphics = Graphics.FromImage(editor.DrawingBitmap)
                    Using pen As New Pen(colorePennino, spessorePennino)
                        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                        g.DrawEllipse(pen, imgPt.X - spessorePennino, imgPt.Y - spessorePennino, spessorePennino * 2, spessorePennino * 2)
                    End Using
                End Using
                If picFrame.Image IsNot Nothing Then
                    Try : picFrame.Image.Dispose() : Catch : End Try
                End If
                picFrame.Image = CType(editor.DrawingBitmap.Clone(), Bitmap)
                hasUnsavedChanges = True
                editor.HasUnsavedChanges = True
                Return
            End If

            ' Fallback: pennino libero (comportamento esistente)
            isDrawing = True
            lastPoint = e.Location
            editor.SaveState()
        Catch ex As Exception
            LogDebug("picFrame_MouseDown EX: " & ex.Message)
        End Try
    End Sub

    Private Sub picFrame_MouseMove(sender As Object, e As MouseEventArgs) Handles picFrame.MouseMove
        Try
            If picFrame Is Nothing OrElse picFrame.Image Is Nothing Then Return

            ' Pennino libero
            If isDrawing Then
                Using g As Graphics = Graphics.FromImage(editor.DrawingBitmap)
                    Using penna As New Pen(colorePennino, spessorePennino)
                        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                        g.DrawLine(penna, lastPoint, ConvertMouseToImagePoint(e.Location))
                    End Using
                End Using
                lastPoint = ConvertMouseToImagePoint(e.Location)
                If picFrame.Image IsNot Nothing Then
                    Try : picFrame.Image.Dispose() : Catch : End Try
                End If
                picFrame.Image = CType(editor.DrawingBitmap.Clone(), Bitmap)
                Return
            End If

            ' Rubberband preview
            If isRubberDragging Then
                rubberCurrent = ConvertMouseToImagePoint(e.Location)
                If overlayRubber IsNot Nothing Then overlayRubber.Invalidate()
                Return
            End If
        Catch ex As Exception
            LogDebug("picFrame_MouseMove EX: " & ex.Message)
        End Try
    End Sub

    Private Sub picFrame_MouseUp(sender As Object, e As MouseEventArgs) Handles picFrame.MouseUp
        Try
            If picFrame Is Nothing OrElse picFrame.Image Is Nothing Then Return

            If isDrawing Then
                isDrawing = False
                hasUnsavedChanges = True
                editor.HasUnsavedChanges = True
                Return
            End If

            If isRubberDragging Then
                isRubberDragging = False
                rubberCurrent = ConvertMouseToImagePoint(e.Location)

                ' Commit su editor.DrawingBitmap (coordinate immagine)
                Using g As Graphics = Graphics.FromImage(editor.DrawingBitmap)
                    Using pen As New Pen(colorePennino, spessorePennino)
                        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                        Dim rect = GetNormalizedRect(rubberStart, rubberCurrent)
                        Select Case currentTool
                            Case ToolType.LineTool
                                g.DrawLine(pen, rubberStart, rubberCurrent)
                            Case ToolType.EllipseTool
                                g.DrawEllipse(pen, rect)
                            Case ToolType.RectangleTool
                                g.DrawRectangle(pen, rect)
                        End Select
                    End Using
                End Using

                If picFrame.Image IsNot Nothing Then
                    Try : picFrame.Image.Dispose() : Catch : End Try
                End If
                picFrame.Image = CType(editor.DrawingBitmap.Clone(), Bitmap)
                hasUnsavedChanges = True
                editor.HasUnsavedChanges = True
                If overlayRubber IsNot Nothing Then overlayRubber.Invalidate()
            End If
        Catch ex As Exception
            LogDebug("picFrame_MouseUp EX: " & ex.Message)
        End Try
    End Sub

    Private Sub OverlayRubber_MouseDown(sender As Object, e As MouseEventArgs)
        Try
            If picFrame Is Nothing OrElse picFrame.Image Is Nothing Then Return
            If editor Is Nothing Then Return

            If currentTool = ToolType.LineTool OrElse currentTool = ToolType.EllipseTool OrElse currentTool = ToolType.RectangleTool Then
                Dim imgPt = ConvertMouseToImagePoint(e.Location)
                isRubberDragging = True
                rubberStart = imgPt
                rubberCurrent = imgPt
                overlayRubber.Invalidate()
                editor.SaveState()
                Return
            End If

            If currentTool = ToolType.PointTool Then
                editor.SaveState()
                Dim imgPt = ConvertMouseToImagePoint(e.Location)
                Using g As Graphics = Graphics.FromImage(editor.DrawingBitmap)
                    Using pen As New Pen(colorePennino, spessorePennino)
                        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                        g.DrawEllipse(pen, imgPt.X - spessorePennino, imgPt.Y - spessorePennino, spessorePennino * 2, spessorePennino * 2)
                    End Using
                End Using
                If picFrame.Image IsNot Nothing Then Try : picFrame.Image.Dispose() : Catch : End Try
                picFrame.Image = CType(editor.DrawingBitmap.Clone(), Bitmap)
                hasUnsavedChanges = True
                editor.HasUnsavedChanges = True
                Return
            End If

            ' fallback: pennino libero
            isDrawing = True
            lastPoint = ConvertMouseToImagePoint(e.Location)
            editor.SaveState()
        Catch ex As Exception
            LogDebug("OverlayRubber_MouseDown EX: " & ex.Message)
        End Try
    End Sub

    Private Sub OverlayRubber_MouseMove(sender As Object, e As MouseEventArgs)
        Try
            If picFrame Is Nothing OrElse picFrame.Image Is Nothing Then Return

            If isDrawing Then
                Dim imgPt = ConvertMouseToImagePoint(e.Location)
                Using g As Graphics = Graphics.FromImage(editor.DrawingBitmap)
                    Using penna As New Pen(colorePennino, spessorePennino)
                        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                        g.DrawLine(penna, lastPoint, imgPt)
                    End Using
                End Using
                lastPoint = imgPt
                If picFrame.Image IsNot Nothing Then Try : picFrame.Image.Dispose() : Catch : End Try
                picFrame.Image = CType(editor.DrawingBitmap.Clone(), Bitmap)
                Return
            End If

            If isRubberDragging Then
                rubberCurrent = ConvertMouseToImagePoint(e.Location)
                overlayRubber.Invalidate()
                Return
            End If
        Catch ex As Exception
            LogDebug("OverlayRubber_MouseMove EX: " & ex.Message)
        End Try
    End Sub

    Private Sub OverlayRubber_MouseUp(sender As Object, e As MouseEventArgs)
        Try
            If picFrame Is Nothing OrElse picFrame.Image Is Nothing Then Return

            If isDrawing Then
                isDrawing = False
                hasUnsavedChanges = True
                editor.HasUnsavedChanges = True
                Return
            End If

            If isRubberDragging Then
                isRubberDragging = False
                rubberCurrent = ConvertMouseToImagePoint(e.Location)

                Using g As Graphics = Graphics.FromImage(editor.DrawingBitmap)
                    Using pen As New Pen(colorePennino, spessorePennino)
                        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                        Dim rect = GetNormalizedRect(rubberStart, rubberCurrent)
                        Select Case currentTool
                            Case ToolType.LineTool
                                g.DrawLine(pen, rubberStart, rubberCurrent)
                            Case ToolType.EllipseTool
                                g.DrawEllipse(pen, rect)
                            Case ToolType.RectangleTool
                                g.DrawRectangle(pen, rect)
                        End Select
                    End Using
                End Using

                If picFrame.Image IsNot Nothing Then Try : picFrame.Image.Dispose() : Catch : End Try
                picFrame.Image = CType(editor.DrawingBitmap.Clone(), Bitmap)
                hasUnsavedChanges = True
                editor.HasUnsavedChanges = True
                overlayRubber.Invalidate()
            End If
        Catch ex As Exception
            LogDebug("OverlayRubber_MouseUp EX: " & ex.Message)
        End Try
    End Sub


    Private Function PictureBoxToImage(point As Point, pb As PictureBox) As Point
        If pb.Image Is Nothing Then Return point
        Dim img = pb.Image
        Dim imgRatio = img.Width / img.Height
        Dim pbRatio = pb.Width / pb.Height
        Dim scale As Double
        Dim offsetX As Integer = 0
        Dim offsetY As Integer = 0
        If imgRatio > pbRatio Then
            scale = pb.Width / img.Width
            offsetY = CInt((pb.Height - img.Height * scale) / 2)
        Else
            scale = pb.Height / img.Height
            offsetX = CInt((pb.Width - img.Width * scale) / 2)
        End If
        Dim ix = CInt((point.X - offsetX) / scale)
        Dim iy = CInt((point.Y - offsetY) / scale)
        Return New Point(Math.Max(0, ix), Math.Max(0, iy))
    End Function

    Private Sub InitOverlayRubber()
        Try
            If picFrame Is Nothing Then
                LogDebug("InitOverlayRubber: picFrame = Nothing")
                Return
            End If
            If overlayRubber IsNot Nothing Then Return

            overlayRubber = New OverlayPanel() With {
            .Name = "overlayRubber",
            .Dock = DockStyle.Fill,
            .BackColor = Color.Transparent,
            .Visible = True,
            .Enabled = True
        }

            picFrame.Controls.Add(overlayRubber)
            overlayRubber.BringToFront()

            AddHandler overlayRubber.Paint, AddressOf OverlayRubber_Paint

            ' Gestori mouse sull'overlay (ora l'overlay riceve i mouse)
            AddHandler overlayRubber.MouseDown, AddressOf OverlayRubber_MouseDown
            AddHandler overlayRubber.MouseMove, AddressOf OverlayRubber_MouseMove
            AddHandler overlayRubber.MouseUp, AddressOf OverlayRubber_MouseUp

            AddHandler picFrame.Resize, Sub(s, e)
                                            If overlayRubber IsNot Nothing Then overlayRubber.Invalidate()
                                        End Sub
        Catch ex As Exception
            LogDebug("InitOverlayRubber ERROR: " & ex.Message)
        End Try
    End Sub

    Private Sub OverlayRubber_Paint(sender As Object, e As PaintEventArgs)
        Try
            If Not isRubberDragging Then Return
            Dim g = e.Graphics
            g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            Using pen As New Pen(colorePennino, Math.Max(1, spessorePennino))
                pen.DashStyle = Drawing2D.DashStyle.Dash
                Dim p1 = ConvertImageToPictureBoxPoint(rubberStart)
                Dim p2 = ConvertImageToPictureBoxPoint(rubberCurrent)
                Dim rect = GetNormalizedRect(p1, p2)
                Select Case currentTool
                    Case ToolType.LineTool
                        g.DrawLine(pen, p1, p2)
                    Case ToolType.EllipseTool
                        g.DrawEllipse(pen, rect)
                    Case ToolType.RectangleTool
                        g.DrawRectangle(pen, rect)
                End Select
            End Using
        Catch ex As Exception
            LogDebug("OverlayRubber_Paint ERROR: " & ex.Message)
        End Try
    End Sub


    Private Function ConvertMouseToImagePoint(mousePt As Point) As Point
        If picFrame Is Nothing OrElse picFrame.Image Is Nothing Then Return mousePt
        Dim img = picFrame.Image
        Dim imgRatio = img.Width / img.Height
        Dim pbRatio = picFrame.Width / picFrame.Height
        Dim scale As Double
        Dim offsetX As Integer = 0
        Dim offsetY As Integer = 0
        If imgRatio > pbRatio Then
            scale = picFrame.Width / img.Width
            offsetY = CInt((picFrame.Height - img.Height * scale) / 2)
        Else
            scale = picFrame.Height / img.Height
            offsetX = CInt((picFrame.Width - img.Width * scale) / 2)
        End If
        Dim ix = CInt((mousePt.X - offsetX) / scale)
        Dim iy = CInt((mousePt.Y - offsetY) / scale)
        ix = Math.Max(0, Math.Min(img.Width - 1, ix))
        iy = Math.Max(0, Math.Min(img.Height - 1, iy))
        Return New Point(ix, iy)
    End Function

    Private Function ConvertImageToPictureBoxPoint(imgPt As Point) As Point
        If picFrame Is Nothing OrElse picFrame.Image Is Nothing Then Return imgPt
        Dim img = picFrame.Image
        Dim imgRatio = img.Width / img.Height
        Dim pbRatio = picFrame.Width / picFrame.Height
        Dim scale As Double
        Dim offsetX As Integer = 0
        Dim offsetY As Integer = 0
        If imgRatio > pbRatio Then
            scale = picFrame.Width / img.Width
            offsetY = CInt((picFrame.Height - img.Height * scale) / 2)
        Else
            scale = picFrame.Height / img.Height
            offsetX = CInt((picFrame.Width - img.Width * scale) / 2)
        End If
        Dim px = CInt(imgPt.X * scale + offsetX)
        Dim py = CInt(imgPt.Y * scale + offsetY)
        Return New Point(px, py)
    End Function

    Private Function GetNormalizedRect(p1 As Point, p2 As Point) As Rectangle
        Dim x = Math.Min(p1.X, p2.X)
        Dim y = Math.Min(p1.Y, p2.Y)
        Dim w = Math.Abs(p1.X - p2.X)
        Dim h = Math.Abs(p1.Y - p2.Y)
        Return New Rectangle(x, y, w, h)
    End Function


    Private Sub LogDebug(msg As String)
        Try
            Dim logPath = IO.Path.Combine(IO.Path.GetTempPath(), "VideoFBF_debug.log")
            IO.File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {msg}{Environment.NewLine}")
        Catch : End Try
    End Sub

    ' --- DTO ListView Note ---
    Public Class NotaFrameInfo
        Public Property FrameIndex As Integer
        Public Property TestoNota As String
        Public Property Autore As String
        Public Property DataNota As DateTime
    End Class

    Public Class OverlayPanel
        Inherits Panel

        Public Sub New()
            MyBase.New()
            ' Abilita double buffering e painting ottimizzato
            Me.SetStyle(ControlStyles.OptimizedDoubleBuffer Or
                        ControlStyles.AllPaintingInWmPaint Or
                        ControlStyles.UserPaint Or
                        ControlStyles.SupportsTransparentBackColor, True)
            Me.UpdateStyles()
            Me.BackColor = Color.Transparent
        End Sub
    End Class

End Class
