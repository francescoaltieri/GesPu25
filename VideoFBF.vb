Imports System.IO
Imports Microsoft.Data.SqlClient
Imports System.Threading.Tasks

Public Class VideoFBF
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

    ' Flag che indica modifiche non salvate sul frame corrente
    Private hasUnsavedChanges As Boolean = False

    Public Property Parametri As RevisioneParametri

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
    End Sub

    Private Sub CaricaUtentiCondivisi()
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

    Private Sub CaricaUtentiDisponibili()
        CaricaUtentiCondivisi()
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

    Private Sub VideoFBF_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se ci sono modifiche non salvate, chiedi conferma
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
            Application.DoEvents()

            Dim videoPath = OpenFileDialog1.FileName
            Dim nomeVideo = Path.GetFileNameWithoutExtension(videoPath)
            Dim baseDir = Path.Combine(OttieniPercorsoFrames(), nomeVideo)
            Dim revisioneID As Integer
            Dim Approvato As Boolean = False

            Dim videoID = Await Task.Run(Function()
                                             Dim id = OttieniVideoID(nomeVideo)
                                             If id = -1 Then
                                                 id = InserisciVideo(nomeVideo, videoPath)
                                                 InserisciRevisioneZero(id, nomeVideo)
                                                 revisioneID = 1
                                             Else
                                                 revisioneID = OttieniProssimoRevisioneIDPerVideo(id)
                                             End If
                                             Return id
                                         End Function)

            Dim revisioneDir = Path.Combine(baseDir, $"Revisione_{revisioneID:D4}")
            Directory.CreateDirectory(baseDir)
            Directory.CreateDirectory(revisioneDir)

            Dim tempEditor As New VideoEditor(videoPath, revisioneDir)
            Await Task.Run(Sub() tempEditor.ExtractFrames())

            Await Task.Run(Sub()
                               InserisciRevisione(videoID, revisioneID, SessioneUtente.NomeUtenteCorrente, "bozza")
                               InserisciPermessoUtente(revisioneID, SessioneUtente.NomeUtenteCorrente)
                           End Sub)

            MessageBox.Show($"Video caricato, frame estratti e Revisione_{revisioneID:D4} registrata.", "Operazione completata", MessageBoxButtons.OK, MessageBoxIcon.Information)

            lblRevAttiva.Text = revisioneID.ToString()

            Dim parametri = New RevisioneParametri(
                videoID,
                revisioneID,
                SessioneUtente.NomeUtenteCorrente,
                $"Revisione {revisioneID}",
                "visualizza",
                DateTime.Now,
                Approvato
            )

            editor = tempEditor
            picFrame.Image = editor.LoadFrame(0)

            RemoveHandler TrackFrame.Scroll, AddressOf trackFrame_Scroll
            TrackFrame.Minimum = 0
            TrackFrame.Maximum = editor.FrameList.Count - 1
            TrackFrame.Value = 0
            AddHandler TrackFrame.Scroll, AddressOf trackFrame_Scroll

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

    Private Function OttieniRevisioneZeroID(videoID As Integer) As Integer
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query = "SELECT TOP 1 RevisioneID FROM Mov_Revisioni WHERE VideoID = @VideoID AND Note LIKE 'Revisione 0%'"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.Add("@VideoID", SqlDbType.Int).Value = videoID
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
                INSERT INTO Mov_Revisione (RevisioneID, VideoID, Autore, DataRevisione, Note, Stato)
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

    Private Sub InserisciRevisione(videoId As Integer, RevisioneID As Integer, nomeUtente As String, stato As String)
        Dim nomeRevisione = $"Revisione {RevisioneID} - {DateTime.Now:dd/MM/yyyy}"
        Dim NumRetake = CalcolaRetake(videoId)

        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Using tran = conn.BeginTransaction()
                Try
                    Dim query = "INSERT INTO Mov_Revisioni (RevisioneID, VideoID, Autore, DataRevisione, NumRetake, Note, Stato)
                                 VALUES (@RevisioneID, @VideoID, @Autore, @Data, @NumRetake, @Note, @Stato)"
                    Using cmd As New SqlCommand(query, conn, tran)
                        cmd.Parameters.Add("@RevisioneID", SqlDbType.Int).Value = RevisioneID
                        cmd.Parameters.Add("@VideoID", SqlDbType.Int).Value = videoId
                        cmd.Parameters.Add("@Autore", SqlDbType.NVarChar, 100).Value = nomeUtente
                        cmd.Parameters.Add("@Data", SqlDbType.DateTime).Value = DateTime.Now
                        cmd.Parameters.Add("@NumRetake", SqlDbType.Int).Value = NumRetake
                        cmd.Parameters.Add("@Note", SqlDbType.NVarChar, 500).Value = nomeRevisione
                        cmd.Parameters.Add("@Stato", SqlDbType.NVarChar, 50).Value = stato
                        cmd.ExecuteNonQuery()
                    End Using

                    tran.Commit()
                Catch ex As Exception
                    tran.Rollback()
                    Throw
                End Try
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

    Private Sub DisabilitaControlliModifica()
        SetControlsEnabled(False)
    End Sub

    Private Sub AbilitaControlliModifica()
        SetControlsEnabled(True)
    End Sub

    Private Sub SetControlsEnabled(enabled As Boolean)
        btnAnnulla.Enabled = enabled
        btnColorePennino.Enabled = enabled
        btnSalvaNote.Enabled = enabled
        btnSalvaVideo.Enabled = enabled
        numSpessorePennino.Enabled = enabled
        txtNote.Enabled = enabled
    End Sub

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
                MDIMessageBox.Show($"Non è possibile ricreare la Revisione {revisioneSuccessiva} perché esistono revisioni successive ({elenco}).\nElimina prima tutte le revisioni superiori.", Me.MdiParent, MessageBoxButtons.OK, "Catena incoerente")
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
    Private Sub trackFrame_Scroll(sender As Object, e As EventArgs) Handles TrackFrame.Scroll
        If editor Is Nothing Then Return

        If Not ConfirmSaveChanges() Then
            RemoveHandler TrackFrame.Scroll, AddressOf trackFrame_Scroll
            TrackFrame.Value = editor.CurrentIndex
            AddHandler TrackFrame.Scroll, AddressOf trackFrame_Scroll
            Return
        End If

        Dim idx = TrackFrame.Value

        ' Impostiamo prima l'indice nell'editor (coerenza)
        editor.CurrentIndex = idx

        ' Carichiamo il frame (LoadFrame NON deve modificare CurrentIndex)
        picFrame.Image = editor.LoadFrame(idx)

        ' Aggiorniamo i campi correlati (nota, autore, tempo)
        AggiornaFrameCorrente(idx)
    End Sub

    Private Sub btnSuccessivo_Click(sender As Object, e As EventArgs) Handles btnSuccessivo.Click
        If editor Is Nothing OrElse picFrame Is Nothing OrElse TrackFrame Is Nothing OrElse txtNote Is Nothing Then
            MessageBox.Show("Impossibile caricare il frame.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If editor.CurrentIndex < editor.FrameList.Count - 1 Then
            If Not ConfirmSaveChanges() Then Return

            Dim nuovoIndex = editor.CurrentIndex + 1
            picFrame.Image = editor.LoadFrame(nuovoIndex)
            editor.CurrentIndex = nuovoIndex
            TrackFrame.Value = nuovoIndex

            Dim revisioneID As Integer
            If TryParseRevisioneID(lblRevAttiva.Text, revisioneID) Then
                txtNote.Text = RecuperaNotaDaDatabase(revisioneID, nuovoIndex)
            Else
                txtNote.Text = ""
            End If
        End If
    End Sub


    Private Sub btnPrecedente_Click(sender As Object, e As EventArgs) Handles btnPrecedente.Click
        If editor Is Nothing OrElse picFrame Is Nothing OrElse TrackFrame Is Nothing OrElse txtNote Is Nothing Then
            MessageBox.Show("Impossibile caricare il frame.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If editor.CurrentIndex > 0 Then
            If Not ConfirmSaveChanges() Then Return

            Dim nuovoIndex = editor.CurrentIndex - 1
            picFrame.Image = editor.LoadFrame(nuovoIndex)
            editor.CurrentIndex = nuovoIndex
            TrackFrame.Value = nuovoIndex

            Dim revisioneID As Integer
            If TryParseRevisioneID(lblRevAttiva.Text, revisioneID) Then
                txtNote.Text = RecuperaNotaDaDatabase(revisioneID, nuovoIndex)
            Else
                txtNote.Text = ""
            End If
        End If
    End Sub


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
        ' Dopo rilascio, se è stato disegnato qualcosa, segna come non salvato
        hasUnsavedChanges = True
        editor.HasUnsavedChanges = True
    End Sub

    Private Function RevisioneModificabile() As Boolean
        Return Parametri IsNot Nothing AndAlso Parametri.RevisioneID <> 0 AndAlso IsRevisioneModificabile(Parametri.RevisioneID)
    End Function

    Private Sub btnAnnulla_Click(sender As Object, e As EventArgs) Handles btnAnnulla.Click
        If picFrame.Image Is Nothing Then Return

        Dim ancoraModifiche = editor.Undo()
        picFrame.Image = CType(editor.DrawingBitmap.Clone(), Bitmap)

        ' Aggiorna flag in base allo stato dell'editor
        hasUnsavedChanges = editor.HasUnsavedChanges
    End Sub

    Private Sub SalvaFrame()
        editor.SaveFrame()
        ' Dopo il salvataggio su disco, resetta lo stato dirty
        hasUnsavedChanges = False
        editor.HasUnsavedChanges = False
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

    Private Sub btnPrimoFrame_Click(sender As Object, e As EventArgs) Handles btnPrimoFrame.Click
        If editor Is Nothing OrElse picFrame Is Nothing OrElse TrackFrame Is Nothing OrElse txtNote Is Nothing Then
            MessageBox.Show("Impossibile caricare il frame.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If Not ConfirmSaveChanges() Then Return

        Dim primoIndex = 0
        picFrame.Image = editor.LoadFrame(primoIndex)
        editor.CurrentIndex = primoIndex
        TrackFrame.Value = primoIndex

        Dim revisioneID As Integer
        If TryParseRevisioneID(lblRevAttiva.Text, revisioneID) Then
            txtNote.Text = RecuperaNotaDaDatabase(revisioneID, primoIndex)
        Else
            txtNote.Text = ""
        End If
    End Sub

    Private Sub btnUltimoFrame_Click(sender As Object, e As EventArgs) Handles btnUltimoFrame.Click
        If editor Is Nothing OrElse picFrame Is Nothing OrElse TrackFrame Is Nothing OrElse txtNote Is Nothing Then
            MessageBox.Show("Impossibile caricare il frame.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If Not ConfirmSaveChanges() Then Return

        Dim ultimoIndex = editor.FrameList.Count - 1
        picFrame.Image = editor.LoadFrame(ultimoIndex)
        editor.CurrentIndex = ultimoIndex
        TrackFrame.Value = ultimoIndex

        Dim revisioneID As Integer
        If TryParseRevisioneID(lblRevAttiva.Text, revisioneID) Then
            txtNote.Text = RecuperaNotaDaDatabase(revisioneID, ultimoIndex)
        Else
            txtNote.Text = ""
        End If
    End Sub


    Private Sub btnAvantiVeloce_Click(sender As Object, e As EventArgs) Handles btnAvantiVeloce.Click
        If editor Is Nothing OrElse picFrame Is Nothing OrElse TrackFrame Is Nothing OrElse txtNote Is Nothing Then
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
        If editor Is Nothing OrElse picFrame Is Nothing OrElse TrackFrame Is Nothing OrElse txtNote Is Nothing Then
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
            ' Se ci sono modifiche non salvate, chiedi conferma
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
            TrackFrame.Value = nuovoIndex

            Dim revisioneID As Integer
            If TryParseRevisioneID(lblRevAttiva.Text, revisioneID) Then
                txtNote.Text = RecuperaNotaDaDatabase(revisioneID, nuovoIndex)
            Else
                txtNote.Text = ""
            End If

            Application.DoEvents()
            Await Task.Delay(60)
        End While
    End Sub

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        If keyData = Keys.Escape Then
            autoScrollActive = False
            btnAvantiVeloce.Text = testoAvantiOriginale
            btnIndietroVeloce.Text = testoIndietroOriginale
            Return True
        End If
        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

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

    Private Sub btnAggiungiNote_Click_1(sender As Object, e As EventArgs)
        If notaPosizione = Point.Empty Then
            MessageBox.Show("Clicca sul frame per scegliere la posizione della nota.")
            Return
        End If

        If String.IsNullOrWhiteSpace(txtNote.Text) Then
            MessageBox.Show("Inserisci del testo nella nota.")
            Return
        End If

        editor.SaveState()

        Using g = Graphics.FromImage(editor.DrawingBitmap)
            Dim font = New Font("Arial", 14, FontStyle.Regular)
            Dim padding = 8

            Dim textSize = g.MeasureString(txtNote.Text, font, 400)
            Dim boxWidth = CInt(textSize.Width) + padding * 2
            Dim boxHeight = CInt(textSize.Height) + padding * 2

            Dim x = notaPosizione.X
            Dim y = notaPosizione.Y

            If x + boxWidth > editor.DrawingBitmap.Width Then
                x = editor.DrawingBitmap.Width - boxWidth - 10
            End If
            If y + boxHeight > editor.DrawingBitmap.Height Then
                y = editor.DrawingBitmap.Height - boxHeight - 10
            End If

            Dim rect = New Rectangle(x, y, boxWidth, boxHeight)

            g.FillRectangle(New SolidBrush(Color.FromArgb(180, Color.Black)), rect)

            Dim format As New StringFormat
            format.Alignment = StringAlignment.Near
            format.LineAlignment = StringAlignment.Near
            format.FormatFlags = StringFormatFlags.LineLimit

            g.DrawString(txtNote.Text, font, Brushes.White, rect, format)
        End Using

        picFrame.Image = CType(editor.DrawingBitmap.Clone, Bitmap)
        notaPosizione = Point.Empty

        ' Segna come non salvato
        hasUnsavedChanges = True
        editor.HasUnsavedChanges = True
    End Sub

    Private Sub btnSalvaNote_Click(sender As Object, e As EventArgs) Handles btnSalvaNote.Click
        If picFrame.Image Is Nothing Then Return

        Dim frameIndex = editor.CurrentIndex
        Dim nota = txtNote.Text.Trim()
        Dim nomeUtente = NomeUtenteCorrente
        Dim revisioneID As Integer
        If Not TryParseRevisioneID(lblRevAttiva.Text, revisioneID) Then
            MessageBox.Show("Revisione non valida.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Try
            ' Salva il disegno/appunti sul frame
            editor.SaveFrame()
            hasUnsavedChanges = False
            editor.HasUnsavedChanges = False

            ' Salva/aggiorna la nota nel DB
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

            AggiornaNoteDaDatabase(revisioneID)
            MessageBox.Show("Nota e appunti salvati.", "Salvataggio completato", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Errore durante il salvataggio: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CaricaListaNote()
        lstNoteFrame.Items.Clear()
        Dim revisioneID As Integer
        If Not TryParseRevisioneID(lblRevAttiva.Text, revisioneID) Then Return

        Using conn As New SqlConnection(ConnString)
            Dim cmd As New SqlCommand("
                SELECT FrameIndex, TestoNota, NomeUtente, DataNota
                FROM Mov_FrameNote
                WHERE RevisioneID = @RevID
                ORDER BY FrameIndex", conn)
            cmd.Parameters.Add("@RevID", SqlDbType.Int).Value = revisioneID
            conn.Open()

            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim frameIndex = Convert.ToInt32(reader("FrameIndex"))
                    Dim testo = reader("TestoNota").ToString()
                    Dim autore = reader("NomeUtente").ToString()
                    Dim data = Convert.ToDateTime(reader("DataNota"))

                    Dim anteprima = If(testo.Length > 30, testo.Substring(0, 30) & "...", testo)
                    Dim voce = $"Frame {frameIndex + 1}: {anteprima}"

                    Dim item As New ListViewItem(voce)
                    item.Tag = New NotaFrameInfo With {
                        .frameIndex = frameIndex,
                        .TestoNota = testo,
                        .autore = autore,
                        .DataNota = data
                    }
                    item.ToolTipText = $"Autore: {autore}{Environment.NewLine}Data: {data:dd/MM/yyyy HH:mm}"
                    lstNoteFrame.Items.Add(item)
                End While
            End Using
        End Using
    End Sub

    Private Sub lstNote_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Delete Then
            EliminaNotaSelezionata()
        End If
    End Sub

    Private Sub EliminaNotaSelezionata()
        If lstNoteFrame.SelectedItems.Count = 0 Then Return

        Dim info = CType(lstNoteFrame.SelectedItems(0).Tag, NotaFrameInfo)
        Dim revisioneID = Int(lblRevAttiva.Text)
        Dim frameIndex = info.FrameIndex
        Dim testoNota = info.TestoNota

        Dim conferma = MessageBox.Show("Vuoi davvero eliminare questa nota?", "Conferma eliminazione", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

        If conferma = DialogResult.Yes Then
            Using conn As New SqlConnection(ConnString)
                conn.Open()
                Dim cmd As New SqlCommand("
                    DELETE FROM Mov_FrameNote 
                    WHERE RevisioneID = @RevID AND FrameIndex = @FrameIndex AND TestoNota = @TestoNota", conn)
                cmd.Parameters.Add("@RevID", SqlDbType.Int).Value = revisioneID
                cmd.Parameters.Add("@FrameIndex", SqlDbType.Int).Value = frameIndex
                cmd.Parameters.Add("@TestoNota", SqlDbType.NVarChar, 2000).Value = testoNota
                cmd.ExecuteNonQuery()
            End Using

            AggiornaNoteDaDatabase(revisioneID)
            txtNote.Clear()
            lblAutore.Text = ""
            lblDataNota.Text = ""
        End If
    End Sub

    Private Sub lstNoteFrame_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstNoteFrame.SelectedIndexChanged
        If lstNoteFrame.SelectedItems.Count = 0 Then Return

        Dim item = lstNoteFrame.SelectedItems(0)
        Dim info = TryCast(item.Tag, NotaFrameInfo)
        If info Is Nothing Then Return

        ' Se ci sono modifiche non salvate, chiedi conferma prima di cambiare frame
        If Not ConfirmSaveChanges() Then
            ' Ripristina selezione al frame corrente per evitare incoerenze
            For Each it As ListViewItem In lstNoteFrame.Items
                Dim inf = TryCast(it.Tag, NotaFrameInfo)
                If inf IsNot Nothing AndAlso inf.FrameIndex = editor.CurrentIndex Then
                    it.Selected = True
                    Exit For
                End If
            Next
            Return
        End If

        Dim frameIndex = info.FrameIndex

        ' Protezioni
        If editor Is Nothing Then Return
        If frameIndex < 0 OrElse frameIndex >= editor.FrameList.Count Then Return

        ' Imposta indice e UI
        editor.CurrentIndex = frameIndex
        TrackFrame.Value = frameIndex
        picFrame.Image = editor.LoadFrame(frameIndex)

        ' Aggiorna campi nota/autore/data
        txtNote.Text = info.TestoNota
        lblAutore.Text = info.Autore
        lblDataNota.Text = $"{info.DataNota:dd/MM/yyyy HH:mm}"

        ' Considera il frame pulito dopo il caricamento
        hasUnsavedChanges = False
        editor.HasUnsavedChanges = False
    End Sub


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

        TrackFrame.Minimum = 0
        TrackFrame.Maximum = editor.FrameList.Count - 1
        TrackFrame.Value = 0
        picFrame.Image = editor.LoadFrame(0)

        txtNote.Text = RecuperaNotaDaDatabase(revisioneID, 0)

        ' Reset stato dirty all'apertura revisione
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
        FrameConNote.Clear()
        lstNoteFrame.Items.Clear()

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

                        If Not FrameConNote.Contains(frameIndex) Then
                            FrameConNote.Add(frameIndex)
                        End If

                        Dim info As New NotaFrameInfo With {
                            .frameIndex = frameIndex,
                            .TestoNota = testo,
                            .autore = autore,
                            .DataNota = data
                        }

                        Dim anteprima = If(testo.Length > 30, testo.Substring(0, 30) & "...", testo)
                        Dim voce = $"Frame {frameIndex + 1}: {anteprima}"

                        Dim item As New ListViewItem(voce) With {
                            .Tag = info,
                            .ToolTipText = $"Autore: {autore}{Environment.NewLine}Data: {data:dd/MM/yyyy HH:mm}"
                        }

                        lstNoteFrame.Items.Add(item)
                    End While
                End Using
            End Using
        End Using

        Dim overlay = Me.Controls.Find("OverlayNotePanel", True).FirstOrDefault()
        If overlay Is Nothing Then
            Dim nuovoOverlay As New Panel With {
                .Width = TrackFrame.Width,
                .Height = 10,
                .Location = New Point(TrackFrame.Left, TrackFrame.Top - 10),
                .BackColor = Color.Transparent,
                .Name = "OverlayNotePanel"
            }
            Me.Controls.Add(nuovoOverlay)
            AddHandler nuovoOverlay.Paint, AddressOf DisegnaSegnaliniNote
        Else
            overlay.Invalidate()
        End If
    End Sub

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

    Private Sub btnRetake_Click(sender As Object, e As EventArgs) Handles btnRetake.Click
        If picFrame.Image Is Nothing Then
            MessageBox.Show("Caricare prima i Frames", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Dim revisioneID As Integer
        If Not TryParseRevisioneID(Me.lblRevAttiva.Text, revisioneID) Then Return

        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query As String = "
                UPDATE Mov_Revisione
                SET Stato = 'Non Conforme', Approvato = 0
                WHERE RevisioneID = @RevisioneID"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.Add("@RevisioneID", SqlDbType.Int).Value = revisioneID
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Private Sub btnApprovazione_Click(sender As Object, e As EventArgs) Handles btnApprovazione.Click
        If picFrame.Image Is Nothing Then
            MessageBox.Show("Caricare prima i Frames", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Dim revisioneID As Integer
        If Not TryParseRevisioneID(Me.lblRevAttiva.Text, revisioneID) Then Return

        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query As String = "
                UPDATE Mov_Revisione
                SET Stato = 'Conforme', Approvato = 1
                WHERE RevisioneID = @RevisioneID"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.Add("@RevisioneID", SqlDbType.Int).Value = revisioneID
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ' Conferma salvataggio modifiche non salvate
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

                ' Salva disegni
                editor.SaveFrame()
                hasUnsavedChanges = False
                editor.HasUnsavedChanges = False

                ' Salva nota
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
            ' Scarta modifiche: ricarica il frame originale
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
            ' Cancel
            Return False
        End If
    End Function

    Public Class NotaFrameInfo
        Public Property FrameIndex As Integer
        Public Property TestoNota As String
        Public Property Autore As String
        Public Property DataNota As DateTime
    End Class

End Class
