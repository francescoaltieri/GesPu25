Imports System.IO
Imports System.Text.RegularExpressions
Imports Microsoft.Data.SqlClient
Imports Microsoft.VisualBasic.Devices

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
        RipristinaPosizioneForm()

        lblRevAttiva.Text = "Nessuna revisione attiva"
        CaricaUtentiDisponibili()

    End Sub

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
                        ' Visualizza Generalità, memorizza NomeUtente
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
                           Dim revisioneID = Int(lblRevAttiva.Text)

                           If e.NewValue = CheckState.Checked Then
                               AggiungiCondivisioneUtente(revisioneID, nomeUtente)
                           ElseIf e.NewValue = CheckState.Unchecked Then
                               RimuoviCondivisioneUtente(revisioneID, nomeUtente)
                           End If
                       End Sub)
    End Sub

    Public Sub AggiornaUtentiCondivisi(revisioneID As Integer)
        aggiornamentoInCorso = True

        ' Deseleziona tutto
        For i = 0 To lstUtentiCondivisi.Items.Count - 1
            lstUtentiCondivisi.SetItemChecked(i, False)
        Next

        ' Seleziona gli utenti condivisi
        Using conn As New SqlConnection(ConnString)
            Dim cmd As New SqlCommand("
            SELECT NomeUtente 
            FROM Mov_UtenteRevisione 
            WHERE RevisioneID = @RevID", conn)
            cmd.Parameters.AddWithValue("@RevID", revisioneID)
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

            ' Verifica se esiste già
            Dim checkCmd As New SqlCommand("
            SELECT COUNT(*) 
            FROM Mov_UtenteRevisione 
            WHERE RevisioneID = @RevID AND NomeUtente = @NomeUtente", conn)
            checkCmd.Parameters.AddWithValue("@RevID", revisioneID)
            checkCmd.Parameters.AddWithValue("@NomeUtente", nomeUtente)

            Dim esiste = Convert.ToInt32(checkCmd.ExecuteScalar()) > 0
            If esiste Then Exit Sub

            ' Inserisci solo se non esiste
            Dim insertCmd As New SqlCommand("
            INSERT INTO Mov_UtenteRevisione (RevisioneID, NomeUtente, Permesso)
            VALUES (@RevID, @NomeUtente, 'visualizza')", conn)
            insertCmd.Parameters.AddWithValue("@RevID", revisioneID)
            insertCmd.Parameters.AddWithValue("@NomeUtente", nomeUtente)
            insertCmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub RimuoviCondivisioneUtente(revisioneID As Integer, nomeUtente As String)
        Using conn As New SqlConnection(ConnString)
            conn.Open()

            Dim cmd As New SqlCommand("
            DELETE FROM Mov_UtenteRevisione 
            WHERE RevisioneID = @RevID AND NomeUtente = @NomeUtente", conn)
            cmd.Parameters.AddWithValue("@RevID", revisioneID)
            cmd.Parameters.AddWithValue("@NomeUtente", nomeUtente)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub RipristinaPosizioneForm()
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query = "SELECT X, Y, Width, Height, WindowsState FROM Sys_Form WHERE FormName = @FormName"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@FormName", Me.Name)
                Using reader = cmd.ExecuteReader()
                    If reader.Read() Then
                        Me.StartPosition = FormStartPosition.Manual
                        Me.Location = New Point(reader("X"), reader("Y"))
                        Me.Size = New Size(reader("Width"), reader("Height"))
                        Me.WindowState = If(reader("WindowsState").ToString = "Maximized", FormWindowState.Maximized, FormWindowState.Normal)
                    Else
                        Me.StartPosition = FormStartPosition.CenterParent
                    End If
                End Using
            End Using
        End Using
    End Sub

    Private Sub VideoFBF_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        SalvaPosizioneForm()
    End Sub

    Private Sub SalvaPosizioneForm()
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
                cmd.Parameters.AddWithValue("@FormName", Me.Name)
                cmd.Parameters.AddWithValue("@X", x)
                cmd.Parameters.AddWithValue("@Y", y)
                cmd.Parameters.AddWithValue("@Width", w)
                cmd.Parameters.AddWithValue("@Height", h)
                cmd.Parameters.AddWithValue("@WindowsState", stato)
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

            ' Disabilita UI
            btnCaricaVideo.Enabled = False
            lblRevAttiva.Text = "Caricamento in corso..."
            Application.DoEvents()

            ' --- IMPOSTA CURSORE DI ATTESA ---
            Application.UseWaitCursor = True
            Me.Cursor = Cursors.WaitCursor
            Cursor.Current = Cursors.WaitCursor
            Application.DoEvents()
            ' ---------------------------------

            Dim videoPath = OpenFileDialog1.FileName
            Dim nomeVideo = Path.GetFileNameWithoutExtension(videoPath)
            Dim baseDir = Path.Combine(OttieniPercorsoFrames(), nomeVideo)
            Dim revisioneID As Integer
            Dim Approvato As Boolean = False

            ' Ottieni o inserisci video
            Dim videoID = Await Task.Run(Function()
                                             Dim id = OttieniVideoID(nomeVideo)
                                             If id = -1 Then
                                                 id = InserisciVideo(nomeVideo, videoPath)
                                                 revisioneID = 1
                                             Else
                                                 revisioneID = OttieniProssimoRevisioneID()
                                             End If
                                             Return id
                                         End Function)

            ' Crea directory revisione
            Dim revisioneDir = Path.Combine(baseDir, $"Revisione_{revisioneID:D4}")
            Directory.CreateDirectory(baseDir)
            Directory.CreateDirectory(revisioneDir)

            ' Estrai i frame in background
            Dim tempEditor As New VideoEditor(videoPath, revisioneDir)
            Await Task.Run(Sub() tempEditor.ExtractFrames())

            ' Inserisci revisione e permesso
            Await Task.Run(Sub()
                               InserisciRevisione(videoID, revisioneID, SessioneUtente.NomeUtenteCorrente, "Supervisione")
                               InserisciPermessoUtente(revisioneID, SessioneUtente.NomeUtenteCorrente)
                           End Sub)

            MessageBox.Show($"Video caricato, frame estratti e Revisione_{revisioneID:D4} registrata.", "Operazione completata", MessageBoxButtons.OK, MessageBoxIcon.Information)

            lblRevAttiva.Text = revisioneID.ToString()

            ' Parametri revisione
            Dim parametri = New RevisioneParametri(
            videoID,
            revisioneID,
            SessioneUtente.NomeUtenteCorrente,
            $"Revisione {revisioneID}",
            "visualizza",
            DateTime.Now,
            Approvato
        )

            ' Carica editor e aggiorna UI
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
            ' Ripristina cursore e UI
            Application.UseWaitCursor = previousUseWait
            Me.Cursor = previousCursor
            Cursor.Current = previousCursor
            btnCaricaVideo.Enabled = True
            Application.DoEvents()
        End Try
    End Sub


    Private Function OttieniPercorsoFrames() As String
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query As String = "SELECT Valore FROM Sys_Parametri WHERE Descrizione = @DescPar"

            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@DescPar", "PercorsoFrames")
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
            Dim query = "SELECT VideoID FROM Mov_Scene WHERE Titolo = @Titolo"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@Titolo", nomeVideo)
                Dim result = cmd.ExecuteScalar()
                Return If(result IsNot Nothing, CInt(result), -1)
            End Using
        End Using
    End Function

    Private Function OttieniRevisioneZeroID(videoID As Integer) As Integer
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query = "SELECT TOP 1 RevisioneID FROM Mov_Revisione WHERE VideoID = @VideoID AND Note LIKE 'Revisione 0%'"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@VideoID", videoID)
                Dim result = cmd.ExecuteScalar()
                Return If(result IsNot Nothing, CInt(result), -1)
            End Using
        End Using
    End Function

    Private Function InserisciVideo(nomeVideo As String, percorsoFile As String) As Integer
        Using conn As New SqlConnection(ConnString)
            conn.Open()

            ' Verifica se il video esiste già
            Dim checkQuery = "SELECT VideoID FROM Mov_Scene WHERE Titolo = @Titolo"
            Using checkCmd As New SqlCommand(checkQuery, conn)
                checkCmd.Parameters.AddWithValue("@Titolo", nomeVideo)
                Dim result = checkCmd.ExecuteScalar()
                If result IsNot Nothing Then
                    Return CInt(result) ' già presente
                End If
            End Using

            ' Inserisce il video con percorso
            Dim insertQuery = "
                                INSERT INTO Mov_Scene (Titolo, CreatoDa, PercorsoFile, DataCreazione)
                                OUTPUT INSERTED.VideoID
                                VALUES (@Titolo, @CreatoDa, @PercorsoFile, @DataCreazione)"
            Using insertCmd As New SqlCommand(insertQuery, conn)
                insertCmd.Parameters.AddWithValue("@Titolo", nomeVideo)
                insertCmd.Parameters.AddWithValue("@CreatoDa", SessioneUtente.NomeUtenteCorrente)
                insertCmd.Parameters.AddWithValue("@PercorsoFile", percorsoFile)
                insertCmd.Parameters.AddWithValue("@DataCreazione", DateTime.Now)
                Return CInt(insertCmd.ExecuteScalar())
            End Using
        End Using
    End Function

    Private Function RevisioneZeroEsiste(videoID As Integer) As Boolean
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query = "
        SELECT COUNT(*) 
        FROM Mov_Revisione 
        WHERE VideoID = @VideoID AND Note LIKE 'Revisione 0%'"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@VideoID", videoID)
                Return CInt(cmd.ExecuteScalar()) > 0
            End Using
        End Using
    End Function

    Private Function InserisciRevisioneZero(videoID As Integer, nomeVideo As String) As Integer
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query = "
        INSERT INTO Mov_Revisione (RevisioneID, VideoID, AutoreNomeUtente, DataRevisione, Note, Stato)
        VALUES (@RevisioneID, @VideoID, @Autore, @Data, @Note, @Stato)"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@RevisioneID", 0)
                cmd.Parameters.AddWithValue("@VideoID", videoID)
                cmd.Parameters.AddWithValue("@Autore", SessioneUtente.NomeUtenteCorrente)
                cmd.Parameters.AddWithValue("@Data", DateTime.Now)
                cmd.Parameters.AddWithValue("@Note", "Revisione 0 - " & nomeVideo)
                cmd.Parameters.AddWithValue("@Stato", "visualizza")
                cmd.ExecuteNonQuery()
            End Using
        End Using

        Return 0 ' perché la revisione è 0
    End Function

    Private Sub InserisciPermessoUtente(revisioneID As Integer, nomeUtente As String)
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query = "
        INSERT INTO Mov_UtenteRevisione (RevisioneID, NomeUtente, Permesso)
        VALUES (@RevID, @NomeUtente, 'visualizza')"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@RevID", revisioneID)
                cmd.Parameters.AddWithValue("@NomeUtente", nomeUtente)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Private Sub trackFrame_Scroll(sender As Object, e As EventArgs) Handles TrackFrame.Scroll
        picFrame.Image = editor.LoadFrame(TrackFrame.Value)
        AggiornaFrameCorrente(TrackFrame.Value)
    End Sub

    Private Sub btnSuccessivo_Click(sender As Object, e As EventArgs) Handles btnSuccessivo.Click
        If editor Is Nothing OrElse picFrame Is Nothing OrElse TrackFrame Is Nothing OrElse txtNote Is Nothing Then
            MessageBox.Show("Impossibile caricare il frame.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If editor.CurrentIndex < editor.FrameList.Count - 1 Then
            Dim nuovoIndex = editor.CurrentIndex + 1
            picFrame.Image = editor.LoadFrame(nuovoIndex)
            TrackFrame.Value = nuovoIndex

            ' Recupera nota dal database
            txtNote.Text = RecuperaNotaDaDatabase(Int(lblRevAttiva.Text), nuovoIndex)
        End If
    End Sub

    Private Sub btnPrecedente_Click(sender As Object, e As EventArgs) Handles btnPrecedente.Click
        If editor Is Nothing OrElse picFrame Is Nothing OrElse TrackFrame Is Nothing OrElse txtNote Is Nothing Then
            MessageBox.Show("Impossibile caricare il frame.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If editor.CurrentIndex > 0 Then
            Dim nuovoIndex = editor.CurrentIndex - 1
            picFrame.Image = editor.LoadFrame(nuovoIndex)
            TrackFrame.Value = nuovoIndex

            ' Recupera nota dal database
            txtNote.Text = RecuperaNotaDaDatabase(Int(lblRevAttiva.Text), nuovoIndex)
        End If
    End Sub

    Private Function RecuperaNotaDaDatabase(revisioneID As Integer, frameIndex As Integer) As String
        Using conn As New SqlConnection(ConnString)
            Dim cmd As New SqlCommand("
            SELECT TestoNota 
            FROM Mov_FrameNote 
            WHERE RevisioneID = @RevID AND FrameIndex = @FrameIndex", conn)
            cmd.Parameters.AddWithValue("@RevID", revisioneID)
            cmd.Parameters.AddWithValue("@FrameIndex", frameIndex)
            conn.Open()

            Dim nota = cmd.ExecuteScalar()
            Return If(nota IsNot Nothing, nota.ToString(), "")
        End Using
    End Function

    Private Sub picFrame_MouseDown(sender As Object, e As MouseEventArgs) Handles picFrame.MouseDown
        If picFrame.Image Is Nothing Then
            ' Ignora il click se non ci sono frame
            Return
        End If
        isDrawing = True
        lastPoint = e.Location
        editor.SaveState()
    End Sub

    Private Sub picFrame_MouseMove(sender As Object, e As MouseEventArgs) Handles picFrame.MouseMove
        If picFrame.Image Is Nothing Then
            ' Ignora il click se non ci sono frame
            Return
        End If
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
        If picFrame.Image Is Nothing Then
            ' Ignora il click se non ci sono frame
            Return
        End If
        isDrawing = False
    End Sub

    Private Function RevisioneModificabile() As Boolean
        Return Parametri IsNot Nothing AndAlso Parametri.RevisioneID <> 0
    End Function

    Private Sub btnAnnulla_Click(sender As Object, e As EventArgs) Handles btnAnnulla.Click
        If picFrame.Image Is Nothing Then
            ' Ignora il click se non ci sono frame
            Return
        End If

        editor.Undo()
        picFrame.Image = editor.DrawingBitmap

        ' Recupera la nota dal database
        Dim revisioneID = Int(lblRevAttiva.Text)
        Dim frameIndex = editor.CurrentIndex
        txtNote.Text = RecuperaNotaDaDatabase(revisioneID, frameIndex)
    End Sub

    Private Sub SalvaFrame()
        editor.SaveFrame()
    End Sub

    Private Sub btnSalvaVideo_Click(sender As Object, e As EventArgs) Handles btnSalvaVideo.Click
        If picFrame.Image Is Nothing Then
            MessageBox.Show("Caricare prima i Frames", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ' Ignora il click se non ci sono frame
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

        Dim primoIndex = 0
        picFrame.Image = editor.LoadFrame(primoIndex)
        TrackFrame.Value = primoIndex

        ' Recupera la nota dal database
        Dim revisioneID = Int(lblRevAttiva.Text)
        txtNote.Text = RecuperaNotaDaDatabase(revisioneID, primoIndex)
    End Sub

    Private Sub btnUltimoFrame_Click(sender As Object, e As EventArgs) Handles btnUltimoFrame.Click
        If editor Is Nothing OrElse picFrame Is Nothing OrElse TrackFrame Is Nothing OrElse txtNote Is Nothing Then
            MessageBox.Show("Impossibile caricare il frame.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim ultimoIndex = editor.FrameList.Count - 1
        picFrame.Image = editor.LoadFrame(ultimoIndex)
        TrackFrame.Value = ultimoIndex

        ' Recupera la nota dal database
        Dim revisioneID = Int(lblRevAttiva.Text)
        txtNote.Text = RecuperaNotaDaDatabase(revisioneID, ultimoIndex)
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
            btnIndietroVeloce.Text = testoIndietroOriginale ' reset se l'altro era attivo
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
            btnAvantiVeloce.Text = testoAvantiOriginale ' reset se l'altro era attivo
            StartAutoScroll()
            AggiornaFrameCorrente(TrackFrame.Value)
        End If
    End Sub

    Private Async Sub StartAutoScroll()
        While autoScrollActive
            Dim nuovoIndex As Integer = editor.CurrentIndex

            If autoScrollDirection = "forward" AndAlso editor.CurrentIndex < editor.FrameList.Count - 1 Then
                nuovoIndex += 1
                picFrame.Image = editor.LoadFrame(nuovoIndex)

            ElseIf autoScrollDirection = "backward" AndAlso editor.CurrentIndex > 0 Then
                nuovoIndex -= 1
                picFrame.Image = editor.LoadFrame(nuovoIndex)

            Else
                ' Fine raggiunta: interrompi e ripristina testo
                autoScrollActive = False
                If autoScrollDirection = "forward" Then
                    btnAvantiVeloce.Text = testoAvantiOriginale
                ElseIf autoScrollDirection = "backward" Then
                    btnIndietroVeloce.Text = testoIndietroOriginale
                End If
            End If

            editor.CurrentIndex = nuovoIndex
            TrackFrame.Value = nuovoIndex

            ' Recupera nota dal database
            Dim revisioneID = Int(lblRevAttiva.Text)
            txtNote.Text = RecuperaNotaDaDatabase(revisioneID, nuovoIndex)

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

            ' Calcola dimensione del testo multilinea
            Dim textSize = g.MeasureString(txtNote.Text, font, 400) ' larghezza massima
            Dim boxWidth = CInt(textSize.Width) + padding * 2
            Dim boxHeight = CInt(textSize.Height) + padding * 2

            ' Regola la posizione se il box esce dai bordi
            Dim x = notaPosizione.X
            Dim y = notaPosizione.Y

            If x + boxWidth > editor.DrawingBitmap.Width Then
                x = editor.DrawingBitmap.Width - boxWidth - 10
            End If
            If y + boxHeight > editor.DrawingBitmap.Height Then
                y = editor.DrawingBitmap.Height - boxHeight - 10
            End If

            Dim rect = New Rectangle(x, y, boxWidth, boxHeight)

            ' Sfondo semitrasparente
            g.FillRectangle(New SolidBrush(Color.FromArgb(180, Color.Black)), rect)

            ' Testo multilinea con layout automatico
            Dim format As New StringFormat
            format.Alignment = StringAlignment.Near
            format.LineAlignment = StringAlignment.Near
            format.FormatFlags = StringFormatFlags.LineLimit

            g.DrawString(txtNote.Text, font, Brushes.White, rect, format)
        End Using

        picFrame.Image = CType(editor.DrawingBitmap.Clone, Bitmap)
        notaPosizione = Point.Empty
    End Sub

    Private Sub btnSalvaNote_Click(sender As Object, e As EventArgs) Handles btnSalvaNote.Click
        If picFrame.Image Is Nothing Then Return

        Dim frameIndex = editor.CurrentIndex
        Dim nota = txtNote.Text.Trim()
        Dim nomeUtente = NomeUtenteCorrente
        Dim revisioneID = Int(lblRevAttiva.Text)

        ' Se la nota è vuota e c'è una selezione attiva, elimina
        If nota = "" AndAlso lstNoteFrame.SelectedItems.Count > 0 Then
            EliminaNotaSelezionata()
            AggiornaNoteDaDatabase(revisioneID)
            txtNote.Clear()
            lblAutore.Text = ""
            lblDataNota.Text = ""
            Return
        End If

        ' Se c'è una nota da salvare
        If nota <> "" Then
            Using conn As New SqlConnection(ConnString)
                conn.Open()

                Dim cmd As New SqlCommand("
                MERGE Mov_FrameNote AS target
                USING (SELECT @RevID AS RevisioneID, @FrameIndex AS FrameIndex) AS source
                ON target.RevisioneID = source.RevisioneID AND target.FrameIndex = source.FrameIndex
                WHEN MATCHED THEN 
                    UPDATE SET TestoNota = @TestoNota, NomeUtente = @NomeUtente, DataNota = GETDATE()
                WHEN NOT MATCHED THEN
                    INSERT (RevisioneID, FrameIndex, NomeUtente, TestoNota, DataNota)
                    VALUES (@RevID, @FrameIndex, @NomeUtente, @TestoNota, GETDATE());", conn)

                cmd.Parameters.AddWithValue("@RevID", revisioneID)
                cmd.Parameters.AddWithValue("@FrameIndex", frameIndex)
                cmd.Parameters.AddWithValue("@NomeUtente", nomeUtente)
                cmd.Parameters.AddWithValue("@TestoNota", nota)
                cmd.ExecuteNonQuery()
            End Using
        Else
            MessageBox.Show("Inserisci una nota prima di salvare.")
            Return
        End If

        ' Aggiorna lista e segnalini
        AggiornaNoteDaDatabase(revisioneID)

        ' Salva modifiche grafiche sul frame
        SalvaFrame()
    End Sub

    Private Sub btnCaricaRevisione_Click(sender As Object, e As EventArgs) Handles btnCaricaRevisione.Click
        If TypeOf Me.MdiParent Is GesPu25 Then
            Dim mainForm As GesPu25 = CType(Me.MdiParent, GesPu25)
            mainForm.ApriModulo2ConPermessi("SceltaVideo", New SceltaVideo(Me)) ' ← passa Me
        Else
            MessageBox.Show("Form principale non disponibile.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub CaricaListaNote()
        lstNoteFrame.Items.Clear()
        Dim revisioneID = Int(lblRevAttiva.Text)

        Using conn As New SqlConnection(ConnString)
            Dim cmd As New SqlCommand("
            SELECT FrameIndex, TestoNota, NomeUtente, DataNota
            FROM Mov_FrameNote
            WHERE RevisioneID = @RevID
            ORDER BY FrameIndex", conn)
            cmd.Parameters.AddWithValue("@RevID", revisioneID)
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
                    item.Tag = frameIndex
                    item.ToolTipText = $"Autore: {autore}{Environment.NewLine}Data: {data:dd/MM/yyyy HH:mm}"
                    lstNoteFrame.Items.Add(item)
                End While
            End Using
        End Using
    End Sub

    Private Sub lstNote_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Delete Then
            EliminaNotaSelezionata
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
                cmd.Parameters.AddWithValue("@RevID", revisioneID)
                cmd.Parameters.AddWithValue("@FrameIndex", frameIndex)
                cmd.Parameters.AddWithValue("@TestoNota", testoNota)
                cmd.ExecuteNonQuery()
            End Using

            AggiornaNoteDaDatabase(revisioneID)
            txtNote.Clear()
            lblAutore.Text = ""
            lblDataNota.Text = ""
        End If
    End Sub

    Private Sub lstNoteFrame_SelectedIndexChanged(sender As Object, e As EventArgs)
        If lstNoteFrame.SelectedItems.Count > 0 Then
            Dim info = lstNoteFrame.SelectedItems(0).Tag
            Dim frameIndex = info.FrameIndex

            TrackFrame.Value = frameIndex
            AggiornaFrameCorrente(frameIndex)
        End If
    End Sub

    Public Sub CaricaRevisione(videoID As Integer, revisioneID As Integer)
        ' Recupera il percorso del video dal database
        Dim videoPath As String = ""
        Dim nomeVideo = OttieniNomeVideo(videoID)
        Dim basePath = OttieniPercorsoFrames()
        Dim frameDir = Path.Combine(basePath, nomeVideo, $"Revisione_{revisioneID:0000}")

        Using conn As New SqlConnection(ConnString)
            Dim query As String = "SELECT PercorsoFile FROM Mov_Scene WHERE VideoID = @VideoID"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@VideoID", videoID)
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

        ' Inizializza l'editor solo se i frame esistono
        If Not Directory.Exists(frameDir) OrElse Directory.GetFiles(frameDir).Length = 0 Then
            MessageBox.Show("Frame non trovati per la revisione selezionata.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        editor = New VideoEditor(videoPath, frameDir)

        ' Carica Annotazioni (solo per disegno segnalini)
        AggiornaNoteDaDatabase(revisioneID)

        Dim overlay = Me.Controls.Find("OverlayNotePanel", True).FirstOrDefault()
        overlay?.Invalidate()

        ' Carica il primo frame
        TrackFrame.Minimum = 0
        TrackFrame.Maximum = editor.FrameList.Count - 1
        TrackFrame.Value = 0
        picFrame.Image = editor.LoadFrame(0)

        ' Recupera la nota dal database
        txtNote.Text = RecuperaNotaDaDatabase(revisioneID, 0)
    End Sub

    Private Function OttieniNomeVideo(videoID As Integer) As String
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query = "SELECT Titolo FROM Mov_Scene WHERE VideoID = @ID"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@ID", videoID)
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
        Dim larghezza = TrackFrame.Width - 28

        For Each index In FrameConNote
            If index < TrackFrame.Minimum OrElse index > TrackFrame.Maximum Then Continue For

            Dim percentuale = (index - TrackFrame.Minimum) / totale
            Dim x = CInt(percentuale * larghezza)

            ' Disegna un pallino arancione sopra il punto
            'e.Graphics.FillEllipse(Brushes.Red, x + 11, 0, 5, 9)
            e.Graphics.FillRectangle(Brushes.Red, x + 11, 0, 5, 10)
        Next
    End Sub

    Public Sub AggiornaFrameCorrente(index As Integer)
        If editor Is Nothing Then Exit Sub
        If index < 0 OrElse index >= editor.FrameList.Count Then Exit Sub

        picFrame.Image = editor.LoadFrame(index)

        Dim revisioneID = Int(lblRevAttiva.Text)

        Using conn As New SqlConnection(ConnString)
            Dim cmd As New SqlCommand("
            SELECT TestoNota, NomeUtente, DataNota 
            FROM Mov_FrameNote 
            WHERE RevisioneID = @RevID AND FrameIndex = @FrameIndex", conn)
            cmd.Parameters.AddWithValue("@RevID", revisioneID)
            cmd.Parameters.AddWithValue("@FrameIndex", index)
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
                cmd.Parameters.AddWithValue("@RevID", revisioneID)
                conn.Open()

                Using reader As SqlDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim frameIndex = Convert.ToInt32(reader("FrameIndex"))
                        Dim testo = reader("TestoNota").ToString()
                        Dim autore = reader("NomeUtente").ToString()
                        Dim data = Convert.ToDateTime(reader("DataNota"))

                        ' Aggiungi frame annotato
                        If Not FrameConNote.Contains(frameIndex) Then
                            FrameConNote.Add(frameIndex)
                        End If

                        ' Crea oggetto informativo
                        Dim info As New NotaFrameInfo With {
                        .FrameIndex = frameIndex,
                        .TestoNota = testo,
                        .Autore = autore,
                        .DataNota = data
                    }

                        ' Crea voce per la lista
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

        ' Ricrea o invalida il pannello dei segnalini
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

            ' Carica revisione e imposta modalità
            CaricaRevisione(videoID, revisioneID)
            AggiornaNoteDaDatabase(revisioneID)
        End If
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

    Private Sub InserisciRevisione(videoId, RevisioneID, nomeUtente, stato)

        Dim Titolo = OttieniNomeVideo(videoId)
        Dim dataRevisione = DateTime.Now

        ' Genera nome revisione
        'Dim numero = OttieniProssimoRevisioneID()
        'Dim nuovaRevisioneID = numero

        Dim NumRetake = CalcolaRetake(videoId)

        Dim nomeRevisione = $"Revisione {RevisioneID} - {dataRevisione:dd/MM/yyyy}"
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query = "INSERT INTO Mov_Revisione (RevisioneID, VideoID, Autore, DataRevisione, NumRetake, Note, Stato)
                        VALUES (@RevisioneID, @VideoID, @Autore, @Data, @NumRetake, @Note, @Stato)"

            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@RevisioneID", RevisioneID)
                cmd.Parameters.AddWithValue("@VideoID", videoId)
                cmd.Parameters.AddWithValue("@Autore", nomeUtente)
                cmd.Parameters.AddWithValue("@Data", dataRevisione)
                cmd.Parameters.AddWithValue("@Note", nomeRevisione)
                cmd.Parameters.AddWithValue("@NumRetake", NumRetake)
                cmd.Parameters.AddWithValue("@Stato", stato)
                cmd.ExecuteNonQuery()
            End Using
        End Using

    End Sub

    Public Function CalcolaRetake(videoID As Integer) As Integer
        Using conn As New SqlConnection(ConnString)
            conn.Open()

            Dim query As String = "SELECT COUNT(*) FROM Mov_Revisione WHERE VideoID = @VideoID"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@VideoID", videoID)
                Return CInt(cmd.ExecuteScalar())
            End Using
        End Using
    End Function

    Private Function OttieniProssimoRevisioneID() As Integer
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query = "SELECT ISNULL(MAX(RevisioneID), 0) + 1 FROM Mov_Revisione"
            Using cmd As New SqlCommand(query, conn)
                Return CInt(cmd.ExecuteScalar())
            End Using
        End Using
    End Function

    Public Function OttieniNumeroRevisione() As Integer
        Dim numero As Integer = 0

        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query As String = "
        SELECT ISNULL(MAX(RevisioneID), -1)
        FROM Mov_Revisione"

            Using cmd As New SqlCommand(query, conn)
                numero = CInt(cmd.ExecuteScalar())
            End Using
        End Using

        ' Se non ci sono revisioni, restituisce -1
        Return numero
    End Function

    Public Function IsRevisioneModificabile(revisioneID As Integer) As Boolean
        Dim modificabile As Boolean = False

        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query As String = "
        SELECT Stato 
        FROM Mov_Revisione 
        WHERE RevisioneID = @ID"

            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@ID", revisioneID)
                Dim stato = CStr(cmd.ExecuteScalar())
                modificabile = (stato = "bozza" OrElse stato = "modifica")
            End Using
        End Using

        Return modificabile
    End Function

    Private Sub DisabilitaControlliModifica()
        btnAnnulla.Enabled = False
        btnColorePennino.Enabled = False
        btnSalvaNote.Enabled = False
        btnSalvaVideo.Enabled = False
        numSpessorePennino.Enabled = False
        txtNote.Enabled = False
    End Sub

    Private Sub AbilitaControlliModifica()
        btnAnnulla.Enabled = True
        btnColorePennino.Enabled = True
        btnSalvaNote.Enabled = True
        btnSalvaVideo.Enabled = True
        numSpessorePennino.Enabled = True
        txtNote.Enabled = True
    End Sub

    Private Function EsistonoRevisioniAttive(videoID As Integer) As Boolean
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query As String = "
        SELECT COUNT(*) 
        FROM Mov_Revisione 
        WHERE VideoID = @VideoID"

            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@VideoID", videoID)
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

        lblRevAttiva.Text = Parametri.RevisioneID
    End Sub

    Private Function VerificaCreazioneRevisione(videoID As Integer, revisioneCorrente As Integer) As Boolean
        Using conn As New SqlConnection(ConnString)
            conn.Open()

            ' Recupera tutte le revisioni esistenti per il video
            Dim revisioniEsistenti As New List(Of Integer)
            Dim query As String = "SELECT RevisioneID FROM Mov_Revisione WHERE VideoID = @VideoID ORDER BY RevisioneID ASC"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@VideoID", videoID)
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

End Class

Public Class NotaFrameInfo
    Public Property FrameIndex As Integer
    Public Property TestoNota As String
    Public Property Autore As String
    Public Property DataNota As DateTime
End Class
