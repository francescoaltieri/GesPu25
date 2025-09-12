Imports System.IO
Imports System.Text.RegularExpressions
Imports Microsoft.Data.SqlClient

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

    Public Property Parametri As RevisioneParametri

    Public Sub New(parametri As RevisioneParametri)
        InitializeComponent()
        Me.Parametri = parametri
    End Sub

    Public Sub New()
        InitializeComponent()
        Me.Parametri = Nothing ' oppure crea un RevisioneParametri vuoto se vuoi
    End Sub

    Private Sub btnCaricaVideo_Click(sender As Object, e As EventArgs) Handles btnCaricaVideo.Click


        OpenFileDialog1.Filter = "Video Files|*.mp4;*.mov"
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim videoPath = OpenFileDialog1.FileName
            Dim nomeVideo = Path.GetFileNameWithoutExtension(videoPath)
            Dim baseDir = Path.Combine("C:\VideoEditor\Frames", nomeVideo)
            Dim revisioneZeroDir = Path.Combine(baseDir, "Revisione_000")
            Dim revisioneID As Integer
            Dim videoID = OttieniVideoID(nomeVideo)

            If videoID > 0 Then
                MessageBox.Show("Il video è già presente nel database. Verranno aggiornate le cartelle relative.", "Video già registrato", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

            ' Crea struttura se mancante
            If Not Directory.Exists(revisioneZeroDir) OrElse Directory.GetFiles(revisioneZeroDir).Length = 0 Then
                If Not Directory.Exists(baseDir) Then Directory.CreateDirectory(baseDir)
                If Not Directory.Exists(revisioneZeroDir) Then Directory.CreateDirectory(revisioneZeroDir)

                ' Estrai i frame
                Dim tempEditor = New VideoEditor(videoPath, revisioneZeroDir)
                tempEditor.ExtractFrames()

                ' Inserisci revisione 0 nel database
                videoID = InserisciVideo(nomeVideo, videoPath)

                If Not RevisioneZeroEsiste(videoID) Then
                    revisioneID = InserisciRevisioneZero(videoID, nomeVideo)
                    InserisciPermessoUtente(revisioneID, SessioneUtente.NomeUtenteCorrente)
                Else
                    revisioneID = OttieniRevisioneZeroID(videoID)
                End If

                MessageBox.Show("Video caricato, frame estratti e Revisione_000 registrata.", "Operazione completata", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                videoID = OttieniVideoID(nomeVideo)
                revisioneID = OttieniRevisioneZeroID(videoID)
                MessageBox.Show("Il video è già stato caricato. Frame e revisione 0 già presenti.", "Informazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

            ' Crea i parametri corretti ora che hai gli ID
            Dim parametri = New RevisioneParametri(
            videoID,
            revisioneID,
            "visualizza",
            SessioneUtente.NomeUtenteCorrente,
            "Revisione 0",
            "visualizza",
            DateTime.Now
        )

            ' Carica editor
            editor = New VideoEditor(videoPath, revisioneZeroDir)
            picFrame.Image = editor.LoadFrame(0)
            TrackFrame.Minimum = 0
            TrackFrame.Maximum = editor.FrameList.Count - 1
            TrackFrame.Value = 0

            Me.Parametri = parametri
            Me.AggiornaRevisioneAttiva()
        End If
    End Sub

    Private Function OttieniVideoID(nomeVideo As String) As Integer
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query = "SELECT VideoID FROM Mov_Video WHERE Titolo = @Titolo"
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
            Dim checkQuery = "SELECT VideoID FROM Mov_Video WHERE Titolo = @Titolo"
            Using checkCmd As New SqlCommand(checkQuery, conn)
                checkCmd.Parameters.AddWithValue("@Titolo", nomeVideo)
                Dim result = checkCmd.ExecuteScalar()
                If result IsNot Nothing Then
                    Return CInt(result) ' già presente
                End If
            End Using

            ' Inserisce il video con percorso
            Dim insertQuery = "
        INSERT INTO Mov_Video (Titolo, CreatoDa, PercorsoFile)
        OUTPUT INSERTED.VideoID
        VALUES (@Titolo, @CreatoDa, @PercorsoFile)"
            Using insertCmd As New SqlCommand(insertQuery, conn)
                insertCmd.Parameters.AddWithValue("@Titolo", nomeVideo)
                insertCmd.Parameters.AddWithValue("@CreatoDa", SessioneUtente.NomeUtenteCorrente)
                insertCmd.Parameters.AddWithValue("@PercorsoFile", percorsoFile)
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
            picFrame.Image = editor.LoadFrame(editor.CurrentIndex + 1)
            TrackFrame.Value = editor.CurrentIndex
            txtNote.Text = editor.GetNotaPerFrame(editor.CurrentIndex)
        End If
    End Sub

    Private Sub btnPrecedente_Click(sender As Object, e As EventArgs) Handles btnPrecedente.Click
        If editor Is Nothing OrElse picFrame Is Nothing OrElse TrackFrame Is Nothing OrElse txtNote Is Nothing Then
            MessageBox.Show("Impossibile caricare il frame.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If editor.CurrentIndex > 0 Then
            picFrame.Image = editor.LoadFrame(editor.CurrentIndex - 1)
            TrackFrame.Value = editor.CurrentIndex
            txtNote.Text = editor.GetNotaPerFrame(editor.CurrentIndex)
        End If
    End Sub

    Private Sub picFrame_MouseDown(sender As Object, e As MouseEventArgs) Handles picFrame.MouseDown
        If Not RevisioneModificabile() Then
            MessageBox.Show("Non è possibile modificare i frame della Revisione 0.", "Operazione non consentita", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        isDrawing = True
        lastPoint = e.Location
        editor.SaveState()
    End Sub

    Private Sub picFrame_MouseMove(sender As Object, e As MouseEventArgs) Handles picFrame.MouseMove
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
        isDrawing = False
    End Sub

    Private Function RevisioneModificabile() As Boolean
        Return Parametri IsNot Nothing AndAlso Parametri.RevisioneID <> 0
    End Function


    Private Sub btnAnnulla_Click(sender As Object, e As EventArgs) Handles btnAnnulla.Click
        If Not RevisioneModificabile() Then
            MessageBox.Show("La Revisione 0 non può essere modificata.", "Operazione non consentita", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        editor.Undo()
        picFrame.Image = editor.DrawingBitmap
        txtNote.Text = editor.GetNotaPerFrame(editor.CurrentIndex)
    End Sub

    Private Sub btnSalvaFrame_Click(sender As Object, e As EventArgs) Handles btnSalvaFrame.Click
        If Not RevisioneModificabile() Then
            MessageBox.Show("La Revisione 0 non può essere modificata.", "Operazione non consentita", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        editor.SaveFrame()
    End Sub

    Private Sub btnSalvaVideo_Click(sender As Object, e As EventArgs) Handles btnSalvaVideo.Click
        Dim outputPath = "C:\VideoEditor\output.mp4"
        editor.RebuildVideo(outputPath)
        MessageBox.Show("Video salvato in: " & outputPath)
    End Sub

    Private Sub btnPrimoFrame_Click(sender As Object, e As EventArgs) Handles btnPrimoFrame.Click
        If editor Is Nothing OrElse picFrame Is Nothing OrElse TrackFrame Is Nothing OrElse txtNote Is Nothing Then
            MessageBox.Show("Impossibile caricare il frame.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        picFrame.Image = editor.LoadFrame(0)
        TrackFrame.Value = editor.CurrentIndex
        txtNote.Text = editor.GetNotaPerFrame(editor.CurrentIndex)
    End Sub

    Private Sub btnUltimoFrame_Click(sender As Object, e As EventArgs) Handles btnUltimoFrame.Click
        If editor Is Nothing OrElse picFrame Is Nothing OrElse TrackFrame Is Nothing OrElse txtNote Is Nothing Then
            MessageBox.Show("Impossibile caricare il frame.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        picFrame.Image = editor.LoadFrame(editor.FrameList.Count - 1)
        TrackFrame.Value = editor.CurrentIndex
        txtNote.Text = editor.GetNotaPerFrame(editor.CurrentIndex)
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
            If autoScrollDirection = "forward" AndAlso editor.CurrentIndex < editor.FrameList.Count - 1 Then
                picFrame.Image = editor.LoadFrame(editor.CurrentIndex + 1)
            ElseIf autoScrollDirection = "backward" AndAlso editor.CurrentIndex > 0 Then
                picFrame.Image = editor.LoadFrame(editor.CurrentIndex - 1)
            Else
                ' Fine raggiunta: interrompi e ripristina testo
                autoScrollActive = False
                If autoScrollDirection = "forward" Then
                    btnAvantiVeloce.Text = testoAvantiOriginale
                ElseIf autoScrollDirection = "backward" Then
                    btnIndietroVeloce.Text = testoIndietroOriginale
                End If
            End If

            txtNote.Text = editor.GetNotaPerFrame(editor.CurrentIndex)
            TrackFrame.Value = editor.CurrentIndex
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
        If Not RevisioneModificabile() Then
            MessageBox.Show("Non è possibile modificare i frame della Revisione 0.", "Operazione non consentita", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        notaPosizione = e.Location
    End Sub

    Private Sub btnColorePennino_Click(sender As Object, e As EventArgs) Handles btnColorePennino.Click
        If colorDialogPennino.ShowDialog() = DialogResult.OK Then
            colorePennino = colorDialogPennino.Color
            btnColorePennino.BackColor = colorePennino
        End If
    End Sub

    Private Sub numSpessorePennino_ValueChanged(sender As Object, e As EventArgs) Handles numSpessorePennino.ValueChanged
        spessorePennino = CInt(numSpessorePennino.Value)
    End Sub

    Private Sub btnAggiungiNote_Click_1(sender As Object, e As EventArgs) Handles btnAggiungiNote.Click
        If Not RevisioneModificabile() Then
            MessageBox.Show("La Revisione 0 non può essere modificata.", "Operazione non consentita", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

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
            Dim format As New StringFormat()
            format.Alignment = StringAlignment.Near
            format.LineAlignment = StringAlignment.Near
            format.FormatFlags = StringFormatFlags.LineLimit


            g.DrawString(txtNote.Text, font, Brushes.White, rect, format)
        End Using

        picFrame.Image = CType(editor.DrawingBitmap.Clone, Bitmap)
        notaPosizione = Point.Empty
    End Sub

    Private Sub btnSalvaNote_Click(sender As Object, e As EventArgs) Handles btnSalvaNote.Click
        If Not RevisioneModificabile() Then
            MessageBox.Show("La Revisione 0 non può essere modificata.", "Operazione non consentita", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim frameIndex = editor.CurrentIndex
        Dim nota = txtNote.Text.Trim
        Dim nomeUtente = NomeUtenteCorrente
        Dim revisioneID = CType(Me.Tag, Object).RevisioneID

        If editor IsNot Nothing AndAlso txtNote.Text.Trim <> "" Then

            Using conn As New SqlConnection(ConnString)
                conn.Open()

                ' Elimina nota precedente se esiste
                Dim deleteCmd As New SqlCommand("
                DELETE FROM Mov_FrameNote 
                WHERE RevisioneID = @RevID AND FrameIndex = @FrameIndex", conn)
                deleteCmd.Parameters.AddWithValue("@RevID", revisioneID)
                deleteCmd.Parameters.AddWithValue("@FrameIndex", frameIndex)
                deleteCmd.ExecuteNonQuery()

                ' Inserisci nuova nota
                Dim insertCmd As New SqlCommand("
                INSERT INTO Mov_FrameNote (RevisioneID, FrameIndex, NomeUtente, TestoNota)
                VALUES (@RevID, @FrameIndex, @NomeUtente, @TestoNota)", conn)
                insertCmd.Parameters.AddWithValue("@RevID", revisioneID)
                insertCmd.Parameters.AddWithValue("@FrameIndex", frameIndex)
                insertCmd.Parameters.AddWithValue("@NomeUtente", nomeUtente)
                insertCmd.Parameters.AddWithValue("@TestoNota", nota)
                insertCmd.ExecuteNonQuery()
            End Using
        Else
            MessageBox.Show("Inserisci una nota prima di salvare.")
        End If

        RicaricaNoteDaDatabase(revisioneID)
        ' Aggiorna la lista
        CaricaListaNote()

    End Sub

    Private Sub btnCaricaRevisione_Click(sender As Object, e As EventArgs) Handles btnCaricaRevisione.Click
        If TypeOf Me.MdiParent Is GesPu25 Then
            Dim mainForm As GesPu25 = CType(Me.MdiParent, GesPu25)
            mainForm.ApriModulo2ConPermessi("SceltaVideo", New SceltaVideo(Me)) ' ← passa Me
        Else
            MessageBox.Show("Form principale non disponibile.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub VideoFBF_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Parametri IsNot Nothing Then
            lblRevAttiva.Text = $"Revisione attiva: {Parametri.RevisioneID}"
            ' Altrimenti, carica la revisione 
            If Parametri.RevisioneID = 0 Then
                DisabilitaControlliModifica()
                MessageBox.Show("Questa revisione non è modificabile.", "Modalità sola lettura", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Else
            lblRevAttiva.Text = "Nessuna revisione attiva"
        End If

    End Sub


    Private Sub CaricaListaNote()
        lstNoteFrame.Items.Clear()

        For Each index In editor.GetFrameAnnotati()
            Dim testo = editor.GetNotaPerFrame(index)
            Dim autore = editor.GetAutorePerFrame(index)
            Dim data = editor.GetDataNotaPerFrame(index)

            Dim anteprima = If(testo.Length > 30, testo.Substring(0, 30) & "...", testo)
            Dim voce = $"Frame {index + 1}: {anteprima}"

            ' Aggiungi voce con tooltip
            Dim item As New ListViewItem(voce)
            item.Tag = index
            item.ToolTipText = $"Autore: {autore}{Environment.NewLine}Data: {data:dd/MM/yyyy HH:mm}"
            lstNoteFrame.Items.Add(voce)
        Next
    End Sub

    Private Sub lstNoteFrame_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstNoteFrame.SelectedIndexChanged
        If lstNoteFrame.SelectedIndex >= 0 Then
            Dim frameIndex = editor.GetFrameAnnotati()(lstNoteFrame.SelectedIndex)
            TrackFrame.Value = frameIndex
            AggiornaFrameCorrente(frameIndex)
        End If
    End Sub


    Public Sub CaricaRevisione(videoID As Integer, revisioneID As Integer)
        ' Recupera il percorso del video dal database
        Dim videoPath As String = ""
        Dim nomeVideo = OttieniNomeVideo(videoID)
        Dim frameDir As String = $"C:\VideoEditor\Frames\{nomeVideo}\Revisione_{revisioneID:000}"

        Using conn As New SqlConnection(ConnString)
            Dim query As String = "SELECT PercorsoFile FROM Mov_Video WHERE VideoID = @VideoID"
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

        ' Carica Annotazioni
        CaricaNoteDaDatabase(revisioneID)

        ' Carica il primo frame
        TrackFrame.Minimum = 0
        TrackFrame.Maximum = editor.FrameList.Count - 1
        TrackFrame.Value = 0
        picFrame.Image = editor.LoadFrame(0)
        txtNote.Text = editor.GetNotaPerFrame(0)

    End Sub

    Private Function OttieniNomeVideo(videoID As Integer) As String
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query = "SELECT Titolo FROM Mov_Video WHERE VideoID = @ID"
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

    Private Sub CaricaNoteDaDatabase(revisioneID As Integer)
        Using conn As New SqlConnection(ConnString)
            Dim query As String = "
            SELECT FrameIndex, TestoNota, NomeUtente, DataNota 
            FROM Mov_FrameNote 
            WHERE RevisioneID = @RevID"

            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@RevID", revisioneID)
                conn.Open()
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim index = Convert.ToInt32(reader("FrameIndex"))
                        Dim nota = reader("TestoNota").ToString()
                        Dim autore = reader("NomeUtente").ToString()
                        Dim data = Convert.ToDateTime(reader("DataNota"))

                        editor.SalvaNotaCompleta(index, nota, autore, data)
                    End While
                End Using
            End Using
        End Using
    End Sub

    Private Sub AggiornaFrameCorrente(index As Integer)
        If editor Is Nothing Then Exit Sub
        If index < 0 OrElse index >= editor.FrameList.Count Then Exit Sub

        picFrame.Image = editor.LoadFrame(index)
        txtNote.Text = editor.GetNotaPerFrame(index)
        lblAutore.Text = $"Autore: {editor.GetAutorePerFrame(index)}"
        lblDataNota.Text = $"Data: {editor.GetDataNotaPerFrame(index):dd/MM/yyyy HH:mm}"
    End Sub

    Private Sub RicaricaNoteDaDatabase(revisioneID As Integer)
        Using conn As New SqlConnection(ConnString)
            Dim query As String = "
            SELECT FrameIndex, TestoNota, NomeUtente, DataNota 
            FROM Mov_FrameNote 
            WHERE RevisioneID = @RevID"

            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@RevID", revisioneID)
                conn.Open()
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim index = Convert.ToInt32(reader("FrameIndex"))
                        Dim nota = reader("TestoNota").ToString()
                        Dim autore = reader("NomeUtente").ToString()
                        Dim data = Convert.ToDateTime(reader("DataNota"))

                        editor.SalvaNotaCompleta(index, nota, autore, data)
                    End While
                End Using
            End Using
        End Using
    End Sub

    Private Sub VideoFBF_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        If Me.Tag IsNot Nothing Then
            Dim param = CType(Me.Tag, Object)
            Dim videoID = param.VideoID
            Dim revisioneID = param.RevisioneID
            Dim permesso = param.Permesso

            ' Carica revisione e imposta modalità
            CaricaRevisione(videoID, revisioneID)
            CaricaListaNote()
        End If
    End Sub

    Private Sub btnNuovaRevisione_Click(sender As Object, e As EventArgs) Handles btnNuovaRevisione.Click
        If Not VerificaCreazioneRevisione(Parametri.VideoID, Parametri.RevisioneID) Then
            MDIMessageBox.Show("Impossibile creare la Revisione sarebbe incoerente rispetto alla catena delle revisioni successive", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If Parametri Is Nothing Then
            MDIMessageBox.Show("Parametri revisione non disponibili.", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Dim risposta = MDIMessageBox.Show("Vuoi creare una nuova revisione ?", Me.MdiParent, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If risposta <> vbYes Then
            Return
        End If

        Dim videoID = Parametri.VideoID
        Dim revisioneCorrenteID = Parametri.RevisioneID
        Dim Titolo = OttieniNomeVideo(videoID)
        Dim nomeUtente = Parametri.NomeUtente
        Dim dataRevisione = DateTime.Now
        Dim stato = ""

        If Not EsistonoRevisioniAttive(videoID) Then
            MDIMessageBox.Show("Non ci sono revisioni attive nella tabella Mov_Revisioni.", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' Genera nome revisione
        Dim numero = OttieniNumeroRevisione(videoID)
        Dim nuovaRevisioneID = numero + 1

        Dim nomeRevisione = $"Revisione {nuovaRevisioneID} - {dataRevisione:dd/MM/yyyy}"
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query = "
                        INSERT INTO Mov_Revisione (RevisioneID, VideoID, AutoreNomeUtente, DataRevisione, Note, Stato)
                        VALUES (@RevisioneID, @VideoID, @Autore, @Data, @Note, @Stato)"

            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@RevisioneID", nuovaRevisioneID)
                cmd.Parameters.AddWithValue("@VideoID", videoID)
                cmd.Parameters.AddWithValue("@Autore", nomeUtente)
                cmd.Parameters.AddWithValue("@Data", dataRevisione)
                cmd.Parameters.AddWithValue("@Note", nomeRevisione)
                cmd.Parameters.AddWithValue("@Stato", stato)
                cmd.ExecuteNonQuery()
            End Using
        End Using

        ' Assegna permesso di modifica
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim cmd As New SqlCommand("
        INSERT INTO Mov_UtenteRevisione (RevisioneID, NomeUtente, Permesso)
        VALUES (@RevID, @NomeUtente, 'modifica')", conn)
            cmd.Parameters.AddWithValue("@RevID", nuovaRevisioneID)
            cmd.Parameters.AddWithValue("@NomeUtente", nomeUtente)
            cmd.ExecuteNonQuery()
        End Using

        ' Crea cartella nuova revisione
        Dim baseDir = $"C:\VideoEditor\Frames\{Titolo}"
        Dim nuovaFrameDir = Path.Combine(baseDir, $"Revisione_{nuovaRevisioneID:000}")
        If Not Directory.Exists(nuovaFrameDir) Then Directory.CreateDirectory(nuovaFrameDir)

        ' Copia frame dalla revisione corrente
        Dim origineFrameDir = Path.Combine(baseDir, $"Revisione_{revisioneCorrenteID:000}")
        If Directory.Exists(origineFrameDir) Then
            For Each filePath In Directory.GetFiles(origineFrameDir)
                Dim fileName = Path.GetFileName(filePath)
                Dim destinazione = Path.Combine(nuovaFrameDir, fileName)
                File.Copy(filePath, destinazione, overwrite:=True)
            Next
        End If

        ' Crea parametri e apri nuova revisione
        'Dim nuoviParametri = New RevisioneParametri(videoID, nuovaRevisioneID, "modifica", nomeUtente, nomeRevisione, stato, dataRevisione)
        'Dim videoForm As New VideoFBF(nuoviParametri)
        'videoForm.MdiParent = Me.MdiParent
        'videoForm.Show()
    End Sub

    Private Function OttieniProssimoRevisioneID() As Integer
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query = "SELECT ISNULL(MAX(RevisioneID), 0) + 1 FROM Mov_Revisione"
            Using cmd As New SqlCommand(query, conn)
                Return CInt(cmd.ExecuteScalar())
            End Using
        End Using
    End Function

    Public Function OttieniNumeroRevisione(videoID As Integer) As Integer
        Dim numero As Integer = 0

        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query As String = "
            SELECT ISNULL(MAX(RevisioneID), -1)
            FROM Mov_Revisione 
            WHERE VideoID = @VideoID"

            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@VideoID", videoID)
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
        btnSalvaFrame.Enabled = False
        btnAggiungiNote.Enabled = False
        btnAnnulla.Enabled = False
        btnColorePennino.Enabled = False
        btnSalvaFrame.Enabled = False
        btnSalvaNote.Enabled = False
        btnSalvaVideo.Enabled = False
        numSpessorePennino.Enabled = False
        txtNote.Enabled = False
    End Sub

    Private Sub AbilitaControlliModifica()
        btnSalvaFrame.Enabled = True
        btnAggiungiNote.Enabled = True
        btnAnnulla.Enabled = True
        btnColorePennino.Enabled = True
        btnSalvaFrame.Enabled = True
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
            lblRevAttiva.ForeColor = Color.DarkRed
            Return
        End If

        lblRevAttiva.Text = $"Revisione {Parametri.RevisioneID}"
        lblRevAttiva.ForeColor = If(Parametri.Permesso.ToLower() = "modifica", Color.Green, Color.Gray)
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
