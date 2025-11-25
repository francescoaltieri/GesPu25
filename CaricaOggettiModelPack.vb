Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Data.SqlClient

Public Class CaricaOggettiModelPack
    Inherits Form


    ' =======================
    ' UI (mantengo i nomi del tuo file)
    ' =======================
    Private WithEvents cmbModelPack As New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList, .Name = "cmbModelPack"}
    Private WithEvents cmbTipoOggetto As New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList, .Name = "cmbTipoOggetto"}
    Private WithEvents btnCaricaGriglia As New Button() With {.Text = "Carica Griglia", .Name = "btnCaricaGriglia"}
    Private WithEvents btnAnnulla As New Button() With {.Text = "Annulla caricamento", .Enabled = False, .Name = "btnAnnulla"}
    Private WithEvents flpGrid As New FlowLayoutPanel() With {.AutoScroll = True, .WrapContents = True}
    Private lblStatus As New Label() With {.AutoSize = True}
    Private prgLoad As New ProgressBar() With {.Visible = False}

    ' BtnSalva e BtnChiudi verranno creati in runtime se non presenti
    Private BtnSalva As Button
    Private BtnChiudi As Button

    ' =======================
    ' Stato
    ' =======================
    Private _currentModelPackId As String
    Private _currentTipoId As String
    Private _items As New List(Of OggettoItem)
    Private _selectedKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    Private _cts As CancellationTokenSource

    ' Flags per stato UI e chiusura
    Private _isLoading As Boolean = False
    Private _isClosing As Boolean = False

    ' =======================
    ' Layout
    ' =======================
    Private Const Columns As Integer = 4
    Private Const CellWidth As Integer = 240
    Private Const CellHeight As Integer = 240

    ' Placeholder
    Private ReadOnly _placeholderImage As Bitmap = CreatePlaceholderBitmap(CellWidth - 20, CellHeight - 80)

    ' Pannelli di layout
    Private panelButtons As Panel
    Private statusPanel As Panel
    Private tlpTop As TableLayoutPanel

    Private _initialModelPackId As String

    ' =======================
    ' Costruttore
    ' =======================
    Public Sub New()
        Me.Text = "Carica Oggetti Model Pack"
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Width = 1100
        Me.Height = 750
        Me.BackColor = Color.White

        EnableDoubleBuffering(Me)
        EnableDoubleBuffering(flpGrid)

        ' --- Top panel (TableLayoutPanel) ---
        tlpTop = New TableLayoutPanel() With {
            .Dock = DockStyle.Top,
            .Height = 56,
            .ColumnCount = 6,
            .Padding = New Padding(8),
            .AutoSize = False
        }
        tlpTop.RowStyles.Clear()
        tlpTop.RowStyles.Add(New RowStyle(SizeType.Absolute, 40))
        tlpTop.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))
        tlpTop.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 40))
        tlpTop.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))
        tlpTop.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 40))
        tlpTop.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))
        tlpTop.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 120))

        ' Label Model Pack
        Dim lblMp As New Label() With {
            .Text = "Model Pack:",
            .Dock = DockStyle.Fill,
            .TextAlign = ContentAlignment.MiddleRight,
            .AutoSize = False
        }
        cmbModelPack.Dock = DockStyle.Fill
        cmbModelPack.Margin = New Padding(4, 6, 4, 6)
        cmbModelPack.Height = 24

        ' Label Tipo
        Dim lblTipo As New Label() With {
            .Text = "Tipo oggetto:",
            .Dock = DockStyle.Fill,
            .TextAlign = ContentAlignment.MiddleRight,
            .AutoSize = False
        }
        cmbTipoOggetto.Dock = DockStyle.Fill
        cmbTipoOggetto.Margin = New Padding(4, 6, 4, 6)
        cmbTipoOggetto.Height = 24

        ' Bottone Carica Griglia
        btnCaricaGriglia.Dock = DockStyle.Fill
        btnCaricaGriglia.AutoSize = False
        btnCaricaGriglia.Height = 26
        btnCaricaGriglia.Margin = New Padding(6, 8, 6, 8)

        ' Bottone Annulla
        btnAnnulla.AutoSize = False
        btnAnnulla.Height = 26
        btnAnnulla.Margin = New Padding(6, 8, 6, 8)
        btnAnnulla.Dock = DockStyle.Fill

        ' Aggiungi controlli al top
        tlpTop.Controls.Add(lblMp, 0, 0)
        tlpTop.Controls.Add(cmbModelPack, 1, 0)
        tlpTop.Controls.Add(lblTipo, 2, 0)
        tlpTop.Controls.Add(cmbTipoOggetto, 3, 0)
        tlpTop.Controls.Add(btnCaricaGriglia, 4, 0)
        tlpTop.Controls.Add(btnAnnulla, 5, 0)

        lblMp.MinimumSize = New Size(0, 24)
        lblTipo.MinimumSize = New Size(0, 24)

        ' --- FlowLayoutPanel (griglia) ---
        flpGrid.Dock = DockStyle.Fill
        flpGrid.Padding = New Padding(12, 12, 12, 12)
        flpGrid.FlowDirection = FlowDirection.LeftToRight
        flpGrid.Margin = New Padding(0)
        flpGrid.BackColor = Color.White
        flpGrid.WrapContents = True
        flpGrid.AutoScroll = True

        ' --- Status panel (non usato per lblStatus ora, ma mantenuto) ---
        statusPanel = New Panel() With {.Dock = DockStyle.Bottom, .Height = 0}

        ' --- Panel bottoni inferiore (creato in runtime) ---
        panelButtons = New Panel() With {
            .Dock = DockStyle.Bottom,
            .Height = 72,
            .Padding = New Padding(8),
            .BackColor = SystemColors.Control
        }

        ' Configura lblStatus: spostato nel panelButtons e con font più grande
        lblStatus.Font = New Font("Segoe UI", 11.0F, FontStyle.Bold)
        lblStatus.AutoSize = False
        lblStatus.TextAlign = ContentAlignment.MiddleLeft
        lblStatus.Dock = DockStyle.None
        lblStatus.Width = 420
        lblStatus.Height = 32

        ' Aggiungi i controlli al form nell'ordine corretto (top, fill, status, buttons)
        Me.Controls.Add(flpGrid)
        Me.Controls.Add(statusPanel)
        Me.Controls.Add(panelButtons)
        Me.Controls.Add(tlpTop)

        ' Crea o trova BtnSalva e BtnChiudi e aggiungili al panelButtons
        CreateOrAttachBottomButtons()

        ' Event handlers
        AddHandler Me.Load, AddressOf Form_Load
        AddHandler Me.Resize, AddressOf Form_Resize
        AddHandler Me.FormClosing, AddressOf CaricaOggettiModelPack_FormClosing
        AddHandler btnCaricaGriglia.Click, Async Sub(sender, e) Await ReloadGridIfTipoSelectedAsync()

        ' Gestori per ripristinare il cursore standard quando si passa su Annulla
        AddHandler btnAnnulla.MouseEnter, Sub(sender, e)
                                              Cursor.Current = Cursors.Default
                                              Me.UseWaitCursor = False
                                          End Sub

        AddHandler btnAnnulla.MouseLeave, Sub(sender, e)
                                              Me.UseWaitCursor = _isLoading
                                              If _isLoading Then
                                                  Cursor.Current = Cursors.WaitCursor
                                              Else
                                                  Cursor.Current = Cursors.Default
                                              End If
                                          End Sub
    End Sub

    ' =======================
    ' Crea o riattacca i bottoni Salva/Chiudi nel panel inferiore
    ' =======================
    Private Sub CreateOrAttachBottomButtons()
        ' Trova BtnSalva nel form (se esiste nel progetto)
        Dim foundSave = Me.Controls.Find("BtnSalva", True)
        If foundSave IsNot Nothing AndAlso foundSave.Length > 0 Then
            BtnSalva = TryCast(foundSave(0), Button)
            If BtnSalva.Parent IsNot Nothing Then BtnSalva.Parent.Controls.Remove(BtnSalva)
        Else
            BtnSalva = New Button() With {
                .Name = "BtnSalva",
                .Text = "Salva",
                .AutoSize = False,
                .Width = 120,
                .Height = 36
            }
        End If

        ' Trova BtnChiudi nel form (se esiste)
        Dim foundClose = Me.Controls.Find("BtnChiudi", True)
        If foundClose IsNot Nothing AndAlso foundClose.Length > 0 Then
            BtnChiudi = TryCast(foundClose(0), Button)
            If BtnChiudi.Parent IsNot Nothing Then BtnChiudi.Parent.Controls.Remove(BtnChiudi)
        Else
            BtnChiudi = New Button() With {
                .Name = "BtnChiudi",
                .Text = "Chiudi",
                .AutoSize = False,
                .Width = 120,
                .Height = 36
            }
        End If

        ' Rimuovi lblStatus da eventuale parent precedente e aggiungila al panelButtons
        If lblStatus.Parent IsNot Nothing Then lblStatus.Parent.Controls.Remove(lblStatus)
        panelButtons.Controls.Add(lblStatus)

        ' Posiziona i bottoni a destra nel panelButtons
        BtnSalva.Anchor = AnchorStyles.Right Or AnchorStyles.Bottom
        BtnChiudi.Anchor = AnchorStyles.Right Or AnchorStyles.Bottom

        ' Aggiungi i bottoni al panel
        panelButtons.Controls.Add(BtnSalva)
        panelButtons.Controls.Add(BtnChiudi)

        ' Collega gli handler
        RemoveHandler BtnSalva.Click, AddressOf BtnSalva_Click
        AddHandler BtnSalva.Click, AddressOf BtnSalva_Click

        RemoveHandler BtnChiudi.Click, AddressOf BtnChiudi_Click
        AddHandler BtnChiudi.Click, AddressOf BtnChiudi_Click

        ' Gestori cursore per BtnChiudi (come per Annulla)
        AddHandler BtnChiudi.MouseEnter, Sub(sender, e)
                                             Cursor.Current = Cursors.Default
                                             Me.UseWaitCursor = False
                                         End Sub
        AddHandler BtnChiudi.MouseLeave, Sub(sender, e)
                                             Me.UseWaitCursor = _isLoading
                                             If _isLoading Then
                                                 Cursor.Current = Cursors.WaitCursor
                                             Else
                                                 Cursor.Current = Cursors.Default
                                             End If
                                         End Sub

        ' Forza layout iniziale
        PositionBottomButtons()
    End Sub

    Private Sub PositionBottomButtons()
        If panelButtons Is Nothing Then Return

        ' Posizione label a sinistra con margine
        Dim leftMargin As Integer = 12
        lblStatus.Left = leftMargin
        lblStatus.Top = (panelButtons.ClientSize.Height - lblStatus.Height) \ 2

        ' Calcola posizione bottoni a destra
        Dim spacing As Integer = 8
        Dim rightMargin As Integer = 12
        Dim totalWidth As Integer = BtnSalva.Width + spacing + BtnChiudi.Width
        Dim startX As Integer = Math.Max(panelButtons.ClientSize.Width - totalWidth - rightMargin, lblStatus.Left + lblStatus.Width + 12)

        BtnChiudi.Left = startX + BtnSalva.Width + spacing
        BtnChiudi.Top = (panelButtons.ClientSize.Height - BtnChiudi.Height) \ 2

        BtnSalva.Left = startX
        BtnSalva.Top = (panelButtons.ClientSize.Height - BtnSalva.Height) \ 2

        ' Assicura visibilità
        lblStatus.BringToFront()
        BtnSalva.BringToFront()
        BtnChiudi.BringToFront()
    End Sub

    Private Sub BtnChiudi_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    ' =======================
    ' FormClosing: cancella token e marca chiusura
    ' =======================
    Private Sub CaricaOggettiModelPack_FormClosing(sender As Object, e As FormClosingEventArgs)
        _isClosing = True
        Try
            If _cts IsNot Nothing AndAlso Not _cts.IsCancellationRequested Then
                _cts.Cancel()
            End If
        Catch
            ' ignore
        End Try
    End Sub

    ' =======================
    ' Load iniziale: carica solo combobox, non la griglia
    ' =======================
    Private Async Sub Form_Load(sender As Object, e As EventArgs)
        HookExistingSaveButton()
        Await LoadModelPacksAsync()
        Await LoadTipiAsync()
        UpdateGridLayout()
    End Sub

    Private Sub HookExistingSaveButton()
        Dim found = Me.Controls.Find("BtnSalva", True)
        If found IsNot Nothing AndAlso found.Length > 0 Then
            BtnSalva = TryCast(found(0), Button)
            If BtnSalva IsNot Nothing Then
                RemoveHandler BtnSalva.Click, AddressOf BtnSalva_Click
                AddHandler BtnSalva.Click, AddressOf BtnSalva_Click
            End If
        End If
    End Sub

    Private Sub Form_Resize(sender As Object, e As EventArgs)
        UpdateGridLayout()
        PositionBottomButtons()
    End Sub

    Private Sub UpdateGridLayout()
        Dim totalCellWidth As Integer = Columns * (CellWidth + 16) + 40
        flpGrid.SuspendLayout()
        flpGrid.Width = Math.Max(Me.ClientSize.Width, totalCellWidth)
        flpGrid.ResumeLayout()
    End Sub

    ' =======================
    ' Eventi combo
    ' =======================
    Private Sub cmbModelPack_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbModelPack.SelectedIndexChanged
        Dim ci = TryCast(cmbModelPack.SelectedItem, ComboItem)
        _currentModelPackId = If(ci IsNot Nothing, ci.Value, Nothing)
        lblStatus.Text = $"ModelPack selezionato: {_currentModelPackId}"
    End Sub

    Private Sub cmbTipoOggetto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbTipoOggetto.SelectedIndexChanged
        Dim ci = TryCast(cmbTipoOggetto.SelectedItem, ComboItem)
        _currentTipoId = If(ci IsNot Nothing, ci.Value, Nothing)
        lblStatus.Text = $"Tipo selezionato: {_currentTipoId}"
        ' Non carichiamo la griglia automaticamente; l'utente deve premere "Carica Griglia"
    End Sub

    Private Async Function ReloadGridIfTipoSelectedAsync() As Task
        If String.IsNullOrWhiteSpace(_currentTipoId) Then
            MessageBox.Show("Seleziona un Tipo Oggetto Lavorazione.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        Await LoadOggettiAndBuildGridAsync(_currentTipoId)
    End Function

    ' =======================
    ' DB loading
    ' =======================
    Private Async Function LoadModelPacksAsync() As Task
        Dim dt As New DataTable()
        Using cn As New SqlConnection(ConnString)
            Using cmd As New SqlCommand("
            SELECT IdModelPack, ISNULL(Descrizione, '') AS Descrizione
            FROM dbo.Mov_ModelPack
            ORDER BY IdModelPack", cn)
                Await cn.OpenAsync()
                Using rdr = Await cmd.ExecuteReaderAsync()
                    dt.Load(rdr)
                End Using
            End Using
        End Using

        cmbModelPack.Items.Clear()

        ' Aggiungi placeholder come primo elemento
        cmbModelPack.Items.Add(New ComboItem("Selezionare un valore...", ""))

        For Each r As DataRow In dt.Rows
            Dim id = r.Field(Of String)("IdModelPack")
            Dim descr = r.Field(Of String)("Descrizione")
            cmbModelPack.Items.Add(New ComboItem($"{id} - {descr}", id))
        Next

        ' Preselezione: se _initialModelPackId è valorizzato, prova a trovarlo negli items
        If Not String.IsNullOrWhiteSpace(_initialModelPackId) Then
            Dim foundIndex As Integer = -1
            For i As Integer = 0 To cmbModelPack.Items.Count - 1
                Dim ci = TryCast(cmbModelPack.Items(i), ComboItem)
                If ci IsNot Nothing AndAlso String.Equals(ci.Value, _initialModelPackId, StringComparison.OrdinalIgnoreCase) Then
                    foundIndex = i
                    Exit For
                End If
            Next

            If foundIndex >= 0 Then
                cmbModelPack.SelectedIndex = foundIndex
            Else
                ' lascia il placeholder selezionato
                cmbModelPack.SelectedIndex = 0
            End If
        Else
            ' nessuna preselezione richiesta: mantieni il placeholder
            cmbModelPack.SelectedIndex = 0
        End If
    End Function

    Private Async Function LoadTipiAsync() As Task
        Dim dt As New DataTable()
        Using cn As New SqlConnection(ConnString)
            Using cmd As New SqlCommand("
            SELECT IdTipoOggettoLavorazione, ISNULL(Descrizione, '') AS Descrizione
            FROM dbo.Tab_TipiOggettoLavorazione
            ORDER BY IdTipoOggettoLavorazione", cn)
                Await cn.OpenAsync()
                Using rdr = Await cmd.ExecuteReaderAsync()
                    dt.Load(rdr)
                End Using
            End Using
        End Using

        cmbTipoOggetto.Items.Clear()

        ' Aggiungi placeholder come primo elemento
        cmbTipoOggetto.Items.Add(New ComboItem("Selezionare un valore...", ""))

        For Each r As DataRow In dt.Rows
            Dim id = r.Field(Of String)("IdTipoOggettoLavorazione")
            Dim descr = r.Field(Of String)("Descrizione")
            cmbTipoOggetto.Items.Add(New ComboItem($"{id} - {descr}", id))
        Next

        ' Mantieni il placeholder selezionato all'apertura
        cmbTipoOggetto.SelectedIndex = 0
    End Function


    Private Async Function LoadOggettiAndBuildGridAsync(tipoId As String) As Task
        _cts?.Dispose()
        _cts = New CancellationTokenSource()
        Dim token = _cts.Token

        SetLoading(True, "Caricamento oggetti ...")
        btnAnnulla.Enabled = True
        _selectedKeys.Clear()
        _items.Clear()

        Dim dt As New DataTable()
        Using cn As New SqlConnection(ConnString)
            Using cmd As New SqlCommand("
                SELECT IdOggettoLavorazione,
                       ISNULL(Descrizione, '') AS Descrizione,
                       TipoOggettoLavorazioneId,
                       ISNULL(FileOggettoLavorazione, '') AS FileOggettoLavorazione,
                       ISNULL(RiusoOggetto, 0) AS RiusoOggetto
                FROM dbo.Tab_OggettiLavorazione
                WHERE TipoOggettoLavorazioneId = @TipoId
                ORDER BY IdOggettoLavorazione", cn)
                cmd.Parameters.AddWithValue("@TipoId", tipoId)
                Await cn.OpenAsync(token)
                Using rdr = Await cmd.ExecuteReaderAsync(token)
                    dt.Load(rdr)
                End Using
            End Using
        End Using

        For Each r As DataRow In dt.Rows
            Dim it As New OggettoItem With {
                .IdOggetto = r.Field(Of String)("IdOggettoLavorazione"),
                .Descrizione = r.Field(Of String)("Descrizione"),
                .TipoId = r.Field(Of String)("TipoOggettoLavorazioneId"),
                .FilePath = r.Field(Of String)("FileOggettoLavorazione"),
                .RiusoOggetto = Convert.ToBoolean(r("RiusoOggetto"))
            }
            _items.Add(it)
        Next

        Try
            Await BuildGridAsync(_items, token)
        Finally
            btnAnnulla.Enabled = False
            SetLoading(False, $"Oggetti caricati: {flpGrid.Controls.Count}")
            _cts?.Dispose()
            _cts = Nothing
        End Try
    End Function

    ' =======================
    ' BuildGridAsync con gestione cancellazione e SafeBeginInvoke
    ' =======================
    Private Async Function BuildGridAsync(items As List(Of OggettoItem), token As CancellationToken) As Task
        flpGrid.SuspendLayout()
        flpGrid.Controls.Clear()
        prgLoad.Visible = True
        prgLoad.Style = ProgressBarStyle.Blocks
        prgLoad.Minimum = 0
        prgLoad.Maximum = Math.Max(items.Count, 1)
        prgLoad.Value = 0

        Dim addedKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        Try
            For i As Integer = 0 To items.Count - 1
                If token.IsCancellationRequested OrElse _isClosing Then Exit For

                Dim it = items(i)
                Dim primaryPath As String = it.FilePath

                ' Determina estensione
                Dim ext As String = ""
                Try
                    ext = If(String.IsNullOrWhiteSpace(primaryPath), "", Path.GetExtension(primaryPath).ToLowerInvariant())
                Catch
                    ext = ""
                End Try

                Dim previewPath As String = String.Empty

                If ext = ".mp4" OrElse ext = ".mov" Then
                    ' Video: usa immagine da Sys_parametri
                    Dim videoPreview As String = Await GetVideoPreviewImagePathAsync()
                    If Not String.IsNullOrWhiteSpace(videoPreview) AndAlso File.Exists(videoPreview) Then
                        previewPath = videoPreview
                    Else
                        previewPath = ResolvePreviewPath(primaryPath)
                    End If
                ElseIf ext = ".png" OrElse ext = ".jpg" OrElse ext = ".jpeg" Then
                    ' Immagini: usa direttamente il file primario
                    previewPath = primaryPath
                Else
                    ' Altri tipi: prova a leggere FilePreviewOggetto da Mov_ModelPackOggetti
                    Dim previewFromMov As String = Await GetPreviewFromMovModelPackOggettiAsync(it.IdOggetto)
                    If Not String.IsNullOrWhiteSpace(previewFromMov) AndAlso File.Exists(previewFromMov) Then
                        previewPath = previewFromMov
                    Else
                        ' fallback standard
                        previewPath = ResolvePreviewPath(primaryPath)
                    End If
                End If

                Dim key As String = GetBaseNameWithoutExtension(primaryPath)
                If String.IsNullOrWhiteSpace(key) Then key = GetBaseNameWithoutExtension(previewPath)

                If addedKeys.Contains(key) Then
                    If Not token.IsCancellationRequested Then
                        SafeBeginInvoke(Sub()
                                            prgLoad.Value = Math.Min(prgLoad.Value + 1, prgLoad.Maximum)
                                            lblStatus.Text = $"Caricati {prgLoad.Value}/{prgLoad.Maximum}"
                                        End Sub)
                    End If
                    Continue For
                End If
                addedKeys.Add(key)

                Dim cell As Panel = CreateCell(it, key, primaryPath)
                SafeBeginInvoke(Sub() flpGrid.Controls.Add(cell))

                Try
                    Dim img As Image = Await Task.Run(Function()
                                                          token.ThrowIfCancellationRequested()
                                                          Return LoadThumbnailSafely(previewPath, _placeholderImage)
                                                      End Function, token)

                    If token.IsCancellationRequested OrElse _isClosing Then
                        If img IsNot Nothing AndAlso Not Object.ReferenceEquals(img, _placeholderImage) Then img.Dispose()
                        Exit For
                    End If

                    SafeBeginInvoke(Sub()
                                        Dim pb = TryCast(cell.Controls("pb"), PictureBox)
                                        If pb IsNot Nothing Then pb.Image = img
                                        prgLoad.Value = Math.Min(prgLoad.Value + 1, prgLoad.Maximum)
                                        lblStatus.Text = $"Caricati {prgLoad.Value}/{prgLoad.Maximum}"
                                    End Sub)
                Catch ex As OperationCanceledException
                    Exit For
                End Try
            Next

        Catch ex As OperationCanceledException
            ' Cancellazione: termina pulito
        Catch ex As Exception
            SafeBeginInvoke(Sub()
                                MessageBox.Show($"Errore durante il caricamento: {ex.Message}", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            End Sub)
        Finally
            SafeBeginInvoke(Sub()
                                prgLoad.Visible = False
                                btnAnnulla.Enabled = False
                                SetLoading(False, $"Oggetti caricati: {flpGrid.Controls.Count}")
                            End Sub)
            flpGrid.ResumeLayout()
        End Try
    End Function

    Private Async Function GetPreviewFromMovModelPackOggettiAsync(oggettoId As String) As Task(Of String)
        If String.IsNullOrWhiteSpace(oggettoId) Then Return String.Empty
        Try
            Dim result As String = String.Empty
            Using cn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand("
                SELECT TOP 1 ISNULL(FilePreviewOggetto, '') AS FilePreviewOggetto
                FROM dbo.Mov_ModelPackOggetti
                WHERE OggettoLavorazioneId = @OggettoId", cn)
                    cmd.Parameters.AddWithValue("@OggettoId", oggettoId)
                    Await cn.OpenAsync()
                    Dim obj = Await cmd.ExecuteScalarAsync()
                    If obj IsNot Nothing AndAlso Not Convert.IsDBNull(obj) Then
                        result = Convert.ToString(obj)
                    End If
                End Using
            End Using
            Return If(String.IsNullOrWhiteSpace(result), String.Empty, result)
        Catch
            Return String.Empty
        End Try
    End Function

    Private Async Function GetVideoPreviewImagePathAsync() As Task(Of String)
        Try
            Dim result As String = String.Empty
            Using cn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand("
                SELECT TOP 1 Valore
                FROM dbo.Sys_parametri
                WHERE Descrizione = @Descrizione", cn)
                    cmd.Parameters.AddWithValue("@Descrizione", "PercorsoImmagineVideo")
                    Await cn.OpenAsync()
                    Dim obj = Await cmd.ExecuteScalarAsync()
                    If obj IsNot Nothing AndAlso Not Convert.IsDBNull(obj) Then
                        result = Convert.ToString(obj)
                    End If
                End Using
            End Using
            Return If(String.IsNullOrWhiteSpace(result), String.Empty, result)
        Catch
            Return String.Empty
        End Try
    End Function

    Private Function CreateCell(it As OggettoItem, key As String, primaryPath As String) As Panel
        Dim panel As New Panel() With {
            .Width = CellWidth,
            .Height = CellHeight,
            .Margin = New Padding(8),
            .BackColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle,
            .Tag = key
        }
        EnableDoubleBuffering(panel)

        Dim pb As New PictureBox() With {
            .Name = "pb",
            .Width = CellWidth - 20,
            .Height = CellHeight - 100,
            .Top = 10,
            .Left = 10,
            .SizeMode = PictureBoxSizeMode.Zoom,
            .Image = _placeholderImage
        }
        AddHandler pb.Click, Sub(s, e) ToggleSelection(panel)

        Dim chk As New CheckBox() With {
            .Name = "chk",
            .Text = "Seleziona",
            .AutoSize = True,
            .Top = pb.Bottom + 4,
            .Left = 10
        }
        AddHandler chk.CheckedChanged, Sub(s, e) ToggleSelection(panel, chk.Checked)

        ' Determina il nome da visualizzare: preferisci sempre il nome primario (FileOggettoLavorazione).
        ' Usa il preview solo come fallback se il primario è vuoto o non disponibile.
        Dim originalFileName As String = ""
        If Not String.IsNullOrWhiteSpace(primaryPath) Then
            Try
                originalFileName = Path.GetFileName(primaryPath)
            Catch
                originalFileName = ""
            End Try
        End If

        Dim displayFileName As String = originalFileName
        If String.IsNullOrWhiteSpace(displayFileName) Then
            Dim previewPath = ResolvePreviewPath(primaryPath)
            If Not String.IsNullOrWhiteSpace(previewPath) Then
                Try
                    displayFileName = Path.GetFileName(previewPath)
                Catch
                    displayFileName = ""
                End Try
            End If
        End If

        If String.IsNullOrWhiteSpace(displayFileName) Then
            displayFileName = "(nessun file)"
        End If

        Dim lblFile As New Label() With {
            .AutoSize = False,
            .Width = CellWidth - 20,
            .Height = 18,
            .Top = chk.Bottom + 2,
            .Left = 10,
            .ForeColor = Color.DimGray,
            .Text = displayFileName,
            .AutoEllipsis = True
        }
        lblFile.Tag = If(String.IsNullOrWhiteSpace(primaryPath), ResolvePreviewPath(primaryPath), primaryPath)

        Dim lblDesc As New Label() With {
            .AutoSize = False,
            .Width = CellWidth - 20,
            .Height = 18,
            .Top = lblFile.Bottom + 2,
            .Left = 10,
            .ForeColor = Color.Black,
            .Text = $"{it.IdOggetto} - {Truncate(it.Descrizione, 50)}",
            .AutoEllipsis = True
        }

        panel.Controls.Add(pb)
        panel.Controls.Add(chk)
        panel.Controls.Add(lblFile)
        panel.Controls.Add(lblDesc)
        Return panel
    End Function

    Private Sub ToggleSelection(panel As Panel, Optional forceState As Boolean? = Nothing)
        Dim key = Convert.ToString(panel.Tag)
        Dim chk = TryCast(panel.Controls("chk"), CheckBox)

        Dim newState As Boolean = If(forceState.HasValue, forceState.Value, Not _selectedKeys.Contains(key))

        If newState Then
            _selectedKeys.Add(key)
            panel.BackColor = Color.AliceBlue
        Else
            _selectedKeys.Remove(key)
            panel.BackColor = Color.White
        End If

        If chk IsNot Nothing AndAlso chk.Checked <> newState Then chk.Checked = newState
    End Sub

    ' Annulla caricamento
    Private Sub btnAnnulla_Click(sender As Object, e As EventArgs) Handles btnAnnulla.Click
        If _cts IsNot Nothing AndAlso Not _cts.IsCancellationRequested Then
            _cts.Cancel()
        End If
        btnAnnulla.Enabled = False
        lblStatus.Text = "Caricamento annullato."
        prgLoad.Visible = False
    End Sub

    ' Salvataggio multiplo
    Private Async Sub BtnSalva_Click(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(_currentModelPackId) Then
            MessageBox.Show("Seleziona un Model Pack.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        If String.IsNullOrWhiteSpace(_currentTipoId) Then
            MessageBox.Show("Seleziona un Tipo Oggetto Lavorazione.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        If _selectedKeys.Count = 0 Then
            MessageBox.Show("Seleziona almeno un oggetto da salvare.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim conferma = MessageBox.Show($"Confermi il salvataggio di {_selectedKeys.Count} oggetti nel ModelPack {_currentModelPackId}?",
                                       "Conferma", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If conferma <> DialogResult.Yes Then Return

        SetLoading(True, "Salvataggio in corso...")

        Dim selectedItems = _items.Where(Function(it)
                                             Dim key = GetBaseNameWithoutExtension(it.FilePath)
                                             If String.IsNullOrWhiteSpace(key) Then
                                                 key = GetBaseNameWithoutExtension(ResolvePreviewPath(it.FilePath))
                                             End If
                                             Return _selectedKeys.Contains(key)
                                         End Function).ToList()

        Dim inserted As Integer = 0
        Try
            Using cn As New SqlConnection(ConnString)
                Await cn.OpenAsync()
                Using tr = cn.BeginTransaction()
                    For Each it In selectedItems
                        Using cmd As New SqlCommand("
                            INSERT INTO dbo.Mov_ModelPackOggetti
                            (Descrizione, ModelPackId, TipoOggettoLavorazioneId, OggettoLavorazioneId, FileOggettoLavorazione)
                            VALUES (@Descrizione, @ModelPackId, @TipoId, @OggettoId, @File)", cn, tr)
                            cmd.Parameters.AddWithValue("@Descrizione", If(String.IsNullOrWhiteSpace(it.Descrizione), CType(DBNull.Value, Object), it.Descrizione))
                            cmd.Parameters.AddWithValue("@ModelPackId", _currentModelPackId)
                            cmd.Parameters.AddWithValue("@TipoId", _currentTipoId)
                            cmd.Parameters.AddWithValue("@OggettoId", it.IdOggetto)
                            cmd.Parameters.AddWithValue("@File", If(String.IsNullOrWhiteSpace(it.FilePath), CType(DBNull.Value, Object), it.FilePath))
                            Await cmd.ExecuteNonQueryAsync()
                            inserted += 1
                        End Using
                    Next
                    tr.Commit()
                End Using
            End Using

            MessageBox.Show($"Salvataggio completato. Inseriti {inserted} record.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show($"Errore nel salvataggio: {ex.Message}", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            SetLoading(False, "Pronto")
        End Try
    End Sub

    ' =======================
    ' SafeBeginInvoke helper
    ' =======================
    Private Sub SafeBeginInvoke(action As Action)
        If _isClosing Then Return
        If Me Is Nothing Then Return
        If Not Me.IsHandleCreated Then Return
        If Me.Disposing OrElse Me.IsDisposed Then Return

        Try
            Me.BeginInvoke(action)
        Catch ex As InvalidOperationException
            ' Il form potrebbe essere in fase di chiusura: ignoriamo in modo sicuro
        Catch
            ' Ignora altre eccezioni di invocazione
        End Try
    End Sub

    ' =======================
    ' Utility
    ' =======================
    Private Shared Sub EnableDoubleBuffering(ctrl As Control)
        Dim pi = ctrl.GetType().GetProperty("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
        If pi IsNot Nothing Then pi.SetValue(ctrl, True)
    End Sub

    Private Sub SetLoading(isLoading As Boolean, Optional statusText As String = Nothing)
        _isLoading = isLoading
        lblStatus.Text = If(statusText, "")
        Me.UseWaitCursor = isLoading
        cmbModelPack.Enabled = Not isLoading
        cmbTipoOggetto.Enabled = Not isLoading
        prgLoad.Visible = isLoading
        prgLoad.Style = If(isLoading, ProgressBarStyle.Marquee, ProgressBarStyle.Blocks)
        If BtnSalva IsNot Nothing Then BtnSalva.Enabled = Not isLoading
        btnAnnulla.Enabled = isLoading
    End Sub

    Private Shared Function Truncate(s As String, maxLen As Integer) As String
        If String.IsNullOrEmpty(s) Then Return ""
        If s.Length <= maxLen Then Return s
        Return s.Substring(0, maxLen - 1) & "…"
    End Function

    Private Shared Function CreatePlaceholderBitmap(w As Integer, h As Integer) As Bitmap
        Dim bmp As New Bitmap(w, h)
        Using g = Graphics.FromImage(bmp)
            g.Clear(Color.Gainsboro)
            Using p As New Pen(Color.DarkGray, 2)
                g.DrawRectangle(p, 2, 2, w - 4, h - 4)
            End Using
            Using f As New Font("Segoe UI", 10, FontStyle.Italic)
                Dim txt = "Anteprima non disponibile"
                Dim sz = g.MeasureString(txt, f)
                g.DrawString(txt, f, Brushes.DimGray, (w - sz.Width) / 2, (h - sz.Height) / 2)
            End Using
        End Using
        Return bmp
    End Function

    Private Shared Function LoadThumbnailSafely(path As String, fallback As Image) As Image
        Try
            If String.IsNullOrWhiteSpace(path) Then Return fallback
            If Not File.Exists(path) Then Return fallback
            Using fs As New FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read)
                Dim img = Image.FromStream(fs)
                Dim maxSide As Integer = 1000
                Dim scale As Double = Math.Min(maxSide / img.Width, maxSide / img.Height)
                If scale < 1 Then
                    Dim w = Math.Max(64, CInt(img.Width * scale))
                    Dim h = Math.Max(64, CInt(img.Height * scale))
                    Dim thumb = New Bitmap(img, New Size(w, h))
                    Return thumb
                Else
                    Return img
                End If
            End Using
        Catch
            Return fallback
        End Try
    End Function

    ' Risolve path anteprima: usa png/jpg se l'originale non è immagine
    Private Shared Function ResolvePreviewPath(originalPath As String) As String
        If String.IsNullOrWhiteSpace(originalPath) Then Return ""
        Dim ext = Path.GetExtension(originalPath).ToLowerInvariant()
        If (ext = ".png" OrElse ext = ".jpg" OrElse ext = ".jpeg") AndAlso File.Exists(originalPath) Then
            Return originalPath
        End If

        Dim dir = Path.GetDirectoryName(originalPath)
        Dim baseName = Path.GetFileNameWithoutExtension(originalPath)
        If String.IsNullOrWhiteSpace(dir) OrElse String.IsNullOrWhiteSpace(baseName) Then
            Return originalPath
        End If

        Dim pngPath = Path.Combine(dir, baseName & ".png")
        If File.Exists(pngPath) Then Return pngPath

        Dim jpgPath = Path.Combine(dir, baseName & ".jpg")
        If File.Exists(jpgPath) Then Return jpgPath

        Dim jpegPath = Path.Combine(dir, baseName & ".jpeg")
        If File.Exists(jpegPath) Then Return jpegPath

        Return originalPath
    End Function

    Private Shared Function GetBaseNameWithoutExtension(path As String) As String
        If String.IsNullOrWhiteSpace(path) Then Return ""
        Return path.BaseNameWithoutExtension()
    End Function

    ' Models
    Private Class ComboItem
        Public Property Text As String
        Public Property Value As String
        Public Sub New(text As String, value As String)
            Me.Text = text
            Me.Value = value
        End Sub
        Public Overrides Function ToString() As String
            Return Text
        End Function
    End Class

    Private Class OggettoItem
        Public Property IdOggetto As String
        Public Property Descrizione As String
        Public Property TipoId As String
        Public Property FilePath As String
        Public Property RiusoOggetto As Boolean
    End Class

End Class
Public Module StringPathExtensions
    <Extension()>
    Public Function BaseNameWithoutExtension(s As String) As String
        If String.IsNullOrWhiteSpace(s) Then Return ""
        Return Path.GetFileNameWithoutExtension(s)
    End Function
End Module
