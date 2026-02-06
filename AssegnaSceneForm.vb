Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Security.Cryptography
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Data.SqlClient

Public Class AssegnaSceneForm
    Inherits Form

    Private ReadOnly _storyboardId As String
    Private ReadOnly _storyboardDesc As String
    Private ReadOnly _outDir As String

    ' UI controls
    Private tblContainer As FlowLayoutPanel
    Private scrollPanel As Panel

    Private cmbScene As ComboBox
    Private btnSave As Button
    Private btnClose As Button
    Private btnSelectAll As Button
    Private btnClearAll As Button
    Private btnRefresh As Button
    Private lblInfo As Label
    Private lblProgress As Label

    ' Overlay (maschera caricamento, senza ProgressBar)
    Private overlay As Panel
    Private overlayLabel As Label

    ' Layout
    Private ReadOnly previewW As Integer = 200
    Private ReadOnly previewH As Integer = 130
    Private ReadOnly extraPanelWidth As Integer = 40
    Private ReadOnly cols As Integer = 4
    Private ReadOnly marginPx As Integer = 8
    Private ReadOnly leftPad As Integer = 12
    Private ReadOnly topPad As Integer = 12

    ' Dati
    Private panelFiles As New List(Of String)()
    Private items As New List(Of PanelItem)()
    Private cts As CancellationTokenSource = Nothing

    Private Shared ReadOnly PlaceholderFont As New Font(SystemFonts.DefaultFont.FontFamily, 9.0F, FontStyle.Regular)

    Private Class PanelItem
        Public Property FilePath As String
        Public Property Pic As PictureBox
        Public Property Cb As CheckBox
        Public Property Assigned As Boolean
        Public Property AssignedInfo As String
    End Class

    ' Compositing per ridurre flicker/strisce grigie
    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim cp = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or &H2000000 ' WS_EX_COMPOSITED
            Return cp
        End Get
    End Property

    Public Sub New(storyboardId As String, storyboardDesc As String, outDir As String)
        _storyboardId = storyboardId
        _storyboardDesc = storyboardDesc
        _outDir = outDir

        InitializeComponent()
        EnableDoubleBuffering()

        AddHandler Me.Load, Sub(s, e)
                                Try
                                    RipristinaPosizioneForm(Me)
                                Catch
                                End Try
                            End Sub

        AddHandler Me.Shown, AddressOf AssegnaSceneForm_Shown
        AddHandler Me.FormClosing, AddressOf AssegnaSceneForm_FormClosing
    End Sub

    Private Sub InitializeComponent()
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = Color.White
        Me.SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.UserPaint Or ControlStyles.OptimizedDoubleBuffer, True)
        Me.UpdateStyles()

        Dim singlePanelWidth = previewW + extraPanelWidth + marginPx * 2
        Dim formWidth As Integer = Math.Max(820, (singlePanelWidth) * cols + 220)
        Dim formHeight As Integer = Math.Max(560, (previewH + marginPx) * 3 + 260)
        Me.Size = New Size(formWidth, formHeight)
        Me.MinimumSize = New Size(760, 520)
        Me.Text = $"Assegna panel - Storyboard {_storyboardId}"

        Dim top As Integer = topPad

        Dim lblId As New Label() With {.Left = leftPad, .Top = top, .Width = 120, .Text = "Id storyboard:"}
        Me.Controls.Add(lblId)

        Dim lblIdVal As New Label() With {.Left = lblId.Right + 8, .Top = top, .Width = 360, .Text = _storyboardId}
        Me.Controls.Add(lblIdVal)

        top += 26
        Dim lblDesc As New Label() With {.Left = leftPad, .Top = top, .Width = 120, .Text = "Descrizione:"}
        Me.Controls.Add(lblDesc)

        Dim lblDescVal As New Label() With {.Left = lblDesc.Right + 8, .Top = top, .Width = 360, .Text = _storyboardDesc}
        Me.Controls.Add(lblDescVal)

        top += 30
        Dim lblScene As New Label() With {.Left = leftPad, .Top = top, .Width = 120, .Text = "Seleziona scena:"}
        Me.Controls.Add(lblScene)

        cmbScene = New ComboBox() With {.Left = lblScene.Right + 8, .Top = top - 2, .Width = 260, .DropDownStyle = ComboBoxStyle.DropDownList}
        Me.Controls.Add(cmbScene)

        btnSelectAll = New Button() With {.Left = cmbScene.Right + 12, .Top = top - 2, .Width = 120, .Text = "Seleziona tutti"}
        AddHandler btnSelectAll.Click, AddressOf BtnSelectAll_Click
        Me.Controls.Add(btnSelectAll)

        btnClearAll = New Button() With {.Left = btnSelectAll.Right + 8, .Top = top - 2, .Width = 120, .Text = "Deseleziona tutti"}
        AddHandler btnClearAll.Click, AddressOf BtnClearAll_Click
        Me.Controls.Add(btnClearAll)

        btnRefresh = New Button() With {.Left = btnClearAll.Right + 8, .Top = top - 2, .Width = 120, .Text = "Aggiorna"}
        AddHandler btnRefresh.Click, AddressOf BtnRefresh_Click
        Me.Controls.Add(btnRefresh)

        top += 36
        Dim gridWidth = Math.Max((cols * singlePanelWidth), Me.ClientSize.Width - 40)
        Dim rowsVisible As Integer = 3
        Dim gridHeight = rowsVisible * (previewH + 82)

        scrollPanel = New Panel() With {
            .Left = leftPad,
            .Top = top,
            .Width = gridWidth,
            .Height = gridHeight,
            .AutoScroll = True,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right,
            .BackColor = Color.White
        }
        Me.Controls.Add(scrollPanel)

        tblContainer = New FlowLayoutPanel() With {
    .FlowDirection = FlowDirection.LeftToRight,
    .WrapContents = True,
    .AutoSize = True,
    .AutoSizeMode = AutoSizeMode.GrowAndShrink,
    .Dock = DockStyle.Top,
    .Padding = New Padding(marginPx),
    .BackColor = Color.White
}
        scrollPanel.Controls.Add(tblContainer)


        ' Overlay senza ProgressBar: aggiunto come figlio di scrollPanel
        overlay = New Panel() With {
         .BackColor = Color.FromArgb(180, Color.White),
        .Visible = False
        }
        overlayLabel = New Label() With {
        .AutoSize = False,
        .Dock = DockStyle.Top,
        .Height = 40,
        .TextAlign = ContentAlignment.MiddleCenter,
        .Font = New Font(SystemFonts.DefaultFont.FontFamily, 11.0F, FontStyle.Bold),
        .ForeColor = Color.DimGray,
        .Text = "Caricamento in corso ..."
        }
        Dim overlaySpacer As New Panel() With {.Dock = DockStyle.Fill, .BackColor = Color.Transparent}
        overlay.Controls.Add(overlaySpacer)
        overlay.Controls.Add(overlayLabel)

        ' Aggiungi overlay come figlio di scrollPanel e ancoralo
        scrollPanel.Controls.Add(overlay)
        overlay.Location = New Point(0, 0)
        overlay.Size = scrollPanel.ClientSize
        overlay.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom

        lblInfo = New Label() With {
            .Left = leftPad,
            .Top = scrollPanel.Bottom + 8,
            .Width = 520,
            .Height = 40,
            .Text = "Seleziona i panel da assegnare alla scena scelta. I panel già assegnati non sono selezionabili.",
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
        }
        Me.Controls.Add(lblInfo)

        ' Label di progresso persistente (più evidente)
        lblProgress = New Label() With {
        .AutoSize = False,
        .Width = 260,
        .Height = 28,
        .TextAlign = ContentAlignment.MiddleLeft,
        .Font = New Font(SystemFonts.DefaultFont.FontFamily, 11.5F, FontStyle.Bold),
        .ForeColor = Color.DarkBlue,
        .BackColor = Color.Transparent,
       .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left,
        .Text = String.Empty
        }
        Me.Controls.Add(lblProgress)

        btnSave = New Button() With {.Text = "Salva", .Width = 120, .Height = 36, .Anchor = AnchorStyles.Bottom}
        AddHandler btnSave.Click, AddressOf BtnSave_Click
        Me.Controls.Add(btnSave)

        btnClose = New Button() With {.Text = "Chiudi", .Width = 120, .Height = 36, .Anchor = AnchorStyles.Bottom}
        AddHandler btnClose.Click, Sub(s, e) Me.Close()
        Me.Controls.Add(btnClose)

        AddHandler Me.Resize,
            Sub(s, e)
                Dim newGridWidth = Math.Max((cols * (previewW + extraPanelWidth + marginPx * 2)), Me.ClientSize.Width - 40)
                scrollPanel.Left = leftPad
                scrollPanel.Top = top
                scrollPanel.Width = newGridWidth
                scrollPanel.Height = gridHeight

                lblInfo.Top = scrollPanel.Bottom + 8

                Dim buttonsTotalWidth = btnSave.Width + 12 + btnClose.Width
                Dim buttonsLeft = scrollPanel.Left + Math.Max(0, (scrollPanel.Width - buttonsTotalWidth) \ 2)
                btnSave.Left = buttonsLeft
                btnSave.Top = lblInfo.Bottom + 12
                btnClose.Left = btnSave.Right + 12
                btnClose.Top = btnSave.Top

                ' Posiziona lblProgress a sinistra dei pulsanti (stessa altezza)
                Dim progressLeft As Integer = Math.Max(leftPad, scrollPanel.Left)
                lblProgress.Left = progressLeft
                lblProgress.Top = btnSave.Top + (btnSave.Height - lblProgress.Height) \ 2
            End Sub

    End Sub

    Private Sub EnableDoubleBuffering()
        SetDoubleBuffered(scrollPanel, True)
        SetDoubleBuffered(tblContainer, True)
    End Sub

    Private Sub SetDoubleBuffered(ctrl As Control, enable As Boolean)
        Dim pi = ctrl.GetType().GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)
        If pi IsNot Nothing Then
            pi.SetValue(ctrl, enable, Nothing)
        End If
    End Sub

    Private Async Sub AssegnaSceneForm_Shown(sender As Object, e As EventArgs)
        cts = New CancellationTokenSource()
        Dim token = cts.Token

        Try
            Me.UseWaitCursor = True
            ToggleTopButtons(False)

            ' Overlay visibile, griglia nascosta finché non è pronta
            overlay.Visible = True
            overlay.BringToFront()
            overlayLabel.Text = "Caricamento Panels ..."
            scrollPanel.Visible = False

            LoadScenes()

            overlayLabel.Text = "Caricamento lista panel ..."
            Await LoadPanelFilesAsync(token)

            overlayLabel.Text = "Costruzione griglia ..."
            BuildGrid() ' griglia completa con placeholder e sfondi bianchi

            ' Ora la griglia è completamente composta, visibile senza strisce
            scrollPanel.Visible = True

            ' Aggiorna testo overlay con conteggio (senza ProgressBar)
            overlayLabel.Text = "Caricamento anteprime ..."
            Await LoadThumbnailsAsync(token,
               Sub(done, total)
                   UpdateProgressLabels($"Caricamento anteprime {done}/{total}")
               End Sub)

                   ' Mostra il totale per un attimo e poi nascondi overlay
                   overlayLabel.Text = $"Caricamento anteprime {items.Count}/{items.Count}"
                   Await Task.Delay(300)
            overlay.Visible = False

        Catch ex As OperationCanceledException
        Catch ex As Exception
            MDIMessageBox.Show("Errore caricamento preview: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.UseWaitCursor = False
            ToggleTopButtons(True)
            If cts IsNot Nothing Then cts.Dispose() : cts = Nothing
        End Try
    End Sub

    Private Sub AssegnaSceneForm_FormClosing(sender As Object, e As FormClosingEventArgs)
        If cts IsNot Nothing Then cts.Cancel()
        Try
            SalvaPosizioneForm(Me)
        Catch
        End Try
    End Sub

    Private Sub ToggleTopButtons(enabled As Boolean)
        btnSave.Enabled = enabled
        btnSelectAll.Enabled = enabled
        btnClearAll.Enabled = enabled
        cmbScene.Enabled = enabled
        btnRefresh.Enabled = enabled
    End Sub

    Private Sub LoadScenes()
        Try
            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand("SELECT IdProgScena, NumScena, Descrizione 
                                             FROM Mov_StoryboardScene 
                                             WHERE StoryboardId = @id 
                                             ORDER BY NumScena", conn)
                    cmd.Parameters.AddWithValue("@id", _storyboardId)

                    Dim dt As New DataTable()
                    Using da As New SqlDataAdapter(cmd)
                        da.Fill(dt)
                    End Using

                    If dt.Rows.Count = 0 Then
                        MDIMessageBox.Show("Nessuna scena trovata per questo storyboard.", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Information)
                        cmbScene.DataSource = Nothing
                        Return
                    End If

                    If Not dt.Columns.Contains("Display") Then
                        dt.Columns.Add("Display", GetType(String))
                        For Each r As DataRow In dt.Rows
                            r("Display") = $"{Convert.ToString(r("NumScena"))} - {Convert.ToString(r("Descrizione"))}"
                        Next
                    End If

                    cmbScene.DisplayMember = "Display"
                    cmbScene.ValueMember = "NumScena"
                    cmbScene.DataSource = dt
                    cmbScene.SelectedIndex = 0
                End Using
            End Using
        Catch ex As Exception
            MDIMessageBox.Show("Errore caricamento scene: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Async Function LoadPanelFilesAsync(token As CancellationToken) As Task
        panelFiles.Clear()
        items.Clear()

        Await Task.Run(Sub()
                           If Not Directory.Exists(_outDir) Then
                               Throw New DirectoryNotFoundException("Cartella panel non trovata: " & _outDir)
                           End If
                           Dim files = Directory.EnumerateFiles(_outDir, "*.png", SearchOption.TopDirectoryOnly) _
                                               .OrderBy(Function(f) f) _
                                               .ToList()
                           SyncLock panelFiles
                               panelFiles.AddRange(files)
                           End SyncLock
                       End Sub, token)
    End Function

    ' Recupera tutte le assegnazioni in un'unica query e ritorna dictionary ImgPath -> NumScena
    Private Function GetAssignedPanels(connString As String, files As List(Of String)) As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        If files Is Nothing OrElse files.Count = 0 Then Return result

        Const batchLimit As Integer = 1000
        Dim batches = Math.Ceiling(files.Count / CDbl(batchLimit))

        Using conn As New SqlConnection(connString)
            conn.Open()
            For b As Integer = 0 To batches - 1
                Dim startIdx = b * batchLimit
                Dim take = Math.Min(batchLimit, files.Count - startIdx)
                Dim subset = files.Skip(startIdx).Take(take).ToList()

                Dim sql As New StringBuilder()
                sql.Append("SELECT ImgVidPanel, NumScena FROM Mov_StoryboardScenePanel WHERE ImgVidPanel IN (")
                For i = 0 To subset.Count - 1
                    If i > 0 Then sql.Append(",")
                    sql.Append("@p" & i.ToString())
                Next
                sql.Append(")")

                Using cmd As New SqlCommand(sql.ToString(), conn)
                    For i = 0 To subset.Count - 1
                        cmd.Parameters.AddWithValue("@p" & i.ToString(), subset(i))
                    Next
                    Using rdr = cmd.ExecuteReader()
                        While rdr.Read()
                            Dim img = Convert.ToString(rdr("ImgVidPanel"))
                            Dim num = Convert.ToString(rdr("NumScena"))
                            If Not result.ContainsKey(img) Then result.Add(img, num)
                        End While
                    End Using
                End Using
            Next
        End Using

        Return result
    End Function

    Private Sub BuildGrid()
        ' Prepara dizionario assegnati in memoria (una sola query)
        Dim assignedDict = GetAssignedPanels(ConnString, panelFiles)

        ' Costruzione off-screen con layout sospeso
        scrollPanel.SuspendLayout()
        tblContainer.SuspendLayout()

        tblContainer.Controls.Clear()
        items.Clear()

        Dim singlePanelWidthLocal = previewW + extraPanelWidth + marginPx * 2

        ' Batch creation per evitare freeze
        Dim batchSize As Integer = 50
        For i As Integer = 0 To panelFiles.Count - 1
            Dim filePath = panelFiles(i)
            Dim assigned As Boolean = False
            Dim assignedInfo As String = String.Empty

            If assignedDict.TryGetValue(filePath, assignedInfo) Then
                assigned = True
                assignedInfo = $"Assegnato: {assignedInfo}"
            End If

            Dim panel As New Panel() With {
                .Width = singlePanelWidthLocal,
                .Height = previewH + 80,
                .Margin = New Padding(marginPx),
                .BackColor = Color.White
            }

            Dim pb As New PictureBox() With {
                .Width = previewW,
                .Height = previewH,
                .SizeMode = PictureBoxSizeMode.Zoom,
                .Left = 8,
                .Top = 8,
                .BorderStyle = BorderStyle.FixedSingle,
                .Cursor = Cursors.Hand,
                .BackColor = Color.White,
                .Image = GetPlaceholderImage()
            }
            panel.Controls.Add(pb)

            Dim cb As New CheckBox() With {
                .Left = 8,
                .Top = pb.Bottom + 8,
                .Width = previewW + extraPanelWidth - 16,
                .Text = Path.GetFileName(filePath),
                .Checked = False,
                .Enabled = Not assigned,
                .BackColor = Color.White
            }
            panel.Controls.Add(cb)

            If assigned Then
                Dim lblAssigned As New Label() With {
                    .Left = 8,
                    .Top = cb.Bottom + 6,
                    .Width = previewW + extraPanelWidth - 16,
                    .Height = 18,
                    .Text = assignedInfo,
                    .ForeColor = Color.DarkRed,
                    .Font = New Font(SystemFonts.DefaultFont.FontFamily, 9.0F, FontStyle.Bold),
                    .BackColor = Color.White
                }
                panel.Controls.Add(lblAssigned)
            End If

            AddHandler pb.Click, Sub(s, ev)
                                     If cb.Enabled Then cb.Checked = Not cb.Checked
                                 End Sub

            tblContainer.Controls.Add(panel)

            Dim pi As New PanelItem With {.FilePath = filePath, .Pic = pb, .Cb = cb, .Assigned = assigned, .AssignedInfo = assignedInfo}
            items.Add(pi)

            ' Batch yield per mantenere UI reattiva
            If (i Mod batchSize) = 0 Then
                Application.DoEvents()
            End If
        Next

        ' Imposta altezza del container in base al numero di elementi e colonne
        Dim rowsNeeded = CInt(Math.Ceiling(tblContainer.Controls.Count / CDbl(cols)))
        tblContainer.Height = rowsNeeded * (previewH + 80) + tblContainer.Padding.Vertical

        tblContainer.ResumeLayout(True)
        scrollPanel.ResumeLayout(True)
        tblContainer.PerformLayout()
        scrollPanel.PerformLayout()
    End Sub

    Private Function GetPlaceholderImage() As Image
        Dim bmp As New Bitmap(previewW, previewH)
        Using g = Graphics.FromImage(bmp)
            g.Clear(Color.WhiteSmoke)
            Using pen As New Pen(Color.Gainsboro)
                g.DrawRectangle(pen, 0, 0, previewW - 1, previewH - 1)
            End Using
            Using s As New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
                g.DrawString("Anteprima", PlaceholderFont, Brushes.Gray, New RectangleF(0, 0, previewW, previewH), s)
            End Using
        End Using
        Return bmp
    End Function

    Private Async Function LoadThumbnailsAsync(token As CancellationToken, progress As Action(Of Integer, Integer)) As Task
        If items.Count = 0 Then Return

        Dim cacheDir = Path.Combine(_outDir, ".thumbcache")
        If Not Directory.Exists(cacheDir) Then Directory.CreateDirectory(cacheDir)

        Dim total As Integer = items.Count
        Dim done As Integer = 0
        Dim maxParallel As Integer = Math.Max(2, Math.Min(8, Environment.ProcessorCount))
        Dim sem As New SemaphoreSlim(maxParallel)
        Dim tasks As New List(Of Task)()

        For i As Integer = 0 To items.Count - 1
            token.ThrowIfCancellationRequested()
            Dim it = items(i)

            tasks.Add(Task.Run(Async Function()
                                   Await sem.WaitAsync(token)
                                   Try
                                       token.ThrowIfCancellationRequested()
                                       Dim bmp As Image = Nothing
                                       Try
                                           Dim thumbPath = GetThumbnailPath(cacheDir, it.FilePath)
                                           If File.Exists(thumbPath) Then
                                               Using fs As New FileStream(thumbPath, FileMode.Open, FileAccess.Read, FileShare.Read)
                                                   bmp = Image.FromStream(fs)
                                                   bmp = New Bitmap(bmp)
                                               End Using
                                           Else
                                               Using fs As New FileStream(it.FilePath, FileMode.Open, FileAccess.Read, FileShare.Read)
                                                   Using tmp = Image.FromStream(fs)
                                                       Dim thumb = New Bitmap(previewW, previewH)
                                                       Using g = Graphics.FromImage(thumb)
                                                           g.Clear(Color.Black)
                                                           g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                                                           g.DrawImage(tmp, 0, 0, previewW, previewH)
                                                       End Using
                                                       bmp = thumb
                                                       Try
                                                           Dim saveBmp = New Bitmap(bmp)
                                                           saveBmp.Save(thumbPath, System.Drawing.Imaging.ImageFormat.Png)
                                                           saveBmp.Dispose()
                                                       Catch
                                                       End Try
                                                   End Using
                                               End Using
                                           End If
                                       Catch
                                           bmp = Nothing
                                       End Try

                                       If bmp IsNot Nothing Then
                                           If it.Pic IsNot Nothing AndAlso Not it.Pic.IsDisposed Then
                                               If it.Pic.InvokeRequired Then
                                                   it.Pic.BeginInvoke(Sub()
                                                                          If it.Pic.Image IsNot Nothing Then
                                                                              Try
                                                                                  Dim old = it.Pic.Image
                                                                                  it.Pic.Image = Nothing
                                                                                  old.Dispose()
                                                                              Catch
                                                                              End Try
                                                                          End If
                                                                          it.Pic.Image = bmp
                                                                      End Sub)
                                               Else
                                                   If it.Pic.Image IsNot Nothing Then
                                                       Try
                                                           Dim old = it.Pic.Image
                                                           it.Pic.Image = Nothing
                                                           old.Dispose()
                                                       Catch
                                                       End Try
                                                   End If
                                                   it.Pic.Image = bmp
                                               End If
                                           Else
                                               bmp.Dispose()
                                           End If
                                       End If
                                   Finally
                                       Interlocked.Increment(done)
                                       If progress IsNot Nothing Then progress(done, total)
                                       sem.Release()
                                   End Try
                               End Function, token))
        Next

        Await Task.WhenAll(tasks)
    End Function

    Private Function GetThumbnailPath(cacheDir As String, originalPath As String) As String
        Using sha As SHA256 = SHA256.Create()
            Dim bytes = Encoding.UTF8.GetBytes(originalPath.ToLowerInvariant())
            Dim hash = sha.ComputeHash(bytes)
            Dim sb As New StringBuilder()
            For Each b In hash
                sb.Append(b.ToString("x2"))
            Next
            Return Path.Combine(cacheDir, sb.ToString() & ".png")
        End Using
    End Function

    Private Async Sub BtnRefresh_Click(sender As Object, e As EventArgs)
        If cts IsNot Nothing Then cts.Cancel()
        cts = New CancellationTokenSource()
        Dim token = cts.Token

        Try
            Me.UseWaitCursor = True
            ToggleTopButtons(False)

            overlay.Visible = True
            overlay.BringToFront()
            overlayLabel.Text = "Aggiornamento Panels ..."
            scrollPanel.Visible = False

            LoadScenes()

            overlayLabel.Text = "Aggiornamento lista panel ..."
            Await LoadPanelFilesAsync(token)

            overlayLabel.Text = "Ricostruzione griglia ..."
            BuildGrid()

            scrollPanel.Visible = True

            overlayLabel.Text = "Caricamento anteprime ..."
            Await LoadThumbnailsAsync(token,
                Sub(done, total)
                    UpdateProgressLabels($"Caricamento anteprime {done}/{total}")
                End Sub)

            overlayLabel.Text = $"Caricamento anteprime {items.Count}/{items.Count}"
            Await Task.Delay(300)
            overlay.Visible = False
        Catch ex As OperationCanceledException
        Catch ex As Exception
            MDIMessageBox.Show("Errore aggiornamento panel: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.UseWaitCursor = False
            ToggleTopButtons(True)
        End Try
    End Sub

    Private Sub BtnSelectAll_Click(sender As Object, e As EventArgs)
        For Each it In items
            If Not it.Assigned Then it.Cb.Checked = True
        Next
    End Sub

    Private Sub BtnClearAll_Click(sender As Object, e As EventArgs)
        For Each it In items
            If Not it.Assigned Then it.Cb.Checked = False
        Next
    End Sub

    Private Sub BtnSave_Click(sender As Object, e As EventArgs)
        If cmbScene.SelectedItem Is Nothing Then
            MDIMessageBox.Show("Seleziona una scena.", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim selectedNumScena As String = Convert.ToString(cmbScene.SelectedValue)

        Dim selectedPaths As New List(Of String)
        For Each it In items
            If Not it.Assigned AndAlso it.Cb.Checked Then
                selectedPaths.Add(it.FilePath)
            End If
        Next

        If selectedPaths.Count = 0 Then
            MDIMessageBox.Show("Seleziona almeno un panel non assegnato.", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim episodioId As String = String.Empty
        Try
            Using connE As New SqlConnection(ConnString)
                Using cmdE As New SqlCommand("SELECT TOP 1 EpisodioId
                                              FROM Mov_StoryboardScene
                                              WHERE StoryboardId = @sb AND NumScena = @num", connE)
                    cmdE.Parameters.AddWithValue("@sb", _storyboardId)
                    cmdE.Parameters.AddWithValue("@num", selectedNumScena)
                    connE.Open()
                    Dim resE = cmdE.ExecuteScalar()
                    If resE IsNot Nothing AndAlso Not Convert.IsDBNull(resE) Then
                        episodioId = Convert.ToString(resE)
                    End If
                End Using
            End Using
        Catch ex As Exception
            MDIMessageBox.Show("Impossibile recuperare EpisodioId: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try

        If String.IsNullOrWhiteSpace(episodioId) Then
            MDIMessageBox.Show("EpisodioId non disponibile per la scena selezionata.", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Try
            Using conn As New SqlConnection(ConnString)
                conn.Open()
                Using tran = conn.BeginTransaction()
                    Dim currentCount As Integer = 0
                    Using cntCmd As New SqlCommand("SELECT COUNT(1)
                                                    FROM Mov_StoryboardScenePanel
                                                    WHERE NumScena = @num", conn, tran)
                        cntCmd.Parameters.AddWithValue("@num", selectedNumScena)
                        currentCount = Convert.ToInt32(cntCmd.ExecuteScalar())
                    End Using

                    Dim inserted As Integer = 0
                    Dim skipped As Integer = 0

                    For Each originalPath In selectedPaths
                        Using chkCmd As New SqlCommand("SELECT COUNT(1)
                                                        FROM Mov_StoryboardScenePanel
                                                        WHERE ImgVidPanel = @img", conn, tran)
                            chkCmd.Parameters.AddWithValue("@img", originalPath)
                            Dim existsCount = Convert.ToInt32(chkCmd.ExecuteScalar())
                            If existsCount > 0 Then
                                skipped += 1
                                Continue For
                            End If
                        End Using

                        Dim seq As Integer = currentCount + inserted + 1
                        Dim panelSeqPadded As String = seq.ToString().PadLeft(4, "0"c)
                        Dim descr As String = panelSeqPadded

                        Using insCmd As New SqlCommand("INSERT INTO Mov_StoryboardScenePanel (Descrizione, NumScena, ImgVidPanel)
                                                        VALUES (@descr, @num, @img)", conn, tran)
                            insCmd.Parameters.AddWithValue("@descr", descr)
                            insCmd.Parameters.AddWithValue("@num", selectedNumScena)
                            insCmd.Parameters.AddWithValue("@img", originalPath)
                            insCmd.ExecuteNonQuery()
                        End Using

                        inserted += 1
                    Next

                    tran.Commit()

                    For Each it In items
                        If Not it.Assigned AndAlso it.Cb.Checked AndAlso selectedPaths.Contains(it.FilePath) Then
                            it.Assigned = True
                            it.Cb.Checked = False
                            it.Cb.Enabled = False

                            Dim lblAssigned =
                                it.Pic.Parent.Controls.OfType(Of Label)().
                                    FirstOrDefault(Function(l) l.Text.StartsWith("Assegnato:", StringComparison.OrdinalIgnoreCase))

                            Dim labelText As String = $"Assegnato: {selectedNumScena}"
                            If lblAssigned Is Nothing Then
                                Dim newLbl As New Label() With {
                                    .Left = 8,
                                    .Top = it.Cb.Bottom + 6,
                                    .Width = previewW + extraPanelWidth - 16,
                                    .Height = 18,
                                    .Text = labelText,
                                    .ForeColor = Color.DarkRed,
                                    .Font = New Font(SystemFonts.DefaultFont.FontFamily, 9.0F, FontStyle.Bold)
                                }
                                it.Pic.Parent.Controls.Add(newLbl)
                            Else
                                lblAssigned.Text = labelText
                            End If
                        End If
                    Next
                End Using
            End Using
        Catch ex As Exception
            MDIMessageBox.Show("Errore salvataggio assegnazioni: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub UpdateProgressLabels(text As String)
        '   If lblProgress IsNot Nothing AndAlso Not lblProgress.IsDisposed Then
        '       If lblProgress.InvokeRequired Then
        '           lblProgress.BeginInvoke(Sub() lblProgress.Text = text)
        '       Else
        '           lblProgress.Text = text
        '       End If
        '   End If
        '   If overlayLabel IsNot Nothing AndAlso Not overlayLabel.IsDisposed Then
        '       If overlayLabel.InvokeRequired Then
        '          overlayLabel.BeginInvoke(Sub() overlayLabel.Text = text)
        '       Else
        '           overlayLabel.Text = text
        ' End If
        'End If
    End Sub

End Class
