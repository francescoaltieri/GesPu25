Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.IO
Imports System.Linq
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
    Private tblContainer As TableLayoutPanel
    Private scrollPanel As Panel
    Private cmbScene As ComboBox
    Private btnSave As Button
    Private btnClose As Button
    Private btnSelectAll As Button
    Private btnClearAll As Button
    Private lblInfo As Label

    ' Layout constants (preview ridotte, pannello leggermente più largo)
    Private ReadOnly previewW As Integer = 200
    Private ReadOnly previewH As Integer = 130
    Private ReadOnly extraPanelWidth As Integer = 40 ' <-- larghezza aggiuntiva per il pannello
    Private ReadOnly cols As Integer = 4
    Private ReadOnly marginPx As Integer = 8
    Private ReadOnly leftPad As Integer = 12
    Private ReadOnly topPad As Integer = 12

    Private panelFiles As New List(Of String)()
    Private cts As CancellationTokenSource = Nothing

    ' Classe per item
    Private Class PanelItem
        Public Property FilePath As String
        Public Property Pic As PictureBox
        Public Property Cb As CheckBox
        Public Property Assigned As Boolean
        Public Property AssignedInfo As String
    End Class

    Private items As New List(Of PanelItem)()

    ' --- Costruttore ---
    Public Sub New(storyboardId As String, storyboardDesc As String, outDir As String)
        _storyboardId = storyboardId
        _storyboardDesc = storyboardDesc
        _outDir = outDir

        InitializeComponent()

        ' Ripristina posizione/size del form all'apertura
        AddHandler Me.Load, Sub(s, e)
                                Try
                                    RipristinaPosizioneForm(Me)
                                Catch
                                    ' Ignora errori se la funzione non è disponibile o fallisce
                                End Try
                            End Sub

        AddHandler Me.Shown, AddressOf AssegnaSceneForm_Shown
        AddHandler Me.FormClosing, AddressOf AssegnaSceneForm_FormClosing
    End Sub

    ' --- InitializeComponent (crea UI dinamicamente) ---
    Private Sub InitializeComponent()
        Me.StartPosition = FormStartPosition.CenterParent

        ' Calcola larghezza pannello singolo e dimensione form
        Dim singlePanelWidth = previewW + extraPanelWidth + marginPx * 2
        Dim formWidth As Integer = Math.Max(820, (singlePanelWidth) * cols + 220)
        Dim formHeight As Integer = Math.Max(560, (previewH + marginPx) * 3 + 260)
        Me.Size = New Size(formWidth, formHeight)
        Me.MinimumSize = New Size(760, 520)
        Me.Text = $"Assegna panel - Storyboard {_storyboardId}"

        Dim top As Integer = topPad

        Dim lblId As New Label() With {
            .Left = leftPad,
            .Top = top,
            .Width = 120,
            .Text = "Id storyboard:"
        }
        Me.Controls.Add(lblId)

        Dim lblIdVal As New Label() With {
            .Left = lblId.Right + 8,
            .Top = top,
            .Width = 360,
            .Text = _storyboardId
        }
        Me.Controls.Add(lblIdVal)

        top += 26

        Dim lblDesc As New Label() With {
            .Left = leftPad,
            .Top = top,
            .Width = 120,
            .Text = "Descrizione:"
        }
        Me.Controls.Add(lblDesc)

        Dim lblDescVal As New Label() With {
            .Left = lblDesc.Right + 8,
            .Top = top,
            .Width = 360,
            .Text = _storyboardDesc
        }
        Me.Controls.Add(lblDescVal)

        top += 30

        Dim lblScene As New Label() With {
            .Left = leftPad,
            .Top = top,
            .Width = 120,
            .Text = "Seleziona scena:"
        }
        Me.Controls.Add(lblScene)

        cmbScene = New ComboBox() With {
            .Left = lblScene.Right + 8,
            .Top = top - 2,
            .Width = 260,
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        Me.Controls.Add(cmbScene)

        btnSelectAll = New Button() With {
            .Left = cmbScene.Right + 12,
            .Top = top - 2,
            .Width = 120,
            .Text = "Seleziona tutti"
        }
        AddHandler btnSelectAll.Click, AddressOf BtnSelectAll_Click
        Me.Controls.Add(btnSelectAll)

        btnClearAll = New Button() With {
            .Left = btnSelectAll.Right + 8,
            .Top = top - 2,
            .Width = 120,
            .Text = "Deseleziona tutti"
        }
        AddHandler btnClearAll.Click, AddressOf BtnClearAll_Click
        Me.Controls.Add(btnClearAll)

        top += 36

        ' Scroll panel e TableLayoutPanel per la griglia
        Dim gridWidth = Math.Max((cols * singlePanelWidth), Me.ClientSize.Width - 40)
        Dim rowsVisible As Integer = 3
        Dim gridHeight = rowsVisible * (previewH + 80)

        scrollPanel = New Panel() With {
            .Left = leftPad,
            .Top = top,
            .Width = gridWidth,
            .Height = gridHeight,
            .AutoScroll = True,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        }
        Me.Controls.Add(scrollPanel)

        tblContainer = New TableLayoutPanel() With {
            .ColumnCount = cols,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .Dock = DockStyle.Top,
            .Padding = New Padding(marginPx)
        }
        For i As Integer = 0 To cols - 1
            tblContainer.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, singlePanelWidth))
        Next
        scrollPanel.Controls.Add(tblContainer)

        lblInfo = New Label() With {
            .Left = leftPad,
            .Top = scrollPanel.Bottom + 8,
            .Width = 520,
            .Height = 40,
            .Text = "Seleziona i panel da assegnare alla scena scelta. I panel già assegnati non sono selezionabili.",
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
        }
        Me.Controls.Add(lblInfo)

        ' Pulsanti sotto il pannello, centrati
        btnSave = New Button() With {
            .Text = "Salva",
            .Width = 120,
            .Height = 36,
            .Anchor = AnchorStyles.Bottom
        }
        AddHandler btnSave.Click, AddressOf BtnSave_Click
        Me.Controls.Add(btnSave)

        btnClose = New Button() With {
            .Text = "Chiudi",
            .Width = 120,
            .Height = 36,
            .Anchor = AnchorStyles.Bottom
        }
        AddHandler btnClose.Click, Sub(s, e) Me.Close()
        Me.Controls.Add(btnClose)

        ' Resize handler per posizionare i pulsanti sotto il pannello
        AddHandler Me.Resize, Sub(s, e)
                                  Dim newGridWidth = Math.Max((cols * singlePanelWidth), Me.ClientSize.Width - 40)
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
                              End Sub
    End Sub

    ' --- Evento Shown: avvia caricamento asincrono ---
    Private Async Sub AssegnaSceneForm_Shown(sender As Object, e As EventArgs)
        cts = New CancellationTokenSource()
        Dim token = cts.Token

        Try
            Me.UseWaitCursor = True
            btnSave.Enabled = False
            btnSelectAll.Enabled = False
            btnClearAll.Enabled = False
            cmbScene.Enabled = False

            LoadScenes()
            Await LoadPanelFilesAsync(token)
            BuildGrid()
            Await LoadThumbnailsAsync(token)

        Catch ex As OperationCanceledException
            ' caricamento annullato
        Catch ex As Exception
            MDIMessageBox.Show("Errore caricamento preview: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.UseWaitCursor = False
            btnSave.Enabled = True
            btnSelectAll.Enabled = True
            btnClearAll.Enabled = True
            cmbScene.Enabled = True
            If cts IsNot Nothing Then
                cts.Dispose()
                cts = Nothing
            End If
        End Try
    End Sub

    Private Sub AssegnaSceneForm_FormClosing(sender As Object, e As FormClosingEventArgs)
        ' Annulla eventuali operazioni asincrone
        If cts IsNot Nothing Then
            cts.Cancel()
        End If

        ' Salva posizione/size del form alla chiusura
        Try
            SalvaPosizioneForm(Me)
        Catch
            ' Ignora errori se la funzione non è disponibile o fallisce
        End Try
    End Sub

    ' --- Carica le scene dallo storyboard ---
    Private Sub LoadScenes()
        Try
            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand("SELECT IdProgScena, NumScena, Descrizione FROM Mov_StoryboardScene WHERE StoryboardId = @id ORDER BY NumScena", conn)
                    cmd.Parameters.AddWithValue("@id", _storyboardId)
                    Dim dt As New DataTable()
                    Using da As New SqlDataAdapter(cmd)
                        da.Fill(dt)
                    End Using

                    If dt.Rows.Count = 0 Then
                        cmbScene.DataSource = Nothing
                        MDIMessageBox.Show("Nessuna scena trovata per questo storyboard.", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Information)
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

    ' --- Carica lista file PNG (asincrono) ---
    Private Async Function LoadPanelFilesAsync(token As CancellationToken) As Task
        panelFiles.Clear()
        items.Clear()

        Await Task.Run(Sub()
                           If Not IO.Directory.Exists(_outDir) Then
                               Throw New IO.DirectoryNotFoundException("Cartella panel non trovata: " & _outDir)
                           End If

                           Dim files = IO.Directory.EnumerateFiles(_outDir, "*.png", SearchOption.TopDirectoryOnly).OrderBy(Function(f) f).ToList()
                           SyncLock panelFiles
                               panelFiles.AddRange(files)
                           End SyncLock
                       End Sub, token)
    End Function

    ' --- Costruisce la griglia (senza immagini) ---
    Private Sub BuildGrid()
        tblContainer.Controls.Clear()
        tblContainer.RowCount = 0
        items.Clear()

        Dim colIndex As Integer = 0
        Dim rowIndex As Integer = -1
        Dim singlePanelWidth = previewW + extraPanelWidth + marginPx * 2

        For i As Integer = 0 To panelFiles.Count - 1
            If colIndex = 0 Then
                tblContainer.RowCount += 1
                tblContainer.RowStyles.Add(New RowStyle(SizeType.Absolute, previewH + 80))
                rowIndex += 1
            End If

            Dim filePath = panelFiles(i)
            Dim assigned As Boolean = False
            Dim assignedInfo As String = String.Empty

            ' Verifica se già assegnato (legge Descrizione)
            Try
                Using conn As New SqlConnection(ConnString)
                    Using cmd As New SqlCommand("SELECT Descrizione FROM Mov_StoryboardScenePanel WHERE ImgVidPanel = @img", conn)
                        cmd.Parameters.AddWithValue("@img", filePath)
                        conn.Open()
                        Dim res = cmd.ExecuteScalar()
                        If res IsNot Nothing AndAlso Not Convert.IsDBNull(res) Then
                            assigned = True
                            assignedInfo = $"Assegnato: {Convert.ToString(res)}"
                        End If
                    End Using
                End Using
            Catch
                assigned = False
            End Try

            Dim panel As New Panel() With {
                .Width = singlePanelWidth,
                .Height = previewH + 80,
                .Margin = New Padding(marginPx)
            }

            Dim pb As New PictureBox() With {
                .Width = previewW,
                .Height = previewH,
                .SizeMode = PictureBoxSizeMode.Zoom,
                .Left = 8,
                .Top = 8,
                .BorderStyle = BorderStyle.FixedSingle,
                .Cursor = Cursors.Hand,
                .Image = Nothing
            }
            panel.Controls.Add(pb)

            Dim cb As New CheckBox() With {
                .Left = 8,
                .Top = pb.Bottom + 8,
                .Width = previewW + extraPanelWidth - 16, ' checkbox più largo per adattarsi al pannello
                .Text = IO.Path.GetFileName(filePath),
                .Checked = False,
                .Enabled = Not assigned
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
                    .Font = New Font(SystemFonts.DefaultFont.FontFamily, 9.0F, FontStyle.Bold)
                }
                panel.Controls.Add(lblAssigned)
            End If

            AddHandler pb.Click, Sub(s, ev)
                                     If cb.Enabled Then cb.Checked = Not cb.Checked
                                 End Sub

            tblContainer.Controls.Add(panel, colIndex, rowIndex)

            Dim pi As New PanelItem With {
                .FilePath = filePath,
                .Pic = pb,
                .Cb = cb,
                .Assigned = assigned,
                .AssignedInfo = assignedInfo
            }
            items.Add(pi)

            colIndex += 1
            If colIndex >= cols Then colIndex = 0
        Next
    End Sub

    ' --- Carica thumbnails in parallelo ---
    Private Async Function LoadThumbnailsAsync(token As CancellationToken) As Task
        If items.Count = 0 Then Return

        Dim maxParallel As Integer = Math.Max(2, Environment.ProcessorCount)
        Dim sem As New SemaphoreSlim(maxParallel)
        Dim tasks As New List(Of Task)()

        For i As Integer = 0 To items.Count - 1
            token.ThrowIfCancellationRequested()
            Dim idx = i
            Dim it = items(idx)

            tasks.Add(Task.Run(Async Function()
                                   Await sem.WaitAsync(token)
                                   Try
                                       token.ThrowIfCancellationRequested()
                                       Dim bmp As Image = Nothing
                                       Try
                                           Using fs As New FileStream(it.FilePath, FileMode.Open, FileAccess.Read, FileShare.Read)
                                               Dim tmp = Image.FromStream(fs)
                                               Dim thumb = New Bitmap(previewW, previewH)
                                               Using g = Graphics.FromImage(thumb)
                                                   g.Clear(Color.Black)
                                                   g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                                                   g.DrawImage(tmp, 0, 0, previewW, previewH)
                                               End Using
                                               tmp.Dispose()
                                               bmp = thumb
                                           End Using
                                       Catch
                                           bmp = Nothing
                                       End Try

                                       If token.IsCancellationRequested Then
                                           If bmp IsNot Nothing Then bmp.Dispose()
                                           Return
                                       End If

                                       If bmp IsNot Nothing Then
                                           If it.Pic IsNot Nothing AndAlso Not it.Pic.IsDisposed Then
                                               If it.Pic.InvokeRequired Then
                                                   it.Pic.Invoke(Sub() it.Pic.Image = bmp)
                                               Else
                                                   it.Pic.Image = bmp
                                               End If
                                           Else
                                               bmp.Dispose()
                                           End If
                                       End If
                                   Finally
                                       sem.Release()
                                   End Try
                               End Function, token))
        Next

        Await Task.WhenAll(tasks)
    End Function

    ' --- Seleziona tutti i checkbox abilitati ---
    Private Sub BtnSelectAll_Click(sender As Object, e As EventArgs)
        For Each it In items
            If Not it.Assigned Then it.Cb.Checked = True
        Next
    End Sub

    ' --- Deseleziona tutti i checkbox abilitati ---
    Private Sub BtnClearAll_Click(sender As Object, e As EventArgs)
        For Each it In items
            If Not it.Assigned Then it.Cb.Checked = False
        Next
    End Sub

    ' --- Salvataggio selezioni in DB ---
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

        Try
            Using conn As New SqlConnection(ConnString)
                conn.Open()
                Using tran = conn.BeginTransaction()
                    ' Conteggio corrente per la scena (numero di panel già presenti per quella scena)
                    Dim currentCount As Integer = 0
                    Using cntCmd As New SqlCommand("SELECT COUNT(1) FROM Mov_StoryboardScenePanel WHERE NumScena = @num", conn, tran)
                        cntCmd.Parameters.AddWithValue("@num", selectedNumScena)
                        currentCount = Convert.ToInt32(cntCmd.ExecuteScalar())
                    End Using

                    Dim seq As Integer = currentCount + 1
                    Dim inserted As Integer = 0
                    Dim skipped As Integer = 0
                    Dim insertedPaths As New List(Of String)

                    For Each p In selectedPaths
                        ' Controllo unicità su ImgVidPanel
                        Using chkCmd As New SqlCommand("SELECT COUNT(1) FROM Mov_StoryboardScenePanel WHERE ImgVidPanel = @img", conn, tran)
                            chkCmd.Parameters.AddWithValue("@img", p)
                            Dim existsCount = Convert.ToInt32(chkCmd.ExecuteScalar())
                            If existsCount > 0 Then
                                skipped += 1
                                Continue For
                            End If
                        End Using

                        ' costruzione descrizione: solo progressivo padded a 4 cifre
                        Dim panelSeqPadded As String = seq.ToString().PadLeft(4, "0"c)
                        Dim descr As String = panelSeqPadded

                        Using insCmd As New SqlCommand("INSERT INTO Mov_StoryboardScenePanel (Descrizione, NumScena, ImgVidPanel) VALUES (@descr, @num, @img)", conn, tran)
                            insCmd.Parameters.AddWithValue("@descr", descr)
                            insCmd.Parameters.AddWithValue("@num", selectedNumScena)
                            insCmd.Parameters.AddWithValue("@img", p)
                            insCmd.ExecuteNonQuery()
                        End Using

                        inserted += 1
                        insertedPaths.Add(p)
                        seq += 1
                    Next

                    tran.Commit()

                    ' Aggiorna stato in memoria e UI: mostra "Assegnato: 0001" (solo progressivo)
                    For Each it In items
                        If insertedPaths.Contains(it.FilePath) Then
                            it.Assigned = True
                            it.Cb.Checked = False
                            it.Cb.Enabled = False

                            Dim idxInInserted = insertedPaths.IndexOf(it.FilePath)
                            If idxInInserted >= 0 Then
                                Dim seqForLabel As Integer = currentCount + 1 + idxInInserted
                                Dim labelDescr As String = seqForLabel.ToString().PadLeft(4, "0"c)

                                Dim lblAssigned = it.Pic.Parent.Controls.OfType(Of Label)().FirstOrDefault(Function(l) l.Text.StartsWith("Assegnato:"))
                                If lblAssigned Is Nothing Then
                                    Dim newLbl As New Label() With {
                                        .Left = 8,
                                        .Top = it.Cb.Bottom + 6,
                                        .Width = previewW + extraPanelWidth - 16,
                                        .Height = 18,
                                        .Text = $"Assegnato: {labelDescr}",
                                        .ForeColor = Color.DarkRed,
                                        .Font = New Font(SystemFonts.DefaultFont.FontFamily, 9.0F, FontStyle.Bold)
                                    }
                                    it.Pic.Parent.Controls.Add(newLbl)
                                Else
                                    lblAssigned.Text = $"Assegnato: {labelDescr}"
                                End If
                            End If
                        End If
                    Next

                    MDIMessageBox.Show($"Assegnazione completata. Inseriti: {inserted}. Saltati (già assegnati): {skipped}.", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Using
            End Using
        Catch ex As Exception
            MDIMessageBox.Show("Errore salvataggio: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class
