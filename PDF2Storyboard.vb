Imports System.IO
Imports System.Text.RegularExpressions
Imports Microsoft.Data.SqlClient
Imports PdfiumViewer
Imports UglyToad.PdfPig
Imports Emgu.CV
Imports Emgu.CV.CvEnum
Imports Emgu.CV.Structure
Imports Emgu.CV.Util
Imports System.Drawing
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Runtime.InteropServices
Imports System.Diagnostics
Imports Emgu.CV.Ocl

Public Class PDF2Storyboard

    ' Configurabili
    Private Const DefaultRenderDpi As Single = 600.0F
    Private ReadOnly SameRowTolerancePx As Integer = 40

    Private exportedFiles As New List(Of String)()
    Private currentIndex As Integer = -1
    Private cts As CancellationTokenSource = Nothing

    ' Editing manuale / stato
    Private originalImage As Image = Nothing          ' immagine originale in memoria per undo
    Private originalFilePath As String = String.Empty ' percorso file corrente mostrato
    Private isEditing As Boolean = False
    Private selectionRect As Rectangle = Rectangle.Empty   ' coordinate in pixel immagine
    Private isDragging As Boolean = False
    Private dragStart As Point = Point.Empty
    Private hasUnsavedChanges As Boolean = False

    ' Pen/Brush overlay (dispose on form close)
    Private ReadOnly overlayPen As New Pen(Color.FromArgb(200, Color.Red), 2)
    Private ReadOnly overlayBrush As New SolidBrush(Color.FromArgb(60, Color.Red))

    Private lastCropRect As Rectangle = Rectangle.Empty

    Private ctsAnimatic As CancellationTokenSource = Nothing
    Private animaticInCorso As Boolean = False
    Private outMp4 As String = String.Empty


    Private Sub PDF2Storyboard_Load(sender As Object, e As EventArgs) Handles Me.Load
        RipristinaPosizioneForm(Me)

        ' popola la combo con Id + Descrizione
        LoadStoryboardCombo()

        If PicPanel IsNot Nothing Then
            PicPanel.SizeMode = PictureBoxSizeMode.Zoom
            PicPanel.Cursor = Cursors.Cross
            AddHandler PicPanel.MouseDown, AddressOf PicPanel_MouseDown
            AddHandler PicPanel.MouseMove, AddressOf PicPanel_MouseMove
            AddHandler PicPanel.MouseUp, AddressOf PicPanel_MouseUp
            AddHandler PicPanel.Paint, AddressOf PicPanel_Paint

            Try
                PicPanel.GetType().GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.NonPublic).SetValue(PicPanel, True, Nothing)
            Catch
            End Try
        End If
    End Sub

    Private Sub PDF2Storyboard_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        ' chiedi di salvare se ci sono modifiche non salvate
        If Not ConfirmDiscardChanges() Then
            e.Cancel = True
            Return
        End If

        Try
            If cts IsNot Nothing Then
                cts.Cancel()
                cts.Dispose()
            End If
        Catch
        End Try

        ' dispose risorse immagine/overlay
        Try
            If originalImage IsNot Nothing Then originalImage.Dispose()
            If PicPanel IsNot Nothing AndAlso PicPanel.Image IsNot Nothing Then PicPanel.Image.Dispose()
            If overlayPen IsNot Nothing Then overlayPen.Dispose()
            If overlayBrush IsNot Nothing Then overlayBrush.Dispose()
        Catch
        End Try

        SalvaPosizioneForm(Me)
    End Sub

    Private Sub BtnAnnulla_Click(sender As Object, e As EventArgs) Handles BtnChiudi.Click
        Me.Close()
    End Sub

    Private Sub LoadStoryboardCombo()
        Try
            ' Carica IdStoryboard + Descrizione dalla tabella Mov_Storyboard
            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand("SELECT IdStoryboard, Descrizione FROM Mov_Storyboard ORDER BY Descrizione", conn)
                    conn.Open()
                    Using rdr = cmd.ExecuteReader()
                        Dim dt As New DataTable()
                        dt.Load(rdr)
                        ' Aggiungi colonna Display combinata
                        If Not dt.Columns.Contains("Display") Then
                            dt.Columns.Add("Display", GetType(String))
                        End If
                        For Each dr As DataRow In dt.Rows
                            Dim id = Convert.ToString(dr("IdStoryboard"))
                            Dim desc = Convert.ToString(dr("Descrizione"))
                            dr("Display") = $"{id} - {desc}"
                        Next
                        ComboStoryboard.DataSource = dt
                        ComboStoryboard.DisplayMember = "Display"
                        ComboStoryboard.ValueMember = "IdStoryboard"
                        ComboStoryboard.SelectedIndex = -1
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MDIMessageBox.Show("Errore caricamento Mov_Storyboard: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
            ComboStoryboard.DataSource = Nothing
        End Try
    End Sub

    Private Async Sub BtnAcquisisciPDF_Click(sender As Object, e As EventArgs) Handles BtnAcquisisciPDF.Click
        Try
            ' Prova a leggere il percorso PDF dalla tabella Mov_Storyboard usando l'Id selezionato nella combo
            Dim pdfPathFromDb = GetStoryboardPdfPathFromDb(ComboStoryboard)

            Dim pdfPath As String = String.Empty
            If Not String.IsNullOrWhiteSpace(pdfPathFromDb) AndAlso File.Exists(pdfPathFromDb) Then
                pdfPath = pdfPathFromDb
            Else
                ' fallback: chiedi all'utente di selezionare manualmente il PDF
                pdfPath = SelezionaPdf()
                If String.IsNullOrWhiteSpace(pdfPath) OrElse Not File.Exists(pdfPath) Then Return
            End If

            ' Ottieni la cartella di output basata su Id dello storyboard (crea se mancante)
            Dim outDir = GetStoryboardOutputFolderByIdOnly(Me, ComboStoryboard)
            If String.IsNullOrWhiteSpace(outDir) Then Return

            ' Se nella cartella esistono già PNG, non procedere: richiedi di rimuovere i vecchi panels
            Try
                Dim existing = Directory.EnumerateFiles(outDir, "*.png", SearchOption.TopDirectoryOnly).Any()
                If existing Then
                    MDIMessageBox.Show("La cartella dello storyboard contiene già alcuni panel. Per procedere, elimina prima i vecchi panels dalla cartella selezionata.", Me.MdiParent, MessageBoxButtons.OK)
                    Return
                End If
            Catch ex As Exception
                ' Se non possiamo enumerare i file (permessi, rete) blocchiamo l'operazione per sicurezza
                MDIMessageBox.Show("Impossibile verificare il contenuto della cartella di output: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
                Return
            End Try

            BtnAcquisisciPDF.Enabled = False
            Cursor = Cursors.WaitCursor

            exportedFiles.Clear()
            currentIndex = -1

            cts = New CancellationTokenSource
            Dim token = cts.Token

            Dim files = Await Task.Run(Function() ProcessPdfToPanels(pdfPath, outDir, token), token)

            If files Is Nothing OrElse files.Count = 0 Then
                MDIMessageBox.Show("Nessun pannello trovato/esportato.", MdiParent, MessageBoxButtons.OK)
                Return
            End If

            exportedFiles = files
            currentIndex = 0
            ShowCurrentPanel()
            MDIMessageBox.Show($"Esportazione completata: {exportedFiles.Count} PNG in{Environment.NewLine}{Path.GetDirectoryName(exportedFiles(0))}", MdiParent, MessageBoxButtons.OK)

        Catch ex As OperationCanceledException
            MDIMessageBox.Show("Operazione annullata.", MdiParent, MessageBoxButtons.OK)
        Catch ex As Exception
            MDIMessageBox.Show("Errore: " & ex.Message, MdiParent, MessageBoxButtons.OK)
        Finally
            Try
                BtnAcquisisciPDF.Enabled = True
                Cursor = Cursors.Default
                If cts IsNot Nothing Then cts.Dispose() : cts = Nothing
            Catch
            End Try
        End Try
    End Sub

    Private Sub BtnAssegnaScene_Click(sender As Object, e As EventArgs) Handles BtnAssegnaScene.Click
        ' Verifica storyboard selezionato
        Dim storyboardId As String = String.Empty
        Dim storyboardDesc As String = String.Empty
        Try
            Dim drv = TryCast(ComboStoryboard.SelectedItem, DataRowView)
            If drv IsNot Nothing Then
                storyboardId = Convert.ToString(drv("IdStoryboard"))
                storyboardDesc = Convert.ToString(drv("Descrizione"))
            Else
                storyboardId = If(ComboStoryboard.SelectedValue IsNot Nothing, ComboStoryboard.SelectedValue.ToString(), String.Empty)
                storyboardDesc = ComboStoryboard.Text
            End If
        Catch
            storyboardId = If(ComboStoryboard.SelectedValue IsNot Nothing, ComboStoryboard.SelectedValue.ToString(), String.Empty)
            storyboardDesc = ComboStoryboard.Text
        End Try

        If String.IsNullOrWhiteSpace(storyboardId) Then
            MDIMessageBox.Show("Seleziona uno storyboard valido.", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim outDir = GetStoryboardOutputFolderByIdOnly(Me, ComboStoryboard)
        If String.IsNullOrWhiteSpace(outDir) Then
            MDIMessageBox.Show("Impossibile determinare la cartella dei panel per lo storyboard.", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' Crea il form come MDI child (non usare Using / ShowDialog)
        Dim frm As New AssegnaSceneForm(storyboardId, storyboardDesc, outDir)
        frm.MdiParent = GesPu25
        frm.TopLevel = False
        frm.FormBorderStyle = FormBorderStyle.Sizable
        frm.StartPosition = FormStartPosition.Manual

        ' Prova a ripristinare posizione/size se disponibile
        Try
            RipristinaPosizioneForm(frm)
        Catch
            ' Ignora errori se la funzione non è disponibile
        End Try

        frm.Show()
    End Sub

    Private Sub BtnPrev_Click(sender As Object, e As EventArgs) Handles BtnPrev.Click
        If exportedFiles.Count = 0 Then Return
        If Not ConfirmDiscardChanges(CheckConfermaSalvataggio.Checked) Then Return
        currentIndex = Math.Max(0, currentIndex - 1)
        ShowCurrentPanel()
    End Sub

    Private Sub BtnNext_Click(sender As Object, e As EventArgs) Handles BtnNext.Click
        If exportedFiles.Count = 0 Then Return
        If Not ConfirmDiscardChanges(CheckConfermaSalvataggio.Checked) Then Return
        currentIndex = Math.Min(exportedFiles.Count - 1, currentIndex + 1)
        ShowCurrentPanel()
    End Sub

    Private Sub ShowCurrentPanel()
        If currentIndex < 0 OrElse currentIndex >= exportedFiles.Count Then
            PicPanel.Image = Nothing
            originalImage = Nothing
            originalFilePath = String.Empty
            isEditing = False
            selectionRect = Rectangle.Empty
            hasUnsavedChanges = False
            UpdateSaveButtonState()
            Return
        End If
        Try
            Dim fp = exportedFiles(currentIndex)
            Using fs As New FileStream(fp, FileMode.Open, FileAccess.Read, FileShare.Read)
                Using mem As New MemoryStream()
                    fs.CopyTo(mem)
                    mem.Seek(0, SeekOrigin.Begin)
                    Dim img = Image.FromStream(mem)
                    If originalImage IsNot Nothing Then
                        Try : originalImage.Dispose() : Catch : End Try
                    End If
                    originalImage = New Bitmap(img)
                    originalFilePath = fp
                    If PicPanel.Image IsNot Nothing Then
                        Try : PicPanel.Image.Dispose() : Catch : End Try
                    End If
                    PicPanel.Image = New Bitmap(img)
                End Using
            End Using
            Me.Text = $"Panel {currentIndex + 1}/{exportedFiles.Count} - {Path.GetFileName(originalFilePath)}"
            ' reset editing state
            isEditing = False
            selectionRect = Rectangle.Empty
            hasUnsavedChanges = False
            UpdateSaveButtonState()
            PicPanel.Invalidate()
        Catch ex As Exception
            Debug.WriteLine("ShowCurrentPanel error: " & ex.Message)
            PicPanel.Image = Nothing
            originalImage = Nothing
            originalFilePath = String.Empty
            isEditing = False
            selectionRect = Rectangle.Empty
            hasUnsavedChanges = False
            UpdateSaveButtonState()
        End Try
    End Sub

    ' Aggiorna lo stato del pulsante Salva Modifica
    Private Sub UpdateSaveButtonState()
        Me.BtnSalvaPanel.Enabled = hasUnsavedChanges
    End Sub

    ' --------------------
    ' Pipeline sync
    ' --------------------
    Private Function ProcessPdfToPanels(pdfPath As String, outDir As String, token As CancellationToken) As List(Of String)
        Dim localExported As New List(Of String)()

        Using doc = PdfiumViewer.PdfDocument.Load(pdfPath)
            Dim pageCount = doc.PageCount
            For i As Integer = 0 To pageCount - 1
                token.ThrowIfCancellationRequested()

                Dim res = TryExtractScenePanelFromPdfPig(pdfPath, i)
                Dim sceneLabel As String = res.Scene
                Dim panelLabel As String = res.Panel

                Using bmpPage As Bitmap = RenderPageBitmap(doc, i, DefaultRenderDpi)
                    If bmpPage Is Nothing Then Continue For

                    Dim crops = SegmentaPannelliDaBitmap(bmpPage)
                    If crops.Count = 0 Then
                        Dim fallbackName = If(Not String.IsNullOrWhiteSpace(sceneLabel) AndAlso Not String.IsNullOrWhiteSpace(panelLabel),
                                          $"{sceneLabel}_{panelLabel}.png",
                                          $"Page_{i + 1:00}.png")
                        fallbackName = MakeValidFileName(fallbackName)
                        Dim outFile = Path.Combine(outDir, fallbackName)
                        SafeSaveBitmap(bmpPage, outFile)
                        localExported.Add(outFile)
                        Continue For
                    End If

                    ' Ordina top->bottom then left->right
                    crops.Sort(Function(a, b)
                                   Dim rowA = a.Y
                                   Dim rowB = b.Y
                                   If Math.Abs(rowA - rowB) < SameRowTolerancePx Then
                                       Return a.X.CompareTo(b.X)
                                   End If
                                   Return rowA.CompareTo(rowB)
                               End Function)

                    Dim idx As Integer = 1
                    For Each r In crops
                        token.ThrowIfCancellationRequested()
                        Using panelBmp As Bitmap = bmpPage.Clone(r, bmpPage.PixelFormat)
                            Dim panelRect As Rectangle = FindPanelRectangle(panelBmp)

                            If Not panelRect.IsEmpty AndAlso panelRect.Width > 10 AndAlso panelRect.Height > 10 Then
                                Using finalBmp As Bitmap = panelBmp.Clone(panelRect, panelBmp.PixelFormat)
                                    Dim idxPadded As String = idx.ToString().PadLeft(4, "0"c)
                                    Dim fileName As String
                                    If Not String.IsNullOrWhiteSpace(sceneLabel) AndAlso Not String.IsNullOrWhiteSpace(panelLabel) Then
                                        fileName = $"{sceneLabel}_{panelLabel}_{idxPadded}.png"
                                    Else
                                        fileName = $"Page_{i + 1:00}_{idxPadded}.png"
                                    End If
                                    fileName = MakeValidFileName(fileName)
                                    Dim outFile = Path.Combine(outDir, fileName)
                                    SafeSaveBitmap(finalBmp, outFile)
                                    localExported.Add(outFile)
                                End Using
                            Else
                                ' fallback: salva panel senza ritagli aggiuntivi
                                Dim idxPadded As String = idx.ToString().PadLeft(4, "0"c)
                                Dim fileName As String = $"Page_{i + 1:00}_{idxPadded}.png"
                                fileName = MakeValidFileName(fileName)
                                Dim outFile = Path.Combine(outDir, fileName)
                                SafeSaveBitmap(panelBmp, outFile)
                                localExported.Add(outFile)
                            End If
                        End Using
                        idx += 1
                    Next
                End Using
            Next
        End Using

        Return localExported
    End Function


    ' --------------------
    ' Save helper
    ' --------------------
    Private Sub SafeSaveBitmap(bmp As Bitmap, outPath As String)
        Dim temp = outPath & ".tmp"
        Try
            If File.Exists(temp) Then File.Delete(temp)
            bmp.Save(temp, Imaging.ImageFormat.Png)
            If File.Exists(outPath) Then File.Delete(outPath)
            File.Move(temp, outPath)
        Catch ex As Exception
            Debug.WriteLine("SafeSaveBitmap error: " & ex.Message)
            Try
                If File.Exists(temp) Then File.Delete(temp)
            Catch
            End Try
        End Try
    End Sub

    ' --------------------
    ' Helpers: PDF select / output folder
    ' --------------------
    Private Function SelezionaPdf() As String
        Using ofd As New OpenFileDialog()
            ofd.Filter = "PDF|*.pdf"
            ofd.Title = "Seleziona storyboard PDF"
            ofd.CheckFileExists = True
            ofd.Multiselect = False
            Return If(ofd.ShowDialog(Me) = DialogResult.OK, ofd.FileName, String.Empty)
        End Using
    End Function

    Private Function GetStoryboardOutputFolder(form As Form, comboStoryboard As ComboBox) As String
        Dim basePath As String = String.Empty
        Try
            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand("SELECT TOP 1 Valore FROM Sys_Parametri WHERE Descrizione = @d", conn)
                    cmd.Parameters.AddWithValue("@d", "PercorsoPanelStoryboard")
                    conn.Open()
                    Dim res = cmd.ExecuteScalar()
                    basePath = Convert.ToString(res).Trim()
                End Using
            End Using
        Catch ex As Exception
            MDIMessageBox.Show("Errore lettura Sys_Parametri: " & ex.Message, form.MdiParent, MessageBoxButtons.OK)
            Return String.Empty
        End Try

        If String.IsNullOrWhiteSpace(basePath) Then
            MDIMessageBox.Show("PercorsoPanelStoryboard non configurato.", form.MdiParent, MessageBoxButtons.OK)
            Return String.Empty
        End If

        If comboStoryboard Is Nothing Then
            MDIMessageBox.Show("ComboStoryboard non trovata.", form.MdiParent, MessageBoxButtons.OK)
            Return String.Empty
        End If

        If comboStoryboard.SelectedItem Is Nothing Then
            MDIMessageBox.Show("Seleziona uno storyboard dalla lista.", form.MdiParent, MessageBoxButtons.OK)
            Return String.Empty
        End If

        ' Estrai Id e Descrizione dalla SelectedItem (DataRowView)
        Dim storyboardId As String = String.Empty
        Dim storyboardDesc As String = String.Empty
        Try
            Dim drv = TryCast(comboStoryboard.SelectedItem, DataRowView)
            If drv IsNot Nothing Then
                storyboardId = Convert.ToString(drv("IdStoryboard"))
                storyboardDesc = Convert.ToString(drv("Descrizione"))
            Else
                storyboardId = If(comboStoryboard.SelectedValue IsNot Nothing, comboStoryboard.SelectedValue.ToString(), String.Empty)
                storyboardDesc = comboStoryboard.Text
            End If
        Catch
            storyboardId = If(comboStoryboard.SelectedValue IsNot Nothing, comboStoryboard.SelectedValue.ToString(), String.Empty)
            storyboardDesc = comboStoryboard.Text
        End Try

        If String.IsNullOrWhiteSpace(storyboardId) Then
            MDIMessageBox.Show("Id dello storyboard non disponibile.", form.MdiParent, MessageBoxButtons.OK)
            Return String.Empty
        End If

        ' Per default cartella con Id.
        Dim folderName = $"{storyboardId}"
        folderName = SanitizeFolderName(folderName)

        Dim outDir = GetStoryboardOutputFolderByIdOnly(Me, comboStoryboard)
        Try
            Directory.CreateDirectory(outDir)
        Catch ex As Exception
            MDIMessageBox.Show("Errore creando cartella output: " & ex.Message, form.MdiParent, MessageBoxButtons.OK)
            Return String.Empty
        End Try
        Return outDir
    End Function

    Private Function GetStoryboardOutputFolderByIdOnly(form As Form, comboStoryboard As ComboBox) As String
        Dim basePath As String = String.Empty
        Try
            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand("SELECT TOP 1 Valore FROM Sys_Parametri WHERE Descrizione = @d", conn)
                    cmd.Parameters.AddWithValue("@d", "PercorsoPanelStoryboard")
                    conn.Open()
                    Dim res = cmd.ExecuteScalar()
                    basePath = Convert.ToString(res).Trim()
                End Using
            End Using
        Catch ex As Exception
            MDIMessageBox.Show("Errore lettura Sys_Parametri: " & ex.Message, form.MdiParent, MessageBoxButtons.OK)
            Return String.Empty
        End Try

        If String.IsNullOrWhiteSpace(basePath) Then
            MDIMessageBox.Show("PercorsoPanelStoryboard non configurato.", form.MdiParent, MessageBoxButtons.OK)
            Return String.Empty
        End If

        If comboStoryboard Is Nothing OrElse comboStoryboard.SelectedItem Is Nothing Then
            MDIMessageBox.Show("Seleziona uno storyboard dalla lista.", form.MdiParent, MessageBoxButtons.OK)
            Return String.Empty
        End If

        Dim storyboardId As String = String.Empty
        Try
            Dim drv = TryCast(comboStoryboard.SelectedItem, DataRowView)
            If drv IsNot Nothing Then
                storyboardId = Convert.ToString(drv("IdStoryboard"))
            Else
                storyboardId = If(comboStoryboard.SelectedValue IsNot Nothing, comboStoryboard.SelectedValue.ToString(), String.Empty)
            End If
        Catch
            storyboardId = If(comboStoryboard.SelectedValue IsNot Nothing, comboStoryboard.SelectedValue.ToString(), String.Empty)
        End Try

        If String.IsNullOrWhiteSpace(storyboardId) Then
            MDIMessageBox.Show("Id dello storyboard non disponibile.", form.MdiParent, MessageBoxButtons.OK)
            Return String.Empty
        End If

        ' Uso solo l'Id per trovare/creare la cartella
        Dim folderName = SanitizeFolderName(storyboardId)
        Dim outDir = Path.Combine(basePath, "PanelsStoryboard", folderName)

        Try
            If Not Directory.Exists(outDir) Then
                Directory.CreateDirectory(outDir)
            End If
        Catch ex As Exception
            MDIMessageBox.Show("Errore creando cartella output: " & ex.Message, form.MdiParent, MessageBoxButtons.OK)
            Return String.Empty
        End Try

        Return outDir
    End Function

    Private Function SanitizeFolderName(name As String) As String
        If String.IsNullOrWhiteSpace(name) Then Return "Storyboard"
        For Each c In Path.GetInvalidFileNameChars()
            name = name.Replace(c, "_"c)
        Next
        Return name
    End Function

    Private Function MakeValidFileName(name As String) As String
        For Each c In Path.GetInvalidFileNameChars()
            name = name.Replace(c, "_"c)
        Next
        Return name
    End Function

    ' --------------------
    ' Pdfium render: pagina -> Bitmap
    ' --------------------
    Private Function RenderPageBitmap(doc As PdfiumViewer.PdfDocument, pageIndex As Integer, dpi As Single) As Bitmap
        Dim pageImage As System.Drawing.Image = Nothing
        Dim bmp As Bitmap '= Nothing
        Try
            pageImage = doc.Render(pageIndex, dpi, dpi, PdfRenderFlags.ForPrinting)
            bmp = New Bitmap(pageImage)
            Return bmp
        Finally
            If pageImage IsNot Nothing Then
                Try
                    pageImage.Dispose()
                Catch
                End Try
            End If
        End Try
    End Function


    ' --------------------
    ' PdfPig: estrazione Scene/Panel dal testo pagina
    ' --------------------
    Private Function TryExtractScenePanelFromPdfPig(pdfPath As String, pageIndex As Integer) As (Scene As String, Panel As String)
        Try
            Using pdf = UglyToad.PdfPig.PdfDocument.Open(pdfPath)
                Dim page = pdf.GetPage(pageIndex + 1)
                Dim text As String = ""
                If page IsNot Nothing Then
                    Dim t = Convert.ToString(page.Text)
                    If Not String.IsNullOrWhiteSpace(t) Then text = t
                End If

                If String.IsNullOrWhiteSpace(text) Then Return ("", "")
                Dim m = Regex.Match(text, "\b(?<scene>\d{1,3})\s*[-_.]?\s*(?<panel>[A-Z])\b")
                If m.Success Then
                    Return (Scene:=m.Groups("scene").Value, Panel:=m.Groups("panel").Value)
                End If
            End Using
        Catch ex As Exception
            Debug.WriteLine("TryExtractScenePanelFromPdfPig error: " & ex.Message)
        End Try
        Return (Scene:="", Panel:="")
    End Function

    ' --------------------
    ' Segmentazione: griglia 3x2
    ' --------------------
    Private Function SegmentaPannelliDaBitmap(bmp As Bitmap) As List(Of Rectangle)
        Return ForceGrid3x2(bmp)
    End Function

    Private Function ForceGrid3x2(bmp As Bitmap) As List(Of Rectangle)
        Dim w = bmp.Width
        Dim h = bmp.Height

        ' Padding proporzionale (circa 0.5% del lato) con minimi
        Dim padX = Math.Max(8, w \ 200)
        Dim padY = Math.Max(8, h \ 200)

        Dim colW = w \ 3
        Dim rowH = h \ 2

        Dim rects As New List(Of Rectangle)()
        For r = 0 To 1
            For c = 0 To 2
                Dim sx = c * colW
                Dim sy = r * rowH
                Dim ex = If(c = 2, w, (c + 1) * colW)
                Dim ey = If(r = 1, h, (r + 1) * rowH)

                Dim x = Math.Max(0, sx + padX)
                Dim y = Math.Max(0, sy + padY)
                Dim rw = Math.Min(w - x, (ex - sx) - 2 * padX)
                Dim rh = Math.Min(h - y, (ey - sy) - 2 * padY)

                rects.Add(New Rectangle(x, y, rw, rh))
            Next
        Next
        Return rects
    End Function

    ' --------------------
    ' Trim adattivo del bianco per ogni panel (lasciato se ti serve)
    ' --------------------
    Private Function ComputeAdaptiveWhiteThreshold(bmp As Bitmap, Optional lowPct As Double = 0.92, Optional highPct As Double = 0.98) As Byte
        Dim rectAll As New Rectangle(0, 0, bmp.Width, bmp.Height)
        Dim data = bmp.LockBits(rectAll, Imaging.ImageLockMode.ReadOnly, bmp.PixelFormat)
        Try
            Dim bpp As Integer
            Select Case data.PixelFormat
                Case Imaging.PixelFormat.Format24bppRgb : bpp = 3
                Case Imaging.PixelFormat.Format32bppArgb, Imaging.PixelFormat.Format32bppPArgb, Imaging.PixelFormat.Format32bppRgb : bpp = 4
                Case Imaging.PixelFormat.Format8bppIndexed : bpp = 1
                Case Else : Return 250
            End Select

            Dim stride = data.Stride
            Dim width = data.Width
            Dim height = data.Height
            Dim ptr As IntPtr = data.Scan0
            Dim totalBytes = stride * height
            Dim buffer(totalBytes - 1) As Byte
            Marshal.Copy(ptr, buffer, 0, totalBytes)

            Dim hist(255) As Integer

            For y As Integer = 0 To height - 1
                Dim offRow = y * stride
                For x As Integer = 0 To width - 1
                    Dim off = offRow + x * bpp
                    Dim lum As Integer
                    If bpp = 1 Then
                        lum = buffer(off)
                    Else
                        Dim b = buffer(off)
                        Dim g = buffer(off + 1)
                        Dim r = buffer(off + 2)
                        lum = Math.Max(r, Math.Max(g, b))
                    End If
                    hist(lum) += 1
                Next
            Next

            Dim total As Integer = width * height
            Dim targetLow As Integer = CInt(total * lowPct)
            Dim targetHigh As Integer = CInt(total * highPct)

            Dim cum As Integer = 0
            Dim thrLow As Integer = 245
            Dim thrHigh As Integer = 252

            For i As Integer = 0 To 255
                cum += hist(i)
                If cum >= targetLow AndAlso thrLow = 245 Then thrLow = i
                If cum >= targetHigh Then
                    thrHigh = i
                    Exit For
                End If
            Next

            Dim thr As Integer = Math.Max(240, Math.Min(255, (thrLow + thrHigh) \ 2))
            Return CByte(thr)
        Finally
            Try : bmp.UnlockBits(data) : Catch : End Try
        End Try
    End Function

    Private Function ComputeTightContentRect(bmp As Bitmap, Optional bgThreshold As Byte = 250, Optional extraPadding As Integer = 2) As Rectangle
        Dim rectAll As New Rectangle(0, 0, bmp.Width, bmp.Height)
        Dim data = bmp.LockBits(rectAll, Imaging.ImageLockMode.ReadOnly, bmp.PixelFormat)
        Try
            Dim bpp As Integer
            Select Case data.PixelFormat
                Case Imaging.PixelFormat.Format24bppRgb : bpp = 3
                Case Imaging.PixelFormat.Format32bppArgb, Imaging.PixelFormat.Format32bppPArgb, Imaging.PixelFormat.Format32bppRgb : bpp = 4
                Case Imaging.PixelFormat.Format8bppIndexed : bpp = 1
                Case Else : Return rectAll
            End Select

            Dim stride = data.Stride
            Dim width = data.Width
            Dim height = data.Height
            Dim ptr As IntPtr = data.Scan0
            Dim totalBytes = stride * height
            Dim buffer(totalBytes - 1) As Byte
            Marshal.Copy(ptr, buffer, 0, totalBytes)

            Dim IsWhite As Func(Of Integer, Boolean) =
                Function(offset As Integer)
                    If bpp = 1 Then
                        Dim v = buffer(offset)
                        Return v >= bgThreshold
                    Else
                        Dim b = buffer(offset)
                        Dim g = buffer(offset + 1)
                        Dim r = buffer(offset + 2)
                        Return (r >= bgThreshold AndAlso g >= bgThreshold AndAlso b >= bgThreshold)
                    End If
                End Function

            Dim minX As Integer = width
            Dim minY As Integer = height
            Dim maxX As Integer = -1
            Dim maxY As Integer = -1

            For y As Integer = 0 To height - 1
                Dim offRow = y * stride
                For x As Integer = 0 To width - 1
                    Dim off = offRow + x * bpp
                    If Not IsWhite(off) Then
                        If x < minX Then minX = x
                        If y < minY Then minY = y
                        If x > maxX Then maxX = x
                        If y > maxY Then maxY = y
                    End If
                Next
            Next

            If maxX < 0 OrElse maxY < 0 Then
                Return Rectangle.Empty
            End If

            minX = Math.Max(0, minX - extraPadding)
            minY = Math.Max(0, minY - extraPadding)
            maxX = Math.Min(width - 1, maxX + extraPadding)
            maxY = Math.Min(height - 1, maxY + extraPadding)

            Dim rw = maxX - minX + 1
            Dim rh = maxY - minY + 1

            Return New Rectangle(minX, minY, rw, rh)
        Finally
            Try : bmp.UnlockBits(data) : Catch : End Try
        End Try
    End Function

    ' Cerca il rettangolo del pannello dentro il panelBmp usando EmguCV
    Private Function FindPanelRectangle(panelBmp As Bitmap) As Rectangle
        Try
            Using src As Mat = Emgu.CV.BitmapExtension.ToMat(panelBmp)
                Using gray As New Mat()
                    CvInvoke.CvtColor(src, gray, ColorConversion.Bgr2Gray)
                    CvInvoke.EqualizeHist(gray, gray)
                    CvInvoke.GaussianBlur(gray, gray, New Size(5, 5), 0)

                    Using th As New Mat()
                        CvInvoke.AdaptiveThreshold(gray, th, 255, AdaptiveThresholdType.GaussianC, ThresholdType.BinaryInv, 15, 7)

                        Using kernel As Mat = CvInvoke.GetStructuringElement(Emgu.CV.CvEnum.MorphShapes.Rectangle, New Size(15, 7), New Point(-1, -1))
                            CvInvoke.MorphologyEx(th, th, MorphOp.Close, kernel, New Point(-1, -1), 2, BorderType.Default, New MCvScalar())

                            Using contours As New VectorOfVectorOfPoint()
                                CvInvoke.FindContours(th, contours, Nothing, RetrType.External, ChainApproxMethod.ChainApproxSimple)

                                Dim bestRect As Rectangle = Rectangle.Empty
                                Dim bestScore As Double = 0

                                For i As Integer = 0 To contours.Size - 1
                                    Using cnt = contours(i)
                                        Dim area = CvInvoke.ContourArea(cnt)
                                        If area < 1000 Then Continue For

                                        Using approx As New VectorOfPoint()
                                            CvInvoke.ApproxPolyDP(cnt, approx, 0.02 * CvInvoke.ArcLength(cnt, True), True)
                                            Dim r = CvInvoke.BoundingRectangle(approx)

                                            Dim extent = area / (r.Width * r.Height + 0.000001)
                                            Dim normalizedArea = area / (panelBmp.Width * panelBmp.Height)
                                            Dim polyScore As Double = extent
                                            If approx.Size = 4 Then polyScore += 0.4
                                            Dim score = normalizedArea * 0.7 + polyScore * 0.3

                                            If score > bestScore Then
                                                bestScore = score
                                                bestRect = r
                                            End If
                                        End Using
                                    End Using
                                Next

                                If Not bestRect.IsEmpty AndAlso bestScore > 0.02 Then
                                    Dim pad = Math.Max(3, Math.Min(panelBmp.Width, panelBmp.Height) \ 100)
                                    Dim x = Math.Max(0, bestRect.X - pad)
                                    Dim y = Math.Max(0, bestRect.Y - pad)
                                    Dim w = Math.Min(panelBmp.Width - x, bestRect.Width + 2 * pad)
                                    Dim h = Math.Min(panelBmp.Height - y, bestRect.Height + 2 * pad)
                                    Return New Rectangle(x, y, w, h)
                                End If
                            End Using
                        End Using
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine("FindPanelRectangle error: " & ex.Message)
        End Try

        Return Rectangle.Empty
    End Function

    ' --------------------
    ' Editing UI helpers (PictureBox Zoom handling)
    ' --------------------
    Private Function GetImageDisplayRect(pb As PictureBox) As Rectangle
        If pb.Image Is Nothing Then Return Rectangle.Empty
        Dim img As Image = pb.Image
        Dim imgRatio = img.Width / CSng(img.Height)
        Dim boxRatio = pb.ClientSize.Width / CSng(pb.ClientSize.Height)
        Dim width As Integer, height As Integer, x As Integer, y As Integer
        If imgRatio > boxRatio Then
            width = pb.ClientSize.Width
            height = CInt(width / imgRatio)
            x = 0
            y = (pb.ClientSize.Height - height) \ 2
        Else
            height = pb.ClientSize.Height
            width = CInt(height * imgRatio)
            y = 0
            x = (pb.ClientSize.Width - width) \ 2
        End If
        Return New Rectangle(x, y, width, height)
    End Function

    Private Function ClientPointToImagePoint(p As Point) As Point
        Dim imgRect = GetImageDisplayRect(PicPanel)
        If imgRect.IsEmpty Then Return Point.Empty
        If Not imgRect.Contains(p) Then
            p.X = Math.Max(imgRect.Left, Math.Min(imgRect.Right, p.X))
            p.Y = Math.Max(imgRect.Top, Math.Min(imgRect.Bottom, p.Y))
        End If
        Dim ix = CInt((p.X - imgRect.Left) * (PicPanel.Image.Width / CSng(imgRect.Width)))
        Dim iy = CInt((p.Y - imgRect.Top) * (PicPanel.Image.Height / CSng(imgRect.Height)))
        Return New Point(ix, iy)
    End Function

    ' --------------------
    ' PictureBox mouse handlers
    ' --------------------
    Private Sub PicPanel_MouseDown(sender As Object, e As MouseEventArgs)
        If PicPanel.Image Is Nothing Then Return
        If e.Button <> MouseButtons.Left Then Return
        isDragging = True
        dragStart = e.Location
        selectionRect = Rectangle.Empty
        PicPanel.Capture = True
    End Sub

    Private Sub PicPanel_MouseMove(sender As Object, e As MouseEventArgs)
        If Not isDragging Then Return
        Dim p1 = ClientPointToImagePoint(dragStart)
        Dim p2 = ClientPointToImagePoint(e.Location)
        Dim x = Math.Min(p1.X, p2.X)
        Dim y = Math.Min(p1.Y, p2.Y)
        Dim w = Math.Abs(p2.X - p1.X)
        Dim h = Math.Abs(p2.Y - p1.Y)
        selectionRect = New Rectangle(x, y, w, h)
        PicPanel.Invalidate()
    End Sub

    Private Sub PicPanel_MouseUp(sender As Object, e As MouseEventArgs)
        If e.Button <> MouseButtons.Left Then Return
        If Not isDragging Then Return
        isDragging = False
        PicPanel.Capture = False

        ' se selezione troppo piccola -> annulla selezione
        If selectionRect.Width < 6 OrElse selectionRect.Height < 6 Then
            selectionRect = Rectangle.Empty
            isEditing = False
            hasUnsavedChanges = False
            UpdateSaveButtonState()
            PicPanel.Invalidate()
            Return
        End If

        ' conferma stato editing e marcatura modifiche (preview creata ma file non sovrascritto)
        isEditing = True
        hasUnsavedChanges = True
        UpdateSaveButtonState()

        ' crea preview ritagliata dall'originalImage (coordinate immagine)
        Try
            Dim previewBmp As Bitmap = Nothing
            SyncLock originalImage
                Using tmp As New Bitmap(originalImage)
                    Dim r As Rectangle = selectionRect
                    r.Intersect(New Rectangle(0, 0, tmp.Width, tmp.Height))
                    If r.IsEmpty Then
                        previewBmp = Nothing
                    Else
                        ' salva il rettangolo usato per la preview in lastCropRect
                        lastCropRect = r
                        previewBmp = New Bitmap(r.Width, r.Height, Imaging.PixelFormat.Format32bppArgb)
                        Using g = Graphics.FromImage(previewBmp)
                            g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                            g.DrawImage(tmp, New Rectangle(0, 0, r.Width, r.Height), r, GraphicsUnit.Pixel)
                        End Using
                    End If

                    ' ... dopo aver impostato la preview: azzera selectionRect ma non lastCropRect
                    selectionRect = Rectangle.Empty
                    isEditing = False
                End Using
            End SyncLock

            ' sostituisci immagine visualizzata con la preview (dispose della precedente immagine del PictureBox)
            If previewBmp IsNot Nothing Then
                If PicPanel.Image IsNot Nothing Then
                    Try : PicPanel.Image.Dispose() : Catch : End Try
                End If
                PicPanel.Image = New Bitmap(previewBmp)
                previewBmp.Dispose()
            End If

            ' azzera selezione e stato editing (la preview è visibile; l'utente può salvare con btnSalvaPanel)
            selectionRect = Rectangle.Empty
            isEditing = False

            PicPanel.Invalidate()
        Catch ex As Exception
            Debug.WriteLine("picPanel_MouseUp preview error: " & ex.Message)
        End Try
    End Sub


    Private Sub PicPanel_Paint(sender As Object, e As PaintEventArgs)
        If PicPanel.Image Is Nothing Then Return
        If selectionRect.IsEmpty Then Return

        Dim imgRect = GetImageDisplayRect(PicPanel)
        If imgRect.IsEmpty Then Return
        Dim scaleX = imgRect.Width / CSng(PicPanel.Image.Width)
        Dim scaleY = imgRect.Height / CSng(PicPanel.Image.Height)
        Dim drawRect As New Rectangle(CInt(imgRect.Left + selectionRect.X * scaleX),
                                      CInt(imgRect.Top + selectionRect.Y * scaleY),
                                      CInt(selectionRect.Width * scaleX),
                                      CInt(selectionRect.Height * scaleY))
        Using b As New SolidBrush(Color.FromArgb(60, Color.Red))
            e.Graphics.FillRectangle(b, drawRect)
        End Using
        Using p As New Pen(Color.FromArgb(200, Color.Red), 2)
            e.Graphics.DrawRectangle(p, drawRect)
        End Using
    End Sub

    Private Sub btnSalvaPanel_Click(sender As Object, e As EventArgs) Handles BtnSalvaPanel.Click
        If originalImage Is Nothing OrElse String.IsNullOrWhiteSpace(originalFilePath) Then Return

        ' usa selectionRect se presente, altrimenti lastCropRect
        Dim r As Rectangle = If(selectionRect.IsEmpty, lastCropRect, selectionRect)
        If r.IsEmpty Then
            MDIMessageBox.Show("Nessuna selezione valida da salvare.", Me.MdiParent, MessageBoxButtons.OK)
            Return
        End If

        Try
            Dim cropBmp As Bitmap = Nothing
            SyncLock originalImage
                Using bmp As New Bitmap(originalImage)
                    r.Intersect(New Rectangle(0, 0, bmp.Width, bmp.Height))
                    If r.IsEmpty Then
                        MDIMessageBox.Show("Selezione fuori dai limiti.", Me.MdiParent, MessageBoxButtons.OK)
                        Return
                    End If
                    cropBmp = New Bitmap(r.Width, r.Height, Imaging.PixelFormat.Format32bppArgb)
                    Using g = Graphics.FromImage(cropBmp)
                        g.DrawImage(bmp, New Rectangle(0, 0, r.Width, r.Height), r, GraphicsUnit.Pixel)
                    End Using
                End Using
            End SyncLock

            ' salva su file come già fatto
            Dim temp = originalFilePath & ".tmp"
            If File.Exists(temp) Then File.Delete(temp)
            cropBmp.Save(temp, Imaging.ImageFormat.Png)
            If File.Exists(originalFilePath) Then File.Delete(originalFilePath)
            File.Move(temp, originalFilePath)

            ' aggiorna immagini in memoria e UI
            If PicPanel.Image IsNot Nothing Then
                Try : PicPanel.Image.Dispose() : Catch : End Try
            End If
            PicPanel.Image = New Bitmap(cropBmp)
            If originalImage IsNot Nothing Then
                Try : originalImage.Dispose() : Catch : End Try
            End If
            originalImage = New Bitmap(PicPanel.Image)

            ' reset stato
            selectionRect = Rectangle.Empty
            lastCropRect = Rectangle.Empty
            isEditing = False
            hasUnsavedChanges = False
            UpdateSaveButtonState() ' <-- disabilita BtnSalvaPanel dopo il salvataggio

            'MDIMessageBox.Show("Immagine salvata.", Me.MdiParent, MessageBoxButtons.OK)
        Catch ex As Exception
            MDIMessageBox.Show("Errore salvataggio: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
        End Try
    End Sub

    Private Function ConfirmDiscardChanges(Optional skipSaveConfirmation As Boolean = False) As Boolean
        If Not hasUnsavedChanges Then Return True

        Dim msg As String = "Ci sono modifiche non salvate per il panel corrente. Vuoi salvare le modifiche prima di cambiare panel?"
        Dim title As String = "Modifiche non salvate"

        Dim res As DialogResult
        If skipSaveConfirmation = False Then
            res = MessageBox.Show(msg, title, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
        Else
            res = DialogResult.Yes
        End If
        Select Case res
            Case DialogResult.Yes
                Try
                    btnSalvaPanel_Click(BtnSalvaPanel, EventArgs.Empty)
                Catch ex As Exception
                    Return False
                End Try
                hasUnsavedChanges = False
                Return True
            Case DialogResult.No
                hasUnsavedChanges = False
                Return True
            Case Else
                Return False
        End Select
    End Function

    Private Function GetStoryboardPdfPathFromDb(comboStoryboard As ComboBox) As String
        If comboStoryboard Is Nothing OrElse comboStoryboard.SelectedItem Is Nothing Then
            Return String.Empty
        End If

        Dim storyboardId As String = String.Empty
        Try
            Dim drv = TryCast(comboStoryboard.SelectedItem, DataRowView)
            If drv IsNot Nothing Then
                storyboardId = Convert.ToString(drv("IdStoryboard"))
            Else
                storyboardId = If(comboStoryboard.SelectedValue IsNot Nothing, comboStoryboard.SelectedValue.ToString(), String.Empty)
            End If
        Catch
            storyboardId = If(comboStoryboard.SelectedValue IsNot Nothing, comboStoryboard.SelectedValue.ToString(), String.Empty)
        End Try

        If String.IsNullOrWhiteSpace(storyboardId) Then Return String.Empty

        Try
            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand("SELECT TOP 1 FilePDFStoryboard FROM Mov_Storyboard WHERE IdStoryboard = @id", conn)
                    cmd.Parameters.AddWithValue("@id", storyboardId)
                    conn.Open()
                    Dim res = cmd.ExecuteScalar()
                    Dim pdfPath = Convert.ToString(res).Trim()
                    If String.IsNullOrWhiteSpace(pdfPath) Then Return String.Empty
                    Return pdfPath
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine("GetStoryboardPdfPathFromDb error: " & ex.Message)
            Return String.Empty
        End Try
    End Function


    Private Sub BtnCaricaStoryboard_Click(sender As Object, e As EventArgs) Handles BtnCaricaStoryboard.Click
        Try
            ' Ottieni la cartella target usando la stessa logica esistente (controlli e creazione se necessario)
            Dim outDir = GetStoryboardOutputFolder(Me, ComboStoryboard)
            If String.IsNullOrWhiteSpace(outDir) Then Return

            ' Leggi tutti i PNG (ordinati per nome)
            Dim files = Directory.EnumerateFiles(outDir, "*.png", SearchOption.TopDirectoryOnly) _
                                 .OrderBy(Function(f) Path.GetFileName(f)) _
                                 .ToList()

            If files Is Nothing OrElse files.Count = 0 Then
                MDIMessageBox.Show("Nessun file PNG trovato nella cartella selezionata.", Me.MdiParent, MessageBoxButtons.OK)
                Return
            End If

            ' Imposta lista e mostra il primo
            exportedFiles = files
            currentIndex = 0
            ShowCurrentPanel()
            'MDIMessageBox.Show($"Caricati {exportedFiles.Count} pannelli da:{Environment.NewLine}{outDir}", Me.MdiParent, MessageBoxButtons.OK)
        Catch ex As Exception
            MDIMessageBox.Show("Errore caricamento storyboard: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub BtnAnnunllaModifica_Click(sender As Object, e As EventArgs) Handles BtnAnnunllaModifica.Click
        Try
            ' Annulla qualsiasi preview o selezione e ripristina l'immagine originale in memoria
            If originalImage IsNot Nothing Then
                ' dispose immagine attuale mostrata (potrebbe essere la preview)
                If PicPanel.Image IsNot Nothing Then
                    Try : PicPanel.Image.Dispose() : Catch : End Try
                End If

                ' ripristina la copia della originalImage
                PicPanel.Image = New Bitmap(originalImage)
            Else
                ' se non abbiamo originalImage, svuota PictureBox
                If PicPanel.Image IsNot Nothing Then
                    Try : PicPanel.Image.Dispose() : Catch : End Try
                End If
                PicPanel.Image = Nothing
            End If

            ' azzera stato di selezione e crop salvato
            selectionRect = Rectangle.Empty
            lastCropRect = Rectangle.Empty
            isEditing = False
            hasUnsavedChanges = False

            ' disattiva pulsante Salva Modifica
            UpdateSaveButtonState()

            PicPanel.Invalidate()
        Catch ex As Exception
            Debug.WriteLine("BtnAnnunllaModifica error: " & ex.Message)
        End Try
    End Sub

    Private Sub BtnCancellaPanel_Click(sender As Object, e As EventArgs) Handles BtnCancellaPanel.Click
        Try
            If exportedFiles Is Nothing OrElse exportedFiles.Count = 0 OrElse currentIndex < 0 OrElse currentIndex >= exportedFiles.Count Then
                MDIMessageBox.Show("Nessun pannello selezionato da cancellare.", Me.MdiParent, MessageBoxButtons.OK)
                Return
            End If

            Dim fileToDelete = exportedFiles(currentIndex)
            If String.IsNullOrWhiteSpace(fileToDelete) OrElse Not File.Exists(fileToDelete) Then
                MDIMessageBox.Show("File del pannello non trovato sul disco.", Me.MdiParent, MessageBoxButtons.OK)
                Return
            End If

            ' Prima conferma
            Dim q1 = MessageBox.Show("Sei sicuro di voler cancellare il pannello selezionato?", "Conferma cancellazione", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)
            If q1 <> DialogResult.Yes Then Return

            ' Seconda conferma (più esplicita)
            Dim q2 = MessageBox.Show("Questa operazione eliminerà definitivamente il file PNG dal disco. Procedere?", "Conferma definitiva", MessageBoxButtons.YesNo, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button2)
            If q2 <> DialogResult.Yes Then Return

            ' Se l'immagine è attualmente mostrata nel PictureBox, rimuovila e dispose in modo sicuro prima di cancellare il file
            If Not String.IsNullOrWhiteSpace(originalFilePath) AndAlso String.Equals(Path.GetFullPath(originalFilePath), Path.GetFullPath(fileToDelete), StringComparison.OrdinalIgnoreCase) Then
                Try
                    If PicPanel.Image IsNot Nothing Then
                        Try : PicPanel.Image.Dispose() : Catch : End Try
                    End If
                Catch
                End Try
                originalImage = Nothing
                originalFilePath = String.Empty
            End If

            ' Cancella file dal disco
            Try
                File.Delete(fileToDelete)
            Catch ex As Exception
                MDIMessageBox.Show("Errore cancellazione file: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
                Return
            End Try

            ' Rimuovi dalla lista e aggiorna indice
            exportedFiles.RemoveAt(currentIndex)
            If exportedFiles.Count = 0 Then
                currentIndex = -1
                PicPanel.Image = Nothing
                originalImage = Nothing
                originalFilePath = String.Empty
                selectionRect = Rectangle.Empty
                lastCropRect = Rectangle.Empty
                isEditing = False
                hasUnsavedChanges = False
                UpdateSaveButtonState()
                MDIMessageBox.Show("Pannello cancellato. Non ci sono più pannelli nella cartella.", Me.MdiParent, MessageBoxButtons.OK)
                Return
            Else
                ' Se abbiamo rimosso l'ultimo elemento, spostiamo l'indice all'ultimo valido
                If currentIndex >= exportedFiles.Count Then
                    currentIndex = exportedFiles.Count - 1
                End If
            End If

            ' Mostra il panel corrente aggiornato
            hasUnsavedChanges = False
            UpdateSaveButtonState()
            ShowCurrentPanel()
            MDIMessageBox.Show("Pannello cancellato correttamente.", Me.MdiParent, MessageBoxButtons.OK)
        Catch ex As Exception
            Debug.WriteLine("BtnCancellaPanel error: " & ex.Message)
            MDIMessageBox.Show("Errore durante la cancellazione del pannello: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
        End Try
    End Sub

    ' --- Helpers: leggi parametro da Sys_Parametri
    Private Function GetSysParamValue(descrizione As String) As String
        Try
            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand("SELECT TOP 1 Valore FROM Sys_Parametri WHERE Descrizione = @d", conn)
                    cmd.Parameters.AddWithValue("@d", descrizione)
                    conn.Open()
                    Dim res = cmd.ExecuteScalar()
                    Return If(res Is Nothing, String.Empty, Convert.ToString(res).Trim())
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine("GetSysParamValue error: " & ex.Message)
            Return String.Empty
        End Try
    End Function

    ' --- BtnAnimatic click handler completo
    Private Async Sub BtnAnimatic_Click(sender As Object, e As EventArgs) Handles BtnAnimatic.Click

        Try
            If exportedFiles Is Nothing OrElse exportedFiles.Count = 0 Then
                MDIMessageBox.Show("Nessun panel disponibile per l'animatic.", Me.MdiParent, MessageBoxButtons.OK)
                Return
            End If

            If Not animaticInCorso Then
                ' avvia creazione
                animaticInCorso = True
                BtnAnimatic.Text = "Annulla Creazione"
                BtnAnimatic.Enabled = True

                ctsAnimatic = New CancellationTokenSource()
                Dim token = ctsAnimatic.Token

                Try
                    ToggleUIForAnimatic(start:=True)

                    Dim success As Boolean = Await GenerateAnimaticAsync(token)

                    If success Then
                        MDIMessageBox.Show($"Animatic creato: {outMp4}", Me.MdiParent, MessageBoxButtons.OK)
                        Try
                            Dim psi As New ProcessStartInfo(outMp4) With {
                                .UseShellExecute = True,
                                .Verb = "open"
                            }
                            Process.Start(psi)
                        Catch ex As Exception
                            Debug.WriteLine("Impossibile aprire l'MP4: " & ex.Message)
                        End Try
                    Else
                        If token.IsCancellationRequested Then
                            MDIMessageBox.Show("Generazione annullata.", Me.MdiParent, MessageBoxButtons.OK)
                        Else
                            MDIMessageBox.Show("Creazione animatic fallita. Controlla ffmpeg e i percorsi.", Me.MdiParent, MessageBoxButtons.OK)
                        End If
                    End If
                Catch exOp As OperationCanceledException
                    MDIMessageBox.Show("Generazione annullata.", Me.MdiParent, MessageBoxButtons.OK)
                Catch ex As Exception
                    Debug.WriteLine("BtnAnimatic_Click error: " & ex.Message)
                    MDIMessageBox.Show("Errore durante la generazione: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
                Finally
                    ToggleUIForAnimatic(start:=False)
                    animaticInCorso = False
                    BtnAnimatic.Text = "Crea Animatic"
                    BtnAnimatic.Enabled = True
                    Try : ctsAnimatic?.Dispose() : Catch : End Try
                    ctsAnimatic = Nothing
                End Try
            Else
                ' richiesta annullamento
                Try
                    ctsAnimatic?.Cancel()
                    ' Forza feedback immediato disabilitando per evitare ripetuti tap
                    BtnAnimatic.Enabled = False
                Catch ex As Exception
                    Debug.WriteLine("Errore nella richiesta di cancellazione animatic: " & ex.Message)
                End Try
            End If
        Catch ex As Exception
            Debug.WriteLine("BtnAnimatic_Click outer error: " & ex.Message)
        End Try
    End Sub

    ' --- Toggle UI durante animatic
    Private Sub ToggleUIForAnimatic(start As Boolean)
        Try
            BtnCaricaStoryboard.Enabled = Not start
            BtnAcquisisciPDF.Enabled = Not start
            BtnPrev.Enabled = Not start
            BtnNext.Enabled = Not start
            Me.Cursor = If(start, Cursors.WaitCursor, Cursors.Default)
        Catch
        End Try
    End Sub

    ' --- GenerateAnimaticAsync (core)
    Private Async Function GenerateAnimaticAsync(token As CancellationToken) As Task(Of Boolean)
        Dim success As Boolean = False

        If exportedFiles Is Nothing OrElse exportedFiles.Count = 0 Then
            Throw New InvalidOperationException("Nessun panel disponibile per l'animatic.")
        End If

        ' chiedo percorso di salvataggio sull'UI thread
        Dim dlgRes As DialogResult = DialogResult.None
        Me.Invoke(New MethodInvoker(Sub()
                                        Using sfd As New SaveFileDialog()
                                            sfd.Filter = "Video|*.mp4"
                                            sfd.Title = "Salva Animatic"
                                            sfd.FileName = "Animatic.mp4"
                                            dlgRes = sfd.ShowDialog(Me)
                                            If dlgRes = DialogResult.OK Then outMp4 = sfd.FileName
                                        End Using
                                    End Sub))

        If dlgRes <> DialogResult.OK OrElse String.IsNullOrWhiteSpace(outMp4) Then
            Return False
        End If

        token.ThrowIfCancellationRequested()

        ' temp dir preso da Sys_Parametri PercorsoTempAnimatic
        Dim tempBase As String = GetSysParamValue("PercorsoTempAnimatic")
        If String.IsNullOrWhiteSpace(tempBase) Then
            Me.Invoke(New MethodInvoker(Sub() MDIMessageBox.Show("PercorsoTempAnimatic non configurato in Sys_Parametri.", Me.MdiParent, MessageBoxButtons.OK)))
            Return False
        End If

        Dim tempDir As String = Path.Combine(tempBase, "GesPu_temp_animatic")
        Try
            If Not Directory.Exists(tempDir) Then Directory.CreateDirectory(tempDir)
        Catch ex As Exception
            Me.Invoke(New MethodInvoker(Sub() MDIMessageBox.Show("Errore creando cartella temporanea: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)))
            Return False
        End Try

        Dim tempListFile As String = Path.Combine(tempDir, "ffmpeg_list.txt")

        ' 1) crea file-list per ffmpeg concat demuxer
        Try
            Using sw As New StreamWriter(tempListFile, False, New System.Text.UTF8Encoding(False))
                For Each f In exportedFiles
                    token.ThrowIfCancellationRequested()
                    If String.IsNullOrWhiteSpace(f) OrElse Not File.Exists(f) Then Continue For
                    ' escape single quote by doubling, ffmpeg accepts path in single quotes; safe path: wrap in single quotes
                    Dim safePath = f.Replace("'", "'\''")
                    sw.WriteLine("file '" & safePath & "'")
                    sw.WriteLine("duration 1")
                Next
                ' ripeti l'ultimo file senza duration per evitare frame final missing
                Dim last = exportedFiles.LastOrDefault(Function(x) Not String.IsNullOrWhiteSpace(x) AndAlso File.Exists(x))
                If Not String.IsNullOrWhiteSpace(last) Then
                    Dim safeLast = last.Replace("'", "'\''")
                    sw.WriteLine("file '" & safeLast & "'")
                End If
            End Using
        Catch ex As Exception
            Debug.WriteLine("Error writing ffmpeg list: " & ex.Message)
            Return False
        End Try

        token.ThrowIfCancellationRequested()

        ' 2) esegui ffmpeg in background e attendi termine (Task.Run)
        Dim ffmpegPath As String = "ffmpeg.exe" ' se necessario metti percorso completo

        ' target resolution (modifica se vuoi un'altra risoluzione)
        Dim targetW As Integer = 1920 / 2
        Dim targetH As Integer = 1080 / 2

        ' filtro per scalare mantenendo aspect ratio e pad per ottenere esattamente targetWxtargetH
        Dim vfFilter As String = $"scale={targetW}:{targetH}:force_original_aspect_ratio=decrease,pad={targetW}:{targetH}:(ow-iw)/2:(oh-ih)/2,format=yuv420p"

        ' opzionali: controlla bitrate/quality
        Dim crf As Integer = 18 ' qualità (più basso = migliore, 0..51); 18 è alta qualità
        Dim preset As String = "medium" ' ultrafast, superfast, veryfast, faster, fast, medium, slow, slower, veryslow

        Dim args As String = String.Format("-y -f concat -safe 0 -i ""{0}"" -vf ""{1}"" -c:v libx264 -preset {2} -crf {3} -r 24 -movflags +faststart ""{4}""", tempListFile, vfFilter, preset, crf, outMp4)

        Dim tcs As New TaskCompletionSource(Of Boolean)()

        Try
            Await Task.Run(Function()
                               Dim proc As Process = Nothing
                               Dim procStarted As Boolean = False
                               Try
                                   Dim psi As New ProcessStartInfo() With {
                                   .FileName = ffmpegPath,
                                   .Arguments = args,
                                   .CreateNoWindow = True,
                                   .UseShellExecute = False,
                                   .RedirectStandardError = True,
                                   .RedirectStandardOutput = True
                               }
                                   proc = Process.Start(psi)
                                   procStarted = (proc IsNot Nothing)
                                   ' leggi stderr in background per debug
                                   Dim stderrTask = Task.Run(Sub()
                                                                 Try
                                                                     Dim err = proc.StandardError.ReadToEnd()
                                                                     If Not String.IsNullOrEmpty(err) Then Debug.WriteLine(err)
                                                                 Catch exRead As Exception
                                                                     Debug.WriteLine("ffmpeg stderr read error: " & exRead.Message)
                                                                 End Try
                                                             End Sub)

                                   ' loop che osserva cancellazione senza bloccare UI
                                   While Not proc.HasExited
                                       If token.IsCancellationRequested Then
                                           Try
                                               proc.Kill()
                                           Catch
                                           End Try
                                           Exit While
                                       End If
                                       Threading.Thread.Sleep(150)
                                   End While

                                   If proc IsNot Nothing Then
                                       Try : proc.WaitForExit() : Catch : End Try
                                   End If

                                   If token.IsCancellationRequested Then
                                       tcs.TrySetResult(False)
                                   ElseIf proc IsNot Nothing AndAlso proc.ExitCode = 0 AndAlso File.Exists(outMp4) Then
                                       tcs.TrySetResult(True)
                                   Else
                                       tcs.TrySetResult(False)
                                   End If

                               Catch exRun As Exception
                                   Debug.WriteLine("ffmpeg spawn error: " & exRun.Message)
                                   tcs.TrySetResult(False)
                               Finally
                                   If proc IsNot Nothing Then
                                       Try : proc.Dispose() : Catch : End Try
                                   End If
                               End Try
                               Return tcs.Task.Result
                           End Function, token)
            success = tcs.Task.Result

        Catch exCancel As OperationCanceledException
            success = False
        Catch ex As Exception
            Debug.WriteLine("GenerateAnimaticAsync ffmpeg error: " & ex.Message)
            success = False
        Finally
            ' pulizia file temporanei
            Try
                If File.Exists(tempListFile) Then File.Delete(tempListFile)
            Catch : End Try
        End Try

        Return success
    End Function


End Class


'
' CLASSE PER GENERARE I RECROS DEI PANNELL ALL'INTERNO DELLA TABELLA MOV_STORYBOARDPANEL
'


