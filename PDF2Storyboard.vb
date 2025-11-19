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

Public Class PDF2Storyboard

    ' Configurabili
    Private Const DefaultRenderDpi As Single = 600.0F
    Private ReadOnly SameRowTolerancePx As Integer = 40

    Private exportedFiles As New List(Of String)()
    Private currentIndex As Integer = -1
    Private cts As CancellationTokenSource = Nothing

    ' --------------------
    ' Form lifecycle
    ' --------------------
    Private Sub PDF2Storyboard_Load(sender As Object, e As EventArgs) Handles Me.Load
        RipristinaPosizioneForm(Me)
        If picPanel IsNot Nothing Then picPanel.SizeMode = PictureBoxSizeMode.Zoom
    End Sub

    Private Sub PDF2Storyboard_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Try
            If cts IsNot Nothing Then
                cts.Cancel()
                cts.Dispose()
            End If
        Catch
        End Try
        SalvaPosizioneForm(Me)
    End Sub

    Private Sub btnAnnulla_Click(sender As Object, e As EventArgs) Handles btnAnnulla.Click
        Me.Close()
    End Sub

    ' --------------------
    ' Avvio pipeline
    ' --------------------
    Private Async Sub btnAcquisisciPDF_Click(sender As Object, e As EventArgs) Handles btnAcquisisciPDF.Click
        Try
            Dim pdfPath = SelezionaPdf()
            If String.IsNullOrWhiteSpace(pdfPath) OrElse Not File.Exists(pdfPath) Then Return

            Dim outDir = GetStoryboardOutputFolder(Me, txtStoryboard)
            If String.IsNullOrWhiteSpace(outDir) Then Return

            btnAcquisisciPDF.Enabled = False
            Cursor = Cursors.WaitCursor

            exportedFiles.Clear()
            currentIndex = -1

            cts = New CancellationTokenSource()
            Dim token = cts.Token

            Dim files = Await Task.Run(Function() ProcessPdfToPanels(pdfPath, outDir, token), token)

            If files Is Nothing OrElse files.Count = 0 Then
                MDIMessageBox.Show("Nessun pannello trovato/esportato.", Me.MdiParent, MessageBoxButtons.OK)
                Return
            End If

            exportedFiles = files
            currentIndex = 0
            ShowCurrentPanel()
            MDIMessageBox.Show($"Esportazione completata: {exportedFiles.Count} PNG in{Environment.NewLine}{Path.GetDirectoryName(exportedFiles(0))}", Me.MdiParent, MessageBoxButtons.OK)

        Catch ex As OperationCanceledException
            MDIMessageBox.Show("Operazione annullata.", Me.MdiParent, MessageBoxButtons.OK)
        Catch ex As Exception
            MDIMessageBox.Show("Errore: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
        Finally
            Try
                btnAcquisisciPDF.Enabled = True
                Cursor = Cursors.Default
                If cts IsNot Nothing Then cts.Dispose() : cts = Nothing
            Catch
            End Try
        End Try
    End Sub

    ' --------------------
    ' Navigazione
    ' --------------------
    Private Sub btnPrev_Click(sender As Object, e As EventArgs) Handles btnPrev.Click
        If exportedFiles.Count = 0 Then Return
        currentIndex = Math.Max(0, currentIndex - 1)
        ShowCurrentPanel()
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        If exportedFiles.Count = 0 Then Return
        currentIndex = Math.Min(exportedFiles.Count - 1, currentIndex + 1)
        ShowCurrentPanel()
    End Sub

    Private Sub ShowCurrentPanel()
        If currentIndex < 0 OrElse currentIndex >= exportedFiles.Count Then
            picPanel.Image = Nothing
            Return
        End If
        Try
            Using fs As New FileStream(exportedFiles(currentIndex), FileMode.Open, FileAccess.Read, FileShare.Read)
                Using mem As New MemoryStream()
                    fs.CopyTo(mem)
                    mem.Seek(0, SeekOrigin.Begin)
                    picPanel.Image = Image.FromStream(mem)
                End Using
            End Using
            Me.Text = $"Panel {currentIndex + 1}/{exportedFiles.Count} - {Path.GetFileName(exportedFiles(currentIndex))}"
        Catch ex As Exception
            Debug.WriteLine("ShowCurrentPanel error: " & ex.Message)
            picPanel.Image = Nothing
        End Try
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

                    Dim crops = SegmentaPannelliDaBitmap(bmpPage) ' griglia 3x2

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

                            ' Soglia bianco adattiva + trim
                            Dim autoThr As Byte = ComputeAdaptiveWhiteThreshold(panelBmp, 0.92, 0.98) ' percentile 92–98
                            Dim tight As Rectangle = ComputeTightContentRect(panelBmp, autoThr, 3)   ' 3 px padding

                            If tight.Width > 0 AndAlso tight.Height > 0 Then
                                Using finalBmp As Bitmap = panelBmp.Clone(tight, panelBmp.PixelFormat)
                                    Dim fileName As String
                                    If Not String.IsNullOrWhiteSpace(sceneLabel) AndAlso Not String.IsNullOrWhiteSpace(panelLabel) Then
                                        fileName = $"{sceneLabel}_{panelLabel}_{idx}.png"
                                    Else
                                        fileName = $"Page_{i + 1:00}_{idx}.png"
                                    End If
                                    fileName = MakeValidFileName(fileName)
                                    Dim outFile = Path.Combine(outDir, fileName)
                                    SafeSaveBitmap(finalBmp, outFile)
                                    localExported.Add(outFile)
                                End Using
                            Else
                                ' fallback: salva panel senza trim
                                Dim fileName As String = $"Page_{i + 1:00}_{idx}.png"
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

    Private Function GetStoryboardOutputFolder(form As Form, txtStoryboard As TextBox) As String
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

        Dim storyboardName = txtStoryboard.Text.Trim()
        If String.IsNullOrWhiteSpace(storyboardName) Then
            MDIMessageBox.Show("Inserisci il nome dello storyboard nel campo txtStoryboard.", form.MdiParent, MessageBoxButtons.OK)
            Return String.Empty
        End If

        Dim outDir = Path.Combine(basePath, "PanelsStoryboard", SanitizeFolderName(storyboardName))
        Try
            Directory.CreateDirectory(outDir)
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
        Dim bmp As Bitmap = Nothing
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
    ' Trim adattivo del bianco per ogni panel
    ' --------------------
    ' Calcola una soglia "bianco" adattiva basata sui percentili della luminosità (0..255)
    ' lowPct e highPct sono percentuali dell'istogramma (es. 0.92 e 0.98)
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
                        ' Luminosità semplice: massimo dei canali (robusto per bianco puro)
                        lum = Math.Max(r, Math.Max(g, b))
                    End If
                    hist(lum) += 1
                Next
            Next

            ' Cumulata e scelta soglia tra i percentili alti
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

            ' Scegli una soglia tra low e high, clamp ragionevole
            Dim thr As Integer = Math.Max(240, Math.Min(255, (thrLow + thrHigh) \ 2))
            Return CByte(thr)
        Finally
            Try : bmp.UnlockBits(data) : Catch : End Try
        End Try
    End Function

    ' Calcola il rettangolo stretto con pixel non bianchi, usando soglia e padding
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
                        ' Considera bianco solo se tutti i canali sono sopra soglia (evita grigi)
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

End Class
