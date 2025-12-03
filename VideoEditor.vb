Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Globalization
Imports System.IO
Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml.Wordprocessing

Public Class VideoEditor
    Public Property VideoPath As String
    Public Property FrameDirectory As String
    Public Property UndoStack As New Stack(Of Bitmap)
    Private framerate As Double = 30.0 ' default
    Public FrameNote As New Dictionary(Of Integer, FrameNota)
    Public Property FrameList As List(Of String)
    Public Property CurrentIndex As Integer = 0
    Public Property DrawingBitmap As Bitmap
    Private stateStack As New Stack(Of Bitmap)
    Public Property HasUnsavedChanges As Boolean = False
    Private framesDirectory As String
    Private baseVideoPath As String

    Public Sub New(videoPath As String, frameDir As String)
        Me.VideoPath = videoPath
        Me.FrameDirectory = frameDir
        Me.framerate = GetFramerate()

        ' Carica i frame esistenti se presenti
        If Directory.Exists(frameDir) Then
            Me.FrameList = Directory.GetFiles(frameDir, "frame_*.png").
            OrderBy(Function(f)
                        Dim nome = Path.GetFileNameWithoutExtension(f)
                        Dim numero = Regex.Match(nome, "\d+").Value
                        Return Integer.Parse(numero)
                    End Function).ToList()
        Else
            Me.FrameList = New List(Of String)
        End If
    End Sub

    Public Sub ExtractFrames()
        If Not Directory.Exists(FrameDirectory) Then
            Directory.CreateDirectory(FrameDirectory)
        End If

        Dim ffmpegArgs As String = $"-i ""{VideoPath}"" ""{FrameDirectory}\frame_%04d.png"""
        Dim proc As New Process()
        proc.StartInfo = New ProcessStartInfo("ffmpeg.exe", ffmpegArgs) With {
            .CreateNoWindow = True,
            .UseShellExecute = False
        }
        proc.Start()
        proc.WaitForExit()

        FrameList = Directory.GetFiles(FrameDirectory, "frame_*.png").OrderBy(Function(f) f).ToList()
    End Sub

    Public Sub SaveState()
        SyncLock Me
            If DrawingBitmap IsNot Nothing Then
                stateStack.Push(CType(DrawingBitmap.Clone(), Bitmap))
                HasUnsavedChanges = True
            End If
        End SyncLock
    End Sub

    Public Function Undo() As Boolean
        SyncLock Me
            If stateStack.Count = 0 Then
                Return False
            End If
            Dim prev = stateStack.Pop()
            If DrawingBitmap IsNot Nothing Then
                DrawingBitmap.Dispose()
            End If
            DrawingBitmap = CType(prev.Clone(), Bitmap)
            prev.Dispose()
            HasUnsavedChanges = (stateStack.Count > 0)
            Return HasUnsavedChanges
        End SyncLock
    End Function

    Public Sub SaveFrame()
        Dim idx = Me.CurrentIndex
        If idx < 0 OrElse idx >= FrameList.Count Then Return
        If String.IsNullOrWhiteSpace(VideoFBF.txtNote.Text) Then
            MDIMessageBox.Show("Nessuna nota da salvare", GesPu25, MessageBoxButtons.OK)
            Return
        End If

        Dim framePath = FrameList(idx)
        If Microsoft.VisualBasic.Strings.Right(framePath, 12) = "_overlay.png" Then
            framePath = Microsoft.VisualBasic.Strings.Left(framePath, Len(framePath) - 12)
        End If
        Dim overlayPath = Path.Combine(Path.GetDirectoryName(framePath), Path.GetFileNameWithoutExtension(framePath) & "_overlay.png")
        Dim tempOverlay = Path.Combine(Path.GetDirectoryName(framePath), Path.GetFileNameWithoutExtension(framePath) & "_overlay_tmp.png")
        Dim logPath = Path.Combine(Path.GetTempPath(), "VideoEditor_save.log")

        SyncLock Me
            Try
                ' Verifica che DrawingBitmap esista
                If DrawingBitmap Is Nothing Then Return

                ' Carica dimensioni del base per garantire che l'overlay abbia la stessa risoluzione
                Dim baseWidth As Integer = 0
                Dim baseHeight As Integer = 0
                Try
                    Using fs As New FileStream(framePath, FileMode.Open, FileAccess.Read, FileShare.Read)
                        Using tmpMs As New MemoryStream()
                            fs.CopyTo(tmpMs)
                            tmpMs.Position = 0
                            Using tmpImg As Image = Image.FromStream(tmpMs)
                                baseWidth = tmpImg.Width
                                baseHeight = tmpImg.Height
                            End Using
                        End Using
                    End Using
                Catch ex As Exception
                    ' Se non riusciamo a leggere il base, usiamo le dimensioni del DrawingBitmap
                    baseWidth = DrawingBitmap.Width
                    baseHeight = DrawingBitmap.Height
                End Try

                ' Prepara bitmap da salvare: non comporre overlay precedente, crea una bitmap trasparente delle dimensioni del base
                Using bmpToSave As New Bitmap(baseWidth, baseHeight, Imaging.PixelFormat.Format32bppArgb)
                    Using g As Graphics = Graphics.FromImage(bmpToSave)
                        g.Clear(System.Drawing.Color.FromArgb(0, 0, 0, 0))   ' ARGB: alpha = 0 (completamente trasparente)
                        ' Se DrawingBitmap ha dimensioni diverse, scala o posiziona al centro a seconda delle esigenze
                        If DrawingBitmap.Width = baseWidth AndAlso DrawingBitmap.Height = baseHeight Then
                            g.DrawImage(DrawingBitmap, 0, 0)
                        Else
                            ' disegna in alto a sinistra senza scalare (evita composizioni multiple)
                            g.DrawImage(DrawingBitmap, 0, 0, Math.Min(DrawingBitmap.Width, baseWidth), Math.Min(DrawingBitmap.Height, baseHeight))
                        End If
                    End Using

                    ' Salva su file temporaneo (sovrascrive temp se esiste)
                    If File.Exists(tempOverlay) Then
                        Try : File.Delete(tempOverlay) : Catch : End Try
                    End If
                    bmpToSave.Save(tempOverlay, ImageFormat.Png)
                End Using

                ' Sostituzione atomica dell'overlay
                Try
                    If File.Exists(overlayPath) Then
                        File.Replace(tempOverlay, overlayPath, Nothing)
                    Else
                        File.Move(tempOverlay, overlayPath)
                    End If
                Catch ex As PlatformNotSupportedException
                    If File.Exists(overlayPath) Then File.Delete(overlayPath)
                    File.Move(tempOverlay, overlayPath)
                End Try

                ' Pulisci stack e stato
                If stateStack IsNot Nothing Then
                    For Each b In stateStack
                        Try : b.Dispose() : Catch : End Try
                    Next
                    stateStack.Clear()
                End If
                HasUnsavedChanges = False

            Catch ex As Exception
                Try
                    File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - SaveFrame error: {ex.Message}{Environment.NewLine}")
                Catch
                End Try
                Throw
            End Try
        End SyncLock
    End Sub


    ' Public method to clear any in-memory annotations for a specific frame index
    Public Sub ClearFrameAnnotations(index As Integer)
        If index < 0 OrElse index >= FrameList.Count Then Return

        ' Remove FrameNote entry if present
        Try
            If FrameNote IsNot Nothing AndAlso FrameNote.ContainsKey(index) Then
                FrameNote.Remove(index)
            End If
        Catch
        End Try

        ' Clear any undo/state entries that refer to this frame (if you store per-frame stacks)
        Try
            ' If you keep a global stateStack for the current frame, clear it
            If stateStack IsNot Nothing Then
                For Each b In stateStack
                    Try : CType(b, Bitmap).Dispose() : Catch : End Try
                Next
                stateStack.Clear()
            End If
        Catch
        End Try

        ' Reset drawing bitmap for safety
        Try
            If DrawingBitmap IsNot Nothing Then
                DrawingBitmap.Dispose()
                DrawingBitmap = Nothing
            End If
        Catch
        End Try

        HasUnsavedChanges = False
    End Sub

    ' ----------------------------
    ' LoadFrame (carica base + overlay se presente, robusto)
    ' ----------------------------
    Public Function LoadFrame(index As Integer) As Bitmap
        If index < 0 OrElse index >= FrameList.Count Then Return Nothing

        Dim basePath = FrameList(index)
        Dim baseBmp As Bitmap = Nothing

        ' Carica immagine principale senza lock del file
        Try
            Using fs As New FileStream(basePath, FileMode.Open, FileAccess.Read, FileShare.Read)
                Using ms As New MemoryStream()
                    fs.CopyTo(ms)
                    ms.Position = 0
                    Using tmpImg As Image = Image.FromStream(ms)
                        baseBmp = New Bitmap(tmpImg) ' copia indipendente
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Throw New IOException($"Impossibile caricare il frame: {basePath}", ex)
        End Try

        ' Applica overlay se presente (file *_overlay.png)
        Try
            Dim overlayPath = Path.Combine(Path.GetDirectoryName(basePath), Path.GetFileNameWithoutExtension(basePath) & "_overlay.png")
            If File.Exists(overlayPath) Then
                Try
                    Using fsOv As New FileStream(overlayPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                        Using msOv As New MemoryStream()
                            fsOv.CopyTo(msOv)
                            msOv.Position = 0
                            Using tmpOv As Image = Image.FromStream(msOv)
                                Using ovBmp As New Bitmap(tmpOv)
                                    Using g As Graphics = Graphics.FromImage(baseBmp)
                                        g.DrawImage(ovBmp, 0, 0)
                                    End Using
                                End Using
                            End Using
                        End Using
                    End Using
                Catch ex As FileNotFoundException
                    ' overlay cancellato tra Exists e Open: ignoriamo
                Catch ex As IOException
                    ' lock o I/O temporaneo: ignoriamo overlay
                Catch
                    ' altri errori: non bloccare il caricamento del frame base
                End Try
            End If
        Catch
            ' sicurezza: non propagare errori di overlay
        End Try

        ' Aggiorna DrawingBitmap
        If DrawingBitmap IsNot Nothing Then
            Try : DrawingBitmap.Dispose() : Catch : End Try
            DrawingBitmap = Nothing
        End If

        DrawingBitmap = CType(baseBmp.Clone(), Bitmap)
        baseBmp.Dispose()

        HasUnsavedChanges = False
        Try : stateStack.Clear() : Catch : End Try

        Return CType(DrawingBitmap.Clone(), Bitmap)
    End Function

    Private Function GetFramerate() As Double
        Dim output As String = ""
        Dim proc As New Process()
        proc.StartInfo = New ProcessStartInfo("ffmpeg.exe", $"-i ""{VideoPath}""") With {
            .RedirectStandardError = True,
            .UseShellExecute = False,
            .CreateNoWindow = True
        }
        proc.Start()
        output = proc.StandardError.ReadToEnd()
        proc.WaitForExit()

        Dim match = Regex.Match(output, "(\d+(\.\d+)?)\s+fps")
        If match.Success Then
            Return Double.Parse(match.Groups(1).Value, CultureInfo.InvariantCulture)
        End If

        Return 30.0 ' fallback
    End Function

    Public Sub RebuildVideo(outputPath As String)
        ' Implementazione esistente per ricostruire il video dai frame
    End Sub

    Public Sub PrepareDrawingBitmapForEditing(frameIndex As Integer)
        If frameIndex < 0 OrElse FrameList Is Nothing OrElse frameIndex >= FrameList.Count Then
            ' fallback: crea una 1x1 trasparente per evitare NullReference
            Try
                If DrawingBitmap IsNot Nothing Then
                    DrawingBitmap.Dispose()
                    DrawingBitmap = Nothing
                End If
            Catch
            End Try
            DrawingBitmap = New System.Drawing.Bitmap(1, 1, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
            HasUnsavedChanges = False
            Try
                If stateStack IsNot Nothing Then
                    stateStack.Clear()
                End If
            Catch
            End Try
        Return
        End If

        Dim basePath = FrameList(frameIndex)
        Dim width As Integer = 1
        Dim height As Integer = 1

        Try
            Using fs As New System.IO.FileStream(basePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                Using ms As New System.IO.MemoryStream()
                    fs.CopyTo(ms)
                    ms.Position = 0
                    Using img As System.Drawing.Image = System.Drawing.Image.FromStream(ms)
                        width = Math.Max(1, img.Width)
                        height = Math.Max(1, img.Height)
                    End Using
                End Using
            End Using
        Catch
            ' se non riesce a leggere il base, mantieni fallback 1x1
        End Try

        ' Sostituisci DrawingBitmap in modo sicuro
        Try
            If DrawingBitmap IsNot Nothing Then
                Try : DrawingBitmap.Dispose() : Catch : End Try
                DrawingBitmap = Nothing
            End If
        Catch
        End Try

        Try
            DrawingBitmap = New System.Drawing.Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
            Using g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(DrawingBitmap)
                g.Clear(System.Drawing.Color.Transparent)
            End Using
        Catch
            ' in caso di errore crea 1x1 trasparente
            Try
                If DrawingBitmap IsNot Nothing Then DrawingBitmap.Dispose()
            Catch
            End Try
            DrawingBitmap = New System.Drawing.Bitmap(1, 1, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
        End Try

        ' Pulisci undo/state stack se presente
        Try
            If stateStack IsNot Nothing Then
                For Each b In stateStack
                    Try : b.Dispose() : Catch : End Try
                Next
                stateStack.Clear()
            End If
        Catch
        End Try

        HasUnsavedChanges = False
    End Sub


End Class


Public Class FrameNota
    Public Property Testo As String
    Public Property Autore As String
    Public Property Data As DateTime
End Class

