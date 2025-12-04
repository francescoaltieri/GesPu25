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

    ' -----------------------------
    ' SaveFrame overload: salva overlay in modo atomico
    ' -----------------------------
    Public Sub SaveFrame(frameIndex As Integer, drawingBmp As Bitmap)
        If frameIndex < 0 OrElse FrameList Is Nothing OrElse frameIndex >= FrameList.Count Then
            Throw New ArgumentOutOfRangeException(NameOf(frameIndex))
        End If

        If drawingBmp Is Nothing Then
            ' Non c'è disegno: niente da salvare qui (gestione DB rimane a chiama)
            Return
        End If

        Dim basePath = FrameList(frameIndex)
        Dim overlayPath = Path.Combine(Path.GetDirectoryName(basePath), Path.GetFileNameWithoutExtension(basePath) & "_overlay.png")
        Dim tempPath As String = Nothing

        SyncLock Me
            Try
                ' Determina dimensioni del base (fallback alle dimensioni del drawing)
                Dim baseWidth As Integer = Math.Max(1, drawingBmp.Width)
                Dim baseHeight As Integer = Math.Max(1, drawingBmp.Height)
                Try
                    Using fs As New FileStream(basePath, FileMode.Open, FileAccess.Read, FileShare.Read)
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
                    ' se non riesce a leggere il base, usiamo dimensioni del drawing
                    baseWidth = Math.Max(1, drawingBmp.Width)
                    baseHeight = Math.Max(1, drawingBmp.Height)
                End Try

                ' Crea bitmap trasparente delle dimensioni del base e disegna il drawing in alto a sinistra
                Using bmpToSave As New Bitmap(baseWidth, baseHeight, Imaging.PixelFormat.Format32bppArgb)
                    Using g As Graphics = Graphics.FromImage(bmpToSave)
                        g.Clear(System.Drawing.Color.Transparent)
                        Dim drawW = Math.Min(drawingBmp.Width, baseWidth)
                        Dim drawH = Math.Min(drawingBmp.Height, baseHeight)
                        g.DrawImage(drawingBmp, 0, 0, drawW, drawH)
                    End Using

                    ' Salvataggio atomico su file (tmp -> replace/move)
                    Dim dir = Path.GetDirectoryName(overlayPath)
                    If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
                    tempPath = Path.Combine(dir, Path.GetFileNameWithoutExtension(overlayPath) & "_tmp" & Path.GetExtension(overlayPath))

                    ' Rimuovi tmp precedente se esiste
                    If File.Exists(tempPath) Then
                        Try : File.Delete(tempPath) : Catch : End Try
                    End If

                    bmpToSave.Save(tempPath, Imaging.ImageFormat.Png)

                    ' Sostituzione atomica
                    If File.Exists(overlayPath) Then
                        Try
                            File.Replace(tempPath, overlayPath, Nothing)
                        Catch ex As PlatformNotSupportedException
                            ' fallback: delete + move
                            Try : File.Delete(overlayPath) : Catch : End Try
                            File.Move(tempPath, overlayPath)
                        End Try
                    Else
                        File.Move(tempPath, overlayPath)
                    End If
                End Using

                ' Pulizia stato undo/state stack: dopo il salvataggio consideriamo lo stato salvato
                Try
                    If stateStack IsNot Nothing Then
                        For Each b In stateStack
                            Try : b.Dispose() : Catch : End Try
                        Next
                        stateStack.Clear()
                    End If
                Catch
                End Try

                ' Aggiorna flag
                HasUnsavedChanges = False
                Try : UndoStack.Clear() : Catch : End Try
            Catch ex As Exception
                ' Se qualcosa va storto, tenta di rimuovere tmp e rilancia per far gestire l'errore al chiamante
                Try
                    If Not String.IsNullOrEmpty(tempPath) AndAlso File.Exists(tempPath) Then
                        Try : File.Delete(tempPath) : Catch : End Try
                    End If
                Catch
                End Try
                Throw
            End Try
        End SyncLock
    End Sub


    Public Sub ClearFrameAnnotations(index As Integer)
        If index < 0 OrElse index >= FrameList.Count Then Return

        Try
            If FrameNote IsNot Nothing AndAlso FrameNote.ContainsKey(index) Then
                FrameNote.Remove(index)
            End If
        Catch
        End Try

        Try
            If stateStack IsNot Nothing Then
                For Each b In stateStack
                    Try : CType(b, Bitmap).Dispose() : Catch : End Try
                Next
                stateStack.Clear()
            End If
        Catch
        End Try

        Try
            If DrawingBitmap IsNot Nothing Then
                DrawingBitmap.Dispose()
                DrawingBitmap = Nothing
            End If
        Catch
        End Try

        HasUnsavedChanges = False
    End Sub

    Public Function LoadFrame(index As Integer) As Bitmap
        If index < 0 OrElse index >= FrameList.Count Then Return Nothing

        Dim basePath = FrameList(index)
        Dim baseBmp As Bitmap = Nothing

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
                    ' overlay cancellato 
                Catch ex As IOException
                    ' lock o I/O temporaneo
                Catch
                    ' altri errori
                End Try
            End If
        Catch
            ' sicurezza
        End Try

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

    Public Sub PrepareDrawingBitmapForEditing(frameIndex As Integer)
        If frameIndex < 0 OrElse FrameList Is Nothing OrElse frameIndex >= FrameList.Count Then
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
            ' se non riesce a leggere base
        End Try

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

    Public Sub RebuildVideo(outputPath As String)
        Dim ffmpegArgs As String = $"-framerate {framerate.ToString(CultureInfo.InvariantCulture)} -i ""{FrameDirectory}\frame_%04d.png"" -c:v libx264 -pix_fmt yuv420p ""{outputPath}"""
        Dim proc As New Process()
        proc.StartInfo = New ProcessStartInfo("ffmpeg.exe", ffmpegArgs) With {
            .CreateNoWindow = True,
            .UseShellExecute = False
        }
        proc.Start()
        proc.WaitForExit()
    End Sub

End Class

Public Class FrameNota
    Public Property Testo As String
    Public Property Autore As String
    Public Property Data As DateTime
End Class

