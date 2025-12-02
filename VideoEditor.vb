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

    ' Salva overlay PNG per il frame corrente
    Public Sub SaveFrame()
        Dim idx = Me.CurrentIndex
        If idx < 0 OrElse idx >= FrameList.Count Then Return

        Dim framePath = FrameList(idx)
        Dim overlayPath = Path.Combine(Path.GetDirectoryName(framePath), Path.GetFileNameWithoutExtension(framePath) & "_overlay.png")

        SyncLock Me
            ' Salva overlay
            DrawingBitmap.Save(overlayPath, System.Drawing.Imaging.ImageFormat.Png)
            ' Dopo il salvataggio, resetta lo stack e lo stato
            For Each b In stateStack
                b.Dispose()
            Next
            stateStack.Clear()
            HasUnsavedChanges = False
        End SyncLock
    End Sub

    ' Carica frame base e compone overlay se presente
    Public Function LoadFrame(index As Integer) As Bitmap
        If index < 0 OrElse index >= FrameList.Count Then Return Nothing
        Dim basePath = FrameList(index)
        Dim baseBmp As Bitmap = CType(Bitmap.FromFile(basePath), Bitmap)
        Dim overlayPath = Path.Combine(Path.GetDirectoryName(basePath), Path.GetFileNameWithoutExtension(basePath) & "_overlay.png")
        If File.Exists(overlayPath) Then
            Using ov = CType(Bitmap.FromFile(overlayPath), Bitmap)
                Using g = Graphics.FromImage(baseBmp)
                    g.DrawImage(ov, 0, 0)
                End Using
            End Using
        End If

        ' Imposta DrawingBitmap con la copia caricata
        If DrawingBitmap IsNot Nothing Then
            DrawingBitmap.Dispose()
        End If
        DrawingBitmap = CType(baseBmp.Clone(), Bitmap)
        baseBmp.Dispose()

        ' Quando carichiamo un frame, lo consideriamo pulito
        HasUnsavedChanges = False
        stateStack.Clear()

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
End Class


Public Class FrameNota
    Public Property Testo As String
    Public Property Autore As String
    Public Property Data As DateTime
End Class

