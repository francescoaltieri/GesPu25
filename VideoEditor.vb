Imports System.Drawing
Imports System.IO
Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports System.Globalization

Public Class VideoEditor
    Public Property VideoPath As String
    Public Property FrameDirectory As String
    Public Property FrameList As List(Of String)
    Public Property CurrentIndex As Integer = 0
    Public Property DrawingBitmap As Bitmap
    Public Property UndoStack As New Stack(Of Bitmap)
    Private framerate As Double = 30.0 ' default
    Public FrameNote As New Dictionary(Of Integer, FrameNota)

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


    ''' <summary>
    ''' Estrae i frame dal video senza testo sovrapposto
    ''' </summary>
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

    ''' <summary>
    ''' Carica il frame e sovrappone numero + timestamp
    ''' </summary>
    Public Function LoadFrame(index As Integer) As Bitmap
        If index >= 0 AndAlso index < FrameList.Count Then
            CurrentIndex = index
            Dim path = FrameList(index)

            ' Caricamento sicuro: copia in memoria per evitare lock
            Dim original As Bitmap
            Using fs As New FileStream(path, FileMode.Open, FileAccess.Read)
                Using ms As New MemoryStream()
                    fs.CopyTo(ms)
                    ms.Position = 0
                    original = New Bitmap(ms)
                End Using
            End Using

            Dim bmp = CType(original.Clone(), Bitmap)
            original.Dispose()

            ' Sovrapposizione testo
            Dim ts = TimeSpan.FromSeconds(index / framerate)
            Dim testo = $"Frame: {index + 1}  {ts:hh\:mm\:ss}"

            Using g As Graphics = Graphics.FromImage(bmp)
                Dim font = New Font("Arial", 16, FontStyle.Bold)
                Dim textSize = g.MeasureString(testo, font)
                Dim padding = 6
                Dim rect = New Rectangle(10, 10, CInt(textSize.Width) + padding * 2, CInt(textSize.Height) + padding * 2)
                g.FillRectangle(Brushes.Black, rect)
                g.DrawString(testo, font, Brushes.White, rect.Left + padding, rect.Top + padding)
            End Using

            DrawingBitmap = bmp
            UndoStack.Clear()
            Return CType(DrawingBitmap.Clone(), Bitmap)
        End If
        Return Nothing
    End Function

    ''' <summary>
    ''' Salva il frame modificato
    ''' </summary>
    Public Sub SaveFrame()
        Try
            Dim path = FrameList(CurrentIndex)

            ' Verifica che il bitmap sia valido
            If DrawingBitmap Is Nothing Then
                MessageBox.Show("Nessun frame da salvare.")
                Exit Sub
            End If

            ' Rimuove attributi di sola lettura se presenti
            If File.Exists(path) Then
                File.SetAttributes(path, FileAttributes.Normal)
            End If

            ' Sovrascrive il file immagine
            Using fs As New FileStream(path, FileMode.Create, FileAccess.Write)
                DrawingBitmap.Save(fs, Imaging.ImageFormat.Png)
            End Using

        Catch ex As Exception
            MessageBox.Show("Errore nel salvataggio del frame: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Annulla l'ultima modifica
    ''' </summary>
    Public Sub Undo()
        If UndoStack.Count > 0 Then
            DrawingBitmap = UndoStack.Pop()
        End If
    End Sub

    ''' <summary>
    ''' Salva lo stato corrente per undo
    ''' </summary>
    Public Sub SaveState()
        If DrawingBitmap IsNot Nothing Then
            UndoStack.Push(CType(DrawingBitmap.Clone(), Bitmap))
        End If
    End Sub

    ''' <summary>
    ''' Ricompone il video dai frame modificati
    ''' </summary>
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

    ''' <summary>
    ''' Estrae il framerate reale dal video
    ''' </summary>
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

    Public Class FrameNota
        Public Property Testo As String
        Public Property Autore As String
        Public Property Data As DateTime
    End Class

    Public Sub SalvaNotaCompleta(index As Integer, nota As String, autore As String, data As DateTime)
        FrameNote(index) = New FrameNota With {
        .Testo = nota,
        .Autore = autore,
        .Data = data
    }
    End Sub

    Public Function GetNotaPerFrame(index As Integer) As String
        If FrameNote.ContainsKey(index) Then Return FrameNote(index).Testo
        Return ""
    End Function

    Public Function GetAutorePerFrame(index As Integer) As String
        If FrameNote.ContainsKey(index) Then Return FrameNote(index).Autore
        Return ""
    End Function

    Public Function GetDataNotaPerFrame(index As Integer) As DateTime
        If FrameNote.ContainsKey(index) Then Return FrameNote(index).Data
        Return DateTime.MinValue
    End Function

    Public Function GetFrameAnnotati() As List(Of Integer)
        Return FrameNote.Keys.ToList()
    End Function

    Public Sub RimuoviNota(index As Integer)
        If FrameNote.ContainsKey(index) Then
            FrameNote.Remove(index)
        End If
    End Sub

End Class

Public Class FrameNota
    Public Property Testo As String
    Public Property Autore As String
    Public Property Data As DateTime
End Class

