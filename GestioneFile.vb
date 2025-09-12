Imports System.IO

Module GestioneFile
    Public Sub CreaStrutturaVideo(videoID As Integer)
        Dim basePath As String = Path.Combine("C:\VideoEditor\Frames", $"Video_{videoID}")
        Dim revisioneZeroPath As String = Path.Combine(basePath, "Revisione_0")

        If Not Directory.Exists(basePath) Then Directory.CreateDirectory(basePath)
        If Not Directory.Exists(revisioneZeroPath) Then Directory.CreateDirectory(revisioneZeroPath)

        ' Flag software per sola lettura
        Dim flagPath As String = Path.Combine(revisioneZeroPath, "readonly.flag")
        File.WriteAllText(flagPath, "true")
    End Sub
End Module

