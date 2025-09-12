Public Class RevisioneParametri
    Public Property VideoID As Integer
    Public Property RevisioneID As Integer
    Public Property Permesso As String
    Public Property NomeUtente As String
    Public Property Note As String
    Public Property Stato As String
    Public Property DataRevisione As DateTime

    Public Sub New(videoID As Integer, revisioneID As Integer, permesso As String,
                   nomeUtente As String, note As String, stato As String, dataRevisione As DateTime)
        Me.VideoID = videoID
        Me.RevisioneID = revisioneID
        Me.Permesso = permesso
        Me.NomeUtente = nomeUtente
        Me.Note = note
        Me.Stato = stato
        Me.DataRevisione = dataRevisione
    End Sub
End Class


