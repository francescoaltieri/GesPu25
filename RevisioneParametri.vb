Public Class RevisioneParametri
    Public Property VideoID As Integer
    Public Property RevisioneID As Integer
    Public Property NomeUtente As String
    Public Property Note As String
    Public Property Stato As String
    Public Property DataRevisione As DateTime
    Public Property Approvato As Boolean

    Public Sub New(videoID As Integer, revisioneID As Integer, nomeUtente As String, note As String, stato As String, dataRevisione As DateTime, Approvato As Boolean)
        Me.VideoID = videoID
        Me.RevisioneID = revisioneID
        Me.NomeUtente = nomeUtente
        Me.Note = note
        Me.Stato = stato
        Me.DataRevisione = dataRevisione
        Me.Approvato = Approvato
    End Sub
End Class


