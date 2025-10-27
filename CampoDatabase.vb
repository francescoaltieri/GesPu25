Public Class CampoDatabase
    Public Property Nome As String
    Public Property Tipo As String
    Public Property MaxLen As Integer
    Public Property IsChiave As Boolean
    Public Property IsIdentity As Boolean
    Public Property TabellaCollegata As String
    Public Property CampoVisuale As String
    Public Property CampoValore As String
    Public Property Lunghezza As String

    ' Proprietà di convalida
    Public Property TipoConvalida As String ' "I" o "E"
    Public Property IntervalloMin As String
    Public Property IntervalloMax As String
    Public Property TabellaElenco As String
    Public Property ChiaveElenco As String
    Public Property DescrizioneChiave As String
    Public Property AbilitaZoom As Boolean
    Public Property AbilitaModifica As Boolean

End Class

