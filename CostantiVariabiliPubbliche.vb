Public Module CostantiVariabiliPubbliche
    Private ReadOnly percorsoIni As String = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.ini")

    Public ReadOnly ConnString As String = CostruisciConnString()

    Private Function CostruisciConnString() As String
        Dim ip = ConfigReader.LeggiValoreIni("SQLServer", "IP", percorsoIni)
        Dim port = ConfigReader.LeggiValoreIni("SQLServer", "Port", percorsoIni)
        Dim db = ConfigReader.LeggiValoreIni("SQLServer", "Database", percorsoIni)
        Dim user = ConfigReader.LeggiValoreIni("SQLServer", "User", percorsoIni)
        Dim pwd = ConfigReader.LeggiValoreIni("SQLServer", "Password", percorsoIni)

        Return $"Server={ip},{port};Database={db};User Id={user};Password={pwd};TrustServerCertificate=True;"
    End Function
End Module