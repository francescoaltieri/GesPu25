
Imports System.IO

Public Module ConfigReader
    Public Function LeggiValoreIni(sezione As String, chiave As String, percorsoFile As String) As String
        Dim linee = File.ReadAllLines(percorsoFile)
        Dim inSezione = False

        For Each linea In linee
            Dim trimmed = linea.Trim()
            If trimmed.StartsWith("[") AndAlso trimmed.EndsWith("]") Then
                inSezione = (trimmed = "[" & sezione & "]")
            ElseIf inSezione AndAlso trimmed.Contains("=") Then
                Dim parts = trimmed.Split("="c)
                If parts(0).Trim() = chiave Then
                    Return parts(1).Trim()
                End If
            End If
        Next

        Return ""
    End Function

End Module

