Imports Microsoft.Data.SqlClient

Module ModuloAutorizzazioni

    Public Function UtenteAutorizzato(formName As String, operazione As String, nomeUtente As String) As Boolean
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim campoPermesso As String = operazione.ToLower()
            Select Case campoPermesso
                Case "view" : campoPermesso = "CanView"
                Case "insert" : campoPermesso = "CanInsert"
                Case "update" : campoPermesso = "CanUpdate"
                Case "delete" : campoPermesso = "CanDelete"
                Case Else : Return False
            End Select

            Dim query As String = $"SELECT {campoPermesso} FROM Tab_UtentiAutorizzazioni WHERE NomeUtente = @NomeUtente AND Form = @Form"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@NomeUtente", nomeUtente)
                cmd.Parameters.AddWithValue("@Form", formName)
                Dim result = cmd.ExecuteScalar()
                Return result IsNot Nothing AndAlso Convert.ToBoolean(result)
            End Using
        End Using
    End Function

End Module
