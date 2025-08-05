Imports Microsoft.Data.SqlClient
Public Class AutorizzazioniUtente
    Private autorizzazioni As New Dictionary(Of String, PermessiForm)

    Public Sub Carica(nomeUtente As String)
        autorizzazioni.Clear()

        Dim query = "SELECT Form, CanView, CanInsert, CanUpdate, CanDelete FROM Sys_Autorizzazioni WHERE NomeUtente = @nome"

        Using conn As New SqlConnection(ConnString)
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@nome", nomeUtente)
                conn.Open()
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim formName = reader.GetString(0)
                        Dim permessi As New PermessiForm With {
                            .CanView = reader.GetBoolean(1),
                            .CanInsert = reader.GetBoolean(2),
                            .CanUpdate = reader.GetBoolean(3),
                            .CanDelete = reader.GetBoolean(4)
                        }
                        autorizzazioni(formName) = permessi
                    End While
                End Using
            End Using
        End Using
    End Sub

    Public Function GetPermessi(formName As String) As PermessiForm
        If autorizzazioni.ContainsKey(formName) Then
            Return autorizzazioni(formName)
        Else
            Return New PermessiForm() ' nessun permesso per default
        End If
    End Function
End Class


