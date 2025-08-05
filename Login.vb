Imports System.Text
Imports Microsoft.Data.SqlClient

Public Class Login
    Private Sub Login_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Verifica che il form MDI parent esista
        If Me.MdiParent IsNot Nothing Then
            ' Calcola la posizione centrata
            Dim centroX As Integer = (Me.MdiParent.ClientSize.Width - Me.Width) \ 2
            Dim centroY As Integer = (Me.MdiParent.ClientSize.Height - Me.Height) \ 2
            ' Imposta la posizione
            Me.Location = New Point(centroX, centroY)
        End If
    End Sub

    Private Sub BtnAnnulla_Click(sender As Object, e As EventArgs) Handles BtnAnnulla.Click
        Application.Exit()
    End Sub

    Private Sub BtnLogin_Click(sender As Object, e As EventArgs) Handles BtnLogin.Click

        Dim query As String = "SELECT NomeUtente FROM Sys_Utenti WHERE NomeUtente = @nome AND Password = @password"
        Dim Cripta As New CriptaHash
        Dim hashedPassword = Cripta.HashPassword(txtPassword.Text)
        Dim utenteAutenticato As Object = Nothing

        Using conn As New SqlConnection(ConnString)
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@nome", txtNomeUtente.Text)
                cmd.Parameters.AddWithValue("@password", hashedPassword)

                conn.Open()
                utenteAutenticato = cmd.ExecuteScalar() ' 🔹 Recupera il NomeUtente se esiste
            End Using
        End Using

        If utenteAutenticato IsNot Nothing Then
            ' ✅ Memorizza nella sessione
            SessioneUtente.NomeUtenteCorrente = txtNomeUtente.Text
            SessioneUtente.Autorizzazioni = New AutorizzazioniUtente()
            SessioneUtente.Autorizzazioni.Carica(txtNomeUtente.Text)

            Me.Hide()
            AttivaTuttiIMenu(GesPu25.MenuStrip1)
        Else
            Dim risposta = MDIMessageBox.Show("Nome utente o password errati.", Me.MdiParent, MessageBoxButtons.OK)
        End If

    End Sub

    Public Sub AttivaTuttiIMenu(menuBar As MenuStrip)
        For Each voce As ToolStripMenuItem In menuBar.Items
            AttivaMenuItem(voce)
        Next
    End Sub

    Private Sub AttivaMenuItem(item As ToolStripMenuItem)
        item.Enabled = True
        For Each subItem As ToolStripItem In item.DropDownItems
            If TypeOf subItem Is ToolStripMenuItem Then
                AttivaMenuItem(CType(subItem, ToolStripMenuItem))
            End If
        Next
    End Sub

End Class