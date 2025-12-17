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
        Try
            Dim query As String = "SELECT NomeUtente FROM Tab_Utenti WHERE NomeUtente = @nome AND Password = @password AND IsActive = @Active"
            Dim Cripta As New CriptaHash
            Dim hashedPassword = Cripta.HashPassword(txtPassword.Text)
            Dim utenteAutenticato As Object = Nothing

            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@nome", txtNomeUtente.Text)
                    cmd.Parameters.AddWithValue("@password", hashedPassword)
                    cmd.Parameters.AddWithValue("@Active", True)

                    conn.Open()
                    utenteAutenticato = cmd.ExecuteScalar()
                End Using
            End Using

            If utenteAutenticato IsNot Nothing Then
                ' Imposta la sessione utente
                SessioneUtente.NomeUtenteCorrente = txtNomeUtente.Text
                SessioneUtente.Autorizzazioni = New AutorizzazioniUtente()
                SessioneUtente.Autorizzazioni.Carica(txtNomeUtente.Text)
                SessioneUtente.DataConnessione = Now

                ' Aggiorna l'interfaccia del MDI parent se disponibile
                If Me.MdiParent IsNot Nothing AndAlso TypeOf Me.MdiParent Is GesPu25 Then
                    Dim parentForm = CType(Me.MdiParent, GesPu25)

                    ' Aggiorna i toolstrip in modo sicuro (Invoke se necessario)
                    If parentForm.InvokeRequired Then
                        parentForm.Invoke(Sub()
                                              parentForm.toolStripUser.Text = "Utente: " & SessioneUtente.NomeUtenteCorrente
                                              parentForm.toolStripDataOra.Text = "Connesso dalle " & SessioneUtente.DataConnessione.ToString("HH:mm - dd/MM/yyyy")
                                              AttivaTuttiIMenu(parentForm.MenuStrip1)
                                          End Sub)
                    Else
                        parentForm.toolStripUser.Text = "Utente: " & SessioneUtente.NomeUtenteCorrente
                        parentForm.toolStripDataOra.Text = "Connesso dalle " & SessioneUtente.DataConnessione.ToString("HH:mm - dd/MM/yyyy")
                        AttivaTuttiIMenu(parentForm.MenuStrip1)
                    End If

                    ' Apri le notifiche
                    Try
                        parentForm.ApriNotifiche(False)
                    Catch ex As Exception
                        MDIMessageBox.Show($"Errore apertura notifiche: {ex.Message}", parentForm, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End Try
                Else
                    GesPu25.toolStripUser.Text = "Utente: " & SessioneUtente.NomeUtenteCorrente
                    GesPu25.toolStripDataOra.Text = "Connesso dalle " & SessioneUtente.DataConnessione.ToString("HH:mm - dd/MM/yyyy")
                    AttivaTuttiIMenu(GesPu25.MenuStrip1)
                End If

                ' Chiudi il form di login
                Me.Close()
            Else
                MDIMessageBox.Show("Nome utente o password errati.", Me.MdiParent, MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            MDIMessageBox.Show("Errore durante l'autenticazione: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
        End Try
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