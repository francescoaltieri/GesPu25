Imports System.ComponentModel
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
            Dim query As String = "SELECT NomeUtente, CambioPassword FROM Tab_Utenti WHERE NomeUtente = @nome AND Password = @password AND IsActive = @Active"
            Dim Cripta As New CriptaHash
            Dim hashedPassword = Cripta.HashPassword(txtPassword.Text)
            Dim nomeTrovato As String = Nothing
            Dim cambioPwdObj As Object = Nothing

            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@nome", txtNomeUtente.Text)
                    cmd.Parameters.AddWithValue("@password", hashedPassword)
                    cmd.Parameters.AddWithValue("@Active", True)

                    conn.Open()
                    Using reader = cmd.ExecuteReader()
                        If reader.Read() Then
                            nomeTrovato = reader("NomeUtente").ToString()
                            cambioPwdObj = If(IsDBNull(reader("CambioPassword")), Nothing, reader("CambioPassword"))
                        End If
                    End Using
                End Using
            End Using

            If nomeTrovato IsNot Nothing Then
                ' Se CambioPassword è True, apri la form di cambio password prima di proseguire
                Dim mustChange As Boolean = False
                If cambioPwdObj IsNot Nothing Then
                    Try
                        mustChange = Convert.ToBoolean(cambioPwdObj)
                    Catch
                        mustChange = False
                    End Try
                End If

                If mustChange Then
                    Using cpForm As New ChangePasswordForm(txtNomeUtente.Text)
                        cpForm.StartPosition = FormStartPosition.CenterParent
                        If cpForm.ShowDialog(Me) <> DialogResult.OK Then
                            ' L'utente ha annullato il cambio password: non proseguire con il login
                            MDIMessageBox.Show("Devi cambiare la password per continuare.", Me.MdiParent, MessageBoxButtons.OK)
                            Return
                        End If
                    End Using
                End If

                ' Imposta la sessione utente
                SessioneUtente.NomeUtenteCorrente = txtNomeUtente.Text
                SessioneUtente.Autorizzazioni = New AutorizzazioniUtente()
                SessioneUtente.Autorizzazioni.Carica(txtNomeUtente.Text)
                SessioneUtente.DataConnessione = Now

                ' Aggiorna l'interfaccia del MDI parent se disponibile
                If Me.MdiParent IsNot Nothing AndAlso TypeOf Me.MdiParent Is GesPu25 Then
                    Dim parentForm = CType(Me.MdiParent, GesPu25)

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

Public Class ChangePasswordForm
    Inherits Form

    Private ReadOnly _username As String

    ' Controlli (se li crei via designer, assicurati che i nomi coincidano)
    Private txtNewPassword As New TextBox With {.UseSystemPasswordChar = True, .Width = 250}
    Private txtConfirmPassword As New TextBox With {.UseSystemPasswordChar = True, .Width = 250}
    Private btnOK As New Button With {.Text = "OK", .DialogResult = DialogResult.None}
    Private btnCancel As New Button With {.Text = "Annulla", .DialogResult = DialogResult.Cancel}
    Private lblMessage As New Label With {.ForeColor = Color.Red, .AutoSize = True}

    Public Sub New(username As String)
        _username = username
        InitializeComponentCustom()
    End Sub

    Private Sub InitializeComponentCustom()
        Me.Text = "Cambia password"
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.StartPosition = FormStartPosition.CenterParent
        Me.ClientSize = New Size(420, 260) ' spazio per statusstrip

        ' Label e textbox per nuova password
        Dim lbl1 As New Label With {
        .Text = "Nuova password:",
        .AutoSize = True,
        .Location = New Point(12, 12)
    }
        txtNewPassword.Location = New Point(12, 32)
        txtNewPassword.Width = 380

        ' Label e textbox per conferma
        Dim lbl2 As New Label With {
        .Text = "Conferma password:",
        .AutoSize = True,
        .Location = New Point(12, 72)
    }
        txtConfirmPassword.Location = New Point(12, 92)
        txtConfirmPassword.Width = 380

        ' Label per messaggi (opzionale, tenuta ma non usata per errori principali)
        lblMessage.Location = New Point(12, 132)
        lblMessage.AutoSize = False
        lblMessage.Size = New Size(380, 28)
        lblMessage.ForeColor = Color.DarkRed
        lblMessage.BackColor = SystemColors.Control
        lblMessage.TextAlign = ContentAlignment.MiddleLeft
        lblMessage.Visible = False

        ' Pulsanti spostati più in basso
        btnOK.Location = New Point(220, 170)
        btnOK.Size = New Size(80, 28)
        btnCancel.Location = New Point(310, 170)
        btnCancel.Size = New Size(80, 28)

        AddHandler btnOK.Click, AddressOf BtnOK_Click
        AddHandler btnCancel.Click, AddressOf BtnCancel_Click

        ' StatusStrip per messaggi di errore sotto i pulsanti
        Dim statusStrip As New StatusStrip()
        statusStrip.Name = "statusStripErrors"
        statusStrip.Dock = DockStyle.Bottom
        statusStrip.SizingGrip = False

        Dim statusLabel As New ToolStripStatusLabel()
        statusLabel.Name = "toolStripStatusLabelError"
        statusLabel.ForeColor = Color.DarkRed
        statusLabel.Text = String.Empty
        statusLabel.Spring = False

        statusStrip.Items.Add(statusLabel)

        ' Aggiungo i controlli al form
        Me.Controls.Add(lbl1)
        Me.Controls.Add(txtNewPassword)
        Me.Controls.Add(lbl2)
        Me.Controls.Add(txtConfirmPassword)
        Me.Controls.Add(lblMessage)
        Me.Controls.Add(btnOK)
        Me.Controls.Add(btnCancel)
        Me.Controls.Add(statusStrip)

        ' Porta la statusstrip in primo piano e assicura visibilità
        statusStrip.BringToFront()
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As EventArgs)
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub BtnOK_Click(sender As Object, e As EventArgs)
        ' Recupera riferimento alla status label
        Dim statusLabel As ToolStripStatusLabel = Nothing
        For Each ctl As Control In Me.Controls
            If TypeOf ctl Is StatusStrip Then
                Dim ss = DirectCast(ctl, StatusStrip)
                For Each it As ToolStripItem In ss.Items
                    If TypeOf it Is ToolStripStatusLabel AndAlso it.Name = "toolStripStatusLabelError" Then
                        statusLabel = DirectCast(it, ToolStripStatusLabel)
                        Exit For
                    End If
                Next
                If statusLabel IsNot Nothing Then Exit For
            End If
        Next

        ' Helper per mostrare errore
        Dim SubShowError = Sub(msg As String)
                               If statusLabel IsNot Nothing Then
                                   statusLabel.Text = msg
                                   statusLabel.ForeColor = Color.DarkRed
                                   ' Assicura che la StatusStrip sia visibile e aggiornata
                                   Dim ssParent As StatusStrip = Nothing
                                   For Each ctl As Control In Me.Controls
                                       If TypeOf ctl Is StatusStrip Then
                                           ssParent = DirectCast(ctl, StatusStrip)
                                           Exit For
                                       End If
                                   Next
                                   If ssParent IsNot Nothing Then
                                       ssParent.Visible = True
                                       ssParent.BringToFront()
                                   End If
                                   statusLabel.Owner.Refresh()
                               Else
                                   ' Fallback alla label interna se statusLabel non trovato
                                   lblMessage.Text = msg
                                   lblMessage.Visible = True
                                   lblMessage.BringToFront()
                                   lblMessage.Refresh()
                               End If
                           End Sub

        lblMessage.Visible = False
        If statusLabel IsNot Nothing Then
            ' nascondi testo precedente
            statusLabel.Text = String.Empty
        End If

        Dim newPwd = txtNewPassword.Text
        Dim confirmPwd = txtConfirmPassword.Text

        If newPwd <> confirmPwd Then
            SubShowError("Le password non corrispondono.")
            Return
        End If

        Dim err As String = String.Empty
        If Not ValidatePasswordComplexity(newPwd, err) Then
            SubShowError(err)
            Return
        End If

        ' Hash della nuova password e update DB
        Try
            Dim cripta As New CriptaHash
            Dim hashed = cripta.HashPassword(newPwd)

            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand("UPDATE Tab_Utenti SET Password = @Password, CambioPassword = 0 WHERE NomeUtente = @NomeUtente", conn)
                    cmd.Parameters.AddWithValue("@Password", hashed)
                    cmd.Parameters.AddWithValue("@NomeUtente", _username)
                    conn.Open()
                    Dim rows = cmd.ExecuteNonQuery()
                    If rows = 0 Then
                        SubShowError("Impossibile aggiornare la password. Riprovare.")
                        Return
                    End If
                End Using
            End Using

            ' Successo: pulisco messaggi e chiudo
            If statusLabel IsNot Nothing Then statusLabel.Text = String.Empty
            lblMessage.Text = String.Empty
            lblMessage.Visible = False

            Me.DialogResult = DialogResult.OK
            Me.Close()
        Catch ex As Exception
            SubShowError("Errore durante l'aggiornamento: " & ex.Message)
        End Try
    End Sub


    Private Function ValidatePasswordComplexity(password As String, ByRef errorMsg As String) As Boolean
        errorMsg = String.Empty
        If String.IsNullOrEmpty(password) Then
            errorMsg = "La password non può essere vuota."
            Return False
        End If

        If password.Length < 10 Then
            errorMsg = "La password deve contenere almeno 10 caratteri."
            Return False
        End If

        ' Almeno una cifra
        If Not System.Text.RegularExpressions.Regex.IsMatch(password, "\d") Then
            errorMsg = "La password deve contenere almeno una cifra."
            Return False
        End If

        ' Almeno un carattere speciale (non lettera né cifra)
        If Not System.Text.RegularExpressions.Regex.IsMatch(password, "[^a-zA-Z0-9]") Then
            errorMsg = "La password deve contenere almeno un carattere speciale (es. !@#%&*)."
            Return False
        End If

        Return True
    End Function

End Class
