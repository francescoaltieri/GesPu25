Imports Microsoft.Data.SqlClient
Imports ModuloCampiDinamici
Imports PdfSharp.Pdf
Imports PdfSharp.Drawing
Imports PdfSharp.Fonts
Imports System.Text.RegularExpressions

Public Class GesPu25
    Private Sub GesPu25_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Name = "GesPu25"
        GestioneStatoForm.CaricaStato(Me)

        GlobalFontSettings.UseWindowsFontsUnderWindows = True

        DisattivaVociMenu()

        Login.MdiParent = Me
        Login.Show()
    End Sub

    Private Sub LogoutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LogoutToolStripMenuItem.Click
        ' Cancella la sessione utente
        SessioneUtente.NomeUtenteCorrente = Nothing
        SessioneUtente.Autorizzazioni = Nothing

        ' Chiude i form MDI aperti (escluso GesPu25)
        For Each frm As Form In Me.MdiChildren
            frm.Close()
        Next

        DisattivaVociMenu()

        LogoutToolStripMenuItem.Enabled = True

        ' Mostra nuovamente il form Login
        Dim loginForm As New Login()
        loginForm.MdiParent = Me
        loginForm.Show()

    End Sub

    Private Sub DisattivaVociMenu()
        For Each voce As ToolStripItem In MenuStrip1.Items
            If TypeOf voce Is ToolStripMenuItem Then
                Dim menuItem = CType(voce, ToolStripMenuItem)
                menuItem.Enabled = False
            End If
        Next
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        Dim infoForm As New InformazioniApp()

        ' Se vuoi aprirlo come finestra figlia all'interno del form principale
        infoForm.MdiParent = Me

        infoForm.Show()
    End Sub

    Private Sub ImportaDaExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportaDaExcelToolStripMenuItem.Click
        Dim ImportForm As New ImportaExcel()

        ' Se vuoi aprirlo come finestra figlia all'interno del form principale
        ImportForm.MdiParent = Me

        ImportForm.Show()
    End Sub

    Public Sub ApriModuloConPermessi(nomeTabella As String, mdiParent As Form)
        Dim permessi = SessioneUtente.Autorizzazioni.GetPermessi(nomeTabella)
        Dim isAdmin = IsUtenteAdmin(SessioneUtente.NomeUtenteCorrente) ' ← uso della funzione helper

        If Not permessi.CanView AndAlso Not isAdmin Then
            MDIMessageBox.Show($"Accesso negato al modulo {nomeTabella}.", mdiParent, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        For Each f As Form In mdiParent.MdiChildren
            If TypeOf f Is DynamicDataForm AndAlso f.Text.Contains(nomeTabella) Then
                f.Activate()
                Return
            End If
        Next

        Dim campi = RecuperaCampiDa(nomeTabella)
        Dim nuovoForm As New DynamicDataForm(campi, nomeTabella)
        nuovoForm.Text = $"Modulo: {nomeTabella}"
        nuovoForm.MdiParent = mdiParent
        nuovoForm.Show()
    End Sub

    Public Function IsUtenteAdmin(nomeUtente As String) As Boolean
        Try
            Using conn As New SqlConnection(ConnString)
                conn.Open()

                Dim query = "SELECT ISNULL(Amministratore, 0) FROM Sys_Utenti WHERE NomeUtente = @utente"
                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@utente", nomeUtente)
                    Return Convert.ToBoolean(cmd.ExecuteScalar())
                End Using
            End Using
        Catch ex As Exception
            MDIMessageBox.Show("Errore nel controllo amministratore: " & ex.Message, Nothing, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Private Sub ParametriToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ParametriToolStripMenuItem1.Click
        ApriModuloConPermessi("Sys_Parametri", Me)
    End Sub

    Private Sub GestioneFormToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GestioneFormToolStripMenuItem.Click
        ApriModuloConPermessi("Sys_Form", Me)
    End Sub

    Private Sub GestioneFormCollegatiToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GestioneFormCollegatiToolStripMenuItem.Click
        ApriModuloConPermessi("Sys_Form_Actions", Me)
    End Sub

    Private Sub TestoEtichetteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TestoEtichetteToolStripMenuItem.Click
        ApriModuloConPermessi("Sys_TestoEtichetta", Me)
    End Sub

    Private Sub GestioneGriglieToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GestioneGriglieToolStripMenuItem.Click
        ApriModuloConPermessi("Sys_VisualizzainDbgrid", Me)
    End Sub

    Private Sub FornitoriToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles FornitoriToolStripMenuItem2.Click
        ApriModuloConPermessi("Tab_Fornitori", Me)
    End Sub

    Private Sub EMailToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EMailToolStripMenuItem.Click
        ApriModuloConPermessi("Tab_EMail", Me)
    End Sub

    Private Sub ContrattiToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ContrattiToolStripMenuItem.Click
        ApriModuloConPermessi("Tab_Contratti", Me)
    End Sub
    Private Sub GesPu25_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        GestioneStatoForm.SalvaStato(Me)
    End Sub

    Private Sub EsciToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EsciToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub GestioneUtentiToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GestioneUtentiToolStripMenuItem.Click
        ApriModuloConPermessi("Sys_Utenti", Me)
    End Sub

    Private Sub EpisodiToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EpisodiToolStripMenuItem.Click
        ApriModuloConPermessi("Tab_Episodi", Me)
    End Sub

    Private Sub BackgroundToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles BackgroundToolStripMenuItem1.Click
        ApriModuloConPermessi("Tab_Backgrounds", Me)
    End Sub

    Private Sub CharactersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CharactersToolStripMenuItem.Click
        ApriModuloConPermessi("Tab_Characters", Me)
    End Sub

    Private Sub PropsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PropsToolStripMenuItem.Click
        ApriModuloConPermessi("Tab_Props", Me)
    End Sub

    Private Sub EffettiToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EffettiToolStripMenuItem.Click
        ApriModuloConPermessi("Tab_Effetti", Me)
    End Sub

    Private Sub TipoLavorazioniToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TipoLavorazioniToolStripMenuItem.Click
        ApriModuloConPermessi("Tab_TipoLavorazioneMS", Me)
    End Sub

    Private Sub StatoLavorazioniToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StatoLavorazioniToolStripMenuItem.Click
        ApriModuloConPermessi("Tab_StatoLavorazioni", Me)
    End Sub

    Private Sub StatoAssegnazioniToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StatoAssegnazioniToolStripMenuItem.Click
        ApriModuloConPermessi("Tab_StatoAssegnazioni", Me)
    End Sub

    Private Sub LavorazioniNonIdentificateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LavorazioniNonIdentificateToolStripMenuItem.Click
        ApriModuloConPermessi("Tab_LavNonIdentificate", Me)
    End Sub
End Class

