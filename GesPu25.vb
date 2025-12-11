Imports Microsoft.Data.SqlClient
Imports PdfSharp.Fonts

Public Class GesPu25
    Private Sub GesPu25_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Name = "GesPu25"
        Me.toolStripUser.Text = "Nessun utente connesso"
        Me.toolStripDataOra.Text = "                                       "

        GestioneStatoForm.CaricaStato(Me)

        GlobalFontSettings.UseWindowsFontsUnderWindows = True

        DisattivaVociMenu()

        Login.MdiParent = Me
        Login.Show()

    End Sub

    Public Sub CaricaNotificheUtente(nomeUtente As String)
        ' Configura la ListView
        ListNotifiche.Clear()
        ListNotifiche.View = View.Details
        ListNotifiche.FullRowSelect = True

        ' Definisci le colonne
        ListNotifiche.Columns.Add("Id Notifica", 80)
        ListNotifiche.Columns.Add("Data", 150)
        ListNotifiche.Columns.Add("Messaggio", 300)
        ListNotifiche.Columns.Add("Letto", 60)

        ListNotifiche.Visible = False

        Try
            Dim query As String = "SELECT IdNotifica, Data, Messaggio, Letto 
                                   FROM Tab_Notifiche 
                                   WHERE Destinatario = @utente and Letto = @letto
                                   ORDER BY Data DESC"

            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@utente", nomeUtente)
                    cmd.Parameters.AddWithValue("@letto", False)

                    conn.Open()
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim item As New ListViewItem(reader("IdNotifica").ToString())
                            item.SubItems.Add(Convert.ToDateTime(reader("Data")).ToString("dd/MM/yyyy HH:mm"))
                            item.SubItems.Add(reader("Messaggio").ToString())
                            item.SubItems.Add(If(Convert.ToBoolean(reader("Letto")), "Sì", "No"))
                            ListNotifiche.Items.Add(item)
                        End While
                    End Using
                End Using
            End Using
            If ListNotifiche.Items.Count = 0 Then
                ListNotifiche.Visible = False
            Else
                ListNotifiche.Visible = True
            End If

        Catch ex As Exception
            MessageBox.Show("Errore nel caricamento notifiche: " & ex.Message)
        End Try
    End Sub

    Private Sub ListNotifiche_DoubleClick(sender As Object, e As EventArgs) Handles ListNotifiche.DoubleClick
        If ListNotifiche.SelectedItems.Count > 0 Then
            Dim idNotifica As Integer = Convert.ToInt32(ListNotifiche.SelectedItems(0).Text)

            ' Aggiorna lo stato nel DB
            AggiornaNotificaComeLetta(idNotifica)

            ' Ricarica la lista notifiche per l’utente corrente
            CaricaNotificheUtente(SessioneUtente.NomeUtenteCorrente)
        End If
    End Sub

    Private Sub AggiornaNotificaComeLetta(idNotifica As Integer)
        Try
            Dim query As String = "UPDATE Tab_Notifiche SET Letto = 1 WHERE IdNotifica = @id"

            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@id", idNotifica)
                    conn.Open()
                    cmd.ExecuteNonQuery()
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show("Errore durante l'aggiornamento della notifica: " & ex.Message)
        End Try
    End Sub


    Private Sub LogoutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LogoutToolStripMenuItem.Click
        ' Cancella la sessione utente
        SessioneUtente.NomeUtenteCorrente = Nothing
        SessioneUtente.Autorizzazioni = Nothing

        ListNotifiche.Visible = False

        Me.toolStripUser.Text = "Nessun utente connesso"
        Me.toolStripDataOra.Text = "                                       "

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
        infoForm.MdiParent = Me
        infoForm.Show()
    End Sub

    Private Sub ImportaDaExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportaDaExcelToolStripMenuItem.Click
        Dim ImportForm As New ImportaExcel()

        ImportForm.MdiParent = Me

        ImportForm.Show()
    End Sub

    Public Sub ApriModuloConPermessi(nomeTabella As String, mdiParent As Form)

        Cursor.Current = Cursors.WaitCursor
        Application.DoEvents()

        Dim permessi = SessioneUtente.Autorizzazioni.GetPermessi(nomeTabella)
        Dim isAdmin = IsUtenteAdmin(SessioneUtente.NomeUtenteCorrente)

        If Not permessi.CanView AndAlso Not isAdmin Then
            MDIMessageBox.Show($"Accesso negato al modulo {nomeTabella}.", mdiParent, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        For Each f As Form In mdiParent.MdiChildren
            If TypeOf f Is DynamicDataForm Then
                Dim ft As String = If(f.Text, "").Trim()
                Dim target As String = If(nomeTabella, "").Trim()

                ' esatto 
                If String.Equals(ft, target, StringComparison.OrdinalIgnoreCase) Then
                    f.Activate()
                    Return
                End If

                ' formato "Modulo: Nomeform" (o simili "Chiave: Valore") -> prendi la parte dopo i due punti
                Dim parts = ft.Split(New Char() {":"c}, 2)
                If parts.Length = 2 Then
                    Dim right = parts(1).Trim()
                    If String.Equals(right, target, StringComparison.OrdinalIgnoreCase) Then
                        f.Activate()
                        Return
                    End If
                End If

            End If
        Next

        Dim campi = RecuperaCampiDa(nomeTabella)
        Dim nuovoForm As New DynamicDataForm(campi, nomeTabella)
        nuovoForm.Text = $"Modulo: {nomeTabella}"
        nuovoForm.MdiParent = mdiParent
        nuovoForm.Show()

        Cursor.Current = Cursors.Default
        Application.DoEvents()

    End Sub

    Public Sub ApriModulo2ConPermessi(nomeTabella As String, pForm As Form)

        Cursor.Current = Cursors.WaitCursor
        Application.DoEvents()

        Dim permessi = SessioneUtente.Autorizzazioni.GetPermessi(nomeTabella)
        Dim isAdmin = IsUtenteAdmin(SessioneUtente.NomeUtenteCorrente) ' ← uso della funzione helper

        If Not permessi.CanView AndAlso Not isAdmin Then
            MDIMessageBox.Show($"Accesso negato al modulo {nomeTabella}.", MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        For Each f As Form In Me.MdiChildren
            If TypeOf f Is DynamicDataForm AndAlso f.Text.Contains(nomeTabella) Then
                f.Activate()
                Return
            End If
        Next

        pForm.MdiParent = Me
        pForm.Show()

        Cursor.Current = Cursors.Default

    End Sub

    Public Function IsUtenteAdmin(nomeUtente As String) As Boolean
        Try
            Using conn As New SqlConnection(ConnString)
                conn.Open()

                Dim query = "SELECT ISNULL(Amministratore, 0) FROM Tab_Utenti WHERE NomeUtente = @utente"
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
    Private Sub GesPu25_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        GestioneStatoForm.SalvaStato(Me)
    End Sub

    Private Sub VideoFbFToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VideoFbFToolStripMenuItem.Click
        Dim modulo As New VideoFBF()
        ApriModulo2ConPermessi("VideoFBF", modulo)
    End Sub

    Private Sub OggettiModelPackToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OggettiModelPackToolStripMenuItem.Click
        Dim modulo As New CaricaOggettiModelPack()
        ApriModulo2ConPermessi("CaricaOggettiModelPack", modulo)
    End Sub

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
    Private Sub GestioneUtentiToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GestioneUtentiToolStripMenuItem.Click
        ApriModuloConPermessi("Tab_Utenti", Me)
    End Sub

    Private Sub FileTemplateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FileTemplateToolStripMenuItem.Click
        ApriModuloConPermessi("Sys_template", Me)
    End Sub
    Private Sub FornitoriToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles FornitoriToolStripMenuItem2.Click
        ApriModuloConPermessi("Tab_Fornitori", Me)
    End Sub

    Private Sub EMailToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ApriModuloConPermessi("Tab_Comunicazioni", Me)
    End Sub

    Private Sub ConvalidaCampiToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConvalidaCampiToolStripMenuItem.Click
        ApriModuloConPermessi("Sys_ConvalidaCampi", Me)
    End Sub

    Private Sub ContrattiToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ContrattiToolStripMenuItem.Click
        ApriModuloConPermessi("Tab_Contratti", Me)
    End Sub

    Private Sub EsciToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EsciToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub EpisodiToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EpisodiToolStripMenuItem.Click
        ApriModuloConPermessi("Tab_Episodi", Me)
    End Sub

    Private Sub CharactersToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ApriModuloConPermessi("Tab_Characters", Me)
    End Sub

    Private Sub PropsToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ApriModuloConPermessi("Tab_Props", Me)
    End Sub

    Private Sub EffettiToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ApriModuloConPermessi("Tab_Effetti", Me)
    End Sub

    Private Sub TipoLavorazioniToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TipoLavorazioniToolStripMenuItem.Click
        ApriModuloConPermessi("Tab_Lavorazioni", Me)
    End Sub

    Private Sub StatoLavorazioniToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StatoLavorazioniToolStripMenuItem.Click
        ApriModuloConPermessi("Tab_StatoLavorazioni", Me)
    End Sub

    Private Sub StatoAssegnazioniToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StatoAssegnazioniToolStripMenuItem.Click
        ApriModuloConPermessi("Tab_StatoAssegnazioni", Me)
    End Sub

    Private Sub LavorazioniNonIdentificateToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ApriModuloConPermessi("Tab_LavNonIdentificate", Me)
    End Sub
    Private Sub OggettiLavorazioneToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OggettiLavorazioneToolStripMenuItem.Click
        ApriModuloConPermessi("Tab_OggettiLavorazione", Me)
    End Sub
    Private Sub ComunicazioniToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ComunicazioniToolStripMenuItem.Click
        ApriModuloConPermessi("Tab_Comunicazioni", Me)
    End Sub

    Private Sub SceneAssegnateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SceneAssegnateToolStripMenuItem.Click
        ApriModuloConPermessi("Mov_ConsegneScene", Me)
    End Sub

    Private Sub AutorizzazioniScenaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AutorizzazioniScenaToolStripMenuItem.Click
        ApriModuloConPermessi("Mov_SceneUtente", Me)
    End Sub

    Private Sub ListaRevisioniToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ListaRevisioniToolStripMenuItem.Click
        ApriModuloConPermessi("Mov_Revisioni", Me)
    End Sub

    Private Sub AutorizzazioniRevisioneToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AutorizzazioniRevisioneToolStripMenuItem.Click
        ApriModuloConPermessi("Mov_RevisioniUtente", Me)
    End Sub

    Private Sub AnnotazioniSuiFramesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AnnotazioniSuiFramesToolStripMenuItem.Click
        ApriModuloConPermessi("Mov_FrameNote", Me)
    End Sub

    Private Sub StoryboardToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles StoryboardToolStripMenuItem1.Click
        ApriModuloConPermessi("Mov_Storyboard", Me)
    End Sub

    Private Sub SceneToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SceneToolStripMenuItem.Click
        ApriModuloConPermessi("Mov_StoryboardScene", Me)
    End Sub

    Private Sub PanelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PanelToolStripMenuItem.Click
        ApriModuloConPermessi("Mov_StoryboardScenePanel", Me)
    End Sub


    Private Sub AssegnazioniToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles AssegnazioniToolStripMenuItem2.Click
        ApriModuloConPermessi("Mov_Assegnazioni", Me)
    End Sub

    Private Sub AssegnazioniAnimazioniToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AssegnazioniAnimazioniToolStripMenuItem.Click
        ApriModuloConPermessi("Mov_AssegnazioniLavA", Me)
    End Sub

    Private Sub AssegnazioniDiverseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AssegnazioniDiverseToolStripMenuItem.Click
        ApriModuloConPermessi("Mov_AssegnazioniLavD", Me)
    End Sub

    Private Sub ConsegneAnimazioniToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConsegneAnimazioniToolStripMenuItem.Click
        ApriModuloConPermessi("Mov_ConsegneLavA", Me)
    End Sub

    Private Sub ConsegneDiverseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConsegneDiverseToolStripMenuItem.Click
        ApriModuloConPermessi("Mov_ConsegneLavD", Me)
    End Sub

    Private Sub OggettiStoryboardToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OggettiStoryboardToolStripMenuItem.Click
        ApriModuloConPermessi("Tab_OggettiLavorazione", Me)
    End Sub

    Private Sub ModelPackToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ModelPackToolStripMenuItem1.Click
        ApriModuloConPermessi("Mov_ModelPack", Me)
    End Sub

    Private Sub VociModelPackToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VociModelPackToolStripMenuItem.Click
        ApriModuloConPermessi("Mov_ModelPackOggetti", Me)
    End Sub

    Private Sub AcquisisciDaPDFToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AcquisisciDaPDFToolStripMenuItem.Click
        Dim modulo As New PDF2Storyboard()
        ApriModulo2ConPermessi("Acquisici da PDF", modulo)
    End Sub

    Private Sub GestioneNotificheToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GestioneNotificheToolStripMenuItem.Click
        ApriModuloConPermessi("Tab_Notifiche", Me)
    End Sub

End Class

