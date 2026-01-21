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

    Private Sub LogoutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LogoutToolStripMenuItem.Click
        ' Cancella la sessione utente
        SessioneUtente.NomeUtenteCorrente = Nothing
        SessioneUtente.Autorizzazioni = Nothing

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
        ' Cerca se il form è già aperto come MDI child
        For Each f As Form In Me.MdiChildren
            If TypeOf f Is InformazioniApp Then
                f.BringToFront()
                f.Focus()
                Return
            End If
        Next

        ' Se non esiste, lo crea
        Dim infoForm As New InformazioniApp()
        infoForm.MdiParent = Me
        infoForm.Show()
    End Sub


    Private Sub ImportaDaExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportaDaExcelToolStripMenuItem.Click
        Dim ImportForm As New ImportaExcel()

        ImportForm.MdiParent = Me

        ImportForm.Show()
    End Sub

    Public Sub ApriModuloConPermessi(nomeTabella As String, mdiParent As Form, Optional pFiltroIniziale As String = "")
        Cursor.Current = Cursors.WaitCursor
        Application.DoEvents()

        Dim permessi = SessioneUtente.Autorizzazioni.GetPermessi(nomeTabella)
        Dim isAdmin = IsUtenteAdmin(SessioneUtente.NomeUtenteCorrente)

        If Not permessi.CanView AndAlso Not isAdmin Then
            MDIMessageBox.Show($"Accesso negato al modulo {nomeTabella}.", mdiParent, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Cerca form già aperto e, se trovato, applica il filtro e ricarica
        For Each f As Form In mdiParent.MdiChildren
            If TypeOf f Is DynamicDataForm Then
                Dim ft As String = If(f.Text, "").Trim()
                Dim target As String = If(nomeTabella, "").Trim()

                Dim matchesTarget As Boolean = False
                If String.Equals(ft, target, StringComparison.OrdinalIgnoreCase) Then
                    matchesTarget = True
                Else
                    Dim parts = ft.Split(New Char() {":"c}, 2)
                    If parts.Length = 2 Then
                        Dim right = parts(1).Trim()
                        If String.Equals(right, target, StringComparison.OrdinalIgnoreCase) Then
                            matchesTarget = True
                        End If
                    End If
                End If

                If matchesTarget Then
                    ' Applica filtro se fornito
                    If Not String.IsNullOrEmpty(pFiltroIniziale) Then
                        Try
                            Dim prop = f.GetType().GetProperty("FiltroIniziale")
                            If prop IsNot Nothing AndAlso prop.CanWrite Then
                                prop.SetValue(f, pFiltroIniziale)
                            Else
                                f.Tag = pFiltroIniziale
                            End If

                            ' Prova a invocare un metodo di ricarica dati esposto dal DynamicDataForm
                            Dim reloadNames = New String() {"RicaricaDati", "CaricaDati", "ApplicaFiltroIniziale", "RefreshData", "RefreshGrid"}
                            For Each NN In reloadNames
                                Dim mi = f.GetType().GetMethod(NN, Reflection.BindingFlags.Instance Or Reflection.BindingFlags.Public Or Reflection.BindingFlags.NonPublic)
                                If mi IsNot Nothing Then
                                    mi.Invoke(f, Nothing)
                                    Exit For
                                End If
                            Next
                        Catch ex As Exception
                            MDIMessageBox.Show($"Impossibile applicare filtro al modulo aperto: {ex.Message}", mdiParent, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End Try
                    End If

                    f.Activate()
                    Cursor.Current = Cursors.Default
                    Return
                End If
            End If
        Next

        ' Non è aperto: crealo, imposta il filtro e mostra
        Dim campi = RecuperaCampiDa(nomeTabella)
        Dim nuovoForm As New DynamicDataForm(campi, nomeTabella)
        nuovoForm.Text = $"Modulo: {nomeTabella}"

        If Not String.IsNullOrEmpty(pFiltroIniziale) Then
            Try
                Dim prop = nuovoForm.GetType().GetProperty("FiltroIniziale")
                If prop IsNot Nothing AndAlso prop.CanWrite Then
                    prop.SetValue(nuovoForm, pFiltroIniziale)
                Else
                    nuovoForm.Tag = pFiltroIniziale
                End If
            Catch
                ' fallback silenzioso
            End Try
        End If

        nuovoForm.MdiParent = mdiParent
        nuovoForm.Show()

        ' opzionale: forzare la ricarica immediata se il form espone il metodo
        Try
            Dim mi = nuovoForm.GetType().GetMethod("RicaricaDati", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.Public Or Reflection.BindingFlags.NonPublic)
            If mi IsNot Nothing Then mi.Invoke(nuovoForm, Nothing)
        Catch
            ' ignorare errori non critici
        End Try

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
            MDIMessageBox.Show($"Errore nel controllo amministratore: {ex.Message}", Nothing, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Private Function HasNotificheDaVisualizzare(nomeUtente As String) As Boolean
        Try
            Dim query As String = "
            SELECT COUNT(1)
            FROM Mov_Notifiche
            WHERE Destinatario = @utente AND Letto = 0"

            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@utente", nomeUtente)
                    conn.Open()
                    Dim countObj = cmd.ExecuteScalar()
                    Dim count As Integer = If(IsDBNull(countObj), 0, Convert.ToInt32(countObj))
                    Return count > 0
                End Using
            End Using
        Catch ex As Exception
            MDIMessageBox.Show($"Errore nel controllo notifiche: {ex.Message}", Me, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ' helper chiamato dopo autenticazione
    Public Sub ApriNotifiche(Optional TuttiIRecord As Boolean = True)
        If String.IsNullOrEmpty(SessioneUtente.NomeUtenteCorrente) Then Return
        If TuttiIRecord <> True Then
            If Not HasNotificheDaVisualizzare(SessioneUtente.NomeUtenteCorrente) Then
                Return
            End If
        End If
        Dim ParteFiltro As String = ""

        If TuttiIRecord = False Then
            ParteFiltro = " AND Letto = 0"
        Else
            ParteFiltro = ""
        End If
        Dim utenteEscaped = SessioneUtente.NomeUtenteCorrente.Replace("'", "''")
        Dim filtro As String = $"Destinatario = '{utenteEscaped}'{ParteFiltro}"

        ApriModuloConPermessi("Mov_Notifiche", Me, filtro)
    End Sub

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
    Private Sub ComunicazioniToolStripMenuItem_Click(sender As Object, e As EventArgs)
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
        ApriNotifiche(True)
    End Sub

    Private Sub TipiFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TipiFileToolStripMenuItem.Click
        ApriModuloConPermessi("Tab_TipiOggettoLavorazione", Me)
    End Sub

    Private Sub TipiFileImgVideoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TipiFileImgVideoToolStripMenuItem.Click
        ApriModuloConPermessi("Tab_TipiFile", Me)
    End Sub

    Private Sub GeneraNuovoDatabaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GeneraNuovoDatabaseToolStripMenuItem.Click
        Dim modulo As New SysNuovoDatabase()
        ApriModulo2ConPermessi("Acquisici da PDF", modulo)
    End Sub

End Class