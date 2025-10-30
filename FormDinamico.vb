Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports AxWMPLib
Imports GesPu25.ModuloCampiDinamici
Imports Microsoft.Data.SqlClient
Imports PdfSharp
Imports PdfSharp.Drawing
Imports PdfSharp.Drawing.Layout
Imports PdfSharp.Events
Imports PdfSharp.Pdf
Imports WMPLib
Imports System.Diagnostics

Public Class DynamicDataForm
    Inherits Form

    Private campoInputs As New Dictionary(Of String, Control)
    Private campiDefiniti As List(Of CampoDatabase)
    Private dgvDati As DataGridView
    Private panelBottoni As FlowLayoutPanel
    Private modalita As String = ""
    Private isModifica As Boolean
    Private pannelloSinistro As TableLayoutPanel
    Private nomeTabellaCorrente As String
    Private ModalitaCorrente As String = "nessuna"
    Private lblModalita As Label
    Private lampeggioAttivo As Boolean = False
    Private Shared visualFormsAttivi As New Dictionary(Of String, VisualMediaForm)
    Private panelBottoniDinamici As FlowLayoutPanel
    Private splitContainer As SplitContainer
    Private colonneModificate As Boolean = False
    Private isInAvvioForm As Boolean = True

    ' Ottimizzazioni: cache lookup e regex precompilate
    Private lookupCache As New Dictionary(Of String, DataTable)(StringComparer.OrdinalIgnoreCase)
    Private regexCache As New Dictionary(Of String, Regex)(StringComparer.OrdinalIgnoreCase)

    ' Flag per soppressione eventi UI quando aggiornamento massivo
    Private isUpdatingControls As Boolean = False

    Public Property FiltroIniziale As String

    Public Sub New(campi As List(Of CampoDatabase), nomeTabella As String)
        Me.Name = nomeTabella
        Me.Text = "Form Dinamico"
        Me.Size = New Size(1100, 600)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.campiDefiniti = campi
        Dim convalide = RecuperaConvalideDaSys(nomeTabella)

        For Each campo In campi
            If convalide.ContainsKey(campo.Nome) Then
                Dim r = convalide(campo.Nome)
                ApplicaConvalidaAlCampo(campo, r)
            End If
        Next

        Me.nomeTabellaCorrente = nomeTabella

        AddHandler Me.Load, AddressOf DynamicDataForm_Load

        GestioneStatoForm.CaricaStato(Me)

        splitContainer = New SplitContainer With {
            .Dock = DockStyle.Fill,
            .Orientation = Orientation.Vertical,
            .FixedPanel = FixedPanel.None
        }
        Me.Controls.Add(splitContainer)

        ' Pannello sinistro con campi e bottoni
        Dim layoutSinistroInterno As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .RowCount = 2,
            .ColumnCount = 1
        }
        layoutSinistroInterno.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        layoutSinistroInterno.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        pannelloSinistro = New TableLayoutPanel With {
            .Dock = DockStyle.Top,
            .AutoSize = True,
            .AutoScroll = True,
            .ColumnCount = 2,
            .Padding = New Padding(20),
            .GrowStyle = TableLayoutPanelGrowStyle.AddRows
        }
        pannelloSinistro.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        pannelloSinistro.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))

        lblModalita = New Label With {
            .Text = "",
            .AutoSize = True,
            .Font = New Font("Verdana", 8, FontStyle.Bold),
            .ForeColor = Color.DarkGreen,
            .Dock = DockStyle.Top,
            .Padding = New Padding(5),
            .TextAlign = ContentAlignment.TopLeft
        }
        pannelloSinistro.Controls.Add(lblModalita)
        pannelloSinistro.SetColumnSpan(lblModalita, 2)

        For i = 0 To campi.Count - 1
            If pannelloSinistro.RowCount <= i + 1 Then
                pannelloSinistro.RowCount += 1
                pannelloSinistro.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            End If

            Dim lbl As New Label With {
                .Text = GetEtichetta(nomeTabella, campi(i).Nome),
                .AutoSize = True,
                .Anchor = AnchorStyles.Left,
                .Margin = New Padding(5)
            }
            Dim ctrl As Control = CreaControllo(campi(i))
            ctrl.Anchor = AnchorStyles.Left
            ctrl.Margin = New Padding(5)
            campoInputs.Add(campi(i).Nome, ctrl)

            pannelloSinistro.Controls.Add(lbl, 0, i + 1)
            pannelloSinistro.Controls.Add(ctrl, 1, i + 1)
        Next

        Dim panelBottoniContenitore As New FlowLayoutPanel With {
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.LeftToRight,
            .AutoSize = True,
            .Padding = New Padding(10),
            .Margin = New Padding(0),
            .WrapContents = True
        }

        panelBottoni = New FlowLayoutPanel With {
            .FlowDirection = FlowDirection.LeftToRight,
            .AutoSize = True,
            .Margin = New Padding(0)
        }
        panelBottoniContenitore.Controls.Add(panelBottoni)

        AggiungiBottone("Inserisci", AddressOf InserisciDati)
        AggiungiBottone("Modifica", AddressOf ModificaDati)
        AggiungiBottone("Salva", AddressOf SalvaDati)
        DisabilitaPulsante("Salva", True)
        AggiungiBottone("Cancella", AddressOf CancellaDati)
        AggiungiBottone("Annulla", AddressOf AnnullaOperazione)
        DisabilitaPulsante("Annulla", True)
        AggiungiBottone("Esporta PDF", AddressOf EsportaPDF)

        AggiungiBottone("Rimuovi filtro", Sub()
                                              Dim dt As DataTable = TryCast(dgvDati.DataSource, DataTable)
                                              If dt IsNot Nothing Then dt.DefaultView.RowFilter = ""
                                              lblModalita.Text = "In Attesa..."
                                              lblModalita.ForeColor = Color.DarkGreen
                                          End Sub)

        panelBottoniDinamici = New FlowLayoutPanel With {
            .FlowDirection = FlowDirection.LeftToRight,
            .AutoSize = True,
            .Margin = New Padding(0, 5, 0, 0)
        }
        panelBottoniContenitore.Controls.Add(panelBottoniDinamici)

        layoutSinistroInterno.Controls.Add(pannelloSinistro, 0, 0)
        layoutSinistroInterno.Controls.Add(panelBottoniContenitore, 0, 1)
        splitContainer.Panel1.Controls.Add(layoutSinistroInterno)

        ' Griglia dati
        dgvDati = New DataGridView With {
            .Dock = DockStyle.Fill,
            .AllowUserToAddRows = False,
            .ReadOnly = True,
            .Name = nomeTabellaCorrente,
            .ScrollBars = ScrollBars.Both,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
            .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        }
        With dgvDati
            .ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
            .ColumnHeadersHeight = 50
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToOrderColumns = False
            .AllowUserToResizeColumns = True
        End With

        AddHandler dgvDati.CellClick, AddressOf dgvDati_CellClick
        AddHandler dgvDati.SelectionChanged, AddressOf dgvDati_SelectionChanged
        AddHandler dgvDati.DataBindingComplete, AddressOf dgvDati_DataBindingComplete
        AddHandler dgvDati.CellDoubleClick, AddressOf dgvDati_CellDoubleClick

        InizializzaEventiGriglia()

        splitContainer.Panel2.Controls.Add(dgvDati)

        ' Caricamenti iniziali
        CaricaBottoniDinamici()

        For Each ctrl As Control In campoInputs.Values
            If TypeOf ctrl Is FlowLayoutPanel Then
                Dim hasVisualBtn = ctrl.Controls.OfType(Of Button)().Any(Function(b) String.Equals(b.Text?.Trim(), "Visualizza", StringComparison.OrdinalIgnoreCase))
                ctrl.Enabled = True

                For Each innerCtrl As Control In ctrl.Controls
                    If TypeOf innerCtrl Is Button AndAlso String.Equals(CType(innerCtrl, Button).Text?.Trim(), "Visualizza", StringComparison.OrdinalIgnoreCase) Then
                        innerCtrl.Enabled = True
                    Else
                        innerCtrl.Enabled = False
                    End If
                Next
            Else
                ctrl.Enabled = False
            End If
        Next

        UniformaDimensioniBottoni()
        ApplicaAutorizzazioni(NomeUtenteCorrente)
        PulisciCampi()
        DisabilitaPulsante("Salva", True)
        DisabilitaPulsante("Annulla", True)

        ' Precompila regex cache per i nomi campo presenti
        For Each kvp In campoInputs
            If Not regexCache.ContainsKey(kvp.Key) Then
                regexCache(kvp.Key) = New Regex($"\b{Regex.Escape(kvp.Key)}\b", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
            End If
        Next

    End Sub

    Private Sub InizializzaEventiGriglia()
        If dgvDati Is Nothing Then Return

        AddHandler dgvDati.ColumnWidthChanged, Sub(s, e)
                                                   If isInAvvioForm Then Return
                                                   colonneModificate = True
                                               End Sub

        AddHandler dgvDati.ColumnDisplayIndexChanged, Sub(s, e)
                                                          If isInAvvioForm Then Return
                                                          colonneModificate = True
                                                      End Sub
    End Sub

    Private Sub dgvDati_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Exit Sub

        Dim colName = dgvDati.Columns(e.ColumnIndex).Name
        Dim cellValue = dgvDati.Rows(e.RowIndex).Cells(e.ColumnIndex).Value?.ToString()
        If String.IsNullOrWhiteSpace(cellValue) Then Exit Sub

        Dim dt As DataTable = TryCast(dgvDati.DataSource, DataTable)
        If dt Is Nothing Then Exit Sub

        Try
            Dim nuovaCondizione = $"[{colName}] = '{cellValue.Replace("'", "''")}'"
            Dim filtroCorrente = dt.DefaultView.RowFilter

            If String.IsNullOrWhiteSpace(filtroCorrente) Then
                dt.DefaultView.RowFilter = nuovaCondizione
            Else
                dt.DefaultView.RowFilter = $"{filtroCorrente} AND {nuovaCondizione}"
            End If

            lblModalita.Text = $"Filtro attivo: {dt.DefaultView.RowFilter}"
            lblModalita.ForeColor = Color.DarkBlue
        Catch ex As Exception
            MessageBox.Show("Errore nel filtro: " & ex.Message)
        End Try
    End Sub

    Private Sub ApplicaConvalidaAlCampo(campo As CampoDatabase, r As DataRow)
        campo.TipoConvalida = r("TipoConvalida").ToString()
        campo.IntervalloMin = r("IntervalloMin").ToString()
        campo.IntervalloMax = r("IntervalloMax").ToString()
        campo.TabellaElenco = r("TabellaElenco").ToString()
        campo.ChiaveElenco = r("ChiaveElenco").ToString()
        campo.DescrizioneChiave = r("DescrizioneChiave").ToString()
        campo.CampoVisuale = r("CampoVisuale").ToString()
        campo.AbilitaZoom = If(r.Table.Columns.Contains("AbilitaZoom"), Convert.ToBoolean(r("AbilitaZoom")), False)
        campo.AbilitaModifica = If(r.Table.Columns.Contains("AbilitaModifica"), Convert.ToBoolean(r("AbilitaModifica")), True)
        If String.IsNullOrWhiteSpace(campo.TabellaElenco) Then
            campo.AbilitaModifica = True
        End If
    End Sub

    Private Function RecuperaConvalideDaSys(nomeTabella As String) As Dictionary(Of String, DataRow)
        Dim convalide As New Dictionary(Of String, DataRow)
        Using conn As New SqlConnection(ConnString)
            Dim query = "SELECT * FROM Sys_ConvalidaCampi WHERE NomeTabella = @Tabella"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@Tabella", nomeTabella)
                Dim da As New SqlDataAdapter(cmd)
                Dim dt As New DataTable()
                da.Fill(dt)
                For Each row As DataRow In dt.Rows
                    Dim nomeCampo = row("NomeCampo").ToString()
                    If Not convalide.ContainsKey(nomeCampo) Then
                        convalide.Add(nomeCampo, row)
                    End If
                Next
            End Using
        End Using
        Return convalide
    End Function

    Private Sub dgvDati_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs)
        If dgvDati.Columns.Count > 0 Then
            ApplicaVisualizzazioneColonne()
            Me.BeginInvoke(New MethodInvoker(Sub()
                                                 ApplicaConfigurazioneGriglia(dgvDati)
                                                 NascondiColonneSensibili()
                                             End Sub))
        Else
            Debug.WriteLine("Le colonne della griglia non sono ancora disponibili.")
        End If
    End Sub

    Private Sub NascondiColonneSensibili()
        For Each col As DataGridViewColumn In dgvDati.Columns
            If col.Name.ToLower().Contains("password") Then
                col.Visible = False
            End If
        Next
    End Sub

    Private Sub DynamicDataForm_Load(sender As Object, e As EventArgs)

        If splitContainer IsNot Nothing Then
            splitContainer.Panel1MinSize = 300
            splitContainer.Panel2MinSize = 300
            splitContainer.SplitterDistance = Me.Width / 2.5
        End If

        ' Carica i dati con eventuale filtro
        CaricaDatiTabellaAsync(Me.Name)

        Me.BeginInvoke(New MethodInvoker(Sub()
                                             isInAvvioForm = False
                                             Me.Refresh()
                                             UpdateButtonsByModalita()
                                         End Sub))

    End Sub

    Private Sub CaricaDatiTabellaAsync(nomeTabella As String)
        Dim filtro = FiltroIniziale
        Task.Run(Sub()
                     Dim dt As New DataTable()
                     Try
                         Using conn As New SqlConnection(ConnString)
                             Using cmd As New SqlCommand($"SELECT * FROM [{nomeTabella}]" & If(String.IsNullOrWhiteSpace(filtro), "", $" WHERE {filtro}"), conn)
                                 conn.Open()
                                 Using adapter As New SqlDataAdapter(cmd)
                                     adapter.Fill(dt)
                                 End Using
                             End Using
                         End Using
                     Catch ex As Exception
                         Dim msg = ex.Message
                         Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show("Errore caricamento dati: " & msg, Me.MdiParent, MessageBoxButtons.OK)))
                         Return
                     End Try

                     Me.BeginInvoke(New MethodInvoker(Sub()
                                                          dgvDati.DataSource = dt
                                                      End Sub))
                 End Sub)
    End Sub

    Private Sub DynamicDataForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        GestioneStatoForm.SalvaStato(Me)

        If colonneModificate Then
            Try
                SalvaConfigurazioneGrigliaBatched(dgvDati)
                colonneModificate = False
            Catch ex As Exception
                System.Diagnostics.Trace.TraceError($"Errore salvando configurazione griglia alla chiusura: {ex.Message}")
            End Try
        End If
    End Sub

    Private Sub FormDinamico_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        PosizionaGrigliaDaSysForm()
    End Sub

    Private Sub AnnullaOperazione()
        Dim risposta = MDIMessageBox.Show("Vuoi annullare l’operazione corrente?", Me.MdiParent, MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If risposta = DialogResult.Yes Then

            For Each ctrl As Control In campoInputs.Values
                If TypeOf ctrl Is FlowLayoutPanel Then
                    For Each innerCtrl As Control In ctrl.Controls
                        If TypeOf innerCtrl Is Button AndAlso CType(innerCtrl, Button).Text = "Visualizza" Then
                            Continue For
                        End If
                        innerCtrl.Enabled = False
                    Next
                Else
                    ctrl.Enabled = False
                End If
            Next

            DisabilitaCampi()
            CaricaDatiTabella(Me.Name)
            DisabilitaPulsante("Salva", True)
            lampeggioAttivo = False
            lblModalita.ForeColor = Color.DarkGreen
            DisabilitaPulsante("Annulla", True)
            ModalitaCorrente = "nessuna"
            lblModalita.Text = "In Attesa..."
            PulisciCampi()
            ResetLabelDescrizioni()
            UpdateButtonsByModalita()

        End If
    End Sub

    Private Sub AbilitaCampi(abilita As Boolean)
        For Each kvp In campoInputs
            Dim nomeCampo As String = kvp.Key
            Dim ctrl As Control = kvp.Value

            Dim campo As CampoDatabase = campiDefiniti.FirstOrDefault(Function(c) c.Nome = nomeCampo)
            If campo Is Nothing Then Continue For

            If nomeCampo.StartsWith("Calc_") Then
                ctrl.Enabled = False
                Continue For
            End If

            Dim joinRow = RecuperaJoinPerCampo(nomeTabellaCorrente, nomeCampo)
            Dim isJoin = joinRow IsNot Nothing

            Dim joinModificabile As Boolean = True
            If isJoin AndAlso joinRow.Table.Columns.Contains("AbilitaModifica") Then
                joinModificabile = Convert.ToBoolean(joinRow("AbilitaModifica"))
            End If

            Dim isBloccato As Boolean = campo.IsIdentity OrElse campo.IsChiave OrElse (isJoin AndAlso Not joinModificabile)

            If TypeOf ctrl Is FlowLayoutPanel Then
                Dim flow = CType(ctrl, FlowLayoutPanel)
                Dim hasVisualBtn = flow.Controls.OfType(Of Button)().Any(Function(b) String.Equals(b.Text?.Trim(), "Visualizza", StringComparison.OrdinalIgnoreCase))
                If hasVisualBtn Then flow.Enabled = True

                For Each innerCtrl As Control In flow.Controls
                    If TypeOf innerCtrl Is Button AndAlso String.Equals(CType(innerCtrl, Button).Text?.Trim(), "Visualizza", StringComparison.OrdinalIgnoreCase) Then
                        innerCtrl.Enabled = True
                    Else
                        innerCtrl.Enabled = Not isBloccato AndAlso abilita
                    End If
                Next
            Else
                ctrl.Enabled = Not isBloccato AndAlso abilita
            End If

        Next
    End Sub

    Private Sub DisabilitaCampi()
        AbilitaCampi(False)
    End Sub

    Private Function CampoIDGestitoManuale() As Boolean
        Return False
    End Function

    Private Sub AggiungiBottone(nome As String, handler As EventHandler)
        Dim btn As New Button With {.Text = nome, .AutoSize = True}
        AddHandler btn.Click, handler
        panelBottoni.Controls.Add(btn)
    End Sub

    Private Sub UniformaDimensioniBottoni()
        Dim larghezzaMassima As Integer = 0
        Dim altezzaMassima As Integer = 0

        For Each ctrl As Control In panelBottoni.Controls
            If TypeOf ctrl Is Button Then
                Dim btn As Button = CType(ctrl, Button)
                If btn.Width > larghezzaMassima Then larghezzaMassima = btn.Width
                If btn.Height > altezzaMassima Then altezzaMassima = btn.Height
            End If
        Next

        For Each ctrl As Control In panelBottoni.Controls
            If TypeOf ctrl Is Button Then
                Dim btn As Button = CType(ctrl, Button)
                btn.AutoSize = False
                btn.Width = larghezzaMassima
                btn.Height = altezzaMassima
            End If
        Next
    End Sub

    Private Sub FocusSulPrimoCampoEditabile()
        Me.BeginInvoke(New MethodInvoker(Sub()
                                             Try
                                                 Dim primo As Control = Nothing

                                                 ' Ordina secondo l'ordine dei controlli nel pannello sinistro
                                                 For i = 0 To pannelloSinistro.RowCount - 1
                                                     For Each c As Control In pannelloSinistro.GetControlFromPosition(1, i)?.Controls
                                                         ' ignora se nulla
                                                     Next
                                                 Next

                                                 ' Scorri i controlli nel TableLayoutPanel rispettando l'ordine di aggiunta
                                                 For Each ctrl As Control In pannelloSinistro.Controls
                                                     If ctrl Is lblModalita Then Continue For
                                                     ' Il controllo reale è nella seconda colonna (index 1) quando è stato aggiunto
                                                     If Not ctrl.Enabled Then Continue For

                                                     ' Se è FlowLayoutPanel, cerca il primo TextBox interno
                                                     If TypeOf ctrl Is FlowLayoutPanel Then
                                                         Dim innerTxt = ctrl.Controls.OfType(Of TextBox)().FirstOrDefault(Function(t) t.Enabled AndAlso t.Visible)
                                                         If innerTxt IsNot Nothing Then
                                                             primo = innerTxt
                                                             Exit For
                                                         End If

                                                         Dim innerCombo = ctrl.Controls.OfType(Of ComboBox)().FirstOrDefault(Function(cb) cb.Enabled AndAlso cb.Visible)
                                                         If innerCombo IsNot Nothing Then
                                                             primo = innerCombo
                                                             Exit For
                                                         End If
                                                     Else
                                                         ' Controlli supportati direttamente
                                                         If (TypeOf ctrl Is TextBox OrElse TypeOf ctrl Is ComboBox OrElse TypeOf ctrl Is DateTimePicker OrElse TypeOf ctrl Is CheckBox) AndAlso ctrl.Enabled AndAlso ctrl.Visible Then
                                                             primo = ctrl
                                                             Exit For
                                                         End If
                                                     End If
                                                 Next

                                                 If primo IsNot Nothing Then
                                                     primo.Focus()
                                                     ' Se è TextBox, seleziona tutto il testo per comodità
                                                     If TypeOf primo Is TextBox Then
                                                         CType(primo, TextBox).SelectAll()
                                                     ElseIf TypeOf primo Is ComboBox Then
                                                         CType(primo, ComboBox).DroppedDown = False
                                                     End If
                                                 End If
                                             Catch ex As Exception
                                                 Trace.TraceWarning($"FocusSulPrimoCampoEditabile errore: {ex.Message}")
                                             End Try
                                         End Sub))
    End Sub


    Private Sub InserisciDati(sender As Object, e As EventArgs)

        isModifica = False
        PulisciCampi()
        AbilitaCampi(True)
        ResetLabelDescrizioni()
        'DisabilitaPulsante("Salva", False)

        ModalitaCorrente = "inserimento"
        lblModalita.Text = "Inserimento in corso..."
        'DisabilitaPulsante("Annulla", False)
        UpdateButtonsByModalita()

        For Each campo In campiDefiniti
            If Not campoInputs.ContainsKey(campo.Nome) Then Continue For
            Dim ctrl = campoInputs(campo.Nome)
            If ctrl Is Nothing Then Continue For

            Dim valoreCorrente As String = ""
            Select Case True
                Case TypeOf ctrl Is ComboBox
                    valoreCorrente = CType(ctrl, ComboBox).SelectedValue?.ToString()

                Case TypeOf ctrl Is TextBox
                    valoreCorrente = CType(ctrl, TextBox).Text

                Case TypeOf ctrl Is FlowLayoutPanel
                    Dim txt = ctrl.Controls.OfType(Of TextBox).FirstOrDefault()
                    If txt IsNot Nothing Then valoreCorrente = txt.Text
            End Select

            Dim campoPrecompilato = Not String.IsNullOrWhiteSpace(valoreCorrente)

            If Not campoPrecompilato Then
                Select Case True
                    Case TypeOf ctrl Is TextBox
                        CType(ctrl, TextBox).Clear()

                    Case TypeOf ctrl Is CheckBox
                        CType(ctrl, CheckBox).Checked = False

                    Case TypeOf ctrl Is ComboBox
                        CType(ctrl, ComboBox).SelectedIndex = -1

                    Case TypeOf ctrl Is DateTimePicker
                        CType(ctrl, DateTimePicker).Value = DateTime.Now

                    Case TypeOf ctrl Is FlowLayoutPanel
                        For Each innerCtrl As Control In ctrl.Controls
                            If TypeOf innerCtrl Is TextBox Then
                                CType(innerCtrl, TextBox).Clear()
                            End If
                        Next
                End Select
            End If

            ctrl.Enabled = Not campo.IsIdentity
        Next

        FocusSulPrimoCampoEditabile()

    End Sub

    Private Function RecuperaJoinPerCampo(nomeTabella As String, nomeCampo As String) As DataRow
        Using conn As New SqlConnection(ConnString)
            Dim query = "SELECT * FROM sys_CampiJoin WHERE NomeTabella = @Tabella AND NomeCampo = @Campo"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@Tabella", nomeTabella)
                cmd.Parameters.AddWithValue("@Campo", nomeCampo)
                Dim da As New SqlDataAdapter(cmd)
                Dim dt As New DataTable()
                da.Fill(dt)
                If dt.Rows.Count > 0 Then Return dt.Rows(0)
            End Using
        End Using
        Return Nothing
    End Function

    Private Function PrelevaValoreJoin(joinRow As DataRow, chiaviFiglia As Dictionary(Of String, Object)) As Object
        Dim tabellaPadre = joinRow("TabellaPadre").ToString()
        Dim campoDaPrelevare = joinRow("CampoDaPrelevare").ToString()

        Dim condizioni As New List(Of String)
        Dim parametri As New Dictionary(Of String, Object)

        For i = 1 To 3
            Dim nomeColonna = $"ChiavePadre{i}"
            If joinRow.Table.Columns.Contains(nomeColonna) Then
                Dim chiavePadre = joinRow(nomeColonna).ToString()
                If Not String.IsNullOrWhiteSpace(chiavePadre) AndAlso chiaviFiglia.ContainsKey($"ChiaveFiglia{i}") Then
                    condizioni.Add($"{chiavePadre} = @param{i}")
                    parametri.Add($"@param{i}", chiaviFiglia($"ChiaveFiglia{i}"))
                End If
            End If
        Next

        If condizioni.Count = 0 Then Return Nothing

        Dim query = $"SELECT {campoDaPrelevare} FROM {tabellaPadre} WHERE {String.Join(" AND ", condizioni)}"

        Using conn As New SqlConnection(ConnString)
            Using cmd As New SqlCommand(query, conn)
                For Each kvp In parametri
                    cmd.Parameters.AddWithValue(kvp.Key, kvp.Value)
                Next
                conn.Open()
                Return cmd.ExecuteScalar()
            End Using
        End Using
    End Function

    Private Sub ModificaDati(sender As Object, e As EventArgs)

        If dgvDati.SelectedRows.Count = 0 Then
            Dim risposta = MDIMessageBox.Show("Seleziona prima una riga dalla griglia.", Me.MdiParent, MessageBoxButtons.OK)
            Return
        End If

        DisabilitaPulsante("Salva", False)
        ModalitaCorrente = "modifica"
        lblModalita.Text = "Modifica in corso..."
        'DisabilitaPulsante("Annulla", False)
        UpdateButtonsByModalita()

        isModifica = True
        AbilitaCampi(True)
        ModalitaCorrente = "modifica"
        lblModalita.Text = "Modifica in corso..."
        lblModalita.ForeColor = Color.Green
        lblModalita.Font = New Font("Segoe UI", 8, FontStyle.Bold)

        Dim rigaSelezionata = dgvDati.SelectedRows(0)

        For Each campo In campiDefiniti
            If Not campoInputs.ContainsKey(campo.Nome) Then Continue For

            Dim joinRow = RecuperaJoinPerCampo(Me.Name, campo.Nome)
            If joinRow IsNot Nothing Then
                Dim chiaveFiglia = joinRow("ChiaveFiglia1").ToString()
                Dim valoreChiave = rigaSelezionata.Cells(chiaveFiglia).Value
                If valoreChiave IsNot Nothing Then
                    Dim chiaviFiglia As New Dictionary(Of String, Object)

                    For i = 1 To 3
                        Dim chiaveNome = $"ChiaveFiglia{i}"
                        If joinRow.Table.Columns.Contains(chiaveNome) Then
                            Dim nomeCampoFiglio = joinRow(chiaveNome).ToString()
                            If Not String.IsNullOrWhiteSpace(nomeCampoFiglio) AndAlso dgvDati.Columns.Contains(nomeCampoFiglio) Then
                                Dim valore = dgvDati.SelectedRows(0).Cells(nomeCampoFiglio).Value
                                chiaviFiglia.Add(chiaveNome, valore)
                            End If
                        End If
                    Next

                    Dim valoreJoin = PrelevaValoreJoin(joinRow, chiaviFiglia)

                    Dim ctrl = campoInputs(campo.Nome)
                    If TypeOf ctrl Is TextBox Then
                        CType(ctrl, TextBox).Text = valoreJoin?.ToString()
                    ElseIf TypeOf ctrl Is ComboBox Then
                        CType(ctrl, ComboBox).SelectedValue = valoreJoin
                    End If
                End If
            End If
        Next

        FocusSulPrimoCampoEditabile()

    End Sub

    Private Async Sub SalvaDati(sender As Object, e As EventArgs)
        Dim sw As New Stopwatch()
        sw.Start()
        Try
            ToggleUIForSaving(True)
            If isModifica Then
                Await SalvaModificaAsync()
            Else
                Await SalvaInserimentoAsync()
            End If
            Me.BeginInvoke(New MethodInvoker(Sub()
                                                 DisabilitaCampi()
                                                 CaricaDatiTabella(Me.Name)
                                                 DisabilitaPulsante("Salva", True)
                                                 lblModalita.ForeColor = Color.DarkGreen
                                                 ModalitaCorrente = "nessuna"
                                                 lblModalita.Text = ""
                                                 DisabilitaPulsante("Annulla", True)
                                                 UpdateButtonsByModalita()
                                             End Sub))
        Catch ex As Exception
            Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show("Errore durante il salvataggio: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)))
        Finally
            ToggleUIForSaving(False)
            sw.Stop()
            Trace.TraceInformation($"SalvaDati durata totale: {sw.ElapsedMilliseconds} ms. ModalitaModifica={isModifica}")
        End Try
    End Sub


    Private Sub CancellaDati(sender As Object, e As EventArgs)
        If dgvDati.SelectedRows.Count = 0 Then
            MDIMessageBox.Show("Seleziona prima una riga da cancellare dalla griglia.", Me.MdiParent, MessageBoxButtons.OK)
            Return
        End If

        campiDefiniti = RecuperaCampiDa(Me.Name)

        Dim campoChiave = campiDefiniti.FirstOrDefault(Function(c) c.IsChiave)
        If campoChiave Is Nothing Then
            MDIMessageBox.Show("Nessuna chiave primaria definita.", Me.MdiParent, MessageBoxButtons.OK)
            Return
        End If

        Dim cella = dgvDati.SelectedRows(0).Cells(campoChiave.Nome)
        If cella.Value Is Nothing Then
            MDIMessageBox.Show("Il valore della chiave è nullo.", Me.MdiParent, MessageBoxButtons.OK)
            Return
        End If

        Dim valoreChiave = cella.Value.ToString()
        Dim conferma = MDIMessageBox.Show($"Sei sicuro di voler cancellare il record con chiave {campoChiave.Nome} = {valoreChiave}?", Me.MdiParent, MessageBoxButtons.YesNo)

        If conferma = DialogResult.Yes Then
            Dim query As String = $"DELETE FROM [{Me.Name}] WHERE {campoChiave.Nome} = @{campoChiave.Nome}"

            Try
                Using conn As New SqlConnection(ConnString)
                    Using cmd As New SqlCommand(query, conn)
                        cmd.Parameters.AddWithValue("@" & campoChiave.Nome, valoreChiave)
                        conn.Open()
                        cmd.ExecuteNonQuery()
                    End Using
                End Using

                CaricaDatiTabella(Me.Name)
                PulisciCampi()

            Catch ex As SqlException
                If ex.Number = 547 Then
                    MDIMessageBox.Show("Impossibile cancellare il record: è referenziato da altre tabelle.", Me.MdiParent, MessageBoxButtons.OK)
                Else
                    MDIMessageBox.Show("Errore SQL durante la cancellazione: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
                End If
            Catch ex As Exception
                MDIMessageBox.Show("Errore imprevisto: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
            End Try
        End If
    End Sub

    Private Sub ResetLabelDescrizioni()
        For Each campo In campiDefiniti
            If Not campoInputs.ContainsKey(campo.Nome) Then Continue For
            Dim ctrl = campoInputs(campo.Nome)

            If TypeOf ctrl Is FlowLayoutPanel Then
                Dim lbl As Label = ctrl.Controls.OfType(Of Label).FirstOrDefault()
                If lbl IsNot Nothing Then lbl.Text = "..."
            End If
        Next
    End Sub

    Private Function GetSqlDbTypePerCampo(campo As CampoDatabase) As SqlDbType
        If campo Is Nothing OrElse String.IsNullOrWhiteSpace(campo.Tipo) Then
            Return SqlDbType.NVarChar
        End If

        Select Case campo.Tipo.ToLower().Trim()
            Case "int", "integer" : Return SqlDbType.Int
            Case "bit", "boolean" : Return SqlDbType.Bit
            Case "datetime", "date", "smalldatetime" : Return SqlDbType.DateTime
            Case "decimal", "numeric", "money" : Return SqlDbType.Decimal
            Case "float", "double" : Return SqlDbType.Float
            Case "uniqueidentifier" : Return SqlDbType.UniqueIdentifier
            Case "varbinary", "image" : Return SqlDbType.VarBinary
            Case Else
                Return SqlDbType.NVarChar
        End Select
    End Function


    Private Async Function SalvaInserimentoAsync() As Task
        Dim sw As New Stopwatch()
        sw.Start()

        ' Prepara calcoli e valori locali
        Dim campiCalcolati = RecuperaCampiCalcolati()
        Dim formule = campiCalcolati.ToDictionary(Function(kvp) kvp.Key, Function(kvp) kvp.Value.Formula)
        Dim tipiValore = campiCalcolati.ToDictionary(Function(kvp) kvp.Key, Function(kvp) kvp.Value.TipoValore)
        Dim valoriCalcolati = CalcolaValoriCampiCalcolati(formule, tipiValore)

        Dim colonne = campiDefiniti.Where(Function(c) Not c.IsIdentity).Select(Function(c) c.Nome).ToList()
        If colonne.Count = 0 Then Return

        Dim query As String = $"INSERT INTO [{Me.Name}] ({String.Join(",", colonne)}) VALUES ({String.Join(",", colonne.Select(Function(n) "@" & n))})"

        Dim errorList As New List(Of String)

        Using conn As New SqlConnection(ConnString)
            Await conn.OpenAsync()
            Using tx = conn.BeginTransaction()
                Using cmd As New SqlCommand(query, conn, tx)
                    cmd.CommandTimeout = 120

                    ' Prepara parametri tipizzati
                    For Each nomeCampo In colonne
                        Dim campoDef = campiDefiniti.FirstOrDefault(Function(c) String.Equals(c.Nome, nomeCampo, StringComparison.OrdinalIgnoreCase))
                        Dim sqlType = If(campoDef IsNot Nothing, GetSqlDbTypePerCampo(campoDef), SqlDbType.NVarChar)
                        Dim size As Integer = 0
                        If campoDef IsNot Nothing AndAlso sqlType = SqlDbType.NVarChar Then
                            Dim l As Integer = 0
                            If Integer.TryParse(Convert.ToString(campoDef.Lunghezza), l) Then size = Math.Min(Math.Max(l, 0), 4000)
                        End If

                        If size > 0 Then
                            cmd.Parameters.Add("@" & nomeCampo, sqlType, size)
                        Else
                            cmd.Parameters.Add("@" & nomeCampo, sqlType)
                        End If
                    Next

                    ' Imposta valori parametri
                    For Each nomeCampo In colonne
                        Dim valore As Object = Nothing
                        If valoriCalcolati.ContainsKey(nomeCampo) Then
                            valore = valoriCalcolati(nomeCampo)
                        Else
                            Try
                                valore = EstraiValoreDaControllo(nomeCampo, campoInputs(nomeCampo))
                            Catch ex As Exception
                                valore = DBNull.Value
                                errorList.Add($"Errore prelevando '{nomeCampo}': {ex.Message}")
                            End Try
                        End If

                        If valore Is Nothing Then
                            cmd.Parameters("@" & nomeCampo).Value = DBNull.Value
                        Else
                            If TypeOf valore Is String AndAlso String.IsNullOrEmpty(CType(valore, String)) Then
                                cmd.Parameters("@" & nomeCampo).Value = DBNull.Value
                            ElseIf TypeOf valore Is Decimal OrElse TypeOf valore Is Double Then
                                cmd.Parameters("@" & nomeCampo).Value = Convert.ToDecimal(valore, Globalization.CultureInfo.InvariantCulture)
                            Else
                                cmd.Parameters("@" & nomeCampo).Value = valore
                            End If
                        End If
                    Next

                    Try
                        Await cmd.ExecuteNonQueryAsync()
                        tx.Commit()
                    Catch ex As Exception
                        Try
                            tx.Rollback()
                        Catch
                        End Try
                        Throw
                    End Try
                End Using
            End Using
        End Using

        If errorList.Count > 0 Then
            Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show(String.Join(Environment.NewLine, errorList), Me.MdiParent, MessageBoxButtons.OK)))
        End If

        sw.Stop()
        Trace.TraceInformation($"SalvaInserimentoAsync durata: {sw.ElapsedMilliseconds} ms.")
    End Function


    Private Async Function SalvaModificaAsync() As Task
        Dim sw As New Stopwatch()
        sw.Start()

        campiDefiniti = RecuperaCampiDa(Me.Name)
        Dim campoChiave = campiDefiniti.FirstOrDefault(Function(c) c.IsChiave)
        If campoChiave Is Nothing OrElse dgvDati.SelectedRows.Count = 0 Then
            Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show("Chiave primaria mancante o nessuna riga selezionata.", Me.MdiParent, MessageBoxButtons.OK)))
            Return
        End If

        Dim valoreChiaveObj = dgvDati.SelectedRows(0).Cells(campoChiave.Nome).Value
        If valoreChiaveObj Is Nothing OrElse valoreChiaveObj Is DBNull.Value Then
            Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show("Valore della chiave non trovato o nullo.", Me.MdiParent, MessageBoxButtons.OK)))
            Return
        End If
        Dim valoreChiave = valoreChiaveObj

        Dim campiCalcolati = RecuperaCampiCalcolati()
        Dim formule = campiCalcolati.ToDictionary(Function(kvp) kvp.Key, Function(kvp) kvp.Value.Formula)
        Dim tipiValore = campiCalcolati.ToDictionary(Function(kvp) kvp.Key, Function(kvp) kvp.Value.TipoValore)
        Dim valoriCalcolati = CalcolaValoriCampiCalcolati(formule, tipiValore)

        Dim colonneValid As New List(Of String)
        For Each campo In campiDefiniti
            If campo.IsChiave OrElse campo.IsIdentity Then Continue For
            Dim input = campoInputs(campo.Nome)
            Dim isPassword = campo.Nome.ToLower().Contains("password")
            If isPassword AndAlso TypeOf input Is TextBox AndAlso String.IsNullOrWhiteSpace(CType(input, TextBox).Text) Then
                Continue For
            End If
            colonneValid.Add(campo.Nome)
        Next

        If colonneValid.Count = 0 Then Return

        Dim query As String = $"UPDATE [{Name}] SET {String.Join(",", colonneValid.Select(Function(n) $"{n} = @{n}"))} WHERE {campoChiave.Nome} = @{campoChiave.Nome}"

        Dim errorList As New List(Of String)

        Using conn As New SqlConnection(ConnString)
            Await conn.OpenAsync()
            Using tx = conn.BeginTransaction()
                Using cmd As New SqlCommand(query, conn, tx)
                    cmd.CommandTimeout = 120

                    ' Prepara parametri tipizzati per colonne e per chiave
                    For Each nomeCampo In colonneValid
                        Dim campoDef = campiDefiniti.FirstOrDefault(Function(c) String.Equals(c.Nome, nomeCampo, StringComparison.OrdinalIgnoreCase))
                        Dim sqlType = If(campoDef IsNot Nothing, GetSqlDbTypePerCampo(campoDef), SqlDbType.NVarChar)
                        Dim size As Integer = 0
                        If campoDef IsNot Nothing AndAlso sqlType = SqlDbType.NVarChar Then
                            Dim l As Integer = 0
                            If Integer.TryParse(Convert.ToString(campoDef.Lunghezza), l) Then size = Math.Min(Math.Max(l, 0), 4000)
                        End If

                        If size > 0 Then
                            cmd.Parameters.Add("@" & nomeCampo, sqlType, size)
                        Else
                            cmd.Parameters.Add("@" & nomeCampo, sqlType)
                        End If
                    Next

                    ' parametro per la chiave
                    Dim keySqlType = GetSqlDbTypePerCampo(campoChiave)
                    cmd.Parameters.Add("@" & campoChiave.Nome, keySqlType).Value = valoreChiave

                    ' Imposta valori parametri
                    For Each nomeCampo In colonneValid
                        Dim valoreCampo As Object
                        If valoriCalcolati.ContainsKey(nomeCampo) Then
                            valoreCampo = valoriCalcolati(nomeCampo)
                        Else
                            Try
                                valoreCampo = EstraiValoreDaControllo(nomeCampo, campoInputs(nomeCampo))
                            Catch ex As Exception
                                valoreCampo = DBNull.Value
                                errorList.Add($"Errore prelevando '{nomeCampo}': {ex.Message}")
                            End Try
                        End If

                        If valoreCampo Is Nothing Then
                            cmd.Parameters("@" & nomeCampo).Value = DBNull.Value
                        Else
                            If TypeOf valoreCampo Is String AndAlso String.IsNullOrEmpty(CType(valoreCampo, String)) Then
                                cmd.Parameters("@" & nomeCampo).Value = DBNull.Value
                            ElseIf TypeOf valoreCampo Is Decimal OrElse TypeOf valoreCampo Is Double Then
                                cmd.Parameters("@" & nomeCampo).Value = Convert.ToDecimal(valoreCampo, Globalization.CultureInfo.InvariantCulture)
                            Else
                                cmd.Parameters("@" & nomeCampo).Value = valoreCampo
                            End If
                        End If
                    Next

                    Try
                        Await cmd.ExecuteNonQueryAsync()
                        tx.Commit()
                    Catch ex As Exception
                        Try
                            tx.Rollback()
                        Catch
                        End Try
                        Throw
                    End Try
                End Using
            End Using
        End Using

        If errorList.Count > 0 Then
            Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show(String.Join(Environment.NewLine, errorList), Me.MdiParent, MessageBoxButtons.OK)))
        End If

        sw.Stop()
        Trace.TraceInformation($"SalvaModificaAsync durata: {sw.ElapsedMilliseconds} ms.")
    End Function


    Private Function EstraiValoreDaControllo(nomeCampo As String, input As Control) As Object
        Dim campiBit As String() = {"CanView", "CanInsert", "CanUpdate", "CanDelete"}
        Dim isPassword = nomeCampo.ToLower().Contains("password")

        Select Case True
            Case campiBit.Contains(nomeCampo, StringComparer.OrdinalIgnoreCase) AndAlso TypeOf input Is CheckBox
                Return If(CType(input, CheckBox).Checked, 1, 0)

            Case TypeOf input Is CheckBox
                Return If(CType(input, CheckBox).Checked, 1, 0)

            Case TypeOf input Is ComboBox
                Return CType(input, ComboBox).SelectedValue

            Case TypeOf input Is FlowLayoutPanel
                Dim txt As TextBox = input.Controls.OfType(Of TextBox).FirstOrDefault()
                If txt IsNot Nothing Then
                    Dim valore = txt.Text.Trim()

                    Dim campoDef As CampoDatabase = Nothing
                    Try
                        campoDef = campiDefiniti.FirstOrDefault(Function(c) String.Equals(c.Nome, nomeCampo, StringComparison.OrdinalIgnoreCase))
                    Catch
                        campoDef = Nothing
                    End Try

                    If campoDef IsNot Nothing AndAlso campoDef.Tipo IsNot Nothing AndAlso campoDef.Tipo.ToLower().Trim() = "imgvid" Then
                        Return valore
                    End If

                    If valore.Contains("-"c) Then
                        valore = valore.Split("-"c)(0).Trim()
                    End If
                    Return valore
                End If
                Return ""

            Case TypeOf input Is TextBox
                Dim valore = CType(input, TextBox).Text
                If isPassword Then
                    If String.IsNullOrWhiteSpace(valore) Then Return Nothing
                    Return (New CriptaHash).HashPassword(valore)
                End If
                Return valore

            Case TypeOf input Is DateTimePicker
                Dim dtp = CType(input, DateTimePicker)
                Dim tag = If(dtp.Tag, "").ToString()
                If tag.EndsWith("|NULL") OrElse dtp.CustomFormat = " " Then
                    Return DBNull.Value
                End If
                Return dtp.Value

            Case Else
                Return input.Text
        End Select
    End Function

    Private Function CreaControllo(campo As CampoDatabase) As Control

        If campo Is Nothing Then Return CreaLabelErrore("Campo non valido.")

        Dim larghezzaBase As Integer = 100
        Dim larghezzaMassima As Integer = 450
        Dim larghezzaStimata As Integer
        Dim CarCtrl As Byte = 2

        If Len(campo.Nome.ToLower()) < 2 Then CarCtrl = 1

        If campo.TipoConvalida = "E" OrElse campo.TipoConvalida = "I" Then
            larghezzaStimata = 50
        ElseIf campo.Lunghezza > 0 Then
            larghezzaStimata = Math.Min(larghezzaBase + campo.Lunghezza * 7, larghezzaMassima)
        ElseIf Not String.IsNullOrEmpty(campo.TabellaCollegata) Then
            larghezzaStimata = 380
        ElseIf campo.Nome.ToLower().Contains("descrizione") OrElse campo.Tipo.ToLower().Contains("text") Then
            larghezzaStimata = 350
        ElseIf campo.Tipo.ToLower().Contains("date") OrElse campo.Nome.ToLower().Substring(0, CarCtrl) = "id" Then
            larghezzaStimata = 120
        Else
            larghezzaStimata = 350
        End If

        campo.Lunghezza = larghezzaStimata

        Dim ctrl As Control

        If Not String.IsNullOrEmpty(campo.TabellaCollegata) Then
            ctrl = CreaComboDaTabella(campo)
        Else
            Select Case campo.Tipo.ToLower()
                Case "string", "string_max", "nvarchar", "varchar", "text"
                    ctrl = CreaTextBoxConGestioneTesto(campo)

                Case "date", "datetime"
                    ctrl = CreaDatePickerConGestioneVuoto(campo)

                Case "checkbox", "boolean", "bit"
                    ctrl = CreaCheckBox()

                Case "money"
                    ctrl = CreaTextBoxNumerico()
                    With CType(ctrl, TextBox)
                        .TextAlign = HorizontalAlignment.Right
                        .Tag = "money"
                        .Width = Math.Max(larghezzaStimata, 150)
                    End With

                Case "int"
                    ctrl = CreaTextBoxNumerico()

                Case "decimal"
                    ctrl = CreaNumericUpDown()

                Case "imgvid"
                    ctrl = CreaPannelloMultimediale()

                Case Else
                    ctrl = CreaLabelErrore($"Tipo campo '{campo.Tipo}' non gestito.")
            End Select
        End If

        If ctrl IsNot Nothing AndAlso Not TypeOf ctrl Is CheckBox AndAlso Not TypeOf ctrl Is Label Then
            If Not campo.Tipo.ToLower().Equals("money") Then
                ctrl.Width = larghezzaStimata
            End If

            If TypeOf ctrl Is ComboBox Then
                CType(ctrl, ComboBox).DropDownWidth = larghezzaStimata + 50
            End If
        End If

        If campo.IsIdentity OrElse Not campo.AbilitaModifica Then
            ctrl.Enabled = False
        End If

        If campo.TipoConvalida = "E" AndAlso
           Not String.IsNullOrEmpty(campo.TabellaElenco) AndAlso
           Not String.IsNullOrEmpty(campo.ChiaveElenco) AndAlso
           Not String.IsNullOrEmpty(campo.CampoVisuale) Then

            Dim dt = RecuperaTabellaCached(campo.TabellaElenco)

            Dim txt = New TextBox With {
                .Width = 100,
                .Tag = campo.Nome
            }

            Dim lblDescrizione = New Label With {
                .Width = 200,
                .AutoSize = False,
                .TextAlign = ContentAlignment.MiddleLeft,
                .ForeColor = Color.DarkSlateGray,
                .Padding = New Padding(5, 3, 0, 0),
                .Text = "..."
            }

            AddHandler txt.Leave, Sub()
                                      Dim codice = txt.Text.Trim()
                                      If String.IsNullOrEmpty(codice) Then
                                          lblDescrizione.Text = "..."
                                      Else
                                          If dt IsNot Nothing AndAlso dt.Columns.Contains(campo.ChiaveElenco) Then
                                              Dim riga = dt.Select($"{campo.ChiaveElenco} = '{codice.Replace("'", "''")}'").FirstOrDefault()
                                              If riga IsNot Nothing AndAlso dt.Columns.Contains(campo.CampoVisuale) Then
                                                  lblDescrizione.Text = riga(campo.CampoVisuale).ToString()
                                              Else
                                                  lblDescrizione.Text = "..."
                                              End If
                                          Else
                                              lblDescrizione.Text = "..."
                                          End If
                                      End If
                                  End Sub

            AddHandler txt.Enter, Sub()
                                      lblDescrizione.Text = ""
                                  End Sub

            If campo.AbilitaZoom Then
                AddHandler txt.DoubleClick, Sub()
                                                ApriSelezioneElenco(campo, txt)
                                            End Sub
            End If

            AddHandler txt.Validated, Sub(sender, e)
                                          ValidazioneElenco(campo, CType(sender, Control))
                                      End Sub

            Dim pannello = New FlowLayoutPanel With {
                .AutoSize = True,
                .FlowDirection = FlowDirection.LeftToRight
            }
            pannello.Controls.Add(txt)
            pannello.Controls.Add(lblDescrizione)

            Return pannello
        End If

        If campo.TipoConvalida = "I" Then
            AddHandler ctrl.Validated, Sub(sender, e)
                                           ValidazioneIntervallo(campo, CType(sender, Control))
                                       End Sub
        End If

        Return ctrl
    End Function

    ' Recupero tabella con caching per evitare roundtrips ripetuti
    Private Function RecuperaTabellaCached(nomeTabella As String) As DataTable
        If String.IsNullOrWhiteSpace(nomeTabella) Then Return New DataTable()
        If lookupCache.ContainsKey(nomeTabella) Then Return lookupCache(nomeTabella)

        Dim dt As New DataTable()
        Try
            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand($"SELECT * FROM [{nomeTabella}]", conn)
                    conn.Open()
                    Using da As New SqlDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using
            End Using
            lookupCache(nomeTabella) = dt
        Catch ex As Exception
            Trace.TraceError($"Errore recupero tabella {nomeTabella}: {ex.Message}")
        End Try
        Return dt
    End Function

    Private Function RecuperaTabella(nomeTabella As String, Optional throwOnError As Boolean = False) As DataTable
        Dim dt As New DataTable()

        If String.IsNullOrWhiteSpace(nomeTabella) OrElse Not System.Text.RegularExpressions.Regex.IsMatch(nomeTabella, "^[\w\.]+$") Then
            Dim msg As String = $"Nome tabella non valido: {nomeTabella}"
            System.Diagnostics.Trace.TraceError(msg)
            If throwOnError Then
                Throw New ArgumentException(msg, NameOf(nomeTabella))
            End If
            Return dt
        End If

        Dim query As String = $"SELECT * FROM [{nomeTabella}]"

        Try
            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand(query, conn)
                    cmd.CommandTimeout = 30
                    Using da As New SqlDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using
            End Using
        Catch ex As SqlException
            System.Diagnostics.Trace.TraceError($"Errore SQL recuperando la tabella {nomeTabella}: {ex.Message}")
            If throwOnError Then
                Throw
            End If
        Catch ex As InvalidOperationException
            System.Diagnostics.Trace.TraceError($"Errore di connessione o comando non valido per {nomeTabella}: {ex.Message}")
            If throwOnError Then
                Throw
            End If
        Catch ex As Exception
            System.Diagnostics.Trace.TraceError($"Errore generico recuperando {nomeTabella}: {ex.Message}")
            If throwOnError Then
                Throw
            End If
        End Try

        Return dt
    End Function

    Private Sub ApplicaFeedbackVisivo(ctrl As Control, isValido As Boolean)
        If isValido Then
            ctrl.BackColor = SystemColors.Window
        Else
            ctrl.BackColor = Color.LightPink
        End If
    End Sub

    Private Sub ValidazioneIntervallo(campo As CampoDatabase, inputControl As Control)
        Dim valoreObj As Object = EstraiValoreDaControllo(campo.Nome, inputControl)
        If valoreObj Is Nothing OrElse Not IsNumeric(valoreObj) Then
            inputControl.BackColor = SystemColors.Window
            Exit Sub
        End If

        Dim valore As Decimal = Convert.ToDecimal(valoreObj)
        Dim minOk As Boolean = True
        Dim maxOk As Boolean = True

        If Not String.IsNullOrWhiteSpace(campo.IntervalloMin) AndAlso IsNumeric(campo.IntervalloMin) Then
            minOk = valore >= Convert.ToDecimal(campo.IntervalloMin)
        End If

        If Not String.IsNullOrWhiteSpace(campo.IntervalloMax) AndAlso IsNumeric(campo.IntervalloMax) Then
            maxOk = valore <= Convert.ToDecimal(campo.IntervalloMax)
        End If

        Dim valido = minOk AndAlso maxOk
        ApplicaFeedbackVisivo(inputControl, valido)
        If Not valido Then
            Me.BeginInvoke(Sub() MDIMessageBox.Show($"Il valore '{valore}' per il campo '{campo.Nome}' è fuori dall'intervallo consentito ({campo.IntervalloMin} - {campo.IntervalloMax}).", Me.MdiParent, MessageBoxButtons.OK))
        End If
    End Sub

    Private Sub ValidazioneElenco(campo As CampoDatabase, inputControl As Control)
        Dim valore As Object = EstraiValoreDaControllo(campo.Nome, inputControl)
        If valore Is Nothing OrElse String.IsNullOrWhiteSpace(valore.ToString()) Then
            inputControl.BackColor = SystemColors.Window
            Exit Sub
        End If

        Dim query = $"SELECT COUNT(*) FROM [{campo.TabellaElenco}] WHERE [{campo.ChiaveElenco}] = @valore"
        Try
            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@valore", valore)
                    conn.Open()
                    Dim count = Convert.ToInt32(cmd.ExecuteScalar())
                    Dim valido = (count > 0)
                    ApplicaFeedbackVisivo(inputControl, valido)
                    If Not valido Then
                        Me.BeginInvoke(Sub() MDIMessageBox.Show($"Il valore '{valore}' non è valido per il campo '{campo.Nome}'.", Me.MdiParent, MessageBoxButtons.OK))
                    End If
                End Using
            End Using
        Catch ex As Exception
            inputControl.BackColor = Color.LightPink
            Me.BeginInvoke(Sub() MDIMessageBox.Show("Errore durante la convalida: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK))
        End Try
    End Sub

    Private Sub ApriSelezioneElenco(campo As CampoDatabase, inputControl As Control)
        Dim form = New Form With {
            .Text = $"Seleziona valore per {campo.Nome}",
            .Size = New Size(800, 500),
            .StartPosition = FormStartPosition.CenterParent
        }

        Dim dtOriginale As New DataTable()

        Dim griglia = New DataGridView With {
            .Dock = DockStyle.Fill,
            .ReadOnly = True,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False
        }

        Dim btnMostraTutti As New Button With {
            .Text = "Mostra tutti",
            .Dock = DockStyle.Fill,
            .Height = 30
        }
        AddHandler btnMostraTutti.Click, Sub(s, e)
                                             griglia.DataSource = dtOriginale
                                         End Sub

        Dim layout = New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .RowCount = 2,
            .ColumnCount = 1
        }
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        layout.Controls.Add(btnMostraTutti, 0, 0)
        layout.Controls.Add(griglia, 0, 1)
        form.Controls.Add(layout)

        AddHandler griglia.CellDoubleClick, Sub(s, e)
                                                If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
                                                    Dim colonnaCliccata = griglia.Columns(e.ColumnIndex).Name
                                                    Dim valoreCellaObj = griglia.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                                                    Dim valoreTesto As String = If(valoreCellaObj Is Nothing OrElse TypeOf valoreCellaObj Is DBNull, "", valoreCellaObj.ToString().Trim())

                                                    If colonnaCliccata.Equals(campo.ChiaveElenco, StringComparison.OrdinalIgnoreCase) Then
                                                        If TypeOf inputControl Is TextBox Then
                                                            CType(inputControl, TextBox).Text = valoreTesto
                                                        ElseIf TypeOf inputControl Is ComboBox Then
                                                            CType(inputControl, ComboBox).SelectedValue = valoreTesto
                                                        End If
                                                        form.Close()
                                                    Else
                                                        valoreTesto = valoreTesto.Replace("'", "''")
                                                        Dim colonnaTipo = dtOriginale.Columns(colonnaCliccata).DataType
                                                        Dim filtro As String

                                                        If colonnaTipo = GetType(Integer) OrElse colonnaTipo = GetType(Decimal) OrElse colonnaTipo = GetType(Double) Then
                                                            If IsNumeric(valoreTesto) Then
                                                                filtro = $"[{colonnaCliccata}] = {valoreTesto}"
                                                            Else
                                                                filtro = "1=0"
                                                                Me.BeginInvoke(Sub() MDIMessageBox.Show("Valore non numerico su colonna numerica → nessuna corrispondenza", Me.MdiParent, MessageBoxButtons.OK))
                                                            End If
                                                        Else
                                                            filtro = $"CONVERT([{colonnaCliccata}], System.String) LIKE '%{valoreTesto}%'"
                                                        End If

                                                        Dim dtFiltrato = dtOriginale.Clone()
                                                        For Each row As DataRow In dtOriginale.Select(filtro)
                                                            dtFiltrato.ImportRow(row)
                                                        Next
                                                        griglia.DataSource = dtFiltrato
                                                    End If
                                                End If
                                            End Sub

        Try
            Using conn As New SqlConnection(ConnString)
                Dim query = $"SELECT * FROM [{campo.TabellaElenco}]"
                Using da As New SqlDataAdapter(query, conn)
                    da.Fill(dtOriginale)
                    griglia.DataSource = dtOriginale
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Errore nel caricamento elenco: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try

        form.ShowDialog(Me)
    End Sub

    Private Function CreaDatePickerConGestioneVuoto(campo As CampoDatabase) As Control
        Dim dtp As New DateTimePicker()
        dtp.Format = DateTimePickerFormat.Custom
        dtp.CustomFormat = " "
        dtp.Width = 140
        dtp.Tag = campo.Nome

        AddHandler dtp.ValueChanged, Sub()
                                         dtp.CustomFormat = "dd/MM/yyyy"
                                         dtp.Tag = campo.Nome
                                     End Sub

        Dim menu As New ContextMenuStrip()
        menu.Items.Add("Cancella data", Nothing, Sub()
                                                     dtp.CustomFormat = " "
                                                     dtp.Value = dtp.MaxDate
                                                     dtp.Tag = campo.Nome & "|NULL"
                                                 End Sub)
        dtp.ContextMenuStrip = menu

        Return dtp
    End Function

    Private Function CreaTextBoxConGestioneTesto(campo As CampoDatabase) As Control

        Dim tipoCampo = campo.Tipo.ToLower()
        Dim isCampoLungo = tipoCampo.Contains("max") OrElse tipoCampo = "text" OrElse tipoCampo.Contains("varchar(max)")

        If campo.Nome.ToLower().Contains("password") Then
            Dim txt = New TextBox() With {
                .Width = campo.Lunghezza,
                .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
                .UseSystemPasswordChar = True,
                .Margin = New Padding(5)
            }
            AddHandler txt.KeyDown, AddressOf TextBoxPassword_KeyDown
            AddHandler txt.MouseDown, AddressOf TextBoxPassword_MouseDown
            Return txt
        End If

        If isCampoLungo Then
            Return New TextBox() With {
                .Width = campo.Lunghezza,
                .Height = 100,
                .Multiline = True,
                .ScrollBars = ScrollBars.Vertical,
                .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
                .Margin = New Padding(5)
            }
        End If

        Return New TextBox() With {
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Margin = New Padding(5)
        }
    End Function

    Private Function CreaLabelErrore(testo As String) As Control
        Return New Label() With {
            .Text = testo,
            .ForeColor = Color.Red,
            .AutoSize = True,
            .Margin = New Padding(5)
        }
    End Function

    Private Function CreaTextBoxIdentity(campo As CampoDatabase) As TextBox
        Dim txt As New TextBox()
        txt.ReadOnly = True
        txt.Enabled = False
        txt.BackColor = Color.LightGray
        txt.Tag = "identity"
        Return txt
    End Function

    Private Function CreaComboDaTabella(campo As CampoDatabase) As Control
        Dim combo As New ComboBox With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Width = campo.Lunghezza,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Tag = campo,
            .Margin = New Padding(5)
        }

        Try
            Using conn As New SqlConnection(ConnString)
                conn.Open()
                Dim query = $"SELECT {campo.CampoValore}, {campo.CampoVisuale} FROM {campo.TabellaCollegata}"
                Using cmd As New SqlCommand(query, conn)
                    Using reader = cmd.ExecuteReader()
                        Dim dt As New DataTable()
                        dt.Load(reader)
                        dt.Columns.Add("VisualeCombo", GetType(String))
                        For Each row As DataRow In dt.Rows
                            row("VisualeCombo") = $"{row(campo.CampoValore)} - {row(campo.CampoVisuale)}"
                        Next
                        combo.DataSource = dt
                        combo.DisplayMember = "VisualeCombo"
                        combo.ValueMember = campo.CampoValore
                    End Using
                End Using
            End Using
        Catch ex As Exception
            combo.Items.Clear()
            combo.Items.Add("Errore nel caricamento")
            combo.Enabled = False
        End Try

        Return combo
    End Function

    Private Function CreaDatePicker(valoreCampo As Object) As Control
        Dim campo As New CampoDatabase
        Dim dtPicker As New DateTimePicker With {
            .Width = campo.Lunghezza,
            .Format = DateTimePickerFormat.Short,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .Margin = New Padding(5)
        }

        If Not IsDBNull(valoreCampo) AndAlso valoreCampo IsNot Nothing Then
            dtPicker.Value = CDate(valoreCampo)
        Else
            dtPicker.Format = DateTimePickerFormat.Custom
            dtPicker.CustomFormat = " "
        End If

        Return dtPicker
    End Function

    Private Function CreaCheckBox() As Control
        Return New CheckBox() With {
            .Text = "",
            .AutoSize = True,
            .Anchor = AnchorStyles.Left,
            .Margin = New Padding(5)
        }
    End Function

    Private Function CreaTextBoxNumerico() As Control
        Dim campo As New CampoDatabase
        Return New TextBox() With {
            .Width = campo.Lunghezza,
            .MaximumSize = New Size(125, 0),
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .TextAlign = HorizontalAlignment.Right,
            .Margin = New Padding(5)
        }
    End Function

    Private Function CreaNumericUpDown() As Control
        Dim campo As New CampoDatabase
        Return New NumericUpDown() With {
            .Width = campo.Lunghezza,
            .Maximum = 1000000,
            .Minimum = 0,
            .DecimalPlaces = 2,
            .Increment = 0.01D,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .TextAlign = HorizontalAlignment.Right,
            .Margin = New Padding(5)
        }
    End Function

    Private Function CreaPannelloMultimediale() As Control
        Dim pannello As New FlowLayoutPanel With {
            .AutoSize = True,
            .FlowDirection = FlowDirection.LeftToRight
        }

        Dim txtFileName As New TextBox With {
            .Width = 250,
            .Text = ""
        }
        pannello.Controls.Add(txtFileName)

        Dim btnView As New Button With {
            .Text = "Visualizza",
            .AutoSize = True,
            .Enabled = True
        }

        AddHandler txtFileName.TextChanged, Sub()
                                                btnView.Enabled = Not String.IsNullOrWhiteSpace(txtFileName.Text)
                                            End Sub

        AddHandler btnView.Click, Sub(sender, e)
                                      If dgvDati Is Nothing OrElse dgvDati.SelectedRows.Count = 0 Then
                                          Dim parent = If(Me.MdiParent, Me)
                                          MDIMessageBox.Show("Seleziona prima una riga nella griglia.", parent, MessageBoxButtons.OK)
                                          Return
                                      End If

                                      Dim nomeFile = txtFileName.Text.Trim()
                                      Dim percorso = OttieniPercorsoImgVid(nomeFile)
                                      If String.IsNullOrWhiteSpace(percorso) Then
                                          MDIMessageBox.Show("Percorso multimediale non configurato.", Me.MdiParent, MessageBoxButtons.OK)
                                          Return
                                      End If

                                      Dim fullPath = Path.Combine(percorso, nomeFile)
                                      If Not File.Exists(fullPath) Then
                                          MDIMessageBox.Show("File non trovato: " & fullPath, Me.MdiParent, MessageBoxButtons.OK)
                                          Return
                                      End If

                                      If visualFormsAttivi.ContainsKey(fullPath) Then
                                          Dim formEsistente = visualFormsAttivi(fullPath)
                                          If Not formEsistente.IsDisposed Then
                                              formEsistente.BringToFront()
                                              formEsistente.Focus()
                                              Return
                                          Else
                                              visualFormsAttivi.Remove(fullPath)
                                          End If
                                      End If

                                      Dim viewer As New VisualMediaForm(fullPath)
                                      visualFormsAttivi(fullPath) = viewer

                                      AddHandler viewer.FormClosed, Sub(senderClosed, args)
                                                                        If visualFormsAttivi.ContainsKey(fullPath) Then
                                                                            visualFormsAttivi.Remove(fullPath)
                                                                        End If
                                                                    End Sub

                                      viewer.Show()
                                  End Sub

        pannello.Controls.Add(btnView)
        Return pannello
    End Function

    Private Sub PulisciCampi()
        For Each ctrl As Control In Me.Controls
            PulisciControllo(ctrl)
        Next
    End Sub

    Private Sub PulisciControllo(ctrl As Control)
        If TypeOf ctrl Is TextBox Then
            CType(ctrl, TextBox).Clear()

        ElseIf TypeOf ctrl Is ComboBox Then
            Dim combo = CType(ctrl, ComboBox)
            combo.SelectedIndex = -1
            combo.SelectedItem = Nothing

        ElseIf TypeOf ctrl Is DateTimePicker Then
            CType(ctrl, DateTimePicker).Value = DateTime.Today

        ElseIf TypeOf ctrl Is CheckBox Then
            CType(ctrl, CheckBox).Checked = False

        ElseIf ctrl.HasChildren Then
            For Each child As Control In ctrl.Controls
                PulisciControllo(child)
            Next
        End If
    End Sub

    Private Function OttieniPercorsoImgVid(NomeFile As String) As String

        If String.IsNullOrWhiteSpace(NomeFile) OrElse NomeFile.Length < 2 Then
            Return ""
        End If

        Try
            Using conn As New SqlConnection(ConnString)
                conn.Open()

                Dim codiceTipo As String = NomeFile.Substring(0, 2).ToUpper().Trim()
                Dim query As String = "SELECT Valore FROM Sys_Parametri WHERE Descrizione = @DescPar"

                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@DescPar", "Percorso" & codiceTipo)

                    Dim result As Object = cmd.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not Convert.IsDBNull(result) Then
                        Return result.ToString()
                    End If
                End Using
            End Using
        Catch ex As Exception
            Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show("Errore nel recupero del percorso: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)))
            Return ""
        End Try

        Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show("Nella Tabella Sys_Parametri non è stato trovato nessun risultato", Me.MdiParent, MessageBoxButtons.OK)))
        Return ""
    End Function

    Private Sub TextBoxPassword_KeyDown(sender As Object, e As KeyEventArgs)
        If e.Control AndAlso (e.KeyCode = Keys.C OrElse e.KeyCode = Keys.V OrElse e.KeyCode = Keys.X) Then
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub TextBoxPassword_MouseDown(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Right Then
            CType(sender, TextBox).ContextMenuStrip = New ContextMenuStrip()
        End If
    End Sub

    Private Sub CaricaBottoniDinamici()

        Dim formNameCorrente As String = Me.Name
        Dim query As String = "SELECT * FROM Sys_BottoniDinamici WHERE FormName = @formName ORDER BY Ordine"

        Try
            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@formName", formNameCorrente)

                    conn.Open()
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim bottoneDinamico = reader("BottoneDinamico").ToString()
                            Dim buttonText = reader("ButtonText").ToString()
                            Dim campoChiavePadre = reader("CampoChiavePadre").ToString()
                            Dim campoChiaveFiglia = reader("CampoChiaveFiglia").ToString()

                            Dim btnDinamico As New Button With {
                                .Text = buttonText,
                                .AutoSize = True,
                                .Margin = New Padding(5),
                                .Tag = New With {
                                    Key .FormName = bottoneDinamico,
                                    Key .CampoPadre = campoChiavePadre,
                                    Key .CampoFiglia = campoChiaveFiglia,
                                    Key .Titolo = buttonText
                                }
                            }

                            AddHandler btnDinamico.Click, Sub(s, e)
                                                              If dgvDati.SelectedRows.Count = 0 Then
                                                                  MDIMessageBox.Show("Seleziona prima una riga dalla griglia per aprire il form collegato.", Me.MdiParent, MessageBoxButtons.OK)
                                                                  Return
                                                              End If

                                                              Dim info = CType(CType(s, Button).Tag, Object)
                                                              Dim valoreCella = dgvDati.SelectedRows(0).Cells(info.CampoPadre).Value?.ToString()
                                                              Dim valoreChiavePadre = If(Not String.IsNullOrEmpty(valoreCella) AndAlso valoreCella.Contains("-"), valoreCella.Split("-"c)(0).Trim(), valoreCella)

                                                              If String.IsNullOrWhiteSpace(valoreChiavePadre) Then
                                                                  MDIMessageBox.Show("Il valore della chiave primaria selezionata è nullo o non valido.", Me.MdiParent, MessageBoxButtons.OK)
                                                                  Return
                                                              End If

                                                              For Each f As Form In GesPu25.MdiChildren
                                                                  If TypeOf f Is DynamicDataForm AndAlso f.Name = info.FormName Then
                                                                      f.WindowState = FormWindowState.Normal
                                                                      f.BringToFront()
                                                                      f.Activate()
                                                                      Return
                                                                  End If
                                                              Next

                                                              Dim campiFigli = RecuperaCampiDa(info.FormName)
                                                              Dim nuovoForm As New DynamicDataForm(campiFigli, info.FormName)
                                                              nuovoForm.MdiParent = GesPu25
                                                              nuovoForm.Text = $"{info.Titolo} - Filtrato per {info.CampoPadre} = {valoreChiavePadre}"
                                                              nuovoForm.FiltroIniziale = $"{info.CampoFiglia} = '{valoreChiavePadre}'"
                                                              nuovoForm.Show()

                                                              Dim campoCollegamento = info.CampoFiglia
                                                              If nuovoForm.campoInputs.ContainsKey(campoCollegamento) Then
                                                                  Dim ctrl = nuovoForm.campoInputs(campoCollegamento)

                                                                  Select Case True
                                                                      Case TypeOf ctrl Is ComboBox
                                                                          CType(ctrl, ComboBox).SelectedValue = valoreChiavePadre

                                                                      Case TypeOf ctrl Is TextBox
                                                                          CType(ctrl, TextBox).Text = valoreChiavePadre

                                                                      Case TypeOf ctrl Is FlowLayoutPanel
                                                                          Dim txt = ctrl.Controls.OfType(Of TextBox).FirstOrDefault()
                                                                          If txt IsNot Nothing Then txt.Text = valoreChiavePadre
                                                                  End Select
                                                              End If

                                                          End Sub

                            panelBottoni.Controls.Add(btnDinamico)
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show("Errore nel caricamento bottoni dinamici: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)))
        End Try
    End Sub

    Private Function IntestazioneMultilinea(nomeColonna As String) As String
        If String.IsNullOrWhiteSpace(nomeColonna) Then Return ""

        Dim sb As New StringBuilder()
        sb.Append(nomeColonna(0))

        For i = 1 To nomeColonna.Length - 1
            Dim c = nomeColonna(i)
            If Char.IsUpper(c) Then
                sb.Append(vbCrLf)
            End If
            sb.Append(c)
        Next

        Return sb.ToString()
    End Function

    Private Sub CaricaDatiTabella(nomeTabella As String)
        Dim query As String = $"SELECT * FROM [{nomeTabella}]"
        If Not String.IsNullOrWhiteSpace(FiltroIniziale) Then
            query &= $" WHERE {FiltroIniziale}"
        End If

        Try
            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand(query, conn)
                    conn.Open()
                    Dim adapter As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    adapter.Fill(dt)

                    dgvDati.DataSource = dt

                    Dim dtCollegamenti As DataTable = EseguiQuery($"
                        SELECT NomeCampo FROM Sys_CollegamentiCampi
                        WHERE NomeTabella = '{nomeTabella}'")

                    For Each col As DataGridViewColumn In dgvDati.Columns
                        Dim nomeCampo = col.Name
                        Dim etichetta = GetEtichetta(Me.Name, nomeCampo)
                        col.HeaderText = etichetta
                    Next
                End Using
            End Using
        Catch ex As Exception
            Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show("Errore nel caricamento dei dati della tabella." & vbCrLf & ex.Message, Me.MdiParent, MessageBoxButtons.OK)))
        End Try
    End Sub

    Public Function EseguiQuery(query As String) As DataTable
        Dim dt As New DataTable()

        Try
            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand(query, conn)
                    cmd.CommandTimeout = 60
                    conn.Open()
                    Using adapter As New SqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                End Using
            End Using
        Catch ex As SqlException
            Trace.TraceError("Errore SQL: " & ex.Message & " Query: " & query)
        Catch ex As Exception
            Trace.TraceError("Errore generico: " & ex.Message)
        End Try

        Return dt
    End Function

    Private Sub dgvDati_SelectionChanged(sender As Object, e As EventArgs)
        If isUpdatingControls Then Return
        If dgvDati.SelectedRows.Count = 0 Then Exit Sub

        CaricaDatiNeiControlli(dgvDati.SelectedRows(0))
    End Sub

    Private Sub dgvDati_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        If isUpdatingControls Then Return
        CaricaDatiNeiControlli(dgvDati.Rows(e.RowIndex))
        ModalitaCorrente = "visualizzazione"
        lblModalita.Text = "Visualizzazione in corso..."
        UpdateButtonsByModalita()
    End Sub

    Private Function GetEtichetta(nomeTabella As String, nomeColonna As String) As String
        Dim etichetta As String = ""

        Using conn As New SqlConnection(ConnString)
            conn.Open()

            Dim querySel = "SELECT TestoEtichetta FROM Sys_TestoEtichetta WHERE NomeTabella = @tab AND NomeColonna = @col"
            Using cmdSel As New SqlCommand(querySel, conn)
                cmdSel.Parameters.AddWithValue("@tab", nomeTabella)
                cmdSel.Parameters.AddWithValue("@col", nomeColonna)

                Dim risultato = cmdSel.ExecuteScalar()
                If risultato IsNot Nothing AndAlso Not Convert.IsDBNull(risultato) Then
                    etichetta = risultato.ToString()
                Else
                    etichetta = SpaziaMaiuscole(nomeColonna)

                    Dim queryIns = "INSERT INTO Sys_TestoEtichetta (NomeTabella, NomeColonna, TestoEtichetta) VALUES (@tab, @col, @txt)"
                    Using cmdIns As New SqlCommand(queryIns, conn)
                        cmdIns.Parameters.AddWithValue("@tab", nomeTabella)
                        cmdIns.Parameters.AddWithValue("@col", nomeColonna)
                        cmdIns.Parameters.AddWithValue("@txt", etichetta)
                        cmdIns.ExecuteNonQuery()
                    End Using
                End If
            End Using
        End Using

        Return etichetta
    End Function

    Private Function SpaziaMaiuscole(text As String) As String
        If String.IsNullOrWhiteSpace(text) Then Return ""
        Dim sb As New StringBuilder()
        sb.Append(text(0))
        For i = 1 To text.Length - 1
            Dim c = text(i)
            If Char.IsUpper(c) Then sb.Append(" ")
            sb.Append(c)
        Next
        Return sb.ToString()
    End Function

    Private Sub ApplicaAutorizzazioni(nomeUtente As String)
        Try
            Using conn As New SqlConnection(ConnString)
                conn.Open()

                Dim isAdmin As Boolean = False
                Dim queryAdmin = "SELECT ISNULL(Amministratore, 0) FROM Sys_Utenti WHERE NomeUtente = @utente"

                Using cmdAdmin As New SqlCommand(queryAdmin, conn)
                    cmdAdmin.Parameters.AddWithValue("@utente", nomeUtente)
                    isAdmin = Convert.ToBoolean(cmdAdmin.ExecuteScalar())
                End Using

                If isAdmin Then
                    For Each ctrl As Control In panelBottoni.Controls
                        If TypeOf ctrl Is Button Then
                            Dim btn As Button = CType(ctrl, Button)
                            If String.Equals(btn.Text, "Salva", StringComparison.OrdinalIgnoreCase) Then
                                btn.Enabled = False
                            Else
                                btn.Enabled = True
                            End If
                        End If
                    Next
                    Return
                End If

                Dim queryAut = "
                SELECT CanInsert, CanUpdate, CanDelete 
                FROM Sys_Autorizzazioni 
                WHERE NomeUtente = @utente AND Form = @form
            "

                Using cmd As New SqlCommand(queryAut, conn)
                    cmd.Parameters.AddWithValue("@utente", nomeUtente)
                    cmd.Parameters.AddWithValue("@form", Me.Name)

                    Using reader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Dim canInsert = Convert.ToBoolean(reader("CanInsert"))
                            Dim canUpdate = Convert.ToBoolean(reader("CanUpdate"))
                            Dim canDelete = Convert.ToBoolean(reader("CanDelete"))

                            DisabilitaPulsante("Inserisci", Not canInsert)
                            DisabilitaPulsante("Modifica", Not canUpdate)
                            DisabilitaPulsante("Cancella", Not canDelete)
                        Else
                            DisabilitaPulsante("Inserisci", True)
                            DisabilitaPulsante("Modifica", True)
                            DisabilitaPulsante("Cancella", True)
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show($"Errore nel controllo autorizzazioni: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)))
        End Try
    End Sub

    Private Sub DisabilitaPulsante(nomeBottone As String, disattiva As Boolean)
        For Each ctrl In panelBottoni.Controls
            If TypeOf ctrl Is Button Then
                Dim btn = CType(ctrl, Button)
                If btn.Text.ToLower() = nomeBottone.ToLower() Then
                    btn.Enabled = Not disattiva
                    Exit For
                End If
            End If
        Next
    End Sub

    Private Sub ApplicaVisualizzazioneColonne()
        Dim tabella As String = Me.Name
        Dim nomeGriglia As String = dgvDati.Name

        Dim querySelect = "SELECT NomeColonna FROM Sys_VisualizzaInDbgrid WHERE NomeTabella = @tab AND NomeDbgrid = @grid AND VisualizzaInDbgrid = 1"
        Dim queryCheck = "SELECT COUNT(*) FROM Sys_VisualizzaInDbgrid WHERE NomeTabella = @tab AND NomeDbgrid = @grid"
        Dim queryInsert = "INSERT INTO Sys_VisualizzaInDbgrid (NomeTabella, NomeColonna, NomeDbgrid, VisualizzaInDbgrid) VALUES (@tab, @col, @grid, 1)"

        Dim colonneDaVisualizzare As New HashSet(Of String)

        Using conn As New SqlConnection(ConnString)
            conn.Open()

            Dim esistonoRecord As Boolean
            Using cmdCheck As New SqlCommand(queryCheck, conn)
                cmdCheck.Parameters.AddWithValue("@tab", tabella)
                cmdCheck.Parameters.AddWithValue("@grid", nomeGriglia)
                esistonoRecord = Convert.ToInt32(cmdCheck.ExecuteScalar()) > 0
            End Using

            If esistonoRecord Then
                Using cmdSel As New SqlCommand(querySelect, conn)
                    cmdSel.Parameters.AddWithValue("@tab", tabella)
                    cmdSel.Parameters.AddWithValue("@grid", nomeGriglia)
                    Using reader = cmdSel.ExecuteReader()
                        While reader.Read()
                            colonneDaVisualizzare.Add(reader.GetString(0))
                        End While
                    End Using
                End Using

                For Each col As DataGridViewColumn In dgvDati.Columns
                    col.Visible = colonneDaVisualizzare.Contains(col.Name)
                Next
            Else
                For Each col As DataGridViewColumn In dgvDati.Columns
                    Using cmdIns As New SqlCommand(queryInsert, conn)
                        cmdIns.Parameters.AddWithValue("@tab", tabella)
                        cmdIns.Parameters.AddWithValue("@col", col.Name)
                        cmdIns.Parameters.AddWithValue("@grid", nomeGriglia)
                        cmdIns.ExecuteNonQuery()
                    End Using
                    col.Visible = True
                Next
            End If
        End Using
    End Sub

    Private Sub SalvaConfigurazioneGrigliaBatched(dgvdati As DataGridView)
        If dgvdati Is Nothing OrElse dgvdati.Columns.Count = 0 Then Return

        Dim nomeDbgrid As String = dgvdati.Name

        Dim righe As New List(Of (NomeColonna As String, ColWidth As Integer, Visualizza As Integer))
        For Each col As DataGridViewColumn In dgvdati.Columns
            righe.Add((col.Name, col.Width, If(col.Visible, 1, 0)))
        Next

        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Using tx = conn.BeginTransaction()
                Try
                    Using cmdDel As New SqlCommand("DELETE FROM Sys_VisualizzaInDbgrid WHERE NomeTabella = @NomeTabella AND NomeDbgrid = @NomeDbgrid", conn, tx)
                        cmdDel.Parameters.AddWithValue("@NomeTabella", nomeTabellaCorrente)
                        cmdDel.Parameters.AddWithValue("@NomeDbgrid", nomeDbgrid)
                        cmdDel.ExecuteNonQuery()
                    End Using

                    Using cmdIns As New SqlCommand("
                    INSERT INTO Sys_VisualizzaInDbgrid (NomeTabella, NomeColonna, NomeDbgrid, ColWidth, VisualizzaInDbgrid)
                    VALUES (@NomeTabella, @NomeColonna, @NomeDbgrid, @ColWidth, @VisualizzaInDbgrid)", conn, tx)

                        cmdIns.Parameters.Add("@NomeTabella", SqlDbType.NVarChar, 200)
                        cmdIns.Parameters.Add("@NomeColonna", SqlDbType.NVarChar, 200)
                        cmdIns.Parameters.Add("@NomeDbgrid", SqlDbType.NVarChar, 200)
                        cmdIns.Parameters.Add("@ColWidth", SqlDbType.Int)
                        cmdIns.Parameters.Add("@VisualizzaInDbgrid", SqlDbType.Int)

                        For Each r In righe
                            cmdIns.Parameters("@NomeTabella").Value = nomeTabellaCorrente
                            cmdIns.Parameters("@NomeColonna").Value = r.NomeColonna
                            cmdIns.Parameters("@NomeDbgrid").Value = nomeDbgrid
                            cmdIns.Parameters("@ColWidth").Value = r.ColWidth
                            cmdIns.Parameters("@VisualizzaInDbgrid").Value = r.Visualizza
                            cmdIns.ExecuteNonQuery()
                        Next
                    End Using

                    tx.Commit()
                Catch ex As Exception
                    Try
                        tx.Rollback()
                    Catch
                    End Try
                    Throw
                End Try
            End Using
        End Using
    End Sub

    Private Sub ApplicaConfigurazioneGriglia(dgv As DataGridView)
        If dgv Is Nothing OrElse dgv.Columns.Count = 0 Then Exit Sub

        dgv.SuspendLayout()
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        Try
            Using conn As New SqlConnection(ConnString)
                conn.Open()

                Dim cmd As New SqlCommand("
                SELECT NomeColonna, ColWidth, VisualizzaInDbgrid 
                FROM Sys_VisualizzaInDbgrid 
                WHERE NomeTabella = @NomeTabella AND NomeDbgrid = @NomeDbgrid", conn)
                cmd.Parameters.AddWithValue("@NomeTabella", nomeTabellaCorrente)
                cmd.Parameters.AddWithValue("@NomeDbgrid", dgv.Name)

                Dim reader = cmd.ExecuteReader()
                Dim listaConfig As New List(Of (Nome As String, Width As Integer, Visible As Boolean))
                While reader.Read()
                    Dim n As String = reader("NomeColonna").ToString()
                    Dim w As Integer = If(IsDBNull(reader("ColWidth")), 0, Convert.ToInt32(reader("ColWidth")))
                    Dim v As Boolean = If(IsDBNull(reader("VisualizzaInDbgrid")), True, Convert.ToBoolean(reader("VisualizzaInDbgrid")))
                    listaConfig.Add((n, w, v))
                End While
                reader.Close()

                Dim mapConfig = listaConfig.ToDictionary(Function(c) c.Nome, Function(c) (c.Width, c.Visible), StringComparer.OrdinalIgnoreCase)
                For Each col As DataGridViewColumn In dgv.Columns
                    If mapConfig.ContainsKey(col.Name) Then
                        Dim cfg = mapConfig(col.Name)
                        col.Visible = cfg.Visible
                    Else
                        col.Visible = True
                    End If
                Next

                For Each cfg In listaConfig
                    Dim col = dgv.Columns.Cast(Of DataGridViewColumn)().FirstOrDefault(Function(c) String.Equals(c.Name, cfg.Nome, StringComparison.OrdinalIgnoreCase))
                    If col Is Nothing Then Continue For

                    Try
                        If cfg.Width > 0 Then
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = cfg.Width
                        End If
                    Catch ex As Exception
                        Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show($"Errore su colonna '{cfg.Nome}' (apply saved width): {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)))
                    End Try
                Next

                For Each cfg In listaConfig
                    If cfg.Width > 0 Then Continue For

                    Dim col = dgv.Columns.Cast(Of DataGridViewColumn)().FirstOrDefault(Function(c) String.Equals(c.Name, cfg.Nome, StringComparison.OrdinalIgnoreCase))
                    If col Is Nothing OrElse Not col.Visible Then Continue For

                    Try
                        dgv.AutoResizeColumn(col.Index, DataGridViewAutoSizeColumnMode.AllCells)
                        Dim computed = col.Width
                        col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                        col.Width = computed
                    Catch ex As Exception
                        Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show($"Errore nel calcolo larghezza per '{cfg.Nome}': {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)))
                    End Try
                Next
            End Using
        Finally
            dgv.ResumeLayout()
            dgv.Refresh()
        End Try
    End Sub

    Private Sub PosizionaGrigliaDaSysForm()

        If dgvDati Is Nothing Then Exit Sub

        Dim posizioneX As Integer = 10

        Using conn As New SqlConnection(ConnString)
            conn.Open()

            Dim cmd As New SqlCommand("
            SELECT UltimoWidth 
            FROM Sys_form 
            WHERE FormName = @NomeTabella", conn)
            cmd.Parameters.AddWithValue("@NomeTabella", nomeTabellaCorrente)

            Dim result = cmd.ExecuteScalar()
            If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                posizioneX = Convert.ToInt32(result)
            End If
        End Using

        Const margineLaterale As Integer = 4
        dgvDati.Location = New Point(posizioneX, margineLaterale)
        dgvDati.Size = New Size(Math.Max(Me.ClientSize.Width - posizioneX - margineLaterale, 100), Me.ClientSize.Height - 2 * margineLaterale)
    End Sub

    Private Function TrovaControlloPiùADestra(root As Control) As Control()
        Dim lista = New List(Of Control)

        For Each c In root.Controls
            If c.Visible AndAlso Not TypeOf c Is DataGridView Then
                lista.Add(c)
            End If
            lista.AddRange(TrovaControlloPiùADestra(c))
        Next

        Return lista.ToArray()
    End Function

    Private Function RecuperaCampiCalcolati() As Dictionary(Of String, (Formula As String, TipoValore As String))
        Dim dizionario As New Dictionary(Of String, (String, String))

        Using conn As New SqlConnection(ConnString)
            Dim query = "SELECT NomeCampo, Formula, Tipovalore FROM Sys_CampiCalcolati WHERE NomeTabella = @NomeTabella"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@NomeTabella", Me.Name)

                conn.Open()
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim nomeCampo = reader("NomeCampo").ToString()
                        Dim formula = reader("Formula").ToString()
                        Dim tipoValore = reader("Tipovalore").ToString().ToLower()
                        dizionario(nomeCampo) = (formula, tipoValore)
                    End While
                End Using
            End Using
        End Using

        Return dizionario
    End Function

    Private Function CalcolaValoriCampiCalcolati(formule As Dictionary(Of String, String), tipiValore As Dictionary(Of String, String)) As Dictionary(Of String, Object)

        Dim risultati As New Dictionary(Of String, Object)

        For Each kvp In formule
            Dim nomeCampo = kvp.Key
            Dim formula = kvp.Value
            Dim tipoValore = If(tipiValore.ContainsKey(nomeCampo), tipiValore(nomeCampo).ToLower(), "numero")

            Dim tipiAmmessi As String() = {"numero", "stringa", "data", "booleano", "raw", "testo"}
            If Not tipiAmmessi.Contains(tipoValore) Then
                tipoValore = "numero"
            End If

            Try
                Dim espressione = formula

                For Each vInput In campoInputs
                    Dim nome = vInput.Key
                    Dim valore = EstraiValoreDaControllo(nome, vInput.Value)

                    Dim valoreStringa As String
                    If valore Is Nothing OrElse valore Is DBNull.Value Then
                        If tipoValore = "stringa" Then
                            valoreStringa = """" & "" & """"
                        Else
                            valoreStringa = "0"
                        End If
                    ElseIf TypeOf valore Is DateTime Then
                        valoreStringa = $"""{CDate(valore).ToString("yyyy-MM-dd")}"""
                    ElseIf tipoValore = "stringa" Then
                        valoreStringa = $"""{valore.ToString().Replace("""", """""")}"""
                    Else
                        valoreStringa = Convert.ToString(valore, Globalization.CultureInfo.InvariantCulture)
                    End If

                    Dim regex As Regex = Nothing
                    If regexCache.ContainsKey(nome) Then
                        regex = regexCache(nome)
                    Else
                        regex = New Regex($"\b{Regex.Escape(nome)}\b", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
                        regexCache(nome) = regex
                    End If
                    espressione = regex.Replace(espressione, valoreStringa)
                Next

                Dim risultato As Object
                If tipoValore = "numero" Then
                    risultato = New DataTable().Compute(espressione, Nothing)
                    If IsNumeric(risultato) Then
                        risultato = Convert.ToDouble(risultato, Globalization.CultureInfo.InvariantCulture)
                    Else
                        risultato = Nothing
                    End If
                ElseIf tipoValore = "stringa" Then
                    risultato = ValutaEspressioneStringa(espressione)
                Else
                    risultato = espressione
                End If

                risultati(nomeCampo) = risultato
            Catch ex As Exception
                Trace.TraceError($"[Campo calcolato] Errore nel calcolo di '{nomeCampo}': {ex.Message}")
                risultati(nomeCampo) = Nothing
            End Try
        Next

        Return risultati
    End Function

    Private Function ValutaEspressioneStringa(expr As String) As String
        Try
            Dim dt As New DataTable()
            dt.Columns.Add("Expr", GetType(String), expr)
            Dim row = dt.NewRow()
            dt.Rows.Add(row)
            Return row("Expr").ToString()
        Catch ex As Exception
            Trace.TraceError($"[StringEval] Errore: {ex.Message}")
            Return ""
        End Try
    End Function

    ' Helpers UI per lock durante salvataggio
    Private Sub ToggleUIForSaving(saving As Boolean)
        Me.BeginInvoke(New MethodInvoker(Sub()
                                             For Each c As Control In panelBottoni.Controls
                                                 c.Enabled = Not saving
                                             Next
                                             For Each c As Control In panelBottoniDinamici.Controls
                                                 c.Enabled = Not saving
                                             Next
                                             If saving Then
                                                 lblModalita.Text = "Salvataggio in corso..."
                                                 lblModalita.ForeColor = Color.Orange
                                             Else
                                                 lblModalita.Text = ""
                                                 lblModalita.ForeColor = Color.DarkGreen
                                             End If
                                         End Sub))
    End Sub

    Private Sub EsportaPDF(sender As Object, e As EventArgs)
        Try
            Dim document As New PdfDocument()
            document.Info.Title = $"Esportazione dati: {Me.Name}"

            Dim page As PdfPage = document.AddPage()
            page.Orientation = PageOrientation.Landscape
            Dim gfx As XGraphics = XGraphics.FromPdfPage(page)
            Dim font As New XFont("Arial", 8, XFontStyleEx.Regular)
            Dim fontBold As New XFont("Arial", 8, XFontStyleEx.Bold)
            Dim formatter As New XTextFormatter(gfx)

            Dim margin As Double = 40
            Dim topOffset As Double = 60
            Dim lineHeight As Double = 20
            Dim pageHeight As Double = page.Height.Point
            Dim usableWidth As Double = page.Width.Point - (2 * margin)

            Dim colonne = dgvDati.Columns.Cast(Of DataGridViewColumn).Where(Function(c) c.Visible).ToList()
            Dim colCount = colonne.Count
            If colCount = 0 Then
                Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show("Nessuna colonna visibile da esportare.", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Information)))
                Return
            End If

            Dim colWidth As Double = usableWidth / colCount

            gfx.DrawString($"Esportazione dati: {Me.Name}", New XFont("Arial", 11, XFontStyleEx.Bold), XBrushes.Black, New XPoint(margin, topOffset))
            topOffset += 30

            For i = 0 To colCount - 1
                Dim header = colonne(i).HeaderText
                Dim rect As New XRect(margin + (i * colWidth), topOffset, colWidth, lineHeight)
                formatter.DrawString(header, fontBold, XBrushes.DarkBlue, rect, XStringFormats.TopLeft)
            Next
            topOffset += lineHeight

            For Each row As DataGridViewRow In dgvDati.Rows
                If topOffset + (lineHeight * 2) > pageHeight - margin Then
                    page = document.AddPage()
                    page.Orientation = PageOrientation.Landscape
                    gfx = XGraphics.FromPdfPage(page)
                    formatter = New XTextFormatter(gfx)
                    topOffset = margin
                    For i = 0 To colCount - 1
                        Dim header = colonne(i).HeaderText
                        Dim rect As New XRect(margin + (i * colWidth), topOffset, colWidth, lineHeight)
                        formatter.DrawString(header, fontBold, XBrushes.DarkBlue, rect, XStringFormats.TopLeft)
                    Next
                    topOffset += lineHeight
                End If

                For i = 0 To colCount - 1
                    Dim valore = If(row.Cells(colonne(i).Name).Value Is Nothing, "", row.Cells(colonne(i).Name).Value.ToString())
                    Dim rect As New XRect(margin + (i * colWidth), topOffset, colWidth, lineHeight * 2)
                    If row.Index Mod 2 = 0 Then gfx.DrawRectangle(XBrushes.LightGray, rect)
                    formatter.DrawString(valore, font, XBrushes.Black, rect, XStringFormats.TopLeft)
                Next
                topOffset += lineHeight * 2
            Next

            Dim filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"{Me.Name}_Esportazione.pdf")
            document.Save(filePath)

            Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show($"PDF esportato con successo:{Environment.NewLine}{filePath}", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Information)))
        Catch ex As Exception
            Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show("Errore durante l'esportazione PDF: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)))
        End Try
    End Sub

    Private Sub CaricaDatiNeiControlli(riga As DataGridViewRow)
        If riga Is Nothing Then Return

        Try
            isUpdatingControls = True

            Dim dtCollegamenti As DataTable = EseguiQuery($"
            SELECT NomeCampo FROM Sys_CollegamentiCampi
            WHERE NomeTabella = '{nomeTabellaCorrente}'")

            Dim campiCollegati As New HashSet(Of String)(
            dtCollegamenti.AsEnumerable().Select(Function(r) r("NomeCampo").ToString()),
            StringComparer.OrdinalIgnoreCase)

            For Each campoNome In campoInputs.Keys
                If Not dgvDati.Columns.Contains(campoNome) Then Continue For

                Dim valoreObj = riga.Cells(campoNome).Value
                Dim valoreRaw = If(valoreObj IsNot DBNull.Value AndAlso valoreObj IsNot Nothing, valoreObj.ToString(), "")
                Dim valore = If(campiCollegati.Contains(campoNome) AndAlso valoreRaw.Contains("-"c),
                           valoreRaw.Split("-"c)(0).Trim(),
                           valoreRaw)

                Dim ctrl = campoInputs(campoNome)
                Dim isPassword As Boolean = campoNome.ToLower().Contains("password")

                Select Case True
                    Case TypeOf ctrl Is TextBox
                        Dim txt As TextBox = CType(ctrl, TextBox)
                        txt.Text = If(isPassword, "", valore)

                    Case TypeOf ctrl Is CheckBox
                        Dim chk As CheckBox = CType(ctrl, CheckBox)
                        Dim booleano As Boolean
                        If Boolean.TryParse(valore, booleano) Then
                            chk.Checked = booleano
                        ElseIf IsNumeric(valore) Then
                            chk.Checked = (Convert.ToInt32(valore) <> 0)
                        Else
                            chk.Checked = False
                        End If

                    Case TypeOf ctrl Is ComboBox
                        ImpostaValoreCombo(CType(ctrl, ComboBox), valore)

                    Case TypeOf ctrl Is DateTimePicker
                        Dim dtPicker As DateTimePicker = CType(ctrl, DateTimePicker)
                        Dim dt As DateTime
                        If DateTime.TryParse(valore, dt) Then
                            dtPicker.Format = DateTimePickerFormat.Short
                            dtPicker.Value = dt
                            dtPicker.Checked = True
                        Else
                            dtPicker.Format = DateTimePickerFormat.Custom
                            dtPicker.CustomFormat = " "
                            dtPicker.Checked = False
                        End If

                    Case TypeOf ctrl Is FlowLayoutPanel
                        Dim txt As TextBox = ctrl.Controls.OfType(Of TextBox)().FirstOrDefault()
                        Dim lbl As Label = ctrl.Controls.OfType(Of Label)().FirstOrDefault()

                        If txt IsNot Nothing Then txt.Text = valore

                        Dim campoDef As CampoDatabase = TrovaDefinizioneCampo(campoNome)

                        If lbl IsNot Nothing Then
                            If campoDef IsNot Nothing AndAlso
                           Not String.IsNullOrEmpty(campoDef.TabellaElenco) AndAlso
                           Not String.IsNullOrEmpty(campoDef.ChiaveElenco) AndAlso
                           Not String.IsNullOrEmpty(campoDef.CampoVisuale) Then

                                Try
                                    Dim dtRef As DataTable = RecuperaTabellaCached(campoDef.TabellaElenco)
                                    If dtRef IsNot Nothing AndAlso dtRef.Columns.Contains(campoDef.ChiaveElenco) Then
                                        Dim filtro = $"{campoDef.ChiaveElenco} = '{valore.Replace("'", "''")}'"
                                        Dim rows = dtRef.Select(filtro)
                                        If rows.Length > 0 AndAlso dtRef.Columns.Contains(campoDef.CampoVisuale) Then
                                            lbl.Text = rows(0)(campoDef.CampoVisuale).ToString()
                                        Else
                                            lbl.Text = "..."
                                        End If
                                    Else
                                        lbl.Text = "..."
                                    End If
                                Catch ex As Exception
                                    lbl.Text = "..."
                                    Trace.TraceError($"Errore recupero descrizione per campo '{campoNome}': {ex.Message}")
                                End Try
                            Else
                                lbl.Text = "..."
                            End If
                        End If

                    Case Else
                        ' fallback: prova a impostare Text se possibile
                        Try
                            ctrl.Text = valore
                        Catch
                        End Try
                End Select
            Next

        Finally
            isUpdatingControls = False
        End Try
    End Sub

    Private Sub ImpostaValoreCombo(combo As ComboBox, valore As Object)
        If combo Is Nothing Then Return

        Try
            ' Se non c'è DataSource o ValueMember, reset e esci
            If combo.DataSource Is Nothing OrElse String.IsNullOrWhiteSpace(combo.ValueMember) Then
                combo.SelectedIndex = -1
                Return
            End If

            If valore Is Nothing OrElse Convert.IsDBNull(valore) Then
                combo.SelectedIndex = -1
                Return
            End If

            Dim stringVal As String = valore.ToString()

            ' Se il DataSource è un DataView/DataTable, cerchiamo il tipo della colonna ValueMember
            Dim targetType As Type = Nothing
            Dim dt As DataTable = TryCast(TryCast(combo.DataSource, DataView)?.Table, DataTable)
            If dt Is Nothing Then
                dt = TryCast(combo.DataSource, DataTable)
            End If

            If dt IsNot Nothing AndAlso dt.Columns.Contains(combo.ValueMember) Then
                targetType = dt.Columns(combo.ValueMember).DataType
            End If

            ' Scorri gli elementi per trovare una corrispondenza (evita eccezioni di bind)
            Dim found As Boolean = False
            For Each itemObj In combo.Items
                Dim drv = TryCast(itemObj, DataRowView)
                If drv IsNot Nothing Then
                    Dim cell = drv(combo.ValueMember)
                    If cell IsNot Nothing AndAlso Not IsDBNull(cell) AndAlso cell.ToString() = stringVal Then
                        found = True
                        Exit For
                    End If
                Else
                    ' se l'item è semplice (lista di valori), confronta ToString()
                    If itemObj IsNot Nothing AndAlso itemObj.ToString() = stringVal Then
                        found = True
                        Exit For
                    End If
                End If
            Next

            If Not found Then
                combo.SelectedIndex = -1
                Return
            End If

            ' Prova a convertire al tipo target quando noto
            If targetType IsNot Nothing Then
                Try
                    Dim converted = Convert.ChangeType(stringVal, targetType, Globalization.CultureInfo.InvariantCulture)
                    combo.SelectedValue = converted
                    Return
                Catch
                    ' fallback: assegna direttamente la stringa (spesso funziona per binding)
                    Try
                        combo.SelectedValue = stringVal
                        Return
                    Catch
                    End Try
                End Try
            Else
                ' Se non conosciamo il tipo target, assegna direttamente
                combo.SelectedValue = stringVal
            End If

        Catch ex As Exception
            Trace.TraceError($"ImpostaValoreCombo error: {ex.Message}")
            Try
                combo.SelectedIndex = -1
            Catch
            End Try
        End Try
    End Sub

    Private Sub UpdateButtonsByModalita()
        ' ModalitaCorrente expected values: "nessuna", "inserimento", "modifica"
        Dim modo = If(String.IsNullOrWhiteSpace(ModalitaCorrente), "nessuna", ModalitaCorrente.ToLowerInvariant())

        ' Default: tutti abilitati tranne Salva e Annulla
        Dim canInsert As Boolean = True
        Dim canEdit As Boolean = True
        Dim canDelete As Boolean = True
        Dim salvaEnabled As Boolean = False
        Dim annullaEnabled As Boolean = False

        Select Case modo
            Case "inserimento"
                salvaEnabled = True
                annullaEnabled = True
                canInsert = True
                canEdit = False
                canDelete = False
            Case "modifica"
                salvaEnabled = True
                annullaEnabled = True
                canInsert = False
                canEdit = True
                canDelete = False
            Case "nessuna"
                salvaEnabled = False
                annullaEnabled = False
                canInsert = True
                canEdit = False
                canDelete = False
            Case "visualizzazione"
                salvaEnabled = False
                annullaEnabled = False
                canInsert = True
                canEdit = True
                canDelete = True
            Case Else ' "nessuna" o altri
                salvaEnabled = False
                annullaEnabled = False
                canInsert = True
                canEdit = True
                canDelete = True
        End Select

        ' Applica alle UI (Invoke-safe)
        Me.BeginInvoke(New MethodInvoker(Sub()
                                             For Each ctrl As Control In panelBottoni.Controls
                                                 If TypeOf ctrl Is Button Then
                                                     Dim btn = CType(ctrl, Button)
                                                     Select Case btn.Text.ToLowerInvariant()
                                                         Case "inserisci"
                                                             btn.Enabled = canInsert
                                                         Case "modifica"
                                                             btn.Enabled = canEdit
                                                         Case "cancella"
                                                             btn.Enabled = canDelete
                                                         Case "salva"
                                                             btn.Enabled = salvaEnabled
                                                         Case "annulla"
                                                             btn.Enabled = annullaEnabled
                                                         Case Else
                                                             ' lascia gli altri pulsanti dinamici com'è
                                                     End Select
                                                 End If
                                             Next
                                         End Sub))
    End Sub


    Private Function TrovaDefinizioneCampo(nomeCampo As String) As CampoDatabase
        If String.IsNullOrWhiteSpace(nomeCampo) Then
            Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show("Il nome del campo è vuoto o nullo.", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Warning)))
            Return Nothing
        End If

        If campiDefiniti Is Nothing OrElse campiDefiniti.Count = 0 Then
            Trace.TraceWarning($"TrovaDefinizioneCampo: lista campiDefiniti vuota; richiesta per '{nomeCampo}'")
            Return Nothing
        End If

        For Each campo As CampoDatabase In campiDefiniti
            If campo IsNot Nothing AndAlso
           Not String.IsNullOrWhiteSpace(campo.Nome) AndAlso
           campo.Nome.Trim().Equals(nomeCampo.Trim(), StringComparison.OrdinalIgnoreCase) Then
                Return campo
            End If
        Next

        Trace.TraceWarning($"TrovaDefinizioneCampo: campo '{nomeCampo}' non trovato nella definizione dei campi per la tabella {Me.Name}")
        Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show($"Il campo '{nomeCampo}' non è stato trovato nella definizione dei campi.", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Warning)))
        Return Nothing
    End Function


End Class

Partial Public Class VisualMediaForm

    Private pictureBox As PictureBox
    Private mediaPlayer As AxWindowsMediaPlayer

    Public Sub New(percorsoFile As String)
        InitializeComponent()

        Me.Name = "VisualMediaForm_" & Path.GetFileNameWithoutExtension(percorsoFile)
        Me.Text = "Visualizzatore Contenuti"
        Me.Size = New Size(800, 600)

        GestioneStatoForm.CaricaStato(Me)

        pictureBox = New PictureBox With {
            .SizeMode = PictureBoxSizeMode.Zoom,
            .Dock = DockStyle.Fill,
            .Visible = False
        }
        Me.Controls.Add(pictureBox)

        mediaPlayer = New AxWindowsMediaPlayer With {
            .Dock = DockStyle.Fill,
            .Visible = False
        }
        CType(mediaPlayer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Controls.Add(mediaPlayer)
        CType(mediaPlayer, System.ComponentModel.ISupportInitialize).EndInit()

        VisualizzaContenuto(percorsoFile)
    End Sub

    Private Sub VisualizzaContenuto(percorsoOriginale As String)
        Dim estensione = Path.GetExtension(percorsoOriginale).ToLower()
        Dim formatoSupportato = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".mp4", ".avi", ".wmv", ".mov"}

        If Not formatoSupportato.Contains(estensione) Then
            Dim baseName = Path.GetFileNameWithoutExtension(percorsoOriginale)
            Dim directory = Path.GetDirectoryName(percorsoOriginale)
            Dim jpgPath = Path.Combine(directory, baseName & ".jpg")

            If File.Exists(jpgPath) Then
                percorsoOriginale = jpgPath
                estensione = ".jpg"
            Else
                Dim pngPath = Path.Combine(directory, baseName & ".png")
                If File.Exists(pngPath) Then
                    percorsoOriginale = pngPath
                    estensione = ".png"
                Else
                    MDIMessageBox.Show("Formato non riconosciuto e nessun file alternativo trovato: " & baseName, Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Me.Close()
                    Return
                End If
            End If
        End If

        If {".jpg", ".jpeg", ".png", ".bmp", ".gif"}.Contains(estensione) Then
            pictureBox.Image = Image.FromFile(percorsoOriginale)
            pictureBox.Visible = True
            mediaPlayer.Visible = False

        ElseIf {".mp4", ".avi", ".wmv", ".mov"}.Contains(estensione) Then
            mediaPlayer.URL = percorsoOriginale
            mediaPlayer.Visible = True
            pictureBox.Visible = False

        Else
            MDIMessageBox.Show("Tipo di file non supportato: " & estensione, Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
        End If
    End Sub

    Private Sub VisualMediaForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        GestioneStatoForm.SalvaStato(Me)
    End Sub

End Class
