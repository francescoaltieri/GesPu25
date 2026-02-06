Imports System.IO
Imports System.Net.Security
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks.Dataflow
Imports AxWMPLib
Imports Microsoft.Data.SqlClient
Imports PdfSharp
Imports PdfSharp.Drawing
Imports PdfSharp.Drawing.Layout
Imports PdfSharp.Pdf
Imports Excel = Microsoft.Office.Interop.Excel
Public Class DynamicDataForm
    Inherits Form

    Private campoInputs As New Dictionary(Of String, Control)
    Private campiDefiniti As List(Of CampoDatabase)
    Private dgvDati As DataGridView
    Private panelBottoni As FlowLayoutPanel
    Private panelBottoniDinamici As FlowLayoutPanel
    Private modalita As String = ""
    Private isModifica As Boolean
    Private pannelloSinistro As TableLayoutPanel
    Private nomeTabellaCorrente As String
    Private ModalitaCorrente As String = "nessuna"
    Private lblModalita As System.Windows.Forms.Label
    Private lampeggioAttivo As Boolean = False
    Private Shared visualFormsAttivi As New Dictionary(Of String, VisualMediaForm)
    Private splitContainer As SplitContainer
    Private colonneModificate As Boolean = False
    Private isInAvvioForm As Boolean = True

    Private lookupCache As New Dictionary(Of String, DataTable)(StringComparer.OrdinalIgnoreCase)
    Private regexCache As New Dictionary(Of String, Regex)(StringComparer.OrdinalIgnoreCase)

    Private isUpdatingControls As Boolean = False

    Private cachedInsertCommand As SqlCommand = Nothing
    Private cachedUpdateCommand As SqlCommand = Nothing
    Private cachedInsertColumns As List(Of String) = Nothing
    Private cachedUpdateColumns As List(Of String) = Nothing

    Private overlayPanel As Panel = Nothing
    Private overlayLabel As System.Windows.Forms.Label = Nothing
    Private overlaySpinner As ProgressBar = Nothing

    Public Property FiltroIniziale As String

    Private campiPath As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

    Private pageIndex As Integer = 0
    Private pageSize As Integer = 50
    Private totalRows As Integer = 0
    Private totalPages As Integer = 0

    Private pnlPaging As FlowLayoutPanel = Nothing
    Private btnFirst As Button = Nothing
    Private btnPrev As Button = Nothing
    Private btnNext As Button = Nothing
    Private btnLast As Button = Nothing
    Private lblPagingInfo As Label = Nothing
    Private cbPageSize As ComboBox = Nothing

    Private campiCalcolatiDettaglio As Dictionary(Of String, (Formula As String, TipoValore As String, SuSeStesso As Boolean)) = Nothing

    ' Cache per i campi join: chiave = NomeTabella + "|" + NomeCampo
    Private campiJoinCache As Dictionary(Of String, DataRow) = Nothing
    Private campiJoinCacheLoadedForTable As String = String.Empty

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

        Try
            campiPath = RecuperaCampiPath()
        Catch ex As Exception
            MDIMessageBox.Show($"Errore caricamento CampiPath: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)
        End Try

        AddHandler Me.Load, AddressOf DynamicDataForm_Load
        AddHandler Me.ResizeEnd, AddressOf DynamicDataForm_Resize

        GestioneStatoForm.CaricaStato(Me)

        splitContainer = New SplitContainer With {
            .Dock = DockStyle.Fill,
            .Orientation = Orientation.Vertical,
            .FixedPanel = FixedPanel.None
        }
        Me.Controls.Add(splitContainer)

        Dim layoutSinistroInterno As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .RowCount = 3,
            .ColumnCount = 1
        }
        layoutSinistroInterno.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        layoutSinistroInterno.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layoutSinistroInterno.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        pannelloSinistro = New TableLayoutPanel With {
            .Dock = DockStyle.Top,
            .AutoSize = True,
            .AutoScroll = True,
            .ColumnCount = 2,
            .Padding = New Padding(20),
            .GrowStyle = TableLayoutPanelGrowStyle.AddRows,
            .BorderStyle = BorderStyle.Fixed3D
        }
        pannelloSinistro.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        pannelloSinistro.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))

        lblModalita = New System.Windows.Forms.Label With {
            .Text = "",
            .AutoSize = True,
            .Font = New Font("Verdana", 8, FontStyle.Bold),
            .ForeColor = Color.DarkGreen,
            .Dock = DockStyle.Top,
            .Padding = New Padding(5),
            .TextAlign = ContentAlignment.TopLeft,
            .AutoEllipsis = False,
            .UseMnemonic = False
        }
        pannelloSinistro.Controls.Add(lblModalita)
        pannelloSinistro.SetColumnSpan(lblModalita, 2)

        For i = 0 To campi.Count - 1
            If pannelloSinistro.RowCount <= i + 1 Then
                pannelloSinistro.RowCount += 1
                pannelloSinistro.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            End If

            Dim lbl As New System.Windows.Forms.Label With {
                .Text = GetEtichetta(nomeTabella, campi(i).Nome),
                .AutoSize = True,
                .Anchor = AnchorStyles.Left,
                .Margin = New Padding(5)
            }
            Dim ctrl As Control = CreaControllo(campi(i))
            ctrl.Anchor = AnchorStyles.Left
            ctrl.Margin = New Padding(5)
            campoInputs.Add(campi(i).Nome, ctrl)
            If campiCalcolatiDettaglio Is Nothing Then CaricaCacheCampiCalcolati()
            If campiCalcolatiDettaglio.ContainsKey(campi(i).Nome) Then
                Dim dettaglio = campiCalcolatiDettaglio(campi(i).Nome)
                If Not dettaglio.SuSeStesso Then
                    ' disabilita subito il controllo calcolato non autoref
                    If TypeOf ctrl Is FlowLayoutPanel Then
                        For Each ic As Control In CType(ctrl, FlowLayoutPanel).Controls
                            If TypeOf ic Is Button AndAlso String.Equals(CType(ic, Button).Text?.Trim(), "Visualizza", StringComparison.OrdinalIgnoreCase) Then
                                ic.Enabled = True
                            Else
                                ic.Enabled = False
                            End If
                        Next
                    Else
                        ctrl.Enabled = False
                    End If
                End If
            End If
            pannelloSinistro.Controls.Add(lbl, 0, i + 1)
            pannelloSinistro.Controls.Add(ctrl, 1, i + 1)
        Next

        Dim panelBottoniContenitore As New FlowLayoutPanel With {
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.TopDown,
            .AutoSize = True,
            .Padding = New Padding(10),
            .Margin = New Padding(0),
            .WrapContents = False,
            .BorderStyle = BorderStyle.Fixed3D
        }

        panelBottoni = New FlowLayoutPanel With {
            .FlowDirection = FlowDirection.LeftToRight,
            .AutoSize = True,
            .Margin = New Padding(0),
            .Padding = New Padding(0)
        }
        panelBottoniContenitore.Controls.Add(panelBottoni)

        panelBottoniDinamici = New FlowLayoutPanel With {
            .FlowDirection = FlowDirection.LeftToRight,
            .AutoSize = True,
            .Margin = New Padding(0, 8, 0, 0),
            .Padding = New Padding(0),
            .WrapContents = True
        }
        panelBottoniContenitore.Controls.Add(panelBottoniDinamici)

        AggiungiBottone("Inserisci", AddressOf InserisciDati)
        AggiungiBottone("Modifica", AddressOf ModificaDati)
        AggiungiBottone("Salva", AddressOf SalvaDati)
        DisabilitaPulsante("Salva", True)
        AggiungiBottone("Cancella", AddressOf CancellaDati)
        AggiungiBottone("Reset", AddressOf AnnullaOperazione)
        DisabilitaPulsante("Annulla", True)
        AggiungiBottone("Esporta", AddressOf EsportaTabella)
        AggiungiBottone("Rimuovi filtro", Sub()
                                              Dim dt As DataTable = TryCast(dgvDati.DataSource, DataTable)
                                              If dt IsNot Nothing Then dt.DefaultView.RowFilter = ""
                                              lblModalita.Text = "In Attesa..."
                                              lblModalita.ForeColor = Color.DarkGreen
                                          End Sub)

        layoutSinistroInterno.Controls.Add(pannelloSinistro, 0, 0)
        layoutSinistroInterno.Controls.Add(panelBottoniContenitore, 0, 1)
        splitContainer.Panel1.Controls.Add(layoutSinistroInterno)

        dgvDati = New DataGridView With {
            .Dock = DockStyle.Fill,
            .AllowUserToAddRows = False,
            .ReadOnly = True,
            .Name = nomeTabellaCorrente,
            .ScrollBars = System.Windows.Forms.ScrollBars.Both,
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

        InitPagingControls()

        CaricaBottoniDinamici()

        For Each ctrl As Control In campoInputs.Values
            If TypeOf ctrl Is FlowLayoutPanel Then
                Dim hasVisualBtn = ctrl.Controls.OfType(Of System.Windows.Forms.Button)().Any(Function(b) String.Equals(b.Text?.Trim(), "Visualizza", StringComparison.OrdinalIgnoreCase))
                ctrl.Enabled = True

                For Each innerCtrl As Control In ctrl.Controls
                    If TypeOf innerCtrl Is System.Windows.Forms.Button AndAlso String.Equals(CType(innerCtrl, System.Windows.Forms.Button).Text?.Trim(), "Visualizza", StringComparison.OrdinalIgnoreCase) Then
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

        For Each kvp In campoInputs
            If Not regexCache.ContainsKey(kvp.Key) Then
                regexCache(kvp.Key) = New Regex($"\b{Regex.Escape(kvp.Key)}\b", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
            End If
        Next

        InitBusyOverlay()

    End Sub

    Private Sub InitPagingControls()
        pnlPaging = New FlowLayoutPanel With {
            .AutoSize = True,
            .FlowDirection = FlowDirection.LeftToRight,
            .Padding = New Padding(5),
            .Margin = New Padding(5)
        }

        btnFirst = New Button With {.Text = "<<", .AutoSize = True}
        btnPrev = New Button With {.Text = "<", .AutoSize = True}
        btnNext = New Button With {.Text = ">", .AutoSize = True}
        btnLast = New Button With {.Text = ">>", .AutoSize = True}
        lblPagingInfo = New Label With {.AutoSize = True, .Text = "Pagina 0/0", .Padding = New Padding(8, 6, 8, 6)}
        cbPageSize = New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDownList, .Width = 80}

        cbPageSize.Items.AddRange(New Object() {10, 25, 50, 100, 250})
        cbPageSize.SelectedItem = pageSize

        AddHandler btnFirst.Click, Sub(s, e) If pageIndex <> 0 Then pageIndex = 0 : CaricaPaginaAsync()
        AddHandler btnPrev.Click, Sub(s, e) If pageIndex > 0 Then pageIndex -= 1 : CaricaPaginaAsync()
        AddHandler btnNext.Click, Sub(s, e) If pageIndex < totalPages - 1 Then pageIndex += 1 : CaricaPaginaAsync()
        AddHandler btnLast.Click, Sub(s, e) If totalPages > 0 Then pageIndex = totalPages - 1 : CaricaPaginaAsync()
        AddHandler cbPageSize.SelectedIndexChanged, Sub(s, e)
                                                        Dim sel = cbPageSize.SelectedItem
                                                        Dim newSize As Integer = pageSize
                                                        If sel IsNot Nothing AndAlso Integer.TryParse(sel.ToString(), newSize) Then
                                                            pageSize = Math.Max(1, newSize)
                                                            pageIndex = 0
                                                            CaricaPaginaAsync()
                                                        End If
                                                    End Sub

        pnlPaging.Controls.Add(btnFirst)
        pnlPaging.Controls.Add(btnPrev)
        pnlPaging.Controls.Add(lblPagingInfo)
        pnlPaging.Controls.Add(btnNext)
        pnlPaging.Controls.Add(btnLast)
        pnlPaging.Controls.Add(New Label With {.Text = "Righe per pagina:", .AutoSize = True, .Padding = New Padding(8, 6, 0, 0)})
        pnlPaging.Controls.Add(cbPageSize)

        Dim container As New TableLayoutPanel With {.Dock = DockStyle.Bottom, .AutoSize = True, .RowCount = 1, .ColumnCount = 1}
        container.Controls.Add(pnlPaging, 0, 0)
        splitContainer.Panel2.Controls.Add(container)
        pnlPaging.BringToFront()
    End Sub

    Private Async Sub CaricaPaginaAsync()
        Dim filtro = If(String.IsNullOrWhiteSpace(FiltroIniziale), "1=1", FiltroIniziale)
        Dim tableName = Me.Name

        ToggleUIForSaving(True)
        ShowBusyOverlay(True, "Caricamento dati pagina...")

        Try
            Dim dt As New DataTable()
            Dim total As Integer = 0

            Await Task.Run(Sub()
                               Try
                                   Using conn As New SqlConnection(ConnString)
                                       conn.Open()

                                       Using cmdCount As New SqlCommand($"SELECT COUNT(*) FROM [{tableName}] WHERE {filtro}", conn)
                                           total = Convert.ToInt32(cmdCount.ExecuteScalar())
                                       End Using

                                       Dim sz = Math.Max(pageSize, 1)
                                       Dim offset = pageIndex * sz

                                       Dim sql = $"SELECT * FROM [{tableName}] WHERE {filtro} ORDER BY (SELECT NULL) OFFSET {offset} ROWS FETCH NEXT {sz} ROWS ONLY"

                                       Using cmd As New SqlCommand(sql, conn)
                                           Using da As New SqlDataAdapter(cmd)
                                               da.Fill(dt)
                                           End Using
                                       End Using
                                   End Using
                               Catch ex As Exception
                                   Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show($"CaricaPaginaAsync errore: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)))
                               End Try
                           End Sub)

            Me.BeginInvoke(New MethodInvoker(Sub()
                                                 totalRows = total
                                                 totalPages = If(pageSize > 0, CInt(Math.Ceiling(totalRows / CSng(pageSize))), 0)
                                                 If pageIndex >= totalPages AndAlso totalPages > 0 Then pageIndex = totalPages - 1

                                                 dgvDati.DataSource = dt
                                                 UpdatePagingUI()
                                             End Sub))
        Catch ex As Exception
            Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show("Errore caricamento dati: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)))
        Finally
            ShowBusyOverlay(False)
            ToggleUIForSaving(False)
        End Try
    End Sub

    Private Sub UpdatePagingUI()
        lblPagingInfo.Text = $"Pagina {Math.Max(pageIndex + 1, 0)}/{Math.Max(totalPages, 0)}  (righe: {totalRows})"

        btnFirst.Enabled = (pageIndex > 0)
        btnPrev.Enabled = (pageIndex > 0)
        btnNext.Enabled = (pageIndex < totalPages - 1)
        btnLast.Enabled = (pageIndex < totalPages - 1)

        If totalPages <= 1 Then
            btnFirst.Enabled = False
            btnPrev.Enabled = False
            btnNext.Enabled = False
            btnLast.Enabled = False
        End If
    End Sub

    Private Sub InitBusyOverlay()
        overlayPanel = New Panel With {
        .Dock = DockStyle.Fill,
        .BackColor = Color.FromArgb(120, Color.LightGray),
        .Visible = False
    }

        overlayLabel = New System.Windows.Forms.Label With {
        .AutoSize = False,
        .TextAlign = ContentAlignment.MiddleCenter,
        .Dock = DockStyle.Top,
        .Height = 40,
        .Font = New Font("Segoe UI", 10, FontStyle.Bold),
        .ForeColor = Color.Black,
        .Text = "Salvataggio in corso..."
    }

        overlaySpinner = New ProgressBar With {
        .Style = ProgressBarStyle.Marquee,
        .MarqueeAnimationSpeed = 30,
        .Height = 18,
        .Dock = DockStyle.Top
    }

        Dim inner As New TableLayoutPanel With {
        .Dock = DockStyle.None,
        .AutoSize = True,
        .BackColor = Color.Transparent,
        .ColumnCount = 1,
        .RowCount = 2
    }
        inner.Controls.Add(overlayLabel, 0, 0)
        inner.Controls.Add(overlaySpinner, 0, 1)
        inner.Padding = New Padding(10)

        inner.Location = New Point((Me.ClientSize.Width - inner.PreferredSize.Width) \ 2, (Me.ClientSize.Height - inner.PreferredSize.Height) \ 2)
        overlayPanel.Controls.Add(inner)

        Me.Controls.Add(overlayPanel)
        overlayPanel.BringToFront()

        AddHandler Me.Resize, Sub()
                                  If overlayPanel IsNot Nothing AndAlso overlayPanel.Visible Then
                                      inner.Location = New Point((Me.ClientSize.Width - inner.PreferredSize.Width) \ 2, (Me.ClientSize.Height - inner.PreferredSize.Height) \ 2)
                                  End If
                              End Sub
    End Sub

    Private Sub ShowBusyOverlay(onOff As Boolean, Optional message As String = "Salvataggio in corso...")
        If overlayPanel Is Nothing Then Return

        Me.BeginInvoke(New MethodInvoker(Sub()
                                             overlayLabel.Text = message
                                             overlayPanel.Visible = onOff
                                             overlayPanel.BringToFront()
                                             Me.Cursor = If(onOff, Cursors.WaitCursor, Cursors.Default)
                                         End Sub))
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
            Dim query = "SELECT * FROM Sys_CampiConvalida WHERE NomeTabella = @Tabella"
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
                                                 AllineaColonne(dgvDati)

                                             End Sub))
        Else

            Debug.WriteLine("Le colonne della griglia non sono ancora disponibili.")
        End If

        For Each col As DataGridViewColumn In dgvDati.Columns
            Try
                col.HeaderText = GetEtichetta(Me.Name, col.Name)
            Catch ex As Exception
                MDIMessageBox.Show($"Impossibile impostare header per {col.Name}: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)
            End Try
        Next

    End Sub

    Private Sub AllineaColonne(dgv As DataGridView)
        If dgv Is Nothing OrElse dgv.DataSource Is Nothing Then Return

        Dim dt As DataTable = TryCast(dgv.DataSource, DataTable)
        For Each col As DataGridViewColumn In dgv.Columns
            Try
                Dim isNumericColumn As Boolean = False
                Dim isBooleanColumn As Boolean = False

                If dt IsNot Nothing AndAlso dt.Columns.Contains(col.Name) Then
                    Dim dataType As Type = dt.Columns(col.Name).DataType
                    If dataType Is GetType(Boolean) Then
                        isBooleanColumn = True
                    ElseIf dataType Is GetType(Integer) OrElse dataType Is GetType(Long) _
                   OrElse dataType Is GetType(Short) OrElse dataType Is GetType(Decimal) _
                   OrElse dataType Is GetType(Double) OrElse dataType Is GetType(Single) _
                   OrElse dataType Is GetType(Byte) Then

                        isNumericColumn = True
                    End If
                Else
                    Dim nome = col.Name.ToLowerInvariant()
                    If nome.EndsWith("id") OrElse nome.Contains("quant") OrElse nome.Contains("prezzo") OrElse nome.Contains("importo") Then
                        isNumericColumn = True
                    End If
                    If nome = "isactive" OrElse nome.StartsWith("is") OrElse nome.Contains("flag") OrElse nome.Contains("can") OrElse nome.Contains("abil") Then
                        ' heuristica per colonne boolean che non hanno tipo disponibile nel DataTable
                        isBooleanColumn = True
                    End If
                End If

                If isBooleanColumn Then
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                ElseIf isNumericColumn Then
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                Else
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                End If

                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            Catch ex As Exception
                MDIMessageBox.Show($"AllineaColonneNumeriche: errore su colonna {col.Name}: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)
            End Try
        Next
    End Sub

    Private Sub NascondiColonneSensibili()
        For Each col As DataGridViewColumn In dgvDati.Columns
            If col.Name.ToLower().Contains("password") Then
                col.Visible = False
            End If
        Next
    End Sub

    Private Sub DynamicDataForm_Load(sender As Object, e As EventArgs)

        campiCalcolatiSet = RecuperaCampiCalcolatiSet(Me.Name)

        If splitContainer IsNot Nothing Then
            splitContainer.Panel1MinSize = 300
            splitContainer.Panel2MinSize = 300
            splitContainer.SplitterDistance = Me.Width / 2.5
        End If

        pageIndex = 0
        CaricaPaginaAsync()
        AggiornaMaxWidthModalita()

        Me.BeginInvoke(New MethodInvoker(Sub()
                                             isInAvvioForm = False
                                             Me.Refresh()
                                             UpdateButtonsByModalita()
                                         End Sub))

    End Sub

    Private Sub DynamicDataForm_Resize(sender As Object, e As EventArgs)
        AggiornaMaxWidthModalita()
    End Sub

    Private Sub AggiornaMaxWidthModalita()
        Try
            Dim padding As Integer = 40
            Dim maxW = Math.Max(200, pannelloSinistro.ClientSize.Width - padding)
            lblModalita.MaximumSize = New Size(maxW, 0)
        Catch
        End Try
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
                MDIMessageBox.Show($"Errore salvando configurazione griglia alla chiusura: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)

            End Try
        End If
    End Sub

    Private Sub FormDinamico_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        PosizionaGrigliaDaSysForm()
    End Sub

    Private Sub AnnullaOperazione()
        Dim risposta = MDIMessageBox.Show("Vuoi eseguire reset del Form?", Me.MdiParent, MessageBoxButtons.YesNo, MessageBoxIcon.Question)

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

            ResetForm()

        End If
    End Sub

    Private Sub AbilitaCampi(abilita As Boolean)
        ' Assicura cache caricata
        If campiCalcolatiDettaglio Is Nothing Then CaricaCacheCampiCalcolati()

        For Each kvp In campoInputs
            Dim nomeCampo As String = kvp.Key
            Dim ctrl As Control = kvp.Value

            Dim campo As CampoDatabase = campiDefiniti.FirstOrDefault(Function(c) c.Nome = nomeCampo)
            If campo Is Nothing Then Continue For

            ' Campo calcolato = solo se presente in Sys_CampiCalcolati
            Dim isCalcolato As Boolean = campiCalcolatiDettaglio IsNot Nothing AndAlso campiCalcolatiDettaglio.ContainsKey(nomeCampo)
            Dim suSeStesso As Boolean = False
            If isCalcolato Then suSeStesso = campiCalcolatiDettaglio(nomeCampo).SuSeStesso

            ' Regola: se è calcolato e NON SuSeStesso => disabilita il controllo
            If isCalcolato AndAlso Not suSeStesso Then
                ' se è un pannello con bottoni Visualizza, mantieni il bottone abilitato
                If TypeOf ctrl Is FlowLayoutPanel Then
                    Dim flow = CType(ctrl, FlowLayoutPanel)
                    For Each innerCtrl As Control In flow.Controls
                        If TypeOf innerCtrl Is Button AndAlso String.Equals(CType(innerCtrl, Button).Text?.Trim(), "Visualizza", StringComparison.OrdinalIgnoreCase) Then
                            innerCtrl.Enabled = True
                        Else
                            innerCtrl.Enabled = False
                        End If
                    Next
                Else
                    ctrl.Enabled = False
                End If
                Continue For
            End If

            ' Altrimenti applica le regole standard (identity, chiave, join)
            Dim joinRow = RecuperaJoinPerCampoCached(nomeTabellaCorrente, nomeCampo)
            Dim isJoin = (joinRow IsNot Nothing)
            Dim joinModificabile As Boolean = True
            If isJoin AndAlso joinRow.Table.Columns.Contains("AbilitaModifica") Then
                joinModificabile = Convert.ToBoolean(joinRow("AbilitaModifica"))
            End If

            Dim isBloccato As Boolean = campo.IsIdentity OrElse (campo.IsChiave And ModalitaCorrente <> "inserimento") OrElse (isJoin AndAlso Not joinModificabile)

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

    Private Sub FocusSulPrimoCampoEditabile()
        Me.BeginInvoke(New MethodInvoker(Sub()
                                             Try
                                                 Dim primo As Control = Nothing

                                                 For i = 0 To pannelloSinistro.RowCount - 1
                                                     For Each c As Control In pannelloSinistro.GetControlFromPosition(1, i)?.Controls
                                                         ' ignora se nulla
                                                     Next
                                                 Next

                                                 For Each ctrl As Control In pannelloSinistro.Controls
                                                     If ctrl Is lblModalita Then Continue For
                                                     If Not ctrl.Enabled Then Continue For
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
                                                         If (TypeOf ctrl Is TextBox OrElse TypeOf ctrl Is ComboBox OrElse TypeOf ctrl Is DateTimePicker OrElse TypeOf ctrl Is CheckBox) AndAlso ctrl.Enabled AndAlso ctrl.Visible Then
                                                             primo = ctrl
                                                             Exit For
                                                         End If
                                                     End If
                                                 Next

                                                 If primo IsNot Nothing Then
                                                     primo.Focus()
                                                     If TypeOf primo Is TextBox Then
                                                         CType(primo, TextBox).SelectAll()
                                                     ElseIf TypeOf primo Is ComboBox Then
                                                         CType(primo, ComboBox).DroppedDown = False
                                                     End If
                                                 End If
                                             Catch ex As Exception
                                                 MDIMessageBox.Show($"FocusSulPrimoCampoEditabile errore: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)
                                             End Try
                                         End Sub))
    End Sub

    Private Sub InserisciDati(sender As Object, e As EventArgs)

        isModifica = False
        ModalitaCorrente = "inserimento"
        lblModalita.Text = "Inserimento in corso..."

        PulisciCampi()
        AbilitaCampi(True)
        ResetLabelDescrizioni()
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

            'ctrl.Enabled = Not campo.IsIdentity
        Next

        FocusSulPrimoCampoEditabile()

    End Sub

    Private Function RecuperaJoinPerCampoCached(nomeTabella As String, nomeCampo As String) As DataRow
        If campiJoinCache Is Nothing OrElse campiJoinCacheLoadedForTable <> nomeTabella Then
            CaricaCampiJoinCachePerTabella(nomeTabella)
        End If

        Dim key = $"{nomeTabella}|{nomeCampo}"
        If campiJoinCache IsNot Nothing AndAlso campiJoinCache.ContainsKey(key) Then
            Return campiJoinCache(key)
        End If
        Return Nothing
    End Function

    Private Sub ModificaDati(sender As Object, e As EventArgs)

        If dgvDati.SelectedRows.Count = 0 Then
            MDIMessageBox.Show("Seleziona prima una riga dalla griglia.", Me.MdiParent, MessageBoxButtons.OK)
            Return
        End If

        ' Carica la cache dei campi join per questa tabella una sola volta
        CaricaCampiJoinCachePerTabella(Me.Name)

        DisabilitaPulsante("Salva", False)
        ModalitaCorrente = "modifica"
        lblModalita.Text = "Modifica in corso..."
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

            Dim joinRow = RecuperaJoinPerCampoCached(Me.Name, campo.Nome)
            If joinRow IsNot Nothing Then
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
                If ctrl Is Nothing Then Continue For

                If TypeOf ctrl Is TextBox Then
                    CType(ctrl, TextBox).Text = If(valoreJoin Is Nothing, String.Empty, valoreJoin.ToString())
                ElseIf TypeOf ctrl Is ComboBox Then
                    Try
                        CType(ctrl, ComboBox).SelectedValue = valoreJoin
                    Catch
                        CType(ctrl, ComboBox).SelectedIndex = -1
                    End Try

                ElseIf TypeOf ctrl Is FlowLayoutPanel Then
                    Dim flow = CType(ctrl, FlowLayoutPanel)

                    ' imposta la TextBox interna se presente
                    Dim innerTxt = flow.Controls.OfType(Of TextBox)().FirstOrDefault()
                    If innerTxt IsNot Nothing Then
                        innerTxt.Text = If(valoreJoin Is Nothing, String.Empty, valoreJoin.ToString())
                    End If

                    ' aggiorna la Label descrizione interna (se presente e se campo ha TabellaElenco/CampoVisuale)
                    Dim lbl = flow.Controls.OfType(Of Label)().FirstOrDefault()
                    Dim campoDef As CampoDatabase = TrovaDefinizioneCampo(campo.Nome)
                    If lbl IsNot Nothing AndAlso campoDef IsNot Nothing _
                    AndAlso Not String.IsNullOrEmpty(campoDef.TabellaElenco) _
                    AndAlso Not String.IsNullOrEmpty(campoDef.ChiaveElenco) _
                    AndAlso Not String.IsNullOrEmpty(campoDef.CampoVisuale) Then

                        Try
                            Dim dtRef As DataTable = RecuperaTabellaCached(campoDef.TabellaElenco)
                            Dim codice = If(valoreJoin Is Nothing, String.Empty, valoreJoin.ToString().Replace("'", "''"))
                            If dtRef IsNot Nothing AndAlso dtRef.Columns.Contains(campoDef.ChiaveElenco) AndAlso dtRef.Columns.Contains(campoDef.CampoVisuale) Then
                                Dim rows = dtRef.Select($"{campoDef.ChiaveElenco} = '{codice}'")
                                If rows.Length > 0 Then
                                    lbl.Text = rows(0)(campoDef.CampoVisuale).ToString()
                                Else
                                    lbl.Text = "..."
                                End If
                            Else
                                lbl.Text = "..."
                            End If
                        Catch ex As Exception
                            lbl.Text = "..."
                        End Try
                    End If
                Else
                    ' fallback generico
                    Try
                        ctrl.Text = If(valoreJoin Is Nothing, String.Empty, valoreJoin.ToString())
                    Catch
                    End Try
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
            ShowBusyOverlay(True, "Salvataggio in corso...")

            If isModifica Then
                Await SalvaModificaAsync()
                CaricaDatiNeiControlli(dgvDati.SelectedRows(0))
            Else
                Await SalvaInserimentoAsync()
                ResetLabelDescrizioni()
            End If

            Dim dt As DataTable = Nothing
            Try
                dt = Await Task.Run(Function()
                                        Dim tmp As New DataTable()
                                        Try
                                            Using conn As New SqlConnection(ConnString)
                                                Dim query = $"SELECT * FROM [{Me.Name}]" & If(String.IsNullOrWhiteSpace(FiltroIniziale), "", $" WHERE {FiltroIniziale}")
                                                Using cmd As New SqlCommand(query, conn)
                                                    cmd.CommandTimeout = 60
                                                    Using da As New SqlDataAdapter(cmd)
                                                        da.Fill(tmp)
                                                    End Using
                                                End Using
                                            End Using
                                        Catch ex As Exception
                                            MDIMessageBox.Show($"Errore ricarica dati background: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)
                                        End Try
                                        Return tmp
                                    End Function)
            Catch ex As Exception
                MDIMessageBox.Show($"Task.Run ricarica error: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)
            End Try

            Me.BeginInvoke(New MethodInvoker(Sub()
                                                 Try
                                                     If dt IsNot Nothing AndAlso dt.Columns.Count > 0 Then
                                                         dgvDati.DataSource = dt
                                                     Else
                                                         CaricaDatiTabellaAsync(Me.Name)
                                                     End If

                                                     DisabilitaCampi()
                                                     DisabilitaPulsante("Salva", True)
                                                     lblModalita.ForeColor = Color.DarkGreen
                                                     ModalitaCorrente = "nessuna"
                                                     lblModalita.Text = "Scheda Salvata"
                                                     DisabilitaPulsante("Annulla", True)
                                                     UpdateButtonsByModalita()
                                                 Catch ex As Exception
                                                     MDIMessageBox.Show($"Errore aggiornamento UI dopo save: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)
                                                 End Try
                                             End Sub))

        Catch ex As Exception
            Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show("Errore durante il salvataggio: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)))
        Finally
            ShowBusyOverlay(False)
            ToggleUIForSaving(False)
            sw.Stop()
        End Try
    End Sub

    Private Sub CancellaDati(sender As Object, e As EventArgs)
        If dgvDati.SelectedRows.Count = 0 Then
            MDIMessageBox.Show("Seleziona prima una riga da cancellare dalla griglia.", Me.MdiParent, MessageBoxButtons.OK)
            Return
        End If

        'campiDefiniti = RecuperaCampiDa(Me.Name)

        Dim campoChiave = campiDefiniti.FirstOrDefault(Function(c) c.IsChiave)
        If campoChiave Is Nothing Then
            MDIMessageBox.Show("Nessuna chiave primaria definita.", Me.MdiParent, MessageBoxButtons.OK)
            ResetForm()
            Return
        End If

        Dim cella = dgvDati.SelectedRows(0).Cells(campoChiave.Nome)
        If cella.Value Is Nothing Then
            MDIMessageBox.Show("Il valore della chiave è nullo.", Me.MdiParent, MessageBoxButtons.OK)
            ResetForm()
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

                If pageIndex > 0 AndAlso (totalRows - 1) <= pageIndex * pageSize Then
                    pageIndex = Math.Max(0, pageIndex - 1)
                End If

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
        ResetForm()
    End Sub

    Private Sub CaricaCampiJoinCachePerTabella(nomeTabella As String)
        If campiJoinCache IsNot Nothing AndAlso campiJoinCacheLoadedForTable = nomeTabella Then
            Return
        End If

        Dim dict As New Dictionary(Of String, DataRow)(StringComparer.OrdinalIgnoreCase)

        Using conn As New SqlConnection(ConnString)
            Dim query = "SELECT * FROM sys_CampiJoin WHERE NomeTabella = @Tabella"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@Tabella", nomeTabella)
                Dim da As New SqlDataAdapter(cmd)
                Dim dt As New DataTable()
                da.Fill(dt)
                For Each r As DataRow In dt.Rows
                    Dim nomeCampo = r("NomeCampo").ToString()
                    Dim key = $"{nomeTabella}|{nomeCampo}"
                    If Not dict.ContainsKey(key) Then
                        dict.Add(key, r)
                    End If
                Next
            End Using
        End Using

        campiJoinCache = dict
        campiJoinCacheLoadedForTable = nomeTabella
    End Sub

    Private Function EstraiValoreDaControllo(ctrl As Control) As Object
        If ctrl Is Nothing Then Return DBNull.Value

        Try
            If TypeOf ctrl Is TextBox Then
                Dim s = CType(ctrl, TextBox).Text
                Return If(String.IsNullOrEmpty(s), DBNull.Value, CType(s, Object))
            ElseIf TypeOf ctrl Is ComboBox Then
                Dim cb = CType(ctrl, ComboBox)
                Dim sv = cb.SelectedValue
                If sv IsNot Nothing AndAlso Not Convert.IsDBNull(sv) Then Return sv
                Dim txt = cb.Text
                Return If(String.IsNullOrEmpty(txt), DBNull.Value, CType(txt, Object))
            ElseIf TypeOf ctrl Is FlowLayoutPanel Then
                Dim innerTxt = ctrl.Controls.OfType(Of TextBox)().FirstOrDefault()
                If innerTxt IsNot Nothing Then
                    Dim s = innerTxt.Text
                    Return If(String.IsNullOrEmpty(s), DBNull.Value, CType(s, Object))
                End If
            End If

            Dim fallback = ctrl.Text
            Return If(String.IsNullOrEmpty(fallback), DBNull.Value, CType(fallback, Object))
        Catch ex As Exception
            Return DBNull.Value
        End Try
    End Function

    Private Sub AggiornaControlloConValoreJoin(ctrl As Control, valoreJoin As Object)
        If ctrl Is Nothing Then Return

        Dim action As Action = Sub()
                                   Try
                                       If TypeOf ctrl Is TextBox Then
                                           CType(ctrl, TextBox).Text = If(valoreJoin Is Nothing OrElse Convert.IsDBNull(valoreJoin), String.Empty, valoreJoin.ToString())
                                       ElseIf TypeOf ctrl Is ComboBox Then
                                           Dim cb = CType(ctrl, ComboBox)
                                           Try
                                               cb.SelectedValue = If(valoreJoin Is Nothing OrElse Convert.IsDBNull(valoreJoin), Nothing, valoreJoin)
                                           Catch
                                               cb.SelectedIndex = -1
                                           End Try
                                       ElseIf TypeOf ctrl Is FlowLayoutPanel Then
                                           Dim flow = CType(ctrl, FlowLayoutPanel)
                                           Dim innerTxt = flow.Controls.OfType(Of TextBox)().FirstOrDefault()
                                           If innerTxt IsNot Nothing Then innerTxt.Text = If(valoreJoin Is Nothing OrElse Convert.IsDBNull(valoreJoin), String.Empty, valoreJoin.ToString())
                                       Else
                                           Try
                                               ctrl.Text = If(valoreJoin Is Nothing OrElse Convert.IsDBNull(valoreJoin), String.Empty, valoreJoin.ToString())
                                           Catch
                                           End Try
                                       End If

                                       ' Forza la scrittura dei binding (se presenti)
                                       If ctrl.DataBindings IsNot Nothing AndAlso ctrl.DataBindings.Count > 0 Then
                                           For Each b As Binding In ctrl.DataBindings
                                               Try
                                                   b.WriteValue()
                                               Catch
                                               End Try
                                           Next
                                       End If
                                   Catch
                                   End Try
                               End Sub

        If ctrl.InvokeRequired Then
            ctrl.Invoke(New MethodInvoker(Sub() action()))
        Else
            action()
        End If
    End Sub

    Private Sub RicalcolaCampiJoinPrimaSalvataggio()
        ' Carica la cache dei campi join per questa tabella una sola volta
        CaricaCampiJoinCachePerTabella(Me.Name)

        For Each campo In campiDefiniti
            If Not campoInputs.ContainsKey(campo.Nome) Then Continue For

            Dim joinRow = RecuperaJoinPerCampoCached(Me.Name, campo.Nome)
            If joinRow Is Nothing Then Continue For

            ' Costruisco dizionario chiavi figlia (se presenti nella definizione join)
            Dim chiaviFiglia As New Dictionary(Of String, Object)
            For i = 1 To 3
                Dim chiaveNome = $"ChiaveFiglia{i}"
                If joinRow.Table.Columns.Contains(chiaveNome) Then
                    Dim nomeCampoFiglio = joinRow(chiaveNome).ToString()
                    If Not String.IsNullOrWhiteSpace(nomeCampoFiglio) Then
                        Dim valore As Object = Nothing
                        ' Se esiste un controllo per il campo figlio, prendi il valore dal controllo (valore modificato)
                        If campoInputs.ContainsKey(nomeCampoFiglio) Then
                            Dim ctrlFiglio = campoInputs(nomeCampoFiglio)
                            valore = EstraiValoreDaControllo(ctrlFiglio)
                        Else
                            ' Altrimenti prendi dalla riga selezionata nella griglia (fallback)
                            valore = If(dgvDati.SelectedRows.Count > 0, dgvDati.SelectedRows(0).Cells(nomeCampoFiglio).Value, Nothing)
                        End If
                        chiaviFiglia.Add(chiaveNome, valore)
                    End If
                End If
            Next

            Dim valoreJoin = PrelevaValoreJoin(joinRow, chiaviFiglia)

            ' Aggiorna il controllo target
            Dim ctrl = campoInputs(campo.Nome)
            If ctrl Is Nothing Then Continue For

            AggiornaControlloConValoreJoin(ctrl, valoreJoin)
        Next
    End Sub

    Private Sub ResetForm()
        DisabilitaCampi()
        pageIndex = 0
        CaricaPaginaAsync()
        DisabilitaPulsante("Salva", True)
        lampeggioAttivo = False
        lblModalita.ForeColor = Color.DarkGreen
        DisabilitaPulsante("Annulla", True)
        ModalitaCorrente = "nessuna"
        lblModalita.Text = "In Attesa..."
        PulisciCampi()
        ResetLabelDescrizioni()
        UpdateButtonsByModalita()
    End Sub

    Private Sub ResetLabelDescrizioni()
        For Each campo In campiDefiniti
            If Not campoInputs.ContainsKey(campo.Nome) Then Continue For
            Dim ctrl = campoInputs(campo.Nome)

            If TypeOf ctrl Is FlowLayoutPanel Then
                Dim lbl As Label = ctrl.Controls.OfType(Of Label)().FirstOrDefault()
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

    Private Function BuildPreparedInsertCommand(conn As SqlConnection, tx As SqlTransaction, tableName As String, columns As List(Of String)) As SqlCommand
        Dim query As String = $"INSERT INTO [{tableName}] ({String.Join(",", columns)}) VALUES ({String.Join(",", columns.Select(Function(n) "@" & n))})"
        Dim cmd As New SqlCommand(query, conn, tx)
        cmd.CommandTimeout = 120

        For Each nomeCampo In columns
            Dim campoDef = campiDefiniti.FirstOrDefault(Function(c) String.Equals(c.Nome, nomeCampo, StringComparison.OrdinalIgnoreCase))
            Dim sqlType = If(campoDef IsNot Nothing, GetSqlDbTypePerCampo(campoDef), SqlDbType.NVarChar)
            Dim size As Integer = 0
            If campoDef IsNot Nothing AndAlso sqlType = SqlDbType.NVarChar Then
                Dim l As Integer = 0
                If Integer.TryParse(Convert.ToString(campoDef.Lunghezza), l) Then size = Math.Min(Math.Max(l, 0), 4000)
            End If

            Dim param As SqlParameter
            If size > 0 Then
                param = cmd.Parameters.Add("@" & nomeCampo, sqlType, size)
            Else
                param = cmd.Parameters.Add("@" & nomeCampo, sqlType)
            End If
            param.Value = DBNull.Value
        Next

        Try
            cmd.Prepare()
        Catch

        End Try

        Return cmd
    End Function

    Private Function BuildPreparedUpdateCommand(conn As SqlConnection, tx As SqlTransaction, tableName As String, columns As List(Of String), keyCampo As CampoDatabase) As SqlCommand
        Dim setPart = String.Join(",", columns.Select(Function(n) $"{n} = @{n}"))
        Dim query As String = $"UPDATE [{tableName}] SET {setPart} WHERE {keyCampo.Nome} = @{keyCampo.Nome}"
        Dim cmd As New SqlCommand(query, conn, tx)
        cmd.CommandTimeout = 120

        For Each nomeCampo In columns
            Dim campoDef = campiDefiniti.FirstOrDefault(Function(c) String.Equals(c.Nome, nomeCampo, StringComparison.OrdinalIgnoreCase))
            Dim sqlType = If(campoDef IsNot Nothing, GetSqlDbTypePerCampo(campoDef), SqlDbType.NVarChar)
            Dim size As Integer = 0
            If campoDef IsNot Nothing AndAlso sqlType = SqlDbType.NVarChar Then
                Dim l As Integer = 0
                If Integer.TryParse(Convert.ToString(campoDef.Lunghezza), l) Then size = Math.Min(Math.Max(l, 0), 4000)
            End If

            If size > 0 Then
                cmd.Parameters.Add("@" & nomeCampo, sqlType, size).Value = DBNull.Value
            Else
                cmd.Parameters.Add("@" & nomeCampo, sqlType).Value = DBNull.Value
            End If
        Next

        Dim keySqlType = GetSqlDbTypePerCampo(keyCampo)
        cmd.Parameters.Add("@" & keyCampo.Nome, keySqlType).Value = DBNull.Value

        Try
            cmd.Prepare()
        Catch
        End Try

        Return cmd
    End Function

    Private Async Function SalvaInserimentoAsync() As Task
        Dim swTotal As New Stopwatch()
        swTotal.Start()

        ' Assicurati di avere i campi definiti
        campiDefiniti = RecuperaCampiDa(Me.Name)

        ' Ricalcola i campi join (aggiorna i controlli se necessario)
        RicalcolaCampiJoinPrimaSalvataggio()

        Dim campiCalcolati = RecuperaCampiCalcolati()
        Dim formule = campiCalcolati.ToDictionary(Function(kvp) kvp.Key, Function(kvp) kvp.Value.Formula)
        Dim tipiValore = campiCalcolati.ToDictionary(Function(kvp) kvp.Key, Function(kvp) kvp.Value.TipoValore)
        Dim valoriCalcolati = CalcolaValoriCampiCalcolati(formule, tipiValore)

        ' Costruisci colonne come in SalvaModifica (escludi identity)
        Dim colonne As New List(Of String)
        For Each c In campiDefiniti
            If Not c.IsIdentity Then
                colonne.Add(c.Nome)
            End If
        Next
        If colonne.Count = 0 Then Return

        ' Determina colonne valide (escludi password vuote come in SalvaModifica)
        Dim colonneValid As New List(Of String)
        For Each campo In campiDefiniti
            If campo.IsIdentity Then Continue For
            Dim input = If(campoInputs.ContainsKey(campo.Nome), campoInputs(campo.Nome), Nothing)
            Dim isPassword = campo.Nome.ToLower().Contains("password")
            If isPassword AndAlso TypeOf input Is TextBox AndAlso String.IsNullOrWhiteSpace(CType(input, TextBox).Text) Then
                Continue For
            End If
            colonneValid.Add(campo.Nome)
        Next
        If colonneValid.Count = 0 Then Return

        ' Prepara dizionario valoriInput 
        Dim valoriInput As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
        For Each nomeCampo In colonneValid
            Dim valore As Object = Nothing
            If valoriCalcolati.ContainsKey(nomeCampo) Then
                valore = valoriCalcolati(nomeCampo)
            Else
                Try
                    If campoInputs.ContainsKey(nomeCampo) Then
                        valore = EstraiValoreDaControllo(campoInputs(nomeCampo))
                    Else
                        valore = DBNull.Value
                    End If
                Catch
                    valore = DBNull.Value
                End Try
            End If
            If valore Is Nothing Then valore = DBNull.Value
            valoriInput(nomeCampo) = valore
        Next

        Try
            Dim convalide = RecuperaConvalideDaSys(Me.Name)
            If convalide IsNot Nothing AndAlso convalide.Count > 0 Then
                For Each campo In campiDefiniti
                    If campo Is Nothing OrElse String.IsNullOrWhiteSpace(campo.Nome) Then Continue For
                    If convalide.ContainsKey(campo.Nome) Then
                        Try
                            ApplicaConvalidaAlCampo(campo, convalide(campo.Nome))
                        Catch ex As Exception
                            MDIMessageBox.Show($"RiapplicaConvalide: errore su campo {campo.Nome}: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)
                        End Try
                    End If
                Next
            End If
        Catch ex As Exception
            MDIMessageBox.Show($"RiapplicaConvalide: errore generale: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)
        End Try

        Dim errorList As New List(Of String)

        ' Cache join è caricata?
        CaricaCampiJoinCachePerTabella(Me.Name)

        Dim joinDefs As New Dictionary(Of String, DataRow)(StringComparer.OrdinalIgnoreCase)
        Dim joinLookupCache As New Dictionary(Of String, DataTable)(StringComparer.OrdinalIgnoreCase)
        Dim joinResultCache As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)

        For Each nomeCampo In colonneValid
            Try
                Dim jr = RecuperaJoinPerCampoCached(Me.Name, nomeCampo)
                If jr IsNot Nothing Then
                    joinDefs(nomeCampo) = jr
                    If jr.Table.Columns.Contains("TabellaElenco") Then
                        Dim tabName = Convert.ToString(jr("TabellaElenco"))
                        If Not String.IsNullOrWhiteSpace(tabName) AndAlso Not joinLookupCache.ContainsKey(tabName) Then
                            Dim dtRef = RecuperaTabellaCached(tabName)
                            If dtRef IsNot Nothing Then joinLookupCache(tabName) = dtRef
                        End If
                    End If
                End If
            Catch ex As Exception
                ' non bloccare l'inserimento per errori di lookup join
            End Try
        Next

        ' Risolvi i join e sovrascrivi i valori in valoriInput quando disponibili
        For Each nomeCampo In colonneValid
            Try
                If Not joinDefs.ContainsKey(nomeCampo) Then Continue For
                Dim jr = joinDefs(nomeCampo)
                Dim chiaviFiglia As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)

                ' Costruisce le chiavi figlie 
                Dim anyKeyAdded As Boolean = False
                For i As Integer = 1 To 3
                    Dim chiaveNome = $"ChiaveFiglia{i}"
                    If jr.Table.Columns.Contains(chiaveNome) Then
                        Dim nomeCampoFiglio = Convert.ToString(jr(chiaveNome))
                        If Not String.IsNullOrWhiteSpace(nomeCampoFiglio) Then
                            Dim valFiglio As Object = DBNull.Value
                            If campoInputs.ContainsKey(nomeCampoFiglio) Then
                                valFiglio = EstraiValoreDaControllo(campoInputs(nomeCampoFiglio))
                            ElseIf valoriInput.ContainsKey(nomeCampoFiglio) Then
                                valFiglio = valoriInput(nomeCampoFiglio)
                            ElseIf dgvDati.SelectedRows.Count > 0 AndAlso dgvDati.Columns.Contains(nomeCampoFiglio) Then
                                valFiglio = dgvDati.SelectedRows(0).Cells(nomeCampoFiglio).Value
                            End If

                            ' filtro coerente
                            chiaviFiglia.Add(chiaveNome, If(valFiglio Is Nothing, DBNull.Value, valFiglio))
                            anyKeyAdded = True
                        End If
                    End If
                Next

                If Not anyKeyAdded Then
                    Continue For
                End If

                ' Prova la lookup cached con fallback al DB
                Dim valoreJoin As Object = Nothing
                Try
                    valoreJoin = PrelevaValoreJoinCached(jr, chiaviFiglia, joinLookupCache, joinResultCache)
                Catch ex As Exception
                    Try
                        valoreJoin = PrelevaValoreJoin(jr, chiaviFiglia)
                    Catch ex2 As Exception
                        valoreJoin = Nothing
                    End Try
                End Try

                ' Assegna anche se valoreJoin è DBNull.Value 
                If valoreJoin IsNot Nothing OrElse (valoreJoin Is DBNull.Value) Then
                    valoriInput(nomeCampo) = If(valoreJoin Is Nothing, DBNull.Value, valoreJoin)
                End If

            Catch ex As Exception
                ' ignore per non bloccare inserimento
            End Try
        Next

        ' --- Rileva campo con pattern di intervallo del tipo <da>0004<a>0015
        Dim intervalRegex As New System.Text.RegularExpressions.Regex("<da>(\d+)<a>(\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        Dim variableFieldName As String = Nothing
        Dim variableOriginalValue As String = Nothing
        Dim generatedValues As List(Of String) = Nothing

        For Each kvp In valoriInput
            Dim nomeCampo = kvp.Key
            Dim valObj = kvp.Value
            If valObj Is Nothing OrElse Convert.IsDBNull(valObj) Then Continue For
            Dim s = Convert.ToString(valObj)
            If String.IsNullOrWhiteSpace(s) Then Continue For
            Dim m = intervalRegex.Match(s)
            If m.Success Then
                If variableFieldName IsNot Nothing Then
                    ' più di un campo con pattern: non supportato in questa implementazione
                    errorList.Add("Più campi con pattern <da>...<a>... rilevati. Al momento è supportato un solo campo variabile per inserimento batch.")
                    Exit For
                End If
                variableFieldName = nomeCampo
                variableOriginalValue = s

                ' estrai start/end
                Dim startStr = m.Groups(1).Value
                Dim endStr = m.Groups(2).Value
                Dim padLen As Integer = Math.Max(startStr.TrimStart("+"c, "-"c).Length, endStr.TrimStart("+"c, "-"c).Length)

                Dim startVal As Long
                Dim endVal As Long
                If Not Int64.TryParse(startStr, startVal) OrElse Not Int64.TryParse(endStr, endVal) Then
                    errorList.Add($"Pattern intervallo non valido nel campo {nomeCampo}.")
                    Exit For
                End If

                ' costruisci template rimuovendo la porzione <da>...<a>...
                Dim baseTemplate As String = intervalRegex.Replace(s, "{0}")

                ' genera lista valori (step = 1, supporto sia start<end che start>end)
                generatedValues = New List(Of String)
                If startVal <= endVal Then
                    For cur As Long = startVal To endVal
                        Dim formatted = If(padLen > 0, cur.ToString("D" & padLen), cur.ToString())
                        generatedValues.Add(String.Format(baseTemplate, formatted))
                    Next
                Else
                    For cur As Long = startVal To endVal Step -1
                        Dim formatted = If(padLen > 0, cur.ToString("D" & padLen), cur.ToString())
                        generatedValues.Add(String.Format(baseTemplate, formatted))
                    Next
                End If
            End If
        Next

        ' Se abbiamo errori di rilevamento, mostriamoli e usciamo
        If errorList.Count > 0 Then
            Dim msg = String.Join(Environment.NewLine, errorList)
            If Me.InvokeRequired Then
                Me.BeginInvoke(Sub() MDIMessageBox.Show(msg, Me.MdiParent, MessageBoxButtons.OK))
            Else
                MDIMessageBox.Show(msg, Me.MdiParent, MessageBoxButtons.OK)
            End If
            Return
        End If

        Using conn As New SqlConnection(ConnString)
            Await conn.OpenAsync()

            Using tx = conn.BeginTransaction()
                Try
                    If cachedInsertCommand Is Nothing OrElse cachedInsertColumns Is Nothing OrElse Not Enumerable.SequenceEqual(cachedInsertColumns, colonneValid, StringComparer.OrdinalIgnoreCase) Then
                        If cachedInsertCommand IsNot Nothing Then cachedInsertCommand.Dispose()
                        cachedInsertCommand = BuildPreparedInsertCommand(conn, tx, Me.Name, colonneValid)
                        cachedInsertColumns = New List(Of String)(colonneValid)
                    Else
                        cachedInsertCommand.Connection = conn
                        cachedInsertCommand.Transaction = tx
                    End If

                    ' Se non c'è campo variabile, esegui un singolo insert come prima
                    If String.IsNullOrWhiteSpace(variableFieldName) OrElse generatedValues Is Nothing OrElse generatedValues.Count = 0 Then
                        For Each nomeCampo In colonneValid
                            Dim param = cachedInsertCommand.Parameters("@" & nomeCampo)
                            Dim v = If(valoriInput.ContainsKey(nomeCampo), valoriInput(nomeCampo), DBNull.Value)
                            param.Value = If(v Is Nothing, DBNull.Value, v)
                        Next

                        Await cachedInsertCommand.ExecuteNonQueryAsync()
                    Else
                        ' Campo variabile: esegui un insert per ogni valore generato
                        For Each generatedVal In generatedValues
                            For Each nomeCampo In colonneValid
                                Dim param = cachedInsertCommand.Parameters("@" & nomeCampo)
                                Dim v As Object = If(valoriInput.ContainsKey(nomeCampo), valoriInput(nomeCampo), DBNull.Value)
                                If String.Equals(nomeCampo, variableFieldName, StringComparison.OrdinalIgnoreCase) Then
                                    ' sovrascrivo il valore del campo variabile con il valore generato
                                    param.Value = If(generatedVal Is Nothing, DBNull.Value, generatedVal)
                                Else
                                    param.Value = If(v Is Nothing, DBNull.Value, v)
                                End If
                            Next
                            Await cachedInsertCommand.ExecuteNonQueryAsync()
                        Next
                    End If

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

        If errorList.Count > 0 Then
            Dim msg = String.Join(Environment.NewLine, errorList)
            If Me.InvokeRequired Then
                Me.BeginInvoke(Sub() MDIMessageBox.Show(msg, Me.MdiParent, MessageBoxButtons.OK))
            Else
                MDIMessageBox.Show(msg, Me.MdiParent, MessageBoxButtons.OK)
            End If
        End If

        swTotal.Stop()
    End Function


    Private Async Function SalvaModificaAsync() As Task
        Dim swTotal As New Stopwatch()
        swTotal.Start()

        campiDefiniti = RecuperaCampiDa(Me.Name)

        ' Ricalcola i campi join prima di costruire i valori da salvare
        RicalcolaCampiJoinPrimaSalvataggio()

        Try
            Dim convalide = RecuperaConvalideDaSys(Me.Name)
            If convalide IsNot Nothing AndAlso convalide.Count > 0 Then
                For Each campo In campiDefiniti
                    If campo Is Nothing OrElse String.IsNullOrWhiteSpace(campo.Nome) Then Continue For
                    If convalide.ContainsKey(campo.Nome) Then
                        Try
                            ApplicaConvalidaAlCampo(campo, convalide(campo.Nome))
                        Catch ex As Exception
                            MDIMessageBox.Show($"RiapplicaConvalide: errore su campo {campo.Nome}: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)
                        End Try
                    End If
                Next
            End If
        Catch ex As Exception
            MDIMessageBox.Show($"RiapplicaConvalide: errore generale: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)
        End Try

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

        Dim valoriInput As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
        For Each nomeCampo In colonneValid
            Dim valore As Object = Nothing
            If valoriCalcolati.ContainsKey(nomeCampo) Then
                valore = valoriCalcolati(nomeCampo)
            Else
                Try
                    valore = EstraiValoreDaControllo(nomeCampo, campoInputs(nomeCampo))
                Catch ex As Exception
                    valore = DBNull.Value
                End Try
            End If
            If valore Is Nothing Then valore = DBNull.Value
            valoriInput(nomeCampo) = valore
        Next

        Using conn As New SqlConnection(ConnString)
            Await conn.OpenAsync()
            Using tx = conn.BeginTransaction()
                Try
                    If cachedUpdateCommand Is Nothing OrElse cachedUpdateColumns Is Nothing OrElse Not Enumerable.SequenceEqual(cachedUpdateColumns, colonneValid, StringComparer.OrdinalIgnoreCase) Then
                        If cachedUpdateCommand IsNot Nothing Then cachedUpdateCommand.Dispose()
                        cachedUpdateCommand = BuildPreparedUpdateCommand(conn, tx, Me.Name, colonneValid, campoChiave)
                        cachedUpdateColumns = New List(Of String)(colonneValid)
                    Else
                        cachedUpdateCommand.Connection = conn
                        cachedUpdateCommand.Transaction = tx
                    End If

                    For Each nomeCampo In colonneValid
                        Dim param = cachedUpdateCommand.Parameters("@" & nomeCampo)
                        param.Value = valoriInput(nomeCampo)
                    Next

                    cachedUpdateCommand.Parameters("@" & campoChiave.Nome).Value = valoreChiaveObj

                    Dim swExec As New Stopwatch()
                    swExec.Start()
                    Await cachedUpdateCommand.ExecuteNonQueryAsync()
                    swExec.Stop()

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

        swTotal.Stop()
    End Function

    Private Function BuildJoinCacheKey(joinRow As DataRow, chiaviFiglia As Dictionary(Of String, Object)) As String
        Dim sb As New System.Text.StringBuilder()
        ' usa un identificatore stabile per la definizione join (es. TabellaJoin o concatenazione colonne)
        If joinRow.Table.Columns.Contains("IdJoin") Then
            sb.Append(joinRow("IdJoin").ToString())
        ElseIf joinRow.Table.Columns.Contains("TabellaElenco") Then
            sb.Append(joinRow("TabellaElenco").ToString())
        Else
            ' fallback: usa nome campo e hash della definizione
            sb.Append(joinRow.Table.TableName & "_" & joinRow.ItemArray.GetHashCode().ToString())
        End If
        ' aggiungi le chiavi figlia in ordine
        For i As Integer = 1 To 3
            Dim k = $"ChiaveFiglia{i}"
            If chiaviFiglia.ContainsKey(k) Then
                Dim v = chiaviFiglia(k)
                sb.Append("|" & If(v Is Nothing OrElse Convert.IsDBNull(v), "<NULL>", v.ToString()))
            Else
                sb.Append("|<MISSING>")
            End If
        Next
        Return sb.ToString()
    End Function

    Private Function PrelevaValoreJoinCached(joinRow As DataRow, chiaviFiglia As Dictionary(Of String, Object), lookupCache As Dictionary(Of String, DataTable), resultCache As Dictionary(Of String, Object)) As Object
        Dim cacheKey = BuildJoinCacheKey(joinRow, chiaviFiglia)
        If resultCache.ContainsKey(cacheKey) Then
            Return resultCache(cacheKey)
        End If

        ' Prova a risolvere usando lookupCache (evita chiamate DB in PrelevaValoreJoin)
        Try
            If joinRow.Table.Columns.Contains("TabellaElenco") Then
                Dim tabName = Convert.ToString(joinRow("TabellaElenco"))
                Dim chiaveElencoCol = If(joinRow.Table.Columns.Contains("ChiaveElenco"), Convert.ToString(joinRow("ChiaveElenco")), String.Empty)
                Dim campoVisuale = If(joinRow.Table.Columns.Contains("CampoVisuale"), Convert.ToString(joinRow("CampoVisuale")), String.Empty)
                Dim campoDaPrelevare = If(joinRow.Table.Columns.Contains("CampoDaPrelevare"), Convert.ToString(joinRow("CampoDaPrelevare")), String.Empty)

                If Not String.IsNullOrWhiteSpace(tabName) AndAlso lookupCache.ContainsKey(tabName) Then
                    Dim dtRef = lookupCache(tabName)

                    ' Costruisci filtro usando le colonne ChiavePadreN (non la stessa chiaveElenco)
                    Dim filtroParts As New List(Of String)
                    For i As Integer = 1 To 3
                        Dim chiavePadreCol = $"ChiavePadre{i}"
                        Dim chiaveFigliaKey = $"ChiaveFiglia{i}"
                        If joinRow.Table.Columns.Contains(chiavePadreCol) AndAlso chiaviFiglia.ContainsKey(chiaveFigliaKey) Then
                            Dim colPadreName = Convert.ToString(joinRow(chiavePadreCol))
                            Dim v = chiaviFiglia(chiaveFigliaKey)
                            If Not String.IsNullOrWhiteSpace(colPadreName) AndAlso v IsNot Nothing AndAlso Not Convert.IsDBNull(v) Then
                                Dim escaped = v.ToString().Replace("'", "''")
                                filtroParts.Add($"[{colPadreName}] = '{escaped}'")
                            End If
                        End If
                    Next

                    If filtroParts.Count > 0 Then
                        Dim filtro = String.Join(" AND ", filtroParts.ToArray())
                        Dim rows() As DataRow = Nothing
                        Try
                            rows = dtRef.Select(filtro)
                        Catch ex As Exception
                            Debug.WriteLine($"[JOIN CACHE] Errore Select su DataTable '{tabName}' con filtro '{filtro}': {ex.ToString()}")
                            rows = New DataRow() {}
                        End Try

                        If rows IsNot Nothing AndAlso rows.Length > 0 Then
                            ' Determina quale colonna restituire: preferisci CampoDaPrelevare, poi CampoVisuale, poi ChiaveElenco
                            Dim colToReturn As String = Nothing
                            If Not String.IsNullOrWhiteSpace(campoDaPrelevare) Then
                                colToReturn = campoDaPrelevare
                            ElseIf Not String.IsNullOrWhiteSpace(campoVisuale) Then
                                colToReturn = campoVisuale
                            ElseIf Not String.IsNullOrWhiteSpace(chiaveElencoCol) Then
                                colToReturn = chiaveElencoCol
                            End If

                            If Not String.IsNullOrWhiteSpace(colToReturn) AndAlso rows(0).Table.Columns.Contains(colToReturn) Then
                                Dim val = rows(0)(colToReturn)
                                resultCache(cacheKey) = val
                                Return val
                            Else
                                Debug.WriteLine($"[JOIN CACHE] Colonna da restituire '{colToReturn}' non trovata in tabella '{tabName}'.")
                            End If
                        Else
                            Debug.WriteLine($"[JOIN CACHE] Nessuna riga trovata in '{tabName}' per filtro: {filtro}")
                        End If
                    Else
                        Debug.WriteLine("[JOIN CACHE] Nessuna chiave figlia valida per costruire il filtro.")
                    End If
                End If
            End If
        Catch ex As Exception
            'Debug.WriteLine($"[JOIN CACHE] Errore durante lookupCache: {ex.ToString()}")
            ' fallback al metodo originale
        End Try

        ' Fallback: chiama la funzione esistente (potrebbe fare query)
        Try
            Dim valore = PrelevaValoreJoin(joinRow, chiaviFiglia)
            resultCache(cacheKey) = valore
            Return valore
        Catch ex As Exception
            'Debug.WriteLine($"[JOIN CACHE] Errore fallback PrelevaValoreJoin: {ex.ToString()}")
            resultCache(cacheKey) = Nothing
            Return Nothing
        End Try
    End Function

    Private Function PrelevaValoreJoin(joinRow As DataRow, chiaviFiglia As Dictionary(Of String, Object)) As Object

        If joinRow Is Nothing Then
            MDIMessageBox.Show("PrelevaValoreJoin: joinRow è Nothing", Me.MdiParent, MessageBoxButtons.OK)
            Return Nothing
        End If

        If chiaviFiglia Is Nothing OrElse chiaviFiglia.Count = 0 Then
            MDIMessageBox.Show("PrelevaValoreJoin: chiaviFiglia è Nothing o vuoto", Me.MdiParent, MessageBoxButtons.OK)
            Return Nothing
        End If

        Dim tabellaPadre As String = ""
        Dim campoDaPrelevare As String = ""

        Try
            tabellaPadre = Convert.ToString(joinRow("TabellaPadre"))
            campoDaPrelevare = Convert.ToString(joinRow("CampoDaPrelevare"))
        Catch ex As Exception
            MDIMessageBox.Show($"PrelevaValoreJoin: mancata lettura colonne joinRow: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)
            Return Nothing
        End Try

        If String.IsNullOrWhiteSpace(tabellaPadre) OrElse String.IsNullOrWhiteSpace(campoDaPrelevare) Then
            MDIMessageBox.Show("PrelevaValoreJoin: tabellaPadre o campoDaPrelevare vuoti", Me.MdiParent, MessageBoxButtons.OK)
            Return Nothing
        End If

        Dim identPattern As String = "^[\w\.]+$"
        If Not Regex.IsMatch(tabellaPadre, identPattern) OrElse Not Regex.IsMatch(campoDaPrelevare, "^[\w]+$") Then
            MDIMessageBox.Show($"PrelevaValoreJoin: nome tabella o campo non valido: {tabellaPadre}.{campoDaPrelevare}", Me.MdiParent, MessageBoxButtons.OK)
            Return Nothing
        End If

        Dim condizioni As New List(Of String)()
        Dim parametri As New Dictionary(Of String, Object)()

        For i = 1 To 3
            Dim nomeColonna = $"ChiavePadre{i}"
            If joinRow.Table.Columns.Contains(nomeColonna) Then
                Dim chiavePadre = Convert.ToString(joinRow(nomeColonna))
                If Not String.IsNullOrWhiteSpace(chiavePadre) AndAlso chiaviFiglia.ContainsKey($"ChiaveFiglia{i}") Then
                    If Not Regex.IsMatch(chiavePadre, "^[\w]+$") Then
                        MDIMessageBox.Show($"PrelevaValoreJoin: nome colonna padre non valido: {chiavePadre}", Me.MdiParent, MessageBoxButtons.OK)
                        Continue For
                    End If

                    Dim paramName = $"@param{i}"
                    condizioni.Add($"{chiavePadre} = {paramName}")
                    parametri.Add(paramName, chiaviFiglia($"ChiaveFiglia{i}"))
                End If
            End If
        Next

        If condizioni.Count = 0 Then
            MDIMessageBox.Show("PrelevaValoreJoin: nessuna condizione costruita, restituisco Nothing", Me.MdiParent, MessageBoxButtons.OK)
            Return Nothing
        End If

        Dim query As String = $"SELECT {campoDaPrelevare} FROM [{tabellaPadre}] WHERE {String.Join(" AND ", condizioni)}"

        Try
            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand(query, conn)
                    cmd.CommandTimeout = 30
                    For Each kvp In parametri
                        Dim val = If(kvp.Value, DBNull.Value)
                        cmd.Parameters.AddWithValue(kvp.Key, val)
                    Next

                    conn.Open()
                    Dim result = cmd.ExecuteScalar()
                    Return If(result Is DBNull.Value, Nothing, result)
                End Using
            End Using

        Catch ex As SqlException
            Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show($"Errore SQL recuperando join da {tabellaPadre}: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)))
            Return Nothing
        Catch ex As Exception
            Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show($"Errore imprevisto recuperando join: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)))
            Return Nothing
        End Try
    End Function

    Private Function RecuperaCampiCalcolatiDettaglio() As Dictionary(Of String, (Formula As String, TipoValore As String, SuSeStesso As Boolean))
        Dim diz As New Dictionary(Of String, (String, String, Boolean))(StringComparer.OrdinalIgnoreCase)
        Using conn As New SqlConnection(ConnString)
            Dim query = "SELECT NomeCampo, Formula, Tipovalore, SuSeStesso FROM Sys_CampiCalcolati WHERE NomeTabella = @NomeTabella"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@NomeTabella", Me.Name)
                conn.Open()
                Using reader = cmd.ExecuteReader()
                    Dim hasSuSeStesso As Boolean = Enumerable.Range(0, reader.FieldCount).Any(Function(i) String.Equals(reader.GetName(i), "SuSeStesso", StringComparison.OrdinalIgnoreCase))
                    While reader.Read()
                        Dim nome = If(reader.IsDBNull(reader.GetOrdinal("NomeCampo")), String.Empty, reader("NomeCampo").ToString()).Trim()
                        If String.IsNullOrEmpty(nome) Then Continue While
                        Dim formula = If(HasColumn(reader, "Formula") AndAlso Not reader.IsDBNull(reader.GetOrdinal("Formula")), reader("Formula").ToString(), String.Empty)
                        Dim tipo = If(HasColumn(reader, "Tipovalore") AndAlso Not reader.IsDBNull(reader.GetOrdinal("Tipovalore")), reader("Tipovalore").ToString().ToLowerInvariant(), "numero")
                        Dim selfRef As Boolean = False
                        If hasSuSeStesso AndAlso Not reader.IsDBNull(reader.GetOrdinal("SuSeStesso")) Then
                            Try
                                selfRef = Convert.ToBoolean(reader("SuSeStesso"))
                            Catch
                                selfRef = False
                            End Try
                        End If
                        diz(nome) = (formula, tipo, selfRef)
                    End While
                End Using
            End Using
        End Using
        Return diz
    End Function

    Private Sub CaricaCacheCampiCalcolati()
        Try
            campiCalcolatiDettaglio = RecuperaCampiCalcolatiDettaglio()
        Catch
            campiCalcolatiDettaglio = New Dictionary(Of String, (String, String, Boolean))(StringComparer.OrdinalIgnoreCase)
        End Try
    End Sub

    Private Function EstraiValoreDaControllo(nomeCampo As String, input As Control) As Object
        Dim campiBit As String() = {
            "CanView", "CanInsert", "CanUpdate", "CanDelete"}
        Dim isPassword = nomeCampo.ToLower().Contains("password")

        Select Case True
            Case campiBit.Contains(nomeCampo, StringComparer.OrdinalIgnoreCase) AndAlso TypeOf input Is CheckBox
                Return If(CType(input, CheckBox).Checked, 1, 0)

            Case TypeOf input Is CheckBox
                Return If(CType(input, CheckBox).Checked, 1, 0)

            Case TypeOf input Is ComboBox
                Return CType(input, ComboBox).SelectedValue

            Case TypeOf input Is FlowLayoutPanel
                Dim txt As TextBox = input.Controls.OfType(Of TextBox)().FirstOrDefault()
                If txt IsNot Nothing Then
                    Dim valore = txt.Text.Trim()
                    Dim campoDef As CampoDatabase = Nothing
                    Try
                        campoDef = campiDefiniti.FirstOrDefault(Function(c) c.Nome.Equals(nomeCampo, StringComparison.OrdinalIgnoreCase))
                    Catch
                        campoDef = Nothing
                    End Try

                    If campoDef IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(campoDef.TabellaElenco) Then
                        Dim m = Regex.Match(valore, "^\s*([^\-]+?)\s*\-\s*(.+)$")
                        If m.Success Then
                            Return m.Groups(1).Value.Trim()
                        End If
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

            Case Else
                Return input.Text
        End Select
    End Function

    Private Function CreaControllo(campo As CampoDatabase) As Control

        If campo Is Nothing Then Return CreaLabelErrore("Campo nothing.")

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

        ' Gestione campi path 
        If IsCampoPath(campo.Nome) Then
            Dim info = RecuperaInfoCampoPath(campo.Nome)
            Dim isFile = info.IsFile
            Dim showButton = info.BottoneVisualizza

            Dim txt As New TextBox With {
        .Width = Math.Max(larghezzaStimata - If(showButton, 90, 0), 100),
        .Height = 23,
        .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
        .Tag = campo,
        .Margin = New Padding(3)
    }

            ' ToolTip per mostrare il percorso completo
            Dim tt As New ToolTip()
            tt.IsBalloon = False
            tt.ShowAlways = True
            tt.AutoPopDelay = 10000
            tt.InitialDelay = 300
            tt.ReshowDelay = 100

            ' funzione di utilità locale per aggiornare tooltip
            Dim AggiornaToolTip = Sub()
                                      Try
                                          If String.IsNullOrWhiteSpace(txt.Text) Then
                                              tt.SetToolTip(txt, String.Empty)
                                          Else
                                              tt.SetToolTip(txt, txt.Text)
                                          End If
                                      Catch
                                      End Try
                                  End Sub

            AggiornaToolTip()

            AddHandler txt.DoubleClick, Sub(s, e)
                                            Try
                                                Dim startDir As String = Nothing
                                                Dim curText = txt.Text.Trim()

                                                If Not String.IsNullOrWhiteSpace(curText) Then
                                                    Try
                                                        If File.Exists(curText) Then
                                                            startDir = Path.GetDirectoryName(curText)
                                                        ElseIf Directory.Exists(curText) Then
                                                            startDir = curText
                                                        Else
                                                            Try
                                                                startDir = Path.GetDirectoryName(curText)
                                                            Catch
                                                                startDir = Nothing
                                                            End Try
                                                        End If
                                                    Catch
                                                        startDir = Nothing
                                                    End Try
                                                End If

                                                If String.IsNullOrWhiteSpace(startDir) Then
                                                    Try
                                                        Using conn As New SqlConnection(ConnString)
                                                            Using cmd As New SqlCommand("SELECT TOP 1 Valore FROM Sys_Parametri WHERE Descrizione = @DescPar", conn)
                                                                cmd.Parameters.AddWithValue("@DescPar", "PercorsoDefaultPath")
                                                                conn.Open()
                                                                Dim res = cmd.ExecuteScalar()
                                                                If res IsNot Nothing AndAlso Not Convert.IsDBNull(res) Then
                                                                    Dim candidate = res.ToString().Trim()
                                                                    If Directory.Exists(candidate) Then
                                                                        startDir = candidate
                                                                    ElseIf File.Exists(candidate) Then
                                                                        startDir = Path.GetDirectoryName(candidate)
                                                                    End If
                                                                End If
                                                            End Using
                                                        End Using
                                                    Catch
                                                        startDir = Nothing
                                                    End Try
                                                End If

                                                If isFile Then
                                                    Using ofd As New OpenFileDialog()
                                                        ofd.CheckFileExists = False
                                                        ofd.CheckPathExists = True
                                                        ofd.Multiselect = False
                                                        ofd.Title = "Seleziona file"
                                                        If Not String.IsNullOrWhiteSpace(startDir) AndAlso Directory.Exists(startDir) Then
                                                            ofd.InitialDirectory = startDir
                                                        End If

                                                        If Not String.IsNullOrWhiteSpace(curText) AndAlso File.Exists(curText) Then
                                                            ofd.FileName = Path.GetFileName(curText)
                                                        End If

                                                        If ofd.ShowDialog(Me) = DialogResult.OK Then
                                                            txt.Text = ofd.FileName
                                                        End If
                                                    End Using
                                                Else
                                                    Using fbd As New FolderBrowserDialog()
                                                        fbd.Description = "Seleziona cartella"
                                                        fbd.ShowNewFolderButton = True
                                                        If Not String.IsNullOrWhiteSpace(startDir) AndAlso Directory.Exists(startDir) Then
                                                            Try
                                                                fbd.SelectedPath = startDir
                                                            Catch
                                                            End Try
                                                        End If

                                                        If fbd.ShowDialog(Me) = DialogResult.OK Then
                                                            txt.Text = fbd.SelectedPath
                                                        End If
                                                    End Using
                                                End If
                                            Catch ex As Exception
                                                MDIMessageBox.Show("Errore selezione path: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
                                            End Try
                                        End Sub


            Dim panel As New FlowLayoutPanel With {
        .AutoSize = True,
        .FlowDirection = FlowDirection.LeftToRight,
        .WrapContents = False,
        .Height = txt.Height + 6,
        .Margin = New Padding(0),
        .Padding = New Padding(0)
    }

            panel.Controls.Add(txt)

            If showButton Then
                Dim btnView As New Button With {
            .Text = "Visualizza",
            .AutoSize = True,
            .Margin = New Padding(3)
        }

                AddHandler txt.TextChanged, Sub()
                                                btnView.Enabled = Not String.IsNullOrWhiteSpace(txt.Text)
                                            End Sub

                AddHandler btnView.Click, Sub(s, e)
                                              Dim val = txt.Text.Trim()
                                              If String.IsNullOrWhiteSpace(val) Then
                                                  MDIMessageBox.Show("Nessun percorso specificato", Me.MdiParent, MessageBoxButtons.OK)
                                                  Return
                                              End If
                                              Try
                                                  If isFile Then
                                                      If Not File.Exists(val) Then
                                                          MDIMessageBox.Show("File non trovato: " & val, Me.MdiParent, MessageBoxButtons.OK)
                                                          Return
                                                      End If
                                                      Process.Start(New ProcessStartInfo(val) With {.UseShellExecute = True})
                                                  Else
                                                      If Not Directory.Exists(val) Then
                                                          MDIMessageBox.Show("Cartella non trovata: " & val, Me.MdiParent, MessageBoxButtons.OK)
                                                          Return
                                                      End If
                                                      Process.Start(New ProcessStartInfo("explorer.exe", """" & val & """") With {.UseShellExecute = True})
                                                  End If
                                              Catch ex As Exception
                                                  MDIMessageBox.Show("Errore apertura risorsa: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
                                              End Try
                                          End Sub

                AddHandler panel.MouseEnter, Sub()
                                                 If Not String.IsNullOrWhiteSpace(txt.Text) Then
                                                     tt.Show(txt.Text, panel, panel.PointToClient(Cursor.Position), 5000)
                                                 End If
                                             End Sub
                AddHandler panel.MouseLeave, Sub() tt.Hide(panel)
                AddHandler txt.TextChanged, Sub()
                                                AggiornaToolTip()
                                            End Sub

                panel.Controls.Add(btnView)

                Try
                    Dim btnW = btnView.PreferredSize.Width + btnView.Margin.Horizontal
                    AggiornaToolTip()
                Catch
                End Try
            End If

            Return panel
        End If

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
            Dim fieldHeight As Integer = 23
            Dim txt = New TextBox With {
                                            .Width = 100,
                                            .Height = fieldHeight,
                                            .Multiline = False,
                                            .AutoSize = False,
                                            .Tag = campo.Nome,
                                            .Margin = New Padding(3, 3, 3, 3)
                                        }

            Dim lblDescrizione = New Label With {
                                            .Width = 200,
                                            .Height = fieldHeight,
                                            .AutoSize = True,
                                            .TextAlign = ContentAlignment.MiddleLeft,
                                            .ForeColor = Color.DarkSlateGray,
                                            .Padding = New Padding(5, 0, 0, 0),
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
                                      txt.BackColor = SystemColors.Window
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
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = False,
                .Height = fieldHeight + 6,
                .Margin = New Padding(0),
                .Padding = New Padding(0)
    }

            txt.Anchor = AnchorStyles.Left
            lblDescrizione.Anchor = AnchorStyles.Left

            pannello.Controls.Add(txt)
            pannello.Controls.Add(lblDescrizione)

            txt.Tag = campo

            Return pannello
        End If

        If campo.TipoConvalida = "I" Then
            AddHandler ctrl.Validated, Sub(sender, e)
                                           ValidazioneIntervallo(campo, CType(sender, Control))
                                       End Sub
        End If

        Return ctrl
    End Function

    Private Function RecuperaInfoCampoPath(nomeCampo As String) As (IsFile As Boolean, BottoneVisualizza As Boolean)
        Try
            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand("SELECT TOP 1 IsFile, BottoneVisualizza FROM Sys_CampiPath WHERE NomeCampoPath = @NomeCampo", conn)
                    cmd.Parameters.AddWithValue("@NomeCampo", nomeCampo)
                    conn.Open()
                    Using reader = cmd.ExecuteReader()
                        If reader.Read() Then
                            Dim isFile As Boolean = False
                            Dim hasIsFile = False
                            Dim bottoneVis As Boolean = False
                            Dim hasBottoneVis = False

                            If HasColumn(reader, "IsFile") AndAlso Not reader.IsDBNull(reader.GetOrdinal("IsFile")) Then
                                hasIsFile = True
                                Try
                                    isFile = Convert.ToBoolean(reader("IsFile"))
                                Catch
                                    isFile = False
                                End Try
                            End If

                            If HasColumn(reader, "BottoneVisualizza") AndAlso Not reader.IsDBNull(reader.GetOrdinal("BottoneVisualizza")) Then
                                hasBottoneVis = True
                                Try
                                    bottoneVis = Convert.ToBoolean(reader("BottoneVisualizza"))
                                Catch
                                    bottoneVis = False
                                End Try
                            End If

                            Return (If(hasIsFile, isFile, True), If(hasBottoneVis, bottoneVis, False))
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show($"RecuperaInfoCampoPath errore: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)))
        End Try

        Return (True, False)
    End Function

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
            MDIMessageBox.Show($"Errore recupero tabella {nomeTabella}: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)
        End Try
        Return dt
    End Function

    Private Function RecuperaTabella(nomeTabella As String, Optional throwOnError As Boolean = False) As DataTable
        Dim dt As New DataTable()

        If String.IsNullOrWhiteSpace(nomeTabella) OrElse Not System.Text.RegularExpressions.Regex.IsMatch(nomeTabella, "^[\w\.]+$") Then
            Dim msg As String = $"Nome tabella non valido: {nomeTabella}"
            System.Diagnostics.Trace.TraceError(msg)
            MDIMessageBox.Show(msg, Me.MdiParent, MessageBoxButtons.OK)
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
            ApplicaFeedbackVisivo(inputControl, false)
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
        dtp.Anchor = AnchorStyles.Left

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
                                      'If String.IsNullOrWhiteSpace(percorso) Then
                                      'MDIMessageBox.Show("Percorso multimediale non configurato.", Me.MdiParent, MessageBoxButtons.OK)
                                      'Return
                                      'End If

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

        AddHandler txtFileName.DoubleClick, Sub(s, e)
                                                Using ofd As New OpenFileDialog()
                                                    ofd.CheckFileExists = True
                                                    ofd.CheckPathExists = True
                                                    ofd.Multiselect = False
                                                    ofd.Title = "Seleziona file Immagine o Video"
                                                    Try
                                                        Dim cur = txtFileName.Text.Trim()
                                                        If Not String.IsNullOrWhiteSpace(cur) Then
                                                            Dim dir = Path.GetDirectoryName(cur)
                                                            If Directory.Exists(dir) Then ofd.InitialDirectory = dir
                                                        End If
                                                    Catch
                                                    End Try

                                                    If ofd.ShowDialog(Me) = DialogResult.OK Then
                                                        txtFileName.Text = ofd.FileName
                                                    End If
                                                End Using
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

        'Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show("Nella Tabella Sys_Parametri non è stato trovato nessun risultato", Me.MdiParent, MessageBoxButtons.OK)))
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
            MDIMessageBox.Show("Errore SQL: " & ex.Message & " Query: " & query, Me.MdiParent, MessageBoxButtons.OK)
        Catch ex As Exception
            MDIMessageBox.Show("Errore generico: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
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

        Try
            dgvDati.ClearSelection()

            Dim row As DataGridViewRow = dgvDati.Rows(e.RowIndex)
            row.Selected = True

            CaricaDatiNeiControlli(row)

            ModalitaCorrente = "visualizzazione"
            lblModalita.Text = "Visualizzazione in corso..."
            UpdateButtonsByModalita()
        Catch ex As Exception
            MDIMessageBox.Show("dgvDati_CellClick error: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
        End Try
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

    Private Function IntestazioneMultilinea(nomeColonna As String) As String

        Dim testo = SpaziaMaiuscole(nomeColonna)
        If String.IsNullOrWhiteSpace(testo) Then Return String.Empty

        Dim parti = testo.Split(" "c).Where(Function(s) Not String.IsNullOrWhiteSpace(s)).ToArray()
        If parti.Length <= 2 Then
            Return testo
        End If

        Dim primaRiga = parti(0)
        Dim secondaRiga = String.Join(" "c, parti.Skip(1))
        Return primaRiga & vbCrLf & secondaRiga
    End Function

    Private Sub ApplicaAutorizzazioni(nomeUtente As String)
        Try
            Using conn As New SqlConnection(ConnString)
                conn.Open()

                Dim isAdmin As Boolean = False
                Dim queryAdmin = "SELECT ISNULL(Amministratore, 0) FROM Tab_Utenti WHERE NomeUtente = @utente"

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
                FROM Tab_UtentiAutorizzazioni 
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
        Dim originalAutoSizeMode = dgv.AutoSizeColumnsMode
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

        Try
            Dim dtConfig As New DataTable()
            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand("
                SELECT NomeColonna, ColWidth, VisualizzaInDbgrid
                FROM Sys_VisualizzaInDbgrid
                WHERE NomeTabella = @NomeTabella AND NomeDbgrid = @NomeDbgrid", conn)
                    cmd.Parameters.AddWithValue("@NomeTabella", nomeTabellaCorrente)
                    cmd.Parameters.AddWithValue("@NomeDbgrid", dgv.Name)
                    Using da As New SqlDataAdapter(cmd)
                        da.Fill(dtConfig)
                    End Using
                End Using
            End Using

            Dim configMap As New Dictionary(Of String, (Width As Integer, Visible As Boolean))(StringComparer.OrdinalIgnoreCase)
            For Each r As DataRow In dtConfig.Rows
                Dim name As String = If(r.IsNull("NomeColonna"), String.Empty, r("NomeColonna").ToString())
                If String.IsNullOrWhiteSpace(name) Then Continue For
                Dim w As Integer = 0
                If Not r.IsNull("ColWidth") Then
                    Integer.TryParse(r("ColWidth").ToString(), w)
                End If
                Dim v As Boolean = True
                If Not r.IsNull("VisualizzaInDbgrid") Then
                    Boolean.TryParse(r("VisualizzaInDbgrid").ToString(), v)
                End If
                configMap(name) = (w, v)
            Next

            Dim dgvColsByName As New Dictionary(Of String, DataGridViewColumn)(StringComparer.OrdinalIgnoreCase)
            For Each col As DataGridViewColumn In dgv.Columns
                dgvColsByName(col.Name) = col
            Next

            For Each kvp In dgvColsByName
                Dim col = kvp.Value
                If configMap.ContainsKey(col.Name) Then
                    col.Visible = configMap(col.Name).Visible
                Else
                    col.Visible = True
                End If
            Next

            For Each kvp In configMap
                Dim colName = kvp.Key
                Dim cfg = kvp.Value
                Dim col As DataGridViewColumn = Nothing
                If dgvColsByName.TryGetValue(colName, col) Then
                    If cfg.Width > 0 Then
                        Try
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            col.Width = cfg.Width
                        Catch ex As Exception
                            MDIMessageBox.Show($"ApplicaConfigurazioneGriglia: errore impostando width su '{colName}': {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)
                        End Try
                    End If
                End If
            Next

            Dim colsToAutoSize As New List(Of Integer)
            For Each col As DataGridViewColumn In dgv.Columns
                Dim hasCfg = configMap.ContainsKey(col.Name)
                If hasCfg Then
                    If configMap(col.Name).Width <= 0 AndAlso col.Visible Then
                        colsToAutoSize.Add(col.Index)
                    End If
                Else
                    If col.Visible Then colsToAutoSize.Add(col.Index)
                End If
            Next

            If colsToAutoSize.Count > 0 Then
                For Each idx In colsToAutoSize
                    Try
                        dgv.AutoResizeColumn(idx, DataGridViewAutoSizeColumnMode.AllCells)
                        Dim c = dgv.Columns(idx)
                        Dim computed = c.Width
                        c.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                        c.Width = computed
                    Catch ex As Exception
                        MDIMessageBox.Show($"ApplicaConfigurazioneGriglia: errore autosize col index {idx}: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)
                    End Try
                Next
            End If

        Finally
            dgv.AutoSizeColumnsMode = originalAutoSizeMode
            dgv.ResumeLayout()
            dgv.Invalidate()
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

    Private Function RecuperaCampiCalcolati() As Dictionary(Of String, (Formula As String, TipoValore As String, SuSeStesso As Boolean))
        Dim diz As New Dictionary(Of String, (String, String, Boolean))(StringComparer.OrdinalIgnoreCase)
        Using conn As New SqlConnection(ConnString)
            Dim query = "SELECT NomeCampo, Formula, Tipovalore, SuSeStesso FROM Sys_CampiCalcolati WHERE NomeTabella = @NomeTabella"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@NomeTabella", Me.Name)
                conn.Open()
                Using reader = cmd.ExecuteReader()
                    Dim hasSuSeStesso As Boolean = Enumerable.Range(0, reader.FieldCount).Any(Function(i) String.Equals(reader.GetName(i), "SuSeStesso", StringComparison.OrdinalIgnoreCase))
                    While reader.Read()
                        Dim nome = reader("NomeCampo").ToString()
                        Dim formula = If(reader.IsDBNull(reader.GetOrdinal("Formula")), String.Empty, reader("Formula").ToString())
                        Dim tipo = If(reader.IsDBNull(reader.GetOrdinal("Tipovalore")), "numero", reader("Tipovalore").ToString().ToLowerInvariant())
                        Dim selfRef As Boolean = False
                        If hasSuSeStesso AndAlso Not reader.IsDBNull(reader.GetOrdinal("SuSeStesso")) Then
                            Try
                                selfRef = Convert.ToBoolean(reader("SuSeStesso"))
                            Catch
                                selfRef = False
                            End Try
                        End If
                        diz(nome) = (formula, tipo, selfRef)
                    End While
                End Using
            End Using
        End Using
        Return diz
    End Function

    Private campiCalcolatiSet As HashSet(Of String) = Nothing

    Private Function RecuperaCampiCalcolatiSet(nomeTabella As String) As HashSet(Of String)
        Dim hs As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Try
            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand("SELECT NomeCampo FROM Sys_CampiCalcolati WHERE NomeTabella = @NomeTabella", conn)
                    cmd.Parameters.AddWithValue("@NomeTabella", nomeTabella)
                    Dim dt As New DataTable()
                    Using da As New SqlDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                    For Each r As DataRow In dt.Rows
                        Dim nome = r("NomeCampo").ToString()
                        If Not String.IsNullOrWhiteSpace(nome) Then hs.Add(nome.Trim())
                    Next
                End Using
            End Using
        Catch
            ' fallback: lista vuota -> nessun campo considerato calcolato
        End Try
        Return hs
    End Function

    Private Function PadWithZeros(val As Object, totalWidth As Integer) As String
        Dim s As String = If(val Is Nothing OrElse val Is DBNull.Value, String.Empty, val.ToString())
        If totalWidth <= 0 Then Return s
        If s.Length >= totalWidth Then Return s
        Return New String("0"c, totalWidth - s.Length) & s
    End Function

    Private Function ReplaceAddZeroForCampo(expr As String,
                                       nomeCampo As String,
                                       campoInputs As Dictionary(Of String, Control),
                                       valoriCalcolati As Dictionary(Of String, Object)) As String
        If String.IsNullOrWhiteSpace(expr) Then Return expr

        ' ottieni il valore corrente da usare per il padding (prima i valori calcolati, poi il controllo)
        Dim valoreRaw As Object = Nothing
        If valoriCalcolati IsNot Nothing AndAlso valoriCalcolati.ContainsKey(nomeCampo) Then
            valoreRaw = valoriCalcolati(nomeCampo)
        ElseIf campoInputs IsNot Nothing AndAlso campoInputs.ContainsKey(nomeCampo) Then
            Try
                valoreRaw = EstraiValoreDaControllo(nomeCampo, campoInputs(nomeCampo))
            Catch
                valoreRaw = Nothing
            End Try
        End If

        ' regex per AddZero( N )
        Dim rx As New Regex("AddZero\s*\(\s*(\d+)\s*\)", RegexOptions.IgnoreCase Or RegexOptions.Compiled)

        Dim result As String = rx.Replace(expr, Function(m As Match)
                                                    Dim digits As Integer = 0
                                                    Integer.TryParse(m.Groups(1).Value, digits)

                                                    Dim padded As String = PadWithZeros(valoreRaw, digits)
                                                    Dim escaped = padded.Replace("""", "\""")
                                                    Return """" & escaped & """"
                                                End Function)

        Return result
    End Function

    Private Function CalcolaValoriCampiCalcolati(formule As Dictionary(Of String, String),
                                              tipiValore As Dictionary(Of String, String)) As Dictionary(Of String, Object)

        Dim risultati As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)

        ' Provo a recuperare i dettagli (Formula, TipoValore, SuSeStesso) se disponibile
        Dim campiDettaglio As Dictionary(Of String, (Formula As String, TipoValore As String, SuSeStesso As Boolean)) = Nothing
        Try
            campiDettaglio = RecuperaCampiCalcolati()
        Catch
            campiDettaglio = New Dictionary(Of String, (String, String, Boolean))(StringComparer.OrdinalIgnoreCase)
        End Try

        For Each kvp In formule
            Dim nomeCampo As String = kvp.Key
            Dim exprRaw As String = If(kvp.Value, String.Empty)
            Dim tipo As String = If(tipiValore IsNot Nothing AndAlso tipiValore.ContainsKey(nomeCampo),
                                tipiValore(nomeCampo).ToLowerInvariant(), "numero")
            If Not {"numero", "stringa", "data", "booleano", "raw", "testo"}.Contains(tipo) Then
                tipo = "numero"
            End If

            Try
                Dim expr As String = exprRaw

                ' 1) sostituisco i placeholder dei campi con i loro valori
                For Each vInput In campoInputs
                    Dim phName As String = vInput.Key
                    Dim controllo As Control = vInput.Value

                    ' Determina se per questo placeholder dobbiamo forzare l'uso del valore dal controllo
                    Dim forceControlValue As Boolean = False
                    If String.Equals(phName, nomeCampo, StringComparison.OrdinalIgnoreCase) Then
                        If campiDettaglio IsNot Nothing AndAlso campiDettaglio.ContainsKey(nomeCampo) Then
                            forceControlValue = campiDettaglio(nomeCampo).SuSeStesso
                        End If
                    End If

                    Dim valoreObj As Object = Nothing
                    If forceControlValue Then
                        valoreObj = EstraiValoreDaControllo(phName, controllo)
                    Else
                        If risultati.ContainsKey(phName) Then
                            valoreObj = risultati(phName)
                        Else
                            valoreObj = EstraiValoreDaControllo(phName, controllo)
                        End If
                    End If

                    ' Costruisco la stringa di sostituzione in modo sicuro
                    Dim sost As String
                    If valoreObj Is Nothing OrElse valoreObj Is DBNull.Value Then
                        If tipo = "stringa" OrElse tipo = "testo" OrElse tipo = "raw" Then
                            sost = """" & "" & """"
                        Else
                            sost = "0"
                        End If
                    ElseIf TypeOf valoreObj Is DateTime Then
                        sost = """" & CType(valoreObj, DateTime).ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture) & """"
                    ElseIf tipo = "stringa" OrElse tipo = "testo" OrElse tipo = "raw" Then
                        ' Escape delle doppie virgolette interne: raddoppio o escape con backslash a seconda delle funzioni di valutazione
                        Dim s As String = valoreObj.ToString()
                        s = s.Replace("""", """""""") ' raddoppia le doppie virgolette per DataTable/Expression
                        sost = """" & s & """"
                    ElseIf TypeOf valoreObj Is Boolean Then
                        sost = If(Convert.ToBoolean(valoreObj), "1", "0")
                    Else
                        ' numero o generico: conversione con invariant culture
                        sost = Convert.ToString(valoreObj, Globalization.CultureInfo.InvariantCulture)
                        If String.IsNullOrWhiteSpace(sost) Then sost = "0"
                    End If

                    ' Usa regex cache per performance e per sostituire solo parole intere
                    Dim rx As Regex = Nothing
                    If regexCache.ContainsKey(phName) Then
                        rx = regexCache(phName)
                    Else
                        rx = New Regex("\b" & Regex.Escape(phName) & "\b", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
                        regexCache(phName) = rx
                    End If

                    expr = rx.Replace(expr, sost)
                Next

                ' 2) Replace AddZero(...) per il campo corrente
                expr = ReplaceAddZeroForCampo(expr, nomeCampo, campoInputs, risultati)

                ' 3) Pulizia espressione: evito espressioni vuote o solo virgolette
                Dim exprPulita = expr.Trim()
                If String.IsNullOrEmpty(exprPulita) OrElse Regex.IsMatch(exprPulita, "^\s*""*\s*""*\s*$") Then
                    ' espressione vuota o solo quote -> risultato nullo o stringa vuota a seconda del tipo
                    Select Case tipo
                        Case "stringa", "testo", "raw"
                            risultati(nomeCampo) = String.Empty
                        Case "data", "booleano"
                            risultati(nomeCampo) = Nothing
                        Case Else
                            risultati(nomeCampo) = Nothing
                    End Select
                    Continue For
                End If

                ' 4) valutazione finale in base al tipo, con try/catch di protezione
                Dim risultato As Object = Nothing
                Select Case tipo
                    Case "numero"
                        Try
                            ' rimuovo eventuali virgolette residue e valuto con DataTable().Compute
                            Dim eNumerica = exprPulita.Replace("""", "")
                            If String.IsNullOrWhiteSpace(eNumerica) Then
                                risultato = Nothing
                            Else
                                risultato = New DataTable().Compute(eNumerica, Nothing)
                                If Not IsNumeric(risultato) Then risultato = Nothing
                            End If
                        Catch exCompute As Exception
                            risultato = Nothing
                        End Try

                    Case "stringa", "testo", "raw"
                        Try
                            risultato = ValutaEspressioneStringa(exprPulita)
                        Catch
                            risultato = ""
                        End Try

                    Case "data"
                        Try
                            Dim rawRes As String = ValutaEspressioneStringa(exprPulita)
                            Dim dt As DateTime
                            If DateTime.TryParseExact(rawRes, "yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, dt) Then
                                risultato = dt
                            ElseIf DateTime.TryParse(rawRes, Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, dt) Then
                                risultato = dt
                            Else
                                risultato = Nothing
                            End If
                        Catch
                            risultato = Nothing
                        End Try

                    Case "booleano"
                        Try
                            Dim rawBool As String = ValutaEspressioneStringa(exprPulita)
                            Dim b As Boolean
                            If Boolean.TryParse(rawBool, b) Then
                                risultato = b
                            ElseIf IsNumeric(rawBool) Then
                                risultato = (Convert.ToInt32(rawBool) <> 0)
                            Else
                                risultato = False
                            End If
                        Catch
                            risultato = False
                        End Try

                    Case Else
                        risultato = exprPulita
                End Select

                risultati(nomeCampo) = risultato

            Catch ex As Exception
                ' non bloccare l'intero processo per un errore su un singolo campo
                Try
                    MDIMessageBox.Show($"[ Campo calcolato] Errore nel calcolo di '{nomeCampo}' : {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)
                Catch
                End Try
                risultati(nomeCampo) = Nothing
            End Try
        Next

        Return risultati
    End Function



    Private Function ValutaEspressioneStringa(expr As String) As String
        If expr Is Nothing Then Return String.Empty
        Dim raw = expr.Trim()

        ' Se l'espressione è vuota o solo virgolette/whitespace -> ritorna stringa vuota
        If String.IsNullOrWhiteSpace(raw) Then Return String.Empty
        If Regex.IsMatch(raw, "^\s*""*\s*""*\s*$") Then Return String.Empty

        ' Se è già una stringa racchiusa tra virgolette, restituisci il contenuto interno
        If raw.StartsWith("""") AndAlso raw.EndsWith("""") AndAlso raw.Length >= 2 Then
            Return raw.Substring(1, raw.Length - 2).Replace("""""", """")
        End If

        ' Tentativo principale: DataColumn expression (lo fai già)
        Try
            Dim dt As New DataTable()
            dt.Columns.Add("Expr", GetType(String), raw)
            Dim row = dt.NewRow()
            dt.Rows.Add(row)
            Dim res = row("Expr")
            Return If(res Is Nothing OrElse Convert.IsDBNull(res), String.Empty, res.ToString())
        Catch ex As Exception
            ' Log locale (non interrompere il flusso)
            Try
                MDIMessageBox.Show($"[StringEval] Errore: {ex.Message}{Environment.NewLine}Expr: {raw}", Me.MdiParent, MessageBoxButtons.OK)
            Catch
            End Try
        End Try

        ' Fallback 1: se sembra un'espressione numerica (solo numeri, spazi, operatori), prova DataTable().Compute
        Try
            Dim numericCandidate = raw.Replace("""", "").Trim()
            If Regex.IsMatch(numericCandidate, "^[0-9\+\-\*\/\.\s\(\)]+$") Then
                Dim computed = New DataTable().Compute(numericCandidate, Nothing)
                Return If(computed Is Nothing OrElse Convert.IsDBNull(computed), String.Empty, computed.ToString())
            End If
        Catch
        End Try

        ' Fallback 2: rimozione di token vuoti generici: sostituisco """" (doppie virgolette vuote) con stringa vuota
        Try
            Dim cleaned = Regex.Replace(raw, """{2,}", """""") ' sostituisce sequenze di doppie virgolette con una sola coppia
            cleaned = cleaned.Trim()
            If String.IsNullOrWhiteSpace(cleaned) Then Return String.Empty

            ' Se dopo pulizia è racchiusa tra virgolette restituisco il contenuto
            If cleaned.StartsWith("""") AndAlso cleaned.EndsWith("""") AndAlso cleaned.Length >= 2 Then
                Return cleaned.Substring(1, cleaned.Length - 2).Replace("""""", """")
            End If

            ' Ultima risorsa: restituisco l'espressione originale (non valutata) per evitare crash a monte
            Return cleaned
        Catch
            Return String.Empty
        End Try
    End Function

    Private Sub ToggleUIForSaving(saving As Boolean)
        Me.BeginInvoke(New MethodInvoker(Sub()
                                             For Each c As Control In panelBottoni.Controls
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

    Private Sub EsportaTabella(sender As Object, e As EventArgs)
        Using chooser As New ExportChoiceForm()
            Dim dr = chooser.ShowDialog(Me)
            If dr = DialogResult.OK Then
                Select Case chooser.SelectedExportType
                    Case ExportType.PDF
                        EsportaPDF()
                    Case ExportType.Excel
                        EsportaExcel()
                    Case Else

                End Select
            End If
        End Using
    End Sub

    Private Sub ReleaseComObject(ByVal obj As Object)
        Try
            If obj IsNot Nothing Then Marshal.ReleaseComObject(obj)
        Catch

        Finally
            obj = Nothing
        End Try
    End Sub

    Public Sub EsportaExcel()

        Using sfd As New SaveFileDialog()
            sfd.Filter = "Excel Workbook|*.xlsx|Excel 97-2003|*.xls"
            sfd.FileName = Me.Name
            sfd.Title = "Salva esportazione Excel"
            If sfd.ShowDialog() <> DialogResult.OK Then
                Dim filePath = sfd.FileName

                Dim previousUseWait = Application.UseWaitCursor
                Dim previousCursor = Me.Cursor

                Try
                    Application.UseWaitCursor = True
                    Me.Cursor = Cursors.WaitCursor
                    Application.DoEvents()

                    Task.Run(Sub()
                                 Dim dgv = Me.dgvDati
                                 If dgv Is Nothing Then
                                     Me.BeginInvoke(New MethodInvoker(Sub()
                                                                          MessageBox.Show("DataGridView non trovata.", "Esporta Excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                                                      End Sub))
                                     Return
                                 End If

                                 Dim colIndexes As New List(Of Integer)
                                 Dim colHeaders As New List(Of String)
                                 For i As Integer = 0 To dgv.Columns.Count - 1
                                     If dgv.Columns(i).Visible Then
                                         colIndexes.Add(i)
                                         colHeaders.Add(dgv.Columns(i).HeaderText)
                                     End If
                                 Next

                                 If colIndexes.Count = 0 Then
                                     Me.BeginInvoke(New MethodInvoker(Sub()
                                                                          MessageBox.Show("Nessuna colonna visibile da esportare.", "Esporta Excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                                                      End Sub))
                                     Return
                                 End If

                                 Dim dtSource As New DataTable()
                                 Try
                                     Using conn As New SqlConnection(ConnString)
                                         conn.Open()
                                         Dim query = $"SELECT * FROM [{Me.Name}]" & If(String.IsNullOrWhiteSpace(FiltroIniziale), "", $" WHERE {FiltroIniziale}")
                                         Using cmd As New SqlCommand(query, conn)
                                             cmd.CommandTimeout = 120
                                             Using da As New SqlDataAdapter(cmd)
                                                 da.Fill(dtSource)
                                             End Using
                                         End Using
                                     End Using
                                 Catch ex As Exception
                                     Me.BeginInvoke(New MethodInvoker(Sub()
                                                                          MessageBox.Show("Errore caricamento dati per esportazione: " & ex.Message, "Esporta Excel", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                                      End Sub))
                                     Return
                                 End Try

                                 If dtSource.Rows.Count = 0 Then
                                     Me.BeginInvoke(New MethodInvoker(Sub()
                                                                          MessageBox.Show("Nessun record da esportare.", "Esporta Excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                                                      End Sub))
                                     Return
                                 End If

                                 Dim excelApp As Excel.Application = Nothing
                                 Dim workBook As Excel.Workbook = Nothing
                                 Dim sheet As Excel.Worksheet = Nothing

                                 Try
                                     excelApp = New Excel.Application()
                                     workBook = excelApp.Workbooks.Add()
                                     sheet = CType(workBook.Sheets(1), Excel.Worksheet)

                                     For c As Integer = 0 To colHeaders.Count - 1
                                         sheet.Cells(1, c + 1) = colHeaders(c)
                                     Next

                                     Dim outRow As Integer = 2
                                     For Each dr As DataRow In dtSource.Rows
                                         For c As Integer = 0 To colIndexes.Count - 1
                                             Dim colIndexInDgv = colIndexes(c)
                                             Dim colName = dgv.Columns(colIndexInDgv).Name
                                             Dim val As Object = If(dtSource.Columns.Contains(colName), dr(colName), DBNull.Value)
                                             sheet.Cells(outRow, c + 1) = If(val Is Nothing OrElse Convert.IsDBNull(val), String.Empty, val.ToString())
                                         Next
                                         outRow += 1
                                     Next

                                     sheet.Columns.AutoFit()

                                     workBook.SaveAs(filePath)
                                     workBook.Close(False)
                                     excelApp.Quit()

                                     Me.BeginInvoke(New MethodInvoker(Sub()
                                                                          MessageBox.Show("Esportazione completata: " & filePath, "Esporta Excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                                                      End Sub))
                                 Catch ex As Exception
                                     Me.BeginInvoke(New MethodInvoker(Sub()
                                                                          MessageBox.Show("Errore durante l'esportazione Excel: " & ex.Message, "Esporta Excel", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                                      End Sub))
                                 Finally
                                     Try
                                         If sheet IsNot Nothing Then Marshal.ReleaseComObject(sheet)
                                         If workBook IsNot Nothing Then Marshal.ReleaseComObject(workBook)
                                         If excelApp IsNot Nothing Then Marshal.ReleaseComObject(excelApp)
                                     Catch
                                     Finally
                                         sheet = Nothing
                                         workBook = Nothing
                                         excelApp = Nothing
                                         GC.Collect()
                                         GC.WaitForPendingFinalizers()
                                     End Try
                                 End Try
                             End Sub).Wait()
                Finally
                    Application.UseWaitCursor = previousUseWait
                    Me.Cursor = previousCursor
                    Application.DoEvents()
                End Try
            End If
        End Using
    End Sub

    Public Sub EsportaPDF()
        Dim defaultName As String = Me.Name & "_Esportazione.pdf"
        For Each ch As Char In System.IO.Path.GetInvalidFileNameChars()
            defaultName = defaultName.Replace(ch, "_"c)
        Next

        Using sfd As New SaveFileDialog()
            sfd.Filter = "PDF File|*.pdf"
            sfd.Title = "Salva PDF esportazione"
            sfd.FileName = defaultName
            sfd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            If sfd.ShowDialog() <> DialogResult.OK Then
                Dim filePath = sfd.FileName

                Dim previousUseWait = Application.UseWaitCursor
                Dim previousCursor = Me.Cursor

                Try
                    Application.UseWaitCursor = True
                    Me.Cursor = Cursors.WaitCursor
                    Application.DoEvents()

                    Task.Run(Sub()
                                 Dim dtSource As New DataTable()
                                 Try
                                     Using conn As New SqlConnection(ConnString)
                                         conn.Open()
                                         Dim query = $"SELECT * FROM [{Me.Name}]" & If(String.IsNullOrWhiteSpace(FiltroIniziale), "", $" WHERE {FiltroIniziale}")
                                         Using cmd As New SqlCommand(query, conn)
                                             cmd.CommandTimeout = 120
                                             Using da As New SqlDataAdapter(cmd)
                                                 da.Fill(dtSource)
                                             End Using
                                         End Using
                                     End Using
                                 Catch ex As Exception
                                     Me.BeginInvoke(New MethodInvoker(Sub()
                                                                          MDIMessageBox.Show("Errore caricamento dati per esportazione: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                                      End Sub))
                                     Return
                                 End Try

                                 If dtSource.Rows.Count = 0 Then
                                     Me.BeginInvoke(New MethodInvoker(Sub()
                                                                          MDIMessageBox.Show("Nessun record da esportare.", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Information)
                                                                      End Sub))
                                     Return
                                 End If

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
                                     Dim lineHeight As Double = 16
                                     Dim pageHeight As Double = page.Height.Point
                                     Dim usableWidth As Double = page.Width.Point - (2 * margin)

                                     Dim colonne = dgvDati.Columns.Cast(Of DataGridViewColumn).Where(Function(c) c.Visible).ToList()
                                     Dim colCount = colonne.Count
                                     If colCount = 0 Then
                                         Me.BeginInvoke(New MethodInvoker(Sub()
                                                                              MDIMessageBox.Show("Nessuna colonna visibile da esportare.", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Information)
                                                                          End Sub))
                                         Return
                                     End If

                                     Dim colWidth As Double = usableWidth / colCount

                                     gfx.DrawString($"Esportazione dati: {Me.Name}", New XFont("Arial", 11, XFontStyleEx.Bold), XBrushes.Black, New XPoint(margin, topOffset - 30))
                                     For i = 0 To colCount - 1
                                         Dim header = colonne(i).HeaderText
                                         Dim rect As New XRect(margin + (i * colWidth), topOffset, colWidth, lineHeight)
                                         formatter.DrawString(header, fontBold, XBrushes.DarkBlue, rect, XStringFormats.TopLeft)
                                     Next
                                     Dim currentY As Double = topOffset + lineHeight

                                     Dim approxCharWidth As Double = gfx.MeasureString("W", font).Width

                                     For r As Integer = 0 To dtSource.Rows.Count - 1
                                         Dim dr As DataRow = dtSource.Rows(r)

                                         Dim rowHeight As Double = lineHeight
                                         For i = 0 To colCount - 1
                                             Dim colName = colonne(i).Name
                                             Dim cellVal As String = If(dtSource.Columns.Contains(colName) AndAlso dr(colName) IsNot DBNull.Value, dr(colName).ToString(), String.Empty)
                                             If String.IsNullOrEmpty(cellVal) Then Continue For
                                             Dim size = gfx.MeasureString(cellVal, font)
                                             Dim linesNeeded As Integer = CInt(Math.Ceiling(size.Width / colWidth))
                                             If linesNeeded < 1 Then linesNeeded = 1
                                             Dim neededHeight = linesNeeded * (font.Size + 2)
                                             If neededHeight > rowHeight Then rowHeight = neededHeight
                                         Next

                                         If currentY + rowHeight > pageHeight - margin Then
                                             page = document.AddPage()
                                             page.Orientation = PageOrientation.Landscape
                                             gfx = XGraphics.FromPdfPage(page)
                                             formatter = New XTextFormatter(gfx)
                                             currentY = margin
                                             For i = 0 To colCount - 1
                                                 Dim header = colonne(i).HeaderText
                                                 Dim rectHeader As New XRect(margin + (i * colWidth), currentY, colWidth, lineHeight)
                                                 formatter.DrawString(header, fontBold, XBrushes.DarkBlue, rectHeader, XStringFormats.TopLeft)
                                             Next
                                             currentY += lineHeight
                                         End If

                                         For i = 0 To colCount - 1
                                             Dim colName = colonne(i).Name
                                             Dim valore As String = If(dtSource.Columns.Contains(colName) AndAlso dr(colName) IsNot DBNull.Value, dr(colName).ToString(), String.Empty)
                                             Dim rect As New XRect(margin + (i * colWidth), currentY, colWidth, rowHeight)
                                             If (r Mod 2) = 0 Then gfx.DrawRectangle(XBrushes.LightGray, rect)
                                             formatter.DrawString(valore, font, XBrushes.Black, rect, XStringFormats.TopLeft)
                                         Next

                                         currentY += rowHeight
                                     Next

                                     document.Save(filePath)

                                     Me.BeginInvoke(New MethodInvoker(Sub()
                                                                          MDIMessageBox.Show($"PDF esportato con successo:{Environment.NewLine}{filePath}", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Information)
                                                                      End Sub))
                                 Catch ex As Exception
                                     Me.BeginInvoke(New MethodInvoker(Sub()
                                                                          MDIMessageBox.Show("Errore durante l'esportazione PDF: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                                      End Sub))
                                 End Try
                             End Sub).Wait()
                Finally
                    Application.UseWaitCursor = previousUseWait
                    Me.Cursor = previousCursor
                    Application.DoEvents()
                End Try
            End If
        End Using
    End Sub

    Private Sub CaricaDatiNeiControlli(riga As DataGridViewRow)
        If riga Is Nothing Then Return

        Try
            isUpdatingControls = True

            Dim dtCollegamenti As DataTable = EseguiQuery($"SELECT NomeCampo FROM Sys_CampiCollegamento WHERE NomeTabella = '{nomeTabellaCorrente}'")
            Dim campiCollegati As New HashSet(Of String)(dtCollegamenti.AsEnumerable().Select(Function(r) r("NomeCampo").ToString()), StringComparer.OrdinalIgnoreCase)

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
                                    MDIMessageBox.Show($"Errore recupero descrizione per campo '{campoNome}': {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)
                                End Try
                            Else
                                lbl.Text = "..."
                            End If
                        End If

                    Case Else
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
            If combo.DataSource Is Nothing OrElse String.IsNullOrWhiteSpace(combo.ValueMember) Then
                combo.SelectedIndex = -1
                Return
            End If

            If valore Is Nothing OrElse Convert.IsDBNull(valore) Then
                combo.SelectedIndex = -1
                Return
            End If

            Dim stringVal As String = valore.ToString()

            Dim targetType As Type = Nothing
            Dim dt As DataTable = TryCast(TryCast(combo.DataSource, DataView)?.Table, DataTable)
            If dt Is Nothing Then
                dt = TryCast(combo.DataSource, DataTable)
            End If

            If dt IsNot Nothing AndAlso dt.Columns.Contains(combo.ValueMember) Then
                targetType = dt.Columns(combo.ValueMember).DataType
            End If

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

            If targetType IsNot Nothing Then
                Try
                    Dim converted = Convert.ChangeType(stringVal, targetType, Globalization.CultureInfo.InvariantCulture)
                    combo.SelectedValue = converted
                    Return
                Catch
                    Try
                        combo.SelectedValue = stringVal
                        Return
                    Catch
                    End Try
                End Try
            Else
                combo.SelectedValue = stringVal
            End If

        Catch ex As Exception
            MDIMessageBox.Show($"ImpostaValoreCombo error: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)
            Try
                combo.SelectedIndex = -1
            Catch
            End Try
        End Try
    End Sub

    Private Sub UpdateButtonsByModalita()
        Dim modo = If(String.IsNullOrWhiteSpace(ModalitaCorrente), "nessuna", ModalitaCorrente.ToLowerInvariant())

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
            Case Else
                salvaEnabled = False
                annullaEnabled = False
                canInsert = True
                canEdit = True
                canDelete = True
        End Select

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
            MDIMessageBox.Show($"TrovaDefinizioneCampo: lista campiDefiniti vuota; richiesta per '{nomeCampo}'", Me.MdiParent, MessageBoxButtons.OK)
            Return Nothing
        End If

        For Each campo As CampoDatabase In campiDefiniti
            If campo IsNot Nothing AndAlso
           Not String.IsNullOrWhiteSpace(campo.Nome) AndAlso
           campo.Nome.Trim().Equals(nomeCampo.Trim(), StringComparison.OrdinalIgnoreCase) Then
                Return campo
            End If
        Next

        MDIMessageBox.Show($"TrovaDefinizioneCampo: campo '{nomeCampo}' non trovato nella definizione dei campi per la tabella {Me.Name}", Me.MdiParent, MessageBoxButtons.OK)
        Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show($"Il campo '{nomeCampo}' non è stato trovato nella definizione dei campi.", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Warning)))
        Return Nothing
    End Function

    Private Function RecuperaCampiPath() As HashSet(Of String)
        Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Try
            Using conn As New SqlConnection(ConnString)
                Dim query As String = "SELECT NomeCampoPath FROM Sys_CampiPath"
                Using cmd As New SqlCommand(query, conn)
                    conn.Open()
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim nome As String = If(reader.IsDBNull(0), String.Empty, reader.GetString(0))
                            If Not String.IsNullOrWhiteSpace(nome) Then
                                result.Add(nome.Trim())
                            End If
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MDIMessageBox.Show($"RecuperaCampiPath errore: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)
        End Try
        Return result
    End Function

    Private Function IsCampoPath(nomeCampo As String) As Boolean
        If String.IsNullOrWhiteSpace(nomeCampo) Then Return False
        Return campiPath.Contains(nomeCampo)
    End Function

    Private Sub AggiungiBottone(nome As String, handler As EventHandler)
        Dim btn As New Button With {.Text = nome, .AutoSize = True}
        AddHandler btn.Click, handler

        panelBottoni.Controls.Add(btn)
    End Sub

    Private Sub UniformaDimensioniBottoni()
        Dim larghezzaMassima As Integer = 0
        Dim altezzaMassima As Integer = 0
        Dim tuttiPannelli As New List(Of Control) From {panelBottoni, panelBottoniDinamici}

        For Each pnl In tuttiPannelli
            If pnl Is Nothing Then Continue For
            For Each ctrl As Control In pnl.Controls
                If TypeOf ctrl Is Button Then
                    Dim btn As Button = CType(ctrl, Button)
                    Dim pref = btn.PreferredSize
                    If pref.Width > larghezzaMassima Then larghezzaMassima = pref.Width
                    If pref.Height > altezzaMassima Then altezzaMassima = pref.Height
                End If
            Next
        Next

        If larghezzaMassima = 0 OrElse altezzaMassima = 0 Then Return

        For Each pnl In tuttiPannelli
            If pnl Is Nothing Then Continue For
            For Each ctrl As Control In pnl.Controls
                If TypeOf ctrl Is Button Then
                    Dim btn As Button = CType(ctrl, Button)
                    btn.AutoSize = False
                    btn.Width = larghezzaMassima
                    btn.Height = altezzaMassima
                End If
            Next
        Next
    End Sub

    Private Function HasColumn(reader As SqlDataReader, columnName As String) As Boolean
        If reader Is Nothing OrElse String.IsNullOrWhiteSpace(columnName) Then Return False
        For i As Integer = 0 To reader.FieldCount - 1
            If String.Equals(reader.GetName(i), columnName, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Function CreateFormInstanceByName(nomeForm As String) As Form
        If String.IsNullOrWhiteSpace(nomeForm) Then Return Nothing

        For Each asm In AppDomain.CurrentDomain.GetAssemblies()
            Try
                Dim t = asm.GetTypes().FirstOrDefault(Function(tt) String.Equals(tt.Name, nomeForm, StringComparison.OrdinalIgnoreCase))
                If t IsNot Nothing AndAlso GetType(Form).IsAssignableFrom(t) Then
                    Try
                        Dim obj = Activator.CreateInstance(t)
                        Return TryCast(obj, Form)
                    Catch

                    End Try
                End If
            Catch

            End Try
        Next

        Return Nothing
    End Function

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
                            Dim bottoneDinamico As String = If(HasColumn(reader, "BottoneDinamico") AndAlso Not reader.IsDBNull(reader.GetOrdinal("BottoneDinamico")), reader("BottoneDinamico").ToString(), String.Empty)
                            Dim buttonText As String = If(HasColumn(reader, "ButtonText") AndAlso Not reader.IsDBNull(reader.GetOrdinal("ButtonText")), reader("ButtonText").ToString(), String.Empty)
                            Dim campoChiavePadre As String = If(HasColumn(reader, "CampoChiavePadre") AndAlso Not reader.IsDBNull(reader.GetOrdinal("CampoChiavePadre")), reader("CampoChiavePadre").ToString(), String.Empty)
                            Dim campoChiaveFiglia As String = If(HasColumn(reader, "CampoChiaveFiglia") AndAlso Not reader.IsDBNull(reader.GetOrdinal("CampoChiaveFiglia")), reader("CampoChiaveFiglia").ToString(), String.Empty)
                            Dim formDaAprire As String = If(HasColumn(reader, "FormDaAprire") AndAlso Not reader.IsDBNull(reader.GetOrdinal("FormDaAprire")), reader("FormDaAprire").ToString(), String.Empty)
                            Dim modulo As String = If(HasColumn(reader, "Modulo") AndAlso Not reader.IsDBNull(reader.GetOrdinal("Modulo")), reader("Modulo").ToString(), String.Empty)

                            Dim displayText = If(String.IsNullOrWhiteSpace(buttonText),
                                                 If(String.IsNullOrWhiteSpace(bottoneDinamico), "Apri", bottoneDinamico),
                                                 buttonText)

                            Dim btnDinamico As New Button With {
                                .Text = displayText,
                                .AutoSize = True,
                                .Margin = New Padding(5),
                                .Tag = New With {
                                    Key .FormName = bottoneDinamico,
                                    Key .CampoPadre = campoChiavePadre,
                                    Key .CampoFiglia = campoChiaveFiglia,
                                    Key .Titolo = displayText,
                                    Key .FormDaAprire = formDaAprire,
                                    Key .Modulo = modulo
                                }
                            }

                            AddHandler btnDinamico.Click, Sub(s, e)
                                                              Try
                                                                  Dim info = CType(CType(s, Button).Tag, Object)
                                                                  Dim nomeFormDaAprire = If(info.FormDaAprire, "").ToString().Trim()
                                                                  If Not String.IsNullOrWhiteSpace(nomeFormDaAprire) Then
                                                                      If dgvDati Is Nothing OrElse dgvDati.SelectedRows.Count = 0 Then
                                                                          MDIMessageBox.Show("Seleziona prima una riga dalla griglia.", Me.MdiParent, MessageBoxButtons.OK)
                                                                          Return
                                                                      End If
                                                                      Dim selRow As DataGridViewRow = dgvDati.SelectedRows(0)
                                                                      Dim recordValues As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                                                                      For Each col As DataGridViewColumn In dgvDati.Columns
                                                                          Try
                                                                              Dim val = selRow.Cells(col.Index).Value
                                                                              recordValues(col.Name) = If(val Is Nothing OrElse Convert.IsDBNull(val), Nothing, val)
                                                                          Catch
                                                                              recordValues(col.Name) = Nothing
                                                                          End Try
                                                                      Next

                                                                      Dim targetForm As Form = Nothing
                                                                      For Each f As Form In GesPu25.MdiChildren
                                                                          If String.Equals(f.Name, nomeFormDaAprire, StringComparison.OrdinalIgnoreCase) _
                                                                                                OrElse String.Equals(f.GetType().Name, nomeFormDaAprire, StringComparison.OrdinalIgnoreCase) Then
                                                                              targetForm = f
                                                                              Exit For
                                                                          End If
                                                                      Next

                                                                      If targetForm Is Nothing Then
                                                                          Try
                                                                              targetForm = CreateFormInstanceByName(nomeFormDaAprire)
                                                                          Catch exCreate As Exception
                                                                              targetForm = Nothing
                                                                          End Try
                                                                      End If

                                                                      If targetForm Is Nothing Then
                                                                          MDIMessageBox.Show($"Impossibile ottenere o creare il form '{nomeFormDaAprire}'. Verifica che la classe esista e sia accessibile.", Me.MdiParent, MessageBoxButtons.OK)
                                                                          Return
                                                                      End If

                                                                      Try
                                                                          Dim existingTag = targetForm.Tag
                                                                          Try
                                                                              targetForm.Tag = New With {
                                                                                Key .Existing = existingTag,
                                                                                Key .TableName = Me.Name,
                                                                                Key .Record = recordValues
            }
                                                                          Catch
                                                                              targetForm.Tag = New With {
                                                                                Key .TableName = Me.Name,
                                                                                Key .Record = recordValues
            }
                                                                          End Try
                                                                      Catch
                                                                          targetForm.Tag = New With {
                                                                                Key .TableName = Me.Name,
                                                                                Key .Record = recordValues
        }
                                                                      End Try

                                                                      Try
                                                                          GesPu25.ApriModulo2ConPermessi(nomeFormDaAprire, targetForm)
                                                                      Catch ex As Exception
                                                                          MDIMessageBox.Show($"Errore aprendo il modulo '{nomeFormDaAprire}' tramite ApriModulo2ConPermessi: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)
                                                                          Return
                                                                      End Try

                                                                      Try
                                                                          If targetForm.WindowState = FormWindowState.Minimized Then targetForm.WindowState = FormWindowState.Normal
                                                                          targetForm.BringToFront()
                                                                          targetForm.Activate()
                                                                      Catch
                                                                      End Try

                                                                      Return
                                                                  End If

                                                                  If dgvDati Is Nothing OrElse dgvDati.SelectedRows.Count = 0 Then
                                                                      MDIMessageBox.Show("Seleziona prima una riga dalla griglia per aprire il form collegato.", Me.MdiParent, MessageBoxButtons.OK)
                                                                      Return
                                                                  End If

                                                                  Dim valoreCellaObj = dgvDati.SelectedRows(0).Cells(info.CampoPadre).Value
                                                                  Dim valoreCella As String = If(valoreCellaObj Is Nothing OrElse Convert.IsDBNull(valoreCellaObj), String.Empty, valoreCellaObj.ToString())
                                                                  Dim valoreChiavePadre = If(Not String.IsNullOrEmpty(valoreCella) AndAlso valoreCella.Contains("-"), valoreCella.Split("-"c)(0).Trim(), valoreCella)

                                                                  If String.IsNullOrWhiteSpace(valoreChiavePadre) Then
                                                                      MDIMessageBox.Show("Il valore della chiave primaria selezionata è nullo o non valido.", Me.MdiParent, MessageBoxButtons.OK)
                                                                      Return
                                                                  End If

                                                                  Dim nomeFormTarget = If(info.FormName, "").ToString().Trim()
                                                                  If String.IsNullOrWhiteSpace(nomeFormTarget) Then
                                                                      MDIMessageBox.Show("Nessun form target definito per il bottone dinamico.", Me.MdiParent, MessageBoxButtons.OK)
                                                                      Return
                                                                  End If

                                                                  For Each f As Form In GesPu25.MdiChildren
                                                                      If TypeOf f Is DynamicDataForm AndAlso String.Equals(f.Name, nomeFormTarget, StringComparison.OrdinalIgnoreCase) Then
                                                                          f.WindowState = FormWindowState.Normal
                                                                          f.BringToFront()
                                                                          f.Activate()
                                                                          Return
                                                                      End If
                                                                  Next

                                                                  Dim campiFigli = RecuperaCampiDa(nomeFormTarget)
                                                                  Dim nuovoForm As New DynamicDataForm(campiFigli, nomeFormTarget)
                                                                  nuovoForm.MdiParent = GesPu25
                                                                  nuovoForm.Text = $"{info.Titolo} - Filtrato per {info.CampoPadre} = {valoreChiavePadre}"
                                                                  nuovoForm.FiltroIniziale = $"{info.CampoFiglia} = '{valoreChiavePadre}'"
                                                                  GesPu25.ApriModulo2ConPermessi(nomeFormTarget, nuovoForm)

                                                                  Dim campoCollegamento = info.CampoFiglia
                                                                  If nuovoForm.campoInputs.ContainsKey(campoCollegamento) Then
                                                                      Dim targetCtrl = nuovoForm.campoInputs(campoCollegamento)
                                                                      Select Case True
                                                                          Case TypeOf targetCtrl Is ComboBox
                                                                              CType(targetCtrl, ComboBox).SelectedValue = valoreChiavePadre
                                                                          Case TypeOf targetCtrl Is TextBox
                                                                              CType(targetCtrl, TextBox).Text = valoreChiavePadre
                                                                          Case TypeOf targetCtrl Is FlowLayoutPanel
                                                                              Dim txt = targetCtrl.Controls.OfType(Of TextBox)().FirstOrDefault()
                                                                              If txt IsNot Nothing Then txt.Text = valoreChiavePadre
                                                                      End Select
                                                                  End If
                                                              Catch ex As Exception
                                                                  MDIMessageBox.Show("Errore esecuzione bottone dinamico: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
                                                              End Try
                                                          End Sub

                            panelBottoniDinamici.Controls.Add(btnDinamico)
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Me.BeginInvoke(New MethodInvoker(Sub() MDIMessageBox.Show("Errore nel caricamento bottoni dinamici: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)))
        End Try
    End Sub

    Private Function ParseInterval(expr As String) As (startVal As Long, endVal As Long, stepVal As Long, padLen As Integer)?
        If String.IsNullOrWhiteSpace(expr) Then Return Nothing

        Dim rx As New Regex("^@intervallo\(\s*([+-]?\d+)\s*-\s*([+-]?\d+)(?:\s+step\s+([+-]?\d+))?\s*\)$", RegexOptions.IgnoreCase)
        Dim m = rx.Match(expr.Trim())
        If Not m.Success Then Return Nothing

        Dim sStr = m.Groups(1).Value
        Dim eStr = m.Groups(2).Value
        Dim stepStr = If(m.Groups(3).Success, m.Groups(3).Value, "1")

        Dim padLen As Integer = Math.Max(sStr.TrimStart("+"c, "-"c).Length, eStr.TrimStart("+"c, "-"c).Length)

        Dim startVal As Long
        Dim endVal As Long
        Dim stepVal As Long

        If Not Int64.TryParse(sStr, startVal) Then Return Nothing
        If Not Int64.TryParse(eStr, endVal) Then Return Nothing
        If Not Int64.TryParse(stepStr, stepVal) Then Return Nothing
        If stepVal = 0 Then Return Nothing

        Return (startVal, endVal, stepVal, padLen)
    End Function

    Private Function ParsePrefixedRange(expr As String) As (prefix As String, startVal As Long, endVal As Long, padLen As Integer)?
        If String.IsNullOrWhiteSpace(expr) Then Return Nothing

        ' Pattern: QUALSIASI_TESTO<da>0004<a>0015
        Dim rx As New Regex("^(?<pref>.*)<da>(?<start>\d+)<a>(?<end>\d+)$", RegexOptions.IgnoreCase)

        Dim m = rx.Match(expr.Trim())
        If Not m.Success Then Return Nothing

        Dim pref = m.Groups("pref").Value
        Dim sStr = m.Groups("start").Value
        Dim eStr = m.Groups("end").Value

        Dim startVal As Long
        Dim endVal As Long

        If Not Int64.TryParse(sStr, startVal) Then Return Nothing
        If Not Int64.TryParse(eStr, endVal) Then Return Nothing

        Dim padLen As Integer = Math.Max(sStr.Length, eStr.Length)

        Return (pref, startVal, endVal, padLen)
    End Function


    Private Function GetColumnSqlDbType(conn As SqlConnection, tableName As String, columnName As String) As SqlDbType
        Dim schemaName As String = "dbo"
        Dim tn = tableName
        Dim col = columnName

        Dim sql = "SELECT t.name AS TypeName FROM sys.columns c " &
              "JOIN sys.types t ON c.user_type_id = t.user_type_id " &
              "JOIN sys.tables tb ON c.object_id = tb.object_id " &
              "WHERE tb.name = @table AND c.name = @col"

        Using cmd As New SqlCommand(sql, conn)
            cmd.Parameters.AddWithValue("@table", tn)
            cmd.Parameters.AddWithValue("@col", col)
            Dim res = cmd.ExecuteScalar()
            If res Is Nothing Then Return SqlDbType.NVarChar
            Dim typeName = Convert.ToString(res).ToLowerInvariant()
            Select Case typeName
                Case "int" : Return SqlDbType.Int
                Case "bigint" : Return SqlDbType.BigInt
                Case "smallint" : Return SqlDbType.SmallInt
                Case "tinyint" : Return SqlDbType.TinyInt
                Case "decimal", "numeric" : Return SqlDbType.Decimal
                Case "float" : Return SqlDbType.Float
                Case "real" : Return SqlDbType.Real
                Case "bit" : Return SqlDbType.Bit
                Case "datetime", "smalldatetime", "datetime2" : Return SqlDbType.DateTime
                Case "date" : Return SqlDbType.Date
                Case "time" : Return SqlDbType.Time
                Case Else : Return SqlDbType.NVarChar
            End Select
        End Using
    End Function

    ' originalFields: Dictionary(Of String,Object) con tutti i campi e i valori del record originale
    ' fieldName: il campo che può contenere @intervallo(...)
    ' connString: stringa di connessione
    ' tableName: nome tabella DB (es. "FormDinamicoTable" o la tua tabella)
    Private Sub SaveRecordsFromInterval(originalFields As Dictionary(Of String, Object),
                                    fieldName As String,
                                    connString As String,
                                    tableName As String)
        Dim rawValueObj As Object = Nothing
        If Not originalFields.TryGetValue(fieldName, rawValueObj) Then
            Throw New ArgumentException("Campo non trovato in originalFields")
        End If

        Dim rawValue = Convert.ToString(rawValueObj)

        ' 1) Provo prima il nuovo pattern con prefisso: SC_005_<da>0004<a>0015
        Dim prefixed = ParsePrefixedRange(rawValue)
        Dim usePrefixed As Boolean = prefixed.HasValue

        Dim startVal As Long
        Dim endVal As Long
        Dim stepVal As Long = 1 ' default
        Dim padLen As Integer
        Dim prefix As String = String.Empty

        If usePrefixed Then
            prefix = prefixed.Value.prefix
            startVal = prefixed.Value.startVal
            endVal = prefixed.Value.endVal
            padLen = prefixed.Value.padLen
        Else
            ' fallback: vecchia sintassi @intervallo(...)
            Dim parsed = ParseInterval(rawValue)
            If parsed Is Nothing Then
                ' Non è un intervallo: salva un singolo record
                SaveSingleRecord(originalFields, connString, tableName)
                Return
            End If

            startVal = parsed.Value.startVal
            endVal = parsed.Value.endVal
            stepVal = parsed.Value.stepVal
            padLen = parsed.Value.padLen
        End If


        ' Calcola iterazioni e protezione
        Dim maxIterations As Integer = 10000
        Dim iterations As Long = 0
        If (stepVal > 0 AndAlso startVal <= endVal) Then
            iterations = ((endVal - startVal) \ stepVal) + 1
        ElseIf (stepVal < 0 AndAlso startVal >= endVal) Then
            iterations = ((startVal - endVal) \ Math.Abs(stepVal)) + 1
        Else
            Throw New InvalidOperationException("Intervallo e step incompatibili (nessuna iterazione).")
        End If

        If iterations <= 0 OrElse iterations > maxIterations Then
            Throw New InvalidOperationException($"Numero di iterazioni non valido o troppo grande: {iterations}. Limite: {maxIterations}.")
        End If

        ' Prepara INSERT dinamico
        Dim columns = originalFields.Keys.ToList()
        Dim insertCols As New List(Of String)
        Dim paramNames As New List(Of String)
        For Each col In columns
            insertCols.Add("[" & col & "]")
            paramNames.Add("@" & col)
        Next
        Dim insertSql = $"INSERT INTO {tableName} ({String.Join(", ", insertCols)}) VALUES ({String.Join(", ", paramNames)})"

        Using conn As New SqlConnection(connString)
            conn.Open()
            Using tran = conn.BeginTransaction()
                Try
                    ' Determina il tipo SQL del campo variabile
                    Dim fieldSqlType = GetColumnSqlDbType(conn, tableName, fieldName)

                    Using cmd As New SqlCommand(insertSql, conn, tran)
                        cmd.Parameters.Clear()
                        ' crea parametri con valori iniziali
                        For Each col In columns
                            Dim param = cmd.Parameters.Add("@" & col, SqlDbType.NVarChar)
                            Dim val = originalFields(col)
                            If val Is Nothing Then
                                param.Value = DBNull.Value
                            Else
                                param.Value = val
                            End If
                        Next

                        Dim current = startVal
                        For i As Long = 1 To iterations
                            ' formatta il valore con padding
                            Dim formattedCore As String = If(padLen > 0, current.ToString("D" & padLen), current.ToString())
                            Dim formattedStr As String

                            If usePrefixed Then
                                ' Nuovo comportamento: SC_005_0004 ... SC_005_0015
                                formattedStr = prefix & formattedCore
                            Else
                                ' Vecchio comportamento numerico puro
                                formattedStr = formattedCore
                            End If

                            ' imposta parametro con tipo corretto
                            Dim p = cmd.Parameters("@" & fieldName)
                            If fieldSqlType = SqlDbType.Int OrElse fieldSqlType = SqlDbType.BigInt OrElse fieldSqlType = SqlDbType.SmallInt OrElse fieldSqlType = SqlDbType.TinyInt Then
                                Dim numericVal As Long
                                If Long.TryParse(formattedStr, numericVal) Then
                                    p.SqlDbType = fieldSqlType
                                    p.Value = numericVal
                                Else
                                    ' fallback: salva come stringa
                                    p.SqlDbType = SqlDbType.NVarChar
                                    p.Value = formattedStr
                                End If
                            ElseIf fieldSqlType = SqlDbType.Decimal OrElse fieldSqlType = SqlDbType.Float OrElse fieldSqlType = SqlDbType.Real Then
                                Dim dbl As Double
                                If Double.TryParse(formattedStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, dbl) Then
                                    p.SqlDbType = fieldSqlType
                                    p.Value = dbl
                                Else
                                    p.SqlDbType = SqlDbType.NVarChar
                                    p.Value = formattedStr
                                End If
                            Else
                                p.SqlDbType = SqlDbType.NVarChar
                                p.Value = formattedStr
                            End If

                            ' esegui insert
                            cmd.ExecuteNonQuery()
                            current = current + stepVal
                        Next
                    End Using

                    tran.Commit()
                Catch ex As Exception
                    Try
                        tran.Rollback()
                    Catch
                    End Try
                    Throw
                End Try
            End Using
        End Using
    End Sub

    Private Sub SaveSingleRecord(fields As Dictionary(Of String, Object), connString As String, tableName As String)
        Dim columns = fields.Keys.ToList()
        Dim insertCols = columns.Select(Function(c) "[" & c & "]").ToArray()
        Dim paramNames = columns.Select(Function(c) "@" & c).ToArray()
        Dim insertSql = $"INSERT INTO {tableName} ({String.Join(", ", insertCols)}) VALUES ({String.Join(", ", paramNames)})"

        Using conn As New SqlConnection(connString)
            conn.Open()
            Using cmd As New SqlCommand(insertSql, conn)
                For Each col In columns
                    Dim val = fields(col)
                    If val Is Nothing Then
                        cmd.Parameters.AddWithValue("@" & col, DBNull.Value)
                    Else
                        cmd.Parameters.AddWithValue("@" & col, val)
                    End If
                Next
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

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
        Dim formatoSupportato = {"jpg", "jpeg", "png", "bmp", "gif", "mp4", "avi", "wmv", "mov"}

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
