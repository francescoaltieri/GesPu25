Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports System.Windows.Forms
Imports AxWMPLib
Imports GesPu25.ModuloCampiDinamici
Imports Microsoft.Data.SqlClient
Imports PdfSharp
Imports PdfSharp.Drawing
Imports PdfSharp.Drawing.Layout
Imports PdfSharp.Pdf
Imports WMPLib

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

    Public Sub New(campi As List(Of CampoDatabase), nomeTabella As String)
        Me.Name = nomeTabella
        Me.Text = "Form Dinamico"
        Me.Size = New Size(1100, 600)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.campiDefiniti = RecuperaCampiDa(nomeTabella)
        Me.nomeTabellaCorrente = nomeTabella

        ' Carica stato del form dal modulo condiviso
        GestioneStatoForm.CaricaStato(Me)

        ' Layout principale a due colonne
        Dim layoutPrincipale As New TableLayoutPanel With {
        .Dock = DockStyle.Fill,
        .ColumnCount = 2
    }
        layoutPrincipale.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
        layoutPrincipale.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
        Me.Controls.Add(layoutPrincipale)

        ' Layout sinistro interno con 2 righe: campi + bottoni
        Dim layoutSinistroInterno As New TableLayoutPanel With {
        .Dock = DockStyle.Fill,
        .RowCount = 2,
        .ColumnCount = 1
    }
        layoutSinistroInterno.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        layoutSinistroInterno.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        ' Pannello sinistro con campi
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

        ' Etichetta modalità
        lblModalita = New Label With {
        .Text = "",
        .AutoSize = True,
        .Font = New Font("Verdana", 10, FontStyle.Bold),
        .ForeColor = Color.DarkGreen,
        .Dock = DockStyle.Top,
        .Padding = New Padding(5),
        .TextAlign = ContentAlignment.TopLeft
    }
        pannelloSinistro.Controls.Add(lblModalita)
        pannelloSinistro.SetColumnSpan(lblModalita, 2)

        ' Campi dinamici con righe uniformi
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
            ctrl.Width = 250
            ctrl.Anchor = AnchorStyles.Left
            ctrl.Margin = New Padding(5)
            campoInputs.Add(campi(i).Nome, ctrl)

            pannelloSinistro.Controls.Add(lbl, 0, i + 1)
            pannelloSinistro.Controls.Add(ctrl, 1, i + 1)
        Next

        'CaricaCollegamentiDinamici(nomeTabella)

        ' Bottoni contenitore
        Dim panelBottoniContenitore As New FlowLayoutPanel With {
        .Dock = DockStyle.Fill,
        .FlowDirection = FlowDirection.LeftToRight,
        .AutoSize = True,
        .Padding = New Padding(10),
        .Margin = New Padding(0),
        .WrapContents = True
    }

        ' Bottoni fissi
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

        ' Bottoni dinamici
        panelBottoniDinamici = New FlowLayoutPanel With {
        .FlowDirection = FlowDirection.LeftToRight,
        .AutoSize = True,
        .Margin = New Padding(0, 5, 0, 0)
    }
        panelBottoniContenitore.Controls.Add(panelBottoniDinamici)

        ' Inserimento layout sinistro completo
        layoutSinistroInterno.Controls.Add(pannelloSinistro, 0, 0)
        layoutSinistroInterno.Controls.Add(panelBottoniContenitore, 0, 1)
        layoutPrincipale.Controls.Add(layoutSinistroInterno, 0, 0)

        ' Sezione destra: Griglia dati
        dgvDati = New DataGridView With {
        .Dock = DockStyle.Fill,
        .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
        .AllowUserToAddRows = False,
        .ReadOnly = True
    }
        With dgvDati
            .ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
            .ColumnHeadersHeight = 50
        End With

        AddHandler dgvDati.CellClick, AddressOf dgvDati_CellClick
        AddHandler dgvDati.SelectionChanged, AddressOf dgvDati_SelectionChanged
        layoutPrincipale.Controls.Add(dgvDati, 1, 0)

        ' Caricamenti iniziali
        CaricaBottoniDinamici()
        CaricaDatiTabella(Me.Name)
        ApplicaVisualizzazioneColonne()

        ' Disabilita i controlli inizialmente
        For Each ctrl As Control In campoInputs.Values
            If TypeOf ctrl Is FlowLayoutPanel Then
                For Each innerCtrl As Control In ctrl.Controls
                    If TypeOf innerCtrl Is Button AndAlso CType(innerCtrl, Button).Text = "Visualizza" Then Continue For
                    innerCtrl.Enabled = False
                Next
            Else
                ctrl.Enabled = False
            End If
        Next

        ' Nasconde colonne sensibili
        For Each col As DataGridViewColumn In dgvDati.Columns
            If col.Name.ToLower().Contains("password") Then col.Visible = False
        Next

        UniformaDimensioniBottoni()
        ApplicaAutorizzazioni(NomeUtenteCorrente)
        PulisciCampi()
    End Sub

    Private Sub CaricaCollegamentiDinamici(nomeTabella As String)
        Dim query = "SELECT * FROM Sys_CollegamentiCampi WHERE NomeTabella = @nome"

        Using conn As New SqlConnection(ConnString)
            conn.Open()

            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@nome", nomeTabella)

                Dim da As New SqlDataAdapter(cmd)
                Dim dt As New DataTable()
                da.Fill(dt)

                For Each riga As DataRow In dt.Rows
                    Dim nomeCampo = riga("NomeCampo").ToString()
                    Dim tabellaCollegata = riga("TabellaCollegata").ToString()
                    Dim campoValore = riga("CampoValore").ToString()
                    Dim campoVisuale = riga("CampoVisuale").ToString()

                    If campoInputs.ContainsKey(nomeCampo) AndAlso TypeOf campoInputs(nomeCampo) Is ComboBox Then
                        Dim combo = CType(campoInputs(nomeCampo), ComboBox)

                        ' Nuova query per la tabella collegata
                        Dim queryCollegata = $"SELECT [{campoValore}], [{campoVisuale}] FROM [{tabellaCollegata}]"
                        Using cmdCollegata As New SqlCommand(queryCollegata, conn)
                            Dim daCollegata As New SqlDataAdapter(cmdCollegata)
                            Dim dtCollegata As New DataTable()
                            daCollegata.Fill(dtCollegata)

                            combo.DataSource = dtCollegata
                            combo.ValueMember = campoValore
                            combo.DisplayMember = campoVisuale
                            combo.DropDownStyle = ComboBoxStyle.DropDownList
                        End Using
                    End If
                Next
            End Using
        End Using
    End Sub

    Private Sub AnnullaOperazione()
        Dim risposta = MDIMessageBox.Show("Vuoi annullare l’operazione corrente?", Me.MdiParent, MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If risposta = DialogResult.Yes Then
            ' Reset modalità e interfaccia
            ModalitaCorrente = "nessuna"
            lblModalita.Text = ""

            ' Disattiva controlli input (escluso "Visualizza")
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
            lblModalita.ForeColor = Color.DarkGreen ' Colore neutro
            DisabilitaPulsante("Annulla", True)
            ModalitaCorrente = "nessuna"
            lblModalita.Text = ""
            PulisciCampi()

        End If
    End Sub

    Private Sub AbilitaCampi(abilita As Boolean)
        For Each ctrl As Control In campoInputs.Values
            If TypeOf ctrl Is FlowLayoutPanel Then
                For Each innerCtrl As Control In ctrl.Controls
                    ' Il bottone "Visualizza" deve rimanere sempre attivo
                    If TypeOf innerCtrl Is Button AndAlso CType(innerCtrl, Button).Text = "Visualizza" Then
                        innerCtrl.Enabled = True
                    Else
                        innerCtrl.Enabled = abilita
                    End If
                Next
            Else
                ctrl.Enabled = abilita
            End If
        Next
    End Sub

    Private Sub DisabilitaCampi()
        AbilitaCampi(False)
    End Sub

    Private Function CampoIDGestitoManuale() As Boolean
        Return False ' personalizzabile
    End Function

    Private Sub AggiungiBottone(nome As String, handler As EventHandler)
        Dim btn As New Button With {.Text = nome, .AutoSize = True}
        AddHandler btn.Click, handler
        panelBottoni.Controls.Add(btn)
    End Sub

    Private Sub UniformaDimensioniBottoni()
        Dim larghezzaMassima As Integer = 0
        Dim altezzaMassima As Integer = 0

        ' Calcola la dimensione massima tra tutti i bottoni
        For Each ctrl As Control In panelBottoni.Controls
            If TypeOf ctrl Is Button Then
                Dim btn As Button = CType(ctrl, Button)
                If btn.Width > larghezzaMassima Then larghezzaMassima = btn.Width
                If btn.Height > altezzaMassima Then altezzaMassima = btn.Height
            End If
        Next

        ' Applica la dimensione a tutti i bottoni
        For Each ctrl As Control In panelBottoni.Controls
            If TypeOf ctrl Is Button Then
                Dim btn As Button = CType(ctrl, Button)
                btn.AutoSize = False
                btn.Width = larghezzaMassima
                btn.Height = altezzaMassima
            End If
        Next
    End Sub

    Private Sub InserisciDati(sender As Object, e As EventArgs)

        isModifica = False
        AbilitaCampi(True)
        PulisciCampi()
        DisabilitaPulsante("Salva", False)

        ModalitaCorrente = "inserimento"
        lblModalita.Text = "Inserimento in corso..."
        DisabilitaPulsante("Annulla", False)

        For Each campo In campiDefiniti
            If Not campoInputs.ContainsKey(campo.Nome) Then Continue For
            Dim ctrl = campoInputs(campo.Nome)
            If ctrl Is Nothing Then Continue For

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
                    ' Pulizia dei componenti all'interno del pannello multimediale (ImgVid)
                    For Each innerCtrl As Control In ctrl.Controls
                        If TypeOf innerCtrl Is TextBox Then
                            CType(innerCtrl, TextBox).Clear()
                        End If
                    Next
            End Select

            ' Gestione campo Identity
            ctrl.Enabled = Not campo.IsIdentity

        Next
    End Sub

    Private Sub CaricaValoridaGriglia()
        ' Carica valori dalla griglia nei controlli
        For Each campo In campiDefiniti
            campoInputs(campo.Nome).Text = dgvDati.SelectedRows(0).Cells(campo.Nome).Value.ToString()
        Next
    End Sub

    Private Sub ModificaDati(sender As Object, e As EventArgs)

        If dgvDati.SelectedRows.Count = 0 Then
            Dim risposta = MDIMessageBox.Show("Seleziona prima una riga dalla griglia.", Me.MdiParent, MessageBoxButtons.OK)
            Return
        End If

        DisabilitaPulsante("Salva", False)
        ModalitaCorrente = "modifica"
        lblModalita.Text = "Modifica in corso..."
        DisabilitaPulsante("Annulla", False)

        isModifica = True
        AbilitaCampi(True)
        ModalitaCorrente = "modifica"
        lblModalita.Text = "Modifica in corso..."
        lblModalita.ForeColor = Color.Green
        lblModalita.Font = New Font("Segoe UI", 8, FontStyle.Bold)

    End Sub

    Private Sub SalvaDati(sender As Object, e As EventArgs)
        If isModifica Then
            SalvaModifica()
        Else
            SalvaInserimento()
        End If

        DisabilitaCampi()
        CaricaDatiTabella(Me.Name)
        DisabilitaPulsante("Salva", True)
        lblModalita.ForeColor = Color.DarkGreen ' Colore neutro

        ModalitaCorrente = "nessuna"
        lblModalita.Text = ""
        DisabilitaPulsante("Annulla", True)

    End Sub

    Private Sub CancellaDati(sender As Object, e As EventArgs)
        If dgvDati.SelectedRows.Count = 0 Then
            Dim risposta = MDIMessageBox.Show("Seleziona prima una riga da cancellare dalla griglia.", Me.MdiParent, MessageBoxButtons.OK)
            Return
        End If

        campiDefiniti = RecuperaCampiDa(Me.Name)

        ' Trova la chiave primaria
        Dim campoChiave = campiDefiniti.FirstOrDefault(Function(c) c.IsChiave)
        If campoChiave Is Nothing Then
            Dim risposta = MDIMessageBox.Show("Nessuna chiave primaria.", Me.MdiParent, MessageBoxButtons.OK)
            Return
        End If

        Dim valoreChiave = dgvDati.SelectedRows(0).Cells(campoChiave.Nome).Value.ToString()
        Dim conferma = MDIMessageBox.Show($"Sei sicuro di voler cancellare il record con chiave {campoChiave.Nome} = {valoreChiave}?", Me.MdiParent, MessageBoxButtons.YesNo
)
        If conferma = DialogResult.Yes Then
            Dim query As String = $"DELETE FROM [{Me.Name}] WHERE {campoChiave.Nome} = @{campoChiave.Nome}"

            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@" & campoChiave.Nome, valoreChiave)
                    conn.Open()
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            CaricaDatiTabella(Me.Name)
        End If
    End Sub

    Private Sub SalvaInserimento()
        Dim campiBit As String() = {"CanView", "CanInsert", "CanUpdate", "CanDelete"}

        Dim colonne = campiDefiniti.
        Where(Function(c) Not c.IsIdentity).
        Select(Function(c) c.Nome).ToList()

        Dim query As String = $"INSERT INTO [{Me.Name}] ({String.Join(",", colonne)}) " &
                          $"VALUES ({String.Join(",", colonne.Select(Function(n) "@" & n))})"

        Try
            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand(query, conn)
                    For Each nomeCampo In colonne
                        Dim input = campoInputs(nomeCampo)
                        Dim valore = EstraiValoreDaControllo(nomeCampo, campoInputs(nomeCampo))
                        cmd.Parameters.AddWithValue("@" & nomeCampo, valore)
                    Next

                    conn.Open()
                    cmd.ExecuteNonQuery()
                End Using
            End Using

        Catch exSql As SqlException
            MDIMessageBox.Show($"Errore SQL: {exSql.Message}", Me.MdiParent, MessageBoxButtons.OK)
        Catch ex As Exception
            MDIMessageBox.Show($"Si è verificato un errore: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub SalvaModifica()
        campiDefiniti = RecuperaCampiDa(Me.Name)

        Dim campoChiave = campiDefiniti.FirstOrDefault(Function(c) c.IsChiave)
        If campoChiave Is Nothing OrElse dgvDati.SelectedRows.Count = 0 Then
            MDIMessageBox.Show("Chiave primaria mancante o nessuna riga selezionata.", Me.MdiParent, MessageBoxButtons.OK)
            Return
        End If

        Dim valoreChiaveObj = dgvDati.SelectedRows(0).Cells(campoChiave.Nome).Value
        If valoreChiaveObj Is Nothing OrElse valoreChiaveObj Is DBNull.Value Then
            MDIMessageBox.Show("Valore della chiave non trovato o nullo.", Me.MdiParent, MessageBoxButtons.OK)
            Return
        End If

        Dim valoreChiave = valoreChiaveObj.ToString()
        Dim campiBit As String() = {"CanView", "CanInsert", "CanUpdate", "CanDelete"}
        Dim cripta As New CriptaHash()
        Dim colonneValid As New List(Of String)

        ' Costruzione lista colonne da aggiornare
        For Each campo In campiDefiniti
            If campo.IsChiave OrElse campo.IsIdentity Then Continue For

            Dim input = campoInputs(campo.Nome)
            Dim isPassword = campo.Nome.ToLower().Contains("password")

            ' Salta password vuote
            If isPassword AndAlso TypeOf input Is TextBox AndAlso String.IsNullOrWhiteSpace(input.Text) Then
                Continue For
            End If

            colonneValid.Add(campo.Nome)
        Next

        Dim query As String = $"UPDATE [{Me.Name}] SET {String.Join(",", colonneValid.Select(Function(n) $"{n} = @{n}"))} WHERE {campoChiave.Nome} = @{campoChiave.Nome}"

        Using conn As New SqlConnection(ConnString)
            Using cmd As New SqlCommand(query, conn)
                For Each nomeCampo In colonneValid
                    Dim input = campoInputs(nomeCampo)
                    Dim valoreCampo = EstraiValoreDaControllo(nomeCampo, campoInputs(nomeCampo))
                    If valoreCampo IsNot Nothing Then
                        cmd.Parameters.AddWithValue("@" & nomeCampo, valoreCampo)
                    End If
                Next

                cmd.Parameters.AddWithValue("@" & campoChiave.Nome, valoreChiave)
                Debug.WriteLine(cmd.CommandText)
                conn.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Private Function EstraiValoreDaControllo(nomeCampo As String, input As Control) As Object
        Dim campiBit As String() = {"CanView", "CanInsert", "CanUpdate", "CanDelete"}
        Dim isPassword = nomeCampo.ToLower().Contains("password")

        Select Case True
            Case campiBit.Contains(nomeCampo, StringComparer.OrdinalIgnoreCase) AndAlso TypeOf input Is CheckBox
                Return CType(input, CheckBox).Checked

            Case TypeOf input Is ComboBox
                Return CType(input, ComboBox).SelectedValue

            Case TypeOf input Is FlowLayoutPanel
                For Each ctrl As Control In input.Controls
                    If TypeOf ctrl Is TextBox Then
                        Return CType(ctrl, TextBox).Text
                    End If
                Next
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
        If campo Is Nothing Then
            Return New Label() With {
            .Text = "Campo non valido.",
            .ForeColor = Color.Red,
            .AutoSize = True
        }
        End If

        Dim ctrl As Control = Nothing
        Dim tipoCampo = campo.Tipo.ToLower()
        Dim larghezzaStandard As Integer = 250

        If campo.IsIdentity Then
            Return New TextBox() With {
            .Width = 200,
            .ReadOnly = True,
            .ForeColor = Color.Gray,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
            .TextAlign = HorizontalAlignment.Left,
            .Margin = New Padding(5)
        }
        End If

        If Not String.IsNullOrEmpty(campo.TabellaCollegata) Then
            Dim combo As New ComboBox With {
        .DropDownStyle = ComboBoxStyle.DropDownList,
        .Width = 250,
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

                            ' 🔧 Aggiungi colonna combinata codice + descrizione
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
        End If

        Select Case tipoCampo
            Case "string"
                If campo.Nome.ToLower().Contains("password") Then
                    ctrl = New TextBox() With {
                    .Width = larghezzaStandard,
                    .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
                    .UseSystemPasswordChar = True
                }
                    AddHandler CType(ctrl, TextBox).KeyDown, AddressOf TextBoxPassword_KeyDown
                    AddHandler CType(ctrl, TextBox).MouseDown, AddressOf TextBoxPassword_MouseDown
                Else
                    ctrl = New TextBox() With {
                    .Width = larghezzaStandard,
                    .Anchor = AnchorStyles.Left Or AnchorStyles.Right
                }
                End If

            Case "date"
                ctrl = New DateTimePicker() With {
                .Width = larghezzaStandard,
                .Format = DateTimePickerFormat.Short,
                .Anchor = AnchorStyles.Left Or AnchorStyles.Right
            }

            Case "textbox", "text"
                ctrl = New TextBox() With {
                .Width = larghezzaStandard,
                .Anchor = AnchorStyles.Left Or AnchorStyles.Right
            }

            Case "combobox"
                ctrl = New ComboBox() With {
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .Width = larghezzaStandard,
                .Anchor = AnchorStyles.Left Or AnchorStyles.Right
            }

            Case "checkbox", "boolean"
                ctrl = New CheckBox() With {
                .Text = "",
                .AutoSize = True,
                .Anchor = AnchorStyles.Left
            }

            Case "int"
                ctrl = New NumericUpDown() With {
                .Width = larghezzaStandard,
                .Maximum = Integer.MaxValue,
                .Minimum = 0,
                .Anchor = AnchorStyles.Left Or AnchorStyles.Right
            }

            Case "imgvid"
                Dim pannelloMultimediale As New FlowLayoutPanel With {
                .AutoSize = True,
                .FlowDirection = FlowDirection.LeftToRight
            }

                Dim txtFileName As New TextBox() With {
                .Width = larghezzaStandard,
                .Text = ""
            }
                pannelloMultimediale.Controls.Add(txtFileName)

                Dim btnSfoglia As New Button() With {
                .Text = "Sfoglia...",
                .AutoSize = True
            }
                AddHandler btnSfoglia.Click, Sub()
                                                 Using dlg As New OpenFileDialog()
                                                     Dim percorso = OttieniPercorsoImgVid()
                                                     If String.IsNullOrWhiteSpace(percorso) Then
                                                         MDIMessageBox.Show("Percorso multimediale non configurato.", Me.MdiParent, MessageBoxButtons.OK)
                                                         Return
                                                     End If
                                                     dlg.InitialDirectory = percorso
                                                     dlg.Filter = "File multimediali|*.jpg;*.png;*.mp4;*.avi"
                                                     If dlg.ShowDialog() = DialogResult.OK Then
                                                         txtFileName.Text = Path.GetFileName(dlg.FileName)
                                                     End If
                                                 End Using
                                             End Sub
                pannelloMultimediale.Controls.Add(btnSfoglia)

                Dim btnView As New Button() With {
                .Text = "Visualizza",
                .AutoSize = True,
                .Enabled = True
            }
                AddHandler btnView.Click, Sub(sender, e)
                                              If dgvDati Is Nothing OrElse dgvDati.SelectedRows.Count = 0 Then
                                                  Dim parent = If(Me.MdiParent, Me)
                                                  MDIMessageBox.Show("Seleziona prima una riga nella griglia.", parent, MessageBoxButtons.OK)
                                                  Return
                                              End If

                                              Dim nomeFile = txtFileName.Text.Trim()
                                              Dim percorso = OttieniPercorsoImgVid()
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
                pannelloMultimediale.Controls.Add(btnView)

                ctrl = pannelloMultimediale

            Case Else
                ctrl = New Label() With {
                .Text = $"Tipo campo '{tipoCampo}' non gestito.",
                .ForeColor = Color.Red,
                .AutoSize = True
            }
        End Select

        If ctrl IsNot Nothing Then
            ctrl.Margin = New Padding(5)
        End If

        Return ctrl
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

    Private Function OttieniPercorsoImgVid() As String
        Using conn As New SqlConnection(ConnString)
            conn.Open()
            Dim query = "SELECT PercorsoImgVid FROM Sys_Parametri"
            Using cmd As New SqlCommand(query, conn)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing Then
                    Return result.ToString()
                End If
            End Using
        End Using
        Return ""
    End Function

    Private Sub TextBoxPassword_KeyDown(sender As Object, e As KeyEventArgs)
        If e.Control AndAlso (e.KeyCode = Keys.C OrElse e.KeyCode = Keys.V OrElse e.KeyCode = Keys.X) Then
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub TextBoxPassword_MouseDown(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Right Then
            CType(sender, TextBox).ContextMenuStrip = New ContextMenuStrip() ' Menu vuoto = blocco tasto destro
        End If
    End Sub


    ' Bottoni dinamici da Sys_Form_Actions
    Private Sub CaricaBottoniDinamici()

        Dim formNameCorrente As String = Me.Name
        Dim query As String = "SELECT FormActionName, ButtonText FROM Sys_Form_Actions WHERE FormName = @formName ORDER BY Ordine"

        Try
            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@formName", formNameCorrente)

                    conn.Open()
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim FormactionName As String = reader("FormActionName").ToString()
                            Dim buttonText As String = reader("ButtonText").ToString()

                            Dim btnDinamico As New Button With {
                            .Text = buttonText,
                            .AutoSize = True,
                            .Margin = New Padding(5),
                            .Tag = FormactionName
                        }

                            AddHandler btnDinamico.Click, Sub(s, e)
                                                              Dim campi = RecuperaCampiDa(FormactionName)
                                                              Dim form As New DynamicDataForm(campi, FormactionName)
                                                              form.MdiParent = GesPu25
                                                              form.Text = CType(s, Button).Text
                                                              form.Show()
                                                          End Sub

                            panelBottoni.Controls.Add(btnDinamico)
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Errore nel caricamento bottoni dinamici: " & ex.Message)
        End Try
    End Sub

    Private Function IntestazioneMultilinea(nomeColonna As String) As String
        If String.IsNullOrWhiteSpace(nomeColonna) Then Return ""

        Dim sb As New StringBuilder()
        sb.Append(nomeColonna(0))

        For i = 1 To nomeColonna.Length - 1
            Dim c = nomeColonna(i)
            If Char.IsUpper(c) Then
                sb.Append(vbCrLf) ' 🔁 Andata a capo prima delle maiuscole
            End If
            sb.Append(c)
        Next

        Return sb.ToString()
    End Function

    Private Sub CaricaDatiTabella(nomeTabella As String)
        Dim query As String = $"SELECT * FROM {nomeTabella}"

        Try
            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand(query, conn)
                    conn.Open()
                    Dim adapter As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    adapter.Fill(dt)

                    ' Assegna sorgente
                    dgvDati.DataSource = dt

                    'For Each col As DataGridViewColumn In dgvDati.Columns
                    'col.HeaderText = SpaziaMaiuscole(col.Name)
                    'Next

                    For Each col As DataGridViewColumn In dgvDati.Columns
                        col.HeaderText = IntestazioneMultilinea(col.Name)
                    Next


                    ' Aggiungi colonna manuale dopo il DataSource
                    If dgvDati.Columns("Miniatura") Is Nothing Then
                        Dim colImage As New DataGridViewImageColumn() With {
                        .Name = "Miniatura",
                        .HeaderText = "Preview",
                        .ImageLayout = DataGridViewImageCellLayout.Zoom
                    }
                        dgvDati.Columns.Insert(0, colImage)
                        dgvDati.RowTemplate.Height = 22
                    End If
                End Using
            End Using
        Catch ex As Exception
            Dim risposta = MDIMessageBox.Show("Errore nel caricamento dei dati della tabella.", Me.MdiParent, MessageBoxButtons.OK)
        End Try

    End Sub

    Private Function MappaTipoVisuale(tipoSql As String) As String
        Select Case tipoSql.ToLower()
            Case "bit" : Return "boolean"
            Case "date", "datetime" : Return "date"
            Case "varchar", "nvarchar", "text" : Return "string"
            Case "int", "bigint", "smallint" : Return "int"
            Case Else : Return "string"
        End Select
    End Function

    Private Sub BottoneDinamico_Click(sender As Object, e As EventArgs)
        Dim btn As Button = CType(sender, Button)
        Dim config As BottoneLogico = CType(btn.Tag, BottoneLogico)

        ' Recupera i campi della tabella destinazione
        Dim campiDestinazione As List(Of CampoDatabase) = RecuperaCampiDa(config.TabellaDestinazione)

        ' Apri form dinamico
        Dim formDestinazione As New DynamicDataForm(campiDestinazione, config.TabellaDestinazione)
        formDestinazione.Size = New Size(config.LarghezzaForm, config.AltezzaForm)
        'formDestinazione.MdiParent = Me
        formDestinazione.Text = config.Etichetta
        formDestinazione.Show()
    End Sub

    Private Sub dgvDati_SelectionChanged(sender As Object, e As EventArgs)
        If dgvDati.SelectedRows.Count > 0 Then
            CaricaDatiNeiControlli(dgvDati.SelectedRows(0))
        End If
    End Sub
    Private Sub dgvDati_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        CaricaDatiNeiControlli(dgvDati.Rows(e.RowIndex))
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
        sb.Append(text(0)) ' Prima maiuscola resta
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

                ' Controllo amministratore
                Dim isAdmin As Boolean = False
                Dim queryAdmin = "SELECT ISNULL(Amministratore, 0) FROM Sys_Utenti WHERE NomeUtente = @utente"

                Using cmdAdmin As New SqlCommand(queryAdmin, conn)
                    cmdAdmin.Parameters.AddWithValue("@utente", nomeUtente)
                    isAdmin = Convert.ToBoolean(cmdAdmin.ExecuteScalar())
                End Using

                If isAdmin Then
                    ' Amministratore: abilita tutto
                    For Each ctrl As Control In panelBottoni.Controls
                        If TypeOf ctrl Is Button Then CType(ctrl, Button).Enabled = True
                    Next
                    Return ' Evita il resto del controllo
                End If

                ' Controlli per utenti non amministratori
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
                            ' Nessuna autorizzazione: disabilita tutto
                            DisabilitaPulsante("Inserisci", True)
                            DisabilitaPulsante("Modifica", True)
                            DisabilitaPulsante("Cancella", True)
                            'Me.Close()
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Errore nel controllo autorizzazioni: " & ex.Message)
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
    Private Sub CaricaDatiNeiControlli(riga As DataGridViewRow)
        For Each campo In campoInputs.Keys
            If Not dgvDati.Columns.Contains(campo) Then Continue For

            Dim valoreObj = riga.Cells(campo).Value
            Dim valore = If(valoreObj IsNot DBNull.Value AndAlso valoreObj IsNot Nothing, valoreObj.ToString(), "")
            Dim ctrl = campoInputs(campo)
            Dim isPassword As Boolean = campo.ToLower().Contains("password")

            Select Case True
                Case TypeOf ctrl Is TextBox
                    CType(ctrl, TextBox).Text = If(isPassword, "", valore)

                Case TypeOf ctrl Is CheckBox
                    Dim booleano As Boolean
                    If Boolean.TryParse(valore, booleano) Then
                        CType(ctrl, CheckBox).Checked = booleano
                    Else
                        CType(ctrl, CheckBox).Checked = False
                    End If

                Case TypeOf ctrl Is ComboBox
                    ImpostaValoreCombo(CType(ctrl, ComboBox), valore)

                Case TypeOf ctrl Is DateTimePicker
                    Dim dt As DateTime
                    If DateTime.TryParse(valore, dt) Then
                        CType(ctrl, DateTimePicker).Value = dt
                    End If

                Case TypeOf ctrl Is FlowLayoutPanel
                    Dim combo As ComboBox = Nothing

                    For Each innerCtrl As Control In ctrl.Controls
                        If TypeOf innerCtrl Is ComboBox Then combo = CType(innerCtrl, ComboBox)
                    Next

                    If combo IsNot Nothing Then
                        ImpostaValoreCombo(combo, valore)
                    End If
            End Select
        Next
    End Sub

    Private Sub ImpostaValoreCombo(combo As ComboBox, valore As Object)
        If combo.DataSource Is Nothing OrElse combo.ValueMember Is Nothing Then
            combo.SelectedIndex = -1
            Return
        End If

        Dim esiste = combo.Items.Cast(Of DataRowView).Any(Function(r) r(combo.ValueMember).ToString() = valore.ToString())

        If esiste Then
            combo.SelectedValue = valore
        Else
            combo.SelectedIndex = -1
        End If
    End Sub


    Private Sub DynamicDataForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        GestioneStatoForm.SalvaStato(Me)
    End Sub

    Private Sub EsportaPDF(sender As Object, e As EventArgs)
        Try
            Dim document As New PdfDocument()
            document.Info.Title = $"Esportazione dati: {Me.Name}"

            ' Prima pagina
            Dim page As PdfPage = document.AddPage()
            page.Orientation = PageOrientation.Landscape
            Dim gfx As XGraphics = XGraphics.FromPdfPage(page)
            Dim font As New XFont("Arial", 8, XFontStyleEx.Regular)
            Dim formatter As New XTextFormatter(gfx)

            Dim margin As Double = 40
            Dim topOffset As Double = 60
            Dim lineHeight As Double = 20
            Dim pageHeight As Double = page.Height.Point
            Dim usableWidth As Double = page.Width.Point - (2 * margin)

            ' Colonne visibili
            Dim colonne = dgvDati.Columns.Cast(Of DataGridViewColumn).Where(Function(c) c.Visible).ToList()
            Dim colCount = colonne.Count
            If colCount = 0 Then
                MDIMessageBox.Show("Nessuna colonna visibile da esportare.", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim colWidth As Double = usableWidth / colCount

            ' Titolo
            gfx.DrawString($"Esportazione dati: {Me.Name}", New XFont("Arial", 11, XFontStyleEx.Bold), XBrushes.Black, New XPoint(margin, topOffset))
            topOffset += 30

            ' Intestazioni
            For i = 0 To colCount - 1
                Dim header = SpaziaPrimaDelleMaiuscole(colonne(i).HeaderText)
                Dim rect As New XRect(margin + (i * colWidth), topOffset, colWidth, lineHeight)
                formatter.DrawString(header, font, XBrushes.DarkBlue, rect, XStringFormats.TopLeft)
            Next
            topOffset += lineHeight

            ' Dati
            For Each row As DataGridViewRow In dgvDati.Rows
                If topOffset + (lineHeight * 2) > pageHeight - margin Then
                    ' Nuova pagina
                    page = document.AddPage()
                    page.Orientation = PageOrientation.Landscape
                    gfx = XGraphics.FromPdfPage(page)
                    formatter = New XTextFormatter(gfx)
                    topOffset = margin

                    ' Intestazioni ripetute
                    For i = 0 To colCount - 1
                        Dim header = SpaziaPrimaDelleMaiuscole(colonne(i).HeaderText)
                        Dim rect As New XRect(margin + (i * colWidth), topOffset, colWidth, lineHeight)
                        formatter.DrawString(header, font, XBrushes.DarkBlue, rect, XStringFormats.TopLeft)
                    Next
                    topOffset += lineHeight
                End If

                ' Righe alternate
                For i = 0 To colCount - 1
                    Dim valore = RipulisciStringa(row.Cells(colonne(i).Name).Value?.ToString())
                    Dim rect As New XRect(margin + (i * colWidth), topOffset, colWidth, lineHeight * 2)
                    If row.Index Mod 2 = 0 Then gfx.DrawRectangle(XBrushes.LightGray, rect)
                    formatter.DrawString(valore, font, XBrushes.Black, rect, XStringFormats.TopLeft)
                Next
                topOffset += lineHeight * 2
            Next

            ' Salvataggio
            Dim filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"{Me.Name}_Esportazione.pdf")
            document.Save(filePath)

            MDIMessageBox.Show($"PDF esportato con successo:\n{filePath}", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MDIMessageBox.Show("Errore durante l'esportazione PDF: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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

        ' Caricamento stato finestra
        GestioneStatoForm.CaricaStato(Me)

        ' Componenti visivi
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

        ' Carica contenuto multimediale
        VisualizzaContenuto(percorsoFile)
    End Sub

    Private Sub VisualizzaContenuto(percorsoOriginale As String)
        Dim estensione = Path.GetExtension(percorsoOriginale).ToLower()
        Dim formatoSupportato = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".mp4", ".avi", ".wmv", ".mov"}

        ' Verifica se il file è apribile
        If Not formatoSupportato.Contains(estensione) Then
            ' Prova JPG
            Dim baseName = Path.GetFileNameWithoutExtension(percorsoOriginale)
            Dim directory = Path.GetDirectoryName(percorsoOriginale)
            Dim jpgPath = Path.Combine(directory, baseName & ".jpg")

            If File.Exists(jpgPath) Then
                percorsoOriginale = jpgPath
                estensione = ".jpg"
            Else
                ' Prova PNG
                Dim pngPath = Path.Combine(directory, baseName & ".png")
                If File.Exists(pngPath) Then
                    percorsoOriginale = pngPath
                    estensione = ".png"
                Else
                    MessageBox.Show("Formato non riconosciuto e nessun file alternativo trovato: " & baseName)
                    Me.Close()
                    Return
                End If
            End If
        End If

        ' Visualizzazione finale
        If {".jpg", ".jpeg", ".png", ".bmp", ".gif"}.Contains(estensione) Then
            pictureBox.Image = Image.FromFile(percorsoOriginale)
            pictureBox.Visible = True
            mediaPlayer.Visible = False

        ElseIf {".mp4", ".avi", ".wmv", ".mov"}.Contains(estensione) Then
            mediaPlayer.URL = percorsoOriginale
            mediaPlayer.Visible = True
            pictureBox.Visible = False

        Else
            MessageBox.Show("Tipo di file non supportato: " & estensione)
            Me.Close()
        End If
    End Sub



    Private Sub VisualMediaForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        ' Salva stato finestra al momento della chiusura
        GestioneStatoForm.SalvaStato(Me)
    End Sub

End Class




