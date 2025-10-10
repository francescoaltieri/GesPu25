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
Imports PdfSharp.Events
Imports PdfSharp.Pdf
Imports WMPLib
Imports System.Text.RegularExpressions

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
    Private ModalitaCorrente As String = "Nessuna"
    Private lblModalita As Label
    Private lampeggioAttivo As Boolean = False
    Private Shared visualFormsAttivi As New Dictionary(Of String, VisualMediaForm)
    Private panelBottoniDinamici As FlowLayoutPanel
    Private splitContainer As SplitContainer
    Public Property FiltroIniziale As String

    Public Sub New(campi As List(Of CampoDatabase), nomeTabella As String)
        Me.Name = nomeTabella
        Me.Text = "Form Dinamico"
        Me.Size = New Size(1100, 600)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.campiDefiniti = RecuperaCampiDa(nomeTabella)
        Me.nomeTabellaCorrente = nomeTabella

        AddHandler Me.Load, AddressOf DynamicDataForm_Load

        GestioneStatoForm.CaricaStato(Me)

        splitContainer = New SplitContainer With {
        .Dock = DockStyle.Fill,
        .Orientation = Orientation.Vertical,
        .FixedPanel = FixedPanel.None
}
        Me.Controls.Add(SplitContainer)

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
        .Font = New Font("Verdana", 10, FontStyle.Bold),
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
            'ctrl.Width = 250
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

        splitContainer.Panel2.Controls.Add(dgvDati)

        ' Caricamenti iniziali
        CaricaBottoniDinamici()

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

        UniformaDimensioniBottoni()
        ApplicaAutorizzazioni(NomeUtenteCorrente)
        PulisciCampi()

    End Sub

    Private Sub dgvDati_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs)
        If dgvDati.Columns.Count > 0 Then
            ApplicaConfigurazioneGriglia(dgvDati)
            ApplicaVisualizzazioneColonne()
            NascondiColonneSensibili()
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
        CaricaDatiTabella(Me.Name)
    End Sub



    Private Sub FormDinamico_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        ApplicaConfigurazioneGriglia(dgvDati)
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
        For Each kvp In campoInputs
            Dim nomeCampo As String = kvp.Key
            Dim ctrl As Control = kvp.Value

            ' Recupera definizione del campo
            Dim campo As CampoDatabase = campiDefiniti.FirstOrDefault(Function(c) c.Nome = nomeCampo)
            Dim isBloccato As Boolean = campo IsNot Nothing AndAlso (campo.IsIdentity OrElse campo.IsChiave)

            If TypeOf ctrl Is FlowLayoutPanel Then
                For Each innerCtrl As Control In ctrl.Controls
                    If TypeOf innerCtrl Is Button AndAlso CType(innerCtrl, Button).Text = "Visualizza" Then
                        innerCtrl.Enabled = True
                    ElseIf Not isBloccato Then
                        innerCtrl.Enabled = abilita
                    End If
                Next
            ElseIf Not isBloccato Then
                ctrl.Enabled = abilita
            End If

            If nomeCampo.StartsWith("Calc_") Then
                ctrl.Enabled = False
                Continue For
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
            MDIMessageBox.Show("Seleziona prima una riga da cancellare dalla griglia.", Me.MdiParent, MessageBoxButtons.OK)
            Return
        End If

        campiDefiniti = RecuperaCampiDa(Me.Name)

        ' Trova la chiave primaria
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

            Catch ex As SqlException
                If ex.Number = 547 Then ' Violazione vincolo FK
                    MDIMessageBox.Show("Impossibile cancellare il record: è referenziato da altre tabelle.", Me.MdiParent, MessageBoxButtons.OK)
                Else
                    MDIMessageBox.Show("Errore SQL durante la cancellazione: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
                End If
            Catch ex As Exception
                MDIMessageBox.Show("Errore imprevisto: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
            End Try
        End If
    End Sub


    Private Sub SalvaInserimento()
        Dim campiBit As String() = {"CanView", "CanInsert", "CanUpdate", "CanDelete"}

        ' Recupera formule e tipi
        Dim campiCalcolati = RecuperaCampiCalcolati()
        Dim formule = campiCalcolati.ToDictionary(Function(kvp) kvp.Key, Function(kvp) kvp.Value.Formula)
        Dim tipiValore = campiCalcolati.ToDictionary(Function(kvp) kvp.Key, Function(kvp) kvp.Value.TipoValore)

        ' Calcola i valori
        Dim valoriCalcolati = CalcolaValoriCampiCalcolati(formule, tipiValore)

        ' Aggiorna i controlli con i valori calcolati
        For Each kvp In valoriCalcolati
            If campoInputs.ContainsKey(kvp.Key) Then
                Dim ctrl = campoInputs(kvp.Key)
                If TypeOf ctrl Is TextBox Then
                    CType(ctrl, TextBox).Text = kvp.Value?.ToString()
                End If
            End If
        Next

        ' Costruisci query di inserimento
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
                        Dim valore = EstraiValoreDaControllo(nomeCampo, input)
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

        ' Recupera formule e tipi
        Dim campiCalcolati = RecuperaCampiCalcolati()
        Dim formule = campiCalcolati.ToDictionary(Function(kvp) kvp.Key, Function(kvp) kvp.Value.Formula)
        Dim tipiValore = campiCalcolati.ToDictionary(Function(kvp) kvp.Key, Function(kvp) kvp.Value.TipoValore)

        ' Calcola i valori
        Dim valoriCalcolati = CalcolaValoriCampiCalcolati(formule, tipiValore)

        ' Aggiorna i controlli con i valori calcolati
        For Each kvp In valoriCalcolati
            If campoInputs.ContainsKey(kvp.Key) Then
                Dim ctrl = campoInputs(kvp.Key)
                If TypeOf ctrl Is TextBox Then
                    CType(ctrl, TextBox).Text = kvp.Value?.ToString()
                End If
            End If
        Next

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

        Try
            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand(query, conn)
                    For Each nomeCampo In colonneValid
                        Dim input = campoInputs(nomeCampo)
                        Dim valoreCampo As Object
                        If valoriCalcolati.ContainsKey(nomeCampo) Then
                            valoreCampo = valoriCalcolati(nomeCampo)
                        Else
                            valoreCampo = EstraiValoreDaControllo(nomeCampo, input)
                        End If
                        cmd.Parameters.AddWithValue("@" & nomeCampo, valoreCampo)
                    Next

                    cmd.Parameters.AddWithValue("@" & campoChiave.Nome, valoreChiave)
                    conn.Open()
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            MDIMessageBox.Show("Modifica salvata correttamente.", Me.MdiParent, MessageBoxButtons.OK)

        Catch ex As SqlException
            If ex.Number = 547 Then
                MDIMessageBox.Show("Impossibile salvare: il record viola vincoli di integrità referenziale.", Me.MdiParent, MessageBoxButtons.OK)
            Else
                MDIMessageBox.Show("Errore SQL durante il salvataggio: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
            End If
        Catch ex As Exception
            MDIMessageBox.Show("Errore imprevisto: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
        End Try
    End Sub


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

            Case TypeOf input Is DateTimePicker
                Dim dtp = CType(input, DateTimePicker)
                Dim tag = dtp.Tag.ToString()

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
        'If campo.IsIdentity Then Return CreaTextBoxIdentity(campo)

        ' Calcolo larghezza dinamica
        Dim larghezzaBase As Integer = 100
        Dim larghezzaMassima As Integer = 450
        Dim larghezzaStimata As Integer
        Dim CarCtrl As Byte = 2

        If Len(campo.Nome.ToLower().ToString) < 2 Then
            CarCtrl = 1
        End If

        If campo.Lunghezza > 0 Then
            larghezzaStimata = Math.Min(larghezzaBase + campo.Lunghezza * 7, larghezzaMassima)
        ElseIf Not String.IsNullOrEmpty(campo.TabellaCollegata) Then
            larghezzaStimata = 380 ' ComboBox con ID + descrizione
        ElseIf campo.Nome.ToLower().Contains("descrizione") OrElse campo.Tipo.ToLower().Contains("text") Then
            larghezzaStimata = 350
        ElseIf campo.Tipo.ToLower().Contains("date") OrElse campo.Nome.ToLower().Substring(0, CarCtrl) = "id" OrElse campo.Nome.ToLower().Substring(0, CarCtrl) = "id" Then
            larghezzaStimata = 120
        Else
            larghezzaStimata = 350
        End If

        campo.Lunghezza = larghezzaStimata

        ' Creazione controllo
        Dim ctrl As Control

        If Not String.IsNullOrEmpty(campo.TabellaCollegata) Then
            ctrl = CreaComboDaTabella(campo)
        Else
            Select Case campo.Tipo.ToLower()
                Case "string", "string_max", "nvarchar", "varchar", "text"
                    ctrl = CreaTextBoxConGestioneTesto(campo)

                Case "date", "datetime"
                    ctrl = CreaDatePickerConGestioneVuoto(campo)

                Case "combobox"
                    ctrl = CreaComboBox()

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

        ' Applica larghezza se il controllo lo supporta
        If ctrl IsNot Nothing AndAlso Not TypeOf ctrl Is CheckBox AndAlso Not TypeOf ctrl Is Label Then
            If Not campo.Tipo.ToLower().Equals("money") Then
                ctrl.Width = larghezzaStimata
            End If

            If TypeOf ctrl Is ComboBox Then
                CType(ctrl, ComboBox).DropDownWidth = larghezzaStimata + 50
            End If
        End If

        If campo.IsIdentity Then
            ctrl.Enabled = False
        End If

        Return ctrl
    End Function

    Private Function CreaDatePickerConGestioneVuoto(campo As CampoDatabase) As Control
        Dim dtp As New DateTimePicker()
        dtp.Format = DateTimePickerFormat.Custom
        dtp.CustomFormat = " " ' campo visivamente vuoto
        dtp.Width = 140
        dtp.Tag = campo.Nome

        ' Mostra la data quando selezionata
        AddHandler dtp.ValueChanged, Sub()
                                         dtp.CustomFormat = "dd/MM/yyyy"
                                         dtp.Tag = campo.Nome ' rimuove eventuale flag di cancellazione
                                     End Sub

        ' Aggiungi menu contestuale per cancellare la data
        Dim menu As New ContextMenuStrip()
        menu.Items.Add("Cancella data", Nothing, Sub()
                                                     dtp.CustomFormat = " "               ' Nasconde la data
                                                     dtp.Value = dtp.MaxDate              ' Imposta una data fittizia non selezionabile
                                                     dtp.Tag = campo.Nome & "|NULL"       ' Flag per salvataggio
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

    Private Function CreaComboBox() As Control
        Dim campo As New CampoDatabase
        Return New ComboBox() With {
        .DropDownStyle = ComboBoxStyle.DropDownList,
        .Width = campo.Lunghezza,
        .Anchor = AnchorStyles.Left Or AnchorStyles.Right,
        .Margin = New Padding(5)
    }
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

        Dim txtFileName As New TextBox() With {.Width = 250, .Text = ""}
        pannello.Controls.Add(txtFileName)

        Dim btnView As New Button() With {.Text = "Visualizza", .AutoSize = True, .Enabled = True}
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
        ' Validazione iniziale: verifica che il nome del file sia almeno di 2 caratteri
        If String.IsNullOrWhiteSpace(NomeFile) OrElse NomeFile.Length < 2 Then
            Return ""
        End If

        Try
            Using conn As New SqlConnection(ConnString)
                conn.Open()

                ' Costruzione del parametro dinamico
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
            MDIMessageBox.Show("Errore nel recupero del percorso ", Me.MdiParent, MessageBoxButtons.OK)
            Return ""
        End Try

        ' Nessun risultato trovato
        MDIMessageBox.Show("Nella Tabella Sys_Paramettri non è statp trovato nessun risultato", Me.MdiParent, MessageBoxButtons.OK)
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
                                                              Dim valoreChiavePadre = dgvDati.SelectedRows(0).Cells(info.CampoPadre).Value?.ToString()

                                                              If String.IsNullOrWhiteSpace(valoreChiavePadre) Then
                                                                  MDIMessageBox.Show("Il valore della chiave primaria selezionata è nullo o non valido.", Me.MdiParent, MessageBoxButtons.OK)
                                                                  Return
                                                              End If

                                                              ' Verifica se il form è già aperto
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
                                                          End Sub



                            panelBottoni.Controls.Add(btnDinamico)
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MDIMessageBox.Show("Errore nel caricamento bottoni dinamici: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
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

                    ' Recupera i collegamenti definiti per la tabella
                    Dim dtCollegamenti As DataTable = EseguiQuery($"
                SELECT NomeCampo, TabellaCollegata, CampoValore, CampoVisuale
                FROM Sys_CollegamentiCampi
                WHERE NomeTabella = '{nomeTabella}'")

                    ' Ricostruisci le intestazioni e applica visualizzazione descrittiva
                    For Each col As DataGridViewColumn In dgvDati.Columns
                        Dim nomeCampo = col.Name
                        Dim etichetta = GetEtichetta(Me.Name, nomeCampo)
                        col.HeaderText = etichetta

                        ' Verifica se il campo ha un collegamento esterno
                        Dim rigaCollegata = dtCollegamenti.AsEnumerable().
                        FirstOrDefault(Function(r) r("NomeCampo").ToString() = nomeCampo)

                        If rigaCollegata IsNot Nothing Then
                            Dim tabellaCollegata = rigaCollegata("TabellaCollegata").ToString()
                            Dim campoValore = rigaCollegata("CampoValore").ToString()
                            Dim campoVisuale = rigaCollegata("CampoVisuale").ToString()

                            ' Costruisci dizionario ID → Descrizione
                            Dim dtValori = EseguiQuery($"
                        SELECT {campoValore}, {campoVisuale}
                        FROM {tabellaCollegata}")

                            Dim dizionario = dtValori.AsEnumerable().
                            ToDictionary(Function(r) r(campoValore).ToString(), Function(r) r(campoVisuale).ToString())

                            ' Sostituisci il valore visualizzato nella cella
                            For Each row As DataGridViewRow In dgvDati.Rows
                                If Not row.IsNewRow Then
                                    Dim id = row.Cells(nomeCampo).Value?.ToString()
                                    If Not String.IsNullOrEmpty(id) AndAlso dizionario.ContainsKey(id) Then
                                        row.Cells(nomeCampo).Value = $"{id} - {dizionario(id)}"
                                    End If
                                End If
                            Next
                        End If
                    Next
                End Using
            End Using
        Catch ex As Exception
            MDIMessageBox.Show("Errore nel caricamento dei dati della tabella." & vbCrLf & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
        End Try
    End Sub

    Public Function EseguiQuery(query As String) As DataTable
        Dim dt As New DataTable()

        Try
            Using conn As New SqlConnection(ConnString)
                Using cmd As New SqlCommand(query, conn)
                    cmd.CommandTimeout = 60 ' Timeout esteso per query complesse
                    conn.Open()
                    Using adapter As New SqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                End Using
            End Using
        Catch ex As SqlException
            ' Log dettagliato per errori SQL
            MDIMessageBox.Show("Errore SQL: " & ex.Message & vbCrLf & "Query: " & query, Nothing, MessageBoxButtons.OK)
        Catch ex As Exception
            ' Log generico per altri errori
            MDIMessageBox.Show("Errore generico: " & ex.Message, Nothing, MessageBoxButtons.OK)
        End Try

        Return dt
    End Function

    'Private Sub BottoneDinamico_Click(sender As Object, e As EventArgs)
    '    Dim btn As Button = CType(sender, Button)
    '    Dim config As BottoneLogico = CType(btn.Tag, BottoneLogico)
    '
    ' Cerca tra i form MDI già aperti
    '    For Each f As Form In GesPu25.MdiChildren
    '        If TypeOf f Is DynamicDataForm AndAlso f.Name = config.TabellaDestinazione Then
    '            f.WindowState = FormWindowState.Normal
    '            f.BringToFront()
    '            f.Activate()
    '            Return
    '        End If
    '    Next

    ' Recupera i campi della tabella destinazione
    '    Dim campiDestinazione As List(Of CampoDatabase) = RecuperaCampiDa(config.TabellaDestinazione)

    ' Crea e mostra nuovo form
    '    Dim formDestinazione As New DynamicDataForm(campiDestinazione, config.TabellaDestinazione)
    '    formDestinazione.Size = New Size(config.LarghezzaForm, config.AltezzaForm)
    '    formDestinazione.MdiParent = GesPu25
    '    formDestinazione.Text = config.Etichetta
    '    formDestinazione.Name = config.TabellaDestinazione
    '    formDestinazione.Show()
    'End Sub

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
        ' Recupera i collegamenti per la tabella corrente
        Dim dtCollegamenti As DataTable = EseguiQuery($"
        SELECT NomeCampo FROM Sys_CollegamentiCampi
        WHERE NomeTabella = '{nomeTabellaCorrente}'")

        Dim campiCollegati As HashSet(Of String) = New HashSet(Of String)(
        dtCollegamenti.AsEnumerable().Select(Function(r) r("NomeCampo").ToString()))

        For Each campo In campoInputs.Keys
            If Not dgvDati.Columns.Contains(campo) Then Continue For

            Dim valoreObj = riga.Cells(campo).Value
            Dim valoreRaw = If(valoreObj IsNot DBNull.Value AndAlso valoreObj IsNot Nothing, valoreObj.ToString(), "")
            Dim valore = If(campiCollegati.Contains(campo) AndAlso valoreRaw.Contains("-"),
                        valoreRaw.Split("-"c)(0).Trim(),
                        valoreRaw)

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
                    For Each innerCtrl As Control In ctrl.Controls
                        If TypeOf innerCtrl Is TextBox Then
                            CType(innerCtrl, TextBox).Text = valore
                        ElseIf TypeOf innerCtrl Is ComboBox Then
                            ImpostaValoreCombo(CType(innerCtrl, ComboBox), valore)
                        End If
                    Next
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
        SalvaConfigurazioneGriglia()
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

    Private Sub SalvaConfigurazioneGriglia()
        If dgvDati Is Nothing OrElse dgvDati.Columns.Count = 0 Then Exit Sub

        Dim nomeDbgrid As String = dgvDati.Name

        Using conn As New SqlConnection(ConnString)
            conn.Open()

            For Each col As DataGridViewColumn In dgvDati.Columns
                Dim nomeColonna As String = col.Name
                Dim colWidth As Integer = col.Width
                Dim visibile As Boolean = col.Visible

                Dim cmdCheck As New SqlCommand("
                SELECT COUNT(*) FROM Sys_VisualizzaInDbgrid 
                WHERE NomeTabella = @NomeTabella AND NomeColonna = @NomeColonna AND NomeDbgrid = @NomeDbgrid", conn)
                cmdCheck.Parameters.AddWithValue("@NomeTabella", nomeTabellaCorrente)
                cmdCheck.Parameters.AddWithValue("@NomeColonna", nomeColonna)
                cmdCheck.Parameters.AddWithValue("@NomeDbgrid", nomeDbgrid)

                Dim esiste As Boolean = Convert.ToInt32(cmdCheck.ExecuteScalar()) > 0

                If esiste Then
                    Dim cmdUpdate As New SqlCommand("
                    UPDATE Sys_VisualizzaInDbgrid 
                    SET ColWidth = @ColWidth, VisualizzaInDbgrid = @VisualizzaInDbgrid 
                    WHERE NomeTabella = @NomeTabella AND NomeColonna = @NomeColonna AND NomeDbgrid = @NomeDbgrid", conn)
                    cmdUpdate.Parameters.AddWithValue("@ColWidth", colWidth)
                    cmdUpdate.Parameters.AddWithValue("@VisualizzaInDbgrid", If(visibile, 1, 0))
                    cmdUpdate.Parameters.AddWithValue("@NomeTabella", nomeTabellaCorrente)
                    cmdUpdate.Parameters.AddWithValue("@NomeColonna", nomeColonna)
                    cmdUpdate.Parameters.AddWithValue("@NomeDbgrid", nomeDbgrid)
                    cmdUpdate.ExecuteNonQuery()
                Else
                    Dim cmdInsert As New SqlCommand("
                    INSERT INTO Sys_VisualizzaInDbgrid (NomeTabella, NomeColonna, NomeDbgrid, ColWidth, VisualizzaInDbgrid)
                    VALUES (@NomeTabella, @NomeColonna, @NomeDbgrid, @ColWidth, @VisualizzaInDbgrid)", conn)
                    cmdInsert.Parameters.AddWithValue("@NomeTabella", nomeTabellaCorrente)
                    cmdInsert.Parameters.AddWithValue("@NomeColonna", nomeColonna)
                    cmdInsert.Parameters.AddWithValue("@NomeDbgrid", nomeDbgrid)
                    cmdInsert.Parameters.AddWithValue("@ColWidth", colWidth)
                    cmdInsert.Parameters.AddWithValue("@VisualizzaInDbgrid", If(visibile, 1, 0))
                    cmdInsert.ExecuteNonQuery()
                End If
            Next
        End Using
    End Sub

    Private Sub ApplicaConfigurazioneGriglia(dgv As DataGridView)

        If dgv Is Nothing OrElse dgv.Columns.Count = 0 Then Exit Sub

        Using conn As New SqlConnection(ConnString)
            conn.Open()

            Dim cmd As New SqlCommand("
            SELECT NomeColonna, ColWidth, VisualizzaInDbgrid 
            FROM Sys_VisualizzaInDbgrid 
            WHERE NomeTabella = @NomeTabella AND NomeDbgrid = @NomeDbgrid", conn)
            cmd.Parameters.AddWithValue("@NomeTabella", nomeTabellaCorrente)
            cmd.Parameters.AddWithValue("@NomeDbgrid", dgv.Name)

            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim nomeColonna As String = reader("NomeColonna").ToString()
                    Dim colWidth As Integer = If(IsDBNull(reader("ColWidth")), 0, Convert.ToInt32(reader("ColWidth")))
                    Dim visibile As Boolean = If(IsDBNull(reader("VisualizzaInDbgrid")), True, Convert.ToBoolean(reader("VisualizzaInDbgrid")))

                    Dim col = dgv.Columns.Cast(Of DataGridViewColumn)().
                    FirstOrDefault(Function(c) String.Equals(c.Name, nomeColonna, StringComparison.OrdinalIgnoreCase))

                    If col IsNot Nothing Then
                        Try
                            col.Visible = visibile
                            If colWidth > 0 Then
                                col.Width = colWidth
                            Else
                                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                            End If
                        Catch ex As Exception
                            MDIMessageBox.Show($"Errore su colonna '{nomeColonna}': {ex.Message}", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                    Else
                        MDIMessageBox.Show($"Colonna non trovata o non inizializzata: {nomeColonna}", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End While
            End Using
        End Using
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

            ' Controllo tipo valido
            Dim tipiAmmessi As String() = {"numero", "stringa", "data", "booleano", "raw", "testo"}
            If Not tipiAmmessi.Contains(tipoValore) Then
                tipoValore = "Numero"
            End If

            Try
                Dim espressione = formula

                ' Sostituisci i nomi dei campi con i valori correnti
                For Each vInput In campoInputs
                    Dim nome = vInput.Key
                    Dim valore = EstraiValoreDaControllo(nome, vInput.Value)

                    Dim valoreStringa As String
                    If valore Is Nothing OrElse valore Is DBNull.Value Then
                        valoreStringa = If(tipoValore = "stringa", """ """, "0")
                        MDIMessageBox.Show($"[Campo calcolato] Il campo '{nome}' è nullo. Usato valore '{valoreStringa}'.", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ElseIf TypeOf valore Is DateTime Then
                        valoreStringa = $"""{CDate(valore).ToString("yyyy-MM-dd")}"""
                    ElseIf tipoValore = "stringa" Then
                        valoreStringa = $"""{valore.ToString().Replace("""", """""")}"""
                    Else
                        valoreStringa = Convert.ToString(valore, Globalization.CultureInfo.InvariantCulture)
                    End If

                    ' Sostituzione sicura con regex
                    Dim regex As New Regex($"\b{Regex.Escape(nome)}\b", RegexOptions.IgnoreCase)
                    espressione = regex.Replace(espressione, valoreStringa)
                Next

                ' Valutazione
                Dim risultato As Object
                If tipoValore = "numero" Then
                    risultato = New DataTable().Compute(espressione, Nothing)
                    If IsNumeric(risultato) Then
                        risultato = Convert.ToDouble(risultato, Globalization.CultureInfo.InvariantCulture)
                    Else
                        MDIMessageBox.Show($"[Campo calcolato] Il risultato di '{nomeCampo}' non è numerico: {risultato}", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)
                        risultato = Nothing
                    End If
                ElseIf tipoValore = "stringa" Then
                    risultato = ValutaEspressioneStringa(espressione)
                Else
                    risultato = espressione ' fallback
                End If

                risultati(nomeCampo) = risultato
            Catch ex As Exception
                MDIMessageBox.Show($"[Campo calcolato] Errore nel calcolo di '{nomeCampo}': {ex.Message}", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)
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
            MDIMessageBox.Show($"[StringEval] Errore: {ex.Message}", Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return ""
        End Try
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
                    MDIMessageBox.Show("Formato non riconosciuto e nessun file alternativo trovato: " & baseName, Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)
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
            MDIMessageBox.Show("Tipo di file non supportato: " & estensione, Me.MdiParent, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
        End If
    End Sub

    Private Sub VisualMediaForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        ' Salva stato finestra al momento della chiusura
        GestioneStatoForm.SalvaStato(Me)
    End Sub

End Class




