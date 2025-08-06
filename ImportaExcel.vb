Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports Microsoft.Data.SqlClient

Public Class ImportaExcel
    Dim excelPath As String
    Dim connStringExcel As String
    Dim connStringSQL As String = ConnString
    Dim campoSelezionato As String = ""

    ' Selezione file
    Private Sub btnCarica_Click(sender As Object, e As EventArgs) Handles btnCarica.Click
        Dim dlg As New OpenFileDialog With {.Filter = "Excel Files (*.xlsm)|*.xlsm"}
        If dlg.ShowDialog() = DialogResult.OK Then
            excelPath = dlg.FileName
            lblNomeFileExcel.Text = System.IO.Path.GetFileName(excelPath)
            connStringExcel = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={excelPath};Extended Properties='Excel 12.0 Macro;HDR=NO;IMEX=1;'"
            lstFogli.DataSource = CaricaFogliExcel()
            lstTabelle.DataSource = CaricaTabelleSQL()
        Else
            MDIMessageBox.Show("Operazione annullata. Nessun file selezionato.", Me.MdiParent, MessageBoxButtons.OK, "Info")
        End If
    End Sub

    ' Carica fogli Excel
    Private Function CaricaFogliExcel() As List(Of String)
        Dim fogli As New List(Of String)
        Using conn As New OleDbConnection(connStringExcel)
            conn.Open()
            Dim schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
            For Each row As DataRow In schema.Rows
                Dim nome = row("TABLE_NAME").ToString()
                If nome.EndsWith("$") Then fogli.Add(nome.TrimEnd("$"c))
            Next
        End Using
        Return fogli
    End Function

    ' Carica tabelle SQL
    Private Function CaricaTabelleSQL() As List(Of String)
        Dim tabelle As New List(Of String)
        Using conn As New SqlConnection(connStringSQL)
            conn.Open()
            Dim cmd = New SqlCommand("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME LIKE '%'", conn)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    tabelle.Add(reader("TABLE_NAME").ToString())
                End While
            End Using
        End Using
        Return tabelle
    End Function

    Private Sub lstFogli_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstFogli.SelectedIndexChanged
        If lstFogli.SelectedItem IsNot Nothing Then
            campoSelezionato = ""
            lstCollegamenti.Items.Clear() ' Azzera i collegamenti qui!
            lstColonne.DataSource = OttieniColonneDaRiga2(lstFogli.SelectedItem.ToString())
        End If
    End Sub

    Private Sub lstTabelle_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstTabelle.SelectedIndexChanged
        If lstTabelle.SelectedItem IsNot Nothing Then

            lstCampi.DataSource = OttieniCampiTabellaSQL(lstTabelle.SelectedItem.ToString())
        End If
    End Sub

    Private Sub lstCampi_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstCampi.SelectedIndexChanged
        campoSelezionato = lstCampi.SelectedItem?.ToString()
    End Sub

    Private Sub lstColonne_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstColonne.SelectedIndexChanged
        If Not String.IsNullOrEmpty(campoSelezionato) Then
            Dim colonnaExcel = lstColonne.SelectedItem?.ToString()
            If Not String.IsNullOrEmpty(colonnaExcel) Then
                Dim collegamento = $"{campoSelezionato} ⇆ {colonnaExcel}"
                If Not lstCollegamenti.Items.Contains(collegamento) Then
                    lstCollegamenti.Items.Add(collegamento)
                End If
                campoSelezionato = ""
            End If
        End If
    End Sub

    ' Importa dati
    Private Sub btnImporta_Click(sender As Object, e As EventArgs) Handles btnImporta.Click
        Try
            ' Validazioni iniziali
            If lstFogli.SelectedItem Is Nothing OrElse lstTabelle.SelectedItem Is Nothing Then
                MDIMessageBox.Show("Seleziona foglio Excel e tabella SQL.", Me.MdiParent, MessageBoxButtons.OK, "Avviso")
                Return
            End If

            If lstCollegamenti.Items.Count = 0 Then
                MDIMessageBox.Show("Nessuna mappatura definita tra colonne Excel e campi SQL.", Me.MdiParent, MessageBoxButtons.OK, "Avviso")
                Return
            End If

            ' Preparazione parametri
            Dim nomeFoglio As String = lstFogli.SelectedItem.ToString() + "$"
            Dim tabellaDestinazione As String = lstTabelle.SelectedItem.ToString()

            ' Costruzione mappature Excel → SQL
            Dim mappature As New Dictionary(Of String, String)
            For Each voce As String In lstCollegamenti.Items
                Dim parti = voce.Split("⇆"c)
                Dim campoSQL = parti(0).Trim()
                Dim colonnaExcel = parti(1).Trim()
                If Not mappature.ContainsKey(campoSQL) Then mappature.Add(campoSQL, colonnaExcel)
            Next

            ' Campi speciali
            Dim campiBit As List(Of String) = RicavaCampiBit(tabellaDestinazione)
            Dim campiMoney As List(Of String) = RicavaCampiMoney(tabellaDestinazione)

            ' Lettura dati da Excel
            Dim intestazioni As New List(Of String)
            Dim excelData As New List(Of Dictionary(Of String, String))

            Using connExcel As New OleDbConnection(connStringExcel)
                connExcel.Open()
                Using cmd As New OleDbCommand($"SELECT * FROM [{nomeFoglio}]", connExcel)
                    Using reader = cmd.ExecuteReader()
                        If reader.Read() AndAlso reader.Read() Then
                            For i = 0 To reader.FieldCount - 1
                                Dim header = reader(i).ToString().Trim()
                                intestazioni.Add(If(String.IsNullOrEmpty(header), $"Col{i + 1}", header))
                            Next
                        End If

                        Dim nomeChiave As String = intestazioni(0)

                        While reader.Read()
                            Dim riga As New Dictionary(Of String, String)
                            For i = 0 To reader.FieldCount - 1
                                riga.Add(intestazioni(i), reader(i).ToString())
                            Next

                            Dim valoreChiave = riga(nomeChiave).Trim()
                            If Not String.IsNullOrEmpty(valoreChiave) Then
                                excelData.Add(riga)
                            End If
                        End While

                        If excelData.Count = 0 Then
                            MDIMessageBox.Show($"Nessuna riga valida trovata: tutte le righe hanno '{intestazioni(0)}' vuoto.", Me.MdiParent,
                                           MessageBoxButtons.OK, "Importazione interrotta")
                            Return
                        End If
                    End Using
                End Using
            End Using

            ' Costruzione DataTable
            Dim dt As New DataTable()
            For Each campoSQL In mappature.Keys
                dt.Columns.Add(campoSQL)
            Next

            For Each riga In excelData
                Dim nuovaRiga = dt.NewRow()
                For Each campoSQL In mappature.Keys
                    Dim colExcel = mappature(campoSQL)
                    If riga.ContainsKey(colExcel) Then
                        Dim valore = PulisciTesto(riga(colExcel).ToString().Trim())

                        If campiBit.Contains(campoSQL) Then
                            nuovaRiga(campoSQL) = ConvertiSNtoBit(valore)
                        ElseIf campiMoney.Contains(campoSQL) Then
                            nuovaRiga(campoSQL) = ConvertiMoney(valore)
                        Else
                            nuovaRiga(campoSQL) = valore.Replace("'", "")
                        End If
                    Else
                        nuovaRiga(campoSQL) = DBNull.Value
                    End If
                Next
                dt.Rows.Add(nuovaRiga)
            Next

            ' Importazione con SqlBulkCopy
            Dim nomeChiaveSQL As String = mappature.Keys.First()

            Using connSQL As New SqlConnection(connStringSQL)
                connSQL.Open()
                Using bulkCopy As New SqlBulkCopy(connSQL)
                    bulkCopy.DestinationTableName = tabellaDestinazione
                    For Each campoSQL In mappature.Keys
                        bulkCopy.ColumnMappings.Add(campoSQL, campoSQL)
                    Next

                    Dim filtro = $"{nomeChiaveSQL} IS NOT NULL AND {nomeChiaveSQL} <> ''"
                    Dim righeValide = dt.Select(filtro)
                    If righeValide.Length = 0 Then
                        MDIMessageBox.Show($"Nessuna riga da importare: chiave '{nomeChiaveSQL}' non valorizzata.", Me.MdiParent,
                                       MessageBoxButtons.OK, "Importazione interrotta")
                        Return
                    End If

                    bulkCopy.WriteToServer(righeValide.CopyToDataTable())
                End Using
            End Using

            MDIMessageBox.Show("Importazione completata con successo!", Me.MdiParent, MessageBoxButtons.OK, "Excel → SQL")

        Catch ex As Exception
            Dim messaggioErrore As String = "Errore durante l'importazione:" & vbCrLf & ex.Message
            MDIMessageBox.Show(messaggioErrore, Me.MdiParent, MessageBoxButtons.OK, "Errore")
            Try
                Clipboard.SetText(messaggioErrore)
            Catch copyEx As Exception
                MDIMessageBox.Show("Errore nella copia negli appunti:" & vbCrLf & copyEx.Message, Me.MdiParent, MessageBoxButtons.OK, "Errore appunti")
            End Try
        End Try
    End Sub

    Private Function RicavaCampiMoney(nomeTabella As String) As List(Of String)
        Dim campiMoney As New List(Of String)

        Try
            Using conn As New SqlConnection(connStringSQL)
                conn.Open()
                Dim query As String = "
                SELECT COLUMN_NAME
                FROM INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_NAME = @Tabella AND DATA_TYPE = 'money'"

                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@Tabella", nomeTabella)

                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            campiMoney.Add(reader("COLUMN_NAME").ToString())
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MDIMessageBox.Show("Errore nel recupero dei campi MONEY:" & vbCrLf & ex.Message, Me.MdiParent, MessageBoxButtons.OK, "Errore SQL")
        End Try

        Return campiMoney
    End Function

    Private Function ConvertiMoney(valore As String) As Object
        valore = valore.Trim()
        If String.IsNullOrEmpty(valore) OrElse valore = "-" Then
            Return DBNull.Value
        End If

        Dim risultato As Decimal
        If Decimal.TryParse(valore, risultato) Then
            Return risultato
        Else
            Return DBNull.Value
        End If
    End Function

    Private Function OttieniColonneDaRiga2(nomeFoglio As String) As List(Of String)
        Dim colonne As New List(Of String)
        Using conn As New OleDbConnection(connStringExcel)
            conn.Open()
            Dim cmd As New OleDbCommand($"SELECT * FROM [{nomeFoglio}$]", conn)
            Using reader = cmd.ExecuteReader()
                If reader.Read() AndAlso reader.Read() Then
                    For i = 0 To reader.FieldCount - 1
                        Dim intestazione = reader(i).ToString().Trim()
                        colonne.Add(If(String.IsNullOrEmpty(intestazione), $"Intest. Col{i + 1} vuota", intestazione))
                    Next
                End If
            End Using
        End Using
        Return colonne
    End Function

    Private Function OttieniCampiTabellaSQL(tabella As String) As List(Of String)
        Dim campi As New List(Of String)
        Using conn As New SqlConnection(connStringSQL)
            conn.Open()
            Dim cmd = New SqlCommand($"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '{tabella}'", conn)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    campi.Add(reader("COLUMN_NAME").ToString())
                End While
            End Using
        End Using
        Return campi
    End Function

    Private Function OttieniColonneExcel(nomeFoglio As String) As List(Of String)
        Dim colonne As New List(Of String)
        Using conn As New OleDbConnection(connStringExcel)
            conn.Open()
            Dim cmd = New OleDbCommand($"SELECT TOP 2 * FROM [{nomeFoglio}]", conn)
            Using reader = cmd.ExecuteReader()
                For i = 0 To reader.FieldCount - 1
                    colonne.Add(reader.GetName(i))
                Next
            End Using
        End Using
        Return colonne
    End Function

    Private Sub lstCollegamenti_DoubleClick(sender As Object, e As EventArgs) Handles lstCollegamenti.DoubleClick
        If lstCollegamenti.SelectedItem IsNot Nothing Then
            Dim voce = lstCollegamenti.SelectedItem.ToString()
            Dim conferma = MDIMessageBox.Show($"Vuoi davvero eliminare il collegamento: {voce}?", Me.MdiParent, MessageBoxButtons.YesNo, "Conferma eliminazione")
            If conferma = DialogResult.Yes Then
                lstCollegamenti.Items.Remove(voce)
            End If
        End If
    End Sub

    Private Function RicavaCampiBit(tabella As String) As List(Of String)
        Dim campi As New List(Of String)
        Using conn As New SqlConnection(connStringSQL)
            conn.Open()
            Dim cmd = New SqlCommand($"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '{tabella}' AND DATA_TYPE = 'bit'", conn)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    campi.Add(reader("COLUMN_NAME").ToString())
                End While
            End Using
        End Using
        Return campi
    End Function

    Private Function ConvertiSNtoBit(valore As String) As Object
        If String.IsNullOrWhiteSpace(valore) Then Return DBNull.Value

        Select Case valore.Trim().ToUpper()
            Case "1", "S", "Si", "SI", "TRUE"
                Return True
            Case "0", "N", "No", "NO", "FALSE"
                Return False
            Case Else
                Return DBNull.Value ' oppure False, se preferisci default rigido
        End Select
    End Function

    Private Function PulisciTesto(valore As String) As String
        Dim sb As New System.Text.StringBuilder()
        For Each c As Char In valore
            If Not Char.IsControl(c) AndAlso AscW(c) < 127 OrElse AscW(c) > 159 Then
                sb.Append(c)
            End If
        Next
        Return sb.ToString().Trim()
    End Function

    Private Sub ImportaExcel_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Carica stato del form dal modulo condiviso
        GestioneStatoForm.CaricaStato(Me)
    End Sub

    Private Sub ImportaExcel_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        ' Salva stato Form nel modulo condiviso
        GestioneStatoForm.SalvaStato(Me)
    End Sub

End Class



