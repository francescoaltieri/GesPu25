'Imports System.Data
'Imports System.Data.SqlClient
Imports System.IO
Imports System.Linq
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Data.SqlClient

Imports System.ComponentModel
Imports System.Drawing

Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Imports a = DocumentFormat.OpenXml.Drawing
Imports pic = DocumentFormat.OpenXml.Drawing.Pictures
Imports wp = DocumentFormat.OpenXml.Drawing.Wordprocessing

Imports SysMath = System.Math

Public Class CreaLettAssegnazione
    Inherits Form

    Private _idAssegnazione As Integer = 0
    Private _recordCorrente As Dictionary(Of String, Object)

    Private Sub CreaLettAssegnazione_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            If Me.Tag IsNot Nothing Then
                Try
                    Dim t = Me.Tag.GetType()
                    Dim pRec = t.GetProperty("Record")
                    Dim pId = t.GetProperty("IdAssegnazione")
                    If pRec IsNot Nothing Then
                        _recordCorrente = TryCast(pRec.GetValue(Me.Tag), Dictionary(Of String, Object))
                    End If
                    If pId IsNot Nothing Then
                        Integer.TryParse(Convert.ToString(pId.GetValue(Me.Tag)), _idAssegnazione)
                    ElseIf _recordCorrente IsNot Nothing AndAlso _recordCorrente.ContainsKey("IdAssegnazione") Then
                        Integer.TryParse(Convert.ToString(_recordCorrente("IdAssegnazione")), _idAssegnazione)
                    End If
                Catch
                End Try
            End If
            RipristinaPosizioneForm(Me)
        Catch
        End Try
    End Sub

    Private Sub BtnCreaLett_Click(sender As Object, e As EventArgs) Handles btnCreaLett.Click
        Try
            If _idAssegnazione <= 0 Then
                If _recordCorrente IsNot Nothing AndAlso _recordCorrente.ContainsKey("IdAssegnazione") Then
                    Integer.TryParse(Convert.ToString(_recordCorrente("IdAssegnazione")), _idAssegnazione)
                End If
            End If
            If _idAssegnazione <= 0 Then
                MDIMessageBox.Show("IdAssegnazione non valido. Seleziona una riga corretta e riprova.", Me.MdiParent, MessageBoxButtons.OK)
                Return
            End If

            ' --- Leggi ComunicazioneId (template) e StudioAssegnatario dall'assegnazione ---
            Dim templateId As Integer = 0
            Dim studioAssegnatario As String = ""
            Using conn As New SqlConnection(ConnString)
                conn.Open()
                Using cmd As New SqlCommand("SELECT ComunicazioneId, StudioAssegnatario FROM Mov_Assegnazioni WHERE IdAssegnazione = @id", conn)
                    cmd.Parameters.AddWithValue("@id", _idAssegnazione)
                    Using rd = cmd.ExecuteReader()
                        If rd.Read() Then
                            If Not rd.IsDBNull(0) Then Integer.TryParse(Convert.ToString(rd.GetValue(0)), templateId)
                            studioAssegnatario = If(rd.IsDBNull(1), "", Convert.ToString(rd.GetValue(1)))
                        End If
                    End Using
                End Using
            End Using

            If templateId <= 0 Then
                MDIMessageBox.Show("Template comunicazione non trovato per questa assegnazione (ComunicazioneId mancante).", Me.MdiParent, MessageBoxButtons.OK)
                Return
            End If

            ' --- Leggi Tab_Comunicazioni usando IdComunicazione (chiave corretta) ---
            Dim marchioPath As String = ""
            Dim oggetto As String = ""
            Dim testoTemplate As String = ""
            Using conn As New SqlConnection(ConnString)
                conn.Open()
                Using cmd As New SqlCommand("SELECT FileMarchioCartaIntestata, Oggetto, TestoMessaggio FROM Tab_Comunicazioni WHERE IdComunicazione = @idCom", conn)
                    cmd.Parameters.AddWithValue("@idCom", templateId)
                    Using rd = cmd.ExecuteReader()
                        If rd.Read() Then
                            marchioPath = If(rd.IsDBNull(0), "", Convert.ToString(rd.GetValue(0)))
                            oggetto = If(rd.IsDBNull(1), "", Convert.ToString(rd.GetValue(1)))
                            testoTemplate = If(rd.IsDBNull(2), "", Convert.ToString(rd.GetValue(2)))
                        Else
                            MDIMessageBox.Show("Template comunicazione non trovato in Tab_Comunicazioni.", Me.MdiParent, MessageBoxButtons.OK)
                            Return
                        End If
                    End Using
                End Using
            End Using

            ' --- Leggi dati società di produzione (Sys_SocietàDiProduzione) ---
            Dim soc As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Using conn As New SqlConnection(ConnString)
                conn.Open()
                Using cmd As New SqlCommand("SELECT TOP 1 NomeSocietà, Indirizzo, CAP, Città, PIva, Rigo1, Rigo2, Rigo3 FROM Sys_SocietàDiProduzione", conn)
                    Using rd = cmd.ExecuteReader()
                        If rd.Read() Then
                            soc("NomeSocietà") = SafeStr(rd, "NomeSocietà")
                            soc("Indirizzo") = SafeStr(rd, "Indirizzo")
                            soc("CAP") = SafeStr(rd, "CAP")
                            soc("Città") = SafeStr(rd, "Città")
                            soc("PIva") = SafeStr(rd, "PIva")
                            soc("Rigo1") = SafeStr(rd, "Rigo1")
                            soc("Rigo2") = SafeStr(rd, "Rigo2")
                            soc("Rigo3") = SafeStr(rd, "Rigo3")
                        End If
                    End Using
                End Using
            End Using

            ' --- Leggi dati destinatario dallo studio (Tab_Fornitori) usando StudioAssegnatario come IdFornitore ---
            Dim destNome As String = ""
            Dim destIndirizzo As String = ""
            Dim destCAP As String = ""
            Dim destCitta As String = ""
            Dim destEmail As String = ""
            Dim destTelefono As String = ""

            If Not String.IsNullOrWhiteSpace(studioAssegnatario) Then
                Using conn As New SqlConnection(ConnString)
                    conn.Open()
                    Using cmd As New SqlCommand("SELECT TOP 1 Descrizione, Indirizzo, CAP, Città, Email, NumeroTelefono FROM Tab_Fornitori WHERE IdFornitore = @id", conn)
                        cmd.Parameters.AddWithValue("@id", studioAssegnatario)
                        Using rd = cmd.ExecuteReader()
                            If rd.Read() Then
                                destNome = If(rd.IsDBNull(0), "", Convert.ToString(rd.GetValue(0)))
                                destIndirizzo = If(rd.IsDBNull(1), "", Convert.ToString(rd.GetValue(1)))
                                destCAP = If(rd.IsDBNull(2), "", Convert.ToString(rd.GetValue(2)))
                                destCitta = If(rd.IsDBNull(3), "", Convert.ToString(rd.GetValue(3)))
                                destEmail = If(rd.IsDBNull(4), "", Convert.ToString(rd.GetValue(4)))
                                destTelefono = If(rd.IsDBNull(5), "", Convert.ToString(rd.GetValue(5)))
                            End If
                        End Using
                    End Using
                End Using
            Else
                Using conn As New SqlConnection(ConnString)
                    conn.Open()
                    Using cmd As New SqlCommand("SELECT TOP 1 StudioNome, StudioIndirizzo, StudioCAP, StudioCitta FROM Mov_Assegnazioni WHERE IdAssegnazione = @id", conn)
                        cmd.Parameters.AddWithValue("@id", _idAssegnazione)
                        Using rd = cmd.ExecuteReader()
                            If rd.Read() Then
                                destNome = If(rd.IsDBNull(0), "", Convert.ToString(rd.GetValue(0)))
                                destIndirizzo = If(rd.IsDBNull(1), "", Convert.ToString(rd.GetValue(1)))
                                destCAP = If(rd.IsDBNull(2), "", Convert.ToString(rd.GetValue(2)))
                                destCitta = If(rd.IsDBNull(3), "", Convert.ToString(rd.GetValue(3)))
                            End If
                        End Using
                    End Using
                End Using
            End If

            ' --- SaveFileDialog ---
            Dim sfd As New SaveFileDialog With {
                .Filter = "Documento Word|*.docx",
                .Title = "Salva comunicazione assegnazione",
                .FileName = "Comunicazione_Assegnazione_" & _idAssegnazione & ".docx",
                .InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            }
            If sfd.ShowDialog(Me) <> DialogResult.OK Then Return
            Dim outPath As String = sfd.FileName

            ' --- Ottieni PercorsoStudio: prima prova a leggere Tab_Fornitori.CartellaAssegnata ---
            Dim percorsoStudio As String = ""

            If Not String.IsNullOrWhiteSpace(studioAssegnatario) Then
                Try
                    Using conn As New SqlConnection(ConnString)
                        conn.Open()
                        Using cmd As New SqlCommand("SELECT CartellaAssegnata FROM Tab_Fornitori WHERE IdFornitore = @id", conn)
                            cmd.Parameters.AddWithValue("@id", studioAssegnatario)
                            Dim o = cmd.ExecuteScalar()
                            If o IsNot Nothing AndAlso Not Convert.IsDBNull(o) Then
                                percorsoStudio = Convert.ToString(o).Trim()
                            End If
                        End Using
                    End Using
                Catch ex As Exception
                    ' Se c'è un errore DB lasciamo percorsoStudio vuoto così da chiedere la cartella all'utente
                    percorsoStudio = ""
                End Try
            End If

            ' --- Risolvi variabili anche nell'oggetto (supporta <BR>, <@PercorsoStudio>, <@Campo ...>) ---
            oggetto = If(oggetto, "")
            oggetto = oggetto.Replace("<BR>", Environment.NewLine)
            oggetto = RisolviCampiInline(oggetto, _idAssegnazione, _recordCorrente, soc, percorsoStudio)

            ' --- Prepara messaggio e risolvi campi inline ---
            Dim messaggio As String = If(testoTemplate, "")
            messaggio = messaggio.Replace("<BR>", Environment.NewLine)
            messaggio = RisolviCampiInline(messaggio, _idAssegnazione, _recordCorrente, soc, percorsoStudio)


            ' --- Split al tag <@Tabella> ---
            Dim parts = messaggio.Split(New String() {"<@Tabella>"}, StringSplitOptions.None)
            Dim beforeTables = If(parts.Length > 0, parts(0), "")
            Dim afterTables = If(parts.Length > 1, String.Join("", parts.Skip(1)), "")

            ' --- Carica dati tabelle ---
            Dim dtLavA = CaricaLavA(_idAssegnazione)
            Dim dtLavD = CaricaLavD(_idAssegnazione)

            ' --- Crea documento con OpenXML ---
            Try
                Using wordDoc = WordprocessingDocument.Create(outPath, WordprocessingDocumentType.Document)
                    Dim mainPart = wordDoc.AddMainDocumentPart()
                    mainPart.Document = New Document(New Body())

                    Dim sect = EnsureDocumentSection(mainPart)

                    ' ---  Intestazione Logo + Ragione Sociale (corretto: aggiungi l'image part alla HeaderPart) ---
                    If Not String.IsNullOrWhiteSpace(marchioPath) AndAlso File.Exists(marchioPath) Then
                        Dim headerPart = mainPart.AddNewPart(Of HeaderPart)()
                        Dim imgId As String = Nothing
                        Using imgStream = File.OpenRead(marchioPath)
                            Dim ext = Path.GetExtension(marchioPath).ToLowerInvariant()
                            Dim imgType = If(ext = ".png", ImagePartType.Png, If(ext = ".gif", ImagePartType.Gif, ImagePartType.Jpeg))
                            ' aggiungi l'immagine direttamente alla HeaderPart (non al MainDocumentPart)
                            Dim imgPart = headerPart.AddImagePart(imgType)
                            imgPart.FeedData(imgStream)
                            imgId = headerPart.GetIdOfPart(imgPart)
                        End Using

                        Dim dims = GetImageDimensionsEMU(marchioPath)
                        Dim cx As Long = dims.Item1
                        Dim cy As Long = dims.Item2

                        Dim header = New Header()
                        Dim tbl = New Table()
                        ' table full page width
                        tbl.Append(New TableProperties(New TableWidth() With {.Type = TableWidthUnitValues.Pct, .Width = "5000"}))
                        Dim hdrRow = New TableRow()

                        ' crea le proprietà comuni per l'intestazione: interlinea 1 e spazi ridotti (Before=0, After=80)
                        Dim hdrParaProps = CreateParaProps(line:="240", before:="0", after:="80")

                        ' left cell: logo
                        Dim tcLeft = New TableCell()
                        Dim pLogo = New Paragraph()
                        pLogo.PrependChild(hdrParaProps.CloneNode(True))
                        Dim rLogo = New Run()
                        rLogo.Append(CreateImageElement(imgId, "Logo", cx, cy))
                        pLogo.Append(rLogo)
                        tcLeft.Append(pLogo)
                        tcLeft.Append(New TableCellProperties(New TableCellWidth() With {.Type = TableWidthUnitValues.Pct, .Width = "2500"}))
                        hdrRow.Append(tcLeft)

                        ' right cell: company header
                        Dim tcRight = New TableCell()
                        Dim pCompany = New Paragraph()
                        pCompany.PrependChild(hdrParaProps.CloneNode(True))
                        Dim rCompany = New Run()
                        Dim rp = New RunProperties()
                        rp.Append(New Bold())
                        rCompany.PrependChild(rp)
                        rCompany.Append(New Text(soc.GetValueOrDefault("NomeSocietà", "")))
                        pCompany.Append(rCompany)
                        tcRight.Append(pCompany)

                        If Not String.IsNullOrWhiteSpace(soc.GetValueOrDefault("Indirizzo", "")) Then
                            Dim pAddr = New Paragraph()
                            pAddr.PrependChild(hdrParaProps.CloneNode(True))
                            pAddr.Append(New Run(New Text(soc.GetValueOrDefault("Indirizzo", ""))))
                            tcRight.Append(pAddr)
                        End If

                        Dim capcitta = $"{soc.GetValueOrDefault("CAP", "")} {soc.GetValueOrDefault("Città", "")}".Trim()
                        If Not String.IsNullOrWhiteSpace(capcitta) Then
                            Dim pCap = New Paragraph()
                            pCap.PrependChild(hdrParaProps.CloneNode(True))
                            pCap.Append(New Run(New Text(capcitta)))
                            tcRight.Append(pCap)
                        End If

                        If Not String.IsNullOrWhiteSpace(soc.GetValueOrDefault("PIva", "")) Then
                            Dim pPiva = New Paragraph()
                            pPiva.PrependChild(hdrParaProps.CloneNode(True))
                            pPiva.Append(New Run(New Text("P.IVA " & soc.GetValueOrDefault("PIva", ""))))
                            tcRight.Append(pPiva)
                        End If

                        For Each r As String In New String() {soc.GetValueOrDefault("Rigo1", ""), soc.GetValueOrDefault("Rigo2", ""), soc.GetValueOrDefault("Rigo3", "")}
                            If Not String.IsNullOrWhiteSpace(r) Then
                                Dim pR = New Paragraph()
                                pR.PrependChild(hdrParaProps.CloneNode(True))
                                pR.Append(New Run(New Text(r)))
                                tcRight.Append(pR)
                            End If
                        Next

                        tcRight.Append(New TableCellProperties(New TableCellWidth() With {.Type = TableWidthUnitValues.Pct, .Width = "2500"}))
                        hdrRow.Append(tcRight)
                        tbl.Append(hdrRow)
                        header.Append(tbl)

                        headerPart.Header = header
                        Dim headerReference = New HeaderReference() With {.Type = HeaderFooterValues.Default, .Id = mainPart.GetIdOfPart(headerPart)}
                        sect.Append(headerReference)
                    Else
                        ' header fallback with company name only
                        Dim headerPart = mainPart.AddNewPart(Of HeaderPart)()
                        Dim header = New Header()
                        Dim p = New Paragraph()
                        p.PrependChild(CreateParaProps("240", "0", "80"))
                        Dim r = New Run()
                        Dim rp = New RunProperties()
                        rp.Append(New Bold())
                        r.PrependChild(rp)
                        r.Append(New Text(soc.GetValueOrDefault("NomeSocietà", "")))
                        p.Append(r)
                        header.Append(p)
                        headerPart.Header = header
                        Dim headerReference = New HeaderReference() With {.Type = HeaderFooterValues.Default, .Id = mainPart.GetIdOfPart(headerPart)}
                        sect.Append(headerReference)
                    End If

                    ' --- Body: recipient on right, then object, message and tables ---
                    Dim body = mainPart.Document.Body

                    ' Recipient block (right aligned) - usa interlinea 1 e piccoli margini
                    If Not String.IsNullOrWhiteSpace(destNome) OrElse Not String.IsNullOrWhiteSpace(destIndirizzo) Then
                        Dim pRecProps = New ParagraphProperties(New Justification() With {.Val = JustificationValues.Right})
                        pRecProps.PrependChild(CreateParaProps("240", "0", "80"))
                        Dim pRec = New Paragraph()
                        pRec.Append(pRecProps)
                        If Not String.IsNullOrWhiteSpace(destNome) Then pRec.Append(New Run(New RunProperties(New Bold()), New Text(destNome)))
                        If Not String.IsNullOrWhiteSpace(destIndirizzo) Then pRec.Append(New Run(New Break(), New Text(destIndirizzo)))
                        Dim capcittaDest = $"{destCAP} {destCitta}".Trim()
                        If Not String.IsNullOrWhiteSpace(capcittaDest) Then pRec.Append(New Run(New Break(), New Text(capcittaDest)))
                        If Not String.IsNullOrWhiteSpace(destEmail) Then pRec.Append(New Run(New Break(), New Text(destEmail)))
                        If Not String.IsNullOrWhiteSpace(destTelefono) Then pRec.Append(New Run(New Break(), New Text(destTelefono)))
                        body.Append(pRec)
                    End If

                    body.Append(New Paragraph())

                    ' Oggetto: può contenere tag HTML base -> converti in runs
                    Dim pObj = New Paragraph()
                    pObj.PrependChild(CreateParaProps("240", "40", "40")) ' piccoli margini attorno all'oggetto
                    For Each el As OpenXmlElement In CreateRunsFromHtml("Oggetto: " & oggetto)
                        pObj.Append(el)
                    Next
                    body.Append(pObj)

                    ' Data comunicazione subito dopo l'oggetto
                    Dim dataComunicazioneStr As String = ""
                    Dim dtCom As DateTime
                    If _recordCorrente IsNot Nothing AndAlso _recordCorrente.ContainsKey("DataComunicazione") AndAlso DateTime.TryParse(Convert.ToString(_recordCorrente("DataComunicazione")), dtCom) Then
                        dataComunicazioneStr = dtCom.ToString("dd/MM/yyyy")
                    Else
                        Using conn As New SqlConnection(ConnString)
                            conn.Open()
                            Using cmd As New SqlCommand("SELECT DataComunicazione FROM Mov_Assegnazioni WHERE IdAssegnazione = @id", conn)
                                cmd.Parameters.AddWithValue("@id", _idAssegnazione)
                                Dim o = cmd.ExecuteScalar()
                                If o IsNot Nothing AndAlso Not Convert.IsDBNull(o) AndAlso DateTime.TryParse(Convert.ToString(o), dtCom) Then
                                    dataComunicazioneStr = dtCom.ToString("dd/MM/yyyy")
                                End If
                            End Using
                        End Using
                    End If
                    body.Append(New Paragraph())
                    AppendParagraph(body, "Data comunicazione: " & dataComunicazioneStr, bold:=False, line:="240", before:="0", after:="80")
                    body.Append(New Paragraph())

                    ' Testo prima tabelle
                    If Not String.IsNullOrWhiteSpace(beforeTables) Then
                        For Each paraLine In beforeTables.Split(New String() {Environment.NewLine}, StringSplitOptions.None)
                            AppendParagraph(body, paraLine, bold:=False, line:="240", before:="0", after:="80")
                        Next
                    End If

                    ' Inserisci titolo e tabella Mov_AssegnazioniLavA (Lavorazioni Animazione)
                    If dtLavA IsNot Nothing AndAlso dtLavA.Rows.Count > 0 Then
                        AppendParagraph(body, "Lavorazioni Animazione", bold:=True, line:="240", before:="0", after:="40")
                        AppendTableFromDataTable(body, dtLavA, "Mov_AssegnazioniLavA")
                    End If

                    ' Inserisci titolo e tabella Mov_AssegnazioniLavD (Lavorazioni diverse)
                    If dtLavD IsNot Nothing AndAlso dtLavD.Rows.Count > 0 Then
                        AppendParagraph(body, "Lavorazioni diverse", bold:=True, line:="240", before:="12", after:="40")
                        AppendTableFromDataTable(body, dtLavD, "Mov_AssegnazioniLavD")
                    End If

                    ' Testo dopo tabelle
                    If Not String.IsNullOrWhiteSpace(afterTables) Then
                        For Each paraLine In afterTables.Split(New String() {Environment.NewLine}, StringSplitOptions.None)
                            AppendParagraph(body, paraLine, bold:=False, line:="240", before:="0", after:="80")
                        Next
                    End If

                    mainPart.Document.Save()
                End Using
            Catch exDoc As Exception
                MessageBox.Show("Errore creazione documento: " & exDoc.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try

            Try
                Process.Start(New ProcessStartInfo(outPath) With {.UseShellExecute = True})
            Catch exOpen As Exception
                MessageBox.Show("Documento creato ma non apribile automaticamente: " & exOpen.Message, "Comunicazione", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Try

        Catch ex As Exception
            MessageBox.Show("Errore creazione comunicazione: " & ex.Message, "Comunicazione", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function CreateParaProps(Optional line As String = "240", Optional before As String = "0", Optional after As String = "0") As ParagraphProperties
        Dim pp As New ParagraphProperties()
        pp.Append(New SpacingBetweenLines() With {
            .Line = line,
            .LineRule = LineSpacingRuleValues.Auto,
            .Before = before,
            .After = after
        })
        Return pp
    End Function

    Private Sub AppendParagraph(body As Body, text As String, Optional bold As Boolean = False)
        AppendParagraph(body, text, bold, "240", "0", "80")
    End Sub

    Private Sub AppendParagraph(body As Body, text As String, Optional bold As Boolean = False, Optional line As String = "240", Optional before As String = "0", Optional after As String = "80")
        Dim p = New Paragraph()
        Dim pp = CreateParaProps(line, before, after)
        p.Append(pp)
        Dim run = New Run()
        Dim rPr = New RunProperties()
        If bold Then rPr.Append(New Bold())
        run.PrependChild(rPr)
        run.Append(New Text(text) With {.Space = SpaceProcessingModeValues.Preserve})
        p.Append(run)
        body.Append(p)
    End Sub

    Private Function IsNumericString(s As String) As Boolean
        If String.IsNullOrWhiteSpace(s) Then Return False
        Dim tmp As Double
        Return Double.TryParse(s, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, tmp)
    End Function

    Private Function IsDateString(s As String) As Boolean
        If String.IsNullOrWhiteSpace(s) Then Return False
        Dim dt As DateTime
        Return DateTime.TryParse(Convert.ToString(s), Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, dt)
    End Function

    Private Sub AppendTableFromDataTable(body As Body, dt As DataTable, nomeTabella As String)
        Dim cols = dt.Columns.Cast(Of DataColumn).Select(Function(c) c.ColumnName).ToList()
        If String.Equals(nomeTabella, "Mov_AssegnazioniLavA", StringComparison.OrdinalIgnoreCase) Then
            If Not cols.Contains("SecondiScena") Then dt.Columns.Add("SecondiScena", GetType(Integer))
            If Not cols.Contains("FrameScena") Then dt.Columns.Add("FrameScena", GetType(Integer))
            If Not cols.Contains("TotFrameScena") Then dt.Columns.Add("TotFrameScena", GetType(Integer))
            cols = dt.Columns.Cast(Of DataColumn).Select(Function(c) c.ColumnName).ToList()
        End If

        Dim etichette = GetEtichettePerTabella(nomeTabella, cols)

        Dim table = New Table()

        Dim tblProps = New TableProperties(
        New TableWidth() With {.Type = TableWidthUnitValues.Pct, .Width = "5000"},
        New TableBorders(
            New TopBorder With {.Val = New EnumValue(Of BorderValues)(BorderValues.Single), .Size = 4},
            New LeftBorder With {.Val = New EnumValue(Of BorderValues)(BorderValues.Single), .Size = 4},
            New BottomBorder With {.Val = New EnumValue(Of BorderValues)(BorderValues.Single), .Size = 4},
            New RightBorder With {.Val = New EnumValue(Of BorderValues)(BorderValues.Single), .Size = 4},
            New InsideHorizontalBorder With {.Val = New EnumValue(Of BorderValues)(BorderValues.Single), .Size = 4},
            New InsideVerticalBorder With {.Val = New EnumValue(Of BorderValues)(BorderValues.Single), .Size = 4}
        )
    )
        table.AppendChild(tblProps)

        Dim lengths As New List(Of Integer)
        For Each c In cols
            Dim maxLen As Integer = If(etichette.ContainsKey(c), etichette(c).Length, c.Length)
            For Each dr As DataRow In dt.Rows
                Dim v = dr(c)
                Dim s = If(v Is Nothing OrElse Convert.IsDBNull(v), "", Convert.ToString(v))
                If s.Length > maxLen Then maxLen = s.Length
            Next
            lengths.Add(SysMath.Max(1, maxLen))
        Next
        Dim totalLen As Integer = lengths.Sum()
        If totalLen = 0 Then totalLen = cols.Count

        Dim hdrRow = New TableRow()
        For i = 0 To cols.Count - 1
            Dim colName = cols(i)
            Dim tc = New TableCell()

            Dim pHdr = New Paragraph()
            Dim ppHdr = New ParagraphProperties()
            ppHdr.Append(New Justification() With {.Val = JustificationValues.Center})
            pHdr.Append(ppHdr)

            Dim rpHdr = New RunProperties()
            rpHdr.Append(New Bold())
            rpHdr.Append(New FontSize() With {.Val = "20"}) ' 20 = 10pt
            pHdr.Append(New Run(rpHdr, New Text(etichette.GetValueOrDefault(colName, colName)) With {.Space = SpaceProcessingModeValues.Preserve}))
            tc.Append(pHdr)

            Dim pctWidth As Integer = CInt(SysMath.Round(5000.0 * (lengths(i) / CDbl(totalLen))))
            If pctWidth < 1 Then pctWidth = 1

            Dim tcp = New TableCellProperties()
            tcp.Append(New TableCellWidth() With {.Type = TableWidthUnitValues.Pct, .Width = pct_width_to_string(pctWidth)})
            tcp.Append(New TableCellVerticalAlignment() With {.Val = TableVerticalAlignmentValues.Center})
            tc.Append(tcp)

            hdrRow.Append(tc)
        Next
        table.Append(hdrRow)

        For Each dr As DataRow In dt.Rows
            Dim tr = New TableRow()
            For i = 0 To cols.Count - 1
                Dim colName = cols(i)
                Dim v = dr(colName)
                Dim s As String = ""

                Dim isDateCol As Boolean = String.Equals(colName, "DataPrevistaFirstRun", StringComparison.OrdinalIgnoreCase) OrElse
                                      String.Equals(colName, "DataPrevistaChiusura", StringComparison.OrdinalIgnoreCase) OrElse
                                      String.Equals(colName, "DataFirstRunPrevista", StringComparison.OrdinalIgnoreCase) OrElse
                                      String.Equals(colName, "DataChiusuraPrevista", StringComparison.OrdinalIgnoreCase)

                If isDateCol Then
                    If v IsNot Nothing AndAlso Not Convert.IsDBNull(v) Then
                        Dim dtVal As DateTime
                        If DateTime.TryParse(Convert.ToString(v), dtVal) Then
                            s = dtVal.ToString("dd/MM/yyyy")
                        Else
                            s = Convert.ToString(v)
                        End If
                    Else
                        s = ""
                    End If
                Else
                    s = If(v Is Nothing OrElse Convert.IsDBNull(v), "", Convert.ToString(v))
                End If

                Dim justification As JustificationValues = JustificationValues.Left
                If isDateCol OrElse IsDateString(s) Then
                    justification = JustificationValues.Center
                ElseIf IsNumericString(s) Then
                    justification = JustificationValues.Right
                Else
                    justification = JustificationValues.Left
                End If

                Dim pCell = New Paragraph()
                Dim ppCell = New ParagraphProperties()
                ppCell.Append(New Justification() With {.Val = justification})
                pCell.Append(ppCell)

                Dim rpCell = New RunProperties()
                rpCell.Append(New FontSize() With {.Val = "18"}) ' 18 = 9pt
                pCell.Append(New Run(rpCell, New Text(s) With {.Space = SpaceProcessingModeValues.Preserve}))

                Dim tc = New TableCell()
                tc.Append(pCell)

                Dim pctW As Integer = CInt(SysMath.Round(5000.0 * (lengths(i) / CDbl(totalLen))))
                If pctW < 1 Then pctW = 1
                Dim tcpell = New TableCellProperties()
                tcpell.Append(New TableCellWidth() With {.Type = TableWidthUnitValues.Pct, .Width = pct_width_to_string(pctW)})
                tcpell.Append(New TableCellVerticalAlignment() With {.Val = TableVerticalAlignmentValues.Center})
                tc.Append(tcpell)

                tr.Append(tc)
            Next
            table.Append(tr)
        Next

        body.Append(table)
    End Sub

    Private Function pct_width_to_string(value As Integer) As String
        ' la libreria usa stringhe numeriche per pct (es "2500")
        Return value.ToString()
    End Function

    Private Function GetEtichettePerTabella(nomeTabella As String, nomiColonne As IEnumerable(Of String)) As Dictionary(Of String, String)
        Dim map As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        For Each c In nomiColonne
            map(c) = c
        Next
        Try
            Using conn As New SqlConnection(ConnString)
                conn.Open()
                Using cmd As New SqlCommand("SELECT NomeColonna, TestoEtichetta FROM Sys_TestoEtichetta WHERE NomeTabella = @t", conn)
                    cmd.Parameters.AddWithValue("@t", nomeTabella)
                    Using rd = cmd.ExecuteReader()
                        While rd.Read()
                            Dim col = SafeStr(rd, "NomeColonna")
                            Dim txt = SafeStr(rd, "TestoEtichetta")
                            If Not String.IsNullOrWhiteSpace(col) AndAlso map.ContainsKey(col) AndAlso Not String.IsNullOrWhiteSpace(txt) Then
                                map(col) = txt
                            End If
                        End While
                    End Using
                End Using
            End Using
        Catch
        End Try
        Return map
    End Function

    Private Function LookupSingolo(tab As String, col As String, whereCol As String, whereVal As Object) As String
        Try
            Using conn As New SqlConnection(ConnString)
                conn.Open()
                Using cmd As New SqlCommand($"SELECT [{col}] FROM [{tab}] WHERE [{whereCol}] = @v", conn)
                    cmd.Parameters.AddWithValue("@v", If(whereVal, DBNull.Value))
                    Dim o = cmd.ExecuteScalar()
                    Return If(o Is Nothing OrElse Convert.IsDBNull(o), "", Convert.ToString(o))
                End Using
            End Using
        Catch
            Return ""
        End Try
    End Function

    Private Shared Function SafeStr(rd As SqlDataReader, colName As String) As String
        Try
            Dim ord = rd.GetOrdinal(colName)
            If ord >= 0 AndAlso Not rd.IsDBNull(ord) Then Return Convert.ToString(rd.GetValue(ord))
        Catch
        End Try
        Return ""
    End Function

    Private Function CaricaLavA(idAssegnazione As Integer) As DataTable
        Dim dt As New DataTable()
        Try
            Using conn As New SqlConnection(ConnString)
                conn.Open()
                Dim sql As String =
                    "SELECT A.EpisodioId, A.NumScena, L.Descrizione AS Lavorazione, A.DataPrevistaFirstRun AS DataPrevistaFirstRun, A.DataPrevistaChiusura AS DataPrevistaChiusura, " &
                    "A.SecondiScena, A.FrameScena, A.TotFrameScena " &
                    "FROM Mov_AssegnazioniLavA A " &
                    "LEFT JOIN Tab_Lavorazioni L ON L.IdLavorazione = A.LavorazioneId " &
                    "WHERE A.AssegnazioneId = @Id ORDER BY A.EpisodioId, A.NumScena"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@Id", idAssegnazione)
                    Using da As New SqlDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Errore lettura Mov_AssegnazioniLavA: " & ex.Message, "Creazione comunicazione", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return dt
    End Function

    Private Function CaricaLavD(idAssegnazione As Integer) As DataTable
        Dim dt As New DataTable()
        Try
            Using conn As New SqlConnection(ConnString)
                conn.Open()
                Dim sql As String =
                    "SELECT D.EpisodioId, L.Descrizione AS Lavorazione, D.DataFirstRunPrevista AS DataFirstRunPrevista, D.DataChiusuraPrevista AS DataChiusuraPrevista " &
                    "FROM Mov_AssegnazioniLavD D " &
                    "LEFT JOIN Tab_Lavorazioni L ON L.IdLavorazione = D.LavorazioneId " &
                    "WHERE D.AssegnazioneId = @Id ORDER BY D.EpisodioId"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@Id", idAssegnazione)
                    Using da As New SqlDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MDIMessageBox.Show("Errore lettura Mov_AssegnazioniLavD: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
        End Try
        Return dt
    End Function

    Private Function EnsureDocumentSection(mainPart As MainDocumentPart) As SectionProperties
        Dim body = mainPart.Document.Body
        Dim sect = body.Elements(Of SectionProperties)().FirstOrDefault()
        If sect Is Nothing Then
            sect = New SectionProperties()
            body.Append(sect)
        End If
        Return sect
    End Function

    Private Function CreateImageElement(imagePartId As String, name As String, cx As Long, cy As Long) As Drawing
        Dim element =
            New Drawing(
                New wp.Inline(
                    New wp.Extent() With {.Cx = cx, .Cy = cy},
                    New wp.EffectExtent() With {.LeftEdge = 0, .TopEdge = 0, .RightEdge = 0, .BottomEdge = 0},
                    New wp.DocProperties() With {.Id = CType(1UL, UInteger), .Name = name},
                    New wp.NonVisualGraphicFrameDrawingProperties(New a.GraphicFrameLocks() With {.NoChangeAspect = True}),
                    New a.Graphic(
                        New a.GraphicData(
                            New pic.Picture(
                                New pic.NonVisualPictureProperties(
                                    New pic.NonVisualDrawingProperties() With {.Id = CType(0UL, UInteger), .Name = name},
                                    New pic.NonVisualPictureDrawingProperties()
                                ),
                                New pic.BlipFill(New a.Blip() With {.Embed = imagePartId}, New a.Stretch(New a.FillRectangle())),
                                New pic.ShapeProperties(New a.Transform2D(New a.Offset() With {.X = 0, .Y = 0}, New a.Extents() With {.Cx = cx, .Cy = cy}), New a.PresetGeometry(New a.AdjustValueList()) With {.Preset = a.ShapeTypeValues.Rectangle})
                            )
                        ) With {.Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"}
                    )
                ) With {.DistanceFromTop = 0UI, .DistanceFromBottom = 0UI, .DistanceFromLeft = 0UI, .DistanceFromRight = 0UI}
            )
        Return element
    End Function

    Private Function GetImageDimensionsEMU(imagePath As String) As Tuple(Of Long, Long)
        Try
            Using img = System.Drawing.Image.FromFile(imagePath)
                Dim horzDpi = img.HorizontalResolution
                Dim vertDpi = img.VerticalResolution
                Dim pxW = img.Width
                Dim pxH = img.Height
                Dim emuPerPixelX = 914400.0! / horzDpi
                Dim emuPerPixelY = 914400.0! / vertDpi
                Dim cx = CLng(pxW * emuPerPixelX)
                Dim cy = CLng(pxH * emuPerPixelY)
                Dim maxEmu = 914400L * 6
                If cx > maxEmu Then
                    Dim ratio = maxEmu / cx
                    cx = maxEmu
                    cy = CLng(cy * ratio)
                End If
                Return Tuple.Create(cx, cy)
            End Using
        Catch
            Return Tuple.Create(914400L * 2, 914400L)
        End Try
    End Function

    Private Function RisolviCampiInline(input As String,
                                    idAssegnazione As Integer,
                                    rec As Dictionary(Of String, Object),
                                    soc As Dictionary(Of String, String),
                                    Optional percorsoStudio As String = "") As String

        If String.IsNullOrEmpty(input) Then Return String.Empty

        Dim result As String = input

        ' --- Gestione speciale per <@PercorsoStudio> ---
        If result.Contains("<@PercorsoStudio>") Then
            If String.IsNullOrWhiteSpace(percorsoStudio) Then
                ' Inserisco un marker speciale che poi verrà convertito in Run evidenziato
                result = result.Replace("<@PercorsoStudio>", "[[MancaPercorso]]")
            Else
                result = result.Replace("<@PercorsoStudio>", percorsoStudio)
            End If
        End If

        ' --- Gestione campi generici <@Campo Tab.Colonna> ---
        Dim rx As New Regex("<@Campo\s+([A-Za-z0-9_\.\-]+)[\.;]([A-Za-z0-9_]+)>", RegexOptions.IgnoreCase)

        result = rx.Replace(result, Function(m)
                                        Dim tabRaw = m.Groups(1).Value.Trim()
                                        Dim col = m.Groups(2).Value.Trim()

                                        Dim tab = tabRaw

                                        If String.Equals(tab, "Mov_Assegnazioni", StringComparison.OrdinalIgnoreCase) Then
                                            If rec IsNot Nothing AndAlso rec.ContainsKey(col) Then
                                                Dim val = rec(col)
                                                Return If(val Is Nothing OrElse Convert.IsDBNull(val), "", Convert.ToString(val))
                                            End If
                                            Return LookupSingolo(tab, col, "IdAssegnazione", idAssegnazione)
                                        End If

                                        If String.Equals(tab, "Sys_SocietàDiProduzione", StringComparison.OrdinalIgnoreCase) OrElse
                                       String.Equals(tab, "Sys_SocietaDiProduzione", StringComparison.OrdinalIgnoreCase) Then
                                            Return If(soc IsNot Nothing AndAlso soc.ContainsKey(col), soc(col), "")
                                        End If

                                        If rec IsNot Nothing AndAlso rec.ContainsKey(col) Then
                                            Dim val = rec(col)
                                            Return If(val Is Nothing OrElse Convert.IsDBNull(val), "", Convert.ToString(val))
                                        End If

                                        Return LookupSingolo(tab, col, "IdAssegnazione", idAssegnazione)
                                    End Function)

        Return result
    End Function

    Private Iterator Function CreateRunsFromHtml(html As String) As IEnumerable(Of OpenXmlElement)
        If String.IsNullOrEmpty(html) Then
            Yield New Run(New Text("") With {.Space = SpaceProcessingModeValues.Preserve})
            Exit Function
        End If

        Dim s = html
        Dim tokenRx As New Regex("(<\/?b>|<\/?i>|<\/?u>)", RegexOptions.IgnoreCase)
        Dim parts = tokenRx.Split(s)

        Dim boldCount As Integer = 0
        Dim italicCount As Integer = 0
        Dim underlineCount As Integer = 0

        For Each part In parts
            If String.IsNullOrEmpty(part) Then Continue For
            Dim lower = part.ToLowerInvariant().Trim()
            Select Case lower
                Case "<b>"
                    boldCount += 1
                Case "</b>"
                    If boldCount > 0 Then boldCount -= 1
                Case "<i>"
                    italicCount += 1
                Case "</i>"
                    If italicCount > 0 Then italicCount -= 1
                Case "<u>"
                    underlineCount += 1
                Case "</u>"
                    If underlineCount > 0 Then underlineCount -= 1
                Case Else
                    Dim run = New Run()
                    Dim rp = New RunProperties()
                    If boldCount > 0 Then rp.Append(New Bold())
                    If italicCount > 0 Then rp.Append(New Italic())
                    If underlineCount > 0 Then rp.Append(New Underline() With {.Val = UnderlineValues.Single})
                    If rp.HasChildren Then run.Append(rp)
                    run.Append(New Text(part) With {.Space = SpaceProcessingModeValues.Preserve})
                    Yield run
            End Select
        Next
    End Function

    Private Sub btnChiudi_Click(sender As Object, e As EventArgs) Handles btnChiudi.Click
        Me.Close()
    End Sub

    Private Sub CreaLettAssegnazione_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        SalvaPosizioneForm(Me)
    End Sub

    Private Class CopiedItem
        Public Property ModelPackId As String
        Public Property TipoOggettoLavorazioneId As String
        Public Property OggettoLavorazioneId As String
        Public Property FileName As String
        Public Property FolderKey As String
        Public Property EpisodioId As String
    End Class

    Private _lastAssegnazioneId As Integer
    Private _lastDataAssegnazioneStr As String
    Private Sub btnCaricaAllegati_Click(sender As Object, e As EventArgs) Handles btnCaricaAllegati.Click
        ' Disabilita UI
        btnCaricaAllegati.Enabled = False
        btnCreaLett.Enabled = False
        btnChiudi.Enabled = False

        Dim errors As New List(Of String)
        Dim copiedItems As New List(Of CopiedItem)

        ' Crea e configura il form di progresso
        Dim progressForm As New CopyProgressForm()
        progressForm.AssegnazioneId = _idAssegnazione

        Dim mdiHost As Form = Nothing
        If Me.MdiParent IsNot Nothing Then
            mdiHost = Me.MdiParent
        ElseIf Application.OpenForms("GesPu25") IsNot Nothing Then
            mdiHost = Application.OpenForms("GesPu25")
        End If

        If mdiHost IsNot Nothing Then
            progressForm.MdiParent = mdiHost
            progressForm.StartPosition = FormStartPosition.Manual
            Dim parentClient = mdiHost.ClientSize
            Dim x = System.Math.Max(0, (parentClient.Width - progressForm.Width) \ 2)
            Dim y = System.Math.Max(0, (parentClient.Height - progressForm.Height) \ 2)
            progressForm.Location = New Point(x, y)
            progressForm.Show()
        Else
            progressForm.StartPosition = FormStartPosition.CenterParent
            progressForm.Show(Me)
        End If
        progressForm.BringToFront()

        _lastAssegnazioneId = _idAssegnazione
        _lastDataAssegnazioneStr = String.Empty

        Try
            Using cn As New SqlConnection(ConnString)
                cn.Open()

                Dim operatoreRaw As Object = Nothing
                Dim dataAssegnazioneObj As Object = Nothing
                Using cmd As New SqlCommand("SELECT OperatoreAssegnatario, DataAssegnazione FROM Mov_Assegnazioni WHERE IdAssegnazione = @Id", cn)
                    cmd.Parameters.Add("@Id", SqlDbType.Int).Value = _idAssegnazione
                    Using rdr = cmd.ExecuteReader()
                        If rdr.Read() Then
                            operatoreRaw = If(rdr.IsDBNull(0), Nothing, rdr.GetValue(0))
                            dataAssegnazioneObj = If(rdr.IsDBNull(1), Nothing, rdr.GetValue(1))
                        End If
                    End Using
                End Using

                If operatoreRaw Is Nothing Then
                    errors.Add($"IdAssegnazione {_idAssegnazione}: OperatoreAssegnatario non trovato.")
                    progressForm.AddMessage($"Errore: Operatore assegnatario non trovato per Id {_idAssegnazione}")
                    GoTo EndProcess
                End If

                Dim operatoreKeyStr As String = operatoreRaw.ToString().Trim()
                If String.IsNullOrWhiteSpace(operatoreKeyStr) Then
                    errors.Add($"IdAssegnazione {_idAssegnazione}: OperatoreAssegnatario vuoto.")
                    progressForm.AddMessage($"Errore: Operatore assegnatario vuoto per Id {_idAssegnazione}")
                    GoTo EndProcess
                End If

                If dataAssegnazioneObj IsNot Nothing AndAlso Not IsDBNull(dataAssegnazioneObj) Then
                    Dim dt As DateTime
                    If DateTime.TryParse(dataAssegnazioneObj.ToString(), dt) Then
                        _lastDataAssegnazioneStr = dt.ToString("yyyy-MM-dd")
                    Else
                        _lastDataAssegnazioneStr = dataAssegnazioneObj.ToString()
                    End If
                End If

                Dim cartellaAssegnata As String = Nothing
                Using cmd As New SqlCommand("SELECT CartellaAssegnata FROM Tab_Fornitori WHERE IdFornitore = @key OR Descrizione = @key", cn)
                    cmd.Parameters.Add("@key", SqlDbType.NVarChar, 100).Value = operatoreKeyStr
                    Dim obj = cmd.ExecuteScalar()
                    If obj IsNot Nothing AndAlso obj IsNot DBNull.Value Then cartellaAssegnata = obj.ToString()
                End Using

                If String.IsNullOrWhiteSpace(cartellaAssegnata) Then
                    errors.Add($"Operatore '{operatoreKeyStr}': CartellaAssegnata non trovata.")
                    progressForm.AddMessage($"Errore: Cartella assegnata non trovata per operatore {operatoreKeyStr}")
                    GoTo EndProcess
                End If
                If Not Directory.Exists(cartellaAssegnata) Then Directory.CreateDirectory(cartellaAssegnata)

                Dim lavA_Scenes As New List(Of (EpisodioId As String, NumScena As String))
                Dim lavD_List As New List(Of (EpisodioId As String, LavorazioneId As String))

                Using cmd As New SqlCommand("SELECT EpisodioId, NumScena FROM Mov_AssegnazioniLavA WHERE AssegnazioneId = @Id", cn)
                    cmd.Parameters.Add("@Id", SqlDbType.Int).Value = _idAssegnazione
                    Using rdr = cmd.ExecuteReader()
                        While rdr.Read()
                            Dim episodio = If(rdr.IsDBNull(0), String.Empty, rdr.GetString(0))
                            Dim numScena = If(rdr.IsDBNull(1), String.Empty, rdr.GetString(1))
                            If Not String.IsNullOrWhiteSpace(numScena) Then
                                lavA_Scenes.Add((episodio, numScena))
                            End If
                        End While
                    End Using
                End Using

                Using cmd As New SqlCommand("SELECT EpisodioId, LavorazioneId FROM Mov_AssegnazioniLavD WHERE AssegnazioneId = @Id", cn)
                    cmd.Parameters.Add("@Id", SqlDbType.Int).Value = _idAssegnazione
                    Using rdr = cmd.ExecuteReader()
                        While rdr.Read()
                            Dim episodio = If(rdr.IsDBNull(0), String.Empty, rdr.GetString(0))
                            Dim lavorazioneId = If(rdr.IsDBNull(1), String.Empty, rdr.GetString(1))
                            If Not String.IsNullOrWhiteSpace(lavorazioneId) Then
                                lavD_List.Add((episodio, lavorazioneId))
                            End If
                        End While
                    End Using
                End Using

                If lavA_Scenes.Count = 0 Then
                    errors.Add($"Nessuna scena trovata in Mov_AssegnazioniLavA per IdAssegnazione {_idAssegnazione}.")
                    progressForm.AddMessage($"Nessuna scena LavA per assegnazione {_idAssegnazione}. Operazione interrotta.")
                    GoTo EndProcess
                End If

                Dim modelPackInfo As New Dictionary(Of String, (EpisodioId As String, NumScena As String))
                Dim paramNames As New List(Of String)
                Dim parameters As New List(Of SqlParameter)
                For i As Integer = 0 To lavA_Scenes.Count - 1
                    Dim pn = $"@ns{i}"
                    paramNames.Add(pn)
                    parameters.Add(New SqlParameter(pn, SqlDbType.NVarChar, 15) With {.Value = lavA_Scenes(i).NumScena})
                Next

                Dim sqlMp As String = "SELECT IdModelPack, EpisodioId, NumScena FROM Mov_ModelPack WHERE NumScena IN (" & String.Join(",", paramNames) & ")"
                Using cmd As New SqlCommand(sqlMp, cn)
                    For Each p In parameters
                        cmd.Parameters.Add(p)
                    Next
                    Using rdr = cmd.ExecuteReader()
                        While rdr.Read()
                            Dim idmpStr = If(rdr.IsDBNull(0), String.Empty, rdr.GetString(0))
                            Dim episodioIdStr = If(rdr.IsDBNull(1), String.Empty, rdr.GetString(1))
                            Dim scenaStr = If(rdr.IsDBNull(2), String.Empty, rdr.GetString(2))
                            If Not String.IsNullOrWhiteSpace(idmpStr) AndAlso Not modelPackInfo.ContainsKey(idmpStr) Then
                                modelPackInfo.Add(idmpStr, (episodioIdStr, scenaStr))
                            End If
                        End While
                    End Using
                End Using

                If modelPackInfo.Count = 0 Then
                    errors.Add($"Nessun ModelPack trovato per le scene collegate all'assegnazione {_idAssegnazione}.")
                    progressForm.AddMessage($"Nessun ModelPack trovato per le scene dell'assegnazione {_idAssegnazione}.")
                    GoTo EndProcess
                End If

                Using cmd As New SqlCommand("
                SELECT ModelPackId, TipoOggettoLavorazioneId, OggettoLavorazioneId, FileOggettoLavorazione 
                FROM Mov_ModelPackOggetti 
                WHERE ModelPackId = @ModelPackId", cn)
                    cmd.Parameters.Add("@ModelPackId", SqlDbType.NVarChar, 10)

                    For Each kvp In modelPackInfo
                        Dim idmp = kvp.Key
                        Dim episodioIdFromMp = kvp.Value.EpisodioId
                        Dim scenaFromMp = kvp.Value.NumScena

                        cmd.Parameters("@ModelPackId").Value = idmp

                        Using rdr = cmd.ExecuteReader()
                            While rdr.Read()
                                Dim mpId = If(rdr.IsDBNull(0), String.Empty, rdr.GetString(0))
                                Dim tipoOggetto = If(rdr.IsDBNull(1), String.Empty, rdr.GetString(1))
                                Dim oggettoLavId = If(rdr.IsDBNull(2), String.Empty, rdr.GetString(2))
                                Dim filePath = If(rdr.IsDBNull(3), String.Empty, rdr.GetString(3))

                                Dim fileNameOnly As String = If(String.IsNullOrWhiteSpace(filePath), String.Empty, Path.GetFileName(filePath))

                                ' Trova le entries LavA che hanno lo stesso NumScena (scenaFromMp)
                                Dim matchingLavA = lavA_Scenes.Where(Function(s) s.NumScena = scenaFromMp).ToList()

                                ' Se non ci sono entry LavA per questa scena, ignora l'oggetto
                                If matchingLavA.Count = 0 Then
                                    errors.Add($"ModelPack {mpId}: scena '{scenaFromMp}' non presente nelle righe LavA dell'assegnazione. Oggetto ignorato.")
                                    progressForm.AddMessage($"Oggetto ignorato per ModelPack {mpId}: scena {scenaFromMp} non collegata all'assegnazione.")
                                    Continue While
                                End If

                                ' Per ogni entry LavA corrispondente copia il file e logga con episodio+scena
                                For Each en In matchingLavA
                                    Dim episodioFolderName As String = If(String.IsNullOrWhiteSpace(en.EpisodioId), If(String.IsNullOrWhiteSpace(episodioIdFromMp), "UnknownEpisodio", episodioIdFromMp), en.EpisodioId)
                                    Dim folderKeyName As String = If(String.IsNullOrWhiteSpace(en.NumScena), "UnknownScene", en.NumScena)
                                    Dim destFolder As String = Path.Combine(cartellaAssegnata, episodioFolderName, "00.Materiali", folderKeyName)

                                    Try
                                        If String.IsNullOrWhiteSpace(filePath) Then
                                            errors.Add($"ModelPack {mpId}: FileOggettoLavorazione vuoto. Oggetto ignorato.")
                                            progressForm.AddMessage($"Ignorato (file vuoto): ModelPack {mpId}")
                                            Continue For
                                        End If
                                        If Not File.Exists(filePath) Then
                                            errors.Add($"ModelPack {mpId}: File non trovato: {filePath}. Oggetto ignorato.")
                                            progressForm.AddMessage($"File non trovato: {filePath}")
                                            Continue For
                                        End If

                                        If Not Directory.Exists(destFolder) Then Directory.CreateDirectory(destFolder)
                                        Dim destPath = Path.Combine(destFolder, fileNameOnly)
                                        File.Copy(filePath, destPath, True)

                                        ' Log nel form con episodio e scena
                                        progressForm.AddMessageFileScene(fileNameOnly, episodioFolderName, folderKeyName)

                                        ' Aggiungo l'item copiato
                                        copiedItems.Add(New CopiedItem With {
                                        .ModelPackId = mpId,
                                        .TipoOggettoLavorazioneId = tipoOggetto,
                                        .OggettoLavorazioneId = oggettoLavId,
                                        .FileName = fileNameOnly,
                                        .EpisodioId = episodioFolderName
                                    })
                                    Catch ex As Exception
                                        errors.Add($"Errore copia ModelPack {mpId} file '{filePath}': {ex.Message}")
                                        progressForm.AddMessage($"Errore copia: {ex.Message}")
                                    End Try
                                Next
                            End While
                        End Using
                    Next
                End Using

EndProcess:
                cn.Close()
            End Using

            Dim desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            Dim logFilePath = Path.Combine(desktopPath, $"CaricaAllegati_Log_{_idAssegnazione}_{DateTime.Now:yyyyMMdd_HHmmss}.txt")
            Try
                Using sw As New StreamWriter(logFilePath, False)
                    sw.WriteLine($"Log operazione Carica Allegati - IdAssegnazione: {_idAssegnazione}")
                    sw.WriteLine($"Data esecuzione: {DateTime.Now}")
                    sw.WriteLine()
                    If errors.Count = 0 Then
                        sw.WriteLine("Nessun errore.")
                    Else
                        sw.WriteLine("Errori:")
                        For Each er In errors
                            sw.WriteLine(er)
                        Next
                    End If
                    sw.WriteLine()
                    sw.WriteLine("Oggetti copiati:")
                    For Each ci In copiedItems
                        sw.WriteLine($"{ci.ModelPackId};{ci.TipoOggettoLavorazioneId};{ci.OggettoLavorazioneId};{ci.FileName};{ci.EpisodioId};{ci.FolderKey}")
                    Next
                End Using
                progressForm.AddMessage($"Log creato: {Path.GetFileName(logFilePath)}")
            Catch ex As Exception
                progressForm.AddMessage($"Impossibile creare log: {ex.Message}")
            End Try

            If copiedItems.Count > 0 Then
                Dim wordFilePath = Path.Combine(desktopPath, $"CaricaAllegati_List_{_idAssegnazione}_{DateTime.Now:yyyyMMdd_HHmmss}.docx")
                Try
                    CreateWordReport_OpenXml(copiedItems, wordFilePath)
                    progressForm.AddMessage($"Report Word creato: {Path.GetFileName(wordFilePath)}")
                Catch ex As Exception
                    errors.Add($"Errore creazione Word: {ex.Message}")
                    progressForm.AddMessage($"Errore creazione Word: {ex.Message}")
                End Try
            Else
                progressForm.AddMessage("Nessun file copiato, report Word non creato.")
            End If

            progressForm.AddMessage("Operazione completata.")
        Catch ex As Exception
            progressForm.AddMessage($"Errore generale: {ex.Message}")
        Finally
            btnCaricaAllegati.Enabled = True
            btnCreaLett.Enabled = True
            btnChiudi.Enabled = True
        End Try
    End Sub

    Private Sub CreateWordReport_OpenXml(items As List(Of CopiedItem), outputPath As String)
        If File.Exists(outputPath) Then
            Try
                File.Delete(outputPath)
            Catch
                outputPath = Path.Combine(Path.GetDirectoryName(outputPath),
                                      Path.GetFileNameWithoutExtension(outputPath) & "_new" & Path.GetExtension(outputPath))
            End Try
        End If

        Using wordDoc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document)
            Dim mainPart = wordDoc.AddMainDocumentPart()
            mainPart.Document = New Document(New Body())
            Dim body = mainPart.Document.Body

            ' Titolo centrato
            Dim titleText = $"Lista file copiati – Assegnazione: {_lastAssegnazioneId}  Data: {_lastDataAssegnazioneStr}"
            Dim titlePara As New Paragraph(
            New ParagraphProperties(New Justification() With {.Val = JustificationValues.Center},
                                    New SpacingBetweenLines() With {.After = "200"}),
            New Run(New RunProperties(New Bold(), New FontSize() With {.Val = "28"}), New Text(titleText))
        )
            body.Append(titlePara)
            body.Append(New Paragraph(New Run(New Text(String.Empty))))

            ' Tabella
            Dim table As New Table()
            Dim tblPr As New TableProperties(
            New TableStyle() With {.Val = "TableGrid"},
            New TableWidth() With {.Type = TableWidthUnitValues.Pct, .Width = "5000"},
            New TableBorders(
                New TopBorder() With {.Val = BorderValues.Single, .Size = 8},
                New BottomBorder() With {.Val = BorderValues.Single, .Size = 8},
                New LeftBorder() With {.Val = BorderValues.Single, .Size = 8},
                New RightBorder() With {.Val = BorderValues.Single, .Size = 8},
                New InsideHorizontalBorder() With {.Val = BorderValues.Single, .Size = 8},
                New InsideVerticalBorder() With {.Val = BorderValues.Single, .Size = 8}
            )
        )
            table.Append(tblPr)

            table.Append(New TableGrid(
            New GridColumn() With {.Width = "2400"},
            New GridColumn() With {.Width = "2200"},
            New GridColumn() With {.Width = "2400"},
            New GridColumn() With {.Width = "1800"},
            New GridColumn() With {.Width = "2200"},
            New GridColumn() With {.Width = "3600"}
        ))

            ' Header row (ripetuta su ogni pagina) - centrata orizz. e vert.
            Dim headerRow As New TableRow()
            headerRow.Append(New TableRowProperties(New TableHeader()))
            headerRow.Append(
            MakeHeaderCell("Model Pack"),
            MakeHeaderCell("Tipo Oggetto"),
            MakeHeaderCell("Oggetto Lavorazione"),
            MakeHeaderCell("Episodio"),
            MakeHeaderCell("Nome file trasmesso")
        )
            table.Append(headerRow)

            ' Righe dati: ultima colonna = FileName (solo nome + estensione)
            For Each it In items
                Dim tr As New TableRow()
                tr.Append(
                MakeCell(NullToEmpty(it.ModelPackId)),
                MakeCell(NullToEmpty(it.TipoOggettoLavorazioneId)),
                MakeCell(NullToEmpty(it.OggettoLavorazioneId)),
                MakeCell(NullToEmpty(it.EpisodioId)),
                MakeCell(NullToEmpty(it.FileName))
            )
                table.Append(tr)
            Next

            body.Append(table)
            mainPart.Document.Save()
        End Using
    End Sub

    ' Helper per header (centrato orizz. e vert., bold)
    Private Function MakeHeaderCell(text As String) As TableCell
        Dim r As New Run(New Text(If(text, String.Empty)))
        r.RunProperties = New RunProperties(New Bold())

        Dim pPr As New ParagraphProperties(New Justification() With {.Val = JustificationValues.Center})
        Dim p As New Paragraph(pPr, r)

        Dim tcPr As New TableCellProperties()
        tcPr.Append(New TableCellVerticalAlignment() With {.Val = TableVerticalAlignmentValues.Center})
        tcPr.Append(New TableCellWidth() With {.Type = TableWidthUnitValues.Auto})

        Return New TableCell(tcPr, p)
    End Function

    ' Helper per celle dati
    Private Function MakeCell(text As String, Optional centerHoriz As Boolean = False, Optional centerVert As Boolean = False) As TableCell
        Dim r As New Run(New Text(If(text, String.Empty)))
        Dim pPr As ParagraphProperties = Nothing
        If centerHoriz Then
            pPr = New ParagraphProperties(New Justification() With {.Val = JustificationValues.Center})
        Else
            pPr = New ParagraphProperties()
        End If
        Dim p As New Paragraph(pPr, r)

        Dim tcPr As New TableCellProperties(New TableCellWidth() With {.Type = TableWidthUnitValues.Auto})
        If centerVert Then tcPr.Append(New TableCellVerticalAlignment() With {.Val = TableVerticalAlignmentValues.Center})

        Return New TableCell(tcPr, p)
    End Function

    Private Function NullToEmpty(s As String) As String
        Return If(String.IsNullOrEmpty(s), String.Empty, s)
    End Function


    ' Helper per creare celle di tabella
    Private Function CreateTableCell(text As String, Optional isHeader As Boolean = False) As TableCell
        Dim tc = New TableCell()
        Dim p = New Paragraph()
        Dim r = New Run()
        r.Append(New Text(text))
        If isHeader Then
            r.RunProperties = New RunProperties(New Bold())
        End If
        p.Append(r)
        tc.Append(p)

        ' Imposta larghezza cella (opzionale)
        Dim tcPr = New TableCellProperties()
        tcPr.Append(New TableCellWidth() With {.Type = TableWidthUnitValues.Dxa, .Width = "2400"})
        tc.Append(tcPr)

        Return tc
    End Function

    Public Class CopyProgressForm
        Inherits Form

        Private txtLog As TextBox
        Private btnClose As Button
        Private btnSaveLog As Button

        ' Espone l'IdAssegnazione per includerlo nel nome del file di log
        Public Property AssegnazioneId As Integer

        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub InitializeComponent()
            Me.Text = "Progresso copia"
            Me.MinimumSize = New Size(640, 360)
            Me.Size = New Size(760, 460)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.Sizable

            ' TextBox log
            txtLog = New TextBox() With {
                .Multiline = True,
                .ScrollBars = ScrollBars.Vertical,
                .ReadOnly = True,
                .Font = New System.Drawing.Font("Consolas", 9),
                .Dock = DockStyle.Fill,
                .BackColor = System.Drawing.Color.White
            }

            ' Panel inferiore con TableLayout per evitare tagli dei bottoni
            Dim pnlBottom As New Panel() With {
                .Dock = DockStyle.Bottom,
                .Height = 56,
                .Padding = New Padding(8)
            }

            Dim tbl As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 3,
                .RowCount = 1
            }
            tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F)) ' spazio flessibile a sinistra
            tbl.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize)) ' Salva log
            tbl.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize)) ' Chiudi
            tbl.RowStyles.Add(New RowStyle(SizeType.AutoSize))

            btnSaveLog = New Button() With {
                .Text = "Salva log",
                .Width = 110,
                .Height = 30,
                .Anchor = AnchorStyles.Right Or AnchorStyles.Bottom,
                .Margin = New Padding(6)
            }
            AddHandler btnSaveLog.Click, AddressOf BtnSaveLog_Click

            btnClose = New Button() With {
                .Text = "Chiudi",
                .Width = 110,
                .Height = 30,
                .Anchor = AnchorStyles.Right Or AnchorStyles.Bottom,
                .Margin = New Padding(6)
            }
            AddHandler btnClose.Click, AddressOf BtnClose_Click

            ' Aggiungo i controlli alla tabella (i bottoni a destra)
            tbl.Controls.Add(New Label() With {.AutoSize = True}, 0, 0) ' placeholder per colonna espandibile
            tbl.Controls.Add(btnSaveLog, 1, 0)
            tbl.Controls.Add(btnClose, 2, 0)

            pnlBottom.Controls.Add(tbl)

            Me.Controls.Add(txtLog)
            Me.Controls.Add(pnlBottom)

        End Sub

        ' Aggiunge un messaggio generico con timestamp (thread-safe)
        Public Sub AddMessage(message As String)
            Dim line As String = $"{DateTime.Now:HH:mm:ss} - {message}"
            If Me.InvokeRequired Then
                Me.Invoke(New Action(Of String)(AddressOf AddMessageInternal), line)
            Else
                AddMessageInternal(line)
            End If
        End Sub

        ' Aggiunge un messaggio specifico per file + episodio + scena (thread-safe)
        Public Sub AddMessageFileScene(fileName As String, episodio As String, scena As String)
            Dim epText As String = If(String.IsNullOrWhiteSpace(episodio), "N/A", episodio)
            Dim scenaText As String = If(String.IsNullOrWhiteSpace(scena), "N/A", scena)
            Dim line As String = $"{DateTime.Now:HH:mm:ss} - Copia: {fileName} (Episodio: {epText}; Scena: {scenaText})"
            If Me.InvokeRequired Then
                Me.Invoke(New Action(Of String)(AddressOf AddMessageInternal), line)
            Else
                AddMessageInternal(line)
            End If
        End Sub

        ' Metodo interno che aggiorna la TextBox
        Private Sub AddMessageInternal(line As String)
            If txtLog.TextLength > 0 Then
                txtLog.AppendText(Environment.NewLine)
            End If
            txtLog.AppendText(line)
            txtLog.SelectionStart = txtLog.TextLength
            txtLog.ScrollToCaret()
        End Sub

        ' Salva il contenuto della TextBox in un file sul Desktop
        Private Sub BtnSaveLog_Click(sender As Object, e As EventArgs)
            SaveLogToDesktop()
        End Sub

        ' Chiude il form
        Private Sub BtnClose_Click(sender As Object, e As EventArgs)
            Me.Close()
        End Sub

        ' Salvataggio automatico alla chiusura (opzionale)
        Private Sub CopyProgressForm_FormClosing(sender As Object, e As FormClosingEventArgs)
            SaveLogToDesktop()
        End Sub

        ' Metodo che effettua il salvataggio del log sul Desktop
        Private Sub SaveLogToDesktop()
            Try
                Dim desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                Dim idPart As String = If(AssegnazioneId > 0, $"_Assegnazione_{AssegnazioneId}", "")
                Dim fileName = $"CaricaAllegati_ProgressLog{idPart}_{DateTime.Now:yyyyMMdd_HHmmss}.txt"
                Dim fullPath = Path.Combine(desktopPath, fileName)
                File.WriteAllText(fullPath, txtLog.Text)
                MessageBox.Show($"Log salvato sul Desktop: {fileName}", "Salva log", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show($"Errore salvataggio log: {ex.Message}", "Salva log", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        ' Metodo pubblico per ottenere il contenuto (opzionale)
        Public Function GetLogText() As String
            Return txtLog.Text
        End Function
    End Class

End Class
