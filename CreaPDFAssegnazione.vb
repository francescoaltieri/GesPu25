Imports System.IO
Imports System.Text
Imports Microsoft.Data.SqlClient
Imports PdfSharp
Imports PdfSharp.Drawing
Imports PdfSharp.Drawing.Layout
Imports PdfSharp.Pdf

Public Class CreaPDFAssegnazione
    Inherits Form

    Private Sub btnAnnulla_Click(sender As Object, e As EventArgs) Handles btnAnnulla.Click
        Me.Close()
    End Sub

    Private Sub btnCreaPDF_Click(sender As Object, e As EventArgs) Handles btnCreaPDF.Click
        Try
            ' Recupera IdAssegnazione da Tag: supporta Tag = Integer oppure Tag = anonymous object con Record dictionary
            Dim idAssegnazione As Integer = -1

            If Me.Tag IsNot Nothing Then
                Try
                    If TypeOf Me.Tag Is Integer OrElse TypeOf Me.Tag Is Int32 Then
                        idAssegnazione = Convert.ToInt32(Me.Tag)
                    Else
                        ' Se Tag contiene .Record (dictionary) o proprieta IdAssegnazione
                        Dim tagType = Me.Tag.GetType()
                        Dim propId = tagType.GetProperty("IdAssegnazione")
                        If propId IsNot Nothing Then
                            idAssegnazione = Convert.ToInt32(propId.GetValue(Me.Tag))
                        ElseIf tagType.GetProperty("Record") IsNot Nothing Then
                            Dim rec = tagType.GetProperty("Record").GetValue(Me.Tag)
                            If rec IsNot Nothing Then
                                Dim recType = rec.GetType()
                                ' se è un Dictionary(Of String,Object)
                                If TypeOf rec Is System.Collections.IDictionary Then
                                    Dim dict = CType(rec, System.Collections.IDictionary)
                                    If dict.Contains("IdAssegnazione") Then idAssegnazione = Convert.ToInt32(dict("IdAssegnazione"))
                                    If idAssegnazione <= 0 AndAlso dict.Contains("Id") Then idAssegnazione = Convert.ToInt32(dict("Id"))
                                Else
                                    Dim recProp = recType.GetProperty("IdAssegnazione")
                                    If recProp IsNot Nothing Then idAssegnazione = Convert.ToInt32(recProp.GetValue(rec))
                                    If idAssegnazione <= 0 Then
                                        Dim recProp2 = recType.GetProperty("Id")
                                        If recProp2 IsNot Nothing Then idAssegnazione = Convert.ToInt32(recProp2.GetValue(rec))
                                    End If
                                End If
                            End If
                        Else
                            ' fallback: se Tag.ToString è numerico
                            Dim tmp As Integer = 0
                            If Integer.TryParse(Me.Tag.ToString(), tmp) Then idAssegnazione = tmp
                        End If
                    End If
                Catch
                End Try
            End If

            If idAssegnazione <= 0 Then
                MDIMessageBox.Show("Impossibile determinare l'IdAssegnazione. Assicurati di aprire il form dal contesto Mov_Assegnazioni oppure imposta Tag con l'Id.", Me.MdiParent, MessageBoxButtons.OK)
                Return
            End If

            Using sfd As New SaveFileDialog()
                sfd.Filter = "PDF File|*.pdf"
                sfd.FileName = $"Assegnazione_{idAssegnazione}.pdf"
                If sfd.ShowDialog(Me) <> DialogResult.OK Then Return
                Dim filePath = sfd.FileName

                ' Carica dati necessari dalla base
                Dim company As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                Dim assegnazione As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                Dim dtA As New DataTable()
                Dim dtD As New DataTable()

                Using conn As New SqlConnection(ConnString)
                    conn.Open()

                    ' Società di produzione (prendo la prima riga disponibile)
                    Using cmd As New SqlCommand("SELECT TOP 1 NomeSocietà, Indirizzo, CAP, Città, PIva, Rigo1, Rigo2, Rigo3 FROM Sys_SocietàDiProduzione", conn)
                        Using rdr = cmd.ExecuteReader()
                            If rdr.Read() Then
                                company("NomeSocietà") = If(rdr.IsDBNull(0), "", rdr.GetString(0))
                                company("Indirizzo") = If(rdr.IsDBNull(1), "", rdr.GetString(1))
                                company("CAP") = If(rdr.IsDBNull(2), "", rdr.GetString(2))
                                company("Città") = If(rdr.IsDBNull(3), "", rdr.GetString(3))
                                company("PIva") = If(rdr.IsDBNull(4), "", rdr.GetString(4))
                                company("Rigo1") = If(rdr.IsDBNull(5), "", rdr.GetString(5))
                                company("Rigo2") = If(rdr.IsDBNull(6), "", rdr.GetString(6))
                                company("Rigo3") = If(rdr.IsDBNull(7), "", rdr.GetString(7))
                            End If
                        End Using
                    End Using

                    ' Assegnazione principale
                    Using cmd As New SqlCommand("SELECT IdAssegnazione, DataAssegnazione, Descrizione, StudioAssegnatario, ContrattoId, DataComunicazione, TemplateId FROM Mov_Assegnazioni WHERE IdAssegnazione = @id", conn)
                        cmd.Parameters.AddWithValue("@id", idAssegnazione)
                        Using rdr = cmd.ExecuteReader()
                            If rdr.Read() Then
                                assegnazione("IdAssegnazione") = rdr("IdAssegnazione")
                                assegnazione("DataAssegnazione") = If(rdr.IsDBNull(1), Nothing, rdr.GetDateTime(1))
                                assegnazione("Descrizione") = If(rdr.IsDBNull(2), "", rdr.GetString(2))
                                assegnazione("StudioAssegnatario") = If(rdr.IsDBNull(3), "", rdr.GetString(3))
                                assegnazione("ContrattoId") = If(rdr.IsDBNull(4), "", rdr.GetString(4))
                                assegnazione("DataComunicazione") = If(rdr.IsDBNull(5), Nothing, rdr.GetDateTime(5))
                                assegnazione("TemplateId") = If(rdr.IsDBNull(6), Nothing, rdr("TemplateId"))
                            End If
                        End Using
                    End Using

                    ' Lavorazioni animazione - includo la descrizione dalla Tab_Lavorazioni
                    Dim sqlA As String = "SELECT a.*, ISNULL(t.Descrizione,'') AS LavorazioneDescrizione FROM Mov_AssegnazioniLavA a LEFT JOIN Tab_Lavorazioni t ON a.LavorazioneId = t.IdLavorazione WHERE a.AssegnazioneId = @id ORDER BY (SELECT NULL)"
                    Using da As New SqlDataAdapter(sqlA, conn)
                        da.SelectCommand.Parameters.AddWithValue("@id", idAssegnazione)
                        da.Fill(dtA)
                    End Using

                    ' Lavorazioni diverse - includo la descrizione dalla Tab_Lavorazioni
                    Dim sqlD As String = "SELECT d.*, ISNULL(t.Descrizione,'') AS LavorazioneDescrizione FROM Mov_AssegnazioniLavD d LEFT JOIN Tab_Lavorazioni t ON d.LavorazioneId = t.IdLavorazione WHERE d.AssegnazioneId = @id ORDER BY (SELECT NULL)"
                    Using da2 As New SqlDataAdapter(sqlD, conn)
                        da2.SelectCommand.Parameters.AddWithValue("@id", idAssegnazione)
                        da2.Fill(dtD)
                    End Using
                End Using

                ' Creazione PDF
                Dim document As New PdfDocument()
                document.Info.Title = $"Comunicazione Assegnazione {idAssegnazione}"

                Dim page As PdfPage = document.AddPage()
                page.Orientation = PageOrientation.Portrait
                Dim gfx As XGraphics = XGraphics.FromPdfPage(page)
                Dim fontTitle As New XFont("Verdana", 14, XFontStyleEx.Bold)
                Dim fontNorm As New XFont("Verdana", 10, XFontStyleEx.Regular)
                Dim fontBold As New XFont("Verdana", 10, XFontStyleEx.Bold)
                Dim tf As New XTextFormatter(gfx)

                Dim margin As Double = 40
                Dim y As Double = margin

                ' Header società
                Dim companyName = If(company.ContainsKey("NomeSocietà"), company("NomeSocietà"), "")
                If Not String.IsNullOrWhiteSpace(companyName) Then
                    gfx.DrawString(companyName, fontTitle, XBrushes.Black, New XRect(margin, y, page.Width.Point - 2 * margin, 30), XStringFormats.TopLeft)
                    y += 28
                End If

                Dim addrSb As New StringBuilder()
                If company.ContainsKey("Indirizzo") AndAlso Not String.IsNullOrWhiteSpace(company("Indirizzo")) Then addrSb.AppendLine(company("Indirizzo"))
                Dim cityLine As String = ""
                If company.ContainsKey("CAP") AndAlso Not String.IsNullOrWhiteSpace(company("CAP")) Then cityLine &= company("CAP") & " "
                If company.ContainsKey("Città") AndAlso Not String.IsNullOrWhiteSpace(company("Città")) Then cityLine &= company("Città")
                If Not String.IsNullOrWhiteSpace(cityLine) Then addrSb.AppendLine(cityLine)
                If company.ContainsKey("PIva") AndAlso Not String.IsNullOrWhiteSpace(company("PIva")) Then addrSb.AppendLine("P.IVA: " & company("PIva"))
                If company.ContainsKey("Rigo1") AndAlso Not String.IsNullOrWhiteSpace(company("Rigo1")) Then addrSb.AppendLine(company("Rigo1"))
                If company.ContainsKey("Rigo2") AndAlso Not String.IsNullOrWhiteSpace(company("Rigo2")) Then addrSb.AppendLine(company("Rigo2"))
                If company.ContainsKey("Rigo3") AndAlso Not String.IsNullOrWhiteSpace(company("Rigo3")) Then addrSb.AppendLine(company("Rigo3"))

                Dim addrText = addrSb.ToString().Trim()
                If Not String.IsNullOrWhiteSpace(addrText) Then
                    tf.DrawString(addrText, fontNorm, XBrushes.Black, New XRect(margin, y, page.Width.Point - 2 * margin, 80), XStringFormats.TopLeft)
                    y += 48
                End If

                y += 4
                ' Intestazione formale
                Dim study As String = If(assegnazione.ContainsKey("StudioAssegnatario"), Convert.ToString(assegnazione("StudioAssegnatario")), "")
                Dim dataAsseg = If(assegnazione.ContainsKey("DataAssegnazione") AndAlso assegnazione("DataAssegnazione") IsNot Nothing, CDate(assegnazione("DataAssegnazione")).ToString("dd/MM/yyyy"), "")

                Dim titolo As String = $"Comunicazione di assegnazione lavori n. {idAssegnazione}"
                tf.DrawString(titolo, fontBold, XBrushes.Black, New XRect(margin, y, page.Width.Point - 2 * margin, 20), XStringFormats.TopLeft)
                y += 20

                Dim destinatario As String = $"All'attenzione di: {study}"
                tf.DrawString(destinatario, fontNorm, XBrushes.Black, New XRect(margin, y, page.Width.Point - 2 * margin, 18), XStringFormats.TopLeft)
                y += 18

                If Not String.IsNullOrWhiteSpace(dataAsseg) Then
                    tf.DrawString($"Data assegnazione: {dataAsseg}", fontNorm, XBrushes.Black, New XRect(margin, y, 300, 18), XStringFormats.TopLeft)
                    y += 18
                End If

                y += 6

                If assegnazione.ContainsKey("Descrizione") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(assegnazione("Descrizione"))) Then
                    tf.DrawString("Oggetto: " & Convert.ToString(assegnazione("Descrizione")), fontNorm, XBrushes.Black, New XRect(margin, y, page.Width.Point - 2 * margin, 60), XStringFormats.TopLeft)
                    y += 36
                End If

                y += 6

                ' Funzione per disegnare tabella migliorata con etichette prese da Sys_TestoEtichetta
                Dim DrawTable = Sub(dt As DataTable, titoloTabella As String, nomeTabellaDb As String)
                                    If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

                                    ' Carica etichette dalla tabella Sys_TestoEtichetta per NomeTabella = nomeTabellaDb
                                    Dim headerLabels As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                                    Try
                                        Using conn As New SqlConnection(ConnString)
                                            conn.Open()
                                            Using cmd As New SqlCommand("SELECT NomeColonna, TestoEtichetta FROM Sys_TestoEtichetta WHERE NomeTabella = @t", conn)
                                                cmd.Parameters.AddWithValue("@t", nomeTabellaDb)
                                                Using rdr = cmd.ExecuteReader()
                                                    While rdr.Read()
                                                        Dim col = If(rdr.IsDBNull(0), String.Empty, rdr.GetString(0))
                                                        Dim txt = If(rdr.IsDBNull(1), String.Empty, rdr.GetString(1))
                                                        If Not String.IsNullOrWhiteSpace(col) Then
                                                            headerLabels(col) = txt
                                                        End If
                                                    End While
                                                End Using
                                            End Using
                                        End Using
                                    Catch
                                        ' se fallisce, prosegui con intestazioni generate automaticamente
                                    End Try

                                    ' Font per tabella: molto piccoli e non bold
                                    Dim fontHdr As New XFont("Arial", 7.5, XFontStyleEx.Regular)
                                    Dim fontCell As New XFont("Arial", 7, XFontStyleEx.Regular)
                                    Dim tfCell As New XTextFormatter(gfx)

                                    Dim minSpace As Double = 100
                                    If y + minSpace > page.Height.Point - margin Then
                                        page = document.AddPage()
                                        gfx = XGraphics.FromPdfPage(page)
                                        tf = New XTextFormatter(gfx)
                                        y = margin
                                    End If

                                    ' Titolo tabella
                                    tf.DrawString(titoloTabella, fontNorm, XBrushes.Black, New XRect(margin, y, page.Width.Point - 2 * margin, 14), XStringFormats.TopLeft)
                                    y += 14

                                    ' Determino larghezza colonne in modo dinamico basato su contenuti (misura testo)
                                    Dim colCount = Math.Max(1, dt.Columns.Count)
                                    Dim usableWidth = page.Width.Point - 2 * margin
                                    Dim maxColWidths(colCount - 1) As Double

                                    ' misuro intestazioni (usando etichetta se presente)
                                    For c As Integer = 0 To dt.Columns.Count - 1
                                        Dim colName = dt.Columns(c).ColumnName
                                        Dim hdrText = If(headerLabels.ContainsKey(colName) AndAlso Not String.IsNullOrWhiteSpace(headerLabels(colName)), headerLabels(colName), SpaziaMaiuscole(colName))
                                        maxColWidths(c) = Math.Min(gfx.MeasureString(hdrText, fontHdr).Width + 8, usableWidth)
                                    Next

                                    ' misuro alcune righe per i contenuti (incluso uso di LavorazioneDescrizione se presente)
                                    Dim sampleCount = Math.Min(200, dt.Rows.Count)
                                    For c As Integer = 0 To dt.Columns.Count - 1
                                        For r As Integer = 0 To sampleCount - 1
                                            Dim s = If(dt.Rows(r).IsNull(c), "", dt.Rows(r)(c).ToString())
                                            Dim w = gfx.MeasureString(s, fontCell).Width + 8
                                            If w > maxColWidths(c) Then maxColWidths(c) = Math.Min(w, usableWidth)
                                        Next
                                    Next

                                    ' scala se la somma supera usableWidth
                                    Dim totalDesired = maxColWidths.Sum()
                                    Dim minColWidth As Double = 30
                                    If totalDesired <= 0 Then
                                        For c As Integer = 0 To maxColWidths.Length - 1
                                            maxColWidths(c) = usableWidth / colCount
                                        Next
                                    ElseIf totalDesired > usableWidth Then
                                        Dim scale = usableWidth / totalDesired
                                        For c As Integer = 0 To maxColWidths.Length - 1
                                            maxColWidths(c) = Math.Max(minColWidth, Math.Floor(maxColWidths(c) * scale))
                                        Next
                                        Dim assigned = maxColWidths.Sum()
                                        Dim diff = usableWidth - assigned
                                        Dim idx = 0
                                        While diff > 0
                                            maxColWidths(idx Mod maxColWidths.Length) += 1
                                            diff -= 1
                                            idx += 1
                                        End While
                                    End If

                                    ' header row height
                                    Dim headerHeight As Double = 12
                                    If y + headerHeight > page.Height.Point - margin Then
                                        page = document.AddPage()
                                        gfx = XGraphics.FromPdfPage(page)
                                        tf = New XTextFormatter(gfx)
                                        y = margin
                                    End If

                                    ' disegno header (usa etichette se presenti)
                                    Dim xPos As Double = margin
                                    For c As Integer = 0 To dt.Columns.Count - 1
                                        Dim colName = dt.Columns(c).ColumnName
                                        Dim hdrText = If(headerLabels.ContainsKey(colName) AndAlso Not String.IsNullOrWhiteSpace(headerLabels(colName)), headerLabels(colName), SpaziaMaiuscole(colName))
                                        Dim rectH As New XRect(xPos, y, maxColWidths(c), headerHeight)
                                        gfx.DrawRectangle(XPens.LightGray, rectH)
                                        tf.DrawString(hdrText, fontHdr, XBrushes.DarkBlue, rectH, XStringFormats.TopLeft)
                                        xPos += maxColWidths(c)
                                    Next
                                    y += headerHeight

                                    ' righe: calcolo altezza riga in base al wrapping del testo in ogni cella
                                    For r As Integer = 0 To dt.Rows.Count - 1
                                        Dim estRowHeight As Double = 0
                                        For c As Integer = 0 To dt.Columns.Count - 1
                                            ' se esiste LavorazioneDescrizione e la colonna corrente è una possibile FK, preferisco la descrizione per la stima/l'output
                                            Dim colName = dt.Columns(c).ColumnName
                                            Dim cellTextCandidate As String = If(dt.Rows(r).IsNull(c), "", dt.Rows(r)(c).ToString())
                                            Dim cellText As String = cellTextCandidate
                                            If dt.Columns.Contains("LavorazioneDescrizione") AndAlso (String.Equals(colName, "LavorazioneId", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(colName, "IdLavorazione", StringComparison.OrdinalIgnoreCase) OrElse colName.ToLower().Contains("lavoraz")) Then
                                                Dim desc = If(dt.Rows(r).IsNull("LavorazioneDescrizione"), "", dt.Rows(r)("LavorazioneDescrizione").ToString())
                                                If Not String.IsNullOrWhiteSpace(desc) Then cellText = desc
                                            End If

                                            If String.IsNullOrEmpty(cellText) Then
                                                estRowHeight = Math.Max(estRowHeight, fontCell.Size + 4)
                                            Else
                                                Dim avgCharWidth = gfx.MeasureString("W", fontCell).Width
                                                Dim maxCharsPerLine = Math.Max(10, CInt(Math.Floor(maxColWidths(c) / avgCharWidth)))
                                                Dim lines = Math.Ceiling(cellText.Length / CSng(maxCharsPerLine))
                                                Dim estimated = lines * (fontCell.Size + 2) + 2
                                                estRowHeight = Math.Max(estRowHeight, estimated)
                                            End If
                                        Next

                                        If y + estRowHeight > page.Height.Point - margin Then
                                            page = document.AddPage()
                                            gfx = XGraphics.FromPdfPage(page)
                                            tf = New XTextFormatter(gfx)
                                            y = margin
                                            ' ridisegno header sulla nuova pagina
                                            xPos = margin
                                            For c As Integer = 0 To dt.Columns.Count - 1
                                                Dim colName = dt.Columns(c).ColumnName
                                                Dim hdrText = If(headerLabels.ContainsKey(colName) AndAlso Not String.IsNullOrWhiteSpace(headerLabels(colName)), headerLabels(colName), SpaziaMaiuscole(colName))
                                                Dim rectH As New XRect(xPos, y, maxColWidths(c), headerHeight)
                                                gfx.DrawRectangle(XPens.LightGray, rectH)
                                                tf.DrawString(hdrText, fontHdr, XBrushes.DarkBlue, rectH, XStringFormats.TopLeft)
                                                xPos += maxColWidths(c)
                                            Next
                                            y += headerHeight
                                        End If

                                        xPos = margin
                                        For c As Integer = 0 To dt.Columns.Count - 1
                                            Dim colName = dt.Columns(c).ColumnName
                                            Dim rawText = If(dt.Rows(r).IsNull(c), "", dt.Rows(r)(c).ToString())
                                            Dim cellText = rawText
                                            If dt.Columns.Contains("LavorazioneDescrizione") AndAlso (String.Equals(colName, "LavorazioneId", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(colName, "IdLavorazione", StringComparison.OrdinalIgnoreCase) OrElse colName.ToLower().Contains("lavoraz")) Then
                                                Dim desc = If(dt.Rows(r).IsNull("LavorazioneDescrizione"), "", dt.Rows(r)("LavorazioneDescrizione").ToString())
                                                If Not String.IsNullOrWhiteSpace(desc) Then cellText = desc
                                            End If

                                            Dim rectCell As New XRect(xPos, y, maxColWidths(c), estRowHeight)
                                            If (r Mod 2) = 0 Then gfx.DrawRectangle(XBrushes.WhiteSmoke, rectCell)
                                            tfCell.Alignment = XParagraphAlignment.Left
                                            tfCell.DrawString(cellText, fontCell, XBrushes.Black, rectCell, XStringFormats.TopLeft)
                                            xPos += maxColWidths(c)
                                        Next

                                        y += estRowHeight + 2
                                    Next

                                    y += 8 ' spazio finale dopo la tabella
                                End Sub

                ' Disegna tabelle: animazione e diverse
                DrawTable(dtA, "Lavorazioni - Animazione", "Mov_AssegnazioniLavA")
                DrawTable(dtD, "Lavorazioni - Diverse", "Mov_AssegnazioniLavD")

                ' Footer / Firma
                If y + 100 > page.Height.Point - margin Then
                    page = document.AddPage()
                    gfx = XGraphics.FromPdfPage(page)
                    tf = New XTextFormatter(gfx)
                    y = margin
                End If

                y += 20
                tf.DrawString("Cordiali saluti,", fontNorm, XBrushes.Black, New XRect(margin, y, 300, 18), XStringFormats.TopLeft)
                y += 26
                tf.DrawString("Assistente Produzione", fontBold, XBrushes.Black, New XRect(margin, y, 300, 18), XStringFormats.TopLeft)
                y += 18

                document.Save(filePath)
                MDIMessageBox.Show("PDF creato correttamente:" & Environment.NewLine & filePath, Me.MdiParent, MessageBoxButtons.OK)
            End Using
        Catch ex As Exception
            MDIMessageBox.Show("Errore durante la creazione del PDF: " & ex.Message, Me.MdiParent, MessageBoxButtons.OK)
        End Try
    End Sub

    Private Function SpaziaMaiuscole(text As String) As String
        If String.IsNullOrWhiteSpace(text) Then Return String.Empty

        Dim sb As New System.Text.StringBuilder()
        sb.Append(text(0))

        For i = 1 To text.Length - 1
            Dim c As Char = text(i)
            Dim prev As Char = text(i - 1)

            If Char.IsUpper(c) AndAlso (Char.IsLower(prev) OrElse Char.IsDigit(prev)) Then
                sb.Append(" "c)
            ElseIf prev = "_"c OrElse prev = " "c Then
                If sb(sb.Length - 1) <> " "c Then sb.Append(" "c)
            End If

            sb.Append(c)
        Next

        Return sb.ToString().Trim()
    End Function

    Private Sub CreaPDFAssegnazione_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Try
            SalvaPosizioneForm(Me)
        Catch
        End Try
    End Sub

    Private Sub CreaPDFAssegnazione_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            RipristinaPosizioneForm(Me)
        Catch
        End Try
    End Sub

End Class

