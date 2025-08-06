Imports System.Text
Imports System.Text.RegularExpressions
Imports Microsoft.Data.SqlClient

Module ModuloCampiDinamici

    ' Conversione tipo SQL → tipo visuale
    Public Function MappaTipoVisuale(sqlType As String) As String
        Select Case sqlType.ToLower()
            Case "bit"
                Return "Boolean"
            Case "date", "datetime"
                Return "Date"
            Case "varchar", "nvarchar", "text"
                Return "String"
            Case Else
                Return "String"
        End Select
    End Function

    ' Recupera l'etichetta (simulazione — personalizzabile)
    Public Function GetEtichetta(tabella As String, nomeCampo As String) As String
        Return nomeCampo.Replace("_", " ") ' Puoi personalizzare o leggere da file config
    End Function

    ' Crea il controllo visivo per ogni campo
    Public Function CreaControllo(campo As CampoDatabase) As Control
        Select Case campo.Tipo.ToLower()
            Case "string"
                Return New TextBox With {.Text = "", .Anchor = AnchorStyles.Left}
            Case "date"
                Return New DateTimePicker With {.Format = DateTimePickerFormat.Short}
            Case "boolean"
                Return New CheckBox With {.Checked = False}
            Case "imgvid"
                Dim pannello As New FlowLayoutPanel With {
                    .FlowDirection = FlowDirection.LeftToRight,
                    .AutoSize = True
                }

                Dim txtPath As New TextBox With {.Width = 200}
                Dim btnVisualizza As New Button With {.Text = "Visualizza"}

                AddHandler btnVisualizza.Click, Sub(s, e)
                                                    MessageBox.Show("Apri form visualizzazione immagine/video qui.", "Anteprima", MessageBoxButtons.OK)
                                                End Sub

                pannello.Controls.Add(txtPath)
                pannello.Controls.Add(btnVisualizza)
                Return pannello

            Case Else
                Return New TextBox With {.Text = "", .Anchor = AnchorStyles.Left}
        End Select
    End Function

    Public Function RipulisciStringa(valore As String) As String

        valore = valore.Trim() ' Rimuove spazi iniziali/finali
        valore = valore.Replace(Chr(0), "") ' Rimuove caratteri null
        valore = Regex.Replace(valore, "[^\u0020-\u007E]", "") ' Rimuove caratteri non ASCII stampabili

        Return valore.ToString()

    End Function

    Public Function SpaziaPrimaDelleMaiuscole(text As String) As String
        If String.IsNullOrWhiteSpace(text) Then Return ""
        Dim sb As New StringBuilder()
        sb.Append(text(0)) ' Mantiene la prima lettera com'è
        For i As Integer = 1 To text.Length - 1
            Dim c = text(i)
            If Char.IsUpper(c) AndAlso Not Char.IsWhiteSpace(text(i - 1)) Then
                sb.Append(" ")
            End If
            sb.Append(c)
        Next
        Return sb.ToString()
    End Function

    Public Function RecuperaCampiDa(nomeTabella As String) As List(Of CampoDatabase)
        Dim campi As New List(Of CampoDatabase)
        Dim nomeChiave As String = ""

        Try
            Using conn As New SqlConnection(ConnString)
                conn.Open()

                ' Recupera tutti i campi con info su tipo e identity
                Dim queryCampi As String = "
                SELECT 
                    c.COLUMN_NAME, 
                    c.DATA_TYPE,
                    COLUMNPROPERTY(OBJECT_ID(c.TABLE_NAME), c.COLUMN_NAME, 'IsIdentity') AS IsIdentity
                FROM INFORMATION_SCHEMA.COLUMNS c
                WHERE c.TABLE_NAME = @nomeTabella
            "

                Using cmdCampi As New SqlCommand(queryCampi, conn)
                    cmdCampi.Parameters.AddWithValue("@nomeTabella", nomeTabella)

                    Using reader = cmdCampi.ExecuteReader()
                        While reader.Read()
                            Dim nomeCampo As String = reader("COLUMN_NAME").ToString()
                            Dim tipoCampo As String = reader("DATA_TYPE").ToString()
                            Dim isIdentity As Boolean = Convert.ToInt32(reader("IsIdentity")) = 1

                            campi.Add(New CampoDatabase With {
                            .Nome = nomeCampo,
                            .Tipo = MappaTipoVisuale(tipoCampo),
                            .IsChiave = False,
                            .IsIdentity = isIdentity
                        })
                        End While
                    End Using
                End Using

                ' Mappa i campi ImgVid come tipo visuale "imgvid"
                For Each campo In campi
                    If campo.Nome.StartsWith("ImgVid", StringComparison.OrdinalIgnoreCase) Then
                        campo.Tipo = "imgvid"
                    End If
                Next

                ' 3. Recupera la chiave primaria
                Dim queryChiave As String = "
                SELECT ccu.COLUMN_NAME 
                FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc
                JOIN INFORMATION_SCHEMA.CONSTRAINT_COLUMN_USAGE ccu
                    ON tc.CONSTRAINT_NAME = ccu.CONSTRAINT_NAME
                WHERE tc.TABLE_NAME = @nomeTabella AND tc.CONSTRAINT_TYPE = 'PRIMARY KEY'
            "

                Using cmdChiave As New SqlCommand(queryChiave, conn)
                    cmdChiave.Parameters.AddWithValue("@nomeTabella", nomeTabella)

                    Using readerChiave = cmdChiave.ExecuteReader()
                        If readerChiave.Read() Then
                            nomeChiave = readerChiave("COLUMN_NAME").ToString()
                        End If
                    End Using
                End Using

                For Each campo In campi
                    If String.Equals(campo.Nome, nomeChiave, StringComparison.OrdinalIgnoreCase) Then
                        campo.IsChiave = True
                        Exit For
                    End If
                Next

                ' Integra i collegamenti da Sys_CollegamentiCampi
                Dim queryCollegamenti = "
                SELECT NomeCampo, TabellaCollegata, CampoValore, CampoVisuale
                FROM Sys_CollegamentiCampi
                WHERE NomeTabella = @nomeTabella
            "

                Using cmdCollegamenti As New SqlCommand(queryCollegamenti, conn)
                    cmdCollegamenti.Parameters.AddWithValue("@nomeTabella", nomeTabella)

                    Using reader = cmdCollegamenti.ExecuteReader()
                        While reader.Read()
                            Dim nomeCampo = reader("NomeCampo").ToString()
                            Dim campo = campi.FirstOrDefault(Function(c) c.Nome.Equals(nomeCampo, StringComparison.OrdinalIgnoreCase))
                            If campo IsNot Nothing Then
                                campo.TabellaCollegata = If(IsDBNull(reader("TabellaCollegata")), Nothing, reader("TabellaCollegata").ToString())
                                campo.CampoValore = If(IsDBNull(reader("CampoValore")), Nothing, reader("CampoValore").ToString())
                                campo.CampoVisuale = If(IsDBNull(reader("CampoVisuale")), Nothing, reader("CampoVisuale").ToString())
                            Else
                                MessageBox.Show("Collegamento ignorato: campo " & nomeCampo & " non trovato nella tabella " & nomeTabella, "Messaggio", MessageBoxButtons.OK)
                            End If
                        End While
                    End Using
                End Using

            End Using

        Catch ex As Exception
            MessageBox.Show("Errore nel recupero dei campi: " & ex.Message, "Errore", MessageBoxButtons.OK)
        End Try

        Return campi
    End Function

End Module

