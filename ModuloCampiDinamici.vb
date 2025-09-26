Imports System.Text
Imports System.Text.RegularExpressions
Imports Microsoft.Data.SqlClient

Module ModuloCampiDinamici

    ' Conversione tipo SQL → tipo visuale
    Private Function MappaTipoVisuale(tipoSql As String) As String
        Select Case tipoSql.ToLower()
            Case "bit", "bit(max)" : Return "boolean"
            Case "int", "bigint", "smallint", "tinyint", "int(max)" : Return "int"
            Case "decimal", "numeric", "money", "float", "real" : Return "decimal"
            Case "date", "date(max)", "datetime", "datetime(max)", "smalldatetime", "datetime2" : Return "date"
            Case "varchar", "nvarchar", "char", "nchar", "text", "ntext" : Return "string"
            Case "varchar(max)", "nvarchar(max)", "char(max)", "nchar(max)", "text(max)", "ntext(max)" : Return "string_max"
            Case "uniqueidentifier" : Return "guid"
            Case "money(max)" : Return "money"
            Case Else : Return "string000" ' campo non riconosciuto
        End Select
    End Function

    Public Function GetEtichetta(tabella As String, nomeCampo As String) As String
        Return nomeCampo.Replace("_", " ")
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

                Dim queryCampi As String = "
                SELECT 
                    c.COLUMN_NAME, 
                    c.DATA_TYPE,
                    c.CHARACTER_MAXIMUM_LENGTH,
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
                            Dim maxLen As Integer = If(IsDBNull(reader("CHARACTER_MAXIMUM_LENGTH")), -1, Convert.ToInt32(reader("CHARACTER_MAXIMUM_LENGTH")))
                            Dim CampoLungo As String

                            If maxLen = -1 Then
                                CampoLungo = "(max)"
                            Else
                                CampoLungo = ""
                            End If

                            campi.Add(New CampoDatabase With {
                            .Nome = nomeCampo,
                            .Tipo = MappaTipoVisuale(tipoCampo & CampoLungo),
                            .MaxLen = maxLen,
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

