Public Class InformazioniApp
    Private Sub InformazioniApp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Informazioni sull'app"
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.StartPosition = FormStartPosition.CenterParent

        GestioneStatoForm.CaricaStato(Me)

        ' Etichette principali
        Dim lblTitolo As New Label With {
            .Text = "GesPu25 - Gestione Produzione Cinematografica",
            .Font = New Font("Segoe UI", 12, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(20, 20)
        }

        Dim lblVersione As New Label With {
            .Text = My.Application.Info.Version.ToString(),
            .AutoSize = True,
            .Location = New Point(20, 50)
        }

        Dim lblAutore As New Label With {
            .Text = "Autore: Francesco Altieri",
            .AutoSize = True,
            .Location = New Point(20, 70)
        }

        Dim txtDescrizione As New TextBox With {
            .Multiline = True,
            .ReadOnly = True,
            .Width = 300,
            .Height = 100,
            .Location = New Point(20, 100),
            .Text = "Questa applicazione consente di gestire ..."
        }

        Dim linkContatto As New LinkLabel With {
            .Text = "francesco.altieri@medinet.it",
            .Location = New Point(20, 210),
            .AutoSize = True
        }

        AddHandler linkContatto.LinkClicked, Sub(s, args)
                                                 Process.Start("mailto:francesco.altieri@medinet.it")
                                             End Sub

        ' Aggiunta controlli al form
        Me.Controls.Add(lblTitolo)
        Me.Controls.Add(lblVersione)
        Me.Controls.Add(lblAutore)
        Me.Controls.Add(txtDescrizione)
        Me.Controls.Add(linkContatto)
    End Sub

    Private Sub InformazioniApp_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        GestioneStatoForm.SalvaStato(Me)
    End Sub
End Class