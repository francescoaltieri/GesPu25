<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CreaLettAssegnazione
    Inherits System.Windows.Forms.Form

    'Form esegue l'override del metodo Dispose per pulire l'elenco dei componenti.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Richiesto da Progettazione Windows Form
    Private components As System.ComponentModel.IContainer

    'NOTA: la procedura che segue è richiesta da Progettazione Windows Form
    'Può essere modificata in Progettazione Windows Form.  
    'Non modificarla mediante l'editor del codice.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        btnCreaLett = New Button()
        btnChiudi = New Button()
        lblCreaPDF = New Label()
        btnCaricaAllegati = New Button()
        SuspendLayout()
        ' 
        ' btnCreaLett
        ' 
        btnCreaLett.Location = New Point(28, 56)
        btnCreaLett.Name = "btnCreaLett"
        btnCreaLett.Size = New Size(102, 23)
        btnCreaLett.TabIndex = 1
        btnCreaLett.Text = "Crea Lettera"
        btnCreaLett.UseVisualStyleBackColor = True
        ' 
        ' btnChiudi
        ' 
        btnChiudi.Location = New Point(92, 96)
        btnChiudi.Name = "btnChiudi"
        btnChiudi.Size = New Size(75, 23)
        btnChiudi.TabIndex = 3
        btnChiudi.Text = "Chiudi"
        btnChiudi.UseVisualStyleBackColor = True
        ' 
        ' lblCreaPDF
        ' 
        lblCreaPDF.AutoSize = True
        lblCreaPDF.Font = New Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        lblCreaPDF.Location = New Point(51, 20)
        lblCreaPDF.Name = "lblCreaPDF"
        lblCreaPDF.Size = New Size(167, 17)
        lblCreaPDF.TabIndex = 0
        lblCreaPDF.Text = "Assegnazione Lavorazioni"
        ' 
        ' btnCaricaAllegati
        ' 
        btnCaricaAllegati.Location = New Point(137, 56)
        btnCaricaAllegati.Name = "btnCaricaAllegati"
        btnCaricaAllegati.Size = New Size(102, 23)
        btnCaricaAllegati.TabIndex = 2
        btnCaricaAllegati.Text = "Carica Allegati"
        btnCaricaAllegati.UseVisualStyleBackColor = True
        ' 
        ' CreaLettAssegnazione
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(272, 130)
        Controls.Add(btnCaricaAllegati)
        Controls.Add(lblCreaPDF)
        Controls.Add(btnChiudi)
        Controls.Add(btnCreaLett)
        Name = "CreaLettAssegnazione"
        Text = "Crea Lettera Assegnazione"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents btnCreaLett As Button
    Friend WithEvents btnChiudi As Button
    Friend WithEvents lblCreaPDF As Label
    Friend WithEvents btnCaricaAllegati As Button
End Class
