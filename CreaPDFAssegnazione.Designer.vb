<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CreaPDFAssegnazione
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
        btnCreaPDF = New Button()
        btnAnnulla = New Button()
        lblCreaPDF = New Label()
        SuspendLayout()
        ' 
        ' btnCreaPDF
        ' 
        btnCreaPDF.Location = New Point(52, 51)
        btnCreaPDF.Name = "btnCreaPDF"
        btnCreaPDF.Size = New Size(75, 23)
        btnCreaPDF.TabIndex = 0
        btnCreaPDF.Text = "Crea PDF"
        btnCreaPDF.UseVisualStyleBackColor = True
        ' 
        ' btnAnnulla
        ' 
        btnAnnulla.Location = New Point(143, 51)
        btnAnnulla.Name = "btnAnnulla"
        btnAnnulla.Size = New Size(75, 23)
        btnAnnulla.TabIndex = 1
        btnAnnulla.Text = "Annulla"
        btnAnnulla.UseVisualStyleBackColor = True
        ' 
        ' lblCreaPDF
        ' 
        lblCreaPDF.AutoSize = True
        lblCreaPDF.Font = New Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        lblCreaPDF.Location = New Point(40, 19)
        lblCreaPDF.Name = "lblCreaPDF"
        lblCreaPDF.Size = New Size(193, 17)
        lblCreaPDF.TabIndex = 2
        lblCreaPDF.Text = "CREAZIONE PDF Assegnazioni"
        ' 
        ' CreaPDFAssegnazione
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(272, 95)
        Controls.Add(lblCreaPDF)
        Controls.Add(btnAnnulla)
        Controls.Add(btnCreaPDF)
        Name = "CreaPDFAssegnazione"
        Text = "Crea PDF Assegnazione"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents btnCreaPDF As Button
    Friend WithEvents btnAnnulla As Button
    Friend WithEvents lblCreaPDF As Label
End Class
