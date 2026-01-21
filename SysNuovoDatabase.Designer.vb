<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SysNuovoDatabase
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
        Label1 = New Label()
        txtNomeNuovoDatabase = New TextBox()
        btnCreaNuovoDatabase = New Button()
        btnAnnulla = New Button()
        lblStato = New Label()
        ProgressBarCopia = New ProgressBar()
        SuspendLayout()
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(12, 21)
        Label1.Name = "Label1"
        Label1.Size = New Size(266, 15)
        Label1.TabIndex = 0
        Label1.Text = "Immettere il nome del nuovo Database da creare:"
        ' 
        ' txtNomeNuovoDatabase
        ' 
        txtNomeNuovoDatabase.Location = New Point(284, 18)
        txtNomeNuovoDatabase.Name = "txtNomeNuovoDatabase"
        txtNomeNuovoDatabase.Size = New Size(243, 23)
        txtNomeNuovoDatabase.TabIndex = 1
        ' 
        ' btnCreaNuovoDatabase
        ' 
        btnCreaNuovoDatabase.Location = New Point(185, 122)
        btnCreaNuovoDatabase.Name = "btnCreaNuovoDatabase"
        btnCreaNuovoDatabase.Size = New Size(75, 23)
        btnCreaNuovoDatabase.TabIndex = 2
        btnCreaNuovoDatabase.Text = "Crea"
        btnCreaNuovoDatabase.UseVisualStyleBackColor = True
        ' 
        ' btnAnnulla
        ' 
        btnAnnulla.Location = New Point(284, 122)
        btnAnnulla.Name = "btnAnnulla"
        btnAnnulla.Size = New Size(75, 23)
        btnAnnulla.TabIndex = 3
        btnAnnulla.Text = "Annulla"
        btnAnnulla.UseVisualStyleBackColor = True
        ' 
        ' lblStato
        ' 
        lblStato.AutoSize = True
        lblStato.Location = New Point(12, 93)
        lblStato.Name = "lblStato"
        lblStato.Size = New Size(0, 15)
        lblStato.TabIndex = 4
        ' 
        ' ProgressBarCopia
        ' 
        ProgressBarCopia.Location = New Point(12, 54)
        ProgressBarCopia.Name = "ProgressBarCopia"
        ProgressBarCopia.Size = New Size(515, 23)
        ProgressBarCopia.TabIndex = 5
        ' 
        ' SysNuovoDatabase
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(539, 155)
        Controls.Add(ProgressBarCopia)
        Controls.Add(lblStato)
        Controls.Add(btnAnnulla)
        Controls.Add(btnCreaNuovoDatabase)
        Controls.Add(txtNomeNuovoDatabase)
        Controls.Add(Label1)
        Name = "SysNuovoDatabase"
        Text = "Crea Nuovo Database"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents txtNomeNuovoDatabase As TextBox
    Friend WithEvents btnCreaNuovoDatabase As Button
    Friend WithEvents btnAnnulla As Button
    Friend WithEvents lblStato As Label
    Friend WithEvents ProgressBarCopia As ProgressBar
End Class
