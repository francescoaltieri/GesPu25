<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SceltaVideo
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
        dgvRevisioni = New DataGridView()
        txtFiltro = New TextBox()
        Label1 = New Label()
        CType(dgvRevisioni, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' dgvRevisioni
        ' 
        dgvRevisioni.AllowUserToAddRows = False
        dgvRevisioni.AllowUserToDeleteRows = False
        dgvRevisioni.AllowUserToResizeColumns = False
        dgvRevisioni.AllowUserToResizeRows = False
        dgvRevisioni.Anchor = AnchorStyles.None
        dgvRevisioni.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        dgvRevisioni.Location = New Point(12, 40)
        dgvRevisioni.Name = "dgvRevisioni"
        dgvRevisioni.Size = New Size(870, 424)
        dgvRevisioni.TabIndex = 0
        ' 
        ' txtFiltro
        ' 
        txtFiltro.Location = New Point(55, 10)
        txtFiltro.Name = "txtFiltro"
        txtFiltro.Size = New Size(219, 23)
        txtFiltro.TabIndex = 1
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(13, 13)
        Label1.Name = "Label1"
        Label1.Size = New Size(37, 15)
        Label1.TabIndex = 2
        Label1.Text = "Video"
        ' 
        ' SceltaVideo
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(894, 474)
        Controls.Add(Label1)
        Controls.Add(txtFiltro)
        Controls.Add(dgvRevisioni)
        Name = "SceltaVideo"
        Text = "Scelta Video da elaborare"
        CType(dgvRevisioni, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents dgvRevisioni As DataGridView
    Friend WithEvents txtFiltro As TextBox
    Friend WithEvents Label1 As Label
End Class
