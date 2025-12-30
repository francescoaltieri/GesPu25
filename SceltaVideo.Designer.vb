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
        txtFiltroLavorazione = New TextBox()
        Label1 = New Label()
        ChkDaApprovare = New CheckBox()
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
        dgvRevisioni.Location = New Point(13, 44)
        dgvRevisioni.Name = "dgvRevisioni"
        dgvRevisioni.Size = New Size(914, 424)
        dgvRevisioni.TabIndex = 0
        ' 
        ' txtFiltroLavorazione
        ' 
        txtFiltroLavorazione.Location = New Point(59, 10)
        txtFiltroLavorazione.Name = "txtFiltroLavorazione"
        txtFiltroLavorazione.Size = New Size(394, 23)
        txtFiltroLavorazione.TabIndex = 1
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(13, 13)
        Label1.Name = "Label1"
        Label1.Size = New Size(40, 15)
        Label1.TabIndex = 2
        Label1.Text = "Cerca:"
        ' 
        ' ChkDaApprovare
        ' 
        ChkDaApprovare.AutoSize = True
        ChkDaApprovare.Location = New Point(552, 12)
        ChkDaApprovare.Name = "ChkDaApprovare"
        ChkDaApprovare.Size = New Size(121, 19)
        ChkDaApprovare.TabIndex = 3
        ChkDaApprovare.Text = "Solo da approvare"
        ChkDaApprovare.UseVisualStyleBackColor = True
        ' 
        ' SceltaVideo
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(939, 480)
        Controls.Add(ChkDaApprovare)
        Controls.Add(Label1)
        Controls.Add(txtFiltroLavorazione)
        Controls.Add(dgvRevisioni)
        FormBorderStyle = FormBorderStyle.FixedToolWindow
        Name = "SceltaVideo"
        Text = "Scelta Lavorazione"
        CType(dgvRevisioni, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents dgvRevisioni As DataGridView
    Friend WithEvents txtFiltroLavorazione As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents ChkDaApprovare As CheckBox
End Class
