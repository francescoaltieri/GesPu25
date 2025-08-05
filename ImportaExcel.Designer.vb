<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ImportaExcel
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
        lstFogli = New ListBox()
        lstTabelle = New ListBox()
        btnImporta = New Button()
        btnCarica = New Button()
        lblNomeFileExcel = New Label()
        Label1 = New Label()
        Label2 = New Label()
        Label3 = New Label()
        lstCampi = New ListBox()
        Label4 = New Label()
        lstColonne = New ListBox()
        Label5 = New Label()
        lstCollegamenti = New ListBox()
        SuspendLayout()
        ' 
        ' lstFogli
        ' 
        lstFogli.BackColor = SystemColors.InactiveCaption
        lstFogli.FormattingEnabled = True
        lstFogli.ItemHeight = 15
        lstFogli.Location = New Point(16, 64)
        lstFogli.Name = "lstFogli"
        lstFogli.Size = New Size(156, 229)
        lstFogli.TabIndex = 0
        ' 
        ' lstTabelle
        ' 
        lstTabelle.BackColor = SystemColors.InactiveCaption
        lstTabelle.FormattingEnabled = True
        lstTabelle.ItemHeight = 15
        lstTabelle.Location = New Point(184, 64)
        lstTabelle.Name = "lstTabelle"
        lstTabelle.Size = New Size(156, 229)
        lstTabelle.TabIndex = 1
        ' 
        ' btnImporta
        ' 
        btnImporta.Location = New Point(692, 88)
        btnImporta.Name = "btnImporta"
        btnImporta.Size = New Size(124, 23)
        btnImporta.TabIndex = 2
        btnImporta.Text = "Importa"
        btnImporta.UseVisualStyleBackColor = True
        ' 
        ' btnCarica
        ' 
        btnCarica.Location = New Point(692, 64)
        btnCarica.Name = "btnCarica"
        btnCarica.Size = New Size(124, 23)
        btnCarica.TabIndex = 3
        btnCarica.Text = "Seleziona File"
        btnCarica.UseVisualStyleBackColor = True
        ' 
        ' lblNomeFileExcel
        ' 
        lblNomeFileExcel.AutoSize = True
        lblNomeFileExcel.Font = New Font("Segoe UI", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        lblNomeFileExcel.ForeColor = SystemColors.HotTrack
        lblNomeFileExcel.Location = New Point(20, 12)
        lblNomeFileExcel.Name = "lblNomeFileExcel"
        lblNomeFileExcel.Size = New Size(183, 25)
        lblNomeFileExcel.TabIndex = 4
        lblNomeFileExcel.Text = "Nome del File Excel"
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Font = New Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label1.Location = New Point(20, 44)
        Label1.Name = "Label1"
        Label1.Size = New Size(135, 15)
        Label1.TabIndex = 5
        Label1.Text = "Lista Fogli nel File Excel"
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Font = New Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label2.Location = New Point(184, 44)
        Label2.Name = "Label2"
        Label2.Size = New Size(148, 15)
        Label2.TabIndex = 6
        Label2.Text = "Lista Tabelle nel Database"
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Font = New Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label3.Location = New Point(356, 44)
        Label3.Name = "Label3"
        Label3.Size = New Size(139, 15)
        Label3.TabIndex = 8
        Label3.Text = "Lista Campi della Tabella"
        ' 
        ' lstCampi
        ' 
        lstCampi.BackColor = SystemColors.InactiveCaption
        lstCampi.FormattingEnabled = True
        lstCampi.ItemHeight = 15
        lstCampi.Location = New Point(356, 64)
        lstCampi.Name = "lstCampi"
        lstCampi.Size = New Size(156, 229)
        lstCampi.TabIndex = 7
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.Font = New Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label4.Location = New Point(524, 44)
        Label4.Name = "Label4"
        Label4.Size = New Size(152, 15)
        Label4.TabIndex = 10
        Label4.Text = "Intestazioni Colonne Excel"
        ' 
        ' lstColonne
        ' 
        lstColonne.BackColor = SystemColors.InactiveCaption
        lstColonne.FormattingEnabled = True
        lstColonne.ItemHeight = 15
        lstColonne.Location = New Point(524, 64)
        lstColonne.Name = "lstColonne"
        lstColonne.Size = New Size(156, 229)
        lstColonne.TabIndex = 9
        ' 
        ' Label5
        ' 
        Label5.AutoSize = True
        Label5.Font = New Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label5.Location = New Point(20, 304)
        Label5.Name = "Label5"
        Label5.Size = New Size(108, 15)
        Label5.TabIndex = 12
        Label5.Text = "Lista Collegamenti"
        ' 
        ' lstCollegamenti
        ' 
        lstCollegamenti.BackColor = SystemColors.InactiveCaption
        lstCollegamenti.FormattingEnabled = True
        lstCollegamenti.ItemHeight = 15
        lstCollegamenti.Location = New Point(16, 328)
        lstCollegamenti.Name = "lstCollegamenti"
        lstCollegamenti.Size = New Size(664, 229)
        lstCollegamenti.TabIndex = 11
        ' 
        ' ImportaExcel
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(822, 583)
        Controls.Add(Label5)
        Controls.Add(lstCollegamenti)
        Controls.Add(Label4)
        Controls.Add(lstColonne)
        Controls.Add(Label3)
        Controls.Add(lstCampi)
        Controls.Add(Label2)
        Controls.Add(Label1)
        Controls.Add(lblNomeFileExcel)
        Controls.Add(btnCarica)
        Controls.Add(btnImporta)
        Controls.Add(lstTabelle)
        Controls.Add(lstFogli)
        Name = "ImportaExcel"
        Text = "ImportaExcel"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents lstFogli As ListBox
    Friend WithEvents lstTabelle As ListBox
    Friend WithEvents btnImporta As Button
    Friend WithEvents btnCarica As Button
    Friend WithEvents lblNomeFileExcel As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents lstCampi As ListBox
    Friend WithEvents Label4 As Label
    Friend WithEvents lstColonne As ListBox
    Friend WithEvents Label5 As Label
    Friend WithEvents lstCollegamenti As ListBox
End Class
