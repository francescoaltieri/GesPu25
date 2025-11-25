<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PDF2Storyboard
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        BtnAcquisisciPDF = New Button()
        BtnChiudi = New Button()
        PicPanel = New PictureBox()
        Label3 = New Label()
        BtnCaricaStoryboard = New Button()
        BtnPrev = New Button()
        BtnNext = New Button()
        BtnAnnunllaModifica = New Button()
        BtnSalvaPanel = New Button()
        ComboStoryboard = New ComboBox()
        BtnCancellaPanel = New Button()
        CheckConfermaSalvataggio = New CheckBox()
        BtnAnimatic = New Button()
        BtnAssegnaScene = New Button()
        CType(PicPanel, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' BtnAcquisisciPDF
        ' 
        BtnAcquisisciPDF.Location = New Point(12, 531)
        BtnAcquisisciPDF.Name = "BtnAcquisisciPDF"
        BtnAcquisisciPDF.Size = New Size(98, 43)
        BtnAcquisisciPDF.TabIndex = 0
        BtnAcquisisciPDF.Text = "Acquisizione Panels da PDF"
        BtnAcquisisciPDF.UseVisualStyleBackColor = True
        ' 
        ' BtnChiudi
        ' 
        BtnChiudi.Location = New Point(781, 531)
        BtnChiudi.Name = "BtnChiudi"
        BtnChiudi.Size = New Size(75, 43)
        BtnChiudi.TabIndex = 1
        BtnChiudi.Text = "Chiudi"
        BtnChiudi.UseVisualStyleBackColor = True
        ' 
        ' PicPanel
        ' 
        PicPanel.BorderStyle = BorderStyle.Fixed3D
        PicPanel.Location = New Point(12, 39)
        PicPanel.Name = "PicPanel"
        PicPanel.Size = New Size(844, 486)
        PicPanel.TabIndex = 4
        PicPanel.TabStop = False
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(12, 9)
        Label3.Name = "Label3"
        Label3.Size = New Size(68, 15)
        Label3.TabIndex = 9
        Label3.Text = "Storyboard:"
        ' 
        ' BtnCaricaStoryboard
        ' 
        BtnCaricaStoryboard.Location = New Point(468, 4)
        BtnCaricaStoryboard.Name = "BtnCaricaStoryboard"
        BtnCaricaStoryboard.Size = New Size(154, 29)
        BtnCaricaStoryboard.TabIndex = 11
        BtnCaricaStoryboard.Text = "Carica Panels dal Drive"
        BtnCaricaStoryboard.UseVisualStyleBackColor = True
        ' 
        ' BtnPrev
        ' 
        BtnPrev.Location = New Point(630, 531)
        BtnPrev.Name = "BtnPrev"
        BtnPrev.Size = New Size(75, 23)
        BtnPrev.TabIndex = 12
        BtnPrev.Text = "<"
        BtnPrev.UseVisualStyleBackColor = True
        ' 
        ' BtnNext
        ' 
        BtnNext.Location = New Point(698, 531)
        BtnNext.Name = "BtnNext"
        BtnNext.Size = New Size(75, 23)
        BtnNext.TabIndex = 13
        BtnNext.Text = ">"
        BtnNext.UseVisualStyleBackColor = True
        ' 
        ' BtnAnnunllaModifica
        ' 
        BtnAnnunllaModifica.Location = New Point(394, 531)
        BtnAnnunllaModifica.Name = "BtnAnnunllaModifica"
        BtnAnnunllaModifica.Size = New Size(69, 43)
        BtnAnnunllaModifica.TabIndex = 14
        BtnAnnunllaModifica.Text = "Annulla Modifiche"
        BtnAnnunllaModifica.UseVisualStyleBackColor = True
        ' 
        ' BtnSalvaPanel
        ' 
        BtnSalvaPanel.Location = New Point(469, 531)
        BtnSalvaPanel.Name = "BtnSalvaPanel"
        BtnSalvaPanel.Size = New Size(69, 43)
        BtnSalvaPanel.TabIndex = 15
        BtnSalvaPanel.Text = "Salva Modifiche"
        BtnSalvaPanel.UseVisualStyleBackColor = True
        ' 
        ' ComboStoryboard
        ' 
        ComboStoryboard.FormattingEnabled = True
        ComboStoryboard.Location = New Point(86, 7)
        ComboStoryboard.Name = "ComboStoryboard"
        ComboStoryboard.Size = New Size(376, 23)
        ComboStoryboard.TabIndex = 16
        ' 
        ' BtnCancellaPanel
        ' 
        BtnCancellaPanel.Location = New Point(544, 531)
        BtnCancellaPanel.Name = "BtnCancellaPanel"
        BtnCancellaPanel.Size = New Size(69, 43)
        BtnCancellaPanel.TabIndex = 18
        BtnCancellaPanel.Text = "Cancella Panel"
        BtnCancellaPanel.UseVisualStyleBackColor = True
        ' 
        ' CheckConfermaSalvataggio
        ' 
        CheckConfermaSalvataggio.AutoSize = True
        CheckConfermaSalvataggio.Location = New Point(630, 555)
        CheckConfermaSalvataggio.Name = "CheckConfermaSalvataggio"
        CheckConfermaSalvataggio.Size = New Size(143, 19)
        CheckConfermaSalvataggio.TabIndex = 19
        CheckConfermaSalvataggio.Text = "Conferma Salvataggio"
        CheckConfermaSalvataggio.UseVisualStyleBackColor = True
        ' 
        ' BtnAnimatic
        ' 
        BtnAnimatic.Location = New Point(116, 531)
        BtnAnimatic.Name = "BtnAnimatic"
        BtnAnimatic.Size = New Size(98, 43)
        BtnAnimatic.TabIndex = 20
        BtnAnimatic.Text = "Crea Animatic"
        BtnAnimatic.UseVisualStyleBackColor = True
        ' 
        ' BtnAssegnaScene
        ' 
        BtnAssegnaScene.Location = New Point(258, 531)
        BtnAssegnaScene.Name = "BtnAssegnaScene"
        BtnAssegnaScene.Size = New Size(98, 43)
        BtnAssegnaScene.TabIndex = 21
        BtnAssegnaScene.Text = "Assegna Scene"
        BtnAssegnaScene.UseVisualStyleBackColor = True
        ' 
        ' PDF2Storyboard
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(865, 586)
        Controls.Add(BtnAssegnaScene)
        Controls.Add(BtnAnimatic)
        Controls.Add(CheckConfermaSalvataggio)
        Controls.Add(BtnCancellaPanel)
        Controls.Add(ComboStoryboard)
        Controls.Add(BtnSalvaPanel)
        Controls.Add(BtnAnnunllaModifica)
        Controls.Add(BtnNext)
        Controls.Add(BtnPrev)
        Controls.Add(BtnCaricaStoryboard)
        Controls.Add(Label3)
        Controls.Add(PicPanel)
        Controls.Add(BtnChiudi)
        Controls.Add(BtnAcquisisciPDF)
        Name = "PDF2Storyboard"
        Text = "Acquisisci Panel Storyboard da PDF"
        CType(PicPanel, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents BtnAcquisisciPDF As Button
    Friend WithEvents BtnChiudi As Button
    Friend WithEvents PicPanel As PictureBox
    Private WithEvents Label3 As Label
    Friend WithEvents BtnCaricaStoryboard As Button
    Friend WithEvents BtnPrev As Button
    Friend WithEvents BtnNext As Button
    Friend WithEvents BtnAnnunllaModifica As Button
    Friend WithEvents BtnSalvaPanel As Button
    Friend WithEvents ComboStoryboard As ComboBox
    Friend WithEvents BtnCancellaPanel As Button
    Friend WithEvents CheckConfermaSalvataggio As CheckBox
    Friend WithEvents BtnAnimatic As Button
    Friend WithEvents BtnAssegnaScene As Button

End Class
