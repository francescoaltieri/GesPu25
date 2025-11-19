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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        btnAcquisisciPDF = New Button()
        btnAnnulla = New Button()
        picPanel = New PictureBox()
        Label1 = New Label()
        Label2 = New Label()
        txtScena = New TextBox()
        txtPanel = New TextBox()
        Label3 = New Label()
        txtStoryboard = New TextBox()
        btnCaricaStoryboard = New Button()
        btnPrev = New Button()
        btnNext = New Button()
        CType(picPanel, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' btnAcquisisciPDF
        ' 
        btnAcquisisciPDF.Location = New Point(12, 577)
        btnAcquisisciPDF.Name = "btnAcquisisciPDF"
        btnAcquisisciPDF.Size = New Size(75, 23)
        btnAcquisisciPDF.TabIndex = 0
        btnAcquisisciPDF.Text = "Acquisisci PDF"
        btnAcquisisciPDF.UseVisualStyleBackColor = True
        ' 
        ' btnAnnulla
        ' 
        btnAnnulla.Location = New Point(93, 577)
        btnAnnulla.Name = "btnAnnulla"
        btnAnnulla.Size = New Size(75, 23)
        btnAnnulla.TabIndex = 1
        btnAnnulla.Text = "Annulla"
        btnAnnulla.UseVisualStyleBackColor = True
        ' 
        ' picPanel
        ' 
        picPanel.BorderStyle = BorderStyle.Fixed3D
        picPanel.Location = New Point(12, 39)
        picPanel.Name = "picPanel"
        picPanel.Size = New Size(844, 532)
        picPanel.TabIndex = 4
        picPanel.TabStop = False
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(878, 49)
        Label1.Name = "Label1"
        Label1.Size = New Size(41, 15)
        Label1.TabIndex = 5
        Label1.Text = "Scena:"
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(880, 86)
        Label2.Name = "Label2"
        Label2.Size = New Size(39, 15)
        Label2.TabIndex = 6
        Label2.Text = "Panel:"
        ' 
        ' txtScena
        ' 
        txtScena.Location = New Point(925, 46)
        txtScena.Name = "txtScena"
        txtScena.Size = New Size(100, 23)
        txtScena.TabIndex = 7
        ' 
        ' txtPanel
        ' 
        txtPanel.Location = New Point(925, 78)
        txtPanel.Name = "txtPanel"
        txtPanel.Size = New Size(100, 23)
        txtPanel.TabIndex = 8
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
        ' txtStoryboard
        ' 
        txtStoryboard.Location = New Point(86, 6)
        txtStoryboard.Name = "txtStoryboard"
        txtStoryboard.Size = New Size(221, 23)
        txtStoryboard.TabIndex = 10
        ' 
        ' btnCaricaStoryboard
        ' 
        btnCaricaStoryboard.Location = New Point(313, 6)
        btnCaricaStoryboard.Name = "btnCaricaStoryboard"
        btnCaricaStoryboard.Size = New Size(129, 24)
        btnCaricaStoryboard.TabIndex = 11
        btnCaricaStoryboard.Text = "Carica Storybord"
        btnCaricaStoryboard.UseVisualStyleBackColor = True
        ' 
        ' btnPrev
        ' 
        btnPrev.Location = New Point(700, 577)
        btnPrev.Name = "btnPrev"
        btnPrev.Size = New Size(75, 23)
        btnPrev.TabIndex = 12
        btnPrev.Text = "<"
        btnPrev.UseVisualStyleBackColor = True
        ' 
        ' btnNext
        ' 
        btnNext.Location = New Point(781, 577)
        btnNext.Name = "btnNext"
        btnNext.Size = New Size(75, 23)
        btnNext.TabIndex = 13
        btnNext.Text = ">"
        btnNext.UseVisualStyleBackColor = True
        ' 
        ' PDF2Storyboard
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1037, 612)
        Controls.Add(btnNext)
        Controls.Add(btnPrev)
        Controls.Add(btnCaricaStoryboard)
        Controls.Add(txtStoryboard)
        Controls.Add(Label3)
        Controls.Add(txtPanel)
        Controls.Add(txtScena)
        Controls.Add(Label2)
        Controls.Add(Label1)
        Controls.Add(picPanel)
        Controls.Add(btnAnnulla)
        Controls.Add(btnAcquisisciPDF)
        Name = "PDF2Storyboard"
        Text = "Acquisisci da Storyboard"
        CType(picPanel, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents btnAcquisisciPDF As Button
    Friend WithEvents btnAnnulla As Button
    Friend WithEvents picPanel As PictureBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents txtScena As TextBox
    Friend WithEvents txtPanel As TextBox
    Private WithEvents Label3 As Label
    Friend WithEvents txtStoryboard As TextBox
    Friend WithEvents btnCaricaStoryboard As Button
    Friend WithEvents btnPrev As Button
    Friend WithEvents btnNext As Button
End Class
