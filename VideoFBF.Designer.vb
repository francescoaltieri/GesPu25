<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class VideoFBF
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        btnCaricaVideo = New Button()
        btnSalvaVideo = New Button()
        picFrame = New PictureBox()
        btnPrecedente = New Button()
        btnSuccessivo = New Button()
        OpenFileDialog1 = New OpenFileDialog()
        btnAnnulla = New Button()
        btnSalvaFrame = New Button()
        btnPrimoFrame = New Button()
        btnUltimoFrame = New Button()
        btnAvantiVeloce = New Button()
        btnIndietroVeloce = New Button()
        TrackFrame = New TrackBar()
        txtNote = New TextBox()
        GroupBox1 = New GroupBox()
        btnSalvaNote = New Button()
        btnAggiungiNote = New Button()
        numSpessorePennino = New NumericUpDown()
        colorDialogPennino = New ColorDialog()
        Label1 = New Label()
        GroupBox2 = New GroupBox()
        btnColorePennino = New Button()
        btnCaricaRevisione = New Button()
        lstNoteFrame = New ListBox()
        GroupBox3 = New GroupBox()
        lblDataNota = New Label()
        lblAutore = New Label()
        btnNuovaRevisione = New Button()
        lblRevAttiva = New Label()
        CType(picFrame, ComponentModel.ISupportInitialize).BeginInit()
        CType(TrackFrame, ComponentModel.ISupportInitialize).BeginInit()
        GroupBox1.SuspendLayout()
        CType(numSpessorePennino, ComponentModel.ISupportInitialize).BeginInit()
        GroupBox2.SuspendLayout()
        GroupBox3.SuspendLayout()
        SuspendLayout()
        ' 
        ' btnCaricaVideo
        ' 
        btnCaricaVideo.Location = New Point(1425, 12)
        btnCaricaVideo.Name = "btnCaricaVideo"
        btnCaricaVideo.Size = New Size(145, 23)
        btnCaricaVideo.TabIndex = 0
        btnCaricaVideo.Text = "Carica Nuovo Video"
        btnCaricaVideo.UseVisualStyleBackColor = True
        ' 
        ' btnSalvaVideo
        ' 
        btnSalvaVideo.Location = New Point(1450, 792)
        btnSalvaVideo.Name = "btnSalvaVideo"
        btnSalvaVideo.Size = New Size(120, 42)
        btnSalvaVideo.TabIndex = 2
        btnSalvaVideo.Text = "Salva Video"
        btnSalvaVideo.UseVisualStyleBackColor = True
        ' 
        ' picFrame
        ' 
        picFrame.BorderStyle = BorderStyle.Fixed3D
        picFrame.Location = New Point(12, 40)
        picFrame.Name = "picFrame"
        picFrame.Size = New Size(1299, 695)
        picFrame.TabIndex = 3
        picFrame.TabStop = False
        ' 
        ' btnPrecedente
        ' 
        btnPrecedente.Location = New Point(12, 792)
        btnPrecedente.Name = "btnPrecedente"
        btnPrecedente.Size = New Size(120, 41)
        btnPrecedente.TabIndex = 4
        btnPrecedente.Text = "Frame Precedente"
        btnPrecedente.UseVisualStyleBackColor = True
        ' 
        ' btnSuccessivo
        ' 
        btnSuccessivo.Location = New Point(138, 792)
        btnSuccessivo.Name = "btnSuccessivo"
        btnSuccessivo.Size = New Size(120, 41)
        btnSuccessivo.TabIndex = 5
        btnSuccessivo.Text = "Frame Successivo"
        btnSuccessivo.UseVisualStyleBackColor = True
        ' 
        ' OpenFileDialog1
        ' 
        OpenFileDialog1.FileName = "OpenFileDialog1"
        ' 
        ' btnAnnulla
        ' 
        btnAnnulla.Location = New Point(1323, 792)
        btnAnnulla.Name = "btnAnnulla"
        btnAnnulla.Size = New Size(120, 42)
        btnAnnulla.TabIndex = 10
        btnAnnulla.Text = "Annulla Ultima Modifica"
        btnAnnulla.UseVisualStyleBackColor = True
        ' 
        ' btnSalvaFrame
        ' 
        btnSalvaFrame.Location = New Point(1197, 792)
        btnSalvaFrame.Name = "btnSalvaFrame"
        btnSalvaFrame.Size = New Size(120, 42)
        btnSalvaFrame.TabIndex = 11
        btnSalvaFrame.Text = "Salva Frame"
        btnSalvaFrame.UseVisualStyleBackColor = True
        ' 
        ' btnPrimoFrame
        ' 
        btnPrimoFrame.Location = New Point(273, 793)
        btnPrimoFrame.Name = "btnPrimoFrame"
        btnPrimoFrame.Size = New Size(120, 41)
        btnPrimoFrame.TabIndex = 12
        btnPrimoFrame.Text = "Primo Frame"
        btnPrimoFrame.UseVisualStyleBackColor = True
        ' 
        ' btnUltimoFrame
        ' 
        btnUltimoFrame.Location = New Point(399, 793)
        btnUltimoFrame.Name = "btnUltimoFrame"
        btnUltimoFrame.Size = New Size(120, 41)
        btnUltimoFrame.TabIndex = 13
        btnUltimoFrame.Text = "Ultimo Frame"
        btnUltimoFrame.UseVisualStyleBackColor = True
        ' 
        ' btnAvantiVeloce
        ' 
        btnAvantiVeloce.Location = New Point(538, 793)
        btnAvantiVeloce.Name = "btnAvantiVeloce"
        btnAvantiVeloce.Size = New Size(120, 41)
        btnAvantiVeloce.TabIndex = 14
        btnAvantiVeloce.Text = "Avanti Veloce"
        btnAvantiVeloce.UseVisualStyleBackColor = True
        ' 
        ' btnIndietroVeloce
        ' 
        btnIndietroVeloce.Location = New Point(664, 793)
        btnIndietroVeloce.Name = "btnIndietroVeloce"
        btnIndietroVeloce.Size = New Size(120, 41)
        btnIndietroVeloce.TabIndex = 15
        btnIndietroVeloce.Text = "Indietro Veloce"
        btnIndietroVeloce.UseVisualStyleBackColor = True
        ' 
        ' TrackFrame
        ' 
        TrackFrame.Location = New Point(12, 741)
        TrackFrame.Name = "TrackFrame"
        TrackFrame.Size = New Size(1299, 45)
        TrackFrame.TabIndex = 16
        ' 
        ' txtNote
        ' 
        txtNote.Location = New Point(8, 22)
        txtNote.Multiline = True
        txtNote.Name = "txtNote"
        txtNote.Size = New Size(245, 166)
        txtNote.TabIndex = 17
        ' 
        ' GroupBox1
        ' 
        GroupBox1.Controls.Add(btnSalvaNote)
        GroupBox1.Controls.Add(btnAggiungiNote)
        GroupBox1.Controls.Add(txtNote)
        GroupBox1.Location = New Point(1317, 433)
        GroupBox1.Name = "GroupBox1"
        GroupBox1.Size = New Size(259, 244)
        GroupBox1.TabIndex = 18
        GroupBox1.TabStop = False
        GroupBox1.Text = "Annotazioni"
        ' 
        ' btnSalvaNote
        ' 
        btnSalvaNote.Location = New Point(133, 194)
        btnSalvaNote.Name = "btnSalvaNote"
        btnSalvaNote.Size = New Size(120, 42)
        btnSalvaNote.TabIndex = 19
        btnSalvaNote.Text = "Salva Note"
        btnSalvaNote.UseVisualStyleBackColor = True
        ' 
        ' btnAggiungiNote
        ' 
        btnAggiungiNote.Location = New Point(8, 194)
        btnAggiungiNote.Name = "btnAggiungiNote"
        btnAggiungiNote.Size = New Size(120, 42)
        btnAggiungiNote.TabIndex = 18
        btnAggiungiNote.Text = "Scrivi Note su Frame"
        btnAggiungiNote.UseVisualStyleBackColor = True
        ' 
        ' numSpessorePennino
        ' 
        numSpessorePennino.Location = New Point(133, 38)
        numSpessorePennino.Maximum = New Decimal(New Integer() {50, 0, 0, 0})
        numSpessorePennino.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        numSpessorePennino.Name = "numSpessorePennino"
        numSpessorePennino.Size = New Size(120, 23)
        numSpessorePennino.TabIndex = 19
        numSpessorePennino.Value = New Decimal(New Integer() {5, 0, 0, 0})
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(8, 40)
        Label1.Name = "Label1"
        Label1.Size = New Size(117, 15)
        Label1.TabIndex = 21
        Label1.Text = "Dimensione Pennino"
        ' 
        ' GroupBox2
        ' 
        GroupBox2.Controls.Add(btnColorePennino)
        GroupBox2.Controls.Add(Label1)
        GroupBox2.Controls.Add(numSpessorePennino)
        GroupBox2.Location = New Point(1317, 312)
        GroupBox2.Name = "GroupBox2"
        GroupBox2.Size = New Size(259, 115)
        GroupBox2.TabIndex = 22
        GroupBox2.TabStop = False
        GroupBox2.Text = "Pennino"
        ' 
        ' btnColorePennino
        ' 
        btnColorePennino.Location = New Point(8, 67)
        btnColorePennino.Name = "btnColorePennino"
        btnColorePennino.Size = New Size(245, 27)
        btnColorePennino.TabIndex = 21
        btnColorePennino.Text = "Scegli Colore"
        btnColorePennino.UseVisualStyleBackColor = True
        ' 
        ' btnCaricaRevisione
        ' 
        btnCaricaRevisione.BackgroundImageLayout = ImageLayout.Stretch
        btnCaricaRevisione.Location = New Point(1015, 12)
        btnCaricaRevisione.Name = "btnCaricaRevisione"
        btnCaricaRevisione.Size = New Size(145, 23)
        btnCaricaRevisione.TabIndex = 23
        btnCaricaRevisione.Text = "Lista Revisione"
        btnCaricaRevisione.UseVisualStyleBackColor = True
        ' 
        ' lstNoteFrame
        ' 
        lstNoteFrame.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Right
        lstNoteFrame.BorderStyle = BorderStyle.None
        lstNoteFrame.FormattingEnabled = True
        lstNoteFrame.ItemHeight = 15
        lstNoteFrame.Location = New Point(1317, 40)
        lstNoteFrame.Name = "lstNoteFrame"
        lstNoteFrame.Size = New Size(255, 270)
        lstNoteFrame.TabIndex = 24
        ' 
        ' GroupBox3
        ' 
        GroupBox3.Controls.Add(lblDataNota)
        GroupBox3.Controls.Add(lblAutore)
        GroupBox3.Location = New Point(1317, 683)
        GroupBox3.Name = "GroupBox3"
        GroupBox3.Size = New Size(259, 89)
        GroupBox3.TabIndex = 25
        GroupBox3.TabStop = False
        GroupBox3.Text = "Autore Modifica"
        ' 
        ' lblDataNota
        ' 
        lblDataNota.AutoSize = True
        lblDataNota.Location = New Point(11, 58)
        lblDataNota.Name = "lblDataNota"
        lblDataNota.Size = New Size(31, 15)
        lblDataNota.TabIndex = 1
        lblDataNota.Text = "Data"
        ' 
        ' lblAutore
        ' 
        lblAutore.AutoSize = True
        lblAutore.Location = New Point(11, 25)
        lblAutore.Name = "lblAutore"
        lblAutore.Size = New Size(43, 15)
        lblAutore.TabIndex = 0
        lblAutore.Text = "Autore"
        ' 
        ' btnNuovaRevisione
        ' 
        btnNuovaRevisione.BackgroundImageLayout = ImageLayout.Stretch
        btnNuovaRevisione.Location = New Point(1166, 12)
        btnNuovaRevisione.Name = "btnNuovaRevisione"
        btnNuovaRevisione.Size = New Size(145, 23)
        btnNuovaRevisione.TabIndex = 26
        btnNuovaRevisione.Text = "Aggiungi Revisione"
        btnNuovaRevisione.UseVisualStyleBackColor = True
        ' 
        ' lblRevAttiva
        ' 
        lblRevAttiva.AutoSize = True
        lblRevAttiva.Font = New Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        lblRevAttiva.ForeColor = SystemColors.HotTrack
        lblRevAttiva.Location = New Point(12, 12)
        lblRevAttiva.Name = "lblRevAttiva"
        lblRevAttiva.Size = New Size(108, 17)
        lblRevAttiva.TabIndex = 27
        lblRevAttiva.Text = "Revisione Attiva"
        ' 
        ' VideoFBF
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1584, 845)
        Controls.Add(lblRevAttiva)
        Controls.Add(btnNuovaRevisione)
        Controls.Add(GroupBox3)
        Controls.Add(lstNoteFrame)
        Controls.Add(btnCaricaRevisione)
        Controls.Add(GroupBox2)
        Controls.Add(GroupBox1)
        Controls.Add(TrackFrame)
        Controls.Add(btnIndietroVeloce)
        Controls.Add(btnAvantiVeloce)
        Controls.Add(btnUltimoFrame)
        Controls.Add(btnPrimoFrame)
        Controls.Add(btnSalvaFrame)
        Controls.Add(btnAnnulla)
        Controls.Add(btnSuccessivo)
        Controls.Add(btnPrecedente)
        Controls.Add(picFrame)
        Controls.Add(btnSalvaVideo)
        Controls.Add(btnCaricaVideo)
        FormBorderStyle = FormBorderStyle.Fixed3D
        Name = "VideoFBF"
        SizeGripStyle = SizeGripStyle.Hide
        StartPosition = FormStartPosition.CenterParent
        Text = "Modifica Frame"
        CType(picFrame, ComponentModel.ISupportInitialize).EndInit()
        CType(TrackFrame, ComponentModel.ISupportInitialize).EndInit()
        GroupBox1.ResumeLayout(False)
        GroupBox1.PerformLayout()
        CType(numSpessorePennino, ComponentModel.ISupportInitialize).EndInit()
        GroupBox2.ResumeLayout(False)
        GroupBox2.PerformLayout()
        GroupBox3.ResumeLayout(False)
        GroupBox3.PerformLayout()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents btnCaricaVideo As Button
    Friend WithEvents btnSalvaVideo As Button
    Friend WithEvents picFrame As PictureBox
    Friend WithEvents btnPrecedente As Button
    Friend WithEvents btnSuccessivo As Button
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents btnAnnulla As Button
    Friend WithEvents btnSalvaFrame As Button
    Friend WithEvents btnPrimoFrame As Button
    Friend WithEvents btnUltimoFrame As Button
    Friend WithEvents btnAvantiVeloce As Button
    Friend WithEvents btnIndietroVeloce As Button
    Friend WithEvents TrackFrame As TrackBar
    Friend WithEvents txtNote As TextBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents btnAggiungiNote As Button
    Friend WithEvents numSpessorePennino As NumericUpDown
    Friend WithEvents colorDialogPennino As ColorDialog
    Friend WithEvents Label1 As Label
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents btnColorePennino As Button
    Friend WithEvents btnSalvaNote As Button
    Friend WithEvents btnCaricaRevisione As Button
    Friend WithEvents lstNoteFrame As ListBox
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents lblDataNota As Label
    Friend WithEvents lblAutore As Label
    Friend WithEvents btnNuovaRevisione As Button
    Friend WithEvents lblRevAttiva As Label

End Class
