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
        btnPrimoFrame = New Button()
        btnUltimoFrame = New Button()
        btnAvantiVeloce = New Button()
        btnIndietroVeloce = New Button()
        TrackFrame = New TrackBar()
        txtNote = New TextBox()
        GroupBox1 = New GroupBox()
        btnSalvaNote = New Button()
        numSpessorePennino = New NumericUpDown()
        colorDialogPennino = New ColorDialog()
        Label1 = New Label()
        GroupBox2 = New GroupBox()
        Label8 = New Label()
        Label7 = New Label()
        PenColor = New PictureBox()
        CmbStrumento = New ComboBox()
        ChkZoom = New CheckBox()
        btnCaricaRevisione = New Button()
        GroupBox3 = New GroupBox()
        Label4 = New Label()
        Label3 = New Label()
        lstUtentiCondivisi = New CheckedListBox()
        lblDataNota = New Label()
        lblAutore = New Label()
        btnRetake = New Button()
        lblRevAttiva = New Label()
        btnApprovazione = New Button()
        Label2 = New Label()
        GroupBox4 = New GroupBox()
        LstNoteFrame = New ListView()
        Frame = New ColumnHeader()
        Nota = New ColumnHeader()
        Utente = New ColumnHeader()
        BtnRipristinaFrame = New Button()
        Label5 = New Label()
        lblFrameAttivo = New Label()
        lblTotFrames = New Label()
        panelViewer = New Panel()
        Label6 = New Label()
        CType(picFrame, ComponentModel.ISupportInitialize).BeginInit()
        CType(TrackFrame, ComponentModel.ISupportInitialize).BeginInit()
        GroupBox1.SuspendLayout()
        CType(numSpessorePennino, ComponentModel.ISupportInitialize).BeginInit()
        GroupBox2.SuspendLayout()
        CType(PenColor, ComponentModel.ISupportInitialize).BeginInit()
        GroupBox3.SuspendLayout()
        GroupBox4.SuspendLayout()
        panelViewer.SuspendLayout()
        SuspendLayout()
        ' 
        ' btnCaricaVideo
        ' 
        btnCaricaVideo.Location = New Point(1089, 5)
        btnCaricaVideo.Name = "btnCaricaVideo"
        btnCaricaVideo.Size = New Size(145, 34)
        btnCaricaVideo.TabIndex = 0
        btnCaricaVideo.Text = "Carica Nuovo Lavoro"
        btnCaricaVideo.UseVisualStyleBackColor = True
        ' 
        ' btnSalvaVideo
        ' 
        btnSalvaVideo.Location = New Point(1314, 828)
        btnSalvaVideo.Name = "btnSalvaVideo"
        btnSalvaVideo.Size = New Size(259, 42)
        btnSalvaVideo.TabIndex = 2
        btnSalvaVideo.Text = "Esporta Video"
        btnSalvaVideo.UseVisualStyleBackColor = True
        ' 
        ' picFrame
        ' 
        picFrame.BorderStyle = BorderStyle.Fixed3D
        picFrame.Location = New Point(3, 3)
        picFrame.Name = "picFrame"
        picFrame.Size = New Size(1283, 727)
        picFrame.TabIndex = 3
        picFrame.TabStop = False
        ' 
        ' btnPrecedente
        ' 
        btnPrecedente.Location = New Point(138, 827)
        btnPrecedente.Name = "btnPrecedente"
        btnPrecedente.Size = New Size(127, 41)
        btnPrecedente.TabIndex = 4
        btnPrecedente.Text = "⟨ Frame Precedente"
        btnPrecedente.UseVisualStyleBackColor = True
        ' 
        ' btnSuccessivo
        ' 
        btnSuccessivo.Location = New Point(271, 827)
        btnSuccessivo.Name = "btnSuccessivo"
        btnSuccessivo.Size = New Size(127, 41)
        btnSuccessivo.TabIndex = 5
        btnSuccessivo.Text = "Frame Successivo ⟩"
        btnSuccessivo.UseVisualStyleBackColor = True
        ' 
        ' OpenFileDialog1
        ' 
        OpenFileDialog1.FileName = "OpenFileDialog1"
        ' 
        ' btnAnnulla
        ' 
        btnAnnulla.Location = New Point(1045, 828)
        btnAnnulla.Name = "btnAnnulla"
        btnAnnulla.Size = New Size(120, 42)
        btnAnnulla.TabIndex = 10
        btnAnnulla.Text = "Annulla Ultima Modifica sul Frame"
        btnAnnulla.UseVisualStyleBackColor = True
        ' 
        ' btnPrimoFrame
        ' 
        btnPrimoFrame.Location = New Point(12, 826)
        btnPrimoFrame.Name = "btnPrimoFrame"
        btnPrimoFrame.Size = New Size(120, 41)
        btnPrimoFrame.TabIndex = 12
        btnPrimoFrame.Text = "│⟨ Primo Frame"
        btnPrimoFrame.UseVisualStyleBackColor = True
        ' 
        ' btnUltimoFrame
        ' 
        btnUltimoFrame.Location = New Point(404, 826)
        btnUltimoFrame.Name = "btnUltimoFrame"
        btnUltimoFrame.Size = New Size(120, 41)
        btnUltimoFrame.TabIndex = 13
        btnUltimoFrame.Text = "Ultimo Frame ⟩│"
        btnUltimoFrame.UseVisualStyleBackColor = True
        ' 
        ' btnAvantiVeloce
        ' 
        btnAvantiVeloce.Location = New Point(706, 827)
        btnAvantiVeloce.Name = "btnAvantiVeloce"
        btnAvantiVeloce.Size = New Size(120, 41)
        btnAvantiVeloce.TabIndex = 14
        btnAvantiVeloce.Text = "Avanti Veloce ⟩⟩⟩"
        btnAvantiVeloce.UseVisualStyleBackColor = True
        ' 
        ' btnIndietroVeloce
        ' 
        btnIndietroVeloce.Location = New Point(580, 827)
        btnIndietroVeloce.Name = "btnIndietroVeloce"
        btnIndietroVeloce.Size = New Size(120, 41)
        btnIndietroVeloce.TabIndex = 15
        btnIndietroVeloce.Text = "⟨⟨⟨ Indietro Veloce"
        btnIndietroVeloce.UseVisualStyleBackColor = True
        ' 
        ' TrackFrame
        ' 
        TrackFrame.Location = New Point(12, 791)
        TrackFrame.Name = "TrackFrame"
        TrackFrame.Size = New Size(1299, 45)
        TrackFrame.TabIndex = 16
        ' 
        ' txtNote
        ' 
        txtNote.Location = New Point(8, 22)
        txtNote.Multiline = True
        txtNote.Name = "txtNote"
        txtNote.Size = New Size(245, 180)
        txtNote.TabIndex = 17
        ' 
        ' GroupBox1
        ' 
        GroupBox1.Controls.Add(btnSalvaNote)
        GroupBox1.Controls.Add(txtNote)
        GroupBox1.Location = New Point(1317, 238)
        GroupBox1.Name = "GroupBox1"
        GroupBox1.Size = New Size(259, 239)
        GroupBox1.TabIndex = 18
        GroupBox1.TabStop = False
        GroupBox1.Text = "Annotazione Frame"
        ' 
        ' btnSalvaNote
        ' 
        btnSalvaNote.Location = New Point(6, 208)
        btnSalvaNote.Name = "btnSalvaNote"
        btnSalvaNote.Size = New Size(245, 25)
        btnSalvaNote.TabIndex = 19
        btnSalvaNote.Text = "Salva Annotazione"
        btnSalvaNote.UseVisualStyleBackColor = True
        ' 
        ' numSpessorePennino
        ' 
        numSpessorePennino.Enabled = False
        numSpessorePennino.Location = New Point(143, 56)
        numSpessorePennino.Maximum = New Decimal(New Integer() {50, 0, 0, 0})
        numSpessorePennino.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        numSpessorePennino.Name = "numSpessorePennino"
        numSpessorePennino.Size = New Size(108, 23)
        numSpessorePennino.TabIndex = 19
        numSpessorePennino.Value = New Decimal(New Integer() {5, 0, 0, 0})
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(8, 61)
        Label1.Name = "Label1"
        Label1.Size = New Size(73, 15)
        Label1.TabIndex = 21
        Label1.Text = "Dimensione:"
        ' 
        ' GroupBox2
        ' 
        GroupBox2.Controls.Add(Label8)
        GroupBox2.Controls.Add(Label7)
        GroupBox2.Controls.Add(PenColor)
        GroupBox2.Controls.Add(CmbStrumento)
        GroupBox2.Controls.Add(ChkZoom)
        GroupBox2.Controls.Add(Label1)
        GroupBox2.Controls.Add(numSpessorePennino)
        GroupBox2.Location = New Point(1317, 704)
        GroupBox2.Name = "GroupBox2"
        GroupBox2.Size = New Size(259, 118)
        GroupBox2.TabIndex = 22
        GroupBox2.TabStop = False
        GroupBox2.Text = "Pennino"
        ' 
        ' Label8
        ' 
        Label8.AutoSize = True
        Label8.Location = New Point(8, 86)
        Label8.Name = "Label8"
        Label8.Size = New Size(45, 15)
        Label8.TabIndex = 39
        Label8.Text = "Colore:"
        ' 
        ' Label7
        ' 
        Label7.AutoSize = True
        Label7.Location = New Point(8, 30)
        Label7.Name = "Label7"
        Label7.Size = New Size(129, 15)
        Label7.TabIndex = 38
        Label7.Text = "Strumento Selezionato:"
        ' 
        ' PenColor
        ' 
        PenColor.BorderStyle = BorderStyle.Fixed3D
        PenColor.Location = New Point(143, 86)
        PenColor.Name = "PenColor"
        PenColor.Size = New Size(108, 21)
        PenColor.TabIndex = 37
        PenColor.TabStop = False
        ' 
        ' CmbStrumento
        ' 
        CmbStrumento.Enabled = False
        CmbStrumento.FormattingEnabled = True
        CmbStrumento.Location = New Point(143, 27)
        CmbStrumento.Name = "CmbStrumento"
        CmbStrumento.Size = New Size(108, 23)
        CmbStrumento.TabIndex = 22
        ' 
        ' ChkZoom
        ' 
        ChkZoom.AutoSize = True
        ChkZoom.Location = New Point(153, 0)
        ChkZoom.Name = "ChkZoom"
        ChkZoom.Size = New Size(106, 19)
        ChkZoom.TabIndex = 36
        ChkZoom.Text = "Abilita Disegno"
        ChkZoom.UseVisualStyleBackColor = True
        ' 
        ' btnCaricaRevisione
        ' 
        btnCaricaRevisione.BackgroundImageLayout = ImageLayout.Stretch
        btnCaricaRevisione.Location = New Point(938, 5)
        btnCaricaRevisione.Name = "btnCaricaRevisione"
        btnCaricaRevisione.Size = New Size(145, 34)
        btnCaricaRevisione.TabIndex = 23
        btnCaricaRevisione.Text = "Lavori in corso"
        btnCaricaRevisione.UseVisualStyleBackColor = True
        ' 
        ' GroupBox3
        ' 
        GroupBox3.Controls.Add(Label4)
        GroupBox3.Controls.Add(Label3)
        GroupBox3.Controls.Add(lstUtentiCondivisi)
        GroupBox3.Controls.Add(lblDataNota)
        GroupBox3.Controls.Add(lblAutore)
        GroupBox3.Location = New Point(1317, 483)
        GroupBox3.Name = "GroupBox3"
        GroupBox3.Size = New Size(259, 215)
        GroupBox3.TabIndex = 25
        GroupBox3.TabStop = False
        GroupBox3.Text = "Autore e Condivisioni"
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.Location = New Point(8, 39)
        Label4.Name = "Label4"
        Label4.Size = New Size(34, 15)
        Label4.TabIndex = 4
        Label4.Text = "Data:"
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(8, 19)
        Label3.Name = "Label3"
        Label3.Size = New Size(46, 15)
        Label3.TabIndex = 3
        Label3.Text = "Autore:"
        ' 
        ' lstUtentiCondivisi
        ' 
        lstUtentiCondivisi.FormattingEnabled = True
        lstUtentiCondivisi.Location = New Point(6, 57)
        lstUtentiCondivisi.Name = "lstUtentiCondivisi"
        lstUtentiCondivisi.Size = New Size(247, 148)
        lstUtentiCondivisi.TabIndex = 2
        ' 
        ' lblDataNota
        ' 
        lblDataNota.AutoSize = True
        lblDataNota.Location = New Point(63, 39)
        lblDataNota.Name = "lblDataNota"
        lblDataNota.Size = New Size(31, 15)
        lblDataNota.TabIndex = 1
        lblDataNota.Text = "Data"
        ' 
        ' lblAutore
        ' 
        lblAutore.AutoSize = True
        lblAutore.Location = New Point(63, 19)
        lblAutore.Name = "lblAutore"
        lblAutore.Size = New Size(43, 15)
        lblAutore.TabIndex = 0
        lblAutore.Text = "Autore"
        ' 
        ' btnRetake
        ' 
        btnRetake.BackgroundImageLayout = ImageLayout.Stretch
        btnRetake.Location = New Point(1317, 5)
        btnRetake.Name = "btnRetake"
        btnRetake.Size = New Size(122, 34)
        btnRetake.TabIndex = 26
        btnRetake.Text = "Retake"
        btnRetake.UseVisualStyleBackColor = True
        ' 
        ' lblRevAttiva
        ' 
        lblRevAttiva.AutoSize = True
        lblRevAttiva.Font = New Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        lblRevAttiva.ForeColor = SystemColors.HotTrack
        lblRevAttiva.Location = New Point(89, 13)
        lblRevAttiva.Name = "lblRevAttiva"
        lblRevAttiva.Size = New Size(74, 17)
        lblRevAttiva.TabIndex = 27
        lblRevAttiva.Text = "Rev. Attiva"
        ' 
        ' btnApprovazione
        ' 
        btnApprovazione.BackgroundImageLayout = ImageLayout.Stretch
        btnApprovazione.Location = New Point(1448, 5)
        btnApprovazione.Name = "btnApprovazione"
        btnApprovazione.Size = New Size(122, 34)
        btnApprovazione.TabIndex = 28
        btnApprovazione.Text = "Approvazione"
        btnApprovazione.UseVisualStyleBackColor = True
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Font = New Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label2.ForeColor = SystemColors.HotTrack
        Label2.Location = New Point(12, 13)
        Label2.Name = "Label2"
        Label2.Size = New Size(71, 17)
        Label2.TabIndex = 29
        Label2.Text = "Revisione:"
        ' 
        ' GroupBox4
        ' 
        GroupBox4.Controls.Add(LstNoteFrame)
        GroupBox4.Location = New Point(1317, 45)
        GroupBox4.Name = "GroupBox4"
        GroupBox4.Size = New Size(259, 187)
        GroupBox4.TabIndex = 30
        GroupBox4.TabStop = False
        GroupBox4.Text = "Lista Annotazioni"
        ' 
        ' LstNoteFrame
        ' 
        LstNoteFrame.Columns.AddRange(New ColumnHeader() {Frame, Nota, Utente})
        LstNoteFrame.Location = New Point(6, 22)
        LstNoteFrame.Name = "LstNoteFrame"
        LstNoteFrame.Size = New Size(247, 159)
        LstNoteFrame.TabIndex = 26
        LstNoteFrame.UseCompatibleStateImageBehavior = False
        ' 
        ' BtnRipristinaFrame
        ' 
        BtnRipristinaFrame.Location = New Point(1171, 828)
        BtnRipristinaFrame.Name = "BtnRipristinaFrame"
        BtnRipristinaFrame.Size = New Size(120, 42)
        BtnRipristinaFrame.TabIndex = 31
        BtnRipristinaFrame.Text = "Ripristina Frame Originale"
        BtnRipristinaFrame.UseVisualStyleBackColor = True
        ' 
        ' Label5
        ' 
        Label5.AutoSize = True
        Label5.Font = New Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label5.ForeColor = SystemColors.HotTrack
        Label5.Location = New Point(629, 13)
        Label5.Name = "Label5"
        Label5.Size = New Size(68, 17)
        Label5.TabIndex = 32
        Label5.Text = "N. Frame:"
        ' 
        ' lblFrameAttivo
        ' 
        lblFrameAttivo.AutoSize = True
        lblFrameAttivo.BackColor = SystemColors.Control
        lblFrameAttivo.Font = New Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        lblFrameAttivo.ForeColor = SystemColors.HotTrack
        lblFrameAttivo.Location = New Point(703, 13)
        lblFrameAttivo.Name = "lblFrameAttivo"
        lblFrameAttivo.Size = New Size(0, 17)
        lblFrameAttivo.TabIndex = 33
        lblFrameAttivo.TextAlign = ContentAlignment.MiddleRight
        ' 
        ' lblTotFrames
        ' 
        lblTotFrames.AutoSize = True
        lblTotFrames.Font = New Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        lblTotFrames.ForeColor = SystemColors.HotTrack
        lblTotFrames.Location = New Point(854, 13)
        lblTotFrames.Name = "lblTotFrames"
        lblTotFrames.Size = New Size(0, 17)
        lblTotFrames.TabIndex = 35
        lblTotFrames.TextAlign = ContentAlignment.MiddleLeft
        ' 
        ' panelViewer
        ' 
        panelViewer.AutoScroll = True
        panelViewer.Controls.Add(picFrame)
        panelViewer.Location = New Point(12, 40)
        panelViewer.Name = "panelViewer"
        panelViewer.Size = New Size(1289, 733)
        panelViewer.TabIndex = 37
        ' 
        ' Label6
        ' 
        Label6.AutoSize = True
        Label6.Font = New Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label6.ForeColor = SystemColors.HotTrack
        Label6.Location = New Point(770, 13)
        Label6.Name = "Label6"
        Label6.Size = New Size(78, 17)
        Label6.TabIndex = 38
        Label6.Text = "Tot. Frame:"
        ' 
        ' VideoFBF
        ' 
        AllowDrop = True
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        AutoScroll = True
        ClientSize = New Size(1585, 879)
        Controls.Add(Label6)
        Controls.Add(panelViewer)
        Controls.Add(lblTotFrames)
        Controls.Add(lblFrameAttivo)
        Controls.Add(Label5)
        Controls.Add(BtnRipristinaFrame)
        Controls.Add(GroupBox4)
        Controls.Add(Label2)
        Controls.Add(btnApprovazione)
        Controls.Add(lblRevAttiva)
        Controls.Add(btnRetake)
        Controls.Add(GroupBox3)
        Controls.Add(btnCaricaRevisione)
        Controls.Add(GroupBox2)
        Controls.Add(GroupBox1)
        Controls.Add(btnIndietroVeloce)
        Controls.Add(btnAvantiVeloce)
        Controls.Add(btnUltimoFrame)
        Controls.Add(btnPrimoFrame)
        Controls.Add(btnAnnulla)
        Controls.Add(btnSuccessivo)
        Controls.Add(btnPrecedente)
        Controls.Add(btnSalvaVideo)
        Controls.Add(btnCaricaVideo)
        Controls.Add(TrackFrame)
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
        CType(PenColor, ComponentModel.ISupportInitialize).EndInit()
        GroupBox3.ResumeLayout(False)
        GroupBox3.PerformLayout()
        GroupBox4.ResumeLayout(False)
        panelViewer.ResumeLayout(False)
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
    Friend WithEvents btnPrimoFrame As Button
    Friend WithEvents btnUltimoFrame As Button
    Friend WithEvents btnAvantiVeloce As Button
    Friend WithEvents btnIndietroVeloce As Button
    Friend WithEvents TrackFrame As TrackBar
    Friend WithEvents txtNote As TextBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents numSpessorePennino As NumericUpDown
    Friend WithEvents colorDialogPennino As ColorDialog
    Friend WithEvents Label1 As Label
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents btnSalvaNote As Button
    Friend WithEvents btnCaricaRevisione As Button
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents lblDataNota As Label
    Friend WithEvents lblAutore As Label
    Friend WithEvents btnRetake As Button
    Friend WithEvents lblRevAttiva As Label
    Friend WithEvents btnApprovazione As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents lstUtentiCondivisi As CheckedListBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents BtnRipristinaFrame As Button
    Friend WithEvents LstNoteFrame As ListView
    Friend WithEvents Label5 As Label
    Friend WithEvents lblFrameAttivo As Label
    Friend WithEvents lblTotFrames As Label
    Friend WithEvents Frame As ColumnHeader
    Friend WithEvents Nota As ColumnHeader
    Friend WithEvents Utente As ColumnHeader
    Friend WithEvents ChkZoom As CheckBox
    Friend WithEvents panelViewer As Panel
    Friend WithEvents CmbStrumento As ComboBox
    Friend WithEvents PenColor As PictureBox
    Friend WithEvents Label7 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents Label6 As Label

End Class
