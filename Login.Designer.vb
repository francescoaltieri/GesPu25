<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Login
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
        BtnAnnulla = New Button()
        BtnLogin = New Button()
        txtPassword = New TextBox()
        Label2 = New Label()
        txtNomeUtente = New TextBox()
        Label1 = New Label()
        SuspendLayout()
        ' 
        ' BtnAnnulla
        ' 
        BtnAnnulla.Location = New Point(160, 80)
        BtnAnnulla.Name = "BtnAnnulla"
        BtnAnnulla.Size = New Size(96, 24)
        BtnAnnulla.TabIndex = 52
        BtnAnnulla.Text = "Esci"
        BtnAnnulla.UseVisualStyleBackColor = True
        ' 
        ' BtnLogin
        ' 
        BtnLogin.Location = New Point(52, 80)
        BtnLogin.Name = "BtnLogin"
        BtnLogin.Size = New Size(96, 24)
        BtnLogin.TabIndex = 51
        BtnLogin.Text = "Login"
        BtnLogin.UseVisualStyleBackColor = True
        ' 
        ' txtPassword
        ' 
        txtPassword.Location = New Point(88, 40)
        txtPassword.Name = "txtPassword"
        txtPassword.PasswordChar = "*"c
        txtPassword.Size = New Size(208, 23)
        txtPassword.TabIndex = 50
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(12, 48)
        Label2.Name = "Label2"
        Label2.Size = New Size(60, 15)
        Label2.TabIndex = 49
        Label2.Text = "Password:"
        ' 
        ' txtNomeUtente
        ' 
        txtNomeUtente.Location = New Point(88, 12)
        txtNomeUtente.Name = "txtNomeUtente"
        txtNomeUtente.Size = New Size(144, 23)
        txtNomeUtente.TabIndex = 48
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(12, 16)
        Label1.Name = "Label1"
        Label1.Size = New Size(58, 15)
        Label1.TabIndex = 47
        Label1.Text = "Id Utente:"
        ' 
        ' Login
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(307, 113)
        Controls.Add(BtnAnnulla)
        Controls.Add(BtnLogin)
        Controls.Add(txtPassword)
        Controls.Add(Label2)
        Controls.Add(txtNomeUtente)
        Controls.Add(Label1)
        Name = "Login"
        StartPosition = FormStartPosition.CenterParent
        Text = "Login"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents BtnAnnulla As Button
    Friend WithEvents BtnLogin As Button
    Friend WithEvents txtPassword As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents txtNomeUtente As TextBox
    Friend WithEvents Label1 As Label
End Class
