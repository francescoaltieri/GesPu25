Public Enum ExportType
    None
    PDF
    Excel
End Enum
Public Class ExportChoiceForm
    Inherits Form

    Public Property SelectedExportType As ExportType = ExportType.None

    Private rbPdf As RadioButton
    Private rbExcel As RadioButton
    Private btnOK As Button
    Private btnCancel As Button

    Public Sub New()
        Me.InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "Scegli tipo di esportazione"
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.StartPosition = FormStartPosition.CenterParent
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.ShowInTaskbar = False
        Me.ClientSize = New Drawing.Size(300, 140)

        rbPdf = New RadioButton() With {
            .Text = "PDF",
            .Location = New Drawing.Point(20, 20),
            .AutoSize = True,
            .Checked = True
        }

        rbExcel = New RadioButton() With {
            .Text = "Excel",
            .Location = New Drawing.Point(20, 50),
            .AutoSize = True
        }

        btnOK = New Button() With {
            .Text = "OK",
            .DialogResult = DialogResult.OK,
            .Location = New Drawing.Point(110, 90),
            .Size = New Drawing.Size(75, 25)
        }
        AddHandler btnOK.Click, AddressOf BtnOK_Click

        btnCancel = New Button() With {
            .Text = "Annulla",
            .DialogResult = DialogResult.Cancel,
            .Location = New Drawing.Point(195, 90),
            .Size = New Drawing.Size(75, 25)
        }

        Me.Controls.Add(rbPdf)
        Me.Controls.Add(rbExcel)
        Me.Controls.Add(btnOK)
        Me.Controls.Add(btnCancel)

        Me.AcceptButton = btnOK
        Me.CancelButton = btnCancel
    End Sub

    Private Sub BtnOK_Click(sender As Object, e As EventArgs)
        If rbPdf.Checked Then
            SelectedExportType = ExportType.PDF
        ElseIf rbExcel.Checked Then
            SelectedExportType = ExportType.Excel
        Else
            SelectedExportType = ExportType.None
        End If
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub
End Class