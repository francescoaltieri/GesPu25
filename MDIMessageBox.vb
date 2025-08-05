Public Class MDIMessageBox
    Public Shared Function Show(message As String, mdiForm As Form, buttons As MessageBoxButtons, Optional title As String = "Avviso") As DialogResult
        Dim box As New Form With {
            .FormBorderStyle = FormBorderStyle.FixedDialog,
            .StartPosition = FormStartPosition.Manual,
            .ShowInTaskbar = False,
            .Text = title,
            .Width = 320,
            .Height = 160,
            .TopMost = True
        }

        ' Posizionamento al centro del MDI form
        Dim x = mdiForm.Left + (mdiForm.Width - box.Width) / 2
        Dim y = mdiForm.Top + (mdiForm.Height - box.Height) / 2
        box.Location = New Point(x, y)

        ' Messaggio
        Dim lbl = New Label With {
            .Text = message,
            .Dock = DockStyle.Top,
            .Height = 80,
            .TextAlign = ContentAlignment.MiddleCenter
        }

        ' Pannello bottoni
        Dim pnlBottoni As New FlowLayoutPanel With {
            .FlowDirection = FlowDirection.RightToLeft,
            .Dock = DockStyle.Bottom,
            .Height = 40
        }

        Dim result As DialogResult = DialogResult.None

        ' Crea bottoni in base a MessageBoxButtons
        Dim creaBottone = Sub(text As String, dialogResult As DialogResult)
                              Dim btn = New Button With {.Text = text, .AutoSize = True, .Margin = New Padding(10)}
                              AddHandler btn.Click, Sub()
                                                        result = dialogResult
                                                        box.Close()
                                                    End Sub
                              pnlBottoni.Controls.Add(btn)
                          End Sub

        Select Case buttons
            Case MessageBoxButtons.OK
                creaBottone("OK", DialogResult.OK)
            Case MessageBoxButtons.YesNo
                creaBottone("No", DialogResult.No)
                creaBottone("Sì", DialogResult.Yes)
            Case MessageBoxButtons.OKCancel
                creaBottone("Annulla", DialogResult.Cancel)
                creaBottone("OK", DialogResult.OK)
                ' Puoi aggiungere altri casi qui
        End Select

        box.Controls.Add(lbl)
        box.Controls.Add(pnlBottoni)

        box.ShowDialog(mdiForm)
        Return result
    End Function
End Class
