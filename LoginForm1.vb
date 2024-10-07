Public Class LoginForm1
    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        If txbx_user.Text = "1" And txbx_pas.Text = "1" Then
            Welding.Visible = True
            Me.Visible = False
        Else
            MsgBox("Wrong user name or Password. Please try again.", MsgBoxStyle.Exclamation, "Error")
        End If
    End Sub
    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub
End Class
