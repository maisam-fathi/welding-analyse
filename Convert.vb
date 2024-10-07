Public Class Convert
    Dim Swap As Double
    Private Sub But_close_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles But_close.Click
        Me.Visible = False
        Welding.Opacity = 100%
        Welding.Enabled = True
    End Sub

    Private Sub But_swap_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles But_swap.Click
        If Swap = 1 Then
            But_cal.Text = "=>>"
            txbx_c.Enabled = False
            txbx_m.Enabled = False
            txbx_k.Enabled = False
            txbx_j.Enabled = False
            txbx_f.Enabled = True
            txbx_i.Enabled = True
            txbx_l.Enabled = True
            txbx_b.Enabled = True
            Swap = 0
        Else
            But_cal.Text = "<<="
            txbx_c.Enabled = True
            txbx_m.Enabled = True
            txbx_k.Enabled = True
            txbx_j.Enabled = True
            txbx_f.Enabled = False
            txbx_i.Enabled = False
            txbx_l.Enabled = False
            txbx_b.Enabled = False
            Swap = 1
        End If
        txbx_c.Text = ""
        txbx_m.Text = ""
        txbx_k.Text = ""
        txbx_j.Text = ""
        txbx_f.Text = ""
        txbx_i.Text = ""
        txbx_l.Text = ""
        txbx_b.Text = ""

    End Sub

    Private Sub But_cal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles But_cal.Click
        If But_cal.Text = "=>>" Then
            txbx_c.Text = 0.556 * (Val(txbx_f.Text)) - 17.778
            txbx_m.Text = 25.4 * Val(txbx_i.Text)
            txbx_k.Text = 0.454 * Val(txbx_l.Text)
            txbx_j.Text = 1055.056 * Val(txbx_b.Text)
        Else
            txbx_f.Text = 1.8 * Val(txbx_c.Text) + 32
            txbx_i.Text = 0.039 * Val(txbx_m.Text)
            txbx_l.Text = 2.205 * Val(txbx_k.Text)
            txbx_b.Text = 0.001 * Val(txbx_j.Text)
        End If
    End Sub
End Class