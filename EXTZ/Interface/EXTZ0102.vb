Public Class EXTZ0102

    Private Sub EXTB0102_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FpSpread1_Sheet2.Visible = False
        FpSpread1_Sheet1.Visible = True
        FpSpread1.Width = 1565
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If RadioButton1.Checked = True Then
            FpSpread1_Sheet2.Visible = False
            FpSpread1_Sheet1.Visible = True
            FpSpread1.Width = 1565
        Else
            FpSpread1_Sheet1.Visible = False
            FpSpread1_Sheet2.Visible = True
            FpSpread1.Width = 1655
        End If
    End Sub
End Class
