Public Class EXTZ0101

    Private Sub EXTZ0101_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FpSpread1_Sheet1.Visible = True
        FpSpread1_Sheet2.Visible = False
        FpSpread1.Width = 1620
    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If RadioButton1.Checked = True Then
            FpSpread1_Sheet1.Visible = True
            FpSpread1_Sheet2.Visible = False
            FpSpread1.Width = 1620
        Else
            FpSpread1_Sheet1.Visible = False
            FpSpread1_Sheet2.Visible = True
            FpSpread1.Width = 1550
        End If
    End Sub
End Class
